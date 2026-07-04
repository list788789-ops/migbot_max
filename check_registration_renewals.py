"""
Периодическая проверка периодов регистрации ("90 из 180") — создаёт Obligation
типа REGISTRATION_RENEWAL, когда до окончания активного периода осталось <= RENEWAL_WARNING_DAYS.

Запускать как ОТДЕЛЬНЫЙ сервис Railway (Cron Job), не внутри процесса bot.py — намеренно:
если встроить в общий asyncio-луп бота, рестарт/редеплой бота сбрасывает состояние джоба.
Cron Job Railway запускает этот скрипт независимо по расписанию (например, раз в сутки),
не завязываясь на то, жив ли сейчас процесс бота.

Идемпотентность: перед созданием Obligation проверяется, нет ли уже созданного renewal-
обязательства с тем же employee_id и deadline_date — повторный запуск в тот же день
не создаст дубликат.

Настройка в Railway: New -> Cron Job Service, команда запуска —
  python check_registration_renewals.py
расписание — например "0 6 * * *" (каждый день в 06:00 UTC).
"""

import os
from datetime import date, timedelta

from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from models import Obligation, ObligationStatus, ObligationType, RegistrationPeriod

load_dotenv()

DATABASE_URL = os.environ["DATABASE_URL"]  # обязателен — без него скрипт бессмысленен
RENEWAL_WARNING_DAYS = int(os.environ.get("RENEWAL_WARNING_DAYS", "7"))


def main():
    engine = create_engine(DATABASE_URL)
    today = date.today()
    warning_threshold = today + timedelta(days=RENEWAL_WARNING_DAYS)

    with Session(engine) as session:
        periods = (
            session.query(RegistrationPeriod)
            .filter(
                RegistrationPeriod.is_active.is_(True),
                RegistrationPeriod.period_end <= warning_threshold,
            )
            .all()
        )

        created, skipped = 0, 0
        for period in periods:
            existing = (
                session.query(Obligation)
                .filter_by(
                    employee_id=period.employee_id,
                    type=ObligationType.REGISTRATION_RENEWAL,
                    deadline_date=period.period_end,
                )
                .first()
            )
            if existing is not None:
                skipped += 1
                continue

            obligation = Obligation(
                employee_id=period.employee_id,
                type=ObligationType.REGISTRATION_RENEWAL,
                trigger_date=period.period_start,
                deadline_value=90,
                deadline_unit="calendar_day",
                deadline_date=period.period_end,
                status=ObligationStatus.PENDING,
            )
            session.add(obligation)
            created += 1

        session.commit()

    print(f"Проверено периодов: {len(periods)}, создано напоминаний: {created}, уже существовало: {skipped}")


if __name__ == "__main__":
    main()
