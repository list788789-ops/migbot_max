"""
Ежедневная проверка Obligation: помечает просроченные как OVERDUE и шлёт напоминания
всем подписчикам (NotificationSubscriber) о том, что горит или уже просрочено.

Запускать как ОТДЕЛЬНЫЙ сервис Railway (Cron Job) — та же причина, что и у
check_registration_renewals.py: встроить в процесс бота значит потерять состояние
при каждом рестарте бота.

Настройка в Railway: New -> Cron Job Service, команда —
  python check_obligation_deadlines.py
расписание — например "0 6 * * *" (06:00 UTC = 09:00 МСК — Railway всегда планирует
cron по UTC, часовой пояс расписания не настраивается, пересчёт делать вручную).

Требует MAX_BOT_TOKEN в окружении — этот скрипт сам создаёт Bot() для отправки
сообщений, независимо от процесса bot.py.
"""

import asyncio
import os
from datetime import date, datetime, timedelta, timezone

from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from maxapi import Bot
from models import Employee, NotificationSubscriber, Obligation, ObligationStatus

load_dotenv()

MSK = timezone(timedelta(hours=3))  # Мурманская обл. — московское время, без перехода на летнее с 2014

DATABASE_URL = os.environ["DATABASE_URL"]
WARNING_DAYS = int(os.environ.get("DEADLINE_WARNING_DAYS", "3"))

OBLIGATION_LABELS = {
    "registration": "постановка на миграционный учёт",
    "contract_notice": "уведомление МВД о трудовом договоре",
    "contract_termination_notice": "уведомление МВД о расторжении договора",
    "medical_exam": "медицинское освидетельствование",
    "patent_payment": "оплата патента",
    "efs1_report": "отчёт ЕФС-1 в СФР",
    "registration_renewal": "продление регистрации (90/180)",
}


async def send_reminders():
    engine = create_engine(DATABASE_URL)
    # date.today() зависит от системного часового пояса контейнера (на Railway обычно UTC) —
    # берём дату явно по МСК, иначе у полуночи МСК/UTC возможен сдвиг на день в сравнениях дедлайнов.
    today = datetime.now(MSK).date()
    warning_threshold = today + timedelta(days=WARNING_DAYS)

    with Session(engine) as session:
        # Сначала переводим просроченные PENDING в OVERDUE — без этого шага статус
        # никогда не меняется сам по себе, дедлайн проходит "молча" в БД.
        overdue_candidates = (
            session.query(Obligation)
            .filter(
                Obligation.status == ObligationStatus.PENDING,
                Obligation.deadline_date < today,
            )
            .all()
        )
        for ob in overdue_candidates:
            ob.status = ObligationStatus.OVERDUE
            session.add(ob)
        session.commit()

        pending_and_overdue = (
            session.query(Obligation)
            .filter(
                Obligation.status.in_([ObligationStatus.PENDING, ObligationStatus.OVERDUE]),
                Obligation.deadline_date <= warning_threshold,
            )
            .order_by(Obligation.deadline_date)
            .all()
        )

        if not pending_and_overdue:
            print("Горящих или просроченных обязательств нет.")
            return

        lines = ["Дедлайны, требующие внимания:\n"]
        for ob in pending_and_overdue:
            employee = session.get(Employee, ob.employee_id)
            name = employee.full_name if employee else f"[сотрудник {ob.employee_id} не найден]"
            label = OBLIGATION_LABELS.get(ob.type.value if hasattr(ob.type, "value") else ob.type, ob.type)
            marker = "🔴 ПРОСРОЧЕНО" if ob.status == ObligationStatus.OVERDUE else "🟡"
            lines.append(
                f"{marker} {name} — {label}, дедлайн {ob.deadline_date.strftime('%d.%m.%Y')}"
            )
        message_text = "\n".join(lines)

        subscribers = session.query(NotificationSubscriber).all()
        if not subscribers:
            print("Подписчиков нет (никто не запускал /start после обновления) — некому слать.")
            print(message_text)
            return

    bot = Bot()
    for sub in subscribers:
        try:
            await bot.send_message(chat_id=sub.chat_id, text=message_text)
        except Exception as e:
            print(f"Не удалось отправить напоминание в chat_id={sub.chat_id}: {e}")

    print(f"Разослано {len(subscribers)} подписчикам. Обязательств в списке: {len(pending_and_overdue)}")


if __name__ == "__main__":
    asyncio.run(send_reminders())
