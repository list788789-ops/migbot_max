"""
Скелет бота миграционного учёта для MAX.

Реализовано (MVP):
  - /start — приветствие, определение роли (пока только по HR_PHONE_WHITELIST)
  - добавление сотрудника (черновик, consent_status=draft) — obligations НЕ создаются
  - выдача текста согласия на языке сотрудника
  - приём скана согласия (paper_scan) -> consent_status=confirmed -> создание obligations

Сознательно НЕ реализовано на этом этапе (см. договорённости в диалоге):
  - вход самого сотрудника в бота под своим аккаунтом (bot_button consent)
  - категории patent/visa/hqs — только eaeu/belarus
  - интеграция invoices с 1С — счёт пока просто файл, без API-обмена
  - производственный календарь праздников для working_day (см. deadlines.py)
"""

import asyncio
import logging
import os
from datetime import datetime

from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from maxapi import Bot, Dispatcher, F
from maxapi.filters.command import CommandStart
from maxapi.types import MessageCreated

from models import (
    Base,
    Category,
    Consent,
    ConsentMethod,
    ConsentStatus,
    Employee,
    Obligation,
    ObligationStatus,
)
from deadlines import DEADLINE_RULES, compute_deadline
from consent_texts import get_consent_text  # см. consent_texts.py

load_dotenv()

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("migbot")

DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./migbot.db")
CONSENT_TEXT_VERSION = os.environ.get("CONSENT_TEXT_VERSION", "v1")
HR_WHITELIST = set(
    p.strip() for p in os.environ.get("HR_PHONE_WHITELIST", "").split(",") if p.strip()
)

engine = create_engine(DATABASE_URL)
Base.metadata.create_all(engine)

bot = Bot()  # токен берётся из MAX_BOT_TOKEN в окружении
dp = Dispatcher()

# Простое in-memory FSM для формы добавления сотрудника.
# На продакшене — заменить на Redis-контекст (maxapi это поддерживает из коробки).
_pending_forms: dict[str, dict] = {}


def is_hr(phone: str | None) -> bool:
    if not HR_WHITELIST:
        return True  # whitelist пуст на этапе разработки — не блокируем
    return phone in HR_WHITELIST


def create_obligations_for_employee(session: Session, employee: Employee) -> None:
    """Вызывается ТОЛЬКО после consent_status=confirmed. Без согласия obligations не создаются —
    это тот самый gate, который обсуждался как обязательное условие."""
    rules = DEADLINE_RULES.get(employee.category, [])
    if not rules:
        log.warning(
            "Нет правил дедлайнов для категории %s (employee_id=%s) — obligations не созданы",
            employee.category,
            employee.id,
        )
        return

    for rule in rules:
        trigger_date = getattr(employee, rule["trigger_field"])
        if trigger_date is None:
            log.warning(
                "Поле %s пустое у employee_id=%s — пропускаю obligation %s",
                rule["trigger_field"],
                employee.id,
                rule["type"],
            )
            continue

        deadline_date = compute_deadline(trigger_date, rule["deadline_value"], rule["deadline_unit"])

        obligation = Obligation(
            employee_id=employee.id,
            type=rule["type"],
            trigger_date=trigger_date,
            deadline_value=rule["deadline_value"],
            deadline_unit=rule["deadline_unit"],
            deadline_date=deadline_date,
            status=ObligationStatus.PENDING,
        )
        session.add(obligation)

    session.commit()


@dp.bot_started()
async def on_bot_started(event):
    await bot.send_message(chat_id=event.chat_id, text=(
        "Бот миграционного учёта.\n"
        "Команды: /add_employee — добавить сотрудника, /pending — список ожидающих согласия."
    ))


@dp.message_created(CommandStart())
async def on_start(event: MessageCreated):
    await event.message.answer(
        "Бот миграционного учёта запущен.\n"
        "Доступные команды:\n"
        "/add_employee — добавить сотрудника\n"
        "/pending — сотрудники без подтверждённого согласия"
    )


@dp.message_created(F.message.body.text == "/add_employee")
async def on_add_employee_start(event: MessageCreated):
    # Упрощённая форма без диалоговых шагов — на первом этапе принимаем данные одним сообщением.
    # Формат (временный, до нормального FSM):
    # ФИО; гражданство; дата_въезда(ГГГГ-ММ-ДД); дата_договора(ГГГГ-ММ-ДД); язык; телефон
    _pending_forms[event.message.sender.user_id] = {"state": "awaiting_employee_data"}
    await event.message.answer(
        "Отправьте данные сотрудника одной строкой через ';':\n"
        "ФИО; гражданство; дата въезда (ГГГГ-ММ-ДД); дата договора (ГГГГ-ММ-ДД); язык; телефон\n\n"
        "Пример:\nИванов Иван; Казахстан; 2026-07-01; 2026-07-03; kk; +7900...\n\n"
        "Категория по умолчанию — eaeu. Для Белоруссии напишите 'belarus' вместо гражданства-триггера "
        "(это временный формат для MVP, потом заменить на нормальную форму с кнопками)."
    )


@dp.message_created(F.message.body.text)
async def on_text(event: MessageCreated):
    user_id = event.message.sender.user_id
    form = _pending_forms.get(user_id)

    if form and form.get("state") == "awaiting_employee_data":
        raw = event.message.body.text
        parts = [p.strip() for p in raw.split(";")]
        if len(parts) != 6:
            await event.message.answer("Не распознал формат. Ожидается 6 полей через ';'.")
            return

        full_name, citizenship, entry_date_s, contract_date_s, language, phone = parts
        category = Category.BELARUS if citizenship.lower() == "belarus" else Category.EAEU

        with Session(engine) as session:
            employee = Employee(
                full_name=full_name,
                citizenship=citizenship,
                category=category,
                entry_date=datetime.strptime(entry_date_s, "%Y-%m-%d").date(),
                contract_date=datetime.strptime(contract_date_s, "%Y-%m-%d").date(),
                language=language or "ru",
                phone=phone or None,
                consent_status=ConsentStatus.DRAFT,
                created_by=str(user_id),
            )
            session.add(employee)
            session.commit()
            session.refresh(employee)

        _pending_forms.pop(user_id, None)

        consent_text = get_consent_text(language)
        await event.message.answer(
            f"Сотрудник {full_name} добавлен как черновик (id={employee.id}).\n"
            f"Обязательства (напоминания, документы) НЕ создаются, пока не подтверждено согласие.\n\n"
            f"--- Текст согласия для передачи сотруднику ({language}) ---\n{consent_text}\n\n"
            f"После подписи пришлите скан командой:\n/confirm_consent {employee.id}"
        )
        return

    # Здесь же ветка для приёма скана после /confirm_consent <id> — упрощена для скелета:
    if raw_text := event.message.body.text:
        if raw_text.startswith("/confirm_consent"):
            try:
                _, employee_id = raw_text.split(maxsplit=1)
            except ValueError:
                await event.message.answer("Формат: /confirm_consent <employee_id>, затем пришлите файл скана.")
                return
            _pending_forms[user_id] = {"state": "awaiting_scan", "employee_id": employee_id}
            await event.message.answer("Пришлите файл скана подписанного согласия.")
            return


@dp.message_created(F.message.body.attachments)
async def on_attachment(event: MessageCreated):
    """Приём файла — используется и для скана согласия, и (в будущем) для перевода паспорта.
    На этом этапе обрабатываем только сценарий awaiting_scan."""
    user_id = event.message.sender.user_id
    form = _pending_forms.get(user_id)

    if not form or form.get("state") != "awaiting_scan":
        await event.message.answer(
            "Файл получен, но я не ожидаю вложение вне сценария подтверждения согласия. "
            "Начните с /confirm_consent <employee_id>."
        )
        return

    employee_id = form["employee_id"]
    attachment = event.message.body.attachments[0]
    file_id = getattr(attachment, "file_id", None) or getattr(attachment, "url", "unknown")

    with Session(engine) as session:
        employee = session.get(Employee, employee_id)
        if employee is None:
            await event.message.answer("Сотрудник не найден — проверьте id.")
            _pending_forms.pop(user_id, None)
            return

        consent = Consent(
            employee_id=employee.id,
            method=ConsentMethod.PAPER_SCAN,
            proof=file_id,
            consent_text_version=CONSENT_TEXT_VERSION,
        )
        session.add(consent)

        employee.consent_status = ConsentStatus.CONFIRMED
        session.add(employee)
        session.commit()
        session.refresh(employee)

        create_obligations_for_employee(session, employee)

    _pending_forms.pop(user_id, None)
    await event.message.answer(
        f"Согласие подтверждено для {employee.full_name}. Обязательства созданы, напоминания активны."
    )


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
