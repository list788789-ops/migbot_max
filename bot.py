"""
Скелет бота миграционного учёта для MAX.

Реализовано (MVP):
  - /start — приветствие, определение роли (пока только по HR_PHONE_WHITELIST)
  - добавление сотрудника (черновик, consent_status=draft) — obligations НЕ создаются
  - выдача текста согласия на языке сотрудника
  - приём скана согласия (paper_scan) -> consent_status=confirmed -> создание obligations
  - /incomplete — список сотрудников без даты въезда (2026-07, для дозаполнения после переноса
    из ручной xlsx-таблицы, где часть записей была без этого поля)
  - /set_entry_date <id> <ГГГГ-ММ-ДД> — точечное дозаполнение даты въезда существующему сотруднику

Сознательно НЕ реализовано на этом этапе (см. договорённости в диалоге):
  - вход самого сотрудника в бота под своим аккаунтом (bot_button consent)
  - категории patent/visa/hqs — только eaeu/belarus
  - интеграция invoices с 1С — счёт пока просто файл, без API-обмена
  - производственный календарь праздников для working_day (см. deadlines.py)
  - дозаполнение employment_status (все 67 перенесённых записей стоят "уточнить") —
    отдельная задача, сознательно не включена в эту итерацию
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
from maxapi.types.input_media import InputMedia

from models import (
    Base,
    Category,
    Consent,
    ConsentMethod,
    ConsentStatus,
    Employee,
    NotificationSubscriber,
    Obligation,
    ObligationStatus,
    ObligationType,
    RegistrationPeriod,
)
from deadlines import DEADLINE_RULES, compute_deadline, calendar_days_add
from consent_texts import get_consent_text  # см. consent_texts.py
from document_templates import generate_consent_docx, generate_medical_referral_docx

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

    # Заводим первый период учёта по правилу "90 из 180" — только для EAEU, формулировка
    # правила подтверждена именно для этой категории. Для BELARUS механизм иной (изначальные
    # 90 дней — это порог, после которого нужна ПЕРВАЯ регистрация, а не лимит на её действие) —
    # переносить сюда по аналогии не стал, нужна отдельная юридическая проверка.
    if employee.category == Category.EAEU and employee.entry_date is not None:
        existing = (
            session.query(RegistrationPeriod)
            .filter_by(employee_id=employee.id, is_active=True)
            .first()
        )
        if existing is None:
            period = RegistrationPeriod(
                employee_id=employee.id,
                period_start=employee.entry_date,
                period_end=calendar_days_add(employee.entry_date, 90),
                is_active=True,
            )
            session.add(period)

    session.commit()


@dp.bot_started()
async def on_bot_started(event):
    # Регистрируем chat_id как получателя проактивных напоминаний здесь, а не в on_start —
    # у BotStarted event.chat_id подтверждён документацией maxapi напрямую; для MessageCreated
    # точный путь к chat_id не проверен, гадать в коде, который реально рассылает
    # уведомления, рискованнее, чем оставить регистрацию только на этом событии.
    with Session(engine) as session:
        existing = (
            session.query(NotificationSubscriber)
            .filter_by(chat_id=str(event.chat_id))
            .first()
        )
        if existing is None:
            session.add(NotificationSubscriber(chat_id=str(event.chat_id)))
            session.commit()

    await bot.send_message(chat_id=event.chat_id, text=(
        "Бот миграционного учёта.\n"
        "Команды: /add_employee — добавить сотрудника, /pending — список ожидающих согласия, "
        "/incomplete — список без даты въезда, /medical_exam_result <id> <done|failed> — "
        "результат медкомиссии.\n\n"
        "Напоминания о горящих дедлайнах будут приходить в этот чат."
    ))


@dp.message_created(CommandStart())
async def on_start(event: MessageCreated):
    await event.message.answer(
        "Бот миграционного учёта запущен.\n"
        "Доступные команды:\n"
        "/add_employee — добавить сотрудника\n"
        "/pending — сотрудники без подтверждённого согласия\n"
        "/incomplete — сотрудники без даты въезда\n"
        "/medical_exam_result <id> <done|failed> — зафиксировать результат медкомиссии"
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


@dp.message_created(F.message.body.text == "/incomplete")
async def on_incomplete(event: MessageCreated):
    """Список сотрудников без даты въезда — целевая группа для дозаполнения после переноса
    из ручной xlsx-таблицы. Паспорт в списке нужен для идентификации: в данных есть тёзки
    (например, две записи 'Кете'), одного ФИО недостаточно, чтобы понять, кому звонить."""
    with Session(engine) as session:
        employees = (
            session.query(Employee)
            .filter(Employee.entry_date.is_(None))
            .order_by(Employee.full_name)
            .all()
        )

        if not employees:
            await event.message.answer("У всех сотрудников указана дата въезда.")
            return

        lines = [f"Без даты въезда: {len(employees)}\n"]
        for emp in employees:
            passport = f"{emp.passport_series or ''} {emp.passport_number or ''}".strip()
            lines.append(
                f"• {emp.full_name} — паспорт: {passport or 'нет данных'}\n  id={emp.id}"
            )
        lines.append("\nЧтобы указать дату: /set_entry_date <id> <ГГГГ-ММ-ДД>")

    await event.message.answer("\n".join(lines))


async def _handle_set_entry_date(event: MessageCreated, raw_text: str) -> None:
    parts = raw_text.split(maxsplit=2)
    if len(parts) != 3:
        await event.message.answer(
            "Формат: /set_entry_date <id> <ГГГГ-ММ-ДД>\n"
            "id сотрудника смотрите в /incomplete."
        )
        return

    _, employee_id, date_s = parts
    try:
        entry_date = datetime.strptime(date_s.strip(), "%Y-%m-%d").date()
    except ValueError:
        await event.message.answer("Не распознал дату. Формат: ГГГГ-ММ-ДД, например 2026-06-15.")
        return

    with Session(engine) as session:
        employee = session.get(Employee, employee_id.strip())
        if employee is None:
            await event.message.answer("Сотрудник с таким id не найден. Проверьте /incomplete.")
            return

        employee.entry_date = entry_date
        session.add(employee)
        session.commit()
        session.refresh(employee)

        # Если согласие уже было подтверждено раньше (в этой партии — маловероятно, но
        # на будущее, если дозаполнение произойдёт уже после /confirm_consent), не оставляем
        # obligations несозданными молча — досоздаём их сейчас же.
        if employee.consent_status == ConsentStatus.CONFIRMED:
            create_obligations_for_employee(session, employee)

    await event.message.answer(
        f"Дата въезда для {employee.full_name} установлена: {entry_date.strftime('%d.%m.%Y')}."
    )


async def _handle_send_document(event: MessageCreated, raw_text: str, generator_func, doc_label: str) -> None:
    parts = raw_text.split(maxsplit=1)
    if len(parts) != 2:
        await event.message.answer(f"Формат: {parts[0]} <id сотрудника>. id смотрите в /incomplete.")
        return

    employee_id = parts[1].strip()
    with Session(engine) as session:
        employee = session.get(Employee, employee_id)
        if employee is None:
            await event.message.answer("Сотрудник с таким id не найден.")
            return

        try:
            path = generator_func(employee)
        except Exception:
            log.exception("Не удалось сгенерировать документ (%s) для employee_id=%s", doc_label, employee_id)
            await event.message.answer(f"Не удалось сгенерировать документ ({doc_label}). Проверьте логи.")
            return

    await event.message.answer(
        text=f"{doc_label.capitalize()} для {employee.full_name}:",
        attachments=[InputMedia(path=path)],
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

    raw_text = event.message.body.text
    if raw_text:
        if raw_text.startswith("/confirm_consent"):
            try:
                _, employee_id = raw_text.split(maxsplit=1)
            except ValueError:
                await event.message.answer("Формат: /confirm_consent <employee_id>, затем пришлите файл скана.")
                return
            _pending_forms[user_id] = {"state": "awaiting_scan", "employee_id": employee_id}
            await event.message.answer("Пришлите файл скана подписанного согласия.")
            return

        if raw_text.startswith("/set_entry_date"):
            await _handle_set_entry_date(event, raw_text)
            return

        if raw_text.startswith("/send_consent_doc"):
            await _handle_send_document(event, raw_text, generate_consent_docx, "согласие на обработку ПД")
            return

        if raw_text.startswith("/send_medical_referral"):
            await _handle_send_document(
                event, raw_text, generate_medical_referral_docx, "направление на медкомиссию"
            )
            return

        if raw_text.startswith("/medical_exam_result"):
            await _handle_medical_exam_result(event, raw_text)
            return


async def _handle_medical_exam_result(event: MessageCreated, raw_text: str) -> None:
    parts = raw_text.split(maxsplit=2)
    if len(parts) != 3 or parts[2].lower() not in ("done", "failed"):
        await event.message.answer(
            "Формат: /medical_exam_result <id> <done|failed>\nid смотрите в /incomplete."
        )
        return

    _, employee_id, result = parts
    result = result.lower()

    with Session(engine) as session:
        employee = session.get(Employee, employee_id.strip())
        if employee is None:
            await event.message.answer("Сотрудник с таким id не найден.")
            return

        obligation = (
            session.query(Obligation)
            .filter_by(employee_id=employee.id, type=ObligationType.MEDICAL_EXAM)
            .order_by(Obligation.deadline_date.desc())
            .first()
        )
        if obligation is None:
            await event.message.answer(
                f"У {employee.full_name} нет активного обязательства по медкомиссии — "
                "нечего отмечать (возможно, согласие ещё не подтверждено)."
            )
            return

        if result == "done":
            obligation.status = ObligationStatus.DONE
            session.add(obligation)
            session.commit()
            await event.message.answer(f"Медкомиссия для {employee.full_name} отмечена как пройденная.")
        else:
            # ВАЖНО: у Obligation нет поля для текстовой причины отказа/незачёта — при "failed"
            # статус НЕ меняется автоматически (остаётся PENDING/OVERDUE), чтобы дедлайн не
            # потерялся молча. Причину нужно фиксировать вне бота (нет места для неё в модели).
            await event.message.answer(
                f"Зафиксировано: медкомиссия для {employee.full_name} не пройдена. "
                "Статус обязательства не изменён — дедлайн остаётся активным, "
                "причину отказа фиксируйте отдельно (в модели нет поля для этого)."
            )


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
