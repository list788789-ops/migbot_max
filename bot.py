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

2026-07: create_obligations_for_employee() вынесена в obligations.py — импортируется оттуда,
чтобы webforms.py (веб-формы кадровика) могла вызывать ту же функцию, не дублируя логику
дедлайнов и не импортируя bot.py целиком (это создало бы второй Bot()/Dispatcher в чужом
процессе). Если меняешь правила создания obligations — правь только obligations.py.

2026-07: _handle_send_document исправлена в двух местах:
  1. Раньше ValueError от _require_fields (document_templates.py) тонул в общем
     "except Exception: ...Проверьте логи" — кадровик в чате не видел, какого именно
     поля не хватает, хотя document_templates.py явно требует показывать текст этого
     исключения. Теперь ValueError перехватывается отдельно и его текст уходит в чат.
  2. Добавлена проверка отсутствующих полей ДО генерации (check_consent_fields /
     check_medical_referral_fields) — как в webforms.py. В тестовом режиме
     (TEST_ALLOW_MISSING_FIELDS=true, флаг живёт в document_templates.py) документ всё
     равно генерируется с прочерками, но к сообщению в чате добавляется тот же
     текст-баннер, что и в docx/HTML-превью веб-форм. В MAX нет способа показать
     произвольный HTML внутри чата (только текст/файлы/кнопки), поэтому здесь это
     текстовый эквивалент HTML-превью, а не сам HTML.
"""

import asyncio
import logging
import os
import shutil
import tempfile
from datetime import date, datetime, timezone, timedelta

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger

MSK = timezone(timedelta(hours=3))  # Мурманская обл. — московское время, без перехода на летнее с 2014

from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from maxapi import Bot, Dispatcher, F
from maxapi.filters.command import CommandStart
from maxapi.types import MessageCreated
from maxapi.types.input_media import InputMedia
from maxapi.types.updates.message_callback import MessageCallback
from maxapi.utils.inline_keyboard import InlineKeyboardBuilder
from maxapi.types.attachments.buttons import CallbackButton

from models import (
    Base,
    Category,
    Consent,
    ConsentMethod,
    ConsentStatus,
    Employee,
    InstructionType,
    NotificationSubscriber,
    Obligation,
    ObligationStatus,
    ObligationType,
    RotationReturn,
    SystemFlag,
    WorkOrder,
    WorkOrderMember,
)
from obligations import create_obligations_for_employee
import tabel
import reports as reports_data
import production as prod
from auth_binding import (
    bind_max_account, find_user_by_max_id, get_role_label,
    confirm_max_code, register_via_max,
    set_max_chat_id, get_admins_with_chat,
)
from common_utils import category_for_citizenship
from consent_texts import get_consent_text  # см. consent_texts.py
from s3_storage import (
    SCAN_TYPES, _s3_list_for_employee, _s3_download, _ext_for,
)
from document_templates import (
    TEST_ALLOW_MISSING_FIELDS,
    check_consent_fields,
    check_medical_referral_fields,
    docx_to_pdf,
    generate_consent_docx,
    generate_employees_xlsx,
    generate_medical_referral_docx,
)

load_dotenv()

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("migbot")

DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./migbot.db")
CONSENT_TEXT_VERSION = os.environ.get("CONSENT_TEXT_VERSION", "v1")
HR_WHITELIST = set(
    p.strip() for p in os.environ.get("HR_PHONE_WHITELIST", "").split(",") if p.strip()
)
# Организация для бланка наряда-допуска — тот же источник, что в webforms.py (ORG_NAME),
# чтобы наряд из бота и из веба печатался с идентичным юрлицом.
ORG_NAME = os.environ.get("COMPANY_NAME", "ИП Буц Сергей Юрьевич")

engine = create_engine(DATABASE_URL)
Base.metadata.create_all(engine)

bot = Bot()  # токен берётся из MAX_BOT_TOKEN в окружении
dp = Dispatcher()

# Простое in-memory FSM для формы добавления сотрудника.
# На продакшене — заменить на Redis-контекст (maxapi это поддерживает из коробки).
_pending_forms: dict[str, dict] = {}

# (user_id, prefix) -> message_id открытого списка-пикера. Нужен, чтобы после подтверждения
# редактировать ИМЕННО этот список (убирая одного сотрудника), а не удалять его целиком.
# In-memory: переживёт до рестарта процесса, после рестарта старые списки просто перестанут
# обновляться editом (упадут в except при попытке отредактировать несуществующий message_id
# в памяти — на практике это будет KeyError при .get, обработано ниже как "просто отправить заново").
_open_pickers: dict[tuple[str, str], str] = {}


def is_hr(phone: str | None) -> bool:
    if not HR_WHITELIST:
        return True  # whitelist пуст на этапе разработки — не блокируем
    return phone in HR_WHITELIST


class _Responder:
    """Абстракция отправки ответа поверх двух разных API: event.message.answer()
    у MessageCreated и bot.send_message(chat_id=...) у MessageCallback (chat_id там
    берётся из подтверждённого документацией event.get_ids(), в отличие от MessageCreated,
    где путь к chat_id не проверен — поэтому для него используется answer(), не send_message)."""

    def __init__(self, event):
        self._event = event

    async def send(self, text: str, attachments=None):
        if isinstance(self._event, MessageCallback):
            chat_id, _ = self._event.get_ids()
            if attachments:
                sent = await bot.send_message(chat_id=chat_id, text=text, attachments=attachments)
            else:
                sent = await bot.send_message(chat_id=chat_id, text=text)
        else:
            if attachments:
                sent = await self._event.message.answer(text=text, attachments=attachments)
            else:
                sent = await self._event.message.answer(text)
        # send_message/answer возвращают объект с .id (см. wiki maxapi: SendedMessage.id) —
        # не проверено на реальном инстансе, что answer() возвращает то же самое, что
        # send_message; если это не так, редактирование списка после подтверждения не сработает
        # и молча упадёт в исключение при вызове bot.edit_message ниже.
        log.info("Responder.send() вернул объект типа %s: %r", type(sent).__name__, sent)
        return getattr(sent, "id", None)

    def user_id(self) -> str:
        if isinstance(self._event, MessageCallback):
            _, user_id = self._event.get_ids()
            return user_id
        return self._event.message.sender.user_id

    def current_message_id(self):
        """ID сообщения, из которого пришёл callback (для редактирования на месте). None, если
        событие не callback или структура иная."""
        try:
            if isinstance(self._event, MessageCallback):
                return self._event.message.body.mid
        except Exception:
            pass
        return None

    async def show_menu(self, text: str, attachments):
        """Показать меню: если это callback (нажали кнопку) — РЕДАКТИРУЕМ текущее сообщение,
        чтобы не плодить копии меню в чате. Если редактирование недоступно/упало — отправляем
        новое (fallback, поведение как раньше)."""
        mid = self.current_message_id()
        if mid is not None:
            try:
                await bot.edit_message(message_id=mid, text=text, attachments=attachments)
                return
            except Exception:
                log.info("show_menu: edit_message не удался, шлём новое сообщение")
        await self.send(text=text, attachments=attachments)


def _build_main_menu(role: str | None = None) -> InlineKeyboardBuilder:
    """Главное меню — разделы. Пункты внутри разделов: _build_section_menu().
    Структура расширяемая: чтобы добавить пункт — допиши в нужный раздел ниже.

    2026-07: слияние с ботом ТабельБелокаменка. "☀️ Утро"/"🌙 Вечер" — разметка
    явки, доступна PRORAB/KADROVIK/ADMIN (см. UserRole в models.py, узкое
    исключение из "PRORAB не пишет в БД"). "⚠️ Требует внимания" скрыт от
    PRORAB совсем — там теперь в том числе флаги межвахты с обязательствами,
    это забота кадровика, не прораба (см. договорённость про "данные
    направлены в отдел кадров")."""
    builder = InlineKeyboardBuilder()
    if role in ("prorab", "kadrovik", "admin"):
        builder.row(CallbackButton(text="☀️ Утро (явка)", payload="menu:morning"))
        builder.row(CallbackButton(text="🌙 Вечер (ночная смена)", payload="menu:evening"))
        builder.row(CallbackButton(text="🧹 Действия с сотрудником", payload="menu:empaction"))
        builder.row(CallbackButton(text="📄 Наряды-допуски", payload="menu:section:workorders"))
        builder.row(CallbackButton(text="📋 Приказы ОТ/ТБ", payload="menu:section:otorders"))
        builder.row(CallbackButton(text="📓 Журналы инструктажей", payload="menu:section:instrjournals"))
    builder.row(CallbackButton(text="👥 Сотрудники", payload="menu:section:employees"))
    if role in ("kadrovik", "admin"):
        builder.row(CallbackButton(text="⚠️ Требует внимания", payload="menu:section:attention"))
    builder.row(CallbackButton(text="📊 Отчёты", payload="menu:section:reports"))
    return builder


def _build_section_menu(section: str, role: str | None = None) -> InlineKeyboardBuilder:
    """Подменю раздела. Внизу каждого — кнопка «Назад» в главное меню (menu:main)."""
    builder = InlineKeyboardBuilder()
    if section == "employees":
        builder.row(CallbackButton(text="➕ Добавить сотрудника", payload="menu:add_employee"))
        builder.row(CallbackButton(text="📋 Список сотрудников", payload="menu:employees"))
        builder.row(CallbackButton(text="📎 Документы работника", payload="menu:docpick"))
        builder.row(CallbackButton(text="🖨 Печать согласия", payload="menu:consent_print"))
    elif section == "attention":
        builder.row(CallbackButton(text="⏳ Без даты въезда", payload="menu:incomplete"))
        builder.row(CallbackButton(text="🗓 Без даты договора", payload="menu:contractdate"))
        builder.row(CallbackButton(text="🖊 Ожидают согласия", payload="menu:pending_consent"))
        if role in ("kadrovik", "admin"):
            builder.row(CallbackButton(text="🚨 ЕСТЬ ЯВКА БЕЗ ДОГОВОРА",
                                        payload="menu:invalid_contract_marks"))
            builder.row(CallbackButton(text="🔄 Межвахта — открытые обязательства",
                                        payload="menu:rotation_flags"))
            builder.row(CallbackButton(text="🧾 На оформлении (не в табеле)",
                                        payload="menu:onboarding"))
            builder.row(CallbackButton(text="❓ Уточнить дату возврата (МЖ)",
                                        payload="menu:pending_rotation"))
    elif section == "workorders":
        builder.row(CallbackButton(text="🟢 Активные наряды", payload="menu:wolist:active"))
        builder.row(CallbackButton(text="🗄 Архив нарядов", payload="menu:wolist:past"))
    elif section == "reports":
        builder.row(CallbackButton(text="📅 Табель за сегодня", payload="menu:tabel_today"))
        builder.row(CallbackButton(text="🐛 Журнал патчей", payload="menu:report_changelog"))
        builder.row(CallbackButton(text="📊 Проблемные за месяц", payload="menu:report_monthly_problems"))
        builder.row(CallbackButton(text="📋 Обязательства", payload="menu:report_obligations"))
        builder.row(CallbackButton(text="🕵️ Активность в табеле", payload="menu:report_activity"))
    builder.row(CallbackButton(text="⬅️ Назад", payload="menu:main"))
    return builder


def _role_for_max_id(session: Session, max_user_id: str) -> str | None:
    """Роль пользователя по MAX user_id, или None, если не привязан ('/login' не выполнен)."""
    user = find_user_by_max_id(session, max_user_id)
    if user is None or user.role is None:
        return None
    return user.role.value


_SECTION_TITLES = {
    "employees": "👥 Сотрудники",
    "attention": "⚠️ Требует внимания",
    "workorders": "📄 Наряды-допуски",
    "reports": "📊 Отчёты (в разработке)",
}


async def _start_add_employee_flow(responder: "_Responder") -> None:
    _pending_forms[responder.user_id()] = {"state": "awaiting_employee_data"}
    await responder.send(
        "Отправьте данные сотрудника одной строкой через ';':\n"
        "ФИО; гражданство; дата въезда (ГГГГ-ММ-ДД); дата договора (ГГГГ-ММ-ДД); язык; телефон\n\n"
        "Пример:\nИванов Иван; Казахстан; 2026-07-01; 2026-07-03; kk; +7900...\n\n"
        "Категория по умолчанию — eaeu. Для Белоруссии напишите 'belarus' вместо гражданства-триггера "
        "(это временный формат для MVP)."
    )


def _tabel_extra_button(prefix: str) -> CallbackButton | None:
    if prefix == "utroday":
        return CallbackButton(text="✅ Отметил всех присутствующих", payload="utro_done")
    if prefix == "eveningnight":
        return CallbackButton(text="✅ Готово", payload="evening_done")
    return None


async def _deliver_employees_list(responder: "_Responder") -> None:
    """Список сотрудников — теперь файлом xlsx, не постраничным текстом. Генерируется
    по текущему состоянию БД на момент запроса, независимо от Google Sheets (тот
    обновляется по крону через export_to_sheets_api.py и может отставать).

    2026-07: исключены уволенные (contract_end_date заполнен) — раньше список
    показывал вообще всех, включая давно уволенных, что не соответствовало
    ожиданию "список действующих сотрудников"."""
    with Session(engine) as session:
        employees = (
            session.query(Employee)
            .filter(Employee.contract_end_date.is_(None))
            .order_by(Employee.full_name)
            .all()
        )
        if not employees:
            await responder.send("Действующих сотрудников в базе нет.")
            return
        path = generate_employees_xlsx(employees)

    await responder.send(
        text=f"Всего сотрудников: {len(employees)}",
        attachments=[InputMedia(path=path)],
    )


PICKER_PAGE_SIZE = 25  # запас от подтверждённого лимита MAX: 30 рядов на сообщение
# (dev.max.ru/docs-api: максимум 210 кнопок, 30 рядов, до 7 в ряду — здесь 1 кнопка
# на ряд, значит лимит по факту 30; проявилось в проде как errors.maxRows на 67 записях)

PICKER_TITLES = {
    "empdate": "Без даты въезда",
    "docpick": "Выберите работника — посмотреть документы",
    "delpick": "Выберите сотрудника для удаления (тест, необратимо)",
    "consentpick": "Ожидают согласия",
    "contractdate": "Без даты договора",
    "utroday": "☀️ Утро — отметьте, кто на месте",
    "eveningnight": "🌙 Вечер — кто заступает в ночь",
    "empaction": "🧹 Действия с сотрудником — выберите",
    "consentprint": "Печать согласия — выберите работника",
}
# Префиксы табеля — там незачем показывать паспорт в подписи кнопки (не тот контекст).
_TABEL_PREFIXES = {"utroday", "eveningnight", "empaction"}


def _picker_employees(session: Session, prefix: str) -> list[Employee]:
    if prefix == "empdate":
        return (
            session.query(Employee)
            .filter(Employee.entry_date.is_(None))
            .order_by(Employee.full_name)
            .all()
        )
    if prefix == "delpick":
        return session.query(Employee).order_by(Employee.full_name).all()
    if prefix == "docpick":
        return session.query(Employee).order_by(Employee.full_name).all()
    if prefix == "consentprint":
        # Печать бланка согласия для ЛЮБОГО сотрудника (не только ожидающих) — весь
        # список по алфавиту, как docpick. Статус согласия здесь не фильтруем.
        return session.query(Employee).order_by(Employee.full_name).all()
    if prefix == "consentpick":
        return (
            session.query(Employee)
            .filter_by(consent_status=ConsentStatus.DRAFT)
            .order_by(Employee.full_name)
            .all()
        )
    if prefix == "contractdate":
        return (
            session.query(Employee)
            .filter(Employee.contract_date.is_(None))
            .order_by(Employee.full_name)
            .all()
        )
    if prefix == "utroday":
        return tabel.get_unmarked_day(session)
    if prefix == "eveningnight":
        return tabel.get_not_worked_day(session)
    if prefix == "empaction":
        return tabel.get_active_employees(session)
    return []


async def _deliver_picker(
    responder: "_Responder", prefix: str, page: int = 0, edit: bool = False,
    only_if_open: bool = False, extra_button: CallbackButton | None = None,
) -> None:
    """Общий постраничный список сотрудников-кнопок для empdate/delpick/consentpick/
    contractdate/utroday/eveningnight.
    Если edit=True и для (user_id, prefix) уже есть открытый список — редактируем его на месте
    (bot.edit_message), а не отправляем новое сообщение. Так после подтверждения действия
    список остаётся на экране, просто без обработанного сотрудника, вместо полного удаления
    и замены единственной строкой результата.

    only_if_open=True: если список не был открыт (например, дата введена через /set_entry_date,
    а не через кнопку) — ничего не отправляем, не заводим список, который никто не открывал.

    extra_button: дополнительная кнопка отдельной строкой над «Выход» (например,
    «✅ Отметил всех присутствующих» для utroday)."""
    key = (responder.user_id(), prefix)
    if only_if_open and key not in _open_pickers:
        return

    # Если для этого пользователя список этого типа УЖЕ открыт — редактируем его
    # на месте, даже если вызвано открытие из меню (edit=False). Иначе повторное
    # нажатие «Утро (явка)»/«Вечер» плодит новый список поверх старого (дубли на
    # экране — см. отчёт с двумя наборами кнопок). edit-ветка ниже сама откатится
    # на отправку нового, если старое сообщение уже устарело/удалено.
    if key in _open_pickers:
        edit = True

    with Session(engine) as session:
        employees = _picker_employees(session, prefix)

        if not employees:
            text = "Список пуст."
            if edit and key in _open_pickers:
                try:
                    await bot.edit_message(message_id=_open_pickers[key], text=text, attachments=[])
                except Exception:
                    log.exception("Не удалось отредактировать пустой список (prefix=%s)", prefix)
                _open_pickers.pop(key, None)
            else:
                await responder.send(text)
            return

        total = len(employees)
        start = page * PICKER_PAGE_SIZE
        chunk = employees[start : start + PICKER_PAGE_SIZE]

        builder = InlineKeyboardBuilder()
        for emp in chunk:
            if prefix in _TABEL_PREFIXES:
                label = emp.full_name
            else:
                passport = f"{emp.passport_series or ''} {emp.passport_number or ''}".strip()
                label = f"{emp.full_name} ({passport})" if passport else emp.full_name
            builder.row(CallbackButton(text=label[:60], payload=f"{prefix}:{emp.id}"))

        total_pages = (total + PICKER_PAGE_SIZE - 1) // PICKER_PAGE_SIZE
        if total_pages > 1:
            nav = []
            if page > 0:
                nav.append(CallbackButton(text="◀️", payload=f"page:{prefix}:{page - 1}"))
            if page < total_pages - 1:
                nav.append(CallbackButton(text="▶️", payload=f"page:{prefix}:{page + 1}"))
            if nav:
                builder.row(*nav)

        if extra_button is not None:
            builder.row(extra_button)
        builder.row(CallbackButton(text="🚪 Выход", payload=f"exitpicker:{prefix}"))

    header = f"{PICKER_TITLES.get(prefix, prefix)}: {total}"
    if total_pages > 1:
        header += f"\nСтраница {page + 1}/{total_pages}"

    # Редактируем сообщение, из которого пришёл колбэк (current_message_id) — оно
    # всегда доступно в событии, НЕ зависит от _open_pickers в памяти. Раньше edit
    # шёл по _open_pickers[key], который теряется при рестарте бота: после деплоя
    # навигация (page:) с only_if_open молча выходила, «Выход» не находил список —
    # кнопки становились мёртвыми. Теперь, как у меню (show_menu), редактируем по
    # mid из колбэка — кнопки работают и после перезапуска.
    if edit:
        mid = getattr(responder, "current_message_id", lambda: None)()
        if mid is None and key in _open_pickers:
            mid = _open_pickers[key]  # fallback на память (напр. edit не из колбэка)
        if mid is not None:
            try:
                await bot.edit_message(
                    message_id=mid, text=header, attachments=[builder.as_markup()]
                )
                _open_pickers[key] = mid  # синхронизируем память с реальным сообщением
                return
            except Exception:
                # Сообщение устарело/удалено — откатываемся на новое, не проваливаем действие.
                log.exception("Не удалось отредактировать список (prefix=%s), отправляю заново", prefix)

    message_id = await responder.send(text=header, attachments=[builder.as_markup()])
    if message_id:
        _open_pickers[key] = message_id


async def _deliver_delete_confirmation(responder: "_Responder", employee_id: str) -> None:
    with Session(engine) as session:
        employee = session.get(Employee, employee_id)
        if employee is None:
            await responder.send("Сотрудник не найден.")
            return
        full_name = employee.full_name

    builder = InlineKeyboardBuilder()
    builder.row(CallbackButton(text="✅ Подтвердить удаление", payload=f"delconfirm:{employee_id}"))
    builder.row(CallbackButton(text="❌ Отмена", payload="cancel:delpick"))

    await responder.send(
        text=f"Удалить {full_name} безвозвратно, вместе со всей историей "
        "(согласия, обязательства, направления)?",
        attachments=[builder.as_markup()],
    )


async def _deliver_consent_confirmation(responder: "_Responder", employee_id: str) -> None:
    with Session(engine) as session:
        employee = session.get(Employee, employee_id)
        if employee is None:
            await responder.send("Сотрудник не найден.")
            return
        full_name = employee.full_name

    builder = InlineKeyboardBuilder()
    builder.row(CallbackButton(text="📄 Бланк согласия (ТСМ)", payload=f"consentblank:tsm:{employee_id}"))
    builder.row(CallbackButton(text="📄 Бланк согласия (ИП Буц)", payload=f"consentblank:ip:{employee_id}"))
    builder.row(CallbackButton(text="✅ Подтвердить (кнопкой, тест)", payload=f"consentconfirm:{employee_id}"))
    builder.row(CallbackButton(text="❌ Отмена", payload="cancel:consentpick"))

    await responder.send(
        text=f"Согласие для {full_name}.\n\n"
        "Как распечатать бланк:\n"
        "1) нажмите «Бланк согласия (ТСМ)» — или «(ИП Буц)», если оператор ИП;\n"
        "2) бот пришлёт PDF — распечатайте его;\n"
        "3) работник подписывает бумажный бланк;\n"
        "4) отсканируйте/сфотографируйте подпись и прикрепите скан "
        "(«✅ Подтвердить» → пришлите файл).\n\n"
        "Бланк печатается даже если в карточке заполнены не все поля — тогда в "
        "документе будет пометка «черновик», а пустые места заполните от руки.\n\n"
        "Кнопка «Подтвердить (кнопкой)» — тестовый способ, юридически слабее "
        "сканированной подписи (ст. 9 152-ФЗ требует осознанного согласия, клик "
        "без верификации личности это не подтверждает).",
        attachments=[builder.as_markup()],
    )


async def _deliver_consent_print_choice(responder: "_Responder", employee_id: str) -> None:
    """Диалог печати бланка согласия для любого сотрудника (пункт «🖨 Печать согласия»).
    В отличие от _deliver_consent_confirmation здесь НЕТ кнопки «Подтвердить» — это меню
    только для печати, статус согласия не меняется. Печатать бланк можно для кого угодно,
    вне зависимости от текущего статуса (генерация PDF от него не зависит)."""
    with Session(engine) as session:
        employee = session.get(Employee, employee_id)
        if employee is None:
            await responder.send("Сотрудник не найден.")
            return
        full_name = employee.full_name

    builder = InlineKeyboardBuilder()
    builder.row(CallbackButton(text="📄 Бланк согласия (ТСМ)", payload=f"consentblank:tsm:{employee_id}"))
    builder.row(CallbackButton(text="📄 Бланк согласия (ИП Буц)", payload=f"consentblank:ip:{employee_id}"))
    builder.row(CallbackButton(text="⬅️ Назад к списку", payload="menu:consent_print"))

    await responder.send(
        text=f"Печать согласия — {full_name}.\n\n"
        "Нажмите «Бланк согласия (ТСМ)» — или «(ИП Буц)», если оператор ИП. "
        "Бот пришлёт PDF, распечатайте и дайте работнику подписать.\n\n"
        "Бланк печатается даже при незаполненной карточке — тогда в документе будет "
        "пометка «черновик», а пустые места заполните от руки.",
        attachments=[builder.as_markup()],
    )


async def _execute_consent_confirm_by_button(responder: "_Responder", employee_id: str) -> None:
    with Session(engine) as session:
        employee = session.get(Employee, employee_id)
        if employee is None:
            await responder.send("Сотрудник не найден.")
            return

        consent = Consent(
            employee_id=employee.id,
            method=ConsentMethod.BOT_BUTTON,
            proof=f"button_click:{responder.user_id()}:{datetime.now(MSK).isoformat()}",
            consent_text_version=CONSENT_TEXT_VERSION,
        )
        session.add(consent)

        employee.consent_status = ConsentStatus.CONFIRMED
        session.add(employee)
        session.commit()
        session.refresh(employee)

        create_obligations_for_employee(session, employee)
        full_name = employee.full_name

    await responder.send(f"Согласие подтверждено (кнопкой) для {full_name}. Обязательства созданы.")


async def _execute_delete_employee(responder: "_Responder", employee_id: str) -> None:
    with Session(engine) as session:
        employee = session.get(Employee, employee_id)
        if employee is None:
            await responder.send("Сотрудник уже удалён или не найден.")
            return
        full_name = employee.full_name
        session.delete(employee)  # каскад удалит consents/obligations/
        # registration_periods/referrals/invoices — см. cascade="all, delete-orphan" в models.py
        session.commit()

    await responder.send(f"Сотрудник {full_name} удалён.")


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

    await bot.send_message(
        chat_id=event.chat_id,
        text=(
            "ТАКТ — учёт персонала и миграционного контроля.\n"
            "Выберите действие или используйте команды: /medical_exam_result <id> <done|failed>, "
            "/set_entry_date <id> <ГГГГ-ММ-ДД>, /send_consent_doc <id>, /send_medical_referral <id>.\n\n"
            "Напоминания о горящих дедлайнах будут приходить в этот чат."
        ),
        attachments=[_build_main_menu().as_markup()],
    )


@dp.message_created(CommandStart())
async def on_start(event: MessageCreated):
    with Session(engine) as session:
        role = _role_for_max_id(session, event.message.sender.user_id)
    text = "ТАКТ — АСУ строительного производства. Выберите действие:"
    if role:
        text = f"Рабочее место: ТАКТ — АСУ строительного производства. Роль: {role}.\nВыберите действие:"
    await event.message.answer(
        text=text,
        attachments=[_build_main_menu(role).as_markup()],
    )


@dp.message_created(F.message.body.text == "/add_employee")
async def on_add_employee_start(event: MessageCreated):
    await _start_add_employee_flow(_Responder(event))


@dp.message_created(F.message.body.text == "/employees")
async def on_employees(event: MessageCreated):
    await _deliver_employees_list(_Responder(event))


@dp.message_created(F.message.body.text == "/incomplete")
async def on_incomplete(event: MessageCreated):
    await _deliver_picker(_Responder(event), "empdate")


@dp.message_created(F.message.body.text == "/login")
async def on_login_start(event: MessageCreated):
    """Привязка MAX-аккаунта к пользователю системы — см. auth_binding.py."""
    user_id = event.message.sender.user_id
    with Session(engine) as session:
        existing = find_user_by_max_id(session, user_id)
        if existing is not None:
            await event.message.answer(
                f"Вы уже вошли как {existing.full_name} ({get_role_label(existing)}). "
                f"Если это не вы — обратитесь к админу."
            )
            return
    _pending_forms[user_id] = {"state": "awaiting_login_phone"}
    await event.message.answer(
        "Введите номер телефона, под которым вас зарегистрировал кадровик "
        "(тот же, что при подаче заявки на веб-форме)."
    )


@dp.message_created(F.message.body.text.startswith("/confirm"))
async def on_confirm_code(event: MessageCreated):
    """Привязка MAX по коду с веб-регистрации — см. auth_binding.confirm_max_code."""
    user_id = event.message.sender.user_id
    parts = event.message.body.text.strip().split(maxsplit=1)
    if len(parts) < 2 or not parts[1].strip():
        await event.message.answer("Формат: /confirm 123456 (код с сайта после регистрации).")
        return
    code = parts[1].strip()
    with Session(engine) as session:
        ok, text = confirm_max_code(session, code, user_id)
    await event.message.answer(text)


@dp.message_created(F.message.body.text == "/register")
async def on_register_start(event: MessageCreated):
    """Регистрация с нуля прямо в боте, без веба — см. auth_binding.register_via_max."""
    user_id = event.message.sender.user_id
    with Session(engine) as session:
        existing = find_user_by_max_id(session, user_id)
        if existing is not None:
            await event.message.answer(
                f"Этот MAX-аккаунт уже привязан к {existing.full_name} "
                f"({get_role_label(existing)}). Обратитесь к админу, если это ошибка."
            )
            return
    _pending_forms[user_id] = {"state": "awaiting_register_name"}
    await event.message.answer("Регистрация. Введите ваше ФИО полностью:")


@dp.message_created(F.message.body.text == "/subscribe")
async def on_subscribe(event: MessageCreated):
    """Подписать ТЕКУЩИЙ чат (в т.ч. групповой) на утреннюю рассылку напоминаний.
    Отвечает кнопкой, а не регистрирует сразу: у MessageCreated путь к chat_id не
    подтверждён (см. комментарий в on_bot_started), а у callback event.get_ids()
    даёт chat_id надёжно. Поэтому регистрация — по нажатию кнопки (payload
    'subscribe:confirm' в on_callback), где chat_id берётся проверенным путём."""
    builder = InlineKeyboardBuilder()
    builder.row(CallbackButton(text="✅ Подписать этот чат", payload="subscribe:confirm"))
    await event.message.answer(
        text=("Подписать этот чат на утренние напоминания (горящие дедлайны, "
              "непроведённые инструктажи)? Нажмите кнопку для подтверждения."),
        attachments=[builder.as_markup()],
    )


@dp.message_created(F.message.body.text == "/unsubscribe")
async def on_unsubscribe(event: MessageCreated):
    """Отписать текущий чат от рассылки. Тоже через кнопку — тот же надёжный путь
    к chat_id, что и при подписке."""
    builder = InlineKeyboardBuilder()
    builder.row(CallbackButton(text="🚫 Отписать этот чат", payload="subscribe:remove"))
    await event.message.answer(
        text="Отписать этот чат от утренних напоминаний? Нажмите кнопку для подтверждения.",
        attachments=[builder.as_markup()],
    )


async def _deliver_document_list(responder: "_Responder", employee_id: str) -> None:
    """Показывает загруженные документы работника кнопками. Скачивание — по нажатию (docget)."""
    with Session(engine) as session:
        employee = session.get(Employee, employee_id)
        if employee is None:
            await responder.send("Сотрудник не найден.")
            return
        full_name = employee.full_name

    present = _s3_list_for_employee(employee_id)
    available = [(st, SCAN_TYPES[st]) for st in SCAN_TYPES if (present.get(st) or {}).get("present")]
    if not available:
        # даже если документов нет — даём кнопку возврата к списку работников (не тупик)
        builder = InlineKeyboardBuilder()
        builder.row(CallbackButton(text="⬅️ Назад к списку", payload="menu:docpick"))
        await responder.send(
            text=f"У {full_name} нет загруженных документов.",
            attachments=[builder.as_markup()],
        )
        return

    builder = InlineKeyboardBuilder()
    for st, label in available:
        builder.row(CallbackButton(text=label[:60], payload=f"docget:{employee_id}:{st}"))
    builder.row(CallbackButton(text="⬅️ Назад к списку", payload="menu:docpick"))
    await responder.send(
        text=f"Документы: {full_name}\nВыберите документ для скачивания:",
        attachments=[builder.as_markup()],
    )


async def _send_employee_document(responder: "_Responder", employee_id: str, scan_type: str) -> None:
    """Скачивает документ из S3 и отправляет файлом в чат. Имя файла — ФИО_тип.ext."""
    if scan_type not in SCAN_TYPES:
        await responder.send("Неизвестный тип документа.")
        return
    with Session(engine) as session:
        employee = session.get(Employee, employee_id)
        if employee is None:
            await responder.send("Сотрудник не найден.")
            return
        full_name = employee.full_name

    try:
        data, ct = _s3_download(scan_type, employee_id)
    except RuntimeError as e:
        await responder.send(f"Не удалось получить документ: {e}")
        return

    # временный файл с осмысленным именем (ФИО_тип.ext) — MAX покажет его как имя вложения
    fio = (full_name or "работник").replace(" ", "_")
    type_name = SCAN_TYPES[scan_type].split("(")[0].strip().replace(" ", "_")
    ext = _ext_for(ct)
    fname = f"{fio}_{type_name}.{ext}"
    tmpdir = tempfile.mkdtemp()
    path = os.path.join(tmpdir, fname)
    try:
        with open(path, "wb") as f:
            f.write(data)
        await responder.send(
            text=f"Документ: {full_name} — {SCAN_TYPES[scan_type]}",
            attachments=[InputMedia(path=path)],
        )
    except Exception:
        log.exception("Не удалось отправить документ %s работника %s", scan_type, employee_id)
        await responder.send("Ошибка отправки файла. Попробуйте ещё раз.")
    finally:
        try:
            os.remove(path)
            os.rmdir(tmpdir)
        except Exception:
            pass


# _docx_to_pdf переехал в document_templates.docx_to_pdf (2026-07): тот же конвертер
# понадобился webforms.py для печати согласий из веба, а импортировать bot.py в webforms
# нельзя — потянет maxapi и второй диспетчер. Здесь оставлен псевдоним, чтобы не править
# все вызовы (_send_work_order_pdf, _send_ot_order_pdf, _send_instruction_journal_pdf,
# _send_consent_pdf). Функция синхронная и блокирующая — вызывать через asyncio.to_thread.
_docx_to_pdf = docx_to_pdf


async def _deliver_work_order_list(responder: "_Responder", scope: str) -> None:
    """Список нарядов-допусков кнопками. scope='active' — действующие, 'past' — архив
    (истёкший срок). Корзину (мягко удалённые) прорабу не показываем. Скачивание — wodl:<id>."""
    with Session(engine) as session:
        if scope == "past":
            orders = prod.get_past_work_orders(session)
            title = "🗄 Архив нарядов (срок истёк):"
            empty = "В архиве нет нарядов."
        else:
            orders = prod.get_active_work_orders(session)
            title = "🟢 Действующие наряды-допуски:"
            empty = "Действующих нарядов нет."
        # подписи собираем внутри сессии (обращаемся к полям, пока объекты привязаны)
        items = [
            (o.id,
             f"№{o.number} — {o.location} "
             f"({o.valid_from:%d.%m}–{o.valid_to:%d.%m.%Y})")
            for o in orders
        ]

    builder = InlineKeyboardBuilder()
    if not items:
        builder.row(CallbackButton(text="⬅️ Назад", payload="menu:section:workorders"))
        # show_menu редактирует текущее сообщение (список открывается всегда по кнопке),
        # чтобы не плодить параллельное меню секции. Fallback на send — внутри show_menu.
        await responder.show_menu(empty, [builder.as_markup()])
        return
    for oid, label in items:
        builder.row(CallbackButton(text=label[:60], payload=f"wodl:{oid}"))
    builder.row(CallbackButton(text="⬅️ Назад", payload="menu:section:workorders"))
    await responder.show_menu(title, [builder.as_markup()])


async def _send_work_order_pdf(responder: "_Responder", work_order_id: str) -> None:
    """Генерирует наряд-допуск (docx), конвертирует в PDF и отправляет файлом в чат.
    Ветвление генератора по work_type_id — как в webforms.work_order_print:
    есть тип работы → высотный бланк (782н), нет → общий."""
    tmp_dir = tempfile.mkdtemp(prefix="wo_")
    try:
        with Session(engine) as session:
            order = session.get(WorkOrder, work_order_id)
            if order is None or order.is_deleted:
                await responder.send("Наряд не найден или удалён.")
                return
            members = (
                session.query(WorkOrderMember)
                .filter_by(work_order_id=order.id)
                .all()
            )
            number = order.number
            try:
                if order.work_type_id:
                    docx_path = prod.generate_height_work_order_docx(
                        order, members, org_name=ORG_NAME, output_dir=tmp_dir)
                else:
                    docx_path = prod.generate_work_order_docx(
                        order, members, org_name=ORG_NAME, output_dir=tmp_dir)
            except Exception:
                log.exception("Не удалось сгенерировать docx наряда %s", work_order_id)
                await responder.send("Не удалось сформировать наряд. Попробуйте позже.")
                return

        try:
            pdf_path = await asyncio.to_thread(_docx_to_pdf, docx_path, tmp_dir)
        except RuntimeError as e:
            await responder.send(f"Не удалось получить PDF: {e}")
            return

        # имя файла для чата — с санацией небезопасных символов в номере (слэш ломал имя)
        safe_num = "".join(c if c not in '/\\:*?"<>|' else "-" for c in (number or "no"))
        nice_name = f"Наряд-допуск_№{safe_num}.pdf"
        nice_path = os.path.join(tmp_dir, nice_name)
        try:
            if nice_path != pdf_path:
                os.replace(pdf_path, nice_path)
            await responder.send(
                text=f"Наряд-допуск №{number}",
                attachments=[InputMedia(path=nice_path)],
            )
        except Exception:
            log.exception("Не удалось отправить PDF наряда %s", work_order_id)
            await responder.send("Ошибка отправки файла. Попробуйте ещё раз.")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# Иконки статуса приказа ОТ (из prod.get_ot_orders_status):
#   ready → есть запись и скан (готово); no_scan → заведён без скана (распечатать/подписать/
#   приложить скан); missing → записи нет вовсе (ещё не создан, печатать нечего).
_OT_STATUS_ICON = {"ready": "✅", "no_scan": "🖨", "missing": "⬜"}


async def _deliver_ot_orders_list(responder: "_Responder") -> None:
    """Плоский список обязательных приказов ОТ/ТБ со статусом (чтобы видеть, что ещё не закрыто).
    Тап печатает бланк (poget:<key>). missing печатать нельзя — нет номера/даты."""
    with Session(engine) as session:
        items = prod.get_ot_orders_status(session)
        rows = [(it["key"], it["status"], it["topic"]) for it in items]

    builder = InlineKeyboardBuilder()
    for key, status, topic in rows:
        icon = _OT_STATUS_ICON.get(status, "•")
        builder.row(CallbackButton(text=f"{icon} {topic}"[:60], payload=f"poget:{key}"))
    builder.row(CallbackButton(text="⬅️ Назад", payload="menu:main"))
    header = ("📋 Приказы ОТ/ТБ:\n"
              "✅ готово (со сканом) · 🖨 распечатать/подписать · ⬜ ещё не создан")
    await responder.show_menu(header, [builder.as_markup()])


async def _send_ot_order_pdf(responder: "_Responder", order_key: str) -> None:
    """Генерирует бланк приказа ОТ по ключу и отправляет PDF. Печать возможна только для
    заведённых приказов (есть номер и дата); missing → печатать нечего."""
    tmp_dir = tempfile.mkdtemp(prefix="ot_")
    try:
        with Session(engine) as session:
            rec = prod.get_order_by_key(session, order_key)
            if rec is None:
                await responder.send(
                    "Этот приказ ещё не заведён в реестре — печатать нечего. "
                    "Создайте его в вебе (нужен номер и дата), затем распечатаете здесь."
                )
                return
            number = rec.number
            order_date = rec.order_date
            try:
                docx_path = prod.generate_ot_order_docx(order_key, number, order_date, tmp_dir)
            except ValueError as e:
                await responder.send(f"Не удалось сформировать приказ: {e}")
                return
            except Exception:
                log.exception("Не удалось сгенерировать docx приказа %s", order_key)
                await responder.send("Не удалось сформировать приказ. Попробуйте позже.")
                return

        try:
            pdf_path = await asyncio.to_thread(_docx_to_pdf, docx_path, tmp_dir)
        except RuntimeError as e:
            await responder.send(f"Не удалось получить PDF: {e}")
            return

        safe_num = "".join(c if c not in '/\\:*?"<>|' else "-" for c in (number or "no"))
        nice_name = f"Приказ_№{safe_num}.pdf"
        nice_path = os.path.join(tmp_dir, nice_name)
        try:
            if nice_path != pdf_path:
                os.replace(pdf_path, nice_path)
            await responder.send(
                text=f"Приказ №{number}",
                attachments=[InputMedia(path=nice_path)],
            )
        except Exception:
            log.exception("Не удалось отправить PDF приказа %s", order_key)
            await responder.send("Ошибка отправки файла. Попробуйте ещё раз.")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# Типы инструктажей, доступные для журнала в боте (заводятся задним числом по датам приезда).
_INSTR_JOURNAL_TYPES = [
    InstructionType.INTRODUCTORY,
    InstructionType.PRIMARY_WORKPLACE,
]


async def _deliver_instr_journals_menu(responder: "_Responder") -> None:
    """Список журналов инструктажей по типам. Тап печатает журнал (ijget:<type>)."""
    builder = InlineKeyboardBuilder()
    for itype in _INSTR_JOURNAL_TYPES:
        label = prod.INSTRUCTION_LABELS.get(itype, itype.value)
        builder.row(CallbackButton(text=f"📓 {label}"[:60], payload=f"ijget:{itype.value}"))
    builder.row(CallbackButton(text="⬅️ Назад", payload="menu:main"))
    header = ("📓 Журналы инструктажей под подписи.\n"
              "Тап формирует журнал: незанумерованные записи получают номера, "
              "приходит полный журнал PDF.")
    await responder.show_menu(header, [builder.as_markup()])


async def _send_instruction_journal_pdf(responder: "_Responder", type_value: str) -> None:
    """Печать журнала инструктажей в бот. Присваивает номера новым записям (как веб —
    сквозная нумерация по типу), затем отдаёт ПОЛНЫЙ журнал PDF (не только свежую партию),
    чтобы на объекте была полная подписная ведомость без веба."""
    try:
        itype = InstructionType(type_value)
    except ValueError:
        await responder.send("Неизвестный тип инструктажа.")
        return

    tmp_dir = tempfile.mkdtemp(prefix="ij_")
    try:
        with Session(engine) as session:
            # 1) присвоить номера новым записям (no-op, если все уже пронумерованы)
            prod.print_new_journal_entries(session, itype)
            # 2) собрать полный журнал из пронумерованных записей
            entries = prod.get_journaled_instructions(session, itype)
            if not entries:
                label = prod.INSTRUCTION_LABELS.get(itype, itype.value)
                await responder.send(
                    f"По журналу «{label}» нет записей. Сначала заведите инструктажи "
                    f"(в вебе — «Заполнить всем», датируется по приезду), затем печать здесь."
                )
                return
            order_ref = prod.get_latest_order_ref(session)
            started_at = prod.get_journal_started_at(session, itype)
            label = prod.INSTRUCTION_LABELS.get(itype, itype.value)
            try:
                xlsx_path = prod.generate_instruction_journal_xlsx(
                    entries, itype, org_name=ORG_NAME, order_ref=order_ref,
                    started_at=started_at, output_dir=tmp_dir,
                )
            except Exception:
                log.exception("Не удалось сгенерировать журнал инструктажа %s", type_value)
                await responder.send("Не удалось сформировать журнал. Попробуйте позже.")
                return

        try:
            pdf_path = await asyncio.to_thread(_docx_to_pdf, xlsx_path, tmp_dir)
        except RuntimeError as e:
            await responder.send(f"Не удалось получить PDF: {e}")
            return

        safe_label = label.replace(" ", "_").replace("/", "-")
        nice_name = f"Журнал_{safe_label}.pdf"
        nice_path = os.path.join(tmp_dir, nice_name)
        try:
            if nice_path != pdf_path:
                os.replace(pdf_path, nice_path)
            await responder.send(
                text=f"Журнал инструктажа: {label} (записей: {len(entries)})",
                attachments=[InputMedia(path=nice_path)],
            )
        except Exception:
            log.exception("Не удалось отправить журнал инструктажа %s", type_value)
            await responder.send("Ошибка отправки файла. Попробуйте ещё раз.")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


async def _send_consent_pdf(responder: "_Responder", employee_id: str, operator: str) -> None:
    """Бланк согласия на обработку ПДн (152-ФЗ) под подпись. operator: 'tsm' | 'ip'."""
    op_label = {"tsm": "ТСМ", "ip": "ИП Буц"}.get(operator, operator)
    tmp_dir = tempfile.mkdtemp(prefix="consent_")
    try:
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await responder.send("Сотрудник не найден.")
                return
            full_name = employee.full_name
            try:
                # require_fields=False: печать бланка под подпись возможна даже при
                # незаполненных полях (прочерки + баннер «черновик») — блокировки по
                # полноте карточки для согласия здесь нет (см. document_templates.py).
                docx_path = generate_consent_docx(
                    employee, operator=operator, output_dir=tmp_dir, require_fields=False
                )
            except Exception:
                log.exception("Не удалось сгенерировать согласие %s (%s)", employee_id, operator)
                await responder.send("Не удалось сформировать бланк согласия. Попробуйте позже.")
                return

        try:
            pdf_path = await asyncio.to_thread(_docx_to_pdf, docx_path, tmp_dir)
        except RuntimeError as e:
            await responder.send(f"Не удалось получить PDF: {e}")
            return

        safe_name = "".join(c if c not in '/\\:*?"<>|' else "-" for c in (full_name or "работник"))
        nice_name = f"Согласие_ПДн_{op_label}_{safe_name}.pdf"
        nice_path = os.path.join(tmp_dir, nice_name)
        try:
            if nice_path != pdf_path:
                os.replace(pdf_path, nice_path)
            await responder.send(
                text=f"Согласие на обработку ПДн — {op_label}: {full_name}",
                attachments=[InputMedia(path=nice_path)],
            )
        except Exception:
            log.exception("Не удалось отправить согласие %s", employee_id)
            await responder.send("Ошибка отправки файла. Попробуйте ещё раз.")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@dp.message_callback()
async def on_callback(event: MessageCallback):
    """Единая точка входа для всех кнопок: главное меню (payload='menu:...') и выбор
    сотрудника из /incomplete (payload='empdate:<id>'). Префиксы нужны, чтобы не перепутать
    два типа кнопок — просто employee.id без префикса раньше не различался бы с меню."""
    payload = event.callback.payload
    if not payload:
        return

    responder = _Responder(event)

    # Кнопка «Подать заявку на регистрацию» (после неудачного /login с ненайденным
    # номером). Запускает тот же флоу, что /register: спрашиваем ФИО, потом телефон,
    # затем register_via_max создаёт PENDING-заявку и уведомляет админов.
    if payload == "register:start":
        _pending_forms[responder.user_id()] = {"state": "awaiting_register_name"}
        await responder.send("Регистрация. Введите ваше ФИО полностью (например: Иванов Пётр):")
        return
    if payload == "subscribe:confirm":
        chat_id, _ = event.get_ids()
        with Session(engine) as session:
            existing = (
                session.query(NotificationSubscriber)
                .filter_by(chat_id=str(chat_id))
                .first()
            )
            if existing is None:
                session.add(NotificationSubscriber(chat_id=str(chat_id)))
                session.commit()
                await responder.send("✅ Чат подписан. Утренние напоминания будут приходить сюда.")
            else:
                await responder.send("Этот чат уже подписан — напоминания приходят сюда.")
        return
    if payload == "subscribe:remove":
        chat_id, _ = event.get_ids()
        with Session(engine) as session:
            existing = (
                session.query(NotificationSubscriber)
                .filter_by(chat_id=str(chat_id))
                .first()
            )
            if existing is not None:
                session.delete(existing)
                session.commit()
                await responder.send("🚫 Чат отписан. Утренние напоминания сюда больше не приходят.")
            else:
                await responder.send("Этот чат не был подписан.")
        return

    # Навигация по разделам меню.
    if payload == "menu:main":
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        await responder.show_menu("Главное меню:", [_build_main_menu(role).as_markup()])
        return
    if payload.startswith("menu:section:"):
        section = payload.split(":", 2)[2]
        title = _SECTION_TITLES.get(section, "Раздел")
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        if section == "attention" and role not in ("kadrovik", "admin"):
            await responder.send("Этот раздел доступен только кадровику/админу.")
            return
        if section == "workorders" and role not in ("prorab", "kadrovik", "admin"):
            await responder.send("Наряды-допуски доступны только зарегистрированным пользователям. "
                                  "Выполните /login.")
            return
        if section == "otorders":
            if role not in ("prorab", "kadrovik", "admin"):
                await responder.send("Приказы доступны только зарегистрированным пользователям. "
                                      "Выполните /login.")
                return
            # плоский список приказов — сразу, без промежуточного подменю
            await _deliver_ot_orders_list(responder)
            return
        if section == "instrjournals":
            if role not in ("prorab", "kadrovik", "admin"):
                await responder.send("Журналы инструктажей доступны только зарегистрированным "
                                      "пользователям. Выполните /login.")
                return
            await _deliver_instr_journals_menu(responder)
            return
        await responder.show_menu(title, [_build_section_menu(section, role).as_markup()])
        return

    if payload == "menu:add_employee":
        await _start_add_employee_flow(responder)
        return

    # ================= ТАБЕЛЬ: Утро/Вечер/Межвахта/Отчёты (2026-07, слияние ботов) =================

    if payload == "menu:morning":
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        if role not in ("prorab", "kadrovik", "admin"):
            await responder.send("Разметка явки доступна только зарегистрированным пользователям. "
                                  "Выполните /login.")
            return
        with Session(engine) as session:
            never_marked = tabel.get_never_marked_employees(session)
        if never_marked:
            names = "\n".join(f"  • {e.full_name} (договор с {e.contract_date:%d.%m.%Y})"
                               for e in never_marked)
            await responder.send(
                f"⚠️ Оформлены, но по ним ещё НИ РАЗУ не было отметки явки:\n{names}\n\n"
                f"Возможно, забыли внести в утренний обход — проверьте."
            )
        await _deliver_picker(responder, "utroday", extra_button=_tabel_extra_button("utroday"))
        return

    if payload == "menu:evening":
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        if role not in ("prorab", "kadrovik", "admin"):
            await responder.send("Разметка явки доступна только зарегистрированным пользователям. "
                                  "Выполните /login.")
            return
        await _deliver_picker(responder, "eveningnight", extra_button=_tabel_extra_button("eveningnight"))
        return

    if payload == "menu:empaction":
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        if role not in ("prorab", "kadrovik", "admin"):
            await responder.send("Доступно только зарегистрированным пользователям. Выполните /login.")
            return
        await _deliver_picker(responder, "empaction")
        return

    if payload.startswith("empaction:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.answer(notification="Сотрудник не найден.")
                return
            full_name = employee.full_name
            day_code = tabel.get_day_slot(session, employee_id) or "—"
            night_code = tabel.get_night_slot(session, employee_id) or "—"
            rr = session.get(RotationReturn, employee_id)
            needs_clarification = rr is not None and rr.expected_return_date is None
        kb = InlineKeyboardBuilder()
        kb.row(CallbackButton(text="✅ Поставить явку (Д)", payload=f"empact_day:{employee_id}"))
        kb.row(CallbackButton(text="🧹 Очистить день", payload=f"empact_clrday:{employee_id}"))
        kb.row(CallbackButton(text="🧹 Очистить ночь", payload=f"empact_clrnight:{employee_id}"))
        if day_code == tabel.ROTATION:
            kb.row(CallbackButton(text="✈️ Отменить межвахту", payload=f"empact_clrrot:{employee_id}"))
        if needs_clarification:
            kb.row(CallbackButton(text="✈️ Уточнить дату возврата",
                                    payload=f"empact_clarify_rot:{employee_id}"))
        kb.row(CallbackButton(text="◀ Назад к списку", payload="menu:empaction"))
        extra_note = "\n⚠️ Дата возврата с межвахты не уточнена!" if needs_clarification else ""
        await responder.send(
            f"{full_name}\nСегодня: день={day_code}, ночь={night_code}{extra_note}\nЧто сделать?",
            attachments=[kb.as_markup()],
        )
        return

    if payload.startswith("empact_clarify_rot:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            full_name = employee.full_name if employee else "—"
        _pending_forms[responder.user_id()] = {
            "state": "awaiting_rotation_return_date",
            "employee_id": employee_id,
        }
        await responder.send(f"{full_name}: укажите дату ВОЗВРАТА на объект "
                              f"(формат ГГГГ-ММ-ДД):")
        return

    if payload.startswith("empact_day:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is not None:
                full_name = employee.full_name
                tabel.mark_day(session, employee, responder.user_id())
        await responder.send(f"✔ {full_name if employee else '—'}: явка (Д) проставлена.")
        return

    if payload.startswith("empact_clrday:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is not None:
                full_name = employee.full_name
                tabel.clear_day_slot(session, employee)
        await responder.send(f"✔ {full_name if employee else '—'}: дневной слот очищен.")
        return

    if payload.startswith("empact_clrnight:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is not None:
                full_name = employee.full_name
                tabel.clear_night_slot(session, employee)
        await responder.send(f"✔ {full_name if employee else '—'}: ночной слот очищен.")
        return

    if payload.startswith("empact_clrrot:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is not None:
                full_name = employee.full_name
                tabel.clear_rotation(session, employee)
        await responder.send(f"✔ {full_name if employee else '—'}: межвахта отменена "
                              f"(отметка и ожидание возврата сняты).")
        return

    if payload.startswith("utroday:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.answer(notification="Сотрудник не найден.")
                return
            conflict = tabel.check_day_conflict(session, employee)
            if conflict:
                kb = InlineKeyboardBuilder()
                kb.row(CallbackButton(text="✅ Всё равно день", payload=f"utroforce:{employee_id}"))
                kb.row(CallbackButton(text="✖ Отмена", payload="menu:morning"))
                await responder.send(f"⚠️ {conflict}\nВсё равно поставить день?",
                                      attachments=[kb.as_markup()])
                return
            rot_conflict = tabel.check_rotation_return_conflict(session, employee)
            if rot_conflict:
                _pending_forms[responder.user_id()] = {
                    "state": "awaiting_actual_return_date",
                    "employee_id": employee_id, "action": "day",
                }
                await responder.send(f"⚠️ {rot_conflict}\nУкажите дату фактического возврата "
                                      f"(обычно — сегодняшнее число, ГГГГ-ММ-ДД):")
                return
            tabel.mark_day(session, employee, responder.user_id())
        await _deliver_picker(responder, "utroday", edit=True,
                               extra_button=_tabel_extra_button("utroday"))
        return

    if payload.startswith("utroforce:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is not None:
                tabel.mark_day(session, employee, responder.user_id())
        await _deliver_picker(responder, "utroday",
                               extra_button=_tabel_extra_button("utroday"))
        return

    if payload.startswith("eveningnight:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.answer(notification="Сотрудник не найден.")
                return
            conflict = tabel.check_night_conflict(session, employee)
            if conflict:
                kb = InlineKeyboardBuilder()
                kb.row(CallbackButton(text="✅ Всё равно ночь", payload=f"eveningforce:{employee_id}"))
                kb.row(CallbackButton(text="✖ Отмена", payload="menu:evening"))
                await responder.send(f"⚠️ {conflict}\nВсё равно поставить ночь?",
                                      attachments=[kb.as_markup()])
                return
            rot_conflict = tabel.check_rotation_return_conflict(session, employee)
            if rot_conflict:
                _pending_forms[responder.user_id()] = {
                    "state": "awaiting_actual_return_date",
                    "employee_id": employee_id, "action": "night",
                }
                await responder.send(f"⚠️ {rot_conflict}\nУкажите дату фактического возврата "
                                      f"(обычно — сегодняшнее число, ГГГГ-ММ-ДД):")
                return
            tabel.mark_night(session, employee, responder.user_id())
        await _deliver_picker(responder, "eveningnight", edit=True,
                               extra_button=_tabel_extra_button("eveningnight"))
        return

    if payload.startswith("eveningforce:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is not None:
                tabel.mark_night(session, employee, responder.user_id())
        await _deliver_picker(responder, "eveningnight",
                               extra_button=_tabel_extra_button("eveningnight"))
        return

    if payload == "utro_done":
        with Session(engine) as session:
            remaining = tabel.get_unmarked_day(session)
            if not remaining:
                await responder.send("Все отмечены. Утро завершено.")
                return
            builder = InlineKeyboardBuilder()
            for e in remaining[:PICKER_PAGE_SIZE]:
                builder.row(CallbackButton(text=e.full_name[:60], payload=f"reasonpick:{e.id}"))
            builder.row(CallbackButton(text="✅ Завершить (остальным неявка)", payload="reason_finish"))
        await responder.send(f"Укажите причину отсутствия (осталось {len(remaining)}):",
                              attachments=[builder.as_markup()])
        return

    if payload.startswith("reasonpick:"):
        employee_id = payload.split(":", 1)[1]
        kb = InlineKeyboardBuilder()
        kb.row(
            CallbackButton(text="❌ Неявка", payload=f"reasoncode:{employee_id}:{tabel.ABSENT}"),
            CallbackButton(text="🤒 Больничный", payload=f"reasoncode:{employee_id}:{tabel.SICK}"),
        )
        kb.row(
            CallbackButton(text="✈️ Межвахта", payload=f"reasoncode:{employee_id}:{tabel.ROTATION}"),
            CallbackButton(text="📋 Мигр.учёт", payload=f"reasoncode:{employee_id}:{tabel.MIGR}"),
        )
        kb.row(CallbackButton(text="🏖 Выходной", payload=f"reasoncode:{employee_id}:{tabel.WEEKEND}"))
        await responder.send("Причина?", attachments=[kb.as_markup()])
        return

    if payload.startswith("reasoncode:") and payload.endswith(":force"):
        _, employee_id, code, _force = payload.split(":", 3)
        # Убираем сообщение «Причина?» с кнопками — оно отработало, чтобы не копилось в чате.
        try:
            await event.message.delete()
        except Exception:
            log.exception("reasoncode:force — не удалось удалить сообщение с кнопками причины")
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            full_name = employee.full_name if employee is not None else "—"
            if employee is not None:
                tabel.set_reason(session, employee, code, responder.user_id())
                today_count = tabel.count_migr_today(session)
                if today_count > tabel.MIGR_DAILY_THRESHOLD:
                    await responder.send(f"⚠️ Сегодня на мигр.учёте уже {today_count} человек "
                                          f"(порог {tabel.MIGR_DAILY_THRESHOLD}). Риск вопросов "
                                          f"от заказчика — проверьте обоснованность.")
        await responder.send(f"{full_name}: причина проставлена.")
        return

    if payload.startswith("reasoncode:"):
        _, employee_id, code = payload.split(":", 2)
        # Убираем сообщение «Причина?» с кнопками — оно отработало, чтобы не копилось в чате.
        try:
            await event.message.delete()
        except Exception:
            log.exception("reasoncode — не удалось удалить сообщение с кнопками причины")
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.answer(notification="Сотрудник не найден.")
                return
            full_name = employee.full_name

            if code == tabel.ROTATION:
                _pending_forms[responder.user_id()] = {
                    "state": "awaiting_rotation_return_date",
                    "employee_id": employee_id,
                }
                await responder.send(f"{full_name}: межвахта.\n"
                                      f"⚠️ Укажите дату ВОЗВРАТА на объект (когда он вернётся "
                                      f"к работе), а не дату отъезда.\nФормат: ГГГГ-ММ-ДД")
                return

            if code == tabel.MIGR:
                migr_warn = tabel.check_migr_after_rotation(session, employee)
                if migr_warn:
                    kb = InlineKeyboardBuilder()
                    kb.row(CallbackButton(text="✅ Всё равно МУ", payload=f"reasoncode:{employee_id}:{tabel.MIGR}:force"))
                    kb.row(CallbackButton(text="✖ Отмена", payload=f"reasonpick:{employee_id}"))
                    await responder.send(f"⚠️ {migr_warn}\nВсё равно поставить МУ?",
                                          attachments=[kb.as_markup()])
                    return
                tabel.set_reason(session, employee, code, responder.user_id())
                today_count = tabel.count_migr_today(session)
                if today_count > tabel.MIGR_DAILY_THRESHOLD:
                    await responder.send(f"⚠️ Сегодня на мигр.учёте уже {today_count} человек "
                                          f"(порог {tabel.MIGR_DAILY_THRESHOLD}). Риск вопросов "
                                          f"от заказчика — проверьте обоснованность.")
            else:
                tabel.set_reason(session, employee, code, responder.user_id())
        await responder.send(f"{full_name}: причина проставлена.")
        return

    if payload == "ack_never_marked":
        today_s = date.today().isoformat()
        with Session(engine) as session:
            flag = session.get(SystemFlag, _NEVER_MARKED_ACK_KEY)
            now = datetime.utcnow()
            if flag is None:
                flag = SystemFlag(key=_NEVER_MARKED_ACK_KEY, value=today_s, updated_at=now)
                session.add(flag)
            else:
                flag.value = today_s
                flag.updated_at = now
            session.commit()
        await responder.send("✔ Принято, сегодня повторно не пришлю. Если проблема не решена — "
                              "напомню завтра утром.")
        return

    if payload.startswith("rotconfirm:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            full_name = employee.full_name if employee else "—"
        await responder.send(f"✔ {full_name}: дата возврата подтверждена, изменений не требуется.")
        return

    if payload.startswith("rotextend:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            full_name = employee.full_name if employee else "—"
        _pending_forms[responder.user_id()] = {
            "state": "awaiting_rotation_extend_date",
            "employee_id": employee_id,
        }
        await responder.send(f"{full_name}: введите НОВУЮ дату возврата (ГГГГ-ММ-ДД):")
        return

    if payload.startswith("depart:"):
        _, employee_id, date_iso, departure_type = payload.split(":", 3)
        return_date = datetime.strptime(date_iso, "%Y-%m-%d").date()
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.answer(notification="Сотрудник не найден.")
                return
            full_name = employee.full_name
            flagged = tabel.set_rotation(session, employee, return_date, responder.user_id(),
                                          departure_type=departure_type)
        if flagged:
            await responder.send(f"✔ {full_name}: межвахта до {return_date:%d.%m.%Y}. "
                                  f"Данные направлены в отдел кадров.")
        else:
            await responder.send(f"✔ {full_name}: межвахта до {return_date:%d.%m.%Y}. "
                                  f"Напомню за 3 дня до возврата.")
        return

    if payload == "reason_finish":
        with Session(engine) as session:
            n = tabel.fill_unmarked_absent(session, responder.user_id())
        await responder.send(f"Утро завершено. Неявка проставлена: {n} чел.")
        return

    if payload == "evening_done":
        await responder.send("🌙 Вечерняя отметка завершена.")
        return

    if payload == "menu:tabel_today":
        with Session(engine) as session:
            s = tabel.day_summary(session)
        lines = [
            "Табель за сегодня:",
            f"☀️ День: {s['day']}   🌙 Ночь: {s['night']}   😴 Отдых: {s['rest']}",
            f"🤒 Больн.: {s['sick']}   ✈️ Межвахта: {s['rotation']}   ❌ Неявка: {s['absent']}",
            f"📋 Мигр.учёт: {s['migr']}",
        ]
        if s["absent_list"]:
            lines.append("\nОтсутствуют/особое:")
            for name, code in s["absent_list"]:
                lines.append(f"  • {name} — {code}")
        await responder.send("\n".join(lines))
        return

    if payload == "menu:report_changelog":
        # Урезанно: только заголовки и счётчик, без «было/исправлено» построчно —
        # полный текст в веб-отчёте (/reports/changelog).
        n_tabel = len(reports_data.CHANGELOG_TABEL)
        n_migbot = len(reports_data.CHANGELOG_MIGBOT)
        lines = [f"🐛 Журнал патчей: {n_tabel + n_migbot} находок ({n_tabel} ТабельБелокаменка, "
                 f"{n_migbot} слияние с миграционным учётом).\n"]
        for h, _problem, _fix in reports_data.CHANGELOG_TABEL + reports_data.CHANGELOG_MIGBOT:
            lines.append(f"  • {h}")
        lines.append("\nПолный текст (было/исправлено по каждому) — в веб-версии, раздел «Отчёты».")
        await responder.send("\n".join(lines))
        return

    if payload == "menu:report_monthly_problems":
        with Session(engine) as session:
            data = reports_data.get_monthly_problems_report(session)
        if not data["problems"]:
            await responder.send(f"📊 Проблемные за {data['month_label']}: никто не превысил пороги.")
            return
        lines = [f"📊 Проблемные за {data['month_label']} ({len(data['problems'])}):"]
        for p in data["problems"]:
            parts = []
            if p["absent_count"] >= data["absent_threshold"]:
                parts.append(f"неявок {p['absent_count']}")
            if p["weekend_count"] >= data["weekend_threshold"]:
                parts.append(f"выходных {p['weekend_count']}")
            lines.append(f"  • {p['name']} — {', '.join(parts)}")
        await responder.send("\n".join(lines))
        return

    if payload == "menu:report_obligations":
        # Урезанно: только сводка по типам/статусам, без детального списка
        # просроченных с карточками — тот список в веб-версии.
        with Session(engine) as session:
            data = reports_data.get_obligations_report(session)
        if not data["counts"]:
            await responder.send("📋 Обязательства: нет данных по активным сотрудникам.")
            return
        lines = ["📋 Обязательства по активным сотрудникам:"]
        for (type_val, status_val), count in sorted(data["counts"].items()):
            lines.append(f"  • {type_val} — {status_val}: {count}")
        n_overdue = len(data["overdue_details"])
        if n_overdue:
            lines.append(f"\n🚨 Просрочено: {n_overdue}. Детали (кто, дедлайн, карточка) — в веб-версии.")
        await responder.send("\n".join(lines))
        return

    if payload == "menu:report_activity":
        # Урезанно: топ-5 по количеству отметок, без полного списка.
        with Session(engine) as session:
            data = reports_data.get_activity_report(session)
        if not data["actors"]:
            await responder.send("🕵️ Активность: за месяц отметок ещё не было.")
            return
        lines = [f"🕵️ Активность за {data['month_start'].strftime('%B %Y')} "
                 f"(всего {data['total']} отметок), топ-5:"]
        for a in data["actors"][:5]:
            lines.append(f"  • {a['label']} — {a['count']}")
        if len(data["actors"]) > 5:
            lines.append(f"\n...и ещё {len(data['actors']) - 5}. Полный список — в веб-версии.")
        await responder.send("\n".join(lines))
        return

    if payload == "menu:rotation_flags":
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
            if role not in ("kadrovik", "admin"):
                await responder.send("Доступно только кадровику/админу.")
                return
            flags = tabel.list_flagged_rotations(session)
            if not flags:
                await responder.send("Нет открытых флагов межвахты.")
                return
            builder = InlineKeyboardBuilder()
            for rr in flags:
                builder.row(CallbackButton(text=rr.employee.full_name[:60],
                                            payload=f"rotflag:{rr.employee_id}"))
        await responder.send(f"Межвахта с открытыми обязательствами ({len(flags)}):",
                              attachments=[builder.as_markup()])
        return

    if payload == "menu:onboarding":
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
            if role not in ("kadrovik", "admin"):
                await responder.send("Доступно только кадровику/админу.")
                return
            onboarding = tabel.get_onboarding_employees(session)
            if not onboarding:
                await responder.send("Нет сотрудников на оформлении.")
                return
            lines = ["На оформлении (не попадают в табель, пока не начнётся договор):"]
            for e in onboarding:
                cd = e.contract_date.strftime("%d.%m.%Y") if e.contract_date else "не указана"
                lines.append(f"  • {e.full_name} — дата договора: {cd}")
        await responder.send("\n".join(lines))
        return

    if payload == "menu:invalid_contract_marks":
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
            if role not in ("kadrovik", "admin"):
                await responder.send("Доступно только кадровику/админу.")
                return
            invalid_marks = tabel.get_marks_without_valid_contract(session)
            if not invalid_marks:
                await responder.send("Нет случаев явки без действующего договора.")
                return
            lines = ["🚨 Явка без действующего договора:"]
            for item in invalid_marks:
                cd = item["contract_date"].strftime("%d.%m.%Y") if item["contract_date"] else "не указана"
                ced = item["contract_end_date"]
                reason = f"уволен {ced:%d.%m.%Y}" if ced else f"дата договора: {cd}"
                marks_str = ", ".join(f"{d:%d.%m}/{slot}={code}" for d, slot, code in item["marks"])
                lines.append(f"\n{item['name']} ({reason})\nОтметки: {marks_str}")
        await responder.send("\n".join(lines))
        return

    if payload == "menu:pending_rotation":
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
            if role not in ("kadrovik", "admin"):
                await responder.send("Доступно только кадровику/админу.")
                return
            pending = tabel.get_pending_clarification_rotations(session)
            if not pending:
                await responder.send("Нет межвахт без уточнённой даты возврата.")
                return
            names = "\n".join(f"  • {p['name']}" for p in pending)
        await responder.send(
            f"❓ Стоят на МЖ, но дата возврата НЕ уточнена ({len(pending)}):\n{names}\n\n"
            f"Попросите прораба уточнить через «☀️ Утро» → выбрать человека → "
            f"«✈️ Уточнить межвахту»."
        )
        return

    if payload.startswith("rotflag:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.answer(notification="Сотрудник не найден.")
                return
            obligations = tabel.get_open_obligations(session, employee_id)
        lines = [f"{employee.full_name} — открытые обязательства при уходе на межвахту:"]
        for o in obligations:
            lines.append(f"  • {o.type.value} — {o.status.value}, дедлайн {o.deadline_date:%d.%m.%Y}")
        kb = InlineKeyboardBuilder()
        kb.row(CallbackButton(text="✅ Разобрано", payload=f"rotflag_resolve:{employee_id}"))
        await responder.send("\n".join(lines), attachments=[kb.as_markup()])
        return

    if payload.startswith("rotflag_resolve:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            user = find_user_by_max_id(session, responder.user_id())
            reviewer_id = user.id if user else "unknown"
            ok = tabel.resolve_rotation_flag(session, employee_id, reviewer_id)
        await responder.send("Отмечено как разобранное." if ok else "Не удалось (флаг уже снят?).")
        return

    # ================= /конец табеля =================

    if payload == "menu:employees":
        await _deliver_employees_list(responder)
        return

    if payload == "menu:incomplete":
        await _deliver_picker(responder, "empdate", edit=True)
        return

    if payload.startswith("page:"):
        _, prefix, page_s = payload.split(":", 2)
        extra = _tabel_extra_button(prefix)
        # only_if_open убран: навигация листает по сообщению из колбэка (оно живое),
        # не сверяясь с _open_pickers. Иначе после рестарта бота (память пуста)
        # листание молча не срабатывало — кнопки ◀️▶️ «не работали».
        await _deliver_picker(responder, prefix, page=int(page_s), edit=True,
                               extra_button=extra)
        return

    if payload.startswith("empdate:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.answer(notification="Сотрудник не найден.")
                return
            full_name = employee.full_name

        _pending_forms[responder.user_id()] = {
            "state": "awaiting_entry_date_button",
            "employee_id": employee_id,
        }
        # Список НЕ удаляем и не трогаем — после ввода даты он обновится через edit
        # в _apply_entry_date, и сотрудник исчезнет из него сам собой (условие фильтра).
        await responder.send(f"Введите дату въезда для {full_name} в формате ГГГГ-ММ-ДД:")
        return

    if payload == "menu:contractdate":
        await _deliver_picker(responder, "contractdate", edit=True)
        return

    if payload.startswith("contractdate:"):
        employee_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.answer(notification="Сотрудник не найден.")
                return
            full_name = employee.full_name

        _pending_forms[responder.user_id()] = {
            "state": "awaiting_contract_date_button",
            "employee_id": employee_id,
        }
        await responder.send(f"Введите дату договора для {full_name} в формате ГГГГ-ММ-ДД:")
        return

    if payload == "menu:delete_employee":
        await _deliver_picker(responder, "delpick", edit=True)
        return

    if payload == "menu:docpick":
        await _deliver_picker(responder, "docpick", edit=True)
        return

    if payload == "menu:consent_print":
        # Печать согласия для любого сотрудника. Доступно всем ролям (решение Валерия),
        # как и docpick в этом же разделе. Генерация PDF от статуса согласия не зависит.
        await _deliver_picker(responder, "consentprint", edit=True)
        return

    if payload.startswith("consentprint:"):
        # Тап по сотруднику в списке печати → диалог только с кнопками печати бланка
        # (ТСМ / ИП Буц), БЕЗ «Подтвердить» — смена статуса согласия живёт в
        # «Ожидают согласия», чтобы её нельзя было случайно нажать на уже подтверждённых.
        employee_id = payload.split(":", 1)[1]
        await _deliver_consent_print_choice(responder, employee_id)
        return

    if payload.startswith("docpick:"):
        # выбран работник — показываем его загруженные документы кнопками
        employee_id = payload.split(":", 1)[1]
        await _deliver_document_list(responder, employee_id)
        return

    if payload.startswith("docget:"):
        # скачать конкретный документ: docget:<employee_id>:<scan_type>
        _, employee_id, scan_type = payload.split(":", 2)
        await _send_employee_document(responder, employee_id, scan_type)
        return

    if payload.startswith("menu:wolist:"):
        # список нарядов: menu:wolist:active | menu:wolist:past
        scope = payload.split(":", 2)[2]
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        if role not in ("prorab", "kadrovik", "admin"):
            await responder.send("Наряды-допуски доступны только зарегистрированным пользователям. "
                                  "Выполните /login.")
            return
        await _deliver_work_order_list(responder, scope)
        return

    if payload.startswith("wodl:"):
        # скачать наряд-допуск PDF: wodl:<work_order_id>
        work_order_id = payload.split(":", 1)[1]
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        if role not in ("prorab", "kadrovik", "admin"):
            await responder.send("Наряды-допуски доступны только зарегистрированным пользователям. "
                                  "Выполните /login.")
            return
        await _send_work_order_pdf(responder, work_order_id)
        return

    if payload.startswith("poget:"):
        # печать приказа ОТ: poget:<order_key>
        order_key = payload.split(":", 1)[1]
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        if role not in ("prorab", "kadrovik", "admin"):
            await responder.send("Приказы доступны только зарегистрированным пользователям. "
                                  "Выполните /login.")
            return
        await _send_ot_order_pdf(responder, order_key)
        return

    if payload.startswith("ijget:"):
        # печать журнала инструктажа: ijget:<instruction_type_value>
        type_value = payload.split(":", 1)[1]
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        if role not in ("prorab", "kadrovik", "admin"):
            await responder.send("Журналы инструктажей доступны только зарегистрированным "
                                  "пользователям. Выполните /login.")
            return
        await _send_instruction_journal_pdf(responder, type_value)
        return

    if payload.startswith("delpick:"):
        employee_id = payload.split(":", 1)[1]
        # Список НЕ удаляем — подтверждение отправляется отдельным сообщением поверх него.
        await _deliver_delete_confirmation(responder, employee_id)
        return

    if payload.startswith("delconfirm:"):
        employee_id = payload.split(":", 1)[1]
        await event.message.delete()  # убираем только диалог подтверждения
        await _execute_delete_employee(responder, employee_id)
        await _deliver_picker(responder, "delpick", edit=True, only_if_open=True)  # список обновится без этого сотрудника
        return

    if payload.startswith("cancel:"):
        prefix = payload.split(":", 1)[1]
        await event.message.delete()  # убираем диалог подтверждения; список не менялся, не трогаем
        return

    if payload.startswith("exitpicker:"):
        prefix = payload.split(":", 1)[1]
        _open_pickers.pop((responder.user_id(), prefix), None)
        # Раньше здесь было event.message.delete(): сообщение со списком удалялось, и
        # пользователь оставался в пустоте (тупик — приходилось снова слать /start).
        # Теперь редактируем это же сообщение в главное меню (как ветка payload=="menu:main"):
        # список согласий заменяется главным меню на месте. Нет пустого экрана и нет
        # висящих живых кнопок для остальных участников общего чата.
        with Session(engine) as session:
            role = _role_for_max_id(session, responder.user_id())
        await responder.show_menu("Главное меню:", [_build_main_menu(role).as_markup()])
        return

    if payload == "menu:pending_consent":
        await _deliver_picker(responder, "consentpick", edit=True)
        return

    if payload.startswith("consentpick:"):
        employee_id = payload.split(":", 1)[1]
        await _deliver_consent_confirmation(responder, employee_id)
        return

    if payload.startswith("consentblank:"):
        # бланк согласия: consentblank:<operator>:<employee_id>
        _, operator, employee_id = payload.split(":", 2)
        await _send_consent_pdf(responder, employee_id, operator)
        return

    if payload.startswith("consentconfirm:"):
        employee_id = payload.split(":", 1)[1]
        await event.message.delete()  # убираем только диалог подтверждения
        await _execute_consent_confirm_by_button(responder, employee_id)
        await _deliver_picker(responder, "consentpick", edit=True, only_if_open=True)  # список обновится без сотрудника
        return


async def _apply_entry_date(event, employee_id: str, entry_date) -> None:
    """Общая логика установки даты въезда — используется и командой /set_entry_date,
    и сценарием после клика по кнопке в /incomplete."""
    with Session(engine) as session:
        employee = session.get(Employee, employee_id.strip())
        if employee is None:
            await event.reply(text="Сотрудник с таким id не найден. Проверьте /incomplete.") \
                if isinstance(event, MessageCallback) else \
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

        full_name = employee.full_name

    text = f"Дата въезда для {full_name} установлена: {entry_date.strftime('%d.%m.%Y')}."
    if isinstance(event, MessageCallback):
        await event.reply(text=text)
    else:
        await event.message.answer(text)

    # Список "Без даты въезда" мог быть открыт (клик по кнопке) — обновляем его,
    # сотрудник уходит из списка сам за счёт фильтра entry_date.is_(None).
    await _deliver_picker(_Responder(event), "empdate", edit=True, only_if_open=True)


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

    await _apply_entry_date(event, employee_id, entry_date)


async def _apply_contract_date(event, employee_id: str, contract_date) -> None:
    """Общая логика установки даты договора — используется и командой /set_contract_date,
    и сценарием после клика по кнопке в списке 'Без даты договора'. Симметрично
    _apply_entry_date: контроль и обязательства (contract_notice, efs1_report) зависят
    от этого поля так же, как registration/medical_exam — от entry_date."""
    with Session(engine) as session:
        employee = session.get(Employee, employee_id.strip())
        if employee is None:
            await event.reply(text="Сотрудник с таким id не найден.") \
                if isinstance(event, MessageCallback) else \
                await event.message.answer("Сотрудник с таким id не найден.")
            return

        employee.contract_date = contract_date
        session.add(employee)
        session.commit()
        session.refresh(employee)

        if employee.consent_status == ConsentStatus.CONFIRMED:
            create_obligations_for_employee(session, employee)

        full_name = employee.full_name

    text = f"Дата договора для {full_name} установлена: {contract_date.strftime('%d.%m.%Y')}."
    if isinstance(event, MessageCallback):
        await event.reply(text=text)
    else:
        await event.message.answer(text)

    await _deliver_picker(_Responder(event), "contractdate", edit=True, only_if_open=True)


async def _handle_set_contract_date(event: MessageCreated, raw_text: str) -> None:
    parts = raw_text.split(maxsplit=2)
    if len(parts) != 3:
        await event.message.answer(
            "Формат: /set_contract_date <id> <ГГГГ-ММ-ДД>\n"
            "id сотрудника смотрите в списке 'Без даты договора'."
        )
        return

    _, employee_id, date_s = parts
    try:
        contract_date = datetime.strptime(date_s.strip(), "%Y-%m-%d").date()
    except ValueError:
        await event.message.answer("Не распознал дату. Формат: ГГГГ-ММ-ДД, например 2026-06-15.")
        return

    await _apply_contract_date(event, employee_id, contract_date)


async def _handle_send_document(
    event: MessageCreated,
    raw_text: str,
    generator_func,
    doc_label: str,
    check_func=None,
) -> None:
    """check_func: функция (employee) -> list[str] отсутствующих полей, напр.
    check_consent_fields / check_medical_referral_fields из document_templates.py.

    2026-07: раньше при незаполненных полях generate_*_docx поднимал ValueError, но
    этот блок ловил его через общий except Exception и показывал только "Проверьте
    логи" — точный список полей терялся. Теперь:
      - список отсутствующих полей запрашивается ЗАРАНЕЕ (check_func), не только через
        перехват исключения;
      - в обычном режиме (TEST_ALLOW_MISSING_FIELDS=false) при непустом списке
        генерация даже не запускается — кадровик сразу видит, чего не хватает;
      - ValueError от самого генератора (на случай прямого вызова без check_func,
        либо иной причины) перехватывается ОТДЕЛЬНО от прочих исключений, и его текст
        уходит в чат, а не проглатывается;
      - если сработал тестовый обход (TEST_ALLOW_MISSING_FIELDS=true и missing непуст),
        к сообщению с файлом добавляется текстовое предупреждение — это MAX-эквивалент
        HTML-баннера в веб-формах, потому что мессенджер не рендерит произвольный HTML
        в чате, только текст/файлы/кнопки."""
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

        missing = check_func(employee) if check_func else []
        if missing and not TEST_ALLOW_MISSING_FIELDS:
            await event.message.answer(
                f"Нельзя сгенерировать документ для {employee.full_name} — "
                f"не заполнены поля: {', '.join(missing)}. "
                f"Заполните их через веб-форму кадровика или командами "
                f"(/set_entry_date, /set_contract_date и т.п.) перед генерацией."
            )
            return

        try:
            path = generator_func(employee)
        except ValueError as e:
            # Точный текст ошибки от _require_fields — раньше тонул в except Exception ниже.
            await event.message.answer(str(e))
            return
        except Exception:
            log.exception("Не удалось сгенерировать документ (%s) для employee_id=%s", doc_label, employee_id)
            await event.message.answer(f"Не удалось сгенерировать документ ({doc_label}). Проверьте логи.")
            return

        full_name = employee.full_name

    warning = ""
    if missing and TEST_ALLOW_MISSING_FIELDS:
        warning = (
            f"\n\n⚠ ТЕСТОВЫЙ ЧЕРНОВИК — не заполнены поля: {', '.join(missing)}. "
            "Документ не имеет юридической силы, пока эти поля не указаны и документ "
            "не перегенерирован."
        )

    await event.message.answer(
        text=f"{doc_label.capitalize()} для {full_name}:{warning}",
        attachments=[InputMedia(path=path)],
    )


@dp.message_created(F.message.body.text)
async def on_text(event: MessageCreated):
    user_id = event.message.sender.user_id
    form = _pending_forms.get(user_id)

    # Вариант 3 (согласовано): при ЛЮБОМ текстовом действии обновляем chat_id
    # личного диалога у привязанного пользователя — чтобы у активного админа
    # chat_id заполнился сам, без спец. /login, и заявки на регистрацию доходили.
    # No-op, если пользователь не привязан или chat_id уже совпадает.
    _chat_id = getattr(event, "chat_id", None)
    if _chat_id is not None:
        try:
            with Session(engine) as session:
                set_max_chat_id(session, user_id, _chat_id)
        except Exception:
            log.exception("on_text: не удалось обновить max_chat_id")

    if form and form.get("state") == "awaiting_login_phone":
        phone = event.message.body.text.strip()
        _pending_forms.pop(user_id, None)
        with Session(engine) as session:
            ok, text = bind_max_account(session, phone, user_id, chat_id=_chat_id)
        await event.message.answer(text)
        # Номер не найден → предлагаем подать заявку на регистрацию (вместо тупика).
        if not ok and "нет в системе" in text:
            builder = InlineKeyboardBuilder()
            builder.row(CallbackButton(text="📝 Подать заявку на регистрацию",
                                       payload="register:start"))
            await event.message.answer(
                text="Хотите подать заявку? Админ проверит и откроет доступ.",
                attachments=[builder.as_markup()],
            )
        return

    if form and form.get("state") == "awaiting_register_name":
        name = event.message.body.text.strip()
        parts = name.split()
        if len(parts) < 2 or not all(all(ch.isalpha() or ch == "-" for ch in p) for p in parts):
            await event.message.answer(
                "⚠️ Похоже, это не ФИО. Введите минимум 2 слова кириллицей "
                "(Фамилия Имя), например: Иванов Пётр."
            )
            return  # состояние остаётся, ждём корректный ввод
        _pending_forms[user_id] = {"state": "awaiting_register_phone", "full_name": name}
        await event.message.answer("Теперь введите номер телефона:")
        return

    if form and form.get("state") == "awaiting_register_phone":
        phone = event.message.body.text.strip()
        full_name = form["full_name"]
        _pending_forms.pop(user_id, None)
        with Session(engine) as session:
            ok, text = register_via_max(session, full_name, phone, user_id, chat_id=_chat_id)
            # Уведомляем админов о новой заявке (только если создана новая заявка).
            if ok:
                admins = get_admins_with_chat(session)
        await event.message.answer(text)
        if ok:
            note = (f"📝 Новая заявка на регистрацию:\n"
                    f"ФИО: {full_name}\nТелефон: {phone}\n"
                    f"Одобрите и назначьте роль в веб-разделе пользователей.")
            for admin in admins:
                try:
                    await bot.send_message(chat_id=admin.max_chat_id, text=note)
                except Exception:
                    log.exception("register: не удалось уведомить админа chat_id=%s", admin.max_chat_id)
        return

    if form and form.get("state") == "awaiting_actual_return_date":
        date_s = event.message.body.text.strip()
        try:
            actual_date = datetime.strptime(date_s, "%Y-%m-%d").date()
        except ValueError:
            await event.message.answer("Не распознал дату. Формат: ГГГГ-ММ-ДД, например 2026-07-10.")
            return
        if actual_date > date.today():
            await event.message.answer(
                "⚠️ Дата в будущем — нужна дата ФАКТИЧЕСКОГО возврата (обычно сегодняшнее "
                "число). Введите ещё раз (ГГГГ-ММ-ДД):"
            )
            return
        employee_id = form["employee_id"]
        action = form["action"]
        _pending_forms.pop(user_id, None)
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.message.answer("Сотрудник не найден.")
                return
            full_name = employee.full_name
            if action == "night":
                tabel.mark_night(session, employee, user_id)
                slot_label = "Ночь"
            else:
                tabel.mark_day(session, employee, user_id)
                slot_label = "День"
            departure_type = tabel.apply_rotation_return(session, employee, actual_date)
        extra = ""
        if departure_type == tabel.DEPARTURE_ABROAD:
            extra = "\nПостановка на учёт по возврату из-за границы создана (кадровику)."
        elif departure_type == tabel.DEPARTURE_DOMESTIC:
            extra = "\nАдрес пребывания обновлён, обязательства пересчитаны (кадровику)."
        await event.message.answer(
            f"✔ {full_name}: фактический возврат с межвахты {date_s} "
            f"зафиксирован. {slot_label} проставлен(а).{extra}"
        )
        return

    if form and form.get("state") == "awaiting_rotation_return_date":
        date_s = event.message.body.text.strip()
        try:
            return_date = datetime.strptime(date_s, "%Y-%m-%d").date()
        except ValueError:
            await event.message.answer("Не распознал дату. Формат: ГГГГ-ММ-ДД, например 2026-07-20.")
            return
        if return_date <= date.today():
            await event.message.answer(
                "⚠️ Дата — сегодня или в прошлом.\nНужна дата ВОЗВРАТА на объект (когда "
                "сотрудник вернётся к работе), а не дата отъезда и не сегодняшнее число.\n"
                "Введите дату возврата ещё раз (ГГГГ-ММ-ДД):"
            )
            return
        employee_id = form["employee_id"]
        _pending_forms.pop(user_id, None)
        # Тип отбытия определяет, какое юридическое событие сработает при
        # фактическом возврате (см. tabel.apply_rotation_return) — не пишем в БД,
        # пока не выберут.
        payload_base = f"{employee_id}:{return_date.isoformat()}"
        kb = InlineKeyboardBuilder()
        kb.row(CallbackButton(text="🌍 За границу", payload=f"depart:{payload_base}:{tabel.DEPARTURE_ABROAD}"))
        kb.row(CallbackButton(text="🇷🇺 В РФ, но с площадки", payload=f"depart:{payload_base}:{tabel.DEPARTURE_DOMESTIC}"))
        kb.row(CallbackButton(text="🏠 Не выезжал", payload=f"depart:{payload_base}:{tabel.DEPARTURE_NONE}"))
        await event.message.answer(
            "Куда убыл на межвахту? (от этого зависит, что сработает при возврате — "
            "новая постановка на учёт, обновление адреса или ничего)",
            attachments=[kb.as_markup()],
        )
        return

    if form and form.get("state") == "awaiting_rotation_extend_date":
        date_s = event.message.body.text.strip()
        try:
            new_return_date = datetime.strptime(date_s, "%Y-%m-%d").date()
        except ValueError:
            await event.message.answer("Не распознал дату. Формат: ГГГГ-ММ-ДД, например 2026-07-25.")
            return
        if new_return_date <= date.today():
            await event.message.answer(
                "⚠️ Дата — сегодня или в прошлом. Нужна дата в будущем. "
                "Введите ещё раз (ГГГГ-ММ-ДД):"
            )
            return
        employee_id = form["employee_id"]
        _pending_forms.pop(user_id, None)
        with Session(engine) as session:
            employee = session.get(Employee, employee_id)
            if employee is None:
                await event.message.answer("Сотрудник не найден.")
                return
            full_name = employee.full_name
            flagged = tabel.extend_rotation(session, employee, new_return_date)
        if flagged:
            await event.message.answer(f"✔ {full_name}: межвахта продлена до {date_s}. "
                                        f"Данные направлены в отдел кадров.")
        else:
            await event.message.answer(f"✔ {full_name}: межвахта продлена до {date_s}. "
                                        f"Напомню за 3 дня до нового возврата.")
        return

    if form and form.get("state") == "awaiting_entry_date_button":
        date_s = event.message.body.text.strip()
        try:
            entry_date = datetime.strptime(date_s, "%Y-%m-%d").date()
        except ValueError:
            await event.message.answer("Не распознал дату. Формат: ГГГГ-ММ-ДД, например 2026-06-15.")
            return

        employee_id = form["employee_id"]
        _pending_forms.pop(user_id, None)
        await _apply_entry_date(event, employee_id, entry_date)
        return

    if form and form.get("state") == "awaiting_contract_date_button":
        date_s = event.message.body.text.strip()
        try:
            contract_date = datetime.strptime(date_s, "%Y-%m-%d").date()
        except ValueError:
            await event.message.answer("Не распознал дату. Формат: ГГГГ-ММ-ДД, например 2026-06-15.")
            return

        employee_id = form["employee_id"]
        _pending_forms.pop(user_id, None)
        await _apply_contract_date(event, employee_id, contract_date)
        return

    if form and form.get("state") == "awaiting_employee_data":
        raw = event.message.body.text
        parts = [p.strip() for p in raw.split(";")]
        if len(parts) != 6:
            await event.message.answer("Не распознал формат. Ожидается 6 полей через ';'.")
            return

        full_name, citizenship, entry_date_s, contract_date_s, language, phone = parts
        category = category_for_citizenship(citizenship)

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

        if raw_text.startswith("/set_contract_date"):
            await _handle_set_contract_date(event, raw_text)
            return

        if raw_text.startswith("/send_consent_doc"):
            # Бланк согласия печатается для любого сотрудника независимо от полноты
            # карточки: check_func не передаём (нет предварительной блокировки), а
            # генератору передаём require_fields=False (не бросает ValueError, ставит
            # прочерки и баннер «черновик»). Затрагивает только согласие.
            await _handle_send_document(
                event, raw_text,
                lambda emp: generate_consent_docx(emp, require_fields=False),
                "согласие на обработку ПД",
            )
            return

        if raw_text.startswith("/send_medical_referral"):
            await _handle_send_document(
                event, raw_text, generate_medical_referral_docx, "направление на медкомиссию",
                check_func=check_medical_referral_fields,
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
            .filter_by(employee_id=employee.id, type=ObligationType.MEDICAL_EXAM, is_current=True)
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


_NEVER_MARKED_ACK_KEY = "never_marked_ack_date"

_MAX_MESSAGE_LEN = 4000  # лимит длины сообщения MAX API


def _split_message(text: str, limit: int = _MAX_MESSAGE_LEN) -> list[str]:
    """Разбить длинный текст на части не длиннее limit символов, по границам
    строк (не рвём строку посреди). Для рассылок со списком, который может
    превысить лимит сообщения MAX (например, много непроведённых инструктажей)."""
    if len(text) <= limit:
        return [text]
    chunks: list[str] = []
    current = ""
    for line in text.split("\n"):
        # +1 на символ переноса, который вернём при склейке
        if current and len(current) + len(line) + 1 > limit:
            chunks.append(current)
            current = line
        else:
            current = f"{current}\n{line}" if current else line
    if current:
        chunks.append(current)
    return chunks


async def morning_job():
    """
    Единый утренний джоб в 9:00 МСК (2026-07, слияние двух прежних джобов —
    rotation_reminders_job и never_marked_job — плюс новая месячная проверка).
    Три части, каждая своим сообщением(-ями), но один запуск планировщика:

    1. Проблемные за месяц (неявки >= ABSENT_THRESHOLD ИЛИ выходные >=
       WEEKEND_THRESHOLD, см. tabel.get_monthly_problems) — просто текстом.
    2. Никогда не отмеченные (см. tabel.get_never_marked_employees) — с кнопкой
       "✅ Ознакомлен", через SystemFlag не шлётся повторно в тот же день.
    3. Возврат с межвахты в течение 3 дней (см. tabel.get_rotation_reminders) —
       С КНОПКАМИ подтвердить/продлить, ТОЛЬКО если человек не встретился в
       пункте 2 (никогда не отмеченные и "возвращается с межвахты" — взаимно
       исключающие случаи на практике, но проверка на дубль сделана явно,
       как просили — "если про неё ещё не упомянуто").

    0. Плюс (2026-07) первая строка отчёта — общее количество активных по
       списку (tabel.get_active_employees), для контекста ко всем остальным
       цифрам ниже.

    Рассылка во ВСЕ чаты NotificationSubscriber безопасна для всех трёх пунктов —
    это работа прораба (вести отметки), не чувствительные детали обязательств
    кадровика (та утечка обсуждалась и НЕ применяется здесь, см. старый
    docstring rotation_reminders_job)."""
    today = date.today()
    today_s = today.isoformat()
    try:
        with Session(engine) as session:
            active_count = len(tabel.get_active_employees(session))
            monthly_problems = tabel.get_monthly_problems(session)
            flag = session.get(SystemFlag, _NEVER_MARKED_ACK_KEY)
            already_acked_today = flag is not None and flag.value == today_s
            never_marked = [] if already_acked_today else tabel.get_never_marked_employees(session)
            reminders = tabel.get_rotation_reminders(session, days_before=3)
            chat_ids = [row.chat_id for row in session.query(NotificationSubscriber).all()]
    except Exception:
        log.exception("morning_job: не удалось прочитать данные")
        return

    if not chat_ids:
        return

    # 0. Общее количество активных по списку.
    for chat_id in chat_ids:
        try:
            await bot.send_message(
                chat_id=chat_id,
                text=f"📋 Табель на {today:%d.%m.%Y}: активных по списку — {active_count}."
            )
        except Exception:
            log.exception("morning_job: не удалось отправить количество активных в chat_id=%s", chat_id)

    # 0.5. Подсказка по командам бота — напоминание раз в день, не нужно листать
    # старую переписку в поисках, как что называется.
    commands_hint = (
        "ℹ️ Команды бота:\n"
        "/login — привязать этот MAX к своей учётной записи (если уже зарегистрированы)\n"
        "/register — зарегистрироваться с нуля прямо здесь, если аккаунта ещё нет\n"
        "/confirm <код> — привязать MAX по коду, полученному при регистрации на сайте\n\n"
        "Разметка явки (Утро/Вечер), отчёты и остальное — через меню, кнопки появятся "
        "после /login или /register."
    )
    for chat_id in chat_ids:
        try:
            await bot.send_message(chat_id=chat_id, text=commands_hint)
        except Exception:
            log.exception("morning_job: не удалось отправить подсказку команд в chat_id=%s", chat_id)

    # 1. Месячные проблемные.
    if monthly_problems:
        lines = ["📊 За текущий месяц накопились неявки/выходные, требующие внимания:"]
        for p in monthly_problems:
            details = []
            if p["absent_count"] >= tabel.ABSENT_THRESHOLD:
                details.append(f"неявок: {p['absent_count']}")
            if p["weekend_count"] >= tabel.WEEKEND_THRESHOLD:
                details.append(f"выходных: {p['weekend_count']}")
            lines.append(f"  • {p['name']} — {', '.join(details)}")
        lines.append("\nℹ️ Это сообщение будет приходить каждое утро, пока показатели "
                      "не опустятся ниже порога — само не исчезнет.")
        text = "\n".join(lines)
        for chat_id in chat_ids:
            try:
                await bot.send_message(chat_id=chat_id, text=text)
            except Exception:
                log.exception("morning_job: не удалось отправить проблемных в chat_id=%s", chat_id)

    # 2. Никогда не отмеченные.
    never_marked_ids = set()
    if never_marked:
        never_marked_ids = {e.id for e in never_marked}
        names = "\n".join(f"  • {e.full_name} (договор с {e.contract_date:%d.%m.%Y})"
                           for e in never_marked)
        text = (f"⚠️ Оформлены, но по ним ещё НИ РАЗУ не было отметки явки:\n{names}\n\n"
                f"Возможно, забыли внести в утренний обход — проверьте и отметьте.\n\n"
                f"ℹ️ Если нажмёте «Ознакомлен» — сегодня повторно не пришлю, но если "
                f"проблема не решена, напомню снова завтра утром.")
        kb = InlineKeyboardBuilder()
        kb.row(CallbackButton(text="✅ Ознакомлен", payload="ack_never_marked"))
        for chat_id in chat_ids:
            try:
                await bot.send_message(chat_id=chat_id, text=text, attachments=[kb.as_markup()])
            except Exception:
                log.exception("morning_job: не удалось отправить never_marked в chat_id=%s", chat_id)

    # 3. Возврат с межвахты — пропускаем тех, кто уже упомянут в пункте 2.
    for r in reminders:
        if r["employee_id"] in never_marked_ids:
            continue
        kb = InlineKeyboardBuilder()
        kb.row(CallbackButton(text="✅ Подтверждаю возврат", payload=f"rotconfirm:{r['employee_id']}"))
        kb.row(CallbackButton(text="📅 Продлить межвахту", payload=f"rotextend:{r['employee_id']}"))
        text = (f"⏰ {r['name']}: ожидаемый возврат с межвахты — "
                f"{r['return_date']:%d.%m.%Y}.\nПодтвердите или продлите:\n\n"
                f"ℹ️ Будет приходить ежедневно, пока дата не пройдёт — "
                f"«Подтверждаю» не останавливает напоминание, только «Продлить» "
                f"сдвигает дату дальше.")
        for chat_id in chat_ids:
            try:
                await bot.send_message(chat_id=chat_id, text=text, attachments=[kb.as_markup()])
            except Exception:
                log.exception("morning_job: не удалось отправить напоминание в chat_id=%s", chat_id)

    # 4. Уточнить дату возврата с межвахты (2026-07) — заглушки без даты
    # (см. tabel.get_pending_clarification_rotations). Это адресовано ПРОРАБУ
    # (он вводит дату), не кадровику — тот минимально работает в MAX, его
    # версия этой же задачи теперь в веб-дашборде (см. webforms.py). Кнопка
    # ведёт прямо в тот же флоу ввода даты, что и "🧹 Действия с сотрудником".
    try:
        with Session(engine) as session:
            pending_rotation = tabel.get_pending_clarification_rotations(session)
    except Exception:
        log.exception("morning_job: не удалось прочитать pending_rotation")
        pending_rotation = []

    for item in pending_rotation:
        kb = InlineKeyboardBuilder()
        kb.row(CallbackButton(text="✈️ Уточнить дату возврата",
                                payload=f"empact_clarify_rot:{item['employee_id']}"))
        text = (f"❓ {item['name']}: стоит на МЖ, но дата возврата не уточнена.\n"
                f"Укажите, когда он вернётся к работе:\n\n"
                f"ℹ️ Это напоминание будет приходить каждое утро, пока не укажете дату — "
                f"после этого исчезнет само.")
        for chat_id in chat_ids:
            try:
                await bot.send_message(chat_id=chat_id, text=text, attachments=[kb.as_markup()])
            except Exception:
                log.exception("morning_job: не удалось отправить pending_rotation в chat_id=%s", chat_id)

    # 5. Непроведённые обязательные инструктажи (вводный / первичный на рабочем
    # месте) — дата начала работы наступила, инструктажа нет. Рассылка "до
    # устранения": приходит каждое утро всем подписчикам, пока инструктаж не
    # проведён (тогда сотрудник уходит из get_instruction_compliance_gaps сам).
    # Стадии помечены эмодзи, чтобы получатель различал давность: 🔴 критично
    # (> порога дней), ⚠️ просрочено (в пределах порога). Отдельным сообщением
    # от миграционных пунктов, чтобы не смешивать охрану труда и 109-ФЗ.
    try:
        with Session(engine) as session:
            instruction_gaps = prod.get_instruction_compliance_gaps(session)
    except Exception:
        log.exception("morning_job: не удалось прочитать instruction_gaps")
        instruction_gaps = []

    if instruction_gaps:
        critical = [g for g in instruction_gaps if g["stage"] == "critical"]
        overdue = [g for g in instruction_gaps if g["stage"] == "overdue"]
        lines = ["🦺 Не проведены обязательные инструктажи (дата начала работы уже прошла):"]
        if critical:
            lines.append("\n🔴 КРИТИЧНО (просрочка большая):")
            for g in critical:
                lines.append(
                    f"  • {g['name']} — {g['type_label']} "
                    f"(с {g['start_date']:%d.%m.%Y}, {g['days_overdue']} дн.)"
                )
        if overdue:
            lines.append("\n⚠️ Просрочено:")
            for g in overdue:
                lines.append(
                    f"  • {g['name']} — {g['type_label']} "
                    f"(с {g['start_date']:%d.%m.%Y}, {g['days_overdue']} дн.)"
                )
        lines.append(
            "\nℹ️ Провести инструктаж и внести в журнал — в веб-разделе «Производство → "
            "Инструктажи». Это сообщение будет приходить каждое утро, пока инструктаж "
            "не проведён — само не исчезнет."
        )
        # MAX ограничивает сообщение 4000 символов — при большом списке режем на части.
        text = "\n".join(lines)
        for chunk in _split_message(text):
            for chat_id in chat_ids:
                try:
                    await bot.send_message(chat_id=chat_id, text=chunk)
                except Exception:
                    log.exception("morning_job: не удалось отправить instruction_gaps в chat_id=%s", chat_id)


async def main():
    # Планировщик вынесен в отдельный модуль scheduler.py (2026-07). Задачи:
    # morning_job (эта, бот-рассылка) + погода из production.update_weather.
    # Запускается здесь, в бот-процессе — единственном always-on (нет дублей).
    from scheduler import start_scheduler
    start_scheduler(engine, morning_job=morning_job)
    while True:
      try:
        await dp.start_polling(bot)
        log.warning("start_polling exited without error; restarting in 5s")
      except asyncio.CancelledError:
        raise
      except Exception:
        log.exception("main: polling crashed; restarting in 5s")
      await asyncio.sleep(5)

if __name__ == "__main__":
    asyncio.run(main())
