"""
production.py — наряды-допуски, инструктажи, удостоверения по профессиям
(2026-07). Отдельный модуль от миграционного учёта — своя доменная область
(охрана труда на производстве), минимально связан с остальной системой:
bot.py/webforms.py дают только пункты меню/ссылки на функции отсюда.

Реализовано "тестовым режимом" — в отдельном файле, чтобы можно было
дорабатывать/переделывать, не трогая стабильный код миграционного учёта
и табеля.
"""

from datetime import date, datetime, timedelta

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from sqlalchemy import select
from sqlalchemy.orm import Session

from models import (
    Brigade,
    BrigadeMember,
    Certificate,
    Employee,
    Instruction,
    InstructionType,
    InternalOrder,
    OrderCategory,
    WorkOrder,
    WorkOrderDailyAdmission,
    WorkOrderMember,
    WorkOrderStatus,
)

# Порог "скоро истекает" для удостоверений — по аналогии с обязательствами
# миграционного учёта. [Предполагаю] 30 дней — не согласовано явно, разумный
# дефолт, легко поменять одной константой.
CERTIFICATE_EXPIRY_WARNING_DAYS = 30


# ================= Бригады =================

def create_brigade(session: Session, name: str, member_employee_ids: list[str]) -> Brigade:
    brigade = Brigade(name=name)
    session.add(brigade)
    session.flush()
    for employee_id in member_employee_ids:
        session.add(BrigadeMember(brigade_id=brigade.id, employee_id=employee_id))
    session.commit()
    return brigade


def get_brigades(session: Session) -> list[Brigade]:
    return session.query(Brigade).order_by(Brigade.name).all()


def get_brigade_member_ids(session: Session, brigade_id: str) -> list[str]:
    return [
        m.employee_id for m in
        session.query(BrigadeMember).filter_by(brigade_id=brigade_id).all()
    ]


def update_brigade_members(session: Session, brigade_id: str, member_employee_ids: list[str]) -> bool:
    """Заменяет состав целиком (не история — просто текущий список, см. docstring
    Brigade в models.py)."""
    brigade = session.get(Brigade, brigade_id)
    if brigade is None:
        return False
    session.query(BrigadeMember).filter_by(brigade_id=brigade_id).delete()
    for employee_id in member_employee_ids:
        session.add(BrigadeMember(brigade_id=brigade_id, employee_id=employee_id))
    session.commit()
    return True


def delete_brigade(session: Session, brigade_id: str) -> bool:
    brigade = session.get(Brigade, brigade_id)
    if brigade is None:
        return False
    session.delete(brigade)
    session.commit()
    return True


# ================= Наряды-допуски =================

def create_work_order(
    session: Session, number: str, work_description: str, location: str,
    responsible_supervisor_id: str, responsible_executor_id: str, issued_by: str,
    valid_from: date, valid_to: date, member_employee_ids: list[str],
    subdivision: str | None = None, materials: str | None = None, tools: str | None = None,
    equipment: str | None = None, special_machinery: str | None = None,
    technological_card_ref: str | None = None, safety_systems: str | None = None,
    special_conditions: str | None = None,
) -> WorkOrder:
    order = WorkOrder(
        number=number,
        subdivision=subdivision,
        work_description=work_description,
        location=location,
        responsible_supervisor_id=responsible_supervisor_id,
        responsible_executor_id=responsible_executor_id,
        issued_by=issued_by,
        valid_from=valid_from,
        valid_to=valid_to,
        status=WorkOrderStatus.ACTIVE,
        materials=materials,
        tools=tools,
        equipment=equipment,
        special_machinery=special_machinery,
        technological_card_ref=technological_card_ref,
        safety_systems=safety_systems,
        special_conditions=special_conditions,
    )
    session.add(order)
    session.flush()  # получить order.id до commit

    for employee_id in member_employee_ids:
        session.add(WorkOrderMember(work_order_id=order.id, employee_id=employee_id))

    # Заготовка строк "Ежедневный допуск к работе" на каждый день периода действия —
    # по образцу реального бланка (там наряд на 9 дней = 9 отдельных строк допуска).
    # Пустые (briefing_time/completion_time=NULL) — заполняются по факту каждый день,
    # см. record_daily_admission/record_daily_completion ниже.
    day = valid_from
    while day <= valid_to:
        session.add(WorkOrderDailyAdmission(work_order_id=order.id, admission_date=day))
        day += timedelta(days=1)

    session.commit()
    return order


def sign_work_order_member(session: Session, work_order_id: str, employee_id: str) -> bool:
    """Подтверждение ознакомления членом бригады (подпись)."""
    member = (
        session.query(WorkOrderMember)
        .filter_by(work_order_id=work_order_id, employee_id=employee_id)
        .first()
    )
    if member is None:
        return False
    member.signed_at = datetime.utcnow()
    session.add(member)
    session.commit()
    return True


def get_daily_admissions(session: Session, work_order_id: str) -> list[WorkOrderDailyAdmission]:
    return (
        session.query(WorkOrderDailyAdmission)
        .filter_by(work_order_id=work_order_id)
        .order_by(WorkOrderDailyAdmission.admission_date)
        .all()
    )


def record_daily_briefing(session: Session, work_order_id: str, admission_date: date,
                           confirmed_by: str) -> bool:
    """Целевой инструктаж на конкретный день выдан — фиксирует время и кто подтвердил
    (ответственный руководитель работ). Одна строка на день, см. WorkOrderDailyAdmission."""
    row = (
        session.query(WorkOrderDailyAdmission)
        .filter_by(work_order_id=work_order_id, admission_date=admission_date)
        .first()
    )
    if row is None:
        return False
    row.briefing_time = datetime.utcnow()
    row.briefing_confirmed_by = confirmed_by
    session.add(row)
    session.commit()
    return True


def record_daily_completion(session: Session, work_order_id: str, admission_date: date,
                             confirmed_by: str) -> bool:
    """Работа за конкретный день закончена, место убрано — фиксирует время и кто
    подтвердил (ответственный исполнитель работ)."""
    row = (
        session.query(WorkOrderDailyAdmission)
        .filter_by(work_order_id=work_order_id, admission_date=admission_date)
        .first()
    )
    if row is None:
        return False
    row.completion_time = datetime.utcnow()
    row.completion_confirmed_by = confirmed_by
    session.add(row)
    session.commit()
    return True


def get_active_work_orders(session: Session) -> list[WorkOrder]:
    today = date.today()
    return (
        session.query(WorkOrder)
        .filter(WorkOrder.status == WorkOrderStatus.ACTIVE)
        .filter(WorkOrder.valid_to >= today)
        .order_by(WorkOrder.valid_from)
        .all()
    )


def close_work_order(session: Session, work_order_id: str) -> bool:
    order = session.get(WorkOrder, work_order_id)
    if order is None:
        return False
    order.status = WorkOrderStatus.CLOSED
    session.add(order)
    session.commit()
    return True


# ================= Инструктажи =================

INSTRUCTION_LABELS = {
    InstructionType.INTRODUCTORY: "Вводный",
    InstructionType.PRIMARY_WORKPLACE: "Первичный на рабочем месте",
    InstructionType.REPEATED: "Повторный",
    InstructionType.UNSCHEDULED: "Внеплановый",
    InstructionType.TARGETED: "Целевой",
}

# Строк на страницу при допечатке журнала — [Предполагаю] не согласовано явно,
# разумный дефолт под таблицу А4 с таким набором граф. Легко поменять.
JOURNAL_ROWS_PER_PAGE = 20


def get_employees_needing_introductory(session: Session) -> list[Employee]:
    """Активные сотрудники, у кого известна дата начала работы (дата договора,
    а если её ещё нет — дата въезда), но вводного инструктажа ещё нет ни одного."""
    existing_ids = {
        i.employee_id for i in
        session.query(Instruction).filter_by(type=InstructionType.INTRODUCTORY).all()
    }
    employees = (
        session.query(Employee)
        .filter(Employee.contract_end_date.is_(None))
        .filter((Employee.contract_date.isnot(None)) | (Employee.entry_date.isnot(None)))
        .all()
    )
    return [e for e in employees if e.id not in existing_ids]


def auto_create_introductory_instructions(session: Session, conducted_by: str) -> list[Instruction]:
    """Заводит вводный инструктаж для ВСЕХ сотрудников разом — по договорённости
    "заполнять всеми сотрудниками с разделением по дате начала работы". Дата
    самого инструктажа = дата начала работы конкретного человека (дата договора,
    а если пусто — дата въезда), НЕ сегодняшняя дата — так порядок строк в
    журнале при печати совпадёт с реальной хронологией приёма, а не с датой,
    когда кто-то нажал кнопку в системе.

    2026-07: защита от гонки (двойное нажатие кнопки создавало дубли — нашли по
    факту на первой же реальной распечатке, см. журнал патчей). Коммит ПО ОДНОМУ
    сотруднику, а не одним общим commit() в конце — если где-то между чтением
    списка и записью появился дубль (например, из-за параллельного запроса),
    unique-индекс в БД (employee_id WHERE type='introductory') отклонит именно
    эту одну вставку, не обрушив всю пачку остальных."""
    from sqlalchemy.exc import IntegrityError

    employees = get_employees_needing_introductory(session)
    created = []
    for e in employees:
        start_date = e.contract_date or e.entry_date
        instr = Instruction(
            employee_id=e.id,
            type=InstructionType.INTRODUCTORY,
            conducted_by=conducted_by,
            conducted_at=datetime.combine(start_date, datetime.min.time()),
        )
        session.add(instr)
        try:
            session.commit()
            created.append(instr)
        except IntegrityError:
            session.rollback()  # уже есть (гонка/повторное нажатие) — пропускаем, не падаем
    return created


def get_unprinted_instructions(session: Session, instruction_type: InstructionType) -> list[Instruction]:
    return (
        session.query(Instruction)
        .filter_by(type=instruction_type, printed_at=None)
        .order_by(Instruction.conducted_at)
        .all()
    )


def get_last_journal_row_number(session: Session, instruction_type: InstructionType) -> int:
    last = (
        session.query(Instruction)
        .filter_by(type=instruction_type)
        .filter(Instruction.journal_row_number.isnot(None))
        .order_by(Instruction.journal_row_number.desc())
        .first()
    )
    return last.journal_row_number if last else 0


def print_new_journal_entries(session: Session, instruction_type: InstructionType) -> list[Instruction]:
    """"Допечатать новые записи" — присваивает номера строк последовательно, продолжая
    с последнего уже выданного номера (отдельная нумерация на каждый InstructionType —
    это разные физические журналы), помечает printed_at. Старые (уже напечатанные и
    подшитые) записи не трогает и не перепечатывает — см. договорённость в чате."""
    unprinted = get_unprinted_instructions(session, instruction_type)
    if not unprinted:
        return []
    next_num = get_last_journal_row_number(session, instruction_type) + 1
    now = datetime.utcnow()
    for instr in unprinted:
        instr.journal_row_number = next_num
        instr.printed_at = now
        next_num += 1
    session.commit()
    return unprinted


def create_instruction(
    session: Session, employee_id: str, instruction_type: InstructionType,
    conducted_by: str, topic: str | None = None, next_due_date: date | None = None,
) -> Instruction:
    instr = Instruction(
        employee_id=employee_id,
        type=instruction_type,
        topic=topic,
        conducted_by=conducted_by,
        next_due_date=next_due_date,
    )
    session.add(instr)
    session.commit()
    return instr


def confirm_instruction(session: Session, instruction_id: str) -> bool:
    """Работник подтвердил, что ознакомлен (подпись)."""
    instr = session.get(Instruction, instruction_id)
    if instr is None:
        return False
    instr.employee_confirmed_at = datetime.utcnow()
    session.add(instr)
    session.commit()
    return True


def get_instructions_for_employee(session: Session, employee_id: str) -> list[Instruction]:
    return (
        session.query(Instruction)
        .filter_by(employee_id=employee_id)
        .order_by(Instruction.conducted_at.desc())
        .all()
    )


def _set_cell(cell, text: str, size: int = 8, bold: bool = False) -> None:
    """Текст ячейки с уменьшенным шрифтом — чтобы длинные ФИО помещались в одну
    строку при печати, не переносились на две-три (см. договорённость)."""
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold


def get_journal_started_at(session: Session, instruction_type: InstructionType) -> date | None:
    """Дата самой ранней записи в журнале этого типа — для графы "Начат" на обложке.
    Не дата печати текущей партии, а дата первой КОГДА-ЛИБО занесённой записи."""
    first = (
        session.query(Instruction)
        .filter_by(type=instruction_type)
        .order_by(Instruction.conducted_at)
        .first()
    )
    return first.conducted_at.date() if first else None


def generate_instruction_journal_docx(
    instructions: list[Instruction], instruction_type: InstructionType,
    org_name: str, order_ref: str, journal_number: int = 1,
    started_at: date | None = None, output_dir: str = "/tmp",
) -> str:
    """
    Печать партии журнала инструктажей — по образцу рекомендуемой формы
    ГОСТ 12.0.004-2015 (Приложение А.4 вводный / А.5 на рабочем месте / А.6
    целевой). Только НОВЫЕ строки (см. print_new_journal_entries — вызывается
    ДО этой функции, здесь просто печать уже пронумерованных записей),
    довешенные прочерками до конца страницы (JOURNAL_ROWS_PER_PAGE) — довесок
    визуальный, не сохраняется как данные, следующая реальная запись получит
    следующий номер как ни в чём не бывало (см. договорённость в чате).

    2026-07 (второй заход, по замечаниям к первой распечатке):
    - journal_number — номер САМОЙ КНИГИ журнала (не строки внутри неё). Книга
      заполняется, закрывается, заводится новая с №2 — это [Предполагаю] пока
      НЕ автоматизировано (нет логики "книга заполнена, начать новую"), номер
      передаётся параметром, по умолчанию 1. Явно отмечаю ограничение, чтобы
      не выглядело как готовая функция ротации книг.
    - started_at — дата "Начат" на обложке (по факту делопроизводства — дата
      самой ранней записи в журнале ЭТОГО типа, не дата печати текущей партии).
      "Окончен" пока всегда пусто — закрытие книги вручную не реализовано.
    - Добавлены графы "Профессия (должность)" и "Подразделение" — были в
      реальной рекомендуемой форме ГОСТ, у нас их не было. В Employee нет
      полей профессии/подразделения — графы печатаются ПУСТЫМИ для заполнения
      от руки, не выдумываю значения.
    - Компактный шрифт (см. _set_cell) — чтобы строка помещалась в одну линию,
      не переносилась на 2-3 (замечание после первой распечатки).

    order_ref — ссылка на приказ (INTERNAL_ORDER_REF), печатается мелким
    шрифтом внизу страницы — см. договорённость про подсказку для проверяющего.
    """
    doc = Document()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    label = INSTRUCTION_LABELS.get(instruction_type, instruction_type.value)
    run = title.add_run(f"ЖУРНАЛ РЕГИСТРАЦИИ ИНСТРУКТАЖА ({label.upper()})")
    run.bold = True
    run.font.size = Pt(14)

    doc.add_paragraph(f"Организация: {org_name}")
    doc.add_paragraph(f"Журнал № {journal_number}")
    started_str = started_at.strftime("%d.%m.%Y") if started_at else "—"
    doc.add_paragraph(f"Начат: {started_str}          Окончен: —")

    headers = ["№", "Дата", "ФИО", "Год рожд.", "Профессия (должность)", "Подразделение",
               "Кто провёл", "Подпись инструкт.", "Подпись инструктирующего"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = False
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        _set_cell(hdr[i], h, size=8, bold=True)

    for instr in instructions:
        row = table.add_row().cells
        emp = instr.employee
        birth_year = str(emp.birth_date.year) if emp and emp.birth_date else ""
        _set_cell(row[0], str(instr.journal_row_number))
        _set_cell(row[1], instr.conducted_at.strftime("%d.%m.%Y"))
        _set_cell(row[2], emp.full_name if emp else "?")
        _set_cell(row[3], birth_year)
        _set_cell(row[4], (emp.position or "") if emp else "")  # из карточки, если заполнено
        _set_cell(row[5], (emp.subdivision or "") if emp else "")  # из карточки, если заполнено
        _set_cell(row[6], instr.conducted_by)
        _set_cell(row[7], "")  # подпись — от руки на распечатке
        _set_cell(row[8], "")

    # Довесок прочерками до конца страницы — визуальный, не данные (не создаёт
    # записей Instruction, следующая допечатка продолжит нумерацию с реального
    # следующего номера, не с номера довеска).
    remainder = len(instructions) % JOURNAL_ROWS_PER_PAGE
    if remainder != 0:
        pad_rows = JOURNAL_ROWS_PER_PAGE - remainder
        for _ in range(pad_rows):
            row = table.add_row().cells
            for cell in row:
                _set_cell(cell, "—")

    footer = doc.add_paragraph()
    footer_run = footer.add_run(
        f"Журнал ведётся по рекомендуемой форме ГОСТ 12.0.004-2015. Порядок регистрации "
        f"определён работодателем самостоятельно (п. 88 Правил №2464 от 24.12.2021; "
        f"разъяснение Роструда №15-2/В-1677 от 30.05.2022) — {order_ref}. "
        f"Незаполненные строки погашены прочерком. Подписи — собственноручные. "
        f"Сроки по закону: вводный — в день начала работы; первичный на рабочем месте — до "
        f"допуска к самостоятельной работе; повторный — не реже 1 раза в 6–12 мес.; "
        f"внеплановый/целевой — по факту события."
    )
    footer_run.font.size = Pt(7)
    footer_run.italic = True

    path = f"{output_dir}/journal_{instruction_type.value}_{datetime.utcnow():%Y%m%d%H%M%S}.docx"
    doc.save(path)
    return path


# ВРЕМЕННО (тест, 2026-07): очистка всех записей инструктажей, чтобы можно было
# заново проверить автозаполнение/допечатку с чистого листа. УДАЛИТЬ эту функцию
# и кнопку в webforms.py, когда тестирование закончится — см. пометку там же.
def test_clear_all_instructions(session: Session) -> int:
    count = session.query(Instruction).count()
    session.query(Instruction).delete()
    session.commit()
    return count


def get_due_instructions(session: Session, days_before: int = 7) -> list[dict]:
    """Повторные инструктажи, у которых next_due_date приближается или уже
    прошёл — для напоминания, по аналогии с get_rotation_reminders в tabel.py."""
    today = date.today()
    rows = (
        session.query(Instruction)
        .filter(Instruction.next_due_date.isnot(None))
        .all()
    )
    result = []
    for instr in rows:
        delta = (instr.next_due_date - today).days
        if delta <= days_before:
            emp = session.get(Employee, instr.employee_id)
            if emp is not None:
                result.append({
                    "employee_id": emp.id, "name": emp.full_name,
                    "due_date": instr.next_due_date, "overdue": delta < 0,
                    "instruction_id": instr.id,
                })
    return result


# ================= Удостоверения (корочки) =================

def create_certificate(
    session: Session, employee_id: str, profession: str,
    issued_by_org: str | None = None, issue_date: date | None = None,
    expiry_date: date | None = None, scan_key: str | None = None,
) -> Certificate:
    cert = Certificate(
        employee_id=employee_id,
        profession=profession,
        issued_by_org=issued_by_org,
        issue_date=issue_date,
        expiry_date=expiry_date,
        scan_key=scan_key,
    )
    session.add(cert)
    session.commit()
    return cert


def set_certificate_scan_key(session: Session, certificate_id: str, scan_key: str) -> None:
    cert = session.get(Certificate, certificate_id)
    if cert is not None:
        cert.scan_key = scan_key
        session.add(cert)
        session.commit()


def get_certificates_for_employee(session: Session, employee_id: str) -> list[Certificate]:
    return (
        session.query(Certificate)
        .filter_by(employee_id=employee_id)
        .order_by(Certificate.expiry_date.is_(None), Certificate.expiry_date)
        .all()
    )


def certificate_status(cert: Certificate, today: date | None = None) -> str:
    """'active' | 'expiring_soon' | 'expired' | 'no_expiry' — для бейджей в UI."""
    if cert.expiry_date is None:
        return "no_expiry"
    today = today or date.today()
    delta = (cert.expiry_date - today).days
    if delta < 0:
        return "expired"
    if delta <= CERTIFICATE_EXPIRY_WARNING_DAYS:
        return "expiring_soon"
    return "active"


def get_expiring_certificates(session: Session, days_before: int = CERTIFICATE_EXPIRY_WARNING_DAYS) -> list[dict]:
    """Удостоверения, срок которых истекает в пределах days_before дней (включая
    уже просроченные) — для напоминаний, та же схема, что и get_due_instructions."""
    today = date.today()
    rows = (
        session.query(Certificate)
        .filter(Certificate.expiry_date.isnot(None))
        .all()
    )
    result = []
    for cert in rows:
        delta = (cert.expiry_date - today).days
        if delta <= days_before:
            emp = session.get(Employee, cert.employee_id)
            if emp is not None:
                result.append({
                    "employee_id": emp.id, "name": emp.full_name,
                    "profession": cert.profession, "expiry_date": cert.expiry_date,
                    "overdue": delta < 0, "certificate_id": cert.id,
                })
    result.sort(key=lambda r: r["expiry_date"])
    return result


# ================= Печатный бланк наряда-допуска =================
# Печатный (не рукописный) наряд-допуск разрешён действующими Правилами по ОТ —
# "документ можно составлять на компьютере или от руки" (проверено, см. журнал
# патчей/обсуждение). Единственное, что должно остаться "живым" — подписи на
# распечатанном экземпляре, сам текст печатать можно.
#
# [Предполагаю] Это ОБЩИЙ бланк, не учитывает разницу форм по видам работ
# (электроустановки — приложение №7 к Приказу №903н, высотные — приложение №2
# и т.д. — формы РАЗНЫЕ и менять их содержание нельзя). Пока WorkOrder не имеет
# поля work_type, бланк один на все случаи — годится для общих/ремонтных работ,
# где унифицированной формы законом не установлено. Для регламентированных видов
# (электро-, высотные, огневые) нужен отдельный шаблон под конкретное приложение
# правил — это следующий шаг, не делать вид, что текущий бланк их закрывает.

def generate_work_order_docx(work_order: WorkOrder, members: list[WorkOrderMember],
                              org_name: str, output_dir: str = "/tmp") -> str:
    """
    2026-07: переписано по образцу реального бланка ООО «Промстроймонтаж»
    (наряд №25 на работы на высоте, фото прислал пользователь) — см. подробный
    docstring WorkOrder в models.py про структурные поправки. Разделы идут в
    том же порядке, что и в реальном бланке, для узнаваемости.
    """
    doc = Document()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(f"НАРЯД-ДОПУСК № {work_order.number}")
    run.bold = True
    run.font.size = Pt(16)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run("на производство работ повышенной опасности").italic = True

    doc.add_paragraph(f"Организация: {org_name}")
    if work_order.subdivision:
        doc.add_paragraph(f"Подразделение: {work_order.subdivision}")
    doc.add_paragraph(f"Выдан: «{work_order.created_at.strftime('%d')}» "
                       f"{work_order.created_at.strftime('%m.%Y')} г.")
    doc.add_paragraph(
        f"Действителен до: «{work_order.valid_to.strftime('%d')}» "
        f"{work_order.valid_to.strftime('%m.%Y')} г."
    )
    doc.add_paragraph(
        f"Ответственному руководителю работ: "
        f"{work_order.responsible_supervisor.full_name if work_order.responsible_supervisor else '—'}"
    )
    doc.add_paragraph(
        f"Ответственному исполнителю работ: "
        f"{work_order.responsible_executor.full_name if work_order.responsible_executor else '—'}"
    )
    doc.add_paragraph()
    doc.add_paragraph(f"На выполнение работ: {work_order.work_description}")
    doc.add_paragraph(f"Место выполнения работ: {work_order.location}")

    # Материалы/инструменты/приспособления/спецтехника — только если заполнены
    # (необязательные поля, не все наряды их требуют).
    for label, value in [
        ("Материалы", work_order.materials),
        ("Инструменты", work_order.tools),
        ("Приспособления", work_order.equipment),
        ("Спецтехника", work_order.special_machinery),
    ]:
        if value:
            doc.add_paragraph(f"{label}: {value}")

    if work_order.technological_card_ref:
        doc.add_paragraph()
        doc.add_paragraph(
            f"Работы производить в соответствии с требованиями технологической карты "
            f"({work_order.technological_card_ref})."
        )

    if work_order.safety_systems:
        doc.add_paragraph()
        doc.add_paragraph("Системы обеспечения безопасности работ:").bold = True
        doc.add_paragraph(work_order.safety_systems)

    if work_order.special_conditions:
        doc.add_paragraph()
        doc.add_paragraph("Особые условия проведения работ:").bold = True
        doc.add_paragraph(work_order.special_conditions)

    doc.add_paragraph()
    doc.add_paragraph("Состав исполнителей работ (члены бригады):").bold = True
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "№"
    hdr[1].text = "Фамилия, имя, отчество"
    hdr[2].text = "С условиями работ ознакомил, инструктаж провёл (подпись)"
    for i, member in enumerate(members, start=1):
        row = table.add_row().cells
        row[0].text = str(i)
        row[1].text = member.employee.full_name if member.employee else "?"
        row[2].text = ""  # печатный бланк — подпись ставится от руки на распечатке

    # Ежедневный допуск к работе — по одной строке на каждый день периода действия
    # наряда (см. WorkOrderDailyAdmission). Пустые графы — заполняются от руки/по
    # факту при подтверждении в системе (record_daily_briefing/completion).
    doc.add_paragraph()
    doc.add_paragraph("Ежедневный допуск к работе:").bold = True
    daily_table = doc.add_table(rows=1, cols=4)
    daily_table.style = "Table Grid"
    dhdr = daily_table.rows[0].cells
    dhdr[0].text = "Дата"
    dhdr[1].text = "Целевой инструктаж выдан (дата, время, подпись руководителя)"
    dhdr[2].text = "Работа закончена, место убрано (дата, время, подпись исполнителя)"
    dhdr[3].text = ""
    day = work_order.valid_from
    while day <= work_order.valid_to:
        drow = daily_table.add_row().cells
        drow[0].text = day.strftime("%d.%m.%Y")
        drow[1].text = ""
        drow[2].text = ""
        drow[3].text = ""
        day += timedelta(days=1)

    doc.add_paragraph()
    doc.add_paragraph("Подписи:")
    doc.add_paragraph("Наряд выдал: _________________________  (подпись, расшифровка)")
    doc.add_paragraph("Ответственный руководитель работ: _________________________  (подпись, расшифровка)")
    doc.add_paragraph("Ответственный исполнитель работ: _________________________  (подпись, расшифровка)")

    path = f"{output_dir}/naryad_{work_order.number}_{work_order.id[:8]}.docx"
    doc.save(path)
    return path


# ================= Реестр приказов =================

def create_order(session: Session, number: str, order_date: date, topic: str,
                  category: OrderCategory = OrderCategory.OTHER,
                  note: str | None = None) -> InternalOrder:
    order = InternalOrder(number=number, order_date=order_date, topic=topic,
                           category=category, note=note)
    session.add(order)
    session.commit()
    return order


def get_orders(session: Session) -> list[InternalOrder]:
    return session.query(InternalOrder).order_by(InternalOrder.order_date.desc()).all()


def get_orders_by_category(session: Session, category: OrderCategory) -> list[InternalOrder]:
    return (
        session.query(InternalOrder)
        .filter_by(category=category)
        .order_by(InternalOrder.order_date.desc())
        .all()
    )


def set_order_scan_key(session: Session, order_id: str, scan_key: str) -> None:
    order = session.get(InternalOrder, order_id)
    if order is not None:
        order.scan_key = scan_key
        session.add(order)
        session.commit()


def get_latest_order_ref(session: Session) -> str:
    """
    Ссылка на актуальный приказ для футеров печатных бланков (наряд-допуск,
    журналы инструктажа) — INTERNAL_ORDER_REF. Берёт последний по дате приказ
    из реестра. Если реестр пуст — явная заглушка, чтобы не выглядело как
    настоящая ссылка на несуществующий документ (см. договорённость)."""
    orders = get_orders(session)
    if not orders:
        return "[Приказ не издан — заполнить в разделе «Приказы»]"
    latest = orders[0]
    return f"Приказ № {latest.number} от {latest.order_date.strftime('%d.%m.%Y')}"
