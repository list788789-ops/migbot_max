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
