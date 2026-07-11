
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

from sqlalchemy import select
from sqlalchemy.orm import Session

from models import (
    Certificate,
    Employee,
    Instruction,
    InstructionType,
    WorkOrder,
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
    responsible_employee_id: str, issued_by: str, valid_from: date, valid_to: date,
    member_employee_ids: list[str],
) -> WorkOrder:
    order = WorkOrder(
        number=number,
        work_description=work_description,
        location=location,
        responsible_employee_id=responsible_employee_id,
        issued_by=issued_by,
        valid_from=valid_from,
        valid_to=valid_to,
        status=WorkOrderStatus.ACTIVE,
    )
    session.add(order)
    session.flush()  # получить order.id до commit

    for employee_id in member_employee_ids:
        session.add(WorkOrderMember(work_order_id=order.id, employee_id=employee_id))

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
