"""
tabel.py — разметка явки (Утро/Вечер/Причины/Межвахта) поверх attendance_marks
и employees migbot. Замена sheets.py бота ТабельБелокаменка (2026-07,
слияние ботов) — та же бизнес-логика, источник данных теперь Postgres,
не Google Sheets.

Роли (см. UserRole в models.py): PRORAB размечает явку без ограничений
(узкое исключение из "PRORAB не пишет в БД" — решение зафиксировано в
models.py). KADROVIK/ADMIN дополнительно видят и разбирают флаги
"Требует внимания" (rotation_returns.flagged) — это отдельный модуль
(см. attention.py), не здесь.
"""

from datetime import date, datetime, timedelta

from sqlalchemy.orm import Session

from models import (
    AttendanceMark,
    Employee,
    Obligation,
    ObligationStatus,
    RotationReturn,
)

# Коды — те же буквы, что были в Google Sheets табеля, для непрерывности
# (люди уже привыкли к этим обозначениям).
DAY = "Д"          # работал день
REST = "О"         # отдых
SICK = "Б"         # больничный
ROTATION = "МЖ"    # межвахта
ABSENT = "Н"       # неявка
MIGR = "МУ"        # миграционный учёт
WEEKEND = "В"      # плановый выходной
NIGHT = "НЧ"       # работал ночь (только ночной слот)

DAY_CODES = [DAY, REST, SICK, ROTATION, ABSENT, MIGR, WEEKEND]
NIGHT_CODES = [NIGHT, REST]
REASON_CODES = [ABSENT, SICK, ROTATION, MIGR, WEEKEND]

MIGR_DAILY_THRESHOLD = 5  # порог одновременных МУ за день (см. договорённость в табеле)


# ================= АКТИВНЫЕ СОТРУДНИКИ =================

def get_active_employees(session: Session) -> list[Employee]:
    """Активные — без даты увольнения (contract_end_date IS NULL), см. договорённость."""
    return (
        session.query(Employee)
        .filter(Employee.contract_end_date.is_(None))
        .order_by(Employee.full_name)
        .all()
    )


# ================= ЧТЕНИЕ ОТМЕТОК =================

def _get_mark(session: Session, employee_id: str, mark_date: date, slot: str) -> str | None:
    row = (
        session.query(AttendanceMark)
        .filter_by(employee_id=employee_id, mark_date=mark_date, slot=slot)
        .first()
    )
    return row.code if row else None


def get_day_slot(session: Session, employee_id: str, mark_date: date | None = None) -> str | None:
    mark_date = mark_date or date.today()
    return _get_mark(session, employee_id, mark_date, "day")


def get_night_slot(session: Session, employee_id: str, mark_date: date | None = None) -> str | None:
    mark_date = mark_date or date.today()
    return _get_mark(session, employee_id, mark_date, "night")


def get_unmarked_day(session: Session, mark_date: date | None = None) -> list[Employee]:
    """Активные, у кого дневной слот ПУСТ (ещё не отмечены утром)."""
    mark_date = mark_date or date.today()
    active = get_active_employees(session)
    marked_ids = {
        row.employee_id for row in
        session.query(AttendanceMark.employee_id)
        .filter_by(mark_date=mark_date, slot="day")
        .all()
    }
    return [e for e in active if e.id not in marked_ids]


def get_marked_day(session: Session, mark_date: date | None = None) -> list[Employee]:
    """Активные, у кого дневной слот НЕ пуст (для «Очистить сотрудника»)."""
    mark_date = mark_date or date.today()
    active = get_active_employees(session)
    marked_ids = {
        row.employee_id for row in
        session.query(AttendanceMark.employee_id)
        .filter_by(mark_date=mark_date, slot="day")
        .all()
    }
    return [e for e in active if e.id in marked_ids]


def get_marked_night(session: Session, mark_date: date | None = None) -> list[Employee]:
    """Активные, у кого ночной слот = НЧ (для «Очистить сотрудника» вечером)."""
    mark_date = mark_date or date.today()
    active = get_active_employees(session)
    night_ids = {
        row.employee_id for row in
        session.query(AttendanceMark.employee_id)
        .filter_by(mark_date=mark_date, slot="night", code=NIGHT)
        .all()
    }
    return [e for e in active if e.id in night_ids]


def get_not_worked_day(session: Session, mark_date: date | None = None) -> list[Employee]:
    """Активные, кто НЕ работал днём (слот != Д) — их можно поставить в ночь."""
    mark_date = mark_date or date.today()
    active = get_active_employees(session)
    day_worked_ids = {
        row.employee_id for row in
        session.query(AttendanceMark.employee_id)
        .filter_by(mark_date=mark_date, slot="day", code=DAY)
        .all()
    }
    return [e for e in active if e.id not in day_worked_ids]


def get_night_rest(session: Session, mark_date: date | None = None) -> list[Employee]:
    """Кто ВЧЕРА работал ночь — им сегодня положен отдых днём."""
    mark_date = mark_date or date.today()
    yday = mark_date - timedelta(days=1)
    active = {e.id: e for e in get_active_employees(session)}
    night_ids = {
        row.employee_id for row in
        session.query(AttendanceMark.employee_id)
        .filter_by(mark_date=yday, slot="night", code=NIGHT)
        .all()
    }
    return [active[eid] for eid in night_ids if eid in active]


def morning_progress(session: Session, mark_date: date | None = None) -> dict:
    """Состояние утренней отметки: marked/unmarked/interrupted."""
    mark_date = mark_date or date.today()
    total = len(get_active_employees(session))
    unmarked = len(get_unmarked_day(session, mark_date))
    marked = total - unmarked
    return {"marked": marked, "unmarked": unmarked,
            "interrupted": marked > 0 and unmarked > 0}


# ================= ЗАПИСЬ ОТМЕТОК =================

def _set_mark(session: Session, employee: Employee, mark_date: date, slot: str,
              code: str, created_by: str) -> None:
    row = (
        session.query(AttendanceMark)
        .filter_by(employee_id=employee.id, mark_date=mark_date, slot=slot)
        .first()
    )
    now = datetime.utcnow()
    if row is None:
        row = AttendanceMark(
            employee_id=employee.id,
            employee_name_snap=employee.full_name,
            mark_date=mark_date,
            slot=slot,
            code=code,
            created_by=created_by,
            created_at=now,
            updated_at=now,
        )
        session.add(row)
    else:
        row.code = code
        row.employee_name_snap = employee.full_name
        row.updated_at = now
    session.commit()


def mark_day(session: Session, employee: Employee, created_by: str,
             mark_date: date | None = None) -> None:
    """Прораб отметил присутствующего днём: ДЕНЬ=Д."""
    mark_date = mark_date or date.today()
    _set_mark(session, employee, mark_date, "day", DAY, created_by)


def mark_night(session: Session, employee: Employee, created_by: str,
               mark_date: date | None = None) -> None:
    """Ночная смена: НОЧЬ=НЧ, ДЕНЬ=О (днём отдыхал)."""
    mark_date = mark_date or date.today()
    _set_mark(session, employee, mark_date, "day", REST, created_by)
    _set_mark(session, employee, mark_date, "night", NIGHT, created_by)


def set_reason(session: Session, employee: Employee, code: str, created_by: str,
               mark_date: date | None = None) -> None:
    """Причина отсутствия в дневной слот: Н/Б/МЖ/МУ/В."""
    mark_date = mark_date or date.today()
    _set_mark(session, employee, mark_date, "day", code, created_by)


def set_rest(session: Session, employee: Employee, created_by: str,
             mark_date: date | None = None) -> None:
    """Автоотдых с ночи: ДЕНЬ=О."""
    mark_date = mark_date or date.today()
    _set_mark(session, employee, mark_date, "day", REST, created_by)


def clear_day_slot(session: Session, employee: Employee, mark_date: date | None = None) -> None:
    mark_date = mark_date or date.today()
    session.query(AttendanceMark).filter_by(
        employee_id=employee.id, mark_date=mark_date, slot="day"
    ).delete()
    session.commit()


def clear_night_slot(session: Session, employee: Employee, mark_date: date | None = None) -> None:
    mark_date = mark_date or date.today()
    session.query(AttendanceMark).filter_by(
        employee_id=employee.id, mark_date=mark_date, slot="night"
    ).delete()
    session.commit()


def fill_unmarked_absent(session: Session, created_by: str,
                          mark_date: date | None = None) -> int:
    """Всем активным с пустым дневным слотом ставит Н. Возвращает число проставленных."""
    mark_date = mark_date or date.today()
    unmarked = get_unmarked_day(session, mark_date)
    for e in unmarked:
        _set_mark(session, e, mark_date, "day", ABSENT, created_by)
    return len(unmarked)


# ================= ПРОВЕРКИ / ПРЕДУПРЕЖДЕНИЯ =================

def check_day_conflict(session: Session, employee: Employee,
                        mark_date: date | None = None) -> str | None:
    """Если сегодня уже стоит ночь, или вчера была ночь — предупреждение."""
    mark_date = mark_date or date.today()
    if get_night_slot(session, employee.id, mark_date) == NIGHT:
        return f"{employee.full_name} уже отмечен в НОЧЬ за этот день."
    yday = mark_date - timedelta(days=1)
    if get_night_slot(session, employee.id, yday) == NIGHT:
        return f"{employee.full_name} вчера работал в НОЧЬ, положен отдых."
    return None


def check_night_conflict(session: Session, employee: Employee,
                          mark_date: date | None = None) -> str | None:
    mark_date = mark_date or date.today()
    if get_day_slot(session, employee.id, mark_date) == DAY:
        return f"{employee.full_name} уже отработал ДЕНЬ за эту дату."
    return None


def check_rotation_return_conflict(session: Session, employee: Employee,
                                    mark_date: date | None = None) -> str | None:
    """Вчера была МЖ, а сегодня ставят Д/Ночь напрямую — строгое предупреждение
    про нарушение миграционного законодательства (см. договорённость в табеле)."""
    mark_date = mark_date or date.today()
    yday = mark_date - timedelta(days=1)
    if get_day_slot(session, employee.id, yday) == ROTATION:
        return (
            f"ВНИМАНИЕ: {employee.full_name} был на межвахте. Возврат сотрудника "
            f"на объект без надлежащего уведомления и постановки на "
            f"миграционный учёт является нарушением миграционного "
            f"законодательства РФ."
        )
    return None


def check_migr_after_rotation(session: Session, employee: Employee,
                               mark_date: date | None = None) -> str | None:
    """Переход МЖ->МУ подозрителен (может маскировать неявку)."""
    mark_date = mark_date or date.today()
    yday = mark_date - timedelta(days=1)
    if get_day_slot(session, employee.id, yday) == ROTATION:
        return (f"{employee.full_name} вчера был на межвахте (МЖ). Переход МЖ→МУ "
                f"подозрителен — проверьте, не маскирует ли это неявку.")
    return None


def count_migr_today(session: Session, mark_date: date | None = None) -> int:
    mark_date = mark_date or date.today()
    active_ids = {e.id for e in get_active_employees(session)}
    rows = (
        session.query(AttendanceMark.employee_id)
        .filter_by(mark_date=mark_date, slot="day", code=MIGR)
        .all()
    )
    return len([r.employee_id for r in rows if r.employee_id in active_ids])


# ================= МЕЖВАХТА: ФЛАГ ДЛЯ КАДРОВИКА =================

def get_open_obligations(session: Session, employee_id: str) -> list[Obligation]:
    return (
        session.query(Obligation)
        .filter_by(employee_id=employee_id, is_current=True)
        .filter(Obligation.status.in_([ObligationStatus.PENDING, ObligationStatus.OVERDUE]))
        .order_by(Obligation.deadline_date)
        .all()
    )


def set_rotation(session: Session, employee: Employee, expected_return_date: date,
                  created_by: str, mark_date: date | None = None) -> bool:
    """
    Ставит МЖ прорабу без задержек. Если есть открытые обязательства — не
    блокирует, но создаёт/обновляет флаг для кадровика (rotation_returns.flagged).
    Возвращает True, если флаг был поднят (для сообщения прорабу "данные
    направлены в отдел кадров" — без деталей обязательств, см. договорённость).
    """
    mark_date = mark_date or date.today()
    set_reason(session, employee, ROTATION, created_by, mark_date)

    open_obligations = get_open_obligations(session, employee.id)
    flagged = len(open_obligations) > 0

    rr = session.get(RotationReturn, employee.id)
    now = datetime.utcnow()
    if rr is None:
        rr = RotationReturn(employee_id=employee.id, expected_return_date=expected_return_date)
        session.add(rr)
    else:
        rr.expected_return_date = expected_return_date
        # Новая межвахта — сбрасываем предыдущий разбор кадровика, если он был.
        rr.reviewed_at = None
        rr.reviewed_by = None

    rr.flagged = flagged
    rr.flagged_at = now if flagged else None
    session.commit()
    return flagged


def list_flagged_rotations(session: Session) -> list[RotationReturn]:
    """Для раздела «Требует внимания» — только KADROVIK/ADMIN (проверка роли —
    на уровне вызывающего кода в bot.py/webforms.py, не здесь)."""
    return (
        session.query(RotationReturn)
        .filter_by(flagged=True)
        .filter(RotationReturn.reviewed_at.is_(None))
        .all()
    )


def resolve_rotation_flag(session: Session, employee_id: str, reviewed_by_user_id: str) -> bool:
    """Кадровик разобрал флаг — явно, руками (не снимается автоматически при
    закрытии обязательств, см. docstring RotationReturn в models.py)."""
    rr = session.get(RotationReturn, employee_id)
    if rr is None or not rr.flagged:
        return False
    rr.reviewed_at = datetime.utcnow()
    rr.reviewed_by = reviewed_by_user_id
    session.commit()
    return True


def get_rotation_reminders(session: Session, days_before: int = 3) -> list[dict]:
    """Кто возвращается с межвахты в пределах days_before дней — для cron-напоминания."""
    today = date.today()
    rows = session.query(RotationReturn).all()
    result = []
    for rr in rows:
        delta = (rr.expected_return_date - today).days
        if 0 <= delta <= days_before:
            employee = session.get(Employee, rr.employee_id)
            if employee is not None:
                result.append({"name": employee.full_name,
                                "return_date": rr.expected_return_date})
    return result


# ================= СВОДКА ЗА ДЕНЬ =================

def day_summary(session: Session, mark_date: date | None = None) -> dict:
    mark_date = mark_date or date.today()
    active = get_active_employees(session)
    active_ids = {e.id for e in active}

    day_rows = (
        session.query(AttendanceMark)
        .filter_by(mark_date=mark_date, slot="day")
        .filter(AttendanceMark.employee_id.in_(active_ids))
        .all()
    )
    night_rows = (
        session.query(AttendanceMark)
        .filter_by(mark_date=mark_date, slot="night", code=NIGHT)
        .filter(AttendanceMark.employee_id.in_(active_ids))
        .all()
    )

    counts = {DAY: 0, REST: 0, SICK: 0, ROTATION: 0, ABSENT: 0, MIGR: 0, WEEKEND: 0}
    absent_list = []
    for row in day_rows:
        if row.code in counts:
            counts[row.code] += 1
        if row.code in (SICK, ROTATION, ABSENT, MIGR):
            absent_list.append((row.employee_name_snap, row.code))

    return {
        "day": counts[DAY], "night": len(night_rows), "rest": counts[REST],
        "sick": counts[SICK], "rotation": counts[ROTATION], "absent": counts[ABSENT],
        "migr": counts[MIGR], "absent_list": absent_list, "total": len(active),
    }
