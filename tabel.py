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
    """Активные — без даты увольнения (contract_end_date IS NULL), И договор уже
    начался (contract_date заполнен и <= сегодня). Кто заведён в базу, но договор
    ещё не наступил или не проставлен — считается "на оформлении", в табель не
    попадает (см. договорённость 2026-07). Такие видны кадровику отдельно —
    см. get_onboarding_employees ниже."""
    today = date.today()
    return (
        session.query(Employee)
        .filter(Employee.off_tabel.is_(False))
        .filter(Employee.contract_end_date.is_(None))
        .filter(Employee.contract_date.isnot(None))
        .filter(Employee.contract_date <= today)
        .order_by(Employee.full_name)
        .all()
    )


def get_never_marked_employees(session: Session) -> list[Employee]:
    """
    Активен по фильтру (договор действует), но НИ РАЗУ не было ни одной отметки
    в attendance_marks — не просто "сегодня не отмечен" (это нормально до конца
    утра), а вообще ни разу с начала договора. Сигнал прорабу: человек оформлен,
    но по нему табель не ведётся вовсе — возможно, забыли внести в утренний обход.
    """
    active = get_active_employees(session)
    marked_ids = {row[0] for row in session.query(AttendanceMark.employee_id).distinct().all()}
    return [e for e in active if e.id not in marked_ids]


def get_onboarding_employees(session: Session) -> list[Employee]:
    """Кто заведён в базу, но договор ещё не начался (нет даты, или дата в будущем) —
    "на оформлении". Не показывается в табеле, отдельная задача кадровику."""
    today = date.today()
    return (
        session.query(Employee)
        .filter(Employee.off_tabel.is_(False))
        .filter(Employee.contract_end_date.is_(None))
        .filter((Employee.contract_date.is_(None)) | (Employee.contract_date > today))
        .order_by(Employee.full_name)
        .all()
    )


def get_marks_without_valid_contract(session: Session) -> list[dict]:
    """
    СРОЧНАЯ проверка (2026-07): у кого есть отметки РЕАЛЬНОЙ явки (Д/НЧ — то есть
    человек физически выходил на работу) ВНЕ периода действия договора. Отличается
    от get_onboarding_employees: там просто "не в табеле по фильтру", здесь —
    уже случившийся факт выхода на работу без действующего договора.

    ВАЖНО: считается нарушением только явка ВНЕ границ договора, а не любая явка
    у уволенного/оформляемого сотрудника. Если уволен 03.07 — явка 02.07 (до
    увольнения) это НОРМА, а не нарушение; нарушение — явка ПОСЛЕ 03.07. Если
    contract_date в будущем — нарушение это явка ДО этой даты; если contract_date
    вообще не указана — любая явка нарушение (нет ни одной подтверждённой даты
    начала работы).

    Возможные причины: (а) договор оформлен, но дата ещё не внесена в систему —
    техническая недоработка данных; (б) человек реально работал без оформления —
    юридический риск. Разбираться должен кадровик, бот только сигнализирует.

    Возвращает [{"employee_id", "name", "contract_date", "contract_end_date",
    "marks": [(date, slot, code), ...]}] — только с реально нарушающими отметками.
    """
    today = date.today()
    candidates = (
        session.query(Employee)
        .filter(Employee.off_tabel.is_(False))
        .filter(
            (Employee.contract_date.is_(None))
            | (Employee.contract_date > today)
            | (Employee.contract_end_date.isnot(None))
        )
        .all()
    )
    result = []
    for e in candidates:
        all_marks = (
            session.query(AttendanceMark)
            .filter_by(employee_id=e.id)
            .filter(AttendanceMark.code.in_([DAY, NIGHT]))
            .order_by(AttendanceMark.mark_date)
            .all()
        )
        bad_marks = []
        for m in all_marks:
            if e.contract_end_date is not None and m.mark_date > e.contract_end_date:
                bad_marks.append(m)  # работал ПОСЛЕ увольнения
            elif e.contract_date is None:
                bad_marks.append(m)  # ни одной подтверждённой даты начала вообще
            elif e.contract_date is not None and m.mark_date < e.contract_date:
                bad_marks.append(m)  # работал ДО официальной даты начала

        if bad_marks:
            result.append({
                "employee_id": e.id,
                "name": e.full_name,
                "contract_date": e.contract_date,
                "contract_end_date": e.contract_end_date,
                "marks": [(m.mark_date, m.slot, m.code) for m in bad_marks],
            })
    return result


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
    """Активные, кого можно поставить в ночь: дневной слот != Д И ещё НЕ отмечены в ночь (НЧ).
    Раньше исключались только отработавшие день, а уже отмеченные ночью оставались в списке —
    из-за этого вечерний picker не обновлялся после простановки (отметка в БД шла, но человек
    не убирался, счётчик застревал, edit «того же самого» залипал). Теперь снимаем обоих."""
    mark_date = mark_date or date.today()
    active = get_active_employees(session)
    day_worked_ids = {
        row.employee_id for row in
        session.query(AttendanceMark.employee_id)
        .filter_by(mark_date=mark_date, slot="day", code=DAY)
        .all()
    }
    night_ids = {
        row.employee_id for row in
        session.query(AttendanceMark.employee_id)
        .filter_by(mark_date=mark_date, slot="night", code=NIGHT)
        .all()
    }
    return [e for e in active if e.id not in day_worked_ids and e.id not in night_ids]


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


def mark_sutki(session: Session, employee: Employee, created_by: str,
               mark_date: date | None = None) -> None:
    """Сутки — отработал и день, и ночь в одну дату (двойная смена). Через обычные
    Утро/Вечер в боте так не поставить (Вечер намеренно исключает уже отработавших
    день, см. get_not_worked_day) — это ручная правка, доступная только из веба."""
    mark_date = mark_date or date.today()
    _set_mark(session, employee, mark_date, "day", DAY, created_by)
    _set_mark(session, employee, mark_date, "night", NIGHT, created_by)


def set_reason(session: Session, employee: Employee, code: str, created_by: str,
               mark_date: date | None = None, drop_rotation: bool = False) -> None:
    """Причина отсутствия в дневной слот: Н/Б/МЖ/МУ/В.

    drop_rotation=True — вместе с кодом снимает ожидание возврата с межвахты
    (строку RotationReturn). Нужно при ПЕРЕЗАПИСИ ранее поставленной МЖ другим
    кодом: сам по себе _set_mark меняет только код дня, а RotationReturn остаётся,
    и бот продолжает ждать возврата и слать напоминания по человеку, который в
    табеле уже помечен неявкой/больничным. Расхождение тихое, поэтому решение
    принимает пользователь кнопкой, а не код молча."""
    mark_date = mark_date or date.today()
    _set_mark(session, employee, mark_date, "day", code, created_by)
    if drop_rotation and code != ROTATION:
        session.query(RotationReturn).filter_by(employee_id=employee.id).delete()
        session.commit()


def has_pending_rotation(session: Session, employee_id: str) -> bool:
    """Есть ли незакрытое ожидание возврата с межвахты. Используется перед
    перезаписью кода МЖ другой причиной."""
    return session.get(RotationReturn, employee_id) is not None


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


def clear_rotation(session: Session, employee: Employee, mark_date: date | None = None) -> None:
    """Полная отмена межвахты: снимает код МЖ за дату (обычно сегодня — день постановки)
    И удаляет строку RotationReturn (ожидание возврата + флаг кадровика), если была.
    Не трогает уже созданные Obligation от apply_rotation_return — если межвахта уже
    была закрыта фактическим возвратом, тот обязательство остаётся (это случившийся
    факт), отменять его отдельно руками."""
    mark_date = mark_date or date.today()
    clear_day_slot(session, employee, mark_date)
    session.query(RotationReturn).filter_by(employee_id=employee.id).delete()
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

DEPARTURE_ABROAD = "abroad"      # пересёк границу РФ — новая постановка на учёт
DEPARTURE_DOMESTIC = "domestic"  # остался в РФ, но покидал место пребывания
DEPARTURE_NONE = "none"          # физически не выезжал с площадки


def get_open_obligations(session: Session, employee_id: str) -> list[Obligation]:
    return (
        session.query(Obligation)
        .filter_by(employee_id=employee_id, is_current=True)
        .filter(Obligation.status.in_([ObligationStatus.PENDING, ObligationStatus.OVERDUE]))
        .order_by(Obligation.deadline_date)
        .all()
    )


def set_rotation(session: Session, employee: Employee, expected_return_date: date,
                  created_by: str, departure_type: str | None = None,
                  mark_date: date | None = None) -> bool:
    """
    Ставит МЖ прорабу без задержек. Если есть открытые обязательства — не
    блокирует, но создаёт/обновляет флаг для кадровика (rotation_returns.flagged).
    Возвращает True, если флаг был поднят (для сообщения прорабу "данные
    направлены в отдел кадров" — без деталей обязательств, см. договорённость).

    departure_type (DEPARTURE_ABROAD/DOMESTIC/NONE) — определяет, какое
    юридическое событие сработает при ФАКТИЧЕСКОМ возврате, см.
    apply_rotation_return().
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

    rr.departure_type = departure_type
    rr.flagged = flagged
    rr.flagged_at = now if flagged else None
    session.commit()
    return flagged


def apply_rotation_return(session: Session, employee: Employee, actual_return_date: date) -> str:
    """
    Вызывается при ФАКТИЧЕСКОМ возврате с межвахты (см. договорённость про три
    типа отбытия). Смотрит departure_type, сохранённый при постановке МЖ:

      DEPARTURE_ABROAD   -> create_registration_obligation_for_return (новая
                            постановка на учёт, БЕЗ дактилоскопии/медосмотра).
      DEPARTURE_DOMESTIC -> employee.address_since = дата возврата, дальше
                            отрабатывает уже существующее правило
                            REGISTRATION/address_since в deadlines.py — но
                            нужен полный create_obligations_for_employee,
                            чтобы это правило реально сработало (он же и
                            версионирует старую REGISTRATION).
      DEPARTURE_NONE / не указано -> ничего не создаём, регистрация не
                            прерывалась.

    Возвращает departure_type (или "none", если не был указан) — для текста
    сообщения человеку.
    """
    from obligations import create_obligations_for_employee, create_registration_obligation_for_return

    rr = session.get(RotationReturn, employee.id)
    departure_type = rr.departure_type if rr else None

    if departure_type == DEPARTURE_ABROAD:
        create_registration_obligation_for_return(session, employee, actual_return_date)
    elif departure_type == DEPARTURE_DOMESTIC:
        employee.address_since = actual_return_date
        session.add(employee)
        session.commit()
        create_obligations_for_employee(session, employee)
    # DEPARTURE_NONE или None — ничего не делаем.

    return departure_type or DEPARTURE_NONE


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


def extend_rotation(session: Session, employee: Employee, new_return_date: date) -> bool:
    """Продление межвахты (кнопка «Продлить» в напоминании за 3 дня) — обновляет только
    дату, сохраняет departure_type, каким он был указан при постановке. Пере-проверяет
    открытые обязательства заново (могли появиться/закрыться за время межвахты)."""
    rr = session.get(RotationReturn, employee.id)
    if rr is None:
        # Продлевать нечего — межвахта не была поставлена через обычный флоу.
        rr = RotationReturn(employee_id=employee.id, expected_return_date=new_return_date)
        session.add(rr)
    else:
        rr.expected_return_date = new_return_date
        rr.reviewed_at = None
        rr.reviewed_by = None

    open_obligations = get_open_obligations(session, employee.id)
    flagged = len(open_obligations) > 0
    rr.flagged = flagged
    rr.flagged_at = datetime.utcnow() if flagged else None
    session.commit()
    return flagged


def get_rotation_reminders(session: Session, days_before: int = 3) -> list[dict]:
    """Кто возвращается с межвахты в пределах days_before дней — для cron-напоминания.
    Пропускает записи с expected_return_date=NULL (заглушки "нужно уточнить") —
    те не про приближающийся срок, а про полное отсутствие даты, см.
    get_pending_clarification_rotations."""
    today = date.today()
    rows = session.query(RotationReturn).all()
    result = []
    for rr in rows:
        if rr.expected_return_date is None:
            continue
        delta = (rr.expected_return_date - today).days
        if 0 <= delta <= days_before:
            employee = session.get(Employee, rr.employee_id)
            if employee is not None:
                result.append({"employee_id": employee.id, "name": employee.full_name,
                                "return_date": rr.expected_return_date})
    return result


def get_pending_clarification_rotations(session: Session) -> list[dict]:
    """
    Кто стоит на МЖ (заглушка создана разовым скриптом при переносе истории —
    2026-07), но дата возврата НЕИЗВЕСТНА (expected_return_date=NULL). Нужна
    ДЛЯ ДВУХ адресатов:
      - прораб уточняет дату (флоу тот же, что при обычной постановке МЖ);
      - кадровик видит список, чтобы дёргать прораба, если тот не уточняет.

    Возвращает [{"employee_id", "name"}].
    """
    rows = (
        session.query(RotationReturn)
        .filter(RotationReturn.expected_return_date.is_(None))
        .all()
    )
    result = []
    for rr in rows:
        employee = session.get(Employee, rr.employee_id)
        if employee is not None:
            result.append({"employee_id": employee.id, "name": employee.full_name})
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


def day_reason_summary(session: Session, mark_date: date | None = None) -> list[tuple[str, str]]:
    """[(ФИО, код)] по всем активным, у кого дневной слот содержит причину отсутствия
    (Н/Б/МЖ/МУ/В), отсортировано по ФИО. Используется ботом для одного итогового
    сообщения-протокола вместо строки на каждое нажатие: в чате остаётся один
    проверяемый список, а не лента «причина проставлена» без указания причины."""
    mark_date = mark_date or date.today()
    active_ids = {e.id for e in get_active_employees(session)}
    if not active_ids:
        return []
    rows = (
        session.query(AttendanceMark)
        .filter_by(mark_date=mark_date, slot="day")
        .filter(AttendanceMark.employee_id.in_(active_ids))
        .filter(AttendanceMark.code.in_(REASON_CODES))
        .all()
    )
    return sorted(((r.employee_name_snap or "—", r.code) for r in rows), key=lambda t: t[0])


def get_month_codes(session: Session, year: int | None = None, month: int | None = None) -> dict:
    """
    Помесячная сетка для веба (аналог старой сетки Google Sheets табеля):
    {employee_id: {"name": ФИО, "codes": [код_день1, код_день2, ...]}}

    День+ночь объединяются в один код на дату — 'С' (сутки), если стоит и
    Д, и НЧ в один день; иначе код дня, если он есть; иначе НЧ, если стоит
    только ночь (день пуст/не Д); иначе пусто ("").
    """
    import calendar as _calendar

    today = date.today()
    year = year or today.year
    month = month or today.month
    days_in_month = _calendar.monthrange(year, month)[1]

    active = get_active_employees(session)
    active_ids = [e.id for e in active]

    marks = (
        session.query(AttendanceMark)
        .filter(AttendanceMark.employee_id.in_(active_ids))
        .filter(AttendanceMark.mark_date >= date(year, month, 1))
        .filter(AttendanceMark.mark_date <= date(year, month, days_in_month))
        .all()
    )
    # (employee_id, day_number, slot) -> code
    lookup = {}
    for m in marks:
        lookup[(m.employee_id, m.mark_date.day, m.slot)] = m.code

    result = {}
    for e in active:
        codes = []
        for d in range(1, days_in_month + 1):
            day_code = lookup.get((e.id, d, "day"))
            night_code = lookup.get((e.id, d, "night"))
            if day_code == DAY and night_code == NIGHT:
                codes.append("С")
            elif day_code:
                codes.append(day_code)
            elif night_code == NIGHT:
                codes.append(NIGHT)
            else:
                codes.append("")
        result[e.id] = {"name": e.full_name, "codes": codes}
    return result


# Пороги месячной проверки "проблемных" (см. get_monthly_problems). Оба порога
# названы явно в задаче: неявки от 2, выходные от 3.
ABSENT_THRESHOLD = 2
WEEKEND_THRESHOLD = 3


def get_monthly_problems(session: Session, year: int | None = None,
                          month: int | None = None) -> list[dict]:
    """
    Кто за текущий месяц накопил >= ABSENT_THRESHOLD неявок (Н) ИЛИ
    >= WEEKEND_THRESHOLD выходных (В). Возвращает
    [{"name", "absent_count", "weekend_count"}] только для тех, кто превысил
    хотя бы один порог.
    """
    today = date.today()
    year = year or today.year
    month = month or today.month

    active = get_active_employees(session)
    active_ids = [e.id for e in active]
    names_by_id = {e.id: e.full_name for e in active}

    import calendar as _calendar
    days_in_month = _calendar.monthrange(year, month)[1]

    rows = (
        session.query(AttendanceMark.employee_id, AttendanceMark.code)
        .filter(AttendanceMark.employee_id.in_(active_ids))
        .filter(AttendanceMark.slot == "day")
        .filter(AttendanceMark.mark_date >= date(year, month, 1))
        .filter(AttendanceMark.mark_date <= date(year, month, days_in_month))
        .filter(AttendanceMark.code.in_([ABSENT, WEEKEND]))
        .all()
    )

    counts = {}  # employee_id -> {"absent": n, "weekend": n}
    for employee_id, code in rows:
        c = counts.setdefault(employee_id, {"absent": 0, "weekend": 0})
        if code == ABSENT:
            c["absent"] += 1
        elif code == WEEKEND:
            c["weekend"] += 1

    result = []
    for employee_id, c in counts.items():
        if c["absent"] >= ABSENT_THRESHOLD or c["weekend"] >= WEEKEND_THRESHOLD:
            result.append({
                "name": names_by_id.get(employee_id, "?"),
                "absent_count": c["absent"],
                "weekend_count": c["weekend"],
            })
    result.sort(key=lambda r: r["name"])
    return result
