from datetime import date
"""
obligations.py — создание Obligation (регистрация/медосмотр/уведомление МВД) и первого
RegistrationPeriod для сотрудника.

Вынесено из bot.py в отдельный модуль, чтобы webforms.py могло вызывать ТУ ЖЕ функцию,
не импортируя bot.py целиком. Импорт bot.py выполнил бы код верхнего уровня (Bot(),
Dispatcher(), load_dotenv()) и создал бы второй экземпляр MAX-бота внутри веб-процесса —
лишняя и опасная связка между двумя независимыми Railway-сервисами.

bot.py теперь должен импортировать эту функцию отсюда же (см. инструкцию внизу файла),
а не определять её у себя — иначе логика снова разойдётся в двух местах.

2026-07: версионирование обязательств. Смена места пребывания (address_since) —
отдельное юридическое событие по 109-ФЗ/п.6 ст.97 Договора о ЕАЭС: миграционный учёт
аннулируется при обращении о постановке на учёт по НОВОМУ месту пребывания, требуется
новая регистрация с тем же сроком (30 дней ЕАЭС / 90 Белоруссия), что при первичном
въезде. Правило добавлено в deadlines.py с trigger_field="address_since".

Раньше повторный вызов этой функции с той же датой создавал бы дубль Obligation.
Теперь: (1) если trigger_date не изменился — новая запись не создаётся (дедуп);
(2) если trigger_date новый (реальный новый въезд ИЛИ смена адреса) — старые
Obligation того же типа помечаются is_current=False, создаётся новая с is_current=True.
История сохраняется, но не задваивает активные списки. Все запросы, которые считают
"активные"/"просроченные" обязательства (дашборд webforms.py, /incomplete и
_handle_medical_exam_result в bot.py), должны фильтровать is_current=True — иначе
устаревшие версии снова начнут задваиваться.

2026-07 (второе изменение, слияние с ботом ТабельБелокаменка): добавлена
create_registration_obligation_for_return() — отдельная функция для возврата с
межвахты "за границу" (см. tabel.apply_rotation_return). НЕ переиспользует
create_obligations_for_employee целиком и НЕ трогает entry_date, потому что тот
триггерит ВСЕ правила категории по entry_date разом, включая MEDICAL_EXAM и
DACTYLOSCOPY — а это разовые обязанности, повторный въезд их НЕ должен создавать
заново (подтверждено явно в диалоге с заказчиком). Создаётся только REGISTRATION,
версионируется так же, как остальные (is_current), срок берётся из того же правила
DEADLINE_RULES, что и для entry_date (30 ЕАЭС / 90 Белоруссия) — по аналогии с
первичным въездом. Это ПРЕДПОЛОЖЕНИЕ, не проверенный юридический факт: возможно,
для ПОВТОРНОГО въезда после временного выезда действует другая норма, не такая же
льгота ЕАЭС, как при первом въезде. Проверить у юриста перед продакшеном (см. общую
пометку в deadlines.py — та же оговорка).
"""

import logging

from sqlalchemy.orm import Session

from models import (
    Category,
    DeadlineUnit,
    Employee,
    Obligation,
    ObligationStatus,
    ObligationType,
    RegistrationPeriod,
    RegistrationStatus,
)
from deadlines import DEADLINE_RULES, compute_deadline, calendar_days_add

log = logging.getLogger("obligations")


def _find_efs1_rule():
    """Ищет правило EFS1_REPORT в DEADLINE_RULES любой категории. ЕФС-1 — федеральная
    обязанность (СФР), срок один для всех (следующий рабочий день после договора/приказа),
    поэтому берём первое найденное правило, не привязываясь к конкретной категории.
    Возвращает (deadline_value, deadline_unit, trigger_field) или None, если правила нет."""
    for cat_rules in DEADLINE_RULES.values():
        for r in cat_rules:
            if r["type"] == ObligationType.EFS1_REPORT:
                return r["deadline_value"], r["deadline_unit"], r["trigger_field"]
    return None


def _create_rf_obligations(session: Session, employee: Employee) -> None:
    """Гражданин РФ (Category.RF): миграционных обязательств НЕТ (ни registration, ни
    medical_exam, ни dactyloscopy, ни contract_notice, ни RegistrationPeriod). Создаётся
    ТОЛЬКО ЕФС-1 в СФР — по дате договора. Договор как ДОКУМЕНТ создаётся отдельно в
    карточке; отдельной Obligation под него нет (в модели такого типа не существует).

    Согласие (152-ФЗ на обработку ПД) — тот же gate, что и для иностранцев: функция
    вызывается только после consent_status=confirmed, поэтому здесь повторно не проверяем.

    Срок ЕФС-1 берём из существующего правила EFS1_REPORT в DEADLINE_RULES (одинаков для
    всех категорий); fallback — 1 рабочий день, если правило почему-то не найдено. Так
    РФ-ветка самодостаточна и не требует отдельной записи для RF в deadlines.py."""
    if employee.contract_date is None:
        log.info(
            "РФ employee_id=%s: дата договора не задана — ЕФС-1 пока не создаётся "
            "(создастся при появлении contract_date)",
            employee.id,
        )
        return

    trigger_date = employee.contract_date

    found = _find_efs1_rule()
    if found is not None:
        deadline_value, deadline_unit, _ = found
    else:
        log.warning(
            "Правило EFS1_REPORT не найдено в DEADLINE_RULES — использую fallback "
            "1 рабочий день (employee_id=%s)",
            employee.id,
        )
        deadline_value, deadline_unit = 1, DeadlineUnit.WORKING_DAY

    already_exists = (
        session.query(Obligation)
        .filter_by(
            employee_id=employee.id,
            type=ObligationType.EFS1_REPORT,
            trigger_date=trigger_date,
        )
        .first()
    )
    if already_exists is not None:
        log.info(
            "ЕФС-1 для РФ employee_id=%s с trigger_date=%s уже существует — пропускаю",
            employee.id, trigger_date,
        )
        return

    deadline_date = compute_deadline(trigger_date, deadline_value, deadline_unit)

    superseded = (
        session.query(Obligation)
        .filter_by(employee_id=employee.id, type=ObligationType.EFS1_REPORT, is_current=True)
        .all()
    )
    for old in superseded:
        old.is_current = False
        session.add(old)

    obligation = Obligation(
        employee_id=employee.id,
        type=ObligationType.EFS1_REPORT,
        trigger_date=trigger_date,
        deadline_value=deadline_value,
        deadline_unit=deadline_unit,
        deadline_date=deadline_date,
        status=ObligationStatus.PENDING,
        is_current=True,
    )
    session.add(obligation)
    session.commit()


def create_obligations_for_employee(session: Session, employee: Employee) -> None:
    """Вызывается ТОЛЬКО после consent_status=confirmed. Без согласия obligations не создаются —
    это тот самый gate, который обсуждался как обязательное условие."""
    # Гражданин РФ — миграционного учёта нет вообще. Ветка стоит ДО чтения rules и ДО
    # проверки registration_status: у РФ нет правил в DEADLINE_RULES (вернулось бы [] →
    # ранний return) и не заполняется registration_status (тоже ранний return) — без этой
    # ветки obligations для РФ молча не создались бы ни одного. Создаём только ЕФС-1.
    if employee.is_rf:
        _create_rf_obligations(session, employee)
        return

    rules = DEADLINE_RULES.get(employee.category, [])
    if not rules:
        log.warning(
            "Нет правил дедлайнов для категории %s (employee_id=%s) — obligations не созданы",
            employee.category,
            employee.id,
        )
        return

    if employee.registration_status is None:
        log.warning(
            "Статус учёта не задан (employee_id=%s) — обязательства НЕ создаются до заполнения",
            employee.id,
        )
        return

    for rule in rules:
        if (
            employee.registration_status == RegistrationStatus.PRIOR
            and rule["trigger_field"] == "entry_date"
        ):
            continue
        trigger_date = getattr(employee, rule["trigger_field"])

        if (
            rule["trigger_field"] == "contract_end_date"
            and trigger_date is not None
            and trigger_date > date.today()
        ):
            log.info(
                "Увольнение employee_id=%s назначено на будущее (%s) — обязательство %s "
                "отложено до наступления даты",
                employee.id, trigger_date, rule["type"],
            )
            continue

        if trigger_date is None:
            log.warning(
                "Поле %s пустое у employee_id=%s — пропускаю obligation %s",
                rule["trigger_field"],
                employee.id,
                rule["type"],
            )
            continue

        # DEPARTURE_NOTICE (снятие с миграционного учёта при убытии) создаётся ТОЛЬКО если
        # работник реально стоял на учёте — есть действующая (is_current, не CANCELLED)
        # REGISTRATION. Нельзя снять с учёта того, кого на учёт не ставили: без этой проверки
        # система вешала уведомление об убытии на каждого уволенного, даже не подававшегося,
        # и задача висела впустую (разбирали вручную по 7 уволенным). Постановки не было —
        # уведомление об убытии не требуется (ст.23 №109-ФЗ применяется к стоявшим на учёте).
        if rule["type"] == ObligationType.DEPARTURE_NOTICE:
            has_active_registration = (
                session.query(Obligation)
                .filter(
                    Obligation.employee_id == employee.id,
                    Obligation.type == ObligationType.REGISTRATION,
                    Obligation.is_current.is_(True),
                    Obligation.status != ObligationStatus.CANCELLED,
                )
                .first()
                is not None
            )
            if not has_active_registration:
                log.info(
                    "employee_id=%s не стоял на учёте (нет действующей REGISTRATION) — "
                    "DEPARTURE_NOTICE не создаётся, снимать нечего",
                    employee.id,
                )
                continue

        already_exists = (
            session.query(Obligation)
            .filter_by(employee_id=employee.id, type=rule["type"], trigger_date=trigger_date)
            .first()
        )
        if already_exists is not None:
            log.info(
                "Obligation %s для employee_id=%s с trigger_date=%s уже существует — пропускаю",
                rule["type"],
                employee.id,
                trigger_date,
            )
            continue

        deadline_date = compute_deadline(trigger_date, rule["deadline_value"], rule["deadline_unit"])

        status = ObligationStatus.PENDING
        if rule["type"] == ObligationType.DACTYLOSCOPY and employee.dactyloscopy_date is not None:
            status = ObligationStatus.DONE

        superseded = (
            session.query(Obligation)
            .filter_by(employee_id=employee.id, type=rule["type"], is_current=True)
            .all()
        )
        for old in superseded:
            old.is_current = False
            session.add(old)

        obligation = Obligation(
            employee_id=employee.id,
            type=rule["type"],
            trigger_date=trigger_date,
            deadline_value=rule["deadline_value"],
            deadline_unit=rule["deadline_unit"],
            deadline_date=deadline_date,
            status=status,
            is_current=True,
        )
        session.add(obligation)

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


def create_registration_obligation_for_return(
    session: Session, employee: Employee, return_date: date
) -> None:
    """
    Возврат с межвахты "за границу" (пересёк границу РФ) — новая постановка на учёт,
    БЕЗ повторной дактилоскопии/медосмотра (разовые, не привязаны к повторному въезду —
    подтверждено явно в диалоге с заказчиком). НЕ трогает employee.entry_date и НЕ
    вызывает create_obligations_for_employee целиком — тот прошёлся бы по ВСЕМ правилам
    entry_date разом, включая дактилоскопию/медосмотр.

    Срок берётся из ТОГО ЖЕ правила REGISTRATION/entry_date в DEADLINE_RULES (30 ЕАЭС /
    90 Белоруссия), по аналогии с первичным въездом — см. предупреждение в докстринге
    модуля насчёт того, что это предположение, не проверенный факт.
    """
    rules = DEADLINE_RULES.get(employee.category, [])
    reg_rule = next(
        (r for r in rules if r["type"] == ObligationType.REGISTRATION
         and r["trigger_field"] == "entry_date"),
        None,
    )
    if reg_rule is None:
        log.warning(
            "Нет правила REGISTRATION/entry_date для категории %s (employee_id=%s) — "
            "обязательство по возврату из-за границы НЕ создано",
            employee.category, employee.id,
        )
        return

    already_exists = (
        session.query(Obligation)
        .filter_by(employee_id=employee.id, type=ObligationType.REGISTRATION, trigger_date=return_date)
        .first()
    )
    if already_exists is not None:
        log.info(
            "Obligation REGISTRATION для employee_id=%s с trigger_date=%s уже существует — пропускаю",
            employee.id, return_date,
        )
        return

    deadline_date = compute_deadline(return_date, reg_rule["deadline_value"], reg_rule["deadline_unit"])

    superseded = (
        session.query(Obligation)
        .filter_by(employee_id=employee.id, type=ObligationType.REGISTRATION, is_current=True)
        .all()
    )
    for old in superseded:
        old.is_current = False
        session.add(old)

    obligation = Obligation(
        employee_id=employee.id,
        type=ObligationType.REGISTRATION,
        trigger_date=return_date,
        deadline_value=reg_rule["deadline_value"],
        deadline_unit=reg_rule["deadline_unit"],
        deadline_date=deadline_date,
        status=ObligationStatus.PENDING,
        is_current=True,
    )
    session.add(obligation)
    session.commit()


# --- Правка bot.py (сделать один раз) ---------------------------------------
# 1. Удалить в bot.py определение функции create_obligations_for_employee целиком
#    (блок от "def create_obligations_for_employee(session: Session, employee: Employee)"
#    до строки "session.commit()" внутри неё).
# 2. Заменить в импортах bot.py:
#      from models import (
#          Base, Category, Consent, ConsentMethod, ConsentStatus, Employee,
#          NotificationSubscriber, Obligation, ObligationStatus, ObligationType,
#          RegistrationPeriod,
#      )
#      from deadlines import DEADLINE_RULES, compute_deadline, calendar_days_add
#    на:
#      from models import (
#          Base, Category, Consent, ConsentMethod, ConsentStatus, Employee,
#          NotificationSubscriber, Obligation, ObligationStatus, ObligationType,
#      )
#      from obligations import create_obligations_for_employee
#    (RegistrationPeriod и импорт deadlines.py в bot.py после переноса функции больше
#    нигде не используются — если это не так, оставь их импорт как есть).
