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
"""

import logging

from sqlalchemy.orm import Session

from models import Category, Employee, Obligation, ObligationStatus, RegistrationPeriod
from deadlines import DEADLINE_RULES, compute_deadline, calendar_days_add

log = logging.getLogger("obligations")


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

        # Идемпотентность по (employee_id, type, trigger_date): функция может быть вызвана
        # повторно для одного и того же сотрудника (согласие подтвердили, потом поправили
        # дату договора — обе ветки в bot.py и webforms.py вызывают эту функцию заново).
        # Если дата-триггер НЕ изменилась — это тот же самый въезд/событие, повторной
        # записи не нужно. Если дата другая (сотрудник выехал и въехал заново, дату
        # исправили на реально другую) — это новое событие, и обязательство должно быть новым.
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

        # Версионирование: новое обязательство того же типа означает, что предыдущее
        # (например, регистрация по старому адресу) больше не актуально — сменился адрес,
        # исправили дату, и т.п. Старые записи НЕ удаляются и не перезаписываются (нужна
        # история), а помечаются is_current=False. Дашборд и разделы бота показывают
        # только is_current=True — см. правки в bot.py/webforms.py, отмеченные ниже.
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
            status=ObligationStatus.PENDING,
            is_current=True,
        )
        session.add(obligation)

    # Первый период учёта "90 из 180" — только для EAEU (см. обоснование в исходном bot.py:
    # для BELARUS механизм иной, перенос по аналогии не делался без отдельной юр. проверки).
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
