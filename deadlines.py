"""
Конфиг сроков по категориям — ТОЛЬКО для eaeu, остальное зарезервировано на будущее.

ВАЖНО: это временное решение для скелета. Как обсуждалось — сроки меняются чаще кода
(изменения в 2026 году шли почти ежемесячно), поэтому перед продакшеном это должно стать
таблицей в БД с полем valid_from/valid_to, а не питоновским словарём в коде.
Иначе обновление срока = деплой, а не запись в БД.

Источники на момент составления (проверять перед продакшеном у юриста):
- ЕАЭС: 30 суток (календарных) с даты въезда — п.6 ст.97 Договора о ЕАЭС от 29.05.2014
- Уведомление о договоре: 3 рабочих дня — ст.13 115-ФЗ, форма МВД №536
- Медосвидетельствование: справка нужна, если с даты въезда прошло больше 30 календарных дней
"""

from models import Category, DeadlineUnit, ObligationType

# category -> list of (obligation_type, trigger_field, deadline_value, deadline_unit)
# trigger_field — какое поле employee считать точкой отсчёта
DEADLINE_RULES: dict[Category, list[dict]] = {
    Category.EAEU: [
        {
            "type": ObligationType.REGISTRATION,
            "trigger_field": "entry_date",
            "deadline_value": 30,
            "deadline_unit": DeadlineUnit.CALENDAR_DAY,
        },
        {
            "type": ObligationType.CONTRACT_NOTICE,
            "trigger_field": "contract_date",
            "deadline_value": 3,
            "deadline_unit": DeadlineUnit.WORKING_DAY,
        },
        {
            "type": ObligationType.MEDICAL_EXAM,
            "trigger_field": "entry_date",
            "deadline_value": 30,
            "deadline_unit": DeadlineUnit.CALENDAR_DAY,
        },
    ],
    Category.BELARUS: [
        {
            "type": ObligationType.REGISTRATION,
            "trigger_field": "entry_date",
            "deadline_value": 90,
            "deadline_unit": DeadlineUnit.CALENDAR_DAY,
        },
        {
            "type": ObligationType.CONTRACT_NOTICE,
            "trigger_field": "contract_date",
            "deadline_value": 3,
            "deadline_unit": DeadlineUnit.WORKING_DAY,
        },
    ],
    # PATENT, VISA, HQS — намеренно не заполнены. При добавлении первого сотрудника
    # этих категорий рулы нужно проверить у юриста, а не копировать по аналогии с EAEU/BELARUS.
}


def working_days_add(start, days: int):
    """Наивная реализация — считает только будни, БЕЗ учёта праздников РФ.
    Для продакшена подключить производственный календарь (напр. библиотеку `workalendar`),
    иначе дедлайн 'contract_notice' будет ошибочно попадать на праздники."""
    from datetime import timedelta

    current = start
    added = 0
    while added < days:
        current += timedelta(days=1)
        if current.weekday() < 5:  # 0-4 = пн-пт
            added += 1
    return current


def calendar_days_add(start, days: int):
    from datetime import timedelta

    return start + timedelta(days=days)


def compute_deadline(trigger_date, value: int, unit: DeadlineUnit):
    if unit == DeadlineUnit.WORKING_DAY:
        return working_days_add(trigger_date, value)
    return calendar_days_add(trigger_date, value)
