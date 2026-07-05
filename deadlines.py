"""
Конфиг сроков по категориям — ТОЛЬКО для eaeu, остальное зарезервировано на будущее.

ВАЖНО: это временное решение для скелета. Как обсуждалось — сроки меняются чаще кода
(изменения в 2026 году шли почти ежемесячно), поэтому перед продакшеном это должно стать
таблицей в БД с полем valid_from/valid_to, а не питоновским словарём в коде.
Иначе обновление срока = деплой, а не запись в БД.

2026-07: добавлено правило EFS1_REPORT (ЕФС-1 в СФР). НЕ добавлено правило продления
регистрации по принципу "90 дней из скользящих 180" — оно принципиально не укладывается
в эту структуру. Все правила здесь — разовое смещение от одной даты, вычисляемое один раз
при создании obligations (см. create_obligations_for_employee в obligations.py). Правило
90/180 требует периодического пересчёта и знания истории продлений, а не разового дедлайна.
Добавлять его сюда как обычную запись — значит один раз посчитать неверный дедлайн и
никогда не обновить его при продлении. Нужен отдельный механизм (recurring obligation
либо periodic job), спроектировать отдельно.

2026-07: добавлено ВТОРОЕ правило REGISTRATION с trigger_field="address_since". Смена
места пребывания в России — отдельное юридическое событие. ИСПРАВЛЕНО (2026-07): срок
переезда НЕ равен первичному въезду. Льгота ЕАЭС (30 суток, п.6 ст.97 Договора) действует
ТОЛЬКО на первичный въезд; переезд внутри РФ — 7 рабочих дней независимо от гражданства
(п.3.1 ст.20 №109-ФЗ), а при заселении на вахту/в общежитие — 1 рабочий день с прибытия.
Площадка Белокаменка = вахта, поэтому у EAEU здесь стоит 1 рабочий день. Оба правила одного
типа (entry_date И address_since) намеренно живут в одном списке — create_obligations_for_employee
в obligations.py дедуплицирует и версионирует по (employee_id, type, trigger_date), а не по
правилу, так что наличие двух правил одного типа не создаёт конфликтов. MEDICAL_EXAM и
DACTYLOSCOPY НЕ дублируются на address_since — они привязаны к факту въезда, не к месту
пребывания, смена адреса их не ретриггерит.

ИСПРАВЛЕНО (2026-07): у Белоруссии address_since тоже была ошибка — стояло 90 календарных
по аналогии с въездом. Переезд по РФ — 7 рабочих дней независимо от гражданства (то же
109-ФЗ). Заменено на 7 рабочих. Вахтовое правило 1 дня на Белоруссию НЕ распространяю —
это специфика площадки на ЕАЭС-казахах, отдельного основания для белорусов нет. Категория
не используется, но раз правило есть — оно должно быть верным, а не миной на будущее.

Источники на момент составления (проверять перед продакшеном у юриста):
- ЕАЭС: 30 суток (календарных) с даты въезда — п.6 ст.97 Договора о ЕАЭС от 29.05.2014
- Уведомление о договоре: 3 рабочих дня — ст.13 115-ФЗ, форма МВД №536
- Медосвидетельствование: справка нужна, если с даты въезда прошло больше 30 календарных дней
- ЕФС-1: не позднее следующего рабочего дня после приказа о приёме/даты договора —
  пп.2 п.5 ст.11 ФЗ №27-ФЗ "О персонифицированном учёте"
- Смена места пребывания = новая постановка на учёт. Срок НЕ равен въезду: 7 рабочих дней
  (п.3.1 ст.20 №109-ФЗ), для вахты/общежития — 1 рабочий день. Льгота ЕАЭС только на въезд.
- Дактилоскопия + фотографирование: разовая, 30 календарных дней с даты въезда —
  п.13 ст.5 №115-ФЗ (Информация МВД от 20.05.2025). Карта на 10 лет, ежегодного повтора нет.
"""

from models import Category, DeadlineUnit, ObligationType

# category -> list of (obligation_type, trigger_field, deadline_value, deadline_unit, lead_days)
# trigger_field — какое поле employee считать точкой отсчёта
# lead_days — за сколько дней до дедлайна чип желтеет и уходит проактивное уведомление.
#   ОДНО число и для чипа, и для крона (не разъезжаются). Для правил в РАБОЧИХ днях
#   (переезд/уведомление/ЕФС-1) обязанность born-amber: жёлтая с создания независимо от
#   lead_days — их окно короче любого порога, а раскладка 'дней до' над выходными ненадёжна.
#   Числовой lead_days содержателен только для 30-дневных календарных (7/14/14).
#   Порог берётся из этих же правил через lead_days_for() — один источник со сроком.
DEADLINE_RULES: dict[Category, list[dict]] = {
    Category.EAEU: [
        {
            "type": ObligationType.REGISTRATION,
            "trigger_field": "entry_date",
            "deadline_value": 30,
            "deadline_unit": DeadlineUnit.CALENDAR_DAY,
            "lead_days": 7,
        },
        {
            # Смена места пребывания. НЕ льгота ЕАЭС (она только на первичный въезд).
            # Вахта (Белокаменка) = 1 рабочий день с прибытия. Born-amber: жёлтая с создания.
            "type": ObligationType.REGISTRATION,
            "trigger_field": "address_since",
            "deadline_value": 1,
            "deadline_unit": DeadlineUnit.WORKING_DAY,
            "lead_days": 1,
        },
        {
            "type": ObligationType.CONTRACT_NOTICE,
            "trigger_field": "contract_date",
            "deadline_value": 3,
            "deadline_unit": DeadlineUnit.WORKING_DAY,
            "lead_days": 3,
        },
        {
            "type": ObligationType.MEDICAL_EXAM,
            "trigger_field": "entry_date",
            "deadline_value": 30,
            "deadline_unit": DeadlineUnit.CALENDAR_DAY,
            "lead_days": 14,
        },
        {
            "type": ObligationType.EFS1_REPORT,
            "trigger_field": "contract_date",
            "deadline_value": 1,
            "deadline_unit": DeadlineUnit.WORKING_DAY,
            "lead_days": 1,
        },
        {
            # Дактилоскопия + фотографирование ("грин карта"). Разовая, 30 календарных дней
            # с даты въезда (п.13 ст.5 №115-ФЗ) — тот же триггер и срок, что медосмотр.
            # Закрывается внесением employee.dactyloscopy_date (webforms переводит в DONE).
            # НЕ годичная: карта на 10 лет, ежегодного пересчёта быть не должно.
            "type": ObligationType.DACTYLOSCOPY,
            "trigger_field": "entry_date",
            "deadline_value": 30,
            "deadline_unit": DeadlineUnit.CALENDAR_DAY,
            "lead_days": 14,
        },
    ],
    Category.BELARUS: [
        {
            "type": ObligationType.REGISTRATION,
            "trigger_field": "entry_date",
            "deadline_value": 90,
            "deadline_unit": DeadlineUnit.CALENDAR_DAY,
            "lead_days": 7,
        },
        {
            # ИСПРАВЛЕНО с 90 календарных: переезд по РФ — 7 рабочих дней независимо от
            # гражданства (п.3.1 ст.20 №109-ФЗ). Вахтовое правило 1 дня сюда не переносим.
            "type": ObligationType.REGISTRATION,
            "trigger_field": "address_since",
            "deadline_value": 7,
            "deadline_unit": DeadlineUnit.WORKING_DAY,
            "lead_days": 7,
        },
        {
            "type": ObligationType.CONTRACT_NOTICE,
            "trigger_field": "contract_date",
            "deadline_value": 3,
            "deadline_unit": DeadlineUnit.WORKING_DAY,
            "lead_days": 3,
        },
        {
            "type": ObligationType.EFS1_REPORT,
            "trigger_field": "contract_date",
            "deadline_value": 1,
            "deadline_unit": DeadlineUnit.WORKING_DAY,
            "lead_days": 1,
        },
    ],
    # PATENT, VISA, HQS — намеренно не заполнены. При добавлении первого сотрудника
    # этих категорий рулы нужно проверить у юриста, а не копировать по аналогии с EAEU/BELARUS.
}


def lead_days_for(category, obligation_type, deadline_unit, default: int = 7) -> int:
    """Порог 'скоро' (жёлтый чип и момент уведомления) для обязанности данного типа.
    Один источник истины со сроком — берётся из того же правила DEADLINE_RULES.

    Ключ поиска — (type, deadline_unit): в пределах категории он уникален даже для
    REGISTRATION, у которой два правила (entry_date=CALENDAR_DAY, address_since=WORKING_DAY).

    Короткоплечие правила (WORKING_DAY) считаются born-amber на стороне webforms
    (жёлтая с создания), поэтому их числовой lead_days номинален. Для 30-дневных
    календарных (регистрация 7, медосмотр/дактилоскопия 14) lead_days содержателен.

    default возвращается, если правило не найдено — например, REGISTRATION_RENEWAL
    создаётся отдельным механизмом и в DEADLINE_RULES отсутствует."""
    for rule in DEADLINE_RULES.get(category, []):
        if rule["type"] == obligation_type and rule["deadline_unit"] == deadline_unit:
            return rule.get("lead_days", default)
    return default


def working_days_add(start, days: int):
    """Наивная реализация — считает только будни, БЕЗ учёта праздников РФ.
    Для продакшена подключить производственный календарь (напр. библиотеку `workalendar`),
    иначе дедлайн 'contract_notice' будет ошибочно попадать на праздники.

    Особенно критично для EFS1_REPORT — там всего 1 рабочий день запаса. Один
    непросчитанный праздник (например, после новогодних каникул) сдвинет реальный
    дедлайн раньше, чем покажет этот расчёт, и просрочка возникнет незаметно."""
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
