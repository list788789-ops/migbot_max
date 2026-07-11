"""
reports.py — данные отчётов, общие для веб-части (webforms.py, полный HTML) и
бота (bot.py, урезанный текст). Здесь только СБОР ДАННЫХ (dict/list простых
значений), никакого HTML и никакого форматирования под конкретный интерфейс —
это сознательное разделение: если понадобится третья поверхность вывода
(например, экспорт в файл), она тоже сможет использовать эти же функции
без изменений.

REPORTS_REGISTRY — список отчётов для витрины /reports в вебе И для
подменю "📊 Отчёты" в боте: (key, web_href, title, description).
Чтобы добавить новый отчёт — одна запись сюда + функция данных ниже +
рендер в webforms.py (HTML) и/или bot.py (текст, можно урезанный).
"""

from datetime import date, datetime

from sqlalchemy import select
from sqlalchemy.orm import Session

from models import AttendanceMark, Employee, Obligation, ObligationStatus
import tabel

REPORTS_REGISTRY = [
    ("changelog", "/reports/changelog", "🐛 Журнал ошибок и патчей",
     "Все найденные баги и внесённые исправления за время разработки табеля/миграционного учёта."),
    ("monthly_problems", "/reports/monthly-problems", "📊 Проблемные за месяц",
     "Неявки и выходные сверх нормы — кто накопил лишнее за текущий месяц."),
    ("obligations", "/reports/obligations", "📋 Обязательства — сводка по статусам",
     "Сколько просрочено/ожидает/закрыто по всем активным, с деталями по просроченным."),
    ("activity", "/reports/activity", "🕵️ Активность в табеле",
     "Кто (прораб/бот/скрипт переноса) сколько отметок поставил за месяц."),
]


# ================= 1. Журнал патчей (статический) =================

# Формат записи: (заголовок, что было не так, как исправлено).
CHANGELOG_TABEL = [
    ("Синтаксис при вставке через телефон",
     "Докстринг в sheets.py ломался при копировании текста на телефоне (смарт-кавычки/тире вместо прямых) — файл не запускался вообще.",
     "Заливать файл как файл (Upload files на GitHub), не вставлять текст в веб-редактор."),
    ("Нечёткое сопоставление ФИО при дублях",
     "Один и тот же человек мог попасть в базу дважды при разном написании отчества.",
     "find_fuzzy_matches — сравнение по Фамилия+Имя, отчество не мешает совпадению."),
    ("Квота Google Sheets API",
     "До 14 запросов на одного сотрудника при массовой загрузке — упирались в лимит 60/мин.",
     "Батчинг: 2 запроса на сотрудника вместо 14 (кэш метаданных + один batchUpdate)."),
    ("day_summary не считал МУ",
     "Отчёт «Табель за сегодня» вообще не показывал строку миграционного учёта — код МУ проваливался мимо всех веток подсчёта.",
     "Добавлена ветка МУ в day_summary + строка «📋 Мигр.учёт» в шаблон отчёта."),
    ("Дата межвахты принималась любая",
     "Можно было ввести сегодняшнюю или прошедшую дату как «дату возврата» — не имеет смысла.",
     "Валидация: дата возврата обязана быть в будущем, иначе просит ввести заново."),
    ("Путаница дата отъезда / дата возврата",
     "Прораб путал, какую именно дату вводит при постановке МЖ.",
     "Текст промпта явно: «дата ВОЗВРАТА, не отъезда»."),
    ("5 пикеров плодили новый список вместо обновления",
     "menu:docpick, menu:incomplete, menu:contractdate, menu:pending_consent, menu:delete_employee — все звали _deliver_picker без edit=True, «Назад к списку» создавал дубль сообщения.",
     "edit=True добавлен во все 5 мест — список редактируется на месте."),
    ("Отчёт «Список сотрудников» кэшировался",
     "Файл всегда сохранялся под одним и тем же именем employees.xlsx — риск показа старой версии.",
     "Имя файла с меткой времени (employees_ГГГГММДД_ЧЧММСС.xlsx) на каждый запрос."),
]

CHANGELOG_MIGBOT = [
    ("Категория гражданства: английский vs русский",
     "bot.py сравнивал citizenship.lower()==\"belarus\" (английское слово), а webforms.py — с русским «Беларусь». Через бота Белоруссию никогда бы не распознали → неверный срок постановки на учёт (30 дней вместо 90).",
     "Вынесено в common_utils.py, единая функция category_for_citizenship для обоих сервисов."),
    ("Нормализация телефона задублирована",
     "webforms.py и auth_binding.py имели свои копии _normalize_phone — риск рассинхрона при правке одного места.",
     "Вынесено в common_utils.py, оба места импортируют оттуда."),
    ("DetachedInstanceError после commit",
     "В 4+ местах bot.py обращение к employee.full_name происходило ПОСЛЕ session.commit() и закрытия сессии — SQLAlchemy инвалидирует атрибуты после коммита.",
     "Имя сотрудника забирается строкой ДО commit, используется вместо повторного обращения к ORM-объекту."),
    ("Обработчик reasoncode:force был недостижим",
     "Порядок проверок: общий if payload.startswith(\"reasoncode:\") перехватывал и force-вариант раньше, чем до него доходила очередь — кнопка «Всё равно МУ» не работала.",
     "Проверка :force вынесена и проверяется первой, до общего обработчика."),
    ("Ложные срабатывания «явка без договора»",
     "Проверка помечала ЛЮБУЮ явку у уволенного — включая дни ДО увольнения, что нормально (последний рабочий день).",
     "Помечается только явка ПОСЛЕ даты увольнения / ДО даты начала договора — реальное нарушение, не рутинный факт."),
    ("Уволенные видны в /employees веба",
     "Отдельный /archive для уволенных уже существовал, но /employees их не исключал — сотрудник виден сразу в обоих местах.",
     "Добавлен фильтр contract_end_date IS NULL в /employees, как и в отчёте бота."),
    ("«Осиротевшие» межвахты после переноса истории",
     "Разовый скрипт переноса истории из Google Sheets копировал код «МЖ», но не создавал запись в rotation_returns (та создаётся только через tabel.set_rotation) — 5 человек стояли на межвахте без даты возврата и не появлялись ни в напоминаниях, ни во флагах кадровику.",
     "Разовый скрипт-заглушка создал записи с expected_return_date=NULL; добавлен постоянный флоу уточнения (кадровику — задача в вебе, прорабу — напоминание с кнопкой в боте)."),
    ("Рассылка могла раскрыть кадровику-only данные прорабу",
     "NotificationSubscriber.chat_id не связан с ролью пользователя — проактивная рассылка ушла бы всем подписавшимся, включая прораба, которому детали обязательств видеть не должны.",
     "Для чувствительных данных (явка без договора) оставлен только пассивный пункт меню с проверкой роли; для нейтральных (межвахта, никогда не отмеченные — это работа самого прораба) рассылка оставлена как есть."),
]


# ================= 2. Проблемные за месяц =================

def get_monthly_problems_report(session: Session) -> dict:
    today = date.today()
    problems = tabel.get_monthly_problems(session)
    return {
        "month_label": today.strftime("%B %Y"),
        "absent_threshold": tabel.ABSENT_THRESHOLD,
        "weekend_threshold": tabel.WEEKEND_THRESHOLD,
        "problems": problems,  # [{"name","absent_count","weekend_count"}]
    }


# ================= 3. Обязательства — сводка =================

def get_obligations_report(session: Session) -> dict:
    active_ids = {e.id for e in tabel.get_active_employees(session)}
    current_obligations = session.scalars(
        select(Obligation).where(Obligation.is_current == True)  # noqa: E712
    ).all()
    current_obligations = [o for o in current_obligations if o.employee_id in active_ids]

    counts = {}  # (type_value, status_value) -> count
    overdue_details = []  # [{"name","type_value","deadline_date","employee_id"}]
    for o in current_obligations:
        key = (o.type.value, o.status.value)
        counts[key] = counts.get(key, 0) + 1
        if o.status == ObligationStatus.OVERDUE:
            emp = session.get(Employee, o.employee_id)
            overdue_details.append({
                "name": emp.full_name if emp else "?",
                "type_value": o.type.value,
                "deadline_date": o.deadline_date,
                "employee_id": o.employee_id,
            })
    overdue_details.sort(key=lambda d: d["deadline_date"])
    return {"counts": counts, "overdue_details": overdue_details}


# ================= 4. Активность в табеле =================

def resolve_actor_label(session: Session, actor_id: str) -> str:
    """Три формата created_by (см. журнал патчей): 'migration_script', числовой
    MAX user_id (бот), готовое ФИО строкой (веб, см. _actor_name в webforms.py)."""
    from auth_binding import find_user_by_max_id, get_role_label  # локальный импорт — избегаем цикла

    if actor_id == "migration_script":
        return "🔄 Перенос истории (разовый скрипт)"
    if actor_id.isdigit():
        user = find_user_by_max_id(session, actor_id)
        if user:
            return f"👤 {user.full_name} ({get_role_label(user)}, из бота)"
        return f"🤖 MAX-аккаунт {actor_id} (не привязан к User — /login не выполнен)"
    return f"👤 {actor_id} (из веба)"


def get_activity_report(session: Session) -> dict:
    today = date.today()
    month_start = date(today.year, today.month, 1)
    marks = session.scalars(
        select(AttendanceMark).where(AttendanceMark.mark_date >= month_start)
    ).all()

    by_actor = {}  # created_by -> {"count","first","last"}
    for m in marks:
        a = by_actor.setdefault(m.created_by, {"count": 0, "first": m.mark_date, "last": m.mark_date})
        a["count"] += 1
        a["first"] = min(a["first"], m.mark_date)
        a["last"] = max(a["last"], m.mark_date)

    actors = [
        {"actor_id": actor, "label": resolve_actor_label(session, actor), **stats}
        for actor, stats in by_actor.items()
    ]
    actors.sort(key=lambda a: -a["count"])
    return {"month_start": month_start, "total": len(marks), "actors": actors}
