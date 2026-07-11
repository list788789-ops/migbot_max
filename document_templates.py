"""
Генерация docx-документов для отправки сотруднику через бота:
  - Согласие на обработку персональных данных (152-ФЗ)
  - Направление на медицинский осмотр (медкомиссию, ст.13.3 115-ФЗ)

ВАЖНО: для согласия по 152-ФЗ обязательна идентификация оператора (работодателя) —
без неё документ не соответствует требованию "конкретности" (ч.1 ст.9 152-ФЗ).
Данные оператора берутся из переменных окружения — см. COMPANY_* ниже.
ПЕРЕД ПЕРВЫМ ИСПОЛЬЗОВАНИЕМ обязательно заполнить реальными данными компании:
плейсхолдеры в квадратных скобках не являются юридически значимыми и их наличие
в отправленном документе означает, что согласие не имеет силы.

С 1 сентября 2025 согласие на обработку ПД должно оформляться ОТДЕЛЬНЫМ документом,
не как часть трудового договора или другой формы — поэтому это отдельный docx,
а не пункт в существующем consent_texts.py (тот — для текста в чате, не для подписи).

2026-07: добавлена проверка обязательных полей сотрудника ПЕРЕД генерацией — раньше
пустые passport_series/passport_number/birth_date/address молча превращались в текст
"[не указано]" внутри готового документа, который уходит в клинику или сотруднику на
подпись. Теперь генератор поднимает ValueError с точным списком недостающих полей —
ВАЖНО: вызывающий код (bot.py, webforms.py) должен показывать текст ЭТОГО исключения
кадровику, а не общее "не удалось сгенерировать документ", иначе смысл проверки теряется.

2026-07: добавлен блок подписей "От Исполнителя / От Заказчика" в направление на
медосмотр — был в бумажном шаблоне (Приложение №1 к договору №176), но отсутствовал
в генераторе; документ обрывался на пункте 10.

2026-07: добавлен ТЕСТОВЫЙ режим TEST_ALLOW_MISSING_FIELDS — если включён, генератор
НЕ поднимает ValueError при незаполненных полях, а подставляет прочерк "—" и добавляет
явное предупреждение внутрь самого документа (чтобы черновик нельзя было спутать
с юридически валидным документом, если он случайно попадёт клинике или сотруднику).
Это временное послабление ТОЛЬКО для тестирования потока — флаг должен быть выключен
(или переменная удалена) до реальной работы с сотрудниками. Один флаг на оба генератора,
чтобы не потерять место, где включали, при отключении.

Требуемые переменные окружения:
  COMPANY_NAME             — полное наименование юрлица-работодателя
  COMPANY_INN              — ИНН
  COMPANY_LEGAL_ADDRESS    — юридический адрес
  HR_SIGNATORY_NAME        — ФИО подписанта со стороны работодателя
  HR_SIGNATORY_POSITION    — должность подписанта
  CLINIC_NAME              — ПОЛНОЕ наименование медицинской организации (шапка бланка)
  CLINIC_SHORT_NAME        — короткое имя МО для блока подписи (напр. ГОАУЗ «МОМЦ»)
  CLINIC_CONTRACT_NUMBER   — номер договора с клиникой
  CLINIC_CONTRACT_DATE     — дата договора с клиникой, формат ДД.ММ.ГГГГ
  CLINIC_CHIEF_DOCTOR_NAME — ФИО главного врача клиники (для блока подписи "От Исполнителя")
  PAYER_NAME               — заказчик услуги (напр. "ИП Буц С.Ю.") — используется в п.5/8 бланка
  PAYER_SIGNATORY_NAME     — ФИО подписанта со стороны заказчика для блока подписи
                             (напр. "С. Ю. Буц") — если не задано, используется PAYER_NAME
  PAYER_PHONE              — телефон заказчика
  TEST_ALLOW_MISSING_FIELDS — "true"/"1" чтобы разрешить генерацию с прочерками вместо
                             незаполненных полей (ТОЛЬКО для теста, см. выше)
"""

import os
from datetime import date, datetime, timedelta, timezone

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Font

from models import Employee

MSK = timezone(timedelta(hours=3))  # Мурманская обл. — московское время, без перехода на летнее с 2014

COMPANY_NAME = os.environ.get("COMPANY_NAME", "[НЕ ЗАПОЛНЕНО — укажите наименование юрлица]")
COMPANY_INN = os.environ.get("COMPANY_INN", "[ИНН не указан]")
COMPANY_ADDRESS = os.environ.get("COMPANY_LEGAL_ADDRESS", "[юридический адрес не указан]")
HR_SIGNATORY_NAME = os.environ.get("HR_SIGNATORY_NAME", "[ФИО подписанта не указано]")
HR_SIGNATORY_POSITION = os.environ.get("HR_SIGNATORY_POSITION", "[должность не указана]")

# ТЕСТОВЫЙ флаг — см. заголовок файла. Читается один раз при импорте модуля;
# если меняешь переменную окружения на Railway, нужен рестарт сервиса, чтобы применилось.
TEST_ALLOW_MISSING_FIELDS = os.environ.get("TEST_ALLOW_MISSING_FIELDS", "false").strip().lower() in (
    "1", "true", "yes",
)

DASH = "—"

# Поля, обязательные для обоих документов — вынесены в константы, чтобы webforms.py/bot.py
# могли получить список отсутствующих полей ДО генерации (для баннера/сообщения), не только
# через перехват ValueError.
CONSENT_REQUIRED_FIELDS = {
    "birth_date": "дата рождения",
    "passport_series": "серия паспорта",
    "passport_number": "номер паспорта",
    "address": "адрес места пребывания",
}
MEDICAL_REFERRAL_REQUIRED_FIELDS = {
    "birth_date": "дата рождения",
    "passport_series": "серия паспорта",
    "passport_number": "номер паспорта",
    # address убран (2026-07): п.3 направления берёт SITE_ADDRESS (константа площадки),
    # employee.address в направление больше не идёт, поэтому его пустота не должна
    # блокировать генерацию. В согласии employee.address по-прежнему используется.
}


def _set_default_style(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)


def _passport_str(employee: Employee) -> str:
    passport = f"{employee.passport_series or ''} {employee.passport_number or ''}".strip()
    return passport or DASH


def _date_or_dash(d: date | None) -> str:
    return d.strftime("%d.%m.%Y") if d else DASH


def _text_or_dash(value: str | None) -> str:
    return value if value else DASH


def _missing_fields(employee: Employee, field_labels: dict[str, str]) -> list[str]:
    """Возвращает список человекочитаемых имён незаполненных полей, ничего не бросая.
    Использовать это, когда нужен просто список (баннер в UI, сообщение бота) —
    без побочного эффекта в виде исключения."""
    return [
        label for attr, label in field_labels.items()
        if getattr(employee, attr) in (None, "")
    ]


def check_consent_fields(employee: Employee) -> list[str]:
    """Список отсутствующих полей для согласия — вызывать из webforms.py/bot.py
    до генерации, если нужно показать баннер независимо от режима TEST_ALLOW_MISSING_FIELDS."""
    return _missing_fields(employee, CONSENT_REQUIRED_FIELDS)


def check_medical_referral_fields(employee: Employee) -> list[str]:
    """То же самое для направления на медосмотр."""
    return _missing_fields(employee, MEDICAL_REFERRAL_REQUIRED_FIELDS)


def _require_fields(employee: Employee, field_labels: dict[str, str]) -> list[str]:
    """Обычный режим: поднимает ValueError с точным списком недостающих полей, вместо
    того чтобы молча вставить в документ текст-плейсхолдер.

    Тестовый режим (TEST_ALLOW_MISSING_FIELDS=true): НЕ поднимает исключение, а
    возвращает список отсутствующих полей — вызывающий генератор обязан сам подставить
    прочерки в текст документа и добавить предупреждение (см. generate_*_docx ниже).
    """
    missing = _missing_fields(employee, field_labels)
    if missing and not TEST_ALLOW_MISSING_FIELDS:
        raise ValueError(
            f"Нельзя сгенерировать документ для {employee.full_name} — "
            f"не заполнены поля: {', '.join(missing)}. "
            f"Заполните их в карточке сотрудника перед генерацией."
        )
    return missing


def _add_test_warning_paragraph(doc: Document, missing: list[str]) -> None:
    """Явное предупреждение прямо в теле документа, если он сгенерирован в тестовом
    режиме с прочерками. Цель — чтобы черновик нельзя было перепутать с юридически
    валидным документом, если он случайно уйдёт клинике или сотруднику на подпись."""
    warning = doc.add_paragraph()
    run = warning.add_run(
        "⚠ ТЕСТОВЫЙ ЧЕРНОВИК — не заполнены поля: "
        + ", ".join(missing)
        + ". Документ не имеет юридической силы, пока эти поля не указаны в карточке "
        "сотрудника и документ не перегенерирован."
    )
    run.bold = True
    doc.add_paragraph()


def generate_consent_docx(employee: Employee, output_dir: str = "/tmp") -> str:
    """Согласие на обработку персональных данных — отдельный документ (152-ФЗ)."""
    missing = _require_fields(employee, CONSENT_REQUIRED_FIELDS)

    doc = Document()
    _set_default_style(doc)

    if missing:
        _add_test_warning_paragraph(doc, missing)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("СОГЛАСИЕ\nна обработку персональных данных")
    run.bold = True
    run.font.size = Pt(14)

    doc.add_paragraph()

    birth = _date_or_dash(employee.birth_date)
    address = _text_or_dash(employee.address)

    body = (
        f"Я, {employee.full_name}, {birth} года рождения, "
        f"паспорт {_passport_str(employee)}, зарегистрированный(-ая) по адресу: {address}, "
        f"в соответствии со ст. 9 Федерального закона от 27.07.2006 № 152-ФЗ "
        f"«О персональных данных» даю согласие {COMPANY_NAME} (ИНН {COMPANY_INN}, "
        f"адрес: {COMPANY_ADDRESS}) (далее — Оператор) на обработку моих "
        f"персональных данных в целях исполнения трудового договора, ведения "
        f"миграционного учёта и исполнения обязанностей работодателя, "
        f"предусмотренных законодательством Российской Федерации о правовом "
        f"положении иностранных граждан."
    )
    doc.add_paragraph(body)

    doc.add_paragraph("Согласие даётся на обработку следующих персональных данных:")
    items = [
        "фамилия, имя, отчество; дата и место рождения; гражданство;",
        "паспортные данные, миграционная карта, данные о постановке на миграционный учёт;",
        "адрес регистрации и фактического проживания; контактный телефон;",
        "сведения о трудовом договоре, занимаемой должности;",
        "сведения о результатах медицинского осмотра, необходимые для допуска к работе;",
        "иные данные, прямо предусмотренные законодательством о миграционном учёте иностранных граждан.",
    ]
    for item in items:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_paragraph(
        "Согласие действует с момента подписания до его отзыва в письменной форме. "
        "Я проинформирован(-а) о праве отозвать настоящее согласие в любой момент, "
        "направив письменное заявление Оператору (ч. 2 ст. 9 152-ФЗ)."
    )

    doc.add_paragraph()
    doc.add_paragraph(f"Дата: {datetime.now(MSK).date().strftime('%d.%m.%Y')}")
    doc.add_paragraph(f"Подпись: _____________________ / {employee.full_name}")

    filename = f"consent_{employee.id}.docx"
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path


# Реквизиты клиники и заказчика. Значения по умолчанию зашиты как fallback (по образцу
# направления «Пирогова», договор №176), но переменная окружения их переопределяет —
# при смене договора/главврача/телефона правь env в Railway, а не этот файл.
# CLINIC_NAME — ПОЛНОЕ имя МО (идёт в шапку "В ..."); CLINIC_SHORT_NAME — короткое,
# для блока подписи "От Исполнителя" (в образце там "ГОАУЗ «МОМЦ»", а не полное имя).
CLINIC_NAME = os.environ.get(
    "CLINIC_NAME",
    "Государственное областное автономное учреждение здравоохранения "
    "«Мурманский областной медицинский центр»",
)
CLINIC_SHORT_NAME = os.environ.get("CLINIC_SHORT_NAME", "ГОАУЗ «МОМЦ»")
CLINIC_CONTRACT_NUMBER = os.environ.get("CLINIC_CONTRACT_NUMBER", "176")
CLINIC_CONTRACT_DATE = os.environ.get("CLINIC_CONTRACT_DATE", "25.06.2026")
CLINIC_CHIEF_DOCTOR_NAME = os.environ.get("CLINIC_CHIEF_DOCTOR_NAME", "А.М. Амозов")
PAYER_NAME = os.environ.get("PAYER_NAME", "ИП Буц С.Ю.")
PAYER_SIGNATORY_NAME = os.environ.get("PAYER_SIGNATORY_NAME", "С. Ю. Буц")
PAYER_PHONE = os.environ.get("PAYER_PHONE", "+7 (985) 415-54-20")

# Адрес площадки — идёт в п.3 направления вместо employee.address (у всех вахтовиков
# один адрес проживания). Только для НАПРАВЛЕНИЯ; в согласии (152-ФЗ) остаётся личный
# employee.address. Переопределяется env, если площадка сменится.
SITE_ADDRESS = os.environ.get(
    "SITE_ADDRESS", "Мурманская обл., Кольский р-н, с. Белокаменка, зд. 1А, 184664"
)

# --- Реквизиты ООО «ТРЕСТСТРОЙМОНТАЖ» (работодатель по трудовому договору) ---
# ВАЖНО: это ДРУГОЕ юрлицо, не ИП Буц (тот — заказчик медуслуги). Не путать.
# Значения по образцу договора №0074 (Уристемов), приняты как факт без проверки.
# Все переопределяемы через env при смене реквизитов.
EMPLOYER_NAME_FULL = os.environ.get(
    "EMPLOYER_NAME_FULL", 'Общество с ограниченной ответственностью "ТРЕСТСТРОЙМОНТАЖ"'
)
EMPLOYER_NAME_SHORT = os.environ.get("EMPLOYER_NAME_SHORT", "ООО «ТРЕСТСТРОЙМОНТАЖ»")
EMPLOYER_INN = os.environ.get("EMPLOYER_INN", "5038107922")
EMPLOYER_KPP = os.environ.get("EMPLOYER_KPP", "770501001")
EMPLOYER_LEGAL_ADDRESS = os.environ.get(
    "EMPLOYER_LEGAL_ADDRESS",
    "115114, Город Москва, вн.тер.г. муниципальный округ Замоскворечье, "
    "наб Шлюзовая, д. 8, стр. 1, помещ. 3ВН",
)
EMPLOYER_ACTUAL_ADDRESS = os.environ.get("EMPLOYER_ACTUAL_ADDRESS", EMPLOYER_LEGAL_ADDRESS)
EMPLOYER_PHONE = os.environ.get("EMPLOYER_PHONE", "+7 (495) 147-82-79")
EMPLOYER_DIRECTOR_FULL = os.environ.get("EMPLOYER_DIRECTOR_FULL", "Железняка Валерия Александровича")
EMPLOYER_DIRECTOR_SHORT = os.environ.get("EMPLOYER_DIRECTOR_SHORT", "Железняк В. А.")
# ОКВЭД и ОГРН работодателя — обязательные поля формы уведомления МВД №8 (приказ №536),
# но в системе их не было. Оставлены пустыми (env-override): по решению — заполняются вручную
# в готовом документе. Когда появятся — задать через переменные окружения EMPLOYER_OKVED/OGRN.
EMPLOYER_OKVED = os.environ.get("EMPLOYER_OKVED", "")
EMPLOYER_OGRN = os.environ.get("EMPLOYER_OGRN", "")
EMPLOYER_SUBDIVISION = os.environ.get("EMPLOYER_SUBDIVISION", "Обособленное подразделение Мурманск")
WORKPLACE_ADDRESS = os.environ.get(
    "WORKPLACE_ADDRESS",
    "Мурманская обл., г. Мурманск, ул. Сполохи 4А, ООО НОВАТЭК-Усть-Луга",
)
DISTRICT_COEFFICIENT = os.environ.get("DISTRICT_COEFFICIENT", "1,500")
CONTRACT_NUMBER_PREFIX = os.environ.get("CONTRACT_NUMBER_PREFIX", "БК-ПСМ-")


# --- Реквизиты госпошлины (квитанция ПД-4сб, налог). УФК по Мурманской области (УМВД).
# Фиксированные, из платёжных реквизитов УМВД (фото от пользователя, приняты как факт).
# Обе пошлины — один получатель, различаются назначением/КБК/суммой (см. DUTY_KINDS ниже).
DUTY_PAYEE_NAME = os.environ.get(
    "DUTY_PAYEE_NAME",
    "УФК по Мурманской области (УМВД России по Мурманской области, л/сч 04491137920)",
)
DUTY_PAYEE_INN = os.environ.get("DUTY_PAYEE_INN", "5191501766")
DUTY_PAYEE_KPP = os.environ.get("DUTY_PAYEE_KPP", "519001001")
DUTY_OKTMO = os.environ.get("DUTY_OKTMO", "47536000")
DUTY_ACCOUNT = os.environ.get("DUTY_ACCOUNT", "03100643000000014900")
DUTY_BANK_NAME = os.environ.get(
    "DUTY_BANK_NAME",
    "ОКЦ № 3 Северо-Западного ГУ Банка России//УФК по Мурманской области г. Мурманск",
)
DUTY_BIC = os.environ.get("DUTY_BIC", "014705901")
# Единый казначейский счёт (ЕКС / корр. счёт для QR по ГОСТ Р 56042). Определяется по БИК,
# подтверждён по официальным источникам (УФК/госорганы Мурманской обл.). Для QR обязателен.
DUTY_CORRESP_ACC = os.environ.get("DUTY_CORRESP_ACC", "40102810745370000041")
# Наименование получателя для QR (без л/сч — в QR идёт чистое наименование).
DUTY_PAYEE_NAME_QR = os.environ.get(
    "DUTY_PAYEE_NAME_QR", "УФК по Мурманской области (УМВД России по Мурманской области)"
)

# Типы пошлин: назначение, КБК (20 цифр, без пробелов), сумма. Ключ передаётся в генератор.
DUTY_KINDS = {
    "registration": {
        "purpose": (
            "Государственная пошлина на постановку иностранного гражданина или лица без "
            "гражданства на учёт по месту пребывания"
        ),
        "kbk": "18810806000010039110",
        "amount": "500",
    },
    "renewal": {
        "purpose": (
            "Государственная пошлина за продление срока временного пребывания иностранного "
            "гражданина в Российской Федерации"
        ),
        "kbk": "18810806000010041110",
        "amount": "1000",
    },
}

MEDICAL_SERVICE_TEXT = (
    "Медицинское освидетельствование на наличие или отсутствие инфекционных заболеваний, "
    "представляющих опасность для окружающих и являющихся основанием для отказа в выдаче "
    "либо аннулирования разрешения на временное проживание иностранных лиц и лиц без "
    "гражданства, или вида на жительство, или патента, или разрешения на работу в Российской "
    "Федерации, если иное не предусмотрено международным договором Российской Федерации, "
    "с проведением лабораторных исследований, проведением осмотра врачом-дерматовенерологом, "
    "осмотра врачом-инфекционистом, с выдачей медицинского заключения и сертификата об "
    "отсутствии у иностранного гражданина заболевания, вызываемого вирусом иммунодефицита "
    "человека (ВИЧ-инфекции)."
)


def _contract_header_parts() -> tuple[str, str, str]:
    """Разбивает CLINIC_CONTRACT_DATE (ожидается формат ДД.ММ.ГГГГ) на день/месяц/год —
    в бумажном бланке это три отдельных поля («25» 06 2026г.), не одна строка.
    [Предполагаю] формат хранения даты в env var — ДД.ММ.ГГГГ, как и в остальных местах
    проекта. Если строку не удалось разобрать — используем как есть одним куском в поле
    "день", чтобы не потерять данные и не упасть, но это не будет визуально соответствовать
    трём полям бланка."""
    try:
        parsed = datetime.strptime(CLINIC_CONTRACT_DATE, "%d.%m.%Y").date()
        return parsed.strftime("%d"), parsed.strftime("%m"), parsed.strftime("%Y")
    except ValueError:
        return CLINIC_CONTRACT_DATE, "", ""


def _format_salary(salary: str) -> str:
    """Оклад из формы (число или строка) -> '30 000' с пробелом-разделителем тысяч.
    Если не парсится как число — возвращаем как есть (кадровик мог ввести строкой)."""
    s = (salary or "").strip().replace(" ", "").replace("\xa0", "")
    if s.isdigit():
        return f"{int(s):,}".replace(",", " ")
    return salary or DASH


def generate_labor_contract_docx(
    employee: Employee,
    position: str,
    salary: str,
    contract_date: date,
    output_dir: str = "/tmp",
) -> str:
    """Трудовой договор с ООО «ТРЕСТСТРОЙМОНТАЖ».

    Переменные из карточки: ФИО, дата рождения, паспорт (серия/номер), адрес (SITE_ADDRESS),
    табельный номер (employee.tab_number). Номер договора = CONTRACT_NUMBER_PREFIX + tab_number.
    Из формы генерации: position (должность), salary (оклад), contract_date (дата договора).
    Прочерк в документе: «паспорт кем/когда выдан» (в модели нет).

    ВНИМАНИЕ: без tab_number номер договора будет обрывком (префикс без номера) — вызывающий
    код (webforms) обязан блокировать генерацию при пустом tab_number. Здесь ставим DASH,
    чтобы не молчать, если всё же вызвали без номера.
    Реквизиты ООО и текст договора приняты как факт без юридической проверки."""
    doc = Document()
    _set_default_style(doc)

    tab = (employee.tab_number or "").strip()
    contract_no = f"{CONTRACT_NUMBER_PREFIX}{tab}" if tab else DASH
    date_str = contract_date.strftime("%d.%m.%Y") if contract_date else DASH
    birth = _date_or_dash(employee.birth_date)
    passport = _passport_str(employee)
    salary_fmt = _format_salary(salary)
    position = (position or "").strip() or DASH

    def h(text):
        p = doc.add_paragraph()
        r = p.add_run(text)
        r.bold = True
        return p

    def para(text):
        return doc.add_paragraph(text)

    # Заголовок
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(f"ТРУДОВОЙ ДОГОВОР № {contract_no}")
    r.bold = True

    # Город слева, дата справа — через таблицу 1×2 без границ (табы съезжают по ширине поля).
    head = doc.add_table(rows=1, cols=2)
    head.autofit = True
    head.cell(0, 0).paragraphs[0].add_run("г. Москва")
    _rp = head.cell(0, 1).paragraphs[0]
    _rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _rp.add_run(date_str)

    para(
        f"{EMPLOYER_NAME_FULL}, именуемое в дальнейшем «Работодатель», в лице "
        f"Генерального директора {EMPLOYER_DIRECTOR_FULL}, действующего на основании Устава, "
        f"с одной стороны, и {employee.full_name}, именуемый в дальнейшем «Работник», "
        f"с другой стороны, заключили настоящий трудовой договор о нижеследующем:"
    )

    h("1. ПРЕДМЕТ И СРОК ДЕЙСТВИЯ ДОГОВОРА")
    para(
        f"1.1. Согласно настоящему договору Работник принимается на работу в "
        f"{EMPLOYER_SUBDIVISION} {EMPLOYER_NAME_SHORT} на должность {position} с {date_str} года."
    )
    para(
        "1.2. Работник обязуется выполнять все работы, обуславливаемые должностью, на которую "
        "он принимается, а также трудовыми обязанностями и конкретными заданиями (поручениями), "
        "устанавливаемыми Работодателем, и должностной инструкцией в случае ее наличия."
    )
    para(f"1.3. Место работы определено: {WORKPLACE_ADDRESS}")
    para("1.4. Работа по настоящему Договору является для Работника основным местом работы.")
    para("1.5. Срок действия настоящего трудового договора устанавливается на неопределенный срок.")

    h("2. УСЛОВИЯ ТРУДА")
    para(
        "2.1. Работнику устанавливается пятидневная рабочая неделя, восьмичасовой рабочий день. "
        "Рабочий день начинается в 9 часов 00 минут утра, если при приеме на работу в связи с "
        "производственной необходимостью не оговорен другой режим рабочего времени."
    )
    para(
        "Продолжительность перерыва для отдыха и питания составляет 1 час (шестьдесят) минут в "
        "день. Время перерыва определяется на усмотрение Работодателя в пределах между 13-00 и "
        "15-00 часами дня."
    )
    para("2.2. Работник имеет право на ежегодный оплачиваемый отпуск продолжительностью 28 календарных дней.")
    para(
        "2.3. Работодатель осуществляет обязательное социальное, медицинское и пенсионное "
        "страхования Работника в порядке, определенном действующим законодательством."
    )
    para(
        "2.4. Работодатель выплачивает Работнику пособие по временной нетрудоспособности в "
        "размере, установленном действующим законодательством РФ."
    )

    h("3. ОПЛАТА ТРУДА")
    para(
        "3.1. Согласно настоящему договору Работнику выплачивается заработная плата в "
        "соответствии со штатным расписанием. На момент заключения договора заработная плата "
        f"состоит из: Оклад: {salary_fmt} руб.; Районный коэффициент: {DISTRICT_COEFFICIENT}."
    )
    para(
        "3.2. Заработная плата выплачивается Работнику не реже чем каждые полмесяца путем "
        "перечисления денежных средств на банковский счёт Работника, реквизиты которого "
        "Работник сообщает Работодателю в письменном виде."
    )

    h("4. ПРАВА И ОБЯЗАННОСТИ СТОРОН")
    para("4.1. Работник имеет права и обязуется исполнять обязанности, предусмотренные статьей 21 ТК РФ.")
    para("4.2. Работодатель имеет права и обязуется исполнять обязанности, предусмотренные статьей 22 ТК РФ.")

    h("5. КОНФИДЕНЦИАЛЬНОСТЬ")
    para(
        "5.1. Работник обязан обеспечить сохранность, не разглашать и не передавать третьим лицам "
        "сведения и документы, составляющие служебную, коммерческую, техническую, технологическую "
        "или экономическую тайну Работодателя и его клиентов (заказчиков) как в течение срока "
        "действия настоящего Договора, так и в течение 3 лет после его прекращения."
    )

    h("6. ПОРЯДОК УРЕГУЛИРОВАНИЯ СПОРОВ")
    para(
        "6.1. Все споры и разногласия, которые могут возникнуть из настоящего трудового договора "
        "или в связи с ним, будут по возможности решаться сторонами путем переговоров. В случае "
        "недостижения согласия спор подлежит урегулированию в порядке, предусмотренном трудовым "
        "законодательством РФ."
    )

    h("7. ИНЫЕ УСЛОВИЯ")
    para(
        "7.1. Во всем остальном, что не предусмотрено настоящим трудовым договором, стороны "
        "руководствуются законодательством РФ, регулирующим трудовые отношения."
    )
    para("7.2. Настоящий договор составлен в 2 экземплярах, по одному для каждой из сторон.")

    h("8. АДРЕСА СТОРОН И ПОДПИСИ")
    # Две стороны рядом: слева Работник, справа Работодатель. Таблица 1x2 без границ.
    sig = doc.add_table(rows=1, cols=2)
    sig.autofit = True

    left = sig.cell(0, 0)
    left.paragraphs[0].add_run("Работник:").bold = True
    left.add_paragraph(f"{employee.full_name}, {birth} года рождения")
    left.add_paragraph(f"Паспорт: {passport}, выдан: {DASH}")
    left.add_paragraph(f"Адрес: {SITE_ADDRESS}")
    left.add_paragraph("")
    left.add_paragraph("Подпись: _______________")

    right = sig.cell(0, 1)
    right.paragraphs[0].add_run("Работодатель:").bold = True
    right.add_paragraph(EMPLOYER_NAME_FULL)
    right.add_paragraph(f"ИНН: {EMPLOYER_INN} КПП: {EMPLOYER_KPP}")
    right.add_paragraph(f"Юридический адрес: {EMPLOYER_LEGAL_ADDRESS}")
    right.add_paragraph(f"Фактический адрес: {EMPLOYER_ACTUAL_ADDRESS}")
    right.add_paragraph(f"Телефон: {EMPLOYER_PHONE}")
    right.add_paragraph("")
    right.add_paragraph(f"Ген. директор _______________")
    right.add_paragraph(EMPLOYER_DIRECTOR_SHORT)
    right.add_paragraph("м.п.")

    safe_tab = tab or "no_tab"
    filename = f"labor_contract_{employee.id}_{safe_tab}.docx"
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path


def _build_duty_qr_payload(kind: str) -> str:
    """Платёжная строка по ГОСТ Р 56042 для банковского QR. Порядок и состав полей —
    обязательные (Name, PersonalAcc, BankName, BIC, CorrespAcc, PayeeINN) плюс KPP, KBK,
    OKTMO, Sum (в копейках), Purpose. Sum: рубли*100."""
    spec = DUTY_KINDS[kind]
    fields = {
        "Name": DUTY_PAYEE_NAME_QR,
        "PersonalAcc": DUTY_ACCOUNT,
        "BankName": DUTY_BANK_NAME,
        "BIC": DUTY_BIC,
        "CorrespAcc": DUTY_CORRESP_ACC,
        "PayeeINN": DUTY_PAYEE_INN,
        "KPP": DUTY_PAYEE_KPP,
        "KBK": spec["kbk"],
        "OKTMO": DUTY_OKTMO,
        "Sum": str(int(spec["amount"]) * 100),
        "Purpose": spec["purpose"],
    }
    return "ST00012|" + "|".join(f"{k}={v}" for k, v in fields.items())


def _generate_duty_qr_png(kind: str, out_path: str) -> bool:
    """Строит QR-картинку платёжной строки в out_path. Возвращает True при успехе.
    Библиотека qrcode может отсутствовать (офлайн-сборка) — тогда возвращаем False,
    и квитанция генерируется без QR (не роняем весь документ из-за отсутствия картинки).
    На Railway qrcode должна быть в зависимостях (добавить 'qrcode[pil]' в requirements)."""
    try:
        import qrcode
    except ImportError:
        return False
    try:
        img = qrcode.make(_build_duty_qr_payload(kind))
        img.save(out_path)
        return True
    except Exception:
        return False


def _fix_table_width(table, widths_mm):
    """Жёстко фиксирует ширину колонок таблицы через XML, иначе Word/LibreOffice игнорируют
    cell.width и растягивают таблицу на всю ширину листа. Ставит tblLayout=fixed, общую
    ширину таблицы и tcW (в twips) на каждую ячейку. widths_mm — список ширин колонок в мм."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    def _twips(mm):
        return str(int(mm * 56.6929))  # 1 мм = 56.6929 twips

    tbl = table._tbl
    tblPr = tbl.tblPr

    # Инлайн-границы (не через стиль): мобильные docx-читалки игнорируют границы стиля
    # Table Grid, поэтому прописываем tblBorders прямо в XML — видно в любой программе.
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")       # 0.5pt
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")

    # layout = fixed
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")

    # общая ширина таблицы
    total = sum(widths_mm)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), _twips(total))
    tblW.set(qn("w:type"), "dxa")

    # tblGrid — ширины колонок
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for gc, w in zip(grid.findall(qn("w:gridCol")), widths_mm):
            gc.set(qn("w:w"), _twips(w))

    # tcW на каждую ячейку
    for row in table.rows:
        for cell, w in zip(row.cells, widths_mm):
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), _twips(w))
            tcW.set(qn("w:type"), "dxa")


def generate_duty_receipt_docx(kind: str, employee=None, payer_name: str = "", output_dir: str = "/tmp") -> str:
    """Квитанция на оплату госпошлины, форма ПД-4сб(налог). kind: "registration" (500, постановка
    на учёт) или "renewal" (1000, продление пребывания). Реквизиты фиксированные (DUTY_*), сумма
    фиксированная по типу. Плательщик (ФИО/адрес) — поля предусмотрены, но пустые: заполняются
    вручную (пользователь: ФИО пока не подставляем). employee — задел на будущее автозаполнение.

    ПД-4сб состоит из двух одинаковых частей: «Извещение» (остаётся в банке) и «Квитанция»
    (у плательщика). Обе части идентичны — строятся одной вспомогательной функцией."""
    if kind not in DUTY_KINDS:
        raise ValueError(f"Неизвестный тип пошлины: {kind!r}. Ожидается один из {list(DUTY_KINDS)}")
    spec = DUTY_KINDS[kind]

    doc = Document()
    _set_default_style(doc)

    # Компоновка на один лист A4: узкие поля + мелкий шрифт таблицы. Две части ПД-4сб (13 строк
    # каждая) + QR иначе не влезают. Шрифт таблицы 8pt — компромисс ради одного листа (запрошено).
    from docx.shared import Pt as _Pt, Inches as _In
    for sec in doc.sections:
        sec.top_margin = _In(0.5)
        sec.bottom_margin = _In(0.5)
        sec.left_margin = _In(0.6)
        sec.right_margin = _In(0.6)

    # ФИО плательщика — из формы (кадровик вводит; плательщик НЕ обязательно работник).
    payer_fio = (payer_name or "").strip() or DASH
    payer_addr = DASH  # адрес плательщика — прочерк (по решению: не вводим)
    # ФИО работника, за кого платёж, добавляется в назначение платежа.
    _worker = (getattr(employee, "full_name", "") or "").strip() if employee else ""

    # QR один раз генерируем во временный файл, вставляем в обе части (извещение и квитанция).
    import tempfile
    qr_path = os.path.join(tempfile.gettempdir(), f"duty_qr_{kind}.png")
    qr_ok = _generate_duty_qr_png(kind, qr_path)

    def _tiny(paragraph, size=10):
        for r in paragraph.runs:
            r.font.size = _Pt(size)
        # убрать вертикальные пробелы между строками таблицы
        pf = paragraph.paragraph_format
        pf.space_before = _Pt(1)
        pf.space_after = _Pt(1)
        pf.line_spacing = 1.0

    def _build_part(part_title: str) -> None:
        head = doc.add_paragraph()
        head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Заголовок части держится со следующим содержимым — часть не начинается в самом низу листа
        # с переносом таблицы/QR на новую страницу.
        head.paragraph_format.keep_with_next = True
        head.paragraph_format.space_before = _Pt(2)
        head.paragraph_format.space_after = _Pt(1)
        r = head.add_run(part_title)
        r.bold = True
        r.font.size = _Pt(13)
        sub = doc.add_paragraph()
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sr = sub.add_run("Форма № ПД-4сб (налог)")
        sr.italic = True
        sr.font.size = _Pt(9)

        rows = [
            ("Наименование получателя", DUTY_PAYEE_NAME),
            ("ИНН / КПП получателя", f"{DUTY_PAYEE_INN} / {DUTY_PAYEE_KPP}"),
            ("Код ОКТМО", DUTY_OKTMO),
            ("Счёт получателя платежа", DUTY_ACCOUNT),
            ("Банк получателя", DUTY_BANK_NAME),
            ("БИК / Кор. счёт", f"{DUTY_BIC} / {DUTY_CORRESP_ACC}"),
            ("КБК", spec["kbk"]),
            ("Наименование платежа", spec["purpose"] + (f" за {_worker}" if _worker else "")),
            ("Ф.И.О. плательщика", payer_fio),
            ("Адрес плательщика", payer_addr),
            ("Сумма платежа", f"{spec['amount']} руб. 00 коп."),
        ]
        table = doc.add_table(rows=len(rows), cols=2)
        table.style = "Table Grid"
        table.autofit = False
        table.allow_autofit = False
        for i, (label, value) in enumerate(rows):
            c0 = table.cell(i, 0).paragraphs[0]
            run = c0.add_run(label)
            run.bold = True
            _tiny(c0)
            c1 = table.cell(i, 1).paragraphs[0]
            c1.add_run(value)
            _tiny(c1)
        # Жёсткая фиксация ширины через XML: подписи 55 мм, значения 110 мм (итого 165 мм —
        # ширина печатного поля A4 при полях 0.5"/0.6"). Без этого таблица растягивается.
        _fix_table_width(table, [50, 95])

        # QR + подпись рядом: таблица 1x2 без границ. Слева QR, справа поля подписи.
        foot = doc.add_table(rows=1, cols=2)
        foot.autofit = True
        if qr_ok:
            foot.cell(0, 0).paragraphs[0].add_run().add_picture(qr_path, width=_In(1.1))
        else:
            foot.cell(0, 0).paragraphs[0].add_run("(QR недоступен)").font.size = _Pt(7)
        rc = foot.cell(0, 1)
        p1 = rc.paragraphs[0]; p1.add_run("Плательщик (подпись) _______________"); _tiny(p1)
        p1.paragraph_format.keep_together = True
        p2 = rc.add_paragraph(); p2.add_run("Дата _______________"); _tiny(p2)
        p2.paragraph_format.keep_together = True
        p3 = rc.add_paragraph(); p3.add_run("Отсканируйте QR в банковском приложении для оплаты"); _tiny(p3, 7)

    _build_part("ИЗВЕЩЕНИЕ")
    _sep = doc.add_paragraph("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
    _sep.paragraph_format.space_before = _Pt(2)
    _sep.paragraph_format.space_after = _Pt(2)
    for _r in _sep.runs:
        _r.font.size = _Pt(8)
    _build_part("КВИТАНЦИЯ")

    filename = f"duty_receipt_{kind}.docx"
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path


def generate_termination_notice_docx(employee: Employee, basis: str = "",
                                     output_dir: str = "/tmp") -> str:
    """Уведомление МВД о прекращении (расторжении) трудового договора с иностранцем.
    Структура по форме №8 приказа МВД России от 30.07.2020 №536 (действует до 01.09.2026 —
    с этой даты приказ №290 от 12.05.2026 вводит новые формы; см. задачу на обновление).

    НЕ пиксельная копия бланка: подача основная через Госуслуги (там своя форма), это документ
    для подготовки данных и архива / запасной бумажной подачи. Обязательные поля, которых нет в
    системе (ОКВЭД, ОГРН работодателя), — прочерком под ручное заполнение.

    basis — основание расторжения (по собственному желанию / инициатива работодателя /
    истечение срока и т.п.), вводится в форме увольнения. Дата расторжения — contract_end_date."""
    doc = Document()
    _set_default_style(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("УВЕДОМЛЕНИЕ\nо прекращении (расторжении) трудового договора\n"
                      "с иностранным гражданином")
    r.bold = True
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run("(форма по приложению №8 к приказу МВД России от 30.07.2020 №536)")
    sr.italic = True
    sr.font.size = Pt(9)

    doc.add_paragraph()
    doc.add_paragraph("В ____________________________________________________________")
    doc.add_paragraph("(наименование территориального органа МВД России)")
    doc.add_paragraph()

    doc.add_paragraph("1. Сведения о работодателе (заказчике работ, услуг):")
    rows = [
        ("Полное наименование", EMPLOYER_NAME_FULL),
        ("ИНН", EMPLOYER_INN),
        ("КПП", EMPLOYER_KPP),
        ("ОГРН", _text_or_dash(EMPLOYER_OGRN)),
        ("ОКВЭД (основной вид деятельности)", _text_or_dash(EMPLOYER_OKVED)),
        ("Юридический адрес", EMPLOYER_LEGAL_ADDRESS),
        ("Фактический адрес (место работы)", WORKPLACE_ADDRESS),
        ("Телефон", EMPLOYER_PHONE),
    ]
    t1 = doc.add_table(rows=len(rows), cols=2)
    t1.style = "Table Grid"
    for i, (k, v) in enumerate(rows):
        t1.cell(i, 0).paragraphs[0].add_run(k).bold = True
        t1.cell(i, 1).paragraphs[0].add_run(v)
    _fix_table_width(t1, [70, 100])

    doc.add_paragraph()
    doc.add_paragraph("2. Сведения об иностранном гражданине (лице без гражданства):")
    name_parts = (employee.full_name or "").split()
    surname = name_parts[0] if name_parts else DASH
    first = name_parts[1] if len(name_parts) > 1 else DASH
    patr = name_parts[2] if len(name_parts) > 2 else DASH
    rows2 = [
        ("Фамилия", surname),
        ("Имя", first),
        ("Отчество", patr),
        ("Гражданство", _text_or_dash(getattr(employee, "citizenship", None) or "Республика Казахстан")),
        ("Дата рождения", _date_or_dash(employee.birth_date)),
        ("Документ, удостоверяющий личность (серия, номер)", _passport_str(employee)),
    ]
    t2 = doc.add_table(rows=len(rows2), cols=2)
    t2.style = "Table Grid"
    for i, (k, v) in enumerate(rows2):
        t2.cell(i, 0).paragraphs[0].add_run(k).bold = True
        t2.cell(i, 1).paragraphs[0].add_run(v)
    _fix_table_width(t2, [70, 100])

    doc.add_paragraph()
    doc.add_paragraph("3. Сведения о трудовом (гражданско-правовом) договоре:")
    _contract_no = (CONTRACT_NUMBER_PREFIX + employee.tab_number.strip()) if (employee.tab_number or "").strip() else DASH
    rows3 = [
        ("Номер договора", _contract_no),
        ("Дата заключения договора", _date_or_dash(employee.contract_date)),
        ("Дата прекращения (расторжения)", _date_or_dash(employee.contract_end_date)),
        ("Основание прекращения (расторжения)", _text_or_dash(basis)),
    ]
    t3 = doc.add_table(rows=len(rows3), cols=2)
    t3.style = "Table Grid"
    for i, (k, v) in enumerate(rows3):
        t3.cell(i, 0).paragraphs[0].add_run(k).bold = True
        t3.cell(i, 1).paragraphs[0].add_run(v)
    _fix_table_width(t3, [70, 100])

    doc.add_paragraph()
    doc.add_paragraph("Уведомление подаётся в срок не более 3 рабочих дней с даты прекращения "
                      "(расторжения) договора (п. 8 ст. 13 Федерального закона от 25.07.2002 №115-ФЗ).")
    doc.add_paragraph()
    doc.add_paragraph(f"Руководитель: _______________ / {EMPLOYER_DIRECTOR_SHORT}")
    doc.add_paragraph(f"Дата подачи: «___» __________ 20__ г.        М.П.")

    filename = f"termination_notice_{employee.id}.docx"
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path


def generate_departure_notice_docx(employee: Employee, output_dir: str = "/tmp") -> str:
    """Уведомление об убытии иностранного гражданина из места пребывания (снятие с
    миграционного учёта). Принимающая сторона — ООО «ТРЕСТСТРОЙМОНТАЖ» (вахта, предоставляет
    жильё). Срок подачи — 7 рабочих дней с даты убытия (ст. 23 №109-ФЗ, п. 45 Правил ПП №9).

    Дата убытия = contract_end_date (дата увольнения). Адрес пребывания, откуда убыл, —
    SITE_ADDRESS (площадка). НЕ пиксельная копия бланка (подача через Госуслуги / бумага —
    запасной путь); структурированный документ со всеми обязательными сведениями."""
    doc = Document()
    _set_default_style(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("УВЕДОМЛЕНИЕ\nоб убытии иностранного гражданина\nиз места пребывания")
    r.bold = True
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run("(снятие с миграционного учёта по месту пребывания)")
    sr.italic = True
    sr.font.size = Pt(9)

    doc.add_paragraph()
    doc.add_paragraph("В ____________________________________________________________")
    doc.add_paragraph("(наименование территориального органа МВД России)")
    doc.add_paragraph()

    doc.add_paragraph("1. Сведения об иностранном гражданине, подлежащем снятию с учёта:")
    name_parts = (employee.full_name or "").split()
    surname = name_parts[0] if name_parts else DASH
    first = name_parts[1] if len(name_parts) > 1 else DASH
    patr = name_parts[2] if len(name_parts) > 2 else DASH
    rows = [
        ("Фамилия", surname),
        ("Имя", first),
        ("Отчество", patr),
        ("Гражданство", _text_or_dash(getattr(employee, "citizenship", None) or "Республика Казахстан")),
        ("Дата рождения", _date_or_dash(employee.birth_date)),
        ("Документ, удостоверяющий личность", _passport_str(employee)),
        ("Адрес места пребывания (откуда убыл)", SITE_ADDRESS),
        ("Дата убытия из места пребывания", _date_or_dash(employee.contract_end_date)),
    ]
    t1 = doc.add_table(rows=len(rows), cols=2)
    t1.style = "Table Grid"
    for i, (k, v) in enumerate(rows):
        t1.cell(i, 0).paragraphs[0].add_run(k).bold = True
        t1.cell(i, 1).paragraphs[0].add_run(v)
    _fix_table_width(t1, [70, 100])

    doc.add_paragraph()
    doc.add_paragraph("2. Сведения о принимающей стороне (организация):")
    rows2 = [
        ("Полное наименование", EMPLOYER_NAME_FULL),
        ("ИНН", EMPLOYER_INN),
        ("КПП", EMPLOYER_KPP),
        ("ОГРН", _text_or_dash(EMPLOYER_OGRN)),
        ("Адрес", EMPLOYER_LEGAL_ADDRESS),
        ("Телефон", EMPLOYER_PHONE),
        ("Ф.И.О. представителя", DASH),
        ("Документ, удостоверяющий личность представителя", DASH),
    ]
    t2 = doc.add_table(rows=len(rows2), cols=2)
    t2.style = "Table Grid"
    for i, (k, v) in enumerate(rows2):
        t2.cell(i, 0).paragraphs[0].add_run(k).bold = True
        t2.cell(i, 1).paragraphs[0].add_run(v)
    _fix_table_width(t2, [70, 100])

    doc.add_paragraph()
    doc.add_paragraph("Уведомление об убытии представляется принимающей стороной в срок 7 рабочих "
                      "дней с даты убытия (ст. 23 Федерального закона от 18.07.2006 №109-ФЗ, "
                      "п. 45 Правил, утв. постановлением Правительства РФ от 15.01.2007 №9).")
    doc.add_paragraph()
    doc.add_paragraph("Представитель принимающей стороны: _______________ / _______________")
    doc.add_paragraph("Дата подачи: «___» __________ 20__ г.        М.П.")

    filename = f"departure_notice_{employee.id}.docx"
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path


def generate_medical_referral_docx(employee: Employee, output_dir: str = "/tmp") -> str:
    """Направление на медицинское освидетельствование — форма ГОАУЗ «МОМЦ» (Приложение №1
    к договору), заполняется по факту согласования конкретной даты/кабинета с клиникой.

    Дата приёма, номер кабинета и время намеренно оставлены пустыми полями для ручного
    заполнения — это отдельный процесс согласования с клиникой, бот не может знать
    расписание клиники заранее и не должен его придумывать."""
    missing = _require_fields(employee, MEDICAL_REFERRAL_REQUIRED_FIELDS)

    doc = Document()
    _set_default_style(doc)

    if missing:
        _add_test_warning_paragraph(doc, missing)

    day, month, year = _contract_header_parts()
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.add_run(
        f"к Договору № {CLINIC_CONTRACT_NUMBER} от «{day}» {month} {year}г.\nПриложение № 1"
    )

    doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("НАПРАВЛЕНИЕ НА МЕДИЦИНСКОЕ ОСВИДЕТЕЛЬСТВОВАНИЕ")
    run.bold = True
    run.font.size = Pt(13)

    doc.add_paragraph()
    doc.add_paragraph(f"В {CLINIC_NAME}")
    doc.add_paragraph("наименование медицинской организации (МО)")

    name_parts = employee.full_name.split()
    surname = name_parts[0] if name_parts else DASH
    first_name = name_parts[1] if len(name_parts) > 1 else DASH
    patronymic = name_parts[2] if len(name_parts) > 2 else DASH

    doc.add_paragraph(f"1. Фамилия {surname}")
    doc.add_paragraph(f"Имя {first_name}")
    doc.add_paragraph(f"Отчество {patronymic}")

    birth = _date_or_dash(employee.birth_date)
    doc.add_paragraph(f"2. Дата рождения (число, месяц, год) {birth}")

    doc.add_paragraph(f"3. Адрес (по месту проживания) {SITE_ADDRESS}")

    doc.add_paragraph(
        f"4. Серия паспорта {_text_or_dash(employee.passport_series)} "
        f"Номер паспорта {_text_or_dash(employee.passport_number)}"
    )

    doc.add_paragraph(f"5. Место работы {PAYER_NAME}")

    doc.add_paragraph("6. Наименование медицинской услуги (медицинского освидетельствования)")
    doc.add_paragraph(MEDICAL_SERVICE_TEXT)

    doc.add_paragraph("7. Дата проведения услуги _____________ кабинет N _____ время _____")

    doc.add_paragraph(
        "8. Полное наименование организации, направившей иностранного гражданина, "
        f"телефон {PAYER_PHONE}"
    )
    doc.add_paragraph(PAYER_NAME)
    doc.add_paragraph("подпись, печать _____________________")

    doc.add_paragraph(f"10. Дата выдачи направления {datetime.now(MSK).date().strftime('%d.%m.%Y')}")

    doc.add_paragraph()

    # Блок подписей "От Исполнителя / От Заказчика" — был в бумажном бланке, отсутствовал
    # в генераторе. Таблица 2×2 без границ, чтобы визуально повторить два столбца оригинала.
    table = doc.add_table(rows=4, cols=2)
    table.autofit = True

    def _cell(row: int, col: int, text: str, bold: bool = False) -> None:
        cell = table.cell(row, col)
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.bold = bold

    _cell(0, 0, "От Исполнителя:", bold=True)
    _cell(0, 1, "От Заказчика:", bold=True)
    _cell(1, 0, f"{CLINIC_SHORT_NAME}")
    _cell(1, 1, "Индивидуальный предприниматель")
    _cell(2, 0, "Главный врач")
    _cell(2, 1, "")
    _cell(3, 0, f"_____________________ {CLINIC_CHIEF_DOCTOR_NAME}\nм.п.")
    _cell(3, 1, f"_____________________ {PAYER_SIGNATORY_NAME}\nм.п.")

    filename = f"medical_referral_{employee.id}.docx"
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path


def generate_employees_xlsx(employees: list[Employee], output_dir: str = "/tmp") -> str:
    """Полный список сотрудников таблицей — замена постраничному тексту в /employees.
    Отдельный файл от Google Sheets: тот обновляется по крону раз в сутки командой
    export_to_sheets_api.py, этот — генерируется по запросу с текущим состоянием БД."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Сотрудники"

    headers = [
        "ФИО", "Гражданство", "Категория", "Дата въезда", "Дата договора",
        "Статус занятости", "Согласие", "Телефон", "Язык",
        "Дата рождения", "Серия паспорта", "Номер паспорта", "Адрес", "Откуда въехал",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for emp in employees:
        ws.append([
            emp.full_name,
            emp.citizenship or "",
            emp.category.value if emp.category else "",
            emp.entry_date.strftime("%d.%m.%Y") if emp.entry_date else "",
            emp.contract_date.strftime("%d.%m.%Y") if emp.contract_date else "",
            emp.employment_status or "",
            "да" if emp.consent_status.value == "confirmed" else "нет",
            emp.phone or "",
            emp.language or "",
            emp.birth_date.strftime("%d.%m.%Y") if emp.birth_date else "",
            emp.passport_series or "",
            emp.passport_number or "",
            emp.address or "",
            emp.entry_country or "",
        ])

    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    path = os.path.join(output_dir, "employees.xlsx")
    wb.save(path)
    return path


# ============================================================================
# НАРЯД-ДОПУСК НА ПРОИЗВОДСТВО РАБОТ НА ВЫСОТЕ
# (Приложение № 2 к Правилам по охране труда при работе на высоте,
#  Приказ Минтруда России от 16.11.2020 № 782н)
#
# v1 (согласовано): читает существующую модель WorkOrder. Упрощения v1:
#   1. Время работ фиксировано (WORK_ORDER_START_TIME/END_TIME = 08:00/19:00) —
#      в модели времени нет, только даты valid_from/valid_to.
#   2. safety_systems — одно текстовое поле; раскладывается по 3 строкам таблицы
#      по переносам строк (недостающие строки — прочерком, лишние — в последнюю).
# Справочники типовых работ/титулов и автонумерация — отдельными шагами.
# ============================================================================

import re as _re

WORK_ORDER_ORG_NAME = os.environ.get("WORK_ORDER_ORG_NAME") or COMPANY_NAME
# Лицо, выдающее наряд (фиксированное, назначается приказом). По умолчанию —
# наименование организации; переопределяется, если выдающий — назначенное лицо.
WORK_ORDER_ISSUER_NAME = os.environ.get("WORK_ORDER_ISSUER_NAME") or WORK_ORDER_ORG_NAME
WORK_ORDER_DEFAULT_SUBDIVISION = os.environ.get("WORK_ORDER_DEFAULT_SUBDIVISION") or "Мурманск"
WORK_ORDER_START_TIME = "08:00"
WORK_ORDER_END_TIME = "19:00"
WORK_ORDER_MAX_DAYS = 15  # 782н: срок действия ≤ 15 календарных дней (+1 продление ≤15)

# Раздел 2 «Мероприятия до начала работ» — типовой текст (одинаков всегда).
WORK_ORDER_PREP_MEASURES = [
    "Оформить наряд-допуск на работы повышенной опасности с обязательным указанием в нём: "
    "ответственного исполнителя работ; ответственного руководителя работ; место выполнения "
    "работ на высоте находится в зоне прямой видимости ответственного исполнителя работ "
    "и/или ответственного руководителя работ.",
    "Ознакомление и обсуждение Плана производства работ (технологической карты) с "
    "ответственным руководителем работ, ответственным исполнителем работ, исполнителями работ.",
    "Обсуждение начала рабочего процесса, разъяснение ответственным руководителем работ всех "
    "специфических обязанностей и процедур всем работникам и соблюдение правил безопасности.",
    "Работники, впервые допускаемые к работам на высоте, должны обладать практическими "
    "навыками применения оборудования, приборов, механизмов и оказания первой помощи "
    "пострадавшим, практическими навыками применения соответствующих СИЗ, их осмотром до и "
    "после использования.",
    "Средства коллективной и индивидуальной защиты работников должны использоваться по "
    "назначению в соответствии с требованиями инструкций изготовителя и нормативной технической "
    "документации, введённой в действие в установленном порядке.",
]


def _height_group_num(value):
    """Из строки группы («2-я гр. по безопасности работ на высоте») достаёт число 2."""
    if not value:
        return None
    m = _re.search(r"(\d+)", str(value))
    return int(m.group(1)) if m else None


def _wo_set_cell(cell, text, *, bold=False, size=9, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    run = p.add_run("" if text in (None, "") else str(text))
    run.bold = bold
    run.font.size = Pt(size)


def _wo_split_safety_systems(raw):
    """safety_systems (одно текстовое поле) -> 3 значения строк таблицы.
    Меньше 3 строк -> недостающие прочерком; больше 3 -> лишние склеиваются в последнюю."""
    lines = [ln.strip() for ln in (raw or "").splitlines() if ln.strip()]
    vals = []
    for i in range(3):
        if i < len(lines):
            vals.append(" ".join(lines[2:]) if (i == 2 and len(lines) > 3) else lines[i])
        else:
            vals.append(DASH)
    return vals


def check_work_order_problems(work_order):
    """Список проблем наряда-допуска (782н + наши правила). Пустой список = ок.
    Документ строится ВСЕГДА (черновик); блокировать выпуск по этому списку — задача
    вызывающего кода (webforms.py/bot.py) через WorkOrderStatus.DRAFT."""
    problems = []
    sup = getattr(work_order, "responsible_supervisor", None)
    ex = getattr(work_order, "responsible_executor", None)

    sup_grp = _height_group_num(getattr(sup, "height_safety_group", None)) if sup else None
    if sup_grp != 3:
        problems.append(
            "Ответственный руководитель работ должен быть 3-й группы по безопасности работ на "
            f"высоте (сейчас: {getattr(sup, 'height_safety_group', None) or DASH})."
        )
    ex_grp = _height_group_num(getattr(ex, "height_safety_group", None)) if ex else None
    if ex_grp is None or ex_grp < 2:
        problems.append(
            "Ответственный исполнитель работ должен быть не ниже 2-й группы "
            f"(сейчас: {getattr(ex, 'height_safety_group', None) or DASH})."
        )

    members = list(getattr(work_order, "members", None) or [])
    if not members:
        problems.append("Состав бригады пуст — нельзя выпустить наряд без исполнителей.")
    for m in members:
        emp = getattr(m, "employee", None)
        name = getattr(emp, "full_name", None)
        if not emp or not name:
            problems.append("В бригаде есть член без привязанного сотрудника (пустая строка).")
            continue
        g = _height_group_num(getattr(emp, "height_safety_group", None))
        if g is None or g < 2:
            problems.append(f"У члена бригады «{name}» не указана группа по высоте (нужна ≥2-й).")

    try:
        span = (work_order.valid_to - work_order.valid_from).days + 1
        if span > WORK_ORDER_MAX_DAYS:
            problems.append(f"Срок действия наряда {span} дн. превышает 15 календарных дней (782н).")
        if span < 1:
            problems.append("Дата окончания раньше даты начала.")
    except Exception:
        problems.append("Не заданы корректные даты периода (valid_from/valid_to).")

    rescue = _wo_split_safety_systems(getattr(work_order, "safety_systems", None))[2].lower()
    if rescue != DASH.lower() and ("привяз" in rescue or "строп" in rescue or "фал" in rescue):
        problems.append(
            "Строка «Эвакуационные и спасательные системы» указывает страховочную привязь/строп/"
            "фал — нужно реальное средство спасения (например, автогидроподъёмник)."
        )
    return problems


def generate_work_order_docx(session, work_order, output_dir="/tmp"):
    """Наряд-допуск на работы на высоте (Приложение № 2 к 782н) из WorkOrder.
    Возвращает (path, problems): документ строится ВСЕГДА (черновик), problems — список
    нарушений (см. check_work_order_problems); блокировать выпуск по нему — дело вызывающего
    кода. session принимается для единообразия и на случай подгрузки связей."""
    problems = check_work_order_problems(work_order)

    doc = Document()
    _set_default_style(doc)
    doc.styles["Normal"].font.size = Pt(11)

    def _p(text="", *, bold=False, italic=False, center=False, size=11):
        p = doc.add_paragraph()
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        return p

    def _label(label, value):
        p = doc.add_paragraph()
        r = p.add_run(label)
        r.bold = True
        r.font.size = Pt(11)
        r2 = p.add_run(value if value not in (None, "") else DASH)
        r2.font.size = Pt(11)
        return p

    def _grid(headers, rows):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"
        for i, h in enumerate(headers):
            _wo_set_cell(t.rows[0].cells[i], h, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row in rows:
            cells = t.add_row().cells
            for i, val in enumerate(row):
                _wo_set_cell(cells[i], val, size=9)
        return t

    sup = getattr(work_order, "responsible_supervisor", None)
    ex = getattr(work_order, "responsible_executor", None)
    subdivision = work_order.subdivision or WORK_ORDER_DEFAULT_SUBDIVISION
    sup_name = getattr(sup, "full_name", None) or "____________"
    ex_name = getattr(ex, "full_name", None) or "____________"

    # Тексты, зависящие от вида работ, берутся из связанного WorkType (если наряд заполнен
    # из справочника). Собственное поле наряда имеет приоритет, справочник — запасной источник.
    wt = getattr(work_order, "work_type", None)
    work_name = work_order.work_description or getattr(wt, "name", None)
    content_val = getattr(wt, "content", None) or work_order.work_description
    conditions_val = work_order.special_conditions or getattr(wt, "conditions", None)
    hazards_val = getattr(work_order, "hazards", None) or getattr(wt, "hazards", None)
    norms_val = getattr(wt, "norms", None)

    _p("Приложение № 2 к Правилам по охране труда при работе на высоте", italic=True, size=9)
    _p("(Приказ Минтруда России от 16.11.2020 № 782н)", italic=True, size=9)
    _p()
    _p("УТВЕРЖДАЮ:", bold=True)
    _p(WORK_ORDER_ORG_NAME)
    _p("_________________ / ____________")
    _p("«____» _____________ 20___ г.", italic=True)
    _p()
    _p(f"НАРЯД-ДОПУСК № {work_order.number or DASH}", bold=True, center=True, size=13)
    _p("НА ПРОИЗВОДСТВО РАБОТ НА ВЫСОТЕ", bold=True, center=True)
    _p()

    _label("Организация: ", WORK_ORDER_ORG_NAME)
    _label("Подразделение: ", subdivision)
    _label("Выдан ", f"{_date_or_dash(work_order.valid_from)} года")
    _label("Действителен до ", f"{_date_or_dash(work_order.valid_to)} года")
    _label("Ответственному руководителю работ: ", getattr(sup, "full_name", None))
    _label("Ответственному исполнителю (производителю) работ: ", getattr(ex, "full_name", None))
    _label("На выполнение работ: ", work_name)

    _p()
    _p("Состав исполнителей работ (члены бригады):")
    member_rows = []
    for i, m in enumerate(getattr(work_order, "members", None) or [], 1):
        emp = getattr(m, "employee", None)
        name = getattr(emp, "full_name", None) or DASH
        pos = getattr(emp, "position", None) or ""
        grp = getattr(emp, "height_safety_group", None) or ""
        pos_grp = ", ".join([x for x in (pos, grp) if x]) or DASH
        member_rows.append([str(i), name, pos_grp, "", ""])
    if not member_rows:
        member_rows = [["", DASH, DASH, "", ""]]
    _grid(
        ["№", "Фамилия, имя, отчество", "Должность (разряд)",
         "Инструктаж провёл (подпись)", "Ознакомлен (подпись)"],
        member_rows,
    )

    _p()
    _label("Место выполнения работ: ", work_order.location)
    _label("Содержание работ: ", content_val)
    _label("Условия проведения работ: ", conditions_val)
    _label(
        "Опасные и вредные производственные факторы, которые действуют или могут возникнуть "
        "в местах выполнения работ: ",
        hazards_val,
    )
    _label("Начало работ: ", f"{WORK_ORDER_START_TIME} {_date_or_dash(work_order.valid_from)}")
    _label("Окончание работ: ", f"{WORK_ORDER_END_TIME} {_date_or_dash(work_order.valid_to)}")
    if norms_val:
        _label("Нормативные основания: ", norms_val)

    _p()
    _p("Системы обеспечения безопасности работ на высоте:", bold=True)
    if (work_order.safety_systems or "").strip():
        ss = _wo_split_safety_systems(work_order.safety_systems)
    elif wt is not None:
        ss = [
            getattr(wt, "sys_restraint", None) or DASH,
            getattr(wt, "sys_fall_arrest", None) or DASH,
            getattr(wt, "sys_rescue", None) or DASH,
        ]
    else:
        ss = [DASH, DASH, DASH]
    _grid(
        ["Системы обеспечения безопасности", "Состав системы"],
        [
            ["Удерживающие системы", ss[0]],
            ["Страховочные системы", ss[1]],
            ["Эвакуационные и спасательные системы", ss[2]],
        ],
    )

    _p()
    _p("1. Необходимые для производства работ:", bold=True)
    for lbl, val in [
        ("Материалы: ", work_order.materials),
        ("Инструмент: ", work_order.tools),
        ("Приспособления: ", work_order.equipment),
        ("Спецтехника: ", work_order.special_machinery),
        ("Шифр ТК: ", work_order.technological_card_ref),
    ]:
        if val:
            _label(lbl, val)

    _p()
    _p("2. До начала работ следует выполнить следующие мероприятия:", bold=True)
    _grid(
        ["Наименование мероприятия", "Срок выполнения", "Ответственный исполнитель"],
        [[m, "До начала работ", ""] for m in WORK_ORDER_PREP_MEASURES],
    )

    _p()
    _p("3. В процессе производства работ необходимо выполнить следующие мероприятия:", bold=True)
    proc_lines = [ln.strip() for ln in (getattr(wt, "process_measures", None) or "").splitlines() if ln.strip()]
    if proc_lines:
        proc_rows = [[m, "Постоянно в процессе работ", ex_name] for m in proc_lines]
    else:
        proc_rows = [["", "", ""] for _ in range(3)]
    _grid(["Наименование мероприятия", "Срок выполнения", "Ответственный исполнитель"], proc_rows)

    _p()
    _p("4. Особые условия проведения работ:", bold=True)
    _grid(["Наименование условий", "Срок выполнения", "Ответственный исполнитель"],
          [["", "", ""] for _ in range(2)])

    _p()
    _p("Отдельные указания: _______________________________________________")
    _p(f"Наряд выдал: ______________ (дата, время)   Подпись: ____________ / {WORK_ORDER_ISSUER_NAME}")
    _p("Наряд продлил: ____________ (дата, время)   Подпись: ____________ / ____________")

    _p()
    _p("5. Разрешение на подготовку рабочих мест и на допуск к выполнению работ:", bold=True)
    _grid(["Разрешение на подготовку и допуск получил", "Дата, время", "Подпись"], [["", "", ""]])
    _p("Рабочие места подготовлены. Ответственный руководитель работ: ____________ / " + sup_name)

    _p()
    _p("6. Ежедневный допуск к работе и время её окончания:", bold=True)
    day_rows = []
    try:
        d = work_order.valid_from
        while d <= work_order.valid_to:
            day_rows.append([d.strftime("%d.%m.%Y"), "", ""])
            d = d + timedelta(days=1)
    except Exception:
        pass
    if not day_rows:
        day_rows = [["", "", ""]]
    _grid(
        ["Дата", "Бригада получила целевой инструктаж и допущена (дата, время, подпись)",
         "Работа закончена, бригада удалена (дата, время, подпись)"],
        day_rows,
    )

    _p()
    _p("7. Изменения в составе бригады:", bold=True)
    _grid(["Введён в состав (ФИО)", "Выведен из состава (ФИО)", "Дата, время", "Разрешил (ФИО, подпись)"],
          [["", "", "", ""] for _ in range(3)])

    _p()
    _p("8. Регистрация целевого инструктажа при первичном допуске:", bold=True)
    _p("Инструктаж провёл: ____________ / " + sup_name)
    _p("Инструктаж прошёл: ____________ / ____________")
    _p("Лицо, выдавшее наряд: ____________ / " + WORK_ORDER_ISSUER_NAME)
    _p("Ответственный руководитель работ: ____________ / " + sup_name)
    _p("Ответственный исполнитель: ____________ / " + ex_name)

    _p()
    _p("9. Письменное разрешение (акт-допуск) действующего предприятия на производство работ "
       "имеется. Мероприятия по безопасности согласованы:", bold=True)
    _p("__________________________________________ (должность, ФИО, подпись)")

    _p()
    _p("10. Рабочее место и условия труда проверены. Мероприятия по безопасности выполнены. "
       "Разрешаю приступить к выполнению работ:", bold=True)
    _p("_______________________ (дата, подпись) ____________________ (ФИО)")
    _p("Наряд-допуск продлён до: ______________ (дата, подпись) ____________ (ФИО)")

    _p()
    _p("11. Работа выполнена в полном объёме. Материалы, инструмент, приспособления убраны. "
       "Члены бригады выведены.", bold=True)
    _p("Ответственный исполнитель (производитель) работ: ____________ (дата, подпись)")
    _p("Наряд-допуск закрыт.")
    _p("Ответственный руководитель работ: ____________ (дата, подпись)     "
       "Лицо, выдавшее наряд-допуск: ____________ (дата, подпись)")

    filename = f"work_order_{work_order.number or 'draft'}.docx".replace("/", "-").replace(" ", "_")
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path, problems
