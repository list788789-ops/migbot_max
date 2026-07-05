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

Требуемые переменные окружения:
  COMPANY_NAME             — полное наименование юрлица-работодателя
  COMPANY_INN              — ИНН
  COMPANY_LEGAL_ADDRESS    — юридический адрес
  HR_SIGNATORY_NAME        — ФИО подписанта со стороны работодателя
  HR_SIGNATORY_POSITION    — должность подписанта
  CLINIC_NAME              — наименование медицинской организации
  CLINIC_CONTRACT_NUMBER   — номер договора с клиникой
  CLINIC_CONTRACT_DATE     — дата договора с клиникой, формат ДД.ММ.ГГГГ
  CLINIC_CHIEF_DOCTOR_NAME — ФИО главного врача клиники (для блока подписи "От Исполнителя")
  PAYER_NAME               — заказчик услуги (напр. "ИП Буц С.Ю.") — используется в п.5/8 бланка
  PAYER_SIGNATORY_NAME     — ФИО подписанта со стороны заказчика для блока подписи
                             (напр. "С. Ю. Буц") — если не задано, используется PAYER_NAME
  PAYER_PHONE              — телефон заказчика
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


def _set_default_style(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)


def _passport_str(employee: Employee) -> str:
    passport = f"{employee.passport_series or ''} {employee.passport_number or ''}".strip()
    return passport or "[паспортные данные не указаны]"


def _require_fields(employee: Employee, field_labels: dict[str, str]) -> None:
    """Поднимает ValueError с точным списком незаполненных полей, вместо того чтобы
    молча вставить в документ текст-плейсхолдер. field_labels: {атрибут: человекочитаемое имя}."""
    missing = [
        label for attr, label in field_labels.items()
        if getattr(employee, attr) in (None, "")
    ]
    if missing:
        raise ValueError(
            f"Нельзя сгенерировать документ для {employee.full_name} — "
            f"не заполнены поля: {', '.join(missing)}. "
            f"Заполните их в карточке сотрудника перед генерацией."
        )


def generate_consent_docx(employee: Employee, output_dir: str = "/tmp") -> str:
    """Согласие на обработку персональных данных — отдельный документ (152-ФЗ)."""
    _require_fields(employee, {
        "birth_date": "дата рождения",
        "passport_series": "серия паспорта",
        "passport_number": "номер паспорта",
        "address": "адрес места пребывания",
    })

    doc = Document()
    _set_default_style(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("СОГЛАСИЕ\nна обработку персональных данных")
    run.bold = True
    run.font.size = Pt(14)

    doc.add_paragraph()

    birth = employee.birth_date.strftime("%d.%m.%Y")
    address = employee.address

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


CLINIC_NAME = os.environ.get("CLINIC_NAME", "[НЕ ЗАПОЛНЕНО — наименование медицинской организации]")
CLINIC_CONTRACT_NUMBER = os.environ.get("CLINIC_CONTRACT_NUMBER", "[номер договора не указан]")
CLINIC_CONTRACT_DATE = os.environ.get("CLINIC_CONTRACT_DATE", "[дата договора не указана]")
CLINIC_CHIEF_DOCTOR_NAME = os.environ.get("CLINIC_CHIEF_DOCTOR_NAME", "[ФИО главного врача не указано]")
PAYER_NAME = os.environ.get("PAYER_NAME", "[НЕ ЗАПОЛНЕНО — заказчик услуги, ИП/юрлицо]")
PAYER_SIGNATORY_NAME = os.environ.get("PAYER_SIGNATORY_NAME", PAYER_NAME)
PAYER_PHONE = os.environ.get("PAYER_PHONE", "[телефон заказчика не указан]")

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


def generate_medical_referral_docx(employee: Employee, output_dir: str = "/tmp") -> str:
    """Направление на медицинское освидетельствование — форма ГОАУЗ «МОМЦ» (Приложение №1
    к договору), заполняется по факту согласования конкретной даты/кабинета с клиникой.

    Дата приёма, номер кабинета и время намеренно оставлены пустыми полями для ручного
    заполнения — это отдельный процесс согласования с клиникой, бот не может знать
    расписание клиники заранее и не должен его придумывать."""
    _require_fields(employee, {
        "birth_date": "дата рождения",
        "passport_series": "серия паспорта",
        "passport_number": "номер паспорта",
        "address": "адрес места пребывания",
    })

    doc = Document()
    _set_default_style(doc)

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
    surname = name_parts[0] if name_parts else "[фамилия не указана]"
    first_name = name_parts[1] if len(name_parts) > 1 else "[имя не указано]"
    patronymic = name_parts[2] if len(name_parts) > 2 else "[отчество не указано]"

    doc.add_paragraph(f"1. Фамилия {surname}")
    doc.add_paragraph(f"Имя {first_name}")
    doc.add_paragraph(f"Отчество {patronymic}")

    birth = employee.birth_date.strftime("%d.%m.%Y")
    doc.add_paragraph(f"2. Дата рождения (число, месяц, год) {birth}")

    doc.add_paragraph(f"3. Адрес (по месту проживания) {employee.address}")

    doc.add_paragraph(
        f"4. Серия паспорта {employee.passport_series} "
        f"Номер паспорта {employee.passport_number}"
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
    _cell(1, 0, f"{CLINIC_NAME}")
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
