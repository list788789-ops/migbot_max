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

Требуемые переменные окружения:
  COMPANY_NAME             — полное наименование юрлица-работодателя
  COMPANY_INN              — ИНН
  COMPANY_LEGAL_ADDRESS    — юридический адрес
  HR_SIGNATORY_NAME        — ФИО подписанта со стороны работодателя
  HR_SIGNATORY_POSITION    — должность подписанта
"""

import os
from datetime import date

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from models import Employee

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


def generate_consent_docx(employee: Employee, output_dir: str = "/tmp") -> str:
    """Согласие на обработку персональных данных — отдельный документ (152-ФЗ)."""
    doc = Document()
    _set_default_style(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("СОГЛАСИЕ\nна обработку персональных данных")
    run.bold = True
    run.font.size = Pt(14)

    doc.add_paragraph()

    birth = employee.birth_date.strftime("%d.%m.%Y") if employee.birth_date else "[дата рождения не указана]"
    address = employee.address or "[адрес не указан]"

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
    doc.add_paragraph(f"Дата: {date.today().strftime('%d.%m.%Y')}")
    doc.add_paragraph(f"Подпись: _____________________ / {employee.full_name}")

    filename = f"consent_{employee.id}.docx"
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path


CLINIC_NAME = os.environ.get("CLINIC_NAME", "[НЕ ЗАПОЛНЕНО — наименование медицинской организации]")
CLINIC_CONTRACT_NUMBER = os.environ.get("CLINIC_CONTRACT_NUMBER", "[номер договора не указан]")
CLINIC_CONTRACT_DATE = os.environ.get("CLINIC_CONTRACT_DATE", "[дата договора не указана]")
PAYER_NAME = os.environ.get("PAYER_NAME", "[НЕ ЗАПОЛНЕНО — заказчик услуги, ИП/юрлицо]")
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


def generate_medical_referral_docx(employee: Employee, output_dir: str = "/tmp") -> str:
    """Направление на медицинское освидетельствование — форма ГОАУЗ «МОМЦ» (Приложение №1
    к договору), заполняется по факту согласования конкретной даты/кабинета с клиникой.

    Дата приёма, номер кабинета и время намеренно оставлены пустыми полями для ручного
    заполнения — это отдельный процесс согласования с клиникой, бот не может знать
    расписание клиники заранее и не должен его придумывать."""
    doc = Document()
    _set_default_style(doc)

    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.add_run(
        f"к Договору № {CLINIC_CONTRACT_NUMBER} от «{CLINIC_CONTRACT_DATE}»\nПриложение № 1"
    )

    doc.add_paragraph("ШАБЛОН")

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

    birth = employee.birth_date.strftime("%d.%m.%Y") if employee.birth_date else "[дата рождения не указана]"
    doc.add_paragraph(f"2. Дата рождения (число, месяц, год) {birth}")

    address = employee.address or "[адрес не указан]"
    doc.add_paragraph(f"3. Адрес (по месту проживания) {address}")

    doc.add_paragraph(
        f"4. Серия паспорта {employee.passport_series or '[не указана]'} "
        f"Номер паспорта {employee.passport_number or '[не указан]'}"
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

    doc.add_paragraph(f"10. Дата выдачи направления {date.today().strftime('%d.%m.%Y')}")

    filename = f"medical_referral_{employee.id}.docx"
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path
