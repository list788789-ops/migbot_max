"""
Одноразовый перенос данных из ручной xlsx-таблицы в Postgres.

Запуск: python import_xlsx_to_db.py путь_к_файлу.xlsx

Требует DATABASE_URL в окружении (тот же, что использует bot.py).
Идемпотентность НЕ реализована намеренно — предполагается разовый запуск на пустой
таблице employees. Повторный запуск создаст дубликаты.
"""

import sys
from datetime import date, datetime

import openpyxl
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from models import Base, Employee

COLUMNS = [
    "n", "full_name", "migration_card_expiry", "registration_deadline_note",
    "medical_exam_note", "citizenship", "category", "birth_date_raw",
    "passport_series", "passport_number", "address", "entry_date_raw",
    "contract_date_raw", "phone", "language", "employment_status",
    "consent_raw", "entry_country",
]


def parse_ru_date(value) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        for fmt in ("%d.%m.%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
    return None  # значение не распознано как дата (например, смешанный текст) — не роняем импорт


def main(xlsx_path: str):
    import os

    database_url = os.environ["DATABASE_URL"]
    engine = create_engine(database_url)
    Base.metadata.create_all(engine)  # создаст новые колонки, если таблицы ещё нет

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    inserted, skipped = 0, 0
    with Session(engine) as session:
        for raw_row in ws.iter_rows(min_row=2, values_only=True):
            row = dict(zip(COLUMNS, raw_row))

            full_name = (row["full_name"] or "").strip()
            if not full_name:
                skipped += 1
                continue  # строка без ФИО — пропускаем, это не сотрудник (пустая/служебная строка)

            category_raw = (row["category"] or "").strip().lower() or None
            citizenship = (row["citizenship"] or "").strip() or None

            emp = Employee(
                full_name=full_name,
                citizenship=citizenship,
                category=category_raw,  # None пройдёт как NULL, значение validируется Enum на уровне БД
                entry_date=parse_ru_date(row["entry_date_raw"]),
                contract_date=parse_ru_date(row["contract_date_raw"]),
                birth_date=parse_ru_date(row["birth_date_raw"]),
                passport_series=(row["passport_series"] or None),
                passport_number=(str(row["passport_number"]) if row["passport_number"] else None),
                address=(row["address"] or None),
                entry_country=(row["entry_country"] or None),
                registration_deadline_note=(row["registration_deadline_note"] or None),
                medical_exam_note=(row["medical_exam_note"] or None),
                employment_status=(row["employment_status"] or None),
                phone=(str(row["phone"]) if row["phone"] else None),
                language=(row["language"] or "ru"),
                consent_status="confirmed" if (row["consent_raw"] or "").strip().lower() == "да" else "draft",
            )
            session.add(emp)
            inserted += 1

        session.commit()

    print(f"Импортировано: {inserted}, пропущено (без ФИО): {skipped}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Использование: python import_xlsx_to_db.py путь_к_файлу.xlsx")
        sys.exit(1)
    main(sys.argv[1])
