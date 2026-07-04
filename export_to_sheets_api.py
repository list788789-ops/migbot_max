"""
Экспорт данных сотрудников из БД бота напрямую в Google Sheets через API.

Требует ДО запуска (см. README):
  - credentials.json (ключ сервис-аккаунта) в этой же папке
  - таблица Google Sheets расшарена сервис-аккаунту с правом "Редактор"
  - в .env заполнен GOOGLE_SHEET_ID (ID из URL таблицы) и GOOGLE_CREDENTIALS_PATH

Как и CSV-версия — это односторонний экспорт БД -> Sheets. Правки в самой
таблице не попадают обратно в БД и будут перезаписаны при следующем запуске.

Запуск: python export_to_sheets_api.py
Для автообновления — повесить на крон (напр. раз в день) на том же хостинге,
где крутится bot.py, т.к. именно там есть сеть и доступ к прод-БД.
"""

import os

import json

import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from models import Employee, Obligation, Consent

load_dotenv()

DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./migbot.db")
SHEET_ID = os.environ["GOOGLE_SHEET_ID"]  # обязателен, без него не понятно, куда писать
CREDENTIALS_PATH = os.environ.get("GOOGLE_CREDENTIALS_PATH", "credentials.json")
CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON")  # для Railway — содержимое файла целиком
WORKSHEET_NAME = os.environ.get("GOOGLE_SHEET_WORKSHEET", "Лист1")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

HEADERS = [
    "ФИО", "Гражданство", "Категория", "Дата въезда", "Дата договора",
    "Статус занятости", "Согласие получено", "Ближайший дедлайн", "Телефон", "Язык",
]


def fetch_rows():
    engine = create_engine(DATABASE_URL)
    with Session(engine) as session:
        employees = session.query(Employee).all()
        rows = []
        for emp in employees:
            consent = (
                session.query(Consent)
                .filter_by(employee_id=emp.id)
                .order_by(Consent.confirmed_at.desc())
                .first()
            )
            obligations = session.query(Obligation).filter_by(employee_id=emp.id).all()
            nearest_deadline = min((o.deadline_date for o in obligations), default=None)

            rows.append([
                emp.full_name,
                emp.citizenship,
                emp.category.value,
                emp.entry_date.strftime("%d.%m.%Y") if emp.entry_date else "",
                emp.contract_date.strftime("%d.%m.%Y") if emp.contract_date else "",
                emp.employment_status.value,
                "да" if consent else "нет",
                nearest_deadline.strftime("%d.%m.%Y") if nearest_deadline else "",
                emp.phone or "",
                emp.language,
            ])
    return rows


def push_to_sheets(rows):
    if CREDENTIALS_JSON:
        # Railway и подобные хостинги без удобной загрузки файлов — секрет лежит в переменной окружения
        info = json.loads(CREDENTIALS_JSON)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        # Локальный запуск / хостинги с файловой системой — путь к скачанному ключу
        creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)

    client = gspread.authorize(creds)

    sheet = client.open_by_key(SHEET_ID)
    try:
        worksheet = sheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        worksheet = sheet.add_worksheet(title=WORKSHEET_NAME, rows=len(rows) + 10, cols=len(HEADERS))

    # Полная перезапись листа — проще и надёжнее, чем построчный diff/update.
    # Для 67 сотрудников это доли секунды, не повод оптимизировать сейчас.
    worksheet.clear()
    worksheet.update([HEADERS] + rows)


if __name__ == "__main__":
    rows = fetch_rows()
    push_to_sheets(rows)
    print(f"Обновлено в Google Sheets: {len(rows)} записей")
