"""
РАЗОВЫЙ скрипт: переносит историю отметок явки из Google Sheets (бот
ТабельБелокаменка, лист «Июль» — единственный реально заполненный на
момент переноса) в таблицу attendance_marks в базе migbot.

Не часть постоянной синхронизации — прогоняется один раз, после чего
можно удалить. Дальше отметки пишутся только в migbot напрямую (Sheets
из процесса записи выключается отдельным изменением бота).

Логика:
  1. Читает лист «Июль» через gspread (те же credentials, что использует
     sheets.py табеля — GOOGLE_CREDENTIALS).
  2. Для каждого активного (по листу «Сотрудники») сотрудника, для каждого
     дня с 1 июля по сегодня — читает дневной и ночной слот отдельно.
  3. Сопоставляет ФИО с employees.full_name в migbot (точное совпадение
     без учёта регистра). Кто не сопоставился — попадает в отчёт
     "не сопоставлено", ничего не пишется для этого человека вообще
     (ни одного из его дней) — решаем позже руками, что с ними делать.
  4. Пустые ячейки — не переносятся (нет отметки — нет записи), чтобы
     не плодить AttendanceMark с бессмысленным пустым code.
  5. Пишет через psycopg2 execute_values одним батчем на человека —
     не по ячейке, чтобы не упереться в write-квоту так же, как раньше
     упирались при работе с самим Google Sheets API.

Запуск: одноразово, руками, с компьютера или через Railway run:
    python migrate_attendance_from_sheets.py
Переменные окружения нужны те же, что у сервиса migbot (DATABASE_URL)
и у сервиса табеля (GOOGLE_CREDENTIALS, доступ к таблице).
"""

import json
import os
import uuid
from datetime import date, datetime, timedelta

import gspread
import psycopg2
import psycopg2.extras
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = "1d7YqIAqWL9_cQQ7JpxqD_qV69q1NpVO3u58BzDlK73M"
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

FIRST_DATA_ROW = 4
NAME_COL = 2          # столбец B
FIRST_DAY_COL = 3     # столбец C = день 1, день/ночь парами

DATABASE_URL = os.environ["DATABASE_URL"]         # база migbot — обязателен
START_DATE = date(2026, 7, 1)                     # "с 1 июля" — по договорённости


def _sheets_client():
    raw = os.environ["GOOGLE_CREDENTIALS"]
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SHEETS_SCOPES)
    return gspread.authorize(creds)


def _read_active_employee_grid():
    """Читает лист «Сотрудники» → множество активных ФИО, и грид листа «Июль»."""
    sp = _sheets_client().open_by_key(SPREADSHEET_ID)

    emp_ws = sp.worksheet("Сотрудники")
    emp_rows = emp_ws.get_all_values()[1:]
    active_names = {
        r[1].strip() for r in emp_rows
        if len(r) >= 3 and r[1].strip() and r[2].strip() == "активен"
    }

    month_ws = sp.worksheet("Июль")
    grid = month_ws.get_all_values()
    return active_names, grid


def _iter_dates(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def _extract_marks(active_names: set, grid: list) -> dict:
    """
    Возвращает {ФИО: [(mark_date, slot, code), ...]} только для активных,
    только для дат с START_DATE по сегодня, только непустые ячейки.
    """
    today = date.today()
    dates = list(_iter_dates(START_DATE, today))
    result = {}

    for r in grid[FIRST_DATA_ROW - 1:]:
        if len(r) <= NAME_COL - 1:
            continue
        name = r[NAME_COL - 1].strip()
        if not name or name not in active_names:
            continue

        marks = []
        for i, d in enumerate(dates):
            d_idx = (FIRST_DAY_COL - 1) + i * 2
            n_idx = d_idx + 1
            dval = r[d_idx].strip() if len(r) > d_idx else ""
            nval = r[n_idx].strip() if len(r) > n_idx else ""
            if dval:
                marks.append((d, "day", dval))
            if nval:
                marks.append((d, "night", nval))
        if marks:
            result[name] = marks

    return result


def _fetch_employee_ids(conn, names: list) -> dict:
    """{full_name (как в migbot): employee_id} для точных совпадений (без учёта регистра)."""
    with conn.cursor() as cur:
        cur.execute(
            "SELECT id, full_name FROM employees WHERE lower(full_name) = ANY(%s)",
            ([n.lower() for n in names],),
        )
        rows = cur.fetchall()
    # ключ — нижний регистр входного имени, чтобы сматчить обратно на исходное ФИО из Sheets
    return {full_name.strip().lower(): (emp_id, full_name) for emp_id, full_name in rows}


def _name_key(name: str) -> str:
    """Фамилия+Имя (первые 2 токена) в нижнем регистре — для fuzzy-сопоставления,
    когда отчество записано по-разному в двух системах (см. find_fuzzy_matches
    в sheets.py табеля — та же самая проблема, тот же приём)."""
    parts = name.split()
    return " ".join(parts[:2]).strip().lower() if len(parts) >= 2 else name.strip().lower()


def _fetch_all_employees(conn) -> list:
    with conn.cursor() as cur:
        cur.execute("SELECT id, full_name FROM employees")
        return cur.fetchall()


def _fuzzy_match(name: str, all_employees: list) -> tuple | None:
    """Ищет среди ВСЕХ сотрудников migbot совпадение по Фамилия+Имя.
    Возвращает (employee_id, migbot_full_name) или None. Если совпадений
    больше одного — не угадываем, возвращаем None (пусть остаётся
    в несопоставленных, разберётся человек)."""
    key = _name_key(name)
    matches = [(emp_id, full_name) for emp_id, full_name in all_employees
               if _name_key(full_name) == key]
    if len(matches) == 1:
        return matches[0]
    return None


def migrate() -> dict:
    active_names, grid = _read_active_employee_grid()
    marks_by_name = _extract_marks(active_names, grid)

    conn = psycopg2.connect(DATABASE_URL)
    conn.autocommit = False

    report = {"migrated": [], "unmatched": [], "fuzzy_matched": [], "rows_written": 0}
    try:
        id_map = _fetch_employee_ids(conn, list(marks_by_name.keys()))
        all_employees = None  # подтягиваем лениво, только если реально понадобится fuzzy

        for name, marks in marks_by_name.items():
            match = id_map.get(name.strip().lower())
            if match is None:
                if all_employees is None:
                    all_employees = _fetch_all_employees(conn)
                match = _fuzzy_match(name, all_employees)
                if match is not None:
                    report["fuzzy_matched"].append((name, match[1]))
            if match is None:
                report["unmatched"].append(name)
                continue
            employee_id, migbot_name = match

            values = [
                (str(uuid.uuid4()), employee_id, migbot_name, d, slot, code, "migration_script")
                for (d, slot, code) in marks
            ]
            with conn.cursor() as cur:
                psycopg2.extras.execute_values(
                    cur,
                    """
                    INSERT INTO attendance_marks
                        (id, employee_id, employee_name_snap, mark_date, slot, code,
                         created_by, created_at, updated_at)
                    VALUES %s
                    ON CONFLICT (employee_id, mark_date, slot)
                    DO UPDATE SET code = EXCLUDED.code, updated_at = now()
                    """,
                    values,
                    template="(%s, %s, %s, %s, %s, %s, %s, now(), now())",
                )
            report["migrated"].append((name, len(marks)))
            report["rows_written"] += len(marks)

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    return report


if __name__ == "__main__":
    result = migrate()
    print(f"Перенесено сотрудников: {len(result['migrated'])}")
    print(f"Записей отметок: {result['rows_written']}")
    if result["fuzzy_matched"]:
        print(f"\nНечёткое совпадение по Фамилия+Имя (проверь сам, что это те же люди) ({len(result['fuzzy_matched'])}):")
        for sheets_name, migbot_name in result["fuzzy_matched"]:
            print(f"  • Sheets: {sheets_name}  →  migbot: {migbot_name}")
    if result["unmatched"]:
        print(f"\nНе сопоставлено ({len(result['unmatched'])}):")
        for n in result["unmatched"]:
            print(f"  • {n}")
