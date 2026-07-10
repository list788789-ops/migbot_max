"""
РАЗОВЫЙ скрипт (2026-07): для сотрудников, у которых последний код в
attendance_marks = 'МЖ' (перенесено разовым скриптом migrate_attendance_from_sheets.py
из старого Google Sheets, БЕЗ прохождения через tabel.set_rotation), но записи в
rotation_returns нет вообще — создаёт заглушку с expected_return_date=NULL.

Дальше это уже подхватывают ПОСТОЯННЫЕ механизмы (не удалять при удалении этого
файла):
  - tabel.get_pending_clarification_rotations — список для кадровика
    ("❓ Уточнить дату возврата (МЖ)" в Требует внимания);
  - "🧹 Действия с сотрудником" в боте — у таких появляется кнопка
    "✈️ Уточнить дату возврата" (прораб вводит реальную дату и тип отбытия).

Запуск: один раз, вручную, с компьютера или через Railway run:
    python fix_orphaned_rotation.py
После успешного прогона — ЭТОТ ФАЙЛ можно и нужно удалить из репозитория.
"""

import os

from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from models import AttendanceMark, Employee, RotationReturn

DATABASE_URL = os.environ["DATABASE_URL"]


def fix_orphaned_rotation() -> list[str]:
    engine = create_engine(DATABASE_URL)
    fixed_names = []

    with Session(engine) as session:
        # Последний известный код по каждому сотруднику (дневной слот) —
        # перебор по возрастанию даты, в конце словаря остаётся последний.
        last_marks = {}
        for m in (
            session.query(AttendanceMark)
            .filter(AttendanceMark.slot == "day")
            .order_by(AttendanceMark.mark_date)
            .all()
        ):
            last_marks[m.employee_id] = m.code

        existing_rr_ids = {row[0] for row in session.query(RotationReturn.employee_id).all()}

        to_fix_ids = [
            employee_id for employee_id, code in last_marks.items()
            if code == "МЖ" and employee_id not in existing_rr_ids
        ]

        for employee_id in to_fix_ids:
            session.add(RotationReturn(
                employee_id=employee_id,
                expected_return_date=None,
                departure_type=None,
                flagged=False,
            ))

        session.commit()

        for employee_id in to_fix_ids:
            emp = session.get(Employee, employee_id)
            if emp is not None:
                fixed_names.append(emp.full_name)

    return fixed_names


if __name__ == "__main__":
    names = fix_orphaned_rotation()
    print(f"Создано заглушек: {len(names)}")
    for n in names:
        print(f"  • {n}")
