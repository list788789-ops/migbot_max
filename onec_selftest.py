"""Временный self-test дельты 1С-выгрузки.

Запуск: venv/bin/python onec_selftest.py
НЕ меняет данные — всё откатывается (rollback). Файл временный, удалить после проверки.
"""
import os
from pathlib import Path

os.chdir(Path(__file__).resolve().parent)
for _line in Path(".env").read_text().splitlines():
    _line = _line.strip()
    if _line.startswith("export "):
        _line = _line[len("export "):]
    if _line and not _line.startswith("#") and "=" in _line:
        _k, _v = _line.split("=", 1)
        os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))

import api_1c
from models import OnecExportLog

db = api_1c._SessionLocal()
res = {}
try:
    changed1, rows1 = api_1c._compute_delta(db)
    res["first_call_changed"] = len(changed1)
    api_1c._record_exported(db, rows1)
    db.flush()
    changed2, _ = api_1c._compute_delta(db)
    res["second_call_changed"] = len(changed2)
    api_1c._record_exported(db, rows1)
    db.flush()
    res["upsert_update_ok"] = True
    if rows1:
        db.query(OnecExportLog).filter_by(
            employee_id=rows1[0]["employee_id"]
        ).update({"content_hash": "changed"})
        db.flush()
        changed3, _ = api_1c._compute_delta(db)
        res["after_one_changed"] = len(changed3)
    print("SELFTEST_RESULT", res)
finally:
    db.rollback()
    db.close()
