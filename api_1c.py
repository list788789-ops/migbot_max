"""
Эндпоинт для интеграции с 1С:Предприятие.

Отдаёт справочник сотрудников из migbot по GET /api/1c/employees.
1С — HTTP-клиент (плановое задание раз в сутки), migbot — сервер.
Авторизация: Bearer-токен из переменной окружения ONEC_API_TOKEN.

Инкремент (вариант B без ack): сервер ведёт журнал выгрузки
(onec_export_log) — хеш последнего отданного профиля по каждому сотруднику.
При запросе отдаём только тех, у кого хеш изменился или его ещё нет,
и тут же фиксируем новые хеши (без отдельного ack-вызова).

Только синхронизация данных; кадровые документы 1С создаёт кадровик вручную.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
from datetime import date

from fastapi import Depends, FastAPI, Header, HTTPException
from sqlalchemy import create_engine, func, select
from sqlalchemy.dialects.postgresql import insert as pg_insert
from sqlalchemy.orm import Session, sessionmaker

from models import Employee, OnecExportLog

# --- Отдельная сессия БД (чтобы не было циклического импорта из webforms) ---
_DATABASE_URL = os.environ["DATABASE_URL"]
if _DATABASE_URL.startswith("postgres://"):
    _DATABASE_URL = _DATABASE_URL.replace("postgres://", "postgresql://", 1)

_engine = create_engine(_DATABASE_URL, pool_pre_ping=True)
_SessionLocal = sessionmaker(bind=_engine)


def _get_db():
    db = _SessionLocal()
    try:
        yield db
    finally:
        db.close()


def _check_token(authorization: str | None = Header(default=None)) -> None:
    expected = os.environ.get("ONEC_API_TOKEN", "")
    if not expected:
        raise HTTPException(status_code=503, detail="1C API token is not configured")
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing bearer token")
    token = authorization[len("Bearer "):].strip()
    if not hmac.compare_digest(token, expected):
        raise HTTPException(status_code=401, detail="Invalid token")


def _iso(d: date | None) -> str | None:
    return d.isoformat() if d else None


def _serialize(e: Employee) -> dict:
    tab = e.tab_number
    return {
        "id_migbot": e.id,
        "full_name": e.full_name,
        "citizenship": e.citizenship,
        "category": e.category.name.lower() if e.category else None,
        "is_rf": e.is_rf,
        "entry_date": _iso(e.entry_date),
        "contract_date": _iso(e.contract_date),
        "contract_end_date": _iso(e.contract_end_date),
        "is_active": e.contract_end_date is None,
        "birth_date": _iso(e.birth_date),
        "doc_type": e.doc_type,
        "passport_series": e.passport_series,
        "passport_number": e.passport_number,
        "iin": e.iin,
        "snils": e.snils,
        "snils_procedure": e.snils_procedure,
        "snils_appointment_date": _iso(e.snils_appointment_date),
        "position": e.position,
        "subdivision": e.subdivision,
        "tab_number": tab,
        "contract_number": f"БК-ПСМ-{tab}" if tab else None,
        "phone": e.phone,
        "registration_status": (
            e.registration_status.name.lower() if e.registration_status else None
        ),
    }


def _payload_hash(payload: dict) -> str:
    canonical = json.dumps(payload, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def _compute_delta(db: Session):
    """Возвращает (changed_payloads, rows_to_persist) — только новые/изменившиеся."""
    employees = db.scalars(select(Employee).order_by(Employee.full_name)).all()
    existing = dict(
        db.execute(
            select(OnecExportLog.employee_id, OnecExportLog.content_hash)
        ).all()
    )
    changed = []
    rows = []
    for e in employees:
        payload = _serialize(e)
        h = _payload_hash(payload)
        if existing.get(e.id) != h:
            changed.append(payload)
            rows.append({"employee_id": e.id, "content_hash": h})
    return changed, rows


def _record_exported(db: Session, rows: list[dict]) -> None:
    """Фиксирует хеши отданных профилей (upsert по employee_id)."""
    if not rows:
        return
    stmt = pg_insert(OnecExportLog).values(rows)
    stmt = stmt.on_conflict_do_update(
        index_elements=["employee_id"],
        set_={"content_hash": stmt.excluded.content_hash, "exported_at": func.now()},
    )
    db.execute(stmt)


def register_1c_routes(app: FastAPI) -> None:
    """Подключает роуты интеграции с 1С и гарантирует наличие таблицы журнала."""

    # Идемпотентно создаём таблицу журнала, если её ещё нет (новая таблица).
    OnecExportLog.__table__.create(bind=_engine, checkfirst=True)

    @app.get("/api/1c/employees")
    def list_employees_for_1c(
        _: None = Depends(_check_token),
        db: Session = Depends(_get_db),
    ):
        changed, rows = _compute_delta(db)
        _record_exported(db, rows)
        db.commit()
        return {"count": len(changed), "employees": changed}
