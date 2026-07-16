"""
Эндпоинт для интеграции с 1С:Предприятие.

Отдаёт справочник сотрудников из migbot по GET /api/1c/employees.
1С — HTTP-клиент (плановое задание раз в сутки), migbot — сервер.
Авторизация: Bearer-токен из переменной окружения ONEC_API_TOKEN.

Первая итерация: полная выгрузка (updated_at у Employee пока нет — инкремент позже).
Только синхронизация данных; кадровые документы 1С создаёт кадровик вручную.
"""
from __future__ import annotations

import hmac
import os
from datetime import date

from fastapi import Depends, FastAPI, Header, HTTPException
from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session, sessionmaker

from models import Employee

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


def register_1c_routes(app: FastAPI) -> None:
    """Подключает роуты интеграции с 1С к переданному FastAPI-приложению."""

    @app.get("/api/1c/employees")
    def list_employees_for_1c(
        _: None = Depends(_check_token),
        db: Session = Depends(_get_db),
    ):
        employees = db.scalars(
            select(Employee).order_by(Employee.full_name)
        ).all()
        data = [_serialize(e) for e in employees]
        return {"count": len(data), "employees": data}
