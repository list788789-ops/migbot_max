"""
webforms.py — веб-интерфейс для кадровика: рабочий стол с задачами и ввод даты въезда.

Отдельный сервис в ТОМ ЖЕ Railway-проекте, что и bot.py: тот же репозиторий, тот же
DATABASE_URL, те же модели из models.py. Деплой — обычный git push, Railway подхватит
Start Command именно этого сервиса (см. инструкцию по деплою внизу файла в комментарии).

ВАЖНО (не пропусти): функция entry_date_submit() ниже пишет entry_date в БД, но НЕ
создаёт obligations (регистрация/медосмотр). Эта логика уже есть в bot.py — там, где
проверяется `employment_status != EmploymentStatus.ACTIVE` перед созданием обязательств.
Нужно скопировать ту функцию сюда (или вынести в общий модуль, например obligations.py,
и импортировать из обоих файлов) и раскомментировать вызов в месте, помеченном TODO.
Без этого разница между вводом даты через бота и через веб-форму — молчаливый баг.
"""

from __future__ import annotations

import os
from datetime import date, timedelta

from fastapi import Depends, FastAPI, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session, sessionmaker
from starlette.middleware.sessions import SessionMiddleware

from models import Employee, EmploymentStatus, Obligation, ObligationStatus

# --- База данных: та же Postgres, что у бота -------------------------------

DATABASE_URL = os.environ["DATABASE_URL"]
# Railway иногда отдаёт "postgres://", SQLAlchemy 2.x требует "postgresql://"
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine)


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# --- Авторизация: один логин/пароль на кадровика, через переменные окружения -

WEBFORMS_USER = os.environ.get("WEBFORMS_USER", "kadrovik")
WEBFORMS_PASSWORD = os.environ["WEBFORMS_PASSWORD"]
SECRET_KEY = os.environ["WEBFORMS_SECRET_KEY"]

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=SECRET_KEY, session_cookie="migbot_session")


def _logged_in(request: Request) -> bool:
    return bool(request.session.get("logged_in"))


# --- Простая HTML-обёртка без отдельных файлов шаблонов ---------------------

PAGE_HEAD = """<!doctype html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
body{{font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:720px;margin:0 auto;
padding:16px;background:#f5f5f7;color:#1c1c1e}}
h1{{font-size:20px}} h2{{font-size:16px;margin-top:28px;color:#444}}
.card{{background:#fff;border-radius:12px;padding:14px;margin-bottom:10px;
box-shadow:0 1px 3px rgba(0,0,0,.08)}}
a.btn,button{{display:inline-block;background:#0a7cff;color:#fff;text-decoration:none;
padding:10px 16px;border-radius:8px;border:none;font-size:15px}}
input[type=date],input[type=text],input[type=password]{{width:100%;padding:10px;
font-size:16px;border:1px solid #ccc;border-radius:8px;margin:6px 0 12px 0;box-sizing:border-box}}
.badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:12px;color:#fff}}
.badge.red{{background:#e6473a}} .badge.orange{{background:#e69a3a}}
.muted{{color:#777;font-size:13px}}
nav{{margin-bottom:16px}}
nav a{{color:#0a7cff;text-decoration:none;margin-right:14px;font-size:14px}}
</style></head><body>
"""
PAGE_FOOT = "</body></html>"
NAV = '<nav><a href="/">Рабочий стол</a><a href="/entry_date">Дата въезда</a><a href="/logout">Выйти</a></nav>'


@app.get("/login", response_class=HTMLResponse)
def login_form():
    return PAGE_HEAD.format(title="Вход") + """
<h1>Кадровик — вход</h1>
<form method="post" action="/login">
<input type="text" name="username" placeholder="Логин" required>
<input type="password" name="password" placeholder="Пароль" required>
<button type="submit">Войти</button>
</form>""" + PAGE_FOOT


@app.post("/login")
def login_submit(request: Request, username: str = Form(...), password: str = Form(...)):
    if username == WEBFORMS_USER and password == WEBFORMS_PASSWORD:
        request.session["logged_in"] = True
        return RedirectResponse("/", status_code=303)
    return HTMLResponse(
        PAGE_HEAD.format(title="Вход")
        + '<h1>Неверный логин или пароль</h1><a class="btn" href="/login">Назад</a>'
        + PAGE_FOOT,
        status_code=401,
    )


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    today = date.today()
    soon = today + timedelta(days=7)

    no_entry_date = db.scalars(
        select(Employee)
        .where(Employee.employment_status == EmploymentStatus.ACTIVE)
        .where(Employee.entry_date.is_(None))
        .order_by(Employee.full_name)
    ).all()

    no_consent = db.scalars(
        select(Employee)
        .where(Employee.employment_status == EmploymentStatus.ACTIVE)
        .where(~Employee.consents.any())
        .order_by(Employee.full_name)
    ).all()

    overdue = db.scalars(
        select(Obligation)
        .where(Obligation.status == ObligationStatus.PENDING)
        .where(Obligation.deadline_date < today)
        .order_by(Obligation.deadline_date)
    ).all()

    due_soon = db.scalars(
        select(Obligation)
        .where(Obligation.status == ObligationStatus.PENDING)
        .where(Obligation.deadline_date >= today)
        .where(Obligation.deadline_date <= soon)
        .order_by(Obligation.deadline_date)
    ).all()

    def obl_row(o: Obligation) -> str:
        emp_name = o.employee.full_name if o.employee else "?"
        badge = "red" if o.deadline_date < today else "orange"
        return (
            f'<div class="card">{emp_name} — {o.type.value}<br>'
            f'<span class="badge {badge}">{o.deadline_date.isoformat()}</span></div>'
        )

    html = (
        PAGE_HEAD.format(title="Рабочий стол")
        + NAV
        + f"""
<h1>Задачи</h1>

<h2>Просрочено ({len(overdue)})</h2>
{''.join(obl_row(o) for o in overdue) or '<p class="muted">Нет просроченных.</p>'}

<h2>Дедлайн в ближайшие 7 дней ({len(due_soon)})</h2>
{''.join(obl_row(o) for o in due_soon) or '<p class="muted">Нет.</p>'}

<h2>Без даты въезда ({len(no_entry_date)})</h2>
{''.join(f'<div class="card">{e.full_name} <a class="btn" href="/entry_date/{e.id}">Указать дату</a></div>' for e in no_entry_date) or '<p class="muted">Все указаны.</p>'}

<h2>Без подтверждённого согласия ({len(no_consent)})</h2>
{''.join(f'<div class="card">{e.full_name}</div>' for e in no_consent) or '<p class="muted">У всех есть согласие.</p>'}
"""
        + PAGE_FOOT
    )
    return html


@app.get("/entry_date", response_class=HTMLResponse)
def entry_date_list(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    employees = db.scalars(
        select(Employee)
        .where(Employee.employment_status == EmploymentStatus.ACTIVE)
        .where(Employee.entry_date.is_(None))
        .order_by(Employee.full_name)
    ).all()

    rows = "".join(
        f'<div class="card">{e.full_name}<br>'
        f'<a class="btn" href="/entry_date/{e.id}">Указать дату</a></div>'
        for e in employees
    ) or '<p class="muted">Все активные сотрудники уже с датой въезда.</p>'

    return PAGE_HEAD.format(title="Дата въезда") + NAV + f"<h1>Кому нужна дата въезда</h1>{rows}" + PAGE_FOOT


@app.get("/entry_date/{employee_id}", response_class=HTMLResponse)
def entry_date_form(employee_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    return (
        PAGE_HEAD.format(title="Дата въезда")
        + NAV
        + f"""
<h1>{emp.full_name}</h1>
<form method="post" action="/entry_date/{emp.id}">
<label>Дата въезда</label>
<input type="date" name="entry_date" required max="{date.today().isoformat()}">
<button type="submit">Сохранить</button>
</form>"""
        + PAGE_FOOT
    )


@app.post("/entry_date/{employee_id}")
def entry_date_submit(
    employee_id: str,
    request: Request,
    entry_date: date = Form(...),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    emp.entry_date = entry_date
    db.commit()

    # TODO: вызвать ту же функцию создания obligations, что использует bot.py
    # (там, где `if employee.employment_status != EmploymentStatus.ACTIVE: return`).
    # Проще всего — вынести эту функцию из bot.py в отдельный модуль (например
    # obligations.py) и импортировать её и здесь, и в bot.py, чтобы логика не жила
    # в двух местах и не могла разойтись.
    #
    # from obligations import create_obligations_for_employee
    # create_obligations_for_employee(db, emp)

    return RedirectResponse("/entry_date", status_code=303)


# --- Деплой на Railway (кратко) ---------------------------------------------
# 1. Файл лежит в том же репозитории, что bot.py и models.py — просто закоммить
#    через GitHub web (Add file → Create new file → webforms.py).
# 2. В requirements.txt добавить строки:
#      fastapi
#      uvicorn[standard]
#      itsdangerous
#      python-multipart
# 3. В Railway: New Service → Deploy from GitHub repo (тот же репозиторий),
#    в Settings → Deploy → Start Command указать:
#      uvicorn webforms:app --host 0.0.0.0 --port $PORT
# 4. В Variables этого нового сервиса добавить:
#      DATABASE_URL          — Reference на существующий Postgres-плагин
#      WEBFORMS_USER          — логин кадровика
#      WEBFORMS_PASSWORD      — пароль кадровика
#      WEBFORMS_SECRET_KEY    — любая случайная длинная строка (для сессий)
# 5. Railway выдаст публичный URL сервиса (Settings → Networking → Generate Domain).
