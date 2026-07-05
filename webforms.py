"""
webforms.py — веб-интерфейс для кадровика: рабочий стол с задачами, дата въезда, медкомиссия.

Отдельный сервис в ТОМ ЖЕ Railway-проекте, что и bot.py: тот же репозиторий, тот же
DATABASE_URL, те же модели из models.py. Деплой — обычный git push, Railway подхватит
Start Command именно этого сервиса (см. инструкцию по деплою внизу файла в комментарии).

Реальный гейт создания obligations — не employment_status (это поле в bot.py не проверяется
при создании obligations), а employee.consent_status == ConsentStatus.CONFIRMED. Логика
create_obligations_for_employee вынесена в отдельный модуль obligations.py и импортируется
и сюда, и в bot.py — см. obligations.py, там же инструкция по правке bot.py.

Раздел "Медкомиссия" зеркалит две команды бота, а не вводит новую логику:
  - /send_medical_referral <id> — генерация направления (generate_medical_referral_docx)
  - /medical_exam_result <id> <done|failed> — отметка результата на Obligation(type=MEDICAL_EXAM)
См. _handle_medical_exam_result в bot.py — статус НЕ меняется при "failed" (в модели нет
поля для причины отказа), это сознательное решение из bot.py, а не упрощение здесь.
"""

from __future__ import annotations

import os
from datetime import date, datetime, timedelta, timezone

from fastapi import Depends, FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session, sessionmaker
from starlette.middleware.sessions import SessionMiddleware

from models import (
    Consent,
    ConsentMethod,
    ConsentStatus,
    Employee,
    ExamStatus,
    Obligation,
    ObligationStatus,
    ObligationType,
    Referral,
)
from obligations import create_obligations_for_employee
from document_templates import generate_medical_referral_docx

MSK = timezone(timedelta(hours=3))  # то же смещение, что в bot.py — для единообразия timestamp'ов proof

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
ORG_NAME = os.environ.get("COMPANY_NAME", "ИП Буц Сергей Юрьевич")
CLINIC_ID = os.environ.get("CLINIC_ID", "pirogova_murmansk")
CONSENT_TEXT_VERSION = os.environ.get("CONSENT_TEXT_VERSION", "v1")

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
*{{box-sizing:border-box}}
body{{font-family:Georgia,'Times New Roman',serif;max-width:720px;margin:0 auto;
padding:16px;background:#efe4c8;color:#3a2a1a;position:relative}}
body::before{{content:"";position:fixed;inset:0;pointer-events:none;z-index:-1;
background-color:#efe4c8;
background-image:repeating-linear-gradient(0deg,rgba(60,40,10,.03) 0px,rgba(60,40,10,.03) 1px,
transparent 1px,transparent 3px)}}
header.org{{background:#7a2e22;color:#fff8ef;border-radius:4px;padding:16px 18px;margin-bottom:16px;
border:1px solid #5c2118}}
header.org .org-name{{font-size:13px;opacity:.85;letter-spacing:.03em}}
header.org .page-title{{font-size:20px;font-weight:600;margin-top:2px}}
h1{{font-size:20px;color:#3a2a1a}}
h2{{font-size:14px;margin:0 0 10px 0;color:#7a2e22;text-transform:uppercase;letter-spacing:.06em;
border-bottom:1px solid #d8c69a;padding-bottom:6px}}
section{{background:#fffdf6;border-radius:6px;padding:14px;margin-bottom:14px;
border:1px solid #d8c69a;box-shadow:0 2px 6px rgba(60,40,10,.10)}}
.card{{background:#f7f0dd;border-radius:5px;padding:12px;margin-bottom:8px;
border:1px solid #e2d3ac;font-family:-apple-system,Segoe UI,Roboto,sans-serif;font-weight:600;
color:#2a1d10}}
.card:last-child{{margin-bottom:0}}
.card .muted-line{{font-family:inherit;font-weight:400;color:#8a7355}}
a.btn,button{{display:inline-block;background:#7a2e22;color:#fff8ef;text-decoration:none;
padding:14px 20px;min-height:48px;line-height:20px;border-radius:6px;border:none;font-size:16px;
font-family:Georgia,serif;letter-spacing:.02em;margin-top:8px;margin-right:8px;cursor:pointer}}
a.btn.secondary,button.secondary{{background:transparent;color:#7a2e22;border:2px solid #7a2e22}}
input[type=date],input[type=text],input[type=password]{{width:100%;padding:12px;
font-size:16px;font-family:-apple-system,Segoe UI,Roboto,sans-serif;border:1px solid #c9b48a;
border-radius:6px;margin:6px 0 12px 0;background:#fffef9;color:#2a1d10}}
.badge{{display:inline-block;padding:3px 10px;border-radius:10px;font-size:12px;color:#fff8ef;
font-family:-apple-system,Segoe UI,Roboto,sans-serif;font-weight:600}}
.badge.red{{background:#7a2e22}} .badge.orange{{background:#b8862b}}
.muted{{color:#8a7355;font-size:13px;font-family:-apple-system,Segoe UI,Roboto,sans-serif}}
nav{{margin-bottom:16px}}
nav a{{color:#7a2e22;text-decoration:none;margin-right:16px;font-size:14px;padding:6px 0;
display:inline-block;font-weight:600}}
form.inline{{display:inline}}

@media (min-width:760px){{
  body{{max-width:1080px;padding:24px 32px}}
  header.org{{padding:20px 28px}}
  header.org .page-title{{font-size:23px}}
  section.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));
  gap:12px;align-items:start}}
  section.grid .card{{margin-bottom:0}}
  section.narrow{{max-width:440px;margin:0 auto}}
  a.btn:hover,button:hover{{opacity:.85}}
  nav a:hover{{text-decoration:underline}}
}}
</style></head><body>
<header class="org">
<div class="org-name">{org_name}</div>
<div class="page-title">Миграционный учёт — {title}</div>
</header>
"""
PAGE_FOOT = "</body></html>"
NAV = (
    '<nav>'
    '<a href="/">Рабочий стол</a>'
    '<a href="/entry_date">Дата въезда</a>'
    '<a href="/contract_date">Дата договора</a>'
    '<a href="/consent">Согласия</a>'
    '<a href="/medical">Медкомиссия</a>'
    '<a href="/logout">Выйти</a>'
    '</nav>'
)


def _render(title: str, body: str) -> str:
    return PAGE_HEAD.format(title=title, org_name=ORG_NAME) + NAV + body + PAGE_FOOT


LOGIN_HEAD = """<!doctype html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Вход</title>
<style>
*{{box-sizing:border-box}}
body.login-page{{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;
background:radial-gradient(ellipse at center,#f4ecd8 0%,#e4d5b0 55%,#cdb98a 100%);
font-family:Georgia,'Times New Roman',serif;position:relative;overflow:hidden;padding:16px}}
body.login-page::before{{content:"";position:absolute;inset:0;pointer-events:none;
background-image:repeating-linear-gradient(0deg,rgba(60,40,10,.035) 0px,rgba(60,40,10,.035) 1px,
transparent 1px,transparent 3px)}}
.vintage-bg{{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;
opacity:.32;transform:rotate(-9deg);pointer-events:none}}
.vintage-bg svg{{width:320px;height:320px}}
.login-card{{position:relative;z-index:1;background:#fffdf6;border:1px solid #c9b48a;
border-radius:4px;padding:28px 24px;width:280px;box-shadow:0 8px 28px rgba(60,40,10,.28)}}
.login-card h1{{font-size:19px;margin:0 0 2px 0;color:#3a2a1a}}
.login-card .subtitle{{font-size:11px;color:#8a7355;margin:0 0 18px 0;letter-spacing:.06em;
text-transform:uppercase}}
.login-card input{{width:100%;padding:10px;font-size:16px;border:1px solid #c9b48a;border-radius:4px;
margin:6px 0 12px 0;background:#fffef9;font-family:inherit}}
.login-card button{{width:100%;background:#7a2e22;color:#fff8ef;border:none;padding:15px;
min-height:50px;border-radius:6px;font-size:16px;cursor:pointer;letter-spacing:.03em;font-family:inherit}}
.login-card a.btn{{display:inline-block;margin-top:12px;color:#7a2e22;text-decoration:underline;
font-size:13px}}
</style></head>
<body class="login-page">
<div class="vintage-bg">
<svg viewBox="0 0 400 400">
<defs>
<path id="stampCircle" d="M 200,200 m -150,0 a 150,150 0 1,1 300,0 a 150,150 0 1,1 -300,0" />
</defs>
<circle cx="200" cy="200" r="150" fill="none" stroke="#7a2e22" stroke-width="4"/>
<circle cx="200" cy="200" r="130" fill="none" stroke="#7a2e22" stroke-width="2"/>
<text font-size="19" fill="#7a2e22" letter-spacing="6">
<textPath href="#stampCircle" startOffset="1%">МИГРАЦИОННЫЙ УЧЁТ • БЕЛОКАМЕННАЯ • MURMANSK •</textPath>
</text>
<g stroke="#7a2e22" fill="none" stroke-width="3">
<line x1="200" y1="118" x2="200" y2="282"/>
<line x1="118" y1="200" x2="282" y2="200"/>
<circle cx="200" cy="200" r="42"/>
<circle cx="200" cy="200" r="6" fill="#7a2e22"/>
</g>
</svg>
</div>
"""


@app.get("/login", response_class=HTMLResponse)
def login_form():
    return LOGIN_HEAD + f"""
<div class="login-card">
<h1>Миграционный учёт</h1>
<p class="subtitle">{ORG_NAME}</p>
<form method="post" action="/login">
<input type="text" name="username" placeholder="Логин" required>
<input type="password" name="password" placeholder="Пароль" required>
<button type="submit">Войти</button>
</form>
</div>
</body></html>"""


@app.post("/login")
def login_submit(request: Request, username: str = Form(...), password: str = Form(...)):
    if username == WEBFORMS_USER and password == WEBFORMS_PASSWORD:
        request.session["logged_in"] = True
        return RedirectResponse("/", status_code=303)
    return HTMLResponse(
        LOGIN_HEAD
        + """
<div class="login-card">
<h1>Неверный логин или пароль</h1>
<a class="btn" href="/login">Назад</a>
</div>
</body></html>""",
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
        .where(Employee.entry_date.is_(None))
        .order_by(Employee.full_name)
    ).all()

    no_consent = db.scalars(
        select(Employee)
        .where(Employee.consent_status == ConsentStatus.DRAFT)
        .order_by(Employee.full_name)
    ).all()

    no_contract_date = db.scalars(
        select(Employee)
        .where(Employee.contract_date.is_(None))
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

    # Медобязательства, ожидающие направления: PENDING medical_exam без связанного Referral.
    referred_obligation_ids = {r.obligation_id for r in db.scalars(select(Referral)).all()}
    need_referral = [
        o for o in db.scalars(
            select(Obligation)
            .where(Obligation.type == ObligationType.MEDICAL_EXAM)
            .where(Obligation.status == ObligationStatus.PENDING)
        ).all()
        if o.id not in referred_obligation_ids
    ]
    awaiting_result = db.scalars(
        select(Referral).where(Referral.exam_status == ExamStatus.REFERRED)
    ).all()

    body = f"""
<h1>Задачи</h1>

<section class="grid">
<h2>Просрочено ({len(overdue)})</h2>
{''.join(obl_row(o) for o in overdue) or '<p class="muted">Нет просроченных.</p>'}
</section>

<section class="grid">
<h2>Дедлайн в ближайшие 7 дней ({len(due_soon)})</h2>
{''.join(obl_row(o) for o in due_soon) or '<p class="muted">Нет.</p>'}
</section>

<section class="grid">
<h2>Без даты въезда ({len(no_entry_date)})</h2>
{''.join(f'<div class="card">{e.full_name} <a class="btn" href="/entry_date/{e.id}">Указать дату</a></div>' for e in no_entry_date) or '<p class="muted">Все указаны.</p>'}
</section>

<section class="grid">
<h2>Без подтверждённого согласия ({len(no_consent)})</h2>
{''.join(f'<div class="card">{e.full_name} <a class="btn" href="/consent/{e.id}">Подтвердить</a></div>' for e in no_consent) or '<p class="muted">У всех есть согласие.</p>'}
</section>

<section class="grid">
<h2>Без даты договора ({len(no_contract_date)})</h2>
{''.join(f'<div class="card">{e.full_name} <a class="btn" href="/contract_date/{e.id}">Указать дату</a></div>' for e in no_contract_date) or '<p class="muted">У всех указана дата договора.</p>'}
</section>

<section>
<h2>Медкомиссия</h2>
<div class="card">Нужно направление: {len(need_referral)}<br>
Ждут результата: {len(awaiting_result)}<br>
<a class="btn" href="/medical">Открыть раздел</a></div>
</section>
"""
    return _render("Рабочий стол", body)


@app.get("/entry_date", response_class=HTMLResponse)
def entry_date_list(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    employees = db.scalars(
        select(Employee)
        .where(Employee.entry_date.is_(None))
        .order_by(Employee.full_name)
    ).all()

    rows = "".join(
        f'<div class="card">{e.full_name}<br>'
        f'<a class="btn" href="/entry_date/{e.id}">Указать дату</a></div>'
        for e in employees
    ) or '<p class="muted">У всех сотрудников уже есть дата въезда.</p>'

    return _render("Дата въезда", f"<h1>Кому нужна дата въезда</h1><section class=\"grid\">{rows}</section>")


@app.get("/entry_date/{employee_id}", response_class=HTMLResponse)
def entry_date_form(employee_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    body = f"""
<h1>{emp.full_name}</h1>
<section class="narrow">
<form method="post" action="/entry_date/{emp.id}">
<label>Дата въезда</label>
<input type="date" name="entry_date" required max="{date.today().isoformat()}">
<button type="submit">Сохранить</button>
</form>
</section>"""
    return _render("Дата въезда", body)


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
    db.refresh(emp)

    # Симметрично _apply_entry_date() в bot.py: если согласие уже подтверждено раньше
    # (маловероятно на практике, но на будущее — дозаполнение может случиться уже после
    # подтверждения), досоздаём obligations сейчас же, а не оставляем их несозданными молча.
    if emp.consent_status == ConsentStatus.CONFIRMED:
        create_obligations_for_employee(db, emp)

    return RedirectResponse("/entry_date", status_code=303)


# --- Согласия (тестовое подтверждение кнопкой) ------------------------------
# Зеркалит _deliver_consent_confirmation / _execute_consent_confirm_by_button в bot.py:
# тот же метод ConsentMethod.BOT_BUTTON, тот же дисклеймер про юридическую слабость
# такого способа по сравнению со сканом (ст.9 152-ФЗ), тот же вызов
# create_obligations_for_employee сразу после подтверждения.

@app.get("/consent", response_class=HTMLResponse)
def consent_list(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    employees = db.scalars(
        select(Employee)
        .where(Employee.consent_status == ConsentStatus.DRAFT)
        .order_by(Employee.full_name)
    ).all()

    rows = "".join(
        f'<div class="card">{e.full_name}<br>'
        f'<a class="btn" href="/consent/{e.id}">Подтвердить</a></div>'
        for e in employees
    ) or '<p class="muted">У всех сотрудников согласие подтверждено.</p>'

    return _render("Согласия", f"<h1>Ожидают согласия</h1><section class=\"grid\">{rows}</section>")


@app.get("/consent/{employee_id}", response_class=HTMLResponse)
def consent_confirm_form(employee_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    body = f"""
<h1>{emp.full_name}</h1>
<section class="narrow">
<p class="muted">Подтвердить согласие кнопкой? Это тестовый способ — юридически слабее,
чем сканированная подпись (ст.9 152-ФЗ требует осознанного согласия, клик без верификации
личности это не подтверждает).</p>
<form method="post" action="/consent/{emp.id}/confirm">
<button type="submit">✅ Подтвердить (кнопкой, тест)</button>
</form>
<a class="btn secondary" href="/consent">Отмена</a>
</section>"""
    return _render("Согласия", body)


@app.post("/consent/{employee_id}/confirm")
def consent_confirm_submit(employee_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    # proof здесь — не user_id из MAX (кадровик авторизован логином/паролем, не MAX-аккаунтом),
    # поэтому используем логин кадровика вместо user_id, чтобы след в аудите оставался осмысленным.
    consent = Consent(
        employee_id=emp.id,
        method=ConsentMethod.BOT_BUTTON,
        proof=f"button_click:webforms:{WEBFORMS_USER}:{datetime.now(MSK).isoformat()}",
        consent_text_version=CONSENT_TEXT_VERSION,
    )
    db.add(consent)

    emp.consent_status = ConsentStatus.CONFIRMED
    db.add(emp)
    db.commit()
    db.refresh(emp)

    create_obligations_for_employee(db, emp)

    return RedirectResponse("/consent", status_code=303)


# --- Дата договора -----------------------------------------------------------
# Симметрично /entry_date выше и _apply_contract_date() в bot.py.

@app.get("/contract_date", response_class=HTMLResponse)
def contract_date_list(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    employees = db.scalars(
        select(Employee)
        .where(Employee.contract_date.is_(None))
        .order_by(Employee.full_name)
    ).all()

    rows = "".join(
        f'<div class="card">{e.full_name}<br>'
        f'<a class="btn" href="/contract_date/{e.id}">Указать дату</a></div>'
        for e in employees
    ) or '<p class="muted">У всех сотрудников уже есть дата договора.</p>'

    return _render("Дата договора", f"<h1>Кому нужна дата договора</h1><section class=\"grid\">{rows}</section>")


@app.get("/contract_date/{employee_id}", response_class=HTMLResponse)
def contract_date_form(employee_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    body = f"""
<h1>{emp.full_name}</h1>
<section class="narrow">
<form method="post" action="/contract_date/{emp.id}">
<label>Дата договора</label>
<input type="date" name="contract_date" required max="{date.today().isoformat()}">
<button type="submit">Сохранить</button>
</form>
</section>"""
    return _render("Дата договора", body)


@app.post("/contract_date/{employee_id}")
def contract_date_submit(
    employee_id: str,
    request: Request,
    contract_date: date = Form(...),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    emp.contract_date = contract_date
    db.commit()
    db.refresh(emp)

    # Симметрично _apply_contract_date() в bot.py: обязательства, зависящие от даты договора
    # (contract_notice, efs1_report), досоздаются сразу, если согласие уже подтверждено.
    if emp.consent_status == ConsentStatus.CONFIRMED:
        create_obligations_for_employee(db, emp)

    return RedirectResponse("/contract_date", status_code=303)


@app.get("/medical", response_class=HTMLResponse)
def medical_list(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    referred_obligation_ids = {r.obligation_id for r in db.scalars(select(Referral)).all()}
    need_referral = [
        o for o in db.scalars(
            select(Obligation)
            .where(Obligation.type == ObligationType.MEDICAL_EXAM)
            .where(Obligation.status == ObligationStatus.PENDING)
            .order_by(Obligation.deadline_date)
        ).all()
        if o.id not in referred_obligation_ids
    ]
    awaiting_result = db.scalars(
        select(Referral).where(Referral.exam_status == ExamStatus.REFERRED)
    ).all()

    def referral_row(o: Obligation) -> str:
        emp = o.employee
        name = emp.full_name if emp else "?"
        return (
            f'<div class="card">{name}<br>'
            f'<span class="badge orange">дедлайн {o.deadline_date.isoformat()}</span>'
            f'<form class="inline" method="post" action="/medical/{o.employee_id}/refer">'
            f'<input type="hidden" name="obligation_id" value="{o.id}">'
            f'<button type="submit">Выписать направление</button>'
            f'</form></div>'
        )

    def awaiting_row(r: Referral) -> str:
        emp = r.employee if hasattr(r, "employee") else None
        name = emp.full_name if emp else db.get(Employee, r.employee_id).full_name
        return (
            f'<div class="card">{name}<br>'
            f'<span class="muted">направлен {r.referral_date.isoformat()}</span><br>'
            f'<form class="inline" method="post" action="/medical/{r.employee_id}/result">'
            f'<input type="hidden" name="result" value="done">'
            f'<button type="submit">✅ Пройдено</button></form>'
            f'<form class="inline" method="post" action="/medical/{r.employee_id}/result">'
            f'<input type="hidden" name="result" value="failed">'
            f'<button type="submit" class="secondary">❌ Не пройдено</button></form>'
            f'</div>'
        )

    body = f"""
<h1>Медкомиссия</h1>

<section>
<h2>Нужно направление ({len(need_referral)})</h2>
{''.join(referral_row(o) for o in need_referral) or '<p class="muted">Все, у кого активно обязательство, уже направлены.</p>'}
</section>

<section>
<h2>Направлены, ждут результата ({len(awaiting_result)})</h2>
{''.join(awaiting_row(r) for r in awaiting_result) or '<p class="muted">Нет ожидающих результата.</p>'}
</section>
"""
    return _render("Медкомиссия", body)


@app.post("/medical/{employee_id}/refer")
def medical_refer(
    employee_id: str,
    request: Request,
    obligation_id: str = Form(...),
    db: Session = Depends(get_db),
):
    """Выписывает направление: генерирует docx (та же генерация, что /send_medical_referral
    в bot.py) и заводит запись Referral, привязанную к обязательству MEDICAL_EXAM. bot.py
    сам Referral не создаёт — это новая часть учёта, добавленная здесь."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    obligation = db.get(Obligation, obligation_id)
    if obligation is None or obligation.employee_id != employee_id:
        raise HTTPException(404, "Обязательство не найдено или не принадлежит этому сотруднику")

    try:
        path = generate_medical_referral_docx(emp)
    except Exception:
        raise HTTPException(500, "Не удалось сгенерировать направление. Проверьте логи сервиса.")

    referral = Referral(
        employee_id=emp.id,
        obligation_id=obligation.id,
        clinic_id=CLINIC_ID,
        referral_date=date.today(),
        exam_status=ExamStatus.REFERRED,
    )
    db.add(referral)
    db.commit()

    filename = f"Направление_{emp.full_name.replace(' ', '_')}.docx"
    return FileResponse(path, filename=filename)


@app.post("/medical/{employee_id}/result")
def medical_result(
    employee_id: str,
    request: Request,
    result: str = Form(...),
    db: Session = Depends(get_db),
):
    """Симметрично /medical_exam_result в bot.py: при 'failed' статус Obligation НЕ меняется
    (в модели нет поля для причины отказа, дедлайн должен остаться активным) — см. комментарий
    в _handle_medical_exam_result в bot.py, это не упрощение, а то же самое решение."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    if result not in ("done", "failed"):
        raise HTTPException(400, "result должен быть 'done' или 'failed'")

    referral = db.scalars(
        select(Referral)
        .where(Referral.employee_id == employee_id)
        .where(Referral.exam_status == ExamStatus.REFERRED)
        .order_by(Referral.referral_date.desc())
    ).first()
    if referral is None:
        raise HTTPException(404, "Нет направления, ожидающего результата, для этого сотрудника")

    referral.exam_status = ExamStatus.COMPLETED
    referral.result_date = date.today()
    db.add(referral)

    if result == "done":
        obligation = db.get(Obligation, referral.obligation_id)
        if obligation is not None:
            obligation.status = ObligationStatus.DONE
            db.add(obligation)

    db.commit()
    return RedirectResponse("/medical", status_code=303)


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
