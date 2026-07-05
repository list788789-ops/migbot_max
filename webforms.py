"""
webforms.py — веб-интерфейс для кадровика: рабочий стол, карточка сотрудника, медкомиссия.

Отдельный сервис в ТОМ ЖЕ Railway-проекте, что и bot.py: тот же репозиторий, тот же
DATABASE_URL, те же модели из models.py. Деплой — обычный git push, Railway подхватит
Start Command именно этого сервиса (см. инструкцию по деплою внизу файла в комментарии).

Реальный гейт создания obligations — не employment_status (это поле в bot.py не проверяется
при создании obligations), а employee.consent_status == ConsentStatus.CONFIRMED. Логика
create_obligations_for_employee вынесена в отдельный модуль obligations.py и импортируется
и сюда, и в bot.py — см. obligations.py, там же инструкция по правке bot.py.

2026-07: консолидация. Раньше дата въезда, место въезда, дата договора, согласие и место
пребывания были отдельными разделами верхнего уровня (каждый — свой список + своя форма).
Теперь единственное место редактирования — карточка сотрудника (/employees/{id}), где все
поля собраны на одной странице несколькими независимыми мини-формами (каждая шлёт только
своё поле, чтобы случайно не затереть остальные при частичном заполнении). Раздел
"Место въезда" (entry_country) при этом обрёл первый реально работающий обработчик —
раньше на него ссылались навигация и дашборд, но сам POST/GET для него не был дописан.

Раздел "Медкомиссия" зеркалит две команды бота, а не вводит новую логику:
  - /send_medical_referral <id> — генерация направления (generate_medical_referral_docx)
  - /medical_exam_result <id> <done|failed> — отметка результата на Obligation(type=MEDICAL_EXAM)
См. _handle_medical_exam_result в bot.py — статус НЕ меняется при "failed" (в модели нет
поля для причины отказа), это сознательное решение из bot.py, а не упрощение здесь.

Место пребывания (address + address_since) — отдельное юридическое событие (см. deadlines.py,
второе правило REGISTRATION с trigger_field="address_since"). Критично: address_since
проставляется ТОЛЬКО когда это реальная смена уже известного адреса — если раньше адреса
не было (первый ввод), это часть первичной регистрации по entry_date, а не новое событие,
и address_since должен остаться NULL, иначе следующий вызов create_obligations_for_employee
создаст лишнее обязательство по несуществующей "смене".
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
from document_templates import (
    CLINIC_CHIEF_DOCTOR_NAME,
    CLINIC_CONTRACT_DATE,
    CLINIC_CONTRACT_NUMBER,
    CLINIC_NAME as REFERRAL_CLINIC_NAME,
    MEDICAL_SERVICE_TEXT,
    PAYER_NAME as REFERRAL_PAYER_NAME,
    PAYER_PHONE,
    PAYER_SIGNATORY_NAME,
    generate_medical_referral_docx,
)

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
font-family:-apple-system,Segoe UI,Roboto,sans-serif;font-weight:600;margin:2px 4px 2px 0}}
.badge.red{{background:#7a2e22}} .badge.orange{{background:#b8862b}} .badge.green{{background:#4a7a3a}}
.muted{{color:#8a7355;font-size:13px;font-family:-apple-system,Segoe UI,Roboto,sans-serif}}
nav{{margin-bottom:16px}}
nav a{{color:#7a2e22;text-decoration:none;margin-right:16px;font-size:14px;padding:6px 0;
display:inline-block;font-weight:600}}
form.inline{{display:inline}}
fieldset{{border:1px solid #e2d3ac;border-radius:6px;padding:12px;margin-bottom:14px}}
fieldset legend{{font-size:13px;color:#7a2e22;text-transform:uppercase;letter-spacing:.05em;
font-weight:600;padding:0 6px}}

@media (min-width:760px){{
  body{{max-width:1080px;padding:24px 32px}}
  header.org{{padding:20px 28px}}
  header.org .page-title{{font-size:23px}}
  section.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));
  gap:12px;align-items:start}}
  section.grid.wide{{grid-template-columns:repeat(auto-fill,minmax(340px,1fr))}}
  section.grid .card{{margin-bottom:0}}
  section.narrow{{max-width:440px;margin:0 auto}}
  section.card-form{{max-width:640px;margin:0 auto}}
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
    '<a href="/employees">Сотрудники</a>'
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


# --- Рабочий стол ------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    today = date.today()
    soon = today + timedelta(days=7)

    overdue = db.scalars(
        select(Obligation)
        .where(Obligation.is_current == True)  # noqa: E712 — SQLAlchemy требует именно так, не `is True`
        .where(Obligation.status == ObligationStatus.PENDING)
        .where(Obligation.deadline_date < today)
        .order_by(Obligation.deadline_date)
    ).all()

    due_soon = db.scalars(
        select(Obligation)
        .where(Obligation.is_current == True)  # noqa: E712
        .where(Obligation.status == ObligationStatus.PENDING)
        .where(Obligation.deadline_date >= today)
        .where(Obligation.deadline_date <= soon)
        .order_by(Obligation.deadline_date)
    ).all()

    OBLIGATION_LABELS = {
        ObligationType.REGISTRATION: "постановка на учёт",
        ObligationType.CONTRACT_NOTICE: "уведомление о договоре",
        ObligationType.CONTRACT_TERMINATION_NOTICE: "уведомление о расторжении",
        ObligationType.MEDICAL_EXAM: "медосмотр",
        ObligationType.EFS1_REPORT: "ЕФС-1",
        ObligationType.REGISTRATION_RENEWAL: "продление регистрации",
        ObligationType.PATENT_PAYMENT: "оплата патента",
    }

    def obl_row(o: Obligation) -> str:
        emp_name = o.employee.full_name if o.employee else "?"
        badge = "red" if o.deadline_date < today else "orange"
        # medical_exam решается в разделе Медкомиссия (направление/результат), остальные типы
        # (registration, contract_notice, efs1_report) — правкой соответствующей даты в карточке
        # сотрудника, откуда их дедлайн и считается.
        action_url = "/medical" if o.type == ObligationType.MEDICAL_EXAM else f"/employees/{o.employee_id}"
        action_label = "Открыть медкомиссию" if o.type == ObligationType.MEDICAL_EXAM else "Открыть карточку"
        type_label = OBLIGATION_LABELS.get(o.type, o.type.value)
        return (
            f'<div class="card">{emp_name} — {type_label}<br>'
            f'<span class="badge {badge}">{o.deadline_date.isoformat()}</span><br>'
            f'<a class="btn" href="{action_url}">{action_label}</a></div>'
        )

    # Сотрудники, у которых не хватает хотя бы одного из ключевых полей карточки —
    # единый список вместо четырёх отдельных разделов, все ведут в /employees/{id}.
    all_employees = db.scalars(select(Employee).order_by(Employee.full_name)).all()

    def missing_badges(e: Employee) -> list[str]:
        badges = []
        if e.entry_date is None:
            badges.append('<span class="badge red">нет даты въезда</span>')
        if e.entry_country is None:
            badges.append('<span class="badge orange">нет места въезда</span>')
        if e.contract_date is None:
            badges.append('<span class="badge orange">нет даты договора</span>')
        if e.consent_status == ConsentStatus.DRAFT:
            badges.append('<span class="badge red">нет согласия</span>')
        return badges

    needs_attention = [(e, missing_badges(e)) for e in all_employees]
    needs_attention = [(e, b) for e, b in needs_attention if b]

    # Медобязательства, ожидающие направления: PENDING medical_exam без связанного Referral.
    referred_obligation_ids = {r.obligation_id for r in db.scalars(select(Referral)).all()}
    need_referral = [
        o for o in db.scalars(
            select(Obligation)
            .where(Obligation.is_current == True)  # noqa: E712
            .where(Obligation.type == ObligationType.MEDICAL_EXAM)
            .where(Obligation.status == ObligationStatus.PENDING)
        ).all()
        if o.id not in referred_obligation_ids
    ]
    awaiting_result = db.scalars(
        select(Referral).where(Referral.exam_status == ExamStatus.REFERRED)
    ).all()

    def attention_row(e: Employee, badges: list[str]) -> str:
        return (
            f'<div class="card">{e.full_name}<br>'
            f'{"".join(badges)}<br>'
            f'<a class="btn" href="/employees/{e.id}">Открыть карточку</a></div>'
        )

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
<h2>Требуют внимания в карточке ({len(needs_attention)})</h2>
{''.join(attention_row(e, b) for e, b in needs_attention) or '<p class="muted">У всех сотрудников заполнены ключевые поля.</p>'}
</section>

<section>
<h2>Медкомиссия</h2>
<div class="card">Нужно направление: {len(need_referral)}<br>
Ждут результата: {len(awaiting_result)}<br>
<a class="btn" href="/medical">Открыть раздел</a></div>
</section>
"""
    return _render("Рабочий стол", body)


# --- Сотрудники: список + единая карточка ------------------------------------

@app.get("/employees", response_class=HTMLResponse)
def employees_list(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    employees = db.scalars(select(Employee).order_by(Employee.full_name)).all()

    def status_badge(e: Employee) -> str:
        return (
            '<span class="badge green">согласие ✓</span>'
            if e.consent_status == ConsentStatus.CONFIRMED
            else '<span class="badge red">без согласия</span>'
        )

    rows = "".join(
        f'<div class="card">{e.full_name}<br>{status_badge(e)}<br>'
        f'<a class="btn" href="/employees/{e.id}">Открыть карточку</a></div>'
        for e in employees
    ) or '<p class="muted">Сотрудников в базе нет.</p>'

    return _render("Сотрудники", f"<h1>Сотрудники ({len(employees)})</h1><section class=\"grid\">{rows}</section>")


@app.get("/employees/{employee_id}", response_class=HTMLResponse)
def employee_card(employee_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    today_s = date.today().isoformat()

    if emp.consent_status == ConsentStatus.CONFIRMED:
        consent_block = '<p><span class="badge green">Согласие подтверждено</span></p>'
    else:
        consent_block = f"""
<p class="muted">Подтвердить согласие кнопкой? Это тестовый способ — юридически слабее,
чем сканированная подпись (ст.9 152-ФЗ требует осознанного согласия, клик без верификации
личности это не подтверждает).</p>
<form method="post" action="/employees/{emp.id}/consent_confirm">
<button type="submit">✅ Подтвердить (кнопкой, тест)</button>
</form>"""

    body = f"""
<h1>{emp.full_name}</h1>
<section class="card-form">

<fieldset>
<legend>Дата въезда</legend>
<form method="post" action="/employees/{emp.id}/entry_date">
<input type="date" name="entry_date" required max="{today_s}"
value="{emp.entry_date.isoformat() if emp.entry_date else ''}">
<button type="submit">Сохранить</button>
</form>
</fieldset>

<fieldset>
<legend>Место въезда</legend>
<form method="post" action="/employees/{emp.id}/entry_country">
<input type="text" name="entry_country" required value="{emp.entry_country or ''}">
<button type="submit">Сохранить</button>
</form>
</fieldset>

<fieldset>
<legend>Место пребывания</legend>
<p class="muted">Текущий адрес: {emp.address or "не указан"}</p>
<form method="post" action="/employees/{emp.id}/address">
<label>Адрес</label>
<input type="text" name="address" required value="{emp.address or ''}">
<label>Дата, с которой действует этот адрес</label>
<input type="date" name="address_since" required max="{today_s}" value="{today_s}">
<button type="submit">Сохранить</button>
</form>
<p class="muted">Если это первый адрес для сотрудника — дата не создаст новое обязательство
по регистрации (оно уже создано по дате въезда). Новое обязательство появляется только
когда адрес реально МЕНЯЕТСЯ.</p>
</fieldset>

<fieldset>
<legend>Дата договора</legend>
<form method="post" action="/employees/{emp.id}/contract_date">
<input type="date" name="contract_date" required max="{today_s}"
value="{emp.contract_date.isoformat() if emp.contract_date else ''}">
<button type="submit">Сохранить</button>
</form>
</fieldset>

<fieldset>
<legend>Согласие на обработку ПД</legend>
{consent_block}
</fieldset>

<a class="btn secondary" href="/employees">← Ко всем сотрудникам</a>
</section>"""
    return _render(emp.full_name, body)


@app.post("/employees/{employee_id}/entry_date")
def employee_entry_date_submit(
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

    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/employees/{employee_id}/entry_country")
def employee_entry_country_submit(
    employee_id: str,
    request: Request,
    entry_country: str = Form(...),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    # Справочное поле — не участвует в deadlines.py (DEADLINE_RULES), obligations не пересчитываются.
    emp.entry_country = entry_country.strip()
    db.commit()

    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/employees/{employee_id}/address")
def employee_address_submit(
    employee_id: str,
    request: Request,
    address: str = Form(...),
    address_since: date = Form(...),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    previous_address = (emp.address or "").strip()
    new_address = address.strip()
    is_real_change = previous_address != "" and previous_address != new_address

    emp.address = new_address
    if is_real_change:
        emp.address_since = address_since
    # Если previous_address был пуст — это первый ввод адреса, address_since НЕ трогаем
    # (остаётся тем, что было — обычно None), чтобы не создать ложное обязательство
    # по несуществующей смене места пребывания.

    db.commit()
    db.refresh(emp)

    if is_real_change and emp.consent_status == ConsentStatus.CONFIRMED:
        create_obligations_for_employee(db, emp)

    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/employees/{employee_id}/contract_date")
def employee_contract_date_submit(
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

    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/employees/{employee_id}/consent_confirm")
def employee_consent_confirm_submit(employee_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    # Зеркалит _execute_consent_confirm_by_button в bot.py: тот же метод ConsentMethod.BOT_BUTTON,
    # тот же дисклеймер про юридическую слабость такого способа по сравнению со сканом (ст.9 152-ФЗ).
    # proof здесь — логин кадровика, а не user_id из MAX (кадровик авторизован логином/паролем,
    # не MAX-аккаунтом), чтобы след в аудите оставался осмысленным.
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

    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


# --- Медкомиссия ---------------------------------------------------------
# Зеркалит две команды бота: /send_medical_referral и /medical_exam_result.
# Не тронуто консолидацией — отдельный раздел, не привязан к общей карточке сотрудника,
# потому что список "кому нужно направление" естественно общий на всех, а не per-employee.

@app.get("/medical", response_class=HTMLResponse)
def medical_list(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    referred_obligation_ids = {r.obligation_id for r in db.scalars(select(Referral)).all()}
    need_referral = [
        o for o in db.scalars(
            select(Obligation)
            .where(Obligation.is_current == True)  # noqa: E712
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

<section class="grid wide">
<h2>Нужно направление ({len(need_referral)})</h2>
{''.join(referral_row(o) for o in need_referral) or '<p class="muted">Все, у кого активно обязательство, уже направлены.</p>'}
</section>

<section class="grid wide">
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
    except ValueError as e:
        # ValueError — сигнал от _require_fields в document_templates.py о конкретных
        # незаполненных полях, а не непредвиденная ошибка. Показываем текст как есть.
        raise HTTPException(400, str(e))
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
