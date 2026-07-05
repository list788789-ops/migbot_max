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

2026-07: два изменения ради теста потока медкомиссии:
  1. Тестовый режим TEST_ALLOW_MISSING_FIELDS (флаг живёт в document_templates.py,
     здесь только читается) — при незаполненных полях (адрес, паспорт, дата рождения)
     направление всё равно генерируется, с прочерками и явным баннером-предупреждением
     в HTML-превью и в самом тексте docx (см. document_templates._add_test_warning_paragraph).
     ВРЕМЕННО — флаг нужно выключить/удалить перед реальной работой с сотрудниками.
  2. Обработчик HTTPException переопределён так, чтобы Content-Type ошибок явно нёс
     charset=utf-8 — без этого мобильный браузер иногда угадывал кодировку JSON-ответа
     неверно и кириллица в сообщении об ошибке превращалась в кракозябры (двойное
     перекодирование на экране, не в самих данных).
"""

from __future__ import annotations

import html
import os
from datetime import date, datetime, timedelta, timezone

from fastapi import Depends, FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session, sessionmaker
from starlette.exceptions import HTTPException as StarletteHTTPException
from starlette.middleware.sessions import SessionMiddleware

from models import (
    Consent,
    ConsentMethod,
    ConsentStatus,
    DeadlineUnit,
    Employee,
    ExamStatus,
    Obligation,
    ObligationStatus,
    ObligationType,
    Referral,
    SystemFlag,
)
from obligations import create_obligations_for_employee
from deadlines import lead_days_for
from document_templates import (
    CLINIC_CHIEF_DOCTOR_NAME,
    CLINIC_CONTRACT_DATE,
    CLINIC_CONTRACT_NUMBER,
    CLINIC_NAME as REFERRAL_CLINIC_NAME,
    CLINIC_SHORT_NAME as REFERRAL_CLINIC_SHORT_NAME,
    MEDICAL_SERVICE_TEXT,
    PAYER_NAME as REFERRAL_PAYER_NAME,
    PAYER_PHONE,
    PAYER_SIGNATORY_NAME,
    TEST_ALLOW_MISSING_FIELDS,
    check_medical_referral_fields,
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


@app.exception_handler(StarletteHTTPException)
async def http_exception_handler(request: Request, exc: StarletteHTTPException):
    """Переопределяет дефолтный обработчик FastAPI ТОЛЬКО ради явного charset=utf-8
    в заголовке ответа. Starlette проставляет charset автоматически лишь для media_type,
    начинающихся с "text/" — для "application/json" charset в Content-Type отсутствует,
    и часть мобильных браузеров/веб-вью в таком случае угадывает кодировку неверно,
    из-за чего кириллица в detail превращается в кракозябры на экране (сами данные
    при этом были в порядке — ломалось только отображение)."""
    return JSONResponse(
        status_code=exc.status_code,
        content={"detail": exc.detail},
        media_type="application/json; charset=utf-8",
    )


def _logged_in(request: Request) -> bool:
    return bool(request.session.get("logged_in"))


def _obligation_status(o, today):
    """(chip_class, label, bucket). bucket: 'overdue' | 'soon' | 'ok'.
    Порог 'скоро' берётся из deadlines.lead_days_for (один источник со сроком).
    Короткоплечие (WORKING_DAY: переезд/уведомление/ЕФС-1) — born-amber: 'soon'
    с момента создания, пока не сделаны. Единая точка окраски для дашборда и списка."""
    days = (o.deadline_date - today).days
    if days < 0:
        return ("red", f"просрочено на {-days} дн", "overdue")
    cat = o.employee.category if o.employee else None
    if o.deadline_unit == DeadlineUnit.WORKING_DAY:
        soon = True
    else:
        soon = days <= lead_days_for(cat, o.type, o.deadline_unit)
    if soon:
        label = "сегодня" if days == 0 else ("завтра" if days == 1 else f"через {days} дн")
        return ("orange", label, "soon")
    return ("green", "в норме", "ok")


# Человекочитаемые названия типов обязанностей — один источник для дашборда и списка.
SAVE_FORM_JS = """
<script>
(function(){
  var f = document.getElementById('saveform');
  if(!f) return;
  f.addEventListener('submit', function(e){
    var warn = [];
    var a = f.querySelector('[name=address]');
    if(a && a.value.trim() && (a.dataset.orig||'').trim() && a.value.trim() !== (a.dataset.orig||'').trim())
      warn.push('адрес места пребывания — создаст обязательство регистрации');
    var d = f.querySelector('[name=dactyloscopy_date]');
    if(d && d.value && d.value !== (d.dataset.orig||''))
      warn.push('дата дактилоскопии — закроет обязательство как пройденное');
    if(warn.length){
      if(confirm('Эти изменения затронут обязательства сотрудника:\n\n\u2022 ' + warn.join('\n\u2022 ') + '\n\nПродолжить?')){
        f.querySelector('[name=confirmed]').value = '1';
      } else {
        e.preventDefault();
      }
    }
  });
})();
</script>
"""


OBLIGATION_LABELS = {
    ObligationType.REGISTRATION: "постановка на учёт",
    ObligationType.CONTRACT_NOTICE: "уведомление о договоре",
    ObligationType.CONTRACT_TERMINATION_NOTICE: "уведомление о расторжении",
    ObligationType.MEDICAL_EXAM: "медосмотр",
    ObligationType.EFS1_REPORT: "ЕФС-1",
    ObligationType.DACTYLOSCOPY: "дактилоскопия",
    ObligationType.REGISTRATION_RENEWAL: "продление регистрации",
    ObligationType.PATENT_PAYMENT: "оплата патента",
}


# --- Простая HTML-обёртка без отдельных файлов шаблонов ---------------------

PAGE_HEAD = """<!doctype html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
:root{{--ink:#111214;--sub:#5b626b;--line:#e6e9ee;--line-2:#eef1f4;--accent:#2f80ed;--accent-ink:#1c63c4;
--red-bg:#fdecec;--red-ink:#c0392b;--amber-bg:#fdf3e2;--amber-ink:#a5720a;--green-bg:#eaf6f0;--green-ink:#1f7a55;
--neutral-bg:#eef1f4;--neutral-ink:#55606b;--serif:Georgia,"Times New Roman",serif;
--sans:-apple-system,Segoe UI,Roboto,Arial,sans-serif}}
*{{box-sizing:border-box}}
body{{font-family:var(--sans);max-width:720px;margin:0 auto;padding:16px;background:#fff;color:var(--ink)}}
header.org{{background:#fff;border:0;border-bottom:1px solid var(--line);border-radius:0;padding:14px 4px 16px;margin-bottom:8px}}
header.org .org-name{{font-size:13px;color:var(--sub);letter-spacing:.01em}}
header.org .page-title{{font-family:var(--serif);font-size:26px;font-weight:700;letter-spacing:-.02em;margin-top:2px}}
h1{{font-family:var(--serif);font-size:24px;font-weight:700;letter-spacing:-.02em}}
h2{{font-size:12px;margin:0 0 12px;color:var(--sub);text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid var(--line-2);padding-bottom:8px;font-weight:700}}
section{{background:#fff;border:1px solid var(--line);border-radius:16px;padding:16px;margin-bottom:14px}}
.card{{background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px;margin-bottom:10px;font-weight:600;color:var(--ink)}}
.card:last-child{{margin-bottom:0}}
.card .muted-line{{font-weight:400;color:var(--sub)}}
a.btn,button{{display:inline-block;background:var(--accent);color:#fff;text-decoration:none;padding:14px 20px;min-height:48px;line-height:20px;border-radius:12px;border:none;font-size:16px;font-family:var(--sans);font-weight:600;margin:8px 8px 0 0;cursor:pointer}}
a.btn.secondary,button.secondary{{background:#fff;color:var(--accent-ink);border:1px solid var(--accent)}}
input[type=date],input[type=text],input[type=password]{{width:100%;padding:12px;font-size:16px;font-family:inherit;border:1px solid #d9dde3;border-radius:12px;margin:6px 0 12px;background:#fff;color:var(--ink)}}
input:focus{{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px #2f80ed22}}
label{{font-size:13px;color:var(--sub)}}
.badge{{display:inline-block;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:600;margin:2px 4px 2px 0}}
.badge.red{{background:var(--red-bg);color:var(--red-ink)}}
.badge.orange{{background:var(--amber-bg);color:var(--amber-ink)}}
.badge.green{{background:var(--green-bg);color:var(--green-ink)}}
.badge.neutral{{background:var(--neutral-bg);color:var(--neutral-ink)}}
.muted{{color:var(--sub);font-size:13px}}
.warning-banner{{background:var(--amber-bg);border:1px solid #f0c674;border-left:4px solid var(--amber-ink);border-radius:12px;padding:12px 14px;margin-bottom:14px;font-weight:600;color:#7a4a00}}
nav{{margin-bottom:16px;background:#fff;border:1px solid var(--line);border-radius:12px;padding:4px 8px;display:flex;gap:2px;overflow-x:auto}}
nav a{{color:var(--sub);text-decoration:none;font-size:15px;padding:10px 12px;white-space:nowrap;font-weight:600;border-radius:8px}}
nav a:hover{{color:var(--ink);background:#f5f7f9}}
form.inline{{display:inline}}
fieldset{{border:1px solid var(--line);border-radius:14px;padding:14px;margin-bottom:14px}}
fieldset legend{{font-size:12px;color:var(--accent-ink);text-transform:uppercase;letter-spacing:.04em;font-weight:700;padding:0 6px}}
@media (min-width:760px){{
  body{{max-width:1080px;padding:24px 32px}}
  header.org .page-title{{font-size:30px}}
  section.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px;align-items:start}}
  section.grid.wide{{grid-template-columns:repeat(auto-fill,minmax(340px,1fr))}}
  section.grid h2{{grid-column:1/-1}}
  section.grid .card{{margin-bottom:0}}
  section.narrow{{max-width:440px;margin:0 auto}}
  section.card-form{{max-width:640px;margin:0 auto}}
  a.btn:hover,button:hover{{opacity:.9}}
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
:root{--accent:#2f80ed;--ink:#111214;--sub:#5b626b;--serif:Georgia,"Times New Roman",serif;--sans:-apple-system,Segoe UI,Roboto,Arial,sans-serif}
*{box-sizing:border-box}
body.login-page{margin:0;min-height:100dvh;display:flex;flex-direction:column;justify-content:flex-start;background:url('/login-bg.svg') no-repeat center bottom / cover, #fff;font-family:var(--sans);color:var(--ink)}
.auth{width:100%;max-width:440px;margin:0 auto;padding:56px 24px 40px}
.auth-row{background:#fff;border:1px solid #e6e9ee;border-radius:16px;padding:14px;box-shadow:0 6px 24px rgba(20,24,30,.10)}
.auth h1{font-family:var(--serif);font-weight:700;letter-spacing:-.02em;font-size:clamp(2.25rem,6vw,3.25rem);line-height:1.03;margin:0 0 .3em}
.auth .subtitle{font-family:var(--serif);font-weight:400;color:var(--sub);font-size:clamp(1.125rem,2.5vw,1.375rem);margin:0 0 1.75rem}
.auth-row{display:flex;gap:10px;flex-wrap:wrap;align-items:stretch}
.auth input{flex:1 1 45%;min-width:130px;font-family:var(--sans);font-size:16px;padding:14px 16px;border:1px solid #b8c0cc;border-radius:12px;background:#fff;margin:0}
.auth input:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px #2f80ed22}
.auth button{flex:1 1 100%;border:0;border-radius:12px;background:var(--accent);color:#fff;font:16px/1 var(--sans);padding:15px 20px;min-height:50px;cursor:pointer}
.auth .err{color:#c0392b;font-size:14px;margin:0 0 12px}
.auth a.btn{display:inline-block;margin-top:14px;color:var(--accent);text-decoration:underline;font-size:14px}
@media(min-width:520px){ .auth button{flex:0 0 auto} }
@media(min-width:768px){
  body.login-page{justify-content:center;background-position:right bottom;background-size:min(46vw,600px) auto}
  .auth{padding:0 24px}
}
</style></head>
<body class="login-page">
"""


@app.get("/login", response_class=HTMLResponse)
def login_form():
    return LOGIN_HEAD + """
<form class="auth" method="post" action="/login" autocomplete="on">
<h1>Миграционный учёт</h1>
<p class="subtitle">Рабочее место кадровика</p>
<div class="auth-row">
<input type="text" name="username" placeholder="Логин" autocomplete="username" required>
<input type="password" name="password" placeholder="Пароль" autocomplete="current-password" required>
<button type="submit">Войти</button>
</div>
</form>
</body></html>"""


@app.post("/login")
def login_submit(request: Request, username: str = Form(...), password: str = Form(...)):
    if username == WEBFORMS_USER and password == WEBFORMS_PASSWORD:
        request.session["logged_in"] = True
        return RedirectResponse("/", status_code=303)
    return HTMLResponse(
        LOGIN_HEAD
        + """
<div class="auth">
<h1>Миграционный учёт</h1>
<p class="err">Неверный логин или пароль</p>
<a class="btn" href="/login">← Назад</a>
</div>
</body></html>""",
        status_code=401,
    )


LOGIN_BG_SVG = """<svg viewBox="0 0 430 760" preserveAspectRatio="xMidYMid slice" xmlns="http://www.w3.org/2000/svg">
<g fill="none" stroke="#c3ccd6" stroke-width="1">
<path d="M-40 300 Q215 280 470 300"/><path d="M-40 360 Q215 340 470 360"/>
<path d="M-40 420 Q215 400 470 420"/><path d="M-40 480 Q215 460 470 480"/>
<path d="M-40 540 Q215 520 470 540"/><path d="M-40 600 Q215 580 470 600"/>
<path d="M-40 660 Q215 640 470 660"/><path d="M-40 720 Q215 700 470 720"/>
<path d="M60 760 Q120 500 175 250"/><path d="M150 760 Q180 500 205 250"/>
<path d="M240 760 Q235 500 235 250"/><path d="M330 760 Q290 500 265 250"/>
<path d="M420 760 Q350 500 295 250"/></g>
<g stroke="#aab4c0" stroke-width="1"><path d="M231 300 L239 300 M231 360 L239 360 M231 420 L239 420 M231 480 L239 480 M231 600 L239 600 M231 660 L239 660 M231 720 L239 720"/></g>
<g fill="none" stroke="#b7c0cb" stroke-width="1">
<path d="M120 760 C150 660 120 600 175 545 C215 505 205 460 250 430"/>
<path d="M170 560 l-6 -4 M158 588 l-6 -4 M150 618 l-6 -4 M146 650 l-6 -4"/></g>
<g stroke="#aab4c0" fill="none" stroke-width="1"><path d="M392 300 L392 330 M392 300 L387 309 M392 300 L397 309"/></g>
<text x="388" y="292" font-family="-apple-system,system-ui,sans-serif" font-size="10" fill="#98a2ae">С</text>
<g stroke="#2f80ed" fill="none" stroke-width="1.4"><circle cx="235" cy="540" r="7"/><path d="M235 522 L235 558 M217 540 L253 540"/></g>
<circle cx="235" cy="540" r="2.4" fill="#2f80ed"/>
<text x="248" y="536" font-family="-apple-system,system-ui,sans-serif" font-size="11" fill="#7f8996">Белокаменка</text>
<text x="248" y="551" font-family="-apple-system,system-ui,sans-serif" font-size="10" fill="#9aa4b0">69°14′ N · 33°17′ E</text>
</svg>"""


@app.get("/login-bg.svg")
def login_bg():
    return Response(content=LOGIN_BG_SVG, media_type="image/svg+xml",
                    headers={"Cache-Control": "public, max-age=86400"})


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


# --- Рабочий стол ------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    today = datetime.now(MSK).date()  # МСК, не UTC — иначе граница "сегодня" съезжает у полуночи

    # Все активные несделанные обязательства; раскладываем по статусу в Python, потому что
    # порог "скоро" теперь per-type (deadlines.lead_days_for), а не плоские 7 дней —
    # одним SQL-диапазоном его не выразить.
    _pending = db.scalars(
        select(Obligation)
        .where(Obligation.is_current == True)  # noqa: E712 — SQLAlchemy требует именно так, не `is True`
        .where(Obligation.status == ObligationStatus.PENDING)
        .order_by(Obligation.deadline_date)
    ).all()
    overdue = [o for o in _pending if _obligation_status(o, today)[2] == "overdue"]
    due_soon = [o for o in _pending if _obligation_status(o, today)[2] == "soon"]

    def obl_row(o: Obligation) -> str:
        emp_name = o.employee.full_name if o.employee else "?"
        chip_class, chip_label, _ = _obligation_status(o, today)
        # medical_exam решается в разделе Медкомиссия (направление/результат), остальные типы
        # (registration, contract_notice, efs1_report) — правкой соответствующей даты в карточке
        # сотрудника, откуда их дедлайн и считается.
        action_url = "/medical" if o.type == ObligationType.MEDICAL_EXAM else f"/employees/{o.employee_id}"
        action_label = "Открыть медкомиссию" if o.type == ObligationType.MEDICAL_EXAM else "Открыть карточку"
        type_label = OBLIGATION_LABELS.get(o.type, o.type.value)
        return (
            f'<div class="card">{emp_name} — {type_label}<br>'
            f'<span class="badge {chip_class}">до {o.deadline_date.strftime("%d.%m.%Y")} · {chip_label}</span><br>'
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

    test_banner = (
        '<div class="warning-banner">⚠ Тестовый режим включён (TEST_ALLOW_MISSING_FIELDS): '
        'документы для медкомиссии генерируются с прочерками вместо незаполненных полей. '
        'Выключите флаг в переменных окружения перед реальной работой с сотрудниками.</div>'
        if TEST_ALLOW_MISSING_FIELDS else ""
    )

    if db.get(SystemFlag, "dactyloscopy_backfill_done") is not None:
        recompute_section = (
            '<section><h2>Разовая операция</h2>'
            '<div class="card"><span class="badge green">Прогон дактилоскопии выполнен</span>'
            '<div class="muted-line">Кнопку и эту секцию можно удалить отдельной правкой '
            '(см. TODO в models.py у таблицы SystemFlag).</div></div></section>'
        )
    else:
        recompute_section = (
            '<section><h2>Разовая операция</h2>'
            '<div class="card">Создать обязанности дактилоскопии для существующих сотрудников — '
            'разово, для тех, кто заведён до добавления правила. Запускать один раз.'
            '<form method="post" action="/admin/recompute-dactyloscopy">'
            '<button type="submit">Запустить прогон</button></form></div></section>'
        )

    body = f"""
{test_banner}
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

{recompute_section}
"""
    return _render("Рабочий стол", body)


# --- Сотрудники: список + единая карточка ------------------------------------

@app.get("/employees", response_class=HTMLResponse)
def employees_list(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    today = datetime.now(MSK).date()
    employees = db.scalars(select(Employee).order_by(Employee.full_name)).all()

    def nearest_pending(e: Employee):
        obs = [o for o in e.obligations if o.is_current and o.status == ObligationStatus.PENDING]
        return min(obs, key=lambda o: o.deadline_date) if obs else None

    def row(e: Employee) -> str:
        cit = e.citizenship or "—"
        if e.consent_status != ConsentStatus.CONFIRMED:
            chip = '<span class="badge neutral">без согласия</span>'
            ob_line = '<div class="muted-line">обязанности создаются после согласия</div>'
        else:
            o = nearest_pending(e)
            if o is None:
                chip = '<span class="badge green">в норме</span>'
                ob_line = '<div class="muted-line">нет активных сроков</div>'
            else:
                cls, lbl, _ = _obligation_status(o, today)
                type_label = OBLIGATION_LABELS.get(o.type, o.type.value)
                chip = f'<span class="badge {cls}">{lbl}</span>'
                ob_line = (
                    f'<div class="muted-line">{type_label} · '
                    f'до {o.deadline_date.strftime("%d.%m.%Y")}</div>'
                )
        return (
            f'<div class="card">{e.full_name} {chip}<br>'
            f'<span class="muted-line">{cit}</span>{ob_line}'
            f'<a class="btn" href="/employees/{e.id}">Открыть карточку</a></div>'
        )

    rows = "".join(row(e) for e in employees) or '<p class="muted">Сотрудников в базе нет.</p>'
    return _render(
        "Сотрудники",
        f'<h1>Сотрудники ({len(employees)})</h1><section class="grid">{rows}</section>',
    )


@app.get("/employees/{employee_id}", response_class=HTMLResponse)
def employee_card(employee_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    today_s = datetime.now(MSK).date().isoformat()

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
<p class="muted">Заполни известные поля и нажми одну кнопку внизу. Пустые поля не трогаются.</p>
<form id="saveform" method="post" action="/employees/{emp.id}/save">
<input type="hidden" name="confirmed" value="">

<fieldset>
<legend>Дата въезда</legend>
<input type="date" name="entry_date" max="{today_s}"
value="{emp.entry_date.isoformat() if emp.entry_date else ''}">
</fieldset>

<fieldset>
<legend>Дактилоскопия («грин карта»)</legend>
<p class="muted">Дата прохождения. Заполнение закрывает обязанность. Пусто — обязанность
горит от даты въезда + 30 дней.</p>
<input type="date" name="dactyloscopy_date" max="{today_s}"
data-orig="{emp.dactyloscopy_date.isoformat() if emp.dactyloscopy_date else ''}"
value="{emp.dactyloscopy_date.isoformat() if emp.dactyloscopy_date else ''}">
</fieldset>

<fieldset>
<legend>Место въезда</legend>
<input type="text" name="entry_country" value="{emp.entry_country or ''}">
</fieldset>

<fieldset>
<legend>Место пребывания</legend>
<p class="muted">Текущий адрес: {emp.address or "не указан"}</p>
<label>Адрес</label>
<input type="text" name="address" data-orig="{emp.address or ''}" value="{emp.address or ''}">
<label>Дата, с которой действует этот адрес</label>
<input type="date" name="address_since" max="{today_s}" value="{today_s}">
<p class="muted">Смена адреса создаёт новое обязательство по регистрации. Первый ввод — нет.</p>
</fieldset>

<fieldset>
<legend>Дата договора</legend>
<input type="date" name="contract_date" max="{today_s}"
value="{emp.contract_date.isoformat() if emp.contract_date else ''}">
</fieldset>

<button type="submit">Сохранить</button>
</form>

<fieldset>
<legend>Согласие на обработку ПД</legend>
{consent_block}
</fieldset>

<a class="btn secondary" href="/employees">← Ко всем сотрудникам</a>
</section>"""
    return _render(emp.full_name, body + SAVE_FORM_JS)


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


@app.post("/employees/{employee_id}/dactyloscopy_date")
def employee_dactyloscopy_date_submit(
    employee_id: str,
    request: Request,
    dactyloscopy_date: date = Form(...),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    emp.dactyloscopy_date = dactyloscopy_date

    # Заполнение даты = дактилоскопия пройдена: закрываем текущую обязанность в DONE.
    # Симметрично тому, как медосмотр закрывается результатом. Если обязанности ещё нет
    # (согласие не подтверждено), просто сохраняем дату — при последующем создании
    # obligations гейт в obligations.py сделает её сразу DONE.
    dact = db.scalars(
        select(Obligation)
        .where(Obligation.employee_id == emp.id)
        .where(Obligation.type == ObligationType.DACTYLOSCOPY)
        .where(Obligation.is_current == True)  # noqa: E712
    ).first()
    if dact is not None:
        dact.status = ObligationStatus.DONE
        db.add(dact)

    db.commit()
    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/employees/{employee_id}/save")
def employee_save(
    employee_id: str,
    request: Request,
    entry_date: str = Form(""),
    entry_country: str = Form(""),
    contract_date: str = Form(""),
    address: str = Form(""),
    address_since: str = Form(""),
    dactyloscopy_date: str = Form(""),
    confirmed: str = Form(""),
    db: Session = Depends(get_db),
):
    """Единое сохранение карточки. Диф по каждому полю относительно БД: применяются только
    реально изменившиеся поля (пустое = не трогаем, не затираем). Опасные изменения — смена
    адреса (создаёт обязательство регистрации) и дата дактилоскопии (закрывает обязательство)
    — требуют подтверждения: JS-confirm() ставит confirmed=1 до отправки, а без JS срабатывает
    серверный второй шаг ниже. Старые пять роутов полей оставлены рабочими рядом."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    def _pdate(s):
        s = (s or "").strip()
        return date.fromisoformat(s) if s else None

    ed = _pdate(entry_date)
    ec = (entry_country or "").strip()
    cd = _pdate(contract_date)
    dd = _pdate(dactyloscopy_date)
    prev_addr = (emp.address or "").strip()
    new_addr = (address or "").strip()

    # что изменилось (сравнение до мутации)
    entry_changed = ed is not None and ed != emp.entry_date
    country_changed = ec != "" and ec != (emp.entry_country or "")
    contract_changed = cd is not None and cd != emp.contract_date
    dact_change = dd is not None and dd != emp.dactyloscopy_date
    addr_real_change = prev_addr != "" and new_addr != "" and prev_addr != new_addr
    addr_first_fill = prev_addr == "" and new_addr != ""

    # опасные изменения -> подтверждение
    dangerous = []
    if addr_real_change:
        dangerous.append("смена адреса места пребывания — создаст обязательство регистрации")
    if dact_change:
        dangerous.append("дата дактилоскопии — закроет обязательство как пройденное")

    if dangerous and confirmed != "1":
        def _hid(name, val):
            return f'<input type="hidden" name="{name}" value="{html.escape(val or "", quote=True)}">'
        items = "".join(f"<li>{html.escape(d)}</li>" for d in dangerous)
        body = f"""
<h1>Подтвердите изменения</h1>
<section class="card-form">
<div class="warning-banner">Эти изменения затронут обязательства сотрудника {html.escape(emp.full_name)}:</div>
<ul>{items}</ul>
<form method="post" action="/employees/{emp.id}/save">
{_hid("entry_date", entry_date)}
{_hid("entry_country", entry_country)}
{_hid("contract_date", contract_date)}
{_hid("address", address)}
{_hid("address_since", address_since)}
{_hid("dactyloscopy_date", dactyloscopy_date)}
<input type="hidden" name="confirmed" value="1">
<button type="submit">Подтвердить и сохранить</button>
</form>
<a class="btn secondary" href="/employees/{emp.id}">Отмена</a>
</section>"""
        return HTMLResponse(_render("Подтверждение", body))

    # --- применяем только изменившееся ---
    if entry_changed:
        emp.entry_date = ed
    if country_changed:
        emp.entry_country = ec
    if contract_changed:
        emp.contract_date = cd
    if dd is not None:
        emp.dactyloscopy_date = dd  # до create_obligations, чтобы гейт увидел
    if addr_real_change:
        emp.address = new_addr
        asd = _pdate(address_since)
        if asd is not None:
            emp.address_since = asd
    elif addr_first_fill:
        emp.address = new_addr  # первый ввод — address_since не трогаем (нет смены)

    db.commit()
    db.refresh(emp)

    # пересоздание обязательств: изменились триггерные поля И согласие подтверждено
    if (entry_changed or contract_changed or addr_real_change) and emp.consent_status == ConsentStatus.CONFIRMED:
        create_obligations_for_employee(db, emp)

    # закрытие дактилоскопии: заполнение даты = пройдено
    if dact_change:
        dact = db.scalars(
            select(Obligation)
            .where(Obligation.employee_id == emp.id)
            .where(Obligation.type == ObligationType.DACTYLOSCOPY)
            .where(Obligation.is_current == True)  # noqa: E712
        ).first()
        if dact is not None and dact.status != ObligationStatus.DONE:
            dact.status = ObligationStatus.DONE
            db.add(dact)
            db.commit()

    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/admin/recompute-dactyloscopy")
def admin_recompute_dactyloscopy(request: Request, db: Session = Depends(get_db)):
    """Разовый прогон: создаёт недостающие обязанности (в т.ч. дактилоскопию) для уже
    заведённых ПОДТВЕРЖДЁННЫХ сотрудников — иначе новое правило оживёт только у тех, чью
    карточку тронут после деплоя. Защищён флагом system_flags: повторный запуск ничего не
    делает (плюс идемпотентность самой create_obligations_for_employee).
    TODO удалить этот эндпоинт и таблицу SystemFlag после успешного прогона."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    if db.get(SystemFlag, "dactyloscopy_backfill_done") is not None:
        return RedirectResponse("/", status_code=303)

    confirmed = db.scalars(
        select(Employee).where(Employee.consent_status == ConsentStatus.CONFIRMED)
    ).all()
    for emp in confirmed:
        create_obligations_for_employee(db, emp)  # сама коммитит и идемпотентна

    db.add(
        SystemFlag(
            key="dactyloscopy_backfill_done",
            value=f"{len(confirmed)} сотрудников; {datetime.now(MSK).isoformat()}",
            updated_at=datetime.now(MSK),
        )
    )
    db.commit()
    return RedirectResponse("/", status_code=303)


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
        missing = check_medical_referral_fields(emp) if emp else []
        missing_line = (
            f'<span class="badge red">нет: {", ".join(missing)}</span><br>' if missing else ""
        )
        return (
            f'<div class="card">{name}<br>'
            f'<span class="badge orange">дедлайн {o.deadline_date.isoformat()}</span><br>'
            f'{missing_line}'
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
            f'<form class="inline" method="post" action="/medical/{r.employee_id}/result" onsubmit="return confirm(&#39;Удалить направление и вернуть сотрудника в очередь на выписку? Действие необратимо.&#39;)"'
            f'<input type="hidden" name="result" value="failed">'
            f'<button type="submit" class="secondary">❌ Не пройдено</button></form>'
            f'</div>'
        )

    test_banner = (
        '<div class="warning-banner">⚠ Тестовый режим включён (TEST_ALLOW_MISSING_FIELDS): '
        'направления с незаполненными полями всё равно выписываются, с прочерками и '
        'предупреждением внутри документа. Выключите флаг перед реальной работой.</div>'
        if TEST_ALLOW_MISSING_FIELDS else ""
    )

    body = f"""
{test_banner}
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
    """Выписывает направление: заводит запись Referral, привязанную к обязательству
    MEDICAL_EXAM (bot.py сам Referral не создаёт — это новая часть учёта, добавленная
    здесь), затем показывает HTML-предпросмотр документа с кнопками "Печать" и
    "Скачать .docx". Сам docx (та же генерация, что /send_medical_referral в bot.py)
    не отдаётся принудительно как download — на экране браузера его не отрисовать
    напрямую (нет нативного просмотра .docx), поэтому печать идёт через HTML-версию
    того же содержания (window.print()), а .docx доступен отдельной кнопкой на скачивание.

    2026-07: список отсутствующих полей запрашивается ЗАРАНЕЕ через
    check_medical_referral_fields — не только чтобы решить, кидать ли 400, но и чтобы
    показать баннер в HTML-превью независимо от того, кинул генератор исключение или
    сработал тестовый обход (TEST_ALLOW_MISSING_FIELDS)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    obligation = db.get(Obligation, obligation_id)
    if obligation is None or obligation.employee_id != employee_id:
        raise HTTPException(404, "Обязательство не найдено или не принадлежит этому сотруднику")

    missing = check_medical_referral_fields(emp)
    if missing and not TEST_ALLOW_MISSING_FIELDS:
        raise HTTPException(
            400,
            f"Нельзя сгенерировать документ для {emp.full_name} — "
            f"не заполнены поля: {', '.join(missing)}. "
            f"Заполните их в карточке сотрудника перед генерацией.",
        )

    try:
        # Генерируем один раз здесь ТОЛЬКО ради проверки обязательных полей
        # (_require_fields в document_templates.py) — если чего-то не хватает,
        # ValueError всплывёт до создания записи Referral, а не после. В тестовом
        # режиме это не бросит исключение даже при missing — документ сгенерируется
        # с прочерками, missing уже посчитан выше для баннера.
        generate_medical_referral_docx(emp)
    except ValueError as e:
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

    return HTMLResponse(_render_referral_preview(emp, obligation_id, missing_fields=missing))


@app.get("/medical/{employee_id}/referral/{obligation_id}/download")
def medical_referral_download(
    employee_id: str,
    obligation_id: str,
    request: Request,
    db: Session = Depends(get_db),
):
    """Отдельный маршрут на скачивание — регенерирует docx по текущим данным сотрудника.
    Не хранит путь к файлу между запросами (файловая система Railway эфемерна между
    процессами), поэтому пересоздаёт документ заново при каждом скачивании. Единственный
    практический нюанс: поле "10. Дата выдачи направления" в документе — дата генерации
    файла, а не дата исходного нажатия "Выписать направление"; если скачать на следующий
    день после того, как направление уже выписано, дата в файле сдвинется на сегодня."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    try:
        path = generate_medical_referral_docx(emp)
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception:
        raise HTTPException(500, "Не удалось сгенерировать направление. Проверьте логи сервиса.")

    filename = f"Направление_{emp.full_name.replace(' ', '_')}.docx"
    return FileResponse(path, filename=filename)


def _render_referral_preview(emp: Employee, obligation_id: str, missing_fields: list[str] | None = None) -> str:
    """HTML-версия направления для экрана/печати — содержание зеркалит
    generate_medical_referral_docx в document_templates.py. Если меняешь текст/поля
    там — поменяй и здесь, иначе предпросмотр разойдётся с реальным .docx.

    missing_fields: если непусто (только в тестовом режиме), показывает баннер над
    документом — то же предупреждение, что вставлено в сам docx."""
    birth = emp.birth_date.strftime("%d.%m.%Y") if emp.birth_date else "—"
    name_parts = emp.full_name.split()
    surname = name_parts[0] if name_parts else "—"
    first_name = name_parts[1] if len(name_parts) > 1 else "—"
    patronymic = name_parts[2] if len(name_parts) > 2 else "—"

    download_url = f"/medical/{emp.id}/referral/{obligation_id}/download"

    warning_banner = ""
    if missing_fields:
        warning_banner = (
            '<div class="warning-banner">⚠ ТЕСТОВЫЙ ЧЕРНОВИК — не заполнены поля: '
            + ", ".join(missing_fields)
            + '. Документ не имеет юридической силы, пока эти поля не указаны в карточке '
            "сотрудника и документ не перегенерирован.</div>"
        )

    body = f"""
<style>
@media print {{
  nav, .no-print {{ display: none !important; }}
  body {{ background: #fff !important; }}
  section {{ box-shadow: none !important; border: none !important; }}
}}
.referral-doc p {{ margin: 4px 0; }}
</style>

<h1>Направление на медосмотр</h1>

{warning_banner}

<div class="no-print" style="margin-bottom:14px">
<button onclick="window.print()">🖨 Печать</button>
<a class="btn secondary" href="{download_url}">⬇ Скачать .docx</a>
<a class="btn secondary" href="/medical">← К медкомиссии</a>
</div>

<section class="narrow referral-doc">
<p style="text-align:right">к Договору № {CLINIC_CONTRACT_NUMBER} от «{CLINIC_CONTRACT_DATE}»<br>Приложение № 1</p>
<p style="text-align:center"><strong>НАПРАВЛЕНИЕ НА МЕДИЦИНСКОЕ ОСВИДЕТЕЛЬСТВОВАНИЕ</strong></p>
<p>В {REFERRAL_CLINIC_NAME}</p>
<p class="muted">наименование медицинской организации (МО)</p>
<p>1. Фамилия {surname}</p>
<p>Имя {first_name}</p>
<p>Отчество {patronymic}</p>
<p>2. Дата рождения (число, месяц, год) {birth}</p>
<p>3. Адрес (по месту проживания) {emp.address or "—"}</p>
<p>4. Серия паспорта {emp.passport_series or "—"} Номер паспорта {emp.passport_number or "—"}</p>
<p>5. Место работы {REFERRAL_PAYER_NAME}</p>
<p>6. Наименование медицинской услуги (медицинского освидетельствования)</p>
<p>{MEDICAL_SERVICE_TEXT}</p>
<p>7. Дата проведения услуги _____________ кабинет N _____ время _____</p>
<p>8. Полное наименование организации, направившей иностранного гражданина, телефон {PAYER_PHONE}</p>
<p>{REFERRAL_PAYER_NAME}</p>
<p>подпись, печать _____________________</p>
<p>10. Дата выдачи направления {date.today().strftime('%d.%m.%Y')}</p>
<br>
<table style="width:100%">
<tr><td><strong>От Исполнителя:</strong></td><td><strong>От Заказчика:</strong></td></tr>
<tr><td>{REFERRAL_CLINIC_SHORT_NAME}</td><td>Индивидуальный предприниматель</td></tr>
<tr><td>Главный врач</td><td>&nbsp;</td></tr>
<tr><td>_____________________ {CLINIC_CHIEF_DOCTOR_NAME}<br>м.п.</td>
<td>_____________________ {PAYER_SIGNATORY_NAME}<br>м.п.</td></tr>
</table>
</section>
"""
    return _render("Направление на медосмотр", body)


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

    if result == "failed":
        # ВРЕМЕННО (тест): "не пройдено/не явился" = удаляем текущее направление, чтобы
        # сотрудник вернулся в список "Выписать направление" (фильтр need_referral —
        # "нет связанного Referral"). Обязательство медосмотра остаётся PENDING, дедлайн жив.
        # ВНИМАНИЕ: cascade="all, delete-orphan" на Referral.invoices — удаление сносит и
        # связанные Invoice. Для теста приемлемо (счетов ещё нет). Позже заменить на статус
        # ExamStatus.CANCELLED с сохранением истории и счёта — отложенная задача.
        db.delete(referral)
        db.commit()
        return RedirectResponse("/medical", status_code=303)

    # result == "done": медосмотр пройден — направление завершается, обязательство закрывается.
    referral.exam_status = ExamStatus.COMPLETED
    referral.result_date = date.today()
    db.add(referral)

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
#      TEST_ALLOW_MISSING_FIELDS — "true", ЧТОБЫ ВРЕМЕННО разрешить генерацию направлений
#                                   с прочерками при незаполненных полях (см. document_templates.py).
#                                   Убрать/поставить "false" перед реальной работой с сотрудниками.
# 5. Railway выдаст публичный URL сервиса (Settings → Networking → Generate Domain).
