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
import json
import logging
import os
import calendar
from datetime import date, datetime, timedelta, timezone

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session, sessionmaker
from starlette.exceptions import HTTPException as StarletteHTTPException
from starlette.middleware.sessions import SessionMiddleware

log = logging.getLogger("webforms")

from models import (
    User,
    UserRole,
    UserStatus,
    AttendanceMark,
    Category,
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
    RegistrationStatus,
    RotationReturn,
    SystemFlag,
    WorkOrderMember,
    WorkOrderMemberChange,
    WorkLogEntry,
    MemberChangeType,
    WorkOrder,
    WorkType,
    InstructionType,
    Certificate,
    InternalOrder,
    OrderCategory,
    Brigade,
    BrigadeMember,
    Instruction,
    Titul,
)
from obligations import create_obligations_for_employee
from auth_binding import get_role_label, find_user_by_max_id
import tabel
from deadlines import lead_days_for
from document_templates import (
    CLINIC_CHIEF_DOCTOR_NAME,
    CLINIC_CONTRACT_DATE,
    CLINIC_CONTRACT_NUMBER,
    CLINIC_NAME as REFERRAL_CLINIC_NAME,
    CLINIC_SHORT_NAME as REFERRAL_CLINIC_SHORT_NAME,
    SITE_ADDRESS,
    MEDICAL_SERVICE_TEXT,
    PAYER_NAME as REFERRAL_PAYER_NAME,
    PAYER_PHONE,
    PAYER_SIGNATORY_NAME,
    TEST_ALLOW_MISSING_FIELDS,
    check_medical_referral_fields,
    generate_medical_referral_docx,
    generate_labor_contract_docx,
    generate_duty_receipt_docx,
    generate_termination_notice_docx,
    generate_departure_notice_docx,
    CONTRACT_NUMBER_PREFIX,
    EMPLOYER_NAME_SHORT,
    EMPLOYER_NAME_FULL,
    EMPLOYER_DIRECTOR_FULL,
    EMPLOYER_DIRECTOR_SHORT,
    EMPLOYER_INN,
    EMPLOYER_KPP,
    EMPLOYER_LEGAL_ADDRESS,
    EMPLOYER_ACTUAL_ADDRESS,
    EMPLOYER_PHONE,
    EMPLOYER_SUBDIVISION,
    WORKPLACE_ADDRESS,
    DISTRICT_COEFFICIENT,
    SITE_ADDRESS as CONTRACT_SITE_ADDRESS,
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

# Старый общий вход убран (перешли на учётки). Переменные оставлены необязательными,
# чтобы их удаление/отсутствие не роняло старт. created_by/proof теперь берут реального
# пользователя из сессии (см. _actor_name).
WEBFORMS_USER = os.environ.get("WEBFORMS_USER", "system")
WEBFORMS_PASSWORD = os.environ.get("WEBFORMS_PASSWORD", "")
SECRET_KEY = os.environ["WEBFORMS_SECRET_KEY"]
ORG_NAME = os.environ.get("COMPANY_NAME", "ИП Буц Сергей Юрьевич")
CLINIC_ID = os.environ.get("CLINIC_ID", "pirogova_murmansk")
CONSENT_TEXT_VERSION = os.environ.get("CONSENT_TEXT_VERSION", "v1")

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=SECRET_KEY, session_cookie="migbot_session")


# --- Хранилище сканов (S3-совместимое, Cloud.ru Object Storage) --------------
# Сканы для пакета Госуслуг (паспорт, миграционная карта, выписка ЕГРН на ВЖК) хранятся в
# приватном бакете Cloud.ru. Ключи — только в переменных окружения Railway, не в коде.
# Доступ к сканам: кадровик и админ (не прораб) — проверяется в роутах.
# Если boto3 не установлен или ключи не заданы — функции бросают понятную ошибку, старт не
# роняется (аналогично graceful-обработке holidays/qrcode).
# S3-хранилище вынесено в общий модуль s3_storage.py (им пользуется и bot.py).
from s3_storage import (
    S3_ENDPOINT, S3_BUCKET, S3_REGION, S3_ACCESS_KEY, S3_SECRET_KEY,
    SCAN_TYPES, PAYMENT_SCAN_TYPES, COMMON_DOC_TYPES, SCAN_COMMON_TYPES,
    _s3_client, _scan_key, _s3_upload, _s3_list_for_employee, _s3_download,
    _s3_delete, _common_key, _s3_upload_common, _s3_list_common,
    _s3_download_common, _s3_delete_common, _s3_clear_check,
)


def _package_missing(emp, present: dict | None = None, common: dict | None = None) -> list:
    """Проверяет комплектность пакета Госуслуг. Принимает уже посчитанные списки сканов present
    (персональные) и common (общие), чтобы НЕ ходить в S3 повторно — это ускоряет карточку, где
    списки уже посчитаны для отображения. Если не переданы — считает сам (для роутов пакета).
    Платёжки пока не обязательны (логика двух госпошлин будет позже)."""
    if present is None:
        present = _s3_list_for_employee(emp.id)
    if common is None:
        common = _s3_list_common()
    missing = []
    # Не требуем для комплектности: платёжки (логика пошлин позже) и уведомление о прибытии
    # (это подтверждение постановки на учёт, а не документ, подаваемый В пакете Госуслуг).
    _optional = PAYMENT_SCAN_TYPES | {"arrival_notice"}
    for st, label in SCAN_TYPES.items():
        if st in _optional:
            continue
        _i = present.get(st) or {}
        if not _i.get("present"):
            missing.append(label)
    for dt, label in COMMON_DOC_TYPES.items():
        if not common.get(dt):
            missing.append(label)
    if emp.contract_date is None:
        missing.append("Трудовой договор (не заключён)")
    # Для загранпаспорта — нужны все страницы (подтверждается чекбоксом кадровика).
    _is_passport = (emp.doc_type == "passport") or (
        not emp.doc_type and (emp.passport_series or "").strip().upper() != "ID"
    )
    if _is_passport and not getattr(emp, "passport_all_pages", False):
        missing.append("Паспорт: подтвердите загрузку всех страниц")
    return missing


def _content_disposition(filename: str) -> str:
    """Заголовок Content-Disposition с именем файла, безопасным для HTTP (latin-1 only).
    Русские имена кодируем по RFC 5987 (filename*=UTF-8''...), плюс ASCII-запасной filename,
    иначе UnicodeEncodeError: latin-1 не кодирует кириллицу в заголовке."""
    from urllib.parse import quote
    ascii_name = filename.encode("ascii", "ignore").decode("ascii") or "file"
    quoted = quote(filename)
    return f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quoted}"


def _ext_for(ct: str) -> str:
    if "pdf" in ct:
        return "pdf"
    if "jpeg" in ct or "jpg" in ct:
        return "jpg"
    if "png" in ct:
        return "png"
    return "bin"



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


def _hash_password(pw: str) -> str:
    """Прямой bcrypt (не passlib-обёртка): passlib 1.7.x несовместим с bcrypt 4.x
    (лезет в bcrypt.__about__, которого больше нет), из-за чего verify падал и вход не
    проходил при верном пароле. bcrypt.hashpw/checkpw читают и $2a$, и $2b$ без проблем."""
    import bcrypt as _bcrypt
    return _bcrypt.hashpw(pw.encode("utf-8"), _bcrypt.gensalt(rounds=12)).decode("utf-8")


def _verify_password(pw: str, pw_hash: str) -> bool:
    import bcrypt as _bcrypt
    try:
        return _bcrypt.checkpw(pw.encode("utf-8"), pw_hash.encode("utf-8"))
    except Exception:
        return False


def _password_problems(pw: str) -> list[str]:
    """Требования: 8+ символов, строчная, заглавная, цифра. Возвращает список нарушений."""
    import re as _re
    problems = []
    if len(pw) < 8:
        problems.append("минимум 8 символов")
    if not _re.search(r"[a-zа-яё]", pw):
        problems.append("строчная буква")
    if not _re.search(r"[A-ZА-ЯЁ]", pw):
        problems.append("заглавная буква")
    if not _re.search(r"\d", pw):
        problems.append("цифра")
    return problems


from common_utils import normalize_phone as _normalize_phone
# _normalize_phone вынесена в common_utils.py — используется и здесь, и в bot.py.
# Оставлена под старым именем (алиас), чтобы не трогать все места вызова в этом файле.


def _current_user(request: Request, db: Session):
    """Пользователь из сессии или None. Проверяет, что запись существует и APPROVED."""
    uid = request.session.get("user_id")
    if not uid:
        return None
    user = db.get(User, uid)
    if user is None or user.status != UserStatus.APPROVED:
        return None
    return user


def _logged_in(request: Request) -> bool:
    return bool(request.session.get("user_id"))


def _require_role(request: Request, db: Session, *roles):
    """Проверяет, что текущий пользователь имеет одну из ролей. Иначе 403.
    ADMIN проходит везде (полный доступ). Возвращает пользователя при успехе."""
    user = _current_user(request, db)
    if user is None:
        raise HTTPException(401, "Требуется вход")
    if user.role == UserRole.ADMIN:
        return user
    if user.role not in roles:
        raise HTTPException(403, "Недостаточно прав для этого действия")
    return user


def _actor_name(request: Request, db: Session) -> str:
    """Имя текущего пользователя для аудита (created_by, proof). Если сессия почему-то пуста —
    'system'. Заменяет прежний фиксированный WEBFORMS_USER, чтобы след показывал, КТО сделал."""
    user = _current_user(request, db)
    return user.full_name if user else "system"


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
    ObligationType.DEPARTURE_NOTICE: "снятие с учёта (убытие)",
    ObligationType.PATENT_PAYMENT: "оплата патента",
}


# --- Простая HTML-обёртка без отдельных файлов шаблонов ---------------------

PAGE_HEAD = """<!doctype html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
:root{{--ink:#111214;--sub:#5b626b;--line:#e6e9ee;--line-2:#eef1f4;--accent:#4a90e2;--accent-ink:#2f6fb0;
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
input[type=date],input[type=text],input[type=password],select{{width:100%;padding:12px;font-size:16px;font-family:inherit;border:1px solid #d9dde3;border-radius:12px;margin:6px 0 12px;background:#fff;color:var(--ink)}}
select{{min-height:48px;-webkit-appearance:none;appearance:none;background-image:url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="%235b626b" stroke-width="2"><path d="M6 9l6 6 6-6"/></svg>');background-repeat:no-repeat;background-position:right 14px center;padding-right:40px}}
.field-help{{margin:2px 0 8px}}
.field-help summary{{list-style:none;cursor:pointer;display:inline-flex;align-items:center;gap:6px;color:var(--accent-ink);font-size:13px;font-weight:600}}
.field-help summary::-webkit-details-marker{{display:none}}
.field-help .i{{display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;border:1.5px solid var(--accent);color:var(--accent-ink);font-size:12px;font-style:italic;font-weight:700;line-height:1}}
.field-help p{{margin:6px 0 0;color:var(--sub);font-size:13px;font-weight:400;line-height:1.4}}
.btn-full{{width:100%;text-align:center;margin:0 0 14px 0}}
.btn-upload{{background:#eaf0fb;color:var(--accent);border:1px solid #cddcf5;font-weight:600}}
.scans-grid > div{{margin-top:0}}
.btn-upload:hover{{background:#dce8fa}}
input:focus{{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px #4a90e222}}
label{{font-size:13px;color:var(--sub)}}
.badge{{display:inline-block;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:600;margin:2px 4px 2px 0}}
.badge.red{{background:var(--red-bg);color:var(--red-ink)}}
.badge.orange{{background:var(--amber-bg);color:var(--amber-ink)}}
.badge.green{{background:var(--green-bg);color:var(--green-ink)}}
.badge.neutral{{background:var(--neutral-bg);color:var(--neutral-ink)}}
.muted{{color:var(--sub);font-size:13px}}
.warning-banner{{background:var(--amber-bg);border:1px solid #f0c674;border-left:4px solid var(--amber-ink);border-radius:12px;padding:12px 14px;margin-bottom:14px;font-weight:600;color:#7a4a00}}
nav{{margin-bottom:16px;background:#fff;border:1px solid var(--line);border-radius:12px;padding:4px 8px;display:flex;flex-wrap:wrap;gap:2px}}
nav a{{color:var(--sub);text-decoration:none;font-size:15px;padding:10px 12px 8px;white-space:nowrap;font-weight:600;border-radius:8px 8px 0 0;border-bottom:2px solid transparent}}
nav a.active{{color:var(--ink);border-bottom-color:var(--accent)}}
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
  section.card-form{{max-width:1000px;margin:0 auto}}
  .card-cols{{display:grid;grid-template-columns:1fr 1fr;gap:20px;align-items:start}}
  a.btn:hover,button:hover{{opacity:.9}}
  /* Десктоп: кнопки на всю ширину (btn-full) не растягивать во всю колонку — ограничить и
     прижать влево, иначе на широком экране кнопка-переросток. Мобильный не затронут (media). */
  .btn-full{{width:auto;min-width:280px;display:inline-block}}
  .btn-upload{{min-width:auto;padding:8px 18px;font-size:14px}}
  .scans-grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px;align-items:start}}
  /* Поля ввода не на всю ширину колонки — читаемая ширина. */
  input[type=date],input[type=text],input[type=password],select{{max-width:420px}}
  /* Форма квитанции: кнопки в ряд, а не столбиком. */
  fieldset form .btn-full{{margin-right:10px}}
}}
</style></head><body>
<header class="org">
<div class="org-name">{org_name}</div>
<div class="page-title">Автоматизированная система учёта на производстве — {title}</div>
</header>
"""
PAGE_FOOT = """
<button id="scrollTopBtn" onclick="window.scrollTo({top:0,behavior:'smooth'})"
  style="display:none;position:fixed !important;right:12px;bottom:calc(16px + env(safe-area-inset-bottom,0px));
  z-index:999;width:46px !important;height:46px !important;min-height:0 !important;min-width:0 !important;
  padding:0 !important;margin:0 !important;border:none;border-radius:50% !important;
  background:rgba(74,144,226,.9);color:#fff;font-size:22px;line-height:46px !important;text-align:center;
  box-shadow:0 3px 10px rgba(20,24,30,.28);cursor:pointer" aria-label="Наверх">&#8593;</button>
<script>
(function(){
  var btn = document.getElementById('scrollTopBtn');
  if(!btn) return;
  window.addEventListener('scroll', function(){
    btn.style.display = (window.scrollY > 300) ? 'block' : 'none';
  });
})();
</script>
</body></html>"""
def _nav(active: str = "", role: str = "") -> str:
    """active: 'home' | 'employees' | 'medical' | 'admin' — подсвечивает корневой раздел.
    role: роль пользователя — 'admin' добавляет пункт «Пользователи» (управление доступом)."""
    items = [
        ("home", "/", "Рабочий стол"),
        ("tabel", "/tabel", "Табель"),
        ("employees", "/employees", "Сотрудники"),
        ("medical", "/medical", "Медкомиссия"),
        ("production", "/production", "Производство"),
        ("reports", "/reports", "Отчёты"),
    ]
    # Общие документы (паспорт директора, основание на адрес) — кадровику и админу, не прорабу.
    if role in ("admin", "kadrovik"):
        items.append(("common", "/common-docs", "Общие документы"))
    if role == "admin":
        items.append(("admin", "/admin/users", "Пользователи"))
    items.append(("", "/logout", "Выйти"))
    links = "".join(
        f'<a href="{href}"{" class=\"active\"" if key and key == active else ""}>{label}</a>'
        for key, href, label in items
    )
    return f"<nav>{links}</nav>"


def _render(title: str, body: str, active: str = "", role: str = "") -> str:
    return PAGE_HEAD.format(title=title, org_name=ORG_NAME) + _nav(active, role) + body + PAGE_FOOT


LOGIN_HEAD = """<!doctype html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Вход</title>
<style>
:root{--accent:#4a90e2;--ink:#111214;--sub:#5b626b;--serif:Georgia,"Times New Roman",serif;--sans:-apple-system,Segoe UI,Roboto,Arial,sans-serif}
*{box-sizing:border-box}
body.login-page{margin:0;min-height:100dvh;display:flex;flex-direction:column;justify-content:flex-start;align-items:center;background:url('/login-bg.svg') no-repeat center bottom / cover, #fff;font-family:var(--sans);color:var(--ink)}
.auth{width:100%;max-width:440px;margin:0 auto;padding:56px 24px 40px}
.auth-row{background:#fff;border:1px solid #e6e9ee;border-radius:16px;padding:14px;box-shadow:0 6px 24px rgba(20,24,30,.10);display:flex;gap:10px;flex-wrap:wrap;align-items:stretch}
.auth h1{font-family:var(--serif);font-weight:700;letter-spacing:-.02em;font-size:clamp(1.5rem,5.5vw,2.25rem);line-height:1.15;margin:0 0 .3em}
.auth .subtitle{font-family:var(--serif);font-weight:400;color:var(--sub);font-size:clamp(1.125rem,2.5vw,1.375rem);margin:0 0 1.75rem}
.auth input{flex:1 1 45%;min-width:130px;font-family:var(--sans);font-size:16px;padding:14px 16px;border:1px solid #b8c0cc;border-radius:12px;background:#fff;margin:0}
.auth input:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px #4a90e222}
.auth button{flex:1 1 100%;border:0;border-radius:12px;background:var(--accent);color:#fff;font:16px/1 var(--sans);padding:15px 20px;min-height:50px;cursor:pointer}
.auth .err{color:#c0392b;font-size:14px;margin:0 0 12px}
.auth a.btn{display:inline-block;margin-top:14px;color:var(--accent);text-decoration:underline;font-size:14px}
@media(min-width:520px){ .auth button{flex:0 0 auto} }
@media(min-width:768px){
  body.login-page{justify-content:center;align-items:flex-start;background-position:center bottom;background-size:min(70vw,820px) auto}
  .auth{margin:0;padding:0 24px 0 clamp(24px,6vw,96px)}
}
</style></head>
<body class="login-page">
"""


@app.get("/login", response_class=HTMLResponse)
def login_form():
    return LOGIN_HEAD + """
<form class="auth" method="post" action="/login" autocomplete="on">
<h1>Автоматизированная система учёта на производстве</h1>
<p class="subtitle">Вход по номеру телефона</p>
<div class="auth-row">
<input type="text" name="phone" placeholder="Телефон" autocomplete="username" required>
<input type="password" name="password" placeholder="Пароль" autocomplete="current-password" required>
<button type="submit">Войти</button>
</div>
<p style="margin-top:14px;text-align:center"><a href="/register">Нет доступа? Подать заявку</a></p>
</form>
</body></html>"""


def _login_error(msg: str, code: int = 401):
    return HTMLResponse(
        LOGIN_HEAD + f"""
<div class="auth">
<h1>Автоматизированная система учёта на производстве</h1>
<p class="err">{msg}</p>
<a class="btn" href="/login">← Назад</a>
</div>
</body></html>""",
        status_code=code,
    )


@app.post("/login")
def login_submit(request: Request, phone: str = Form(...), password: str = Form(...),
                 db: Session = Depends(get_db)):
    """Вход по телефону+паролю. Защита от подбора: 5 неудачных попыток -> блок на час.
    Проверяем статус (APPROVED) и временную блокировку (locked_until) до проверки пароля."""
    norm = _normalize_phone(phone)
    user = db.scalars(select(User).where(User.phone == norm)).first()

    # неизвестный телефон — общая ошибка (не раскрываем, есть ли такой пользователь)
    if user is None:
        log.warning("Вход: неизвестный телефон %s", norm)
        return _login_error("Неверный телефон или пароль")

    now = datetime.utcnow()

    # временная блокировка за подбор
    if user.locked_until is not None and user.locked_until > now:
        mins = int((user.locked_until - now).total_seconds() // 60) + 1
        log.warning("Вход: телефон %s заблокирован ещё %s мин", norm, mins)
        return _login_error(f"Слишком много попыток. Вход заблокирован на ~{mins} мин.")

    # статус аккаунта
    if user.status == UserStatus.PENDING:
        return _login_error("Заявка ещё не одобрена администратором.")
    if user.status == UserStatus.BLOCKED:
        return _login_error("Доступ заблокирован администратором.")

    # проверка пароля
    if not _verify_password(password, user.password_hash):
        user.failed_attempts = (user.failed_attempts or 0) + 1
        if user.failed_attempts >= 5:
            from datetime import timedelta
            user.locked_until = now + timedelta(hours=1)
            user.failed_attempts = 0
            log.warning("Вход: телефон %s — 5 неудач, блок на час", norm)
            db.commit()
            return _login_error("Слишком много попыток. Вход заблокирован на 1 час.")
        db.commit()
        return _login_error("Неверный телефон или пароль")

    # успех — сброс счётчика, установка сессии
    user.failed_attempts = 0
    user.locked_until = None
    db.commit()
    request.session["user_id"] = user.id
    request.session["role"] = user.role.value if user.role else ""
    return RedirectResponse("/", status_code=303)


@app.get("/register", response_class=HTMLResponse)
def register_form():
    return LOGIN_HEAD + """
<form class="auth" method="post" action="/register" autocomplete="on">
<h1>Заявка на доступ</h1>
<p class="subtitle">Заполните — администратор одобрит и назначит роль</p>
<div class="auth-row">
<input type="text" name="full_name" placeholder="ФИО" required>
<input type="text" name="phone" placeholder="Телефон" autocomplete="username" required>
<input type="password" name="password" placeholder="Пароль (8+, буквы, цифра, регистр)" autocomplete="new-password" required>
<button type="submit">Подать заявку</button>
</div>
<p style="margin-top:14px;text-align:center"><a href="/login">← Уже есть доступ</a></p>
</form>
</body></html>"""


@app.post("/register")
def register_submit(request: Request, full_name: str = Form(...), phone: str = Form(...),
                    password: str = Form(...), db: Session = Depends(get_db)):
    """Открытая регистрация: создаёт пользователя со status=PENDING. Роль НЕ назначается
    (её даст админ при одобрении). Проверка сложности пароля. Уведомление админу — в bot.py."""
    norm = _normalize_phone(phone)
    problems = _password_problems(password)
    if problems:
        return _login_error("Пароль слабый: " + ", ".join(problems) + ".")
    if not norm or len(norm) < 10:
        return _login_error("Неверный номер телефона.")
    existing = db.scalars(select(User).where(User.phone == norm)).first()
    if existing is not None:
        return _login_error("Пользователь с таким телефоном уже существует.")
    user = User(
        phone=norm,
        password_hash=_hash_password(password),
        full_name=full_name.strip(),
        status=UserStatus.PENDING,
    )
    db.add(user)
    db.commit()
    log.info("Новая заявка на доступ: %s (%s)", full_name, norm)
    # 2026-07: код для привязки MAX-аккаунта — см. auth_binding.generate_max_confirm_code
    # и docstring поля pending_max_code в models.py.
    from auth_binding import generate_max_confirm_code
    confirm_code = generate_max_confirm_code(db, user)
    # TODO(bot): уведомить админа в MAX о новой заявке (реализуется в bot.py)
    return HTMLResponse(
        LOGIN_HEAD + f"""
<div class="auth">
<h1>Заявка отправлена</h1>
<p class="subtitle">Администратор одобрит доступ и назначит роль.</p>
<div class="card" style="text-align:center">
<p style="margin:0 0 8px">Чтобы пользоваться ботом в MAX без повторного ввода телефона —
отправьте боту команду:</p>
<p style="font-size:22px;font-weight:700;letter-spacing:2px;margin:8px 0">/confirm {confirm_code}</p>
<p class="muted" style="margin:0">Код действует 30 минут. Можно пропустить этот шаг и
позже выполнить /login с телефоном — тоже сработает.</p>
</div>
<a class="btn" href="/login">← К входу</a>
</div>
</body></html>""",
    )


LOGIN_BG_SVG = """<svg viewBox="0 0 430 760" preserveAspectRatio="xMidYMid slice" xmlns="http://www.w3.org/2000/svg">
<g fill="none" stroke="#aab4c2" stroke-width="1">
<path d="M-40 300 Q215 280 470 300"/><path d="M-40 360 Q215 340 470 360"/>
<path d="M-40 420 Q215 400 470 420"/><path d="M-40 480 Q215 460 470 480"/>
<path d="M-40 540 Q215 520 470 540"/><path d="M-40 600 Q215 580 470 600"/>
<path d="M-40 660 Q215 640 470 660"/><path d="M-40 720 Q215 700 470 720"/>
<path d="M60 760 Q120 500 175 250"/><path d="M150 760 Q180 500 205 250"/>
<path d="M240 760 Q235 500 235 250"/><path d="M330 760 Q290 500 265 250"/>
<path d="M420 760 Q350 500 295 250"/></g>
<g stroke="#8f9aa8" stroke-width="1"><path d="M231 300 L239 300 M231 360 L239 360 M231 420 L239 420 M231 480 L239 480 M231 600 L239 600 M231 660 L239 660 M231 720 L239 720"/></g>
<g fill="none" stroke="#9ba6b3" stroke-width="1">
<path d="M120 760 C150 660 120 600 175 545 C215 505 205 460 250 430"/>
<path d="M170 560 l-6 -4 M158 588 l-6 -4 M150 618 l-6 -4 M146 650 l-6 -4"/></g>
<g stroke="#8f9aa8" fill="none" stroke-width="1"><path d="M392 300 L392 330 M392 300 L387 309 M392 300 L397 309"/></g>
<text x="388" y="292" font-family="-apple-system,system-ui,sans-serif" font-size="10" fill="#7f8a98">С</text>
<g stroke="#2f80ed" fill="none" stroke-width="1.4"><circle cx="235" cy="540" r="7"/><path d="M235 522 L235 558 M217 540 L253 540"/></g>
<circle cx="235" cy="540" r="2.4" fill="#2f80ed"/>
<text x="248" y="536" font-family="-apple-system,system-ui,sans-serif" font-size="11" fill="#68737f">Белокаменка</text>
<text x="248" y="551" font-family="-apple-system,system-ui,sans-serif" font-size="10" fill="#828d9a">69°14′ N · 33°17′ E</text>
</svg>"""


@app.get("/login-bg.svg")
def login_bg():
    return Response(content=LOGIN_BG_SVG, media_type="image/svg+xml",
                    headers={"Cache-Control": "public, max-age=86400"})


@app.get("/admin/users", response_class=HTMLResponse)
def admin_users(request: Request, db: Session = Depends(get_db)):
    """Управление пользователями. Только ADMIN. Список: заявки (PENDING) сверху с кнопками
    одобрения и выбором роли; активные и заблокированные ниже с действиями."""
    user = _current_user(request, db)
    if user is None:
        return RedirectResponse("/login", status_code=303)
    if user.role != UserRole.ADMIN:
        raise HTTPException(403, "Доступ только для администратора")

    users = db.scalars(select(User).order_by(User.created_at.desc())).all()
    now = datetime.utcnow()

    def _row(u):
        locked = u.locked_until is not None and u.locked_until > now
        role_ru = {"prorab": "прораб", "kadrovik": "кадровик", "admin": "админ"}.get(
            u.role.value if u.role else "", "—")
        parts = [f'<b>{u.full_name}</b> · {u.phone} · роль: {role_ru}']
        if u.status == UserStatus.PENDING:
            parts.append(
                f'<form method="post" action="/admin/users/{u.id}/approve" style="margin:6px 0">'
                '<select name="role" required>'
                '<option value="">— назначить роль —</option>'
                '<option value="prorab">Прораб (только чтение и скачивание)</option>'
                '<option value="kadrovik">Кадровик (всё по работникам)</option>'
                '<option value="admin">Админ</option>'
                '</select> '
                '<button type="submit">Одобрить</button></form>'
            )
        else:
            if u.status == UserStatus.BLOCKED:
                parts.append('<span class="badge red">заблокирован</span> ')
                parts.append(
                    f'<form method="post" action="/admin/users/{u.id}/unblock" style="display:inline">'
                    '<button type="submit">Разблокировать</button></form>')
            else:
                parts.append('<span class="badge green">активен</span> ')
                if u.id != user.id:  # себя не блокируем
                    parts.append(
                        f'<form method="post" action="/admin/users/{u.id}/block" style="display:inline">'
                        '<button type="submit" class="secondary">Заблокировать</button></form>')
            if locked:
                mins = int((u.locked_until - now).total_seconds() // 60) + 1
                parts.append(
                    f' <span class="badge orange">замок за попытки ~{mins} мин</span>'
                    f'<form method="post" action="/admin/users/{u.id}/unlock" style="display:inline">'
                    '<button type="submit">Снять замок</button></form>')
        return '<fieldset><legend>' + ("заявка" if u.status == UserStatus.PENDING else "пользователь") + '</legend>' + "".join(parts) + '</fieldset>'

    pending = [u for u in users if u.status == UserStatus.PENDING]
    others = [u for u in users if u.status != UserStatus.PENDING]
    body = "<h1>Пользователи</h1>"
    if pending:
        body += "<h2>Заявки на доступ</h2>" + "".join(_row(u) for u in pending)
    else:
        body += '<p class="muted">Новых заявок нет.</p>'
    body += "<h2>Все пользователи</h2>" + "".join(_row(u) for u in others)
    return _render("Пользователи", body, active="admin", role="admin")


@app.post("/admin/users/{user_id}/approve")
def admin_approve(user_id: str, request: Request, role: str = Form(...),
                  db: Session = Depends(get_db)):
    admin = _current_user(request, db)
    if admin is None or admin.role != UserRole.ADMIN:
        raise HTTPException(403, "Только администратор")
    u = db.get(User, user_id)
    if u is None:
        raise HTTPException(404, "Пользователь не найден")
    if role not in ("prorab", "kadrovik", "admin"):
        raise HTTPException(400, "Неверная роль")
    u.role = UserRole(role)
    u.status = UserStatus.APPROVED
    u.approved_at = datetime.utcnow()
    db.commit()
    return RedirectResponse("/admin/users", status_code=303)


@app.post("/admin/users/{user_id}/block")
def admin_block(user_id: str, request: Request, db: Session = Depends(get_db)):
    admin = _current_user(request, db)
    if admin is None or admin.role != UserRole.ADMIN:
        raise HTTPException(403, "Только администратор")
    u = db.get(User, user_id)
    if u is None:
        raise HTTPException(404, "Пользователь не найден")
    if u.id == admin.id:
        raise HTTPException(400, "Нельзя заблокировать себя")
    u.status = UserStatus.BLOCKED
    db.commit()
    return RedirectResponse("/admin/users", status_code=303)


@app.post("/admin/users/{user_id}/unblock")
def admin_unblock(user_id: str, request: Request, db: Session = Depends(get_db)):
    admin = _current_user(request, db)
    if admin is None or admin.role != UserRole.ADMIN:
        raise HTTPException(403, "Только администратор")
    u = db.get(User, user_id)
    if u is None:
        raise HTTPException(404, "Пользователь не найден")
    u.status = UserStatus.APPROVED
    u.failed_attempts = 0
    u.locked_until = None
    db.commit()
    return RedirectResponse("/admin/users", status_code=303)


@app.post("/admin/users/{user_id}/unlock")
def admin_unlock(user_id: str, request: Request, db: Session = Depends(get_db)):
    """Снять временный замок за попытки подбора (не меняя статус)."""
    admin = _current_user(request, db)
    if admin is None or admin.role != UserRole.ADMIN:
        raise HTTPException(403, "Только администратор")
    u = db.get(User, user_id)
    if u is None:
        raise HTTPException(404, "Пользователь не найден")
    u.failed_attempts = 0
    u.locked_until = None
    db.commit()
    return RedirectResponse("/admin/users", status_code=303)


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


# --- Рабочий стол ------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _current_user_obj = _current_user(request, db)

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
        if e.registration_status is None:
            badges.append('<span class="badge red">статус учёта не задан</span>')
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
    # Из тех, кому нужно направление — у кого прошло >=5 дней от въезда ("пора выдавать").
    _today_dash = date.today()
    urge_referral = 0
    for _o in need_referral:
        _emp_u = db.get(Employee, _o.employee_id)
        if _emp_u and _emp_u.entry_date and (_today_dash - _emp_u.entry_date).days >= 5:
            urge_referral += 1

    def attention_row(e: Employee, badges: list[str]) -> str:
        return (
            f'<div class="card">{e.full_name}<br>'
            f'{"".join(badges)}<br>'
            f'<a class="btn" href="/employees/{e.id}">Открыть карточку</a></div>'
        )

    # Непроведённые обязательные инструктажи (вводный / первичный на рабочем месте):
    # дата начала работы наступила, а инструктажа нет. Отдельный раздел от "пробелов
    # в карточке" — это не незаполненное поле, а невыполненное действие по охране труда.
    instruction_gaps = prod.get_instruction_compliance_gaps(db)

    def instruction_gap_row(g: dict) -> str:
        badge_class = "red" if g["stage"] == "critical" else "orange"
        stage_word = "критично" if g["stage"] == "critical" else "просрочено"
        return (
            f'<div class="card">{html.escape(g["name"])} — {g["type_label"]}<br>'
            f'<span class="badge {badge_class}">{stage_word} · с {g["start_date"].strftime("%d.%m.%Y")} '
            f'({g["days_overdue"]} дн.)</span><br>'
            f'<a class="btn" href="/production/instructions">К инструктажам</a> '
            f'<a class="btn secondary" href="/employees/{g["employee_id"]}">Карточка</a></div>'
        )

    # Межвахта с открытыми обязательствами (2026-07, слияние с ботом ТабельБелокаменка).
    # Видит только кадровик/админ — прораб ставит межвахту не глядя на это (см.
    # UserRole.PRORAB в models.py и договорённость "данные направлены в отдел кадров").
    role = request.session.get("role", "")
    rotation_flags = []
    if role in ("kadrovik", "admin"):
        rotation_flags = db.scalars(
            select(RotationReturn)
            .where(RotationReturn.flagged == True)  # noqa: E712
            .where(RotationReturn.reviewed_at.is_(None))
        ).all()

    def rotation_row(rr: RotationReturn) -> str:
        emp = rr.employee
        open_obl = [
            o for o in (emp.obligations if emp else [])
            if o.is_current and o.status in (ObligationStatus.PENDING, ObligationStatus.OVERDUE)
        ]
        obl_labels = ", ".join(
            f"{OBLIGATION_LABELS.get(o.type, o.type.value)} ({o.status.value})" for o in open_obl
        ) or "—"
        return (
            f'<div class="card">{emp.full_name if emp else "?"} — межвахта до '
            f'{rr.expected_return_date.strftime("%d.%m.%Y") if rr.expected_return_date else "не уточнена"}<br>'
            f'<span class="badge red">Открытые обязательства: {obl_labels}</span><br>'
            f'<form method="post" action="/attention/rotation/{rr.employee_id}/resolve" style="display:inline">'
            f'<button type="submit" class="btn">✅ Разобрано</button></form> '
            f'<a class="btn" href="/employees/{rr.employee_id}">Открыть карточку</a></div>'
        )

    # На оформлении (2026-07): договор ещё не начался (нет даты или дата в будущем) —
    # такие сотрудники исключены из табеля (см. tabel.get_active_employees). Видит
    # только кадровик/админ, как и флаги межвахты.
    onboarding_employees = []
    if role in ("kadrovik", "admin"):
        onboarding_employees = tabel.get_onboarding_employees(db)

    def onboarding_row(e: Employee) -> str:
        cd = e.contract_date.strftime("%d.%m.%Y") if e.contract_date else "не указана"
        return (
            f'<div class="card">{e.full_name} — дата договора: {cd}<br>'
            f'<a class="btn" href="/employees/{e.id}">Открыть карточку</a></div>'
        )

    # Уточнить дату возврата с межвахты (2026-07): заглушки RotationReturn с
    # expected_return_date=NULL, см. tabel.get_pending_clarification_rotations.
    # Кадровик минимально работает в MAX — эта задача теперь в первую очередь
    # здесь, в веб-дашборде (прорабу отдельно шлётся напоминание в боте, см.
    # bot.py morning_job).
    pending_rotation = []
    if role in ("kadrovik", "admin"):
        pending_rotation = tabel.get_pending_clarification_rotations(db)

    def pending_rotation_row(item: dict) -> str:
        return (
            f'<div class="card">{item["name"]} — дата возврата с межвахты не уточнена<br>'
            f'<a class="btn" href="/employees/{item["employee_id"]}">Открыть карточку</a></div>'
        )

    # СРОЧНАЯ проверка (2026-07): явка есть, а договор не действует. См. подробный
    # docstring tabel.get_marks_without_valid_contract. Веб безопасен для проактивного
    # показа (вход персональный, не рассылка) — в отличие от бота, где это оставлено
    # только пассивным пунктом меню (см. решение про NotificationSubscriber).
    invalid_contract_marks = []
    if role in ("kadrovik", "admin"):
        invalid_contract_marks = tabel.get_marks_without_valid_contract(db)

    def invalid_contract_row(item: dict) -> str:
        cd = item["contract_date"].strftime("%d.%m.%Y") if item["contract_date"] else "не указана"
        ced = item["contract_end_date"]
        reason = f"уволен {ced:%d.%m.%Y}" if ced else f"дата договора: {cd}"
        marks_str = ", ".join(f"{d:%d.%m}" for d, _slot, _code in item["marks"])
        return (
            f'<div class="card">{item["name"]} ({reason})<br>'
            f'<span class="badge red">Явка: {marks_str}</span><br>'
            f'<a class="btn" href="/employees/{item["employee_id"]}">Открыть карточку</a></div>'
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
{f'''<div class="warning-banner" style="background:#fdecec;border-left-color:#c0392b;color:#7a1f1f">
🚨 СРОЧНО: явка без действующего договора у {len(invalid_contract_marks)} чел. — проверьте ниже.
</div>''' if invalid_contract_marks else ""}
<p class="muted" style="margin:0 0 10px">Рабочее место: Автоматизированная система учёта на производстве.
Вы вошли как {html.escape(_current_user_obj.full_name if _current_user_obj else "?")},
роль: {role or "—"}.</p>
<h1>Задачи</h1>
<p><a class="btn" href="/employees/new">+ Добавить сотрудника</a></p>

{f'''<section class="grid">
<h2>🚨 Явка без действующего договора ({len(invalid_contract_marks)})</h2>
{"".join(invalid_contract_row(item) for item in invalid_contract_marks)}
</section>''' if invalid_contract_marks else ""}

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

<section class="grid">
<h2>🦺 Инструктажи не проведены ({len(instruction_gaps)}{f", из них критично: {sum(1 for g in instruction_gaps if g['stage'] == 'critical')}" if any(g['stage'] == 'critical' for g in instruction_gaps) else ""})</h2>
{''.join(instruction_gap_row(g) for g in instruction_gaps) or '<p class="muted">Вводный и первичный проведены у всех, кто уже начал работу.</p>'}
</section>

{f'''<section class="grid">
<h2>🔄 Межвахта с открытыми обязательствами ({len(rotation_flags)})</h2>
{"".join(rotation_row(rr) for rr in rotation_flags) or '<p class="muted">Нет.</p>'}
</section>''' if role in ("kadrovik", "admin") else ""}

{f'''<section class="grid">
<h2>🧾 На оформлении, не в табеле ({len(onboarding_employees)})</h2>
{"".join(onboarding_row(e) for e in onboarding_employees) or '<p class="muted">Нет.</p>'}
</section>''' if role in ("kadrovik", "admin") else ""}

{f'''<section class="grid">
<h2>❓ Уточнить дату возврата с межвахты ({len(pending_rotation)})</h2>
{"".join(pending_rotation_row(item) for item in pending_rotation) or '<p class="muted">Нет.</p>'}
</section>''' if role in ("kadrovik", "admin") else ""}

<section>
<h2>Медкомиссия</h2>
<div class="card">Нужно направление: {len(need_referral)}<br>
{('<b style="color:#c47f00">Пора выдать (5+ дней от въезда): ' + str(urge_referral) + '</b><br>') if urge_referral else ''}Ждут результата: {len(awaiting_result)}<br>
<a class="btn" href="/medical">Открыть раздел</a></div>
</section>

{recompute_section}
"""
    return _render("Рабочий стол", body, active="home", role=request.session.get("role",""))


@app.post("/attention/rotation/{employee_id}/resolve")
def attention_rotation_resolve(employee_id: str, request: Request, db: Session = Depends(get_db)):
    """Кадровик отмечает флаг межвахты как разобранный — вручную, не снимается
    автоматически при закрытии обязательств (см. docstring RotationReturn в models.py)."""
    user = _require_role(request, db, UserRole.KADROVIK)
    rr = db.get(RotationReturn, employee_id)
    if rr is not None and rr.flagged:
        rr.reviewed_at = datetime.now(MSK)
        rr.reviewed_by = user.id
        db.add(rr)
        db.commit()
    return RedirectResponse("/", status_code=303)


# --- Табель (2026-07, слияние с ботом ТабельБелокаменка) --------------------
# Видят все три роли (PRORAB тоже — узкое исключение из "не пишет в БД", см.
# UserRole.PRORAB в models.py). Разметка через веб — та же tabel.py, что и бот,
# один источник правды (attendance_marks), просто два интерфейса ввода.

_TABEL_CODE_OPTIONS = [
    ("", "—"),
    (tabel.DAY, "Д — день"),
    ("С", "С — сутки"),
    (tabel.NIGHT, "НЧ — ночь"),
    (tabel.REST, "О — отдых"),
    (tabel.SICK, "Б — больничный"),
    (tabel.ROTATION, "МЖ — межвахта"),
    (tabel.ABSENT, "Н — неявка"),
    (tabel.MIGR, "МУ — мигр.учёт"),
    (tabel.WEEKEND, "В — выходной"),
]


@app.get("/tabel", response_class=HTMLResponse)
def tabel_page(request: Request, year: int | None = None, month: int | None = None,
                db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    today = datetime.now(MSK).date()
    year = year or today.year
    month = month or today.month
    days_in_month = calendar.monthrange(year, month)[1]

    s = tabel.day_summary(db)
    active_count = len(tabel.get_active_employees(db))
    grid = tabel.get_month_codes(db, year, month)

    # Навигация месяц назад/вперёд.
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    month_names = ["", "январь", "февраль", "март", "апрель", "май", "июнь",
                    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]

    _stat = lambda emoji, label, value: (
        f'<div style="flex:1 1 auto;min-width:70px;text-align:center;padding:8px 4px">'
        f'<div style="font-size:20px">{emoji}</div>'
        f'<div style="font-size:20px;font-weight:700">{value}</div>'
        f'<div style="font-size:11px;color:var(--sub)">{label}</div></div>'
    )
    summary_html = f"""
<div class="card" style="font-weight:400">
<div style="font-size:13px;color:var(--sub);margin-bottom:6px">Табель за {today.strftime("%d.%m.%Y")} — активных по списку: {active_count}</div>
<div style="display:flex;flex-wrap:wrap;gap:2px">
{_stat("☀️", "День", s['day'])}{_stat("🌙", "Ночь", s['night'])}{_stat("😴", "Отдых", s['rest'])}
{_stat("🤒", "Больн.", s['sick'])}{_stat("✈️", "Межвахта", s['rotation'])}{_stat("❌", "Неявка", s['absent'])}
{_stat("📋", "Мигр.учёт", s['migr'])}
</div>
</div>
"""
    if s["absent_list"]:
        _absent_badge = {
            tabel.ABSENT: "red", tabel.SICK: "orange", tabel.ROTATION: "neutral",
            tabel.MIGR: "neutral", tabel.WEEKEND: "green",
        }
        absent_cards = "".join(
            f'<div class="card" style="display:flex;justify-content:space-between;'
            f'align-items:center;gap:8px;font-weight:400">'
            f'<span>{html.escape(name)}</span>'
            f'<span class="badge {_absent_badge.get(code, "neutral")}" style="margin:0">{code}</span></div>'
            for name, code in s["absent_list"]
        )
        summary_html += f'<section><h2>Отсутствуют/особое ({len(s["absent_list"])})</h2>{absent_cards}</section>'

    # ✈️ На межвахте — с датой ожидаемого возврата (не просто бейдж в общем списке).
    rotation_rows = db.scalars(
        select(RotationReturn)
        .join(Employee, Employee.id == RotationReturn.employee_id)
        .where(Employee.contract_end_date.is_(None))
    ).all()
    if rotation_rows:
        _dep_label = {
            "abroad": "за границу", "domestic": "в РФ, с площадки", "none": "не выезжал",
        }
        rotation_cards = "".join(
            f'<div class="card" style="display:flex;justify-content:space-between;'
            f'align-items:center;gap:8px;font-weight:400">'
            f'<span>{html.escape(rr.employee.full_name if rr.employee else "?")}'
            f'<br><span class="muted">{_dep_label.get(rr.departure_type, "не указано")}</span></span>'
            + (
                f'<span class="badge neutral" style="margin:0">до {rr.expected_return_date:%d.%m.%Y}</span>'
                if rr.expected_return_date
                else '<span class="badge red" style="margin:0">дата не уточнена</span>'
            )
            + '</div>'
            for rr in rotation_rows
        )
        summary_html += f'<section><h2>✈️ На межвахте ({len(rotation_rows)})</h2>{rotation_cards}</section>'

    # Закреплённый (sticky) стиль для колонки номера — при горизонтальной прокрутке

    # Закреплённый (sticky) стиль для колонки номера — при горизонтальной прокрутке
    # ФИО и дни уходят влево вместе с остальной таблицей, а № остаётся на месте.
    _num_sticky = "position:sticky;left:0;background:#fff;z-index:2;"

    # Заголовок с числами месяца.
    header_cells = "".join(f'<th style="min-width:30px">{d}</th>' for d in range(1, days_in_month + 1))
    is_current_month = (year == today.year and month == today.month)

    rows_html = ""
    for row_num, (emp_id, data) in enumerate(
        sorted(grid.items(), key=lambda kv: kv[1]["name"]), start=1
    ):
        cells = ""
        for i, code in enumerate(data["codes"]):
            day_num = i + 1
            day_date = date(year, month, day_num)
            is_today_col = is_current_month and day_num == today.day
            bg = " background:#eaf0fb;" if is_today_col else ""
            code_bg = {
                tabel.DAY: "#eaf6f0", tabel.NIGHT: "#eef1f4", tabel.REST: "#eef1f4",
                tabel.SICK: "#fdf3e2", tabel.ROTATION: "#eef1f4", tabel.ABSENT: "#fdecec",
                tabel.MIGR: "#eef1f4", tabel.WEEKEND: "#eaf6f0", "С": "#eaf6f0",
            }.get(code, "#fff")
            options_html = "".join(
                f'<option value="{val}"{" selected" if val == code else ""}>{val or "—"}</option>'
                for val, _label in _TABEL_CODE_OPTIONS
            )
            cells += (
                f'<td style="padding:2px;{bg}">'
                f'<form method="post" action="/tabel/mark" class="inline">'
                f'<input type="hidden" name="employee_id" value="{emp_id}">'
                f'<input type="hidden" name="mark_date" value="{day_date.isoformat()}">'
                f'<input type="hidden" name="year" value="{year}">'
                f'<input type="hidden" name="month" value="{month}">'
                f'<select name="code" onchange="this.form.submit()" '
                f'style="min-height:32px;padding:2px;margin:0;font-size:12px;width:56px;'
                f'background-color:{code_bg}">'
                f'{options_html}</select></form></td>'
            )
        rows_html += (
            f'<tr>'
            f'<td style="{_num_sticky}padding:4px 8px;text-align:right;color:var(--sub)">{row_num}</td>'
            f'<td style="white-space:nowrap;font-weight:600;padding:4px 8px 4px 0">'
            f'{html.escape(data["name"])}</td>{cells}</tr>'
        )

    body = f"""
<h1>Табель</h1>
{summary_html}
<section style="overflow-x:auto">
<h2>{month_names[month]} {year}
&nbsp;
<a class="btn secondary" href="/tabel?year={prev_year}&month={prev_month}">← </a>
<a class="btn secondary" href="/tabel?year={next_year}&month={next_month}"> →</a>
</h2>
<table style="border-collapse:collapse;font-size:13px">
<thead><tr><th style="position:sticky;left:0;background:#fff;z-index:2">№</th><th></th>{header_cells}</tr></thead>
<tbody>{rows_html}</tbody>
</table>
<p class="muted">Изменение ячейки применяется сразу при выборе. "С" (сутки) и полноценная
постановка "МЖ" с указанием даты возврата и типа отбытия — доступны с полными проверками
через MAX-бота; здесь — быстрая правка кода без пошагового флоу.</p>
</section>
"""
    return _render("Табель", body, active="tabel", role=request.session.get("role", ""))


@app.post("/tabel/mark")
def tabel_mark(
    request: Request,
    employee_id: str = Form(...),
    mark_date: str = Form(...),
    code: str = Form(""),
    year: int = Form(...),
    month: int = Form(...),
    db: Session = Depends(get_db),
):
    """Запись одной ячейки табеля из веба. Роль не ограничивается (см. договорённость —
    все три роли могут ставить/менять отметки), но вход обязателен."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    actor = _actor_name(request, db)
    employee = db.get(Employee, employee_id)
    if employee is None:
        return RedirectResponse(f"/tabel?year={year}&month={month}", status_code=303)

    d = datetime.strptime(mark_date, "%Y-%m-%d").date()

    if code == "":
        tabel.clear_day_slot(db, employee, d)
        tabel.clear_night_slot(db, employee, d)
    elif code == tabel.DAY:
        tabel.mark_day(db, employee, actor, d)
    elif code == "С":
        tabel.mark_sutki(db, employee, actor, d)
    elif code == tabel.NIGHT:
        tabel.mark_night(db, employee, actor, d)
    elif code == tabel.REST:
        tabel.set_rest(db, employee, actor, d)
    elif code in (tabel.SICK, tabel.ROTATION, tabel.ABSENT, tabel.MIGR, tabel.WEEKEND):
        # МЖ через веб — без даты возврата/типа отбытия (упрощённая правка кода,
        # см. предупреждение на странице). Полноценная постановка — через бота.
        tabel.set_reason(db, employee, code, actor, d)

    return RedirectResponse(f"/tabel?year={year}&month={month}", status_code=303)


# --- Отчёты (2026-07) --------------------------------------------------------
# Данные — в reports.py (общие с ботом, см. bot.py "📊 Отчёты"). Здесь только
# HTML-рендер. Чтобы добавить отчёт: новая запись в reports.REPORTS_REGISTRY +
# функция данных в reports.py + маршрут здесь (полный HTML) + опционально
# урезанный текстовый рендер в bot.py.

import reports as reports_data


@app.get("/reports", response_class=HTMLResponse)
def reports_page(request: Request):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    cards = "".join(
        f'<div class="card"><h3 style="margin:0 0 6px">{title}</h3>'
        f'<p class="muted" style="margin:0 0 10px">{desc}</p>'
        f'<a class="btn" href="{href}">Открыть</a></div>'
        for _key, href, title, desc in reports_data.REPORTS_REGISTRY
    )
    body = f"""
<h1>Отчёты</h1>
<section class="grid">{cards}</section>
"""
    return _render("Отчёты", body, active="reports", role=request.session.get("role", ""))


@app.get("/reports/changelog", response_class=HTMLResponse)
def report_changelog(request: Request):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    def section(title: str, entries: list[tuple[str, str, str]]) -> str:
        rows = "".join(
            f'<div class="card"><h3 style="margin:0 0 6px">{h}</h3>'
            f'<p style="margin:0 0 6px"><b>Было:</b> {problem}</p>'
            f'<p style="margin:0"><b>Исправлено:</b> {fix}</p></div>'
            for h, problem, fix in entries
        )
        return f'<section class="grid"><h2>{title}</h2>{rows}</section>'

    body = f"""
<h1>🐛 Журнал ошибок и патчей</h1>
<p class="muted">Собрано по ходу разработки табеля (ТабельБелокаменка) и слияния с ботом
миграционного учёта. Не автоматический отчёт — фиксирует находки вручную по мере работы.</p>
{section(f"ТабельБелокаменка ({len(reports_data.CHANGELOG_TABEL)})", reports_data.CHANGELOG_TABEL)}
{section(f"Слияние с миграционным учётом ({len(reports_data.CHANGELOG_MIGBOT)})", reports_data.CHANGELOG_MIGBOT)}
<p><a class="btn secondary" href="/reports">← Все отчёты</a></p>
"""
    return _render("Журнал патчей", body, active="reports", role=request.session.get("role", ""))


@app.get("/reports/monthly-problems", response_class=HTMLResponse)
def report_monthly_problems(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    data = reports_data.get_monthly_problems_report(db)
    rows = "".join(
        f'<div class="card">{p["name"]}<br>'
        + (f'<span class="badge red" style="margin-right:6px">неявок: {p["absent_count"]}</span>'
           if p["absent_count"] >= data["absent_threshold"] else "")
        + (f'<span class="badge orange">выходных: {p["weekend_count"]}</span>'
           if p["weekend_count"] >= data["weekend_threshold"] else "")
        + '</div>'
        for p in data["problems"]
    )
    body = f"""
<h1>📊 Проблемные за месяц</h1>
<p class="muted">{data["month_label"]} — пороги: неявки от {data["absent_threshold"]},
выходные от {data["weekend_threshold"]}. Учитываются только активные сотрудники (действующий договор).</p>
<section class="grid">
<h2>Всего: {len(data["problems"])}</h2>
{rows or '<p class="muted">Никто не превысил пороги.</p>'}
</section>
<p><a class="btn secondary" href="/reports">← Все отчёты</a></p>
"""
    return _render("Проблемные за месяц", body, active="reports", role=request.session.get("role", ""))


@app.get("/reports/obligations", response_class=HTMLResponse)
def report_obligations(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    data = reports_data.get_obligations_report(db)
    summary_rows = "".join(
        f'<div class="card" style="display:flex;justify-content:space-between;align-items:center">'
        f'<span>{OBLIGATION_LABELS.get(_t, _t)} — {_s}</span>'
        f'<span class="badge {"red" if _s == "overdue" else ("green" if _s == "done" else "neutral")}">{_c}</span>'
        f'</div>'
        for (_t, _s), _c in sorted(data["counts"].items())
    )
    overdue_rows = "".join(
        f'<div class="card">{item["name"]} — {OBLIGATION_LABELS.get(item["type_value"], item["type_value"])}<br>'
        f'<span class="badge red">Дедлайн был {item["deadline_date"]:%d.%m.%Y}</span><br>'
        f'<a class="btn" href="/employees/{item["employee_id"]}">Открыть карточку</a></div>'
        for item in data["overdue_details"]
    )
    body = f"""
<h1>📋 Обязательства — сводка по статусам</h1>
<p class="muted">Только активные сотрудники (действующий договор). Только актуальные версии обязательств (is_current).</p>
<section class="grid"><h2>По типам и статусам</h2>{summary_rows or '<p class="muted">Нет данных.</p>'}</section>
<section class="grid"><h2>🚨 Просроченные, с деталями ({len(data["overdue_details"])})</h2>
{overdue_rows or '<p class="muted">Просроченных нет.</p>'}</section>
<p><a class="btn secondary" href="/reports">← Все отчёты</a></p>
"""
    return _render("Обязательства", body, active="reports", role=request.session.get("role", ""))


@app.get("/reports/activity", response_class=HTMLResponse)
def report_activity(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    data = reports_data.get_activity_report(db)
    rows = "".join(
        f'<div class="card">{a["label"]}<br>'
        f'<span class="badge neutral">{a["count"]} отметок</span> '
        f'<span class="muted">{a["first"]:%d.%m}–{a["last"]:%d.%m}</span></div>'
        for a in data["actors"]
    )
    body = f"""
<h1>🕵️ Активность в табеле</h1>
<p class="muted">{data["month_start"].strftime("%B %Y")} — кто и сколько отметок поставил
(включая перенесённую историю из Google Sheets).</p>
<section class="grid"><h2>Всего отметок за месяц: {data["total"]}</h2>{rows or '<p class="muted">Нет данных.</p>'}</section>
<p><a class="btn secondary" href="/reports">← Все отчёты</a></p>
"""
    return _render("Активность", body, active="reports", role=request.session.get("role", ""))


# --- Производство (2026-07) --------------------------------------------------
# Наряды-допуски, инструктажи, удостоверения — отдельный модуль (production.py),
# минимально связанный с остальной системой. См. docstring в production.py.

import production as prod


@app.get("/production", response_class=HTMLResponse)
def production_page(request: Request):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    body = """
<h1>Производство</h1>
<section class="grid">
<div class="card"><h3 style="margin:0 0 6px">📋 Наряды-допуски</h3>
<p class="muted" style="margin:0 0 10px">Официальные документы на производство работ — ответственный, бригада, срок действия.</p>
<a class="btn" href="/production/work-orders">Открыть</a></div>
<div class="card"><h3 style="margin:0 0 6px">🎓 Инструктажи</h3>
<p class="muted" style="margin:0 0 10px">Вводный, первичный, повторный, внеплановый, целевой — учёт по каждому сотруднику.</p>
<a class="btn" href="/production/instructions">Открыть</a></div>
<div class="card"><h3 style="margin:0 0 6px">🪪 Удостоверения</h3>
<p class="muted" style="margin:0 0 10px">Удостоверения по профессиям — сроки действия, сканы, напоминания об истечении.</p>
<a class="btn" href="/production/certificates">Открыть</a></div>
<div class="card"><h3 style="margin:0 0 6px">📑 Приказы</h3>
<p class="muted" style="margin:0 0 10px">Реестр внутренних приказов — сканы, номера, используется в футерах печатных бланков.</p>
<a class="btn" href="/production/orders">Открыть</a></div>
<div class="card"><h3 style="margin:0 0 6px">👷 Бригады</h3>
<p class="muted" style="margin:0 0 10px">Сохранённые составы — выбор бригадой целиком при создании наряда, без ручной отметки каждого.</p>
<a class="btn" href="/production/brigades">Открыть</a></div>
<div class="card"><h3 style="margin:0 0 6px">🏗 Титулы</h3>
<p class="muted" style="margin:0 0 10px">Справочник объектов (шифр + наименование) — заполняет «Место выполнения работ» при создании наряда, без ручного набора.</p>
<a class="btn" href="/production/tituly">Открыть</a></div>
<div class="card"><h3 style="margin:0 0 6px">📓 Журналы</h3>
<p class="muted" style="margin:0 0 10px">Журнал инструктажей, журнал учёта нарядов, общий журнал работ — учёт и выгрузка для печати/показа.</p>
<a class="btn" href="/production/journals">Открыть</a></div>
</section>
"""
    return _render("Производство", body, active="production", role=request.session.get("role", ""))


# --- Журналы ------------------------------------------------------------------

@app.get("/production/journals", response_class=HTMLResponse)
def journals_hub(request: Request):
    """Хаб журналов. Показываем только рабочее (инструктажи, наряды, ОЖР);
    журналы лесов и трёхступенчатого контроля появятся по мере готовности."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    body = """
<h1>📓 Журналы</h1>
<section class="grid">
<div class="card"><h3 style="margin:0 0 6px">🎓 Журнал инструктажей</h3>
<p class="muted" style="margin:0 0 10px">Вводный, первичный, повторный, целевой — допечатка партиями, отдельная нумерация по каждому виду.</p>
<a class="btn" href="/production/instructions">Открыть</a></div>
<div class="card"><h3 style="margin:0 0 6px">📋 Журнал учёта нарядов</h3>
<p class="muted" style="margin:0 0 10px">Учёт работ по наряду-допуску (Приложение №5 к 782н) — по данным выпущенных нарядов, выгрузка в xlsx.</p>
<a class="btn" href="/production/journals/work-orders">Открыть</a></div>
<div class="card"><h3 style="margin:0 0 6px">🏗 Общий журнал работ</h3>
<p class="muted" style="margin:0 0 10px">Ежедневные записи по наряду — что сделано, погода. Подпись УКЭП (КриптоПро) подключается отдельно.</p>
<a class="btn" href="/production/journals/work-log">Открыть</a></div>
</section>
<p><a class="btn secondary" href="/production">← Производство</a></p>
"""
    return _render("Журналы", body, active="production", role=request.session.get("role", ""))


# --- Журнал учёта нарядов -----------------------------------------------------

@app.get("/production/journals/work-orders", response_class=HTMLResponse)
def wo_journal_page(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    unprinted = len(prod.get_unprinted_work_orders(db))
    last_num = prod.get_last_wo_journal_row_number(db)
    body = f"""
<h1>📋 Журнал учёта нарядов</h1>
<p class="muted">Учёт работ по наряду-допуску (Приложение №5 к 782н). Строится по выпущенным нарядам.
Допечатка партиями: печатаются только новые наряды, ещё не попавшие в журнал, со сквозной нумерацией.</p>
<section class="grid">
<div class="card">
<span class="badge {'orange' if unprinted else 'neutral'}">Ждут внесения в журнал: {unprinted}</span><br>
<span class="muted">Последний номер строки в журнале: {last_num or '—'}</span><br>
<form method="post" action="/production/journals/work-orders/print" style="margin-top:8px">
<button type="submit" class="btn"{' disabled' if not unprinted else ''}>Допечатать новые записи ({unprinted})</button></form>
</div>
</section>
<p><a class="btn secondary" href="/production/journals">← Журналы</a></p>
"""
    return _render("Журнал нарядов", body, active="production", role=request.session.get("role", ""))


@app.post("/production/journals/work-orders/print")
def wo_journal_print(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    printed = prod.print_new_wo_journal_entries(db)
    if not printed:
        return RedirectResponse("/production/journals/work-orders", status_code=303)
    path = prod.generate_wo_journal_xlsx(printed, org_name=ORG_NAME)
    with open(path, "rb") as f:
        data = f.read()
    fn = f"Журнал_нарядов_{datetime.now(MSK):%d.%m.%Y}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": _content_disposition(fn)},
    )


# --- Общий журнал работ (ОЖР) -------------------------------------------------

# JS тестового прогона КриптоПро — вынесен из f-string (фигурные скобки JS иначе ломают
# подстановку). Подключается cadesplugin_api.js; кнопка обращается к плагину, показывает
# сертификаты на токене и НИЧЕГО не подписывает. Нужен файл cadesplugin_api.js в статике
# (из дистрибутива «КриптоПро ЭЦП Browser plug-in»); без него кнопка честно скажет
# «плагин не обнаружен».
_CP_TEST_JS = """
<script src="/cadesplugin_api.js"></script>
<script>
function _cpShow(t){var el=document.getElementById('cpResult');el.style.display='block';el.textContent=t;}
function _cpTest(){
  _cpShow('Обращаюсь к плагину…');
  if (typeof cadesplugin === 'undefined'){
    _cpShow('❌ Плагин не обнаружен.\\nПроверьте: установлен «КриптоПро ЭЦП Browser plug-in», включено расширение браузера, и подключён файл cadesplugin_api.js.');
    return;
  }
  cadesplugin.then(function(){
    cadesplugin.async_spawn(function*(){
      try{
        var oStore = yield cadesplugin.CreateObjectAsync("CAdESCOM.Store");
        yield oStore.Open(2, "My", 0);
        var certs = yield oStore.Certificates;
        var count = yield certs.Count;
        if (!count){ _cpShow('⚠ Плагин работает, но сертификатов не найдено. Вставьте токен или установите сертификат.'); yield oStore.Close(); return; }
        var lines = ['✅ Плагин и ключ доступны. Сертификатов найдено: ' + count, ''];
        for (var i=1; i<=count; i++){
          var c = yield certs.Item(i);
          var subj = yield c.SubjectName;
          lines.push('• ' + subj);
        }
        _cpShow(lines.join('\\n'));
        yield oStore.Close();
      }catch(e){ _cpShow('❌ Ошибка обращения к хранилищу: ' + e); }
    });
  }, function(err){ _cpShow('❌ Плагин не инициализирован: ' + err); });
}
</script>
"""


@app.get("/production/journals/work-log", response_class=HTMLResponse)
def work_log_page(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    entries = prod.get_work_log_entries(db)
    orders = db.scalars(select(WorkOrder).order_by(WorkOrder.created_at.desc())).all()
    wo_options = "".join(
        f'<option value="{o.id}">№{o.number} — {html.escape(o.work_description)} ({o.location})</option>'
        for o in orders
    )
    unprinted = len(prod.get_unprinted_work_log(db))

    rows = ""
    for e in entries:
        wo = e.work_order
        place = f"№{wo.number} · {html.escape(wo.location)} — {html.escape(wo.work_description)}" if wo else "—"
        signed = e.sign_status == WorkLogSignStatus.SIGNED
        sign_badge = (
            f'<span class="badge neutral">Подписано ({html.escape(e.signed_by or "")})</span>'
            if signed else '<span class="badge orange">Черновик — не подписано</span>'
        )
        # ТЕСТОВЫЙ набор кнопок. На проде: убрать «Редактировать» и «Снять подпись»,
        # оставить только постановку подписи (необратимо) — см. пометки в production.py.
        if signed:
            btns = (
                f'<form method="post" action="/production/journals/work-log/{e.id}/unsign" style="display:inline"'
                f' onsubmit="return confirm(\'Снять подпись (тест)? Запись вернётся в черновик.\')">'
                f'<button type="submit" class="btn secondary">↩ Снять подпись (тест)</button></form>'
            )
        else:
            btns = (
                f'<a class="btn secondary" href="/production/journals/work-log/{e.id}/edit">✏ Редактировать</a> '
                f'<form method="post" action="/production/journals/work-log/{e.id}/sign-test" style="display:inline">'
                f'<button type="submit" class="btn">🖊 Подписать (тест)</button></form> '
                f'<form method="post" action="/production/journals/work-log/{e.id}/delete" style="display:inline"'
                f' onsubmit="return confirm(\'Удалить запись журнала?\')">'
                f'<button type="submit" class="btn secondary">🗑 Удалить</button></form>'
            )
        rows += (
            f'<div class="card"><b>{e.entry_date:%d.%m.%Y}</b> · {place}<br>'
            f'<span>{html.escape(e.work_done)}</span><br>'
            f'{f"<span class=\"muted\">Погода: {html.escape(e.weather)}</span><br>" if e.weather else ""}'
            f'{f"<span class=\"muted\">{html.escape(e.note)}</span><br>" if e.note else ""}'
            f'{sign_badge} {btns}</div>'
        )

    body = f"""
<h1>🏗 Общий журнал работ</h1>
<p class="muted">Внутренний электронный журнал: одна запись — один день работ по наряду. Место, работа и
состав берутся из наряда. Погодные условия важны для зимнего бетонирования. Подпись УКЭП (КриптоПро)
подключается отдельно — пока записи в статусе «черновик».</p>
<details style="margin:8px 0"><summary style="cursor:pointer;color:#555">ℹ️ О законности электронного ведения</summary>
<div class="muted" style="margin-top:6px;font-size:0.92em;line-height:1.5">
<p>Это <b>внутренний</b> журнал ИП для собственного учёта и как основание для актов/отчётов. Вести его
электронно вы вправе свободно — это ваш документ, он не обязан соответствовать государственному формату.</p>
<p>Он <b>не заменяет</b> официальный Общий журнал работ по Градостроительному кодексу (форма и порядок —
приказ Минстроя № 1026/пр, электронное ведение — № 344/пр, действуют с 01.09.2023). Официальный ОЖР ведёт
<b>генподрядчик</b>, ответственный за строительство объекта, а не субподрядчик; его электронная форма требует
XML-схемы Минстроя, УКЭП на каждой записи и выгрузки в нередактируемые файлы.</p>
<p>УКЭП в этом журнале — <b>ваше решение</b> для юридической значимости записей перед заказчиком/генподрядчиком,
а не обязательное требование закона к внутреннему журналу.</p>
</div></details>

<section class="grid">
<h2>Записи журнала ({len(entries)})</h2>
{rows or '<p class="muted">Записей пока нет.</p>'}
</section>

<section>
<h2>Печать журнала</h2>
<p class="muted">Ждут внесения в журнал: {unprinted}. Допечатка партиями со сквозной нумерацией.</p>
<form method="post" action="/production/journals/work-log/print">
<button type="submit"{' disabled' if not unprinted else ''}>Допечатать новые записи ({unprinted})</button></form>
</section>

<section>
<h2>Новая запись</h2>
<form method="get" action="/production/journals/work-log/prepare">
<label>Наряд-допуск: <select name="work_order_id" required>{wo_options}</select></label>
<label>Дата: <input type="date" name="entry_date" required></label>
<button type="submit">Подготовить запись →</button>
</form>
<p class="muted">Операции, инструмент и погода подставятся из наряда автоматически — останется
подтвердить отмеченные операции и сохранить.</p>
</section>
<section>
<h2>Проверка ключа (КриптоПро)</h2>
<p class="muted">Тестовый прогон подписи: вставьте токен и нажмите. Браузер обратится к
«КриптоПро ЭЦП Browser plug-in» и покажет найденные сертификаты. Ничего не подписывает
и на сервер не отправляет — только проверка, что связка браузер↔плагин↔ключ работает.</p>
<button type="button" class="btn secondary" onclick="_cpTest()">🔑 Проверить ключ</button>
<pre id="cpResult" style="white-space:pre-wrap;margin-top:8px;font-size:13px;color:#333;background:#f7f8fa;padding:8px;border-radius:8px;display:none"></pre>
</section>
<p><a class="btn secondary" href="/production/journals">← Журналы</a></p>
"""
    return _render("Общий журнал работ", body + _CP_TEST_JS, active="production", role=request.session.get("role", ""))



@app.get("/production/journals/work-log/prepare", response_class=HTMLResponse)
def work_log_prepare(
    request: Request,
    work_order_id: str,
    entry_date: str,
    db: Session = Depends(get_db),
):
    """Экран подготовки записи ОЖР: операции вида работ подставлены чекбоксами (уже сделанные
    по этому наряду помечены, следующие невыполненные предложены галочками), инструмент и
    место — из наряда. Если это бетонирование, а по объекту нет записей этапа 1 — мягкое
    предупреждение СП 435. Человек подтверждает операции и сохраняет — «выполнено за день»
    собирается на сервере."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    wo = db.get(WorkOrder, work_order_id)
    if wo is None:
        raise HTTPException(404, "Наряд не найден")
    from datetime import date as _date
    try:
        d = _date.fromisoformat(entry_date)
    except ValueError:
        raise HTTPException(400, "Неверная дата.")

    wt = wo.work_type
    ops = prod.parse_operations(wt.content if wt else None)
    done_before = prod.get_done_operations_for_order(db, work_order_id)

    # Мягкое предупреждение СП 435: бетонирование (этап 2) без записей этапа 1 по объекту.
    warn = ""
    if wt and wt.stage_order == 2 and not prod.stage1_documented_for_object(db, wo.titul_id):
        warn = (
            '<div style="margin:8px 0;padding:10px;border:1px solid #e0a800;border-radius:8px;'
            'background:#fff8e1;color:#7a5c00">⚠ По этому объекту ещё нет записей журнала об '
            'опалубке/армировании (этап 1). По СП 435.1325800.2018 бетонирование ведётся после '
            'приёмки опалубки и арматуры. Проверьте последовательность — запись можно сохранить.</div>'
        )

    checks = ""
    for op in ops:
        already = op in done_before
        checked = "" if already else " checked"  # невыполненные предложены, выполненные сняты
        suffix = ' <span class="muted">(отмечено ранее)</span>' if already else ""
        op_esc = html.escape(op)
        checks += (
            f'<label style="display:block;margin:3px 0">'
            f'<input type="checkbox" name="ops" value="{op_esc}"{checked}> {op_esc}{suffix}</label>'
        )

    tools = (wt.tools if wt else "") or ""
    place = f"№{wo.number} · {html.escape(wo.location)}"
    vid = html.escape(wt.name) if wt else "— (наряд без вида работ)"

    body = f"""
<h1>🏗 Запись ОЖР — подтверждение</h1>
{warn}
<p class="muted">Наряд: {place}<br>Дата: {d:%d.%m.%Y} · Вид работ: {vid}</p>
<form method="post" action="/production/journals/work-log/new">
<input type="hidden" name="work_order_id" value="{work_order_id}">
<input type="hidden" name="entry_date" value="{entry_date}">
<section>
<h3>Операции за день</h3>
<p class="muted">Отмечены операции, ещё не выполненные по этому наряду. Снимите лишние или добавьте нужные.</p>
{checks or '<p class="muted">У вида работ нет списка операций — «выполнено за день» останется по названию вида.</p>'}
</section>
<section>
<h3>Инструмент (из вида работ)</h3>
<p class="muted">{html.escape(tools) or '—'}</p>
</section>
<label>Погода: <input type="text" name="weather" placeholder="пусто = подтянуть автоматически по Белокаменке"></label>
<label>Примечания: <input type="text" name="note" placeholder="необязательно"></label>
<button type="submit">Сохранить запись</button>
</form>
<p><a class="btn secondary" href="/production/journals/work-log">← Отмена</a></p>
"""
    return _render("Запись ОЖР", body, active="production", role=request.session.get("role", ""))


@app.post("/production/journals/work-log/new")
def work_log_create(
    request: Request,
    work_order_id: str = Form(...),
    entry_date: str = Form(...),
    ops: list[str] = Form(default=[]),
    weather: str = Form(""),
    note: str = Form(""),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    from datetime import date as _date
    try:
        d = _date.fromisoformat(entry_date)
    except ValueError:
        raise HTTPException(400, "Неверная дата.")
    actor = _actor_name(request, db)
    wo = db.get(WorkOrder, work_order_id)
    wt = wo.work_type if wo else None
    # «Выполнено за день» собирается из вида работ + отмеченных операций + инструмента.
    work_done = prod.build_work_done_text(wt, ops, tools=(wt.tools if wt else None))
    if not work_done.strip():
        work_done = (wt.name if wt else "Работы по наряду")
    # Погода: ручной ввод в приоритете; пусто — тянем из Open-Meteo по Белокаменке.
    weather_val = weather.strip()
    if not weather_val:
        weather_val = prod.fetch_weather_belokamenka(d) or ""
    prod.create_work_log_entry(
        db, work_order_id, d, work_done,
        weather=weather_val or None, note=note.strip() or None, created_by=actor,
        done_operations=json.dumps(ops, ensure_ascii=False) if ops else None,
    )
    return RedirectResponse("/production/journals/work-log", status_code=303)


@app.post("/production/journals/work-log/{entry_id}/delete")
def work_log_delete(entry_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.delete_work_log_entry(db, entry_id)
    return RedirectResponse("/production/journals/work-log", status_code=303)


# --- ТЕСТОВЫЙ блок работы с подписью ОЖР (на проде убрать edit и unsign) -------

@app.get("/production/journals/work-log/{entry_id}/edit", response_class=HTMLResponse)
def work_log_edit_form(entry_id: str, request: Request, db: Session = Depends(get_db)):
    """ТЕСТ: форма редактирования черновика ОЖР. Подписанные не редактируются."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    e = db.get(WorkLogEntry, entry_id)
    if e is None:
        raise HTTPException(404, "Запись не найдена")
    if e.sign_status == WorkLogSignStatus.SIGNED:
        return RedirectResponse("/production/journals/work-log", status_code=303)
    wo = e.work_order
    place = f"№{wo.number} · {html.escape(wo.location)}" if wo else "—"
    body = f"""
<h1>✏ Редактирование записи ОЖР</h1>
<p class="muted">Наряд: {place}. Тестовый режим — на проде подписанные записи не редактируются.</p>
<form method="post" action="/production/journals/work-log/{e.id}/edit">
<label>Дата: <input type="date" name="entry_date" value="{e.entry_date.isoformat()}" required></label>
<label>Выполнено за день:<br><textarea name="work_done" rows="3" required style="width:100%">{html.escape(e.work_done)}</textarea></label>
<input type="text" name="weather" value="{html.escape(e.weather or '')}" placeholder="Погода">
<input type="text" name="note" value="{html.escape(e.note or '')}" placeholder="Примечания">
<button type="submit">Сохранить</button>
</form>
<p><a class="btn secondary" href="/production/journals/work-log">← Назад к журналу</a></p>
"""
    return _render("Редактирование ОЖР", body, active="production", role=request.session.get("role", ""))


@app.post("/production/journals/work-log/{entry_id}/edit")
def work_log_edit_submit(
    entry_id: str,
    request: Request,
    entry_date: str = Form(...),
    work_done: str = Form(...),
    weather: str = Form(""),
    note: str = Form(""),
    db: Session = Depends(get_db),
):
    """ТЕСТ: сохранение отредактированного черновика."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    from datetime import date as _date
    try:
        d = _date.fromisoformat(entry_date)
    except ValueError:
        raise HTTPException(400, "Неверная дата.")
    prod.update_work_log_entry(
        db, entry_id, entry_date=d, work_done=work_done.strip(),
        weather=weather.strip() or None, note=note.strip() or None,
    )
    return RedirectResponse("/production/journals/work-log", status_code=303)


@app.post("/production/journals/work-log/{entry_id}/sign-test")
def work_log_sign_test(entry_id: str, request: Request, db: Session = Depends(get_db)):
    """ТЕСТ-подпись: помечает запись подписанной БЕЗ реальной криптографии. Серийник = «ТЕСТ»,
    хеш = sha256 содержимого. Реальная подпись придёт с клиента (КриптоПро) в ту же
    prod.sign_work_log_entry, когда подключим плагин — здесь только имитация потока."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    e = db.get(WorkLogEntry, entry_id)
    if e is None:
        raise HTTPException(404, "Запись не найдена")
    import hashlib
    payload = f"{e.entry_date.isoformat()}|{e.work_done}|{e.weather or ''}|{e.note or ''}"
    content_hash = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    prod.sign_work_log_entry(
        db, entry_id, signed_by=_actor_name(request, db),
        cert_serial="ТЕСТ", content_hash=content_hash,
    )
    return RedirectResponse("/production/journals/work-log", status_code=303)


@app.post("/production/journals/work-log/{entry_id}/unsign")
def work_log_unsign(entry_id: str, request: Request, db: Session = Depends(get_db)):
    """ТЕСТ: снять подпись, вернуть в черновик. На проде убрать — подпись необратима."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.unsign_work_log_entry(db, entry_id)
    return RedirectResponse("/production/journals/work-log", status_code=303)


@app.post("/production/journals/work-log/print")
def work_log_print(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    printed = prod.print_new_work_log_entries(db)
    if not printed:
        return RedirectResponse("/production/journals/work-log", status_code=303)
    path = prod.generate_work_log_xlsx(printed, org_name=ORG_NAME)
    with open(path, "rb") as f:
        data = f.read()
    fn = f"Общий_журнал_работ_{datetime.now(MSK):%d.%m.%Y}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": _content_disposition(fn)},
    )


# --- Наряды-допуски -----------------------------------------------------------

@app.get("/production/work-orders", response_class=HTMLResponse)
def work_orders_page(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    orders = prod.get_active_work_orders(db)
    # Автонумерация наряда: формат «{порядковый}-{год}» со сбросом по годам.
    # Берём максимальный порядковый среди ВСЕХ нарядов текущего года (включая
    # закрытые — чтобы не переиспользовать номер), +1. Старые номера без «-ГГГГ»
    # под шаблон не подпадают и игнорируются. Поле редактируемое — можно заменить.
    import re as _re_num
    _today = datetime.now(MSK).date()
    _year = _today.year
    _all_numbers = db.scalars(select(WorkOrder.number)).all()
    _seqs = [
        int(_m.group(1))
        for _n in _all_numbers
        if (_m := _re_num.match(rf"^(\d+)-{_year}$", (_n or "").strip()))
    ]
    _next_number = f"{(max(_seqs) + 1) if _seqs else 1}-{_year}"
    # Даты по умолчанию: «с» = сегодня, «по» = +14 дней (предел срока по 782н).
    _valid_from_default = _today.isoformat()
    _valid_to_default = (_today + timedelta(days=14)).isoformat()
    employees = db.scalars(
        select(Employee).where(Employee.contract_end_date.is_(None)).order_by(Employee.full_name)
    ).all()
    # Разведение по ролям 782н:
    #  - руководитель — только с 3-й группой по высоте (по объекту это Буц; список
    #    сам расширится, если кто-то ещё пройдёт обучение на 3-ю группу);
    #  - исполнитель и члены бригады — табельные (off_tabel=False), т.к. руководитель
    #    и исполнитель по Правилам не могут быть одним лицом.
    supervisors = [e for e in employees if e.height_safety_group and "3-я гр." in e.height_safety_group]
    workers = [e for e in employees if not e.off_tabel]
    # Если ни у кого нет 3-й группы — не оставляем выпадашку пустой (иначе наряд не
    # создать вообще); показываем всех и полагаемся на проверку check_work_order_problems.
    supervisor_pool = supervisors or employees
    # Повтор последнего наряда (по времени, статус не важен): ответственный исполнитель
    # и состав бригады по умолчанию берутся из самого свежего наряда. Не хардкод имени —
    # дефолт следует за реальностью: сменится исполнитель/состав, поедет и он. Если
    # человек уволился/стал off_tabel, его нет в workers → просто не отметится.
    _last_order = db.scalars(
        select(WorkOrder).order_by(WorkOrder.created_at.desc())
    ).first()
    _last_executor_id = _last_order.responsible_executor_id if _last_order else ""
    _last_member_ids = set()
    if _last_order:
        _last_member_ids = set(db.scalars(
            select(WorkOrderMember.employee_id).where(
                WorkOrderMember.work_order_id == _last_order.id
            )
        ).all())
    emp_options = "".join(f'<option value="{e.id}">{e.full_name}</option>' for e in supervisor_pool)
    worker_options = "".join(
        f'<option value="{e.id}"{" selected" if e.id == _last_executor_id else ""}>{e.full_name}</option>'
        for e in workers
    )
    emp_checkboxes = "".join(
        f'<label style="display:block;margin:2px 0"><input type="checkbox" name="member_ids" '
        f'id="memberCb_{e.id}" value="{e.id}"{" checked" if e.id in _last_member_ids else ""}> {e.full_name}</label>'
        for e in workers
    )
    brigades = prod.get_brigades(db)
    brigade_options = "".join(f'<option value="{b.id}">{b.name}</option>' for b in brigades)
    work_types = db.scalars(
        select(WorkType).where(WorkType.active.is_(True)).order_by(WorkType.name)
    ).all()
    work_type_options = "".join(
        f'<option value="{w.id}">{w.name}</option>' for w in work_types
    )
    # Титулы: value = готовая строка «код — наименование», ею наполняется поле location.
    tituly = prod.get_tituly(db)
    titul_options = "".join(
        f'<option value="{t.id}" data-text="{html.escape(f"{t.code} — {t.name}")}">'
        f'{html.escape(f"{t.code} — {t.name}")}</option>'
        for t in tituly
    )
    import json as _json
    brigades_js_map = _json.dumps({
        b.id: prod.get_brigade_member_ids(db, b.id) for b in brigades
    })
    # Словарь id->название типовой работы для автоподстановки в «На выполнение работ» (JS).
    # <  ->  \u003c, чтобы название с "<" не могло закрыть <script> раньше времени.
    work_types_js_map = _json.dumps(
        {w.id: w.name for w in work_types}, ensure_ascii=False
    ).replace("<", "\\u003c")

    def order_row(o) -> str:
        members = db.query(WorkOrderMember).filter_by(work_order_id=o.id).all()
        signed = sum(1 for m in members if m.signed_at is not None)
        member_ids_set = {m.employee_id for m in members}
        # Дозаполнение состава уже созданного наряда (узкое редактирование): чекбоксы
        # табельных сотрудников, текущие члены отмечены; сохранение заменяет состав.
        member_cbs = "".join(
            f'<label style="display:block;margin:2px 0"><input type="checkbox" name="member_ids" '
            f'value="{e.id}"{" checked" if e.id in member_ids_set else ""}> {html.escape(e.full_name)}</label>'
            for e in workers
        )
        edit_members = (
            f'<details style="margin-top:8px"><summary style="cursor:pointer">👷 Состав бригады '
            f'({len(members)})</summary>'
            f'<form method="post" action="/production/work-orders/{o.id}/members" style="margin-top:8px">'
            f'<fieldset><legend>Отметьте членов бригады</legend>{member_cbs}</fieldset>'
            f'<button type="submit">Сохранить состав</button></form></details>'
        )
        # --- Пункт 7: изменения состава действующего наряда (782н) ---
        stats = prod.work_order_change_stats(db, o.id)
        # Ввести можно тех, кого НЕТ в составе; вывести — тех, кто В составе.
        not_in = [e for e in workers if e.id not in member_ids_set]
        in_crew = [e for e in workers if e.id in member_ids_set]
        add_opts = "".join(f'<option value="{e.id}">{html.escape(e.full_name)}</option>' for e in not_in)
        rem_opts = "".join(f'<option value="{e.id}">{html.escape(e.full_name)}</option>' for e in in_crew)
        changes_hist = "".join(
            f'<li class="muted">{"➕ " + html.escape(getattr(ch.employee, "full_name", "?")) if ch.change_type == MemberChangeType.ADDED else "➖ " + html.escape(getattr(ch.employee, "full_name", "?"))} · '
            f'{ch.changed_at:%d.%m %H:%M} · разрешил: {html.escape(ch.ordered_by or "—")}</li>'
            for ch in prod.get_member_changes(db, o.id)
        )
        # Счётчик и блокировка по правилу половины.
        if stats["at_limit"]:
            limit_note = (
                f'<p class="badge red">Достигнут предел изменений состава '
                f'({stats["changes"]} из {stats["limit"]}). По 782н дальнейшее изменение '
                f'аннулирует наряд — требуется НОВЫЙ наряд-допуск.</p>'
            )
            change_forms = ""
        else:
            limit_note = (
                f'<p class="muted">Изменено: {stats["changes"]} из {stats["limit"]} допустимых '
                f'(первоначально {stats["initial"]} чел.). Свыше половины — новый наряд.</p>'
            )
            change_forms = (
                (f'<form method="post" action="/production/work-orders/{o.id}/member-change" style="margin-top:8px">'
                 f'<input type="hidden" name="change_type" value="added">'
                 f'<label>Ввести работника: <select name="employee_id" required>{add_opts}</select></label>'
                 f'<input type="text" name="ordered_by" placeholder="Кто дал указание (ФИО)" required>'
                 f'<button type="submit">➕ Ввести в состав</button></form>' if add_opts else '')
                +
                (f'<form method="post" action="/production/work-orders/{o.id}/member-change" style="margin-top:8px">'
                 f'<input type="hidden" name="change_type" value="removed">'
                 f'<label>Вывести работника: <select name="employee_id" required>{rem_opts}</select></label>'
                 f'<input type="text" name="ordered_by" placeholder="Кто дал указание (ФИО)" required>'
                 f'<button type="submit">➖ Вывести из состава</button></form>' if rem_opts else '')
            )
        p7 = (
            f'<details style="margin-top:8px"><summary style="cursor:pointer">📝 Пункт 7: изменения состава '
            f'({stats["changes"]})</summary>'
            f'{limit_note}'
            f'{("<ul>" + changes_hist + "</ul>") if changes_hist else ""}'
            f'{change_forms}</details>'
        )
        return (
            f'<div class="card">№{o.number} — {o.work_description}<br>'
            f'<span class="muted">{o.location} · {o.valid_from:%d.%m}–{o.valid_to:%d.%m.%Y}</span><br>'
            f'<span class="muted">Руководитель: {o.responsible_supervisor.full_name if o.responsible_supervisor else "?"} · '
            f'Исполнитель: {o.responsible_executor.full_name if o.responsible_executor else "?"}</span><br>'
            f'<span class="badge neutral">Подписали: {signed}/{len(members)}</span> '
            f'<a class="btn secondary" href="/production/work-orders/{o.id}/print">Печатный бланк</a> '
            f'<form method="post" action="/production/work-orders/{o.id}/close" style="display:inline">'
            f'<button type="submit" class="btn secondary">Закрыть наряд</button></form>'
            f'<form method="post" action="/production/work-orders/{o.id}/delete" style="display:inline"'
            f' onsubmit="return confirm(\'Удалить наряд №{o.number} безвозвратно? Вместе с ним удалятся состав, ежедневные допуски и изменения состава.\')">'
            f'<button type="submit" class="btn secondary">🗑 Удалить</button></form>'
            f'{edit_members}{p7}</div>'
        )

    rows = "".join(order_row(o) for o in orders)
    body = f"""
<h1>📋 Наряды-допуски</h1>
<section class="grid"><h2>Активные ({len(orders)})</h2>{rows or '<p class="muted">Нет активных нарядов.</p>'}</section>
<section>
<h2>Новый наряд</h2>
<form method="post" action="/production/work-orders/new">
<input type="text" name="number" placeholder="Номер наряда" value="{_next_number}" required>
<input type="text" name="subdivision" placeholder="Подразделение (например: ОС)">
<label>Типовая работа из справочника (необязательно — заполнит условия, ОВПФ, системы безопасности, раздел 3 и нормы):
<select name="work_type_id" onchange="_applyWorkType()"><option value="">— не выбрано —</option>{work_type_options}</select></label>
<textarea name="work_description" id="workDescription" placeholder="На выполнение работ" required rows="3"></textarea>
<label style="display:block;margin-top:6px">Титул из справочника (заполнит место выполнения работ; можно вписать свой):
<select id="titulSelect" onchange="_applyTitul()"><option value="">— вручную —</option>{titul_options}</select></label>
<input type="hidden" name="titul_id" id="titul_id_field">
<input type="text" id="location" name="location" placeholder="Место выполнения работ" required>
<label>Ответственный руководитель работ:
<select name="responsible_supervisor_id" required>{emp_options}</select></label>
<label>Ответственный исполнитель работ (бригадир):
<select name="responsible_executor_id" required>{worker_options}</select></label>
<label>Действует с: <input type="date" id="validFrom" name="valid_from" value="{_valid_from_default}" onchange="_applyValidTo()" required></label>
<label>Действует по: <input type="date" id="validTo" name="valid_to" value="{_valid_to_default}" required></label>
<label>Выбрать готовую бригаду (необязательно):
<select id="brigadeSelect" onchange="_applyBrigade()"><option value="">— вручную —</option>{brigade_options}</select></label>
<button type="button" class="secondary" onclick="_applyBrigade()">Применить состав</button>
<fieldset><legend>Члены бригады</legend>{emp_checkboxes}</fieldset>
<fieldset><legend>Дополнительно (необязательно)</legend>
<textarea name="materials" placeholder="Материалы" rows="2"></textarea>
<textarea name="tools" placeholder="Инструменты" rows="2"></textarea>
<textarea name="equipment" placeholder="Приспособления" rows="2"></textarea>
<textarea name="special_machinery" placeholder="Спецтехника" rows="2"></textarea>
<input type="text" name="technological_card_ref" placeholder="Ссылка на технологическую карту (шифр ТК)">
<textarea name="safety_systems" placeholder="Системы обеспечения безопасности (страховочные/поддерживающие/эвакуационные)" rows="2"></textarea>
<textarea name="special_conditions" placeholder="Особые условия (погодные ограничения и т.п.)" rows="2"></textarea>
</fieldset>
<button type="submit">Создать наряд</button>
</form>
</section>
<p><a class="btn secondary" href="/production">← Производство</a></p>
<script>
var _brigadesData = {brigades_js_map};
var _workTypesData = {work_types_js_map};
function _applyBrigade(){{
  var sel = document.getElementById('brigadeSelect');
  var brigadeId = sel.value;
  document.querySelectorAll('input[name=member_ids]').forEach(function(cb){{ cb.checked = false; }});
  if (!brigadeId) return;
  var ids = _brigadesData[brigadeId] || [];
  ids.forEach(function(id){{
    var cb = document.getElementById('memberCb_' + id);
    if (cb) cb.checked = true;
  }});
}}
function _applyTitul(){{
  var s = document.getElementById('titulSelect');
  var opt = s.options[s.selectedIndex];
  var hid = document.getElementById('titul_id_field');
  if (s.value) {{
    document.getElementById('location').value = (opt && opt.getAttribute('data-text')) || '';
    if (hid) hid.value = s.value;
  }} else {{
    if (hid) hid.value = '';
  }}
}}
function _applyWorkType(){{
  // Подставляет название выбранной типовой работы в «На выполнение работ».
  // Перезаписывает, если поле пустое ИЛИ было заполнено этой же автоподстановкой
  // (метка data-auto="1") — тогда смена работы обновляет название. Если кадровик
  // правил поле руками (метка снята обработчиком input ниже) — не трогаем.
  var sel = document.querySelector('select[name=work_type_id]');
  var ta = document.getElementById('workDescription');
  if (!sel.value) return;
  var manual = ta.value.trim() && ta.dataset.auto !== '1';
  if (manual) return;
  var name = _workTypesData[sel.value];
  if (name) {{ ta.value = name; ta.dataset.auto = '1'; }}
}}
function _applyValidTo(){{
  // Автозаполнение "Действует по" = "с" + 14 дней (предельный срок наряда по
  // 782н: не более 15 дней, т.е. дата начала + 14). Ставим ТОЛЬКО если поле ещё
  // пустое — иначе не затираем срок, вписанный кадровиком вручную.
  var from = document.getElementById('validFrom');
  var to = document.getElementById('validTo');
  if (!from.value || to.value) return;
  var d = new Date(from.value + 'T00:00:00');
  d.setDate(d.getDate() + 14);
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  to.value = y + '-' + m + '-' + day;
}}
// При загрузке применяем автоподстановки к значениям, которые браузер восстановил
// сам (bfcache/автозаполнение формы) — для них событие change НЕ срабатывает, и без
// этого поля остались бы пустыми, хотя селект уже показывает выбор. Вызываем только
// функции с guard на пустоту (не затирают заполненное вручную). _applyTitul НЕ здесь:
// он всегда перезаписывает «место работ», на загрузке это стёрло бы ручной ввод.
document.addEventListener('DOMContentLoaded', function(){{
  _applyWorkType();
  _applyValidTo();
  // Ручная правка «На выполнение работ» снимает метку авто — после этого смена
  // типовой работы больше не перезаписывает вписанное кадровиком.
  var _ta = document.getElementById('workDescription');
  if (_ta) _ta.addEventListener('input', function(){{ _ta.dataset.auto = ''; }});
}});
</script>
"""
    return _render("Наряды-допуски", body, active="production", role=request.session.get("role", ""))


@app.post("/production/work-orders/new")
def work_order_create(
    request: Request,
    number: str = Form(...),
    subdivision: str = Form(""),
    work_description: str = Form(...),
    location: str = Form(...),
    responsible_supervisor_id: str = Form(...),
    responsible_executor_id: str = Form(...),
    valid_from: str = Form(...),
    valid_to: str = Form(...),
    member_ids: list[str] = Form(default=[]),
    materials: str = Form(""),
    tools: str = Form(""),
    equipment: str = Form(""),
    special_machinery: str = Form(""),
    technological_card_ref: str = Form(""),
    safety_systems: str = Form(""),
    special_conditions: str = Form(""),
    work_type_id: str = Form(""),
    titul_id: str = Form(""),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    actor = _actor_name(request, db)
    order = prod.create_work_order(
        db, number, work_description, location,
        responsible_supervisor_id, responsible_executor_id, actor,
        datetime.strptime(valid_from, "%Y-%m-%d").date(),
        datetime.strptime(valid_to, "%Y-%m-%d").date(),
        member_ids,
        subdivision=subdivision or None,
        materials=materials or None,
        tools=tools or None,
        equipment=equipment or None,
        special_machinery=special_machinery or None,
        technological_card_ref=technological_card_ref or None,
        safety_systems=safety_systems or None,
        special_conditions=special_conditions or None,
        work_type_id=work_type_id or None,
        titul_id=titul_id or None,
    )
    members = db.query(WorkOrderMember).filter_by(work_order_id=order.id).all()
    problems = prod.check_work_order_problems(order, members) if work_type_id else []
    if problems:
        items = "".join(f"<li>{p}</li>" for p in problems)
        body = f"""
<h1>⚠️ Наряд №{order.number} создан, но есть замечания</h1>
<section class="grid">
<p class="muted">Наряд сохранён как черновик. По требованиям Правил 782н перед выпуском устраните:</p>
<ul>{items}</ul>
<p><a class="btn secondary" href="/production/work-orders/{order.id}/print">Всё равно распечатать</a>
<a class="btn" href="/production/work-orders">К нарядам</a></p>
</section>
"""
        return _render("Наряд: замечания", body, active="production",
                       role=request.session.get("role", ""))
    return RedirectResponse("/production/work-orders", status_code=303)


@app.post("/production/work-orders/{work_order_id}/members")
def work_order_set_members(
    work_order_id: str,
    request: Request,
    member_ids: list[str] = Form(default=[]),
    db: Session = Depends(get_db),
):
    """Замена состава бригады уже созданного наряда (дозаполнение черновика).
    Полностью заменяет членов: снимает прежних, добавляет отмеченных."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    order = db.get(WorkOrder, work_order_id)
    if order is None:
        raise HTTPException(404, "Наряд не найден.")
    db.query(WorkOrderMember).filter_by(work_order_id=work_order_id).delete()
    for eid in member_ids:
        db.add(WorkOrderMember(work_order_id=work_order_id, employee_id=eid))
    # Фиксируем первоначальный размер бригады при первом заполнении состава
    # (момент выпуска) — база для правила 782н о половине. Позже не трогаем.
    if not order.initial_member_count and member_ids:
        order.initial_member_count = len(member_ids)
    db.commit()
    return RedirectResponse("/production/work-orders", status_code=303)


@app.post("/production/work-orders/{work_order_id}/member-change")
def work_order_member_change(
    work_order_id: str,
    request: Request,
    change_type: str = Form(...),
    employee_id: str = Form(...),
    ordered_by: str = Form(...),
    db: Session = Depends(get_db),
):
    """Оформление изменения состава по пункту 7 (ввод/вывод) с проверкой правила
    половины (782н). При превышении половины изменение не вносится."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    try:
        ct = MemberChangeType(change_type)
    except ValueError:
        raise HTTPException(400, "Неверный тип изменения.")
    actor = _actor_name(request, db)
    ok, msg = prod.add_member_change(
        db, work_order_id, employee_id, ct, ordered_by.strip(), created_by=actor
    )
    # Сообщение (в т.ч. блокировку «нужен новый наряд») показываем на отдельной странице.
    body = (
        f'<h1>{"✅" if ok else "⚠️"} Пункт 7: изменение состава</h1>'
        f'<section class="grid"><p class="{"muted" if ok else "warning-banner"}">{html.escape(msg)}</p>'
        f'<p><a class="btn" href="/production/work-orders">К нарядам</a></p></section>'
    )
    return HTMLResponse(_render("Изменение состава", body, active="production",
                                role=request.session.get("role", "")))


@app.post("/production/work-orders/{work_order_id}/close")
def work_order_close(work_order_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.close_work_order(db, work_order_id)
    return RedirectResponse("/production/work-orders", status_code=303)


@app.post("/production/work-orders/{work_order_id}/delete")
def work_order_delete(work_order_id: str, request: Request, db: Session = Depends(get_db)):
    """Физическое удаление наряда через ORM: каскад delete-orphan снимает состав,
    ежедневные допуски и изменения состава автоматически (в отличие от сырого
    DELETE в psql, где внешние ключи приходится чистить вручную). Необратимо —
    подтверждение спрашивается на кнопке. Для штатного завершения есть «Закрыть»."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    order = db.get(WorkOrder, work_order_id)
    if order is None:
        raise HTTPException(404, "Наряд не найден.")
    db.delete(order)
    db.commit()
    return RedirectResponse("/production/work-orders", status_code=303)


@app.get("/production/work-orders/{work_order_id}/print")
def work_order_print(work_order_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    order = db.get(WorkOrder, work_order_id)
    if order is None:
        raise HTTPException(404, "Наряд не найден.")
    members = db.query(WorkOrderMember).filter_by(work_order_id=order.id).all()
    if order.work_type_id:
        path = prod.generate_height_work_order_docx(order, members, org_name=ORG_NAME)
    else:
        path = prod.generate_work_order_docx(order, members, org_name=ORG_NAME)
    with open(path, "rb") as f:
        data = f.read()
    fn = f"Наряд-допуск_№{order.number}.docx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": _content_disposition(fn)},
    )


# --- Инструктажи ---------------------------------------------------------------

@app.get("/production/instructions", response_class=HTMLResponse)
def instructions_page(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    due = prod.get_due_instructions(db)
    employees = db.scalars(
        select(Employee).where(Employee.contract_end_date.is_(None)).order_by(Employee.full_name)
    ).all()
    emp_options = "".join(f'<option value="{e.id}">{e.full_name}</option>' for e in employees)
    type_options = "".join(f'<option value="{t.value}">{label}</option>' for t, label in prod.INSTRUCTION_LABELS.items())

    due_rows = "".join(
        f'<div class="card">{d["name"]} — {"просрочен" if d["overdue"] else "срок"} '
        f'{d["due_date"]:%d.%m.%Y}<br>'
        f'<a class="btn" href="/employees/{d["employee_id"]}">Открыть карточку</a></div>'
        for d in due
    )

    # Автозаполнение вводного всем + допечатка журнала партиями по каждому типу —
    # см. договорённость: журналы заполняются ВСЕМИ сотрудниками (не по одному
    # вручную), печать — только новых непечатанных записей, отдельная нумерация
    # на каждый InstructionType.
    need_intro = len(prod.get_employees_needing_introductory(db))
    need_primary = len(prod.get_employees_needing_instruction(db, InstructionType.PRIMARY_WORKPLACE))
    journal_rows = ""
    for t, label in prod.INSTRUCTION_LABELS.items():
        unprinted = len(prod.get_unprinted_instructions(db, t))
        journal_rows += (
            f'<div class="card">{label}<br>'
            f'<span class="badge {"orange" if unprinted else "neutral"}">Ждут печати: {unprinted}</span><br>'
            f'<form method="post" action="/production/instructions/print" style="display:inline">'
            f'<input type="hidden" name="instruction_type" value="{t.value}">'
            f'<button type="submit" class="btn secondary">Допечатать новые записи</button></form></div>'
        )

    body = f"""
<h1>🎓 Инструктажи</h1>
{f'<section class="grid"><h2>⏰ Требуют повторного проведения ({len(due)})</h2>{due_rows}</section>' if due else ''}

<section class="grid">
<h2>Журналы — допечатка партиями</h2>
{journal_rows}
</section>

<section>
<h2>Вводный инструктаж — автозаполнение</h2>
<p class="muted">Сотрудников без вводного инструктажа: {need_intro}. Заводит запись каждому
датой начала работы (дата договора, а если её нет — дата въезда) — не сегодняшним числом,
чтобы порядок строк в журнале при печати совпадал с реальной хронологией приёма.</p>
<form method="post" action="/production/instructions/auto-introductory">
<button type="submit"{" disabled" if not need_intro else ""}>Заполнить вводный всем ({need_intro})</button>
</form>
</section>

<section>
<h2>Первичный на рабочем месте — автозаполнение</h2>
<p class="muted">Сотрудников без первичного инструктажа: {need_primary}. Проводится вместе с
вводным в день начала работы, но регистрируется в ОТДЕЛЬНОМ журнале со своей нумерацией.
Заводит запись каждому датой начала работы — так же, как вводный.</p>
<form method="post" action="/production/instructions/auto-primary">
<button type="submit"{" disabled" if not need_primary else ""}>Заполнить первичный всем ({need_primary})</button>
</form>
</section>

<section>
<h2>⚠️ ВРЕМЕННО (тест)</h2>
<p class="muted">Очищает ВСЕ записи инструктажей всех видов — вводный, на рабочем месте,
повторный, внеплановый, целевой. Только для проверки автозаполнения/допечатки с чистого
листа. Удалить эту кнопку и функцию production.test_clear_all_instructions после теста.</p>
<form method="post" action="/production/instructions/test-clear"
onsubmit="return confirm('Удалить ВСЕ записи инструктажей без возможности отмены?')">
<button type="submit" class="secondary">🧪 Очистить инструктажи (тест)</button>
</form>
</section>

<section>
<h2>Провести инструктаж (по одному)</h2>
<form method="post" action="/production/instructions/new">
<label>Сотрудник: <select name="employee_id" required>{emp_options}</select></label>
<label>Вид: <select name="instruction_type" required>{type_options}</select></label>
<input type="text" name="topic" placeholder="Тема (для целевого/внепланового)">
<label>Следующий срок (для повторного, необязательно): <input type="date" name="next_due_date"></label>
<button type="submit">Зафиксировать проведение</button>
</form>
</section>
<p><a class="btn secondary" href="/production">← Производство</a></p>
"""
    return _render("Инструктажи", body, active="production", role=request.session.get("role", ""))


@app.post("/production/instructions/auto-introductory")
def instruction_auto_introductory(request: Request, db: Session = Depends(get_db)):
    """Вводный инструктаж всем сотрудникам, у кого его ещё нет — датой начала работы,
    не сегодняшним числом (см. договорённость в чате)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    actor = _actor_name(request, db)
    prod.auto_create_introductory_instructions(db, actor)
    return RedirectResponse("/production/instructions", status_code=303)


@app.post("/production/instructions/auto-primary")
def instruction_auto_primary(request: Request, db: Session = Depends(get_db)):
    """Первичный на рабочем месте всем, у кого его ещё нет — датой начала работы.
    Проводится вместе с вводным, но отдельный журнал. Защищён unique-индексом
    uq_primary_workplace_once от дублей при двойном нажатии (см. auto_create_instructions)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    actor = _actor_name(request, db)
    prod.auto_create_instructions(db, InstructionType.PRIMARY_WORKPLACE, actor)
    return RedirectResponse("/production/instructions", status_code=303)


@app.post("/production/instructions/test-clear")
def instruction_test_clear(request: Request, db: Session = Depends(get_db)):
    """ВРЕМЕННО (тест) — удаляет ВСЕ записи инструктажей, чтобы проверить
    автозаполнение/допечатку с чистого листа. УДАЛИТЬ этот роут вместе с
    production.test_clear_all_instructions и кнопкой в instructions_page,
    когда тестирование закончится."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.test_clear_all_instructions(db)
    return RedirectResponse("/production/instructions", status_code=303)


@app.post("/production/instructions/print")
def instruction_print_journal(
    request: Request, instruction_type: str = Form(...), db: Session = Depends(get_db),
):
    """Допечатать новые записи журнала — присваивает номера строк, отдаёт xlsx
    (2026-07: переведено с docx на Excel, см. production.print_new_journal_entries /
    generate_instruction_journal_xlsx). Довеска прочерками нет — журнал кончается
    на последней записи, под таблицей итог "Внесено записей: N"."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    t = InstructionType(instruction_type)
    printed = prod.print_new_journal_entries(db, t)
    if not printed:
        return RedirectResponse("/production/instructions", status_code=303)
    order_ref = prod.get_latest_order_ref(db)
    started_at = prod.get_journal_started_at(db, t)
    path = prod.generate_instruction_journal_xlsx(
        printed, t, org_name=ORG_NAME, order_ref=order_ref, started_at=started_at,
    )
    with open(path, "rb") as f:
        data = f.read()
    label = prod.INSTRUCTION_LABELS.get(t, t.value)
    fn = f"Журнал_{label.replace(' ', '_')}_{datetime.now(MSK):%d.%m.%Y}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": _content_disposition(fn)},
    )


@app.post("/production/instructions/new")
def instruction_create(
    request: Request,
    employee_id: str = Form(...),
    instruction_type: str = Form(...),
    topic: str = Form(""),
    next_due_date: str = Form(""),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    actor = _actor_name(request, db)
    due = datetime.strptime(next_due_date, "%Y-%m-%d").date() if next_due_date else None
    prod.create_instruction(
        db, employee_id, InstructionType(instruction_type), actor,
        topic=topic or None, next_due_date=due,
    )
    return RedirectResponse("/production/instructions", status_code=303)


# --- Удостоверения (корочки) ---------------------------------------------------

@app.get("/production/certificates", response_class=HTMLResponse)
def certificates_page(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    expiring = prod.get_expiring_certificates(db)
    employees = db.scalars(
        select(Employee).where(Employee.contract_end_date.is_(None)).order_by(Employee.full_name)
    ).all()
    emp_options = "".join(f'<option value="{e.id}">{e.full_name}</option>' for e in employees)

    expiring_rows = "".join(
        f'<div class="card">{item["name"]} — {item["profession"]}<br>'
        f'<span class="badge {"red" if item["overdue"] else "orange"}">'
        f'{"Просрочено" if item["overdue"] else "Истекает"} {item["expiry_date"]:%d.%m.%Y}</span><br>'
        f'<a class="btn" href="/employees/{item["employee_id"]}">Открыть карточку</a></div>'
        for item in expiring
    )

    all_certs = db.query(Certificate).order_by(Certificate.created_at.desc()).all()

    def cert_row(c) -> str:
        status = prod.certificate_status(c)
        badge_class = {"expired": "red", "expiring_soon": "orange", "active": "green", "no_expiry": "neutral"}[status]
        badge_text = {"expired": "Просрочено", "expiring_soon": "Истекает", "active": "Действует", "no_expiry": "Бессрочное"}[status]
        exp = f' до {c.expiry_date:%d.%m.%Y}' if c.expiry_date else ''
        scan_link = (f'<a class="btn secondary" href="/production/certificates/{c.id}/download">Скачать скан</a>'
                     if c.scan_key else '<span class="muted">без скана</span>')
        return (
            f'<div class="card">{c.employee.full_name if c.employee else "?"} — {c.profession}<br>'
            f'<span class="badge {badge_class}">{badge_text}{exp}</span><br>{scan_link}</div>'
        )

    all_rows = "".join(cert_row(c) for c in all_certs)
    body = f"""
<h1>🪪 Удостоверения по профессиям</h1>
{f'<section class="grid"><h2>⏰ Истекают/просрочены ({len(expiring)})</h2>{expiring_rows}</section>' if expiring else ''}
<section class="grid"><h2>Все удостоверения ({len(all_certs)})</h2>{all_rows or '<p class="muted">Пока нет ни одного.</p>'}</section>
<section>
<h2>Добавить удостоверение</h2>
<form method="post" action="/production/certificates/new" enctype="multipart/form-data">
<label>Сотрудник: <select name="employee_id" required>{emp_options}</select></label>
<input type="text" name="profession" placeholder="Профессия / вид допуска (например «Электробезопасность IV группа»)" required>
<input type="text" name="issued_by_org" placeholder="Кем выдано">
<label>Дата выдачи: <input type="date" name="issue_date"></label>
<label>Действует до (пусто — бессрочное): <input type="date" name="expiry_date"></label>
<label>Скан:</label>
<input type="file" name="scan_file" accept="application/pdf,image/*,.doc,.docx" style="display:block;width:100%;margin:8px 0;padding:10px;border:1px solid #d9dde3;border-radius:8px;background:#fff;font-size:16px">
<button type="submit">Добавить</button>
</form>
</section>
<p><a class="btn secondary" href="/production">← Производство</a></p>
"""
    return _render("Удостоверения", body, active="production", role=request.session.get("role", ""))


@app.post("/production/certificates/new")
async def certificate_create(
    request: Request,
    employee_id: str = Form(...),
    profession: str = Form(...),
    issued_by_org: str = Form(""),
    issue_date: str = Form(""),
    expiry_date: str = Form(""),
    scan_file: UploadFile = File(None),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    cert = prod.create_certificate(
        db, employee_id, profession,
        issued_by_org=issued_by_org or None,
        issue_date=datetime.strptime(issue_date, "%Y-%m-%d").date() if issue_date else None,
        expiry_date=datetime.strptime(expiry_date, "%Y-%m-%d").date() if expiry_date else None,
    )

    if scan_file is not None and scan_file.filename:
        # Свой scan_type на каждый сертификат (сотрудник может иметь несколько
        # удостоверений разных профессий) — _s3_upload/_scan_key строят ключ
        # как employee_{id}/{scan_type}, поэтому тип должен быть уникальным
        # в рамках сотрудника, не фиксированным "certificate".
        from s3_storage import _s3_upload, _scan_key
        scan_type = f"certificate_{cert.id}"
        content = await scan_file.read()
        content_type = scan_file.content_type or "application/octet-stream"
        try:
            _s3_upload(scan_type, employee_id, content, content_type)
            prod.set_certificate_scan_key(db, cert.id, _scan_key(scan_type, employee_id))
        except RuntimeError as e:
            log.warning("certificate_create: загрузка скана не удалась: %s", e)
            # Удостоверение уже создано без скана — не откатываем, просто без файла.

    return RedirectResponse("/production/certificates", status_code=303)


@app.get("/production/certificates/{certificate_id}/download")
def certificate_download(certificate_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    cert = db.get(Certificate, certificate_id)
    if cert is None or not cert.scan_key:
        raise HTTPException(404, "Скан не найден.")
    from s3_storage import _s3_download
    # scan_key хранится как employee_{id}/certificate_{cert.id} — scan_type для
    # _s3_download нужен без префикса "employee_{id}/", восстанавливаем.
    scan_type = f"certificate_{cert.id}"
    try:
        data, ct = _s3_download(scan_type, cert.employee_id)
    except RuntimeError as e:
        raise HTTPException(404, str(e))
    ext = "pdf" if "pdf" in ct else ("jpg" if "jpeg" in ct or "jpg" in ct else "bin")
    fio = (cert.employee.full_name if cert.employee else "").strip().replace(" ", "_")
    fn = f"{fio}_{cert.profession.replace(' ', '_')}.{ext}" if fio else f"certificate.{ext}"
    return Response(content=data, media_type=ct,
                    headers={"Content-Disposition": _content_disposition(fn)})


# --- Приказы (2026-07) ---------------------------------------------------------

@app.get("/production/orders", response_class=HTMLResponse)
def orders_page(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    orders = prod.get_orders(db)

    _category_label = {
        OrderCategory.PERSONNEL: "Кадровый", OrderCategory.PRODUCTION: "Производственный",
        OrderCategory.LABOR_PROTECTION: "Охрана труда", OrderCategory.TRAINING: "Обучение и инструктажи",
        OrderCategory.HEIGHT_WORK: "Работы на высоте", OrderCategory.FIRE_SAFETY: "Пожарная безопасность",
        OrderCategory.ELECTRICAL: "Электробезопасность", OrderCategory.OTHER: "Прочее",
    }
    _category_badge = {
        OrderCategory.PERSONNEL: "orange", OrderCategory.PRODUCTION: "green",
        OrderCategory.LABOR_PROTECTION: "green", OrderCategory.TRAINING: "neutral",
        OrderCategory.HEIGHT_WORK: "amber", OrderCategory.FIRE_SAFETY: "red",
        OrderCategory.ELECTRICAL: "neutral", OrderCategory.OTHER: "neutral",
    }

    def order_row(o) -> str:
        scan_link = (f'<a class="btn secondary" href="/production/orders/{o.id}/download">Скачать скан</a>'
                     if o.scan_key else '<span class="muted">без скана</span>')
        note = f'<br><span class="muted">{o.note}</span>' if o.note else ''
        badge = f'<span class="badge {_category_badge[o.category]}">{_category_label[o.category]}</span>'
        return (
            f'<div class="card">№ {o.number} от {o.order_date:%d.%m.%Y} — {o.topic} {badge}{note}<br>'
            f'{scan_link}</div>'
        )

    rows = "".join(order_row(o) for o in orders)
    # Опции генератора приказов по ОТ — сгруппированы по разделам (optgroup) из справочника
    # prod.OT_ORDERS. topic (2-й элемент кортежа) — подпись; ключ приказа — value.
    _ot_groups = {}
    for _key in prod.OT_ORDER_KEYS:
        _cat, _topic = prod.OT_ORDERS[_key][0], prod.OT_ORDERS[_key][1]
        _ot_groups.setdefault(_cat, []).append((_key, _topic))
    _ot_options = ""
    for _cat, _items in _ot_groups.items():
        _ot_options += f'<optgroup label="{html.escape(prod.OT_SECTIONS[_cat])}">'
        for _key, _topic in _items:
            _ot_options += f'<option value="{_key}">{html.escape(_topic)}</option>'
        _ot_options += '</optgroup>'
    _today_iso = datetime.now(MSK).date().isoformat()
    # Предзаполнение первой записи данными уже готового приказа №20-ПСМ/2026 —
    # только пока реестр пуст, чтобы не подставлять эти значения повторно для
    # следующих приказов.
    _prefill = not orders
    _pf_number = "20-ПСМ/2026" if _prefill else ""
    _pf_date = "2026-06-01" if _prefill else ""
    _pf_topic = ("О порядке ведения журналов инструктажей по охране труда и учёта "
                 "нарядов-допусков" if _prefill else "")
    category_options = "".join(
        f'<option value="{c.value}"{" selected" if _prefill and c == OrderCategory.PRODUCTION else ""}>{label}</option>'
        for c, label in _category_label.items()
    )
    body = f"""
<h1>📑 Приказы</h1>
<section class="grid"><h2>Реестр ({len(orders)})</h2>{rows or '<p class="muted">Пока пусто.</p>'}</section>
<section>
<h2>Сгенерировать приказ по охране труда</h2>
<p class="muted">Готовые шаблоны приказов по ОТ (реквизиты ИП Буц подставляются автоматически).
Выберите приказ, укажите номер и дату — система сформирует .docx для печати. Подписанный скан
загрузите обратно через форму ниже.</p>
<form method="post" action="/production/orders/generate">
<label>Приказ:<br><select name="order_key" required>{_ot_options}</select></label>
<label>Номер: <input type="text" name="number" placeholder="например: 01-ОТ/2026" required></label>
<label>Дата: <input type="date" name="order_date" value="{_today_iso}" required></label>
<button type="submit">Сгенерировать .docx</button>
</form>
</section>
<section>
<h2>Новый приказ (в реестр)</h2>
<form method="post" action="/production/orders/new" enctype="multipart/form-data">
<input type="text" name="number" id="orderNumber" placeholder="Номер (например: 20-ПСМ/2026)" value="{_pf_number}" required>
<label>Дата: <input type="date" name="order_date" id="orderDate" value="{_pf_date}" required></label>
<input type="text" name="topic" id="orderTopic" placeholder="Тема приказа" value="{html.escape(_pf_topic)}" required>
<label>Раздел: <select name="category">{category_options}</select></label>
<textarea name="note" placeholder="Примечание (необязательно)" rows="2"></textarea>
<label>Скан приказа (PDF/фото):</label>
<input type="file" name="scan_file" id="orderScanFile" accept="application/pdf,image/*,.doc,.docx" style="display:block;width:100%;margin:8px 0;padding:10px;border:1px solid #d9dde3;border-radius:8px;background:#fff;font-size:16px">
<p class="muted" style="margin:0 0 10px">Имя файла вида Prikaz_20-ПСМ2026_Тема_01-06-2026 —
поля заполнятся сами при выборе файла.</p>
<button type="submit">Добавить в реестр</button>
</form>
</section>
<p><a class="btn secondary" href="/production">← Производство</a></p>
<script>
(function(){{
  var input = document.getElementById('orderScanFile');
  if (!input) return;
  // Известные темы по короткому обозначению в имени файла — расширять по мере
  // появления новых приказов с другими темами.
  var topicMap = {{
    'OT-zhurnaly-naryady': 'О порядке ведения журналов инструктажей по охране труда и учёта нарядов-допусков',
    'zhurnaly-naryady': 'О порядке ведения журналов инструктажей по охране труда и учёта нарядов-допусков'
  }};
  var customerMap = {{'PSM': 'ПСМ'}};
  input.addEventListener('change', function(e){{
    var f = e.target.files && e.target.files[0];
    if (!f) return;
    var name = f.name.replace(/\\.[^.]+$/, '');
    // Формат: Prikaz_{{номер}}-{{заказчик}}{{год}}_{{тема}}_{{дд}}-{{мм}}-{{гггг}}
    var m = name.match(/^Prikaz_(\\d+)-([A-Za-zА-Яа-я]+)(\\d{{4}})_(.+)_(\\d{{2}})-(\\d{{2}})-(\\d{{4}})$/);
    if (!m) return;  // имя не по формату — не мешаем ручному вводу
    var num = m[1], customerRaw = m[2].toUpperCase(), year = m[3];
    var topicSlug = m[4], dd = m[5], mm = m[6], yyyy = m[7];
    var customerRu = customerMap[customerRaw] || customerRaw;
    document.getElementById('orderNumber').value = num + '-' + customerRu + '/' + year;
    document.getElementById('orderDate').value = yyyy + '-' + mm + '-' + dd;
    document.getElementById('orderTopic').value = topicMap[topicSlug] || topicSlug.replace(/-/g, ' ');
  }});
}})();
</script>
"""
    return _render("Приказы", body, active="production", role=request.session.get("role", ""))


@app.post("/production/orders/new")
async def order_create(
    request: Request,
    number: str = Form(...),
    order_date: str = Form(...),
    topic: str = Form(...),
    category: str = Form(OrderCategory.OTHER.value),
    note: str = Form(""),
    scan_file: UploadFile = File(None),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    order = prod.create_order(
        db, number, datetime.strptime(order_date, "%Y-%m-%d").date(), topic,
        category=OrderCategory(category), note=note or None,
    )
    if scan_file is not None and scan_file.filename:
        from s3_storage import _s3_upload, _scan_key
        scan_type = f"order_{order.id}"
        content = await scan_file.read()
        content_type = scan_file.content_type or "application/octet-stream"
        try:
            _s3_upload(scan_type, None, content, content_type)
            prod.set_order_scan_key(db, order.id, _scan_key(scan_type, None))
        except RuntimeError as e:
            log.warning("order_create: загрузка скана не удалась: %s", e)
    return RedirectResponse("/production/orders", status_code=303)


@app.post("/production/orders/generate")
def order_generate(
    request: Request,
    order_key: str = Form(...),
    number: str = Form(...),
    order_date: str = Form(...),
    db: Session = Depends(get_db),
):
    """Генерирует приказ по ОТ из справочника prod.OT_ORDERS и отдаёт .docx для печати.
    В реестр запись НЕ создаётся автоматически — после печати и подписи пользователь
    загружает подписанный скан через форму «Новый приказ» (номер/тема заполнятся из имени)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    if order_key not in prod.OT_ORDERS:
        raise HTTPException(400, "Неизвестный приказ.")
    try:
        d = datetime.strptime(order_date, "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(400, "Неверная дата.")
    import tempfile, os
    tmpdir = tempfile.mkdtemp()
    path = prod.generate_ot_order_docx(order_key, number, d, tmpdir)
    with open(path, "rb") as f:
        data = f.read()
    topic = prod.OT_ORDERS[order_key][1]
    safe_topic = topic[:40].replace("/", "-").replace(" ", "_")
    fn = f"Приказ_{number.replace('/', '-')}_{safe_topic}.docx"
    try:
        os.remove(path); os.rmdir(tmpdir)
    except OSError:
        pass
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": _content_disposition(fn)},
    )
def order_download(order_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    order = db.get(InternalOrder, order_id)
    if order is None or not order.scan_key:
        raise HTTPException(404, "Скан не найден.")
    from s3_storage import _s3_download
    scan_type = f"order_{order.id}"
    try:
        data, ct = _s3_download(scan_type, None)
    except RuntimeError as e:
        raise HTTPException(404, str(e))
    ext = "pdf" if "pdf" in ct else ("jpg" if "jpeg" in ct or "jpg" in ct else "bin")
    fn = f"Приказ_{order.number.replace('/', '-')}.{ext}"
    return Response(content=data, media_type=ct,
                    headers={"Content-Disposition": _content_disposition(fn)})


# --- Бригады (2026-07) ---------------------------------------------------------

@app.get("/production/brigades", response_class=HTMLResponse)
def brigades_page(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    brigades = prod.get_brigades(db)
    employees = db.scalars(
        select(Employee).where(Employee.contract_end_date.is_(None)).order_by(Employee.full_name)
    ).all()
    # off_tabel-сотрудники (ИП-руководитель) не входят в состав бригады-исполнителей
    # по Правилам 782н. Пул выбора — без них; имена уже сохранённых членов берём из
    # полного employees, чтобы состав не пропадал из сводки.
    workers = [e for e in employees if not e.off_tabel]
    emp_checkboxes = "".join(
        f'<label style="display:block;margin:2px 0"><input type="checkbox" name="member_ids" '
        f'value="{e.id}"> {e.full_name}</label>'
        for e in workers
    )

    def brigade_row(b) -> str:
        member_ids = set(prod.get_brigade_member_ids(db, b.id))
        names = [e.full_name for e in employees if e.id in member_ids]
        # Чекбоксы редактирования: текущие члены отмечены.
        edit_checkboxes = "".join(
            f'<label style="display:block;margin:2px 0"><input type="checkbox" name="member_ids" '
            f'value="{e.id}"{" checked" if e.id in member_ids else ""}> {html.escape(e.full_name)}</label>'
            for e in workers
        )
        return (
            f'<div class="card">{html.escape(b.name)}<br>'
            f'<span class="muted">{", ".join(html.escape(n) for n in names) or "пусто"}</span><br>'
            f'<details style="margin-top:6px">'
            f'<summary style="cursor:pointer">✏️ Изменить</summary>'
            f'<form method="post" action="/production/brigades/{b.id}/update" style="margin-top:8px">'
            f'<input type="text" name="name" value="{html.escape(b.name)}" required '
            f'style="display:block;margin-bottom:6px">'
            f'<fieldset><legend>Состав</legend>{edit_checkboxes}</fieldset>'
            f'<button type="submit">Сохранить изменения</button>'
            f'</form></details>'
            f'<form method="post" action="/production/brigades/{b.id}/delete" style="display:inline;margin-top:6px"'
            f' onsubmit="return confirm(\'Удалить бригаду?\')">'
            f'<button type="submit" class="btn secondary">Удалить</button></form></div>'
        )

    rows = "".join(brigade_row(b) for b in brigades)
    body = f"""
<h1>👷 Бригады</h1>
<section class="grid"><h2>Сохранённые ({len(brigades)})</h2>{rows or '<p class="muted">Пока нет ни одной.</p>'}</section>
<section>
<h2>Новая бригада</h2>
<form method="post" action="/production/brigades/new">
<input type="text" name="name" placeholder="Название (например: Бригада бетонщиков №1)" required>
<fieldset><legend>Состав</legend>{emp_checkboxes}</fieldset>
<button type="submit">Сохранить</button>
</form>
</section>
<p><a class="btn secondary" href="/production">← Производство</a></p>
"""
    return _render("Бригады", body, active="production", role=request.session.get("role", ""))


@app.post("/production/brigades/new")
def brigade_create(
    request: Request, name: str = Form(...),
    member_ids: list[str] = Form(default=[]), db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.create_brigade(db, name, member_ids)
    return RedirectResponse("/production/brigades", status_code=303)


@app.post("/production/brigades/{brigade_id}/delete")
def brigade_delete(brigade_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.delete_brigade(db, brigade_id)
    return RedirectResponse("/production/brigades", status_code=303)


@app.post("/production/brigades/{brigade_id}/update")
def brigade_update(
    brigade_id: str, request: Request, name: str = Form(...),
    member_ids: list[str] = Form(default=[]), db: Session = Depends(get_db),
):
    """Переименование + замена состава одной формой (см. prod.update_brigade)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.update_brigade(db, brigade_id, name, member_ids)
    return RedirectResponse("/production/brigades", status_code=303)


# --- Титулы (справочник объектов) --------------------------------------------


@app.get("/production/tituly", response_class=HTMLResponse)
def tituly_page(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    tituly = prod.get_tituly(db)

    def titul_row(t) -> str:
        return (
            f'<div class="card"><b>{html.escape(t.code)}</b> — {html.escape(t.name)}'
            f'<details style="margin-top:6px">'
            f'<summary style="cursor:pointer">✏️ Изменить</summary>'
            f'<form method="post" action="/production/tituly/{t.id}/update" style="margin-top:8px">'
            f'<input type="text" name="code" value="{html.escape(t.code)}" placeholder="Шифр" required '
            f'style="display:block;margin-bottom:6px">'
            f'<input type="text" name="name" value="{html.escape(t.name)}" placeholder="Наименование" required '
            f'style="display:block;margin-bottom:6px">'
            f'<button type="submit">Сохранить изменения</button>'
            f'</form></details>'
            f'<form method="post" action="/production/tituly/{t.id}/delete" style="display:inline;margin-top:6px"'
            f' onsubmit="return confirm(\'Удалить титул?\')">'
            f'<button type="submit" class="btn secondary">Удалить</button></form></div>'
        )

    rows = "".join(titul_row(t) for t in tituly)
    body = f"""
<h1>🏗 Титулы</h1>
<section class="grid"><h2>Сохранённые ({len(tituly)})</h2>{rows or '<p class="muted">Пока нет ни одного.</p>'}</section>
<section>
<h2>Новый титул</h2>
<form method="post" action="/production/tituly/new">
<input type="text" name="code" placeholder="Шифр (например: 15.21)" required>
<input type="text" name="name" placeholder="Наименование (например: Лаборатория)" required>
<button type="submit">Сохранить</button>
</form>
</section>
<p><a class="btn secondary" href="/production">← Производство</a></p>
"""
    return _render("Титулы", body, active="production", role=request.session.get("role", ""))


@app.post("/production/tituly/new")
def titul_create(
    request: Request, code: str = Form(...), name: str = Form(...),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.create_titul(db, code, name)
    return RedirectResponse("/production/tituly", status_code=303)


@app.post("/production/tituly/{titul_id}/delete")
def titul_delete(titul_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.delete_titul(db, titul_id)
    return RedirectResponse("/production/tituly", status_code=303)


@app.post("/production/tituly/{titul_id}/update")
def titul_update(
    titul_id: str, request: Request, code: str = Form(...), name: str = Form(...),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    prod.update_titul(db, titul_id, code, name)
    return RedirectResponse("/production/tituly", status_code=303)


# --- Сотрудники: список + единая карточка ------------------------------------


# --- Маршруты домена сотрудников вынесены в отдельный файл (2026-07) ----------
# webforms.py разросся до 5300+ строк и медленно открывался в мобильном
# редакторе. Домен сотрудников (список, карточка, медкомиссия, трудовые
# договоры, госпошлина, увольнения, сканы, общие документы) вынесен в
# webforms_employees.py. Импорт СТРОГО последней строкой: к этому моменту app и
# все хелперы/константы уже определены выше, поэтому `from webforms import ...`
# в том файле резолвится без циклической ошибки. Роуты регистрируются на том же
# экземпляре app (декораторы @app.get/@app.post внутри импортируемого модуля).
import webforms_employees  # noqa: E402,F401
