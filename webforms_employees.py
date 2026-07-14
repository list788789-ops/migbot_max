"""
webforms_employees.py — маршруты домена сотрудников migbot.

Выделено из webforms.py (2026-07) ради скорости открытия файла в мобильном
редакторе. Здесь ТОЛЬКО роуты и хелперы, относящиеся к сотрудникам: список,
карточка, создание/OCR, медкомиссия, трудовые договоры, госпошлина, увольнения,
обязательства, сканы (S3) и общие документы.

Общий экземпляр FastAPI `app`, функция `get_db` и хелперы рендера/авторизации
импортируются из webforms — там же они и остаются определёнными. Библиотечные,
модельные и шаблонные имена импортируются напрямую из исходных модулей (как в
webforms.py), чтобы этот файл не зависел от того, что webforms что-то реэкспортит.

ВАЖНО: этот модуль импортируется ПОСЛЕДНЕЙ строкой webforms.py. Обратный импорт
`from webforms import ...` ниже работает именно потому, что к моменту импорта
всё перечисленное в webforms уже определено.
"""
from __future__ import annotations

import html
import os
import json
import calendar
import logging
from datetime import date, datetime, timedelta, timezone

from fastapi import Depends, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import (
    FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response,
)
from sqlalchemy import select
from sqlalchemy.orm import Session

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
    EventDeletionLog,
    ExamStatus,
    Obligation,
    ObligationStatus,
    ObligationType,
    Referral,
    RegistrationStatus,
    RotationReturn,
    SystemFlag,
    WorkOrderMember,
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
from s3_storage import (
    S3_ENDPOINT, S3_BUCKET, S3_REGION, S3_ACCESS_KEY, S3_SECRET_KEY,
    SCAN_TYPES, PAYMENT_SCAN_TYPES, COMMON_DOC_TYPES, SCAN_COMMON_TYPES,
    _s3_client, _scan_key, _s3_upload, _s3_list_for_employee, _s3_download,
    _s3_delete, _common_key, _s3_upload_common, _s3_list_common,
    _s3_download_common, _s3_delete_common, _s3_clear_check,
)

# --- Общий app, сессия БД, хелперы и локальные константы — из webforms.py -----
from webforms import (
    app,
    get_db,
    _render,
    _logged_in,
    _current_user,
    _require_role,
    _actor_name,
    _obligation_status,
    _package_missing,
    _content_disposition,
    _ext_for,
    MSK,
    CLINIC_ID,
    CONSENT_TEXT_VERSION,
    OBLIGATION_LABELS,
    SAVE_FORM_JS,
)

log = logging.getLogger("webforms_employees")


@app.get("/employees", response_class=HTMLResponse)
def employees_list(request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    today = datetime.now(MSK).date()
    # 2026-07: исключены уволенные (contract_end_date заполнен) — для них уже есть
    # отдельный /archive; раньше показывались в обоих местах одновременно.
    employees = db.scalars(
        select(Employee).where(Employee.contract_end_date.is_(None))
        .order_by(Employee.full_name)
    ).all()

    def nearest_pending(e: Employee):
        obs = [o for o in e.obligations if o.is_current and o.status == ObligationStatus.PENDING]
        return min(obs, key=lambda o: o.deadline_date) if obs else None

    # Медстатусы для чипа в списке: одним запросом все Referral (последний на работника).
    from models import Referral as _Ref, ExamStatus as _ES
    _all_refs = db.scalars(select(_Ref).order_by(_Ref.referral_date.desc())).all()
    _ref_by_emp = {}
    for _r in _all_refs:
        _ref_by_emp.setdefault(_r.employee_id, _r)  # первый = самый свежий (order desc)

    def _med_chip(e: Employee) -> str:
        """Краткий чип статуса медкомиссии для списка. Пусто, если работник ещё не прибыл."""
        r = _ref_by_emp.get(e.id)
        if r is None:
            if e.entry_date and (today - e.entry_date).days >= 5:
                return '<span class="badge amber">мед: пора направление</span>'
            return ''
        if r.exam_status == _ES.COMPLETED:
            return '<span class="badge green">мед: пройдена</span>'
        _md = (today - r.referral_date).days
        if _md > 14:
            return '<span class="badge red">мед: справка просрочена</span>'
        if _md >= 10:
            return '<span class="badge amber">мед: ждём справку</span>'
        return '<span class="badge neutral">мед: направлен</span>'

    def row(e: Employee) -> str:
        cit = e.citizenship or "—"
        # Ожидающие прибытия (без даты въезда): обязательства по въезду не создаются,
        # поэтому чип нейтральный "ожидает прибытия", а не "в норме" (которое врёт — человека
        # ещё нет). Дедлайны появятся, когда в карточку впишут дату въезда.
        if e.entry_date is None:
            chip = '<span class="badge neutral">ожидает прибытия</span>'
            ob_line = '<div class="muted-line">дата въезда не указана — дедлайны не идут</div>'
        elif e.consent_status != ConsentStatus.CONFIRMED:
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
        # data-search: ФИО + табельный в нижнем регистре — по нему фильтрует JS-поиск.
        _search_key = f"{e.full_name or ''} {e.tab_number or ''}".lower()
        return (
            f'<div class="card emp-row" data-search="{html.escape(_search_key, quote=True)}">{e.full_name} {chip} {_med_chip(e)}<br>'
            f'<span class="muted-line">{cit}</span>{ob_line}'
            f'<a class="btn" href="/employees/{e.id}">Открыть карточку</a></div>'
        )

    # Уволенные (contract_end_date заполнен) -> в архив, из основного списка убираем.
    working = [e for e in employees if e.contract_end_date is None]
    active = [e for e in working if e.entry_date is not None]
    awaiting = [e for e in working if e.entry_date is None]
    archived_count = sum(1 for e in employees if e.contract_end_date is not None)

    active_rows = "".join(row(e) for e in active) or '<p class="muted">Нет активных сотрудников.</p>'
    active_section = (
        f'<section class="grid"><h2>Активные ({len(active)})</h2>{active_rows}</section>'
    )

    awaiting_section = ""
    if awaiting:
        awaiting_rows = "".join(row(e) for e in awaiting)
        awaiting_section = (
            f'<section class="grid"><h2>Ожидают прибытия ({len(awaiting)})</h2>{awaiting_rows}</section>'
        )

    _search_box = (
        '<input type="text" id="empSearch" placeholder="Поиск по ФИО или табельному номеру…" '
        'oninput="_filterEmployees()" '
        'style="width:100%;font-size:16px;padding:12px 14px;margin:8px 0 4px;border:1px solid #b8c0cc;'
        'border-radius:12px;background:#fff">'
        '<p class="muted" id="empSearchEmpty" style="display:none">Никого не найдено.</p>'
    )
    _search_js = """
<script>
function _filterEmployees(){
  var q = (document.getElementById('empSearch').value || '').toLowerCase().trim();
  var rows = document.querySelectorAll('.emp-row');
  var shown = 0;
  rows.forEach(function(r){
    var key = r.getAttribute('data-search') || '';
    var match = q === '' || key.indexOf(q) !== -1;
    r.style.display = match ? '' : 'none';
    if(match) shown++;
  });
  var empty = document.getElementById('empSearchEmpty');
  if(empty) empty.style.display = (shown === 0 && q !== '') ? 'block' : 'none';
  // скрыть заголовки секций, если в них никого не осталось
  document.querySelectorAll('.grid').forEach(function(sec){
    var vis = sec.querySelectorAll('.emp-row:not([style*="display: none"])').length;
    sec.style.display = (q !== '' && vis === 0) ? 'none' : '';
  });
}
</script>"""
    _archive_link = (
        f'<a class="btn secondary" href="/archive">Архив уволенных ({archived_count})</a>'
        if archived_count else ''
    )
    return _render(
        "Сотрудники",
        f'<h1>Сотрудники ({len(working)})</h1>'
        f'<p><a class="btn" href="/employees/new">+ Добавить сотрудника</a> {_archive_link}</p>'
        f'{_search_box}'
        f'{active_section}{awaiting_section}'
        f'{_search_js}',
        active="employees",
        role=request.session.get("role", ""),
    )


@app.get("/archive", response_class=HTMLResponse)
def employees_archive(request: Request, db: Session = Depends(get_db)):
    """Архив уволенных сотрудников (contract_end_date заполнен). Их обязательства (уведомление
    об убытии и пр.) остаются в задачах/дашборде — здесь только список для истории/справок."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    archived = db.scalars(
        select(Employee).where(Employee.contract_end_date.is_not(None))
        .order_by(Employee.full_name)
    ).all()
    if not archived:
        body = ('<h1>Архив уволенных</h1>'
                '<p class="muted">Уволенных сотрудников нет.</p>'
                '<p><a class="btn secondary" href="/employees">← К сотрудникам</a></p>')
        return _render("Архив", body, active="employees", role=request.session.get("role", ""))
    rows = ""
    for e in archived:
        _end = e.contract_end_date.isoformat() if e.contract_end_date else "—"
        rows += (
            f'<div class="card"><b>{html.escape(e.full_name)}</b><br>'
            f'<span class="muted">уволен: {_end}</span><br>'
            f'<a class="btn secondary" href="/employees/{e.id}">Открыть карточку</a></div>'
        )
    body = (
        f'<h1>Архив уволенных ({len(archived)})</h1>'
        f'<p><a class="btn secondary" href="/employees">← К действующим сотрудникам</a></p>'
        f'<section class="grid">{rows}</section>'
    )
    return _render("Архив", body, active="employees", role=request.session.get("role", ""))


# --- Создание нового сотрудника ---------------------------------------------
# Гражданство выбирается из стран ЕАЭС; категория ВЫВОДИТСЯ из него (не отдельное поле),
# чтобы исключить рассинхрон "Казахстан + BELARUS". Беларусь -> BELARUS, остальные -> EAEU.
CITIZENSHIP_OPTIONS = ["Казахстан", "Киргизия", "Армения", "Беларусь"]
from common_utils import (
    CITIZENSHIP_TO_CATEGORY,
    category_for_citizenship as _category_for_citizenship,
)
# Вынесено в common_utils.py — используется и здесь, и в bot.py (там раньше была
# другая, неверная проверка на английское "belarus" вместо русского "Беларусь" —
# несовпадение обнаружилось только при выносе в общий модуль).


def _new_employee_form_html(values: dict, error: str = "") -> str:
    v = values
    cit_sel = v.get("citizenship", "Казахстан")
    opts = "".join(
        f'<option value="{c}"{" selected" if c == cit_sel else ""}>{c}</option>'
        for c in CITIZENSHIP_OPTIONS
    )
    err = f'<div class="warning-banner">Заполни обязательные поля: {html.escape(error)}</div>' if error else ""
    ocr_note = ""
    if v.get("_ocr_filled"):
        ocr_note = ('<div style="margin:12px 0;padding:12px;border:1px solid #e0a800;border-radius:8px;'
                    'background:#fff8e1;color:#7a5c00">⚠ Поля предзаполнены распознаванием (ТЕСТ). '
                    'ОБЯЗАТЕЛЬНО проверьте, особенно ФИО — сверьте с кириллицей на лицевой стороне '
                    'удостоверения (транслит может отличаться от официального написания).</div>')
    return f"""
<h1>Новый сотрудник</h1>
<section class="card-form">
{err}
<fieldset>
<legend>⚡ Распознать из фото удостоверения (ТЕСТ)</legend>
<p class="muted">Фото стороной с MRZ (3 строки латиницей внизу). Поля заполнятся распознанными
данными — проверьте их. Тестовая функция.</p>
<form method="post" action="/employees/new/ocr" enctype="multipart/form-data">
<input type="file" name="photo" accept="image/*" required style="display:block;width:100%;margin:8px 0;padding:10px;border:1px solid #d9dde3;border-radius:8px;background:#fff;font-size:16px">
<button type="submit" class="secondary btn-full">Распознать и заполнить</button>
</form>
</fieldset>
{ocr_note}
<form method="post" action="/employees/new">

<fieldset>
<legend>ФИО (обязательно)</legend>
<input type="text" name="full_name" value="{html.escape(v.get('full_name',''))}">
</fieldset>

<fieldset>
<legend>Гражданство (обязательно)</legend>
<select name="citizenship">{opts}</select>
<p class="muted">Категория учёта определяется автоматически: Беларусь — 90 дней, остальные ЕАЭС — 30.</p>
</fieldset>

<fieldset>
<legend>Дата рождения (обязательно)</legend>
<input type="date" name="birth_date" max="{datetime.now(MSK).date().isoformat()}" value="{html.escape(v.get('birth_date',''))}">
</fieldset>

<fieldset>
<legend>Паспорт (обязательно)</legend>
<label>Серия</label>
<input type="text" name="passport_series" value="{html.escape(v.get('passport_series','ID'))}">
<label>Номер</label>
<input type="text" name="passport_number" value="{html.escape(v.get('passport_number',''))}">
<p class="muted">Для нац. удостоверения РК серия «ID». Для загранпаспорта — впиши свою серию.</p>
</fieldset>

<fieldset>
<legend>ИИН (необязательно)</legend>
<input type="text" name="iin" value="{html.escape(v.get('iin',''))}" placeholder="12 цифр">
<p class="muted">Индивидуальный идентификационный номер (Казахстан). Можно распознать из фото.</p>
</fieldset>
<input type="hidden" name="doc_type" value="{html.escape(v.get('doc_type',''))}">

<fieldset>
<legend>Дата въезда (необязательно)</legend>
<input type="date" name="entry_date" max="{datetime.now(MSK).date().isoformat()}" value="{html.escape(v.get('entry_date',''))}">
<p class="muted">Пусто — если сотрудник ещё не прибыл. Дедлайны начнут считаться после ввода даты въезда.</p>
</fieldset>

<fieldset>
<legend>Дата договора (необязательно)</legend>
<input type="date" name="contract_date" max="{datetime.now(MSK).date().isoformat()}" value="{html.escape(v.get('contract_date',''))}">
</fieldset>

<fieldset>
<legend>Телефон (необязательно)</legend>
<input type="text" name="phone" value="{html.escape(v.get('phone',''))}">
</fieldset>

<button type="submit">Создать</button>
</form>
<p class="muted">Адрес пребывания проставится автоматически (адрес площадки). Согласие на обработку ПД
подтверждается отдельно в карточке — до этого обязательства не создаются.</p>
<a class="btn secondary" href="/employees">← Ко всем сотрудникам</a>
</section>"""


@app.get("/employees/new", response_class=HTMLResponse)
def employee_new_form(request: Request):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    return _render("Новый сотрудник", _new_employee_form_html({}), active="employees", role=request.session.get("role",""))


@app.post("/employees/new/ocr", response_class=HTMLResponse)
async def employee_new_ocr(request: Request, photo: UploadFile = File(...),
                           db: Session = Depends(get_db)):
    """ТЕСТ: распознаёт удостоверение и перерисовывает форму создания с предзаполненными полями
    (ФИО-транслит черновик, дата рождения, номер паспорта, ИИН, гражданство). Ничего не сохраняет
    — только заполняет форму для проверки кадровиком."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    data = await photo.read()
    if not data:
        return HTMLResponse(_render("Новый сотрудник",
            _new_employee_form_html({}, "пустой файл фото"), active="employees",
            role=request.session.get("role", "")))
    ocr = _ocr_id_card(data)
    if not ocr:
        vals = {"_ocr_failed": True}
        note = ("Не удалось распознать MRZ (нужна библиотека passporteye и tesseract, либо фото "
                "нечёткое/не та сторона). Заполните форму вручную.")
        return HTMLResponse(_render("Новый сотрудник",
            _new_employee_form_html(vals, note), active="employees",
            role=request.session.get("role", "")))
    _dtype = ocr.get("doc_type", "id")
    # Серия: для удостоверения (id) — "ID"; для загранпаспорта — обычно нет отдельной серии,
    # весь номер в поле «номер», серию оставляем пустой (кадровик уточнит).
    _series = "ID" if _dtype == "id" else ""
    values = {
        "full_name": ocr.get("full_name_translit", ""),
        "citizenship": ocr.get("citizenship") if ocr.get("citizenship") in CITIZENSHIP_OPTIONS else "Казахстан",
        "birth_date": ocr.get("birth_date", ""),
        "passport_series": _series,
        "passport_number": ocr.get("passport_number", ""),
        "iin": ocr.get("iin", ""),
        "doc_type": _dtype,
        "_ocr_filled": True,
    }
    return HTMLResponse(_render("Новый сотрудник", _new_employee_form_html(values),
        active="employees", role=request.session.get("role", "")))


@app.post("/employees/new")
def employee_create(
    request: Request,
    full_name: str = Form(""),
    citizenship: str = Form("Казахстан"),
    birth_date: str = Form(""),
    passport_series: str = Form("ID"),
    passport_number: str = Form(""),
    iin: str = Form(""),
    doc_type: str = Form(""),
    entry_date: str = Form(""),
    contract_date: str = Form(""),
    phone: str = Form(""),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)

    values = {
        "full_name": full_name, "citizenship": citizenship, "birth_date": birth_date,
        "passport_series": passport_series, "passport_number": passport_number, "iin": iin,
        "entry_date": entry_date, "contract_date": contract_date, "phone": phone,
    }

    def _pdate(s):
        s = (s or "").strip()
        return date.fromisoformat(s) if s else None

    errors = []
    fn = full_name.strip()
    if not fn:
        errors.append("ФИО")
    cit = citizenship.strip()
    if cit not in CITIZENSHIP_OPTIONS:
        errors.append("гражданство")
    try:
        bd = _pdate(birth_date)
    except ValueError:
        bd = None
    if bd is None:
        errors.append("дата рождения")
    ps = passport_series.strip()
    if not ps:
        errors.append("серия паспорта")
    pn = passport_number.strip()
    if not pn:
        errors.append("номер паспорта")

    # необязательные даты: если введены с ошибкой формата — тоже ошибка
    ed = cd = None
    for label, raw, setter in (("дата въезда", entry_date, "ed"), ("дата договора", contract_date, "cd")):
        try:
            val = _pdate(raw)
        except ValueError:
            errors.append(f"{label} (неверный формат)")
            val = None
        if setter == "ed":
            ed = val
        else:
            cd = val

    if errors:
        return HTMLResponse(_render("Новый сотрудник", _new_employee_form_html(values, ", ".join(errors)), active="employees", role=request.session.get("role","")))

    emp = Employee(
        full_name=fn,
        citizenship=cit,
        category=_category_for_citizenship(cit),
        birth_date=bd,
        passport_series=ps,
        passport_number=pn,
        iin=(iin.strip() or None),
        doc_type=(doc_type.strip() or None),
        entry_date=ed,
        contract_date=cd,
        phone=(phone.strip() or None),
        address=SITE_ADDRESS,          # адрес площадки по умолчанию
        # address_since НЕ ставим: первый адрес, не переезд — обязательство регистрации не плодим
        consent_status=ConsentStatus.DRAFT,  # согласие отдельно; до него obligations не создаются
        created_by=_actor_name(request, db),
        language="ru",
    )
    db.add(emp)
    db.commit()
    db.refresh(emp)
    return RedirectResponse(f"/employees/{emp.id}", status_code=303)


def _help(text: str) -> str:
    """Значок-подсказка (i) с раскрытием по тапу. Надёжно на мобильном (нативный <details>,
    без JS и без title, который на телефоне не показывается)."""
    return (
        '<details class="field-help"><summary><span class="i">i</span> подсказка</summary>'
        f'<p>{html.escape(text)}</p></details>'
    )


@app.get("/employees/{employee_id}", response_class=HTMLResponse)
def employee_card(employee_id: str, request: Request, db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)

    _warn_banner = ""
    if request.query_params.get("warn") == "payment":
        _warn_banner = (
            '<div style="margin:12px 0;padding:12px;border:1px solid #e0a800;border-radius:8px;'
            'background:#fff8e1;color:#7a5c00">⚠ Платёжка помечена «требует проверки»: не совпала '
            'фамилия работника и/или сумма (возможно, платёжка не на ту госпошлину). Проверьте '
            'вручную и подтвердите в блоке платёжки (или текст не распознан, если это скан).</div>'
        )

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    today_s = datetime.now(MSK).date().isoformat()

    # Плательщик госпошлины, запомненный на сессию (prefill поля). Экранируем — уходит в HTML-атрибут.
    _last_payer = html.escape(request.session.get("last_payer", ""), quote=True)


    # Роль: прораб — только чтение и скачивание, формы записи ему не показываем (сервер их
    # тоже режет 403, но кнопки-впустую путают). can_write = не прораб.
    _cu = _current_user(request, db)
    can_write = bool(_cu and _cu.role != UserRole.PRORAB)

    # Блок обязательств и сроков в карточке: показывает активные (с кнопкой «Отметить поданным»
    # для тех, что подаются вовне — ЕФС-1, уведомление МВД, регистрация) и выполненные
    # (с датой/автором отметки и кнопкой отмены). Только для пишущих (кадровик/админ).
    _obligations_section = ""
    if can_write:
        # Медосмотр и дактилоскопия исключены — они в объединённой зоне «Медкомиссия и
        # дактилоскопия» выше (со своим закрытием: скан справки / дата). Здесь — остальные.
        _obs = sorted(
            [o for o in emp.obligations if o.is_current
             and o.type not in (ObligationType.MEDICAL_EXAM, ObligationType.DACTYLOSCOPY)],
            key=lambda o: (o.status == ObligationStatus.DONE, o.deadline_date),
        )
        if _obs:
            _ob_rows = ""
            for _o in _obs:
                _olabel = OBLIGATION_LABELS.get(_o.type, _o.type.value)
                _dl = _o.deadline_date.strftime("%d.%m.%Y")
                if _o.status == ObligationStatus.DONE:
                    _who = html.escape(_o.done_by or "—")
                    _when = _o.done_date.strftime("%d.%m.%Y") if _o.done_date else "—"
                    _ob_rows += f'''<div style="margin:8px 0;padding:10px;border:1px solid #e6e9ee;border-radius:8px;background:#f6faf7">
<b>{_olabel}</b> — <span style="color:#1a7f37">выполнено ✓</span><br>
<span class="muted">отмечено {_when}, {_who}. Срок был до {_dl}.</span>
<form method="post" action="/employees/{emp.id}/obligation/reopen" style="margin-top:6px"
onsubmit="return confirm(&#39;Вернуть обязательство в работу?&#39;)">
<input type="hidden" name="obligation_id" value="{_o.id}">
<button type="submit" class="secondary">Отменить отметку</button></form>
</div>'''
                elif _o.status == ObligationStatus.CANCELLED:
                    _ob_rows += f'''<div style="margin:8px 0;padding:10px;border:1px solid #e6e9ee;border-radius:8px;background:#f5f5f5">
<b>{_olabel}</b> — <span class="muted">снято при увольнении</span><br>
<span class="muted">Работник уволен/убыл, обязательство неактуально. Срок был до {_dl}.</span>
</div>'''
                else:
                    _overdue = _o.status == ObligationStatus.OVERDUE
                    _mark = "🔴 просрочено" if _overdue else "🟡 в работе"
                    # Связь с СНИЛС: для ЕФС-1 нужен СНИЛС. Не блокируем (срок ЕФС-1 жёсткий),
                    # но предупреждаем, если СНИЛС нет — форму подать в срок, потом корректировку.
                    _snils_warn = ""
                    if _o.type == ObligationType.EFS1_REPORT and not emp.snils:
                        if emp.snils_appointment_date:
                            _snils_warn = ('<br><span style="color:#c47f00;font-size:13px">⚠ СНИЛС оформляется '
                                           '(запись в СФР ' + emp.snils_appointment_date.strftime("%d.%m.%Y")
                                           + '). Подайте ЕФС-1 в срок, при получении СНИЛС — корректировку.</span>')
                        else:
                            _snils_warn = ('<br><span style="color:#b00;font-size:13px">⚠ СНИЛС отсутствует '
                                           '(нужен для ЕФС-1). Оформите запись в СФР в блоке «СНИЛС». '
                                           'ЕФС-1 подайте в срок, потом корректировку.</span>')
                    _ob_rows += f'''<div style="margin:8px 0;padding:10px;border:1px solid #e6e9ee;border-radius:8px">
<b>{_olabel}</b> — {_mark}, срок до {_dl}{_snils_warn}
<form method="post" action="/employees/{emp.id}/obligation/mark_done" style="margin-top:6px">
<input type="hidden" name="obligation_id" value="{_o.id}">
<button type="submit" class="btn-full">Отметить поданным</button></form>
</div>'''
            _obligations_section = f'''
<fieldset>
<legend>Обязательства и сроки</legend>
<p class="muted">Отметьте «поданным» то, что уже подали в ведомство (ЕФС-1 в СФР, уведомление
МВД, постановка на учёт). Отметка фиксирует дату и кто отметил. Обязательства с собственным
закрытием (медосмотр — результатом, дактилоскопия — датой) закрываются в своих блоках.</p>
{_ob_rows}
</fieldset>'''

    # --- Блок «Удалить событие» — ТОЛЬКО ADMIN (2026-07) --------------------------
    # Мягкое удаление: obligations помечаются is_current=False/CANCELLED, даты-триггеры
    # обнуляются. Из базы ничего не стирается, всё для отката пишется в EventDeletionLog.
    # Семантика дат (маршрут /event/delete):
    #   entry_date       → полный сброс: ВСЕ активные obligations сотрудника → CANCELLED
    #                      (вариант A, согласовано — сотрудник в «чистый лист»);
    #   contract_date    → efs1_report + contract_notice → CANCELLED;
    #   address_since    → registration (по этой дате) → CANCELLED;
    #   dactyloscopy_date→ dactyloscopy ВОЗВРАЩАЕТСЯ в PENDING (дата закрывала обязанность,
    #                      а не порождала — снятие закрытия, не отмена).
    _admin_delete_section = ""
    if _cu and _cu.role == UserRole.ADMIN:
        _del_rows = ""

        # 1) Активные обязательства (кроме уже снятых)
        _del_obs = [o for o in emp.obligations
                    if o.is_current and o.status != ObligationStatus.CANCELLED]
        _del_obs.sort(key=lambda o: o.deadline_date)
        for _do in _del_obs:
            _dolabel = html.escape(OBLIGATION_LABELS.get(_do.type, _do.type.value))
            _dodl = _do.deadline_date.strftime("%d.%m.%Y")
            _del_rows += f'''<div style="margin:6px 0;padding:8px;border:1px solid #f0d6d6;border-radius:8px;background:#fdf7f7">
<b>Обязательство:</b> {_dolabel} <span class="muted">(срок до {_dodl})</span>
<form method="post" action="/employees/{emp.id}/event/delete" style="margin-top:6px"
onsubmit="return confirm(&#39;Удалить обязательство «{_dolabel}»? Оно будет снято (можно откатить по журналу).&#39;)">
<input type="hidden" name="event_kind" value="obligation">
<input type="hidden" name="target" value="{_do.id}">
<button type="submit" class="secondary">Удалить</button></form>
</div>'''

        # 2) Даты-триггеры (только непустые). Для каждой — свой текст подтверждения.
        _date_events = [
            ("entry_date", "Дата въезда", emp.entry_date,
             "Удалить дату въезда? Будут сняты ВСЕ обязательства сотрудника, включая ЕФС-1 (СФР) "
             "и уведомление МВД. Откат — только вручную по журналу. Продолжить?"),
            ("contract_date", "Дата договора", emp.contract_date,
             "Удалить дату договора? Будут сняты ЕФС-1 и уведомление МВД о договоре. Продолжить?"),
            ("address_since", "Дата начала пребывания (адрес)", emp.address_since,
             "Удалить дату начала пребывания? Будет снята постановка на учёт по этому адресу. Продолжить?"),
            ("dactyloscopy_date", "Дата дактилоскопии", emp.dactyloscopy_date,
             "Удалить дату дактилоскопии? Обязанность «дактилоскопия» вернётся в работу. Продолжить?"),
        ]
        for _field, _flabel, _fval, _confirm in _date_events:
            if _fval is None:
                continue
            _fval_s = _fval.strftime("%d.%m.%Y")
            _confirm_js = html.escape(_confirm, quote=True).replace("'", "&#39;")
            _del_rows += f'''<div style="margin:6px 0;padding:8px;border:1px solid #f0d6d6;border-radius:8px;background:#fdf7f7">
<b>Дата:</b> {_flabel} — {_fval_s}
<form method="post" action="/employees/{emp.id}/event/delete" style="margin-top:6px"
onsubmit="return confirm(&#39;{_confirm_js}&#39;)">
<input type="hidden" name="event_kind" value="date">
<input type="hidden" name="target" value="{_field}">
<button type="submit" class="secondary">Удалить</button></form>
</div>'''

        # 3) Полный сброс («чистый лист») — если есть что сбрасывать.
        _has_active_obs = any(o.is_current and o.status != ObligationStatus.CANCELLED
                              for o in emp.obligations)
        _can_full_reset = bool(emp.entry_date) or emp.consent_status == ConsentStatus.CONFIRMED or _has_active_obs
        _full_reset_row = ""
        if _can_full_reset:
            _full_reset_row = f'''<div style="margin:12px 0 2px;padding:10px;border:1px solid #b00;border-radius:8px;background:#fff0f0">
<b style="color:#b00">Полный сброс сотрудника («чистый лист»)</b>
<p class="muted" style="margin:4px 0 8px">Согласие → черновик, дата въезда снимается, все
обязательства снимаются, миграционные сканы уходят в корзину. Для повторного оформления
с нуля.</p>
<form method="post" action="/employees/{emp.id}/event/delete"
onsubmit="return confirm(&#39;ПОЛНЫЙ СБРОС: согласие, дата въезда и ВСЕ обязательства будут сняты, сканы — в корзину. Сотрудник вернётся в исходное состояние. Продолжить?&#39;)">
<input type="hidden" name="event_kind" value="full_reset">
<input type="hidden" name="target" value="full_reset">
<button type="submit" class="secondary" style="border-color:#b00;color:#b00">Полный сброс</button></form>
</div>'''

        if _del_rows or _full_reset_row:
            _admin_delete_section = f'''
<fieldset style="border-color:#e0a0a0">
<legend style="color:#b00">⚠ Удаление событий (админ)</legend>
<details>
<summary style="cursor:pointer;color:#b00">Показать удаляемые события</summary>
<p class="muted" style="margin-top:8px">Удаление мягкое: запись снимается, но сохраняется в
журнале для отката. Даты обнуляются, обязательства помечаются снятыми. Действие только для
администратора.</p>
{_del_rows}
{_full_reset_row}
</details>
</fieldset>'''

    # Сброс согласия ПД — рабочий инструмент кадровика (и админа). Отдельно от админского
    # блока удаления: доступен всем пишущим, но только когда согласие подтверждено.
    _consent_reset_section = ""
    if can_write and emp.consent_status == ConsentStatus.CONFIRMED:
        _consent_reset_section = f'''
<fieldset>
<legend>Сброс согласия ПД</legend>
<p class="muted">Вернуть согласие в статус «черновик» (например, при ошибочном подтверждении).
Запись о согласии остаётся в журнале; при повторном подтверждении фиксируется заново.</p>
<form method="post" action="/employees/{emp.id}/event/delete"
onsubmit="return confirm(&#39;Сбросить согласие ПД в черновик?&#39;)">
<input type="hidden" name="event_kind" value="consent">
<input type="hidden" name="target" value="consent">
<button type="submit" class="secondary">Сбросить согласие</button></form>
</fieldset>'''

    # Секция сканов для пакета Госуслуг — только для пишущих (кадровик/админ), паспортные
    # данные прорабу недоступны. Показывает, какие сканы загружены, даёт загрузить/скачать/удалить.
    _scans_section = ""
    # Прораб: только просмотр и скачивание сканов работника (без загрузки/удаления/пакета).
    if _cu and _cu.role == UserRole.PRORAB:
        _pr_present = _s3_list_for_employee(emp.id)
        _pr_rows = ""
        for _st, _label in SCAN_TYPES.items():
            if _st == "medical_certificate":
                continue  # справка — в объединённой зоне
            _info = _pr_present.get(_st) or {}
            if not _info.get("present"):
                continue
            _pr_rows += (
                '<div style="margin:12px 0;padding:12px;border:1px solid #e6e9ee;border-radius:8px">'
                f'<b>{_label}</b> — <span style="color:#1a7f37">загружен ✓</span>'
                '<div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap">'
                f'<a href="/employees/{emp.id}/scan/view?scan_type={_st}" target="_blank" class="btn secondary" style="display:inline-block">Просмотр</a>'
                f'<form method="post" action="/employees/{emp.id}/scan/download" style="display:inline">'
                f'<input type="hidden" name="scan_type" value="{_st}">'
                '<button type="submit" class="secondary">Скачать</button></form>'
                '</div></div>'
            )
        if not _pr_rows:
            _pr_rows = '<p class="muted">Документы ещё не загружены.</p>'
        _scans_section = (
            '<fieldset id="scans-section">'
            '<legend>Документы работника</legend>'
            '<p class="muted">Просмотр и скачивание загруженных сканов.</p>'
            '<div class="scans-grid">'
            + _pr_rows +
            '</div></fieldset>'
        )
    if can_write:
        # Считаем наличие сканов ОДИН раз (персональные + общие) и переиспользуем ниже и в
        # _package_missing — иначе S3 опрашивается дважды за рендер карточки (было медленно).
        _present = _s3_list_for_employee(emp.id)
        _common_present = _s3_list_common()
        _rows = ""
        for _st, _label in SCAN_TYPES.items():
            # Справка медкомиссии грузится в объединённой зоне «Медкомиссия и дактилоскопия»,
            # в общей секции сканов не дублируем (остаётся обязательной для пакета — см. _package_missing).
            if _st == "medical_certificate":
                continue
            _info = _present.get(_st) or {}
            _has = bool(_info.get("present"))
            _check = bool(_info.get("check"))
            _status = ('<span style="display:inline-block;padding:2px 10px;border-radius:12px;background:#e3f5e9;color:#1a7f37;font-weight:600;font-size:14px">✓ загружен</span>' if _has else '<span style="display:inline-block;padding:2px 10px;border-radius:12px;background:#f0f1f3;color:#889;font-size:14px">нет</span>')
            # Скан помечен «требует проверки» (косячная платёжка: фамилия/сумма не сошлись).
            _check_row = ""
            if _has and _check:
                _check_row = f'''<div style="margin-top:8px;padding:8px;border:1px solid #e0a800;border-radius:8px;background:#fff8e1;color:#7a5c00">
⚠ Требует проверки: фамилия или сумма в платёжке не совпали. Проверьте вручную и подтвердите.
<form method="post" action="/employees/{emp.id}/scan/confirm" style="margin-top:6px">
<input type="hidden" name="scan_type" value="{_st}">
<button type="submit" class="secondary">Подтвердить (снять метку)</button></form>
</div>'''
            _actions = ""
            if _has:
                _actions = f'''<a href="/employees/{emp.id}/scan/view?scan_type={_st}" target="_blank" class="btn secondary" style="display:inline-block">Просмотр</a>
<form method="post" action="/employees/{emp.id}/scan/download" style="display:inline">
<input type="hidden" name="scan_type" value="{_st}">
<button type="submit" class="secondary">Скачать</button></form>
<form method="post" action="/employees/{emp.id}/scan/delete" style="display:inline"
onsubmit="return confirm(&#39;Удалить скан?&#39;)">
<input type="hidden" name="scan_type" value="{_st}">
<button type="submit" class="secondary">Удалить</button></form>'''
            _border = "#e0a800" if (_has and _check) else "#e6e9ee"
            _rows += f'''<div style="margin:12px 0;padding:12px;border:1px solid {_border};border-radius:8px">
<b>{_label}</b> — {_status}
{_check_row}
<form method="post" action="/employees/{emp.id}/scan/upload" enctype="multipart/form-data" style="margin-top:8px">
<input type="hidden" name="scan_type" value="{_st}">
<input type="file" name="files" accept="application/pdf,image/*" multiple required style="display:block;width:100%;margin:8px 0;padding:10px;border:1px solid #d9dde3;border-radius:8px;background:#fff;font-size:16px">
<p class="muted" style="margin:4px 0 0;font-size:13px">Можно выбрать несколько файлов сразу (напр. 2 фото удостоверения — лицевая и оборотная) — склеятся в один PDF.</p>
<button type="submit" class="btn-full btn-upload">Загрузить</button></form>
<div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap">{_actions}</div>
</div>'''
        # Проверка комплектности пакета: если чего-то нет — показываем список, кнопку выгрузки
        # не даём (пакет не должен выходить дырявым и уходить на Госуслуги с отказом).
        _missing = _package_missing(emp, present=_present, common=_common_present)
        if _missing:
            _pkg_block = ('<p class="muted">Пакет пока неполный. Не хватает:</p><ul>'
                          + "".join(f"<li>{html.escape(_m)}</li>" for _m in _missing)
                          + "</ul><p class=\"muted\">Догрузите недостающее (общие документы — на "
                          '<a href="/common-docs">странице общих документов</a>).</p>')
        else:
            _pkg_block = f'''<p style="color:#1a7f37">Пакет полный ✓</p>
<form method="post" action="/employees/{emp.id}/package">
<button type="submit" class="btn-full">Выгрузить пакет (ZIP) для Госуслуг</button>
</form>'''

        # Тип документа: passport явно, или (пусто И серия не "ID") -> считаем паспортом.
        _is_passport = (emp.doc_type == "passport") or (
            not emp.doc_type and (emp.passport_series or "").strip().upper() != "ID"
        )
        _passport_pages_block = ""
        if _is_passport:
            _checked = "checked" if getattr(emp, "passport_all_pages", False) else ""
            _passport_pages_block = f'''<div style="margin:12px 0;padding:12px;border:1px solid #e0a800;border-radius:8px;background:#fff8e1">
<b>Загранпаспорт:</b> нужны сканы ВСЕХ страниц (не только разворот с фото).
<form method="post" action="/employees/{emp.id}/passport_pages" style="margin-top:8px">
<label style="display:flex;align-items:center;gap:8px">
<input type="checkbox" name="all_pages" {_checked} onchange="this.form.submit()">
Все страницы паспорта загружены
</label></form></div>'''
        _scans_section = f'''
<fieldset id="scans-section">
<legend>Пакет для Госуслуг</legend>
<p class="muted">Персональные сканы работника: паспорт, миграционная карта, исполненная платёжка
(подтверждение оплаты госпошлины из банка). PDF или фото, до 15 МБ. Общие документы (паспорт
директора, основание на адрес) — на <a href="/common-docs">отдельной странице</a>.</p>
<div class="scans-grid">
{_rows}
</div>
{_passport_pages_block}
<hr>
{_pkg_block}
</fieldset>'''

    # Тексты справок у полей карточки (значок i с раскрытием). Коротко: что это и что делает.
    help_entry = _help(
        "Дата пересечения ГРАНИЦЫ РФ по миграционной карте/штампу — когда иностранец въехал в "
        "Россию. НЕ дата переезда внутри страны (например, из Москвы в Мурманск): для переезда "
        "есть поле «Дата, с которой действует адрес» ниже. Пример: въехал в РФ 03.04, приехал на "
        "объект 15.06 — сюда ставится 03.04. От этой даты идут сроки: регистрация (30 дней ЕАЭС), "
        "медосмотр, дактилоскопия. Пусто — ещё не прибыл в РФ, сроки не идут."
    )
    help_dact = _help(
        "Дата прохождения дактилоскопии и фотографирования («грин карта»). Заполнение закрывает "
        "обязанность. Пусто — обязанность горит от даты въезда + 30 дней. Изменение потребует "
        "подтверждения (закроет обязательство)."
    )
    help_country = _help(
        "Государство, из которого сотрудник въехал в РФ. Информационное поле, на сроки не влияет."
    )
    help_address = _help(
        "Фактическое место пребывания в РФ (адрес объекта/общежития). «Дата, с которой действует "
        "адрес» — день прибытия на ЭТО место (переезд внутри РФ), НЕ дата въезда в страну. Пример: "
        "въехал в РФ 03.04, прибыл на объект в Белокаменке 15.06 — здесь 15.06. "
        "СРОК ПОСТАНОВКИ: для граждан Казахстана (ЕАЭС) без ВНЖ и без статуса ВКС — 30 суток с "
        "даты въезда в РФ (Договор о ЕАЭС, п.6 ст.97). При переезде в другой регион встать на "
        "учёт по новому адресу нужно, но общий 30-дневный режим сохраняется. Правило «7 дней при "
        "смене региона» к ним НЕ относится — оно только для ВКС и обладателей ВНЖ. Смена адреса "
        "создаёт новое обязательство постановки; первый ввод адреса обязательство не создаёт."
    )
    help_contract_date = _help(
        "Дата заключения трудового договора. Это поле (слева) сохраняется кнопкой «Сохранить» и "
        "создаёт обязательства: уведомление МВД (3 рабочих дня) и ЕФС-1 (1 рабочий день). "
        "Отдельная «Дата договора» справа, в блоке заключения — это дата, с которой сформируется "
        "документ при нажатии «Заключить»; она не сохраняется сама по себе, а применяется в момент "
        "заключения. Если нужно просто зафиксировать дату — используйте это поле слева и «Сохранить»."
    )
    help_contract = _help(
        "Заключение проставит дату договора в карточку и создаст обязательства (уведомление МВД, "
        "ЕФС-1). Скачается .docx с реквизитами ООО «ТРЕСТСТРОЙМОНТАЖ». Номер договора — из "
        "табельного. «Кем/когда выдан паспорт» в документе — прочерк, заполняется вручную."
    )
    help_position = _help(
        "Должность работника для договора. Идёт в документ как есть, не проверяется. По умолчанию "
        "«Монтажник» — измените под конкретного работника."
    )
    help_salary = _help(
        "Оклад в рублях (только оклад, без районного коэффициента — он добавляется в договоре "
        "отдельно, 1,500). Идёт в документ как есть, не проверяется."
    )
    help_status = _help(
        "Первичный учёт — сотрудник впервые встаёт на учёт в РФ: регистрация, медосмотр и "
        "дактилоскопия считаются от даты въезда. Ранее стоял на учёте — приехал на вахту из "
        "другого региона РФ: регистрация от прибытия (даты адреса), медосмотр и дактилоскопия "
        "заново НЕ требуются. Пустой статус блокирует создание обязательств."
    )

    # Блок трудового договора. Предусловия: (1) задан статус учёта — от него зависят
    # обязательства, договор без статуса заключать нельзя; (2) есть табельный номер — иначе
    # номер договора соберётся как "БК-ПСМ-" без хвоста. Проверяем статус первым.
    if emp.registration_status is None:
        contract_block = (
            '<p class="muted">Нельзя заключить договор: не задан статус миграционного учёта. '
            'Сначала выберите статус выше — от него зависит расчёт обязательств.</p>'
        )
    elif not (emp.tab_number or "").strip():
        contract_block = (
            '<p class="muted">Нельзя заключить договор: у сотрудника нет табельного номера. '
            'Он идёт в номер договора (' + CONTRACT_NUMBER_PREFIX + '{таб}). '
            'Присвойте табельный номер.</p>'
        )
    else:
        _contract_no = CONTRACT_NUMBER_PREFIX + emp.tab_number.strip()
        # Кнопка отмены — только если договор уже заключён (стоит дата). Отмена откатывает
        # contract_date и снимает НЕ выполненные обязательства от договора (МВД/ЕФС-1);
        # выполненные (DONE) сохраняются как след исполнения.
        # Скачать и отменить доступны только ПОСЛЕ заключения (стоит contract_date).
        if emp.contract_date is not None:
            # Отмена договора: админ — всегда; кадровик — только в день заключения (свежая ошибка
            # ввода). После — только админ. cancel-роут это тоже проверяет на сервере.
            _cancel_allowed = (_cu and _cu.role == UserRole.ADMIN) or (
                _cu and _cu.role == UserRole.KADROVIK
                and emp.contract_date == date.today()
            )
            _cancel = ""
            if _cancel_allowed:
                _cancel = f'''<form method="post" action="/employees/{emp.id}/labor_contract/cancel"
onsubmit="return confirm(&#39;Отменить договор? Дата договора будет снята, а незакрытые обязательства (уведомление МВД, ЕФС-1) удалены. Выполненные останутся.&#39;)">
<button type="submit" class="secondary btn-full">Отменить договор</button>
</form>'''
            elif _cu and _cu.role == UserRole.KADROVIK:
                _cancel = '<p class="muted">Отмена договора доступна только в день заключения. Позже — обратитесь к администратору.</p>'

            # Секция увольнения — только для пишущих (кадровик/админ), после заключения договора.
            _termination = ""
            if can_write:
                if emp.contract_end_date is not None:
                    _termination = f'''
<hr>
<p><b>Увольнение оформлено:</b> {emp.contract_end_date.strftime("%d.%m.%Y")}.</p>
<p class="muted">Созданы обязательства: уведомление МВД о расторжении (3 раб. дня) и снятие с
учёта / уведомление об убытии (7 раб. дней). Если дата в будущем — обязательства включатся в день увольнения.</p>
<form method="post" action="/employees/{emp.id}/termination_notice">
<button type="submit" class="secondary btn-full">Уведомление о расторжении — скачать</button>
</form>
<form method="post" action="/employees/{emp.id}/departure_notice">
<button type="submit" class="secondary btn-full">Уведомление об убытии — скачать</button>
</form>
<form method="post" action="/employees/{emp.id}/termination/cancel"
onsubmit="return confirm(&#39;Отменить оформление увольнения? Дата увольнения снимется, связанные обязательства удалятся.&#39;)">
<button type="submit" class="secondary btn-full">Отменить увольнение</button>
</form>'''
                else:
                    _termination = f'''
<hr>
<p><b>Увольнение / расторжение договора</b></p>
<p class="muted">Отдельно от «Отменить договор»: отмена = договора не было (ошибка ввода);
увольнение = работник был трудоустроён, история сохраняется. Дата не раньше даты договора.
Будущую дату можно (обязательства включатся в день увольнения).</p>
<form method="post" action="/employees/{emp.id}/termination">
<label>Дата увольнения (расторжения договора)</label>
<input type="date" name="termination_date" min="{emp.contract_date.isoformat()}" value="{today_s}">
<label>Основание расторжения</label>
<select name="basis" id="basis_select" onchange="_toggleBasisNote()">
<option value="по собственному желанию">По собственному желанию</option>
<option value="по инициативе работодателя">По инициативе работодателя</option>
<option value="по соглашению сторон">По соглашению сторон</option>
<option value="истечение срока договора">Истечение срока договора</option>
<option value="иное">Иное (указать в примечании)</option>
</select>
<div id="basis_note_wrap" style="display:none">
<label>Примечание к основанию</label>
<input type="text" name="basis_note" placeholder="">
</div>
<button type="submit" class="btn-full">Оформить увольнение</button>
</form>
<script>
function _toggleBasisNote(){{
  var sel = document.getElementById('basis_select');
  var wrap = document.getElementById('basis_note_wrap');
  if (sel && wrap) wrap.style.display = (sel.value === 'иное') ? 'block' : 'none';
}}
_toggleBasisNote();
</script>'''

            _post = f'''
<p class="muted">Договор заключён {emp.contract_date.strftime("%d.%m.%Y")}.</p>
<form method="post" action="/employees/{emp.id}/labor_contract/download">
<input type="hidden" name="position" value="Монтажник">
<input type="hidden" name="salary" value="30000">
<button type="submit" class="btn-full">Скачать .docx</button>
</form>
<form method="post" action="/employees/{emp.id}/labor_contract/download_pdf">
<input type="hidden" name="position" value="Монтажник">
<input type="hidden" name="salary" value="30000">
<button type="submit" class="secondary btn-full">Скачать .pdf (для Госуслуг)</button>
</form>
{_cancel}
{_termination}'''
        else:
            _post = ""
        # До заключения: форма с предпросмотром и заключением. Предпросмотр (formaction preview)
        # ничего не пишет — только показывает HTML по введённым данным. Заключение пишет дату
        # и создаёт обязательства. Скачивание доступно только после заключения (блок _post).
        # Форма заключения показывается ТОЛЬКО пока договор НЕ заключён. После заключения
        # (contract_date задана) видны лишь скачивание и «Отменить договор» (блок _post) —
        # чтобы нельзя было заключить повторно. Отмена договора вернёт форму заключения.
        if emp.contract_date is None:
            _conclude_form = f"""
<p class="muted">Номер договора: {_contract_no}. Дата по умолчанию — сегодня; можно изменить.</p>
{help_contract}
<form method="post" action="/employees/{emp.id}/labor_contract">
<label>Должность</label>
<input type="text" name="position" value="Монтажник">
{help_position}
<label>Оклад (руб.)</label>
<input type="text" name="salary" value="30000">
{help_salary}
<label>Дата договора</label>
<input type="date" name="contract_date" max="{today_s}" value="{today_s}">
<button type="submit" formaction="/employees/{emp.id}/labor_contract/preview" class="secondary btn-full">Предпросмотр</button>
<button type="submit" class="btn-full">Заключить трудовой договор</button>
</form>"""
        else:
            _conclude_form = ""
        contract_block = f"""{_conclude_form}
{_post}"""

    # Секция статуса учёта. Отдельная форма со своим confirm — смена пересоздаёт
    # обязательства по новому статусу, это не рутинное сохранение, мешать с saveform нельзя.
    _rs = emp.registration_status
    _rs_val = _rs.value if _rs is not None else ""
    def _opt(v, label):
        sel = " selected" if _rs_val == v else ""
        return f'<option value="{v}"{sel}>{label}</option>'
    if _rs is None:
        _status_warn = (
            '<p><span class="badge red">Статус учёта не задан</span></p>'
            '<p class="muted">Пока статус не выбран, обязательства НЕ создаются '
            '(ни регистрация, ни медосмотр, ни уведомления). Выберите статус.</p>'
        )
    else:
        _status_warn = ""
    _confirm = (
        "return confirm(&#39;Сменить статус учёта? Обязательства будут пересозданы: "
        "лишние незакрытые удалены, недостающие добавлены. Выполненные останутся.&#39;)"
    )
    _sel_empty = " selected" if _rs_val == "" else ""
    status_block = (
        _status_warn + help_status
        + f'<form method="post" action="/employees/{emp.id}/registration_status" onsubmit="{_confirm}">'
        + '<label>Статус миграционного учёта</label>'
        + '<select name="registration_status">'
        + f'<option value=""{_sel_empty}>— не задан —</option>'
        + _opt("primary", "Первичный учёт (сроки от даты въезда)")
        + _opt("prior", "Ранее стоял на учёте в РФ (сроки от прибытия на вахту)")
        + '</select>'
        + '<button type="submit" class="btn-full">Сохранить статус</button>'
        + '</form>'
    )

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

    # Левая колонка: пишущим — форма сохранения; прорабу — те же данные текстом (только чтение).
    if can_write:
        left_col = f"""
<p class="muted">Заполни известные поля и нажми одну кнопку внизу. Пустые поля не трогаются.</p>
<form id="saveform" method="post" action="/employees/{emp.id}/save">
<input type="hidden" name="confirmed" value="">
<fieldset>
<legend>Дата въезда</legend>
<input type="date" name="entry_date" max="{today_s}"
value="{emp.entry_date.isoformat() if emp.entry_date else ''}">
{help_entry}
</fieldset>
<fieldset>
<legend>Место пребывания</legend>
<p class="muted">Текущий адрес: {emp.address or "не указан"}</p>
<label>Адрес</label>
<input type="text" name="address" data-orig="{emp.address or ''}" value="{emp.address or ''}">
<label>Дата, с которой действует этот адрес</label>
<input type="date" name="address_since" max="{today_s}" value="{emp.address_since.isoformat() if emp.address_since else today_s}">
<label>Срок регистрации до (из уведомления Госуслуг)</label>
<input type="date" name="registration_valid_until" value="{emp.registration_valid_until.isoformat() if emp.registration_valid_until else ''}">
<p class="muted">Дата окончания срока пребывания — та, что стоит в отрывной части уведомления
(«срок пребывания до»). На Госуслугах даты начала нет, печатается только эта. Справочно, для
напоминания о продлении.</p>
{help_address}
</fieldset>
<fieldset>
<legend>Дата договора</legend>
<input type="date" name="contract_date" max="{today_s}"
value="{emp.contract_date.isoformat() if emp.contract_date else ''}">
{help_contract_date}
</fieldset>
<fieldset>
<legend>Должность и подразделение</legend>
<label>Должность (профессия)</label>
<input type="text" name="position" value="{html.escape(emp.position or '')}" placeholder="например: Бетонщик">
<label>Подразделение</label>
<input type="text" name="subdivision" value="{html.escape(emp.subdivision or '')}" placeholder="например: ОС">
<p class="muted">Подставляется автоматически в журналы инструктажей (графы по ГОСТ) —
заполнить один раз здесь, дальше не нужно вписывать вручную в каждой распечатке.</p>
</fieldset>
<button type="submit" class="btn-full">Сохранить</button>
</form>"""
    else:
        _d = lambda v: v if v else "—"
        left_col = f"""
<p class="muted">Режим чтения. Изменение данных доступно кадровику.</p>
<fieldset><legend>Дата въезда</legend><p>{_d(emp.entry_date.isoformat() if emp.entry_date else None)}</p></fieldset>
<fieldset><legend>Дактилоскопия</legend><p>{_d(emp.dactyloscopy_date.isoformat() if emp.dactyloscopy_date else None)}</p></fieldset>
<fieldset><legend>Место пребывания</legend><p>{_d(emp.address)}</p></fieldset>
<fieldset><legend>Дата договора</legend><p>{_d(emp.contract_date.isoformat() if emp.contract_date else None)}</p></fieldset>"""

    # Прораб — только чтение: убираем формы записи из правой колонки. Согласие/статус
    # показываем текстом-статусом; договор оставляем ТОЛЬКО предпросмотр и скачивание
    # (генерация документа — разрешена прорабу), но без «Заключить»/«Отменить».
    if not can_write:
        # согласие: только статус, без кнопки подтверждения
        if emp.consent_status == ConsentStatus.CONFIRMED:
            consent_block = '<p><span class="badge green">Согласие подтверждено</span></p>'
        else:
            consent_block = '<p class="muted">Согласие не подтверждено. Подтверждение доступно кадровику.</p>'
        # статус: только текущее значение, без формы смены
        _rs_ru = {"primary": "Первичный учёт", "prior": "Ранее стоял на учёте в РФ"}.get(
            (emp.registration_status.value if emp.registration_status else ""), "не задан")
        status_block = f'<p>Статус: <b>{_rs_ru}</b></p><p class="muted">Изменение доступно кадровику.</p>'
        # договор: предпросмотр + скачивание (если заключён), без заключения/отмены
        if (emp.tab_number or "").strip() and emp.registration_status is not None:
            _dl = ""
            if emp.contract_date is not None:
                _dl = f"""<p class="muted">Договор заключён {emp.contract_date.strftime("%d.%m.%Y")}.</p>
<form method="post" action="/employees/{emp.id}/labor_contract/download">
<input type="hidden" name="position" value="Монтажник">
<input type="hidden" name="salary" value="30000">
<button type="submit" class="btn-full">Скачать .docx</button>
</form>"""
            contract_block = f"""
<p class="muted">Режим чтения. Заключение договора доступно кадровику.</p>
<form method="post" action="/employees/{emp.id}/labor_contract/preview">
<input type="hidden" name="position" value="Монтажник">
<input type="hidden" name="salary" value="30000">
<input type="hidden" name="contract_date" value="{today_s}">
<button type="submit" class="secondary btn-full">Предпросмотр договора</button>
</form>
{_dl}"""
        else:
            contract_block = '<p class="muted">Договор недоступен: не задан статус учёта или табельный номер.</p>'

    # === Объединённая зона: Медкомиссия и дактилоскопия ===
    from models import Referral, ExamStatus
    # законные дедлайны (30 дней от въезда) из obligations — показать в зоне, т.к. из блока
    # обязательств медосмотр/дактилоскопия убраны.
    _med_ob = next((o for o in emp.obligations if o.is_current and o.type == ObligationType.MEDICAL_EXAM), None)
    _dact_ob = next((o for o in emp.obligations if o.is_current and o.type == ObligationType.DACTYLOSCOPY), None)
    _med_legal = f' · срок по закону до {_med_ob.deadline_date.strftime("%d.%m.%Y")}' if _med_ob else ''
    _dact_legal = f' · срок по закону до {_dact_ob.deadline_date.strftime("%d.%m.%Y")}' if _dact_ob else ''
    _ref = db.scalars(
        select(Referral).where(Referral.employee_id == emp.id)
        .order_by(Referral.referral_date.desc())
    ).first()
    if _ref is None:
        if emp.entry_date:
            _since = (date.today() - emp.entry_date).days
            if _since >= 5:
                _med_html = f'<b style="color:#c47f00">Пора выдать направление</b> (прошло {_since} дн. от въезда)'
            else:
                _med_html = f'<span class="muted">Направление ещё не требуется (прошло {_since} дн. из 5)</span>'
        else:
            _med_html = '<span class="muted">Ожидает прибытия — сроки не идут</span>'
        _med_html += ' · <a href="/medical">выписать направление</a>'
        _med_done = False
    elif _ref.exam_status == ExamStatus.COMPLETED:
        _med_html = '<b style="color:#1a7f37">Медкомиссия пройдена</b>'
        if _ref.result_date:
            _med_html += f' <span class="muted">({_ref.result_date.isoformat()})</span>'
        _med_done = True
    else:
        _md = (date.today() - _ref.referral_date).days
        if _md > 14:
            _med_html = f'<b style="color:#b00">Справка просрочена</b> (прошло {_md} дн., внутр. срок 14)'
        elif _md >= 10:
            _med_html = f'<b style="color:#c47f00">Ждём справку</b> (прошло {_md} дн. из 14)'
        else:
            _med_html = f'<span>Направлен {_ref.referral_date.isoformat()}, прошло {_md} дн. из 14</span>'
        _med_done = False
    # Форму загрузки справки показываем, пока СКАНА СПРАВКИ НЕТ — независимо от статуса медкомиссии
    # (иначе тупик: медкомиссия «пройдена», а скана нет и загрузить негде; пакет требует справку).
    if emp.registration_status == RegistrationStatus.PRIOR:
        _med_html = ('<b style="color:#1a7f37">Грин-карта с прошлого въезда</b> '
                     '<span class="muted">— медкомиссия и справка не требуются (повторный въезд)</span>')
    _cert_uploaded = False
    try:
        _cert_uploaded = bool((_s3_list_for_employee(emp.id).get("medical_certificate") or {}).get("present"))
    except Exception:
        _cert_uploaded = False
    # Повторный въезд (PRIOR): грин-карта с прошлого раза, медкомиссия и дактилоскопия
    # заново не проходятся — сканы справок для пакета не требуются (правило согласовано 2026-07).
    _is_prior = emp.registration_status == RegistrationStatus.PRIOR
    # Подсказка: медкомиссия пройдена, но скана справки нет — он нужен для пакета Госуслуг.
    _cert_hint = ""
    if _med_done and not _cert_uploaded and not _is_prior:
        _cert_hint = '<br><span style="color:#c47f00;font-size:13px">⚠ Медкомиссия пройдена, но скан справки не загружен — он нужен для пакета Госуслуг. Загрузите ниже.</span>'
    _med_upload = ""
    if not _cert_uploaded and not _is_prior:
        _med_upload = f'''<form method="post" action="/employees/{emp.id}/scan/upload" enctype="multipart/form-data" style="margin-top:8px">
<input type="hidden" name="scan_type" value="medical_certificate">
<input type="file" name="files" accept="application/pdf,image/*" multiple required style="display:block;width:100%;margin:6px 0;padding:8px;border:1px solid #d9dde3;border-radius:8px;background:#fff;font-size:15px">
<button type="submit" class="btn-full">Загрузить справку (закроет медкомиссию)</button></form>'''
    _today_s_dact = date.today().isoformat()
    _dact_action = f'/employees/{emp.id}/dactyloscopy_date'
    if emp.dactyloscopy_date:
        _dv = emp.dactyloscopy_date.isoformat()
        _dact_form = (
            '<form method="post" action="' + _dact_action + '" style="margin-top:8px">'
            '<label style="font-size:13px" class="muted">Изменить дату:</label>'
            '<input type="date" name="dactyloscopy_date" max="' + _today_s_dact + '" value="' + _dv + '" '
            'style="display:block;margin:4px 0;padding:8px;border:1px solid #d9dde3;border-radius:8px">'
            '<button type="submit" class="secondary">Сохранить дату</button></form>'
        )
        _dact_html = (
            '<b style="color:#1a7f37">Дактилоскопия сделана</b> '
            '<span class="muted">(' + _dv + ')</span>' + _dact_form
        )
    else:
        _dact_seq = "" if _med_done else ' <span class="muted">(обычно после медкомиссии)</span>'
        _dact_form = (
            '<form method="post" action="' + _dact_action + '" style="margin-top:8px">'
            '<label style="font-size:13px" class="muted">Дата прохождения (заполнение закроет обязательство):</label>'
            '<input type="date" name="dactyloscopy_date" max="' + _today_s_dact + '" '
            'style="display:block;margin:4px 0;padding:8px;border:1px solid #d9dde3;border-radius:8px">'
            '<button type="submit" class="btn-full">Сохранить дату дактилоскопии</button></form>'
        )
        _dact_html = (
            '<b style="color:#c47f00">Дактилоскопия не сделана</b>' + _dact_seq + '<br>' + _dact_form
        )
    if emp.registration_status == RegistrationStatus.PRIOR:
        _dact_html = ('<b style="color:#1a7f37">Грин-карта с прошлого въезда</b> '
                      '<span class="muted">— дактилоскопия не требуется (повторный въезд)</span>')
    _medzone_section = f'''
<fieldset>
<legend>Медкомиссия и дактилоскопия</legend>
<div style="padding:10px 0;border-bottom:1px solid #eee">
<div style="font-size:13px;color:#889;text-transform:uppercase;letter-spacing:.5px">Медкомиссия{_med_legal}</div>
{_med_html}{_cert_hint}
{_med_upload}
</div>
<div style="padding:10px 0">
<div style="font-size:13px;color:#889;text-transform:uppercase;letter-spacing:.5px">Дактилоскопия{_dact_legal}</div>
{_dact_html}
</div>
</fieldset>'''

    # === Блок СНИЛС (нужен для корректного ЕФС-1) ===
    _snils_action = '/employees/' + emp.id + '/snils'
    if emp.snils:
        _snils_status = '<b style="color:#1a7f37">СНИЛС: ' + html.escape(emp.snils) + '</b>'
        _snils_form_extra = ''
    else:
        if emp.snils_appointment_date:
            _proc = {"new": "первичное получение", "merge": "объединение дублей"}.get(emp.snils_procedure or "", "получение")
            _appt_passed = emp.snils_appointment_date < date.today()
            if _appt_passed:
                _snils_status = ('<b style="color:#b00">СНИЛС: дата записи прошла ('
                                 + emp.snils_appointment_date.strftime("%d.%m.%Y") + ', ' + _proc
                                 + '), а номер не внесён — проверьте, получен ли СНИЛС</b>')
            else:
                _snils_status = ('<b style="color:#c47f00">СНИЛС оформляется</b> — запись в СФР на '
                                 + emp.snils_appointment_date.strftime("%d.%m.%Y") + ' (' + _proc + ')')
        else:
            _snils_status = '<b style="color:#b00">СНИЛС отсутствует</b> — нужна запись в СФР'
        _sel_new = ' selected' if emp.snils_procedure == "new" else ''
        _sel_merge = ' selected' if emp.snils_procedure == "merge" else ''
        _appt_val = emp.snils_appointment_date.isoformat() if emp.snils_appointment_date else ''
        _snils_form_extra = (
            '<label style="font-size:13px" class="muted">Вид процедуры:</label>'
            '<select name="snils_procedure" style="display:block;margin:4px 0;padding:8px;border:1px solid #d9dde3;border-radius:8px">'
            '<option value=""></option>'
            '<option value="new"' + _sel_new + '>Первичное получение</option>'
            '<option value="merge"' + _sel_merge + '>Объединение дублей (было несколько СНИЛС)</option>'
            '</select>'
            '<label style="font-size:13px" class="muted">Дата записи в СФР:</label>'
            '<input type="date" name="snils_appointment_date" value="' + _appt_val + '" '
            'style="display:block;margin:4px 0;padding:8px;border:1px solid #d9dde3;border-radius:8px">'
        )
    _snils_section = (
        '<fieldset><legend>СНИЛС</legend>' + _snils_status +
        '<form method="post" action="' + _snils_action + '" style="margin-top:8px">'
        '<label style="font-size:13px" class="muted">Номер СНИЛС (если есть):</label>'
        '<input type="text" name="snils" value="' + html.escape(emp.snils or "") + '" placeholder="XXX-XXX-XXX YY" '
        'style="display:block;margin:4px 0;padding:8px;border:1px solid #d9dde3;border-radius:8px">'
        + _snils_form_extra +
        '<button type="submit" class="btn-full">Сохранить СНИЛС</button></form>'
        '<p class="muted" style="font-size:13px">Нужен для ЕФС-1. Форму подавайте в срок '
        '(1 раб. день от договора) даже без СНИЛС, при получении — корректировка.</p></fieldset>'
    )

    body = f"""
{_warn_banner}
<h1>{emp.full_name}</h1>
<section class="card-form">
<div class="card-cols">
<div class="card-col">
{left_col}
</div>
<div class="card-col">
<fieldset>
<legend>Согласие на обработку ПД</legend>
{consent_block}
</fieldset>

<fieldset>
<legend>Статус миграционного учёта</legend>
{status_block}
</fieldset>

<fieldset>
<legend>Трудовой договор ({EMPLOYER_NAME_SHORT})</legend>
{contract_block}
</fieldset>

<fieldset>
<legend>Госпошлина</legend>
<p class="muted">Квитанция ПД-4сб. Введите ФИО плательщика (кто вносит деньги — необязательно
сам работник). В назначение платежа автоматически попадёт ФИО работника, за кого платёж.</p>
<form method="post" action="/employees/{emp.id}/duty_receipt">
<label>Ф.И.О. плательщика (инициалы)</label>
<input type="text" name="payer_name" placeholder="Иванов И. И." value="{_last_payer}">
<button type="submit" name="kind" value="registration" class="btn-full">Квитанция: постановка на учёт (500 ₽) — .docx</button>
<button type="submit" name="kind" value="renewal" class="btn-full">Квитанция: продление пребывания (1000 ₽) — .docx</button>
</form>
</fieldset>
</div>
</div>
{_medzone_section}
{_snils_section}
{_obligations_section}
{_consent_reset_section}
{_admin_delete_section}
{_scans_section}
<a class="btn secondary" href="/employees">← Ко всем сотрудникам</a>
</section>"""
    # Кнопка «вниз к сканам» — только на карточке (там есть секция #scans-section). Скроллит к
    # секции пакета/сканов; видна, пока секция ниже экрана. Стиль как у глобальной «вверх».
    _down_btn = """
<button id="scrollDownBtn" onclick="document.getElementById('scans-section') && document.getElementById('scans-section').scrollIntoView({behavior:'smooth'})"
  style="display:none;position:fixed !important;right:12px;bottom:calc(70px + env(safe-area-inset-bottom,0px));
  z-index:999;width:46px !important;height:46px !important;min-height:0 !important;min-width:0 !important;
  padding:0 !important;margin:0 !important;border:none;border-radius:50% !important;
  background:rgba(74,144,226,.9);color:#fff;font-size:22px;line-height:46px !important;text-align:center;
  box-shadow:0 3px 10px rgba(20,24,30,.28);cursor:pointer" aria-label="К сканам">&#8595;</button>
<script>
(function(){
  var btn = document.getElementById('scrollDownBtn');
  var sec = document.getElementById('scans-section');
  if(!btn || !sec) return;
  function upd(){
    var r = sec.getBoundingClientRect();
    // видна, пока верх секции ниже нижней границы экрана (ещё не доскроллили)
    btn.style.display = (r.top > window.innerHeight - 60) ? 'block' : 'none';
  }
  window.addEventListener('scroll', upd);
  window.addEventListener('resize', upd);
  upd();
})();
</script>"""
    return _render(emp.full_name, body + SAVE_FORM_JS + _down_btn, active="employees", role=request.session.get("role",""))


@app.post("/employees/{employee_id}/entry_date")
def employee_entry_date_submit(
    employee_id: str,
    request: Request,
    entry_date: date = Form(...),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)

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
    registration_valid_until: str = Form(""),
    db: Session = Depends(get_db),
):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    # Дата окончания регистрации — справочная, из уведомления Госуслуг. Пустая строка -> None.
    if registration_valid_until.strip():
        try:
            emp.registration_valid_until = date.fromisoformat(registration_valid_until.strip())
        except ValueError:
            raise HTTPException(400, "Некорректная дата окончания регистрации.")
    else:
        emp.registration_valid_until = None

    previous_address = (emp.address or "").strip()
    new_address = address.strip()
    previous_since = emp.address_since
    first_address_ever = previous_address == ""

    # Дата теперь сохраняется ВСЕГДА (вариант Б), а не только при смене адреса — раньше правка
    # только даты (адрес тот же) игнорировалась, дата "не сохранялась". Исключение: самый первый
    # ввод адреса не создаёт обязательство регистрации (это не переезд), но дату всё равно пишем.
    emp.address = new_address
    emp.address_since = address_since

    # Пересоздать обязательство регистрации нужно, если это НЕ первый ввод адреса И реально
    # изменилось то, от чего считается дедлайн: адрес или дата начала пребывания.
    address_changed = (not first_address_ever) and previous_address != new_address
    since_changed = (not first_address_ever) and previous_since != address_since
    need_reobligate = address_changed or since_changed

    db.commit()
    db.refresh(emp)

    if need_reobligate and emp.consent_status == ConsentStatus.CONFIRMED:
        # Удаляем незакрытое обязательство регистрации-ПЕРЕЕЗДА по СТАРОЙ дате и создаём по новой,
        # чтобы дедлайн соответствовал текущей address_since (вариант Б). Различаем от первичной
        # регистрации (та привязана к entry_date) по trigger_date == старая address_since:
        # сносим только обязательство, чей триггер совпадал с прежним address_since.
        from models import Obligation, ObligationStatus
        if previous_since is not None:
            db.query(Obligation).filter(
                Obligation.employee_id == emp.id,
                Obligation.type == ObligationType.REGISTRATION,
                Obligation.trigger_date == previous_since,
                Obligation.status != ObligationStatus.DONE,
            ).delete(synchronize_session=False)
            db.commit()
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
    _require_role(request, db, UserRole.KADROVIK)

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
    _require_role(request, db, UserRole.KADROVIK)

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
        proof=f"button_click:webforms:{_actor_name(request, db)}:{datetime.now(MSK).isoformat()}",
        consent_text_version=CONSENT_TEXT_VERSION,
    )
    db.add(consent)

    emp.consent_status = ConsentStatus.CONFIRMED
    db.add(emp)
    db.commit()
    db.refresh(emp)

    create_obligations_for_employee(db, emp)

    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/employees/{employee_id}/snils")
def employee_snils_submit(
    employee_id: str,
    request: Request,
    snils: str = Form(""),
    snils_procedure: str = Form(""),
    snils_appointment_date: str = Form(""),
    db: Session = Depends(get_db),
):
    """Сохраняет СНИЛС и/или данные записи в СФР на его получение/объединение. Если введён номер
    СНИЛС — данные записи очищаются (СНИЛС уже получен)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    _snils_clean = snils.strip()
    emp.snils = _snils_clean or None
    if _snils_clean:
        # СНИЛС получен — данные записи больше не нужны
        emp.snils_procedure = None
        emp.snils_appointment_date = None
    else:
        emp.snils_procedure = (snils_procedure.strip() or None)
        emp.snils_appointment_date = _pdate(snils_appointment_date)
    db.commit()
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
    _require_role(request, db, UserRole.KADROVIK)

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
    registration_valid_until: str = Form(""),
    dactyloscopy_date: str = Form(""),
    position: str = Form(""),
    subdivision: str = Form(""),
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
    _require_role(request, db, UserRole.KADROVIK)

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
{_hid("position", position)}
{_hid("subdivision", subdivision)}
<input type="hidden" name="confirmed" value="1">
<button type="submit">Подтвердить и сохранить</button>
</form>
<a class="btn secondary" href="/employees/{emp.id}">Отмена</a>
</section>"""
        return HTMLResponse(_render("Подтверждение", body, role=request.session.get("role","")))

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

    # Должность/подразделение — простые справочные поля, без побочных эффектов на
    # обязательства (в отличие от адреса/дактилоскопии выше), пишем всегда как есть.
    if position.strip():
        emp.position = position.strip()
    if subdivision.strip():
        emp.subdivision = subdivision.strip()

    # Срок регистрации до — справочное поле из уведомления Госуслуг, пишется всегда
    # (пустое -> None). Не привязано к смене адреса, не влияет на обязательства.
    emp.registration_valid_until = _pdate(registration_valid_until)

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
        # Внутренний дедлайн 14 дней от направления: 10 дней на врачей + 4 на справку.
        _days = (date.today() - r.referral_date).days
        _left = 14 - _days
        if _days > 14:
            _status = f'<b style="color:#b00">⚠ Справка просрочена (прошло {_days} дн., внутренний срок 14)</b>'
        elif _days >= 10:
            _status = f'<b style="color:#c47f00">Ждём справку — прошло {_days} дн. из 14 (осталось {_left})</b>'
        else:
            _status = f'<span class="muted">идёт: прошло {_days} дн. из 14 (комиссию пройти к 10-му дню)</span>'
        return (
            f'<div class="card">{name}<br>'
            f'<span class="muted">направлен {r.referral_date.isoformat()}</span><br>'
            f'{_status}<br>'
            f'<span class="muted" style="font-size:13px">Завершение — загрузка скана справки '
            f'(кнопка в карточке работника). Отметка «пройдено» без скана убрана.</span><br>'
            f'<a class="btn secondary" href="/employees/{r.employee_id}">Открыть карточку → загрузить справку</a><br>'
            f'<form class="inline" method="post" action="/medical/{r.employee_id}/result" onsubmit="return confirm(&#39;Удалить направление и вернуть сотрудника в очередь на выписку? Действие необратимо.&#39;)">'
            f'<input type="hidden" name="result" value="failed">'
            f'<button type="submit" class="secondary">❌ Не пройдено (отменить направление)</button></form>'
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
    return _render("Медкомиссия", body, active="medical", role=request.session.get("role",""))


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


@app.post("/employees/{employee_id}/labor_contract")
def employee_labor_contract(
    employee_id: str,
    request: Request,
    position: str = Form("Монтажник"),
    salary: str = Form("30000"),
    contract_date: date = Form(...),
    db: Session = Depends(get_db),
):
    """Заключение договора: пишет contract_date (создаёт обязательства МВД/ЕФС-1 при согласии),
    возвращает в карточку. Docx скачивается ОТДЕЛЬНО после заключения (роут /download).
    Блокируется при пустом статусе учёта и пустом табельном. Должность/оклад из формы как есть."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    if emp.registration_status is None:
        raise HTTPException(400, "Не задан статус миграционного учёта — сначала выберите статус.")
    if not (emp.tab_number or "").strip():
        raise HTTPException(400, "У сотрудника нет табельного номера — номер договора собрать нельзя.")

    # Заключение: фиксируем дату договора (создаёт обязательства МВД/ЕФС-1 при согласии) и
    # возвращаем в карточку. Скачивание .docx — отдельной кнопкой ПОСЛЕ заключения (роут
    # /download). Так «Заключить» — это акт фиксации, а не выдача файла: кадровик сперва
    # смотрит предпросмотр, заключает, потом скачивает.
    emp.contract_date = contract_date
    db.commit()
    db.refresh(emp)
    if emp.consent_status == ConsentStatus.CONFIRMED:
        create_obligations_for_employee(db, emp)

    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


def _render_labor_contract_preview(emp, position, salary, contract_date_str, tab):
    """HTML-предпросмотр трудового договора. Зеркалит текст generate_labor_contract_docx
    в document_templates.py. Меняешь текст там — поменяй и здесь, иначе предпросмотр
    разойдётся с .docx. Вид «похожий», не пиксель-в-пиксель (решение пользователя)."""
    import html as _h
    contract_no = f"{CONTRACT_NUMBER_PREFIX}{tab}" if tab else "—"
    s = (salary or "").strip().replace(" ", "")
    salary_fmt = f"{int(s):,}".replace(",", " ") if s.isdigit() else (salary or "—")
    pos = _h.escape((position or "").strip() or "—")
    body = f"""
<div style="max-width:760px">
<h1 style="text-align:center;font-size:18px">ТРУДОВОЙ ДОГОВОР № {_h.escape(contract_no)}</h1>
<p style="display:flex;justify-content:space-between"><span>г. Москва</span><span>{_h.escape(contract_date_str)}</span></p>
<p>{_h.escape(EMPLOYER_NAME_FULL)}, именуемое «Работодатель», в лице Генерального директора
{_h.escape(EMPLOYER_DIRECTOR_FULL)}, действующего на основании Устава, с одной стороны, и
{_h.escape(emp.full_name)}, именуемый «Работник», с другой стороны, заключили настоящий договор:</p>
<h3>1. Предмет и срок</h3>
<p>1.1. Работник принимается в {_h.escape(EMPLOYER_SUBDIVISION)} {_h.escape(EMPLOYER_NAME_SHORT)}
на должность {pos} с {_h.escape(contract_date_str)}.</p>
<p>1.3. Место работы: {_h.escape(WORKPLACE_ADDRESS)}</p>
<p>1.5. Срок — неопределённый.</p>
<h3>3. Оплата труда</h3>
<p>3.1. Оклад: {_h.escape(salary_fmt)} руб.; Районный коэффициент: {_h.escape(DISTRICT_COEFFICIENT)}.</p>
<h3>8. Адреса и подписи</h3>
<table style="width:100%;border-collapse:collapse"><tr>
<td style="width:50%;vertical-align:top;padding-right:12px">
<b>Работник:</b><br>
{_h.escape(emp.full_name)}, {emp.birth_date.strftime("%d.%m.%Y") if emp.birth_date else "—"} г.р.<br>
Паспорт: {_h.escape((emp.passport_series or "") + " " + (emp.passport_number or ""))}, выдан: —<br>
Адрес: {_h.escape(CONTRACT_SITE_ADDRESS)}<br><br>
Подпись: _______________
</td>
<td style="width:50%;vertical-align:top;padding-left:12px">
<b>Работодатель:</b><br>
{_h.escape(EMPLOYER_NAME_FULL)}<br>
ИНН {_h.escape(str(EMPLOYER_INN))} КПП {_h.escape(str(EMPLOYER_KPP))}<br>
{_h.escape(EMPLOYER_LEGAL_ADDRESS)}<br>
Телефон: {_h.escape(EMPLOYER_PHONE)}<br><br>
Ген. директор _______________<br>{_h.escape(EMPLOYER_DIRECTOR_SHORT)}<br>м.п.
</td>
</tr></table>
</div>
<p class="muted">Это предпросмотр (упрощённый вид). Полный документ — в скачанном .docx после заключения.</p>
<div style="display:flex;gap:8px;flex-wrap:wrap">
<a class="btn secondary" href="/employees/{emp.id}">← Назад к карточке</a>
<button class="btn" onclick="window.print()">Печать</button>
</div>"""
    return _render("Предпросмотр договора", body, active="employees", role=request.session.get("role",""))


@app.post("/employees/{employee_id}/labor_contract/preview")
def employee_labor_contract_preview(
    employee_id: str,
    request: Request,
    position: str = Form("Монтажник"),
    salary: str = Form("30000"),
    contract_date: date = Form(...),
    db: Session = Depends(get_db),
):
    """Предпросмотр договора. НИЧЕГО не пишет в БД (не заключает). Показывает HTML по данным
    формы. Доступен до заключения — иначе кадровик не смог бы проверить документ глазами."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    tab = (emp.tab_number or "").strip()
    return HTMLResponse(_render_labor_contract_preview(
        emp, position, salary, contract_date.strftime("%d.%m.%Y"), tab))


def _find_soffice() -> str | None:
    """Ищет бинарь soffice. На Railway/Nixpacks LibreOffice ставится в Nix store и НЕ попадает
    в PATH — поэтому недостаточно вызвать 'soffice'. Проверяем PATH, затем типовые пути и Nix
    store. Возвращает путь к бинарю или None."""
    import shutil, glob
    # 1) в PATH
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    # 2) типовые пути
    for p in ("/usr/bin/soffice", "/usr/local/bin/soffice",
              "/usr/lib/libreoffice/program/soffice", "/opt/libreoffice/program/soffice"):
        if os.path.exists(p):
            return p
    # 3) Nix store (Railway) — ищем soffice в /nix/store/*/bin и program-каталогах
    for pattern in ("/nix/store/*/bin/soffice",
                    "/nix/store/*libreoffice*/lib/libreoffice/program/soffice",
                    "/nix/store/*libreoffice*/program/soffice"):
        hits = glob.glob(pattern)
        if hits:
            return hits[0]
    return None


def _docx_to_pdf(docx_path: str) -> str:
    """Конвертирует docx в pdf через LibreOffice (soffice --headless). Возвращает путь к pdf.
    Требует libreoffice в контейнере (nixpacks.toml). Ищет soffice не только в PATH, но и в
    Nix store — на Railway бинарь туда и попадает, минуя PATH (частая причина 'soffice not found').
    Для пакета Госуслуг: договор и квитанция генерируются из одного docx — расхождения нет."""
    import subprocess
    soffice = _find_soffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice (soffice) не найден в контейнере. Проверьте, что nixpacks.toml с "
            "'libreoffice' в nixPkgs задеплоен и в build-логе есть его установка."
        )
    out_dir = os.path.dirname(docx_path)
    # HOME нужен soffice для профиля; в некоторых контейнерах он не задан -> падение.
    env = dict(os.environ)
    env.setdefault("HOME", out_dir)
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
            check=True, capture_output=True, timeout=90, env=env,
        )
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Конвертация docx->pdf не удалась: {e.stderr.decode(errors='ignore')[:200]}")
    except subprocess.TimeoutExpired:
        raise RuntimeError("Конвертация docx->pdf превысила таймаут")
    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    if not os.path.exists(pdf_path):
        raise RuntimeError("PDF не создан после конвертации")
    return pdf_path


@app.post("/employees/{employee_id}/labor_contract/download")
def employee_labor_contract_download(
    employee_id: str,
    request: Request,
    position: str = Form("Монтажник"),
    salary: str = Form("30000"),
    db: Session = Depends(get_db),
):
    """Скачивание .docx. Доступно ТОЛЬКО после заключения (contract_date стоит) — иначе 400.
    Не пишет в БД, только генерирует файл по уже зафиксированной дате договора."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    if emp.contract_date is None:
        raise HTTPException(400, "Договор не заключён — сначала нажмите «Заключить».")
    try:
        path = generate_labor_contract_docx(emp, position=position, salary=salary,
                                            contract_date=emp.contract_date)
    except Exception:
        raise HTTPException(500, "Не удалось сгенерировать договор. Проверьте логи сервиса.")
    filename = f"Трудовой_договор_{emp.full_name.replace(' ', '_')}.docx"
    return FileResponse(path, filename=filename)


@app.post("/employees/{employee_id}/labor_contract/download_pdf")
def employee_labor_contract_download_pdf(
    employee_id: str,
    request: Request,
    position: str = Form("Монтажник"),
    salary: str = Form("30000"),
    db: Session = Depends(get_db),
):
    """PDF-версия договора для пакета Госуслуг. Генерирует docx и конвертирует в PDF через
    LibreOffice — один источник (docx), PDF всегда совпадает. Требует libreoffice в контейнере."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    if emp.contract_date is None:
        raise HTTPException(400, "Договор не заключён — сначала нажмите «Заключить».")
    try:
        docx_path = generate_labor_contract_docx(emp, position=position, salary=salary,
                                                 contract_date=emp.contract_date)
        pdf_path = _docx_to_pdf(docx_path)
    except RuntimeError as e:
        raise HTTPException(500, f"PDF недоступен: {e}")
    except Exception:
        raise HTTPException(500, "Не удалось сгенерировать PDF договора. Проверьте логи.")
    filename = f"Трудовой_договор_{emp.full_name.replace(' ', '_')}.pdf"
    return FileResponse(pdf_path, filename=filename, media_type="application/pdf")


@app.post("/employees/{employee_id}/duty_receipt")
def employee_duty_receipt(
    employee_id: str,
    request: Request,
    kind: str = Form(...),
    payer_name: str = Form(""),
    db: Session = Depends(get_db),
):
    """Генерация квитанции госпошлины (ПД-4сб) и отдача docx. kind — из нажатой кнопки
    (registration/renewal). payer_name — ФИО плательщика из формы (плательщик не обязательно
    работник). ФИО работника уходит в назначение платежа автоматически внутри генератора."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    if kind not in ("registration", "renewal"):
        raise HTTPException(400, "Неизвестный тип пошлины")
    # Запомнить плательщика на сессию (prefill в следующих карточках). Пустой ввод не затирает.
    if (payer_name or "").strip():
        request.session["last_payer"] = payer_name.strip()
    try:
        path = generate_duty_receipt_docx(kind, employee=emp, payer_name=payer_name)
    except Exception:
        raise HTTPException(500, "Не удалось сгенерировать квитанцию. Проверьте логи сервиса.")
    kind_ru = "постановка" if kind == "registration" else "продление"
    filename = f"Квитанция_{kind_ru}_{emp.full_name.replace(' ', '_')}.docx"
    return FileResponse(path, filename=filename)


@app.post("/employees/{employee_id}/registration_status")
def employee_registration_status(
    employee_id: str,
    request: Request,
    registration_status: str = Form(""),
    db: Session = Depends(get_db),
):
    """Смена статуса учёта. Пишет статус и ПЕРЕСОЗДАЁТ обязательства под него:
    - create_obligations_for_employee создаёт недостающие по новому статусу (если согласие есть);
    - лишние НЕ выполненные (PENDING) обязательства, которые новый статус делает ненужными,
      удаляются. Выполненные (DONE) НЕ трогаются — след исполнения.
    PRIMARY->PRIOR: сносятся PENDING медосмотр, дактилоскопия, регистрация-от-въезда.
    PRIOR->PRIMARY: create_obligations досоздаёт медосмотр/дактилоскопию (ничего не сносим).
    Пустой статус: обязательства не создаются (гейт в obligations), существующие PENDING
    не трогаем — просто перестаёт быть валидным для новых расчётов."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    raw = (registration_status or "").strip()
    if raw == "":
        emp.registration_status = None
    elif raw == RegistrationStatus.PRIMARY.value:
        emp.registration_status = RegistrationStatus.PRIMARY
    elif raw == RegistrationStatus.PRIOR.value:
        emp.registration_status = RegistrationStatus.PRIOR
    else:
        raise HTTPException(400, "Недопустимый статус учёта")
    db.commit()
    db.refresh(emp)

    # PRIOR делает лишними обязательства, привязанные к факту въезда — снять их PENDING.
    if emp.registration_status == RegistrationStatus.PRIOR:
        entry_bound = (
            ObligationType.MEDICAL_EXAM,
            ObligationType.DACTYLOSCOPY,
        )
        obs = db.scalars(
            select(Obligation)
            .where(Obligation.employee_id == emp.id)
            .where(Obligation.type.in_(entry_bound))
            .where(Obligation.is_current == True)  # noqa: E712
            .where(Obligation.status == ObligationStatus.PENDING)
        ).all()
        for o in obs:
            db.delete(o)
        # регистрация-от-въезда (trigger_date == entry_date) тоже лишняя при PRIOR
        if emp.entry_date is not None:
            reg_entry = db.scalars(
                select(Obligation)
                .where(Obligation.employee_id == emp.id)
                .where(Obligation.type == ObligationType.REGISTRATION)
                .where(Obligation.trigger_date == emp.entry_date)
                .where(Obligation.is_current == True)  # noqa: E712
                .where(Obligation.status == ObligationStatus.PENDING)
            ).all()
            for o in reg_entry:
                db.delete(o)
        db.commit()

    # создать недостающие по новому статусу (сама функция гейтит по статусу и согласию)
    if emp.registration_status is not None and emp.consent_status == ConsentStatus.CONFIRMED:
        create_obligations_for_employee(db, emp)

    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/employees/{employee_id}/termination")
def employee_termination(
    employee_id: str,
    request: Request,
    termination_date: str = Form(...),
    basis: str = Form(""),
    basis_note: str = Form(""),
    db: Session = Depends(get_db),
):
    """Оформление увольнения: пишет contract_end_date, создаёт обязательства (расторжение +
    убытие) через create_obligations_for_employee. Будущая дата разрешена — обязательства
    отложатся до наступления (логика в obligations.py). Валидация: дата не раньше договора."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    if emp.contract_date is None:
        raise HTTPException(400, "Нельзя оформить увольнение: договор не заключён.")
    try:
        term_date = date.fromisoformat(termination_date)
    except ValueError:
        raise HTTPException(400, "Некорректная дата увольнения.")
    # Валидация: дата увольнения не раньше даты договора.
    if term_date < emp.contract_date:
        raise HTTPException(400, "Дата увольнения не может быть раньше даты договора.")
    # основание: если "иное" — берём примечание
    final_basis = basis_note.strip() if basis == "иное" and basis_note.strip() else basis
    emp.contract_end_date = term_date
    db.commit()
    # создаём обязательства увольнения (расторжение + убытие; будущая дата -> отложатся внутри)
    create_obligations_for_employee(db, emp)
    db.commit()

    # Гасим СТАРЫЕ обязательства (от приёма): работник уволен и убывает, они неактуальны и не
    # должны висеть в просроченных. НЕ трогаем обязательства САМОГО увольнения (расторжение,
    # снятие с учёта) — их надо подать. Не трогаем уже DONE. Помечаем CANCELLED (след остаётся).
    _stale_types = [
        ObligationType.REGISTRATION,
        ObligationType.MEDICAL_EXAM,
        ObligationType.DACTYLOSCOPY,
        ObligationType.EFS1_REPORT,
        ObligationType.CONTRACT_NOTICE,
        ObligationType.REGISTRATION_RENEWAL,
    ]
    db.query(Obligation).filter(
        Obligation.employee_id == emp.id,
        Obligation.type.in_(_stale_types),
        Obligation.status.in_([ObligationStatus.PENDING, ObligationStatus.OVERDUE]),
    ).update({Obligation.status: ObligationStatus.CANCELLED}, synchronize_session=False)
    db.commit()

    # сохраним основание в сессию для генерации уведомления (в модели поля нет — не плодим ALTER)
    request.session[f"term_basis_{emp.id}"] = final_basis
    return RedirectResponse(f"/employees/{emp.id}", status_code=303)


@app.post("/employees/{employee_id}/termination_notice")
def employee_termination_notice(employee_id: str, request: Request, db: Session = Depends(get_db)):
    """Скачать уведомление о расторжении договора (форма №8, приказ №536). Доступно всем
    вошедшим (прораб тоже может скачать документ)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    if emp.contract_end_date is None:
        raise HTTPException(400, "Увольнение не оформлено — нет даты расторжения.")
    basis = request.session.get(f"term_basis_{emp.id}", "")
    try:
        path = generate_termination_notice_docx(emp, basis=basis)
    except Exception:
        raise HTTPException(500, "Не удалось сгенерировать уведомление о расторжении.")
    fn = f"Уведомление_расторжение_{emp.full_name.replace(' ', '_')}.docx"
    return FileResponse(path, filename=fn)


@app.post("/employees/{employee_id}/departure_notice")
def employee_departure_notice(employee_id: str, request: Request, db: Session = Depends(get_db)):
    """Скачать уведомление об убытии (снятие с миграционного учёта). Доступно всем вошедшим."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    if emp.contract_end_date is None:
        raise HTTPException(400, "Увольнение не оформлено — нет даты убытия.")
    try:
        path = generate_departure_notice_docx(emp)
    except Exception:
        raise HTTPException(500, "Не удалось сгенерировать уведомление об убытии.")
    fn = f"Уведомление_убытие_{emp.full_name.replace(' ', '_')}.docx"
    return FileResponse(path, filename=fn)


@app.post("/employees/{employee_id}/obligation/mark_done")
def employee_obligation_mark_done(
    employee_id: str,
    request: Request,
    obligation_id: str = Form(...),
    db: Session = Depends(get_db),
):
    """Ручная отметка обязательства как поданного (ЕФС-1, уведомление МВД, регистрация и др.,
    что подаётся вовне и не имеет своего закрывателя). Пишет DONE + дату + автора. Кадровик/админ."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    ob = db.get(Obligation, obligation_id)
    if ob is None or ob.employee_id != employee_id:
        raise HTTPException(404, "Обязательство не найдено")
    if ob.status == ObligationStatus.DONE:
        return RedirectResponse(f"/employees/{employee_id}", status_code=303)
    ob.status = ObligationStatus.DONE
    ob.done_date = date.today()
    ob.done_by = _actor_name(request, db)
    db.add(ob)
    db.commit()
    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/employees/{employee_id}/obligation/reopen")
def employee_obligation_reopen(
    employee_id: str,
    request: Request,
    obligation_id: str = Form(...),
    db: Session = Depends(get_db),
):
    """Отмена ошибочной отметки: возвращает обязательство в работу. Пересчёт статуса
    (PENDING/OVERDUE) сделает cron при следующем прогоне; ставим PENDING, снимаем дату/автора."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    ob = db.get(Obligation, obligation_id)
    if ob is None or ob.employee_id != employee_id:
        raise HTTPException(404, "Обязательство не найдено")
    ob.status = ObligationStatus.PENDING
    ob.done_date = None
    ob.done_by = None
    db.add(ob)
    db.commit()
    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


# --- Перенос сканов в корзину при удалении события (2026-07) -----------------
# Сканы привязаны к типу события. Удаление НЕ физическое: объект переносится в
# префикс deleted/<исходный ключ>/<timestamp> того же бакета (copy+delete), откат
# возможен. Timestamp в ключе — чтобы повторное удаление скана того же типа не
# затёрло предыдущий в корзине. Паспорт (passport) не трогаем нигде — базовый
# документ, к миграционным событиям не привязан.
_SCANS_BY_OBLIGATION = {
    ObligationType.MEDICAL_EXAM: ["medical_certificate"],
    ObligationType.REGISTRATION: ["arrival_notice", "payment_registration"],
    ObligationType.REGISTRATION_RENEWAL: ["payment_renewal"],
}
_SCANS_BY_DATE = {
    # Дата въезда — полный сброс: всё миграционное (кроме паспорта).
    "entry_date": ["migration_card", "medical_certificate", "arrival_notice",
                   "payment_registration", "payment_renewal"],
    "address_since": ["arrival_notice"],
    "contract_date": [],       # договор — генерируемый docx, не скан
    "dactyloscopy_date": [],   # возврат в работу; грин-карты в сканах нет
}


def _soft_delete_scan(employee_id: str, scan_type: str) -> str | None:
    """Переносит скан в корзину deleted/ (copy+delete) вместо физического удаления.
    Возвращает новый ключ в корзине или None, если скана нет / S3 недоступен. НЕ бросает
    исключение наружу: сбой S3 не должен откатывать уже сделанное снятие обязательств в БД
    (скан в худшем случае останется на месте, но не потеряется)."""
    try:
        client = _s3_client()
    except Exception:
        return None
    src_key = _scan_key(scan_type, employee_id)
    try:
        client.head_object(Bucket=S3_BUCKET, Key=src_key)  # есть ли что переносить
    except Exception:
        return None  # скана нет — нечего переносить
    ts = datetime.now(MSK).strftime("%Y%m%dT%H%M%S")
    dst_key = f"deleted/{src_key}/{ts}"
    try:
        client.copy_object(
            Bucket=S3_BUCKET, Key=dst_key,
            CopySource={"Bucket": S3_BUCKET, "Key": src_key},
        )
        client.delete_object(Bucket=S3_BUCKET, Key=src_key)
        return dst_key
    except Exception as e:
        log.warning("Не удалось перенести скан %s employee=%s в корзину: %s",
                    scan_type, employee_id, str(e)[:200])
        return None


@app.post("/employees/{employee_id}/event/delete")
def employee_event_delete(
    employee_id: str,
    request: Request,
    event_kind: str = Form(...),
    target: str = Form(...),
    db: Session = Depends(get_db),
):
    """Удаление/сброс события из карточки. Мягкое: obligations помечаются
    is_current=False/CANCELLED, даты-триггеры обнуляются, согласие → DRAFT; из базы
    ничего не стирается, всё для отката пишется в EventDeletionLog (snapshot).

    Права по типу события:
      event_kind='consent'    → кадровик + админ (рабочий инструмент);
      всё остальное           → только админ.

    event_kind='obligation'   → снять одно обязательство (target = Obligation.id).
    event_kind='date'         → обнулить дату-триггер (target = имя поля) + каскад:
      entry_date        → ВСЕ активные obligations сотрудника → CANCELLED (вариант A);
      contract_date     → efs1_report + contract_notice (по этой дате) → CANCELLED;
      address_since     → registration (по этой дате) → CANCELLED;
      dactyloscopy_date → dactyloscopy ВОЗВРАЩАЕТСЯ в PENDING (снятие закрытия, не отмена).
    event_kind='consent'      → сброс согласия ПД: consent_status → DRAFT (запись Consent
                                НЕ удаляется, остаётся следом; при повторном подтверждении
                                создастся новая). target игнорируется.
    event_kind='full_reset'   → «чистый лист» (только админ): согласие → DRAFT + дата въезда
                                → NULL + ВСЕ активные obligations → CANCELLED + миграционные
                                сканы в корзину. Один снимок в аудит.
    """
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    # Гейт по типу события: сброс согласия — рабочий инструмент кадровика; удаление
    # обязательств/дат и полный сброс — только админ.
    if event_kind == "consent":
        _require_role(request, db, UserRole.KADROVIK)   # ADMIN проходит внутри _require_role
    else:
        _require_role(request, db, UserRole.ADMIN)

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    actor = _actor_name(request, db)

    if event_kind == "obligation":
        ob = db.get(Obligation, target)
        if ob is None or ob.employee_id != employee_id:
            raise HTTPException(404, "Обязательство не найдено")
        snapshot = {
            "obligation_id": ob.id,
            "type": ob.type.value,
            "prev_status": ob.status.value,
            "prev_is_current": ob.is_current,
            "trigger_date": ob.trigger_date.isoformat() if ob.trigger_date else None,
            "deadline_date": ob.deadline_date.isoformat() if ob.deadline_date else None,
        }
        ob.is_current = False
        ob.status = ObligationStatus.CANCELLED
        db.add(ob)
        # Сканы, привязанные к типу этого обязательства → в корзину deleted/.
        _moved = []
        for _st in _SCANS_BY_OBLIGATION.get(ob.type, []):
            _dk = _soft_delete_scan(employee_id, _st)
            if _dk:
                _moved.append({"scan_type": _st, "deleted_key": _dk})
        snapshot["moved_scans"] = _moved
        db.add(EventDeletionLog(
            employee_id=employee_id, actor=actor, event_kind="obligation",
            target=ob.type.value, snapshot=json.dumps(snapshot, ensure_ascii=False),
        ))
        db.commit()
        log.info("ADMIN %s удалил obligation %s (%s) у employee_id=%s, сканов в корзину: %d",
                 actor, ob.id, ob.type.value, employee_id, len(_moved))
        return RedirectResponse(f"/employees/{employee_id}", status_code=303)

    if event_kind == "date":
        DATE_FIELDS = {"entry_date", "contract_date", "address_since", "dactyloscopy_date"}
        if target not in DATE_FIELDS:
            raise HTTPException(400, "Неизвестное поле даты")
        old_value = getattr(emp, target, None)
        affected = []

        if target == "dactyloscopy_date":
            # Возврат обязанности в работу, НЕ отмена: дата закрывала DACTYLOSCOPY.
            for o in emp.obligations:
                if (o.is_current and o.type == ObligationType.DACTYLOSCOPY
                        and o.status != ObligationStatus.CANCELLED):
                    affected.append({
                        "obligation_id": o.id, "type": o.type.value,
                        "prev_status": o.status.value, "prev_is_current": o.is_current,
                    })
                    o.status = ObligationStatus.PENDING
                    o.done_date = None
                    o.done_by = None
                    db.add(o)
            emp.dactyloscopy_date = None
        else:
            if target == "entry_date":
                # Вариант A (согласовано): полный сброс — все активные обязательства.
                to_cancel = [o for o in emp.obligations
                             if o.is_current and o.status != ObligationStatus.CANCELLED]
            elif target == "contract_date":
                _types = {ObligationType.EFS1_REPORT, ObligationType.CONTRACT_NOTICE}
                to_cancel = [o for o in emp.obligations
                             if o.is_current and o.type in _types
                             and o.trigger_date == old_value
                             and o.status != ObligationStatus.CANCELLED]
            else:  # address_since
                to_cancel = [o for o in emp.obligations
                             if o.is_current and o.type == ObligationType.REGISTRATION
                             and o.trigger_date == old_value
                             and o.status != ObligationStatus.CANCELLED]
            for o in to_cancel:
                affected.append({
                    "obligation_id": o.id, "type": o.type.value,
                    "prev_status": o.status.value, "prev_is_current": o.is_current,
                })
                o.is_current = False
                o.status = ObligationStatus.CANCELLED
                db.add(o)
            setattr(emp, target, None)

        db.add(emp)
        # Сканы, привязанные к этой дате-триггеру → в корзину deleted/.
        _moved = []
        for _st in _SCANS_BY_DATE.get(target, []):
            _dk = _soft_delete_scan(employee_id, _st)
            if _dk:
                _moved.append({"scan_type": _st, "deleted_key": _dk})
        snapshot = {
            "field": target,
            "old_value": old_value.isoformat() if old_value else None,
            "affected": affected,
            "moved_scans": _moved,
        }
        db.add(EventDeletionLog(
            employee_id=employee_id, actor=actor, event_kind="date",
            target=target, snapshot=json.dumps(snapshot, ensure_ascii=False),
        ))
        db.commit()
        log.info("ADMIN %s обнулил дату %s (было %s), обязательств: %d, сканов в корзину: %d, employee_id=%s",
                 actor, target, old_value, len(affected), len(_moved), employee_id)
        return RedirectResponse(f"/employees/{employee_id}", status_code=303)

    if event_kind == "consent":
        # Сброс согласия ПД (кадровик/админ). CONFIRMED → DRAFT. Запись Consent НЕ удаляем —
        # остаётся следом; при повторном подтверждении создастся новая. Обязательства при этом
        # НЕ снимаются (для этого есть отдельные кнопки/полный сброс) — сбрасывается только гейт.
        prev = emp.consent_status.value if emp.consent_status else None
        if emp.consent_status != ConsentStatus.CONFIRMED:
            return RedirectResponse(f"/employees/{employee_id}", status_code=303)
        emp.consent_status = ConsentStatus.DRAFT
        db.add(emp)
        db.add(EventDeletionLog(
            employee_id=employee_id, actor=actor, event_kind="consent",
            target="consent",
            snapshot=json.dumps({"prev_consent_status": prev}, ensure_ascii=False),
        ))
        db.commit()
        log.info("%s сбросил согласие ПД (было %s) у employee_id=%s", actor, prev, employee_id)
        return RedirectResponse(f"/employees/{employee_id}", status_code=303)

    if event_kind == "full_reset":
        # «Чистый лист» (только админ): согласие → DRAFT + дата въезда → NULL + ВСЕ активные
        # obligations → CANCELLED + миграционные сканы (маппинг entry_date) в корзину. Один
        # снимок в аудит. Возвращает сотрудника в исходное состояние до заведения.
        prev_consent = emp.consent_status.value if emp.consent_status else None
        prev_entry = emp.entry_date.isoformat() if emp.entry_date else None
        affected = []
        for o in emp.obligations:
            if o.is_current and o.status != ObligationStatus.CANCELLED:
                affected.append({
                    "obligation_id": o.id, "type": o.type.value,
                    "prev_status": o.status.value, "prev_is_current": o.is_current,
                })
                o.is_current = False
                o.status = ObligationStatus.CANCELLED
                db.add(o)
        emp.entry_date = None
        emp.consent_status = ConsentStatus.DRAFT
        db.add(emp)
        _moved = []
        for _st in _SCANS_BY_DATE.get("entry_date", []):
            _dk = _soft_delete_scan(employee_id, _st)
            if _dk:
                _moved.append({"scan_type": _st, "deleted_key": _dk})
        db.add(EventDeletionLog(
            employee_id=employee_id, actor=actor, event_kind="full_reset",
            target="full_reset",
            snapshot=json.dumps({
                "prev_consent_status": prev_consent,
                "prev_entry_date": prev_entry,
                "affected": affected,
                "moved_scans": _moved,
            }, ensure_ascii=False),
        ))
        db.commit()
        log.info("ADMIN %s ПОЛНЫЙ СБРОС employee_id=%s: согласие %s→draft, въезд %s→null, "
                 "снято обязательств: %d, сканов в корзину: %d",
                 actor, employee_id, prev_consent, prev_entry, len(affected), len(_moved))
        return RedirectResponse(f"/employees/{employee_id}", status_code=303)

    raise HTTPException(400, "Неизвестный тип события")


# Ожидаемая сумма (руб.) по типу слота платёжки. Постановка 500, продление 1000.
PAYMENT_EXPECTED_AMOUNT = {"payment_registration": 500, "payment_renewal": 1000}


def _extract_payment_amounts(text: str) -> set:
    """Извлекает суммы платежа по МАРКЕРАМ «Сумма платежа» / «Итого» / «Сумма», а не по всему
    тексту — иначе числа в реквизитах (ОКТМО, КБК, счета, УИН) дают ложные совпадения.
    Возвращает набор целых рублей, найденных рядом с маркерами."""
    import re
    t = text.replace("\xa0", " ")
    amounts = set()
    for marker in ["Сумма платежа", "Итого", "Сумма"]:
        for m in re.finditer(marker, t, re.IGNORECASE):
            tail = t[m.end():m.end() + 40]
            num = re.search(r"([\d\s]{1,12}),?\d{0,2}\s*(?:руб|₽|р\.)", tail)
            if num:
                val = num.group(1).replace(" ", "").strip()
                if val.isdigit():
                    amounts.add(int(val))
    return amounts


def _payment_amount_check(pdf_bytes: bytes, scan_type: str) -> bool | None:
    """Проверяет сумму платёжки по маркеру «Сумма платежа»/«Итого» (надёжнее, чем поиск числа
    по всему тексту — реквизиты не мешают). Сравнивает с ожидаемой суммой слота (500/1000).
    True — ожидаемая сумма найдена и чужой нет; False — найдена чужая сумма (метка);
    None — текст не прочитан (скан) или маркер суммы не найден (другой формат чека)."""
    expected = PAYMENT_EXPECTED_AMOUNT.get(scan_type)
    if expected is None:
        return None
    other = 1000 if expected == 500 else 500
    try:
        import io
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(pdf_bytes))
        text = "".join((p.extract_text() or "") for p in reader.pages)
        if not text.strip():
            return None
        amounts = _extract_payment_amounts(text)
        if not amounts:
            return None  # маркер суммы не найден — не проверяем (не блокируем зря)
        if other in amounts and expected not in amounts:
            return False  # найдена ТОЛЬКО чужая сумма -> точно не тот слот
        if expected in amounts:
            return True   # ожидаемая сумма есть
        return False      # ожидаемой нет, но какая-то сумма есть -> подозрительно
    except Exception:
        return None


def _payment_surname_check(pdf_bytes: bytes, full_name: str) -> bool | None:
    """Проверяет, встречается ли фамилия работника в тексте PDF-платёжки (назначение платежа).
    True — нашли; False — не нашли (повод предупредить); None — не смогли прочитать (скан без
    текста/не PDF). Ищем по ФАМИЛИИ как подстроке — устойчиво к падежам/инициалам.
    ЭТО ХЕЛПЕР, НЕ РОУТ — декоратор @app.post должен стоять над employee_scan_upload ниже."""
    if not full_name or not full_name.strip():
        return None
    surname = full_name.strip().split()[0].lower()
    if len(surname) < 3:
        return None
    try:
        import io
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(pdf_bytes))
        text = "".join((p.extract_text() or "") for p in reader.pages).lower()
        if not text.strip():
            return None
        return surname in text
    except Exception:
        return None


@app.post("/employees/{employee_id}/passport_pages")
def employee_passport_pages(
    employee_id: str,
    request: Request,
    all_pages: str = Form(""),
    db: Session = Depends(get_db),
):
    """Сохраняет чекбокс «все страницы паспорта загружены» (подтверждение кадровика)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    emp.passport_all_pages = bool(all_pages)  # чекбокс отмечен -> "on", иначе пусто
    db.commit()
    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


def _translit_iso(s: str) -> str:
    """Обратный транслит латиница->кириллица для ЧЕРНОВИКА ФИО. Неоднозначен — только подсказка
    под ручную сверку с кириллицей на карте. Эвристика: казахское -AYEV/-AYEVA -> -аев/-аева."""
    if not s:
        return ""
    s = s.upper().strip()
    suffix = ""
    for lat, cyr in [("AYEVA", "аева"), ("AYEV", "аев"), ("EEVA", "еева"), ("EEV", "еев")]:
        if s.endswith(lat):
            suffix = cyr
            s = s[:-len(lat)]
            break
    for lat, cyr in [("SHCH","Щ"),("KH","Х"),("ZH","Ж"),("CH","Ч"),("SH","Ш"),
                     ("YU","Ю"),("YA","Я"),("YO","Ё"),("TS","Ц")]:
        s = s.replace(lat, cyr)
    single = {"A":"А","B":"Б","V":"В","G":"Г","D":"Д","E":"Е","Z":"З","I":"И","Y":"Й",
              "K":"К","L":"Л","M":"М","N":"Н","O":"О","P":"П","R":"Р","S":"С","T":"Т",
              "U":"У","F":"Ф","H":"Х","C":"К","J":"Ж","Q":"К","W":"В","X":"КС"}
    out = "".join(single.get(ch, ch) for ch in s) + suffix
    return out.capitalize()


def _ocr_id_card(image_bytes: bytes) -> dict:
    """Распознаёт MRZ удостоверения (passporteye) с перебором поворотов. Возвращает dict полей
    для формы: full_name_translit, birth_date(ISO), passport_number, iin, citizenship. Пустой
    dict, если MRZ не распознана или passporteye не установлен. ФИО — черновик под сверку."""
    try:
        import io
        from passporteye import read_mrz
        from PIL import Image, ImageOps
    except ImportError:
        return {}
    try:
        base = Image.open(io.BytesIO(image_bytes))
        base = ImageOps.exif_transpose(base)
        if base.mode != "RGB":
            base = base.convert("RGB")
    except Exception:
        return {}
    best, best_score = None, -1
    for angle in (0, 90, 180, 270):
        try:
            rot = base.rotate(angle, expand=True)
            buf = io.BytesIO(); rot.save(buf, format="PNG"); buf.seek(0)
            m = read_mrz(buf)
            if m is not None:
                sc = m.to_dict().get("valid_score", 0)
                if sc > best_score:
                    best, best_score = m, sc
        except Exception:
            continue
    if best is None:
        return {}
    d = best.to_dict()
    res = {}

    def _clean_mrz_name(raw):
        # MRZ-имя: разделители '<' -> пробел, оставляем только части длиной >1 (одиночные буквы —
        # артефакты распознавания, как хвост 'K' от <<K<<). Возвращаем очищенную строку.
        parts = [p for p in (raw or "").replace("<", " ").split() if len(p) > 1]
        return " ".join(parts)

    surname = _clean_mrz_name(d.get("surname"))
    names = _clean_mrz_name(d.get("names"))
    res["full_name_translit"] = " ".join(p for p in [_translit_iso(surname), _translit_iso(names)] if p)
    res["passport_number"] = (d.get("number") or "").replace("<", "").strip()
    dob = (d.get("date_of_birth") or "").strip()
    res["birth_date"] = ""
    if len(dob) == 6 and dob.isdigit():
        import datetime as _dt
        yy, mm, dd = int(dob[:2]), dob[2:4], dob[4:6]
        cur_yy = _dt.date.today().year % 100
        year = 1900 + yy if yy > cur_yy else 2000 + yy
        try:
            _dt.date(year, int(mm), int(dd))
            res["birth_date"] = f"{year:04d}-{mm}-{dd}"
        except Exception:
            res["birth_date"] = ""
    # ИИН (12 цифр) — ищем в OPTIONAL-поле MRZ по СТРУКТУРЕ формата (не regex по всей строке,
    # иначе цепляются невыровненные окна из соседних полей). ID-карта (TD1, 3 строки ~30):
    # optional в конце 1-й строки. Загранпаспорт (TD3, 2 строки ~44): optional в 2-й строке
    # (позиции 28-42). Подтверждаем: первые 6 цифр ИИН = дата рождения.
    iin = ""
    try:
        import re as _re
        raw = (d.get("raw_text", "") or "")
        dob6 = (d.get("date_of_birth") or "").strip()  # ГГММДД
        lines = [ln.replace(" ", "") for ln in raw.splitlines() if ln.strip()]
        optionals = []
        if len(lines) >= 3 and len(lines[0]) <= 32:      # TD1 — ID-карта
            optionals.append(lines[0][15:])
        if len(lines) >= 2 and len(lines[1]) >= 40:      # TD3 — загранпаспорт
            optionals.append(lines[1][28:42])
        for opt in optionals:
            digits = opt.replace("<", "")
            matched = False
            for m in _re.finditer(r"\d{12}", digits):
                if dob6 and m.group()[:6] == dob6:
                    iin = m.group()
                    matched = True
                    break
            if matched:
                break
            if len(digits) == 12 and digits.isdigit():
                iin = digits
                break
    except Exception:
        pass
    res["iin"] = iin
    # Вид документа по формату MRZ: TD1 (3 строки ~30) — ID-карта; TD3 (2 строки ~44) — паспорт.
    # Плюс тип из первой буквы MRZ: "P" = паспорт, "I"/"A"/"C" = ID/удостоверение.
    _lines = [ln.replace(" ", "") for ln in (d.get("raw_text", "") or "").splitlines() if ln.strip()]
    _doc_letter = (d.get("type") or "").upper()[:1]
    if _doc_letter == "P" or (len(_lines) == 2 and len(_lines[0]) >= 40):
        res["doc_type"] = "passport"
    else:
        res["doc_type"] = "id"
    nat = (d.get("nationality") or "").strip()
    res["citizenship"] = "Казахстан" if nat[:2] == "KA" else nat
    return res


def _process_image(data: bytes) -> bytes:
    """Обработка фото перед сохранением: автоповорот по EXIF (чтобы документ не был боком),
    ужатие разрешения (не больше 2000px по длинной стороне — качество документа не страдает) и
    сжатие в JPG качества 85. Уменьшает вес телефонных снимков в разы. Возвращает JPG-байты.
    Если это не изображение — возвращает исходные байты без изменений."""
    try:
        import io
        from PIL import Image, ImageOps
        im = Image.open(io.BytesIO(data))
        im = ImageOps.exif_transpose(im)  # поворот по EXIF
        if im.mode != "RGB":
            im = im.convert("RGB")
        # ужать, если больше 2000px по длинной стороне
        maxside = 2000
        if max(im.size) > maxside:
            ratio = maxside / max(im.size)
            im = im.resize((int(im.width * ratio), int(im.height * ratio)))
        out = io.BytesIO()
        im.save(out, format="JPEG", quality=85, optimize=True)
        return out.getvalue()
    except Exception:
        return data  # не изображение или ошибка — как есть


def _close_medical_on_cert(db, employee_id: str) -> None:
    """Скан справки загружен -> медкомиссия пройдена: Referral=COMPLETED (+result_date),
    связанное Obligation(MEDICAL_EXAM)=DONE. Останавливает отсчёт 14 дней."""
    from models import Referral, ExamStatus, Obligation, ObligationType, ObligationStatus
    refs = db.scalars(
        select(Referral).where(Referral.employee_id == employee_id)
        .where(Referral.exam_status != ExamStatus.COMPLETED)
    ).all()
    for r in refs:
        r.exam_status = ExamStatus.COMPLETED
        if not r.result_date:
            r.result_date = date.today()
        ob = db.get(Obligation, r.obligation_id)
        if ob is not None and ob.status != ObligationStatus.DONE:
            ob.status = ObligationStatus.DONE
    db.commit()


def _reopen_medical_on_cert_delete(db, employee_id: str) -> None:
    """Скан справки удалён -> откат медкомиссии: Referral обратно REFERRED, Obligation в PENDING.
    Возобновляет отсчёт (скан — единственный критерий завершения)."""
    from models import Referral, ExamStatus, Obligation, ObligationStatus
    refs = db.scalars(
        select(Referral).where(Referral.employee_id == employee_id)
        .where(Referral.exam_status == ExamStatus.COMPLETED)
    ).all()
    for r in refs:
        r.exam_status = ExamStatus.REFERRED
        r.result_date = None
        ob = db.get(Obligation, r.obligation_id)
        if ob is not None and ob.status == ObligationStatus.DONE:
            ob.status = ObligationStatus.PENDING
    db.commit()


@app.post("/employees/{employee_id}/scan/upload")
async def employee_scan_upload(
    employee_id: str,
    request: Request,
    scan_type: str = Form(...),
    files: list[UploadFile] = File(default=[]),
    db: Session = Depends(get_db),
):
    """Загрузка скана в хранилище. Доступ — кадровик/админ. Можно выбрать НЕСКОЛЬКО файлов
    (напр. паспорт на 2 страницах — два PDF): они склеиваются в один PDF под одним ключом,
    чтобы вторая загрузка не затирала первую. Один файл — сохраняется как есть.
    files необязателен на уровне FastAPI (default=[]), пустоту проверяем ниже — иначе строгий
    File(...) даёт 422 на некоторых браузерах (десктоп по-разному шлёт multiple file input)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    if scan_type not in SCAN_TYPES:
        raise HTTPException(400, "Неизвестный тип скана.")
    if not files:
        raise HTTPException(400, "Файл не выбран. Выберите PDF или фото и нажмите «Загрузить».")
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    # Читаем все выбранные файлы.
    parts = []
    for f in files:
        b = await f.read()
        if b:
            parts.append((b, f.content_type or "application/octet-stream"))
    if not parts:
        raise HTTPException(400, "Пустой файл.")

    if len(parts) == 1:
        data, ct = parts[0]
        # одиночное фото — обрабатываем (поворот/сжатие), PDF — оставляем как есть
        if "pdf" not in (ct or "").lower() and "image" in (ct or "").lower():
            data = _process_image(data)
            ct = "image/jpeg"
    else:
        # Несколько файлов — склеиваем в один PDF. Поддерживаем PDF И изображения (PNG/JPG):
        # фото страниц удостоверения конвертируются в страницы PDF, PDF-части добавляются как есть.
        # Итог — один многостраничный PDF под одним ключом (вторая страница не затирает первую).
        try:
            import io
            from pypdf import PdfWriter, PdfReader
            from PIL import Image
            writer = PdfWriter()
            for b, pct in parts:
                low = (pct or "").lower()
                if "pdf" in low:
                    reader = PdfReader(io.BytesIO(b))
                    for page in reader.pages:
                        writer.add_page(page)
                else:
                    # изображение -> обработка (поворот/сжатие) -> одностраничный PDF -> страница
                    pb = _process_image(b)
                    im = Image.open(io.BytesIO(pb))
                    if im.mode != "RGB":
                        im = im.convert("RGB")
                    tmp = io.BytesIO()
                    im.save(tmp, format="PDF")
                    tmp.seek(0)
                    for page in PdfReader(tmp).pages:
                        writer.add_page(page)
            out = io.BytesIO()
            writer.write(out)
            data = out.getvalue()
            ct = "application/pdf"
        except Exception:
            raise HTTPException(400, "Не удалось объединить файлы. Загрузите страницы как PDF "
                                     "или фото (PNG/JPG) — они склеятся в один PDF.")

    if len(data) > 15 * 1024 * 1024:
        raise HTTPException(400, "Итоговый файл больше 15 МБ — сожмите страницы.")

    # Для платёжек: проверяем фамилию И сумму. Если хоть что-то не сошлось (или чужая сумма) —
    # помечаем скан «требует проверки» (метка check=1 в метаданных S3), но НЕ блокируем загрузку.
    # Проверки обёрнуты — они не должны ронять загрузку ни при каких условиях.
    _meta = None
    warn = ""
    if scan_type in PAYMENT_SCAN_TYPES:
        need_check = False
        try:
            if _payment_surname_check(data, emp.full_name or "") is False:
                need_check = True
        except Exception:
            pass
        try:
            if _payment_amount_check(data, scan_type) is False:
                need_check = True
        except Exception:
            pass
        if need_check:
            _meta = {"check": "1"}
            warn = "?warn=payment"

    eid = None if scan_type in SCAN_COMMON_TYPES else employee_id
    try:
        _s3_upload(scan_type, eid, data, ct, metadata=_meta)
    except RuntimeError as e:
        raise HTTPException(500, str(e))
    # Скан справки медкомиссии -> закрываем медобязательство (COMPLETED/DONE).
    if scan_type == "medical_certificate":
        _close_medical_on_cert(db, employee_id)
    return RedirectResponse(f"/employees/{employee_id}{warn}", status_code=303)


@app.post("/employees/{employee_id}/scan/download")
def employee_scan_download(
    employee_id: str,
    request: Request,
    scan_type: str = Form(...),
    db: Session = Depends(get_db),
):
    """Скачивание скана. Доступ — кадровик/админ (паспортные данные прорабу недоступны)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK, UserRole.PRORAB)
    if scan_type not in SCAN_TYPES:
        raise HTTPException(400, "Неизвестный тип скана.")
    emp = db.get(Employee, employee_id)
    _fio = (emp.full_name if emp else "").strip().replace(" ", "_")
    eid = None if scan_type in SCAN_COMMON_TYPES else employee_id
    try:
        data, ct = _s3_download(scan_type, eid)
    except RuntimeError as e:
        raise HTTPException(404, str(e))
    ext = "pdf" if "pdf" in ct else ("jpg" if "jpeg" in ct or "jpg" in ct else "bin")
    _type_name = SCAN_TYPES[scan_type].split('(')[0].strip().replace(' ', '_')
    # ФИО впереди, затем тип — чтобы в загрузках не путать файлы разных работников.
    fn = f"{_fio}_{_type_name}.{ext}" if _fio else f"{_type_name}.{ext}"
    return Response(content=data, media_type=ct,
                    headers={"Content-Disposition": _content_disposition(fn)})


@app.get("/employees/{employee_id}/scan/view")
def employee_scan_view(
    employee_id: str,
    request: Request,
    scan_type: str,
    db: Session = Depends(get_db),
):
    """Просмотр скана в браузере (inline, открывается в новой вкладке — ссылка target=_blank
    в карточке). Отличие от скачивания: Content-Disposition inline, а не attachment. GET —
    чтобы работала обычная ссылка. Доступ — кадровик/админ."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK, UserRole.PRORAB)
    if scan_type not in SCAN_TYPES:
        raise HTTPException(400, "Неизвестный тип скана.")
    eid = None if scan_type in SCAN_COMMON_TYPES else employee_id
    try:
        data, ct = _s3_download(scan_type, eid)
    except RuntimeError as e:
        raise HTTPException(404, str(e))
    return Response(content=data, media_type=ct,
                    headers={"Content-Disposition": "inline"})


@app.post("/employees/{employee_id}/scan/delete")
def employee_scan_delete(
    employee_id: str,
    request: Request,
    scan_type: str = Form(...),
    db: Session = Depends(get_db),
):
    """Удаление скана. Доступ — кадровик/админ."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    if scan_type not in SCAN_TYPES:
        raise HTTPException(400, "Неизвестный тип скана.")
    eid = None if scan_type in SCAN_COMMON_TYPES else employee_id
    try:
        _s3_delete(scan_type, eid)
    except RuntimeError as e:
        raise HTTPException(500, str(e))
    # Удалён скан справки -> откат медкомиссии в ожидание.
    if scan_type == "medical_certificate":
        _reopen_medical_on_cert_delete(db, employee_id)
    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


@app.post("/employees/{employee_id}/scan/confirm")
def employee_scan_confirm(
    employee_id: str,
    request: Request,
    scan_type: str = Form(...),
    db: Session = Depends(get_db),
):
    """Подтверждение платёжки «требует проверки»: кадровик сверил вручную, снимаем метку check."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    if scan_type not in SCAN_TYPES:
        raise HTTPException(400, "Неизвестный тип скана.")
    eid = None if scan_type in SCAN_COMMON_TYPES else employee_id
    try:
        _s3_clear_check(scan_type, eid)
    except RuntimeError as e:
        raise HTTPException(500, str(e))
    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


# --- Общие документы (страница кадровика/админа) --------------------------------------------
@app.get("/common-docs", response_class=HTMLResponse)
def common_docs_page(request: Request, db: Session = Depends(get_db)):
    """Страница общих документов: паспорт директора, документ-основание на адрес подразделения.
    Один файл на всех работников. Доступ — кадровик/админ, не прораб (паспортные данные)."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    present = _s3_list_common()
    rows = ""
    for dt, label in COMMON_DOC_TYPES.items():
        has = present.get(dt, False)
        status = '<span style="color:#1a7f37">загружен ✓</span>' if has else '<span class="muted">нет</span>'
        actions = ""
        if has:
            actions = f'''<form method="post" action="/common-docs/download" style="display:inline">
<input type="hidden" name="doc_type" value="{dt}">
<button type="submit" class="secondary">Скачать</button></form>
<form method="post" action="/common-docs/delete" style="display:inline"
onsubmit="return confirm(&#39;Удалить документ?&#39;)">
<input type="hidden" name="doc_type" value="{dt}">
<button type="submit" class="secondary">Удалить</button></form>'''
        rows += f'''<div style="margin:12px 0;padding:12px;border:1px solid #e6e9ee;border-radius:8px">
<b>{label}</b> — {status}
<form method="post" action="/common-docs/upload" enctype="multipart/form-data" style="margin-top:8px">
<input type="hidden" name="doc_type" value="{dt}">
<input type="file" name="file" accept="application/pdf,image/*" required style="display:block;width:100%;margin:8px 0;padding:10px;border:1px solid #d9dde3;border-radius:8px;background:#fff;font-size:16px">
<button type="submit" class="btn-full btn-upload">Загрузить</button></form>
<div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap">{actions}</div>
</div>'''
    body = f'''<section class="card">
<h1>Общие документы</h1>
<p class="muted">Документы, единые для всех работников: паспорт директора (принимающая сторона)
и документ-основание на адрес подразделения. Входят в каждый пакет Госуслуг. PDF или фото, до 15 МБ.</p>
{rows}
<a class="btn secondary" href="/employees">← К сотрудникам</a>
</section>'''
    return _render("Общие документы", body, active="common", role=request.session.get("role", ""))


@app.post("/common-docs/upload")
async def common_docs_upload(request: Request, doc_type: str = Form(...),
                             file: UploadFile = File(...), db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    if doc_type not in COMMON_DOC_TYPES:
        raise HTTPException(400, "Неизвестный тип документа.")
    data = await file.read()
    if not data:
        raise HTTPException(400, "Пустой файл.")
    if len(data) > 15 * 1024 * 1024:
        raise HTTPException(400, "Файл больше 15 МБ.")
    try:
        _s3_upload_common(doc_type, data, file.content_type or "application/octet-stream")
    except RuntimeError as e:
        raise HTTPException(500, str(e))
    return RedirectResponse("/common-docs", status_code=303)


@app.post("/common-docs/download")
def common_docs_download(request: Request, doc_type: str = Form(...), db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    if doc_type not in COMMON_DOC_TYPES:
        raise HTTPException(400, "Неизвестный тип документа.")
    try:
        data, ct = _s3_download_common(doc_type)
    except RuntimeError as e:
        raise HTTPException(404, str(e))
    fn = f"{COMMON_DOC_TYPES[doc_type].split('(')[0].strip().replace(' ', '_')}.{_ext_for(ct)}"
    return Response(content=data, media_type=ct,
                    headers={"Content-Disposition": _content_disposition(fn)})


@app.post("/common-docs/delete")
def common_docs_delete(request: Request, doc_type: str = Form(...), db: Session = Depends(get_db)):
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    if doc_type not in COMMON_DOC_TYPES:
        raise HTTPException(400, "Неизвестный тип документа.")
    try:
        _s3_delete_common(doc_type)
    except RuntimeError as e:
        raise HTTPException(500, str(e))
    return RedirectResponse("/common-docs", status_code=303)


@app.post("/employees/{employee_id}/package")
def employee_package(employee_id: str, request: Request, db: Session = Depends(get_db)):
    """Пакет для Госуслуг одним ZIP: персональные сканы (паспорт, миграционная карта, платёжка)
    + общие документы (паспорт директора, основание на адрес) + договор PDF (без печати —
    подписывается ЭЦП). Неполный пакет — отказ с перечнем недостающего."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    missing = _package_missing(emp)
    if missing:
        raise HTTPException(400, "Пакет неполный, не хватает: " + "; ".join(missing))

    import io, zipfile
    safe_name = (emp.full_name or "работник").replace(" ", "_")
    buf = io.BytesIO()
    try:
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for st, label in SCAN_TYPES.items():
                data, ct = _s3_download(st, employee_id)
                _tn = label.split('(')[0].strip().replace(' ', '_')
                z.writestr(f"{safe_name}_{_tn}.{_ext_for(ct)}", data)
            for dt, label in COMMON_DOC_TYPES.items():
                data, ct = _s3_download_common(dt)
                z.writestr(f"{label.split('(')[0].strip()}.{_ext_for(ct)}", data)
            docx_path = generate_labor_contract_docx(
                emp, position="Монтажник", salary="30000", contract_date=emp.contract_date,
            )
            pdf_path = _docx_to_pdf(docx_path)
            with open(pdf_path, "rb") as f:
                z.writestr("Трудовой_договор.pdf", f.read())
    except RuntimeError as e:
        raise HTTPException(500, f"Не удалось собрать пакет: {e}")
    except Exception:
        raise HTTPException(500, "Ошибка сборки пакета. Проверьте логи.")
    buf.seek(0)
    return Response(content=buf.read(), media_type="application/zip",
                    headers={"Content-Disposition": _content_disposition(f"Пакет_Госуслуги_{safe_name}.zip")})


@app.post("/employees/{employee_id}/termination/cancel")
def employee_termination_cancel(employee_id: str, request: Request, db: Session = Depends(get_db)):
    """Отмена оформления увольнения: снимает contract_end_date, удаляет связанные незакрытые
    обязательства (расторжение/убытие). Кадровик/админ."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)
    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")
    from models import Obligation, ObligationType, ObligationStatus
    # удалить незакрытые обязательства увольнения (расторжение/убытие)
    db.query(Obligation).filter(
        Obligation.employee_id == emp.id,
        Obligation.type.in_([ObligationType.CONTRACT_TERMINATION_NOTICE, ObligationType.DEPARTURE_NOTICE]),
        Obligation.status != ObligationStatus.DONE,
    ).delete(synchronize_session=False)
    # вернуть в работу старые обязательства, снятые при оформлении увольнения (CANCELLED -> PENDING).
    # Cron при следующем прогоне пересчитает просрочку по дедлайну. Так отмена увольнения
    # полностью откатывает и гашение старых обязательств.
    db.query(Obligation).filter(
        Obligation.employee_id == emp.id,
        Obligation.status == ObligationStatus.CANCELLED,
    ).update({Obligation.status: ObligationStatus.PENDING}, synchronize_session=False)
    emp.contract_end_date = None
    db.commit()
    request.session.pop(f"term_basis_{emp.id}", None)
    return RedirectResponse(f"/employees/{emp.id}", status_code=303)


@app.post("/employees/{employee_id}/labor_contract/cancel")
def employee_labor_contract_cancel(
    employee_id: str,
    request: Request,
    db: Session = Depends(get_db),
):
    """Отмена договора: откатывает contract_date в NULL и снимает НЕ выполненные обязательства,
    порождённые договором (CONTRACT_NOTICE — уведомление МВД, EFS1_REPORT — ЕФС-1). Выполненные
    (DONE) НЕ трогаются — они след того, что документы подавались в срок, стирать нельзя.
    Само поле contract_date общее с ручным вводом в карточке, отмена его тоже обнулит."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _cu = _current_user(request, db)
    if _cu is None:
        raise HTTPException(401, "Требуется вход")

    emp = db.get(Employee, employee_id)
    if emp is None:
        raise HTTPException(404, "Сотрудник не найден")

    # Отмена договора: админ — всегда; кадровик — только в день заключения (свежая ошибка ввода).
    # Позже кадровику нельзя (заметание следов/поздний откат обязательств) — только админ.
    if _cu.role != UserRole.ADMIN:
        if _cu.role != UserRole.KADROVIK:
            raise HTTPException(403, "Недостаточно прав")
        if emp.contract_date != date.today():
            raise HTTPException(
                403, "Отмена договора доступна кадровику только в день заключения. Обратитесь к администратору."
            )

    emp.contract_date = None

    # снять только PENDING обязательства от договора; DONE оставить
    contract_types = (ObligationType.CONTRACT_NOTICE, ObligationType.EFS1_REPORT)
    obs = db.scalars(
        select(Obligation)
        .where(Obligation.employee_id == emp.id)
        .where(Obligation.type.in_(contract_types))
        .where(Obligation.is_current == True)  # noqa: E712
        .where(Obligation.status == ObligationStatus.PENDING)
    ).all()
    for o in obs:
        db.delete(o)

    db.commit()
    return RedirectResponse(f"/employees/{employee_id}", status_code=303)


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
    return _render("Направление на медосмотр", body, active="medical", role=request.session.get("role",""))


@app.post("/medical/{employee_id}/result")
def medical_result(
    employee_id: str,
    request: Request,
    result: str = Form(...),
    db: Session = Depends(get_db),
):
    """result='done': медосмотр пройден — направление -> COMPLETED, обязательство -> DONE.
    result='failed': ВРЕМЕННОЕ тестовое поведение — направление УДАЛЯЕТСЯ, сотрудник
    возвращается в очередь на выписку, обязательство остаётся PENDING (дедлайн жив).
    ВНИМАНИЕ: это РАСХОДИТСЯ с bot.py (_handle_medical_exam_result, где 'failed' лишь
    оставляет статус без изменений). Расхождение осознанное и временное — позже 'failed'
    заменяется на статус ExamStatus.CANCELLED с сохранением истории, и обе точки (bot.py и
    webforms.py) снова синхронизируются. См. отложенную задачу по полноценной истории."""
    if not _logged_in(request):
        return RedirectResponse("/login", status_code=303)
    _require_role(request, db, UserRole.KADROVIK)

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


# --- ВРЕМЕННО: тестовый роут OCR (/ocr-test). Убрать после проверки passporteye ---------------
# Изолированный тест распознавания MRZ с фото удостоверения. Требует passporteye в requirements
# и tesseract-ocr в RAILPACK_DEPLOY_APT_PACKAGES. После теста удалить: эти 4 строки, файл
# ocr_test.py, passporteye из requirements, tesseract из apt-переменной.
try:
    import ocr_test
    ocr_test.register(app)
except Exception:
    pass  # если ocr_test.py не подключён/удалён — рабочая система не падает


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
