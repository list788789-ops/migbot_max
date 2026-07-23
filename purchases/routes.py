"""
purchases/routes.py — HTTP-слой раздела «Покупки». Логики здесь нет, вся в service.py.

Подключение в webforms.py — одной строкой рядом с register_1c_routes(app):

    from purchases.routes import register_purchases_routes
    register_purchases_routes(
        app,
        render=_render,
        get_db=get_db,
        logged_in=_logged_in,
        current_user=_current_user,
        actor_name=_actor_name,
    )

Зависимости передаются явно, а не импортируются из webforms. Причина простая:
webforms импортирует этот модуль, и обратный импорт даст цикл. Плюс так роуты
тестируются без поднятия всего приложения.

ДОСТУП. Раздел виден только BUHGALTER и ADMIN. Кадровик и прораб сюда не ходят:
у прораба вообще нет права записи в основные таблицы (см. UserRole в models.py),
а кадровику поставщики и номенклатура не нужны. Проверка — в _require на каждом
роуте, не только в навигации: скрытая ссылка не защита.

ФОРМА СВЕРКИ. Строки показываются светофором:
  зелёное  — сопоставлено автоматически (LEARNED/ARTICLE/NAME_EXACT), свёрнуто;
  жёлтое   — есть похожие кандидаты, нужен клик;
  красное  — не найдено, завести позицию или выбрать вручную.
При 300 позициях глазами обрабатываются только жёлтые и красные. Зелёные не мозолят
глаза — иначе смысл автосопоставления теряется, человек всё равно листает всё.

Все POST отвечают редиректом 303 (как остальные формы проекта): без этого F5 после
отправки повторяет действие, а повторное проведение документа — не то, что нужно.
"""

from __future__ import annotations

import html as _html
import logging
from datetime import date, datetime

from fastapi import Depends, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from sqlalchemy.orm import Session

from models import (
    GoodsReceipt,
    GoodsReceiptLine,
    Nomenclature,
    ReceiptStatus,
    Supplier,
    UserRole,
)

from . import service as svc
from .matching import CandidateIndex, MatchLevel
from .parser import HeaderNotFound, ParseError

log = logging.getLogger("purchases.routes")

ALLOWED_ROLES = (UserRole.BUHGALTER, UserRole.ADMIN)

STATUS_BADGE = {
    ReceiptStatus.DRAFT: ("neutral", "черновик"),
    ReceiptStatus.REVIEW: ("orange", "на проверке"),
    ReceiptStatus.POSTED: ("green", "проведён"),
    ReceiptStatus.EXPORTED: ("green", "выгружен в 1С"),
    ReceiptStatus.CANCELLED: ("red", "отменён"),
}


def esc(v) -> str:
    return _html.escape(str(v)) if v is not None else ""


def _money(v) -> str:
    return f"{v:,.2f}".replace(",", " ") if v is not None else "—"


def register_purchases_routes(app, *, render, get_db, logged_in, current_user,
                              actor_name=None) -> None:
    """Регистрирует роуты раздела. render/get_db/... — функции из webforms.py."""

    def _require(request: Request, db: Session):
        """Пользователь с правом на раздел, иначе None."""
        user = current_user(request, db)
        if user is None or user.role not in ALLOWED_ROLES:
            return None
        return user

    def _guard(request: Request, db: Session):
        """Возвращает (user, redirect). redirect не None → сразу вернуть его."""
        if not logged_in(request):
            return None, RedirectResponse("/login", status_code=303)
        user = _require(request, db)
        if user is None:
            return None, RedirectResponse("/", status_code=303)
        return user, None

    def _page(request, title, body):
        return render(title, body, active="purchases",
                      role=request.session.get("role", ""))

    # ---------------------------------------------------------------- список

    @app.get("/purchases", response_class=HTMLResponse)
    def purchases_hub(request: Request, db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir

        health = svc.matching_health(db)
        receipts = svc.get_receipts(db, limit=50)
        suppliers = {s.id: s.name for s in svc.get_suppliers(db)}

        share = health["auto_share"]
        tone = "green" if share >= 90 else ("orange" if share >= 70 else "red")
        health_block = (
            f'<section><h2>Качество сопоставления за 90 дней</h2>'
            f'<div class="card wo-card">'
            f'<span class="badge {tone}">{share:.0f}% автоматически</span>'
            f'<div class="muted-line">Документов {health["docs"]}, '
            f'строк {health["lines"]}, из них вручную {health["lines"] - health["auto"]}.</div>'
            f'<details class="field-help"><summary><span class="i">i</span>что это значит</summary>'
            f'<p>Доля строк, сопоставленных без участия человека. Ниже 90% — обычно '
            f'признак дублей в справочнике номенклатуры, а не проблемы разбора. '
            f'Отчёты по ценам считаются по сопоставленным строкам, поэтому при низкой '
            f'доле им нельзя доверять.</p></details>'
            f'</div></section>'
        ) if health["lines"] else ""

        if receipts:
            cards = []
            for r in receipts:
                cls, label = STATUS_BADGE.get(r.status, ("neutral", r.status.value))
                num = esc(r.number) if r.number else "б/н"
                dt = r.doc_date.strftime("%d.%m.%Y") if r.doc_date else "дата не указана"
                unmatched = r.lines_count - r.auto_matched
                extra = (f' · <span class="badge orange">вручную {unmatched}</span>'
                         if unmatched else "")
                cards.append(
                    f'<div class="card wo-card">'
                    f'<div class="wo-title">{esc(suppliers.get(r.supplier_id, "—"))} · {num}</div>'
                    f'<div class="wo-meta">{dt} · {r.lines_count} поз. · '
                    f'{_money(r.amount_lines)} ₽</div>'
                    f'<div><span class="badge {cls}">{label}</span>{extra}</div>'
                    f'<div class="wo-actions">'
                    f'<a class="btn secondary" href="/purchases/receipts/{r.id}">Открыть</a>'
                    f'</div></div>'
                )
            list_block = f'<section><h2>Поступления</h2>{"".join(cards)}</section>'
        else:
            list_block = ('<section><h2>Поступления</h2>'
                          '<p class="muted">Пока ничего не загружено.</p></section>')

        body = (
            '<header class="org"><div class="page-title">Покупки</div></header>'
            '<a class="btn btn-full" href="/purchases/upload">Загрузить накладную</a>'
            f'{health_block}{list_block}'
            '<section><h2>Справочники</h2>'
            '<a class="btn secondary" href="/purchases/suppliers">Поставщики</a>'
            '<a class="btn secondary" href="/purchases/nomenclature">Номенклатура</a>'
            '</section>'
        )
        return _page(request, "Покупки", body)

    # -------------------------------------------------------------- загрузка

    @app.get("/purchases/upload", response_class=HTMLResponse)
    def upload_form(request: Request, err: str | None = None,
                    db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir

        suppliers = svc.get_suppliers(db)
        if not suppliers:
            body = ('<header class="org"><div class="page-title">Загрузка накладной</div></header>'
                    '<div class="warning-banner">Сначала заведите поставщика — '
                    'без него не к чему привязать профиль разбора файла.</div>'
                    '<a class="btn" href="/purchases/suppliers">К поставщикам</a>')
            return _page(request, "Загрузка накладной", body)

        opts = "".join(
            f'<option value="{s.id}">{esc(s.name)}'
            f'{" · ЭДО" if s.is_edo else ""}</option>' for s in suppliers
        )
        err_block = f'<div class="warning-banner">{esc(err)}</div>' if err else ""

        body = (
            '<header class="org"><div class="page-title">Загрузка накладной</div></header>'
            f'{err_block}'
            '<section><form method="post" action="/purchases/upload" '
            'enctype="multipart/form-data">'
            '<label>Поставщик</label>'
            f'<select name="supplier_id" required>{opts}</select>'
            '<details class="field-help"><summary><span class="i">i</span>'
            'почему поставщика выбираем вручную</summary>'
            '<p>По нему подтягивается сохранённая раскладка файла и все прошлые '
            'сопоставления номенклатуры. Определять поставщика из содержимого файла '
            'ненадёжно: ИНН в шапке есть не у всех.</p></details>'
            '<label>Номер накладной</label>'
            '<input type="text" name="number" placeholder="напр. 4417">'
            '<label>Дата накладной</label>'
            '<input type="date" name="doc_date">'
            '<label>Файл Excel (.xlsx)</label>'
            '<input type="file" name="file" accept=".xlsx" required '
            'style="margin:6px 0 12px">'
            '<button type="submit" class="btn-full">Разобрать</button>'
            '</form></section>'
            '<p class="muted">Оригинал файла сохраняется вместе с документом — '
            'спор с поставщиком решается сверкой с ним, а не с разобранными строками.</p>'
        )
        return _page(request, "Загрузка накладной", body)

    @app.post("/purchases/upload")
    async def upload_submit(
        request: Request,
        supplier_id: str = Form(...),
        number: str = Form(""),
        doc_date: str = Form(""),
        file: UploadFile = File(...),
        db: Session = Depends(get_db),
    ):
        user, redir = _guard(request, db)
        if redir:
            return redir

        data = await file.read()
        if not data:
            return RedirectResponse("/purchases/upload?err=Файл пустой.", status_code=303)

        parsed_date = None
        if doc_date:
            try:
                parsed_date = date.fromisoformat(doc_date)
            except ValueError:
                parsed_date = None

        try:
            out = svc.load_receipt(
                db,
                data=data,
                filename=file.filename,
                supplier_id=supplier_id,
                user_id=user.id,
                number=number.strip() or None,
                doc_date=parsed_date,
            )
        except svc.DuplicateReceipt as e:
            # Не ошибка пользователя, а самый частый сценарий: показываем документ.
            return RedirectResponse(
                f"/purchases/receipts/{e.receipt.id}?dup=1", status_code=303
            )
        except HeaderNotFound:
            return RedirectResponse(
                "/purchases/upload?err=Не удалось найти таблицу в файле. "
                "Проверьте, что это накладная, а не счёт или письмо.",
                status_code=303,
            )
        except ParseError as e:
            return RedirectResponse(f"/purchases/upload?err={esc(e)}", status_code=303)

        # Оригинал в S3 тем же механизмом, что сканы приказов.
        try:
            from s3_storage import _s3_upload_entity, _scan_key
            scan_type = f"receipt_{out.receipt.id}"
            _s3_upload_entity(scan_type, None, data,
                              file.content_type or "application/octet-stream")
            out.receipt.scan_key = _scan_key(scan_type, None)
            db.commit()
        except Exception as e:  # noqa: BLE001 — потеря скана не должна ронять загрузку
            log.warning("Оригинал накладной %s не сохранён: %s", out.receipt.id, e)

        return RedirectResponse(f"/purchases/receipts/{out.receipt.id}", status_code=303)

    # --------------------------------------------------------------- документ

    @app.get("/purchases/receipts/{receipt_id}", response_class=HTMLResponse)
    def receipt_page(receipt_id: str, request: Request, dup: int = 0,
                     db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir

        r = db.get(GoodsReceipt, receipt_id)
        if r is None:
            raise HTTPException(404, "Документ не найден.")
        supplier = db.get(Supplier, r.supplier_id)
        lines = svc.get_receipt_lines(db, receipt_id)
        editable = r.status in (ReceiptStatus.DRAFT, ReceiptStatus.REVIEW)

        index = svc.build_candidate_index(db) if editable else None
        noms = {n.id: n for n in db.query(Nomenclature).all()} if lines else {}

        cls, label = STATUS_BADGE.get(r.status, ("neutral", r.status.value))
        dup_block = ('<div class="warning-banner">Этот файл уже загружался — '
                     'открыт существующий документ.</div>') if dup else ""

        problems = svc.check_postable(db, receipt_id) if editable else []
        prob_block = ""
        if problems:
            items = "".join(f"<li>{esc(p)}</li>" for p in problems)
            prob_block = (f'<div class="warning-banner">Перед проведением:'
                          f'<ul style="margin:6px 0 0;padding-left:18px;font-weight:400">'
                          f'{items}</ul></div>')

        # --- строки со светофором
        green, yellow, red = [], [], []
        for ln in lines:
            nom = noms.get(ln.nomenclature_id) if ln.nomenclature_id else None
            head = (f'<div class="wo-title">{ln.line_no}. {esc(ln.name_raw)}</div>'
                    f'<div class="wo-meta">{ln.qty} {esc(ln.unit_raw or ln.unit or "")}'
                    f'{" · " + _money(ln.price) + " ₽" if ln.price else ""}'
                    f'{" · " + _money(ln.amount) + " ₽" if ln.amount else ""}</div>')

            if nom is not None:
                warn = ('<span class="badge orange">ед. изм. разошлись: '
                        f'{esc(ln.unit)} / {esc(nom.base_unit)}</span>'
                        if ln.unit_mismatch else "")
                lvl = ln.match_level or ""
                mark = "вручную" if lvl == "MANUAL" else MatchLevel(lvl).label if lvl in MatchLevel.__members__ else lvl
                card = (f'<div class="card wo-card">{head}'
                        f'<div class="muted-line">→ {esc(nom.name)} '
                        f'<span class="muted">({esc(mark)})</span></div>{warn}'
                        + (_relink_form(ln, editable)) + '</div>')
                (yellow if ln.unit_mismatch else green).append(card)
                continue

            # не сопоставлено — предлагаем кандидатов
            proposals = []
            if index is not None:
                from .matching import norm_name
                proposals = index.fuzzy(norm_name(ln.name_raw))

            if proposals:
                opts = "".join(
                    f'<option value="{p.candidate.id}">{esc(p.candidate.name)} '
                    f'— {p.score:.0f}%</option>' for p in proposals
                )
                yellow.append(
                    f'<div class="card wo-card">{head}'
                    f'<span class="badge orange">похожие найдены</span>'
                    f'<form method="post" action="/purchases/lines/{ln.id}/confirm" '
                    f'style="margin-top:8px">'
                    f'<select name="nomenclature_id" required>{opts}</select>'
                    f'<button type="submit">Подтвердить</button>'
                    f'</form>'
                    f'{_new_nom_form(ln)}</div>'
                )
            else:
                red.append(
                    f'<div class="card wo-card">{head}'
                    f'<span class="badge red">нет в справочнике</span>'
                    f'{_new_nom_form(ln)}</div>'
                )

        blocks = ""
        if red:
            blocks += f'<section><h2>Новые позиции — {len(red)}</h2>{"".join(red)}</section>'
        if yellow:
            blocks += (f'<section><h2>Требуют подтверждения — {len(yellow)}</h2>'
                       f'{"".join(yellow)}</section>')
        if green:
            blocks += (f'<section><h2>Сопоставлено — {len(green)}</h2>'
                       f'<details><summary class="muted" style="cursor:pointer;'
                       f'padding:6px 0">показать {len(green)} строк</summary>'
                       f'<div style="margin-top:10px">{"".join(green)}</div>'
                       f'</details></section>')

        actions = ""
        if editable:
            actions = (
                '<section><h2>Действия</h2>'
                f'<form method="post" action="/purchases/receipts/{receipt_id}/post" '
                'style="display:inline">'
                '<button type="submit">Провести</button></form>'
                f'<form method="post" action="/purchases/receipts/{receipt_id}/rematch" '
                'style="display:inline">'
                '<button type="submit" class="secondary">Пересопоставить</button></form>'
                f'<form method="post" action="/purchases/receipts/{receipt_id}/cancel" '
                'style="display:inline">'
                '<button type="submit" class="ghost-danger">Отменить документ</button>'
                '</form></section>'
            )

        orig = (f'<a class="btn secondary" href="/purchases/receipts/{receipt_id}/file">'
                f'Скачать оригинал</a>' if r.scan_key else
                '<p class="muted">Оригинал файла не сохранён.</p>')

        body = (
            f'<header class="org"><div class="org-name">'
            f'{esc(supplier.name if supplier else "—")}</div>'
            f'<div class="page-title">Поступление {esc(r.number) if r.number else "б/н"}'
            f'</div></header>'
            f'{dup_block}{prob_block}'
            f'<section><h2>Документ</h2><div class="card wo-card">'
            f'<span class="badge {cls}">{label}</span>'
            f'<div class="wo-meta">'
            f'{r.doc_date.strftime("%d.%m.%Y") if r.doc_date else "дата не указана"} · '
            f'{r.lines_count} позиций</div>'
            f'<div class="muted-line">Сумма по строкам: {_money(r.amount_lines)} ₽<br>'
            f'Итог в файле: {_money(r.amount_declared)} ₽</div>'
            f'<div class="wo-actions">{orig}</div>'
            f'</div></section>'
            f'{blocks}{actions}'
            f'<a class="btn secondary" href="/purchases">← К списку</a>'
        )
        return _page(request, "Поступление", body)

    # ------------------------------------------------------------ действия

    @app.post("/purchases/lines/{line_id}/confirm")
    def line_confirm(line_id: str, request: Request,
                     nomenclature_id: str = Form(...),
                     db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir
        line = db.get(GoodsReceiptLine, line_id)
        if line is None:
            raise HTTPException(404, "Строка не найдена.")
        rid = line.receipt_id
        try:
            svc.confirm_line(db, line_id=line_id,
                             nomenclature_id=nomenclature_id, user_id=user.id)
        except svc.NotPostable as e:
            log.info("confirm_line отклонён: %s", e)
        return RedirectResponse(f"/purchases/receipts/{rid}", status_code=303)

    @app.post("/purchases/lines/{line_id}/new-nomenclature")
    def line_new_nom(line_id: str, request: Request,
                     name: str = Form(...), base_unit: str = Form(""),
                     article: str = Form(""),
                     db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir
        line = db.get(GoodsReceiptLine, line_id)
        if line is None:
            raise HTTPException(404, "Строка не найдена.")
        rid = line.receipt_id
        svc.create_nomenclature_from_line(
            db, line_id=line_id, name=name.strip(),
            base_unit=base_unit.strip() or None,
            article=article.strip() or None, user_id=user.id,
        )
        return RedirectResponse(f"/purchases/receipts/{rid}", status_code=303)

    @app.post("/purchases/receipts/{receipt_id}/post")
    def receipt_post(receipt_id: str, request: Request,
                     db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir
        try:
            svc.post_receipt(db, receipt_id=receipt_id, user_id=user.id, force=True)
        except svc.NotPostable as e:
            log.info("Проведение %s отклонено: %s", receipt_id, e)
        return RedirectResponse(f"/purchases/receipts/{receipt_id}", status_code=303)

    @app.post("/purchases/receipts/{receipt_id}/rematch")
    def receipt_rematch(receipt_id: str, request: Request,
                        db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir
        try:
            svc.rematch_receipt(db, receipt_id)
        except (ValueError, svc.NotPostable) as e:
            log.info("Пересопоставление %s отклонено: %s", receipt_id, e)
        return RedirectResponse(f"/purchases/receipts/{receipt_id}", status_code=303)

    @app.post("/purchases/receipts/{receipt_id}/cancel")
    def receipt_cancel(receipt_id: str, request: Request,
                       db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir
        try:
            svc.cancel_receipt(db, receipt_id=receipt_id, user_id=user.id)
        except (ValueError, svc.NotPostable) as e:
            log.info("Отмена %s отклонена: %s", receipt_id, e)
        return RedirectResponse(f"/purchases/receipts/{receipt_id}", status_code=303)

    @app.get("/purchases/receipts/{receipt_id}/file")
    def receipt_file(receipt_id: str, request: Request,
                     db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir
        r = db.get(GoodsReceipt, receipt_id)
        if r is None or not r.scan_key:
            raise HTTPException(404, "Оригинал не найден.")
        from s3_storage import _s3_download_entity
        try:
            data, ct = _s3_download_entity(f"receipt_{r.id}", None)
        except RuntimeError as e:
            raise HTTPException(404, str(e))
        fn = r.file_name or f"nakladnaya_{r.number or r.id[:8]}.xlsx"
        return Response(content=data, media_type=ct,
                        headers={"Content-Disposition": f'attachment; filename="{fn}"'})

    # ----------------------------------------------------------- справочники

    @app.get("/purchases/suppliers", response_class=HTMLResponse)
    def suppliers_page(request: Request, db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir
        rows = svc.get_suppliers(db)
        cards = "".join(
            f'<div class="card wo-card"><div class="wo-title">{esc(s.name)}</div>'
            f'<div class="wo-meta">ИНН {esc(s.inn) or "—"}'
            f'{" · ЭДО" if s.is_edo else ""}</div></div>'
            for s in rows
        ) or '<p class="muted">Пока никого.</p>'

        body = (
            '<header class="org"><div class="page-title">Поставщики</div></header>'
            '<section><h2>Добавить</h2>'
            '<form method="post" action="/purchases/suppliers">'
            '<label>Наименование</label><input type="text" name="name" required>'
            '<label>ИНН</label><input type="text" name="inn">'
            '<label><input type="checkbox" name="is_edo" value="1" '
            'style="width:auto;margin-right:8px">Работает через ЭДО</label>'
            '<details class="field-help"><summary><span class="i">i</span>'
            'зачем отмечать ЭДО</summary>'
            '<p>По этому флагу считается, какая доля поставок вообще может уйти '
            'из ручной загрузки. Это цифра для решения, а не для отображения.</p>'
            '</details>'
            '<button type="submit" class="btn-full">Добавить</button>'
            '</form></section>'
            f'<section><h2>Список</h2>{cards}</section>'
            '<a class="btn secondary" href="/purchases">← Назад</a>'
        )
        return _page(request, "Поставщики", body)

    @app.post("/purchases/suppliers")
    def supplier_create(request: Request, name: str = Form(...),
                        inn: str = Form(""), is_edo: str = Form(""),
                        db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir
        db.add(Supplier(name=name.strip(), inn=(inn.strip() or None),
                        is_edo=bool(is_edo)))
        db.commit()
        return RedirectResponse("/purchases/suppliers", status_code=303)

    @app.get("/purchases/nomenclature", response_class=HTMLResponse)
    def nomenclature_page(request: Request, q: str = "",
                          db: Session = Depends(get_db)):
        user, redir = _guard(request, db)
        if redir:
            return redir
        query = db.query(Nomenclature).filter(Nomenclature.is_active.is_(True))
        if q:
            query = query.filter(Nomenclature.name.ilike(f"%{q}%"))
        rows = query.order_by(Nomenclature.name).limit(200).all()

        cards = "".join(
            f'<div class="card wo-card"><div class="wo-title">{esc(n.name)}</div>'
            f'<div class="wo-meta">{esc(n.base_unit)}'
            f'{" · арт. " + esc(n.article) if n.article else ""}</div></div>'
            for n in rows
        ) or '<p class="muted">Ничего не найдено.</p>'

        body = (
            '<header class="org"><div class="page-title">Номенклатура</div></header>'
            '<section><form method="get" action="/purchases/nomenclature">'
            f'<input type="text" name="q" placeholder="поиск" value="{esc(q)}">'
            '<button type="submit" class="secondary">Найти</button>'
            '</form></section>'
            f'<section><h2>Позиции ({len(rows)})</h2>{cards}</section>'
            '<a class="btn secondary" href="/purchases">← Назад</a>'
        )
        return _page(request, "Номенклатура", body)


# --- Кусочки разметки -------------------------------------------------------

def _new_nom_form(ln: GoodsReceiptLine) -> str:
    """Форма заведения новой позиции из строки накладной.

    Название подставляется из накладной, но остаётся редактируемым: формулировка
    поставщика редко совпадает с вашей номенклатурной дисциплиной, а молча
    записать её в справочник — способ получить дубль.
    """
    return (
        f'<details style="margin-top:8px"><summary class="muted" '
        f'style="cursor:pointer">завести новую позицию</summary>'
        f'<form method="post" action="/purchases/lines/{ln.id}/new-nomenclature" '
        f'style="margin-top:8px">'
        f'<label>Наименование</label>'
        f'<input type="text" name="name" value="{esc(ln.name_raw)}" required>'
        f'<label>Единица</label>'
        f'<input type="text" name="base_unit" value="{esc(ln.unit or "шт")}">'
        f'<label>Артикул</label>'
        f'<input type="text" name="article" value="{esc(ln.article_raw or "")}">'
        f'<button type="submit" class="secondary">Создать и привязать</button>'
        f'</form></details>'
    )


def _relink_form(ln: GoodsReceiptLine, editable: bool) -> str:
    """Перепривязка уже сопоставленной строки — свёрнуто, чтобы не мешать."""
    if not editable:
        return ""
    return (
        f'<details style="margin-top:6px"><summary class="muted" '
        f'style="cursor:pointer;font-size:13px">изменить привязку</summary>'
        f'<form method="post" action="/purchases/lines/{ln.id}/confirm" '
        f'style="margin-top:8px">'
        f'<input type="text" name="nomenclature_id" placeholder="ID номенклатуры" required>'
        f'<button type="submit" class="secondary">Привязать</button>'
        f'</form></details>'
    )
