"""
purchases/service.py — склейка parser + matching + БД. Аналог production.py:
вся логика здесь, в routes.py только HTTP.

Порядок обработки файла и почему он такой:

  1. Хэш → проверка дубля. ДО разбора: SELECT по уникальному индексу дешевле парсинга,
     а повторная загрузка той же накладной — самый частый пользовательский сценарий
     («не понял, загрузилось или нет, нажму ещё раз»).
  2. Профиль поставщика → разбор. Профиль перекрывает автопоиск шапки.
  3. Индекс номенклатуры + регистр соответствий → сопоставление.
  4. Документ создаётся ВСЕГДА, даже если сопоставилось не всё. Статус REVIEW честнее,
     чем отказ грузить: отказ заставляет человека править Excel, а это потеря оригинала.
  5. Профиль сохраняется после удачного разбора — следующий файл этого поставщика
     пойдёт без эвристики.

Что здесь сознательно НЕ делается:

  - Не проводится автоматически. POSTED ставит только человек через post_receipt().
  - Не пишется в регистр соответствий по результату нечёткого сопоставления. Только
    confirm_line() после явного подтверждения. Автозапись догадки закрепляет ошибку
    навсегда: следующая накладная получит её уже как LEARNED, с максимальным доверием.
  - Не создаётся номенклатура «на лету» из строки накладной. Только явным вызовом
    create_nomenclature_from_line(). Иначе справочник за квартал зарастёт дублями
    вида «Арматура А500С ф12 ГОСТ 34028» рядом с существующей «Арматура А500С ф12» —
    ровно тем, из-за чего сопоставление и ломается.

Транзакции: каждая публичная функция сама делает commit, как в production.py.
Загрузка файла — одна транзакция на документ со строками: половина накладной в базе
хуже, чем её отсутствие.
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from models import (
    GoodsReceipt,
    GoodsReceiptLine,
    Nomenclature,
    ReceiptStatus,
    Supplier,
    SupplierNomenclatureMap,
    SupplierParseProfile,
)

from .matching import (
    Candidate,
    CandidateIndex,
    MatchLevel,
    match_lines,
    supplier_key,
)
from .parser import (
    ParseError,
    ParseResult,
    SupplierProfile,
    file_hash,
    parse_invoice,
    profile_from_result,
)

log = logging.getLogger("purchases.service")


class DuplicateReceipt(Exception):
    """Файл с таким хэшем уже загружен. Несёт ссылку на существующий документ,
    чтобы роут показал его, а не просто ругнулся."""

    def __init__(self, receipt: GoodsReceipt):
        self.receipt = receipt
        super().__init__(f"Накладная уже загружена (документ от {receipt.created_at:%d.%m.%Y}).")


class NotPostable(Exception):
    """Документ нельзя провести — в тексте перечислено, что мешает."""


# --- Справочники: загрузка в память -----------------------------------------

def build_candidate_index(db: Session) -> CandidateIndex:
    """Индекс активной номенклатуры. Строится один раз на загрузку накладной.

    supplier_articles подтягиваются из регистра соответствий: если поставщик когда-то
    прислал артикул и человек его подтвердил, этот артикул работает как ключ и для
    других поставщиков тоже — артикулы производителя часто совпадают.
    """
    rows = db.scalars(
        select(Nomenclature).where(Nomenclature.is_active.is_(True))
    ).all()

    arts: dict[str, list[str]] = {}
    for nid, art in db.execute(
        select(SupplierNomenclatureMap.nomenclature_id,
               SupplierNomenclatureMap.supplier_article_raw)
        .where(SupplierNomenclatureMap.supplier_article_raw.isnot(None))
    ).all():
        arts.setdefault(nid, []).append(art)

    return CandidateIndex([
        Candidate(
            id=n.id,
            name=n.name,
            unit=n.base_unit or "",
            article=n.article,
            supplier_articles=tuple(arts.get(n.id, ())),
        )
        for n in rows
    ])


def load_learned(db: Session, supplier_id: str) -> dict[str, str]:
    """{supplier_key: nomenclature_id} для одного поставщика."""
    rows = db.execute(
        select(SupplierNomenclatureMap.supplier_key,
               SupplierNomenclatureMap.nomenclature_id)
        .where(SupplierNomenclatureMap.supplier_id == supplier_id)
    ).all()
    return {k: v for k, v in rows}


def get_profile(db: Session, supplier_id: str) -> SupplierProfile | None:
    """Сохранённая раскладка файла поставщика → структура для parser."""
    row = db.scalar(
        select(SupplierParseProfile).where(SupplierParseProfile.supplier_id == supplier_id)
    )
    if row is None or not row.header_row:
        return None
    try:
        cols = json.loads(row.columns_json or "{}")
    except json.JSONDecodeError:
        log.warning("Профиль поставщика %s: битый columns_json, игнорирую", supplier_id)
        return None
    return SupplierProfile(
        sheet=row.sheet,
        header_row=row.header_row,
        columns={k: int(v) for k, v in cols.items()},
        skip_rows_after_header=row.skip_rows_after_header or 0,
    )


def save_profile(db: Session, supplier_id: str, res: ParseResult) -> None:
    """Запомнить раскладку после удачного разбора. Перезаписывает молча: поставщик
    мог сменить форму файла, и актуальна последняя сработавшая."""
    prof = profile_from_result(res)
    row = db.scalar(
        select(SupplierParseProfile).where(SupplierParseProfile.supplier_id == supplier_id)
    )
    if row is None:
        row = SupplierParseProfile(supplier_id=supplier_id)
        db.add(row)
    row.sheet = prof.sheet
    row.header_row = prof.header_row
    row.columns_json = json.dumps(prof.columns, ensure_ascii=False)
    row.skip_rows_after_header = prof.skip_rows_after_header
    row.updated_at = datetime.utcnow()


# --- Загрузка накладной -----------------------------------------------------

@dataclass
class LoadOutcome:
    receipt: GoodsReceipt
    auto: int
    fuzzy: int
    unmatched: int
    warnings: list[str]

    @property
    def auto_share(self) -> float:
        total = self.auto + self.fuzzy + self.unmatched
        return (self.auto / total * 100.0) if total else 0.0


def find_by_hash(db: Session, data: bytes) -> GoodsReceipt | None:
    return db.scalar(
        select(GoodsReceipt).where(GoodsReceipt.file_hash == file_hash(data))
    )


def load_receipt(
    db: Session,
    *,
    data: bytes,
    filename: str | None,
    supplier_id: str,
    user_id: str | None = None,
    number: str | None = None,
    doc_date: date | None = None,
    titul_id: str | None = None,
) -> LoadOutcome:
    """Главная точка входа: файл → документ «Поступление товаров и услуг».

    Бросает DuplicateReceipt, если файл уже грузили, и ParseError, если разобрать
    не вышло. Роут ловит оба и показывает разное: дубль — ссылкой на документ,
    ошибку разбора — формой ручного указания колонок.
    """
    existing = find_by_hash(db, data)
    if existing is not None:
        raise DuplicateReceipt(existing)

    profile = get_profile(db, supplier_id)
    res = parse_invoice(data, profile=profile)          # ParseError уходит наверх

    index = build_candidate_index(db)
    learned = load_learned(db, supplier_id)
    matches, stats = match_lines(res.lines, index, learned=learned, supplier_id=supplier_id)

    receipt = GoodsReceipt(
        supplier_id=supplier_id,
        number=number,
        doc_date=doc_date,
        titul_id=titul_id,
        status=ReceiptStatus.DRAFT,
        amount_declared=res.declared_amount,
        amount_lines=res.lines_total,
        file_hash=res.file_hash,
        file_name=filename,
        auto_matched=stats.auto,
        lines_count=stats.total,
        warnings=json.dumps(res.warnings, ensure_ascii=False) if res.warnings else None,
        created_by=user_id,
    )
    db.add(receipt)
    db.flush()  # нужен receipt.id для строк

    for i, (ln, m) in enumerate(zip(res.lines, matches), start=1):
        db.add(GoodsReceiptLine(
            receipt_id=receipt.id,
            line_no=i,
            row_in_file=ln.row,
            name_raw=ln.name,
            article_raw=ln.article,
            unit_raw=ln.unit_raw,
            unit=ln.unit,
            qty=ln.qty,
            price=ln.price,
            amount=ln.amount if ln.amount is not None else ln.computed_amount,
            vat_rate=ln.vat_rate,
            nomenclature_id=m.nomenclature_id,
            match_level=m.level.value,
            unit_mismatch=m.unit_mismatch,
        ))

    unmatched = stats.total - stats.auto
    if unmatched or res.warnings or stats.unit_warnings:
        receipt.status = ReceiptStatus.REVIEW

    save_profile(db, supplier_id, res)
    db.commit()

    log.info(
        "Загружена накладная %s: строк %s, авто %s (%.0f%%), поставщик %s",
        receipt.id, stats.total, stats.auto, stats.auto_share, supplier_id,
    )
    return LoadOutcome(
        receipt=receipt,
        auto=stats.auto,
        fuzzy=stats.fuzzy,
        unmatched=stats.none,
        warnings=list(res.warnings),
    )


# --- Ручное сопоставление ---------------------------------------------------

def confirm_line(
    db: Session,
    *,
    line_id: str,
    nomenclature_id: str,
    user_id: str | None = None,
    learn_it: bool = True,
) -> GoodsReceiptLine:
    """Человек выбрал номенклатуру для строки.

    learn_it=True записывает связку в регистр — со следующей накладной этого поставщика
    строка подтянется сама. Отключать имеет смысл для разовых поставок, где строка
    заведомо больше не повторится: мусор в регистре мешает не меньше, чем его нехватка.
    """
    line = db.get(GoodsReceiptLine, line_id)
    if line is None:
        raise ValueError("Строка не найдена.")
    receipt = db.get(GoodsReceipt, line.receipt_id)
    if receipt.status in (ReceiptStatus.POSTED, ReceiptStatus.EXPORTED):
        raise NotPostable("Документ уже проведён — строки не редактируются.")

    nom = db.get(Nomenclature, nomenclature_id)
    if nom is None:
        raise ValueError("Номенклатура не найдена.")

    line.nomenclature_id = nomenclature_id
    line.match_level = "MANUAL"
    line.unit_mismatch = bool(
        line.unit and nom.base_unit and line.unit.strip().lower() != nom.base_unit.strip().lower()
    )

    if learn_it:
        key = supplier_key(line.article_raw, line.name_raw)
        existing = db.scalar(
            select(SupplierNomenclatureMap).where(
                SupplierNomenclatureMap.supplier_id == receipt.supplier_id,
                SupplierNomenclatureMap.supplier_key == key,
            )
        )
        if existing is None:
            db.add(SupplierNomenclatureMap(
                supplier_id=receipt.supplier_id,
                supplier_key=key,
                supplier_name_raw=line.name_raw,
                supplier_article_raw=line.article_raw,
                nomenclature_id=nomenclature_id,
                confirmed_by=user_id,
            ))
        else:
            # Перепривязка: человек исправил прошлое решение. Перезаписываем, но в лог —
            # частые перепривязки одного ключа означают дубли в справочнике.
            if existing.nomenclature_id != nomenclature_id:
                log.info("Перепривязка ключа %s: %s → %s",
                         key, existing.nomenclature_id, nomenclature_id)
            existing.nomenclature_id = nomenclature_id
            existing.confirmed_by = user_id
            existing.confirmed_at = datetime.utcnow()

    _refresh_status(db, receipt)
    db.commit()
    return line


def create_nomenclature_from_line(
    db: Session,
    *,
    line_id: str,
    name: str | None = None,
    base_unit: str | None = None,
    article: str | None = None,
    user_id: str | None = None,
) -> Nomenclature:
    """Завести новую позицию справочника из строки накладной и сразу привязать.

    name по умолчанию берётся из накладной как есть. Это осознанный компромисс:
    название поставщика редко совпадает с вашей номенклатурной дисциплиной, поэтому
    в форме поле должно быть редактируемым, а не подставляться молча.
    """
    line = db.get(GoodsReceiptLine, line_id)
    if line is None:
        raise ValueError("Строка не найдена.")

    nom = Nomenclature(
        name=(name or line.name_raw).strip(),
        article=article,
        base_unit=(base_unit or line.unit or "шт"),
        vat_rate=line.vat_rate,
    )
    db.add(nom)
    db.flush()
    confirm_line(db, line_id=line_id, nomenclature_id=nom.id, user_id=user_id)
    return nom


def rematch_receipt(db: Session, receipt_id: str) -> LoadOutcome:
    """Пересопоставить строки документа заново — после чистки справочника или
    пополнения регистра. Файл не перечитывается, исходные строки на месте.

    Уже подтверждённые вручную (MANUAL) не трогаются: пересопоставление не должно
    отменять решение человека.
    """
    receipt = db.get(GoodsReceipt, receipt_id)
    if receipt is None:
        raise ValueError("Документ не найден.")
    if receipt.status in (ReceiptStatus.POSTED, ReceiptStatus.EXPORTED):
        raise NotPostable("Документ проведён — пересопоставление недоступно.")

    index = build_candidate_index(db)
    learned = load_learned(db, receipt.supplier_id)

    lines = db.scalars(
        select(GoodsReceiptLine)
        .where(GoodsReceiptLine.receipt_id == receipt_id)
        .order_by(GoodsReceiptLine.line_no)
    ).all()

    # match_lines ждёт объекты с .name/.article/.unit — подсовываем адаптер
    class _L:
        def __init__(self, r):
            self.name, self.article, self.unit = r.name_raw, r.article_raw, r.unit or ""

    open_lines = [ln for ln in lines if ln.match_level != "MANUAL"]
    matches, stats = match_lines([_L(ln) for ln in open_lines], index,
                                 learned=learned, supplier_id=receipt.supplier_id)

    for ln, m in zip(open_lines, matches):
        ln.nomenclature_id = m.nomenclature_id
        ln.match_level = m.level.value
        ln.unit_mismatch = m.unit_mismatch

    manual = len(lines) - len(open_lines)
    receipt.auto_matched = stats.auto + manual
    _refresh_status(db, receipt)
    db.commit()
    return LoadOutcome(receipt=receipt, auto=stats.auto + manual,
                       fuzzy=stats.fuzzy, unmatched=stats.none, warnings=[])


# --- Статус и проведение ----------------------------------------------------

def _refresh_status(db: Session, receipt: GoodsReceipt) -> None:
    """DRAFT ↔ REVIEW по факту наличия проблем. POSTED/EXPORTED не трогаем."""
    if receipt.status in (ReceiptStatus.POSTED, ReceiptStatus.EXPORTED,
                          ReceiptStatus.CANCELLED):
        return
    problems = db.scalar(
        select(func.count(GoodsReceiptLine.id)).where(
            GoodsReceiptLine.receipt_id == receipt.id,
            (GoodsReceiptLine.nomenclature_id.is_(None))
            | (GoodsReceiptLine.unit_mismatch.is_(True)),
        )
    ) or 0
    receipt.status = ReceiptStatus.REVIEW if problems else ReceiptStatus.DRAFT


def check_postable(db: Session, receipt_id: str) -> list[str]:
    """Что мешает провести. Пустой список — можно проводить.

    Возвращается списком, а не первой ошибкой: человек должен увидеть весь объём
    работы сразу, а не чинить по одному пункту с перезагрузкой формы.
    """
    receipt = db.get(GoodsReceipt, receipt_id)
    if receipt is None:
        return ["Документ не найден."]

    problems: list[str] = []
    if receipt.status in (ReceiptStatus.POSTED, ReceiptStatus.EXPORTED):
        return ["Документ уже проведён."]
    if receipt.status is ReceiptStatus.CANCELLED:
        return ["Документ отменён."]

    unmatched = db.scalar(
        select(func.count(GoodsReceiptLine.id)).where(
            GoodsReceiptLine.receipt_id == receipt_id,
            GoodsReceiptLine.nomenclature_id.is_(None),
        )
    ) or 0
    if unmatched:
        problems.append(f"Не сопоставлено строк: {unmatched}.")

    if not receipt.number:
        problems.append("Не указан номер накладной поставщика.")
    if not receipt.doc_date:
        problems.append("Не указана дата накладной.")

    if receipt.amount_declared is not None and receipt.amount_lines is not None:
        diff = abs(receipt.amount_lines - receipt.amount_declared)
        if diff > Decimal("0.01") * max(receipt.lines_count, 1):
            problems.append(
                f"Сумма строк {receipt.amount_lines} не сходится с итогом "
                f"{receipt.amount_declared} (расхождение {diff})."
            )

    if not receipt.scan_key:
        # Не блокирует, но должно быть видно: без оригинала спор с поставщиком
        # разрешить нечем.
        problems.append("Не приложен оригинал накладной (можно провести, но нежелательно).")

    return problems


def post_receipt(db: Session, *, receipt_id: str, user_id: str,
                 force: bool = False) -> GoodsReceipt:
    """Провести документ. force пропускает мягкие замечания, но НЕ пропускает
    несопоставленные строки: проводить документ с неизвестной номенклатурой нельзя,
    в 1С такая строка всё равно не ляжет."""
    receipt = db.get(GoodsReceipt, receipt_id)
    if receipt is None:
        raise ValueError("Документ не найден.")

    problems = check_postable(db, receipt_id)
    hard = [p for p in problems if p.startswith("Не сопоставлено")
            or p.startswith("Документ")]
    blocking = hard if force else problems
    if blocking:
        raise NotPostable("Проведение невозможно:\n- " + "\n- ".join(blocking))

    receipt.status = ReceiptStatus.POSTED
    receipt.posted_by = user_id
    receipt.posted_at = datetime.utcnow()
    db.commit()
    log.info("Документ %s проведён пользователем %s", receipt_id, user_id)
    return receipt


def cancel_receipt(db: Session, *, receipt_id: str, user_id: str) -> GoodsReceipt:
    """Отменить документ. Строки и файл остаются: отмена — это состояние, а не удаление.
    Хэш тоже остаётся, поэтому повторно загрузить тот же файл не выйдет, пока документ
    не удалён физически. Это защита от «отменю и залью заново», которая маскирует ошибку.
    """
    receipt = db.get(GoodsReceipt, receipt_id)
    if receipt is None:
        raise ValueError("Документ не найден.")
    if receipt.status is ReceiptStatus.EXPORTED:
        raise NotPostable("Документ выгружен в 1С — отменять нужно там.")
    receipt.status = ReceiptStatus.CANCELLED
    db.commit()
    log.info("Документ %s отменён пользователем %s", receipt_id, user_id)
    return receipt


# --- Выборки для списков ----------------------------------------------------

def get_receipts(db: Session, *, limit: int = 100,
                 status: ReceiptStatus | None = None) -> list[GoodsReceipt]:
    q = select(GoodsReceipt).order_by(GoodsReceipt.created_at.desc()).limit(limit)
    if status is not None:
        q = q.where(GoodsReceipt.status == status)
    return list(db.scalars(q).all())


def get_receipt_lines(db: Session, receipt_id: str) -> list[GoodsReceiptLine]:
    return list(db.scalars(
        select(GoodsReceiptLine)
        .where(GoodsReceiptLine.receipt_id == receipt_id)
        .order_by(GoodsReceiptLine.line_no)
    ).all())


def get_suppliers(db: Session) -> list[Supplier]:
    return list(db.scalars(select(Supplier).order_by(Supplier.name)).all())


def matching_health(db: Session, days: int = 90) -> dict:
    """Сводка качества сопоставления за период — на экран загрузки и в отчёты.

    Считается по сохранённым auto_matched/lines_count, а не пересчётом: нужна картина
    на момент приёмки, а не то, как сопоставилось бы сегодняшним справочником.
    """
    cutoff = datetime.utcnow() - timedelta(days=days)
    rows = db.execute(
        select(func.sum(GoodsReceipt.auto_matched), func.sum(GoodsReceipt.lines_count),
               func.count(GoodsReceipt.id))
        .where(GoodsReceipt.created_at >= cutoff,
               GoodsReceipt.status != ReceiptStatus.CANCELLED)
    ).one()
    auto, total, docs = (rows[0] or 0), (rows[1] or 0), (rows[2] or 0)
    return {
        "docs": docs,
        "lines": total,
        "auto": auto,
        "auto_share": (auto / total * 100.0) if total else 0.0,
    }
