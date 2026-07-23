"""
purchases/matching.py — сопоставление строк накладной с вашей номенклатурой.

Модуль намеренно без БД и без ORM: на вход подаются кандидаты (список Candidate),
на выход — решение с уровнем доверия. Так его можно тестировать на голых данных и
не тащить models.py в юнит-тесты. Загрузку кандидатов делает service.py.

Порядок правил — от самого надёжного к самому шаткому, первое сработавшее выигрывает:

  1. LEARNED  — в регистре соответствий уже есть связка (поставщик + его строка).
                Это то, что человек подтвердил руками в прошлый раз. Доверие максимальное.
  2. ARTICLE  — совпал артикул поставщика, записанный в карточке номенклатуры.
  3. NAME_EXACT — совпало нормализованное наименование (один в один после чистки).
  4. FUZZY    — похожее наименование. НЕ применяется автоматически: отдаётся человеку
                как список кандидатов с оценкой. Автоприменение нечёткого совпадения —
                это способ тихо положить бетон на счёт метизов.
  5. NONE     — не нашли, позиция новая.

Почему FUZZY не автоприменяется даже при score 95: в стройке названия отличаются одним
символом при принципиально разном товаре — «Арматура А500С ф12» и «Арматура А500С ф16»
дают очень высокую похожесть, а это разные позиции с разной ценой. Порог тут не спасает,
спасает только человек. Зато человеку показываем три кандидата, и вместо ввода — клик.

Нормализация названий под стройматериалы (_norm_name):
  - убираем ГОСТ/ТУ/СТО с номерами — они у поставщика есть, у вас в карточке нет;
  - «ф12», «д12», «Ø12», «диам. 12» → «d12», иначе одинаковый товар не сходится;
  - «х»/«*» между числами → «x» (кириллическая «х» в размерах — классическая ловушка);
  - множественные пробелы, кавычки, скобки — вон.
Токены дальше сортируются: «Арматура ф12 А500С» и «Арматура А500С ф12» — одно и то же.

rapidfuzz быстрее difflib в разы и на справочнике в десятки тысяч позиций разница
заметна. Но жёсткой зависимости нет: если пакета в окружении не оказалось, работает
difflib. Для прода в requirements.txt лучше положить rapidfuzz.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum

try:  # предпочитаем rapidfuzz, но не требуем его
    from rapidfuzz import fuzz as _fuzz

    def _ratio(a: str, b: str) -> float:
        return float(_fuzz.token_sort_ratio(a, b))

    FUZZ_BACKEND = "rapidfuzz"
except ImportError:  # pragma: no cover — зависит от окружения
    from difflib import SequenceMatcher

    def _ratio(a: str, b: str) -> float:
        return SequenceMatcher(None, a, b).ratio() * 100.0

    FUZZ_BACKEND = "difflib"


# --- Пороги -----------------------------------------------------------------

FUZZY_MIN_SCORE = 70.0      # ниже — не показываем даже как кандидата, это шум
FUZZY_TOP_N = 3             # сколько вариантов предлагать человеку
AUTO_APPLY_LEVELS = ("LEARNED", "ARTICLE", "NAME_EXACT")  # что можно ставить без человека


class MatchLevel(str, Enum):
    LEARNED = "LEARNED"
    ARTICLE = "ARTICLE"
    NAME_EXACT = "NAME_EXACT"
    FUZZY = "FUZZY"
    NONE = "NONE"

    @property
    def auto(self) -> bool:
        return self.value in AUTO_APPLY_LEVELS

    @property
    def label(self) -> str:
        return {
            "LEARNED": "по сохранённому соответствию",
            "ARTICLE": "по артикулу",
            "NAME_EXACT": "по наименованию",
            "FUZZY": "похожее — подтвердите",
            "NONE": "новая позиция",
        }[self.value]


# --- Входные структуры ------------------------------------------------------

@dataclass(frozen=True)
class Candidate:
    """Позиция вашего справочника номенклатуры (то, во что сопоставляем)."""
    id: str
    name: str
    unit: str = ""
    article: str | None = None          # ваш внутренний артикул
    supplier_articles: tuple[str, ...] = ()  # артикулы поставщиков из карточки


@dataclass(frozen=True)
class LearnedLink:
    """Строка регистра соответствий: что человек подтвердил в прошлый раз."""
    supplier_id: str
    supplier_key: str        # нормализованная строка поставщика (артикул или имя)
    nomenclature_id: str


@dataclass
class MatchProposal:
    """Один вариант для показа человеку."""
    candidate: Candidate
    score: float


@dataclass
class MatchResult:
    level: MatchLevel
    nomenclature_id: str | None = None
    proposals: list[MatchProposal] = field(default_factory=list)
    unit_mismatch: bool = False
    note: str | None = None

    @property
    def resolved(self) -> bool:
        return self.nomenclature_id is not None and self.level.auto


# --- Нормализация -----------------------------------------------------------

_GOST_RE = re.compile(r"\b(гост|ту|сто|din|iso)[\s\-]*[\d\.\-–/]+", re.IGNORECASE)
_DIAM_RE = re.compile(r"(?:ф|d|ø|диам\.?|диаметр)\s*(\d+(?:[.,]\d+)?)", re.IGNORECASE)
_SIZE_SEP_RE = re.compile(r"(?<=\d)\s*[хx*×]\s*(?=\d)", re.IGNORECASE)
_JUNK_RE = re.compile(r"[«»\"'()\[\],;]")
_SPACE_RE = re.compile(r"\s+")
# «18мм» → «18 мм»: поставщики пишут слитно, в карточке чаще раздельно.
# Без этого совершенно одинаковые позиции уходят в нечёткое сопоставление.
_NUM_UNIT_RE = re.compile(r"(?<=\d)\s*(мм|см|дм|мкм|кг|гр|г|тн|т|мл|л|шт)\b", re.IGNORECASE)


def norm_name(s: str | None) -> str:
    """Каноническая форма наименования для сравнения. Не для показа человеку."""
    if not s:
        return ""
    t = str(s).replace("\xa0", " ").lower()
    t = _GOST_RE.sub(" ", t)
    t = _DIAM_RE.sub(lambda m: "d" + m.group(1).replace(",", "."), t)
    t = _SIZE_SEP_RE.sub("x", t)
    t = _NUM_UNIT_RE.sub(r" \1", t)
    t = _JUNK_RE.sub(" ", t)
    t = t.replace("ё", "е")
    t = _SPACE_RE.sub(" ", t).strip()
    # токенная сортировка: порядок слов у поставщика произвольный
    return " ".join(sorted(t.split()))


def norm_article(s: str | None) -> str:
    """Артикулы сравниваем без разделителей и регистра: 'A500-12' == 'a500 12'."""
    if not s:
        return ""
    return re.sub(r"[^0-9a-zа-я]", "", str(s).lower())


def supplier_key(article: str | None, name: str | None) -> str:
    """Ключ строки поставщика для регистра соответствий.

    Артикул надёжнее имени (поставщик правит формулировки, артикул — нет),
    поэтому при наличии артикула ключ строится по нему.
    """
    a = norm_article(article)
    return f"a:{a}" if a else f"n:{norm_name(name)}"


# --- Индекс кандидатов ------------------------------------------------------

class CandidateIndex:
    """Предрасчитанные индексы, чтобы не нормализовать справочник на каждой строке.

    Строится один раз на загрузку накладной. На 300 строк и 20 000 позиций это
    разница между секундами и минутами.
    """

    def __init__(self, candidates: list[Candidate]):
        self.candidates = candidates
        self.by_article: dict[str, Candidate] = {}
        self.by_name: dict[str, Candidate] = {}
        self.by_id: dict[str, Candidate] = {}
        self._norm_names: list[tuple[str, Candidate]] = []

        for c in candidates:
            self.by_id[c.id] = c
            for a in (c.article, *c.supplier_articles):
                na = norm_article(a)
                if na and na not in self.by_article:
                    self.by_article[na] = c
            nn = norm_name(c.name)
            if nn:
                self._norm_names.append((nn, c))
                self.by_name.setdefault(nn, c)

    def fuzzy(self, query: str, top_n: int = FUZZY_TOP_N,
              min_score: float = FUZZY_MIN_SCORE) -> list[MatchProposal]:
        if not query:
            return []
        scored = [
            MatchProposal(candidate=c, score=_ratio(query, nn))
            for nn, c in self._norm_names
        ]
        scored = [p for p in scored if p.score >= min_score]
        scored.sort(key=lambda p: p.score, reverse=True)
        return scored[:top_n]


# --- Основное правило -------------------------------------------------------

def match_line(
    *,
    name: str,
    article: str | None,
    unit: str,
    index: CandidateIndex,
    learned: dict[str, str] | None = None,
    supplier_id: str | None = None,
) -> MatchResult:
    """Сопоставить одну строку накладной.

    learned — {supplier_key: nomenclature_id} для ЭТОГО поставщика; готовит service.py.
    """
    key = supplier_key(article, name)

    # 1. Выученное соответствие
    if learned:
        nid = learned.get(key)
        if nid and nid in index.by_id:
            cand = index.by_id[nid]
            return MatchResult(
                level=MatchLevel.LEARNED,
                nomenclature_id=nid,
                unit_mismatch=_unit_differs(unit, cand.unit),
            )

    # 2. Артикул
    na = norm_article(article)
    if na and na in index.by_article:
        cand = index.by_article[na]
        return MatchResult(
            level=MatchLevel.ARTICLE,
            nomenclature_id=cand.id,
            unit_mismatch=_unit_differs(unit, cand.unit),
        )

    # 3. Точное наименование
    nn = norm_name(name)
    if nn and nn in index.by_name:
        cand = index.by_name[nn]
        return MatchResult(
            level=MatchLevel.NAME_EXACT,
            nomenclature_id=cand.id,
            unit_mismatch=_unit_differs(unit, cand.unit),
        )

    # 4. Нечёткое — только предложения, без автоприменения
    proposals = index.fuzzy(nn)
    if proposals:
        return MatchResult(
            level=MatchLevel.FUZZY,
            nomenclature_id=None,
            proposals=proposals,
            note="Похожие позиции найдены — требуется подтверждение.",
        )

    # 5. Не нашли
    return MatchResult(level=MatchLevel.NONE, note="Позиция отсутствует в справочнике.")


def _unit_differs(invoice_unit: str, cand_unit: str) -> bool:
    """Единицы разошлись — не блокируем, но помечаем: возможна ошибка в количестве."""
    if not invoice_unit or not cand_unit:
        return False
    return invoice_unit.strip().lower() != cand_unit.strip().lower()


# --- Пакетная обработка -----------------------------------------------------

@dataclass
class MatchStats:
    total: int = 0
    auto: int = 0
    fuzzy: int = 0
    none: int = 0
    unit_warnings: int = 0

    @property
    def auto_share(self) -> float:
        """Доля автосопоставления — главная метрика качества справочника.

        Держать на виду рядом с бизнес-цифрами: если она проседает, отчёты по ценам
        начинают врать раньше, чем это заметят.
        """
        return (self.auto / self.total * 100.0) if self.total else 0.0


def match_lines(lines, index: CandidateIndex,
                learned: dict[str, str] | None = None,
                supplier_id: str | None = None) -> tuple[list[MatchResult], MatchStats]:
    """lines — итерируемое с атрибутами .name/.article/.unit (InvoiceLine из parser.py)."""
    results, st = [], MatchStats()
    for ln in lines:
        r = match_line(
            name=ln.name, article=ln.article, unit=ln.unit,
            index=index, learned=learned, supplier_id=supplier_id,
        )
        st.total += 1
        if r.level.auto:
            st.auto += 1
        elif r.level is MatchLevel.FUZZY:
            st.fuzzy += 1
        else:
            st.none += 1
        if r.unit_mismatch:
            st.unit_warnings += 1
        results.append(r)
    return results, st


def learn(article: str | None, name: str, nomenclature_id: str,
          supplier_id: str) -> LearnedLink:
    """Собрать запись регистра после подтверждения человеком.

    Вызывать ТОЛЬКО после явного подтверждения. Записывать автоматически то, что
    алгоритм сам угадал нечётко, — значит закрепить ошибку навсегда.
    """
    return LearnedLink(
        supplier_id=supplier_id,
        supplier_key=supplier_key(article, name),
        nomenclature_id=nomenclature_id,
    )
