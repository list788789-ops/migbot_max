"""
common_utils.py — мелкие функции, используемые и в bot.py (MAX-бот), и в
webforms.py (веб-формы кадровика). Раньше были продублированы в обоих
местах по отдельности — риск в том, что поправят одно место и забудут
про второе (ровно так и обнаружилось: category_for_citizenship в bot.py
проверял английское "belarus", а webforms.py — русское "Беларусь", то
есть уже успели разъехаться до выноса сюда).

Не тянет модели/БД само по себе — чистые функции, чтобы можно было
импортировать и в бота, и в веб-сервис без побочных эффектов при импорте.
"""

from models import Category


def normalize_phone(phone: str) -> str:
    """Телефон-логин к единому виду: только цифры, ведущая 8 -> 7. Чтобы +7/8/пробелы
    не создавали разных логинов одному человеку."""
    digits = "".join(ch for ch in (phone or "") if ch.isdigit())
    if digits.startswith("8") and len(digits) == 11:
        digits = "7" + digits[1:]
    return digits


# Единственное явное сопоставление — Беларусь имеет отдельный срок постановки
# на учёт (90 календарных дней вместо 30 у остального ЕАЭС, см. models.py).
# Всё, чего нет в этом словаре, по умолчанию считается EAEU.
CITIZENSHIP_TO_CATEGORY = {"Беларусь": Category.BELARUS}


def category_for_citizenship(citizenship: str) -> Category:
    return CITIZENSHIP_TO_CATEGORY.get((citizenship or "").strip(), Category.EAEU)
