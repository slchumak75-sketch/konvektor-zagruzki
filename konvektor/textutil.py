"""Разбор и нормализация значений из ячеек Excel.

Всё, что здесь есть, ― чистые функции без побочных эффектов: их удобно
тестировать (см. tests/test_textutil.py) и переиспользовать.
"""

from __future__ import annotations

import datetime as _dt
import re
from typing import Any, Iterable, Optional

_SPACES = re.compile(r"[\s\u00a0\u202f\u2009\u2007\u200b]+")
_ONLY_DIGITS = re.compile(r"^\d{4,14}$")
_SCI = re.compile(r"^[+-]?\d+(?:[.,]\d+)?[eE][+-]?\d+$")
_BC = re.compile(r"(?<![\d.,])(\d{8}|\d{12}|\d{13}|\d{14})(?!\d)(?![.,]\d)")

_ART_PATTERNS = (
    re.compile(
        r"\bарт(?:икул)?\s*\.?\s*(?:№|N)?\s*\.?\s*[:№]?\s*([A-Za-zА-Яа-яЁё0-9][A-Za-z0-9\-/_.]{0,29})",
        re.IGNORECASE,
    ),
    re.compile(
        r"\b(?:код|code)\s*(?:товара|артикул)?\s*[:№]?\s*([A-Za-zА-Яа-яЁё0-9][A-Za-z0-9\-/_.]{0,29})",
        re.IGNORECASE,
    ),
)
#: Слова, которые не могут быть артикулом.
_ART_STOP = {
    "товара", "товар", "продукции", "продукция", "тнвэд", "окп", "унп", "штрих",
    "штрихкод", "штрих-код", "по", "и", "или", "для", "без", "с", "в",
}
#: «штрих-код» / «штрих код» — это штрихкод, а не артикул.
_STRICH = re.compile(r"штрих[\s-]?\s*код", re.IGNORECASE)
_LEADING_CODE = re.compile(r"^\s*(\d{4,10})\s+(?=[A-Za-zА-Яа-яЁё«\"])")

#: Значения, которые означают «пусто» в бухгалтерских выгрузках.
_EMPTY_TOKENS = {
    "",
    "-",
    "--",
    "---",
    "—",
    "–",
    "−",
    "x",
    "х",
    "н/д",
    "нд",
    "нет",
    "б/н",
    "без ндс",
    "х",
}

#: Мусорные слова, которые вырезаем из денежных строк.
_CURRENCY_WORDS = (
    "руб.",
    "руб",
    "р.",
    "р",
    "byn",
    "byr",
    "рублей",
    "рубля",
    "рубль",
    "коп.",
    "коп",
    "копеек",
    "копейки",
    "копейка",
    "бел.",
)


_CURRENCY_WORDS_SORTED = tuple(sorted(_CURRENCY_WORDS, key=len, reverse=True))


def to_text(value: Any) -> str:
    """Привести значение ячейки к строке (без 'None')."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "да" if value else "нет"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, _dt.datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, _dt.date):
        return value.isoformat()
    return str(value)


def clean_text(value: Any) -> str:
    """Схлопнуть пробелы/переводы строк, обрезать края."""
    return _SPACES.sub(" ", to_text(value)).strip()


def norm_header(value: Any) -> str:
    """Нормализованный текст заголовка колонки (для словарей синонимов)."""
    t = clean_text(value).lower().replace("ё", "е")
    return t.strip(" .")


def is_blank(value: Any) -> bool:
    return clean_text(value) == ""


def looks_empty_token(value: Any) -> bool:
    t = clean_text(value).lower().strip(" .")
    return t in _EMPTY_TOKENS


def parse_number(value: Any) -> Optional[float]:
    """Разобрать число из ячейки: '1 006,75', '          68,80', '20 %', 6.0.

    Возвращает None, если значение не является числом ('х', 'б/н', пусто...).
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        v = float(value)
        if v != v or v in (float("inf"), float("-inf")):
            return None
        return v
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return None

    s = to_text(value).replace("\u00a0", " ").replace("\u202f", " ")
    s = s.strip()
    if not s:
        return None
    low = s.lower()
    if low.strip(" .") in _EMPTY_TOKENS:
        return None

    # убираем только «денежный» мусор; любые другие буквы -> это не число
    t = low
    for w in _CURRENCY_WORDS_SORTED:
        t = t.replace(w, " ")
    t = t.replace("%", " ").replace("\u202f", " ").replace("\u00a0", " ")
    t = _SPACES.sub(" ", t).strip()
    if not t:
        return None
    if not re.fullmatch(r"[+\-\u2212]?\d[\d ]*(?:[.,]\d+)?", t):
        return None

    t = t.replace(" ", "").replace("\u2212", "-")
    if "," in t and "." in t:
        # 1.006,75 -> 1006.75 ; 1,006.75 -> 1006.75
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    else:
        t = t.replace(",", ".")
    try:
        v = float(t)
    except ValueError:
        return None
    if v != v or v in (float("inf"), float("-inf")):
        return None
    return v


def parse_rate(value: Any) -> Optional[float]:
    """Ставка НДС: '20%', '20 %', 20, 'без НДС' -> 0.0."""
    t = clean_text(value).lower()
    if not t:
        return None
    if "без ндс" in t or t in ("0%", "0"):
        num = parse_number(t)
        return 0.0 if num is None else num
    return parse_number(value)


def ean_ok(code: str) -> bool:
    """Проверка контрольной цифры GTIN/EAN (8, 12, 13, 14 цифр)."""
    if not code or len(code) not in (8, 12, 13, 14) or not code.isdigit():
        return False
    body = [int(c) for c in code[:-1]]
    chk = int(code[-1])
    total = 0
    for i, d in enumerate(reversed(body)):
        total += d * (3 if i % 2 == 0 else 1)
    return (10 - total % 10) % 10 == chk


def parse_barcode(value: Any) -> Optional[str]:
    """Достать штрихкод из значения колонки (число, текст, '4,6E+12'…)."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        v = float(value)
        if v <= 0 or v != v:
            return None
        if v.is_integer():
            s = str(int(v))
            return s if len(s) >= 6 else None
        return None
    if isinstance(value, _dt.datetime):
        return None
    s = clean_text(value)
    if not s:
        return None
    if _ONLY_DIGITS.match(s):
        return s
    if _SCI.match(s.replace(" ", "")):
        try:
            v = float(s.replace(" ", "").replace(",", "."))
        except ValueError:
            return None
        return str(int(v)) if v.is_integer() and v > 0 else None
    return find_barcode_in_text(s)


def find_barcode_in_text(text: Any) -> Optional[str]:
    """Найти штрихкод внутри текста наименования.

    Берём последовательности из 8/12/13/14 цифр, отдавая предпочтение тем,
    у которых сходится контрольная цифра EAN (это отсекает номера ТУ,
    телефоны и прочие длинные числа).
    """
    if text is None:
        return None
    s = to_text(text)
    if not s:
        return None
    # Последовательность после «арт.» — артикул, даже если она похожа на EAN.
    cands = [m.group(1) for m in _BC.finditer(s)
             if not re.search(r"(?:арт(?:икул)?\.?|article)\s*[:№#.-]?\s*$", s[:m.start(1)], re.IGNORECASE)
             and not re.match(r"\s*[-–—:]\s*артикул\b", s[m.end(1):], re.IGNORECASE)]
    cands = [c for c in cands if not c.strip("0") == ""]
    if not cands:
        return None
    valid = [c for c in cands if ean_ok(c)]
    pool = valid or cands
    return max(pool, key=lambda c: (len(c) == 13, len(c), c))


def _mask(text: str, *tokens: Optional[str]) -> str:
    out = text
    for t in tokens:
        if t:
            out = out.replace(t, " ")
    return out


def find_article_in_text(text: Any, barcode: Optional[str] = None) -> Optional[str]:
    """Найти артикул внутри наименования ('арт. 08135', 'код: N451', '8160612 Вода…')."""
    if text is None:
        return None
    s = clean_text(text)
    s = _STRICH.sub(" ", s)
    s = _mask(s, barcode)
    if not s:
        return None
    for pat in _ART_PATTERNS:
        for m in pat.finditer(s):
            cand = m.group(1).strip(" .,;:/-")
            if len(cand) < 2:
                continue
            if cand.lower() in _ART_STOP:
                continue
            if not any(ch.isdigit() for ch in cand):
                continue  # артикул почти всегда содержит цифры
            if cand.isdigit() and len(cand) >= 8:
                # это скорее ещё один штрихкод/длинный код
                continue
            return cand
    m = _LEADING_CODE.match(s)
    if m:
        cand = m.group(1)
        if cand != barcode:
            return cand
    return None


def fmt_money(value: Optional[float]) -> str:
    if value is None:
        return ""
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def fmt_num(value: Optional[float]) -> str:
    if value is None:
        return ""
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.4f}".rstrip("0").rstrip(".").replace(".", ",")


def first_non_empty(values: Iterable[Any]) -> Any:
    for v in values:
        if not is_blank(v):
            return v
    return None
