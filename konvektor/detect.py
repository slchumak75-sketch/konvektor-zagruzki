"""Распознавание шапки таблицы и типов колонок.

Идея: у каждого поставщика свой «образец» счёта, но набор смысловых колонок
один и тот же (штрихкод, артикул, наименование, ед. изм., количество, цена,
сумма, ставка НДС, сумма НДС, сумма с НДС). Поэтому мы не привязываемся к
номерам колонок и к названиям файлов, а сопоставляем текст заголовка со
словарём синонимов. Такой подход переживает появление новых поставщиков.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from .textutil import norm_header, parse_number

#: Смысловые поля итоговой таблицы.
NAME = "name"
BARCODE = "barcode"
ARTICLE = "article"
UNIT = "unit"
QTY = "qty"
PRICE = "price"
SUM = "sum"  # сумма без НДС / стоимость
VAT_RATE = "vat_rate"
VAT_SUM = "vat_sum"
TOTAL = "total"  # сумма с НДС / всего

NUMERIC_KINDS = (QTY, PRICE, SUM, TOTAL)

#: Колонки, которые нам никогда не нужны (и которые иначе ложно опознаются).
_REJECT = (
    "тнвэд",
    "окп",
    "унп",
    "бик",
    "грузовых мест",
    "масса",
    "вес",
    "объем",
    "объём",
    "страна",
    "примечание",
    "сертификат",
    "фасовка",
    "шт. в кр",
    "декларац",
    "iban",
    "р/с",
    "расчетный счет",
)

_NAME_HINTS = ("наимен", "номенклатур", "товар", "описан", "продукц", "материал", "услуг")
_VAT_EXCL_HINTS = ("без ндс", "б/ндс", "б ндс", "безндс")
_VAT_INCL_HINTS = ("с ндс", "сндс")
_BAD_PRICE_HINTS = (
    "отпускн",
    "производител",
    "импортер",
    "закуп",
    "розн",
    "опт",
    "доставк",
    "транспорт",
    "с ндс",
)


@dataclass
class HeaderInfo:
    """Найденная шапка: номер строки и карта «смысл -> индекс колонки»."""

    row: int
    cols: Dict[str, int]
    quality: Dict[str, float] = field(default_factory=dict)
    raw: List[Any] = field(default_factory=list)

    @property
    def score(self) -> float:
        return sum(self.quality.values()) + len(self.cols)

    def col(self, kind: str) -> Optional[int]:
        return self.cols.get(kind)

    def has(self, kind: str) -> bool:
        return kind in self.cols

    @property
    def numeric_count(self) -> int:
        return sum(1 for k in NUMERIC_KINDS if k in self.cols)


def classify_header(value: Any) -> Optional[Tuple[str, float]]:
    """Определить смысл колонки по тексту заголовка.

    Возвращает (kind, quality) или None. quality используется, чтобы выбрать
    лучшую колонку, если подходящих несколько (например «Цена шт. б/НДС» и
    «Цена шт. с НДС»).
    """
    t = norm_header(value)
    if not t:
        return None

    if any(x in t for x in _REJECT):
        return None

    if "штрих" in t or "barcode" in t or "ean" in t or "штрихкод" in t:
        return (BARCODE, 3.0)

    if "артикул" in t or re.search(r"\bарт\b", t):
        return (ARTICLE, 3.0)

    if "ставка" in t:
        return (VAT_RATE, 3.0)

    if "цена" in t or "price" in t:
        q = 2.0
        if any(x in t for x in _VAT_EXCL_HINTS):
            q += 2.0
        if "за единицу" in t or t.startswith("цена"):
            q += 0.5
        if any(x in t for x in _BAD_PRICE_HINTS):
            q -= 3.0
        return (PRICE, q)

    if "ндс" in t:
        if any(x in t for x in _VAT_EXCL_HINTS):
            return (SUM, 3.0)
        if any(x in t for x in _VAT_INCL_HINTS) or "всего" in t or "к оплате" in t:
            return (TOTAL, 3.0)
        return (VAT_SUM, 3.0)

    if "всего" in t or "итого" in t or "к оплате" in t:
        return (TOTAL, 2.0)

    if "кол" in t or "к-во" in t or "количество" in t or "qty" in t or "количест" in t:
        return (QTY, 3.0)

    if any(x in t for x in _NAME_HINTS):
        return (NAME, 3.0)

    if ("ед" in t and "изм" in t) or t in ("ед", "ед.", "единица", "ед.изм", "ед. изм", "ед. измер"):
        return (UNIT, 3.0)

    if "сумм" in t or "стоимост" in t:
        return (SUM, 2.0)

    if "код" in t or "code" in t:
        return (ARTICLE, 1.0)

    return None


def detect_header(values: List[Any], row_index: int = 0) -> Optional[HeaderInfo]:
    """Проверить, является ли строка шапкой таблицы."""
    cols: Dict[str, int] = {}
    quality: Dict[str, float] = {}
    for c, v in enumerate(values):
        res = classify_header(v)
        if not res:
            continue
        kind, q = res
        if kind not in cols or q > quality[kind] + 1e-9:
            cols[kind] = c
            quality[kind] = q

    if NAME not in cols:
        return None
    numeric = sum(1 for k in NUMERIC_KINDS if k in cols)
    if len(cols) < 4 or numeric < 2:
        return None
    return HeaderInfo(row=row_index, cols=cols, quality=quality, raw=list(values))


# --------------------------------------------------------------------------- #
#  Эвристики для строк
# --------------------------------------------------------------------------- #
_TOTAL_PREFIXES = (
    "итого",
    "всего",
    "в том числе",
    "сумма ндс",
    "сумма без ндс",
    "сумма с ндс",
    "всего сумма",
    "к оплате",
    "всего к оплате",
    "итого с ндс",
    "итого по",
)
_PAGE_TOTAL_HINTS = ("по листу", "по странице", "по стр.", "итого по стр")


def total_label(values: List[Any]) -> Optional[str]:
    """Если строка — итоговая, вернуть её текстовую метку."""
    for v in values:
        if v is None:
            continue
        t = norm_header(v)
        if not t:
            continue
        for p in _TOTAL_PREFIXES:
            if t.startswith(p):
                return t
        return None
    return None


_TOTAL_SKIP_HINTS = ("масса", "объем", "объём", "грузовых мест", "прописью", "доставка")
_TOTAL_TOTAL_HINTS = ("с ндс", "к оплате", "итого с ндс")
_TOTAL_VAT_HINTS = ("ндс",)
_TOTAL_SUM_HINTS = ("без ндс", "б/ндс")
_TOTAL_TOTAL_HINTS_EXTRA = ("стоимость",)


def total_field_from_label(label: str) -> str:
    """Какое поле означает итоговая строка.

    Возвращает 'sum' | 'vat_sum' | 'total' | 'qty', '' если непонятно
    (тогда сверяем все распознанные колонки), 'skip' — если строку сверять
    не нужно (масса, доставка, суммы прописью).
    """
    t = norm_header(label)
    if any(x in t for x in _TOTAL_SKIP_HINTS):
        return "skip"
    if any(x in t for x in _TOTAL_SUM_HINTS):
        return "sum"
    if any(x in t for x in _TOTAL_TOTAL_HINTS):
        return "total"
    if any(x in t for x in _TOTAL_VAT_HINTS):
        return "vat_sum"
    if "количест" in t or "кол" in t:
        return "qty"
    return ""


#: Признаки «служебных» строк документа — их пропуск не считается потерей.
_FURNITURE_HINTS = (
    "номер листа",
    "товарный раздел",
    "поставщик",
    "покупатель",
    "всего",
    "итого",
    "подпись",
    "руководитель",
    "бухгалтер",
    "приложение к",
    "договор",
    "адрес",
    "телефон",
    "унп",
    "р/сч",
    "контактное лицо",
    "с товаром переданы",
    "отпуск разрешил",
    "сдал",
    "принял",
    "масса",
    "прописью",
    "страница",
    "переадресовка",
    "водитель",
    "автомобиль",
    "заказчик",
    "грузоотправитель",
    "грузополучатель",
    "основание отпуска",
    "пункт погрузки",
    "пункт разгрузки",
    "прицеп",
    "к путевому листу",
    "расчет стоимости",
    "счет на оплату является",
    "оплата по реализации",
    "банк",
    "бик",
)


def is_furniture(text: Any) -> bool:
    """Служебная строка документа (шапка, подписи, реквизиты), а не товар."""
    t = norm_header(text)
    if not t:
        return True
    return any(h in t for h in _FURNITURE_HINTS)


def total_scope(label: str) -> str:
    """'block' — итог по листу/странице, 'document' — накопительный итог."""
    t = label.lower()
    if any(h in t for h in _PAGE_TOTAL_HINTS):
        return "block"
    return "document"


def is_numbering_row(values: List[Any]) -> bool:
    """Строка-нумерация «1 2 3 4 …» под шапкой таблицы."""
    nums: List[float] = []
    for v in values:
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        n = parse_number(v)
        if n is None or not float(n).is_integer() or not (1 <= n <= 60):
            return False
        nums.append(n)
    if len(nums) < 4:
        return False
    return all(b > a for a, b in zip(nums, nums[1:]))


def row_is_empty(values: List[Any]) -> bool:
    for v in values:
        if v is None:
            continue
        if isinstance(v, str) and not v.strip():
            continue
        return False
    return True
