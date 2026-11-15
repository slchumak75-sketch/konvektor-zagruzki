"""Извлечение позиций товара из листа.

Лист разбивается на «блоки» — участки под каждой найденной шапкой. Это
позволяет корректно читать документы, где таблица повторяется на каждой
странице (ТТН-1, приложения к ТТН, многостраничные счета).

Всё, что не удалось однозначно отнести к позиции, попадает в warnings, чтобы
пользователь видел, а не терял данные молча.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set

from . import detect, layout
from .detect import HeaderInfo
from .readers import Sheet, SourceFile
from .textutil import (
    clean_text,
    find_article_in_text,
    find_barcode_in_text,
    is_blank,
    parse_barcode,
    parse_number,
    parse_rate,
)

MAX_UNIT_LEN = 16
MAX_NAME_LEN = 1000


@dataclass
class Item:
    """Одна позиция товара в нормализованном виде."""

    source: str
    sheet: str
    row: int
    block: int = 0
    name: str = ""
    barcode: Optional[str] = None
    article: Optional[str] = None
    unit: Optional[str] = None
    qty: Optional[float] = None
    price: Optional[float] = None
    total_sum: Optional[float] = None  # «сумма» (без НДС)
    vat_rate: Optional[float] = None
    vat_sum: Optional[float] = None
    sum_with_vat: Optional[float] = None
    computed: Set[str] = field(default_factory=set)

    def as_dict(self) -> Dict[str, Any]:
        return {
            "source": self.source,
            "sheet": self.sheet,
            "row": self.row,
            "block": self.block,
            "barcode": self.barcode,
            "article": self.article,
            "name": self.name,
            "unit": self.unit,
            "qty": self.qty,
            "price": self.price,
            "sum": self.total_sum,
            "vat_rate": self.vat_rate,
            "vat_sum": self.vat_sum,
            "sum_with_vat": self.sum_with_vat,
            "computed": sorted(self.computed),
        }


@dataclass
class DocTotal:
    """Итоговая строка, найденная в документе (для сверки сумм)."""

    source: str
    sheet: str
    row: int
    scope: str = "document"  # block | document
    block: int = 0
    label: str = ""
    values: Dict[str, float] = field(default_factory=dict)  # по распознанным колонкам
    numbers: List[float] = field(default_factory=list)     # все числа строки


@dataclass
class SheetExtraction:
    items: List[Item] = field(default_factory=list)
    totals: List[DocTotal] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    headers: int = 0
    #: решения анализа раскладки (для --explain и листа «Проверки»)
    decisions: List[layout.AnalysisDecision] = field(default_factory=list)
    #: разобранные блоки таблиц (шапка, роли, качество)
    blocks: List[layout.BlockInfo] = field(default_factory=list)


def _col_value(values: List[Any], col: Optional[int]) -> Any:
    if col is None or col < 0 or col >= len(values):
        return None
    return values[col]


def _num(values: List[Any], col: Optional[int]) -> Optional[float]:
    return parse_number(_col_value(values, col))


def _cell_text(values: List[Any], col: Optional[int]) -> str:
    return clean_text(_col_value(values, col))


def _build_item(
    source: str,
    sheet: str,
    row_no: int,
    values: List[Any],
    h: HeaderInfo,
    block: int = 0,
) -> Optional[Item]:
    name = _cell_text(values, h.col(detect.NAME))
    if not name or len(name) > MAX_NAME_LEN:
        return None

    barcode = parse_barcode(_col_value(values, h.col(detect.BARCODE)))
    if not barcode:
        barcode = find_barcode_in_text(name)

    article = _cell_text(values, h.col(detect.ARTICLE))
    if len(article) > 64:
        article = article[:64]
    if not article:
        article = find_article_in_text(name, barcode=barcode) or ""

    unit = _cell_text(values, h.col(detect.UNIT))
    if len(unit) > MAX_UNIT_LEN:
        unit = unit[:MAX_UNIT_LEN].strip()

    item = Item(
        source=source,
        sheet=sheet,
        row=row_no,
        block=block,
        name=name,
        barcode=barcode or None,
        article=article or None,
        unit=unit or None,
        qty=_num(values, h.col(detect.QTY)),
        price=_num(values, h.col(detect.PRICE)),
        total_sum=_num(values, h.col(detect.SUM)),
        vat_rate=parse_rate(_col_value(values, h.col(detect.VAT_RATE))),
        vat_sum=_num(values, h.col(detect.VAT_SUM)),
        sum_with_vat=_num(values, h.col(detect.TOTAL)),
    )

    has_money = any(
        v is not None for v in (item.qty, item.price, item.total_sum, item.vat_sum, item.sum_with_vat)
    )
    if not has_money:
        return None
    return item


_WORDS_CELLS = {"руб.", "руб", "коп.", "коп", "(цифрами)", "(прописью)", "цифрами", "прописью"}


def row_looks_like_words(values: List[Any]) -> bool:
    """Строка с суммой прописью: числа разбиты на «руб.» и «коп.» по клеткам."""
    for v in values:
        if not isinstance(v, str):
            continue
        t = clean_text(v).lower()
        if not t:
            continue
        if t in _WORDS_CELLS or "прописью" in t or "цифрами" in t:
            return True
    return False


def _all_numbers(values: List[Any]) -> List[float]:
    out: List[float] = []
    if row_looks_like_words(values):
        return out
    for v in values:
        n = parse_number(v)
        if n is not None:
            out.append(n)
    return out


def _total_values(values: List[Any], h: HeaderInfo) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for kind in (detect.QTY, detect.SUM, detect.VAT_SUM, detect.TOTAL):
        col = h.col(kind)
        if col is None:
            continue
        v = parse_number(_col_value(values, col))
        if v is not None:
            out[kind] = v
    return out


def extract_sheet(sf: SourceFile, sheet: Sheet) -> SheetExtraction:
    """Разобрать один лист файла."""
    res = SheetExtraction()
    grid = sheet.rows
    if not grid:
        return res

    headers = layout.find_headers(grid)
    transposed = False
    if not headers:
        # шапки в строках нет — возможно, таблица «на бок» (поля в первом столбце)
        block = layout.find_transposed_block(grid)
        if block is not None:
            _start, tgrid, decision = block
            theaders = layout.find_headers(tgrid)
            if theaders:
                grid = tgrid
                headers = theaders
                transposed = True
                res.decisions.append(decision)

    res.headers = len(headers)
    if not headers:
        res.warnings.append(f"{sf.path} [{sheet.name}]: шапка таблицы не найдена")
        return res

    for idx, (hrow, h) in enumerate(headers):
        end = headers[idx + 1][0] if idx + 1 < len(headers) else len(grid)
        data_start = hrow + max(1, h.span)

        # --- проверка гипотез по данным, а не только по подписям -----------
        prof = layout.profiles(grid, data_start, end)
        decision = layout.infer_barcode_column(h, prof)
        if decision is not None:
            res.decisions.append(decision)
        decision = layout.repair_numeric_roles(h, grid, data_start, end)
        if decision is not None:
            res.decisions.append(decision)

        res.blocks.append(
            layout.BlockInfo(
                sheet=sheet.name,
                row=hrow + 1,
                span=max(1, h.span),
                roles=dict(h.cols),
                headers=layout.header_texts(h),
                quality=dict(h.quality),
                arithmetic=layout.arithmetic_ratio(grid, h, data_start, end),
                transposed=transposed,
            )
        )

        prev_item: Optional[Item] = None
        for r in range(data_start, end):
            values = grid[r]
            # в развёрнутой таблице строка = столбец исходного листа
            row_no = (r + 1) if not transposed else r + 1
            if detect.row_is_empty(values):
                continue
            if detect.is_numbering_row(values):
                continue

            label = detect.total_label(values)
            if label is not None:
                first_text = ""
                for v in values:
                    if not is_blank(v):
                        first_text = clean_text(v)
                        break
                res.totals.append(
                    DocTotal(
                        source=sf.path,
                        sheet=sheet.name,
                        row=row_no,
                        block=idx,
                        label=first_text,
                        scope=detect.total_scope(first_text),
                        values=_total_values(values, h),
                        numbers=_all_numbers(values),
                    )
                )
                prev_item = None
                continue

            item = _build_item(sf.path, sheet.name, row_no, values, h, block=idx)
            if item is None:
                # Возможен перенос длинного наименования на следующую строку
                if (
                    prev_item is not None
                    and _only_name_text(values, h)
                ):
                    extra = _cell_text(values, h.col(detect.NAME))
                    if extra and not extra.lower().startswith(("итого", "всего")):
                        prev_item.name = f"{prev_item.name} {extra}".strip()
                        continue
                if not detect.row_is_empty(values):
                    txt = clean_text(_col_value(values, h.col(detect.NAME)))
                    if txt and not detect.is_furniture(txt):
                        res.warnings.append(
                            f"{sf.path} [{sheet.name}] стр.{row_no}: строка пропущена "
                            f"(не найдены числовые значения) — «{txt[:80]}»"
                        )
                prev_item = None
                continue
            res.items.append(item)
            prev_item = item

    return res


def _only_name_text(values: List[Any], h: HeaderInfo) -> bool:
    """В строке есть только текст наименования, все числовые колонки пусты."""
    for kind in (detect.QTY, detect.PRICE, detect.SUM, detect.VAT_RATE, detect.VAT_SUM, detect.TOTAL):
        col = h.col(kind)
        if col is None:
            continue
        if parse_number(_col_value(values, col)) is not None:
            return False
    name_col = h.col(detect.NAME)
    if name_col is None:
        return False
    txt = _cell_text(values, name_col)
    return len(txt) >= 10 and txt[:2].upper() not in ("I.", "II")


def extract_file(sf: SourceFile) -> SheetExtraction:
    """Разобрать все листы файла."""
    total = SheetExtraction()
    for sheet in sf.sheets:
        part = extract_sheet(sf, sheet)
        total.items.extend(part.items)
        total.totals.extend(part.totals)
        total.warnings.extend(part.warnings)
        total.decisions.extend(part.decisions)
        total.blocks.extend(part.blocks)
        total.headers += part.headers
    if not total.items and not total.warnings:
        total.warnings.append(f"{sf.path}: позиции не найдены")
    return total
