"""Анализ раскладки: гипотезы о смысле колонок проверяются по данным.

`detect.py` отвечает на вопрос «что написано в шапке». Этого мало: поставщики
путают подписи, оставляют колонки без подписи, кладут таблицу «на бок» и
склеивают шапку из двух строк. Поэтому здесь работает второй контур анализа,
который **не доверяет шапке на слово**:

1. **Профиль колонки** (`column_profile`) — что реально лежит в столбце:
   числа, целые, коды из 8–14 цифр, текст, длины, повторы.
2. **Достройка ролей по данным** (`infer_barcode_column`) — колонка без подписи,
   в которой лежат 13-значные коды, становится штрихкодом.
3. **Проверка гипотез арифметикой** (`repair_numeric_roles`) — если разметка
   «цена/сумма/НДС» не сходится с числами (количество × цена ≠ сумма), роли
   переставляются по столбцам так, чтобы сходилось. Это и есть общий механизм,
   а не правило под конкретный файл.
4. **Ориентация таблицы** (`find_transposed_block`) — если шапки в строках нет,
   ищется блок, где названия полей стоят в первом столбце, и таблица
   разворачивается.
5. **Склейка шапки из двух строк** (`compose_header`) — «Сумма» + «без НДС».

Все решения записываются в `AnalysisDecision`, чтобы их можно было показать
пользователю (`--explain`) и объяснить, почему получилось именно так.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from itertools import permutations
from typing import Any, Dict, List, Optional, Sequence, Tuple

from . import detect
from .detect import HeaderInfo
from .textutil import clean_text, norm_header, parse_number

#: Роли, которые проверяются арифметикой.
MONEY_ROLES = (detect.PRICE, detect.SUM, detect.VAT_SUM, detect.TOTAL)
RATE_ROLE = detect.VAT_RATE
NUMERIC_ROLES = (detect.QTY,) + MONEY_ROLES + (RATE_ROLE,)

_RE_DIGITS = re.compile(r"^\d{8,14}$")
_RE_CODE_WITH_PREFIX = re.compile(r"^(?:штрих[- ]?код[:\s]*)?(\d{8,14})$", re.IGNORECASE)

#: Допуск сверки: как в validate.py — цент и полпроцента.
def close(a: float, b: float) -> bool:
    return abs(a - b) <= max(0.02, 0.005 * max(abs(a), abs(b)))


@dataclass
class BlockInfo:
    """Один распознанный блок таблицы: шапка, роли, качество."""

    sheet: str = ""
    row: int = 0              # строка шапки, 1-based
    span: int = 1             # сколько строк занимает шапка
    roles: Dict[str, int] = field(default_factory=dict)     # роль -> колонка (0-based)
    headers: Dict[int, str] = field(default_factory=dict)   # колонка -> текст шапки
    quality: Dict[str, float] = field(default_factory=dict)
    #: доля сошедшихся проверок (None — проверять было нечего)
    arithmetic: Optional[float] = 0.0
    transposed: bool = False

    @property
    def roles_count(self) -> int:
        return len(self.roles)


@dataclass
class ColumnProfile:
    """Что фактически лежит в столбце под шапкой."""

    col: int
    filled: int = 0
    numeric: int = 0
    integers: int = 0
    codes: int = 0
    texts: int = 0
    small_ints: int = 0   # 1..999 — похоже на количество
    rates: int = 0        # 0..100
    distinct: int = 0
    avg_len: float = 0.0

    @property
    def numeric_ratio(self) -> float:
        return self.numeric / self.filled if self.filled else 0.0

    @property
    def code_ratio(self) -> float:
        return self.codes / self.filled if self.filled else 0.0

    @property
    def text_ratio(self) -> float:
        return self.texts / self.filled if self.filled else 0.0

    @property
    def qty_ratio(self) -> float:
        return self.small_ints / self.filled if self.filled else 0.0

    @property
    def rate_ratio(self) -> float:
        return self.rates / self.filled if self.filled else 0.0


@dataclass
class AnalysisDecision:
    """Одно решение анализа: что сделали и почему."""

    kind: str          # barcode | roles | transpose | compose | header
    message: str
    detail: Dict[str, Any] = field(default_factory=dict)

    def as_text(self) -> str:
        if not self.detail:
            return self.message
        detail = ", ".join(f"{k}={v}" for k, v in self.detail.items())
        return f"{self.message} ({detail})"


def header_texts(header: HeaderInfo) -> Dict[int, str]:
    """Текст шапки по колонкам (для отчёта)."""
    out: Dict[int, str] = {}
    for col in sorted(set(header.cols.values())):
        text = ""
        if col < len(header.raw):
            text = clean_text(header.raw[col])
        out[col] = text
    return out


def arithmetic_ratio(
    grid: Sequence[Sequence[Any]], header: HeaderInfo, start: int, end: int
) -> Optional[float]:
    """Доля сошедшихся проверок в блоке (0..1) или None, если сверять нечего."""
    rows = _data_rows(grid, start, end)
    if not rows:
        return None
    checks = _arith_checks(rows, {role: header.col(role) for role in NUMERIC_ROLES})
    if not any(n for _m, n in checks.values()):
        return None
    return _score(checks)


# --------------------------------------------------------------------------- #
#  Профиль колонки
# --------------------------------------------------------------------------- #
def _data_rows(grid: Sequence[Sequence[Any]], start: int, end: int) -> List[Sequence[Any]]:
    """Строки-данные: непустые, без итоговых подписей и нумерации."""
    rows: List[Sequence[Any]] = []
    for r in range(max(start, 0), min(end, len(grid))):
        values = grid[r]
        if detect.row_is_empty(values):
            continue
        if detect.is_numbering_row(values):
            continue
        if detect.total_label(values) is not None:
            continue
        rows.append(values)
    return rows


def column_profile(
    rows: Sequence[Sequence[Any]], col: int, max_rows: int = 400
) -> ColumnProfile:
    p = ColumnProfile(col=col)
    seen = set()
    lengths: List[int] = []
    for values in rows[:max_rows]:
        value = values[col] if col < len(values) else None
        if value is None or (isinstance(value, str) and not value.strip()):
            continue
        p.filled += 1
        text = clean_text(value)
        lengths.append(len(text))
        seen.add(text)
        num = parse_number(value)
        if num is not None:
            p.numeric += 1
            if float(num).is_integer():
                p.integers += 1
                if 1 <= num <= 999:
                    p.small_ints += 1
            if 0 <= num <= 100:
                p.rates += 1
        digits = text.replace(" ", "")
        m = _RE_DIGITS.match(digits)
        if m or (isinstance(value, (int, float)) and _is_code_number(value)):
            p.codes += 1
        if num is None:
            p.texts += 1
    p.distinct = len(seen)
    p.avg_len = sum(lengths) / len(lengths) if lengths else 0.0
    return p


def _is_code_number(value: Any) -> bool:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    if not float(value).is_integer():
        return False
    return 10**7 <= int(value) < 10**14


def profiles(
    grid: Sequence[Sequence[Any]], start: int, end: int, width: Optional[int] = None
) -> Dict[int, ColumnProfile]:
    rows = _data_rows(grid, start, end)
    width = width if width is not None else max((len(r) for r in grid), default=0)
    return {c: column_profile(rows, c) for c in range(width)}


# --------------------------------------------------------------------------- #
#  Достройка ролей по данным
# --------------------------------------------------------------------------- #
def infer_barcode_column(
    header: HeaderInfo, prof: Dict[int, ColumnProfile]
) -> Optional[AnalysisDecision]:
    """Колонка без подписи, целиком состоящая из 8–14-значных кодов."""
    if header.has(detect.BARCODE):
        return None
    used = set(header.cols.values())
    best: Optional[Tuple[float, int]] = None
    for col, p in prof.items():
        if col in used or p.filled < 2:
            continue
        if p.code_ratio < 0.8:
            continue
        score = (p.code_ratio, -col)
        if best is None or score > (best[0], -best[1]):
            best = (p.code_ratio, col)
    if best is None:
        return None
    col = best[1]
    header.cols[detect.BARCODE] = col
    header.quality[detect.BARCODE] = 2.0
    return AnalysisDecision(
        kind="barcode",
        message="колонка штрихкода без подписи найдена по данным",
        detail={"колонка": col + 1, "коды": f"{best[0]:.0%}"},
    )


def _arith_checks(
    rows: Sequence[Sequence[Any]], cols: Dict[str, Optional[int]]
) -> Dict[str, Tuple[int, int]]:
    """Сколько раз сходятся равенства при такой разметке ролей.

    Возвращает словарь «проверка -> (совпало, проверено)». Для НДС
    рассматриваются оба режима (сверху и «в том числе»), чтобы не путать
    разметку с режимом ценообразования.
    """
    checks = {"qty_price": [0, 0], "sum_vat": [0, 0], "vat": [0, 0]}

    def val(values: Sequence[Any], role: str) -> Optional[float]:
        col = cols.get(role)
        if col is None or col >= len(values):
            return None
        return parse_number(values[col])

    for values in rows:
        qty = val(values, detect.QTY)
        price = val(values, detect.PRICE)
        total_sum = val(values, detect.SUM)
        vat_sum = val(values, detect.VAT_SUM)
        total = val(values, detect.TOTAL)
        rate = val(values, RATE_ROLE)

        if qty is not None and price is not None and total_sum is not None:
            checks["qty_price"][1] += 1
            checks["qty_price"][0] += 1 if close(qty * price, total_sum) else 0
        if total_sum is not None and vat_sum is not None and total is not None:
            checks["sum_vat"][1] += 1
            ok = close(total_sum + vat_sum, total) or close(total_sum, total)
            checks["sum_vat"][0] += 1 if ok else 0
        if total_sum is not None and vat_sum is not None and rate is not None:
            checks["vat"][1] += 1
            excl = close(total_sum * rate / 100.0, vat_sum)
            incl = close(total_sum * rate / (100.0 + rate), vat_sum) if rate > -100 else False
            checks["vat"][0] += 1 if (excl or incl) else 0
    return {k: (v[0], v[1]) for k, v in checks.items()}


def _score(checks: Dict[str, Tuple[int, int]]) -> float:
    """Взвешенная доля сошедшихся проверок (0..1)."""
    weights = {"qty_price": 2.0, "sum_vat": 1.0, "vat": 1.0}
    total = 0.0
    got = 0.0
    for name, weight in weights.items():
        matched, checked = checks.get(name, (0, 0))
        if not checked:
            continue
        total += weight
        got += weight * matched / checked
    return got / total if total else 0.0


def repair_numeric_roles(
    header: HeaderInfo,
    grid: Sequence[Sequence[Any]],
    start: int,
    end: int,
    min_improvement: float = 0.25,
) -> Optional[AnalysisDecision]:
    """Переставить роли «цена/сумма/НДС/итог» так, чтобы числа сходились.

    Гипотезы проверяются по данным: количество × цена = сумма,
    сумма + НДС = всего, сумма × ставка = НДС. Если текущая разметка даёт
    меньше совпадений, ищется другая расстановка тех же столбцов.
    """
    rows = _data_rows(grid, start, end)
    if len(rows) < 2:
        return None
    current = {role: header.col(role) for role in NUMERIC_ROLES}
    checks = _arith_checks(rows, current)
    if not any(n for _m, n in checks.values()):
        return None
    base = _score(checks)
    if base >= 0.8:
        return None

    cols = sorted({c for c in current.values() if c is not None})
    if len(cols) < 3:
        return None

    best: Optional[Tuple[float, Dict[str, int]]] = None
    roles = [r for r in NUMERIC_ROLES if current.get(r) is not None]
    if len(roles) > 5 or len(cols) > 6:
        return None
    for perm in permutations(cols, len(roles)):
        assignment = dict(zip(roles, perm))
        got = _arith_checks(rows, assignment)
        if not any(n for _m, n in got.values()):
            continue
        score = _score(got)
        if best is None or score > best[0] + 1e-9:
            best = (score, assignment)
    if best is None or best[0] < base + min_improvement:
        return None

    changed = {r: c for r, c in best[1].items() if current.get(r) != c}
    if not changed:
        return None
    for role, col in best[1].items():
        header.cols[role] = col
    return AnalysisDecision(
        kind="roles",
        message="роли числовых колонок исправлены по арифметике",
        detail={
            "было": {k: (v + 1 if v is not None else None) for k, v in current.items()},
            "стало": {k: v + 1 for k, v in changed.items()},
            "совпадений": f"{best[0]:.0%}",
        },
    )


# --------------------------------------------------------------------------- #
#  Шапка из двух строк
# --------------------------------------------------------------------------- #
def compose_header(row_top: Sequence[Any], row_bottom: Sequence[Any], index: int) -> List[str]:
    """Склеить двухуровневую шапку: «Сумма» + «без НДС» -> «Сумма без НДС»."""
    width = max(len(row_top), len(row_bottom))
    out: List[str] = []
    for c in range(width):
        top = clean_text(row_top[c]) if c < len(row_top) else ""
        bottom = clean_text(row_bottom[c]) if c < len(row_bottom) else ""
        if top and bottom and norm_header(bottom) != norm_header(top):
            out.append(f"{top} {bottom}")
        else:
            out.append(top or bottom)
    return out


def _filled(values: Sequence[Any]) -> int:
    return sum(
        1 for v in values if v is not None and not (isinstance(v, str) and not v.strip())
    )


def _is_label_row(values: Sequence[Any]) -> bool:
    """Строка целиком из подписей (чисел нет) — годна как вторая строка шапки."""
    text = 0
    for v in values:
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        if parse_number(v) is not None:
            return False
        text += 1
    return text >= 2


def find_headers(
    grid: Sequence[Sequence[Any]], max_rows: Optional[int] = None
) -> List[Tuple[int, HeaderInfo]]:
    """Все шапки листа: одиночные строки и склейки из двух строк.

    Склейка рассматривается, только если вторая строка — подписи (без чисел),
    и даёт больше смысловых колонок, чем любая из строк по отдельности.
    """
    limit = len(grid) if max_rows is None else min(len(grid), max_rows)
    singles: Dict[int, HeaderInfo] = {}
    for i in range(limit):
        h = detect.detect_header(grid[i], i)
        if h is not None:
            singles[i] = h

    composed: Dict[int, Tuple[HeaderInfo, int]] = {}
    for i in range(limit - 1):
        bottom = grid[i + 1]
        top = grid[i]
        if not _is_label_row(bottom):
            continue
        # верхняя строка должна быть «шапкоподобной»: несколько подписей, а не
        # одинокий заголовок раздела («I. ТОВАРНЫЙ РАЗДЕЛ»)
        if not _is_label_row(top) or _filled(top) < 3:
            continue
        merged = compose_header(top, bottom, i)
        h = detect.detect_header(merged, i)
        if h is None:
            continue
        # склейка должна давать больше смысла, чем любая строка по отдельности
        single_roles = max(
            len(singles[i].cols) if i in singles else 0,
            len(singles[i + 1].cols) if i + 1 in singles else 0,
        )
        if len(h.cols) <= single_roles:
            continue
        if h.numeric_count < 2:
            continue
        h.raw = merged
        h.span = 2
        composed[i] = (h, 2)

    headers: List[Tuple[int, HeaderInfo]] = []
    i = 0
    while i < limit:
        if i in composed:
            h, span = composed[i]
            headers.append((i, h))
            i += span
            continue
        if i in singles:
            headers.append((i, singles[i]))
        i += 1
    return headers


# --------------------------------------------------------------------------- #
#  Ориентация: таблица «на бок»
# --------------------------------------------------------------------------- #
TRANSPOSE = "transpose"


def transpose(rows: Sequence[Sequence[Any]]) -> List[List[Any]]:
    width = max((len(r) for r in rows), default=0)
    out: List[List[Any]] = []
    for c in range(width):
        out.append([rows[r][c] if c < len(rows[r]) else None for r in range(len(rows))])
    return out


def find_transposed_block(
    grid: Sequence[Sequence[Any]],
) -> Optional[Tuple[int, List[List[Any]], AnalysisDecision]]:
    """Найти блок, где названия полей стоят в первом столбце, и развернуть его."""
    rows: List[int] = []
    for i, values in enumerate(grid):
        if detect.row_is_empty(values):
            continue
        filled = [v for v in values if v is not None and not (isinstance(v, str) and not v.strip())]
        if len(filled) < 3:
            continue
        first = values[0] if values else None
        role = detect.classify_header(first)
        if role is None:
            # «Стоимость с НДС» и подобное распознаётся как роль; иначе — мимо
            if detect.classify_header(clean_text(first).replace(":", "")) is None:
                continue
        rows.append(i)

    if len(rows) < 4:
        return None
    start, stop = rows[0], rows[-1] + 1
    if any(i not in rows for i in range(start, stop)) and len(rows) < 5:
        return None

    block = [list(grid[i]) for i in range(start, stop) if i in rows]
    transposed = transpose(block)
    if not transposed:
        return None
    decision = AnalysisDecision(
        kind=TRANSPOSE,
        message="таблица развёрнута: названия полей стоят в первом столбце",
        detail={"строки": f"{start + 1}–{stop}", "полей": len(transposed)},
    )
    return start, transposed, decision
