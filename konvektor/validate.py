"""Математическая проверка и «достройка» строк.

Правила, которые проверяются:

    сумма          ≈ количество × цена
    сумма с НДС    ≈ сумма + сумма НДС
    сумма НДС      ≈ сумма × ставка / 100        (цены без НДС)
    сумма НДС      ≈ сумма × ставка / (100+ставка) (цены с НДС)

Режим ценообразования (с НДС или без) определяется по самому документу: какой
вариант даёт меньше расхождений, тот и считается верным.

Если значение в документе отсутствует, но его можно однозначно посчитать, —
оно рассчитывается и помечается как вычисленное (в Excel заливается жёлтым).
Каждое расхождение попадает в отчёт.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .detect import total_field_from_label
from .extract import DocTotal, Item
from .textutil import fmt_money

#: Допуск по деньгам, руб. (два округления до копеек + округление НДС).
MONEY_TOL = 0.02
#: Допуск по количеству.
QTY_TOL = 0.002
EXCLUSIVE = "exclusive"
INCLUSIVE = "inclusive"

MODE_TITLES = {
    EXCLUSIVE: "цены без НДС (НДС сверху)",
    INCLUSIVE: "цены с НДС (НДС в том числе)",
}


@dataclass
class Issue:
    """Одно замечание проверки."""

    severity: str  # error | warning | info
    source: str
    sheet: str
    row: int
    name: str
    field: str
    message: str
    doc_value: Optional[float] = None
    calc_value: Optional[float] = None

    @property
    def delta(self) -> Optional[float]:
        if self.doc_value is None or self.calc_value is None:
            return None
        return round(self.calc_value - self.doc_value, 4)


@dataclass
class ValidationReport:
    items: List[Item]
    issues: List[Issue] = field(default_factory=list)
    vat_mode: Dict[str, str] = field(default_factory=dict)  # файл -> режим

    @property
    def errors(self) -> List[Issue]:
        return [i for i in self.issues if i.severity == "error"]

    @property
    def warnings(self) -> List[Issue]:
        return [i for i in self.issues if i.severity == "warning"]

    @property
    def infos(self) -> List[Issue]:
        return [i for i in self.issues if i.severity == "info"]

    @property
    def computed_cells(self) -> int:
        return sum(len(i.computed) for i in self.items)


def _round2(v: float) -> float:
    return round(v + 0.0, 2)


def expected_vat(total_sum_item: float, rate: Optional[float], mode: str) -> Optional[float]:
    """НДС по сумме без НДС и ставке."""
    if rate is None:
        return None
    if mode == INCLUSIVE:
        return _round2(total_sum_item * rate / (100.0 + rate)) if rate else 0.0
    return _round2(total_sum_item * rate / 100.0)


def detect_vat_mode(items: Sequence[Item]) -> str:
    """Определить, включён ли НДС в «сумму», по данным документа."""
    excl_score = 0.0
    incl_score = 0.0
    used = 0
    for it in items:
        if it.total_sum is None or it.vat_sum is None or not it.vat_rate:
            continue
        used += 1
        excl_score += abs(it.vat_sum - it.total_sum * it.vat_rate / 100.0)
        incl_score += abs(it.vat_sum - it.total_sum * it.vat_rate / (100.0 + it.vat_rate))
    if used == 0:
        return EXCLUSIVE
    return INCLUSIVE if incl_score < excl_score else EXCLUSIVE


def _fill(item: Item, mode: str, issues: List[Issue]) -> None:
    """Достроить отсутствующие значения строки (с пометкой «вычислено»)."""

    def note(field_name: str, value: float) -> None:
        issues.append(
            Issue(
                severity="info",
                source=item.source,
                sheet=item.sheet,
                row=item.row,
                name=item.name,
                field=field_name,
                message=f"{field_name} рассчитана программой: {fmt_money(value)}",
                calc_value=value,
            )
        )
        item.computed.add(field_name)

    rate = item.vat_rate

    # 1. сумма без НДС
    if item.total_sum is None:
        if item.qty is not None and item.price is not None:
            item.total_sum = _round2(item.qty * item.price)
            note("сумма", item.total_sum)
        elif mode == INCLUSIVE and item.sum_with_vat is not None:
            item.total_sum = _round2(item.sum_with_vat)
            note("сумма", item.total_sum)
        elif item.sum_with_vat is not None and item.vat_sum is not None:
            item.total_sum = _round2(item.sum_with_vat - item.vat_sum)
            note("сумма", item.total_sum)
        elif item.sum_with_vat is not None and rate:
            item.total_sum = _round2(item.sum_with_vat / (1 + rate / 100.0))
            note("сумма", item.total_sum)

    # 2. сумма НДС
    if item.vat_sum is None:
        if mode == EXCLUSIVE and item.total_sum is not None and item.sum_with_vat is not None:
            item.vat_sum = _round2(item.sum_with_vat - item.total_sum)
            note("сумма НДС", item.vat_sum)
        elif item.total_sum is not None and rate is not None:
            v = expected_vat(item.total_sum, rate, mode)
            if v is not None:
                item.vat_sum = v
                note("сумма НДС", v)
        elif item.total_sum is not None and item.sum_with_vat is not None:
            item.vat_sum = _round2(item.sum_with_vat - item.total_sum)
            note("сумма НДС", item.vat_sum)

    # 3. сумма с НДС
    if item.sum_with_vat is None:
        if mode == INCLUSIVE and item.total_sum is not None:
            item.sum_with_vat = _round2(item.total_sum)
            note("сумма с НДС", item.sum_with_vat)
        elif item.total_sum is not None and item.vat_sum is not None:
            item.sum_with_vat = _round2(item.total_sum + item.vat_sum)
            note("сумма с НДС", item.sum_with_vat)
        elif item.total_sum is not None and rate is not None:
            item.sum_with_vat = _round2(item.total_sum + (expected_vat(item.total_sum, rate, mode) or 0.0))
            note("сумма с НДС", item.sum_with_vat)
        elif item.qty is not None and item.price is not None:
            base = _round2(item.qty * item.price)
            if rate:
                if mode == INCLUSIVE:
                    item.sum_with_vat = base
                else:
                    item.sum_with_vat = _round2(base * (1 + rate / 100.0))
            else:
                item.sum_with_vat = base
            note("сумма с НДС", item.sum_with_vat)

    # 4. цена
    if item.price is None and item.total_sum is not None and item.qty:
        item.price = round(item.total_sum / item.qty, 4)
        note("цена", item.price)

    # 5. количество
    if item.qty is None and item.price and item.total_sum is not None:
        item.qty = round(item.total_sum / item.price, 3)
        note("количество", item.qty)

    # 6. ставка
    if item.vat_rate is None:
        if item.total_sum and item.vat_sum is not None:
            item.vat_rate = round(item.vat_sum / item.total_sum * 100.0, 2)
            note("ставка НДС", item.vat_rate)
        elif item.sum_with_vat and item.vat_sum is not None:
            base = item.sum_with_vat - item.vat_sum
            if base > 0:
                item.vat_rate = round(item.vat_sum / base * 100.0, 2)
                note("ставка НДС", item.vat_rate)


def _normalize_vat_rate(item: Item) -> None:
    """Ставка НДС всегда в процентах: 0.2 -> 20.

    В части выгрузок ставка лежит долей (0.2), и в таблице появлялось «0,2»
    вместо «20». Значения 0 < ставка < 1 переводим в проценты.
    """
    rate = item.vat_rate
    if rate is None or isinstance(rate, bool):
        return
    if 0 < rate < 1:
        item.vat_rate = round(float(rate) * 100.0, 2)


def validate_items(items: List[Item]) -> ValidationReport:
    """Проверить и достроить позиции. Меняет items «на месте»."""
    report = ValidationReport(items=list(items))

    # режим НДС определяем отдельно для каждого исходного файла
    by_source: Dict[str, List[Item]] = {}
    for it in items:
        by_source.setdefault(it.source, []).append(it)
    for src, group in by_source.items():
        report.vat_mode[src] = detect_vat_mode(group)

    for it in items:
        _normalize_vat_rate(it)
        mode = report.vat_mode.get(it.source, EXCLUSIVE)
        _fill(it, mode, report.issues)
        _check_item(it, mode, report.issues)

    return report


def _issue(
    issues: List[Issue],
    it: Item,
    field_name: str,
    message: str,
    doc_value: Optional[float],
    calc_value: Optional[float],
    severity: str = "error",
) -> None:
    issues.append(
        Issue(
            severity=severity,
            source=it.source,
            sheet=it.sheet,
            row=it.row,
            name=it.name,
            field=field_name,
            message=message,
            doc_value=doc_value,
            calc_value=calc_value,
        )
    )


def _check_item(it: Item, mode: str, issues: List[Issue]) -> None:
    """Проверить арифметику строки, которой не хватало для контроля."""
    # количество × цена = сумма
    if it.qty is not None and it.price is not None and it.total_sum is not None:
        expected = _round2(it.qty * it.price)
        if abs(it.total_sum - expected) > MONEY_TOL:
            _issue(
                issues,
                it,
                "сумма",
                f"сумма {fmt_money(it.total_sum)} ≠ количество × цена "
                f"({fmt_money(it.qty)} × {fmt_money(it.price)} = {fmt_money(expected)})",
                it.total_sum,
                expected,
            )

    # сумма + НДС = сумма с НДС (или сумма = сумма с НДС, если НДС «в том числе»)
    if (
        it.total_sum is not None
        and it.vat_sum is not None
        and it.sum_with_vat is not None
        and "сумма с НДС" not in it.computed
    ):
        if mode == INCLUSIVE:
            expected = _round2(it.total_sum)
            formula = f"сумма (с НДС) {fmt_money(it.total_sum)}"
        else:
            expected = _round2(it.total_sum + it.vat_sum)
            formula = f"сумма + сумма НДС ({fmt_money(it.total_sum)} + {fmt_money(it.vat_sum)} = {fmt_money(expected)})"
        if abs(it.sum_with_vat - expected) > MONEY_TOL:
            _issue(
                issues,
                it,
                "сумма с НДС",
                f"сумма с НДС {fmt_money(it.sum_with_vat)} ≠ {formula}",
                it.sum_with_vat,
                expected,
            )

    # НДС по ставке
    if (
        it.total_sum is not None
        and it.vat_sum is not None
        and it.vat_rate
        and "ставка НДС" not in it.computed
        and "сумма" not in it.computed
    ):
        expected = expected_vat(it.total_sum, it.vat_rate, mode)
        if expected is not None and abs(it.vat_sum - expected) > max(MONEY_TOL, 0.011 * it.vat_rate):
            _issue(
                issues,
                it,
                "сумма НДС",
                f"сумма НДС {fmt_money(it.vat_sum)} не соответствует ставке "
                f"{it.vat_rate:g}% от {fmt_money(it.total_sum)} (ожидалось {fmt_money(expected)})",
                it.vat_sum,
                expected,
            )

    if it.vat_rate is not None and it.vat_rate not in (0.0, 10.0, 20.0, 5.0, 7.0, 25.0):
        _issue(
            issues,
            it,
            "ставка НДС",
            f"нестандартная ставка НДС: {it.vat_rate:g}%",
            it.vat_rate,
            None,
            severity="warning",
        )


def check_document_totals(
    report: ValidationReport, totals: Sequence[DocTotal], blocks: Dict[Tuple[str, str, int], List[Item]]
) -> None:
    """Сверить найденные в документах итоговые строки с суммой наших строк."""
    labels = {"qty": "количество", "sum": "сумма", "vat_sum": "сумма НДС", "total": "сумма с НДС"}
    attrs = {"qty": "qty", "sum": "total_sum", "vat_sum": "vat_sum", "total": "sum_with_vat"}

    def aggregate(group: Sequence[Item], kind: str) -> float:
        return round(sum(getattr(i, attrs[kind]) or 0.0 for i in group), 4)

    for t in totals:
        field = total_field_from_label(t.label)
        if field == "skip":
            continue

        if t.scope == "block":
            group = list(blocks.get((t.source, t.sheet, t.block), []))
            scope_text = "по листу"
        else:
            group = []
            for (src, sh, blk), items in blocks.items():
                if src == t.source and sh == t.sheet and blk <= t.block:
                    group.extend(items)
            scope_text = "накопительно"
        if not group:
            continue

        pairs: List[Tuple[str, float]] = []  # (kind, значение в документе)
        if field in attrs:
            value = None
            if t.values.get(field) is not None:
                value = t.values[field]
            elif t.numbers:
                value = t.numbers[0]
            if value is not None:
                pairs.append((field, value))
        elif field == "":
            pairs.extend((k, v) for k, v in t.values.items() if k in attrs)
            if not pairs and t.numbers:
                # не знаем, что за итог — проверяем, сходится ли он хоть с чем-то
                nums = [n for n in t.numbers if n]
                ok = False
                for kind in attrs:
                    calc = aggregate(group, kind)
                    if any(abs(calc - n) <= max(0.05, abs(n) * 0.001) for n in nums):
                        ok = True
                        break
                if not ok:
                    report.issues.append(
                        Issue(
                            severity="warning",
                            source=t.source,
                            sheet=t.sheet,
                            row=t.row,
                            name=f"«{t.label}» ({scope_text})",
                            field="итог",
                            message=(
                                "итог документа "
                                + ", ".join(fmt_money(n) for n in t.numbers)
                                + " не сходится ни с количеством, ни с суммой, ни с НДС, ни с суммой с НДС"
                            ),
                            calc_value=aggregate(group, "total"),
                        )
                    )
                continue

        for kind, doc_val in pairs:
            calc = aggregate(group, kind)
            tol = QTY_TOL if kind == "qty" else max(0.05, abs(doc_val) * 0.001)
            if abs(calc - doc_val) > tol:
                report.issues.append(
                    Issue(
                        severity="warning",
                        source=t.source,
                        sheet=t.sheet,
                        row=t.row,
                        name=f"«{t.label}» ({scope_text})",
                        field=labels.get(kind, kind),
                        message=(
                            f"итог документа ({labels.get(kind, kind)}) {fmt_money(doc_val)} "
                            f"не совпадает с суммой строк {fmt_money(calc)} "
                            f"(разница {fmt_money(calc - doc_val)})"
                        ),
                        doc_value=doc_val,
                        calc_value=calc,
                    )
                )
