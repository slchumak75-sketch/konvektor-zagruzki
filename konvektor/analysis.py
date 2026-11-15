"""Отчёт о разборе файла: что программа поняла и насколько уверенно.

Идея «общей концепции» в терминах, которые можно проверить:

* **Покрытие полей** — сколько из 10 смысловых колонок найдено (роли).
* **Арифметическая согласованность** — доля сошедшихся проверок
  (кол-во × цена = сумма, сумма + НДС = всего, сумма × ставка = НДС).
* **Полнота строк** — какая доля позиций имеет все ключевые поля.
* **Доверие** (`confidence`) — свёртка этих чисел; ниже порога разбор
  помечается «требует проверки» (abstain), а не выдаётся молча.

Отчёт печатается ключом `--explain`, кладётся в `--json` и в `--diagnose`,
и попадает на лист «Сводка» столбцом «Качество разбора».
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence

from . import detect
from .extract import Item, SheetExtraction
from .layout import AnalysisDecision, BlockInfo
from .meta import DocMeta
from .readers import SourceFile

#: Порог, ниже которого результат помечается «требует проверки».
CONFIDENCE_OK = 0.75
CONFIDENCE_MID = 0.5

ROLE_TITLES = {
    detect.BARCODE: "штрихкод",
    detect.ARTICLE: "артикул",
    detect.NAME: "наименование",
    detect.UNIT: "ед. изм.",
    detect.QTY: "количество",
    detect.PRICE: "цена",
    detect.SUM: "сумма",
    detect.VAT_RATE: "ставка НДС",
    detect.VAT_SUM: "сумма НДС",
    detect.TOTAL: "сумма с НДС",
}


def col_letter(index: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA."""
    index = int(index)
    letters = ""
    while True:
        letters = chr(ord("A") + index % 26) + letters
        index = index // 26 - 1
        if index < 0:
            return letters


@dataclass
class AnalysisReport:
    """Разбор одного файла: метрики, блоки, решения, доверие."""

    path: str
    reader: str = ""
    sheets: List[str] = field(default_factory=list)
    items: int = 0
    roles_found: int = 0
    roles_total: int = 10
    #: доля сошедшейся арифметики; None — сверять было нечего
    arithmetic: Optional[float] = None
    filled_ratio: float = 0.0
    confidence: float = 0.0
    quality: str = "низкое"
    reasons: List[str] = field(default_factory=list)
    blocks: List[BlockInfo] = field(default_factory=list)
    decisions: List[AnalysisDecision] = field(default_factory=list)
    meta: Optional[DocMeta] = None
    warnings: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.confidence >= CONFIDENCE_OK

    @property
    def needs_review(self) -> bool:
        return self.confidence < CONFIDENCE_OK

    def as_dict(self) -> Dict[str, Any]:
        return {
            "file": self.path,
            "reader": self.reader,
            "sheets": self.sheets,
            "items": self.items,
            "coverage": f"{self.roles_found}/{self.roles_total}",
            "arithmetic": None if self.arithmetic is None else round(self.arithmetic, 3),
            "filled": round(self.filled_ratio, 3),
            "confidence": round(self.confidence, 3),
            "quality": self.quality,
            "needs_review": self.needs_review,
            "reasons": self.reasons,
            "blocks": [
                {
                    "sheet": b.sheet,
                    "header_row": b.row,
                    "span": b.span,
                    "transposed": b.transposed,
                    "arithmetic": None if b.arithmetic is None else round(b.arithmetic, 3),
                    "columns": {
                        role: f"{col_letter(col)} ({b.headers.get(col, '')})".strip()
                        for role, col in sorted(b.roles.items())
                    },
                }
                for b in self.blocks
            ],
            "decisions": [d.as_text() for d in self.decisions],
            "warnings": self.warnings,
        }


_KEY_FIELDS = ("name", "qty", "price", "total_sum", "vat_rate", "vat_sum", "sum_with_vat")


def _filled_ratio(items: Sequence[Item]) -> float:
    if not items:
        return 0.0
    got = 0
    for it in items:
        got += sum(1 for f in _KEY_FIELDS if getattr(it, f) is not None)
    return got / (len(items) * len(_KEY_FIELDS))


def build_report(
    path: str,
    source: Optional[SourceFile],
    extraction: SheetExtraction,
    meta: Optional[DocMeta] = None,
) -> AnalysisReport:
    """Свести метрики разбора файла в один отчёт."""
    report = AnalysisReport(path=path)
    if source is not None:
        report.reader = source.reader
        report.sheets = [s.name for s in source.sheets]
    report.items = len(extraction.items)
    report.blocks = list(extraction.blocks)
    report.decisions = list(extraction.decisions)
    report.warnings = list(extraction.warnings)
    report.meta = meta

    roles = set()
    for block in report.blocks:
        roles.update(block.roles.keys())
    report.roles_found = len(roles)
    known = [b.arithmetic for b in report.blocks if b.arithmetic is not None]
    report.arithmetic = sum(known) / len(known) if known else None
    report.filled_ratio = _filled_ratio(extraction.items)

    # --- доверие: покрытие полей, арифметика, полнота строк ---------------
    if not extraction.items:
        report.confidence = 0.0
        report.quality = "низкое"
        report.reasons.append("позиции не найдены")
        return report

    coverage = report.roles_found / report.roles_total
    # если сверять нечего (в документе нет колонок для сверки) — нейтральная оценка
    arithmetic = 0.7 if report.arithmetic is None else report.arithmetic
    confidence = 0.35 + 0.35 * coverage + 0.2 * arithmetic + 0.1 * report.filled_ratio
    if not report.blocks:
        confidence -= 0.1
    report.confidence = round(max(0.0, min(1.0, confidence)), 3)

    if coverage < 1.0:
        missing = [ROLE_TITLES[r] for r in ROLE_TITLES if r not in roles]
        report.reasons.append("не найдены колонки: " + ", ".join(missing))
    if report.arithmetic is None:
        report.reasons.append("арифметику сверить не с чем: в документе нет нужных колонок")
    elif report.arithmetic < 0.95:
        report.reasons.append(
            f"арифметика сходится на {report.arithmetic:.0%} строк — проверьте лист «Проверки»"
        )
    if report.filled_ratio < 0.9:
        report.reasons.append(
            f"заполнено {report.filled_ratio:.0%} ключевых полей — часть значений посчитана программой"
        )
    if report.decisions:
        report.reasons.extend(d.as_text() for d in report.decisions)

    if report.confidence >= CONFIDENCE_OK:
        report.quality = "высокое"
    elif report.confidence >= CONFIDENCE_MID:
        report.quality = "среднее"
    else:
        report.quality = "низкое"
    return report


def format_report(report: AnalysisReport) -> str:
    """Человекочитаемый отчёт для консоли (ключ --explain)."""
    lines: List[str] = []
    lines.append(f"Файл: {os.path.basename(report.path)}")
    if report.reader:
        lines.append(f"  формат: {report.reader}; листы: {', '.join(report.sheets) or '—'}")
    if report.meta is not None:
        meta = report.meta
        lines.append(
            "  документ: "
            + (meta.doc_label or "— не определён")
            + (f" от {meta.date_text}" if meta.doc_date else "")
            + "; поставщик: "
            + (meta.supplier or "— не найден")
            + (f" (УНП {meta.supplier_unp})" if meta.supplier_unp else "")
        )
    lines.append(
        f"  позиций: {report.items}; колонок распознано: "
        f"{report.roles_found}/{report.roles_total}; "
        + (
            "арифметика: не проверялась"
            if report.arithmetic is None
            else f"арифметика: {report.arithmetic:.0%}"
        )
        + f"; поля заполнены: {report.filled_ratio:.0%}"
    )
    lines.append(
        f"  доверие: {report.confidence:.2f} ({report.quality})"
        + (" — требует проверки" if report.needs_review else "")
    )

    for block in report.blocks:
        header = f"  блок: лист «{block.sheet}», шапка в строке {block.row}"
        if block.span > 1:
            header += f" (склеена из {block.span} строк)"
        if block.transposed:
            header += ", таблица развёрнута"
        lines.append(header)
        for role, col in sorted(block.roles.items(), key=lambda kv: kv[1]):
            title = block.headers.get(col, "")
            quality = block.quality.get(role, 0.0)
            lines.append(
                f"    {ROLE_TITLES.get(role, role):<14} -> {col_letter(col)}  «{title}»"
                + (f"  (вес {quality:.1f})" if quality else "")
            )

    if report.reasons:
        lines.append("  замечания и решения:")
        for reason in report.reasons:
            lines.append("    - " + reason)
    return "\n".join(lines)
