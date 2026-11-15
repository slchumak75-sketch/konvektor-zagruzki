"""Формирование итогового Excel-файла по образцу заказчика.

Лист «Данные» — ровно те колонки, что в образце:

    Штрихкод | Артикул | Наименование товара | ед. изм. | количество |
    цена | сумма | ставка НДС | сумма НДС | сумма с НДС

Дополнительно создаются листы «Проверки» (все расхождения и расчёты) и
«Сводка» (что и откуда взято). Значения, которые пришлось рассчитать
программе, залиты жёлтым — визуально видно, что это не данные поставщика.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from . import APP_NAME, TEMPLATE_COLUMNS, __version__
from .extract import Item
from .validate import MODE_TITLES, Issue, ValidationReport

# --- стили ------------------------------------------------------------------
COLOR_HEADER = "1F4E78"
COLOR_TOTAL = "DDEBF7"
COLOR_COMPUTED = "FFF2CC"
COLOR_ERROR = "F8CBAD"
COLOR_WARNING = "FFE699"
COLOR_INFO = "E7E6E6"

THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill("solid", fgColor=COLOR_HEADER)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP_TOP = Alignment(vertical="top", wrap_text=True)
CENTER_TOP = Alignment(horizontal="center", vertical="top")

COLUMN_WIDTHS = [16, 18, 58, 9, 11, 12, 12, 10, 12, 13]
MONEY_FMT = "#,##0.00"
QTY_FMT = "#,##0.###"
RATE_FMT = "0.##"

INVALID_SHEET = re.compile(r"[\\/*?:\[\]]")


@dataclass
class FileStat:
    """Сводка по одному исходному файлу."""

    path: str
    sheets: str = ""
    reader: str = ""
    items: int = 0
    sum_without_vat: float = 0.0
    vat_sum: float = 0.0
    sum_with_vat: float = 0.0
    errors: int = 0
    warnings: int = 0
    computed: int = 0
    vat_mode: str = ""
    #: реквизиты документа, прочитанные из файла (konvektor/meta.py)
    supplier: str = ""
    supplier_unp: str = ""
    doc_type: str = ""
    doc_number: str = ""
    doc_date: str = ""
    meta_warnings: List[str] = field(default_factory=list)
    #: качество разбора (konvektor/analysis.py)
    quality: str = ""
    confidence: float = 0.0
    roles_found: int = 0
    review: bool = False

    @property
    def doc_label(self) -> str:
        """«Счет № 542» — тип и номер документа одной строкой."""
        parts = [self.doc_type, f"№ {self.doc_number}" if self.doc_number else ""]
        return " ".join(p for p in parts if p).strip()


@dataclass
class WorkbookPayload:
    """Всё, что нужно для записи книги."""

    data_sheets: List[Tuple[str, List[Item]]]
    report: ValidationReport
    stats: List[FileStat] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)
    with_source: bool = False


def safe_sheet_title(title: str, used: Sequence[str] = ()) -> str:
    t = INVALID_SHEET.sub(" ", title).strip() or "Лист"
    t = t[:31]
    base = t
    i = 2
    existing = {u.lower() for u in used}
    while t.lower() in existing:
        suffix = f" ({i})"
        t = (base[: 31 - len(suffix)] + suffix).strip()
        i += 1
    return t


def _write_header(ws: Worksheet, row: int, columns: Sequence[str]) -> None:
    for c, title in enumerate(columns, start=1):
        cell = ws.cell(row=row, column=c, value=title)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def _write_data_sheet(ws: Worksheet, items: Sequence[Item], with_source: bool = False) -> int:
    columns = list(TEMPLATE_COLUMNS)
    widths = list(COLUMN_WIDTHS)
    if with_source:
        columns += ["файл", "строка"]
        widths += [38, 8]
    _write_header(ws, 1, columns)
    for i, _ in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]

    computed_fill = PatternFill("solid", fgColor=COLOR_COMPUTED)
    r = 2
    for it in items:
        values: List[Optional[object]] = [
            it.barcode or "",
            it.article or "",
            it.name,
            it.unit or "",
            it.qty,
            it.price,
            it.total_sum,
            it.vat_rate,
            it.vat_sum,
            it.sum_with_vat,
        ]
        if with_source:
            values += [os.path.basename(it.source), it.row]
        for c, val in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=val if val != "" else None)
            cell.border = BORDER
            if c == 3:
                cell.alignment = WRAP_TOP
            elif c in (1, 2, 4):
                cell.alignment = CENTER_TOP
                cell.number_format = "@"
            else:
                cell.alignment = CENTER_TOP
        # штрихкод и артикул — всегда текстом (иначе Excel съест ведущие нули)
        for c in (1, 2):
            cell = ws.cell(row=r, column=c)
            if cell.value is not None:
                cell.value = str(cell.value)
                cell.number_format = "@"
        for c, fmt in ((5, QTY_FMT), (6, MONEY_FMT), (7, MONEY_FMT), (8, RATE_FMT), (9, MONEY_FMT), (10, MONEY_FMT)):
            ws.cell(row=r, column=c).number_format = fmt

        for fname, col in (
            ("количество", 5),
            ("цена", 6),
            ("сумма", 7),
            ("ставка НДС", 8),
            ("сумма НДС", 9),
            ("сумма с НДС", 10),
        ):
            if fname in it.computed:
                ws.cell(row=r, column=col).fill = computed_fill
        r += 1

    # итоговая строка с формулами
    if items:
        last = r - 1
        total_row = r
        labels = {3: "ИТОГО"}
        for col, text in labels.items():
            cell = ws.cell(row=total_row, column=col, value=text)
            cell.font = Font(bold=True)
        for col in (7, 9, 10):
            letter = get_column_letter(col)
            cell = ws.cell(row=total_row, column=col, value=f"=SUM({letter}2:{letter}{last})")
            cell.font = Font(bold=True)
            cell.number_format = MONEY_FMT
        for col in range(1, len(columns) + 1):
            cell = ws.cell(row=total_row, column=col)
            cell.fill = PatternFill("solid", fgColor=COLOR_TOTAL)
            cell.border = Border(left=THIN, right=THIN, top=Side(style="double", color="808080"), bottom=THIN)
        r += 1

    ws.freeze_panes = "A2"
    if items:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{r - 2}"
    return r


def _write_issues_sheet(ws: Worksheet, issues: Sequence[Issue]) -> None:
    headers = [
        "№",
        "Тип",
        "Файл",
        "Лист",
        "Строка",
        "Наименование",
        "Поле",
        "Значение в документе",
        "Расчёт программы",
        "Разница",
        "Комментарий",
    ]
    widths = [5, 12, 32, 14, 8, 44, 14, 18, 18, 12, 70]
    _write_header(ws, 1, headers)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    severity_title = {"error": "ОШИБКА", "warning": "ВНИМАНИЕ", "info": "РАСЧЁТ"}
    severity_fill = {
        "error": PatternFill("solid", fgColor=COLOR_ERROR),
        "warning": PatternFill("solid", fgColor=COLOR_WARNING),
        "info": PatternFill("solid", fgColor=COLOR_INFO),
    }

    if not issues:
        ws.cell(row=2, column=1, value="Расхождений не найдено — все суммы сошлись.").font = Font(bold=True)
        return

    for i, issue in enumerate(issues, start=1):
        values = [
            i,
            severity_title.get(issue.severity, issue.severity),
            os.path.basename(issue.source),
            issue.sheet,
            issue.row,
            issue.name,
            issue.field,
            issue.doc_value,
            issue.calc_value,
            issue.delta,
            issue.message,
        ]
        for c, val in enumerate(values, start=1):
            cell = ws.cell(row=i + 1, column=c, value=val)
            cell.border = BORDER
            cell.alignment = WRAP_TOP if c in (3, 6, 11) else CENTER_TOP
            if c in (8, 9, 10):
                cell.number_format = MONEY_FMT
        ws.cell(row=i + 1, column=2).fill = severity_fill.get(issue.severity, severity_fill["info"])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{len(issues) + 1}"


def _write_summary_sheet(ws: Worksheet, payload: WorkbookPayload) -> None:
    headers = [
        "Файл",
        "Формат",
        "Лист(ы)",
        "Поставщик",
        "УНП поставщика",
        "Документ",
        "Дата документа",
        "Позиций",
        "Сумма",
        "Сумма НДС",
        "Сумма с НДС",
        "Ошибок",
        "Предупреждений",
        "Рассчитано значений",
        "Цены",
        "Качество разбора",
        "Доверие",
        "Проверить",
    ]
    widths = [46, 9, 26, 38, 15, 26, 14, 10, 15, 15, 16, 9, 15, 20, 26, 18, 10, 10]
    _write_header(ws, 1, headers)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 2
    for st in payload.stats:
        values = [
            st.path,
            st.reader,
            st.sheets,
            st.supplier or "— не найден",
            st.supplier_unp,
            st.doc_label or "—",
            st.doc_date or "—",
            st.items,
            st.sum_without_vat,
            st.vat_sum,
            st.sum_with_vat,
            st.errors,
            st.warnings,
            st.computed,
            MODE_TITLES.get(st.vat_mode, st.vat_mode),
            st.quality,
            round(st.confidence, 2) if st.confidence else None,
            "да" if st.review else "",
        ]
        for c, val in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=val if val != "" else None)
            cell.border = BORDER
            cell.alignment = WRAP_TOP if c in (1, 3, 4, 6, 15) else CENTER_TOP
            if c == 18 and st.review and val:
                cell.fill = PatternFill("solid", fgColor=COLOR_WARNING)
            if c in (9, 10, 11):
                cell.number_format = MONEY_FMT
            if c in (5, 7):
                cell.number_format = "@"
        r += 1

    # итог
    if payload.stats:
        cell = ws.cell(row=r, column=1, value="ВСЕГО")
        cell.font = Font(bold=True)
        ws.cell(row=r, column=8, value=sum(s.items for s in payload.stats)).font = Font(bold=True)
        for col, attr in ((9, "sum_without_vat"), (10, "vat_sum"), (11, "sum_with_vat")):
            c = ws.cell(row=r, column=col, value=round(sum(getattr(s, attr) for s in payload.stats), 2))
            c.font = Font(bold=True)
            c.number_format = MONEY_FMT
        for col in range(1, len(headers) + 1):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=COLOR_TOTAL)
            ws.cell(row=r, column=col).border = BORDER
        r += 1

    r += 1
    info = [
        ("Отчёт сформирован", f"{APP_NAME} {__version__}"),
        ("Дата формирования", _dt.datetime.now().strftime("%d.%m.%Y %H:%M")),
        ("Проверки", "сумма = кол-во × цена; сумма с НДС = сумма + НДС; НДС = сумма × ставка"),
        ("Жёлтая заливка", "значение рассчитано программой (в документе отсутствовало)"),
        ("Лист «Проверки»", "все расхождения, предупреждения и расчёты"),
        ("Поставщик и дата", "прочитаны из файла поставщика (см. столбцы «Поставщик» и «Дата документа»)"),
    ]
    for title, text in info:
        ws.cell(row=r, column=1, value=title).font = Font(bold=True)
        ws.cell(row=r, column=2, value=text)
        r += 1
    if payload.notes:
        r += 1
        ws.cell(row=r, column=1, value="Замечания по чтению файлов").font = Font(bold=True)
        r += 1
        for note in payload.notes:
            ws.cell(row=r, column=1, value=note)
            r += 1


def write_workbook(path: str, payload: WorkbookPayload) -> str:
    """Записать итоговую книгу. Возвращает путь к файлу."""
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    used: List[str] = []
    for title, items in payload.data_sheets:
        name = safe_sheet_title(title, used)
        used.append(name)
        ws = wb.create_sheet(name)
        _write_data_sheet(ws, items, with_source=payload.with_source)

    issues_ws = wb.create_sheet(safe_sheet_title("Проверки", used))
    used.append("Проверки")
    _write_issues_sheet(issues_ws, payload.report.issues)

    summary_ws = wb.create_sheet(safe_sheet_title("Сводка", used))
    _write_summary_sheet(summary_ws, payload)

    wb.properties.title = "Конвертированная таблица"
    wb.properties.creator = f"{APP_NAME} {__version__}"
    wb.properties.description = "Таблица по образцу: штрихкод, артикул, наименование, ед. изм., количество, цена, сумма, ставка НДС, сумма НДС, сумма с НДС"

    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    wb.save(path)
    return path
