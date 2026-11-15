"""Чтение файлов в единое представление (список листов с «сырыми» значениями).

Поддерживается:
  * .xlsx / .xlsm  — openpyxl;
  * .xls / .et      — xlrd (в т.ч. «кривые» выгрузки 1С: ignore_workbook_corruption);
  * .html / .htm / .mht — таблицы из HTML (часто так отдают прайсы);
  * .csv / .txt / .tsv — CSV с автоопределением кодировки и разделителя.

Формат определяется по содержимому файла (магические байты), а не по
расширению, потому что поставщики регулярно присылают .xls, внутри которого
лежит xlsx или HTML.
"""

from __future__ import annotations

import csv
import io
import re
import zipfile
from dataclasses import dataclass, field
from html.parser import HTMLParser
from typing import Any, List, Optional

from .textutil import clean_text, to_text


@dataclass
class Sheet:
    """Один лист: имя + матрица значений (None там, где пусто)."""

    name: str
    rows: List[List[Any]] = field(default_factory=list)

    @property
    def width(self) -> int:
        return max((len(r) for r in self.rows), default=0)

    @property
    def height(self) -> int:
        return len(self.rows)

    def cell(self, r: int, c: int) -> Any:
        if 0 <= r < len(self.rows):
            row = self.rows[r]
            if 0 <= c < len(row):
                return row[c]
        return None

    def row(self, r: int) -> List[Any]:
        return self.rows[r] if 0 <= r < len(self.rows) else []


@dataclass
class SourceFile:
    path: str
    sheets: List[Sheet] = field(default_factory=list)
    reader: str = ""
    warnings: List[str] = field(default_factory=list)


MAGIC_OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
MAGIC_ZIP = b"PK\x03\x04"


# --------------------------------------------------------------------------- #
#  Определение формата
# --------------------------------------------------------------------------- #
def sniff_kind(data: bytes, filename: str = "") -> str:
    head = data[:4096]
    if head.startswith(MAGIC_ZIP):
        return "xlsx"
    if head.startswith(MAGIC_OLE):
        return "xls"
    if zipfile.is_zipfile(io.BytesIO(data)):
        return "xlsx"
    stripped = head.lstrip(b"\xef\xbb\xbf \t\r\n")
    low = stripped[:512].lower()
    if stripped[:1] == b"<" or b"<html" in low or b"<?xml" in low or b"<table" in low:
        if b"<table" in low or b"<html" in low or b"urn:schemas-microsoft-com:office:spreadsheet" in low:
            return "html"
        return "html"
    name = filename.lower()
    if name.endswith((".html", ".htm", ".mht", ".mhtml")):
        return "html"
    if name.endswith((".csv", ".txt", ".tsv")):
        return "csv"
    return "csv"


# --------------------------------------------------------------------------- #
#  XLSX
# --------------------------------------------------------------------------- #
def _read_xlsx(data: bytes) -> List[Sheet]:
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    sheets: List[Sheet] = []
    try:
        for ws in wb.worksheets:
            rows: List[List[Any]] = []
            for row in ws.iter_rows(values_only=True):
                if row is None:
                    rows.append([])
                    continue
                vals = list(row)
                while vals and vals[-1] is None:
                    vals.pop()
                rows.append([_norm_cell(v) for v in vals])
            # убрать хвостовые пустые строки
            while rows and not any(v is not None for v in rows[-1]):
                rows.pop()
            sheets.append(Sheet(name=ws.title or f"Лист{len(sheets) + 1}", rows=rows))
    finally:
        wb.close()
    return sheets


def _norm_cell(v: Any) -> Any:
    if isinstance(v, str):
        return v
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v
    # datetime, date, time, Decimal, прочее
    return to_text(v)


# --------------------------------------------------------------------------- #
#  XLS (BIFF)
# --------------------------------------------------------------------------- #
_MOJIBAKE_MARKS = "ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞß"


def _mojibake_score(sheets: List[Sheet]) -> int:
    """>0 — кириллица прочиталась нормально, <0 — похоже на cp1251-as-latin1."""
    cyr = 0
    bad = 0
    for sh in sheets[:3]:
        for row in sh.rows[:200]:
            for v in row:
                if not isinstance(v, str):
                    continue
                for ch in v:
                    o = ord(ch)
                    if 0x0400 <= o <= 0x04FF:
                        cyr += 1
                    elif ch in _MOJIBAKE_MARKS:
                        bad += 1
    return cyr - bad


def _read_xls(data: bytes) -> List[Sheet]:
    import xlrd

    last_err: Optional[Exception] = None
    for enc in (None, "cp1251"):
        try:
            wb = xlrd.open_workbook(
                file_contents=data,
                ignore_workbook_corruption=True,
                encoding_override=enc,
                logfile=io.StringIO(),
            )
        except Exception as exc:  # noqa: BLE001 - нужен фолбэк на другие ридеры
            last_err = exc
            continue
        sheets: List[Sheet] = []
        datemode = getattr(wb, "datemode", 0)
        for sh in wb.sheets():
            rows: List[List[Any]] = []
            for r in range(sh.nrows):
                vals: List[Any] = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    vals.append(_xls_value(cell, datemode))
                while vals and vals[-1] is None:
                    vals.pop()
                rows.append(vals)
            while rows and not any(v is not None for v in rows[-1]):
                rows.pop()
            sheets.append(Sheet(name=sh.name or "Лист", rows=rows))
        if enc is None:
            if _mojibake_score(sheets) > 0:
                return sheets
            continue  # похоже на cp1251-баг — пробуем с явной кодировкой
        return sheets
    if last_err:
        raise last_err
    raise ValueError("не удалось прочитать файл .xls")


def _xls_value(cell, datemode: int = 0) -> Any:
    import xlrd

    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return None
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        v = float(cell.value)
        return int(v) if v.is_integer() and abs(v) < 1e15 else v
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    if cell.ctype == xlrd.XL_CELL_DATE:
        try:
            return xlrd.xldate.xldate_as_datetime(cell.value, datemode).date()
        except Exception:  # noqa: BLE001 - битая дата в файле
            return None
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return None
    return cell.value


# --------------------------------------------------------------------------- #
#  HTML
# --------------------------------------------------------------------------- #
class _TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: List[List[List[str]]] = []
        self._stack: List[List[List[str]]] = []
        self._row: Optional[List[str]] = None
        self._cell: Optional[List[str]] = None
        self._span: List[tuple] = []
        self._in_cell = False
        self._colspans: List[tuple] = []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "table":
            self._stack.append([])
        elif tag == "tr" and self._stack:
            self._row = []
            self._span = []
            self._stack[-1].append(self._row)
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []
            self._in_cell = True
            try:
                cs = max(1, int(a.get("colspan", 1)))
            except (TypeError, ValueError):
                cs = 1
            try:
                rs = max(1, int(a.get("rowspan", 1)))
            except (TypeError, ValueError):
                rs = 1
            self._cell_span = (cs, rs)
        elif tag == "br" and self._in_cell and self._cell is not None:
            self._cell.append("\n")

    def handle_data(self, data):
        if self._in_cell and self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._in_cell and self._cell is not None:
            text = "".join(self._cell)
            text = re.sub(r"[ \t\r]+", " ", text).strip()
            cs, rs = getattr(self, "_cell_span", (1, 1))
            self._row.append(text)
            for _ in range(cs - 1):
                self._row.append("")
            for _ in range(rs - 1):
                self._span.append((len(self._row) - 1, cs))
            self._cell = None
            self._in_cell = False
        elif tag == "tr" and self._stack:
            self._row = None
        elif tag == "table" and self._stack:
            self.tables.append(self._stack.pop())


def _read_html(data: bytes) -> List[Sheet]:
    for enc in ("utf-8-sig", "cp1251", "utf-16", "latin-1"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = data.decode("utf-8", "replace")
    # SpreadsheetML 2003 (XML) — таблицы тоже читаем как «строки/ячейки»
    if "urn:schemas-microsoft-com:office:spreadsheet" in text:
        return _read_spreadsheetml(text)
    parser = _TableParser()
    parser.feed(text)
    parser.close()
    sheets: List[Sheet] = []
    for i, table in enumerate(parser.tables):
        rows: List[List[Any]] = []
        for row in table:
            vals = [v if v != "" else None for v in row]
            while vals and vals[-1] is None:
                vals.pop()
            rows.append(vals)
        while rows and not any(v is not None for v in rows[-1]):
            rows.pop()
        if any(any(v is not None for v in r) for r in rows):
            sheets.append(Sheet(name=f"Таблица{i + 1}", rows=rows))
    if not sheets:
        raise ValueError("в HTML не найдено ни одной таблицы")
    return sheets


def _read_spreadsheetml(text: str) -> List[Sheet]:
    import xml.etree.ElementTree as ET

    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    root = ET.fromstring(text.encode("utf-8"))
    sheets: List[Sheet] = []
    for ws in root.findall(".//ss:Worksheet", ns):
        name = ws.get("{urn:schemas-microsoft-com:office:spreadsheet}Name") or "Лист"
        rows: List[List[Any]] = []
        for row in ws.findall(".//ss:Row", ns):
            vals: List[Any] = []
            for cell in row.findall("ss:Cell", ns):
                idx = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
                if idx:
                    while len(vals) < int(idx) - 1:
                        vals.append(None)
                data = cell.find("ss:Data", ns)
                vals.append(data.text if data is not None else None)
            while vals and vals[-1] is None:
                vals.pop()
            rows.append(vals)
        while rows and not any(v is not None for v in rows[-1]):
            rows.pop()
        sheets.append(Sheet(name=clean_text(name) or "Лист", rows=rows))
    if not sheets:
        raise ValueError("в SpreadsheetML не найдено листов")
    return sheets


# --------------------------------------------------------------------------- #
#  CSV / TXT
# --------------------------------------------------------------------------- #
def _read_csv(data: bytes) -> List[Sheet]:
    text = None
    for enc in ("utf-8-sig", "cp1251", "utf-16", "latin-1"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = data.decode("utf-8", "replace")

    sample = text[:8000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,\t|")
        delim = dialect.delimiter
    except csv.Error:
        counts = {d: sample.count(d) for d in ";,\t|"}
        delim = max(counts, key=counts.get) if max(counts.values()) else ";"

    reader = csv.reader(io.StringIO(text), delimiter=delim)
    rows: List[List[Any]] = []
    for raw in reader:
        vals = [v.strip() or None for v in raw]
        while vals and vals[-1] is None:
            vals.pop()
        rows.append(vals)
    while rows and not any(v is not None for v in rows[-1]):
        rows.pop()
    return [Sheet(name="CSV", rows=rows)]


# --------------------------------------------------------------------------- #
#  Публичный API
# --------------------------------------------------------------------------- #
def read_source(path: str) -> SourceFile:
    """Прочитать файл и вернуть SourceFile (без падений: ошибки в warnings)."""
    with open(path, "rb") as fh:
        data = fh.read()

    sf = SourceFile(path=path)
    kind = sniff_kind(data, path)

    attempts = []
    if kind == "xlsx":
        attempts = [("xlsx", _read_xlsx), ("xls", _read_xls), ("html", _read_html)]
    elif kind == "xls":
        attempts = [("xls", _read_xls), ("xlsx", _read_xlsx), ("html", _read_html)]
    else:
        attempts = [("html", _read_html), ("xlsx", _read_xlsx), ("xls", _read_xls), ("csv", _read_csv)]

    errors: List[str] = []
    for name, fn in attempts:
        try:
            sheets = fn(data)
        except Exception as exc:  # noqa: BLE001 - пробуем следующий ридер
            errors.append(f"{name}: {exc}")
            continue
        if sheets:
            sf.sheets = sheets
            sf.reader = name
            return sf
        errors.append(f"{name}: пустой результат")

    raise ValueError("не удалось прочитать файл: " + "; ".join(errors))


def iter_source_files(paths) -> List[SourceFile]:
    out: List[SourceFile] = []
    for p in paths:
        out.append(read_source(p))
    return out
