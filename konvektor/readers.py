"""Чтение файлов в единое представление (список листов с «сырыми» значениями).

Поддерживается:
  * .xlsx / .xlsm  — openpyxl;
  * .xls / .et      — xlrd (в т.ч. «кривые» выгрузки 1С: ignore_workbook_corruption);
  * .html / .htm / .mht — таблицы из HTML (часто так отдают прайсы);
  * .csv / .txt / .tsv — CSV с автоопределением кодировки и разделителя;
  * .rtf            — RTF с таблицами (МАВ: через striprtf);
  * .pdf            — PDF с таблицами (Элсви: через pdfplumber).

Формат определяется по содержимому файла (магические байты), а не по
расширению, потому что поставщики регулярно присылают .xls, внутри которого
лежит xlsx или HTML.
"""

from __future__ import annotations

import csv
import io
import re
import struct
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
    #: Надписи из текстовых полей/рисунков (в xlsx реквизиты часто нарисованы
    #: поверх листа и в ячейках их нет: «ПОСТАВЩИК: ООО "ТД Комплект" УНП …»).
    texts: List[str] = field(default_factory=list)


MAGIC_OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
MAGIC_ZIP = b"PK\x03\x04"


# --------------------------------------------------------------------------- #
#  Определение формата
# --------------------------------------------------------------------------- #
MAGIC_RTF = b"{\\rtf"
MAGIC_PDF = b"%PDF"


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
    if stripped.startswith(b"{\\rtf") or head.strip().startswith(MAGIC_RTF):
        return "rtf"
    if stripped.startswith(b"%PDF") or head.strip().startswith(MAGIC_PDF):
        return "pdf"
    if stripped[:1] == b"<" or b"<html" in low or b"<?xml" in low or b"<table" in low:
        if b"<table" in low or b"<html" in low or b"urn:schemas-microsoft-com:office:spreadsheet" in low:
            return "html"
        return "html"
    name = filename.lower()
    if name.endswith((".html", ".htm", ".mht", ".mhtml")):
        return "html"
    if name.endswith(".rtf"):
        return "rtf"
    if name.endswith(".pdf"):
        return "pdf"
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


def xlsx_drawing_texts(data: bytes) -> List[str]:
    """Тексты из текстовых полей и фигур xlsx (по одному абзацу на строку).

    Реквизиты в счетах, выгруженных из 1С, нередко лежат не в ячейках,
    а в текстовом поле поверх листа — без этого поставщик «не находится».
    """
    import xml.etree.ElementTree as ET

    texts: List[str] = []
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except Exception:  # noqa: BLE001 - не zip или битый файл
        return texts
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    try:
        names = [n for n in zf.namelist() if re.match(r"xl/drawings/.*\.xml$", n)]
        for name in names:
            try:
                root = ET.fromstring(zf.read(name))
            except Exception:  # noqa: BLE001 - битый рисунок не мешает чтению
                continue
            for paragraph in root.iter(ns + "p"):
                parts = [node.text or "" for node in paragraph.iter(ns + "t")]
                line = re.sub(r"\s+", " ", "".join(parts)).strip()
                if line:
                    texts.append(line)
    finally:
        zf.close()
    return texts


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
        if enc is None and _mojibake_score(sheets) <= 0:
            continue  # похоже на cp1251-баг — пробуем с явной кодировкой
        _xls_embedded_barcodes(data, sheets)
        return sheets
    if last_err:
        raise last_err
    raise ValueError("не удалось прочитать файл .xls")


def _xls_embedded_barcodes(data: bytes, sheets: List[Sheet]) -> None:
    """Считать штрихкоды BarcodeCtrl, привязанные к строкам старого XLS.

    Такие объекты находятся поверх ячеек и не читаются xlrd. Привязку берём
    из BIFF-записи OBJ, код — из соответствующего OLE-потока Contents; порядок
    потоков сам по себе не связан с порядком товаров.
    """
    import olefile

    marker = b"BARCODE.BarcodeCtrl.1\x00"
    try:
        with olefile.OleFileIO(io.BytesIO(data)) as ole:
            stream = "Workbook" if ole.exists("Workbook") else "Book"
            book = ole.openstream(stream).read()
            records = []
            pos = 0
            while pos + 4 <= len(book):
                offset = pos
                kind, size = struct.unpack_from("<HH", book, pos)
                pos += 4
                if pos + size > len(book):
                    break
                records.append((offset, kind, book[pos:pos + size]))
                pos += size
            bounds = [struct.unpack_from("<I", value)[0] for _, kind, value in records
                      if kind == 0x0085 and len(value) >= 4]
            if len(bounds) != len(sheets) or bounds != sorted(bounds):
                return
            codes = {}
            for offset, kind, value in records:
                if kind != 0x005D or marker not in value or len(value) < 18:
                    continue
                sheet_index = next((i for i in range(len(bounds) - 1, -1, -1)
                                    if bounds[i] <= offset), None)
                if sheet_index is None:
                    continue
                row = struct.unpack_from("<H", value, 14)[0]
                marker_end = value.index(marker) + len(marker)
                if marker_end + 4 > len(value):
                    continue
                object_id = struct.unpack_from("<I", value, marker_end)[0]
                path = [f"MBD{object_id:08x}", "Contents"]
                if not ole.exists(path):
                    continue
                content = ole.openstream(path).read()
                matches = re.findall(rb"EAN 13\x0d([0-9]{13})", content)
                if len(matches) != 1 or row >= len(sheets[sheet_index].rows):
                    continue
                key = (sheet_index, row)
                code = matches[0].decode("ascii")
                if key in codes and codes[key] != code:
                    codes[key] = None  # неоднозначность не превращаем в чужой код
                else:
                    codes[key] = code
            for sheet_index, sheet in enumerate(sheets):
                selected = [(row, code) for (index, row), code in codes.items()
                            if index == sheet_index and code]
                if not selected:
                    continue
                col = sheet.width
                for row, code in selected:
                    values = sheet.rows[row]
                    values.extend([None] * (col + 1 - len(values)))
                    values[col] = code
    except (OSError, ValueError, KeyError, struct.error):
        # Обычные XLS без соответствующего OLE-объекта читаются как прежде.
        return


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
#  RTF (МАВ)
# --------------------------------------------------------------------------- #
def _read_rtf(data: bytes) -> List[Sheet]:
    """Прочитать RTF с таблицами (МАВ). Использует striprtf для извлечения текста."""

    # Декодируем как надо — RTF от МАВ в cp1251
    for enc in ("utf-8", "cp1251", "windows-1251", "utf-16"):
        try:
            rtf_text = data.decode(enc)
            if "{\\rtf" in rtf_text:
                break
        except UnicodeDecodeError:
            continue
    else:
        rtf_text = data.decode("utf-8", "replace")

    try:
        from striprtf.striprtf import rtf_to_text

        plain = rtf_to_text(rtf_text)
    except Exception as exc:  # noqa: BLE001 - striprtf не установлен или битый RTF
        raise ValueError(f"rtf: {exc}") from exc

    # plain — строки с '|' как разделителями ячеек (как в Word-экспорте)
    raw_lines = [ln for ln in plain.splitlines() if ln.strip()]

    rows: List[List[Any]] = []
    for line in raw_lines:
        if "|" in line:
            # разделить по '|' — это колонки таблицы
            cells = [c.strip() or None for c in line.split("|")]
            # убрать крайние пустые (строка начинается и кончается с |)
            while cells and cells[0] is None:
                cells.pop(0)
            while cells and cells[-1] is None:
                cells.pop()
            # пустые итоговые строки пропускаем? Нет, «ИТОГО» нужен
            if any(v is not None and str(v).strip() for v in cells):
                rows.append(cells)
        else:
            stripped = line.strip()
            if not stripped:
                continue
            # служебные строки (заголовки счёта) — кладём как одну ячейку, чтобы layout их пропустил,
            # но meta нашла поставщика
            # Если это похоже на заголовок документа (СЧЕТ-ФАКТУРА, Поставщик, Покупатель) — оставляем
            if any(kw in stripped for kw in ("СЧЕТ-ФАКТУРА", "Поставщик:", "Покупатель:", "УНП", "IBAN", "Всего к оплате", "ИТОГО")):
                rows.append([stripped])

    if not rows:
        raise ValueError("в RTF не найдено таблиц")

    # --- Склейка раздробленного заголовка МАВ ---
    # В RTF от МАВ заголовок таблицы разбит на 5 физических строк из-за переносов внутри ячеек:
    #   ['№','Наименование товара','Инвент.'] + ['Номер','Един.'] + ['Изм.','Количество','Цена без НДС,'] + ...
    # Детектим этот паттерн и заменяем на одну каноническую строку из 9 колонок.
    # Признак: подряд идут строки где есть 'Инвент.'/'Номер'/'Изм.'/'бел. руб' и следующая строка начинается с '1' (данные)
    canonical_header = [
        "№",
        "Наименование товара",
        "Инвент. Номер",
        "Един. Изм.",
        "Количество",
        "Цена без НДС, бел. руб",
        "Стоимость без НДС, бел. руб",
        "Сумма НДС по ставке 20%, бел. руб",
        "Стоимость с НДС, бел. руб",
    ]
    # ищем фрагмент заголовка
    frag_start = None
    frag_end = None
    for idx, row in enumerate(rows):
        first = str(row[0] or "").strip() if row else ""
        # начало фрагмента — строка с '№' и 'Наименование'
        if first == "№" and any("Наименование" in str(c or "") for c in row):
            frag_start = idx
        # конец фрагмента — строки с 'бел. руб' перед данными
        if frag_start is not None and first == "1" and len(row) == 9:
            # нашли первую строку данных — фрагмент заканчивается перед ней
            frag_end = idx
            break
    if frag_start is not None and frag_end is not None and frag_end - frag_start >= 3:
        # проверяем что между ними были фрагменты с 'Инвент.'/'Номер'/'Изм.'
        middle_text = " ".join(" ".join(str(c or "") for c in rows[i]) for i in range(frag_start, frag_end))
        if "Инвент." in middle_text and "Един." in middle_text:
            # заменяем фрагмент на одну строку
            rows = rows[:frag_start] + [canonical_header] + rows[frag_end:]

    return [Sheet(name="RTF", rows=rows)]


# --------------------------------------------------------------------------- #
#  PDF (Элсви)
# --------------------------------------------------------------------------- #
def _read_pdf(data: bytes) -> List[Sheet]:
    """Прочитать PDF с таблицами (Элсви). Использует pdfplumber."""

    try:
        import pdfplumber
    except ImportError as exc:
        raise ValueError("pdf: нужен pdfplumber (pip install pdfplumber)") from exc

    rows: List[List[Any]] = []
    texts: List[str] = []

    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                # сначала пробуем таблицы
                tables = page.extract_tables() or []
                for tbl in tables:
                    for raw_row in tbl:
                        if not raw_row:
                            continue
                        cells = [str(c).strip().replace("\n", " ") if c is not None else None for c in raw_row]
                        # очистить «2 160,00» -> оставить как есть, парсинг чисел делает extract позже
                        while cells and cells[-1] is None:
                            cells.pop()
                        if any(v is not None and str(v).strip() for v in cells):
                            rows.append(cells)
                # также текст страницы для meta (поставщик, покупатель)
                t = page.extract_text() or ""
                if t.strip():
                    for ln in t.splitlines():
                        if ln.strip():
                            texts.append(ln.strip())
    except Exception as exc:  # noqa: BLE001 - битый PDF
        raise ValueError(f"pdf: {exc}") from exc

    if not rows:
        # Фолбэк: попробовать разобрать текст построчно как псевдотаблицу
        if texts:
            for ln in texts:
                if any(kw in ln for kw in ("СЧЕТ", "Продавец", "Покупатель", "Итого", "Сумма НДС")):
                    rows.append([ln])
        if not rows:
            raise ValueError("в PDF не найдено распознаваемых таблиц; если это скан, нужен PDF с текстом или Excel")

    sheet = Sheet(name="PDF", rows=rows)
    return [sheet]


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
    elif kind == "rtf":
        attempts = [("rtf", _read_rtf), ("html", _read_html), ("xlsx", _read_xlsx), ("xls", _read_xls), ("csv", _read_csv)]
    elif kind == "pdf":
        attempts = [("pdf", _read_pdf)]
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
            if name == "xlsx":
                # реквизиты могут быть нарисованы текстовым полем поверх листа
                sf.texts = xlsx_drawing_texts(data)
            return sf
        errors.append(f"{name}: пустой результат")

    raise ValueError("не удалось прочитать файл: " + "; ".join(errors))


def iter_source_files(paths) -> List[SourceFile]:
    out: List[SourceFile] = []
    for p in paths:
        out.append(read_source(p))
    return out
