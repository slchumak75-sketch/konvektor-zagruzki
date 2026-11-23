"""Низкоуровневое распознавание отдельных ячеек скана.

Результат — только подсказка для человека, а не подтверждённое значение:
даже очень уверенное OCR может ошибиться в одной цифре. Поддерживается только
экспериментальное распознавание стандартной 11-колоночной формы ТТН.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Protocol

from PIL import Image, ImageOps


@dataclass(frozen=True)
class CellCandidate:
    text: str
    confidence: int
    # Не снимать этот флаг на основании уверенности OCR или математики.
    requires_confirmation: bool = True


class CellEngine(Protocol):
    def recognize(self, image: Image.Image, numeric: bool) -> tuple[str, int]: ...


def prepare_cell(image: Image.Image, threshold: int = 110) -> Image.Image:
    """Уменьшить влияние цветной сетки, увеличить и добавить белые поля."""
    if image.width < 1 or image.height < 1:
        raise ValueError("пустое изображение ячейки")
    gray = ImageOps.grayscale(image)
    binary = gray.point(lambda value: 0 if value < threshold else 255, mode="1").convert("L")
    binary = binary.resize((binary.width * 3, binary.height * 3), Image.Resampling.LANCZOS)
    return ImageOps.expand(binary, border=25, fill="white")


def recognize_cell(image: Image.Image, engine: CellEngine, *, numeric: bool = False) -> CellCandidate:
    """Предложить прочитанное значение, а не подтверждать его автоматически."""
    text, confidence = engine.recognize(prepare_cell(image), numeric)
    text = " ".join(text.split()).strip()
    if numeric:
        text = text.replace(" ", "")
        # Сомнительный набор символов не должен превращаться в число.
        if not re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?%?", text):
            text = ""
    return CellCandidate(text=text or "???", confidence=max(0, min(100, int(confidence))))


class TesseractCellEngine:
    """Tesseract с заранее установленными моделями rus и eng."""

    def __init__(self, tessdata: str):
        try:
            import tesserocr
        except ImportError as exc:
            raise RuntimeError("Для OCR требуется tesserocr и языковые модели Tesseract") from exc
        self._tesserocr = tesserocr
        self._api = tesserocr.PyTessBaseAPI(path=tessdata, lang="rus+eng")

    def recognize(self, image: Image.Image, numeric: bool) -> tuple[str, int]:
        self._api.SetVariable("tessedit_char_whitelist", "0123456789,.%+-" if numeric else "")
        self._api.SetPageSegMode(self._tesserocr.PSM.SINGLE_LINE if numeric else self._tesserocr.PSM.SINGLE_BLOCK)
        self._api.SetImage(image)
        text = self._api.GetUTF8Text() or ""
        return text, self._api.MeanTextConf()

    def close(self) -> None:
        self._api.End()

    def __enter__(self):
        return self

    def __exit__(self, *_args) -> None:
        self.close()


@dataclass(frozen=True)
class GridLines:
    """Координаты линий в изображении; наличие сетки не доказывает её полноту."""
    vertical: tuple[int, ...]
    horizontal: tuple[int, ...]

    def cell(self, row: int, column: int, inset: int = 2) -> tuple[int, int, int, int]:
        x1, x2 = self.vertical[column:column + 2]
        y1, y2 = self.horizontal[row:row + 2]
        if x2 - x1 <= 2 * inset or y2 - y1 <= 2 * inset:
            raise ValueError("ячейка слишком мала")
        return x1 + inset, y1 + inset, x2 - inset, y2 - inset


def _peaks(counts, minimum: int, gap: int = 3) -> tuple[int, ...]:
    """Объединить соседние пиксели одной линии; центр берём по максимуму."""
    candidates = [i for i, count in enumerate(counts) if count >= minimum]
    if not candidates:
        return ()
    groups = [[candidates[0]]]
    for pos in candidates[1:]:
        if pos - groups[-1][-1] <= gap:
            groups[-1].append(pos)
        else:
            groups.append([pos])
    return tuple(max(group, key=lambda i: counts[i]) for group in groups)


def detect_grid(image: Image.Image, bounds: tuple[int, int, int, int]) -> GridLines:
    """Найти линии внутри заранее определённой области таблицы.

    Область должна быть предварительно выровнена по горизонтали; функция не
    определяет заголовок и не гарантирует, что все строки накладной найдены.
    Не использовать результат как подтверждение полноты документа.
    """
    try:
        import cv2
        import numpy as np
    except ImportError as exc:
        raise RuntimeError("Для поиска линий OCR нужны numpy и opencv-python-headless") from exc
    left, top, right, bottom = bounds
    if not (0 <= left < right <= image.width and 0 <= top < bottom <= image.height):
        raise ValueError("неверные границы области таблицы")
    gray = np.asarray(ImageOps.grayscale(image.crop(bounds)))
    height, width = gray.shape
    if width < 100 or height < 50:
        raise ValueError("область таблицы слишком мала")
    thresholded = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                        cv2.THRESH_BINARY_INV, 31, 12)
    horizontal = cv2.morphologyEx(
        thresholded, cv2.MORPH_OPEN,
        cv2.getStructuringElement(cv2.MORPH_RECT, (max(30, width // 14), 1)))
    vertical = cv2.morphologyEx(
        thresholded, cv2.MORPH_OPEN,
        cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(25, height // 30))))
    y_counts = (horizontal > 0).sum(axis=1)
    x_counts = (vertical > 0).sum(axis=0)
    ys = tuple(top + y for y in _peaks(y_counts, int(width * .32), gap=4))
    xs = tuple(left + x for x in _peaks(x_counts, int(height * .22), gap=5))
    if len(xs) < 2 or len(ys) < 2:
        raise ValueError("не удалось достаточно уверенно выделить линии таблицы")
    return GridLines(vertical=xs, horizontal=ys)


def deskew_page(image: Image.Image) -> tuple[Image.Image, float]:
    """Выровнять длинные горизонтальные линии скана; вернуть угол их наклона.

    Если для оценки недостаточно линий, явно отказываемся вместо угадывания.
    """
    try:
        import cv2
        import numpy as np
    except ImportError as exc:
        raise RuntimeError("Для выравнивания OCR нужны numpy и opencv-python-headless") from exc
    import math

    rgb = np.asarray(image.convert("RGB"))
    height, width = rgb.shape[:2]
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV, 31, 12)
    lines = cv2.HoughLinesP(binary, 1, np.pi / 1800,
                            threshold=max(40, width // 9),
                            minLineLength=max(80, width // 3),
                            maxLineGap=max(10, width // 40))
    angles = []
    if lines is not None:
        for x1, y1, x2, y2 in lines.reshape(-1, 4):
            angle = math.degrees(math.atan2(float(y2-y1), float(x2-x1)))
            if -5 <= angle <= 5:
                angles.append(angle)
    if len(angles) < 3:
        raise ValueError("не удалось надёжно определить наклон скана")
    angle = float(np.median(angles))
    # Расширяем холст, чтобы не обрезать содержимое по углам.
    matrix = cv2.getRotationMatrix2D((width / 2, height / 2), angle, 1)
    cos, sin = abs(matrix[0, 0]), abs(matrix[0, 1])
    new_width = int(math.ceil(height * sin + width * cos))
    new_height = int(math.ceil(height * cos + width * sin))
    matrix[0, 2] += (new_width - width) / 2
    matrix[1, 2] += (new_height - height) / 2
    rotated = cv2.warpAffine(rgb, matrix, (new_width, new_height),
                             flags=cv2.INTER_CUBIC, borderValue=(255, 255, 255))
    return Image.fromarray(rotated), angle


@dataclass(frozen=True)
class TableRegion:
    """Область кандидатной товарной таблицы, без гарантии полноты."""
    bounds: tuple[int, int, int, int]
    lines: GridLines


def find_ttn_region(image: Image.Image) -> TableRegion:
    """Предложить область товарного раздела стандартной 11-колоночной ТТН.

    Это экспериментальное обнаружение геометрии, не проверка полноты и не
    универсальный детектор накладных. Иной формат должен быть отклонён, а не
    принят за ТТН. На изображении предварительно вызывают deskew_page.
    """
    try:
        import cv2
        import numpy as np
    except ImportError as exc:
        raise RuntimeError("Для поиска таблиц OCR нужны numpy и opencv-python-headless") from exc
    gray = np.asarray(ImageOps.grayscale(image))
    height, width = gray.shape
    if width < 400 or height < 300:
        raise ValueError("страница слишком мала для поиска таблицы")
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV, 31, 12)
    v = cv2.morphologyEx(binary, cv2.MORPH_OPEN,
                          cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(25, height // 50))))
    h = cv2.morphologyEx(binary, cv2.MORPH_OPEN,
                          cv2.getStructuringElement(cv2.MORPH_RECT, (max(40, width // 16), 1)))
    # Формат ТТН: 11 колонок; проверяются относительные позиции линий, а не
    # фиксированное положение таблицы на листе или имя поставщика.
    template = (0, .163, .223, .306, .392, .476, .531, .613, .695, .780, .862, 1)
    window = max(150, int(height * .16))
    step = max(40, window // 3)
    candidates = []
    for top in range(0, height, step):
        bottom = min(height, top + window)
        if bottom - top < window // 2:
            continue
        xcounts = (v[top:bottom] > 0).sum(axis=0)
        xs = _peaks(xcounts, max(25, int((bottom-top) * .26)), gap=6)
        for left in xs:
            for right in xs:
                span = right - left
                if span < width * .55:
                    continue
                matched = []
                for fraction in template:
                    target = left + span * fraction
                    nearby = [x for x in xs if abs(x - target) <= span * .013]
                    if not nearby:
                        break
                    matched.append(min(nearby, key=lambda x: abs(x-target)))
                if len(matched) == len(template) and len(set(matched)) == len(template):
                    candidates.append((top, bottom, tuple(matched)))
    if not candidates:
        raise ValueError("не найдена товарная таблица поддерживаемой формы ТТН")
    # Соседние окна одной таблицы дают почти одинаковые координаты колонок.
    clusters = []
    for candidate in candidates:
        top, bottom, xs = candidate
        for cluster in clusters:
            anchor = cluster[0][2]
            if abs(xs[0] - anchor[0]) < width * .02 and abs(xs[-1] - anchor[-1]) < width * .02 and all(abs(a-b) < width*.015 for a,b in zip(xs, anchor)):
                cluster.append(candidate)
                break
        else:
            clusters.append([candidate])
    cluster = max(clusters, key=len)
    xs = tuple(int(np.median([entry[2][i] for entry in cluster])) for i in range(len(template)))
    left, right = xs[0], xs[-1]
    ycounts = (h[:, left:right+1] > 0).sum(axis=1)
    ys = _peaks(ycounts, int((right-left)*.34), gap=4)
    valid = []
    for y in ys:
        # Около границы строки должно быть достаточно линий тех же колонок.
        band = v[max(0, y-18):min(height, y+19)]
        matches = sum(bool(band[:, max(0, x-7):min(width, x+8)].any()) for x in xs)
        if matches >= 9:
            valid.append(y)
    if len(valid) < 3:
        raise ValueError("недостаточно границ строк товарной таблицы")
    # Выбираем непрерывную последовательность; другие табличные разделы не
    # должны попадать в товарную часть.
    groups = [[valid[0]]]
    for y in valid[1:]:
        if y - groups[-1][-1] <= height * .15:
            groups[-1].append(y)
        else:
            groups.append([y])
    ys = max(groups, key=lambda group: (len(group), group[-1]-group[0]))
    if len(ys) < 3:
        raise ValueError("не удалось определить непрерывную товарную таблицу")
    return TableRegion((left, ys[0], right, ys[-1]), GridLines(xs, tuple(ys)))


@dataclass(frozen=True)
class TableRowCandidate:
    """Поля строки и их OCR-подсказки, все требуют сверки с оригиналом."""
    kind: str  # item | total | header
    cells: dict[str, CellCandidate]
    top: int
    bottom: int


def read_ttn_rows(image: Image.Image, engine: CellEngine) -> list[TableRowCandidate]:
    """Предложить строки ТТН из уже выровненной страницы.

    Назначение коротких строк заголовка и итога предварительное; полноту и
    правильность значений должен подтвердить пользователь. До подключения к
    интерфейсу этот результат нельзя использовать как готовую накладную.
    """
    from difflib import SequenceMatcher

    grid = find_ttn_region(image).lines
    fields = (("name", 0), ("unit", 1), ("qty", 2), ("price", 3),
              ("total_sum", 4), ("vat_rate", 5), ("vat_sum", 6), ("sum_with_vat", 7))
    out = []
    names = []
    for row in range(len(grid.horizontal) - 1):
        name = recognize_cell(image.crop(grid.cell(row, 0, inset=3)), engine)
        names.append(name)
    for row, name in enumerate(names):
        normalized = name.text.casefold().translate(str.maketrans({"a": "а", "e": "е", "o": "о", "t": "т", "r": "г", "i": "и"}))
        text = re.sub(r"[^а-яё]", "", normalized)
        y1, y2 = grid.horizontal[row:row + 2]
        # Только первый и последний фрагменты могут быть шапкой/итогом.
        is_header = row == 0 and ("наимен" in text or "наимено" in text)
        numbered_header = row == 1 and out and out[0].kind == "header" and bool(re.fullmatch(r"\W*\d+\W*", name.text))
        is_total = row == len(names)-1 and len(text) <= 8 and SequenceMatcher(None, text, "итого").ratio() >= .65
        kind = "header" if is_header or numbered_header else ("total" if is_total else "item")
        cells = {"name": name}
        if kind != "header":
            for field, column in fields[1:]:
                if kind == "total" and field not in ("qty", "total_sum", "vat_sum", "sum_with_vat"):
                    continue
                box = grid.cell(row, column, inset=3)
                if kind == "item":
                    # Числа и единица измерения напечатаны в первой строке ячейки.
                    box = (box[0], box[1], box[2], min(box[3], box[1] + max(30, image.width // 38)))
                cells[field] = recognize_cell(image.crop(box), engine, numeric=field != "unit")
        out.append(TableRowCandidate(kind, cells, y1, y2))
    return out


def candidates_to_extraction(source: str, sheet: str, rows: list[TableRowCandidate]):
    """Преобразовать OCR в неподтверждённые позиции и итог, без домыслов.

    Сама сегментация может пропустить строку: перед сохранением интерфейс
    обязан потребовать от пользователя подтвердить полноту таблицы.
    """
    from .extract import DocTotal, Item, SheetExtraction
    from .textutil import parse_number, find_barcode_in_text

    extraction = SheetExtraction()
    numeric = {"qty", "price", "total_sum", "vat_rate", "vat_sum", "sum_with_vat"}
    total_fields = {"qty": "qty", "total_sum": "sum", "vat_sum": "vat_sum", "sum_with_vat": "total"}
    for row_number, row in enumerate(rows, 1):
        if row.kind == "header":
            continue
        if row.kind == "total":
            values = {}
            for name, key in total_fields.items():
                candidate = row.cells.get(name)
                value = parse_number(candidate.text) if candidate else None
                if value is not None:
                    values[key] = value
            extraction.totals.append(DocTotal(source=source, sheet=sheet, row=row_number,
                                               label="Итого (OCR)", values=values,
                                               needs_confirmation=True))
            continue
        item = Item(source=source, sheet=sheet, row=row_number, barcode_from_name_allowed=False)
        for name, candidate in row.cells.items():
            if name not in numeric | {"name", "unit"}:
                continue
            value = parse_number(candidate.text) if name in numeric else candidate.text
            if value == "???" or value is None:
                item.unrecognized.add(name)
                value = "???" if name == "name" else None
            setattr(item, name, value)
            item.needs_confirmation.add(name)
        # Штрихкод из OCR — только кандидат: его обязательно сверяют отдельно.
        barcode = find_barcode_in_text(item.name or "")
        if barcode:
            item.barcode = barcode
        else:
            item.unrecognized.add("barcode")
        item.needs_confirmation.add("barcode")
        extraction.items.append(item)
    extraction.warnings.append("OCR: подтвердите все значения и полноту строк по оригиналу")
    return extraction


def read_scanned_pdf(path: str, tessdata: str):
    """Прочитать все страницы поддерживаемого скана; частичный разбор запрещён."""
    import pdfplumber
    from .extract import SheetExtraction
    from .readers import Sheet, SourceFile

    extraction = SheetExtraction()
    source = SourceFile(path=path, reader="pdf-ocr")
    with pdfplumber.open(path) as pdf, TesseractCellEngine(tessdata) as engine:
        if not pdf.pages or len(pdf.pages) > 30:
            raise ValueError("OCR: документ пустой или содержит более 30 страниц")
        for page_number, page in enumerate(pdf.pages, 1):
            # Не принимать логотип на текстовой странице за скан накладной.
            page_area = page.width * page.height
            if not any(abs(im.get("x1", 0) - im.get("x0", 0)) * abs(im.get("bottom", 0) - im.get("top", 0)) > page_area * .3 for im in page.images):
                raise ValueError(f"OCR: страница {page_number} не является полностраничным сканом")
            image = page.to_image(resolution=200).original
            image, _angle = deskew_page(image)
            rows = read_ttn_rows(image, engine)
            sheet = f"PDF {page_number}"
            page_extraction = candidates_to_extraction(path, sheet, rows)
            if not page_extraction.items:
                raise ValueError(f"OCR: на странице {page_number} не найдены товары")
            extraction.items.extend(page_extraction.items)
            extraction.totals.extend(page_extraction.totals)
            extraction.warnings.extend(page_extraction.warnings)
            source.sheets.append(Sheet(name=sheet))
    return source, extraction
