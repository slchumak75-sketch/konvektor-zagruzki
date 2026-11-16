"""Извлечение реквизитов документа: поставщик, тип, номер и дата.

Нужно, чтобы в итоговой таблице и в имени файла было видно, от кого и какой
документ. Данные берутся из содержимого листа (подписи «Поставщик»,
«Продавец», «Грузоотправитель», «Счет № … от …», «Серия номер накладной»)
и, если в документе ничего нет, — из имени файла.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple

from .readers import Sheet, SourceFile
from .textutil import clean_text, norm_header

#: С какого дня Excel считает даты.
_EXCEL_EPOCH = _dt.date(1899, 12, 30)

_MONTH_STEMS = {
    "январ": 1,
    "феврал": 2,
    "март": 3,
    "апрел": 4,
    "ма": 5,  # май / мая
    "июн": 6,
    "июл": 7,
    "август": 8,
    "сентябр": 9,
    "октябр": 10,
    "ноябр": 11,
    "декабр": 12,
}

_RE_NUM_DATE = re.compile(r"(?<!\d)(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})(?!\d)")
_RE_RU_DATE = re.compile(
    r"(?<!\d)(\d{1,2})\s+(январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр)[а-я]*\.?\s+(\d{4})(?!\d)",
    re.IGNORECASE,
)

#: Подписи, рядом с которыми стоит поставщик.
_SUPPLIER_LABELS = (
    "поставщик",
    "продавец",
    "грузоотправитель",
    "отправитель",
    "продавец (поставщик)",
    "поставщик и его адрес",
)
#: Названия нашей организации (покупателя). Поставщиком они быть не могут,
#: даже если стоят в самом верху документа: в счетах сверху часто печатают
#: именно получателя/плательщика. Свой список можно задать переменной
#: окружения KONVEKTOR_OUR_COMPANIES (через запятую).
DEFAULT_OUR_COMPANIES = ("мирстроймаркет", "мирсититорг")

#: Надёжные подписи получателя — кто получает товар, тот и покупатель.
_BUYER_STRONG_LABELS = (
    "покупатель",
    "грузополучатель",
    "конечный получатель",
    "получатель",
)
#: Слабые подписи: плательщиком/заказчиком перевозки бывает и поставщик.
_BUYER_WEAK_LABELS = ("плательщик", "заказчик")
#: Подписи «другой стороны» — их значения поставщиком не считаем.
_COUNTERPARTY_LABELS = _BUYER_STRONG_LABELS + _BUYER_WEAK_LABELS
#: Подписи с номером документа и датой.
_NUMBER_LABELS = (
    "серия номер накладной",
    "номер накладной",
    "номер счета",
    "номер документа",
    "номер ттн",
    "документ №",
    "счет №",
)
_DATE_LABELS = ("дата", "дата документа", "дата составления")

_LEGAL_FORMS = (
    "ООО",
    "ОАО",
    "ЗАО",
    "ПАО",
    "АО",
    "УП",
    "КУП",
    "РУП",
    "ЧУП",
    "ЧТУП",
    "ИП",
    "ЧП",
    "ОДО",
    "СП",
    "ГП",
    "ТОО",
    "ПТ",
    "ООО «",
    "ФХ",
    "КФХ",
)

_RE_NUMBER_IN_TEXT = re.compile(r"№\s*([0-9A-Za-zА-Яа-яЁё][0-9A-Za-zА-Яа-яЁё\-/_\.]{0,24})")
_RE_DOC_WORD_NUMBER = re.compile(
    r"\b(ТТН(?:-1)?|счет|счёт|накладная|приложение|акт)\b[^\d№]{0,20}№?\s*(\d[0-9A-Za-zА-Яа-яЁё\-/_\.]{1,24})",
    re.IGNORECASE,
)
_RE_BETWEEN = re.compile(r"между\s+(.{3,120}?)\s+и\s+(.{3,120})", re.IGNORECASE)
_RE_UNP = re.compile(r"УНП\s*:?\s*(\d{6,12})", re.IGNORECASE)
_RE_ONLY_DIGITS = re.compile(r"^[\d\s]+$")


# --------------------------------------------------------------------------- #
#  Даты
# --------------------------------------------------------------------------- #
def parse_date(value) -> Optional[_dt.date]:
    """Достать дату из строки ('12 ноября 2026', '12.11.26', '19.10.2026')."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.date() if isinstance(value, _dt.datetime) else value
    if isinstance(value, (int, float)):
        return excel_serial_date(value)
    text = clean_text(value)
    if not text:
        return None

    m = _RE_RU_DATE.search(text)
    if m:
        day, stem, year = int(m.group(1)), m.group(2).lower(), int(m.group(3))
        month = None
        for key, num in _MONTH_STEMS.items():
            if stem.startswith(key):
                month = num
                break
        if month:
            try:
                return _dt.date(year, month, day)
            except ValueError:
                return None

    m = _RE_NUM_DATE.search(text)
    if m:
        day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if year < 100:
            year += 2000
        try:
            return _dt.date(year, month, day)
        except ValueError:
            return None
    return None


def excel_serial_date(value: float) -> Optional[_dt.date]:
    """Число Excel (46316) -> дата. Отсекаем всё, что не похоже на дату."""
    try:
        num = float(value)
    except (TypeError, ValueError):
        return None
    if not (20000 <= num <= 80000):
        return None
    return _EXCEL_EPOCH + _dt.timedelta(days=int(num))


def fmt_date(value: Optional[_dt.date]) -> str:
    return value.strftime("%d.%m.%Y") if value else ""


# --------------------------------------------------------------------------- #
#  Модель
# --------------------------------------------------------------------------- #
@dataclass
class DocMeta:
    """Реквизиты документа, извлечённые из файла."""

    supplier: str = ""
    supplier_unp: str = ""
    buyer: str = ""
    doc_type: str = ""
    doc_number: str = ""
    doc_date: Optional[_dt.date] = None
    origin: str = ""  # «документ» | «имя файла» | пусто
    warnings: List[str] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        if self.warnings is None:
            self.warnings = []

    @property
    def date_text(self) -> str:
        return fmt_date(self.doc_date)

    @property
    def doc_label(self) -> str:
        """Короткая подпись документа: «Счет № 3596616», «ТТН-1 № 1019250»."""
        parts = []
        if self.doc_type:
            parts.append(self.doc_type)
        if self.doc_number:
            parts.append(f"№ {self.doc_number}" if self.doc_type else f"№ {self.doc_number}")
        return " ".join(parts)

    @property
    def found(self) -> bool:
        return bool(self.supplier or self.doc_number or self.doc_date)


# --------------------------------------------------------------------------- #
#  Поставщик
# --------------------------------------------------------------------------- #
def _is_label(text: str) -> bool:
    t = norm_header(text)
    if not t:
        return False
    if any(t.startswith(x) for x in _SUPPLIER_LABELS):
        return True
    if any(t.startswith(x) for x in _COUNTERPARTY_LABELS):
        return True
    if any(t.startswith(x) for x in _NUMBER_LABELS) or any(t.startswith(x) for x in _DATE_LABELS):
        return True
    return False


_RE_COMPANY_QUOTED = re.compile(
    r"^(?:ООО|ОАО|ЗАО|ПАО|АО|УП|КУП|РУП|ЧУП|ЧТУП|ИП|ЧП|ОДО|СП|ТОО|ГП|ФХ)\s*[«\"„']",
    re.IGNORECASE,
)
_RE_COMPANY_BARE = re.compile(
    r"^(?:ООО|ОАО|ЗАО|ПАО|ЧУП|ОДО|ИП|УП)\s+[А-ЯЁA-Z][а-яёa-z]"
)
_RE_ADDRESSISH = re.compile(r"(ул\.|г\.\s?п\.|маг-н|обл\.|\bр-н\b|тел\.|д\.\s?\d)")


def looks_like_company(text: str) -> bool:
    """Похоже ли значение на название организации (а не адрес/УНП/подпись)."""
    t = clean_text(text)
    if len(t) < 4:
        return False
    if not (_RE_COMPANY_QUOTED.match(t) or _RE_COMPANY_BARE.match(t)):
        return False
    head = t[:40].lower()
    # «гп. Мир, ул. Красноармейская…» — это адрес
    return not _RE_ADDRESSISH.search(head)


def clean_supplier(text: str) -> str:
    """Из «ООО "ОМА", УНП 100005713, 223021, Минский р-н…» делаем «ООО "ОМА"»."""
    t = clean_text(text)
    if not t:
        return ""
    t = re.split(r"\s+УНП\b", t, flags=re.IGNORECASE)[0]
    t = t.split(",")[0]
    t = t.strip(" .,;:")
    t = re.sub(r"\s{2,}", " ", t)
    if len(t) > 120:
        t = t[:120].rsplit(" ", 1)[0]
    if not any(ch.isalpha() for ch in t):
        return ""
    return t


def _value_right(row: Sequence, col: int) -> str:
    """Ближайшее справа непустое текстовое значение (не подпись и не число)."""
    for cc in range(col + 1, len(row)):
        value = row[cc]
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            continue  # УНП, коды, суммы — не название организации
        text = clean_text(value)
        if not text or _RE_ONLY_DIGITS.match(text):
            continue
        if _is_label(text):
            continue
        return text
    return ""


def _after_colon(text: str) -> str:
    if ":" not in text:
        return ""
    tail = text.split(":", 1)[1].strip()
    return tail


def our_companies() -> Tuple[str, ...]:
    """Названия нашей организации: по умолчанию заказчик + переменная окружения."""
    raw = os.environ.get("KONVEKTOR_OUR_COMPANIES", "")
    extra = [n.strip().lower() for n in raw.split(",") if n.strip()]
    return tuple(extra) if extra else DEFAULT_OUR_COMPANIES


def core_name(text: str) -> str:
    """«ООО "МирСтройМаркет"» -> «мирстроймаркет» (для сравнения названий)."""
    t = norm_header(text)
    for form in _LEGAL_FORMS:
        t = t.replace(form.lower(), " ")
    return re.sub(r"[^0-9a-zа-яё]+", "", t)


def is_our_company(name: str) -> bool:
    """Это наша организация (покупатель), а не поставщик."""
    core = core_name(name)
    if not core:
        return False
    return any(ours and ours in core for ours in our_companies())


def _team_values(
    sheets: Sequence[Sheet], labels: Sequence[str], max_rows: int = 80
) -> List[Tuple[int, int, str]]:
    """Значения рядом с подписями: (строка, колонка, очищенное значение)."""
    found: List[Tuple[int, int, str]] = []
    for sheet in sheets:
        for r, row in enumerate(sheet.rows[:max_rows]):
            for col, value in enumerate(row):
                if not isinstance(value, str):
                    continue
                t = norm_header(value)
                if not t or not any(t.startswith(x) for x in labels):
                    continue
                if any(t.startswith(x) for x in _COUNTERPARTY_LABELS) and labels is _SUPPLIER_LABELS:
                    continue
                for candidate in (_after_colon(value), _value_right(row, col)):
                    cleaned = clean_supplier(candidate)
                    if cleaned and len(cleaned) > 3 and not _RE_ONLY_DIGITS.match(cleaned):
                        found.append((r, col, cleaned))
                        break
    return found


def _supplier_and_buyer(sheets: Sequence[Sheet], max_rows: int = 80) -> Tuple[str, str]:
    """Поставщик и покупатель (наша организация).

    Правила:
      1. значения под подписями «Поставщик», «Продавец», «Грузоотправитель»;
      2. «между ООО «А» и ООО «Б»» — первое поставщик, второе покупатель;
      3. название организации в верхней части листа (первые 8 строк и 6 столбцов,
         затем до 20-й строки) — с меньшим весом;
      4. наша организация (Мирстроймаркет / Мирсититорг / свой список) и
         покупатель из документа поставщиком не считаются никогда.
    """
    # покупатель: сначала надёжные подписи, из них — наша организация, если есть
    strong = _team_values(sheets, _BUYER_STRONG_LABELS, max_rows)
    weak = _team_values(sheets, _BUYER_WEAK_LABELS, max_rows)
    buyer = ""
    for _row, _col, name in strong:
        if is_our_company(name):
            buyer = name
            break
    if not buyer and strong:
        buyer = strong[0][2]
    if not buyer:
        for _row, _col, name in weak:
            if is_our_company(name):
                buyer = name
                break
    if not buyer and weak:
        buyer = weak[0][2]
    buyer_core = core_name(buyer)

    strong_cores = {core_name(n) for _r, _c, n in strong if len(n) > 3}

    def allowed(name: str) -> bool:
        if not name or len(name) < 4:
            return False
        if is_our_company(name):
            return False
        core = core_name(name)
        if not core:
            return False
        if buyer_core and (core in buyer_core or buyer_core in core):
            return False
        # получатель по надёжной подписи поставщиком не бывает
        if any(core in sc or sc in core for sc in strong_cores if sc):
            return False
        return True

    candidates: List[Tuple[float, str]] = []

    # помеченные подписи поставщика — самый надёжный источник
    for row, _col, name in _team_values(sheets, _SUPPLIER_LABELS, max_rows):
        candidates.append((100.0 - row, name))

    # «между А и Б»
    for sheet in sheets:
        for value in sheet.rows[:max_rows]:
            for cell in value:
                if not isinstance(cell, str):
                    continue
                m = _RE_BETWEEN.search(cell)
                if m:
                    cand = clean_supplier(m.group(1))
                    if looks_like_company(cand) and allowed(cand):
                        candidates.append((90.0, cand))
                    other = clean_supplier(m.group(2))
                    if not buyer and looks_like_company(other):
                        buyer, buyer_core = other, core_name(other)

    # название организации в верхней части листа
    for sheet in sheets:
        for r, row in enumerate(sheet.rows[:20]):
            for cell in row[:6]:
                if not isinstance(cell, str) or _is_label(cell):
                    continue
                if not looks_like_company(cell):
                    continue
                name = clean_supplier(cell)
                if len(name) > 4:
                    candidates.append(((60.0 - 2 * r) if r < 8 else (40.0 - r), name))

    for _score, name in sorted(candidates, key=lambda c: -c[0]):
        if allowed(name):
            return name, buyer
    return "", buyer


def _label_inline(text: str, labels: Sequence[str]) -> str:
    """Значение сразу после подписи: «ПОСТАВЩИК: ООО …» -> «ООО …»."""
    low = norm_header(text)
    for label in labels:
        if not low.startswith(label):
            continue
        rest = text[len(label) :].lstrip(" .:-\t")
        if rest.startswith(":"):
            rest = rest.lstrip(":\t ")
        return rest.strip()
    return ""


def _supplier_from_texts(texts: Sequence[str]) -> Tuple[str, str, str]:
    """Поставщик, наш покупатель и УНП из надписей (текстовые поля поверх листа).

    Так устроены счета из 1С: реквизиты нарисованы текстовым полем, в ячейках
    их нет. Правила те же, что и для ячеек: подпись «Поставщик/Продавец/
    Отправитель» и название организации сразу после неё; наша организация
    поставщиком не бывает.
    """
    supplier, buyer, unp = "", "", ""
    for text in texts:
        if not text:
            continue
        rest = _label_inline(text, _SUPPLIER_LABELS)
        if rest and not supplier:
            name = clean_supplier(rest)
            if name and looks_like_company(name) and not is_our_company(name):
                supplier = name
                m = _RE_UNP.search(text)
                if m:
                    unp = m.group(1)
        rest = _label_inline(text, _BUYER_STRONG_LABELS + _BUYER_WEAK_LABELS)
        if rest and not buyer:
            name = clean_supplier(rest)
            if name and looks_like_company(name) and is_our_company(name):
                buyer = name
    return supplier, buyer, unp


def _supplier_from_sheets(sheets: Sequence[Sheet], max_rows: int = 80) -> str:
    # 1. по подписям
    for sheet in sheets:
        for row in sheet.rows[:max_rows]:
            for col, value in enumerate(row):
                if not isinstance(value, str):
                    continue
                t = norm_header(value)
                if not t or not any(t.startswith(x) for x in _SUPPLIER_LABELS):
                    continue
                if any(t.startswith(x) for x in _COUNTERPARTY_LABELS):
                    continue
                inline = _after_colon(value)
                if inline and not _RE_ONLY_DIGITS.match(inline):
                    cleaned = clean_supplier(inline)
                    if cleaned and len(cleaned) > 3:
                        return cleaned
                right = _value_right(row, col)
                if right:
                    cleaned = clean_supplier(right)
                    if cleaned and len(cleaned) > 3:
                        return cleaned

    # 2. «между ООО "А" и ООО "Б"» — поставщик первый
    for sheet in sheets:
        for row in sheet.rows[:max_rows]:
            for value in row:
                if not isinstance(value, str):
                    continue
                m = _RE_BETWEEN.search(value)
                if m:
                    cleaned = clean_supplier(m.group(1))
                    if cleaned and looks_like_company(cleaned):
                        return cleaned

    # 3. название организации в самом верху листа (шапка выгрузки 1С)
    for sheet in sheets:
        for row in sheet.rows[:6]:
            for value in row[:4]:
                if not isinstance(value, str):
                    continue
                if looks_like_company(value) and not _is_label(value):
                    cleaned = clean_supplier(value)
                    if cleaned and len(cleaned) > 4:
                        return cleaned
    return ""


# --------------------------------------------------------------------------- #
#  Номер и дата документа
# --------------------------------------------------------------------------- #
_RE_T_PRILOZH = re.compile(r"приложение\s+к\s+товарно", re.IGNORECASE)
_RE_T_TTN1 = re.compile(r"\bттн-1\b", re.IGNORECASE)
_RE_T_TTN = re.compile(r"\bттн\b|товарно-транспортн", re.IGNORECASE)
_RE_T_SF = re.compile(r"\bсч[её]т-фактура\b", re.IGNORECASE)
_RE_T_SCHET = re.compile(r"\bсч[её]т\b", re.IGNORECASE)
_RE_T_NAKL = re.compile(r"\bнакладная\b|\bнакладной\b", re.IGNORECASE)
_RE_T_AKT = re.compile(r"\bакт\b", re.IGNORECASE)


def _doc_type_from(text: str) -> str:
    """Тип документа по тексту. Слова ищем целиком: «компакт» — не «акт»."""
    t = clean_text(text)
    if not t:
        return ""
    if _RE_T_PRILOZH.search(t):
        return "Приложение к ТТН"
    if _RE_T_TTN1.search(t):
        return "ТТН-1"
    if _RE_T_TTN.search(t):
        return "ТТН"
    if _RE_T_SF.search(t):
        return "Счет-фактура"
    if _RE_T_SCHET.search(t):
        return "Счет"
    if _RE_T_NAKL.search(t):
        return "Накладная"
    if _RE_T_AKT.search(t):
        return "Акт"
    return ""


def _date_near(text: str) -> Optional[_dt.date]:
    return parse_date(text)


_RE_OT = re.compile(r"^от\b", re.IGNORECASE)
_RE_CONTRACT_WORD = re.compile(r"договор|контракт|спецификаци|заявк|соглашени", re.IGNORECASE)


def _number_date_from_texts(texts: Sequence[str]) -> Tuple[str, str, Optional[_dt.date], str]:
    """Ищем «… № N от дата» в текстах документа.

    Кандидаты оцениваются: документ со «Счет/ТТН/Накладная» и «от» — хороший,
    «Договор № 528 от 31.05.2016» — плохой (это не номер накладной).
    """
    best: Optional[Tuple[int, str, str, Optional[_dt.date]]] = None

    def consider(score: int, number: str, doc_type: str, date: Optional[_dt.date]) -> None:
        nonlocal best
        if best is None or score > best[0]:
            best = (score, number, doc_type, date)

    for index, text in enumerate(texts):
        for m in _RE_NUMBER_IN_TEXT.finditer(text):
            number = m.group(1).strip(" .,;:-/")
            if not number or not any(ch.isdigit() for ch in number):
                continue
            doc_type = _doc_type_from(text)
            tail = text[m.end() : m.end() + 40]
            score = 0
            if doc_type:
                score += 3
            if _RE_OT.match(tail.strip()):
                score += 1
            if _RE_CONTRACT_WORD.search(text):
                score -= 5
            date = _date_near(tail) or _date_near(text)
            if date is None:
                # «серия ЯЛ № 0066065», а «от 12.11.26» — в соседней ячейке
                for nxt in texts[index + 1 : index + 3]:
                    date = _date_near(nxt)
                    if date is not None:
                        score += 1
                        break
            if date is not None:
                score += 1
            consider(score, number, doc_type, date)

    # «ТТН 2880985 от 15.10.26» — без знака №
    for text in texts:
        m = _RE_DOC_WORD_NUMBER.search(text)
        if not m:
            continue
        number = m.group(2).strip(" .,;:-/")
        if not any(ch.isdigit() for ch in number):
            continue
        doc_type = _doc_type_from(text) or _doc_type_from(m.group(1))
        tail = text[m.end() : m.end() + 40]
        score = 3 if doc_type else 0
        if _RE_OT.match(tail.strip()):
            score += 1
        if _RE_CONTRACT_WORD.search(text):
            score -= 5
        date = _date_near(tail) or _date_near(text)
        if date is not None:
            score += 1
        consider(score, number, doc_type, date)

    if best is not None and best[0] > 0:
        _, number, doc_type, date = best
        return number, doc_type, date, "документ"
    return "", "", None, ""


def _label_values(sheets: Sequence[Sheet], max_rows: int = 80) -> Tuple[str, Optional[_dt.date]]:
    """Значения подписей «Серия номер накладной» и «Дата»."""
    number = ""
    date: Optional[_dt.date] = None
    for sheet in sheets:
        for row in sheet.rows[:max_rows]:
            for col, value in enumerate(row):
                if not isinstance(value, str):
                    continue
                t = norm_header(value)
                if not t:
                    continue
                if not number and any(t.startswith(x) for x in _NUMBER_LABELS):
                    inline = _after_colon(value)
                    candidate = inline or _value_right(row, col)
                    candidate = clean_text(candidate)
                    if candidate and any(ch.isdigit() for ch in candidate):
                        number = candidate.split()[0].strip(" .,;:")
                if date is None and any(t == x or t.startswith(x + " ") or t.startswith(x + ":") for x in _DATE_LABELS):
                    # дата может стоять справа (в т.ч. числом Excel)
                    for cc in range(col + 1, min(len(row), col + 4)):
                        cell = row[cc]
                        if cell is None:
                            continue
                        parsed = parse_date(cell)
                        if parsed:
                            date = parsed
                            break
                    if date is None:
                        date = parse_date(_after_colon(value))
                if number and date:
                    return number, date
    return number, date


def from_filename(path: str) -> Tuple[str, Optional[_dt.date], str]:
    """Резервный источник: номер и дата из имени файла."""
    stem = os.path.splitext(os.path.basename(path))[0]
    if not stem:
        return "", None, ""

    date = parse_date(stem)
    number = ""
    m = _RE_NUMBER_IN_TEXT.search(stem)
    if m and any(ch.isdigit() for ch in m.group(1)):
        number = m.group(1).strip(" .,;:-/")
    if not number:
        m = _RE_DOC_WORD_NUMBER.search(stem)
        if m and any(ch.isdigit() for ch in m.group(2)):
            number = m.group(2).strip(" .,;:-/")
    if not number:
        m = re.search(r"([А-ЯA-Z]{2,}[-\s]?\d{4,})", stem)
        if m:
            number = m.group(1).replace(" ", "-")
    if not number:
        m = re.search(r"(\d{5,})", stem)
        if m:
            number = m.group(1)
    doc_type = _doc_type_from(stem)
    return number, date, doc_type


# --------------------------------------------------------------------------- #
#  Публичный API
# --------------------------------------------------------------------------- #
def extract_meta(source: SourceFile) -> DocMeta:
    """Собрать реквизиты документа из прочитанных листов."""
    meta = DocMeta()
    sheets = source.sheets
    if not sheets:
        return meta

    texts: List[str] = []
    for sheet in sheets[:3]:
        for row in sheet.rows[:80]:
            for value in row:
                if isinstance(value, str):
                    text = clean_text(value)
                    if text:
                        texts.append(text)

    # надписи поверх листа (текстовые поля): там же часто лежат реквизиты
    extra = [clean_text(t) for t in (source.texts or []) if clean_text(t)]

    meta.supplier, meta.buyer = _supplier_and_buyer(sheets)
    if not meta.supplier and extra:
        drawn_supplier, drawn_buyer, drawn_unp = _supplier_from_texts(extra)
        if drawn_supplier:
            meta.supplier = drawn_supplier
            if drawn_unp:
                meta.supplier_unp = drawn_unp
        if not meta.buyer and drawn_buyer:
            meta.buyer = drawn_buyer

    for text in texts:
        m = _RE_UNP.search(text)
        if m:
            meta.supplier_unp = meta.supplier_unp or m.group(1)
            break

    number, doc_type, date, _ = _number_date_from_texts(texts)
    if not number or date is None:
        # номер и дата бывают только в надписи поверх листа
        drawn_number, drawn_type, drawn_date, _ = _number_date_from_texts(extra)
        if not number and drawn_number:
            number = drawn_number
        if date is None and drawn_date is not None:
            date = drawn_date
        if not doc_type and drawn_type:
            doc_type = drawn_type
    label_number, label_date = _label_values(sheets)
    from_doc = bool(number or date is not None or label_number)

    if date is None and label_date is None:
        # дата без номера: берём первую дату в документе
        for text in texts:
            found = _date_near(text)
            if found is not None:
                date, from_doc = found, True
                break
    if not number and label_number:
        number, from_doc = label_number, True
    if date is None and label_date is not None:
        date, from_doc = label_date, True
    if not doc_type:
        for text in texts:
            doc_type = _doc_type_from(text)
            if doc_type:
                break

    if not number or date is None:
        fname_number, fname_date, fname_type = from_filename(source.path)
        if not number and fname_number:
            number = fname_number
        if date is None and fname_date is not None:
            date = fname_date
        if not doc_type and fname_type:
            doc_type = fname_type

    meta.doc_number = number
    meta.doc_type = doc_type
    meta.doc_date = date
    meta.origin = "документ" if from_doc else ("имя файла" if (number or date or doc_type) else "")

    if not meta.supplier:
        if meta.buyer and is_our_company(meta.buyer):
            meta.warnings.append(
                "поставщик не найден: в документе указана только наша организация "
                f"({meta.buyer})"
            )
        else:
            meta.warnings.append("поставщик не найден в документе")
    if meta.doc_date is None:
        meta.warnings.append("дата документа не найдена")
    return meta


def meta_from_path(path: str) -> DocMeta:
    """Прочитать файл и вернуть только реквизиты (для списка файлов в окне)."""
    from .readers import read_source

    source = read_source(path)
    return extract_meta(source)
