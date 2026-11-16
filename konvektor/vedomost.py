"""Сверка рабочего файла с ведомостью остатков ТМЦ.

«Ведомость» — выгрузка из учётной системы. В ней есть колонки «Склад»,
«Группа номенклатуры», «Артикул», «Номенклатура», «Ед.», «Конечный остаток»,
«Цена последнего поставщика», «Последний поставщик» и далее месяцы. Для
подбора нужны только номенклатура, артикул, штрихкод, единица и последний
поставщик — остальное игнорируется.

Как подбирается номенклатура:

1. **сверка по поставщику** — сначала ищем только среди позиций того же
   поставщика, что указан в документе (в ведомости на 30 тысяч строк иначе
   найдётся «похожее» у чужого поставщика);
2. внутри поставщика — по выбранному ключу: **штрихкод**, **артикул** или
   **наименование**; названия сравниваются нечётко, с учётом сокращений
   («акрил.» ↔ «акриловая», «бел.» ↔ «белая»);
3. если выбранный ключ ничего не дал, пробуем остальные ключи по порядку,
   чтобы позиция всё-таки нашлась (в отчёте видно, каким ключом найдено);
4. **ручной выбор** и **«не выбирать»** перекрывают любой автоподбор.

Модуль не знает про интерфейс: те же функции использует и окно, и консоль.
"""

from __future__ import annotations

import difflib
import os
import re
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .readers import read_source
from .textutil import (
    clean_text,
    ean_ok,
    find_barcode_in_text,
    norm_header,
    parse_number,
    to_text,
)

# ---------------------------------------------------------------- ключи подбора
MATCH_BARCODE = "barcode"
MATCH_ARTICLE = "article"
MATCH_NAME = "name"

#: Подписи кнопок и режимов.
MATCH_TITLES = OrderedDict(
    [
        (MATCH_BARCODE, "по штрихкоду"),
        (MATCH_ARTICLE, "по артикулу"),
        (MATCH_NAME, "по наименованию"),
    ]
)

#: Порядок подбора: выбранный ключ, затем остальные как запасные варианты.
MATCH_ORDER = {
    MATCH_BARCODE: (MATCH_BARCODE, MATCH_ARTICLE, MATCH_NAME),
    MATCH_ARTICLE: (MATCH_ARTICLE, MATCH_BARCODE, MATCH_NAME),
    MATCH_NAME: (MATCH_NAME, MATCH_ARTICLE, MATCH_BARCODE),
}

#: Как в отчёте называется подбор, сделанный пользователем вручную.
HOW_MANUAL = "вручную"

#: С какого сходства названий считаем подбор удавшимся.
AUTO_SCORE = 0.62
#: С какого сходства подбор настолько уверенный, что проверять не нужно.
CONFIDENT_SCORE = 0.80

#: Латинские буквы, которые в выгрузках путают с русскими (SOLАRIS).
_LAT2CYR = {
    "a": "а",
    "b": "в",
    "c": "с",
    "e": "е",
    "h": "н",
    "k": "к",
    "m": "м",
    "o": "о",
    "p": "р",
    "t": "т",
    "x": "х",
    "y": "у",
}

_NON_WORD = re.compile(r"[^0-9a-zа-я]+")
_DIGITS = re.compile(r"\D+")

#: Организационно-правовые формы — в сверке поставщиков они не значат ничего.
_LEGAL_FORMS = {
    "ооо",
    "одо",
    "оао",
    "зао",
    "сооо",
    "чтуп",
    "итуп",
    "ип",
    "чп",
    "сп",
    "тд",
    "уп",
}

#: Заголовки колонок ведомости → роль.
_ROLE_ALIASES = OrderedDict(
    [
        ("nomenclature", ("номенклатура", "наименование", "название", "товар")),
        ("article", ("артикул", "код", "артикул поставщика")),
        ("barcode", ("штрихкод", "штрих код", "баркод", "ean", "штриховой код")),
        ("unit", ("ед", "единица", "единица измерения", "ед изм", "еденица")),
        ("supplier", ("последний поставщик", "поставщик", "контрагент")),
        ("group", ("группа номенклатуры", "группа", "категория", "вид номенклатуры")),
        ("stock", ("конечный остаток", "остаток", "количество")),
        ("price", ("цена последнего поставщика", "цена", "розничная цена")),
    ]
)

#: Служебные строки, которые не являются позицией.
_TOTAL_WORDS = {"итого", "всего", "итог", "итого по", "итого:"}

_MISSING = object()

#: Решение пользователя по позиции: строка ведомости, свободный текст или
#: None («не выбирать»).
Override = Optional[object]


# --------------------------------------------------------------------- тексты
_JUNK_PARENS = re.compile(r"\([^)]*(?:страна|пр-в[ао]|произв|ввоз)[^)]*\)", re.IGNORECASE)
_JUNK_BARCODE = re.compile(r"/\s*\d{8,14}\s*/|\b\d{8,14}\b")


def norm_text(value: Any) -> str:
    """Привести текст к сравнимому виду: нижний регистр, без пунктуации и «ё»."""
    s = to_text(value)
    s = _JUNK_PARENS.sub(" ", s)
    s = _JUNK_BARCODE.sub(" ", s)
    s = s.lower().replace("ё", "е")
    s = "".join(_LAT2CYR.get(ch, ch) for ch in s)
    s = _NON_WORD.sub(" ", s)
    return " ".join(s.split())


def tokens(value: Any) -> List[str]:
    """Значимые слова текста (для сравнения названий)."""
    return [t for t in norm_text(value).split() if t]


def norm_barcode(value: Any) -> str:
    """Только цифры штрихкода (8/12/13/14 знаков), иначе пустая строка.

    8-значные коды принимаем, только если сходится контрольная цифра EAN,
    иначе любой восьмизначный артикул выглядел бы штрихкодом.
    """
    if value is None or isinstance(value, bool):
        return ""
    digits = _DIGITS.sub("", to_text(value))
    if len(digits) in (12, 13, 14):
        return digits
    if len(digits) == 8 and ean_ok(digits):
        return digits
    return ""


def norm_article(value: Any) -> str:
    """Артикул без пробелов/знаков: «01_0356659» → «010356659»."""
    s = to_text(value).lower().replace("ё", "е")
    s = "".join(_LAT2CYR.get(ch, ch) for ch in s)
    return _NON_WORD.sub("", s)


def supplier_key(value: Any) -> str:
    """Ключ поставщика: без кавычек, ОПФ и лишних слов.

    «ООО"Ома"», «ООО «ОМА»» и «ООО  Ома» дают один ключ «ома».
    """
    words = [w for w in tokens(value) if w not in _LEGAL_FORMS]
    return " ".join(words)


def same_supplier(a: Any, b: Any) -> bool:
    """Один ли это поставщик (с точностью до кавычек, ОПФ и опечаток)."""
    ka, kb = supplier_key(a), supplier_key(b)
    if not ka or not kb:
        return False
    if ka == kb:
        return True
    if ka in kb or kb in ka:
        return True
    return difflib.SequenceMatcher(None, ka, kb).ratio() >= 0.9


def _token_pair(a: str, b: str) -> float:
    """Похожесть двух слов с учётом сокращений («акрил.» ↔ «акриловая»)."""
    if a == b:
        return 1.0
    if min(len(a), len(b)) >= 3:
        short, long_ = (a, b) if len(a) <= len(b) else (b, a)
        if long_.startswith(short):
            return 0.86 + 0.08 * (len(short) / len(long_))
    # разные первые буквы — почти наверняка разные слова: не тратим время
    if a[:1] != b[:1]:
        return 0.0
    ratio = difflib.SequenceMatcher(None, a, b).ratio()
    return ratio if ratio >= 0.78 else 0.0


def _coverage(src: Sequence[str], dst: Sequence[str]) -> float:
    """Какая доля слов src нашлась в dst (длинные слова весомее)."""
    total = weight = 0.0
    for token in src:
        w = min(len(token), 8)
        best = 0.0
        for other in dst:
            value = _token_pair(token, other)
            if value > best:
                best = value
                if best >= 1.0:
                    break
        total += w * best
        weight += w
    return total / weight if weight else 0.0


def _score_words(
    doc_tokens: Sequence[str],
    ved_tokens: Sequence[str],
    doc_norm: str,
    ved_norm: str,
    unit: str = "",
    ved_unit: str = "",
) -> float:
    """Оценка сходства по уже разобранным словам (быстро: без повторного разбора)."""
    if not doc_tokens or not ved_tokens:
        return 0.0
    cover_doc = _coverage(doc_tokens, ved_tokens)
    cover_ved = _coverage(ved_tokens, doc_tokens)
    seq = difflib.SequenceMatcher(None, doc_norm, ved_norm).ratio()
    score = 0.5 * cover_doc + 0.25 * cover_ved + 0.25 * seq

    numbers_doc = {t for t in doc_tokens if t.isdigit()}
    numbers_ved = {t for t in ved_tokens if t.isdigit()}
    if numbers_doc and numbers_ved and not (numbers_doc & numbers_ved):
        score *= 0.55

    # главное слово (первое не числовое) должно совпасть: «Бачок для унитаза»
    # и «Сиденье для унитаза» — разные товары, хотя слов общего много.
    # Но если почти все слова ведомости (cover_ved >= 0.85) уже есть в документе,
    # значит в документе просто длинное описание («Средство родентицидное ...»).
    head_doc = next((t for t in doc_tokens if not t.isdigit()), "")
    head_ved = next((t for t in ved_tokens if not t.isdigit()), "")
    if cover_ved < 0.85 and head_doc and head_ved and _token_pair(head_doc, head_ved) < 0.5:
        score *= 0.7
    if cover_ved >= 0.85 and not (numbers_doc and numbers_ved and not (numbers_doc & numbers_ved)):
        score = max(score, 0.85 * cover_ved)

    units = (norm_text(unit)[:2], norm_text(ved_unit)[:2])
    if units[0] and units[0] == units[1]:
        score = min(1.0, score + 0.03)
    return round(score, 4)


def similarity(doc_name: Any, ved_name: Any, unit: str = "", ved_unit: str = "") -> float:
    """Сходство названий документа и ведомости, 0…1.

    Сравниваются слова (чтобы «Таз 35 л» не путался с «Таз 40 л»), числа
    проверяются отдельно, а главное слово — с штрафом: «Бачок для унитаза» и
    «Сиденье для унитаза» это разные товары.
    """
    return _score_words(
        tokens(doc_name), tokens(ved_name), norm_text(doc_name), norm_text(ved_name), unit, ved_unit
    )


# ------------------------------------------------------------------- ведомость
@dataclass
class VedRow:
    """Одна позиция ведомости (только нужные для подбора поля)."""

    nomenclature: str
    article: str = ""
    barcode: str = ""
    unit: str = ""
    supplier: str = ""
    group: str = ""
    stock: float = 0.0
    price: float = 0.0
    line: int = 0
    #: номер строки во внутреннем списке (для быстрых индексов)
    index: int = field(default=-1, repr=False, compare=False)
    _supplier_key: Optional[str] = field(default=None, repr=False, compare=False)

    @property
    def supplier_key(self) -> str:
        if self._supplier_key is None:
            self._supplier_key = supplier_key(self.supplier)
        return self._supplier_key

    def title(self) -> str:
        """Строка для списка выбора: «Наименование · арт. X · ООО …»."""
        parts = [self.nomenclature]
        if self.article:
            parts.append("арт. " + self.article)
        if self.supplier:
            parts.append(self.supplier)
        return " · ".join(parts)


@dataclass
class Vedomost:
    """Ведомость остатков ТМЦ: позиции + быстрый поиск по ним."""

    rows: List[VedRow] = field(default_factory=list)
    path: str = ""
    sheet: str = ""
    header_row: int = 0
    roles: Dict[str, int] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    _by_barcode: Dict[str, VedRow] = field(default_factory=dict, repr=False)
    _by_article: Dict[str, VedRow] = field(default_factory=dict, repr=False)
    _by_name: Dict[str, VedRow] = field(default_factory=dict, repr=False)
    _by_supplier: Dict[str, List[VedRow]] = field(default_factory=dict, repr=False)
    _postings: Dict[str, List[int]] = field(default_factory=dict, repr=False)
    _prefix_postings: Dict[str, List[int]] = field(default_factory=dict, repr=False)
    _norm_names: List[str] = field(default_factory=list, repr=False)
    _word_lists: List[List[str]] = field(default_factory=list, repr=False)
    _scope_cache: Dict[str, Tuple[Dict[str, List[int]], Dict[str, List[int]]]] = field(
        default_factory=dict, repr=False
    )

    # ------------------------------------------------------------------ индекс
    def build_index(self) -> "Vedomost":
        """Построить индексы: штрихкод, артикул, точное название, поставщик, слова."""
        self._by_barcode.clear()
        self._by_article.clear()
        self._by_name.clear()
        self._by_supplier.clear()
        self._postings.clear()
        self._prefix_postings.clear()
        self._scope_cache.clear()
        self._norm_names = []
        self._word_lists = []
        for i, row in enumerate(self.rows):
            row.index = i
            if row.barcode and row.barcode not in self._by_barcode:
                self._by_barcode[row.barcode] = row
            article = norm_article(row.article)
            if article and article not in self._by_article:
                self._by_article[article] = row
            name = norm_text(row.nomenclature)
            words = name.split()
            self._norm_names.append(name)
            self._word_lists.append(words)
            if name and name not in self._by_name:
                self._by_name[name] = row
            key = row.supplier_key
            if key:
                self._by_supplier.setdefault(key, []).append(row)
            for word in set(words):
                self._postings.setdefault(word, []).append(i)
                if len(word) >= 3:
                    self._prefix_postings.setdefault(word[:3], []).append(i)
        return self

    @property
    def count(self) -> int:
        return len(self.rows)

    @property
    def suppliers(self) -> List[Tuple[str, int]]:
        """Поставщики ведомости по убыванию числа позиций: [(имя, сколько), …]."""
        names: Dict[str, Tuple[str, int]] = {}
        for row in self.rows:
            key = row.supplier_key
            if not key:
                continue
            name, number = names.get(key, (row.supplier, 0))
            names[key] = (name, number + 1)
        return sorted(names.values(), key=lambda item: (-item[1], item[0].lower()))

    def rows_of_supplier(self, key: str) -> List[VedRow]:
        if not key:
            return []
        return self._by_supplier.get(key, [])

    def _scope(
        self, supplier: str
    ) -> Optional[Tuple[Dict[str, List[int]], Dict[str, List[int]]]]:
        """Индексы слов только по позициям одного поставщика (строятся один раз)."""
        if not supplier:
            return None
        cached = self._scope_cache.get(supplier)
        if cached is not None:
            return cached
        rows = self._by_supplier.get(supplier) or []
        if not rows:
            return None
        postings: Dict[str, List[int]] = {}
        prefix: Dict[str, List[int]] = {}
        for row in rows:
            for word in set(self._word_lists[row.index]):
                postings.setdefault(word, []).append(row.index)
                if len(word) >= 3:
                    prefix.setdefault(word[:3], []).append(row.index)
        result = (postings, prefix)
        self._scope_cache[supplier] = result
        return result

    def supplier_group(self, supplier: Any) -> Tuple[str, bool]:
        """Найти группу поставщика в ведомости: (ключ, найден ли).

        Если точного совпадения нет, пробуем «похожего» поставщика (опечатка,
        другое написание), но только при очень высоком сходстве.
        """
        key = supplier_key(supplier)
        if not key:
            return "", False
        if key in self._by_supplier:
            return key, True
        best_key, best_score = "", 0.0
        for candidate in self._by_supplier:
            if key in candidate or candidate in key:
                score = 0.95
            else:
                score = difflib.SequenceMatcher(None, key, candidate).ratio()
            if score > best_score:
                best_key, best_score = candidate, score
        if best_score >= 0.88:
            return best_key, True
        return "", False

    # ------------------------------------------------------------------- поиск
    def by_barcode(self, code: Any) -> Optional[VedRow]:
        return self._by_barcode.get(norm_barcode(code))

    def by_article(self, article: Any) -> Optional[VedRow]:
        key = norm_article(article)
        return self._by_article.get(key) if key else None

    def by_name(self, name: Any) -> Optional[VedRow]:
        return self._by_name.get(norm_text(name))

    def candidates(
        self, name: Any, supplier: str = "", limit: int = 40
    ) -> List[Tuple[float, VedRow]]:
        """Похожие позиции ведомости по названию: [(сходство, позиция), …].

        supplier — ключ поставщика: если он задан и такие позиции есть, поиск
        идёт только среди них. Перебирать все 30 тысяч строк на каждую позицию
        нельзя, поэтому сначала отбираем кандидатов по общим словам (точным и
        по первым трём буквам — так находятся сокращения «акрил.»/«акриловая»),
        и только их сравниваем полностью.
        """
        doc_words = tokens(name)
        if not doc_words:
            return []
        index = self._scope(supplier) or (self._postings, self._prefix_postings)
        exact_index, prefix_index = index
        hits: Dict[int, int] = {}
        for word in doc_words:
            for i in exact_index.get(word, ()):
                hits[i] = hits.get(i, 0) + 3
            if len(word) >= 3:
                for i in prefix_index.get(word[:3], ()):
                    hits[i] = hits.get(i, 0) + 1
        if not hits:
            return []
        doc_norm = norm_text(name)
        scored: List[Tuple[float, VedRow]] = []
        for i, _weight in sorted(hits.items(), key=lambda kv: -kv[1])[:600]:
            row = self.rows[i]
            score = _score_words(
                doc_words, self._word_lists[i], doc_norm, self._norm_names[i], "", row.unit
            )
            if score > 0:
                scored.append((score, row))
        scored.sort(key=lambda pair: (-pair[0], pair[1].nomenclature.lower()))
        unique: List[Tuple[float, VedRow]] = []
        seen = set()
        for score, row in scored:
            key = (self._norm_names[row.index], norm_article(row.article))
            if key in seen:
                continue
            seen.add(key)
            unique.append((score, row))
            if len(unique) >= limit:
                break
        return unique

    def search(self, text: Any, supplier: str = "", limit: int = 200) -> List[VedRow]:
        """Ручной поиск по ведомости: подстрока в названии или артикуле.

        Пустой запрос — первые позиции выбранного поставщика (чтобы список
        не был пустым и его можно было листать).
        """
        query = norm_text(text)
        scope = self.rows_of_supplier(supplier) if supplier else []
        rows: Sequence[VedRow] = scope if scope else self.rows
        if not query:
            return list(rows)[:limit]
        query_article = norm_article(text)
        result: List[VedRow] = []
        for row in rows:
            if query in self._norm_names[row.index] or (
                query_article and query_article in norm_article(row.article)
            ):
                result.append(row)
                if len(result) >= limit:
                    break
        return result

    def find(
        self, name: Any, article: Any, barcode: Any, kind: str, supplier: str = ""
    ) -> Optional[Tuple[VedRow, float]]:
        """Найти позицию по одному ключу: (позиция, сходство) или None."""
        if kind == MATCH_BARCODE:
            code = norm_barcode(barcode)
            if not code:
                return None
            row = self._lookup_barcode(code, supplier)
            return (row, 1.0) if row is not None else None
        if kind == MATCH_ARTICLE:
            key = norm_article(article)
            if not key:
                return None
            row = self._lookup_article(key, supplier)
            return (row, 1.0) if row is not None else None
        # по наименованию: сначала точное совпадение, затем нечёткое
        exact = self.by_name(name)
        if exact is not None and self._in_scope(exact, supplier):
            return exact, 1.0
        for score, row in self.candidates(name, supplier, limit=1):
            return row, score
        return None

    def _lookup_barcode(self, code: str, supplier: str) -> Optional[VedRow]:
        row = self._by_barcode.get(code)
        if row is not None and self._in_scope(row, supplier):
            return row
        for candidate in self.rows_of_supplier(supplier):
            if candidate.barcode and candidate.barcode == code:
                return candidate
        return None

    def _lookup_article(self, key: str, supplier: str) -> Optional[VedRow]:
        row = self._by_article.get(key)
        if row is not None and self._in_scope(row, supplier):
            return row
        for candidate in self.rows_of_supplier(supplier):
            if norm_article(candidate.article) == key:
                return candidate
        return None

    @staticmethod
    def _in_scope(row: VedRow, supplier: str) -> bool:
        if not supplier:
            return True
        return row.supplier_key == supplier


# -------------------------------------------------------------------- разбор
def _role_of(header: Any) -> str:
    """Роль колонки по её заголовку (пустая строка — колонка не нужна)."""
    text = norm_header(header)
    if not text:
        return ""
    for role, aliases in _ROLE_ALIASES.items():
        for alias in aliases:
            if text == alias or text.startswith(alias):
                return role
    return ""


def _find_header(rows: Sequence[Sequence[Any]]) -> Tuple[int, Dict[str, int]]:
    """Найти строку заголовков ведомости и колонки нужных полей."""
    best: Tuple[int, Dict[str, int]] = (0, {})
    for index, row in enumerate(rows[:200]):
        roles: Dict[str, int] = {}
        for col, cell in enumerate(row):
            role = _role_of(cell)
            if role and role not in roles:
                roles[role] = col
        if "nomenclature" not in roles:
            continue
        if len(roles) > len(best[1]):
            best = (index, roles)
        if len(roles) >= 4:
            break
    return best


def _cell(row: Sequence[Any], col: Optional[int]) -> Any:
    if col is None or col < 0 or col >= len(row):
        return None
    return row[col]


def parse_vedomost(source: Any) -> Vedomost:
    """Разобрать уже прочитанный файл ведомости (SourceFile)."""
    warnings: List[str] = []
    chosen_sheet = None
    header_index = 0
    roles: Dict[str, int] = {}
    for sheet in source.sheets:
        index, found = _find_header(sheet.rows)
        if "nomenclature" in found:
            chosen_sheet, header_index, roles = sheet, index, found
            break
    if chosen_sheet is None:
        raise ValueError(
            "в файле не найдена колонка «Номенклатура»: это не похоже на ведомость "
            "остатков. Нужна выгрузка, где есть заголовки «Артикул», "
            "«Номенклатура», «Ед.», «Последний поставщик»."
        )

    columns = {
        "nomenclature": roles.get("nomenclature"),
        "article": roles.get("article"),
        "barcode": roles.get("barcode"),
        "unit": roles.get("unit"),
        "supplier": roles.get("supplier"),
        "group": roles.get("group"),
        "stock": roles.get("stock"),
        "price": roles.get("price"),
    }
    if columns["supplier"] is None:
        warnings.append(
            "в ведомости нет колонки «Последний поставщик» — сверка по поставщику "
            "невозможна, подбор идёт по всей ведомости"
        )

    rows: List[VedRow] = []
    for offset, raw in enumerate(chosen_sheet.rows[header_index + 1 :], start=header_index + 2):
        name = clean_text(_cell(raw, columns["nomenclature"]))
        if not name or norm_text(name) in _TOTAL_WORDS:
            continue
        article = clean_text(_cell(raw, columns["article"]))
        if article in ("-", "—", "–"):
            article = ""
        barcode = norm_barcode(_cell(raw, columns["barcode"])) if columns["barcode"] is not None else ""
        if not barcode:
            barcode = norm_barcode(article) or (find_barcode_in_text(name) or "")
        price = 0.0
        if columns["price"] is not None:
            price = parse_number(_cell(raw, columns["price"])) or 0.0
        stock = 0.0
        if columns["stock"] is not None:
            stock = parse_number(_cell(raw, columns["stock"])) or 0.0
        rows.append(
            VedRow(
                nomenclature=name,
                article=article,
                barcode=barcode,
                unit=clean_text(_cell(raw, columns["unit"])),
                supplier=clean_text(_cell(raw, columns["supplier"]))
                if columns["supplier"] is not None
                else "",
                group=clean_text(_cell(raw, columns["group"]))
                if columns["group"] is not None
                else "",
                stock=stock,
                price=price,
                line=offset,
            )
        )

    vedomost = Vedomost(
        rows=rows,
        path=getattr(source, "path", "") or "",
        sheet=chosen_sheet.name,
        header_row=header_index + 1,
        roles={role: col for role, col in columns.items() if col is not None},
        warnings=warnings,
    )
    vedomost.build_index()
    if not rows:
        raise ValueError(
            "в ведомости не нашлось ни одной позиции: под заголовками нет строк "
            "с заполненной номенклатурой"
        )
    return vedomost


def load_vedomost(path: str) -> Vedomost:
    """Прочитать ведомость из файла (xls/xlsx/csv — любой поддерживаемый)."""
    if not path or not os.path.isfile(path):
        raise ValueError(f"файл ведомости не найден: {path}")
    try:
        source = read_source(path)
    except Exception as exc:  # noqa: BLE001 - показываем пользователю причину
        raise ValueError(
            f"не удалось прочитать ведомость «{os.path.basename(path)}» "
            f"({type(exc).__name__}: {exc}). Закройте файл в Excel и попробуйте снова."
        ) from exc
    vedomost = parse_vedomost(source)
    if not vedomost.path:
        vedomost.path = path
    return vedomost


# --------------------------------------------------------------------- сверка
@dataclass
class MatchEntry:
    """Результат подбора для одной позиции документа (для окна и отчёта)."""

    key: Tuple[str, str, int]
    source: str = ""
    sheet: str = ""
    row: int = 0
    doc_name: str = ""
    doc_article: str = ""
    doc_barcode: str = ""
    name: str = ""
    how: str = ""  # штрихкод / артикул / наименование / вручную / не выбирать
    score: float = 0.0
    matched: bool = False
    skipped: bool = False  # «не выбирать»
    review: bool = False  # подбор не уверенный — стоит проверить
    supplier: str = ""  # поставщик ведомости, среди которого искали
    supplier_ok: bool = False
    ved_row: Optional[VedRow] = None

    @property
    def doc_title(self) -> str:
        return self.doc_name

    @property
    def how_title(self) -> str:
        if self.skipped:
            return "не выбирать"
        if self.how:
            return self.how
        return "не подобрано"

    @property
    def score_title(self) -> str:
        if self.skipped:
            return "—"
        if not self.matched:
            return "—"
        if self.how in (MATCH_BARCODE, MATCH_ARTICLE) and self.score >= 1.0:
            return "точно"
        if self.how == HOW_MANUAL:
            return HOW_MANUAL
        return f"{self.score:.0%}"


class Matcher:
    """Подбор номенклатуры ведомости для позиций документа.

    Один объект — на одну конвертацию: pipeline вызывает begin_file() перед
    разбором файла и apply() для каждой позиции. Ручные решения хранятся в
    overrides и переживают повторную конвертацию.
    """

    def __init__(
        self,
        vedomost: Vedomost,
        strategy: str = MATCH_NAME,
        overrides: Optional[Dict[Tuple[str, str, int], "Override"]] = None,
        min_score: float = AUTO_SCORE,
        supplier: str = "",
    ) -> None:
        self.vedomost = vedomost
        self.strategy = strategy if strategy in MATCH_TITLES else MATCH_NAME
        #: поставщик, выбранный пользователем в окне (пусто — берём из документа)
        self.forced_supplier = clean_text(supplier)
        self.overrides: Dict[Tuple[str, str, int], Optional[VedRow]] = dict(overrides or {})
        self.min_score = min_score
        self.entries: List[MatchEntry] = []
        self.by_key: Dict[Tuple[str, str, int], MatchEntry] = {}
        self._supplier_key = ""
        self._supplier_text = ""
        self._supplier_found = False
        self._file = ""

    # -------------------------------------------------------------- служебное
    def begin_file(self, path: str, supplier: str = "") -> None:
        """Начать новый файл: поставщик задаёт область поиска.

        Поставщик, выбранный пользователем в окне, важнее найденного в
        документе: в счёте поставщик может быть не распознан или написан
        иначе, а ведомость знает правильное название.
        """
        self._file = path
        self._supplier_text = self.forced_supplier or clean_text(supplier)
        key, found = self.vedomost.supplier_group(self._supplier_text)
        self._supplier_key = key
        self._supplier_found = found

    @property
    def supplier_text(self) -> str:
        return self._supplier_text

    @property
    def strategy_title(self) -> str:
        return MATCH_TITLES.get(self.strategy, MATCH_TITLES[MATCH_NAME])

    @property
    def summary(self) -> Dict[str, Any]:
        """Короткий отчёт для подсказки в окне и для --json."""
        matched = [e for e in self.entries if e.matched]
        counters = {kind: 0 for kind in MATCH_TITLES}
        counters[HOW_MANUAL] = 0
        for entry in matched:
            if entry.how in counters:
                counters[entry.how] += 1
        return {
            "strategy": self.strategy,
            "supplier": self._supplier_text,
            "supplier_found": self._supplier_found,
            "supplier_key": self._supplier_key,
            "items": len(self.entries),
            "matched": len(matched),
            "barcode": counters[MATCH_BARCODE],
            "article": counters[MATCH_ARTICLE],
            "name": counters[MATCH_NAME],
            "manual": counters[HOW_MANUAL],
            "novinka": sum(1 for e in self.entries if not e.matched),
            "skipped": sum(1 for e in self.entries if e.skipped),
            "review": sum(1 for e in self.entries if e.review),
            "unmatched": sum(1 for e in self.entries if not e.matched and not e.skipped),
        }

    def summary_text(self) -> str:
        """Фраза для строки состояния: «сверено: 9 из 11 (наименование 9)…»."""
        s = self.summary
        if not s["items"]:
            return ""
        parts = []
        for kind in MATCH_TITLES:
            if s[kind]:
                parts.append(f"{MATCH_TITLES[kind]} {s[kind]}")
        if s["manual"]:
            parts.append(f"вручную {s['manual']}")
        detail = f" ({', '.join(parts)})" if parts else ""
        text = f"сверено с ведомостью: {s['matched']} из {s['items']}{detail}"
        if s["skipped"]:
            text += f", не выбирать {s['skipped']}"
        if s["review"]:
            text += f", проверить {s['review']}"
        if s["unmatched"]:
            text += f", не подобрано {s['unmatched']} (пойдут как «Новинка»)"
        if s["supplier"] and not s["supplier_found"]:
            text += (
                " · поставщика «" + str(s["supplier"]) + "» в ведомости нет — "
                "подбор только вручную"
            )
        return text

    # --------------------------------------------------------------- подбор
    def apply(self, item: Any) -> MatchEntry:
        """Подобрать номенклатуру для позиции и переименовать её."""
        key = (getattr(item, "source", "") or "", getattr(item, "sheet", "") or "", int(getattr(item, "row", 0) or 0))
        override = self.overrides.get(key, _MISSING)
        entry = MatchEntry(
            key=key,
            source=key[0],
            sheet=key[1],
            row=key[2],
            doc_name=getattr(item, "name", "") or "",
            doc_article=getattr(item, "article", "") or "",
            doc_barcode=getattr(item, "barcode", "") or "",
            supplier=self._supplier_text,
            supplier_ok=self._supplier_found,
        )

        if override is not _MISSING:
            if override is None:
                # «не выбирать»: остаётся наименование документа
                entry.skipped = True
                entry.name = entry.doc_name
            elif isinstance(override, str):
                # пользователь вписал название прямо в таблице
                entry.matched = True
                entry.how = HOW_MANUAL
                entry.score = 1.0
                entry.name = override
                try:
                    item.name = override
                except Exception:  # noqa: BLE001 - позиция может быть неизменяемой
                    pass
            else:
                self._use(entry, override, HOW_MANUAL, 1.0, item)
        else:
            found = self._auto(item, key)
            if found is not None:
                row, how, score, supplier = found
                entry.supplier = supplier or self._supplier_text
                self._use(entry, row, how, score, item)
            else:
                entry.name = entry.doc_name

        # «Новинка» — номенклатуру не выбрали (нет в ведомости или отказались)
        try:
            item.novinka = not entry.matched
        except Exception:  # noqa: BLE001 - позиция может быть неизменяемой
            pass
        self.entries.append(entry)
        self.by_key[key] = entry
        return entry

    def _use(
        self, entry: MatchEntry, row: VedRow, how: str, score: float, item: Any
    ) -> None:
        entry.matched = True
        entry.how = how
        entry.score = score
        entry.ved_row = row
        entry.name = row.nomenclature
        entry.review = how in (MATCH_NAME, MATCH_ARTICLE, MATCH_BARCODE) and score < CONFIDENT_SCORE
        # в документе наименование заменяем на ведомостное — этого просил заказчик
        try:
            item.name = row.nomenclature
            if not (getattr(item, "article", "") or "") and row.article:
                item.article = row.article
            if not (getattr(item, "barcode", "") or "") and row.barcode:
                item.barcode = row.barcode
            if not (getattr(item, "unit", "") or "") and row.unit:
                item.unit = row.unit
        except Exception:  # noqa: BLE001 - позиция может быть неизменяемой
            pass

    def _auto(
        self, item: Any, key: Tuple[str, str, int]
    ) -> Optional[Tuple[VedRow, str, float, str]]:
        name = getattr(item, "name", "") or ""
        article = getattr(item, "article", "") or ""
        barcode = getattr(item, "barcode", "") or ""
        if not barcode:
            barcode = find_barcode_in_text(name) or ""

        # Сверка по поставщику: если поставщик документа есть в ведомости —
        # ищем только среди его позиций (иначе легко подобрать товар чужого
        # поставщика). Поставщика в ведомости нет — не подбираем вовсе: пусть
        # пользователь выберет вручную, чем получит неверное наименование.
        # Поставщик в документе вообще не распознан — ищем по всей ведомости.
        if self._supplier_key:
            scopes = [self._supplier_key]
        elif self._supplier_text:
            scopes = []
        else:
            scopes = [""]

        for scope in scopes:
            for kind in MATCH_ORDER[self.strategy]:
                found = self.vedomost.find(name, article, barcode, kind, scope)
                if found is None:
                    continue
                row, score = found
                if kind == MATCH_NAME and score < self.min_score:
                    continue
                return row, kind, score, row.supplier
        return None

    # ------------------------------------------------------------- отчёты
    def report_lines(self, limit: int = 20) -> List[str]:
        """Строки отчёта: что подобрано, что требует проверки, что не найдено."""
        lines = [self.summary_text()] if self.summary_text() else []
        for entry in self.entries:
            if entry.matched and not entry.review:
                continue
            if entry.skipped:
                lines.append(f"  не выбирать: {entry.doc_name}")
            elif entry.matched:
                lines.append(
                    f"  проверить: {entry.doc_name} → {entry.name} ({entry.how}, {entry.score:.0%})"
                )
            else:
                lines.append(f"  не подобрано: {entry.doc_name}")
            if len(lines) > limit:
                lines.append("  …")
                break
        return lines


def apply_vedomost(
    items: Sequence[Any],
    vedomost: Vedomost,
    strategy: str = MATCH_NAME,
    supplier: str = "",
    overrides: Optional[Dict[Tuple[str, str, int], "Override"]] = None,
) -> Matcher:
    """Подобрать номенклатуру для готового списка позиций (для консоли и тестов)."""
    matcher = Matcher(vedomost, strategy, overrides)
    matcher.begin_file("", supplier)
    for item in items:
        matcher.apply(item)
    return matcher
