"""Проверки сверки с ведомостью остатков ТМЦ (konvektor/vedomost.py).

Ведомость в тестах — синтетическая: та же раскладка колонок, что в выгрузке
заказчика («Склад», «Группа номенклатуры», «Артикул», «Номенклатура», «Ед.»,
«Конечный остаток», «Цена последнего поставщика», «Последний поставщик»), но
5–8 строк вместо 33 тысяч. Так тесты не зависят от конкретного файла.
"""

from __future__ import annotations

import os

import openpyxl
import pytest

from konvektor.extract import Item
from konvektor.pipeline import ConvertOptions, convert
from konvektor.readers import Sheet, SourceFile
from konvektor.vedomost import (
    MATCH_ARTICLE,
    MATCH_BARCODE,
    MATCH_NAME,
    Matcher,
    Vedomost,
    load_vedomost,
    norm_article,
    norm_barcode,
    parse_vedomost,
    same_supplier,
    similarity,
    supplier_key,
)

HEADER = [
    "Склад",
    "Группа номенклатуры",
    "Артикул",
    "Номенклатура",
    "Ед.",
    "Расход",
    "Конечный\nостаток",
    "Свободный остаток",
    "Цена\nпоследнего \nпоставщика",
    "Розничная \nцена",
    "Последний\nпоставщик",
]

ROWS = [
    ("Склад Мир", "АКРИЛ", " ", "Грунтовка акрил.Тайфун Мастер №100 бел.1кг.", "шт", "0.00", "12.00", "12.00", "8.50", " ", 'ООО "ОМА"'),
    ("Склад Мир", "АКРИЛ", " ", "Грунтовка акрил.Тайфун Мастер №102 бел.1кг.", "шт", "0.00", "4.00", "4.00", "8.60", " ", 'ООО "ОМА"'),
    ("Склад Мир", "АКРИЛ", " "  , "Грунтовка акрил.Тайфун Мастер №100 бел.1кг.", "шт", "0.00", "3.00", "3.00", "9.10", " ", 'ООО "Культбытхозторгплюс"'),
    ("Склад Мир", "ГАЗ", " ", "Газ сжижен.в однораз.баллоне WERKER 520см3", "шт", "0.00", "7.00", "7.00", "3.20", " ", 'ООО"Ома"'),
    ("Склад Мир", "КРЕПЁЖ", "02610-005002", "Бита Impact торсионная ударная PH2 50мм", "упак", "0.00", "2.00", "2.00", "5.40", " ", 'ООО "СИСТЕМА ООО"'),
    ("Склад Мир", "КРЕПЁЖ", " "  , "Ботинки рабочие кожаные стальной подносок", "пар", "0.00", "1.00", "1.00", "44.00", " ", 'ООО "СИСТЕМА ООО"'),
    ("Склад Мир", "КРЕПЁЖ", " "  , "Инвертор сварочный SOLАRIS MMA-200D", "шт", "0.00", "1.00", "1.00", "310.00", " ", 'ООО "ТД Комплект"'),
    ("Склад Мир", "ЭМАЛЬ", " "  , "Эмаль металлик универсальная 520 мл VX19100", "шт", "0.00", "5.00", "5.00", "10.90", " ", 'ООО "ТД Комплект"'),
]


def grid() -> list:
    """Матрица листа ведомости: 3 пустые строки, название, шапка, данные."""
    out = [[None] * len(HEADER) for _ in range(6)]
    out[3][3] = "Ведомость по остаткам ТМЦ на складах"
    out[5] = list(HEADER)
    for row in ROWS:
        out.append(list(row))
    return out


def make_vedomost(rows=None) -> Vedomost:
    source = SourceFile(
        path="Ведомость.xlsx",
        sheets=[Sheet(name="Лист_1", rows=grid() if rows is None else rows)],
        reader="xlsx",
    )
    return parse_vedomost(source)


def item(name: str, supplier: str = "", **kwargs) -> Item:
    """Позиция документа (то, что отдаёт extract)."""
    return Item(source="счет.xls", sheet="Лист1", row=kwargs.pop("row", 1), name=name, **kwargs)


# ------------------------------------------------------------------ разбор
def test_columns_are_found_by_headers():
    ved = make_vedomost()
    assert ved.count == len(ROWS)
    assert ved.header_row == 6
    # ключевые колонки найдены по названиям, а не по номерам
    assert ved.roles["nomenclature"] == 3
    assert ved.roles["article"] == 2
    assert ved.roles["supplier"] == 10
    first = ved.rows[0]
    assert first.nomenclature.startswith("Грунтовка акрил.")
    assert first.unit == "шт"
    assert first.supplier == 'ООО "ОМА"'
    assert first.price == pytest.approx(8.5)


def test_file_without_nomenclature_is_rejected():
    rows = [["Склад", "Количество"], ["Мир", 5]]
    with pytest.raises(ValueError) as err:
        make_vedomost(rows)
    assert "Номенклатура" in str(err.value)


def test_warning_when_no_supplier_column():
    rows = [["Артикул", "Номенклатура", "Ед."], ["A1", "Шуруп 4х40", "шт"]]
    ved = make_vedomost(rows)
    assert any("Последний поставщик" in w for w in ved.warnings)
    assert ved.count == 1


def test_empty_ведомость_is_rejected():
    rows = [list(HEADER)]
    with pytest.raises(ValueError) as err:
        make_vedomost(rows)
    assert "ни одной позиции" in str(err.value)


def test_vedomost_is_loaded_from_xlsx_file(tmp_path):
    path = tmp_path / "Ведомость 08.11.2026.xlsx"
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Лист_1"
    for row in grid():
        sheet.append(list(row))
    book.save(path)

    ved = load_vedomost(str(path))
    assert ved.count == len(ROWS)
    assert os.path.basename(ved.path) == path.name


def test_missing_file_gives_clear_message(tmp_path):
    with pytest.raises(ValueError) as err:
        load_vedomost(str(tmp_path / "нет такого.xlsx"))
    assert "не найден" in str(err.value)


# ---------------------------------------------------- нормализация и сходство
def test_supplier_key_ignores_quotes_spaces_and_legal_form():
    assert supplier_key('ООО"Ома"') == supplier_key("ООО «ОМА»") == "ома"
    assert same_supplier('ООО "Торговый дом "Залог успеха"', "ООО «Торговый дом «Залог успеха»")
    assert not same_supplier('ООО "ОМА"', 'ООО "Авидо"')


def test_article_and_barcode_normalisation():
    assert norm_article("01_0356659") == norm_article("01 0356659")
    # 8 цифр без верной контрольной суммы — это артикул, а не штрихкод
    assert norm_barcode("00000115") == ""
    assert norm_barcode("4607011539220") == "4607011539220"
    assert norm_barcode("4 607011 539220") == "4607011539220"


def test_similarity_prefers_right_product():
    base = similarity(
        "Грунтовка акриловая Тайфун Мастер № 100 белая, 1кг",
        "Грунтовка акрил.Тайфун Мастер №100 бел.1кг.",
    )
    other = similarity(
        "Грунтовка акриловая Тайфун Мастер № 100 белая, 1кг",
        "Грунтовка акрил.Тайфун Мастер №102 бел.1кг.",
    )
    assert base > other
    assert base >= 0.8


def test_similarity_penalises_other_head_noun():
    """«Бачок для унитаза» и «Сиденье для унитаза» — разные товары."""
    score = similarity("Бачок для унитаза-компакт AGIO A", "Сиденье для унитаза AV Engineering")
    assert score < 0.62


# -------------------------------------------------------------------- подбор
def test_match_by_name_stays_inside_supplier():
    ved = make_vedomost()
    matcher = Matcher(ved, MATCH_NAME)
    matcher.begin_file("счет.xls", 'ООО "ОМА"')
    doc = item("Грунтовка акриловая Тайфун Мастер № 100 белая, 1кг")
    entry = matcher.apply(doc)

    assert entry.matched
    assert entry.how == MATCH_NAME
    assert doc.name == "Грунтовка акрил.Тайфун Мастер №100 бел.1кг."
    # позиция именно «своего» поставщика, хотя та же номенклатура есть у чужого
    assert entry.ved_row.supplier == 'ООО "ОМА"'


def test_position_of_foreign_supplier_is_not_taken():
    ved = make_vedomost()
    matcher = Matcher(ved, MATCH_NAME)
    matcher.begin_file("счет.xls", 'ООО "СИСТЕМА ООО"')
    doc = item("Грунтовка акриловая Тайфун Мастер № 100 белая 1 кг")
    entry = matcher.apply(doc)

    assert not entry.matched
    assert doc.name == "Грунтовка акриловая Тайфун Мастер № 100 белая 1 кг"


def test_unknown_supplier_means_manual_only():
    ved = make_vedomost()
    matcher = Matcher(ved, MATCH_NAME)
    matcher.begin_file("ттн.xls", 'ООО "КАРИО"')
    entry = matcher.apply(item("Газ сжиженный WERKER 520см3"))

    assert not entry.matched
    assert "в ведомости нет" in matcher.summary_text()


def test_no_supplier_in_document_searches_whole_vedomost():
    ved = make_vedomost()
    matcher = Matcher(ved, MATCH_NAME)
    matcher.begin_file("прайс.xls", "")
    entry = matcher.apply(item("Газ сжиженный WERKER в одноразовом баллоне 520см3"))

    assert entry.matched
    assert entry.how == MATCH_NAME


def test_match_by_article():
    ved = make_vedomost()
    matcher = Matcher(ved, MATCH_ARTICLE)
    matcher.begin_file("счет.xls", 'ООО "СИСТЕМА ООО"')
    doc = item("Бита ударная PH2 50мм", article="02610-005002")
    entry = matcher.apply(doc)

    assert entry.matched
    assert entry.how == MATCH_ARTICLE
    assert entry.score == 1.0
    assert entry.score_title == "точно"
    assert doc.name == "Бита Impact торсионная ударная PH2 50мм"


def test_match_by_barcode():
    rows = [list(HEADER)] + [
        ["Склад", "ЭМАЛЬ", " ", "Эмаль металлик универсальная 520 мл", "шт", "0", "5", "0",
         "10,90", " ", 'ООО "ТД Комплект"']
    ]
    ved = make_vedomost(rows)
    # штрихкод берём из наименования: в ведомости он может быть в названии,
    # а в счёте — в отдельной колонке
    ved.rows[0].barcode = "4607011539220"
    ved.build_index()

    matcher = Matcher(ved, MATCH_BARCODE)
    matcher.begin_file("счет.xls", 'ООО "ТД Комплект"')
    doc = item("Эмаль металлик 520 мл", barcode="4607011539220")
    entry = matcher.apply(doc)

    assert entry.matched
    assert entry.how == MATCH_BARCODE
    assert doc.name == "Эмаль металлик универсальная 520 мл"


def test_barcode_strategy_falls_back_to_name():
    """Выбран штрихкод, но его в ведомости нет — подбираем по названию."""
    ved = make_vedomost()
    matcher = Matcher(ved, MATCH_BARCODE)
    matcher.begin_file("счет.xls", 'ООО "ОМА"')
    doc = item("Грунтовка акриловая Тайфун Мастер № 100 белая, 1кг", barcode="4605544048424")
    entry = matcher.apply(doc)

    assert entry.matched
    assert entry.how == MATCH_NAME


def test_manual_choice_and_refusal():
    ved = make_vedomost()
    chosen = ved.rows[1]  # «…№102 бел.1кг.»
    key = ("счет.xls", "Лист1", 7)
    matcher = Matcher(ved, MATCH_NAME, overrides={key: chosen})
    matcher.begin_file("счет.xls", 'ООО "ОМА"')
    doc = item("Что-то совсем непохожее", row=7)
    entry = matcher.apply(doc)

    assert entry.matched and entry.how == "вручную"
    assert doc.name == chosen.nomenclature

    # «не выбирать»: наименование документа остаётся как есть
    matcher = Matcher(ved, MATCH_NAME, overrides={key: None})
    matcher.begin_file("счет.xls", 'ООО "ОМА"')
    doc = item("Грунтовка акриловая Тайфун Мастер № 100 белая, 1кг", row=7)
    entry = matcher.apply(doc)

    assert entry.skipped
    assert not entry.matched
    assert entry.how_title == "не выбирать"
    assert doc.name == "Грунтовка акриловая Тайфун Мастер № 100 белая, 1кг"


def test_summary_counts_and_text():
    ved = make_vedomost()
    matcher = Matcher(ved, MATCH_NAME)
    matcher.begin_file("счет.xls", 'ООО "ОМА"')
    matcher.apply(item("Грунтовка акриловая Тайфун Мастер № 100 белая, 1кг", row=1))
    matcher.apply(item("Газ сжиженный WERKER в одноразовом баллоне 520см3", row=2))
    matcher.apply(item("Совершенно другой товар", row=3))

    summary = matcher.summary
    assert summary["items"] == 3
    assert summary["matched"] == 2
    assert summary["name"] == 2
    assert summary["unmatched"] == 1
    text = matcher.summary_text()
    assert text.startswith("сверено с ведомостью: 2 из 3")
    assert matcher.report_lines()  # отчёт есть всегда, когда есть позиции


def test_missing_article_is_filled_from_ведомость():
    ved = make_vedomost()
    matcher = Matcher(ved, MATCH_ARTICLE)
    matcher.begin_file("счет.xls", 'ООО "СИСТЕМА ООО"')
    doc = item("Бита ударная PH2", article="02610 005002")
    matcher.apply(doc)
    assert doc.article == "02610 005002"  # свой артикул не перетираем


def test_search_finds_by_name_and_article():
    ved = make_vedomost()
    assert [r.nomenclature for r in ved.search("грунтовка")] == [
        "Грунтовка акрил.Тайфун Мастер №100 бел.1кг.",
        "Грунтовка акрил.Тайфун Мастер №102 бел.1кг.",
        "Грунтовка акрил.Тайфун Мастер №100 бел.1кг.",
    ]
    assert ved.search("02610")[0].nomenclature.startswith("Бита")
    # пустой запрос — первые позиции выбранного поставщика
    scoped = ved.search("", supplier_key('ООО "СИСТЕМА ООО"'))
    assert len(scoped) == 2


def test_candidates_are_ordered_by_similarity():
    ved = make_vedomost()
    found = ved.candidates("Грунтовка акриловая Тайфун Мастер 100 белая 1кг")
    assert found
    assert found[0][1].nomenclature.startswith("Грунтовка акрил.Тайфун Мастер №100")
    assert found[0][0] >= found[-1][0]


# ------------------------------------------------------------------ pipeline
def _write_book(path, rows):
    book = openpyxl.Workbook()
    sheet = book.active
    for row in rows:
        sheet.append(list(row))
    book.save(path)


def test_convert_renames_items_with_vedomost(tmp_path):
    """Кнопка конвертации с загруженной ведомостью переименовывает позиции."""
    ved_path = tmp_path / "Ведомость.xlsx"
    _write_book(ved_path, grid())
    ved = load_vedomost(str(ved_path))

    doc_path = tmp_path / "счет.xlsx"
    _write_book(
        doc_path,
        [
            ["Наименование", "Количество", "Цена", "Сумма"],
            ["Газ сжиженный WERKER в однораз.баллоне (220гр/520см3)", 2, 3.2, 6.4],
        ],
    )

    matcher = Matcher(ved, MATCH_NAME)
    result = convert([str(doc_path)], "", ConvertOptions(write=False, matcher=matcher), log=lambda m: None)

    assert result.items_count == 1
    assert result.items[0].name == "Газ сжижен.в однораз.баллоне WERKER 520см3"
    assert result.match_summary["matched"] == 1
    assert result.match_log
    assert result.items[0].unit == "шт"  # единица подставлена из ведомости


def test_convert_without_vedomost_keeps_names(tmp_path):
    doc_path = tmp_path / "счет.xlsx"
    _write_book(
        doc_path,
        [
            ["Наименование", "Количество", "Цена", "Сумма"],
            ["Газ сжиженный WERKER (220гр/520см3)", 2, 3.2, 6.4],
        ],
    )
    result = convert([str(doc_path)], "", ConvertOptions(write=False), log=lambda m: None)
    assert result.items[0].name == "Газ сжиженный WERKER (220гр/520см3)"
    assert result.match_summary == {}


def test_free_text_override_is_used_as_is():
    """Название, вписанное пользователем в таблице, попадает в результат."""
    ved = make_vedomost()
    key = ("счет.xls", "Лист1", 3)
    matcher = Matcher(ved, MATCH_NAME, overrides={key: "Своё название 1 кг"})
    matcher.begin_file("счет.xls", 'ООО "ОМА"')
    doc = item("Грунтовка акриловая Тайфун Мастер № 100 белая, 1кг", row=3)
    entry = matcher.apply(doc)

    assert entry.matched and entry.how == "вручную"
    assert entry.how_title == "вручную"
    assert doc.name == "Своё название 1 кг"
    assert matcher.summary["manual"] == 1
