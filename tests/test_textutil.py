"""Тесты разбора значений: числа, штрихкоды, артикулы."""

from __future__ import annotations

import pytest

from konvektor.textutil import (
    clean_text,
    ean_ok,
    find_article_in_text,
    find_barcode_in_text,
    fmt_money,
    norm_header,
    parse_barcode,
    parse_number,
    parse_rate,
)


@pytest.mark.parametrize(
    "raw,expected",
    [
        (6.0, 6.0),
        (10.73, 10.73),
        ("1 006.75", 1006.75),
        ("1 006,75", 1006.75),
        ("          68,80", 68.8),
        ("1 234 567,89", 1234567.89),
        ("20 %", 20.0),
        ("20%", 20.0),
        ("руб. 15,47", 15.47),
        ("0,006", 0.006),
        ("-5,5", -5.5),
        ("х", None),
        ("Х", None),
        ("-", None),
        ("б/н", None),
        ("", None),
        (None, None),
        ("Шестьсот сорок четыре руб. 29 копеек", None),
        ("29 коп", 29.0),
        ("1,0л", None),
        ("ТУ BY 812004107.016-2018", None),
        (True, None),
    ],
)
def test_parse_number(raw, expected):
    assert parse_number(raw) == expected


@pytest.mark.parametrize(
    "raw,expected",
    [(4607011539220.0, "4607011539220"), ("4605922046691", "4605922046691"), ("46265531", "46265531"), ("", None), ("нет", None)],
)
def test_parse_barcode(raw, expected):
    assert parse_barcode(raw) == expected


def test_parse_barcode_from_scientific():
    assert parse_barcode("4,60701153922E+12") == "4607011539220"


@pytest.mark.parametrize(
    "code,ok",
    [
        ("4607011539220", True),  # реальный EAN-13 из образца
        ("4810285007566", True),
        ("4607011539221", False),
        ("812004107", False),  # 9 цифр — не EAN
    ],
)
def test_ean_ok(code, ok):
    assert ean_ok(code) is ok


def test_find_barcode_in_name_with_commas():
    """Штрихкод в тексте наименования обрамлён запятыми (выгрузка 1С)."""
    text = "VX19100 VIXEN Эмаль металлик, алюминий, аэрозоль 520 мл, 4607011539220, страна происх. РОССИЯ"
    assert find_barcode_in_text(text) == "4607011539220"


def test_find_barcode_variants():
    assert find_barcode_in_text("Нож, штрих-код 4610273692930, Китай") == "4610273692930"
    assert find_barcode_in_text("Грунт 6,5л /4607141950148/ Страна пр-ва РФ") == "4607141950148"
    assert find_barcode_in_text("Средство 750 мл, ТУ BY 812004107.016-2018") is None
    assert find_barcode_in_text("") is None


@pytest.mark.parametrize(
    "text,expected",
    [
        ('Таз 35 л "строительный" (h255;640x470) арт. 08135, Изготовитель: ООО "Милих"', "08135"),
        ("Кружка эмал. арт.01-1612/4-Лимоны штрих код 4680284100359, Россия", "01-1612/4"),
        ("8160612 Вода питьевая обработанная «АУРА» негаз. ПЭТ 1,0л.", "8160612"),
        ("Обычное наименование без кода", None),
    ],
)
def test_find_article_in_text(text, expected):
    assert find_article_in_text(text) == expected


def test_find_article_ignores_barcode_and_company():
    text = 'Ковш ст.эмал. 1,0л. Черешенка штрих код 4603784385460 ООО ТД "МеталТрейд" Россия'
    article = find_article_in_text(text, barcode="4603784385460")
    assert article is None


def test_norm_header():
    assert norm_header("Ставка \nНДС, %") == "ставка ндс, %"
    assert norm_header("  Коли-\nчество ") == "коли- чество"
    assert norm_header("Единица\nизме-\nрения") == "единица изме- рения"
    assert norm_header(None) == ""


def test_clean_text():
    assert clean_text("  а   б\nс ") == "а б с"
    assert clean_text(None) == ""


def test_parse_rate():
    assert parse_rate("20%") == 20.0
    assert parse_rate("20 %") == 20.0
    assert parse_rate(20.0) == 20.0
    assert parse_rate("без НДС") == 0.0
    assert parse_rate(None) is None


def test_fmt_money():
    assert fmt_money(1234.5) == "1 234,50"
    assert fmt_money(None) == ""
