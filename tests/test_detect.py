"""Тесты распознавания шапки и колонок (на реальных заголовках из образцов)."""

from __future__ import annotations

import pytest

from konvektor import detect


@pytest.mark.parametrize(
    "header,kind",
    [
        # все встречающиеся в образцах названия колонок
        ("Штрихкод", detect.BARCODE),
        ("Штрих-код", detect.BARCODE),
        ("EAN-PEACE (штуки)", detect.BARCODE),
        ("Артикул", detect.ARTICLE),
        ("Код", detect.ARTICLE),
        ("Наименование товара", detect.NAME),
        ("Наименование\nтовара, продукции", detect.NAME),
        ("Наименование материала", detect.NAME),
        ("Номенклатура", detect.NAME),
        ("Товар", detect.NAME),
        ("Товары (работы, услуги)", detect.NAME),
        ("Единица\nизме-\nрения", detect.UNIT),
        ("Ед. изм.", detect.UNIT),
        ("Ед.", detect.UNIT),
        ("Коли-\nчество", detect.QTY),
        ("Количество", detect.QTY),
        ("К-во", detect.QTY),
        ("Кол-во", detect.QTY),
        ("Заказанное количество (шт.)", detect.QTY),
        ("Цена,\nруб., коп.", detect.PRICE),
        ("Цена", detect.PRICE),
        ("Цена шт. б/НДС, руб", detect.PRICE),
        ("Стоимость,\nруб., коп.", detect.SUM),
        ("Сумма", detect.SUM),
        ("Сумма без НДС,     руб. коп.", detect.SUM),
        ("Ставка \nНДС, %", detect.VAT_RATE),
        ("Ставка НДС%", detect.VAT_RATE),
        ("Сумма НДС,\nруб. коп.", detect.VAT_SUM),
        ("НДС", detect.VAT_SUM),
        ("Стоимость с НДС, \nруб., коп.", detect.TOTAL),
        ("Сумма с НДС, руб", detect.TOTAL),
        ("Всего\nс НДС", detect.TOTAL),
        ("Всего", detect.TOTAL),
        ("Всего с НДС", detect.TOTAL),
        # то, что нам не нужно
        ("№", None),
        ("№ п/п", None),
        ("Код ТНВЭД", None),
        ("Количество грузовых мест", None),
        ("Масса груза", None),
        ("Вес", None),
        ("Страна производитель", None),
        ("Примечание", None),
        ("Фасовка (шт. в кр.)", None),
        ("Объём, M3", None),
        ("Номер заказа", None),
        ("Сертификат", None),
        ("", None),
    ],
)
def test_classify_header(header, kind):
    res = detect.classify_header(header)
    assert (res[0] if res else None) == kind


def test_price_prefers_without_vat():
    a = detect.classify_header("Цена шт. б/НДС, руб")
    b = detect.classify_header("Цена шт. с НДС, руб")
    assert a[0] == b[0] == detect.PRICE
    assert a[1] > b[1]


def test_detect_header_tht():
    row = [
        "№",
        "Наименование товара",
        "Единица измерения",
        "Количество",
        "Цена, руб. коп.",
        "Стоимость, руб. коп.",
        "Ставка НДС, %",
        "Сумма НДС, руб. коп.",
        "Стоимость с НДС, руб. коп.",
        "Количество грузовых мест",
        "Масса груза",
        "Примечание",
    ]
    h = detect.detect_header(row)
    assert h is not None
    assert h.col(detect.NAME) == 1
    assert h.col(detect.UNIT) == 2
    assert h.col(detect.QTY) == 3
    assert h.col(detect.PRICE) == 4
    assert h.col(detect.SUM) == 5
    assert h.col(detect.VAT_RATE) == 6
    assert h.col(detect.VAT_SUM) == 7
    assert h.col(detect.TOTAL) == 8


def test_detect_header_needs_name_and_numbers():
    assert detect.detect_header(["Наименование", "Единица", "Количество"]) is None
    assert detect.detect_header(["1", "2", "3", "4", "5"]) is None


def test_total_label_and_scope():
    assert detect.total_label(["ИТОГО ПО СТРАНИЦЕ", "х", 16]) == "итого по странице"
    assert detect.total_label(["Итого по листу №  1"]) == "итого по листу № 1"
    assert detect.total_label(["Всего сумма НДС", "прописью"]) == "всего сумма ндс"
    assert detect.total_label(["Таз 35 л", 1, 15.47]) is None
    assert detect.total_scope("итого по листу № 1") == "block"
    assert detect.total_scope("итого") == "document"


def test_total_field_from_label():
    assert detect.total_field_from_label("Сумма без НДС :") == "sum"
    assert detect.total_field_from_label("Сумма НДС :") == "vat_sum"
    assert detect.total_field_from_label("Итого с НДС:") == "total"
    assert detect.total_field_from_label("Итого:") == ""
    assert detect.total_field_from_label("Всего масса груза") == "skip"
    assert detect.total_field_from_label("В том числе доставка:") == "skip"


def test_is_numbering_row():
    assert detect.is_numbering_row(["1", "2", "3", "4", "5"])
    assert not detect.is_numbering_row(["1", "Товар", "2", "3", "4"])
    assert not detect.is_numbering_row(["2", "1", "3", "4"])


def test_is_furniture():
    assert detect.is_furniture("Номер листа: 2")
    assert detect.is_furniture("I. ТОВАРНЫЙ РАЗДЕЛ")
    assert detect.is_furniture("Контактное лицо: Специалист")
    assert not detect.is_furniture("Газ сжиженный WERKER в баллоне")
