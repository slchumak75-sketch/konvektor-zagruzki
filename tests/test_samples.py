"""Разбор реальных файлов заказчика: количество позиций, поля, арифметика."""

from __future__ import annotations

import pytest

from konvektor.extract import extract_file
from konvektor.readers import read_source
from konvektor.validate import check_document_totals, validate_items
from tests.conftest import KNOWN_FILES


@pytest.fixture(scope="module")
def parsed(known_paths):
    out = {}
    for p in known_paths:
        src = read_source(str(p))
        out[p.name] = (src, extract_file(src))
    return out


def test_all_known_files_present(known_paths):
    missing = [n for n in KNOWN_FILES if n not in {p.name for p in known_paths}]
    assert not missing, f"нет файлов-образцов: {missing}"


@pytest.mark.parametrize("name,expected", sorted(KNOWN_FILES.items()))
def test_item_count(parsed, name, expected):
    if name not in parsed:
        pytest.skip(f"файл {name} отсутствует")
    _, res = parsed[name]
    assert len(res.items) == expected


@pytest.mark.parametrize("name", sorted(KNOWN_FILES))
def test_items_have_name_and_money(parsed, name):
    if name not in parsed:
        pytest.skip(f"файл {name} отсутствует")
    _, res = parsed[name]
    assert res.items, "позиции не найдены"
    for it in res.items:
        assert it.name.strip(), f"пустое наименование, строка {it.row}"
        assert any(
            v is not None for v in (it.qty, it.price, it.total_sum, it.vat_sum, it.sum_with_vat)
        ), f"нет ни одного числового значения, строка {it.row}"


@pytest.mark.parametrize("name", sorted(KNOWN_FILES))
def test_no_parser_warnings(parsed, name):
    if name not in parsed:
        pytest.skip(f"файл {name} отсутствует")
    _, res = parsed[name]
    assert res.warnings == []


def test_barcodes_from_column(parsed):
    """ИТ0685355: штрихкод отдельной колонкой, ведущие нули не теряются."""
    _, res = parsed["ИТ0685355.xlsx"]
    assert res.items[0].barcode == "4605922046691"
    assert all(i.barcode for i in res.items if i.row > 6)
    codes = {i.barcode for i in res.items}
    assert "46265531" in codes  # 8-значный код


def test_barcodes_inside_name(parsed):
    """Приложение к ТТН: штрихкод записан прямо в наименовании."""
    _, res = parsed["00000003370Приложение2Листа.xls"]
    assert res.items[0].barcode == "4607011539220"
    assert all(i.barcode for i in res.items)


def test_barcodes_in_ttn1(parsed):
    _, res = parsed["ТТН-1 (вертикальная) № 1019250 от 12.11.2026.xls"]
    assert all(i.barcode for i in res.items)
    assert res.items[0].barcode == "4610273692930"
    # трёхстраничный документ: все страницы прочитаны
    assert len({i.block for i in res.items}) == 3


def test_articles(parsed):
    _, kh = parsed["КХ2104033 (1).xls"]
    assert kh.items[0].article == "8160612"

    _, sch = parsed["Счет СЧЦБ-17059130.xlsx"]
    assert sch.items[0].article == "EC4213-1"

    _, torg = parsed["000004781#0066065.xls"]
    assert torg.items[0].article == "02610-005002"

    _, kor = parsed["Кореличи (4).xls"]
    assert kor.items[0].article == "01_0356659"


def test_article_from_name(parsed):
    _, res = parsed["Счет на оплату № УТ-542 от 12.11.2026.xlsx"]
    assert res.items[0].article == "08135"


def test_multi_page_application(parsed):
    _, res = parsed["00000003370Приложение2Листа.xls"]
    assert len({i.block for i in res.items}) == 4


def test_khaz_encoding_and_quantity(parsed):
    """Файл без CODEPAGE-записи читается в cp1251, количество не теряется."""
    src, res = parsed["000004781#0066065.xls"]
    assert "Бита" in res.items[0].name
    assert res.items[0].qty == 15
    assert res.items[0].total_sum == pytest.approx(26.55)
    assert res.items[0].sum_with_vat == pytest.approx(31.86)


def test_corrupted_xls_is_read(parsed):
    """КХ2104033 — повреждённый OLE-контейнер, но данные читаются."""
    src, res = parsed["КХ2104033 (1).xls"]
    assert src.reader == "xls"
    assert len(res.items) == 19


def test_math_checks_pass_on_samples(parsed):
    items, totals, blocks = [], [], {}
    for name, (src, res) in parsed.items():
        items.extend(res.items)
        totals.extend(res.totals)
        for it in res.items:
            blocks.setdefault((it.source, it.sheet, it.block), []).append(it)
    report = validate_items(items)
    check_document_totals(report, totals, blocks)
    assert report.errors == [], [i.message for i in report.errors]
    assert report.warnings == [], [i.message for i in report.warnings]


def test_vat_mode_is_exclusive_for_samples(parsed):
    items = [i for _, res in parsed.values() for i in res.items]
    report = validate_items(items)
    assert set(report.vat_mode.values()) == {"exclusive"}


def test_document_totals_are_found(parsed):
    _, res = parsed["Счет на оплату № УТ-542 от 12.11.2026.xlsx"]
    assert res.totals and res.totals[0].values.get("total") == pytest.approx(481.0)
