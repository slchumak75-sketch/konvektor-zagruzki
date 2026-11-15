"""Реквизиты документа: поставщик, номер и дата (konvektor/meta.py).

Здесь же — регрессия по всем файлам-образцам: что именно программа считает
поставщиком и датой каждого документа.
"""

from __future__ import annotations

import datetime as dt

import pytest

from konvektor.meta import (
    DocMeta,
    extract_meta,
    fmt_date,
    from_filename,
    looks_like_company,
    meta_from_path,
    parse_date,
)
from konvektor.meta import _doc_type_from
from konvektor.readers import read_source

DATE = dt.date

#: Ожидаемые реквизиты файлов-образцов: имя -> (поставщик, документ, дата).
EXPECTED = {
    "00000003370Приложение2Листа.xls": ("", "Приложение к ТТН № 2880985", "15.10.2026"),
    "00000003370ЭТТН.xls": ("", "ТТН № 2880985", "15.10.2026"),
    "000004781#0066065.xls": ('ООО "СИСТЕМА ООО"', "ТТН № 0066065", "12.11.2026"),
    "_№ 10802 от 12 ноября 2026 г .XLS": (
        'ООО "Культбытхозторгплюс"',
        "Счет № 3596616",
        "12.11.2026",
    ),
    "ИТ0685355.xlsx": ("", "№ ИТ0685355", "29.10.2026"),
    "КХ2104033 (1).xls": ('ОАО "Лидское пиво"', "Накладная № КХ2104033", "21.10.2026"),
    "Кореличи (4).xls": ('ООО "ОМА"', "Счет № 11121-000020449", "19.10.2026"),
    "Счет СЧЦБ-17059130.xlsx": ("", "Счет № 17059130", ""),
    "Счет на оплату № УТ-542 от 12.11.2026.xlsx": (
        "ООО «Торговый дом «Залог успеха»",
        "Счет № 542",
        "12.11.2026",
    ),
    "ТТН-1 (вертикальная) № 1019250 от 12.11.2026.xls": (
        'ООО "КАРИО"',
        "ТТН-1 № 1019250",
        "12.11.2026",
    ),
}


# ---------------------------------------------------------------- даты
@pytest.mark.parametrize(
    "value, expected",
    [
        ("12 ноября 2026", DATE(2026, 11, 12)),
        ("12 ноября 2026 г.", DATE(2026, 11, 12)),
        ("1 января 2026", DATE(2026, 1, 1)),
        ("19.10.2026", DATE(2026, 10, 19)),
        ("19.10.20", DATE(2020, 10, 19)),
        ("19/10/2026", DATE(2026, 10, 19)),
        ("от 19.10.2026 г.", DATE(2026, 10, 19)),
        (dt.date(2026, 3, 4), DATE(2026, 3, 4)),
        (dt.datetime(2026, 3, 4, 12, 30), DATE(2026, 3, 4)),
        (46316, DATE(2026, 10, 21)),  # число Excel
        (46316.0, DATE(2026, 10, 21)),
    ],
)
def test_parse_date(value, expected):
    assert parse_date(value) == expected


@pytest.mark.parametrize("value", [None, "", "без даты", 10, 999999, True, "32.13.2026"])
def test_parse_date_rejects_garbage(value):
    assert parse_date(value) is None


def test_fmt_date():
    assert fmt_date(DATE(2026, 11, 12)) == "12.11.2026"
    assert fmt_date(None) == ""


# ---------------------------------------------------------------- мелочи
@pytest.mark.parametrize(
    "text, expected",
    [
        ("Счет № 542 от 12.11.2026", "Счет"),
        ("Счёт на оплату", "Счет"),
        ("ТТН-1 № 1019250", "ТТН-1"),
        ("Приложение к товарно-транспортной накладной № 2880985", "Приложение к ТТН"),
        ("Товарно-транспортная накладная", "ТТН"),
        ("Накладная № КХ2104033", "Накладная"),
        ("Счет-фактура № 7", "Счет-фактура"),
        ("Акт № ИТ0685355", "Акт"),
        # «компакт» — не «акт», «расчет» — не «счет»
        ("Бачок для унитаза-компакт AGIO A", ""),
        ("Итого расчет", ""),
        ("", ""),
    ],
)
def test_doc_type(text, expected):
    assert _doc_type_from(text) == expected


@pytest.mark.parametrize(
    "text, expected",
    [
        ('ООО "ОМА"', True),
        ("ОАО «Лидское пиво»", True),
        ("ООО Культбытхозторгплюс", True),
        ("УНП 100005713", False),
        ("гп. Мир, ул. Красноармейская 1", False),
        ("Итого", False),
        ("", False),
    ],
)
def test_looks_like_company(text, expected):
    assert looks_like_company(text) is expected


# ---------------------------------------------------------------- имя файла
def test_from_filename_number_date_type():
    number, date, doc_type = from_filename("ТТН-1 (вертикальная) № 1019250 от 12.11.2026.xls")
    assert (number, date, doc_type) == ("1019250", DATE(2026, 11, 12), "ТТН-1")


def test_from_filename_russian_month():
    number, date, _ = from_filename("_№ 10802 от 12 ноября 2026 г .XLS")
    assert (number, date) == ("10802", DATE(2026, 11, 12))


def test_from_filename_without_date():
    assert meta_from_path.__name__  # имя файла без даты — дата не выдумывается
    number, date, _ = from_filename("ИТ0685355.xlsx")
    assert number == "ИТ0685355" and date is None


# ---------------------------------------------------------------- модель
def test_doc_label_and_flags():
    meta = DocMeta(doc_type="Счет", doc_number="542", doc_date=DATE(2026, 11, 12))
    assert meta.doc_label == "Счет № 542"
    assert meta.date_text == "12.11.2026"
    assert meta.found

    empty = DocMeta()
    assert empty.doc_label == "" and empty.date_text == "" and not empty.found


# ---------------------------------------------------------------- образцы
@pytest.mark.parametrize("name, expected", sorted(EXPECTED.items()))
def test_sample_meta(known_paths, name, expected):
    path = next((p for p in known_paths if p.name == name), None)
    if path is None:  # образца нет в этой копии репозитория
        pytest.skip("нет файла-образца")
    meta = meta_from_path(str(path))
    supplier, doc_label, date_text = expected
    assert meta.supplier == supplier
    assert meta.doc_label == doc_label
    assert meta.date_text == date_text


def test_meta_warns_when_nothing_found(known_paths):
    path = next((p for p in known_paths if p.name == "Счет СЧЦБ-17059130.xlsx"), None)
    if path is None:
        pytest.skip("нет файла-образца")
    meta = extract_meta(read_source(str(path)))
    assert "поставщик не найден в документе" in meta.warnings
    assert "дата документа не найдена" in meta.warnings


def test_meta_does_not_take_contract_number(known_paths):
    """«Договор № 528 от 31.05.2016» — не номер и не дата накладной."""
    path = next(
        (p for p in known_paths if p.name.startswith("ТТН-1 (вертикальная)")), None
    )
    if path is None:
        pytest.skip("нет файла-образца")
    meta = extract_meta(read_source(str(path)))
    assert meta.doc_number == "1019250"
    assert meta.doc_date == DATE(2026, 11, 12)
