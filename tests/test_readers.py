"""Тесты чтения файлов разных форматов."""

from __future__ import annotations

import pytest

from konvektor.readers import read_source, sniff_kind


def test_sniff_kind():
    assert sniff_kind(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1rest") == "xls"
    assert sniff_kind(b"PK\x03\x04rest") == "xlsx"
    assert sniff_kind(b"<html><body><table></table></body></html>") == "html"
    assert sniff_kind(b"a;b;c\n1;2;3") == "csv"
    assert sniff_kind(b"", "file.CSV") == "csv"


@pytest.mark.parametrize(
    "name,reader",
    [
        ("00000003370ЭТТН.xls", "xls"),
        ("000004781#0066065.xls", "xls"),
        ("КХ2104033 (1).xls", "xls"),
        ("ТТН-1 (вертикальная) № 1019250 от 12.11.2026.xls", "xls"),
        ("ИТ0685355.xlsx", "xlsx"),
        ("Счет СЧЦБ-17059130.xlsx", "xlsx"),
        ("Счет на оплату № УТ-542 от 12.11.2026.xlsx", "xlsx"),
    ],
)
def test_reader_kind(known_paths, name, reader):
    path = next((p for p in known_paths if p.name == name), None)
    if path is None:
        pytest.skip(f"файл {name} отсутствует")
    sf = read_source(str(path))
    assert sf.reader == reader
    assert sf.sheets


def test_sheets_have_rows(known_paths):
    for p in known_paths:
        sf = read_source(str(p))
        assert sf.sheets[0].height > 5
        assert sf.sheets[0].width > 3


def test_html_table(tmp_path):
    html = """<html><body><table>
      <tr><th>Наименование товара</th><th>Ед. изм.</th><th>Количество</th><th>Цена</th><th>Сумма</th></tr>
      <tr><td>Товар &laquo;А&raquo;</td><td>шт</td><td>2</td><td>10,50</td><td>21,00</td></tr>
    </table></body></html>"""
    path = tmp_path / "price.xls"  # HTML с «неправильным» расширением — частая ситуация
    path.write_text(html, encoding="utf-8")
    sf = read_source(str(path))
    assert sf.reader == "html"
    from konvektor.extract import extract_file

    res = extract_file(sf)
    assert len(res.items) == 1
    assert res.items[0].name == "Товар «А»"
    assert res.items[0].qty == 2
    assert res.items[0].total_sum == 21.0


def test_csv_with_semicolon_and_comma(tmp_path):
    content = "Наименование;Ед.;Количество;Цена;Сумма\nТовар Б;шт;3;5,25;15,75\n"
    path = tmp_path / "price.csv"
    path.write_text(content, encoding="cp1251")
    sf = read_source(str(path))
    assert sf.reader == "csv"
    from konvektor.extract import extract_file

    res = extract_file(sf)
    assert len(res.items) == 1
    assert res.items[0].total_sum == 15.75


def test_xlsx_inside_xls_container(tmp_path):
    """Поставщики иногда кладут xlsx в файл с расширением .xls."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["Штрихкод", "Наименование товара", "Ед. изм.", "Количество", "Цена", "Сумма"])
    ws.append(["1234567890123", "Тестовый товар", "шт", 1, 100, 100])
    tmp = tmp_path / "fake.xls"
    wb.save(tmp)
    sf = read_source(str(tmp))
    assert sf.reader == "xlsx"
    from konvektor.extract import extract_file

    res = extract_file(sf)
    assert len(res.items) == 1
    assert res.items[0].barcode == "1234567890123"


def test_broken_file_returns_empty_source(tmp_path):
    """Нечитаемый файл не должен ронять программу: возвращаем пустой результат."""
    from konvektor.extract import extract_file

    path = tmp_path / "broken.xlsx"
    path.write_bytes(b"\x00\x01\x02\x03 not a spreadsheet")
    sf = read_source(str(path))
    res = extract_file(sf)
    assert res.items == []
    assert res.warnings  # «шапка таблицы не найдена»
