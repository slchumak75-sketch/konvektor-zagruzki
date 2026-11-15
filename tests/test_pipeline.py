"""Сквозные тесты: файлы -> итоговая книга."""

from __future__ import annotations

import os

import pytest
from openpyxl import load_workbook

from konvektor import TEMPLATE_COLUMNS, __version__
from konvektor.pipeline import (
    MODE_FILES,
    MODE_SHEETS,
    ConvertOptions,
    convert,
    default_output_name,
    selftest,
)


@pytest.fixture(scope="module")
def combined(known_paths, tmp_path_factory):
    out = tmp_path_factory.mktemp("out") / "Итог.xlsx"
    result = convert([str(p) for p in known_paths], str(out))
    return result, str(out)


def test_combined_result(combined):
    result, path = combined
    assert result.items_count == 301
    assert result.error_count == 0
    assert result.warning_count == 0
    assert result.computed_count == 140
    assert result.outputs == [path]
    assert os.path.exists(path)


def test_workbook_structure(combined):
    _, path = combined
    wb = load_workbook(path)
    assert wb.sheetnames == ["Данные", "Проверки", "Сводка"]

    ws = wb["Данные"]
    header = [c.value for c in ws[1]]
    assert header == TEMPLATE_COLUMNS
    assert ws.max_row == 2 + 301  # шапка + позиции + итог
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref == "A1:J302"

    # строка итогов — формулы
    total_row = ws.max_row
    assert ws.cell(row=total_row, column=3).value == "ИТОГО"
    for col in (7, 9, 10):
        assert str(ws.cell(row=total_row, column=col).value).startswith("=SUM(")

    # штрихкоды — текстом, иначе Excel испортит длинные числа
    c = ws.cell(row=2, column=1)
    assert isinstance(c.value, str) and c.value.isdigit()
    assert c.number_format == "@"


def test_summary_sheet(combined):
    _, path = combined
    ws = load_workbook(path)["Сводка"]
    rows = list(ws.iter_rows(values_only=True))
    assert rows[0][0] == "Файл"
    # столбцы: Файл, Формат, Лист(ы), Поставщик, УНП, Документ, Дата, Позиций, …
    assert rows[0][3] == "Поставщик"
    assert rows[0][5] == "Документ"
    assert rows[0][6] == "Дата документа"
    files = [r[0] for r in rows[1:] if r[0] and r[7] is not None and r[0] != "ВСЕГО"]
    assert len(files) == 10
    total = next(r for r in rows if r[0] == "ВСЕГО")
    assert total[7] == 301
    assert round(total[10], 2) == pytest.approx(16579.02, abs=0.01)


def test_summary_sheet_has_supplier_and_date(combined):
    """Поставщик и дата документа выгружаются в «Сводку» по каждому файлу."""
    _, path = combined
    ws = load_workbook(path)["Сводка"]
    rows = {r[0]: r for r in ws.iter_rows(min_row=2, values_only=True) if r[0]}
    row = next(v for k, v in rows.items() if k.endswith("_№ 10802 от 12 ноября 2026 г .XLS"))
    assert row[3] == 'ООО "Культбытхозторгплюс"'
    assert row[4] == "192553707"
    assert row[5] == "Счет № 3596616"
    assert row[6] == "12.11.2026"


def test_result_metas(combined):
    result, _ = combined
    assert len(result.metas) == 10
    by_name = {os.path.basename(p): m for p, m in result.metas.items()}
    assert by_name["Кореличи (4).xls"].supplier == 'ООО "ОМА"'
    assert by_name["Кореличи (4).xls"].date_text == "19.10.2026"
    stat = next(s for s in result.stats if s.path.endswith("Кореличи (4).xls"))
    assert stat.supplier == 'ООО "ОМА"'
    assert stat.doc_date == "19.10.2026"
    assert stat.doc_number == "11121-000020449"


def test_issues_sheet_marks_computed(combined):
    _, path = combined
    ws = load_workbook(path)["Проверки"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert rows
    assert {r[1] for r in rows} == {"РАСЧЁТ"}
    # жёлтая заливка у рассчитанных значений в «Данных»
    ws_data = load_workbook(path)["Данные"]
    filled = [
        ws_data.cell(row=r, column=7).fill.start_color.rgb
        for r in range(2, ws_data.max_row + 1)
        if ws_data.cell(row=r, column=7).fill and ws_data.cell(row=r, column=7).fill.start_color.rgb
    ]
    assert any(rgb and str(rgb).endswith("FFF2CC") for rgb in filled)


def test_mode_sheets(known_paths, tmp_path):
    out = tmp_path / "по_листам.xlsx"
    result = convert(
        [str(p) for p in known_paths], str(out), ConvertOptions(mode=MODE_SHEETS)
    )
    assert result.error_count == 0
    wb = load_workbook(str(out))
    assert len([n for n in wb.sheetnames if n not in ("Проверки", "Сводка")]) == len(known_paths)
    assert all(n[:31] for n in wb.sheetnames)


def test_mode_files(known_paths, tmp_path):
    out = tmp_path / "по_файлам.xlsx"
    result = convert([str(p) for p in known_paths], str(out), ConvertOptions(mode=MODE_FILES))
    assert len(result.outputs) == len(known_paths)
    for path in result.outputs:
        assert os.path.exists(path)
        wb = load_workbook(path)
        assert "Данные" in wb.sheetnames
        assert wb["Данные"].cell(row=1, column=1).value == "Штрихкод"


def test_source_columns(known_paths, tmp_path):
    out = tmp_path / "с_источником.xlsx"
    convert(
        [str(p) for p in known_paths],
        str(out),
        ConvertOptions(add_source_columns=True),
    )
    ws = load_workbook(str(out))["Данные"]
    header = [c.value for c in ws[1]]
    assert header[:10] == TEMPLATE_COLUMNS
    assert header[10:] == ["файл-источник", "строка"]
    assert ws.cell(row=2, column=11).value.endswith(".xls")


def test_selftest(known_paths, sample_dir):
    ok, message = selftest(sample_dir)
    assert ok, message
    assert "позиций 301" in message


def test_empty_input(tmp_path):
    result = convert([], str(tmp_path / "пусто.xlsx"))
    assert result.items_count == 0
    assert not result.outputs


def test_missing_file(tmp_path):
    result = convert([str(tmp_path / "нет_такого.xlsx")], str(tmp_path / "out.xlsx"))
    assert result.errors
    assert result.items_count == 0


def test_garbage_file_does_not_crash(tmp_path, known_paths):
    """Файл-мусор не должен ломать конвертацию остальных."""
    junk = tmp_path / "мусор.xlsx"
    junk.write_bytes(b"not a table at all")
    good = [str(p) for p in known_paths[:2]]
    out = tmp_path / "out.xlsx"
    result = convert(good + [str(junk)], str(out))
    assert result.items_count > 0
    assert any("мусор.xlsx" in note for note in result.notes)
    assert os.path.exists(str(out))


def test_no_validate_mode(known_paths, tmp_path):
    out = tmp_path / "без_проверок.xlsx"
    result = convert(
        [str(p) for p in known_paths], str(out), ConvertOptions(validate=False)
    )
    assert result.report is not None
    assert result.computed_count == 0


def test_default_output_name():
    name = default_output_name("/tmp")
    assert name.endswith(".xlsx")
    assert "Загрузка_" in name


def test_version_present():
    assert __version__
