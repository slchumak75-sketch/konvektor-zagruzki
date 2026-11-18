"""Сквозные тесты: файлы -> итоговая книга."""

from __future__ import annotations

import datetime as dt
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
    assert header[10:] == ["файл", "строка"]
    assert ws.cell(row=2, column=11).value.endswith(".xls")


def test_selftest(known_paths, sample_dir):
    ok, message = selftest(sample_dir)
    # с новыми файлами (МАВ/Элсви/цветогамма) самопроверка имеет ошибки в служебных файлах, но позиции считаются
    assert "позиций 388" in message
    assert "файлов с данными 15 из 19" in message


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


def test_write_falls_back_when_folder_is_locked(known_paths, tmp_path, monkeypatch):
    """Если рядом с файлом писать нельзя — результат сохраняется в домашнюю папку."""
    import konvektor.pipeline as pipeline
    from konvektor import writer

    real = writer.write_workbook
    calls = {"n": 0}

    def flaky(path, payload):
        calls["n"] += 1
        if calls["n"] == 1:
            raise PermissionError("папка только для чтения")
        return real(path, payload)

    monkeypatch.setattr(pipeline, "write_workbook", flaky)
    out = tmp_path / "out.xlsx"
    result = convert([str(known_paths[0])], str(out))
    assert result.outputs and os.path.exists(result.outputs[0])
    assert result.outputs[0] != str(out)
    assert any("не удалось записать" in n for n in result.notes)


def test_analysis_is_built_for_every_file(combined):
    result, _ = combined
    assert len(result.analysis) == 10
    for report in result.analysis.values():
        assert report.items > 0
        assert report.blocks
        assert 0.0 <= report.confidence <= 1.0


# ---------------------------------------------------------- имя результата
def test_output_path_for():
    import datetime as dt

    from konvektor.meta import DocMeta
    from konvektor.pipeline import documents_dir, output_filename, output_path_for, safe_path_part

    meta = DocMeta(
        supplier='ООО "Культбытхозторгплюс"', doc_date=dt.date(2026, 11, 12)
    )
    assert safe_path_part('ООО "Культбытхозторгплюс"') == "ООО Культбытхозторгплюс"
    assert safe_path_part("ООО «ОМА»") == "ООО ОМА"
    assert safe_path_part('a/b:c*?"<>|') == "a b c"

    # папка поставщика не создаётся: поставщик и дата — в имени файла.
    # Дата в имени — дата сохранения (сегодняшняя), а не дата документа.
    today = f"{dt.date.today():%d.%m.%Y}"
    assert output_filename(meta) == f"ООО Культбытхозторгплюс {today}.xlsx"
    # дату можно передать явно (например, для тестов)
    assert (
        output_filename(meta, date=dt.date(2026, 11, 12))
        == "ООО Культбытхозторгплюс 12.11.2026.xlsx"
    )
    path = output_path_for(meta, "/tmp/база")
    assert path == os.path.join("/tmp/база", f"ООО Культбытхозторгплюс {today}.xlsx")
    # по умолчанию — папка «Документы»
    assert output_path_for(meta) == os.path.join(
        documents_dir(), f"ООО Культбытхозторгплюс {today}.xlsx"
    )
    # нет реквизитов — понятные имена, а не «None.xlsx»
    fallback = output_path_for(DocMeta(), "/tmp/база")
    assert fallback.endswith(f"Поставщик не определён {today}.xlsx")


def test_documents_dir_exists():
    from konvektor.pipeline import documents_dir

    assert os.path.isdir(documents_dir())


def test_auto_naming_uses_supplier_and_date_file(known_paths, tmp_path):
    """naming="auto": один файл «поставщик дата.xlsx», без подпапок."""
    path = next((p for p in known_paths if p.name.startswith("Кореличи")), None)
    if path is None:
        pytest.skip("нет файла-образца")
    result = convert(
        [str(path)],
        "",
        ConvertOptions(naming="auto", add_source_columns=True, output_dir=str(tmp_path)),
    )
    assert result.outputs, result.notes
    written = result.outputs[0]
    today = f"{dt.date.today():%d.%m.%Y}"
    assert written == str(tmp_path / f"ООО ОМА {today}.xlsx"), written
    assert os.path.exists(written)
    # никаких подпапок поставщика
    assert [p.name for p in tmp_path.iterdir()] == [f"ООО ОМА {today}.xlsx"]


def test_write_false_only_plans_file(known_paths, tmp_path):
    """write=False: считаем и показываем таблицу, но файл не создаём (кнопка «Сохранить»)."""
    path = next((p for p in known_paths if p.name.startswith("Кореличи")), None)
    if path is None:
        pytest.skip("нет файла-образца")
    result = convert(
        [str(path)], "", ConvertOptions(naming="auto", output_dir=str(tmp_path), write=False)
    )
    assert result.items_count == 11
    assert result.outputs == []
    assert list(tmp_path.iterdir()) == []
    expected = str(tmp_path / f"ООО ОМА {dt.date.today():%d.%m.%Y}.xlsx")
    assert result.planned_output == expected, result.planned_output
    # предпросмотр строится из тех же позиций, что попали бы в книгу
    assert len(result.items) == 11
    assert result.items[0].name


def test_no_columns_after_sum_with_vat(combined):
    """После «сумма с НДС» колонок нет: ровно 10 колонок образца."""
    _, path = combined
    ws = load_workbook(path)["Данные"]
    header = [c.value for c in ws[1] if c.value is not None]
    assert header == TEMPLATE_COLUMNS
    assert ws.max_column == len(TEMPLATE_COLUMNS)


def test_quantity_and_vat_are_plain_numbers(combined):
    """Количество и ставка НДС записаны целыми и с узким форматом — без «12,00»."""
    _, path = combined
    ws = load_workbook(path)["Данные"]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row - 1):
        qty, rate = row[4], row[7]
        if isinstance(qty.value, float):
            assert qty.value.is_integer(), qty.value
        if isinstance(rate.value, float):
            assert rate.value.is_integer(), rate.value
        # целые — формат «0», чтобы в Excel не было ни «12,00», ни «12,»
        assert qty.number_format == "0", qty.number_format
        assert rate.number_format == "0", rate.number_format


def test_field_number_format():
    from konvektor.writer import field_number_format, plain_number

    assert field_number_format("количество", 12) == "0"
    assert field_number_format("количество", 0.5) == "0.###"
    assert field_number_format("ставка НДС", 20) == "0"
    assert field_number_format("ставка НДС") == "0.##"
    assert field_number_format("цена", 2.5) == "#,##0.00"
    assert field_number_format("") == "#,##0.00"
    assert plain_number(12.0) == 12 and isinstance(plain_number(12.0), int)
    assert plain_number(0.5) == 0.5
    # «6,» и «1 234,5» из текстовых ячеек поставщика тоже становятся числами
    assert plain_number("6,") == 6
    assert plain_number("1 234,5") == 1234.5
    assert plain_number("") == ""
    assert plain_number("нет") == "нет"


def test_repeated_conversion_creates_new_file(known_paths, tmp_path):
    """Повторная конвертация не затирает прошлый файл: появляется «… (2).xlsx»."""
    path = next((p for p in known_paths if p.name.startswith("Кореличи")), None)
    if path is None:
        pytest.skip("нет файла-образца")
    today = f"{dt.date.today():%d.%m.%Y}"
    first = convert(
        [str(path)], "", ConvertOptions(naming="auto", output_dir=str(tmp_path))
    )
    second = convert(
        [str(path)], "", ConvertOptions(naming="auto", output_dir=str(tmp_path))
    )
    assert first.outputs == [str(tmp_path / f"ООО ОМА {today}.xlsx")]
    assert second.outputs == [str(tmp_path / f"ООО ОМА {today} (2).xlsx")]
    assert len(list(tmp_path.iterdir())) == 2
    # первый файл остался на месте
    assert os.path.exists(first.outputs[0])


def test_planned_output_is_unique_too(known_paths, tmp_path):
    """Предпросмотр показывает то имя, под которым файл реально появится."""
    path = next((p for p in known_paths if p.name.startswith("Кореличи")), None)
    if path is None:
        pytest.skip("нет файла-образца")
    today = f"{dt.date.today():%d.%m.%Y}"
    (tmp_path / f"ООО ОМА {today}.xlsx").write_text("старый результат")
    result = convert(
        [str(path)],
        "",
        ConvertOptions(naming="auto", output_dir=str(tmp_path), write=False),
    )
    assert result.planned_output == str(tmp_path / f"ООО ОМА {today} (2).xlsx")


def test_downloads_dir_exists():
    from konvektor.pipeline import downloads_dir

    assert os.path.isdir(downloads_dir())


def test_vat_rate_fraction_becomes_percent():
    """Ставка НДС, записанная долей (0.2), превращается в 20 — без «0,2» в таблице."""
    from konvektor.extract import Item
    from konvektor.validate import validate_items

    item = Item(
        source="файл.xls", sheet="Лист1", row=5, name="Товар", qty=2.0, price=10.0, vat_rate=0.2
    )
    validate_items([item])
    assert item.vat_rate == 20.0


def test_unavailable_folder_raises_clear_error(known_paths, tmp_path):
    """Недоступная папка — понятная ошибка, а не «[Errno 2] No such file»."""
    path = next((p for p in known_paths if p.name.startswith("Кореличи")), None)
    if path is None:
        pytest.skip("нет файла-образца")
    blocker = tmp_path / "файл_вместо_папки.txt"
    blocker.write_text("не папка")
    bogus = str(blocker / "внутри")
    with pytest.raises(ValueError) as exc:
        convert([str(path)], "", ConvertOptions(naming="auto", output_dir=bogus))
    assert "недоступна для записи" in str(exc.value)
    assert "Куда сохранить" in str(exc.value)


# ------------------------------------------- папка «склад / поставщик / дата»
def test_output_path_with_sklad_and_supplier():
    """1.7.0: при выбранном складе — <папка>/<склад>/<поставщик>/<дата>.xlsx."""
    import datetime as dt

    from konvektor.meta import DocMeta
    from konvektor.pipeline import output_path_for

    day = dt.date(2026, 11, 15)
    meta = DocMeta(supplier='ООО "ОМА"')
    path = output_path_for(meta, "/tmp/база", date=day, sklad="Мир")
    assert path == os.path.join("/tmp/база", "Мир", "ООО ОМА", "15.11.2026.xlsx"), path

    # поставщик, выбранный пользователем, важнее найденного в документе
    path = output_path_for(meta, "/tmp/база", date=day, sklad="Кореличи", supplier="ООО ТД Комплект")
    assert path == os.path.join("/tmp/база", "Кореличи", "ООО ТД Комплект", "15.11.2026.xlsx")

    # поставщика нет ни там, ни там — понятное имя папки
    path = output_path_for(DocMeta(), "/tmp/база", date=day, sklad="Мир")
    assert path.endswith(os.path.join("Мир", "Поставщик не определён", "15.11.2026.xlsx"))

    # без склада — прежнее имя файла
    assert output_path_for(meta, "/tmp/база", date=day).endswith("ООО ОМА 15.11.2026.xlsx")


def test_convert_puts_file_into_sklad_supplier_folder(known_paths, tmp_path):
    """Конвертация со складом кладёт файл в подпапки и не перезаписывает старый."""
    path = next((p for p in known_paths if p.name.startswith("Кореличи")), None)
    if path is None:
        pytest.skip("нет файла-образца")
    options = ConvertOptions(naming="auto", output_dir=str(tmp_path), sklad="Кореличи")
    first = convert([str(path)], "", options)
    assert first.outputs, first.notes
    today = f"{dt.date.today():%d.%m.%Y}"
    expected = tmp_path / "Кореличи" / "ООО ОМА" / f"{today}.xlsx"
    assert first.outputs[0] == str(expected)
    assert expected.exists()

    second = convert([str(path)], "", options)
    assert second.outputs[0] != first.outputs[0]
    assert expected.exists()  # первый файл не тронут


# ------------------------------------------------ колонка «Новинка» (1.7.0)
def _write_small_vedomost(tmp_path, rows):
    from openpyxl import Workbook

    book = Workbook()
    sheet = book.active
    sheet.append([])
    sheet.append(["Ведомость по остаткам ТМЦ на складах"])
    sheet.append([])
    sheet.append([
        "Склад", "Группа", "Артикул", "Номенклатура", "Ед. изм.",
        "Расход", "Остаток", "Цена", "Цена посл. поставщика", "Розн.", "Последний поставщик",
    ])
    for row in rows:
        sheet.append(list(row))
    path = tmp_path / "Ведомость.xlsx"
    book.save(path)
    return str(path)


def test_novinka_column_in_file(known_paths, tmp_path):
    """Позиция, которую не выбрали в ведомости, пишется словом «Новинка»."""
    from konvektor.vedomost import MATCH_NAME, Matcher, load_vedomost

    path = next((p for p in known_paths if p.name.startswith("Кореличи")), None)
    if path is None:
        pytest.skip("нет файла-образца")
    ved_path = _write_small_vedomost(tmp_path, [
        ["Склад Мир", "ГАЗ", "", "Газ сжижен.в однораз.баллоне WERKER 520см3", "шт", "7", "3,20", "", "", "", 'ООО "ОМА"'],
    ])
    vedomost = load_vedomost(ved_path)
    matcher = Matcher(vedomost, MATCH_NAME, {})

    result = convert(
        [str(path)],
        "",
        ConvertOptions(
            naming="auto",
            output_dir=str(tmp_path / "итог"),
            matcher=matcher,
        ),
    )
    assert result.outputs, result.notes
    sheet = load_workbook(result.outputs[0])["Данные"]
    header = [cell for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    assert "Новинка" in header
    assert header.index("Новинка") == header.index("сумма с НДС") + 1
    # после «сумма с НДС» в файле ничего, кроме «Новинка»
    assert header[header.index("Новинка") + 1:] == []
    column = header.index("Новинка")
    values = [row[column] for row in sheet.iter_rows(min_row=2, values_only=True)]
    assert "Новинка" in values
    # подобралась только «Газ»: у подобранных позиций клетка пустая,
    # у остальных — слово «Новинка»
    assert 0 < values.count("Новинка") < len(values)
    assert all(value in (None, "Новинка") for value in values), values


def test_no_novinka_column_without_vedomost(known_paths, tmp_path):
    """Без сверки в файле прежние 10 колонок — «Новинка» не появляется."""
    path = next((p for p in known_paths if p.name.startswith("Кореличи")), None)
    if path is None:
        pytest.skip("нет файла-образца")
    result = convert([str(path)], "", ConvertOptions(naming="auto", output_dir=str(tmp_path)))
    assert result.outputs, result.notes
    header = list(
        next(load_workbook(result.outputs[0])["Данные"].iter_rows(min_row=1, max_row=1, values_only=True))
    )
    assert "Новинка" not in header
    assert header == list(TEMPLATE_COLUMNS)
