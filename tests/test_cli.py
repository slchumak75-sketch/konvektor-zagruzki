"""Консольный интерфейс: JSON-вывод и режим --diagnose."""

from __future__ import annotations

import datetime as dt
import json

import pytest

from konvektor.cli import main


def _known(known_paths, startswith: str):
    return next((p for p in known_paths if p.name.startswith(startswith)), None)


def _first_json(text: str) -> dict:
    """Первый JSON-объект в выводе (после него печатаются пути файлов)."""
    start = text.find("{")
    assert start >= 0, text
    obj, _end = json.JSONDecoder().raw_decode(text[start:])
    return obj


def test_json_contains_documents(known_paths, tmp_path, capsys):
    path = _known(known_paths, "_№ 10802")
    if path is None:
        pytest.skip("нет файла-образца")
    out = tmp_path / "out.xlsx"
    assert main([str(path), "-o", str(out), "--json", "-q"]) == 0
    payload = _first_json(capsys.readouterr().out)
    assert payload["ok"] is True
    doc = payload["documents"][0]
    assert doc["supplier"] == 'ООО "Культбытхозторгплюс"'
    assert doc["supplier_unp"] == "192553707"
    assert doc["doc_type"] == "Счет"
    assert doc["doc_number"] == "3596616"
    assert doc["doc_date"] == "12.11.2026"
    assert doc["origin"] == "документ"


def test_diagnose_good_file(known_paths, capsys):
    path = _known(known_paths, "КХ2104033")
    if path is None:
        pytest.skip("нет файла-образца")
    assert main(["--diagnose", str(path)]) == 0
    out = capsys.readouterr().out
    assert "Прочитан как: xls" in out
    assert "Позиций найдено: 19" in out
    assert "ОАО" in out


def test_diagnose_missing_file(tmp_path, capsys):
    assert main(["--diagnose", str(tmp_path / "нет.xls")]) == 1
    assert "файл не найден" in capsys.readouterr().out


def test_diagnose_garbage_file(tmp_path, capsys):
    bad = tmp_path / "мусор.xls"
    bad.write_bytes(b"not an excel at all")
    code = main(["--diagnose", str(bad)])
    out = capsys.readouterr().out
    assert code in (0, 1)
    assert "Файл:" in out
    assert "Итог:" in out


def test_explain_prints_roles(known_paths, capsys):
    path = _known(known_paths, "Счет на оплату № УТ-542")
    if path is None:
        pytest.skip("нет файла-образца")
    assert main(["--explain", str(path)]) == 0
    out = capsys.readouterr().out
    assert "колонок распознано:" in out
    assert "наименование" in out and "-> D" in out
    assert "доверие:" in out


def test_json_contains_analysis(known_paths, tmp_path, capsys):
    path = _known(known_paths, "КХ2104033")
    if path is None:
        pytest.skip("нет файла-образца")
    out = tmp_path / "out.xlsx"
    assert main([str(path), "-o", str(out), "--json", "-q"]) == 0
    payload = _first_json(capsys.readouterr().out)
    analysis = payload["analysis"][0]
    assert analysis["items"] == 19
    assert analysis["coverage"].endswith("/10")
    assert analysis["quality"] in ("высокое", "среднее", "низкое")
    assert analysis["blocks"][0]["header_row"] == 10


def test_output_is_folder_and_name_is_supplier_date(known_paths, tmp_path, capsys):
    """-o — это папка: имя файла всегда «поставщик дата.xlsx», а не «Результат.xlsx»."""
    path = _known(known_paths, "Кореличи")
    if path is None:
        pytest.skip("нет файла-образца")
    # вариант 1: указали папку
    assert main([str(path), "-o", str(tmp_path), "--json", "-q"]) == 0
    payload = _first_json(capsys.readouterr().out)
    expected = str(tmp_path / f"ООО ОМА {dt.date.today():%d.%m.%Y}.xlsx")
    assert payload["outputs"] == [expected], payload["outputs"]
    # вариант 2: по привычке указали путь к файлу — имя всё равно своё
    other = tmp_path / "вторая"
    other.mkdir()
    assert main([str(path), "-o", str(other / "Результат.xlsx"), "--json", "-q"]) == 0
    payload = _first_json(capsys.readouterr().out)
    assert payload["outputs"] == [str(other / f"ООО ОМА {dt.date.today():%d.%m.%Y}.xlsx")]
