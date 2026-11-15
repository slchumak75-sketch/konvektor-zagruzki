"""Консольный интерфейс: JSON-вывод и режим --diagnose."""

from __future__ import annotations

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
