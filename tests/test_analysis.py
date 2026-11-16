"""Метрики разбора: покрытие полей, арифметика, доверие (konvektor/analysis.py)."""

from __future__ import annotations

import pytest

from konvektor import detect
from konvektor.analysis import CONFIDENCE_OK, build_report, col_letter, format_report
from konvektor.extract import SheetExtraction
from konvektor.meta import DocMeta

from synthetic import CANON, VARIANTS, extract


def test_col_letter():
    assert col_letter(0) == "A"
    assert col_letter(25) == "Z"
    assert col_letter(26) == "AA"
    assert col_letter(41) == "AP"


@pytest.mark.parametrize("variant", sorted(VARIANTS))
def test_report_confidence_is_high_on_clean_layouts(variant):
    res = extract(VARIANTS[variant]())
    report = build_report(f"{variant}.xlsx", None, res)
    assert report.items > 0
    assert report.confidence >= CONFIDENCE_OK, (
        f"{variant}: доверие {report.confidence}, причины {report.reasons}"
    )
    assert report.quality == "высокое"
    assert not report.needs_review


def test_report_describes_roles_and_columns():
    res = extract(VARIANTS["classic"]())
    report = build_report("classic.xlsx", None, res)
    assert report.roles_found == 10
    block = report.blocks[0]
    assert block.row == 1
    assert block.roles[detect.NAME] == 2
    assert block.headers[2] == "Наименование товара"
    data = report.as_dict()
    assert data["coverage"] == "10/10"
    assert data["blocks"][0]["columns"]["name"].startswith("C ")
    text = format_report(report)
    assert "колонок распознано: 10/10" in text
    assert "наименование   -> C" in text


def test_report_marks_missing_items_as_low_confidence():
    report = build_report("пусто.xlsx", None, SheetExtraction())
    assert report.items == 0
    assert report.confidence == 0.0
    assert report.quality == "низкое"
    assert report.needs_review
    assert "позиции не найдены" in report.reasons


def test_report_missing_role_lowers_confidence():
    res = extract(VARIANTS["no_barcode_header"]())
    # колонка штрихкода без подписи достраивается по данным -> все 10 ролей
    assert build_report("x.xlsx", None, res).roles_found == 10

    # а вот таблица без артикула и штрихкода — это меньше ролей и ниже доверие
    grid = [["Наименование", "Кол-во", "Цена", "Сумма", "Ставка НДС", "Сумма НДС", "Всего"]]
    for rec in CANON:
        grid.append([rec[2], rec[4], rec[5], rec[6], rec[7], rec[8], rec[9]])
    report = build_report("y.xlsx", None, extract(grid))
    assert report.roles_found == 7
    assert report.confidence < 1.0
    assert any("не найдены колонки" in r for r in report.reasons)


def test_report_includes_meta():
    res = extract(VARIANTS["classic"]())
    meta = DocMeta(supplier='ООО "Ромашка"', doc_type="Счет", doc_number="7")
    report = build_report("z.xlsx", None, res, meta)
    text = format_report(report)
    assert 'ООО "Ромашка"' in text
    assert "Счет № 7" in text


def test_real_samples_are_all_high_confidence(known_paths):
    """Замок качества: все образцы заказчика разбираются с высоким доверием."""
    from konvektor.analysis import build_report
    from konvektor.extract import extract_file
    from konvektor.readers import read_source

    for path in known_paths:
        src = read_source(str(path))
        report = build_report(str(path), src, extract_file(src))
        assert report.items > 0, f"{path.name}: позиции не найдены"
        assert report.confidence >= CONFIDENCE_OK, (
            f"{path.name}: доверие {report.confidence} ({report.quality}); "
            f"причины: {report.reasons}"
        )
