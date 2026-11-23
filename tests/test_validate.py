"""Тесты математических проверок и «достройки» строк."""

from __future__ import annotations

import pytest

from konvektor.extract import DocTotal, Item
from konvektor.validate import (
    EXCLUSIVE,
    INCLUSIVE,
    check_document_totals,
    detect_vat_mode,
    validate_items,
)


def make_item(**kwargs) -> Item:
    base = dict(source="f.xlsx", sheet="Лист", row=2, name="Товар")
    base.update(kwargs)
    return Item(**base)


def test_valid_row_no_issues():
    it = make_item(qty=2, price=10.0, total_sum=20.0, vat_rate=20.0, vat_sum=4.0, sum_with_vat=24.0)
    report = validate_items([it])
    assert report.errors == []


def test_sum_mismatch_is_error():
    it = make_item(qty=2, price=10.0, total_sum=25.0, vat_rate=20.0, vat_sum=5.0, sum_with_vat=30.0)
    report = validate_items([it])
    errors = [e for e in report.errors if e.field == "сумма"]
    assert errors
    assert "≠ количество × цена" in errors[0].message
    assert errors[0].doc_value == 25.0
    assert errors[0].calc_value == 20.0


def test_total_mismatch_is_error():
    it = make_item(qty=1, price=10.0, total_sum=10.0, vat_rate=20.0, vat_sum=2.0, sum_with_vat=13.0)
    report = validate_items([it])
    assert any(e.field == "сумма с НДС" for e in report.errors)


def test_vat_by_rate_mismatch_is_error():
    it = make_item(qty=1, price=100.0, total_sum=100.0, vat_rate=20.0, vat_sum=10.0, sum_with_vat=110.0)
    report = validate_items([it])
    assert any(e.field == "сумма НДС" for e in report.errors)


def test_missing_sum_is_calculated():
    it = make_item(qty=6, price=10.73, vat_rate=20.0, sum_with_vat=77.26)
    report = validate_items([it])
    assert it.total_sum == pytest.approx(64.38)
    assert it.vat_sum == pytest.approx(12.88)
    assert "сумма" in it.computed
    assert "сумма НДС" in it.computed
    assert all(e.severity == "info" for e in report.issues)


def test_missing_vat_from_total():
    it = make_item(qty=1, price=10.0, total_sum=10.0, sum_with_vat=12.0)
    validate_items([it])
    assert it.vat_sum == pytest.approx(2.0)
    assert it.vat_rate == pytest.approx(20.0)


def test_missing_price_and_qty_are_calculated():
    it = make_item(total_sum=50.0, price=12.5)
    validate_items([it])
    assert it.qty == pytest.approx(4.0)
    assert "количество" in it.computed

    it2 = make_item(total_sum=50.0, qty=4.0)
    validate_items([it2])
    assert it2.price == pytest.approx(12.5)
    assert "цена" in it2.computed


def test_inclusive_vat_mode_detected():
    items = [
        make_item(row=2, qty=1, price=120.0, total_sum=120.0, vat_rate=20.0, vat_sum=20.0, sum_with_vat=120.0),
        make_item(row=3, qty=1, price=60.0, total_sum=60.0, vat_rate=20.0, vat_sum=10.0, sum_with_vat=60.0),
    ]
    assert detect_vat_mode(items) == INCLUSIVE
    report = validate_items(items)
    assert report.errors == []


def test_exclusive_vat_mode_detected():
    items = [
        make_item(row=2, qty=1, price=100.0, total_sum=100.0, vat_rate=20.0, vat_sum=20.0, sum_with_vat=120.0),
        make_item(row=3, qty=1, price=50.0, total_sum=50.0, vat_rate=20.0, vat_sum=10.0, sum_with_vat=60.0),
    ]
    assert detect_vat_mode(items) == EXCLUSIVE


def test_inclusive_mode_used_for_calculation():
    items = [
        make_item(row=2, qty=1, price=120.0, total_sum=120.0, vat_rate=20.0, vat_sum=20.0, sum_with_vat=120.0),
        make_item(row=3, qty=1, price=60.0, total_sum=60.0, vat_rate=20.0, sum_with_vat=60.0),
    ]
    validate_items(items)
    # НДС «в том числе»: 60 × 20/120 = 10
    assert items[1].vat_sum == pytest.approx(10.0)


def test_rounding_within_tolerance():
    # 3 × 0,333 = 0,999 ≈ 1,00 — расхождение в копейку допустимо
    it = make_item(qty=3, price=0.333, total_sum=1.0, vat_rate=20.0, vat_sum=0.2, sum_with_vat=1.2)
    report = validate_items([it])
    assert report.errors == []


def test_nonstandard_vat_rate_is_warning():
    it = make_item(qty=1, price=100.0, total_sum=100.0, vat_rate=12.5, vat_sum=12.5, sum_with_vat=112.5)
    report = validate_items([it])
    assert any(e.severity == "warning" and e.field == "ставка НДС" for e in report.issues)


def test_document_total_mismatch_detected():
    items = [make_item(qty=1, price=100.0, total_sum=100.0, vat_rate=20.0, vat_sum=20.0, sum_with_vat=120.0)]
    report = validate_items(items)
    blocks = {("f.xlsx", "Лист", 0): items}
    total = DocTotal(
        source="f.xlsx",
        sheet="Лист",
        row=10,
        label="Итого:",
        scope="document",
        values={"total": 999.0},
    )
    check_document_totals(report, [total], blocks)
    assert any("итог документа" in i.message for i in report.warnings)


def test_document_total_match_no_warning():
    items = [make_item(qty=1, price=100.0, total_sum=100.0, vat_rate=20.0, vat_sum=20.0, sum_with_vat=120.0)]
    report = validate_items(items)
    blocks = {("f.xlsx", "Лист", 0): items}
    total = DocTotal(
        source="f.xlsx", sheet="Лист", row=10, label="Итого:", scope="document",
        values={"total": 120.0, "sum": 100.0, "vat_sum": 20.0},
    )
    check_document_totals(report, [total], blocks)
    assert report.warnings == []


def test_words_row_ignored():
    """Строка «прописью» не должна давать ложных сверок."""
    items = [make_item(qty=1, price=100.0, total_sum=100.0, vat_rate=20.0, vat_sum=20.0, sum_with_vat=120.0)]
    report = validate_items(items)
    blocks = {("f.xlsx", "Лист", 0): items}
    total = DocTotal(
        source="f.xlsx", sheet="Лист", row=10, label="Всего сумма НДС", scope="document",
        values={}, numbers=[],
    )
    check_document_totals(report, [total], blocks)
    assert report.warnings == []


def test_missing_values_cannot_be_counted_as_zero_for_totals():
    from konvektor.validate import ValidationReport
    items = [make_item(row=2, total_sum=5), make_item(row=3, total_sum=None)]
    report = ValidationReport(items=items)
    total = DocTotal(source="f.xlsx", sheet="Лист", row=10, label="Итого:", values={"sum": 5})
    check_document_totals(report, [total], {("f.xlsx", "Лист", 0): items})
    assert any("не проверен" in issue.message and issue.field == "сумма" for issue in report.warnings)


def test_explicitly_grouped_pages_compare_all_lines():
    from konvektor.validate import ValidationReport
    first = make_item(source="1.pdf", row=2, qty=60)
    second = make_item(source="2.pdf", row=2, qty=14)
    total = DocTotal(source="2.pdf", sheet="Лист", row=10, label="Итого:", values={"qty": 74})
    blocks = {("1.pdf", "Лист", 0): [first], ("2.pdf", "Лист", 0): [second]}
    report = ValidationReport(items=[first, second])
    check_document_totals(report, [total], blocks, prior_sources={"2.pdf": ["1.pdf"]})
    assert not report.warnings
    first.qty = 59
    report = ValidationReport(items=[first, second])
    check_document_totals(report, [total], blocks, prior_sources={"2.pdf": ["1.pdf"]})
    assert any(i.doc_value == 74 and i.calc_value == 73 for i in report.warnings)


def test_group_missing_page_never_reports_match():
    from konvektor.validate import ValidationReport
    second = make_item(source="2.pdf", qty=74)
    total = DocTotal(source="2.pdf", sheet="Лист", row=10, label="Итого:", values={"qty": 74})
    report = ValidationReport(items=[second])
    check_document_totals(report, [total], {("2.pdf", "Лист", 0): [second]}, prior_sources={"2.pdf": ["1.pdf"]})
    assert any("не все выбранные страницы" in i.message for i in report.warnings)


def test_unrecognized_cells_are_not_filled_by_arithmetic():
    item = make_item(qty=2, price=10, total_sum=None, vat_rate=20, vat_sum=4, sum_with_vat=24,
                     unrecognized={"total_sum"})
    validate_items([item])
    assert item.total_sum is None
    assert "сумма" not in item.computed
