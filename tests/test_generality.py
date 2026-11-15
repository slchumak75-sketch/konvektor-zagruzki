"""Общность разбора: одни и те же данные из разных раскладок (синтетика).

В отличие от `test_samples.py`, здесь нет файлов заказчика — раскладки
генерируются кодом (см. `tests/synthetic.py`), поэтому тест проверяет общий
алгоритм, а не совпадение с конкретным образцом.
"""

from __future__ import annotations

import pytest

from konvektor.validate import validate_items

from synthetic import CANON, DOUBLED, VARIANTS, extract


@pytest.mark.parametrize("variant", sorted(VARIANTS))
def test_layout_is_recognised(variant):
    res = extract(VARIANTS[variant]())
    expected_rows = len(CANON) * (2 if variant in DOUBLED else 1)
    assert not res.warnings, f"{variant}: неожиданные замечания {res.warnings}"
    assert len(res.items) == expected_rows, (
        f"{variant}: найдено {len(res.items)} позиций вместо {expected_rows}"
    )


@pytest.mark.parametrize("variant", sorted(VARIANTS))
def test_all_fields_recovered(variant):
    res = extract(VARIANTS[variant]())
    assert res.items, f"{variant}: позиции не найдены"
    expected = CANON * (2 if variant in DOUBLED else 1)
    for item, rec in zip(res.items, expected):
        barcode, article, name, unit, qty, price, total_sum, vat_rate, vat_sum, total = rec
        assert item.name.startswith(name[:20]), f"{variant}: «{item.name}» != «{name}»"
        assert item.unit == unit, f"{variant}: ед. изм. {item.unit!r} != {unit!r}"
        assert item.qty == pytest.approx(qty), f"{variant}: кол-во {item.qty} != {qty}"
        assert item.price == pytest.approx(price), f"{variant}: цена {item.price} != {price}"
        assert item.total_sum == pytest.approx(total_sum), f"{variant}: сумма {item.total_sum}"
        assert item.vat_rate == pytest.approx(vat_rate), f"{variant}: ставка {item.vat_rate}"
        assert item.vat_sum == pytest.approx(vat_sum), f"{variant}: НДС {item.vat_sum}"
        assert item.sum_with_vat == pytest.approx(total), f"{variant}: с НДС {item.sum_with_vat}"
        if barcode:
            assert item.barcode == barcode, f"{variant}: штрихкод {item.barcode!r}"
        if article:
            assert item.article == article, f"{variant}: артикул {item.article!r}"


@pytest.mark.parametrize("variant", sorted(VARIANTS))
def test_math_is_consistent(variant):
    """Данные из любой раскладки проходят арифметическую проверку без ошибок."""
    res = extract(VARIANTS[variant]())
    report = validate_items(res.items)
    assert report.errors == [], (
        f"{variant}: расхождения {[e.message for e in report.errors]}"
    )
