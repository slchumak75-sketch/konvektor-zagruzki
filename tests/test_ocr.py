from PIL import Image

from konvektor.ocr import prepare_cell, recognize_cell


class FakeEngine:
    def __init__(self, text, confidence):
        self.text, self.confidence = text, confidence

    def recognize(self, image, numeric):
        assert image.mode == "L"
        return self.text, self.confidence


def test_cell_is_preprocessed_and_given_padding():
    image = Image.new("RGB", (10, 8), (190, 210, 210))
    result = prepare_cell(image)
    assert result.size == (80, 74)
    assert result.getpixel((0, 0)) == 255


def test_ocr_always_requires_human_confirmation_even_if_confident():
    result = recognize_cell(Image.new("RGB", (10, 8)), FakeEngine("49,71", 99), numeric=True)
    assert result.text == "49,71"
    assert result.requires_confirmation


def test_unparseable_number_never_silently_changes_to_number():
    result = recognize_cell(Image.new("RGB", (10, 8)), FakeEngine("12x5", 90), numeric=True)
    assert result.text == "???"
    assert result.requires_confirmation


def test_unrecognized_name_is_marked():
    result = recognize_cell(Image.new("RGB", (10, 8)), FakeEngine("   ", 0))
    assert result.text == "???"


def test_detect_grid_on_synthetic_table():
    import pytest
    pytest.importorskip("numpy")
    pytest.importorskip("cv2")
    from PIL import ImageDraw
    from konvektor.ocr import detect_grid

    image = Image.new("RGB", (600, 400), "white")
    draw = ImageDraw.Draw(image)
    for x in (40, 180, 330, 550):
        draw.line((x, 30, x, 350), fill="black", width=2)
    for y in (30, 100, 190, 270, 350):
        draw.line((40, y, 550, y), fill="black", width=2)
    grid = detect_grid(image, (30, 20, 560, 360))
    assert len(grid.vertical) == 4
    assert len(grid.horizontal) == 5
    assert all(abs(a-b) <= 2 for a, b in zip(grid.vertical, (40, 180, 330, 550)))
    assert all(abs(a-b) <= 2 for a, b in zip(grid.horizontal, (30, 100, 190, 270, 350)))
    assert grid.cell(0, 0)[0] > 40


def test_detect_grid_rejects_invalid_area():
    import pytest
    pytest.importorskip("numpy")
    pytest.importorskip("cv2")
    from konvektor.ocr import detect_grid
    with pytest.raises(ValueError):
        detect_grid(Image.new("RGB", (200, 200)), (0, 0, 201, 100))


def test_deskew_preserves_page_and_recovers_rotation():
    import pytest
    pytest.importorskip("numpy")
    pytest.importorskip("cv2")
    from PIL import ImageDraw
    from konvektor.ocr import deskew_page
    image = Image.new("RGB", (600, 400), "white")
    draw = ImageDraw.Draw(image)
    for y in range(60, 350, 50):
        draw.line((50, y, 550, y), fill="black", width=2)
    tilted = image.rotate(2, expand=True, fillcolor="white")
    corrected, angle = deskew_page(tilted)
    assert -2.5 < angle < -1.5
    assert corrected.width >= tilted.width
    assert corrected.height >= tilted.height


def test_deskew_without_lines_fails_explicitly():
    import pytest
    pytest.importorskip("numpy")
    pytest.importorskip("cv2")
    from konvektor.ocr import deskew_page
    with pytest.raises(ValueError, match="наклон"):
        deskew_page(Image.new("RGB", (300, 300), "white"))


def test_find_ttn_table_region_on_synthetic_page():
    import pytest
    pytest.importorskip("numpy")
    pytest.importorskip("cv2")
    from PIL import ImageDraw
    from konvektor.ocr import find_ttn_region
    image = Image.new("RGB", (1000, 1200), "white")
    draw = ImageDraw.Draw(image)
    xs = (60, 203, 256, 329, 405, 479, 527, 599, 671, 746, 818, 940)
    ys = (250, 370, 490, 610, 730, 850)
    for x in xs:
        draw.line((x, ys[0], x, ys[-1]), fill="black", width=2)
    for y in ys:
        draw.line((xs[0], y, xs[-1], y), fill="black", width=2)
    region = find_ttn_region(image)
    assert len(region.lines.vertical) == len(xs)
    assert len(region.lines.horizontal) == len(ys)
    assert all(abs(a-b) < 4 for a,b in zip(region.lines.vertical, xs))
    assert all(abs(a-b) < 4 for a,b in zip(region.lines.horizontal, ys))


def test_unsupported_grid_not_assumed_to_be_ttn():
    import pytest
    pytest.importorskip("numpy")
    pytest.importorskip("cv2")
    from PIL import ImageDraw
    from konvektor.ocr import find_ttn_region
    image = Image.new("RGB", (1000, 1200), "white")
    draw = ImageDraw.Draw(image)
    for x in (100, 300, 600, 900):
        draw.line((x, 200, x, 1000), fill="black", width=2)
    for y in (200, 400, 600, 800, 1000):
        draw.line((100, y, 900, y), fill="black", width=2)
    with pytest.raises(ValueError, match="не найдена"):
        find_ttn_region(image)


def test_ttn_rows_keep_candidates_for_user_verification(monkeypatch):
    from konvektor import ocr
    # Две строки шапки, товар и итог, без зависимости от OCR-движка.
    xs = tuple(range(10, 610, 50))
    ys = (10, 50, 90, 180, 220)
    monkeypatch.setattr(ocr, "find_ttn_region", lambda _im: ocr.TableRegion((10, 10, 560, 220), ocr.GridLines(xs, ys)))
    class Engine:
        def __init__(self):
            self.names = iter(["Наименование товара", "1", "Товар", "ИТЭГО"])
        def recognize(self, _image, numeric):
            if not numeric:
                return next(self.names, "шт"), 95
            return "12,00", 95
    rows = ocr.read_ttn_rows(Image.new("RGB", (620, 240), "white"), Engine())
    assert [row.kind for row in rows] == ["header", "header", "item", "total"]
    assert rows[2].cells["qty"].text == "12,00"
    assert all(c.requires_confirmation for row in rows for c in row.cells.values())
    assert set(rows[3].cells) == {"name", "qty", "total_sum", "vat_sum", "sum_with_vat"}


def test_ocr_extraction_never_trusts_candidates_or_totals():
    from konvektor.ocr import CellCandidate, TableRowCandidate, candidates_to_extraction
    from konvektor.validate import validate_items, check_document_totals
    c = lambda value: CellCandidate(value, 99)
    rows = [TableRowCandidate("item", {"name": c("Товар"), "qty": c("2"), "price": c("5"),
                                      "total_sum": c("???")}, 0, 40),
            TableRowCandidate("total", {"qty": c("2"), "total_sum": c("10")}, 40, 60)]
    extraction = candidates_to_extraction("1.pdf", "PDF 1", rows)
    item = extraction.items[0]
    assert item.qty == 2
    assert item.total_sum is None and "total_sum" in item.unrecognized
    assert item.needs_confirmation == {"name", "qty", "price", "total_sum", "barcode"}
    assert "barcode" in item.unrecognized
    report = validate_items(extraction.items)
    assert item.total_sum is None
    check_document_totals(report, extraction.totals, {("1.pdf", "PDF 1", 0): [item]})
    assert any("требуют подтверждения" in issue.message for issue in report.warnings)


def test_ocr_does_not_rederive_barcode_from_name_after_user_rejection():
    from konvektor.ocr import CellCandidate, TableRowCandidate, candidates_to_extraction
    c = lambda value: CellCandidate(value, 95)
    row = TableRowCandidate("item", {"name": c("Товар 4811221006186")}, 0, 20)
    item = candidates_to_extraction("1.pdf", "PDF 1", [row]).items[0]
    assert not item.barcode_from_name_allowed
    assert "barcode" in item.needs_confirmation
