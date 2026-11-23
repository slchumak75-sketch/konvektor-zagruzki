"""Проверка настоящих Tk-виджетов в Windows CI; не заменяет визуальную проверку."""
from konvektor.extract import DocTotal, Item
from konvektor.gui import KonvertorApp
from konvektor.pipeline import ConvertResult
from konvektor.validate import validate_items


def main():
    app = KonvertorApp()
    try:
        app.update()
        item = Item(source="scan.pdf", sheet="PDF 1", row=1, name="Товар", unit="шт",
                    qty=2, price=5, total_sum=10, vat_rate=20, vat_sum=2, sum_with_vat=12,
                    needs_confirmation={"name", "qty"})
        total = DocTotal(source="scan.pdf", sheet="PDF 1", row=2, label="Итого (OCR)",
                         values={"qty": 2, "sum": 10, "vat_sum": 2, "total": 12}, needs_confirmation=True)
        result = ConvertResult(items=[item], report=validate_items([item]), document_totals=[total],
                               ocr_source_hashes={"scan.pdf": "hash"})
        app._finish(result)
        app.update()
        rid = next(iter(app._preview_rows))
        app._refresh_barcode_highlights()
        assert (rid, "количество") in app._issue_highlights, "Подсветка OCR не показана"
        assert (rid, "Из документа") in app._issue_highlights, "Подсветка названия не показана"
        assert app._total_editor_vars, "Поля итога OCR отсутствуют"
        app.edit_source_cell(rid, "количество")
        app.update()
        assert app._editor is not None and app._editor.winfo_viewable(), "Редактор ячейки не показан"
        app._cancel_edit()
        print("Windows Tk GUI: OK")
    finally:
        app.destroy()


if __name__ == "__main__":
    main()
