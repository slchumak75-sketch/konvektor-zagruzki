"""Консольный интерфейс (Konvertor-cli.exe / python -m konvektor.cli)."""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import List, Optional

from . import APP_NAME, __version__
from .vedomost import MATCH_ARTICLE, MATCH_BARCODE, MATCH_NAME, MATCH_TITLES, Matcher, load_vedomost
from .pipeline import (
    MODE_COMBINED,
    MODE_FILES,
    MODE_SHEETS,
    ConvertOptions,
    convert,
    diagnose,
    explain,
    is_supported,
    selftest,
)


def _safe_streams() -> None:
    """Чтобы русский текст не падал в консоли Windows с cp866."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001 - не критично
            pass


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="Konvertor-cli",
        description=f"{APP_NAME} {__version__} — конвертер счетов/накладных/прайсов "
        "в единую Excel-таблицу по образцу.",
        epilog="Пример: Konvertor-cli.exe счет.xls ТТН.xlsx -o Загрузка.xlsx",
    )
    p.add_argument("files", nargs="*", help="файлы и/или папки для конвертации")
    p.add_argument(
        "-o",
        "--output",
        help="папка для результата (имя файла всегда «поставщик дата.xlsx»; "
        "по умолчанию — папка «Документы»)",
    )
    p.add_argument(
        "-m",
        "--mode",
        choices=(MODE_COMBINED, MODE_SHEETS, MODE_FILES),
        default=MODE_COMBINED,
        help="combined — всё в один лист; sheets — по листу на файл; files — отдельный xlsx на файл",
    )
    p.add_argument("-r", "--recursive", action="store_true", help="обходить вложенные папки")
    p.add_argument("-q", "--quiet", action="store_true", help="не печатать прогресс")
    p.add_argument("--json", action="store_true", help="напечатать итог в JSON")
    p.add_argument(
        "--source-columns",
        action="store_true",
        help="добавить в таблицу колонки «файл» и «строка» (в окне они добавляются всегда)",
    )
    p.add_argument(
        "--vedomost",
        metavar="ФАЙЛ",
        help="ведомость остатков ТМЦ: сверить номенклатуру и переименовать позиции",
    )
    p.add_argument(
        "--match",
        choices=(MATCH_BARCODE, MATCH_ARTICLE, MATCH_NAME),
        default=MATCH_NAME,
        help="чем сверять с ведомостью: barcode — штрихкод, article — артикул, "
        "name — наименование (по умолчанию)",
    )
    p.add_argument(
        "--sklad",
        choices=("Мир", "Кореличи"),
        default="",
        help="склад заказчика: результат ляжет в <папка>/<склад>/<поставщик>/<дата>.xlsx",
    )
    p.add_argument(
        "--supplier",
        metavar="ИМЯ",
        default="",
        help="поставщик для папки и сверки (по умолчанию — как в документе)",
    )
    p.add_argument("--strict", action="store_true", help="вернуть код 3, если найдены расхождения в суммах")
    p.add_argument("--no-validate", action="store_true", help="не проверять арифметику (только перенос данных)")
    p.add_argument("--version", action="version", version=f"{APP_NAME} {__version__}")
    p.add_argument("--ocr-selftest", action="store_true", help="проверить наличие движка и моделей OCR")
    p.add_argument(
        "--explain",
        nargs="+",
        metavar="ФАЙЛ",
        help="показать, как разобран файл: шапка, роли колонок, арифметика, доверие",
    )
    p.add_argument(
        "--diagnose",
        metavar="ФАЙЛ",
        help="подробно разобрать один файл (почему не читается, что найдено)",
    )
    p.add_argument(
        "--selftest",
        nargs="?",
        const=".",
        metavar="ПАПКА",
        help="самопроверка: сконвертировать все файлы папки во временный файл (для CI)",
    )
    return p


def main(argv: Optional[List[str]] = None) -> int:
    _safe_streams()
    args = build_parser().parse_args(argv)

    log = (lambda msg: None) if args.quiet else (lambda msg: print(msg, flush=True))

    if args.ocr_selftest:
        from PIL import Image
        from .ocr import TesseractCellEngine, recognize_cell
        try:
            import cv2
            import numpy
            base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(__file__)))
            tessdata = os.environ.get("KONVEKTOR_TESSDATA", os.path.join(base, "tessdata"))
            with TesseractCellEngine(tessdata) as engine:
                recognize_cell(Image.new("RGB", (100, 40), "white"), engine, numeric=True)
            print(f"OCR: OK (OpenCV {cv2.__version__}, NumPy {numpy.__version__})")
            return 0
        except Exception as exc:
            print(f"OCR: FAIL — {exc}", file=sys.stderr)
            return 1

    if args.selftest is not None:
        ok, msg = selftest(args.selftest, log=log)
        print(("САМОПРОВЕРКА: OK — " if ok else "САМОПРОВЕРКА: FAIL — ") + msg)
        return 0 if ok else 1

    if args.explain:
        for text in explain(args.explain, log=lambda m: None):
            print(text)
            print()
        return 0

    if args.diagnose:
        ok, lines = diagnose(args.diagnose, log=lambda m: None)
        for line in lines:
            print(line)
        return 0 if ok else 1

    if not args.files:
        build_parser().print_help()
        return 2

    files = [
        f for f in args.files if os.path.isdir(f) or (os.path.isfile(f) and is_supported(f))
    ]
    for f in args.files:
        if os.path.isfile(f) and not is_supported(f):
            print(f"пропущен (неподдерживаемый формат): {f}", file=sys.stderr)
        elif not os.path.exists(f):
            print(f"пропущен (файл не найден): {f}", file=sys.stderr)

    # -o — это папка: имя файла всегда «поставщик дата.xlsx», иначе легко
    # получить безликое «Результат.xlsx». Если передали путь к .xlsx, берём его папку.
    out = args.output or ""
    if os.path.splitext(out)[1].lower() in (".xlsx", ".xlsm", ".xls"):
        out = os.path.dirname(out)
    matcher = None
    if args.vedomost:
        try:
            vedomost = load_vedomost(args.vedomost)
        except ValueError as exc:
            print("ОШИБКА ведомости: " + str(exc), file=sys.stderr)
            return 2
        matcher = Matcher(vedomost, args.match, supplier=args.supplier)
        log(
            f"ведомость: {os.path.basename(args.vedomost)} — позиций {vedomost.count}, "
            f"поставщиков {len(vedomost.suppliers)}; сверка {MATCH_TITLES[matcher.strategy]}"
        )

    options = ConvertOptions(
        mode=args.mode,
        validate=not args.no_validate,
        # раскладываем как «<папка>/поставщик дата.xlsx»
        naming="auto",
        output_dir=out,
        add_source_columns=args.source_columns,
        matcher=matcher,
        sklad=args.sklad,
        supplier=args.supplier,
    )
    result = convert(files, "", options, log=log, recursive=args.recursive)

    if args.json:
        payload = {
            "ok": result.ok,
            "outputs": result.outputs,
            "items": result.items_count,
            "errors": [
                {"source": i.source, "row": i.row, "field": i.field, "message": i.message}
                for i in (result.report.errors if result.report else [])
            ],
            "warnings": [
                {"source": i.source, "row": i.row, "field": i.field, "message": i.message}
                for i in (result.report.warnings if result.report else [])
            ],
            "computed": result.computed_count,
            "documents": [
                {
                    "file": path,
                    "supplier": meta.supplier,
                    "supplier_unp": meta.supplier_unp,
                    "doc_type": meta.doc_type,
                    "doc_number": meta.doc_number,
                    "doc_date": meta.date_text,
                    "origin": meta.origin,
                    "warnings": meta.warnings,
                }
                for path, meta in result.metas.items()
            ],
            "match": result.match_summary,
            "analysis": [report.as_dict() for report in result.analysis.values()],
            "stats": [
                {
                    "file": s.path,
                    "items": s.items,
                    "sum": s.sum_without_vat,
                    "vat": s.vat_sum,
                    "total": s.sum_with_vat,
                    "errors": s.errors,
                    "warnings": s.warnings,
                }
                for s in result.stats
            ],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))

    if matcher is not None:
        print(matcher.summary_text(), file=sys.stderr)
    if result.errors:
        for e in result.errors:
            print("ОШИБКА: " + e, file=sys.stderr)
    if result.report:
        for issue in result.report.errors:
            print(f"РАСХОЖДЕНИЕ: {os.path.basename(issue.source)} стр.{issue.row}: {issue.message}", file=sys.stderr)

    for path in result.outputs:
        print(path)

    if not result.items:
        return 1
    if args.strict and result.error_count:
        return 3
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
