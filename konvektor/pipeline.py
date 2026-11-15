"""Оркестрация: файлы -> позиции -> проверки -> итоговая книга.

Единая точка входа для GUI и CLI. Ничего не знает про интерфейс, поэтому
используется в тестах напрямую (tests/test_pipeline.py).
"""

from __future__ import annotations

import os
import tempfile
import traceback
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from .extract import Item, extract_file
from .readers import read_source
from .validate import ValidationReport, check_document_totals, validate_items
from .writer import FileStat, WorkbookPayload, write_workbook

SUPPORTED_EXTENSIONS = (
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".csv",
    ".txt",
    ".tsv",
    ".html",
    ".htm",
    ".mht",
    ".mhtml",
)

MODE_COMBINED = "combined"  # все файлы в один лист «Данные»
MODE_SHEETS = "sheets"      # каждый файл — отдельный лист в одной книге
MODE_FILES = "files"        # каждый файл — отдельный xlsx

MODE_TITLES = OrderedDict(
    [
        (MODE_COMBINED, "Все файлы в один лист «Данные»"),
        (MODE_SHEETS, "Каждый файл — отдельный лист одной книги"),
        (MODE_FILES, "Каждый файл — отдельный Excel-файл"),
    ]
)

Logger = Callable[[str], None]


@dataclass
class ConvertOptions:
    mode: str = MODE_COMBINED
    validate: bool = True
    keep_warnings: bool = True
    #: добавить в итоговую таблицу колонки «файл-источник» и «строка»
    add_source_columns: bool = False


@dataclass
class ConvertResult:
    outputs: List[str] = field(default_factory=list)
    items: List[Item] = field(default_factory=list)
    report: Optional[ValidationReport] = None
    stats: List[FileStat] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    @property
    def items_count(self) -> int:
        return len(self.items)

    @property
    def error_count(self) -> int:
        return len(self.report.errors) if self.report else 0

    @property
    def warning_count(self) -> int:
        return len(self.report.warnings) if self.report else 0

    @property
    def computed_count(self) -> int:
        if not self.report:
            return 0
        return sum(len(i.computed) for i in self.items)

    @property
    def ok(self) -> bool:
        return bool(self.items) and not self.errors


def _noop(msg: str) -> None:  # pragma: no cover - значение по умолчанию
    pass


def is_supported(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in SUPPORTED_EXTENSIONS


def collect_files(paths: Iterable[str], recursive: bool = False) -> Tuple[List[str], List[str]]:
    """Развернуть аргументы (файлы и папки) в список файлов."""
    files: List[str] = []
    skipped: List[str] = []
    seen = set()
    for p in paths:
        if os.path.isdir(p):
            walker = os.walk(p) if recursive else [(p, [], os.listdir(p))]
            for root, _dirs, names in walker:
                for name in sorted(names):
                    full = os.path.join(root, name)
                    if os.path.isfile(full) and is_supported(full):
                        if full not in seen:
                            seen.add(full)
                            files.append(full)
        elif os.path.isfile(p):
            if p not in seen:
                seen.add(p)
                files.append(p)
        else:
            skipped.append(p)
    return files, skipped


def convert(
    paths: Sequence[str],
    out_path: str,
    options: Optional[ConvertOptions] = None,
    log: Logger = _noop,
    recursive: bool = False,
) -> ConvertResult:
    """Конвертировать файлы в итоговую книгу.

    out_path — путь к .xlsx (для режима MODE_FILES — папка или базовое имя).
    """
    options = options or ConvertOptions()
    result = ConvertResult()

    files, missing = collect_files(paths, recursive=recursive)
    for m in missing:
        result.errors.append(f"файл не найден: {m}")
    if not files:
        result.notes.append("Не найдено ни одного подходящего файла")
        return result

    all_items: List[Item] = []
    totals = []
    blocks: Dict[Tuple[str, str, int], List[Item]] = {}
    per_file_items: List[Tuple[str, List[Item]]] = []

    for path in files:
        log(f"Читаю: {os.path.basename(path)}")
        try:
            src = read_source(path)
            extracted = extract_file(src)
        except Exception as exc:  # noqa: BLE001 - один битый файл не должен рушить всё
            msg = f"{os.path.basename(path)}: не удалось прочитать ({exc})"
            result.errors.append(msg)
            log("  ! " + msg)
            if os.environ.get("KONVEKTOR_DEBUG"):
                log(traceback.format_exc())
            continue

        sheet_names = ", ".join(s.name for s in src.sheets)
        log(
            f"  формат: {src.reader}; листов: {len(src.sheets)}; "
            f"позиций: {len(extracted.items)}"
        )
        for w in extracted.warnings:
            result.notes.append(w)
            log("  ! " + w)

        all_items.extend(extracted.items)
        totals.extend(extracted.totals)
        for it in extracted.items:
            blocks.setdefault((it.source, it.sheet, it.block), []).append(it)
        per_file_items.append((path, extracted.items))
        result.notes.append(f"{os.path.basename(path)}: листы [{sheet_names}], формат {src.reader}")
        if not extracted.items:
            result.notes.append(f"{os.path.basename(path)}: позиции не найдены")
            log("  ! позиции не найдены")

    if not all_items:
        result.notes.append("Позиции товаров не найдены — проверьте структуру файлов")
        return result

    report = validate_items(all_items) if options.validate else ValidationReport(items=all_items)
    check_document_totals(report, totals, blocks)
    result.report = report
    result.items = all_items

    # ---- статистика по файлам -------------------------------------------
    stats: List[FileStat] = []
    for path, items in per_file_items:
        if not items:
            continue
        st = FileStat(path=path, items=len(items))
        st.sum_without_vat = round(sum(i.total_sum or 0.0 for i in items), 2)
        st.vat_sum = round(sum(i.vat_sum or 0.0 for i in items), 2)
        st.sum_with_vat = round(sum(i.sum_with_vat or 0.0 for i in items), 2)
        st.computed = sum(len(i.computed) for i in items)
        st.errors = sum(1 for i in report.issues if i.source == path and i.severity == "error")
        st.warnings = sum(1 for i in report.issues if i.source == path and i.severity == "warning")
        st.vat_mode = report.vat_mode.get(path, "")
        for note in result.notes:
            if note.startswith(os.path.basename(path) + ": листы"):
                st.sheets = note.split("[", 1)[1].split("]", 1)[0]
                st.reader = note.rsplit("формат ", 1)[-1]
                break
        stats.append(st)
    result.stats = stats

    # ---- запись ----------------------------------------------------------
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)

    if options.mode == MODE_FILES:
        base_dir = os.path.dirname(os.path.abspath(out_path)) or "."
        stem = os.path.splitext(os.path.basename(out_path))[0]
        for path, items in per_file_items:
            if not items:
                continue
            name = _unique_path(base_dir, stem, path)
            payload = WorkbookPayload(
                data_sheets=[("Данные", items)],
                report=report,
                stats=stats,
                notes=result.notes,
                with_source=options.add_source_columns,
            )
            result.outputs.append(write_workbook(name, payload))
            log(f"  записан: {name}")
    else:
        if options.mode == MODE_SHEETS:
            sheets = [(os.path.splitext(os.path.basename(p))[0], items) for p, items in per_file_items]
        else:
            sheets = [("Данные", all_items)]
        payload = WorkbookPayload(
            data_sheets=sheets,
            report=report,
            stats=stats,
            notes=result.notes,
            with_source=options.add_source_columns,
        )
        result.outputs.append(write_workbook(out_path, payload))
        log(f"  записан: {out_path}")

    log(
        f"Готово: позиций {len(all_items)}, ошибок {result.error_count}, "
        f"предупреждений {result.warning_count}, рассчитано {result.computed_count}"
    )
    return result


def _unique_path(base_dir: str, stem: str, source: str) -> str:
    src_stem = os.path.splitext(os.path.basename(source))[0]
    src_stem = "".join(ch if ch.isalnum() or ch in " -_()" else "_" for ch in src_stem).strip() or "файл"
    candidate = os.path.join(base_dir, f"{stem}_{src_stem}.xlsx")
    i = 2
    while os.path.exists(candidate):
        candidate = os.path.join(base_dir, f"{stem}_{src_stem}_{i}.xlsx")
        i += 1
    return candidate


def default_output_name(directory: str = "") -> str:
    stamp = __import__("datetime").datetime.now().strftime("%Y-%m-%d_%H%M")
    return os.path.join(directory, f"Загрузка_{stamp}.xlsx")


def selftest(directory: str = ".", log: Logger = _noop) -> Tuple[bool, str]:
    """Быстрая самопроверка: сконвертировать все файлы папки и сверить итоги."""
    files, _ = collect_files([directory])
    files = [f for f in files if is_supported(f)]
    if not files:
        return False, f"в папке {directory} нет файлов для проверки"
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "selftest.xlsx")
        res = convert(files, out, ConvertOptions(), log=log)
        if not res.items:
            return False, "позиции не найдены"
        if not os.path.exists(out):
            return False, "файл результата не создан"
        msg = (
            f"файлов {len(files)}, позиций {res.items_count}, "
            f"ошибок {res.error_count}, предупреждений {res.warning_count}, "
            f"рассчитано {res.computed_count}"
        )
        return res.error_count == 0, msg
