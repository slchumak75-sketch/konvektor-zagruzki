"""Оркестрация: файлы -> позиции -> проверки -> итоговая книга.

Единая точка входа для GUI и CLI. Ничего не знает про интерфейс, поэтому
используется в тестах напрямую (tests/test_pipeline.py).
"""

from __future__ import annotations

import datetime as _dt
import os
import math
import tempfile
import traceback
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from .analysis import AnalysisReport, build_report, format_report
from .extract import DocTotal, Item, extract_file
from .meta import DocMeta, extract_meta, fmt_date
from .readers import read_source
from .validate import ValidationReport, check_document_totals, validate_items
from .writer import FileStat, WorkbookPayload, write_workbook

#: Символы, недопустимые в именах файлов Windows.
_ILLEGAL_IN_NAME = '<>:"/\\|?*'


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
    ".rtf",
    ".pdf",
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
    #: добавить в итоговую таблицу колонки «файл» и «строка»
    add_source_columns: bool = False
    #: "auto" — папка «Документы» и имя «поставщик дата.xlsx», "path" — как указано
    naming: str = "path"
    #: куда сохранять при naming="auto" (по умолчанию — папка «Документы»)
    output_dir: str = ""
    #: писать ли файл сразу (False — только разобрать и показать таблицу)
    write: bool = True
    #: не сохранять неполный результат, если один из выбранных файлов не прочитан
    require_all: bool = False
    #: Исправления исходных ячеек: (файл, лист, строка, блок) -> поля Item.
    field_edits: Dict[Tuple[str, str, int, int], Dict[str, object]] = field(default_factory=dict)
    #: Явно подтверждённые страницы одного документа, в порядке страниц.
    document_groups: List[Tuple[str, ...]] = field(default_factory=list)
    #: склад («Мир», «Кореличи»): подпапка результата <папка>/<склад>/<поставщик>/
    sklad: str = ""
    #: поставщик, выбранный пользователем (пусто — берём из документа)
    supplier: str = ""
    #: сверка с ведомостью: объект vedomost.Matcher (None — без сверки).
    #: Переименовывает номенклатуру под ведомость ещё до проверок и записи.
    matcher: Optional[object] = None


@dataclass
class ConvertResult:
    outputs: List[str] = field(default_factory=list)
    items: List[Item] = field(default_factory=list)
    report: Optional[ValidationReport] = None
    document_totals: List[DocTotal] = field(default_factory=list)
    stats: List[FileStat] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    #: реквизиты по каждому файлу: путь -> DocMeta (поставщик, документ, дата)
    metas: Dict[str, DocMeta] = field(default_factory=dict)
    #: метрики разбора по каждому файлу (покрытие полей, арифметика, доверие)
    analysis: Dict[str, AnalysisReport] = field(default_factory=dict)
    #: куда файл был бы записан (заполняется и при write=False)
    planned_output: str = ""
    #: итог сверки с ведомостью (пусто, если сверки не было): см. Matcher.summary
    match_summary: Dict[str, object] = field(default_factory=dict)
    #: строки отчёта о сверке (что проверено, что осталось без подбора)
    match_log: List[str] = field(default_factory=list)

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


def safe_path_part(name: str, limit: int = 80) -> str:
    """Привести название к виду, допустимому в имени папки/файла Windows.

    «ООО "Культбытхозторгплюс"» -> «ООО Культбытхозторгплюс».
    """
    text = (name or "").replace("«", " ").replace("»", " ").replace('"', " ")
    for ch in _ILLEGAL_IN_NAME:
        text = text.replace(ch, " ")
    text = " ".join(text.split())
    text = text.strip(" .")
    if len(text) > limit:
        text = text[:limit].rstrip(" .")
    return text


def documents_dir() -> str:
    """Папка «Документы» пользователя (по умолчанию результат кладём туда).

    Windows: C:\\Users\\<имя>\\Documents. Если папки с таким именем нет —
    пробуем русские названия, затем домашнюю папку.
    """
    home = os.path.expanduser("~")
    for name in ("Documents", "Документы", "Мои документы"):
        candidate = os.path.join(home, name)
        if os.path.isdir(candidate):
            return candidate
    return home


def downloads_dir() -> str:
    """Папка «Загрузки» пользователя (Windows: C:\\Users\\<имя>\\Downloads)."""
    home = os.path.expanduser("~")
    for name in ("Downloads", "Загрузки", "Закачки"):
        candidate = os.path.join(home, name)
        if os.path.isdir(candidate):
            return candidate
    return documents_dir()


def unique_output_path(path: str) -> str:
    """Не перезаписывать готовый файл: «имя.xlsx» → «имя (2).xlsx».

    Пользователь просил, чтобы каждая конвертация давала новый файл,
    а не затирала предыдущий результат.
    """
    if not os.path.exists(path):
        return path
    folder, name = os.path.split(path)
    stem, ext = os.path.splitext(name)
    i = 2
    while True:
        candidate = os.path.join(folder, f"{stem} ({i}){ext}")
        if not os.path.exists(candidate):
            return candidate
        i += 1


def output_filename(
    meta: Optional[DocMeta],
    separator: str = " ",
    date: Optional[_dt.date] = None,
) -> str:
    """Имя файла результата: «поставщик дата.xlsx».

    Дата в имени — **дата сохранения** (сегодняшняя), а не дата документа:
    так просил пользователь.
    Слэш в имени файла Windows запрещён, поэтому «поставщик/дата» из
    технического задания превращается в «поставщик дата» (разделитель —
    пробел; можно задать другой).
    """
    supplier = safe_path_part(meta.supplier if meta else "") or "Поставщик не определён"
    day = date or _dt.date.today()
    return safe_path_part(f"{supplier}{separator}{fmt_date(day)}", limit=150) + ".xlsx"


def output_path_for(
    meta: Optional[DocMeta],
    base_dir: str = "",
    date: Optional[_dt.date] = None,
    sklad: str = "",
    supplier: str = "",
) -> str:
    """Результат в папке склада; склад, поставщики и дата — в имени файла."""
    base = base_dir or documents_dir()
    if sklad:
        name = supplier or (meta.supplier if meta else "") or "Поставщик не определён"
        day = date or _dt.date.today()
        prefix = safe_path_part(f"{sklad} {name}", limit=135)
        filename = f"{prefix} {fmt_date(day)}.xlsx"
        return os.path.join(base, safe_path_part(sklad), filename)
    if supplier:
        day = date or _dt.date.today()
        return os.path.join(base, safe_path_part(f"{supplier} {fmt_date(day)}", limit=150) + ".xlsx")
    return os.path.join(base, output_filename(meta, date=date))


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

    prior_sources: Dict[str, Sequence[str]] = {}
    group_errors: List[str] = []
    grouped = set()
    for group in options.document_groups:
        if len(group) < 2 or len(set(group)) != len(group) or grouped.intersection(group) or any(p not in files for p in group):
            group_errors.append("Некорректная группа страниц накладной: " + ", ".join(group))
            continue
        grouped.update(group)
        prior_sources[group[-1]] = group[:-1]
    result.errors.extend(group_errors)

    all_items: List[Item] = []
    totals = []
    blocks: Dict[Tuple[str, str, int], List[Item]] = {}
    per_file_items: List[Tuple[str, List[Item]]] = []
    metas: Dict[str, DocMeta] = {}
    analyses: Dict[str, AnalysisReport] = {}
    edit_counts: Dict[Tuple[str, str, int, int], int] = {}
    edit_errors: List[str] = []
    text_fields = {"name", "barcode", "article", "unit"}
    number_fields = {"qty", "price", "total_sum", "vat_rate", "vat_sum", "sum_with_vat"}

    for path in files:
        log(f"Читаю: {os.path.basename(path)}")
        try:
            src = read_source(path)
            extracted = extract_file(src)
        except Exception as exc:  # noqa: BLE001 - один битый файл не должен рушить всё
            msg = (
                f"{os.path.basename(path)}: не удалось прочитать "
                f"({type(exc).__name__}: {exc})"
            )
            result.errors.append(msg)
            result.notes.append(
                f"{os.path.basename(path)}: подсказка — закройте файл в Excel и проверьте, "
                "что это настоящий Excel; разбор файла: Konvertor-cli.exe --diagnose "
                f"\"{os.path.basename(path)}\""
            )
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

        meta = extract_meta(src)
        metas[path] = meta
        result.metas[path] = meta
        log(
            f"  документ: {meta.doc_label or '—'}"
            + (f" от {meta.date_text}" if meta.doc_date else "")
            + f"; поставщик: {meta.supplier or '—'}"
            + (f" (УНП {meta.supplier_unp})" if meta.supplier_unp else "")
        )
        for w in meta.warnings:
            log("  ! " + w)

        # Редактируем исходные значения до подбора, валидации и записи.
        for item in extracted.items:
            key = (item.source, item.sheet, item.row, item.block)
            if key not in options.field_edits:
                continue
            edit_counts[key] = edit_counts.get(key, 0) + 1
            changes = options.field_edits[key]
            for name, value in changes.items():
                valid = (name in text_fields and isinstance(value, str) and (name != "name" or bool(value.strip()))) or (
                    name in number_fields and isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value))
                if not valid:
                    edit_errors.append(f"Недопустимое исправление {name}: {key}")
                    continue
                setattr(item, name, value)
                item.unrecognized.discard(name)

        matcher = options.matcher
        if matcher is not None:
            # сверка по поставщику документа: begin_file задаёт область поиска
            matcher.begin_file(path, meta.supplier)
            for item in extracted.items:
                if "name" not in item.unrecognized:
                    matcher.apply(item)
            log(f"  сверка с ведомостью ({getattr(matcher, 'strategy_title', '')}): " + (matcher.summary_text() or "—"))
            for line in matcher.report_lines():
                log("  " + line)

        analysis = build_report(path, src, extracted, meta)
        analyses[path] = analysis
        result.analysis[path] = analysis
        log(
            f"  разбор: колонок {analysis.roles_found}/{analysis.roles_total}, "
            + (
                "арифметика н/д, "
                if analysis.arithmetic is None
                else f"арифметика {analysis.arithmetic:.0%}, "
            )
            + f"доверие {analysis.confidence:.2f} ({analysis.quality})"
        )
        if analysis.needs_review:
            log("  ! разбор требует проверки: " + "; ".join(analysis.reasons[:3]))

        all_items.extend(extracted.items)
        totals.extend(extracted.totals)
        for it in extracted.items:
            blocks.setdefault((it.source, it.sheet, it.block), []).append(it)
        per_file_items.append((path, extracted.items))
        result.notes.append(f"{os.path.basename(path)}: листы [{sheet_names}], формат {src.reader}")
        if not extracted.items:
            note = f"{os.path.basename(path)}: позиции не найдены"
            result.notes.append(note)
            if options.require_all:
                result.errors.append(note)
            log("  ! позиции не найдены")

    for key in options.field_edits:
        if edit_counts.get(key, 0) != 1:
            edit_errors.append(f"Исправленная строка не найдена или неоднозначна: {key}")
    result.errors.extend(edit_errors)

    if not all_items:
        result.notes.append("Позиции товаров не найдены — проверьте структуру файлов")
        return result

    if options.matcher is not None:
        result.match_summary = dict(options.matcher.summary)
        result.match_log = list(options.matcher.report_lines())
        # в лист «Проверки» уходит одна строка-итог, подробности — в match_log,
        # чтобы ведомость на 30 тысяч строк не превратила отчёт в простыню
        summary = options.matcher.summary_text()
        if summary:
            result.notes.append("сверка с ведомостью: " + summary)

    report = validate_items(all_items) if options.validate else ValidationReport(items=all_items)
    check_document_totals(report, totals, blocks, prior_sources=prior_sources)
    result.report = report
    result.items = all_items
    result.document_totals = totals

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
        meta = metas.get(path)
        if meta is not None:
            st.supplier = meta.supplier
            st.supplier_unp = meta.supplier_unp
            st.doc_type = meta.doc_type
            st.doc_number = meta.doc_number
            st.doc_date = fmt_date(meta.doc_date)
            st.meta_warnings = list(meta.warnings)
        analysis = analyses.get(path)
        if analysis is not None:
            st.quality = analysis.quality
            st.confidence = analysis.confidence
            st.roles_found = analysis.roles_found
            st.review = analysis.needs_review
        for note in result.notes:
            if note.startswith(os.path.basename(path) + ": листы"):
                st.sheets = note.split("[", 1)[1].split("]", 1)[0]
                st.reader = note.rsplit("формат ", 1)[-1]
                break
        stats.append(st)
    result.stats = stats

    # ---- куда писать -----------------------------------------------------
    today = _dt.date.today()  # дата сохранения — одна на весь запуск
    if options.naming == "auto":
        base_dir = options.output_dir or documents_dir()
        if len(per_file_items) == 1:
            out_path = unique_output_path(
                output_path_for(
                    metas.get(per_file_items[0][0]),
                    base_dir,
                    date=today,
                    sklad=options.sklad,
                    supplier=options.supplier,
                )
            )
        else:
            if options.supplier:
                combined_supplier = options.supplier
            else:
                unique_suppliers = sorted(
                    {
                        m.supplier.strip()
                        for p, _ in per_file_items
                        if (m := metas.get(p)) and m.supplier and m.supplier.strip()
                    },
                    key=str.casefold,
                )
                combined_supplier = ", ".join(unique_suppliers)
            if combined_supplier or options.sklad:
                first_meta = metas.get(per_file_items[0][0])
                out_path = unique_output_path(
                    output_path_for(
                        first_meta,
                        base_dir,
                        date=today,
                        sklad=options.sklad,
                        supplier=combined_supplier,
                    )
                )
            else:
                out_path = unique_output_path(default_output_name(base_dir))
        result.planned_output = out_path
        log(f"  файл результата: {out_path}")
    elif not out_path:
        base_dir = options.output_dir or os.path.dirname(os.path.abspath(files[0])) or os.getcwd()
        out_path = default_output_name(base_dir)
    result.planned_output = result.planned_output or out_path

    if not options.write:
        # только разбор и предпросмотр — файл запишет кнопка «Сохранить»
        log(f"  файл не записан: нажмите «Сохранить» ({result.planned_output})")
        log(
            f"Готово: позиций {len(all_items)}, ошибок {result.error_count}, "
            f"предупреждений {result.warning_count}, рассчитано {result.computed_count}"
        )
        return result

    unresolved = [item for item in all_items if item.unrecognized]
    if unresolved:
        result.errors.append(f"Не распознаны значения в {len(unresolved)} строках: исправьте ячейки с ??? перед сохранением")
        log("  файл не записан: есть нераспознанные значения")
        return result
    if edit_errors or group_errors or (options.require_all and result.errors):
        log("  файл не записан: не все исходные файлы удалось обработать")
        return result

    # ---- запись ----------------------------------------------------------
    target_dir = os.path.dirname(os.path.abspath(out_path)) or "."
    try:
        os.makedirs(target_dir, exist_ok=True)
    except OSError as exc:
        # понятная ошибка вместо «[Errno 2] No such file or directory»
        raise ValueError(
            f"папка «{target_dir}» недоступна для записи ({type(exc).__name__}: {exc}). "
            "Проверьте путь в поле «Куда сохранить» — возможно, диск не подключён."
        ) from exc

    if options.mode == MODE_FILES:
        base_dir = os.path.dirname(os.path.abspath(out_path)) or "."
        stem = os.path.splitext(os.path.basename(out_path))[0]
        for path, items in per_file_items:
            if not items:
                continue
            if options.naming == "auto":
                name = unique_output_path(
                    output_path_for(
                        metas.get(path),
                        options.output_dir or documents_dir(),
                        date=today,
                        sklad=options.sklad,
                        supplier=options.supplier,
                    )
                )
            else:
                name = _unique_path(base_dir, stem, path)
            payload = WorkbookPayload(
                data_sheets=[("Данные", items)],
                report=report,
                stats=stats,
                notes=result.notes,
                with_source=options.add_source_columns,
                with_novinka=options.matcher is not None,
            )
            result.outputs.append(_write_with_fallback(name, payload, result, log))
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
            with_novinka=options.matcher is not None,
        )
        result.outputs.append(_write_with_fallback(out_path, payload, result, log))

    log(
        f"Готово: позиций {len(all_items)}, ошибок {result.error_count}, "
        f"предупреждений {result.warning_count}, рассчитано {result.computed_count}"
    )
    return result


def _write_with_fallback(
    path: str, payload: WorkbookPayload, result: ConvertResult, log: Logger
) -> str:
    """Записать книгу; если папка недоступна — сохранить в «Документы»/домашнюю.

    Типичный случай на Windows: файл открыт из защищённой папки или с сетевого
    диска только для чтения — рядом писать нельзя. Молча падать нельзя, поэтому
    результат уходит в доступное место, а пользователь видит, куда именно.
    """
    try:
        written = write_workbook(path, payload)
        log(f"  записан: {written}")
        return written
    except OSError as exc:
        base = os.path.join(os.path.expanduser("~"), "Документы")
        if not os.path.isdir(base):
            base = os.path.expanduser("~")
        alt = os.path.join(base, os.path.basename(path))
        written = write_workbook(alt, payload)
        message = (
            f"не удалось записать {path} ({type(exc).__name__}: {exc}); "
            f"результат сохранён в {written}"
        )
        result.notes.append(message)
        log("  ! " + message)
        return written


def explain(paths: Sequence[str], log: Logger = _noop) -> List[str]:
    """Текстовые отчёты о разборе файлов (ключ --explain)."""
    out: List[str] = []
    files, missing = collect_files(paths)
    for m in missing:
        out.append(f"файл не найден: {m}")
    for path in files:
        try:
            src = read_source(path)
            extracted = extract_file(src)
        except Exception as exc:  # noqa: BLE001 - показываем причину
            out.append(f"{path}: не удалось прочитать ({type(exc).__name__}: {exc})")
            continue
        meta = extract_meta(src)
        report = build_report(path, src, extracted, meta)
        out.append(format_report(report))
        log(format_report(report))
    return out


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


def diagnose(path: str, log: Logger = _noop) -> Tuple[bool, List[str]]:
    """Подробный разбор одного файла: почему не читается / что нашлось.

    Нужен, когда у пользователя файл «не конвертируется»: вывод можно
    переслать разработчику как есть.
    """
    lines: List[str] = []

    def say(text: str) -> None:
        lines.append(text)
        log(text)

    say(f"Файл: {path}")
    if not os.path.exists(path):
        say("ОШИБКА: файл не найден по этому пути")
        return False, lines
    try:
        say(f"Размер: {os.path.getsize(path)} байт")
        say(f"Расширение: {os.path.splitext(path)[1] or '(нет)'}")
        say(f"Поддерживается: {'да' if is_supported(path) else 'нет'}")
    except OSError as exc:  # noqa: BLE001 - проблемы доступа к файлу
        say(f"ОШИБКА доступа к файлу: {type(exc).__name__}: {exc}")
        return False, lines

    src = None
    try:
        src = read_source(path)
    except Exception as exc:  # noqa: BLE001 - показываем пользователю причину
        say(f"ОШИБКА чтения: {type(exc).__name__}: {exc}")
        say("")
        say("Возможные причины: файл открыт в Excel (закройте его), файл повреждён,")
        say("это не настоящий Excel (например, HTML/CSV с чужим расширением),")
        say("или формат новее, чем поддерживает программа.")
        say("Подробности (traceback):")
        for line in traceback.format_exc().splitlines():
            say("  " + line)
        return False, lines

    try:
        say(f"Прочитан как: {src.reader}; листов: {len(src.sheets)}")
        for sheet in src.sheets[:10]:
            cols = max((len(r) for r in sheet.rows), default=0)
            say(f"  лист «{sheet.name}»: строк {len(sheet.rows)}, столбцов {cols}")
        meta = extract_meta(src)
        say(
            "Реквизиты: поставщик "
            f"{meta.supplier or '— не найден'}"
            + (f" (УНП {meta.supplier_unp})" if meta.supplier_unp else "")
            + f"; документ {meta.doc_label or '— не определён'}"
            + (f" от {meta.date_text}" if meta.doc_date else ", дата не найдена")
        )
        extracted = extract_file(src)
        say(f"Позиций найдено: {len(extracted.items)}")
        report = build_report(path, src, extracted, meta)
        say(
            f"Качество разбора: {report.quality}; колонок {report.roles_found}"
            f"/{report.roles_total}; "
            + (
                "арифметика не проверялась; "
                if report.arithmetic is None
                else f"арифметика {report.arithmetic:.0%}; "
            )
            + f"доверие {report.confidence:.2f}"
        )
        for reason in report.reasons[:8]:
            say("  - " + reason)
        for warn in extracted.warnings[:10]:
            say("  ! " + warn)
        say("Итог: файл читается" if extracted.items else "Итог: файл читается, но позиции не найдены")
        return True, lines
    except Exception as exc:  # noqa: BLE001 - разбор тоже может упасть
        say(f"ОШИБКА разбора: {type(exc).__name__}: {exc}")
        for line in traceback.format_exc().splitlines():
            say("  " + line)
        return False, lines


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
        used = len(res.stats)
        msg = (
            f"файлов с данными {used} из {len(files)}, позиций {res.items_count}, "
            f"ошибок {res.error_count}, предупреждений {res.warning_count}, "
            f"рассчитано {res.computed_count}"
        )
        return res.error_count == 0, msg
