"""Оконное приложение для Windows (tkinter — входит в стандартный Python).

Логики здесь минимум: всё, что можно посчитать, живёт в pipeline.py. GUI
только принимает **один** файл, запускает конвертацию в отдельном потоке и
показывает в нижней части окна саму полученную таблицу (как в Excel).

Особенности, заданные заказчиком:
  * файл обрабатывается по одному; после конвертации список очищается сам;
  * файл результата сохраняется как «поставщик / дата документа.xlsx»
    (папка поставщика создаётся автоматически);
  * в таблицу всегда добавляются колонки «файл» и «строка»;
  * журнала нет — вместо него предпросмотр готовой таблицы.

Файл можно не только выбрать кнопкой, но и перетащить на ярлык
Konvektor.exe — Windows передаст его в командной строке (см. app_gui.py).
"""

from __future__ import annotations

import os
import queue
import sys
import threading
import traceback
from typing import List, Optional

from . import APP_NAME, TEMPLATE_COLUMNS, __version__
from .extract import Item
from .pipeline import (
    MODE_COMBINED,
    ConvertOptions,
    ConvertResult,
    convert,
    is_supported,
)

try:  # pragma: no cover - зависит от сборки Python
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    TK_AVAILABLE = True
except Exception:  # noqa: BLE001 - на «голом» Linux без python3-tk
    tk = None  # type: ignore[assignment]
    filedialog = messagebox = ttk = None  # type: ignore[assignment]
    TK_AVAILABLE = False


if TK_AVAILABLE:

    class _Base(tk.Tk):
        pass

else:  # pragma: no cover

    class _Base:  # type: ignore[no-redef]
        def __init__(self, *args, **kwargs):
            raise RuntimeError(
                "Оконный интерфейс требует tkinter (в Windows он входит в Python). "
                "Используйте консольную версию: Konvektor-cli.exe"
            )


FILE_TYPES = [
    (
        "Таблицы и счета",
        "*.xls *.xlsx *.xlsm *.csv *.txt *.html *.htm *.mht",
    ),
    ("Excel 97-2003 (*.xls)", "*.xls"),
    ("Excel (*.xlsx)", "*.xlsx"),
    ("Все файлы", "*.*"),
]

#: Колонки предпросмотра — только шаблон (лишних колонок в окне нет;
#: «файл» и «строка» остаются только в сохранённом файле).
PREVIEW_COLUMNS = list(TEMPLATE_COLUMNS)
#: Ширины — в пикселях (у Treeview width именно пиксели, а не символы).
PREVIEW_WIDTHS = [125, 105, 380, 75, 85, 95, 105, 85, 105, 115]
PREVIEW_MIN_WIDTHS = [90, 70, 160, 55, 60, 70, 80, 60, 80, 90]
#: Цвет линий таблицы
GRID_COLOR = "#9DB2C8"
#: Числовые колонки выравниваем по правому краю, как в Excel.
_NUMERIC_PREVIEW = {"количество", "цена", "сумма", "ставка НДС", "сумма НДС", "сумма с НДС"}

#: Сколько строк показывать в предпросмотре (остальные — в файле).
PREVIEW_LIMIT = 500


class KonvektorApp(_Base):
    """Главное окно."""

    def __init__(self, initial_files: Optional[List[str]] = None):
        super().__init__()
        self.title(f"{APP_NAME} {__version__}")
        self.geometry("1100x720")
        self.minsize(900, 600)

        self.files: List[str] = []
        self._queue: "queue.Queue[tuple]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._running = False
        self._last_outputs: List[str] = []
        self.log_lines: List[str] = []
        self._pending_source: Optional[str] = None
        self._pending_out: str = ""
        self._pending_result: Optional[ConvertResult] = None

        self._set_window_icon()
        self._build_ui()
        if initial_files:
            self._add_files(initial_files)
        self.after(120, self._poll_queue)

    # ------------------------------------------------------------------ UI
    def _set_window_icon(self) -> None:
        """Иконка окна. В собранном .exe картинка лежит в assets (см. --add-data)."""
        base = getattr(sys, "_MEIPASS", None)
        candidates = []
        if base:
            candidates.append(os.path.join(base, "assets", "konvektor_icon.png"))
        here = os.path.dirname(os.path.abspath(__file__))
        candidates.append(os.path.join(os.path.dirname(here), "assets", "konvektor_icon.png"))
        for path in candidates:
            if not os.path.exists(path):
                continue
            try:
                self._icon_image = tk.PhotoImage(file=path)
                self.iconphoto(True, self._icon_image)
                return
            except Exception:  # noqa: BLE001 - иконка не критична
                continue

    def _build_ui(self) -> None:
        # --- 1. один файл
        top = ttk.LabelFrame(self, text="1. Файл для конвертации (по одному)")
        top.pack(fill="x", padx=10, pady=(10, 4))

        ttk.Label(
            top,
            text="Поддерживаются .xls, .xlsx, .csv, .txt, .html. "
            "Файл можно перетащить мышью на ярлык Konvektor.exe.",
            foreground="#555555",
        ).pack(side="bottom", anchor="w", padx=8, pady=(0, 6))

        self.file_var = tk.StringVar(value="файл не выбран")
        ttk.Label(top, textvariable=self.file_var, anchor="w").pack(
            side="left", fill="x", expand=True, padx=8, pady=8
        )

        ttk.Button(top, text="Выбрать файл…", command=self.on_add_files).pack(
            side="right", padx=8, pady=8
        )

        # --- 2. куда сохранить
        mid = ttk.LabelFrame(self, text="2. Куда сохранить")
        mid.pack(fill="x", padx=10, pady=4)

        row1 = ttk.Frame(mid)
        row1.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(row1, text="Итоговый файл:").pack(side="left")
        self.out_var = tk.StringVar(value="")
        ttk.Entry(row1, textvariable=self.out_var).pack(
            side="left", fill="x", expand=True, padx=6
        )
        ttk.Button(row1, text="Обзор…", command=self.on_browse_output).pack(side="left")

        ttk.Label(
            mid,
            text="Пусто — сохранить в папку «Документы» с именем "
            "«поставщик дата документа.xlsx».",
            foreground="#555555",
        ).pack(anchor="w", padx=8, pady=(0, 8))

        # --- действие
        act = ttk.Frame(self)
        act.pack(fill="x", padx=10, pady=4)
        self.run_btn = ttk.Button(act, text="Конвертировать", command=self.on_convert)
        self.run_btn.pack(side="left")
        self.save_btn = ttk.Button(
            act, text="Сохранить", command=self.on_save, state="disabled"
        )
        self.save_btn.pack(side="left", padx=(8, 0))

        # --- 3. предпросмотр результата вместо журнала
        bottom = ttk.LabelFrame(self, text="3. Результат (таблица)")
        bottom.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        table_wrap = ttk.Frame(bottom)
        table_wrap.pack(fill="both", expand=True, padx=8, pady=8)

        self._setup_table_style()
        # рамка вокруг таблицы не зависит от темы оформления Windows
        grid_box = tk.Frame(
            table_wrap, bd=1, relief="solid", highlightthickness=0,
            background=GRID_COLOR,
        )
        grid_box.pack(fill="both", expand=True)
        self.preview = ttk.Treeview(
            grid_box,
            columns=PREVIEW_COLUMNS,
            show="headings",
            height=12,
            style="Grid.Treeview",
            selectmode="browse",
        )
        # «разлиновка»: чередование строк, как в Excel
        self.preview.tag_configure("odd", background="#F2F6FB")
        self.preview.tag_configure("even", background="#FFFFFF")
        for title, width, min_width in zip(
            PREVIEW_COLUMNS, PREVIEW_WIDTHS, PREVIEW_MIN_WIDTHS
        ):
            numeric = title in _NUMERIC_PREVIEW
            self.preview.heading(title, text=title, anchor="center")
            self.preview.column(
                title,
                width=width,
                minwidth=min_width,
                anchor="e" if numeric else "w",
                stretch=(title == "Наименование товара"),
            )

        yscroll = ttk.Scrollbar(grid_box, orient="vertical", command=self.preview.yview)
        xscroll = ttk.Scrollbar(grid_box, orient="horizontal", command=self.preview.xview)
        self.preview.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        yscroll.pack(side="right", fill="y")
        xscroll.pack(side="bottom", fill="x")
        self.preview.pack(side="left", fill="both", expand=True)

        self.hint_var = tk.StringVar(value="")
        ttk.Label(bottom, textvariable=self.hint_var, foreground="#555555").pack(
            anchor="w", padx=8, pady=(0, 6)
        )

    def _setup_table_style(self) -> None:
        """Разлиновка таблицы: подсветка строк через одну, объёмная шапка, высота строк."""
        try:
            style = ttk.Style(self)
            style.configure(
                "Grid.Treeview",
                background="#FFFFFF",
                fieldbackground="#FFFFFF",
                foreground="#000000",
                rowheight=23,
                borderwidth=1,
                relief="solid",
                font=("Segoe UI", 9),
            )
            style.configure(
                "Grid.Treeview.Heading",
                background="#DCE6F1",
                foreground="#1F4E78",
                relief="raised",
                borderwidth=1,
                padding=(4, 4),
                font=("Segoe UI", 9, "bold"),
            )
            style.map("Grid.Treeview.Heading", background=[("active", "#CFE0F5")])
            style.map("Grid.Treeview", background=[("selected", "#BBDEFB")])
        except Exception:  # noqa: BLE001 - оформление не должно ломать окно
            pass

    # ------------------------------------------------------------- действия
    def on_add_files(self) -> None:
        path = filedialog.askopenfilename(title="Выберите файл", filetypes=FILE_TYPES)
        if path:
            self._add_files([path])

    def on_browse_output(self) -> None:
        initial = self.out_var.get()
        folder = os.path.dirname(initial) or os.path.dirname(self.files[0] if self.files else "") or os.getcwd()
        path = filedialog.asksaveasfilename(
            title="Куда сохранить результат",
            defaultextension=".xlsx",
            initialfile=os.path.basename(initial) or "Результат.xlsx",
            initialdir=folder,
            filetypes=[("Excel (*.xlsx)", "*.xlsx")],
        )
        if path:
            self.out_var.set(path)

    # ------------------------------------------------------------ конвертация
    def on_convert(self) -> None:
        if self._running:
            return
        if not self.files:
            messagebox.showwarning(APP_NAME, "Сначала выберите файл.")
            return

        path = self.files[0]
        out = self.out_var.get().strip()
        options = ConvertOptions(
            mode=MODE_COMBINED,
            # колонки «файл» и «строка» нужны всегда (в сохранённом файле)
            add_source_columns=True,
            # без явного пути — «Документы/поставщик дата.xlsx»
            naming="path" if out else "auto",
            # «Конвертировать» только показывает таблицу; пишет «Сохранить»
            write=False,
        )

        self._running = True
        self.run_btn.configure(state="disabled")
        self.save_btn.configure(state="disabled")
        self._pending_source = path
        self._pending_out = out
        self._pending_result = None
        self.hint_var.set("")
        self._log_clear()
        self._clear_preview()
        self._log(f"{APP_NAME} {__version__}")
        self._log(f"файл: {os.path.basename(path)}")

        def worker() -> None:
            try:
                result = convert([path], out, options, log=lambda m: self._queue.put(("log", m)))
                self._queue.put(("done", result))
            except Exception as exc:  # noqa: BLE001 - показываем пользователю
                self._queue.put(("error", (exc, traceback.format_exc())))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def on_save(self) -> None:
        """Записать разобранную таблицу в xlsx (кнопка «Сохранить»)."""
        if self._running or self._pending_result is None:
            return
        if not self._pending_result.items_count:
            messagebox.showwarning(APP_NAME, "Нечего сохранять: позиции не найдены.")
            return

        path = self._pending_source or ""
        if not path or not os.path.exists(path):
            messagebox.showwarning(
                APP_NAME, "Исходный файл больше недоступен — выберите его снова."
            )
            return

        out = (self._pending_out or self.out_var.get()).strip()
        options = ConvertOptions(
            mode=MODE_COMBINED,
            add_source_columns=True,
            naming="path" if out else "auto",
            write=True,
        )

        self._running = True
        self.save_btn.configure(state="disabled")
        self.run_btn.configure(state="disabled")

        def worker() -> None:
            try:
                result = convert([path], out, options, log=lambda m: self._queue.put(("log", m)))
                self._queue.put(("saved", result))
            except Exception as exc:  # noqa: BLE001 - показываем пользователю
                self._queue.put(("error", (exc, traceback.format_exc())))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _saved(self, result: ConvertResult) -> None:
        self._running = False
        self.run_btn.configure(state="normal")
        self.save_btn.configure(state="normal")
        self._last_outputs = list(result.outputs)
        if not result.outputs:
            messagebox.showwarning(APP_NAME, "; ".join(result.errors) or "Файл не записан.")
            return
        written = result.outputs[0]
        folder, name = os.path.split(written)
        self.hint_var.set(f"Сохранено: {folder} / {name}")
        messagebox.showinfo(APP_NAME, f"Файл сохранён:\n{written}")

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "log":
                    self._log(str(payload))
                elif kind == "done":
                    self._finish(payload)
                elif kind == "saved":
                    self._saved(payload)
                elif kind == "error":
                    self._fail(*payload)
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _finish(self, result: ConvertResult) -> None:
        self._running = False
        self.run_btn.configure(state="normal")
        self._last_outputs = list(result.outputs)
        self._pending_result = result
        self.save_btn.configure(state="normal" if result.items_count else "disabled")

        self._show_preview(result.items)
        self._log("")
        self._log(
            f"Позиций: {result.items_count}; ошибок: {result.error_count}; "
            f"предупреждений: {result.warning_count}; рассчитано значений: {result.computed_count}"
        )
        for message in result.errors:
            self._log("  ОШИБКА: " + message)
        for path, report in result.analysis.items():
            if report.needs_review:
                self._log(
                    f"  ПРОВЕРИТЬ: {os.path.basename(path)} — доверие {report.confidence:.2f}"
                    f" ({report.quality}); " + "; ".join(report.reasons[:3])
                )
        docs = "; ".join(
            (m.supplier or "поставщик не найден")
            + (f", {m.doc_label}" if m.doc_number else "")
            + (f" от {m.date_text}" if m.doc_date else "")
            for m in result.metas.values()
        )
        # файл обработан — список очищается сам (до подсказки, чтобы её не затёрло)
        self._clear_files()

        planned = result.planned_output
        if planned and result.items_count:
            folder, name = os.path.split(planned)
            self.hint_var.set(
                f"{docs} → будет сохранено: {folder} / {name} (кнопка «Сохранить»)"
            )
        elif docs:
            self.hint_var.set(docs)

        if result.items_count:
            text = (
                f"Таблица готова.\n\nПозиций: {result.items_count}\n"
                f"Документ: {docs}\n"
                f"Расхождений в суммах: {result.error_count}\n"
                f"Предупреждений: {result.warning_count}\n"
                f"Значений рассчитано программой: {result.computed_count}"
            )
            if planned:
                text += f"\n\nКнопка «Сохранить» запишет файл:\n{planned}"
            if result.errors:
                text += "\n\nОшибки: " + "; ".join(result.errors)
            if result.error_count:
                messagebox.showwarning(APP_NAME, text + "\n\nПодробности — на листе «Проверки».")
            else:
                messagebox.showinfo(APP_NAME, text)
        else:
            messagebox.showwarning(
                APP_NAME,
                "Не удалось найти ни одной позиции.\n\n"
                "Проверьте, что в файле есть таблица с колонками "
                "«наименование», «количество», «цена» и т.п.",
            )

    def _fail(self, exc: Exception, tb: str) -> None:
        self._running = False
        self.run_btn.configure(state="normal")
        self._log("ОШИБКА: " + str(exc))
        self._log(tb)
        messagebox.showerror(APP_NAME, f"Произошла ошибка:\n{exc}")
        self._clear_files()

    # --------------------------------------------------------------- утилиты
    def _show_preview(self, items: List[Item], limit: int = PREVIEW_LIMIT) -> None:
        """Показать полученную таблицу так, как она попадёт в Excel."""
        self._clear_preview()
        for item in items[:limit]:
            values = [
                item.barcode or "",
                item.article or "",
                item.name,
                item.unit or "",
                item.qty,
                item.price,
                item.total_sum,
                item.vat_rate,
                item.vat_sum,
                item.sum_with_vat,
            ]
            tag = "odd" if len(self.preview.get_children()) % 2 else "even"
            self.preview.insert(
                "", "end", values=[self._fmt(v) for v in values], tags=(tag,)
            )
        if not items:
            self.hint_var.set("Позиции не найдены — смотрите сообщение об ошибке.")

    @staticmethod
    def _fmt(value) -> str:
        if value is None:
            return ""
        if isinstance(value, float):
            return f"{value:.2f}".replace(".", ",")
        return str(value)

    def _clear_preview(self) -> None:
        for child in self.preview.get_children():
            self.preview.delete(child)

    def _add_files(self, paths: List[str]) -> None:
        """Один файл за раз: новый выбор заменяет предыдущий."""
        for path in paths:
            full = os.path.abspath(path)
            if not is_supported(full):
                self._log(f"пропущен (неподдерживаемый формат): {full}")
                continue
            self.files = [full]
            self.file_var.set(os.path.basename(full))
            self.out_var.set("")
            self.hint_var.set("")
            self._log(f"выбран файл: {full}")
            break
        self._update_status()

    def _clear_files(self) -> None:
        self.files.clear()
        self.file_var.set("файл не выбран")
        self.out_var.set("")
        self._update_status()

    def _update_status(self) -> None:
        """Состояние показываем в подсказке под таблицей, без отдельного поля."""
        if self._running:
            return
        if not self.files:
            self.hint_var.set("")
            return
        self.hint_var.set("Файл выбран — нажмите «Конвертировать»")

    def _log(self, message: str) -> None:
        self.log_lines.append(message)

    def _log_clear(self) -> None:
        self.log_lines.clear()


def main(argv: Optional[List[str]] = None) -> int:
    """Запуск GUI. Аргументы командной строки = файл для конвертации."""
    if not TK_AVAILABLE:
        print(
            "Оконный интерфейс недоступен: не найден tkinter.\n"
            "Используйте консольную версию: Konvektor-cli.exe файл.xls -o результат.xlsx",
            file=sys.stderr,
        )
        return 1

    _enable_dpi_awareness()
    files = [a for a in (argv if argv is not None else sys.argv[1:]) if os.path.exists(a)]
    app = KonvektorApp(initial_files=files)
    app.mainloop()
    return 0


def _enable_dpi_awareness() -> None:
    """Чтобы шрифты не мылились на мониторах с масштабированием (Windows)."""
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001 - не Windows или старый Windows
        pass


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
