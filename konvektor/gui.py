"""Оконное приложение для Windows (tkinter — входит в стандартный Python).

Логики здесь минимум: всё, что можно посчитать, живёт в pipeline.py.  GUI
только собирает файлы, запускает конвертацию в отдельном потоке и показывает
журнал.  Благодаря этому интерфейс не мешает тестировать конвертер.

Файлы можно не только выбирать кнопками, но и перетащить на ярлык
Konvektor.exe — Windows передаст их в командной строке (см. app_gui.py).
"""

from __future__ import annotations

import os
import queue
import sys
import threading
import traceback
from typing import List, Optional

from . import APP_NAME, __version__
from .pipeline import (
    MODE_COMBINED,
    MODE_FILES,
    MODE_SHEETS,
    MODE_TITLES,
    ConvertOptions,
    ConvertResult,
    convert,
    default_output_name,
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


class KonvektorApp(_Base):
    """Главное окно."""

    def __init__(self, initial_files: Optional[List[str]] = None):
        super().__init__()
        self.title(f"{APP_NAME} {__version__}")
        self.geometry("980x680")
        self.minsize(820, 560)

        self.files: List[str] = []
        self._queue: "queue.Queue[tuple]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._running = False
        self._last_outputs: List[str] = []

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
        # --- список файлов
        top = ttk.LabelFrame(self, text="1. Файлы для конвертации")
        top.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        ttk.Label(
            top,
            text="Поддерживаются .xls, .xlsx, .csv, .txt, .html. "
            "Файлы можно перетащить мышью на ярлык Konvektor.exe.",
            foreground="#555555",
        ).pack(side="bottom", anchor="w", padx=8, pady=(0, 6))

        list_wrap = ttk.Frame(top)
        list_wrap.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        self.listbox = tk.Listbox(list_wrap, selectmode="extended", height=8, activestyle="dotbox")
        scroll = ttk.Scrollbar(list_wrap, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        btns = ttk.Frame(top)
        btns.pack(side="right", fill="y", padx=8, pady=8)
        ttk.Button(btns, text="Добавить файлы…", command=self.on_add_files).pack(fill="x", pady=2)
        ttk.Button(btns, text="Добавить папку…", command=self.on_add_folder).pack(fill="x", pady=2)
        ttk.Button(btns, text="Убрать выбранные", command=self.on_remove).pack(fill="x", pady=2)
        ttk.Button(btns, text="Очистить список", command=self.on_clear).pack(fill="x", pady=2)

        # --- настройки
        mid = ttk.LabelFrame(self, text="2. Куда и как сохранить")
        mid.pack(fill="x", padx=10, pady=4)

        row1 = ttk.Frame(mid)
        row1.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(row1, text="Итоговый файл:").pack(side="left")
        self.out_var = tk.StringVar(value="")
        ttk.Entry(row1, textvariable=self.out_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row1, text="Обзор…", command=self.on_browse_output).pack(side="left")

        row2 = ttk.Frame(mid)
        row2.pack(fill="x", padx=8, pady=(2, 8))
        ttk.Label(row2, text="Режим:").pack(side="left")
        self.mode_var = tk.StringVar(value=MODE_COMBINED)
        self.mode_box = ttk.Combobox(
            row2,
            textvariable=self.mode_var,
            state="readonly",
            values=[MODE_TITLES[k] for k in (MODE_COMBINED, MODE_SHEETS, MODE_FILES)],
            width=46,
        )
        self.mode_box.current(0)
        self.mode_box.pack(side="left", padx=6)

        self.add_source_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row2,
            text="добавить колонки «файл» и «строка»",
            variable=self.add_source_var,
        ).pack(side="left", padx=10)

        self.open_after_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2, text="открыть результат", variable=self.open_after_var).pack(side="left", padx=6)

        # --- действие
        act = ttk.Frame(self)
        act.pack(fill="x", padx=10, pady=4)
        self.run_btn = ttk.Button(act, text="Конвертировать", command=self.on_convert)
        self.run_btn.pack(side="left")
        self.progress = ttk.Progressbar(act, mode="indeterminate", length=260)
        self.progress.pack(side="left", padx=12)
        self.status = ttk.Label(act, text="Готов к работе")
        self.status.pack(side="left", padx=6)

        # --- журнал
        bottom = ttk.LabelFrame(self, text="3. Журнал")
        bottom.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        log_wrap = ttk.Frame(bottom)
        log_wrap.pack(fill="both", expand=True, padx=8, pady=8)
        self.log_text = tk.Text(log_wrap, height=12, wrap="word", state="disabled")
        log_scroll = ttk.Scrollbar(log_wrap, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

    # ------------------------------------------------------------- действия
    def on_add_files(self) -> None:
        paths = filedialog.askopenfilenames(title="Выберите файлы", filetypes=FILE_TYPES)
        if paths:
            self._add_files(list(paths))

    def on_add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Выберите папку с файлами")
        if not folder:
            return
        found = []
        for name in sorted(os.listdir(folder)):
            full = os.path.join(folder, name)
            if os.path.isfile(full) and is_supported(full):
                found.append(full)
        if not found:
            messagebox.showinfo(APP_NAME, "В папке нет поддерживаемых файлов.")
            return
        self._add_files(found)

    def on_remove(self) -> None:
        for idx in sorted(self.listbox.curselection(), reverse=True):
            self.listbox.delete(idx)
            del self.files[idx]
        self._update_status()

    def on_clear(self) -> None:
        self.listbox.delete(0, "end")
        self.files.clear()
        self._update_status()

    def on_browse_output(self) -> None:
        initial = self.out_var.get() or default_output_name(os.getcwd())
        path = filedialog.asksaveasfilename(
            title="Куда сохранить результат",
            defaultextension=".xlsx",
            initialfile=os.path.basename(initial),
            initialdir=os.path.dirname(initial) or os.getcwd(),
            filetypes=[("Excel (*.xlsx)", "*.xlsx")],
        )
        if path:
            self.out_var.set(path)

    # ------------------------------------------------------------ конвертация
    def on_convert(self) -> None:
        if self._running:
            return
        if not self.files:
            messagebox.showwarning(APP_NAME, "Сначала добавьте файлы.")
            return

        out = self.out_var.get().strip()
        if not out:
            out = default_output_name(os.path.dirname(self.files[0]) or os.getcwd())
            self.out_var.set(out)

        mode = {
            MODE_TITLES[MODE_COMBINED]: MODE_COMBINED,
            MODE_TITLES[MODE_SHEETS]: MODE_SHEETS,
            MODE_TITLES[MODE_FILES]: MODE_FILES,
        }.get(self.mode_var.get(), MODE_COMBINED)

        options = ConvertOptions(mode=mode, add_source_columns=self.add_source_var.get())
        files = list(self.files)

        self._running = True
        self.run_btn.configure(state="disabled")
        self.progress.start(12)
        self.status.configure(text="Конвертирую…")
        self._log_clear()
        self._log(f"{APP_NAME} {__version__}")
        self._log(f"Файлов: {len(files)}; режим: {MODE_TITLES[mode]}")
        self._log("")

        def worker() -> None:
            try:
                result = convert(
                    files,
                    out,
                    options,
                    log=lambda m: self._queue.put(("log", m)),
                )
                self._queue.put(("done", result))
            except Exception as exc:  # noqa: BLE001 - показываем пользователю
                self._queue.put(("error", (exc, traceback.format_exc())))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "log":
                    self._log(str(payload))
                elif kind == "done":
                    self._finish(payload)
                elif kind == "error":
                    self._fail(*payload)
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _finish(self, result: ConvertResult) -> None:
        self._running = False
        self.progress.stop()
        self.run_btn.configure(state="normal")
        self._last_outputs = list(result.outputs)

        self._log("")
        self._log(
            f"Позиций: {result.items_count}; ошибок: {result.error_count}; "
            f"предупреждений: {result.warning_count}; рассчитано значений: {result.computed_count}"
        )
        for issue in (result.report.errors if result.report else [])[:20]:
            self._log(
                f"  РАСХОЖДЕНИЕ: {os.path.basename(issue.source)} стр.{issue.row}: {issue.message}"
            )
        for path in result.outputs:
            self._log(f"Сохранено: {path}")
        self.status.configure(text=f"Готово: {result.items_count} позиций")

        if result.items_count:
            text = (
                f"Готово.\n\nПозиций: {result.items_count}\n"
                f"Расхождений в суммах: {result.error_count}\n"
                f"Предупреждений: {result.warning_count}\n"
                f"Значений рассчитано программой: {result.computed_count}\n\n"
                + "\n".join(result.outputs)
            )
            if result.error_count:
                messagebox.showwarning(APP_NAME, text + "\n\nПодробности — на листе «Проверки».")
            else:
                messagebox.showinfo(APP_NAME, text)
            if self.open_after_var.get() and result.outputs:
                self._open(result.outputs[0])
        else:
            messagebox.showwarning(
                APP_NAME,
                "Не удалось найти ни одной позиции.\n\n"
                "Проверьте, что в файлах есть таблица с колонками "
                "«наименование», «количество», «цена» и т.п.\n"
                "Подробности — в журнале.",
            )

    def _fail(self, exc: Exception, tb: str) -> None:
        self._running = False
        self.progress.stop()
        self.run_btn.configure(state="normal")
        self.status.configure(text="Ошибка")
        self._log("ОШИБКА: " + str(exc))
        self._log(tb)
        messagebox.showerror(APP_NAME, f"Произошла ошибка:\n{exc}")

    # --------------------------------------------------------------- утилиты
    def _add_files(self, paths: List[str]) -> None:
        added = 0
        for path in paths:
            full = os.path.abspath(path)
            if full in self.files:
                continue
            if not is_supported(full):
                self._log(f"пропущен (неподдерживаемый формат): {full}")
                continue
            self.files.append(full)
            self.listbox.insert("end", os.path.basename(full))
            added += 1
        if added:
            if not self.out_var.get():
                folder = os.path.dirname(self.files[0])
                self.out_var.set(default_output_name(folder))
            self._log(f"добавлено файлов: {added}")
        self._update_status()

    def _update_status(self) -> None:
        self.status.configure(text=f"Файлов в списке: {len(self.files)}")

    def _log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _log_clear(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _open(self, path: str) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}" >/dev/null 2>&1 &')
        except Exception:  # noqa: BLE001 - открыть файл не критично
            pass


def main(argv: Optional[List[str]] = None) -> int:
    """Запуск GUI. Аргументы командной строки = файлы для добавления в список."""
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
