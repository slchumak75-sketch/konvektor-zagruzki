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
import subprocess
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
    documents_dir,
    is_supported,
)
from .writer import plain_number

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
#: Деньги — всегда две цифры после запятой.
_MONEY_PREVIEW = {"цена", "сумма", "сумма НДС", "сумма с НДС"}
#: Количество и ставка НДС — целые, если дробной части нет (5, а не 5,00; 20, а не 20,00).
_WHOLE_PREVIEW = {"количество", "ставка НДС"}

#: Сколько строк показывать в предпросмотре (остальные — в файле).
PREVIEW_LIMIT = 500

#: Оформление окна: спокойный светлый фон, тёмная шапка, один акцент.
NAVY = "#1F3A5F"
NAVY_LIGHT = "#2A4C77"
BG = "#F4F6FA"
CARD = "#FFFFFF"
BORDER = "#D8E0EC"
TEXT = "#1B2733"
MUTED = "#5B6B7F"
ACCENT = "#2F6FED"
ACCENT_DARK = "#1F55C0"
OK_GREEN = "#1E8E5A"
ROW_ODD = "#F7F9FC"


class KonvektorApp(_Base):
    """Главное окно."""

    def __init__(self, initial_files: Optional[List[str]] = None):
        super().__init__()
        self.title(f"{APP_NAME} {__version__}")
        self.geometry("1180x780")
        self.minsize(980, 640)
        try:
            self.configure(background=BG)
        except Exception:  # noqa: BLE001 - оформление не критично
            pass

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
        self._setup_theme()

        # --- шапка: иконка, название, подзаголовок
        header = tk.Frame(self, background=NAVY, height=76)
        header.pack(fill="x")
        header.pack_propagate(False)

        icon = self._header_icon(header)
        if icon is not None:
            icon.pack(side="left", padx=(18, 12), pady=12)

        titles = tk.Frame(header, background=NAVY)
        titles.pack(side="left", anchor="w", pady=12)
        ttk.Label(
            titles, text=APP_NAME, style="Header.TLabel"
        ).pack(anchor="w")
        ttk.Label(
            titles,
            text="счета, накладные и прайсы → одна таблица по образцу",
            style="HeaderSub.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            header, text=f"версия {__version__}", style="HeaderVer.TLabel"
        ).pack(side="right", padx=18)

        body = tk.Frame(self, background=BG)
        body.pack(fill="both", expand=True, padx=16, pady=14)

        # --- шаг 1: файл
        top = self._card(body, "ШАГ 1 · Файл для конвертации (по одному)")
        row = ttk.Frame(top, style="Card.TFrame")
        row.pack(fill="x", padx=14, pady=(10, 4))
        self.file_var = tk.StringVar(value="файл не выбран")
        ttk.Label(
            row, textvariable=self.file_var, style="File.TLabel", anchor="w"
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(
            row, text="Выбрать файл…", style="Secondary.TButton",
            command=self.on_add_files,
        ).pack(side="right")
        ttk.Label(
            top,
            text="Поддерживаются .xls, .xlsx, .csv, .txt, .html. "
            "Файл можно перетащить мышью на ярлык Konvektor.exe.",
            style="Hint.TLabel",
        ).pack(anchor="w", padx=14, pady=(0, 10))

        # --- шаг 2: куда сохранить
        mid = self._card(body, "ШАГ 2 · Куда сохранить")
        prow = ttk.Frame(mid, style="Card.TFrame")
        prow.pack(fill="x", padx=14, pady=(10, 4))
        ttk.Label(prow, text="Путь:", style="Card.TLabel").pack(side="left")
        # обычное поле пути, как в любой программе Windows: можно вписать любую
        # папку (в том числе на другом диске), можно стереть — тогда «Документы»
        self.out_var = tk.StringVar(value=documents_dir())
        ttk.Entry(prow, textvariable=self.out_var, style="Path.TEntry").pack(
            side="left", fill="x", expand=True, padx=8
        )
        # «Обзор…» — привычный выбор папки; поле остаётся обычным
        ttk.Button(
            prow, text="Обзор…", style="Secondary.TButton", command=self.on_browse_output
        ).pack(side="left")

        ttk.Label(
            mid,
            text="Впишите путь, вставьте его из проводника или нажмите «Обзор…». "
            "Пусто — папка «Документы». Имя файла программа задаёт сама: "
            "«поставщик + сегодняшняя дата.xlsx» (например, «ООО ОМА 15.11.2026.xlsx»); "
            "если такой файл уже есть, новый сохранится как «… (2).xlsx» — "
            "старый не затирается.",
            style="Hint.TLabel",
            wraplength=1060,
            justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 10))

        # --- кнопки действий
        act = ttk.Frame(body, style="Card.TFrame")
        act.pack(fill="x", pady=(2, 10))

        self.run_btn = ttk.Button(
            act, text="Конвертировать", style="Primary.TButton", command=self.on_convert
        )
        self.run_btn.pack(side="left")
        self.save_btn = ttk.Button(
            act, text="Сохранить", style="Secondary.TButton",
            command=self.on_save, state="disabled",
        )
        self.save_btn.pack(side="left", padx=(10, 0))

        self.open_file_btn = ttk.Button(
            act, text="Открыть файл", style="Ghost.TButton",
            command=self.on_open_result, state="disabled",
        )
        self.open_file_btn.pack(side="left", padx=(10, 0))
        self.open_folder_btn = ttk.Button(
            act, text="Открыть папку", style="Ghost.TButton",
            command=self.on_open_folder, state="disabled",
        )
        self.open_folder_btn.pack(side="left", padx=(8, 0))

        # --- шаг 3: предпросмотр результата вместо журнала
        bottom = self._card(body, "ШАГ 3 · Результат (таблица)")
        bottom.pack_configure(fill="both", expand=True)

        table_wrap = ttk.Frame(bottom, style="Card.TFrame")
        table_wrap.pack(fill="both", expand=True, padx=14, pady=(10, 6))

        self._setup_table_style()
        # рамка вокруг таблицы не зависит от темы оформления Windows
        grid_box = tk.Frame(
            table_wrap, bd=1, relief="solid", highlightthickness=0,
            background=BORDER,
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
        self.preview.tag_configure("odd", background=ROW_ODD)
        self.preview.tag_configure("even", background=CARD)
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
        ttk.Label(
            bottom, textvariable=self.hint_var, style="Hint.TLabel"
        ).pack(anchor="w", padx=14, pady=(0, 10))

        self._clear_preview()

    def _card(self, parent, title: str):
        """Белая карточка раздела с подписью-шагом."""
        frame = ttk.LabelFrame(parent, text=title, style="Card.TLabelframe")
        frame.pack(fill="x", pady=(0, 10))
        return frame

    def _header_icon(self, parent):
        """Иконка в шапке (та же, что у .exe). Если картинки нет — просто пропускаем."""
        image = getattr(self, "_icon_image", None)
        if image is None:
            return None
        try:
            # исходник 256×256 — для шапки берём 64×64
            self._icon_header = image.subsample(4, 4)
            return tk.Label(parent, image=self._icon_header, background=NAVY, bd=0)
        except Exception:  # noqa: BLE001 - иконка не критична
            return None

    def _setup_theme(self) -> None:
        """Единое оформление: шрифты, цвета, кнопки, карточки."""
        try:
            style = ttk.Style(self)
            if "clam" in style.theme_names():
                # clam позволяет задавать цвета кнопок и рамок (vista их игнорирует)
                style.theme_use("clam")
            style.configure(
                ".", background=BG, foreground=TEXT, font=("Segoe UI", 10)
            )
            style.configure("Card.TFrame", background=CARD)
            style.configure("Card.TLabel", background=CARD, foreground=TEXT)
            style.configure(
                "Card.TLabelframe",
                background=CARD,
                bordercolor=BORDER,
                relief="solid",
                borderwidth=1,
            )
            style.configure(
                "Card.TLabelframe.Label",
                background=BG,
                foreground=NAVY,
                font=("Segoe UI", 9, "bold"),
            )
            style.configure(
                "Header.TLabel",
                background=NAVY,
                foreground="#FFFFFF",
                font=("Segoe UI", 16, "bold"),
            )
            style.configure(
                "HeaderSub.TLabel",
                background=NAVY,
                foreground="#BFD0E6",
                font=("Segoe UI", 9),
            )
            style.configure(
                "HeaderVer.TLabel",
                background=NAVY,
                foreground="#8FA9C8",
                font=("Segoe UI", 8),
            )
            style.configure(
                "File.TLabel",
                background=CARD,
                foreground=TEXT,
                font=("Segoe UI", 11, "bold"),
            )
            style.configure("Hint.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))
            style.configure(
                "Primary.TButton",
                font=("Segoe UI", 10, "bold"),
                padding=(20, 9),
                background=ACCENT,
                foreground="#FFFFFF",
                borderwidth=0,
            )
            style.map(
                "Primary.TButton",
                background=[("disabled", "#B9C7DA"), ("active", ACCENT_DARK)],
                foreground=[("disabled", "#F0F4FA")],
            )
            style.configure(
                "Secondary.TButton",
                font=("Segoe UI", 10),
                padding=(16, 8),
                background=CARD,
                foreground=NAVY,
                bordercolor=BORDER,
            )
            style.map("Secondary.TButton", background=[("active", "#E8EEF8")])
            style.configure(
                "Ghost.TButton",
                font=("Segoe UI", 9),
                padding=(12, 7),
                background=BG,
                foreground=MUTED,
                borderwidth=0,
            )
            style.map("Ghost.TButton", background=[("active", "#E6ECF5")])
            style.configure("Path.TEntry", padding=(6, 5), fieldbackground=CARD)
            style.configure("TSeparator", background=BORDER)
        except Exception:  # noqa: BLE001 - оформление не должно ломать окно
            pass

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
        """Выбрать папку в диалоге «Обзор…» (имя файла задаёт программа)."""
        current = self.output_dir()
        start = current if os.path.isdir(current) else documents_dir()
        path = filedialog.askdirectory(
            title="Куда сохранять результат", initialdir=start, mustexist=False
        )
        if path:
            self.out_var.set(os.path.normpath(path))

    def output_dir(self) -> str:
        """Папка сохранения из поля «Путь» (пустая строка — «Документы»).

        Можно вписать любую папку, в том числе на другом диске. Если вставили
        полный путь к файлу — берём папку из него: имя всё равно будет
        «поставщик дата.xlsx», как и просили.
        """
        raw = self.out_var.get().strip().strip('"').strip("'").strip()
        if not raw:
            return ""
        # «%USERPROFILE%\Documents», «~\Documents», путь из проводника в кавычках
        value = os.path.expandvars(os.path.expanduser(raw)).strip()
        if not value:
            return ""
        if os.path.isfile(value):
            # указали файл, а не папку — берём папку из него
            value = os.path.dirname(value)
        elif os.path.splitext(value)[1].lower() in (".xlsx", ".xls", ".xlsm", ".csv"):
            value = os.path.dirname(value)
        if not value:
            return ""
        # «D:\Отчёты\» -> «D:\Отчёты»; «.» -> текущая папка
        return os.path.normpath(value)

    # ------------------------------------------------------------ конвертация
    def on_convert(self) -> None:
        if self._running:
            return
        if not self.files:
            messagebox.showwarning(APP_NAME, "Сначала выберите файл.")
            return

        path = self.files[0]
        folder = self.output_dir()
        options = ConvertOptions(
            mode=MODE_COMBINED,
            # в файле ровно те же 10 колонок, что в образце:
            # после «сумма с НДС» ничего не добавляем
            add_source_columns=False,
            # имя всегда «поставщик дата.xlsx» — в выбранной папке или «Документах»
            naming="auto",
            output_dir=folder,
            # «Конвертировать» только показывает таблицу; пишет «Сохранить»
            write=False,
        )

        self._running = True
        self.run_btn.configure(state="disabled")
        self.save_btn.configure(state="disabled")
        self._pending_source = path
        self._pending_out = folder
        self._pending_result = None
        self.open_file_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.hint_var.set("")
        self._log_clear()
        self._clear_preview()
        self._log(f"{APP_NAME} {__version__}")
        self._log(f"файл: {os.path.basename(path)}")

        def worker() -> None:
            try:
                result = convert([path], folder, options, log=lambda m: self._queue.put(("log", m)))
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

        # ВАЖНО: путь читаем из поля заново — пользователь мог изменить его
        # после «Конвертировать». Раньше здесь брался снимок, и файл уходил
        # в старую папку.
        folder = self.output_dir()
        self._pending_out = folder
        options = ConvertOptions(
            mode=MODE_COMBINED,
            add_source_columns=False,
            naming="auto",
            output_dir=folder,
            write=True,
        )

        self._running = True
        self.save_btn.configure(state="disabled")
        self.run_btn.configure(state="disabled")

        def worker() -> None:
            try:
                result = convert([path], folder, options, log=lambda m: self._queue.put(("log", m)))
                self._queue.put(("saved", result))
            except Exception as exc:  # noqa: BLE001 - показываем пользователю
                self._queue.put(("error", (exc, traceback.format_exc())))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def on_open_result(self) -> None:
        """Открыть сохранённый файл в Excel (или чем он открыт в системе)."""
        path = self._last_outputs[0] if self._last_outputs else ""
        if not path or not os.path.exists(path):
            messagebox.showinfo(APP_NAME, "Сначала нажмите «Сохранить» — файла ещё нет.")
            return
        self._open_path(path)

    def on_open_folder(self) -> None:
        """Открыть папку, в которую сохранён результат."""
        path = self._last_outputs[0] if self._last_outputs else ""
        folder = os.path.dirname(path) if path else (self.output_dir() or documents_dir())
        if not os.path.isdir(folder):
            messagebox.showinfo(APP_NAME, "Папки пока нет — она появится после сохранения.")
            return
        self._open_path(folder)

    @staticmethod
    def _open_path(path: str) -> None:
        """Открыть файл или папку средствами системы (Windows: os.startfile)."""
        try:
            if hasattr(os, "startfile"):
                os.startfile(path)  # type: ignore[attr-defined]  # только Windows
                return
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen([opener, path])  # noqa: S603,S607 - путь наш, не из сети
        except Exception as exc:  # noqa: BLE001 - не критично для работы
            messagebox.showinfo(APP_NAME, f"Не удалось открыть:\n{path}\n\n{exc}")

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
        self.open_file_btn.configure(state="normal")
        self.open_folder_btn.configure(state="normal")

        planned = result.planned_output
        if planned and os.path.abspath(planned) != os.path.abspath(written):
            # записать по указанному пути не удалось (нет прав / диск отключён)
            reason = next((n for n in result.notes if "не удалось записать" in n), "")
            messagebox.showwarning(
                APP_NAME,
                f"По указанному пути записать не удалось:\n{planned}\n\n"
                f"Файл сохранён в:\n{written}\n\n{reason}".strip(),
            )
            return
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
            cells = [
                self._fmt(value, title)
                for value, title in zip(values, PREVIEW_COLUMNS)
            ]
            self.preview.insert("", "end", values=cells, tags=(tag,))
        if not items:
            self.hint_var.set("Позиции не найдены — смотрите сообщение об ошибке.")

    @staticmethod
    def _fmt(value, column: str = "") -> str:
        """Как число будет выглядеть в Excel: деньги — 2 знака, количество — без нулей."""
        if value is None:
            return ""
        if isinstance(value, bool):
            return str(value).lower()
        if isinstance(value, int):
            return str(value)
        if column in _WHOLE_PREVIEW:
            number = plain_number(value)
            if isinstance(number, str):
                return number
            value = number
        if isinstance(value, float):
            if column in _WHOLE_PREVIEW:
                if value.is_integer():
                    return str(int(value))
                return f"{value:.3f}".rstrip("0").rstrip(".").replace(".", ",")
            if column in _MONEY_PREVIEW:
                return f"{value:.2f}".replace(".", ",")
            return f"{value:.2f}".rstrip("0").rstrip(".").replace(".", ",")
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
            self.hint_var.set("")
            self._log(f"выбран файл: {full}")
            break
        self._update_status()

    def _clear_files(self) -> None:
        self.files.clear()
        self.file_var.set("файл не выбран")
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
