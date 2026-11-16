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

from . import APP_NAME, __version__
from .extract import Item
from .pipeline import (
    MODE_COMBINED,
    ConvertOptions,
    ConvertResult,
    convert,
    documents_dir,
    is_supported,
)
from .vedomost import (
    AUTO_SCORE,
    MATCH_ARTICLE,
    MATCH_BARCODE,
    MATCH_NAME,
    MATCH_TITLES,
    Matcher,
    load_vedomost,
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

#: В окне две колонки наименования — «из документа» и «из ведомости» — и
#: процент совпадения: так видно, что было в счёте и что подобрала программа.
#: В сохранённом файле по-прежнему ровно 10 колонок образца (см. TEMPLATE_COLUMNS).
COLUMN_FROM_DOC = "Из документа"
COLUMN_FROM_VED = "Из ведомости"
COLUMN_SCORE = "% совпадения"
#: Колонки, которые показываются только при загруженной ведомости.
_VED_ONLY_COLUMNS = (COLUMN_FROM_VED, COLUMN_SCORE)

PREVIEW_COLUMNS = [
    "Штрихкод",
    "Артикул",
    COLUMN_FROM_DOC,
    COLUMN_FROM_VED,
    COLUMN_SCORE,
    "ед. изм.",
    "количество",
    "цена",
    "сумма",
    "ставка НДС",
    "сумма НДС",
    "сумма с НДС",
]
#: Ширины — в пикселях (у Treeview width именно пиксели, а не символы).
#: Числовые колонки и «ед. изм.» в полтора раза уже, чем были: они короткие,
#: а место нужно двум колонкам наименования.
PREVIEW_WIDTHS = [110, 95, 320, 320, 85, 50, 62, 63, 70, 60, 70, 68]
PREVIEW_MIN_WIDTHS = [85, 70, 150, 150, 60, 40, 45, 50, 55, 45, 55, 60]
#: Показываемые колонки без ведомости: без «из ведомости» и «% совпадения».
PREVIEW_WITHOUT_VED = [
    title for title in PREVIEW_COLUMNS if title not in _VED_ONLY_COLUMNS
]
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

#: «решения нет» — отличается от None, которым помечено «не выбирать»
_NO_OVERRIDE = object()

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
        # сверка с ведомостью: загруженная ведомость, выбранный способ подбора
        # и ручные решения пользователя (они переживают повторную конвертацию)
        self._vedomost = None
        self._ved_path = ""
        self._strategy = MATCH_NAME
        self._matcher: Optional[Matcher] = None
        self._overrides: dict = {}
        # строки таблицы: id строки -> {"item": позиция, "key": ключ сверки}
        self._preview_rows: dict = {}
        # открытый редактор названия (Combobox поверх ячейки)
        self._editor = None
        self._edit_var = None
        self._editing: Optional[str] = None

        self._set_window_icon()
        self._build_ui()
        if initial_files:
            self._add_files(initial_files)
        self._maximize()
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

        # --- шаг 1: файл и ведомость (плашка минимальной высоты)
        top = self._card(body, "ШАГ 1 · Файл и ведомость", gap=6)
        row = ttk.Frame(top, style="Card.TFrame")
        row.pack(fill="x", padx=10, pady=(4, 2))
        ttk.Label(row, text="Файл:", style="Small.TLabel").pack(side="left")
        self.file_var = tk.StringVar(value="файл не выбран")
        ttk.Label(
            row, textvariable=self.file_var, style="File.TLabel", anchor="w"
        ).pack(side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Button(
            row, text="Выбрать файл…", style="Compact.TButton",
            command=self.on_add_files,
        ).pack(side="right")

        vrow = ttk.Frame(top, style="Card.TFrame")
        vrow.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Label(vrow, text="Ведомость:", style="Small.TLabel").pack(side="left")
        self.ved_var = tk.StringVar(value="не загружена")
        ttk.Label(
            vrow, textvariable=self.ved_var, style="VedSmall.TLabel", anchor="w"
        ).pack(side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Button(
            vrow, text="Загрузить ведомость…", style="Compact.TButton",
            command=self.on_load_vedomost,
        ).pack(side="right")

        # --- шаг 2: куда сохранить (тоже одна строка, без пояснений)
        mid = self._card(body, "ШАГ 2 · Куда сохранить", gap=6)
        prow = ttk.Frame(mid, style="Card.TFrame")
        prow.pack(fill="x", padx=10, pady=(4, 4))
        ttk.Label(prow, text="Путь:", style="Small.TLabel").pack(side="left")
        # обычное поле пути, как в любой программе Windows: можно вписать любую
        # папку (в том числе на другом диске), можно стереть — тогда «Документы»
        self.out_var = tk.StringVar(value=documents_dir())
        ttk.Entry(prow, textvariable=self.out_var, style="Path.TEntry").pack(
            side="left", fill="x", expand=True, padx=6
        )
        # «Обзор…» — привычный выбор папки; поле остаётся обычным
        ttk.Button(
            prow, text="Обзор…", style="Compact.TButton", command=self.on_browse_output
        ).pack(side="left")

        # --- кнопки: три способа сверки, затем сохранение и открытие
        act = ttk.Frame(body, style="Card.TFrame")
        act.pack(fill="x", pady=(2, 10))

        convert_row = ttk.Frame(act, style="Card.TFrame")
        convert_row.pack(fill="x")
        self.convert_buttons = {}
        for strategy in (MATCH_BARCODE, MATCH_ARTICLE, MATCH_NAME):
            button = ttk.Button(
                convert_row,
                text="Конвертировать " + MATCH_TITLES[strategy],
                style="Primary.TButton",
                command=lambda s=strategy: self.on_convert(s),
            )
            button.pack(side="left", padx=(0, 8))
            self.convert_buttons[strategy] = button
        # «по наименованию» — основная кнопка (её и нажимают чаще всего)
        self.run_btn = self.convert_buttons[MATCH_NAME]

        tools_row = ttk.Frame(act, style="Card.TFrame")
        tools_row.pack(fill="x", pady=(8, 0))
        self.save_btn = ttk.Button(
            tools_row, text="Сохранить", style="Secondary.TButton",
            command=self.on_save, state="disabled",
        )
        self.save_btn.pack(side="left")
        # подсказка про правку названия прямо в таблице (одна строка, не мешает)
        self.edit_hint = ttk.Label(
            tools_row,
            text="двойной щелчок по «Из ведомости» — изменить название",
            style="Hint.TLabel",
        )
        self.edit_hint.pack(side="left", padx=(14, 0))

        self.open_folder_btn = ttk.Button(
            tools_row, text="Открыть папку", style="Ghost.TButton",
            command=self.on_open_folder, state="disabled",
        )
        self.open_folder_btn.pack(side="right")
        self.open_file_btn = ttk.Button(
            tools_row, text="Открыть файл", style="Ghost.TButton",
            command=self.on_open_result, state="disabled",
        )
        self.open_file_btn.pack(side="right", padx=(0, 8))

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
        # строки, где название выбрано/вписано вручную — светло-жёлтые
        self.preview.tag_configure("manual", background="#FFF7E0")
        # без ведомости двух «лишних» колонок нет — таблица как раньше, 10 колонок
        self.preview.configure(displaycolumns=list(PREVIEW_WITHOUT_VED))
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
        # название можно править прямо в таблице: двойной щелчок по «Из ведомости»
        self.preview.bind("<Double-1>", self._on_preview_double_click)

        self.hint_var = tk.StringVar(value="")
        ttk.Label(
            bottom, textvariable=self.hint_var, style="Hint.TLabel"
        ).pack(anchor="w", padx=14, pady=(0, 10))

        self._clear_preview()

    def _maximize(self) -> None:
        """Развернуть окно на весь экран сразу при запуске.

        Windows понимает state("zoomed"), Linux/X11 — attributes("-zoomed").
        Если оконный менеджер не поддерживает ни то, ни другое, окно просто
        остаётся 1180×780 (и его можно развернуть обычной кнопкой).
        """
        try:
            self.state("zoomed")
            return
        except Exception:  # noqa: BLE001 - не Windows
            pass
        try:
            self.attributes("-zoomed", True)
            return
        except Exception:  # noqa: BLE001 - оконный менеджер не умеет
            pass
        try:
            width = self.winfo_screenwidth()
            height = self.winfo_screenheight()
            self.geometry(f"{width}x{height}+0+0")
        except Exception:  # noqa: BLE001 - не критично
            pass

    def _card(self, parent, title: str, gap: int = 10):
        """Белая карточка раздела с подписью-шагом (gap — отступ снизу)."""
        frame = ttk.LabelFrame(parent, text=title, style="Card.TLabelframe")
        frame.pack(fill="x", pady=(0, gap))
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
                padding=(0, 0),
            )
            # узкие плашки: минимум внутренних отступов у рамки
            style.configure("Card.TLabelframe", padding=(2, 2))
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
                "Ved.TLabel",
                background=CARD,
                foreground=OK_GREEN,
                font=("Segoe UI", 9),
            )
            # компактные кнопки и подписи для узких плашек ШАГ 1 / ШАГ 2
            style.configure(
                "Compact.TButton",
                font=("Segoe UI", 9),
                padding=(10, 2),
                background=CARD,
                foreground=NAVY,
                bordercolor=BORDER,
            )
            style.map("Compact.TButton", background=[("active", "#E8EEF8")])
            style.configure(
                "Small.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 9)
            )
            style.configure(
                "VedSmall.TLabel", background=CARD, foreground=OK_GREEN, font=("Segoe UI", 9)
            )
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
                padding=(2, 3),
                # 8 пт: узкие числовые колонки подписываются без обрезки
                font=("Segoe UI", 8, "bold"),
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

    # ------------------------------------------------------------- ведомость
    def on_load_vedomost(self) -> None:
        """Выбрать файл ведомости остатков ТМЦ (читается в фоне)."""
        path = filedialog.askopenfilename(
            title="Выберите ведомость остатков ТМЦ",
            filetypes=[
                ("Ведомость", "*.xlsx *.xls *.xlsm *.csv"),
                ("Все файлы", "*.*"),
            ],
        )
        if not path:
            return
        self._set_busy(True)
        self.hint_var.set(f"читаю ведомость: {os.path.basename(path)}…")

        def worker() -> None:
            try:
                vedomost = load_vedomost(path)
                self._queue.put(("ved", (path, vedomost)))
            except Exception as exc:  # noqa: BLE001 - показываем пользователю
                self._queue.put(("ved_error", (path, exc)))

        threading.Thread(target=worker, daemon=True).start()

    def load_vedomost(self, path: str) -> bool:
        """Прочитать ведомость синхронно (так же, как фоновая загрузка)."""
        try:
            vedomost = load_vedomost(path)
        except Exception as exc:  # noqa: BLE001 - понятное сообщение пользователю
            self._vedomost_failed(path, exc)
            return False
        self._vedomost_loaded(path, vedomost)
        return True

    def _vedomost_loaded(self, path: str, vedomost) -> None:
        self._set_busy(False)
        self._vedomost = vedomost
        self._ved_path = path
        # ручные решения относились к прошлой ведомости — сбрасываем
        self._overrides = {}
        self._matcher = None
        self.ved_var.set(self._vedomost_text(vedomost))
        hint = (
            f"Ведомость загружена: {os.path.basename(path)}. "
            "Выберите файл и нажмите одну из кнопок «Конвертировать»."
        )
        if vedomost.warnings:
            hint = vedomost.warnings[0]
        self.hint_var.set(hint)
        self._log(f"ведомость: {path} ({vedomost.count} позиций)")

    def _vedomost_failed(self, path: str, exc: Exception) -> None:
        self._set_busy(False)
        message = str(exc) or type(exc).__name__
        self.ved_var.set("не загружена")
        self.hint_var.set("Ведомость не загружена: " + message)
        messagebox.showwarning(APP_NAME, "Не удалось загрузить ведомость\n\n" + message)

    @staticmethod
    def _vedomost_text(vedomost) -> str:
        """Короткая строка о ведомости: имя, позиции, поставщики."""
        name = os.path.basename(getattr(vedomost, "path", "") or "")
        suppliers = len(vedomost.suppliers)
        text = f"{name} — позиций {vedomost.count}"
        if suppliers:
            text += f", поставщиков {suppliers}"
        return text

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
    def on_convert(self, strategy: str = MATCH_NAME) -> None:
        """«Конвертировать по штрихкоду / артикулу / наименованию».

        Без загруженной ведомости все три кнопки делают обычную конвертацию:
        сверять не с чем. С ведомостью — подбирают номенклатуру выбранным
        способом (по поставщику документа) и переименовывают позиции.
        """
        if self._running:
            return
        if not self.files:
            messagebox.showwarning(APP_NAME, "Сначала выберите файл.")
            return
        self._start_convert(self.files[0], strategy, write=False)

    def _start_convert(self, path: str, strategy: str = MATCH_NAME, write: bool = False) -> None:
        """Запустить разбор файла в фоне (write=False — только показать таблицу)."""
        folder = self.output_dir()
        self._strategy = strategy if strategy in MATCH_TITLES else MATCH_NAME
        matcher = None
        if self._vedomost is not None:
            # ручные решения пользователя (выбор/«не выбирать») переживают
            # повторную конвертацию — иначе они бы терялись при каждом нажатии
            matcher = Matcher(self._vedomost, self._strategy, self._overrides)
        self._matcher = matcher
        options = ConvertOptions(
            mode=MODE_COMBINED,
            # в файле ровно те же 10 колонок, что в образце:
            # после «сумма с НДС» ничего не добавляем
            add_source_columns=False,
            # имя всегда «поставщик дата.xlsx» — в выбранной папке или «Документах»
            naming="auto",
            output_dir=folder,
            # «Конвертировать» только показывает таблицу; пишет «Сохранить»
            write=write,
            matcher=matcher,
        )

        self._set_busy(True)
        self.save_btn.configure(state="disabled")
        self._pending_source = path
        self._pending_out = folder
        if not write:
            self._pending_result = None
        self.open_file_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.hint_var.set("")
        self._log_clear()
        self._clear_preview()
        self._log(f"{APP_NAME} {__version__}")
        self._log(f"файл: {os.path.basename(path)}")
        if matcher is not None:
            self._log("сверка с ведомостью: " + matcher.strategy_title)

        def worker() -> None:
            try:
                result = convert([path], folder, options, log=lambda m: self._queue.put(("log", m)))
                self._queue.put(("saved" if write else "done", result))
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
        # в старую папку. Способ сверки тоже берём тот, которым конвертировали.
        self._pending_out = self.output_dir()
        self._start_convert(path, self._strategy, write=True)

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
        self._set_busy(False)
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
                elif kind == "ved":
                    self._vedomost_loaded(*payload)
                elif kind == "ved_error":
                    self._vedomost_failed(*payload)
                elif kind == "error":
                    self._fail(*payload)
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _finish(self, result: ConvertResult) -> None:
        self._set_busy(False)
        self._last_outputs = list(result.outputs)
        self._pending_result = result
        self.save_btn.configure(state="normal" if result.items_count else "disabled")
        self._update_preview_columns()

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
        parts = []
        if result.match_summary:
            parts.append(self._match_hint())
        if planned and result.items_count:
            folder, name = os.path.split(planned)
            parts.append(f"будет сохранено: {folder} / {name} (кнопка «Сохранить»)")
        elif docs:
            parts.append(docs)
        if parts:
            self.hint_var.set(" · ".join(part for part in parts if part))

        if result.items_count:
            text = (
                f"Таблица готова.\n\nПозиций: {result.items_count}\n"
                f"Документ: {docs}\n"
                f"Расхождений в суммах: {result.error_count}\n"
                f"Предупреждений: {result.warning_count}\n"
                f"Значений рассчитано программой: {result.computed_count}"
            )
            if result.match_summary:
                text += "\n\nСверка с ведомостью (" + self._strategy_title() + "): "
                text += self._match_hint()
                if result.match_summary.get("review"):
                    text += (
                        "\nСпорные позиции видны в таблице: там две колонки наименования"
                        " («Из документа» и «Из ведомости») и процент совпадения."
                        " Двойной щелчок по «Из ведомости» — выбрать другое название"
                        " или вписать своё; пустое значение — оставить название документа."
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
        self._set_busy(False)
        self._log("ОШИБКА: " + str(exc))
        self._log(tb)
        messagebox.showerror(APP_NAME, f"Произошла ошибка:\n{exc}")
        self._clear_files()

    # ------------------------------------------------------- сверка/состояние
    def _strategy_title(self) -> str:
        return MATCH_TITLES.get(self._strategy, MATCH_TITLES[MATCH_NAME])

    def _match_hint(self) -> str:
        """Одна строка об итогах сверки с ведомостью."""
        if self._matcher is None:
            return ""
        return self._matcher.summary_text()

    def _set_busy(self, busy: bool) -> None:
        """Пока идёт работа, все кнопки запуска недоступны."""
        self._running = busy
        state = "disabled" if busy else "normal"
        for button in getattr(self, "convert_buttons", {}).values():
            button.configure(state=state)
        if busy:
            self.save_btn.configure(state="disabled")

    # --------------------------------------------------------------- утилиты
    def _show_preview(self, items: List[Item], limit: int = PREVIEW_LIMIT) -> None:
        """Показать таблицу так, как она попадёт в Excel.

        В окне две колонки наименования — «Из документа» и «Из ведомости» — и
        процент совпадения. В сохранённом файле по-прежнему одна колонка
        «Наименование товара» с итоговым (ведомостным) названием.
        """
        self._clear_preview()
        for item in items[:limit]:
            stored = {
                "item": item,
                "key": (item.source, item.sheet, item.row),
            }
            row_id = self.preview.insert(
                "", "end", values=self._preview_values(item, stored), tags=("even",)
            )
            self._preview_rows[row_id] = stored
            self._retag_row(row_id, stored)

        # две «лишние» колонки нужны только при загруженной ведомости
        self._update_preview_columns()
        if not items:
            self.hint_var.set("Позиции не найдены — смотрите сообщение об ошибке.")

    def _update_preview_columns(self) -> None:
        """Показать/скрыть колонки «Из ведомости» и «% совпадения»."""
        with_vedomost = self._vedomost is not None
        columns = PREVIEW_COLUMNS if with_vedomost else PREVIEW_WITHOUT_VED
        try:
            self.preview.configure(displaycolumns=list(columns))
        except Exception:  # noqa: BLE001 - оформление не критично
            pass
        # подсказка про правку нужна только когда есть что править
        try:
            self.edit_hint.configure(
                text="двойной щелчок по «" + COLUMN_FROM_VED + "» — изменить название"
                if with_vedomost
                else ""
            )
        except Exception:  # noqa: BLE001 - не критично
            pass

    def _entry_for(self, key):
        """Результат сверки для позиции (или None, если сверки не было)."""
        if self._matcher is None:
            return None
        return self._matcher.by_key.get(key)

    def _preview_values(self, item: Item, stored: dict) -> List[str]:
        """Значения строки таблицы (12 колонок окна)."""
        entry = self._entry_for(stored["key"])
        override = self._overrides.get(stored["key"], _NO_OVERRIDE)
        from_doc = entry.doc_name if entry is not None else item.name
        if override is not _NO_OVERRIDE:
            if override is None:
                from_ved = ""
                score = "не выбирать"
            elif isinstance(override, str):
                from_ved = override
                score = "вручную"
            else:
                from_ved = override.nomenclature
                score = "вручную"
        elif entry is not None and entry.matched:
            from_ved = entry.name
            score = entry.score_title
        else:
            from_ved = ""
            score = entry.score_title if entry is not None and entry.skipped else ""
        values = [
            item.barcode or "",
            item.article or "",
            from_doc,
            from_ved,
            score,
            item.unit or "",
            item.qty,
            item.price,
            item.total_sum,
            item.vat_rate,
            item.vat_sum,
            item.sum_with_vat,
        ]
        return [self._fmt(value, title) for value, title in zip(values, PREVIEW_COLUMNS)]

    def _retag_row(self, row_id: str, stored: dict) -> None:
        """Чередование строк + подсветка ручных правок."""
        try:
            index = list(self._preview_rows).index(row_id)
        except ValueError:
            index = 0
        if stored["key"] in self._overrides:
            tags = ("manual",)
        else:
            tags = ("odd",) if index % 2 else ("even",)
        try:
            self.preview.item(row_id, tags=tags)
        except Exception:  # noqa: BLE001 - не критично
            pass

    # ------------------------------------------------- правка названия в таблице
    def _on_preview_double_click(self, event) -> None:
        """Двойной щелчок по ячейке «Из ведомости» открывает поле правки."""
        if self._vedomost is None or self._matcher is None:
            return
        try:
            row_id = self.preview.identify_row(event.y)
            column = self.preview.identify_column(event.x)
        except Exception:  # noqa: BLE001 - старая версия Tk
            return
        if not row_id or not column:
            return
        try:
            index = int(str(column).lstrip("#")) - 1
        except ValueError:
            return
        if index < 0 or index >= len(PREVIEW_COLUMNS):
            return
        if PREVIEW_COLUMNS[index] != COLUMN_FROM_VED:
            return
        self.edit_cell(row_id)

    def edit_cell(self, row_id: str) -> None:
        """Открыть редактор названия поверх ячейки «Из ведомости»."""
        stored = self._preview_rows.get(row_id)
        if stored is None:
            return
        entry = self._entry_for(stored["key"])
        if entry is None:
            # сверки для этой позиции нет — править нечего
            return
        try:
            bbox = self.preview.bbox(row_id, f"#{PREVIEW_COLUMNS.index(COLUMN_FROM_VED) + 1}")
        except Exception:  # noqa: BLE001 - не критично
            bbox = None
        if not bbox:
            return
        x, y, width, height = bbox
        suggestions = [row.nomenclature for _score, row in self._suggestions(entry)]
        self._edit_var = tk.StringVar(value=self._current_from_ved(stored))
        self._editor = ttk.Combobox(
            self.preview, textvariable=self._edit_var, values=suggestions
        )
        try:
            self._editor.place(x=x, y=y, width=max(width, 280), height=height)
            self._editor.focus_set()
            self._editor.bind("<Return>", lambda _e: self.commit_edit())
            self._editor.bind("<FocusOut>", lambda _e: self.commit_edit())
            self._editor.bind("<Escape>", lambda _e: self.cancel_edit())
        except Exception:  # noqa: BLE001 - без редактора таблица остаётся
            self._editor = None
            return
        self._editing = row_id

    def _suggestions(self, entry):
        """Похожие номенклатуры ведомости — подсказкой в выпадающем списке."""
        if self._vedomost is None:
            return []
        supplier = ""
        if self._matcher is not None:
            supplier = str(self._matcher.summary.get("supplier_key") or "")
        return self._vedomost.candidates(entry.doc_name, supplier, limit=10)

    def _current_from_ved(self, stored: dict) -> str:
        entry = self._entry_for(stored["key"])
        override = self._overrides.get(stored["key"], _NO_OVERRIDE)
        if override is not _NO_OVERRIDE:
            if override is None:
                return ""
            if isinstance(override, str):
                return override
            return override.nomenclature
        if entry is not None and entry.matched:
            return entry.name
        return ""

    def _destroy_editor(self) -> None:
        editor = getattr(self, "_editor", None)
        self._editor = None
        self._editing = None
        if editor is not None:
            try:
                editor.destroy()
            except Exception:  # noqa: BLE001 - уже уничтожен
                pass
        try:
            self.preview.focus_set()
        except Exception:  # noqa: BLE001 - не критично
            pass

    def cancel_edit(self) -> None:
        self._destroy_editor()

    def commit_edit(self) -> None:
        """Записать выбранное/вписанное название и перерисовать строку."""
        row_id = self._editing
        if row_id is None:
            return
        text = ""
        try:
            text = self._edit_var.get() if self._edit_var is not None else ""
        except Exception:  # noqa: BLE001 - не критично
            text = ""
        self._destroy_editor()
        stored = self._preview_rows.get(row_id)
        if stored is None:
            return
        entry = self._entry_for(stored["key"])
        if entry is None:
            return
        text = " ".join(str(text or "").split())
        if not text:
            self._overrides[stored["key"]] = None  # пусто = «не выбирать»
        else:
            row = self._find_ved_row(entry, text)
            # строка ведомости — переименовываем по ней; своего названия нет —
            # оставляем вписанный текст как есть
            self._overrides[stored["key"]] = row if row is not None else text
        self._apply_edit_locally(stored, entry)
        self._refresh_row(row_id, stored)
        self.hint_var.set(
            "Название изменено вручную — не забудьте «Сохранить»."
        )

    def _find_ved_row(self, entry, text: str):
        """Найти номенклатуру ведомости по вписанному названию (если есть)."""
        if self._vedomost is None:
            return None
        exact = self._vedomost.by_name(text)
        if exact is not None:
            return exact
        supplier = ""
        if self._matcher is not None:
            supplier = str(self._matcher.summary.get("supplier_key") or "")
        found = self._vedomost.candidates(text, supplier, limit=1)
        if found and found[0][0] >= AUTO_SCORE:
            return found[0][1]
        return None

    def _apply_edit_locally(self, stored: dict, entry) -> None:
        """Сразу показать правку в таблице (файл пишется при «Сохранить»)."""
        item = stored["item"]
        override = self._overrides.get(stored["key"], _NO_OVERRIDE)
        if override is _NO_OVERRIDE or override is None:
            item.name = entry.doc_name
            item.article = entry.doc_article or ""
            item.barcode = entry.doc_barcode or ""
            return
        if isinstance(override, str):
            item.name = override
            item.article = entry.doc_article or ""
            item.barcode = entry.doc_barcode or ""
            return
        item.name = override.nomenclature
        item.article = entry.doc_article or override.article or ""
        item.barcode = entry.doc_barcode or override.barcode or ""
        if not entry.doc_article and override.unit and not item.unit:
            item.unit = override.unit

    def _refresh_row(self, row_id: str, stored: dict) -> None:
        values = self._preview_values(stored["item"], stored)
        try:
            self.preview.item(row_id, values=values)
        except Exception:  # noqa: BLE001 - не критично
            pass
        self._retag_row(row_id, stored)

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
        self._destroy_editor()
        self._preview_rows.clear()
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
        self.hint_var.set("Файл выбран — нажмите «Конвертировать по наименованию»")

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
