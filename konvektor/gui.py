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
from typing import List, Optional, Tuple

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
from .pick import (
    CHOICE_CUSTOM,
    SourceDialog,
    _VedChooser,
    install_styles,
)
from .writer import plain_number
from . import settings as app_settings

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
COLUMN_PICK = "▼"
COLUMN_SCORE = "% совпадения"
#: Колонки, которые показываются только при загруженной ведомости.
_VED_ONLY_COLUMNS = (COLUMN_FROM_VED, COLUMN_PICK, COLUMN_SCORE)

PREVIEW_COLUMNS = [
    "Штрихкод",
    "Артикул",
    COLUMN_FROM_DOC,
    COLUMN_FROM_VED,
    COLUMN_PICK,
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
PREVIEW_WIDTHS = [110, 95, 320, 320, 26, 85, 50, 62, 63, 70, 60, 70, 68]
PREVIEW_MIN_WIDTHS = [85, 70, 150, 150, 24, 60, 40, 45, 50, 55, 45, 55, 60]
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
#: Обновлённая палитра 1.9.17 — чуть светлее, мягче контраст, современный вид
NAVY = "#1E3A5F"
NAVY_LIGHT = "#2B4C7A"
NAVY_HOVER = "#264A7A"
BG = "#F2F6FC"
BG_SOFT = "#EAF0FA"
CARD = "#FFFFFF"
BORDER = "#D4DFEE"
BORDER_LIGHT = "#E6EDF7"
TEXT = "#1A2B42"
MUTED = "#64748B"
ACCENT = "#2563EB"
ACCENT_HOVER = "#1D4ED8"
ACCENT_DARK = "#1E3A8A"
OK_GREEN = "#059669"
OK_GREEN_LIGHT = "#ECFDF5"
ROW_ODD = "#F8FAFC"
ROW_HOVER = "#EFF6FF"


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
        # склад и поставщик, выбранные в диалоге при загрузке файла
        self._sklad = ""
        self._supplier_choice = ""
        # открытый выпадающий список номенклатуры ведомости
        self._chooser = None
        self._chooser_row: Optional[str] = None
        # сторож открытого окна: если строка «уехала» — закрываем
        self._popup_row: Optional[str] = None
        self._popup_y = 0
        # открытый редактор названия (Combobox поверх ячейки)
        self._editor = None
        self._edit_var = None
        self._editing: Optional[str] = None
        # сохранение полей таблицы
        self._saved_fields: dict = {}
        self._column_widths_override: dict = {}

        self._set_window_icon()
        self._build_ui()
        self._restore_settings()
        if initial_files:
            self._add_files(initial_files)
        self._maximize()
        self.after(120, self._poll_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

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

        # --- шапка: иконка, название, подзаголовок + все кнопки действий
        header = tk.Frame(self, background=NAVY, height=76)
        header.pack(fill="x")
        header.pack_propagate(False)

        icon = self._header_icon(header)
        if icon is not None:
            icon.pack(side="left", padx=(14, 10), pady=12)

        titles = tk.Frame(header, background=NAVY)
        titles.pack(side="left", anchor="w", pady=12)
        ttk.Label(
            titles, text=APP_NAME, style="Header.TLabel"
        ).pack(anchor="w")
        ttk.Label(
            titles,
            text=f"версия {__version__}",
            style="HeaderSub.TLabel",
        ).pack(anchor="w")

        # переменные состояния (используются в обработчиках)
        self.file_var = tk.StringVar(value="файл не выбран")
        self.ved_var = tk.StringVar(value="не загружена")
        self.out_var = tk.StringVar(value=documents_dir())
        self.source_var = tk.StringVar(value="не выбраны")

        # панель кнопок на синем фоне (сгруппирована в выпадающие меню)
        nav = tk.Frame(header, background=NAVY)
        nav.pack(side="right", padx=14, pady=18)

        # 1. Меню «Файл и папка ▾»
        file_mb = ttk.Menubutton(nav, text="Файл и ведомость ▾", style="Header.TMenubutton")
        file_menu = tk.Menu(file_mb, tearoff=0)
        file_menu.add_command(label="Выбрать файл…", command=self.on_add_files)
        file_menu.add_command(label="Загрузить ведомость…", command=self.on_load_vedomost)
        file_menu.add_separator()
        file_menu.add_command(label="Склад и поставщик…", command=self.on_choose_source)
        file_menu.add_command(label="Папка сохранения…", command=self.on_browse_output)
        file_mb["menu"] = file_menu
        file_mb.pack(side="left", padx=(0, 6))

        # 2. Меню «Конвертировать ▾»
        self.convert_mb = ttk.Menubutton(nav, text="Конвертировать ▾", style="HeaderPrimary.TMenubutton")
        convert_menu = tk.Menu(self.convert_mb, tearoff=0)
        self.convert_buttons = {}
        for strategy in (MATCH_BARCODE, MATCH_ARTICLE, MATCH_NAME):
            convert_menu.add_command(
                label="Конвертировать " + MATCH_TITLES[strategy],
                command=lambda s=strategy: self.on_convert(s),
            )
            self.convert_buttons[strategy] = self.convert_mb
        self.convert_mb["menu"] = convert_menu
        self.convert_mb.pack(side="left", padx=(0, 6))
        self.run_btn = self.convert_mb

        # 3. Кнопка «Сохранить»
        self.save_btn = ttk.Button(
            nav, text="Сохранить", style="Secondary.TButton",
            command=self.on_save, state="disabled",
        )
        self.save_btn.pack(side="left", padx=(0, 6))

        # 4. Кнопка «Очистить таблицу»
        self.clear_btn = ttk.Button(
            nav, text="Очистить таблицу", style="Secondary.TButton",
            command=self.on_clear_table,
        )
        self.clear_btn.pack(side="left", padx=(0, 6))

        # 4. Кнопки «Открыть файл» и «Открыть папку» в меню «Открыть ▾»
        open_mb = ttk.Menubutton(nav, text="Открыть ▾", style="Header.TMenubutton")
        open_menu = tk.Menu(open_mb, tearoff=0)
        open_menu.add_command(label="Открыть файл", command=self.on_open_result, state="disabled")
        open_menu.add_command(label="Открыть папку", command=self.on_open_folder, state="disabled")
        open_mb["menu"] = open_menu
        open_mb.pack(side="left", padx=(0, 6))

        # 5. Меню «Таблица ▾» — сохранение полей таблицы
        table_mb = ttk.Menubutton(nav, text="Таблица ▾", style="Header.TMenubutton")
        table_menu = tk.Menu(table_mb, tearoff=0)
        table_menu.add_command(label="Сохранить поля таблицы", command=self.on_save_table_fields)
        table_menu.add_command(label="Загрузить поля таблицы", command=self.on_load_table_fields)
        table_menu.add_separator()
        table_menu.add_command(label="Сбросить ширину колонок", command=self.on_reset_column_widths)
        table_menu.add_command(label="Сохранить настройки", command=self.on_save_settings)
        table_mb["menu"] = table_menu
        table_mb.pack(side="left")
        self.table_mb = table_mb

        class _MenuItemProxy:
            def __init__(self, menu: tk.Menu, index: int) -> None:
                self._menu = menu
                self._index = index

            def configure(self, **kw) -> None:
                self._menu.entryconfigure(self._index, **kw)

            def cget(self, key: str):
                return self._menu.entrycget(self._index, key)

        self.open_file_btn = _MenuItemProxy(open_menu, 0)
        self.open_folder_btn = _MenuItemProxy(open_menu, 1)

        # --- сразу под синим заголовком — таблица
        body = tk.Frame(self, background=BG)
        body.pack(fill="both", expand=True, padx=10, pady=(8, 8))

        self.edit_hint = ttk.Label(
            body,
            text="двойной щелчок по «Из ведомости» — изменить название",
            style="Hint.TLabel",
        )

        table_wrap = ttk.Frame(body, style="Card.TFrame")
        table_wrap.pack(fill="both", expand=True)

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
        # кратковременная подсветка при копировании штрихкода
        self.preview.tag_configure("copied", background="#C8E6C9")
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

        yscroll = ttk.Scrollbar(grid_box, orient="vertical", command=self._preview_yview)
        xscroll = ttk.Scrollbar(grid_box, orient="horizontal", command=self._preview_xview)
        self.preview.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        yscroll.pack(side="right", fill="y")
        xscroll.pack(side="bottom", fill="x")
        self.preview.pack(side="left", fill="both", expand=True)
        # «▾» справа от названия открывает список номенклатуры ведомости
        self.preview.bind("<Button-1>", self._on_preview_click)
        self.preview.bind("<Double-1>", self._on_preview_double_click)
        # прокрутка и колесо закрывают открытый список/поле ввода: иначе они
        # «висят» поверх уехавшей строки
        self.preview.bind("<MouseWheel>", self._on_preview_wheel)
        self.preview.bind("<Button-4>", self._on_preview_wheel)
        self.preview.bind("<Button-5>", self._on_preview_wheel)

        status_row = ttk.Frame(body, style="Card.TFrame")
        status_row.pack(fill="x", pady=(4, 0))
        self.hint_var = tk.StringVar(value="")
        ttk.Label(
            status_row, textvariable=self.hint_var, style="Hint.TLabel"
        ).pack(side="left", padx=4)
        ttk.Label(
            status_row, textvariable=self.ved_var, style="Hint.TLabel"
        ).pack(side="right", padx=4)
        ttk.Label(
            status_row, text="Ведомость:", style="Hint.TLabel"
        ).pack(side="right")

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
            install_styles(style)
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
            style.map("Compact.TButton", background=[("active", BG_SOFT), ("pressed", BORDER)])
            style.configure(
                "Small.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 9)
            )
            style.configure(
                "VedSmall.TLabel", background=CARD, foreground=OK_GREEN, font=("Segoe UI", 9)
            )
            style.configure("Path.TEntry", padding=(6, 5), fieldbackground=CARD)
            style.configure("TSeparator", background=BORDER)
            style.configure("Header.TMenubutton", background=NAVY, foreground="#FFFFFF", font=("Segoe UI", 9, "bold"), padding=(12, 6))
            style.map("Header.TMenubutton", background=[("active", NAVY_HOVER), ("pressed", NAVY_LIGHT)])
            style.configure("HeaderPrimary.TMenubutton", background=ACCENT, foreground="#FFFFFF", font=("Segoe UI", 9, "bold"), padding=(14, 7))
            style.map("HeaderPrimary.TMenubutton", background=[("active", ACCENT_HOVER), ("pressed", ACCENT_DARK)])
            style.configure("Secondary.TButton", background=CARD, foreground=NAVY, font=("Segoe UI", 9), padding=(12, 5), bordercolor=BORDER)
            style.map("Secondary.TButton", background=[("active", BG_SOFT), ("pressed", BORDER_LIGHT)], foreground=[("disabled", MUTED)])
        except Exception:  # noqa: BLE001 - оформление не должно ломать окно
            pass

    def _setup_table_style(self) -> None:
        """Разлиновка таблицы: подсветка строк через одну, объёмная шапка, высота строк."""
        try:
            style = ttk.Style(self)
            style.configure(
                "Grid.Treeview",
                background=CARD,
                fieldbackground=CARD,
                foreground=TEXT,
                rowheight=24,
                borderwidth=1,
                bordercolor=BORDER,
                lightcolor=BORDER,
                darkcolor=BORDER,
                relief="solid",
                font=("Segoe UI", 9),
            )
            style.configure(
                "Grid.Treeview.Cell",
                borderwidth=1,
                relief="solid",
            )
            style.configure(
                "Grid.Treeview.Heading",
                background=BG_SOFT,
                foreground=NAVY,
                relief="raised",
                borderwidth=1,
                padding=(4, 6),
                font=("Segoe UI", 8, "bold"),
            )
            style.map("Grid.Treeview.Heading", background=[("active", BORDER_LIGHT)])
            style.map("Grid.Treeview", background=[("selected", ROW_HOVER)], foreground=[("selected", NAVY)])
        except Exception:  # noqa: BLE001 - оформление не должно ломать окно
            pass

    # ------------------------------------------------- сохранение полей таблицы
    def _save_table_fields(self) -> dict:
        """Собрать текущие поля таблицы для сохранения."""
        fields = {
            "columns": list(self.preview.columns_list) if hasattr(self.preview, "columns_list") else PREVIEW_COLUMNS,
            "widths": {},
            "displaycolumns": list(self.preview.cget("displaycolumns")) if hasattr(self.preview, "cget") else [],
        }
        try:
            for col in PREVIEW_COLUMNS:
                try:
                    w = self.preview.column(col, "width")
                    fields["widths"][col] = int(w) if str(w).isdigit() else w
                except Exception:
                    pass
        except Exception:
            pass
        fields["sklad"] = self._sklad
        fields["supplier"] = self._supplier_choice
        fields["output_dir"] = self.out_var.get()
        fields["ved_path"] = self._ved_path
        return fields

    def on_save_table_fields(self) -> None:
        """Сохранить поля таблицы в настройки."""
        try:
            fields = self._save_table_fields()
            data = app_settings.load_settings()
            data["table_fields"] = fields
            data["column_widths"] = fields["widths"]
            data["sklad"] = self._sklad
            data["supplier"] = self._supplier_choice
            data["output_dir"] = self.out_var.get()
            app_settings.save_settings(data)
            self.hint_var.set("Поля таблицы сохранены ✓")
            try:
                from tkinter import messagebox
                messagebox.showinfo(APP_NAME, "Поля таблицы сохранены.\nОни восстановятся при следующем запуске.")
            except Exception:
                pass
        except Exception as e:
            self.hint_var.set(f"Не удалось сохранить: {e}")

    def on_load_table_fields(self) -> None:
        """Загрузить сохранённые поля таблицы."""
        try:
            data = app_settings.load_settings()
            fields = data.get("table_fields") or {}
            widths = data.get("column_widths") or fields.get("widths") or {}
            if widths:
                for col, w in widths.items():
                    try:
                        self.preview.column(col, width=int(w))
                    except Exception:
                        pass
                self.hint_var.set("Поля таблицы загружены ✓")
            else:
                self.hint_var.set("Сохранённых полей нет")
        except Exception as e:
            self.hint_var.set(f"Не удалось загрузить: {e}")

    def on_reset_column_widths(self) -> None:
        for col, w in zip(PREVIEW_COLUMNS, PREVIEW_WIDTHS):
            try:
                self.preview.column(col, width=w)
            except Exception:
                pass
        self.hint_var.set("Ширина колонок сброшена")

    def on_save_settings(self) -> None:
        self.on_save_table_fields()

    def _restore_settings(self) -> None:
        """Восстановить настройки при запуске."""
        try:
            data = app_settings.load_settings()
            if data.get("output_dir"):
                self.out_var.set(data["output_dir"])
            if data.get("sklad"):
                self._sklad = data["sklad"]
            if data.get("supplier"):
                self._supplier_choice = data["supplier"]
                self.source_var.set(self.source_text())
            widths = data.get("column_widths") or {}
            if widths:
                for col, w in widths.items():
                    try:
                        self.preview.column(col, width=int(w))
                    except Exception:
                        pass
            geom = data.get("window_geometry")
            if geom:
                try:
                    self.geometry(geom)
                except Exception:
                    pass
        except Exception:
            pass

    def _on_close(self) -> None:
        try:
            data = app_settings.load_settings()
            data["window_geometry"] = self.geometry()
            data["output_dir"] = self.out_var.get()
            data["sklad"] = self._sklad
            data["supplier"] = self._supplier_choice
            # сохранить ширины колонок
            widths = {}
            for col in PREVIEW_COLUMNS:
                try:
                    widths[col] = self.preview.column(col, "width")
                except Exception:
                    pass
            data["column_widths"] = widths
            data["table_fields"] = self._save_table_fields()
            app_settings.save_settings(data)
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

    # ------------------------------------------------------------- действия
    # ------------------------------------------------- склад и поставщик
    def suppliers_of_vedomost(self) -> List[str]:
        """Поставщики из загруженной ведомости — для списка в диалоге."""
        if self._vedomost is None:
            return []
        try:
            found = list(self._vedomost.suppliers)
        except Exception:  # noqa: BLE001 - не критично
            return []
        # ведомость отдаёт пары «поставщик, позиций» — в списке нужно имя
        names = []
        for item in found:
            name = item[0] if isinstance(item, tuple) else item
            name = str(name or "").strip()
            if name and name not in names:
                names.append(name)
        return names

    def on_choose_source(self) -> None:
        """Кнопка «Изменить…»: спросить склад и поставщика."""
        self.ask_source(force=True)

    def ask_source(self, force: bool = False) -> bool:
        """Спросить склад и поставщика (диалог поверх окна)."""
        dialog = SourceDialog(
            self,
            self.suppliers_of_vedomost(),
            self._sklad,
            self._supplier_choice,
            on_done=self._source_chosen,
        )
        self._source_dialog = dialog
        try:
            dialog.win.grab_set()  # пока диалог открыт, окно ждёт ответа
        except Exception:  # noqa: BLE001 - модальность не критична
            pass
        if force:
            self.hint_var.set("Выберите склад и поставщика — папка и сверка пойдут по ним.")
        try:
            dialog.win.wait_window()
        except Exception:  # noqa: BLE001 - без ожидания результат тот же
            pass
        return True

    def _source_chosen(self, sklad: str, supplier: str) -> None:
        self._sklad = sklad or ""
        self._supplier_choice = supplier or ""
        self._matcher = None
        self.source_var.set(self.source_text())
        if self._sklad:
            self.hint_var.set(
                "Результат будет сохранён в папке "
                f"«{self._sklad}{os.sep}{self._supplier_choice or 'Поставщик не определён'}»."
            )
        else:
            self.hint_var.set("Склад не выбран — файл сохранится сразу в папку «Путь».")

    def source_text(self) -> str:
        """Строка «склад / поставщик» под полем пути."""
        if not self._sklad:
            return "не выбран (как в документе)"
        supplier = self._supplier_choice or "как в документе"
        return f"{self._sklad} / {supplier}"

    def on_add_files(self) -> None:
        path = filedialog.askopenfilename(title="Выберите файл", filetypes=FILE_TYPES)
        if not path:
            return
        self._add_files([path])
        # заказчик просил спрашивать склад и поставщика именно при загрузке файла
        self.ask_source()

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
        # выбрать поставщика теперь можно из этой ведомости — обновляем строку
        self.source_var.set(self.source_text())
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
            matcher = Matcher(
                self._vedomost,
                self._strategy,
                self._overrides,
                supplier=self._supplier_choice,
                sklad=self._sklad,
            )
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
            sklad=self._sklad,
            # поставщик для папки: выбранный в диалоге (пусто — из документа)
            supplier=self._supplier_choice,
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
        # файл остаётся в буфере до нажатия «Очистить таблицу» — можно
        # поочерёдно конвертировать разными способами без повторной загрузки

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
                        " Щёлкните «▼» справа от названия — выберите похожую позицию"
                        " ведомости или «не выбирать (новый товар)»: тогда в файле"
                        " будет слово «Новинка»."
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
        """Показать/скрыть колонки «Из ведомости», «▾» и «% совпадения»."""
        with_vedomost = self._vedomost is not None
        columns = PREVIEW_COLUMNS if with_vedomost else PREVIEW_WITHOUT_VED
        try:
            self.preview.configure(displaycolumns=list(columns))
        except Exception:  # noqa: BLE001 - оформление не критично
            pass
        # подсказка про правку нужна только когда есть что править
        try:
            self.edit_hint.configure(
                text="«" + COLUMN_PICK + "» справа от названия — выбрать позицию ведомости"
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
        if not score and entry is not None:
            score = "Новинка"
        values = [
            item.barcode or "",
            item.article or "",
            from_doc,
            from_ved,
            COLUMN_PICK if self._vedomost is not None else "",
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

    # ------------------------------------- выбор номенклатуры из ведомости
    def _on_preview_click(self, event) -> None:
        """Одиночный щелчок: по штрихкоду — копирует в буфер, по «▾» — открывает список."""
        target = self._cell_at(event)
        if target is None:
            return
        row_id, title = target
        if title == "Штрихкод":
            stored = self._preview_rows.get(row_id)
            if stored is not None:
                barcode = str(stored["item"].barcode or "").strip()
                if barcode:
                    try:
                        self.clipboard_clear()
                        self.clipboard_append(barcode)
                        self.update()
                    except Exception:  # noqa: BLE001
                        pass
                    try:
                        box = self.preview.bbox(row_id, "#1")
                        if box:
                            x, y, w, h = box
                            parent = self.preview.master or self.preview
                            flash = tk.Label(
                                parent,
                                text=barcode,
                                background="#C8E6C9",
                                foreground="#1B5E20",
                                font=("Segoe UI", 11, "bold"),
                                relief="solid",
                                bd=1,
                                anchor="w",
                                padx=4,
                            )
                            flash.place(x=x, y=y, width=w, height=h)
                            flash.lift()
                            self.after(1000, flash.destroy)
                    except Exception:  # noqa: BLE001
                        pass
                    self.hint_var.set(f"✓ Штрихкод {barcode} скопирован в буфер обмена.")
            return
        if self._vedomost is None or self._matcher is None:
            return
        if title != COLUMN_PICK:
            return
        self.open_chooser(row_id)

    def _on_preview_double_click(self, event) -> None:
        """Двойной щелчок по «Из ведомости»/«▾» — тот же список."""
        if self._vedomost is None or self._matcher is None:
            return
        target = self._cell_at(event)
        if target is None:
            return
        row_id, title = target
        if title not in (COLUMN_PICK, COLUMN_FROM_VED):
            return
        self.open_chooser(row_id)

    def _cell_at(self, event):
        """(id строки, название колонки) под курсором или None."""
        try:
            row_id = self.preview.identify_row(event.y)
            column = self.preview.identify_column(event.x)
        except Exception:  # noqa: BLE001 - старая версия Tk
            return None
        if not row_id or not column:
            return None
        try:
            index = int(str(column).lstrip("#")) - 1
        except ValueError:
            return None
        if index < 0 or index >= len(PREVIEW_COLUMNS):
            return None
        return row_id, PREVIEW_COLUMNS[index]

    def open_chooser(self, row_id: str) -> bool:
        """Показать под ячейкой список похожих позиций ведомости с процентами.

        Первый пункт — «не выбирать (новый товар)», последний — «Своё
        название…» (ввод текста вручную).
        """
        stored = self._preview_rows.get(row_id)
        if stored is None or self._vedomost is None:
            return False
        entry = self._entry_for(stored["key"])
        if entry is None:
            return False
        self._close_chooser()
        self._cancel_edit()

        items = [("— не выбирать (новый товар) —", None)]
        for score, row in self._suggestions(entry):
            label = f"{score:.0%} · {row.nomenclature}"
            if row.article:
                label += f" · арт. {row.article}"
            if row.barcode:
                label += f" · штрихкод {row.barcode}"
            items.append((label, row))
        items.append(("Своё название…", CHOICE_CUSTOM))

        bbox = None
        try:
            bbox = self.preview.bbox(
                row_id, f"#{PREVIEW_COLUMNS.index(COLUMN_FROM_VED) + 1}"
            )
        except Exception:  # noqa: BLE001 - не критично
            bbox = None
        x = y = 0
        width = 460
        if bbox:
            x = int(bbox[0])
            y = int(bbox[1] + bbox[3])
            width = max(360, int(bbox[2]) + 60)
        try:
            left = self.preview.winfo_rootx() + x
            top = self.preview.winfo_rooty() + y
        except Exception:  # noqa: BLE001 - не критично
            left = top = 100
        self._chooser_row = row_id
        self._chooser = _VedChooser(self, left, top, width, items, self._chooser_picked)
        self._watch_popup(row_id, bbox)
        return True

    def _chooser_picked(self, value) -> None:
        """Пользователь выбрал строку в списке."""
        row_id = self._chooser_row
        self._close_chooser()
        if row_id is None:
            return
        stored = self._preview_rows.get(row_id)
        if stored is None:
            return
        entry = self._entry_for(stored["key"])
        if entry is None:
            return
        if value is CHOICE_CUSTOM:
            # своё название: обычное поле ввода поверх ячейки
            self.edit_cell(row_id)
            return
        self._overrides[stored["key"]] = value
        self._apply_edit_locally(stored, entry)
        self._refresh_row(row_id, stored)
        what = "не выбирать (новый товар)" if value is None else getattr(value, "nomenclature", "")
        self.hint_var.set(f"Выбрано: {what} — не забудьте «Сохранить».")

    def _close_chooser(self) -> None:
        chooser, self._chooser = self._chooser, None
        self._chooser_row = None
        if chooser is not None:
            self._cancel_popup_watch()
        if chooser is not None:
            try:
                chooser.close()
            except Exception:  # noqa: BLE001 - уже закрыт
                pass

    def _close_popups(self) -> None:
        """Закрыть всё, что «висит» над таблицей (список и поле ввода)."""
        self._close_chooser()
        self._cancel_edit()

    def _on_preview_wheel(self, event):
        """Колесо мыши: закрыть список/поле и прокрутить таблицу."""
        self._close_popups()
        delta = 0
        if getattr(event, "num", None) == 4:
            delta = -1
        elif getattr(event, "num", None) == 5:
            delta = 1
        elif getattr(event, "delta", 0):
            delta = -1 if event.delta > 0 else 1
        if delta:
            try:
                self.preview.yview_scroll(delta, "units")
            except Exception:  # noqa: BLE001 - не критично
                pass
        return "break"

    def _preview_yview(self, *args):
        self._close_popups()
        return self.preview.yview(*args)

    def _preview_xview(self, *args):
        self._close_popups()
        return self.preview.xview(*args)

    def _cancel_edit(self) -> None:
        """Убрать поле ввода без сохранения (прокрутка, Esc, новый выбор)."""
        self._destroy_editor()

    def edit_cell(self, row_id: str) -> None:
        """Поле ввода названия поверх ячейки «Из ведомости» (свой текст)."""
        stored = self._preview_rows.get(row_id)
        if stored is None:
            return
        entry = self._entry_for(stored["key"])
        if entry is None:
            return
        try:
            bbox = self.preview.bbox(row_id, f"#{PREVIEW_COLUMNS.index(COLUMN_FROM_VED) + 1}")
        except Exception:  # noqa: BLE001 - не критично
            bbox = None
        if not bbox:
            return
        x, y, width, height = bbox
        suggestions = [row.nomenclature for _score, row in self._suggestions(entry)]
        self._cancel_edit()
        self._edit_var = tk.StringVar(value=self._current_from_ved(stored))
        self._editor = ttk.Combobox(
            self.preview, textvariable=self._edit_var, values=suggestions
        )
        try:
            self._editor.place(x=x, y=y, width=max(width, 280), height=height)
            self._editor.focus_set()
            self._editor.bind("<Return>", lambda _e: self.commit_edit())
            self._editor.bind("<Escape>", lambda _e: self._cancel_edit())
            self._editor.bind("<MouseWheel>", self._on_preview_wheel)
        except Exception:  # noqa: BLE001 - без поля таблица остаётся
            self._editor = None
            return
        self._editing = row_id
        self._watch_popup(row_id, bbox)

    # ------------------------------------- «окно не должно висеть на месте»
    def _watch_popup(self, row_id: str, bbox) -> None:
        """Запомнить, где было окно, и следить, не уехала ли строка."""
        self._popup_row = row_id
        self._popup_y = int(bbox[1]) if bbox else 0
        self._popup_watch = None
        try:
            self._popup_watch = self.after(150, self._check_popup)
        except Exception:  # noqa: BLE001 - не критично
            self._popup_watch = None

    def _check_popup(self) -> None:
        """Закрыть список/поле, если строка прокрутилась или исчезла."""
        if self._chooser is None and getattr(self, "_editor", None) is None:
            self._popup_row = None
            return
        row_id = self._popup_row
        moved = True
        if row_id:
            try:
                box = self.preview.bbox(row_id, "#1")
            except Exception:  # noqa: BLE001 - не критично
                box = ""
            moved = not box or int(box[1]) != int(self._popup_y)
        if moved:
            self._close_popups()
            return
        try:
            self._popup_watch = self.after(150, self._check_popup)
        except Exception:  # noqa: BLE001 - не критично
            self._popup_watch = None

    def _cancel_popup_watch(self) -> None:
        watch, self._popup_watch = getattr(self, "_popup_watch", None), None
        self._popup_row = None
        if watch is not None:
            try:
                self.after_cancel(watch)
            except Exception:  # noqa: BLE001 - уже сработал
                pass

    def _suggestions(self, entry):
        """Похожие номенклатуры ведомости — подсказкой в выпадающем списке.

        Показывает только номенклатуру того склада, куда загружается документ,
        и только выбранного поставщика(ов) — как просил заказчик: «при выборе
        птичкой иной номенклатуры должна отображаться номенклатура только этого
        склада куда был загружен счет и этого поставщика».
        """
        if self._vedomost is None:
            return []
        # поставщик(и) — из сводки Matcher (может быть несколько)
        suppliers: List[str] = []
        if self._matcher is not None:
            keys = self._matcher.summary.get("supplier_keys") or []
            if keys:
                suppliers = [str(k) for k in keys if k]
            else:
                single = self._matcher.summary.get("supplier_key") or ""
                if single:
                    suppliers = [str(single)]
        if not suppliers:
            return self._vedomost.candidates(entry.doc_name, "", limit=10, sklad=self._sklad)
        # несколько поставщиков — объединяем результаты и сортируем по сходству
        seen: dict = {}
        merged: List[Tuple[float, object]] = []
        for sup in suppliers:
            for score, row in self._vedomost.candidates(entry.doc_name, sup, limit=10, sklad=self._sklad):
                key = (row.nomenclature, row.article)
                if key in seen and seen[key] >= score:
                    continue
                seen[key] = score
                merged.append((score, row))
        merged.sort(key=lambda pair: (-pair[0], pair[1].nomenclature.lower()))
        # дедупликация как в candidates уже частично сделана, но нужна и здесь
        unique: List[Tuple[float, object]] = []
        seen_norm = set()
        for score, row in merged:
            k = (self._vedomost._norm_names[row.index] if hasattr(row, "index") else row.nomenclature.lower(), row.article)
            if k in seen_norm:
                continue
            seen_norm.add(k)
            unique.append((score, row))
            if len(unique) >= 10:
                break
        return unique

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

    def cancel_edit(self) -> None:
        """Публичный отказ от правки (Esc в окне)."""
        self._destroy_editor()

    def _destroy_editor(self) -> None:
        editor = getattr(self, "_editor", None)
        if editor is not None:
            self._cancel_popup_watch()
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
        """Найти номенклатуру ведомости по вписанному названию (если есть).

        Фильтрует по складу и поставщику — чтобы свободный ввод не
        подставил товар чужого склада (см. требование заказчика). Если
        выбрано несколько поставщиков — ищет среди всех выбранных.
        """
        if self._vedomost is None:
            return None
        suppliers: List[str] = []
        if self._matcher is not None:
            keys = self._matcher.summary.get("supplier_keys") or []
            if keys:
                suppliers = [str(k) for k in keys if k]
            else:
                single = self._matcher.summary.get("supplier_key") or ""
                if single:
                    suppliers = [str(single)]
        # точное совпадение — с проверкой склада и поставщика
        exact = self._vedomost.by_name(text)
        if exact is not None and self._vedomost._sklad_ok(exact.sklad, self._sklad):
            if not suppliers or exact.supplier_key in suppliers:
                return exact
        # нечёткий поиск среди нужных поставщиков и склада
        for sup in (suppliers or [""]):
            found = self._vedomost.candidates(text, sup, limit=1, sklad=self._sklad)
            if found and found[0][0] >= AUTO_SCORE:
                return found[0][1]
        # последняя попытка: точное без учёта склада
        if exact is not None and (not suppliers or exact.supplier_key in suppliers):
            return exact
        return None

    def _apply_edit_locally(self, stored: dict, entry) -> None:
        """Сразу показать правку в таблице (файл пишется при «Сохранить»)."""
        item = stored["item"]
        override = self._overrides.get(stored["key"], _NO_OVERRIDE)
        if override is not _NO_OVERRIDE:
            # «не выбирать» = товара нет в ведомости -> в файле слово «Новинка»
            item.novinka = override is None
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
        self._close_popups()
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

    def on_clear_table(self) -> None:
        """Очистить таблицу и сбросить загруженный файл из буфера."""
        if self._running:
            return
        self._clear_preview()
        self._clear_files()
        self._pending_source = None
        self._pending_result = None
        self.save_btn.configure(state="disabled")
        self.open_file_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.hint_var.set("Таблица очищена.")

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
