"""Диалог «склад и поставщик» — спрашивается при выборе файла.

Заказчик просил: при загрузке файла программа спрашивает **склад**
(«Мир» или «Кореличи») и даёт выбрать **поставщика** из тех, что есть в
загруженной ведомости. Склад и поставщик задают папку результата
(``<папка>/<склад>/<поставщик>/<дата>.xlsx``) и то, среди чьих позиций искать
номенклатуру в ведомости.
"""

from __future__ import annotations

from typing import Callable, Iterable, List, Optional

try:  # pragma: no cover - зависит от сборки Python
    import tkinter as tk
    from tkinter import ttk

    TK_AVAILABLE = True
except Exception:  # noqa: BLE001 - без tkinter диалог не нужен
    tk = None  # type: ignore[assignment]
    ttk = None  # type: ignore[assignment]
    TK_AVAILABLE = False

#: Склады заказчика.
SKLADS = ("Мир", "Кореличи")
#: Пункт «как в документе» убран по просьбе заказчика (1.9.18) — остаётся
#: константой для совместимости, но в диалоге не показывается.
DOC_SUPPLIER = "— как в документе —"

BG = "#F2F6FC"
BG_SOFT = "#EAF0FA"
CARD = "#FFFFFF"
TEXT = "#1A2B42"
MUTED = "#64748B"
NAVY = "#1E3A5F"
NAVY_LIGHT = "#2B4C7A"
BORDER = "#D4DFEE"
BORDER_LIGHT = "#E6EDF7"
ACCENT = "#2563EB"
ACCENT_HOVER = "#1D4ED8"


class SourceDialog:
    """Маленькое окно: склад + поставщик перед конвертацией."""

    def __init__(
        self,
        master,
        suppliers: Iterable[str] = (),
        sklad: str = "",
        supplier: str = "",
        on_done: Optional[Callable[[str, str], None]] = None,
    ) -> None:
        self.master = master
        self.on_done = on_done
        self.suppliers: List[str] = [s for s in suppliers if str(s).strip()]
        self.win = tk.Toplevel(master)
        try:
            self.win.title("Склад и поставщик")
            self.win.configure(background=BG)
            self.win.transient(master)
            try:
                self.win.resizable(False, False)
                self.win.attributes("-topmost", True)
                self.win.after(300, lambda: self.win.attributes("-topmost", False))
            except Exception:
                pass
            screen_h = self.win.winfo_screenheight()
            screen_w = self.win.winfo_screenwidth()
            if screen_h and screen_h > 200:
                # строго по высоте экрана — как таблица, оставляем запас под панель задач и приподнимаем низ
                full_h = max(520, screen_h - 80)
                x = max(0, (screen_w - 640) // 2) if screen_w else 100
                y = 10
                self.win.geometry(f"640x{full_h}+{x}+{y}")
                self._locked_x = x
                self._locked_y = y
                self._locked_h = full_h
                try:
                    self.win.bind("<Configure>", self._lock_position)
                except Exception:
                    pass
            else:
                self.win.geometry("640x700+100+20")
        except Exception:  # noqa: BLE001 - оформление не критично
            pass

        self.sklad_var = tk.StringVar(value=sklad or SKLADS[0])
        # поставщики сортируются по алфавиту (требование 1.9.18)
        self.suppliers = sorted(self.suppliers, key=str.casefold)
        self.supplier_var = tk.StringVar(value=supplier or "")
        self._build()

    def _lock_position(self, event=None) -> None:
        """Не давать двигать окно мышкой — возвращаем на зафиксированное место."""
        try:
            if getattr(event, "widget", None) is not self.win:
                return
            geom = self.win.geometry()
            wanted = f"640x{self._locked_h}+{self._locked_x}+{self._locked_y}"
            if "+" in geom:
                parts = geom.split("+")
                if len(parts) == 3:
                    cur_x, cur_y = int(parts[1]), int(parts[2])
                    if cur_x != self._locked_x or cur_y != self._locked_y:
                        self.win.geometry(wanted)
                        return
                wh = parts[0]
                if "x" in wh:
                    cur_w, cur_h = wh.split("x")
                    if int(cur_w) != 640 or int(cur_h) != int(self._locked_h):
                        self.win.geometry(wanted)
        except Exception:
            pass

    # ------------------------------------------------------------------ вёрстка
    def _build(self) -> None:
        pad = {"padx": 20}
        # Используем grid, чтобы кнопки внизу были ВСЕГДА видны независимо от высоты списка.
        # Колонка 0 растягивается, строка со списком (row=4) растягивается.
        self.win.columnconfigure(0, weight=1)
        self.win.rowconfigure(4, weight=1)

        # 0: заголовок
        header = tk.Frame(self.win, background=NAVY, height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Label(header, text="  \U0001F4E6  Склад и поставщик", background=NAVY, foreground="#FFFFFF", font=("Segoe UI", 12, "bold"), anchor="w").pack(fill="both", expand=True, padx=14, pady=10)

        # 1: карточка склада — поуже по высоте
        card = tk.Frame(self.win, background=CARD, highlightbackground=BORDER, highlightthickness=1, bd=0)
        card.grid(row=1, column=0, sticky="ew", padx=16, pady=(14, 8))
        tk.Label(card, text="Выберите склад — куда положить результат", background=CARD, foreground=NAVY, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14, pady=(8, 2))
        sklad_row = tk.Frame(card, background=CARD)
        sklad_row.pack(fill="x", padx=12, pady=(4, 8))
        self.sklad_buttons = {}
        for name in SKLADS:
            icon = "\U0001F3E2" if name == "Мир" else "\U0001F3ED"
            is_selected = self.sklad_var.get() == name
            btn = None
            try:
                ButtonCls = getattr(tk, "Button", None)
                if ButtonCls is not None:
                    btn = ButtonCls(
                        sklad_row,
                        text=f"{icon}  {name}",
                        font=("Segoe UI", 10, "bold" if is_selected else "normal"),
                        background=ACCENT if is_selected else CARD,
                        foreground="#FFFFFF" if is_selected else NAVY,
                        activebackground=ACCENT_HOVER if is_selected else BG_SOFT,
                        relief="flat" if is_selected else "solid",
                        bd=1,
                        highlightbackground=BORDER,
                        padx=12,
                        pady=4,
                        cursor="hand2",
                        command=lambda n=name: self._select_sklad(n),
                    )
                else:
                    raise AttributeError("no tk.Button")
            except Exception:
                try:
                    btn = ttk.Radiobutton(
                        sklad_row,
                        text=f"{icon}  {name}",
                        value=name,
                        variable=self.sklad_var,
                        style="Sklad.TRadiobutton",
                    )
                except Exception:
                    btn = tk.Frame(sklad_row)  # type: ignore
            try:
                btn.pack(side="left", padx=8, expand=True, fill="x", ipady=2)
            except Exception:
                pass
            self.sklad_buttons[name] = btn
        tk.Label(card, text="Файл попадёт в папку «склад / поставщик / дата.xlsx»", background=CARD, foreground=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=14, pady=(0, 6))

        # 2: заголовок поставщиков
        ttk.Label(self.win, text="  Поставщик (по алфавиту):", background=BG, foreground=TEXT, font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", **pad)

        # 3: поиск как в Excel
        search_frame = tk.Frame(self.win, background=BG)
        search_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=(6, 4))
        search_frame.columnconfigure(1, weight=1)
        self.supplier_search_var = tk.StringVar()
        try:
            tk.Label(search_frame, text="\U0001F50D", background=BG, foreground=MUTED, font=("Segoe UI", 9)).grid(row=0, column=0, padx=(0, 6))
            self.supplier_search_entry = ttk.Entry(search_frame, textvariable=self.supplier_search_var, style="Pick.Search.TEntry", font=("Segoe UI", 9))
            self.supplier_search_entry.grid(row=0, column=1, sticky="ew")
            self.search_count_label = tk.Label(search_frame, text="", background=BG, foreground=MUTED, font=("Segoe UI", 8))
            self.search_count_label.grid(row=0, column=2, padx=(8, 0))
            # крестик для очистки выбранных поставщиков (для следующей загрузки)
            self.clear_supplier_btn = ttk.Button(search_frame, text="✕", width=3, style="Pick.Tool.TButton", command=self._clear_supplier_selection)
            self.clear_supplier_btn.grid(row=0, column=3, padx=(6, 0))
            try:
                self.clear_supplier_btn.configure(cursor="hand2")
            except Exception:
                pass
            self._search_placeholder = "Поиск поставщика…"
            self._search_has_placeholder = True
            self.supplier_search_entry.insert(0, self._search_placeholder)
            try:
                self.supplier_search_entry.configure(foreground=MUTED)
            except Exception:
                pass
            self.supplier_search_entry.bind("<FocusIn>", self._on_search_focus_in)
            self.supplier_search_entry.bind("<FocusOut>", self._on_search_focus_out)
            self.supplier_search_entry.bind("<KeyRelease>", self._on_supplier_key)
            self.supplier_search_entry.bind("<Escape>", self._on_search_escape)
        except Exception:
            self.supplier_search_var = tk.StringVar()  # type: ignore

        # 4: список поставщиков — растягивается, без фиксированной высоты
        list_frame = tk.Frame(self.win, background=CARD, highlightbackground=BORDER, highlightthickness=1, bd=0)
        list_frame.grid(row=4, column=0, sticky="nsew", padx=16, pady=(4, 6))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        canvas = tk.Canvas(list_frame, highlightthickness=0, background=CARD)
        vscroll = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, style="Pick.TFrame")
        inner.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vscroll.grid(row=0, column=1, sticky="ns")
        self._supplier_canvas = canvas
        self._supplier_inner = inner

        initial_set = {
            s.strip() for s in (self.supplier_var.get() or "").split(",") if s.strip()
        }
        self.supplier_vars: dict[str, tk.BooleanVar] = {}
        self.supplier_checks: dict[str, ttk.Checkbutton] = {}
        for name in self.suppliers:
            checked = name in initial_set
            var = tk.BooleanVar(value=checked)
            cb = ttk.Checkbutton(
                inner,
                text=name,
                variable=var,
                style="Supplier.TCheckbutton",
            )
            cb.pack(anchor="w", padx=8, pady=1, fill="x")
            self.supplier_vars[name] = var
            self.supplier_checks[name] = cb
        if not self.suppliers:
            ttk.Label(inner, text="Загрузите ведомость — здесь появятся поставщики", style="PickHint.TLabel").pack(anchor="w", padx=8, pady=6)
        self._update_search_count(len(self.suppliers), len(self.suppliers))

        # 5: кнопки — ВСЕГДА внизу, видны (grid row=5 weight=0)
        buttons = ttk.Frame(self.win, style="Pick.TFrame")
        buttons.grid(row=5, column=0, sticky="ew", pady=(10, 12), padx=20)
        self.ok_btn = ttk.Button(buttons, text="✓  Продолжить", style="PickPrimary.TButton", command=self.on_ok)
        self.ok_btn.pack(side="right")
        self.cancel_btn = ttk.Button(buttons, text="Отмена", style="Pick.TButton", command=self.on_cancel)
        self.cancel_btn.pack(side="right", padx=(0, 10))
        self.hint_label = None
        self._hint_value = self._hint()

    def _on_search_focus_in(self, _event=None) -> None:
        try:
            if getattr(self, "_search_has_placeholder", False):
                self.supplier_search_entry.delete(0, "end")
                self.supplier_search_entry.configure(foreground=TEXT)
                self._search_has_placeholder = False
                self.supplier_search_var.set("")
        except Exception:
            pass

    def _on_search_focus_out(self, _event=None) -> None:
        try:
            if not self.supplier_search_var.get().strip():
                self.supplier_search_entry.delete(0, "end")
                self.supplier_search_entry.insert(0, self._search_placeholder)
                self.supplier_search_entry.configure(foreground=MUTED)
                self._search_has_placeholder = True
        except Exception:
            pass

    def _on_search_escape(self, _event=None) -> None:
        try:
            self.supplier_search_var.set("")
            self.supplier_search_entry.delete(0, "end")
            self.supplier_search_entry.insert(0, self._search_placeholder)
            self.supplier_search_entry.configure(foreground=MUTED)
            self._search_has_placeholder = True
            self._apply_supplier_filter("")
        except Exception:
            pass
        return "break"

    def _on_supplier_key(self, _event=None) -> None:
        # как в Excel — фильтр на каждый KeyRelease
        try:
            if getattr(self, "_search_has_placeholder", False):
                return
            q = self.supplier_search_var.get()
            self._apply_supplier_filter(q)
        except Exception:
            pass

    def _apply_supplier_filter(self, query: str) -> None:
        q = (query or "").strip().lower()
        # учитываем placeholder
        if q == self._search_placeholder.lower():
            q = ""
        visible = 0
        for name, cb in self.supplier_checks.items():
            show = not q or q in name.lower()
            try:
                if show:
                    cb.pack(anchor="w", padx=8, pady=1, fill="x")
                    visible += 1
                else:
                    cb.pack_forget()
            except Exception:
                # fake_tk has no pack_forget — просто считаем, без реального скрытия
                if show:
                    try:
                        cb.pack(anchor="w", padx=8, pady=1, fill="x")
                    except Exception:
                        pass
                    visible += 1
        self._update_search_count(visible, len(self.supplier_checks))
        # прокрутка в начало как в Excel
        try:
            self._supplier_canvas.yview_moveto(0)
        except Exception:
            pass

    def _update_search_count(self, shown: int, total: int) -> None:
        try:
            if hasattr(self, "search_count_label"):
                if shown == total:
                    self.search_count_label.configure(text=f"{total}")
                else:
                    self.search_count_label.configure(text=f"{shown} из {total}")
        except Exception:
            pass

    def _clear_supplier_selection(self) -> None:
        """Сбросить выбранных поставщиков — для следующей загрузки."""
        try:
            for var in self.supplier_vars.values():
                var.set(False)
            # важно: supplier_var хранит итоговый список через запятую — его тоже сбросить
            try:
                self.supplier_var.set("")
            except Exception:
                pass
            # очистить поиск и показать всех
            self.supplier_search_var.set("")
            try:
                self.supplier_search_entry.delete(0, "end")
                self.supplier_search_entry.insert(0, self._search_placeholder)
                self.supplier_search_entry.configure(foreground=MUTED)
                self._search_has_placeholder = True
            except Exception:
                pass
            self._apply_supplier_filter("")
        except Exception:
            pass

    def _select_sklad(self, name: str) -> None:
        """Наглядный выбор склада — подсветка активной кнопки."""
        self.sklad_var.set(name)
        for sklad_name, btn in self.sklad_buttons.items():
            is_sel = sklad_name == name
            try:
                btn.configure(
                    background=ACCENT if is_sel else CARD,
                    foreground="#FFFFFF" if is_sel else NAVY,
                    font=("Segoe UI", 10, "bold" if is_sel else "normal"),
                    relief="flat" if is_sel else "solid",
                )
            except Exception:
                pass

    def _hint(self) -> str:
        if not self.suppliers:
            return (
                "Ведомость не загружена — выберите склад, поставщик будет из документа. "
                "Загрузите ведомость, чтобы выбрать поставщика из списка."
            )
        return (
            f"В ведомости {len(self.suppliers)} поставщик(ов) по алфавиту. "
            "Отметьте своего — номенклатура подберётся только среди его позиций, "
            "а файл попадёт в папку «склад / поставщик»."
        )

    # ------------------------------------------------------------------ данные
    def _on_supplier_toggle(self, name: str) -> None:
        pass

    @property
    def supplier_value(self) -> str:
        """Выбранный поставщик (пустая строка — берём из документа, если ничего не выбрано)."""
        checked = [
            name
            for name, var in self.supplier_vars.items()
            if var.get()
        ]
        if not checked:
            return ""
        return ", ".join(sorted(checked, key=str.casefold))

    @property
    def sklad_value(self) -> str:
        return str(self.sklad_var.get() or "").strip()

    # ---------------------------------------------------------------- события
    def on_ok(self) -> None:
        sklad, supplier = self.sklad_value, self.supplier_value
        self.close()
        if self.on_done is not None:
            self.on_done(sklad, supplier)

    def on_cancel(self) -> None:
        self.close()

    def close(self) -> None:
        try:
            self.win.destroy()
        except Exception:  # noqa: BLE001 - окно могло уже закрыться
            pass


#: Пункт выпадающего списка номенклатуры: вписать своё название.
CHOICE_CUSTOM = object()


def install_styles(style) -> None:
    """Завести в ttk-тему стили диалога и выпадающего списка — современный вид 1.9.17."""
    try:
        style.configure("Pick.TFrame", background=BG)
        style.configure("Pick.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("PickHint.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 8))
        style.configure("Pick.TButton", font=("Segoe UI", 10), padding=(14, 6), background=CARD, foreground=TEXT)
        style.map("Pick.TButton", background=[("active", BG_SOFT)])
        style.configure("PickPrimary.TButton", font=("Segoe UI", 10, "bold"), padding=(18, 7), background=ACCENT, foreground="#FFFFFF")
        style.map("PickPrimary.TButton", background=[("active", ACCENT_HOVER)])
        style.configure("Sklad.TRadiobutton", font=("Segoe UI", 10), padding=(8, 4))
        style.configure("Supplier.TCheckbutton", font=("Segoe UI", 11), padding=(2, 1))
        style.configure("Pick.Search.TEntry", fieldbackground="#FFFFFF", padding=(8, 6))
        style.configure("Pick.Tool.TButton", font=("Segoe UI", 9), padding=(4, 2))
        style.configure("Chooser.Toplevel", background=CARD)
        style.configure("Chooser.TLabel", background="#F8FAFC", foreground=MUTED, font=("Segoe UI", 8))
        style.configure("Chooser.TButton", font=("Segoe UI", 9), padding=(6, 3))
        style.configure("Chooser.Search.TEntry", fieldbackground="#FFFFFF", padding=(8, 6))
        style.configure("Chooser.Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT, rowheight=27, font=("Segoe UI", 11))
        style.configure("Chooser.Treeview.Heading", background=BG_SOFT, foreground=NAVY, font=("Segoe UI", 10, "bold"))
        style.map("Chooser.Treeview", background=[("selected", "#EFF6FF")], foreground=[("selected", NAVY)])
    except Exception:  # noqa: BLE001 - оформление не критично
        pass


class _VedChooser:
    """Выпадающий список номенклатуры под ячейкой «Из ведомости» — таблицей."""

    def __init__(self, app, x: int, y: int, width: int, items, on_pick) -> None:
        self.app = app
        self.items = list(items)
        self.on_pick = on_pick
        self.closed = False
        self.win = tk.Toplevel(app)
        try:
            self.win.overrideredirect(True)
            self.win.configure(background=BORDER, bd=1, relief="solid")
        except Exception:
            pass
        try:
            self.win.configure(background=CARD)
        except Exception:
            pass
        try:
            screen_h = app.winfo_screenheight() if hasattr(app, "winfo_screenheight") else 800
            max_h = max(400, screen_h - y - 40) if screen_h else 600
            height = min(max_h, max(360, 28 * len(self.items) + 64))
            # если не вмещается вниз — переворачиваем вверх
            if screen_h and y + height > screen_h - 10:
                y = max(10, y - height - 30)
        except Exception:
            height = min(600, 28 * len(self.items) + 64)
        try:
            popup_w = max(int(width), 640)
            popup_w = max(popup_w, 760)
            self.win.geometry(f"{popup_w}x{int(height)}+{int(x)}+{int(y)}")
        except Exception:
            pass

        # поиск убран по ТЗ — выпадающее без окна поиска
        self.search_var = None

        # Таблица в 4 столбика: штрихкод | артикул | % | наименование — без слов "штрихкод"/"арт."
        self._is_tree = False
        self._tree_to_index = {}
        try:
            cols = ("barcode", "article", "percent", "name")
            self.tree = ttk.Treeview(self.win, columns=cols, show="headings", height=12, style="Chooser.Treeview")
            self.tree.heading("barcode", text="Штрихкод", anchor="center")
            self.tree.heading("article", text="Артикул", anchor="center")
            self.tree.heading("percent", text="%", anchor="center")
            self.tree.heading("name", text="Наименование", anchor="w")
            self.tree.column("barcode", width=130, minwidth=110, anchor="center", stretch=False)
            self.tree.column("article", width=110, minwidth=90, anchor="center", stretch=False)
            self.tree.column("percent", width=60, minwidth=50, anchor="center", stretch=False)
            self.tree.column("name", width=400, minwidth=200, anchor="w", stretch=True)
            # скролл
            vsb = ttk.Scrollbar(self.win, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=vsb.set)
            self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(2, 4))
            vsb.pack(side="right", fill="y", pady=(2, 4), padx=(0, 8))
            # заполнение — CHOICE_CUSTOM убран по ТЗ (больше нет «Своё название»)
            for label, val in self.items:
                if val is CHOICE_CUSTOM:
                    continue
                if val is None:
                    vals = ("", "", "—", label)
                else:
                    # row is VedRow
                    try:
                        vals = (getattr(val, "barcode", "") or "", getattr(val, "article", "") or "", label.split(" · ")[0] if " · " in label else "", getattr(val, "nomenclature", "") or label)
                        # label already contains percent as first part when we built in gui.py: percent is first after split? Actually gui builds as barcode·article·percent·name, but we have old label with percent first. For tree we need percent separately.
                        # For backward compat, extract percent from label if possible
                        # label format from gui 1.9.25: "штрихкод X · арт. Y · 79% · Name" — percent is third part
                        # So find part containing %
                        pct = ""
                        for part in label.split(" · "):
                            if "%" in part or part.strip() == "—":
                                pct = part
                                break
                        if not pct and " · " in label:
                            pct = label.split(" · ")[0]
                        # name is last part
                        name = label.split(" · ")[-1] if " · " in label else label
                        # override with actual row data for name
                        name = getattr(val, "nomenclature", name)
                        vals = (getattr(val, "barcode", "") or "", getattr(val, "article", "") or "", pct, name)
                    except Exception:
                        vals = ("", "", "", label)
                iid = self.tree.insert("", "end", values=vals)
                self._tree_to_index[iid] = len(self._tree_to_index)
            # alias для тестов (box)
            self.box = self.tree
            self._is_tree = True
            # для совместимости с тестами: curselection должен вернуть индекс
            # добавим методы
            orig_curselection = getattr(self.tree, "curselection", None)
            def _curselection():
                sel = self.tree.selection()
                if not sel:
                    return ()
                # map tree iid to index in self.items
                idx = self._tree_to_index.get(sel[0])
                return (idx,) if idx is not None else ()
            self.tree.curselection = _curselection  # type: ignore
            def _selection_set(idx):
                # idx is row index in current filtered items
                try:
                    children = self.tree.get_children()
                    if 0 <= idx < len(children):
                        iid = children[idx]
                        self.tree.selection_set(iid)
                        self.tree.see(iid)
                        self.tree.focus(iid)
                except Exception:
                    pass
            self.tree.selection_set = _selection_set  # type: ignore
            def _delete(first, last=None):
                # for search: delete all
                try:
                    for iid in self.tree.get_children():
                        self.tree.delete(iid)
                except Exception:
                    pass
                self._tree_to_index.clear()
            self.tree.delete = _delete  # type: ignore
            def _insert(index, element):
                # not used for tree but for compatibility
                try:
                    vals = ("", "", "", str(element))
                    iid = self.tree.insert("", "end", values=vals)
                    self._tree_to_index[iid] = len(self._tree_to_index)
                except Exception:
                    pass
            # keep original insert for search fallback
            self._orig_tree_insert = self.tree.insert
            # select first
            if self.tree.get_children():
                first = self.tree.get_children()[0]
                self.tree.selection_set(first)
                self.tree.focus(first)
            self.tree.bind("<ButtonRelease-1>", self._on_click)
            self.tree.bind("<Double-1>", self._on_activate)
            self.tree.bind("<Return>", self._on_activate)
            self.tree.bind("<Escape>", lambda _e: self.close())
        except Exception as e:
            # fallback на Listbox если Treeview не доступен (fake_tk)
            try:
                self.box = tk.Listbox(
                    self.win,
                    activestyle="none",
                    borderwidth=0,
                    relief="flat",
                    highlightthickness=0,
                    exportselection=False,
                    font=("Segoe UI", 10),
                    selectbackground="#EFF6FF",
                    selectforeground=NAVY,
                    background=CARD,
                    foreground=TEXT,
                )
                for label, _value in self.items:
                    self.box.insert("end", label)
                self.box.pack(fill="both", expand=True, padx=8, pady=(2, 4))
                if self.items:
                    self.box.selection_set(0)
                self._is_tree = False
            except Exception:
                self.box = None  # type: ignore
                self._is_tree = False

        hint = "  \u25BC  «не выбирать» → товар будет «Новинка»   ·   Esc — закрыть"
        try:
            ttk.Label(self.win, text=hint, style="Chooser.TLabel").pack(fill="x", padx=8, pady=(4, 8))
        except Exception:
            pass

        self._all_items = list(self.items)
        if not self._is_tree:
            try:
                self.box.bind("<ButtonRelease-1>", self._on_click)
                self.box.bind("<Return>", self._on_activate)
                self.box.bind("<Double-1>", self._on_activate)
                self.box.bind("<Escape>", lambda _e: self.close())
            except Exception:
                pass
        try:
            self.win.bind("<FocusOut>", lambda _e: self.close())
            self.win.bind("<Escape>", lambda _e: self.close())
            if self._is_tree:
                self.tree.focus_set()
            else:
                self.box.focus_set()
        except Exception:
            pass

    def _on_click(self, _event=None):
        self._pick_current()
        return "break"

    def _on_search(self, _event=None):
        # поиск убран — заглушка для совместимости
        return
        q = self.search_var.get().strip().lower()
        if self._is_tree:
            # фильтр Treeview
            try:
                for iid in self.tree.get_children():
                    self.tree.delete(iid)
                self._tree_to_index.clear()
            except Exception:
                pass
            self.items = []
            for label, val in self._all_items:
                if not q or q == "поиск…" or q in label.lower():
                    # вставить с теми же 4 колонками
                    if val is None:
                        vals = ("", "", "—", label)
                    elif val is CHOICE_CUSTOM:
                        vals = ("", "", "", label)
                    else:
                        # извлечь percent из label
                        pct = ""
                        for part in label.split(" · "):
                            if "%" in part or part.strip() == "—":
                                pct = part
                                break
                        if not pct:
                            pct = label.split(" · ")[0] if " · " in label else ""
                        name = getattr(val, "nomenclature", label.split(" · ")[-1] if " · " in label else label)
                        vals = (getattr(val, "barcode", "") or "", getattr(val, "article", "") or "", pct, name)
                    iid = self.tree.insert("", "end", values=vals)
                    self._tree_to_index[iid] = len(self.items)
                    self.items.append((label, val))
            if self.tree.get_children():
                first = self.tree.get_children()[0]
                self.tree.selection_set(first)
                self.tree.focus(first)
        else:
            self.box.delete(0, "end")
            self.items = []
            for label, val in self._all_items:
                if not q or q == "поиск…" or q in label.lower():
                    self.box.insert("end", label)
                    self.items.append((label, val))
            if self.items:
                self.box.selection_set(0)

    def _on_activate(self, _event=None):
        self._pick_current()
        return "break"

    def _pick_current(self) -> None:
        index = 0
        try:
            if self._is_tree:
                sel = self.tree.selection()
                if sel:
                    iid = sel[0]
                    # найти индекс в текущем отфильтрованном списке
                    children = list(self.tree.get_children())
                    try:
                        index = children.index(iid)
                    except ValueError:
                        index = 0
                else:
                    # fallback на первую строку
                    index = 0
            else:
                selected = self.box.curselection()
                if selected:
                    index = int(selected[0])
        except Exception:
            index = 0
        if index < 0 or index >= len(self.items):
            self.close()
            return
        value = self.items[index][1]
        self.close()
        if self.on_pick is not None:
            self.on_pick(value)

    def labels(self) -> List[str]:
        return [label for label, _value in self.items]

    def close(self) -> None:
        if self.closed:
            return
        self.closed = True
        for widget in (self.win,):
            try:
                widget.destroy()
            except Exception:
                pass
