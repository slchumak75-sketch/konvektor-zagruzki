"""Диалог «склад и поставщик» — спрашивается при выборе файла.

Заказчик просил: при загрузке файла программа спрашивает **склад**
(«Мир» или «Кореличи») и даёт выбрать **поставщика** из тех, что есть в
загруженной ведомости. Склад и поставщик задают папку результата
(`<папка>/<склад>/<поставщик>/<дата>.xlsx`) и то, среди чьих позиций искать
номенклатуру в ведомости.
"""

from __future__ import annotations

from typing import List

try:  # pragma: no cover - зависит от сборки Python
    import tkinter as tk
    from tkinter import ttk, messagebox

    TK_AVAILABLE = True
except Exception:  # noqa: BLE001 - без tkinter диалог не нужен
    tk = None  # type: ignore[assignment]
    ttk = None  # type: ignore[assignment]
    TK_AVAILABLE = False

#: Склады заказчика.
SKLADS = ("Мир", "Кореличи")
#: Пункт списка поставщиков: не выбирать, взять поставщика из документа.
DOC_SUPPLIER = "— как в документе —"

BG = "#F4F6FA"
CARD = "#FFFFFF"
TEXT = "#1B2733"
MUTED = "#5B6B7F"
NAVY = "#1F3A5F"
BORDER = "#D8E0EC"


class SourceDialog:
    """Выбор склада и нескольких поставщиков только из этого склада."""

    def __init__(self, master, suppliers=(), sklad="", supplier=(), on_done=None, require_selection=False):
        self.master, self.on_done = master, on_done
        self.require_selection = require_selection
        self._supplier_source = suppliers
        self._initial = [supplier] if isinstance(supplier, str) and supplier else list(supplier or ())
        self.suppliers = []
        self.win = tk.Toplevel(master)
        self.win.title("Склад и поставщики")
        self.win.geometry("600x480")
        self.win.configure(background=BG)
        self.win.transient(master)
        self.sklad_var = tk.StringVar(value=sklad or SKLADS[0])
        self._build()
        self._change_sklad(initial=True)

    def _build(self):
        pad = {"padx": 18}
        ttk.Label(self.win, text="Склад:", style="Pick.TLabel").pack(anchor="w", pady=(14, 2), **pad)
        row = ttk.Frame(self.win, style="Pick.TFrame")
        row.pack(fill="x", **pad)
        self.sklad_buttons = {}
        for name in SKLADS:
            button = ttk.Radiobutton(row, text=name, value=name, variable=self.sklad_var,
                                     command=self._change_sklad)
            button.pack(side="left", padx=(0, 16))
            self.sklad_buttons[name] = button
        ttk.Label(self.win, text="Поставщики выбранного склада (можно выбрать несколько):",
                  style="Pick.TLabel").pack(anchor="w", pady=(10, 2), **pad)
        box = ttk.Frame(self.win, style="Pick.TFrame")
        box.pack(fill="both", expand=True, **pad)
        self.supplier_box = tk.Listbox(box, selectmode="multiple", exportselection=False,
                                       height=10, font=("Segoe UI", 10))
        scrollbar = ttk.Scrollbar(box, orient="vertical", command=self.supplier_box.yview)
        self.supplier_box.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.supplier_box.pack(side="left", fill="both", expand=True)
        buttons = ttk.Frame(self.win, style="Pick.TFrame")
        buttons.pack(fill="x", pady=(8, 8), **pad)
        ttk.Button(buttons, text="Выбрать всех", command=self.select_all).pack(side="left")
        ttk.Button(buttons, text="Снять выбор", command=self.clear_selection).pack(side="left", padx=8)
        ttk.Button(buttons, text="Продолжить", style="PickPrimary.TButton", command=self.on_ok).pack(side="right")
        ttk.Button(buttons, text="Отмена", command=self.on_cancel).pack(side="right", padx=8)
        self.hint_label = ttk.Label(self.win, text="Выберите поставщиков из ведомости. Если ведомость не загружена — как в документе.",
                                    style="PickHint.TLabel", wraplength=550)
        self.hint_label.pack(anchor="w", pady=(0, 10), **pad)

    def _change_sklad(self, initial=False):
        source = self._supplier_source
        values = source(self.sklad_value) if callable(source) else source
        self.suppliers = sorted(set(str(v).strip() for v in values if str(v).strip()), key=str.casefold)
        self.supplier_box.delete(0, "end")
        for i, name in enumerate(self.suppliers):
            self.supplier_box.insert("end", name)
            if initial and name in self._initial:
                self.supplier_box.selection_set(i)

    def select_all(self):
        self.supplier_box.selection_set(0, "end")

    def clear_selection(self):
        self.supplier_box.selection_clear(0, "end")

    @property
    def supplier_value(self):
        return [self.suppliers[int(i)] for i in self.supplier_box.curselection()]

    @property
    def sklad_value(self):
        return str(self.sklad_var.get() or "").strip()

    def on_ok(self):
        sklad, suppliers = self.sklad_value, self.supplier_value
        if self.require_selection and not suppliers:
            messagebox.showwarning("Склад и поставщики", "Выберите хотя бы одного поставщика этого склада. Если список пуст, проверьте колонку «Склад» ведомости.", parent=self.win)
            return
        self.close()
        if self.on_done is not None:
            self.on_done(sklad, suppliers)

    def on_cancel(self):
        self.close()

    def close(self):
        self.win.destroy()


#: Пункт выпадающего списка номенклатуры: вписать своё название.
CHOICE_CUSTOM = object()


def install_styles(style) -> None:
    """Завести в ttk-тему стили диалога и выпадающего списка."""
    try:
        style.configure("Pick.TFrame", background=BG)
        style.configure("Pick.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 10))
        style.configure(
            "PickHint.TLabel",
            background=BG,
            foreground=MUTED,
            font=("Segoe UI", 9),
        )
        style.configure("Pick.TButton", font=("Segoe UI", 10), padding=(14, 5))
        style.configure("PickPrimary.TButton", font=("Segoe UI", 10), padding=(18, 5))
        style.configure("Chooser.Toplevel", background=CARD)
        style.configure("Chooser.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("Chooser.TButton", font=("Segoe UI", 9), padding=(6, 2))
    except Exception:  # noqa: BLE001 - оформление не критично
        pass


class _VedChooser:
    """Выпадающий список номенклатуры под ячейкой «Из ведомости».

    Обычное безрамочное окно поверх таблицы: щелчок по строке выбирает
    номенклатуру, Esc или щелчок мимо — закрывает. Окно не «висит» на месте:
    таблица закрывает его при прокрутке и при любом новом действии.
    """

    def __init__(self, app, x: int, y: int, width: int, items, on_pick) -> None:
        self.app = app
        self.items = list(items)
        self.on_pick = on_pick
        self.closed = False
        self.win = tk.Toplevel(app)
        try:
            self.win.overrideredirect(True)
        except Exception:  # noqa: BLE001 - не критично
            pass
        try:
            self.win.configure(background=CARD)
        except Exception:  # noqa: BLE001 - не критично
            pass
        height = min(320, 26 * len(self.items) + 34)
        try:
            self.win.geometry(f"{int(width)}x{int(height)}+{int(x)}+{int(y)}")
        except Exception:  # noqa: BLE001 - не критично
            pass

        self.box = tk.Listbox(
            self.win,
            activestyle="none",
            borderwidth=1,
            relief="solid",
            highlightthickness=0,
            exportselection=False,
            font=("Segoe UI", 10),
        )
        for label, _value in self.items:
            self.box.insert("end", label)
        self.box.pack(fill="both", expand=True, padx=1, pady=(1, 0))
        if self.items:
            self.box.selection_set(0)
        hint = "«не выбирать» → товар будет «Новинка»   ·   Esc — закрыть"
        try:
            ttk.Label(self.win, text=hint, style="Chooser.TLabel").pack(
                fill="x", padx=6, pady=(2, 4)
            )
        except Exception:  # noqa: BLE001 - подсказка не критична
            pass

        self.box.bind("<ButtonRelease-1>", self._on_click)
        self.box.bind("<Return>", self._on_activate)
        self.box.bind("<Double-1>", self._on_activate)
        self.box.bind("<Escape>", lambda _e: self.close())
        try:
            self.win.bind("<FocusOut>", lambda _e: self.close())
            self.win.bind("<Escape>", lambda _e: self.close())
            self.box.focus_set()
        except Exception:  # noqa: BLE001 - не критично
            pass

    # ------------------------------------------------------------------ события
    def _on_click(self, _event=None):
        self._pick_current()
        return "break"

    def _on_activate(self, _event=None):
        self._pick_current()
        return "break"

    def _pick_current(self) -> None:
        index = 0
        try:
            selected = self.box.curselection()
            if selected:
                index = int(selected[0])
        except Exception:  # noqa: BLE001 - не критично
            index = 0
        if index < 0 or index >= len(self.items):
            self.close()
            return
        value = self.items[index][1]
        self.close()
        if self.on_pick is not None:
            self.on_pick(value)

    # ------------------------------------------------------------------ данные
    def labels(self) -> List[str]:
        return [label for label, _value in self.items]

    def close(self) -> None:
        if self.closed:
            return
        self.closed = True
        for widget in (self.win,):
            try:
                widget.destroy()
            except Exception:  # noqa: BLE001 - уже уничтожен
                pass
