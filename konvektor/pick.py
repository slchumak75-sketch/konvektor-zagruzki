"""Диалог «склад и поставщик» — спрашивается при выборе файла.

Заказчик просил: при загрузке файла программа спрашивает **склад**
(«Мир» или «Кореличи») и даёт выбрать **поставщика** из тех, что есть в
загруженной ведомости. Склад и поставщик задают папку результата
(`<папка>/<склад>/<поставщик>/<дата>.xlsx`) и то, среди чьих позиций искать
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
#: Пункт списка поставщиков: не выбирать, взять поставщика из документа.
DOC_SUPPLIER = "— как в документе —"

BG = "#F4F6FA"
CARD = "#FFFFFF"
TEXT = "#1B2733"
MUTED = "#5B6B7F"
NAVY = "#1F3A5F"
BORDER = "#D8E0EC"


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
            self.win.geometry("560x300")
            self.win.configure(background=BG)
            self.win.transient(master)
            screen_h = self.win.winfo_screenheight()
            if screen_h and screen_h > 200:
                self.win.geometry(f"560x{screen_h - 80}+100+20")
        except Exception:  # noqa: BLE001 - оформление не критично
            pass

        self.sklad_var = tk.StringVar(value=sklad or SKLADS[0])
        self.supplier_var = tk.StringVar(value=supplier or DOC_SUPPLIER)
        self._build()

    # ------------------------------------------------------------------ вёрстка
    def _build(self) -> None:
        pad = {"padx": 18}
        tk.Label(
            self.win,
            text="Куда положить результат",
            background=BG,
            foreground=NAVY,
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", pady=(14, 2), **pad)

        sklad_row = ttk.Frame(self.win, style="Pick.TFrame")
        sklad_row.pack(fill="x", pady=(2, 8), **pad)
        ttk.Label(sklad_row, text="Склад:", style="Pick.TLabel").pack(side="left")
        self.sklad_buttons = {}
        for name in SKLADS:
            button = ttk.Radiobutton(
                sklad_row, text=name, value=name, variable=self.sklad_var
            )
            button.pack(side="left", padx=(12, 0))
            self.sklad_buttons[name] = button

        ttk.Label(self.win, text="Поставщик:", style="Pick.TLabel").pack(anchor="w", **pad)
        buttons = ttk.Frame(self.win, style="Pick.TFrame")
        buttons.pack(side="bottom", fill="x", pady=(8, 12), **pad)
        self.ok_btn = ttk.Button(
            buttons, text="Продолжить", style="PickPrimary.TButton", command=self.on_ok
        )
        self.ok_btn.pack(side="right")
        self.cancel_btn = ttk.Button(
            buttons, text="Отмена", style="Pick.TButton", command=self.on_cancel
        )
        self.cancel_btn.pack(side="right", padx=(0, 8))

        self.hint_label = ttk.Label(
            self.win,
            text=self._hint(),
            style="PickHint.TLabel",
            wraplength=500,
            justify="left",
        )
        self.hint_label.pack(side="bottom", anchor="w", **pad)

        list_frame = ttk.Frame(self.win, style="Pick.TFrame")
        list_frame.pack(fill="both", expand=True, pady=(2, 6), **pad)

        canvas = tk.Canvas(list_frame, height=160, highlightthickness=1, highlightbackground="#C5D2E0", background=BG)
        vscroll = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, style="Pick.TFrame")
        inner.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        initial_set = {
            s.strip() for s in (self.supplier_var.get() or "").split(",") if s.strip()
        }
        if not initial_set or DOC_SUPPLIER in initial_set:
            initial_set = set()

        self.supplier_vars: dict[str, tk.BooleanVar] = {}
        self.supplier_checks: dict[str, ttk.Checkbutton] = {}
        for name in [DOC_SUPPLIER] + self.suppliers:
            checked = (name == DOC_SUPPLIER and not initial_set) or (name in initial_set)
            var = tk.BooleanVar(value=checked)
            cb = ttk.Checkbutton(
                inner,
                text=name,
                variable=var,
                command=lambda n=name: self._on_supplier_toggle(n),
            )
            cb.pack(anchor="w", padx=6, pady=2)
            self.supplier_vars[name] = var
            self.supplier_checks[name] = cb

    def _hint(self) -> str:
        if not self.suppliers:
            return (
                "Ведомость не загружена — поставщик берётся из документа. "
                "Загрузите ведомость, чтобы выбрать поставщика из неё."
            )
        return (
            f"В ведомости {len(self.suppliers)} поставщик(ов). Выберите своего — "
            "номенклатура будет подбираться только среди его позиций, а файл "
            "попадёт в папку «склад / поставщик»."
        )

    # ------------------------------------------------------------------ данные
    def _on_supplier_toggle(self, name: str) -> None:
        if name == DOC_SUPPLIER:
            if self.supplier_vars[DOC_SUPPLIER].get():
                for k, v in self.supplier_vars.items():
                    if k != DOC_SUPPLIER:
                        v.set(False)
        else:
            if self.supplier_vars[name].get():
                self.supplier_vars[DOC_SUPPLIER].set(False)

    @property
    def supplier_value(self) -> str:
        """Выбранный поставщик (пустая строка — как в документе)."""
        if self.supplier_vars.get(DOC_SUPPLIER) and self.supplier_vars[DOC_SUPPLIER].get():
            return ""
        checked = [
            name
            for name, var in self.supplier_vars.items()
            if name != DOC_SUPPLIER and var.get()
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
            popup_w = max(int(width), 620)
            self.win.geometry(f"{popup_w}x{int(height)}+{int(x)}+{int(y)}")
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
