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
            self.win.geometry("620x420")
            self.win.configure(background=BG)
            self.win.transient(master)
            # модальное затемнение
            try:
                self.win.attributes("-topmost", True)
                self.win.after(200, lambda: self.win.attributes("-topmost", False))
            except Exception:
                pass
            screen_h = self.win.winfo_screenheight()
            if screen_h and screen_h > 200:
                full_h = max(520, screen_h - 60)
                self.win.geometry(f"640x{full_h}+100+20")
        except Exception:  # noqa: BLE001 - оформление не критично
            pass

        self.sklad_var = tk.StringVar(value=sklad or SKLADS[0])
        # поставщики сортируются по алфавиту (требование 1.9.18)
        self.suppliers = sorted(self.suppliers, key=str.casefold)
        self.supplier_var = tk.StringVar(value=supplier or "")
        self._build()

    # ------------------------------------------------------------------ вёрстка
    def _build(self) -> None:
        pad = {"padx": 20}
        # заголовок-карточка
        header = tk.Frame(self.win, background=NAVY, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="  \U0001F4E6  Склад и поставщик", background=NAVY, foreground="#FFFFFF", font=("Segoe UI", 12, "bold"), anchor="w").pack(fill="both", expand=True, padx=14, pady=10)

        # карточка склада — наглядный выбор (крупные кнопки с иконками)
        card = tk.Frame(self.win, background=CARD, highlightbackground=BORDER, highlightthickness=1, bd=0)
        card.pack(fill="x", padx=16, pady=(14, 10))
        tk.Label(card, text="Выберите склад — куда положить результат", background=CARD, foreground=NAVY, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14, pady=(10, 6))
        sklad_row = tk.Frame(card, background=CARD)
        sklad_row.pack(fill="x", padx=12, pady=(6, 12))
        self.sklad_buttons = {}
        for name in SKLADS:
            icon = "\U0001F3E2" if name == "Мир" else "\U0001F3ED"
            is_selected = self.sklad_var.get() == name
            # Пытаемся сделать настоящую tk.Button с подсветкой ACCENT (красиво в реальном окне).
            # В тестах (fake_tk) tk.Button отсутствует — откатываемся к ttk.
            btn = None
            try:
                ButtonCls = getattr(tk, "Button", None)
                if ButtonCls is not None:
                    btn = ButtonCls(
                        sklad_row,
                        text=f"{icon}  {name}",
                        font=("Segoe UI", 11, "bold" if is_selected else "normal"),
                        background=ACCENT if is_selected else CARD,
                        foreground="#FFFFFF" if is_selected else NAVY,
                        activebackground=ACCENT_HOVER if is_selected else BG_SOFT,
                        relief="flat" if is_selected else "solid",
                        bd=1,
                        highlightbackground=BORDER,
                        padx=20,
                        pady=10,
                        cursor="hand2",
                        command=lambda n=name: self._select_sklad(n),
                    )
                else:
                    raise AttributeError("no tk.Button")
            except Exception:
                # fallback для fake_tk / окружений без tk.Button
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
                btn.pack(side="left", padx=8, expand=True, fill="x", ipady=6)
            except Exception:
                pass
            self.sklad_buttons[name] = btn
        # подсказка под складом
        tk.Label(card, text="Файл попадёт в папку «склад / поставщик / дата.xlsx»", background=CARD, foreground=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=14, pady=(0, 10))

        ttk.Label(self.win, text="  Поставщик (по алфавиту):", background=BG, foreground=TEXT, font=("Segoe UI", 10, "bold")).pack(anchor="w", **pad)
        buttons = ttk.Frame(self.win, style="Pick.TFrame")
        buttons.pack(side="bottom", fill="x", pady=(12, 16), **pad)
        self.ok_btn = ttk.Button(buttons, text="✓  Продолжить", style="PickPrimary.TButton", command=self.on_ok)
        self.ok_btn.pack(side="right")
        self.cancel_btn = ttk.Button(buttons, text="Отмена", style="Pick.TButton", command=self.on_cancel)
        self.cancel_btn.pack(side="right", padx=(0, 10))

        self.hint_label = ttk.Label(self.win, text=self._hint(), style="PickHint.TLabel", wraplength=540, justify="left")
        self.hint_label.pack(side="bottom", anchor="w", **pad)

        # список поставщиков — на всю высоту экрана (требование 1.9.18)
        list_frame = tk.Frame(self.win, background=CARD, highlightbackground=BORDER, highlightthickness=1, bd=0)
        list_frame.pack(fill="both", expand=True, padx=16, pady=(6, 8))

        # высота по экрану
        try:
            screen_h = self.win.winfo_screenheight()
            canvas_h = max(300, screen_h - 260) if screen_h else 400
        except Exception:
            canvas_h = 400
        canvas = tk.Canvas(list_frame, height=canvas_h, highlightthickness=0, background=CARD)
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

        # без пункта «как в документе» — только реальные поставщики по алфавиту
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
            )
            cb.pack(anchor="w", padx=8, pady=3)
            self.supplier_vars[name] = var
            self.supplier_checks[name] = cb
        # если ведомость не загружена — подсказка, иначе счётчик
        if not self.suppliers:
            ttk.Label(inner, text="Загрузите ведомость — здесь появятся поставщики", style="PickHint.TLabel").pack(anchor="w", padx=8, pady=6)

    def _select_sklad(self, name: str) -> None:
        """Наглядный выбор склада — подсветка активной кнопки."""
        self.sklad_var.set(name)
        for sklad_name, btn in self.sklad_buttons.items():
            is_sel = sklad_name == name
            try:
                btn.configure(
                    background=ACCENT if is_sel else CARD,
                    foreground="#FFFFFF" if is_sel else NAVY,
                    font=("Segoe UI", 11, "bold" if is_sel else "normal"),
                    relief="flat" if is_sel else "solid",
                )
            except Exception:
                pass
        # подсказку тоже можно освежить
        try:
            self.hint_label.configure(text=self._hint())
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
        # оставлено для совместимости, но «как в документе» убран
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
        style.configure("Sklad.TRadiobutton", font=("Segoe UI", 11), padding=(10, 8))
        style.configure("Chooser.Toplevel", background=CARD)
        style.configure("Chooser.TLabel", background="#F8FAFC", foreground=MUTED, font=("Segoe UI", 8))
        style.configure("Chooser.TButton", font=("Segoe UI", 9), padding=(6, 3))
        style.configure("Chooser.Search.TEntry", fieldbackground="#FFFFFF", padding=(8, 6))
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
            # тень/граница для современного вида
            self.win.configure(background=BORDER, bd=1, relief="solid")
        except Exception:  # noqa: BLE001 - не критично
            pass
        try:
            self.win.configure(background=CARD)
        except Exception:  # noqa: BLE001 - не критично
            pass
        # на всю высоту экрана (требование 1.9.18)
        try:
            screen_h = app.winfo_screenheight() if hasattr(app, "winfo_screenheight") else 800
            max_h = max(400, screen_h - y - 40) if screen_h else 600
            height = min(max_h, max(360, 28 * len(self.items) + 64))
        except Exception:
            height = min(600, 28 * len(self.items) + 64)
        try:
            popup_w = max(int(width), 640)
            self.win.geometry(f"{popup_w}x{int(height)}+{int(x)}+{int(y)}")
        except Exception:  # noqa: BLE001 - не критично
            pass

        # поиск внутри выпадающего списка — современный дизайн
        try:
            search_row = tk.Frame(self.win, background="#F8FAFC", highlightbackground=BORDER, highlightthickness=1)
            search_row.pack(fill="x", padx=8, pady=(8, 6))
            tk.Label(search_row, text="\U0001F50D", background="#F8FAFC", foreground=MUTED, font=("Segoe UI", 9)).pack(side="left", padx=(8, 4))
            self.search_var = tk.StringVar()
            self.search_entry = ttk.Entry(search_row, textvariable=self.search_var, style="Chooser.Search.TEntry", font=("Segoe UI", 9))
            self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 6), pady=4)
            self.search_entry.insert(0, "Поиск…")
            self.search_entry.bind("<FocusIn>", lambda e: self.search_var.set("") if self.search_var.get()=="Поиск…" else None)
            self.search_entry.bind("<KeyRelease>", self._on_search)
        except Exception:
            self.search_var = None

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
        hint = "  \u25BC  «не выбирать» → товар будет «Новинка»   ·   Esc — закрыть   ·   начните печатать для поиска"
        try:
            ttk.Label(self.win, text=hint, style="Chooser.TLabel").pack(fill="x", padx=8, pady=(4, 8))
        except Exception:  # noqa: BLE001 - подсказка не критична
            pass

        self._all_items = list(self.items)
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

    def _on_search(self, _event=None):
        if not hasattr(self, "search_var") or self.search_var is None:
            return
        q = self.search_var.get().strip().lower()
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
