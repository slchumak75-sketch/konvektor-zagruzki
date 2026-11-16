"""Окно ручной сверки с ведомостью.

Показывается по кнопке «Сверить с ведомостью…» после конвертации. Слева —
позиции документа с тем, что подобралось автоматически, справа — позиции
ведомости: их можно искать по словам и выбирать вручную. Кнопка «Не выбирать»
оставляет в документе исходное наименование.

Логики подбора здесь нет: всё в konvektor/vedomost.py, окно только показывает
и записывает решения пользователя в matcher.overrides.
"""

from __future__ import annotations

from typing import Callable, Dict, List, Optional, Tuple

from . import APP_NAME
from .vedomost import Matcher, VedRow, similarity

try:  # pragma: no cover - зависит от сборки Python
    import tkinter as tk
    from tkinter import ttk

    TK_AVAILABLE = True
except Exception:  # noqa: BLE001 - без tkinter окно не нужно
    tk = None  # type: ignore[assignment]
    ttk = None  # type: ignore[assignment]
    TK_AVAILABLE = False

#: Оформление — те же цвета, что в главном окне.
BG = "#F4F6FA"
CARD = "#FFFFFF"
BORDER = "#D8E0EC"
TEXT = "#1B2733"
MUTED = "#5B6B7F"
NAVY = "#1F3A5F"

FILTER_ALL = "Все позиции"
FILTER_REVIEW = "Только требующие проверки"
FILTER_NONE = "Только не подобранные"
FILTERS = (FILTER_ALL, FILTER_REVIEW, FILTER_NONE)


class ReviewDialog:
    """Окно «Сверка с ведомостью»: ручной выбор номенклатуры."""

    def __init__(
        self,
        master,
        matcher: Matcher,
        on_done: Optional[Callable[[Dict[Tuple[str, str, int], Optional[VedRow]]], None]] = None,
    ) -> None:
        self.master = master
        self.matcher = matcher
        self.ved = matcher.vedomost
        self.overrides: Dict[Tuple[str, str, int], Optional[VedRow]] = dict(matcher.overrides)
        self.on_done = on_done
        self.changed = False

        self._entry_by_row: Dict[str, object] = {}
        self._candidate_by_row: Dict[str, VedRow] = {}

        self.win = tk.Toplevel(master)
        try:
            self.win.title(f"{APP_NAME} — сверка с ведомостью")
            self.win.geometry("1180x680")
            self.win.minsize(900, 520)
            self.win.configure(background=BG)
        except Exception:  # noqa: BLE001 - оформление не критично
            pass
        self.search_var = tk.StringVar(value="")
        self.scope_var = tk.BooleanVar(value=bool(matcher.summary.get("supplier_found")))
        self.filter_var = tk.StringVar(value=FILTER_ALL)
        self.status_var = tk.StringVar(value="")

        self._build()
        self.refresh()

    # ------------------------------------------------------------------ вёрстка
    def _build(self) -> None:
        head = tk.Frame(self.win, background=NAVY, height=56)
        head.pack(fill="x")
        head.pack_propagate(False)
        tk.Label(
            head,
            text="Сверка с ведомостью: выберите номенклатуру вручную",
            background=NAVY,
            foreground="#FFFFFF",
            font=("Segoe UI", 12, "bold"),
        ).pack(side="left", padx=16, pady=14)
        self.stats_label = tk.Label(
            head,
            text="",
            background=NAVY,
            foreground="#BFD0E6",
            font=("Segoe UI", 9),
        )
        self.stats_label.pack(side="right", padx=16)

        body = tk.Frame(self.win, background=BG)
        body.pack(fill="both", expand=True, padx=12, pady=10)

        # --- слева: позиции документа
        left = ttk.LabelFrame(body, text="Позиции документа", style="Review.TLabelframe")
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))

        filter_row = ttk.Frame(left, style="Review.TFrame")
        filter_row.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Label(filter_row, text="Показать:", style="Review.TLabel").pack(side="left")
        self.filter_box = ttk.Combobox(
            filter_row,
            textvariable=self.filter_var,
            values=list(FILTERS),
            state="readonly",
            width=32,
        )
        self.filter_box.pack(side="left", padx=6)
        self.filter_box.bind("<<ComboboxSelected>>", lambda _e: self.refresh())

        item_box = ttk.Frame(left, style="Review.TFrame")
        item_box.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.item_tree = ttk.Treeview(
            item_box,
            columns=("doc", "ved", "how", "score"),
            show="headings",
            selectmode="browse",
        )
        for column, title, width in (
            ("doc", "В документе", 300),
            ("ved", "Из ведомости", 300),
            ("how", "Как", 90),
            ("score", "Совпадение", 90),
        ):
            self.item_tree.heading(column, text=title)
            self.item_tree.column(
                column, width=width, anchor="w" if column in ("doc", "ved") else "center"
            )
        item_scroll = ttk.Scrollbar(item_box, orient="vertical", command=self.item_tree.yview)
        self.item_tree.configure(yscrollcommand=item_scroll.set)
        item_scroll.pack(side="right", fill="y")
        self.item_tree.pack(side="left", fill="both", expand=True)
        self.item_tree.bind("<<TreeviewSelect>>", lambda _e: self.on_select_item())

        # --- справа: ведомость
        right = ttk.LabelFrame(body, text="Ведомость", style="Review.TLabelframe")
        right.pack(side="left", fill="both", expand=True)

        search_row = ttk.Frame(right, style="Review.TFrame")
        search_row.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Entry(search_row, textvariable=self.search_var, width=28).pack(
            side="left", fill="x", expand=True
        )
        self.search_btn = ttk.Button(
            search_row, text="Найти", style="Review.TButton", command=self.on_search
        )
        self.search_btn.pack(side="left", padx=(6, 0))
        self.scope_check = ttk.Checkbutton(
            right,
            text="искать только у поставщика документа",
            variable=self.scope_var,
            command=self.on_search,
        )
        self.scope_check.pack(anchor="w", padx=8)

        cand_box = ttk.Frame(right, style="Review.TFrame")
        cand_box.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        self.cand_tree = ttk.Treeview(
            cand_box,
            columns=("name", "supplier", "article", "unit"),
            show="headings",
            selectmode="browse",
        )
        for column, title, width in (
            ("name", "Наименование в ведомости", 330),
            ("supplier", "Поставщик", 190),
            ("article", "Артикул", 100),
            ("unit", "Ед.", 50),
        ):
            self.cand_tree.heading(column, text=title)
            self.cand_tree.column(
                column, width=width, anchor="w" if column != "unit" else "center"
            )
        cand_scroll = ttk.Scrollbar(cand_box, orient="vertical", command=self.cand_tree.yview)
        self.cand_tree.configure(yscrollcommand=cand_scroll.set)
        cand_scroll.pack(side="right", fill="y")
        self.cand_tree.pack(side="left", fill="both", expand=True)
        self.cand_tree.bind("<Double-1>", lambda _e: self.on_apply())

        # --- кнопки
        buttons = ttk.Frame(self.win, style="Review.TFrame")
        buttons.pack(fill="x", padx=12, pady=(0, 8))
        self.apply_btn = ttk.Button(
            buttons, text="Выбрать эту номенклатуру", style="ReviewPrimary.TButton",
            command=self.on_apply,
        )
        self.apply_btn.pack(side="left")
        self.skip_btn = ttk.Button(
            buttons, text="Не выбирать", style="Review.TButton", command=self.on_skip
        )
        self.skip_btn.pack(side="left", padx=(8, 0))
        self.reset_btn = ttk.Button(
            buttons, text="Сбросить ручной выбор", style="Review.TButton",
            command=self.on_reset,
        )
        self.reset_btn.pack(side="left", padx=(8, 0))
        self.done_btn = ttk.Button(
            buttons, text="Готово", style="ReviewPrimary.TButton", command=self.on_close
        )
        self.done_btn.pack(side="right")
        ttk.Label(buttons, textvariable=self.status_var, style="ReviewHint.TLabel").pack(
            side="left", padx=12
        )

    # ------------------------------------------------------------------ данные
    @property
    def scope_supplier(self) -> str:
        """Ключ поставщика для поиска (пустая строка — вся ведомость)."""
        if not self.scope_var.get():
            return ""
        return str(self.matcher.summary.get("supplier_key") or "")

    def entries(self) -> List[object]:
        """Позиции с учётом выбранного фильтра."""
        entries = list(self.matcher.entries)
        mode = self.filter_var.get()
        if mode == FILTER_REVIEW:
            return [e for e in entries if e.review]
        if mode == FILTER_NONE:
            return [e for e in entries if not e.matched and not e.skipped]
        return entries

    def refresh(self) -> None:
        """Перерисовать список позиций (после ручного выбора)."""
        for row in list(self.item_tree.get_children()):
            self.item_tree.delete(row)
        self._entry_by_row.clear()
        for entry in self.entries():
            row = self.item_tree.insert("", "end", values=self._row_values(entry))
            self._entry_by_row[row] = entry
        summary = self.matcher.summary
        self.stats_label.configure(
            text=(
                f"подобрано {summary['matched']} из {summary['items']} · "
                f"проверить {summary['review']} · не подобрано {summary['unmatched']}"
            )
        )
        self.status_var.set(
            "Выберите позицию слева и номенклатуру справа — или «Не выбирать»."
        )

    def _row_values(self, entry) -> Tuple[str, str, str, str]:
        if entry.key in self.overrides:
            chosen = self.overrides[entry.key]
            if chosen is None:
                return (entry.doc_name, entry.doc_name, "не выбирать", "—")
            return (entry.doc_name, chosen.nomenclature, "вручную", "вручную")
        return (entry.doc_name, entry.name, entry.how_title, entry.score_title)

    def selected_entry(self):
        selection = self.item_tree.selection()
        if not selection:
            return None
        return self._entry_by_row.get(selection[0])

    def selected_candidate(self) -> Optional[VedRow]:
        selection = self.cand_tree.selection()
        if not selection:
            return None
        return self._candidate_by_row.get(selection[0])

    # ---------------------------------------------------------------- события
    def on_select_item(self) -> None:
        """Показать для выбранной позиции похожие номенклатуры ведомости."""
        entry = self.selected_entry()
        if entry is None:
            return
        self.search_var.set(entry.doc_name)
        self._fill_candidates(self.ved.candidates(entry.doc_name, self.scope_supplier, limit=60))

    def on_search(self) -> None:
        query = self.search_var.get()
        rows = self.ved.search(query, self.scope_supplier, limit=300)
        scored = [(similarity(query, row.nomenclature), row) for row in rows] if query else []
        if scored:
            self._fill_candidates(scored)
        else:
            self._fill_candidates([(0.0, row) for row in rows])

    def _fill_candidates(self, pairs) -> None:
        for row in list(self.cand_tree.get_children()):
            self.cand_tree.delete(row)
        self._candidate_by_row.clear()
        for _score, row in pairs:
            key = self.cand_tree.insert(
                "", "end",
                values=(row.nomenclature, row.supplier, row.article, row.unit),
            )
            self._candidate_by_row[key] = row
        if not pairs:
            self.status_var.set("Ничего не найдено — измените запрос или снимите галочку поставщика.")

    def on_apply(self) -> None:
        entry = self.selected_entry()
        candidate = self.selected_candidate()
        if entry is None:
            self.status_var.set("Сначала выберите позицию слева.")
            return
        if candidate is None:
            self.status_var.set("Сначала выберите номенклатуру справа.")
            return
        self.overrides[entry.key] = candidate
        self.changed = True
        self.status_var.set(f"Выбрано: {candidate.nomenclature}")
        self.refresh()
        self._reselect(entry)

    def on_skip(self) -> None:
        entry = self.selected_entry()
        if entry is None:
            self.status_var.set("Сначала выберите позицию слева.")
            return
        self.overrides[entry.key] = None
        self.changed = True
        self.status_var.set("Наименование документа остаётся как есть.")
        self.refresh()
        self._reselect(entry)

    def on_reset(self) -> None:
        entry = self.selected_entry()
        if entry is None:
            self.status_var.set("Сначала выберите позицию слева.")
            return
        if entry.key in self.overrides:
            self.overrides.pop(entry.key)
            self.changed = True
        self.refresh()
        self._reselect(entry)

    def _reselect(self, entry) -> None:
        for row, kept in self._entry_by_row.items():
            if kept is entry:
                try:
                    self.item_tree.selection_set(row)
                except Exception:  # noqa: BLE001 - не критично
                    pass
                break

    def on_close(self) -> None:
        if self.changed and self.on_done is not None:
            self.on_done(dict(self.overrides))
        try:
            self.win.destroy()
        except Exception:  # noqa: BLE001 - окно могло уже закрыться
            pass
