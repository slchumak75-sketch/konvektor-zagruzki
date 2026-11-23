"""Минимальная «заглушка» tkinter для тестов без графического окружения.

Позволяет проверить логику окна (сбор файлов, запуск конвертации, обработку
журнала и результата) в CI на Linux, где tkinter может отсутствовать.
Это не проверка внешнего вида, а проверка того, что код окна не падает.
"""

from __future__ import annotations

import sys
import types
from typing import Dict, List


#: Тексты всех надписей, созданных в окне (для проверки «нет простыней текста»).
TEXTS: List[str] = []
#: Все созданные виджеты ttk (аргументы конструктора) — для проверки стилей.
RECORDS: List[Dict[str, object]] = []


class Widget:
    """Универсальный виджет: принимает любые аргументы, умеет pack/configure."""

    def __init__(self, master=None, *args, **kwargs):
        self.master = master
        self.options: Dict[str, object] = dict(kwargs)
        self.packed = False

    # --- геометрия и свойства
    def pack(self, *a, **k):
        self.packed = True
        return self

    def grid(self, *a, **k):
        return self

    def place(self, *a, **k):
        return self

    def pack_propagate(self, *a, **k):
        return None

    def grid_propagate(self, *a, **k):
        return None

    def pack_configure(self, *a, **k):
        self.options.update(k)
        return None

    def columnconfigure(self, *a, **k):
        return None

    def rowconfigure(self, *a, **k):
        return None

    def __setitem__(self, key, value):
        self.options[key] = value

    def configure(self, *a, **k):
        self.options.update(k)
        return None

    config = configure

    def cget(self, key):
        return self.options.get(key)

    def bind(self, *a, **k):
        return None

    def set(self, *a):
        return None

    def destroy(self):
        return None

    def winfo_exists(self):
        return True

    def focus_set(self):
        self.focused = True

    def focus(self):
        return None

    def place_forget(self):
        self.placed = False

    def after(self, ms, func=None, *args):
        return None

    def winfo_rootx(self):
        return 0

    def winfo_rooty(self):
        return 0

    def winfo_screenwidth(self):
        return 1920

    def winfo_screenheight(self):
        return 1080


class Tk(Widget):
    def __init__(self, *a, **k):
        super().__init__(None, *a, **k)
        self._after_id = 0
        self.mainloop_called = False

    def title(self, *a):
        if a:
            self.options["title"] = a[0]
        return self.options.get("title")

    def geometry(self, *a):
        if a:
            self.options["geometry"] = a[0]
        return self.options.get("geometry")

    def minsize(self, *a):
        return None

    def state(self, *a):
        """Как у настоящего Tk: state("zoomed") разворачивает окно."""
        if a:
            self.options["state"] = a[0]
            return None
        return self.options.get("state", "normal")

    def attributes(self, *a, **k):
        if len(a) == 2:
            self.options["attr:" + str(a[0])] = a[1]
            return None
        return self.options.get("attr:" + str(a[0])) if a else None

    def protocol(self, *a, **k):
        return None

    def after(self, ms, func=None, *args):
        self._after_id += 1
        return f"after#{self._after_id}"  # коллбэки намеренно не вызываются

    def after_cancel(self, *a):
        return None

    def update(self):
        return None

    def mainloop(self, *a, **k):
        self.mainloop_called = True

    def call(self, *a, **k):  # pragma: no cover - используется редко
        return None


class Frame(Widget):
    """Рамка: используется для видимой границы вокруг таблицы и шапки окна."""


class Label(Widget):
    """Надпись (tk.Label): шапка окна, иконка."""


class Toplevel(Widget):
    """Окно (tk.Toplevel): окно ручной сверки с ведомостью."""

    def __init__(self, master=None, *a, **k):
        super().__init__(master, *a, **k)
        self.destroyed = False

    def title(self, *a):
        if a:
            self.options["title"] = a[0]
        return self.options.get("title")

    def geometry(self, *a):
        if a:
            self.options["geometry"] = a[0]
        return self.options.get("geometry")

    def minsize(self, *a):
        return None

    def transient(self, *a, **k):
        return None

    def grab_set(self, *a, **k):
        return None

    def wait_window(self, *a, **k):
        return None

    def resizable(self, *a, **k):
        return None

    def overrideredirect(self, *a, **k):
        return None

    def focus_force(self):
        return None

    def destroy(self):
        self.destroyed = True


class Listbox(Widget):
    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.items: List[str] = []
        self._selection: tuple = ()
        self.row_tags: Dict[str, tuple] = {}

    def insert(self, index, element):
        self.items.append(str(element))

    def delete(self, first, last=None):
        if first == 0 and last == "end":
            self.items.clear()
        elif isinstance(first, int):
            if first < len(self.items):
                del self.items[first]

    def get(self, first=0, last=None):
        if isinstance(first, int):
            return self.items[first] if first < len(self.items) else ""
        return tuple(self.items)

    def size(self):
        return len(self.items)

    def curselection(self):
        return self._selection

    def selection_set(self, idx):
        self._selection = (idx,)

    def see(self, *a):
        return None

    def yview(self, *a):
        return None


class Text(Widget):
    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.content = ""
        self.state = k.get("state", "normal")

    def insert(self, index, chars):
        self.content += str(chars)

    def delete(self, first, last=None):
        self.content = ""

    def see(self, *a):
        return None

    def configure(self, *a, **k):
        self.options.update(k)
        if "state" in k:
            self.state = k["state"]
        return None

    config = configure

    def yview(self, *a):
        return None


class Variable:
    def __init__(self, master=None, value=None, name=None):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value

    def trace_add(self, *a, **k):
        return None


class Combobox(Widget):
    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self._current = -1
        self.values_list: List[str] = list(k.get("values") or [])

    def current(self, index=None):
        if index is None:
            return self._current
        self._current = index
        if 0 <= index < len(self.values_list):
            var = self.options.get("textvariable")
            if isinstance(var, Variable):
                var.set(self.values_list[index])
        return self._current

    def set(self, value):
        var = self.options.get("textvariable")
        if isinstance(var, Variable):
            var.set(value)


class Treeview(Widget):
    """Минимальная таблица: заголовки, колонки, строки."""

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.columns_list: List[str] = list(k.get("columns") or [])
        self.headings: Dict[str, str] = {}
        self.widths: Dict[str, int] = {}
        self.rows: List[tuple] = []
        self.tags: Dict[str, dict] = {}
        self._next_id = 0
        self._selection: tuple = ()
        self.row_tags: Dict[str, tuple] = {}

    def tag_configure(self, name, **k):
        self.tags[name] = dict(k)

    def heading(self, column, **k):
        self.headings[column] = k.get("text", "")

    def column(self, column, option=None, **k):
        if "width" in k:
            self.widths[column] = k["width"]
        if option == "width":
            return self.widths.get(column, 0)

    def insert(self, parent, index, values=None, **k):
        self._next_id += 1
        row_id = f"I{self._next_id:03d}"
        self.rows.append((row_id, tuple(values or ())))
        self.last_tags = tuple(k.get("tags") or ())
        return row_id

    def delete(self, *items):
        drop = {i for i in items if isinstance(i, str)}
        self.rows = [r for r in self.rows if r[0] not in drop]

    def get_children(self, *a, **k):
        return tuple(r[0] for r in self.rows)

    def configure(self, *a, **k):
        self.options.update(k)
        return None

    config = configure

    def yview(self, *a):
        return None

    def xview(self, *a):
        return None

    def identify_region(self, x=0, y=0):
        return getattr(self, "fake_region", "cell")

    def identify_row(self, y=0):
        if getattr(self, "fake_row", None):
            return self.fake_row
        return self.rows[0][0] if self.rows else ""

    def identify_column(self, x=0):
        return getattr(self, "fake_column", "#1")

    def bbox(self, item=None, column=None):
        if item is None or not any(row[0] == item for row in self.rows):
            return ""
        return (0, 0, 200, 23)

    def focus_set(self):
        self.focused = True

    def set(self, item=None, column=None, value=None):
        """Как у настоящего Treeview: прочитать/записать значение ячейки."""
        if column is None:
            return ""
        if item not in self.columns_list:
            # вызов вида set(row_id, column)
            row_id, column = item, column
            if value is not None:
                return self._set_cell(row_id, column, value)
            return self._get_cell(row_id, column)
        return self._get_cell(None, item)

    def _index_of(self, column):
        return self.columns_list.index(column) if column in self.columns_list else -1

    def _get_cell(self, row_id, column):
        index = self._index_of(column)
        for row in self.rows:
            if row_id is None or row[0] == row_id:
                values = list(row[1])
                return values[index] if 0 <= index < len(values) else ""
        return ""

    def _set_cell(self, row_id, column, value):
        index = self._index_of(column)
        for position, row in enumerate(self.rows):
            if row[0] == row_id:
                values = list(row[1])
                while len(values) <= index:
                    values.append("")
                values[index] = value
                self.rows[position] = (row[0], tuple(values))
                break
        return value

    def selection(self, *a):
        return tuple(self._selection)

    def selection_set(self, *items):
        self._selection = tuple(items)
        return None

    def exists(self, item):
        return any(row[0] == item for row in self.rows)

    def item(self, item=None, *a, **k):
        if "values" in k:
            index = None
            for position, row in enumerate(self.rows):
                if row[0] == item:
                    index = position
                    break
            if index is not None:
                self.rows[index] = (item, tuple(k["values"]))
        if "tags" in k:
            self.row_tags[item] = tuple(k["tags"])
        for row in self.rows:
            if row[0] == item:
                return {"values": row[1], "tags": self.row_tags.get(item, ())}
        return {}


class Style(Widget):
    """Заглушка ttk.Style: тема, настройки и карты состояний."""

    _installed_theme = ""

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.settings: Dict[str, dict] = {}
        self.maps: Dict[str, dict] = {}

    def theme_names(self):
        return ("clam", "default")

    def theme_use(self, name=None):
        if name is None:
            return Style._installed_theme or "default"
        Style._installed_theme = name
        return name

    def configure(self, name, **k):
        self.settings.setdefault(name, {}).update(k)

    def map(self, name, **k):
        self.maps.setdefault(name, {}).update(k)

    def layout(self, *a, **k):
        return None


class Progressbar(Widget):
    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.running = False

    def start(self, *a):
        self.running = True

    def stop(self):
        self.running = False


class _Dialog:
    """Подмена filedialog/messagebox: запоминает вызовы, ничего не показывает."""

    def __init__(self):
        self.calls: List[tuple] = []
        self.open_result: tuple = ()
        self.dir_result: str = ""
        self.save_result: str = ""

    def __getattr__(self, name):
        def fake(*a, **k):
            self.calls.append((name, a, k))
            if name == "askopenfilenames":
                return self.open_result
            if name == "askopenfilename":
                return self.open_result[0] if self.open_result else ""
            if name == "askdirectory":
                return self.dir_result
            if name == "asksaveasfilename":
                return self.save_result
            return None

        return fake


def install() -> types.ModuleType:
    """Создать модуль tkinter-заглушки и положить его в sys.modules."""
    tk = types.ModuleType("tkinter")
    ttk = types.ModuleType("tkinter.ttk")
    filedialog = types.ModuleType("tkinter.filedialog")
    messagebox = types.ModuleType("tkinter.messagebox")

    for cls in (
        Widget,
        Tk,
        Frame,
        Label,
        Listbox,
        Text,
        Toplevel,
        Variable,
        Combobox,
        Progressbar,
        Treeview,
        Style,
    ):
        setattr(tk, cls.__name__, cls)
    tk.StringVar = Variable
    tk.BooleanVar = Variable
    tk.IntVar = Variable
    tk.END = "end"
    tk.TkVersion = 8.6

    class _RecordingLabel(Widget):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            if "text" in k:
                TEXTS.append(str(k["text"]))
            RECORDS.append(dict(k))

    class Menu(Widget):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self.items: List[Dict[str, object]] = []

        def add_command(self, **k):
            self.items.append(dict(k))
            if "label" in k:
                TEXTS.append(str(k["label"]))
                RECORDS.append({"text": str(k["label"]), **k})

        def add_separator(self):
            self.items.append({"separator": True})

        def entryconfigure(self, index: int, **k):
            if 0 <= index < len(self.items):
                self.items[index].update(k)

        def entrycget(self, index: int, key: str):
            if 0 <= index < len(self.items):
                return self.items[index].get(key)
            return None

    class Canvas(Widget):
        def create_window(self, *a, **k):
            return 1

        def yview(self, *a, **k):
            return (0.0, 1.0)

        def bbox(self, *a, **k):
            return (0, 0, 100, 100)

    tk.Menu = Menu
    tk.Canvas = Canvas

    TEXTS.clear()
    RECORDS.clear()
    for name in (
        "Frame",
        "Label",
        "LabelFrame",
        "Button",
        "Menubutton",
        "Entry",
        "Checkbutton",
        "Radiobutton",
        "Scrollbar",
        "Separator",
        "Notebook",
    ):
        setattr(ttk, name, type(name, (_RecordingLabel,), {}))
    ttk.Combobox = Combobox
    ttk.Progressbar = Progressbar
    ttk.Treeview = Treeview
    ttk.Style = Style

    dialogs = _Dialog()
    for name in ("askopenfilenames", "askopenfilename", "askdirectory", "asksaveasfilename"):
        setattr(filedialog, name, getattr(dialogs, name))
    for name in ("showinfo", "showwarning", "showerror", "askyesno"):
        setattr(messagebox, name, getattr(dialogs, name))

    tk.ttk = ttk
    tk.filedialog = filedialog
    tk.messagebox = messagebox

    sys.modules["tkinter"] = tk
    sys.modules["tkinter.ttk"] = ttk
    sys.modules["tkinter.filedialog"] = filedialog
    sys.modules["tkinter.messagebox"] = messagebox
    return dialogs  # type: ignore[return-value]
