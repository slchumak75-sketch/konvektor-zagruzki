"""Минимальная «заглушка» tkinter для тестов без графического окружения.

Позволяет проверить логику окна (сбор файлов, запуск конвертации, обработку
журнала и результата) в CI на Linux, где tkinter может отсутствовать.
Это не проверка внешнего вида, а проверка того, что код окна не падает.
"""

from __future__ import annotations

import sys
import types
from typing import Dict, List


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


class Listbox(Widget):
    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.items: List[str] = []
        self._selection: tuple = ()

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

    for cls in (Widget, Tk, Listbox, Text, Variable, Combobox, Progressbar):
        setattr(tk, cls.__name__, cls)
    tk.StringVar = Variable
    tk.BooleanVar = Variable
    tk.IntVar = Variable
    tk.END = "end"
    tk.TkVersion = 8.6

    for name in (
        "Frame",
        "Label",
        "LabelFrame",
        "Button",
        "Entry",
        "Checkbutton",
        "Scrollbar",
        "Separator",
        "Notebook",
    ):
        setattr(ttk, name, type(name, (Widget,), {}))
    ttk.Combobox = Combobox
    ttk.Progressbar = Progressbar

    dialogs = _Dialog()
    for name in ("askopenfilenames", "askdirectory", "asksaveasfilename"):
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
