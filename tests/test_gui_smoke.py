"""Проверка логики окна без графического окружения (tkinter-заглушка).

Тест не смотрит на внешний вид, но ловит опечатки в обработчиках: добавление
файлов, запуск конвертации в потоке, обработку очереди журнала и результата.
"""

from __future__ import annotations

import importlib
import sys
import time

import pytest

from tests import fake_tk


@pytest.fixture()
def gui(monkeypatch, known_paths):
    dialogs = fake_tk.install()
    import konvektor.gui as gui_module

    module = importlib.reload(gui_module)
    monkeypatch.setattr(gui_module, "Tk", fake_tk.Tk, raising=False)
    app = module.KonvektorApp(initial_files=[str(p) for p in known_paths])
    app.fake_dialogs = dialogs
    yield app, module, dialogs
    # вернуть «настоящее» состояние модуля
    for name in ("tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox"):
        sys.modules.pop(name, None)
    importlib.reload(gui_module)


def test_window_builds_and_lists_files(gui, known_paths):
    app, module, _ = gui
    assert len(app.files) == len(known_paths)
    assert app.listbox.size() == len(known_paths)
    assert "Конвектор" in app.title()
    assert app.out_var.get().endswith(".xlsx")


def test_duplicates_and_unsupported_are_filtered(gui, tmp_path):
    app, _, _ = gui
    before = len(app.files)
    app._add_files(list(app.files))  # повторно те же файлы
    assert len(app.files) == before
    junk = tmp_path / "документ.docx"
    junk.write_text("не таблица")
    app._add_files([str(junk)])
    assert len(app.files) == before


def test_remove_and_clear(gui):
    app, _, _ = gui
    app.listbox.selection_set(0)
    app.on_remove()
    assert len(app.files) == app.listbox.size()
    app.on_clear()
    assert app.files == []
    assert app.listbox.size() == 0


def test_convert_flow(gui, tmp_path):
    app, _, dialogs = gui
    app.out_var.set(str(tmp_path / "gui_out.xlsx"))
    app.open_after_var.set(False)
    app.on_convert()

    assert app._running is True
    assert app.run_btn.options.get("state") == "disabled"

    # дождаться фонового потока и разобрать очередь
    deadline = time.time() + 120
    while time.time() < deadline:
        if app._worker is not None and not app._worker.is_alive() and app._queue.empty():
            break
        time.sleep(0.05)
        app._poll_queue()

    app._poll_queue()
    assert app._running is False
    assert "позиций" in app.status.cget("text").lower()
    assert (tmp_path / "gui_out.xlsx").exists()
    assert "Готово" in app.log_text.content or "позиций" in app.log_text.content
    assert any(call[0] == "showinfo" for call in dialogs.calls)


def test_convert_without_files_shows_warning(gui):
    app, _, dialogs = gui
    app.on_clear()
    app.on_convert()
    assert any(call[0] == "showwarning" for call in dialogs.calls)
    assert app._running is False


def test_log_clear_and_status(gui):
    app, _, _ = gui
    app._log("привет")
    assert "привет" in app.log_text.content
    app._log_clear()
    assert app.log_text.content == ""
    app._update_status()
    assert "Файлов" in app.status.cget("text")
