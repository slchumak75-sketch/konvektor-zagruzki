"""Проверка логики окна без графического окружения (tkinter-заглушка).

Проверяется поведение, заданное заказчиком: один файл за раз, автоочистка
списка после конвертации, предпросмотр таблицы вместо журнала, сохранение
в «поставщик/дата.xlsx».
"""

from __future__ import annotations

import importlib
import os
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
    # файл с известным поставщиком и датой: ООО "ОМА", 19.10.2026
    sample = next(
        (p for p in known_paths if p.name.startswith("Кореличи")), known_paths[0]
    )
    app = module.KonvektorApp(initial_files=[str(sample)])
    app.fake_dialogs = dialogs
    yield app, module, dialogs
    # вернуть «настоящее» состояние модуля
    for name in ("tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox"):
        sys.modules.pop(name, None)
    importlib.reload(gui_module)


def _run(app):
    deadline = time.time() + 120
    while time.time() < deadline:
        if app._worker is not None and not app._worker.is_alive() and app._queue.empty():
            break
        time.sleep(0.05)
        app._poll_queue()
    app._poll_queue()


def test_window_builds_with_single_file(gui, known_paths):
    app, module, _ = gui
    assert len(app.files) == 1
    assert app.file_var.get().startswith("Кореличи")
    assert "Конвектор" in app.title()
    # предпросмотр — таблица шаблона + «файл» и «строка»
    assert len(app.preview.columns_list) == 12
    assert app.preview.headings["файл"] == "файл"
    assert app.preview.headings["строка"] == "строка"


def test_only_one_file_at_a_time(gui, known_paths):
    app, _, _ = gui
    app._add_files([str(known_paths[1])])
    assert len(app.files) == 1
    assert app.files[0].endswith(known_paths[1].name)
    app._add_files([str(known_paths[2])])
    assert len(app.files) == 1
    assert app.files[0].endswith(known_paths[2].name)


def test_unsupported_files_are_ignored(gui, tmp_path):
    app, _, _ = gui
    before = list(app.files)
    junk = tmp_path / "документ.docx"
    junk.write_text("не таблица")
    app._add_files([str(junk)])
    assert app.files == before


def test_default_output_name_is_supplier_and_date(gui):
    """Пустое поле «Итоговый файл» = поставщик/дата.xlsx рядом с исходником."""
    app, module, _ = gui
    assert app.out_var.get() == ""
    app.on_convert()
    _run(app)
    assert app._last_outputs, "результат не записан"
    path = app._last_outputs[0]
    folder, name = os.path.split(path)
    assert name == "19.10.2026.xlsx", path
    assert os.path.basename(folder) == "ООО ОМА", path
    assert os.path.exists(path)


def test_explicit_output_path_is_used(gui, tmp_path):
    app, _, _ = gui
    target = tmp_path / "свой.xlsx"
    app.out_var.set(str(target))
    app.on_convert()
    _run(app)
    assert target.exists()
    assert app._last_outputs == [str(target)]


def test_preview_shows_converted_table(gui):
    app, _, _ = gui
    app.on_convert()
    _run(app)
    rows = app.preview.rows
    assert len(rows) == 11  # «Кореличи» — 11 позиций
    first = rows[0][1]
    assert first[2].startswith("Грунтовка") or first[2]
    assert first[10].startswith("Кореличи")  # колонка «файл»
    assert first[11]  # колонка «строка»


def test_list_is_cleared_after_conversion(gui):
    app, _, _ = gui
    app.on_convert()
    _run(app)
    assert app.files == []
    assert app.file_var.get() == "файл не выбран"
    assert app.out_var.get() == ""


def test_convert_without_file_shows_warning(gui):
    app, _, dialogs = gui
    app._clear_files()
    app.on_convert()
    assert any(call[0] == "showwarning" for call in dialogs.calls)
    assert app._running is False


def test_no_folder_or_list_buttons_left(gui):
    """Кнопок «Добавить папку», «Удалить выбранные», «Очистить список» больше нет."""
    app, module, _ = gui
    assert not hasattr(app, "on_add_folder")
    assert not hasattr(app, "on_remove")
    assert not hasattr(app, "on_clear")
    assert not hasattr(app, "open_after_var")
    assert not hasattr(app, "log_text")
    assert hasattr(app, "preview")
