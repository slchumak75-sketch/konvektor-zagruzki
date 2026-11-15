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
def gui(monkeypatch, known_paths, tmp_path):
    dialogs = fake_tk.install()
    # «Документы» подменяем временной папкой, чтобы тесты не писали в профиль
    import konvektor.pipeline as pipeline_module

    documents = tmp_path / "Документы"
    documents.mkdir()
    monkeypatch.setattr(pipeline_module, "documents_dir", lambda: str(documents))
    import konvektor.gui as gui_module

    module = importlib.reload(gui_module)
    monkeypatch.setattr(gui_module, "Tk", fake_tk.Tk, raising=False)
    # файл с известным поставщиком и датой: ООО "ОМА", 19.10.2026
    sample = next(
        (p for p in known_paths if p.name.startswith("Кореличи")), known_paths[0]
    )
    app = module.KonvektorApp(initial_files=[str(sample)])
    app.fake_dialogs = dialogs
    app.test_documents = str(documents)
    yield app, module, dialogs
    # вернуть «настоящее» состояние модуля
    for name in ("tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox"):
        sys.modules.pop(name, None)
    importlib.reload(gui_module)


def _run(app):
    """Дождаться фонового потока и разобрать очередь сообщений."""
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
    # предпросмотр — ровно 10 колонок шаблона, без лишних
    assert len(app.preview.columns_list) == 10
    assert app.preview.headings["Наименование товара"] == "Наименование товара"
    assert "файл" not in app.preview.headings
    assert "строка" not in app.preview.headings
    # разлиновка: чередование строк настроено
    assert "odd" in app.preview.tags and "even" in app.preview.tags


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


def test_convert_only_previews_and_save_writes_file(gui):
    """«Конвертировать» показывает таблицу, файл пишет только «Сохранить»."""
    app, module, _ = gui
    app.on_convert()
    _run(app)
    # ничего не записано, но путь известен
    assert app._last_outputs == []
    assert not os.listdir(app.test_documents)
    assert app._pending_result is not None and app._pending_result.items_count == 11
    expected = os.path.join(app.test_documents, "ООО ОМА 19.10.2026.xlsx")
    assert app._pending_result.planned_output == expected
    assert app.save_btn.cget("state") == "normal"
    # подсказка не затирается автоочисткой списка
    assert "будет сохранено" in app.hint_var.get()

    app.on_save()
    _run(app)
    assert app._last_outputs == [expected]
    assert os.path.exists(expected)
    assert os.listdir(app.test_documents) == ["ООО ОМА 19.10.2026.xlsx"]


def test_no_status_field_and_no_progressbar(gui, known_paths):
    """Надпись «Готов к работе» и шкала прогресса убраны."""
    app, _, _ = gui
    assert not hasattr(app, "status")
    assert not hasattr(app, "progress")

    app._clear_files()
    assert app.hint_var.get() == ""

    other = next(p for p in known_paths if p.name.startswith("КХ2104033"))
    app._add_files([str(other)])
    assert app.hint_var.get().startswith("Файл выбран")


def test_explicit_output_path_is_used(gui, tmp_path):
    app, _, _ = gui
    target = tmp_path / "свой.xlsx"
    app.out_var.set(str(target))
    app.on_convert()
    _run(app)
    assert not target.exists()  # конвертация ничего не пишет
    app.on_save()
    _run(app)
    assert target.exists()
    assert app._last_outputs == [str(target)]


def test_save_button_is_disabled_until_conversion(gui):
    app, _, dialogs = gui
    assert app.save_btn.cget("state") == "disabled"
    app.on_save()  # ничего не должно произойти
    assert app._last_outputs == []


def test_preview_shows_converted_table(gui):
    app, _, _ = gui
    app.on_convert()
    _run(app)
    rows = app.preview.rows
    assert len(rows) == 11  # «Кореличи» — 11 позиций
    first = rows[0][1]
    assert len(first) == 10  # только колонки шаблона
    assert first[2]  # наименование
    assert "," in first[5]  # числа показываем как в Excel — с запятой


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
