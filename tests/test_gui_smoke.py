"""Проверка логики окна без графического окружения (tkinter-заглушка).

Проверяется поведение, заданное заказчиком: один файл за раз, автоочистка
списка после конвертации, предпросмотр таблицы вместо журнала, сохранение
в «поставщик/дата.xlsx».
"""

from __future__ import annotations

import importlib
import datetime as dt
import os
import sys
import time

import pytest

from konvektor.vedomost import MATCH_ARTICLE, MATCH_BARCODE, MATCH_NAME
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
    # файл с известным поставщиком: ООО "ОМА" (дата в имени — сегодняшняя)
    sample = next(
        (p for p in known_paths if p.name.startswith("Кореличи")), known_paths[0]
    )
    app = module.KonvertorApp(initial_files=[str(sample)])
    app.fake_dialogs = dialogs
    app.test_documents = str(documents)
    yield app, module, dialogs
    # вернуть «настоящее» состояние модуля
    for name in ("tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox"):
        sys.modules.pop(name, None)
    importlib.reload(gui_module)


TODAY = f"{dt.date.today():%d.%m.%Y}"


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
    assert "Конвертор" in app.title()
    # в окне две колонки наименования, «▼» для выбора и процент совпадения
    assert list(app.preview.columns_list) == [
        "Штрихкод", "Артикул", "Из документа", "Из ведомости", "▼", "% совпадения",
        "ед. изм.", "количество", "цена", "сумма", "ставка НДС", "сумма НДС",
        "сумма с НДС",
    ]
    # «▼» — узкая колонка-кнопка сразу справа от названия
    assert app.preview.widths["▼"] <= 30
    assert app.preview.headings["Из документа"] == "Из документа"
    assert app.preview.headings["Из ведомости"] == "Из ведомости"
    assert "файл" not in app.preview.headings
    assert "строка" not in app.preview.headings
    # числовые колонки и «ед. изм.» в полтора раза уже, чем наименования
    assert app.preview.widths["количество"] < app.preview.widths["Из документа"] / 3
    assert app.preview.widths["сумма с НДС"] <= 80
    assert app.preview.widths["ед. изм."] == 50
    # без ведомости лишние колонки скрыты — таблица как раньше
    assert app.preview.options["displaycolumns"] == [
        "Штрихкод", "Артикул", "Из документа", "ед. изм.", "количество", "цена",
        "сумма", "ставка НДС", "сумма НДС", "сумма с НДС",
    ]
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
    expected = os.path.join(app.test_documents, f"ООО ОМА {TODAY}.xlsx")
    assert app._pending_result.planned_output == expected
    assert app.save_btn.cget("state") == "normal"
    # подсказка не затирается автоочисткой списка
    assert "будет сохранено" in app.hint_var.get()

    app.on_save()
    _run(app)
    assert app._last_outputs == [expected]
    assert os.path.exists(expected)
    assert os.listdir(app.test_documents) == [f"ООО ОМА {TODAY}.xlsx"]

    # в файле ровно 10 колонок образца: после «сумма с НДС» ничего нет
    from konvektor import TEMPLATE_COLUMNS
    from openpyxl import load_workbook

    header = [c.value for c in load_workbook(expected)["Данные"][1]]
    assert header == TEMPLATE_COLUMNS
    assert "файл" not in header and "строка" not in header


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


def test_output_field_is_folder_and_name_is_supplier_date(gui, tmp_path):
    """В поле — только папка: имя файла всегда «поставщик дата.xlsx»."""
    app, _, _ = gui
    folder = tmp_path / "выбранная папка"
    app.out_var.set(str(folder))
    app.on_convert()
    _run(app)
    assert not folder.exists() or list(folder.iterdir()) == []  # ничего не пишем
    expected = str(folder / f"ООО ОМА {TODAY}.xlsx")
    assert app._pending_result.planned_output == expected
    app.on_save()
    _run(app)
    assert [p.name for p in folder.iterdir()] == [f"ООО ОМА {TODAY}.xlsx"]
    assert app._last_outputs == [expected]


def test_file_path_in_folder_field_still_gives_supplier_date_name(gui, tmp_path):
    """Даже если вставить полный путь к файлу, «Результат.xlsx» не появится."""
    app, _, _ = gui
    folder = tmp_path / "вставленный путь"
    app.out_var.set(str(folder / "Результат.xlsx"))
    assert app.output_dir() == str(folder)
    app.on_convert()
    _run(app)
    app.on_save()
    _run(app)
    names = [p.name for p in folder.iterdir()]
    assert names == [f"ООО ОМА {TODAY}.xlsx"], names


def test_quantity_and_vat_have_no_decimal_comma(gui):
    """В «количество» и «ставке НДС» лишней запятой нет: 5, а не 5,00; 20, а не 20,00."""
    app, _, _ = gui
    assert app._fmt(5.0, "количество") == "5"
    assert app._fmt(12.0, "количество") == "12"
    assert app._fmt(0.5, "количество") == "0,5"
    assert app._fmt(20.0, "ставка НДС") == "20"
    assert app._fmt(10.0, "ставка НДС") == "10"
    # деньги остаются с копейками
    assert app._fmt(2.5, "цена") == "2,50"
    assert app._fmt(15.47, "сумма с НДС") == "15,47"

    app.on_convert()
    _run(app)
    rows = [dict(zip(app.preview.columns_list, row[1])) for row in app.preview.rows]
    quantities = [r["количество"] for r in rows if r["количество"]]
    rates = {r["ставка НДС"] for r in rows if r["ставка НДС"]}
    assert quantities and rates
    assert not any(q.endswith(",00") for q in quantities), quantities
    assert all("," not in rate for rate in rates), rates


def test_path_field_with_browse_button(gui, tmp_path):
    """Обычное поле пути + кнопка «Обзор…»; других кнопок выбора нет."""
    app, module, dialogs = gui
    # по умолчанию подставлена папка «Документы»
    assert app.out_var.get() == module.documents_dir()
    assert hasattr(app, "on_browse_output")
    assert not hasattr(app, "on_use_documents")
    assert not hasattr(app, "on_use_downloads")

    # «Обзор…» подставляет выбранную папку в то же поле
    chosen = tmp_path / "выбранная"
    chosen.mkdir()
    dialogs.dir_result = str(chosen)
    app.on_browse_output()
    assert app.out_var.get() == str(chosen)
    assert any(call[0] == "askdirectory" for call in dialogs.calls)
    # путь можно стереть — тогда «Документы»
    app.out_var.set("   ")
    assert app.output_dir() == ""
    # и вписать любой другой
    app.out_var.set("D:\\Мои документы\\Конвертор")
    assert app.output_dir() == "D:\\Мои документы\\Конвертор"


def test_second_save_does_not_overwrite(gui):
    """Повторное сохранение создаёт новый файл, а не затирает предыдущий."""
    app, _, _ = gui
    source = str(app.files[0])

    app.on_convert()
    _run(app)
    app.on_save()
    _run(app)
    first = app._last_outputs[0]
    assert os.path.exists(first)

    # тот же документ разбираем и сохраняем ещё раз
    app._add_files([source])
    app.on_convert()
    _run(app)
    app.on_save()
    _run(app)
    second = app._last_outputs[0]
    assert second != first
    assert second.endswith("(2).xlsx"), second
    assert os.path.exists(first) and os.path.exists(second)
    assert sorted(os.listdir(app.test_documents)) == [
        f"ООО ОМА {TODAY} (2).xlsx",
        f"ООО ОМА {TODAY}.xlsx",
    ]


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
    # 10 колонок шаблона + «из ведомости», «▼» и «% совпадения»
    assert len(first) == 13 == len(app.preview.columns_list)
    assert first[2]  # наименование из документа
    assert first[3] == ""  # ведомость не загружена — подбирать не из чего
    assert first[4] == ""  # «▼» появляется только с ведомостью
    row = dict(zip(app.preview.columns_list, first))
    # числа показываем как в Excel — с запятой
    assert "," in row["цена"] or "," in row["сумма с НДС"]


def test_list_is_cleared_after_conversion(gui):
    app, module, _ = gui
    app.out_var.set(r"D:\Итоги")
    app.on_convert()
    _run(app)
    # файл остаётся в буфере до нажатия «Очистить таблицу»
    assert len(app.files) == 1
    app.on_clear_table()
    assert app.files == []
    assert app.file_var.get() == "файл не выбран"
    # путь сохранения — настройка, а не имя файла: он остаётся
    assert app.out_var.get() == r"D:\Итоги"


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


def test_icon_assets_exist():
    """Иконка приложения: PNG для окна и многоразмерный .ico для .exe."""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    png = os.path.join(root, "assets", "konvektor_icon.png")
    ico = os.path.join(root, "assets", "konvektor.ico")
    assert os.path.exists(png) and os.path.getsize(png) > 2000
    assert os.path.exists(ico) and os.path.getsize(ico) > 2000
    with open(ico, "rb") as fh:
        head = fh.read(6)
    assert head[:4] == b"\x00\x00\x01\x00"  # сигнатура ICO


def test_open_buttons_follow_save(gui, monkeypatch):
    """«Открыть файл» и «Открыть папку» доступны только после сохранения."""
    app, module, _ = gui
    assert app.open_file_btn.cget("state") == "disabled"
    assert app.open_folder_btn.cget("state") == "disabled"

    opened = []
    monkeypatch.setattr(module.KonvertorApp, "_open_path", staticmethod(opened.append))

    app.on_open_result()          # файла ещё нет — только сообщение
    assert opened == []

    app.on_convert()
    _run(app)
    app.on_save()
    _run(app)
    assert app.open_file_btn.cget("state") == "normal"
    assert app.open_folder_btn.cget("state") == "normal"

    app.on_open_result()
    assert opened and opened[-1] == app._last_outputs[0]
    app.on_open_folder()
    assert opened[-1] == os.path.dirname(app._last_outputs[0])


def test_new_conversion_disables_open_buttons(gui):
    app, _, _ = gui
    app.on_convert()
    _run(app)
    app.on_save()
    _run(app)
    assert app.open_file_btn.cget("state") == "normal"
    app._add_files([str(app._pending_source)])
    app.on_convert()
    _run(app)
    assert app.open_file_btn.cget("state") == "disabled"
    assert app.open_folder_btn.cget("state") == "disabled"


def test_path_changed_after_convert_is_used(gui, tmp_path):
    """Путь, изменённый после «Конвертировать», уважается при «Сохранить».

    Раньше сохранялся снимок пути времён конвертации, и файл уходил
    в старую папку — «сохраняет не туда, куда указываю».
    """
    app, _, _ = gui
    new_folder = tmp_path / "куда указал"
    app.on_convert()
    _run(app)
    app.out_var.set(str(new_folder))
    app.on_save()
    _run(app)
    assert [p.name for p in new_folder.iterdir()] == [f"ООО ОМА {TODAY}.xlsx"]
    assert app._last_outputs == [str(new_folder / f"ООО ОМА {TODAY}.xlsx")]


def test_path_is_normalised(gui, tmp_path, monkeypatch):
    """Кавычки, лишние пробелы, «~», полный путь к файлу — всё приводится к папке."""
    app, module, _ = gui
    folder = tmp_path / "итоги"
    folder.mkdir()
    app.out_var.set(f'  "{folder}"  ')
    assert app.output_dir() == str(folder)

    app.out_var.set(str(folder) + os.sep)
    assert app.output_dir() == str(folder)

    # указали файл — берём его папку
    app.out_var.set(str(folder / "ООО ОМА 15.11.2026.xlsx"))
    assert app.output_dir() == str(folder)
    target = folder / "старый.xlsx"
    target.write_text("x")
    app.out_var.set(str(target))
    assert app.output_dir() == str(folder)

    # «~» раскрывается в домашнюю папку
    app.out_var.set("~")
    assert app.output_dir() == os.path.expanduser("~")


def test_save_warns_when_written_elsewhere(gui, tmp_path):
    """Если записать по указанному пути не вышло, видно и куда просили, и куда сохранили."""
    app, module, dialogs = gui
    app.on_convert()
    _run(app)

    planned = str(tmp_path / "куда просили" / f"ООО ОМА {TODAY}.xlsx")
    actual = str(tmp_path / f"ООО ОМА {TODAY}.xlsx")
    result = app._pending_result
    result.planned_output = planned
    result.outputs = [actual]
    result.notes = ["не удалось записать ...; результат сохранён в ..."]

    dialogs.calls.clear()
    app._saved(result)
    assert any(call[0] == "showwarning" for call in dialogs.calls)
    warned = [c for c in dialogs.calls if c[0] == "showwarning"][0]
    text = warned[1][1]
    assert planned in text and actual in text


# --------------------------------------------------- ведомость и максимизация
VEDOMOST_HEADER = [
    "Склад",
    "Группа номенклатуры",
    "Артикул",
    "Номенклатура",
    "Ед.",
    "Конечный\nостаток",
    "Цена\nпоследнего \nпоставщика",
    "Последний\nпоставщик",
]

VEDOMOST_ROWS = [
    ["Склад Мир", "ГАЗ", "", "Газ сжижен.в однораз.баллоне WERKER 520см3", "шт", "7", "3,20", 'ООО "ОМА"'],
    ["Склад Мир", "ГЕРМ", "", "Герметик акрил. бел. 280мл WUNDER", "шт", "4", "5,10", 'ООО "ОМА"'],
    # та же номенклатура, но у чужого поставщика — подбираться не должна
    ["Склад Мир", "ГАЗ", "", "Газ сжижен.в однораз.баллоне WERKER 520см3", "шт", "2", "3,00", 'ООО "ТД Комплект"'],
]


def _write_vedomost(tmp_path) -> str:
    """Сделать файл ведомости остатков ТМЦ (те же колонки, что у заказчика)."""
    import openpyxl

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Лист_1"
    sheet.append([])
    sheet.append(["Ведомость по остаткам ТМЦ на складах"])
    sheet.append([])
    sheet.append(list(VEDOMOST_HEADER))
    for row in VEDOMOST_ROWS:
        sheet.append(list(row))
    path = tmp_path / "Ведомость 08.11.2026.xlsx"
    book.save(path)
    return str(path)


def test_window_starts_maximized(gui):
    """Пользователь просил, чтобы окно сразу было развёрнуто на весь экран."""
    app, _, _ = gui
    assert app.state() == "zoomed"


def test_cards_are_compact_without_helper_texts(gui):
    """Плашки «ШАГ 1» и «ШАГ 2» компактные: длинных пояснений в окне нет."""
    app, _, _ = gui
    long_texts = [t for t in fake_tk.TEXTS if len(t) > 120]
    assert long_texts == []
    joined = " ".join(fake_tk.TEXTS)
    assert "Поддерживаются .xls" not in joined  # старый текст подсказки
    assert "Впишите путь" not in joined
    # а нужные элементы на месте
    assert any("Загрузить ведомость" in t for t in fake_tk.TEXTS)
    assert app.ved_var.get() == "не загружена"


def test_three_convert_buttons(gui):
    """Три кнопки конвертации: по штрихкоду, по артикулу, по наименованию."""
    app, module, _ = gui
    assert set(app.convert_buttons) == {MATCH_BARCODE, MATCH_ARTICLE, MATCH_NAME}
    texts = [
        t
        for t in fake_tk.TEXTS
        if t.startswith("Конвертировать ") and t != "Конвертировать ▾"
    ]
    assert texts == [
        "Конвертировать по штрихкоду",
        "Конвертировать по артикулу",
        "Конвертировать по наименованию",
    ]
    # основная кнопка — «по наименованию»
    assert app.run_btn is app.convert_buttons[MATCH_NAME]


def test_load_vedomost(gui, tmp_path):
    app, _, _ = gui
    path = _write_vedomost(tmp_path)
    assert app.load_vedomost(path) is True
    assert app._vedomost is not None
    assert app._vedomost.count == len(VEDOMOST_ROWS)
    assert "позиций 3" in app.ved_var.get()
    assert "Ведомость 08.11.2026.xlsx" in app.ved_var.get()
    assert "Ведомость загружена" in app.hint_var.get()


def test_load_bad_vedomost_shows_reason(gui, tmp_path):
    app, _, dialogs = gui
    junk = tmp_path / "не ведомость.txt"
    junk.write_text("просто текст\nбез таблицы", encoding="utf-8")
    assert app.load_vedomost(str(junk)) is False
    assert app.ved_var.get() == "не загружена"
    assert "не загружена" in app.hint_var.get()
    assert any(call[0] == "showwarning" for call in dialogs.calls)


def test_convert_with_vedomost_shows_two_name_columns(gui, tmp_path):
    """С ведомостью в таблице видно и название документа, и подобранное."""
    app, _, _ = gui
    app.load_vedomost(_write_vedomost(tmp_path))
    app.on_convert(MATCH_NAME)
    _run(app)

    assert "сверено с ведомостью" in app.hint_var.get()
    assert "будет сохранено" in app.hint_var.get()
    # колонки «из ведомости» и «% совпадения» появились
    assert "Из ведомости" in app.preview.options["displaycolumns"]
    assert "% совпадения" in app.preview.options["displaycolumns"]
    from_doc = [row[1][2] for row in app.preview.rows]
    from_ved = [row[1][3] for row in app.preview.rows]
    assert any(name.startswith("Газ сжиженный WERKER") for name in from_doc)
    assert "Газ сжижен.в однораз.баллоне WERKER 520см3" in from_ved
    assert "Герметик акрил. бел. 280мл WUNDER" in from_ved
    # позиции, которых нет в ведомости, остаются с наименованием документа
    assert any(name and "Гидроизол ХКП S10м2 сер. к-сх-б-к/пп-3,5" in name for name in from_doc)
    assert any(name and "Гидроизол" in name for name in from_ved) is False


def test_save_uses_vedomost_names(gui, tmp_path):
    """Сохранённый файл содержит наименования из ведомости."""
    import openpyxl

    app, _, _ = gui
    app.load_vedomost(_write_vedomost(tmp_path))
    out_dir = tmp_path / "Куда сохранить"
    app.out_var.set(str(out_dir))
    app.on_convert(MATCH_NAME)
    _run(app)
    app.on_save()
    _run(app)

    written = app._last_outputs[0]
    assert os.path.dirname(written) == str(out_dir)
    book = openpyxl.load_workbook(written)
    names = [row[2] for row in book["Данные"].iter_rows(min_row=2, values_only=True)]
    assert "Газ сжижен.в однораз.баллоне WERKER 520см3" in names


def test_convert_strategy_is_passed_to_matcher(gui, tmp_path):
    app, _, _ = gui
    app.load_vedomost(_write_vedomost(tmp_path))
    app.on_convert(MATCH_ARTICLE)
    _run(app)
    assert app._strategy == MATCH_ARTICLE
    assert app._matcher.strategy == MATCH_ARTICLE
    assert "по артикулу" in app.hint_var.get() or "сверено" in app.hint_var.get()




# --------------------------------- выбор номенклатуры из ведомости («▼»)
class _Event:
    """Простое событие мыши для щелчка по таблице."""

    def __init__(self, x=0, y=0):
        self.x = x
        self.y = y
        self.num = 0
        self.delta = 0


def _open_vedomost_and_convert(app, tmp_path):
    """Загрузить ведомость, сконвертировать образец «Кореличи»."""
    app.load_vedomost(_write_vedomost(tmp_path))
    app.on_convert(MATCH_NAME)
    _run(app)


def _row_with_doc(app, startswith: str):
    for row_id, stored in app._preview_rows.items():
        entry = app._entry_for(stored["key"])
        if entry is not None and entry.doc_name.startswith(startswith):
            return row_id, stored
    raise AssertionError("позиция не найдена: " + startswith)


def _click_cell(app, row_id, column: str, handler=None):
    """Щёлкнуть по ячейке указанной колонки этой строки."""
    app.preview.fake_row = row_id
    app.preview.fake_column = "#" + str(app.preview.columns_list.index(column) + 1)
    (handler or app._on_preview_click)(_Event(10, 10))


def _choose(app, index: int):
    """Выбрать пункт `index` в открытом выпадающем списке."""
    assert app._chooser is not None
    app._chooser.box.selection_set(index)
    app._chooser.box._selection = (index,)
    app._chooser._on_click()


def _values(app, row_id):
    target = next(row[1] for row in app.preview.rows if row[0] == row_id)
    return dict(zip(app.preview.columns_list, target))


def test_arrow_opens_list_only_in_vedomost_columns(gui, tmp_path):
    """«▼» открывает список номенклатуры; другие колонки — нет."""
    app, _, _ = gui
    _open_vedomost_and_convert(app, tmp_path)
    row_id, _stored = _row_with_doc(app, "Газ")

    # щелчок по «Из документа» список не открывает
    _click_cell(app, row_id, "Из документа")
    assert app._chooser is None

    _click_cell(app, row_id, "▼")
    assert app._chooser is not None
    labels = app._chooser.labels()
    # первый пункт — «не выбирать», пункта «Своё название» нет
    assert labels[0].startswith("— не выбирать")
    assert "Своё название…" not in labels
    # похожие позиции идут с процентом совпадения
    assert any("%" in label for label in labels[1:])


def test_choosing_row_from_vedomost_renames_item(gui, tmp_path):
    app, _, _ = gui
    _open_vedomost_and_convert(app, tmp_path)
    row_id, stored = _row_with_doc(app, "Газ")

    _click_cell(app, row_id, "▼")
    wanted = "Газ сжижен.в однораз.баллоне WERKER 520см3"
    index = next(
        i
        for i, (_label, value) in enumerate(app._chooser.items)
        if getattr(value, "nomenclature", "") == wanted
    )
    _choose(app, index)

    override = app._overrides[stored["key"]]
    assert getattr(override, "nomenclature", "") == wanted
    values = _values(app, row_id)
    assert values["Из ведомости"] == wanted
    assert values["% совпадения"] == "вручную"
    assert app.preview.row_tags[row_id] == ("manual",)
    assert stored["item"].name == wanted
    # список закрылся
    assert app._chooser is None


def test_do_not_choose_keeps_doc_name_and_marks_novinka(gui, tmp_path):
    app, _, _ = gui
    _open_vedomost_and_convert(app, tmp_path)
    row_id, stored = _row_with_doc(app, "Газ")

    _click_cell(app, row_id, "▼")
    _choose(app, 0)  # «— не выбирать (новый товар) —»

    assert app._overrides[stored["key"]] is None
    values = _values(app, row_id)
    assert values["Из ведомости"] == ""
    assert values["% совпадения"] == "не выбирать"
    # наименование документа не тронуто
    assert stored["item"].name.startswith("Газ сжиженный WERKER")


def test_custom_item_opens_editor(gui, tmp_path):
    """Редактирование ячейки остаётся доступным отдельно от списка."""
    app, _, _ = gui
    _open_vedomost_and_convert(app, tmp_path)
    row_id, _stored = _row_with_doc(app, "Гидроизол")

    app.edit_cell(row_id)

    assert app._editing == row_id
    assert app._editor is not None


def test_own_text_is_kept_as_is(gui, tmp_path):
    """Своё название, которого нет в ведомости, остаётся свободным текстом."""
    app, _, _ = gui
    _open_vedomost_and_convert(app, tmp_path)
    row_id, stored = _row_with_doc(app, "Гидроизол")

    app.edit_cell(row_id)
    assert app._editing == row_id
    app._edit_var.set("Гидроизол ХКП S10м2 серый 3,5 мм (своё)")
    app.commit_edit()

    assert app._overrides[stored["key"]] == "Гидроизол ХКП S10м2 серый 3,5 мм (своё)"
    assert stored["item"].name == "Гидроизол ХКП S10м2 серый 3,5 мм (своё)"


def test_empty_typed_text_means_do_not_choose(gui, tmp_path):
    app, _, _ = gui
    _open_vedomost_and_convert(app, tmp_path)
    row_id, stored = _row_with_doc(app, "Газ")

    app.edit_cell(row_id)
    app._edit_var.set("")
    app.commit_edit()

    assert app._overrides[stored["key"]] is None
    assert _values(app, row_id)["Из ведомости"] == ""


def test_edited_name_goes_to_the_saved_file(gui, tmp_path):
    """Правка в таблице попадает в сохранённый файл."""
    import openpyxl

    app, _, _ = gui
    out_dir = tmp_path / "Итог"
    app.out_var.set(str(out_dir))
    _open_vedomost_and_convert(app, tmp_path)
    row_id, _stored = _row_with_doc(app, "Гидроизол")

    app.edit_cell(row_id)
    app._edit_var.set("Гидроизол ХКП 3,5 мм (проверено)")
    app.commit_edit()

    app.on_save()
    _run(app)

    written = app._last_outputs[0]
    names = [
        row[2]
        for row in openpyxl.load_workbook(written)["Данные"].iter_rows(min_row=2, values_only=True)
    ]
    assert "Гидроизол ХКП 3,5 мм (проверено)" in names


def test_scroll_closes_list_and_editor(gui, tmp_path):
    """Прокрутка закрывает список и поле ввода — окно не «висит» на месте."""
    app, _, _ = gui
    _open_vedomost_and_convert(app, tmp_path)
    row_id, _stored = _row_with_doc(app, "Газ")

    _click_cell(app, row_id, "▼")
    assert app._chooser is not None
    app._on_preview_wheel(_Event(10, 10))
    assert app._chooser is None

    app.edit_cell(row_id)
    assert app._editor is not None
    app._preview_yview("moveto", 0.5)  # прокрутка полосой прокрутки
    assert app._editor is None and app._editing is None


def test_popup_closes_when_row_scrolled_away(gui, tmp_path):
    """Сторож закрывает список, если строка «уехала» при прокрутке."""
    app, _, _ = gui
    _open_vedomost_and_convert(app, tmp_path)
    row_id, _stored = _row_with_doc(app, "Газ")

    _click_cell(app, row_id, "▼")
    assert app._chooser is not None
    # строка ушла из видимой области — bbox пустой, как при прокрутке
    app.preview.bbox = lambda *a, **k: ""
    app._check_popup()
    assert app._chooser is None
    assert app._popup_row is None


def test_chooser_requires_vedomost(gui):
    """Без ведомости выбирать нечего — список не открывается."""
    app, _, _ = gui
    app.on_convert()
    _run(app)
    row_id = next(iter(app._preview_rows))
    _click_cell(app, row_id, "▼")
    assert app._chooser is None


# ------------------------------------------------------ узкие плашки шагов
def test_step_cards_use_compact_buttons(gui):
    """Плашки ШАГ 1 и ШАГ 2 — минимальной высоты: мелкий шрифт и узкие кнопки."""
    app, _, _ = gui
    compact = [
        record for record in fake_tk.RECORDS if record.get("style") == "Compact.TButton"
    ]
    texts = [str(record.get("text", "")) for record in fake_tk.RECORDS]
    assert "Выбрать файл…" in texts
    assert any("Загрузить ведомость" in text for text in texts)
    assert "Папка сохранения…" in texts
    # кнопки сгруппированы в выпадающие меню на синем заголовке
    header_menus = [
        record for record in fake_tk.RECORDS if record.get("style") == "Header.TMenubutton"
    ]
    assert len(header_menus) >= 3
    assert {record.get("text") for record in header_menus} >= {
        "Файл и ведомость ▾", "Конвертировать ▾", "Настройки ▾"
    }
    header_buttons = [r for r in fake_tk.RECORDS if r.get("style") == "Header.TButton"]
    assert {r.get("text") for r in header_buttons} >= {"Сохранить", "Очистить таблицу"}


# -------------------------------------------- склад и поставщик (диалог 1.7.0)
def _dialog(app):
    dialog = getattr(app, "_source_dialog", None)
    assert dialog is not None, "диалог склада/поставщика не открылся"
    return dialog


def test_file_choice_asks_sklad_and_supplier(gui, tmp_path):
    """При выборе файла спрашиваем склад и поставщика из ведомости."""
    app, module, dialogs = gui
    app.load_vedomost(_write_vedomost(tmp_path))
    path = str(next(iter(known_sample_paths())))
    dialogs.open_result = (path,)
    app.on_add_files()
    dialog = _dialog(app)
    # склады заказчика — наглядные кнопки
    assert list(dialog.sklad_buttons) == ["Мир", "Кореличи"]
    # в списке поставщиков — все поставщики ведомости по алфавиту (без «как в документе»)
    values = list(dialog.supplier_checks)
    assert 'ООО "ОМА"' in values
    assert values == sorted(values, key=str.casefold)
    assert not any(v.startswith("— как в документе") for v in values)

    # выбор склада через наглядную кнопку
    dialog._select_sklad("Кореличи")
    dialog.supplier_vars['ООО "ОМА"'].set(True)
    dialog.on_ok()
    assert app._sklad == "Кореличи"
    assert app._supplier_choice == 'ООО "ОМА"'
    assert app.source_var.get() == 'Кореличи / ООО "ОМА"'
    assert not dialog.supplier_vars['ООО "ОМА"'].get()
    assert not hasattr(dialog, "clear_supplier_btn")
    app.ask_source()
    assert not app._source_dialog.supplier_vars['ООО "ОМА"'].get()
    assert app._supplier_choice == 'ООО "ОМА"'


def test_sklad_goes_to_folder_and_matcher(gui, tmp_path):
    """Выбранный склад задаёт папку результата и область поиска в ведомости."""
    app, _, _ = gui
    out_dir = tmp_path / "Итоги"
    app.out_var.set(str(out_dir))
    app.load_vedomost(_write_vedomost(tmp_path))
    app._source_chosen("Мир", 'ООО "ОМА"')
    app._add_files([str(next(iter(known_sample_paths())))])
    app.on_convert(MATCH_NAME)
    _run(app)

    assert app._matcher is not None
    assert app._matcher.forced_supplier == 'ООО "ОМА"'
    planned = app._pending_result.planned_output
    import datetime as dt

    today = f"{dt.date.today():%d.%m.%Y}"
    assert planned == str(out_dir / "Мир" / f"Мир ООО ОМА {today}.xlsx"), planned

    app.on_save()
    _run(app)
    assert app._last_outputs and app._last_outputs[0] == planned
    import os

    assert os.path.exists(planned)


def test_source_line_is_visible_in_window(gui):
    """Пункт «Склад и поставщик…» есть в окне и его можно изменить."""
    app, _, _ = gui
    assert "Склад и поставщик…" in fake_tk.TEXTS
    app._source_chosen("Кореличи", "")
    assert app.source_var.get() == "Кореличи"


def known_sample_paths():
    from pathlib import Path

    root = Path(__file__).resolve().parent.parent
    samples = [
        path
        for path in root.iterdir()
        if path.suffix.lower() in {".xls", ".xlsx"} and path.name.startswith("Кореличи")
    ]
    if not samples:
        import pytest

        pytest.skip("нет файла-образца")
    return samples


def test_chooser_shows_codes_before_long_name(gui, tmp_path):
    import openpyxl
    app, _, _ = gui
    path = _write_vedomost(tmp_path)
    book = openpyxl.load_workbook(path)
    sheet = book.active
    sheet.cell(4, 9, "Штрихкод")
    sheet.cell(5, 3, "ГАЗ-001")
    sheet.cell(5, 9, "0481234567890")
    book.save(path)
    assert app.load_vedomost(path)
    app.on_convert(MATCH_NAME)
    _run(app)
    row_id, stored = _row_with_doc(app, "Газ")
    _click_cell(app, row_id, "▼")
    matches = [(i, label, row) for i, (label, row) in enumerate(app._chooser.items)
               if getattr(row, "article", "") == "ГАЗ-001"]
    assert len(matches) == 1
    index, label, row = matches[0]
    assert label.index("арт. ГАЗ-001") < label.index(row.nomenclature)
    assert label.index("штрихкод 0481234567890") < label.index(row.nomenclature)
    _choose(app, index)
    assert app._overrides[stored["key"]].barcode == "0481234567890"
    assert app._overrides[stored["key"]].article == "ГАЗ-001"


def test_chooser_missing_codes_are_explicit(gui, tmp_path):
    app, _, _ = gui
    _open_vedomost_and_convert(app, tmp_path)
    row_id, _ = _row_with_doc(app, "Газ")
    _click_cell(app, row_id, "▼")
    choices = [label for label, row in app._chooser.items
               if getattr(row, "nomenclature", "").startswith("Газ")]
    assert choices and "арт. —" in choices[0] and "штрихкод —" in choices[0]


def test_chooser_uses_codes_from_real_vedomost(gui):
    from pathlib import Path
    path = Path(__file__).resolve().parent.parent / "Ведомость 16.11.2026 (ОК).xlsx"
    if not path.exists():
        pytest.skip("нет примера ведомости")
    app, _, _ = gui
    assert app.load_vedomost(str(path))
    app._source_chosen("Мир", 'ООО"Ома"')
    app.on_convert(MATCH_NAME)
    _run(app)
    row_id, _ = _row_with_doc(app, "Газ")
    _click_cell(app, row_id, "▼")
    labels = [label for label, row in app._chooser.items if getattr(row, "article", "") == "0356659"]
    assert labels and "штрихкод 4813426172277" in labels[0]
    assert "арт. 0356659" in labels[0]


def test_exact_barcode_with_low_name_similarity_is_pink(gui, monkeypatch):
    from types import SimpleNamespace
    from konvektor.extract import Item
    from konvektor.vedomost import MatchEntry, VedRow, MATCH_BARCODE
    app, module, _ = gui
    monkeypatch.setattr(module.app_settings, "load_settings", lambda: {"name_threshold": 90})
    key = ("file.xls", "sheet", 1)
    ved = VedRow(nomenclature="Эмаль белая 1 кг", barcode="4812355010858")
    entry = MatchEntry(key=key, row=1, doc_name="Грунтовка зелёная 5 кг",
                       doc_barcode="4812355010858", ved_row=ved,
                       matched=True, score=1.0, how=MATCH_BARCODE)
    app._matcher = SimpleNamespace(by_key={key: entry})
    stored = {"key": key, "item": Item(source=key[0], sheet=key[1], row=1)}
    app._preview_rows["row1"] = stored
    assert entry.score_title == "ОК"
    assert app._is_mismatch(stored)
    monkeypatch.setattr(module.app_settings, "load_settings", lambda: {"name_threshold": 0})
    assert not app._is_mismatch(stored)


def test_new_items_are_not_pink(gui):
    from types import SimpleNamespace
    app, _, _ = gui
    key = ("file.xls", "sheet", 1)
    entry = SimpleNamespace(skipped=False, matched=False, ved_row=None)
    app._matcher = SimpleNamespace(by_key={key: entry})
    stored = {"key": key}
    app._preview_rows["row1"] = stored
    assert not app._is_mismatch(stored)
    app._retag_row("row1", stored)
    assert app.preview.row_tags["row1"] == ("even",)


def test_conversion_completes_without_info_popup(gui):
    app, _, dialogs = gui
    dialogs.calls.clear()
    app.on_convert()
    _run(app)
    assert app._pending_result.items_count > 0
    assert app.save_btn.cget("state") == "normal"
    assert not any(call[0] == "showinfo" for call in dialogs.calls)


def test_header_displays_selected_sklad_and_suppliers(gui):
    app, _, _ = gui
    assert app.source_header.cget("textvariable") is app.source_var
    assert app.source_header.cget("style") == "HeaderSource.TLabel"
    app._source_chosen("Мир", 'ООО "ОМА", ООО Альфа')
    assert app.source_var.get() == 'Мир / ООО "ОМА", ООО Альфа'
    app._source_chosen("Кореличи", "")
    assert app.source_var.get() == "Кореличи"


def test_settings_menu_does_not_offer_save_settings(gui):
    app, _, _ = gui
    menu = app.table_mb.options["menu"]
    assert not any(item.get("label") == "Сохранить настройки" for item in menu.items)
    assert any(item.get("label") == "Сохранить поля таблицы" for item in menu.items)


def test_header_shows_document_supplier_without_placeholder(gui):
    app, _, _ = gui
    app._source_chosen("Мир", "")
    assert app.source_var.get() == "Мир"
    app.on_convert()
    _run(app)
    assert app.source_var.get() == 'Мир / ООО "ОМА"'
    app._source_chosen("Мир", "Другой поставщик")
    assert app.source_var.get() == "Мир / Другой поставщик"
    app._source_chosen("Мир", "")
    assert app.source_var.get() == 'Мир / ООО "ОМА"'
    app.on_clear_table()
    assert app.source_var.get() == "Мир"


def test_barcode_highlight_only_when_document_code_is_in_vedomost(gui):
    from types import SimpleNamespace
    from konvektor.extract import Item
    from konvektor.vedomost import Vedomost, VedRow, MatchEntry
    app, _, _ = gui
    code = "4812355010858"
    app._vedomost = Vedomost(rows=[VedRow(nomenclature="Товар", barcode=code)]).build_index()
    keys = [("x.xls", "s", i) for i in range(1, 4)]
    entries = {keys[0]: MatchEntry(key=keys[0], doc_barcode=code),
               keys[1]: MatchEntry(key=keys[1], doc_barcode="4812355010889"),
               keys[2]: MatchEntry(key=keys[2], doc_barcode="")}
    app._matcher = SimpleNamespace(by_key=entries)
    for key, displayed in zip(keys, [code, "4812355010889", code]):
        item = Item(source=key[0], sheet=key[1], row=key[2], barcode=displayed)
        row = app.preview.insert("", "end", values=[displayed])
        app._preview_rows[row] = {"key": key, "item": item}
    assert [app._barcode_in_vedomost(app._preview_rows[r]) for r in app._preview_rows] == [True, False, False]
    app._refresh_barcode_highlights()
    assert len(app._barcode_highlights) == 1
    label = next(iter(app._barcode_highlights.values()))
    assert label.cget("background") == "#EDFFD9"
    app._clear_preview()
    assert not app._barcode_highlights


def test_update_requires_confirmation_and_launches_selected_exe(gui, monkeypatch, tmp_path):
    app, module, dialogs = gui
    installer = tmp_path / "обновление setup.exe"
    installer.write_bytes(b"test")
    monkeypatch.setattr(module.sys, "platform", "win32")
    dialogs.open_result = (str(installer),)
    calls = []
    monkeypatch.setattr(module.subprocess, "Popen", lambda args, **kwargs: calls.append((args, kwargs)))
    monkeypatch.setattr(app, "_on_close", lambda: calls.append("closed"))
    monkeypatch.setattr(module.messagebox, "askyesno", lambda *a, **k: False)
    app.on_update_from_file()
    assert not calls
    monkeypatch.setattr(module.messagebox, "askyesno", lambda *a, **k: True)
    app.on_update_from_file()
    assert calls == [([str(installer)], {"shell": False}), "closed"]


def test_update_invalid_file_and_launch_failure_do_not_close(gui, monkeypatch, tmp_path):
    app, module, dialogs = gui
    monkeypatch.setattr(module.sys, "platform", "win32")
    closed = []
    monkeypatch.setattr(app, "_on_close", lambda: closed.append(True))
    monkeypatch.setattr(module.messagebox, "askyesno", lambda *a, **k: True)
    dialogs.open_result = (str(tmp_path / "missing.exe"),)
    app.on_update_from_file()
    installer = tmp_path / "setup.exe"
    installer.write_bytes(b"test")
    dialogs.open_result = (str(installer),)
    def fail(*args, **kwargs):
        raise OSError("ошибка")
    monkeypatch.setattr(module.subprocess, "Popen", fail)
    app.on_update_from_file()
    assert not closed
    assert sum(name == "showerror" for name, _, _ in dialogs.calls) >= 2


def test_update_not_started_during_conversion(gui, monkeypatch):
    app, module, dialogs = gui
    monkeypatch.setattr(module.sys, "platform", "win32")
    app._running = True
    app.on_update_from_file()
    assert not any(name == "askopenfilename" for name, _, _ in dialogs.calls)
