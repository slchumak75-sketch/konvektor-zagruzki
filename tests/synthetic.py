"""Синтетические раскладки таблиц — корпус для проверки общности разбора.

Идея: один и тот же логический документ (3 позиции, 10 полей) раскладывается
по листу десятком разных способов — как это делают разные поставщики. Разбор
обязан вытащить одни и те же данные из **любой** раскладки. Это защита от
«подгонки под конкретные файлы»: тесты не привязаны к образцам заказчика и
проверяют именно общий алгоритм.

Каждая функция возвращает матрицу ячеек (`List[List[Any]]`), как её отдаёт
`readers.Sheet.rows`.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from konvektor.extract import SheetExtraction, extract_file
from konvektor.readers import Sheet, SourceFile

#: Эталонные позиции: (штрихкод, артикул, наименование, ед., кол-во, цена,
#: сумма без НДС, ставка НДС, сумма НДС, сумма с НДС).
CANON = [
    ("4607011539220", "VX19100", "Эмаль металлик универсальная 520 мл", "шт", 6, 10.73, 64.38, 20, 12.88, 77.26),
    ("4605544048424", "GR9003", "Грунт-эмаль белая RAL 9003 0,9 кг", "шт", 6, 15.37, 92.22, 20, 18.44, 110.66),
    ("4605544048707", "GR3005", "Грунт-эмаль вишнёвая RAL 3005 0,9 кг", "шт", 2, 14.11, 28.22, 20, 5.64, 33.86),
]

#: Заголовки по умолчанию (порядок = порядок полей в CANON).
TITLES = [
    "Штрихкод",
    "Артикул",
    "Наименование товара",
    "Ед. изм.",
    "Количество",
    "Цена",
    "Сумма",
    "Ставка НДС",
    "Сумма НДС",
    "Сумма с НДС",
]

TOTAL_SUM = round(sum(r[6] for r in CANON), 2)      # 184.82
TOTAL_VAT = round(sum(r[8] for r in CANON), 2)      # 36.96
TOTAL_WITH_VAT = round(sum(r[9] for r in CANON), 2)  # 221.78


def extract(grid: List[List[Any]], name: str = "grid.xlsx") -> SheetExtraction:
    """Разобрать матрицу так же, как настоящий файл."""
    source = SourceFile(path=name, sheets=[Sheet(name="Лист1", rows=grid)], reader="grid")
    return extract_file(source)


def _data_row(rec, cols: Dict[int, Any]) -> List[Any]:
    """Строка данных: значения по номерам колонок."""
    width = max(cols) + 1
    row: List[Any] = [None] * width
    for col, value in cols.items():
        row[col] = value
    return row


def _table(cols: Optional[Dict[int, Any]] = None, titles: Optional[List[str]] = None,
           rows: Optional[List[List[Any]]] = None) -> List[List[Any]]:
    """Простая таблица: шапка + строки."""
    titles = titles or TITLES
    nums = list(range(len(titles)))
    grid: List[List[Any]] = [list(titles)]
    if rows is not None:
        grid.extend(rows)
        return grid
    for i, rec in enumerate(CANON):
        grid.append([rec[nums[j]] for j in range(len(titles))] if False else list(rec))
    return grid


# --------------------------------------------------------------------------- #
#  Варианты раскладок
# --------------------------------------------------------------------------- #
def v_classic() -> List[List[Any]]:
    """Шапка в первой строке, данные сразу под ней."""
    return _table()


def v_header_deep() -> List[List[Any]]:
    """Шапка на 7-й строке, выше — реквизиты документа."""
    top: List[List[Any]] = [
        ["Счет на оплату № 777 от 1 февраля 2026 г."],
        [],
        ["Поставщик:", 'ООО "Ромашка"', "", "УНП 100000001"],
        ["Покупатель:", 'ООО "Лютик"', "", "УНП 200000002"],
        [],
        [],
    ]
    return top + _table()


def v_shuffled() -> List[List[Any]]:
    """Колонки переставлены: цена раньше количества, артикул в конце."""
    order = [2, 4, 5, 6, 7, 8, 9, 3, 1, 0]  # индексы полей CANON
    titles = [TITLES[i] for i in order]
    rows = [[rec[i] for i in order] for rec in CANON]
    return [titles] + rows


def v_junk_columns() -> List[List[Any]]:
    """Между нужными колонками — лишние (страна, ТНВЭД, вес, примечание)."""
    grid = [[
        "№", "Наименование", "Страна", "Код ТНВЭД", "Кол-во", "Ед. изм.", "Вес",
        "Цена", "Примечание", "Артикул", "Сумма", "Ставка НДС", "Сумма НДС",
        "Всего с НДС", "Штрихкод",
    ]]
    for i, rec in enumerate(CANON, start=1):
        grid.append([
            i, rec[2], "РОССИЯ", "3208909109", rec[4], rec[3], 0.9,
            rec[5], "", rec[1], rec[6], f"{rec[7]}%", rec[8], rec[9], int(rec[0]),
        ])
    return grid


def v_multirow_header() -> List[List[Any]]:
    """Двухуровневая шапка: часть подписей во второй строке (как в 1С)."""
    return [
        ["Штрихкод", "Артикул", "Наименование", "Ед.", "Кол-во", "Цена", "Сумма", "НДС", "Сумма", "Всего"],
        ["", "", "", "изм.", "", "за единицу", "без НДС", "ставка", "НДС", "с НДС"],
    ] + [list(rec) for rec in CANON]


def v_two_tables() -> List[List[Any]]:
    """Две таблицы на одном листе (как страницы ТТН)."""
    first = _table()
    second = [list(TITLES)] + [list(rec) for rec in CANON]
    return first + [[], ["Итого по странице:", "", "", "", "", "", TOTAL_SUM]] + second


def v_totals_above() -> List[List[Any]]:
    """Итоги сверху, таблица ниже."""
    return [
        ["Итого с НДС:", "", "", "", "", "", "", "", "", TOTAL_WITH_VAT],
        [],
    ] + _table()


def v_synonyms() -> List[List[Any]]:
    """Другие подписи тех же колонок."""
    titles = [
        "Штрих-код",
        "Код",
        "Товары (работы, услуги)",
        "Ед.",
        "К-во",
        "Цена, руб. (BYN)",
        "Стоимость",
        "Ставка НДС, %",
        "В том числе НДС",
        "Всего к оплате",
    ]
    return [titles] + [list(rec) for rec in CANON]


def v_gaps() -> List[List[Any]]:
    """Пустые строки и колонки-разделители внутри таблицы."""
    grid: List[List[Any]] = [[None, None] + list(TITLES)]
    for rec in CANON:
        grid.append([None, None] + list(rec))
        grid.append([])
    return grid


def v_no_barcode_header() -> List[List[Any]]:
    """У колонки штрихкода нет подписи — только 13-значные числа."""
    titles = list(TITLES)
    titles[0] = ""
    return [titles] + [list(rec) for rec in CANON]


def v_transposed() -> List[List[Any]]:
    """Таблица «на бок»: роли по строкам, позиции по колонкам."""
    rows: List[List[Any]] = []
    for field, title in enumerate(TITLES):
        row: List[Any] = [title]
        for rec in CANON:
            row.append(rec[field])
        rows.append(row)
    return rows


def v_transposed_with_meta() -> List[List[Any]]:
    """«Боковая» таблица под шапкой документа."""
    return [
        ["Поставщик:", 'ООО "Ромашка"'],
        ["Счет № 777 от 1 февраля 2026 г."],
        [],
    ] + v_transposed()


def v_no_totals_no_meta() -> List[List[Any]]:
    """Голая таблица без итогов и реквизитов."""
    return _table()


VARIANTS = {
    "classic": v_classic,
    "header_deep": v_header_deep,
    "shuffled_columns": v_shuffled,
    "junk_columns": v_junk_columns,
    "multirow_header": v_multirow_header,
    "two_tables": v_two_tables,
    "totals_above": v_totals_above,
    "synonyms": v_synonyms,
    "gaps": v_gaps,
    "no_barcode_header": v_no_barcode_header,
    "transposed": v_transposed,
    "transposed_with_meta": v_transposed_with_meta,
    "bare_table": v_no_totals_no_meta,
}

#: Раскладки, которые дают удвоенный набор позиций (две таблицы).
DOUBLED = {"two_tables"}
