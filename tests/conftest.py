"""Общие фикстуры: образцы счетов лежат в корне репозитория."""

from __future__ import annotations

from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]

#: Файлы-образцы, на которых настраивался парсер (данные заказчика).
KNOWN_FILES = {
    "00000003370Приложение2Листа.xls": 70,
    "00000003370ЭТТН.xls": 70,
    "000004781#0066065.xls": 31,
    "_№ 10802 от 12 ноября 2026 г .XLS": 29,
    "ИТ0685355.xlsx": 25,
    "КХ2104033 (1).xls": 19,
    "Кореличи (4).xls": 11,
    "Счет СЧЦБ-17059130.xlsx": 6,
    "Счет на оплату № УТ-542 от 12.11.2026.xlsx": 6,
    "ТТН-1 (вертикальная) № 1019250 от 12.11.2026.xls": 34,
}

_EXCLUDE_MARKERS = ("Загрузка_", "konvektor", "selftest", "~$")


def _sample_paths():
    out = []
    for p in sorted(ROOT.iterdir()):
        if not p.is_file():
            continue
        if p.suffix.lower() not in (".xls", ".xlsx", ".xlsm"):
            continue
        if any(m in p.name for m in _EXCLUDE_MARKERS):
            continue
        out.append(p)
    return out


@pytest.fixture(scope="session")
def sample_paths():
    paths = _sample_paths()
    if not paths:
        pytest.skip("в корне репозитория нет файлов-образцов")
    return paths


@pytest.fixture(scope="session")
def known_paths(sample_paths):
    """Только те образцы, для которых зафиксированы ожидаемые количества."""
    return [p for p in sample_paths if p.name in KNOWN_FILES]


@pytest.fixture(scope="session")
def sample_dir():
    return str(ROOT)
