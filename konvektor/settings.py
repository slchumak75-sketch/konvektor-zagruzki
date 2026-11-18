"""Сохранение настроек окна — поля таблицы, склад, папки, размеры.

Хранится в `%APPDATA%/Konvertor/settings.json` (Windows) или `~/.config/konvertor/settings.json`
(Linux). Если записать не удалось — молча пропускаем, работа программы не ломается.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict

APP_DIR_NAME = "Konvertor"
SETTINGS_FILE = "settings.json"
# старое имя для миграции с предыдущих версий (Конвектор)
LEGACY_APP_DIR_NAME = "Konvektor"

DEFAULTS: Dict[str, Any] = {
    "window_geometry": "",
    "output_dir": "",
    "last_vedomost": "",
    "sklad": "",
    "supplier": "",
    "column_widths": {},  # title -> width px
    "preview_columns_order": [],
    "overrides": {},  # vedomost_path -> {key -> value} (для будущих сессий)
    "table_fields": {},  # произвольные поля таблицы, сохраняемые пользователем
    "name_threshold": 60,  # процент совпадения по наименованию (икс) — если штрихкод/артикул совпал, а имя ниже икс — подсветить
}


def settings_dir() -> Path:
    """Папка для настроек."""
    appdata = os.environ.get("APPDATA")
    if appdata and os.path.isdir(appdata):
        return Path(appdata) / APP_DIR_NAME
    # Linux / fallback
    xdg = os.environ.get("XDG_CONFIG_HOME")
    if xdg:
        return Path(xdg) / "konvertor"
    return Path.home() / ".config" / "konvertor"


def _legacy_settings_dir() -> Path:
    """Старая папка настроек (Конвектор) для миграции."""
    appdata = os.environ.get("APPDATA")
    if appdata and os.path.isdir(appdata):
        return Path(appdata) / LEGACY_APP_DIR_NAME
    xdg = os.environ.get("XDG_CONFIG_HOME")
    if xdg:
        return Path(xdg) / "konvektor"
    return Path.home() / ".config" / "konvektor"


def settings_path() -> Path:
    return settings_dir() / SETTINGS_FILE


def load_settings() -> Dict[str, Any]:
    """Прочитать настройки, вернуть копию DEFAULTS + файл."""
    data = dict(DEFAULTS)
    path = settings_path()
    legacy = _legacy_settings_dir() / SETTINGS_FILE
    # миграция: если новой папки нет, а старая есть — читаем оттуда
    if not path.is_file() and legacy.is_file():
        path = legacy
    if not path.is_file():
        return data
    try:
        with open(path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        if isinstance(loaded, dict):
            data.update(loaded)
    except Exception:
        pass
    return data


def save_settings(data: Dict[str, Any]) -> None:
    """Записать настройки (тихо игнорируем ошибки)."""
    path = settings_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        # мержим с дефолтами, чтобы не потерять ключи
        to_save = dict(DEFAULTS)
        to_save.update(data)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(to_save, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def update_settings(**kwargs: Any) -> None:
    """Обновить отдельные ключи."""
    data = load_settings()
    data.update(kwargs)
    save_settings(data)
