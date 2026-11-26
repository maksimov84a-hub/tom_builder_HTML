# -*- coding: utf-8 -*-
"""Иконка приложения и его идентификатор для панели задач Windows.

Файл иконки — `app/assets/app.ico` (несколько размеров: от 16 до 256 px),
его рисует `tmp/make_icon.py`. Здесь — только пути и небольшая обвязка:

* `app_icon()` — QIcon для окна и диалогов;
* `set_app_user_model_id()` — до создания окна сообщает Windows, что у
  программы свой AppUserModelID. Без этого задача группируется с
  `python.exe`/`pythonw.exe` и в панели задач показывается иконка Python,
  а не наша.
"""
from __future__ import annotations

import os
import sys

#: Идентификатор приложения в панели задач (формат «Компания.Продукт.Версия»).
APP_USER_MODEL_ID = "TomBuilder.VolumeBuilder.1"

_ICON_CACHE = None


def assets_dir() -> str:
    """Каталог ресурсов приложения (app/assets)."""
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "assets")


def icon_path() -> str:
    """Путь к файлу иконки (пустая строка, если файла нет)."""
    for name in ("app.ico", "app.png"):
        path = os.path.join(assets_dir(), name)
        if os.path.exists(path):
            return path
    return ""


def app_icon():
    """QIcon приложения (None, если файла иконки нет).

    Импорт PySide6 сделан внутри: модуль используют и сценарии без GUI.
    """
    global _ICON_CACHE
    if _ICON_CACHE is not None:
        return _ICON_CACHE
    path = icon_path()
    if not path:
        return None
    try:
        from PySide6.QtGui import QIcon
        _ICON_CACHE = QIcon(path)
    except Exception:
        return None
    return _ICON_CACHE


def set_app_user_model_id(app_id: str | None = None) -> None:
    """Задаёт AppUserModelID процесса — от него зависит иконка в панели задач.

    Вызывать ДО создания главного окна. На не-Windows и при недоступном
    shell32 тихо ничего не делает.
    """
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            ctypes.c_wchar_p(app_id or APP_USER_MODEL_ID))
    except Exception:
        # Идентификатор — косметика для панели задач: без него программа
        # работает, иконка окна всё равно берётся из QIcon.
        pass
