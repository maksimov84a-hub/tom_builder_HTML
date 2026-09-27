# -*- coding: utf-8 -*-
"""Инициализация pywin32 (vendored): настройка путей и DLL-каталогов.

Реплицирует содержимое pywin32.pth, которое при установке через pip
выполняется автоматически. Поскольку зависимости лежат в ./vendor,
эту инициализацию делаем вручную.
"""
from __future__ import annotations

import os
import sys

_initialized = False


def ensure_pywin32() -> None:
    """Добавляет пути pywin32 в sys.path и регистрирует DLL-каталог
    (pywin32_system32). Выполняется один раз."""
    global _initialized
    if _initialized:
        return

    # Каталог vendor (где лежат пакеты)
    app_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    vendor = os.path.join(app_root, "vendor")
    if vendor not in sys.path:
        sys.path.insert(0, vendor)

    # Пути pywin32 (как в pywin32.pth)
    for rel in ("win32", os.path.join("win32", "lib"), "pythonwin"):
        p = os.path.join(vendor, rel)
        if p not in sys.path:
            sys.path.append(p)

    # Регистрация DLL-каталога
    try:
        import pywin32_system32
        for p in getattr(pywin32_system32, "__path__", []):
            if os.path.isdir(p):
                os.add_dll_directory(p)
                break
    except Exception:
        pass

    _initialized = True
