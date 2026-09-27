# -*- coding: utf-8 -*-
"""Точка входа приложения «Сборка тома по ГОСТ Р 21.101-2020»."""
from __future__ import annotations

import os
import sys


def ensure_vendor() -> None:
    """Добавляет каталог vendor (зависимости) в sys.path, если библиотек нет в системе."""
    root = os.path.dirname(os.path.abspath(__file__))
    vendor = os.path.join(root, "vendor")
    if vendor not in sys.path:
        sys.path.insert(0, vendor)


def setup_webengine_env() -> None:
    """Настраивает окружение QtWebEngine (перед созданием QApplication).

    Флаги отключают GPU-подсистему и sandbox, чтобы QWebEngine корректно
    запускался в локальных средах (в т.ч. при запуске из-под ограничений).
    """
    # Каталог для данных WebEngine (не системный temp, чтобы избежать «Отказано в доступе»)
    root = os.path.dirname(os.path.abspath(__file__))
    user_data = os.path.join(root, "tmp", "qt_webdata")
    os.makedirs(user_data, exist_ok=True)
    os.environ.setdefault("QTWEBENGINE_USER_DATA_DIR", user_data)

    flags = os.environ.get("QTWEBENGINE_CHROMIUM_FLAGS", "")
    add = "--single-process --no-sandbox --disable-gpu --disable-software-rasterizer --disable-dev-shm-usage"
    # Не перезаписываем уже заданные пользователем флаги
    if flags.strip():
        return
    os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = add
    os.environ.setdefault("QTWEBENGINE_DISABLE_SANDBOX", "1")


def main() -> int:
    ensure_vendor()
    setup_webengine_env()

    from PySide6.QtWidgets import QApplication
    from app.ui.main_window import MainWindow

    app = QApplication(sys.argv)
    app.setApplicationName("Сборка тома по ГОСТ Р 21.101-2020")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())

