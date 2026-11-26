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

    Флаг `--single-process` здесь НЕ используется: в нём Chromium держит
    сетевой поток и очередь ввода в одном процессе с GUI, и после потери
    фокуса окном (переключение в другую программу и обратно) очередь ввода
    перестаёт обрабатываться — кнопки в окне перестают нажиматься. Без него
    QtWebEngine запускает вспомогательные процессы как обычно.
    """
    # Каталог для данных WebEngine (не системный temp, чтобы избежать «Отказано в доступе»)
    root = os.path.dirname(os.path.abspath(__file__))
    user_data = os.path.join(root, "tmp", "qt_webdata")
    os.makedirs(user_data, exist_ok=True)
    os.environ.setdefault("QTWEBENGINE_USER_DATA_DIR", user_data)

    flags = os.environ.get("QTWEBENGINE_CHROMIUM_FLAGS", "")
    add = ("--no-sandbox --disable-gpu --disable-software-rasterizer "
           "--disable-dev-shm-usage")
    # Не перезаписываем уже заданные пользователем флаги
    if flags.strip():
        return
    os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = add
    os.environ.setdefault("QTWEBENGINE_DISABLE_SANDBOX", "1")


def setup_streams() -> None:
    """Подстраховка запуска без консоли (pythonw.exe из run.bat).

    У pythonw.exe нет стандартных потоков: `sys.stdout` и `sys.stderr` равны
    None, и первая же запись (print, предупреждение QtWebEngine) привела бы к
    ошибке в C-коде или исключению. Перенаправляем их в tmp/build.log — файл
    заодно служит журналом запуска.
    """
    if sys.stdout is not None and sys.stderr is not None:
        return
    root = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(root, "tmp", "build.log")
    stream = None
    try:
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        stream = open(log_path, "a", encoding="utf-8", buffering=1)
    except OSError:
        try:
            stream = open(os.devnull, "w", encoding="utf-8")
        except OSError:
            return
    if sys.stdout is None:
        sys.stdout = stream
    if sys.stderr is None:
        sys.stderr = stream


def main() -> int:
    ensure_vendor()
    setup_webengine_env()
    setup_streams()

    from PySide6.QtCore import QLockFile
    from PySide6.QtWidgets import QApplication, QMessageBox
    from app.ui.main_window import MainWindow
    from app.ui.app_icon import app_icon, set_app_user_model_id

    # Идентификатор до создания окна: иначе Windows группирует задачу с
    # python.exe и в панели задач показывает иконку Python, а не нашу.
    set_app_user_model_id()

    lock_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "tmp", "application.lock")
    lock = QLockFile(lock_path)
    if not lock.tryLock(0):
        app = QApplication(sys.argv)
        app.setApplicationName("Сборка тома по ГОСТ Р 21.101-2020")
        QMessageBox.information(None, app.applicationName(),
                                "Программа уже запущена.")
        return 1

    app = QApplication(sys.argv)
    app.setApplicationName("Сборка тома по ГОСТ Р 21.101-2020")
    # Иконка приложения: окно, панель задач и все диалоги.
    icon = app_icon()
    if icon is not None:
        app.setWindowIcon(icon)
    window = MainWindow()
    if icon is not None:
        window.setWindowIcon(icon)
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())

