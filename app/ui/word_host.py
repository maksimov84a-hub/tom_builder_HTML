# -*- coding: utf-8 -*-
"""Встроенное окно Microsoft Word для предпросмотра листов ПЗ.

Chromium (QWebEngine) не умеет держать нативное окно, поэтому лист Word
показывается отдельным нативным виджетом: окно документа Word переносится
внутрь виджета через ``SetParent``. Пользователь видит настоящий лист Word —
с лентой, курсором и правкой, а не картинку.

Почему так, а не OLE-контейнер: ``QAxWidget("Word.Document.8")`` создаёт
контейнер, но загрузить в него готовый файл вызовом из Python не удалось —
Word создаёт пустой документ «Без имени». Перенос окна работает устойчиво и
даёт тот же результат.

Правки остаются в РАБОЧЕЙ КОПИИ файла (см. core.word_preview): исходный DOCX
записки не меняется, пока пользователь не перенесёт правки явно.
"""
from __future__ import annotations

import os

from PySide6.QtCore import QTimer, Signal, Qt
from PySide6.QtWidgets import QLabel, QVBoxLayout, QWidget


class WordHost(QWidget):
    """Виджет, показывающий документ Word внутри окна программы.

    Сигналы:
      closed()          — окно Word закрылось (пользователь или сбой);
      modified_changed(bool) — появились/исчезли несохранённые правки.
    """

    closed = Signal()
    modified_changed = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_NativeWindow, True)
        self.setAttribute(Qt.WidgetAttribute.WA_DontCreateNativeAncestors, True)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        self._hint = QLabel(
            "Лист Word загружается…\n"
            "Правки сохраняются в рабочую копию и не меняют исходный DOCX.")
        self._hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._hint.setStyleSheet("color:#475467; background:#f9fafb;")
        lay.addWidget(self._hint)

        self._word = None          # Word.Application (COM)
        self._doc = None           # открытый документ
        self._hwnd = 0             # окно документа Word
        self._path = ""
        self._own_word = True
        self._poll = QTimer(self)
        self._poll.setInterval(700)
        self._poll.timeout.connect(self._check_state)

    # --- Управление ---

    def is_open(self) -> bool:
        """Открыт ли документ (Word жив и не закрыт пользователем)."""
        if self._doc is None or self._word is None:
            return False
        try:
            _ = self._doc.Name
            return True
        except Exception:
            return False

    def current_path(self) -> str:
        return self._path

    def open_document(self, path: str, *, read_only: bool = False) -> bool:
        """Открывает DOCX во встроенном Word. True — документ показан."""
        if not path or not os.path.exists(path):
            self._show_error(f"Файл не найден: {path}")
            return False
        if self._path == os.path.abspath(path) and self.is_open():
            self._embed()
            return True
        self.close_document()
        result = self._open_live(path, read_only)
        if result is None:
            # Файл держит другой экземпляр Word: открываем копию, чтобы вместо
            # диалога «Файл уже используется» пользователь увидел лист.
            copy_path = self._locked_copy(path)
            if copy_path:
                result = self._open_live(copy_path, read_only)
                if result is not None:
                    self._path = os.path.abspath(path)
        if result is None:
            return False
        self._word, self._doc, self._hwnd = result
        self._hint.hide()
        QTimer.singleShot(150, self._embed)
        self._poll.start()
        return True

    def _open_live(self, path: str, read_only: bool):
        """Открывает файл в новом экземпляре Word: (word, doc, hwnd) или None."""
        try:
            from ..core.word_com import ensure_pywin32
            ensure_pywin32()
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = True
            word.DisplayAlerts = 0
            doc = word.Documents.Open(
                os.path.abspath(path), ReadOnly=bool(read_only),
                AddToRecentFiles=False, Visible=True)
            try:
                doc.ActiveWindow.WindowState = 0      # wdWindowStateNormal
            except Exception:
                pass
        except Exception as e:  # нет Word, сбой COM — сообщаем на панели
            self._show_error(f"Не удалось открыть Word: {type(e).__name__}: {e}")
            return None
        self._path = os.path.abspath(path)
        return word, doc, 0

    @staticmethod
    def _locked_copy(path: str) -> str:
        """Копия файла во временном каталоге для случая «файл занят»."""
        import shutil
        import tempfile
        try:
            folder = tempfile.mkdtemp(prefix="pz_preview_")
            target = os.path.join(folder, os.path.basename(path))
            shutil.copy2(path, target)
            return target
        except OSError:
            return ""

    def save(self) -> bool:
        """Сохраняет правки в рабочую копию. True — сохранено."""
        if not self.is_open():
            return False
        try:
            self._doc.Save()
            return True
        except Exception:
            return False

    def close_document(self, save: bool = True) -> None:
        """Закрывает документ и освобождает Word.

        :param save: True — сохранить правки в рабочую копию. False — закрыть
            без сохранения (например, при выходе из программы).
        """
        self._poll.stop()
        doc, word = self._doc, self._word
        self._doc = None
        self._word = None
        self._hwnd = 0
        self._path = ""
        if doc is not None:
            try:
                doc.Close(SaveChanges=-1 if save else 0)   # wdSaveChanges / wdDoNotSave
            except Exception:
                pass
        if word is not None and self._own_word:
            try:
                word.Quit(SaveChanges=-1 if save else 0)
            except Exception:
                try:
                    word.Quit()
                except Exception:
                    pass

    def reload(self, path: str) -> bool:
        """Переоткрывает файл (после пересборки рабочей копии)."""
        self.close_document()
        return self.open_document(path)

    # --- Встраивание ---

    def _embed(self) -> None:
        """Переносит окно документа Word внутрь этого виджета."""
        if not self.is_open():
            return
        try:
            from ..core.word_com import ensure_pywin32
            ensure_pywin32()
            import win32gui
            import win32con
        except Exception as e:
            self._show_error(f"Нет доступа к Win32: {type(e).__name__}: {e}")
            return
        try:
            hwnd = int(self._doc.ActiveWindow.Hwnd)
        except Exception:
            # Окно ещё не создано (документ открывается) — повторим.
            QTimer.singleShot(200, self._embed)
            return
        if not hwnd:
            QTimer.singleShot(200, self._embed)
            return
        host = int(self.winId())
        try:
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            win32gui.SetWindowLong(
                hwnd, win32con.GWL_STYLE,
                (style | win32con.WS_CHILD) & ~win32con.WS_POPUP)
            win32gui.SetParent(hwnd, host)
        except Exception as e:
            self._show_error(f"Не удалось встроить окно Word: {type(e).__name__}: {e}")
            return
        self._hwnd = hwnd
        self._resize_embedded()

    def _resize_embedded(self) -> None:
        """Подгоняет окно Word под размер виджета."""
        if not self._hwnd:
            return
        try:
            import win32gui
            win32gui.MoveWindow(self._hwnd, 0, 0, max(1, self.width()),
                                max(1, self.height()), True)
        except Exception:
            pass

    def resizeEvent(self, event) -> None:      # noqa: N802 (Qt API)
        super().resizeEvent(event)
        self._resize_embedded()

    def showEvent(self, event) -> None:        # noqa: N802 (Qt API)
        super().showEvent(event)
        self._resize_embedded()

    # --- Состояние ---

    def _check_state(self) -> None:
        """Следит за закрытием Word, правками и целостностью встраивания."""
        if self._doc is None:
            self._poll.stop()
            return
        try:
            self._doc.Name
        except Exception:
            self._poll.stop()
            self._doc = None
            self._word = None
            self._hwnd = 0
            self.closed.emit()
            return
        self._keep_embedded()
        try:
            self.modified_changed.emit(not bool(self._doc.Saved))
        except Exception:
            pass

    def _keep_embedded(self) -> None:
        """Возвращает окно Word внутрь виджета, если Word его пересоздал.

        Word — отдельный процесс: при смене вида, показе диалога или открытии
        документа он может пересоздать окно, и встраивание теряется — окно
        уезжает на рабочий стол. Дешёвая проверка раз в такт возвращает его на
        место.
        """
        try:
            from ..core.word_com import ensure_pywin32
            ensure_pywin32()
            import win32gui
        except Exception:
            return
        try:
            hwnd = int(self._doc.ActiveWindow.Hwnd)
        except Exception:
            return
        if not hwnd:
            return
        try:
            parent = win32gui.GetParent(hwnd)
        except Exception:
            return
        if hwnd == self._hwnd and parent == int(self.winId()):
            return
        self._hwnd = 0
        self._embed()

    def _show_error(self, message: str) -> None:
        self._hint.setText(message)
        self._hint.show()

    def message(self) -> str:
        """Текст подсказки на панели (пусто — документ показан)."""
        return "" if self._hwnd else self._hint.text()

    def closeEvent(self, event) -> None:       # noqa: N802 (Qt API)
        self.close_document()
        super().closeEvent(event)
