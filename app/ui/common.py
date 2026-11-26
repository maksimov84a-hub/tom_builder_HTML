# -*- coding: utf-8 -*-
"""Общие устройства правой панели.

Здесь собрано то, что раньше копировалось в каждой форме:
  • строка «поле с путём + кнопка выбора файла» (документ, приложение,
    логотип графы 9, шаблоны обложки и титула, подпись подписанта);
  • каркас «таблица + Добавить/Изменить/Удалить» (изменения, приложения,
    подписанты);
  • пересчёт числа листов исходного файла.
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton,
    QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox, QCheckBox,
)
from PySide6.QtCore import Signal


class FileRow(QWidget):
    """Путь к файлу (только для чтения) и кнопка выбора.

    Выбор файла не меняет модель сам по себе: форма подписывается на сигнал
    ``picked`` и решает, что сделать с новым путём.
    """

    picked = Signal(str)

    def __init__(self, dialog_title: str, filters: str, *,
                 button_text: str = "Выбрать файл…",
                 placeholder: str = "", clearable: bool = False, parent=None):
        super().__init__(parent)
        self._title = dialog_title
        self._filters = filters
        row = QHBoxLayout(self)
        row.setContentsMargins(0, 0, 0, 0)
        self.edit = QLineEdit()
        self.edit.setReadOnly(True)
        if placeholder:
            self.edit.setPlaceholderText(placeholder)
        self.btn = QPushButton(button_text)
        self.btn.clicked.connect(self._pick)
        row.addWidget(self.edit, stretch=1)
        row.addWidget(self.btn)
        self.btn_clear = None
        if clearable:
            self.btn_clear = QPushButton("Очистить")
            self.btn_clear.clicked.connect(self._clear)
            row.addWidget(self.btn_clear)

    def path(self) -> str:
        return self.edit.text().strip()

    def set_path(self, path: str) -> None:
        self.edit.setText(path or "")

    def set_filters(self, filters: str) -> None:
        """Меняет фильтр выбора (например, «только PDF» для графики)."""
        self._filters = filters

    def set_button_visible(self, visible: bool) -> None:
        self.btn.setVisible(visible)

    def _pick(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, self._title, self.path(), self._filters)
        if path:
            self.set_path(path)
            self.picked.emit(path)

    def _clear(self) -> None:
        self.set_path("")
        self.picked.emit("")


class TableCrudForm(QWidget):
    """Каркас реестра: таблица и кнопки «Добавить / Изменить / Удалить».

    Наследник задаёт заголовки колонок (``headers``), содержимое строки
    (``cells``) и диалог ввода (``ask_item``). Список хранится по ссылке —
    правки сразу видны в модели проекта.

    ``checkable = True`` добавляет первую колонку «В сборку» с флажком:
    он управляет признаком ``needed`` элемента (снятый флажок исключает
    элемент из сборки, не удаляя его). Наследник переопределяет
    ``is_checked`` / ``set_checked`` и получает уведомление ``items_changed``.
    """

    headers: tuple = ()
    delete_question = "Удалить запись?"
    checkable = False
    check_header = "В сборку"

    #: Изменился признак «в сборку» у элемента.
    items_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        v = QVBoxLayout(self)
        v.setContentsMargins(8, 8, 8, 8)
        ncols = len(self.headers) + (1 if self.checkable else 0)
        self.table = QTableWidget(0, ncols)
        labels = ([self.check_header] if self.checkable else []) + list(self.headers)
        self.table.setHorizontalHeaderLabels(labels)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        v.addWidget(self.table)

        row = QHBoxLayout()
        self.btn_add = QPushButton("Добавить")
        self.btn_edit = QPushButton("Изменить")
        self.btn_del = QPushButton("Удалить")
        self.btn_add.clicked.connect(self._add)
        self.btn_edit.clicked.connect(self._edit)
        self.btn_del.clicked.connect(self._del)
        row.addWidget(self.btn_add)
        row.addWidget(self.btn_edit)
        row.addWidget(self.btn_del)
        row.addStretch(1)
        v.addLayout(row)

        self._items: list = []

    # ---- переопределяется наследником ----
    def cells(self, item) -> list:
        """Тексты ячеек строки таблицы для элемента реестра."""
        return []

    def ask_item(self, item):
        """Диалог ввода. Возвращает элемент или None (отмена)."""
        return None

    def is_checked(self, item) -> bool:
        """Включён ли элемент в сборку (для checkable-таблиц)."""
        return bool(getattr(item, "needed", True))

    def set_checked(self, item, checked: bool) -> None:
        """Записывает признак «в сборку» в модель."""
        if hasattr(item, "needed"):
            item.needed = bool(checked)

    # ---- общее поведение ----
    def set_items(self, items: list) -> None:
        self._items = items if items is not None else []
        self._refresh()

    def get_items(self) -> list:
        return self._items

    def _refresh(self) -> None:
        self.table.setRowCount(0)
        offset = 1 if self.checkable else 0
        for item in self._items:
            r = self.table.rowCount()
            self.table.insertRow(r)
            if self.checkable:
                self.table.setCellWidget(r, 0, self._make_check(item))
            for c, text in enumerate(self.cells(item)):
                self.table.setItem(r, c + offset, QTableWidgetItem(text))
        if self.checkable:
            self.table.setColumnWidth(0, 64)
        self.table.resizeColumnsToContents()
        if self.checkable:
            self.table.setColumnWidth(0, 64)

    def _make_check(self, item):
        """Флажок «в сборку» в ячейке таблицы."""
        box = QCheckBox()
        box.setChecked(self.is_checked(item))
        box.setToolTip("Включать в сборку тома")
        box.setStyleSheet("margin-left:14px;")
        box.toggled.connect(lambda checked, it=item: self._on_check_toggled(it, checked))
        return box

    def _on_check_toggled(self, item, checked: bool) -> None:
        self.set_checked(item, checked)
        self.items_changed.emit()

    def _add(self) -> None:
        item = self.ask_item(None)
        if item is not None:
            self._items.append(item)
            self._refresh()
            self.items_changed.emit()

    def _edit(self) -> None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self._items):
            return
        item = self.ask_item(self._items[row])
        if item is not None:
            self._items[row] = item
            self._refresh()
            self.items_changed.emit()

    def _del(self) -> None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self._items):
            return
        if QMessageBox.question(self, "Удалить", self.delete_question) == QMessageBox.Yes:
            del self._items[row]
            self._refresh()
            self.items_changed.emit()


def refresh_document_meta(doc) -> None:
    """Пересчитывает листы (и размер листа) документа по его файлу."""
    path = getattr(doc, "path", "") or ""
    if not path:
        return
    try:
        from ..core.normalizer import normalize_file
        normalized = normalize_file(path, force_kind=getattr(doc, "doc_kind", None))
        doc.pages = normalized.pages
        if normalized.format:
            doc.size = (normalized.format.width_mm, normalized.format.height_mm)
    except Exception:
        pass


def refresh_appendix_meta(app) -> None:
    """Пересчитывает число листов приложения по его PDF."""
    path = getattr(app, "file_path", "") or ""
    if not path:
        return
    try:
        from ..core.pdf_reader import compute_pages
        app.pages = compute_pages(path)
    except Exception:
        pass
