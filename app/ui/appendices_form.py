# -*- coding: utf-8 -*-
"""Список приложений (PDF) и их параметры."""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
    QFormLayout, QLineEdit, QDialog, QLabel, QFileDialog, QMessageBox,
)
from PySide6.QtCore import Qt

from ..models import Appendix


class AppendixDialog(QDialog):
    """Диалог добавления/редактирования приложения (PDF)."""

    def __init__(self, parent=None, app: Appendix | None = None):
        super().__init__(parent)
        self.setWindowTitle("Приложение")
        self.result_app: Appendix | None = None
        form = QFormLayout(self)
        self.number = QLineEdit()
        self.name = QLineEdit()
        self.file = QLineEdit()
        self.file.setReadOnly(True)
        self.target = QLineEdit()
        btn_file = QPushButton("Выбрать PDF...")
        btn_file.clicked.connect(self._pick)
        file_row = QHBoxLayout()
        file_row.addWidget(self.file)
        file_row.addWidget(btn_file)

        form.addRow("Номер (А, Б, В...):", self.number)
        form.addRow("Наименование:", self.name)
        form.addRow("Файл (PDF):", file_row)
        form.addRow("Раздел-привязка:", self.target)

        btns = QHBoxLayout()
        ok = QPushButton("OK"); ok.clicked.connect(self._ok)
        cancel = QPushButton("Отмена"); cancel.clicked.connect(self.reject)
        btns.addStretch(1); btns.addWidget(ok); btns.addWidget(cancel)
        form.addRow(btns)

        self._app = app
        if app:
            self.number.setText(app.number)
            self.name.setText(app.name)
            self.file.setText(app.file_path)
            self.target.setText(app.target)

    def _pick(self) -> None:
        f, _ = QFileDialog.getOpenFileName(self, "Приложение PDF", "", "PDF (*.pdf)")
        if f:
            self.file.setText(f)
            if not self.name.text().strip():
                self.name.setText(os.path.splitext(os.path.basename(f))[0])

    def _ok(self) -> None:
        a = self._app or Appendix()
        a.number = self.number.text().strip()
        a.name = self.name.text().strip()
        a.file_path = self.file.text().strip()
        a.target = self.target.text().strip()
        a.pages = self._count_pages(a.file_path)
        self.result_app = a
        self.accept()

    @staticmethod
    def _count_pages(path: str) -> int:
        if not path:
            return 0
        try:
            import pymupdf
            with pymupdf.open(path) as doc:
                return doc.page_count
        except Exception:
            return 0


class AppendicesForm(QWidget):
    """Список приложений."""

    def __init__(self, parent=None):
        super().__init__(parent)
        v = QVBoxLayout(self)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["№", "Наименование", "Файл", "Листы"])
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
        row.addWidget(self.btn_add); row.addWidget(self.btn_edit); row.addWidget(self.btn_del)
        row.addStretch(1)
        v.addLayout(row)

        self._apps: list[Appendix] = []

    def set_appendices(self, apps: list) -> None:
        self._apps = apps
        self._refresh()

    def get_appendices(self) -> list:
        return self._apps

    def _refresh(self) -> None:
        self.table.setRowCount(0)
        for a in self._apps:
            i = self.table.rowCount()
            self.table.insertRow(i)
            self.table.setItem(i, 0, QTableWidgetItem(a.number))
            self.table.setItem(i, 1, QTableWidgetItem(a.name))
            self.table.setItem(i, 2, QTableWidgetItem(a.file_path))
            self.table.setItem(i, 3, QTableWidgetItem(str(a.pages)))
        self.table.resizeColumnsToContents()

    def _add(self) -> None:
        dlg = AppendixDialog(self)
        if dlg.exec() == QDialog.Accepted and dlg.result_app:
            self._apps.append(dlg.result_app)
            self._refresh()

    def _edit(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            return
        dlg = AppendixDialog(self, self._apps[row])
        if dlg.exec() == QDialog.Accepted and dlg.result_app:
            self._apps[row] = dlg.result_app
            self._refresh()

    def _del(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            return
        if QMessageBox.question(self, "Удалить", "Удалить приложение?") == QMessageBox.Yes:
            del self._apps[row]
            self._refresh()
