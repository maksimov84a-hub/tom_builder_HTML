# -*- coding: utf-8 -*-
"""Список приложений (PDF) и их параметры."""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QHBoxLayout, QPushButton, QFormLayout, QLineEdit, QDialog,
)
from ..models import Appendix
from .common import FileRow, TableCrudForm


class AppendixDialog(QDialog):
    """Диалог добавления/редактирования приложения (PDF)."""

    def __init__(self, parent=None, app: Appendix | None = None):
        super().__init__(parent)
        self.setWindowTitle("Приложение")
        self.result_app: Appendix | None = None
        form = QFormLayout(self)
        self.number = QLineEdit()
        self.name = QLineEdit()
        self.file = FileRow("Приложение PDF", "PDF (*.pdf)", button_text="Выбрать PDF…")
        self.file.picked.connect(self._on_file_picked)
        self.target = QLineEdit()

        form.addRow("Номер (А, Б, В...):", self.number)
        form.addRow("Наименование:", self.name)
        form.addRow("Файл (PDF):", self.file)
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
            self.file.set_path(app.file_path)
            self.target.setText(app.target)

    def _on_file_picked(self, path: str) -> None:
        if path and not self.name.text().strip():
            self.name.setText(os.path.splitext(os.path.basename(path))[0])

    def _ok(self) -> None:
        a = self._app or Appendix()
        a.number = self.number.text().strip()
        a.name = self.name.text().strip()
        a.file_path = self.file.path()
        a.target = self.target.text().strip()
        from .common import refresh_appendix_meta
        refresh_appendix_meta(a)
        self.result_app = a
        self.accept()


class AppendicesForm(TableCrudForm):
    """Список приложений. Флажок «В сборку» исключает приложение из тома."""

    headers = ("№", "Наименование", "Файл", "Листы")
    delete_question = "Удалить приложение?"
    checkable = True

    def set_appendices(self, apps: list) -> None:
        self.set_items(apps)

    def get_appendices(self) -> list:
        return self.get_items()

    def cells(self, app) -> list:
        return [app.number, app.name, app.file_path, str(app.pages)]

    def ask_item(self, app):
        dlg = AppendixDialog(self, app)
        if dlg.exec() == QDialog.Accepted and dlg.result_app:
            return dlg.result_app
        return None
