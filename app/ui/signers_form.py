# -*- coding: utf-8 -*-
"""Форма реестра подписантов: настройка подписей по Фамилии."""
from __future__ import annotations

from PySide6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QPushButton, QFormLayout, QLineEdit, QDialog,
)
from ..models import Signer
from .common import FileRow, TableCrudForm


class SignerDialog(QDialog):
    """Диалог добавления/редактирования подписанта."""

    def __init__(self, parent=None, signer: Signer | None = None):
        super().__init__(parent)
        self.setWindowTitle("Подписант")
        self.result_signer: Signer | None = None
        form = QFormLayout(self)
        self.surname = QLineEdit()
        self.name = QLineEdit()
        self.position = QLineEdit()
        self.role = QLineEdit()
        self.sig = FileRow("Изображение подписи",
                           "Изображения (*.png *.jpg *.jpeg *.bmp);;Все файлы (*)",
                           button_text="Выбрать изображение…", clearable=True)

        form.addRow("Фамилия:", self.surname)
        form.addRow("Имя/Отчество:", self.name)
        form.addRow("Должность:", self.position)
        form.addRow("Роль:", self.role)
        form.addRow("Подпись:", self.sig)

        btns = QHBoxLayout()
        ok = QPushButton("OK"); ok.clicked.connect(self._ok)
        cancel = QPushButton("Отмена"); cancel.clicked.connect(self.reject)
        btns.addStretch(1); btns.addWidget(ok); btns.addWidget(cancel)
        form.addRow(btns)

        self._signer = signer
        if signer:
            self.surname.setText(signer.surname)
            self.name.setText(signer.name)
            self.position.setText(signer.position)
            self.role.setText(signer.role)
            self.sig.set_path(signer.signature_image)

    def _ok(self) -> None:
        s = self._signer or Signer()
        s.surname = self.surname.text().strip()
        s.name = self.name.text().strip()
        s.position = self.position.text().strip()
        s.role = self.role.text().strip()
        s.signature_image = self.sig.path()
        self.result_signer = s
        self.accept()


class SignersForm(TableCrudForm):
    """Реестр подписантов (просмотр + добавление/редактирование)."""

    headers = ("Фамилия", "Имя", "Должность", "Роль", "Подпись")
    delete_question = "Удалить подписанта?"

    def set_signers(self, signers: list) -> None:
        self.set_items(signers)

    def get_signers(self) -> list:
        return self.get_items()

    def cells(self, signer) -> list:
        return [signer.surname, signer.name, signer.position, signer.role,
                signer.signature_image or "—"]

    def ask_item(self, signer):
        dlg = SignerDialog(self, signer)
        if dlg.exec() == QDialog.Accepted and dlg.result_signer:
            return dlg.result_signer
        return None


class SignersDialog(QDialog):
    """Модальное окно редактирования реестра подписантов.

    Использует переданную форму, чтобы изменения сразу попадали в общий реестр.
    """

    def __init__(self, parent=None, signers_form: SignersForm | None = None):
        super().__init__(parent)
        self.setWindowTitle("Подписанты")
        self.setMinimumSize(640, 420)
        v = QVBoxLayout(self)
        self.form = signers_form or SignersForm()
        v.addWidget(self.form)

        btns = QHBoxLayout()
        close = QPushButton("Готово")
        close.clicked.connect(self.accept)
        btns.addStretch(1)
        btns.addWidget(close)
        v.addLayout(btns)
