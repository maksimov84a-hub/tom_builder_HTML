# -*- coding: utf-8 -*-
"""Форма реестра подписантов: настройка подписей по Фамилии."""
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
    QFormLayout, QLineEdit, QDialog, QLabel, QFileDialog, QMessageBox,
)
from PySide6.QtCore import Qt

from ..models import Signer


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
        self.sig_path = QLineEdit()
        self.sig_path.setReadOnly(True)
        btn_sig = QPushButton("Выбрать изображение подписи...")
        btn_sig.clicked.connect(self._pick_sig)
        sig_row = QHBoxLayout()
        sig_row.addWidget(self.sig_path)
        sig_row.addWidget(btn_sig)

        form.addRow("Фамилия:", self.surname)
        form.addRow("Имя/Отчество:", self.name)
        form.addRow("Должность:", self.position)
        form.addRow("Роль:", self.role)
        form.addRow("Подпись:", sig_row)

        btns = QHBoxLayout()
        ok = QPushButton("OK"); ok.clicked.connect(self._ok)
        cancel = QPushButton("Отмена"); cancel.clicked.connect(self.reject)
        btns.addStretch(1); btns.addWidget(ok); btns.addWidget(cancel)
        form.addRow(btns)

        if signer:
            self._signer = signer
            self.surname.setText(signer.surname)
            self.name.setText(signer.name)
            self.position.setText(signer.position)
            self.role.setText(signer.role)
            self.sig_path.setText(signer.signature_image)
        else:
            self._signer = None

    def _pick_sig(self) -> None:
        f, _ = QFileDialog.getOpenFileName(self, "Изображение подписи", "",
                                           "Изображения (*.png *.jpg *.jpeg *.bmp);;Все файлы (*)")
        if f:
            self.sig_path.setText(f)

    def _ok(self) -> None:
        s = self._signer or Signer()
        s.surname = self.surname.text().strip()
        s.name = self.name.text().strip()
        s.position = self.position.text().strip()
        s.role = self.role.text().strip()
        s.signature_image = self.sig_path.text().strip()
        self.result_signer = s
        self.accept()


class SignersForm(QWidget):
    """Реестр подписантов (просмотр + добавление/редактирование)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        v = QVBoxLayout(self)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Фамилия", "Имя", "Должность", "Роль", "Подпись"])
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

        self._signers: list[Signer] = []

    def set_signers(self, signers: list) -> None:
        self._signers = signers
        self._refresh()

    def get_signers(self) -> list:
        return self._signers

    def _refresh(self) -> None:
        self.table.setRowCount(0)
        for s in self._signers:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(s.surname))
            self.table.setItem(r, 1, QTableWidgetItem(s.name))
            self.table.setItem(r, 2, QTableWidgetItem(s.position))
            self.table.setItem(r, 3, QTableWidgetItem(s.role))
            self.table.setItem(r, 4, QTableWidgetItem(s.signature_image or "—"))
        self.table.resizeColumnsToContents()

    def _add(self) -> None:
        dlg = SignerDialog(self)
        if dlg.exec() == QDialog.Accepted and dlg.result_signer:
            self._signers.append(dlg.result_signer)
            self._refresh()

    def _edit(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            return
        dlg = SignerDialog(self, self._signers[row])
        if dlg.exec() == QDialog.Accepted and dlg.result_signer:
            self._signers[row] = dlg.result_signer
            self._refresh()

    def _del(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            return
        if QMessageBox.question(self, "Удалить", "Удалить подписанта?") == QMessageBox.Yes:
            del self._signers[row]
            self._refresh()


class SignersDialog(QDialog):
    """Модальное окно редактирования реестра подписантов.

    Содержит SignersForm (добавление/редактирование/удаление подписантов).
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
