# -*- coding: utf-8 -*-
"""Журнал записей о внесении изменений."""
from __future__ import annotations

from PySide6.QtWidgets import (
    QHBoxLayout, QPushButton, QFormLayout, QLineEdit, QComboBox,
    QDialog, QDateEdit,
)
from PySide6.QtCore import QDate

from ..models import ChangeRecord, ChangeStatus
from .common import TableCrudForm


class ChangeDialog(QDialog):
    """Диалог добавления/редактирования записи об изменении."""

    def __init__(self, parent=None, rec: ChangeRecord | None = None,
                 signer_names: list | None = None):
        super().__init__(parent)
        self.setWindowTitle("Запись об изменении")
        self.result_rec: ChangeRecord | None = None
        signer_names = signer_names or []

        form = QFormLayout(self)
        self.number = QLineEdit()
        self.date = QDateEdit(QDate.currentDate())
        self.date.setCalendarPopup(True)
        self.desc = QLineEdit()
        self.target = QLineEdit()
        self.signer = QComboBox()
        self.signer.addItem("—", "")
        for nm in signer_names:
            self.signer.addItem(nm, nm)
        self.status = QComboBox()
        for st in ChangeStatus:
            self.status.addItem(st.value, st)

        form.addRow("№ изменения:", self.number)
        form.addRow("Дата:", self.date)
        form.addRow("Содержание изменения:", self.desc)
        form.addRow("Раздел/лист:", self.target)
        form.addRow("Исполнитель:", self.signer)
        form.addRow("Статус:", self.status)

        btns = QHBoxLayout()
        ok = QPushButton("OK"); ok.clicked.connect(self._ok)
        cancel = QPushButton("Отмена"); cancel.clicked.connect(self.reject)
        btns.addStretch(1); btns.addWidget(ok); btns.addWidget(cancel)
        form.addRow(btns)

        self._rec = rec
        if rec:
            self.number.setText(str(rec.number))
            self.desc.setText(rec.description)
            self.target.setText(rec.target)
            idx = self.signer.findData(rec.signer_surname)
            self.signer.setCurrentIndex(idx if idx >= 0 else 0)
            sidx = self.status.findData(rec.status)
            self.status.setCurrentIndex(sidx if sidx >= 0 else 0)
            if rec.date:
                self.date.setDate(QDate.fromString(rec.date, "dd.MM.yyyy"))

    def _ok(self) -> None:
        r = self._rec or ChangeRecord()
        r.number = int(self.number.text()) if self.number.text().isdigit() else 0
        r.date = self.date.date().toString("dd.MM.yyyy")
        r.description = self.desc.text().strip()
        r.target = self.target.text().strip()
        r.signer_surname = self.signer.currentData() or ""
        r.status = self.status.currentData()
        self.result_rec = r
        self.accept()


class ChangesForm(TableCrudForm):
    """Журнал изменений."""

    headers = ("№", "Дата", "Содержание", "Раздел/Лист", "Исполнитель", "Статус")
    delete_question = "Удалить запись об изменении?"

    def __init__(self, parent=None, get_signers=None):
        self._get_signers = get_signers or (lambda: [])
        super().__init__(parent)

    def set_records(self, records: list) -> None:
        self.set_items(records)

    def get_records(self) -> list:
        return self.get_items()

    def cells(self, record) -> list:
        return [str(record.number), record.date, record.description,
                record.target, record.signer_surname, record.status.value]

    def ask_item(self, record):
        names = [s.surname for s in self._get_signers() if s.surname]
        dlg = ChangeDialog(self, record, names)
        if dlg.exec() == QDialog.Accepted and dlg.result_rec:
            return dlg.result_rec
        return None
