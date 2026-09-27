# -*- coding: utf-8 -*-
"""Журнал записей о внесении изменений."""
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
    QFormLayout, QLineEdit, QComboBox, QDialog, QLabel, QMessageBox, QDateEdit,
)
from PySide6.QtCore import Qt, QDate

from ..models import ChangeRecord, ChangeStatus


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


class ChangesForm(QWidget):
    """Журнал изменений."""

    def __init__(self, parent=None, get_signers=None):
        super().__init__(parent)
        self._get_signers = get_signers or (lambda: [])
        v = QVBoxLayout(self)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["№", "Дата", "Содержание", "Раздел/Лист", "Исполнитель", "Статус"])
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

        self._records: list[ChangeRecord] = []

    def set_records(self, records: list) -> None:
        self._records = records
        self._refresh()

    def get_records(self) -> list:
        return self._records

    def _signer_names(self) -> list:
        return [s.surname for s in self._get_signers() if s.surname]

    def _refresh(self) -> None:
        self.table.setRowCount(0)
        for r in self._records:
            i = self.table.rowCount()
            self.table.insertRow(i)
            self.table.setItem(i, 0, QTableWidgetItem(str(r.number)))
            self.table.setItem(i, 1, QTableWidgetItem(r.date))
            self.table.setItem(i, 2, QTableWidgetItem(r.description))
            self.table.setItem(i, 3, QTableWidgetItem(r.target))
            self.table.setItem(i, 4, QTableWidgetItem(r.signer_surname))
            self.table.setItem(i, 5, QTableWidgetItem(r.status.value))
        self.table.resizeColumnsToContents()

    def _add(self) -> None:
        dlg = ChangeDialog(self, signer_names=self._signer_names())
        if dlg.exec() == QDialog.Accepted and dlg.result_rec:
            self._records.append(dlg.result_rec)
            self._refresh()

    def _edit(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            return
        dlg = ChangeDialog(self, self._records[row], self._signer_names())
        if dlg.exec() == QDialog.Accepted and dlg.result_rec:
            self._records[row] = dlg.result_rec
            self._refresh()

    def _del(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            return
        if QMessageBox.question(self, "Удалить", "Удалить запись об изменении?") == QMessageBox.Yes:
            del self._records[row]
            self._refresh()
