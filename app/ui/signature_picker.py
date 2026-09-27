# -*- coding: utf-8 -*-
"""Диалог выбора подписанта для поля подписи."""
from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QListWidgetItem, QPushButton,
    QLabel, QMessageBox,
)
from PySide6.QtCore import Qt


class SignaturePickerDialog(QDialog):
    """Выбор подписанта для конкретной роли (key)."""

    def __init__(self, parent=None, role_label="", signers=None,
                 current_surname="", allow_clear=True):
        super().__init__(parent)
        self.setWindowTitle(f"Выбор подписанта — {role_label}")
        self.resize(400, 300)
        self.result_surname: str | None = None
        self.result_signature_image: str = ""

        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(f"Роль: {role_label}"))
        self.list = QListWidget()
        signers = signers or []
        for s in signers:
            item = QListWidgetItem()
            text = f"{s.get('surname', '')} — {s.get('position', '')}"
            item.setText(text)
            item.setData(Qt.UserRole, s)
            # отметить текущего
            if s.get("surname") == current_surname:
                item.setSelected(True)
            self.list.addItem(item)
        lay.addWidget(self.list)

        row = QHBoxLayout()
        self.btn_ok = QPushButton("OK")
        self.btn_clear = QPushButton("Убрать подпись")
        self.btn_cancel = QPushButton("Отмена")
        self.btn_ok.clicked.connect(self._ok)
        self.btn_clear.clicked.connect(self._clear)
        self.btn_cancel.clicked.connect(self.reject)
        row.addStretch(1)
        if allow_clear:
            row.addWidget(self.btn_clear)
        row.addWidget(self.btn_ok)
        row.addWidget(self.btn_cancel)
        lay.addLayout(row)

    def _ok(self) -> None:
        cur = self.list.currentItem()
        if cur is None:
            self.result_surname = None
            self.accept()
            return
        data = cur.data(Qt.UserRole)
        self.result_surname = data.get("surname", "")
        self.result_signature_image = data.get("signature_image", "")
        self.accept()

    def _clear(self) -> None:
        self.result_surname = None
        self.result_signature_image = ""
        self.accept()
