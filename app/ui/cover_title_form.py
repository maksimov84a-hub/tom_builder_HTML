# -*- coding: utf-8 -*-
"""Настройка обложки (прил. Е) и титульного листа (прил. Ж) по ГОСТ."""
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QGroupBox, QCheckBox, QLineEdit,
    QPushButton, QFileDialog, QLabel, QHBoxLayout,
)
from ..models import CoverSheet, TitleSheet


class CoverTitleForm(QWidget):
    """Форма настройки обложки и титульного листа."""

    def __init__(self, parent=None, cover: CoverSheet | None = None, title: TitleSheet | None = None):
        super().__init__(parent)
        self._cover = cover or CoverSheet()
        self._title = title or TitleSheet()
        v = QVBoxLayout(self)

        # --- Обложка ---
        g_cover = QGroupBox("Обложка (приложение Е ГОСТ Р 21.101-2020)")
        cf = QFormLayout(g_cover)
        self.cover_auto = QCheckBox("Формировать обложку автоматически по ГОСТ")
        self.cover_auto.toggled.connect(self._toggle_cover)
        self.cover_logo = QCheckBox("Показывать логотип организации")
        self.cover_custom_path = QLineEdit()
        self.cover_custom_path.setReadOnly(True)
        btn_cover = QPushButton("Выбрать шаблон...")
        btn_cover.clicked.connect(self._pick_cover)
        cover_row = QHBoxLayout()
        cover_row.addWidget(self.cover_custom_path); cover_row.addWidget(btn_cover)
        cf.addRow(self.cover_auto)
        cf.addRow(self.cover_logo)
        cf.addRow("Шаблон:", cover_row)
        v.addWidget(g_cover)

        # --- Титульный лист ---
        g_title = QGroupBox("Титульный лист (приложение Ж ГОСТ Р 21.101-2020)")
        tf = QFormLayout(g_title)
        self.title_auto = QCheckBox("Формировать титульный лист автоматически по ГОСТ")
        self.title_auto.toggled.connect(self._toggle_title)
        self.title_custom_path = QLineEdit()
        self.title_custom_path.setReadOnly(True)
        btn_title = QPushButton("Выбрать шаблон...")
        btn_title.clicked.connect(self._pick_title)
        title_row = QHBoxLayout()
        title_row.addWidget(self.title_custom_path); title_row.addWidget(btn_title)
        tf.addRow(self.title_auto)
        tf.addRow("Шаблон:", title_row)
        v.addWidget(g_title)

        v.addStretch(1)

        # Загрузить текущие значения
        self._load()

    def _toggle_cover(self, checked: bool) -> None:
        self.cover_custom_path.setEnabled(not checked)

    def _toggle_title(self, checked: bool) -> None:
        self.title_custom_path.setEnabled(not checked)

    def _pick_cover(self) -> None:
        f, _ = QFileDialog.getOpenFileName(self, "Шаблон обложки", "",
                                           "Документы (*.docx *.doc *.pdf);;Изображения (*.png *.jpg);;Все файлы (*)")
        if f:
            self.cover_custom_path.setText(f)

    def _pick_title(self) -> None:
        f, _ = QFileDialog.getOpenFileName(self, "Шаблон титульного листа", "",
                                           "Документы (*.docx *.doc *.pdf);;Изображения (*.png *.jpg);;Все файлы (*)")
        if f:
            self.title_custom_path.setText(f)

    def _load(self) -> None:
        self.cover_auto.setChecked(not self._cover.use_custom)
        self.cover_logo.setChecked(self._cover.show_logo)
        self.cover_custom_path.setText(self._cover.template_path)
        self.title_auto.setChecked(not self._title.use_custom)
        self.title_custom_path.setText(self._title.template_path)
        self._toggle_cover(self.cover_auto.isChecked())
        self._toggle_title(self.title_auto.isChecked())

    def apply_cover(self, cover: CoverSheet) -> None:
        cover.use_custom = not self.cover_auto.isChecked()
        cover.show_logo = self.cover_logo.isChecked()
        cover.template_path = self.cover_custom_path.text().strip()
        if cover.template_path:
            ext = cover.template_path.lower()
            cover.is_image = ext in (".png", ".jpg", ".jpeg", ".bmp")

    def apply_title(self, title: TitleSheet) -> None:
        title.use_custom = not self.title_auto.isChecked()
        title.template_path = self.title_custom_path.text().strip()
        if title.template_path:
            ext = title.template_path.lower()
            title.is_image = ext in (".png", ".jpg", ".jpeg", ".bmp")
