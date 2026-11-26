# -*- coding: utf-8 -*-
"""Настройка обложки (прил. Е) и титульного листа (прил. Ж) по ГОСТ."""
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QGroupBox, QCheckBox, QPushButton,
)
from ..models import CoverSheet, TitleSheet
from .common import FileRow

TEMPLATE_FILTERS = ("Документы (*.docx *.doc *.pdf);;"
                    "Изображения (*.png *.jpg);;Все файлы (*)")
_IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".bmp")


class _TemplateGroup(QGroupBox):
    """Лист тома: авто по ГОСТ либо кастомный шаблон."""

    def __init__(self, title: str, auto_text: str, dialog_title: str, parent=None):
        super().__init__(title, parent)
        form = QFormLayout(self)
        self.auto = QCheckBox(auto_text)
        self.auto.toggled.connect(lambda checked: self.template.setEnabled(not checked))
        self.template = FileRow(dialog_title, TEMPLATE_FILTERS,
                                button_text="Выбрать шаблон…")
        form.addRow(self.auto)
        form.addRow("Шаблон:", self.template)
        self.toggle()

    def toggle(self) -> None:
        self.template.setEnabled(not self.auto.isChecked())

    def load(self, sheet) -> None:
        self.auto.setChecked(not sheet.use_custom)
        self.template.set_path(sheet.template_path)
        self.toggle()

    def apply(self, sheet) -> None:
        sheet.use_custom = not self.auto.isChecked()
        sheet.template_path = self.template.path()
        sheet.is_image = sheet.template_path.lower().endswith(_IMAGE_EXTENSIONS)


class CoverTitleForm(QWidget):
    """Форма настройки обложки и титульного листа.

    Группа показывается только та, что относится к выбранному листу дерева:
    на «Обложке» — настройки обложки, на «Титульном листе» — настройки титула.
    """

    def __init__(self, parent=None, cover: CoverSheet | None = None,
                 title: TitleSheet | None = None):
        super().__init__(parent)
        self._cover = cover or CoverSheet()
        self._title = title or TitleSheet()
        v = QVBoxLayout(self)
        self.cover_group = _TemplateGroup(
            "Обложка (приложение Е ГОСТ Р 21.101-2020)",
            "Формировать обложку автоматически по ГОСТ",
            "Шаблон обложки")
        self.title_group = _TemplateGroup(
            "Титульный лист (приложение Ж ГОСТ Р 21.101-2020)",
            "Формировать титульный лист автоматически по ГОСТ",
            "Шаблон титульного листа")
        v.addWidget(self.cover_group)
        v.addWidget(self.title_group)
        v.addStretch(1)
        self._load()
        self.show_part("cover")

    def show_part(self, kind: str) -> None:
        """Показывает настройки одного листа: 'cover' либо 'title'."""
        is_title = kind == "title"
        self.cover_group.setVisible(not is_title)
        self.title_group.setVisible(is_title)

    def _load(self) -> None:
        self.cover_group.load(self._cover)
        self.title_group.load(self._title)

    def apply_cover(self, cover: CoverSheet) -> None:
        self.cover_group.apply(cover)

    def apply_title(self, title: TitleSheet) -> None:
        self.title_group.apply(title)
