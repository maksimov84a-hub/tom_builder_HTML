# -*- coding: utf-8 -*-
"""Мини-окно «Собрать том»: выбор форматов и кнопки «Отмена» / «Собрать».

Открывается пунктом меню «Сборка → Собрать том…». Форматы отмечаются
флажками; выбранный набор запоминается в config.json (`Config.build_formats`)
и восстанавливается при следующем открытии окна.
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QCheckBox, QDialog, QDialogButtonBox, QLabel, QVBoxLayout,
)
from PySide6.QtCore import Qt


class BuildDialog(QDialog):
    """Окно выбора форматов сборки тома: флажки и «Отмена» / «Собрать»."""

    #: Форматы тома в порядке флажков: (ключ, подпись).
    FORMATS = (
        ("pdf", "PDF (*.pdf)"),
        ("docx", "DOC — Word (*.docx)"),
        ("html", "HTML (*.html)"),
    )
    DEFAULT_FORMATS = ("pdf", "docx", "html")

    def __init__(self, parent=None, selected=None, chapters: bool = False):
        super().__init__(parent)
        self.setWindowTitle("Собрать том")
        self.setModal(True)
        # Небольшое окно по содержимому: без растягивания на весь экран.
        self.setMinimumWidth(320)

        v = QVBoxLayout(self)
        v.setContentsMargins(14, 12, 14, 12)
        v.setSpacing(8)

        title = QLabel("Форматы сборки")
        title.setStyleSheet("font-weight:bold;")
        v.addWidget(title)

        hint = QLabel("Отметьте форматы, в которых нужно собрать том.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#555;")
        v.addWidget(hint)

        # selected=None — набор не задан, отмечаем все форматы. Пустой список
        # это именно «ничего не отмечено»: иначе окно нельзя было бы открыть
        # в состоянии «форматы не выбраны».
        if selected is None:
            chosen = list(self.DEFAULT_FORMATS)
        else:
            chosen = [key for key in selected]
        self._boxes: dict = {}
        for key, label in self.FORMATS:
            box = QCheckBox(label)
            box.setChecked(key in chosen)
            box.toggled.connect(self._update_ok_state)
            v.addWidget(box)
            self._boxes[key] = box

        # Режим Word: один документ на том либо отдельные файлы по разделам.
        self.chapters_box = QCheckBox("DOC — по разделам (отдельные файлы)")
        self.chapters_box.setToolTip(
            "Каждый раздел тома (обложка, титул, содержание, состав ПД, ПЗ,\n"
            "графическая часть, приложения, изменения) — в свой файл *.docx")
        self.chapters_box.setChecked(bool(chapters))
        self.chapters_box.toggled.connect(self._update_ok_state)
        v.addWidget(self.chapters_box)

        # Кнопки со стандартными ролями: «Отмена» закрывает окно, «Собрать»
        # подтверждает выбор (кнопка Ok с подписью «Собрать»).
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        self.buttons.button(QDialogButtonBox.Ok).setText("Собрать")
        self.buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        v.addWidget(self.buttons)

        self._update_ok_state()

    def _update_ok_state(self) -> None:
        """«Собрать» недоступна, когда не отмечен ни один формат."""
        ok = self.buttons.button(QDialogButtonBox.Ok)
        ok.setEnabled(bool(self.selected_formats()))
        # Отдельные файлы по разделам — это только Word.
        self.chapters_box.setEnabled("docx" in self.selected_formats())

    def selected_formats(self) -> list:
        """Отмеченные форматы в порядке флажков (пустой список — ничего)."""
        return [key for key, _label in self.FORMATS
                if self._boxes[key].isChecked()]

    def chapters(self) -> bool:
        """True — DOCX собирать по разделам в отдельные файлы."""
        return self.chapters_box.isChecked() and self.chapters_box.isEnabled()
