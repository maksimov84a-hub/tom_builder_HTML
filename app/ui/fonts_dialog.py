# -*- coding: utf-8 -*-
"""Настройки шрифтов текста пояснительной записки.

Один пункт меню задаёт шрифт и кегль для трёх групп текста ПЗ:

    обычный текст | заголовки | таблицы

Значения хранятся в конфигурации оформления (`config.json`, ключ `text_fonts`)
и используются всеми выводами — предпросмотром, Word (DOCX) и PDF, — поэтому
кегль в предпросмотре совпадает с собранным томом.
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox, QLabel,
    QPushButton, QFontComboBox, QSpinBox, QWidget,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from ..config import Config, TEXT_FONT_GROUPS, TextFonts

#: Текст образца для группы: показывает, как выглядит вывод.
_SAMPLE_TEXT = {
    "body": "Основной текст записки: абзац начинается с отступа 5 знаков.",
    "heading": "Заголовок раздела",
    "table": "Ячейка таблицы",
}


class TextFontsDialog(QDialog):
    """Модальное окно «Шрифт и кегль текста»: настройки по группам."""

    def __init__(self, parent=None, config: Config | None = None):
        super().__init__(parent)
        self.setWindowTitle("Шрифт и кегль текста пояснительной записки")
        self.setMinimumWidth(560)
        self._config = config or Config()
        self._fonts = TextFonts.from_dict(self._config.text_fonts.to_dict())
        #: (группа, поле шрифта, поле кегля, исходное имя шрифта).
        self._rows: list = []
        #: группа -> QLabel образца.
        self._samples: dict = {}
        self._build()
        self._update_samples()

    # ---- построение окна ---------------------------------------------------

    def _build(self) -> None:
        v = QVBoxLayout(self)

        hint = QLabel(
            "Шрифт и кегль применяются к тексту пояснительной записки во всех "
            "выводах: предпросмотр, Word и PDF.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#555;")
        v.addWidget(hint)

        grp = QGroupBox("Текст пояснительной записки")
        form = QFormLayout(grp)
        form.setLabelAlignment(Qt.AlignRight)
        for group, label in TEXT_FONT_GROUPS.items():
            row = QWidget()
            h = QHBoxLayout(row)
            h.setContentsMargins(0, 0, 0, 0)
            font_box = QFontComboBox()
            current = self._fonts.font(group)
            font_box.setCurrentFont(self._font_by_name(current))
            size_box = QSpinBox()
            size_box.setRange(4, 72)
            size_box.setSuffix(" pt")
            size_box.setValue(int(self._fonts.size(group)))
            h.addWidget(font_box, stretch=1)
            h.addWidget(size_box)
            form.addRow(f"{label}:", row)
            # Исходное имя шрифта: если семейства нет в системе, QFontComboBox
            # подставит первый из списка — тогда сохраняем исходное значение,
            # чтобы настройка не потерялась молча.
            self._rows.append((group, font_box, size_box, current))
            font_box.currentFontChanged.connect(self._update_samples)
            size_box.valueChanged.connect(self._update_samples)
        v.addWidget(grp)

        sample_grp = QGroupBox("Образец")
        sample_form = QFormLayout(sample_grp)
        sample_form.setLabelAlignment(Qt.AlignRight)
        for group, label in TEXT_FONT_GROUPS.items():
            sample = QLabel(_SAMPLE_TEXT.get(group, ""))
            sample.setWordWrap(True)
            sample.setTextFormat(Qt.PlainText)
            sample.setMinimumHeight(24)
            self._samples[group] = sample
            sample_form.addRow(f"{label}:", sample)
        self._missing_hint = QLabel("")
        self._missing_hint.setWordWrap(True)
        self._missing_hint.setStyleSheet("color:#a00;")
        sample_form.addRow("", self._missing_hint)
        v.addWidget(sample_grp)

        note = QLabel(
            "Абзацный отступ (5 знаков) считается от кегля обычного текста, "
            "поэтому при смене кегля он пересчитывается автоматически.")
        note.setWordWrap(True)
        note.setStyleSheet("color:#555;")
        v.addWidget(note)

        btns = QHBoxLayout()
        reset = QPushButton("По умолчанию")
        reset.clicked.connect(self._reset)
        btns.addWidget(reset)
        btns.addStretch(1)
        ok = QPushButton("OK")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        cancel = QPushButton("Отмена")
        cancel.clicked.connect(self.reject)
        btns.addWidget(ok)
        btns.addWidget(cancel)
        v.addLayout(btns)

    # ---- вспомогательное ---------------------------------------------------

    @staticmethod
    def _font_by_name(name: str) -> QFont:
        """QFont по имени семейства (для QFontComboBox)."""
        return QFont((name or "").strip())

    @staticmethod
    def _families() -> list:
        """Установленные в системе семейства шрифтов (список Qt)."""
        from PySide6.QtGui import QFontDatabase
        return QFontDatabase.families()

    @classmethod
    def _family_available(cls, name: str) -> bool:
        """Есть ли семейство в системе (иначе Qt/Word подставят своё).

        Если Qt не видит ни одного шрифта (например, окно создаётся без
        графической сессии), проверка считается успешной: иначе настройка
        молча не сохранилась бы.
        """
        families = cls._families()
        if not families:
            return True
        return bool(name) and name in families

    def _update_samples(self) -> None:
        """Показывает выбранные шрифт и кегль в образце текста."""
        missing = []
        for group, font_box, size_box, _orig in self._rows:
            family = font_box.currentFont().family()
            sample = self._samples.get(group)
            if sample is not None:
                font = QFont(font_box.currentFont())
                font.setPointSize(int(size_box.value()))
                sample.setFont(font)
            if not self._family_available(family):
                missing.append(family or "—")
        if missing:
            self._missing_hint.setText(
                "Не установлены в системе: " + ", ".join(sorted(set(missing))) +
                ". Word и PDF подставят своё начертание — выберите шрифт из списка.")
        else:
            self._missing_hint.setText("")

    def _reset(self) -> None:
        """Возвращает значения по умолчанию (не применяя их до «OK»)."""
        defaults = TextFonts()
        for group, font_box, size_box, _orig in self._rows:
            font_box.setCurrentFont(self._font_by_name(defaults.font(group)))
            size_box.setValue(int(defaults.size(group)))
        self._update_samples()

    # ---- результат ---------------------------------------------------------

    #: Служебные подстановки Qt, когда запрошенного семейства нет в системе.
    _QT_SUBSTITUTES = frozenset({"sans serif", "serif", "monospace", "cursive",
                                 "fantasy", "system"})

    def result_fonts(self) -> TextFonts:
        """Шрифты и кегли, выбранные в окне."""
        fonts = TextFonts.from_dict(self._fonts.to_dict())
        for group, font_box, size_box, original in self._rows:
            family = font_box.currentFont().family()
            # Семейства нет в системе — QFontComboBox показал бы чужой шрифт
            # или служебную подстановку Qt; сохраняем исходное имя, чтобы
            # настройка не подменялась молча.
            if family.strip().lower() in self._QT_SUBSTITUTES:
                family = original or family
            elif not self._family_available(original) and not self._family_available(family):
                family = original
            fonts.set(group, family, size_box.value())
        return fonts

    def apply_to(self, config: Config) -> None:
        """Записывает выбранные значения в конфигурацию."""
        config.text_fonts = self.result_fonts()
