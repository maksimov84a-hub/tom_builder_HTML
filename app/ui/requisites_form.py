# -*- coding: utf-8 -*-
"""Форма ввода реквизитов тома.

Дополнительно содержит настройку графы 9 основной надписи (ГОСТ Р 21.101-2020,
приложение Ж): «наименование или различительный индекс организации». Графа
заполняется ЛИБО текстом (по центру ячейки), ЛИБО логотипом организации из
PDF-файла — логотип автоматически масштабируется под ячейку.
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QFormLayout, QLineEdit, QComboBox, QSpinBox, QVBoxLayout,
    QGroupBox, QScrollArea, QLabel, QCheckBox,
)
from PySide6.QtCore import Qt, Signal

from ..models import Requisites, Stage, STAGE_LABELS, STAMP9_TEXT, STAMP9_LOGO
from .common import FileRow


class RequisitesForm(QWidget):
    """Собирает реквизиты тома через форму ввода."""

    #: Изменилась настройка графы 9 (режим, текст или файл логотипа).
    cell9_changed = Signal()
    #: Изменилась настройка автоматического оглавления пояснительной записки.
    toc_changed = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        #: True, пока форма заполняется из модели (сигналы подавляются).
        self._loading_cell9 = False
        #: True, пока флажок оглавления заполняется программно.
        self._loading_toc = False
        self._build()

    def _build(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        scroll.setWidget(inner)
        outer.addWidget(scroll)

        form = QFormLayout(inner)
        form.setLabelAlignment(Qt.AlignRight)

        # --- Основные реквизиты ---
        grp_main = QGroupBox("Основные реквизиты")
        main_form = QFormLayout(grp_main)
        self.org = QLineEdit()
        self.org.setPlaceholderText("ООО «Проектная организация»")
        self.volume_id = QLineEdit()
        self.volume_id.setPlaceholderText("Например: 12-2024-ПЗ")
        self.object_name = QLineEdit()
        self.volume_name = QLineEdit()
        self.stage = QComboBox()
        self.stage.addItem(f"{STAGE_LABELS[Stage.P]} (П)", Stage.P.value)
        self.stage.addItem(f"{STAGE_LABELS[Stage.R]} (Р)", Stage.R.value)
        self.city = QLineEdit()
        self.year = QSpinBox()
        self.year.setRange(2000, 2100)
        main_form.addRow("Организация:", self.org)
        main_form.addRow("Обозначение тома:", self.volume_id)
        main_form.addRow("Наименование объекта:", self.object_name)
        main_form.addRow("Наименование тома:", self.volume_name)
        main_form.addRow("Стадия:", self.stage)
        main_form.addRow("Город:", self.city)
        main_form.addRow("Год:", self.year)
        form.addRow(grp_main)

        # --- Графа 9 основной надписи: текст ИЛИ логотип (PDF) ---
        grp9 = QGroupBox("Графа 9 основной надписи (организация)")
        self._cell9_form = QFormLayout(grp9)
        self.cell9_mode = QComboBox()
        self.cell9_mode.addItem("Текст (по центру ячейки)", STAMP9_TEXT)
        self.cell9_mode.addItem("Логотип (PDF, автомасштаб)", STAMP9_LOGO)
        self.cell9_mode.currentIndexChanged.connect(self._on_cell9_mode)
        self.cell9_logo = FileRow(
            "Логотип организации (PDF)", "PDF (*.pdf)",
            button_text="Файл логотипа…",
            placeholder="PDF-файл с логотипом организации", clearable=True)
        self.cell9_logo.picked.connect(self._on_cell9_logo_picked)
        self.cell9_hint = QLabel("")
        self.cell9_hint.setWordWrap(True)
        self.cell9_hint.setStyleSheet("color:#666;")
        self._cell9_form.addRow("Заполнение:", self.cell9_mode)
        self._cell9_form.addRow("Логотип:", self.cell9_logo)
        self._logo_row = self._cell9_form.rowCount() - 1
        self._cell9_form.addRow("", self.cell9_hint)
        form.addRow(grp9)

        # --- Автоматическое оглавление пояснительной записки ---
        # Оглавление строится по уровням заголовков исходного DOCX
        # (app.core.pz_toc) и вставляется первым листом записки в предпросмотр,
        # Word и PDF. Настройка хранится в конфигурации оформления (config).
        grp_toc = QGroupBox("Оглавление пояснительной записки")
        toc_form = QVBoxLayout(grp_toc)
        self.toc_enabled = QCheckBox(
            "Автоматическое оглавление по уровням текста из Word")
        self.toc_enabled.setToolTip(
            "Строится по заголовкам исходного DOCX и вставляется первым листом "
            "пояснительной записки. Отключите, если оглавление уже есть в файле.")
        toc_form.addWidget(self.toc_enabled)
        self.toc_enabled.toggled.connect(self._on_toc_toggled)
        form.addRow(grp_toc)

        self._on_cell9_mode()
        self.load_toc()

    # --- Автоматическое оглавление ПЗ ---

    def load_toc(self) -> None:
        """Загружает настройку оглавления из конфигурации оформления."""
        enabled = True
        try:
            from ..config import load_config
            enabled = bool(load_config().toc_enabled)
        except Exception:
            enabled = True
        self._loading_toc = True
        self.toc_enabled.setChecked(enabled)
        self._loading_toc = False

    def _on_toc_toggled(self, checked: bool) -> None:
        """Сохраняет настройку оглавления в конфигурацию и обновляет предпросмотр."""
        if self._loading_toc:
            return
        try:
            from ..config import load_config, save_config
            cfg = load_config()
            cfg.toc_enabled = bool(checked)
            save_config(cfg)
        except Exception:
            pass
        self.toc_changed.emit()

    # --- Графа 9 ---

    def _on_cell9_mode(self, *_args) -> None:
        """Показывает поля, относящиеся к выбранному режиму заполнения графы 9."""
        is_logo = (self.cell9_mode.currentData() or STAMP9_TEXT) == STAMP9_LOGO
        self._cell9_form.setRowVisible(self._logo_row, is_logo)
        self._update_cell9_hint()
        # При программной загрузке формы сигнал не нужен — иначе предпросмотр
        # перерисовывался бы лишний раз прямо в процессе открытия проекта.
        if not self._loading_cell9:
            self.cell9_changed.emit()

    def _on_cell9_logo_picked(self, path: str) -> None:
        """Выбран (или очищен) файл логотипа организации."""
        if path:
            try:
                from ..core.logo import logo_pixel_size
                w, h = logo_pixel_size(path)
                if w and h:
                    self.cell9_hint.setText(
                        f"Логотип подготовлен: {w}x{h} px, пропорции {w / h:.2f}:1. "
                        "Автомасштабируется под ячейку графы 9.")
                else:
                    self.cell9_hint.setText(
                        "Не удалось прочитать PDF — будет выведен текст графы 9.")
            except Exception:
                self.cell9_hint.setText("Не удалось прочитать PDF.")
        else:
            self._update_cell9_hint()
        self.cell9_changed.emit()

    def _update_cell9_hint(self) -> None:
        if self.cell9_mode.currentData() == STAMP9_LOGO:
            if not self.cell9_logo.path():
                self.cell9_hint.setText("Файл логотипа не выбран — выводится текст.")
            elif not self.cell9_hint.text():
                self.cell9_hint.setText("Логотип вписывается в ячейку без искажений.")
        else:
            self.cell9_hint.setText(
                "Текст выравнивается по центру ячейки графы 9 "
                "(форма 5 — 50 x 10 мм, форма 3 — 50 x 15 мм).")

    # --- Связывание с моделью ---

    def sync_requisites(self, req: Requisites) -> None:
        """Обновляет только реквизиты тома, не трогая настройки графы 9.

        Используется после inline-правки поля в предпросмотре: значения должны
        совпасть с моделью, но выбранный режим/файл логотипа сбрасывать нельзя.
        """
        self.load_requisites(req)

    def load(self, req: Requisites, stamp9=None) -> None:
        self.load_requisites(req)
        self.load_cell9(stamp9)

    def load_requisites(self, req: Requisites) -> None:
        """Заполняет основные реквизиты тома."""
        self.org.setText(req.organization)
        self.volume_id.setText(req.volume_id)
        self.object_name.setText(req.object_name)
        self.volume_name.setText(req.volume_name)
        idx = self.stage.findData(req.stage.value)
        self.stage.setCurrentIndex(idx if idx >= 0 else 0)
        self.city.setText(req.city)
        self.year.setValue(req.year)

    def load_cell9(self, stamp9) -> None:
        """Загружает настройки графы 9 (models.StampCell9) в форму."""
        self._loading_cell9 = True
        self.cell9_mode.blockSignals(True)
        if stamp9 is None:
            self.cell9_mode.setCurrentIndex(0)
            self.cell9_logo.set_path("")
        else:
            idx = self.cell9_mode.findData(stamp9.mode)
            self.cell9_mode.setCurrentIndex(idx if idx >= 0 else 0)
            self.cell9_logo.set_path(getattr(stamp9, "logo_path", "") or "")
        self.cell9_mode.blockSignals(False)
        self.cell9_hint.setText("")
        self._on_cell9_mode()
        self._loading_cell9 = False

    def apply_cell9(self, stamp9) -> None:
        """Сохраняет настройки графы 9 из формы в модель."""
        if stamp9 is None:
            return
        stamp9.mode = self.cell9_mode.currentData() or STAMP9_TEXT
        stamp9.logo_path = self.cell9_logo.path()

    def apply(self, req: Requisites, stamp9=None) -> None:
        req.organization = self.org.text().strip()
        req.volume_id = self.volume_id.text().strip()
        req.object_name = self.object_name.text().strip()
        req.volume_name = self.volume_name.text().strip()
        req.stage = self.stage.currentData()
        req.city = self.city.text().strip()
        req.year = self.year.value()
        self.apply_cell9(stamp9)
