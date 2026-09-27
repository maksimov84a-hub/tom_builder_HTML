# -*- coding: utf-8 -*-
"""Форма ввода реквизитов тома.

Дополнительно содержит настройку графы 9 основной надписи (ГОСТ Р 21.101-2020,
приложение Ж): «наименование или различительный индекс организации». Графа
заполняется ЛИБО текстом (по центру ячейки), ЛИБО логотипом организации из
PDF-файла — логотип автоматически масштабируется под ячейку.
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QFormLayout, QLineEdit, QComboBox, QSpinBox, QTextEdit, QVBoxLayout,
    QGroupBox, QScrollArea, QLabel, QPushButton, QHBoxLayout, QFileDialog,
    QCheckBox,
)
from PySide6.QtCore import Qt, Signal

from ..models import Requisites, Stage, STAMP9_TEXT, STAMP9_LOGO


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
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        scroll.setWidget(inner)
        outer.addWidget(scroll)

        self._form = QFormLayout(inner)
        self._form.setLabelAlignment(Qt.AlignRight)

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
        self.stage.addItem("Проектная (П)", Stage.P.value)
        self.stage.addItem("Рабочая (Р)", Stage.R.value)
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
        self._form.addRow(grp_main)

        # --- Графа 9 основной надписи: текст ИЛИ логотип (PDF) ---
        grp9 = QGroupBox("Графа 9 основной надписи (организация)")
        f9 = QFormLayout(grp9)
        self.cell9_mode = QComboBox()
        self.cell9_mode.addItem("Текст (по центру ячейки)", STAMP9_TEXT)
        self.cell9_mode.addItem("Логотип (PDF, автомасштаб)", STAMP9_LOGO)
        self.cell9_mode.currentIndexChanged.connect(self._on_cell9_mode)
        self.cell9_logo = QLineEdit()
        self.cell9_logo.setReadOnly(True)
        self.cell9_logo.setPlaceholderText("PDF-файл с логотипом организации")
        self.btn_cell9_logo = QPushButton("Файл логотипа…")
        self.btn_cell9_logo.clicked.connect(self._pick_cell9_logo)
        self.btn_cell9_clear = QPushButton("Очистить")
        self.btn_cell9_clear.clicked.connect(self._clear_cell9_logo)
        logo_row = QHBoxLayout()
        logo_row.addWidget(self.cell9_logo, stretch=1)
        logo_row.addWidget(self.btn_cell9_logo)
        logo_row.addWidget(self.btn_cell9_clear)
        self.cell9_hint = QLabel("")
        self.cell9_hint.setWordWrap(True)
        self.cell9_hint.setStyleSheet("color:#666;")
        f9.addRow("Заполнение:", self.cell9_mode)
        f9.addRow("Логотип:", logo_row)
        f9.addRow("", self.cell9_hint)
        self._cell9_logo_row = self._row_label(f9, logo_row)
        self._form.addRow(grp9)

        # --- Разделы (для содержания) ---
        grp_sec = QGroupBox("Разделы тома (для содержания)")
        sec_form = QVBoxLayout(grp_sec)
        self.sections = QTextEdit()
        self.sections.setPlaceholderText("Каждый раздел — с новой строки")
        self.sections.setFixedHeight(120)
        sec_form.addWidget(self.sections)
        self._form.addRow(grp_sec)

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
        self._form.addRow(grp_toc)

        self._form.addRow(QLabel(""))
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

    @staticmethod
    def _row_label(form: QFormLayout, widget) -> QLabel:
        """Возвращает подпись строки формы, содержащую widget (для скрытия).

        widget — либо виджет, либо компоновка (строка «Логотип» — это QHBoxLayout).
        """
        for i in range(form.rowCount()):
            item = form.itemAt(i, QFormLayout.LabelRole)
            field = form.itemAt(i, QFormLayout.FieldRole)
            if field is None:
                continue
            matched = field.widget() is widget if hasattr(widget, "isWidgetType") \
                else field.layout() is widget
            if matched:
                return item.widget() if item is not None else QLabel()
        return QLabel()

    # --- Графа 9 ---

    def _on_cell9_mode(self, *_args) -> None:
        """Показывает поля, относящиеся к выбранному режиму заполнения графы 9."""
        mode = self.cell9_mode.currentData() or STAMP9_TEXT
        is_logo = mode == STAMP9_LOGO
        self._cell9_logo_row.setVisible(is_logo)
        for w in (self.cell9_logo, self.btn_cell9_logo, self.btn_cell9_clear):
            w.setVisible(is_logo)
        self._update_cell9_hint()
        # При программной загрузке формы сигнал не нужен — иначе предпросмотр
        # перерисовывался бы лишний раз прямо в процессе открытия проекта.
        if not self._loading_cell9:
            self.cell9_changed.emit()

    def _pick_cell9_logo(self) -> None:
        """Выбор PDF-файла логотипа организации."""
        start = self.cell9_logo.text().strip()
        f, _ = QFileDialog.getOpenFileName(
            self, "Логотип организации (PDF)", start, "PDF (*.pdf)")
        if not f:
            return
        self.cell9_logo.setText(f)
        try:
            from ..core.logo import logo_pixel_size
            w, h = logo_pixel_size(f)
            if w and h:
                self.cell9_hint.setText(
                    f"Логотип подготовлен: {w}x{h} px, пропорции {w / h:.2f}:1. "
                    "Автомасштабируется под ячейку графы 9.")
            else:
                self.cell9_hint.setText(
                    "Не удалось прочитать PDF — будет выведен текст графы 9.")
        except Exception:
            self.cell9_hint.setText("Не удалось прочитать PDF.")
        self.cell9_changed.emit()

    def _clear_cell9_logo(self) -> None:
        self.cell9_logo.setText("")
        self._update_cell9_hint()
        self.cell9_changed.emit()

    def _update_cell9_hint(self) -> None:
        if self.cell9_mode.currentData() == STAMP9_LOGO:
            if not self.cell9_logo.text().strip():
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
        self.load(req, self._stamp9_from_form())

    def _stamp9_from_form(self):
        """Снимок текущих настроек графы 9 из формы (или None)."""
        from ..models import StampCell9
        s = StampCell9()
        self.apply_cell9(s)
        return s

    def load(self, req: Requisites, stamp9=None) -> None:
        self.org.setText(req.organization)
        self.volume_id.setText(req.volume_id)
        self.object_name.setText(req.object_name)
        self.volume_name.setText(req.volume_name)
        # стадия
        idx = self.stage.findData(req.stage.value)
        self.stage.setCurrentIndex(idx if idx >= 0 else 0)
        self.city.setText(req.city)
        self.year.setValue(req.year)
        self.sections.setPlainText("\n".join(req.sections))
        self.load_cell9(stamp9)

    # --- Графа 9: загрузка/сохранение ---

    def load_cell9(self, stamp9) -> None:
        """Загружает настройки графы 9 (models.StampCell9) в форму."""
        self._loading_cell9 = True
        self.cell9_mode.blockSignals(True)
        if stamp9 is None:
            self.cell9_mode.setCurrentIndex(0)
            self.cell9_logo.setText("")
        else:
            idx = self.cell9_mode.findData(stamp9.mode)
            self.cell9_mode.setCurrentIndex(idx if idx >= 0 else 0)
            self.cell9_logo.setText(getattr(stamp9, "logo_path", "") or "")
        self.cell9_mode.blockSignals(False)
        self.cell9_hint.setText("")
        self._on_cell9_mode()
        self._loading_cell9 = False

    def apply_cell9(self, stamp9) -> None:
        """Сохраняет настройки графы 9 из формы в модель."""
        if stamp9 is None:
            return
        stamp9.mode = self.cell9_mode.currentData() or STAMP9_TEXT
        stamp9.logo_path = self.cell9_logo.text().strip()

    def apply(self, req: Requisites, stamp9=None) -> None:
        req.organization = self.org.text().strip()
        req.volume_id = self.volume_id.text().strip()
        req.object_name = self.object_name.text().strip()
        req.volume_name = self.volume_name.text().strip()
        req.stage = self.stage.currentData()
        req.city = self.city.text().strip()
        req.year = self.year.value()
        req.sections = [s.strip() for s in self.sections.toPlainText().split("\n") if s.strip()]
        self.apply_cell9(stamp9)
