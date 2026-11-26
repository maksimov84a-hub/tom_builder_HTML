# -*- coding: utf-8 -*-
"""Настройки раздела тома: дополнительное обозначение и перечень листов ГЧ.

Обозначение раздела = обозначение тома + это дополнение («02-2026-ТР» +
«.ПЗ» = «02-2026-ТР.ПЗ»). Значение показывается в графе 1 основной надписи
листов раздела и в содержании тома.

Для графической части дополнительно настраивается лист «Графическая часть»
с перечнем листов: включение, источник строк (по листам раздела или из Excel)
и сам файл Excel.
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QGroupBox, QLineEdit, QLabel,
    QCheckBox, QComboBox, QPushButton, QFileDialog, QHBoxLayout,
)
from PySide6.QtCore import Signal

from ..models import (VolumeProject, SECTION_LABELS, SECTION_CODE_DEFAULTS,
                      SECTION_GRAPHIC, SHEETLIST_AUTO, SHEETLIST_EXCEL,
                      SHEETLIST_PDF)


class SectionSettingsForm(QWidget):
    """Форма «Обозначение раздела» (+ перечень листов для графики)."""

    #: Дополнительное обозначение изменено.
    changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.project: VolumeProject | None = None
        self.section_kind: str | None = None
        v = QVBoxLayout(self)
        v.setContentsMargins(8, 8, 8, 8)

        self.grp = QGroupBox("Обозначение раздела")
        form = QFormLayout(self.grp)
        self.lbl_section = QLabel("—")
        self.edit_code = QLineEdit()
        self.lbl_result = QLabel("—")
        self.lbl_result.setWordWrap(True)

        form.addRow("Раздел:", self.lbl_section)
        form.addRow("Доп. обозначение:", self.edit_code)
        form.addRow("Обозначение:", self.lbl_result)
        self.edit_code.textEdited.connect(self._on_code_typing)
        self.edit_code.editingFinished.connect(self._on_code_committed)
        v.addWidget(self.grp)

        # ---- Лист «Графическая часть» с перечнем листов ----
        self.grp_list = QGroupBox("Лист «Графическая часть»")
        list_form = QFormLayout(self.grp_list)
        self.chk_list = QCheckBox("Вставлять лист с перечнем листов")
        self.cmb_source = QComboBox()
        self.cmb_source.addItem("Из PDF листов (штамп)", SHEETLIST_PDF)
        self.cmb_source.addItem("По листам раздела", SHEETLIST_AUTO)
        self.cmb_source.addItem("Из Excel", SHEETLIST_EXCEL)
        self.edit_template = QLineEdit()
        self.edit_template.setPlaceholderText("{volume} л.{no}")
        self.edit_template.setToolTip(
            "Шаблон обозначения листа: {volume} — шифр раздела, {no} — номер "
            "листа из штампа")
        self.edit_excel = QLineEdit()
        self.edit_excel.setPlaceholderText("Путь к книге .xlsx")
        self.btn_browse = QPushButton("…")
        self.btn_browse.setFixedWidth(28)
        self.btn_browse.setToolTip("Выбрать файл Excel с перечнем листов")
        row_excel = QWidget()
        h = QHBoxLayout(row_excel)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(4)
        h.addWidget(self.edit_excel, stretch=1)
        h.addWidget(self.btn_browse)
        self.lbl_list_hint = QLabel("")
        self.lbl_list_hint.setWordWrap(True)
        self.lbl_list_hint.setStyleSheet("color:#555;")

        list_form.addRow("", self.chk_list)
        list_form.addRow("Строки перечня:", self.cmb_source)
        list_form.addRow("Шаблон обозначения:", self.edit_template)
        list_form.addRow("Файл Excel:", row_excel)
        list_form.addRow("", self.lbl_list_hint)
        v.addWidget(self.grp_list)
        v.addStretch(1)

        self.chk_list.toggled.connect(self._on_list_changed)
        self.cmb_source.currentIndexChanged.connect(self._on_source_changed)
        self.edit_template.editingFinished.connect(self._on_template_committed)
        self.edit_excel.editingFinished.connect(self._on_excel_committed)
        self.btn_browse.clicked.connect(self._choose_excel)

    # ---- загрузка / показ ----
    def show_section(self, project: VolumeProject, section_kind: str) -> None:
        """Показывает код раздела и настройки перечня из проекта."""
        self.project = project
        self.section_kind = section_kind
        self.lbl_section.setText(
            SECTION_LABELS.get(section_kind, section_kind or "—"))
        self.edit_code.blockSignals(True)
        self.edit_code.setText(project.section_code(section_kind))
        self.edit_code.blockSignals(False)
        self.edit_code.setPlaceholderText(
            SECTION_CODE_DEFAULTS.get(section_kind, ""))
        self._update_result()
        self._load_sheet_list()

    def _load_sheet_list(self) -> None:
        """Заполняет блок перечня листов (только для графической части)."""
        is_graphic = self.section_kind == SECTION_GRAPHIC
        self.grp_list.setVisible(is_graphic)
        if not is_graphic or self.project is None:
            return
        from ..core.preview_html import _graphic_sheet_list_cfg
        gl = _graphic_sheet_list_cfg(self.project)
        # Сигналы блокируются: иначе установка значений запускает
        # обработчики и пишет в модель то, что только что прочитали.
        for w in (self.chk_list, self.cmb_source, self.edit_excel,
                  self.edit_template):
            w.blockSignals(True)
        self.chk_list.setChecked(bool(gl.enabled))
        index = self.cmb_source.findData(gl.source)
        self.cmb_source.setCurrentIndex(index if index >= 0 else 0)
        self.edit_excel.setText(gl.excel_path)
        self.edit_template.setText(gl.designation_template)
        for w in (self.chk_list, self.cmb_source, self.edit_excel,
                  self.edit_template):
            w.blockSignals(False)
        self._update_list_state()

    def _update_list_state(self) -> None:
        """Показывает/скрывает поля в зависимости от выбора."""
        enabled = self.chk_list.isChecked()
        source = self.cmb_source.currentData()
        is_excel = source == SHEETLIST_EXCEL
        is_pdf = source == SHEETLIST_PDF
        self.cmb_source.setEnabled(enabled)
        self.edit_template.setEnabled(enabled and is_pdf)
        self.edit_excel.setEnabled(enabled and is_excel)
        self.btn_browse.setEnabled(enabled and is_excel)
        hint = ""
        if enabled and is_excel:
            path = self.edit_excel.text().strip()
            if not path:
                hint = "Укажите файл .xlsx с таблицей."
            elif not os.path.exists(path):
                hint = "Файл не найден — будут использованы листы раздела."
            else:
                try:
                    from ..core.xlsx_reader import read_sheet_list
                    rows = read_sheet_list(path)
                    hint = (f"Прочитано строк: {len(rows)}. "
                            "Первая строка — шапка, пустые строки пропускаются."
                            if rows else
                            "В файле не найдено строк перечня — будут "
                            "использованы листы раздела.")
                except Exception:
                    hint = "Не удалось прочитать файл."
        elif enabled and is_pdf:
            hint = self._pdf_hint()
        elif enabled:
            hint = "Строки строятся по листам графической части раздела."
        self.lbl_list_hint.setText(hint)

    def _pdf_hint(self) -> str:
        """Подсказка для источника «Из PDF»: сколько листов распознано."""
        if self.project is None:
            return ""
        try:
            from ..core.preview_html import build_graphic_sheet_list
            rows = build_graphic_sheet_list(self.project)
        except Exception:
            return "Не удалось прочитать штампы листов."
        if rows:
            return (f"Распознано листов по штампу: {len(rows)}. "
                    "Листы без названия или номера в своих квадратах в "
                    "перечень не попадают.")
        return ("Штампы не распознаны: ни у одного листа нет названия и номера "
                "в заданных квадратах.")

    def _update_result(self) -> None:
        if self.project is None:
            self.lbl_result.setText("—")
            return
        self.lbl_result.setText(self.project.designation(self.section_kind or ""))

    # ---- правка ----
    def _on_code_typing(self, text: str) -> None:
        """Идёт ввод: показываем будущее обозначение, но в проект не пишем.

        Сохранение — только по снятию фокуса со строки (editingFinished):
        иначе каждая буква пересобирала предпросмотр, и вид «перескакивал»
        на начало тома.
        """
        if self.project is None or self.section_kind is None:
            self.lbl_result.setText("—")
            return
        base = (self.project.requisites.volume_id or "").strip()
        code = text.strip()
        self.lbl_result.setText(f"{base}{code}" if base else code)

    def _on_code_committed(self) -> None:
        """Ввод завершён (фокус ушёл со строки) — сохраняем доп. обозначение."""
        if self.project is None or self.section_kind is None:
            return
        code = self.edit_code.text().strip()
        if code != self.edit_code.text():
            self.edit_code.blockSignals(True)
            self.edit_code.setText(code)
            self.edit_code.blockSignals(False)
        if self.project.section_code(self.section_kind) == code:
            self._update_result()
            return
        self.project.section_codes[self.section_kind] = code
        self._update_result()
        self.changed.emit()

    def _apply_list_settings(self, **kwargs) -> None:
        """Пишется настройка перечня в модель проекта."""
        if self.project is None:
            return
        from ..core.preview_html import _graphic_sheet_list_cfg
        gl = _graphic_sheet_list_cfg(self.project)
        for key, value in kwargs.items():
            if hasattr(gl, key):
                setattr(gl, key, value)
        self.project.graphic_sheet_list = gl
        self._update_list_state()
        self.changed.emit()

    def _on_list_changed(self, checked: bool) -> None:
        self._apply_list_settings(enabled=bool(checked))

    def _on_source_changed(self, _index: int) -> None:
        source = self.cmb_source.currentData() or SHEETLIST_PDF
        self._apply_list_settings(source=source)

    def _on_template_committed(self) -> None:
        """Шаблон обозначения листа сохранён (фокус ушёл со строки)."""
        if self.project is None:
            return
        from ..core.preview_html import _graphic_sheet_list_cfg
        text = self.edit_template.text().strip() or "{volume} л.{no}"
        if _graphic_sheet_list_cfg(self.project).designation_template == text:
            return
        self._apply_list_settings(designation_template=text)

    def _on_excel_committed(self) -> None:
        """Путь к книге Excel сохранён (фокус ушёл со строки)."""
        if self.project is None:
            return
        from ..core.preview_html import _graphic_sheet_list_cfg
        text = self.edit_excel.text().strip()
        if _graphic_sheet_list_cfg(self.project).excel_path == text:
            return
        self._apply_list_settings(excel_path=text)

    def _choose_excel(self) -> None:
        """Выбор файла Excel с перечнем листов."""
        start = self.edit_excel.text().strip() or ""
        path, _ = QFileDialog.getOpenFileName(
            self, "Файл перечня листов", start,
            "Книга Excel (*.xlsx *.xlsm)")
        if not path:
            return
        self.edit_excel.setText(path)
        self._apply_list_settings(excel_path=path,
                                  source=SHEETLIST_EXCEL)
        index = self.cmb_source.findData(SHEETLIST_EXCEL)
        if index >= 0:
            self.cmb_source.blockSignals(True)
            self.cmb_source.setCurrentIndex(index)
            self.cmb_source.blockSignals(False)
        self._update_list_state()
