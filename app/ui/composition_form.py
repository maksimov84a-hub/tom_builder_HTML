# -*- coding: utf-8 -*-
"""Форма раздела «Состав проектной документации»: таблица томов.

Таблица заполняется вручную (правка ячеек прямо в форме) либо импортом из
книги Excel. Импорт КОПИРУЕТ строки в модель проекта, поэтому после загрузки
таблицу можно править, а сборка тома не зависит от наличия файла.

Колонки и их ширины — единый источник с листом (models.COMPOSITION_*):
Номер тома 15 | Обозначение 60 | Наименование 90 | Примечание 20 мм.
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QLabel, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox,
    QAbstractItemView, QAbstractScrollArea, QHeaderView,
)
from PySide6.QtCore import Signal

from ..models import (VolumeProject, COMPOSITION_COLS, COMPOSITION_LABELS,
                      SECTION_COMPOSITION)

#: Ширины колонок формы в пикселях — те же пропорции, что на листе (15/60/90/20 мм).
_COL_WIDTHS = (70, 170, 240, 110)


class CompositionForm(QWidget):
    """Таблица «Состав проектной документации»: ручной ввод и импорт из Excel."""

    #: Данные таблицы изменились (нужно перерисовать листы раздела).
    changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.project: VolumeProject | None = None
        self._loading = False      # защита от записи в модель при заполнении

        v = QVBoxLayout(self)
        v.setContentsMargins(8, 8, 8, 8)
        v.setSpacing(6)

        title = QLabel("Состав проектной документации")
        title.setStyleSheet("font-weight:bold;")
        v.addWidget(title)

        # Дополнительное обозначение раздела — как у остальных разделов:
        # обозначение тома + это дополнение выводится в графе 1 штампа.
        self.grp_code = QGroupBox("Обозначение раздела")
        code_layout = QVBoxLayout(self.grp_code)
        code_layout.setContentsMargins(8, 6, 8, 6)
        code_layout.setSpacing(4)
        self.edit_code = QLineEdit()
        self.edit_code.setPlaceholderText("например .СПД")
        self.edit_code.setToolTip(
            "Дополнительное обозначение раздела: дописывается к обозначению "
            "тома в графе 1 основной надписи листов")
        self.lbl_code_result = QLabel("—")
        self.lbl_code_result.setWordWrap(True)
        code_layout.addWidget(self.edit_code)
        code_layout.addWidget(self.lbl_code_result)
        v.addWidget(self.grp_code)
        self.edit_code.editingFinished.connect(self._on_code_committed)

        hint = QLabel(
            "Строки таблицы выводятся на листах с рамкой и штампом формы 5 "
            "(первый лист) и формы 6 (последующие): шапка 15 мм, строка 8 мм. "
            "Заполняйте вручную или загрузите книгу Excel (первая строка — "
            "шапка, далее по строке на том).")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#555;")
        v.addWidget(hint)
        self.lbl_hint = hint

        self.table = QTableWidget(0, len(COMPOSITION_LABELS))
        self.table.setHorizontalHeaderLabels(list(COMPOSITION_LABELS))
        for i, width in enumerate(_COL_WIDTHS):
            self.table.setColumnWidth(i, width)
        # Колонки растягиваются по ширине панели, а таблица не держит минимум по
        # ширине колонок: иначе форма становится шире панели и появляется
        # горизонтальная прокрутка.
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setMinimumSectionSize(30)
        self.table.setMinimumWidth(0)
        self.table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustIgnored)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.verticalHeader().setDefaultSectionSize(22)
        self.table.itemChanged.connect(self._on_item_changed)
        v.addWidget(self.table, stretch=1)

        self.btn_add = QPushButton("Добавить строку")
        self.btn_del = QPushButton("Удалить строку")
        self.btn_excel = QPushButton("Загрузить из Excel…")
        self.btn_clear = QPushButton("Очистить")
        self.btn_add.clicked.connect(self._add_row)
        self.btn_del.clicked.connect(self._del_rows)
        self.btn_excel.clicked.connect(self._load_excel)
        self.btn_clear.clicked.connect(self._clear)
        # Кнопки — в столбец во всю ширину панели: в две колонки подписи
        # обрезаются, в строку они не помещаются.
        buttons = QVBoxLayout()
        buttons.setContentsMargins(0, 0, 0, 0)
        buttons.setSpacing(4)
        for widget in (self.btn_add, self.btn_del, self.btn_excel, self.btn_clear):
            buttons.addWidget(widget)
        v.addLayout(buttons)

        # Правка ячейки завершена (Enter или уход фокуса) — только теперь
        # перерисовываем листы: побуквенное обновление сбивало бы ввод.
        self.table.itemDelegate().closeEditor.connect(
            lambda *_: self.changed.emit())

    # ---- загрузка / показ ----
    def show_project(self, project: VolumeProject | None) -> None:
        """Заполняет таблицу строками проекта."""
        self.project = project
        self._refresh()
        self._load_code()

    def _load_code(self) -> None:
        """Показывает доп. обозначение раздела из проекта."""
        code = ""
        if self.project is not None:
            code = (self.project.section_codes or {}).get(SECTION_COMPOSITION, "")
        self.edit_code.blockSignals(True)
        self.edit_code.setText(code)
        self.edit_code.blockSignals(False)
        self._update_code_result()

    def _update_code_result(self) -> None:
        """Показывает обозначение раздела: обозначение тома + доп. обозначение."""
        if self.project is None:
            self.lbl_code_result.setText("—")
            return
        self.lbl_code_result.setText(
            self.project.designation(SECTION_COMPOSITION) or "—")

    def _on_code_committed(self) -> None:
        """Доп. обозначение сохранено (фокус ушёл из строки)."""
        if self.project is None:
            return
        code = self.edit_code.text().strip()
        if code != self.edit_code.text():
            self.edit_code.blockSignals(True)
            self.edit_code.setText(code)
            self.edit_code.blockSignals(False)
        if self.project.section_code(SECTION_COMPOSITION) == code:
            self._update_code_result()
            return
        self.project.section_codes[SECTION_COMPOSITION] = code
        self._update_code_result()
        self.changed.emit()

    def _refresh(self) -> None:
        rows = (self.project.composition.rows if self.project is not None else [])
        self._loading = True
        try:
            self.table.setRowCount(0)
            for values in rows:
                self._insert_table_row(self.table.rowCount(), values)
        finally:
            self._loading = False
        self._update_hint()

    def _insert_table_row(self, index: int, values=None) -> None:
        values = list(values or [""] * len(COMPOSITION_COLS))
        self.table.insertRow(index)
        for col in range(len(COMPOSITION_COLS)):
            text = str(values[col] if col < len(values) else "")
            item = QTableWidgetItem(text)
            self.table.setItem(index, col, item)

    def _update_hint(self) -> None:
        count = len(self.project.composition.rows) if self.project is not None else 0
        self.lbl_hint.setText(
            "Строк в таблице: {n}. Таблица выводится на листах с рамкой и "
            "штампом формы 5 (первый лист) и формы 6 (последующие): шапка "
            "15 мм, строка 8 мм. Заполняйте вручную или загрузите книгу Excel "
            "(первая строка — шапка).".format(n=count))

    # ---- правка ----
    def _on_item_changed(self, item) -> None:
        """Значение ячейки записано в модель проекта."""
        if self._loading or self.project is None:
            return
        rows = self.project.composition.rows
        row, col = item.row(), item.column()
        while len(rows) <= row:
            rows.append({key: "" for key in COMPOSITION_COLS})
        rows[row][COMPOSITION_COLS[col]] = item.text().strip()

    def _add_row(self) -> None:
        if self.project is None:
            return
        self.project.composition.rows.append({key: "" for key in COMPOSITION_COLS})
        self._loading = True
        try:
            self._insert_table_row(self.table.rowCount())
        finally:
            self._loading = False
        self.table.setCurrentCell(self.table.rowCount() - 1, 0)
        self.table.editItem(self.table.currentItem())
        self._update_hint()
        self.changed.emit()

    def _del_rows(self) -> None:
        if self.project is None:
            return
        selected = sorted({index.row() for index in self.table.selectedIndexes()})
        if not selected:
            return
        self._loading = True
        try:
            for row in reversed(selected):
                if 0 <= row < len(self.project.composition.rows):
                    del self.project.composition.rows[row]
                self.table.removeRow(row)
        finally:
            self._loading = False
        self._update_hint()
        self.changed.emit()

    def _clear(self) -> None:
        if self.project is None or not self.project.composition.rows:
            return
        if QMessageBox.question(
                self, "Очистить таблицу",
                "Удалить все строки таблицы «Состав проектной документации»?"
        ) != QMessageBox.Yes:
            return
        self.project.composition.rows.clear()
        self._refresh()
        self.changed.emit()

    def _load_excel(self) -> None:
        """Импорт строк из книги Excel: строки копируются в модель."""
        if self.project is None:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Книга Excel со составом проектной документации", "",
            "Книга Excel (*.xlsx *.xlsm)")
        if not path:
            return
        if not os.path.exists(path):
            QMessageBox.warning(self, "Файл не найден", path)
            return
        from ..core.xlsx_reader import read_composition, sheet_names
        names = sheet_names(path)
        sheet = ""
        if len(names) > 1:
            sheet = self._choose_sheet(names)
            if sheet is None:
                return
        try:
            rows = read_composition(path, sheet)
        except Exception as e:
            QMessageBox.warning(self, "Не удалось прочитать файл",
                                f"{type(e).__name__}: {e}")
            return
        if not rows:
            QMessageBox.information(
                self, "Таблица не найдена",
                "В книге не найдено непустых строк: первая строка считается "
                "шапкой, пустые строки пропускаются.")
            return
        current = len(self.project.composition.rows)
        if current and QMessageBox.question(
                self, "Заменить строки",
                f"В таблице уже {current} строк. Заменить их строками из файла "
                f"({len(rows)})?") != QMessageBox.Yes:
            return
        self.project.composition.rows = [
            {key: (values[i].strip() if i < len(values) else "")
             for i, key in enumerate(COMPOSITION_COLS)}
            for values in rows
        ]
        self._refresh()
        self.changed.emit()

    def _choose_sheet(self, names: list) -> str | None:
        """Выбор листа книги, если их несколько."""
        from PySide6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getItem(
            self, "Лист книги", "Какой лист читать?", list(names), 0, False)
        return name if ok else None
