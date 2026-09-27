# -*- coding: utf-8 -*-
"""Окно редактирования текста пояснительной записки (DOCX) по листам.

Документ разбивается на листы тем же алгоритмом, что и HTML-предпросмотр
(app.core.preview_html._paginate_docx), поэтому список листов в редакторе
совпадает с листами предпросмотра и сборки. Слева — перечень всех листов,
справа — текст выбранного листа.

При сохранении изменяется текст абзацев исходного DOCX. Таблицы, изображения,
колонтитулы и разметка страниц не удаляются: нетронутые абзацы (например,
внутри таблиц) остаются как есть.
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPlainTextEdit, QPushButton,
    QMessageBox, QListWidget, QListWidgetItem, QSplitter,
)
from PySide6.QtCore import Qt


class PzEditorDialog(QDialog):
    """Постраничный редактор текста исходного DOCX пояснительной записки."""

    def __init__(self, path: str, parent=None, first_form=2):
        super().__init__(parent)
        self.path = path
        # Форма штампа первого листа потока документа в разделе (2 / "2a") —
        # по ней считается бюджет первого листа, чтобы разбивка редактора
        # совпадала с предпросмотром (см. preview_html.pz_doc_first_form).
        self.first_form = first_form
        self._document = None
        # Список индексов абзацев DOCX для каждого листа.
        self._pages: list[list[int]] = []
        # Текущий текст каждого листа (редактируемый).
        self._page_text: list[str] = []
        self._current_row = -1

        self.setWindowTitle(
            f"Редактирование пояснительной записки — {os.path.basename(path)}")
        self.resize(1100, 760)

        root = QVBoxLayout(self)
        hint = QLabel(
            "Выберите лист слева и отредактируйте текст справа. "
            "Сохраняются текстовые абзацы; таблицы и изображения не изменяются.")
        hint.setWordWrap(True)
        root.addWidget(hint)

        split = QSplitter(Qt.Horizontal)
        self.pages_list = QListWidget()
        self.pages_list.setMinimumWidth(220)
        self.pages_list.currentRowChanged.connect(self._show_page)
        split.addWidget(self.pages_list)

        self.editor = QPlainTextEdit()
        self.editor.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        split.addWidget(self.editor)
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([240, 820])
        root.addWidget(split, stretch=1)

        self.status = QLabel("—")
        root.addWidget(self.status)

        buttons = QHBoxLayout()
        self.btn_prev = QPushButton("← Предыдущий лист")
        self.btn_next = QPushButton("Следующий лист →")
        self.btn_cancel = QPushButton("Отмена")
        self.btn_save = QPushButton("Сохранить в файл")
        self.btn_prev.clicked.connect(lambda: self._select_relative(-1))
        self.btn_next.clicked.connect(lambda: self._select_relative(1))
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_save.clicked.connect(self._save)
        buttons.addWidget(self.btn_prev)
        buttons.addWidget(self.btn_next)
        buttons.addStretch(1)
        buttons.addWidget(self.btn_cancel)
        buttons.addWidget(self.btn_save)
        root.addLayout(buttons)

        self._load()

    # ---- Загрузка ----
    def _load(self) -> None:
        try:
            from docx import Document
            from ..core.docx_reader import read_docx
            from ..core.preview_html import _paginate_docx

            self._document = Document(self.path)
            content = read_docx(self.path)
            # Путь нужен для фактических ширин колонок таблиц — иначе листы в
            # редакторе не совпали бы с листами предпросмотра.
            pages = _paginate_docx(content, self.path,
                                   first_form=self.first_form) or [[]]

            # Индексы непустых абзацев в порядке следования: именно они попадают
            # в поток пагинации (пустые абзацы алгоритм пропускает).
            nonempty = [i for i, p in enumerate(content.paragraphs) if p.text.strip()]
            cursor = 0
            for page in pages:
                indices = []
                for item in page:
                    if item[0] == "line" and cursor < len(nonempty):
                        indices.append(nonempty[cursor])
                        cursor += 1
                self._pages.append(indices)
            # Абзацы, не попавшие в оценку пагинации, показываем на последнем листе.
            if cursor < len(nonempty):
                self._pages[-1].extend(nonempty[cursor:])
            if not self._pages:
                self._pages = [[]]

            paragraphs = self._document.paragraphs
            for number, indices in enumerate(self._pages, 1):
                item = QListWidgetItem(f"Лист {number} — {len(indices)} абз.")
                self.pages_list.addItem(item)
                self._page_text.append(
                    "\n".join(paragraphs[i].text for i in indices))

            self.status.setText(
                f"Листов в записке: {len(self._pages)}. Источник: {self.path}")
            self.pages_list.setCurrentRow(0)
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка чтения",
                                 f"Не удалось открыть DOCX:\n{exc}")
            self.btn_save.setEnabled(False)
            self.editor.setEnabled(False)

    # ---- Переключение листов ----
    def _show_page(self, row: int) -> None:
        # Сохраняем правки предыдущего листа перед переключением.
        if 0 <= self._current_row < len(self._page_text):
            self._page_text[self._current_row] = self.editor.toPlainText()
        self._current_row = row
        if 0 <= row < len(self._page_text):
            self.editor.setPlainText(self._page_text[row])
        else:
            self.editor.setPlainText("")
        total = len(self._page_text)
        self.btn_prev.setEnabled(row > 0)
        self.btn_next.setEnabled(0 <= row < total - 1)
        if 0 <= row < total:
            self.status.setText(
                f"Лист {row + 1} из {total}. Источник: {self.path}")

    def _select_relative(self, delta: int) -> None:
        row = self.pages_list.currentRow() + delta
        if 0 <= row < self.pages_list.count():
            self.pages_list.setCurrentRow(row)

    def flush_current_page(self) -> None:
        """Переносит текст текущего листа из поля ввода в буфер редактора."""
        if 0 <= self._current_row < len(self._page_text):
            self._page_text[self._current_row] = self.editor.toPlainText()

    def collect_edits(self) -> dict:
        """Собирает правки абзацев как {индекс абзаца DOCX: новый текст}.

        Тот же формат использует интерактивный предпросмотр, поэтому запись
        выполняется единым модулем app.core.docx_edit.
        """
        edits: dict[int, str] = {}
        if self._document is None:
            return edits
        self.flush_current_page()
        paragraphs = list(self._document.paragraphs)
        for page_no, indices in enumerate(self._pages):
            lines = self._page_text[page_no].splitlines()
            for pos, index in enumerate(indices):
                if index >= len(paragraphs):
                    continue
                edits[index] = lines[pos] if pos < len(lines) else ""
        return edits

    def _save(self) -> None:
        if self._document is None:
            return
        try:
            from ..core.docx_edit import apply_docx_edits
            edits = self.collect_edits()
            # Изменёнными считаются только отличающиеся абзацы.
            current = list(self._document.paragraphs)
            edits = {i: t for i, t in edits.items()
                     if i < len(current) and current[i].text != t}
            # Окно редактора правит только абзацы; таблицы в нём не показываются
            # (они редактируются прямо в предпросмотре).
            apply_docx_edits(self.path, paragraph_edits=edits)
            QMessageBox.information(
                self, "Сохранено", f"Изменения записаны в файл:\n{self.path}")
            self.accept()
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка сохранения",
                                 f"Не удалось сохранить DOCX:\n{exc}")
