# -*- coding: utf-8 -*-
"""Контекстная панель настроек (правая часть окна).

Заменяет вкладки-формы. Показывает настройки, относящиеся к выбранному пункту
в дереве «Листы тома» (слева):
  • обложка / титульный лист  — форма обложки/титула;
  • содержание / ПЗ / графика — реквизиты тома;
  • приложения (раздел)       — список приложений;
  • лист регистрации изменений— журнал изменений;
  • документ                  — параметры конкретного документа;
  • приложение (пункт)        — параметры конкретного приложения.

Реестр подписантов вынесен в отдельную кнопку «Подписанты…», открывающую
модальное окно редактирования (SignersDialog).
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QStackedWidget, QPushButton,
    QGroupBox, QFormLayout, QLabel, QCheckBox, QDialog, QFileDialog,
    QMessageBox,
)
from PySide6.QtCore import Signal

from ..models import VolumeProject, SECTION_APPENDIX, SECTION_PZ, SECTION_GRAPHIC, DocKind
from .requisites_form import RequisitesForm
from .cover_title_form import CoverTitleForm
from .changes_form import ChangesForm
from .appendices_form import AppendicesForm
from .signers_form import SignersForm, SignersDialog


class SectionDocsWidget(QWidget):
    """Свойства раздела с документами (ПЗ / Графика): реквизиты + добавление документов."""

    item_modified = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.project: VolumeProject | None = None
        self.section_kind: str | None = None
        v = QVBoxLayout(self)
        self.requisites = RequisitesForm()
        # Настройка графы 9 влияет на все листы со штампом — обновляем предпросмотр.
        self.requisites.cell9_changed.connect(self._on_cell9_changed)
        # Оглавление ПЗ добавляет листы записки — обновляем предпросмотр.
        self.requisites.toc_changed.connect(self.item_modified)
        v.addWidget(self.requisites)

        grp = QGroupBox("Документы раздела")
        f = QVBoxLayout(grp)
        self.lbl_docs = QLabel("—")
        self.btn_add_doc = QPushButton("Добавить документ…")
        self.btn_add_doc.clicked.connect(self._add_doc)
        row = QHBoxLayout()
        row.addWidget(self.btn_add_doc)
        row.addStretch(1)
        f.addWidget(self.lbl_docs)
        f.addLayout(row)
        v.addWidget(grp)
        v.addStretch(1)

    def set_project(self, project: VolumeProject) -> None:
        self.project = project

    def _on_cell9_changed(self) -> None:
        """Настройка графы 9 изменена — переносим её в модель и обновляем листы."""
        if self.project is not None:
            self.requisites.apply_cell9(getattr(self.project, "stamp_cell9", None))
        self.item_modified.emit()

    def show(self, project: VolumeProject, section_kind: str) -> None:
        self.project = project
        self.section_kind = section_kind
        self.requisites.load(project.requisites, getattr(project, "stamp_cell9", None))
        kind = DocKind(section_kind)
        docs = project.documents_for(kind)
        if docs:
            self.lbl_docs.setText("\n".join(
                (os.path.basename(d.path) if d.path else "—") for d in docs))
        else:
            self.lbl_docs.setText("Документы отсутствуют")

    def _add_doc(self) -> None:
        if self.project is None or self.section_kind is None:
            return
        kind = DocKind(self.section_kind)
        files, _ = QFileDialog.getOpenFileNames(
            self, "Добавить документы", "",
            "Документы (*.docx *.doc *.pdf);;Все файлы (*)")
        if not files:
            return
        for f in files:
            from ..models import SourceFile
            sf = SourceFile(path=f, doc_kind=kind)
            try:
                from ..core.normalizer import normalize_file
                ni = normalize_file(f, force_kind=kind)
                sf.pages = ni.pages
                if ni.format:
                    sf.size = (ni.format.width_mm, ni.format.height_mm)
            except Exception:
                pass
            self.project.documents.append(sf)
        self.item_modified.emit()
        self.show(self.project, self.section_kind)


class DocInfoWidget(QWidget):
    """Настройки конкретного документа (ПЗ или графический лист)."""

    changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._doc = None
        # Проект нужен для расчёта формы первого листа документа (pz_doc_first_form);
        # проставляется контекстной панелью при загрузке проекта/показе формы.
        self.project = None
        v = QVBoxLayout(self)
        grp = QGroupBox("Документ")
        f = QFormLayout(grp)
        self.lbl_path = QLabel("—")
        self.lbl_kind = QLabel("—")
        self.lbl_pages = QLabel("—")
        self.chk_needed = QCheckBox("Включать в сборку")
        self.chk_needed.toggled.connect(self._on_needed)
        # Кнопка выбора файла (загрузка выделенной части).
        self.btn_file = QPushButton("Выбрать файл…")
        self.btn_file.clicked.connect(self._pick_file)
        # Редактор текста пояснительной записки по листам (только для DOCX).
        self.btn_edit = QPushButton("Редактировать ПЗ…")
        self.btn_edit.setToolTip(
            "Открыть все листы пояснительной записки и отредактировать их текст")
        self.btn_edit.clicked.connect(self._edit_pz)
        file_row = QHBoxLayout()
        file_row.addWidget(self.lbl_path, stretch=1)
        file_row.addWidget(self.btn_file)
        self.edit_row = QHBoxLayout()
        self.edit_row.addWidget(self.btn_edit)
        self.edit_row.addStretch(1)
        f.addRow("Файл:", file_row)
        f.addRow("Тип:", self.lbl_kind)
        f.addRow("Листы:", self.lbl_pages)
        f.addRow("", self.edit_row)
        f.addRow(self.chk_needed)
        v.addWidget(grp)
        v.addStretch(1)

    def load(self, info: dict) -> None:
        self._doc = info.get("obj")
        self.lbl_path.setText(info.get("path", "—") or "—")
        self.lbl_kind.setText(info.get("section_label", ""))
        pages = getattr(self._doc, "pages", None) if self._doc else None
        self.lbl_pages.setText(str(pages) if pages else "—")
        self.chk_needed.blockSignals(True)
        self.chk_needed.setChecked(bool(getattr(self._doc, "needed", True)))
        self.chk_needed.blockSignals(False)
        # Редактор доступен только для редактируемого DOCX (не PDF).
        path = getattr(self._doc, "path", "") or ""
        is_docx = path.lower().endswith((".docx", ".doc")) and os.path.exists(path)
        self.btn_edit.setEnabled(bool(is_docx))
        if not is_docx:
            self.btn_edit.setToolTip(
                "Редактирование доступно только для исходного документа Word (DOCX)")

    def _edit_pz(self) -> None:
        """Открывает редактор листов пояснительной записки."""
        path = getattr(self._doc, "path", "") or ""
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "Файл не найден",
                                "Сначала выберите существующий DOCX пояснительной записки.")
            return
        if not path.lower().endswith((".docx", ".doc")):
            QMessageBox.information(
                self, "Формат не поддерживается",
                "Редактирование доступно только для исходного документа Word (DOCX).")
            return
        from .pz_editor import PzEditorDialog
        from ..core.preview_html import pz_doc_first_form
        dlg = PzEditorDialog(path, self,
                             first_form=pz_doc_first_form(self.project, path))
        if dlg.exec() == QDialog.Accepted:
            # Число листов могло измениться — пересчитываем и обновляем предпросмотр.
            # Используется та же пагинация, что и в предпросмотре/сборке.
            try:
                from ..core.preview_html import _doc_sheet_count
                self._doc.pages = max(1, _doc_sheet_count(
                    self._doc,
                    first_form=pz_doc_first_form(self.project, path)))
                self.lbl_pages.setText(str(self._doc.pages))
            except Exception:
                pass
            self.changed.emit()

    def _pick_file(self) -> None:
        if self._doc is None:
            return
        kind = getattr(self._doc, "doc_kind", None)
        # Для ПЗ — Word/PDF, для графики — PDF.
        if kind is not None and getattr(kind, "value", "") == "Графическая часть":
            filt = "Документы (*.pdf);;Все файлы (*)"
        else:
            filt = "Документы (*.docx *.doc *.pdf);;Все файлы (*)"
        f, _ = QFileDialog.getOpenFileName(self, "Выбрать файл документа", "", filt)
        if f:
            self._doc.path = f
            self.lbl_path.setText(f)
            # Пересчитать листы.
            try:
                from ..core.normalizer import normalize_file
                from ..models import DocKind
                ni = normalize_file(f, force_kind=self._doc.doc_kind)
                self._doc.pages = ni.pages
                if ni.format:
                    self._doc.size = (ni.format.width_mm, ni.format.height_mm)
                self.lbl_pages.setText(str(self._doc.pages))
            except Exception:
                pass
            self.changed.emit()

    def _on_needed(self, checked: bool) -> None:
        if self._doc is not None:
            self._doc.needed = checked
            self.changed.emit()


class AppendixInfoWidget(QWidget):
    """Настройки конкретного приложения."""

    changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._app = None
        v = QVBoxLayout(self)
        grp = QGroupBox("Приложение")
        f = QFormLayout(grp)
        self.lbl_number = QLabel("—")
        self.lbl_name = QLabel("—")
        self.lbl_path = QLabel("—")
        self.lbl_pages = QLabel("—")
        self.chk_needed = QCheckBox("Включать в сборку")
        self.chk_needed.toggled.connect(self._on_needed)
        self.btn_file = QPushButton("Выбрать файл…")
        self.btn_file.clicked.connect(self._pick_file)
        file_row = QHBoxLayout()
        file_row.addWidget(self.lbl_path, stretch=1)
        file_row.addWidget(self.btn_file)
        f.addRow("Номер:", self.lbl_number)
        f.addRow("Наименование:", self.lbl_name)
        f.addRow("Файл:", file_row)
        f.addRow("Листы:", self.lbl_pages)
        f.addRow(self.chk_needed)
        v.addWidget(grp)
        v.addStretch(1)

    def load(self, info: dict) -> None:
        self._app = info.get("obj")
        self.lbl_number.setText(info.get("number", "") or "—")
        self.lbl_name.setText(info.get("name", "") or "—")
        self.lbl_path.setText(info.get("path", "—") or "—")
        pages = getattr(self._app, "pages", None) if self._app else None
        self.lbl_pages.setText(str(pages) if pages else "—")
        self.chk_needed.blockSignals(True)
        self.chk_needed.setChecked(bool(getattr(self._app, "needed", True)))
        self.chk_needed.blockSignals(False)

    def _pick_file(self) -> None:
        if self._app is None:
            return
        f, _ = QFileDialog.getOpenFileName(self, "Выбрать PDF приложения", "", "PDF (*.pdf)")
        if f:
            self._app.file_path = f
            self.lbl_path.setText(f)
            try:
                from ..core.pdf_reader import compute_pages
                self._app.pages = compute_pages(f)
                self.lbl_pages.setText(str(self._app.pages))
            except Exception:
                pass
            self.changed.emit()

    def _on_needed(self, checked: bool) -> None:
        if self._app is not None:
            self._app.needed = checked
            self.changed.emit()


class ContextPanel(QWidget):
    """Контекстная панель настроек: стек форм + кнопка «Подписанты…»."""

    item_modified = Signal()   # изменился параметр документа/приложения (needed)

    def __init__(self, project: VolumeProject | None = None,
                 signers_form: SignersForm | None = None, parent=None):
        super().__init__(parent)
        self.project = project
        self.signers_form = signers_form or SignersForm()

        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        # ---- Шапка: кнопка «Подписанты…» (вместо вкладки) ----
        head = QHBoxLayout()
        head.setContentsMargins(8, 8, 8, 8)
        self.btn_signers = QPushButton("Подписанты…")
        self.btn_signers.clicked.connect(self._open_signers)
        head.addWidget(self.btn_signers)
        head.addStretch(1)
        v.addLayout(head)

        # ---- Стек контекстных форм ----
        self.stack = QStackedWidget()
        self.requisites = RequisitesForm()
        # Настройка графы 9 (текст/логотип) влияет на все листы со штампом,
        # поэтому любое её изменение сразу пишется в модель и обновляет
        # предпросмотр (см. _on_context_cell9_changed).
        self.requisites.cell9_changed.connect(self._on_context_cell9_changed)
        # Автоматическое оглавление ПЗ: пересобираем листы после переключения.
        self.requisites.toc_changed.connect(self.item_modified)
        self.cover_title = CoverTitleForm(
            cover=(project.cover_sheet if project else None),
            title=(project.title_sheet if project else None))
        self.changes = ChangesForm(get_signers=lambda: self.signers_form.get_signers())
        self.appendices = AppendicesForm()
        self.doc_info = DocInfoWidget()
        self.doc_info.project = project
        self.doc_info.changed.connect(self.item_modified)
        self.appendix_info = AppendixInfoWidget()
        self.appendix_info.changed.connect(self.item_modified)
        self.section_docs = SectionDocsWidget()
        self.section_docs.item_modified.connect(self.item_modified)

        self.stack.addWidget(self.requisites)        # 0 — реквизиты (по умолчанию)
        self.stack.addWidget(self.cover_title)       # 1 — обложка/титул
        self.stack.addWidget(self.changes)           # 2 — изменения
        self.stack.addWidget(self.appendices)        # 3 — приложения
        self.stack.addWidget(self.doc_info)          # 4 — документ
        self.stack.addWidget(self.appendix_info)     # 5 — приложение (пункт)
        self.stack.addWidget(self.section_docs)      # 6 — раздел с документами (ПЗ/Графика)
        v.addWidget(self.stack, stretch=1)

    # ---- Загрузка / сбор данных ----
    def load_project(self, project: VolumeProject) -> None:
        self.project = project
        self.doc_info.project = project
        self.requisites.load(project.requisites, getattr(project, "stamp_cell9", None))
        self.cover_title._load()
        self.changes.set_records(project.changes)
        self.appendices.set_appendices(project.appendices)
        self.signers_form.set_signers(project.signers)

    def collect(self, project: VolumeProject) -> None:
        self.requisites.apply(project.requisites, getattr(project, "stamp_cell9", None))
        self.cover_title.apply_cover(project.cover_sheet)
        self.cover_title.apply_title(project.title_sheet)
        project.changes = self.changes.get_records()
        project.appendices = self.appendices.get_appendices()
        project.signers = self.signers_form.get_signers()

    # ---- Показ контекстной формы по выбранному пункту ----
    def show_context(self, selection: dict) -> None:
        kind = selection.get("kind")
        if kind in ("cover", "title"):
            self.stack.setCurrentWidget(self.cover_title)
        elif kind == "content":
            self.stack.setCurrentWidget(self.requisites)
        elif kind == "section":
            sec = selection.get("section")
            if sec == SECTION_APPENDIX:
                self.stack.setCurrentWidget(self.appendices)
            elif sec in (SECTION_PZ, SECTION_GRAPHIC):
                self.section_docs.show(self.project, sec)
                self.stack.setCurrentWidget(self.section_docs)
            else:
                self.stack.setCurrentWidget(self.requisites)
        elif kind == "changes":
            self.stack.setCurrentWidget(self.changes)
        elif kind == "appendix":
            self.appendix_info.load(selection)
            self.stack.setCurrentWidget(self.appendix_info)
        elif kind == "doc":
            self.doc_info.project = self.project
            self.doc_info.load(selection)
            self.stack.setCurrentWidget(self.doc_info)
        else:
            self.stack.setCurrentWidget(self.requisites)

    def _on_context_cell9_changed(self) -> None:
        """Настройка графы 9 изменена на форме реквизитов (стек 0)."""
        if self.project is not None:
            self.requisites.apply_cell9(getattr(self.project, "stamp_cell9", None))
        self.item_modified.emit()

    # ---- Подписанты ----
    def _open_signers(self) -> None:
        dlg = SignersDialog(self, self.signers_form)
        dlg.exec()
