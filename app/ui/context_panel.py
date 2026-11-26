# -*- coding: utf-8 -*-
"""Контекстная панель настроек (правая часть окна).

Показывает настройки, относящиеся к выбранному пункту в дереве «Листы тома»
(слева):
  • обложка / титульный лист  — форма обложки/титула;
  • содержание / ПЗ / графика — реквизиты тома (+ документы раздела для ПЗ и
    графики);
  • приложения (раздел)       — список приложений;
  • лист регистрации изменений— журнал изменений;
  • документ или приложение   — свойства элемента (одна форма на оба вида).

Реестр подписантов вынесен в отдельную кнопку «Подписанты…», открывающую
модальное окно редактирования (SignersDialog).
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QStackedWidget, QPushButton,
    QGroupBox, QFormLayout, QLabel, QCheckBox, QDialog, QFileDialog,
    QListWidget, QListWidgetItem, QMessageBox, QScrollArea, QFrame,
    QSizePolicy,
)
from PySide6.QtCore import Qt, Signal

from ..models import VolumeProject, SECTION_APPENDIX, SECTION_PZ, SECTION_GRAPHIC
from ..models import SECTION_CONTENT, SECTION_COMPOSITION, DocKind, SourceFile
from .requisites_form import RequisitesForm
from .cover_title_form import CoverTitleForm
from .changes_form import ChangesForm
from .appendices_form import AppendicesForm
from .section_settings_form import SectionSettingsForm
from .composition_form import CompositionForm
from .signers_form import SignersForm, SignersDialog
from .common import FileRow, refresh_document_meta, refresh_appendix_meta

DOC_FILTERS = "Документы (*.docx *.doc *.pdf);;Все файлы (*)"
PDF_FILTERS = "Документы (*.pdf);;Все файлы (*)"


class RequisitesDialog(QDialog):
    """Модальное окно реквизитов тома.

    Реквизиты — свойства тома, а не отдельного листа, поэтому в контекстной
    панели их нет: окно открывается кнопкой «Реквизиты тома…». Использует
    переданную форму, чтобы изменения сразу попадали в модель.
    """

    def __init__(self, parent=None, form: RequisitesForm | None = None):
        super().__init__(parent)
        self.setWindowTitle("Реквизиты тома")
        self.setMinimumSize(620, 560)
        v = QVBoxLayout(self)
        self.form = form or RequisitesForm()
        v.addWidget(self.form, stretch=1)

        btns = QHBoxLayout()
        ok = QPushButton("OK")
        ok.clicked.connect(self.accept)
        cancel = QPushButton("Отмена")
        cancel.clicked.connect(self.reject)
        btns.addStretch(1)
        btns.addWidget(ok)
        btns.addWidget(cancel)
        v.addLayout(btns)


class SectionDocsGroup(QGroupBox):
    """«Документы раздела»: список файлов с флажком «В сборку», добавление и удаление.

    Флажок управляет признаком ``needed``: снятый флажок исключает документ из
    сборки, не удаляя его из проекта. Показываются ВСЕ документы раздела, в том
    числе исключённые, — иначе исключённый документ нельзя было бы вернуть.

    Кнопка «Убрать из проекта» удаляет документ из списка совсем (файл на диске
    остаётся) — это необратимо, поэтому спрашивает подтверждение.
    """

    item_modified = Signal()

    def __init__(self, parent=None):
        super().__init__("Документы раздела", parent)
        self.project: VolumeProject | None = None
        self.section_kind: str | None = None
        self._loading = False
        v = QVBoxLayout(self)
        self.list = QListWidget()
        self.list.setToolTip(
            "Флажок «В сборку» исключает документ из тома, оставляя его в "
            "проекте; кнопка «Убрать из проекта» удаляет его из списка совсем")
        self.list.setMaximumHeight(120)
        self.list.setSelectionMode(QListWidget.ExtendedSelection)
        self.list.itemChanged.connect(self._on_item_changed)
        self.lbl_empty = QLabel("Документы отсутствуют")
        self.lbl_empty.setStyleSheet("color:#666;")
        self.btn_add_doc = QPushButton("Добавить документ…")
        self.btn_add_doc.clicked.connect(self._add_doc)
        self.btn_del_doc = QPushButton("Убрать из проекта")
        self.btn_del_doc.setToolTip(
            "Удалить выбранные документы из проекта (файлы на диске не удаляются)")
        self.btn_del_doc.clicked.connect(self._remove_docs)
        row = QHBoxLayout()
        row.addWidget(self.btn_add_doc)
        row.addWidget(self.btn_del_doc)
        row.addStretch(1)
        v.addWidget(self.lbl_empty)
        v.addWidget(self.list)
        v.addLayout(row)

    def show_section(self, project: VolumeProject, section_kind: str) -> None:
        self.project = project
        self.section_kind = section_kind
        kind = DocKind(section_kind)
        # Все документы раздела, включая исключённые из сборки.
        docs = [d for d in project.documents if d.doc_kind == kind]
        self._loading = True
        self.list.clear()
        for d in docs:
            name = os.path.basename(d.path) if d.path else "—"
            item = QListWidgetItem(name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if d.needed else Qt.Unchecked)
            item.setToolTip(d.path or "")
            item.setData(Qt.UserRole, d)
            if not d.needed:
                item.setForeground(Qt.gray)
            self.list.addItem(item)
        self._loading = False
        self.list.setVisible(bool(docs))
        self.lbl_empty.setVisible(not docs)

    def _on_item_changed(self, item) -> None:
        """Флажок переключён: пишем needed в модель и обновляем сборку."""
        if self._loading:
            return
        doc = item.data(Qt.UserRole)
        if doc is None:
            return
        doc.needed = item.checkState() == Qt.Checked
        item.setForeground(Qt.black if doc.needed else Qt.gray)
        self.item_modified.emit()

    def _add_doc(self) -> None:
        if self.project is None or self.section_kind is None:
            return
        kind = DocKind(self.section_kind)
        files, _ = QFileDialog.getOpenFileNames(
            self, "Добавить документы", "",
            "Документы (*.docx *.doc *.pdf);;Все файлы (*)")
        if not files:
            return
        for path in files:
            source = SourceFile(path=path, doc_kind=kind)
            refresh_document_meta(source)
            self.project.documents.append(source)
        self.item_modified.emit()
        self.show_section(self.project, self.section_kind)

    def _remove_docs(self) -> None:
        """Убирает выбранные документы из проекта (не только из сборки)."""
        if self.project is None:
            return
        docs = [it.data(Qt.UserRole) for it in self.list.selectedItems()]
        docs = [d for d in docs if d is not None]
        if not docs:
            QMessageBox.information(
                self, "Документ не выбран",
                "Выделите в списке документ, который нужно убрать из проекта.")
            return
        names = "\n".join(os.path.basename(d.path) if d.path else "—" for d in docs)
        answer = QMessageBox.question(
            self, "Убрать из проекта",
            f"Убрать документы из списка раздела?\n\n{names}\n\n"
            f"Файлы на диске не удаляются, но из проекта они исчезнут — "
            f"вернуть их можно только повторным добавлением.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            return
        for d in docs:
            if d in self.project.documents:
                self.project.documents.remove(d)
        self.item_modified.emit()
        self.show_section(self.project, self.section_kind)


class ItemInfoWidget(QWidget):
    """Свойства выбранного элемента дерева: документ или приложение.

    Одна форма на оба вида: у приложения дополнительно видны строки
    «Номер» и «Наименование», у документа — кнопка редактора записки.
    """

    changed = Signal()
    edit_pz_requested = Signal(object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._obj = None
        self._is_appendix = False
        v = QVBoxLayout(self)
        self.grp = QGroupBox("Документ")
        self._form = QFormLayout(self.grp)
        self.lbl_number = QLabel("—")
        self.lbl_name = QLabel("—")
        self.file = FileRow("Выбрать файл документа", DOC_FILTERS)
        self.file.picked.connect(self._on_file_picked)
        self.lbl_pages = QLabel("—")
        # Включение в сборку — тот же признак needed, что и флажок в дереве.
        self.chk_needed = QCheckBox("Включать в сборку")
        self.chk_needed.setToolTip(
            "Снятый флажок исключает элемент из сборки тома; файл остаётся в проекте")
        self.chk_needed.toggled.connect(self._on_needed)
        self.btn_edit = QPushButton("Редактировать ПЗ…")
        self.btn_edit.setToolTip(
            "Открыть все листы пояснительной записки и отредактировать их текст")
        self.btn_edit.clicked.connect(
            lambda: self.edit_pz_requested.emit(self._obj))
        self._form.addRow("Номер:", self.lbl_number)
        self._form.addRow("Наименование:", self.lbl_name)
        self._form.addRow("Файл:", self.file)
        self._form.addRow("Листы:", self.lbl_pages)
        self._form.addRow(self.chk_needed)
        self._form.addRow("", self.btn_edit)
        v.addWidget(self.grp)
        v.addStretch(1)
        self.set_appendix_mode(False)

    # ---- вид формы ----
    def set_appendix_mode(self, is_appendix: bool) -> None:
        """Переключает форму между документом и приложением."""
        self._is_appendix = is_appendix
        self.grp.setTitle("Приложение" if is_appendix else "Документ")
        self._form.setRowVisible(0, is_appendix)   # Номер
        self._form.setRowVisible(1, is_appendix)   # Наименование
        self.btn_edit.setVisible(not is_appendix)
        self.file.set_filters(PDF_FILTERS if is_appendix else DOC_FILTERS)

    def load(self, info: dict) -> None:
        obj = info.get("obj")
        self._obj = obj
        self.set_appendix_mode(info.get("kind") == "appendix")
        if self._is_appendix:
            self.lbl_number.setText(info.get("number", "") or "—")
            self.lbl_name.setText(info.get("name", "") or "—")
            path = getattr(obj, "file_path", "") or ""
        else:
            path = getattr(obj, "path", "") or ""
            kind = getattr(obj, "doc_kind", None)
            if getattr(kind, "value", "") == SECTION_GRAPHIC:
                self.file.set_filters(PDF_FILTERS)
        self.file.set_path(path)
        pages = getattr(obj, "pages", None) if obj else None
        self.lbl_pages.setText(str(pages) if pages else "—")
        self.btn_edit.setEnabled(self._is_docx(path))
        self.chk_needed.blockSignals(True)
        self.chk_needed.setChecked(bool(getattr(obj, "needed", True)))
        self.chk_needed.blockSignals(False)

    def _on_needed(self, checked: bool) -> None:
        """Флажок «включать в сборку» переключён."""
        if self._obj is not None:
            self._obj.needed = bool(checked)
            self.changed.emit()

    @staticmethod
    def _is_docx(path: str) -> bool:
        return bool(path) and path.lower().endswith((".docx", ".doc")) \
            and os.path.exists(path)

    # ---- правки ----
    def _on_file_picked(self, path: str) -> None:
        if not path or self._obj is None:
            return
        if self._is_appendix:
            self._obj.file_path = path
            refresh_appendix_meta(self._obj)
        else:
            self._obj.path = path
            refresh_document_meta(self._obj)
            self.btn_edit.setEnabled(self._is_docx(path))
        self.lbl_pages.setText(
            str(self._obj.pages) if self._obj.pages else "—")
        self.changed.emit()


class ContextPanel(QWidget):
    """Контекстная панель настроек: стек форм + кнопка «Подписанты…»."""

    item_modified = Signal()            # изменился параметр документа/приложения
    edit_pz_requested = Signal(object)  # запрошен редактор листов записки

    def __init__(self, project: VolumeProject | None = None,
                 signers_form: SignersForm | None = None, parent=None):
        super().__init__(parent)
        self.project = project
        self.signers_form = signers_form or SignersForm()

        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        # ---- Шапка: реквизиты тома и реестр подписантов ----
        head = QHBoxLayout()
        head.setContentsMargins(8, 8, 8, 8)
        self.btn_requisites = QPushButton("Реквизиты тома…")
        self.btn_requisites.setToolTip(
            "Организация, обозначение тома, объект, стадия, город, год, "
            "графа 9 и оглавление записки")
        self.btn_requisites.clicked.connect(self._open_requisites)
        self.btn_signers = QPushButton("Подписанты…")
        self.btn_signers.clicked.connect(self._open_signers)
        head.addWidget(self.btn_requisites)
        head.addWidget(self.btn_signers)
        head.addStretch(1)
        v.addLayout(head)

        # ---- Стек контекстных форм ----
        # Формы бывают высокими (таблица состава ПД с кнопками): в прокрутке
        # они не обрезаются, даже если панель узкая или низкая.
        self.stack = QStackedWidget()
        # Стек не должен раздвигать панель по самой широкой форме: иначе узкая
        # панель получает горизонтальную прокрутку. Ширину задаёт панель.
        self.stack.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setWidget(self.stack)
        self.requisites = RequisitesForm()
        # Настройка графы 9 (текст/логотип) влияет на все листы со штампом,
        # поэтому любое её изменение сразу пишется в модель и обновляет
        # предпросмотр.
        self.requisites.cell9_changed.connect(self._on_cell9_changed)
        # Автоматическое оглавление ПЗ: пересобираем листы после переключения.
        self.requisites.toc_changed.connect(self.item_modified)

        # Настройки раздела (обозначение) и документы раздела — ОДНА страница:
        # для «Содержания» блок документов скрыт, для ПЗ и графики — показан.
        self.docs_group = SectionDocsGroup()
        self.docs_group.item_modified.connect(self.item_modified)
        self.docs_group.setVisible(False)
        self.section_settings = SectionSettingsForm()
        self.section_settings.changed.connect(self.item_modified)
        self.section_page = QWidget()
        page = QVBoxLayout(self.section_page)
        page.setContentsMargins(0, 0, 0, 0)
        page.setSpacing(4)
        page.addWidget(self.section_settings)
        page.addWidget(self.docs_group)
        page.addStretch(1)

        self.cover_title = CoverTitleForm(
            cover=(project.cover_sheet if project else None),
            title=(project.title_sheet if project else None))
        self.changes = ChangesForm(get_signers=lambda: self.signers_form.get_signers())
        self.appendices = AppendicesForm()
        # Флажок «В сборку» в списке приложений меняет состав тома.
        self.appendices.items_changed.connect(self.item_modified)
        # Состав проектной документации: таблица томов (вручную или из Excel).
        self.composition = CompositionForm()
        self.composition.changed.connect(self.item_modified)
        self.item_info = ItemInfoWidget()
        self.item_info.changed.connect(self.item_modified)
        self.item_info.edit_pz_requested.connect(self.edit_pz_requested)

        self.stack.addWidget(self.section_page)      # 0 — раздел (обозначение + документы)
        self.stack.addWidget(self.cover_title)       # 1 — обложка/титул
        self.stack.addWidget(self.appendices)        # 2 — приложения
        self.stack.addWidget(self.changes)           # 3 — изменения
        self.stack.addWidget(self.item_info)         # 4 — документ/приложение
        self.stack.addWidget(self.composition)       # 5 — состав проектной документации
        v.addWidget(self.scroll, stretch=1)

    # ---- Загрузка / сбор данных ----
    def load_project(self, project: VolumeProject) -> None:
        self.project = project
        self.requisites.load(project.requisites,
                             getattr(project, "stamp_cell9", None))
        self.cover_title._load()
        self.changes.set_records(project.changes)
        self.appendices.set_appendices(project.appendices)
        self.composition.show_project(project)
        self.signers_form.set_signers(project.signers)

    def collect(self, project: VolumeProject) -> None:
        self.requisites.apply(project.requisites,
                              getattr(project, "stamp_cell9", None))
        self.cover_title.apply_cover(project.cover_sheet)
        self.cover_title.apply_title(project.title_sheet)
        project.changes = self.changes.get_records()
        project.appendices = self.appendices.get_appendices()
        project.signers = self.signers_form.get_signers()

    # ---- Показ контекстной формы по выбранному пункту ----
    def show_context(self, selection: dict) -> None:
        kind = selection.get("kind")
        if kind in ("cover", "title"):
            # На обложке — настройки обложки, на титуле — настройки титула.
            self.cover_title.show_part(kind)
            self.stack.setCurrentWidget(self.cover_title)
        elif kind == "changes":
            self.stack.setCurrentWidget(self.changes)
        elif kind == "appendix" or kind == "doc":
            self.item_info.load(selection)
            self.stack.setCurrentWidget(self.item_info)
        elif kind == "section" and selection.get("section") == SECTION_APPENDIX:
            self.stack.setCurrentWidget(self.appendices)
        elif kind == "section" and selection.get("section") == SECTION_COMPOSITION:
            # Состав проектной документации: таблица томов.
            self.composition.show_project(self.project)
            self.stack.setCurrentWidget(self.composition)
        else:
            # Содержание, ПЗ и графика: обозначение раздела; для ПЗ и графики
            # дополнительно показывается блок документов раздела.
            sec = (selection.get("section") if kind == "section" else None) or kind
            with_docs = sec in (SECTION_PZ, SECTION_GRAPHIC)
            self.docs_group.setVisible(with_docs)
            if with_docs and self.project is not None:
                self.docs_group.show_section(self.project, sec)
            self.section_settings.show_section(self.project, sec)
            self.stack.setCurrentWidget(self.section_page)

    def _on_cell9_changed(self) -> None:
        """Настройка графы 9 изменена — переносим её в модель."""
        if self.project is not None:
            self.requisites.apply_cell9(getattr(self.project, "stamp_cell9", None))
        self.item_modified.emit()

    # ---- Реквизиты тома и подписанты (модальные окна) ----
    def _open_requisites(self) -> None:
        """Окно реквизитов тома: основные реквизиты, графа 9, оглавление ПЗ."""
        dlg = RequisitesDialog(self, self.requisites)
        # Код результата берётся у КЛАССА: у экземпляра диалога атрибута
        # Accepted нет, `dlg.Accepted` давал AttributeError.
        if dlg.exec() == QDialog.Accepted and self.project is not None:
            self.requisites.apply(self.project.requisites,
                                  getattr(self.project, "stamp_cell9", None))
            self.item_modified.emit()

    def _open_signers(self) -> None:
        dlg = SignersDialog(self, self.signers_form)
        dlg.exec()
