# -*- coding: utf-8 -*-
"""Дерево листов тома (левая панель) — редактор структуры.

Показывает разделы тома в настраиваемом порядке и их содержимое
(документы ПЗ/Графики, приложения). Позволяет:
  • переставлять разделы и документы местами (Вверх/Вниз);
  • отмечать необходимость (включение/исключение из сборки) чекбоксом;
  • добавлять/удалять документы и приложения;
  • выбирать лист для предпросмотра (сигнал selection_changed).

Отметка «необходимость» для раздела/документа/приложения хранится в модели
(needed / структура), поэтому исключённый элемент не попадает в сборку,
в «Листов M» и в содержание.
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTreeWidget, QTreeWidgetItem,
    QLabel, QMessageBox,
)
from PySide6.QtCore import Qt, Signal

from ..models import (
    VolumeProject, DocKind, SourceFile, Appendix,
    SECTION_COVER, SECTION_TITLE, SECTION_CONTENT, SECTION_PZ, SECTION_GRAPHIC,
    SECTION_APPENDIX, SECTION_CHANGES, SECTION_LABELS,
)

# Роли хранения данных в элементах дерева
_SEC_ROLE = Qt.UserRole          # ключ раздела (str)
_DOC_ROLE = Qt.UserRole + 1      # объект SourceFile
_APP_ROLE = Qt.UserRole + 2      # объект Appendix
_NODE_ROLE = Qt.UserRole + 3     # тип узла: "section" | "doc" | "appendix"


class SheetTree(QWidget):
    """Дерево структуры тома. Выбор листа -> сигнал selection_changed(dict)."""

    selection_changed = Signal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.project: VolumeProject | None = None
        self.setMinimumWidth(320)
        self._updating = False   # защита от реентерабельности при перестройке

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        header = QLabel("Листы тома")
        header.setStyleSheet(
            "font-size:13pt; font-weight:bold; padding:8px 10px; "
            "background:#eef2f7; border-bottom:1px solid #c9d2dc;"
        )
        lay.addWidget(header)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Раздел / лист", "Листы"])
        self.tree.setSelectionMode(QTreeWidget.SingleSelection)
        self.tree.setIndentation(18)
        self.tree.setFocusPolicy(Qt.NoFocus)
        # Ширина колонок: название раздела шире, листы — узкая колонка.
        self.tree.header().resizeSection(0, 220)
        self.tree.header().resizeSection(1, 55)
        self.tree.header().setStretchLastSection(False)
        f = self.tree.font()
        f.setPointSize(10)
        self.tree.setFont(f)
        self.tree.currentItemChanged.connect(self._on_select)
        self.tree.itemChanged.connect(self._on_item_changed)
        lay.addWidget(self.tree)

        # ---- Панель действий ----
        row = QHBoxLayout()
        self.btn_del = QPushButton("Удалить")
        self.btn_up = QPushButton("Вверх")
        self.btn_down = QPushButton("Вниз")
        self.btn_del.clicked.connect(self._delete)
        self.btn_up.clicked.connect(lambda: self._move(-1))
        self.btn_down.clicked.connect(lambda: self._move(1))
        row.addWidget(self.btn_del)
        row.addWidget(self.btn_up)
        row.addWidget(self.btn_down)
        row.addStretch(1)
        lay.addLayout(row)

    # ---- публичное ----
    def set_project(self, project: VolumeProject) -> None:
        self.project = project
        self._refresh()

    # ---- построение дерева ----
    def _refresh(self, select_key: str | None = None) -> None:
        """Перестраивает дерево. select_key — ключ раздела для восстановления выбора."""
        self._updating = True
        self.tree.clear()
        try:
            if self.project is None:
                return
            p = self.project
            # Отслеживаем, какой объект был выбран (для восстановления после перестройки).
            selected = self._current_ref()
            for entry in p.structure:
                kind = entry.get("kind")
                needed = bool(entry.get("needed", True))
                label = SECTION_LABELS.get(kind, kind)
                sec = QTreeWidgetItem([label, ""])
                sec.setData(0, _NODE_ROLE, "section")
                sec.setData(0, _SEC_ROLE, kind)
                sec.setCheckState(0, Qt.Checked if needed else Qt.Unchecked)
                # дочерние элементы
                if kind == SECTION_PZ or kind == SECTION_GRAPHIC:
                    dkind = DocKind(kind)
                    for d in p.documents_for(dkind):
                        sec.addChild(self._doc_item(d))
                elif kind == SECTION_APPENDIX:
                    for a in p.appendices:
                        sec.addChild(self._app_item(a))
                self.tree.addTopLevelItem(sec)
            self.tree.expandAll()
            self._restore_selection(selected)
        finally:
            self._updating = False

    def _doc_item(self, d: SourceFile) -> QTreeWidgetItem:
        name = os.path.basename(d.path) if d.path else "—"
        item = QTreeWidgetItem([name, str(d.pages) if d.pages else ""])
        item.setData(0, _NODE_ROLE, "doc")
        item.setData(0, _DOC_ROLE, d)
        item.setCheckState(0, Qt.Checked if d.needed else Qt.Unchecked)
        return item

    def _app_item(self, a: Appendix) -> QTreeWidgetItem:
        label = a.number or "?"
        if a.name:
            label += f". {a.name}"
        item = QTreeWidgetItem([label, str(a.pages) if a.pages else ""])
        item.setData(0, _NODE_ROLE, "appendix")
        item.setData(0, _APP_ROLE, a)
        item.setCheckState(0, Qt.Checked if a.needed else Qt.Unchecked)
        return item

    # ---- навигация по данным ----
    def _current_ref(self):
        """Возвращает ссылку на текущий выбранный объект (для восстановления)."""
        cur = self.tree.currentItem()
        if cur is None:
            return None
        node = cur.data(0, _NODE_ROLE)
        if node == "doc":
            return ("doc", id(cur.data(0, _DOC_ROLE)))
        if node == "appendix":
            return ("appendix", id(cur.data(0, _APP_ROLE)))
        return ("section", cur.data(0, _SEC_ROLE))

    def _restore_selection(self, ref) -> None:
        if ref is None:
            return
        kind, key = ref
        for i in range(self.tree.topLevelItemCount()):
            sec = self.tree.topLevelItem(i)
            if kind == "section":
                if sec.data(0, _SEC_ROLE) == key:
                    self.tree.setCurrentItem(sec)
                    return
            else:
                for j in range(sec.childCount()):
                    ch = sec.child(j)
                    if kind == "doc" and id(ch.data(0, _DOC_ROLE)) == key:
                        self.tree.setCurrentItem(ch)
                        return
                    if kind == "appendix" and id(ch.data(0, _APP_ROLE)) == key:
                        self.tree.setCurrentItem(ch)
                        return

    def _section_kind(self, item) -> str | None:
        """Ключ раздела, которому принадлежит элемент (сам или родитель)."""
        if item is None:
            return None
        sec = item.data(0, _SEC_ROLE)
        if sec:
            return sec
        parent = item.parent()
        if parent is not None:
            return parent.data(0, _SEC_ROLE)
        return None

    def _node_type(self, item) -> str | None:
        if item is None:
            return None
        return item.data(0, _NODE_ROLE)

    # ---- выбор листа для предпросмотра ----
    def _on_select(self, cur, prev) -> None:
        if cur is None:
            return
        node = cur.data(0, _NODE_ROLE)
        if node == "doc":
            d = cur.data(0, _DOC_ROLE)
            sec = self._section_kind(cur)
            self.selection_changed.emit({
                "kind": "doc", "section": sec,
                "section_label": SECTION_LABELS.get(sec, sec),
                "label": os.path.basename(d.path) if d.path else "—",
                "path": d.path,
                "obj": d,
            })
        elif node == "appendix":
            a = cur.data(0, _APP_ROLE)
            self.selection_changed.emit({
                "kind": "appendix", "section": SECTION_APPENDIX,
                "section_label": SECTION_LABELS.get(SECTION_APPENDIX, SECTION_APPENDIX),
                "label": (a.number + ". " if a.number else "") + a.name,
                "path": a.file_path, "number": a.number, "name": a.name,
                "obj": a,
            })
        else:
            # раздел
            kind = cur.data(0, _SEC_ROLE)
            label = SECTION_LABELS.get(kind, kind)
            # Для ПЗ/Графики/Приложений — placeholder-предпросмотр раздела.
            if kind in (SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX):
                self.selection_changed.emit({
                    "kind": "section", "section": kind,
                    "section_label": label,
                    "label": label,
                })
            else:
                # Листы ГОСТ: cover/title/content/changes (или штампы).
                self.selection_changed.emit({
                    "kind": self._gost_kind(kind), "section": kind,
                    "section_label": label,
                    "label": label,
                })

    def _gost_kind(self, section_kind: str) -> str:
        """Сопоставляет ключ раздела с листом предпросмотра ГОСТ."""
        if section_kind == SECTION_COVER:
            return "cover"
        if section_kind == SECTION_TITLE:
            return "title"
        if section_kind == SECTION_CONTENT:
            return "content"
        if section_kind == SECTION_CHANGES:
            return "changes"
        return "cover"

    # ---- необходимость (чекбокс) ----
    def _on_item_changed(self, item, column) -> None:
        if self._updating or self.project is None or column != 0:
            return
        checked = item.checkState(0) == Qt.Checked
        node = item.data(0, _NODE_ROLE)
        if node == "doc":
            item.data(0, _DOC_ROLE).needed = checked
        elif node == "appendix":
            item.data(0, _APP_ROLE).needed = checked
        elif node == "section":
            kind = item.data(0, _SEC_ROLE)
            for entry in self.project.structure:
                if entry.get("kind") == kind:
                    entry["needed"] = checked
                    break

    # ---- действия ----
    def _delete(self) -> None:
        if self.project is None:
            return
        cur = self.tree.currentItem()
        if cur is None:
            return
        if cur.parent() is None:
            QMessageBox.information(self, "Выберите элемент",
                                    "Удаление раздела невозможно. Выделите документ или приложение.")
            return
        if QMessageBox.question(self, "Удалить", f"Удалить «{cur.text(0)}»?") == QMessageBox.Yes:
            node = cur.data(0, _NODE_ROLE)
            if node == "doc":
                d = cur.data(0, _DOC_ROLE)
                if d in self.project.documents:
                    self.project.documents.remove(d)
            elif node == "appendix":
                a = cur.data(0, _APP_ROLE)
                if a in self.project.appendices:
                    self.project.appendices.remove(a)
            self._refresh()

    def _move(self, delta: int) -> None:
        """Перемещает выбранный элемент вверх/вниз.
        Разделы — в пределах структуры тома; документы/приложения — в пределах своего раздела."""
        if self.project is None:
            return
        cur = self.tree.currentItem()
        if cur is None:
            return
        node = cur.data(0, _NODE_ROLE)
        if node == "section":
            self._move_section(cur, delta)
        elif node == "doc":
            self._move_doc(cur, delta)
        elif node == "appendix":
            self._move_appendix(cur, delta)

    def _move_section(self, item, delta: int) -> None:
        p = self.project
        kind = item.data(0, _SEC_ROLE)
        idx = next((i for i, e in enumerate(p.structure) if e.get("kind") == kind), None)
        if idx is None:
            return
        new_idx = idx + delta
        if new_idx < 0 or new_idx >= len(p.structure):
            return
        p.structure[idx], p.structure[new_idx] = p.structure[new_idx], p.structure[idx]
        self._refresh(select_key=kind)

    def _move_doc(self, item, delta: int) -> None:
        p = self.project
        d = item.data(0, _DOC_ROLE)
        kind = d.doc_kind
        group = [x for x in p.documents if x.doc_kind == kind]
        idx = group.index(d)
        new_idx = idx + delta
        if new_idx < 0 or new_idx >= len(group):
            return
        # Переставляем в общем списке documents.
        pos = p.documents.index(d)
        other = group[new_idx]
        pos_other = p.documents.index(other)
        p.documents[pos], p.documents[pos_other] = p.documents[pos_other], p.documents[pos]
        self._refresh(select_key=kind)

    def _move_appendix(self, item, delta: int) -> None:
        p = self.project
        a = item.data(0, _APP_ROLE)
        idx = p.appendices.index(a)
        new_idx = idx + delta
        if new_idx < 0 or new_idx >= len(p.appendices):
            return
        p.appendices[idx], p.appendices[new_idx] = p.appendices[new_idx], p.appendices[idx]
        self._refresh(select_key=SECTION_APPENDIX)

