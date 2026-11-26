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

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTreeWidget, QTreeWidgetItem,
    QLabel, QMessageBox,
)
from PySide6.QtCore import Qt, Signal

from ..models import (
    VolumeProject,
    SECTION_COVER, SECTION_TITLE, SECTION_CONTENT, SECTION_PZ, SECTION_GRAPHIC,
    SECTION_APPENDIX, SECTION_CHANGES, SECTION_COMPOSITION, SECTION_LABELS,
)

# Роли хранения данных в элементах дерева
_SEC_ROLE = Qt.UserRole          # ключ раздела (str)
_NODE_ROLE = Qt.UserRole + 1     # тип узла (в дереве только "section")


def _section_sheets(project: VolumeProject, kind: str) -> int:
    """Число листов раздела (для колонки «Листы»); 0 — раздел пуст.

    Считается по единому источнику нумерации (toc.build_layout), чтобы число
    в дереве совпадало с содержанием и собранным томом.
    """
    try:
        from ..core.toc import build_layout
        total = 0
        for it in build_layout(project).items:
            if it.section_key == kind:
                total += max(0, int(it.count or 0))
        return total
    except Exception:
        return 0


class SheetTree(QWidget):
    """Дерево структуры тома. Выбор листа -> сигнал selection_changed(dict)."""

    selection_changed = Signal(dict)
    #: Изменена отметка «включать в сборку» (раздел, документ или приложение).
    item_modified = Signal()

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
        """Перестраивает дерево. select_key — ключ раздела для восстановления выбора.

        В дереве показываются ТОЛЬКО разделы тома: документы и приложения здесь
        не выводятся, чтобы список отражал состав тома, а не файлы. Управление
        документами раздела (добавление, флажок «В сборку», удаление) — в
        контекстной панели справа («Документы раздела»).
        """
        self._updating = True
        # Выбранный раздел запоминается ДО очистки дерева: после clear()
        # currentItem уже пуст, и выбор терялся — открытый раздел «исчезал»
        # из дерева после каждой правки настроек.
        selected = (("section", select_key) if select_key
                    else self._current_ref())
        self.tree.clear()
        try:
            if self.project is None:
                return
            p = self.project
            for entry in p.structure:
                kind = entry.get("kind")
                needed = bool(entry.get("needed", True))
                label = SECTION_LABELS.get(kind, kind)
                sec = QTreeWidgetItem([label, str(_section_sheets(p, kind))])
                sec.setData(0, _NODE_ROLE, "section")
                sec.setData(0, _SEC_ROLE, kind)
                sec.setCheckState(0, Qt.Checked if needed else Qt.Unchecked)
                if not needed:
                    sec.setForeground(0, Qt.gray)
                self.tree.addTopLevelItem(sec)
            self._restore_selection(selected)
        finally:
            self._updating = False

    def _doc_item(self, d: SourceFile) -> QTreeWidgetItem:
        name = os.path.basename(d.path) if d.path else "—"
        item = QTreeWidgetItem([name, str(d.pages) if d.pages else ""])
        item.setData(0, _NODE_ROLE, "doc")
        item.setData(0, _DOC_ROLE, d)
        item.setCheckState(0, Qt.Checked if d.needed else Qt.Unchecked)
        if not d.needed:
            item.setForeground(0, Qt.gray)
        return item

    def _app_item(self, a: Appendix) -> QTreeWidgetItem:
        label = a.number or "?"
        if a.name:
            label += f". {a.name}"
        item = QTreeWidgetItem([label, str(a.pages) if a.pages else ""])
        item.setData(0, _NODE_ROLE, "appendix")
        item.setData(0, _APP_ROLE, a)
        item.setCheckState(0, Qt.Checked if a.needed else Qt.Unchecked)
        if not a.needed:
            item.setForeground(0, Qt.gray)
        return item

    # ---- навигация по данным ----
    def _current_ref(self):
        """Возвращает ссылку на текущий выбранный раздел (для восстановления)."""
        cur = self.tree.currentItem()
        if cur is None:
            return None
        return ("section", cur.data(0, _SEC_ROLE))

    def _restore_selection(self, ref) -> None:
        """Возвращает подсветку выбранного раздела после перестройки дерева.

        Сигналы на время выбора блокируются: иначе повторный выбор запускал бы
        новую перерисовку предпросмотра и заново заполнял контекстную форму —
        правка настроек выглядела как «перескок» на начальный вид.
        """
        if ref is None:
            return
        kind, key = ref
        for i in range(self.tree.topLevelItemCount()):
            sec = self.tree.topLevelItem(i)
            if sec.data(0, _SEC_ROLE) == key:
                self.tree.blockSignals(True)
                try:
                    self.tree.setCurrentItem(sec)
                finally:
                    self.tree.blockSignals(False)
                return

    def _section_kind(self, item) -> str | None:
        """Ключ раздела элемента (в дереве все элементы — разделы)."""
        if item is None:
            return None
        return item.data(0, _SEC_ROLE)

    def _node_type(self, item) -> str | None:
        if item is None:
            return None
        return item.data(0, _NODE_ROLE)

    # ---- выбор листа для предпросмотра ----
    def _on_select(self, cur, prev) -> None:
        # Во время перестройки дерева выбор не обрабатывается. QTreeWidget при
        # clear() и наполнении сам двигает currentItem и шлёт currentItemChanged
        # (в т.ч. с cur=None) — без этой проверки завершение пересчёта и любая
        # синхронизация проекта перебрасывали предпросмотр на другой раздел
        # («фокус сбрасывался»). Восстановление выбора делает _restore_selection,
        # сигналы на это время блокируются.
        if self._updating:
            return
        if cur is None:
            return
        kind = cur.data(0, _SEC_ROLE)
        label = SECTION_LABELS.get(kind, kind)
        # ПЗ/Графика/Приложения и «Состав проектной документации» — предпросмотр
        # раздела целиком (все листы).
        if kind in (SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX,
                    SECTION_COMPOSITION):
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
        kind = item.data(0, _SEC_ROLE)
        for entry in self.project.structure:
            if entry.get("kind") == kind:
                entry["needed"] = checked
                break
        # Исключённый раздел остаётся в дереве, но помечается серым.
        item.setForeground(0, Qt.black if checked else Qt.gray)
        self.item_modified.emit()

    # ---- действия ----
    def _delete(self) -> None:
        """Удаление из дерева недоступно: здесь только разделы тома."""
        if self.project is None:
            return
        cur = self.tree.currentItem()
        if cur is None:
            return
        kind = cur.data(0, _SEC_ROLE)
        QMessageBox.information(
            self, "Удаление недоступно",
            f"Раздел «{SECTION_LABELS.get(kind, kind)}» удалить нельзя — "
            "структура тома задаётся порядком разделов (кнопки «Вверх»/«Вниз»).\n\n"
            "Документы и приложения удаляются в контекстной панели справа.")

    def _move(self, delta: int) -> None:
        """Перемещает выбранный раздел вверх/вниз по структуре тома."""
        if self.project is None:
            return
        cur = self.tree.currentItem()
        if cur is None:
            return
        self._move_section(cur, delta)

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
        # Выбор раздела после перестройки восстанавливается без сигналов,
        # поэтому о смене порядка сообщаем явно — иначе предпросмотр и
        # нумерация листов остались бы от прежнего порядка.
        self.item_modified.emit()

