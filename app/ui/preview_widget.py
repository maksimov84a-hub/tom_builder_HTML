# -*- coding: utf-8 -*-
"""
Виджет HTML-предпросмотра листа (QWebEngineView + QWebChannel).

Показывает один лист тома (обложка/титул/содержание/осн. надпись/изменения)
в HTML. Поля редактируются inline; изменения и клики по подписи идут через
QWebChannel в Python.

Масштаб: по умолчанию лист отображается в реальном размере (100%). Есть
панель управления: 100% / вписать в окно / +/- (зум через CSS-масштаб).

Сигналы:
  field_changed(key, value) — отредактировано поле (общие реквизиты/подписи).
  signature_clicked(key)    — клик по полю подписи (открыть выбор подписанта).
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtCore import QUrl, QTimer, Signal

from ..core.preview_html import render_sheet, render_sheet_fragments
from .preview_bridge import PreviewBridge

ZOOM_STEP = 0.1


class SheetPreview(QWidget):
    """Предпросмотр одного листа в HTML."""

    field_changed = Signal(str, str)
    signature_clicked = Signal(str)
    paragraph_changed = Signal(int, str, str)   # (индекс абзаца, путь DOCX, текст)
    cell_changed = Signal(int, int, int, str, str)  # (табл., строка, ячейка, DOCX, текст)
    save_pz_requested = Signal()                # «Сохранить правки ПЗ» нажата

    def __init__(self, parent=None):
        super().__init__(parent)
        self._bridge = PreviewBridge()
        self._bridge.field_changed.connect(self.field_changed)
        self._bridge.signature_clicked.connect(self.signature_clicked)
        self._bridge.paragraph_changed.connect(self._on_paragraph_changed)
        self._bridge.table_cell_changed.connect(self._on_table_cell_changed)
        self._bridge.zoom_changed.connect(self._on_zoom_changed)
        # Несохранённые правки абзацев ПЗ: {путь DOCX: {индекс: текст}}.
        # Применяются поверх текста файла, чтобы правки не терялись при
        # перерисовке листов и при переключении по дереву.
        self._pz_edits: dict = {}
        # Несохранённые правки ячеек таблиц ПЗ: {путь DOCX: {"т:с:я": текст}}.
        self._pz_cell_edits: dict = {}

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # ---- Панель управления масштабом ----
        bar = QHBoxLayout()
        bar.setContentsMargins(6, 4, 6, 4)
        bar.setSpacing(4)
        self.btn_zoom_out = QPushButton("−")
        self.btn_zoom_in = QPushButton("+")
        self.btn_100 = QPushButton("100%")
        self.btn_fit = QPushButton("Вписать в окно")
        self.lbl_zoom = QLabel("100%")
        self.btn_zoom_out.clicked.connect(lambda: self._zoom(-ZOOM_STEP))
        self.btn_zoom_in.clicked.connect(lambda: self._zoom(+ZOOM_STEP))
        self.btn_100.clicked.connect(lambda: self._set_zoom(1.0))
        self.btn_fit.clicked.connect(self._fit_to_window)
        # Кнопка сохранения правок ПЗ, сделанных прямо в предпросмотре.
        # Скрыта, пока правок нет.
        self.btn_save_pz = QPushButton("Сохранить правки ПЗ")
        self.btn_save_pz.setToolTip(
            "Записать изменения текста пояснительной записки в исходный DOCX")
        self.btn_save_pz.clicked.connect(self.save_pz_requested)
        self.btn_save_pz.setVisible(False)
        # Индикатор числа листов в предпросмотре (документ ПЗ показывается
        # целиком, несколькими листами подряд).
        self.lbl_sheets = QLabel("Листов: 1")
        bar.addWidget(QLabel("Масштаб:"))
        bar.addWidget(self.btn_zoom_out)
        bar.addWidget(self.btn_zoom_in)
        bar.addWidget(self.btn_100)
        bar.addWidget(self.btn_fit)
        bar.addStretch(1)

        bar.addWidget(self.btn_save_pz)
        bar.addWidget(self.lbl_sheets)
        bar.addWidget(self.lbl_zoom)
        lay.addLayout(bar)

        self.view = QWebEngineView()
        lay.addWidget(self.view, stretch=1)

        self.channel = QWebChannel()
        self.channel.registerObject("bridge", self._bridge)
        self.view.page().setWebChannel(self.channel)

        self._current_html = ""
        self._pending_snapshot = {}
        self._pending_kind = "cover"
        self._zoom = 1.0

        # Прокрутка к началу после загрузки
        self.view.loadFinished.connect(self._on_load)

    # --- Загрузка листа ---
    def show_sheet(self, kind: str, snapshot: dict, project=None) -> None:
        """Отрисовывает ВСЕ листы выбранного пункта дерева по снимку данных.

        Для документа ПЗ, раздела ПЗ/графики/приложений или их документа
        показываются все листы подряд — с той же пагинацией и сквозной
        нумерацией, что и при сборке. Для стандартных листов ГОСТ — один лист.
        """
        self._pending_kind = kind
        self._pending_snapshot = dict(snapshot)
        self._pending_snapshot["pz_edits"] = self.pz_pending_edits()
        self._pending_snapshot["pz_cell_edits"] = self.pz_pending_cell_edits()
        html, count = self._build_html(kind, self._pending_snapshot, project)
        self._current_html = html
        self.lbl_sheets.setText(f"Листов: {count}")
        self.view.setHtml(html, QUrl("about:blank"))

    def _build_html(self, kind: str, snapshot: dict, project=None):
        """Собирает HTML предпросмотра. Возвращает (html, число листов)."""
        # Правка текста ПЗ доступна и для документа ПЗ, и для раздела ПЗ:
        # в разделе каждый DOCX получает свои data-para-file, поэтому правки
        # адресуются точно (файл + индекс абзаца).
        if snapshot.get("section") == "Пояснительная записка":
            if kind == "doc":
                path = snapshot.get("path", "")
                if path.lower().endswith((".docx", ".doc")):
                    snapshot["pz_editable"] = True
            elif kind == "section":
                snapshot["pz_editable"] = True
        try:
            fragments = render_sheet_fragments(kind, snapshot, project=project)
        except Exception:
            fragments = []
        if not fragments:
            return render_sheet(kind, snapshot), 1
        from ..core.preview_html import _preview_document
        return _preview_document(fragments), len(fragments)

    def refresh(self, kind: str, snapshot: dict, project=None) -> None:
        """Перерисовать текущий лист после изменения данных."""
        self.show_sheet(kind, snapshot, project)

    def _on_load(self, ok: bool) -> None:
        QTimer.singleShot(100, self._fit_to_window)

    def _apply_zoom(self) -> None:
        """Применяет текущий масштаб к телу страницы."""
        self.view.page().runJavaScript(
            "document.body.style.zoom = '%f';" % self._zoom)

    def _set_zoom(self, z: float) -> None:
        self._zoom = max(0.3, min(3.0, round(z, 2)))
        self.lbl_zoom.setText(f"{int(self._zoom * 100)}%")
        self._apply_zoom()

    def _zoom(self, delta: float) -> None:
        self._set_zoom(self._zoom + delta)

    def _fit_to_window(self) -> None:
        """Вписывает лист в окно (по умолчанию — под размер окна).

        Если показано несколько листов (документ ПЗ или раздел), масштаб
        подбирается по самому широкому и самому высокому листу, чтобы ни один
        лист не выходил за границы окна по горизонтали.
        """
        self.view.page().runJavaScript(
            """
            (function(){
              var sheets = document.querySelectorAll('.sheet');
              if (!sheets.length) return;
              var w = 0, h = 0;
              sheets.forEach(function(s){
                w = Math.max(w, s.offsetWidth);
                h = Math.max(h, s.offsetHeight);
              });
              if (!w || !h) return;
              var vw = window.innerWidth, vh = window.innerHeight;
              var margin = 32;
              // Внешние поля .sheet (12/16 мм) — тоже часть макета: без их учёта
              // масштаб получался больше нужного, и правая часть листа (вместе с
              // рамкой) уходила за границу окна — появлялся горизонтальный скролл.
              var cs = getComputedStyle(sheets[0]);
              var mx = (parseFloat(cs.marginLeft) || 0) + (parseFloat(cs.marginRight) || 0);
              var my = (parseFloat(cs.marginTop) || 0) + (parseFloat(cs.marginBottom) || 0);
              var scale = Math.min((vw - margin) / (w + mx), (vh - margin) / (h + my));
              if (scale > 0.1) {
                document.body.style.zoom = scale;
                // обновить индикатор масштаба
                window.bridge && window.bridge.onZoomChanged && window.bridge.onZoomChanged(scale);
              } else {
                document.body.style.zoom = 1;
              }
            })();
            """)

    # --- Служебное ---
    def set_html(self, html: str) -> None:
        self._current_html = html
        self.view.setHtml(html, QUrl("about:blank"))

    def current_html(self) -> str:
        """Возвращает HTML текущего отображаемого листа (для экспорта)."""
        return self._current_html

    def _on_zoom_changed(self, scale: float) -> None:
        """Обновляет индикатор масштаба после автоподгонки под окно."""
        self._zoom = scale
        self.lbl_zoom.setText(f"{int(scale * 100)}%")

    # --- Inline-правка текста ПЗ ---
    def _on_paragraph_changed(self, index: int, path: str, text: str) -> None:
        """Абзац ПЗ отредактирован в предпросмотре — копим правку (без записи)."""
        self._pz_edits.setdefault(path, {})[int(index)] = text
        self._update_save_button()
        self.paragraph_changed.emit(int(index), path, text)

    def _on_table_cell_changed(self, table: int, row: int, cell: int,
                               path: str, text: str) -> None:
        """Ячейка таблицы ПЗ отредактирована — копим правку (без записи в файл)."""
        address = f"{int(table)}:{int(row)}:{int(cell)}"
        self._pz_cell_edits.setdefault(path, {})[address] = text
        self._update_save_button()
        self.cell_changed.emit(int(table), int(row), int(cell), path, text)

    def _update_save_button(self) -> None:
        """Показывает кнопку сохранения, когда есть несохранённые правки ПЗ."""
        count = self.pending_edit_count()
        self.btn_save_pz.setVisible(count > 0)
        if count:
            self.btn_save_pz.setText(f"Сохранить правки ПЗ ({count})")

    def pz_pending_edits(self) -> dict:
        """Несохранённые правки абзацев: {путь DOCX: {индекс: текст}}."""
        return {p: dict(v) for p, v in self._pz_edits.items() if v}

    def pz_pending_cell_edits(self) -> dict:
        """Несохранённые правки ячеек таблиц: {путь DOCX: {"т:с:я": текст}}."""
        return {p: dict(v) for p, v in self._pz_cell_edits.items() if v}

    def pending_edit_count(self) -> int:
        """Общее число несохранённых правок ПЗ (абзацы + ячейки таблиц)."""
        return (sum(len(v) for v in self._pz_edits.values())
                + sum(len(v) for v in self._pz_cell_edits.values()))

    def has_pz_edits(self) -> bool:
        """Есть ли несохранённые правки текста ПЗ (абзацы или таблицы)."""
        return self.pending_edit_count() > 0

    def clear_pz_edits(self, path: str | None = None) -> None:
        """Сбрасывает учёт правок (после сохранения или при смене проекта)."""
        if path is None:
            self._pz_edits.clear()
            self._pz_cell_edits.clear()
        else:
            self._pz_edits.pop(path, None)
            self._pz_cell_edits.pop(path, None)
        self._update_save_button()
