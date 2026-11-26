# -*- coding: utf-8 -*-
"""
Мост Python <-> JavaScript для HTML-предпросмотра.

QWebChannel-объект, который JS вызывает через window.bridge.
  - onFieldChanged(key, value)      — поле отредактировано (contenteditable input);
  - onParagraphChanged(index, file, text) — абзац ПЗ отредактирован inline;
  - onTableCellChanged(table, row, cell, file, text) — ячейка таблицы ПЗ;
  - onSignatureClicked(key)         — клик по полю подписи (открыть выбор подписанта);
  - onSheetReady()                  — страница загружена и готова к взаимодействию.
"""
from __future__ import annotations

from PySide6.QtCore import QObject, Slot, Signal


class PreviewBridge(QObject):
    """Мост для JS-вызовов из HTML-предпросмотра."""

    # Сигналы для безопасного (queued) уведомления главного окна
    field_changed = Signal(str, str)        # (key, value)
    paragraph_changed = Signal(int, str, str)   # (индекс абзаца, путь DOCX, текст)
    table_cell_changed = Signal(int, int, int, str, str)   # (таблица, строка, ячейка, DOCX, текст)
    signature_clicked = Signal(str)         # (key)
    sheet_ready = Signal()
    zoom_changed = Signal(float)            # (scale)
    scroll_changed = Signal(float)          # (положение прокрутки, px)
    active_sheet_changed = Signal(int)      # (индекс активного листа)

    @Slot(str, str)
    def onFieldChanged(self, key: str, value: str) -> None:
        self.field_changed.emit(key, value)

    @Slot(int, str, str)
    def onParagraphChanged(self, index: int, file_path: str, text: str) -> None:
        """Абзац пояснительной записки отредактирован прямо в предпросмотре."""
        self.paragraph_changed.emit(int(index), file_path, text)

    @Slot(int, int, int, str, str)
    def onTableCellChanged(self, table: int, row: int, cell: int,
                            file_path: str, text: str) -> None:
        """Ячейка таблицы пояснительной записки отредактирована в предпросмотре."""
        self.table_cell_changed.emit(int(table), int(row), int(cell), file_path, text)

    @Slot(str)
    def onSignatureClicked(self, key: str) -> None:
        self.signature_clicked.emit(key)

    @Slot()
    def onSheetReady(self) -> None:
        self.sheet_ready.emit()

    @Slot(float)
    def onZoomChanged(self, scale: float) -> None:
        self.zoom_changed.emit(scale)

    @Slot(float)
    def onScrollChanged(self, position: float) -> None:
        """Положение прокрутки страницы предпросмотра (px от начала)."""
        self.scroll_changed.emit(float(position))

    @Slot(int)
    def onActiveSheet(self, index: int) -> None:
        """Сменился активный лист (тот, что занимает большую часть окна)."""
        self.active_sheet_changed.emit(int(index))
