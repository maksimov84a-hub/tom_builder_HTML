# -*- coding: utf-8 -*-
"""Компактное окно прогресса сборки тома."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QLabel, QProgressBar, QVBoxLayout


class BuildProgressDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Сборка тома")
        self.setWindowModality(Qt.WindowModal)
        self.setFixedWidth(430)
        self.setMinimumHeight(164)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowCloseButtonHint)
        self.setStyleSheet("""
            QDialog {
                background: #ffffff;
                border: 1px solid #d9e1ec;
                border-radius: 12px;
            }
            QLabel#title {
                color: #172b4d;
                font-size: 17px;
                font-weight: 600;
            }
            QLabel#stage {
                color: #344563;
                font-size: 13px;
            }
            QLabel#detail {
                color: #7a869a;
                font-size: 11px;
            }
            QProgressBar {
                min-height: 10px;
                max-height: 10px;
                border: none;
                border-radius: 5px;
                background: #e9eef5;
                text-align: center;
            }
            QProgressBar::chunk {
                border-radius: 5px;
                background: #2878d0;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(11)

        title = QLabel("Собираем том")
        title.setObjectName("title")
        layout.addWidget(title)

        self.stage = QLabel("Подготовка…")
        self.stage.setObjectName("stage")
        layout.addWidget(self.stage)

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setRange(0, 0)
        layout.addWidget(self.progress)

        self.detail = QLabel("Пожалуйста, подождите")
        self.detail.setObjectName("detail")
        layout.addWidget(self.detail)

    def set_stage(self, text: str) -> None:
        self.stage.setText(text)
        self.detail.setText("Пожалуйста, подождите")
        self.progress.setRange(0, 0)

    def set_sheet_progress(self, current: int, total: int) -> None:
        self.progress.setRange(0, max(1, total))
        self.progress.setValue(min(current + 1, total))
        self.progress.setFormat("%p%")
        self.progress.setTextVisible(True)
        self.stage.setText("Формирование PDF")
        self.detail.setText(f"Обработан лист {min(current + 1, total)} из {total}")

    def finish(self) -> None:
        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        self.stage.setText("Сборка завершена")
        self.detail.setText("Подготовлены выбранные файлы")
