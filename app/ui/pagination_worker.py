# -*- coding: utf-8 -*-
"""Фоновый пересчёт страниц DOCX средствами Microsoft Word."""
from __future__ import annotations

import logging

from PySide6.QtCore import QObject, Signal, Slot


class PaginationWorker(QObject):
    progress = Signal(str)
    finished = Signal(object)

    def __init__(self, job_id: int, paths: list[tuple[str, object]]):
        super().__init__()
        self.job_id = job_id
        self.paths = list(paths)

    @Slot()
    def run(self):
        from ..core.word_com import word_application
        from ..core.word_pagination import calculate_docx_pages
        results = {}
        errors = {}
        try:
            # Один экземпляр Word на всю пачку документов: запуск приложения
            # стоит ~2 с, столько же добавлял каждый документ при отдельном
            # экземпляре. Экземпляр создаётся здесь, в потоке воркера, —
            # COM-объект нельзя переносить между потоками без маршалинга.
            with word_application() as word:
                for path, first_form in self.paths:
                    self.progress.emit(f"Word рассчитывает страницы: {path}")
                    try:
                        results[path] = calculate_docx_pages(path, first_form,
                                                             word=word)
                    except Exception as exc:
                        logging.getLogger("tom_builder.pagination").exception(
                            "Word pagination failed for %s", path)
                        errors[path] = f"{type(exc).__name__}: {exc}"
        except Exception as exc:
            # Не поднялся сам Word (не установлен, занят COM). Ошибку относим к
            # каждому документу без результата, а не к пустому ключу: в статусе
            # должен быть виден файл, иначе непонятно, что именно не посчитано.
            logging.getLogger("tom_builder.pagination").exception(
                "Pagination worker failed")
            detail = f"{type(exc).__name__}: {exc}"
            for path, _ in self.paths:
                if path not in results:
                    errors.setdefault(path, detail)
        self.finished.emit({"job_id": self.job_id, "results": results,
                            "errors": errors})
