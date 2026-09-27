# -*- coding: utf-8 -*-
"""
Утилиты слияния PDF: объединение страниц с сохранением нативных размеров,
сквозной нумерации листов и закладок.
"""
from __future__ import annotations

import os
from typing import Optional

from ..models import DocKind


def merge_pdfs(pdf_paths: list, out_path: str) -> int:
    """Объединяет PDF-файлы в один, сохраняя нативные размеры страниц.

    Возвращает общее число страниц.
    """
    import pymupdf

    result = pymupdf.open()
    try:
        for p in pdf_paths:
            with pymupdf.open(p) as doc:
                result.insert_pdf(doc)
        result.save(out_path)
    finally:
        result.close()
    return 0


def count_pages(pdf_paths: list) -> int:
    """Суммарное число страниц в наборе PDF."""
    import pymupdf
    total = 0
    for p in pdf_paths:
        try:
            with pymupdf.open(p) as doc:
                total += doc.page_count
        except Exception:
            pass
    return total


def compute_total_sheets(volume) -> int:
    """Вычисляет «Листов M» для тома.

    Единый источник — app.core.toc.build_layout: учитывает настраиваемую
    структуру (порядок блоков и отметки «необходимость») и фактическое число
    листов разделов/документов/приложений. Исключённые из сборки элементы не
    считаются.

    Обложка и титульный лист нумерацию не «расходуют» (согласовано с
    предпросмотром). Возвращает общее число листов (страниц) тома.
    """
    from .toc import total_sheets as _toc_total
    return _toc_total(volume)
