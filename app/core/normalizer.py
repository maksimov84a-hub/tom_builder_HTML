# -*- coding: utf-8 -*-
"""
Нормализация входа: принимает файлы (DOCX/PDF) и определяет
их тип (текст для ПЗ / графика), формат листов и число страниц.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field

from ..models import DocKind, SourceFile
from .formats import SheetFormat, classify_sheet
from .pdf_reader import read_pdf, compute_pages, PdfDocument
from .docx_reader import read_docx, DocxContent

# Допустимые расширения
SUPPORTED_EXT = {".docx", ".doc", ".pdf"}


@dataclass
class NormalizedInput:
    """Результат нормализации одного исходного файла."""
    path: str
    kind: DocKind
    pages: int = 0
    format: SheetFormat | None = None
    is_graphic: bool = False
    content: DocxContent | None = None   # для DOCX
    pdf: PdfDocument | None = None       # для PDF
    errors: list = field(default_factory=list)


def kind_from_ext(ext: str) -> DocKind | None:
    """По расширению — базовый тип документа."""
    ext = ext.lower()
    if ext == ".pdf":
        return DocKind.GRAPHIC     # PDF по умолчанию считается графикой (уточним по содержимому)
    if ext in (".docx", ".doc"):
        return DocKind.POYASNITELNAYA
    return None


def normalize_file(path: str, force_kind: DocKind | None = None) -> NormalizedInput:
    """Разбирает один файл и определяет его характеристики.

    force_kind — если задан, переопределяет определение типа.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext not in SUPPORTED_EXT:
        raise ValueError(f"Неподдерживаемый формат: {ext}")

    kind = force_kind or kind_from_ext(ext)

    if ext == ".pdf":
        return _normalize_pdf(path, kind)
    else:
        return _normalize_docx(path, kind)


def _normalize_pdf(path: str, kind: DocKind) -> NormalizedInput:
    result = NormalizedInput(path=path, kind=kind)
    try:
        pdf = read_pdf(path)
        result.pdf = pdf
        result.pages = pdf.page_count
        result.is_graphic = pdf.is_graphical()
        # ВАЖНО: тип не переопределяем автоматически — раздел задан пользователем.
        # Автоопределение сохраняем как подсказку (is_graphic), но kind не меняем.
        if pdf.pages:
            result.format = pdf.pages[0].format()
    except Exception as e:
        result.errors.append(str(e))
    return result


def _normalize_docx(path: str, kind: DocKind) -> NormalizedInput:
    result = NormalizedInput(path=path, kind=kind)
    try:
        content = read_docx(path)
        result.content = content
        # страницы DOCX точно вычисляются только после рендера; пока оценим по объёму текста
        result.pages = _estimate_docx_pages(content.plain_text)
        # kind уже задан (force_kind или по расширению) — не переопределяем
    except Exception as e:
        result.errors.append(str(e))
    return result


def _estimate_docx_pages(text: str) -> int:
    """Оценка числа страниц по тексту: ~2500 знаков на страницу А4."""
    char_est = 2500
    import math
    n = len(text.strip())
    return max(1, int(math.ceil(n / char_est))) if n else 0


def normalize_many(paths: list, force_kind: DocKind | None = None) -> list:
    """Нормализует список файлов."""
    out = []
    for p in paths:
        try:
            out.append(normalize_file(p, force_kind))
        except Exception as e:
            out.append(NormalizedInput(path=p, kind=DocKind.POYASNITELNAYA, errors=[str(e)]))
    return out
