# -*- coding: utf-8 -*-
"""Чтение PDF: страницы с геометрией, текстом и признаком «текст/графика»."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Optional

from .formats import classify_sheet

# Превосод точек в мм: 1 pt = 25.4/72 мм
PT_TO_MM = 25.4 / 72.0


@dataclass
class PdfPage:
    """Страница PDF с геометрией и содержимым."""
    index: int                 # номер страницы (0-based)
    width_pt: float
    height_pt: float
    has_text: bool = False
    text: str = ""
    image_count: int = 0
    drawings: int = 0          # количество векторных объектов

    @property
    def width_mm(self) -> float:
        return self.width_pt * PT_TO_MM

    @property
    def height_mm(self) -> float:
        return self.height_pt * PT_TO_MM

    def format(self):
        return classify_sheet(self.width_mm, self.height_mm)


@dataclass
class PdfDocument:
    """Разобранный PDF."""
    path: str
    pages: list = field(default_factory=list)

    @property
    def page_count(self) -> int:
        return len(self.pages)

    def is_graphical(self) -> bool:
        """Страница — графическая, если в ней есть векторные объекты/изображения
        и мало текста (типичный чертёж/схема)."""
        total_text = sum(len(p.text.strip()) for p in self.pages)
        total_draw = sum(p.drawings for p in self.pages)
        total_img = sum(p.image_count for p in self.pages)
        # Если в документе есть чертёжные объекты или встроенные изображения
        # и текст короткий — считаем графикой.
        if (total_draw > 0 or total_img > 0) and total_text < 2000:
            return True
        return False

    def is_textual(self) -> bool:
        return not self.is_graphical()


def read_pdf(path: str) -> PdfDocument:
    """Открывает PDF и возвращает документ с геометрией каждой страницы."""
    import pymupdf  # локальный импорт (вынесен для скорости загрузки)

    doc = PdfDocument(path=path)
    try:
        with pymupdf.open(path) as pdf:
            for i, page in enumerate(pdf):
                rect = page.rect
                p = PdfPage(index=i, width_pt=rect.width, height_pt=rect.height)
                # текст
                txt = page.get_text("text") or ""
                p.text = txt.strip()
                p.has_text = bool(p.text)
                # изображения
                imgs = page.get_images(full=True) or []
                p.image_count = len(imgs)
                # векторные рисунки
                try:
                    drawings = page.get_drawings() or []
                    p.drawings = len(drawings)
                except Exception:
                    p.drawings = 0
                doc.pages.append(p)
    except Exception as e:
        raise ValueError(f"Не удалось прочитать PDF «{path}»: {e}") from e

    return doc


def compute_pages(path: str) -> int:
    """Быстрый подсчёт числа страниц PDF без полного разбора."""
    import pymupdf
    try:
        with pymupdf.open(path) as pdf:
            return pdf.page_count
    except Exception:
        return 0
