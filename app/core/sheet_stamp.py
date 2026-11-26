# -*- coding: utf-8 -*-
"""Извлечение названия и номера листа из штампа исходного PDF.

Штамп на листе всегда стоит в правом нижнем углу с отступом 5 мм от правого
края и 5 мм от низа. Все размеры отсчитываются ОТ КРАЯ ЛИСТА, поэтому
геометрия не зависит от формата (A4, A1, A3) и ориентации:

    название листа — влево 55..125 мм, вверх 5..20 мм;
    номер листа    — влево 25..40 мм, вверх 20..30 мм.

Области вынесены в настройки проекта (models.GraphicSheetList.name_box /
number_box), значения по умолчанию — в models.SHEET_NAME_BOX / SHEET_NO_BOX.

Главное правило: если название или номер в своих квадратах НЕ найдены, лист
считается неоформленным и в перечень НЕ ПОПАДАЕТ (см. read_sheet_from_pdf,
возвращает None). Так в перечень не попадают титулы, схемы без штампа и
служебные страницы.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass

#: Пунктов в миллиметре (для перевода координат PyMuPDF в мм и обратно).
PT_PER_MM = 72.0 / 25.4

#: Допустимые символы номера листа: цифры, буквенные индексы, «/», «-», «.».
#: В штампе номер обычно просто число («16»), но встречаются «12а», «5.1».
_SHEET_NO_RE = re.compile(r"^[0-9]{1,4}\s*[а-яa-z]?$", re.IGNORECASE)


@dataclass
class SheetStamp:
    """Данные штампа одного листа, прочитанные из PDF."""
    page: int                     # номер страницы (0-based)
    name: str = ""                # название листа (из области названия)
    number: str = ""              # номер листа (из области номера)
    designation: str = ""         # обозначение: шаблон(шифр, номер)


def _box_rect(page, box: tuple):
    """Прямоугольник области в координатах PyMuPDF.

    :param box: (влево от, влево до, вверх от, вверх до) в мм от правого
        нижнего края листа. Например (55, 125, 5, 20) — название листа.

    «Влево N» — от правого края листа, «вверх N» — от нижнего края. В
    координатах PyMuPDF начало — левый верхний угол, поэтому ось Y
    переворачивается: y = высота листа − «вверх».
    """
    import pymupdf
    w_mm = page.rect.width / PT_PER_MM
    h_mm = page.rect.height / PT_PER_MM
    left_from, left_to, up_from, up_to = box
    # «Влево» отсчитывается от правого края, поэтому большая величина «влево»
    # даёт меньший x. Оси нормализуются: прямоугольник всегда неотрицательный.
    x0_mm = w_mm - left_to
    x1_mm = w_mm - left_from
    y0_mm = h_mm - up_to
    y1_mm = h_mm - up_from
    if x1_mm < x0_mm:
        x0_mm, x1_mm = x1_mm, x0_mm
    if y1_mm < y0_mm:
        y0_mm, y1_mm = y1_mm, y0_mm
    return pymupdf.Rect(x0_mm * PT_PER_MM, y0_mm * PT_PER_MM,
                        x1_mm * PT_PER_MM, y1_mm * PT_PER_MM)


def _clean(text: str) -> str:
    """Нормализует текст из PDF: убирает переносы и лишние пробелы."""
    text = (text or "").replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _cell_text(page, box: tuple) -> str:
    """Текст области, склеенный по строкам сверху вниз.

    get_text("text") отдаёт текст в порядке потока, для узкой ячейки штампа
    этого достаточно; пустые строки и пробельные символы отбрасываются.
    """
    rect = _box_rect(page, box)
    raw = page.get_text("text", clip=rect) or ""
    return _clean(raw)


def _normalize_number(text: str) -> str:
    """Оставляет от текста области номера только сам номер листа.

    В область номера может попасть подпись «Лист» или соседние символы;
    берём первый токен, похожий на номер («16», «12а», «5.1»).
    """
    text = _clean(text)
    if not text:
        return ""
    for token in text.split():
        candidate = token.strip(" .,;:|")
        if _SHEET_NO_RE.match(candidate):
            return candidate
    return ""


def read_sheet_from_pdf(path: str, page: int = 0,
                        name_box: tuple | None = None,
                        number_box: tuple | None = None,
                        template: str = "{volume} л.{no}",
                        volume: str = "") -> SheetStamp | None:
    """Читает название и номер листа из штампа страницы PDF.

    Возвращает None, если лист не оформлен: название или номер в своих
    квадратах не найдены. Такие листы в перечень листов не попадают.

    :param path: путь к PDF.
    :param page: номер страницы (0-based).
    :param name_box: область названия (по умолчанию models.SHEET_NAME_BOX).
    :param number_box: область номера (по умолчанию models.SHEET_NO_BOX).
    :param template: шаблон обозначения «{volume} л.{no}».
    :param volume: шифр тома для подстановки в шаблон.
    """
    from ..models import SHEET_NAME_BOX, SHEET_NO_BOX
    if not path or not os.path.exists(path):
        return None
    name_box = name_box or SHEET_NAME_BOX
    number_box = number_box or SHEET_NO_BOX
    try:
        import pymupdf
        with pymupdf.open(path) as doc:
            if page < 0 or page >= doc.page_count:
                return None
            pg = doc[page]
            name = _cell_text(pg, name_box)
            number = _normalize_number(_cell_text(pg, number_box))
    except Exception:
        return None
    # Лист не оформлен — в перечень не берём: без названия или номера строки
    # перечня получались бы пустыми и не отражали реальный состав раздела.
    if not name or not number:
        return None
    stamp = SheetStamp(page=page, name=name, number=number)
    try:
        stamp.designation = (template or "{volume} л.{no}").format(
            volume=volume or "", no=number).strip()
    except (KeyError, IndexError, ValueError):
        stamp.designation = f"{volume} л.{number}".strip()
    return stamp


def read_sheets_from_pdf(path: str, name_box: tuple | None = None,
                         number_box: tuple | None = None,
                         template: str = "{volume} л.{no}",
                         volume: str = "") -> list:
    """Читает штампы всех листов PDF: только оформленные листы.

    Возвращает список SheetStamp в порядке страниц. Страницы без названия или
    без номера в заданных квадратах пропускаются.
    """
    if not path or not os.path.exists(path):
        return []
    out = []
    try:
        import pymupdf
        with pymupdf.open(path) as doc:
            count = doc.page_count
    except Exception:
        return []
    for i in range(count):
        stamp = read_sheet_from_pdf(path, i, name_box=name_box,
                                    number_box=number_box,
                                    template=template, volume=volume)
        if stamp is not None:
            out.append(stamp)
    return out
