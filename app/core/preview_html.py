# -*- coding: utf-8 -*-
"""
Генератор HTML-листов для предпросмотра по ГОСТ Р 21.101-2020.

Каждый лист — HTML-страница с рамкой (левое 20 мм, остальные 5 мм), основной
надписью (формы 3 / 5 / 6) и кликабельными полями. Поля имеют data-атрибуты,
по которым JS-мост сообщает в Python, какое поле отредактировано.

Листы предпросмотра (автоматические по ГОСТ):
    cover    — обложка (приложение Е);
    title    — титульный лист (приложение Ж), форма 5;
    content  — содержание / опись, форма 5;
    stamp_f1 — основная надпись, форма 3 (графика), 185x55;
    stamp_f2 — основная надпись, форма 5 (первый лист текста), 185x40;
    stamp_f2a— основная надпись, форма 6 (последующие листы), 185x15;
    changes  — лист регистрации изменений, форма 6/5.

Внутренние ключи превью (1 / 2 / "2a") соответствуют формам ГОСТ 3 / 5 / 6:
    1      -> Форма 3 (графический лист, 185x55)
    2      -> Форма 5 (первый/заглавный лист текста, 185x40)
    "2a"   -> Форма 6 (последующие листы текста, 185x15)

Поля редактируются inline (contenteditable). Подписи отображаются изображениями
и открывают выбор подписанта (data-kind="signature").
"""
from __future__ import annotations

import base64
import html as html_mod
import math
import os
import re

from ..config import FontMap, A4_MM, load_config
from ..core.gost import LEFT_MM, RIGHT_MM, TOP_MM, BOTTOM_MM
from .pz_format import (PT_PER_MM, DEFAULT_LINE_HEIGHT, paragraph_css,
                        marker_css)
# Якоря заголовков: по ним строки оглавления ссылаются на пункт записки.
from .pz_toc import heading_anchor, heading_emphasis

# Поля рамки в мм (по ГОСТ)
PAGE_LEFT_MM = LEFT_MM
PAGE_RIGHT_MM = RIGHT_MM
PAGE_TOP_MM = TOP_MM
PAGE_BOTTOM_MM = BOTTOM_MM

# Размеры листа предпросмотра (A4, мм)
PAGE_W_MM = A4_MM[0]
PAGE_H_MM = A4_MM[1]

# Ширина основной надписи (185 мм) и высоты по формам
STAMP_W_MM = 185.0
FORM_HEIGHT = {1: 55.0, 2: 40.0, "2a": 15.0}

# ---- Квадрат сквозной нумерации листов (правый верхний угол) ----------------
# Квадрат с номером листа привязан к верхнему правому углу листа с отступами
# 5 мм вниз и 5 мм влево. Его размер и положение используются и для отступа
# текстовой области, чтобы текст не заезжал под квадрат нумерации.
SHEET_NUM_BOX_MM = 10.0          # размер квадрата (сторона), мм
SHEET_NUM_OFFSET_MM = 5.0        # отступ от верхнего/правого края листа, мм
SHEET_NUM_BOTTOM_MM = SHEET_NUM_OFFSET_MM + SHEET_NUM_BOX_MM   # низ квадрата = 15 мм

# Верхний отступ текстовой области, учитывающий квадрат сквозной нумерации.
# Квадрат занимает зону от 5 мм до 15 мм от верха листа; текстовая область
# начинается ниже этой зоны, чтобы нумерация не перекрывала текст.
TEXT_TOP_INDENT_MM = SHEET_NUM_BOTTOM_MM + 2.0   # = 17 мм от верха листа

# ---- Модель высот для разбивки содержимого DOCX на листы -------------------
# Разбивка на «страницы» должна быть консервативной: она выполняется ДО
# реального рендера, поэтому любые расхождения между оценкой и фактом
# приводят к тому, что строка/таблица уходит под штамп и обрезается
# overflow:hidden. Ниже — измеренные (через печать QtWebEngine) высоты
# элементов и ширина колонки текста, чтобы оценка не занижала реальный
# объём (тогда контент не налезал бы на штамп).
#
# Ширина текстовой колонки: лист A4 (210) - левое поле (20) - правое (5)
# - отступы .sheet-content (6 + 6) = 173 мм.
_PAGE_TEXT_W_MM = PAGE_W_MM - PAGE_LEFT_MM - PAGE_RIGHT_MM - 12.0

# Число символов (примерно), помещающихся в одну строку текстовой колонки
# для кеглей 11pt (основной текст) и 13pt (заголовок). Измерено печатью
# QtWebEngine для реалистичного кириллического текста.
_CHARS_PER_LINE = {11: 81, 13: 61}

# Высота строки — см. _LINE_H_MM ниже (измерено по собранному тому).
# Измерено по печати: одна строка 11pt ≈ 5.5 мм, 13pt ≈ 6.1 мм.

# Высота одной строки таблицы DOCX (мм). Строка из 4 ячеек с коротким
# текстом занимает ~4.3 мм; с длинным текстом, переносящимся на несколько
# визуальных строк, — кратно числу таких строк. Число визуальных строк ячейки
# оценивается по переносу текста при фактической ширине ячейки.
#: Кегль, при котором измерены метрики таблицы ниже. Кегль таблиц задаётся в
#: настройках (config.TextFonts, группа «table»), поэтому метрики
#: пропорционально пересчитываются под него — иначе при смене кегля разбивка
#: листов занижала бы высоту таблиц и они уходили бы под штамп.
_TABLE_BASE_PT = 9.0
_TABLE_CELL_LINE_MM = 9.0 * 1.3 * 25.4 / 72.0  # CSS: 9pt × line-height 1.3
# Отступы и рамка ячейки (.dt .dt-cell): padding 1mm сверху и снизу +
# границы 0.25mm. Вертикальные отступы совпадают с нормализованной копией,
# по которой Word рассчитывает страницы; расхождение на длинной таблице
# накапливалось и обрезало следующие абзацы в HTML-предпросмотре.
_TABLE_CELL_PAD_MM = 2.0
_TABLE_CELL_BORDER_MM = 0.5
# Знаков, помещающихся в 1 мм строки ячейки при базовом кегле таблиц.
# Подобрано по измерению: ячейка 4-колоночной таблицы (~43 мм) вмещает
# ~19 знаков в строку, 3-колоночной (~58 мм) — ~30 знаков.
_TABLE_CHARS_PER_MM = 0.5

# Кэш фактических ширин колонок таблиц DOCX: ключ (путь файла, индекс таблицы,
# число колонок). Ширины читаются из самого DOCX (w:tblGrid), поэтому
# предпросмотр рисует таблицу с той же геометрией, что и исходный документ,
# а не с равными колонками.
_TABLE_GRID_CACHE: dict = {}

def table_col_widths_mm(path: str, table_index: int, ncols: int) -> list:
    """Ширины колонок таблицы в мм (из w:tblGrid исходного DOCX).

    Если ширины в файле не заданы (или файл недоступен), колонки считаются
    равными по ширине текстовой области листа. Сумма ширин нормируется к
    ширине текстовой области, чтобы пропорции колонок сохранялись на всех
    листах.
    """
    ncols = max(1, int(ncols))
    equal = [_PAGE_TEXT_W_MM / ncols] * ncols
    if not path or table_index < 0:
        return equal
    key = (os.path.abspath(path), int(table_index), ncols)
    cached = _TABLE_GRID_CACHE.get(key)
    if cached is not None:
        return cached
    widths = []
    try:
        from docx import Document
        table = Document(path).tables[table_index]
        for col in table.columns:
            w = getattr(col, "width", None)
            if not w:
                widths = []
                break
            widths.append(float(w.mm))
        if len(widths) != ncols or sum(widths) <= 0:
            widths = []
    except Exception:
        widths = []
    if widths:
        scale = _PAGE_TEXT_W_MM / sum(widths)
        result = [w * scale for w in widths]
    else:
        result = equal
    _TABLE_GRID_CACHE[key] = result
    return result

def table_colspec(path: str, table_index: int, ncols: int) -> str:
    """CSS grid-template-columns для таблицы, например "52mm 40mm 81mm"."""
    return " ".join(_css_mm(w) for w in table_col_widths_mm(path, table_index, ncols))

def _css_mm(value: float) -> str:
    """Число миллиметров для CSS (без лишних нулей)."""
    return f"{float(value):.2f}".rstrip("0").rstrip(".") + "mm"

def _text_visual_lines(text: str, chars_per_line: int) -> int:
    """Число визуальных строк, в которые переносится текст при заданной ширине."""
    if not text:
        return 1
    import textwrap
    return max(1, len(textwrap.wrap(text, chars_per_line) or [text]))

def _paragraph_lines(text: str, size: int, fmt=None,
                     font: str | None = None) -> int:
    """Число строк абзаца по МЕТРИКАМ ШРИФТА — как переносит Word.

    Раньше строки считались по «числу знаков в строке» (81 знак при 11 pt).
    На обычном тексте это совпадало, но длинные слова («санитарно-
    эпидемиологические», «встроенно-пристроенный») в узкой колонке не влезают
    целиком: Word переносит их раньше, и строк выходит больше. Оценка по знакам
    этого не видела, пагинация HTML вмещала больше текста, и листы расходились
    с вёрсткой Word.

    Ширина берётся из фактической колонки: текстовая область минус отступы
    абзаца (включая «висячий» отступ списка).

    :param font: имя шрифта текста (config.TextFonts); None — шрифт оформления.
    """
    if not text:
        return 1
    width = _PAGE_TEXT_W_MM
    first_line = 0.0
    if fmt is not None:
        width -= float(getattr(fmt, "left_mm", 0.0) or 0.0)
        width -= float(getattr(fmt, "right_mm", 0.0) or 0.0)
        # «Красная строка» сужает ПЕРВУЮ строку — Word переносит на неё раньше.
        first_line = float(getattr(fmt, "first_line_mm", 0.0) or 0.0)
    try:
        from .textfit import visual_lines as _fit_lines, font_path as _fpath
        return _fit_lines(text, width, float(size),
                          path=_fpath(font) if font else None,
                          first_line_mm=first_line)
    except Exception:
        # Метрики шрифта недоступны — прежняя оценка по числу знаков.
        return _text_visual_lines(text, _chars_per_line(size, fmt))

#: Высота одной визуальной строки текста (мм) по кеглю.
#: Замер по собранному тому (`том.pdf`): при кегле 12 pt шаг строк — 5.56 мм
#: (12 × 0.463), при 14 pt — 6.35 мм (14 × 0.454). Раньше 12 pt оценивался
#: как 0.5 × 12 = 6.0 мм, и абзацы выходили на 8 % выше, чем рисует браузер:
#: на лист влезало меньше текста, и пагинация HTML расходилась с вёрсткой.
_LINE_H_MM = {11: 5.56, 12: 5.56, 13: 6.35, 14: 6.35}

def _line_height(size: int) -> float:
    """Высота одной визуальной строки текста (мм) для кегля size.

    Для кеглей из _LINE_H_MM используются измеренные значения; для остальных —
    расчёт по кеглю (0.46 × кегль: та же пропорция, что у измеренных).
    """
    if size in _LINE_H_MM:
        return _LINE_H_MM[size]
    return round(float(size) * 0.463, 2)

def pz_line_mm(size, fmt=None) -> float:
    """Высота строки абзаца ПЗ (мм) — ЕДИНЫЙ источник для Word, HTML и оценки.

    Разбивка листов выполняется Word'ом по «нормализованной» копии DOCX, а
    рисует листы браузер. Если Word отводит строке больше места, чем рисует
    HTML, листы, нарезанные Word, заполняются лишь частично: Word кладёт на
    лист 20 строк там, где браузер нарисовал бы 30, и низ листа остаётся
    пустым до самого штампа. Раньше так и было: Word брал свой межстрочный
    интервал (≈7.8 мм при 12 pt), а CSS `line-height` для абзацев вообще не
    выводился — браузер рисовал строку 5.0 мм (`normal`).

    :param size: кегль текста абзаца, pt.
    :param fmt: оформление абзаца Word (app.core.pz_format.ParagraphFormat).
        Заданный в исходнике интервал переносится буквально: точный — как есть,
        множитель — умножением на кегль (так же его считает CSS line-height).
    """
    if fmt is not None:
        exact_pt = getattr(fmt, "line_spacing_pt", None)
        if exact_pt:
            return float(exact_pt) / PT_PER_MM
        multiple = getattr(fmt, "line_spacing", None)
        if multiple:
            return float(size) * float(multiple) / PT_PER_MM
    return _line_height(size)

def _chars_per_line(size: int, fmt=None) -> int:
    """Число знаков в строке с учётом абзацных отступов Word (мм).

    Отступы (в т.ч. отступы списка) сужают текстовую колонку, поэтому оценка
    переноса без их учёта занижала бы число строк, и абзац уезжал бы под штамп.
    """
    if size in _CHARS_PER_LINE:
        chars = _CHARS_PER_LINE[size]
    else:
        # Ширина колонки в мм / средняя ширина знака (≈ 0.5 кегля в pt).
        chars = max(1, int(_PAGE_TEXT_W_MM / max(0.5, float(size) * 0.5 / PT_PER_MM)))
    if fmt is None:
        return chars
    left = float(getattr(fmt, "left_mm", 0.0) or 0.0)
    right = float(getattr(fmt, "right_mm", 0.0) or 0.0)
    width = _PAGE_TEXT_W_MM - left - right
    if width <= 10.0:
        return max(1, chars // 2)
    if width >= _PAGE_TEXT_W_MM:
        return chars
    return max(1, int(chars * width / _PAGE_TEXT_W_MM))

def _estimate_paragraph_h(text: str, size: int, fmt=None,
                          font: str | None = None) -> float:
    """Оценка высоты абзаца (мм) по фактической высоте строки в HTML."""
    line_h = pz_line_mm(size, fmt)
    height = _paragraph_lines(text, size, fmt, font) * line_h
    if fmt is None:
        return height
    # Интервалы заданы в ПУНКТАХ (Word/CSS: padding-top: Npt), а высота листа
    # считается в миллиметрах: мм = пт / (пт в одном мм). PT_PER_MM = 72/25.4
    # — это «пунктов в миллиметре», поэтому делим, а не умножаем (умножение
    # давало ×8 и обрывало листы на середине — много пустого места внизу).
    height += (float(getattr(fmt, "space_before_pt", 0.0) or 0.0)
               + float(getattr(fmt, "space_after_pt", 0.0) or 0.0)) / PT_PER_MM
    return height

def _table_ncols(table) -> int:
    """Число колонок таблицы DOCX (по самой широкой строке)."""
    if not table.rows:
        return 1
    return max((len(row) for row in table.rows), default=1)

def _table_scale(size: float | None) -> float:
    """Множитель метрик таблицы относительно измеренных при базовом кегле."""
    try:
        size = float(size)
    except (TypeError, ValueError):
        return 1.0
    if size <= 0:
        return 1.0
    return size / _TABLE_BASE_PT

def _table_cell_chars_per_line(ncols: int, cell_index: int = 0,
                              cell_w_mm: float = 0.0,
                              size: float | None = None) -> int:
    """Число печатных знаков, помещающихся в одну строку ячейки таблицы.

    :param size: кегль таблиц (config.TextFonts, группа «table»); None или
        базовый кегль — измеренные значения.
    """
    cell_w = cell_w_mm if cell_w_mm > 0 else _PAGE_TEXT_W_MM / max(1, ncols)
    per_mm = _TABLE_CHARS_PER_MM / max(0.1, _table_scale(size))
    return max(1, int((cell_w - 4.0) * per_mm))

def _estimate_table_row_h(row, chars_per_line: int, cell_chars: list | None = None,
                          col_widths: list | None = None,
                          size: float | None = None,
                          font: str | None = None) -> float:
    """Оценка высоты одной строки таблицы DOCX (мм).

    :param cell_chars: число знаков в строке по колонкам (если ширины неизвестны).
    :param col_widths: ширины колонок в мм из DOCX (w:tblGrid). Для объединённой
        ячейки ширина складывается по занятым колонкам (col_span) — иначе текст
        широкой ячейки оценивался бы по ширине одной узкой колонки, строки
        выходили бы завышенными и таблица разрывалась по строке на каждый лист.
    :param size: кегль таблиц из настроек — от него зависят высота строки и
        число знаков в строке.
    :param font: шрифт таблиц; None — шрифт оформления.
    """
    row_lines = 1
    ncols = len(col_widths) if col_widths else 0
    table_pt = float(size) if size else _TABLE_BASE_PT
    for i, cell in enumerate(row):
        txt = getattr(cell, "text", None) or ""
        if not txt.strip():
            continue
        per_line = chars_per_line
        span = int(getattr(cell, "col_span", 1) or 1)
        width = 0.0
        if col_widths:
            if i < ncols:
                width = sum(col_widths[i:i + span])
            else:
                width = col_widths[-1] * span
            per_line = _table_cell_chars_per_line(ncols, i, width, size)
        elif cell_chars and i < len(cell_chars) and cell_chars[i] > 0:
            per_line = cell_chars[i]
        # Строки — по метрикам шрифта: перенос в ячейке Word делает по ширинам
        # слов, а оценка по числу знаков на длинных словах занижала результат.
        lines = 0
        if width > 0:
            try:
                from .textfit import visual_lines as _fit_lines, font_path as _fpath
                # Минус внутренние отступы ячейки (по 2 мм слева и справа в CSS).
                lines = _fit_lines(txt, max(1.0, width - 4.0), table_pt,
                                   path=_fpath(font) if font else None)
            except Exception:
                lines = 0
        if lines <= 0:
            lines = _text_visual_lines(txt, per_line)
        row_lines = max(row_lines, lines)
    return (row_lines * _TABLE_CELL_LINE_MM * _table_scale(size)
            + _TABLE_CELL_PAD_MM + _TABLE_CELL_BORDER_MM)

def _page_text_height(stamp_h_mm: float) -> float:
    """Доступная высота текстовой области листа (мм) при данной высоте штампа.

    Учитывает верхний отступ, введённый для квадрата сквозной нумерации
    (TEXT_TOP_INDENT_MM): текстовая область начинается ниже зоны квадрата
    (5..15 мм от верха листа), поэтому её высота меньше на эту величину.
    """
    reserved_bottom = PAGE_BOTTOM_MM + stamp_h_mm + 2.0
    return PAGE_H_MM - reserved_bottom - TEXT_TOP_INDENT_MM - 12.0

def esc(value) -> str:
    """Экранирует текст для HTML."""
    return html_mod.escape(str(value if value is not None else ""))

def _field(key, value, *, cls="", size=10, kind="text", placeholder="",
           title="", usage="", align="", extra_style="") -> str:
    """Редактируемое поле (span с contenteditable).

    align — выравнивание текста ВНУТРИ поля ('left' | 'center' | 'right').
    Поле имеет min-width (1.5em), поэтому без явного text-align короткое
    значение («1», «П») прижималось к левому краю своего бокса и выглядело
    смещённым из центра ячейки.

    extra_style — дополнительные inline-стили поля (например, растянуть его на
    всю ширину ячейки и разрешить перенос строк для графы 9).
    """
    val = "" if value is None else str(value)
    attrs = [f'data-key="{esc(key)}"', f'data-kind="{esc(kind)}"', f'data-size="{size}"']
    if title:
        attrs.append(f'data-title="{esc(title)}"')
    if usage:
        attrs.append(f'data-usage="{esc(usage)}"')
    if placeholder:
        attrs.append(f'data-placeholder="{esc(placeholder)}"')
    if not val:
        cls = (cls + " empty").strip()
    ta = f"text-align:{align};" if align else ""
    return (f'<span class="field {cls}" {" ".join(attrs)} contenteditable="true" '
            f'style="font-size:{size}pt;{ta}{extra_style}">{esc(val)}</span>')

#: Кэш встроенных изображений подписи: путь -> data-URI. Картинка читается
#: один раз за сессию: подписи повторяются на всех листах со штампом.
_SIG_IMG_CACHE: dict = {}

#: Порог «белого» при поиске границ подписи (0..255). Подпись — это росчерк
#: на белом листе, поэтому белые поля вокруг него обрезаются.
_SIG_WHITE_LEVEL = 240
#: Поля вокруг обрезанной подписи, в долях её высоты (чтобы росчерк не
#: касался линий графы).
_SIG_PAD_RATIO = 0.06

def _signature_pixmap(path: str):
    """Растр подписи: обрезан по границам «чернил», белый фон прозрачный.

    Скан или фото подписи почти всегда снят вместе с большими белыми полями:
    сам росчерк занимает небольшую часть картинки. Если вписать такую картинку
    в ячейку 15 × 5 мм целиком, росчерк получается в несколько миллиметров и
    выглядит крошечным. Поэтому белые поля обрезаются, и в ячейку вписывается
    уже сам росчерк.

    Дополнительно белый фон делается ПРОЗРАЧНЫМ (`set_alpha(..., opaque=...)`):
    иначе картинка с белой подложкой перекрывала бы линии сетки штампа, и
    подпись выглядела как «заплатка» поперёк рамки. Прозрачный фон оставляет
    линии штампа целыми при любом размере подписи.

    :return: (Pixmap, ширина_px, высота_px) или None, если файл не читается.
    """
    try:
        import pymupdf
        pix = pymupdf.Pixmap(path)
    except Exception:
        return None
    try:
        gray = pymupdf.Pixmap(pymupdf.csGRAY, pix)
    except Exception:
        return None
    w, h, stride = gray.width, gray.height, gray.stride
    if w <= 0 or h <= 0:
        return None
    samples = gray.samples
    level = _SIG_WHITE_LEVEL

    def row_has_ink(y: int) -> bool:
        base = y * stride
        return any(samples[base + x] < level for x in range(w))

    def col_has_ink(x: int, y0: int, y1: int) -> bool:
        return any(samples[y * stride + x] < level for y in range(y0, y1))

    top, bottom = 0, h
    while top < bottom and not row_has_ink(top):
        top += 1
    while bottom > top and not row_has_ink(bottom - 1):
        bottom -= 1
    if top >= bottom:
        return None                      # картинка целиком белая
    left, right = 0, w
    while left < right and not col_has_ink(left, top, bottom):
        left += 1
    while right > left and not col_has_ink(right - 1, top, bottom):
        right -= 1
    if left >= right:
        return None
    # Запас по краям, чтобы росчерк не лип к границам графы.
    pad_x = max(1, int((right - left) * _SIG_PAD_RATIO))
    pad_y = max(1, int((bottom - top) * _SIG_PAD_RATIO))
    ix0, iy0 = max(0, left - pad_x), max(0, top - pad_y)
    ix1, iy1 = min(w, right + pad_x), min(h, bottom + pad_y)
    try:
        clip = pymupdf.IRect(ix0, iy0, ix1, iy1)
        out = pymupdf.Pixmap(pix, ix1 - ix0, iy1 - iy0, clip)
    except Exception:
        out = pix
    else:
        out = _with_transparent_white(out)
    return out, out.width, out.height

def _with_transparent_white(pix):
    """Копия растра, у которой белый фон стал прозрачным.

    Подпись — тёмный росчерк на белом листе. Если оставить белую подложку,
    картинка перекрывает линии сетки штампа: подпись на печати выглядит как
    белая заплатка, перечёркивающая рамку. Прозрачный фон сохраняет линии.

    Используется `Pixmap.set_alpha(opaque=...)`: пиксели указанного цвета
    получают нулевую непрозрачность, остальные остаются как есть. Вызов
    делается на обычной копии растра: в этой версии PyMuPDF у неё появляется
    альфа-канал (конструктор `Pixmap(src, alpha=True)` не поддерживается).
    """
    try:
        import pymupdf
        out = pymupdf.Pixmap(pix)
        n = out.colorspace.n
        white = (255, 255, 255) if n >= 3 else (255,) * max(1, n)
        out.set_alpha(opaque=white)
        return out
    except Exception:
        # Прозрачность недоступна — отдаём картинку как есть (подпись видна,
        # но с белой подложкой).
        try:
            import pymupdf
            return pymupdf.Pixmap(pix)
        except Exception:
            return pix

def signature_src(path: str) -> str:
    """Источник изображения подписи для HTML: встроенный data-URI.

    Картинка читается с диска, обрезается по границам росчерка
    (`_signature_pixmap`) и встраивается в разметку как data-URI.

    Ссылку `file:///…` страница предпросмотра и панели штампа открыть не может:
    они грузятся через `setHtml(…, about:blank)`, и QtWebEngine запрещает такой
    странице читать локальные файлы — картинка молча не отображалась.

    :return: `data:`-URI, либо "" — если файла нет или он не читается.
    """
    path = (path or "").strip()
    if not path:
        return ""
    cached = _SIG_IMG_CACHE.get(path)
    if cached is not None:
        return cached
    src = ""
    made = _signature_pixmap(path)
    if made is not None:
        pix, _w, _h = made
        try:
            data = pix.tobytes("png")
            if data:
                src = ("data:image/png;base64,"
                       + base64.b64encode(data).decode("ascii"))
        except Exception:
            src = ""
    _SIG_IMG_CACHE[path] = src
    return src

#: Насколько подпись может выйти за границы своей ячейки. Живой росчерк в
#: ГОСТ-графе 15 × 5 мм не помещается целиком, поэтому подпись разрешено
#: увеличивать до 1.5 размеров ячейки — она выходит за линии графы, как при
#: подписывании от руки. Линии при этом остаются целыми: у картинки подписи
#: прозрачный фон (`_with_transparent_white`), поэтому она не перекрывает
#: сетку штампа белой подложкой. Расчёт единый для предпросмотра, Word и PDF.
SIGNATURE_OVERFLOW = 1.5

def signature_fit_mm(path: str, box_w_mm: float, box_h_mm: float,
                     pad_mm: float = 0.4,
                     overflow: float = SIGNATURE_OVERFLOW) -> tuple:
    """Размер подписи (мм) при вписывании ОБРЕЗАННОГО росчерка в ячейку.

    Тот же расчёт, что и в HTML, но по обрезанной картинке: Word и PDF
    вписывают в ячейку именно росчерк, без белых полей. Раньше документ и
    предпросмотр давали разный размер подписи.

    :param overflow: во сколько раз подпись может превысить ячейку. По
        умолчанию 1.0 — строго внутри графы: подпись не должна перекрывать
        линии сетки штампа.

    :return: (ширина, высота) в мм; (0, 0) — картинка не читается.
    """
    made = _signature_pixmap(path)
    if made is None:
        return 0.0, 0.0
    _pix, w_px, h_px = made
    if w_px <= 0 or h_px <= 0:
        return 0.0, 0.0
    avail_w = max(0.1, (float(box_w_mm) - 2 * pad_mm) * overflow)
    avail_h = max(0.1, (float(box_h_mm) - 2 * pad_mm) * overflow)
    scale = min(avail_w / w_px, avail_h / h_px)
    return w_px * scale, h_px * scale

def _sig_field(key, role_label, signer, *, size=10, compact=False,
               hint=True) -> str:
    """Поле подписи: изображение (или плейсхолдер) + Фамилия.

    :param hint: показывать подсказку «(подпись)», когда подписи нет. Она нужна
        только в интерактивном предпросмотре; в собранном томе (HTML/PDF/DOCX)
        графа остаётся пустой (snapshot["hide_signature_hint"]).
    """
    ig = ' sig-img-c' if compact else ''
    sn = ' sig-surname-c' if compact else ''
    src = signature_src(signer.get("signature_image", "") if signer else "")
    if compact:
        # В компактном режиме (узкая ячейка) показываем только фамилию мелким кеглем.
        surname = f'<span class="sig-surname{sn}">{esc(signer.get("surname", "") if signer else "")}</span>'
        img = ''
    elif src:
        img = (f'<img class="sig-img{ig}" src="{src}" '
               f'alt="{esc(signer.get("surname", ""))}">')
        surname = f'<span class="sig-surname{sn}">{esc(signer.get("surname", ""))}</span>'
    else:
        img = ('<span class="sig-img sig-placeholder">(подпись)</span>'
               if hint else '')
        surname = f'<span class="sig-surname{sn}">{esc(signer.get("surname", "") if signer else "")}</span>'
    return (f'<span class="field signature" data-key="{esc(key)}" data-kind="signature" '
            f'data-size="{size}" data-title="{esc(role_label)}" contenteditable="false">'
            f'{img}{surname}</span>')

# ---- CSS -------------------------------------------------------------------

_CSS_TEMPLATE = """
* { box-sizing: border-box; }
html, body { margin:0; padding:0; }
body {
  background:#e8e8e8;
  font-family:'Segoe UI', Arial, sans-serif;
  color:#111;
  -webkit-user-select:none; user-select:none;
}
.sheet {
  position:relative;
  width:{page_w}mm;
  height:{page_h}mm;
  margin:12mm 16mm;
  background:#fff;
  box-shadow:0 1px 6px rgba(0,0,0,.35);
}
/* Рамка листа и квадрат номера рисуются полосами (см. _frame_rules и
   _sheet_number_rules), а не CSS-рамкой: у border штрих уходит ВНУТРЬ блока,
   а у линий сетки штампа — наружу от координаты, из-за чего на стыке рамки со
   штампом получалась «двойная» линия толщиной 0.25 мм. */
.sheet-number {
  position:absolute;
  z-index:3;
  right:5mm;
  top:5mm;
  width:10mm;
  height:10mm;
  background:#fff;
  font-family:'ISOCPEUR','Segoe UI',Arial,sans-serif;
  /* Цифра отцентрирована по квадрату и увеличена, чтобы на листе читалась
     высотой ~2.5-3 мм (совпадает с PDF-выводом draw_sheet_number_box). */
  font-size:12pt;
  line-height:10mm;
  text-align:center;
}
.sheet-content {
  position:absolute;
  left:{left}mm;
  top:{top}mm;
  right:{right}mm;
  bottom:{bottom}mm;
  padding:6mm;
  overflow:hidden;
}
.field { outline:none; border-bottom:1px dashed #99a; min-width:1.5em; display:inline-block; }
/* Абзацы и списки пояснительной записки (в предпросмотре редактируются). */
.pz-para { outline:none; margin:0; }
/* Пункт списка Word: маркер — отдельный бокс в поле «висячего» отступа
   (w:hanging) слева от текста, сам текст пункта начинается на его отступе. */
.pz-item { position:relative; }
.pz-item > .pz-mark { font:inherit; color:inherit; }
.pz-list-marker { margin:0; }
.pz-para:hover { background:#f4f8ff; box-shadow:inset 0 0 0 1px #cfe0f5; }
.pz-para:focus { background:#fff7d6; box-shadow:inset 0 0 0 1px #2a6; }
/* Таблицы DOCX: CSS-grid с фактическими ширинами колонок из исходного файла.
   Объединённые ячейки занимают свой span, поэтому дублей текста не возникает.
   В предпросмотре ячейки редактируются (contenteditable), как и абзацы ПЗ. */
.dt { display:grid; margin:1mm 0; border:0.25mm solid #000; }
.dt .dt-row { display:contents; }
.dt .dt-cell {
  border:0.25mm solid #000;
  padding:1mm 2mm;
  /* Шрифт и кегль таблиц — из настроек (config.TextFonts, группа «table»). */
  font-family:{table_font};
  font-size:{table_size}pt;
  line-height:1.3;
  min-height:4.3mm;
  white-space:pre-wrap;
  overflow-wrap:break-word;
  outline:none;
}
.dt .pz-cell:hover { background:#f4f8ff; box-shadow:inset 0 0 0 1px #cfe0f5; }
.dt .pz-cell:focus { background:#fff7d6; box-shadow:inset 0 0 0 1px #2a6; }
/* Продолжения объединённых ячеек: место занимают, но рамку не рисуют
   (её рисует начало объединения, растянутое на свой span). */
.dt .dt-merged { border:none; padding:0; min-height:0; }
.field:focus { background:#fff7d6; border-bottom:2px solid #2a6; }
.field.empty { color:#aaa; font-style:italic; }
.field.empty::after { content:attr(data-placeholder); }
.field[contenteditable="false"] { cursor:pointer; }
.field.signature { border:none; display:inline-flex; align-items:center; gap:4px; vertical-align:middle; }
.sig-img { height:6mm; max-width:30mm; object-fit:contain; }
.sig-placeholder { color:#bbb; font-style:italic; font-size:9pt; }
.sig-surname { font-size:9pt; }
.sig-img-c { height:3mm; max-width:10mm; object-fit:contain; }
.sig-surname-c { font-size:6pt; }
/* Графа «Подп.» строки подписей: изображение подписи из реестра подписантов,
   подобранное по фамилии в соседней ячейке. Клик открывает выбор подписанта.
   Росчерк в графу 15 × 5 мм не помещается, поэтому подпись занимает 1.5
   размера ячейки (SIGNATURE_OVERFLOW) и выходит за её линии. Линии сетки при
   этом остаются целыми: фон картинки прозрачный (_with_transparent_white),
   белая подложка не перекрывает рамку. Пропорции сохраняет object-fit:contain. */
.sig-img-row { width:150%; height:150%; object-fit:contain; display:block; }
.field.sig-cell { border:none; display:flex; align-items:center;
  justify-content:center; width:100%; height:100%; overflow:visible;
  position:relative; }
.field.sig-cell img { width:150%; height:150%; object-fit:contain;
  display:block; }
/* Графа 9 основной надписи: логотип организации (PDF -> PNG), автомасштаб
   под ячейку без искажения пропорций. */
.stamp-logo { width:100%; height:100%; object-fit:contain; display:block; }
.stamp {
  position:absolute;
  right:{right}mm;
  bottom:{bottom}mm;
  width:{stamp_w}mm;
  font-family:'ISOCPEUR','Segoe UI',Arial,sans-serif;
  border:0.25mm solid #000;
}
.stamp table { width:100%; height:100%; border-collapse:collapse; table-layout:fixed; }
.stamp td { border:0.25mm solid #000; text-align:center; vertical-align:middle; padding:0 1mm; }
.stamp .lbl { font-size:7pt; color:#333; }
.stamp .val { font-size:9pt; }
/* Ячейки CSS-grid основной надписи */
.stamp .gcell { border:0.25mm solid #000; display:flex; align-items:center;
  justify-content:center; text-align:center; padding:0 1mm; }
.stamp .gcell table { width:100%; height:100%; }
.cover { text-align:center; }
.cover .logo { font-size:18pt; font-weight:bold; margin-top:20mm; }
h1,h2 { margin:0; text-align:center; font-weight:normal; }
/* AutoTOC table: Name | Designation | Sheets.
   Ширины колонок фиксированы в мм (models.TOC_WIDTHS = 95/50/40): таблица
   занимает всю ширину между линиями рамки листа и не «плавает» от длины
   текста. Перенос в ячейках разрешён — длинное наименование или шифр без
   пробелов иначе распирал строку и выводил правую часть содержания за рамку.
   Границы — полная сетка (все вертикальные и горизонтальные линии), как у
   обычной таблицы: у каждой ячейки рамка 0.25 мм со всех сторон, внешний
   контур — по краям таблицы. */
.toc-table {
  width:100%; border-collapse:collapse; table-layout:fixed;
  border:0.25mm solid #000;
}
.toc-table th, .toc-table td {
  border:0.25mm solid #000;
  padding:1.2mm 1.5mm; font-weight:normal; vertical-align:top;
  overflow-wrap:break-word; word-break:break-word;
}
.toc-table .toc-name { text-align:left; }
.toc-table .toc-des, .toc-table .toc-sheets { text-align:center; }
/* Шапка таблицы содержания — полужирная и по центру (как в DOCX: bold только у
   заголовков колонок). Строки содержания выводятся ОБЫЧНЫМ шрифтом. */
.toc-table .toc-head th { font-weight:bold; font-size:9pt; text-align:center; }
/* Листы, на которых таблица упирается в рамку листа («Содержание» и
   «Перечень листов графической части»): ширины колонок в сумме дают 185 мм —
   ровно расстояние между линиями рамки. Внутренний отступ .sheet-content
   (6 мм с каждой стороны) компенсируется отрицательными полями, иначе
   таблица была бы уже рамки на 12 мм. */
.bleed-x { margin-left:-6mm; margin-right:-6mm; }
/* Внешний контур таблицы ложится ТОЧНО на ось линии рамки листа.
   При border-collapse граница рисуется по центру линии сетки, поэтому
   коробка таблицы на 100% ширины даёт контур, смещённый внутрь на половину
   толщины штриха (0.125 мм) с каждой стороны — правая линия таблицы шла на
   0.13 мм левее линии рамки, и на печати получалась «двойная» линия.
   Таблица расширена на толщину штриха и сдвинута на половину с каждой
   стороны: оси крайних линий совпадают с осями линий рамки (20 и 205 мм). */
.bleed-x .toc-table, .bleed-x .gsheetlist, .bleed-x .comp-table {
  width:calc(100% + 0.25mm);
  margin-left:-0.125mm;
  margin-right:-0.125mm;
}
/* Таблица «Состав проектной документации» начинается сразу ПОД квадратом
   сквозной нумерации: над таблицей печатается только надпись раздела, поэтому
   снимается верхний внутренний отступ .sheet-content (6 мм). */
.bleed-top { margin-top:-6mm; }
/* Перечень листов графической части: таблица во всю ширину листа,
   колонки Обозначение | Наименование | Примечание. */
.gsheetlist { width:100%; border-collapse:collapse; table-layout:fixed; }
.gsheetlist th, .gsheetlist td {
  border:0.25mm solid #000;
  padding:0.6mm 1.5mm;
  font-size:10pt;
  line-height:1.25;
  vertical-align:middle;
}
.gsheetlist th { font-weight:normal; text-align:center; }
.gsheetlist td { word-break:break-word; }
/* Состав проектной документации: Номер тома 15 | Обозначение 60 |
   Наименование 90 | Примечание 20 мм (185 мм — вся ширина между линиями
   рамки). Высота шапки 15 мм, высота строки 8 мм — те же значения, что в
   Word и PDF (models.COMPOSITION_HEAD_MM / COMPOSITION_ROW_MM). */
.comp-table { width:100%; border-collapse:collapse; table-layout:fixed;
  border:0.25mm solid #000; }
.comp-table th, .comp-table td {
  border:0.25mm solid #000;
  padding:0 1.5mm;
  font-size:10pt;
  line-height:1.15;
  vertical-align:middle;
  overflow-wrap:break-word;
  word-break:break-word;
}
.comp-table th { height:15mm; font-weight:normal; text-align:center; }
.comp-table td { height:8mm; }
.comp-table td:first-child { text-align:center; }
/* Надпись над таблицей состава: по центру, полужирная, без отступа снизу —
   таблица идёт сразу под ней. Высота надписи входит в расчёт ёмкости листа
   (COMPOSITION_TITLE_MM). */
.comp-title {
  height:8mm; line-height:8mm;
  font-size:12pt; font-weight:bold; text-align:center;
  margin-bottom:4mm;
}
"""

def _text_fonts() -> "TextFonts":
    """Шрифты текста ПЗ из конфигурации (обычный текст / заголовки / таблицы).

    Читаются из config.json; если файла нет — значения по умолчанию. Кегли
    используются и в разметке, и в оценке высоты абзацев, поэтому предпросмотр
    и собранный том дают одинаковое число листов.
    """
    try:
        return load_config().text_fonts
    except Exception:
        from ..config import TextFonts
        return TextFonts()

def _css() -> str:
    """Подставляет размеры листа в шаблон CSS (без .format, т.к. CSS содержит {})."""
    css = _CSS_TEMPLATE
    css = css.replace("{page_w}", str(PAGE_W_MM))
    css = css.replace("{page_h}", str(PAGE_H_MM))
    css = css.replace("{left}", str(PAGE_LEFT_MM))
    css = css.replace("{right}", str(PAGE_RIGHT_MM))
    css = css.replace("{top}", str(PAGE_TOP_MM))
    css = css.replace("{bottom}", str(PAGE_BOTTOM_MM))
    css = css.replace("{stamp_w}", str(STAMP_W_MM))
    css = css.replace("{change_w}", str(sum(CHANGE_COLS)))
    # Шрифт и кегль таблиц — из настроек текста ПЗ.
    tf = _text_fonts()
    css = css.replace("{table_font}", _css_font_family(tf.font("table")))
    css = css.replace("{table_size}", str(tf.size("table")))
    return css

def _css_font_family(name: str) -> str:
    """CSS-значение font-family по имени шрифта (с запасными семействами)."""
    family = (name or "").strip() or "ISOCPEUR"
    return "'%s','Segoe UI',Arial,sans-serif" % family.replace("'", "")

# ---- Основная надпись ------------------------------------------------------
# Точная геометрия по ГОСТ Р 21.101-2020, приложение Ж.
#   Форма 3 (графика): 185 x 55 мм
#   Форма 5 (первый лист текста): 185 x 40 мм
#   Форма 6 (последующий лист текста): 185 x 15 мм
# Таблица изменений (графы 14-19): 6 колонок 10,10,10,10,15,10 = 65 мм
#   (Изм. | Кол.уч. | Лист | № док. | Подл. | Дата)
# Обозначение (графа 1) и правая узкая колонка «Лист / Листов» — остаток до 185 мм.

CHANGE_COLS = [10, 10, 10, 10, 15, 10]       # = 65 мм
#: Подписи колонок в штампе. Ключи колонок — единый источник
#: models.STAMP_CHANGE_COLS / geometry.CHANGE_KEYS (те же, что в модели проекта).
CHANGE_LABELS = ["Изм.", "Кол.уч", "Лист", "№док.", "Подп.", "Дата"]
CHANGE_KEYS = ["izm", "uch", "list", "doc", "sign", "date"]
#: X-координаты колонок таблицы изменений (мм от левого края штампа).
CHANGE_X = [0.0, 10.0, 20.0, 30.0, 40.0, 55.0]

# Кегль надписей и значений граф «Стадия / Лист / Листов» в основной надписи
# (формы 3, 5, 6 ГОСТ). По требованию — единый размер 12 pt.
STAGE_CELL_PT = 12
# Кегль текста графы 9 (наименование/индекс организации) — как у «Стадия/Лист».
CELL9_PT = 12
SHEET_COL_W = 10.0                           # правая колонка «Лист / Листов»

def _grid_cell(c, r, cs=1, rs=1, content="", cls="val", extra=""):
    """Ячейка CSS-grid: c,r — 1-based колонка/строка."""
    style = f"grid-column:{c} / span {cs}; grid-row:{r} / span {rs};"
    if extra:
        style += extra
    return f'<div class="gcell {cls}" style="{style}">{content}</div>'

# ---- Форма 3 (ГОСТ, графический лист): 185 x 55 мм — сетка из образца «Форма 5.pdf» ----
# ВАЖНО: образец «Форма 5.pdf» от пользователя — это ФОРМА 3 ГОСТ (графический лист).
# Верхние 2 полосы (обозначение + наименование) идут на всю длину.
FORM3_H = 55.0
FORM3_W = 185.0
FORM3_VLINES = [
    (0.0,  [(0.0, 55.0)]),
    (10.0, [(0.0, 25.0)]),
    (20.0, [(0.0, 55.0)]),
    (30.0, [(0.0, 25.0)]),
    (40.0, [(0.0, 55.0)]),
    (55.0, [(0.0, 55.0)]),
    (65.0, [(0.0, 55.0)]),
    (135.0,[(25.0, 55.0)]),
    (150.0,[(25.0, 40.0)]),
    (165.0,[(25.0, 40.0)]),
    (185.0,[(0.0, 55.0)]),
]
FORM3_HLINES = [
    (0.0,  [(0.0, 185.0)]),
    (5.0,  [(0.0, 65.0)]),
    (10.0, [(0.0, 65.0), (65.0, 185.0)]),
    (15.0, [(0.0, 65.0)]),
    (20.0, [(0.0, 65.0)]),
    (25.0, [(0.0, 185.0)]),
    (30.0, [(0.0, 65.0), (135.0, 185.0)]),
    (35.0, [(0.0, 65.0)]),
    (40.0, [(0.0, 65.0), (65.0, 185.0)]),
    (45.0, [(0.0, 65.0)]),
    (50.0, [(0.0, 65.0)]),
    (55.0, [(0.0, 185.0)]),
]

# ---- Форма 5 (ГОСТ, первый/заглавный лист текста): 185 x 40 мм — из «Форма 3.pdf» ----
# ВАЖНО: образец «Форма 3.pdf» от пользователя — это ФОРМА 5 ГОСТ (первый лист текста).
FORM5_H = 40.0          # высота штампа, мм
FORM5_W = 185.0         # ширина штампа, мм
# Вертикальные линии: (x, [(y0, y1)...]) — y от верха штампа вниз.
FORM5_VLINES = [
    (0.0,  [(0.0, 40.0)]),
    (10.0, [(0.0, 15.0)]),
    (20.0, [(0.0, 40.0)]),
    (30.0, [(0.0, 15.0)]),
    (40.0, [(0.0, 40.0)]),
    (55.0, [(0.0, 40.0)]),
    (65.0, [(0.0, 40.0)]),
    (135.0,[(15.0, 40.0)]),
    (150.0,[(15.0, 25.0)]),
    (165.0,[(15.0, 25.0)]),
    (185.0,[(0.0, 40.0)]),
]
# Горизонтальные линии: (y, [(x0, x1)...]).
FORM5_HLINES = [
    (0.0,  [(0.0, 185.0)]),
    (5.0,  [(0.0, 65.0)]),
    (10.0, [(0.0, 65.0)]),
    (15.0, [(0.0, 185.0)]),
    (20.0, [(0.0, 65.0), (135.0, 185.0)]),
    (25.0, [(0.0, 65.0), (135.0, 185.0)]),
    (30.0, [(0.0, 65.0)]),
    (35.0, [(0.0, 65.0)]),
    (40.0, [(0.0, 185.0)]),
]
LN = "0.25mm solid #000"

#: Толщина всех линий оформления (рамка, квадрат нумерации, сетка штампа), мм.
#: Единая величина: линии на стыке рамки и штампа должны совпадать и по
#: положению, и по толщине.
LINE_MM = 0.25

#: Половина толщины линии. На неё сдвигается штрих, чтобы его ОСЬ (а не край)
#: легла на заданную координату: у рамки листа и у границ ячеек CSS-рамка
#: уходит внутрь блока, а у отдельной линии border-left/top — наружу. Без
#: такого сдвига линии, нарисованные разными приёмами, расходились ровно на
#: толщину линии, и на стыке рамки со штампом получалась «двойная» линия.
_HALF_LINE_MM = LINE_MM / 2.0

def _rule_h(y, x0, x1):
    """Горизонтальная линия с осью ровно на y, от x0 до x1 (мм).

    Рисуется залитой полосой высотой LINE_MM со сдвигом на половину толщины,
    поэтому центр штриха приходится точно на координату — совпадает с линией
    рамки листа и с границами ячеек.
    """
    return (f'<div style="position:absolute;left:{x0:g}mm;'
            f'top:{y - _HALF_LINE_MM:.4f}mm;width:{x1 - x0:g}mm;'
            f'height:{LINE_MM:g}mm;background:#000;"></div>')

def _rule_v(x, y0, y1):
    """Вертикальная линия с осью ровно на x, от y0 до y1 (мм)."""
    return (f'<div style="position:absolute;left:{x - _HALF_LINE_MM:.4f}mm;'
            f'top:{y0:g}mm;width:{LINE_MM:g}mm;height:{y1 - y0:g}mm;'
            f'background:#000;"></div>')

# ---- Форма 6 (185 x 15): сетка из образца «Форма 6.pdf» ----
# Вертикали: 0,10,20,30,40,55,65,175,185.
#   Левый блок (0..65): таблица изменений 10|10|10|10|15|10.
#   Правый блок (65..185): обозначение (65..175) | Лист (175..185).
# Горизонтали: лево 0,5,10,15 (два ряда по 5мм); право 0,15; зона «Лист» 0,7,15.
FORM6_H = 15.0
FORM6_W = 185.0
FORM6_VLINES = [
    (0.0,  [(0.0, 15.0)]),
    (10.0, [(0.0, 15.0)]),
    (20.0, [(0.0, 15.0)]),
    (30.0, [(0.0, 15.0)]),
    (40.0, [(0.0, 15.0)]),
    (55.0, [(0.0, 15.0)]),
    (65.0, [(0.0, 15.0)]),
    (175.0,[(0.0, 15.0)]),
    (185.0,[(0.0, 15.0)]),
]
FORM6_HLINES = [
    (0.0,  [(0.0, 185.0)]),
    (5.0,  [(0.0, 65.0)]),
    (7.0,  [(175.0, 185.0)]),
    (10.0, [(0.0, 65.0)]),
    (15.0, [(0.0, 185.0)]),
]
FORM6_CHANGE_LABELS = ["Изм.", "Кол.уч", "Лист", "№док.", "Подп.", "Дата"]

def _hline(y, x0, x1, cut=0.0):
    """Горизонтальная линия на высоте y (мм от верха штампа, вниз), от x0 до x1.

    cut — срез сверху (мм): координаты сдвигаются вверх, линии выше среза отбрасываются.
    """
    if y < cut:
        return ""
    return _rule_h(y - cut, x0, x1)

def _vline(x, y0, y1, cut=0.0):
    """Вертикальная линия на x, от y0 до y1 (мм от верха штампа, вниз).

    cut — срез сверху (мм): координаты сдвигаются вверх; отрезок обрезается снизу
    до среза (y0 < cut) или отбрасывается целиком.
    """
    if y1 <= cut:
        return ""
    y0c = max(y0, cut)
    return _rule_v(x, y0c - cut, y1 - cut)

def _lbl(x0, y0, w, h, text, *, size=9, align="left", cls="lbl", pad=0.0, fit=False, cut=0.0):
    """Надпись (статичный текст) в прямоугольнике (x0,y0,w,h) — мм от верха штампа.

    align: 'left' | 'center' | 'right' (горизонтальное выравнивание).
    pad:   отступ слева/справа (мм); применяется при выравнивании влево.
    fit:   True — автоподгонка кегля, чтобы текст влезал в ячейку (через JS).
    cut:   срез сверху (мм); координаты сдвигаются вверх.
    По вертикали текст центрируется.
    """
    if y0 < cut:
        return ""
    if cls == "val":
        color = "#111"
    else:
        color = "#333"
    justify = {"left": "flex-start", "center": "center", "right": "flex-end"}[align]
    textalign = {"left": "left", "center": "center", "right": "right"}[align]
    style = (f'position:absolute;left:{x0}mm;top:{y0 - cut}mm;'
             f'width:{w}mm;height:{h}mm;'
             f'display:flex;align-items:center;justify-content:{justify};'
             f'font-size:{size}pt;color:{color};'
             f'white-space:nowrap;overflow:hidden;text-align:{textalign};')
    if pad:
        style += f'padding-left:{pad}mm;padding-right:{pad}mm;'
    cls_attr = ' class="fit-lbl"' if fit else ""
    return f'<div{cls_attr} style="{style}">{esc(text)}</div>'

def _cell_field(x0, y0, w, h, key, value, *, size=9, align="left", pad=0.0, cut=0.0,
                placeholder="", title="", usage="", wrap=False, field_style=""):
    """Редактируемое поле в ячейке (x0,y0,w,h) — мм от верха штампа.

    wrap=True — разрешить перенос строк (для графы 9: длинное наименование
    организации раскладывается на строки, а блок текста центрируется по высоте
    ячейки flex-выравниванием align-items:center).
    """
    if y0 < cut:
        return ""
    justify = {"left": "flex-start", "center": "center", "right": "flex-end"}[align]
    wrap_css = "white-space:normal;word-break:break-word;overflow-wrap:anywhere;" \
        if wrap else "white-space:nowrap;"
    style = (f'position:absolute;left:{x0}mm;top:{y0 - cut}mm;'
             f'width:{w}mm;height:{h}mm;'
             f'display:flex;align-items:center;justify-content:{justify};'
             f'font-size:{size}pt;'
             f'{wrap_css}overflow:hidden;')
    if pad:
        style += f'padding-left:{pad}mm;padding-right:{pad}mm;'
    field = _field(key, value, size=size, placeholder=placeholder, title=title,
                   usage=usage, align=align, extra_style=field_style)
    return f'<div class="fit-lbl" style="{style}">{field}</div>'

def _cell_logo(x0, y0, w, h, png_path, *, pad=0.0, cut=0.0, title="",
               usage="") -> str:
    """Логотип организации в ячейке графы 9 (x0,y0,w,h) — мм от верха штампа.

    Автомасштабирование под ячейку выполняет CSS (object-fit:contain) — логотип
    занимает максимум места, пропорции сохраняются, выход за границы ячейки
    невозможен. png_path — растеризованный PDF-логотип (app.core.logo).
    """
    if y0 < cut:
        return ""
    style = (f'position:absolute;left:{x0}mm;top:{y0 - cut}mm;'
             f'width:{w}mm;height:{h}mm;display:flex;'
             f'align-items:center;justify-content:center;overflow:hidden;')
    if pad:
        style += f'padding:{pad}mm;'
    attrs = ""
    if title:
        attrs += f' data-title="{esc(title)}"'
    if usage:
        attrs += f' data-usage="{esc(usage)}"'
    # Логотип встраивается как data-URI: страница предпросмотра грузится через
    # setHtml(about:blank), и ссылку `file:///` QtWebEngine не отдаёт — картинка
    # молча не отображалась.
    src = signature_src(png_path)
    if not src:
        return ""
    return (f'<div class="cell9-logo"{attrs} style="{style}">'
            f'<img class="stamp-logo" src="{src}" alt="Логотип организации">'
            f'</div>')

def _cell9(snapshot, x0, y0, w, h, *, cut=0.0) -> str:
    """Графа 9 основной надписи: логотип организации ИЛИ текст по центру ячейки.

    Режим берётся из снимка (snapshot["cell9"], формируется по модели
    project.stamp_cell9). Если логотип не задан или не читается — выводится
    текст (реквизит «Организация»), выровненный по центру ячейки и по
    горизонтали, и по вертикали.
    """
    c9 = snapshot.get("cell9") or {}
    # Текст графы 9 — единый реквизит «Организация» (правится inline прямо в
    # графе; см. FIELD_USAGE), поэтому значение всегда согласовано со штампом.
    req = snapshot.get("requisites") or {}
    text = req.get("organization") or ""
    png = c9.get("png") or ""
    if c9.get("mode") == "logo" and png:
        pad = c9.get("padding_mm", 0.4)
        try:
            pad = max(0.0, float(pad))
        except (TypeError, ValueError):
            pad = 0.4
        return _cell_logo(x0, y0, w, h, png, pad=pad, cut=cut,
                          title="Графа 9 — логотип организации",
                          usage="Основная надпись (графа 9)")
    # Ключ поля — "organization": правка текста прямо в графе 9 меняет
    # наименование организации во всех листах (единый источник данных).
    # Ячейка помечена fit-lbl: длинное наименование автоматически уменьшает
    # кегль (JS-автоподгонка), не вылезая за рамку графы.
    # Центрирование по ГОРИЗОНТАЛИ: поле растягивается на всю ширину ячейки
    # (display:block + width:100% + text-align:center), иначе span с min-width
    # прижимался бы к левому краю flex-контейнера.
    # Центрирование по ВЕРТИКАЛИ: контейнер выше — flex с align-items:center;
    # длинное наименование переносится на строки (white-space:normal) и блок
    # текста остаётся центрированным по высоте ячейки.
    return _cell_field(x0, y0, w, h, "organization", text, size=CELL9_PT,
                       align="center", cut=cut,
                       title="Графа 9 — организация",
                       usage="Основная надпись (графа 9)",
                       wrap=True,
                       field_style="display:block;width:100%;white-space:normal;")

def _stamp_signature_rows(snapshot, form: int) -> list:
    """Строки подписей левого блока штампа для формы 5 (текст) или 3 (графика).

    Строки приходят в снимке (preview_model.build_preview_snapshot) — единый
    источник значений для предпросмотра, Word и PDF.
    """
    kind = snapshot.get("section") or ""
    table = (snapshot.get("stamp_signatures") or {}).get(kind) or {}
    rows = table.get(str(5 if form != 3 else 3))
    return list(rows or [])

def _sig_signer(snapshot, surname: str, row: dict | None = None) -> dict:
    """Данные подписанта по фамилии (для изображения подписи в графе).

    Подпись подтягивается из реестра подписантов по фамилии: правка фамилии
    в графе сразу показывает подпись соответствующего подписанта. Готовая
    ссылка на изображение приходит в строке (`signature_image`) — она уже
    подобрана по фамилии при сборке снимка.
    """
    image = str((row or {}).get("signature_image") or "")
    if image:
        return {"signature_image": image}
    target = (surname or "").strip().lower()
    if not target:
        return {}
    for signer in (snapshot.get("signers") or []):
        if (signer.get("surname") or "").strip().lower() == target:
            return signer
    return {}

def _cell_sig_image(x0, y0, w, h, key, signer, *, size=7, pad=0.5, cut=0.0,
                    hint=True) -> str:
    """Ячейка графы «Подп.»: изображение подписи, подобранной по фамилии.

    Клик по ячейке открывает выбор подписанта (data-kind="signature"): выбранная
    фамилия записывается в строку, а изображение берётся из реестра подписантов.

    :param hint: подсказка «(подпись)» при отсутствии подписи — только для
        интерактивного предпросмотра (см. _sig_field).
    """
    if y0 < cut:
        return ""
    src = signature_src((signer or {}).get("signature_image") or "")
    if src:
        inner = f'<img class="sig-img sig-img-row" src="{src}" alt="">'
    else:
        inner = ('<span class="sig-placeholder">(подпись)</span>'
                 if hint else '')
    style = (f'position:absolute;left:{x0}mm;top:{y0 - cut}mm;'
             f'width:{w}mm;height:{h}mm;'
             f'display:flex;align-items:center;justify-content:center;'
             f'font-size:{size}pt;'
             # Подпись может выходить за ячейку до 1.5 размеров
             # (SIGNATURE_OVERFLOW): фон картинки прозрачный, поэтому линии
             # сетки штампа остаются целыми.
             f'overflow:visible;z-index:2;')
    if pad:
        style += f'padding-left:{pad}mm;padding-right:{pad}mm;'
    return (f'<div class="fit-lbl" style="{style}">'
            f'<span class="field signature sig-cell" data-key="{esc(key)}" '
            f'data-kind="signature" data-size="{size}" '
            f'data-title="Подпись" contenteditable="false">{inner}</span></div>')

def _sig_cells(snapshot, form: int, *, size=7, pad=1.0, date_pad=0.5,
               cut=0.0) -> list:
    """Ячейки строк подписей левого блока: должность | фамилия | подпись | дата.

    Координаты колонок — по сетке ГОСТ (0..20 должность, 20..40 фамилия,
    40..55 графа «Подп.», 55..65 дата), одинаковой для форм 5 и 3. Должность
    по умолчанию берётся из формы, фамилия — из реестра подписантов (по роли
    строки); все ячейки редактируются в предпросмотре.
    """
    kind = snapshot.get("section") or ""
    out = []
    for row in _stamp_signature_rows(snapshot, form):
        y = float(row.get("y") or 0.0)
        index = int(row.get("row") or 0)
        post = str(row.get("post") or "")
        surname = str(row.get("surname") or "")
        date = str(row.get("date") or "")
        out.append(_cell_field(0.0, y, 20.0, 5.0,
                               f"sp|{kind}|{index}|post", post,
                               size=size, pad=pad, cut=cut,
                               title="Должность подписи", align="left"))
        out.append(_cell_field(20.0, y, 20.0, 5.0,
                               f"ss|{kind}|{index}|surname", surname,
                               size=size, pad=pad, cut=cut,
                               title="Фамилия (подпись — из реестра подписантов)",
                               align="left"))
        out.append(_cell_sig_image(40.0, y, 15.0, 5.0,
                                   f"ss|{kind}|{index}|signature",
                                   _sig_signer(snapshot, surname, row),
                                   size=size, cut=cut,
                                   hint=not snapshot.get("hide_signature_hint")))
        out.append(_cell_field(55.0, y, 10.0, 5.0,
                               f"sp|{kind}|{index}|date", date,
                               size=size, pad=date_pad, cut=cut,
                               title="Дата подписи", align="left"))
    return out

def _stamp_change_table(snapshot) -> dict:
    """Таблица изменений раздела из снимка: строки и даты подписей.

    Ключ раздела — snapshot["section"] (содержание, ПЗ и т.д.), поэтому у каждого
    раздела своя таблица изменений, как и в самой основной надписи.
    """
    kind = snapshot.get("section") or ""
    tables = snapshot.get("stamp_changes") or {}
    table = tables.get(kind) or {}
    return {"rows": list(table.get("rows") or []),
            "sign_dates": dict(table.get("sign_dates") or {})}

def _sheet_designation(snapshot) -> str:
    """Обозначение в графе 1: обозначение тома + доп. обозначение раздела.

    Например «02-2026-ТР» + «.ПЗ» = «02-2026-ТР.ПЗ». Код раздела берётся из
    project.section_codes по ключу раздела листа (snapshot["section"]).
    """
    base = (snapshot.get("requisites", {}).get("volume_id") or "").strip()
    codes = snapshot.get("section_codes") or {}
    code = (codes.get(snapshot.get("section", "")) or "").strip()
    if not base:
        return code
    return f"{base}{code}" if code else base

def _form3_labels(snapshot) -> str:
    """Надписи/поля формы 3 ГОСТ (графический лист, 185x55). Значения — редактируемые поля."""
    req = snapshot["requisites"]
    L = []
    # ---- Левый блок (0..65): таблица изменений, 11 рядов по 5 мм ----
    # Заголовок таблицы изменений (ряд y=20..25), 6 колонок.
    chg_hdr = [("Изм.", 0, 10), ("Кол.уч", 10, 10), ("Лист", 20, 10),
               ("№док.", 30, 10), ("Подпись", 40, 15), ("Дата", 55, 10)]
    for text, x0, w in chg_hdr:
        L.append(_lbl(x0, 20.0, w, 5.0, text, size=8, align="left", cls="val", pad=1.0, fit=True))
    # Строки подписей: должность | фамилия | подпись | дата (y=25..55).
    # Первая строка («Разраб.») одновременно служит строкой значений таблицы
    # изменений — как в образце формы.
    L.extend(_sig_cells(snapshot, 3, size=8, pad=1.0, date_pad=1.0))

    # ---- Правый блок (65..185) ----
    # Графа обозначения (y=0..10) — вся ширина. Обозначение тома + код раздела.
    L.append(_cell_field(65.0, 0.0, 120.0, 10.0, "volume_id", _sheet_designation(snapshot),
                         size=12, align="center", title="Обозначение"))
    # Графы 2/3 – наименование (y=10..25).
    L.append(_cell_field(65.0, 10.0, 120.0, 15.0, "object_name", req.get("object_name", ""),
                         size=9, align="center", title="Наименование объекта"))
    # Заголовки Стадия | Лист | Листов (y=25..30), колонки 135-150-165-185.
    # Надписи и значения граф «Стадия / Лист / Листов» — кегль 12 (общий размер).
    L.append(_lbl(135.0, 25.0, 15.0, 5.0, "Стадия", size=STAGE_CELL_PT, align="center", cls="lbl"))
    L.append(_lbl(150.0, 25.0, 15.0, 5.0, "Лист", size=STAGE_CELL_PT, align="center", cls="lbl"))
    L.append(_lbl(165.0, 25.0, 20.0, 5.0, "Листов", size=STAGE_CELL_PT, align="center", cls="lbl"))
    # Значения П | 1 | 6 (y=30..40).
    L.append(_cell_field(135.0, 30.0, 15.0, 10.0, "stage", req.get("stage", ""), size=STAGE_CELL_PT, align="center", title="Стадия"))
    L.append(_cell_field(150.0, 30.0, 15.0, 10.0, "sheet", snapshot.get("sheet_no", ""), size=STAGE_CELL_PT, align="center", title="Лист"))
    L.append(_cell_field(165.0, 30.0, 20.0, 10.0, "sheets_total", snapshot.get("sheets_total", ""), size=STAGE_CELL_PT, align="center", title="Листов"))
    # Графа наименования листа (y=40..55, колонка 65..135).
    L.append(_cell_field(65.0, 40.0, 70.0, 15.0, "volume_name", req.get("volume_name", ""),
                         size=10, align="center", title="Наименование листа"))
    # Графа 9 — наименование/индекс организации (y=40..55, колонка 135..185).
    # Раньше здесь была заглушка-надпись «domus».
    L.append(_cell9(snapshot, 135.0, 40.0, 50.0, 15.0))
    return "".join(L)

# Срезаемая сверху полоса (2 верхние строки формы 5): обозначение + наименование.
FORM5_CUT = 25.0

def _form5_labels(snapshot) -> str:
    """Надписи/поля формы 5 ГОСТ (первый/заглавный лист текста, 185x40). Значения — редактируемые поля."""
    req = snapshot["requisites"]
    L = []
    # ---- Левый блок (0..65): таблица изменений + строки подписей ----
    # Заголовок таблицы изменений (ряд y=10..15), 6 колонок.
    chg_hdr = [("Изм.", 0, 10, 1.0), ("Кол.уч", 10, 10, 0.5), ("Лист", 20, 10, 0.5),
               ("№док.", 30, 10, 0.5), ("Подп.", 40, 15, 0.5), ("Дата", 55, 10, 0.5)]
    for text, x0, w, pad in chg_hdr:
        L.append(_lbl(x0, 10.0, w, 5.0, text, size=7, align="left", cls="val", pad=pad, fit=True))
    # Строки подписей: должность | фамилия | подпись | дата (y=15..40).
    # Первая строка («Разраб.») одновременно служит строкой значений таблицы
    # изменений — так же, как в образце формы.
    L.extend(_sig_cells(snapshot, 5, size=7, pad=1.0, date_pad=0.5))

    # ---- Правый блок (65..185) ----
    # Обозначение документа (y=0..15) — вся ширина правого блока.
    L.append(_cell_field(65.0, 0.0, 120.0, 15.0, "volume_id", _sheet_designation(snapshot),
                         size=12, align="center", title="Обозначение"))
    # Заголовки Стадия | Лист | Листов (y=15..20), колонки 135-150-165-185.
    # Надписи и значения граф «Стадия / Лист / Листов» — кегль 12 (общий размер).
    L.append(_lbl(135.0, 15.0, 15.0, 5.0, "Стадия", size=STAGE_CELL_PT, align="center", cls="lbl", fit=True))
    L.append(_lbl(150.0, 15.0, 15.0, 5.0, "Лист", size=STAGE_CELL_PT, align="center", cls="lbl", fit=True))
    L.append(_lbl(165.0, 15.0, 20.0, 5.0, "Листов", size=STAGE_CELL_PT, align="center", cls="lbl", fit=True))
    # Значения П | 1 | 18 (y=20..25).
    L.append(_cell_field(135.0, 20.0, 15.0, 5.0, "stage", req.get("stage", ""), size=STAGE_CELL_PT, align="center", title="Стадия"))
    L.append(_cell_field(150.0, 20.0, 15.0, 5.0, "sheet", snapshot.get("sheet_no", ""), size=STAGE_CELL_PT, align="center", title="Лист"))
    L.append(_cell_field(165.0, 20.0, 20.0, 5.0, "sheets_total", snapshot.get("sheets_total", ""), size=STAGE_CELL_PT, align="center", title="Листов"))
    # Наименование документа «Пояснительная записка» (колонка 65..135, y=15..40).
    L.append(_cell_field(65.0, 15.0, 70.0, 25.0, "volume_name", req.get("volume_name", ""),
                         size=13, align="center", title="Наименование"))
    # Графа 9 — наименование/индекс организации: ячейка 135..185 x 25..40 мм.
    # Вертикали x=150/165 обрываются на y=25, поэтому ниже строки
    # «Стадия | Лист | Листов» (y=15..25) правый блок — одна ячейка графы 9.
    # По образцу (Форма 3.pdf) текст организации центрируется в ней по обеим
    # осям и переносится на несколько строк.
    L.append(_cell9(snapshot, 135.0, 25.0, 50.0, 15.0))
    return "".join(L)

def _form3_stamp(snapshot, h, design, sheet_cell, org, obj, dev, chk, nrm, cut=0.0) -> str:
    """Форма 3 ГОСТ (графический лист, 185 x 55) — сетка линий + надписи."""
    lines = []
    for x, spans in FORM3_VLINES:
        for y0, y1 in spans:
            lines.append(_vline(x, y0, y1))
    for y, spans in FORM3_HLINES:
        for x0, x1 in spans:
            lines.append(_hline(y, x0, x1))
    return (f'<div class="stamp" style="height:{FORM3_H}mm;width:{FORM3_W}mm;border:none;">'
            f'{ "".join(lines) }{ _form3_labels(snapshot) }</div>')

def _form5_stamp(snapshot, h, design, sheet_cell, org, obj, dev, chk, nrm) -> str:
    """Форма 5 ГОСТ (первый/заглавный лист текста, 185 x 40) — сетка линий + надписи."""
    lines = []
    for x, spans in FORM5_VLINES:
        for y0, y1 in spans:
            lines.append(_vline(x, y0, y1))
    for y, spans in FORM5_HLINES:
        for x0, x1 in spans:
            lines.append(_hline(y, x0, x1))
    return (f'<div class="stamp" style="height:{FORM5_H}mm;width:{FORM5_W}mm;border:none;">'
            f'{ "".join(lines) }{ _form5_labels(snapshot) }</div>')

def _form6_labels(snapshot) -> str:
    """Надписи/поля формы 6 (185x15). Значения — редактируемые поля.

    Левый блок по образцу «Форма 6.pdf»: сверху два ряда значений таблицы
    изменений (y 0..5 и 5..10), подписи — в НИЖНЕЙ полосе (y 10..15).
    """
    req = snapshot["requisites"]
    L = []
    # ---- Левый блок (0..65): подписи таблицы изменений (полоса y=10..15) ----
    chg = [("Изм.", 0, 10), ("Кол.уч", 10, 10), ("Лист", 20, 10),
           ("№док.", 30, 10), ("Подп.", 40, 15), ("Дата", 55, 10)]
    for text, x0, w in chg:
        L.append(_lbl(x0, 10.0, w, 5.0, text, size=7, align="left", cls="val", pad=0.5, fit=True))
    # ---- Правый блок (65..185) ----
    # Обозначение документа (зона 65..175, y=0..15) — по центру.
    L.append(_cell_field(65.0, 0.0, 110.0, 15.0, "volume_id", _sheet_designation(snapshot),
                         size=11, align="center", title="Обозначение"))
    # Зона «Лист» (175..185): заголовок y=0..7, значение y=7..15.
    # Надпись и значение — кегль 12 (общий размер).
    L.append(_lbl(175.0, 0.0, 10.0, 7.0, "Лист", size=STAGE_CELL_PT, align="center", cls="lbl"))
    L.append(_cell_field(175.0, 7.0, 10.0, 8.0, "sheet", snapshot.get("sheet_no", ""),
                         size=STAGE_CELL_PT, align="center", title="Лист"))
    return "".join(L)

def _form6_stamp(snapshot, h, design, sheet_cell, org, obj, dev, chk, nrm) -> str:
    """Форма 6 (185 x 15) — сетка 1:1 из «Форма 6.pdf» + надписи."""
    lines = []
    for x, spans in FORM6_VLINES:
        for y0, y1 in spans:
            lines.append(_vline(x, y0, y1))
    for y, spans in FORM6_HLINES:
        for x0, x1 in spans:
            lines.append(_hline(y, x0, x1))
    return (f'<div class="stamp" style="height:{FORM6_H}mm;width:{FORM6_W}mm;border:none;">'
            f'{ "".join(lines) }{ _form6_labels(snapshot) }</div>')

def _stamp_html(form, snapshot, font_map: FontMap) -> str:
    """Основная надпись по форме (1/2/2а) как точная ГОСТ-сетка.

    form: 1 (графика = форма 3 ГОСТ, 185x55),
          2 (первый лист текста = форма 5, 185x40),
          '2a' (последующий лист = форма 6, 185x15).
    """
    req = snapshot["requisites"]
    roles = snapshot["roles"]
    h = FORM_HEIGHT.get(form, 40.0)
    sheet_no = snapshot.get("sheet_no", "")
    change_w = sum(CHANGE_COLS)

    design = _field("volume_id", req.get("volume_id", ""), size=11, usage="Обозначение тома")
    sheet_cell = _field("sheet", sheet_no, size=9)
    org = _field("organization", req.get("organization", ""), size=8, usage="Организация")
    obj = _field("object_name", req.get("object_name", ""), size=8, usage="Наименование здания")

    # Подсказка «(подпись)» — только в интерактивном предпросмотре: в собранном
    # томе графа «Подп.» без выбранной подписи остаётся пустой.
    hint = not snapshot.get("hide_signature_hint")
    dev = _sig_field("develop", "Разработал", roles.get("develop"), hint=hint)
    chk = _sig_field("check", "Проверил", roles.get("check"), hint=hint)
    nrm = _sig_field("norm", "Нормоконтроль", roles.get("norm"), hint=hint)

    # --- Форма 6 (2а): 185x15 — сетка 1:1 из «Форма 6.pdf» + надписи ---
    if form == "2a":
        return _form6_stamp(snapshot, h, design, sheet_cell, org, obj, dev, chk, nrm)

    # --- Форма 5 (2): 185x55 — сетка 1:1 из «Форма 5.pdf» + надписи ---
    if form == 2:
        return _form5_stamp(snapshot, h, design, sheet_cell, org, obj, dev, chk, nrm)

    # --- Форма 3 (1): 185x40 — сетка 1:1 из «Форма 3.pdf», только линии ---
    if form == 1:
        return _form3_stamp(snapshot, h, design, sheet_cell, org, obj, dev, chk, nrm)

# ---- Листы ----------------------------------------------------------------

def render_cover(snapshot, font_map) -> tuple:
    """Обложка (приложение Е). Возвращает (content, stamp)."""
    req = snapshot["requisites"]
    org = _field("organization", req.get("organization", ""), size=18, cls="logo",
                 placeholder="Наименование организации", usage="Обложка, все листы")
    vid = _field("volume_id", req.get("volume_id", ""), size=14,
                 placeholder="Обозначение тома", usage="Обложка, все листы")
    volname = _field("volume_name", req.get("volume_name", ""), size=14,
                     placeholder="Наименование тома")
    obj = _field("object_name", req.get("object_name", ""), size=12,
                 placeholder="Наименование объекта")
    # Полное наименование стадии («Проектная документация»); пустое значение
    # показывается подсказкой. В основной надписи выводится «П» / «Р».
    stage = _field("stage", req.get("stage_label", ""), size=12,
                   placeholder="Стадия", usage="Обложка, титульный лист")
    city = _field("city", req.get("city", ""), size=12, placeholder="Город")
    # Год по умолчанию 0 — незаполненный год показывается подсказкой, а не «0».
    year = _field("year", req.get("year") or "", size=12, placeholder="Год")
    content = f"""
    <div class="cover">
      <div class="logo">{org}</div>
      <div style="height:30mm"></div>
      <h1>ОБЛОЖКА</h1>
      <div style="height:15mm"></div>
      <div style="font-size:14pt">{vid}</div>
      <div style="height:10mm"></div>
      <div style="font-size:14pt">{volname}</div>
      <div style="height:10mm"></div>
      <div style="font-size:12pt">{obj}</div>
      <div style="height:20mm"></div>
      <div style="font-size:12pt">{stage}</div>
      <div style="font-size:12pt">{city} — {year}</div>
    </div>
    """
    return content, ""

def render_title(snapshot, font_map) -> tuple:
    """Титульный лист (приложение Ж) — без основной надписи (штампа)."""
    req = snapshot["requisites"]
    roles = snapshot["roles"]
    org = _field("organization", req.get("organization", ""), size=12, cls="logo")
    obj = _field("object_name", req.get("object_name", ""), size=18)
    volname = _field("volume_name", req.get("volume_name", ""), size=16)
    # Обозначение тома — самостоятельная строка без надписи-префикса (ГОСТ,
    # приложение Ж). Пока значение пустое, объяснение даёт подсказка поля.
    vid = _field("volume_id", req.get("volume_id", ""), size=12,
                 placeholder="Обозначение тома")
    # Стадия — полное наименование, без надписи-префикса (как обозначение).
    stage = _field("stage", req.get("stage_label", ""), size=12,
                   placeholder="Стадия", usage="Обложка, титульный лист")
    city = _field("city", req.get("city", ""), size=12)
    year = _field("year", req.get("year") or "", size=12)
    approve = _sig_field("approve", "Утвердил", roles.get("approve"),
                         hint=not snapshot.get("hide_signature_hint"))
    content = f"""
    <div class="cover">
      <div style="font-size:12pt; margin-top:8mm">{org}</div>
      <div style="height:20mm"></div>
      <div style="font-size:18pt">{obj}</div>
      <div style="height:8mm"></div>
      <div style="font-size:16pt">{volname}</div>
      <div style="height:12mm"></div>
      <div style="font-size:12pt">{vid}</div>
      <div style="height:12mm"></div>
      <div style="font-size:12pt">{stage}</div>
      <div style="height:8mm"></div>
      <div style="font-size:12pt">{city} — {year}</div>
      <div style="height:20mm"></div>
      <div style="text-align:left; font-size:11pt">Утвердил: {approve}</div>
    </div>
    """
    # Титульный лист — без основной надписи (штампа).
    return content, ""

def _cols_html(widths) -> str:
    """<colgroup> таблицы во всю ширину листа: колонки в процентах от суммы.

    Ширины задаются в миллиметрах (models.TOC_WIDTHS / SHEETLIST_WIDTHS), чтобы
    они были единым источником для Word и PDF. В HTML те же мм пересчитываются
    в проценты: сумма 185 мм не переводится в целое число пикселей, и браузер
    округлял ширину таблицы вверх на 1 px (0.26 мм) — правая линия таблицы
    выходила за линию рамки и на печати давала утолщение. Проценты считаются от
    ширины блока .bleed-x (ровно 185 мм между линиями рамки), поэтому правая
    линия таблицы совпадает с рамкой.
    """
    total = float(sum(widths)) or 1.0
    return "".join(f'<col style="width:{w / total * 100:.4f}%">' for w in widths)

def render_content(snapshot, font_map) -> tuple:
    """Содержание / опись — автоматический сбор по фактической структуре тома.

    Строки формируются единым источником (app.core.toc) и отражают сквозные
    номера листов («Лист N — M»). Если автоматических строк нет, используется
    ручной перечень requisites.sections (обратная совместимость).
    """
    rows = snapshot.get("toc_rows") or []
    if not rows:
        rows = [(s, "", "") for s in snapshot.get("sections", [])]
    lines = []
    for name, des, sheets in rows:
        lines.append(
            f'<tr class="toc-row">'
            f'<td class="toc-name">{esc(name)}</td>'
            f'<td class="toc-des">{esc(des)}</td>'
            f'<td class="toc-sheets">{esc(sheets)}</td>'
            f'</tr>')
    # Ширины колонок — единый источник с Word (models.TOC_WIDTHS): Наименование
    # 95 | Обозначение 50 | Листы 40 мм. Сумма 185 мм — вся ширина между
    # рамками листа, поэтому блок выведен с классом bleed-x.
    from ..models import TOC_WIDTHS
    cols_html = _cols_html(TOC_WIDTHS)
    content = f"""
    <div class="bleed-x">
      <h1>СОДЕРЖАНИЕ</h1>
      <div style="height:8mm"></div>
      <table class="toc-table">
        <colgroup>{cols_html}</colgroup>
        <thead>
          <tr class="toc-head">
            <th class="toc-name">Наименование</th>
            <th class="toc-des">Обозначение</th>
            <th class="toc-sheets">Листы</th>
          </tr>
        </thead>
        <tbody>{''.join(lines)}</tbody>
      </table>
    </div>
    """
    return content, _stamp_html(2, snapshot, font_map)

def _toc_font_map() -> FontMap:
    """Карта кеглей конфигурации (для листа оглавления ПЗ).

    Предпросмотр получает только снимок значений и карту кеглей листа, поэтому
    кегль строк оглавления (toc_row) читается из конфигурации. Если файл
    конфигурации недоступен — используются значения по умолчанию.
    """
    try:
        return load_config().font_map
    except Exception:
        return FontMap()

def _graphic_sheet_list_cfg(project):
    """Настройки листа «Графическая часть» из проекта (с умолчаниями).

    Единый помощник для предпросмотра и Word: если проект сохранён старой
    версией (поля нет) — возвращаются умолчания, а не падение.
    """
    from ..models import GraphicSheetList
    cfg = getattr(project, "graphic_sheet_list", None)
    if not isinstance(cfg, GraphicSheetList):
        return GraphicSheetList()
    return cfg

def build_graphic_sheet_list(project) -> list:
    """Строки перечня листов графической части: [(обозначение, наименование, примечание)].

    Источник задаётся в настройках проекта (models.GraphicSheetList):
      * pdf   — название и номер листа читаются из самого PDF по геометрии
        штампа (правый нижний угол); обозначение собирается по шаблону
        «шифр + л. + номер». Лист, у которого название или номер в своих
        квадратах не найдены, в перечень НЕ ПОПАДАЕТ;
      * excel — читается из указанной книги (app.core.xlsx_reader);
      * auto  — по фактическим листам раздела: обозначение = обозначение
        раздела (том + код «.ГЧ»), наименование — имя файла листа.

    Строки строятся по ОДНОМУ правилу для предпросмотра, Word и PDF, поэтому
    таблица на листе и в собранном томе совпадает.
    """
    from ..models import GraphicSheetList, SECTION_GRAPHIC, DocKind, section_designation
    cfg = _graphic_sheet_list_cfg(project)
    if cfg.use_excel():
        try:
            from .xlsx_reader import read_sheet_list
            rows = read_sheet_list(cfg.excel_path, cfg.sheet_name)
            if rows:
                return rows
        except Exception:
            pass
    if cfg.use_pdf():
        rows = _graphic_rows_from_pdf(project, cfg)
        if rows:
            return rows
        # Ни один лист не оформлен (нет названия/номера в штампе) — перечень
        # по PDF пуст. Ниже он не подменяется именами файлов: пользователь
        # ожидает именно данные штампа.
        return []
    # Автоматический перечень: по листам раздела. Обозначение листа — то же,
    # что в содержании тома (том + код раздела); наименование — имя файла,
    # без расширения (в файле графики обычно уже лежит осмысленное имя).
    rows = []
    base = section_designation(project, SECTION_GRAPHIC)
    try:
        docs = project.documents_for(DocKind.GRAPHIC)
    except Exception:
        docs = []
    for d in docs:
        path = getattr(d, "path", "") or ""
        name = os.path.basename(path) if path else ""
        root, ext = os.path.splitext(name)
        if ext.lower() in (".pdf", ".docx", ".doc"):
            name = root
        rows.append((base, name or "—", ""))
    return rows

#: Кэш строк перечня листов ГЧ, прочитанных из PDF. Ключ — контрольная сумма
#: состояния файлов (путь, размер, первые/последние байты) + области штампа и
#: шаблон обозначения. Пока PDF-листы графической части не менялись, штампы
#: повторно не разбираются: чтение сотен страниц — самая дорогая часть
#: построения перечня.
_SHEETLIST_CACHE: dict = {}
_SHEETLIST_CACHE_LIMIT = 8

def _pdf_content_key(path: str) -> tuple | None:
    """Контрольная сумма «состояния» PDF: путь, размер и крайние байты.

    Читается только начало и конец файла (не весь файл), поэтому проверка
    дешёвая даже для PDF в сотни мегабайт. Смена содержимого меняет размер
    или крайние блоки — этого достаточно, чтобы понять, что листы изменились
    и перечень надо перестроить.
    """
    if not path or not os.path.exists(path):
        return None
    try:
        size = os.path.getsize(path)
        with open(path, "rb") as f:
            head = f.read(4096)
            if size > 8192:
                f.seek(-4096, os.SEEK_END)
                tail = f.read(4096)
            else:
                tail = head
    except OSError:
        return None
    import hashlib
    digest = hashlib.md5()
    digest.update(head)
    digest.update(tail)
    digest.update(str(size).encode("ascii"))
    return (os.path.abspath(path), size, digest.hexdigest())

def clear_sheetlist_cache() -> None:
    """Сбрасывает кэш перечня листов ГЧ (например, после пересборки раздела)."""
    _SHEETLIST_CACHE.clear()

def _graphic_rows_from_pdf(project, cfg) -> list:
    """Строки перечня по данным штампов исходных PDF-файлов графики.

    Для каждого PDF раздела берутся только ОФОРМЛЕННЫЕ листы — те, у которых
    в заданных квадратах штампа нашлись и название, и номер
    (app.core.sheet_stamp). Порядок строк — порядок файлов в разделе и
    страниц в файле.

    Обозначение собирается по шаблону из настроек: по умолчанию
    «шифр + л. + номер» (например «A1-003.24-КР л.16»).

    Результат кэшируется по контрольной сумме PDF-файлов: перечень
    перестраивается только когда листы графической части действительно
    изменились.
    """
    from ..models import SECTION_GRAPHIC, DocKind, section_designation
    from .sheet_stamp import read_sheets_from_pdf
    volume = section_designation(project, SECTION_GRAPHIC)
    try:
        docs = project.documents_for(DocKind.GRAPHIC)
    except Exception:
        docs = []
    # Ключ кэша: состояние каждого файла + параметры чтения. Настройки входят
    # в ключ, иначе смена областей штампа или шаблона вернула бы старый набор.
    file_keys = []
    for d in docs:
        file_keys.append(_pdf_content_key(getattr(d, "path", "") or ""))
    key = (tuple(file_keys), tuple(cfg.name_box or ()), tuple(cfg.number_box or ()),
           str(cfg.designation_template or ""), volume)
    cached = _SHEETLIST_CACHE.get(key)
    if cached is not None:
        return list(cached)
    rows = []
    for d in docs:
        path = getattr(d, "path", "") or ""
        if not path or not os.path.exists(path):
            continue
        try:
            stamps = read_sheets_from_pdf(
                path, name_box=cfg.name_box, number_box=cfg.number_box,
                template=cfg.designation_template, volume=volume)
        except Exception:
            stamps = []
        for st in stamps:
            rows.append((st.designation, st.name, ""))
    if len(_SHEETLIST_CACHE) >= _SHEETLIST_CACHE_LIMIT:
        _SHEETLIST_CACHE.clear()
    _SHEETLIST_CACHE[key] = list(rows)
    return rows

def render_graphic_sheet_list(snapshot, font_map) -> tuple:
    """Лист «Графическая часть» с перечнем листов (рамка + осн. надпись формы 3).

    Таблица — во всю ширину листа, колонки: Обозначение | Наименование |
    Примечание. Строки приходят в снимке (snapshot["graphic_sheet_list_rows"]),
    поэтому предпросмотр, Word и PDF показывают одно и то же.
    """
    from ..models import GraphicSheetList, SHEETLIST_LABELS
    cfg = (snapshot.get("graphic_sheet_list") or {})
    title = str(cfg.get("title") or "Графическая часть")
    subtitle = str(cfg.get("subtitle") or "")
    rows = list(snapshot.get("graphic_sheet_list_rows") or [])

    body = []
    for designation, name, note in rows:
        body.append(
            f'<tr><td class="val" style="text-align:left">{esc(designation)}</td>'
            f'<td class="val" style="text-align:left">{esc(name)}</td>'
            f'<td class="val" style="text-align:left">{esc(note)}</td></tr>')
    if not body:
        body.append(f'<tr><td class="lbl" colspan="3">'
                    f'Перечень листов не задан</td></tr>')

    headers = "".join(f'<th class="lbl">{esc(label)}</th>'
                      for label in SHEETLIST_LABELS)
    subtitle_html = (f'<div style="font-size:11pt;margin-top:1mm">{esc(subtitle)}</div>'
                     if subtitle else "")
    # Ширины колонок заданы в мм (models.SHEETLIST_WIDTHS = 50/95/40): единый
    # источник с Word и PDF, поэтому таблица одинакова во всех выводах. Сумма
    # 185 мм — ровно ширина между линиями рамки листа (класс bleed-x ниже);
    # в HTML колонки переводятся в проценты, чтобы правая линия таблицы
    # совпала с рамкой (см. _cols_html).
    from ..models import SHEETLIST_WIDTHS
    cols_html = _cols_html(SHEETLIST_WIDTHS)
    content = f"""
    <div class="bleed-x">
      <div style="text-align:center">
        <div style="font-size:14pt;font-weight:bold">{esc(title)}</div>
        {subtitle_html}
      </div>
      <div style="height:6mm"></div>
      <table class="gsheetlist">
        <colgroup>{cols_html}</colgroup>
        <thead><tr>{headers}</tr></thead>
        <tbody>{''.join(body)}</tbody>
      </table>
    </div>
    """
    # Лист перечня — лист графической части, поэтому основная надпись формы 3.
    return content, _stamp_html(1, snapshot, font_map)

# ---- Состав проектной документации ------------------------------------------

#: Надпись над таблицей состава и её высота вместе с отбивкой, мм.
COMPOSITION_TITLE = "Состав проектной документации"
COMPOSITION_TITLE_MM = 12.0

def composition_rows_per_sheet(first: bool) -> int:
    """Сколько строк таблицы состава ПД входит на лист.

    Считается по фактической высоте листа и основной надписи: первый лист
    раздела несёт форму 5 (40 мм), последующие — форму 6 (15 мм). Область
    начинается ПОД квадратом сквозной нумерации (TEXT_TOP_INDENT_MM = 17 мм),
    поэтому квадрат не перекрывает надпись над таблицей. Из области вычитается
    надпись с отбивкой (COMPOSITION_TITLE_MM), затем шапка 15 мм; остаток
    делится на строки высотой 8 мм. По этому числу строится разбиение на
    листы, поэтому предпросмотр и собранный HTML дают одинаковое число листов.
    """
    stamp_h = FORM_HEIGHT[2 if first else "2a"]
    area = PAGE_H_MM - (PAGE_BOTTOM_MM + stamp_h + 2.0) - TEXT_TOP_INDENT_MM
    from ..models import COMPOSITION_HEAD_MM, COMPOSITION_ROW_MM
    area -= COMPOSITION_TITLE_MM
    return max(1, int((area - COMPOSITION_HEAD_MM) // COMPOSITION_ROW_MM))

def composition_pages_from_rows(rows) -> list:
    """Разбивает строки таблицы состава ПД на листы (минимум один лист)."""
    rows = list(rows or [])
    pages = []
    pos = 0
    first = True
    while True:
        capacity = composition_rows_per_sheet(first)
        pages.append(rows[pos:pos + capacity])
        pos += capacity
        first = False
        if pos >= len(rows):
            break
    return pages

def composition_pages(project) -> list:
    """Страницы таблицы «Состав проектной документации» по строкам модели."""
    from ..models import composition_rows
    return composition_pages_from_rows(composition_rows(project))

def render_composition(snapshot, font_map) -> tuple:
    """Лист «Состав проектной документации» (рамка + основная надпись формы 5/6).

    Над таблицей печатается надпись «Состав проектной документации», под ней —
    таблица во всю ширину листа: Номер тома 15 | Обозначение 60 | Наименование
    90 | Примечание 20 мм (сумма 185 мм — ширина между линиями рамки). Высота
    шапки 15 мм, высота строки 8 мм. Строки страницы приходят в снимке
    (snapshot["composition_rows"] + snapshot["composition_page"]), поэтому
    предпросмотр и собранный том показывают одно и то же.
    """
    from ..models import COMPOSITION_LABELS, COMPOSITION_WIDTHS
    rows = [tuple(r or ()) for r in (snapshot.get("composition_rows") or [])]
    pages = composition_pages_from_rows(rows)
    index = int(snapshot.get("composition_page", 0) or 0)
    index = min(max(0, index), len(pages) - 1)
    page_rows = pages[index]

    body = []
    for values in page_rows:
        cells = "".join(f'<td>{esc(v)}</td>' for v in values)
        body.append(f'<tr class="comp-row">{cells}</tr>')
    headers = "".join(f'<th>{esc(label)}</th>' for label in COMPOSITION_LABELS)
    cols_html = _cols_html(COMPOSITION_WIDTHS)
    # Классы: bleed-x — таблица упирается в линии рамки по ширине,
    # bleed-top — начинается сразу под квадратом сквозной нумерации: верхний
    # внутренний отступ .sheet-content (6 мм) снимается. Над таблицей печатается
    # только надпись раздела.
    content = f"""
    <div class="bleed-x bleed-top">
      <div class="comp-title">{esc(COMPOSITION_TITLE)}</div>
      <table class="comp-table">
        <colgroup>{cols_html}</colgroup>
        <thead><tr class="comp-head">{headers}</tr></thead>
        <tbody>{''.join(body)}</tbody>
      </table>
    </div>
    """
    # Первый лист раздела — форма 5, последующие — форма 6.
    return content, _stamp_html(2 if index == 0 else "2a", snapshot, font_map)

def render_pz_toc(snapshot, font_map) -> tuple:
    """Лист «ОГЛАВЛЕНИЕ» пояснительной записки (автоматически по уровням Word).

    Строки берутся из снимка предпросмотра (snapshot["pz_toc_items"], единый
    источник app.core.pz_toc): наименование заголовка, линия-выноска (отточие)
    и номер листа записки. Листы оглавления 2..N берут свой набор строк по
    номеру листа (snapshot["pz_toc_page"]).

    Штамп: первый лист записки — форма 5 (большой), последующие — форма 6.
    """
    from .pz_toc import items_from_dicts, paginate_items, toc_html
    fm = font_map or FontMap()
    try:
        toc_fm = _toc_font_map()
    except Exception:
        toc_fm = fm
    raw = snapshot.get("pz_toc_items")
    page_no = int(snapshot.get("pz_toc_page", 0) or 0)
    if page_no > 0:
        # Листы 2..N оглавления: текст листа уже отобран при разбивке.
        pages = paginate_items(items_from_dicts(raw))
        index = min(max(1, page_no), len(pages)) - 1
        items = pages[index]
    else:
        items = items_from_dicts(raw)
    # Строки оглавления набираются шрифтом и кеглем обычного текста записки
    # (config.TextFonts, группа «body») — требование к пояснительной записке.
    tf = _text_fonts()
    content = toc_html(items,
                       title_size=getattr(toc_fm, "content_header", 14),
                       body_size=tf.size("body"),
                       font_family=tf.font("body"),
                       offset=int(snapshot.get("pz_toc_offset", 0) or 0))
    form = 2 if bool(snapshot.get("toc_first")) else "2a"
    return content, _stamp_html(form, snapshot, font_map)

def render_changes(snapshot, font_map) -> tuple:
    """Лист регистрации изменений."""
    changes = snapshot.get("changes", [])
    rows = []
    for c in changes:
        rows.append(
            f'<tr><td class="val">{esc(c.get("number", ""))}</td>'
            f'<td class="val">{esc(c.get("date", ""))}</td>'
            f'<td class="val" style="text-align:left">{esc(c.get("description", ""))}</td>'
            f'<td class="val">{esc(c.get("target", ""))}</td>'
            f'<td class="val">{esc(c.get("signer_surname", ""))}</td>'
            f'<td class="val">{esc(c.get("status", ""))}</td></tr>')
    if not rows:
        rows.append('<tr><td colspan="6" class="lbl">Записи о внесении изменений отсутствуют</td></tr>')
    content = f"""
    <h1 style="font-size:12pt">ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ</h1>
    <div style="height:6mm"></div>
    <table style="width:100%; border-collapse:collapse; border:0.25mm solid #000;">
      <tr>
        <th class="lbl">Изм.</th><th class="lbl">Дата</th>
        <th class="lbl">Содержание изменения</th><th class="lbl">Раздел/Лист</th>
        <th class="lbl">Исполнитель</th><th class="lbl">Статус</th>
      </tr>
      {''.join(rows)}
    </table>
    """
    return content, _stamp_html("2a", snapshot, font_map)

def render_stamp_sheet(form, snapshot, font_map) -> tuple:
    """Отдельный лист с основной надписью (форма 3/5/6 ГОСТ)."""
    label = {1: "ФОРМА 3 — ГРАФИЧЕСКИЙ ЛИСТ", 2: "ФОРМА 5 — ПЕРВЫЙ ЛИСТ ТЕКСТА",
             "2a": "ФОРМА 6 — ПОСЛЕДУЮЩИЙ ЛИСТ ТЕКСТА"}.get(form, "")
    content = f'<h1 style="font-size:11pt">{label}</h1>'
    return content, _stamp_html(form, snapshot, font_map)

# ---- Сборка полной страницы ------------------------------------------------

SHEET_KINDS = ["cover", "title", "content", "stamp_f1", "stamp_f2", "stamp_f2a", "changes"]

def render_preview_error(kind: str, snapshot: dict, reason: str,
                         log_path: str = "") -> str:
    """Полный HTML-документ с листом-сообщением о сбое предпросмотра.

    Показывает причину вместо безликого «Не удалось отобразить содержимое
    файла», чтобы сбой можно было диагностировать по самому экрану.
    """
    snapshot = dict(snapshot or {})
    if not snapshot.get("requisites"):
        snapshot["requisites"] = {"volume_id": ""}
    note = f"Сбой предпросмотра: {reason}"
    if log_path:
        note += f"\nПодробности: {log_path}"
    kind_key = kind if kind in ("doc", "appendix", "section") else "section"
    fragment = _placeholder_sheet_fragment(kind_key, snapshot, FontMap(),
                                           note=note)
    return _preview_document([fragment], snapshot)

def render_placeholder(label: str, *, subtitle: str = "", path: str = "",
                       kind: str = "", note: str = "") -> str:
    """Заглушка-лист для документа/приложения/раздела (не стандартный лист ГОСТ).

    Используется, когда содержимое не удалось получить (нет файла,
    неподдерживаемый формат) или когда документ прочитан, но текста в нём нет.
    Показывает наименование, источник и причину.

    :param note: пояснение причины (пустой — общий текст «не удалось отобразить»).
    """
    import html as _html
    def e(v):
        return _html.escape(str(v if v is not None else ""))
    lines = []
    lines.append(f'<div style="font-size:16pt; font-weight:bold; margin-bottom:8mm">{e(label)}</div>')
    if subtitle:
        lines.append(f'<div style="font-size:12pt; color:#333; margin-bottom:6mm">{e(subtitle)}</div>')
    if path:
        lines.append(f'<div style="font-size:11pt; color:#555; margin-bottom:6mm">Источник: {e(path)}</div>')
    lines.append('<div style="font-size:11pt; color:#888; margin-top:20mm">'
                 + e(note or "Не удалось отобразить содержимое файла.")
                 + '</div>')
    return "".join(lines)

# --------------------------------------------------------------------------
# Рендер содержимого документов (DOCX -> текст/таблицы, PDF -> страницы-картинки)
# --------------------------------------------------------------------------

def _frame_rules() -> str:
    """Линии рамки листа (по ГОСТ: левое поле 20 мм, остальные 5 мм).

    Рисуются полосами с осью ровно на границе полей — тем же приёмом, что и
    сетка штампа, поэтому линии штампа и квадрата нумерации, попадающие на
    рамку, ложатся ровно на неё (без «двойной» линии).
    """
    x0, x1 = PAGE_LEFT_MM, PAGE_W_MM - PAGE_RIGHT_MM
    y0, y1 = PAGE_TOP_MM, PAGE_H_MM - PAGE_BOTTOM_MM
    return "".join([_rule_h(y0, x0, x1), _rule_h(y1, x0, x1),
                    _rule_v(x0, y0, y1), _rule_v(x1, y0, y1)])

def _sheet_number_rules(page_w_mm: float | None = None,
                        page_h_mm: float | None = None) -> str:
    """Линии квадрата сквозной нумерации листов (правый верхний угол листа).

    :param page_w_mm: фактическая ширина листа. None — A4. У листов своего
        формата (графическая часть, приложения) рамка и цифра привязаны к
        ФАКТИЧЕСКОМУ листу, поэтому координаты A4 дали бы квадрат не в углу,
        а в середине листа.
    :param page_h_mm: фактическая высота листа. None — A4.

    Рисуются ВСЕ четыре стороны, хотя верхняя и правая совпадают с линиями
    рамки: у квадрата есть белая подложка под цифру (z-index выше рамки), она
    перекрывает половину толщины линии рамки, и верхняя граница квадрата
    выглядела «непрорисованной». Линии квадрата лежат в слое выше подложки,
    поэтому восстановленная линия ложится точно на линию рамки (та же
    координата и толщина) и второй линии не образует.
    """
    w = float(page_w_mm) if page_w_mm else PAGE_W_MM
    x1 = w - SHEET_NUM_OFFSET_MM
    x0 = x1 - SHEET_NUM_BOX_MM
    y0 = SHEET_NUM_OFFSET_MM
    y1 = y0 + SHEET_NUM_BOX_MM
    return "".join([_rule_h(y0, x0, x1), _rule_h(y1, x0, x1),
                    _rule_v(x0, y0, y1), _rule_v(x1, y0, y1)])

def extract_stamp_html(fragment: str) -> str:
    """Основная надпись (div.stamp) из разметки готового листа.

    Нужна панели внизу окна: она показывает штамп активного листа в
    натуральном масштабе. Штамп — div с вложенными div (сетка, ячейки,
    поля), поэтому границы берутся подсчётом открытий и закрытий тегов, а не
    регулярным выражением.

    :return: HTML штампа или "" — если у листа надписи нет (обложка, титул).
    """
    start = fragment.find('<div class="stamp"')
    if start < 0:
        return ""
    depth = 0
    pos = start
    while True:
        open_at = fragment.find("<div", pos)
        close_at = fragment.find("</div>", pos)
        if close_at < 0:
            return ""
        if 0 <= open_at < close_at:
            depth += 1
            pos = open_at + 4
            continue
        depth -= 1
        pos = close_at + len("</div>")
        if depth == 0:
            return fragment[start:pos]

def stamp_size_mm(stamp_html: str) -> tuple:
    """Размер штампа в мм по его inline-стилю: (ширина, высота).

    Панель надписи по этому размеру задаёт свою высоту, чтобы штамп
    отображался целиком в масштабе 100%, без прокрутки внутри панели.
    """
    match = re.search(r'height:([\d.]+)mm', stamp_html or "", re.I)
    h = float(match.group(1)) if match else 40.0
    match = re.search(r'width:([\d.]+)mm', stamp_html or "", re.I)
    w = float(match.group(1)) if match else 185.0
    return w, h

def _wrap_sheet(content_html: str, stamp_html: str, w_mm: float | None = None,
                h_mm: float | None = None, data_pdf: tuple[str, int] | None = None,
                sheet_no: str = "", no_frame: bool = False,
                indent_top: bool = False, full_bleed: bool = False) -> str:
    """Оборачивает содержимое листа в div.sheet с рамкой и осн. надписью.

    Если заданы w_mm/h_mm — лист получает собственный размер (для листов
    произвольного формата: графика, приложения). Иначе — A4 по умолчанию.

    data_pdf — (путь_к_исходному_PDF, индекс_страницы). Когда указан, лист —
    графическая часть, которая при экспорте в PDF вставляется напрямую из
    исходного файла (вектор, родной размер), а не растеризуется через HTML.

    no_frame  — True: не рисовать рамку листа (для графической части). Вместо
                неё накладывается только квадрат сквозной нумерации, привязанный
                к верхнему правому углу листа (5 мм вниз и 5 мм влево).
    indent_top — True: сдвинуть текстовую область вниз, чтобы текст не заезжал
                под квадрат сквозной нумерации (учитывая его высоту от верха).
    full_bleed — True: содержимое занимает весь лист (без внутренних отступов
                и рамки). Используется для графической части, чтобы чертёж
                сохранял родной размер и его рамка совмещалась с рамкой листа,
                а квадрат сквозной нумерации — с верхним правым углом рамки.
    """
    size_style = ""
    data_attrs = ""
    if w_mm and h_mm:
        size_style = f' style="width:{w_mm:g}mm;height:{h_mm:g}mm"'
    if data_pdf:
        src_path, page = data_pdf
        data_attrs = (f' data-pdf-src="{esc(src_path)}" '
                      f'data-pdf-page="{int(page)}"')

    # Штамп позиционируется поверх листа, поэтому область контента должна
    # заканчиваться выше него. Раньше .sheet-content доходил до нижнего поля
    # листа и длинный текст (особенно из DOCX) рисовался под штампом при PDF-
    # экспорте. Высота берётся из inline-стиля конкретной формы (3/5/6).
    stamp_height = 0.0
    if stamp_html:
        match = re.search(r'\bheight:([\d.]+)mm', stamp_html, re.I)
        if match:
            stamp_height = float(match.group(1))
    content_parts = []
    if full_bleed:
        # Графический лист: содержимое занимает весь лист (без внутренних
        # отступов), чтобы чертёж сохранял родной размер и его рамка совпадала
        # с рамкой листа — тогда квадрат сквозной нумерации, привязанный к
        # верхнему правому углу, совмещается с рамкой чертежа.
        content_parts.append('left:0;top:0;right:0;bottom:0;padding:0')
    else:
        # Отступ сверху с учётом квадрата сквозной нумерации: текстовое
        # содержимое начинается ниже зоны квадрата (5..15 мм от верха листа),
        # чтобы нумерация не перекрывала текст.
        if indent_top:
            content_parts.append(f'top:{TEXT_TOP_INDENT_MM:g}mm')
        if stamp_height:
            reserved_bottom = PAGE_BOTTOM_MM + stamp_height + 2.0
            content_parts.append(f'bottom:{reserved_bottom:g}mm')
    content_style = f' style="{";".join(content_parts)}"' if content_parts else ""

    # Квадрат номера: линии рисуются полосами (ось на координате), а div
    # остаётся только под цифру — иначе его CSS-рамка расходилась с линиями
    # рамки листа и штампа. Линии кладутся ПОВЕРХ белого фона квадрата:
    # у графических листов фон закрывает чертёж, и полуштрих внутри квадрата
    # иначе был бы срезан. Координаты — по ФАКТИЧЕСКОМУ размеру листа: у
    # графической части он свой, и привязка к A4 уводила квадрат от угла.
    number_html = ""
    if sheet_no:
        number_html = (
            f'<div class="sheet-number" aria-label="Номер листа">{esc(sheet_no)}</div>'
            f'<div style="position:absolute;left:0;top:0;right:0;bottom:0;'
            f'z-index:4;pointer-events:none;">'
            f'{_sheet_number_rules(w_mm, h_mm)}</div>')
    frame_html = "" if no_frame else _frame_rules()
    return (f'<div class="sheet"{size_style}{data_attrs}>{frame_html}'
            f'{number_html}<div class="sheet-content"{content_style}>{content_html}</div>'
            f'{stamp_html}</div>')

def _is_docx(path: str) -> bool:
    return path.lower().endswith((".docx", ".doc"))

def _para_format(content, para_index: int):
    """Оформление абзаца Word по его индексу в document.paragraphs (или None)."""
    formats = getattr(content, "formats", None)
    if not formats:
        return None
    try:
        return formats.get(int(para_index))
    except Exception:
        return None

#: Кэш разбивки DOCX по листам. Ключ — (файл, время изменения, размер, форма
#: первого листа, правки); значение — (объект содержимого, страницы). За одно
#: переключение раздела разбивка запрашивается несколько раз (содержание,
#: штамп, оглавление, сами листы) по одному и тому же файлу.
_PAGINATE_CACHE: dict = {}
_PAGINATE_CACHE_LIMIT = 32

def _paginate_docx(content, path: str = "", first_form=2,
                   overrides: dict | None = None,
                   word_layout: dict | None = None) -> list:
    """Разбивает содержимое DOCX на «страницы» (списки строк/таблиц).

    Каждая страница — список элементов ("line", {...}) / ("table_rows", [row...]).
    Абзацы остаются в исходном порядке, но не расходуют бюджет листа: Word
    свободно переносит их между страницами. Высота таблиц оценивается по числу
    визуальных строк самой высокой ячейки; разрывы и перенос шапок таблиц
    планируются отдельно.

    Результат кэшируется: разбивка одного и того же файла запрашивается
    несколько раз за перерисовку, а стоит она дороже самого разбора.

    :param first_form: форма основной надписи ПЕРВОГО листа потока —
        2 (форма 5, штамп 40 мм) или "2a" (форма 6, штамп 15 мм). Бюджет
        первого листа берётся по ней; последующие листы всегда форма 6.
        Если лист оглавления уже занял «первый лист» раздела (форма 5),
        содержимое документа начинается с формы 6 — иначе ~25 мм низа
        первого листа остаются бы пустыми.

    Параметр first_form учитывается только в бюджете размещения таблиц; высота
    обычных абзацев при расчёте листов не используется.
    """
    from .docx_reader import docx_cache_key
    if word_layout:
        try:
            valid_layout = (
                int(word_layout.get("source_mtime_ns", -1)) == os.stat(path).st_mtime_ns
                and int(word_layout.get("source_size", -1)) == os.path.getsize(path))
        except OSError:
            valid_layout = False
        if valid_layout:
            return _paginate_docx_by_word(content, word_layout, overrides=overrides)
    key = None
    if not overrides:
        file_key = docx_cache_key(path)
        if file_key:
            tf = _text_fonts()
            key = (file_key, first_form, tf.size("body"), tf.size("heading"),
                   tf.size("table"))
            cached = _PAGINATE_CACHE.get(key)
            # Сверяем и сам объект содержимого: ключ мог достаться другому
            # разбору того же файла (например, после ручного read_docx).
            if cached is not None and cached[0] is content:
                return cached[1]
    pages = _paginate_docx_uncached(content, path, first_form=first_form,
                                    overrides=overrides)
    if key is not None:
        if len(_PAGINATE_CACHE) >= _PAGINATE_CACHE_LIMIT:
            _PAGINATE_CACHE.clear()
        _PAGINATE_CACHE[key] = (content, pages)
    return pages

def _split_text_at_word_boundaries(text: str, offsets: list[int]) -> list[str]:
    """Делит текст по границам листов, которые задал Word.

    Смещения приходят из Word: это концы фрагментов абзаца, разнесённых по
    листам. Word переносит строку по границе слова, поэтому смещение уже стоит
    между словами — там разрез и остаётся, символ в символ.

    Если смещение попало ВНУТРЬ слова (расхождение координат), разрез
    откатывается к началу этого слова: слово целиком уходит на следующий лист.
    Дотягивать разрез вперёд до конца слова нельзя — тогда на предыдущем листе
    появлялось слово, которое Word перенёс дальше, и лист расходился с Word.

    Если разрез попал сразу после маркера пункта («-», «12.» и т. п.), он
    откатывается ещё и к началу маркера: иначе внизу листа оставался одинокий
    знак, а текст пункта уезжал на следующую страницу.
    """
    parts = []
    start = 0
    for offset in offsets:
        target = min(max(start, int(offset)), len(text))
        if target <= start:
            parts.append("")
            continue
        split_at = target
        # Смещение внутри слова (пробела нет ни перед ним, ни после него):
        # откатываемся к началу слова.
        if (start < split_at < len(text)
                and not text[split_at].isspace()
                and not text[split_at - 1].isspace()):
            while split_at > start and not text[split_at - 1].isspace():
                split_at -= 1
        if split_at <= start:
            split_at = target
        # Маркер пункта не остаётся последним на листе. Абзац с мягкими
        # переносами (w:br) — это несколько пунктов в ОДНОМ абзаце Word, и
        # выравнивание разреза по границе слова могло оставить в конце листа
        # одинокий «-» от пункта, чей текст уехал на следующий лист.
        # Разрез откатывается к началу маркера: он уходит вместе со своим
        # пунктом, как это сделал бы Word.
        marker_tail = re.search(r"\n\s*([-–—•·▪◦*]|\d+[.)])(?:[ \t]*)$",
                                text[start:split_at])
        if marker_tail and start + marker_tail.start(1) > start:
            split_at = start + marker_tail.start(1)
        if (start == 0 and re.fullmatch(
                r"\s*(?:[-–—•·▪◦*]|\d+[.)])\s*", text[:split_at])):
            while split_at < len(text) and text[split_at].isspace():
                split_at += 1
            while split_at < len(text) and not text[split_at].isspace():
                split_at += 1
        parts.append(text[start:split_at])
        start = split_at
    parts.append(text[start:])
    return parts


def _keep_list_markers_with_text(pages: list) -> list:
    marker_pattern = re.compile(r"(?:[-–—•·▪◦*]|\d+[.)])")
    markers = []
    for page_index, page in enumerate(pages):
        for item_index, item in enumerate(page):
            if item[0] != "line":
                continue
            line = item[1]
            text = (line.get("text", "") or "").strip()
            fmt = line.get("fmt")
            if marker_pattern.fullmatch(text) or (
                    not text and fmt is not None and getattr(fmt, "has_list", False)):
                markers.append((page_index, item_index, item))

    for page_index, item_index, marker in reversed(markers):
        if page_index + 1 >= len(pages):
            continue
        next_item_index = next((index for index, item in enumerate(pages[page_index + 1])
                                if item[0] == "line" and item[1].get("text", "").strip()),
                               None)
        if next_item_index is None:
            continue
        pages[page_index].remove(marker)
        pages[page_index + 1].insert(next_item_index, marker)
    return pages

def _paginate_docx_by_word(content, layout: dict, overrides: dict | None = None) -> list:
    """Группирует HTML-блоки по номерам страниц, полученным от Word."""
    pages = [[] for _ in range(max(1, int(layout.get("page_count", 1) or 1)))]
    paragraphs = layout.get("paragraphs") or []
    tables = layout.get("tables") or []
    paragraph_index = -1
    table_index = -1
    blocks = getattr(content, "blocks", None) or []
    for kind, value in blocks:
        if kind == "paragraph":
            if getattr(value, "from_textbox", False):
                continue
            paragraph_index += 1
            text = (overrides or {}).get(paragraph_index, value.text)
            if paragraph_index >= len(paragraphs):
                continue
            page_info = paragraphs[paragraph_index] or {}
            # Пустой абзац, которого НЕТ в разметке Word, на листе не выводится:
            # ведущие пустые абзацы перед первым текстом из нормализованной
            # копии удаляются, и Word их не размечал. Пустые абзацы ВНУТРИ
            # текста разметку имеют — они занимают строку и должны быть видны,
            # иначе низ листа остаётся пустым.
            if not text.strip() and page_info.get("start") is None:
                continue
            fragments = page_info.get("fragments") or []
            # Заголовок — только абзац стиля Word (см. pz_toc). Он выводится
            # полужирным, курсив — как в исходнике; обычный текст записки —
            # единый, без полужирного и курсива.
            bold, italic = heading_emphasis(value)
            is_heading = bold
            line_data = {
                "text": text,
                "bold": bold,
                "italic": italic,
                "align": {"LEFT": "left", "CENTER": "center", "RIGHT": "right",
                          "JUSTIFY": "justify"}.get(value.alignment, "justify"),
                "size": _text_fonts().size("heading" if is_heading else "body"),
                "font": _text_fonts().font("heading" if is_heading else "body"),
                "para_index": paragraph_index,
                # Якорь гиперссылки оглавления — только у первого фрагмента
                # абзаца: разрезанный листом абзац не должен давать дубли id.
                "anchor": heading_anchor(paragraph_index),
                "fmt": _para_format(content, paragraph_index),
                "page_break": False,
            }
            if len(fragments) > 1 and text == value.text:
                fmt = line_data.get("fmt")
                list_prefix = (f"{fmt.marker_text} " if fmt is not None
                               and getattr(fmt, "has_list", False) else "")
                offsets = [int(item.get("end", 0)) - len(list_prefix)
                           for item in fragments[:-1]]
                parts = _split_text_at_word_boundaries(text, offsets)
                visible_parts = [(fragment, part)
                                 for fragment, part in zip(fragments, parts)
                                 if part.strip() or not text.strip()]
                for part_index, (fragment, part) in enumerate(visible_parts):
                    page_index = min(max(1, int(fragment.get("page", 1))), len(pages)) - 1
                    item_data = dict(line_data, text=part,
                                     suppress_marker=(bool(list_prefix)
                                                      and part_index > 0),
                                     anchor=(line_data["anchor"] if part_index == 0
                                             else ""))
                    pages[page_index].append(("line", item_data))
            else:
                page = page_info.get("start") or 1
                page_index = min(max(1, int(page)), len(pages)) - 1
                pages[page_index].append(("line", line_data))
        elif kind == "table":
            table_index += 1
            row_pages = (tables[table_index] if table_index < len(tables) else [])
            by_page = {}
            # Word повторяет при переносе ВСЕ строки-шапки (w:tblHeader), а не
            # только первую: в штатном расписании их 10, и повтор лишь одной
            # давал на листе-продолжении на 9 строк меньше, чем разметил Word.
            header_count = int(getattr(value, "header_rows", 0) or 0)
            if header_count <= 0 and len(value.rows) > 1:
                header_count = 1
            header = value.rows[:header_count] if header_count else None
            for row_index, row in enumerate(value.rows):
                row_info = row_pages[row_index] if row_index < len(row_pages) else {}
                page = row_info.get("start") or 1
                page_end = row_info.get("end") or page
                for row_page in range(int(page), int(page_end) + 1):
                    by_page.setdefault(row_page, []).append(row)
            previous_page = None
            for page, rows in sorted(by_page.items()):
                page_index = min(max(1, page), len(pages)) - 1
                if previous_page is not None and page != previous_page and header:
                    rows = list(header) + rows
                pages[page_index].append(("table_rows", rows))
                previous_page = page
    return _keep_list_markers_with_text(pages or [[]])

def _paginate_docx_uncached(content, path: str = "", first_form=2,
                            overrides: dict | None = None) -> list:
    stream = []
    # Шрифты и кегли текста ПЗ из настроек: используются и в разметке, и в
    # оценке высоты, поэтому разбивка совпадает с отрисовкой.
    tf = _text_fonts()
    # Критично обходить единый список blocks: paragraphs и tables в DOCX
    # находятся в одном body и могут чередоваться. Обход отдельных списков
    # сначала paragraphs, затем tables переносит все таблицы в конец ПЗ.
    blocks = getattr(content, "blocks", None)
    if blocks is None:
        # Совместимость с объектами DocxContent, созданными старой версией.
        blocks = [("paragraph", para) for para in content.paragraphs]
        blocks.extend(("table", tbl) for tbl in content.tables)

    # Сквозной индекс абзаца в content.paragraphs: порядок блоков совпадает с
    # порядком абзацев документа, поэтому индекс годится для записи правок
    # обратно в DOCX (document.paragraphs[i]).
    para_index = -1
    for kind, value in blocks:
        if kind == "paragraph":
            para = value
            # Абзацы из «надписей» (фигуры Word) не входят в document.paragraphs:
            # их нельзя адресовать индексом, иначе правка записалась бы в чужой
            # абзац файла. Они выводятся, но не редактируются (para_index = -1).
            if getattr(para, "from_textbox", False):
                editable_idx = -1
            else:
                para_index += 1
                editable_idx = para_index
            if not para.text.strip():
                continue
            is_heading = para.is_heading
            # Заголовок — только абзац стиля Word (см. pz_toc): выводится
            # полужирным, курсив — как в исходнике; обычный текст записки —
            # единый, без полужирного и курсива.
            bold, italic = heading_emphasis(para)
            align = {"LEFT": "left", "CENTER": "center", "RIGHT": "right",
                     "JUSTIFY": "justify"}.get(para.alignment, "justify")
            # Кегль — из настроек текста ПЗ (config.TextFonts): обычный текст и
            # заголовки. Раньше здесь были зашиты 11/13 pt, из-за чего кегль
            # предпросмотра не совпадал с Word (14 pt) и число листов расходилось.
            size = tf.size("heading") if is_heading else tf.size("body")
            fmt = _para_format(content, para_index) if editable_idx >= 0 else None
            text = para.text
            if overrides and editable_idx >= 0:
                text = overrides.get(editable_idx, text)
            stream.append(("line", {"text": text, "bold": bold, "italic": italic,
                                    "align": align, "size": size,
                                    "font": tf.font("heading" if is_heading else "body"),
                                    "para_index": editable_idx, "fmt": fmt,
                                    "anchor": heading_anchor(editable_idx),
                                    "page_break": bool(fmt is not None
                                                       and fmt.page_break_before)}))
        elif kind == "table":
            stream.append(("table", value))

    # Доступная высота текстовой области: лист A4 - верхнее поле - нижнее поле,
    # уменьшенные на высоту штампа (форма 5 — 40 мм, форма 6 — 15 мм) и зазор.
    # Запас 2 мм сверху/снизу (border/отступы), чтобы оценка не занижала объём.
    first_budget = _page_text_height(FORM_HEIGHT[first_form])  # форма 5 или 6
    next_budget = _page_text_height(FORM_HEIGHT["2a"])         # форма 6
    pages, cur, cur_h, budget = [], [], 0.0, first_budget

    def flush_page() -> None:
        nonlocal cur, cur_h, budget
        if cur:
            pages.append(cur)
        cur = []
        cur_h = 0.0
        budget = next_budget

    for idx, item in enumerate(stream):
        if item[0] == "line":
            line = item[1]
            h = 0.0
            # Разрыв страницы из Word закрывает лист даже тогда, когда место
            # на нём ещё есть (кроме случая, когда абзац и так начинает лист).
            if line.get("page_break") and cur:
                flush_page()
            cur.append(item)
            cur_h += h
            continue

        # Таблица разбивается по строкам: целая таблица не влезает на лист
        # даже при переполнении — иначе нижняя часть обрезается CSS-ом и
        # пользователь видит половину строки/таблицы. При переносе таблицы на
        # следующий лист повторяются ВСЕ строки-шапки (w:tblHeader) — так же
        # ведёт себя Word; в штатном расписании их 10.
        table = item[1]
        rows = list(table.rows)
        ncols = _table_ncols(table)
        # Фактические ширины колонок из исходного DOCX (w:tblGrid): по ним
        # считаются и перенос текста в ячейке, и CSS-grid при отрисовке —
        # предпросмотр совпадает с исходной таблицей.
        grid = table_col_widths_mm(path, getattr(table, "table_index", -1), ncols)
        # Кегль таблиц — из настроек текста ПЗ: от него зависят и перенос в
        # ячейке, и высота строки, поэтому разбивка совпадает с отрисовкой.
        table_pt = tf.size("table")
        chars = _table_cell_chars_per_line(ncols, size=table_pt)
        cell_chars = [_table_cell_chars_per_line(ncols, i, w, table_pt)
                      for i, w in enumerate(grid)]
        header_count = int(getattr(table, "header_rows", 0) or 0)
        if header_count <= 0 and len(rows) > 1:
            header_count = 1
        header = rows[:header_count] if header_count else []
        header_cost = sum(_estimate_table_row_h(r, chars, cell_chars,
                                               col_widths=grid, size=table_pt,
                                               font=tf.font('table'))
                          for r in header)
        chunk = []
        chunk_h = 0.0
        for row in rows:
            cost = _estimate_table_row_h(row, chars, cell_chars,
                                         col_widths=grid, size=table_pt,
                                         font=tf.font('table'))
            if chunk and cur_h + chunk_h + cost > budget:
                cur.append(("table_rows", chunk))
                cur_h += chunk_h
                flush_page()
                # Продолжение таблицы начинается с повторённых строк-шапок.
                chunk = list(header)
                chunk_h = header_cost
            elif not chunk and cur and cur_h + cost > budget:
                flush_page()
                if header:
                    chunk = list(header)
            chunk.append(row)
            chunk_h += cost
        if chunk:
            cur.append(("table_rows", chunk))
            cur_h += chunk_h
    if cur:
        pages.append(cur)
    if not pages:
        pages = [[]]
    return _keep_list_markers_with_text(pages)

def _grid_row_render(row, *,
                     editable: bool,
                     para_file: str,
                     overrides: dict | None = None,
                     cell_overrides: dict | None = None) -> str:
    """HTML одной строки таблицы DOCX. Ячейки адресуемы и редактируемы.

    Адрес ячейки — data-table / data-row / data-cell (индексы в document.tables),
    он же используется при записи правок в файл (apply_table_cell_edits).
    """
    overrides = overrides or {}
    cell_overrides = cell_overrides or {}
    cells_html = []
    for c, cell in enumerate(row, start=1):
        text = getattr(cell, "text", None) or ""
        # Позиция задаётся явно (колонка), чтобы продолжения объединённых ячеек
        # не сдвигали соседние ячейки (CSS auto-placement).
        continuation = getattr(cell, "continuation", "")
        if continuation:
            # Ячейка внутри объединения: место занимает, но текст не повторяет —
            # его выводит начало объединения.
            # Вертикальное объединение не рисуется как grid-row:span: при
            # переносе таблицы на следующий лист такое объединение «уезжает»
            # вместе с листом, и продолжение осталось бы без текста. Пустая
            # ячейка сохраняет колонку (и её ширину), а текст показан один раз —
            # в начале объединения.
            span = int(getattr(cell, "col_span", 1) or 1)
            if continuation == "horizontal":
                # Колонку целиком перекрывает начало объединения (у него
                # grid-column / span) — отдельную ячейку рисовать нельзя,
                # иначе она наложится на объединение и строка «поедет».
                continue
            # Продолжение вертикального объединения: рамку рисуем, иначе в
            # строке появляется «дыра» и таблица выглядит разорванной.
            style = (f"grid-column:{c} / span {span};" if span > 1
                     else f"grid-column:{c};")
            cells_html.append(f'<div class="dt-cell" style="{style}"></div>')
            continue
        address = getattr(cell, "address", "")
        if editable and address and address in cell_overrides:
            text = cell_overrides[address]
        style = f"grid-column:{c};"
        colspan_cell = int(getattr(cell, "col_span", 1) or 1)
        if colspan_cell > 1:
            style = f"grid-column:{c} / span {colspan_cell};"
        if getattr(cell, "bold", False):
            style += "font-weight:bold;"
        alignment = getattr(cell, "alignment", "")
        if alignment:
            style += f"text-align:{alignment.lower()};"
        if editable and address:
            table_index, row_index, cell_index = address.split(":")
            cells_html.append(
                f'<div class="dt-cell pz-cell" contenteditable="true" '
                f'data-table="{table_index}" data-row="{row_index}" '
                f'data-cell="{cell_index}" data-cell-addr="{esc(address)}" '
                f'data-para-file="{esc(para_file)}" style="{style}">{esc(text)}</div>')
        else:
            cells_html.append(f'<div class="dt-cell" style="{style}">{esc(text)}</div>')
    return f'<div class="dt-row">{"".join(cells_html)}</div>'

def _pz_text(text) -> str:
    """Текст абзаца ПЗ для HTML: экранирование + переносы строк как <br>.

    python-docx отдаёт мягкий перенос Word (w:br) символом «\n»; в HTML он бы
    свернулся в пробел, поэтому переносы превращаем в <br> — иначе пропадало бы
    разбиение строки, сделанное в Word.
    """
    return esc(text).replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")

def _line_html(line: dict, text: str, *, editable: bool, para_file: str,
               idx: int) -> str:
    """HTML одного абзаца DOCX: обычный абзац со стилями Word или пункт списка.

    Отступы, интервалы и выравнивание берутся из оформления абзаца в Word
    (app.core.pz_format), кегль — по карте ГОСТ (строка "size").

    Пункт списка выводится как «маркер + абзац» в обёртке .pz-item: маркер
    (номер или знак Word) занимает поле «висячего» отступа абзаца, а текст
    пункта начинается на отступе left — ровно так же, как расставляет список
    Word. Поэтому выключка пункта (в т.ч. по ширине) сохраняется.
    """
    fmt = line.get("fmt")
    size = line.get("size", 11)
    # Шрифт абзаца — из настроек: заголовки и обычный текст могут различаться.
    family = line.get("font") or ""
    font_css = f"font-family:{_css_font_family(family)};" if family else ""
    align = line.get("align", "justify")
    is_list = bool(fmt is not None and fmt.has_list)
    style = f"font-size:{size}pt;{font_css}"
    if fmt is not None:
        style = paragraph_css(fmt, extra=f"font-size:{size}pt;{font_css}",
                              omit_text_indent=is_list,
                              line_mm=pz_line_mm(size, fmt))
        if para_align_is_default(fmt, align):
            style += f"text-align:{align};"
    else:
        style += (f"text-align:{align};margin:0;"
                  f"line-height:{pz_line_mm(size):g}mm;")
    if line.get("bold"):
        style += "font-weight:bold;"
    if line.get("italic"):
        style += "font-style:italic;"

    marker = ""
    if is_list and not line.get("suppress_marker"):
        # Маркер выносится в отдельный бокс рядом с абзацем: он не входит в
        # редактируемый текст (пользователь правит только текст пункта, а
        # номер/маркер остаётся как в Word) и не удаляется клавишей Backspace.
        marker = (f'<span class="pz-mark" contenteditable="false" '
                  f'style="{marker_css(fmt)}">{esc(fmt.marker_text)}</span>')

    body = _pz_text(text)
    if not body:
        # Пустой абзац Word занимает на листе целую строку, а в HTML <p> без
        # содержимого схлопывается в ноль: без min-height строки пропадали бы
        # и низ листа оставался пустым.
        style += f"min-height:{pz_line_mm(size, fmt):g}mm;"
    # idx < 0 — абзац из «надписи» Word: его нет в document.paragraphs, поэтому
    # он выводится, но не редактируется (правка записалась бы в чужой абзац).
    # id — цель гиперссылки строки оглавления (#MIMO_H{индекс абзаца}).
    anchor = (line.get("anchor") or "").strip()
    id_attr = f' id="{esc(anchor)}"' if anchor else ""
    if editable and idx >= 0:
        para = (f'<p class="pz-para" data-para-index="{idx}" '
                f'data-para-file="{esc(para_file)}" '
                f'contenteditable="true"{id_attr} style="{style}">{body}</p>')
    else:
        # В собранном документе (HTML/PDF) редактируемая разметка не выводится.
        para = f'<p{id_attr} style="{style}">{body}</p>'
    if is_list:
        return (f'<div class="pz-item" data-list-marker="{esc(fmt.marker_text)}">'
                f'{marker}{para}</div>')
    if re.fullmatch(r"\s*(?:[-–—•·▪◦*]|\d+[.)])\s*", text):
        return (f'<p class="pz-list-marker pz-para" data-para-index="{idx}" '
                f'data-para-file="{esc(para_file)}" contenteditable="true"'
                f'{id_attr} style="{style}">{body}</p>')
    return para

def para_align_is_default(fmt, align: str) -> bool:
    """Нужно ли добавить ГОСТ-выключку поверх оформления Word.

    Word не задаёт выравнивание, если оно наследуется от стиля; в этом случае
    используется выключка по ГОСТ (по ширине). Явно заданное в Word
    выравнивание НЕ переопределяется — как в исходном документе.
    """
    if fmt is None:
        return True
    return not getattr(fmt, "alignment", "")

def _render_docx_page(page_items: list, editable: bool = False,
                      para_file: str = "", overrides: dict | None = None,
                      cell_overrides: dict | None = None) -> list:
    """Возвращает HTML-строки для одной страницы DOCX.

    editable=True — абзацы и ячейки таблиц становятся редактируемыми прямо в
    предпросмотре (contenteditable с data-para-index / data-cell-addr). Правка
    уходит в Python через мост и записывается в исходный DOCX
    («Сохранить правки ПЗ»).

    overrides      — несохранённые правки абзацев ({индекс абзаца: текст});
    cell_overrides — несохранённые правки ячеек ({"т:с:я": текст}).
    Оба применяются поверх текста файла, чтобы правки не пропадали при
    перерисовке листов.

    """
    overrides = overrides or {}
    cell_overrides = cell_overrides or {}
    parts = []
    for item in page_items:
        if item[0] == "line":
            line = item[1]
            idx = int(line.get("para_index", -1))
            text = line["text"]
            if editable and idx in overrides:
                text = overrides[idx]
            parts.append(_line_html(line, text, editable=editable,
                                    para_file=para_file, idx=idx))
        else:
            # item[1] — список строк (row), каждая row — список ячеек (TableCell).
            # Таблица рисуется CSS-grid с фактическими ширинами колонок из DOCX,
            # а не как <table> (иначе объединённые ячейки отображались бы
            # дублями: python-docx отдаёт одну и ту же ячейку на весь span).
            chunk = item[1]
            table_index = -1
            if chunk:
                table_index = int(getattr(chunk[0][0], "table_index", -1)) if chunk[0] else -1
            ncols = max((len(r) for r in chunk), default=1)
            colspec = table_colspec(para_file, table_index, ncols)
            rows_html = []
            for row in chunk:
                rows_html.append(_grid_row_render(
                    row, editable=editable, para_file=para_file, overrides=overrides,
                    cell_overrides=cell_overrides))
            parts.append(f'<div class="dt" style="grid-template-columns:{colspec}">'
                         f'{"".join(rows_html)}</div>')
    return parts

#: Кэш растрированных страниц PDF: (путь, mtime, размер, индекс, dpi) -> base64.
#: Графические листы вставляются в предпросмотр картинкой, и без кэша каждая
#: перерисовка заново растрировала ВСЕ страницы: на томе из 169 листов это
#: 12 секунд и десятки мегабайт HTML.
_PDF_IMAGE_CACHE: dict = {}
_PDF_IMAGE_CACHE_LIMIT = 64
#: Кэш размеров страниц PDF: (путь, индекс) -> (ширина, высота) в мм.
_PDF_SIZE_CACHE: dict = {}

def _pdf_page_image_b64(path: str, index: int, max_side: int = 1200) -> str:
    """Рендерит страницу PDF в PNG и возвращает base64 (для встраивания в HTML).

    Результат кэшируется по (файл, его состояние, номер страницы, разрешение):
    за одну перерисовку листа картинка запрашивается один раз, но при возврате
    на раздел она нужна снова, и повторная растеризация была самой дорогой
    частью переключения.

    :param max_side: длинная сторона картинки в пикселях. 1200 достаточно для
        экранного превью (лист показывается масштабированным под окно); в
        собранный PDF графика попадает вектором из исходного файла, поэтому
        качество предпросмотра на результат не влияет.
    """
    from .docx_reader import docx_cache_key
    file_key = docx_cache_key(path)
    if file_key is None:
        return ""
    key = (file_key, int(index), int(max_side))
    cached = _PDF_IMAGE_CACHE.get(key)
    if cached is not None:
        return cached
    import pymupdf
    with pymupdf.open(path) as doc:
        page = doc[index]
        rect = page.rect
        mm_w = rect.width * 25.4 / 72.0
        mm_h = rect.height * 25.4 / 72.0
        dpi = max_side / (max(mm_w, mm_h) / 25.4)
        dpi = int(max(50.0, min(150.0, dpi)))
        pix = page.get_pixmap(dpi=dpi)
        data = base64.b64encode(pix.tobytes("png")).decode("ascii")
    if len(_PDF_IMAGE_CACHE) >= _PDF_IMAGE_CACHE_LIMIT:
        _PDF_IMAGE_CACHE.clear()
    _PDF_IMAGE_CACHE[key] = data
    return data

def _pdf_page_size_mm(path: str, index: int) -> tuple:
    """Размеры страницы PDF в мм (с кэшем: вызывается на каждую страницу)."""
    from .docx_reader import docx_cache_key
    file_key = docx_cache_key(path)
    key = (file_key or path, int(index))
    cached = _PDF_SIZE_CACHE.get(key)
    if cached is not None:
        return cached
    import pymupdf
    with pymupdf.open(path) as doc:
        rect = doc[index].rect
        size = (round(rect.width * 25.4 / 72.0, 1), round(rect.height * 25.4 / 72.0, 1))
    _PDF_SIZE_CACHE[key] = size
    return size

def _pz_toc_fragments(snapshot: dict, font_map: FontMap,
                      counter: _SheetCounter | None = None,
                      sheets_total: int | None = None) -> list:
    """Листы автоматического оглавления пояснительной записки.

    Оглавление построено единым источником (app.core.pz_toc) по ПЕРВОМУ
    документу раздела ПЗ и передано в снимке предпросмотра. Оно вставляется
    первым листом записки — до её листов; разбивка строк по листам выполняется
    тем же способом, что и пагинация записки, поэтому число листов оглавления
    в предпросмотре, Word и PDF совпадает.

    Пустой список — оглавление не выводится (нет документа ПЗ или оглавление
    отключено в конфигурации).
    """
    if not snapshot.get("pz_toc_items"):
        return []
    # Оглавление принадлежит записке целиком и показывается только перед её
    # первым документом (документ, по которому оно построено). Для остальных
    # документов раздела лист оглавления не повторяется.
    doc_path = snapshot.get("path", "") or ""
    toc_path = snapshot.get("pz_toc_path", "") or ""
    if toc_path and doc_path and os.path.abspath(doc_path) != os.path.abspath(toc_path):
        return []
    from .pz_toc import items_from_dicts, paginate_items
    pages = paginate_items(items_from_dicts(snapshot.get("pz_toc_items")))
    frags = []
    for i in range(len(pages)):
        if counter is not None:
            sheet_no = counter.next()
        else:
            sheet_no = i + 1
        num = _numbered_snapshot(snapshot, sheet_no, sheets_total)
        # Номера листов в строках оглавления — как внизу листа (графа «Лист»
        # штампа): нумерация ведётся ВНУТРИ раздела ПЗ, листы оглавления — его
        # первые листы. Поэтому сдвиг — только число листов самого оглавления;
        # сквозная нумерация тома (square_no) сюда не входит.
        num["pz_toc_offset"] = len(pages)
        num["pz_toc_pages"] = len(pages)
        # Первый лист записки несёт большой штамп (форма 5), остальные — форму 6.
        num["toc_first"] = (i == 0)
        if i > 0:
            num["pz_toc_page"] = i + 1
        frag = _sheet_fragment("pz_toc", num, font_map, counter=None,
                               sheets_total=sheets_total)
        frags.append(frag)
    return frags

def _empty_docx_note(path: str) -> str:
    """Пояснение к пустому DOCX: почему в документе не нашлось текста."""
    import re as _re
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            xml = z.read("word/document.xml").decode("utf-8", "ignore") \
                if "word/document.xml" in names else ""
    except Exception:
        return "Файл прочитан, но текста в нём не найдено."
    # Непустой текстовый тег — есть ли вообще текст в основном документе.
    has_text_tags = bool(_re.search(r"<w:t[\s/>]", xml))
    hints = []
    if "w:txbxContent" in xml:
        hints.append("текст находится в надписях/фигурах Word")
    if "w:drawing" in xml or "w:pict" in xml:
        hints.append("содержимое — рисунки или фигуры")
    if not has_text_tags:
        if any(n.startswith("word/header") or n.startswith("word/footer")
               for n in names):
            hints.append("текст находится в колонтитулах")
        if "word/footnotes.xml" in names:
            hints.append("текст находится в сносках")
    if not hints:
        hints.append("в основном тексте нет абзацев и таблиц с текстом")
    return ("Файл прочитан, но текста для вывода не найдено: "
            + "; ".join(hints) + ".")

def _placeholder_sheet_fragment(kind: str, snapshot: dict,
                                font_map: FontMap | None = None,
                                counter: "_SheetCounter | None" = None,
                                sheets_total: int | None = None,
                                note: str = "") -> str:
    """Лист-заглушка вместо содержимого файла.

    Не обращается к исходному файлу вовсе. Это важно: если файл существует, но
    не читается (битый DOCX, .doc, повреждённый PDF), возврат в
    `_sheet_fragment("doc")` снова уходил бы в загрузчик и снова падал —
    получалась бесконечная рекурсия `_sheet_fragment → _docx_sheet_fragments →
    _sheet_fragment` и «maximum recursion depth exceeded» при сборке тома.

    :param note: текст-пояснение под заголовком (например, почему документ пуст).
    """
    font_map = font_map or FontMap()
    if counter is not None:
        sheet_no = counter.next()
    else:
        sheet_no = snapshot.get("sheet_no", "") or ""
    num = _numbered_snapshot(snapshot, sheet_no, sheets_total)
    if kind == "appendix":
        label = num.get("label", "Приложение")
        content = render_placeholder(label, subtitle="Приложение",
                                     path=num.get("path", ""), kind="appendix")
        stamp = ""
    elif kind == "section":
        content = render_placeholder(num.get("label", "Раздел"),
                                     subtitle="Раздел тома", kind="section",
                                     note=note)
        stamp = (_stamp_html(2, num, font_map)
                 if num.get("section") == "Пояснительная записка" else "")
    else:
        content = render_placeholder(num.get("label", "Документ"),
                                     subtitle=num.get("section_label", ""),
                                     path=num.get("path", ""), kind="doc",
                                     note=note)
        stamp = (_stamp_html(2, num, font_map)
                 if num.get("section") == "Пояснительная записка" else "")
    has_no = bool(num.get("sheet_no"))
    return _wrap_sheet(content, stamp,
                       sheet_no=num.get("square_no", num.get("sheet_no", "")),
                       indent_top=has_no)

def _docx_sheet_fragments(path: str, snapshot: dict, font_map: FontMap,
                          first: bool = True, max_pages: int | None = None,
                          counter: _SheetCounter | None = None,
                          sheets_total: int | None = None) -> list:
    """Страницы DOCX (ПЗ) как отдельные листы.

    Для ПЕРВОГО документа раздела ПЗ перед его листами добавляются листы
    автоматического оглавления (app.pz_toc); они несут большой штамп (форма 5),
    а листы самой записки — последующие (форма 6). Для остальных документов
    раздела форма листа задаётся аргументом first (первый лист документа —
    форма 5), поэтому нумерация не зависит от листов оглавления.
    """
    toc_frags = []
    if first:
        # Оглавление занимает номер листа (сквозная нумерация тома), но не
        # считается «первым листом» записки: его листы несут большой штамп
        # (форма 5), а первый лист самой записки — уже последующий (форма 6).
        toc_frags = _pz_toc_fragments(snapshot, font_map, counter=counter,
                                      sheets_total=sheets_total)
        first = not toc_frags         # без оглавления форма 5 — у первого листа ПЗ
    from .docx_reader import read_docx
    try:
        content = read_docx(path)
    except Exception as e:
        # Файл не читается вообще — выводим заглушку с причиной, чтобы сбой
        # был понятен по самому листу.
        return toc_frags + [_placeholder_sheet_fragment(
            "doc", snapshot, font_map, counter=counter, sheets_total=sheets_total,
            note=f"Не удалось прочитать файл: {type(e).__name__}: {e}")]
    # Документ прочитан, но показывать нечего: пустой лист без пояснения
    # выглядел как «записка не прикрепилась». Выводим заглушку с причиной.
    if getattr(content, "is_empty", False):
        return toc_frags + [_placeholder_sheet_fragment(
            "doc", snapshot, font_map, counter=counter, sheets_total=sheets_total,
            note=_empty_docx_note(path))]
    layouts = snapshot.get("word_pagination") or {}
    word_layout = layouts.get(path)
    if word_layout is None:
        from .word_pagination import cached_docx_pages
        word_layout = cached_docx_pages(path, first_form=(2 if first else "2a"))
    pages = _paginate_docx(content, path,
                           # Форма первого листа та же, что и у штампа ниже:
                           # если лист оглавления уже занял форму 5, первая
                           # страница содержимого несёт форму 6 (и бюджет 246 мм).
                           first_form=(2 if first else "2a"),
                           # Несохранённые правки абзацев: дописанный текст
                           # должен увеличивать высоту, иначе он обрежется.
                           overrides=(snapshot.get("pz_edits") or {}).get(path),
                           word_layout=word_layout)
    if max_pages is not None:
        pages = pages[:max_pages]
    # Несохранённые правки ячеек таблиц этого документа ({"т:с:я": текст}).
    cell_overrides = (snapshot.get("pz_cell_edits") or {}).get(path) or {}
    frags = []
    for i, page_items in enumerate(pages):
        form = 2 if (first and i == 0) else "2a"
        # Номер листа в пределах раздела: берём из счётчика (или 1-й лист).
        if counter is not None:
            sheet_no = counter.next()
        else:
            sheet_no = i + 1
        num = _numbered_snapshot(snapshot, sheet_no, sheets_total)
        stamp = _stamp_html(form, num, font_map)
        # Текст ПЗ редактируется прямо в предпросмотре (когда открыт документ ПЗ).
        body = "".join(_render_docx_page(
            page_items, editable=bool(num.get("pz_editable")),
            para_file=path,
            overrides=(num.get("pz_edits") or {}).get(path),
            cell_overrides=cell_overrides))
        # Отступ сверху с учётом квадрата сквозной нумерации (indent_top),
        # чтобы текст не заезжал под квадрат в правом верхнем углу.
        # Границы листов задаёт Word, и предпросмотр их не пересчитывает:
        # блоки раскладываются строго по номерам страниц из разметки.
        frags.append(_wrap_sheet(body, stamp,
                                 sheet_no=num.get("square_no", num.get("sheet_no", "")),
                                 indent_top=True))
    if not frags:
        frags = [_placeholder_sheet_fragment("doc", snapshot, font_map,
                                            counter=counter,
                                            sheets_total=sheets_total)]
    # После оглавления форма первого листа записки — 6 (малый штамп): оглавление
    # уже было «первым листом» записки с формой 5.
    if toc_frags:
        return toc_frags + frags
    return frags

def _pdf_sheet_fragments(path: str, snapshot: dict, font_map: FontMap,
                         first: bool = True, is_appendix: bool = False,
                         max_pages: int | None = None,
                         counter: _SheetCounter | None = None,
                         sheets_total: int | None = None) -> list:
    """Страницы PDF как отдельные листы-картинки (графика / приложение / ПЗ-PDF)."""
    import pymupdf
    fallback = ("appendix" if is_appendix else "doc")
    try:
        with pymupdf.open(path) as pdf:
            count = pdf.page_count
    except Exception:
        return [_placeholder_sheet_fragment(fallback, snapshot, font_map,
                                           counter=counter,
                                           sheets_total=sheets_total)]
    if count <= 0:
        return [_placeholder_sheet_fragment(fallback, snapshot, font_map,
                                           counter=counter,
                                           sheets_total=sheets_total)]
    if max_pages is not None:
        count = min(count, max_pages)

    frags = []
    # Правило формата: только ГРАФИЧЕСКАЯ ЧАСТЬ сохраняет родной размер листа
    # из вложенного PDF. Все остальные листы (приложения, ПЗ-PDF) — A4.
    section = snapshot.get("section", "")
    is_graphic = (not is_appendix and section != "Пояснительная записка")
    # Предпросмотр графической части: содержимое листов не рендерится вовсе —
    # показываются пустые листы с квадратом нумерации. Растр чертежей в
    # предпросмотре не нужен (он тяжёлый и всё равно не используется при
    # сборке: в PDF графика попадает вектором из исходного файла), а
    # нумерация нужна, чтобы листать и сверять номера с содержанием.
    # В собранном HTML-томе флаг не ставится — там листы встраиваются.
    blank_graphic = bool(snapshot.get("preview_blank_graphic"))
    for i in range(count):
        # Номер листа в пределах раздела: из счётчика (или i+1).
        if counter is not None:
            sheet_no = counter.next()
        else:
            sheet_no = i + 1
        num = _numbered_snapshot(snapshot, sheet_no, sheets_total)
        w_mm, h_mm = _pdf_page_size_mm(path, i)
        if is_appendix:
            header = (f'<div style="font-size:13pt;font-weight:bold;text-align:center;'
                      f'margin-bottom:4mm">{esc(num.get("label", "Приложение"))}</div>')
            stamp = ""
        else:
            header = ""
            if section == "Пояснительная записка":
                stamp = _stamp_html(2 if (first and i == 0) else "2a", num, font_map)
            else:
                stamp = ""
        # Картинка страницы. В предпросмотре графической части её нет совсем
        # (blank_graphic): лист остаётся пустым, но получает квадрат номера.
        # В собранном HTML-томе и для приложений/ПЗ-PDF страница встраивается
        # сразу готовым data-URL.
        if is_graphic and blank_graphic:
            img = ""
        else:
            b64 = _pdf_page_image_b64(path, i)
            img = (f'<img src="data:image/png;base64,{b64}" '
                   f'style="width:100%;height:100%;object-fit:contain;"/>')
        if is_graphic:
            # Графическая часть: НЕ рисуем новую рамку и штамп — сохраняем
            # ссылку на исходную страницу PDF. При экспорте она вставляется
            # напрямую из PDF (вектор, родной размер).
            # В предпросмотре лист остаётся пустым (blank_graphic): виден
            # только квадрат номера, чтобы листать и сверять нумерацию.
            # Пояснительная надпись в центре говорит, что лист не сломан, а
            # намеренно не рендерится.
            if blank_graphic:
                note = ('<div style="position:absolute;left:0;top:0;right:0;bottom:0;'
                        'display:flex;align-items:center;justify-content:center;'
                        'font-size:10pt;color:#999;">'
                        f'Лист графической части {esc(num.get("sheet_no", ""))}'
                        '</div>')
                header += note
            frags.append(_wrap_sheet(header + img, stamp, w_mm=w_mm, h_mm=h_mm,
                                     data_pdf=(path, i),
                                     sheet_no=num.get("square_no", num.get("sheet_no", "")),
                                     no_frame=True, full_bleed=True))
        else:
            # Текстовые листы (ПЗ) и приложения: отступ сверху с учётом
            # квадрата сквозной нумерации, чтобы текст не заезжал под него.
            frags.append(_wrap_sheet(header + img, stamp,
                                     sheet_no=num.get("square_no", num.get("sheet_no", "")),
                                     indent_top=True))
    return frags

def _doc_sheet_fragments(d, snapshot: dict, font_map: FontMap,
                         first: bool = True, max_pages: int | None = None,
                         counter: _SheetCounter | None = None,
                         sheets_total: int | None = None) -> list:
    """Листы документа (d — SourceFile): DOCX -> текст, PDF -> страницы-картинки."""
    path = d.path
    if not path or not os.path.exists(path):
        return [_placeholder_sheet_fragment("doc", snapshot, font_map,
                                           counter=counter,
                                           sheets_total=sheets_total)]
    if _is_docx(path):
        return _docx_sheet_fragments(path, snapshot, font_map, first=first,
                                     max_pages=max_pages, counter=counter,
                                     sheets_total=sheets_total)
    return _pdf_sheet_fragments(path, snapshot, font_map, first=first,
                                is_appendix=False, max_pages=max_pages,
                                counter=counter, sheets_total=sheets_total)

def _appendix_sheet_fragments(a, snapshot: dict, font_map: FontMap,
                              max_pages: int | None = None,
                              counter: _SheetCounter | None = None,
                              sheets_total: int | None = None) -> list:
    """Листы приложения (a — Appendix): PDF как страницы-картинки с заголовком."""
    path = a.file_path
    if not path or not os.path.exists(path):
        return [_placeholder_sheet_fragment("appendix", snapshot, font_map,
                                           counter=counter,
                                           sheets_total=sheets_total)]
    return _pdf_sheet_fragments(path, snapshot, font_map, first=True,
                                is_appendix=True, max_pages=max_pages,
                                counter=counter, sheets_total=sheets_total)

# ---- Сквозная нумерация листов внутри раздела --------------------------------
class _SheetCounter:
    """Мутируемый счётчик номеров листов одного раздела (начинается с 1)."""

    __slots__ = ("value",)

    def __init__(self, start: int = 1):
        self.value = start

    def next(self) -> int:
        v = self.value
        self.value += 1
        return v

def _numbered_snapshot(snapshot: dict, sheet_no, sheets_total) -> dict:
    """Копия snapshot с номером листа и общим числом листов раздела.

    sheet_no / sheets_total — значения для граф «Лист» / «Листов» основной
    надписи (формы 5/6): листы ТОЛЬКО этого раздела (части тома).

    square_no — номер в квадрате сквозной нумерации (правый верхний угол).
    В сборке тома в snapshot передаётся square_base (первый номер раздела
    по toc.build_layout), поэтому квадрат остаётся сквозным по тому и
    совпадает с содержанием; без base квадрат равен локальному номеру.
    """
    s = dict(snapshot)
    s["sheet_no"] = str(sheet_no)
    if sheets_total is not None:
        s["sheets_total"] = str(sheets_total)
    base = snapshot.get("square_base")
    try:
        s["square_no"] = str(int(base) + int(sheet_no) - 1)
    except (TypeError, ValueError):
        s["square_no"] = str(sheet_no) if sheet_no not in (None, "") else ""
    return s

def _pdf_count(path: str) -> int:
    """Число страниц PDF (для подсчёта листов раздела)."""
    import pymupdf
    try:
        with pymupdf.open(path) as pdf:
            return max(1, pdf.page_count)
    except Exception:
        return 1

def _doc_sheet_count(d, first_form=2) -> int:
    """Число листов одного документа (DOCX — пагинация, PDF — страницы).

    :param first_form: форма штампа первого листа потока документа (2 или "2a")
        — по ней считается бюджет первого листа (см. _paginate_docx).
    """
    path = d.path
    if not path or not os.path.exists(path):
        return 1          # будет показан лист-заглушка
    if _is_docx(path):
        try:
            from .docx_reader import read_docx
            # Путь передаётся обязательно: от него зависят фактические ширины
            # колонок таблиц, а значит и разбивка на листы. Иначе число листов
            # в содержании/штампе не совпало бы с предпросмотром.
            from .word_pagination import cached_docx_pages
            layout = cached_docx_pages(path, first_form)
            if layout is not None:
                return max(1, int(layout.get("page_count", 1) or 1))
            return max(1, len(_paginate_docx(read_docx(path), path,
                                             first_form=first_form)))
        except Exception:
            return max(1, int(getattr(d, "pages", 0) or 1))
    return _pdf_count(path)

def pz_doc_first_forms(project) -> list:
    """Форма первого листа содержимого каждого документа ПЗ в разделе.

    Возвращает список [2 | "2a"] параллельно documents_for(PZ):
    первый лист раздела несёт форму 5, только если перед ним НЕТ листов
    оглавления; все остальные листы (листы оглавления уже заняли форму 5,
    последующие документы) — форма 6. Единое правило для предпросмотра,
    подсчёта листов и редактора ПЗ.
    """
    from ..models import DocKind
    try:
        docs = project.documents_for(DocKind.POYASNITELNAYA)
    except Exception:
        return []
    if not docs:
        return []
    try:
        from .pz_toc import project_toc_sheets
        toc_sheets = project_toc_sheets(project)
    except Exception:
        toc_sheets = 0
    return [2 if (i == 0 and toc_sheets == 0) else "2a"
            for i in range(len(docs))]

def pz_doc_first_form(project, path: str) -> int | str:
    """Форма первого листа содержимого документа ПЗ по его пути (2 | "2a").

    Для вызовов, которым известен только путь (редактор ПЗ, пересчёт числа
    листов в дереве/свойствах): 2 — документ открывается первым листом
    раздела без оглавления, "2a" — все остальные случаи (см. pz_doc_first_forms).
    """
    from ..models import DocKind
    try:
        docs = project.documents_for(DocKind.POYASNITELNAYA)
        forms = pz_doc_first_forms(project)
        for i, d in enumerate(docs):
            if d.path and path and os.path.abspath(d.path) == os.path.abspath(path):
                return forms[i] if i < len(forms) else 2
    except Exception:
        pass
    return 2

def _section_sheet_count(project, section_kind: str) -> int:
    """Общее число листов раздела (для графы «Листов» в штампе)."""
    from ..models import (SECTION_COVER, SECTION_TITLE, SECTION_CONTENT,
                          SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX,
                          SECTION_CHANGES, SECTION_COMPOSITION, DocKind)
    if section_kind == SECTION_COVER:
        return 1 if project.structure_needed(SECTION_COVER) else 0
    if section_kind == SECTION_TITLE:
        return 1 if project.structure_needed(SECTION_TITLE) else 0
    if section_kind == SECTION_CONTENT:
        return 1 if project.structure_needed(SECTION_CONTENT) else 0
    if section_kind == SECTION_COMPOSITION:
        # Раздел занимает лист даже с пустой таблицей (на листе — шапка).
        return len(composition_pages(project))
    if section_kind == SECTION_CHANGES:
        if not project.structure_needed(SECTION_CHANGES) or not project.changes:
            return 0
        return 1          # лист регистрации изменений формируется одним листом
    if section_kind == SECTION_PZ:
        # + листы автоматического оглавления записки (app.core.pz_toc).
        # Форма первого листа каждого документа — как в предпросмотре.
        from .pz_toc import project_toc_sheets
        docs = project.documents_for(DocKind.POYASNITELNAYA)
        forms = pz_doc_first_forms(project)
        return (project_toc_sheets(project)
                + sum(_doc_sheet_count(d, first_form=forms[i] if i < len(forms) else 2)
                      for i, d in enumerate(docs)))
    if section_kind == SECTION_GRAPHIC:
        # В разделе графики листы рисуются с формы 6 (first=False в дереве).
        total = sum(_doc_sheet_count(d, first_form="2a")
                    for d in project.documents_for(DocKind.GRAPHIC))
        # Лист «Графическая часть» с перечнем листов — первый лист раздела,
        # поэтому он входит в графу «Листов» (если включён в настройках).
        gl = getattr(project, "graphic_sheet_list", None)
        if gl is not None and gl.enabled and total:
            total += 1
        return total
    if section_kind == SECTION_APPENDIX:
        total = 0
        for a in project.appendices_needed():
            if not a.file_path or not os.path.exists(a.file_path):
                total += 1
            else:
                total += max(1, a.pages) if a.pages else _pdf_count(a.file_path)
        return total
    return 0

def _sheet_fragment(kind: str, snapshot: dict, font_map: FontMap | None = None,
                    counter: _SheetCounter | None = None,
                    sheets_total: int | None = None) -> str:
    """Возвращает разметку ОДНОГО листа (div.sheet) без обёртки документа.

    Для стандартных листов (cover/title/content/changes/stamp) — один лист.
    Для документа/приложения — первый лист его содержимого (полный набор листов
    собирают build_volume_html / build_section_html).
    """
    font_map = font_map or FontMap()
    # Обложка и титульный лист не входят в сквозную нумерацию.
    # Номер отображается квадратом в правом верхнем углу листа.
    if kind in ("cover", "title"):
        sheet_no = ""
    elif counter is not None:
        sheet_no = counter.next()
    else:
        sheet_no = snapshot.get("sheet_no", "") or ""
    num_snapshot = _numbered_snapshot(snapshot, sheet_no, sheets_total)

    if kind == "cover":
        content, stamp = render_cover(num_snapshot, font_map)
    elif kind == "title":
        content, stamp = render_title(num_snapshot, font_map)
    elif kind == "content":
        content, stamp = render_content(num_snapshot, font_map)
    elif kind == "stamp_f1":
        content, stamp = render_stamp_sheet(1, num_snapshot, font_map)
    elif kind == "stamp_f2":
        content, stamp = render_stamp_sheet(2, num_snapshot, font_map)
    elif kind == "stamp_f2a":
        content, stamp = render_stamp_sheet("2a", num_snapshot, font_map)
    elif kind == "changes":
        content, stamp = render_changes(num_snapshot, font_map)
    elif kind == "pz_toc":
        content, stamp = render_pz_toc(num_snapshot, font_map)
    elif kind == "graphic_list":
        content, stamp = render_graphic_sheet_list(num_snapshot, font_map)
    elif kind == "composition":
        content, stamp = render_composition(num_snapshot, font_map)
    elif kind == "doc":
        path = num_snapshot.get("path", "")
        if path and os.path.exists(path):
            if _is_docx(path):
                frags = _docx_sheet_fragments(path, num_snapshot, font_map, first=True,
                                              max_pages=1)
            else:
                frags = _pdf_sheet_fragments(path, num_snapshot, font_map, first=True,
                                             is_appendix=False, max_pages=1)
            if frags:
                return frags[0]
        content = render_placeholder(num_snapshot.get("label", "Документ"),
                                     subtitle=num_snapshot.get("section_label", ""),
                                     path=path, kind="doc")
        if num_snapshot.get("section") == "Пояснительная записка":
            stamp = _stamp_html(2, num_snapshot, font_map)
        else:
            stamp = ""
    elif kind == "appendix":
        path = num_snapshot.get("path", "")
        if path and os.path.exists(path):
            frags = _pdf_sheet_fragments(path, num_snapshot, font_map, first=True,
                                         is_appendix=True, max_pages=1)
            if frags:
                return frags[0]
        content = render_placeholder(num_snapshot.get("label", "Приложение"),
                                     subtitle="Приложение",
                                     path=path, kind="appendix")
        stamp = ""
    elif kind == "section":
        content = render_placeholder(num_snapshot.get("label", "Раздел"),
                                     subtitle="Раздел тома",
                                     path="", kind="section")
        if num_snapshot.get("section") == "Пояснительная записка":
            stamp = _stamp_html(2, num_snapshot, font_map)
        else:
            stamp = ""
    else:
        content, stamp = f'<div>Неизвестный лист: {esc(kind)}</div>', ""

    # Отступ сверху с учётом квадрата сквозной нумерации (indent_top) —
    # для листов, на которых отображается номер (кроме обложки/титульного листа).
    _has_no = bool(num_snapshot.get("sheet_no"))
    return _wrap_sheet(content, stamp,
                       sheet_no=num_snapshot.get("square_no",
                                                 num_snapshot.get("sheet_no", "")),
                       indent_top=_has_no)

def _section_square_base(project, section_kind: str) -> int | None:
    """Первый сквозной номер раздела в томе (по toc.build_layout).

    Используется для квадрата сквозной нумерации при сборке тома: квадрат
    листа = base + локальный номер внутри раздела − 1, поэтому он совпадает
    с содержанием, а графы «Лист/Листов» штампа остаются по разделу.
    """
    from ..models import (SECTION_COVER, SECTION_TITLE)
    if section_kind in (SECTION_COVER, SECTION_TITLE):
        return None          # обложка/титул не нумеруются
    try:
        from .toc import build_layout
        for it in build_layout(project).items:
            if it.section_key == section_kind and it.first_sheet >= 1:
                return int(it.first_sheet)
    except Exception:
        return None
    return None

def _section_fragments(project, section_kind: str, snapshot: dict,
                       font_map: FontMap, counter: _SheetCounter | None = None,
                       sheets_total: int | None = None,
                       volume_numbered: bool = False) -> list:
    """Все листы раздела section_kind, включая многостраничные документы.

    Графы «Лист»/«Листов» форм 5/6 ведутся ПО РАЗДЕЛУ (части тома): счётчик
    раздела стартует с 1, «Листов» = число листов раздела.

    volume_numbered=True (сборка тома): в snapshot подставляется square_base —
    первый сквозной номер раздела, поэтому квадрат сквозной нумерации и номера
    строк оглавления ПЗ остаются сквозными по тому (как в содержании).
    Этот же режим использует предпросмотр раздела: квадрат в правом верхнем
    углу обязан совпадать с содержанием, поэтому он сквозной и на экране.
    Без base (отдельная сборка раздела через build_section_html) квадрат
    локальный — раздел нумеруется с 1. Обложка и титульный лист номер не
    получают и счётчик не увеличивают.
    """
    from ..models import (SECTION_COVER, SECTION_TITLE, SECTION_CONTENT,
                          SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX,
                          SECTION_CHANGES, SECTION_COMPOSITION, DocKind)
    frags = []
    if volume_numbered:
        # Том: локальный счётчик раздела + квадрат от сквозного base.
        counter = _SheetCounter(1)
        sheets_total = _section_sheet_count(project, section_kind)
        base = _section_square_base(project, section_kind)
        if base is not None:
            snapshot = dict(snapshot)
            snapshot["square_base"] = base
    else:
        # Standalone-раздел: локальная нумерация (или от вызывающего кода).
        counter = counter or _SheetCounter(1)
        sheets_total = (sheets_total if sheets_total is not None
                        else _section_sheet_count(project, section_kind))
    if section_kind == SECTION_COVER:
        frags.append(_sheet_fragment("cover", snapshot, font_map,
                                     counter=counter, sheets_total=sheets_total))
    elif section_kind == SECTION_TITLE:
        frags.append(_sheet_fragment("title", snapshot, font_map,
                                     counter=counter, sheets_total=sheets_total))
    elif section_kind == SECTION_CONTENT:
        sub = dict(snapshot)
        sub.update({"section": SECTION_CONTENT, "section_label": "Содержание"})
        frags.append(_sheet_fragment("content", sub, font_map,
                                     counter=counter, sheets_total=sheets_total))
    elif section_kind == SECTION_COMPOSITION:
        # Все листы таблицы: первый — форма 5, последующие — форма 6.
        from ..models import composition_rows
        rows = composition_rows(project)
        for index in range(len(composition_pages_from_rows(rows))):
            sub = dict(snapshot)
            sub.update({"section": SECTION_COMPOSITION,
                        "section_label": "Состав проектной документации",
                        "composition_rows": rows,
                        "composition_page": index})
            frags.append(_sheet_fragment("composition", sub, font_map,
                                         counter=counter,
                                         sheets_total=sheets_total))
    elif section_kind == SECTION_CHANGES:
        if project.changes:
            sub = dict(snapshot)
            sub.update({"section": SECTION_CHANGES,
                        "section_label": "Лист регистрации изменений"})
            frags.append(_sheet_fragment("changes", sub, font_map,
                                         counter=counter, sheets_total=sheets_total))
    elif section_kind == SECTION_PZ:
        docs = project.documents_for(DocKind.POYASNITELNAYA)
        # Нет включённых документов — раздел не занимает лист в томе
        # (см. toc.build_layout): иначе после снятия флажка «в сборку»
        # в томе оставался бы пустой лист-заглушка.
        first = True
        for d in docs:
            sub = dict(snapshot)
            sub.update({"path": d.path, "section": SECTION_PZ,
                        "section_label": "Пояснительная записка",
                        "label": os.path.basename(d.path) if d.path else "—",
                        "word_pagination": snapshot.get("word_pagination", {}),
                        "pz_edits": snapshot.get("pz_edits", {}),
                        "pz_cell_edits": snapshot.get("pz_cell_edits", {})})
            dfrags = _doc_sheet_fragments(d, sub, font_map, first=first,
                                          counter=counter, sheets_total=sheets_total)
            frags.extend(dfrags)
            if dfrags:
                first = False
    elif section_kind == SECTION_GRAPHIC:
        # Лист «Графическая часть» с перечнем листов — ПЕРВЫМ листом раздела
        # (рамка + основная надпись формы 3). Отключается в настройках проекта
        # (models.GraphicSheetList.enabled). Пустой раздел (нет включённых
        # документов) листа не занимает — как и в содержании (toc.build_layout):
        # иначе в томе оставался бы одинокий лист перечня без самих листов.
        gl = getattr(project, "graphic_sheet_list", None)
        graphic_docs = project.documents_for(DocKind.GRAPHIC)
        if gl is not None and gl.enabled and graphic_docs:
            sub = dict(snapshot)
            sub.update({"section": SECTION_GRAPHIC,
                        "section_label": "Графическая часть",
                        "graphic_sheet_list": gl.to_dict(),
                        "graphic_sheet_list_rows": build_graphic_sheet_list(project)})
            frags.append(_sheet_fragment("graphic_list", sub, font_map,
                                         counter=counter,
                                         sheets_total=sheets_total))
        for d in graphic_docs:
            sub = dict(snapshot)
            sub.update({"path": d.path, "section": SECTION_GRAPHIC,
                        "section_label": "Графическая часть",
                        "label": os.path.basename(d.path) if d.path else "—"})
            frags.extend(_doc_sheet_fragments(d, sub, font_map, first=False,
                                              counter=counter, sheets_total=sheets_total))
    elif section_kind == SECTION_APPENDIX:
        for a in project.appendices_needed():
            sub = dict(snapshot)
            sub.update({"path": a.file_path, "section": SECTION_APPENDIX,
                        "section_label": "Приложение",
                        "label": (f"Приложение {a.number}. {a.name}"
                                  if a.number else a.name)})
            frags.extend(_appendix_sheet_fragments(a, sub, font_map,
                                                   counter=counter,
                                                   sheets_total=sheets_total))
    return frags

def render_sheet(kind: str, snapshot: dict, font_map: FontMap | None = None) -> str:
    """Возвращает полный HTML-документ одного листа (первый лист содержимого)."""
    fragment = _sheet_fragment(kind, snapshot, font_map)
    return _preview_document([fragment], snapshot)

def render_sheet_fragments(kind: str, snapshot: dict,
                           font_map: FontMap | None = None,
                           project=None) -> list:
    """Возвращает разметку ВСЕХ листов выбранного пункта дерева.

    Для стандартных листов (обложка/титул/содержание/изменения) и отдельных
    форм — один лист. Для раздела тома (ПЗ / графика / приложения) и для
    многостраничного документа — все его листы подряд. Графы «Лист/Листов»
    форм 5/6 считаются в пределах раздела (части тома), а квадрат сквозной
    нумерации — СКВОЗНОЙ по тому (как в содержании и в собранном HTML).
    """
    from ..models import (SECTION_COVER, SECTION_TITLE, SECTION_CONTENT,
                          SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX,
                          SECTION_CHANGES, SECTION_COMPOSITION, DocKind)
    font_map = font_map or FontMap()

    # Лист/документ с явно заданным путём — отдаём все листы этого документа.
    if kind in ("doc", "appendix"):
        return _all_document_fragments(kind, snapshot, font_map, project)

    # Раздел тома, для которого нужен полный набор листов.
    # Раздел определяется по полю section; если его нет (старый проект, другой
    # вызывающий код), пробуем по section_label — иначе вместо листов раздела
    # показывалась заглушка «Раздел тома».
    section = snapshot.get("section", "")
    doc_sections = (SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX,
                    SECTION_COMPOSITION)
    if kind == "section" and section not in doc_sections:
        label = (snapshot.get("section_label") or snapshot.get("label") or "").strip()
        if label in doc_sections:
            section = label
    if kind == "section" and section in doc_sections:
        sub = dict(snapshot)
        sub["section"] = section
        return _all_section_fragments(section, sub, font_map, project)

    # Обложка/титул/содержание/изменения и отдельные формы ГОСТ — один лист.
    if kind in ("cover", "title", "content", "changes",
                "stamp_f1", "stamp_f2", "stamp_f2a"):
        # Содержание/изменения при известном project: «Лист 1 / Листов N»
        # (листы только этого раздела); без project — как в снимке.
        if kind in ("content", "changes") and project is not None:
            sec_kind = (SECTION_CONTENT if kind == "content" else SECTION_CHANGES)
            sub = dict(snapshot)
            sub["sheet_no"] = "1"
            sub["section"] = sec_kind
            sub["sheets_total"] = str(_section_sheet_count(project, sec_kind))
            # Квадрат сквозной по тому: у «Листа изменений» номер в томе не 1.
            sub = _with_square_base(project, sec_kind, sub)
            return [_sheet_fragment(kind, sub, font_map)]
        return [_sheet_fragment(kind, snapshot, font_map)]

    # Прочие документы/разделы — один лист (заглушка или первая страница).
    return [_sheet_fragment(kind, snapshot, font_map)]

def _all_document_fragments(kind: str, snapshot: dict, font_map: FontMap,
                             project=None) -> list:
    """Все листы документа/приложения из снимка предпросмотра (path в snapshot).

    Нумерация продолжает РАЗДЕЛ (часть тома), как и в предпросмотре раздела:
    счётчик стартует со смещением — листами, стоящими перед этим документом
    в разделе (для первого документа ПЗ смещение 0: листы оглавления входят
    в его поток). «Листов» — общее число листов раздела.
    """
    from ..models import (SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX, DocKind)
    path = snapshot.get("path", "")
    if not path or not os.path.exists(path):
        return [_sheet_fragment(kind, snapshot, font_map)]
    if project is None:
        if kind == "appendix":
            return _pdf_sheet_fragments(path, snapshot, font_map, first=True,
                                        is_appendix=True)
        if _is_docx(path):
            return _docx_sheet_fragments(path, snapshot, font_map, first=True)
        return _pdf_sheet_fragments(path, snapshot, font_map, first=True,
                                    is_appendix=False)

    def _find(docs, attr="path") -> int | None:
        return next((i for i, d in enumerate(docs)
                     if getattr(d, attr, "") and path
                     and os.path.abspath(getattr(d, attr)) == os.path.abspath(path)),
                    None)

    # --- Приложение: продолжение нумерации раздела приложений ---
    if kind == "appendix":
        from .toc import appendix_sheet_count
        snapshot = _with_square_base(project, SECTION_APPENDIX, snapshot)
        apps = project.appendices_needed()
        idx = _find(apps, "file_path")
        counter = None
        if idx is not None:
            offset = sum(appendix_sheet_count(a) for a in apps[:idx])
            counter = _SheetCounter(offset + 1)
        return _pdf_sheet_fragments(path, snapshot, font_map, first=True,
                                    is_appendix=True, counter=counter,
                                    sheets_total=_section_sheet_count(
                                        project, SECTION_APPENDIX))

    # --- Документ ПЗ: сдвиг и форма первого листа — как в разделе ---
    docs = project.documents_for(DocKind.POYASNITELNAYA)
    idx = _find(docs)
    if idx is not None:
        snapshot = _with_square_base(project, SECTION_PZ, snapshot)
        forms = pz_doc_first_forms(project)
        offset = sum(_doc_sheet_count(d, first_form=forms[i] if i < len(forms) else 2)
                     for i, d in enumerate(docs[:idx]))
        if idx > 0:
            # Оглавление стоит перед ПЕРВЫМ документом — для последующих
            # документов оно тоже входит в число предшествующих листов.
            try:
                from .pz_toc import project_toc_sheets
                offset += project_toc_sheets(project)
            except Exception:
                pass
        counter = _SheetCounter(offset + 1)
        sheets_total = _section_sheet_count(project, SECTION_PZ)
        # Показ документа совпадает с показом раздела: оглавление и
        # форма 5 — только у ПЕРВОГО листа раздела.
        doc_first = (idx == 0)
        if _is_docx(path):
            return _docx_sheet_fragments(path, snapshot, font_map, first=doc_first,
                                         counter=counter, sheets_total=sheets_total)
        return _pdf_sheet_fragments(path, snapshot, font_map, first=doc_first,
                                    is_appendix=False, counter=counter,
                                    sheets_total=sheets_total)

    # --- Документ графической части: квадрат продолжает раздел ---
    gdocs = project.documents_for(DocKind.GRAPHIC)
    gidx = _find(gdocs)
    if gidx is not None:
        snapshot = _with_square_base(project, SECTION_GRAPHIC, snapshot)
        offset = sum(_doc_sheet_count(d, first_form="2a") for d in gdocs[:gidx])
        counter = _SheetCounter(offset + 1)
        return _pdf_sheet_fragments(path, snapshot, font_map, first=False,
                                    is_appendix=False, counter=counter,
                                    sheets_total=_section_sheet_count(
                                        project, SECTION_GRAPHIC))

    # --- Прочее (вне структуры) — прежнее поведение ---
    if _is_docx(path):
        return _docx_sheet_fragments(path, snapshot, font_map, first=True)
    return _pdf_sheet_fragments(path, snapshot, font_map, first=True,
                                is_appendix=False)

def _with_square_base(project, section_kind: str, snapshot: dict) -> dict:
    """Копия снимка с базой сквозной нумерации раздела.

    Квадрат в правом верхнем углу — СКВОЗНОЙ по тому (он совпадает с номерами
    в содержании), поэтому предпросмотр обязан знать первый номер раздела так
    же, как сборка тома. Без базы квадрат начинался бы с 1 в каждом разделе.
    """
    if project is None or not section_kind:
        return snapshot
    base = _section_square_base(project, section_kind)
    if base is None:
        return snapshot
    sub = dict(snapshot)
    sub["square_base"] = base
    return sub

def _all_section_fragments(section_kind: str, snapshot: dict,
                           font_map: FontMap, project=None) -> list:
    """Все листы раздела тома (ПЗ / графика / приложения) для предпросмотра.

    Графы «Лист/Листов» ведутся по разделу, а квадрат сквозной нумерации —
    сквозной по тому (volume_numbered=True), как в «Собрать том в HTML»:
    номера на листах совпадают с содержанием.
    """
    if project is None:
        # Без проекта доступен только первый лист раздела.
        return [_sheet_fragment("section", snapshot, font_map)]
    counter = _SheetCounter(1)
    # Интерактивный предпросмотр: листы графической части показываются
    # пустыми, без растра чертежей — остаётся только квадрат нумерации, по
    # которому видно, какой лист открыт и совпадает ли номер с содержанием.
    # Растр в предпросмотре не нужен: он тяжёлый (десятки мегабайт на том) и
    # не используется при сборке — в PDF графика попадает вектором из
    # исходного файла. Приложения и ПЗ-PDF рендерятся как раньше.
    snapshot = dict(snapshot)
    snapshot["preview_blank_graphic"] = True
    frags = _section_fragments(project, section_kind, snapshot, font_map,
                               counter=counter, volume_numbered=True)
    if frags:
        return frags
    # В разделе нет включённых в сборку документов — это не ошибка, поэтому
    # вместо пустого листа показываем подсказку (в сборку раздел не попадает).
    note = ("В разделе нет документов, включённых в сборку. "
            "Включите документ флажком «В сборку».")
    sub = dict(snapshot)
    sub["section"] = section_kind
    return [_placeholder_sheet_fragment("section", sub, font_map, note=note)]

def _signer_registry_html(snapshot) -> str:
    """Скрытый реестр подписантов: фамилия → изображение подписи.

    Нужен скрипту предпросмотра, чтобы подобрать подпись по фамилии сразу
    после правки ячейки фамилии, не дожидаясь перерисовки листа. Изображение
    встраивается как data-URI (`signature_src`): ссылки `file:///` со страницы,
    загруженной через `setHtml(about:blank)`, QtWebEngine не отдаёт.
    """
    items = []
    for signer in (snapshot.get("signers") or []):
        surname = str(signer.get("surname") or "")
        if not surname:
            continue
        src = signature_src(str(signer.get("signature_image") or ""))
        items.append(f'<span data-signer-surname="{esc(surname)}" '
                     f'data-signer-img="{esc(src)}"></span>')
    if not items:
        return ""
    return ('<div id="signer-registry" style="display:none" aria-hidden="true">'
            + "".join(items) + "</div>")

def _preview_document(fragments: list, snapshot: dict | None = None) -> str:
    """Полный HTML-документ интерактивного предпросмотра (один или много листов)."""
    fragment = "\n".join(fragments)
    registry = _signer_registry_html(snapshot or {})
    return f"""<!doctype html>
<html><head><meta charset="utf-8">
<script src="qrc:///qtwebchannel/qwebchannel.js"></script>
<style>{_css()}</style></head>
<body>{registry}{fragment}
<script>
(function() {{
  // Автоподгонка кегля для надписей/полей, помеченных class="fit-lbl":
  // уменьшает font-size, пока содержимое не перестанет переполнять ячейку.
  function fitLabels() {{
    var nodes = document.querySelectorAll('.fit-lbl');
    nodes.forEach(function(el) {{
      var text = el.textContent;
      if (!text) {{ return; }}
      // стартовый кегль берём из inline-стиля
      var m = /font-size:([\\d.]+)pt/.exec(el.getAttribute('style') || '');
      var start = m ? parseFloat(m[1]) : 7;
      // целевой элемент: сам или вложенное поле/подпись
      var target = el.querySelector('.field') || el;
      var found = false;
      for (var fs = start; fs >= 3.5; fs -= 0.5) {{
        target.style.fontSize = fs + 'pt';
        // Уменьшаем кегль, пока текст не перестанет переполнять ячейку и по
        // горизонтали, и по вертикали (в графе 9 текст переносится на строки,
        // поэтому высота тоже может выйти за рамку).
        if (el.scrollWidth <= el.clientWidth + 1 &&
            el.scrollHeight <= el.clientHeight + 1) {{ found = true; break; }}
      }}
      if (!found) {{
        target.style.fontSize = Math.max(3.5, start) + 'pt';
      }}
    }});
  }}

  // Подсказка поля (data-placeholder) выводится правилом .field.empty::after,
  // поэтому класс empty должен сниматься, как только в поле появился текст:
  // иначе подсказка дописывается после значения («12-2026-ТРОобозначение тома»)
  // и значение остаётся серым курсивом. У полей без подсказки атрибута нет,
  // но класс empty ставится при пустом значении — его тоже нужно снимать.
  function syncEmpty(field) {{
    if (!field || !field.classList) {{ return; }}
    if (!field.hasAttribute('data-placeholder') &&
        !field.classList.contains('empty')) {{ return; }}
    field.classList.toggle('empty', !field.innerText.trim());
  }}

  // ---- Границы листов задаёт Word -------------------------------------
  // Разбивка на листы рассчитывается ТОЛЬКО в Word (см. word_pagination):
  // предпросмотр раскладывает блоки строго по номерам страниц из разметки и
  // никогда не переливает их между листами сам. После правки текста Python
  // помечает разметку устаревшей и запускает новый расчёт в Word.

  // Абзац DOCX адресуется data-para-index. Word может разрезать абзац между
  // листами — тогда на каждом листе выводится своя часть с тем же адресом.
  // Правка собирает текст всех частей и уходит в Python одним значением,
  // иначе записалась бы только часть абзаца. Порядок частей — порядок листов,
  // то есть document order выборки.
  function paraParts(host) {{
    var idx = host.getAttribute('data-para-index');
    if (idx === null) {{ return [host]; }}
    var file = host.getAttribute('data-para-file') || '';
    var same = Array.prototype.filter.call(
      document.querySelectorAll('.pz-para[data-para-index="' + idx + '"]'),
      function (el) {{
        return (el.getAttribute('data-para-file') || '') === file && el.isConnected;
      }});
    return same.length ? same : [host];
  }}

  // Текст абзаца целиком: части склеиваются через пробел, если он потерялся
  // на границе разреза Word.
  function joinParts(parts) {{
    var out = '';
    parts.forEach(function (el) {{
      var t = el.innerText;
      if (!t) {{ return; }}
      if (out && !/\\s$/.test(out) && !/^\\s/.test(t)) {{ out += ' '; }}
      out += t;
    }});
    return out;
  }}

  function init() {{
    // Автоподгонка подписей (до создания моста, т.к. мост может быть недоступен
    // в оффскрин-рендере; в приложении повторно вызовется после onSheetReady).
    fitLabels();
    // Редактируемые поля: подсказка и передача значения в Python. Обработчик
    // ставится до проверки моста, чтобы подсказка снималась и без QWebChannel.
    var fields = document.querySelectorAll('.field[contenteditable="true"]');
    fields.forEach(function(f) {{
      syncEmpty(f);
      f.addEventListener('input', function(e) {{
        syncEmpty(f);
        if (!window.bridge || !window.bridge.onFieldChanged) {{ return; }}
        window.bridge.onFieldChanged(f.getAttribute('data-key'), f.innerText);
      }});
    }});
    if (typeof QWebChannel === 'undefined' || typeof qt === 'undefined') {{ return; }}
    new QWebChannel(qt.webChannelTransport, function(channel) {{
      window.bridge = channel.objects.bridge;
      if (!window.bridge) {{ return; }}
      // Положение прокрутки: Python запоминает его с задержкой после остановки
      // прокрутки, чтобы после перерисовки листа (правка поля) вернуться на то
      // же место, а не «перескочить» на начало тома.
      // Активный лист — тот, что занимает наибольшую часть окна: по нему
      // панель внизу окна показывает основную надпись в масштабе 100%.
      var activeIdx = -1;
      function updateActiveSheet() {{
        var sheets = document.querySelectorAll('.sheet');
        if (!sheets.length) {{ return; }}
        var vh = window.innerHeight;
        var best = 0, bestVis = -1;
        for (var i = 0; i < sheets.length; i++) {{
          var r = sheets[i].getBoundingClientRect();
          var vis = Math.min(r.bottom, vh) - Math.max(r.top, 0);
          if (vis > bestVis) {{ bestVis = vis; best = i; }}
        }}
        if (best !== activeIdx) {{
          activeIdx = best;
          if (window.bridge.onActiveSheet) {{ window.bridge.onActiveSheet(best); }}
        }}
      }}
      var scrollTimer = null;
      window.addEventListener('scroll', function () {{
        updateActiveSheet();
        if (scrollTimer) {{ clearTimeout(scrollTimer); }}
        scrollTimer = setTimeout(function () {{
          if (window.bridge.onScrollChanged) {{
            window.bridge.onScrollChanged(window.scrollY || 0);
          }}
        }}, 150);
      }});
      // Сообщаем активный лист ещё раз: при первом вызове моста могло не
      // быть (QWebChannel подключается асинхронно), и панель внизу окна
      // осталась бы на прежнем листе.
      activeIdx = -1;
      updateActiveSheet();
      // Поля подписей (клик -> выбор подписанта)
      var sigs = document.querySelectorAll('.field[contenteditable="false"].signature');
      sigs.forEach(function(f) {{
        f.addEventListener('click', function(e) {{
          var key = f.getAttribute('data-key');
          if (window.bridge.onSignatureClicked) {{
            window.bridge.onSignatureClicked(key);
          }}
        }});
      }});
      // Строки подписей левого блока штампа: правка фамилии сразу показывает
      // подпись этого подписанта из реестра (изображение подбирается по
      // фамилии), а клик по графе «Подп.» открывает выбор подписанта — тогда
      // фамилия подставляется из реестра и подпись появляется сама.
      var EMPTY_SIG = '<span class="sig-placeholder">(подпись)</span>';
      function sigImageFor(surname) {{
        var s = String(surname || '').trim().toLowerCase();
        if (!s) {{ return ''; }}
        var found = '';
        document.querySelectorAll('[data-signer-surname]').forEach(function(el) {{
          if (found) {{ return; }}
          if ((el.getAttribute('data-signer-surname') || '').toLowerCase() === s) {{
            found = el.getAttribute('data-signer-img') || '';
          }}
        }});
        return found;
      }}
      function syncSigPreview(rowKey) {{
        var nameCell = document.querySelector('.field[data-key="ss|' + rowKey + '|surname"]');
        var sigCell = document.querySelector('.field[data-key="ss|' + rowKey + '|signature"]');
        if (!sigCell) {{ return; }}
        var src = sigImageFor(nameCell ? nameCell.innerText : '');
        if (src) {{
          sigCell.innerHTML = '<img class="sig-img sig-img-row" src="' + src + '" alt="">';
        }} else {{
          sigCell.innerHTML = EMPTY_SIG;
        }}
      }}
      document.querySelectorAll('.field[data-key^="ss|"]').forEach(function(f) {{
        var key = f.getAttribute('data-key') || '';
        var parts = key.split('|');
        if (parts.length !== 4 || parts[3] !== 'surname') {{ return; }}
        f.addEventListener('input', function(e) {{
          // Ключ строки — «<раздел>|<номер>»: подпись стоит на этом же листе.
          // Значение в проект записывает общий обработчик .field (см. init),
          // здесь только сразу показываем подпись этого подписанта.
          syncSigPreview(parts[1] + '|' + parts[2]);
        }});
      }});
      // Абзацы пояснительной записки: правка прямо в предпросмотре.
      // Индекс абзаца в DOCX приходит в data-para-index; по нему Python
      // записывает текст в исходный файл (кнопка «Сохранить в файл»).
      var paras = document.querySelectorAll('.pz-para[contenteditable="true"]');
      paras.forEach(function(p) {{
        p.addEventListener('input', function(e) {{
          var idx = parseInt(p.getAttribute('data-para-index'), 10);
          if (isNaN(idx) || idx < 0) {{ return; }}
          var file = p.getAttribute('data-para-file') || '';
          // Абзац Word мог быть разрезан между листами — в Python уходит
          // текст всех частей, иначе записалась бы только одна. Границы
          // листов пересчитает Word после сохранения правок.
          if (window.bridge.onParagraphChanged) {{
            window.bridge.onParagraphChanged(idx, file, joinParts(paraParts(p)));
          }}
        }});
        // Enter в абзаце не создаёт новый <p> (это одна запись текста), но
        // вставляет мягкий перенос — как Shift+Enter в Word.
        p.addEventListener('keydown', function(e) {{
          if (e.key !== 'Enter') {{ return; }}
          e.preventDefault();
          if (!document.execCommand('insertLineBreak')) {{
            document.execCommand('insertHTML', false, '<br>');
          }}
          p.dispatchEvent(new Event('input', {{ bubbles: true }}));
        }});
      }});
      // Ячейки таблиц пояснительной записки: правка прямо в предпросмотре.
      // Адрес ячейки — data-table/data-row/data-cell (индексы в document.tables);
      // по нему Python записывает текст обратно в исходный DOCX.
      var cells = document.querySelectorAll('.pz-cell[contenteditable="true"]');
      cells.forEach(function(c) {{
        c.addEventListener('input', function(e) {{
          var table = parseInt(c.getAttribute('data-table'), 10);
          var row = parseInt(c.getAttribute('data-row'), 10);
          var cell = parseInt(c.getAttribute('data-cell'), 10);
          if (isNaN(table) || isNaN(row) || isNaN(cell)) {{ return; }}
          if (table < 0 || row < 0 || cell < 0) {{ return; }}
          var file = c.getAttribute('data-para-file') || '';
          if (window.bridge.onTableCellChanged) {{
            window.bridge.onTableCellChanged(table, row, cell, file, c.innerText);
          }}
        }});
      }});
      // Пере-подгонка после готовности моста
      fitLabels();
      if (window.bridge.onSheetReady) {{
        window.bridge.onSheetReady();
      }}
    }});
  }}
  if (document.readyState === 'loading') {{
    document.addEventListener('DOMContentLoaded', init);
  }} else {{
    init();
  }}
}})();
</script>
</body></html>"""

# --------------------------------------------------------------------------
# Сборка тома / раздела в HTML-файл
# --------------------------------------------------------------------------

def _html_document(fragments: list, title: str) -> str:
    """Оборачивает список листов (fragments) в единый HTML-документ.

    Все листы идут подряд (каждый в своём блоке .sheet). Кнопки «Собрать»/
    «Собрать раздел» для собранного файла не нужны — это готовый документ.
    """
    sheets = "\n".join(fragments)
    return f"""<!doctype html>
<html><head><meta charset="utf-8">
<title>{esc(title)}</title>
<style>{_css()}</style>
<style>
  @media print {{ .sheet {{ page-break-after: always; }} }}
  /* В собранном документе (HTML/PDF) скрываем редакторские подсказки:
     пунктирную линию под полем, подсветку фокуса и серый placeholder.
     Они нужны только в интерактивном предпросмотре (render_sheet), где
     поле редактируется inline. В итоговом файле поле должно выглядеть чисто. */
  .field {{ border-bottom:none !important; }}
  .field:focus {{ background:transparent !important; border-bottom:none !important; }}
  .field.empty {{ color:#111 !important; font-style:normal !important; }}
  .field.empty::after {{ content:"" !important; }}
  /* Подсказка «подпись» в графе «Подп.» — тоже только редакторская: в
     собранном томе графа остаётся пустой, если подпись не выбрана (иначе в
     PDF/HTML печатался служебный текст). Изображение подписи выводится как есть. */
  .sig-placeholder {{ display:none !important; }}
</style>
</head>
<body>
{sheets}
<script>
// Автоподгонка кегля в ячейках, помеченных class="fit-lbl" (в т.ч. графа 9 —
// наименование организации). В собранном документе скрипт предпросмотра не
// выполняется, поэтому уменьшаем кегль здесь: без этого длинное наименование
// вылезало за рамку графы при печати PDF/HTML.
(function() {{
  function fit() {{
    document.querySelectorAll('.fit-lbl').forEach(function(el) {{
      if (!el.textContent) {{ return; }}
      var m = /font-size:([0-9.]+)pt/.exec(el.getAttribute('style') || '');
      var start = m ? parseFloat(m[1]) : 7;
      var target = el.querySelector('.field') || el;
      var found = false;
      for (var fs = start; fs >= 3.5; fs -= 0.5) {{
        target.style.fontSize = fs + 'pt';
        // Проверяем и горизонталь, и вертикаль: текст графы 9 переносится
        // на строки и может переполнить ячейку по высоте.
        if (el.scrollWidth <= el.clientWidth + 1 &&
            el.scrollHeight <= el.clientHeight + 1) {{ found = true; break; }}
      }}
      if (!found) {{ target.style.fontSize = Math.max(3.5, start) + 'pt'; }}
    }});
  }}
  if (document.readyState === 'complete') {{ fit(); }}
  else {{ window.addEventListener('load', fit); }}
}})();
</script>
</body></html>"""

def _basename(path: str) -> str:
    import os
    return os.path.basename(path) if path else "—"

def build_volume_html(project, out_path: str, font_map: FontMap | None = None) -> str:
    """Собирает весь том в единый HTML-файл (листы идут подряд, с разрывом при печати).

    :param project: VolumeProject
    :param out_path: путь к создаваемому .html
    :return: путь к файлу
    """
    from ..core.preview_model import build_preview_snapshot
    font_map = font_map or FontMap()
    snapshot = build_preview_snapshot(project)
    # Собранный документ: без редакторских подсказок (в т.ч. «(подпись)»).
    snapshot["hide_signature_hint"] = True
    fragments = []
    # Обходим структуру тома: каждый раздел нумеруется отдельно (графы
    # «Лист/Листов» форм 5/6 — листы только этого раздела). Квадрат
    # сквозной нумерации считается внутри раздела от square_base (см.
    # _section_fragments / volume_numbered).
    for entry in project.structure:
        if not entry.get("needed", True):
            continue
        kind = entry.get("kind")
        fragments.extend(_section_fragments(project, kind, snapshot, font_map,
                                            volume_numbered=True))
    title = project.requisites.volume_name or project.name or "Том"
    html = _html_document(fragments, title)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path

def build_section_html(project, section_kind: str, out_path: str,
                       font_map: FontMap | None = None) -> str:
    """Собирает раздел тома в единый HTML-файл (листы идут подряд).

    :param project: VolumeProject
    :param section_kind: ключ раздела ('cover'|'title'|'content'|'changes'|
        DocKind.POYASNITELNAYA.value|DocKind.GRAPHIC.value|'appendix')
    :param out_path: путь к создаваемому .html
    :return: путь к файлу
    """
    from ..core.preview_model import build_preview_snapshot
    from ..models import SECTION_LABELS
    font_map = font_map or FontMap()
    snapshot = build_preview_snapshot(project)
    # Собранный документ: без редакторских подсказок (в т.ч. «(подпись)»).
    snapshot["hide_signature_hint"] = True
    fragments = _section_fragments(project, section_kind, snapshot, font_map)
    if not fragments:
        # В разделе нет включённых в сборку документов: пустой файл без
        # пояснения выглядел бы как ошибка сборки.
        sub = dict(snapshot)
        sub["section"] = section_kind
        sub["label"] = SECTION_LABELS.get(section_kind, section_kind)
        fragments = [_placeholder_sheet_fragment(
            "section", sub, font_map,
            note="В разделе нет документов, включённых в сборку. "
                 "Включите документ флажком «В сборку».")]
    title = SECTION_LABELS.get(section_kind, section_kind)
    html = _html_document(fragments, title)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path
