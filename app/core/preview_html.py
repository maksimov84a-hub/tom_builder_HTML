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

# Высота одной строки текста (мм) при line-height:1.4 для кеглей 11/13pt.
# Измерено по печати: одна строка 11pt ≈ 5.5 мм, 13pt ≈ 6.1 мм.
_LINE_H_MM = {11: 5.5, 13: 6.1}

# Высота одной строки таблицы DOCX (мм). Строка из 4 ячеек с коротким
# текстом занимает ~4.3 мм; с длинным текстом, переносящимся на несколько
# визуальных строк, — кратно числу таких строк. Число визуальных строк ячейки
# оценивается по переносу текста при фактической ширине ячейки.
_TABLE_CELL_LINE_MM = 4.3

# Средняя ширина одного печатного знака (мм) для кегля 9pt (ячейки таблиц).
# Подобрано по измерению: ячейка 4-колоночной таблицы (~43 мм) вмещает
# ~19 знаков в строку, 3-колоночной (~58 мм) — ~30 знаков.
_TABLE_AVG_CHAR_MM = 0.5

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


def _line_height(size: int) -> float:
    """Высота одной визуальной строки текста (мм) для кегля size."""
    return _LINE_H_MM.get(size, 5.5)


def _chars_per_line(size: int, fmt=None) -> int:
    """Число знаков в строке с учётом абзацных отступов Word (мм).

    Отступы (в т.ч. отступы списка) сужают текстовую колонку, поэтому оценка
    переноса без их учёта занижала бы число строк, и абзац уезжал бы под штамп.
    """
    chars = _CHARS_PER_LINE.get(size, 81)
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


def _estimate_paragraph_h(text: str, size: int, fmt=None) -> float:
    """Оценка высоты абзаца (мм): визуальные строки * высота строки + интервалы.

    fmt — оформление абзаца Word (app.core.pz_format.ParagraphFormat): из него
    берутся межстрочный интервал (множитель к базовой высоте), абзацные
    отступы до/после (space before/after) и сужение колонки отступами.
    """
    height = _text_visual_lines(text, _chars_per_line(size, fmt)) * _line_height(size)
    if fmt is None:
        return height
    height *= fmt.line_height_factor()
    height += (float(getattr(fmt, "space_before_pt", 0.0) or 0.0)
               + float(getattr(fmt, "space_after_pt", 0.0) or 0.0)) * PT_PER_MM
    return height


def _table_ncols(table) -> int:
    """Число колонок таблицы DOCX (по самой широкой строке)."""
    if not table.rows:
        return 1
    return max((len(row) for row in table.rows), default=1)


def _table_cell_chars_per_line(ncols: int, cell_index: int = 0, cell_w_mm: float = 0.0) -> int:
    """Число печатных знаков, помещающихся в одну строку ячейки таблицы."""
    cell_w = cell_w_mm if cell_w_mm > 0 else _PAGE_TEXT_W_MM / max(1, ncols)
    return max(1, int((cell_w - 4.0) * _TABLE_AVG_CHAR_MM))


def _estimate_table_row_h(row, chars_per_line: int, cell_chars: list | None = None) -> float:
    """Оценка высоты одной строки таблицы DOCX (мм).

    cell_chars — число знаков в строке для каждой колонки (если ширины колонок
    известны из DOCX); иначе используется общая оценка chars_per_line.
    """
    row_lines = 1
    for i, cell in enumerate(row):
        txt = getattr(cell, "text", None) or ""
        per_line = chars_per_line
        if cell_chars and i < len(cell_chars) and cell_chars[i] > 0:
            per_line = cell_chars[i]
        row_lines = max(row_lines, _text_visual_lines(txt, per_line))
    return row_lines * _TABLE_CELL_LINE_MM


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


def _sig_field(key, role_label, signer, *, size=10, compact=False) -> str:
    """Поле подписи: изображение (или плейсхолдер) + Фамилия."""
    ig = ' sig-img-c' if compact else ''
    sn = ' sig-surname-c' if compact else ''
    if compact:
        # В компактном режиме (узкая ячейка) показываем только фамилию мелким кеглем.
        surname = f'<span class="sig-surname{sn}">{esc(signer.get("surname", "") if signer else "")}</span>'
        img = ''
    elif signer and signer.get("signature_image"):
        img = (f'<img class="sig-img{ig}" src="file:///{esc(signer["signature_image"])}" '
               f'alt="{esc(signer.get("surname", ""))}">')
        surname = f'<span class="sig-surname{sn}">{esc(signer.get("surname", ""))}</span>'
    else:
        img = '<span class="sig-img sig-placeholder">(подпись)</span>'
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
.frame {
  position:absolute;
  left:{left}mm;
  top:{top}mm;
  right:{right}mm;
  bottom:{bottom}mm;
  border:0.25mm solid #000;
}
.sheet-number {
  position:absolute;
  z-index:3;
  right:5mm;
  top:5mm;
  width:10mm;
  height:10mm;
  background:#fff;
  border:0.25mm solid #000;
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
  font-size:9pt;
  line-height:1.3;
  min-height:4.3mm;
  white-space:pre-wrap;
  overflow-wrap:break-word;
  outline:none;
}
.dt .pz-cell:hover { background:#f4f8ff; box-shadow:inset 0 0 0 1px #cfe0f5; }
.dt .pz-cell:focus { background:#fff7d6; box-shadow:inset 0 0 0 1px #2a6; }
.dt .dt-cont { text-align:center; font-style:italic; color:#333; }
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
   Колонки обозначения и листов имеют фиксированную ширину и НЕ сжимаются,
   поэтому наименование обязано переноситься: без min-width:0 длинное слово
   (например, шифр или наименование без пробелов) распирало flex-строку, и
   правая часть содержания вместе с рамкой уходила за край листа. */
.toc-table { width:100%; border-collapse:collapse; }
.toc-line { display:flex; justify-content:space-between; align-items:center;
  border-bottom:0.25mm solid #999; padding:1.2mm 0; font-weight:normal; }
/* Шапка таблицы содержания — полужирная (как в DOCX: bold только у заголовков
   колонок). Строки содержания выводятся ОБЫЧНЫМ шрифтом — как в Word/PDF. */
.toc-head { border-bottom:0.4mm solid #000; font-weight:bold; font-size:9pt; }
.toc-line .toc-name { flex:1 1 auto; min-width:0; text-align:left;
  overflow-wrap:break-word; word-break:break-word; }
.toc-line .toc-des { flex:0 0 34mm; max-width:34mm; text-align:center;
  overflow-wrap:break-word; word-break:break-word; }
.toc-line .toc-sheets { flex:0 0 18mm; max-width:18mm; text-align:center;
  overflow-wrap:break-word; word-break:break-word; }
"""


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
    return css


# ---- Основная надпись ------------------------------------------------------
# Точная геометрия по ГОСТ Р 21.101-2020, приложение Ж.
#   Форма 3 (графика): 185 x 55 мм
#   Форма 5 (первый лист текста): 185 x 40 мм
#   Форма 6 (последующий лист текста): 185 x 15 мм
# Таблица изменений (графы 14-19): 6 колонок 10,10,10,10,15,10 = 65 мм
#   (Изм. | Кол.уч. | Лист | № док. | Подл. | Дата)
# Обозначение (графа 1) и правая узкая колонка «Лист / Листов» — остаток до 185 мм.

CHANGE_COLS = [10, 10, 10, 10, 15, 10]       # = 65 мм
CHANGE_LABELS = ["Изм.", "Кол.уч.", "Лист", "№ док.", "Подл.", "Дата"]

# Кегль надписей и значений граф «Стадия / Лист / Листов» в основной надписи
# (формы 3, 5, 6 ГОСТ). По требованию — единый размер 12 pt.
STAGE_CELL_PT = 12
# Кегль текста графы 9 (наименование/индекс организации) — как у «Стадия/Лист».
CELL9_PT = 12
SHEET_COL_W = 10.0                           # правая колонка «Лист / Листов»


def _change_rows(snapshot, header=True):
    """Строки таблицы изменений: заголовок и значения (2 строки)."""
    vals = "".join(
        f'<td style="width:{w}mm" class="val">{_field(f"ch_{i}", "", size=8)}</td>'
        for i, w in enumerate(CHANGE_COLS))
    hdr = "".join(f'<td style="width:{w}mm" class="lbl">{esc(l)}</td>'
                  for w, l in zip(CHANGE_COLS, CHANGE_LABELS))
    return hdr, vals


def _change_table(snapshot, row_h_mm=7.5):
    """Таблица изменений (графы 14-19): 2 строки заданной высоты."""
    hdr, vals = _change_rows(snapshot)
    return (f'<table style="table-layout:fixed;width:100%;height:100%;">'
            f'<tr style="height:{row_h_mm}mm">{hdr}</tr>'
            f'<tr style="height:{row_h_mm}mm">{vals}</tr>'
            f'</table>')


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
    return (f'<div style="position:absolute;left:{x0}mm;top:{y - cut}mm;'
            f'width:{x1 - x0}mm;height:0;border-top:{LN};"></div>')


def _vline(x, y0, y1, cut=0.0):
    """Вертикальная линия на x, от y0 до y1 (мм от верха штампа, вниз).

    cut — срез сверху (мм): координаты сдвигаются вверх; отрезок обрезается снизу
    до среза (y0 < cut) или отбрасывается целиком.
    """
    if y1 <= cut:
        return ""
    y0c = max(y0, cut)
    return (f'<div style="position:absolute;left:{x}mm;top:{y0c - cut}mm;'
            f'width:0;height:{y1 - y0c}mm;border-left:{LN};"></div>')


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
    # Корректный file-URI (прямые слэши) — QtWebEngine иначе не открывает картинку.
    try:
        from pathlib import Path
        src = Path(png_path).resolve().as_uri()
    except Exception:
        src = "file:///" + str(png_path).replace("\\", "/")
    return (f'<div class="cell9-logo"{attrs} style="{style}">'
            f'<img class="stamp-logo" src="{esc(src)}" alt="Логотип организации">'
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


def _cell_sig(x0, y0, w, h, key, role_label, signer, *, size=9, align="left", pad=0.0, cut=0.0):
    """Поле подписи в ячейке (x0,y0,w,h) — мм от верха штампа."""
    if y0 < cut:
        return ""
    justify = {"left": "flex-start", "center": "center", "right": "flex-end"}[align]
    style = (f'position:absolute;left:{x0}mm;top:{y0 - cut}mm;'
             f'width:{w}mm;height:{h}mm;'
             f'display:flex;align-items:center;justify-content:{justify};'
             f'font-size:{size}pt;'
             f'white-space:nowrap;overflow:hidden;')
    if pad:
        style += f'padding-left:{pad}mm;padding-right:{pad}mm;'
    sig = _sig_field(key, role_label, signer, size=size, compact=True)
    return f'<div class="fit-lbl" style="{style}">{sig}</div>'


def _form3_labels(snapshot) -> str:
    """Надписи/поля формы 3 ГОСТ (графический лист, 185x55). Значения — редактируемые поля."""
    req = snapshot["requisites"]
    roles = snapshot["roles"]
    L = []
    # ---- Левый блок (0..65): таблица изменений, 11 рядов по 5 мм ----
    # Заголовок таблицы изменений (ряд y=20..25), 6 колонок.
    chg_hdr = [("Изм.", 0, 10), ("Кол.уч", 10, 10), ("Лист", 20, 10),
               ("№док.", 30, 10), ("Подпись", 40, 15), ("Дата", 55, 10)]
    for text, x0, w in chg_hdr:
        L.append(_lbl(x0, 20.0, w, 5.0, text, size=8, align="left", cls="val", pad=1.0, fit=True))
    # Строка «Разраб.» (ряд y=25..30): должность | фамилия | ... | дата.
    L.append(_lbl(0.0, 25.0, 20.0, 5.0, "Разраб.", size=8, cls="lbl", pad=1.0, fit=True))
    L.append(_cell_sig(20.0, 25.0, 20.0, 5.0, "develop", "Разработал", roles.get("develop"), size=8, pad=1.0))
    L.append(_cell_field(55.0, 25.0, 10.0, 5.0, "ch_date", "12/25", size=8, pad=1.0))
    # Строка «Н.контр.» (ряд y=45..50).
    L.append(_lbl(0.0, 45.0, 20.0, 5.0, "Н.контр.", size=8, cls="lbl", pad=1.0, fit=True))
    L.append(_cell_sig(20.0, 45.0, 20.0, 5.0, "norm", "Нормоконтроль", roles.get("norm"), size=8, pad=1.0))
    L.append(_cell_field(55.0, 45.0, 10.0, 5.0, "ch_date2", "12/25", size=8, pad=1.0))
    # Строка «ГИП» (ряд y=50..55) — самая нижняя.
    L.append(_lbl(0.0, 50.0, 20.0, 5.0, "ГИП", size=8, cls="lbl", pad=1.0, fit=True))
    L.append(_cell_sig(20.0, 50.0, 20.0, 5.0, "approve", "Утвердил", roles.get("approve"), size=8, pad=1.0))
    L.append(_cell_field(55.0, 50.0, 10.0, 5.0, "ch_date3", "12/25", size=8, pad=1.0))

    # ---- Правый блок (65..185) ----
    # Графа обозначения (y=0..10) — вся ширина.
    L.append(_cell_field(65.0, 0.0, 120.0, 10.0, "volume_id", req.get("volume_id", ""),
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
    roles = snapshot["roles"]
    L = []
    # ---- Левый блок (0..65): таблица изменений + подписи, ряды по 5 мм ----
    # Заголовок таблицы изменений (ряд y=10..15), 6 колонок.
    chg_hdr = [("Изм.", 0, 10, 1.0), ("Кол.уч", 10, 10, 0.5), ("Лист", 20, 10, 0.5),
               ("№док.", 30, 10, 0.5), ("Подп.", 40, 15, 0.5), ("Дата", 55, 10, 0.5)]
    for text, x0, w, pad in chg_hdr:
        L.append(_lbl(x0, 10.0, w, 5.0, text, size=7, align="left", cls="val", pad=pad, fit=True))
    # Строка «Разраб.» (ряд y=15..20): должность | фамилия | ... | дата.
    L.append(_lbl(0.0, 15.0, 20.0, 5.0, "Разраб.", size=7, cls="lbl", pad=1.0, fit=True))
    L.append(_cell_sig(20.0, 15.0, 20.0, 5.0, "develop", "Разработал", roles.get("develop"), size=7, pad=1.0))
    L.append(_cell_field(55.0, 15.0, 10.0, 5.0, "ch_date", "12/25", size=7, pad=0.5))
    # Строка «Н.контр.» (ряд y=30..35).
    L.append(_lbl(0.0, 30.0, 20.0, 5.0, "Н.контр.", size=7, cls="lbl", pad=1.0, fit=True))
    L.append(_cell_sig(20.0, 30.0, 20.0, 5.0, "norm", "Нормоконтроль", roles.get("norm"), size=7, pad=1.0))
    L.append(_cell_field(55.0, 30.0, 10.0, 5.0, "ch_date2", "12/25", size=7, pad=0.5))
    # Строка «ГИП» (ряд y=30..35).
    L.append(_lbl(0.0, 35.0, 20.0, 5.0, "ГИП", size=7, cls="lbl", pad=1.0, fit=True))
    L.append(_cell_sig(20.0, 35.0, 20.0, 5.0, "approve", "Утвердил", roles.get("approve"), size=7, pad=1.0))
    L.append(_cell_field(55.0, 35.0, 10.0, 5.0, "ch_date3", "12/25", size=7, pad=0.5))

    # ---- Правый блок (65..185) ----
    # Обозначение документа (y=0..15) — вся ширина правого блока.
    L.append(_cell_field(65.0, 0.0, 120.0, 15.0, "volume_id", req.get("volume_id", ""),
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
    """Надписи/поля формы 6 (185x15). Значения — редактируемые поля."""
    req = snapshot["requisites"]
    L = []
    # ---- Левый блок (0..65): заголовок таблицы изменений (ряд y=0..5) ----
    chg = [("Изм.", 0, 10), ("Кол.уч", 10, 10), ("Лист", 20, 10),
           ("№док.", 30, 10), ("Подп.", 40, 15), ("Дата", 55, 10)]
    for text, x0, w in chg:
        L.append(_lbl(x0, 0.0, w, 5.0, text, size=7, align="left", cls="val", pad=0.5, fit=True))
    # ---- Правый блок (65..185) ----
    # Обозначение документа (зона 65..175, y=0..15) — по центру.
    L.append(_cell_field(65.0, 0.0, 110.0, 15.0, "volume_id", req.get("volume_id", ""),
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

    dev = _sig_field("develop", "Разработал", roles.get("develop"))
    chk = _sig_field("check", "Проверил", roles.get("check"))
    nrm = _sig_field("norm", "Нормоконтроль", roles.get("norm"))

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
    stage = _field("stage", req.get("stage", ""), size=12, placeholder="Стадия")
    city = _field("city", req.get("city", ""), size=12, placeholder="Город")
    year = _field("year", req.get("year", ""), size=12, placeholder="Год")
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
      <div style="font-size:12pt">Стадия {stage}</div>
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
    vid = _field("volume_id", req.get("volume_id", ""), size=12)
    stage = _field("stage", req.get("stage", ""), size=12)
    city = _field("city", req.get("city", ""), size=12)
    year = _field("year", req.get("year", ""), size=12)
    approve = _sig_field("approve", "Утвердил", roles.get("approve"))
    content = f"""
    <div class="cover">
      <div style="font-size:12pt; margin-top:8mm">{org}</div>
      <div style="height:20mm"></div>
      <div style="font-size:18pt">{obj}</div>
      <div style="height:8mm"></div>
      <div style="font-size:16pt">{volname}</div>
      <div style="height:12mm"></div>
      <div style="font-size:12pt">Обозначение: {vid}</div>
      <div style="height:12mm"></div>
      <div style="font-size:12pt">Стадия {stage}</div>
      <div style="height:8mm"></div>
      <div style="font-size:12pt">{city} — {year}</div>
      <div style="height:20mm"></div>
      <div style="text-align:left; font-size:11pt">Утвердил: {approve}</div>
    </div>
    """
    # Титульный лист — без основной надписи (штампа).
    return content, ""


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
            f'<div class="toc-line">'
            f'<span class="toc-name">{esc(name)}</span>'
            f'<span class="toc-des">{esc(des)}</span>'
            f'<span class="toc-sheets">{esc(sheets)}</span>'
            f'</div>')
    content = f"""
    <h1>СОДЕРЖАНИЕ</h1>
    <div style="height:8mm"></div>
    <div class="toc-table">
      <div class="toc-line toc-head">
        <span class="toc-name">Наименование</span>
        <span class="toc-des">Обозначение</span>
        <span class="toc-sheets">Листы</span>
      </div>
      {''.join(lines)}
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
    content = toc_html(items,
                       title_size=getattr(toc_fm, "content_header", 14),
                       body_size=getattr(toc_fm, "toc_row", 12),
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


def render_placeholder(label: str, *, subtitle: str = "", path: str = "",
                       kind: str = "") -> str:
    """Заглушка-лист для документа/приложения/раздела (не стандартный лист ГОСТ).

    Используется только когда содержимое не удалось получить (нет файла /
    неподдерживаемый формат). Показывает наименование и источник.
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
                 'Не удалось отобразить содержимое файла.</div>')
    return "".join(lines)


# --------------------------------------------------------------------------
# Рендер содержимого документов (DOCX -> текст/таблицы, PDF -> страницы-картинки)
# --------------------------------------------------------------------------

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

    number_html = (f'<div class="sheet-number" aria-label="Номер листа">{esc(sheet_no)}</div>'
                   if sheet_no else "")
    frame_html = "" if no_frame else '<div class="frame"></div>'
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


def _paginate_docx(content, path: str = "") -> list:
    """Разбивает содержимое DOCX на «страницы» (списки строк/таблиц).

    Каждая страница — список элементов ("line", {...}) / ("table_rows", [row...]).
    Оценка заполнения страницы ведётся по измеренной модели высот: для абзаца —
    число визуальных строк x высота строки, для строки таблицы — число строк
    самой высокой ячейки x высота строки ячейки. Бюджет считается консервативно
    (с запасом), чтобы контент не уходил под штамп и не обрезался overflow:hidden.

    Высота абзаца учитывает оформление Word (app.core.pz_format): межстрочный
    интервал, абзацные отступы до/после и сужение колонки абзацными отступами,
    поэтому разбивка совпадает с отрисовкой.
    """
    stream = []
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
            para_index += 1
            if not para.text.strip():
                continue
            is_heading = para.is_heading
            bold = is_heading or any(r.bold for r in para.runs)
            italic = any(r.italic for r in para.runs)
            align = {"LEFT": "left", "CENTER": "center", "RIGHT": "right",
                     "JUSTIFY": "justify"}.get(para.alignment, "justify")
            size = 13 if is_heading else 11
            fmt = _para_format(content, para_index)
            # Один абзац — один элемент потока (а не построчно): его высота
            # оценивается по числу визуальных строк с учётом ширины колонки.
            # para_index нужен для inline-редактирования текста ПЗ в предпросмотре.
            stream.append(("line", {"text": para.text, "bold": bold, "italic": italic,
                                    "align": align, "size": size,
                                    "para_index": para_index, "fmt": fmt,
                                    # Разрыв страницы, заданный в Word («с новой
                                    # страницы»): абзац начинает новый лист.
                                    "page_break": bool(fmt is not None
                                                       and fmt.page_break_before)}))
        elif kind == "table":
            stream.append(("table", value))

    # Доступная высота текстовой области: лист A4 - верхнее поле - нижнее поле,
    # уменьшенные на высоту штампа (форма 5 — 40 мм, форма 6 — 15 мм) и зазор.
    # Запас 2 мм сверху/снизу (border/отступы), чтобы оценка не занижала объём.
    first_budget = _page_text_height(FORM_HEIGHT[2])    # форма 5: штамп 40 мм
    next_budget = _page_text_height(FORM_HEIGHT["2a"])  # форма 6: штамп 15 мм
    pages, cur, cur_h, budget = [], [], 0.0, first_budget

    def flush_page() -> None:
        nonlocal cur, cur_h, budget
        if cur:
            pages.append(cur)
        cur = []
        cur_h = 0.0
        budget = next_budget

    for item in stream:
        if item[0] == "line":
            line = item[1]
            h = _estimate_paragraph_h(line["text"], line["size"], line.get("fmt"))
            # Разрыв страницы из Word закрывает лист даже тогда, когда место на
            # нём ещё есть (кроме случая, когда абзац и так начинает лист).
            if line.get("page_break") and cur:
                flush_page()
            elif cur and cur_h + h > budget:
                flush_page()
            cur.append(item)
            cur_h += h
            continue

        # Таблица разбивается по строкам: целая таблица не влезает на лист
        # даже при переполнении — иначе нижняя часть обрезается CSS-ом и
        # пользователь видит половину строки/таблицы. При переносе таблицы на
        # следующий лист первая строка повторяется как шапка (обычная практика
        # оформления таблиц по ГОСТ).
        table = item[1]
        rows = list(table.rows)
        ncols = _table_ncols(table)
        # Фактические ширины колонок из исходного DOCX (w:tblGrid): по ним
        # считаются и перенос текста в ячейке, и CSS-grid при отрисовке —
        # предпросмотр совпадает с исходной таблицей.
        grid = table_col_widths_mm(path, getattr(table, "table_index", -1), ncols)
        chars = _table_cell_chars_per_line(ncols)
        cell_chars = [_table_cell_chars_per_line(ncols, i, w)
                      for i, w in enumerate(grid)]
        header_row = rows[0] if len(rows) > 1 else None
        header_cost = (_estimate_table_row_h(header_row, chars, cell_chars)
                       if header_row is not None else 0.0)
        chunk = []
        chunk_h = 0.0
        for row in rows:
            cost = _estimate_table_row_h(row, chars, cell_chars)
            if chunk and cur_h + chunk_h + cost > budget:
                cur.append(("table_rows", chunk))
                cur_h += chunk_h
                flush_page()
                # Продолжение таблицы начинается с повторённой шапки.
                chunk = [header_row] if header_row is not None else []
                chunk_h = header_cost if header_row is not None else 0.0
            elif not chunk and cur and cur_h + cost > budget:
                flush_page()
                if header_row is not None:
                    chunk = [header_row]
                    chunk_h = header_cost
            chunk.append(row)
            chunk_h += cost
        if chunk:
            cur.append(("table_rows", chunk))
            cur_h += chunk_h
    if cur:
        pages.append(cur)
    if not pages:
        pages = [[]]
    return pages


def _continued_tables(pages: list) -> list:
    """Определяет, какие таблицы продолжаются на каждой странице DOCX.

    Пагинация повторяет шапку таблицы на новом листе, поэтому лист, первая
    строка которого повторяет последнюю строку предыдущего листа, — это
    продолжение. Возвращает список (по числу страниц) списков индексов таблиц,
    для которых на данном листе нужно вывести строку «Продолжение таблицы».
    """
    result = [[] for _ in pages]
    prev_rows = None
    for i, page in enumerate(pages):
        chunks = [item[1] for item in page if item[0] == "table_rows"]
        first_row = chunks[0][0] if chunks and chunks[0] else None
        if first_row is not None and prev_rows and prev_rows[-1] is not None:
            last_prev = prev_rows[-1][0] if prev_rows[-1] else None
            if last_prev is not None and last_prev is first_row:
                result[i].append(int(getattr(first_row[0], "table_index", -1)))
        prev_rows = chunks
    return result


def _grid_row_render(row, *,
                     editable: bool,
                     para_file: str,
                     overrides: dict | None = None,
                     cell_overrides: dict | None = None,
                     colspan: int = 0) -> str:
    """HTML одной строки таблицы DOCX. Ячейки адресуемы и редактируемы.

    Адрес ячейки — data-table / data-row / data-cell (индексы в document.tables),
    он же используется при записи правок в файл (apply_table_cell_edits).
    colspan > 1 нужен строке-заглушке «Продолжение таблицы» на новом листе.
    """
    overrides = overrides or {}
    cell_overrides = cell_overrides or {}
    if colspan:
        style = f"grid-column:1 / span {int(colspan)};"
        return (f'<div class="dt-row"><div class="dt-cell dt-cont" style="{style}">'
                f'Продолжение таблицы</div></div>')
    cells_html = []
    for c, cell in enumerate(row, start=1):
        text = getattr(cell, "text", None) or ""
        # Позиция задаётся явно (колонка), чтобы продолжения объединённых ячеек
        # не сдвигали соседние ячейки (CSS auto-placement).
        continuation = getattr(cell, "continuation", "")
        if continuation:
            # Ячейка внутри объединения: место занимает, но рамку не рисует и
            # текст не повторяет — его выводит начало объединения.
            # Вертикальное объединение не рисуется как grid-row:span: при
            # переносе таблицы на следующий лист такое объединение «уезжает»
            # вместе с листом, и продолжение осталось бы без текста. Пустая
            # ячейка столбец сохраняет, а текст показан один раз — в начале.
            cells_html.append(f'<div class="dt-cell dt-merged" style="grid-column:{c};"></div>')
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
    align = line.get("align", "justify")
    is_list = bool(fmt is not None and fmt.has_list)
    style = f"font-size:{size}pt;"
    if fmt is not None:
        style = paragraph_css(fmt, extra=f"font-size:{size}pt",
                              omit_text_indent=is_list)
        if para_align_is_default(fmt, align):
            style += f"text-align:{align};"
    else:
        style += f"text-align:{align};margin:0;line-height:{DEFAULT_LINE_HEIGHT:g};"
    if line.get("bold"):
        style += "font-weight:bold;"
    if line.get("italic"):
        style += "font-style:italic;"

    marker = ""
    if is_list:
        # Маркер выносится в отдельный бокс рядом с абзацем: он не входит в
        # редактируемый текст (пользователь правит только текст пункта, а
        # номер/маркер остаётся как в Word) и не удаляется клавишей Backspace.
        marker = (f'<span class="pz-mark" contenteditable="false" '
                  f'style="{marker_css(fmt)}">{esc(fmt.marker_text)}</span>')

    body = _pz_text(text)
    if editable:
        para = (f'<p class="pz-para" data-para-index="{idx}" '
                f'data-para-file="{esc(para_file)}" '
                f'contenteditable="true" style="{style}">{body}</p>')
    else:
        # В собранном документе (HTML/PDF) редактируемая разметка не выводится.
        para = f'<p style="{style}">{body}</p>'
    if is_list:
        return f'<div class="pz-item">{marker}{para}</div>'
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
                      cell_overrides: dict | None = None,
                      continued: list | None = None) -> list:
    """Возвращает HTML-строки для одной страницы DOCX.

    editable=True — абзацы и ячейки таблиц становятся редактируемыми прямо в
    предпросмотре (contenteditable с data-para-index / data-cell-addr). Правка
    уходит в Python через мост и записывается в исходный DOCX
    («Сохранить правки ПЗ»).

    overrides      — несохранённые правки абзацев ({индекс абзаца: текст});
    cell_overrides — несохранённые правки ячеек ({"т:с:я": текст}).
    Оба применяются поверх текста файла, чтобы правки не пропадали при
    перерисовке листов.

    continued — список индексов таблиц, продолжение которых начинается на этой
    странице (после разрыва таблицы) — над ними выводится строка-заглушка.
    """
    overrides = overrides or {}
    cell_overrides = cell_overrides or {}
    continued = list(continued or ())
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
            if table_index in continued:
                rows_html.append(_grid_row_render([], editable=False, para_file=para_file,
                                                  colspan=ncols))
            for row in chunk:
                rows_html.append(_grid_row_render(
                    row, editable=editable, para_file=para_file, overrides=overrides,
                    cell_overrides=cell_overrides))
            parts.append(f'<div class="dt" style="grid-template-columns:{colspec}">'
                         f'{"".join(rows_html)}</div>')
    return parts


def _pdf_page_image_b64(path: str, index: int, max_side: int = 1800) -> str:
    """Рендерит страницу PDF в PNG и возвращает base64 (для встраивания в HTML)."""
    import pymupdf
    with pymupdf.open(path) as doc:
        page = doc[index]
        rect = page.rect
        mm_w = rect.width * 25.4 / 72.0
        mm_h = rect.height * 25.4 / 72.0
        dpi = max_side / (max(mm_w, mm_h) / 25.4)
        dpi = int(max(60.0, min(200.0, dpi)))
        pix = page.get_pixmap(dpi=dpi)
        return base64.b64encode(pix.tobytes("png")).decode("ascii")


def _pdf_page_size_mm(path: str, index: int) -> tuple:
    """Размеры страницы PDF в мм."""
    import pymupdf
    with pymupdf.open(path) as doc:
        rect = doc[index].rect
        return round(rect.width * 25.4 / 72.0, 1), round(rect.height * 25.4 / 72.0, 1)


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
        # Номера листов в строках оглавления — сквозные по тому: сдвиг равен
        # (номер ПЕРВОГО листа оглавления − 1) + число листов оглавления, т.е.
        # номеру листа записки с локальным номером 1 минус 1. Сдвиг одинаков для
        # всех листов оглавления, поэтому считается по номеру первого листа.
        try:
            first_no = int(sheet_no) - i
        except (TypeError, ValueError):
            first_no = 1
        num["pz_toc_offset"] = max(0, first_no - 1) + len(pages)
        num["pz_toc_pages"] = len(pages)
        # Первый лист записки несёт большой штамп (форма 5), остальные — форму 6.
        num["toc_first"] = (i == 0)
        if i > 0:
            num["pz_toc_page"] = i + 1
        frag = _sheet_fragment("pz_toc", num, font_map, counter=None,
                               sheets_total=sheets_total)
        frags.append(frag)
    return frags


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
    except Exception:
        return toc_frags + [_sheet_fragment("doc", snapshot, font_map, counter=counter,
                                            sheets_total=sheets_total)]
    pages = _paginate_docx(content, path)
    if max_pages is not None:
        pages = pages[:max_pages]
    # Несохранённые правки ячеек таблиц этого документа ({"т:с:я": текст}).
    cell_overrides = (snapshot.get("pz_cell_edits") or {}).get(path) or {}
    # Какие таблицы продолжаются на каждом листе: если первая строка листа
    # совпадает с последней строкой предыдущего листа, значит таблица была
    # разорвана и её шапка повторена — над ней выводится «Продолжение таблицы».
    continued = _continued_tables(pages)
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
            cell_overrides=cell_overrides,
            continued=continued[i] if i < len(continued) else None))
        # Отступ сверху с учётом квадрата сквозной нумерации (indent_top),
        # чтобы текст не заезжал под квадрат в правом верхнем углу.
        frags.append(_wrap_sheet(body, stamp, sheet_no=num.get("sheet_no", ""),
                                 indent_top=True))
    if not frags:
        frags = [_sheet_fragment("doc", snapshot, font_map, counter=counter,
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
    try:
        with pymupdf.open(path) as pdf:
            count = pdf.page_count
    except Exception:
        return [_sheet_fragment("appendix" if is_appendix else "doc", snapshot, font_map,
                                counter=counter, sheets_total=sheets_total)]
    if count <= 0:
        return [_sheet_fragment("appendix" if is_appendix else "doc", snapshot, font_map,
                                counter=counter, sheets_total=sheets_total)]
    if max_pages is not None:
        count = min(count, max_pages)

    frags = []
    # Правило формата: только ГРАФИЧЕСКАЯ ЧАСТЬ сохраняет родной размер листа
    # из вложенного PDF. Все остальные листы (приложения, ПЗ-PDF) — A4.
    section = snapshot.get("section", "")
    is_graphic = (not is_appendix and section != "Пояснительная записка")
    for i in range(count):
        # Номер листа в пределах раздела: из счётчика (или i+1).
        if counter is not None:
            sheet_no = counter.next()
        else:
            sheet_no = i + 1
        num = _numbered_snapshot(snapshot, sheet_no, sheets_total)
        b64 = _pdf_page_image_b64(path, i)
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
        img = (f'<img src="data:image/png;base64,{b64}" '
               f'style="width:100%;height:100%;object-fit:contain;"/>')
        if is_graphic:
            # Графическая часть: НЕ рисуем новую рамку и штамп — сохраняем
            # ссылку на исходную страницу PDF. При экспорте она вставляется
            # напрямую из PDF (вектор, родной размер), а растровая base64-
            # картинка здесь служит только для предпросмотра. Поверх листа
            # накладывается только квадрат сквозной нумерации (no_frame=True),
            # привязанный к верхнему правому углу с отступами 5 мм.
            # full_bleed=True: содержимое занимает весь лист, чтобы чертёж
            # сохранял родной размер и его рамка совмещалась с рамкой листа,
            # а квадрат номера — с верхним правым углом рамки чертежа.
            frags.append(_wrap_sheet(header + img, stamp, w_mm=w_mm, h_mm=h_mm,
                                     data_pdf=(path, i), sheet_no=num.get("sheet_no", ""),
                                     no_frame=True, full_bleed=True))
        else:
            # Текстовые листы (ПЗ) и приложения: отступ сверху с учётом
            # квадрата сквозной нумерации, чтобы текст не заезжал под него.
            frags.append(_wrap_sheet(header + img, stamp,
                                     sheet_no=num.get("sheet_no", ""),
                                     indent_top=True))
    return frags


def _doc_sheet_fragments(d, snapshot: dict, font_map: FontMap,
                         first: bool = True, max_pages: int | None = None,
                         counter: _SheetCounter | None = None,
                         sheets_total: int | None = None) -> list:
    """Листы документа (d — SourceFile): DOCX -> текст, PDF -> страницы-картинки."""
    path = d.path
    if not path or not os.path.exists(path):
        return [_sheet_fragment("doc", snapshot, font_map, counter=counter,
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
        return [_sheet_fragment("appendix", snapshot, font_map, counter=counter,
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
    """Копия snapshot с номером листа и общим числом листов раздела."""
    s = dict(snapshot)
    s["sheet_no"] = str(sheet_no)
    if sheets_total is not None:
        s["sheets_total"] = str(sheets_total)
    return s


def _pdf_count(path: str) -> int:
    """Число страниц PDF (для подсчёта листов раздела)."""
    import pymupdf
    try:
        with pymupdf.open(path) as pdf:
            return max(1, pdf.page_count)
    except Exception:
        return 1


def _doc_sheet_count(d) -> int:
    """Число листов одного документа (DOCX — пагинация, PDF — страницы)."""
    path = d.path
    if not path or not os.path.exists(path):
        return 1          # будет показан лист-заглушка
    if _is_docx(path):
        try:
            from .docx_reader import read_docx
            # Путь передаётся обязательно: от него зависят фактические ширины
            # колонок таблиц, а значит и разбивка на листы. Иначе число листов
            # в содержании/штампе не совпало бы с предпросмотром.
            return max(1, len(_paginate_docx(read_docx(path), path)))
        except Exception:
            return max(1, int(getattr(d, "pages", 0) or 1))
    return _pdf_count(path)


def _section_sheet_count(project, section_kind: str) -> int:
    """Общее число листов раздела (для графы «Листов» в штампе)."""
    from ..models import (SECTION_COVER, SECTION_TITLE, SECTION_CONTENT,
                          SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX,
                          SECTION_CHANGES, DocKind)
    if section_kind == SECTION_COVER:
        return 1 if project.structure_needed(SECTION_COVER) else 0
    if section_kind == SECTION_TITLE:
        return 1 if project.structure_needed(SECTION_TITLE) else 0
    if section_kind == SECTION_CONTENT:
        return 1 if project.structure_needed(SECTION_CONTENT) else 0
    if section_kind == SECTION_CHANGES:
        if not project.structure_needed(SECTION_CHANGES) or not project.changes:
            return 0
        return 1          # лист регистрации изменений формируется одним листом
    if section_kind == SECTION_PZ:
        # + листы автоматического оглавления записки (app.core.pz_toc).
        from .pz_toc import project_toc_sheets
        return (project_toc_sheets(project)
                + sum(_doc_sheet_count(d) for d in project.documents_for(DocKind.POYASNITELNAYA)))
    if section_kind == SECTION_GRAPHIC:
        return sum(_doc_sheet_count(d) for d in project.documents_for(DocKind.GRAPHIC))
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
    return _wrap_sheet(content, stamp, sheet_no=num_snapshot.get("sheet_no", ""),
                       indent_top=_has_no)


def _section_fragments(project, section_kind: str, snapshot: dict,
                       font_map: FontMap, counter: _SheetCounter | None = None,
                       sheets_total: int | None = None) -> list:
    """Все листы раздела section_kind, включая многостраничные документы.

    Нумерация листов (графы «Лист» и «Листов») ведётся сквозная по всему тому.
    При standalone-сборке раздела счётчик локальный; обложка и титульный лист
    не получают номер и не увеличивают счётчик.
    """
    from ..models import (SECTION_COVER, SECTION_TITLE, SECTION_CONTENT,
                          SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX,
                          SECTION_CHANGES, DocKind)
    frags = []
    # Для standalone-раздела сохраняем локальную нумерацию; для тома
    # вызывающий код передаёт единый счётчик и общее число листов.
    counter = counter or _SheetCounter(1)
    sheets_total = sheets_total if sheets_total is not None else _section_sheet_count(project, section_kind)
    if section_kind == SECTION_COVER:
        frags.append(_sheet_fragment("cover", snapshot, font_map,
                                     counter=counter, sheets_total=sheets_total))
    elif section_kind == SECTION_TITLE:
        frags.append(_sheet_fragment("title", snapshot, font_map,
                                     counter=counter, sheets_total=sheets_total))
    elif section_kind == SECTION_CONTENT:
        frags.append(_sheet_fragment("content", snapshot, font_map,
                                     counter=counter, sheets_total=sheets_total))
    elif section_kind == SECTION_CHANGES:
        if project.changes:
            frags.append(_sheet_fragment("changes", snapshot, font_map,
                                         counter=counter, sheets_total=sheets_total))
    elif section_kind == SECTION_PZ:
        docs = project.documents_for(DocKind.POYASNITELNAYA)
        if not docs:
            sub = dict(snapshot)
            sub.update({"section": SECTION_PZ, "section_label": "Пояснительная записка"})
            frags.append(_sheet_fragment("doc", sub, font_map,
                                         counter=counter, sheets_total=sheets_total))
        first = True
        for d in docs:
            sub = dict(snapshot)
            sub.update({"path": d.path, "section": SECTION_PZ,
                        "section_label": "Пояснительная записка",
                        "label": os.path.basename(d.path) if d.path else "—"})
            dfrags = _doc_sheet_fragments(d, sub, font_map, first=first,
                                          counter=counter, sheets_total=sheets_total)
            frags.extend(dfrags)
            if dfrags:
                first = False
    elif section_kind == SECTION_GRAPHIC:
        for d in project.documents_for(DocKind.GRAPHIC):
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
    return _preview_document([fragment])


def render_sheet_fragments(kind: str, snapshot: dict,
                           font_map: FontMap | None = None,
                           project=None) -> list:
    """Возвращает разметку ВСЕХ листов выбранного пункта дерева.

    Для стандартных листов (обложка/титул/содержание/изменения) и отдельных
    форм — один лист. Для раздела тома (ПЗ / графика / приложения) и для
    многостраничного документа — все его листы подряд, с той же сквозной
    нумерацией, что и при сборке тома.
    """
    from ..models import (SECTION_COVER, SECTION_TITLE, SECTION_CONTENT,
                          SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX,
                          SECTION_CHANGES)
    font_map = font_map or FontMap()

    # Лист/документ с явно заданным путём — отдаём все листы этого документа.
    if kind in ("doc", "appendix"):
        return _all_document_fragments(kind, snapshot, font_map)

    # Раздел тома, для которого нужен полный набор листов.
    section = snapshot.get("section", "")
    if kind == "section" and section in (SECTION_PZ, SECTION_GRAPHIC,
                                         SECTION_APPENDIX):
        return _all_section_fragments(section, snapshot, font_map, project)

    # Обложка/титул/содержание/изменения и отдельные формы ГОСТ — один лист.
    if kind in ("cover", "title", "content", "changes",
                "stamp_f1", "stamp_f2", "stamp_f2a"):
        return [_sheet_fragment(kind, snapshot, font_map)]

    # Прочие документы/разделы — один лист (заглушка или первая страница).
    return [_sheet_fragment(kind, snapshot, font_map)]


def _all_document_fragments(kind: str, snapshot: dict, font_map: FontMap) -> list:
    """Все листы документа/приложения из снимка предпросмотра (path в snapshot)."""
    path = snapshot.get("path", "")
    if not path or not os.path.exists(path):
        return [_sheet_fragment(kind, snapshot, font_map)]
    if kind == "appendix":
        return _pdf_sheet_fragments(path, snapshot, font_map, first=True,
                                    is_appendix=True)
    if _is_docx(path):
        return _docx_sheet_fragments(path, snapshot, font_map, first=True)
    return _pdf_sheet_fragments(path, snapshot, font_map, first=True,
                                is_appendix=False)


def _all_section_fragments(section_kind: str, snapshot: dict,
                           font_map: FontMap, project=None) -> list:
    """Все листы раздела тома (ПЗ / графика / приложения) для предпросмотра.

    Сквозная нумерация ведётся внутри раздела, как при сборке отдельного
    раздела (build_section_html), — номера совпадают с «Собрать раздел».
    """
    if project is None:
        # Без проекта доступен только первый лист раздела.
        return [_sheet_fragment("section", snapshot, font_map)]
    counter = _SheetCounter(1)
    frags = _section_fragments(project, section_kind, snapshot, font_map,
                               counter=counter)
    return frags or [_sheet_fragment("section", snapshot, font_map)]


def _preview_document(fragments: list) -> str:
    """Полный HTML-документ интерактивного предпросмотра (один или много листов)."""
    fragment = "\n".join(fragments)
    return f"""<!doctype html>
<html><head><meta charset="utf-8">
<script src="qrc:///qtwebchannel/qwebchannel.js"></script>
<style>{_css()}</style></head>
<body>{fragment}
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

  function init() {{
    // Автоподгонка подписей (до создания моста, т.к. мост может быть недоступен
    // в оффскрин-рендере; в приложении повторно вызовется после onSheetReady).
    fitLabels();
    if (typeof QWebChannel === 'undefined' || typeof qt === 'undefined') {{ return; }}
    new QWebChannel(qt.webChannelTransport, function(channel) {{
      window.bridge = channel.objects.bridge;
      if (!window.bridge) {{ return; }}
      // Редактируемые поля (contenteditable)
      var fields = document.querySelectorAll('.field[contenteditable="true"]');
      fields.forEach(function(f) {{
        f.addEventListener('input', function(e) {{
          var key = f.getAttribute('data-key');
          var val = f.innerText;
          if (window.bridge.onFieldChanged) {{
            window.bridge.onFieldChanged(key, val);
          }}
        }});
      }});
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
      // Абзацы пояснительной записки: правка прямо в предпросмотре.
      // Индекс абзаца в DOCX приходит в data-para-index; по нему Python
      // записывает текст в исходный файл (кнопка «Сохранить в файл»).
      var paras = document.querySelectorAll('.pz-para[contenteditable="true"]');
      paras.forEach(function(p) {{
        p.addEventListener('input', function(e) {{
          var idx = parseInt(p.getAttribute('data-para-index'), 10);
          if (isNaN(idx) || idx < 0) {{ return; }}
          var file = p.getAttribute('data-para-file') || '';
          if (window.bridge.onParagraphChanged) {{
            window.bridge.onParagraphChanged(idx, file, p.innerText);
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
    fragments = []
    # Один счётчик на весь том: обложка и титульный лист его не расходуют.
    counter = _SheetCounter(1)
    sheets_total = sum(_section_sheet_count(project, e.get("kind"))
                       for e in project.structure if e.get("needed", True)
                       and e.get("kind") not in ("cover", "title"))
    # Обходим структуру тома и для каждого раздела добавляем ВСЕ его листы
    # (включая многостраничные документы и приложения).
    for entry in project.structure:
        if not entry.get("needed", True):
            continue
        kind = entry.get("kind")
        fragments.extend(_section_fragments(project, kind, snapshot, font_map,
                                            counter=counter, sheets_total=sheets_total))
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
    fragments = _section_fragments(project, section_kind, snapshot, font_map)
    title = SECTION_LABELS.get(section_kind, section_kind)
    html = _html_document(fragments, title)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path
