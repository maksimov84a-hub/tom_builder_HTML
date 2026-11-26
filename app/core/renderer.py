# -*- coding: utf-8 -*-
"""
Оформляющий слой ГОСТ: отрисовка рамки и основной надписи на PDF-странице
через PyMuPDF (векторная отрисовка). Шрифт ISOCPEUR, кегль по карте.

Геометрия основной надписи (ГОСТ Р 21.101-2020, приложение Ж):
  форма 3  (лист графической части)                 — 185 x 55 мм;
  форма 5  (первый/заглавный лист текстового)       — 185 x 40 мм;
  форма 6  (последующие листы текстового)           — 185 x 15 мм.

Внутренние ключи формы (см. gost.FORM_*): 1 -> форма 3, 2 -> форма 5,
"2a" -> форма 6.

Раскладка строится в миллиметрах от левого-верхнего угла штампа, поэтому
не зависит от ориентации/размера листа.
"""
from __future__ import annotations

import os

import pymupdf

from .gost import Frame, FORM_HEIGHT, STAMP_WIDTH_MM, \
    FORM_GRAPHIC, FORM_TEXT_FIRST, FORM_TEXT_NEXT
from .geometry import STAGE_CELL_KEYS, STAGE_CELL_PT, CELL9_KEYS, CELL9_PT
from ..config import FontMap
from .fonts import register_font

LINE_WIDTH_PT = 0.7        # ~0.25 мм
BLACK = (0, 0, 0)

# перевод мм -> pt (для координат PyMuPDF)
_PT = 25.4 / 72.0


def mm_to_pt(mm: float) -> float:
    return mm / _PT


def _font_for(key: str, font_map: FontMap, role: str = "") -> float:
    """Кегль по ключу и роли ячейки (карта расположения).

    Графы «Стадия / Лист / Листов» (ключи stage/sheet/qty) набираются
    единым кеглем STAGE_CELL_PT (12 pt) во всех формах (3, 5, 6).
    Ключ "sheet" встречается также в таблице изменений (графа 16) — там
    кегль остаётся прежним, поэтому учитывается роль ячейки ("value").
    """
    if key in STAGE_CELL_KEYS and role != "value":
        return float(STAGE_CELL_PT)
    # Графа 9 (наименование/индекс организации) — кегль как у «Стадия/Лист».
    if key in CELL9_KEYS:
        return float(CELL9_PT)
    m = {
        "volume": font_map.stamp_volume,
        "docno": font_map.stamp_columns,
        "sign": font_map.stamp_columns,
        "date": font_map.stamp_columns,
        "ch": font_map.stamp_columns,
        "organization": font_map.stamp_columns,
        "object": font_map.stamp_columns,
    }
    return float(m.get(key, font_map.stamp_columns)) if key else font_map.stamp_columns


# ---- Вертикальное центрирование текста в ячейке ----------------------------
# insert_textbox размещает первую строку от верхней границы прямоугольника,
# из-за чего текст «прилипал» к верху ячейки. Ниже прямоугольник сдвигается
# так, чтобы видимый центр текстового блока совпал с центром ячейки.

#: кэш измеренных габаритов текста: (текст, кегль, ширина) ->
#: (смещение верха текста от верха прямоугольника, высота текста), в pt.
_TEXT_BOX_CACHE: dict = {}


def _text_box(text: str, size_pt: float, width_pt: float) -> tuple | None:
    """Измеряет фактические габариты текста при заданной ширине ячейки.

    Возвращает (top_offset, height) в pt относительно верхней границы
    прямоугольника. Замер делается пробной вставкой в прямоугольник ТОЙ ЖЕ
    ширины (перенос строк получается таким же, как в реальной ячейке) с
    заведомо большой высотой. Результат кэшируется по (текст, кегль, ширина).
    """
    key = (text, round(float(size_pt), 2), round(float(width_pt), 2))
    if key in _TEXT_BOX_CACHE:
        return _TEXT_BOX_CACHE[key]

    result = None
    try:
        tmp = pymupdf.open()
        pg = tmp.new_page(width=595, height=842)
        fname = register_font(pg) or "helv"
        rect = pymupdf.Rect(50.0, 100.0, 50.0 + float(width_pt), 5000.0)
        rc = pg.insert_textbox(rect, text, fontname=fname,
                               fontsize=float(size_pt),
                               align=pymupdf.TEXT_ALIGN_CENTER)
        if rc >= 0:
            spans = []
            for b in pg.get_text("dict", clip=rect)["blocks"]:
                if b["type"] != 0:
                    continue
                for ln in b["lines"]:
                    for s in ln["spans"]:
                        if s["text"].strip():
                            spans.append(s["bbox"])
            if spans:
                top = min(s[1] for s in spans) - rect.y0
                bottom = max(s[3] for s in spans) - rect.y0
                result = (float(top), float(bottom - top))
        tmp.close()
    except Exception:
        result = None
    _TEXT_BOX_CACHE[key] = result
    return result


def _vcenter_rect(rect, text: str, size_pt: float):
    """Прямоугольник, сдвинутый так, чтобы текст был по центру ячейки.

    Высота прямоугольника сохраняется (вписывание текста не меняется),
    меняется только вертикальное положение: insert_textbox ставит первую
    строку от верхней границы, поэтому прямоугольник сдвигается вниз так,
    чтобы видимый центр текста совпал с центром ячейки.
    """
    box = _text_box(text, size_pt, rect.width)
    if not box:
        return rect
    top_off, text_h = box
    cell_cy = (rect.y0 + rect.y1) / 2.0
    new_y0 = max(rect.y0, cell_cy - (top_off + text_h / 2.0))
    return pymupdf.Rect(rect.x0, new_y0, rect.x1, new_y0 + rect.height)


# ---- Таблица изменений (графы 14-19) ---------------------------------------
# 5 колонок: Изм | Лист | № докум. | Подп. | Дата  (сумма 65 мм по ГОСТ)
CELL9_PAD_MM = 0.4        # внутренний отступ логотипа в ячейке графы 9

CHANGE_COLS = [10.0, 10.0, 15.0, 15.0, 15.0]
CHANGE_KEYS = ["ch", "sheet", "docno", "sign", "date"]
CHANGE_LABELS = ["Изм", "Лист", "№ докум.", "Подп.", "Дата"]
SHEET_COL_W = 10.0        # правая узкая колонка «Лист / Листов» (графы 7, 8)


def _header_value_h(form) -> tuple:
    """Высоты строки заголовка и строки значений таблицы изменений (мм).

    Заголовок таблицы изменений стоит над первой строкой подписей: у формы 5
    это y 10..15, у формы 3 — y 20..25 (см. FORM5_HLINES/FORM3_HLINES).
    """
    return (10.0, 5.0) if form != FORM_GRAPHIC else (20.0, 5.0)


# ---- Сетка правого блока штампа (мм от левого-верхнего угла) ----------------
# Согласовано с эталонами (Форма 3.pdf / Форма 5.pdf) и с HTML-предпросмотром
# (preview_html.FORM5_VLINES / FORM5_HLINES):
#   обозначение (графа 1)  — 65..135 на всю высоту штампа;
#   Стадия | Лист | Листов — 135..150 | 150..165 | 165..185, y 15..25;
#   графа 9 (организация)  — 135..185, y 25..40 (под строкой Стадия/Лист/Листов).
# Раньше правый блок делился по 175 мм и «наезжал» на ячейку графы 9 — теперь
# границы совпадают с сеткой ГОСТ.
DESIGN_X0 = 65.0
DESIGN_X1 = 135.0
STAGE_COLS = ((135.0, 15.0), (150.0, 15.0), (165.0, 20.0))   # (x, ширина) x3


def _layout_form6(stamp_h: float) -> list:
    """Раскладка формы 6 (185 x 15) по сетке app.core.geometry.

    Левый блок (0..65) — таблица изменений из ШЕСТИ граф 10/10/10/10/15/10
    (``geometry.CHANGE_COLS``): два ряда значений (0..5 и 5..10) и подписи в
    нижней полосе 10..15. Раньше у PDF-рендера была своя разбивка на пять граф
    (10/10/15/15/15), и сетка расходилась с Word и HTML-предпросмотром.

    «Лист» (175..185): заголовок 0..7, значение 7..15 — как на образце
    «Форма 6.pdf»; линия на 7 мм проходит только в этой графе.
    """
    from .geometry import CHANGE_COLS as F6_COLS, CHANGE_LABELS as F6_LABELS

    cells = []
    # Два ряда значений таблицы изменений.
    x = 0.0
    for width in F6_COLS:
        for y in (0.0, 5.0):
            cells.append({"x": x, "y": y, "w": width, "h": 5.0,
                          "text": "", "key": "", "role": "value"})
        x += width
    # Подписи — нижняя полоса.
    x = 0.0
    for width, label in zip(F6_COLS, F6_LABELS):
        cells.append({"x": x, "y": 10.0, "w": width, "h": 5.0,
                      "text": label, "key": "", "role": "header"})
        x += width
    # Обозначение документа.
    cells.append({"x": DESIGN_X0, "y": 0.0, "w": 175.0 - DESIGN_X0, "h": stamp_h,
                  "text": "", "key": "volume", "role": "design"})
    # Колонка «Лист».
    cells.append({"x": 175.0, "y": 0.0, "w": 10.0, "h": 7.0,
                  "text": "Лист", "key": "", "role": "header"})
    cells.append({"x": 175.0, "y": 7.0, "w": 10.0, "h": stamp_h - 7.0,
                  "text": "", "key": "sheet", "role": "sheet"})
    return cells


def _layout(form) -> list:
    """Раскладка основной надписи (ячейки в мм от левого-верхнего угла штампа).

    Общая схема:
      слева — таблица изменений (заголовок + строка значений);
      по центру — ячейка обозначения документа (графа 1, 65..135);
      справа — у форм 5/3 блок «Стадия/Лист/Листов» (135..185, y 15..25) и
      графа 9 «организация» (135..185, y 25..40); у формы 6 — узкая колонка
      «Лист» (175..185).

    Сетка согласована с эталонными образцами и HTML-предпросмотром.
    """
    stamp_h = FORM_HEIGHT.get(form, 40.0)
    if form == FORM_TEXT_NEXT:
        return _layout_form6(stamp_h)
    # header_y — верх строки заголовка таблицы изменений (мм от верха штампа).
    header_y, header_h = _header_value_h(form)
    change_w = sum(CHANGE_COLS)

    cells = []

    # --- Таблица изменений: только заголовок. Строка значений совпадает с
    # первой строкой подписей («Разраб.») — так на образцах форм 5 и 3.
    # У формы 6 своя раскладка — см. _layout_form6.
    hdr_y = header_y
    hdr_h = 5.0
    x = 0.0
    for wd, lb in zip(CHANGE_COLS, CHANGE_LABELS):
        cells.append({"x": x, "y": hdr_y, "w": wd, "h": hdr_h,
                      "text": lb, "key": "", "role": "header"})
        x += wd

    # --- Обозначение (графа 1): 65..135 на всю высоту штампа ---
    cells.append({"x": DESIGN_X0, "y": 0.0, "w": DESIGN_X1 - DESIGN_X0, "h": stamp_h,
                  "text": "", "key": "volume", "role": "design"})

    # --- Строки подписей левого блока (должность | фамилия | подпись | дата) ---
    # Раскладка строк — единый источник geometry.stamp_signature_rows: те же
    # строки и координаты, что в Word-таблице и HTML-предпросмотре.
    # Колонки по сетке: 0..20 должность, 20..40 фамилия, 40..55 графа «Подп.»,
    # 55..65 дата.
    from .geometry import stamp_signature_rows
    form_num = FORM_GRAPHIC if form == FORM_GRAPHIC else FORM_TEXT_FIRST
    for _r0, _r1, y, _role, index in stamp_signature_rows(form_num):
        cells.append({"x": 0.0, "y": y, "w": 20.0, "h": 5.0,
                      "text": "", "key": f"sigpost{index}", "role": "sig_post"})
        cells.append({"x": 20.0, "y": y, "w": 20.0, "h": 5.0,
                      "text": "", "key": f"sigsurname{index}", "role": "sig_name"})
        cells.append({"x": 40.0, "y": y, "w": 15.0, "h": 5.0,
                      "text": "", "key": f"sigimage{index}", "role": "sig_sign"})
        cells.append({"x": 55.0, "y": y, "w": 10.0, "h": 5.0,
                      "text": "", "key": f"sigdate{index}", "role": "sig_date"})

    # Вертикальные границы блоков зависят от формы (см. FORM3_HLINES/FORM5_HLINES
    # в preview_html и сетки в geometry):
    #   форма 5 (40 мм): Стадия/Лист/Листов — y 15..25, графа 9 — y 25..40;
    #   форма 3 (55 мм): Стадия/Лист/Листов — y 25..40, графа 9 — y 40..55.
    stage_y, cell9_y = (25.0, 40.0) if form == FORM_GRAPHIC else (15.0, 25.0)

    # --- Стадия | Лист | Листов (заголовок и значение — одна ячейка) ---
    for key, (cx, cw) in zip(("stage", "sheet", "qty"), STAGE_COLS):
        cells.append({"x": cx, "y": stage_y, "w": cw, "h": 10.0,
                      "text": "", "key": key, "role": "field"})

    # --- Графа 9 (наименование/различительный индекс организации) ---
    # 135..185 на высоту 15 мм под строкой «Стадия/Лист/Листов».
    # В форме 6 графы 9 нет (по ГОСТ), поэтому ячейка не создаётся.
    cells.append({"x": 135.0, "y": cell9_y, "w": 50.0, "h": 15.0,
                  "text": "", "key": "organization", "role": "field"})

    # --- Наименование (графа 5/2) в левой части под обозначением ---
    name_y = cell9_y
    cells.append({"x": DESIGN_X0, "y": name_y, "w": DESIGN_X1 - DESIGN_X0,
                  "h": max(0.0, stamp_h - name_y),
                  "text": "", "key": "object", "role": "design"})

    return cells


# Кегль цифры в квадрате сквозной нумерации (pt). Подобран так, чтобы
# видимая высота цифры составляла ~2.5-3 мм на листе (см. измерение:
# 12pt -> видимая высота цифры ~2.71 мм).
SHEET_NUMBER_FONT_PT = 12.0

# Коэффициенты видимых размеров цифры ISOCPEUR (мм на 1 pt кегля).
# Измерено рендером: для кегля 12pt видимая высота цифры = 2.709 мм,
# ширина = 1.397 мм. Это меньше полного кегля (высота строки = 12pt),
# т.к. видимая цифра занимает лишь часть межстрочного интервала.
_SHEET_NUM_DIGIT_H_PER_PT = 2.709 / 12.0    # ~0.2258 мм/pt
_SHEET_NUM_DIGIT_W_PER_PT = 1.397 / 12.0    # ~0.1164 мм/pt


def draw_sheet_number_box(page, page_w_mm: float, page_h_mm: float,
                          sheet_no: object, box_mm: float = 10.0) -> None:
    """Накладывает квадрат со сквозным номером листа на исходную страницу PDF.

    Верхний правый угол квадрата расположен на (-5; -5) мм относительно
    верхнего правого угла листа: на 5 мм левее и на 5 мм ниже границы листа.
    Координаты отсчитываются от левого верхнего угла страницы PyMuPDF.

    Цифра выравнивается по центру квадрата (по горизонтали и вертикали) и
    набирается кеглем SHEET_NUMBER_FONT_PT, чтобы на листе читалась высотой
    ~2.5-3 мм.
    """
    sheet_no = str(sheet_no or "")
    if not sheet_no:
        return
    x0 = page_w_mm - 5.0 - box_mm
    y0 = 5.0
    page.draw_rect((mm_to_pt(x0), mm_to_pt(y0),
                    mm_to_pt(x0 + box_mm), mm_to_pt(y0 + box_mm)),
                   color=BLACK, width=LINE_WIDTH_PT)
    # Цифра должна быть отцентрирована по квадрату (и по горизонтали, и по
    # вертикали) и набрана кеглем SHEET_NUMBER_FONT_PT, чтобы на листе
    # читалась высотой ~2.5-3 мм.
    fs = SHEET_NUMBER_FONT_PT
    fname = register_font(page) or "helv"
    # Видимые размеры цифры ISOCPEUR (мм) при данном кегле — измерены рендером.
    digit_h_mm = fs * _SHEET_NUM_DIGIT_H_PER_PT
    digit_w_mm = fs * _SHEET_NUM_DIGIT_W_PER_PT
    # Горизонтальное центрирование: левая точка базовой линии так, чтобы
    # центр видимой ширины цифры совпал с центром квадрата.
    tx = x0 + (box_mm - digit_w_mm) / 2.0
    # Вертикальное центрирование: видимая цифра опирается на базовую линию,
    # её верх = baseline - digit_h. Базовая линия ставится так, чтобы центр
    # видимой высоты цифры совпал с центром квадрата (y0 + box_mm/2).
    baseline_y = y0 + (box_mm + digit_h_mm) / 2.0
    page.insert_text((mm_to_pt(tx), mm_to_pt(baseline_y)), sheet_no,
                     fontname=fname, fontsize=fs, color=BLACK)


def _draw_cell9_logo(page, sx_mm: float, sy_mm: float, cell: dict,
                     texts: dict) -> bool:
    """Рисует логотип организации в графе 9 основной надписи.

    Логотип вписывается в ячейку с сохранением пропорций (автомасштабирование
    под ячейку). Возвращает True, если логотип нарисован.
    """
    path = texts.get("cell9_logo") or ""
    if not path:
        return False
    try:
        from .logo import render_logo_png, logo_pixel_size, fit_into
        png = render_logo_png(path)
        if not png:
            return False
        px_w, px_h = logo_pixel_size(path)
    except Exception:
        return False
    try:
        pad = max(0.0, float(texts.get("cell9_pad", CELL9_PAD_MM) or 0.0))
    except (TypeError, ValueError):
        pad = CELL9_PAD_MM
    fit_w, fit_h = fit_into(px_w, px_h,
                            float(cell["w"]) - 2 * pad, float(cell["h"]) - 2 * pad)
    if fit_w <= 0 or fit_h <= 0:
        return False
    # Центрируем вписанный логотип в ячейке (и по горизонтали, и по вертикали).
    x = mm_to_pt(sx_mm + cell["x"] + (cell["w"] - fit_w) / 2.0)
    y = mm_to_pt(sy_mm + cell["y"] + (cell["h"] - fit_h) / 2.0)
    try:
        page.insert_image(pymupdf.Rect(x, y, x + mm_to_pt(fit_w), y + mm_to_pt(fit_h)),
                          filename=png, keep_proportion=True)
        return True
    except Exception:
        return False


def draw_frame_and_stamp(page, page_w_mm: float, page_h_mm: float,
                         form=None, font_map: FontMap = None, texts: dict = None) -> None:
    """Рисует рамку листа и основную надпись по ГОСТ.

    texts — значения граф по ключам. Дополнительно:
      texts["cell9_logo"] — путь к PDF-файлу логотипа организации: если задан,
      графа 9 заполняется логотипом (вписывается в ячейку без искажений);
      texts["signatures"] — строки подписей левого блока (models.stamp_signature_rows):
      должность, фамилия, дата и путь к изображению подписи каждой строки.
    """
    font_map = font_map or FontMap()
    texts = texts or {}
    frame = Frame(page_w_mm, page_h_mm)

    # Внутренняя рамка листа
    fx0, fy0 = mm_to_pt(frame.x0), mm_to_pt(frame.y0)
    fx1, fy1 = mm_to_pt(frame.x1), mm_to_pt(frame.y1)
    page.draw_rect((fx0, fy0, fx1, fy1), color=BLACK, width=LINE_WIDTH_PT)

    fname = register_font(page) or "helv"

    # Высота и положение штампа (прижат к правому нижнему углу внутренней рамки)
    form = form if form is not None else FORM_TEXT_FIRST
    stamp_h = FORM_HEIGHT.get(form, 40.0)
    stamp_w = frame.width
    sx = frame.x0
    sy = frame.y1 - stamp_h

    # Внешняя рамка штампа
    page.draw_rect((mm_to_pt(sx), mm_to_pt(sy), mm_to_pt(sx + stamp_w), mm_to_pt(frame.y1)),
                   color=BLACK, width=LINE_WIDTH_PT)

    cells = _layout(form)

    # Заполнение из texts по ключу
    for c in cells:
        c["text"] = texts.get(c.get("key", ""), c.get("text", ""))

    # Строки подписей: значение каждой ячейки берётся из строки по её номеру.
    signatures = texts.get("signatures") or []
    for c in cells:
        role = c.get("role") or ""
        if not role.startswith("sig_"):
            continue
        row = _signature_row(signatures, c.get("key", ""))
        if role == "sig_post":
            c["text"] = str(row.get("post") or "")
        elif role == "sig_name":
            c["text"] = str(row.get("surname") or "")
        elif role == "sig_date":
            c["text"] = str(row.get("date") or "")
        elif role == "sig_sign":
            c["text"] = ""

    for c in cells:
        if c.get("role") == "design" and c["key"] == "volume":
            c["text"] = texts.get("volume", c["text"])
        # Графа 9: если задан логотип — он заменяет текст в ячейке.
        if c.get("key") in CELL9_KEYS and _draw_cell9_logo(page, sx, sy, c, texts):
            _draw_cell_border(page, sx, sy, c)
            continue
        # Графа «Подп.» строки: изображение подписи подписанта по фамилии строки.
        if c.get("role") == "sig_sign":
            row = _signature_row(signatures, c.get("key", ""))
            if _draw_signature_image(page, sx, sy, c, row):
                _draw_cell_border(page, sx, sy, c)
                continue
        _draw_cell(page, sx, sy, c, fname, font_map)

    return


def _signature_row(signatures, key) -> dict:
    """Строка подписей штампа по номеру в ключе ячейки (`sig...<номер>`)."""
    if not signatures:
        return {}
    digits = "".join(ch for ch in str(key or "") if ch.isdigit())
    if not digits:
        return {}
    try:
        index = int(digits)
    except ValueError:
        return {}
    for row in signatures:
        if int((row or {}).get("row", -1)) == index:
            return row
    return {}


def _draw_signature_image(page, sx_mm: float, sy_mm: float, cell: dict,
                          row: dict) -> bool:
    """Рисует изображение подписи в ячейке графы «Подп.» основной надписи.

    В ячейку вписывается САМ РОСЧЕРК: размеры берутся по обрезанному
    изображению (`preview_html.signature_fit_mm`). Иначе скан подписи с белыми
    полями уменьшался целиком, и росчерк получался в несколько миллиметров.
    Возвращает True, если картинка нарисована; False — подписи нет.
    """
    path = str((row or {}).get("signature_image") or "")
    if not path or not os.path.exists(path):
        return False
    fit_w_mm = fit_h_mm = 0.0
    try:
        from .preview_html import signature_fit_mm
        fit_w_mm, fit_h_mm = signature_fit_mm(path, cell["w"], cell["h"])
    except Exception:
        fit_w_mm = fit_h_mm = 0.0
    if fit_w_mm <= 0 or fit_h_mm <= 0:
        return False
    x = mm_to_pt(sx_mm + cell["x"])
    y = mm_to_pt(sy_mm + cell["y"])
    w = mm_to_pt(cell["w"])
    h = mm_to_pt(cell["h"])
    fit_w, fit_h = mm_to_pt(fit_w_mm), mm_to_pt(fit_h_mm)
    x0 = x + (w - fit_w) / 2.0
    y0 = y + (h - fit_h) / 2.0
    try:
        page.insert_image(pymupdf.Rect(x0, y0, x0 + fit_w, y0 + fit_h),
                          filename=path, keep_proportion=True)
        return True
    except Exception:
        return False


#: Минимальный кегль при автоподгонке текста в ячейку основной надписи.
MIN_FIT_PT = 5.0


def _fit_font_size(rect, text: str, fname: str, start_pt: float,
                   min_pt: float = MIN_FIT_PT) -> float:
    """Подбирает кегль, при котором текст целиком помещается в прямоугольник.

    Измерение выполняется на отдельной (черновой) странице, поэтому на исходную
    страницу ничего лишнего не попадает. Возвращает найденный кегль либо min_pt.
    """
    if not text:
        return start_pt
    try:
        scratch = pymupdf.open()
        try:
            sp = scratch.new_page(width=rect.width, height=rect.height)
            fs = float(start_pt)
            while fs > min_pt:
                rc = sp.insert_textbox(pymupdf.Rect(0, 0, rect.width, rect.height),
                                       text, fontname=fname, fontsize=fs,
                                       align=pymupdf.TEXT_ALIGN_CENTER)
                if rc >= 0:
                    return fs
                fs = round(fs - 0.5, 2)
        finally:
            scratch.close()
    except Exception:
        pass
    return min_pt


def _draw_cell_border(page, sx_mm, sy_mm, cell) -> None:
    """Рисует только рамку ячейки основной надписи (без текста).

    Используется, когда содержимое ячейки — изображение (логотип в графе 9).
    """
    x = mm_to_pt(sx_mm + cell["x"])
    y = mm_to_pt(sy_mm + cell["y"])
    w = mm_to_pt(cell["w"])
    h = mm_to_pt(cell["h"])
    page.draw_rect((x, y, x + w, y + h), color=BLACK, width=LINE_WIDTH_PT)


def _draw_cell(page, sx_mm, sy_mm, cell, fname, font_map: FontMap) -> None:
    x = mm_to_pt(sx_mm + cell["x"])
    y = mm_to_pt(sy_mm + cell["y"])
    w = mm_to_pt(cell["w"])
    h = mm_to_pt(cell["h"])
    page.draw_rect((x, y, x + w, y + h), color=BLACK, width=LINE_WIDTH_PT)

    text = cell.get("text", "")
    if not text:
        return
    is_cell9 = cell.get("key") in CELL9_KEYS
    if is_cell9:
        # Графа 9 — наименование/индекс организации: кегль как у «Стадия/Лист»,
        # при длинном наименовании он уменьшается, чтобы текст влез в ячейку.
        size_pt = CELL9_PT
    else:
        size_pt = _font_for(cell.get("key", ""), font_map, cell.get("role", ""))

    if is_cell9:
        size_pt = _fit_font_size(pymupdf.Rect(x, y, x + w, y + h), text,
                                 fname, size_pt)

    # insert_textbox ставит первую строку от ВЕРХА прямоугольника. Сдвигаем
    # прямоугольник вниз, чтобы текст оказался по центру ячейки (по вертикали).
    rect = _vcenter_rect(pymupdf.Rect(x, y, x + w, y + h), text, size_pt)
    try:
        rc = page.insert_textbox(rect, text, fontname=fname, fontsize=size_pt,
                                 color=BLACK, align=pymupdf.TEXT_ALIGN_CENTER)
        if rc < 0 and size_pt > 4:
            page.insert_textbox(rect, text, fontname=fname,
                                fontsize=max(4.0, size_pt - 1),
                                color=BLACK, align=pymupdf.TEXT_ALIGN_CENTER)
    except Exception:
        tx = x + w / 2
        ty = y + h / 2
        page.insert_text((tx, ty), text, fontname=fname, fontsize=size_pt, color=BLACK)
