# -*- coding: utf-8 -*-
"""
Помощники для работы с таблицами DOCX: установка границ (рамок) ячеек,
объединение ячеек и настройка размеров колонок. Используется для отрисовки
основной надписи по ГОСТ Р 21.101-2020 в формате таблицы.
"""
from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell

#: Толщина линий оформления в Word — в восьмых пункта (6 = 0.75 pt ≈ 0.26 мм,
#: как 0.25 мм в HTML-предпросмотре и PDF). Единая величина для рамки листа
#: (page border) и границ ячеек основной надписи: если они различаются, на
#: стыке рамки и штампа видно утолщение линии.
LINE_BORDER_SZ = 6


def set_cell_border(cell, edges=("top", "bottom", "left", "right"),
                    sz: int = LINE_BORDER_SZ, color: str = "000000"):
    """Устанавливает границу ячейки (сплошная тонкая линия).

    :param edges: кортеж сторон ('top','bottom','left','right')
    :param sz: толщина линии в восьмых пункта (6 = 0.75pt)
    :param color: цвет линии (шестнадцатеричный RGB)
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in edges:
        tag = "w:" + edge
        el = tcBorders.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            tcBorders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def remove_cell_border(cell, edges=("top", "bottom", "left", "right")):
    """Удаляет границы ячейки (если заданы)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        return
    for edge in edges:
        tag = "w:" + edge
        el = tcBorders.find(qn(tag))
        if el is not None:
            tcBorders.remove(el)


def set_cell_edges(cell, top=True, bottom=True, left=True, right=True,
                   sz: int = LINE_BORDER_SZ, color: str = "000000"):
    """Задаёт границы ячейки по каждой стороне: рисовать или гасить.

    Нужно для основной надписи: у вертикально объединённых ячеек Word берёт
    границу из ``tcPr`` КАЖДОЙ строки объединения, поэтому одной ячейки мало —
    ``docx_exporter._add_stamp_table`` вызывает функцию для каждой строки.

    :param top, bottom, left, right: True — сплошная линия, False — ``w:val=nil``
        (сторона не рисуется; так убирают задвоение с линией рамки листа).
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge, draw in (("top", top), ("left", left),
                       ("bottom", bottom), ("right", right)):
        el = tcBorders.find(qn("w:" + edge))
        if el is None:
            el = OxmlElement("w:" + edge)
            tcBorders.append(el)
        el.set(qn("w:val"), "single" if draw else "nil")
        el.set(qn("w:sz"), str(sz) if draw else "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color if draw else "auto")


def set_stamp_cell_borders(table, regions, ncols: int, nrows: int,
                           sz: int = LINE_BORDER_SZ, color: str = "000000"):
    """Рисует сетку основной надписи по ячейкам, но без задвоения с рамкой.

    Особенности, из-за которых обычные приёмы не работают:

    * у вертикально объединённой ячейки Word берёт границу из ``tcPr`` КАЖДОЙ
      строки объединения, а не только первой. Поэтому обход идёт по XML
      (``w:tr`` → ``w:tc``), а не через ``table.cell()``: тот для объединённой
      ячейки всегда возвращает верхний левый ``tc``, и линии нижних строк
      оставались пустыми — сетка выходила рваной;
    * левая, правая и нижняя линии таблицы совпадают с линиями рамки листа.
      Их гасим (``w:val="nil"``): иначе Word рисует две параллельные линии в
      0.25 мм — видно задвоение по краю штампа.

    :param regions: список прямоугольников ячеек (row0, col0, col1, row1).
    """
    owner = {}
    for row0, col0, col1, row1 in regions:
        for r in range(row0, row1):
            for c in range(col0, col1):
                owner[(r, c)] = (row0, col0, col1, row1)

    for r, tr in enumerate(table._tbl.tr_lst):
        column = 0
        for tc in tr.tc_lst:
            span = tc.grid_span
            region = owner.get((r, column))
            if region is None:
                column += span
                continue
            row0, col0, col1, row1 = region
            set_cell_edges(
                _Cell(tc, table),
                top=(r == row0),
                # Нижнюю линию таблицы отдаём рамке листа.
                bottom=(r == row1 - 1 and row1 < nrows),
                left=(column == col0 and col0 > 0),
                right=(column + span == col1 and col1 < ncols),
                sz=sz, color=color)
            column += span


def set_cell_vertical_alignment(cell, align="center"):
    """Задаёт вертикальное выравнивание содержимого ячейки."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tcPr.append(vAlign)
    vAlign.set(qn("w:val"), align)


def set_cell_shading(cell, fill: str = "FFFFFF"):
    """Заливка ячейки цветом."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_width(cell, twips: int):
    """Устанавливает ширину ячейки в twips (1/1440 дюйма)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(twips))
    tcW.set(qn("w:type"), "dxa")


def set_table_grid_columns(table, widths_twips):
    """Устанавливает ширины колонок таблицы через w:tblGrid (в twips).

    Обновляет и gridCol, и ширину каждой ячейки в колонке, чтобы Word
    при фиксированной раскладке рисовал колонки нужной ширины.
    """
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        # вставляем после tblPr
        tblPr = tbl.tblPr
        tblPr.addnext(grid)
    # очистить существующие gridCol
    for gc in grid.findall(qn("w:gridCol")):
        grid.remove(gc)
    for w in widths_twips:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w)))
        grid.append(gc)
    # задать ширины ячеек по колонкам
    for row in table.rows:
        for ci, cell in enumerate(row.cells):
            if ci < len(widths_twips):
                set_cell_width(cell, int(widths_twips[ci]))


def merge_cells(table, row1, col1, row2, col2):
    """Объединяет прямоугольный диапазон ячеек (индексы 0-базисные)."""
    a = table.cell(row1, col1)
    b = table.cell(row2, col2)
    return a.merge(b)


def set_row_height(row, mm: float, rule="atLeast"):
    """Устанавливает высоту строки и запрещает разрыв строки между листами."""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cant_split = trPr.find(qn("w:cantSplit"))
    if cant_split is None:
        cant_split = OxmlElement("w:cantSplit")
        trPr.append(cant_split)
    trHeight = trPr.find(qn("w:trHeight"))
    if trHeight is None:
        trHeight = OxmlElement("w:trHeight")
        trPr.append(trHeight)
    # 1 мм = 56.6929 twips
    trHeight.set(qn("w:val"), str(int(round(mm * 56.6929))))
    trHeight.set(qn("w:hRule"), rule)


def mm_to_twips(mm: float) -> int:
    """Миллиметры -> twips."""
    return int(round(mm * 56.6929))


def set_table_cell_margins(table, left_mm: float = 0.0, right_mm: float = 0.0,
                           top_mm: float = 0.0, bottom_mm: float = 0.0):
    """Задаёт внутренние отступы ячеек всей таблицы (мм).

    По умолчанию Word оставляет по 1.9 мм слева и справа в каждой ячейке:
    в графе шириной 10 мм на текст остаётся 6.2 мм, и значение переносится.
    Основная надпись по ГОСТ рисуется без отступов — текст центрируется в
    графе, поэтому отступы обнуляются.
    """
    tblPr = table._tbl.tblPr
    mar = tblPr.find(qn("w:tblCellMar"))
    if mar is None:
        mar = OxmlElement("w:tblCellMar")
        tblPr.append(mar)
    for edge, value in (("top", top_mm), ("left", left_mm),
                        ("bottom", bottom_mm), ("right", right_mm)):
        el = mar.find(qn("w:" + edge))
        if el is None:
            el = OxmlElement("w:" + edge)
            mar.append(el)
        el.set(qn("w:w"), str(mm_to_twips(value)))
        el.set(qn("w:type"), "dxa")


def set_cell_no_wrap(cell, no_wrap: bool = True):
    """Запрещает перенос строк внутри ячейки (w:noWrap).

    В графах штампа шириной 10–20 мм перенос недопустим: значение должно
    остаться одной строкой, иначе высота строки растёт и штамп разъезжается.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    el = tcPr.find(qn("w:noWrap"))
    if no_wrap and el is None:
        tcPr.append(OxmlElement("w:noWrap"))
    elif not no_wrap and el is not None:
        tcPr.remove(el)


def set_table_borders(table, sz: int = LINE_BORDER_SZ, color: str = "000000",
                      edges=("top", "left", "bottom", "right",
                             "insideH", "insideV")):
    """Задаёт границы ВСЕЙ таблицы (w:tblBorders) — полная сетка.

    Границы ячеек (`set_cell_border`) задаются по одной ячейке и теряются у
    ячеек, полученных объединением: Word рисует объединение по своей логике, и
    часть линий пропадала. Границы на уровне таблицы действуют на все ячейки,
    включая объединённые, поэтому сетка всегда полная.

    :param edges: какие стороны рисовать. Для основной надписи левая, правая и
        нижняя совпадают с рамкой листа, поэтому их отдают рамке (`w:val=none`):
        иначе Word рисует две параллельные линии в 0.25 мм — «задвоение».
    """
    tblPr = table._tbl.tblPr
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn("w:" + edge))
        if el is None:
            el = OxmlElement("w:" + edge)
            borders.append(el)
        if edge not in edges:
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
            continue
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def set_table_indent(table, left_mm: float):
    """Сдвигает таблицу вправо от левой границы текста (w:tblInd).

    Word по умолчанию выводит таблицу с отрицательным отступом: ячейка имеет
    внутренний отступ 1.9 мм, и Word компенсирует его, уводя линии таблицы
    ЛЕВЕЕ текста. Без явного отступа таблица шириной 185 мм заходила за левую
    линию рамки (замер: x = 17.9 мм при рамке 19.7 мм).
    """
    tblPr = table._tbl.tblPr
    el = tblPr.find(qn("w:tblInd"))
    if el is None:
        el = OxmlElement("w:tblInd")
        tblPr.append(el)
    el.set(qn("w:w"), str(mm_to_twips(left_mm)))
    el.set(qn("w:type"), "dxa")


def set_table_center(table):
    """Выравнивает таблицу по центру текстовой области (w:jc = center).

    Текстовая область Word совпадает с внутренней областью рамки (185 мм), а
    таблицы нормируются либо на всю её ширину (содержание, состав, перечень —
    185 мм), либо на ширину с отступом 6 мм (таблицы ПЗ — 173 мм). По центру
    обе встают ровно: первая — линиями по рамке, вторая — с равными полями по
    6 мм, как `padding` у `.sheet-content` в HTML.

    Выравниванием, а не отступом (`w:tblInd`): Word выводит таблицу с
    отрицательным отступом на величину внутреннего отступа ячейки, и подбор
    `tblInd` под каждую комбинацию отступов ненадёжен. Центр не зависит от них.
    """
    tblPr = table._tbl.tblPr
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tblPr.append(jc)
    jc.set(qn("w:val"), "center")


def set_table_fixed_layout(table, width_mm: float):
    """Фиксированная раскладка таблицы с заданной общей шириной (мм)."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    # fixed layout
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")
    # общая ширина
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(mm_to_twips(width_mm)))
    tblW.set(qn("w:type"), "dxa")
