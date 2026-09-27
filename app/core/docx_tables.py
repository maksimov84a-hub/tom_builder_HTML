# -*- coding: utf-8 -*-
"""
Помощники для работы с таблицами DOCX: установка границ (рамок) ячеек,
объединение ячеек и настройка размеров колонок. Используется для отрисовки
основной надписи по ГОСТ Р 21.101-2020 в формате таблицы.
"""
from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_cell_border(cell, edges=("top", "bottom", "left", "right"),
                    sz: int = 6, color: str = "000000"):
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
    """Устанавливает высоту строки (в мм). rule: 'atLeast'|'exact'."""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
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
