# -*- coding: utf-8 -*-
"""Чтение DOCX: извлечение абзацев, заголовков, таблиц и рисунков
для последующего переоформления под ГОСТ."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Run:
    """Фрагмент (run) текста с форматированием."""
    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font: str = ""
    size_pt: float = 0.0


@dataclass
class Paragraph:
    """Абзац текста."""
    style_name: str = ""
    level: int = 0          # уровень заголовка (0 = обычный текст)
    alignment: str = ""     # LEFT/CENTER/RIGHT/JUSTIFY
    runs: list = field(default_factory=list)
    is_heading: bool = False
    # Оформление абзаца Word (отступы, интервалы, выравнивание, список) —
    # app.core.pz_format.ParagraphFormat. Заполняется read_docx; None означает
    # «оформление неизвестно» (объект создан вручную или из старого кэша).
    fmt: object = None

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class TableCell:
    text: str = ""
    bold: bool = False
    # Адрес ячейки в исходном DOCX (document.tables[table_index]
    # .rows[row_index].cells[cell_index]). Нужен для записи правок таблиц
    # обратно в файл из предпросмотра (app.core.docx_edit.apply_table_cell_edits).
    table_index: int = -1
    row_index: int = -1
    cell_index: int = -1
    # Размер объединённой ячейки и признак ячейки-продолжения:
    #   ""           — обычная ячейка (или начало объединения);
    #   "horizontal" — та же ячейка в объединении по горизонтали (не рисуется,
    #                  её перекрывает начало объединения с col_span > 1);
    #   "vertical"   — та же ячейка в объединении по вертикали (рисуется пустой,
    #                  если начало объединения осталось на предыдущем листе).
    row_span: int = 1
    col_span: int = 1
    continuation: str = ""

    @property
    def address(self) -> str:
        """Адрес ячейки в виде «таблица:строка:ячейка» ('' — если неизвестен)."""
        if self.table_index < 0 or self.row_index < 0 or self.cell_index < 0:
            return ""
        return f"{self.table_index}:{self.row_index}:{self.cell_index}"


@dataclass
class DocTable:
    rows: list = field(default_factory=list)   # list of list[TableCell]
    table_index: int = -1                      # индекс таблицы в document.tables


@dataclass
class DocImage:
    """Встроенное изображение внутри DOCX."""
    r_id: str = ""
    width_pt: float = 0.0
    height_pt: float = 0.0
    alt_text: str = ""


@dataclass
class DocxContent:
    """Извлечённое содержимое DOCX."""
    paragraphs: list = field(default_factory=list)
    tables: list = field(default_factory=list)
    images: list = field(default_factory=list)
    headings: list = field(default_factory=list)   # list[Paragraph]
    # Элементы основного тела в исходном порядке. Отдельные списки paragraphs
    # и tables сохраняются для обратной совместимости с остальным кодом.
    blocks: list = field(default_factory=list)     # ("paragraph", Paragraph) / ("table", DocTable)
    # Оформление абзацев: app.core.pz_format.ParagraphFormats
    # ({индекс абзаца в document.paragraphs: ParagraphFormat} + .headings).
    formats: object = None

    @property
    def plain_text(self) -> str:
        parts = [p.text for p in self.paragraphs if p.text]
        for t in self.tables:
            for row in t.rows:
                parts.append(" | ".join(c.text for c in row))
        return "\n".join(parts)


def _alignment_of(pf) -> str:
    try:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        val = pf.alignment
        if val is None:
            return ""
        mapping = {
            WD_ALIGN_PARAGRAPH.LEFT: "LEFT",
            WD_ALIGN_PARAGRAPH.CENTER: "CENTER",
            WD_ALIGN_PARAGRAPH.RIGHT: "RIGHT",
            WD_ALIGN_PARAGRAPH.JUSTIFY: "JUSTIFY",
        }
        return mapping.get(val, "")
    except Exception:
        return ""


def _cell_merge_flags(tc) -> tuple:
    """Флаги объединения ячейки Word по её XML (w:tcPr).

    :return: (vmerge, grid_span, hmerged), где
        vmerge    — "restart" | "continue" | ""   (объединение по вертикали);
        grid_span — размер объединения по горизонтали (w:gridSpan, >= 1);
        hmerged   — True, если ячейка поглощена объединением по горизонтали
                    (её w:tcPr содержит w:hMerge; §17.4.34 OOXML).
    """
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    try:
        pr = tc.find(W + "tcPr")
    except Exception:
        pr = None
    if pr is None:
        return "", 1, False
    vmerge = ""
    el = pr.find(W + "vMerge")
    if el is not None:
        val = el.get(W + "val")
        # <w:vMerge/> без значения или val="continue" — продолжение объединения,
        # val="restart" — его начало.
        vmerge = "continue" if val in (None, "", "continue") else str(val)
    grid_span = 1
    gs = pr.find(W + "gridSpan")
    if gs is not None:
        try:
            grid_span = max(1, int(gs.get(W + "val") or 1))
        except (TypeError, ValueError):
            grid_span = 1
    # <w:hMerge/> — ячейка поглощена объединением по горизонтали;
    # val="restart" (или отсутствие значения) — начало объединения.
    hmerge_el = pr.find(W + "hMerge")
    hmerged = False
    if hmerge_el is not None:
        hval = hmerge_el.get(W + "val")
        hmerged = hval not in ("restart",)
    return vmerge, grid_span, hmerged


def table_grid_cells(table, table_index: int) -> list:
    """Сетка ячеек таблицы DOCX для предпросмотра и записи правок.

    Возвращает список строк, каждая строка — список TableCell по числу
    физических ячеек w:tc в строке (len(row.cells)) с адресом
    (table_index, row_index, cell_index) и размерами объединения row_span /
    col_span. Адресация совпадает с document.tables[t].rows[r].cells[c], поэтому
    по ней можно записать правку обратно в файл.

    Объединённые ячейки распознаются по XML строки (w:gridSpan / w:hMerge /
    w:vMerge), потому что python-docx разрешает любую позицию объединения в
    ячейку-начало и по его API отличить ячейки внутри объединения нельзя.
    В предпросмотре:
      * начало горизонтального объединения получает col_span > 1;
      * начало вертикального объединения получает row_span > 1;
      * поглощённые ячейки помечаются continuation="horizontal"/"vertical":
        первые не рисуются (их перекрывает начало объединения), вторые рисуются
        пустыми там, где начало объединения осталось на предыдущем листе.
    """
    # Важно: row.cells в python-docx РАЗРЕШАЕТ объединённые ячейки в ячейку-
    # начало, поэтому флаги объединения читаем у «сырых» элементов w:tc.
    trs = list(table._tbl.tr_lst)
    raw_tcs = [[tc for tc in tr.tc_lst] for tr in trs]
    raw_flags = [[_cell_merge_flags(tc) for tc in tr_tcs] for tr_tcs in raw_tcs]
    # Схлопываем горизонтальные объединения логически: одна запись на колонку.
    flags = []
    for r, row_flags in enumerate(raw_flags):
        logical = []
        pending = 0
        for (vmerge, grid_span, hmerged) in row_flags:
            if pending > 0:
                # Продолжение уже учтённого объединения по горизонтали.
                logical.append((vmerge, 0, True))
                pending -= 1
                continue
            logical.append((vmerge, grid_span, hmerged))
            if grid_span > 1:
                pending = grid_span - 1
        flags.append(logical)

    from docx.table import _Cell

    # Число логических колонок таблицы (по сумме gridSpan строки).
    total_cols = 1
    for row_flags in flags:
        total_cols = max(total_cols, sum(max(1, f[1]) for f in row_flags))

    rows_out = []
    total_rows = len(raw_tcs)
    for r, tr_tcs in enumerate(raw_tcs):
        cells_out = []
        logical_pos = 0
        for c, tc in enumerate(tr_tcs):
            cell = _Cell(tc, table)
            vmerge, grid_span, hmerged = flags[r][c]
            if hmerged:
                # Поглощена горизонтальным объединением: колонку занимает начало.
                cells_out.append(TableCell(text="", bold=False,
                                           continuation="horizontal"))
                logical_pos += 1
                continue
            if vmerge == "continue":
                # Продолжение вертикального объединения: текст берём из начала
                # (Word хранит его только там), адрес оставляем свой — Word
                # при записи перенаправит правку в ячейку-начало.
                text = ""
                bold = False
                for rr in range(r - 1, -1, -1):
                    prev_flag = flags[rr][c] if c < len(flags[rr]) else ("", 1, False)
                    if prev_flag[0] == "restart":
                        try:
                            src = _Cell(raw_tcs[rr][c], table)
                            text = (src.text or "").strip()
                            bold = _cell_bold(src)
                        except (IndexError, ValueError):
                            text, bold = "", False
                        break
                    if prev_flag[0] != "continue":
                        break
                row_span = 1
                for rr in range(r + 1, total_rows):
                    nxt = flags[rr][c] if c < len(flags[rr]) else ("", 1, False)
                    if nxt[0] != "continue":
                        break
                    row_span += 1
                cells_out.append(TableCell(text=text, bold=bold,
                                           table_index=table_index,
                                           row_index=r, cell_index=c,
                                           row_span=row_span,
                                           continuation="vertical"))
                logical_pos += 1
                continue
            row_span = 1
            for rr in range(r + 1, total_rows):
                nxt = flags[rr][c] if c < len(flags[rr]) else ("", 1, False)
                if nxt[0] != "continue":
                    break
                row_span += 1
            span = max(1, grid_span)
            cells_out.append(TableCell(text=(cell.text or "").strip(),
                                       bold=_cell_bold(cell),
                                       table_index=table_index,
                                       row_index=r, cell_index=c,
                                       row_span=row_span,
                                       col_span=span))
            logical_pos += span
            # Колонки, занятые объединением, дополняем невидимыми ячейками,
            # чтобы строки таблицы выравнивались по своим колонкам.
            for _ in range(span - 1):
                cells_out.append(TableCell(text="", bold=False,
                                           continuation="horizontal"))
        # Выравниваем строку по числу колонок таблицы (строки с объединениями
        # короче остальных).
        while logical_pos < total_cols:
            cells_out.append(TableCell(text="", bold=False,
                                       continuation="horizontal"))
            logical_pos += 1
        rows_out.append(cells_out)
    return rows_out


def _cell_bold(cell) -> bool:
    """Первый run ячейки выделен полужирным?"""
    try:
        paras = cell.paragraphs
        return bool(paras[0].runs[0].bold) if paras and paras[0].runs else False
    except Exception:
        return False


def read_docx(path: str) -> DocxContent:
    """Открывает DOCX и извлекает структурированное содержимое."""
    from docx import Document
    from docx.document import Document as DocType
    from docx.table import Table as DocTableType
    from docx.text.paragraph import Paragraph as ParaType

    content = DocxContent()
    doc = Document(path)
    # Оформление абзацев (отступы, интервалы, выравнивание, списки) читается из
    # XML документа: python-docx не отдаёт наследование свойств стилей, а без
    # него предпросмотр и экспорт теряли абзацные отступы и списки Word.
    try:
        from .pz_format import read_paragraph_formats
        content.formats = read_paragraph_formats(path, base_size_pt=11.0)
    except Exception:
        content.formats = None
    # Индекс таблицы в document.tables. Порядок таблиц тела документа совпадает
    # с порядком document.tables, поэтому он годится для записи правок.
    table_index = -1

    # Итеративный обход: параграфы и таблицы в теле документа
    def iter_block_items(parent):
        from docx.oxml.text.paragraph import CT_P
        from docx.oxml.table import CT_Tbl
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        for child in parent.element.body.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    # Заголовок = определить по стилю
    def style_is_heading(style_name: str) -> bool:
        if not style_name:
            return False
        low = style_name.lower()
        return low.startswith("heading") or "heading" in low or low.startswith("заголовок")

    formats = getattr(content, "formats", None) or {}
    para_index = -1
    for block in iter_block_items(doc):
        if isinstance(block, ParaType):
            para_index += 1
            style_nm = block.style.name if block.style else ""
            fmt = formats.get(para_index)
            para = Paragraph(
                style_name=style_nm,
                alignment=_alignment_of(block.paragraph_format) if block.paragraph_format else "",
                is_heading=style_is_heading(style_nm),
                fmt=fmt,
            )
            if style_nm.lower().startswith("heading"):
                # уровень
                try:
                    lvl = int("".join(ch for ch in style_nm if ch.isdigit()) or 0)
                    para.level = lvl
                except Exception:
                    para.level = 0
            for run in block.runs:
                r = Run(
                    text=run.text,
                    bold=bool(run.bold),
                    italic=bool(run.italic),
                    underline=bool(run.underline),
                    font=run.font.name or "",
                    size_pt=(run.font.size.pt if run.font.size else 0.0),
                )
                para.runs.append(r)
            # Выравнивание берём из разобранного оформления, если прямое
            # выравнивание абзаца в файле не задано (в Word оно наследуется
            # от стиля).
            if not para.alignment and fmt is not None and getattr(fmt, "alignment", ""):
                para.alignment = str(fmt.alignment).upper()
            content.paragraphs.append(para)
            content.blocks.append(("paragraph", para))
            if para.is_heading and para.text.strip():
                content.headings.append(para)
        elif isinstance(block, DocTableType):
            table_index += 1
            rows = table_grid_cells(block, table_index)
            table = DocTable(rows=rows, table_index=table_index)
            content.tables.append(table)
            content.blocks.append(("table", table))

    return content


def read_table_layout(table, table_index: int = 0) -> list:
    """Сетка ячеек таблицы в виде списков словарей (для внешних потребителей).

    Тонкая обёртка над table_grid_cells: возвращает те же данные, но не
    датаклассами TableCell, а обычными словарями (удобно для отчётов/сериализации).
    Учитывает объединения ячеек так же, как предпросмотр.

    :param table: объект docx.table.Table
    :param table_index: индекс таблицы в document.tables (для адресации)
    :return: список строк; каждая строка — список словарей с ключами
        text, bold, row_index, cell_index, row_span, col_span, continuation.
    """
    grid = table_grid_cells(table, table_index)
    return [[{"text": c.text, "bold": c.bold, "row_index": c.row_index,
              "cell_index": c.cell_index, "row_span": c.row_span,
              "col_span": c.col_span, "continuation": c.continuation}
             for c in row] for row in grid]


def read_docx_tables(path: str) -> list:
    """Все таблицы тела документа в виде сеток (см. read_table_layout)."""
    from docx import Document
    try:
        doc = Document(path)
    except Exception:
        return []
    grids = []
    for index, table in enumerate(doc.tables):
        try:
            grids.append(read_table_layout(table, index))
        except Exception:
            grids.append([])
    return grids
