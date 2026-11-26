# -*- coding: utf-8 -*-
"""Чтение DOCX: извлечение абзацев, заголовков, таблиц и рисунков
для последующего переоформления под ГОСТ."""
from __future__ import annotations

import os
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
    #: True — абзац взят из «надписи» (фигуры Word, w:txbxContent). Такого
    #: абзаца нет в document.paragraphs, поэтому он НЕ адресуется для правок:
    #: иначе правка записалась бы в чужой абзац файла.
    from_textbox: bool = False

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class TableCell:
    text: str = ""
    bold: bool = False
    alignment: str = ""
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
    #: Число строк-шапок (w:tblHeader) в начале таблицы. Word повторяет их ВСЕ
    #: при переносе таблицы на следующий лист, поэтому пагинация должна
    #: повторять столько же — иначе на листе-продолжении окажется меньше строк,
    #: чем разметил Word, и низ листа останется пустым.
    header_rows: int = 0


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
    #: True — текст найден в «надписях» (фигуры Word, w:txbxContent), а не в
    #: обычных абзацах. Такая записка требует внимания: правки в предпросмотре
    #: к ней не применяются, разметка страниц исходника не сохраняется.
    text_in_textboxes: bool = False

    @property
    def is_empty(self) -> bool:
        """В документе нет ни текста, ни таблиц (показывать нечего)."""
        has_text = any((p.text or "").strip() for p in self.paragraphs)
        return not has_text and not self.tables

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


def _header_row_count(table) -> int:
    """Сколько строк в начале таблицы помечено как шапка (w:tblHeader).

    Word повторяет при переносе таблицы на следующий лист ВСЕ такие строки, а не
    только первую. Пагинация должна повторять столько же: иначе на
    листе-продолжении окажется меньше строк, чем разметил Word.
    """
    from docx.oxml.ns import qn

    count = 0
    for row in table.rows:
        tr_pr = row._tr.find(qn("w:trPr"))
        header = tr_pr.find(qn("w:tblHeader")) if tr_pr is not None else None
        if header is None:
            break
        val = header.get(qn("w:val"))
        if val is not None and str(val).strip().lower() in ("0", "false", "off"):
            break
        count += 1
    return count


def table_grid_cells(table, table_index: int) -> list:
    """Сетка ячеек таблицы DOCX, выровненная по КОЛОНКАМ сетки (w:tblGrid).

    В каждой строке ровно столько ячеек, сколько колонок в таблице, поэтому все
    строки встают по своим колонкам.

    Почему не по «сырым» w:tc: Word описывает объединения двумя способами
    (w:gridSpan и legacy w:hMerge), и в одном файле они встречаются вместе.
    Подсчёт физических w:tc давал лишние колонки — таблица разваливалась:
    правые ячейки уезжали за границу и вставали со сдвигом по вертикали
    (например, колонка «Пол» в штатном расписании). python-docx сам разрешает
    объединения по сетке: row.cells отдаёт по одной ячейке на колонку, повторяя
    объединённую, поэтому совпадение элемента w:tc и есть признак объединения.

    Адрес ячейки (table_index, row_index, cell_index) — индексы в
    document.tables[t].rows[r].cells[c]; по нему записываются правки из
    предпросмотра (Word перенаправляет запись в начало объединения).

    В предпросмотре:
      * начало объединения по горизонтали получает col_span > 1;
      * начало объединения по вертикали получает row_span > 1;
      * поглощённые ячейки помечаются continuation="horizontal" (не рисуются,
        их перекрывает начало объединения) либо "vertical" (рисуются пустыми
        там, где начало объединения осталось на предыдущем листе).
    """
    grid = [list(row.cells) for row in table.rows]
    total_rows = len(grid)
    if not total_rows:
        return []
    total_cols = max(len(row) for row in grid)
    # Элементы w:tc по колонкам: объединённая ячейка повторяется на каждой
    # колонке, которую занимает.
    elem = [[cell._tc for cell in row] for row in grid]

    rows_out = []
    for r in range(total_rows):
        row_elems = elem[r]
        row_cells = grid[r]
        cells_out = []
        for c in range(total_cols):
            tc = row_elems[c] if c < len(row_elems) else None
            prev = row_elems[c - 1] if c > 0 and (c - 1) < len(row_elems) else None
            if tc is None:
                # Строка короче сетки (нестандартная разметка) — добираем пустой.
                cells_out.append(TableCell(continuation="horizontal"))
                continue
            if c > 0 and tc is prev:
                # Та же ячейка, что и слева — колонку занимает начало объединения.
                cells_out.append(TableCell(continuation="horizontal"))
                continue
            # Размер объединения по горизонтали.
            col_span = 1
            while (c + col_span < len(row_elems)
                   and row_elems[c + col_span] is tc):
                col_span += 1
            # Та же ячейка выше — это продолжение вертикального объединения.
            # Горизонтальный охват (col_span) считается и здесь: продолжение
            # занимает столько же колонок, сколько начало объединения, иначе в
            # строке оставалась бы незанятая колонка — дыра в таблице.
            above = elem[r - 1] if r > 0 else None
            if above is not None and c < len(above) and above[c] is tc:
                cells_out.append(TableCell(table_index=table_index,
                                           row_index=r, cell_index=c,
                                           row_span=1, col_span=col_span,
                                           continuation="vertical"))
                continue
            # Размер объединения по вертикали.
            row_span = 1
            while (r + row_span < total_rows
                   and c < len(elem[r + row_span])
                   and elem[r + row_span][c] is tc):
                row_span += 1
            cell = row_cells[c]
            cells_out.append(TableCell(text=(cell.text or "").strip(),
                                       bold=_cell_bold(cell),
                                       alignment=_cell_alignment(cell),
                                       table_index=table_index,
                                       row_index=r, cell_index=c,
                                       row_span=row_span,
                                       col_span=col_span))
        rows_out.append(cells_out)
    return rows_out


def _cell_bold(cell) -> bool:
    """Первый run ячейки выделен полужирным?"""
    try:
        paras = cell.paragraphs
        return bool(paras[0].runs[0].bold) if paras and paras[0].runs else False
    except Exception:
        return False


def _cell_alignment(cell) -> str:
    try:
        if not cell.paragraphs:
            return ""
        paragraph = cell.paragraphs[0]
        alignment = _alignment_of(paragraph.paragraph_format)
        if alignment:
            return alignment
        style = paragraph.style
        seen = set()
        while style is not None and style.style_id not in seen:
            seen.add(style.style_id)
            alignment = _alignment_of(style.paragraph_format)
            if alignment:
                return alignment
            style = style.base_style
    except Exception:
        pass
    return ""


def style_name_of(paragraph, cache: dict | None = None) -> str:
    """Имя стиля абзаца; разрешение кэшируется по styleId.

    ``paragraph.style`` поднимает styles.xml и разрешает наследование при
    каждом обращении — около 1 мс на абзац (0.27 с на записке в 240 абзацев),
    хотя различных стилей в документе единицы. Кэш по styleId сохраняет именно
    имя стиля, а не сырой идентификатор: у типовых записок styleId числовые
    («1», «20» → «heading 1», «heading 2»), и по ним заголовок не распознать.

    :param cache: словарь {styleId: имя} на документ (см. _read_docx_uncached).
    """
    try:
        style_id = str(getattr(paragraph._p, "style", "") or "")
    except Exception:
        style_id = ""
    if cache is not None and style_id in cache:
        return cache[style_id]
    try:
        name = paragraph.style.name if paragraph.style else ""
    except Exception:
        name = ""
    if cache is not None:
        cache[style_id] = name
    return name


def _read_paragraph(block, formats, para_index: int, style_is_heading,
                    style_cache: dict | None = None) -> Paragraph:
    """Разбирает один абзац DOCX: стиль, выравнивание, runs и оформление."""
    # Имя стиля разрешается через кэш: повторные абзацы того же стиля не
    # поднимают styles.xml заново (см. style_name_of).
    style_nm = style_name_of(block, style_cache)
    fmt = formats.get(para_index)
    try:
        alignment = _alignment_of(block.paragraph_format) if block.paragraph_format else ""
    except Exception:
        alignment = ""
    para = Paragraph(
        style_name=style_nm,
        alignment=alignment,
        is_heading=style_is_heading(style_nm),
        fmt=fmt,
    )
    if style_nm.lower().startswith("heading"):
        try:
            para.level = int("".join(ch for ch in style_nm if ch.isdigit()) or 0)
        except Exception:
            para.level = 0
    for run in block.runs:
        try:
            r = Run(
                text=run.text,
                bold=bool(run.bold),
                italic=bool(run.italic),
                underline=bool(run.underline),
                font=run.font.name or "",
                size_pt=(run.font.size.pt if run.font.size else 0.0),
            )
        except Exception:
            r = Run(text=getattr(run, "text", "") or "")
        para.runs.append(r)
    # Выравнивание берём из разобранного оформления, если прямое
    # выравнивание абзаца в файле не задано (в Word оно наследуется
    # от стиля).
    if not para.alignment and fmt is not None and getattr(fmt, "alignment", ""):
        para.alignment = str(fmt.alignment).upper()
    return para


def _textbox_paragraphs(block) -> list:
    """Абзацы из «надписей» внутри абзаца DOCX (фигуры/текстовые поля Word).

    Word хранит текст фигуры в `w:txbxContent` — внутри `v:textbox` (VML) или
    `wps:txbx` (DrawingML). `python-docx` не включает это содержимое в
    `paragraph.runs`, поэтому его читаем из XML напрямую, сохраняя порядок
    абзацев и их оформление (полужирный, выравнивание по стилю фигуры теряется —
    это осознанное упрощение).
    """
    try:
        element = block._p
    except Exception:
        return []
    from docx.oxml.ns import qn
    out = []
    # w:txbxContent встречается и в VML (v:textbox), и в DrawingML (wps:txbx).
    for holder in element.iter(qn("w:txbxContent")):
        for p_el in holder.findall(qn("w:p")):
            runs = []
            for r_el in p_el.findall(qn("w:r")):
                text = "".join(t.text or "" for t in r_el.iter(qn("w:t")))
                if not text:
                    continue
                bold = r_el.find(qn("w:rPr") + "/" + qn("w:b")) is not None
                italic = r_el.find(qn("w:rPr") + "/" + qn("w:i")) is not None
                runs.append(Run(text=text, bold=bold, italic=italic))
            if not runs:
                continue
            out.append(Paragraph(runs=runs, alignment="", from_textbox=True))
    return out


def docx_cache_key(path: str) -> tuple | None:
    """Ключ состояния файла для кэшей: (путь, время изменения, размер).

    None — файла нет или он недоступен (кэшировать нечего). Общий помощник:
    по этому ключу кэшируется и разбор DOCX, и разбивка по листам, поэтому
    любая запись в файл автоматически сбрасывает оба кэша.
    """
    if not path or not os.path.exists(path):
        return None
    try:
        st = os.stat(path)
    except OSError:
        return None
    return (os.path.abspath(path), int(st.st_mtime_ns), int(st.st_size))


def paragraph_formats(path: str):
    """Оформление абзацев DOCX по настройкам проекта (см. pz_format).

    Вынесено из разбора документа: часть потребителей (подготовка
    нормализованной копии для Word) нуждается только в оформлении абзацев и не
    использует ни тексты, ни таблицы. Полный ``read_docx`` тратил бы на них
    время зря — разбор стоит 0.5–0.9 с на записке в 240 абзацев.
    """
    try:
        from .pz_format import read_paragraph_formats, DEFAULT_INDENT_CHARS
        # Абзацный отступ (5 знаков) берётся из конфигурации оформления:
        # его величина в миллиметрах считается от кегля основного текста ПЗ
        # (5 × 14 pt × 0.5 = 35 pt = 12.35 мм), а не от кегля Word.
        indent_chars = DEFAULT_INDENT_CHARS
        indent_size_pt = 11.0
        try:
            from ..config import load_config
            cfg = load_config()
            indent_chars = cfg.paragraph_indent_chars
            # Кегль для пересчёта знаков в мм — кегль обычного текста из
            # настроек (config.TextFonts): при его смене отступ пересчитывается.
            indent_size_pt = float(cfg.text_fonts.body_size)
        except Exception:
            pass
        return read_paragraph_formats(path, base_size_pt=11.0,
                                      indent_chars=indent_chars,
                                      indent_size_pt=indent_size_pt)
    except Exception:
        return None


def read_docx(path: str) -> DocxContent:
    """Открывает DOCX и извлекает структурированное содержимое.

    Результат кэшируется по (путь, время изменения, размер): за одно
    переключение раздела предпросмотр читает один и тот же DOCX несколько раз
    (подсчёт листов для содержания и штампа, оглавление, сами листы), и каждый
    разбор файла занимал сотни миллисекунд — переключение ощущалось медленным.
    Кэш сбрасывается сам при изменении файла (правка абзаца, автосохранение).
    """
    path = path or ""
    if not path or not os.path.exists(path):
        return _read_docx_uncached(path)
    key = docx_cache_key(path)
    if key is None:
        return _read_docx_uncached(path)
    cached = _DOCX_CACHE.get(key)
    if cached is not None:
        return cached
    content = _read_docx_uncached(path)
    if len(_DOCX_CACHE) >= _DOCX_CACHE_LIMIT:
        _DOCX_CACHE.clear()
    _DOCX_CACHE[key] = content
    return content


#: Кэш разобранных DOCX: (путь, время изменения, размер) -> DocxContent.
_DOCX_CACHE: dict = {}
_DOCX_CACHE_LIMIT = 24


def clear_docx_cache(path: str | None = None) -> None:
    """Сбрасывает кэш разбора DOCX (после записи правок в файл)."""
    if path is None:
        _DOCX_CACHE.clear()
        return
    target = os.path.abspath(path)
    for key in [k for k in _DOCX_CACHE if k[0] == target]:
        _DOCX_CACHE.pop(key, None)


def _read_docx_uncached(path: str) -> DocxContent:
    from docx import Document
    from docx.document import Document as DocType
    from docx.table import Table as DocTableType
    from docx.text.paragraph import Paragraph as ParaType

    content = DocxContent()
    doc = Document(path)
    # Оформление абзацев (отступы, интервалы, выравнивание, списки) читается из
    # XML документа: python-docx не отдаёт наследование свойств стилей, а без
    # него предпросмотр и экспорт теряли абзацные отступы и списки Word.
    content.formats = paragraph_formats(path)
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
    # Кэш «styleId -> имя стиля» на документ: разные абзацы почти всегда
    # используют одни и те же стили, а разрешение имени стоит около 1 мс.
    style_cache: dict = {}
    for block in iter_block_items(doc):
        if isinstance(block, ParaType):
            para_index += 1
            # Один «нестандартный» абзац (неизвестный стиль, битый элемент)
            # не должен ронять разбор всего документа: раньше исключение здесь
            # приводило к тому, что в предпросмотре вместо записки появлялась
            # заглушка. При сбое берём из абзаца хотя бы текст.
            try:
                para = _read_paragraph(block, formats, para_index, style_is_heading,
                                       style_cache)
            except Exception:
                para = Paragraph(runs=[Run(text=block.text or "")])
            content.paragraphs.append(para)
            content.blocks.append(("paragraph", para))
            if para.is_heading and para.text.strip():
                content.headings.append(para)
            # Текст внутри «надписи» (фигуры Word: w:pict/v:textbox,
            # mc:AlternateContent/wps:txbx → w:txbxContent). python-docx его в
            # runs не отдаёт, а типовые записки верстают текст именно так —
            # раньше такой документ считался пустым и лист выходил пустым.
            try:
                tb_paras = _textbox_paragraphs(block)
            except Exception:
                tb_paras = []
            for tb_para in tb_paras:
                content.paragraphs.append(tb_para)
                content.blocks.append(("paragraph", tb_para))
                if tb_para.is_heading and tb_para.text.strip():
                    content.headings.append(tb_para)
                content.text_in_textboxes = True
        elif isinstance(block, DocTableType):
            table_index += 1
            rows = table_grid_cells(block, table_index)
            table = DocTable(rows=rows, table_index=table_index,
                             header_rows=_header_row_count(block))
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
