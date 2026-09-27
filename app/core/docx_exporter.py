# -*- coding: utf-8 -*-
"""
Экспорт текстовой части тома в DOCX (обложка + титульный лист + содержание +
пояснительная записка) с оформлением по ГОСТ Р 21.101-2020.

Основная надпись (форма 5 / 6) выводится как таблица в нижнем колонтитуле
(разные для первой и последующих страниц). Шрифт — ISOCPEUR, кегль — по карте.

Рендер DOCX не зависит от Word — только python-docx. Конвертация в PDF
(через Word COM) — в word_com.py.
"""
from __future__ import annotations

import os
from typing import Optional

from ..models import (
    VolumeProject, DocKind, Stage, SECTION_LABELS,
    SECTION_COVER, SECTION_TITLE, SECTION_CONTENT, SECTION_PZ,
    SECTION_CHANGES,
)
from ..config import Config, FontMap, PageMargins, A4_MM, FONT_NAME
from .docx_reader import read_docx
from .toc import content_rows
# Автоматическое оглавление ПЗ по уровням текста из Word (единый источник).
from .pz_toc import (project_toc_items, paginate_items, toc_page_offset,
                     DOCX_TAB_POS_MM, LEVEL_INDENT_MM)
from .gost import (LEFT_MM, RIGHT_MM, TOP_MM, BOTTOM_MM,
                   FORM_GRAPHIC, FORM_TEXT_FIRST, FORM_TEXT_NEXT,
                   STAMP_WIDTH_MM)
# Кегль граф «Стадия / Лист / Листов» (единый источник — geometry.py).
from .geometry import STAGE_CELL_KEYS, STAGE_CELL_PT, CELL9_KEYS, CELL9_PT
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

# python-docx использует EMU (English Metric Units): 1 мм = 36000 EMU
MM_TO_EMU = 36000

def _mm(v: float) -> int:
    """Миллиметры -> EMU (для page_width / margins в python-docx)."""
    return int(round(v * MM_TO_EMU))


def _set_section(page_w_mm, page_h_mm, margins: PageMargins, section):
    """Настраивает размер страницы и поля раздела DOCX."""
    section.page_width = _mm(page_w_mm)
    section.page_height = _mm(page_h_mm)
    section.left_margin = _mm(margins.left)
    section.right_margin = _mm(margins.right)
    section.top_margin = _mm(margins.top)
    section.bottom_margin = _mm(margins.bottom)


def _set_run(run, font: str, size_pt: float, bold=False):
    from docx.shared import Pt
    run.font.name = font
    run.font.size = Pt(size_pt)
    run.bold = bold
    # восточно-европейский / кириллический шрифт
    rpr = run._element.get_or_add_rPr()
    rFonts = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    if rFonts is None:
        rFonts = rpr.makeelement('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts', {})
        rpr.append(rFonts)
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', font)
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', font)
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs', font)


# Текстовые kinds, участвующие в DOCX-томе (графика/приложения не входят).
TEXT_SECTION_KINDS = (
    SECTION_COVER, SECTION_TITLE, SECTION_CONTENT, SECTION_PZ, SECTION_CHANGES,
)


def _structure_text_kinds(project: VolumeProject) -> list:
    """Порядок текстовых разделов из project.structure (только needed)."""
    kinds = []
    for entry in project.structure or []:
        if not entry.get("needed", True):
            continue
        kind = entry.get("kind") or ""
        if kind in TEXT_SECTION_KINDS:
            kinds.append(kind)
    return kinds


def build_text_docx(project: VolumeProject, out_path: str,
                    config: Config | None = None,
                    kinds: list | None = None) -> str:
    """Собирает DOCX текстовой части тома.

    Порядок и состав разделов берутся из ``project.structure`` (отфильтрованной
    по ``needed``). Параметр ``kinds`` позволяет собрать только указанные
    текстовые разделы (используется при сегментации PDF-сборки).

    Графика и приложения в текстовый DOCX не включаются (SPEC: DOC — без
    графики и приложений). Возвращает путь к созданному DOCX.
    """
    config = config or Config()
    fm: FontMap = config.font_map
    mm: PageMargins = config.margins
    font = config.font_name

    from docx import Document
    from docx.enum.section import WD_SECTION

    if kinds is None:
        text_kinds = _structure_text_kinds(project)
    else:
        text_kinds = [k for k in kinds if k in TEXT_SECTION_KINDS]
    # changes в DOCX попадает только при непустом журнале
    text_kinds = [
        k for k in text_kinds
        if k != SECTION_CHANGES or project.changes
    ]

    doc = Document()
    # ---- Раздел по умолчанию: A4, поля ГОСТ ----
    _set_section(A4_MM[0], A4_MM[1], mm, doc.sections[0])

    for i, kind in enumerate(text_kinds):
        if i == 0:
            sec = doc.sections[0]
        else:
            sec = doc.add_section(WD_SECTION.NEW_PAGE)
            _set_section(A4_MM[0], A4_MM[1], mm, sec)

        if kind == SECTION_COVER:
            _add_cover(doc, project, fm, font)
            _no_stamp_section(sec)
        elif kind == SECTION_TITLE:
            _add_title_sheet(doc, project, fm, font)
            _no_stamp_section(sec)
        elif kind == SECTION_CONTENT:
            _add_content(doc, project, fm, font)
            _add_stamp_footers(doc, project, config, sec, first_form=FORM_TEXT_FIRST)
        elif kind == SECTION_PZ:
            _add_pz_toc(doc, project, fm, font)
            for src in project.documents_for(DocKind.POYASNITELNAYA):
                _add_pz_doc(doc, src.path, fm, font)
            _add_stamp_footers(doc, project, config, sec, first_form=FORM_TEXT_FIRST)
        elif kind == SECTION_CHANGES:
            _add_changes_sheet(doc, project, fm, font)
            _add_stamp_footers(doc, project, config, sec, first_form=FORM_TEXT_FIRST)

    # Пустой текстовый том — минимальный лист, чтобы файл был валиден
    if not text_kinds:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run("Нет текстовых разделов в структуре тома"),
                 font, fm.content_header)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    return out_path


def _no_stamp_section(sec) -> None:
    """Убирает основную надпись и рамку у раздела (обложка / титульный лист)."""
    sec.different_first_page_header_footer = False
    for footer in (sec.footer, sec.first_page_footer, sec.header, sec.first_page_header):
        footer.is_linked_to_previous = False
        _clear_footer(footer)
    # удаляем рамку страницы, если она была установлена
    from docx.oxml.ns import qn
    sectPr = sec._sectPr
    for el in sectPr.findall(qn("w:pgBorders")):
        sectPr.remove(el)


def _add_cover(doc, project, fm, font):
    """Обложка — приложение Е ГОСТ."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    req = project.requisites
    # логотип/организация (центр, верх)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run(req.organization if req.organization else ""), font, fm.title_doc_name, bold=True)
    body = doc.add_paragraph()
    body.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(body.add_run("ОБЛОЖКА"), font, fm.stamp_volume, bold=True)
    # обозначение тома
    if req.volume_id:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run(req.volume_id), font, fm.stamp_volume)
    doc.add_page_break()


def _add_title_sheet(doc, project, fm, font):
    """Титульный лист — приложение Ж ГОСТ."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    req = project.requisites
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run(req.object_name if req.object_name else ""), font, fm.title_doc_name, bold=True)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run(req.volume_name if req.volume_name else ""), font, fm.title_doc_name)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run("ТИТУЛЬНЫЙ ЛИСТ"), font, fm.stamp_volume, bold=True)
    # стадия / город / год
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    stage = req.stage.value if isinstance(req.stage, Stage) else str(req.stage)
    loc = " ".join(x for x in [req.city, str(req.year)] if x)
    _set_run(p.add_run(f"Стадия {stage}    {loc}"), font, fm.stamp_columns)
    doc.add_page_break()


def _add_pz_toc(doc, project, fm, font, offset: int | None = None,
                standalone: bool = False) -> int:
    """Добавляет автоматическое оглавление записки первым листом раздела ПЗ.

    Заголовки берутся из исходного DOCX по уровням текста Word
    (app.core.pz_toc): строки выводятся с абзацным отступом по уровню, номером
    листа записки и отточием. Отточие — правая табуляция с заполнителем DOTS
    (как «Содержание» в Word), поэтому номер листа стоит у правого края
    текстовой области, а промежуток заполнен точками.

    :param offset: сдвиг номеров листов (номер листа записки в томе минус 1) —
        чтобы напечатанный номер совпал с квадратом сквозной нумерации листа.
    :return: число добавленных листов (0 — оглавление не выводится).
    """
    from docx.shared import Mm, Pt
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER

    items, _path = project_toc_items(project)
    if not items:
        return 0
    pages = paginate_items(items)
    # Номера листов в оглавлении — сквозные по тому: сдвиг = номер первого
    # листа записки минус 1 (в отдельном разделе — число листов оглавления).
    if offset is None:
        offset = toc_page_offset(project, standalone=standalone)
    size = fm.toc_row
    title_size = fm.content_header

    for page_no, page_items in enumerate(pages):
        # Оглавление разбито по листам той же моделью высоты, что и записка:
        # каждый следующий лист начинается с новой страницы.
        if page_no:
            doc.add_page_break()
        for item in page_items:
            p = doc.add_paragraph()
            if item.is_title:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(12)
                _set_run(p.add_run(item.text), font, title_size, bold=True)
                continue
            align_left = WD_ALIGN_PARAGRAPH.LEFT
            p.alignment = align_left
            indent_mm = LEVEL_INDENT_MM * max(0, int(item.level) - 1)
            pf = p.paragraph_format
            if indent_mm:
                pf.left_indent = Mm(indent_mm)
            # Номер листа у правого края с отточием; позиция — от левого края
            # текстовой области, поэтому отступ уровня из неё вычитается.
            try:
                pf.tab_stops.add_tab_stop(
                    Mm(max(20.0, DOCX_TAB_POS_MM - indent_mm)),
                    WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            except Exception:
                pass
            # Строка оглавления: номер раздела и наименование — полужирным.
            if item.num:
                _set_run(p.add_run(item.num + " "), font, size, bold=True)
            _set_run(p.add_run(item.text), font, size, bold=True)
            if item.page:
                _set_run(p.add_run("\t" + str(int(item.page) + max(0, int(offset or 0)))),
                         font, size)
    return len(pages)


def _add_content(doc, project, fm, font):
    """Содержание (опись) — автоматический сбор по фактической структуре тома.

    Строки содержания формируются единым источником (app.core.toc) по реальной
    структуре тома и фактическому числу листов каждого раздела/документа/
    приложения. В отличие от ручного перечня requisites.sections, содержание
    корректно показывает сквозные номера листов («Лист N — M»).
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from .docx_tables import (set_table_fixed_layout, set_table_grid_columns,
                              set_cell_border, mm_to_twips)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run("СОДЕРЖАНИЕ"), font, fm.content_header, bold=True)

    rows = content_rows(project)    # [(наименование, обозначение, листы), ...]

    # Таблица содержания (3 колонки): Наименование | Обозначение | Листы.
    if rows:
        ncols = 3
        # +1 строка-заголовок («шапка») таблицы.
        t = doc.add_table(rows=len(rows) + 1, cols=ncols)
        t.autofit = False
        set_table_fixed_layout(t, A4_MM[0] - 25.0)   # ~185 мм текстовой области
        # Ширины колонок (мм): наименование (широкая) | обозначение | листы.
        set_table_grid_columns(t, [mm_to_twips(w) for w in (110.0, 50.0, 25.0)])
        headers = ["Наименование", "Обозначение", "Листы"]
        for ci, hdr in enumerate(headers):
            cell = t.cell(0, ci)
            set_cell_border(cell, ("top", "bottom", "left", "right"))
            cell.text = ""
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_run(para.add_run(hdr), font, fm.content_row, bold=True)
        for ri, (name, des, sheets) in enumerate(rows, start=1):
            cells = [name, des, sheets]
            for ci, val in enumerate(cells):
                cell = t.cell(ri, ci)
                set_cell_border(cell, ("top", "bottom", "left", "right"))
                cell.text = ""
                para = cell.paragraphs[0]
                para.alignment = (WD_ALIGN_PARAGRAPH.LEFT if ci == 0
                                  else WD_ALIGN_PARAGRAPH.CENTER)
                _set_run(para.add_run(val), font, fm.content_row)
    else:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run("Состав тома не заполнен"), font, fm.content_row)
    doc.add_page_break()


# Соответствие выравнивания ГОСТ-текста значениям python-docx.
_ALIGN_ENUM = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _apply_paragraph_format(p, fmt, align_default=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """Переносит оформление абзаца Word (отступы, интервалы, выравнивание) в DOCX.

    Оформление берётся из app.core.docx_reader (app.core.pz_format): отступ
    первой строки («красная строка»), абзацные отступы, интервалы до/после и
    межстрочный интервал — как в исходном документе. Если в исходнике
    выравнивание не задано, применяется выключка по ширине (ГОСТ).

    Кегль и шрифт здесь не переносятся: их задаёт карта кеглей (config.FontMap).
    """
    from docx.shared import Mm, Pt
    from .pz_format import docx_indent_kwargs

    kwargs = docx_indent_kwargs(fmt) if fmt is not None else {}
    pf = p.paragraph_format
    alignment = (kwargs.get("alignment") or "").lower()
    p.alignment = _ALIGN_ENUM.get(alignment, align_default)
    is_list = bool(fmt is not None and getattr(fmt, "has_list", False))
    try:
        if kwargs.get("left_mm"):
            pf.left_indent = Mm(kwargs["left_mm"])
        if kwargs.get("right_mm"):
            pf.right_indent = Mm(kwargs["right_mm"])
        if kwargs.get("first_line_mm") and not is_list:
            # Отрицательное значение python-docx сам переводит в «выступ»
            # (w:hanging) — так в Word и задаётся позиция маркера списка.
            pf.first_line_indent = Mm(kwargs["first_line_mm"])
        if is_list:
            # Пункт списка: «висячий» отступ (w:ind left + hanging) оставляем
            # как в исходном документе, а позицию табуляции задаём на уровне
            # текста пункта. Тогда маркер («1.», «•») стоит слева, а текст
            # начинается ровно на отступе — точно как в Word.
            from docx.enum.text import WD_TAB_ALIGNMENT
            if kwargs.get("first_line_mm"):
                pf.first_line_indent = Mm(kwargs["first_line_mm"])
            try:
                pf.tab_stops.add_tab_stop(Mm(kwargs.get("left_mm") or 0.0),
                                          WD_TAB_ALIGNMENT.LEFT)
            except Exception:
                pass
        if kwargs.get("space_before_pt"):
            pf.space_before = Pt(kwargs["space_before_pt"])
        if kwargs.get("space_after_pt"):
            pf.space_after = Pt(kwargs["space_after_pt"])
        if kwargs.get("line_spacing_pt"):
            from docx.shared import Pt as _Pt
            pf.line_spacing = _Pt(kwargs["line_spacing_pt"])
        elif kwargs.get("line_spacing"):
            pf.line_spacing = float(kwargs["line_spacing"])
    except Exception:
        # Оформление не критично для сборки: при ошибке остаётся ГОСТ-вид.
        pass


def _list_marker_run(p, fmt, size, font):
    """Добавляет в абзац маркер списка Word («1.», «а)», «•»).

    Нумерация Word воспроизводится текстом маркера, а не ссылкой на
    numbering.xml: определение списка живёт в исходном пакете DOCX, и простой
    его перенос ломается при объединении нескольких документов в один том.
    Отступ маркера при этом задаётся настоящими отступами абзаца
    (w:ind left + hanging), а табуляция после маркера ставит текст пункта на
    отступ left — поэтому пункт выглядит как список Word.
    """
    if fmt is None or not getattr(fmt, "has_list", False):
        return False
    text = fmt.marker_text
    if not text:
        return False
    _set_run(p.add_run(text + "\t"), font, size)
    return True


def _add_pz_doc(doc, path, fm, font):
    """Импортирует DOCX-исходник в ПЗ с переоформлением под ГОСТ."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    if not path or not os.path.exists(path):
        return
    try:
        content = read_docx(path)
    except Exception:
        return
    for para in content.paragraphs:
        text = para.text
        if not text.strip():
            continue
        size = fm.text_heading if para.is_heading else fm.text_body
        fmt = getattr(para, "fmt", None)
        p = doc.add_paragraph()
        # Оформление абзаца из Word: отступы, интервалы, выравнивание, отступ
        # маркера списка (пункт списка сохраняет свой вид и нумерацию).
        _apply_paragraph_format(p, fmt)
        _list_marker_run(p, fmt, size, font)
        for run in para.runs:
            _set_run(p.add_run(run.text), font, size, bold=run.bold)
    # таблицы
    for tbl in content.tables:
        rows = len(tbl.rows)
        if rows == 0:
            continue
        cols = max((len(r) for r in tbl.rows), default=1)
        t = doc.add_table(rows=rows, cols=cols)
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row):
                # Ячейки внутри объединения заполняются самим объединением
                # (иначе текст начала объединения продублировался бы, а текст
                # вертикального объединения потерялся бы).
                if getattr(cell, "continuation", ""):
                    continue
                if ci >= cols:
                    continue
                t.cell(ri, ci).text = cell.text
                row_span = int(getattr(cell, "row_span", 1) or 1)
                col_span = int(getattr(cell, "col_span", 1) or 1)
                if row_span > 1 or col_span > 1:
                    # Восстанавливаем объединение по фактическим размерам ячейки
                    # исходного документа (w:gridSpan / w:vMerge).
                    try:
                        t.cell(ri, ci).merge(
                            t.cell(min(ri + row_span - 1, rows - 1),
                                   min(ci + col_span - 1, cols - 1)))
                    except (IndexError, ValueError):
                        pass
    doc.add_page_break()


def _add_stamp_footers(doc, project, config, sec=None, first_form=None,
                       default_form=None):
    """Добавляет основную надпись по ГОСТ (форма 5 / 6) в нижний колонтитул.

    :param first_form: форма для первого листа раздела.
        None  — не использовать «первый лист» отдельно (все страницы — форма 6);
        FORM_TEXT_FIRST — первый лист форма 5, последующие — форма 6.
    Если first_form не задан — используются только последующие (форма 6).

    :param default_form: форма для обычных (не первых) страниц. По умолчанию
        FORM_TEXT_NEXT. Используется, когда нужен отдельный лист с формой 3/5.
    :param sec: раздел, для которого строятся колонтитулы/рамка. По умолчанию 0-й.
    """
    if sec is None:
        if not doc.sections:
            return
        sec = doc.sections[0]
    req = project.requisites
    fm: FontMap = config.font_map
    font = config.font_name
    default_form = default_form if default_form is not None else FORM_TEXT_NEXT

    from .docx_tables import (set_table_fixed_layout, set_table_grid_columns,
                              set_cell_border, set_row_height,
                              set_cell_vertical_alignment, mm_to_twips)
    from docx.shared import Mm
    from .gost import STAMP_WIDTH_MM

    stamp_w = STAMP_WIDTH_MM  # 185 мм

    # --- Нижний колонтитул (последующие страницы) — форма 6 ---
    footer = sec.footer
    footer.is_linked_to_previous = False
    _clear_footer(footer)
    _add_stamp_table(footer, stamp_w, form=default_form, req=req, fm=fm, font=font,
                     set_table_fixed_layout=set_table_fixed_layout,
                     set_table_grid_columns=set_table_grid_columns,
                     set_cell_border=set_cell_border,
                     set_row_height=set_row_height,
                     set_cell_vertical_alignment=set_cell_vertical_alignment,
                     mm_to_twips=mm_to_twips,
                     stamp9=getattr(project, "stamp_cell9", None))

    # --- Первый лист отдельно (форма 5) через different_first_page ---
    if first_form is not None:
        sec.different_first_page_header_footer = True
        first_footer = sec.first_page_footer
        first_footer.is_linked_to_previous = False
        _clear_footer(first_footer)
        _add_stamp_table(first_footer, stamp_w, form=first_form, req=req, fm=fm, font=font,
                         set_table_fixed_layout=set_table_fixed_layout,
                         set_table_grid_columns=set_table_grid_columns,
                         set_cell_border=set_cell_border,
                         set_row_height=set_row_height,
                         set_cell_vertical_alignment=set_cell_vertical_alignment,
                         mm_to_twips=mm_to_twips,
                         stamp9=getattr(project, "stamp_cell9", None))
    else:
        sec.different_first_page_header_footer = False

    # --- Рамка страницы (линия по периметру) через page border ---
    _set_page_frame(sec)


def _clear_footer(footer):
    """Удаляет все абзацы (пустой колонтитул)."""
    for p in list(footer.paragraphs):
        p._element.getparent().remove(p._element)


def _add_stamp_table(container, stamp_w_mm, form, req, fm, font,
                     set_table_fixed_layout, set_table_grid_columns,
                     set_cell_border, set_row_height,
                     set_cell_vertical_alignment, mm_to_twips, stamp9=None):
    """Строит таблицу-основную надпись в нижнем колонтитуле по ГОСТ.

    Геометрия (сетка, разбивка колонок, высоты строк) берётся из
    app.core.geometry — единый источник для Word и PDF, чтобы штампы
    соответствовали образцам Форма 3/5/6.

    :param container: футер (BlockItemContainer), куда добавляется таблица
    :param form: номер формы ГОСТ (3/5/6) либо FORM_TEXT_FIRST/FORM_TEXT_NEXT.
    :param stamp9: настройки графы 9 (models.StampCell9) — текст по центру ячейки
        либо логотип организации из PDF (вписывается в ячейку без искажений).
    """
    from docx.shared import Mm
    from .geometry import grid, form5_cells, form6_cells, STAMP_WIDTH_MM
    from .docx_tables import merge_cells

    # Нормализуем форму к номеру ГОСТ: 3 / 5 / 6.
    if form in (FORM_TEXT_FIRST, 5):
        form_num = 5
    elif form in (FORM_TEXT_NEXT, 6):
        form_num = 6
    elif form in (FORM_GRAPHIC, 3):
        form_num = 3
    else:
        form_num = 5

    g = grid(form_num)
    vlines = g["vlines"]
    hlines = g["hlines"]
    cells = g["cells"]
    ncols = len(vlines) - 1
    nrows = len(hlines) - 1

    # Ширины колонок (мм) — из сетки.
    col_mm = [vlines[i + 1] - vlines[i] for i in range(ncols)]
    row_mm = [hlines[i + 1] - hlines[i] for i in range(nrows)]

    table = container.add_table(rows=nrows, cols=ncols, width=Mm(stamp_w_mm))
    table.autofit = False
    set_table_fixed_layout(table, stamp_w_mm)
    set_table_grid_columns(table, [mm_to_twips(c) for c in col_mm])

    # Задаём высоты строк (мм).
    for ri in range(nrows):
        set_row_height(table.rows[ri], row_mm[ri])

    # Сначала объединяем ячейки, затем заполняем.
    # Ячейка: (col0, col1, row0, row1, role, key, label)
    # Отслеживаем занятые (row, col), чтобы не слить уже слитую ячейку.
    covered = set()
    merged_cells = []
    for col0, col1, row0, row1, role, key, label in cells:
        region = [(r, c) for r in range(row0, row1) for c in range(col0, col1)]
        if any((r, c) in covered for r, c in region):
            # Часть региона уже слита — пропускаем (не должно случиться).
            continue
        if len(region) > 1:
            a = table.cell(row0, col0)
            b = table.cell(row1 - 1, col1 - 1)
            try:
                a.merge(b)
            except Exception:
                pass
        for r, c in region:
            covered.add((r, c))
        merged_cells.append((row0, col0, col1, row1, role, key, label))

    # Заполняем содержимое.
    for row0, col0, col1, row1, role, key, label in merged_cells:
        cell = table.cell(row0, col0)
        set_cell_border(cell, ("top", "bottom", "left", "right"))
        set_cell_vertical_alignment(cell, "center")
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Графа 9: логотип организации вместо текста (если выбран режим «логотип»).
        cell_w_mm = sum(col_mm[col0:col1])
        cell_h_mm = sum(row_mm[row0:row1])
        if (role == "field" and key in CELL9_KEYS
                and _insert_cell9_logo(para, stamp9, cell_w_mm, cell_h_mm)):
            continue
        text, size = _stamp_cell_value(role, key, label, req, fm)
        if text:
            _set_run(para.add_run(text), font, size)


def _insert_cell9_logo(para, stamp9, cell_w_mm: float, cell_h_mm: float) -> bool:
    """Вставляет логотип организации в абзац ячейки графы 9 основной надписи.

    Возвращает True, если логотип вставлен (тогда текст в графу не выводится).
    PDF-логотип предварительно растеризуется (app.core.logo), затем вписывается
    в ячейку с сохранением пропорций (автомасштабирование под ячейку).
    """
    if stamp9 is None or not hasattr(stamp9, "use_logo") or not stamp9.use_logo():
        return False
    try:
        from .logo import render_logo_png, logo_pixel_size, fit_into
        from docx.shared import Mm, Pt
    except Exception:
        return False
    png = render_logo_png(stamp9.logo_path)
    if not png:
        return False
    try:
        pad = max(0.0, float(getattr(stamp9, "padding_mm", 0.0) or 0.0))
    except (TypeError, ValueError):
        pad = 0.0
    px_w, px_h = logo_pixel_size(stamp9.logo_path)
    fit_w, fit_h = fit_into(px_w, px_h,
                            float(cell_w_mm) - 2 * pad, float(cell_h_mm) - 2 * pad)
    if fit_w <= 0 or fit_h <= 0:
        return False
    # Убираем межстрочный интервал/отбивки абзаца, чтобы картинка не смещалась.
    try:
        pf = para.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
    except Exception:
        pass
    try:
        para.add_run().add_picture(png, width=Mm(fit_w), height=Mm(fit_h))
        return True
    except Exception:
        return False


def _stamp_cell_value(role, key, label, req, fm):
    """Возвращает (текст, кегль) для ячейки основной надписи по её роли."""
    if role == "change_hdr":
        return label, fm.stamp_columns
    if role == "change_val":
        # Значения таблицы изменений пока пустые (заполняются вручную).
        return "", fm.stamp_columns
    if role == "design":
        return req.volume_id or "", fm.stamp_volume
    if role == "sheet_hdr":
        # Надпись «Лист» в правой узкой колонке — кегль 12 (как «Стадия/Листов»).
        return label, STAGE_CELL_PT
    if role == "sheet_val":
        return "", STAGE_CELL_PT
    if role == "label":
        # Надписи «Стадия / Лист / Листов» — кегль 12.
        if label in ("Стадия", "Лист", "Листов"):
            return label, STAGE_CELL_PT
        return label, fm.stamp_columns
    if role == "sig":
        # Подпись (Фамилия) — поле с фамилией подписанта.
        surname = req and ""
        return surname, fm.stamp_columns
    if role == "field":
        # Графа 9 — наименование/различительный индекс организации: собственное
        # значение графы, иначе наименование организации (кегль как у «Стадия»).
        if key in CELL9_KEYS:
            return (req.organization or ""), CELL9_PT
        # Значение из реквизитов.
        val = {
            "stage": req.stage.value if hasattr(req.stage, "value") else str(req.stage),
            "organization": req.organization or "",
            "name": req.volume_name or "",
        }.get(key, "")
        # Значения граф «Стадия / Лист / Листов» — кегль 12.
        if key in STAGE_CELL_KEYS:
            return val, STAGE_CELL_PT
        return val, fm.stamp_columns
    return "", fm.stamp_columns


def _set_page_frame(sec):
    """Устанавливает рамку страницы (ГОСТ Р 21.101-2020) в виде линии по периметру.

    Поля рамки: левое 20 мм, верх/низ/право 5 мм — берутся из свойств раздела.
    """
    from docx.oxml.ns import qn
    sectPr = sec._sectPr
    # удаляем существующую рамку (если есть)
    for el in sectPr.findall(qn("w:pgBorders")):
        sectPr.remove(el)

    pgBorders = OxmlElement("w:pgBorders")
    # offsetFrom="text" — рамка рисуется по границе полей (ГОСТ: левое 20 мм,
    # остальные 5 мм). offsetFrom="page" с space=0 уводит линию на край страницы,
    # где Word её обрезает, поэтому рамка «не видна».
    pgBorders.set(qn("w:offsetFrom"), "text")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "8")       # ~1 pt — заметная линия рамки
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        pgBorders.append(el)

    # В OOXML w:pgBorders должен стоять ПОСЛЕ pgMar и ДО cols/docGrid.
    # Ищем точку вставки: перед первым из <cols|docGrid|...>.
    order = [qn("w:pgMar"), qn("w:pgNumType"), qn("w:pgBorders"), qn("w:lnNumType"),
             qn("w:cols"), qn("w:formProt"), qn("w:vAlign"), qn("w:noEndnote"),
             qn("w:titlePg"), qn("w:textDirection"), qn("w:bidi"), qn("w:rtlGutter"),
             qn("w:docGrid"), qn("w:printerSettings")]
    # простой способ: вставить после pgMar (или в начало, если pgMar нет)
    anchor = sectPr.find(qn("w:pgMar"))
    if anchor is not None:
        anchor.addnext(pgBorders)
    else:
        sectPr.insert(0, pgBorders)


# --------------------------------------------------------------------------
# Экспорт ОДНОГО листа в редактируемый DOCX (из HTML-предпросмотра)
# --------------------------------------------------------------------------

def build_sheet_docx(project: VolumeProject, sheet_kind: str, out_path: str,
                     config: Config | None = None) -> str:
    """Собирает ОДИН лист предпросмотра (kind) в редактируемый DOCX.

    Лист восстанавливается как настоящие абзацы/таблицы Word (текст редактируется),
    с рамкой и основной надписью по ГОСТ — аналогично сборке всего тома.

    :param sheet_kind: ключ листа предпросмотра — 'cover' | 'title' | 'content'
        | 'changes' | 'stamp_f1' | 'stamp_f2' | 'stamp_f2a'.
    :param out_path: путь к создаваемому .docx.
    """
    config = config or Config()
    fm: FontMap = config.font_map
    font = config.font_name

    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    # Раздел по умолчанию: A4, поля ГОСТ
    _set_section(A4_MM[0], A4_MM[1], config.margins, doc.sections[0])
    sec = doc.sections[0]

    if sheet_kind == "cover":
        _add_cover(doc, project, fm, font)
        _no_stamp_section(sec)
    elif sheet_kind == "title":
        _add_title_sheet(doc, project, fm, font)
        _no_stamp_section(sec)
    elif sheet_kind == "content":
        _add_content(doc, project, fm, font)
        _add_stamp_footers(doc, project, config, sec, first_form=FORM_TEXT_FIRST)
    elif sheet_kind == "changes":
        _add_changes_sheet(doc, project, fm, font)
        _add_stamp_footers(doc, project, config, sec, first_form=FORM_TEXT_FIRST)
    elif sheet_kind in ("stamp_f1", "stamp_f2", "stamp_f2a"):
        # Отдельный лист с основной надписью (форма 3/5/6 ГОСТ).
        form = {"stamp_f1": FORM_GRAPHIC, "stamp_f2": FORM_TEXT_FIRST,
                "stamp_f2a": FORM_TEXT_NEXT}.get(sheet_kind, FORM_TEXT_FIRST)
        _add_stamp_sheet(doc, project, config, sec, form=form)
    else:
        # Неизвестный лист — просто текст заголовка.
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run(str(sheet_kind)), font, fm.content_header, bold=True)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    return out_path


def _add_changes_sheet(doc, project, fm, font):
    """Лист регистрации изменений — таблица записей + заголовок."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run("ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ"), font, fm.content_header, bold=True)

    cols = ["Изм.", "Дата", "Содержание изменения", "Раздел/Лист",
            "Исполнитель", "Статус"]
    rows = len(project.changes) or 1
    t = doc.add_table(rows=rows + 1, cols=len(cols))
    for ci, label in enumerate(cols):
        cell = t.cell(0, ci)
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(para.add_run(label), font, fm.stamp_columns, bold=True)
    for ri, c in enumerate(project.changes):
        values = [str(c.number), c.date, c.description, c.target,
                  c.signer_surname, c.status.value if hasattr(c.status, "value") else str(c.status)]
        for ci, v in enumerate(values):
            cell = t.cell(ri + 1, ci)
            cell.text = ""
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci == 2 else WD_ALIGN_PARAGRAPH.CENTER
            _set_run(para.add_run(v), font, fm.stamp_columns)
    doc.add_page_break()


def _add_stamp_sheet(doc, project, config, sec, form):
    """Отдельный лист с основной надписью (форма 3/5/6) — как таблица внизу."""
    from docx.shared import Mm
    # Заголовок-метка формы
    label = {FORM_GRAPHIC: "ФОРМА 3 — ГРАФИЧЕСКИЙ ЛИСТ",
             FORM_TEXT_FIRST: "ФОРМА 5 — ПЕРВЫЙ ЛИСТ ТЕКСТА",
             FORM_TEXT_NEXT: "ФОРМА 6 — ПОСЛЕДУЮЩИЙ ЛИСТ ТЕКСТА"}.get(form, "")
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run(label), config.font_name, config.font_map.content_header, bold=True)
    # Рамка и основная надпись — в нижнем колонтитуле
    _add_stamp_footers(doc, project, config, sec, first_form=None, default_form=form)


# --------------------------------------------------------------------------
# Экспорт РАЗДЕЛА тома в редактируемый DOCX
# --------------------------------------------------------------------------

def build_section_docx(project: VolumeProject, section_kind: str, out_path: str,
                       config: Config | None = None) -> str:
    """Собирает выбранный раздел тома целиком в редактируемый DOCX.

    :param section_kind: ключ раздела — 'content' | 'changes' |
        DocKind.POYASNITELNAYA.value | DocKind.GRAPHIC.value | 'appendix'.
    :param out_path: путь к создаваемому .docx.
    """
    config = config or Config()
    fm: FontMap = config.font_map
    font = config.font_name
    req = project.requisites

    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    _set_section(A4_MM[0], A4_MM[1], config.margins, doc.sections[0])
    sec = doc.sections[0]

    if section_kind == "content":
        _add_content(doc, project, fm, font)
        _add_stamp_footers(doc, project, config, sec, first_form=FORM_TEXT_FIRST)
    elif section_kind == "cover":
        _add_cover(doc, project, fm, font)
        _no_stamp_section(sec)
    elif section_kind == "title":
        _add_title_sheet(doc, project, fm, font)
        _no_stamp_section(sec)
    elif section_kind == "changes":
        _add_changes_sheet(doc, project, fm, font)
        _add_stamp_footers(doc, project, config, sec, first_form=FORM_TEXT_FIRST)
    elif section_kind in (DocKind.POYASNITELNAYA.value, DocKind.GRAPHIC.value):
        # Пояснительная записка / Графическая часть — все документы раздела.
        dkind = DocKind(section_kind)
        if dkind == DocKind.POYASNITELNAYA:
            # ПЗ: первый лист — форма 5, последующие — форма 6.
            # Раздел собирается отдельно — нумерация листов начинается с 1.
            _add_pz_toc(doc, project, fm, font, standalone=True)
            _add_section_docs(doc, project, dkind, fm, font)
            _add_stamp_footers(doc, project, config, sec, first_form=FORM_TEXT_FIRST)
        else:
            # Графическая часть: документы как есть (листы уже со штампом форма 3).
            _add_section_docs(doc, project, dkind, fm, font)
    elif section_kind == "appendix":
        _add_appendix_section(doc, project, fm, font)
    else:
        # Неизвестный раздел — заголовок.
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run(SECTION_LABELS.get(section_kind, section_kind)),
                 font, fm.content_header, bold=True)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    return out_path


def _add_section_docs(doc, project, dkind, fm, font):
    """Добавляет все документы раздела dkind в DOCX."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    docs = project.documents_for(dkind)
    if not docs:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run("Документы отсутствуют"), font, fm.content_row)
        return
    for idx, d in enumerate(docs):
        if d.path and os.path.exists(d.path):
            _add_pz_doc(doc, d.path, fm, font)
        else:
            p = doc.add_paragraph()
            _set_run(p.add_run(os.path.basename(d.path) if d.path else "—"), font, fm.content_row)
        # Разрыв между документами (кроме последнего).
        if idx < len(docs) - 1:
            doc.add_page_break()


def _add_appendix_section(doc, project, fm, font):
    """Добавляет раздел приложений: заголовки + описание каждого приложения."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    apps = project.appendices_needed()
    if not apps:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run("Приложения отсутствуют"), font, fm.content_row)
        return
    for a in apps:
        title = f"Приложение {a.number}" if a.number else "Приложение"
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run(title), font, fm.appendix_header, bold=True)
        if a.name:
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_run(p.add_run(a.name), font, fm.content_row)
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _set_run(p.add_run(f"Файл: {a.file_path or '—'}"), font, fm.content_row)
        doc.add_page_break()
