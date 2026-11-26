# -*- coding: utf-8 -*-
"""Подготовка DOCX и получение раскладки страниц Microsoft Word."""
from __future__ import annotations

import glob
import hashlib
import json
import os
import re
import tempfile


def heading_emphasis(paragraph, is_heading: bool) -> tuple:
    """Начертание абзаца в «нормализованной» копии: (полужирный, курсив).

    Повторяет правило вывода (app.core.pz_toc.heading_emphasis): заголовок
    печатается полужирным (курсив — как в исходнике), обычный текст записки —
    единым, без начертаний. Заголовок — только абзац стиля Word «Heading N»
    (см. app.core.pz_toc.is_heading_paragraph, по которому строится оглавление).

    Нужно потому, что начертание влияет на ширину строк: копия, по которой Word
    размечает листы, должна совпадать с выводом, иначе число листов разойдётся.
    """
    if not is_heading:
        return False, False
    return True, any(bool(run.italic) for run in paragraph.runs)


def _normalized_copy(path: str, first_form=2) -> str:
    from docx import Document
    from docx.shared import Mm, Pt
    from ..config import A4_MM, load_config
    from .docx_reader import paragraph_formats, style_name_of
    from .preview_html import FORM_HEIGHT, pz_doc_first_forms, pz_line_mm
    from .pz_format import PT_PER_MM

    config = load_config()
    # Нужны только форматы абзацев: полный разбор документа (тексты, таблицы)
    # здесь не используется, а стоит 0.5–0.9 с на записке в 240 абзацев.
    formats = paragraph_formats(path) or {}
    source = Document(path)
    # Кэш «styleId -> имя стиля»: разрешение имени стоит ~1 мс на абзац, тогда
    # как различных стилей в записке единицы (см. docx_reader.style_name_of).
    style_cache: dict = {}
    section = source.sections[0]
    section.different_first_page_header_footer = False
    section.page_width = Mm(A4_MM[0])
    section.page_height = Mm(A4_MM[1])
    section.left_margin = Mm(20)
    section.right_margin = Mm(5)
    section.top_margin = Mm(23)
    stamp_h = FORM_HEIGHT["2a"]
    section.bottom_margin = Mm(5 + stamp_h + 2 + 6)

    fonts = config.text_fonts
    body_size = float(fonts.body_size)
    heading_size = float(fonts.heading_size)
    table_size = float(fonts.table_size)
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    # Автоперенос ОТКЛЮЧАЕТСЯ в копии, по которой Word размечает страницы.
    # В исходниках ПЗ он обычно включён: Word рвёт слова по слогам
    # («мало-мобильных» → «мало-» + «мобильных») и укладывает абзац в меньшее
    # число строк, чем браузер, который переносит только по словам. Из-за этого
    # на листе, размеченном Word, браузер рисовал лишнюю строку — последний
    # абзац не помещался и обрезался (текст не доходил до штампа). Без
    # автопереноса Word режет строки так же, как их отображает HTML.
    settings = source.settings.element
    for tag in ("w:autoHyphenation", "w:doNotHyphenateCaps"):
        element = settings.find(qn(tag))
        if element is not None:
            settings.remove(element)
    auto_hyphen = OxmlElement("w:autoHyphenation")
    auto_hyphen.set(qn("w:val"), "0")
    settings.append(auto_hyphen)

    def mark_paragraph(paragraph, name: str, bookmark_id: int,
                       *, start_only: bool = False, end_only: bool = False) -> None:
        if not end_only:
            start = OxmlElement("w:bookmarkStart")
            start.set(qn("w:id"), str(bookmark_id))
            start.set(qn("w:name"), name)
            paragraph._p.insert(0, start)
        if not start_only:
            end = OxmlElement("w:bookmarkEnd")
            end.set(qn("w:id"), str(bookmark_id))
            paragraph._p.append(end)

    original_paragraphs = list(source.paragraphs)
    first_content_index = next(
        (index for index, paragraph in enumerate(original_paragraphs)
         if paragraph.text.strip()), None)
    if first_content_index is not None:
        paragraph = original_paragraphs[first_content_index]
        first_fmt = formats.get(first_content_index)
        has_leading_page_break = (
            paragraph.paragraph_format.page_break_before
            or bool(getattr(first_fmt, "page_break_before", False))
            or any(br.get(qn("w:type")) == "page"
                   for br in paragraph._p.iter(qn("w:br"))))
        if first_form == 2 and has_leading_page_break:
            leading_break = source.add_paragraph()
            leading_break.paragraph_format.space_before = Pt(0)
            leading_break.paragraph_format.space_after = Pt(0)
            extra_top_mm = FORM_HEIGHT[2] - FORM_HEIGHT["2a"]
            leading_break.paragraph_format.line_spacing = Pt(extra_top_mm * 72 / 25.4)
            leading_break._p.get_or_add_pPr().append(OxmlElement("w:pageBreakBefore"))
            paragraph._p.addprevious(leading_break._p)
        if has_leading_page_break:
            paragraph.paragraph_format.page_break_before = False
            ppr = paragraph._p.pPr
            if ppr is not None:
                page_break = ppr.find(qn("w:pageBreakBefore"))
                if page_break is not None:
                    ppr.remove(page_break)
            for break_element in list(paragraph._p.iter(qn("w:br"))):
                if break_element.get(qn("w:type")) == "page":
                    break_element.getparent().remove(break_element)
    for index, paragraph in enumerate(original_paragraphs):
        fmt = formats.get(index)
        if not paragraph.text.strip() and (first_content_index is None or index < first_content_index):
            paragraph._element.getparent().remove(paragraph._element)
            continue
        if first_content_index is not None and index == first_content_index and first_form != 2:
            paragraph.paragraph_format.page_break_before = False
            ppr = paragraph._p.pPr
            if ppr is not None:
                page_break = ppr.find(qn("w:pageBreakBefore"))
                if page_break is not None:
                    ppr.remove(page_break)
            for break_element in list(paragraph._p.iter(qn("w:br"))):
                if break_element.get(qn("w:type")) == "page":
                    break_element.getparent().remove(break_element)
        # Метка ставится и ПУСТОМУ абзацу: Word отводит ему строку на листе, и
        # без метки неизвестно, где он стоит — предпросмотр не мог бы
        # воспроизвести его высоту, и низ листа оставался пустым.
        mark_paragraph(paragraph, f"MIMO_P{index}", index)
        style = style_name_of(paragraph, style_cache).lower()
        is_heading = style.startswith(("heading", "заголовок")) or "heading" in style
        group = "heading" if is_heading else "body"
        size = heading_size if is_heading else body_size
        for run in paragraph.runs:
            run.font.name = fonts.font(group)
            run.font.size = Pt(size)
            # Начертание — по тому же правилу, что и в выводе: у заголовков
            # полужирный, у обычного текста начертания сняты (он печатается
            # единым). Копия должна мерить те же строки, что рисует вывод.
            bold, italic = heading_emphasis(paragraph, is_heading)
            run.font.bold = bold
            run.font.italic = italic
        pf = paragraph.paragraph_format
        if fmt and fmt.has_list:
            pf.left_indent = Mm(float(6.0 + fmt.marker_mm + fmt.marker_width_mm))
            pf.first_line_indent = Mm(-float(fmt.marker_width_mm))
        else:
            pf.left_indent = Mm(float(6.0 + (getattr(fmt, "left_mm", 0) or 0)))
            pf.first_line_indent = Mm(float(getattr(fmt, "first_line_mm", 0) or 0)) if fmt else None
        pf.right_indent = Mm(float(6.0 + (getattr(fmt, "right_mm", 0) or 0)))
        if fmt:
            pf.space_before = Pt(float(fmt.space_before_pt or 0))
            pf.space_after = Pt(float(fmt.space_after_pt or 0))
            if fmt.line_spacing_pt:
                pf.line_spacing = Pt(float(fmt.line_spacing_pt))
            elif fmt.line_spacing:
                pf.line_spacing = float(fmt.line_spacing)
            if fmt.page_break_before and not (first_content_index is not None
                                                and index == first_content_index
                                                and first_form != 2):
                pf.page_break_before = True
            if fmt.has_list:
                text = paragraph.text
                if paragraph.runs:
                    paragraph.runs[0].text = f"{fmt.marker_text} {paragraph.runs[0].text}"
                else:
                    paragraph.add_run(f"{fmt.marker_text} {text}")
        pf.widow_control = False
        # Высота строки задаётся ЯВНО и берётся из того же источника, что
        # рисует HTML (preview_html.pz_line_mm). Без этого Word размечал
        # страницы по своему интервалу (≈7.8 мм при 12 pt), а браузер рисовал
        # ту же строку 5.0 мм — листы, нарезанные Word, заполнялись лишь на
        # две трети, и низ листа оставался пустым до самого штампа.
        line_mm = pz_line_mm(size, fmt)
        pf.line_spacing = Pt(line_mm * PT_PER_MM)
        # Заголовок не отрывается от следующего за ним абзаца: иначе Word
        # оставлял заголовок последней строкой листа.
        pf.keep_with_next = is_heading

    bookmark_id = len(original_paragraphs)
    from .preview_html import table_col_widths_mm
    from .docx_tables import (set_table_fixed_layout, set_table_grid_columns,
                              set_table_cell_margins, mm_to_twips)
    text_width_mm = A4_MM[0] - 20.0 - 5.0 - 12.0
    for table_index, table in enumerate(source.tables):
        widths = table_col_widths_mm(path, table_index, len(table.columns))
        widths = [w * text_width_mm / sum(widths) for w in widths]
        set_table_fixed_layout(table, text_width_mm)
        set_table_grid_columns(table, [mm_to_twips(w) for w in widths])
        set_table_cell_margins(table, left_mm=2.0, right_mm=2.0,
                               top_mm=1.0, bottom_mm=1.0)
        for row_index, row in enumerate(table.rows):
            row_pr = row._tr.get_or_add_trPr()
            cant_split = OxmlElement("w:cantSplit")
            row_pr.append(cant_split)
            paragraphs = [p for cell in row.cells for p in cell.paragraphs]
            nonempty = [p for p in paragraphs if p.text.strip()]
            if nonempty:
                name = f"MIMO_T{table_index}R{row_index}"
                mark_paragraph(nonempty[0], name, bookmark_id, start_only=True)
                mark_paragraph(nonempty[-1], name, bookmark_id, end_only=True)
                bookmark_id += 1
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = fonts.table_font
                        run.font.size = Pt(table_size)
                    # Высота строки ячейки задаётся ЯВНО — как CSS
                    # `.dt .dt-cell { line-height:1.3 }`. Без этого Word брал
                    # свой интервал (≈8.8 мм при кегле 10 pt против 6.6 мм в
                    # HTML), в лист влезало меньше строк таблицы, и разрывы
                    # расходились с предпросмотром.
                    cell_fmt = paragraph.paragraph_format
                    cell_fmt.line_spacing = Pt(1.3 * table_size)
                    cell_fmt.space_before = Pt(0)
                    cell_fmt.space_after = Pt(0)

    _, directory, cache_path, _ = _cache_paths(path, first_form)
    os.makedirs(directory, exist_ok=True)
    temp = tempfile.NamedTemporaryFile(prefix=".tom_pz_pagination_tmp_", suffix=".docx",
                                       dir=directory, delete=False)
    temp.close()
    try:
        source.save(temp.name)
        os.replace(temp.name, cache_path)
    finally:
        if os.path.exists(temp.name):
            os.remove(temp.name)
    return cache_path


_LAYOUT_CACHE: dict = {}
_CACHE_VERSION = 14


def _cache_identity(path: str, first_form):
    from ..config import load_config
    stat = os.stat(path)
    fonts = load_config().text_fonts
    return {
        "version": _CACHE_VERSION,
        "source": os.path.abspath(path),
        "source_mtime_ns": stat.st_mtime_ns,
        "source_size": stat.st_size,
        "first_form": first_form,
        "fonts": [fonts.body_font, fonts.body_size, fonts.heading_font,
                  fonts.heading_size, fonts.table_font, fonts.table_size],
        "page": "A4:20,5,23,stamp+8",
    }


def _cache_paths(path: str, first_form):
    identity = _cache_identity(path, first_form)
    folder = os.path.splitext(os.path.abspath(path))[0] + ".pagination"
    return identity, folder, os.path.join(folder, "normalized.docx"), os.path.join(folder, "layout.json")


def _clear_legacy_cache_files(path: str) -> None:
    stem = os.path.splitext(os.path.abspath(path))[0]
    for pattern in (stem + ".pagination-*.docx", stem + ".pagination-*.json"):
        for old_path in glob.glob(pattern):
            try:
                os.remove(old_path)
            except OSError:
                pass


def _cache_key(path: str, first_form):
    identity, _, _, _ = _cache_paths(path, first_form)
    return (identity["source"], identity["source_mtime_ns"], identity["source_size"],
            identity["first_form"], *identity["fonts"])


def cached_docx_pages(path: str, first_form=2) -> dict | None:
    try:
        key = _cache_key(path, first_form)
    except OSError:
        return None
    cached = _LAYOUT_CACHE.get(key)
    if cached is not None:
        return cached
    try:
        identity, _, normalized_path, manifest_path = _cache_paths(path, first_form)
        with open(manifest_path, "r", encoding="utf-8") as stream:
            manifest = json.load(stream)
        if manifest.get("identity") != identity or not os.path.isfile(normalized_path):
            return None
        layout = manifest["layout"]
        _LAYOUT_CACHE[key] = layout
        return layout
    except (OSError, ValueError, KeyError, TypeError):
        return None


def store_docx_pages(path: str, first_form, layout: dict) -> None:
    try:
        key = _cache_key(path, first_form)
        _LAYOUT_CACHE[key] = layout
        identity, _, normalized_path, manifest_path = _cache_paths(path, first_form)
        if os.path.isfile(normalized_path):
            with open(manifest_path, "w", encoding="utf-8") as stream:
                json.dump({"identity": identity, "layout": layout}, stream,
                          ensure_ascii=False, separators=(",", ":"))
    except (OSError, TypeError, ValueError):
        pass


def invalidate_docx_pages(path: str | None = None) -> None:
    if path is None:
        _LAYOUT_CACHE.clear()
        return
    target = os.path.abspath(path)
    for key in list(_LAYOUT_CACHE):
        if key[0] == target:
            _LAYOUT_CACHE.pop(key, None)


def calculate_docx_pages(path: str, first_form=2, word=None) -> dict:
    """Загружает карту из дискового кэша или рассчитывает её через Word.

    :param word: уже открытый ``Word.Application`` (см. ``word_com.paginate_docx``).
        Позволяет посчитать несколько документов подряд на одном экземпляре
        Word: запуск приложения стоит ~2 с, и на каждом следующем документе
        это время тратилось бы заново.
    """
    if not path or not os.path.isfile(path):
        raise FileNotFoundError(path or "Не задан путь к DOCX")
    from .word_com import paginate_docx

    cached = cached_docx_pages(path, first_form)
    if cached is not None:
        return cached
    normalized = _normalized_copy(path, first_form)
    result = paginate_docx(normalized, word=word)
    identity, _, _, manifest_path = _cache_paths(path, first_form)
    result["source_mtime_ns"] = identity["source_mtime_ns"]
    result["source_size"] = identity["source_size"]
    key = _cache_key(path, first_form)
    _LAYOUT_CACHE[key] = result
    with open(manifest_path, "w", encoding="utf-8") as stream:
        json.dump({"identity": identity, "layout": result}, stream,
                  ensure_ascii=False, separators=(",", ":"))
    _clear_legacy_cache_files(path)
    return result
