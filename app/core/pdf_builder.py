# -*- coding: utf-8 -*-
"""
Сборка полного тома: текстовая часть (титул+содержание+ПЗ+лист изменений) +
графические листы + приложения в единый PDF, с рамками/осн. надписью по ГОСТ,
сквозной нумерацией листов и закладками.
"""
from __future__ import annotations

import os
import tempfile

from ..models import (
    VolumeProject, DocKind,
    SECTION_COVER, SECTION_TITLE, SECTION_CONTENT, SECTION_PZ,
    SECTION_GRAPHIC, SECTION_APPENDIX, SECTION_CHANGES, SECTION_COMPOSITION,
)
from ..config import Config
from .gost import (FORM_GRAPHIC, FORM_TEXT_FIRST, FORM_TEXT_NEXT,
                   A4_W_PT, A4_H_PT, A4_W_MM, A4_H_MM)
from .pdf_merge import count_pages, compute_total_sheets
from .renderer import draw_frame_and_stamp, draw_sheet_number_box
from .toc import content_rows as toc_content_rows
# Единый источник текстовых kinds (не дублировать — риск расхождения).
from .docx_exporter import TEXT_SECTION_KINDS

#: kinds, попадающие в текстовый DOCX/PDF-сегмент (graphic/appendix — отдельно)
_SEGMENT_TEXT = frozenset(TEXT_SECTION_KINDS)
_SEGMENT_GRAPHIC = frozenset({SECTION_GRAPHIC})
_SEGMENT_APPENDIX = frozenset({SECTION_APPENDIX})


def structure_segments(project: VolumeProject) -> list:
    """Группирует project.structure в сегменты для PDF-сборки.

    Соседние текстовые kinds (cover/title/content/pz/changes) склеиваются в один
    сегмент ("text", kinds); graphic и appendix — каждый свой сегмент.
    Порядок сегментов = порядку structure (после фильтра needed).

    Пример по умолчанию:
        [("text", [cover,title,content,pz]), ("graphic", []),
         ("appendix", []), ("text", [changes])]
    """
    segments = []
    cur_text: list = []

    def _flush_text():
        nonlocal cur_text
        if cur_text:
            segments.append(("text", list(cur_text)))
            cur_text = []

    for entry in project.structure or []:
        if not entry.get("needed", True):
            continue
        kind = entry.get("kind") or ""
        if kind in _SEGMENT_TEXT:
            # changes без журнала не создаёт сегмент
            if kind == SECTION_CHANGES and not project.changes:
                continue
            cur_text.append(kind)
        elif kind in _SEGMENT_GRAPHIC:
            _flush_text()
            segments.append(("graphic", [kind]))
        elif kind in _SEGMENT_APPENDIX:
            _flush_text()
            segments.append(("appendix", [kind]))
        # неизвестные kinds игнорируются
    _flush_text()
    return segments


def _segment_start_sheet(project: VolumeProject, segment_kinds: list) -> int:
    """Первый номер листа сегмента по карте toc.build_layout."""
    from .toc import build_layout
    lay = build_layout(project)
    for it in lay.items:
        if it.section_key in segment_kinds and it.count > 0 and it.first_sheet >= 1:
            return it.first_sheet
    return 1


def _text_pdf(project: VolumeProject, cfg: Config, out_dir: str,
              use_word: bool = True,
              kinds: list | None = None,
              tag: str = "") -> tuple[str | None, bool, list]:
    """Формирует PDF текстовой части.

    Если use_word и доступен Word — конвертируем DOCX через Word COM.
    DOCX уже содержит рамку и основную надпись (форма 5/6), поэтому
    дополнительно штамповать НЕ нужно.

    Иначе текстовая часть собирается напрямую через PyMuPDF (fallback) и
    требует последующего наложения штампов.

    :param kinds: подмножество текстовых разделов (для сегментации PDF);
        None — все текстовые разделы из project.structure.
    :param tag: метка для уникальности временных файлов (сегментация PDF).

    Возвращает (путь_к_pdf | None, пришёл_из_word, page_forms).
    page_forms — список форм основной надписи по каждой странице текстовой части:
        None — без рамки и штампа (обложка);
        2    — большой штамп (первый/заглавный лист раздела);
        '2a' — малый штамп (последующий лист).
    """
    # Собираем DOCX текстовой части (всегда доступно)
    from .docx_exporter import build_text_docx
    suffix = f"_{tag}" if tag else ""
    docx_path = os.path.join(out_dir, f"_text_part{suffix}.docx")
    build_text_docx(project, docx_path, cfg, kinds=kinds)

    # 1) Если разрешён Word и он доступен — конвертируем DOCX -> PDF (максимальная точность)
    if use_word:
        try:
            from .word_com import word_available, docx_to_pdf
            if word_available():
                pdf_path = os.path.join(out_dir, f"_text_part{suffix}.pdf")
                docx_to_pdf(docx_path, pdf_path)
                return pdf_path, True, []
        except Exception:
            pass

    # 2) Fallback: собрать текстовые страницы напрямую через PyMuPDF (без Word)
    pdf_path, page_forms = _text_pdf_pymupdf(project, out_dir, cfg,
                                             kinds=kinds, tag=tag)
    return pdf_path, False, page_forms


def _add_toc_pages_pymupdf(project, new_page, fname, fm, forms,
                           offset: int | None = None) -> int:
    """Рисует листы автоматического оглавления ПЗ (запасной путь без Word).

    Строки оглавления берутся из единого источника (app.core.pz_toc) по уровням
    текста исходного DOCX; отточие воспроизводится точками, номер листа —
    справа. Листы оглавления идут ПЕРВЫМИ листами записки, поэтому первый из
    них несёт большой штамп (форма 5), остальные — малый (форма 6).

    :param new_page: функция создания страницы new_page(form, width, height).
    :param forms: (форма первого листа, форма последующих листов).
    :param offset: сдвиг номеров листов (None — посчитать по структуре тома).
    :return: число добавленных листов (0 — оглавление не выводится).
    """
    from .pz_toc import (paginate_items, project_toc_items, toc_page_offset,
                         toc_pdf_lines, toc_text_style)
    items, _path = project_toc_items(project)
    if not items:
        return 0
    # Кегль строк — как у обычного текста ПЗ (тот же источник, что в DOCX/HTML).
    _font_name, body_size = toc_text_style()
    pages = paginate_items(items, body_size)
    if offset is None:
        offset = toc_page_offset(project)
    first_form, next_form = forms
    for i, page_items in enumerate(pages):
        pg = new_page(first_form if i == 0 else next_form)
        y = 90
        for j, line in enumerate(toc_pdf_lines(page_items, offset=offset)):
            size = fm.content_header if j == 0 else body_size
            pg.insert_text((72, y), line, fontname=fname, fontsize=size)
            y += 20
    return len(pages)


def _register_text_font(doc):
    """Регистрирует шрифт ISOCPEUR (кириллица) в PDF-документе.

    Возвращает имя шрифта для insert_text; при недоступности — 'helv'.

    Шрифт регистрируется НА СТРАНИЦЕ (page.insert_font): у объекта Document
    PyMuPDF метода insert_font нет, и прежний вызов молча падал в 'helv' —
    тогда кириллица выводилась точками. Поэтому имя шрифта запоминается и
    подставляется на каждой новой странице (см. _ensure_page_font).
    """
    from .fonts import font_file
    fpath = font_file()
    if not fpath:
        return "helv"
    try:
        with open(fpath, "rb") as f:
            f.read(1)
    except Exception:
        return "helv"
    return "ISOCPEUR"


def _ensure_page_font(page, fontname: str) -> str:
    """Регистрирует ISOCPEUR на конкретной странице (иначе PyMuPDF не знает шрифт)."""
    if fontname != "ISOCPEUR":
        return fontname
    from .fonts import register_font
    return register_font(page, "ISOCPEUR") or "helv"


#: Псевдонимы шрифтов текста ПЗ по группам (PyMuPDF требует ASCII-имя).
_TEXT_FONT_ALIAS = {"body": "TomBody", "heading": "TomHead", "table": "TomTable"}


def _text_font_files(cfg) -> dict:
    """Файл шрифта для каждой группы текста ПЗ (обычный текст/заголовки/таблицы).

    Берётся из настроек оформления (config.TextFonts); если выбранного шрифта
    нет в системе — используется основной шрифт проекта, чтобы кириллица не
    выводилась подставным начертанием.
    """
    from .fonts import font_file
    tf = getattr(cfg, "text_fonts", None)
    base = font_file()
    files = {}
    for group in _TEXT_FONT_ALIAS:
        family = tf.font(group) if tf is not None else ""
        files[group] = (font_file(family) if family else "") or base
    return files


def _register_group_fonts(page, files: dict) -> dict:
    """Регистрирует шрифты групп на странице; возвращает имя шрифта по группам."""
    aliases = {}
    for group, alias in _TEXT_FONT_ALIAS.items():
        path = files.get(group) or ""
        if not path or not os.path.exists(path):
            aliases[group] = "helv"
            continue
        try:
            page.insert_font(fontname=alias, fontfile=path)
            aliases[group] = alias
        except Exception:
            aliases[group] = "helv"
    return aliases


def _text_pdf_pymupdf(project, out_dir, cfg, kinds: list | None = None,
                      tag: str = ""):
    """Собирает текстовую часть как PDF-страницы напрямую через PyMuPDF —
    запасной путь, когда Word COM недоступен.

    :param kinds: подмножество текстовых разделов в порядке следования
        (None — все текстовые разделы из project.structure).
    :param tag: метка для уникальности временного файла.

    Возвращает (путь_к_pdf, page_forms). page_forms[i] — форма осн. надписи для
    страницы i: None (обложка — без рамки/штампа), FORM_TEXT_FIRST (первый лист
    раздела — большой штамп), FORM_TEXT_NEXT (последующий лист — малый штамп)."""
    import pymupdf
    suffix = f"_{tag}" if tag else ""
    pdf_path = os.path.join(out_dir, f"_text_part{suffix}.pdf")
    doc = pymupdf.open()
    fm = cfg.font_map
    req = project.requisites
    page_forms = []         # форма по каждой странице текстовой части

    if kinds is None:
        from .docx_exporter import _structure_text_kinds
        kinds = _structure_text_kinds(project)
        if SECTION_CHANGES in kinds and not project.changes:
            kinds = [k for k in kinds if k != SECTION_CHANGES]

    # Шрифт с поддержкой кириллицы (ISOCPEUR, если доступен).
    fname = _register_text_font(doc)
    # Шрифты групп текста ПЗ из настроек (обычный текст/заголовки/таблицы):
    # встраиваются на каждой странице, имя по группе — в page_fonts.
    group_files = _text_font_files(cfg)
    page_fonts = {"body": fname, "heading": fname, "table": fname}

    # Хелпер: добавить страницу и зафиксировать её форму основной надписи.
    # Шрифт регистрируется на КАЖДОЙ странице: PyMuPDF хранит встроенный шрифт
    # в ресурсах страницы, поэтому без этого кириллица выводилась бы точками.
    def new_page(form, width=A4_W_PT, height=A4_H_PT):
        nonlocal page_forms
        pg = doc.new_page(width=width, height=height)
        page_forms.append(form)
        _ensure_page_font(pg, fname)
        if group_files:
            page_fonts.update(_register_group_fonts(pg, group_files))
        return pg

    for kind in kinds:
        if kind == SECTION_COVER:
            # Обложка: без рамки и без штампа
            pg = new_page(None)
            pg.insert_text((72, 100), req.organization or "", fontname=fname, fontsize=18)
            pg.insert_text((72, 140), "ОБЛОЖКА", fontname=fname, fontsize=14)
            pg.insert_text((72, 180), req.volume_id or "", fontname=fname, fontsize=14)
        elif kind == SECTION_TITLE:
            # Титульный лист — без рамки и штампа, form=None.
            pg = new_page(None)
            pg.insert_text((72, 100), req.volume_name or "", fontname=fname, fontsize=16)
            pg.insert_text((72, 140), "ТИТУЛЬНЫЙ ЛИСТ", fontname=fname, fontsize=14)
            pg.insert_text((72, 180), req.object_name or "", fontname=fname, fontsize=12)
        elif kind == SECTION_CONTENT:
            # Содержание — первый лист форма 5, №1 сквозной нумерации.
            pg = new_page(FORM_TEXT_FIRST)
            pg.insert_text((72, 90), "СОДЕРЖАНИЕ", fontname=fname, fontsize=14)
            y = 120
            for name, des, sheets in toc_content_rows(project):
                line = f"{name}"
                if des:
                    line += f"  {des}"
                if sheets:
                    line += f"  {sheets}"
                pg.insert_text((72, y), line, fontname=fname, fontsize=11)
                y += 20
        elif kind == SECTION_COMPOSITION:
            # Состав проектной документации: надпись и таблица томов.
            # Разбиение на листы — единый источник (composition_pages):
            # первый лист несёт форму 5, последующие — форму 6.
            pages = _composition_pages(project)
            for index, page_rows in enumerate(pages):
                pg = new_page(FORM_TEXT_FIRST if index == 0 else FORM_TEXT_NEXT)
                _draw_composition_page(pg, page_rows, fname, cfg)
        elif kind == SECTION_PZ:
            # Оглавление ПЗ — первыми листами раздела.
            from .docx_reader import read_docx
            from .pz_toc import toc_page_offset
            toc_pages = _add_toc_pages_pymupdf(project, new_page, fname, fm,
                                               (FORM_TEXT_FIRST, FORM_TEXT_NEXT),
                                               offset=toc_page_offset(project))
            first_pz = not toc_pages
            for d in project.documents_for(DocKind.POYASNITELNAYA):
                if d.path and os.path.exists(d.path):
                    try:
                        content = read_docx(d.path)
                    except Exception:
                        continue
                    pg = new_page(FORM_TEXT_FIRST if first_pz else FORM_TEXT_NEXT)
                    first_pz = False
                    y = 80
                    # Шрифт и кегль текста ПЗ — из настроек (config.TextFonts):
                    # те же значения, что в предпросмотре и Word.
                    tf = getattr(cfg, "text_fonts", None)
                    body_size = tf.size("body") if tf else 11
                    head_size = tf.size("heading") if tf else 11
                    for para in content.paragraphs:
                        if not para.text.strip():
                            continue
                        fmt = getattr(para, "fmt", None)
                        group = "heading" if para.is_heading else "body"
                        size = head_size if para.is_heading else body_size
                        # Абзацный отступ (5 знаков) — из оформления абзаца: он уже
                        # учитывает исключения (заголовки и пункты списков его не
                        # получают). Отрицательный отступ (выступ списка) не сдвигает
                        # первую строку вправо.
                        first_mm = float(getattr(fmt, "first_line_mm", 0.0) or 0.0)
                        indent_pt = first_mm * 2.8346 if first_mm > 0 else 0.0
                        lines = _wrap(para.text, width=90)
                        for li, ln in enumerate(lines):
                            if y > 800:
                                pg = new_page(FORM_TEXT_NEXT)
                                y = 80
                            x = 72 + (indent_pt if li == 0 else 0)
                            pg.insert_text((x, y), ln,
                                           fontname=page_fonts.get(group, fname),
                                           fontsize=size)
                            y += size * 1.45
        elif kind == SECTION_CHANGES:
            # Лист регистрации изменений (fallback-путь).
            pg = new_page(FORM_TEXT_FIRST)
            pg.insert_text((72, 90), "ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ",
                           fontname=fname, fontsize=13)
            y = 130
            for c in project.changes:
                line = f"{c.number}  {c.date}  {c.description}"
                pg.insert_text((72, y), line, fontname=fname, fontsize=10)
                y += 18

    # Минимум одна страница, чтобы файл был валидным
    if not page_forms:
        new_page(None).insert_text((72, 100), "", fontname=fname, fontsize=11)

    doc.save(pdf_path)
    doc.close()
    return pdf_path, page_forms


def _wrap(text: str, width: int = 90) -> list:
    """Простой перенос строк по ширине (для fallback-рендера)."""
    import textwrap
    return textwrap.wrap(text, width) or [text]


def _composition_pages(project) -> list:
    """Страницы таблицы состава ПД: единый источник разбиения на листы."""
    from .preview_html import composition_pages
    return composition_pages(project)


def _draw_composition_page(pg, page_rows, fname: str, cfg) -> None:
    """Рисует лист состава ПД: надпись сверху и таблица томов.

    Геометрия — как в предпросмотре: надпись «Состав проектной документации»,
    шапка 15 мм и строки 8 мм; колонки 15 / 60 / 90 / 20 мм (сумма 185 мм —
    ширина между линиями рамки). Область начинается под квадратом сквозной
    нумерации (17 мм от верха листа).
    """
    from ..models import (COMPOSITION_HEAD_MM, COMPOSITION_LABELS,
                          COMPOSITION_ROW_MM, COMPOSITION_WIDTHS)
    from .preview_html import COMPOSITION_TITLE
    from .renderer import mm_to_pt

    left_mm, right_mm = 20.0, 5.0
    x0 = left_mm
    total_w = sum(COMPOSITION_WIDTHS)
    # Надпись над таблицей — по центру текстовой области.
    y = 19.0
    title_box = (mm_to_pt(x0), mm_to_pt(y),
                 mm_to_pt(x0 + total_w), mm_to_pt(y + 8.0))
    pg.insert_textbox(pymupdf.Rect(*title_box), COMPOSITION_TITLE,
                      fontname=fname, fontsize=13,
                      align=pymupdf.TEXT_ALIGN_CENTER)
    y += 12.0

    cols = tuple(COMPOSITION_WIDTHS)
    # Шапка таблицы.
    _draw_composition_row(pg, x0, y, cols, COMPOSITION_HEAD_MM, fname,
                          COMPOSITION_LABELS, header=True)
    y += COMPOSITION_HEAD_MM
    # Строки таблицы.
    for values in page_rows:
        _draw_composition_row(pg, x0, y, cols, COMPOSITION_ROW_MM, fname,
                              [str(v or "") for v in values])
        y += COMPOSITION_ROW_MM


def _draw_composition_row(pg, x0: float, y: float, cols, row_h: float,
                          fname: str, values, header: bool = False) -> None:
    """Рисует одну строку таблицы состава: рамка, вертикали и текст ячеек."""
    from .renderer import mm_to_pt

    total_w = sum(cols)
    pg.draw_rect((mm_to_pt(x0), mm_to_pt(y),
                  mm_to_pt(x0 + total_w), mm_to_pt(y + row_h)),
                 color=(0, 0, 0), width=0.7)
    cx = x0
    for i, value in enumerate(values):
        if i:
            pg.draw_line((mm_to_pt(cx), mm_to_pt(y)),
                         (mm_to_pt(cx), mm_to_pt(y + row_h)),
                         color=(0, 0, 0), width=0.7)
        text = str(value or "")
        size = 9 if header else 9
        # Подбор кегля по ширине ячейки: длинное наименование не выходит
        # за линию колонки (перенос в PDF-таблице не делается).
        max_chars = max(3, int(cols[i] / 2.0))
        while text and len(text) > max_chars and size > 5:
            size -= 0.5
            max_chars = max(3, int(cols[i] / (size * 0.22)))
        if len(text) > max_chars:
            text = text[:max_chars - 1] + "…"
        tx = mm_to_pt(cx + 1.5)
        pg.insert_text((tx, mm_to_pt(y + row_h - 2.4)), text,
                       fontname=fname, fontsize=size)
        cx += cols[i]


def _copy_graphic(src_pdf: str, out_pdf: str) -> None:
    """Копирует графический PDF без оформления (совместимость)."""
    import pymupdf
    doc = pymupdf.open(src_pdf)
    doc.save(out_pdf)
    doc.close()


def _graphic_sheet_list_pdf(project, cfg, out_pdf: str, sheet_no: int,
                            total: int) -> str | None:
    """Лист «Графическая часть» с перечнем листов как PDF-страница.

    Строится стандартным векторным путём (рамка + основная надпись формы 3 +
    таблица), поэтому лист совпадает с предпросмотром и Word. Строки берутся
    из единого источника (preview_html.build_graphic_sheet_list): таблица или
    из Excel, или по фактическим листам раздела.

    Возвращает путь к PDF или None, если лист отключён в настройках проекта.
    """
    from .preview_html import _graphic_sheet_list_cfg, build_graphic_sheet_list
    gl = _graphic_sheet_list_cfg(project)
    if not gl.enabled:
        return None
    import pymupdf
    from ..models import SHEETLIST_LABELS
    from .renderer import draw_frame_and_stamp, draw_sheet_number_box, mm_to_pt

    doc = pymupdf.open()
    try:
        page = doc.new_page(width=A4_W_PT, height=A4_H_PT)
        texts = {"volume": str(project.requisites.volume_id or ""),
                 "sheet": str(sheet_no),
                 "qty": str(total),
                 "signatures": _signature_texts(project, SECTION_GRAPHIC,
                                                FORM_GRAPHIC)}
        texts.update(_cell9_texts(project))
        draw_frame_and_stamp(page, A4_W_MM, A4_H_MM, form=FORM_GRAPHIC,
                             font_map=cfg.font_map, texts=texts)
        draw_sheet_number_box(page, A4_W_MM, A4_H_MM, sheet_no)

        # Заголовок, подзаголовок и таблица — в текстовой области листа
        # (левое поле 20 мм, остальные 5 мм; верх учитывает квадрат номера).
        from .fonts import register_font
        fname = register_font(page) or "helv"
        left_mm, right_mm = 20.0, 5.0
        width_mm = A4_W_MM - left_mm - right_mm
        y = 20.0
        # Заголовок и подзаголовок — по центру текстовой области
        # (insert_textbox с выравниванием CENTER, чтобы не подбирать отступ).
        for text, size, box_h in ((gl.title or "Графическая часть", 14, 8.0),
                                  (gl.subtitle, 10, 6.0)):
            if not text:
                continue
            box = pymupdf.Rect(mm_to_pt(left_mm), mm_to_pt(y),
                               mm_to_pt(left_mm + width_mm), mm_to_pt(y + box_h))
            page.insert_textbox(box, text, fontname=fname, fontsize=size,
                                align=pymupdf.TEXT_ALIGN_CENTER)
            y += box_h
        y += 2.0

        # Колонки таблицы (мм) — единый источник с предпросмотром и Word:
        # Обозначение 50 | Наименование 95 | Примечание 40.
        from ..models import SHEETLIST_WIDTHS
        cols = tuple(SHEETLIST_WIDTHS)
        row_h = 6.0
        x0 = left_mm
        rows = build_graphic_sheet_list(project)
        # Шапка таблицы.
        page.draw_rect((mm_to_pt(x0), mm_to_pt(y),
                        mm_to_pt(x0 + sum(cols)), mm_to_pt(y + row_h)),
                       color=(0, 0, 0), width=0.7)
        cx = x0
        for i, label in enumerate(SHEETLIST_LABELS):
            page.draw_line((mm_to_pt(cx), mm_to_pt(y)),
                           (mm_to_pt(cx), mm_to_pt(y + row_h)),
                           color=(0, 0, 0), width=0.7)
            page.insert_text((mm_to_pt(cx + 2), mm_to_pt(y + row_h - 2)),
                             label, fontname=fname, fontsize=9)
            cx += cols[i]
        page.draw_line((mm_to_pt(x0 + sum(cols)), mm_to_pt(y)),
                       (mm_to_pt(x0 + sum(cols)), mm_to_pt(y + row_h)),
                       color=(0, 0, 0), width=0.7)
        y += row_h
        # Строки перечня; при нехватке места таблица обрывается (лист один).
        limit_y = A4_H_MM - 5.0 - 55.0 - 3.0     # над штампом формы 3
        for designation, name, note in rows:
            if y + row_h > limit_y:
                break
            page.draw_rect((mm_to_pt(x0), mm_to_pt(y),
                            mm_to_pt(x0 + sum(cols)), mm_to_pt(y + row_h)),
                           color=(0, 0, 0), width=0.7)
            cx = x0
            for i, value in enumerate((designation, name, note)):
                if i < len(cols) - 1:
                    page.draw_line((mm_to_pt(cx), mm_to_pt(y)),
                                   (mm_to_pt(cx), mm_to_pt(y + row_h)),
                                   color=(0, 0, 0), width=0.7)
                text = str(value or "")
                # Обрезка по ширине колонки (без переноса: таблица перечня
                # коротких строк, длинные наименования режутся с отточием).
                max_chars = max(4, int(cols[i] / 2.1))
                if len(text) > max_chars:
                    text = text[:max_chars - 1] + "…"
                page.insert_text((mm_to_pt(cx + 2), mm_to_pt(y + row_h - 2)),
                                 text, fontname=fname, fontsize=9)
                cx += cols[i]
            y += row_h
        doc.save(out_pdf)
        return out_pdf
    finally:
        doc.close()


def _stamp_graphic_number_box(src_pdf: str, out_pdf: str, sheet_start: int) -> int:
    """Накладывает только квадрат со сквозным номером на исходную графику.

    Исходный PDF сохраняется в родном размере и не растеризуется; рамка и
    основная надпись чертежа не дублируются. Возвращает следующий номер листа.
    """
    import pymupdf

    doc = pymupdf.open(src_pdf)
    sheet = sheet_start
    try:
        for page in doc:
            rect = page.rect
            page_w_mm = rect.width * 25.4 / 72.0
            page_h_mm = rect.height * 25.4 / 72.0
            draw_sheet_number_box(page, page_w_mm, page_h_mm, sheet)
            sheet += 1
        doc.save(out_pdf)
    finally:
        doc.close()
    return sheet


def _cell9_texts(project) -> dict:
    """Значения графы 9 для PyMuPDF-рендера: текст организации, наименование
    или логотип (путь к PDF). Используется и для текстовых, и для графических
    листов — графа 9 есть в формах 5 и 3."""
    cell9 = getattr(project, "stamp_cell9", None)
    req = project.requisites
    org = req.organization or ""
    if cell9 is None:
        return {"organization": org}
    texts = {"organization": org}
    if cell9.use_logo():
        texts["cell9_logo"] = cell9.logo_path
        try:
            texts["cell9_pad"] = float(cell9.padding_mm)
        except (TypeError, ValueError):
            pass
    return texts


def _segment_signature_kind(seg_kinds: list) -> str:
    """Раздел, по которому берутся строки подписей для текстового сегмента.

    Текстовый сегмент склеивает соседние разделы (например содержание + ПЗ),
    поэтому берётся последний документный раздел — именно его штамп видят
    страницы записки.
    """
    for kind in reversed(list(seg_kinds or [])):
        if kind in (SECTION_PZ,):
            return kind
    for kind in reversed(list(seg_kinds or [])):
        if kind:
            return kind
    return ""


def _signature_texts(project, kind: str | None, form) -> list:
    """Строки подписей левого блока штампа для PDF-рендера.

    Значения берутся из модели (models.stamp_signature_rows) — те же, что в
    HTML-предпросмотре и Word: должность, фамилия, дата и изображение подписи
    подписанта, подобранного по фамилии.
    """
    from ..models import stamp_signature_rows
    if not kind:
        return []
    form_num = 3 if form == FORM_GRAPHIC else 5
    return stamp_signature_rows(project, kind, form_num)


def _stamp_text_pages(text_pdf: str, cfg: Config, volume_id: str, total: int,
                      page_forms: list | None = None,
                      cell9_texts: dict | None = None,
                      sheet_start: int | None = None,
                      signatures: list | None = None) -> str:
    """Накладывает рамку и основную надпись на текстовые страницы.

    page_forms[i] определяет форму основной надписи для страницы i:
        None            — не рисовать рамку и штамп (обложка);
        FORM_TEXT_FIRST — большой штамп (первый/заглавный лист раздела);
        FORM_TEXT_NEXT  — малый штамп (последующий лист).
    Если page_forms не задан — первая страница FORM_TEXT_FIRST, остальные
    FORM_TEXT_NEXT.

    :param sheet_start: первый номер листа сегмента (сквозная нумерация тома);
        None — нумерация с 1 (полный текстовый блок «от начала»).
    :param signatures: строки подписей левого блока (models.stamp_signature_rows)
        — должность, фамилия, дата и изображение подписи каждой строки.

    Возвращает путь к новому PDF (или исходный при ошибке)."""
    import pymupdf
    doc = pymupdf.open(text_pdf)
    try:
        sheet_no = sheet_start if sheet_start else 1
        for i, pg in enumerate(doc):
            rect = pg.rect
            w_mm = rect.width / 2.8346
            h_mm = rect.height / 2.8346
            if page_forms is not None:
                form = page_forms[i] if i < len(page_forms) else FORM_TEXT_NEXT
            else:
                form = FORM_TEXT_FIRST if i == 0 else FORM_TEXT_NEXT
            # обложка: без рамки и без штампа
            if form is None:
                continue
            texts = {"volume": str(volume_id),
                     "sheet": str(sheet_no),
                     "qty": str(total),
                     "signatures": signatures or []}
            texts.update(cell9_texts or {})
            draw_frame_and_stamp(pg, w_mm, h_mm, form=form,
                                 font_map=cfg.font_map, texts=texts)
            sheet_no += 1
        base = os.path.splitext(text_pdf)[0]
        out = base + "_stamped.pdf"
        doc.save(out)
        return out
    finally:
        doc.close()


def _stamp_graphic(src_pdf: str, out_pdf: str, page_w_mm: float,
                   page_h_mm: float, cfg: Config, sheet_start: int,
                   total: int, volume_id: str, cell9_texts: dict | None = None,
                   signatures: list | None = None) -> int:
    """Накладывает рамку и основную надпись (форма 3) на листы графической части.

    Нумерация листов — сквозная (начиная с sheet_start). Возвращает следующий
    свободный номер листа.

    Размер страницы берётся из каждого листа PDF (если он ненулевой), иначе
    используются переданные page_w_mm/page_h_mm (на случай нестандартной графики).

    :param signatures: строки подписей левого блока (models.stamp_signature_rows).
    """
    import pymupdf
    doc = pymupdf.open(src_pdf)
    sheet = sheet_start
    for pg in doc:
        rect = pg.rect
        # 1 pt = 25.4/72 мм
        real_w = rect.width * 0.352778
        real_h = rect.height * 0.352778
        w_mm = real_w if real_w > 1 else page_w_mm
        h_mm = real_h if real_h > 1 else page_h_mm
        texts = {"volume": str(volume_id),
                 "sheet": str(sheet),
                 "qty": str(total),
                 "signatures": signatures or []}
        texts.update(cell9_texts or {})
        draw_frame_and_stamp(pg, w_mm, h_mm, form=FORM_GRAPHIC,
                             font_map=cfg.font_map, texts=texts)
        sheet += 1
    doc.save(out_pdf)
    doc.close()
    return sheet


def _add_text_colontitules(text_pdf: str, cfg: Config, volume_id: str,
                           org_name: str = "") -> str:
    """Добавляет колонтитулы на текстовые страницы по ГОСТ Р 21.101-2020 (п. 4.1.5).

    Текстовые документы со сплошным текстом выполняются БЕЗ основных надписей,
    дополнительных граф и рамок. Вместо этого:
      — верхний колонтитул: обозначение документа (в левом углу);
      — нижний колонтитул: наименование организации + номер листа (в правом углу).

    Возвращает путь к новому PDF (или исходный при ошибке)."""
    import pymupdf
    doc = pymupdf.open(text_pdf)
    try:
        for i, pg in enumerate(doc):
            rect = pg.rect
            # верхний колонтитул: обозначение документа, слева
            pg.insert_text((rect.width * 0.05, rect.height * 0.035),
                           volume_id, fontname="helv", fontsize=9)
            # нижний колонтитул: организация слева + номер листа справа
            bottom_y = rect.height * 0.975
            if org_name:
                pg.insert_text((rect.width * 0.05, bottom_y), org_name, fontname="helv", fontsize=8)
            pg.insert_text((rect.width * 0.9, bottom_y), str(i + 1), fontname="helv", fontsize=9)
        # сохранить в новый файл
        import os as _os
        base = _os.path.splitext(text_pdf)[0]
        out = base + "_colontitules.pdf"
        doc.save(out)
        return out
    finally:
        doc.close()


def build_full_pdf(project: VolumeProject, out_path: str, cfg: Config | None = None,
                   use_word: bool = True, work_dir: str | None = None) -> str:
    """Собирает полный том в единый PDF.

    Порядок сегментов берётся из project.structure (см. structure_segments):
    соседние текстовые разделы склеиваются в один DOCX→PDF-сегмент, графика
    и приложения идут своими сегментами. Сквозная нумерация считается по
    toc.build_layout и прокидывается в квадраты номеров листов.
    """
    import pymupdf
    import logging

    log = logging.getLogger("tom_builder.build")
    cfg = cfg or Config()
    # По умолчанию рабочий каталог — внутри проекта (tmp), т.к. системный temp
    # в некоторых средах недоступен для записи zip/DOCX.
    if work_dir is None:
        proj_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        work_dir = os.path.join(proj_root, "tmp", "_build")
    os.makedirs(work_dir, exist_ok=True)

    total = compute_total_sheets(project)  # «Листов M»
    volume_id = project.requisites.volume_id
    segments = structure_segments(project)
    log.info("build_full_pdf: segments=%s total_sheets=%s",
             [(k, list(kk)) for k, kk in segments], total)

    # Начальный номер каждого сегмента — из единого источника toc.build_layout.
    # Для сегмента, начинающегося «до» первой нумерованной позиции (обложка,
    # титул), — 1.
    seg_starts: list[int] = []
    for _kind, kk in segments:
        seg_starts.append(_segment_start_sheet(project, kk))

    merged = pymupdf.open()
    segment_files: list[str | None] = []   # путь PDF каждого сегмента
    segment_meta: list[tuple] = []         # для диагностики/закладок

    for idx, ((seg_kind, seg_kinds), start_sheet) in enumerate(zip(segments, seg_starts)):
        if seg_kind == "text":
            text_pdf, from_word, page_forms = _text_pdf(
                project, cfg, work_dir, use_word, kinds=seg_kinds,
                tag=f"seg{idx}")
            log.info("segment[%d] text kinds=%s start=%s from_word=%s",
                     idx, seg_kinds, start_sheet, from_word)
            if text_pdf and os.path.exists(text_pdf) and not from_word:
                # Наложение рамки/штампа — критично: без него PDF уходит
                # без основных надписей. Ошибка пробрасывается наверх.
                # Подписи — по разделу сегмента: в текстовом сегменте обычно
                # один документный раздел (ПЗ), поэтому берём последний из них.
                seg_section = _segment_signature_kind(seg_kinds)
                text_pdf = _stamp_text_pages(
                    text_pdf, cfg, volume_id, total, page_forms,
                    cell9_texts=_cell9_texts(project), sheet_start=start_sheet,
                    signatures=_signature_texts(project, seg_section,
                                                FORM_TEXT_FIRST))
            segment_files.append(text_pdf if text_pdf and os.path.exists(text_pdf) else None)
            segment_meta.append(("text", seg_kinds, from_word))

        elif seg_kind == "graphic":
            graphic_pdfs = []
            g_index = 0
            sheet_counter = start_sheet
            # Документы графической части сегмента. Пустой раздел листа не
            # занимает: лист с перечнем без самих листов в том не вставляется
            # (иначе в PDF оставался бы одинокий лист перечня).
            graphic_docs = [d for d in project.documents_for(DocKind.GRAPHIC)
                            if d.path and os.path.exists(d.path)]
            # Лист «Графическая часть» с перечнем листов — ПЕРВЫМ листом
            # раздела (если включён в настройках проекта).
            from .preview_html import _graphic_sheet_list_cfg
            if graphic_docs and _graphic_sheet_list_cfg(project).enabled:
                gl_out = os.path.join(work_dir, f"_g{idx}_list.pdf")
                made = _graphic_sheet_list_pdf(project, cfg, gl_out,
                                               sheet_counter, total)
                if made:
                    graphic_pdfs.append(made)
                    sheet_counter += 1
                    g_index += 1
            for d in graphic_docs:
                src = d.path
                out = os.path.join(work_dir, f"_g{idx}_{g_index}.pdf")
                sheet_counter = _stamp_graphic_number_box(src, out, sheet_counter)
                graphic_pdfs.append(out)
                g_index += 1
            # Склеиваем документы графики сегмента в один файл (или None)
            if graphic_pdfs:
                gseg = os.path.join(work_dir, f"_seg_graphic_{idx}.pdf")
                gm = pymupdf.open()
                for gp in graphic_pdfs:
                    with pymupdf.open(gp) as g:
                        gm.insert_pdf(g)
                gm.save(gseg)
                gm.close()
                segment_files.append(gseg)
            else:
                segment_files.append(None)
            log.info("segment[%d] graphic docs=%d start=%s",
                     idx, g_index, start_sheet)
            segment_meta.append(("graphic", seg_kinds, g_index))

        elif seg_kind == "appendix":
            appendix_pdfs = [a.file_path for a in project.appendices_needed()
                             if a.file_path and os.path.exists(a.file_path)]
            if appendix_pdfs:
                aseg = os.path.join(work_dir, f"_seg_appendix_{idx}.pdf")
                am = pymupdf.open()
                for ap in appendix_pdfs:
                    with pymupdf.open(ap) as ap_:
                        am.insert_pdf(ap_)
                am.save(aseg)
                am.close()
                segment_files.append(aseg)
            else:
                segment_files.append(None)
            log.info("segment[%d] appendix files=%d", idx, len(appendix_pdfs))
            segment_meta.append(("appendix", seg_kinds, len(appendix_pdfs)))

    # ---- Склейка сегментов строго в порядке structure ----
    for sf in segment_files:
        if sf and os.path.exists(sf):
            with pymupdf.open(sf) as part:
                merged.insert_pdf(part)

    if merged.page_count == 0:
        pg = merged.new_page(width=A4_W_PT, height=A4_H_PT)
        pg.insert_text((72, 100), "Том пуст — нет включённых разделов",
                       fontname="helv", fontsize=11)

    n_graphic = sum(n for k, _kk, n in segment_meta if k == "graphic")
    n_appendix = sum(n for k, _kk, n in segment_meta if k == "appendix")
    has_text = any(k == "text" for k, _kk, _n in segment_meta)
    _add_bookmarks(merged, project, has_text, n_graphic, n_appendix)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    merged.save(out_path)
    merged.close()
    log.info("build_full_pdf: saved %s (%d pages)", out_path, total)
    return out_path


def build_section_pdf(project: VolumeProject, section_kind: str, out_path: str,
                      cfg: Config | None = None, use_word: bool = True,
                      work_dir: str | None = None) -> str:
    """Собирает выбранный раздел тома в единый PDF.

    :param section_kind: ключ раздела — 'content' | 'changes' |
        DocKind.POYASNITELNAYA.value | DocKind.GRAPHIC.value | 'appendix'.
    """
    import pymupdf
    cfg = cfg or Config()
    if work_dir is None:
        proj_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        work_dir = os.path.join(proj_root, "tmp", "_build")
    os.makedirs(work_dir, exist_ok=True)

    total = compute_total_sheets(project)
    volume_id = project.requisites.volume_id

    # Текстовые разделы (content / changes / cover / title / ПЗ) — DOCX -> PDF.
    text_pdf = None
    from_word = False
    page_forms = []
    if section_kind in ("content", "changes", "cover", "title",
                        DocKind.POYASNITELNAYA.value):
        text_pdf, from_word, page_forms = _section_text_pdf(project, section_kind, cfg, work_dir, use_word)

    # Графическая часть — исходные документы с квадратом номера листа.
    graphic_pdfs = []
    sheet_counter = 1
    if section_kind == DocKind.GRAPHIC.value:
        gi = 0
        for d in project.documents_for(DocKind.GRAPHIC):
            if d.path and os.path.exists(d.path):
                out = os.path.join(work_dir, f"_sec_g{gi}.pdf")
                sheet_counter = _stamp_graphic_number_box(d.path, out, sheet_counter)
                graphic_pdfs.append(out)
                gi += 1

    # Приложения — PDF как есть.
    appendix_pdfs = []
    if section_kind == "appendix":
        for a in project.appendices_needed():
            if a.file_path and os.path.exists(a.file_path):
                appendix_pdfs.append(a.file_path)

    # Наложить рамку/штамп на текстовые страницы (если собраны через PyMuPDF).
    # Ошибка пробрасывается: PDF без основных надписей — не результат сборки.
    if text_pdf and os.path.exists(text_pdf) and not from_word:
        text_pdf = _stamp_text_pages(text_pdf, cfg, volume_id, total, page_forms,
                                     cell9_texts=_cell9_texts(project))

    merged = pymupdf.open()
    if text_pdf and os.path.exists(text_pdf):
        with pymupdf.open(text_pdf) as t:
            merged.insert_pdf(t)
    for gp in graphic_pdfs:
        with pymupdf.open(gp) as g:
            merged.insert_pdf(g)
    for ap in appendix_pdfs:
        with pymupdf.open(ap) as ap_:
            merged.insert_pdf(ap_)

    # Если раздел пуст (нет ни одной страницы) — создаём пустую страницу-заглушку.
    if merged.page_count == 0:
        pg = merged.new_page(width=A4_W_PT, height=A4_H_PT)
        pg.insert_text((72, 100), f"Раздел: {section_kind} (нет содержимого)",
                       fontname="helv", fontsize=11)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    merged.save(out_path)
    merged.close()
    return out_path


def _section_text_pdf(project, section_kind, cfg, work_dir, use_word):
    """Собирает DOCX раздела и конвертирует в PDF (через Word или PyMuPDF)."""
    from .docx_exporter import build_section_docx
    docx_path = os.path.join(work_dir, f"_sec_{section_kind}.docx")
    build_section_docx(project, section_kind, docx_path, cfg)

    if use_word:
        try:
            from .word_com import word_available, docx_to_pdf
            if word_available():
                pdf_path = os.path.join(work_dir, f"_sec_{section_kind}.pdf")
                docx_to_pdf(docx_path, pdf_path)
                return pdf_path, True, []
        except Exception:
            pass

    # Fallback: собрать через PyMuPDF с кириллицей.
    pdf_path, page_forms = _section_text_pdf_pymupdf(project, section_kind, docx_path, cfg, work_dir)
    return pdf_path, False, page_forms


def _section_text_pdf_pymupdf(project, section_kind, docx_path, cfg, work_dir):
    """Запасной путь: текстовая часть раздела напрямую через PyMuPDF.

    Извлекает текст из DOCX раздела и размещает его на страницах A4 с рамкой
    и осн. надписью. Возвращает (путь_к_pdf, page_forms)."""
    import pymupdf
    from .docx_reader import read_docx
    pdf_path = os.path.join(work_dir, f"_sec_{section_kind}.pdf")
    doc = pymupdf.open()
    fname = _register_text_font(doc)
    req = project.requisites
    page_forms = []
    # Шрифты групп текста ПЗ из настроек — регистрируются на каждой странице.
    group_files = _text_font_files(cfg)
    page_fonts = {"body": fname, "heading": fname, "table": fname}

    def new_page(form, width=A4_W_PT, height=A4_H_PT):
        nonlocal page_forms
        pg = doc.new_page(width=width, height=height)
        page_forms.append(form)
        # Шрифт встраивается в ресурсы страницы: без этого кириллица
        # выводилась бы точками (PyMuPDF искал бы его в стандартных 14).
        _ensure_page_font(pg, fname)
        if group_files:
            page_fonts.update(_register_group_fonts(pg, group_files))
        return pg

    if section_kind == "cover":
        pg = new_page(None)
        pg.insert_text((72, 100), req.organization or "", fontname=fname, fontsize=18)
        pg.insert_text((72, 140), "ОБЛОЖКА", fontname=fname, fontsize=14)
        pg.insert_text((72, 180), req.volume_id or "", fontname=fname, fontsize=14)
    elif section_kind == "title":
        pg = new_page(None)
        pg.insert_text((72, 100), req.volume_name or "", fontname=fname, fontsize=16)
        pg.insert_text((72, 140), "ТИТУЛЬНЫЙ ЛИСТ", fontname=fname, fontsize=14)
        pg.insert_text((72, 180), req.object_name or "", fontname=fname, fontsize=12)
    elif section_kind == "content":
        pg = new_page(FORM_TEXT_FIRST)
        pg.insert_text((72, 90), "СОДЕРЖАНИЕ", fontname=fname, fontsize=14)
        y = 120
        # Автоматический сбор содержания по фактической структуре тома.
        for name, des, sheets in toc_content_rows(project):
            line = name
            if des:
                line += f"  {des}"
            if sheets:
                line += f"  {sheets}"
            pg.insert_text((72, y), line, fontname=fname, fontsize=11)
            y += 20
    elif section_kind == "changes":
        pg = new_page(FORM_TEXT_FIRST)
        pg.insert_text((72, 90), "ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ", fontname=fname, fontsize=13)
        y = 130
        for c in project.changes:
            line = f"{c.number}  {c.date}  {c.description}"
            pg.insert_text((72, y), line, fontname=fname, fontsize=10)
            y += 18
    else:
        # ПЗ: оглавление (по уровням текста DOCX), затем документы раздела.
        # Номера в оглавлении — как внизу листа (графа «Лист» штампа):
        # нумерация внутри раздела, сдвиг — число листов самого оглавления.
        from .pz_toc import toc_page_offset
        toc_pages = _add_toc_pages_pymupdf(
            project, new_page, fname, cfg.font_map,
            (FORM_TEXT_FIRST, FORM_TEXT_NEXT),
            offset=toc_page_offset(project))
        first = not toc_pages
        for d in project.documents_for(DocKind.POYASNITELNAYA):
            if not (d.path and os.path.exists(d.path)):
                continue
            try:
                content = read_docx(d.path)
            except Exception:
                continue
            pg = new_page(FORM_TEXT_FIRST if first else FORM_TEXT_NEXT)
            first = False
            y = 80
            # Шрифт и кегль текста ПЗ — из настроек (config.TextFonts), как в
            # предпросмотре и в Word.
            tf = getattr(cfg, "text_fonts", None)
            body_size = tf.size("body") if tf else 11
            head_size = tf.size("heading") if tf else 11
            for para in content.paragraphs:
                if not para.text.strip():
                    continue
                group = "heading" if para.is_heading else "body"
                size = head_size if para.is_heading else body_size
                for ln in _wrap(para.text, width=90):
                    if y > 800:
                        pg = new_page(FORM_TEXT_NEXT)
                        y = 80
                    pg.insert_text((72, y), ln,
                                   fontname=page_fonts.get(group, fname),
                                   fontsize=size)
                    y += size * 1.45
                    y += 16

    doc.save(pdf_path)
    doc.close()
    return pdf_path, page_forms


def _add_bookmarks(doc, project, has_text: bool, n_graphic: int, n_appendix: int) -> None:
    """Добавляет закладки (оглавление) по разделам тома."""
    toc = []
    if has_text:
        toc.append([1, "Титульный лист", 1])
        toc.append([1, "Содержание", 1])
        toc.append([1, "Пояснительная записка", 1])
    if n_graphic:
        toc.append([1, "Графическая часть", 1])
    if n_appendix:
        toc.append([1, "Приложения", 1])
    if project.changes:
        toc.append([1, "Лист регистрации изменений", doc.page_count])
    if toc:
        try:
            doc.set_toc(toc)
        except Exception:
            pass
