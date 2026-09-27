# -*- coding: utf-8 -*-
"""
Сборка полного тома: текстовая часть (титул+содержание+ПЗ+лист изменений) +
графические листы + приложения в единый PDF, с рамками/осн. надписью по ГОСТ,
сквозной нумерацией листов и закладками.
"""
from __future__ import annotations

import os
import tempfile

from ..models import VolumeProject, DocKind
from ..config import Config
from .gost import FORM_GRAPHIC, FORM_TEXT_FIRST, FORM_TEXT_NEXT
from .pdf_merge import count_pages, compute_total_sheets
from .renderer import draw_frame_and_stamp, draw_sheet_number_box
from .toc import content_rows as toc_content_rows


def _text_pdf(project: VolumeProject, cfg: Config, out_dir: str, use_word: bool = True) -> tuple[str | None, bool, list]:
    """Формирует PDF текстовой части.

    Если use_word и доступен Word — конвертируем DOCX через Word COM.
    DOCX уже содержит рамку и основную надпись (форма 5/6), поэтому
    дополнительно штамповать НЕ нужно.

    Иначе текстовая часть собирается напрямую через PyMuPDF (fallback) и
    требует последующего наложения штампов.

    Возвращает (путь_к_pdf | None, пришёл_из_word, page_forms).
    page_forms — список форм основной надписи по каждой странице текстовой части:
        None — без рамки и штампа (обложка);
        2    — большой штамп (первый/заглавный лист раздела);
        '2a' — малый штамп (последующий лист).
    """
    # Собираем DOCX текстовой части (всегда доступно)
    from .docx_exporter import build_text_docx
    docx_path = os.path.join(out_dir, "_text_part.docx")
    build_text_docx(project, docx_path, cfg)

    # 1) Если разрешён Word и он доступен — конвертируем DOCX -> PDF (максимальная точность)
    if use_word:
        try:
            from .word_com import word_available, docx_to_pdf
            if word_available():
                pdf_path = os.path.join(out_dir, "_text_part.pdf")
                docx_to_pdf(docx_path, pdf_path)
                return pdf_path, True, []
        except Exception:
            pass

    # 2) Fallback: собрать текстовые страницы напрямую через PyMuPDF (без Word)
    pdf_path, page_forms = _text_pdf_pymupdf(project, out_dir, cfg)
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
                         toc_pdf_lines)
    items, _path = project_toc_items(project)
    if not items:
        return 0
    pages = paginate_items(items)
    if offset is None:
        offset = toc_page_offset(project)
    first_form, next_form = forms
    for i, page_items in enumerate(pages):
        pg = new_page(first_form if i == 0 else next_form)
        y = 90
        for j, line in enumerate(toc_pdf_lines(page_items, offset=offset)):
            size = fm.content_header if j == 0 else fm.toc_row
            pg.insert_text((72, y), line, fontname=fname, fontsize=size)
            y += 20
    return len(pages)


def _register_text_font(doc):
    """Регистрирует шрифт ISOCPEUR (кириллица) в PDF-документе.

    Возвращает имя шрифта для insert_text; при недоступности — 'helv'.
    """
    from .fonts import font_file
    fpath = font_file()
    if not fpath:
        return "helv"
    try:
        with open(fpath, "rb") as f:
            buf = f.read()
        # Регистрируем шрифт один раз для документа (вне страницы).
        import pymupdf
        doc.insert_font(fontname="ISOCPEUR", fontbuffer=buf)
        return "ISOCPEUR"
    except Exception:
        return "helv"


def _text_pdf_pymupdf(project, out_dir, cfg):
    """Собирает текстовую часть (обложка/титул/содержание/ПЗ) как PDF-страницы
    напрямую через PyMuPDF — запасной путь, когда Word COM недоступен.

    Возвращает (путь_к_pdf, page_forms). page_forms[i] — форма осн. надписи для
    страницы i: None (обложка — без рамки/штампа), FORM_TEXT_FIRST (первый лист
    раздела — большой штамп), FORM_TEXT_NEXT (последующий лист — малый штамп)."""
    import pymupdf
    pdf_path = os.path.join(out_dir, "_text_part.pdf")
    doc = pymupdf.open()
    fm = cfg.font_map
    req = project.requisites
    page_forms = []         # форма по каждой странице текстовой части

    # Шрифт с поддержкой кириллицы (ISOCPEUR, если доступен).
    fname = _register_text_font(doc)

    # Хелпер: добавить страницу и зафиксировать её форму основной надписи.
    def new_page(form, width=595, height=842):
        nonlocal page_forms
        pg = doc.new_page(width=width, height=height)
        page_forms.append(form)
        return pg

    # ---- Обложка (без рамки и без штампа) ----
    pg = new_page(None)
    pg.insert_text((72, 100), req.organization or "", fontname=fname, fontsize=18)
    pg.insert_text((72, 140), "ОБЛОЖКА", fontname=fname, fontsize=14)
    pg.insert_text((72, 180), req.volume_id or "", fontname=fname, fontsize=14)

    # ---- Титульный лист (как и обложка — БЕЗ рамки и БЕЗ штампа) ----
    # По ГОСТ и в предпросмотре титульный лист не штампуется и не нумеруется;
    # первый нумерованный лист тома — содержание. Поэтому form=None.
    pg = new_page(None)
    pg.insert_text((72, 100), req.volume_name or "", fontname=fname, fontsize=16)
    pg.insert_text((72, 140), "ТИТУЛЬНЫЙ ЛИСТ", fontname=fname, fontsize=14)
    pg.insert_text((72, 180), req.object_name or "", fontname=fname, fontsize=12)

    # ---- Содержание (раздел — первый лист форма 5, №1 сквозной нумерации) ----
    # Автоматический сбор по фактической структуре тома (единый источник toc).
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

    # ---- Пояснительная записка (первый лист — форма 5, далее форма 6) ----
    from .docx_reader import read_docx
    # Оглавление записки — первыми листами раздела ПЗ (первый лист — форма 5).
    from .pz_toc import toc_page_offset
    toc_pages = _add_toc_pages_pymupdf(project, new_page, fname, fm,
                                       (FORM_TEXT_FIRST, FORM_TEXT_NEXT),
                                       offset=toc_page_offset(project))
    # Без оглавления большой штамп (форма 5) достаётся первому листу записки.
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
            for para in content.paragraphs:
                if not para.text.strip():
                    continue
                lines = _wrap(para.text, width=90)
                for ln in lines:
                    if y > 800:
                        # следующая страница раздела — малый штамп (форма 6)
                        pg = new_page(FORM_TEXT_NEXT)
                        y = 80
                    pg.insert_text((72, y), ln, fontname=fname, fontsize=11)
                    y += 16

    doc.save(pdf_path)
    doc.close()
    return pdf_path, page_forms


def _wrap(text: str, width: int = 90) -> list:
    """Простой перенос строк по ширине (для fallback-рендера)."""
    import textwrap
    return textwrap.wrap(text, width) or [text]


def _copy_graphic(src_pdf: str, out_pdf: str) -> None:
    """Копирует графический PDF без оформления (совместимость)."""
    import pymupdf
    doc = pymupdf.open(src_pdf)
    doc.save(out_pdf)
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


def _stamp_text_pages(text_pdf: str, cfg: Config, volume_id: str, total: int,
                      page_forms: list | None = None,
                      cell9_texts: dict | None = None) -> str:
    """Накладывает рамку и основную надпись на текстовые страницы.

    page_forms[i] определяет форму основной надписи для страницы i:
        None            — не рисовать рамку и штамп (обложка);
        FORM_TEXT_FIRST — большой штамп (первый/заглавный лист раздела);
        FORM_TEXT_NEXT  — малый штамп (последующий лист).
    Если page_forms не задан — первая страница FORM_TEXT_FIRST, остальные
    FORM_TEXT_NEXT.

    Возвращает путь к новому PDF (или исходный при ошибке)."""
    import pymupdf
    doc = pymupdf.open(text_pdf)
    try:
        sheet_no = 1
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
                     "qty": str(total)}
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
                   total: int, volume_id: str, cell9_texts: dict | None = None) -> int:
    """Накладывает рамку и основную надпись (форма 3) на листы графической части.

    Нумерация листов — сквозная (начиная с sheet_start). Возвращает следующий
    свободный номер листа.

    Размер страницы берётся из каждого листа PDF (если он ненулевой), иначе
    используются переданные page_w_mm/page_h_mm (на случай нестандартной графики).
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
                 "qty": str(total)}
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
    """Собирает полный том в единый PDF."""
    import pymupdf
    cfg = cfg or Config()
    # По умолчанию рабочий каталог — внутри проекта (tmp), т.к. системный temp
    # в некоторых средах недоступен для записи zip/DOCX.
    if work_dir is None:
        proj_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        work_dir = os.path.join(proj_root, "tmp", "_build")
    os.makedirs(work_dir, exist_ok=True)

    total = compute_total_sheets(project)  # «Листов M»
    volume_id = project.requisites.volume_id

    # 1) текстовая часть (DOCX -> PDF через Word, если возможно)
    text_pdf, from_word, page_forms = _text_pdf(project, cfg, work_dir, use_word)

    # 2) графические листы. Исходная графика сохраняется без изменений, поверх
    # неё добавляется только квадрат со сквозным номером листа.
    graphic_pdfs = []
    g_index = 0
    # Начальный номер листа графики = число НУМЕРОВАННЫХ страниц текстовой части + 1.
    # Обложка и титульный лист не нумеруются (согласовано с содержанием toc), поэтому
    # нельзя брать просто физический page_count: он включает обложку/титул.
    sheet_counter = 1
    if text_pdf and os.path.exists(text_pdf):
        import pymupdf
        with pymupdf.open(text_pdf) as t:
            if page_forms is not None:
                # fallback-путь: нумеруются только страницы с формой (не None).
                sheet_counter = sum(1 for f in page_forms if f is not None) + 1
            else:
                # Word-путь: обложка/титул без номера, штампы внутри DOCX.
                # Используем единый источник toc — последний номер листа перед
                # графической частью.
                from .toc import build_layout as _build_layout
                lay = _build_layout(project)
                pre_graphic = [it for it in lay.items
                               if it.kind not in ("graphic", "appendix", "changes")]
                sheet_counter = max((it.last_sheet for it in pre_graphic), default=0) + 1
    graphic_docs = [d for d in project.documents_for(DocKind.GRAPHIC)]
    for d in graphic_docs:
        src = d.path
        if not src or not os.path.exists(src):
            continue
        out = os.path.join(work_dir, f"_g{g_index}.pdf")
        sheet_counter = _stamp_graphic_number_box(src, out, sheet_counter)
        graphic_pdfs.append(out)
        g_index += 1

    # 3) приложения (PDF как есть, без рамки — самостоятельный документ)
    appendix_pdfs = [a.file_path for a in project.appendices_needed()
                     if a.file_path and os.path.exists(a.file_path)]

    # 3.5) наложить рамку + осн. надпись (ф.5 / ф.6) на текстовые страницы.
    # Только для PDF, собранного напрямую через PyMuPDF. Если текстовая часть
    # пришла из Word (DOCX уже содержит рамку и осн. надпись) — не накладываем,
    # чтобы не было двойного штампа.
    if text_pdf and os.path.exists(text_pdf) and not from_word:
        try:
            text_pdf = _stamp_text_pages(text_pdf, cfg, volume_id, total, page_forms,
                                         cell9_texts=_cell9_texts(project))
        except Exception:
            pass

    # 4) слить всё в единый PDF в порядке: текст, графика, приложения
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

    # 5) закладки по разделам
    _add_bookmarks(merged, project, text_pdf is not None, len(graphic_pdfs), len(appendix_pdfs))

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    merged.save(out_path)
    merged.close()
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
    if text_pdf and os.path.exists(text_pdf) and not from_word:
        try:
            text_pdf = _stamp_text_pages(text_pdf, cfg, volume_id, total, page_forms,
                                         cell9_texts=_cell9_texts(project))
        except Exception:
            pass

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
        pg = merged.new_page(width=595, height=842)
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

    def new_page(form, width=595, height=842):
        nonlocal page_forms
        pg = doc.new_page(width=width, height=height)
        page_forms.append(form)
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
        # В отдельном разделе нумерация начинается заново, поэтому сдвиг номеров
        # в оглавлении — число листов самого оглавления.
        from .pz_toc import toc_page_offset
        toc_pages = _add_toc_pages_pymupdf(
            project, new_page, fname, cfg.font_map,
            (FORM_TEXT_FIRST, FORM_TEXT_NEXT),
            offset=toc_page_offset(project, standalone=True))
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
            for para in content.paragraphs:
                if not para.text.strip():
                    continue
                for ln in _wrap(para.text, width=90):
                    if y > 800:
                        pg = new_page(FORM_TEXT_NEXT)
                        y = 80
                    pg.insert_text((72, y), ln, fontname=fname, fontsize=11)
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
