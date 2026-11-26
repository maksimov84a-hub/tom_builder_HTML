# -*- coding: utf-8 -*-
"""
Экспорт текстовой части тома в DOCX (обложка + титульный лист + содержание +
пояснительная записка) с оформлением по ГОСТ Р 21.101-2020.

Основная надпись (форма 3 / 6) выводится как таблица в нижнем колонтитуле
(разные для первой и последующих страниц). Шрифт — ISOCPEUR, кегль — по карте.

Рендер DOCX не зависит от Word — только python-docx. Конвертация в PDF
(через Word COM) — в word_com.py.
"""
from __future__ import annotations

import os
import re
from typing import Optional

from ..models import (
    VolumeProject, DocKind, Stage, SECTION_LABELS,
    SECTION_COVER, SECTION_TITLE, SECTION_CONTENT, SECTION_PZ,
    SECTION_CHANGES, SECTION_COMPOSITION, SECTION_GRAPHIC, SECTION_APPENDIX,
)
from ..config import Config, FontMap, PageMargins, A4_MM, FONT_NAME
from .docx_reader import read_docx
from .toc import content_rows
# Автоматическое оглавление ПЗ по уровням текста из Word (единый источник).
from .pz_toc import (project_toc_items, paginate_items, toc_page_offset,
                     toc_text_style, anchor_map, is_heading_paragraph,
                     heading_emphasis, level_indent_mm, DOCX_TAB_POS_MM)
from .gost import (LEFT_MM, RIGHT_MM, TOP_MM, BOTTOM_MM,
                   FORM_GRAPHIC, FORM_TEXT_FIRST, FORM_TEXT_FIRST_WORD,
                   FORM_TEXT_NEXT,
                   STAMP_WIDTH_MM)
# Кегль граф «Стадия / Лист / Листов» (единый источник — geometry.py).
from .geometry import STAGE_CELL_KEYS, STAGE_CELL_PT, CELL9_KEYS, CELL9_PT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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


def _set_run(run, font: str, size_pt: float, bold=False, italic=False):
    """Задаёт шрифт, кегль и начертание run-а.

    :param italic: курсив; по умолчанию снимается (обычный текст записки
        выводится без курсива, заголовок передаёт курсив исходника).
    """
    from docx.shared import Pt
    run.font.name = font
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
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
    SECTION_COVER, SECTION_TITLE, SECTION_COMPOSITION, SECTION_CONTENT,
    SECTION_PZ, SECTION_CHANGES,
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
    # Пустые разделы листа не занимают — так же, как в содержании и PDF:
    # changes без журнала, ПЗ без включённых документов. Иначе в томе
    # оставался бы пустой лист с одним штампом.
    text_kinds = [
        k for k in text_kinds
        if (k != SECTION_CHANGES or project.changes)
        and (k != SECTION_PZ or project.documents_for(DocKind.POYASNITELNAYA))
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
        elif kind == SECTION_COMPOSITION:
            _add_composition(doc, project, fm, font)
            _add_stamp_footers(doc, project, config, sec,
                               first_form=FORM_TEXT_FIRST_WORD,
                               section_kind=kind)
        elif kind == SECTION_CONTENT:
            _add_content(doc, project, fm, font)
            _add_stamp_footers(doc, project, config, sec,
                               first_form=FORM_TEXT_FIRST_WORD,
                               section_kind=kind)
        elif kind == SECTION_PZ:
            # Оглавление и закладки-цели строятся по одному документу (первому
            # в разделе ПЗ) — ссылки из оглавления ведут в него.
            _toc_items, toc_path = project_toc_items(project)
            toc_anchors = anchor_map(_toc_items)
            toc_pages = _add_pz_toc(doc, project, fm, font,
                                     text_fonts=getattr(config, "text_fonts", None))
            # Если оглавление заняло первый лист раздела, большая надпись уже
            # использована — листы самой записки начинаются с формы 6, как в
            # предпросмотре (иначе первый лист ПЗ получал бы большой штамп).
            pz_first = FORM_TEXT_NEXT if toc_pages else FORM_TEXT_FIRST_WORD
            for index, src in enumerate(
                    project.documents_for(DocKind.POYASNITELNAYA)):
                _add_pz_doc(
                    doc, src.path, fm, font,
                    text_fonts=getattr(config, "text_fonts", None),
                    first_form=pz_first,
                    break_before_first=bool(toc_pages or index),
                    toc_source=toc_path, toc_anchors=toc_anchors)
                pz_first = FORM_TEXT_NEXT
            _add_stamp_footers(doc, project, config, sec,
                               first_form=FORM_TEXT_FIRST_WORD,
                               section_kind=kind)
        elif kind == SECTION_CHANGES:
            _add_changes_sheet(doc, project, fm, font)
            _add_stamp_footers(doc, project, config, sec,
                               first_form=FORM_TEXT_FIRST_WORD,
                               section_kind=kind)

    # Пустой текстовый том — минимальный лист, чтобы файл был валиден
    if not text_kinds:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run("Нет текстовых разделов в структуре тома"),
                 font, fm.content_header)

    return _save_docx(doc, out_path)


#: Внутренний отступ ячейки таблицы по умолчанию (мм) — столько Word
#: оставляет слева и справа от текста, если не задано иное.
_CELL_TEXT_PAD_MM = 1.9

#: Коэффициент высоты строки абзаца ПЗ в Word: высота = K × кегль (мм).
#: Оценка пагинации считает 0.5 × кегль, но HTML рисует строку 11 pt как
#: 5.56 мм (замер том.pdf), то есть 0.505. Word с точным межстрочным 5.5 мм
#: давал 5.1–5.2 мм по метрикам шрифта — на 8 % меньше, и на лист влезало
#: больше строк, чем в предпросмотре. 0.5 совпадает с HTML.
_LINE_MM_K = 0.5


#: Ключи форм основной надписи в терминах HTML-предпросмотра
#: (`preview_html.FORM_HEIGHT`): 1 — форма 3, 2 — форма 5, "2a" — форма 6.
#: Пагинация листов (preview_html._paginate_docx) принимает именно их.
#: Здесь важно только число миллиметров под штамп: у форм 3 и 5 оно одно (40),
#: поэтому первый лист текста берёт ключ 2 независимо от номера формы ГОСТ.
_PZ_FIRST_FORM = 2
_PZ_NEXT_FORM = "2a"


#: Внутренний отступ текста от линий рамки (мм) — как `padding` у
#: `.sheet-content` в HTML-предпросмотре. Текст начинается на 20+6 = 26 мм
#: слева и на 17+6 = 23 мм сверху (17 мм — зона квадрата сквозной нумерации).
SHEET_PAD_MM = 6.0

#: Отступ сверху до начала текста в HTML: TEXT_TOP_INDENT_MM (17) + padding (6).
SHEET_TEXT_TOP_MM = 23.0

#: Просвет между низом текста и верхом штампа в HTML (`+2.0` в
#: preview_html._page_text_height).
SHEET_STAMP_GAP_MM = 2.0


def _apply_sheet_margins(sec, config, stamp_h_mm: float = 0.0) -> None:
    """Поля листа, колонтитулы и текстовая область как в HTML-предпросмотре.

    Рамка в Word привязана к границе текста, а граница текста — к ПОЛЯМ, поэтому
    поля остаются равными линиям рамки (20 мм слева, по 5 мм с других сторон) и
    линии проходят ровно по ним: 19.7…205.1 × 4.7…292.1 мм (замер в Word).

    Внутренний отступ текста 6 мм задаётся не полями (они сдвинули бы рамку), а:
      * сверху — пустым абзацем в верхнем колонтитуле (`_set_top_indent`);
      * слева/справа — отступами абзацев и таблиц (`SHEET_PAD_MM`).

    Нижняя граница НЕ добавляет отступ: в HTML текст заканчивается прямо у
    штампа (`bottom = PAGE_BOTTOM_MM + штамп + 2 мм`), и Word должен
    остановиться там же. Раньше сюда добавлялись ещё 6 мм отступа, из-за чего
    текст обрывался на 8 мм выше и листы не совпадали с предпросмотром.

    :param stamp_h_mm: высота штампа раздела (мм) — от неё зависит нижняя
        граница текста. 0 — раздела без штампа (нижнее поле по настройкам).
    """
    from docx.shared import Mm
    m = config.margins
    sec.left_margin = Mm(m.left)
    sec.right_margin = Mm(m.right)
    sec.top_margin = Mm(m.top)
    sec.header_distance = Mm(m.top)
    sec.footer_distance = Mm(m.bottom)
    if stamp_h_mm > 0:
        # Резерв снизу выровнен с бюджетом пагинации (`_page_text_height`):
        #   297 −(5 + штамп + 2) −17 −12 = 246 мм  (форма 6)
        # Текст начинается на 23 мм (17 зоны квадрата + 6 отступа), поэтому
        # нижнее поле = 297 − 23 − 246 = 28 мм, то есть
        # 5 (рамка) + 15 (штамп) + 2 (просвет) + 6 (запас бюджета).
        # Лишние 6 мм отступа снизу НЕ добавляются: в HTML текст снизу доходит
        # до штампа, и Word не должен обрываться раньше — иначе на лист
        # влезает на одну строку меньше и разрывы расходятся.
        # Резерв по САМОМУ МАЛОМУ штампу раздела (форма 6): поле задаётся на
        # секцию, а штампы разные — при резерве 40 мм листы с формой 6 рвались
        # много раньше предпросмотра.
        from .geometry import FORM_HEIGHT
        reserve = min(stamp_h_mm, FORM_HEIGHT[6])
        # 297 − 23 − бюджет, где бюджет = 297 −(5+reserve+2) −17 −12.
        # Это тождественно: 5 + reserve + 2 + 6.
        sec.bottom_margin = Mm(m.bottom + reserve + SHEET_STAMP_GAP_MM + 6.0)
    else:
        sec.bottom_margin = Mm(m.bottom)


def _set_top_indent(header, top_mm: float = SHEET_TEXT_TOP_MM) -> None:
    """Сдвигает текст вниз пустым абзацем верхнего колонтитула.

    Поле остаётся 5 мм (рамка проходит по нему), а текст начинается на
    `top_mm` от верха листа — как в HTML, где над текстом зона квадрата
    сквозной нумерации (17 мм) и внутренний отступ 6 мм.

    Высота абзаца = `top_mm` минус расстояние колонтитула: колонтитул стоит на
    `header_distance` от края листа, поэтому его абзац заданной высоты H
    заканчивается на distance + H. При 5 + 18 = 23 мм текст встаёт ровно на
    23 мм (замер в Word: 22.7 мм).
    """
    from docx.shared import Pt
    from docx.enum.text import WD_LINE_SPACING
    header.is_linked_to_previous = False
    for extra in list(header.paragraphs)[1:]:
        extra._element.getparent().remove(extra._element)
    para = header.paragraphs[0]
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.left_indent = Pt(0)
    pf.right_indent = Pt(0)
    pf.line_spacing = Pt(max(0.0, top_mm - 5.0) * 72.0 / 25.4)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY


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
    # Стадия — полным наименованием, без надписи-префикса (как в предпросмотре);
    # в основной надписи остаётся обозначение «П» / «Р».
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    from ..models import stage_label
    stage_text = stage_label(req.stage)
    if stage_text:
        _set_run(p.add_run(stage_text), font, fm.stamp_columns)
    loc = " ".join(x for x in [req.city, str(req.year)] if x)
    if loc:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run(loc), font, fm.stamp_columns)


def _disable_external_update_prompt(path: str) -> None:
    """Просит Word пересчитать поля при открытии документа.

    В оглавлении номера листов — поля `PAGEREF` на закладки заголовков: их
    значения зависят от фактической раскладки, и Word должен обновить их сам.
    Флаг `w:updateFields` заставляет его сделать это при открытии, поэтому в
    предпросмотре и в PDF сразу видны верные номера.

    Внешних ссылок в собираемых документах нет (все ссылки — внутренние
    закладки), поэтому запроса «обновить внешние ссылки?» не возникает.
    """
    import zipfile
    from lxml import etree

    settings_path = "word/settings.xml"
    with zipfile.ZipFile(path, "r") as source:
        files = {name: source.read(name) for name in source.namelist()}
    xml = etree.fromstring(files[settings_path])
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    update_fields = xml.find("w:updateFields", ns)
    if update_fields is None:
        update_fields = etree.SubElement(
            xml, "{%s}updateFields" % ns["w"])
    update_fields.set("{%s}val" % ns["w"], "true")
    files[settings_path] = etree.tostring(
        xml, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in files.items():
            target.writestr(name, data)


def _save_docx(doc, out_path: str) -> str:
    """Сохраняет документ, не теряя работу при занятом файле.

    Файл пишется во временный «.saving» и затем подменяет целевой: так на месте
    не остаётся полузаписанного DOCX, если запись прервётся.

    Если целевой файл открыт в Word (Windows блокирует запись — `PermissionError:
    [Errno 13]`), результат сохраняется рядом под свободным именем
    («том (2).docx»), а сборка не падает: иначе пользователь терял всю сборку
    из-за открытого документа.

    :return: фактический путь сохранённого файла.
    """
    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    tmp = out_path + ".saving"
    doc.save(tmp)
    _disable_external_update_prompt(tmp)
    try:
        os.replace(tmp, out_path)
        return out_path
    except PermissionError:
        pass
    stem, ext = os.path.splitext(out_path)
    for n in range(2, 100):
        alt = f"{stem} ({n}){ext}"
        try:
            os.replace(tmp, alt)
            return alt
        except PermissionError:
            continue
    try:
        os.remove(tmp)
    except OSError:
        pass
    raise PermissionError(
        f"Не удалось записать файл «{os.path.basename(out_path)}»: он открыт в "
        f"другой программе (вероятно, в Word). Закройте файл и повторите сборку.")


def _fit_cell_run(para, text, font, prefer_pt: float, cell_w_mm: float,
                  fpath: str = "", bold: bool = False) -> None:
    """Добавляет текст в ячейку, подбирая кегль под её ширину.

    Внутри ячейки Word по умолчанию остаётся ~1.9 мм с каждой стороны, поэтому
    доступная ширина меньше ширины графы. Кегль уменьшается только при
    необходимости: длинное наименование остаётся ОДНОЙ строкой, иначе строка
    таблицы вырастает и разбиение на листы разъезжается.
    """
    from .textfit import fit_size
    pad = _CELL_TEXT_PAD_MM * 2
    size = fit_size(text, max(1.0, cell_w_mm - pad), prefer_pt, path=fpath)
    _set_run(para.add_run(text), font, size, bold=bold)


def _wrap_size(text: str, avail_mm: float, prefer_pt: float,
               fpath: str = "") -> float:
    """Кегль для текста, который РАЗРЕШЕНО переносить по словам.

    Замер ведётся по самому широкому слову, а не по всей строке: надпись
    «Номер тома» в графе 15 мм целиком шире ячейки (20.5 мм при 12 pt), и подбор
    «в одну строку» сжимал её до 6.5 pt — вдвое мельче остальных заголовков
    таблицы. С переносом достаточно, чтобы влезало слово: «Номер» 10.1 мм.
    """
    from .textfit import fit_size, text_width_mm
    words = text.split()
    if not words:
        return float(prefer_pt)
    widest = max(words, key=lambda w: text_width_mm(w, prefer_pt, fpath))
    return fit_size(widest, avail_mm, prefer_pt, path=fpath)


def _wrap_lines(text: str, size_pt: float, avail_mm: float,
                fpath: str = "") -> list:
    """Разбивает текст на строки по ширине графы — по словам, как Word.

    Перенос ставится явным разрывом строки внутри абзаца: высота строки таблицы
    задана правилом `exact`, поэтому лишний абзац в ячейке её не увеличивает, но
    и не «съедает» межстрочный интервал.
    """
    from .textfit import text_width_mm
    lines, current = [], ""
    for word in text.split():
        candidate = f"{current} {word}".strip()
        if current and text_width_mm(candidate, size_pt, fpath) > avail_mm:
            lines.append(current)
            current = word
        else:
            current = candidate
    if current:
        lines.append(current)
    return lines or [text]


def _add_wrapped_run(para, lines, font, size_pt: float, bold: bool = False):
    """Кладёт строки в абзац одной ячейки через разрыв строки Word."""
    run = None
    for index, line in enumerate(lines):
        if index:
            run.add_break()
        run = para.add_run(line)
        _set_run(run, font, size_pt, bold=bold)
    return run


def _add_bookmark(paragraph, name: str, bookmark_id: int) -> None:
    """Ставит в абзац закладку Word (цель гиперссылки из оглавления).

    `w:pPr` обязан оставаться первым дочерним элементом `w:p`, поэтому маркер
    начала закладки вставляется СРАЗУ за ним, а не в начало абзаца.
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(int(bookmark_id)))
    start.set(qn("w:name"), name)
    ppr = paragraph._p.pPr
    if ppr is not None:
        ppr.addnext(start)
    else:
        paragraph._p.insert(0, start)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(int(bookmark_id)))
    paragraph._p.append(end)


def _add_internal_link(paragraph, text: str, anchor: str, font: str,
                       size_pt: float):
    """Добавляет в абзац гиперссылку на закладку (внутри самого документа).

    Оформление — как у обычного текста строки оглавления (без синего цвета и
    подчёркивания: это печатный документ, ссылка нужна для навигации).
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    run = paragraph.add_run(text)
    _set_run(run, font, size_pt)
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), anchor)
    link.set(qn("w:history"), "1")
    run._r.getparent().remove(run._r)
    link.append(run._r)
    paragraph._p.append(link)


def _add_pz_toc(doc, project, fm, font, offset: int | None = None,
                text_fonts=None) -> int:
    """Добавляет автоматическое оглавление записки первым листом раздела ПЗ.

    Заголовки берутся из исходного DOCX по уровням текста Word
    (app.core.pz_toc): строки выводятся с абзацным отступом по уровню, номером
    листа записки и отточием. Отточие — правая табуляция с заполнителем DOTS
    (как «Содержание» в Word), поэтому номер листа стоит у правого края
    текстовой области, а промежуток заполнен точками.

    Оформление строк — как у обычного текста записки: шрифт и кегль группы
    «body» (config.TextFonts), без полужирного. Наименование заголовка —
    гиперссылка на закладку заголовка в самой записке.

    :param offset: сдвиг номеров листов (номер листа записки в томе минус 1) —
        чтобы напечатанный номер совпал с квадратом сквозной нумерации листа.
    :return: число добавленных листов (0 — оглавление не выводится).
    """
    from docx.shared import Mm, Pt
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER

    items, _path = project_toc_items(project)
    if not items:
        return 0
    # Шрифт и кегль строк — как у обычного текста записки.
    size = fm.toc_row
    toc_font = font
    if text_fonts is not None:
        size = text_fonts.size("body")
        toc_font = text_fonts.font("body") or font
    else:
        _style_font, _style_size = toc_text_style()
        size = _style_size or size
        toc_font = _style_font or font
    title_size = fm.content_header
    # Разбивка по листам — той же моделью высоты строки, что и вывод.
    pages = paginate_items(items, size)
    # Номера листов в оглавлении — как внизу листа (графа «Лист» штампа):
    # нумерация внутри раздела ПЗ, сдвиг = число листов самого оглавления.
    if offset is None:
        offset = toc_page_offset(project)

    for page_no, page_items in enumerate(pages):
        # Оглавление разбито по листам той же моделью высоты, что и записка:
        # следующий лист начинается с разрыва перед первым его абзацем.
        for item_index, item in enumerate(page_items):
            p = doc.add_paragraph()
            if page_no and item_index == 0:
                p.paragraph_format.page_break_before = True
            # Высота строки — как в пагинации оглавления (кегль × 1.4),
            # иначе стиль Normal добавляет отбивку и лист переполняется.
            p.paragraph_format.line_spacing = Pt(size * 1.4)
            p.paragraph_format.space_before = Pt(0)
            if item.is_title:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(12)
                _set_run(p.add_run(item.text), toc_font, title_size, bold=True)
                continue
            p.paragraph_format.space_after = Pt(0)
            align_left = WD_ALIGN_PARAGRAPH.LEFT
            p.alignment = align_left
            indent_mm = level_indent_mm(item.level)
            pf = p.paragraph_format
            # Внутренний отступ листа: строки оглавления начинаются там же,
            # где текст записки (26 мм от края листа, а не 20 мм).
            pf.left_indent = Mm(SHEET_PAD_MM + indent_mm)
            pf.right_indent = Mm(SHEET_PAD_MM)
            # Номер листа у правого края текстовой области с отточием; позиция
            # табуляции — от левого поля, поэтому отступ уровня вычитается.
            try:
                pf.tab_stops.add_tab_stop(
                    Mm(max(20.0, DOCX_TAB_POS_MM - indent_mm)),
                    WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            except Exception:
                pass
            # Строка оглавления: номер раздела и наименование — как обычный
            # текст записки (без полужирного), наименование — ссылка на
            # закладку заголовка (в DOCX и в PDF из Word она кликабельна).
            if item.num:
                _set_run(p.add_run(item.num + " "), toc_font, size)
            if item.anchor:
                _add_internal_link(p, item.text, item.anchor, toc_font, size)
            else:
                _set_run(p.add_run(item.text), toc_font, size)
            if item.page:
                _set_run(p.add_run("\t"), toc_font, size)
                # Номер листа — ПОЛЕ PAGEREF на закладку заголовка: Word сам
                # подставляет номер того листа, где заголовок оказался в
                # собранном файле. Литеральное число приходилось бы считать по
                # раскладке ИСХОДНОЙ записки, а собранный раздел занимает больше
                # листов — напечатанные номера отставали от настоящих.
                # `\h` делает номер гиперссылкой, `w:dirty` заставляет Word
                # пересчитать поле при открытии.
                cached = str(int(item.page) + max(0, int(offset or 0)))
                if item.anchor:
                    _field_runs(p, " PAGEREF %s \\h " % item.anchor,
                                cached=cached, font=toc_font, size=size)
                else:
                    _set_run(p.add_run(cached), toc_font, size)
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
                              set_cell_border, set_table_borders,
                              set_table_center, mm_to_twips)
    from ..models import TOC_WIDTHS

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run("СОДЕРЖАНИЕ"), font, fm.content_header, bold=True)

    rows = content_rows(project)    # [(наименование, обозначение, листы), ...]

    # Таблица содержания (3 колонки): Наименование | Обозначение | Листы.
    if rows:
        ncols = 3
        # +1 строка-заголовок («шапка») таблицы.
        t = doc.add_table(rows=len(rows) + 1, cols=ncols)
        t.autofit = False
        set_table_borders(t)
        set_table_center(t)
        set_table_fixed_layout(t, A4_MM[0] - 25.0)   # ~185 мм текстовой области
        # Ширины колонок — единый источник с предпросмотром и PDF
        # (models.TOC_WIDTHS = 95/50/40 мм, сумма 185 мм = вся ширина между
        # линиями рамки листа).
        set_table_grid_columns(t, [mm_to_twips(w) for w in TOC_WIDTHS])
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


def _add_composition(doc, project, fm, font):
    """Раздел «Состав проектной документации»: таблица томов на листах формы 5/6.

    Над таблицей печатается надпись «Состав проектной документации», под ней —
    таблица во всю ширину между линиями рамки (15 + 60 + 90 + 20 = 185 мм).
    Высота шапки 15 мм, высота строки 8 мм (models.COMPOSITION_HEAD_MM /
    COMPOSITION_ROW_MM). Разбиение на листы — по единому источнику
    (preview_html.composition_pages_from_rows), поэтому число листов и номера
    совпадают с предпросмотром.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from .docx_tables import (set_table_fixed_layout, set_table_grid_columns,
                              set_cell_border, set_row_height,
                              set_cell_vertical_alignment, set_table_borders,
                              set_table_center, mm_to_twips)
    from ..models import (COMPOSITION_LABELS, COMPOSITION_WIDTHS,
                          COMPOSITION_HEAD_MM, COMPOSITION_ROW_MM,
                          composition_rows)
    from .preview_html import COMPOSITION_TITLE, composition_pages_from_rows
    from .textfit import font_path

    # Надпись над таблицей — как на листе предпросмотра (единый источник).
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = 0
    title.paragraph_format.space_after = 0
    _set_run(title.add_run(COMPOSITION_TITLE), font, fm.content_header, bold=True)

    headers = list(COMPOSITION_LABELS)
    pages = composition_pages_from_rows(composition_rows(project))
    fpath = font_path(font)
    # Кегль надписей шапки считается с переносом по словам (узкая графа «Номер
    # тома» 15 мм), поэтому надпись выводится в две строки тем же кеглем, что и
    # «Обозначение», а не сжимается до 6.5 pt.
    head_avail = [max(1.0, w - _CELL_TEXT_PAD_MM * 2) for w in COMPOSITION_WIDTHS]
    head_sizes = [_wrap_size(label, head_avail[ci], fm.content_row, fpath)
                  for ci, label in enumerate(headers)]
    for page_no, page_rows in enumerate(pages):
        # Каждый следующий лист начинается с новой страницы; шапка таблицы есть
        # на каждом листе.
        t = doc.add_table(rows=len(page_rows) + 1, cols=len(headers))
        if page_no:
            p = OxmlElement("w:p")
            t._tbl.addprevious(p)
            from docx.text.paragraph import Paragraph as DocxParagraph
            from docx.shared import Pt
            paragraph_format = DocxParagraph(p, doc).paragraph_format
            paragraph_format.page_break_before = True
            paragraph_format.space_before = Pt(0)
            paragraph_format.space_after = Pt(0)
            paragraph_format.line_spacing = Pt(1)
        t.autofit = False
        set_table_borders(t)
        set_table_center(t)
        set_table_fixed_layout(t, sum(COMPOSITION_WIDTHS))
        set_table_grid_columns(t, [mm_to_twips(w) for w in COMPOSITION_WIDTHS])
        # Высоты строк — строго по сетке (exact). При `atLeast` строка росла от
        # переносов текста, таблица переставала влезать на лист, Word разрывал
        # её сам — и наш разрыв между листами давал пустую страницу.
        set_row_height(t.rows[0], COMPOSITION_HEAD_MM, rule="exact")
        for ci, label in enumerate(headers):
            cell = t.cell(0, ci)
            set_cell_border(cell, ("top", "bottom", "left", "right"))
            set_cell_vertical_alignment(cell, "center")
            cell.text = ""
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _tight_paragraph(para)
            size = head_sizes[ci]
            lines = _wrap_lines(label, size, head_avail[ci], fpath)
            _add_wrapped_run(para, lines, font, size, bold=True)
        for ri, values in enumerate(page_rows, start=1):
            set_row_height(t.rows[ri], COMPOSITION_ROW_MM, rule="exact")
            for ci, value in enumerate(values):
                cell = t.cell(ri, ci)
                set_cell_border(cell, ("top", "bottom", "left", "right"))
                set_cell_vertical_alignment(cell, "center")
                cell.text = ""
                para = cell.paragraphs[0]
                para.alignment = (WD_ALIGN_PARAGRAPH.CENTER if ci == 0
                                  else WD_ALIGN_PARAGRAPH.LEFT)
                _tight_paragraph(para)
                _fit_cell_run(para, str(value or ""), font, fm.content_row,
                              COMPOSITION_WIDTHS[ci], fpath)



def _add_graphic_sheet_list(doc, project, fm, font):
    """Лист «Графическая часть» с перечнем листов (перед листами графики).

    Заголовок, подзаголовок и таблица во всю ширину текстовой области:
    Обозначение | Наименование | Примечание. Строки берутся из единого
    источника (preview_html.build_graphic_sheet_list), поэтому Word, PDF и
    предпросмотр совпадают.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from .docx_tables import (set_table_fixed_layout, set_table_grid_columns,
                              set_cell_border, set_table_borders,
                              set_table_center, mm_to_twips)
    from .preview_html import build_graphic_sheet_list, _graphic_sheet_list_cfg
    from ..models import SHEETLIST_LABELS, SHEETLIST_WIDTHS

    cfg = _graphic_sheet_list_cfg(project)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run(cfg.title or "Графическая часть"), font,
             fm.content_header, bold=True)
    if cfg.subtitle:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run(cfg.subtitle), font, fm.content_row)

    rows = build_graphic_sheet_list(project)
    headers = list(SHEETLIST_LABELS)
    ncols = len(headers)
    t = doc.add_table(rows=len(rows) + 1, cols=ncols)
    t.autofit = False
    set_table_borders(t)
    set_table_center(t)
    set_table_fixed_layout(t, A4_MM[0] - 25.0)     # ~185 мм текстовой области
    # Ширины колонок — единый источник с предпросмотром и PDF (50/95/40 мм).
    set_table_grid_columns(t, [mm_to_twips(w) for w in SHEETLIST_WIDTHS])
    for ci, hdr in enumerate(headers):
        cell = t.cell(0, ci)
        set_cell_border(cell, ("top", "bottom", "left", "right"))
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(para.add_run(hdr), font, fm.content_row)
    for ri, (designation, name, note) in enumerate(rows, start=1):
        for ci, val in enumerate((designation, name, note)):
            cell = t.cell(ri, ci)
            set_cell_border(cell, ("top", "bottom", "left", "right"))
            cell.text = ""
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _set_run(para.add_run(val or ""), font, fm.content_row)
    if not rows:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run("Перечень листов не задан"), font, fm.content_row)


# Соответствие выравнивания ГОСТ-текста значениям python-docx.
_ALIGN_ENUM = {    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _align_paragraph_metrics(p, fmt, size) -> None:
    """Приводит высоту абзаца в Word к той, что считает пагинация HTML.

    Высота строки берётся из ЕДИНОГО источника — `preview_html.pz_line_mm`
    (то же значение получает и Word при разметке страниц в
    `word_pagination._normalized_copy`): интервал исходника, а при его
    отсутствии — измеренная высота строки (12 pt → 5.56 мм, 14 pt → 6.35 мм).
    Раньше здесь стоял коэффициент 0.5 × кегль (6.0 мм при 12 pt): абзацы в
    Word выходили выше, лист вмещал меньше строк, и Word рвал страницу там,
    где в предпросмотре разрыва нет.
    """
    from docx.shared import Pt
    pf = p.paragraph_format
    from .preview_html import pz_line_mm
    line_mm = pz_line_mm(size, fmt)
    pf.line_spacing = Pt(line_mm * 72.0 / 25.4)   # мм -> пункты
    try:
        pf.space_before = Pt(float(getattr(fmt, "space_before_pt", 0.0) or 0.0))
        pf.space_after = Pt(float(getattr(fmt, "space_after_pt", 0.0) or 0.0))
    except Exception:
        pass


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


def _pz_run_style(para, fm, font, text_fonts):
    """Шрифт и кегль абзаца ПЗ: (размер, шрифт) по группе абзаца."""
    group = "heading" if getattr(para, "is_heading", False) else "body"
    if text_fonts is not None:
        return text_fonts.size(group), (text_fonts.font(group) or font)
    size = fm.text_heading if getattr(para, "is_heading", False) else fm.text_body
    return size, font


def _table_run_style(fm, font, text_fonts):
    """Шрифт и кегль текста таблиц ПЗ (группа «table»)."""
    if text_fonts is not None:
        return text_fonts.size("table"), (text_fonts.font("table") or font)
    return fm.text_body, font


def _fit_widths_mm(widths, total_mm: float) -> list:
    """Пропорционально приводит ширины колонок к общей ширине total_mm.

    Таблицы ПЗ в исходном DOCX часто имеют `w:tblW=auto`: Word распределяет
    ширину сам, и таблица не доходит до линий рамки. Здесь ширины колонок
    берутся из исходной сетки (w:tblGrid) и масштабируются на всю ширину листа
    между линиями рамки (185 мм) — так же, как таблица выглядит в предпросмотре.
    """
    vals = [float(w) for w in (widths or []) if float(w) > 0]
    if not vals:
        n = max(1, len(widths or []))
        return [total_mm / n] * n
    k = total_mm / sum(vals)
    return [w * k for w in vals]


def _add_pz_table(doc, chunk, table_index, path, fm, font, text_fonts,
                  break_before: bool = False):
    """Добавляет на лист часть таблицы ПЗ (строка или несколько строк).

    Ширина и положение таблицы повторяют HTML-предпросмотр: таблица занимает
    ширину текстовой области (173 мм = 185 − 2 × 6 мм отступа) и стоит с
    отступом 6 мм от левой линии рамки. Внутренние отступы и высота строк
    совпадают с CSS-сеткой. Объединения ячеек исходника
    (w:gridSpan / w:vMerge) восстанавливаются в пределах попавших на лист строк.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from .docx_tables import (set_table_fixed_layout, set_table_grid_columns,
                              set_cell_border, set_table_cell_margins,
                              set_table_borders, set_table_center, mm_to_twips)
    from .preview_html import table_col_widths_mm

    if not chunk:
        return
    nrows = len(chunk)
    ncols = max((len(r) for r in chunk), default=1)
    if ncols <= 0:
        return
    try:
        widths = table_col_widths_mm(path, table_index, ncols)
    except Exception:
        widths = []
    widths = _fit_widths_mm(widths, STAMP_WIDTH_MM - 2 * SHEET_PAD_MM)

    size, run_font = _table_run_style(fm, font, text_fonts)
    cont_rows = 0
    nrows_all = nrows
    ncols_all = ncols

    t = doc.add_table(rows=nrows_all, cols=ncols_all)
    if break_before:
        p = OxmlElement("w:p")
        t._tbl.addprevious(p)
        from docx.text.paragraph import Paragraph as DocxParagraph
        from docx.shared import Pt
        paragraph_format = DocxParagraph(p, doc).paragraph_format
        paragraph_format.page_break_before = True
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)
        paragraph_format.line_spacing = Pt(1)
    t.autofit = False
    set_table_fixed_layout(t, sum(widths))
    set_table_grid_columns(t, [mm_to_twips(w) for w in widths])
    # Отступ таблицы от левой линии рамки — как padding .sheet-content.
    set_table_center(t)
    # Полная сетка на уровне таблицы: границы отдельных ячеек терялись у
    # объединённых, и часть линий в таблице отсутствовала.
    set_table_borders(t)
    # Небольшие внутренние отступы: текст не должен касаться линий графы.
    set_table_cell_margins(t, left_mm=2.0, right_mm=2.0,
                           top_mm=1.0, bottom_mm=1.0)

    covered = set()
    from .preview_html import (_estimate_table_row_h, _table_cell_chars_per_line,
                               _TABLE_CELL_PAD_MM, _TABLE_CELL_BORDER_MM,
                               _TABLE_CELL_LINE_MM, _table_scale)
    table_pt = float(size)
    cell_chars = [_table_cell_chars_per_line(ncols, i, widths[i], table_pt)
                  for i in range(ncols)]
    for ri, row in enumerate(chunk):
        row_height_mm = _estimate_table_row_h(
            row, max(cell_chars), cell_chars,
            col_widths=widths, size=table_pt, font=run_font)
        row_height_mm = max(row_height_mm, 4.3 * _table_scale(table_pt)
                            + _TABLE_CELL_PAD_MM + _TABLE_CELL_BORDER_MM)
        from .docx_tables import set_row_height
        set_row_height(t.rows[ri], row_height_mm, rule="exact")
        for ci, cell in enumerate(row):
            if ci >= ncols or (ri, ci) in covered:
                continue
            row_span = int(getattr(cell, "row_span", 1) or 1)
            col_span = int(getattr(cell, "col_span", 1) or 1)
            r2 = min(ri + row_span - 1, nrows_all - 1)
            c2 = min(ci + col_span - 1, ncols - 1)
            if row_span > 1 or col_span > 1:
                try:
                    t.cell(ri, ci).merge(t.cell(r2, c2))
                except (IndexError, ValueError):
                    pass
            # Ячейки внутри объединения уже закрыты им — иначе текст начала
            # объединения продублировался бы.
            for rr in range(ri, r2 + 1):
                for cc in range(ci, c2 + 1):
                    covered.add((rr, cc))
            if getattr(cell, "continuation", ""):
                continue
            tc = t.cell(ri, ci)
            set_cell_border(tc, ("top", "bottom", "left", "right"))
            tc.text = ""
            para = tc.paragraphs[0]
            alignment = getattr(cell, "alignment", "")
            para.alignment = _ALIGN_ENUM.get(alignment.lower(), WD_ALIGN_PARAGRAPH.LEFT)
            _tight_paragraph(para)
            from docx.shared import Pt
            para.paragraph_format.line_spacing = Pt(
                _TABLE_CELL_LINE_MM * _table_scale(table_pt) * 72.0 / 25.4)
            text = cell.text or ""
            if text:
                _set_run(para.add_run(text), run_font, size)


def _apply_sheet_indent(p, pad_mm: float = SHEET_PAD_MM) -> None:
    """Добавляет внутренний отступ листа к отступам абзаца.

    В HTML текст ПЗ стоит внутри `.sheet-content` с `padding:6mm`, поэтому он
    начинается на 26 мм от края листа (20 мм рамка + 6 мм отступ). В Word тот
    же отступ задаётся отступами слева/справа — к уже перенесённым отступам
    исходника (абзацный отступ, «висячий» отступ списка) прибавляется 6 мм.
    """
    from docx.shared import Mm
    pf = p.paragraph_format
    left = pf.left_indent.mm if pf.left_indent is not None else 0.0
    right = pf.right_indent.mm if pf.right_indent is not None else 0.0
    pf.left_indent = Mm(left + pad_mm)
    pf.right_indent = Mm(right + pad_mm)


def _apply_heading_style(paragraph, para) -> None:
    """Переносит в собранный документ стиль заголовка исходной записки.

    Стиль (`Heading 1` … `Heading 3`) сохраняется, хотя оформление задаётся
    наше: шрифт, кегль, отступы и интервалы проставляются явно и перекрывают
    параметры стиля шаблона. Благодаря стилю Word видит структуру документа —
    работает навигация по заголовкам и обновление поля оглавления, если его
    вставят вручную.

    Если стиля с таким именем в собранном документе нет (шаблон python-docx
    знает только английские имена), берётся встроенный `Heading N` по номеру
    из имени стиля исходника.
    """
    style_name = (getattr(para, "style_name", "") or "").strip()
    if not style_name:
        return
    styles = paragraph.part.document.styles
    candidates = [style_name]
    match = re.search(r"(\d+)", style_name)
    if match:
        candidates.append("Heading %s" % match.group(1))
        candidates.append("Заголовок %s" % match.group(1))
    for name in candidates:
        try:
            paragraph.style = styles[name]
            return
        except KeyError:
            continue
        except Exception:
            return


def _set_pz_pagination(paragraph, keep_with_next: bool = False) -> None:
    """Отключает переносы абзаца, не заданные исходной вёрсткой.

    keep_with_next=True — заголовок: он удерживается со следующим абзацем
    (`w:keepNext`), иначе Word оставлял бы заголовок последней строкой листа.
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    ppr = paragraph._p.get_or_add_pPr()
    for tag in ("w:keepNext", "w:widowControl"):
        el = ppr.find(qn(tag))
        if el is not None:
            ppr.remove(el)
    if keep_with_next:
        ppr.append(OxmlElement("w:keepNext"))
    widow = OxmlElement("w:widowControl")
    widow.set(qn("w:val"), "0")
    ppr.append(widow)


def _add_pz_page(doc, page_items, content, row_owner, path, fm, font,
                 text_fonts, break_before=False, toc_anchors=None):
    """Выводит один лист ПЗ: абзацы и таблицы в порядке исходного документа.

    :param toc_anchors: {индекс абзаца: имя закладки} — заголовки, на которые
        ссылаются строки оглавления: в них ставится закладка Word.
    """
    paragraphs = list(getattr(content, "paragraphs", []) or [])
    toc_anchors = toc_anchors or {}
    break_pending = break_before
    for item_index, (kind, value) in enumerate(page_items):
        if kind == "line":
            idx = int(value.get("para_index", -1))
            para = paragraphs[idx] if 0 <= idx < len(paragraphs) else None
            if para is None:
                # Абзац из «надписи» Word (from_textbox): адреса нет — выводим
                # только текст, без оформления исходника.
                p = doc.add_paragraph()
                _set_pz_pagination(p)
                if break_pending:
                    p.paragraph_format.page_break_before = True
                    break_pending = False
                _apply_sheet_indent(p)
                _set_run(p.add_run(value.get("text", "")), font, fm.text_body)
                continue
            size, run_font = _pz_run_style(para, fm, font, text_fonts)
            # Заголовок — только абзац стиля Word (см. pz_toc): выводится
            # полужирным, курсив — как в исходнике; обычный текст записки —
            # единый, без полужирного и курсива.
            is_heading = is_heading_paragraph(para)
            bold, italic = heading_emphasis(para)
            p = doc.add_paragraph()
            if is_heading:
                # Стиль заголовка сохраняется: Word видит структуру документа
                # (навигация, обновление поля оглавления), хотя оформление
                # задаётся наше — шрифт, кегль и отступы проставлены явно и
                # перекрывают стиль шаблона.
                _apply_heading_style(p, para)
            _set_pz_pagination(p, keep_with_next=is_heading)
            fmt = getattr(para, "fmt", None)
            if break_pending or bool(getattr(fmt, "page_break_before", False)):
                p.paragraph_format.page_break_before = True
                break_pending = False
            _apply_paragraph_format(p, fmt)
            _align_paragraph_metrics(p, fmt, size)
            _apply_sheet_indent(p)
            anchor = toc_anchors.get(idx)
            if anchor:
                _add_bookmark(p, anchor, 10000 + idx)
            _list_marker_run(p, fmt, size, run_font)
            for run in para.runs:
                _set_run(p.add_run(run.text), run_font, size, bold=bold,
                         italic=italic)
            continue

        # Строки таблицы. Владелец (индекс исходной таблицы) нужен для ширин
        # колонок — ищем по первой строке, известной карте.
        chunk = value
        table_index = -1
        for row in chunk or []:
            if row and id(row) in row_owner:
                table_index = row_owner[id(row)]
                break
        _add_pz_table(doc, chunk, table_index, path, fm, font, text_fonts,
                      break_before=break_pending)
        break_pending = False


def _add_pz_doc(doc, path, fm, font, text_fonts=None, first_form=None,
                break_before_first=False, toc_source="", toc_anchors=None):
    """Импортирует DOCX-исходник в ПЗ с переоформлением под ГОСТ.

    Разбиение на листы берётся из ЕДИНОГО источника — пагинации
    HTML-предпросмотра (`preview_html._paginate_docx`): она считает высоту
    абзацев и строк таблиц по тем же шрифтам, что и вывод, и решает, что
    попадёт на лист. Разрыв страницы ставится там же, где в предпросмотре
    начинается новый лист, поэтому Word не разливает текст по-своему и число
    листов совпадает с HTML.

    :param text_fonts: настройки шрифтов текста ПЗ (config.TextFonts): шрифт и
        кегль отдельно для обычного текста, заголовков и таблиц. None — берутся
        значения из карты кеглей (fm.text_body / fm.text_heading).
    :param first_form: форма ГОСТ первой страницы документа (FORM_TEXT_FIRST_WORD /
        FORM_TEXT_NEXT). От неё зависит бюджет первой страницы в пагинации.
    :param toc_source: путь к документу, по которому построено оглавление;
        закладки-цели ставятся только в нём (иначе ссылки вели бы в чужой файл).
    :param toc_anchors: {индекс абзаца: имя закладки} — цели гиперссылок.
    :return: число выведенных листов (0 — документ не прочитан).
    """
    if not path or not os.path.exists(path):
        return 0
    try:
        content = read_docx(path)
    except Exception:
        return 0
    try:
        from .preview_html import _paginate_docx
    except Exception:
        return 0
    # Формы в терминах предпросмотра: 2 — большая надпись первого листа (40 мм,
    # это и форма 3 в Word, и форма 5 в предпросмотре), "2a" — форма 6 (15 мм).
    pz_first = (_PZ_FIRST_FORM if first_form in (None, FORM_TEXT_FIRST_WORD)
                else _PZ_NEXT_FORM)
    from .word_pagination import cached_docx_pages
    word_layout = cached_docx_pages(path, pz_first)
    pages = _paginate_docx(content, path, first_form=pz_first,
                            word_layout=word_layout)

    # Карта «строка таблицы -> индекс исходной таблицы» (для ширин колонок).
    row_owner = {}
    for ti, tbl in enumerate(getattr(content, "tables", []) or []):
        for row in getattr(tbl, "rows", []) or []:
            row_owner[id(row)] = ti

    is_toc_source = bool(toc_source) and (os.path.abspath(path)
                                          == os.path.abspath(toc_source))
    anchors = dict(toc_anchors or {}) if is_toc_source else {}

    for page_no, page_items in enumerate(pages):
        _add_pz_page(
            doc, page_items, content, row_owner, path, fm, font, text_fonts,
            break_before=bool(page_no or (break_before_first and page_no == 0)),
            toc_anchors=anchors)
    return len(pages)


def _add_stamp_footers(doc, project, config, sec=None, first_form=None,
                       default_form=None, section_kind=None):
    """Добавляет основную надпись по ГОСТ (форма 3 / 6) в нижний колонтитул.

    :param first_form: форма для первого листа раздела.
        None  — не использовать «первый лист» отдельно (все страницы — форма 6);
        FORM_TEXT_FIRST_WORD — первый лист форма 3, последующие — форма 6.
    Если first_form не задан — используются только последующие (форма 6).

    :param default_form: форма для обычных (не первых) страниц. По умолчанию
        FORM_TEXT_NEXT. Используется, когда нужен отдельный лист с формой 3/5.
    :param sec: раздел, для которого строятся колонтитулы/рамка. По умолчанию 0-й.
    :param section_kind: ключ раздела тома — по нему в графу «Обозначение»
        дописывается дополнительное обозначение («.ПЗ», «.ГЧ», «.С»).
    """
    if sec is None:
        if not doc.sections:
            return
        sec = doc.sections[0]
    req = project.requisites
    fm: FontMap = config.font_map
    font = config.font_name
    default_form = default_form if default_form is not None else FORM_TEXT_NEXT
    # Обозначение в графе 1: том + код раздела (если раздел известен).
    designation = (project.designation(section_kind) if section_kind
                   else (req.volume_id or ""))

    # Строки подписей берутся из модели по разделу и форме: значения по
    # умолчанию + правки, сделанные в предпросмотре. Форма 6 графы подписей
    # не содержит, поэтому для неё строки не нужны.
    from ..models import stamp_signature_rows
    sig5 = stamp_signature_rows(project, section_kind, 5) if section_kind else []
    sig3 = stamp_signature_rows(project, section_kind, 3) if section_kind else []

    # --- Текстовая область как в HTML: отступ сверху и граница снизу ---
    # Высота штампа — по форме первой/обычной страницы: от неё зависит нижняя
    # граница текста (в HTML она считается так же).
    stamp_h = max((_form_height_mm(f) for f in (first_form, default_form)
                   if f is not None), default=0.0)
    _apply_sheet_margins(sec, config, stamp_h)

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
                     stamp9=getattr(project, "stamp_cell9", None),
                     designation=designation,
                     signature_rows=(sig3 if default_form in (FORM_GRAPHIC, 3)
                                     else sig5),
                     section_sheets=_section_sheet_count(project, section_kind))

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
                         stamp9=getattr(project, "stamp_cell9", None),
                         designation=designation,
                         signature_rows=(sig3 if first_form in (FORM_GRAPHIC, 3)
                                         else sig5),
                         section_sheets=_section_sheet_count(project, section_kind))
    else:
        sec.different_first_page_header_footer = False

    # --- Рамка страницы (линия по периметру) через page border ---
    _set_page_frame(sec, config.margins)

    # Верхний отступ — для всех страниц раздела: поле одно, колонтитул один.
    _set_top_indent(sec.header)
    if first_form is not None:
        _set_top_indent(sec.first_page_header)

    # --- Квадрат сквозной нумерации в правом верхнем углу ---
    # Графы «Лист»/«Листов» штампа ведутся В ПРЕДЕЛАХ раздела, поэтому нумерация
    # секции перезапускается с 1 (w:pgNumType). Квадрат в углу — сквозной по
    # тому, поэтому в него ставится формула с базой раздела.
    _set_pgnum_start(sec, 1)
    base = _section_square_base(project, section_kind)
    _add_square_box(sec.header, base=base, page_w_mm=A4_MM[0])
    if first_form is not None:
        _add_square_box(sec.first_page_header, base=base, page_w_mm=A4_MM[0])
    # PAGE и SECTIONPAGES пересчитываются самим Word при пагинации.
    # Глобальный updateFields также запускал бы обновление внешних LINK-полей,
    # из-за которых Word показывает запрос на обновление ссылок при открытии.


def _form_height_mm(form) -> float:
    """Высота основной надписи (мм) по ключу формы ГОСТ."""
    from .geometry import FORM_HEIGHT
    if form in (FORM_TEXT_FIRST, 5):
        return FORM_HEIGHT[5]
    if form in (FORM_TEXT_NEXT, 6):
        return FORM_HEIGHT[6]
    if form in (FORM_GRAPHIC, 3):
        return FORM_HEIGHT[3]
    return FORM_HEIGHT[5]



def _clear_footer(footer):
    """Удаляет все абзацы (пустой колонтитул)."""
    for p in list(footer.paragraphs):
        p._element.getparent().remove(p._element)


#: Графы штампа, текст которых выравнивается по ЛЕВОМУ краю, как в
#: HTML-предпросмотре: таблица изменений и строки подписей левого блока.
_STAMP_LEFT_ROLES = ("change_hdr", "change_val",
                     "sig_post", "sig_name", "sig_date")


def _stamp_cell_align(role: str, form_num: int) -> tuple:
    """Выравнивание текста графы штампа: (сторона, отступ слева в мм).

    Левый блок (таблица изменений и строки подписей) в HTML-предпросмотре
    выровнен по левому краю с внутренним отступом; остальные графы — по
    центру. Word центрировал всё, и «Разраб. / Максимов» стояли иначе, чем
    в предпросмотре.

    :param form_num: форма ГОСТ (3/5/6) — у формы 6 графы узкие (10 мм), поэтому
        отступ меньше, иначе «Кол.уч» не помещается в строку.
    """
    if role not in _STAMP_LEFT_ROLES:
        return "center", 0.0
    return "left", 0.5 if form_num == 6 else 1.0


def _add_stamp_table(container, stamp_w_mm, form, req, fm, font,
                     set_table_fixed_layout, set_table_grid_columns,
                     set_cell_border, set_row_height,
                     set_cell_vertical_alignment, mm_to_twips, stamp9=None,
                     designation=None, signature_rows=None, section_sheets=1):
    """Строит таблицу-основную надпись в нижнем колонтитуле по ГОСТ.

    Геометрия (сетка, разбивка колонок, высоты строк) берётся из
    app.core.geometry.grid — того же источника, что у HTML-предпросмотра и PDF:
    формы 3 / 5 / 6 по ГОСТ Р 21.101-2020, ширина 185 мм (между линиями рамки).
    Готовых шаблонов DOCX больше нет: рамка, линии и графы рисуются таблицей,
    поэтому любую правку геометрии достаточно сделать в geometry.py.

    :param container: футер (BlockItemContainer), куда добавляется таблица
    :param form: номер формы ГОСТ (3/5/6) либо FORM_TEXT_FIRST/FORM_TEXT_NEXT.
    :param designation: значение графы «Обозначение» (том + код раздела);
        None — обозначение тома из реквизитов.
    :param stamp9: настройки графы 9 (models.StampCell9) — текст по центру ячейки
        либо логотип организации из PDF (вписывается в ячейку без искажений).
    :param signature_rows: строки подписей левого блока (models.stamp_signature_rows):
        должность, фамилия и дата каждой строки. None — строки не заполняются.
    """
    from docx.shared import Mm
    from .geometry import grid, STAMP_WIDTH_MM
    from .docx_tables import (merge_cells, set_table_cell_margins,
                              set_cell_no_wrap, set_table_borders,
                              set_stamp_cell_borders)
    from .textfit import fit_size, font_path

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
    # Границы всей таблицы гасятся: сетка рисуется по ячейкам (`set_stamp_cell_borders`),
    # потому что у вертикально объединённых ячеек Word берёт границу из tcPr
    # КАЖДОЙ строки объединения и границы уровня таблицы давали рваную сетку.
    set_table_borders(table, edges=())
    # Внутренние отступы ячеек обнуляются: иначе в графе 10 мм на текст
    # остаётся ~6 мм, значение переносится, и штамп «разъезжается».
    set_table_cell_margins(table)

    # Задаём высоты строк (мм). Правило exact: строка не может вырасти —
    # геометрия штампа должна совпадать с сеткой при любом тексте.
    for ri in range(nrows):
        set_row_height(table.rows[ri], row_mm[ri], rule="exact")

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

    # Сетка по ячейкам. Обход идёт по XML (каждая строка объединения), поэтому
    # линии доходят до низа и до правой рамки у всех блоков, а стороны,
    # совпадающие с рамкой листа, не дублируются.
    set_stamp_cell_borders(table, [(r0, c0, c1, r1)
                                   for r0, c0, c1, r1, _ro, _k, _l in merged_cells],
                           ncols, nrows)

    # Заполняем содержимое.
    fpath = font_path(font)
    for row0, col0, col1, row1, role, key, label in merged_cells:
        cell = table.cell(row0, col0)
        # Границы НЕ ставятся по ячейке: сетка задана на уровне таблицы
        # (`set_table_borders`). При поячеечных границах Word рисует вторую
        # линию рядом с линией таблицы (и с линией рамки) — видно задвоение,
        # а у объединённых ячеек часть линий вообще пропадала.
        set_cell_vertical_alignment(cell, "center")
        set_cell_no_wrap(cell)
        cell.text = ""
        para = cell.paragraphs[0]
        align, pad_mm = _stamp_cell_align(role, form_num)
        para.alignment = (WD_ALIGN_PARAGRAPH.LEFT if align == "left"
                          else WD_ALIGN_PARAGRAPH.CENTER)
        # Интервалы и межстрочный — строго по высоте строки: иначе абзац
        # стиля Normal добавляет отбивку и строка выходит за сетку.
        _tight_paragraph(para, sum(row_mm[row0:row1]))
        # Отступ слева у граф с текстом по левому краю (как `pad` в HTML).
        if pad_mm:
            from docx.shared import Mm as _Mm
            para.paragraph_format.left_indent = _Mm(pad_mm)
        # Графа 9: логотип организации вместо текста (если выбран режим «логотип»).
        cell_w_mm = sum(col_mm[col0:col1])
        cell_h_mm = sum(row_mm[row0:row1])
        if (role == "field" and key in CELL9_KEYS
                and _insert_cell9_logo(para, stamp9, cell_w_mm, cell_h_mm)):
            continue
        # Графа «Подп.» строки подписи: изображение подписи подписанта,
        # подобранного по фамилии строки. Если подписи нет — ячейка пустая.
        if role == "sig_sign":
            row = _signature_row(signature_rows, key)
            if _insert_signature_image(para, row, cell_w_mm, cell_h_mm):
                continue
        # Графы «Лист» / «Листов»: значения ставит Word полями, а не Python —
        # номер листа и число листов раздела зависят от разбивки на страницы.
        # Форма 6: единственная ячейка sheet_val. Формы 3 и 5: поля field
        # с ключами sheet и qty (не путать с колонкой «Лист» таблицы изменений —
        # у неё роль change_val, её заполняет пользователь).
        if role == "sheet_val" or (role == "field" and key in ("sheet", "qty")):
            _add_sheet_number(
                para, "sheet" if key != "qty" else "total",
                font=font, size=STAGE_CELL_PT,
                total=section_sheets)
            continue
        text, size = _stamp_cell_value(role, key, label, req, fm,
                                       designation=designation,
                                       signature_rows=signature_rows)
        if text:
            # Кегль подбирается по фактической ширине графы (метрики шрифта):
            # «Стадия», «17.11.26», «Кол.уч» в узких графах при 10–12 pt
            # шире ячейки и переносились. Отступ слева вычитается: без этого
            # «Кол.уч» в графе 10 мм обрезался до «Кол.у».
            size = fit_size(text, cell_w_mm - pad_mm - _STAMP_TEXT_PAD_MM, size,
                            path=fpath)
            _set_run(para.add_run(text), font, size)


#: Запас по ширине графы штампа (мм) при подборе кегля: текст не должен
#: касаться линий сетки.
_STAMP_TEXT_PAD_MM = 0.3


# --------------------------------------------------------------------------
# Номер листа в штампе — поле PAGE Word; количество листов берётся из общей
# раскладки тома (`toc.build_layout`), как в HTML, PDF и содержании.
# --------------------------------------------------------------------------
#
# Почему поля, а не готовые числа:
#   * графы «Лист»/«Листов» считаются В ПРЕДЕЛАХ РАЗДЕЛА (так же, как в
#     HTML-предпросмотре и PDF), поэтому в каждой секции нумерация
#     перезапускается (w:pgNumType w:start=1);
#   * квадрат в правом верхнем углу — СКВОЗНОЙ по тому, поэтому он ставится
#     формулой { = PAGE + base - 1 }, где base — первый сквозной номер листа
#     раздела по toc.build_layout().


def _field_runs(paragraph, instr, cached="", font=None, size=None):
    """Вставляет в абзац простое поле Word `{ instr }`.

    :param instr: инструкция поля без фигурных скобок, например " PAGE ".
    :param cached: отображаемое значение до пересчёта. Word пересчитает поле
        при открытии документа и при экспорте в PDF, поэтому это значение
        нужно только для предпросмотра в приложениях, не считающих поля.
    """
    from docx.shared import Pt

    def _run():
        run = paragraph.add_run()
        if font and size:
            # Через _set_run — чтобы поле получило тот же rFonts (ascii/hAnsi/cs),
            # что и остальной текст штампа: иначе цифра в графе выпадает из
            # ISOCPEUR и печатается подставным шрифтом.
            _set_run(run, font, size)
        elif font:
            run.font.name = font
        elif size:
            run.font.size = Pt(size)
        return run

    begin = _run()
    el = OxmlElement("w:fldChar")
    el.set(qn("w:fldCharType"), "begin")
    el.set(qn("w:dirty"), "true")
    begin._r.append(el)

    instr_run = _run()
    el = OxmlElement("w:instrText")
    el.set(qn("xml:space"), "preserve")
    el.text = instr
    instr_run._r.append(el)

    sep = _run()
    el = OxmlElement("w:fldChar")
    el.set(qn("w:fldCharType"), "separate")
    sep._r.append(el)

    value = _run()
    el = OxmlElement("w:t")
    el.text = cached or ""
    value._r.append(el)

    end = _run()
    el = OxmlElement("w:fldChar")
    el.set(qn("w:fldCharType"), "end")
    end._r.append(el)
    return paragraph


def _add_sheet_number(paragraph, key, font=None, size=None, total=1):
    """Ставит номер листа полем Word или число листов из общей раскладки."""
    if key == "total":
        from docx.shared import Pt
        run = paragraph.add_run(str(max(0, int(total))))
        if font and size:
            _set_run(run, font, size)
        elif font:
            run.font.name = font
        elif size:
            run.font.size = Pt(size)
        return paragraph
    return _field_runs(paragraph, " PAGE ", cached="1", font=font, size=size)


def _add_square_number(paragraph, base, font=None, size=None):
    """Ставит в абзац СКВОЗНОЙ номер листа формулой `{ = PAGE + base - 1 }`.

    Формула Word не умеет ссылаться на поле PAGE по имени — номер страницы
    должен быть ВЛОЖЕННЫМ полем внутри формулы. Если записать инструкцию
    строкой "= PAGE + N", Word ищет закладку PAGE и печатает
    «!Неопределенная закладка». Поэтому строится составное поле:
    begin → "=" → begin → "PAGE" → end → "+ N" → end.

    :param base: первый сквозной номер листа раздела по toc.build_layout().
        При base <= 1 номер совпадает с PAGE, и формула не нужна.
    """
    if base is None or base <= 1:
        return _field_runs(paragraph, " PAGE ", cached="1", font=font, size=size)

    from docx.shared import Pt

    def _run():
        run = paragraph.add_run()
        if font and size:
            _set_run(run, font, size)
        elif size:
            run.font.size = Pt(size)
        return run

    def _fld(kind):
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), kind)
        if kind == "begin":
            el.set(qn("w:dirty"), "true")
        _run()._r.append(el)

    def _instr(text):
        el = OxmlElement("w:instrText")
        el.set(qn("xml:space"), "preserve")
        el.text = text
        _run()._r.append(el)

    def _text(text):
        el = OxmlElement("w:t")
        el.text = text
        _run()._r.append(el)

    _fld("begin")            # внешняя формула
    _instr(" = ")
    _fld("begin")            # вложенное поле PAGE
    _instr(" PAGE ")
    _fld("separate")
    _text("1")
    _fld("end")
    _instr(" + %d " % (base - 1))
    _fld("separate")
    _text(str(base))
    _fld("end")
    return paragraph


def _set_pgnum_start(section, start=1):
    """Перезапускает нумерацию страниц секции с `start` (w:pgNumType).

    Без этого поля PAGE/SECTIONPAGES продолжали бы сквозную нумерацию тома, а
    графы «Лист»/«Листов» по ГОСТ ведутся в пределах раздела.
    """
    sectPr = section._sectPr
    for el in sectPr.findall(qn("w:pgNumType")):
        sectPr.remove(el)
    el = OxmlElement("w:pgNumType")
    el.set(qn("w:start"), str(int(start)))
    anchor = sectPr.find(qn("w:pgMar"))
    if anchor is not None:
        anchor.addnext(el)
    else:
        sectPr.insert(0, el)


def _add_square_box(header, base=None, page_w_mm: float = 210.0) -> None:
    """Рисует квадрат сквозной нумерации в правом верхнем углу листа.

    Квадрат — плавающая таблица 10x10 мм в ВЕРХНЕМ колонтитуле, привязанная к
    краю листа (`vertAnchor`/`horzAnchor` = page), поэтому он не сдвигает текст
    и не зависит от полей раздела. Геометрия — как у HTML-предпросмотра и PDF
    (`preview_html.SHEET_NUM_*`, `renderer.draw_sheet_number_box`): сторона
    10 мм на расстоянии 5 мм от верхнего и правого края, то есть 195…205 x
    5…15 мм на A4.

    Внутри — поле Word, поскольку секция может занимать несколько страниц.
    Для сквозной нумерации используется база раздела из `toc.build_layout`.
    """
    from docx.shared import Mm, Pt
    from .preview_html import SHEET_NUM_BOX_MM, SHEET_NUM_OFFSET_MM
    from .renderer import SHEET_NUMBER_FONT_PT
    from .docx_tables import (LINE_BORDER_SZ, set_row_height,
                              set_cell_vertical_alignment)

    box = SHEET_NUM_BOX_MM
    left_mm = page_w_mm - SHEET_NUM_OFFSET_MM - box
    top_mm = SHEET_NUM_OFFSET_MM
    tw = 56.6929  # twips в 1 мм

    tbl = header.add_table(rows=1, cols=1, width=Mm(box))
    tblPr = tbl._tbl.tblPr
    pos = OxmlElement("w:tblpPr")
    pos.set(qn("w:leftFromText"), "0")
    pos.set(qn("w:rightFromText"), "0")
    pos.set(qn("w:topFromText"), "0")
    pos.set(qn("w:bottomFromText"), "0")
    pos.set(qn("w:vertAnchor"), "page")
    pos.set(qn("w:horzAnchor"), "page")
    pos.set(qn("w:tblpX"), str(int(round(left_mm * tw))))
    pos.set(qn("w:tblpY"), str(int(round(top_mm * tw))))
    tblPr.insert(0, pos)
    # Квадрат не должен расталкивать текст колонтитула: обтекание запрещено.
    overlap = OxmlElement("w:tblOverlap")
    overlap.set(qn("w:val"), "never")
    tblPr.insert(1, overlap)

    # Сетка квадрата: линия той же толщины, что рамка листа и штамп.
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(LINE_BORDER_SZ))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.insert_element_before(
        borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        "w:tblCaption", "w:tblDescription")

    # Внутренние отступы ячейки нулевые: иначе цифра смещается от центра.
    mar = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.insert_element_before(mar, "w:tblLook", "w:tblCaption",
                               "w:tblDescription")

    tbl.autofit = False
    set_row_height(tbl.rows[0], box, rule="exact")

    cell = tbl.cell(0, 0)
    cell.text = ""
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _tight_paragraph(para, box)
    set_cell_vertical_alignment(cell, "center")
    _add_square_number(para, base or 1, font=FONT_NAME, size=SHEET_NUMBER_FONT_PT)


def _section_sheet_count(project, section_kind):
    """Количество листов раздела из общей карты `toc.build_layout`."""
    if not section_kind:
        return 0
    from .toc import build_layout
    try:
        layout = build_layout(project)
    except Exception:
        return 0
    return sum(item.count for item in layout.items
               if item.section_key == section_kind)


def _section_square_base(project, section_kind):
    """Первый сквозной номер листа раздела по toc.build_layout().

    None — раздел не найден в карте тома (тогда квадрат считается локально,
    то есть с первого листа раздела).
    """
    if not section_kind:
        return None
    from .toc import build_layout
    try:
        layout = build_layout(project)
    except Exception:
        return None
    for item in layout.items:
        if item.section_key == section_kind:
            return item.first_sheet
    return None


def _tight_paragraph(para, row_h_mm: float = 0.0) -> None:
    """Убирает отбивки и растянутый межстрочный интервал у абзаца ячейки.

    Стиль Normal в шаблоне python-docx задаёт отбивку после абзаца и
    межстрочный 1.08 — в строке высотой 5 мм это выталкивает текст за сетку
    и увеличивает высоту строки, из-за чего штамп разъезжается.
    """
    from docx.shared import Pt
    from docx.enum.text import WD_LINE_SPACING
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    try:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    except Exception:
        pass
    pf.first_line_indent = Pt(0)
    pf.left_indent = Pt(0)
    pf.right_indent = Pt(0)


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


def _signature_row(signature_rows, key) -> dict:
    """Строка подписей штампа по номеру (key — номер строки строкой)."""
    if not signature_rows:
        return {}
    try:
        index = int(key)
    except (TypeError, ValueError):
        return {}
    for row in signature_rows:
        if int(row.get("row", -1)) == index:
            return row
    return {}


def _signature_image_path(row: dict) -> str:
    """Путь к изображению подписи строки: по фамилии ищется в реестре томов."""
    if not row:
        return ""
    path = str(row.get("signature_image") or "")
    return path if path and os.path.exists(path) else ""


def _image_size_px(path: str) -> tuple:
    """Пиксельные размеры изображения (0,0 — если не удалось прочитать).

    Pillow есть не во всех окружениях проекта, поэтому размеры читаются
    сначала им, затем PyMuPDF (он уже используется для логотипа графы 9).
    """
    try:
        from PIL import Image
        with Image.open(path) as im:
            return im.size
    except Exception:
        pass
    try:
        import pymupdf
        with pymupdf.open(path) as im:
            rect = im[0].rect
            return rect.width, rect.height
    except Exception:
        return (0, 0)


def _insert_signature_image(para, row: dict, cell_w_mm: float,
                            cell_h_mm: float) -> bool:
    """Вставляет изображение подписи подписанта в ячейку графы «Подп.».

    В ячейку вписывается САМ РОСЧЕРК: у картинки обрезаются белые поля
    (`preview_html.signature_fit_mm` считает размер по обрезанному изображению).
    Иначе скан подписи с полями уменьшался целиком, и росчерк в ячейке 15 × 5 мм
    получался в несколько миллиметров — так же, как было в предпросмотре.
    Возвращает True, если изображение вставлено; False — подписи нет.
    """
    path = _signature_image_path(row)
    if not path:
        return False
    try:
        from docx.shared import Mm, Pt
    except Exception:
        return False
    fit_w = fit_h = 0.0
    image_stream = None
    try:
        from io import BytesIO
        from .preview_html import _signature_pixmap, signature_fit_mm
        fit_w, fit_h = signature_fit_mm(path, cell_w_mm, cell_h_mm)
        processed = _signature_pixmap(path)
        if processed is not None:
            image_stream = BytesIO(processed[0].tobytes("png"))
    except Exception:
        fit_w = fit_h = 0.0
    if fit_w <= 0 or fit_h <= 0:
        # Обработка недоступна — используем исходное изображение как запасной путь.
        px_w, px_h = _image_size_px(path)
        if px_w <= 0 or px_h <= 0:
            return False
        # Отступы по 0.5 мм: подпись не должна касаться линий графы.
        avail_w = max(1.0, float(cell_w_mm) - 1.0)
        avail_h = max(1.0, float(cell_h_mm) - 1.0)
        scale = min(avail_w / px_w, avail_h / px_h)
        fit_w, fit_h = px_w * scale, px_h * scale
    if fit_w <= 0 or fit_h <= 0:
        return False
    try:
        pf = para.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
    except Exception:
        pass
    try:
        para.add_run().add_picture(
            image_stream if image_stream is not None else path,
            width=Mm(fit_w), height=Mm(fit_h))
        return True
    except Exception:
        return False


def _stamp_cell_value(role, key, label, req, fm, designation=None,
                      signature_rows=None):
    """Возвращает (текст, кегль) для ячейки основной надписи по её роли.

    :param signature_rows: строки подписей левого блока (models.stamp_signature_rows).
        Роли sig_post/sig_name/sig_sign/sig_date берут значение строки по её
        номеру (key) — должность, фамилию, подпись и дату.
    """
    if role == "change_hdr":
        return label, fm.stamp_columns
    if role == "change_val":
        # Значения таблицы изменений пока пустые (заполняются вручную).
        return "", fm.stamp_columns
    # Строки подписей штампа: номер строки приходит в key.
    if role in ("sig_post", "sig_name", "sig_sign", "sig_date"):
        row = _signature_row(signature_rows, key)
        if role == "sig_post":
            return str(row.get("post") or ""), fm.stamp_columns
        if role == "sig_name":
            return str(row.get("surname") or ""), fm.stamp_columns
        if role == "sig_date":
            return str(row.get("date") or ""), fm.stamp_columns
        # Графа «Подп.»: изображение подписи вставляется отдельно (см. ниже).
        return "", fm.stamp_columns
    if role == "design":
        # Графа 1: обозначение тома + дополнительное обозначение раздела.
        if designation is not None:
            return designation, fm.stamp_volume
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


def _set_page_frame(sec, margins=None):
    """Устанавливает рамку страницы (ГОСТ Р 21.101-2020) по периметру.

    Рамка привязывается к ГРАНИЦЕ ТЕКСТА (`offsetFrom="text"`, все отступы 0):
    поля листа заданы по линиям рамки (20 мм слева, по 5 мм с других сторон),
    поэтому линия проходит ровно по ним — 20..205 мм по ширине и 5..292 мм по
    высоте (проверено сборкой в Word).

    Привязка к краю листа (`offsetFrom="page"`) для этого листа не работает:
    Word ограничивает отступ рамки 31 pt (≈11 мм), поэтому левая линия
    вставала на 8.8 мм вместо 20 мм и не совпадала с таблицей шириной 185 мм.
    """
    from docx.oxml.ns import qn
    from .docx_tables import LINE_BORDER_SZ
    sectPr = sec._sectPr
    # удаляем существующую рамку (если есть)
    for el in sectPr.findall(qn("w:pgBorders")):
        sectPr.remove(el)

    pgBorders = OxmlElement("w:pgBorders")
    pgBorders.set(qn("w:offsetFrom"), "text")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        # Толщина — как у границ ячеек штампа (docx_tables.set_cell_border, sz=6
        # = 0.75 pt): рамка листа и основная надпись должны быть одной толщины,
        # иначе на стыке видно утолщение.
        el.set(qn("w:sz"), str(LINE_BORDER_SZ))
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
        _add_stamp_footers(doc, project, config, sec,
                           first_form=FORM_TEXT_FIRST_WORD,
                           section_kind=SECTION_CONTENT)
    elif sheet_kind == SECTION_COMPOSITION:
        _add_composition(doc, project, fm, font)
        _add_stamp_footers(doc, project, config, sec,
                           first_form=FORM_TEXT_FIRST_WORD,
                           section_kind=SECTION_COMPOSITION)
    elif sheet_kind == "changes":
        _add_changes_sheet(doc, project, fm, font)
        _add_stamp_footers(doc, project, config, sec,
                           first_form=FORM_TEXT_FIRST_WORD,
                           section_kind=SECTION_CHANGES)
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

    return _save_docx(doc, out_path)


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


def _add_stamp_sheet(doc, project, config, sec, form, section_kind=None):
    """Отдельный лист с основной надписью (форма 3/5/6) — как таблица внизу.

    :param section_kind: раздел, чьи строки подписей показывает штамп. Для
        отдельного листа формы 3 это графическая часть (по умолчанию).
    """
    from docx.shared import Mm
    if section_kind is None:
        from ..models import SECTION_GRAPHIC
        section_kind = SECTION_GRAPHIC if form == FORM_GRAPHIC else SECTION_CONTENT
    # Заголовок-метка формы
    label = {FORM_GRAPHIC: "ФОРМА 3 — ГРАФИЧЕСКИЙ ЛИСТ",
             FORM_TEXT_FIRST: "ФОРМА 5 — ПЕРВЫЙ ЛИСТ ТЕКСТА",
             FORM_TEXT_NEXT: "ФОРМА 6 — ПОСЛЕДУЮЩИЙ ЛИСТ ТЕКСТА"}.get(form, "")
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(p.add_run(label), config.font_name, config.font_map.content_header, bold=True)
    # Рамка и основная надпись — в нижнем колонтитуле
    _add_stamp_footers(doc, project, config, sec, first_form=None, default_form=form,
                       section_kind=section_kind)


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
        _add_stamp_footers(doc, project, config, sec,
                           first_form=FORM_TEXT_FIRST_WORD,
                           section_kind=SECTION_CONTENT)
    elif section_kind == SECTION_COMPOSITION:
        _add_composition(doc, project, fm, font)
        _add_stamp_footers(doc, project, config, sec,
                           first_form=FORM_TEXT_FIRST_WORD,
                           section_kind=SECTION_COMPOSITION)
    elif section_kind == "cover":
        _add_cover(doc, project, fm, font)
        _no_stamp_section(sec)
    elif section_kind == "title":
        _add_title_sheet(doc, project, fm, font)
        _no_stamp_section(sec)
    elif section_kind == "changes":
        _add_changes_sheet(doc, project, fm, font)
        _add_stamp_footers(doc, project, config, sec,
                           first_form=FORM_TEXT_FIRST_WORD,
                           section_kind=SECTION_CHANGES)
    elif section_kind in (DocKind.POYASNITELNAYA.value, DocKind.GRAPHIC.value):
        # Пояснительная записка / Графическая часть — все документы раздела.
        dkind = DocKind(section_kind)
        if dkind == DocKind.POYASNITELNAYA:
            # ПЗ: первый лист — форма 3, последующие — форма 6.
            # Раздел собирается отдельно — нумерация листов начинается с 1.
            toc_pages = _add_pz_toc(doc, project, fm, font,
                                    text_fonts=getattr(config, "text_fonts", None))
            _add_section_docs(doc, project, dkind, fm, font,
                              text_fonts=getattr(config, "text_fonts", None),
                              first_form=(FORM_TEXT_NEXT if toc_pages
                                          else FORM_TEXT_FIRST_WORD),
                              break_before_first=bool(toc_pages))
            _add_stamp_footers(
                doc, project, config, sec,
                first_form=FORM_TEXT_FIRST_WORD,
                section_kind=section_kind)
        else:
            # Графическая часть в Word — только генерируемый лист «Графическая
            # часть» с перечнем листов (форма 3). Сами листы — исходные PDF:
            # их страницы в DOCX не переносятся, поэтому раздела с «документами»
            # здесь нет (раньше он печатал «Документы отсутствуют», хотя листы
            # в томе есть).
            from .preview_html import _graphic_sheet_list_cfg
            gl = _graphic_sheet_list_cfg(project)
            if gl.enabled:
                _add_graphic_sheet_list(doc, project, fm, font)
                _add_stamp_footers(doc, project, config, sec,
                                   first_form=FORM_GRAPHIC,
                                   default_form=FORM_GRAPHIC,
                                   section_kind=section_kind)
            else:
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _set_run(p.add_run("Листы графической части — исходные PDF; "
                                   "в Word они не выводятся. Включите перечень "
                                   "листов, чтобы получить лист «Графическая "
                                   "часть»."), font, fm.content_row)

    elif section_kind == "appendix":
        _add_appendix_section(doc, project, fm, font)
    else:
        # Неизвестный раздел — заголовок.
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run(SECTION_LABELS.get(section_kind, section_kind)),
                 font, fm.content_header, bold=True)

    return _save_docx(doc, out_path)


def _add_section_docs(doc, project, dkind, fm, font, text_fonts=None,
                      first_form=None, break_before_first=False):
    """Добавляет документы раздела dkind в DOCX.

    Импортируется только текст документов Word (.docx). Исходные PDF (графика,
    приложения) в Word-документ НЕ выводятся: страницы PDF нельзя перенести
    редактируемым текстом, а их вставка картинками делает файл нередактируемым.
    Раньше для таких файлов печаталась строка с именем файла — в томе это
    выглядело как «пустой лист со ссылкой», поэтому она убрана.

    :param first_form: форма ГОСТ первой страницы ПЕРВОГО документа раздела
        (от неё зависит бюджет первой страницы при разбиении на листы).
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    docs = [d for d in project.documents_for(dkind)
            if d.path and d.path.lower().endswith((".docx", ".docm"))]
    if not docs:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run("Документы отсутствуют"), font, fm.content_row)
        return
    # Цели гиперссылок оглавления живут в документе, по которому оно построено.
    toc_path, toc_anchors = "", {}
    if dkind == DocKind.POYASNITELNAYA:
        _items, toc_path = project_toc_items(project)
        toc_anchors = anchor_map(_items)
    form = first_form
    for idx, d in enumerate(docs):
        if os.path.exists(d.path):
            _add_pz_doc(doc, d.path, fm, font, text_fonts=text_fonts,
                        first_form=form,
                        break_before_first=bool(break_before_first or idx),
                        toc_source=toc_path, toc_anchors=toc_anchors)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.page_break_before = bool(idx)
            _set_run(p.add_run(os.path.basename(d.path)), font, fm.content_row)
        # Первый лист раздела уже занят первым документом: следующие документы
        # начинаются с формы 6.
        form = FORM_TEXT_NEXT



def _add_appendix_section(doc, project, fm, font):
    """Добавляет раздел приложений: титульные листы «Приложение …».

    Сами приложения — исходные PDF; их страницы в Word не переносятся
    (см. `_add_section_docs`). Выводятся только генерируемые листы: номер,
    наименование и число листов приложения — по ним приложение вкладывается
    в том при печати.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    apps = project.appendices_needed()
    if not apps:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run("Приложения отсутствуют"), font, fm.content_row)
        return
    for idx, a in enumerate(apps):
        # Разделитель листов — только МЕЖДУ приложениями: после последнего
        # разрыв дал бы пустой лист в конце раздела.
        if idx:
            doc.add_page_break()
        title = f"Приложение {a.number}" if a.number else "Приложение"
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(p.add_run(title), font, fm.appendix_header, bold=True)
        if a.name:
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_run(p.add_run(a.name), font, fm.content_row)
        pages = int(getattr(a, "pages", 0) or 0)
        if pages:
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_run(p.add_run(f"Листов: {pages}"), font, fm.content_row)


# --------------------------------------------------------------------------
# Экспорт тома ПО РАЗДЕЛАМ (главам) — отдельный файл на раздел
# --------------------------------------------------------------------------

def _safe_filename(name: str) -> str:
    """Имя файла без символов, запрещённых в Windows."""
    cleaned = "".join("_" if ch in '<>:"/\\|?*' else ch for ch in name)
    return " ".join(cleaned.split()).strip(" .") or "раздел"


def chapter_sections(project: VolumeProject) -> list:
    """Разделы тома, дающие отдельный файл: [(kind, подпись), …].

    Порядок и состав — из `project.structure` (только включённые разделы).
    Пустые разделы не выводятся — так же, как в общем Word-документе:
    лист изменений без журнала и ПЗ без документов файлов не дают.
    """
    kinds = []
    for entry in project.structure or []:
        if not entry.get("needed", True):
            continue
        kind = entry.get("kind") or ""
        if kind not in SECTION_LABELS:
            continue
        kinds.append(kind)

    # Лист «Графическая часть» — единственное, что раздел даёт в Word.
    from .preview_html import _graphic_sheet_list_cfg
    try:
        graphic_list = bool(_graphic_sheet_list_cfg(project).enabled)
    except Exception:
        graphic_list = True

    # Пустые разделы не дают файла.
    kinds = [
        k for k in kinds
        if (k != SECTION_CHANGES or project.changes)
        and (k != SECTION_PZ or project.documents_for(DocKind.POYASNITELNAYA))
        and (k != SECTION_GRAPHIC
             or (graphic_list and project.documents_for(DocKind.GRAPHIC)))
        and (k != SECTION_APPENDIX or project.appendices_needed())
    ]
    return [(k, SECTION_LABELS.get(k, k)) for k in kinds]


def build_chapters_docx(project: VolumeProject, out_dir: str,
                        config: Config | None = None,
                        kinds: list | None = None) -> list:
    """Собирает том по разделам: отдельный DOCX на каждый раздел.

    :param out_dir: каталог для файлов (создаётся).
    :param kinds: ограничить набор разделов (ключи structure).
    :return: список собранных путей.
    """
    config = config or Config()
    os.makedirs(out_dir, exist_ok=True)
    sections = chapter_sections(project)
    if kinds is not None:
        wanted = list(kinds)
        sections = [(k, label) for k, label in sections if k in wanted]

    paths = []
    for kind, label in sections:
        designation = project.designation(kind)
        stem = f"{designation} {label}" if designation else label
        out_path = os.path.join(out_dir, _safe_filename(stem) + ".docx")
        build_section_docx(project, kind, out_path, config)
        paths.append(out_path)
    return paths

