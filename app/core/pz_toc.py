# -*- coding: utf-8 -*-
"""Автоматическое оглавление пояснительной записки по уровням текста из Word.

Модуль собирает заголовки ПЗ из исходного DOCX (стили Word + оформление) и
формирует оглавление, которое вставляется ПЕРВЫМ листом пояснительной записки:

  * HTML-предпросмотр (app.core.preview_html) — лист «ОГЛАВЛЕНИЕ» (листы) перед
    листами записки, с номерами листов записки;
  * Word / DOCX (app.core.docx_exporter) — абзацы оглавления с отточием и
    номером листа (отточие — правая табуляция с заполнителем DOTS, как в Word);
  * PDF (app.core.pdf_builder) — через DOCX (Word COM), а в запасном
    PyMuPDF-рендере — теми же строками оглавления.

НОРМАЛИЗАЦИЯ ТЕКСТА (требование проекта)
-----------------------------------------
При сопоставлении и нормализации текста НЕ учитываются пробелы и перенос
текста на другую страницу Word:

  * сравнение выполняется по «сжатому» виду (normalize_text): удаляются все
    пробельные символы, в том числе неразрывные (NBSP), узкие пробелы,
    табуляции, мягкие переносы и BOM;
  * разрыв страницы Word (w:pageBreakBefore, w:br type="page") на структуру
    заголовков не влияет: заголовок, разорванный между листами, остаётся одним
    заголовком, а абзацы-«переносы» (пустые строки, пробелы, символы разрыва)
    заголовками не считаются и в оглавление не попадают;
  * для вывода текст нормализуется до одного пробела между словами
    (clean_text) — перенос строки Word внутри абзаца не разрывает пункт.

Единый источник данных для предпросмотра, DOCX и PDF — build_toc_items().
"""
from __future__ import annotations

import html as html_mod
import os
import re
from dataclasses import dataclass

#: Заголовок оглавления.
TOC_TITLE = "ОГЛАВЛЕНИЕ"

#: Максимальный уровень заголовка, учитываемый в оглавлении.
MAX_LEVEL = 4

#: Отступ строки оглавления на один уровень (мм) — для DOCX и HTML.
LEVEL_INDENT_MM = 8.0

#: Позиция правой табуляции с отточием в DOCX (мм от левого края текста).
DOCX_TAB_POS_MM = 175.0

#: Пунктов в миллиметре (для оценки высоты строки оглавления).
PT_PER_MM = 72.0 / 25.4

# --- Нормализация текста ----------------------------------------------------

# Пробельные символы (включая NBSP, узкие пробелы, табуляции, BOM, ZWSP).
_SPACE_RE = re.compile(r"[ \t\u00a0\u1680\u2000-\u200a\u202f\u205f\u3000\u200b\ufeff]+")
# Символы переноса строки/страницы (в DOCX никогда не входят в текст абзаца, но
# могут встретиться в тексте, скопированном из PDF или из другого источника).
_BREAK_RE = re.compile(r"[\r\n\v\f\u2028\u2029]+")
# Мягкий перенос и нулевой пробел — «невидимые» символы, не влияющие на слова.
_INVISIBLE_RE = re.compile(r"[\u00ad\u200b\ufeff]")


def normalize_text(text) -> str:
    """Сжатый вид текста: без пробелов, табуляций и переносов.

    Используется для СРАВНЕНИЯ текстов (поиск заголовков, дедупликация,
    проверка «та же строка, но с другим переносом страницы»). Все пробельные
    символы и мягкие переносы/нулевые пробелы удаляются полностью — поэтому
    «Раздел 1. Общие положения» и «Раздел\u00a01.\nОбщие положения» равны.
    """
    if not text:
        return ""
    s = _INVISIBLE_RE.sub("", str(text))
    s = _BREAK_RE.sub("", s)
    return _SPACE_RE.sub("", s)


def clean_text(text) -> str:
    """Текст для вывода: пробелы/переносы приведены к одному пробелу.

    Пробелы и перенос строки Word (w:br) внутри абзаца не разрывают пункт
    оглавления: «Содержание   раздела» -> «Содержание раздела». Пустой или
    состоящий только из пробелов/переносов абзац даёт пустую строку и в
    оглавление не попадает.
    """
    if not text:
        return ""
    s = _INVISIBLE_RE.sub("", str(text))
    s = _BREAK_RE.sub(" ", s)
    s = s.replace("\u00a0", " ")
    return " ".join(s.split())


def same_text(left, right) -> bool:
    """Равны ли тексты без учёта пробелов и переносов."""
    return normalize_text(left) == normalize_text(right)


# --- Распознавание заголовков -----------------------------------------------

# Стиль абзаца — заголовок (как его называет Word и ГОСТ-шаблоны).
_STYLE_HEADING_RE = re.compile(
    r"(heading|заголов|название\s*раздела|раздел\b|section\s*heading|title)", re.I)
# Номер уровня в имени стиля («Heading 2», «Заголовок 3», «H2»).
_STYLE_LEVEL_RE = re.compile(r"(\d+)")
# Номер раздела в начале текста заголовка: «2.1.3 Наименование».
_LEADING_NUM_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+")
# Строка-заголовок без стиля: короткая, без точки в конце, без «пустого» текста.
_SENTENCE_END_RE = re.compile(r"[.:;,!?]$")


@dataclass
class Heading:
    """Заголовок исходного DOCX."""
    index: int = -1          # индекс абзаца в document.paragraphs
    text: str = ""           # текст абзаца (нормализованный для вывода)
    level: int = 1           # уровень 1..MAX_LEVEL
    sheet: int = 0           # номер листа ПЗ (по пагинации предпросмотра)


@dataclass
class TocItem:
    """Строка автоматического оглавления ПЗ.

    kind:
        "title"   — заголовок «ОГЛАВЛЕНИЕ» (шапка листа);
        "heading" — заголовок раздела (num — номер раздела, level — уровень).
    """
    kind: str = "heading"
    text: str = ""
    num: str = ""
    level: int = 1
    page: int = 0
    para_index: int = -1

    @property
    def is_title(self) -> bool:
        return self.kind == "title"

    @property
    def label(self) -> str:
        """Строка оглавления «номер + наименование»."""
        return (self.num + " " + self.text).strip() if self.num else self.text


def _paragraph_level(para) -> int:
    """Уровень заголовка (1..MAX_LEVEL) — стиль Word, затем эвристика.

    Стиль «Heading N» / «Заголовок N» задаёт уровень точно. Если стиль не
    проставлен, уровень определяется по нумерации текста («2.1 Название» -> 2),
    иначе заголовок считается первого уровня.
    """
    try:
        level = int(getattr(para, "level", 0) or 0)
    except (TypeError, ValueError):
        level = 0
    if level > 0:
        return max(1, min(MAX_LEVEL, level))
    style = (getattr(para, "style_name", "") or "").strip().lower()
    m = _STYLE_LEVEL_RE.search(style or "")
    if m and _STYLE_HEADING_RE.search(style):
        try:
            return max(1, min(MAX_LEVEL, int(m.group(1))))
        except ValueError:
            pass
    m = _LEADING_NUM_RE.match(clean_text(getattr(para, "text", "")))
    if m:
        return max(1, min(MAX_LEVEL, m.group(1).count(".") + 1))
    return 1


def is_heading_paragraph(para) -> bool:
    """Абзац является заголовком для оглавления?

    Порядок проверки:
      1) пустой абзац (только пробелы/переносы/разрывы страницы) — не заголовок;
      2) стиль Word «Heading N» / «Заголовок N» — заголовок;
      3) пункт списка (w:numPr) — не заголовок (нумерованные перечисления ПЗ
         в оглавление не входят);
      4) заголовок без стиля: короткий абзац (< 200 знаков), без завершающей
         точки и выделенный (полужирный) или центрированный.
    """
    text = clean_text(getattr(para, "text", ""))
    if not text:
        return False                      # пустой абзац / только пробелы / перенос
    if len(text) > 200:
        return False
    if not same_text(text, text):         # защита от непечатных «мусорных» строк
        return False
    fmt = getattr(para, "fmt", None)
    has_list = bool(fmt is not None and getattr(fmt, "has_list", False))
    style = (getattr(para, "style_name", "") or "").strip().lower()
    if not has_list and _STYLE_HEADING_RE.search(style or ""):
        return True
    try:
        level = int(getattr(para, "level", 0) or 0)
    except (TypeError, ValueError):
        level = 0
    if level > 0:
        return True
    if getattr(para, "is_heading", False):
        return True
    if has_list:
        return False
    if _SENTENCE_END_RE.search(text):
        return False
    bold = any(getattr(r, "bold", False) for r in (getattr(para, "runs", None) or []))
    align = (getattr(para, "alignment", "") or "").upper()
    if bold or align == "CENTER":
        return True
    return False


def collect_headings(content, sheets: dict | None = None) -> list:
    """Заголовки DOCX в порядке документа с номерами листов ПЗ.

    :param content: DocxContent (app.core.docx_reader.read_docx).
    :param sheets: {индекс абзаца: номер листа} — из sheet_numbers(); если не
        передан, номера листов не заполняются (0).
    """
    sheets = sheets or {}
    out = []
    for index, para in enumerate(getattr(content, "paragraphs", None) or []):
        if not is_heading_paragraph(para):
            continue
        out.append(Heading(index=index, text=clean_text(para.text),
                           level=_paragraph_level(para),
                           sheet=int(sheets.get(index, 0) or 0)))
    return out


# --- Пагинация: номера листов после разрывов Word ---------------------------

def sheet_numbers(content, path: str = "") -> dict:
    """{индекс абзаца: номер листа ПЗ} по единой пагинации предпросмотра.

    Используется та же разбивка, что и в предпросмотре/сборке
    (preview_html._paginate_docx), поэтому номера листов в оглавлении совпадают
    с листами записки. Разрыв страницы, заданный в Word, учтён алгоритмом
    пагинации: абзац с «с новой страницы» начинает следующий лист.
    """
    try:
        from .preview_html import _paginate_docx
        pages = _paginate_docx(content, path or "")
    except Exception:
        return {}
    numbers: dict = {}
    for sheet_no, page in enumerate(pages or [], start=1):
        for item in page or []:
            if not item or item[0] != "line":
                continue
            try:
                index = int((item[1] or {}).get("para_index", -1))
            except (TypeError, ValueError):
                continue
            if index >= 0:
                numbers.setdefault(index, sheet_no)
    return numbers


def paginate_items(items) -> list:
    """Разбивает строки оглавления по листам записки.

    Высота строки берётся из той же модели, что и пагинация ПЗ (кегль 11pt,
    line-height 1.4), а доступная высота — как у листов ПЗ: первый лист несёт
    штамп формы 5 (40 мм), последующие — формы 6 (15 мм). Поэтому число листов
    оглавления совпадает с фактическим размещением в предпросмотре и сборке.
    """
    from .preview_html import _page_text_height, FORM_HEIGHT
    # Высота строки оглавления с запасом (кегль 12pt, line-height 1.4): оценка
    # должна быть консервативной, чтобы строка не уехала под штамп листа.
    per_line_mm = 12.0 * 1.4 / PT_PER_MM     # ~5.9 мм на строку
    title_mm = 18.0                          # заголовок «ОГЛАВЛЕНИЕ» + отступ
    first_budget = _page_text_height(FORM_HEIGHT[2])     # форма 5 (40 мм)
    next_budget = _page_text_height(FORM_HEIGHT["2a"])   # форма 6 (15 мм)
    pages, cur, used, budget = [], [], 0.0, first_budget
    for item in items or []:
        height = per_line_mm + (title_mm if getattr(item, "is_title", False) else 0.0)
        if cur and used + height > budget:
            pages.append(cur)
            cur, used, budget = [], 0.0, next_budget
        cur.append(item)
        used += height
    if cur:
        pages.append(cur)
    return pages or [[]]


def toc_pages(content, path: str = "", sheets: dict | None = None) -> list:
    """Строки оглавления ПЗ (build_toc_items), разбитые по листам записки."""
    return paginate_items(build_toc_items(content, path, sheets))


def item_dicts(items) -> list:
    """Строки оглавления как простые словари (для снимка предпросмотра)."""
    return [{"kind": it.kind, "text": it.text, "num": it.num,
             "level": int(it.level), "page": int(it.page)} for it in (items or [])]


def items_from_dicts(raw) -> list:
    """Восстанавливает строки оглавления из словарей снимка предпросмотра."""
    out = []
    for item in raw or []:
        if isinstance(item, dict):
            out.append(TocItem(kind=item.get("kind", "heading"),
                               text=item.get("text", ""),
                               num=item.get("num", ""),
                               level=int(item.get("level", 1) or 1),
                               page=int(item.get("page", 0) or 0)))
        elif isinstance(item, (list, tuple)) and len(item) >= 2:
            out.append(TocItem(kind=item[0], text=item[1],
                               num=item[2] if len(item) > 2 else "",
                               level=int(item[3] if len(item) > 3 else 1),
                               page=int(item[4] if len(item) > 4 else 0)))
        elif item:
            out.append(TocItem(kind="heading", text=str(item)))
    return out


def toc_sheet_count(content, path: str = "") -> int:
    """Сколько листов займёт оглавление ПЗ (минимум 1)."""
    return max(1, len(toc_pages(content, path)))


# --- Сборка строк оглавления -------------------------------------------------

def build_toc_items(content, path: str = "", sheets: dict | None = None) -> list:
    """Строки автоматического оглавления ПЗ (единый источник для всех выводов).

    Возвращает список TocItem: первая строка — заголовок «ОГЛАВЛЕНИЕ», далее по
    одной строке на каждый заголовок записки (номер раздела, наименование,
    номер листа) в порядке документа и по уровням текста из Word.

    Номера разделов НЕ придумываются: номер берётся из самого заголовка
    («2.1 Наименование» -> номер 2.1, «Раздел 2. Наименование» -> номера нет,
    он уже есть в тексте). Иерархия уровней передаётся отступом строки.
    Текст заголовков сравнивается и выводится без учёта пробелов и переносов
    текста на другую страницу Word.

    :param content: DocxContent исходной записки.
    :param path: путь к DOCX (нужен для фактических ширин колонок таблиц —
        от них зависит пагинация, а значит и номера листов в оглавлении).
    :param sheets: заранее посчитанные номера листов (sheet_numbers).
    """
    if sheets is None:
        sheets = sheet_numbers(content, path)
    headings = collect_headings(content, sheets)
    items = [TocItem(kind="title", text=TOC_TITLE)]
    for heading in headings:
        level = max(1, min(MAX_LEVEL, int(heading.level or 1)))
        num = ""
        text = heading.text
        match = _LEADING_NUM_RE.match(text)
        if match:
            # Номер уже есть в тексте: выносим его отдельно («2.1» + название),
            # чтобы он не дублировался при выводе строки оглавления. Заголовки
            # вида «Раздел 2. Наименование» номера не получают — он уже в тексте.
            num = match.group(1)
            text = text[match.end():].strip() or text
        items.append(TocItem(kind="heading", text=text, num=num,
                             level=level, page=heading.sheet, para_index=heading.index))
    return items


def pz_sheet_offset(project) -> int:
    """Сдвиг номеров листов записки для печати в оглавлении (номер − 1).

    В оглавлении номера считаются от начала ЗАПИСКИ (лист 1 — первый лист после
    оглавления), а в томе записка начинается позже (после содержания). Сдвиг
    равен «номер первого листа раздела ПЗ в томе минус 1», поэтому напечатанный
    номер совпадает с квадратом сквозной нумерации на листе.

    Сдвиг = (номер первого листа раздела ПЗ − 1) + число листов оглавления:
    листы оглавления занимают начало раздела, поэтому лист записки с локальным
    номером 1 идёт сразу после них.

    Единый источник номера — app.core.toc.build_layout (учитывает структуру
    тома и листы самого оглавления).
    """
    try:
        from .toc import build_layout
        for item in build_layout(project).items:
            if item.kind == "pz" and item.count > 0:
                base = max(0, int(item.first_sheet) - 1)
                return base + project_toc_sheets(project)
    except Exception:
        pass
    return 0


def toc_page_offset(project, standalone: bool = False) -> int:
    """Сдвиг номеров листов в оглавлении для печати строк.

    :param standalone: True — раздел собирается отдельно (нумерация листов
        начинается с 1), тогда сдвиг равен числу листов самого оглавления;
        False — том целиком, сдвиг считается по структуре тома
        (см. pz_sheet_offset).
    """
    if standalone:
        return max(0, int(project_toc_sheets(project) or 0))
    return pz_sheet_offset(project)


def project_toc_enabled(project) -> bool:
    """Включено ли автоматическое оглавление ПЗ в конфигурации проекта."""
    cfg = _project_config(project)
    if cfg is None:
        return True
    return bool(getattr(cfg, "toc_enabled", True))


def _project_config(project):
    """Конфигурация оформления проекта ('' -> значения по умолчанию)."""
    cfg = getattr(project, "config", None)
    if cfg is not None:
        return cfg
    try:
        from ..config import load_config
        return load_config()
    except Exception:
        return None


def project_toc_sheets(project) -> int:
    """Число листов автоматического оглавления ПЗ для проекта (0 — не нужно).

    Оглавление строится по ПЕРВОМУ документу раздела ПЗ (как и в предпросмотре,
    и при сборке). Если документ недоступен — оглавление не добавляется.
    """
    if project is None or not project_toc_enabled(project):
        return 0
    try:
        from ..models import DocKind, SECTION_PZ
        if not project.structure_needed(SECTION_PZ):
            return 0            # раздел ПЗ исключён из сборки — оглавление не нужно
        docs = project.documents_for(DocKind.POYASNITELNAYA)
    except Exception:
        return 0
    if not docs:
        return 0
    first = docs[0]
    path = getattr(first, "path", "") or ""
    if not path or not os.path.exists(path):
        return 0
    try:
        from .docx_reader import read_docx
        return toc_sheet_count(read_docx(path), path)
    except Exception:
        return 0


def first_pz_path(project) -> str:
    """Путь к первому документу раздела ПЗ ('' — если документа нет)."""
    try:
        from ..models import DocKind
        docs = project.documents_for(DocKind.POYASNITELNAYA)
    except Exception:
        return ""
    for d in docs:
        path = getattr(d, "path", "") or ""
        if path and os.path.exists(path):
            return path
    return ""


def project_toc_items(project) -> tuple:
    """Строки автоматического оглавления ПЗ проекта и документ-источник.

    Единый источник для предпросмотра, Word и PDF: оглавление строится по
    ПЕРВОМУ документу раздела ПЗ (той же пагинацией, что и листы записки).

    :return: (список TocItem, путь к DOCX). Пустой список — оглавление не
        выводится (нет документа ПЗ, раздел исключён из сборки, оглавление
        отключено в конфигурации или документ не читается).
    """
    if not project_toc_sheets(project):
        return [], ""
    path = first_pz_path(project)
    if not path:
        return [], ""
    try:
        from .docx_reader import read_docx
        items = build_toc_items(read_docx(path), path)
    except Exception:
        return [], ""
    if len(items) <= 1:
        # Заголовков уровней текста в записке нет — оглавление не выводится
        # (лист с одной шапкой «ОГЛАВЛЕНИЕ» не создаётся).
        return [], ""
    return items, path


# --- Вывод: HTML, DOCX, PDF --------------------------------------------------

def toc_html(items, *, title: str = TOC_TITLE, body_size: int = 11,
             title_size: int = 14, offset: int = 0) -> str:
    """HTML-разметка листа оглавления (для предпросмотра и сборки в HTML).

    Отточие — точечная нижняя граница «резинки» между текстом и номером листа,
    как заполнитель табуляции в Word. Размеры — в пунктах карты кеглей
    (app.config.FontMap).

    :param offset: сдвиг номеров листов: номера в строках оглавления считаются
        от начала ЗАПИСКИ (лист 1 — первый лист после оглавления), а в томе
        записка начинается не с первого листа. offset = номер первого листа
        записки в текущей сборке минус 1, поэтому напечатанный номер совпадает
        с квадратом сквозной нумерации на листе.
    """
    def esc(value) -> str:
        return html_mod.escape(str(value if value is not None else ""))

    shift = max(0, int(offset or 0))
    lines = []
    for item in items:
        if item.is_title:
            lines.append(
                f'<h1 style="font-size:{title_size}pt;text-align:center;'
                f'font-weight:bold;margin:0 0 8mm 0;">{esc(title)}</h1>')
            continue
        indent = LEVEL_INDENT_MM * max(0, int(item.level) - 1)
        size = body_size
        weight = "bold"
        num_html = (f'<span style="flex:0 0 auto;white-space:nowrap;font-weight:bold;'
                    f'margin-right:1.5mm;">{esc(item.num)}</span>'
                    if item.num else "")
        page = int(item.page or 0)
        page_html = (f'<span style="flex:0 0 auto;white-space:nowrap;margin-left:1.5mm;">'
                     f'{esc(page + shift)}</span>' if page else
                     '<span style="flex:0 0 auto;margin-left:1.5mm;">&nbsp;</span>')
        # Текст заголовка может быть длинным (в т.ч. без пробелов), поэтому
        # он СЖИМАЕТСЯ (flex:0 1 auto + min-width:0) и переносится на строки.
        # Иначе длинное наименование распирало flex-строку, и номер листа с
        # отточием уходили за правую рамку листа (правый край был не виден).
        lines.append(
            f'<div style="display:flex;align-items:baseline;font-size:{size}pt;'
            f'font-weight:{weight};margin-left:{indent:g}mm;padding:0.4mm 0;">'
            f'{num_html}'
            f'<span style="flex:0 1 auto;min-width:0;overflow-wrap:break-word;'
            f'word-break:break-word;">{esc(item.text)}</span>'
            f'<span style="flex:1 1 auto;border-bottom:0.3mm dotted #000;'
            f'margin:0 0.5mm 0.9mm 0.5mm;min-width:6mm;"></span>'
            f'{page_html}</div>')
    return '<div class="pz-toc">' + "".join(lines) + '</div>'


def toc_pdf_lines(items, *, title: str = TOC_TITLE, width: int = 95,
                  offset: int = 0) -> list:
    """Строки оглавления для прямого рендера PyMuPDF (запасной путь без Word).

    Отточие воспроизводится точками: «2.1 Наименование .......... 7».
    offset — сдвиг номеров листов (см. toc_html).
    """
    shift = max(0, int(offset or 0))
    out = [title]
    for item in items:
        if item.is_title:
            continue
        indent = " " * (2 * max(0, int(item.level) - 1))
        left = f"{indent}{item.label}"
        page = str(int(item.page) + shift) if item.page else ""
        space = max(1, width - len(left) - len(page) - 2)
        out.append(f"{left} {'.' * space} {page}".rstrip())
    return out
