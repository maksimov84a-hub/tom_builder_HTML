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

#: Максимальный уровень заголовка, попадающего в оглавление (1..3).
#: Заголовки 4-го уровня и глубже в оглавление НЕ входят.
MAX_LEVEL = 3

#: Предел уровня заголовка при разборе DOCX (выше MAX_LEVEL — отбрасывается).
LEVEL_LIMIT = 9

#: Отступ строки оглавления на один уровень (мм) — для DOCX и HTML.
#: Считается по метрикам шрифта: ровно ОДИН символ текста ПЗ. Строка 2-го
#: уровня сдвинута на символ от 1-го, строка 3-го — ещё на символ от 2-го.
#: Значение — запасное, если метрики шрифта недоступны (см. level_indent_mm).
LEVEL_INDENT_MM = 2.2

#: Символ, по ширине которого считается отступ уровня.
LEVEL_INDENT_CHAR = "0"

#: Внутренний отступ текста от линий рамки (мм) — как `padding` .sheet-content.
SHEET_PAD_MM = 6.0

#: Правая граница текстовой области листа (мм от левого поля): 185 мм рамки
#: минус 6 мм внутреннего отступа. По ней выравниваются номера листов.
DOCX_TAB_POS_MM = 179.0

#: Префикс имён закладок заголовков (гиперссылки оглавления → заголовок).
ANCHOR_PREFIX = "MIMO_H"


def heading_anchor(para_index) -> str:
    """Имя закладки (HTML-якоря) заголовка по индексу его абзаца."""
    try:
        index = int(para_index)
    except (TypeError, ValueError):
        return ""
    return f"{ANCHOR_PREFIX}{index}" if index >= 0 else ""


def anchor_map(items) -> dict:
    """{индекс абзаца: имя закладки} для строк оглавления."""
    return {int(item.para_index): item.anchor
            for item in (items or []) if item.anchor}


def toc_text_style() -> tuple:
    """Шрифт и кегль строк оглавления — как у обычного текста ПЗ.

    Требование к записке: оглавление набирается тем же шрифтом и кеглем, что и
    обычный текст (config.TextFonts, группа «body»), а не кеглем карты кеглей.
    """
    try:
        from ..config import load_config
        fonts = load_config().text_fonts
        return (fonts.font("body") or "", float(fonts.size("body")))
    except Exception:
        return "", 12.0

#: Пунктов в миллиметре (для оценки высоты строки оглавления).
PT_PER_MM = 72.0 / 25.4


def level_indent_mm(level) -> float:
    """Отступ строки оглавления для уровня: один СИМВОЛ текста ПЗ на уровень.

    Требование к записке: строка 2-го уровня сдвинута вправо на один символ
    относительно 1-го, строка 3-го — ещё на один символ относительно 2-го.
    Ширина символа берётся из файла шрифта оглавления при его кегле, поэтому
    отступ масштабируется вместе с настройкой шрифта текста ПЗ; если метрик
    нет — используется LEVEL_INDENT_MM.
    """
    try:
        step = float(level) if level else 1.0
    except (TypeError, ValueError):
        step = 1.0
    count = max(0, int(step) - 1)
    if not count:
        return 0.0
    try:
        from .textfit import font_path, text_width_mm
        family, size = toc_text_style()
        char_mm = text_width_mm(LEVEL_INDENT_CHAR, size, font_path(family or None))
        if not char_mm or char_mm <= 0:
            char_mm = LEVEL_INDENT_MM
    except Exception:
        char_mm = LEVEL_INDENT_MM
    return char_mm * count

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

# Номер уровня в имени стиля («Heading 2», «Заголовок 3», «H2»).
_STYLE_LEVEL_RE = re.compile(r"(\d+)")
# Номер раздела в начале текста заголовка: «2.1.3 Наименование».
_LEADING_NUM_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+")


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

    @property
    def anchor(self) -> str:
        """Имя закладки заголовка (цель гиперссылки строки оглавления)."""
        return heading_anchor(self.para_index)


def _paragraph_level(para) -> int:
    """Уровень заголовка (1..LEVEL_LIMIT) — только из стиля Word.

    Стиль «Heading N» / «Заголовок N» задаёт уровень своим номером. Стиль-
    заголовок без номера («Заголовок») считается первым уровнем.

    Уровень НЕ ограничивается сверху: решение о том, попадает ли заголовок в
    оглавление, принимает collect_headings (уровни 1..MAX_LEVEL).
    """
    try:
        level = int(getattr(para, "level", 0) or 0)
    except (TypeError, ValueError):
        level = 0
    if level > 0:
        return max(1, min(LEVEL_LIMIT, level))
    style = (getattr(para, "style_name", "") or "").strip().lower()
    m = _STYLE_LEVEL_RE.search(style or "")
    if m and _STYLE_HEADING_RE.search(style or ""):
        try:
            return max(1, min(LEVEL_LIMIT, int(m.group(1))))
        except ValueError:
            pass
    return 1


def is_heading_paragraph(para) -> bool:
    """Абзац является заголовком для оглавления?

    Единственный признак — **стиль Word**: «Heading N» / «Заголовок N» (включая
    стили, созданные на основе них) или имя стиля, содержащее «heading».
    Полужирный шрифт, центрирование, длина строки и завершающая точка НЕ
    учитываются — это было эвристикой, и она втягивала в оглавление формулы и
    подписи таблиц.

    Отбрасывается только абзац без текста: выводить в оглавление нечего.
    """
    if not getattr(para, "is_heading", False):
        return False
    return bool(clean_text(getattr(para, "text", "")))


def heading_emphasis(para) -> tuple:
    """Начертание абзаца в выводе ПЗ: (полужирный, курсив).

    Заголовок выводится **полужирным** — выделение, как было в записке; курсив
    берётся из исходника (у заголовков он сохраняется). Обычный текст записки
    печатается единым — без полужирного и курсива.

    Единое правило для предпросмотра, DOCX и «нормализованной» копии, по которой
    Word размечает листы: начертание меняет ширину строк, и копия должна мерить
    те же строки, что вывод.
    """
    if not is_heading_paragraph(para):
        return False, False
    runs = getattr(para, "runs", None) or []
    return True, any(bool(getattr(r, "italic", False)) for r in runs)


def collect_headings(content, sheets: dict | None = None) -> list:
    """Заголовки DOCX в порядке документа с номерами листов ПЗ.

    Отбор — только по стилю Word (см. is_heading_paragraph): каждый абзац со
    стилем-заголовком даёт отдельный пункт, других условий нет.

    В оглавление попадают только заголовки уровней 1..MAX_LEVEL (1–3):
    заголовки 4-го уровня и глубже отбрасываются.

    :param content: DocxContent (app.core.docx_reader.read_docx).
    :param sheets: {индекс абзаца: номер листа} — из sheet_numbers(); если не
        передан, номера листов не заполняются (0).
    """
    sheets = sheets or {}
    out = []
    for index, para in enumerate(getattr(content, "paragraphs", None) or []):
        if not is_heading_paragraph(para):
            continue
        level = _paragraph_level(para)
        if level > MAX_LEVEL:
            continue
        out.append(Heading(index=index, text=clean_text(para.text),
                           level=level,
                           sheet=int(sheets.get(index, 0) or 0)))
    return out


# --- Пагинация: номера листов после разрывов Word ---------------------------

def sheet_numbers(content, path: str = "") -> dict:
    """{индекс абзаца: номер листа ПЗ} по единой пагинации предпросмотра.

    Используется та же разбивка, что и в предпросмотре/сборке
    (preview_html._paginate_docx), поэтому номера листов в оглавлении совпадают
    с листами записки. Разрыв страницы, заданный в Word, учтён алгоритмом
    пагинации: абзац с «с новой страницы» начинает следующий лист.

    Границы листов определяет РАЗМЕТКА WORD (`word_pagination`), а не оценка
    по высоте строк: оценочная раскладка давала 2 листа вместо 13, и половина
    заголовков получала один и тот же номер. Кэш раскладки читается с диска;
    если его нет — работает оценка (предпросмотр пересчитает через Word).
    """
    try:
        from .preview_html import _paginate_docx
        # Лист оглавления несёт форму 5 и стоит ПЕРЕД содержимым первого
        # документа, поэтому его страницы (в т.ч. страницы заголовков)
        # рисуются и пагинируются с формы 6 (штамп 15 мм) — бюджет первого
        # листа должен совпадать с фактическим штампом, иначе разбивка и
        # номера листов в оглавлении разойдутся с предпросмотром.
        word_layout = None
        if path:
            try:
                from .word_pagination import cached_docx_pages
                word_layout = cached_docx_pages(path, first_form="2a")
            except Exception:
                word_layout = None
        pages = _paginate_docx(content, path or "", first_form="2a",
                               word_layout=word_layout)
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


def _toc_row_height_mm(item, size_pt: float, line_mm: float, path: str = "") -> float:
    """Высота строки оглавления с учётом ПЕРЕНОСА длинного наименования.

    Строка «номер — наименование — отточие — номер листа» переносится, если
    наименование не влезает в одну строку (в ПЗ наименования бывают по 200
    знаков). Раньше каждая строка считалась однострочной, и пагинация
    оглавления давала 2 листа вместо 3 — номера листов в оглавлении сдвигались.
    """
    from .textfit import font_path, text_width_mm, visual_lines
    if not path:
        path = font_path()
    indent = level_indent_mm(getattr(item, "level", 1))
    # Справа от наименования стоит номер листа (до трёх знаков) и отточие.
    reserve = text_width_mm("000", size_pt, path) + 2.0
    avail = DOCX_TAB_POS_MM - SHEET_PAD_MM - indent - reserve
    label = getattr(item, "label", None) or getattr(item, "text", "") or ""
    lines = max(1, visual_lines(label, avail, size_pt, path))
    return lines * line_mm + 0.8


def paginate_items(items, size_pt: float | None = None) -> list:
    """Разбивает строки оглавления по листам записки.

    Высота строки берётся из той же модели, что и пагинация ПЗ, и совпадает с
    тем, что рисует HTML: строка `toc_html` — это кегль обычного текста ПЗ
    (`toc_text_style`, по умолчанию 12 pt) с высотой строки
    `preview_html._line_height`, плюс внутренние отступы `padding:0.4mm 0`
    сверху и снизу. Длинное наименование занимает несколько строк (см.
    `_toc_row_height_mm`). Доступная высота — как у листов ПЗ: первый лист несёт
    штамп формы 5 (40 мм), последующие — формы 6.

    :param size_pt: кегль строк оглавления; None — как у текста ПЗ (кегль
        строк оглавления равен кеглю обычного текста записки).
    """
    from .preview_html import _page_text_height, _line_height, FORM_HEIGHT
    if size_pt is None:
        size_pt = toc_text_style()[1]
    # Высота одной строки кегля; отступы строки (0.4 мм сверху и снизу)
    # добавляются к строке-строке один раз — в _toc_row_height_mm.
    line_mm = _line_height(size_pt)
    title_mm = 18.0                          # заголовок «ОГЛАВЛЕНИЕ» + отступ
    first_budget = _page_text_height(FORM_HEIGHT[2])     # форма 5 (40 мм)
    next_budget = _page_text_height(FORM_HEIGHT["2a"])   # форма 6 (15 мм)
    pages, cur, used, budget = [], [], 0.0, first_budget
    for item in items or []:
        height = _toc_row_height_mm(item, size_pt, line_mm)
        if getattr(item, "is_title", False):
            height += title_mm
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
    """Строки оглавления как простые словари (для снимка предпросмотра).

    para_index нужен предпросмотру: по нему строка оглавления связывается с
    абзацем-целью гиперссылки (якорь `MIMO_H{индекс}`).
    """
    return [{"kind": it.kind, "text": it.text, "num": it.num,
             "level": int(it.level), "page": int(it.page),
             "para_index": int(it.para_index)} for it in (items or [])]


def items_from_dicts(raw) -> list:
    """Восстанавливает строки оглавления из словарей снимка предпросмотра."""
    out = []
    for item in raw or []:
        if isinstance(item, dict):
            out.append(TocItem(kind=item.get("kind", "heading"),
                               text=item.get("text", ""),
                               num=item.get("num", ""),
                               level=int(item.get("level", 1) or 1),
                               page=int(item.get("page", 0) or 0),
                               para_index=int(item.get("para_index", -1) or -1)))
        elif isinstance(item, (list, tuple)) and len(item) >= 2:
            out.append(TocItem(kind=item[0], text=item[1],
                               num=item[2] if len(item) > 2 else "",
                               level=int(item[3] if len(item) > 3 else 1),
                               page=int(item[4] if len(item) > 4 else 0)))
        elif item:
            out.append(TocItem(kind="heading", text=str(item)))
    return out


def toc_sheet_count(content, path: str = "") -> int:
    """Сколько листов займёт оглавление ПЗ (0 — оглавление не выводится).

    Как и в предпросмотре/сборке (project_toc_items): если в записке нет
    заголовков по уровням текста, лист-«пустышка» с одной шапкой
    «ОГЛАВЛЕНИЕ» не создаётся — тогда и счётчик должен быть 0, иначе
    графа «Листов» в штампе расходится с числом реально показанных листов.
    """
    items = build_toc_items(content, path)
    if len(items) <= 1:
        return 0
    return max(1, len(paginate_items(items)))


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


def toc_page_offset(project) -> int:
    """Сдвиг номеров листов в строках оглавления.

    Номера в оглавлении должны совпадать с номером, напечатанным внизу листа
    записки (графа «Лист» штампа), то есть считаться **внутри раздела ПЗ**, без
    сквозной нумерации тома. Штамп ведёт нумерацию по разделу, и листы
    оглавления — его первые листы: первый лист самой записки получает номер
    «число листов оглавления + 1».

    Поэтому сдвиг всегда равен числу листов оглавления и не зависит от того,
    каким по счёту идёт раздел ПЗ в томе (раньше для тома прибавлялся номер
    первого листа раздела, и оглавление показывало сквозные номера).
    """
    return max(0, int(project_toc_sheets(project) or 0))


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

def toc_html(items, *, title: str = TOC_TITLE, body_size: float | None = None,
             title_size: int = 14, offset: int = 0,
             font_family: str = "") -> str:
    """HTML-разметка листа оглавления (для предпросмотра и сборки в HTML).

    Отточие — точечная нижняя граница «резинки» между текстом и номером листа,
    как заполнитель табуляции в Word.

    Оформление строк — как у обычного текста записки: тот же шрифт и кегль
    (`toc_text_style`), без полужирного; номер листа прижат к правому краю
    текстовой области. Наименование заголовка — ссылка на сам заголовок
    (`#MIMO_H{индекс абзаца}`), поэтому в предпросмотре и собранном HTML
    работает переход к пункту.

    :param body_size: кегль строк, pt; None — кегль обычного текста ПЗ.
    :param font_family: имя шрифта ПЗ; пусто — из настроек текста ПЗ.
    :param offset: сдвиг номеров листов: номера в строках оглавления считаются
        от начала ЗАПИСКИ (лист 1 — первый лист после оглавления), а в томе
        записка начинается не с первого листа. offset = номер первого листа
        записки в текущей сборке минус 1, поэтому напечатанный номер совпадает
        с квадратом сквозной нумерации на листе.
    """
    def esc(value) -> str:
        return html_mod.escape(str(value if value is not None else ""))

    style_font, style_size = toc_text_style()
    if body_size is None:
        body_size = style_size
    family = (font_family or style_font or "").strip()
    family_css = f"font-family:'{family}',Arial,sans-serif;" if family else ""

    shift = max(0, int(offset or 0))
    lines = []
    for item in items:
        if item.is_title:
            lines.append(
                f'<h1 style="font-size:{title_size}pt;text-align:center;'
                f'{family_css}'
                f'font-weight:bold;margin:0 0 8mm 0;">{esc(title)}</h1>')
            continue
        indent = level_indent_mm(item.level)
        num_html = (f'<span style="flex:0 0 auto;white-space:nowrap;'
                    f'margin-right:1.5mm;">{esc(item.num)}</span>'
                    if item.num else "")
        page = int(item.page or 0)
        page_html = (f'<span style="flex:0 0 auto;white-space:nowrap;margin-left:1.5mm;">'
                     f'{esc(page + shift)}</span>' if page else
                     '<span style="flex:0 0 auto;margin-left:1.5mm;">&nbsp;</span>')
        text_html = esc(item.text)
        if item.anchor:
            text_html = (f'<a href="#{esc(item.anchor)}" '
                         f'style="color:inherit;text-decoration:none;">'
                         f'{text_html}</a>')
        # Текст заголовка может быть длинным (в т.ч. без пробелов), поэтому
        # он СЖИМАЕТСЯ (flex:0 1 auto + min-width:0) и переносится на строки.
        # Иначе длинное наименование распирало flex-строку, и номер листа с
        # отточием уходили за правую рамку листа (правый край был не виден).
        lines.append(
            f'<div style="display:flex;align-items:baseline;'
            f'font-size:{body_size:g}pt;{family_css}font-weight:normal;'
            f'margin-left:{indent:g}mm;padding:0.4mm 0;">'
            f'{num_html}'
            f'<span style="flex:0 1 auto;min-width:0;overflow-wrap:break-word;'
            f'word-break:break-word;">{text_html}</span>'
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
        # Отступ уровня — один символ текста на уровень (как в HTML и DOCX).
        indent = " " * max(0, int(item.level) - 1)
        left = f"{indent}{item.label}"
        page = str(int(item.page) + shift) if item.page else ""
        space = max(1, width - len(left) - len(page) - 2)
        out.append(f"{left} {'.' * space} {page}".rstrip())
    return out
