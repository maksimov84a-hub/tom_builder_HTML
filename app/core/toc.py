# -*- coding: utf-8 -*-
"""
Автоматический сбор содержания тома по ГОСТ Р 21.101-2020.

Задача: сформировать «Содержание / Опись» не из ручного списка (requisites.sections),
а автоматически — по фактической структуре тома (project.structure) и реальному
числу листов каждого раздела/документа/приложения.

Модуль строит ЛИНЕЙНУЮ КАРТУ ЛИСТОВ тома: упорядоченный перечень позиций
(разделы, документы, приложения, лист изменений), каждая из которых знает свой
диапазон сквозной нумерации (первый и последний номер листа). На её основе:

  * формируются строки содержания («Обозначение / Наименование … Лист / Листов»);
  * вычисляется фактическое общее число листов «Листов M»;
  * все экспортёры (DOCX, PDF, HTML-предпросмотр) используют один и тот же
    порядок и одинаковые номера листов, чтобы содержания не расходились.

Единый источник сквозной нумерации = build_layout() в этом модуле. Он применяется
и в штампе (графы «Лист N / Листов M»), и в содержании.

Порядок следования блоков определяется настраиваемой структурой project.structure
(список {kind, needed}); исключённые из сборки элементы не входят в содержание и
не занимают листы.
"""
from __future__ import annotations

import os
import textwrap
from dataclasses import dataclass, field

from ..models import (
    VolumeProject, DocKind,
    SECTION_COVER, SECTION_TITLE, SECTION_CONTENT,
    SECTION_PZ, SECTION_GRAPHIC, SECTION_APPENDIX, SECTION_CHANGES,
    SECTION_COMPOSITION, SECTION_LABELS, section_designation,
)


# --- Число листов одного документа / приложения ---------------------------------

def _is_docx(path: str) -> bool:
    return bool(path) and path.lower().endswith((".docx", ".doc"))


def _pdf_count(path: str) -> int:
    """Число страниц PDF (минимум 1 лист)."""
    try:
        import pymupdf
        with pymupdf.open(path) as pdf:
            return max(1, pdf.page_count)
    except Exception:
        return 1


def _paginate_docx(content) -> list:
    """Разбивает содержимое DOCX на «страницы» для подсчёта листов.

    Переносим логику из preview_html._paginate_docx, чтобы число листов,
    показываемое в содержании и в штампе, совпадало с предпросмотром.
    Возвращает список «страниц» (списков элементов).
    """
    stream = []
    blocks = getattr(content, "blocks", None)
    if blocks is None:
        stream = [("line", {"text": getattr(content, "plain_text", ""), "bold": False})]
    else:
        stream = list(blocks)

    budget_mm = 255.0     # доступная высота текстовой области листа (А4)
    line_h = 5.5          # высота строки текста 11pt
    cell_h = 4.3          # высота строки таблицы

    pages = []
    cur = []
    cur_h = 0.0
    for item in stream:
        if isinstance(item, (list, tuple)) and len(item) > 1:
            kind, data = item[0], item[1]
        else:
            continue
        if kind == "line":
            text = (data.get("text") if isinstance(data, dict) else "") or ""
            if not str(text).strip():
                continue
            lines = max(1, len(textwrap.wrap(str(text), 81) or [str(text)]))
            h = lines * line_h
        elif kind == "table_rows":
            rows = data if isinstance(data, list) else []
            h = len(rows) * cell_h
            if h <= 0:
                continue
        else:
            continue
        if cur and cur_h + h > budget_mm:
            pages.append(cur)
            cur = []
            cur_h = 0.0
        cur.append(item)
        cur_h += h
    if cur:
        pages.append(cur)
    if not pages:
        pages = [[]]
    return pages


def doc_sheet_count(d, first_form=2) -> int:
    """Число листов одного документа (DOCX — пагинация, PDF — страницы).

    Единый источник подсчёта — preview_html._doc_sheet_count (та же пагинация,
    что и в предпросмотре). Это гарантирует, что число листов в содержании
    совпадает с тем, что пользователь видит в предпросмотре листа.

    :param first_form: форма штампа первого листа потока документа
        (2 — форма 5, "2a" — форма 6; влияет на бюджет первого листа).

    Если файл недоступен — используется заявленное число листов d.pages
    (минимум 1), т.к. лист-заглушка всё равно занимает место.
    """
    path = getattr(d, "path", "") or ""
    if not path or not os.path.exists(path):
        return max(1, int(getattr(d, "pages", 0) or 1))   # лист-заглушка
    try:
        # Ленивый импорт, чтобы избежать жёсткого цикла (preview_html не
        # импортирует toc на верхнем уровне).
        from .preview_html import _doc_sheet_count as _pv_doc_sheet
        return max(1, _pv_doc_sheet(d, first_form=first_form))
    except Exception:
        pass
    if _is_docx(path):
        try:
            from .docx_reader import read_docx
            return max(1, len(_paginate_docx(read_docx(path))))
        except Exception:
            return max(1, int(getattr(d, "pages", 0) or 1))
    return _pdf_count(path)


def appendix_sheet_count(a) -> int:
    """Число листов одного приложения (PDF).

    Если файл недоступен — используется заявленное число листов a.pages
    (минимум 1).
    """
    f = getattr(a, "file_path", "") or ""
    if not f or not os.path.exists(f):
        return max(1, int(getattr(a, "pages", 0) or 1))
    pages = getattr(a, "pages", 0) or 0
    return max(1, pages) if pages else _pdf_count(f)


def composition_sheet_count(project) -> int:
    """Число листов раздела «Состав проектной документации» (минимум 1).

    Пагинация таблицы (сколько строк входит на лист формы 5 / формы 6) —
    в едином источнике preview_html.composition_pages: по нему же строятся
    предпросмотр, Word и PDF, поэтому номера листов в содержании совпадают.
    Раздел занимает лист даже с пустой таблицей — на нём видна шапка таблицы.
    """
    try:
        from .preview_html import composition_pages
        return max(1, len(composition_pages(project)))
    except Exception:
        return 1


# --- Карта листов (линейная нумерация) ------------------------------------------

@dataclass
class ContentItem:
    """Позиция содержания: раздел/документ/приложение.

    title       — наименование для строки содержания;
    designation — обозначение (шифр, номер приложения и т.п.);
    first_sheet — номер первого листа позиции (сквозной);
    last_sheet  — номер последнего листа позиции;
    count       — число листов позиции;
    kind        — тип: 'cover' | 'title' | 'content' | 'pz' | 'graphic' |
                  'appendix' | 'changes';
    section_key — ключ раздела, к которому относится позиция (для группировки);
    is_section_header — True, если строка является «разделом» (не документом).
    """
    title: str = ""
    designation: str = ""
    first_sheet: int = 0
    last_sheet: int = 0
    count: int = 0
    kind: str = ""
    section_key: str = ""
    is_section_header: bool = False

    @property
    def sheet_label(self) -> str:
        """Номер первого листа позиции («N»); для диапазона — «N-M» (0 — «—»)."""
        if self.count <= 0:
            return "—"
        return f"{self.first_sheet}-{self.last_sheet}" if self.count > 1 else str(self.first_sheet)


@dataclass
class VolumeLayout:
    """Линейная карта листов тома с итоговым числом листов.

    items — упорядоченный перечень ContentItem;
    total_sheets — общее число листов тома («Листов M»).
    content_sheets — число листов, занимаемых самим содержанием (как правило 1).
    """
    items: list = field(default_factory=list)
    total_sheets: int = 0
    content_sheets: int = 1


def _pz_toc_sheets(project) -> int:
    """Число листов автоматического оглавления ПЗ (0 — оглавление не нужно).

    Единый источник — app.core.pz_toc. Импорт ленивый: pz_toc использует
    пагинацию предпросмотра, а модуль toc импортируется предпросмотром.
    """
    try:
        from .pz_toc import project_toc_sheets
        return int(project_toc_sheets(project) or 0)
    except Exception:
        return 0


def doc_label(d) -> str:
    """Человекочитаемое имя документа (по имени файла)."""
    p = getattr(d, "path", "") or ""
    return os.path.basename(p) if p else "—"


def appendix_designation(a) -> str:
    """Обозначение приложения для графы «Обозначение» содержания.

    По ГОСТ Р 21.101-2020 в этой графе у приложения пишется его номер со
    словом «Приложение»: «Приложение 1», «Приложение А». Без номера остаётся
    просто «Приложение».
    """
    number = (getattr(a, "number", "") or "").strip()
    return f"Приложение {number}" if number else "Приложение"


def build_layout(project: VolumeProject) -> VolumeLayout:
    """Строит линейную карту листов тома (единый источник нумерации).

    Возвращает VolumeLayout: упорядоченные позиции (ContentItem) с фактическими
    диапазонами листов и общее число листов «Листов M».
    """
    layout = VolumeLayout()
    sheet = 1                       # текущий сквозной номер листа
    req = project.requisites
    volume_id = req.volume_id or ""

    def design(kind: str) -> str:
        """Обозначение раздела: обозначение тома + доп. обозначение раздела."""
        return section_designation(project, kind)

    def _sections_in_order():
        for entry in project.structure:
            if entry.get("needed", True):
                yield entry.get("kind")

    for kind in _sections_in_order():
        if kind == SECTION_COVER:
            # Обложка — без рамки/штампа, в сквозную нумерацию не входит.
            layout.items.append(ContentItem(
                title=SECTION_LABELS.get(SECTION_COVER, "Обложка"),
                designation=volume_id, first_sheet=0, last_sheet=0, count=0,
                kind="cover", section_key=SECTION_COVER, is_section_header=True))
            continue
        if kind == SECTION_TITLE:
            # Титульный лист (как и обложка) не нумеруется: в предпросмотре и
            # экспорте квадрат/номер листа на нём не проставляется. В содержании
            # он указывается без номера («—»).
            layout.items.append(ContentItem(
                title=SECTION_LABELS.get(SECTION_TITLE, "Титульный лист"),
                designation=volume_id, first_sheet=0, last_sheet=0, count=0,
                kind="title", section_key=SECTION_TITLE, is_section_header=True))
            continue
        if kind == SECTION_COMPOSITION:
            # Состав проектной документации: листы с рамкой и основной надписью
            # формы 5 (первый) и формы 6 (последующие). Раздел занимает лист
            # даже с пустой таблицей — на нём видна шапка таблицы.
            n = composition_sheet_count(project)
            layout.items.append(ContentItem(
                title=SECTION_LABELS.get(SECTION_COMPOSITION,
                                         "Состав проектной документации"),
                designation=design(SECTION_COMPOSITION), first_sheet=sheet,
                last_sheet=sheet + n - 1, count=n, kind="composition",
                section_key=SECTION_COMPOSITION, is_section_header=True))
            sheet += n
            continue
        if kind == SECTION_CONTENT:
            # Содержание — первый нумерованный лист тома (№1).
            first = sheet
            layout.content_sheets = 1
            sheet += layout.content_sheets
            layout.items.append(ContentItem(
                title=SECTION_LABELS.get(SECTION_CONTENT, "Содержание"),
                designation=design(SECTION_CONTENT), first_sheet=first,
                last_sheet=first + layout.content_sheets - 1,
                count=layout.content_sheets, kind="content",
                section_key=SECTION_CONTENT, is_section_header=True))
            continue
        if kind == SECTION_PZ:
            docs = project.documents_for(DocKind.POYASNITELNAYA)
            if not docs:
                # В разделе нет включённых документов: он не занимает лист и не
                # попадает в содержание. Иначе после снятия флажка «в сборку»
                # у документа оставался бы пустой лист-заглушка, и том выглядел
                # бы так, будто документ не исключён.
                continue
            # Автоматическое оглавление записки (app.core.pz_toc) занимает свои
            # листы ПЕРЕД первым документом раздела и также входит в «Листов M».
            toc_sheets = _pz_toc_sheets(project)
            # Форма первого листа каждого документа — как в предпросмотре
            # (preview_html.pz_doc_first_forms): сказывается на разбивке DOCX.
            try:
                from .preview_html import pz_doc_first_forms
                forms = pz_doc_first_forms(project)
            except Exception:
                forms = []
            for index, d in enumerate(docs):
                ff = forms[index] if index < len(forms) else 2
                n = doc_sheet_count(d, first_form=ff) + (toc_sheets if index == 0 else 0)
                layout.items.append(ContentItem(
                    title=doc_label(d), designation=design(SECTION_PZ),
                    first_sheet=sheet, last_sheet=sheet + n - 1, count=n,
                    kind="pz", section_key=SECTION_PZ, is_section_header=False))
                sheet += n
            continue
        if kind == SECTION_GRAPHIC:
            docs = project.documents_for(DocKind.GRAPHIC)
            if not docs:
                # См. ПЗ: без включённых документов раздел не занимает лист.
                continue
            # Лист «Графическая часть» с перечнем листов идёт ПЕРВЫМ листом
            # раздела (если включён в настройках), поэтому сдвигает нумерацию
            # всех листов графики: без этого содержание разошлось бы с
            # предпросмотром и с самим томом.
            gl = getattr(project, "graphic_sheet_list", None)
            if gl is not None and gl.enabled:
                layout.items.append(ContentItem(
                    title="Графическая часть. Перечень листов",
                    designation=design(SECTION_GRAPHIC),
                    first_sheet=sheet, last_sheet=sheet, count=1,
                    kind="graphic", section_key=SECTION_GRAPHIC,
                    is_section_header=False))
                sheet += 1
            for d in docs:
                # В разделе графики первый лист документа тоже рисуется
                # с формы 6 (first=False в дереве) — та же модель высот.
                n = doc_sheet_count(d, first_form="2a")
                layout.items.append(ContentItem(
                    title=doc_label(d), designation=design(SECTION_GRAPHIC),
                    first_sheet=sheet, last_sheet=sheet + n - 1, count=n,
                    kind="graphic", section_key=SECTION_GRAPHIC, is_section_header=False))
                sheet += n
            continue
        if kind == SECTION_APPENDIX:
            apps = project.appendices_needed()
            if not apps:
                continue
            for a in apps:
                n = appendix_sheet_count(a)
                design = appendix_designation(a)
                layout.items.append(ContentItem(
                    # Наименование строки содержания — название приложения;
                    # номер вынесен в графу «Обозначение» («Приложение 1»).
                    title=(getattr(a, "name", "") or "").strip() or design,
                    designation=design,
                    first_sheet=sheet, last_sheet=sheet + n - 1, count=n,
                    kind="appendix", section_key=SECTION_APPENDIX, is_section_header=False))
                sheet += n
            continue
        if kind == SECTION_CHANGES:
            if not project.changes:
                continue
            # Лист регистрации изменений: 1 лист на каждые ~20 записей.
            n = max(1, (len(project.changes) + 19) // 20)
            layout.items.append(ContentItem(
                title=SECTION_LABELS.get(SECTION_CHANGES, "Лист регистрации изменений"),
                designation=volume_id, first_sheet=sheet, last_sheet=sheet + n - 1,
                count=n, kind="changes", section_key=SECTION_CHANGES,
                is_section_header=True))
            sheet += n
            continue

    layout.total_sheets = sheet - 1
    return layout


def build_content_items(project: VolumeProject) -> list:
    """Перечень строк содержания (ContentItem) для вставки в том.

    Группирует карту листов так, как она отображается в содержании: документы
    внутри ПЗ/Графики объединены заголовком раздела, приложения идут отдельным
    блоком, лист изменений — последней строкой.

    :return: упорядоченный список ContentItem (только включённые позиции).
    """
    layout = build_layout(project)
    if not layout.items:
        return []
    return layout.items


def content_rows(project: VolumeProject) -> list:
    """Строки содержания в виде кортежей (наименование, обозначение, листы).

    Удобно для построения таблицы содержания в DOCX/PDF/HTML.
    Обложка и титульный лист в содержание НЕ включаются (это служебные листы
    тома); сквозная нумерация при этом сохраняется как была.
    Ручной перечень requisites.sections добавляется как дополнительный блок,
    если он заполнен и не дублирует автоматические разделы.

    В содержании — РАЗДЕЛЫ тома: документы внутри раздела объединяются в одну
    строку. Исключение — приложения: каждое идёт ОТДЕЛЬНОЙ строкой, потому что
    у него своё обозначение («Приложение 1», «Приложение 2») и своё название.
    В графе «Листы» стоит номер ПЕРВОГО листа строки (без слов «Лист»/«Листы»
    и без диапазона) — так требует ГОСТ Р 21.101-2020. Таким образом содержание
    отражает состав тома, а не перечень файлов.
    """
    items = build_content_items(project)
    # Строки собираются списком словарей: у приложений своё наименование и
    # обозначение, поэтому ключ группировки и подпись строки хранятся раздельно.
    rows = []
    for it in items:
        # Обложка (cover) и титульный лист (title) не входят в содержание.
        if it.kind in ("cover", "title"):
            continue
        # Позиции одного раздела склеиваются в одну строку: у ПЗ/графики это
        # несколько документов + лист перечня, в содержании они — один раздел.
        key = it.section_key or it.kind
        if it.kind == "appendix":
            # Каждое приложение — своя строка: ключ уникален, поэтому условие
            # склейки ниже для них не срабатывает.
            rows.append({"key": f"{key}#{len(rows)}", "title": it.title,
                         "designation": it.designation or "",
                         "label": it.sheet_label, "last": it.last_sheet})
            continue
        if rows and rows[-1]["key"] == key:
            # Диапазон расширяется до последнего листа раздела.
            rows[-1]["last"] = max(rows[-1]["last"], it.last_sheet)
            continue
        rows.append({"key": key, "title": _row_title(key),
                     "designation": it.designation or "",
                     "label": it.sheet_label, "last": it.last_sheet})
    # Ручные строки из реквизитов (если заданы) — после автоматических.
    manual = []
    for s in project.requisites.sections or []:
        s = s.strip()
        if not s:
            continue
        if any(r["designation"] == s or r["title"] == s for r in rows):
            continue
        manual.append((s, "", "—"))
    return [(r["title"], r["designation"], _row_sheets(r["label"], r["last"]))
            for r in rows] + manual


def _row_title(key: str) -> str:
    """Наименование раздела для строки содержания (ключ -> подпись)."""
    return SECTION_LABELS.get(key, key)


def _row_sheets(label: str, last: int) -> str:
    """Графа «Листы» строки содержания: номер ПЕРВОГО листа строки.

    По ГОСТ Р 21.101-2020 в графе указывается только номер листа, с которого
    начинается раздел, — без слов «Лист»/«Листы» и без диапазона: «3», а не
    «Листы 3 — 10». Для не занимающих лист позиций (обложка, титул) — «—».
    """
    if not last:
        return "—"
    # Первый лист — из метки позиции («N» / «N-M»).
    import re
    m = re.search(r"(\d+)", label or "")
    return m.group(1) if m else (label or "—")


def total_sheets(project: VolumeProject) -> int:
    """Общее число листов тома «Листов M» (единый источник для штампа)."""
    return build_layout(project).total_sheets


def changes_sheet_count(project: VolumeProject) -> int:
    """Число листов журнала изменений."""
    if not project.changes:
        return 0
    return max(1, (len(project.changes) + 19) // 20)
