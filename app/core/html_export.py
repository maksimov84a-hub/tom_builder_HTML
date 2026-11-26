# -*- coding: utf-8 -*-
"""Экспорт HTML-предпросмотра в PDF через QtWebEngine."""
from __future__ import annotations

import os
import re
from PySide6.QtCore import QEventLoop, QMarginsF, QSizeF, QTimer, QUrl
from PySide6.QtGui import QPageLayout, QPageSize
from PySide6.QtWebEngineCore import QWebEnginePage
from .preview_html import build_section_html, build_volume_html

SHEET_SIZE_MM = {"cover": (210.0, 297.0), "title": (210.0, 297.0),
                 "content": (210.0, 297.0), "changes": (210.0, 297.0),
                 "stamp_f1": (210.0, 297.0), "stamp_f2": (210.0, 297.0),
                 "stamp_f2a": (210.0, 297.0)}
DEFAULT_SIZE_MM = (210.0, 297.0)


def _page_layout(w_mm: float, h_mm: float) -> QPageLayout:
    # QPageSize в миллиметрах уже задаёт физический размер листа (ширина x высота).
    # Явная ориентация Landscape/Portrait в QPageLayout ПОВОРАЧИВАЕТ страницу:
    # для альбомного листа (w > h) это даёт перевёрнутый формат (h x w). Поэтому
    # всегда используем Portrait — он не вращает страницу, а размер остаётся w x h.
    #
    # Размер задаётся в ТОЧКАХ, а не в миллиметрах: QPageSize(Millimeter)
    # округляет стороны до целых пунктов (A4 → 595 x 842 pt = 209.9 мм), и
    # правая линия рамки расходилась с краем листа. Точные пункты A4 —
    # 595.276 x 841.890 (gost.A4_W_PT / A4_H_PT).
    ps = QPageSize(QSizeF(_mm_to_pt(w_mm), _mm_to_pt(h_mm)), QPageSize.Point)
    return QPageLayout(ps, QPageLayout.Portrait, QMarginsF(0.0, 0.0, 0.0, 0.0))


def _mm_to_pt(mm: float) -> float:
    """Миллиметры в типографские пункты (1 pt = 25.4/72 мм)."""
    return float(mm) * 72.0 / 25.4


def _fix_page_size(pdf_path: str, w_mm: float, h_mm: float) -> None:
    """Ставит страницам точный размер листа в пунктах.

    Chromium печатает страницу на 0.15 мм уже и на 0.1 мм выше заданного
    размера: на печати правая линия рамки уходила бы от края листа. Размер
    выставляется через mediabox — содержимое при этом не масштабируется.
    """
    import pymupdf
    w_pt, h_pt = _mm_to_pt(w_mm), _mm_to_pt(h_mm)
    with pymupdf.open(pdf_path) as doc:
        changed = False
        for page in doc:
            box = page.mediabox
            if abs(box.width - w_pt) > 0.01 or abs(box.height - h_pt) > 0.01:
                page.set_mediabox(pymupdf.Rect(0, 0, w_pt, h_pt))
                changed = True
        if changed:
            doc.saveIncr()


def html_to_pdf(html: str, out_path: str, w_mm: float = 210.0,
                h_mm: float = 297.0, timeout_ms: int = 20000) -> str:
    """Печатает HTML в PDF через QtWebEngine в GUI-потоке."""
    page = QWebEnginePage()
    loop = QEventLoop()
    result = {"data": None, "error": None}

    def on_load(ok: bool) -> None:
        if not ok:
            result["error"] = "HTML не удалось загрузить"
            loop.quit()
            return

        def on_done(data) -> None:
            if data:
                result["data"] = bytes(data)
            else:
                result["error"] = "QtWebEngine вернул пустой PDF"
            loop.quit()

        page.printToPdf(on_done, _page_layout(w_mm, h_mm))

    # HTML-предпросмотр рассчитан на экран (с внешним отступом и тенью).
    # Перед печатью убираем эти декоративные свойства, чтобы лист целиком
    # попадал в PDF без дополнительного поля вокруг рамки ГОСТ.
    # @page{size} фиксирует точный физический размер страницы: без него Chrome
    # печатает чуть ниже заданной высоты и «выплёскивает» пустой второй лист.
    print_html = html.replace(
        "</head>",
        f"<style>@page{{size:{w_mm:g}mm {h_mm:g}mm;margin:0;}} "
        "body{background:#fff;} "
        ".sheet{margin:0 !important;box-shadow:none !important;}</style></head>",
        1,
    )
    # Большие тома часто превышают лимит setHtml() в Chromium. Поэтому
    # печатаем через временный локальный файл — это также даёт корректный
    # документный URL для QtWebEngine.
    source_path = os.path.abspath(out_path + ".source.html")
    os.makedirs(os.path.dirname(source_path), exist_ok=True)
    with open(source_path, "w", encoding="utf-8") as source:
        source.write(print_html)
    page.loadFinished.connect(on_load)
    page.load(QUrl.fromLocalFile(source_path))
    QTimer.singleShot(timeout_ms, loop.quit)
    loop.exec()
    try:
        os.remove(source_path)
    except OSError:
        pass
    if result["error"]:
        raise RuntimeError(f"HTML → PDF: {result['error']}")
    if result["data"] is None:
        raise RuntimeError("HTML → PDF: превышено время ожидания")

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "wb") as f:
        f.write(result["data"])
    # Chromium печатает страницу на доли миллиметра меньше заданной: размер
    # выставляется точно по медиабоксу, содержимое не масштабируется.
    try:
        _fix_page_size(out_path, w_mm, h_mm)
    except Exception:
        # Не удалось поправить размер — PDF уже записан и пригоден к печати.
        pass
    return out_path


def _extract_sheet_fragments(html: str) -> list[str]:
    """Извлекает блоки .sheet, включая вложенные div-элементы."""
    starts = list(re.finditer(r"<div\s+class=[\"']sheet(?:[\"']|\s)", html, re.I))
    fragments = []
    tag_re = re.compile(r'<div\b[^>]*>|</div\s*>', re.I)
    for match in starts:
        start = match.start()
        depth = 0
        end = None
        for tag in tag_re.finditer(html, start):
            if tag.group(0).lower().startswith('</'):
                depth -= 1
                if depth == 0:
                    end = tag.end()
                    break
            elif not tag.group(0).rstrip().endswith('/>'):
                depth += 1
        if end is not None:
            fragments.append(html[start:end])
    return fragments


def _sheet_size(fragment: str) -> tuple[float, float]:
    """Возвращает физический размер листа из inline-стиля или A4."""
    m = re.search(r"<div\s+class=[\"']sheet[\"']\s+style=[\"'][^\"']*?width\s*:\s*([\d.]+)mm[^\"']*?height\s*:\s*([\d.]+)mm", fragment, re.I)
    if m:
        return float(m.group(1)), float(m.group(2))
    return DEFAULT_SIZE_MM


def _single_sheet_html(html: str, fragment: str) -> str:
    """Оставляет в документе только один лист, сохраняя общие стили."""
    body_start = re.search(r'<body\b[^>]*>', html, re.I)
    body_end = re.search(r'</body\s*>', html, re.I)
    if not body_start or not body_end:
        return html
    return html[:body_start.end()] + fragment + html[body_end.start():]


_PDF_SRC_RE = re.compile(
    r'<div\s+class=[\"\']sheet[\"\'][^>]*?'
    r'data-pdf-src="([^"]+)"[^>]*?data-pdf-page="(\d+)"',
    re.I | re.S,
)

# Регэксп для извлечения номера листа из квадрата сквозной нумерации
# (<div class="sheet-number" ...>N</div>) в HTML-фрагменте листа.
_SHEET_NUMBER_RE = re.compile(
    r'<div\s+class=[\"\']sheet-number[\"\'][^>]*>\s*([^<]*?)\s*</div>',
    re.I | re.S,
)


def _fragment_sheet_number(fragment: str) -> str:
    """Возвращает номер листа из квадрата сквозной нумерации в фрагменте."""
    m = _SHEET_NUMBER_RE.search(fragment)
    return (m.group(1).strip() if m else "")


def _fragment_source_pdf(fragment: str) -> tuple[str, int] | None:
    """Возвращает (путь_к_PDF, индекс_страницы) для листа графической части.

    Если лист помечен data-pdf-src/data-pdf-page (графическая часть) — возвращаем
    пару; иначе None (обычный лист, рендерится через HTML).
    """
    m = _PDF_SRC_RE.search(fragment)
    if not m:
        return None
    return m.group(1), int(m.group(2))


def _insert_source_pdf_page(src_pdf: str, page: int, out_path: str,
                            sheet_no: str = "") -> None:
    """Копирует одну страницу исходного PDF как есть (вектор, родной размер).

    Если задан sheet_no — поверх страницы накладывается квадрат сквозной
    нумерации, привязанный к верхнему правому углу (5 мм вниз и 5 мм влево),
    совмещённый с рамкой листа. Размер листа графической части не изменяется.
    """
    import pymupdf
    from .renderer import draw_sheet_number_box
    src = pymupdf.open(src_pdf)
    try:
        pdf = pymupdf.open()
        try:
            pdf.insert_pdf(src, from_page=page, to_page=page)
            if sheet_no:
                dst_page = pdf[0]
                rect = dst_page.rect
                w_mm = rect.width * 25.4 / 72.0
                h_mm = rect.height * 25.4 / 72.0
                draw_sheet_number_box(dst_page, w_mm, h_mm, sheet_no)
            pdf.save(out_path)
        finally:
            pdf.close()
    finally:
        src.close()


def _merge_pdfs(paths: list[str], out_path: str) -> None:
    import pymupdf
    result = pymupdf.open()
    try:
        for path in paths:
            with pymupdf.open(path) as part:
                result.insert_pdf(part)
        result.save(out_path)
    finally:
        result.close()


def default_work_dir() -> str:
    """Рабочий каталог сборки (tmp/_html_pdf рядом с корнем проекта).

    Системный каталог temp в некоторых средах недоступен для записи, поэтому
    промежуточные PDF складываются внутрь проекта — так же, как в pdf_builder.
    """
    root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(root, "tmp", "_html_pdf")


def html_file_to_pdf(html_path: str, out_path: str, work_dir: str | None = None,
                     progress_callback=None, **kwargs) -> str:
    """Преобразует HTML в PDF, сохраняя размер каждого листа.

    Листы графической части, помеченные data-pdf-src/data-pdf-page, вставляются
    напрямую из исходного PDF (вектор, родной размер) — без растеризации через
    HTML. Остальные листы рендерятся через QtWebEngine.

    :param work_dir: каталог для промежуточных PDF (по умолчанию tmp/_html_pdf).
    """
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()
    fragments = _extract_sheet_fragments(html)
    if not fragments:
        return html_to_pdf(html, out_path, **kwargs)
    tmp = work_dir or default_work_dir()
    os.makedirs(tmp, exist_ok=True)
    parts = []
    total = len(fragments)
    for index, fragment in enumerate(fragments):
        if progress_callback:
            progress_callback(index, total)
        part_path = os.path.join(tmp, f"page_{index:04d}.pdf")
        src = _fragment_source_pdf(fragment)
        if src is not None:
            # Графическая часть: вставляем исходную страницу PDF как есть,
            # накладывая квадрат сквозной нумерации (совмещённый с рамкой).
            src_pdf, page = src
            if os.path.exists(src_pdf):
                sheet_no = _fragment_sheet_number(fragment)
                _insert_source_pdf_page(src_pdf, page, part_path, sheet_no=sheet_no)
            else:
                # Файл недоступен — запасной путь через HTML (растровый вид).
                w_mm, h_mm = _sheet_size(fragment)
                html_to_pdf(_single_sheet_html(html, fragment), part_path,
                            w_mm=w_mm, h_mm=h_mm, **kwargs)
        else:
            w_mm, h_mm = _sheet_size(fragment)
            html_to_pdf(_single_sheet_html(html, fragment), part_path,
                        w_mm=w_mm, h_mm=h_mm, **kwargs)
        parts.append(part_path)
    _merge_pdfs(parts, out_path)
    if progress_callback:
        progress_callback(total, total)
    # Промежуточные листы больше не нужны.
    for part in parts:
        try:
            os.remove(part)
        except OSError:
            pass
    return out_path


def render_sheet_to_pdf(html: str, out_path: str, sheet_kind: str = "cover",
                        timeout_ms: int = 20000) -> str:
    w, h = SHEET_SIZE_MM.get(sheet_kind, DEFAULT_SIZE_MM)
    return html_to_pdf(html, out_path, w_mm=w, h_mm=h, timeout_ms=timeout_ms)


def build_volume_html_pdf(project, out_path: str, font_map=None,
                          timeout_ms: int = 20000) -> str:
    """Собирает весь том по цепочке HTML → PDF."""
    html_path = os.path.splitext(out_path)[0] + ".html"
    build_volume_html(project, html_path, font_map)
    return html_file_to_pdf(html_path, out_path, timeout_ms=timeout_ms)


def build_section_html_pdf(project, section_kind: str, out_path: str,
                           font_map=None, timeout_ms: int = 20000) -> str:
    """Собирает раздел по цепочке HTML → PDF."""
    html_path = os.path.splitext(out_path)[0] + ".html"
    build_section_html(project, section_kind, html_path, font_map)
    return html_file_to_pdf(html_path, out_path, timeout_ms=timeout_ms)


def build_html_pdf(project, out_path: str, section_kind: str | None = None,
                   font_map=None, timeout_ms: int = 20000) -> str:
    if section_kind:
        return build_section_html_pdf(project, section_kind, out_path, font_map, timeout_ms)
    return build_volume_html_pdf(project, out_path, font_map, timeout_ms)
