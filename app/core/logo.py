# -*- coding: utf-8 -*-
"""Подготовка логотипа организации для графы 9 основной надписи.

Логотип задаётся PDF-файлом. Для встраивания в HTML-предпросмотр и DOCX
(Word не умеет вставлять PDF как картинку) PDF растеризуется в PNG,
дополнительно обрезаются белые поля страницы.

Растеризация двухпроходная и с ограничениями по памяти:
  1) пробный рендер при низком разрешении — ищем границы «непустой» области;
  2) рендер ТОЛЬКО этой области с зумом, подобранным так, чтобы логотип на
     листе имел полиграфическое разрешение (TARGET_PX пикселей по большей
     стороне), но суммарное число пикселей не превышало MAX_PIXELS.
Благодаря этому логотип, занимающий угол листа A4, не растеризует весь лист
в огромный bitmap.

Результат кэшируется в каталоге tmp/logo_cache и пересобирается только при
изменении исходного PDF или параметров растеризации.

Модуль независим от Qt и пригоден для фоновых расчётов. Любая ошибка
приводит к пустой строке — вызывающий код тогда выводит графу 9 текстом.
"""
from __future__ import annotations

import hashlib
import os

# Разрешение пробного рендера (dpi) — только для поиска границ изображения.
PROBE_DPI = 150

# Целевая ширина/высота логотипа в пикселях (по большей стороне).
# Логотип печатается в ячейке шириной ~50 мм: 1400 px / 50 мм ~ 710 dpi.
TARGET_PX = 1400

# Ограничения зума и общего числа пикселей (защита от «мегабитмапа»).
MAX_ZOOM = 24.0
MAX_PIXELS = 4_000_000

# Порог «белого» при поиске границ изображения (0..255).
_WHITE_LEVEL = 245

# Небольшой запас вокруг обрезанного изображения (в точках PDF).
_CROP_PAD_PT = 1.0

# Каталог кэша (tmp/logo_cache рядом с корнем проекта).
_CACHE_DIRNAME = os.path.join("tmp", "logo_cache")


def project_root() -> str:
    """Корень проекта (каталог, содержащий app/ и tmp/)."""
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def cache_dir() -> str:
    """Каталог кэша растеризованных логотипов (создаётся при обращении)."""
    path = os.path.join(project_root(), _CACHE_DIRNAME)
    os.makedirs(path, exist_ok=True)
    return path


def _signature(pdf_path: str) -> str:
    """Ключ кэша: путь + размер + mtime исходного PDF и параметры растеризации."""
    st = os.stat(pdf_path)
    raw = "%s|%d|%d|%d|%d" % (os.path.abspath(pdf_path), st.st_size,
                              int(st.st_mtime), PROBE_DPI, TARGET_PX)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def _ink_bbox(samples: bytes, stride: int, w: int, h: int):
    """Границы «непустой» (не белой) области в оттенках серого.

    Возвращает (left, top, right, bottom) в пикселях или None, если
    изображение целиком белое.
    """
    level = _WHITE_LEVEL

    def row_has_ink(y: int) -> bool:
        base = y * stride
        for x in range(w):
            if samples[base + x] < level:
                return True
        return False

    def col_has_ink(x: int, y0: int, y1: int) -> bool:
        for y in range(y0, y1):
            if samples[y * stride + x] < level:
                return True
        return False

    top, bottom = 0, h
    while top < bottom and not row_has_ink(top):
        top += 1
    while bottom > top and not row_has_ink(bottom - 1):
        bottom -= 1
    if top >= bottom:
        return None
    left, right = 0, w
    while left < right and not col_has_ink(left, top, bottom):
        left += 1
    while right > left and not col_has_ink(right - 1, top, bottom):
        right -= 1
    if left >= right:
        return None
    return (left, top, right, bottom)


def _render_page_png(page, out_path: str) -> bool:
    """Растеризует страницу PDF в PNG с обрезкой белых полей.

    Возвращает True при успехе.
    """
    import pymupdf

    # 1) Пробный рендер — ищем границы изображения.
    probe = page.get_pixmap(dpi=PROBE_DPI, alpha=False)
    gray = pymupdf.Pixmap(pymupdf.csGRAY, probe)
    bbox = _ink_bbox(gray.samples, gray.stride, gray.width, gray.height)

    page_rect = page.rect
    if bbox is None:
        # Пустая страница — рендерим как есть, без обрезки.
        clip = page_rect
    else:
        k = 72.0 / PROBE_DPI            # пиксели пробного рендера -> точки PDF
        clip = pymupdf.Rect(bbox[0] * k, bbox[1] * k,
                            bbox[2] * k, bbox[3] * k)
        clip = pymupdf.Rect(clip.x0 - _CROP_PAD_PT, clip.y0 - _CROP_PAD_PT,
                            clip.x1 + _CROP_PAD_PT, clip.y1 + _CROP_PAD_PT)
        clip = clip & page_rect
        if clip.is_empty:
            clip = page_rect

    # 2) Зум под целевую ширину/высоту с ограничением по числу пикселей.
    w_pt = max(1.0, clip.width)
    h_pt = max(1.0, clip.height)
    zoom = TARGET_PX / max(w_pt, h_pt)
    zoom = max(0.05, min(MAX_ZOOM, zoom))
    # Ограничение суммарного разрешения.
    while w_pt * zoom * h_pt * zoom > MAX_PIXELS and zoom > 0.05:
        zoom *= 0.85

    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip, alpha=False)
    pix.save(out_path)
    return True


def render_logo_png(pdf_path: str, *, force: bool = False) -> str:
    """Растеризует первую страницу PDF-логотипа в PNG и возвращает путь к файлу.

    Ошибки не поднимаются: при любой проблеме возвращается пустая строка, и
    вызывающий код отрисует графу 9 текстом.
    """
    if not pdf_path or not os.path.exists(pdf_path):
        return ""
    try:
        key = _signature(pdf_path)
    except OSError:
        return ""
    out_path = os.path.join(cache_dir(), "logo_%s.png" % key)
    if os.path.exists(out_path) and not force:
        return out_path
    try:
        import pymupdf
    except Exception:
        return ""
    # PyMuPDF определяет формат по расширению, поэтому временный файл тоже .png;
    # запись через временный файл + os.replace делает кэш атомарным.
    tmp_path = out_path + ".part.png"
    try:
        with pymupdf.open(pdf_path) as doc:
            if doc.page_count < 1:
                return ""
            if not _render_page_png(doc[0], tmp_path):
                return ""
        os.replace(tmp_path, out_path)
        return out_path
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        return ""


def logo_pixel_size(pdf_path: str) -> tuple:
    """Возвращает (ширина, высота) подготовленного PNG в пикселях (0, 0 — ошибка)."""
    png = render_logo_png(pdf_path)
    if not png:
        return (0, 0)
    try:
        import pymupdf
        with pymupdf.open(png) as doc:
            if doc.page_count < 1:
                return (0, 0)
            rect = doc[0].rect
            return (int(round(rect.width)), int(round(rect.height)))
    except Exception:
        return (0, 0)


def fit_into(width: float, height: float, box_w: float, box_h: float) -> tuple:
    """Вписывает прямоугольник (width x height) в бокс, сохраняя пропорции.

    Возвращает (ширина, высота) вписанного прямоугольника. Если исходные или
    целевые размеры некорректны — возвращает (0.0, 0.0).
    """
    try:
        width, height = float(width), float(height)
        box_w, box_h = float(box_w), float(box_h)
    except (TypeError, ValueError):
        return (0.0, 0.0)
    if width <= 0 or height <= 0 or box_w <= 0 or box_h <= 0:
        return (0.0, 0.0)
    scale = min(box_w / width, box_h / height)
    return (width * scale, height * scale)


def logo_fit_mm(pdf_path: str, box_w_mm: float, box_h_mm: float, *,
                padding_mm: float = 0.0) -> tuple:
    """Размер логотипа (ширина, высота) в мм, вписанного в ячейку box_w x box_h.

    Это и есть автомасштабирование: логотип занимает максимум доступного места,
    но не искажается и не выходит за границы ячейки.
    """
    px = logo_pixel_size(pdf_path)
    if px == (0, 0):
        return (0.0, 0.0)
    pad = max(0.0, float(padding_mm or 0.0))
    avail_w = float(box_w_mm) - 2 * pad
    avail_h = float(box_h_mm) - 2 * pad
    return fit_into(px[0], px[1], avail_w, avail_h)
