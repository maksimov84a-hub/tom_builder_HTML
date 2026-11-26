# -*- coding: utf-8 -*-
"""Измерение ширины текста в шрифте оформления и подбор кегля под ячейку.

Основная надпись по ГОСТ имеет узкие графы (10–20 мм), а значения в них
(«17.11.26», «Кол.уч», «Стадия») при штатном кегле 10–12 pt шире графы.
Word не уменьшает кегль сам, поэтому текст переносится, высоты строк
(`atLeast`) растут, и штамп «разъезжается» — занимает вдвое больше места и
наезжает на текст листа.

Модуль даёт точный замер: ширины берутся из файла шрифта (PyMuPDF читает
метрики глифов), поэтому подобранный кегль совпадает с тем, что видит Word.
Если файл шрифта недоступен, используется консервативная оценка, чтобы текст
гарантированно влез (пусть чуть мельче — но без переносов).
"""
from __future__ import annotations

#: Кэш объектов шрифтов PyMuPDF по пути к файлу.
_FONT_CACHE: dict = {}
#: Кэш ширины строки: (путь, текст, кегль) -> ширина в мм.
_WIDTH_CACHE: dict = {}

#: Средняя ширина знака относительно кегля — запасной вариант без метрик.
#: Для ISOCPEUR (узкий чертёжный шрифт) среднее отношение ≈ 0.55.
_FALLBACK_ADVANCE = 0.56


def _font(path: str | None):
    """Объект шрифта PyMuPDF для файла (None — метрики недоступны)."""
    if not path:
        return None
    if path in _FONT_CACHE:
        return _FONT_CACHE[path]
    font = None
    try:
        import pymupdf
        font = pymupdf.Font(fontfile=path)
    except Exception:
        font = None
    _FONT_CACHE[path] = font
    return font


def font_path(family: str | None = None) -> str:
    """Путь к файлу шрифта оформления (пустая строка — метрик нет)."""
    try:
        from .fonts import font_file
        return font_file(family) or ""
    except Exception:
        return ""


def text_width_mm(text: str, size_pt: float, path: str | None = None) -> float:
    """Ширина однострочного текста в мм при кегле size_pt.

    :param path: файл шрифта; None — взять шрифт оформления из настроек.
    """
    if not text or size_pt <= 0:
        return 0.0
    if path is None:
        path = font_path()
    key = (path, text, round(size_pt, 2))
    cached = _WIDTH_CACHE.get(key)
    if cached is not None:
        return cached
    font = _font(path)
    if font is not None:
        try:
            # text_length возвращает пункты.
            width = font.text_length(text, fontsize=size_pt) * 25.4 / 72.0
        except Exception:
            width = len(text) * size_pt * _FALLBACK_ADVANCE * 25.4 / 72.0
    else:
        width = len(text) * size_pt * _FALLBACK_ADVANCE * 25.4 / 72.0
    _WIDTH_CACHE[key] = width
    return width


def fit_size(text: str, avail_mm: float, prefer_pt: float,
             min_pt: float = 5.0, path: str | None = None,
             step: float = 0.5) -> float:
    """Наибольший кегль ≤ prefer_pt, при котором текст шириной влезает в avail_mm.

    Шаг подбора — 0.5 pt (как в Word). Если текст не влезает даже при min_pt,
    возвращается min_pt: перенос в узкой графе всё равно недопустим, а
    незначительный выход за линию незаметнее разъехавшегося штампа.
    """
    if avail_mm <= 0 or not text:
        return prefer_pt
    if not path:
        path = font_path()
    size = float(prefer_pt)
    while size > min_pt:
        if text_width_mm(text, size, path) <= avail_mm:
            return size
        size = round(size - step, 2)
    return float(min_pt)


def widest_size(texts, avail_mm: float, prefer_pt: float,
                min_pt: float = 5.0, path: str | None = None) -> float:
    """Общий кегль для группы строк: минимальный из подобранных по каждой.

    Нужен, когда несколько значений стоят в одинаковых графах (строки подписей,
    столбцы «Изм./Кол.уч/...») — кегль у них должен быть один.
    """
    sizes = [fit_size(t, avail_mm, prefer_pt, min_pt, path)
             for t in texts if t]
    return min(sizes) if sizes else float(prefer_pt)


def _tokens(text: str) -> list:
    """Слова с сохранением пробелов: ['Слово', ' ', 'слово', ...].

    Перенос в Word идёт по пробелам И ДЕФИСАМ, поэтому единица переноса —
    часть слова до дефиса включительно («мало-» / «мобильных»): без этого
    оценка строк занижалась, браузер рисовал лишнюю строку, и последний абзац
    листа обрезался (текст не доходил до штампа). Пробелы держим отдельными
    элементами, чтобы их ширина учитывалась.
    """
    out = []
    word = ""
    for ch in text:
        if ch in " \t\u00a0":
            if word:
                out.append(word)
                word = ""
            out.append(ch)
        else:
            word += ch
            if ch == "-":
                # Дефис — точка переноса: слово разбивается после него.
                out.append(word)
                word = ""
    if word:
        out.append(word)
    return out


def visual_lines(text: str, width_mm: float, size_pt: float,
                 path: str | None = None,
                 first_line_mm: float = 0.0) -> int:
    """Число визуальных строк текста при заданной ширине колонки (метрики шрифта).

    Перенос считается по фактической ширине слов из файла шрифта — так же, как
    это делает Word. Оценка «знаков в строке» ошибалась на длинных словах
    («санитарно-эпидемиологические»): в узкой колонке такое слово не влезает
    целиком, и Word переносит раньше, а по числу знаков выходило меньше строк.
    Из-за этого пагинация HTML и вёрстка Word расходились.

    Длинное слово шире колонки Word переносит по символам — здесь оно
    разбивается на части по ширине, чтобы число строк совпало.

    :param width_mm: доступная ширина строки (мм), уже без отступов абзаца.
    :param size_pt: кегль текста (pt).
    :param first_line_mm: абзацный отступ первой строки («красная строка», мм):
        первая строка короче остальных на эту величину.
    """
    if not text:
        return 1
    if width_mm <= 1.0 or size_pt <= 0:
        return 1
    if path is None:
        path = font_path()
    limit = float(width_mm)
    first_limit = max(1.0, limit - max(0.0, float(first_line_mm)))

    lines = 1
    current = 0.0

    def advance(token: str) -> float:
        return text_width_mm(token, size_pt, path)

    def room() -> float:
        return first_limit if lines == 1 else limit

    for token in _tokens(text):
        w = advance(token)
        if token.isspace():
            # Пробел в конце строки не переносится — он «съедается».
            if current + w > room() and current > 0:
                lines += 1
                current = 0.0
            else:
                current += w
            continue
        if current + w <= room():
            current += w
            continue
        # Слово не влезает в остаток строки.
        if current > 0:
            lines += 1
            current = 0.0
        if w <= room():
            current = w
            continue
        # Слово шире целой строки: Word переносит по символам.
        parts = 0
        used = 0.0
        for ch in token:
            cw = advance(ch)
            if used + cw > room() and used > 0:
                parts += 1
                used = 0.0
            used += cw
        current = used
        lines += parts
    return max(1, lines)

