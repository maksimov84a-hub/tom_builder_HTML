# -*- coding: utf-8 -*-
"""Оформление абзацев DOCX: отступы, интервалы, выравнивание и списки.

Word описывает абзац не «сам по себе», а через наследование свойств:

    свойства абзаца (w:pPr)
      -> стиль абзаца (w:style/w:pPr)
      -> базовый стиль (w:basedOn)
      -> умолчания документа (w:docDefaults)
      -> умолчания приложения

Списки задаются ссылкой w:numPr (numId + ilvl) на определение нумерации
(word/numbering.xml), причём и уровень, и сам список могут быть заданы в стиле,
а параметры отступа уровня — в w:lvl/w:pPr/w:ind.

Модуль собирает эту цепочку в одну структуру ParagraphFormat, которая:

  * в HTML-предпросмотре и при сборке тома/раздела в HTML (preview_html)
    превращается в CSS (отступ первой строки, абзацные отступы, интервалы,
    выравнивание) и в разметку списков <ul>/<ol>;
  * при сборке DOCX/PDF (docx_exporter) переносится в абзацы выходного
    документа.

Вертикальные интервалы абзацев (space_before / space_after) в HTML выражаются
через padding, а не margin: соседние margin-ы схлопываются, а в Word интервалы
соседних абзацев суммируются (важно, когда в документе есть пустые строки).

Шрифт и кегль здесь намеренно НЕ читаются: по требованиям проекта шрифт и кегль
текста задаёт ГОСТ-оформление (карта кеглей config.FontMap), а из Word берётся
только «геометрия» абзаца — отступы, интервалы, выравнивание и списки.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Единицы измерения Word.
TWIPS_PER_MM = 1440.0 / 25.4          # 1 мм = 56.6929 twips
PT_PER_MM = 72.0 / 25.4               # 1 мм = 2.8346 pt
MM_PER_TWIP = 25.4 / 1440.0

# «Одинарный» межстрочный интервал Word в терминах CSS line-height.
# Word хранит множитель относительно одинарного интервала (w:line/w:lineRule),
# а CSS line-height задаётся относительно кегля; для типовых шрифтов одинарный
# интервал Word ≈ 1.15 кегля.
WORD_SINGLE_LINE = 1.15

# Базовый межстрочный интервал ГОСТ-текста проекта: он же использовался до
# появления этого модуля. Применяется, когда Word интервал не задаёт.
DEFAULT_LINE_HEIGHT = 1.4

# Ширина маркера списка по умолчанию (мм) — стандартная «висячая» позиция Word
# (360 twips). Нужна, если уровень списка не задаёт собственных отступов.
DEFAULT_LIST_HANGING_MM = 360.0 * MM_PER_TWIP


# --------------------------------------------------------------------------
# Разбор значений
# --------------------------------------------------------------------------

def twips_to_mm(value) -> float:
    """Twips -> мм (0.0, если значение не число)."""
    try:
        return float(value) * MM_PER_TWIP
    except (TypeError, ValueError):
        return 0.0


def _on_off(el) -> bool:
    """Значение «включено/выключено» OOXML (w:b, w:hanging и т. п.)."""
    if el is None:
        return False
    val = el.get(W + "val")
    if val is None:
        return True
    return str(val).strip().lower() not in ("0", "false", "off", "none")


_ROMAN = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
          (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
          (5, "V"), (4, "IV"), (1, "I"))


def _roman(n: int) -> str:
    n = max(1, int(n))
    out = []
    for value, sym in _ROMAN:
        while n >= value:
            out.append(sym)
            n -= value
    return "".join(out)


def _alpha(n: int) -> str:
    """1 -> a, 26 -> z, 27 -> aa (как в Word)."""
    n = max(1, int(n))
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("a") + rem) + out
    return out


def format_number(number: int, num_fmt: str) -> str:
    """Строка номера списка для формата нумерации Word (w:numFmt)."""
    fmt = (num_fmt or "decimal").strip()
    low = fmt.lower()
    if low in ("decimal", "decimalenclosedcircle", "decimalfullwidth"):
        return str(int(number))
    if low == "decimalzero":
        return "%02d" % int(number)
    if low == "lowerletter":
        return _alpha(number)
    if low == "upperletter":
        return _alpha(number).upper()
    if low == "lowerroman":
        return _roman(number).lower()
    if low == "upperroman":
        return _roman(number)
    if low in ("russianlower", "russianupper"):
        # В OOXML такие форматы встречаются редко и обычно как обычный текст.
        text = _alpha(number)
        return text.upper() if low.endswith("upper") else text
    if low == "bullet":
        return "\u2022"
    if low == "none":
        return ""
    return str(int(number))


_LVL_PLACEHOLDER = re.compile(r"%\d")


def format_list_text(lvl_text: str, num_fmt: str, number: int) -> str:
    """Текст маркера списка: «1.», «а)», «•» — как в Word (w:lvlText)."""
    pattern = lvl_text or ""
    if num_fmt.lower() == "bullet" and not pattern:
        return "\u2022"
    if not pattern:
        pattern = "%1." if num_fmt.lower() != "bullet" else "\u2022"
    if num_fmt.lower() == "bullet":
        # Символ маркера задаётся шрифтом (Symbol/Wingdings): \uF0B7 и т. п.
        ch = pattern.replace("%1", "")
        return _normalize_bullet(ch)
    out = pattern
    if "%1" in out:
        out = out.replace("%1", format_number(number, num_fmt))
    # Остальные уровни (%2, %3...) в плоском выводе заменяем их номерами
    # неизвестны — оставляем как есть, но убираем служебный символ %.
    out = _LVL_PLACEHOLDER.sub("", out)
    return out.strip()


# Символы маркеров из «служебных» шрифтов Word -> обычные.
_BULLET_MAP = {
    "\uf0b7": "\u2022", "\uf0a7": "\u25aa", "\uf0d8": "\u25aa",
    "\uf0fc": "\u2713", "\uf0e0": "\u27a4", "\uf075": "\u25cf",
    "o": "\u25cb", "\uf06f": "\u25cb",
}


def _normalize_bullet(ch: str) -> str:
    """Приводит символ маркера к знаку, который отображается в HTML."""
    ch = (ch or "").strip()
    if not ch:
        return "\u2022"
    # В служебных шрифтах маркер — один символ из области Private Use.
    low = ch.lower()
    if low in _BULLET_MAP:
        return _BULLET_MAP[low]
    if len(ch) == 1 and 0xE000 <= ord(ch) <= 0xF8FF:
        return "\u2022"
    return ch


# --------------------------------------------------------------------------
# Свойства абзаца
# --------------------------------------------------------------------------

@dataclass
class ParagraphFormat:
    """Оформление абзаца Word, приведённое к миллиметрам и пунктам."""

    # Отступы внутри текстовой области, мм.
    left_mm: float = 0.0
    right_mm: float = 0.0
    # Отступ первой строки, мм: > 0 — «красная строка», < 0 — выступ.
    first_line_mm: float = 0.0
    # Интервалы, pt.
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    # Межстрочный интервал: множитель CSS line-height (line_spacing = None —
    # интервал не задан, используется DEFAULT_LINE_HEIGHT).
    line_spacing: float | None = None
    line_spacing_pt: float | None = None
    # Выравнивание: "" (по умолчанию) | left | center | right | justify.
    alignment: str = ""
    # Абзац начинается с новой страницы.
    page_break_before: bool = False

    # Список: "" | "bullet" | "number".
    list_kind: str = ""
    list_level: int = 0
    list_id: str = ""            # w:numId — один счётчик на список
    list_fmt: str = ""           # w:numFmt
    list_text: str = ""          # w:lvlText («%1.», «•»)
    list_jc: str = ""            # w:lvlJc — выравнивание маркера в его поле
    list_number: int = 0         # фактический номер абзаца в списке
    # Отступ маркера от левого края текстовой колонки, мм (позиция «висячего»
    # отступа Word: left + firstLine, где firstLine < 0).
    marker_mm: float = 0.0
    # Ширина поля маркера, мм (расстояние от маркера до текста).
    marker_width_mm: float = DEFAULT_LIST_HANGING_MM

    # ---- вычисляемые свойства -------------------------------------------

    @property
    def has_list(self) -> bool:
        return bool(self.list_kind)

    @property
    def leading_space_mm(self) -> float:
        """Отступ первой строки «в терминах CSS»: positive — вправо."""
        return self.first_line_mm

    @property
    def marker_text(self) -> str:
        """Готовый текст маркера («1.», «а)», «•»)."""
        if not self.list_kind:
            return ""
        return format_list_text(self.list_text, self.list_fmt, self.list_number)

    def css_line_height(self) -> str | None:
        """line-height для CSS (None — значение не задано)."""
        if self.line_spacing_pt:
            return "%gpt" % self.line_spacing_pt
        if self.line_spacing:
            return "%g" % self.line_spacing
        return None

    def line_height_factor(self) -> float:
        """Во сколько раз высота строки отличается от базовой (1.4)."""
        if self.line_spacing_pt:
            return max(0.2, float(self.line_spacing_pt) / (11.0 * DEFAULT_LINE_HEIGHT))
        value = self.line_spacing if self.line_spacing else DEFAULT_LINE_HEIGHT
        return max(0.2, float(value) / DEFAULT_LINE_HEIGHT)


class ParagraphFormats(dict):
    """{индекс абзаца: ParagraphFormat} + карта текущих заголовков.

    Значения по индексу абзаца в document.paragraphs — та же адресация, что
    используется для правок текста ПЗ (app.core.docx_edit).
    """

    def __init__(self, formats: dict | None = None, headings: dict | None = None):
        super().__init__(formats or {})
        # {индекс абзаца: текст последнего заголовка к этому месту документа}
        # — используется для колонтитула «текущий раздел» (как в Word).
        self.headings = headings or {}


# --------------------------------------------------------------------------
# Разбор OOXML
# --------------------------------------------------------------------------

def _parse_ind(ind) -> dict:
    """Отступы w:ind (twips и символьные единицы)."""
    if ind is None:
        return {}
    keys = ("left", "right", "firstLine", "hanging",
            "leftChars", "rightChars", "firstLineChars", "hangingChars",
            "start", "end")
    out = {}
    for key in keys:
        val = ind.get(W + key)
        if val is not None:
            out[key] = val
    return out


def _parse_spacing(sp) -> dict:
    """Интервалы w:spacing."""
    if sp is None:
        return {}
    out = {}
    for key in ("before", "after", "line", "beforeLines", "afterLines", "lineRule"):
        val = sp.get(W + key)
        if val is not None:
            out[key] = val
    return out


def _parse_pPr(pPr) -> dict:
    """Значимые для оформления свойства из элемента w:pPr."""
    out = {}
    if pPr is None:
        return out
    jc = pPr.find(W + "jc")
    if jc is not None:
        out["jc"] = (jc.get(W + "val") or "left").strip().lower()
    ind = _parse_ind(pPr.find(W + "ind"))
    if ind:
        out["ind"] = ind
    spacing = _parse_spacing(pPr.find(W + "spacing"))
    if spacing:
        out["spacing"] = spacing
    numPr = pPr.find(W + "numPr")
    if numPr is not None:
        num = {}
        num_id = numPr.find(W + "numId")
        if num_id is not None:
            num["numId"] = num_id.get(W + "val")
        ilvl = numPr.find(W + "ilvl")
        if ilvl is not None:
            num["ilvl"] = ilvl.get(W + "val")
        if num:
            out["numPr"] = num
    if pPr.find(W + "pageBreakBefore") is not None:
        out["pageBreakBefore"] = _on_off(pPr.find(W + "pageBreakBefore"))
    return out


def _merge_pPr(chain) -> dict:
    """Сливает цепочку pPr «от общего к частному» (последний побеждает)."""
    merged: dict = {}
    for pPr in chain:
        if not pPr:
            continue
        for key, value in pPr.items():
            if key in ("ind", "spacing"):
                current = dict(merged.get(key) or {})
                # Внутри w:ind / w:spacing затираются только указанные атрибуты.
                current.update(value)
                merged[key] = current
            elif key == "numPr":
                if str(value.get("numId", "")) == "0":
                    # numId=0 в Word означает «нумерацию отключить».
                    merged.pop("numPr", None)
                else:
                    current = dict(merged.get("numPr") or {})
                    current.update(value)
                    merged["numPr"] = current
            else:
                merged[key] = value
    return merged


def _indent_values(ind: dict, base_size_pt: float) -> tuple:
    """(left_mm, right_mm, first_line_mm) из собранных w:ind.

    Символьные единицы (w:leftChars и т. п.) переводятся в миллиметры по
    кеглю: 100 единиц = 1 знак.
    """
    def side_mm(name: str, chars_name: str) -> float:
        if name in ind:
            return twips_to_mm(ind[name])
        if chars_name in ind:
            try:
                chars = float(ind[chars_name]) / 100.0
            except (TypeError, ValueError):
                return 0.0
            return chars * base_size_pt * 0.5 * PT_PER_MM
        return 0.0

    left = side_mm("left", "leftChars")
    if not left:
        left = side_mm("start", "leftChars")
    right = side_mm("right", "rightChars")
    if not right:
        right = side_mm("end", "rightChars")

    first = 0.0
    if "hanging" in ind:
        first = -twips_to_mm(ind["hanging"])
    elif "hangingChars" in ind:
        try:
            first = -float(ind["hangingChars"]) / 100.0 * base_size_pt * 0.5 * PT_PER_MM
        except (TypeError, ValueError):
            first = 0.0
    elif "firstLine" in ind:
        first = twips_to_mm(ind["firstLine"])
    elif "firstLineChars" in ind:
        try:
            first = float(ind["firstLineChars"]) / 100.0 * base_size_pt * 0.5 * PT_PER_MM
        except (TypeError, ValueError):
            first = 0.0
    return left, right, first


_ALIGN_MAP = {"left": "left", "start": "left", "center": "center",
              "right": "right", "end": "right", "both": "justify",
              "justify": "justify", "distribute": "justify",
              "thaiDistribute": "justify"}


def _parse_numbering(root) -> tuple:
    """Разбирает word/numbering.xml -> (уровни abstractNum, карта numId)."""
    abstract: dict = {}
    for an in root.findall(W + "abstractNum"):
        aid = an.get(W + "abstractNumId")
        if aid is None:
            continue
        levels = {}
        for lvl in an.findall(W + "lvl"):
            levels[lvl.get(W + "ilvl") or "0"] = _parse_lvl(lvl)
        abstract[aid] = levels

    num_ids: dict = {}
    for num in root.findall(W + "num"):
        nid = num.get(W + "numId")
        if nid is None:
            continue
        aid_el = num.find(W + "abstractNumId")
        overrides = {}
        for ov in num.findall(W + "lvlOverride"):
            ilvl = ov.get(W + "ilvl") or "0"
            start = ov.find(W + "startOverride")
            lvl_el = ov.find(W + "lvl")
            if lvl_el is not None:
                overrides[ilvl] = _parse_lvl(lvl_el)
            elif start is not None:
                overrides[ilvl] = {"start": start.get(W + "val")}
        num_ids[nid] = {"abstract": aid_el.get(W + "val") if aid_el is not None else "",
                        "overrides": overrides}
    return abstract, num_ids


def _parse_lvl(lvl) -> dict:
    """Один уровень списка (w:lvl): формат, шаблон текста, отступы."""
    fmt = lvl.find(W + "numFmt")
    txt = lvl.find(W + "lvlText")
    start = lvl.find(W + "start")
    jc = lvl.find(W + "lvlJc")
    pPr = lvl.find(W + "pPr")
    ind = _parse_ind(pPr.find(W + "ind")) if pPr is not None else {}
    return {
        "fmt": (fmt.get(W + "val") if fmt is not None else "") or "",
        "text": (txt.get(W + "val") if txt is not None else "") or "",
        "start": (start.get(W + "val") if start is not None else "") or "",
        "jc": (jc.get(W + "val") if jc is not None else "") or "",
        "ind": ind,
    }


def _read_parts(path: str):
    """Читает styles/numbering/document из DOCX (без python-docx)."""
    import zipfile
    from xml.etree import ElementTree as ET

    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        root = ET.fromstring(z.read("word/document.xml"))
        styles_root = ET.fromstring(z.read("word/styles.xml")) if "word/styles.xml" in names else None
        numbering_root = (ET.fromstring(z.read("word/numbering.xml"))
                          if "word/numbering.xml" in names else None)
    return root, styles_root, numbering_root


def _parse_styles(styles_root) -> tuple:
    """(стили, умолчания документа) из word/styles.xml."""
    styles: dict = {}
    defaults: dict = {}
    if styles_root is None:
        return styles, defaults
    for st in styles_root.findall(W + "style"):
        sid = st.get(W + "styleId")
        if not sid:
            continue
        based = st.find(W + "basedOn")
        size_el = None
        rPr = st.find(W + "rPr")
        if rPr is not None:
            size_el = rPr.find(W + "sz")
        styles[sid] = {
            "basedOn": (based.get(W + "val") if based is not None else "") or "",
            "pPr": _parse_pPr(st.find(W + "pPr")),
            "size_half_pt": (size_el.get(W + "val") if size_el is not None else ""),
        }
    dd = styles_root.find(W + "docDefaults")
    if dd is not None:
        p_def = dd.find(W + "pPrDefault")
        if p_def is not None:
            defaults = _parse_pPr(p_def.find(W + "pPr"))
    return styles, defaults


def _style_chain(style_id: str, styles: dict) -> list:
    """Цепочка pPr стилей «от базового к производному» (с защитой от циклов)."""
    chain = []
    seen = set()
    current = style_id
    while current and current in styles and current not in seen:
        seen.add(current)
        chain.append(styles[current].get("pPr") or {})
        current = styles[current].get("basedOn") or ""
    chain.reverse()
    return chain


def _style_base_size(style_id: str, styles: dict, default_pt: float) -> float:
    """Кегль стиля абзаца (для перевода символьных отступов), если задан."""
    current = style_id
    seen = set()
    while current and current in styles and current not in seen:
        seen.add(current)
        size = styles[current].get("size_half_pt")
        if size:
            try:
                return float(size) / 2.0
            except (TypeError, ValueError):
                pass
        current = styles[current].get("basedOn") or ""
    return default_pt


# --------------------------------------------------------------------------
# Публичное чтение
# --------------------------------------------------------------------------

_CACHE: dict = {}
_CACHE_LIMIT = 64


def clear_cache() -> None:
    """Сбрасывает кэш разобранного оформления (после правки DOCX)."""
    _CACHE.clear()


def read_paragraph_formats(path: str, base_size_pt: float = 11.0) -> ParagraphFormats:
    """Оформление всех абзацев тела DOCX.

    :param path: путь к DOCX.
    :param base_size_pt: кегль по умолчанию — нужен только для перевода
        символьных отступов (w:leftChars и т. п.) в миллиметры.
    :return: ParagraphFormats ({индекс абзаца: ParagraphFormat} + .headings —
        текст текущего заголовка для каждого абзаца, для колонтитула).
        Пустой результат, если файл не читается.
    """
    if not path or not os.path.exists(path):
        return ParagraphFormats()
    try:
        stat = os.stat(path)
        key = (os.path.abspath(path), int(stat.st_mtime_ns), int(stat.st_size))
    except OSError:
        return ParagraphFormats()
    cached = _CACHE.get(key)
    if cached is not None:
        return cached

    try:
        result = _read_formats(path, base_size_pt)
    except Exception:
        result = ParagraphFormats()
    if len(_CACHE) >= _CACHE_LIMIT:
        _CACHE.clear()
    _CACHE[key] = result
    return result


def _read_formats(path: str, base_size_pt: float) -> ParagraphFormats:
    root, styles_root, numbering_root = _read_parts(path)
    styles, defaults = _parse_styles(styles_root)
    abstract, num_ids = ({}, {})
    if numbering_root is not None:
        abstract, num_ids = _parse_numbering(numbering_root)

    formats: dict = {}
    headings: dict = {}
    counters: dict = {}          # (numId, ilvl) -> текущий номер
    style_of: dict = {}
    is_heading: dict = {}
    last_heading = ""

    paragraphs = root.find(W + "body").findall(W + "p") if root.find(W + "body") is not None else []
    for index, p in enumerate(paragraphs):
        pPr_el = p.find(W + "pPr")
        pPr = _parse_pPr(pPr_el)
        p_style_el = pPr_el.find(W + "pStyle") if pPr_el is not None else None
        style_id = p_style_el.get(W + "val") if p_style_el is not None else ""
        style_of[index] = style_id

        chain = _style_chain(style_id, styles)
        merged = _merge_pPr([defaults] + chain + [pPr])

        size_pt = _style_base_size(style_id, styles, base_size_pt)
        fmt = ParagraphFormat()
        left, right, first = _indent_values(merged.get("ind") or {}, size_pt)
        fmt.left_mm = round(left, 3)
        fmt.right_mm = round(right, 3)
        fmt.first_line_mm = round(first, 3)

        spacing = merged.get("spacing") or {}
        fmt.space_before_pt = _space_pt(spacing, "before", "beforeLines", size_pt)
        fmt.space_after_pt = _space_pt(spacing, "after", "afterLines", size_pt)
        _apply_line_spacing(fmt, spacing)
        fmt.alignment = _ALIGN_MAP.get((merged.get("jc") or "").strip().lower(), "")
        fmt.page_break_before = bool(merged.get("pageBreakBefore"))

        num = merged.get("numPr") or {}
        num_id = str(num.get("numId") or "")
        if num_id and num_id != "0":
            level = _numbering_level(num_id, num.get("ilvl"), num_ids, abstract)
            if level is not None:
                kind = "bullet" if (level["fmt"] or "").lower() == "bullet" else "number"
                fmt.list_kind = kind
                fmt.list_id = num_id
                fmt.list_fmt = level["fmt"] or "decimal"
                fmt.list_text = level["text"]
                fmt.list_jc = level.get("jc") or ""
                try:
                    fmt.list_level = int(num.get("ilvl") or 0)
                except (TypeError, ValueError):
                    fmt.list_level = 0
                _apply_level_indent(fmt, level)
            else:
                # Список «отсутствует» в numbering.xml — параграф остаётся
                # обычным абзацем с отступами.
                fmt.list_kind = ""

        if fmt.has_list:
            key = (fmt.list_id, fmt.list_level)
            if key not in counters:
                counters[key] = _list_start(fmt, num_ids, abstract) - 1
            counters[key] += 1
            fmt.list_number = counters[key]

        # Отступ маркера и ширина его поля (для HTML/CSS).
        if fmt.has_list:
            fmt.marker_mm = round(fmt.left_mm + min(0.0, fmt.first_line_mm), 3)
            width = -fmt.first_line_mm if fmt.first_line_mm < 0 else 0.0
            fmt.marker_width_mm = round(width if width > 0.4 else DEFAULT_LIST_HANGING_MM, 3)

        formats[index] = fmt

        # --- текущий заголовок (для колонтитула «продолжение раздела») ----
        heading = _is_heading_style(style_id, styles) or _is_heading_by_styles(style_id, styles)
        is_heading[index] = heading
        if heading and _paragraph_text(p).strip():
            text = _paragraph_text(p).strip()
            if fmt.has_list and fmt.list_kind == "number":
                text = "%s %s" % (fmt.marker_text, text)
            last_heading = text
        headings[index] = last_heading

    return ParagraphFormats(formats, headings)


def _paragraph_text(p) -> str:
    """Текст абзаца из XML (w:t + переносы/табуляции как пробелы)."""
    parts = []
    for node in p.iter():
        tag = node.tag
        if tag == W + "t":
            parts.append(node.text or "")
        elif tag in (W + "tab", W + "br", W + "cr"):
            parts.append(" ")
    return "".join(parts)


def _is_heading_style(style_id: str, styles: dict) -> bool:
    """Абзац — заголовок? (по цепочке стилей: Heading / Заголовок)."""
    current = style_id
    seen = set()
    while current and current not in seen:
        seen.add(current)
        low = current.lower()
        if low.startswith("heading") or low.startswith("заголовок") or "heading" in low:
            return True
        style = styles.get(current)
        if not style:
            break
        current = style.get("basedOn") or ""
    return False


# Стили абзаца, которые помечаются как заголовок (проверяется по цепочке
# наследования стиля) — как их называет Word и ГОСТ-шаблоны.
_HEADING_STYLE_RE = re.compile(
    r"(heading|заголов|название\s*раздела|section\s*heading|title\b)", re.I)


def _is_heading_by_styles(style_id: str, styles: dict) -> bool:
    """Абзац — заголовок? Проверяются и стиль, и его базовые стили.

    Нужно как запасной путь: python-docx может не знать имя стиля (стиль не
    описан в части styles.xml так, как ожидает библиотека), тогда как в XML он
    есть — и заголовок в предпросмотре терял бы свой кегль.
    """
    current = style_id
    seen = set()
    while current and current not in seen:
        seen.add(current)
        if _HEADING_STYLE_RE.search(current):
            return True
        style = styles.get(current)
        if not style:
            break
        current = style.get("basedOn") or ""
    return False


def _space_pt(spacing: dict, key: str, char_key: str, size_pt: float) -> float:
    """Интервал до/после абзаца в pt (учитывает «строчные» единицы Word)."""
    if key in spacing:
        try:
            return round(float(spacing[key]) / 20.0, 2)     # twips -> pt
        except (TypeError, ValueError):
            return 0.0
    if char_key in spacing:
        try:
            return round(float(spacing[char_key]) / 100.0 * size_pt, 2)
        except (TypeError, ValueError):
            return 0.0
    return 0.0


def _apply_line_spacing(fmt: ParagraphFormat, spacing: dict) -> None:
    """Межстрочный интервал Word -> CSS line-height (множитель или pt)."""
    raw = spacing.get("line")
    if raw is None:
        return
    try:
        line = float(raw)
    except (TypeError, ValueError):
        return
    if line <= 0:
        return
    rule = (spacing.get("lineRule") or "auto").strip().lower()
    if rule in ("exact", "atleast"):
        fmt.line_spacing_pt = round(line / 20.0, 2)
    else:
        multiple = line / 240.0
        # Word: «одинарный» = 240/240. В CSS одинарный ≈ 1.15 кегля.
        fmt.line_spacing = round(multiple * WORD_SINGLE_LINE, 3)


def _numbering_level(num_id: str, ilvl, num_ids: dict, abstract: dict):
    """Определение уровня списка для пары (numId, ilvl) — с учётом override."""
    info = num_ids.get(num_id)
    if not info:
        return None
    level_index = str(ilvl if ilvl is not None else 0)
    override = (info.get("overrides") or {}).get(level_index)
    if override and override.get("fmt"):
        return override
    levels = abstract.get(info.get("abstract") or "") or {}
    level = levels.get(level_index)
    if level is None and levels:
        level = levels.get("0")
    return level


def _apply_level_indent(fmt: ParagraphFormat, level: dict) -> None:
    """Отступы уровня списка, если абзац/стиль их не переопределяет."""
    ind = level.get("ind") or {}
    if not ind:
        return
    left = twips_to_mm(ind.get("left")) if "left" in ind else None
    hanging = twips_to_mm(ind.get("hanging")) if "hanging" in ind else None
    first = twips_to_mm(ind.get("firstLine")) if "firstLine" in ind else None
    if not fmt.left_mm and left:
        fmt.left_mm = round(left, 3)
    if abs(fmt.first_line_mm) < 0.05:
        if hanging:
            fmt.first_line_mm = round(-hanging, 3)
        elif first:
            fmt.first_line_mm = round(first, 3)


def _list_start(fmt: ParagraphFormat, num_ids: dict, abstract: dict) -> int:
    """Начальный номер списка (w:start уровня или startOverride)."""
    level = _numbering_level(fmt.list_id, str(fmt.list_level), num_ids, abstract)
    if level and level.get("start"):
        try:
            return int(level["start"])
        except (TypeError, ValueError):
            return 1
    return 1


# --------------------------------------------------------------------------
# HTML
# --------------------------------------------------------------------------

def paragraph_css(fmt: ParagraphFormat | None, *, extra: str = "",
                  omit_text_indent: bool = False) -> str:
    """CSS абзаца (без шрифта и кегля) — отступы, интервалы, выравнивание.

    omit_text_indent=True — отступ первой строки не выводится. Нужно для пункта
    списка: «висячий» отступ Word там занят маркером, который в HTML рисуется
    отдельным абсолютным боксом, а сам текст пункта начинается на left_mm.
    """
    fmt = fmt or ParagraphFormat()
    parts = ["margin:0"]
    if fmt.alignment:
        parts.append("text-align:%s" % fmt.alignment)
    if fmt.left_mm:
        parts.append("margin-left:%smm" % _num(fmt.left_mm))
    if fmt.right_mm:
        parts.append("margin-right:%smm" % _num(fmt.right_mm))
    if fmt.first_line_mm and not omit_text_indent:
        parts.append("text-indent:%smm" % _num(fmt.first_line_mm))
    if fmt.space_before_pt:
        parts.append("padding-top:%spt" % _num(fmt.space_before_pt, 2))
    if fmt.space_after_pt:
        parts.append("padding-bottom:%spt" % _num(fmt.space_after_pt, 2))
    line = fmt.css_line_height()
    if line:
        parts.append("line-height:%s" % line)
    if extra:
        parts.append(extra)
    return ";".join(parts) + ";"


def marker_css(fmt: ParagraphFormat | None) -> str:
    """CSS бокса маркера списка.

    Маркер выносится в поле «висячего» отступа Word (w:hanging) абсолютным
    позиционированием относительно абзаца: текст пункта начинается на left_mm,
    маркер занимает полосу слева от него и выровнен по правому краю — точно так
    же, как Word ставит номер/маркер списка. Абсолютное позиционирование выбрано
    специально: маркер не попадает в текст абзаца при inline-правке в
    предпросмотре (иначе нумерация «затиралась» бы вместе с текстом).
    """
    fmt = fmt or ParagraphFormat()
    width = fmt.marker_width_mm or DEFAULT_LIST_HANGING_MM
    # Позиция маркера — начало поля «висячего» отступа Word (left + hanging)
    # относительно текстовой колонки, в которой стоит пункт списка.
    # Выравнивание маркера в его поле — как задано в Word (w:lvlJc); по
    # умолчанию «по левому краю»: так Word и ставит маркеры списков.
    align = {"right": "right", "center": "center"}.get(
        (fmt.list_jc or "").strip().lower(), "left")
    return ("position:absolute;left:%smm;width:%smm;text-align:%s;"
            "white-space:nowrap;" % (_num(fmt.marker_mm), _num(width), align))


def docx_indent_kwargs(fmt: ParagraphFormat | None) -> dict:
    """Отступы/интервалы абзаца для выходного DOCX (значения в мм и pt)."""
    fmt = fmt or ParagraphFormat()
    out = {
        "left_mm": fmt.left_mm,
        "right_mm": fmt.right_mm,
        "first_line_mm": fmt.first_line_mm,
        "space_before_pt": fmt.space_before_pt,
        "space_after_pt": fmt.space_after_pt,
    }
    if fmt.line_spacing_pt:
        out["line_spacing_pt"] = fmt.line_spacing_pt
    elif fmt.line_spacing:
        # Обратный перевод CSS-множителя в «множитель Word».
        out["line_spacing"] = round(float(fmt.line_spacing) / WORD_SINGLE_LINE, 3)
    if fmt.alignment:
        out["alignment"] = fmt.alignment
    return out


def _num(value: float, digits: int = 3) -> str:
    """Число без лишних нулей (для CSS/DOCX)."""
    text = ("%." + str(digits) + "f") % float(value)
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"
