# -*- coding: utf-8 -*-
"""
Конфигурация оформления: карта «расположение -> кегль» для шрифта ISOCPEUR,
полей листа по ГОСТ и прочих настроек по умолчанию.

Правила вынесены сюда (не зашиты в код), чтобы менять без правки логики.
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass, field, asdict


# Шрифт по решению проекта — используется во всех элементах
FONT_NAME = "ISOCPEUR"

# Поля листа по ГОСТ Р 21.101-2020 (стандарт): левое 20 мм, остальные 5 мм
@dataclass
class PageMargins:
    left: float = 20.0      # мм
    right: float = 5.0
    top: float = 5.0
    bottom: float = 5.0


# Формат основного листа для текстовой части
A4_MM = (210.0, 297.0)


# Карта «расположение -> кегль» (ISOCPEUR). Кегль в пунктах.
# Значения по умолчанию — пример; пользователь меняет в конфиге.
@dataclass
class FontMap:
    # (ключ, описание, кегль по умолчанию)
    title_doc_name:   int = 18   # титульный лист: наименование тома
    stamp_volume:     int = 14   # осн. надпись: обозначение тома
    stamp_qty:        int = 12   # осн. надпись: номер листа/листов
    stamp_columns:    int = 10   # осн. надпись: графы, ФИО, должности
    content_header:   int = 14   # содержание: заголовок
    content_row:      int = 12   # содержание: строки (разделы, приложения)
    toc_row:          int = 12   # оглавление ПЗ: строки (уровни текста из Word)
    text_body:        int = 14   # пояснительная записка: основной текст
    text_heading:     int = 14   # пояснительная записка: заголовки разделов
    appendix_header:  int = 12   # заголовок приложения / лист изменений
    signature:        int = 10   # должности / расшифровки подписей

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "FontMap":
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


# Описание локаций для UI (подсказки)
FONT_LOCATIONS = [
    ("title_doc_name",  "Титульный лист — наименование тома"),
    ("stamp_volume",    "Осн. надпись — обозначение тома"),
    ("stamp_qty",       "Осн. надпись — номер листа / листов"),
    ("stamp_columns",   "Осн. надпись — графы, ФИО, должности"),
    ("content_header",  "Содержание — заголовок"),
    ("content_row",     "Содержание — строки (разделы, приложения)"),
    ("toc_row",         "Оглавление ПЗ — строки по уровням заголовков"),
    ("text_body",       "Пояснительная записка — основной текст"),
    ("text_heading",    "Пояснительная записка — заголовки разделов"),
    ("appendix_header", "Заголовок приложения / лист изменений"),
    ("signature",       "Должности / расшифровки подписей"),
]


# --------------------------------------------------------------------------
# Шрифты текста пояснительной записки: обычный текст / заголовки / таблицы
# --------------------------------------------------------------------------

#: Группы текста ПЗ, для которых задаётся свой шрифт и кегль.
#: Ключ -> подпись в настройках.
TEXT_FONT_GROUPS = {
    "body": "Обычный текст",
    "heading": "Заголовки",
    "table": "Таблицы",
}

#: Шрифт по умолчанию для всех групп текста ПЗ.
DEFAULT_TEXT_FONT = FONT_NAME

#: Кегли по умолчанию для групп текста ПЗ (pt).
DEFAULT_TEXT_SIZES = {"body": 14, "heading": 14, "table": 10}


@dataclass
class TextFonts:
    """Шрифт и кегль текста ПЗ по группам: обычный текст, заголовки, таблицы.

    Один источник для предпросмотра, Word и PDF: раньше кегли были зашиты в
    коде (11/13 pt в предпросмотре, 14 pt в Word), из-за чего число листов в
    предпросмотре и в собранном томе расходилось.
    """

    body_font: str = DEFAULT_TEXT_FONT
    body_size: int = DEFAULT_TEXT_SIZES["body"]
    heading_font: str = DEFAULT_TEXT_FONT
    heading_size: int = DEFAULT_TEXT_SIZES["heading"]
    table_font: str = DEFAULT_TEXT_FONT
    table_size: int = DEFAULT_TEXT_SIZES["table"]

    def font(self, group: str) -> str:
        """Шрифт группы ('body' | 'heading' | 'table')."""
        return {"body": self.body_font, "heading": self.heading_font,
                "table": self.table_font}.get(group, self.body_font)

    def size(self, group: str) -> int:
        """Кегль группы ('body' | 'heading' | 'table')."""
        return {"body": self.body_size, "heading": self.heading_size,
                "table": self.table_size}.get(group, self.body_size)

    def set(self, group: str, font: str | None = None,
            size: int | None = None) -> None:
        """Записывает шрифт и/или кегль группы."""
        if group not in TEXT_FONT_GROUPS:
            return
        if font is not None:
            setattr(self, f"{group}_font", str(font))
        if size is not None:
            try:
                setattr(self, f"{group}_size", max(4, int(size)))
            except (TypeError, ValueError):
                pass

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "TextFonts":
        tf = cls()
        for group in TEXT_FONT_GROUPS:
            font = data.get(f"{group}_font")
            if font:
                setattr(tf, f"{group}_font", str(font))
            try:
                size = int(data.get(f"{group}_size", tf.size(group)))
                setattr(tf, f"{group}_size", max(4, size))
            except (TypeError, ValueError):
                pass
        return tf

CONFIG_FILENAME = "config.json"

#: Допустимые форматы сборки тома (чекбоксы меню «Сборка»).
BUILD_FORMAT_KEYS = ("pdf", "docx", "html")

#: Абзацный отступ («красная строка») в знаках. Один знак Word — половина
#: кегля, поэтому 5 знаков при кегле текста ПЗ 14 pt дают
#: 5 × 14 × 0.5 = 35 pt = 12.35 мм (ГОСТ: 1.25 см).
DEFAULT_INDENT_CHARS = 5.0

#: Пунктов в миллиметре (для перевода кегля в мм).
PT_PER_MM = 72.0 / 25.4


def indent_mm_for_chars(chars: float, size_pt: float) -> float:
    """Абзацный отступ в мм по числу знаков и кеглю (1 знак = 0.5 кегля)."""
    try:
        chars = float(chars)
        size_pt = float(size_pt)
    except (TypeError, ValueError):
        return 0.0
    if chars <= 0 or size_pt <= 0:
        return 0.0
    return chars * size_pt * 0.5 / PT_PER_MM


@dataclass
class Config:
    font_name: str = FONT_NAME
    margins: PageMargins = field(default_factory=PageMargins)
    font_map: FontMap = field(default_factory=FontMap)
    #: Автоматическое оглавление в начале пояснительной записки: строится по
    #: уровням заголовков исходного DOCX (app.core.pz_toc) и вставляется первым
    #: листом ПЗ в предпросмотр, Word и PDF. Отключается здесь, без правки кода.
    toc_enabled: bool = True
    #: Абзацный отступ текста ПЗ в знаках (5 знаков ≈ 12.35 мм при 14 pt).
    #: Применяется к абзацам и спискам, у которых Word отступ не задаёт.
    paragraph_indent_chars: float = DEFAULT_INDENT_CHARS
    #: Шрифт и кегль текста ПЗ по группам: обычный текст, заголовки, таблицы.
    #: Единый источник для предпросмотра, Word и PDF.
    text_fonts: TextFonts = field(default_factory=TextFonts)
    #: Форматы сборки тома, отмеченные чекбоксами в меню «Сборка»:
    #: "pdf" / "docx" / "html". Собираются только отмеченные.
    build_formats: list = field(default_factory=lambda: ["pdf", "docx", "html"])
    #: DOCX собирать по разделам тома — отдельный файл на раздел
    #: (app.core.docx_exporter.build_chapters_docx) вместо одного документа.
    build_docx_chapters: bool = False
    #: Панель основной надписи внизу окна развёрнута (True) или свёрнута в
    #: узкую полосу (False). Состояние запоминается между запусками.
    stamp_panel_open: bool = True

    def paragraph_indent_mm(self) -> float:
        """Абзацный отступ в мм по кеглю основного текста ПЗ."""
        return round(indent_mm_for_chars(self.paragraph_indent_chars,
                                         self.text_fonts.body_size), 3)

    def to_dict(self) -> dict:
        return {
            "font_name": self.font_name,
            "margins": asdict(self.margins),
            "font_map": self.font_map.to_dict(),
            "toc_enabled": bool(self.toc_enabled),
            "paragraph_indent_chars": float(self.paragraph_indent_chars),
            "text_fonts": self.text_fonts.to_dict(),
            "build_formats": list(self.build_formats),
            "build_docx_chapters": bool(self.build_docx_chapters),
            "stamp_panel_open": bool(self.stamp_panel_open),
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Config":
        cfg = cls()
        cfg.font_name = data.get("font_name", cfg.font_name)
        cfg.margins = PageMargins(**{k: v for k, v in (data.get("margins") or {}).items() if k in PageMargins.__dataclass_fields__})
        cfg.font_map = FontMap.from_dict(data.get("font_map") or {})
        cfg.toc_enabled = bool(data.get("toc_enabled", cfg.toc_enabled))
        try:
            cfg.paragraph_indent_chars = float(
                data.get("paragraph_indent_chars", cfg.paragraph_indent_chars))
        except (TypeError, ValueError):
            pass
        cfg.text_fonts = TextFonts.from_dict(data.get("text_fonts") or {})
        formats = [str(f) for f in (data.get("build_formats") or [])
                   if str(f) in BUILD_FORMAT_KEYS]
        # Пустой набор означает «все форматы»: иначе после сбоя файла
        # настроек сборка не произвела бы ничего.
        cfg.build_formats = formats or list(BUILD_FORMAT_KEYS)
        cfg.build_docx_chapters = bool(
            data.get("build_docx_chapters", cfg.build_docx_chapters))
        cfg.stamp_panel_open = bool(data.get("stamp_panel_open",
                                             cfg.stamp_panel_open))
        return cfg


def load_config(path: str | None = None) -> Config:
    """Загрузка конфигурации; если файла нет — возвращает значения по умолчанию."""
    path = path or os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), CONFIG_FILENAME)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return Config.from_dict(json.load(f))
        except Exception:
            pass
    return Config()


def save_config(cfg: Config, path: str | None = None) -> None:
    path = path or os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), CONFIG_FILENAME)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg.to_dict(), f, ensure_ascii=False, indent=2)
