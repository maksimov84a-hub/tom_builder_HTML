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

CONFIG_FILENAME = "config.json"


@dataclass
class Config:
    font_name: str = FONT_NAME
    margins: PageMargins = field(default_factory=PageMargins)
    font_map: FontMap = field(default_factory=FontMap)
    #: Автоматическое оглавление в начале пояснительной записки: строится по
    #: уровням заголовков исходного DOCX (app.core.pz_toc) и вставляется первым
    #: листом ПЗ в предпросмотр, Word и PDF. Отключается здесь, без правки кода.
    toc_enabled: bool = True

    def to_dict(self) -> dict:
        return {
            "font_name": self.font_name,
            "margins": asdict(self.margins),
            "font_map": self.font_map.to_dict(),
            "toc_enabled": bool(self.toc_enabled),
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Config":
        cfg = cls()
        cfg.font_name = data.get("font_name", cfg.font_name)
        cfg.margins = PageMargins(**{k: v for k, v in (data.get("margins") or {}).items() if k in PageMargins.__dataclass_fields__})
        cfg.font_map = FontMap.from_dict(data.get("font_map") or {})
        cfg.toc_enabled = bool(data.get("toc_enabled", cfg.toc_enabled))
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
