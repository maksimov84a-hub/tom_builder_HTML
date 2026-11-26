# -*- coding: utf-8 -*-
"""Подключение шрифтов (ISOCPEUR по ГОСТ и системных) к PyMuPDF.

PyMuPDF встраивает в PDF ФАЙЛ шрифта, а не имя семейства, поэтому выбор
шрифта в настройках оформления переводится здесь в путь к TTF/OTF через
реестр установленных шрифтов Windows.
"""
from __future__ import annotations

import os

from ..config import FONT_NAME

#: Системная папка шрифтов
_win = os.environ.get("WINDIR", "C:\\Windows")
_FONTS_DIR = os.path.join(_win, "Fonts")

#: Путь к шрифту ISOCPEUR (основной шрифт проекта по ГОСТ)
ISOCPEUR_TTF = os.path.join(_FONTS_DIR, "isocpeur.ttf")

#: Индекс установленных шрифтов: имя семейства (нижний регистр) -> путь к файлу.
_FONT_INDEX: dict | None = None

#: Ключ реестра со списком установленных шрифтов.
_FONT_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"


def _font_index() -> dict:
    """Индекс установленных шрифтов: имя семейства -> путь к файлу.

    Читается из реестра Windows: имя семейства лежит в имени значения
    («ISOCPEUR (TrueType)»), имя файла — в самом значении. Без индекса
    выбранный в настройках шрифт не действовал бы в запасном PDF.
    """
    global _FONT_INDEX
    if _FONT_INDEX is not None:
        return _FONT_INDEX
    index: dict = {}
    try:
        import winreg
    except ImportError:
        winreg = None
    if winreg is not None:
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, _FONT_KEY) as key:
                    count = winreg.QueryInfoKey(key)[1]
                    for i in range(count):
                        try:
                            name, value, _type = winreg.EnumValue(key, i)
                        except OSError:
                            break
                        if not isinstance(value, str) or not name:
                            continue
                        path = (value if os.path.isabs(value)
                                else os.path.join(_FONTS_DIR, value))
                        if not os.path.exists(path):
                            continue
                        family = name.split("(")[0].strip().lower()
                        # Базовое начертание важнее курсива/полужирного: первым
                        # в реестре может идти «ISOCPEUR Italic», и текст
                        # выводился бы курсивом.
                        if family and ("italic" in name.lower()
                                       or "bold" in name.lower()):
                            index.setdefault(family, path)
                        else:
                            index[family] = path
            except OSError:
                continue
    _FONT_INDEX = index
    return index


def font_file(family: str | None = None) -> str:
    """Путь к файлу шрифта по имени семейства.

    Без имени (или если шрифта нет в системе) — ISOCPEUR; если недоступен и он,
    возвращается пустая строка (вызывающий код откатывается на встроенный
    шрифт PyMuPDF).
    """
    name = (family or "").strip()
    if name:
        path = _font_index().get(name.lower())
        if path:
            return path
        # Шрифт мог быть установлен без записи в реестр — ищем файл по имени.
        compact = name.replace(" ", "")
        for ext in (".ttf", ".otf", ".ttc"):
            candidate = os.path.join(_FONTS_DIR, compact + ext)
            if os.path.exists(candidate):
                return candidate
    return ISOCPEUR_TTF if os.path.exists(ISOCPEUR_TTF) else ""


def register_font(page, fontname: str = "ISOCPEUR",
                  family: str | None = None) -> str | None:
    """Регистрирует шрифт на странице PyMuPDF и возвращает имя для insert_text.

    :param family: имя семейства из настроек; None — шрифт ГОСТ ISOCPEUR.
    Возвращает имя, под которым шрифт доступен; None если недоступен.
    """
    path = ISOCPEUR_TTF if family is None else font_file(family)
    if not path or not os.path.exists(path):
        return None
    try:
        page.insert_font(fontname=fontname, fontfile=path)
        return fontname
    except Exception:
        return None
