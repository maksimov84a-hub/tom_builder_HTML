# -*- coding: utf-8 -*-
"""Подключение шрифта ISOCPEUR (TrueType из Windows) к PyMuPDF."""
from __future__ import annotations

import os

from ..config import FONT_NAME

#: Путь к шрифту ISOCPEUR в системной папке Windows
_win = os.environ.get("WINDIR", "C:\\Windows")
ISOCPEUR_TTF = os.path.join(_win, "Fonts", "isocpeur.ttf")


def font_file() -> str:
    """Полный путь к TTF-файлу ISOCPEUR."""
    return ISOCPEUR_TTF if os.path.exists(ISOCPEUR_TTF) else ""


def register_font(page, fontname: str = "ISOCPEUR") -> str | None:
    """Регистрирует ISOCPEUR на странице PyMuPDF и возвращает имя шрифта.

    Возвращает имя, под которым шрифт доступен; None если недоступен.
    """
    try:
        if not os.path.exists(ISOCPEUR_TTF):
            return None
        page.insert_font(fontname=fontname, fontfile=ISOCPEUR_TTF)
        return fontname
    except Exception:
        return None
