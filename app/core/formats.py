# -*- coding: utf-8 -*-
"""Стандарты форматов листов по ГОСТ: размеры в мм и классификация."""
from __future__ import annotations

from dataclasses import dataclass

# Размеры листов по ГОСТ 2.301 (мм) — портретная ориентация (ширина x высота)
PAPER_SIZES_MM = {
    "A0": (841, 1189),
    "A1": (594, 841),
    "A2": (420, 594),
    "A3": (297, 420),
    "A4": (210, 297),
}

# Некоторые дополнительные нестандартные для графики кратные форматы
EXTRA_SIZES_MM = {
    "A0+": (841, 1682),
    "A1+": (594, 1261),
}

# Допуск при определении формата по размеру (мм)
TOLERANCE_MM = 3.0


@dataclass
class SheetFormat:
    """Определённый формат листа."""
    label: str          # A0..A4, или "Произвольный"
    width_mm: float
    height_mm: float
    portrait: bool
    standard: bool      # True — стандартный по ГОСТ, False — произвольный

    def __str__(self) -> str:
        orient = "портрет" if self.portrait else "альбом"
        return f"{self.label} ({self.width_mm:g}x{self.height_mm:g} мм, {orient})"


def _classify(w: float, h: float, tol: float = TOLERANCE_MM) -> SheetFormat | None:
    """Возвращает стандартный формат по максимальной из сторон, либо None."""
    # Пробуем портрет и альбом
    for label, (pw, ph) in PAPER_SIZES_MM.items():
        # портрет: ширина pw, высота ph ; альбом: ширина ph, высота pw
        if abs(w - pw) <= tol and abs(h - ph) <= tol:
            return SheetFormat(label, pw, ph, True, True)
        if abs(w - ph) <= tol and abs(h - pw) <= tol:
            return SheetFormat(label, pw, ph, False, True)
    return None


def classify_sheet(width_mm: float, height_mm: float) -> SheetFormat:
    """Классифицирует лист по физическим размерам (мм)."""
    st = _classify(width_mm, height_mm)
    if st:
        return st
    # нестандартный
    # определяем ориентацию: ширина больше высоты = альбом
    portrait = width_mm < height_mm
    return SheetFormat("Произвольный", width_mm, height_mm, portrait, False)
