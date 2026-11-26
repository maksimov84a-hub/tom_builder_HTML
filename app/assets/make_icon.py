# -*- coding: utf-8 -*-
"""Рисует иконку приложения (том по ГОСТ) и сохраняет app/assets/app.ico.

Иконка — собирательный образ документа: лист с рамкой и основной надписью
(полоса снизу) на синем «переплёте» тома. Рисуется в нескольких размерах,
поэтому остаётся читаемой и в панели задач (16-24 px), и в крупном превью.

Скрипт **не входит в рантайм**: он нужен только чтобы перерисовать иконку
(Pillow в vendor/ нет — запускать интерпретатором с установленным Pillow):

    python app/assets/make_icon.py
"""
import os

from PIL import Image, ImageDraw

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
OUT_DIR = os.path.join(ROOT, "app", "assets")
SIZES = (16, 24, 32, 48, 64, 128, 256)

#: Базовый размер отрисовки: дальше уменьшаем с высоким качеством.
BASE = 1024

INK = (28, 44, 68)          # тёмно-синий контур тома
COVER = (37, 99, 168)       # синий переплёт
COVER_DARK = (26, 74, 128)
PAPER = (255, 255, 255)
LINE = (52, 68, 90)
ACCENT = (214, 73, 58)      # красная линия рамки — «ГОСТ-овский» акцент


def draw_icon(size: int = BASE) -> Image.Image:
    """Иконка в заданном размере (рисуется в BASE и уменьшается)."""
    s = BASE
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    pad = int(s * 0.06)
    r = int(s * 0.20)

    # Переплёт тома (синяя подложка со скруглением).
    d.rounded_rectangle([pad, pad, s - pad, s - pad], radius=r, fill=COVER)
    # Тень-полутень справа снизу — объём без градиентов.
    d.rounded_rectangle(
        [pad + int(s * 0.02), pad + int(s * 0.03), s - pad, s - pad],
        radius=r, outline=COVER_DARK, width=max(2, int(s * 0.016)))

    # Лист бумаги.
    left = int(s * 0.26)
    top = int(s * 0.16)
    right = int(s * 0.86)
    bottom = int(s * 0.84)
    d.rounded_rectangle([left, top, right, bottom], radius=int(s * 0.035),
                        fill=PAPER, outline=INK, width=max(2, int(s * 0.014)))

    # Рамка листа (тонкая линия с отступом) — узнаваемый признак чертежа.
    fl, ft = int(s * 0.31), int(s * 0.22)
    fr, fb = int(s * 0.81), int(s * 0.78)
    d.rectangle([fl, ft, fr, fb], outline=ACCENT, width=max(2, int(s * 0.011)))

    # Строки текста: три полосы разной длины.
    lw = max(2, int(s * 0.030))
    for i, wf in enumerate((0.66, 0.52, 0.60)):
        y = int(s * (0.32 + i * 0.11))
        d.rounded_rectangle([fl + int(s * 0.04), y,
                             fl + int((fr - fl) * wf), y + lw],
                            radius=lw // 2, fill=LINE)

    # Основная надпись — полоса внизу рамки (штамп).
    stamp_top = int(s * 0.64)
    d.rectangle([fl, stamp_top, fr, fb], fill=PAPER, outline=INK,
                width=max(2, int(s * 0.011)))
    # Вертикали штампа.
    for xf in (0.42, 0.62):
        x = int(fl + (fr - fl) * xf)
        d.line([x, stamp_top, x, fb], fill=INK, width=max(1, int(s * 0.008)))
    # Горизонталь в штампе (строка «Лист / Листов»).
    ymid = int(stamp_top + (fb - stamp_top) * 0.45)
    d.line([fl, ymid, fr, ymid], fill=INK, width=max(1, int(s * 0.008)))

    return img


def main() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    base = draw_icon()
    frames = [base.resize((n, n), Image.LANCZOS) for n in SIZES]
    ico_path = os.path.join(OUT_DIR, "app.ico")
    frames[-1].save(ico_path, format="ICO",
                    sizes=[(n, n) for n in SIZES])

    # PNG 256×256 — для проверки и как запасной вариант (окна/трей).
    png_path = os.path.join(OUT_DIR, "app.png")
    frames[-1].save(png_path, format="PNG")
    print("сохранено:", ico_path, os.path.getsize(ico_path), "байт")
    print("сохранено:", png_path, os.path.getsize(png_path), "байт")


if __name__ == "__main__":
    main()
