# -*- coding: utf-8 -*-
"""Отрисовка основных надписей формы 3 / 5 / 6 в Word — по одной на лист A4.

Инструмент для доводки геометрии: собирает DOCX, где на каждом листе стоит
рамка ГОСТ (20 мм слева, по 5 мм с других сторон) и основная надпись своей
формы, затем (если установлен Word) конвертирует файл в PDF, растеризует
страницы в PNG и делает крупный план штампа. Заодно печатает координаты линий
в системе штампа (мм от его левого-верхнего угла) и сравнивает их с сеткой
app.core.geometry и с образцом из корня проекта («Форма 3.pdf» / «Форма 6.pdf»).

Геометрия не зашита здесь: сетки берутся из app.core.geometry (grid(3),
grid(5), grid(6)) — того же источника, что у HTML-предпросмотра и PDF.
Правишь geometry.py, перезапускаешь скрипт, сравниваешь замер.

Запуск:    py -3.12 tools/stamp_forms_docx.py [--form 6|3|5|356]
           (по умолчанию — форма 6, последующий лист текста)
Результат: tmp/форма_6.docx (+ .pdf, .png и крупный план штампа)
"""
from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
VENDOR = os.path.join(ROOT, "vendor")
if os.path.isdir(VENDOR) and VENDOR not in sys.path:
    sys.path.insert(0, VENDOR)

from docx import Document                                    # noqa: E402
from docx.enum.section import WD_SECTION                      # noqa: E402

from app.config import Config, A4_MM                          # noqa: E402
from app.core import docx_exporter as dx                      # noqa: E402
from app.core.gost import (FORM_GRAPHIC, FORM_TEXT_FIRST,     # noqa: E402
                           FORM_TEXT_NEXT)
from app.models import (VolumeProject, Requisites, Stage,     # noqa: E402
                        SECTION_GRAPHIC, SECTION_CONTENT)

#: Что рисуем: форма ГОСТ -> (ключ формы для docx_exporter, раздел подписей,
#: файл-образец в корне проекта).
FORMS = {
    3: (FORM_GRAPHIC, SECTION_GRAPHIC, "Форма 3.pdf"),
    5: (FORM_TEXT_FIRST, SECTION_CONTENT, "Форма 5.pdf"),
    6: (FORM_TEXT_NEXT, SECTION_CONTENT, "Форма 6.pdf"),
}

#: Название формы для подписи на листе и для имени файла.
FORM_NAMES = {3: "ФОРМА 3 — ГРАФИЧЕСКИЙ ЛИСТ",
              5: "ФОРМА 5 — ПЕРВЫЙ ЛИСТ ТЕКСТА",
              6: "ФОРМА 6 — ПОСЛЕДУЮЩИЙ ЛИСТ ТЕКСТА"}


def demo_project() -> VolumeProject:
    """Проект с заполненными реквизитами — чтобы графы штампа были не пустыми."""
    project = VolumeProject()
    project.requisites = Requisites(
        organization="ООО «ТверьУниверсалПроект»",
        volume_id="0325-ТХ",
        object_name="Реконструкция школы в д. Мокшино",
        volume_name="Пояснительная записка",
        stage=Stage.P,
        city="Тверь",
        year=2026,
    )
    project.signature_assignments = {"develop": "Максимов",
                                     "norm": "Сидоров",
                                     "approve": "Абрамов"}
    return project


def build(forms: list, out_path: str, config: Config) -> str:
    """Собирает DOCX: по листу A4 на каждую форму, с рамкой и штампом."""
    project = demo_project()
    fm = config.font_map
    font = config.font_name

    doc = Document()
    for index, form in enumerate(forms):
        if index:
            sec = doc.add_section(WD_SECTION.NEW_PAGE)
        else:
            sec = doc.sections[0]
        dx._set_section(A4_MM[0], A4_MM[1], config.margins, sec)
        # Подпись формы на листе — чтобы при просмотре не путать страницы.
        paragraph = doc.add_paragraph()
        dx._set_run(paragraph.add_run(FORM_NAMES.get(form, f"ФОРМА {form}")),
                    font, fm.content_header, bold=True)
        default_form, section_kind, _ref = FORMS[form]
        # Штамп и рамка листа: first_form=None — лист без «первого листа»
        # (на листе одна надпись своей формы).
        dx._add_stamp_footers(doc, project, config, sec, first_form=None,
                             default_form=default_form,
                             section_kind=section_kind)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    return dx._save_docx(doc, out_path)


def to_mm(pt: float) -> float:
    """Пункты PDF -> миллиметры."""
    return pt * 25.4 / 72.0


def to_pt(mm: float) -> float:
    """Миллиметры -> пункты PDF."""
    return mm * 72.0 / 25.4


def page_lines(page) -> tuple:
    """Собирает линии страницы PDF в мм: (вертикали, горизонтали).

    Word экспортирует границы ячеек заливками: каждый штрих — тонкий
    прямоугольник (0.25 мм), поэтому линия распознаётся по «длинной» стороне
    прямоугольника. Элементы — отрезки ``(координата, начало, конец)``:
    по ним видно, на каком участке штампа линия проходит (например, линия
    7 мм в форме 6 есть только в графе «Лист»).
    """
    horizontals, verticals = [], []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] != "re":
                continue
            rect = item[1]
            width, height = to_mm(rect.width), to_mm(rect.height)
            if width > 2 and height < 0.6:
                horizontals.append((to_mm((rect.y0 + rect.y1) / 2),
                                    to_mm(rect.x0), to_mm(rect.x1)))
            elif height > 2 and width < 0.6:
                verticals.append((to_mm((rect.x0 + rect.x1) / 2),
                                  to_mm(rect.y0), to_mm(rect.y1)))
    return verticals, horizontals


def cluster_segments(segments, tol: float = 0.4) -> list:
    """Группирует отрезки по координате: [(координата, (начало, конец), ...)].

    Возвращает [(координата, [(начало, конец), ...]), ...] по возрастанию —
    так видно протяжённость каждой линии сетки.
    """
    groups = []
    for coord, start, end in sorted(segments):
        if groups and abs(coord - groups[-1][0][-1]) <= tol:
            groups[-1][0].append(coord)
            groups[-1][1].append((start, end))
        else:
            groups.append(([coord], [(start, end)]))
    return [(round(sum(coords) / len(coords), 2), spans)
            for coords, spans in groups]


def cluster(values, tol: float = 0.4) -> list:
    """Склеивает близкие координаты (двойные линии Word) в одну."""
    out = []
    for value in sorted(values):
        if out and abs(value - out[-1][-1]) <= tol:
            out[-1].append(value)
        else:
            out.append([value])
    return [round(sum(group) / len(group), 2) for group in out]


def _coords(segments) -> list:
    """Только координаты линий (без протяжённости) по возрастанию, мм."""
    return [coord for coord, _spans in cluster_segments(segments)]


def stamp_frame(verticals: list, horizontals: list, page_h_mm: float) -> tuple:
    """Границы штампа на листе: (x_лево, y_верх, x_право, y_низ) в мм.

    Штамп стоит в нижнем колонтитуле: его линии — самая нижняя группа
    горизонталей, крайние вертикали листа — его боковые границы.
    """
    low = [y for y in horizontals if y > page_h_mm * 0.7]
    if not low or not verticals:
        return (0.0, 0.0, 0.0, 0.0)
    return (min(verticals), min(low), max(verticals), max(low))


def ref_grid(pdf_path: str) -> tuple:
    """Сетка образца из корня проекта: (вертикали, горизонтали) в мм.

    Образцы («Форма 3.pdf», «Форма 6.pdf») — фрагмент чертежа в масштабе 1:1,
    может быть повёрнут на 90°: у формы 6 фигура вытянута по вертикали листа.
    Поэтому направление линий в PDF не совпадает с направлением в форме:
    линии, идущие в PDF вдоль длинной стороны фигуры, — горизонтали формы,
    поперёк — её вертикали. Возвращается сетка в системе штампа: от левого
    края и от верха фигуры.

    Кромки фигуры в образцах нарисованы не отрезками, а замкнутым контуром,
    поэтому в результат попадают только внутренние линии сетки.
    """
    import pymupdf
    with pymupdf.open(pdf_path) as pdf:
        page = pdf[0]
        segs = []
        for drawing in page.get_drawings():
            for item in drawing["items"]:
                if item[0] != "l":
                    continue
                x0, y0 = to_mm(item[1].x), to_mm(item[1].y)
                x1, y1 = to_mm(item[2].x), to_mm(item[2].y)
                segs.append((x0, y0, x1, y1))
        # Направление оси Y (от верха штампа) — по заголовку «Лист»: он стоит в
        # верхней полосе (до 7 мм), поэтому начало оси у той кромки, к которой
        # заголовок ближе. Без этого координаты считались от противоположного
        # края и линия 7 мм «читалась» как 8 мм.
        head = _ref_head_coord(page)
    if not segs:
        return ([], [])
    xs = [v for s in segs for v in (s[0], s[2])]
    ys = [v for s in segs for v in (s[1], s[3])]
    long_is_y = (max(ys) - min(ys)) >= (max(xs) - min(xs))
    near_min = True
    if head is not None:
        coords = xs if long_is_y else ys
        near_min = abs(head - min(coords)) <= abs(head - max(coords))
    verticals, horizontals = [], []
    for x0, y0, x1, y1 in segs:
        along_pdf_x = abs(y1 - y0) < 0.3      # отрезок горизонтален в PDF
        along_pdf_y = abs(x1 - x0) < 0.3      # отрезок вертикален в PDF
        cross_min = min(xs) if long_is_y else min(ys)
        cross_max = max(xs) if long_is_y else max(ys)
        if long_is_y:
            if along_pdf_y:
                coord = (x0 + x1) / 2
                horizontals.append(coord - cross_min if near_min
                                   else cross_max - coord)
            elif along_pdf_x:
                verticals.append((y0 + y1) / 2 - min(ys))
        else:
            if along_pdf_x:
                coord = (y0 + y1) / 2
                horizontals.append(coord - cross_min if near_min
                                   else cross_max - coord)
            elif along_pdf_y:
                verticals.append((x0 + x1) / 2 - min(xs))
    long_span = (max(ys) - min(ys)) if long_is_y else (max(xs) - min(xs))
    short_span = (max(xs) - min(xs)) if long_is_y else (max(ys) - min(ys))
    # Кромки фигуры нарисованы замкнутым контуром (не отрезками) — добавляем
    # их из габаритов, иначе они выглядели бы «пропущенными» в сравнении.
    return (cluster(verticals + [0.0, long_span]),
            cluster(horizontals + [0.0, short_span]))


#: Заголовки графы «Лист» в образцах: по ним определяется верх штампа.
_REF_HEAD_TEXTS = ("Лист", "Листов")


def _ref_head_coord(page):
    """Поперечная координата (мм) заголовка «Лист» в образце или None.

    Образец может быть повёрнут на 90°, поэтому координата берётся по
    поперечной оси фигуры — той же, по которой отсчитывается высота штампа.
    """
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                if span["text"].strip() not in _REF_HEAD_TEXTS:
                    continue
                x0, y0, x1, y1 = span["bbox"]
                return to_mm((x0 + x1) / 2) if abs(y1 - y0) > abs(x1 - x0) \
                    else to_mm((y0 + y1) / 2)
    return None


def render(docx_path: str, forms: list, dpi: int = 150, zoom_dpi: int = 400) -> tuple:
    """DOCX -> PDF (Word COM) -> PNG страниц + крупный план штампа.

    Файлы называются по номеру формы листа (форма_6.png, форма_6_штамп.png),
    чтобы прогон одной формы не затирал снимки другой.
    Возвращает (pdf, [png страниц], [png крупным планом]) или ('', [], []).
    """
    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    try:
        from app.core.word_com import docx_to_pdf
        docx_to_pdf(docx_path, pdf_path)
    except Exception as exc:
        print(f"PDF не собран (Word недоступен): {exc}")
        return ("", [], [])
    shots, zooms = [], []
    try:
        import pymupdf
        with pymupdf.open(pdf_path) as pdf:
            for number, page in enumerate(pdf, start=1):
                page.set_rotation(0)
                form = forms[number - 1] if number <= len(forms) else number
                base = os.path.join(os.path.dirname(docx_path), f"форма_{form}")
                shot = base + ".png"
                page.get_pixmap(dpi=dpi).save(shot)
                shots.append(shot)
                # Крупный план штампа: по замеренным линиям, с полями 2 мм.
                verticals, horizontals = page_lines(page)
                left, top, right, bottom = stamp_frame(
                    _coords(verticals), _coords(horizontals),
                    to_mm(page.rect.height))
                if right > left and bottom > top:
                    clip = pymupdf.Rect(
                        to_pt(max(0.0, left - 2)), to_pt(max(0.0, top - 2)),
                        to_pt(min(right + 2, to_mm(page.rect.width))),
                        to_pt(min(bottom + 2, to_mm(page.rect.height))))
                    zoom = base + "_штамп.png"
                    page.get_pixmap(dpi=zoom_dpi, clip=clip).save(zoom)
                    zooms.append(zoom)
    except Exception as exc:
        print(f"PNG не собран: {exc}")
    return (pdf_path, shots, zooms)


def measure(pdf_path: str, forms: list) -> list:
    """Печатает координаты линий штампа и шаги сетки, сверяет с geometry.py.

    Абсолютные координаты считаются от левого края штампа (первая вертикаль).
    Крайние линии штампа совпадают с рамкой листа, а рамка при экспорте
    Word→PDF смещена наружу на ~0.15 мм, поэтому точнее сравнивать ШАГИ между
    линиями — они не зависят от положения рамки.

    Возвращает замеры по листам: [(вертикали, горизонтали), ...] в мм.
    """
    import pymupdf
    from app.core.geometry import grid
    out = []
    with pymupdf.open(pdf_path) as pdf:
        for number, page in enumerate(pdf, start=1):
            vsegs, hsegs = page_lines(page)
            if not vsegs or not hsegs:
                print(f"\n--- лист {number}: линии не найдены ---")
                out.append(([], []))
                continue
            vgroups, hgroups = cluster_segments(vsegs), cluster_segments(hsegs)
            page_h = to_mm(page.rect.height)
            v = [coord for coord, _ in vgroups]
            left, top = v[0], min(coord for coord in _coords(hsegs)
                                  if coord > page_h * 0.7)
            rel_v = [round(x - left, 2) for x in v]
            rel_h = [round(y - top, 2) for y in _coords(hsegs)
                     if y > page_h * 0.7]
            out.append((rel_v, rel_h))
            print(f"\n--- лист {number} ({to_mm(page.rect.width):.1f} x "
                  f"{page_h:.1f} мм) ---")
            print(f"штамп: левый край {left:.2f} мм, верх {top:.2f} мм, "
                  f"высота {max(rel_h):.2f} мм")
            print("вертикали (от левого края штампа):", rel_v)
            print("горизонтали (от верха штампа):     ", rel_h)
            print("шаги по X:", _steps(rel_v), "\nшаги по Y:", _steps(rel_h))
            # Протяжённость горизонталей: видно, на каком участке идёт линия
            # (у формы 6 линия 7 мм должна быть только в графе «Лист»).
            print("участки горизонталей по X:")
            for coord, spans in hgroups:
                if coord < top - 0.5:
                    continue
                parts = ", ".join(f"{a - left:.0f}..{b - left:.0f}"
                                  for a, b in spans)
                print(f"   y={coord - top:5.2f}: {parts}")
            if number <= len(forms):
                g = grid(forms[number - 1])
                exp_v, exp_h = _steps(g["vlines"]), _steps(g["hlines"])
                print("сетка geometry.py: шаги по X", exp_v,
                      "\n                   шаги по Y", exp_h)
                print("расхождение: шаги по X", _diff(_steps(rel_v), exp_v),
                      "шаги по Y", _diff(_steps(rel_h), exp_h))
    return out


def _steps(values: list) -> list:
    """Шаги между соседними координатами (мм)."""
    return [round(values[i + 1] - values[i], 2) for i in range(len(values) - 1)]


def _diff(measured: list, expected: list) -> str:
    """Строка расхождения замера с сеткой: «нет» либо расхождения по номерам.

    При одинаковом числе линий сравнение поэлементное (важен порядок: шаги
    2+3 и 3+2 — разные сетки), иначе — по наличию значений.
    """
    tol = 0.6
    if len(measured) == len(expected):
        bad = []
        for index, (m, e) in enumerate(zip(measured, expected)):
            if abs(m - e) > tol:
                bad.append(f"#{index + 1}: {m:g} вместо {e:g}")
        return "нет" if not bad else "; ".join(bad)
    missing = [e for e in expected
               if not any(abs(m - e) <= tol for m in measured)]
    extra = [m for m in measured
             if not any(abs(m - e) <= tol for e in expected)]
    parts = []
    if missing:
        parts.append("нет линий " + ", ".join(f"{v:g}" for v in missing))
    if extra:
        parts.append("лишние " + ", ".join(f"{v:g}" for v in extra))
    return "; ".join(parts) if parts else "нет"


def compare_reference(form: int, rel_v: list, rel_h: list) -> None:
    """Печатает сетку образца из корня проекта и расхождение шагов с замером."""
    ref = os.path.join(ROOT, FORMS[form][2])
    if not os.path.isfile(ref):
        print(f"образец не найден: {FORMS[form][2]}")
        return
    ref_v, ref_h = ref_grid(ref)
    print(f"образец «{FORMS[form][2]}»: вертикали {ref_v},\n"
          f"                            горизонтали {ref_h}")
    print("шаги образца: по X", _steps(ref_v), "по Y", _steps(ref_h))
    print("расхождение шагов с образцом: по X",
          _diff(_steps(rel_v), _steps(ref_v)), "по Y",
          _diff(_steps(rel_h), _steps(ref_h)))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--form", default="6",
                        help="какие формы рисовать: 3, 5, 6 или их набор "
                             "(по умолчанию 6 — последующий лист текста)")
    parser.add_argument("--no-pdf", action="store_true",
                        help="только DOCX, без конвертации в PDF/PNG")
    args = parser.parse_args()
    forms = [int(ch) for ch in str(args.form) if ch.isdigit()]
    forms = sorted({f for f in forms if f in FORMS})
    if not forms:
        print("Нечего рисовать: --form 3, 5, 6 или набор из них")
        return 2

    config = Config()
    name = ("форма_" + "_".join(str(f) for f in forms)) if len(forms) < 3 else "формы"
    out_path = os.path.join(ROOT, "tmp", name + ".docx")
    path = build(forms, out_path, config)
    print(f"DOCX: {path}")
    if os.path.normcase(os.path.abspath(path)) != os.path.normcase(out_path):
        print(f"ВНИМАНИЕ: {os.path.basename(out_path)} открыт в Word — сборка "
              f"сохранена рядом. Закройте документ, чтобы перезаписывать "
              f"{os.path.basename(out_path)}.")
    if args.no_pdf:
        return 0
    pdf_path, shots, zooms = render(path, forms)
    if not pdf_path:
        return 0
    print(f"PDF:  {pdf_path}")
    for shot in shots:
        print(f"PNG:  {shot}")
    for zoom in zooms:
        print(f"PNG (крупно): {zoom}")
    measurements = measure(pdf_path, forms)
    # Сверка с образцом — только для формы 6: «Форма 6.pdf» содержит ровно
    # фрагмент штампа, а «Форма 3.pdf» и «Форма 5.pdf» — целые листы чертежа,
    # где кроме штампа есть рамка и другие линии, и автоматическая сверка
    # сравнивала бы разнородные наборы.
    if forms == [6] and measurements and measurements[0][0]:
        compare_reference(6, *measurements[0])
    return 0


if __name__ == "__main__":
    sys.exit(main())
