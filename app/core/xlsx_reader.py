# -*- coding: utf-8 -*-
"""Чтение таблицы перечня листов из Excel (.xlsx) без сторонних библиотек.

В проекте нет openpyxl (vendor/ собран под другой набор зависимостей, и его
нельзя дополнять), поэтому xlsx читается напрямую: это ZIP с XML внутри.
Разбираются только нужные части:

    xl/workbook.xml        — имена листов и их порядок
    xl/_rels/workbook.xml.rels — связь имени листа с файлом листа
    xl/sharedStrings.xml   — общий пул строк (в ячейках лежат ссылки на него)
    xl/worksheets/sheetN.xml — сами ячейки

Этого достаточно, чтобы прочитать текст таблицы; форматирование, формулы и
стили намеренно не разбираются — перечень листов их не использует.
"""
from __future__ import annotations

import os
import re
import zipfile
import xml.etree.ElementTree as ET

#: Пространства имён OOXML (в разных файлах встречаются разные префиксы,
#: поэтому сравниваем по локальному имени тега).
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _local(tag: str) -> str:
    """Локальное имя тега без пространства имён."""
    return tag.rsplit("}", 1)[-1]


def _col_index(ref: str) -> int:
    """Индекс колонки из адреса ячейки: 'A1' -> 0, 'C7' -> 2, 'AA1' -> 26."""
    letters = re.match(r"([A-Za-z]+)", ref or "")
    if not letters:
        return 0
    index = 0
    for ch in letters.group(1).upper():
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def sheet_names(path: str) -> list:
    """Имена листов книги в порядке книги (пустой список — файл не читается)."""
    if not path or not os.path.exists(path):
        return []
    try:
        with zipfile.ZipFile(path) as zf:
            data = zf.read("xl/workbook.xml")
    except (KeyError, OSError, zipfile.BadZipFile):
        return []
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []
    names = []
    for el in root.iter():
        if _local(el.tag) == "sheet":
            name = el.get("name") or ""
            if name:
                names.append(name)
    return names


def _sheet_path(zf: zipfile.ZipFile, sheet_name: str) -> str:
    """Путь к XML листа внутри пакета по его имени."""
    # По умолчанию — первый лист: связь rId1 -> worksheets/sheet1.xml.
    default = "xl/worksheets/sheet1.xml"
    try:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, ET.ParseError):
        return default
    # Имя листа -> r:id
    rid = ""
    order = []
    for el in wb.iter():
        if _local(el.tag) != "sheet":
            continue
        name = el.get("name") or ""
        order.append(name)
        if not rid and (not sheet_name or name == sheet_name):
            rid = el.get("{%s}id" % _NS_REL) or ""
    # r:id -> путь
    for el in rels:
        if _local(el.tag) != "Relationship":
            continue
        if rid and el.get("Id") != rid:
            continue
        target = el.get("Target") or ""
        if not target:
            continue
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        return target
    return default


def _shared_strings(zf: zipfile.ZipFile) -> list:
    """Пул строк книги (в ячейках хранятся индексы этого пула)."""
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []
    out = []
    for si in root:
        if _local(si.tag) != "si":
            continue
        # Текст может быть разбит на несколько <t> (например, при
        # форматировании части строки) — склеиваем все в порядке следования.
        parts = [t.text or "" for t in si.iter() if _local(t.tag) == "t"]
        out.append("".join(parts))
    return out


def _cell_text(cell, shared: list) -> str:
    """Текст ячейки: значение из пула строк или вложенное (число, дата)."""
    ctype = cell.get("t") or ""
    value = ""
    inline = ""
    for el in cell.iter():
        tag = _local(el.tag)
        if tag == "v":
            value = el.text or ""
        elif tag == "is":
            inline = "".join(t.text or "" for t in el.iter() if _local(t.tag) == "t")
    if ctype == "s":
        try:
            return shared[int(value)]
        except (ValueError, IndexError):
            return ""
    if inline:
        return inline
    return value


def read_rows(path: str, sheet_name: str = "") -> list:
    """Строки таблицы из xlsx: список списков строк (по колонкам).

    Пустые хвостовые колонки обрезаются, полностью пустые строки сохраняются
    как пустые списки — вызывающий код решает, пропускать их или нет.
    """
    if not path or not os.path.exists(path):
        return []
    try:
        with zipfile.ZipFile(path) as zf:
            shared = _shared_strings(zf)
            try:
                data = zf.read(_sheet_path(zf, sheet_name))
            except KeyError:
                return []
    except (OSError, zipfile.BadZipFile):
        return []
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []

    rows = []
    for row_el in root.iter():
        if _local(row_el.tag) != "row":
            continue
        cells = {}
        for cell in row_el:
            if _local(cell.tag) != "c":
                continue
            text = _cell_text(cell, shared).strip()
            if not text:
                continue
            cells[_col_index(cell.get("r") or "A1")] = text
        if not cells:
            rows.append([])
            continue
        width = max(cells) + 1
        rows.append([cells.get(i, "") for i in range(width)])
    return rows


def read_sheet_list(path: str, sheet_name: str = "") -> list:
    """Перечень листов из таблицы: [(обозначение, наименование, примечание)].

    Первая непустая строка считается шапкой и пропускается. Строки, где пусты
    и обозначение, и наименование, отбрасываются — так в таблице можно
    оставлять пустые строки-разделители.
    """
    rows = read_rows(path, sheet_name)
    out = []
    seen_header = False
    for row in rows:
        if not row or not any(cell.strip() for cell in row):
            continue
        cells = (list(row) + ["", "", ""])[:3]
        if not seen_header:
            # Шапка: считаем ею первую строку, если она не похожа на данные.
            seen_header = True
            low0 = cells[0].strip().lower()
            low1 = cells[1].strip().lower()
            if ("обозн" in low0 or "наимен" in low1
                    or "наимен" in cells[0].lower()):
                continue
        designation, name, note = (c.strip() for c in cells)
        if not designation and not name:
            continue
        out.append((designation, name, note))
    return out


#: Признаки шапки таблицы «Состав проектной документации»: если первая строка
#: содержит эти слова, она считается заголовком колонок и не выводится.
_COMPOSITION_HEAD_HINTS = ("номер том", "обозн", "наимен", "примечан")


def read_composition(path: str, sheet_name: str = "") -> list:
    """Таблица «Состав проектной документации» из книги Excel.

    :return: строки [(номер тома, обозначение, наименование, примечание)].

    Первая непустая строка считается шапкой (пропускается), полностью пустые
    строки отбрасываются — так в книге можно оставлять пустые строки-разделители.
    Числовые ячейки (номер тома, год) читаются как есть текстом.
    """
    from ..models import COMPOSITION_COLS
    ncols = len(COMPOSITION_COLS)
    rows = read_rows(path, sheet_name)
    out = []
    seen_header = False
    for row in rows:
        if not row or not any(cell.strip() for cell in row):
            continue
        cells = (list(row) + [""] * ncols)[:ncols]
        if not seen_header:
            seen_header = True
            low = " ".join(c.strip().lower() for c in cells)
            if any(hint in low for hint in _COMPOSITION_HEAD_HINTS):
                continue
        values = tuple(c.strip() for c in cells)
        if not any(values):
            continue
        out.append(values)
    return out
