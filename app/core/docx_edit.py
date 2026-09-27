# -*- coding: utf-8 -*-
"""Запись правок обратно в исходный DOCX.

Поддерживаются два вида правок, которые пользователь делает прямо в предпросмотре
пояснительной записки (или в отдельном окне редактора):

  * текст абзацев — {индекс абзаца в document.paragraphs: новый текст};
  * текст в ячейках таблиц — {"таблица:строка:ячейка": новый текст}, где индексы
    берутся из document.tables (та же адресация, что и в app.core.docx_reader).

Оба вида записываются ОДНИМ открытием и сохранением файла (apply_docx_edits),
поэтому правки не конфликтуют друг с другом и создаётся одна резервная копия
<path>.docx.bak.

Сохраняются таблицы, изображения, колонтитулы и разметка страниц: изменяется
только текст указанных абзацев/ячеек.
"""
from __future__ import annotations

import os
import shutil


def replace_paragraph_text(paragraph, text: str) -> None:
    """Заменяет текст абзаца, сохраняя формат первого run."""
    runs = paragraph.runs
    if not runs:
        paragraph.add_run(text)
        return
    runs[0].text = text
    for run in runs[1:]:
        run.text = ""


def replace_cell_text(cell, text: str) -> None:
    """Заменяет текст ячейки таблицы, сохраняя формат первого run.

    Многострочный текст (переносы строки из contenteditable) раскладывается по
    абзацам ячейки: первый абзац получает первую строку, остальные строки —
    следующие абзацы (при нехватке абзацы добавляются). Лишние абзацы очищаются,
    но не удаляются, чтобы не потерять форматирование ячейки.
    """
    rows = str(text if text is not None else "").split("\n")
    paragraphs = list(getattr(cell, "paragraphs", []) or [])
    if not paragraphs:
        cell.text = rows[0] if rows else ""
        return
    replace_paragraph_text(paragraphs[0], rows[0] if rows else "")
    for i, row_text in enumerate(rows[1:], start=1):
        if i < len(paragraphs):
            replace_paragraph_text(paragraphs[i], row_text)
        else:
            try:
                cell.add_paragraph(row_text)
            except Exception:          # pragma: no cover
                pass
    for paragraph in paragraphs[len(rows):]:
        replace_paragraph_text(paragraph, "")


def parse_cell_address(address) -> tuple | None:
    """Разбирает адрес ячейки "таблица:строка:ячейка" в кортеж (t, r, c).

    Возвращает None, если адрес не разбирается. Поддерживается и словарь с
    ключами table/row/cell — на случай передачи адреса структурой.
    """
    if isinstance(address, dict):
        try:
            return (int(address.get("table", -1)), int(address.get("row", -1)),
                    int(address.get("cell", -1)))
        except (TypeError, ValueError):
            return None
    if isinstance(address, (list, tuple)) and len(address) == 3:
        try:
            return (int(address[0]), int(address[1]), int(address[2]))
        except (TypeError, ValueError):
            return None
    try:
        parts = str(address).split(":")
        if len(parts) != 3:
            return None
        return (int(parts[0]), int(parts[1]), int(parts[2]))
    except (TypeError, ValueError):
        return None


def _backup(path: str) -> None:
    """Создаёт резервную копию <path>.bak, если её ещё нет."""
    bak = path + ".bak"
    if os.path.exists(bak):
        return
    try:
        shutil.copy2(path, bak)
    except OSError:
        pass


def _open_document(path: str):
    """Открывает DOCX или поднимает понятную ошибку."""
    if not path or not os.path.exists(path):
        raise FileNotFoundError(path or "(пустой путь)")
    try:
        from docx import Document
    except Exception as exc:                      # pragma: no cover
        raise RuntimeError(f"python-docx недоступен: {exc}") from exc
    return Document(path)


def _save_document(document, path: str) -> None:
    try:
        document.save(path)
    except Exception as exc:
        raise RuntimeError(f"не удалось сохранить DOCX: {exc}") from exc


def apply_docx_edits(path: str, paragraph_edits: dict | None = None,
                     table_edits: dict | None = None, *, backup: bool = True) -> dict:
    """Записывает правки абзацев и ячеек таблиц за одно сохранение файла.

    :param path: путь к исходному DOCX (перезаписывается).
    :param paragraph_edits: {индекс абзаца в document.paragraphs: новый текст}.
    :param table_edits: {"таблица:строка:ячейка": новый текст}.
    :param backup: создать рядом резервную копию <path>.bak при первом сохранении.
    :return: {"paragraphs": n, "table_cells": m} — сколько правок применено.
    :raises FileNotFoundError: файл не найден.
    :raises RuntimeError: файл не удалось открыть или сохранить.
    """
    paragraph_edits = paragraph_edits or {}
    table_edits = table_edits or {}
    result = {"paragraphs": 0, "table_cells": 0}
    if not paragraph_edits and not table_edits:
        return result

    document = _open_document(path)
    paragraphs = list(document.paragraphs)
    for raw_index, text in paragraph_edits.items():
        try:
            index = int(raw_index)
        except (TypeError, ValueError):
            continue
        if not (0 <= index < len(paragraphs)):
            continue
        paragraph = paragraphs[index]
        new_text = "" if text is None else str(text)
        if paragraph.text == new_text:
            continue
        replace_paragraph_text(paragraph, new_text)
        result["paragraphs"] += 1

    if table_edits:
        tables = list(document.tables)
        for raw_address, text in table_edits.items():
            address = parse_cell_address(raw_address)
            if address is None:
                continue
            table_index, row_index, cell_index = address
            if not (0 <= table_index < len(tables)):
                continue
            try:
                cell = tables[table_index].rows[row_index].cells[cell_index]
            except (IndexError, ValueError):
                continue
            new_text = "" if text is None else str(text)
            if (cell.text or "").strip() == new_text.strip():
                continue
            replace_cell_text(cell, new_text)
            result["table_cells"] += 1

    if result["paragraphs"] or result["table_cells"]:
        if backup:
            _backup(path)
        _save_document(document, path)
    return result


def apply_paragraph_edits(path: str, edits: dict, *, backup: bool = True) -> int:
    """Записывает правки абзацев ({индекс: текст}) в DOCX. Возвращает число
    фактически изменённых абзацев."""
    return apply_docx_edits(path, paragraph_edits=edits, backup=backup)["paragraphs"]


def apply_table_cell_edits(path: str, edits: dict, *, backup: bool = True) -> int:
    """Записывает правки ячеек таблиц ({"таблица:строка:ячейка": текст}) в DOCX.

    Возвращает число фактически изменённых ячеек.
    """
    return apply_docx_edits(path, table_edits=edits, backup=backup)["table_cells"]
