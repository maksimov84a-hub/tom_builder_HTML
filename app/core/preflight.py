# -*- coding: utf-8 -*-
"""Preflight-проверки проекта тома перед экспортом.

Разделяет проблемы на errors (блокируют сборку) и warnings (показываются
пользователю с возможностью продолжить). Единая точка вызова — UI экспорта
(app.ui.main_window); ядро не зависит от Qt.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field

from ..models import VolumeProject, SECTION_PZ


@dataclass
class PreflightResult:
    """Результат проверки: ошибки блокируют экспорт, warnings — нет."""
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors

    @property
    def has_warnings(self) -> bool:
        return bool(self.warnings)


def _file_problem(path: str) -> str | None:
    """Возвращает текст ошибки, если файл отсутствует/нечитаем, иначе None."""
    if not path:
        return "не указан файл"
    if not os.path.exists(path):
        return f"файл не найден: {path}"
    if not os.path.isfile(path):
        return f"не файл: {path}"
    try:
        with open(path, "rb") as f:
            f.read(1)
    except OSError as e:
        return f"не удалось прочитать {path}: {e}"
    return None


def _content_problem(path: str) -> str | None:
    """Проверяет, что файл действительно открывается как DOCX или PDF.

    Файл может существовать и читаться, но быть битым или иметь неверное
    расширение (например, старый .doc под именем .docx). Раньше такой файл
    проходил проверку и ронял сборку.
    """
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".docx", ".doc"):
            import docx
            docx.Document(path)
        elif ext == ".pdf":
            import pymupdf
            with pymupdf.open(path) as pdf:
                pdf.page_count
    except Exception as e:
        return (f"файл не открывается: {path} ({type(e).__name__}: {e}). "
                f"Проверьте формат: поддерживаются DOCX и PDF")
    return None


def _readable_problem(path: str) -> str | None:
    """Документ Word читается, но текста для вывода в нём нет.

    Возвращает предупреждение (лист выйдет заглушкой) либо None.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".docx", ".doc"):
        return None
    try:
        from .docx_reader import read_docx
        content = read_docx(path)
    except Exception:
        return None            # ошибку открытия уже сообщил _content_problem
    if getattr(content, "is_empty", False):
        return ("файл прочитан, но текста для вывода в нём нет — "
                "лист будет показан заглушкой (проверьте, не свёрстан ли текст "
                "в надписях/фигурах или колонтитулах)")
    if getattr(content, "text_in_textboxes", False):
        return ("текст найден в надписях/фигурах Word — он будет выведен, но "
                "правка в предпросмотре к такому тексту не применяется")
    return None


def run_preflight(project: VolumeProject) -> PreflightResult:
    """Проверяет проект перед сборкой тома/раздела.

    Ошибки:
      * structure пуст или ни один раздел не включён;
      * у включённого документа/приложения отсутствует или нечитается файл.

    Предупреждения:
      * включён раздел изменений, но журнал пуст;
      * пустые ключевые реквизиты тома;
      * Word COM недоступен (PDF пойдёт через PyMuPDF-fallback).
    """
    res = PreflightResult()

    # --- structure ---
    entries = [e for e in (project.structure or []) if e.get("needed", True)]
    if not project.structure:
        res.errors.append("структура тома пуста — нечего собирать")
    elif not entries:
        res.errors.append("в структуре тома не отмечено ни одного раздела")
    enabled_kinds = {e.get("kind") for e in entries}

    # --- документы (только разделы, включённые в structure) ---
    for d in project.documents:
        if not d.needed:
            continue
        if d.doc_kind.value not in enabled_kinds:
            continue  # раздел выключен в дереве — файл не участвует в сборке
        problem = _file_problem(d.path)
        if problem:
            res.errors.append(f"документ «{d.doc_kind.value}»: {problem}")
            continue
        # Ошибка только по включённым документам; битый файл ловится здесь,
        # а не падением сборки.
        problem = _content_problem(d.path)
        if problem:
            res.errors.append(f"документ «{d.doc_kind.value}»: {problem}")
            continue
        note = _readable_problem(d.path)
        if note:
            res.warnings.append(f"документ «{d.doc_kind.value}»: {note}")

    # --- приложения (только если раздел приложений включён) ---
    if "appendix" in enabled_kinds:
        for a in project.appendices:
            if not a.needed:
                continue
            problem = _file_problem(a.file_path)
            if problem:
                label = a.number or a.name or a.file_path or "?"
                res.errors.append(f"приложение {label}: {problem}")
                continue
            problem = _content_problem(a.file_path)
            if problem:
                label = a.number or a.name or a.file_path or "?"
                res.errors.append(f"приложение {label}: {problem}")

    # --- журнал изменений ---
    changes_needed = any(
        e.get("kind") == "changes" for e in entries
    )
    if changes_needed and not project.changes:
        res.warnings.append(
            "раздел «Лист регистрации изменений» включён, но журнал пуст — "
            "лист в сборку не попадёт"
        )

    # --- реквизиты ---
    req = project.requisites
    if not (req.volume_id or "").strip():
        res.warnings.append("не заполнено обозначение тома")
    if not (req.volume_name or "").strip() and not (req.object_name or "").strip():
        res.warnings.append("не заполнено наименование тома и объекта")
    if not (req.organization or "").strip():
        res.warnings.append("не заполнена организация")

    # --- Word COM (только если вообще есть что конвертировать) ---
    has_text = any(
        e.get("kind") in ("cover", "title", "content", SECTION_PZ, "changes")
        for e in entries
    )
    if has_text:
        try:
            from .word_com import word_available
            if not word_available():
                res.warnings.append(
                    "Microsoft Word недоступен — текстовая часть PDF будет "
                    "собрана упрощённым способом (PyMuPDF)"
                )
        except Exception as e:  # noqa: BLE001 — preflight не должен падать
            res.warnings.append(f"не удалось проверить Word COM: {e}")

    return res