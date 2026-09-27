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

    # --- приложения (только если раздел приложений включён) ---
    if "appendix" in enabled_kinds:
        for a in project.appendices:
            if not a.needed:
                continue
            problem = _file_problem(a.file_path)
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