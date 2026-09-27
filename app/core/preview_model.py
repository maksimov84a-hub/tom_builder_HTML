# -*- coding: utf-8 -*-
"""
Модель предпросмотра: единый источник данных для HTML-листов.

Собирает из проекта (реквизиты + реестр подписантов + журнал изменений) один
«снимок» значений, который используется для отрисовки всех листов. Правка
любого поля в HTML меняет именно этот снимок (через модель), поэтому значение
обновляется везде, где оно используется.

Роли подписей (ключи signature_assignments):
    develop — Разработал
    check   — Проверил
    norm    — Нормоконтроль
    approve — Утвердил
    (дополнительно могут использоваться произвольные строки).
"""
from __future__ import annotations

from ..models import (VolumeProject, Signer, ChangeRecord, DocKind, Stage,
                      STAMP9_LOGO)
from .pdf_merge import compute_total_sheets
from .toc import content_rows, total_sheets as toc_total_sheets

# Канонические роли подписей в основной надписи / титуле.
ROLE_KEYS = ["develop", "check", "norm", "approve"]
ROLE_LABELS = {
    "develop": "Разработал",
    "check": "Проверил",
    "norm": "Нормоконтроль",
    "approve": "Утвердил",
}


def find_signer(project: VolumeProject, surname: str) -> Signer | None:
    """Находит подписанта по Фамилии (ключ сопоставления)."""
    if not surname:
        return None
    for s in project.signers:
        if s.surname.strip().lower() == surname.strip().lower():
            return s
    return None


def signer_for_role(project: VolumeProject, role_key: str) -> Signer | None:
    """Возвращает подписанта, назначенного на роль, или None."""
    surname = project.signature_assignments.get(role_key, "")
    return find_signer(project, surname)


def role_label(role_key: str) -> str:
    return ROLE_LABELS.get(role_key, role_key)


# ---- Карта «поле -> где используется» (для понимания пользователем) ---------

FIELD_USAGE = {
    "organization": "Обложка, титульный лист, основная надпись (графа 9 всех форм)",
    "cell9": "Основная надпись: графа 9 — текст организации или логотип (PDF)",
    "volume_id": "Обложка, титульный лист, основная надпись (все листы)",
    "object_name": "Титульный лист, основная надпись",
    "volume_name": "Титульный лист, содержание",
    "stage": "Титульный лист, основная надпись",
    "city": "Титульный лист",
    "year": "Титульный лист",
    "develop": "Основная надпись (графа «Разработал»), титульный лист",
    "check": "Основная надпись (графа «Проверил»), титульный лист",
    "norm": "Основная надпись (графа «Нормоконтроль»)",
    "approve": "Титульный лист (графа «Утвердил»)",
}


def _pz_toc(project: VolumeProject) -> tuple:
    """Строки автоматического оглавления ПЗ (словари) и документ-источник.

    Единый источник — app.core.pz_toc.project_toc_items (заголовки уровней
    текста из Word). Строки сериализуются в простые словари, чтобы снимок
    передавался в HTML. Пустой список означает «оглавление не выводится»
    (нет документа ПЗ, раздел исключён или оглавление отключено в конфигурации).
    """
    if not project.documents_for(DocKind.POYASNITELNAYA):
        return [], ""
    try:
        from .pz_toc import item_dicts, project_toc_items
        items, path = project_toc_items(project)
    except Exception:
        return [], ""
    return item_dicts(items), path


def build_preview_snapshot(project: VolumeProject) -> dict:
    """Собирает снимок значений для предпросмотра.

    Возвращает словарь, который легко сериализуется в JSON и передаётся в HTML.
    """
    req = project.requisites
    # Автоматическое оглавление ПЗ (app.core.pz_toc): строки уровней текста из
    # Word с номерами листов записки и документ-источник оглавления.
    _toc_items, _toc_path = _pz_toc(project)
    # Роли подписей: и Фамилия, и данные подписанта (если назначен).
    roles = {}
    for key in ROLE_KEYS:
        s = signer_for_role(project, key)
        roles[key] = {
            "key": key,
            "label": role_label(key),
            "surname": s.surname if s else "",
            "name": s.name if s else "",
            "position": s.position if s else "",
            "signature_image": s.signature_image if s else "",
            "has_signer": s is not None,
        }

    # Графа 9 основной надписи: текст по центру ячейки ЛИБО логотип (PDF).
    # PDF растеризуется в PNG (app.core.logo); результат кэшируется, поэтому
    # повторные сборки снимка не тратят время на рендер.
    c9 = getattr(project, "stamp_cell9", None)
    cell9 = {"mode": "text", "padding_mm": 0.4, "png": ""}
    if c9 is not None:
        cell9 = {"mode": c9.mode, "padding_mm": c9.padding_mm, "png": ""}
        if c9.mode == STAMP9_LOGO and c9.logo_path:
            try:
                from .logo import render_logo_png
                cell9["png"] = render_logo_png(c9.logo_path)
            except Exception:
                cell9["png"] = ""

    # Список подписантов для выбора в выпадающем списке.
    signer_list = [
        {"surname": s.surname, "name": s.name, "position": s.position,
         "signature_image": s.signature_image}
        for s in project.signers
    ]

    return {
        "requisites": {
            "organization": req.organization,
            "volume_id": req.volume_id,
            "object_name": req.object_name,
            "volume_name": req.volume_name,
            "stage": req.stage.value if isinstance(req.stage, Stage) else str(req.stage),
            "city": req.city,
            "year": req.year,
        },
        "sections": list(req.sections),
        "cell9": cell9,
        "roles": roles,
        "signers": signer_list,
        "changes": [
            {
                "number": c.number,
                "date": c.date,
                "description": c.description,
                "target": c.target,
                "signer_surname": c.signer_surname,
                "status": c.status.value if hasattr(c.status, "value") else str(c.status),
            }
            for c in project.changes
        ],
        "field_usage": FIELD_USAGE,
        # Автоматические строки содержания (единый источник app.core.toc):
        # список (наименование, обозначение, листы) по фактической структуре тома.
        "toc_rows": content_rows(project),
        # Автоматическое оглавление пояснительной записки (app.core.pz_toc):
        # заголовки уровней текста из Word с номерами листов записки. Пустой
        # список — оглавление не выводится (нет DOCX или отключено в конфиге).
        "pz_toc_items": _toc_items,
        # Документ, по которому построено оглавление: лист оглавления
        # показывается только для него (и для раздела ПЗ), а не для каждого
        # документа записки по отдельности.
        "pz_toc_path": _toc_path,
        # Сквозная нумерация: общее число листов тома и текущий номер листа.
        # «Листов M» заполняется автоматически при сборке записки.
        "sheets_total": str(toc_total_sheets(project)),
        "sheet_no": "",
        "count_sheets": {
            # число листов журнала изменений (1 лист на ~20 строк, мин. 1)
            "changes_pages": max(1, (len(project.changes) + 19) // 20),
        },
    }
