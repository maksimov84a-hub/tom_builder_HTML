# -*- coding: utf-8 -*-
"""Рабочая копия раздела ПЗ для предпросмотра во встроенном Word.

Предпросмотр пояснительной записки показывает НЕ HTML-вёрстку, а настоящий
лист Word: файл собирается тем же кодом, что и «Собрать раздел», поэтому в нём
уже есть автоматическое оглавление с номерами листов и основные надписи.

Файл живёт в ``tmp/word_preview``. Это РАБОЧАЯ КОПИЯ: правки, сделанные в
Word, остаются в ней и не затрагивают исходный DOCX записки, пока пользователь
явно не перенесёт их в исходник. Пересборка выполняется по кнопке — тогда
копия создаётся заново.

Ключ кэша — состояние проекта, влияющее на вид листов (реквизиты, структура,
шрифты текста ПЗ) плюс состояние исходных DOCX. Пока ключ не менялся, файл
переиспользуется, и Word открывает его мгновенно.
"""
from __future__ import annotations

import hashlib
import json
import os


def _project_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def preview_dir() -> str:
    """Каталог рабочих копий предпросмотра (создаётся при обращении).

    Каталог — во временной папке системы, а НЕ в каталоге проекта: проект лежит
    в синхронизируемой папке (YandexDisk), и её клиент мешает Word открывать
    файл — COM-вызов ``Documents.Open`` зависал. Во временной папке файл
    открывается за полсекунды, а рабочая копия заодно не попадает в git.
    """
    import tempfile
    path = os.path.join(tempfile.gettempdir(), "tom_builder_preview")
    os.makedirs(path, exist_ok=True)
    return path


def preview_path(project=None, section_kind: str = "") -> str:
    """Базовый путь рабочей копии DOCX для раздела/документа предпросмотра."""
    from ..models import SECTION_PZ
    section = section_kind or SECTION_PZ
    safe = "".join(ch if ch.isalnum() else "_" for ch in str(section))[:40]
    return os.path.join(preview_dir(), f"preview_{safe}.docx")


def _manifest_path(path: str) -> str:
    return path + ".key"


def _read_manifest(path: str) -> dict:
    """Читает манифест сборки: ключ состояния и фактический путь файла."""
    try:
        with open(_manifest_path(path), "r", encoding="utf-8") as stream:
            data = json.load(stream)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _write_manifest(path: str, key: str, actual: str) -> None:
    try:
        with open(_manifest_path(path), "w", encoding="utf-8") as stream:
            json.dump({"key": key, "path": actual}, stream, ensure_ascii=False)
    except OSError:
        pass


def current_preview_path(project=None, section_kind: str = "") -> str:
    """Фактический файл последней сборки ("" — сборки ещё не было).

    Имя может отличаться от базового: если Word держал файл открытым, сборка
    уходит в следующее свободное имя (см. free_target).
    """
    base = preview_path(project, section_kind)
    stored = str(_read_manifest(base).get("path") or "")
    if stored and os.path.exists(stored):
        return stored
    return base if os.path.exists(base) else ""


def free_target(path: str) -> str:
    """Свободное имя для сборки: файл, занятый Word, заменить нельзя.

    Word открывает рабочую копию на запись и держит её, поэтому ``os.replace``
    на это имя падает с «файл занят». Сборка уходит в следующее имя
    ``preview_<раздел>_2.docx`` и т.д., а прежний файл удаляется, когда Word его
    отпустит.
    """
    if not os.path.exists(path):
        return path
    try:
        os.remove(path)          # прежняя сборка свободна — переиспользуем имя
        return path
    except OSError:
        pass                     # занята Word — берём следующее имя
    base, ext = os.path.splitext(path)
    for index in range(2, 1000):
        candidate = f"{base}_{index}{ext}"
        if not os.path.exists(candidate):
            return candidate
        try:
            os.remove(candidate)
            return candidate
        except OSError:
            continue
    return f"{base}_{os.getpid()}{ext}"


def _drop_other_versions(path: str, keep: str) -> None:
    """Удаляет прежние версии рабочей копии (занятые Word — молча пропускает)."""
    base, ext = os.path.splitext(path)
    folder = os.path.dirname(path) or "."
    prefix = os.path.basename(base) + "_"
    try:
        names = os.listdir(folder)
    except OSError:
        return
    for name in names:
        full = os.path.join(folder, name)
        if os.path.abspath(full) == os.path.abspath(keep):
            continue
        if name == os.path.basename(path):
            # Прежняя сборка под базовым именем: она больше не актуальна.
            try:
                os.remove(full)
            except OSError:
                pass
            continue
        if not name.startswith(prefix) or not name.endswith(ext):
            continue
        # Отсекаем посторонние имена: интересуют только «base_<число>.docx».
        tail = name[len(prefix):-len(ext)] if ext else name[len(prefix):]
        if not tail.isdigit():
            continue
        try:
            os.remove(full)
        except OSError:
            pass


def _file_state(path: str) -> list:
    """Состояние файла для ключа кэша: путь, время и размер."""
    try:
        stat = os.stat(path)
        return [os.path.abspath(path), stat.st_mtime_ns, stat.st_size]
    except OSError:
        return [os.path.abspath(path or ""), 0, 0]


def source_signature(project) -> list:
    """Состояние исходных документов ПЗ и настроек, влияющих на вид листов.

    Меняется — рабочая копия устарела и её нужно пересобрать.
    """
    from ..models import DocKind
    docs = []
    try:
        docs = [_file_state(d.path) for d in project.documents_for(DocKind.POYASNITELNAYA)
                if d.path]
    except Exception:
        docs = []
    cfg = getattr(project, "config", None)
    fonts = getattr(cfg, "text_fonts", None)
    font_state = [getattr(fonts, name, None) for name in
                  ("body_font", "body_size", "heading_font", "heading_size",
                   "table_font", "table_size")] if fonts is not None else []
    req = getattr(project, "requisites", None)
    req_state = {}
    if req is not None:
        req_state = {key: getattr(req, key, None) for key in
                     ("organization", "volume_id", "object_name", "stage", "city", "year")}
    return [docs, font_state, len(getattr(project, "structure", []) or []),
            _safe_json(req_state)]


def _safe_json(value) -> str:
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
    except Exception:
        return str(value)


def _signature_key(signature) -> str:
    return hashlib.sha256(_safe_json(signature).encode("utf-8")).hexdigest()


def is_stale(project, section_kind: str = "") -> bool:
    """Устарела ли рабочая копия (её нет или изменились исходники/настройки)."""
    base = preview_path(project, section_kind)
    manifest = _read_manifest(base)
    actual = str(manifest.get("path") or "")
    if not actual or not os.path.exists(actual):
        return True
    return str(manifest.get("key") or "") != _signature_key(
        source_signature(project))


def _upgrade_compatibility(path: str) -> None:
    """Переводит файл python-docx в режим совместимости Word 2013+.

    Шаблон python-docx объявляет ``w:compatibilityMode=14`` (Word 2010) и
    содержит устаревший ``w:useFELayout``. Word 2016/2019 открывает такой файл
    с пометкой «Режим ограниченной функциональности» и запрещает правку —
    проверить это можно по заголовку окна документа. Режим 15 и удаление
    useFELayout снимают пометку.
    """
    import zipfile

    tmp = path + ".compat"
    changed = False
    try:
        with zipfile.ZipFile(path) as source, \
                zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "word/settings.xml":
                    text = data.decode("utf-8")
                    new = text.replace(
                        'w:name="compatibilityMode" '
                        'w:uri="http://schemas.microsoft.com/office/word" w:val="14"',
                        'w:name="compatibilityMode" '
                        'w:uri="http://schemas.microsoft.com/office/word" w:val="15"')
                    new = new.replace("<w:useFELayout/>", "")
                    if new != text:
                        data = new.encode("utf-8")
                        changed = True
                target.writestr(item, data)
        if changed:
            os.replace(tmp, path)
        else:
            os.remove(tmp)
    except (OSError, zipfile.BadZipFile):
        try:
            os.remove(tmp)
        except OSError:
            pass


def build_preview_docx(project, section_kind: str = "", *, force: bool = False,
                       config=None) -> str:
    """Собирает (или возвращает готовую) рабочую копию DOCX раздела для Word.

    Сборка идёт штатным кодом раздела (`docx_exporter.build_section_docx`),
    поэтому в файле есть оглавление с номерами листов и основные надписи —
    ровно то, что попадёт в том.

    :param force: собрать заново, даже если копия актуальна (кнопка
        «Пересобрать»).
    :return: путь к рабочей копии.
    """
    from ..models import SECTION_PZ
    from ..config import load_config
    from .docx_exporter import build_section_docx

    section = section_kind or SECTION_PZ
    base = preview_path(project, section)
    if not force and not is_stale(project, section):
        return current_preview_path(project, section)
    signature = _signature_key(source_signature(project))
    # Сборка идёт в СВОБОДНОЕ имя: рабочую копию держит открытой Word, и
    # заменить занятый файл нельзя — иначе предпросмотр падал бы с диалогом
    # «Файл уже используется».
    target = free_target(base)
    tmp_path = target + ".build.docx"
    build_section_docx(project, section, tmp_path, config=config or load_config())
    try:
        os.replace(tmp_path, target)
    except OSError:
        # Файл успели открыть между проверкой и заменой — берём новое имя.
        target = free_target(target + "")
        os.replace(tmp_path, target)
    _upgrade_compatibility(target)
    _write_manifest(base, signature, target)
    _drop_other_versions(base, target)
    return target


def mark_stale(project, section_kind: str = "") -> None:
    """Помечает рабочую копию устаревшей (следующая сборка её перезапишет)."""
    try:
        os.remove(_manifest_path(preview_path(project, section_kind)))
    except OSError:
        pass


def drop_preview(project, section_kind: str = "") -> None:
    """Удаляет рабочую копию и её прежние версии (при закрытии проекта)."""
    base = preview_path(project, section_kind)
    actual = current_preview_path(project, section_kind)
    if actual:
        try:
            os.remove(actual)
        except OSError:
            pass
    _drop_other_versions(base, "")
    for suffix in (".key", ".build.docx", ".compat"):
        try:
            os.remove(base + suffix)
        except OSError:
            pass
