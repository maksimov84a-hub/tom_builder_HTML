# -*- coding: utf-8 -*-
"""
Модель тома: данные, которые программа собирает и оформляет по ГОСТ.

Проект тома = объект VolumeProject, сохраняемый в JSON (сборочное задание).
Том можно пересобрать при изменении исходников, т.к. проект хранит
ссылки на файлы, а не их содержимое.
"""
from __future__ import annotations

import json
import os
import uuid
from dataclasses import dataclass, field, asdict
from enum import Enum
from typing import Optional


def as_stage(v):
    """Коэрция стадии: принимает Stage или его строковое значение."""
    if isinstance(v, Stage):
        return v
    if isinstance(v, str):
        for s in Stage:
            if s.value == v:
                return s
        return Stage(v)
    return Stage.P


def as_change_status(v):
    """Коэрция статуса изменения: принимает ChangeStatus или его строковое значение."""
    if isinstance(v, ChangeStatus):
        return v
    if isinstance(v, str):
        for s in ChangeStatus:
            if s.value == v:
                return s
    return ChangeStatus.APPLIED



# ---- Перечисления ----------------------------------------------------------

class Stage(str, Enum):
    """Стадия проектной документации."""
    P = "П"   # проектная
    R = "Р"   # рабочая


class ChangeStatus(str, Enum):
    """Статус записи о внесении изменений."""
    APPLIED = "Внесено"
    IN_WORK = "В работе"
    CANCELLED = "Отменено"


class DocKind(str, Enum):
    """Тип документа в томе."""
    COVER = "Обложка"
    TITLE = "Титульный лист"
    CONTENT = "Содержание"
    POYASNITELNAYA = "Пояснительная записка"
    GRAPHIC = "Графическая часть"
    APPENDIX = "Приложение"
    CHANGES = "Лист регистрации изменений"


# ---- Структура тома (настраиваемый порядок и необходимость блоков) --------

# Ключи разделов тома. Два из них совпадают со значениями DocKind (ПЗ/Графика),
# остальные — служебные строки для фиксированных листов.
SECTION_COVER = "cover"
SECTION_TITLE = "title"
SECTION_CONTENT = "content"
SECTION_PZ = DocKind.POYASNITELNAYA.value     # "Пояснительная записка"
SECTION_GRAPHIC = DocKind.GRAPHIC.value       # "Графическая часть"
SECTION_APPENDIX = "appendix"
SECTION_CHANGES = "changes"

# Человекочитаемые заголовки разделов (для дерева и содержания).
SECTION_LABELS = {
    SECTION_COVER: "Обложка",
    SECTION_TITLE: "Титульный лист",
    SECTION_CONTENT: "Содержание",
    SECTION_PZ: "Пояснительная записка",
    SECTION_GRAPHIC: "Графическая часть",
    SECTION_APPENDIX: "Приложения",
    SECTION_CHANGES: "Лист регистрации изменений",
}


def default_structure() -> list:
    """Стандартный порядок блоков тома по ГОСТ (по умолчанию)."""
    return [
        {"kind": SECTION_COVER, "needed": True},
        {"kind": SECTION_TITLE, "needed": True},
        {"kind": SECTION_CONTENT, "needed": True},
        {"kind": SECTION_PZ, "needed": True},
        {"kind": SECTION_GRAPHIC, "needed": True},
        {"kind": SECTION_APPENDIX, "needed": True},
        {"kind": SECTION_CHANGES, "needed": True},
    ]


# ---- Источник (входной файл) ----------------------------------------------

@dataclass
class SourceFile:
    """Ссылка на исходный файл документа (Word или PDF)."""
    path: str = ""
    doc_kind: DocKind = DocKind.POYASNITELNAYA
    pages: int = 0     # число листов (для PDF вычисляется; для DOCX — после рендера)
    size: tuple = (0, 0)  # (ширина, высота) листа в мм (для графики/приложений)
    needed: bool = True    # True — включается в сборку; False — исключён (отметка необходимости)

    def to_dict(self) -> dict:
        d = asdict(self)
        d["doc_kind"] = self.doc_kind.value
        d["size"] = list(self.size)
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "SourceFile":
        sf = cls()
        sf.path = d.get("path", "")
        sf.doc_kind = DocKind(d.get("doc_kind", sf.doc_kind.value))
        sf.pages = int(d.get("pages", 0))
        size = d.get("size") or [0, 0]
        sf.size = (size[0], size[1])
        sf.needed = bool(d.get("needed", True))
        return sf


# ---- Подписант -------------------------------------------------------------

@dataclass
class Signer:
    """Участник, подписывающий документы. Подпись — изображение."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    surname: str = ""            # Фамилия — ключ для сопоставления в графах
    name: str = ""               # Имя/Отчество (И.И.)
    position: str = ""           # Должность (ГИП, Нач. отд. ...)
    role: str = ""               # Роль (Разраб., Пров., Н.контр., Утв.)
    signature_image: str = ""    # путь к изображению подписи

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Signer":
        s = cls()
        for k in ("id", "surname", "name", "position", "role", "signature_image"):
            if k in d:
                setattr(s, k, d[k])
        return s


# ---- Запись об изменении ----------------------------------------------------

@dataclass
class ChangeRecord:
    """Запись о внесении изменения в том."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    number: int = 0             # № изменения (порядковый)
    date: str = ""              # дата внесения
    description: str = ""       # содержание изменения
    target: str = ""            # раздел/лист, к которому относится
    signer_surname: str = ""    # исполнитель (Фамилия -> подпись из реестра)
    status: ChangeStatus = ChangeStatus.APPLIED

    def to_dict(self) -> dict:
        d = asdict(self)
        d["status"] = as_change_status(self.status).value
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "ChangeRecord":
        c = cls()
        for k in ("id", "number", "date", "description", "target", "signer_surname"):
            if k in d:
                setattr(c, k, d[k])
        if "status" in d and d["status"]:
            c.status = as_change_status(d["status"])
        return c


# ---- Приложение --------------------------------------------------------------

@dataclass
class Appendix:
    """Приложение (PDF-файл), вставляемое в том."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    number: str = ""            # номер: буквенный А, Б, В... или числовой
    name: str = ""              # наименование приложения
    file_path: str = ""         # путь к PDF
    target: str = ""            # раздел-привязка (опционально)
    pages: int = 0              # количество листов (вычисляется из PDF)
    needed: bool = True         # True — включается в сборку

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Appendix":
        a = cls()
        for k in ("id", "number", "name", "file_path", "target", "pages"):
            if k in d:
                setattr(a, k, d[k])
        a.needed = bool(d.get("needed", True))
        return a


# ---- Реквизиты тома ----------------------------------------------------------

@dataclass
class Requisites:
    """Реквизиты тома — заполняются через форму."""
    organization: str = ""               # наименование организации
    volume_id: str = ""                  # обозначение/шифр тома
    object_name: str = ""                # наименование объекта
    volume_name: str = ""                # наименование тома (раздела)
    stage: Stage = Stage.P               # П / Р
    city: str = ""                       # город
    year: int = 0                        # год
    # разделы тома (для содержания) — список строк
    sections: list = field(default_factory=list)

    def to_dict(self) -> dict:
        d = asdict(self)
        d["stage"] = as_stage(self.stage).value
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "Requisites":
        r = cls()
        for k in ("organization", "volume_id", "object_name", "volume_name", "city", "year", "sections"):
            if k in d:
                setattr(r, k, d[k])
        if "stage" in d and d["stage"]:
            r.stage = as_stage(d["stage"])
        return r


# ---- Обложка (приложение Е ГОСТ Р 21.101-2020) -------------------------------

@dataclass
class CoverSheet:
    """Обложка тома — первый лист по ГОСТ (шаблон или авто-генерация)."""
    template_path: str = ""      # путь к шаблону (DOCX/изображение) или "" — авто
    is_image: bool = False       # True — изображение, False — DOCX
    use_custom: bool = False     # True — использовать кастомный шаблон, False — авто по ГОСТ
    show_logo: bool = True       # показывать ли логотип организации

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "CoverSheet":
        t = cls()
        for k in ("template_path", "is_image", "use_custom", "show_logo"):
            if k in d:
                setattr(t, k, d[k])
        return t


# ---- Титульный лист (приложение Ж ГОСТ Р 21.101-2020) -----------------------

@dataclass
class TitleSheet:
    """Титульный лист тома — второй лист по ГОСТ (кастомный шаблон или авто)."""
    template_path: str = ""      # путь к шаблону
    is_image: bool = False       # True — изображение, False — DOCX
    use_custom: bool = False     # True — использовать кастомный шаблон, False — авто по ГОСТ

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "TitleSheet":
        t = cls()
        for k in ("template_path", "is_image", "use_custom"):
            if k in d:
                setattr(t, k, d[k])
        return t


# ---- Графа 9 основной надписи (наименование/индекс организации) ---------------

STAMP9_TEXT = "text"
STAMP9_LOGO = "logo"


@dataclass
class StampCell9:
    """Содержимое графы 9 основной надписи (ГОСТ Р 21.101-2020, прил. Ж).

    Графа 9 — «наименование или различительный индекс организации, разработавшей
    документ». В формах 5 и 3 это правая верхняя ячейка основной надписи
    (50 x 10 мм); в форме 6 графы 9 нет.

    Заполняется одним из двух способов (переключатель «Текст / Логотип»):
      * текст — выводится по центру ячейки;
      * логотип — изображение из PDF-файла, автоматически масштабируется под
        ячейку с сохранением пропорций (поля PDF обрезаются, см. app/core/logo.py).
    """
    mode: str = STAMP9_TEXT       # "text" — текст, "logo" — логотип
    logo_path: str = ""           # путь к PDF-файлу логотипа
    padding_mm: float = 0.4       # внутренний отступ логотипа в ячейке, мм

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "StampCell9":
        s = cls()
        for k in ("mode", "logo_path"):
            v = d.get(k)
            if v is not None:
                setattr(s, k, v)
        try:
            s.padding_mm = float(d.get("padding_mm", s.padding_mm))
        except (TypeError, ValueError):
            pass
        if s.mode not in (STAMP9_TEXT, STAMP9_LOGO):
            s.mode = STAMP9_TEXT
        return s

    def use_logo(self) -> bool:
        """True — графа 9 заполняется логотипом (файл задан и существует)."""
        return (self.mode == STAMP9_LOGO and bool(self.logo_path)
                and os.path.exists(self.logo_path))


# ---- Проект тома --------------------------------------------------------------

@dataclass
class VolumeProject:
    """Собранное сборочное задание тома."""
    name: str = "Новый том"
    requisites: Requisites = field(default_factory=Requisites)
    cover_sheet: CoverSheet = field(default_factory=CoverSheet)   # обложка (прил. Е)
    title_sheet: TitleSheet = field(default_factory=TitleSheet)   # титульный лист (прил. Ж)
    signers: list = field(default_factory=list)          # Signer
    changes: list = field(default_factory=list)          # ChangeRecord
    appendices: list = field(default_factory=list)       # Appendix
    documents: list = field(default_factory=list)        # SourceFile, в порядке структуры тома
    page_margins_override: Optional[dict] = None         # переопределение полей (опционально)
    # Единый источник ролей подписей: role (develop/check/norm/approve) -> Фамилия подписанта.
    # Одна и та же роль используется на всех листах, поэтому правка роли обновляется везде.
    signature_assignments: dict = field(default_factory=dict)
    # Настраиваемая структура тома: список {kind, needed} в порядке следования блоков.
    structure: list = field(default_factory=default_structure)
    # Графа 9 основной надписи: текст или логотип (PDF).
    stamp_cell9: StampCell9 = field(default_factory=StampCell9)

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "requisites": self.requisites.to_dict(),
            "cover_sheet": self.cover_sheet.to_dict(),
            "title_sheet": self.title_sheet.to_dict(),
            "signers": [s.to_dict() for s in self.signers],
            "changes": [c.to_dict() for c in self.changes],
            "appendices": [a.to_dict() for a in self.appendices],
            "documents": [d.to_dict() for d in self.documents],
            "page_margins_override": self.page_margins_override,
            "signature_assignments": dict(self.signature_assignments),
            "structure": list(self.structure),
            "stamp_cell9": self.stamp_cell9.to_dict(),
        }

    @classmethod
    def from_dict(cls, d: dict) -> "VolumeProject":
        vp = cls()
        vp.name = d.get("name", vp.name)
        vp.requisites = Requisites.from_dict(d.get("requisites") or {})
        vp.cover_sheet = CoverSheet.from_dict(d.get("cover_sheet") or {})
        vp.title_sheet = TitleSheet.from_dict(d.get("title_sheet") or {})
        vp.signers = [Signer.from_dict(x) for x in (d.get("signers") or [])]
        vp.changes = [ChangeRecord.from_dict(x) for x in (d.get("changes") or [])]
        vp.appendices = [Appendix.from_dict(x) for x in (d.get("appendices") or [])]
        vp.documents = [SourceFile.from_dict(x) for x in (d.get("documents") or [])]
        vp.page_margins_override = d.get("page_margins_override")
        vp.signature_assignments = dict(d.get("signature_assignments") or {})
        vp.structure = list(d.get("structure") or default_structure())
        vp.stamp_cell9 = StampCell9.from_dict(d.get("stamp_cell9") or {})
        return vp

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path: str) -> "VolumeProject":
        with open(path, "r", encoding="utf-8") as f:
            return cls.from_dict(json.load(f))

    # ---- Структура тома (настраиваемый порядок и необходимость) ----

    def structure_kinds(self) -> list:
        """Порядок ключей разделов согласно структуре тома."""
        return [entry.get("kind") for entry in self.structure]

    def structure_needed(self, kind: str) -> bool:
        """Нужен ли раздел kind в сборке (по структуре). По умолчанию — да."""
        for entry in self.structure:
            if entry.get("kind") == kind:
                return bool(entry.get("needed", True))
        return True

    def documents_for(self, kind: DocKind) -> list:
        """Документы раздела kind, включённые в сборку (needed=True)."""
        return [d for d in self.documents if d.doc_kind == kind and d.needed]

    def appendices_needed(self) -> list:
        """Приложения, включённые в сборку (needed=True)."""
        return [a for a in self.appendices if a.needed]

    def structure_document_kinds(self) -> list:
        """Документные разделы, идущие в структуре (ПЗ/Графика)."""
        return [k for k in self.structure_kinds()
                if k in (DocKind.POYASNITELNAYA.value, DocKind.GRAPHIC.value)]


# ---- Быстрая проверка --------------------------------------------------------

if __name__ == "__main__":
    vp = VolumeProject()
    vp.requisites.organization = "ООО «Пример»"
    vp.requisites.volume_id = "12-2024-ПЗ"
    vp.requisites.stage = Stage.P
    vp.signers.append(Signer(surname="Иванов", name="И.И.", position="ГИП", role="Утв."))
    vp.changes.append(ChangeRecord(number=1, date="12.03.2026", description="Замена таблицы в ПЗ, л. 5", target="ПЗ, л.5", signer_surname="Иванов"))
    vp.appendices.append(Appendix(number="А", name="Спецификация оборудования", file_path="app_A.pdf", pages=3))
    vp.documents.append(SourceFile(path="ПЗ.docx", doc_kind=DocKind.POYASNITELNAYA))
    vp.save(os.path.join(os.path.dirname(__file__), "..", "tmp", "sample_project.json"))
    print("OK — sample project saved")
