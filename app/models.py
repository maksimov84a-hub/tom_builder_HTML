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



# ---- Перечень листов графической части ---------------------------------------

#: Источники перечня листов графической части.
SHEETLIST_AUTO = "auto"       # по фактическим листам раздела (из структуры тома)
SHEETLIST_EXCEL = "excel"     # из таблицы Excel (обозначение/наименование/примечание)
SHEETLIST_PDF = "pdf"         # из самого PDF: название и номер из штампа листа

#: Колонки перечня в порядке граф таблицы.
SHEETLIST_COLS = ("designation", "name", "note")
SHEETLIST_LABELS = ("Обозначение", "Наименование", "Примечание")
#: Ширины колонок перечня в мм: Обозначение | Наименование | Примечание.
#: Единый источник для предпросмотра, Word и PDF, чтобы таблица на листе
#: выглядела одинаково во всех выводах. Сумма (185 мм) равна ширине листа
#: между линиями рамки (210 − 20 − 5), поэтому таблица занимает всю рамку.
SHEETLIST_WIDTHS = (50.0, 95.0, 40.0)

#: Ширины колонок содержания в мм: Наименование | Обозначение | Листы.
#: Тот же набор ширин, что у перечня листов (95 мм — наименование, 50 —
#: обозначение, 40 — листы), сумма также 185 мм.
TOC_WIDTHS = (95.0, 50.0, 40.0)

# ---- Геометрия штампа листа в исходном PDF -----------------------------------
# Все размеры — от КРАЯ ЛИСТА. Штамп всегда в правом нижнем углу с отступом
# 5 мм от низа и 5 мм от правого края. По этой геометрии из PDF берутся
# название листа и его номер.
STAMP_EDGE_MM = 5.0

#: Название листа: влево 55..125 мм от правого края, вверх 5..20 мм от низа.
SHEET_NAME_BOX = (55.0, 125.0, 5.0, 20.0)      # (← от, ← до, ↑ от, ↑ до)
#: Номер листа: влево 25..40 мм от правого края, вверх 20..30 мм от низа.
SHEET_NO_BOX = (25.0, 40.0, 20.0, 30.0)

#: Шаблон обозначения листа для перечня: шифр тома + «л.» + номер листа.
#: {volume} — обозначение тома, {no} — номер листа из штампа.
SHEET_DESIGNATION_TEMPLATE = "{volume} л.{no}"


@dataclass
class GraphicSheetList:
    """Лист «Графическая часть» с перечнем листов (заголовок + таблица).

    Вставляется ПЕРЕД листами графической части, с рамкой и основной надписью
    формы 3. Таблица — во всю ширину листа, колонки:
    Обозначение | Наименование | Примечание.

    Источник строк:
      * pdf   — название и номер берутся из самого PDF по геометрии штампа
        (правый нижний угол). Обозначение собирается по шаблону
        «шифр + л. + номер». Лист, у которого название или номер в заданных
        квадратах не найдены, В ПЕРЕЧЕНЬ НЕ ПОПАДАЕТ;
      * auto  — по фактическим листам раздела (файлы графики), обозначение =
        обозначение раздела, наименование = имя файла;
      * excel — читается из указанного xlsx: первая строка — шапка, далее
        по строке на лист. Пустые строки пропускаются.
    """
    enabled: bool = True                  # вставлять лист в том
    source: str = SHEETLIST_PDF           # "pdf" | "auto" | "excel"
    excel_path: str = ""                  # путь к xlsx при source="excel"
    sheet_name: str = ""                  # имя листа книги ("" — первый лист)
    title: str = "Графическая часть"      # заголовок листа (крупно)
    subtitle: str = "Перечень листов графической части"
    #: Шаблон обозначения листа для перечня ({volume}, {no}).
    designation_template: str = SHEET_DESIGNATION_TEMPLATE
    #: Область названия листа: (влево от, влево до, вверх от, вверх до), мм.
    name_box: tuple = SHEET_NAME_BOX
    #: Область номера листа: (влево от, влево до, вверх от, вверх до), мм.
    number_box: tuple = SHEET_NO_BOX

    def to_dict(self) -> dict:
        d = asdict(self)
        # Кортежи областей сохраняются списками (JSON не знает кортежей).
        d["name_box"] = list(self.name_box)
        d["number_box"] = list(self.number_box)
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "GraphicSheetList":
        g = cls()
        d = d or {}
        g.enabled = bool(d.get("enabled", g.enabled))
        source = str(d.get("source") or g.source)
        g.source = (source if source in (SHEETLIST_PDF, SHEETLIST_AUTO,
                                         SHEETLIST_EXCEL) else SHEETLIST_PDF)
        g.excel_path = str(d.get("excel_path") or "")
        g.sheet_name = str(d.get("sheet_name") or "")
        g.title = str(d.get("title") or g.title)
        g.subtitle = str(d.get("subtitle") if d.get("subtitle") is not None else g.subtitle)
        g.designation_template = str(d.get("designation_template")
                                     or SHEET_DESIGNATION_TEMPLATE)
        g.name_box = _as_box(d.get("name_box"), SHEET_NAME_BOX)
        g.number_box = _as_box(d.get("number_box"), SHEET_NO_BOX)
        return g

    def use_excel(self) -> bool:
        """True — строки берутся из Excel (файл задан и существует)."""
        return (self.source == SHEETLIST_EXCEL and bool(self.excel_path)
                and os.path.exists(self.excel_path))

    def use_pdf(self) -> bool:
        """True — название и номер читаются из самого PDF листов."""
        return self.source == SHEETLIST_PDF


def _as_box(value, default: tuple) -> tuple:
    """Область из JSON: список/кортеж из 4 чисел (иначе — значение по умолчанию)."""
    try:
        box = tuple(float(v) for v in (value or ()))
    except (TypeError, ValueError):
        return default
    return box if len(box) == 4 else default


# ---- Состав проектной документации -----------------------------------------

#: Колонки таблицы «Состав проектной документации» в порядке граф.
COMPOSITION_COLS = ("volume", "designation", "name", "note")
COMPOSITION_LABELS = ("Номер тома", "Обозначение", "Наименование", "Примечание")
#: Ширины колонок в мм: Номер тома 15 | Обозначение 60 | Наименование 90 |
#: Примечание 20. Сумма 185 мм — ровно ширина между линиями рамки листа
#: (210 − 20 − 5), поэтому таблица занимает лист от рамки до рамки.
COMPOSITION_WIDTHS = (15.0, 60.0, 90.0, 20.0)
#: Высота шапки таблицы и высота строки — в мм.
COMPOSITION_HEAD_MM = 15.0
COMPOSITION_ROW_MM = 8.0


@dataclass
class VolumeComposition:
    """Раздел «Состав проектной документации» — таблица томов (документов).

    Таблица заполняется вручную (строки хранятся в модели) либо импортом из
    книги Excel. Импорт КОПИРУЕТ строки в модель: после загрузки таблицу можно
    править вручную, а сборка тома не зависит от наличия файла Excel.
    """
    #: Строки таблицы: [{"volume": "", "designation": "", "name": "", "note": ""}].
    rows: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {"rows": [ {key: str((row or {}).get(key) or "")
                           for key in COMPOSITION_COLS}
                          for row in self.rows if isinstance(row, dict) ]}

    @classmethod
    def from_dict(cls, d: dict) -> "VolumeComposition":
        comp = cls()
        for row in (d or {}).get("rows") or []:
            if not isinstance(row, dict):
                continue
            comp.rows.append({key: str(row.get(key) or "")
                              for key in COMPOSITION_COLS})
        return comp

    def table_rows(self) -> list:
        """Строки для вывода: [(номер тома, обозначение, наименование, примечание)]."""
        return [tuple(str((row or {}).get(key) or "") for key in COMPOSITION_COLS)
                for row in self.rows]


def composition_rows(project) -> list:
    """Строки таблицы «Состав проектной документации» из модели проекта.

    Единый источник для предпросмотра, Word и PDF: таблица на листе и в
    собранном томе совпадает.
    """
    comp = getattr(project, "composition", None)
    if not isinstance(comp, VolumeComposition):
        return []
    return comp.table_rows()


# ---- Перечисления ----------------------------------------------------------

class Stage(str, Enum):
    """Стадия проектной документации."""
    P = "П"   # проектная
    R = "Р"   # рабочая


#: Полные наименования стадий (для обложки и титульного листа). В основной
#: надписи и содержании выводится обозначение стадии — «П» / «Р».
STAGE_LABELS = {
    Stage.P: "Проектная документация",
    Stage.R: "Рабочая документация",
}


def stage_label(value) -> str:
    """Полное наименование стадии по её значению («П» → «Проектная документация»)."""
    return STAGE_LABELS.get(as_stage(value), "")


def parse_stage(text: str):
    """Распознаёт стадию по обозначению («П», «Р») или полному наименованию.

    Принимает то, что пользователь может ввести в поле листа или выбрать в
    списке: «П», «п», «Проектная документация», «Проектная (П)».
    """
    s = (text or "").strip()
    if not s:
        return None
    low = s.lower()
    # Обозначение стадии (в т.ч. в скобках: «Проектная (П)»).
    for st in Stage:
        if low == st.value.lower() or f"({st.value.lower()})" in low:
            return st
    # Полное наименование либо его начало без второго слова:
    # «Проектная документация», «Рабочая документация», «Проектная».
    for st, label in STAGE_LABELS.items():
        if low == label.lower() or low.startswith(label.split()[0].lower()):
            return st
    return None


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
SECTION_COMPOSITION = "composition"           # Состав проектной документации

# Человекочитаемые заголовки разделов (для дерева и содержания).
SECTION_LABELS = {
    SECTION_COVER: "Обложка",
    SECTION_TITLE: "Титульный лист",
    SECTION_COMPOSITION: "Состав проектной документации",
    SECTION_CONTENT: "Содержание",
    SECTION_PZ: "Пояснительная записка",
    SECTION_GRAPHIC: "Графическая часть",
    SECTION_APPENDIX: "Приложения",
    SECTION_CHANGES: "Лист регистрации изменений",
}

#: Дополнительное обозначение раздела — дописывается к обозначению тома
#: («02-2026-ТР» + «.ПЗ» = «02-2026-ТР.ПЗ»). Задаётся для разделов, у которых
#: по ГОСТ есть собственное обозначение; значения по умолчанию можно менять
#: в контекстной панели (форма «Обозначение раздела»).
SECTION_CODE_DEFAULTS = {
    SECTION_CONTENT: ".С",
    SECTION_PZ: ".ПЗ",
    SECTION_GRAPHIC: ".ГЧ",
    # У состава ПД своего кода по умолчанию нет: поле есть, значение пустое.
    SECTION_COMPOSITION: "",
}


def section_designation(project, kind: str) -> str:
    """Обозначение раздела: обозначение тома + дополнительное обозначение.

    Для разделов без своего кода (обложка, титул, приложения, изменения)
    возвращается только обозначение тома.
    """
    base = (getattr(project.requisites, "volume_id", "") or "").strip()
    codes = getattr(project, "section_codes", None) or {}
    code = (codes.get(kind) or "").strip()
    if not base:
        return code
    return f"{base}{code}" if code else base


# ---- Таблица изменений основной надписи (графы 14–19) ------------------------

#: Колонки таблицы изменений в порядке граф ГОСТ: Изм. | Кол.уч | Лист |
#: №док. | Подп. | Дата. Те же ключи использует geometry.CHANGE_KEYS.
STAMP_CHANGE_COLS = ("izm", "uch", "list", "doc", "sign", "date")
STAMP_CHANGE_LABELS = ("Изм.", "Кол.уч", "Лист", "№док.", "Подп.", "Дата")


def _empty_stamp_table() -> dict:
    """Пустая таблица изменений раздела: строки + даты подписей."""
    return {"rows": [], "sign_dates": {}}


def stamp_change_table(project, kind: str) -> dict:
    """Таблица изменений раздела (создаёт пустую, если её ещё нет)."""
    tables = getattr(project, "stamp_changes", None)
    if tables is None:
        project.stamp_changes = {}
        tables = project.stamp_changes
    table = tables.get(kind)
    if not isinstance(table, dict):
        table = _empty_stamp_table()
        tables[kind] = table
    table.setdefault("rows", [])
    table.setdefault("sign_dates", {})
    return table


def stamp_change_rows(project, kind: str) -> list:
    """Строки таблицы изменений раздела: [{колонка: значение}, ...]."""
    return list(stamp_change_table(project, kind).get("rows") or [])


def stamp_sign_dates(project, kind: str) -> dict:
    """Даты подписей раздела: {роль: дата} (роли — develop/norm/approve)."""
    return dict(stamp_change_table(project, kind).get("sign_dates") or {})


def stamp_row_value(project, kind: str, row: int, col: str) -> str:
    """Значение одной ячейки таблицы изменений ('' — если не заполнено)."""
    rows = stamp_change_rows(project, kind)
    if 0 <= row < len(rows) and isinstance(rows[row], dict):
        return str(rows[row].get(col) or "")
    return ""


def set_stamp_row_value(project, kind: str, row: int, col: str, value: str) -> None:
    """Записывает значение ячейки таблицы изменений (строка создаётся при нужде)."""
    if col not in STAMP_CHANGE_COLS or row < 0:
        return
    rows = stamp_change_table(project, kind)["rows"]
    while len(rows) <= row:
        rows.append({})
    rows[row][col] = (value or "").strip()


def set_stamp_sign_date(project, kind: str, role: str, value: str) -> None:
    """Записывает дату подписи раздела (роль — develop/norm/approve)."""
    stamp_change_table(project, kind)["sign_dates"][role] = (value or "").strip()


# ---- Строки подписей основной надписи (Разраб./Пров./Н.контр./ГИП) ------------
# Все ячейки левого блока штампа заполняются прямо в предпросмотре: должность,
# фамилия и дата. Значения по умолчанию задаёт отрисовка формы (preview_html),
# а здесь хранятся только ИЗМЕНЁННЫЕ пользователем значения — поэтому смена
# подписанта в реестре продолжает обновлять незатронутые строки.

#: Поля строки подписи.
STAMP_SIG_FIELDS = ("post", "surname", "date")

#: Строки подписей левого блока штампа: (y_mm от верха штампа, должность по
#: умолчанию, роль подписи). Роль связывает строку с реестром подписантов
#: (`signature_assignments`); пустая роль — свободная строка: её должность и
#: фамилию заполняют вручную. Умолчания совпадают с прежним видом штампа,
#: поэтому добавление правки не меняет уже настроенные листы.
#: Координаты согласованы с сетками форм (geometry.F5_HLINES / F3_HLINES).
STAMP_SIG_ROWS = {
    5: [(15.0, "Разраб.", "develop"), (20.0, "", ""), (25.0, "", ""),
        (30.0, "Н.контр.", "norm"), (35.0, "ГИП", "approve")],
    3: [(15.0, "Разраб.", "develop"), (20.0, "", ""), (25.0, "", ""),
        (30.0, "Н.контр.", "norm"), (35.0, "ГИП.", "approve")],
}

#: Связь строк формы 5 (ПЗ, содержание) со строками формы 3 (графическая часть).
#: Ключ — «слот» строки: одно и то же значение данных.
#:
#: У формы 3 (по образцу) пять строк подписей, как и у формы 5, поэтому строки
#: совпадают один в один. Карта оставлена отдельной: если у графической части
#: появится своя раскладка, менять придётся только её.
STAMP_SIG_SLOTS = {
    5: {0: 1, 1: 2, 2: 3, 3: 4, 4: 5},      # строка формы 5 -> слот
    3: {0: 1, 1: 2, 2: 3, 3: 4, 4: 5},      # строка формы 3 -> слот
}

#: Должность в строке подписи по умолчанию — по её роли (для форм 5 и 3).
#: Используется при заполнении строки «с чистого листа» (например, в Word/PDF).
STAMP_ROLE_POSTS = {"develop": "Разраб.", "check": "Пров.",
                    "norm": "Н.контр.", "approve": "ГИП"}


def stamp_signature_rows(project, kind: str, form: int) -> list:
    """Строки подписей основной надписи для формы 5 (текст) или 3 (графика).

    Значения по умолчанию — должность из сетки формы и фамилия из реестра
    подписантов (по роли строки). Правки, сделанные в предпросмотре, хранятся
    ОДИН раз на «слот строки» (`project.stamp_signatures`) и НЕ зависят от
    раздела: должности, фамилии, подписи и даты в штампах содержания,
    пояснительной записки и графической части заполняются одинаково. Правка в
    любом штампе сразу видна во всех остальных.

    Исключение — разделы из `STAMP_SIG_INDEPENDENT` (состав ПД): у них свой
    набор правок и НЕТ наследования от общего штампа, в том числе фамилий из
    реестра по роли. Такой штамп заполняется независимо.

    :param kind: раздел тома. Определяет набор правок: у разделов из
        `STAMP_SIG_INDEPENDENT` он свой, у остальных — общий. Также используется
        для дат подписей старого формата (`stamp_changes[kind]["sign_dates"]`).

    В строке также отдаётся `signature_image` — изображение подписи подписанта,
    найденного по фамилии (для Word/PDF и предпросмотра).
    """
    layout = STAMP_SIG_ROWS.get(5 if form != 3 else 3, STAMP_SIG_ROWS[5])
    overrides = stamp_signature_overrides(project, kind)
    assignments = getattr(project, "signature_assignments", None) or {}
    legacy_dates = stamp_sign_dates(project, kind)
    signers = getattr(project, "signers", None) or []
    # Независимый штамп не берёт фамилии из общего реестра ролей: иначе правка
    # подписанта в другом разделе меняла бы и лист состава ПД.
    independent = kind in STAMP_SIG_INDEPENDENT
    # Строка адресуется «слотом», а не номером: у формы 3 «Н.контр.» и «ГИП»
    # стоят на строках 5 и 6, но несут те же данные, что строки 4 и 5 формы 5.
    slots = STAMP_SIG_SLOTS.get(5 if form != 3 else 3) or {}
    rows = []
    for index, (y, post, role) in enumerate(layout):
        slot = slots.get(index, 0)
        row = {
            "row": index,
            "slot": slot,
            "y": y,
            "role": role,
            "post": post,
            "surname": ("" if independent else
                        (assignments.get(role, "") if role else "")),
            "date": ("" if independent else
                     (legacy_dates.get(role, "") if role else "")),
        }
        if slot:
            row.update(overrides.get(slot) or {})
        # Подпись подбирается по фамилии в графе, а не по роли: правка фамилии
        # в предпросмотре сразу меняет и изображение подписи.
        row["signature_image"] = signer_signature_image(signers, row["surname"])
        rows.append(row)
    return rows


#: Ключ «общего» набора подписей: строки не зависят от раздела, поэтому
#: правки хранятся под одним ключом и видны во всех штампах.
STAMP_SIG_SHARED = "*"

#: Разделы с НЕЗАВИСИМЫМ штампом: строки подписей (должность, фамилия, дата)
#: хранятся отдельно для раздела и не связаны с общими штампами тома. У состава
#: проектной документации свой штамп: его правки не переносятся в содержание,
#: записку и графическую часть и наоборот.
STAMP_SIG_INDEPENDENT = (SECTION_COMPOSITION,)


def stamp_signature_scope(kind: str | None) -> str:
    """Ключ хранения строк подписей раздела: собственный или общий."""
    return str(kind) if kind in STAMP_SIG_INDEPENDENT else STAMP_SIG_SHARED


def signer_signature_image(signers, surname: str) -> str:
    """Изображение подписи подписанта с указанной фамилией ('' — если нет)."""
    target = (surname or "").strip().lower()
    if not target:
        return ""
    for signer in signers or []:
        if (signer.surname or "").strip().lower() == target:
            return signer.signature_image or ""
    return ""


def stamp_signature_by_role(project, kind: str, form: int, role: str) -> dict | None:
    """Строка подписи, назначенная роли (для обратной совместимости ключей)."""
    for row in stamp_signature_rows(project, kind, form):
        if row["role"] == role:
            return row
    return None


def stamp_signature_overrides(project, kind: str | None = None) -> dict:
    """Изменённые строки подписей: {номер строки: {поле: значение}}.

    Правки ОБЩИЕ для всех разделов (ключ `STAMP_SIG_SHARED`): штампы
    содержания, пояснительной записки и графической части заполняются
    одинаково, и правка в одном сразу применяется ко всем. Исключение —
    разделы из `STAMP_SIG_INDEPENDENT`: у них собственный набор подписей.

    :param kind: раздел. Оставлен для обратной совместимости: старые проекты
        хранили правки по разделам, поэтому если общего набора ещё нет, они
        подхватываются из любых сохранённых разделов.
    """
    tables = getattr(project, "stamp_signatures", None) or {}
    scope = stamp_signature_scope(kind)
    table = tables.get(scope)
    if not isinstance(table, dict):
        if scope != STAMP_SIG_SHARED:
            # Независимый раздел: чужие правки не наследуются.
            return {}
        # Старый проект: правки лежали по разделам. Переносим их в общий набор
        # (первый найденный раздел — приоритет у явно запрошенного kind).
        # Наборы независимых разделов не наследуются: у них свой штамп.
        order = []
        if kind:
            order.append(kind)
        order.extend(k for k in tables
                     if k != STAMP_SIG_SHARED and k not in STAMP_SIG_INDEPENDENT)
        table = None
        for key in order:
            candidate = tables.get(key)
            if isinstance(candidate, dict) and candidate:
                table = candidate
                break
        if not isinstance(table, dict):
            return {}
    result = {}
    for row, fields in table.items():
        try:
            index = int(row)
        except (TypeError, ValueError):
            continue
        if index < 0 or not isinstance(fields, dict):
            continue
        result[index] = {f: str(fields.get(f, "") or "")
                         for f in STAMP_SIG_FIELDS if f in fields}
    return result


def stamp_signature_row(project, kind: str, row: int, *,
                        post: str = "", surname: str = "", date: str = "") -> dict:
    """Значения строки подписи: значения по умолчанию с наложенными правками.

    :param post/surname/date: значения по умолчанию (из отрисовки формы).
    """
    overrides = stamp_signature_overrides(project, kind).get(row) or {}
    result = {"post": post, "surname": surname, "date": date}
    result.update(overrides)
    return result


def stamp_signature_slot(form: int, row: int) -> int:
    """Слот строки подписи: строка формы -> общий слот данных.

    0 — строка не участвует в обмене (у формы 3 четвёртая строка свободна).
    Слот и есть ключ хранения правок: одна и та же правка действует и в
    форме 5 (записка, содержание), и в форме 3 (графика).
    """
    slots = STAMP_SIG_SLOTS.get(5 if form != 3 else 3) or {}
    return int(slots.get(int(row), 0))


def set_stamp_signature(project, kind: str, row: int, field: str, value: str,
                        form: int | None = None) -> None:
    """Записывает ячейку строки подписи (должность/фамилия/дата).

    Значение сохраняется в наборе по «слоту» строки: у разделов из
    `STAMP_SIG_INDEPENDENT` — в собственном наборе раздела, у остальных — в
    ОБЩЕМ. Правка в штампе общего раздела сразу действует во всех остальных
    общих штампах, а строки формы 5 и формы 3 связываются по схеме
    `STAMP_SIG_SLOTS` (1-1, 2-2, 3-3, 4-5, 5-6). Независимый штамп состава ПД
    не связан с общими.

    :param form: форма, в которой сделана правка (5 или 3). Нужна, чтобы
        перевести номер строки в слот: у формы 3 «Н.контр.» — строка 5, у
        формы 5 — строка 4. None — считается форма 5.
    """
    if field not in STAMP_SIG_FIELDS or row < 0:
        return
    slot = stamp_signature_slot(5 if not form else form, row)
    if not slot:
        # Свободная строка формы 3, не связанная со строками записки: правка
        # не сохраняется (в записке ей нечего соответствовать).
        return
    tables = getattr(project, "stamp_signatures", None)
    if not isinstance(tables, dict):
        project.stamp_signatures = {}
        tables = project.stamp_signatures
    scope = stamp_signature_scope(kind)
    if scope == STAMP_SIG_SHARED and STAMP_SIG_SHARED not in tables:
        # Правки старого формата (по разделам) переносим в общий набор, чтобы
        # далее источником был только он.
        inherited = stamp_signature_overrides(project, kind)
        shared = {}
        for index, fields in inherited.items():
            shared[str(index)] = dict(fields)
        tables[STAMP_SIG_SHARED] = shared
    table = tables.get(scope)
    if not isinstance(table, dict):
        table = {}
        tables[scope] = table
    fields = table.get(str(slot))
    if not isinstance(fields, dict):
        fields = {}
        table[str(slot)] = fields
    fields[field] = (value or "").strip()


def default_structure() -> list:
    """Стандартный порядок блоков тома по ГОСТ (по умолчанию)."""
    return [
        {"kind": SECTION_COVER, "needed": True},
        {"kind": SECTION_TITLE, "needed": True},
        {"kind": SECTION_CONTENT, "needed": True},
        {"kind": SECTION_COMPOSITION, "needed": True},
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
    # Дополнительное обозначение раздела: {kind: ".ПЗ"} — дописывается к
    # обозначению тома в графе 1 основной надписи и в содержании.
    section_codes: dict = field(default_factory=lambda: dict(SECTION_CODE_DEFAULTS))
    # Таблица изменений основной надписи (графы 14–19) и даты подписей по
    # разделам: {kind: {"rows": [{"izm": "", ...}], "sign_dates": {"develop": ""}}}.
    # Значения заполняются прямо в предпросмотре (см. preview_html).
    stamp_changes: dict = field(default_factory=dict)
    #Изменённые строки подписей штампа по разделам:
    # {kind: {"0": {"post": "", "surname": "", "date": ""}}}.
    # Хранятся только правки — значения по умолчанию берутся из формы и реестра.
    stamp_signatures: dict = field(default_factory=dict)
    # Лист «Графическая часть» с перечнем листов перед графической частью
    # (рамка + основная надпись формы 3, таблица во всю ширину листа).
    graphic_sheet_list: GraphicSheetList = field(default_factory=GraphicSheetList)
    # Раздел «Состав проектной документации»: таблица томов (номер, обозначение,
    # наименование, примечание). Листы — форма 5 (первый) и форма 6 (последующие).
    composition: VolumeComposition = field(default_factory=VolumeComposition)

    def section_code(self, kind: str) -> str:
        return (self.section_codes or {}).get(kind, "")

    def designation(self, kind: str) -> str:
        """Обозначение раздела: том + дополнительное обозначение раздела."""
        return section_designation(self, kind)

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
            "section_codes": dict(self.section_codes or {}),
            "stamp_changes": {
                str(k): {
                    "rows": [dict(r) for r in (v or {}).get("rows") or []],
                    "sign_dates": dict((v or {}).get("sign_dates") or {}),
                }
                for k, v in (self.stamp_changes or {}).items()
            },
            "stamp_signatures": {
                str(k): {str(r): dict(f) for r, f in (v or {}).items()
                         if isinstance(f, dict)}
                for k, v in (self.stamp_signatures or {}).items()
                if isinstance(v, dict)
            },
            "graphic_sheet_list": self.graphic_sheet_list.to_dict(),
            "composition": self.composition.to_dict(),
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
        # Проекты, сохранённые до появления раздела «Состав проектной
        # документации», не содержат его в структуре — без переноса раздел не
        # появился бы в дереве и его нельзя было бы заполнить. Вставляем его
        # сразу после «Содержания» (как в структуре по умолчанию: содержание —
        # первый лист тома, состав — второй); исключить раздел из сборки можно
        # флажком в дереве.
        if not any(isinstance(e, dict) and e.get("kind") == SECTION_COMPOSITION
                   for e in vp.structure):
            def _index_of(kind):
                return next((i for i, e in enumerate(vp.structure)
                             if isinstance(e, dict) and e.get("kind") == kind), None)
            content_at = _index_of(SECTION_CONTENT)
            title_at = _index_of(SECTION_TITLE)
            if content_at is not None:
                after = content_at + 1
            elif title_at is not None:
                after = title_at + 1
            else:
                after = 0
            vp.structure.insert(after,
                                {"kind": SECTION_COMPOSITION, "needed": True})
        vp.stamp_cell9 = StampCell9.from_dict(d.get("stamp_cell9") or {})
        # Старые проекты не содержат section_codes — берём значения по умолчанию,
        # а сохранённые коды (в т.ч. намеренно пустые) уважаем.
        vp.section_codes = dict(SECTION_CODE_DEFAULTS)
        vp.section_codes.update({
            str(k): str(v) for k, v in (d.get("section_codes") or {}).items()
        })
        # Таблица изменений штампа: строки и даты подписей по разделам.
        vp.stamp_changes = {}
        for kind, table in (d.get("stamp_changes") or {}).items():
            if not isinstance(table, dict):
                continue
            vp.stamp_changes[str(kind)] = {
                "rows": [dict(r) for r in (table.get("rows") or [])
                         if isinstance(r, dict)],
                "sign_dates": {str(k): str(v)
                               for k, v in (table.get("sign_dates") or {}).items()},
            }
        # Правки строк подписей штампа: только изменённые поля.
        vp.stamp_signatures = {}
        for kind, table in (d.get("stamp_signatures") or {}).items():
            if not isinstance(table, dict):
                continue
            vp.stamp_signatures[str(kind)] = {
                str(row): {str(f): str(v) for f, v in fields.items()
                           if f in STAMP_SIG_FIELDS}
                for row, fields in table.items() if isinstance(fields, dict)
            }
        vp.graphic_sheet_list = GraphicSheetList.from_dict(
            d.get("graphic_sheet_list") or {})
        vp.composition = VolumeComposition.from_dict(d.get("composition") or {})
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
