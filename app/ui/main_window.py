# -*- coding: utf-8 -*-
"""Главное окно программы: HTML-предпросмотр листов + контекстная панель настроек.

Компоновка:
  слева  — дерево «Листы тома» (структура тома, SheetTree);
  центр  — HTML-предпросмотр выбранного листа (SheetPreview);
  справа — контекстная панель настроек (ContextPanel): показывает настройки,
           относящиеся к выбранному пункту. Реестр подписантов — отдельная
           кнопка «Подписанты…».

Предпросмотр — главный интерфейс: поля редактируются inline, общие реквизиты
и подписи обновляются во всех листах.
"""
from __future__ import annotations

import json
import os

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout,
    QFileDialog, QMessageBox, QSplitter, QDialog,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QKeySequence

from ..models import VolumeProject
from ..core.preview_model import build_preview_snapshot, signer_for_role
from .context_panel import ContextPanel
from .sheet_tree import SheetTree
from .preview_widget import SheetPreview
from .signature_picker import SignaturePickerDialog


class MainWindow(QMainWindow):
    """Главное окно: предпросмотр (HTML) + контекстная панель настроек."""

    def __init__(self, project: VolumeProject | None = None):
        super().__init__()
        self.project = project or VolumeProject()
        self._current_file: str | None = None   # путь к JSON-проекту
        self._current_sheet: dict = {"kind": "cover"}   # текущий лист предпросмотра
        self.setWindowTitle("Сборка тома по ГОСТ Р 21.101-2020")

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # ---- Левая панель: дерево листов (структура тома) ----
        self.sheet_tree = SheetTree()
        self.sheet_tree.selection_changed.connect(self._on_sheet_selected)

        # ---- Центр: HTML-предпросмотр ----
        self.preview = SheetPreview()
        self.preview.field_changed.connect(self._on_field_changed)
        self.preview.signature_clicked.connect(self._on_signature_clicked)
        self.preview.save_pz_requested.connect(self._save_pz_edits)

        # ---- Правая панель: контекстные настройки (вместо вкладок) ----
        self.context = ContextPanel(project=self.project)
        self.context.item_modified.connect(self._on_item_modified)

        # ---- Компоновка ----
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.sheet_tree)
        splitter.addWidget(self.preview)
        splitter.addWidget(self.context)
        self.context.setMinimumWidth(300)
        splitter.setSizes([320, 880, 360])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 0)
        splitter.setCollapsible(0, False)
        splitter.setCollapsible(2, False)
        root.addWidget(splitter, stretch=1)

        # ---- Главное меню (кнопки проекта вынесены наверх как меню) ----
        self._build_menu()

        self._load_project_into_forms()

        # Окно на весь экран (максимизировано при первом показе)
        self._was_maximized = False

    def showEvent(self, event) -> None:
        """При первом показе — окно на весь экран."""
        super().showEvent(event)
        if not self._was_maximized:
            self._was_maximized = True
            self.showMaximized()

    def _confirm_pending_pz_edits(self, action: str) -> bool:
        """Спрашивает о несохранённых правках ПЗ перед действием.

        Возвращает True, если можно продолжать (правок нет или пользователь
        решил продолжить). Если пользователь выбрал «Сохранить» — правки
        записываются, и возвращается True при успехе.
        """
        if not self.preview.has_pz_edits():
            return True
        answer = QMessageBox.question(
            self, "Несохранённые правки ПЗ",
            f"В пояснительной записке есть несохранённые правки.\n\n"
            f"Сохранить их перед тем, как {action}?",
            QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
            QMessageBox.Save)
        if answer == QMessageBox.Cancel:
            return False
        if answer == QMessageBox.Save:
            self._save_pz_edits()
            # Если сохранение не удалось, правки остались — не продолжаем.
            return not self.preview.has_pz_edits()
        # Discard — правки отбрасываются.
        self.preview.clear_pz_edits()
        return True

    def closeEvent(self, event) -> None:
        """Не даём потерять правки ПЗ, сделанные в предпросмотре."""
        if not self._confirm_pending_pz_edits("закрыть программу"):
            event.ignore()
            return
        super().closeEvent(event)

    # ---- Главное меню ----
    def _build_menu(self) -> None:
        """Строит главное меню окна. Кнопки «Новый/Открыть/Сохранить/
        Сохранить как» вынесены наверх как пункты меню «Файл»."""
        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu("&Файл")

        # Пункты меню из бывших кнопок.
        act_new = QAction("&Новый", self)
        act_new.setShortcut(QKeySequence.StandardKey.New)
        act_new.triggered.connect(self._new)
        file_menu.addAction(act_new)

        act_open = QAction("&Открыть…", self)
        act_open.setShortcut(QKeySequence.StandardKey.Open)
        act_open.triggered.connect(self._open)
        file_menu.addAction(act_open)

        file_menu.addSeparator()

        act_save = QAction("&Сохранить", self)
        act_save.setShortcut(QKeySequence.StandardKey.Save)
        act_save.triggered.connect(self._save)
        file_menu.addAction(act_save)

        act_save_as = QAction("Сохранить &как…", self)
        act_save_as.setShortcut(QKeySequence.StandardKey.SaveAs)
        act_save_as.triggered.connect(self._save_as)
        file_menu.addAction(act_save_as)

        file_menu.addSeparator()

        act_quit = QAction("&Выход", self)
        act_quit.setShortcut(QKeySequence.StandardKey.Quit)
        act_quit.triggered.connect(self.close)
        file_menu.addAction(act_quit)

        # Меню «Недавние» — отдельный пункт в строке меню (до 5 записей).
        self._recent_menu = menu_bar.addMenu("&Недавние")
        self._rebuild_recent_menu()

        # Пункты сборки — отдельными элементами в строке меню (без подменю).
        act_build = QAction("&Собрать том…", self)
        act_build.triggered.connect(self._build)
        menu_bar.addAction(act_build)

        act_build_html = QAction("Собрать том в &HTML…", self)
        act_build_html.triggered.connect(self._build_html_volume)
        menu_bar.addAction(act_build_html)

        act_build_pdf = QAction("Экспорт HTML → &PDF…", self)
        act_build_pdf.triggered.connect(self._build_html_pdf_volume)
        menu_bar.addAction(act_build_pdf)

        # Сборка раздела — только через меню (кнопки из панелей убраны).
        act_build_section = QAction("&Собрать раздел…", self)
        act_build_section.triggered.connect(self._on_build_section_current)
        menu_bar.addAction(act_build_section)

        act_build_section_html = QAction("Собрать раздел в &HTML…", self)
        act_build_section_html.triggered.connect(self._build_html_section)
        menu_bar.addAction(act_build_section_html)

        # Редактирование пояснительной записки — из окна программы.
        act_edit_pz = QAction("&Редактировать пояснительную записку…", self)
        act_edit_pz.triggered.connect(self._edit_current_pz)
        menu_bar.addAction(act_edit_pz)

        self._actions = {
            "new": act_new, "open": act_open,
            "save": act_save, "save_as": act_save_as,
            "build": act_build,
        }

    # ---- Недавние файлы ----
    RECENT_MAX = 5
    # Файл со списком недавних проектов храним в каталоге пользователя
    # (надёжно для Windows, в отличие от реестра через QSettings).
    RECENT_FILENAME = "recent_projects.json"

    def _recent_store_path(self) -> str:
        """Путь к файлу со списком недавних проектов (в каталоге пользователя)."""
        config_dir = os.path.join(os.path.expanduser("~"), ".tom_builder")
        return os.path.join(config_dir, self.RECENT_FILENAME)

    def _recent_files(self) -> list[str]:
        """Возвращает список путей к недавно открытым проектам (свежие первыми)."""
        path = self._recent_store_path()
        if not os.path.exists(path):
            return []
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return []
        paths = data if isinstance(data, list) else []
        return [p for p in paths if isinstance(p, str) and p]

    def _set_recent_files(self, paths: list[str]) -> None:
        """Сохраняет список недавних проектов (не более RECENT_MAX)."""
        # Держим максимум RECENT_MAX, отбрасываем дубликаты (свежие первыми).
        cleaned: list[str] = []
        for p in paths:
            if p and p not in cleaned:
                cleaned.append(p)
        cleaned = cleaned[:self.RECENT_MAX]
        store = self._recent_store_path()
        try:
            os.makedirs(os.path.dirname(store), exist_ok=True)
            with open(store, "w", encoding="utf-8") as f:
                json.dump(cleaned, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _add_recent_file(self, path: str) -> None:
        """Добавляет путь в начало списка недавних и обновляет меню."""
        if not path:
            return
        paths = self._recent_files()
        if path in paths:
            paths.remove(path)
        paths.insert(0, path)
        self._set_recent_files(paths)
        self._rebuild_recent_menu()

    def _rebuild_recent_menu(self) -> None:
        """Перестраивает пункты подменю «Недавние» из сохранённого списка."""
        menu = getattr(self, "_recent_menu", None)
        if menu is None:
            return
        menu.clear()
        paths = self._recent_files()

        if not paths:
            empty = QAction("(пусто)", self)
            empty.setEnabled(False)
            menu.addAction(empty)
            return

        for path in paths:
            act = QAction(os.path.basename(path), self)
            act.setToolTip(path)
            # Показываем и имя файла, и полный путь.
            act.setText(f"{os.path.basename(path)}")
            act.triggered.connect(lambda _=False, p=path: self._open_recent(p))
            menu.addAction(act)
        menu.addSeparator()
        clear = QAction("Очистить недавние", self)
        clear.triggered.connect(self._clear_recent_files)
        menu.addAction(clear)

    def _open_recent(self, path: str) -> None:
        """Открывает проект из списка «Недавние»."""
        if not os.path.exists(path):
            QMessageBox.warning(
                self, "Не найдено",
                f"Файл не найден:\n{path}\n\nОн будет удалён из списка «Недавние».")
            self._remove_recent_file(path)
            return
        self._collect_forms()
        try:
            self.project = VolumeProject.load(path)
            self._current_file = path
            self._load_project_into_forms()
            self._add_recent_file(path)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть проект:\n{e}")
            self._remove_recent_file(path)

    def _remove_recent_file(self, path: str) -> None:
        """Удаляет путь из списка недавних (при отсутствии файла)."""
        paths = self._recent_files()
        if path in paths:
            paths.remove(path)
            self._set_recent_files(paths)
            self._rebuild_recent_menu()

    def _clear_recent_files(self) -> None:
        """Очищает список недавних файлов."""
        self._set_recent_files([])
        self._rebuild_recent_menu()

    # ---- Перенос данных формы <-> модель ----
    def _load_project_into_forms(self) -> None:
        # Правки ПЗ предыдущего проекта к новому отношения не имеют.
        self.preview.clear_pz_edits()
        self.context.load_project(self.project)
        self.sheet_tree.set_project(self.project)
        self._refresh_preview()

    def _collect_forms(self) -> None:
        self.context.collect(self.project)

    # ---- Предпросмотр ----
    def _refresh_preview(self) -> None:
        snapshot = build_preview_snapshot(self.project)
        # Клеймим снимок метаданными текущего листа (для заглушек документов/приложений).
        if self._current_sheet.get("kind") in ("doc", "appendix", "section"):
            snapshot["label"] = self._current_sheet.get("label", "")
            snapshot["section_label"] = self._current_sheet.get("section_label", "")
            snapshot["path"] = self._current_sheet.get("path", "")
            snapshot["section"] = self._current_sheet.get("section", "")
        # Показываются ВСЕ листы выбранного пункта (документ/раздел — постранично).
        self.preview.show_sheet(self._current_sheet.get("kind", "cover"),
                                snapshot, project=self.project)

    def _on_sheet_selected(self, info: dict) -> None:
        self._current_sheet = info
        self._refresh_preview()
        # Показываем настройки, относящиеся к выбранному пункту.
        self.context.show_context(info)

    def _on_item_modified(self) -> None:
        """Параметр документа/приложения (needed) изменён — обновляем предпросмотр и дерево."""
        self._refresh_preview()
        self.sheet_tree.set_project(self.project)

    # ---- Редактирование полей ----
    def _on_field_changed(self, key: str, value: str) -> None:
        """Поле отредактировано inline. Обновляем модель и формы (без перезагрузки
        текущего листа, чтобы не прерывать ввод). Другие листы перерисуются,
        когда пользователь переключится на них (из свежего снимка данных)."""
        req = self.project.requisites
        req_map = {
            "organization": "organization",
            "volume_id": "volume_id",
            "object_name": "object_name",
            "volume_name": "volume_name",
            "stage": "stage",
            "city": "city",
            "year": "year",
        }
        if key in req_map:
            attr = req_map[key]
            if attr == "stage":
                from ..models import Stage
                try:
                    req.stage = Stage(value)
                except Exception:
                    pass
            elif attr == "year":
                try:
                    req.year = int(value)
                except Exception:
                    pass
            else:
                setattr(req, attr, value)
            # Синхронизируем форму реквизитов (чтобы не было расхождений).
            # Настройки графы 9 (текст/логотип) не трогаем — иначе правка
            # наименования организации сбрасывала бы выбранный логотип.
            self.context.requisites.sync_requisites(req)
            return

        # Прочие ключи (ch_* — ячейки таблицы изменений) пока не сохраняем.

    def _on_signature_clicked(self, key: str) -> None:
        """Клик по полю подписи — выбор подписанта."""
        from ..core.preview_model import role_label, build_preview_snapshot
        # Сначала синхронизируем формы с моделью (чтобы реестр подписантов был актуален)
        self.context.signers_form.set_signers(self.project.signers)
        self._collect_forms()
        snapshot = build_preview_snapshot(self.project)
        role_label_text = role_label(key)
        current = signer_for_role(self.project, key)
        current_surname = current.surname if current else ""
        dlg = SignaturePickerDialog(self, role_label=role_label_text,
                                    signers=snapshot["signers"],
                                    current_surname=current_surname)
        if dlg.exec() == SignaturePickerDialog.Accepted:
            if dlg.result_surname:
                self.project.signature_assignments[key] = dlg.result_surname
            else:
                self.project.signature_assignments.pop(key, None)
            self._refresh_preview()

    # ---- Файловые операции ----
    def _new(self) -> None:
        if not self._confirm_pending_pz_edits("создать новый проект"):
            return
        self._collect_forms()
        self.project = VolumeProject()
        self._load_project_into_forms()
    def _open(self) -> None:
        if not self._confirm_pending_pz_edits("открыть другой проект"):
            return
        self._collect_forms()
        f, _ = QFileDialog.getOpenFileName(self, "Открыть проект тома", "",
                                           "Проект тома (*.json);;Все файлы (*)")
        if not f:
            return
        try:
            self.project = VolumeProject.load(f)
            self._current_file = f
            self._load_project_into_forms()
            self._add_recent_file(f)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть проект:\n{e}")

    def _save(self) -> None:
        if not getattr(self, "_current_file", None):
            self._save_as()
            return
        self._collect_forms()
        try:
            self.project.save(self._current_file)
            QMessageBox.information(self, "Сохранено", "Проект сохранён.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить:\n{e}")

    def _save_as(self) -> None:
        self._collect_forms()
        f, _ = QFileDialog.getSaveFileName(self, "Сохранить проект тома", "том.json",
                                           "Проект тома (*.json)")
        if not f:
            return
        self._current_file = f
        try:
            self.project.save(f)
            self._add_recent_file(f)
            QMessageBox.information(self, "Сохранено", "Проект сохранён.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить:\n{e}")

    def _build(self) -> None:
        """Собрать весь том: и Word (DOCX), и PDF."""
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        if not self._confirm_pending_pz_edits("собирать том"):
            return
        self._collect_forms()
        # Запрашиваем базовое имя файла; DOCX и PDF создаются рядом.
        base = QFileDialog.getSaveFileName(self, "Сохранить том", "том.docx",
                                           "Документ Word (*.docx)")[0]
        if not base:
            return
        try:
            from ..config import load_config
            cfg = load_config()

            docx_out = base if base.lower().endswith(".docx") else base + ".docx"
            pdf_out = os.path.splitext(docx_out)[0] + ".pdf"

            from ..core.docx_exporter import build_text_docx
            build_text_docx(self.project, docx_out, cfg)

            full_pdf = None
            try:
                from ..core.pdf_builder import build_full_pdf
                full_pdf = build_full_pdf(self.project, pdf_out, cfg)
            except Exception as pe:
                full_pdf = None

            msg = f"Том собран в Word (DOCX):\n{docx_out}"
            if full_pdf:
                msg += f"\n\nТом собран в PDF:\n{full_pdf}"
            else:
                msg += "\n\nНе удалось собрать PDF."
            QMessageBox.information(self, "Готово", msg)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось собрать том:\n{e}")

    def _on_build_section(self, info: dict) -> None:
        """Собрать выбранный раздел тома целиком (Word / PDF)."""
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        if not self._confirm_pending_pz_edits("собирать раздел"):
            return
        self._collect_forms()
        section = info.get("section", "")
        label = info.get("label", section)
        # Диалог выбора формата и файла.
        out, selected_filter = QFileDialog.getSaveFileName(
            self, f"Собрать раздел «{label}»", f"{section}.docx",
            "Документ Word (*.docx);;PDF (*.pdf)")
        if not out:
            return
        try:
            from ..config import load_config
            cfg = load_config()
            if out.lower().endswith(".pdf"):
                from ..core.pdf_builder import build_section_pdf
                build_section_pdf(self.project, section, out, cfg)
                QMessageBox.information(self, "Готово",
                                        f"Раздел «{label}» собран в PDF:\n{out}")
            else:
                from ..core.docx_exporter import build_section_docx
                build_section_docx(self.project, section, out, cfg)
                QMessageBox.information(self, "Готово",
                                        f"Раздел «{label}» собран в Word:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось собрать раздел:\n{e}")

    def _on_build_section_current(self) -> None:
        """«Собрать раздел» (пункт меню): собрать раздел текущего листа."""
        info = self._section_info_for_current()
        if not info:
            return
        self._on_build_section(info)

    def _build_html_volume(self) -> None:
        """«Собрать» в HTML: собрать весь том в один HTML-файл."""
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        if not self._confirm_pending_pz_edits("собирать том"):
            return
        self._collect_forms()
        out = QFileDialog.getSaveFileName(self, "Собрать том в HTML", "том.html",
                                          "HTML (*.html)")[0]
        if not out:
            return
        try:
            from ..core.preview_html import build_volume_html
            build_volume_html(self.project, out)
            QMessageBox.information(self, "Готово", f"Том собран в HTML:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось собрать том в HTML:\n{e}")

    def _build_html_pdf_volume(self) -> None:
        """Экспорт всего тома по цепочке данные → HTML → PDF."""
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        if not self._confirm_pending_pz_edits("экспортировать том"):
            return
        self._collect_forms()
        out = QFileDialog.getSaveFileName(
            self, "Экспорт HTML → PDF", "том.pdf", "PDF (*.pdf)")[0]
        if not out:
            return
        if not out.lower().endswith(".pdf"):
            out += ".pdf"
        try:
            from ..core.html_export import build_volume_html_pdf
            build_volume_html_pdf(self.project, out)
            QMessageBox.information(
                self, "Готово",
                f"Том экспортирован по цепочке HTML → PDF:\n{out}\n\n"
                f"Промежуточный HTML сохранён рядом с PDF.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка HTML → PDF", f"Не удалось экспортировать том:\n{e}")

    def _build_html_section(self) -> None:
        """«Собрать раздел» в HTML: собрать текущий раздел в один HTML-файл."""
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        if not self._confirm_pending_pz_edits("собирать раздел"):
            return
        self._collect_forms()
        info = self._section_info_for_current()
        if not info:
            return
        section = info.get("section", "")
        label = info.get("label", section)
        out = QFileDialog.getSaveFileName(self, f"Собрать раздел «{label}» в HTML",
                                          f"{section}.html", "HTML (*.html)")[0]
        if not out:
            return
        try:
            from ..core.preview_html import build_section_html
            build_section_html(self.project, section, out)
            QMessageBox.information(self, "Готово",
                                    f"Раздел «{label}» собран в HTML:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось собрать раздел в HTML:\n{e}")

    # ---- Правка текста и таблиц ПЗ прямо в предпросмотре ----
    def _save_pz_edits(self) -> None:
        """Записывает в DOCX правки абзацев и ячеек таблиц из предпросмотра."""
        pending = self.preview.pz_pending_edits()
        pending_cells = self.preview.pz_pending_cell_edits()
        if not pending and not pending_cells:
            return
        try:
            from ..core.docx_edit import apply_docx_edits
            total = 0
            details = []
            # Один проход по каждому файлу: абзацы и таблицы пишутся за одно
            # открытие/сохранение DOCX.
            for path in dict.fromkeys(list(pending) + list(pending_cells)):
                result = apply_docx_edits(path, pending.get(path), pending_cells.get(path))
                total += result["paragraphs"] + result["table_cells"]
                details.append(
                    f"{os.path.basename(path)}: абзацев — {result['paragraphs']}, "
                    f"ячеек таблиц — {result['table_cells']}")
            if total == 0:
                QMessageBox.information(
                    self, "Нет изменений",
                    "Текст пояснительной записки не изменился.")
                self.preview.clear_pz_edits()
                return
            # Правки записаны — учитываем их как применённые.
            for path in dict.fromkeys(list(pending) + list(pending_cells)):
                self.preview.clear_pz_edits(path)
            self._refresh_preview()
            # Обновляем число листов документов ПЗ (могло измениться).
            from ..models import DocKind
            for d in self.project.documents:
                if d.doc_kind == DocKind.POYASNITELNAYA:
                    self._refresh_doc_pages(d)
            self.sheet_tree.set_project(self.project)
            QMessageBox.information(
                self, "Сохранено",
                "Правки пояснительной записки и её таблиц записаны "
                "в исходный DOCX:\n"
                + "\n".join(details))
        except Exception as e:
            QMessageBox.critical(
                self, "Ошибка сохранения",
                f"Не удалось записать правки пояснительной записки:\n{e}")

    # ---- Редактор пояснительной записки ----
    def _pz_document(self):
        """Возвращает DOCX пояснительной записки для текущего выбора в дереве.

        Если выбран документ раздела «Пояснительная записка», берётся он;
        иначе — первый доступный DOCX этого раздела.
        """
        from ..models import DocKind
        sheet = self._current_sheet
        obj = sheet.get("obj")
        if obj is not None and getattr(obj, "doc_kind", None) == DocKind.POYASNITELNAYA:
            path = getattr(obj, "path", "") or ""
            if path.lower().endswith((".docx", ".doc")) and os.path.exists(path):
                return obj
        # Иначе — любой DOCX раздела ПЗ (в т.ч. исключённый из сборки).
        for d in self.project.documents:
            if getattr(d, "doc_kind", None) != DocKind.POYASNITELNAYA:
                continue
            path = getattr(d, "path", "") or ""
            if path.lower().endswith((".docx", ".doc")) and os.path.exists(path):
                return d
        return None

    def _edit_current_pz(self) -> None:
        """Открывает редактор всех листов пояснительной записки."""
        self._collect_forms()
        doc = self._pz_document()
        if doc is None:
            QMessageBox.information(
                self, "Нет пояснительной записки",
                "В разделе «Пояснительная записка» нет исходного документа Word (DOCX).\n"
                "Добавьте документ в разделе или выберите DOCX в дереве «Листы тома».")
            return
        from .pz_editor import PzEditorDialog
        dlg = PzEditorDialog(doc.path, self)
        if dlg.exec() == QDialog.Accepted:
            # Текст изменился — пересчитываем листы и обновляем предпросмотр и дерево.
            self._refresh_doc_pages(doc)
            self._refresh_preview()
            self.sheet_tree.set_project(self.project)

    @staticmethod
    def _refresh_doc_pages(doc) -> None:
        """Пересчитывает число листов документа той же пагинацией, что и сборка.

        В дереве и в свойствах показывается именно то число листов, которое
        реально отображается в предпросмотре (а не грубая оценка по объёму).
        """
        path = getattr(doc, "path", "") or ""
        if not path or not os.path.exists(path):
            return
        try:
            from ..core.preview_html import _doc_sheet_count
            doc.pages = max(1, _doc_sheet_count(doc))
        except Exception:
            pass

    def _section_info_for_current(self) -> dict | None:
        """Возвращает информацию о разделе, к которому относится текущий лист."""
        sheet = self._current_sheet
        # Раздел уже известен для документов/приложений/разделов.
        sec = sheet.get("section")
        if not sec:
            # Для стандартных листов (cover/title/content/changes) — по ключу.
            kind = sheet.get("kind")
            mapping = {
                "cover": "cover", "title": "title", "content": "content",
                "changes": "changes",
            }
            sec = mapping.get(kind, kind)
        if not sec:
            return None
        from ..models import SECTION_LABELS
        return {
            "kind": "section", "section": sec,
            "section_label": SECTION_LABELS.get(sec, sec),
            "label": SECTION_LABELS.get(sec, sec),
        }
