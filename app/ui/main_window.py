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
    QFileDialog, QMessageBox, QSplitter, QDialog, QProgressBar, QLabel,
)
from PySide6.QtCore import Qt, QEvent, QTimer, QThread, Slot
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
        #: DOCX, у которых текст правили, но число листов ещё не пересчитано
        #: (пересчёт занимает около секунды — откладываем до первой нужды).
        self._pages_dirty: set = set()
        self.setWindowTitle("Сборка тома по ГОСТ Р 21.101-2020")

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # ---- Левая панель: дерево листов (структура тома) ----
        self.sheet_tree = SheetTree()
        self.sheet_tree.selection_changed.connect(self._on_sheet_selected)
        self.sheet_tree.item_modified.connect(self._on_structure_changed)

        # ---- Центр: HTML-предпросмотр ----
        self.preview = SheetPreview()
        self.preview.field_changed.connect(self._on_field_changed)
        self.preview.signature_clicked.connect(self._on_signature_clicked)
        self.preview.save_pz_requested.connect(self._save_pz_edits)
        # Автосохранение: правка ПЗ сама уходит в связанный DOCX.
        self.preview.autosave_requested.connect(self._on_pz_autosave)
        self.preview.repaginate_requested.connect(self._start_word_repagination)
        self.preview.paragraph_changed.connect(self._mark_pagination_stale)
        self.preview.cell_changed.connect(self._mark_pagination_stale)
        self._pagination_stale: set = set()
        self._pagination_invalidated_at: dict = {}
        self._pagination_layouts: dict = {}
        self._pagination_threads: dict = {}
        self._pagination_busy: set = set()
        self._pagination_job_id = 0
        self._pagination_paths_by_job: dict = {}
        self._pagination_thread_by_job: dict = {}
        self._pagination_errors: dict = {}
        self._pagination_status_label = QLabel("ПЗ: не рассчитана")
        self._pagination_progress = QProgressBar()
        self._pagination_progress.setRange(0, 0)
        self._pagination_progress.setFixedWidth(110)
        self._pagination_progress.setFixedHeight(12)
        self._pagination_progress.setTextVisible(False)
        self._pagination_progress.hide()
        self.statusBar().addPermanentWidget(self._pagination_status_label)
        self.statusBar().addPermanentWidget(self._pagination_progress)

        # ---- Правая панель: контекстные настройки (вместо вкладок) ----
        self.context = ContextPanel(project=self.project)
        self.context.item_modified.connect(self._on_item_modified)
        self.context.edit_pz_requested.connect(self._open_pz_editor)

        # ---- Нижняя панель: основная надпись активного листа (масштаб 100%) ----
        from .stamp_panel import StampPanel
        self.stamp_panel = StampPanel()
        self.stamp_panel.restore_state()
        self.preview.stamps_changed.connect(self.stamp_panel.set_stamps)
        self.preview.active_stamp_changed.connect(self.stamp_panel.set_active)
        # Поля панели правятся теми же обработчиками, что и на самом листе:
        # правка уходит в модель проекта, подпись выбирается по клику.
        # Сохранение в панели — по потере фокуса, поэтому после записи листы
        # можно перерисовать: обновляются все связанные штампы.
        self.stamp_panel.field_changed.connect(self._on_panel_field_changed)
        self.stamp_panel.signature_clicked.connect(self._on_signature_clicked)

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
        root.addWidget(self.stamp_panel)

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

    def changeEvent(self, event) -> None:
        """Возврат в окно: отдаём фокус обратно в предпросмотр.

        После переключения в другую программу и обратно клавиатура и мышь
        могли «залипнуть»: активным элементом оставалось поле внутри
        QWebEngineView, которое потеряло фокус на уровне Windows. Возврат
        активирует окно и заново ставит фокус на тот виджет, который был
        активен, — иначе кнопки и поля в окне не реагируют.
        """
        super().changeEvent(event)
        if event.type() == QEvent.ActivationChange and self.isActiveWindow():
            self._restore_input_focus()

    def _restore_input_focus(self) -> None:
        """Восстанавливает фокус ввода после возврата в окно (см. changeEvent)."""
        if not self.isActiveWindow():
            return

        def apply() -> None:
            if not self.isActiveWindow():
                return
            widget = self.focusWidget()
            # Фокус мог остаться на скрытом/неактивном виджете — тогда берём
            # предпросмотр: он занимает центр окна.
            if widget is None or not widget.isVisible() or not widget.isEnabled():
                widget = self.preview
            try:
                widget.setFocus(Qt.OtherFocusReason)
            except Exception:
                pass
            # Странице WebEngine нужен явный «щелчок фокуса»: без него
            # обработчики ввода в HTML не оживают до реального клика.
            view = getattr(self.preview, "view", None)
            if view is not None:
                view.setFocus(Qt.OtherFocusReason)
                try:
                    view.page().runJavaScript(
                        "window.focus(); if (document.body) {"
                        "document.body.dispatchEvent("
                        "new Event('focus', {bubbles: true})); }")
                except Exception:
                    pass

        # Небольшая задержка: Windows отдаёт фокус окну уже после события.
        QTimer.singleShot(60, apply)
        QTimer.singleShot(220, apply)

    def _confirm_pending_pz_edits(self, action: str) -> bool:
        """Спрашивает о несохранённых правках ПЗ перед действием.

        Возвращает True, если можно продолжать (правок нет или пользователь
        решил продолжить). Если пользователь выбрал «Сохранить» — правки
        записываются, и возвращается True при успехе.
        """
        # Сначала дописываем то, что ещё ждёт паузы автосохранения: обычно
        # после этого несохранённых правок не остаётся и вопрос не нужен.
        self.preview.flush_autosave()
        # Число листов понадобится сборке — пересчитываем отложенное.
        self._refresh_dirty_pages()
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

        # «Сборка» — не выпадающее меню, а пункт строки меню: нажатие сразу
        # открывает окно сборки (форматы + «Отмена» / «Собрать»). Промежуточного
        # подменю с единственным пунктом «Собрать том…» больше нет.
        act_build = QAction("&Сборка", self)
        act_build.setToolTip("Открыть окно сборки тома")
        act_build.setShortcut(QKeySequence("Ctrl+B"))
        act_build.triggered.connect(self._build)
        menu_bar.addAction(act_build)

        # Настройки оформления текста ПЗ: шрифт и кегль по группам.
        settings_menu = menu_bar.addMenu("&Настройки")
        act_fonts = QAction("&Шрифты текста ПЗ…", self)
        act_fonts.triggered.connect(self._open_text_fonts)
        settings_menu.addAction(act_fonts)

    # ---- Форматы сборки (окно «Собрать том») ----
    #: Форматы тома в порядке флажков окна: (ключ, подпись).
    BUILD_FORMATS = (
        ("pdf", "PDF (*.pdf)"),
        ("docx", "DOC — Word (*.docx)"),
        ("html", "HTML (*.html)"),
    )

    def _saved_build_formats(self) -> list:
        """Форматы из config.json (пустой/битый набор — все форматы)."""
        from ..config import load_config
        try:
            saved = load_config().build_formats
        except Exception:
            saved = None
        if not saved:
            saved = [key for key, _label in self.BUILD_FORMATS]
        keys = {key for key, _label in self.BUILD_FORMATS}
        return [key for key in saved if key in keys] or [
            key for key, _label in self.BUILD_FORMATS]

    def _save_build_formats(self, chosen: list, chapters: bool | None = None) -> None:
        """Запоминает выбранный набор в config.json (пустой не пишется)."""
        if not chosen:
            return
        from ..config import load_config, save_config
        try:
            cfg = load_config()
        except Exception:
            return
        cfg.build_formats = list(chosen)
        if chapters is not None:
            cfg.build_docx_chapters = bool(chapters)
        try:
            save_config(cfg)
        except Exception:
            # Не удалось записать файл — набор остаётся действующим до выхода.
            pass

    def _ask_build_formats(self) -> tuple | None:
        """Показывает окно выбора форматов.

        Возвращает (форматы, «по разделам»); None — пользователь отменил.
        """
        from PySide6.QtWidgets import QDialog
        from .build_dialog import BuildDialog
        from ..config import load_config
        try:
            chapters = bool(load_config().build_docx_chapters)
        except Exception:
            chapters = False
        dlg = BuildDialog(self, self._saved_build_formats(), chapters=chapters)
        # Сравнение идёт с QDialog.Accepted: у ЭКЗЕМПЛЯРА диалога атрибута
        # Accepted нет, и `dlg.Accepted` падал с AttributeError — слот
        # обрывался, поэтому нажатие «Собрать» ничего не делало.
        if dlg.exec() != QDialog.Accepted:
            return None
        chosen = dlg.selected_formats()
        chapters = dlg.chapters()
        self._save_build_formats(chosen, chapters)
        return chosen, chapters

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
        self._pagination_stale.clear()
        self._pagination_invalidated_at.clear()
        self._pagination_layouts.clear()
        self._pagination_errors.clear()
        self._set_pagination_status("не рассчитана")
        from ..core.word_pagination import invalidate_docx_pages
        invalidate_docx_pages()
        self.context.load_project(self.project)
        self.sheet_tree.set_project(self.project)
        self._refresh_preview()
        QTimer.singleShot(0, lambda: self._start_word_repagination(force=False))

    def _collect_forms(self) -> None:
        self.context.collect(self.project)

    # ---- Предпросмотр ----
    def _refresh_preview(self) -> None:
        snapshot = build_preview_snapshot(self.project)
        current_path = self._current_sheet.get("path", "")
        pz_section = self._current_sheet.get("section") == "Пояснительная записка"
        paths = self._current_pz_paths() if pz_section else []
        self._load_cached_pagination_layouts(paths)
        snapshot["word_pagination"] = {
            p: self._pagination_layouts[p] for p in paths
            if p in self._pagination_layouts and p not in self._pagination_stale
        }
        snapshot["pagination_stale"] = any(p in self._pagination_stale for p in paths)
        snapshot["pagination_busy"] = any(p in self._pagination_busy for p in paths)
        snapshot["pagination_error"] = self._pagination_errors.get(current_path, "")
        # Клеймим снимок метаданными текущего листа (для заглушек документов/приложений).
        if self._current_sheet.get("kind") in ("doc", "appendix", "section"):
            snapshot["label"] = self._current_sheet.get("label", "")
            snapshot["section_label"] = self._current_sheet.get("section_label", "")
            snapshot["path"] = self._current_sheet.get("path", "")
        # Раздел нужен ВСЕМ листам: по нему адресуются правки ячеек штампа
        # (таблица изменений и даты подписей). Без него ключ правки получал
        # пустой раздел, и значение не попадало в модель проекта.
        snapshot["section"] = self._current_sheet.get("section", "")
        # Показываются ВСЕ листы выбранного пункта (документ/раздел — постранично).
        self.preview.show_sheet(self._current_sheet.get("kind", "cover"),
                                snapshot, project=self.project)

    def _load_cached_pagination_layouts(self, paths: list[str]) -> None:
        from ..core.word_pagination import cached_docx_pages
        for path in paths:
            if path in self._pagination_layouts or path in self._pagination_stale:
                continue
            layout = cached_docx_pages(path, self._first_form_for_path(path))
            if layout is not None:
                self._pagination_layouts[path] = layout

    def _current_pz_paths(self) -> list[str]:
        from ..models import DocKind
        return [d.path for d in self.project.documents_for(DocKind.POYASNITELNAYA)
                if d.path and os.path.isfile(d.path)]

    def _first_form_for_path(self, path: str):
        from ..core.preview_html import pz_doc_first_form
        return pz_doc_first_form(self.project, path)

    def _cached_word_layout(self, path: str, first_form):
        from ..core.word_pagination import cached_docx_pages
        return cached_docx_pages(path, first_form)

    def _mark_pagination_stale(self, *_args) -> None:
        paths = set(self._current_pz_paths())
        path = next((value for value in _args if isinstance(value, str) and value in paths), "")
        if path:
            self._pagination_stale.add(path)
            self._pagination_invalidated_at[path] = os.stat(path).st_mtime_ns if os.path.exists(path) else 0
            self._pagination_layouts.pop(path, None)
            from ..core.word_pagination import invalidate_docx_pages
            invalidate_docx_pages(path)
        else:
            paths = self._current_pz_paths()
            self._pagination_stale.update(paths)
            for item in paths:
                self._pagination_invalidated_at[item] = os.stat(item).st_mtime_ns
                self._pagination_layouts.pop(item, None)
            from ..core.word_pagination import invalidate_docx_pages
            invalidate_docx_pages()
        paths = self._current_pz_paths()
        self.preview.set_pagination_state(
            stale=any(p in self._pagination_stale for p in paths),
            busy=any(p in self._pagination_busy for p in paths))
        self._pagination_status_for_current()

    def _start_word_repagination(self, force: bool = True) -> None:
        paths = self._current_pz_paths()
        if self.preview.has_pz_edits():
            self.preview.flush_autosave()
        pending = [(path, self._first_form_for_path(path)) for path in paths
                   if path not in self._pagination_busy and
                   (force or path in self._pagination_stale or
                    self._cached_word_layout(path, self._first_form_for_path(path)) is None)]
        if not pending:
            self._pagination_status_for_current()
            return
        from .pagination_worker import PaginationWorker
        self._pagination_job_id += 1
        job_id = self._pagination_job_id
        thread = QThread(self)
        worker = PaginationWorker(job_id, pending)
        worker.moveToThread(thread)
        worker.progress.connect(self._on_pagination_progress)
        thread.started.connect(worker.run)
        worker.finished.connect(thread.quit)
        worker.finished.connect(self._word_repagination_finished)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        key = id(thread)
        self._pagination_threads[key] = (thread, worker)
        self._pagination_paths_by_job[job_id] = [path for path, _ in pending]
        self._pagination_thread_by_job[job_id] = thread
        self._pagination_busy.update(path for path, _ in pending)
        self.preview.set_pagination_state(
            stale=any(path in self._pagination_stale for path, _ in pending),
            busy=True)
        self._set_pagination_status("Пересчитывается — подготовка DOCX…", busy=True)
        self.preview.set_pagination_state(
            stale=bool(self._pagination_stale), busy=True,
            message="Подготовка нормализованной копии DOCX…")
        thread.start()

    @Slot(str)
    def _on_pagination_progress(self, message: str) -> None:
        self._set_pagination_status(message, busy=True)
        self.preview.set_pagination_state(
            stale=bool(self._pagination_stale), busy=True, message=message)

    def _set_pagination_status(self, text: str, *, busy: bool = False,
                               stale: bool = False, error: bool = False) -> None:
        self._pagination_status_label.setText("ПЗ: " + text)
        self._pagination_progress.setVisible(busy)
        if error:
            color = "#b42318"
        elif stale:
            color = "#b54708"
        elif busy:
            color = "#175cd3"
        else:
            color = "#067647"
        self._pagination_status_label.setStyleSheet(f"color:{color}; font-weight:600;")

    def _pagination_status_for_current(self) -> None:
        if self._current_sheet.get("section") != "Пояснительная записка":
            return
        paths = self._current_pz_paths()
        if not paths:
            self._set_pagination_status("нет документов")
            return
        busy_paths = [p for p in paths if p in self._pagination_busy]
        if busy_paths:
            self._set_pagination_status(
                f"пересчитывается ({len(busy_paths)}/{len(paths)})…", busy=True)
            return
        errors = [self._pagination_errors[p] for p in paths if p in self._pagination_errors]
        if errors:
            self._set_pagination_status("ошибка: " + errors[0], error=True)
            return
        missing = [p for p in paths if p in self._pagination_stale
                   or p not in self._pagination_layouts
                   or self._pagination_layouts[p].get("page_count") is None]
        if missing:
            count = len(missing)
            self._set_pagination_status(
                f"требует пересчёта ({count} документ.)", stale=True)
            return
        total = sum(int(self._pagination_layouts[p].get("page_count", 0) or 0)
                    for p in paths)
        self._set_pagination_status(f"актуальна — {total} листов")

    @Slot(object)
    def _word_repagination_finished(self, result: dict) -> None:
        job_id = result.get("job_id")
        thread = self._pagination_thread_by_job.pop(job_id, None)
        if thread is None:
            return
        if job_id != self._pagination_job_id:
            for path in self._pagination_paths_by_job.pop(job_id, []):
                self._pagination_busy.discard(path)
            self._pagination_threads.pop(id(thread), None)
            self.preview.set_pagination_state(
                stale=bool(self._pagination_stale), busy=bool(self._pagination_busy))
            self._refresh_preview()
            return
        for path in self._pagination_paths_by_job.pop(job_id, []):
            self._pagination_busy.discard(path)
        from ..models import DocKind
        errors = result.get("errors", {})
        for path, layout in result.get("results", {}).items():
            invalidated_at = self._pagination_invalidated_at.get(path)
            current_mtime = os.stat(path).st_mtime_ns if os.path.exists(path) else 0
            accepted = invalidated_at is None or current_mtime == invalidated_at
            if accepted:
                self._pagination_stale.discard(path)
                self._pagination_invalidated_at.pop(path, None)
                self._pagination_layouts[path] = layout
            self._pagination_errors.pop(path, None)
            if accepted:
                from ..core.word_pagination import store_docx_pages
                store_docx_pages(path, self._first_form_for_path(path), layout)
            if accepted:
                for doc in self.project.documents:
                    if doc.doc_kind == DocKind.POYASNITELNAYA and os.path.abspath(doc.path) == os.path.abspath(path):
                        doc.pages = max(1, int(layout.get("page_count", 1)))
        for path, error in errors.items():
            if path:
                self._pagination_errors[path] = error
                self._set_pagination_status(
                    f"ошибка {os.path.basename(path)}: {error}", error=True)
                self.statusBar().showMessage(
                    f"Не удалось пересчитать страницы {os.path.basename(path)}: {error}", 10000)
        self._pagination_threads.pop(id(thread), None)
        if errors and not any(path in self._pagination_errors for path in self._current_pz_paths()):
            detail = next(iter(errors.values()))
            self._set_pagination_status(f"ошибка: {detail}", error=True)
        else:
            self._pagination_status_for_current()
        paths = self._current_pz_paths() if self._current_sheet.get("section") == "Пояснительная записка" else []
        self.preview.set_pagination_state(
            stale=any(path in self._pagination_stale for path in paths),
            busy=any(path in self._pagination_busy for path in paths),
            error=self._pagination_errors.get(self._current_sheet.get("path", ""), ""))
        self.sheet_tree.set_project(self.project)
        self._refresh_preview()
        self._pagination_status_for_current()
        if any(path in self._pagination_stale for path in self._current_pz_paths()) and not self._pagination_busy:
            QTimer.singleShot(0, lambda: self._start_word_repagination(force=False))
        if not errors:
            self.statusBar().showMessage("Пересчёт страниц ПЗ завершён", 4000)

    def _on_sheet_selected(self, info: dict) -> None:
        self._current_sheet = info
        # Правки текущего листа дописываем в файл до перерисовки.
        self.preview.flush_autosave()
        self._refresh_preview()
        if info.get("section") == "Пояснительная записка":
            paths = self._current_pz_paths()
            missing = [path for path in paths
                       if path not in self._pagination_layouts
                       or path in self._pagination_stale]
            if missing:
                self.statusBar().showMessage(
                    f"ПЗ: в фоне проверяю раскладку ({len(missing)} документ.)")
                self._set_pagination_status(
                    f"требует пересчёта ({len(missing)} документ.)", stale=True)
                self._start_word_repagination(force=False)
            else:
                self.statusBar().showMessage("ПЗ: раскладка страниц актуальна", 4000)
                self._pagination_status_for_current()
        # Показываем настройки, относящиеся к выбранному пункту.
        self.context.show_context(info)

    def _on_item_modified(self) -> None:
        """Состав тома изменён — обновляем предпросмотр и дерево.

        Документ или приложение могли быть не только исключены, но и убраны из
        проекта: тогда выбор в дереве указывает на несуществующий элемент, и
        предпросмотр показывал бы заглушку удалённого файла. В этом случае
        переключаемся на раздел, к которому элемент относился.
        """
        if not self._current_object_alive():
            info = self._section_info_for_current()
            if info:
                self._current_sheet = info
                self.context.show_context(info)
        self._refresh_preview()
        self.sheet_tree.set_project(self.project)

    def _current_object_alive(self) -> bool:
        """Выбранный в дереве объект всё ещё есть в проекте?"""
        obj = self._current_sheet.get("obj")
        if obj is None:
            return True
        return obj in self.project.documents or obj in self.project.appendices

    def _on_structure_changed(self) -> None:
        """Состав изменён из дерева — обновляем только листы.

        Дерево уже отражает новое состояние (флажок нажат самим пользователем,
        удаление перестроило дерево), поэтому перестраивать его не нужно —
        иначе сбрасывался бы выбор.
        """
        if not self._current_object_alive():
            info = self._section_info_for_current()
            if info:
                self._current_sheet = info
                self.context.show_context(info)
        self._refresh_preview()

    # ---- Редактирование полей ----
    def _on_panel_field_changed(self, key: str, value: str) -> None:
        """Правка поля в нижней панели штампа сохранена (по потере фокуса).

        Значение пишется в модель теми же обработчиками, что и правка на самом
        листе. После записи листы перерисовываются: правка уже завершена (фокус
        ушёл), поэтому курсор не пострадает, а все связанные штампы — на
        остальных листах пункта, в содержании и в самой панели — показывают
        новое значение.
        """
        self._on_field_changed(key, value)
        self._refresh_preview()

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
                # Поле на обложке/титуле показывает полное наименование стадии;
                # принимаем и его, и обозначение «П» / «Р».
                from ..models import parse_stage
                stage = parse_stage(value)
                if stage is not None:
                    req.stage = stage
                # Нераспознанный текст в модель не пишем; лист перерисуется
                # из свежего снимка.
                self.context.requisites.sync_requisites(req)
                return
            elif attr == "year":
                try:
                    req.year = int(value)
                except Exception:
                    pass
            elif attr == "volume_id":
                # В графе «Обозначение» показано том + код раздела
                # («02-2026-ТР.ПЗ»); в модель пишем только обозначение тома.
                req.volume_id = self._strip_section_code(value)
            else:
                setattr(req, attr, value)
            # Синхронизируем форму реквизитов (чтобы не было расхождений).
            # Настройки графы 9 (текст/логотип) не трогаем — иначе правка
            # наименования организации сбрасывала бы выбранный логотип.
            self.context.requisites.sync_requisites(req)
            return

        # Таблица изменений штампа и даты подписей: ключи `ch|<раздел>|<строка>|
        # <колонка>` и `sd|<раздел>|<роль>`. Правка сразу уходит в модель, поэтому
        # значение обновляется во всех листах раздела.
        if self._apply_stamp_edit(key, value):
            return

    def _apply_stamp_edit(self, key: str, value: str) -> bool:
        """Записывает правку ячейки штампа в модель. True — ключ распознан.

        Ключи: `ch|<раздел>|<строка>|<колонка>` (таблица изменений, графы 14–19),
        `sd|<раздел>|<роль>` (дата подписи — старый формат) и
        `sp|/ss|<раздел>|<строка>|<поле>` (строки подписей левого блока:
        должность, фамилия, дата).
        """
        from ..models import (set_stamp_row_value, set_stamp_sign_date,
                              set_stamp_signature, STAMP_CHANGE_COLS,
                              STAMP_SIG_FIELDS)
        parts = (key or "").split("|")
        prefix = parts[0] if parts else ""
        if prefix == "ch" and len(parts) == 4:
            _, kind, row_text, col = parts
            try:
                row = int(row_text)
            except ValueError:
                return False
            if not kind or col not in STAMP_CHANGE_COLS:
                return False
            set_stamp_row_value(self.project, kind, row, col, value)
            return True
        if prefix == "sd" and len(parts) == 3:
            _, kind, role = parts
            if not kind or not role:
                return False
            set_stamp_sign_date(self.project, kind, role, value)
            return True
        # Строки подписей: должность, фамилия и дата. Значение пишется в
        # проект, поэтому оно сохраняется вместе с томом и попадает в
        # Word/PDF при сборке. Форма нужна, чтобы перевести номер строки в
        # общий слот данных (у графики строки 5 и 6 — это 4 и 5 записки).
        if prefix in ("sp", "ss") and len(parts) == 4:
            _, kind, row_text, field = parts
            if not kind or field not in STAMP_SIG_FIELDS:
                return False
            try:
                row = int(row_text)
            except ValueError:
                return False
            if row < 0:
                return False
            form = self._stamp_row_form(kind)
            set_stamp_signature(self.project, kind, row, field, value, form=form)
            return True
        return False

    @staticmethod
    def _stamp_row_form(kind: str) -> int:
        """Форма штампа раздела: 3 — графическая часть, иначе 5 (текст).

        Нужна для связи строк подписей: у формы 3 шесть строк, и её строки
        5 и 6 несут те же данные, что строки 4 и 5 формы 5.
        """
        from ..models import SECTION_GRAPHIC
        return 3 if kind == SECTION_GRAPHIC else 5

    def _strip_section_code(self, value: str) -> str:
        """Убирает доп. обозначение раздела из значения графы «Обозначение».

        В графе показано «обозначение тома + код раздела»; правая инлайн-правка
        этой ячейки не должна записывать код в обозначение тома.
        """
        text = (value or "").strip()
        codes = [c for c in (self.project.section_codes or {}).values() if c]
        for code in sorted(codes, key=len, reverse=True):
            if text.endswith(code):
                return text[: -len(code)].strip()
        return text

    def _on_signature_clicked(self, key: str) -> None:
        """Клик по полю подписи — выбор подписанта.

        Для графы «Подп.» строки штампа (`ss|<раздел>|<строка>|signature`)
        выбор записывается в эту строку: фамилия подставляется из реестра, а
        изображение подписи подтягивается по фамилии.
        """
        from ..core.preview_model import role_label, build_preview_snapshot
        from ..models import set_stamp_signature
        # Сначала синхронизируем формы с моделью (чтобы реестр подписантов был актуален)
        self.context.signers_form.set_signers(self.project.signers)
        self._collect_forms()
        snapshot = build_preview_snapshot(self.project)
        row_ref = self._stamp_row_ref(key)
        if row_ref is not None:
            kind, row = row_ref
            label_text = self._stamp_row_title(kind, row)
            current_surname = self._stamp_row_surname(kind, row)
            dlg = SignaturePickerDialog(self, role_label=label_text,
                                        signers=snapshot["signers"],
                                        current_surname=current_surname)
            if dlg.exec() == SignaturePickerDialog.Accepted:
                set_stamp_signature(self.project, kind, row, "surname",
                                    dlg.result_surname or "",
                                    form=self._stamp_row_form(kind))
                self._refresh_preview()
            return
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

    @staticmethod
    def _stamp_row_ref(key: str):
        """(раздел, номер строки) для ключа строки подписи штампа, иначе None."""
        parts = (key or "").split("|")
        if len(parts) != 4 or parts[0] != "ss" or parts[3] != "signature":
            return None
        try:
            row = int(parts[2])
        except ValueError:
            return None
        if not parts[1] or row < 0:
            return None
        return parts[1], row

    def _stamp_row(self, kind: str, row: int) -> dict:
        """Строка подписи раздела по номеру (значения по умолчанию + правки)."""
        from ..models import stamp_signature_rows, SECTION_GRAPHIC
        form = 3 if kind == SECTION_GRAPHIC else 5
        rows = stamp_signature_rows(self.project, kind, form)
        if 0 <= row < len(rows):
            return rows[row]
        return {}

    def _stamp_row_title(self, kind: str, row: int) -> str:
        """Подпись строки для окна выбора подписанта (должность + фамилия)."""
        data = self._stamp_row(kind, row)
        post = (data.get("post") or "").strip()
        surname = (data.get("surname") or "").strip()
        if post and surname:
            return f"{post} — {surname}"
        return post or surname or "Подпись"

    def _stamp_row_surname(self, kind: str, row: int) -> str:
        return (self._stamp_row(kind, row).get("surname") or "").strip()

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
        """Собрать том: форматы выбираются в мини-окне «Собрать том»."""
        from PySide6.QtWidgets import QFileDialog, QMessageBox

        if not self._confirm_pending_pz_edits("собирать том"):
            return

        # Форматы спрашиваем окном: «Отмена» в нём отменяет всю сборку,
        # поэтому проверки и preflight идут уже после выбора форматов.
        answer = self._ask_build_formats()
        if not answer:
            return
        formats, chapters = answer

        self._collect_forms()

        # Preflight: ошибки блокируют сборку, warnings — с подтверждением.
        try:
            from ..core.preflight import run_preflight
            pf = run_preflight(self.project)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось выполнить preflight:\n{e}")
            return
        if pf.errors:
            QMessageBox.critical(
                self, "Сборка отменена",
                "Обнаружены ошибки — сборка тома не запущена:\n\n• "
                + "\n• ".join(pf.errors))
            return
        if pf.warnings:
            ans = QMessageBox.warning(
                self, "Предупреждения",
                "Предупреждения:\n\n• " + "\n• ".join(pf.warnings)
                + "\n\nПродолжить сборку?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ans != QMessageBox.Yes:
                return

        # Имя по умолчанию: обозначение тома и дата сборки.
        from datetime import date
        from ..core.docx_exporter import _safe_filename
        volume_id = (self.project.requisites.volume_id or "том").strip()
        today = date.today()
        default_name = _safe_filename(
            f"{volume_id}_{today.year % 100:02d}_{today.month:02d}_{today.day:02d}")
        filters = {"pdf": "PDF (*.pdf)", "docx": "Документ Word (*.docx)",
                   "html": "HTML (*.html)"}
        ext = {"pdf": ".pdf", "docx": ".docx", "html": ".html"}[formats[0]]
        if chapters:
            out_dir = QFileDialog.getExistingDirectory(
                self, "Каталог для файлов по разделам")
            if not out_dir:
                return
            base = os.path.join(out_dir, default_name + ext)
        else:
            filter_str = ";;".join(filters[f] for f in formats)
            base = QFileDialog.getSaveFileName(
                self, "Сохранить том", default_name + ext, filter_str)[0]
        if not base:
            return

        # Лог сборки — в tmp/build.log (причины ошибок не теряются).
        log_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            "tmp", "build.log")
        import logging
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        logging.basicConfig(
            filename=log_path, filemode="a", level=logging.INFO,
            format="%(asctime)s %(levelname)s %(name)s: %(message)s")
        log = logging.getLogger("tom_builder.build")
        log.info("=== Сборка тома: %s (форматы: %s) ===", base, ", ".join(formats))

        stem = os.path.splitext(base)[0]
        from PySide6.QtCore import QEventLoop
        from .build_worker import BuildWorker
        from .build_progress import BuildProgressDialog

        progress = BuildProgressDialog(self)
        progress.show()

        thread = QThread(self)
        worker = BuildWorker(self.project, formats, chapters, base)
        worker.moveToThread(thread)
        loop = QEventLoop(self)
        result = {}
        worker.progress.connect(progress.set_stage)
        worker.finished.connect(lambda data: (result.update(data), loop.quit()))
        thread.started.connect(worker.run)
        worker.finished.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.start()
        loop.exec()
        thread.quit()
        thread.wait()

        done = result.get("done", [])
        failed = result.get("failed", [])
        html_path = result.get("html_path")

        if "pdf" in formats:
            out = stem + ".pdf"
            try:
                from ..core.html_export import html_file_to_pdf
                if html_path and os.path.exists(html_path):
                    progress.set_stage("Подготовка PDF…")
                    progress.show()

                    def update_pdf_progress(current, total):
                        progress.set_sheet_progress(current, total)
                        from PySide6.QtWidgets import QApplication
                        QApplication.processEvents()

                    html_file_to_pdf(html_path, out,
                                     progress_callback=update_pdf_progress)
                    progress.finish()
                    log.info("PDF (из HTML): %s", out)
                    done.append(f"PDF: {out}")
                else:
                    raise RuntimeError(
                        "HTML-том не собран — печатать PDF не из чего")
            except Exception as e:
                log.exception("html_file_to_pdf failed")
                failed.append(f"PDF: {e}")
        progress.close()

        # Итог: что собрано и что не удалось (частичный успех не скрывается).
        msg = ""
        if done:
            msg = "Том собран:\n\n" + "\n".join(done)
        if failed:
            msg += ("\n\n" if msg else "") + "Не удалось собрать:\n\n" \
                + "\n".join(failed) + f"\n\nПодробности: {log_path}"
            QMessageBox.warning(self, "Сборка завершена с ошибками", msg)
        else:
            QMessageBox.information(self, "Готово", msg)

    def _log_build_error(self, message: str) -> str:
        """Пишет traceback в tmp/build.log и возвращает путь к логу."""
        import logging
        log_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            "tmp", "build.log")
        try:
            os.makedirs(os.path.dirname(log_path), exist_ok=True)
            logging.basicConfig(filename=log_path, filemode="a", level=logging.INFO,
                                format="%(asctime)s %(levelname)s %(name)s: %(message)s")
            logging.getLogger("tom_builder.build").exception(message)
        except Exception:
            pass
        return log_path

    # ---- Правка текста и таблиц ПЗ прямо в предпросмотре ----
    def _save_pz_edits(self, silent: bool = False) -> None:
        """Записывает в DOCX правки абзацев и ячеек таблиц из предпросмотра.

        :param silent: автосохранение — без диалогов и без перерисовки листов
            (перерисовка сбросила бы курсор посреди ввода). Дерево и
            предпросмотр обновляются при следующем переключении.
        """
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
                if not silent:
                    QMessageBox.information(
                        self, "Нет изменений",
                        "Текст пояснительной записки не изменился.")
                self.preview.clear_pz_edits()
                return
            # Правки записаны — учитываем их как применённые.
            for path in dict.fromkeys(list(pending) + list(pending_cells)):
                self.preview.clear_pz_edits(path)
            changed_paths = set(pending) | set(pending_cells)
            self._pagination_stale.update(changed_paths)
            for path in changed_paths:
                self._pagination_invalidated_at[path] = os.stat(path).st_mtime_ns if os.path.exists(path) else 0
                self._pagination_layouts.pop(path, None)
            from ..core.word_pagination import invalidate_docx_pages
            for path in changed_paths:
                invalidate_docx_pages(path)
            if silent:
                # Лист уже показывает введённый текст — перерисовка не нужна.
                self._pages_dirty.update(changed_paths)
                self._autosave_notice()
                return
            # Число листов и план Word будут обновлены при следующем переходе
            # в раздел ПЗ либо вручную по кнопке.
            self._refresh_preview()
            self.sheet_tree.set_project(self.project)
            self._autosave_notice()
            QMessageBox.information(
                self, "Сохранено",
                "Правки пояснительной записки и её таблиц записаны "
                "в исходный DOCX:\n"
                + "\n".join(details))
        except Exception as e:
            # Молча терять правки нельзя: причина — в лог, пользователю — сообщение.
            log_path = self._log_build_error(f"не удалось записать правки ПЗ: {e}")
            QMessageBox.critical(
                self, "Ошибка сохранения",
                f"Не удалось записать правки пояснительной записки:\n{e}\n\n"
                f"Подробности: {log_path}")

    def _on_pz_autosave(self) -> None:
        """Автосохранение: правки уходят в связанный DOCX без диалогов.

        Формы сюда не собираются: правки ПЗ живут в предпросмотре и в модели
        реквизитов не участвуют — лишняя синхронизация только мешала бы вводу.
        """
        self._save_pz_edits(silent=True)

    def _autosave_notice(self) -> None:
        """Показывает, что правки уже записаны в файл (кратковременная подсказка)."""
        sb = self.statusBar()
        sb.showMessage("Правки пояснительной записки сохранены в файл", 4000)

    def _refresh_dirty_pages(self) -> None:
        """Пересчитывает число листов у документов, которые правили.

        Вызывается, когда число листов действительно нужно (перерисовка листа,
        сборка): при обычном вводе пересчёт не выполняется, чтобы не тормозить.
        """
        if not self._pages_dirty:
            return
        from ..models import DocKind
        for d in self.project.documents:
            if d.doc_kind != DocKind.POYASNITELNAYA:
                continue
            path = getattr(d, "path", "") or ""
            if path in self._pages_dirty:
                self._refresh_doc_pages(d, self.project)
        self._pages_dirty.clear()

    # ---- Настройки оформления ----
    def _open_text_fonts(self) -> None:
        """Настройки шрифтов текста ПЗ: обычный текст, заголовки, таблицы.

        Значения сохраняются в конфигурацию оформления (config.json) и сразу
        применяются к предпросмотру: кегль в предпросмотре и в собранном томе
        должен совпадать, иначе число листов расходится.
        """
        from ..config import load_config, save_config
        from .fonts_dialog import TextFontsDialog
        cfg = load_config()
        dlg = TextFontsDialog(self, cfg)
        if dlg.exec() != QDialog.Accepted:
            return
        dlg.apply_to(cfg)
        try:
            save_config(cfg)
        except Exception as e:
            QMessageBox.warning(
                self, "Настройки не сохранены",
                f"Не удалось записать config.json:\n{e}\n\n"
                "Значения применены до перезапуска программы.")
        # Кегли влияют на Word-компоновку: помечаем планы устаревшими.
        for path in self._current_pz_paths():
            self._pagination_stale.add(path)
            self._pagination_invalidated_at[path] = os.stat(path).st_mtime_ns
            self._pagination_layouts.pop(path, None)
        from ..core.word_pagination import invalidate_docx_pages
        invalidate_docx_pages()
        self._refresh_preview()
        if self._current_sheet.get("section") == "Пояснительная записка":
            self._start_word_repagination()

    def _open_pz_editor(self, doc) -> None:
        """Постраничный редактор записки — единственный вход для всех кнопок."""
        if doc is None:
            return
        path = getattr(doc, "path", "") or ""
        if not self._is_editable_docx(path):
            QMessageBox.information(
                self, "Формат не поддерживается",
                "Редактирование доступно только для исходного документа Word (DOCX).")
            return
        from .pz_editor import PzEditorDialog
        from ..core.preview_html import pz_doc_first_form
        dlg = PzEditorDialog(path, self,
                             first_form=pz_doc_first_form(self.project, path))
        if dlg.exec() == QDialog.Accepted:
            # Текст изменился — план Word нужно обновить.
            self._mark_pagination_stale(path)
            self._refresh_preview()
            self.sheet_tree.set_project(self.project)

    @staticmethod
    def _is_editable_docx(path: str) -> bool:
        return bool(path) and path.lower().endswith((".docx", ".doc")) \
            and os.path.exists(path)

    @staticmethod
    def _refresh_doc_pages(doc, project=None) -> None:
        """Пересчитывает число листов документа той же пагинацией, что и сборка.

        В дереве и в свойствах показывается именно то число листов, которое
        реально отображается в предпросмотре (а не грубая оценка по объёму).
        """
        path = getattr(doc, "path", "") or ""
        if not path or not os.path.exists(path):
            return
        try:
            from ..core.preview_html import _doc_sheet_count, pz_doc_first_form
            first_form = 2 if project is None else pz_doc_first_form(project, path)
            doc.pages = max(1, _doc_sheet_count(doc, first_form=first_form))
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
