# -*- coding: utf-8 -*-
"""
Виджет HTML-предпросмотра листа (QWebEngineView + QWebChannel).

Показывает один лист тома (обложка/титул/содержание/осн. надпись/изменения)
в HTML. Поля редактируются inline; изменения и клики по подписи идут через
QWebChannel в Python.

Масштаб: по умолчанию лист отображается в реальном размере (100%). Есть
панель управления: 100% / вписать в окно / +/- (зум через CSS-масштаб).

Сигналы:
  field_changed(key, value) — отредактировано поле (общие реквизиты/подписи).
  signature_clicked(key)    — клик по полю подписи (открыть выбор подписанта).
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QProgressBar, QStackedLayout,
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtCore import QUrl, QTimer, Signal, Qt

from ..core.preview_html import (render_sheet, render_sheet_fragments,
                                 render_preview_error)
from .preview_bridge import PreviewBridge
from .word_host import WordHost

ZOOM_STEP = 0.1
#: Пауза после остановки ввода, через которую правка ПЗ пишется в DOCX.
AUTOSAVE_DELAY_MS = 1000


def _log_preview_error(kind: str, reason: str) -> str:
    """Пишет сбой предпросмотра в tmp/build.log и возвращает путь к логу."""
    import logging
    log_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        "tmp", "build.log")
    try:
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        logging.basicConfig(filename=log_path, filemode="a", level=logging.INFO,
                            format="%(asctime)s %(levelname)s %(name)s: %(message)s")
        logging.getLogger("tom_builder.preview").exception(
            "предпросмотр: сбой для пункта kind=%s: %s", kind, reason)
    except Exception:
        pass
    return log_path


class SheetPreview(QWidget):
    """Предпросмотр одного листа в HTML."""

    field_changed = Signal(str, str)
    signature_clicked = Signal(str)
    paragraph_changed = Signal(int, str, str)   # (индекс абзаца, путь DOCX, текст)
    cell_changed = Signal(int, int, int, str, str)  # (табл., строка, ячейка, DOCX, текст)
    save_pz_requested = Signal()                # «Сохранить правки ПЗ» нажата
    #: Правки накопились и их пора записать в связанный DOCX (автосохранение).
    autosave_requested = Signal()
    repaginate_requested = Signal()
    #: Перерисован набор листов: основные надписи (HTML) по порядку листов и
    #: скрытый реестр подписантов — панель внизу окна правит те же поля.
    stamps_changed = Signal(list, str)
    #: Сменился активный лист — индекс в списке основных надписей.
    active_stamp_changed = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._bridge = PreviewBridge()
        self._bridge.field_changed.connect(self.field_changed)
        self._bridge.signature_clicked.connect(self.signature_clicked)
        self._bridge.paragraph_changed.connect(self._on_paragraph_changed)
        self._bridge.table_cell_changed.connect(self._on_table_cell_changed)
        self._bridge.zoom_changed.connect(self._on_zoom_changed)
        self._bridge.scroll_changed.connect(self._on_scroll_changed)
        self._bridge.active_sheet_changed.connect(self.active_stamp_changed)
        #: Основные надписи листов текущего набора (HTML, по порядку листов).
        self._stamps: list = []
        #: Реестр подписантов (скрытые span) — чтобы панель внизу окна сразу
        #: показывала подпись по фамилии, не дожидаясь перерисовки листа.
        self._signer_registry: str = ""
        # Несохранённые правки абзацев ПЗ: {путь DOCX: {индекс: текст}}.
        # Применяются поверх текста файла, чтобы правки не терялись при
        # перерисовке листов и при переключении по дереву.
        self._pz_edits: dict = {}
        # Несохранённые правки ячеек таблиц ПЗ: {путь DOCX: {"т:с:я": текст}}.
        self._pz_cell_edits: dict = {}
        # Автосохранение: правка уходит в связанный DOCX сама, через паузу
        # после остановки ввода (а также при переключении листа и сборке).
        self._autosave = QTimer(self)
        self._autosave.setSingleShot(True)
        self._autosave.setInterval(AUTOSAVE_DELAY_MS)
        self._autosave.timeout.connect(self.flush_autosave)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # ---- Панель управления масштабом ----
        bar = QHBoxLayout()
        bar.setContentsMargins(6, 4, 6, 4)
        bar.setSpacing(4)
        self.btn_zoom_out = QPushButton("−")
        self.btn_zoom_in = QPushButton("+")
        self.btn_100 = QPushButton("100%")
        self.btn_fit = QPushButton("Вписать в окно")
        self.lbl_zoom = QLabel("100%")
        self.btn_zoom_out.clicked.connect(lambda: self._zoom(-ZOOM_STEP))
        self.btn_zoom_in.clicked.connect(lambda: self._zoom(+ZOOM_STEP))
        self.btn_100.clicked.connect(lambda: self._set_zoom(1.0))
        self.btn_fit.clicked.connect(self._fit_to_window)
        # Кнопка сохранения правок ПЗ, сделанных прямо в предпросмотре.
        # Скрыта, пока правок нет.
        self.btn_save_pz = QPushButton("Сохранить правки ПЗ")
        self.btn_save_pz.setToolTip(
            "Записать изменения текста пояснительной записки в исходный DOCX")
        self.btn_save_pz.clicked.connect(self.save_pz_requested)
        self.btn_save_pz.setVisible(False)
        self.btn_repaginate = QPushButton("Пересчитать страницы")
        self.btn_repaginate.setVisible(False)
        self.btn_repaginate.clicked.connect(self._on_repaginate_clicked)
        # Индикатор числа листов в предпросмотре (документ ПЗ показывается
        # целиком, несколькими листами подряд).
        self.lbl_sheets = QLabel("Листов: 1")
        bar.addWidget(QLabel("Масштаб:"))
        bar.addWidget(self.btn_zoom_out)
        bar.addWidget(self.btn_zoom_in)
        bar.addWidget(self.btn_100)
        bar.addWidget(self.btn_fit)
        bar.addStretch(1)

        bar.addWidget(self.btn_save_pz)
        bar.addWidget(self.btn_repaginate)
        bar.addWidget(self.lbl_sheets)
        bar.addWidget(self.lbl_zoom)
        lay.addLayout(bar)

        self.preview_stack = QStackedLayout()
        self.preview_stack.setStackingMode(QStackedLayout.StackingMode.StackAll)
        self.view = QWebEngineView()
        self.preview_stack.addWidget(self.view)
        # Листы пояснительной записки показываются настоящим Word: файл уже
        # собран с оглавлением и номерами листов (см. core.word_preview).
        self.word_host = WordHost()
        self.preview_stack.addWidget(self.word_host)
        self.word_host.hide()
        self._word_mode = False
        self._word_path = ""
        #: Проект текущего показа — нужен кнопке «Пересобрать листы».
        self._project = None
        self._word_error = ""
        self.progress_overlay = QWidget()
        self.progress_overlay.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.progress_overlay.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.progress_overlay.setStyleSheet("background:rgba(245,247,250,210);")
        overlay_layout = QVBoxLayout(self.progress_overlay)
        overlay_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        overlay_layout.setContentsMargins(20, 20, 20, 20)
        self.progress_card = QWidget()
        self.progress_card.setFixedWidth(380)
        self.progress_card.setStyleSheet(
            "QWidget { background:#ffffff; border:1px solid #d0d5dd; "
            "border-radius:10px; }")
        card_layout = QVBoxLayout(self.progress_card)
        card_layout.setContentsMargins(24, 20, 24, 20)
        card_layout.setSpacing(12)
        self.progress_title = QLabel("Пересчёт страниц ПЗ")
        self.progress_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_title.setStyleSheet(
            "border:0; background:transparent; color:#344054; font-size:15px; "
            "font-weight:600;")
        self.progress_message = QLabel("Подготовка документа…")
        self.progress_message.setWordWrap(True)
        self.progress_message.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_message.setStyleSheet(
            "border:0; background:transparent; color:#475467; font-size:12px;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(10)
        card_layout.addWidget(self.progress_title)
        card_layout.addWidget(self.progress_message)
        card_layout.addWidget(self.progress_bar)
        overlay_layout.addWidget(self.progress_card)
        self.preview_stack.addWidget(self.progress_overlay)
        self.progress_overlay.hide()
        lay.addLayout(self.preview_stack, stretch=1)

        # Разрешаем странице читать локальные файлы: логотип графы 9 и
        # изображения подписей лежат на диске. Страница грузится через
        # setHtml(about:blank), и по умолчанию QtWebEngine такие запросы
        # блокирует — картинки молча не отображались.
        try:
            from PySide6.QtWebEngineCore import QWebEngineSettings
            st = self.view.settings()
            st.setAttribute(
                QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls,
                True)
            st.setAttribute(
                QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls,
                True)
        except Exception:
            pass

        self.channel = QWebChannel()
        self.channel.registerObject("bridge", self._bridge)
        self.view.page().setWebChannel(self.channel)

        self._current_html = ""
        self._pending_snapshot = {}
        self._pending_kind = "cover"
        self._zoom = 1.0
        #: Положение прокрутки (приходит из JS) и признак «лист перерисован
        #: после правки»: в этом случае масштаб и место прокрутки сохраняются,
        #: иначе вид перескакивал бы на начало тома после каждой правки поля.
        self._last_scroll = 0.0
        self._keep_view = False
        #: Вписывать по ПЕРВОМУ листу (графическая часть: перечень листов A4,
        #: остальные листы раздела — чертёжные A1/A2, они в разы больше).
        self._fit_first = False

        # Прокрутка к началу после загрузки
        self.view.loadFinished.connect(self._on_load)

    # --- Загрузка листа ---
    def show_sheet(self, kind: str, snapshot: dict, project=None) -> None:
        """Отрисовывает ВСЕ листы выбранного пункта дерева по снимку данных.

        Для документа ПЗ, раздела ПЗ/графики/приложений или их документа
        показываются все листы подряд — с той же пагинацией и сквозной
        нумерацией, что и при сборке. Для стандартных листов ГОСТ — один лист.
        """
        # Тот же самый лист (тот же раздел и тот же документ) перерисовывается
        # после правки — тогда сохраняем масштаб и место прокрутки.
        self._keep_view = (self._pending_kind == kind
                           and self._pending_snapshot.get("section")
                           == snapshot.get("section")
                           and self._pending_snapshot.get("path")
                           == snapshot.get("path"))
        self._pending_kind = kind
        self._pending_snapshot = dict(snapshot)
        self._project = project
        self.btn_repaginate.setVisible(snapshot.get("section") == "Пояснительная записка")
        self.btn_repaginate.setEnabled(not bool(snapshot.get("pagination_busy")))
        self.btn_repaginate.setText("Пересчёт…" if snapshot.get("pagination_busy") else
                                    "Пересчитать страницы")
        stale = bool(snapshot.get("pagination_stale"))
        self.btn_repaginate.setStyleSheet(
            "QPushButton { background:#f59e0b; color:#1f2937; font-weight:600; }"
            if stale else "")
        self.btn_repaginate.setToolTip(snapshot.get("pagination_error") or
                                       "Пересчитать разбиение DOCX средствами Word")
        self._pending_snapshot["pz_edits"] = self.pz_pending_edits()
        self._pending_snapshot["pz_cell_edits"] = self.pz_pending_cell_edits()
        # Пояснительная записка показывается НАСТОЯЩИМ Word: файл собран тем же
        # кодом, что и раздел, — в нём уже есть оглавление с номерами листов.
        if self._show_word_preview(kind, snapshot, project):
            return
        self._leave_word_mode()
        # Графическая часть: вписываем в окно первый лист — перечень листов.
        # Чертежи (A1/A2) в разы больше, и масштаб по самому большому листу
        # делал перечень нечитаемо мелким.
        from ..models import SECTION_GRAPHIC
        self._fit_first = snapshot.get("section") == SECTION_GRAPHIC
        html, count = self._build_html(kind, self._pending_snapshot, project)
        self._current_html = html
        self.lbl_sheets.setText(f"Листов: {count}")
        self.view.setHtml(html, QUrl("about:blank"))
        # Панель внизу окна показывает и правит основную надпись активного листа.
        self.stamps_changed.emit(list(self._stamps), self._signer_registry)

    def _on_repaginate_clicked(self) -> None:
        """Кнопка справа от предпросмотра.

        В режиме Word (пояснительная записка) она пересобирает документ с
        оглавлением и номерами листов; в HTML-режиме — пересчитывает разбиение
        средствами Word, как раньше.
        """
        if self._word_mode:
            self.rebuild_word_preview(self._project)
            return
        self.repaginate_requested.emit()

    def _show_word_preview(self, kind: str, snapshot: dict, project=None) -> bool:
        """Показывает листы ПЗ во встроенном Word. True — режим Word включён.

        Файл для Word — рабочая копия раздела ПЗ (core.word_preview): она
        собрана штатным кодом раздела, поэтому в ней уже есть автоматическое
        оглавление с номерами листов и основные надписи. Правки в Word
        остаются в этой копии и не меняют исходный DOCX.

        Пересборка — ТОЛЬКО по кнопке «Пересобрать листы»: если настройки
        изменились, документ помечается устаревшим, но не пересобирается сам —
        иначе несохранённые правки в Word пропадали бы без предупреждения.

        Word недоступен (не установлен, сбой COM) — возвращаем False, и
        предпросмотр откатывается на прежнюю HTML-вёрстку.
        """
        from ..models import SECTION_PZ
        if snapshot.get("section") != SECTION_PZ:
            return False
        if kind not in ("doc", "section"):
            return False
        if project is None:
            return False
        try:
            from ..core.word_preview import (current_preview_path,
                                             build_preview_docx, is_stale)
        except Exception:
            return False
        path = current_preview_path(project, SECTION_PZ)
        if not path:
            # Первый показ: документа ещё нет, собираем.
            self.progress_title.setText("Подготовка листов ПЗ")
            self.progress_message.setText(
                "Собираю документ с оглавлением и номерами листов…")
            self.progress_overlay.setVisible(True)
            self.progress_overlay.raise_()
            QApplication.processEvents()
            try:
                path = build_preview_docx(project, SECTION_PZ)
            except Exception as e:
                self._word_error = (f"Не удалось собрать документ ПЗ: "
                                    f"{type(e).__name__}: {e}")
                return False
            finally:
                self.progress_overlay.setVisible(False)
        self._word_path = path
        if not self.word_host.is_open() or self.word_host.current_path() != os.path.abspath(path):
            # Прежний документ закрываем сами и без диалогов: если файл открыт,
            # Word показал бы «Файл уже используется».
            self.word_host.close_document(save=True)
            if not self.word_host.open_document(path):
                self._word_error = self.word_host.message()
                return False
        self._enter_word_mode()
        stale = False
        try:
            stale = bool(is_stale(project, SECTION_PZ))
        except Exception:
            stale = False
        self._set_word_stale(stale)
        self.lbl_sheets.setText(
            "Лист Word" + (" · требует пересборки" if stale else ""))
        return True

    def _set_word_stale(self, stale: bool) -> None:
        """Помечает кнопку пересборки, когда документ отстал от настроек."""
        self.btn_repaginate.setText("Пересобрать листы")
        self.btn_repaginate.setStyleSheet(
            "QPushButton { background:#f59e0b; color:#1f2937; font-weight:600; }"
            if stale else "")
        self.btn_repaginate.setToolTip(
            ("Настройки изменились — документ собран раньше. " if stale else "")
            + "Собрать документ ПЗ заново, с оглавлением и номерами листов. "
              "Несохранённые правки в Word будут сохранены в рабочую копию.")

    def _enter_word_mode(self) -> None:
        """Переключает предпросмотр на окно Word (скрывает HTML-вид)."""
        self._word_mode = True
        self.view.hide()
        self.word_host.show()
        self.word_host.raise_()
        # Масштаб и «вписать в окно» относятся к HTML-листу: в Word их роль
        # выполняет сам Word (масштаб — в его строке состояния).
        for widget in (self.btn_zoom_out, self.btn_zoom_in, self.btn_100,
                       self.btn_fit, self.lbl_zoom):
            widget.setVisible(False)
        self.btn_repaginate.setVisible(True)
        self.btn_repaginate.setText("Пересобрать листы")
        self.btn_repaginate.setToolTip(
            "Собрать документ ПЗ заново — с оглавлением и номерами листов. "
            "Несохранённые правки в Word будут потеряны.")

    def _leave_word_mode(self) -> None:
        """Возвращает HTML-предпросмотр (другие листы тома)."""
        if not self._word_mode:
            return
        self._word_mode = False
        self.word_host.hide()
        self.view.show()
        for widget in (self.btn_zoom_out, self.btn_zoom_in, self.btn_100,
                       self.btn_fit, self.lbl_zoom):
            widget.setVisible(True)

    def rebuild_word_preview(self, project=None) -> bool:
        """Пересобирает рабочую копию и переоткрывает её в Word.

        Вызывается кнопкой «Пересобрать листы»: правки в Word сначала
        сохраняются в рабочую копию, затем файл собирается заново — уже с
        оглавлением и номерами листов по текущим настройкам.
        """
        from ..models import SECTION_PZ
        project = project or self._project
        if project is None or not self._word_mode:
            return False
        self.word_host.close_document(save=True)
        self.progress_title.setText("Пересборка листов ПЗ")
        self.progress_message.setText(
            "Собираю документ заново — с оглавлением и номерами листов…")
        self.progress_overlay.setVisible(True)
        self.progress_overlay.raise_()
        QApplication.processEvents()
        try:
            from ..core.word_preview import build_preview_docx
            path = build_preview_docx(project, SECTION_PZ, force=True)
        except Exception as e:
            self.progress_overlay.setVisible(False)
            self.word_host._show_error(
                f"Не удалось пересобрать: {type(e).__name__}: {e}")
            return False
        finally:
            self.progress_overlay.setVisible(False)
        self._word_path = path
        opened = self.word_host.open_document(path)
        self._set_word_stale(not opened)
        if opened:
            self.lbl_sheets.setText("Лист Word")
        return opened

    def _build_html(self, kind: str, snapshot: dict, project=None):
        """Собирает HTML предпросмотра. Возвращает (html, число листов)."""
        # Правка текста ПЗ доступна и для документа ПЗ, и для раздела ПЗ:
        # в разделе каждый DOCX получает свои data-para-file, поэтому правки
        # адресуются точно (файл + индекс абзаца).
        if snapshot.get("section") == "Пояснительная записка":
            if kind == "doc":
                path = snapshot.get("path", "")
                if path.lower().endswith((".docx", ".doc")):
                    snapshot["pz_editable"] = True
            elif kind == "section":
                snapshot["pz_editable"] = True
        try:
            fragments = render_sheet_fragments(kind, snapshot, project=project)
        except Exception as e:
            # Раньше исключение молча превращалось в заглушку «Не удалось
            # отобразить содержимое файла» — причина сбоя терялась. Теперь она
            # пишется в tmp/build.log и выводится на самом листе.
            reason = f"{type(e).__name__}: {e}"
            log_path = _log_preview_error(kind, reason)
            self._stamps = []
            self._signer_registry = ""
            return render_preview_error(kind, snapshot, reason, log_path), 1
        if not fragments:
            self._stamps = []
            self._signer_registry = ""
            return render_preview_error(
                kind, snapshot,
                "не удалось получить ни одного листа для выбранного пункта"), 1
        from ..core.preview_html import (_preview_document, extract_stamp_html,
                                         _signer_registry_html)
        # Основные надписи листов — для панели внизу окна (у обложки и титула
        # штампа нет, в списке остаётся пустая строка).
        self._stamps = [extract_stamp_html(f) for f in fragments]
        self._signer_registry = _signer_registry_html(snapshot)
        return _preview_document(fragments, snapshot), len(fragments)

    def refresh(self, kind: str, snapshot: dict, project=None) -> None:
        """Перерисовать текущий лист после изменения данных."""
        self.show_sheet(kind, snapshot, project)

    def set_pagination_state(self, *, stale: bool, busy: bool, error: str = "",
                             message: str = "") -> None:
        self.btn_repaginate.setEnabled(not busy)
        if self._word_mode:
            # В режиме Word кнопка пересобирает документ, а не пересчитывает
            # HTML-разбиение: текст «Пересчёт…» и индикатор здесь не нужны.
            self._set_word_stale(bool(stale))
            self.progress_overlay.setVisible(False)
            return
        self.btn_repaginate.setText("Пересчёт…" if busy else "Пересчитать страницы")
        self.btn_repaginate.setStyleSheet(
            "QPushButton { background:#f59e0b; color:#1f2937; font-weight:600; }"
            if stale else "")
        self.btn_repaginate.setToolTip(error or
                                       "Пересчитать разбиение DOCX средствами Word")
        show_overlay = bool(busy and self._pending_snapshot.get("section") ==
                            "Пояснительная записка")
        self.progress_overlay.setVisible(show_overlay)
        if show_overlay:
            self.progress_message.setText(message or "Word рассчитывает разбиение страниц…")

    def _on_load(self, ok: bool) -> None:
        if self._keep_view:
            # Лист перерисован после правки: возвращаем прежний масштаб и место
            # прокрутки, иначе пользователя «перебрасывало» на начало тома.
            QTimer.singleShot(60, self._restore_view)
        else:
            QTimer.singleShot(100, self._fit_to_window)

    def _restore_view(self) -> None:
        """Возвращает масштаб и положение прокрутки, бывшие до перерисовки."""
        self.view.page().runJavaScript(
            "document.body.style.zoom = '%f'; window.scrollTo(0, %f);"
            % (self._zoom, self._last_scroll))

    def _on_scroll_changed(self, position: float) -> None:
        """JS сообщил положение прокрутки — запоминаем для возврата после правки."""
        self._last_scroll = float(position)

    def _apply_zoom(self) -> None:
        """Применяет текущий масштаб к телу страницы."""
        self.view.page().runJavaScript(
            "document.body.style.zoom = '%f';" % self._zoom)

    def _set_zoom(self, z: float) -> None:
        self._zoom = max(0.3, min(3.0, round(z, 2)))
        self.lbl_zoom.setText(f"{int(self._zoom * 100)}%")
        self._apply_zoom()

    def _zoom(self, delta: float) -> None:
        self._set_zoom(self._zoom + delta)

    def _fit_to_window(self) -> None:
        """Вписывает лист в окно.

        Для обычных разделов (записка, приложения) масштаб подбирается по
        самому широкому и самому высокому листу, чтобы ни один лист не выходил
        за границы окна по горизонтали.

        Для ГРАФИЧЕСКОЙ ЧАСТИ вписывается ПЕРВЫЙ лист — перечень листов (A4).
        Чертёжные листы (A1, A2) в разы больше, и масштаб по самому большому
        листу делал перечень нечитаемо мелким. Остальные листы раздела
        листаются прокруткой.
        """
        fit_first = "true" if self._fit_first else "false"
        self.view.page().runJavaScript(
            """
            (function(fitFirst){
              var sheets = document.querySelectorAll('.sheet');
              if (!sheets.length) return;
              var w = 0, h = 0;
              if (fitFirst) {
                // Графическая часть: масштаб по первому листу (перечень).
                w = sheets[0].offsetWidth;
                h = sheets[0].offsetHeight;
              } else {
                sheets.forEach(function(s){
                  w = Math.max(w, s.offsetWidth);
                  h = Math.max(h, s.offsetHeight);
                });
              }
              if (!w || !h) return;
              var vw = window.innerWidth, vh = window.innerHeight;
              var margin = 32;
              // Внешние поля .sheet (12/16 мм) — тоже часть макета: без их учёта
              // масштаб получался больше нужного, и правая часть листа (вместе с
              // рамкой) уходила за границу окна — появлялся горизонтальный скролл.
              var cs = getComputedStyle(sheets[0]);
              var mx = (parseFloat(cs.marginLeft) || 0) + (parseFloat(cs.marginRight) || 0);
              var my = (parseFloat(cs.marginTop) || 0) + (parseFloat(cs.marginBottom) || 0);
              var scale = Math.min((vw - margin) / (w + mx), (vh - margin) / (h + my));
              if (scale > 0.1) {
                document.body.style.zoom = scale;
                // обновить индикатор масштаба
                window.bridge && window.bridge.onZoomChanged && window.bridge.onZoomChanged(scale);
              } else {
                document.body.style.zoom = 1;
              }
            })(%s);
            """ % fit_first)

    # --- Служебное ---
    def set_html(self, html: str) -> None:
        self._current_html = html
        self.view.setHtml(html, QUrl("about:blank"))

    def current_html(self) -> str:
        """Возвращает HTML текущего отображаемого листа (для экспорта)."""
        return self._current_html

    def _on_zoom_changed(self, scale: float) -> None:
        """Обновляет индикатор масштаба после автоподгонки под окно."""
        self._zoom = scale
        self.lbl_zoom.setText(f"{int(scale * 100)}%")

    # --- Inline-правка текста ПЗ ---
    def _on_paragraph_changed(self, index: int, path: str, text: str) -> None:
        """Абзац ПЗ отредактирован в предпросмотре — копим правку и планируем запись."""
        self._pz_edits.setdefault(path, {})[int(index)] = text
        self._update_save_button()
        self._schedule_autosave()
        self.paragraph_changed.emit(int(index), path, text)

    def _on_table_cell_changed(self, table: int, row: int, cell: int,
                               path: str, text: str) -> None:
        """Ячейка таблицы ПЗ отредактирована — копим правку и планируем запись."""
        address = f"{int(table)}:{int(row)}:{int(cell)}"
        self._pz_cell_edits.setdefault(path, {})[address] = text
        self._update_save_button()
        self._schedule_autosave()
        self.cell_changed.emit(int(table), int(row), int(cell), path, text)

    # --- Автосохранение правок в связанный DOCX ---
    def _schedule_autosave(self) -> None:
        """Перезапускает паузу автосохранения (ждём остановки ввода)."""
        self._autosave.start()

    def flush_autosave(self) -> None:
        """Немедленно записывает накопленные правки (без паузы).

        Вызывается по таймеру, а также принудительно — перед переключением
        листа, сборкой и закрытием, чтобы правки не потерялись.
        """
        if self._autosave.isActive():
            self._autosave.stop()
        if self.has_pz_edits():
            self.autosave_requested.emit()

    def cancel_autosave(self) -> None:
        """Отменяет отложенную запись (например, при смене проекта)."""
        self._autosave.stop()

    def _update_save_button(self) -> None:
        """Показывает кнопку сохранения, когда есть несохранённые правки ПЗ."""
        count = self.pending_edit_count()
        self.btn_save_pz.setVisible(count > 0)
        if count:
            self.btn_save_pz.setText(f"Сохранить правки ПЗ ({count})")

    def pz_pending_edits(self) -> dict:
        """Несохранённые правки абзацев: {путь DOCX: {индекс: текст}}."""
        return {p: dict(v) for p, v in self._pz_edits.items() if v}

    def pz_pending_cell_edits(self) -> dict:
        """Несохранённые правки ячеек таблиц: {путь DOCX: {"т:с:я": текст}}."""
        return {p: dict(v) for p, v in self._pz_cell_edits.items() if v}

    def pending_edit_count(self) -> int:
        """Общее число несохранённых правок ПЗ (абзацы + ячейки таблиц)."""
        return (sum(len(v) for v in self._pz_edits.values())
                + sum(len(v) for v in self._pz_cell_edits.values()))

    def has_pz_edits(self) -> bool:
        """Есть ли несохранённые правки текста ПЗ (абзацы или таблицы)."""
        return self.pending_edit_count() > 0

    def clear_pz_edits(self, path: str | None = None) -> None:
        """Сбрасывает учёт правок (после сохранения или при смене проекта)."""
        if path is None:
            self._pz_edits.clear()
            self._pz_cell_edits.clear()
            self._autosave.stop()
        else:
            self._pz_edits.pop(path, None)
            self._pz_cell_edits.pop(path, None)
        self._update_save_button()

    def closeEvent(self, event) -> None:
        """Закрывает встроенный Word вместе с предпросмотром.

        Иначе процесс WINWORD остаётся в памяти с окном, встроенным в уже
        уничтоженный виджет.
        """
        try:
            self.word_host.close_document()
        except Exception:
            pass
        super().closeEvent(event)
