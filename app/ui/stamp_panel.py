# -*- coding: utf-8 -*-
"""Панель основной надписи активного листа внизу главного окна.

Показывает нижний штамп (основную надпись) того листа, который сейчас
занимает большую часть окна предпросмотра. Масштаб — 100%: размеры штампа
заданы в миллиметрах, а CSS-миллиметр равен физическому миллиметру при
штатном DPI экрана, поэтому надпись видна в натуральную величину.

Поля штампа в панели **редактируются так же, как на самом листе**: правка
уходит в тот же QWebChannel-мост (PreviewBridge) и в те же обработчики
главного окна, что и правка в предпросмотре. Клик по полю подписи открывает
выбор подписанта, правка фамилии сразу подставляет подпись из реестра.

Панель сворачивается в узкую полосу внизу окна — кнопкой в её заголовке.
Свёрнутое состояние запоминается в config.json (`Config.stamp_panel_open`).
"""
from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame,
)
from PySide6.QtWebChannel import QWebChannel

from ..core.preview_html import _css, stamp_size_mm
from .preview_bridge import PreviewBridge

#: Пикселей в миллиметре для CSS (1 CSS мм = 1/25.4 дюйма при 96 dpi).
PX_PER_MM = 96.0 / 25.4
#: Поля вокруг штампа внутри панели, мм.
PANEL_PAD_MM = 4.0
#: Высота свёрнутой панели (полоса внизу окна), px.
COLLAPSED_HEIGHT = 26


class StampPanel(QWidget):
    """Нижняя панель с основной надписью активного листа (масштаб 100%)."""

    #: Поле штампа отредактировано (key, value) — как в предпросмотре.
    field_changed = Signal(str, str)
    #: Клик по полю подписи (key) — открыть выбор подписанта.
    signature_clicked = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._stamps: list = []
        self._active = -1
        self._open = True
        self._stamp_html = ""
        self._registry = ""
        # Мост тот же, что у предпросмотра: правки панели и листа идут в
        # одни и те же обработчики главного окна.
        self._bridge = PreviewBridge()
        self._bridge.field_changed.connect(self.field_changed)
        self._bridge.signature_clicked.connect(self.signature_clicked)

        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        # ---- Заголовок: название, масштаб и кнопка сворачивания ----
        self.header = QFrame()
        self.header.setObjectName("stampHeader")
        self.header.setStyleSheet(
            "#stampHeader{background:#eef2f7;border-top:1px solid #c9d2dc;}")
        h = QHBoxLayout(self.header)
        h.setContentsMargins(8, 2, 6, 2)
        h.setSpacing(6)
        self.lbl_title = QLabel("Основная надпись листа")
        self.lbl_title.setStyleSheet("font-weight:bold;")
        self.lbl_size = QLabel("")
        self.lbl_size.setStyleSheet("color:#555;")
        self.lbl_source = QLabel("масштаб 100%")
        self.lbl_source.setStyleSheet("color:#555;")
        self.btn_toggle = QPushButton("Свернуть")
        self.btn_toggle.setFixedHeight(20)
        self.btn_toggle.setToolTip(
            "Свернуть панель в полосу внизу окна или развернуть обратно")
        self.btn_toggle.clicked.connect(self.toggle)
        h.addWidget(self.lbl_title)
        h.addWidget(self.lbl_size)
        h.addStretch(1)
        h.addWidget(self.lbl_source)
        h.addWidget(self.btn_toggle)
        v.addWidget(self.header)

        # ---- Сам штамп ----
        from PySide6.QtWebEngineWidgets import QWebEngineView
        self.view = QWebEngineView()
        self.view.setMinimumHeight(0)
        # Разрешаем читать локальные файлы: подписи и логотип графы 9 лежат на
        # диске, а страница грузится через setHtml(about:blank).
        try:
            from PySide6.QtWebEngineCore import QWebEngineSettings
            st = self.view.settings()
            st.setAttribute(
                QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls,
                True)
        except Exception:
            pass
        # Мост нужен ДО загрузки страницы: иначе JS не получит window.bridge
        # и поля останутся нередактируемыми.
        self.channel = QWebChannel()
        self.channel.registerObject("bridge", self._bridge)
        self.view.page().setWebChannel(self.channel)
        v.addWidget(self.view, stretch=1)

        self._set_open(self._open, save=False)
        self._show_message("Загрузка предпросмотра…")

    # ---- внешнее API ----
    def set_stamps(self, stamps: list, registry: str = "") -> None:
        """Новый набор основных надписей (по порядку листов) и реестр подписантов.

        Активный лист сохраняется: предпросмотр перерисовывается и после правки
        поля, и при смене раздела — панель не должна «перескакивать» на первый
        лист, если пользователь смотрел другой.
        """
        self._stamps = list(stamps or [])
        self._registry = registry or ""
        keep = self._active
        self._stamp_html = ""
        self.set_active(keep if keep >= 0 else 0)

    def set_active(self, index: int) -> None:
        """Показать надпись листа с указанным индексом."""
        if not self._stamps:
            self._stamp_html = ""
            self.lbl_size.setText("")
            self._show_message("Нет данных предпросмотра")
            return
        index = max(0, min(int(index), len(self._stamps) - 1))
        self._active = index
        html = self._stamps[index]
        if not html:
            self._stamp_html = ""
            self.lbl_size.setText("")
            self._show_message(
                "Основной надписи на этом листе нет (обложка, титул)")
            return
        if html == self._stamp_html:
            return
        self._stamp_html = html
        w_mm, h_mm = stamp_size_mm(html)
        self.lbl_size.setText(f"{w_mm:g} × {h_mm:g} мм")
        self.view.setHtml(self._html(html), self._base_url())
        self._apply_size(h_mm)
        self.btn_toggle.setEnabled(True)

    def is_open(self) -> bool:
        return self._open

    def set_open(self, opened: bool) -> None:
        """Развернуть/свернуть панель (состояние сохраняется в config.json)."""
        self._set_open(bool(opened), save=True)

    def toggle(self) -> None:
        self._set_open(not self._open, save=True)

    def restore_state(self) -> None:
        """Применяет сохранённое состояние из конфигурации при запуске."""
        from ..config import load_config
        try:
            opened = bool(load_config().stamp_panel_open)
        except Exception:
            opened = True
        self._set_open(opened, save=False)

    # ---- внутреннее ----
    def _commit_pending(self) -> None:
        """Завершает ввод в поле панели (значение уходит в модель по focusout).

        Нужно перед сворачиванием панели: правка сохраняется по потере
        фокуса, а скрытое поле фокус уже не теряет — без этого последняя
        недописанная ячейка пропала бы.
        """
        self.view.page().runJavaScript(
            "(function(){var el=document.activeElement;"
            "if(el&&el.blur){el.blur();}"
            "if(el&&el.dispatchEvent){"
            "el.dispatchEvent(new FocusEvent('focusout',{bubbles:true}));}"
            "return 1;})();")

    def _set_open(self, opened: bool, save: bool) -> None:
        self._open = bool(opened)
        if not self._open and self._stamp_html:
            self._commit_pending()
        self.view.setVisible(self._open)
        self.btn_toggle.setText("Свернуть" if self._open else "Развернуть")
        self.btn_toggle.setToolTip(
            "Свернуть панель в полосу внизу окна" if self._open
            else "Развернуть панель основной надписи")
        if self._open:
            self._apply_size_from_current()
        else:
            self.setFixedHeight(COLLAPSED_HEIGHT)
        if save:
            self._save_state()

    def _apply_size_from_current(self) -> None:
        """Высота развёрнутой панели — по размеру штампа (масштаб 100%)."""
        if not self._stamp_html:
            # Штампа нет (или он ещё не пришёл): панель занимает только
            # строку-пояснение, а не половину окна.
            head = max(self.header.sizeHint().height(), COLLAPSED_HEIGHT)
            self.setFixedHeight(head + 40)
            return
        _w_mm, h_mm = stamp_size_mm(self._stamp_html)
        self._apply_size(h_mm)

    def _apply_size(self, h_mm: float) -> None:
        """Фиксирует высоту панели по высоте штампа плюс поля."""
        if not self._open:
            return
        head = max(self.header.sizeHint().height(), COLLAPSED_HEIGHT)
        need = int(round((h_mm + 2 * PANEL_PAD_MM) * PX_PER_MM)) + head
        self.setFixedHeight(max(head + 24, need))

    def _show_message(self, text: str) -> None:
        """Вместо штампа — пояснение (нет надписи / нет данных)."""
        self.view.setHtml(
            "<!doctype html><html><head><meta charset='utf-8'></head>"
            "<body style=\"margin:0;background:#e8e8e8;font-family:'Segoe UI',"
            "Arial,sans-serif;color:#666;\">"
            f"<div style='padding:10px 12px;font-size:12px'>{text}</div>"
            "</body></html>", self._base_url())
        if self._open:
            head = max(self.header.sizeHint().height(), COLLAPSED_HEIGHT)
            self.setFixedHeight(head + 40)

    @staticmethod
    def _base_url():
        from PySide6.QtCore import QUrl
        return QUrl("about:blank")

    def _html(self, stamp_html: str) -> str:
        """Документ панели: штамп на сером фоне в натуральном размере.

        Поля остаются редактируемыми: скрипт подключается к тому же мосту,
        что и предпросмотр, поэтому правка уходит в модель теми же ключами
        (`organization`, `volume_id`, `ss|<раздел>|<строка>|surname` и т.д.).
        """
        # CSS собирается конкатенацией: в нём есть символы «%» (width:100%),
        # из-за которых форматирование через % или .format() падает.
        return (
            "<!doctype html><html><head><meta charset=\"utf-8\">"
            "<script src=\"qrc:///qtwebchannel/qwebchannel.js\"></script>"
            "<style>" + _css() + "</style>"
            "<style>"
            "html,body{margin:0;padding:0;background:#e8e8e8;}"
            "/* Штамп по центру панели; в предпросмотре он позиционируется "
            "от листа, здесь стоит в потоке. */"
            "body{text-align:center;}"
            ".host{padding:" + f"{PANEL_PAD_MM:g}" + "mm;display:inline-block;"
            "text-align:initial;}"
            ".host .stamp{position:relative;right:auto;bottom:auto;"
            "background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.35);}"
            ".host .field{outline:none;}"
            "</style></head><body>"
            + self._registry
            + "<div class=\"host\">" + stamp_html + "</div>"
            + _PANEL_SCRIPT + "</body></html>")

    def _save_state(self) -> None:
        from ..config import load_config, save_config
        try:
            cfg = load_config()
        except Exception:
            return
        cfg.stamp_panel_open = self._open
        try:
            save_config(cfg)
        except Exception:
            pass


#: Скрипт правки штампа в панели. Повторяет обработчики предпросмотра,
#: которые относятся к штампу: подсказки пустых полей, автоподгонка кегля
#: (графа 9), отправка правки в модель, клик по подписи и подстановка
#: подписи по фамилии из реестра. Перелив текста по листам и правка абзацев
#: ПЗ здесь не нужны — в панели только основная надпись.
_PANEL_SCRIPT = """
<script>
(function() {
  // Подсказка поля (data-placeholder) выводится правилом .field.empty::after.
  // Класс empty снимается, как только в поле появился текст: иначе введённое
  // значение остаётся серым курсивом (у полей без подсказки атрибута нет, но
  // класс empty ставится при пустом значении).
  function syncEmpty(field) {
    if (!field || !field.classList) { return; }
    if (!field.hasAttribute('data-placeholder') &&
        !field.classList.contains('empty')) { return; }
    field.classList.toggle('empty', !field.innerText.trim());
  }
  // Автоподгонка кегля: уменьшает font-size, пока содержимое не помещается.
  function fitLabels() {
    var nodes = document.querySelectorAll('.fit-lbl');
    nodes.forEach(function(el) {
      var text = el.textContent;
      if (!text) { return; }
      var m = /font-size:([\\d.]+)pt/.exec(el.getAttribute('style') || '');
      var start = m ? parseFloat(m[1]) : 7;
      var target = el.querySelector('.field') || el;
      var found = false;
      for (var fs = start; fs >= 3.5; fs -= 0.5) {
        target.style.fontSize = fs + 'pt';
        if (el.scrollWidth <= el.clientWidth + 1 &&
            el.scrollHeight <= el.clientHeight + 1) { found = true; break; }
      }
      if (!found) { target.style.fontSize = Math.max(3.5, start) + 'pt'; }
    });
  }
  function sigImageFor(surname) {
    var s = String(surname || '').trim().toLowerCase();
    if (!s) { return ''; }
    var found = '';
    document.querySelectorAll('[data-signer-surname]').forEach(function(el) {
      if (found) { return; }
      if ((el.getAttribute('data-signer-surname') || '').toLowerCase() === s) {
        found = el.getAttribute('data-signer-img') || '';
      }
    });
    return found;
  }
  var EMPTY_SIG = '<span class="sig-placeholder">(подпись)</span>';
  function syncSigPreview(rowKey) {
    var nameCell = document.querySelector('.field[data-key="ss|' + rowKey + '|surname"]');
    var sigCell = document.querySelector('.field[data-key="ss|' + rowKey + '|signature"]');
    if (!sigCell) { return; }
    var src = sigImageFor(nameCell ? nameCell.innerText : '');
    sigCell.innerHTML = src
      ? '<img class="sig-img sig-img-row" src="' + src + '" alt="">'
      : EMPTY_SIG;
  }
  function init() {
    var fields = document.querySelectorAll('.field[contenteditable="true"]');
    fields.forEach(function(f) {
      syncEmpty(f);
      // Значение на момент получения фокуса: по нему видно, была ли правка.
      var valueAtFocus = f.innerText;
      f.addEventListener('focus', function() {
        valueAtFocus = f.innerText;
      });
      f.addEventListener('input', function() {
        // Во время ввода обновляем только вид поля (подсказка, подпись по
        // фамилии). В модель значение уходит по потере фокуса: иначе каждое
        // нажатие клавиши пересобирало бы листы и панель «дёргалась».
        syncEmpty(f);
        var key = f.getAttribute('data-key') || '';
        var parts = key.split('|');
        if (parts.length === 4 && parts[0] === 'ss' && parts[3] === 'surname') {
          syncSigPreview(parts[1] + '|' + parts[2]);
        }
      });
      var commit = function() {
        var text = f.innerText;
        if (text === valueAtFocus) { return; }   // правки не было
        valueAtFocus = text;
        if (window.bridge && window.bridge.onFieldChanged) {
          window.bridge.onFieldChanged(f.getAttribute('data-key'), text);
        }
      };
      // Сохранение — по потере фокуса. focusout добавлен к blur: в QtWebEngine
      // он приходит надёжнее, а повторный вызов безопасен (значение уже
      // записано и повторно не отправляется).
      f.addEventListener('blur', commit);
      f.addEventListener('focusout', commit);
      // Enter в поле штампа переносить строку не должен: это одна ячейка.
      // Завершаем ввод: значение сохраняется, как при уходе фокуса. commit()
      // вызывается явно — blur() в QtWebEngine не всегда порождает событие
      // (например, когда окно не активно), и правка иначе бы потерялась.
      f.addEventListener('keydown', function(e) {
        if (e.key !== 'Enter') { return; }
        e.preventDefault();
        commit();
        f.blur();
      });
    });
    // Клик по полю подписи — выбор подписанта (как на листе).
    document.querySelectorAll('.field[contenteditable="false"].signature')
      .forEach(function(f) {
        f.addEventListener('click', function() {
          if (window.bridge && window.bridge.onSignatureClicked) {
            window.bridge.onSignatureClicked(f.getAttribute('data-key'));
          }
        });
      });
    fitLabels();
    if (typeof QWebChannel === 'undefined' || typeof qt === 'undefined') { return; }
    new QWebChannel(qt.webChannelTransport, function(channel) {
      window.bridge = channel.objects.bridge;
    });
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
</script>
"""
