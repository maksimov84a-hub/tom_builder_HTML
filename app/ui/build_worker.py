# -*- coding: utf-8 -*-
"""Фоновая подготовка HTML и DOCX для сборки тома."""
from __future__ import annotations

import logging
import os

from PySide6.QtCore import QObject, Signal, Slot


class BuildWorker(QObject):
    progress = Signal(str)
    finished = Signal(object)

    def __init__(self, project, formats, chapters, base):
        super().__init__()
        self.project = project
        self.formats = list(formats)
        self.chapters = chapters
        self.base = base

    @Slot()
    def run(self):
        try:
            self._run()
        except Exception as e:
            logging.getLogger("tom_builder.build").exception("build worker failed")
            self.finished.emit({"done": [], "failed": [f"Сборка: {e}"],
                                "html_path": None})

    def _run(self):
        from ..config import load_config

        log = logging.getLogger("tom_builder.build")
        stem = os.path.splitext(self.base)[0]
        done, failed = [], []
        html_path = None

        if "html" in self.formats or "pdf" in self.formats or "docx" in self.formats:
            try:
                from ..models import DocKind
                from ..core.preview_html import pz_doc_first_forms
                from ..core.word_com import word_application
                from ..core.word_pagination import calculate_docx_pages
                docs = self.project.documents_for(DocKind.POYASNITELNAYA)
                forms = pz_doc_first_forms(self.project)
                # Один экземпляр Word на все записки тома (см. PaginationWorker):
                # запуск приложения стоит ~2 с на документ.
                with word_application() as word:
                    for index, doc in enumerate(docs):
                        if not doc.path or not doc.path.lower().endswith(".docx"):
                            continue
                        self.progress.emit(f"Word рассчитывает страницы: {os.path.basename(doc.path)}")
                        calculate_docx_pages(doc.path,
                                             forms[index] if index < len(forms) else 2,
                                             word=word)
            except Exception as e:
                log.exception("Word pagination failed before export")
                failed.append(f"Пагинация ПЗ: {e}")

        if "html" in self.formats or "pdf" in self.formats:
            out = stem + ".html"
            self.progress.emit("Формирование HTML тома…")
            try:
                from ..core.preview_html import build_volume_html
                build_volume_html(self.project, out)
                html_path = out
                log.info("HTML: %s", out)
                if "html" in self.formats:
                    done.append(f"HTML: {out}")
            except Exception as e:
                log.exception("build_volume_html failed")
                failed.append(f"HTML: {e}")

        if "docx" in self.formats:
            self.progress.emit("Формирование документа Word…")
            try:
                from ..core.docx_exporter import build_text_docx, build_chapters_docx
                cfg = load_config()
                if self.chapters:
                    paths = build_chapters_docx(
                        self.project, os.path.dirname(self.base), cfg)
                    log.info("DOCX по разделам: %s", paths)
                    done.append("DOC (Word) по разделам ("
                                + str(len(paths)) + " файлов):\n  "
                                + "\n  ".join(os.path.basename(p) for p in paths))
                else:
                    out = stem + ".docx"
                    actual = build_text_docx(self.project, out, cfg)
                    log.info("DOCX: %s", actual)
                    if os.path.abspath(actual) != os.path.abspath(out):
                        done.append(
                            f"DOC (Word): {actual}\n  (файл {os.path.basename(out)}"
                            " занят другой программой — сохранено рядом)")
                    else:
                        done.append(f"DOC (Word): {actual}")
            except Exception as e:
                log.exception("build_text_docx failed")
                failed.append(f"DOC (Word): {e}")

        self.finished.emit({"done": done, "failed": failed, "html_path": html_path})
