# -*- coding: utf-8 -*-
"""Обёртка над Word COM: конвертация DOCX -> PDF через установленный Word."""
from __future__ import annotations

import os
from contextlib import contextmanager

from .pywin32_boot import ensure_pywin32


def word_available(timeout_ms: int = 8000) -> bool:
    """Быстро проверяет, доступен ли Word (COM) в системе.

    Запускает проверку в отдельном процессе с таймаутом, чтобы НЕ зависнуть,
    если Word не установлен или COM-инициализация блокируется.

    :param timeout_ms: максимальное время ожидания (мс).
    :return: True, если Word.Application можно создать.
    """
    import subprocess
    import sys
    script = (
        "import sys; sys.path.insert(0, {v!r}); "
        "from app.core.pywin32_boot import ensure_pywin32; ensure_pywin32(); "
        "import win32com.client; "
        "w = win32com.client.DispatchEx('Word.Application'); w.Quit(); print('OK')"
    ).format(v=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "vendor"))
    try:
        proc = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True, timeout=timeout_ms / 1000.0,
        )
        return proc.returncode == 0 and b"OK" in proc.stdout
    except Exception:
        return False


@contextmanager
def word_application(visible: bool = False):
    """Открывает экземпляр Word.Application (COM) и корректно закрывает его."""
    ensure_pywin32()
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = visible
    try:
        yield word
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def docx_to_pdf(docx_path: str, pdf_path: str, visible: bool = False) -> str:
    """Конвертирует DOCX в PDF через Word COM. Возвращает путь к PDF.

    Word экспортирует документ с сохранением форматирования (максимальная
    точность). Формат PDF из Word = формат 17 (wdFormatPDF).
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(docx_path)
    ensure_pywin32()
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = visible
    try:
        doc = None
        try:
            # 0 = wdDoNotSaveChanges
            doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True,
                                      AddToRecentFiles=False, Visible=visible)
            # 17 = wdFormatPDF
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.abspath(pdf_path),
                ExportFormat=17,
                OpenAfterExport=False,
            )
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return pdf_path
