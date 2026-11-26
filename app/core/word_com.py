# -*- coding: utf-8 -*-
"""Обёртка над Word COM: конвертация DOCX -> PDF через установленный Word."""
from __future__ import annotations

import bisect
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
    """Открывает экземпляр Word.Application (COM) и корректно закрывает его.

    Незавершённая инициализация COM освобождается и при сбое запуска Word:
    контекст используют фоновые воркеры, и утечка CoInitialize оставила бы
    поток в «занятом» COM-состоянии до его завершения.
    """
    ensure_pywin32()
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = visible
        yield word
    finally:
        if word is not None:
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


def paginate_docx(docx_path: str, word=None) -> dict:
    """Возвращает число страниц и страницы абзацев/строк таблиц DOCX через Word.

    Номер страницы позиции берётся сравнением с началами страниц, а НЕ вызовом
    ``Range.Information(3)``. Обращение к COM стоит около 30 мс, и на документе
    в 250 абзацев/строк таблиц таких обращений набиралось полтысячи — это давало
    ~15 с из ~22 с общего времени пересчёта. Массив начал страниц собирается
    заранее пачкой (``GoTo`` What=1, Which=1) и используется как таблица
    поиска; позиции закладок берутся из ``Range.Start/End``, которые в COM
    бесплатны. Результат совпадает с прежним расчётом до последнего символа.

    :param word: уже открытый ``Word.Application``. Передан — используется он,
        а свой экземпляр не создаётся и не закрывается: запуск Word стоит ~2 с,
        и на нескольких документах подряд это заметная доля времени. Живёт
        обычно внутри :func:`word_application` в одном потоке с вызовом —
        COM-объект нельзя переносить между потоками без маршалинга.
    """
    ensure_pywin32()
    if word is not None:
        return _paginate_with_word(word, docx_path)
    with word_application() as own_word:
        return _paginate_with_word(own_word, docx_path)


def _paginate_with_word(word, docx_path: str) -> dict:
    """Разметка одного DOCX на переданном экземпляре Word (см. paginate_docx)."""
    doc = None
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            os.path.abspath(docx_path), ReadOnly=True,
            AddToRecentFiles=False, Visible=False,
        )
        doc.Repaginate()

        page_count = int(doc.ComputeStatistics(2))
        page_starts = []
        for page_no in range(1, page_count + 1):
            try:
                page_starts.append(int(doc.GoTo(What=1, Which=1, Count=page_no).Start))
            except Exception:
                page_starts.append(int(doc.Content.Start))
        content_start = int(doc.Content.Start)
        content_end = int(doc.Content.End)
        page_starts.append(content_end)

        # Таблица поиска «позиция -> номер страницы»: начала страниц идут по
        # возрастанию, поэтому номер — это количество начал, не превышающих
        # позицию. Дубликаты (вырожденные страницы, сбойный GoTo) НЕ убираются:
        # позиция после повторного начала принадлежит ПОСЛЕДНЕЙ из страниц,
        # начинающихся в этой точке, — ровно так отвечает и Word.
        bounds = list(page_starts[:-1])

        def page_number(position):
            index = bisect.bisect_right(bounds, position)
            if index <= 0:
                return 1
            return min(index, page_count)

        def range_pages(word_range):
            try:
                start = int(word_range.Start)
                end = int(word_range.End)
            except Exception:
                return {"start": None, "end": None}
            if start < content_start or end > content_end:
                return {"start": None, "end": None}
            if end <= start:
                # Вырожденный диапазон — пустой абзац: закладка охватывает одну
                # позицию (знак абзаца), и страница берётся по началу. Раньше
                # такой абзац получал None и пропадал из разметки, а Word
                # отводил ему на листе целую строку — низ листов оставался
                # пустым, текст не доходил до штампа.
                page = page_number(start)
                return {"start": page, "end": page,
                        "range_start": start, "range_end": end, "fragments": []}
            start_page = page_number(start)
            end_page = page_number(end - 1)
            fragments = []
            for page_no in range(start_page, end_page + 1):
                left = max(start, page_starts[page_no - 1])
                right = min(end, page_starts[page_no])
                if right <= left:
                    continue
                fragments.append({
                    "page": page_no,
                    "start": left - start,
                    "end": right - start,
                })
            return {"start": start_page, "end": end_page,
                    "range_start": start, "range_end": end,
                    "fragments": fragments}

        paragraph_pages = {}
        tables = []
        for bookmark in doc.Bookmarks:
            name = str(bookmark.Name)
            if name.startswith("MIMO_P"):
                try:
                    paragraph_pages[int(name[6:])] = range_pages(bookmark.Range)
                except ValueError:
                    continue
            elif name.startswith("MIMO_T"):
                try:
                    table_index, row_index = map(int, name[6:].split("R", 1))
                except ValueError:
                    continue
                while len(tables) <= table_index:
                    tables.append([])
                rows = tables[table_index]
                while len(rows) <= row_index:
                    rows.append({"start": None, "end": None})
                rows[row_index] = range_pages(bookmark.Range)
        paragraph_count = max(paragraph_pages, default=-1) + 1
        paragraphs = [paragraph_pages.get(i, {"start": None, "end": None})
                      for i in range(paragraph_count)]
        return {
            "page_count": page_count,
            "paragraphs": paragraphs,
            "tables": tables,
        }
    finally:
        if doc is not None:
            try:
                doc.Close(0)
            except Exception:
                pass
