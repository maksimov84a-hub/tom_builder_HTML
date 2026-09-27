---
feature: export-sync
status: delivered
updated: 2026-02-14
branch: feature/export-sync
commits: 38eb2a9..a2a0a4b
---

# Синхронизация экспорта тома со структурой (этап A)

## Report

**What was built** — Сборка DOCX/PDF теперь идёт строго по `project.structure` (порядок + `needed`): `build_text_docx` обходит текстовые разделы по дереву и включает лист регистрации изменений при непустом журнале; `build_full_pdf` группирует structure в сегменты (text / graphic / appendix) и склеивает PDF в порядке дерева. Перед сборкой в UI выполняется preflight (`app/core/preflight.py`): ошибки блокируют экспорт, предупреждения — диалогом «Продолжить/Отмена»; проверка файлов учитывает `structure.needed`. Диагностика: падение PDF показывает причину и путь к `tmp/build.log`, ошибка наложения штампов пробрасывается (не «Готово» без рамок), `merge_pdfs` возвращает фактический `page_count` и сообщает путь битого входа.

**Verification** — `python tmp/check_preflight.py` PASS; `python tmp/check_t2_docx.py` PASS (порядок/needed/subset/changes); `python tmp/check_t3_pdf.py` PASS (сегменты, перестановка, pages=7); `python tmp/check_t4_merge.py` PASS (page_count=5, ошибка с путём); stamp-propagation one-liner PASS; `py_compile` 5 файлов PASS; structure-gate preflight one-liner PASS. Независимое ревью (general-1): критичных нет; major (preflight без structure-gate) исправлен и перепроверен; minor с дублем `TEXT_SECTION_KINDS` устранён. `tests/` отсутствует — PRE-EXISTING (этап B).

**Journey log**
1. Работа в основном чекауте на `feature/export-sync` — `git worktree add` заблокирован песочницей (общий ref store).
2. Перегрузка tool-calls: дубли запросов отменялись системой; полезно держать партии ≤3 вызовов.
3. Кириллица в fallback-PDF не извлекается как текст (шрифт) — smoke-тесты проверяют ASCII-содержимое (цифры/даты).
4. Два текстовых сегмента затирали один временный файл — добавлен уникальный `tag` (`_text_part_seg0.pdf`, `_text_part_seg3.pdf`).
5. Ревью поймало structure-gate в preflight: файлы отключённых разделов не должны блокировать сборку.

## [S1] Problem

Сборка DOCX/PDF расходится с HTML-предпросмотром и содержанием (`toc.py`):

1. `build_text_docx` (`app/core/docx_exporter.py`) безусловно добавляет обложку, титул, содержание и ПЗ; не читает `project.structure` / `needed`, не добавляет лист регистрации изменений, не учитывает порядок разделов.
2. `build_full_pdf` (`app/core/pdf_builder.py`) собирает том фиксированным порядком «текст → графика → приложения»; лист регистрации изменений не вставляется; порядок и флаги `project.structure` игнорируются; исчезнувшие входные файлы молча выпадают из тома.
3. Ошибки PDF-этапа (`build_full_pdf`) проглатываются в `main_window._build` без причины; падение наложения штампов (`_stamp_text_pages`) скрывается `except: pass` — PDF уходит без рамок с сообщением «Готово».

В результате «Собрать том» выдаёт том, состав и порядок которого не совпадают с предпросмотром и содержанием — нарушение принципа SPEC §6.

## [S2] Design

Единый источник порядка и состава блоков — `project.structure` (список `{kind, needed}`) и `toc.build_layout`. Оба экспортёра обязаны идти строго по нему.

**Порядок блоков.** Для DOCX-тома и PDF-тома последовательность разделов = `project.structure`, отфильтрованная по `needed=True`. Внутри документного раздела — `project.documents_for(kind)` (уже фильтрует `needed`); внутри приложений — `appendices_needed()`.

**Состав текстового DOCX (`build_text_docx`).** Последовательно, для каждого `kind` из отфильтрованной структуры:

| kind | Действие |
|------|----------|
| `cover` | `_add_cover` (или кастомный шаблон — вне объёма этапа A, см. Out of Scope) |
| `title` | `_add_title_sheet` |
| `content` | `_add_content` + колонтитулы форма 5 |
| `pz` | оглавление ПЗ + `documents_for(POYASNITELNAYA)` + колонтитулы 5/6 |
| `graphic` | в текстовом DOCX раздел графики не включается (как и сейчас — DOCX без графики) |
| `appendix` | в текстовый DOCX приложения не включаются (SPEC §4.8: DOC — без приложений) |
| `changes` | `_add_changes_sheet`, если `project.changes` не пуст; иначе раздел пропускается |

Отключённый раздел не создаёт вообще. Разделы, не перечисленные в structure, в том не участвуют (страховка: известные kinds берутся из structure; отсутствие kind в structure = не включать).

**Состав полного PDF (`build_full_pdf`).** PDF строится сегментами в порядке `project.structure`:

1. Отфильтровать structure по `needed=True`.
2. Сгруппировать соседние текстовые kinds (`cover`, `title`, `content`, `pz`, `changes`) в один **текстовый сегмент**; `graphic` и `appendix` — каждый свой сегмент.

Пример (структура по умолчанию): `text(cover,title,content,pz)` → `graphic` → `appendix` → `text(changes)`.

3. Каждый текстовый сегмент → `build_text_docx(project, out, kinds_subset)` (новый необязательный параметр `kinds: list[str] | None = None`) → Word COM или PyMuPDF-fallback → PDF-сегмент.
4. Сегмент `graphic` — документы `documents_for(GRAPHIC)` с квадратом номера (как сейчас).
5. Сегмент `appendix` — PDF `appendices_needed()` (как сейчас).
6. Склеить сегменты через `pymupdf.insert_pdf` строго в порядке structure; сквозная нумерация листов считается по `toc.build_layout` до сборки (как сейчас) и прокидывается в квадраты графики.

**Изменения (`changes`).** Если `structure_needed("changes")` и журнал непуст — лист изменений входит в текстовый сегмент в той позиции, где он стоит в structure. Если журнал пуст — блок пропускается (warning preflight).

**preflight (минимальный).** Перед запуском экспорта в `main_window._build` — список ошибок (блокируют) и предупреждений (не блокируют):

- ошибки: отсутствует/нечитаем файл включённого DOCX/PDF-документа или приложения; structure пуст; включён `changes`, но журнал пуст (это warning, не error);
- предупреждения: реквизиты тома пустые (наименование/обозначение); Word COM недоступен (fallback).

Новый модуль `app/core/preflight.py`: `run_preflight(project) -> PreflightResult(errors: list[str], warnings: list[str])`. При наличии errors — `QMessageBox` со списком, экспорт не запускать. Warnings — один диалог с продолжением/отменой.

**Диагностика PDF.** В `main_window._build` убрать `except Exception pe: full_pdf = None`; вместо этого логировать (`logging` в stderr/файл `tmp/build.log`) и показывать пользователю текст исключения. В `pdf_builder` заменить `except Exception: pass` вокруг `_stamp_text_pages` на лог + проброс (PDF без штампов не считается успехом).

**Мелкие фиксы в рамках A:**
- `merge_pdfs` возвращает фактическое число страниц (не 0) и не глушит ошибку открытия входного файла (пробрасывает с путём);
- `count_pages` — логировать/собирать ошибки в список вместо голого `pass` (минимум: возвращать кортеж или писать в logger).

**Не входит в A:** кастомные обложка/титул, подписанты/подписи в экспорте, форма листа изменений строго по ГОСТ, прогресс/отмена, тесты (этап B), `gost.py` vs `geometry.py` (этап C).

## [S3] Out of Scope

- Подписи (`signature_image`) в DOCX/PDF — этап C.
- Форма листа регистрации изменений по приложению ГОСТ (колонки, подписи, пагинация) — этап C; в A лист использует существующий `_add_changes_sheet`.
- Кастомные шаблоны обложки/титула (`use_custom/template_path`) — отдельный фича-документ.
- Прогресс, отмена, фоновые потоки.
- Автотесты и pytest-инфраструктура (этап B).
- Редактирование `vendor/`, `tmp/` артефактов.
- Изменение публичного поведения HTML-предпросмотра и `build_volume_html`.

## Tasks

- [x] T1: `app/core/preflight.py` — `run_preflight(project)` с errors/warnings по контракту S2; acceptance: отсутствующий путь документа → error; включён `changes` при пустом журнале → warning; пустые реквизиты → warning (covers: S2)
- [x] T2: `build_text_docx` — сборка строго по `project.structure` (порядок + needed), пропуск `changes` при пустом журнале; acceptance: структура без `content` → в DOCX нет листа содержания; structure с changes и записями → лист изменений присутствует; порядок разделов в DOCX = порядку structure (covers: S2)
- [x] T3: `build_full_pdf` — сегментация structure, склейка PDF в порядке structure; `build_text_docx` получает параметр `kinds: list[str] | None`; acceptance: по умолчанию PDF = text(cover,title,content,pz) → graphic → appendix → changes; отключённый раздел отсутствует; перестановка graphic/appendix в structure меняет порядок в PDF; `changes` в середине structure попадает в свою позицию (covers: S2; depends: T1, T2)
- [x] T4: диагностика: `main_window._build` показывает причину падения PDF и пишет `tmp/build.log`; `pdf_builder` не глушит ошибку наложения штампов; `merge_pdfs` возвращает фактический page_count; acceptance: падение `_stamp_text_pages` → пользователь видит текст ошибки, не «Готово»; `merge_pdfs([a], out)` == page_count файла a (covers: S2; depends: T2)