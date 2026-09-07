# Документация libre-macros

Единый каталог документов проекта. Номера задают **рекомендуемый порядок** знакомства: от установки и первого прогона к архитектуре и ручной проверке.

## Быстрый старт

| № | Документ | Для кого | Содержание |
|---|----------|----------|------------|
| [01](01_QUICKSTART.md) | QUICKSTART | все | Установка, первый запуск, ссылки |
| [02](02_INSTALL_USER.md) | Установка (пользователь) | пользователь | `MacroInstaller_bundled.ods` |
| [03](03_INSTALL_DEV.md) | Установка (разработчик) | разработчик | `install.sh`, пути LO, AO 2026, отладка |

## Сбор данных и параметры

| № | Документ | Содержание |
|---|----------|------------|
| [15](15_PARAM_REFERENCE.md) | **Справочник параметров** | Все имена колонки A, значения B/C, ключи постобработки |
| [04](04_DATA_COLLECTION.md) | Сбор книг | Лист `Параметры_Объединения*`, режимы (в т.ч. «Вовлечь листы»), фильтры, пути |
| [18](18_ORCHESTRATOR.md) | **Оркестратор `collect_pack`** | Несколько листов параметров, порядок, журналы `_N`, сброс DatabaseRange |
| [05](05_POSTPROCESS.md) | Постобработка | `merge_pp_*`, RANGE/ROW, pipeline, финал |
| [17](17_POSTPROCESS_XML.md) | **Постобработка_xml** | Двухфазный пайплайн `xml_excel_ods`, direct/UNO |
| [06](06_PARAM_WIZARD.md) | Визард параметров | `set_merge_param`, пресеты, диалоги |
| [07](07_USER_FUNCTIONS.md) | Пользовательские функции | `functions_pp.py`, `functions_final.py` |
| [08](08_PICK_SOURCE.md) | Выбор файла | `pick_source_file`, FilePicker |

## Архитектура и код

| № | Документ | Содержание |
|---|----------|------------|
| [09](09_ARCHITECTURE.md) | Архитектура | Поток `collect_workbooks` / `collect_pack`, модули `pythonpath/` |
| [10](10_MERGE_ALGORITHM.md) | Алгоритм «На один лист» | Слияние листов, индексы |
| [11](11_LIB_REFERENCE.md) | Справочник библиотеки | `libre_macros_lib.py`, `merge_pp_*` |
| [16](16_CODE_SANITIZE.md) | Санация Python | AST / file# перед eval |

## Проверка и формат JSON

| № | Документ | Содержание |
|---|----------|------------|
| [12](12_MANUAL_TESTING.md) | Ручное тестирование | Шаблоны в `test/collect_workbooks/`, прогон в Calc |
| [13](13_JSON_PARAMS.md) | **JSON в колонке C** | Поля блоков по каждой функции кодека |
| [14](14_TZ_UNIFICATION.md) | ТЗ унификации | История и полная спецификация (архив) |

## Каскадная (водопадная) проектная документация

Формальный комплект по состоянию кода: [technical_documentation/00_INDEX.md](technical_documentation/00_INDEX.md) — технические требования, ТЗ, пояснительная записка, методика испытаний (версия макросов **3.10.523**).

## Связанные каталоги

```text
docs/                          ← эта папка (вся prose-документация)
docs/docx/                     ← Word (.docx), сборка: python3 docs/build_docx.py --all
docs/technical_documentation/  ← ТТ, ТЗ, ПЗ, методика испытаний (водопад)
macro-lib/                     ← исходники макросов (.py); при установке → Scripts/python/
macro-lib/pythonpath/          ← библиотеки (обычный import; фильтр AO 2026 не применяется)
test/collect_workbooks/        ← шаблоны сценариев, источники, результаты ручных прогонов
todo/                          ← черновики ТЗ на новые функции
.cursor/rules/                 ← правила для ассистента (JSON-only, AO 2026, без автотестов)
```

## Версия макроса

Каноническая версия — `macro-lib/version.txt` и `MACRO_VERSION` в `macro-lib/collect_workbooks.py` (на момент обновления документации: **3.10.523**). Синхронизируется по пакету (`sync_macro_version.py` / сборка bundled).

После обновления кода переустановите макросы в LibreOffice / AlterOffice и сверьте строку **Версия** на листе параметров и запись в «Сбор_книг_лог».

## Ключевые возможности (состояние 3.10.523)

| Тема | Документ |
|------|----------|
| Четыре режима сбора, в т.ч. **«Вовлечь листы»** | [04](04_DATA_COLLECTION.md) |
| Оркестратор **`collect_pack`** | [18](18_ORCHESTRATOR.md) |
| Маркер текущей книги (`эта_книга` / `ThisWorkbook`, гибкое написание) | [04](04_DATA_COLLECTION.md) §5.3, [15](15_PARAM_REFERENCE.md) |
| Pipeline: range ↔ ВПР / merge / split ↔ row; JSON-only в C | [05](05_POSTPROCESS.md), [13](13_JSON_PARAMS.md) |
| `разделить_листы` (полный split, `run_phase`) | [13](13_JSON_PARAMS.md), [15](15_PARAM_REFERENCE.md) |
| `xml_excel_ods` + `Постобработка_xml` | [17](17_POSTPROCESS_XML.md) |
| AO 2026: flatten + compat check для entry-скриптов | [03](03_INSTALL_DEV.md) |
| **Многоэтажные заголовки** (`Строка_Заголовков`, JSON в C) | [04](04_DATA_COLLECTION.md) §6, [15](15_PARAM_REFERENCE.md), [13](13_JSON_PARAMS.md) |
| **Разделить по столбцам**, **условный столбец**, **столбец по лямбде** | [05](05_POSTPROCESS.md), [13](13_JSON_PARAMS.md), [15](15_PARAM_REFERENCE.md) |
| Word-копии документации (`.docx`) | `python3 docs/build_docx.py --all` → [03](03_INSTALL_DEV.md) |
