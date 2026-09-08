# Быстрый старт

Полное оглавление: **[00_INDEX.md](00_INDEX.md)**.  
**Версия макросов:** **3.10.711** (`macro-lib/version.txt`).  
Безопасность (гейты, доверенные корни): [19_SECURITY.md](19_SECURITY.md).

## Установка (Linux / Windows)

Полное руководство: [03_INSTALL_DEV.md](03_INSTALL_DEV.md). Пользовательский установщик: [02_INSTALL_USER.md](02_INSTALL_USER.md).

```bash
# Linux
sudo apt install libreoffice-script-provider-python   # при необходимости
./macro-lib/install.sh
```

На Windows: `macro-lib/install.ps1` или копирование `.py` из `macro-lib/` в  
`%APPDATA%\LibreOffice\4\user\Scripts\python\` (и `pythonpath/` рядом). Перед установкой `install.sh`/`install.ps1` гоняют flatten + AO compat check.

---

## Сбор данных: `collect_workbooks` и `collect_pack`

1. Откройте в Calc книгу с листом `Параметры_Объединения` (или `Параметры_Объединения_1`, …).

2. Заполните параметры:
   - вручную в колонках B, C, D…;
   - **pick_source_file** — файл(ы) в строку «Файлы-Источники»;
   - **param_wizard** — любой параметр, пресеты ([06_PARAM_WIZARD.md](06_PARAM_WIZARD.md)).

3. Запуск:
   - **`collect_workbooks`** — один конфиг (при нескольких листах группы активируйте нужный).
   - **`collect_pack`** — оркестратор: выбрать несколько листов параметров, порядок, опционально сброс DatabaseRange ([18_ORCHESTRATOR.md](18_ORCHESTRATOR.md)).

4. Режим «Диалог» (`Отображение_прогресса`) — сводка и **комментарий** перед сбором.

### Режимы (параметр «Режим»)

| Режим | Кратко |
|-------|--------|
| На один лист / На разные листы / Копирование листов | Сбор из файлов (или `эта_книга`) |
| **Вовлечь листы** | Без копирования: вкладки текущей книги сразу в постобработку/финал |

Подробно: [04_DATA_COLLECTION.md](04_DATA_COLLECTION.md).

Постобработка: колонка **C** — **только JSON** ([13_JSON_PARAMS.md](13_JSON_PARAMS.md)).

При **`Способ_переноса` = `xml_excel_ods`** книга должна быть **сохранена** на диск; см. [17_POSTPROCESS_XML.md](17_POSTPROCESS_XML.md).

### Ручной прогон

Подробно: [12_MANUAL_TESTING.md](12_MANUAL_TESTING.md).

```bash
pip install openpyxl
./test/collect_workbooks/generate.sh --skip-sources
# открыть workbooks/… в Calc → collect_workbooks / collect_pack
```

---

## Другие макросы

| Макрос | Назначение |
|--------|------------|
| `param_wizard` (`set_merge_param`) | Визард листа параметров |
| `pick_source_file` | FilePicker в «Файлы-Источники» |
| `file_list_macro` | Список/анализ файлов |
| `generate_qr_codes` | QR |
| `convert_tables_to_ranges` | Таблицы → диапазоны |
| `set_menu` / `setContextMenu` | Меню / контекст |
| `show_version` | Версия пакета |

```bash
./macro-lib/install.sh file_list_macro
```
