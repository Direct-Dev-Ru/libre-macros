# Ручное тестирование collect_workbooks

**Версия макроса:** `MACRO_VERSION` в `macro-lib/collect_workbooks.py` и `macro-lib/version.txt` (актуально: **3.10.523**).

Автоматические pytest-скрипты в проекте **не используются**. Проверка — вручную в LibreOffice Calc по шаблонам сценариев.

## Содержание

1. [Что проверяем](#что-проверяем)
2. [Требования](#требования)
3. [Структура каталогов](#структура-каталогов)
4. [Генерация шаблонов](#генерация-шаблонов)
5. [Прогон в Calc](#прогон-в-calc)
6. [Сценарии](#сценарии)
7. [Матрица постобработки JSON](#матрица-постобработки-json)
8. [Типичные проблемы](#типичные-проблемы)

---

## Что проверяем

Макрос **`collect_workbooks`** читает лист **`Параметры_Объединения*`**, копирует данные из источников и выполняет постобработку.

Параметры постобработки в колонке **C** — **только JSON** ([13_JSON_PARAMS.md](13_JSON_PARAMS.md)).

Проверяем:

- режимы «На один лист» / «На разные листы» / **«Копирование листов»**;
- **`Способ_переноса` = `xml_excel_ods`** (direct + отложенная UNO-фаза);
- фильтры листов и столбцов;
- «Ячейка_В_Столбец»;
- цепочки `Постобработка_Диапазон` / `Постобработка_Строка` / **`Постобработка_xml`** с JSON;
- **`MERGE_XML_CONVERT_TO_NUMBERS`** — текстовые **числа** (`1 000 000,00`, `'1 000 000.00`) и **даты** (copy + апостроф `'02.02.1976`);
- журнал **«Сбор_книг_лог»** (ошибки, пропуски, версия макроса);
- при xml: **Постобработка_Диапазон/Строка после показа** книги, прогресс в **строке состояния**.

---

## Требования

| Компонент | Зачем |
|-----------|--------|
| Python 3.10+ | генератор шаблонов |
| **openpyxl** | `pip install openpyxl` |
| LibreOffice Calc | прогон макроса |
| Установленные макросы | `./macro-lib/install.sh collect_workbooks param_wizard` |

---

## Структура каталогов

```text
test/collect_workbooks/
├── generate_manual_workbooks.py   ← генератор (источники + .xltx)
├── manual_json_presets.py         ← примеры JSON для колонки C
├── generate.sh                    ← обёртка
├── sources/                       ← source_01..03, source_products
├── start_param_file.xlsx          ← основа шаблонов (кнопка «Обзор»)
└── workbooks/
    ├── MANUAL_CHECK.md            ← чек-лист сценариев (генерируется)
    ├── 01_merge_one_sheet.xltx    ← в каждом .xltx есть лист «Справка_макроса»
    ├── …
    ├── 12_postprocess_json.xltx ← матрица всех JSON-функций
    ├── ODF/*.ots                  ← опционально (--odf)
    └── results/                   ← сюда сохранять результат после макроса
```

---

## Генерация шаблонов

```bash
cd /path/to/libre-macros
python3 -m venv venv && source venv/bin/activate
pip install openpyxl

./test/collect_workbooks/generate.sh
# или с конвертацией в ODF для Calc:
python test/collect_workbooks/generate_manual_workbooks.py --odf
```

Опции:

- `--skip-sources` — не пересоздавать `sources/*.xlsx` (если уже есть);
- `--odf` — дополнительно `workbooks/ODF/*.ots` (нужен `soffice` в PATH).

После генерации откройте **`workbooks/MANUAL_CHECK.md`** — краткий чек-лист.

В каждый шаблон добавляется лист **`Справка_макроса`** (второй после параметров): docstring макроса, оглавление `docs/`, каталог функций постобработки с подсказками JSON. Версия в строке «Версия» и в заголовке справки берётся из `libre_macros_lib.py`.

---

## Прогон в Calc

1. Установите макросы ([03_INSTALL_DEV.md](03_INSTALL_DEV.md)).
2. Откройте `workbooks/ODF/01_merge_one_sheet.ots` (или `.xltx` в Excel).
3. При необходимости поправьте пути в **Файлы-Источники** на абсолютные к `sources/`.
4. **Инструменты → Макросы → Python → collect_workbooks**.
5. Сверьте результат с комментарием сценария и журналом «Сбор_книг_лог».
6. Сохраните копию в `workbooks/results/` для сравнения при следующих изменениях.

---

## Headless-прогон (ACell / soffice)

Автоматический прогон **манипуляций данными** без визарда. Код и **подробные доки**: [`qa/collect_headless/README.md`](../qa/collect_headless/README.md) (оглавление QA: [`qa/README.md`](../qa/README.md)). ТЗ: `todo/модуль_headless_тестирования_сбора.md`. Справочник проверок: [`qa/collect_headless/CHECKS.md`](../qa/collect_headless/CHECKS.md).

Нужны установленный ACell или LibreOffice, шаблоны `workbooks/ODF/*.ots` (при необходимости `./test/collect_workbooks/generate.sh` + `--odf`), Python 3 с `openpyxl` для `.xlsx`. Если нет `DISPLAY` — опционально `Xvfb`.

```bash
# манифест по умолчанию (01, 02, 03, 11_direct, 11_vlookup_1to1_low, 12)
python3 qa/collect_headless/run.py

# один сценарий
python3 qa/collect_headless/run.py --only 01_merge_one_sheet

# свой каталог шаблонов
python3 qa/collect_headless/run.py --templates-dir test/collect_workbooks/workbooks/ODF \
  --results-dir /tmp/lm-qa --glob '01_*.ots'
```

Результаты и `report.md`: `test/collect_workbooks/workbooks/results/headless/` (в `.gitignore`).  
Макросы копируются в изолированный профиль `qa/collect_headless/.profile` (не трогает GUI-сессию). Офис: `LM_QA_OFFICE`, порт UNO: `LM_QA_PORT`.

Визард, диалоги «Запуск»/«Завершить», градиент/сетка как UI — **не проверяются**. Откройте сохранённый `.ods` в обычном ACell, если нужна глазами вёрстка.

---

## Сценарии

| Файл | Назначение |
|------|------------|
| `00_all_parameters.xltx` | Справочник параметров сбора |
| `01_merge_one_sheet.xltx` | На один лист + базовая JSON-постобработка |
| `02_merge_many_sheets.xltx` | На разные листы |
| `03_columns_filter.xltx` | Фильтр столбцов |
| `04_sheets_wildcard.xltx` | Шаблоны листов `*2024`, `Отчет*` |
| `05_all_columns.xltx` | Столбцы = «Все» |
| `06_sheets_by_index.xltx` | Лист по номеру |
| `07_columns_by_index.xltx` | Столбцы по номерам |
| `08_columns_by_letters.xltx` | Столбцы A, C, F |
| `09_fixed_row_count.xltx` | Лимит строк |
| `10_product_groups.xltx` | Каталог товаров |
| `11_vlookup_1to1_low.xltx` | ВПР 1:1, low (~180 строк, ключ Код_позиции) |
| `11_vlookup_1to1_middle.xltx` | ВПР 1:1, middle (составной ключ 2 поля) |
| `11_vlookup_1to1_high.xltx` | ВПР 1:1, high (ключ 3 поля, ~3000 строк) |
| `11_vlookup_1toN_low.xltx` | ВПР 1:N, low — дубли ключей справа |
| `11_vlookup_1toN_middle.xltx` | ВПР 1:N, middle |
| `11_vlookup_1toN_high.xltx` | ВПР 1:N, high |
| **`11_direct_xml_excel_ods.xltx`** | **xml_excel_ods**: direct-перенос, источники .xlsm/.xls/.ods |
| **`12_postprocess_json.xltx`** | **Матрица всех основных JSON-функций** |
| **`13_merge_sheets_pivot.xltx`** | **На разные листы → `объединить_листы_в_один` (A) → сводная** |

---

## Матрица постобработки JSON

Сценарий **`12_postprocess_json.xltx`** — один прогон для визуальной проверки RANGE/ROW функций с JSON в C.

После макроса пройдите по списку в `workbooks/MANUAL_CHECK.md` и убедитесь:

- нет записей «ошибка» / неожиданных «пропуск» в логе;
- каждая функция дала ожидаемый визуальный эффект на листе «Отчет_январь».

Полные примеры JSON: [13_JSON_PARAMS.md](13_JSON_PARAMS.md).

---

## Сценарий объединить_листы + сводная (`13_merge_sheets_pivot.xltx`)

1. Источник `sources/source_pivot.xlsx` (листы `Отчет_Q1`…`Q3`).
2. На листе параметров: **Режим** = «На разные листы»; строка A **`объединить_листы_в_один`** с JSON в C (не B в «Постобработка_Диапазон»).
3. Ожидание: `Отчет_Q123`, сводная `Сводная_Отдел` (Отдел × Квартал).
4. Аналог пресета визарда **Я05 / Я14**.

---

## Сценарий xml_excel_ods (`11_direct_xml_excel_ods.xltx`)

1. Книга результата **сохранена** (`.ods` или `.xlsx`).
2. Запустите `collect_workbooks`; во время direct-фазы книга может быть закрыта — это нормально.
3. После сбора окно книги **появляется**, затем идёт UNO-оформление и `Постобработка_Диапазон` / `Строка` (статус — строка состояния).
4. Сверьте версию в «Сбор_книг_лог» (**≥ 3.10.523** рекомендуется; **≥ 3.10.81** — числа с разрядами при copy; **≥ 3.10.79** — отложенная UNO-фаза).
5. Многоэтажные заголовки: источники `test/collect_workbooks/sources/multilevel_headers/`, параметр `Строка_Заголовков` + JSON в C ([04](04_DATA_COLLECTION.md) §6.0.2).

Дополнительно для **Копирование листов + xml**: текстовые числа (`1 000 000,00`) и даты с апострофом, `MERGE_XML_CONVERT_TO_NUMBERS=Да` — см. [17_POSTPROCESS_XML.md](17_POSTPROCESS_XML.md).

---

## Типичные проблемы

| Симптом | Решение |
|---------|---------|
| Макрос не виден | Переустановка, `libreoffice-script-provider-python` |
| Пустой результат | Проверьте пути в «Файлы-Источники», активный лист параметров |
| Постобработка пропущена | JSON в C невалиден или нет блока для текущего листа (`sheet`) |
| Старая версия в логе | Переустановите макросы, сверьте строку «Версия» на листе |
| Градиент / сортировка не сработали | Убедитесь, что в логе версия ≥ 3.10.x и C содержит JSON, не legacy-текст |
| xml: постобработка «до показа» / freeze не работает | При xml_excel_ods RANGE/ROW и freeze — **после** показа книги; см. [17_POSTPROCESS_XML.md](17_POSTPROCESS_XML.md) |
| xml: «книга не сохранена» | Сохраните результат на диск до запуска |
| Числа остались текстом (copy + xml) | `MERGE_XML_CONVERT_TO_NUMBERS=Да`; в столбце ≥3 распознанных значения; формат `1 000 000,00` / `1 000 000.00` |
