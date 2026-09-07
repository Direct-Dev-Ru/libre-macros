# Как добавить функцию в libre_macros_lib и подключить в collect_workbooks

Инструкция для текущей схемы проекта (один макрос + общая библиотека постобработки).

| Файл | Роль |
|------|------|
| `macro-lib/collect_workbooks.py` | макрос: сбор данных, оформление, оркестрация постобработки |
| `macro-lib/pythonpath/libre_macros_lib.py` | функции постобработки (`lm_pp_*`) и хелперы |
| `macro-lib/param_wizard.py` | визард листа «Параметры_Объединения» (в т.ч. выпадающие списки B/C) |

Постобработку **правим в lib** (`lm_pp_*`). В `collect_workbooks.py` — регистрация имён для листа параметров, карты `_PP_*_MAP_SPEC`, подсказки колонки C и вызов `_ensure_pp_from_lib()`.

Отдельного «full/lite» и генератора макроса **нет**: единственный источник — `collect_workbooks.py`.

---

## Именование

| Где | Префикс | Пример |
|-----|---------|--------|
| Библиотека, публичный колбэк | `lm_pp_range_*` / `lm_pp_row_*` | `lm_pp_range_fill_down_empty` |
| Библиотека, внутренний хелпер | `_lm_pp_*` | `_lm_pp_parse_column_list` |
| Библиотека, константа | `_LM_PP_*` | `_LM_PP_DEFAULT_FONT_SIZE` |
| В макросе после загрузки lib (shim) | `merge_pp_*` / `_merge_pp_*` / `_merge_*` | `merge_pp_range_fill_down_empty` |
| Ключ в колонке **B** листа параметров | русский идентификатор | `заполнение_вниз` |

Сравнение заголовков столбцов и имён листов при **сборе**: `merge_identity_key()` в макросе / `lm_identity_key()` в lib (trim + `casefold`).

В **постобработке** при поиске столбца по заголовку предпочтительно `lm_identity_key`, не прямое `==`.

---

## Сигнатуры колбэков

**Диапазон (RANGE)** — один проход по листу:

```python
def lm_pp_range_my_feature(doc, sheet, data_range, header_row_range, *extra_args):
    ...
```

**Строка (ROW)** — для каждой строки данных:

```python
def lm_pp_row_my_feature(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    ...
```

`extra_args` — фрагменты из колонки **C** (разбивка через `parse_pp_extra_args`: запятая с учётом кавычек/скобок). Если в C один текст «`ФИО;Отдел;3`», он часто приходит одним элементом кортежа — объединяйте: `", ".join(str(x) for x in extra_args)`.

Полезно внутри колбэка:

- `lm_pp_active_context()` — границы листа после предыдущих шагов цепочки;
- `cell_text`, `lo_color_rgb`, `_lm_pp_range_address`, `_lm_pp_header_titles`;
- `_lm_pp_parse_columns_param` — номер (1-based), буква (`A`) или подстрока в заголовке;
- журнал: `_lm_log_postprocess` (хост вешает hook через `lm_pp_set_log_hook` → `_merge_log_postprocess` в макросе).

---

## Шаг 1. Код в `libre_macros_lib.py`

1. Добавьте функцию `lm_pp_range_*` или `lm_pp_row_*`.
2. При необходимости — хелперы `_lm_pp_*` (только в lib).
3. Добавьте имя в кортеж **`LM_PP_PUBLIC_NAMES`** в конце файла.

Без записи в `LM_PP_PUBLIC_NAMES` функция **не попадёт** в `macros_lib` и карты MAP.

Проверка синтаксиса:

```bash
python3 -m py_compile macro-lib/pythonpath/libre_macros_lib.py
```

---

## Шаг 2. Подключение в `collect_workbooks.py`

Блок **«Постобработка из pythonpath/libre_macros_lib.py»** (≈ строки 10600+).

### 2.1. Карта имён для колонки B

`_PP_RANGE_MAP_SPEC` / `_PP_ROW_MAP_SPEC`:

```python
_PP_RANGE_MAP_SPEC = (
    ...
    ("заполнение_вниз", "merge_pp_range_fill_down_empty"),
)
```

- Первый элемент — ключ в **B** (`Постобработка_Диапазон`).
- Второй — имя в `macros_lib`: всегда `merge_pp_*` (не `lm_pp_*`).

При запуске `_fill_postprocess_maps()` заполняются `MERGE_RESULT_POSTPROCESS_RANGE_MAP` / `MERGE_RESULT_POSTPROCESS_ROW_MAP`.

### 2.2. Подсказка для колонки C

`MERGE_POSTPROCESS_RANGE_REST_HINTS` / `MERGE_POSTPROCESS_ROW_REST_HINTS`:

```python
MERGE_POSTPROCESS_RANGE_REST_HINTS = {
    ...
    "заполнение_вниз": (
        "столбцы: номер (1-based), буква A или имя заголовка; через , или ;",
        "ФИО;Отдел;3",
    ),
}
```

Используется в справке, `param_wizard` и генераторах тестовых книг.

### 2.3. (Опционально) Жёсткая цепочка в коде

```python
_MERGE_CODE_POSTPROCESS_RANGE = (
    # merge_pp_range_fill_down_empty,
)
```

Обычно достаточно строк на листе параметров.

### 2.4. (Опционально) Особый разбор колонки C

По умолчанию C режется через `parse_pp_extra_args`. Свой разбор — в `parse_postprocess_from_sheet` в `collect_workbooks.py`:

```python
elif str(fn_name or "").strip().lower() == "заполнение_вниз":
    extra_args = _merge_pp_parse_column_list_extra(extra_raw)
```

Парсер можно держать в lib (`_lm_pp_parse_column_list_extra`) — после `_install_pp_shims()` в макросе он доступен как `_merge_pp_parse_column_list_extra`.

### 2.5. Локальный колбэк только в макросе

Region **«Свои колбэки постобработки»** в `collect_workbooks.py` — `merge_pp_*` без lib. В `_PP_*_MAP_SPEC` укажите локальное имя; в `LM_PP_PUBLIC_NAMES` не добавляйте.

### 2.6. Визард параметров (`param_wizard.py`) — выпадающие списки

Списки в колонке **B** для строк «Постобработка_Диапазон» / «Постобработка_Строка» собираются в `_build_wizard_catalog()`.

**Основной путь (рекомендуется):** ключи из `collect_workbooks.py` попадают в список автоматически:

```
_build_wizard_catalog()
    → _try_import_collect()
    → _ensure_collect_pp_maps(cw)     # cw._ensure_pp_from_lib()
    → sorted(cw.MERGE_RESULT_POSTPROCESS_RANGE_MAP.keys())
    → sorted(cw.MERGE_RESULT_POSTPROCESS_ROW_MAP.keys())
```

После правок в `libre_macros_lib.py` и `collect_workbooks.py` достаточно `./install.sh collect_workbooks param_wizard` и **перезапуска Calc** — в B появится новый ключ (например `заполнение_вниз`).

**Запасные списки** (если `collect_workbooks` не загрузился или карта пуста) — править вручную в `param_wizard.py`:

| Место | Назначение |
|-------|------------|
| `_FALLBACK_PP_RANGE_CHOICES` | ключи диапазонной постобработки для B |
| `_FALLBACK_PP_ROW_CHOICES` | ключи постобработки строки для B |
| `_FALLBACK_PP_RANGE_REST_HINTS` | подсказки колонки C (диапазон) |
| `_FALLBACK_PP_ROW_REST_HINTS` | подсказки колонки C (строка) |

Подсказки из макроса подмешиваются через `MERGE_POSTPROCESS_*_REST_HINTS` (если collect загружен).

**Функции из `functions_pp.py` и `functions_final.py`** в список B добавляются **автоматически**: `_scan_functions_pp_choices()` / `_scan_functions_final_choices()` ищут `def pp_*` в файлах рядом с макросом (`functions_pp.py#имя`, `functions_final.py#имя`).

**Валидация на листе** ставится в `_apply_param_sheet_validations()` / `_set_cell_list_validation()` при создании листа, пресетах и записи из визарда.

Чеклист для визарда при новой функции **в lib**:

- [ ] шаги 1–2 (lib + `collect_workbooks`) и `./install.sh collect_workbooks param_wizard`
- [ ] (опционально) `_FALLBACK_PP_RANGE_CHOICES` + `_FALLBACK_PP_RANGE_REST_HINTS` — если нужен запасной список без collect
- [ ] перезапуск Calc

Пользовательские функции в отдельном файле — см. **`macro-lib/USER_PP_FUNCTIONS_GUIDE.md`** (без правки fallback-списков, если имя `pp_*` в `functions_pp.py`).

---

## Шаг 3. Установка в LibreOffice

```bash
cd macro-lib
./install.sh collect_workbooks param_wizard
```

Копируется:

- `collect_workbooks.py` → `~/.config/libreoffice/4/user/Scripts/python/`
- `pythonpath/libre_macros_lib.py` → `.../Scripts/python/pythonpath/`
- `param_wizard.py` — по желанию

Перезапустите Calc.

Опционально без docstring (если LO ругается на кириллицу в docstring макроса):

```bash
python3 generate_no_docstrings.py
./install.sh collect_workbooks_no_docstrings
```

---

## Как это связывается при запуске

```
collect_workbooks() / collect_workbooks_run()
    → _ensure_pp_from_lib()
        → _load_macros_lib()       # LM_PP_PUBLIC_NAMES → macros_lib["merge_pp_*"]
        → _install_pp_shims()      # lm_pp_* → merge_pp_*; _lm_pp_* → _merge_pp_*;
                                   # _lm_finalize_* → _merge_finalize_* и т.д.
        → _fill_postprocess_maps() # _PP_*_MAP_SPEC → MERGE_RESULT_POSTPROCESS_*_MAP
```

Во время постобработки:

```
apply_result_sheet_postprocessing()
    → _merge_pp_sync_lib_active_context(pp_ctx)   # lm_pp_active_context() в lib
```

`param_wizard` перед чтением списков функций вызывает `_ensure_pp_from_lib()` через `_ensure_collect_pp_maps()`, чтобы в выпадающих списках B были актуальные ключи постобработки.

В Python-коде колонки **B** (`lambda` / `def`) доступны имена из `_merge_pp_postprocess_eval_namespace()` — в т.ч. `merge_pp_*` после shims.

---

## Чеклист для новой RANGE-функции

- [ ] `lm_pp_range_*` в `libre_macros_lib.py`
- [ ] имя в `LM_PP_PUBLIC_NAMES`
- [ ] `("ключ_b", "merge_pp_range_*")` в `_PP_RANGE_MAP_SPEC`
- [ ] запись в `MERGE_POSTPROCESS_RANGE_REST_HINTS`
- [ ] при особом C — ветка в `parse_postprocess_from_sheet` (опционально)
- [ ] `py_compile` lib + `./install.sh collect_workbooks param_wizard`
- [ ] (опционально) `_FALLBACK_PP_*` в `param_wizard.py` — см. п. 2.6
- [ ] строка на листе: A=`Постобработка_Диапазон`, B=`ключ_b`, C=…
- [ ] при необходимости — перегенерировать `collect_workbooks_no_docstrings.py`

## Чеклист для новой ROW-функции

То же, но `lm_pp_row_*`, `_PP_ROW_MAP_SPEC`, `MERGE_POSTPROCESS_ROW_REST_HINTS`, A=`Постобработка_Строка`.

---

## Полный пример: «заполнение вниз» по столбцам

Задача: после сбора пройти указанные столбцы сверху вниз; если ячейка **пуста**, скопировать значение из строки выше (как Fill Down в Calc, только для выбранных колонок).

**Колонка C:** список столбцов через запятую или точку с запятой:

- номер столбца (1 = A, как в параметрах макроса);
- буква (`B`, `AA`);
- имя или часть заголовка (`ФИО`, `Отдел`).

Пример C: `ФИО;Отдел;3` или `ФИО, Отдел, 3`.

Порядок шагов ниже — от кода до строки на листе параметров.

---

### Шаг A. Хелперы и функция в `libre_macros_lib.py`

Добавьте **перед** блоком `LM_PP_PUBLIC_NAMES` (рядом с другими range-функциями):

```python
import re


def _lm_pp_join_extra_args(extra_args):
    """Склеить rest из колонки C в одну строку."""
    if extra_args is None:
        return ""
    parts = []
    i = 0
    while i < len(extra_args):
        s = str(extra_args[i]).strip()
        if s != "":
            parts.append(s)
        i = i + 1
    return ", ".join(parts)


def _lm_pp_split_column_tokens(spec_text):
    """Разбить C на токены столбцов по запятой или точке с запятой."""
    text = str(spec_text or "").strip()
    if text == "":
        return []
    raw = re.split(r"[,;]", text)
    out = []
    i = 0
    while i < len(raw):
        t = str(raw[i]).strip()
        if t != "":
            out.append(t)
        i = i + 1
    return out


def _lm_pp_resolve_column_token(token, sheet, header_row_range, sc, ec):
    """
    Один токен → 0-based индекс столбца в пределах листа.
    Номер (1-based), буква A/AA или заголовок (lm_identity_key).
    """
    token = str(token).strip()
    if token == "":
        return None
    try:
        idx = int(token) - 1
        if idx >= 0:
            return idx
    except (TypeError, ValueError):
        pass
    if is_col_letters(token):
        return col_letters_to_index(token)
    titles = _lm_pp_header_titles(sheet, header_row_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    key = lm_identity_key(token)
    c = h_sc
    while c <= h_ec:
        pos = c - h_sc
        title = titles[pos] if pos < len(titles) else ""
        if lm_identity_key(title) == key or key in lm_identity_key(title):
            return c
        c = c + 1
    return None


def _lm_pp_resolve_column_list(spec_text, sheet, header_row_range, sc, ec):
    """Список уникальных 0-based индексов столбцов из C."""
    tokens = _lm_pp_split_column_tokens(spec_text)
    cols = []
    i = 0
    while i < len(tokens):
        col = _lm_pp_resolve_column_token(
            tokens[i], sheet, header_row_range, sc, ec
        )
        if col is not None and col not in cols:
            cols.append(col)
        i = i + 1
    cols.sort()
    return cols


def _lm_pp_copy_cell_value_down(src_cell, dst_cell):
    """Скопировать отображаемое значение из src в dst (для fill-down)."""
    if src_cell is None or dst_cell is None:
        return False
    if cell_text(dst_cell) != "":
        return False
    if cell_text(src_cell) == "":
        return False
    try:
        dst_cell.String = src_cell.String
        return True
    except Exception:
        try:
            dst_cell.Value = src_cell.Value
            return True
        except Exception:
            return False


def lm_pp_range_fill_down_empty(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Заполнение вниз: в указанных столбцах пустые ячейки получают значение из строки выше.

    Rest (колонка C): столбцы через , или ; — номер (1-based), буква или имя заголовка.
    Пример: «ФИО;Отдел;3».

    Карта: «заполнение_вниз».
    """
    sheet_name = sheet.Name if sheet is not None else ""
    spec_text = _lm_pp_join_extra_args(extra_args)
    if spec_text == "":
        _lm_log_postprocess(
            doc, sheet_name, "заполнение_вниз", "—", "пропуск", "пустая колонка C"
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    if er <= sr:
        _lm_log_postprocess(
            doc, sheet_name, "заполнение_вниз", spec_text, "пропуск", "нет строк данных"
        )
        return

    cols = _lm_pp_resolve_column_list(spec_text, sheet, header_row_range, sc, ec)
    if len(cols) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "заполнение_вниз",
            spec_text,
            "ошибка",
            "не найдены столбцы",
        )
        return

    filled = 0
    ci = 0
    while ci < len(cols):
        col = cols[ci]
        row = sr + 1
        while row <= er:
            if _lm_pp_copy_cell_value_down(
                sheet.getCellByPosition(col, row - 1),
                sheet.getCellByPosition(col, row),
            ):
                filled = filled + 1
            row = row + 1
        ci = ci + 1

    _lm_log_postprocess(
        doc,
        sheet_name,
        "заполнение_вниз",
        spec_text,
        "ok",
        "столбцов: %d, заполнено ячеек: %d" % (len(cols), filled),
    )
```

В кортеж **`LM_PP_PUBLIC_NAMES`** добавьте строку:

```python
    "lm_pp_range_fill_down_empty",
```

---

### Шаг B. Регистрация в `collect_workbooks.py`

В `_PP_RANGE_MAP_SPEC`:

```python
    ("заполнение_вниз", "merge_pp_range_fill_down_empty"),
```

В `MERGE_POSTPROCESS_RANGE_REST_HINTS`:

```python
    "заполнение_вниз": (
        "столбцы: номер (1-based), буква A или имя заголовка; через , или ;",
        "ФИО;Отдел;3",
    ),
```

Сохраните файл. Shims подставят `merge_pp_range_fill_down_empty` автоматически — отдельно объявлять `def merge_pp_*` в макросе не нужно.

---

### Шаг C. Установка и проверка синтаксиса

```bash
cd macro-lib
python3 -m py_compile pythonpath/libre_macros_lib.py collect_workbooks.py
./install.sh collect_workbooks param_wizard
```

Копируются **оба** файла: `collect_workbooks.py` и `pythonpath/libre_macros_lib.py`. Без lib в `pythonpath/` будет ошибка вида `'merge_pp_range_…'` при заполнении карт постобработки.

Перезапустите Calc.

---

### Шаг D. Строка на листе «Параметры_Объединения»

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | заполнение_вниз | ФИО;Отдел;Табельный_номер |

Или несколько шагов подряд (выполняются сверху вниз):

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | тонкая_сетка | |
| Постобработка_Диапазон | заполнение_вниз | ФИО;Отдел |
| Постобработка_Диапазон | перенос | |

Через **визард параметров** (`param_wizard.set_merge_param`): вкладка постобработки → в B выберите `заполнение_вниз` из списка → в C введите столбцы.

---

### Шаг E. Поведение на листе результата

До (столбец «Отдел», фрагменты пустые после сбора из разных файлов):

| #Путь | ФИО | Отдел |
|-------|-----|-------|
| src_01#… | Иванов | Продажи |
| src_01#… | | |
| src_02#… | Петров | |

После `заполнение_вниз` с `C = ФИО;Отдел`:

| #Путь | ФИО | Отдел |
|-------|-----|-------|
| src_01#… | Иванов | Продажи |
| src_01#… | Иванов | Продажи |
| src_02#… | Петров | IT |

Первая строка данных в диапазоне постобработки не меняется; заполнение идёт со второй строки данных (`sr + 1` … `er`).

---

### Шаг F. Отладка

- Журнал на листе **«Сбор_книг_лог»** — записи с меткой `заполнение_вниз`.
- Если функция «не видна» в списке B визарда: проверьте `LM_PP_PUBLIC_NAMES`, `_PP_RANGE_MAP_SPEC` и что установлены **оба** файла (`collect_workbooks.py` + `pythonpath/libre_macros_lib.py`).
- Если столбец не находится: проверьте заголовок в **строке 0** листа результата и используйте точное имя или номер/букву.

---

## Полный пример: «заполнение вниз с вычислением»

Расширенная версия `заполнение_вниз`: в пустые ячейки записывается **вычисленное значение** по формуле или константа, а не копия из строки выше.

**Колонка C:** пары `столбец/формула` через `,` или `;`:

- **столбец**: номер (1-based), буква (`A`, `AA`) или имя заголовка;
- **формула**: формула Calc (начинается с `=`) или константа;
- **шаблоны**: `{row_id}` → номер строки (1-based), `{row_id-1}` → предыдущая строка.

Примеры пар:
- `ФИО/=A{row_id}&" "&B{row_id}` — конкатенация столбцов A и B с пробелом;
- `Сумма/=C{row_id-1}*2` — удвоить значение из предыдущей строки;
- `Категория/общий` — константа "общий" во все пустые ячейки;
- `3/=SUM(C{row_id-1}:E{row_id-1})` — сумма диапазона в столбце 3 (C).

Пример C: `ФИО/=A{row_id}&" "&B{row_id};Сумма/=C{row_id-1}*2`.

---

### Регистрация

**`libre_macros_lib.py`** — добавлено в `LM_PP_PUBLIC_NAMES`:

```python
    "lm_pp_range_fill_down_calculate"
```

Для `замена_значений`:

```python
    "lm_pp_range_replace_values"
```

**`collect_workbooks.py`** — в `_PP_RANGE_MAP_SPEC`:

```python
    ("заполнение_вниз_вычислить", "merge_pp_range_fill_down_calculate")
```

И для `замена_значений`:

```python
    ("замена_значений", "merge_pp_range_replace_values")
```

И для `текстовые_операции`:

```python
    ("текстовые_операции", "merge_pp_range_normalize_text")
```

И для `добавить_столбец`:

```python
    ("добавить_столбец", "merge_pp_range_add_column")  # + FINAL: merge_final_add_column
```

Публичное имя: `lm_pp_range_add_column` / `lm_final_add_column`. JSON и визард: [13_JSON_PARAMS.md](13_JSON_PARAMS.md), [05_POSTPROCESS.md](05_POSTPROCESS.md).

---

### Лямбды `rows` / `upd_rows`

Единое ядро для всех лямбд проекта живёт в `libre_macros_lambda_column_lib.py`.

Что считается стандартом:

- `rows("Колонка")` — текущая строка;
- `rows("Колонка")[-1]` — предыдущая строка листа;
- `upd_rows("Колонка")` — последняя уже обработанная строка выше;
- `upd_rows_seen("Колонка", value)` — поиск значения выше;
- `upd_rows_count("Колонка", value)`, `upd_rows_len()` — статистика по накопителю.

Где это уже применяется:

- `замена_значений`: `row_filter`, `match_expr`;
- `текстовые_операции`: `row_filter_lambda` у отдельных `ops`;
- `столбец_по_лямбде`: `expr`, `new_column_expr`, `sheet_expr`.

Рабочий JSON-пример:

```json
[{
  "v": 1,
  "fn": "столбец_по_лямбде",
  "new_column": "Первый_ИНН",
  "expr": "lambda rows: rows(\"ИНН\") if not upd_rows_seen(\"ИНН\", rows(\"ИНН\")) else \"\"",
  "position": "_конец_",
  "on_error": "empty",
  "result_type": "auto",
  "output_policy": "new"
}]
```

Ключевые поля:

- `expr` — основная лямбда `lambda rows: ...`, возвращает значение ячейки для текущей строки;
- `new_column_expr` — лямбда `lambda sheet, headers: ...`, вычисляет имя столбца один раз на лист;
- `sheet_expr` — лямбда `lambda sheets: ...`, выбирает лист или список листов из кандидатов;
- `on_error` — поведение при ошибке в строке: `empty`, `keep`, `fail`;
- `result_type` — тип результата: `auto`, `text`, `number`;
- `output_policy` — поведение при найденном заголовке: `new` или `replace`.

Минимальные примеры:

```python
lambda rows: str(rows("Отдел") or "").strip() != ""
lambda rows: rows("Статус") != rows("Статус")[-1]
lambda rows: not upd_rows_seen("ИНН", rows("ИНН"))
lambda rows: upd_rows_len() == 0 or rows("Сумма") != upd_rows("Сумма")
lambda rows: rows("Фамилия") + " " + rows("Имя")
lambda rows: rows("Сумма") if upd_rows_len() == 0 else rows("Сумма") - upd_rows("Сумма")
```

В `MERGE_POSTPROCESS_RANGE_REST_HINTS`:

```python
    "заполнение_вниз_вычислить": (
        "JSON: columns + rules (column, formula, expand_to_right); диалог «Параметры…»",
        '[{"v":1,"fn":"заполнение_вниз_вычислить","columns":["A"],"rules":[{"column":"A","formula":"=A{row-1}"}]}]',
    ),
```

**`param_wizard.py`** — диалог `_show_block_rules_param_dialog`: правило `Столбец` / `Формула` / ☑ «Вправо» (`expand_to_right`).

---

### Строка на листе параметров

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | заполнение_вниз_вычислить | ФИО/=A{row_id}&" "&B{row_id} |

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | заполнение_вниз | ФИО;Отдел |
| Постобработка_Диапазон | заполнение_вниз_вычислить | Сумма/=C{row_id-1}*2;Категория/итог |

---

### Поведение

До:

| ФИО (пусто) | Сумма (пусто) |
|-------------|---------------|
| Иванов | 100 |
| | |
| Петров | 200 |
| | |

После `C = ФИО/=A{row_id}&"_"&{row_id};Сумма/=C{row_id-1}*2`:

| ФИО | Сумма |
|-----|-------|
| Иванов | 100 |
| Иванов_2 | 200 (100*2) |
| Петров | 200 |
| Петров_4 | 400 (200*2) |

Значение вычисляется заново для каждой пустой ячейки; если формула начинается с `=`, она записывается в `cell.Formula`, иначе как `cell.String`.

---

## См. также

- `macro-lib/USER_PP_FUNCTIONS_GUIDE.md` — свои функции: `functions_pp.py`, lambda/def в B, eval
- `docs/05_POSTPROCESS.md` — поведение встроенных колбэков для пользователя
- `macro-lib/functions_pp.py` — примеры пользовательских колбэков
- `macro-lib/DATA_COLLECTION_DESCRIPTION.md` — параметры сбора и постобработки
