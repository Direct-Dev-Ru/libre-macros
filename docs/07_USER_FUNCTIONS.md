# Пользовательские функции постобработки

Как писать **свои** шаги постобработки без правки `libre_macros_lib.py`: отдельный файл `functions_pp.py`, inline-код в колонке **B** или комбинация с параметрами в **C**.

Актуально для макроса **collect_workbooks 3.10.x** (встроенные ключи B, визард `param_wizard`, диалог сводной — `show_pivot_table_param_dialog` внутри `param_wizard.py`).

Для функций в общей библиотеке (`lm_pp_*` → ключи вроде `заполнение_вниз`) см. **[11_LIB_REFERENCE.md](11_LIB_REFERENCE.md)**. Справочник всех встроенных ключей B и параметров C — **`05_POSTPROCESS.md`**.

---

## Три способа указать функцию в колонке B

| Способ | Пример B | Когда использовать |
|--------|----------|-------------------|
| **Ключ карты** | `тонкая_сетка` | Встроенная функция из `libre_macros_lib.py` |
| **Файл#функция** | `functions_pp.py#pp_range_zagotovka` | Свой `.py` рядом с макросом, удобно версионировать |
| **Python-код** | `lambda doc,sheet,dr,hr: …` или `def my_pp(...):` | Быстрый одноразовый шаг без отдельного файла |

Порядок поиска в `resolve_postprocess_function()` (`collect_workbooks.py`):

1. ключ в `MERGE_RESULT_POSTPROCESS_*_MAP`;
2. `__name__` функции из карты;
3. `путь.py#имя_функции`;
4. inline `lambda` / `def` через `eval` / `exec`.

Колонка **C** передаётся как `*extra_args`. Для встроенных ключей с кодеком (`libre_macros_param_codec.py`) JSON в C автоматически преобразуется в канонический текст перед вызовом `lm_pp_*`. Особые разборы остаются у legacy-функций без C (`тонкая_сетка` как пресет и т.д.).

---

## Установка `functions_pp.py`, `functions_final.py` и связанных модулей

```bash
cd macro-lib
./install.sh collect_workbooks param_wizard
# или явно:
./install.sh functions_pp.py functions_final.py collect_workbooks param_wizard
```

- `functions_pp.py` — пользовательские `pp_*` (колонка B: `functions_pp.py#имя`);
- `functions_final.py` — пользовательские функции финальной обработки (`functions_final.py#имя`);
- `param_wizard` — визард параметров, пресеты и диалог макета для ключа **`сводная_таблица`**;
- `collect_workbooks` + `pythonpath/libre_macros_lib.py` — встроенные `merge_pp_*` / ключи карты.

**Кнопка «Параметры…»** в визарде (`param_wizard`) доступна для большинства встроенных ключей постобработки (**Постобработка_Диапазон**, **Постобработка_Строка**, **Финальная_обработка**): открывает микродиалог с полями или JSON-редактором. По умолчанию параметры сохраняются как **JSON-массив** в колонке C (см. `05_POSTPROCESS.md`, раздел «Формат параметров»). Исключения: **ВПР**, **объединить_листы_в_один** / **разделить_листы** (отдельные диалоги, имя в A, JSON в C), **сводная_таблица** (диалог макета).

Файлы копируются в `~/.config/libreoffice/4/user/Scripts/python/` (или каталог из `LO_MACROS_DIR`). Относительный путь в B: `functions_pp.py#имя`.

После изменений — снова `install.sh` или копирование вручную; для **списков визарда** из `libre_macros_lib` иногда нужен перезапуск Calc. Кэш `файл.py#func` — по `mtime` файла.

## Сигнатуры

### Диапазон (A = `Постобработка_Диапазон`)

```python
def pp_range_имя(doc, sheet, data_range, header_row_range, *extra_args):
    ...
```

- `data_range` — все строки данных (заголовок обычно отдельно в `header_row_range`);
- `extra_args` — кортеж строк из колонки C.

### Строка (A = `Постобработка_Строка`)

Имя функции лучше начинать с **`pp_row_`** — тогда визард добавит её в список строки.

```python
def pp_row_имя(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    ...
```

- `data_row_range` — одна строка данных;
- `row_index` — 0-based индекс строки на листе.

---

## Механизм `файл.py#функция`

1. **Разбор B:** одна `#`, слева путь к `.py`, справа имя функции (`isidentifier()`).
2. **Путь:** абсолютный или относительно каталога `collect_workbooks.py` (`_merge_macro_script_dir()`).
3. **Извлечение:** из файла вырезается текст `def имя(...)` (AST или построчно).
4. **Компиляция:** тот же `exec`, что и для inline-кода в B.
5. **Кэш:** `(abs_path, func_name) → (mtime, callable)` — при изменении файла перечитывается.

Импорты в `functions_pp.py` **не обязательны** для хелперов макроса: при `exec` подставляется namespace `_merge_pp_postprocess_eval_namespace()`.

---

## Механизм `lambda` / `def` в колонке B

Код выполняется в ограниченном окружении `_merge_pp_postprocess_eval_namespace()`:

- **Доступно:** все ключи карт (`тонкая_сетка`, `переименовать_лист`, …) как вызываемые функции, `merge_pp_*`, `cell_text`, `lo_color_rgb`, `merge_pp_active_context`, `_merge_pp_range_address`, `_merge_pp_foreach_row`, `merge_ui_yield`, `merge_ui_status_set`, `merge_pp_user_log`, `traceback`, константы `MERGE_UI_*_EVERY_ROWS`.
- **Запрещено в builtins:** `__import__`, `eval`, `exec`, `open`, `compile`, `globals`, `locals` и др. (`_MERGE_PP_EVAL_BUILTIN_BLOCKLIST`).

Вызов встроенной функции **по ключу карты** (имя из колонки B как функция):

```python
lambda doc,sheet,dr,hr: переименовать_лист(doc, sheet, dr, hr, "Итоговый_отчёт")
```

То же через `merge_pp_*`:

```python
lambda doc,sheet,dr,hr: merge_pp_range_rename_sheet(doc, sheet, dr, hr, "Итоговый_отчёт")
```

**Lambda (одна операция):**

```python
lambda doc,sheet,dr,hr: merge_pp_user_log(doc, sheet, "готово", "ok")
```

**Def (рекомендуется для нескольких операций и циклов по заголовкам):**

```python
def fn(doc, sheet, data_range, header_row_range):
    sc, sr, ec, er = _merge_pp_range_address(data_range)
    merge_pp_user_log(doc, sheet, "строк %d" % (er - sr + 1), "ok")
```

Многострочный `def` в ячейке Calc неудобен — для сложной логики используйте `functions_pp.py`.

Важно: это **не** те же лямбды, что `row_filter` / `match_expr` / `row_filter_lambda` в JSON-параметрах.  
Inline `lambda` в колонке **B** получает аргументы `doc, sheet, ...` и работает как мини-колбэк постобработки.  
Лямбды в JSON работают в sugar-режиме `lambda rows: …` и понимают `rows(...)`, `upd_rows(...)`, `upd_rows_seen(...)`, `upd_rows_count(...)`, `upd_rows_len()`.

Примеры JSON-лямбд для встроенных функций:

```python
lambda rows: str(rows("Отдел") or "").strip() != ""
lambda rows: rows("Статус") != rows("Статус")[-1]
lambda rows: not upd_rows_seen("ИНН", rows("ИНН"))
lambda rows: upd_rows_len() == 0 or rows("Сумма") != upd_rows("Сумма")
```

---

## Шаблон: диапазон (`functions_pp.py`)

Полный каркас с циклом по строкам, статусом и логом (как `pp_range_zagotovka`):

```python
def pp_range_moya(doc, sheet, data_range, header_row_range, *extra_args):
    """
  Параметры C: param1, param2 (через запятую).
  B = functions_pp.py#pp_range_moya
  C = highlight,Сумма
    """
    param1 = extra_args[0] if len(extra_args) > 0 else ""
    param2 = extra_args[1] if len(extra_args) > 1 else ""

    try:
        sc, sr, ec, er = _merge_pp_range_address(data_range)
        _hc, hdr_r, _he, _hr = _merge_pp_range_address(header_row_range)
        if sr <= hdr_r:
            sr = hdr_r + 1
        if er < sr:
            merge_pp_user_log(doc, sheet, "нет строк данных", "пропуск")
            return

        row = sr
        n = 0
        while row <= er:
            # --- логика на строку row (0-based) ---
            col = sc
            while col <= ec:
                cell = sheet.getCellByPosition(col, row)
                # пример: if param1 == "highlight": cell.CellBackColor = 0xFFFF00
                col = col + 1
            n = n + 1
            if n == 1 or row == er or n % MERGE_UI_STATUS_EVERY_ROWS == 0:
                merge_ui_status_set(n, "pp_range_moya — %d/%d" % (row + 1, er + 1))
            merge_ui_yield(counter=n)
            row = row + 1

        merge_pp_user_log(doc, sheet, "обработано строк: %d" % n, "ok")
    except Exception as err:
        merge_pp_user_log(doc, sheet, "ОШИБКА: %s\n%s" % (err, traceback.format_exc()), "ошибка")
        raise
```

---

## Шаблон: строка (`pp_row_*`)

```python
def pp_row_podsvetka(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
  B = functions_pp.py#pp_row_podsvetka
  C = Отдел
    """
    col_name = extra_args[0] if len(extra_args) > 0 else ""
    if col_name == "":
        return
    try:
        sc, sr, ec, er = _merge_pp_range_address(data_row_range)
        # найти столбец по заголовку, изменить ячейку в row_index
        hdr_sc, hdr_sr, hdr_ec, hdr_er = _merge_pp_range_address(header_row_range)
        c = hdr_sc
        while c <= hdr_ec:
            title = cell_text(sheet.getCellByPosition(c, hdr_sr))
            if col_name.lower() in title.lower():
                cell = sheet.getCellByPosition(c, row_index)
                cell.CellBackColor = 0xE0E0E0
                break
            c = c + 1
    except Exception as err:
        merge_pp_user_log(doc, sheet, "ОШИБКА: %s" % err, "ошибка")
        raise
```

---

## Шаблон: заполнение вниз (`functions_pp.py#pp_range_fill_down`)

Пользовательская реализация функции «заполнение вниз» (альтернатива встроенной `заполнение_вниз` из `libre_macros_lib`). Пустые ячейки в указанных столбцах заполняются копией из строки выше.

```python
def pp_range_fill_down(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Пользовательский вариант заполнения вниз: пустые ячейки → значение из строки выше.

    Параметры из колонки C: имена столбцов через , или ; — номер (1=A), буква или имя заголовка.
    Пример C: ФИО;Отдел;3

    Использование в B: functions_pp.py#pp_range_fill_down
    """
    try:
        sc, sr, ec, er = _merge_pp_range_address(data_range)
        _hc, hdr_r, _he, _hr = _merge_pp_range_address(header_row_range)
        if sr <= hdr_r:
            sr = hdr_r + 1
        if er < sr:
            merge_pp_user_log(doc, sheet, "нет строк данных", "пропуск")
            return

        # Разобрать параметры C
        spec = ", ".join(str(x) for x in extra_args) if extra_args else ""
        cols_to_fill = []
        if spec:
            for part in spec.replace(";", ",").split(","):
                p = part.strip()
                if not p:
                    continue
                # Поиск столбца по номеру, букве или имени заголовка
                col_idx = None
                try:
                    col_idx = int(p) - 1  # 1-based → 0-based
                except ValueError:
                    pass
                if col_idx is None and len(p) <= 3 and p.isalpha():
                    # Буква столбца (A, AA...)
                    col_idx = 0
                    for ch in p.upper():
                        col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)
                    col_idx -= 1
                if col_idx is None:
                    # Поиск по заголовкам
                    c = sc
                    while c <= ec:
                        hdr_text = cell_text(sheet.getCellByPosition(c, hdr_r))
                        if p.lower() in hdr_text.lower():
                            col_idx = c
                            break
                        c = c + 1
                if col_idx is not None and col_idx not in cols_to_fill:
                    cols_to_fill.append(col_idx)

        filled = 0
        for col in cols_to_fill:
            row = sr + 1
            while row <= er:
                cell = sheet.getCellByPosition(col, row)
                if cell_text(cell) == "":
                    src = sheet.getCellByPosition(col, row - 1)
                    try:
                        cell.String = src.String
                        filled = filled + 1
                    except Exception:
                        try:
                            cell.Value = src.Value
                            filled = filled + 1
                        except Exception:
                            pass
                row = row + 1

        merge_pp_user_log(
            doc, sheet,
            "заполнение_вниз (functions_pp.py): столбцов %d, заполнено %d" % (len(cols_to_fill), filled),
            "ok"
        )
    except Exception as err:
        merge_pp_user_log(doc, sheet, "ОШИБКА: %s" % err, "ошибка")
        raise
```

**Отличия от встроенной `заполнение_вниз`:**
- Это образец для копирования и модификации;
- Можно изменить логику копирования (например, копировать не всё подряд, а по условию);
- Использует `functions_pp.py#pp_range_fill_down` в колонке B вместо ключа `заполнение_вниз`.

**Строка на листе параметров:**

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | functions_pp.py#pp_range_fill_down | ФИО;Отдел;3 |

---

## Несколько операций в одном шаге (пример: лист + заголовки столбцов)

Встроенный ключ **`переименовать_лист`** меняет только имя листа (C = новое имя). Заголовки столбцов — **`переименовать_столбцы`** (JSON `mappings`). Порядок столбцов — **`переставить_столбцы`** (JSON `columns` + `position`).

### Вариант 1 — две строки на листе (проще всего)

Порядок сверху вниз: сначала лист, потом заголовки.

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | переименовать_лист | Итоговый_отчёт |
| Постобработка_Диапазон | `def fn(doc,sheet,dr,hr): ...` (только заголовки, см. вариант 2) | |

Во второй строке B — короткий `def fn` без переименования листа (только замена `#Путь` → `Источник` и префикс `поле_` для остальных непустых ячеек заголовка).

### Вариант 2 — один шаг `def fn` в колонке B

Подходит, когда обе операции нужны в одной ячейке. Имя функции — **`fn`** (так ожидает `exec` в макросе). C можно оставить пустой.

```python
def fn(doc, sheet, data_range, header_row_range):
    merge_pp_range_rename_sheet(doc, sheet, data_range, header_row_range, "Итоговый_отчёт")
    hc, hr, he, _ = _merge_pp_range_address(header_row_range)
    subs = (("#Путь", "Источник"), ("#ДатаВремяФайла", "Дата_файла"))
    c = hc
    while c <= he:
        cell = sheet.getCellByPosition(c, hr)
        t = cell_text(cell)
        for old, new in subs:
            if t == old:
                cell.String = new
                break
        else:
            if t != u"" and not t.startswith(u"_"):
                cell.String = u"поле_" + t
        c = c + 1
    merge_pp_user_log(doc, sheet, u"лист и заголовки переименованы", u"ok")
```

Тот же код удобнее хранить в **`functions_pp.py`** (см. вариант 3).

### Вариант 3 — `functions_pp.py` (рекомендуется для сопровождения)

Добавьте в `functions_pp.py`:

```python
def pp_range_rename_sheet_and_headers(doc, sheet, data_range, header_row_range, *extra_args):
    """
    C: новое_имя_листа [, старый=новый ...]
    Пример C: Итоговый_отчёт, #Путь=Источник, #ДатаВремяФайла=Дата_файла
    B: functions_pp.py#pp_range_rename_sheet_and_headers
    """
    new_sheet = extra_args[0].strip() if extra_args else u""
    if new_sheet:
        merge_pp_range_rename_sheet(doc, sheet, data_range, header_row_range, new_sheet)
    hc, hr, he, _ = _merge_pp_range_address(header_row_range)
    subs = []
    i = 1
    while i < len(extra_args):
        pair = extra_args[i].strip()
        if u"=" in pair:
            old, new = pair.split(u"=", 1)
            subs.append((old.strip(), new.strip()))
        i = i + 1
    c = hc
    n = 0
    while c <= he:
        cell = sheet.getCellByPosition(c, hr)
        t = cell_text(cell)
        for old, new in subs:
            if t == old:
                cell.String = new
                n = n + 1
                break
        c = c + 1
    merge_pp_user_log(doc, sheet, u"лист + %d заголовков" % n, u"ok")
```

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | functions_pp.py#pp_range_rename_sheet_and_headers | Итоговый_отчёт, #Путь=Источник, #ДатаВремяФайла=Дата_файла |

### Вариант 4 — одна `lambda` (только простые цепочки)

`lambda` — **одно выражение**. Две встроенные операции без своего цикла:

```python
lambda doc,sheet,dr,hr: (merge_pp_range_thin_grid_borders(doc,sheet,dr,hr), merge_pp_range_rename_sheet(doc,sheet,dr,hr,"Итог"))[1]
```

Переименование листа **и** правка заголовков в одной `lambda` получается нечитаемым (нужен цикл `while` или `def`). Для «лист + заголовки» используйте **вариант 1–3**.

Компактная `lambda` **только для заголовков** (замена одной подписи):

```python
lambda doc,sheet,dr,hr: (lambda hc,hr,he,_: [sheet.getCellByPosition(c,hr).setString(u"Источник") for c in range(hc,he+1) if cell_text(sheet.getCellByPosition(c,hr))==u"#Путь"])(*_merge_pp_range_address(hr)) or merge_pp_user_log(doc,sheet,u"заголовок #Путь","ok")
```

---

## Шаблоны lambda (колонка B)

**Лог в журнал:**

```python
lambda doc,sheet,dr,hr: merge_pp_user_log(doc, sheet, "шаг выполнен", "ok")
```

**Вызов встроенной функции по имени из карты** (в namespace есть ключи карт):

```python
lambda doc,sheet,dr,hr: merge_pp_range_thin_grid_borders(doc, sheet, dr, hr)
```

**С параметром из C** (C не попадает в lambda в B — нужен `def` в файле или обёртка через `functions_pp.py`).

Для параметров из C без отдельного файла — только `def` в B:

```python
def fn(doc, sheet, data_range, header_row_range, color_name):
    pass
```

Но тогда C должен содержать один аргумент; обёртка `_merge_pp_wrap_callback` добавит его в вызов.

---

## Строка на листе параметров

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | functions_pp.py#pp_range_zagotovka | 42,тест |
| Постобработка_Диапазон | тонкая_сетка | |
| Постобработка_Диапазон | переименовать_лист | Итоговый_отчёт |
| Постобработка_Диапазон | сводная_таблица | `{"v":1,"sheet_name":"Сводная",...}` (через визард) |
| Постобработка_Диапазон | lambda doc,sheet,dr,hr: merge_pp_user_log(doc,sheet,"ok","ok") | |
| Постобработка_Строка | functions_pp.py#pp_row_podsvetka | Отдел |

Порядок шагов — **сверху вниз**. Сначала выполняются все строки `Постобработка_Диапазон`, затем все `Постобработка_Строка`.

---

## Визард параметров (`param_wizard`)

- Ключи вроде `тонкая_сетка`, `переименовать_лист`, `сводная_таблица` — из `MERGE_RESULT_POSTPROCESS_RANGE_MAP` (после `install.sh collect_workbooks`; списки B иногда требуют перезапуск Calc).
- **`сводная_таблица`** — в колонке B выберите ключ, нажмите **«Параметры…»** справа от поля C: откроется диалог макета из `param_wizard` (поля, имя листа, «как значения»). JSON пишется в C. Достаточно установленного `param_wizard.py`.
- `functions_pp.py#pp_*` — подхватываются `_scan_functions_pp_choices()` автоматически, если файл установлен рядом с макросом.
- Inline `lambda` / `def` — ввод вручную в B (визард разрешает значение вне списка).

Подробнее про списки визарда — **`06_PARAM_WIZARD.md`**, **[11_LIB_REFERENCE.md](11_LIB_REFERENCE.md)** (п. 2.6).

---

## Встроенные функции (libre_macros_lib, выборка)

Помимо `functions_pp.py#pp_range_fill_down` есть **встроенные** функции в `libre_macros_lib.py`:

| Функция | Ключ в B | Параметр C | Назначение |
|---------|----------|------------|------------|
| `lm_pp_range_fill_down_empty` | `заполнение_вниз` | Столбцы через `,` или `;` | Пустые ячейки → копия из строки выше |
| `lm_pp_range_fill_down_calculate` | `заполнение_вниз_вычислить` | JSON: `columns` + `rules` | Пустые ячейки → формула/константа; опц. `expand_to_right` |
| `lm_pp_range_replace_values` | `замена_значений` | JSON: `columns`, `find`, `replace`, флаги, `row_filter`, `match_pick`; `<<Переменные…>>` в find/replace | Замена подстрок в текстовых столбцах (один getDataArray/setDataArray) |
| `lm_pp_range_add_column` / `lm_final_add_column` | `добавить_столбец` | JSON: `new_column`, `data_type`, `format`, `choices` / `data_validation`+`min`/`max` | Пустой столбец + формат + контроль данных (LIST / DECIMAL / DATE) |
| `lm_pp_range_rename_sheet` | `переименовать_лист` | Новое имя листа | При конфликте — суффикс `_2`, `_3`, … |
| `lm_pp_range_pivot_table` | `сводная_таблица` | JSON макета (визард) | Новый лист со сводной таблицей по данным диапазона |
| `lm_pp_range_value_count` | `количество_значений` | JSON: `new_column`, `key_columns` | Число строк с тем же ключом в диапазоне |
| `lm_pp_range_remove_duplicates` | `удалить_дубликаты` | JSON: `key_columns`, `keep`, `output`, … | Удалить повторы строк по ключу |
| `lm_pp_range_copy_sheet` | `копировать_лист` | JSON: `source_sheet`, `dest_sheet` | Полная копия вкладки в книге |
| `lm_pp_range_copy_ranges` / `lm_final_copy_ranges` | `копирование_диапазонов` | JSON: `source_*` → `dest_*`, `mode`, `content`, … | Блок ячеек между листами (replace/insert) |

### `копирование_диапазонов`

Копирует прямоугольник / строки / столбцы с **листа-источника** на один или несколько **листов-приёмников** той же книги (RANGE и Финал).

| A | B | C |
|---|---|---|
| Финальная_обработка | копирование_диапазонов | JSON (см. ниже) |
| Постобработка_Диапазон | копирование_диапазонов | тот же JSON |

Алиасы B: `копировать_диапазон`, `copy_ranges`, `copy_range`, `вставить_диапазон`.

```json
[{"v":1,"fn":"копирование_диапазонов","source_sheet":"Сводная","source_range":"A1:C10","dest_sheet":"Отчет","dest_cell":"A1","mode":"replace"}]
```

- `mode`: `replace` (поверх) или `insert` (сдвиг строк/столбцов перед якорем);
- `content`: `values` (default) или `formulas`;
- после копирования приёмник по умолчанию **вовлекается** в скоп следующих шагов (`involve_dest`).

Полная таблица полей и примеры — [13_JSON_PARAMS.md](13_JSON_PARAMS.md); описание режимов — [05_POSTPROCESS.md](05_POSTPROCESS.md); форма визарда — [06_PARAM_WIZARD.md](06_PARAM_WIZARD.md).

### `переименовать_лист`

**C:** одно имя, например `Итоговый_отчёт`. Переименовывается **текущий лист результата**.

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | переименовать_лист | Сводка_2024 |

### `сводная_таблица`

Создаёт сводную на новом листе по собранному диапазону. **C** — JSON (удобно через визард → «Параметры…»).

Функции данных: `SUM` / `COUNT` / … и псевдоагрегат **`LIST_ROWS`** («Список_строк»): перечень исходных строк по `list_fields`. Требует `as_values: true`. Внутри — COUNT, после материализации заголовок → «Список_строк». Режим вывода: в одной ячейке (`\n`) или **`list_rows_expand`** — каждая запись на отдельной строке листа (ключ группы только в первой).
### `заполнение_вниз`

**Простое копирование** значения из предыдущей строки.

- **C:** `ФИО;Отдел;3` — список столбцов (номер 1-based, буква `A`, или имя заголовка);
- Пустые ячейки заполняются копией из `row-1`;
- Первая строка данных не меняется.

**Пример:**

До (пропуски после сбора из разных файлов):

| ФИО | Отдел |
|-----|-------|
| Иванов | Продажи |
| *(пусто)* | *(пусто)* |
| Петров | IT |

→ после `C = ФИО;Отдел`:

| ФИО | Отдел |
|-----|-------|
| Иванов | Продажи |
| Иванов | Продажи |
| Петров | IT |

### `заполнение_вниз_вычислить`

**Вычисление значения** по формуле или константа в пустых ячейках указанных столбцов.

**JSON (кнопка «Параметры…» в визарде):**

```json
[{"v":1,"fn":"заполнение_вниз_вычислить","columns":["A"],"rules":[{"column":"A","formula":"=A{row-1}","expand_to_right":true}]}]
```

| Поле | Описание |
|------|----------|
| `columns` | Столбцы с пропусками |
| `rules[].column` | Столбец правила (сопоставляется с `columns`) |
| `rules[].formula` | `=…` или константа |
| `rules[].expand_to_right` | Заполнить пустые ячейки **вправо** в тех же строках (до границы диапазона) |

**Шаблоны:** `{row_id}`, `{row-1}`, `R[-1]C` (в A1-листе → A1-ссылки).

**Legacy-текст C:** пары `столбец/формула` через `,` или `;` (без `expand_to_right`).

**Примеры пар (legacy):**

| Пара C | Результат |
|--------|-----------|
| `ФИО/=A{row_id}&" "&B{row_id}` | Конкатенация колонок A и B |
| `Сумма/=C{row_id-1}*2` | Удвоить значение из предыдущей строки |
| `Категория/общий` | Константа "общий" во все пустые ячейки |

**Пример строки (legacy):**

| A | B | C |
|---|---|---|
| Постобработка_Диапазон | заполнение_вниз_вычислить | ФИО/=A{row_id}&" "&B{row_id};Сумма/=C{row_id-1}*2 |

**Результат:**

До:

| Фамилия | Имя | Сумма |
|---------|-----|-------|
| Иванов | Алексей | 100 |
| *(пусто)* | *(пусто)* | *(пусто)* |

→ после (для строки 3, row_id=3):

| Фамилия | Имя | Сумма |
|---------|-----|-------|
| Иванов | Алексей | 100 |
| `=A3&" "&B3` → "Иванов Алексей" | `=B3` → "Алексей" | `=C2*2` → 200 |

**Когда использовать какую:**

- **`заполнение_вниз`** — простое дублирование (категории, статусы, ФИО из предыдущей строки);
- **`заполнение_вниз_вычислить`** — вычисляемые значения (код из префикса+номера, прогрессивная сумма, округление, конкатенация нескольких колонок).

---

## Доступные хелперы в пользовательском коде

| Имя | Назначение |
|-----|------------|
| `_merge_pp_range_address(range)` | `(start_col, start_row, end_col, end_row)` 0-based |
| `_merge_pp_foreach_row(sheet, sr, er, fn, status_prefix)` | Цикл по строкам с yield |
| `_merge_pp_number_format_key(doc, fmt)` | Ключ числового формата |
| `cell_text(cell)` | Текст ячейки |
| `lo_color_rgb(r, g, b)` | Цвет LibreOffice |
| `merge_pp_active_context()` | dict контекста (границы, `total_row`, …) |
| `merge_pp_user_log(doc, sheet, msg, status)` | Запись в «Сбор_книг_лог» |
| `merge_ui_yield(counter)` | Отдача UI |
| `merge_ui_status_set(n, text)` | Статусная строка |
| `merge_pp_range_*` / `merge_pp_row_*` | Все встроенные колбэки после загрузки lib |
| `переименовать_лист`, `тонкая_сетка`, … | Ключи карт — вызываются как функции в `lambda`/`def` в B |

---

## Отладка

| Симптом | Что проверить |
|---------|----------------|
| `⚠ Файл постобработки не найден` | `./install.sh functions_pp.py`, путь B = `functions_pp.py#…` |
| `⚠ В файле … нет функции def …` | имя в B совпадает с `def` в файле |
| `⚠ Ошибка Python-постобработки в B` | синтаксис, traceback в консоли Python / «Сбор_книг_лог» |
| `⚠ Неизвестная функция постобработки` | опечатка в B, нет ключа в карте, неверный `#` |
| `'merge_pp_range_…'` KeyError | не установлен `pythonpath/libre_macros_lib.py` — `./install.sh collect_workbooks` |
| Функция не в выпадающем списке визарда | для lib — перезапуск Calc; для `functions_pp.py` / `functions_final.py` — имя `pp_*` и install |

---

## Пример из репозитория

Готовые образцы: **`macro-lib/functions_pp.py`**

- `pp_range_zagotovka` — цикл по диапазону, `extra_args`, UI;
- `pp_stroka_itogo` — строка «Итого» с формулами `SUM`;
- `pp_range_fill_down` — заполнение вниз по столбцам из C.

Шаблон финальной обработки: **`macro-lib/functions_final.py`** (`pp_range_zagotovka` — проход по всем листам книги).

---

## См. также

- `11_LIB_REFERENCE.md` — добавление функций в общую библиотеку
- `05_POSTPROCESS.md` — справочник встроенных ключей B и параметров C
- `06_PARAM_WIZARD.md` — визард, пресеты, ВПР, сводная таблица
- `04_DATA_COLLECTION.md` — лист параметров сбора
