# JSON-параметры постобработки (колонка C)

**Версия макроса:** **3.10.696** (`macro-lib/version.txt`).

**Только JSON.** Строки `Постобработка_Диапазон` / `Постобработка_Строка` / `Финальная_обработка` / `Постобработка_xml`, а также отдельные параметры pipeline **`объединить_листы_в_один`**, **`разделить_листы`**, **`ВПР`** (если выбран JSON) хранят в **колонке C** компактный JSON:

```json
[{"v":1,"fn":"тонкая_сетка"}]
```

Несколько листов в одной функции — несколько объектов в массиве. Поле `sheet` (или `sheets` для списковых функций) ограничивает выполнение конкретными листами результата.

Кодек: `macro-lib/pythonpath/libre_macros_param_codec.py`  
Готовые строки для шаблонов: `test/collect_workbooks/manual_json_presets.py`  
Матрица ручной проверки: сценарий `12_postprocess_json.xltx`.

---

## ВПР (JSON в C)

Строка в A: `ВПР`. Предпочтительная запись — JSON-массив в **C** (legacy B–F ещё читается).

```json
[{
  "v": 1,
  "fn": "впр",
  "left":  {"sheet": "Merge", "range": "A1:G200"},
  "right": {"sheet": "Справочник", "range": "A1:D50"},
  "join_type": "left",
  "keys_same": true,
  "keys_left": ["'Код'"],
  "keys_right": ["'Код'"],
  "extract_columns": ["'Наименование'", "'Цена'"],
  "highlight_color": "голубой",
  "fill_duplicates": true,
  "not_found_fill": "#Н/Д",
  "multi_match": "all",
  "multi_match_one_cell": false,
  "trim_keys": false,
  "column_suffix": true
}]
```

| Поле | Описание |
|------|----------|
| `left` / `right` | `sheet`; опционально `range` (A1); при `range` — derived `start_row` / `start_col` / `header_row` (заголовок = первая строка диапазона). Без `range` — used-area, как раньше |
| `join_type` | `left` (по умолчанию), `inner` (только совпадения; строки без ключа справа удаляются), `full` (два прохода + строки только справа на левом листе) |
| `multi_match` | При нескольких строках справа по ключу: `all` (по умолчанию — все, с учётом «Дубли»/«Мульти»), `first` (первая), `last` (последняя) |
| `extract_columns` | Колонки **в левую из правой**; `'Имя'`, `'Имя -> Новое'`, буква/номер |
| `extract_columns_right` | При `full` — колонки **в правую из левой** |
| `keys_left` / `keys_right` | Ключи; при `keys_same` достаточно `keys_left` |
| `highlight_color` | Подсветка новых колонок (hex / пресет); пусто / «Нет» — без |
| `fill_duplicates` | При `multi_match=all`: заполнять дублирующие строки (Да) или только первую |
| `multi_match_one_cell` | Все совпадения в одну ячейку через перевод строки |
| `not_found_fill` | Значение при отсутствии ключа; `_ПУСТО_` / `_EMPTY_` — пустая ячейка |
| `trim_keys` | Обрезать пробелы у ключей перед сравнением |
| `column_suffix` | Суффикс к именам новых колонок с правой (по умолчанию `true`) |

В поддиалоге ВПР кнопка **«Справка»** показывает полное описание полей с примерами.

---

## Общие поля блока

| Поле | Тип | Описание |
|------|-----|----------|
| `v` | int | Версия кодека (сейчас `1`) |
| `fn` | string | Ключ функции (как в колонке B) |
| `sheet` | string | Один лист (имя, индекс, шаблон) |
| `sheets` | string[] | Список листов (для `тонкая_сетка`, `авто_ширина`, …) |

---

## Список листов (`sheets`)

Функции: `тонкая_сетка`, `толстая_сетка`, `авто_высота`, `авто_ширина`, `левое_выравнивание`, `сброс_стрипов`, `закрепить_заголовок`, `автофильтр`, `вертикаль_центр`, `только_значения`, `формат_деньги`, `серые_служебные`, `полосы_по_пути`, `чередующиеся_границы`, `копировать_формат_заголовка`, `подсветка_по_заголовку`, `жирный_по_пути`, …

```json
[{"v":1,"fn":"тонкая_сетка"}]
[{"v":1,"fn":"толстая_сетка","sheets":["Отчет_январь"]}]
[{"v":1,"fn":"авто_ширина","sheets":["Отчет_январь","Сводка"]}]
```

Пустой массив `sheets` или отсутствие поля — все листы.

---

## Блоки по листу (`sheet` + поля)

### `заголовок_плюс_высота`

```json
[{"v":1,"fn":"заголовок_плюс_высота","height_mm":13,"h_align":"center","v_align":"center","bold":true}]
```

Дополнительно: `font_name`, `font_size`, `fill_color`, `font_color`, `font_color_auto`.

### `отступ`

```json
[{"v":1,"fn":"отступ","h_align":"left","steps":1,"columns":["Сумма_руб"]}]
```

### `шрифт`

```json
[{"v":1,"fn":"шрифт","name":"PT Sans","size_data":10,"size_header":11}]
```

### `ширина_столбцов` / `высота_строки`

```json
[{"v":1,"fn":"ширина_столбцов","width_mm":30,"columns":"all"}]
[{"v":1,"fn":"высота_строки","height_mm":5,"rows":"all"}]
```

`columns` / `rows`: `"all"` или массив номеров/имён.

### `сетка`

```json
[{"v":1,"fn":"сетка","width":"тонкая","color":"серый","include_header":true}]
```

### `зебра_диапазон`

```json
[{"v":1,"fn":"зебра_диапазон","roles":{"header":{"fill":"excel_header","font":"белый"},"even":{"fill":"голубой","font":"синий"},"odd":{"fill":"белый","font":"синий"}}}]
```

### `перенос` / `перенос_и_авто_высота`

```json
[{"v":1,"fn":"перенос","wrap":true}]
[{"v":1,"fn":"перенос_и_авто_высота","wrap":true,"include_header":true}]
```

### `формат_даты`

```json
[{"v":1,"fn":"формат_даты","format":"DD.MM.YYYY","markers":["дата","время"],"convert_existing":true,"parse_formats":"FIRST_DAY.MMMM.YEAR"}]
```

`parse_formats` — разбор текста при `convert_existing` (тот же синтаксис, что `MERGE_XML_DATETIME_FORMATS`, включая month-only). Пусто — bound merge-флаг, если был.

### `формат_столбцы`

```json
[{"v":1,"fn":"формат_столбцы","rules":[{"column":"Сумма_руб","format":"# ##0,00"}],"convert_existing":false}]
[{"v":1,"fn":"формат_столбцы","rules":[{"column":"'Месяц'","format":"DD.MM.YYYY"}],"convert_existing":true,"parse_formats":"LAST_DAY.MMMM.NEXT_YEAR"}]
```

### `подсветка_по_порогу` (алиас `условное_форматирование`)

```json
[{"v":1,"fn":"подсветка_по_порогу","marker":"Сумма_руб","threshold":70000,"color":"светло_красный"}]
```

### `градиент`

```json
[{"v":1,"fn":"градиент","marker":"Сумма_руб","color_min":"зеленый","color_max":"красный","whole_row":true,"sort":true,"sort_dir":"по_возр"}]
```

### `сортировка`

```json
[{"v":1,"fn":"сортировка","keys":[{"column":"ФИО","desc":false},{"column":"Сумма_руб","desc":true}]}]
```

С фильтром по листу:

```json
[{"v":1,"fn":"сортировка","sheet":"Отчет_январь","keys":[{"column":"ФИО","desc":false}]}]
```

### `раскрасить_блоки` (алиас `раскрасить`)

```json
[{"v":1,"fn":"раскрасить_блоки","key_columns":["ФИО"],"colors":["голубой","желтый","зеленый"],"sort":"asc","outline_blocks":true,"outline_color":"темно_синий"}]
```

### `удалить_столбцы`

```json
[{"v":1,"fn":"удалить_столбцы","markers":["маркер","ДатаВремя_строка"]}]
```

### `конкатенация_столбцов`

```json
[{"v":1,"fn":"конкатенация_столбцов","new_column":"ФИО_Отдел","separator":" / ","columns":[1,3]}]
```

### `применить_формулу`

```json
[{"v":1,"fn":"применить_формулу","column":"Бонус_итог","formula":"=D{row_id}*0.1","as_values":false}]
```

Нумерация строк (эквивалентные шаблоны):

```json
[{"v":1,"fn":"применить_формулу","column":"№ п/п","formula":"=IFERROR({col_id}{row_id-1}+1;1)","as_values":false}]
```

```json
[{"v":1,"fn":"применить_формулу","column":"№ п/п","formula":"=IFERROR(R[-1]C+1;1)","as_values":false}]
```

Плейсхолдеры: `{row_id}`, `{row_id-1}`, `{col_id}` → A1; `R[-1]C` — R1C1 (ячейка выше в том же столбце).

Разделитель аргументов в шаблоне — **`;`** (как в русском Calc). Макрос пишет через **FormulaLocal**; при R1C1-документе `{col_id}{row_id-1}` автоматически превращается в `R[-1]C`.

### `удаление_строк` (алиас `пропуск_пустых_строк`)

По столбцам:

```json
[{"v":1,"fn":"удаление_строк","mode":"columns","columns":["маркер"]}]
```

По формуле:

```json
[{"v":1,"fn":"удаление_строк","mode":"formula","formula":"E{row_id}+F{row_id}=0"}]
```

### `группировка_по_столбцу`

Один уровень итога:

```json
[{"v":1,"fn":"группировка_по_столбцу","marker":"'Отдел'","agg_fn":"sum","agg_columns":["'Сумма_руб'"],"sort":true,"outline":true,"sheet":"МСУЗ"}]
```

Несколько уровней в одном блоке (`aggs`):

```json
[{"v":1,"fn":"группировка_по_столбцу","marker":"'Отдел'","sheet":"МСУЗ",
  "aggs":[{"fn":"count","columns":["'Сумма_руб'"]},{"fn":"sum","columns":["'Сумма_руб'"]}],
  "sort":true,"outline":true}]
```

| Поле | Описание |
|------|----------|
| `marker` | Столбец группировки (`'…'` = точное имя заголовка) |
| `agg_fn` / `agg_columns` | Один уровень: функция и столбцы формул `ИТОГ` |
| `aggs` | Несколько уровней: `[{"fn":"count","columns":[…]},{"fn":"sum","columns":[…]}]` |
| `sort` / `outline` / `grand_total` | Сортировка, outline, общие итоги (по умолчанию `true`) |

**Склейка:** подряд идущие строки параметров с одним `marker`+`sheet` и разными `agg_fn` склеиваются в один проход (как `aggs[]`). Иначе второй проход портит итоги. Эквивалент — один блок с `aggs`. Подробнее: [05_POSTPROCESS.md](05_POSTPROCESS.md) §22.

### `стиль_печати`

```json
[{"v":1,"fn":"стиль_печати","orientation":"landscape","fit_pages":1}]
```

### `заполнение_вниз` / `заполнение_вниз_вычислить`

```json
[{"v":1,"fn":"заполнение_вниз","columns":["Отдел"]}]
[{"v":1,"fn":"заполнение_вниз_вычислить","columns":["A"],"rules":[{"column":"A","formula":"=A{row-1}"}]}]
[{"v":1,"fn":"заполнение_вниз_вычислить","columns":[1],"sheet":"Сводная","rules":[{"column":"A","formula":"=R[-1]C","expand_to_right":true}]}]
```

| Поле (блок) | Описание |
|-------------|----------|
| `columns` | Столбцы, в которых ищутся пустые ячейки (номер с 1, буква, имя заголовка) |
| `rules` | Правила: `column`, `formula`, опционально `expand_to_right` |
| `formula` | Общая формула для всех `columns`, если для столбца нет своего rule |

| Поле (rule) | Описание |
|-------------|----------|
| `column` | Столбец правила (сопоставляется с `columns` по индексу) |
| `formula` | Формула Calc или константа; `{row_id}`, `{row-1}`, `R[-1]C` |
| `expand_to_right` | После заполнения столбца — протянуть формулу в пустые ячейки строки вправо (до `end_col` диапазона) |

### `копировать_значения`

```json
[{"v":1,"fn":"копировать_значения","columns":["Сумма_руб","Сумма_бонус"]}]
```

Весь лист:

```json
[{"v":1,"fn":"копировать_значения","whole_sheet":true}]
```

### `замена_значений`

Заменяет подстроки в **текстовых** столбцах в памяти (`getDataArray` → Python → один `setDataArray`).

```json
[{"v":1,"fn":"замена_значений",
  "columns":["ФИО","Комментарий"],
  "find":["ООО","  "],
  "replace":["ОБЩЕСТВО"," "],
  "case_insensitive":true,
  "squeeze_spaces":true,
  "row_filter":"lambda rows: rows(\"Статус\") != \"архив\"",
  "match_pick":"first"
}]
```

| Поле | Описание |
|------|----------|
| `columns` | Столбцы (номер с 1, буква, имя заголовка) |
| `find` | Список подстрок поиска |
| `replace` | Либо 1 значение (для всех `find`), либо список той же длины; пусто / `_ПУСТО_` — стирание; токены с `,`/`;` — в `'…'` |
| `case_insensitive` | Искать без учёта регистра (по умолчанию `false`) |
| `squeeze_spaces` | `trim` + схлопнуть пробельные символы в один пробел (по умолчанию `false`) |
| `<<Переменные.…>>` | В `find`/`replace`, лямбдах, **путях/масках «Файлы-Источники»**: подстановка из карты; в `'…'` (find/replace) — литерал без подстановки |
| `row_filter` | Опционально: `lambda rows: …` — в каких строках искать. Sugar как у `столбец_по_лямбде`: `rows("Col")`, `rows("Col")[-1]`, `rows(1)` / `rows("A")` |
| `match_pick` | Если в столбце несколько совпадений: `all` (по умолчанию), `first`, `last`, `odd`, `even` (1-based среди совпавших **строк**), `lambda` |
| `match_expr` | При `match_pick=lambda`: `lambda rows: …`; в namespace также `idx` (0-based), `n`, `i` (1-based), `text` |
| `sheet` | Лист результата; пусто — все листы |

Лямбды `row_filter` и `match_expr` понимают sugar:

- `rows("Столбец")` — значение в текущей строке;
- `rows("Столбец")[-1]` — значение в предыдущей строке листа;
- `upd_rows("Столбец")` — значение в последней **уже обработанной** строке выше текущей;
- `upd_rows_seen("Столбец", value)` — было ли такое значение выше;
- `upd_rows_count("Столбец", value)` / `upd_rows_len()` — счётчики по уже обработанным строкам.

Примеры:

```json
[{"v":1,"fn":"замена_значений","columns":["ИНН"],"find":["-"],"replace":[""],"row_filter":"lambda rows: not upd_rows_seen(\"ИНН\", rows(\"ИНН\"))"}]
[{"v":1,"fn":"замена_значений","columns":["Комментарий"],"find":["ООО"],"replace":["Общество"],"row_filter":"lambda rows: rows(\"Статус\") != rows(\"Статус\")[-1]"}]
[{"v":1,"fn":"замена_значений","columns":["Сумма"],"find":["."],"replace":[","],"match_pick":"lambda","match_expr":"lambda rows: idx == 0 or text != upd_rows(\"Сумма\")"}]
```

### `нормализация_текста`

Стек текстовых преобразований (`ops[]`) по столбцам. RANGE / FINAL / `Постобработка_xml`.

```json
[{"v":1,"fn":"нормализация_текста","sheet":"Сводная","columns":["ФИО"],
  "ops":[
    {"op":"collapse_ws"},
    {"op":"lang_homoglyphs"},
    {"op":"case","mode":"proper"}
  ]}]
```

С фильтрами по строкам у отдельных операций:

```json
[{"v":1,"fn":"нормализация_текста","columns":["ФИО"],
  "ops":[
    {"op":"collapse_ws","row_filter":"non_empty"},
    {"op":"case","mode":"proper","row_filter":"lambda","row_filter_lambda":"lambda rows: not upd_rows_seen(\"ФИО\", rows(\"ФИО\"))"}
  ]}]
```

| Поле | Описание |
|------|----------|
| `columns` | Столбцы (**обязательно**) |
| `ops` | Упорядоченный стек: `trim`, `collapse_ws`, `case`(+`mode`), `lang_homoglyphs`, `reverse`, `base64_encode`/`decode`, `sha256`, `encrypt`/`decrypt` |
| `non_text` | `skip` (default) \| `coerce` \| `error` |
| `on_error` | `keep` (default) \| `empty` \| `stop` |
| `key` / `key_cell` / `key_file` | ключ для encrypt/decrypt (взаимоисключающие) |
| `sheet` | Лист результата; пусто — все |

Для op с `row_filter="lambda"` поле `row_filter_lambda` поддерживает те же helper'ы, что и `замена_значений`:

- `rows(...)` — текущая/соседняя строка листа;
- `upd_rows(...)` — уже обработанные строки выше;
- `upd_rows_seen(...)`, `upd_rows_count(...)`, `upd_rows_len()` — быстрый поиск по накопителю.

Алиасы `fn`: `нормализовать_текст`, `normalize_text`.

### `переименовать_лист`

```json
[{"v":1,"fn":"переименовать_лист","new_name":"Сводка_тест"}]
```

### `переименовать_столбцы`

Переименование заголовков в строке заголовков (не ячеек данных).

```json
[{"v":1,"fn":"переименовать_столбцы","sheet":"Заказы","mappings":[{"old":"Сумма_руб","new":"Сумма"}]}]
```

Текстовый вид соответствий в визарде: `Старое=Новое, Другое=Ещё`. Старое имя должно однозначно указывать столбец.

| Поле | Описание |
|------|----------|
| `sheet` | Лист результата; пусто — все листы (как у других sheet-block функций) |
| `mappings` | Список `{old, new}` или строка `A=B, C=D` |

### `переставить_столбцы`

Физический перенос или **копирование** группы столбцов. **`sheet` обязателен** (режим «все листы» не применяется).

```json
[{"v":1,"fn":"переставить_столбцы","sheet":"Заказы","columns":["C","Сумма"],"position":"_перед_","relative_column":"Итого"}]
```

Копия столбцов (исходные остаются, вставляются дубликаты):

```json
[{"v":1,"fn":"переставить_столбцы","sheet":"Заказы","columns":["Сумма"],"position":"_после_","relative_column":"ФИО","copy":true,"new_names":["Сумма_копия"]}]
```

В режиме `copy` столбцы **вставляются** (`insertByIndex`), затем копируются значения (буфер → SVD), форматы (буфер → Flags=T) и ширина; диапазон листа для следующих шагов расширяется. Перенос (`copy` нет) по-прежнему через перестановку DataArray.

| Поле | Описание |
|------|----------|
| `sheet` | Конкретный лист (обязателен) |
| `columns` | Столбцы для переноса/копии: номер (1-based), буква, имя или шаблон с `*` |
| `position` | `_начало_`, `_конец_`, `_перед_`, `_после_` (или `start`/`end`/`before`/`after`) |
| `relative_column` | Опорный столбец для `_перед_`/`_после_` — должен быть **единственным**; при переносе если он в `columns`, исключается из переноса (при копировании — нет) |
| `copy` | `true` — копировать (не перемещать); алиас `mode`: `копировать` / `copy` |
| `new_names` | Новые заголовки копий (параллельно `columns`); пусто или короче списка — автоимя `Исходное_2`, … |

### `количество_значений`

Столбец с числом вхождений составного ключа в диапазоне данных (один проход — карта частот).

```json
[{"v":1,"fn":"количество_значений","new_column":"Количество","key_columns":["ФИО","Отдел"],"trim":true,"case_sensitive":false}]
```

| Поле | Описание |
|------|----------|
| `new_column` | Имя нового столбца; пусто — `_Количество_значений` (с суффиксом `_2`, … при конфликте) |
| `key_columns` | Столбцы ключа: номер, буква или имя заголовка |
| `trim` | Обрезка пробелов при сборке ключа (по умолчанию `true`) |
| `case_sensitive` | Учёт регистра при сравнении (по умолчанию `false`) |
| `sheet` | Лист результата; пусто — все листы |

Несколько листов с разными параметрами — несколько объектов в массиве.

### `удалить_дубликаты`

Удаление повторяющихся строк по ключу (аналог Excel «Удалить дубликаты» по выбранным столбцам).

```json
[{"v":1,"fn":"удалить_дубликаты","key_columns":["ФИО","Отдел"],"keep":"first","pre_sort":[{"column":"Дата","order":"desc"}],"output":"inplace"}]
```

| Поле | Описание |
|------|----------|
| `key_columns` | Столбцы ключа; пусто — все столбцы `data_range` |
| `exclude_columns` | Все столбцы, кроме перечисленных (взаимоисключимо с непустым `key_columns`) |
| `keep` | `first` (по умолч.) или `last` — какое вхождение ключа оставить |
| `scan_order` | `top_down` / `bottom_up` (опционально; по умолчанию по `keep`) |
| `pre_sort` | Сортировка до dedup: `[{"column":"…","order":"asc"|"desc"}]` |
| `trim`, `case_sensitive`, `compare_as`, `empty_key_policy`, `as_values` | Как у ключа в `количество_значений` + `compare_as`: `text`/`strict` |
| `output` | `inplace` (по умолч.), `new_sheet`, `offset` |
| `dest_sheet` | Для `new_sheet` |
| `dest_cell` | Якорь для `offset` (по умолч. `A1`) |
| `overwrite_overlap` | Разрешить пересечение при `offset` |
| `mark_duplicates` | Не удалять; столбец-маркер для повторов |
| `mark_column` | Имя столбца маркера (по умолч. `Дубликаты`) |
| `sheet` | Лист результата; пусто — все листы |

Алиасы fn: `убрать_повторы_строк`, `dedup`.

### `копировать_лист`

Полная копия вкладки в той же книге (`Sheets.copyByName`).

```json
[{"v":1,"fn":"копировать_лист","source_sheet":"Шаблон","dest_sheet":"Копия_Шаблон"}]
```

| Поле | Описание |
|------|----------|
| `source_sheet` | Имя листа-источника |
| `dest_sheet` | Имя нового листа; при конфликте — суффикс `_2`, `_3`, … |
| `sheet` | На каком листе результата выполнять шаг (фильтр строки таблицы); для копирования обычно не важен |
| `columns` | **Опционально (ODS в `xml_excel_ods`)**: список столбцов (имя/индекс/буква/шаблон) для копирования |
| `skip_empty` | **Опционально (ODS в `xml_excel_ods`)**: `true` — переносить строку только если среди выбранных столбцов есть непустая |
| `skip_empty_columns` | **Опционально (ODS в `xml_excel_ods`)**: по каким колонкам (в *итоговом* наборе `columns`) проверять пустоту строки; пусто → по первой колонке |
| `header_row` | **Опционально (ODS в `xml_excel_ods`)**: строка заголовков (1-based, по умолчанию 1) — для `columns` |

Доступна в постобработке диапазона и в финальной обработке.

---

## Отдельные параметры pipeline (A ≠ Постобработка_*)

### `объединить_листы_в_один`

Параметр в колонке **A** (не ключ B «Постобработка_Диапазон»). JSON только в C.

```json
[{"v":1,"fn":"объединить_листы_в_один","sheets":["Отчет_Q1","Отчет_Q2","Отчет_Q3"],"dest_sheet":"Отчет_Q123","header_row":1,"with_formatting":false}]
```

| Поле | Описание |
|------|----------|
| `sheets` | Листы-источники в книге результата |
| `dest_sheet` | Лист-приёмник (создаётся или очищается) |
| `header_row` | Строка заголовков на листах результата (**1-based**); пусто — авто |
| `with_formatting` | `true` — перенос оформления столбцов (буфер); иначе формат чисел + дефолтный шрифт данных |

В pipeline: `kind=merge_sheets` (один раз на шаг; dest → очередь оформления как после `копировать_лист`).

### `разделить_листы`

Отдельная строка A; JSON в C. Pipeline: `kind=split_sheets` (`libre_macros_split_lib.lm_split_sheets_run`).

```json
[{
  "v": 1,
  "fn": "разделить_листы",
  "run_phase": "after_postprocess",
  "source": {"book": "эта_книга", "sheet": "Общий"},
  "split_books": ["Организация"],
  "split_sheets": ["Подразделение", "Участок"],
  "header_row": 1,
  "pre_sort": true,
  "book_template": "эта_книга",
  "sheet_template": "[Подразделение]-[Участок]",
  "split_mode": "copy_range",
  "paste_formats": false,
  "copy_header": true,
  "create_folders": true,
  "overwrite_sheets": true,
  "append_if_exists": false,
  "save_immediately": false
}]
```

| Поле | Смысл |
|------|--------|
| `run_phase` | `after_collect` / `after_postprocess` (default) / `after_final` |
| `source.book` | Всегда `эта_книга` |
| `source.sheet` | Лист-источник в книге результата |
| `split_books` / `split_sheets` | Поля-триггеры разбиения по книгам/листам |
| `book_template` / `sheet_template` | Шаблоны имён с `[Поле]`; `эта_книга` — только листы в текущей книге |
| `split_mode` | `copy_range` (и др. по реализации) |
| `header_row` | 1-based строка заголовков на источнике |

При `xml_excel_ods` допустим только `run_phase=after_final`. Подробнее — [15_PARAM_REFERENCE.md](15_PARAM_REFERENCE.md).

---

## Параметр сбора и RANGE: `строка_заголовков`

Используется в строке **`Строка_Заголовков`**: B = `смотри в C`, JSON в C (или как `fn` в `Постобработка_Диапазон` при flatten на копии листа).

```json
[{"v":1,"fn":"строка_заголовков","rows":"2-5","separator":"пробел"}]
```

| Поле | Смысл |
|------|--------|
| `rows` | Строки шапки 1-based: `2-5`, `2:5`, `3` |
| `separator` | Между уровнями: `пробел`, `_`, `-`, `/`, пусто |

Объединения при сборке: горизонтально — значение во все столбцы строки; вертикально — только верхняя строка.

---

## RANGE: `разделить_по_столбцам`

```json
[{"v":1,"fn":"разделить_по_столбцам","column":"'ФИО'","delimiter":",","position":"_после_","source_policy":"replace"}]
```

| Поле | Смысл |
|------|--------|
| `column` | Исходный столбец (`'…'` для exact-match) |
| `delimiter` | Разделитель текста |
| `position` | `_начало_`, `_конец_`, `_перед_`, `_после_` + `relative_column` |
| `source_policy` | `replace` — первая часть в исходном; `keep` — исходный без изменений |
| `split_mode` | `each` / `leftmost`; `max_columns`, `new_names`, `trim_parts`, `non_text` |

---

## RANGE: `условный_столбец`

```json
[{"v":1,"fn":"условный_столбец","new_column":"Категория","rules":[{"column":"'Сумма'","op":"ge","value":1000,"then":"Крупный"}],"otherwise":"Мелкий"}]
```

| Поле | Смысл |
|------|--------|
| `rules[]` | `column`, `op` (`eq`, `ne`, `gt`, `ge`, `lt`, `le`, `contains`, …), `value`, `then` |
| `otherwise` | Значение ELSE |
| `compare_as` | `auto`, `text`, `number`, `date`; `trim`, `case_sensitive` |

---

## RANGE: `столбец_по_лямбде`

```json
[{"v":1,"fn":"столбец_по_лямбде","new_column":"ИНН_копия","expr":"lambda rows: rows('ИНН')","position":"_конец_","on_error":"keep","result_type":"auto","output_policy":"new"}]
```

Sugar в `expr`: `rows("Col")`, `rows("Col")[-1]`, `upd_rows("Col")`, `upd_rows_len()`, … Санация через `libre_macros_sanitize_lib.py`.

---

## RANGE / Финал: `добавить_столбец`

Алиасы: `add_column`, `новый_столбец`.  
Lib: `lm_pp_range_add_column` / `lm_final_add_column` (шимы `merge_*`).  
Только UNO (RANGE / Финал); **не** в `Постобработка_xml`.

```json
[
  {"v":1,"fn":"добавить_столбец","new_column":"Статус","data_type":"text","choices":["Активен","Архив"]},
  {"v":1,"fn":"добавить_столбец","sheet":"Сводка","new_column":"Сумма","data_type":"number",
   "format":"# ##0,00","data_validation":true,"min":"0","max":"1000000"},
  {"v":1,"fn":"добавить_столбец","new_column":"Дата_отчета","data_type":"date",
   "format":"DD.MM.YYYY","data_validation":true,"min":"01.01.2020","max":"31.12.2030"}
]
```

| Поле | Тип | Описание |
|------|-----|----------|
| `new_column` | string | Имя столбца |
| `data_type` | string | `text` \| `number` \| `date` |
| `format` | string | NumberFormat; пусто — по типу (`@` / глобальный числовой / дата) |
| `choices` | list | Для `text`: допустимые значения → LIST-валидация на диапазоне данных |
| `data_validation` | bool | Для `number`/`date`: контроль мин/макс |
| `min` / `max` | string | Границы числа или даты |
| `sheet` | string | Лист; пусто — все листы шага |

Столбец в конец used-области; при существующем заголовке — только формат и контроль.

---

## RANGE: `текстовые_операции`

Алиас legacy: `нормализация_текста`.

Стек `ops[]` по `columns`. Полный каталог — кнопка «Справка» в визарде / `libre_macros_wizard_help_cfg.py`.

| Поле | Default | Смысл |
|------|---------|--------|
| `columns` | `[]` | Столбцы (`'Имя'` = exact) |
| `ops` | `[]` | Список `{op, …}`; см. docs/05 |
| `non_text` | `skip` | `skip` \| `coerce` \| `error` |
| `on_error` | `keep` | `keep` \| `empty` \| `stop` |
| `key` / `key_cell` / `key_file` | — | Ровно один канал для encrypt/decrypt |

```json
[{"v":1,"fn":"текстовые_операции","columns":["'ФИО'"],
  "ops":[{"op":"collapse_ws"},{"op":"case","mode":"proper"}],
  "non_text":"skip","on_error":"keep"}]
```

---

## Финал: `активировать_лист`

```json
[{"v":1,"fn":"активировать_лист","target_sheet":"Merge","tab_color":"голубой"}]
```

`tab_color` — опционально: цвет ярлычка (имя из пресетов / `#RRGGBB`); `нет` / `сброс` — системный цвет; пусто — не менять.

---

## Финал / RANGE: `копирование_диапазонов`

Копирует прямоугольник / строки / столбцы с листа-источника на один или несколько листов-приёмников **той же книги результата** (не внешние файлы).

Алиасы: `копировать_диапазон`, `copy_ranges`, `copy_range`, `вставить_диапазон`.  
Lib: `lm_final_copy_ranges` / `merge_final_copy_ranges` (+ RANGE: `lm_pp_range_copy_ranges`).

Не путать с `копировать_значения` (формулы→значения на том же диапазоне), `копировать_переместить_лист` (целая вкладка), `разделить_листы` / `split_mode=copy_range`.

| Поле | Default | Смысл |
|------|---------|--------|
| `source_sheet` | обязателен | Лист-источник (exact; без glob) |
| `source_range` | `""` | A1: `A1:D20`, одна ячейка `A1`, столбец `C:C`, строка `5:5` |
| `source_rows` | `""` | Альтернатива: строки 1-based (`"2:10"`, `"2,5,7"`, `[2,3,4]`) — непрерывный блок |
| `source_columns` | `""` | Альтернатива: столбцы (`"A:C"`, `"B,D"`, имена в `'…'`) — непрерывный блок |
| `dest_sheet` | `""` | Один лист-приёмник |
| `dest_sheets` | `[]` | Несколько приёмников (приоритетнее `dest_sheet`) |
| `dest_cell` | `A1` | Якорь левого верхнего угла (одна ячейка A1, не диапазон) |
| `mode` | `replace` | `replace` — поверх; `insert` — вставка строк/столбцов перед якорем |
| `insert_axis` | `auto` | Только для `insert`: `rows` \| `cols` \| `auto` |
| `content` | `values` | `values` — вычисленные; `formulas` — сами формулы (clipboard + автосдвиг ссылок) |
| `with_formatting` | `false` | Копировать оформление (clipboard; best effort) |
| `clear_source` | `false` | После успешных вставок очистить исходный прямоугольник |
| `create_missing_dest` | `false` | Создать лист-приёмник, если нет |
| `involve_dest` | `true` | Вовлечь приёмник(и) в скоп следующих шагов pipeline |
| `header_row` | из pipeline / `1` | 1-based строка заголовков на приёмнике (для регистрации скопа) |

Приоритет источника: `source_range` → иначе `source_rows`/`source_columns` (оба = пересечение; один = полоса по used-area) → иначе ошибка.

Алиасы полей при нормализации: `source`→`source_range`, `from_sheet`→`source_sheet`, `to_sheet`/`target_sheet`→`dest_sheet`, `to_sheets`/`target_sheets`→`dest_sheets`, `anchor`/`at`→`dest_cell`, `paste_mode`→`mode`.

Минимум — замена блока на другом листе:

```json
[{"v":1,"fn":"копирование_диапазонов","source_sheet":"Сводная","source_range":"A1:C10","dest_sheet":"Отчет","dest_cell":"A1","mode":"replace"}]
```

Строки источника → вставка со сдвигом строк:

```json
[{
  "v":1,
  "fn":"копирование_диапазонов",
  "source_sheet":"Данные",
  "source_rows":"2:50",
  "dest_sheet":"Сбор",
  "dest_cell":"A2",
  "mode":"insert",
  "insert_axis":"rows"
}]
```

Столбцы по заголовкам → несколько листов:

```json
[{
  "v":1,
  "fn":"копирование_диапазонов",
  "source_sheet":"Справочник",
  "source_columns":["'Код'","'Название'"],
  "dest_sheets":["Лист1","Лист2"],
  "dest_cell":"F1",
  "mode":"replace"
}]
```

Целый столбец A1 + очистка источника:

```json
[{
  "v":1,
  "fn":"копирование_диапазонов",
  "source_sheet":"Temp",
  "source_range":"B:B",
  "dest_sheet":"Итог",
  "dest_cell":"D1",
  "mode":"replace",
  "clear_source":true
}]
```

«Заморозить» формулы в значения на другом листе:

```json
[{
  "v":1,
  "fn":"копирование_диапазонов",
  "source_sheet":"Расчет",
  "source_range":"E2:E100",
  "dest_sheet":"Отчет",
  "dest_cell":"C2",
  "content":"values",
  "mode":"replace"
}]
```

Перенести формулы (с автосдвигом ссылок Calc):

```json
[{
  "v":1,
  "fn":"копирование_диапазонов",
  "source_sheet":"Шаблон",
  "source_range":"A1:D20",
  "dest_sheet":"Рабочий",
  "dest_cell":"A1",
  "content":"formulas",
  "mode":"replace"
}]
```

Полный набор полей (пример):

```json
[{
  "v":1,
  "fn":"копирование_диапазонов",
  "source_sheet":"Сводная",
  "source_range":"A1:D20",
  "dest_sheets":["Отчет","Архив"],
  "dest_cell":"B3",
  "mode":"replace",
  "content":"values",
  "with_formatting":false,
  "clear_source":false,
  "create_missing_dest":false,
  "involve_dest":true,
  "header_row":1
}]
```

Ограничения v1: при `source_sheet == dest_sheet` и пересечении диапазонов + `clear_source=true` — ошибка; cut / мульти-диапазоны / внешние книги — вне scope. Подробнее — [05_POSTPROCESS.md](05_POSTPROCESS.md) (`копирование_диапазонов`).

---

## RANGE / Финал: `транспонировать_таблицу`

Алиасы: `transpose`, `transpose_table`. Только UNO (не `Постобработка_xml` в v1). Ядро: `libre_macros_transpose_lib.py`.

```json
[{"v":1,"fn":"транспонировать_таблицу","sheet":"Матрица","output":"new_sheet","dest_sheet":"Матрица_T","as_values":true}]
```

| Поле | Тип | Описание |
|------|-----|----------|
| `output` | string | `inplace` \| `new_sheet` \| `offset` |
| `dest_sheet` / `dest` | string | обязателен при `new_sheet` |
| `dest_cell` | string | якорь A1 при `offset` |
| `range` | string | A1-диапазон; пусто = авто |
| `header_row` / `data_start` | int | якоря авто-диапазона |
| `columns` | list | `columns_pick` |
| `headers_from_column` | bool | шапка из колонки меток |
| `header_column` | int/string | колонка меток |
| `result_headers` | list/string | кастомные имена столбцов результата (prepend); алиас `new_headers` |
| `skip_header_row` / `skip_first_column` | bool | вырез до транспонирования (ограничения с `inplace`) |
| `as_values` | bool | default true |
| `overwrite_overlap` / `clear_source` / `with_formatting` | bool | см. справку визарда |

```json
[{"v":1,"fn":"транспонировать_таблицу","sheet":"Метки_колонка","output":"new_sheet","dest_sheet":"Метки_T","headers_from_column":true}]
```

```json
[{"v":1,"fn":"транспонировать_таблицу","sheet":"Матрица","output":"new_sheet","dest_sheet":"Матрица_H","result_headers":["Метка","Строка1","Строка2"]}]
```

Сценарий: `16_transpose_table.xltx`. Подробности — [05_POSTPROCESS.md](05_POSTPROCESS.md), встроенная «Справка» визарда.

---

## RANGE / Финал: `анкета_в_таблицу` / `таблица_в_анкету`

Алиасы: `form_to_table` / `form_to_wide`; `table_to_form` / `wide_to_form`. Ядро: `libre_macros_form_table_lib.py`.

```json
[{"v":1,"fn":"анкета_в_таблицу","question_column":"A","answer_column":"B","block_mode":"by_repeat_key","block_start_question":"ФИО","output":"new_sheet","dest_sheet":"Анкеты_wide"}]
```

```json
[{"v":1,"fn":"таблица_в_анкету","body_columns":["'ФИО'","'Возраст'"],"question_header":"Вопрос","answer_header":"Ответ","has_header_in_output":true,"dest_sheet":"Анкеты"}]
```

Сценарии: `14_form_to_table.xltx`, `15_table_to_form.xltx`.

---

## ROW-функции

### `преобразовать_числа`

```json
[{"v":1,"fn":"преобразовать_числа","columns":["Сумма_руб","Сумма_бонус"]}]
```

### `цвет_текста_по_значению`

```json
[{"v":1,"fn":"цвет_текста_по_значению","marker":"статус"}]
```

Остальные ROW (`серые_служебные`, `полосы_по_пути`, …) — как список листов, см. раздел «Список листов».

---

## Сводная таблица

Отдельный объект (не массив блоков с `fn` в общем списке):

```json
{"v":1,"fn":"сводная_таблица","source_sheet":"Отчет_январь","header_row":2,"data_start":3,"sheet_name":"Сводная","as_values":true,"row_fields":["Отдел"],"column_fields":[],"data_fields":[{"field":"Сумма_руб","function":"SUM"}]}
```

Псевдофункция **Список_строк** (`LIST_ROWS`) — детализация строк источника в data-ячейке (после `as_values`):

```json
{"v":1,"fn":"сводная_таблица","source_sheet":"Merge","sheet_name":"Сводная","as_values":true,"list_rows_expand":true,"row_fields":["Филиал"],"column_fields":[],"filter_fields":[],"data_fields":[{"field":"ФИО","function":"LIST_ROWS","list_fields":["ФИО","Должность","Сумма_руб"]}]}
```

| Поле | Описание |
|------|----------|
| `data_fields[].function` | UNO-агрегат (`SUM`, `COUNT`, …) или `LIST_ROWS` |
| `data_fields[].list_fields` | только для `LIST_ROWS`: колонки детализации |
| `as_values` | при `LIST_ROWS` всегда `true` |
| `list_rows_expand` | `true` — каждая запись списка на своей строке; `false`/нет — всё в одной ячейке через `\n` |
---

## `функция_плагин`

Пользовательский код из `functions_pp.py` / `functions_final.py` или inline Python:

```json
[{"v":1,"fn":"функция_плагин","ref":"functions_pp.py#pp_range_zagotovka","extra":"42,тест","sheet":"Сводка"}]
```

| Поле | Описание |
|------|----------|
| `ref` | `functions_*.py#имя_функции` |
| `code` | inline `lambda` / `def` (вместо `ref`) |
| `extra` | доп. аргументы через запятую (аналог колонки D) |
| `sheet` | один лист результата; пусто — все листы |
| `sheets` | список листов (как у встроенных функций) |

Несколько листов с разными параметрами — несколько объектов в массиве (как у «зебра_диапазон»):

```json
[
  {"v":1,"fn":"функция_плагин","sheet":"1_Заказы","ref":"functions_pp.py#pp_range_zagotovka","extra":"4,5"},
  {"v":1,"fn":"функция_плагин","sheet":"Сводка","ref":"functions_pp.py#pp_stroka_itogo"}
]
```

Галочка «Одинаково на всех листах» в визарде — один блок без `sheet`.

---

## Визард и пресеты

В `param_wizard` диалог «Параметры…» сохраняет JSON автоматически. Пресеты визарда (вкладка «Пресеты») соответствуют сценариям 01–10 в `generate_manual_workbooks.py`.
