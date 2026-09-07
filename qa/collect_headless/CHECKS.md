# Справочник проверок (`checks.py`)

Проверки читают **сохранённый** файл результата (`.ods` / `.ots` / `.xlsx`) без UNO.  
Вызов: `run_checks(path, checks, version_hint=…)`.

Статусы: `pass` | `fail` | `skip`.

Общие поля любого check:

| Поле | Смысл |
|------|--------|
| `type` | Тип проверки (ниже) |
| `gui` | Если `true` — `skip` («gui-check в headless пропущен») |

---

## Журнал `Сбор_книг_лог`

### `log_exists`

Есть лист с префиксом имени `Сбор_книг_лог`.

```json
{"type": "log_exists"}
```

### `log_macro_version`

В тексте журнала есть строка версии. По умолчанию — `macro-lib/version.txt` (передаётся раннером). Можно задать явно:

```json
{"type": "log_macro_version"}
{"type": "log_macro_version", "version": "3.10.567"}
```

Fail, если журнал пуст или версии нет.

### `log_no_status`

В колонке статуса (F) журнала **нет** подстрок из `values` (без учёта регистра).

```json
{"type": "log_no_status", "values": ["ошибка", "error", "critical"]}
```

Не ставить на матричные сценарии вроде `12_postprocess_json`, где отдельные шаги PP ожидаемо пишут «ошибка».

### `log_contains`

В объединённом тексте журнала есть подстрока `text`:

```json
{"type": "log_contains", "text": "режим на один лист"}
```

---

## Листы и строки данных

Служебные листы **не** считаются «данными» при автовыборе листа для `min_data_rows` без `sheet`:

- `collect_params*`, `Параметры_Объединения*`
- `Сбор_книг_лог*`
- `Справка_макроса*`

Подсчёт строк данных: число **непустых** строк начиная со 2-й (строка 1 = заголовок).

### `sheet_exists` / `sheet_absent`

```json
{"type": "sheet_exists", "name": "Merge"}
{"type": "sheet_absent", "name": "Лишний"}
```

### `min_data_rows` / `max_data_rows`

```json
{"type": "min_data_rows", "sheet": "Merge", "n": 10}
{"type": "max_data_rows", "sheet": "Merge", "n": 100}
{"type": "min_data_rows", "n": 1}
```

Без `sheet` / `name` берётся первый неслужебный лист.

### `header_contains`

Заголовки первой строки содержат все токены из `names` (подстрока / casefold):

```json
{"type": "header_contains", "sheet": "Merge", "names": ["ФИО", "Отдел"]}
```

---

## Ячейки

### `cell_equals`

```json
{"type": "cell_equals", "sheet": "Merge", "cell": "A2", "value": "Иванов"}
{"type": "cell_equals", "sheet": "Merge", "cell": "C2", "value": 100, "abs_tol": 0.01}
```

`cell` — A1-адрес. `abs_tol` — сравнение как float (запятая/пробелы в числе допускаются).

### `cell_not_empty`

```json
{"type": "cell_not_empty", "sheet": "Merge", "cell": "B2"}
```

---

## Эталонный лист

### `golden_sheet`

Сравнение прямоугольника результата с эталонным файлом (тот же формат, читается `load_workbook_grid`).

```json
{
  "type": "golden_sheet",
  "path": "test/collect_workbooks/goldens/01_merge.xlsx",
  "sheet": "Merge",
  "range": "A1:J20",
  "abs_tol": 1e-6
}
```

| Поле | Смысл |
|------|--------|
| `path` | Файл эталона |
| `sheet` | Имя листа в обоих файлах |
| `range` | Опционально `A1:D10`; иначе весь используемый прямоугольник |
| `abs_tol` | Допуск для чисел (как у `cell_equals`) |

Fail при первом отличии ячейки (в `detail` — координаты и значения).

---

## Значения по умолчанию

Если у case нет `checks`, раннер использует:

```python
[
  {"type": "log_exists"},
  {"type": "log_macro_version"},
  {"type": "log_no_status", "values": ["ошибка", "error", "critical"]},
  {"type": "min_data_rows", "n": 1},
]
```

---

## Примеры наборов

Строгий smoke (один лист Merge):

```json
[
  {"type": "log_exists"},
  {"type": "log_macro_version"},
  {"type": "log_no_status", "values": ["ошибка", "error", "critical"]},
  {"type": "sheet_exists", "name": "Merge"},
  {"type": "header_contains", "sheet": "Merge", "names": ["ФИО"]},
  {"type": "min_data_rows", "sheet": "Merge", "n": 1},
  {"type": "cell_not_empty", "sheet": "Merge", "cell": "A2"}
]
```

Матрица PP (допускаем ошибки шагов в журнале):

```json
[
  {"type": "log_exists"},
  {"type": "min_data_rows", "n": 1}
]
```

GUI-оформление в манифесте на будущее (будет skip):

```json
{"type": "cell_equals", "sheet": "Merge", "cell": "A1", "value": "x", "gui": true}
```
