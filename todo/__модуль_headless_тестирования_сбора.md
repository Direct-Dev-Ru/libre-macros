# ТЗ: автоматизированный модуль тестирования сбора данных (headless)

**Статус:** реализовано (v1) — `qa/collect_headless/run.py`, макрос `qa_collect_run`. Визард вне scope.

Модуль **не** является pytest-набором в `test/` (политика проекта: автотесты Calc не пишут в `test/` как pytest). Это **раннер сценариев**: открыть книгу с листом `Параметры_Объединения*` в **ACell / LibreOffice headless**, вызвать макрос сбора, сохранить результат, проверить **данные** скриптом и/или вручную.

**Объект:** `collect_workbooks` (и при необходимости `collect_workbooks_for_param_sheet` / `collect_pack` без диалога).  
**Не объект v1:** визард `param_wizard`, file picker, оркестратор с GUI, контекстное меню, тосты.

Связь: [docs/12_MANUAL_TESTING.md](../docs/12_MANUAL_TESTING.md), шаблоны `test/collect_workbooks/workbooks/`, журнал `Сбор_книг_лог`.

---

## 1. Цель

1. Прогонять **готовые шаблоны** (листы параметров + источники) без кликов в GUI.
2. Сохранять книгу-результат в указанный каталог.
3. Проверять **манипуляции данными**: состав листов, заголовки, значения ячеек, число строк, ключи ВПР, JSON-постобработка (dedupe, split, формулы-как-значения, карта переменных).
4. Оформление (градиент, сетка, диалоги, freeze как «видно на экране») — **вне обязательной проверки** в headless; допустим `skip` / WARN.

Критерий успеха v1: один манифест + одна команда → N шаблонов отработали, результаты лежат в `results_dir`, отчёт pass/fail/skip по проверкам данных.

---

## 2. Ограничения среды (GUI-less)

| Что | В headless |
|-----|------------|
| Визард параметров, AWT-диалоги «Запуск» / «Завершить», consent, QueryBox удаления листов | **Не тестировать.** Обходить флагами макроса (ниже). |
| `Отображение_прогресса` = диалог | Зависнуть может. Форсировать `status` / выкл. |
| `param_wizard`, `pick_source_file`, `collect_pack` с окном выбора листов | Вне scope |
| Градиент / зебра / ширина столбца / автофильтр как UI | Не assert’ить в v1 (опционально позже через XML стилей) |
| Сбор, xml/direct, RANGE/ROW/FINAL **по значениям**, ВПР, журнал | **Основной контур** |

Макрос уже умеет тихий прогон (оркестратор):

```python
collect_workbooks_for_param_sheet(
    sheet_name,
    suppress_pre_clear=False,
    skip_consent=True,
    skip_confirm=True,
    quiet_finish=True,
)
```

Раннер **обязан** вызывать сбор через этот путь (или тонкую обёртку `collect_workbooks_headless` в pythonpath), а не через меню с MessageBox.

Дополнительно на время прогона (в settings / флаги раннера, не обязательно писать в шаблон):

- `_MERGE_PACK_SKIP_CONSENT` / `SKIP_CONFIRM` / `QUIET_FINISH` — как у pack;
- `MERGE_PROGRESS_UI_MODE=status` (не `dialog`);
- `MERGE_CONFIRM_DELETION_OF_SHEETS=false`.

Если AWT всё равно требует дисплей (типично для Toolkit даже при `--headless`): поднимать **Xvfb** (Linux) и считать это частью «headless-обвязки», не полноценным GUI-тестом визарда.

---

## 3. Архитектура

```text
qa/collect_headless/          ← раннер (не pytest в test/)
  README.md                   ← подробные доки по тестированию
  CHECKS.md                   ← справочник проверок
  ADDING_CASES.md             ← как добавить сценарий
  run.py / run.sh             ← CLI
  office.py                   ← старт ACell/soffice, UNO-сокет, профиль
  patch_sources.py            ← копия шаблона + подмена путей
  uno_job.py                  ← load doc, invoke qa_collect_run, store
  checks.py                   ← проверки результата без офиса
  report.py                   ← Markdown/JSON отчёт
  suites/default.yaml         ← манифест сценариев

macro-lib/
  qa_collect_run.py
  pythonpath/libre_macros_qa_lib.py

test/collect_workbooks/
  workbooks/*.xltx | ODF/*.ots
  sources/
  workbooks/results/          ← ручные прогоны
  workbooks/results/headless/ ← артефакты раннера (gitignore)
```
Генератор шаблонов (`generate_manual_workbooks.py`) **не заменяется**. Раннер потребляет уже сгенерированные / пользовательские книги.

Системный Python раннера ≠ Python внутри офиса. Раннер:

1. стартует офис;
2. по UNO открывает шаблон, вызывает макрос **внутри** процесса офиса;
3. сохраняет файл;
4. закрывает документ;
5. **отдельным** процессом читает сохранённый файл и гоняет checks.

Так проверки не зависят от зависания UNO после сбора.

---

## 4. Вход: манифест (путь к шаблонам и куда класть результат)

Файл YAML или JSON (v1 — YAML). Пути — абсолютные или относительно корня репозитория / `LM_QA_ROOT`.

```yaml
v: 1
office: acell                 # acell | soffice | auto
templates_dir: test/collect_workbooks/workbooks/ODF
sources_dir: test/collect_workbooks/sources
results_dir: test/collect_workbooks/workbooks/results/headless
report_path: test/collect_workbooks/workbooks/results/headless/report.md
timeout_sec: 300
copy_template: true           # не портить исходный .ots — копия в results_dir/_run
param_sheet: ""               # пусто = активный / get_active_param_sheet_name
cases:
  - id: 01_merge_one_sheet
    template: 01_merge_one_sheet.ots
    param_sheet: Параметры_Объединения   # если несколько листов
    save_as: 01_merge_one_sheet.ods
    checks_file: suites/checks/01_merge.yaml   # или встроенный checks:
    checks:
      - {type: log_macro_version}
      - {type: log_no_status, values: [ошибка, error]}
      - {type: sheet_exists, name: Merge}
      - {type: header_contains, sheet: Merge, names: ["ФИО"]}
      - {type: min_data_rows, sheet: Merge, n: 1}
```

CLI:

```bash
python qa/collect_headless/run.py --manifest qa/collect_headless/suites/default.yaml
python qa/collect_headless/run.py --templates-dir /path/to/ots --results-dir /tmp/out --glob '01_*.ots'
python qa/collect_headless/run.py --only 01_merge_one_sheet --keep-office  # отладка
```

Минимальный контракт без YAML: **`--templates-dir` + `--results-dir`**. Каждый `*.ots`/`*.ods`/`*.xlsx` с листом параметров = один case; проверки по умолчанию — «макрос вернул OK + есть журнал + нет статуса ошибка». Детальные assert — в манифесте.

Переменные окружения:

| Переменная | Смысл |
|------------|--------|
| `LM_QA_OFFICE` | путь к `acell` / `soffice` / `soffice.bin` |
| `LM_QA_USER_INSTALL` | `file:///…/qa-profile` — изолированный профиль с уже установленными макросами |
| `LM_QA_PORT` | порт UNO (default 2002) |
| `LM_QA_DISPLAY` | если задан — не стартовать Xvfb |

---

## 5. Запуск офиса (ACell / soffice)

### 5.1. Процесс

Типовые флаги (уточнить на AlterOS для бинаря ACell):

```text
--headless --invisible --nologo --norestore --nofirststartwizard --nodefault
--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext
-env:UserInstallation=file:///…/qa-profile
```

Профиль **отдельный** от пользовательского `~/.config/alteroffice` / `libreoffice`, чтобы не трогать GUI-сессию. В профиль один раз копируются макросы (`install.sh` с `AO_MACROS_DIR` / `LO_MACROS_DIR` на `qa-profile/user/Scripts/python`).

Поиск бинаря (`office: auto`):

1. `$LM_QA_OFFICE`
2. `acell`, затем `soffice` в PATH
3. известные пути AlterOffice / LibreOffice на Linux

Windows v1: тот же сокет; `soffice.exe` vs `soffice.bin` — как в `run_calc_sal_log.ps1` (ждать `soffice.bin`).

### 5.2. Вызов макроса

Не полагаться на `macro:///…` из argv (разные URI у AO/LO). Предпочтительно:

1. `python3` офиса **или** системный с `uno` (если установлен пакет office-python);
2. `desktop.loadComponentFromURL(template_url, …)` с `Hidden=True`, `AsTemplate=True` для `.ots`/`.xltx`;
3. активировать лист параметров;
4. импорт/вызов `collect_workbooks_for_param_sheet(...)` через script provider **или** bootstrap: добавить `pythonpath` профиля и вызвать функцию, если раннер крутится внутри макро-контекста.

Рекомендуемый v1-мост: маленький скрипт `qa/collect_headless/uno_job.py`, который офис выполняет как:

`vnd.sun.star.script:…`  
либо `soffice … macro:///` только если на целевой сборке URI стабилен.

Аргументы job (файл JSON рядом с копией шаблона):

```json
{
  "param_sheet": "Параметры_Объединения",
  "save_path": "/abs/results/01_merge_one_sheet.ods",
  "skip_consent": true,
  "skip_confirm": true,
  "quiet_finish": true
}
```

После `collect_workbooks_for_param_sheet` → `storeAsURL` / `store` на `save_path` → `close(True)`. Код возврата job: 0 только если `started and ok` (как `_MERGE_LAST_COLLECT_OK`).

### 5.3. Таймаут и зависание

- `timeout_sec` на case: kill дерева процесса офиса, case = fail (`timeout`).
- Не переиспользовать один процесс офиса между case в v1 (утечки Hidden-документов). Опция `--keep-office` — только отладка.

---

## 6. Проверки (скрипт)

Читать **сохранённый** файл (не живой UNO), чтобы не держать lock.

### 6.1. Обязательные типы v1

| `type` | Поля | Смысл |
|--------|------|--------|
| `log_exists` | — | лист журнала (`Сбор_книг_лог` / имя из cfg) |
| `log_macro_version` | — | в логе есть `MACRO_VERSION` / `version.txt` на момент прогона |
| `log_no_status` | `values: [ошибка]` | ни одна строка статуса не содержит подстроки (casefold) |
| `log_contains` | `text` | подстрока в журнале |
| `sheet_exists` | `name` | лист есть |
| `sheet_absent` | `name` | листа нет |
| `min_data_rows` | `sheet`, `n` | строк данных ≥ n (ниже заголовка) |
| `max_data_rows` | `sheet`, `n` | |
| `header_contains` | `sheet`, `names[]` | заголовки содержат токены (как `column_header_matches_token` — желательно переиспользовать codec) |
| `cell_equals` | `sheet`, `cell` (A1), `value` | точное сравнение (строка после trim); числа — с `abs_tol` |
| `cell_not_empty` | `sheet`, `cell` | |
| `golden_sheet` | `path`, `sheet`, `range` | сравнить диапазон с эталоном `.xlsx`/`.ods` |

### 6.2. Вне v1 (не блокируют ETL-приёмку)

Цвета, borders, print page style, видимость автофильтра, диалоги.

Политика: check с `"gui": true` → **skip** в headless, в отчёте отдельная секция «только ручной GUI».

### 6.3. Ручная проверка

Раннер всегда оставляет файлы в `results_dir`. Человек открывает их в обычном ACell. Отчёт Markdown: таблица case / pass / fail / skip + путь к файлу + выдержка из журнала.

Режим `--no-checks`: только прогон и сохранение (как сейчас шаг 6 в `12_MANUAL_TESTING.md`).

---

## 7. Что входит в первый набор сценариев

Переиспользовать существующие ODF-шаблоны из `test/collect_workbooks/workbooks/ODF/` (после `generate.sh --odf`). Приоритет data-manipulation:

| Case | Зачем |
|------|--------|
| `01_merge_one_sheet` | append на один лист, базовые значения |
| `02_merge_many_sheets` | несколько вкладок |
| `03_columns_filter` | набор столбцов |
| `11_direct_xml_excel_ods` | xml/direct + значения |
| `11_vlookup_1to1_low` | join left, извлечённые поля |
| `12_postprocess_json` | не все RANGE (сетка skip); assert: `замена_значений`, `количество_значений`, `удалить_дубликаты` если есть в шаблоне, журнал без json-ошибок |

Новые fn (unpivot, group by, …) — отдельные case в манифесте по мере появления шаблонов.

Пути `Файлы-Источники` в шаблоне должны резолвиться: либо генератор пишет абсолютные пути, либо раннер **перед** запуском подменяет ячейки B на `sources_dir` (v1: обязательная подмена, если путь не существует).

---

## 8. Изменения в макросах (минимальные, под раннер)

Не раздувать `collect_workbooks.py`. Предпочтительно `pythonpath/libre_macros_qa_lib.py` + тонкая entry `qa_collect_run.py` **только если** script provider иначе не вызывает `collect_workbooks_for_param_sheet`.

Обязательно для стабильного headless:

1. Гарантия: при `quiet_finish` + skip consent/confirm **нет** блокирующего `execute()` MessageBox (сейчас pack это обещает — зафиксировать тестом «прогон 01 без DISPLAY», на Xvfb).
2. Прогресс: если Toolkit падает — fallback без UI, запись в журнал, сбор продолжается.
3. Не показывать диалог «Завершить»; не `ask_yes_no`.
4. Документ Hidden: `getCurrentFrame()` может быть пустым — все `message_parent_window` / progress не должны падать (уже есть переборы родителей — проверить на Hidden).

Визард **не** адаптировать под headless.

---

## 9. Отчёт и код выхода

- stdout: краткая таблица.
- `report.md` + `report.json` в `results_dir`.
- Exit code: `0` — все обязательные checks pass; `1` — есть fail; `2` — офис не стартовал / UNO.

Не падать из-за skip GUI-checks.

---

## 10. Безопасность и изоляция

- Не коммитить результаты прогонов и lock-файлы LO (`.~lock.*`).
- `results_dir/headless/` в `.gitignore`.
- Профиль qa не содержит пользовательских документов.
- Макросы в qa-профиле = копия из `macro-lib` после `install.sh` (тот же `MACRO_VERSION`).

---

## 11. Приёмка v1

- [ ] CLI: `--templates-dir` + `--results-dir` прогоняет ≥1 `.ots`.
- [ ] Манифест с 3+ case из таблицы §7.
- [ ] Consent / confirm / finish не блокируют (Xvfb допустим).
- [ ] Результат открывается вручную в ACell; журнал содержит версию.
- [ ] Checks: `log_no_status` + `min_data_rows` на Merge-подобном листе.
- [ ] `param_wizard` не вызывается.
- [ ] Документация: короткий раздел в `docs/12_MANUAL_TESTING.md` («Headless-прогон») + этот файл остаётся ТЗ.
- [ ] Правило «не pytest в `test/`» соблюдено: код раннера в `qa/collect_headless/`.

---

## 12. Вне объёма v1

- CI без установленного ACell/LO (job только если в образе есть офис).
- Сравнение скриншотов.
- Нагрузочные 1toN high как обязательные (можно `optional: true`).
- Оркестратор `collect_pack` с несколькими листами параметров (позже: цикл `param_sheet` в одном файле).

---

## 13. Ручная проверка самого раннера

1. `./macro-lib/install.sh` в qa-профиль.
2. `python test/collect_workbooks/generate_manual_workbooks.py --odf` при необходимости.
3. Прогон `01_merge_one_sheet.ots`.
4. Открыть `results_dir/*.ods` глазами: данные на месте, лог без ошибки.
5. Намеренно сломать путь источника → fail + запись в report, офис завершился по таймауту или вернул ok=false.
