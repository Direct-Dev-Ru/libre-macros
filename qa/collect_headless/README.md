# Headless-тестирование сбора (`collect_workbooks`)

Подробное руководство по раннеру в `qa/collect_headless/`.  
Кратко: [qa/README.md](../README.md) · ТЗ: [todo/модуль_headless_тестирования_сбора.md](../../todo/модуль_headless_тестирования_сбора.md) · ручной Calc: [docs/12_MANUAL_TESTING.md](../../docs/12_MANUAL_TESTING.md).

**Версия макросов** при прогоне берётся из `macro-lib/version.txt` (отчёт и проверка `log_macro_version`).

---

## 1. Что проверяем и что нет

| В scope | Вне scope (v1) |
|---------|----------------|
| Сбор по листу `Параметры_Объединения*` | Визард `param_wizard` |
| Режимы «на один лист» / «на разные» / копирование листов | Диалоги «Запуск» / «Завершить» / consent |
| `Способ_переноса` = `xml_excel_ods` | File picker, оркестратор с GUI |
| Постобработка RANGE/ROW/FINAL и ВПР **по данным** | Визуал: градиент, зебра, ширина, автофильтр «на глаз» |
| Журнал `Сбор_книг_лог` (версия, статусы, строки) | Pytest в `test/` |

Раннер **обязан** вызывать тихий путь:

`collect_workbooks_for_param_sheet(..., skip_consent=True, skip_confirm=True, quiet_finish=True)`  

через макрос `qa_collect_run` + флаг `_MERGE_QA_HEADLESS` (прогресс = status, без MessageBox).

---

## 2. Архитектура прогона

```text
qa/collect_headless/
  run.py / run.sh     ← CLI, манифест, цикл case, отчёт
  office.py           ← поиск soffice/acell, UserInstallation, Xvfb, копирование макросов
  patch_sources.py    ← копия шаблона + подмена Файлы-Источники / Папка_результата / прогресс
  uno_job.py          ← UNO: открыть книгу → invoke qa_collect_run → закрыть
  checks.py           ← проверки .ods/.xlsx без UNO
  report.py           ← report.md + report.json
  suites/default.yaml ← манифест по умолчанию
  .profile/           ← изолированный профиль (gitignore)

macro-lib/
  qa_collect_run.py              ← entry (меню / script provider)
  pythonpath/libre_macros_qa_lib.py  ← job JSON → сбор → storeAsURL + *.qa.json
```

Поток одного case:

1. Копия шаблона в `results_dir/_run/<id>/template.*`.
2. Патч путей источников на `sources_dir`, папки xml-результата — на `_run/<id>/xml_result`.
3. Запись `job.json` (пути, порт UNO, флаги quiet).
4. Старт офиса `--headless` с `--accept=socket…` и `UserInstallation` = QA-профиль.
5. `uno_job.py` (интерпретатор с UNO): `AsTemplate` → при необходимости **сохранить на диск** (нужно для `xml_excel_ods`) → `createScriptProvider(doc)` → `qa_collect_run(job_path)`.
6. Макрос пишет результат в `save_path` и маркер `save_path.qa.json`.
7. Раннер читает сохранённый файл через `checks.py`, пишет сводку в `report.md` / `report.json`.

Системный Python раннера ≠ Python внутри офиса. Проверки данных идут **после** сохранения, отдельным процессом.

---

## 3. Требования к среде

| Компонент | Зачем |
|-----------|--------|
| ACell или LibreOffice (`soffice` / `acell`) | Headless + UNO |
| Python 3 | CLI раннера |
| Пакет `uno` для выбранного Python | Обычно системный `/usr/bin/python3` с `python3-uno`, либо `…/program/python` офиса |
| `openpyxl` | Только если результат / эталон — `.xlsx` |
| Шаблоны ODF | `test/collect_workbooks/workbooks/ODF/*.ots` |
| Источники | `test/collect_workbooks/sources/` |
| `Xvfb` (опционально) | Если нет `DISPLAY` и AWT всё же дергается |

Генерация шаблонов (если ODF нет):

```bash
./test/collect_workbooks/generate.sh
python3 test/collect_workbooks/generate_manual_workbooks.py --odf
```

Переменные окружения:

| Переменная | Смысл |
|------------|--------|
| `LM_QA_OFFICE` | Путь или имя бинарника офиса |
| `LM_QA_OFFICE_PYTHON` | Python с UNO для `uno_job.py` |
| `LM_QA_USER_INSTALL` | Каталог профиля (не `file://`) |
| `LM_QA_HOST` / `LM_QA_PORT` | UNO-сокет (по умолчанию `127.0.0.1:2002`) |
| `LM_QA_DISPLAY` / `LM_QA_XVFB_DISPLAY` | Дисплей / номер Xvfb (`:91`) |

---

## 4. Запуск

Из **корня репозитория**:

```bash
# весь default-манифест
python3 qa/collect_headless/run.py
./qa/collect_headless/run.sh

# один или несколько id
python3 qa/collect_headless/run.py --only 01_merge_one_sheet
python3 qa/collect_headless/run.py --only 01_merge_one_sheet 03_columns_filter

# быстрее на пачке: один процесс офиса
python3 qa/collect_headless/run.py --keep-office --timeout 300

# свой манифест / каталоги
python3 qa/collect_headless/run.py --manifest qa/collect_headless/suites/default.yaml
python3 qa/collect_headless/run.py \
  --templates-dir test/collect_workbooks/workbooks/ODF \
  --sources-dir test/collect_workbooks/sources \
  --results-dir /tmp/lm-qa \
  --glob '01_*.ots'

# без переустановки макросов в профиль (уже стоят)
python3 qa/collect_headless/run.py --no-install --only 01_merge_one_sheet

# только прогон и сохранение, без checks
python3 qa/collect_headless/run.py --no-checks --only 02_merge_many_sheets
```

### CLI

| Флаг | Описание |
|------|----------|
| `--manifest PATH` | JSON/YAML; по умолчанию `suites/default.yaml` |
| `--templates-dir` | Каталог шаблонов |
| `--sources-dir` | Каталог для подмены `Файлы-Источники` |
| `--results-dir` | Куда класть `.ods` и отчёты |
| `--glob` | Маска, если манифест без `cases` (по умолчанию `*.ots`) |
| `--only ID…` | Фильтр по `id` case |
| `--office auto\|acell\|soffice` | Предпочтение бинарника |
| `--timeout N` | Секунд на case (перекрывает манифест) |
| `--keep-office` | Один soffice на все case |
| `--no-install` | Не копировать `macro-lib` в профиль |
| `--no-checks` | Не вызывать `checks.py` |
| `--profile DIR` | Свой UserInstallation |

Код выхода: `0` — все case ok и checks без fail; `1` — есть fail; `2` — нет манифеста/шаблонов/файлов результата.

---

## 5. Манифест

Формат: **JSON** (как `suites/default.yaml` сейчас) или YAML при установленном PyYAML.

```json
{
  "v": 1,
  "office": "auto",
  "templates_dir": "test/collect_workbooks/workbooks/ODF",
  "sources_dir": "test/collect_workbooks/sources",
  "results_dir": "test/collect_workbooks/workbooks/results/headless",
  "report_path": "test/collect_workbooks/workbooks/results/headless/report.md",
  "timeout_sec": 300,
  "copy_template": true,
  "param_sheet": "",
  "cases": [
    {
      "id": "01_merge_one_sheet",
      "template": "01_merge_one_sheet.ots",
      "save_as": "01_merge_one_sheet.ods",
      "param_sheet": "Параметры_Объединения",
      "suppress_pre_clear": false,
      "checks": [
        {"type": "log_exists"},
        {"type": "sheet_exists", "name": "Merge"},
        {"type": "min_data_rows", "sheet": "Merge", "n": 1}
      ]
    }
  ]
}
```

Пути — абсолютные или относительно корня репозитория.

Поля case:

| Поле | Смысл |
|------|--------|
| `id` | Имя в отчёте и фильтре `--only` |
| `template` | Файл в `templates_dir` |
| `save_as` | Имя результата в `results_dir` |
| `param_sheet` | Лист параметров; пусто — единственный / активный |
| `suppress_pre_clear` | Передаётся в `collect_workbooks_for_param_sheet` |
| `checks` | Список проверок (см. [CHECKS.md](CHECKS.md)) |

Если `checks` нет — берутся `default_checks()` из `checks.py`.

---

## 6. Сценарии default

Манифест: [suites/default.yaml](suites/default.yaml).

| id | Шаблон | Что смотрим |
|----|--------|-------------|
| `01_merge_one_sheet` | `01_merge_one_sheet.ots` | Merge, ФИО, версия в журнале, без статусов ошибка |
| `02_merge_many_sheets` | `02_merge_many_sheets.ots` | Есть данные на листах отчёта |
| `03_columns_filter` | `03_columns_filter.ots` | Merge, заголовки ФИО/Отдел |
| `11_direct_xml_excel_ods` | `11_direct_xml_excel_ods.ots` | xml_excel_ods + журнал/строки |
| `11_vlookup_1to1_low` | `11_vlookup_1to1_low.ots` | Копирование листов + ВПР |
| `12_postprocess_json` | `12_postprocess_json.ots` | Матрица JSON-PP: **только** факт сбора (`log_exists` + `min_data_rows`). В журнале допустимы ошибки отдельных шагов матрицы (нет листа «Шаблон» и т.п.) |

Добавление своего case: скопировать блок в манифесте или отдельный JSON + `--manifest`.

---

## 7. Артефакты после прогона

```text
results/headless/
  report.md                 ← сводка pass/fail + фрагмент журнала
  report.json               ← машинночитаемый отчёт
  <id>.ods                  ← сохранённая книга
  <id>.ods.qa.json          ← маркер макроса (ok, parse_err, log_rows, …)
  _run/<id>/
    template.ots            ← пропатченная копия
    job.json                ← задание для uno_job / макроса
    xml_result/             ← Папка_результата для xml-режима
```

Маркер `*.qa.json` (поля):

| Поле | Смысл |
|------|--------|
| `ok` | Успех сбора (`_MERGE_LAST_COLLECT_OK` / started) |
| `save_path` | Куда storeAsURL |
| `param_sheet` | Фактический лист параметров |
| `started_ok` / `last_ok` | Диагностика |
| `parse_err` | Ошибка `parse_collect_settings` до/вокруг старта |
| `has_xsc` | Был ли `XSCRIPTCONTEXT` |
| `log_rows` | Снимок строк журнала через UNO |
| `error` | Исключение entry |

Профиль `qa/collect_headless/.profile/` после первого прогона содержит копию макросов (`user/Scripts/python/`, в т.ч. `qa_collect_run.py`). При обычном запуске макросы **переустанавливаются** из `macro-lib/` (если нет `--no-install`).

---

## 8. Как устроен тихий макрос

1. Entry: `macro-lib/qa_collect_run.py` → `libre_macros_qa_lib.entry(doc, XSCRIPTCONTEXT, job_path)`.
2. `collect_workbooks` грузится по абсолютному пути из `Scripts/python/` (UNO-import часто не видит соседний entry).
3. В модуль инжектится `XSCRIPTCONTEXT` (иначе `NameError` при закрытии orphan docs и т.п.).
4. Untitled / `.ots` после `AsTemplate` **сохраняется на диск** до сбора (требование `xml_excel_ods`).
5. `_MERGE_QA_HEADLESS=True`: `show_message`/`ask_yes_no` без окон, прогресс = status.
6. Вызов `collect_workbooks_for_param_sheet(..., skip_consent/skip_confirm/quiet_finish=True)`.
7. `storeAsURL` + запись маркера.

URI скрипта в UNO:

`vnd.sun.star.script:qa_collect_run.py$qa_collect_run?language=Python&location=user`

---

## 9. Подмена путей в шаблоне

`patch_sources.py` ищет лист `Параметры_Объединения*` / `collect_params*` и правит:

- **Файлы-Источники** — пути с `/sources/` или известными именами файлов → `sources_dir`;
- **Папка_результата** — на каталог xml-результата case (если параметр есть);
- **Отображение_прогресса** — в сторону `status` (чтобы не зависнуть на dialog);
- **Способ_переноса** (отдельная строка B или колонка C у «Режим») — пусто / `Буфер*` → явно `По_API` (буфер в headless пуст); `xml_excel_ods` не меняется.

Исходные `.ots` в git **не портятся** — правится только копия в `_run/`.

---

## 10. Чеклист перед релизом / после правок сбора

1. Обновить макросы: правки в `macro-lib/` → bump `version.txt` → `sync_macro_version.py` → при AO — flatten + compat check.
2. При необходимости перегенерировать ODF-шаблоны.
3. `python3 qa/collect_headless/run.py --keep-office` — ожидать exit 0.
4. Открыть один `.ods` из `results/headless/` в Calc глазами (оформление).
5. При падении — смотреть `report.md`, `*.qa.json`, лист `Сбор_книг_лог` в результате.

---

## 11. Типичные проблемы

| Симптом | Что проверить |
|---------|----------------|
| `No module named 'collect_workbooks'` | Старая копия `libre_macros_qa_lib` в профиле; перезапуск **без** `--no-install` |
| `нет открытой книги` | `createScriptProvider(doc)`, не пустой provider; не `Hidden=True` без привязки к doc |
| `книга ещё не сохранена` (xml) | Pre-save в `libre_macros_qa_lib` до collect |
| `Неизвестная функция … 'зебра'` | Алиас `зебра` → `зебра_диапазон` в `param_codec`; переустановить макросы |
| Пустой журнал / нет Merge | `parse_err` в `*.qa.json`; пути источников после патча; `ok: false` |
| Зависание на timeout | Прогресс dialog в шаблоне; убрать `--keep-office` и убить `soffice`; Xvfb |
| Порт занят | Сменить `LM_QA_PORT` / убить старый headless |
| Checks fail, макрос ok | Смотреть `report.md` по типам; для 12 — не требовать `log_no_status` |
| UNO / python | `LM_QA_OFFICE_PYTHON` или системный python с `uno` |

Убить зависший офис QA (осторожно с обычными сессиями):

```bash
pkill -f 'UserInstallation=.*qa/collect_headless/.profile' || true
```

---

## 12. Расширение

1. Новый шаблон в `test/collect_workbooks/` (+ `--odf` при необходимости).
2. Case в манифесте + осмысленные `checks` ([CHECKS.md](CHECKS.md)).
3. Для эталонного листа — `golden_sheet` с эталоном вне gitignore или в отдельном каталоге fixtures.
4. Не тащить GUI-assert’ы: пометку `"gui": true` в check раннер пропускает как `skip`.

Связанный код макроса: `collect_workbooks.py`, `libre_macros_collect_cfg.py` (`_MERGE_QA_HEADLESS`), `libre_macros_qa_lib.py`.
