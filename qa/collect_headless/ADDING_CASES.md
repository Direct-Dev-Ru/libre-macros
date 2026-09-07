# Как добавить сценарий headless

Краткая памятка. Полный контекст: [README.md](README.md), проверки: [CHECKS.md](CHECKS.md).

## 1. Шаблон

1. Добавить/обновить сценарий в `test/collect_workbooks/generate_manual_workbooks.py` (или взять готовый `.xltx`).
2. Сгенерировать ODF при необходимости:

```bash
python3 test/collect_workbooks/generate_manual_workbooks.py --odf
```

3. Убедиться, что в листе параметров пути к источникам содержат `/sources/` или узнаваемые имена файлов — иначе `patch_sources.py` не перепишет пути.

## 2. Case в манифесте

Скопировать блок в `suites/default.yaml` или создать свой JSON:

```json
{
  "id": "04_sheets_wildcard",
  "template": "04_sheets_wildcard.ots",
  "save_as": "04_sheets_wildcard.ods",
  "checks": [
    {"type": "log_exists"},
    {"type": "log_no_status", "values": ["ошибка", "error", "critical"]},
    {"type": "min_data_rows", "n": 1}
  ]
}
```

Запуск только нового:

```bash
python3 qa/collect_headless/run.py --only 04_sheets_wildcard
```

## 3. Выбор checks

| Ситуация | Рекомендация |
|----------|----------------|
| Обычный сбор без ожидаемых ошибок в журнале | `log_exists` + `log_no_status` + данные |
| Матрица / шаги, которые могут писать «ошибка» | Без `log_no_status`; `min_data_rows` / `sheet_exists` |
| Регрессия конкретных ячеек | `cell_equals` / `golden_sheet` |
| Только «макрос отработал» | `--no-checks` или минимум `log_exists` |

## 4. Отладка нового case

1. `run.py --only <id>` → смотреть `results/headless/<id>.ods.qa.json`.
2. Если `ok: false` — `parse_err` / `error` / `log_rows`.
3. Открыть `<id>.ods` → лист `Сбор_книг_лог` и лист результата.
4. Проверить `_run/<id>/template.ots` — подставились ли пути источников.
5. После правок макроса — **без** `--no-install` (перекопировать в `.profile`).

## 5. Не делать

- Не писать pytest в `test/` под этот контур.
- Не assert’ить цвет заливки / ширину столбца в v1 (или помечать `"gui": true`).
- Не править исходные `.ots` в git через раннер — правится только копия в `_run/`.
