# QA libre-macros

Автоматизированные **сценарии** (не pytest в `test/`). Цель — прогонять сбор данных и проверять **значения** в сохранённых книгах без кликов в визарде.

| Каталог | Назначение |
|---------|------------|
| [collect_headless/](collect_headless/README.md) | Headless-прогон `collect_workbooks` через ACell / LibreOffice + UNO |
| [collect_headless/CHECKS.md](collect_headless/CHECKS.md) | Справочник типов проверок |
| [collect_headless/ADDING_CASES.md](collect_headless/ADDING_CASES.md) | Как добавить свой сценарий |

Ручное тестирование в Calc по-прежнему описано в [docs/12_MANUAL_TESTING.md](../docs/12_MANUAL_TESTING.md).  
ТЗ модуля: [todo/модуль_headless_тестирования_сбора.md](../todo/модуль_headless_тестирования_сбора.md).

## Быстрый старт

```bash
# из корня репозитория
python3 qa/collect_headless/run.py
# или
./qa/collect_headless/run.sh --only 01_merge_one_sheet
```

Нужны: LibreOffice / ACell (`soffice` в PATH или `LM_QA_OFFICE`), шаблоны `test/collect_workbooks/workbooks/ODF/*.ots`, для `.xlsx`-результатов — `openpyxl`.

## Политика

- **Не** добавлять pytest в `test/` под этот контур.
- Визард (`param_wizard`), file picker, GUI-диалоги «Запуск»/«Завершить» — **вне** headless-проверок.
- Результаты headless лежат в `test/collect_workbooks/workbooks/results/headless/` (gitignore).
- QA-профиль офиса: `qa/collect_headless/.profile/` (gitignore) — не смешивать с личной установкой макросов.
