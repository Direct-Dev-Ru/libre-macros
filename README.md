# alter-macros / libre-macros

Python-макросы для **LibreOffice Calc** и **AlterOffice**: сбор книг, постобработка листов и визард параметров.

## Возможности

- **`collect_workbooks` / `collect_pack`** — сбор данных из `.xlsx` / `.xls` / `.ods` / CSV / XML по листу `Параметры_Объединения*`
- постобработка RANGE / ROW / FINAL и `Постобработка_xml` (`xml_excel_ods`)
- параметры в колонке **C только JSON**
- визард `param_wizard`, совместимость с фильтром скриптов **AlterOffice 2026**
- shape-функции: **`транспонировать_таблицу`**, **`анкета_в_таблицу` / `таблица_в_анкету`**, unpivot, split, …

## Быстрый старт

```bash
./macro-lib/install.sh
```

Документация: [docs/00_INDEX.md](docs/00_INDEX.md) · установка: [docs/01_QUICKSTART.md](docs/01_QUICKSTART.md)

Версия макросов: `macro-lib/version.txt` (сейчас **3.10.696**).
