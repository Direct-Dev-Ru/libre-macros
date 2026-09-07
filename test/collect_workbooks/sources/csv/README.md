# CSV-источники для collect_workbooks

Сгенерировано: `python test/collect_workbooks/generate_csv_sources.py`

Кодировка файлов: **Windows-1251** (как default в `Доп_Параметры_Источника`).

| Файл | Строк данных | Разделитель | Кодировка | Назначение |
|------|-------------:|:-----------:|:---------:|------------|
| `source_csv_tiny.csv` | 25 | `;` | `cp1251` | объём **tiny** |
| `source_csv_small.csv` | 250 | `,` | `cp1251` | объём **small** |
| `source_csv_medium.csv` | 2500 | `;` | `cp1251` | объём **medium** |
| `source_csv_large.csv` | 25000 | `;` | `cp1251` | объём **large** |
| `source_csv_xlarge.csv` | 100000 | `;` | `cp1251` | объём **xlarge** |

## Дополнительные примеры для отладки (tab + UTF-16 LE BOM)

| Файл | Строк данных | Разделитель | Кодировка |
|------|-------------:|:-----------:|:---------:|
| `source_csv_tab_utf16le_bom_1.csv` | 420 | `tab` | `utf-16` (UTF-16LE + BOM) |
| `source_csv_tab_utf16le_bom_2.csv` | 350 | `tab` | `utf-16` (UTF-16LE + BOM) |

Параметры на листе (пример `Доп_Параметры_Источника`):
```json
{"delimiter":"\\t","encoding":"utf-16"}
```

Чтобы принудительно использовать «теневую» схему `CSV -> temp XLSX`
(и обрабатывать источник как обычный xlsx), добавьте:
```json
{"delimiter":"\\t","encoding":"utf-16","open_as_xlsx":true}
```

## Параметры на листе

```
Файлы-Источники | …/sources/csv/source_csv_tiny.csv | …/source_csv_small.csv | …
Доп_Параметры_Источника | {"delimiter":";","encoding":"windows-1251"} | {"delimiter":",","encoding":"windows-1251"} | …
```

Для режима «Копирование листов» + перенос листом целиком (Calc/`importSheet`):

```json
{"delimiter":";","encoding":"windows-1251","copy_whole_sheet":true}
```

Пустая `Доп_Параметры_Источника` → разделитель `;`, кодировка `windows-1251`, `copy_whole_sheet=false`.
Для UTF-8-файлов укажите `"encoding":"utf-8"` (или `utf-8-sig`).

Колонки: ФИО, Табельный_номер, Отдел, Сумма_руб, Сумма_бонус, маркер, ДатаВремя_строка, Филиал.

В режиме «Копирование листов» / «На один лист» / «На разные листы» CSV
по умолчанию сначала читается в память, там же распознаются числа/даты (`MERGE_XML_CONVERT_TO_NUMBERS`),
затем пишется на лист(ы) результата (`setDataArray`; матчинг заголовков или позиция).
Опция `copy_whole_sheet` влияет только на «Копирование листов».
В консоль Calc — лог `[CSV] …` с таймингами.