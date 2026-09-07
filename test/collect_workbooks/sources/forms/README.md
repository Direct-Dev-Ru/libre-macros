# Источники анкета ↔ таблица

`source_forms.xlsx` — листы:

| Лист | Назначение |
|------|------------|
| **Анкета_ФИО** | Q/A, `by_repeat_key` (ключ ФИО) → `анкета_в_таблицу` |
| **Анкета_пусто** | Q/A через пустую строку → `by_blank_row` |
| **Анкета_шапка** | шапка листа + шапка блока + тело |
| **Wide_анкеты** | wide → `таблица_в_анкету` |

Сценарии: `14_form_to_table.xltx`, `15_table_to_form.xltx` (режим «Копирование листов»).

Генерация:

```bash
python test/collect_workbooks/generate_form_sources.py
./test/collect_workbooks/generate.sh
```
