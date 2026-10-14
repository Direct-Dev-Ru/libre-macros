# atext_letters

- `source_recipients.xlsx` — лист «Получатели» (ФИО, Адрес, Сумма).
- `template_letter.odt` — `block_rows=1`, плейсхолдеры `<<ФИО>>` и т.д.
- `template_pair.odt` — `block_rows=2`, `<<…>>{1}` / `{2}`.

Сценарий: `20_atext_letters.xltx`

Пересоздать:
```bash
python test/collect_workbooks/generate_atext_letters_sources.py
```
