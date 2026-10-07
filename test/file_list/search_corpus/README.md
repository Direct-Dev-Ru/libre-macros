# search_corpus — тестовые файлы для вкладки «Поиск»

Словарь паттернов: `../dlp_name_patterns_example.csv`.

| Папка | Имена | Содержимое |
|-------|-------|------------|
| `hit_name/` | с маркерами (конфиденц, secret, гриф…) | чистое |
| `hit_content/` | нейтральные | с маркерами из словаря |
| `hit_both/` | с маркерами | с маркерами |
| `clean/` | нейтральные | чистое |
| `nested/` | смешанные, вложенные папки | смешанное |
| `legacy/` | .doc/.dot/.xls/.xlt (LibreOffice convert); `.xlsb` — stub (копия .xls, export filter часто отсутствует) | hit/clean |

Пересоздать: `python3 test/file_list/generate_search_corpus.py`
