# Презентации libre-macros

Материалы для руководства: переход с Excel на AlterOffice + пакет макросов **libre-macros** (импортозамещение).

| Раскадровка | Слайдов | ID | Файл |
|-------------|---------|-----|------|
| Полная | 14 | `full` | [STORYBOARD.md](STORYBOARD.md) |
| Краткая | 5 (титул + 4 смысловых) | `short` | [STORYBOARD_short.md](STORYBOARD_short.md) |

Картинки — в [assets/](assets/) (имена и промпты в раскадровках). Нет файла → заглушка; у `short` есть fallback на картинки полной версии.

`python-pptx` в macro-lib **не** бандлится — только локальная зависимость для сборки.

## Установка и сборка

Из корня репозитория или из этой папки:

```bash
# зависимость (лучше в .venv проекта)
.venv/bin/pip install -r presentation/requirements.txt

# полная
.venv/bin/python presentation/build_pptx.py --storyboard full --out libre_macros_import_substitution.pptx

# краткая (killer: часть VBA не нужно переписывать с нуля)
.venv/bin/python presentation/build_pptx.py --storyboard short --out libre_macros_short.pptx
```

Раскадровку можно указать файлом:

```bash
.venv/bin/python presentation/build_pptx.py --storyboard presentation/STORYBOARD_short.md --out short.pptx
```

### Параметры

| Флаг | Смысл | По умолчанию |
|------|--------|--------------|
| `--storyboard` | `full` / `short` или путь к `STORYBOARD*.md` | `full` |
| `--out` | Имя или путь выходного `.pptx` | из раскадровки |
| `--org` | Подпись организации в футере | `Предприятие` |
| `--year` | Год на титуле | `2026` |

Имя файла без пути (`--out short.pptx`) пишется в `presentation/`.

## Файлы

| Путь | Назначение |
|------|------------|
| `build_pptx.py` | Генератор PPTX |
| `STORYBOARD.md` | Полная раскадровка + промпты картинок |
| `STORYBOARD_short.md` | Краткая (5 слайдов, акцент на VBA) |
| `assets/` | Картинки для вставки |
| `requirements.txt` | `python-pptx` |
| `*.pptx` | Собранные презентации (пересобрать скриптом) |

## Краткая версия — суть

1. Титул: Excel → AlterOffice без переписывания всего VBA  
2. Проблема: парк VBA нужно переносить — дорого  
3. Killer: libre-macros закрывает типовой класс «собрать → очистить → свести → оформить»  
4. На практике: шаги прогона и выгоды (сценарий, визард, AO 2026)  
5. Предложение: пилот 1–2 отчёта  

Открывать в LibreOffice Impress / AlterOffice / PowerPoint.
