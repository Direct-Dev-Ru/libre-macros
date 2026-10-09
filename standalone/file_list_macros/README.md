# Standalone: file_list_macro

Самодостаточная книга Calc с макросом сбора файлов/папок, поиском по именам/содержимому,
листом «поиск_шаблоны» (словарь по умолчанию) и вшитой справкой.

## Пересборка

```bash
python3 standalone/file_list_macros/build_file_list_ods.py
```

Результат:

| Путь | Что |
|------|-----|
| `file_list_macro.ods` | книга с макросами и кнопками |
| `docs/file_list_user_guide.odt` | справка (вшивается в ODS as-is; править в Writer) |
| `_staged_python/` | копии скриптов, попавших в ODS |

## Зависимости внутри ODS

Один файл `Scripts/python/file_list_macro.py` (cfg + lib + help blob + entry).

**Почему не pythonpath:** в AlterOffice document-макросы копируются во временный файл;
`Scripts/python/pythonpath` из `tdoc:` на `sys.path` не попадает → ImportError и
«Could not compile file and find macro».

Константы: `class _C` (file_list cfg) и `class _Ui` (тема soft-gray/green).
Стиль заголовка и меню-автофильтр — локальные fallback’и в entry (без `libre_macros_lib`).

## Справка

Источник — только `docs/file_list_user_guide.odt`. Скрипт сборки **не** пересобирает ODT;
после правок в Writer снова запустите `build_file_list_ods.py`, чтобы обновить blob в ODS.
