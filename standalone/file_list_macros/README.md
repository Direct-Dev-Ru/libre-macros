# Standalone: file_list_macro

Самодостаточная книга Calc с макросом сбора файлов/папок, мини-DLP и вшитой справкой.

## Пересборка

```bash
python3 standalone/file_list_macros/build_file_list_ods.py
```

Результат:

| Путь | Что |
|------|-----|
| `file_list_macro.ods` | книга с макросами и кнопками |
| `docs/file_list_user_guide.md` | исходник справки |
| `docs/file_list_user_guide.odt` | справка (и blob внутри ODS) |
| `_staged_python/` | копии скриптов, попавших в ODS |

## Зависимости внутри ODS

Один файл `Scripts/python/file_list_macro.py` (cfg + lib + help blob + entry).

**Почему не pythonpath:** в AlterOffice document-макросы копируются во временный файл;
`Scripts/python/pythonpath` из `tdoc:` на `sys.path` не попадает → ImportError и
«Could not compile file and find macro».

Константы: `class _C` (file_list cfg) и `class _Ui` (тема soft-gray/green).
`libre_macros_lib` (автофильтр/стиль заголовка) по-прежнему optional.

## Скриншоты в справке

В `docs/file_list_user_guide.md` маркеры:

```text
{{IMAGE: dialog_main | Диалог запуска — вкладка «Основное»}}
```

После правки md пересоберите скриптом — в ODT появятся блоки «МЕСТО ДЛЯ РИСУНКА»; вставьте туда картинки вручную в Writer.
