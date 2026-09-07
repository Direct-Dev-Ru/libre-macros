# -*- coding: utf-8 -*-
"""
Единый кодек параметров постобработки / финала — только JSON.

param_decode(fn_key, raw_text) -> list[dict]
param_encode(fn_key, blocks) -> str
"""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.688"
import json
import re

try:
    unicode
except NameError:
    unicode = str

PARAM_CODEC_VERSION = 1

DS_LIST = ","

_FN_ALIASES = {
    "раскрасить": "раскрасить_блоки",
    "условное_форматирование": "подсветка_по_порогу",
    "подкрасить_пороги": "подсветка_по_порогу",
    "пропуск_пустых_строк": "удаление_строк",
        "пропуск_строк_источника": "пропуск_строк_источника",
    "строка_заголовков": "строка_заголовков",
    "заголовок": "строка_заголовков",
    "header_row": "строка_заголовков",
    "удалить_листы": "удаление_листов",
    "удалить_служебные_листы": "удаление_листов",
    "нормализация_текста": "текстовые_операции",
    "нормализовать_текст": "текстовые_операции",
    "normalize_text": "текстовые_операции",
    "убрать_повторы_строк": "удалить_дубликаты",
    "dedup": "удалить_дубликаты",
    "разделить_столбец": "разделить_по_столбцам",
    "split_column": "разделить_по_столбцам",
    "разбить_столбец": "разделить_по_столбцам",
    "conditional_column": "условный_столбец",
    "if_column": "условный_столбец",
    "условный_столбец_добавить": "условный_столбец",
    "lambda_column": "столбец_по_лямбде",
    "custom_column": "столбец_по_лямбде",
    "вычислить_столбец": "столбец_по_лямбде",
    "новый_столбец_лямбда": "столбец_по_лямбде",
    "добавить_столбец": "добавить_столбец",
    "add_column": "добавить_столбец",
    "новый_столбец": "добавить_столбец",
    "копировать_лист": "копировать_переместить_лист",
    "fill_up": "заполнить_вверх",
    "заполнение_вверх": "заполнить_вверх",
    "fillup": "заполнить_вверх",
    "unpivot": "развернуть_столбцы",
    "unpivot_columns": "развернуть_столбцы",
    "развернуть столбцы": "развернуть_столбцы",
    "копировать_диапазон": "копирование_диапазонов",
    "copy_ranges": "копирование_диапазонов",
    "copy_range": "копирование_диапазонов",
    "вставить_диапазон": "копирование_диапазонов",
    "создать_лист": "создать_лист",
    "create_sheet": "создать_лист",
    "новый_лист": "создать_лист",
    "зебра": "зебра_диапазон",
    "добавить_в_карту_переменных": "добавить_в_карту_переменных",
    "ручной_ввод_в_карту_переменных": "ручной_ввод_в_карту_переменных",
    "ручной_ввод": "ручной_ввод_в_карту_переменных",
    "manual_variable_input": "ручной_ввод_в_карту_переменных",
    "email": "отправить_по_почте",
    "mailto": "отправить_по_почте",
    "send_mail": "отправить_по_почте",
    "почта": "отправить_по_почте",
}

_SHEET_BLOCK_FORM_FNS = frozenset(
    {
        "градиент",
        "подсветка_по_порогу",
        "формат_даты",
        "формат_деньги",
        "удалить_столбцы",
        "сетка",
        "высота_строки",
        "отступ",
        "шрифт",
        "ширина_столбцов",
        "перенос",
        "перенос_и_авто_высота",
        "заголовок_плюс_высота",
        "зебра_диапазон",
        "конкатенация_столбцов",
        "разделить_по_столбцам",
        "условный_столбец",
        "столбец_по_лямбде",
        "добавить_столбец",
        "применить_формулу",
        "удаление_строк",
        "группировка_по_столбцу",
        "стиль_печати",
        "заполнение_вниз",
        "заполнить_вверх",
        "развернуть_столбцы",
        "копировать_значения",
        "замена_значений",
        "переименовать_лист",
        "переименовать_столбцы",
        "переставить_столбцы",
        "количество_значений",
        "удалить_дубликаты",
        "копировать_переместить_лист",
        "активировать_лист",
        "копирование_диапазонов",
        "создать_лист",
        "преобразовать_числа",
        "цвет_текста_по_значению",
        "удаление_верхних_строк",
        "ячейка_в_столбец",
        "добавить_в_карту_переменных",
        "ручной_ввод_в_карту_переменных",
        "пропуск_строк_источника",
        "строка_заголовков",
        "отправить_по_почте",
    }
)

_SHEET_BLOCK_RULES_FNS = frozenset({"формат_столбцы", "заполнение_вниз_вычислить"})

_SHEET_LIST_FNS = frozenset(
    {
        "тонкая_сетка",
        "толстая_сетка",
        "авто_высота",
        "авто_ширина",
        "левое_выравнивание",
        "сброс_стрипов",
        "закрепить_заголовок",
        "автофильтр",
        "вертикаль_центр",
        "только_значения",
        "удаление_листов",
        "скрытие_листов",
        "серые_служебные",
        "полосы_по_пути",
        "чередующиеся_границы",
        "копировать_формат_заголовка",
        "подсветка_по_заголовку",
        "жирный_по_пути",
    }
)

_PIVOT_FIELDS = (
    "source_sheet",
    "header_row",
    "data_start",
    "data_end",
    "sheet_name",
    "as_values",
    "list_rows_expand",
    "list_rows_include_names",
    "list_rows_name_sep",
    "list_rows_record_sep",
    "row_fields",
    "column_fields",
    "filter_fields",
    "data_fields",
)

def normalize_fn_key(fn_key):
    key = unicode(fn_key or "").strip()
    if key == "":
        return key
    return _FN_ALIASES.get(key.casefold(), key)
def is_sheet_list_fn(fn_key):
    return normalize_fn_key(fn_key) in _SHEET_LIST_FNS
def sheet_list_allows_convert(fn_key):
    return False
def is_sheet_block_form_fn(fn_key):
    return normalize_fn_key(fn_key) in _SHEET_BLOCK_FORM_FNS
def is_sheet_block_rules_fn(fn_key):
    return normalize_fn_key(fn_key) in _SHEET_BLOCK_RULES_FNS
def sheet_block_form_allows_convert(fn_key):
    return normalize_fn_key(fn_key) in (
        "формат_даты",
        "формат_деньги",
        "формат_столбцы",
    )
def is_codec_fn(fn_key):
    fn_key = normalize_fn_key(fn_key)
    if fn_key in (
        "сортировка",
        "раскрасить_блоки",
        "объединить_листы_в_один",
        "разделить_листы",
    ):
        return True
    if is_sheet_list_fn(fn_key) or is_sheet_block_form_fn(fn_key):
        return True
    if is_sheet_block_rules_fn(fn_key):
        return True
    return False
def _normalize_block_for_fn(fn_key, block):
    fn_key = normalize_fn_key(fn_key)
    if fn_key == "сортировка":
        return _normalize_sort_block(block)
    if fn_key == "раскрасить_блоки":
        return _normalize_colorize_block(block)
    if fn_key == "сводная_таблица":
        return _normalize_pivot_block(block)
    if is_sheet_list_fn(fn_key):
        return _normalize_sheet_list_block(fn_key, block)
    if fn_key == "градиент":
        return _normalize_gradient_block(block)
    if fn_key == "подсветка_по_порогу":
        return _normalize_threshold_block(block)
    if fn_key == "формат_деньги":
        return _normalize_format_money_block(block)
    if fn_key == "формат_даты":
        return _normalize_format_date_block(block)
    if fn_key == "удалить_столбцы":
        return _normalize_delete_columns_block(block)
    if fn_key == "сетка":
        return _normalize_grid_block(block)
    if fn_key == "формат_столбцы":
        return _normalize_format_columns_block(block)
    if fn_key == "высота_строки":
        return _normalize_row_height_block(block)
    if fn_key == "отступ":
        return _normalize_indent_block(block)
    if fn_key == "шрифт":
        return _normalize_font_block(block)
    if fn_key == "ширина_столбцов":
        return _normalize_column_width_block(block)
    if fn_key in ("перенос", "перенос_и_авто_высота"):
        return _normalize_wrap_block(fn_key, block)
    if fn_key == "заголовок_плюс_высота":
        return _normalize_header_plus_height_block(block)
    if fn_key == "зебра_диапазон":
        return _normalize_zebra_block(block)
    if fn_key == "конкатенация_столбцов":
        return _normalize_concat_columns_block(block)
    if fn_key == "разделить_по_столбцам":
        return _normalize_split_by_columns_block(block)
    if fn_key == "условный_столбец":
        return _normalize_conditional_column_block(block)
    if fn_key == "столбец_по_лямбде":
        return _normalize_lambda_column_block(block)
    if fn_key == "добавить_столбец":
        return _normalize_add_column_block(block)
    if fn_key == "применить_формулу":
        return _normalize_apply_formula_block(block)
    if fn_key == "удаление_строк":
        return _normalize_delete_rows_block(block)
    if fn_key == "пропуск_строк_источника":
        return _normalize_skip_source_rows_block(block)
    if fn_key == "строка_заголовков":
        return _normalize_header_row_block(block)
    if fn_key == "группировка_по_столбцу":
        return _normalize_group_by_column_block(block)
    if fn_key == "стиль_печати":
        return _normalize_print_style_block(block)
    if fn_key == "заполнение_вниз":
        return _normalize_fill_down_block(block)
    if fn_key == "заполнить_вверх":
        return _normalize_fill_up_block(block)
    if fn_key == "развернуть_столбцы":
        return _normalize_unpivot_columns_block(block)
    if fn_key == "заполнение_вниз_вычислить":
        return _normalize_fill_down_calculate_block(block)
    if fn_key == "копировать_значения":
        return _normalize_copy_values_block(block)
    if fn_key == "замена_значений":
        return _normalize_replace_values_block(block)
    if fn_key == "текстовые_операции":
        return _normalize_normalize_text_block(block)
    if fn_key == "переименовать_лист":
        return _normalize_rename_sheet_block(block)
    if fn_key == "переименовать_столбцы":
        return _normalize_rename_columns_block(block)
    if fn_key == "переставить_столбцы":
        return _normalize_reorder_columns_block(block)
    if fn_key == "количество_значений":
        return _normalize_value_count_block(block)
    if fn_key == "удалить_дубликаты":
        return _normalize_remove_duplicates_block(block)
    if fn_key == "копировать_переместить_лист":
        return _normalize_copy_sheet_block(block)
    if fn_key == "активировать_лист":
        return _normalize_activate_sheet_block(block)
    if fn_key == "копирование_диапазонов":
        return _normalize_copy_ranges_block(block)
    if fn_key == "создать_лист":
        return _normalize_create_sheet_block(block)
    if fn_key == "отправить_по_почте":
        return _normalize_send_mail_block(block)
    if fn_key == "объединить_листы_в_один":
        return _normalize_merge_sheets_into_one_block(block)
    if fn_key == "разделить_листы":
        return _normalize_split_sheets_block(block)
    if fn_key == "преобразовать_числа":
        return _normalize_convert_numeric_block(block)
    if fn_key == "цвет_текста_по_значению":
        return _normalize_text_color_by_value_block(block)
    if fn_key == "удаление_верхних_строк":
        return _normalize_delete_top_rows_block(block)
    if fn_key == "ячейка_в_столбец":
        return _normalize_cell_to_column_block(block)
    if fn_key == "добавить_в_карту_переменных":
        return _normalize_add_to_variables_map_block(block)
    if fn_key == "ручной_ввод_в_карту_переменных":
        return _normalize_manual_variable_input_block(block)
    return _normalize_block(fn_key, block)
def _looks_like_json_payload(raw):
    s = unicode(raw or "").strip()
    return bool(s) and s[0] in "[{"
def _parse_json_relaxed(text):
    """
    Разбор JSON без libre_macros_pivot_lib (избегаем циклического импорта
    pivot_lib ↔ libre_macros_lib при первом открытии диалога в LO).
    """
    s = unicode(text or "").strip()
    if s == "":
        return None
    # Частый кейс из Calc: строка выглядит как JSON, но с удвоенными кавычками,
    # например: "[{""v"":1,""colors"":[""a"",""b""]}]".
    # Это невалидный JSON, но легко чинится заменой "" → " (после снятия внешних ").
    try:
        if len(s) >= 2 and s[0] == u"\"" and s[-1] == u"\"" and u"\"\"" in s:
            unq = unicode(s[1:-1])
            fixed_quotes = unq.replace(u"\"\"", u"\"").strip()
            if fixed_quotes and fixed_quotes[0] in u"[{":
                try:
                    return json.loads(fixed_quotes)
                except (ValueError, TypeError):
                    pass
    except Exception:
        pass
    try:
        return json.loads(s)
    except (ValueError, TypeError):
        pass
    # Если строка не обёрнута во внешние кавычки, но содержит "" внутри — тоже пробуем починить.
    if u"\"\"" in s and (u"{\"\"" in s or u"[{\"\"" in s):
        try:
            fixed_quotes = s.replace(u"\"\"", u"\"")
            return json.loads(fixed_quotes)
        except (ValueError, TypeError):
            pass
    try:
        import ast

        val = ast.literal_eval(s)
        if isinstance(val, (dict, list)):
            return val
    except Exception:
        pass
    fixed = s
    fixed = re.sub(r"\bTrue\b", "true", fixed)
    fixed = re.sub(r"\bFalse\b", "false", fixed)
    fixed = re.sub(r"\bNone\b", "null", fixed)
    fixed = re.sub(r",\s*}", "}", fixed)
    fixed = re.sub(r",\s*]", "]", fixed)
    try:
        return json.loads(fixed)
    except (ValueError, TypeError):
        return None
def _parse_json_root(raw_text):
    """Корень JSON: массив блоков; объект — один блок; «голый» объект сводной."""
    s = unicode(raw_text or "").strip()
    if s == "":
        return []
    if not _looks_like_json_payload(s):
        raise ValueError("not json")
    data = _parse_json_relaxed(s)
    if data is None:
        raise ValueError("json parse failed")
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        return [data]
    raise ValueError("JSON root must be array or object")
def _normalize_block(fn_key, block):
    if not isinstance(block, dict):
        raise ValueError("block must be object")
    # В LibreOffice (особенно со старыми Python/UNO) ключи dict иногда приходят как bytes/str
    # (python-literal через ast.literal_eval), а в коде здесь почти везде используются
    # unicode-ключи из-за unicode_literals. Приводим ключи к тексту, чтобы b.get("colors")
    # работал независимо от исходного типа ключа.
    out = {}
    for k, v in block.items():
        try:
            if isinstance(k, bytes):
                kk = k.decode("utf-8", "ignore")
            else:
                kk = unicode(k)
        except Exception:
            try:
                kk = str(k)
            except Exception:
                kk = u""
        out[kk] = v
    out["v"] = int(out.get("v") or PARAM_CODEC_VERSION)
    out["fn"] = normalize_fn_key(out.get("fn") or fn_key)
    return out
def _decode_json_blocks(fn_key, items):
    fn_key = normalize_fn_key(fn_key)
    if fn_key == "сводная_таблица":
        return _decode_pivot_json(items)
    if fn_key == "сортировка":
        return [_normalize_sort_block(b) for b in items]
    if fn_key == "раскрасить_блоки":
        return [_normalize_colorize_block(b) for b in items]
    if is_sheet_list_fn(fn_key):
        return [_normalize_sheet_list_block(fn_key, b) for b in items]
    if fn_key == "градиент":
        return [_normalize_gradient_block(b) for b in items]
    if fn_key == "подсветка_по_порогу":
        return [_normalize_threshold_block(b) for b in items]
    if fn_key == "формат_деньги":
        return [_normalize_format_money_block(b) for b in items]
    if fn_key == "формат_даты":
        return [_normalize_format_date_block(b) for b in items]
    if fn_key == "удалить_столбцы":
        return [_normalize_delete_columns_block(b) for b in items]
    if fn_key == "сетка":
        return [_normalize_grid_block(b) for b in items]
    if fn_key == "формат_столбцы":
        return [_normalize_format_columns_block(b) for b in items]
    if fn_key == "высота_строки":
        return [_normalize_row_height_block(b) for b in items]
    if fn_key == "отступ":
        return [_normalize_indent_block(b) for b in items]
    if fn_key == "шрифт":
        return [_normalize_font_block(b) for b in items]
    if fn_key == "ширина_столбцов":
        return [_normalize_column_width_block(b) for b in items]
    if fn_key in ("перенос", "перенос_и_авто_высота"):
        return [_normalize_wrap_block(fn_key, b) for b in items]
    if fn_key == "заголовок_плюс_высота":
        return [_normalize_header_plus_height_block(b) for b in items]
    if fn_key == "зебра_диапазон":
        return [_normalize_zebra_block(b) for b in items]
    if fn_key == "конкатенация_столбцов":
        return [_normalize_concat_columns_block(b) for b in items]
    if fn_key == "разделить_по_столбцам":
        return [_normalize_split_by_columns_block(b) for b in items]
    if fn_key == "условный_столбец":
        return [_normalize_conditional_column_block(b) for b in items]
    if fn_key == "столбец_по_лямбде":
        return [_normalize_lambda_column_block(b) for b in items]
    if fn_key == "добавить_столбец":
        return [_normalize_add_column_block(b) for b in items]
    if fn_key == "замена_значений":
        return [_normalize_replace_values_block(b) for b in items]
    if fn_key == "текстовые_операции":
        return [_normalize_normalize_text_block(b) for b in items]
    if fn_key == "применить_формулу":
        return [_normalize_apply_formula_block(b) for b in items]
    if fn_key == "удаление_строк":
        return [_normalize_delete_rows_block(b) for b in items]
    if fn_key == "пропуск_строк_источника":
        return [_normalize_skip_source_rows_block(b) for b in items]
    if fn_key == "строка_заголовков":
        return [_normalize_header_row_block(b) for b in items]
    if fn_key == "группировка_по_столбцу":
        return [_normalize_group_by_column_block(b) for b in items]
    if fn_key == "стиль_печати":
        return [_normalize_print_style_block(b) for b in items]
    if fn_key == "заполнение_вниз":
        return [_normalize_fill_down_block(b) for b in items]
    if fn_key == "заполнить_вверх":
        return [_normalize_fill_up_block(b) for b in items]
    if fn_key == "развернуть_столбцы":
        return [_normalize_unpivot_columns_block(b) for b in items]
    if fn_key == "заполнение_вниз_вычислить":
        return [_normalize_fill_down_calculate_block(b) for b in items]
    if fn_key == "копировать_значения":
        return [_normalize_copy_values_block(b) for b in items]
    if fn_key == "переименовать_лист":
        return [_normalize_rename_sheet_block(b) for b in items]
    if fn_key == "переименовать_столбцы":
        return [_normalize_rename_columns_block(b) for b in items]
    if fn_key == "переставить_столбцы":
        return [_normalize_reorder_columns_block(b) for b in items]
    if fn_key == "количество_значений":
        return [_normalize_value_count_block(b) for b in items]
    if fn_key == "удалить_дубликаты":
        return [_normalize_remove_duplicates_block(b) for b in items]
    if fn_key == "копировать_переместить_лист":
        return [_normalize_copy_sheet_block(b) for b in items]
    if fn_key == "активировать_лист":
        return [_normalize_activate_sheet_block(b) for b in items]
    if fn_key == "копирование_диапазонов":
        return [_normalize_copy_ranges_block(b) for b in items]
    if fn_key == "создать_лист":
        return [_normalize_create_sheet_block(b) for b in items]
    if fn_key == "отправить_по_почте":
        return [_normalize_send_mail_block(b) for b in items]
    if fn_key == "объединить_листы_в_один":
        return [_normalize_merge_sheets_into_one_block(b) for b in items]
    if fn_key == "разделить_листы":
        return [_normalize_split_sheets_block(b) for b in items]
    if fn_key == "преобразовать_числа":
        return [_normalize_convert_numeric_block(b) for b in items]
    if fn_key == "цвет_текста_по_значению":
        return [_normalize_text_color_by_value_block(b) for b in items]
    if fn_key == "удаление_верхних_строк":
        return [_normalize_delete_top_rows_block(b) for b in items]
    if fn_key == "ячейка_в_столбец":
        return [_normalize_cell_to_column_block(b) for b in items]
    if fn_key == "добавить_в_карту_переменных":
        return [_normalize_add_to_variables_map_block(b) for b in items]
    if fn_key == "ручной_ввод_в_карту_переменных":
        return [_normalize_manual_variable_input_block(b) for b in items]
    return [_normalize_block(fn_key, b) for b in items]
def _decode_pivot_json(items):
    if len(items) == 0:
        return [_pivot_empty_block()]
    if len(items) != 1:
        raise ValueError("сводная_таблица: ожидается один блок")
    return [_normalize_pivot_block(items[0])]
def _pivot_empty_block():
    return {"v": PARAM_CODEC_VERSION, "fn": "сводная_таблица"}
def _normalize_pivot_block(block):
    out = _pivot_empty_block()
    if not isinstance(block, dict):
        return out
    for key in _PIVOT_FIELDS:
        if key in block:
            out[key] = block[key]
    if block.get("as_values"):
        out["as_values"] = True
    if block.get("list_rows_expand"):
        out["list_rows_expand"] = True
    if "list_rows_include_names" in block:
        out["list_rows_include_names"] = bool(block.get("list_rows_include_names"))
    for sep_key in ("list_rows_name_sep", "list_rows_record_sep"):
        if sep_key in block and block.get(sep_key) is not None:
            out[sep_key] = unicode(block.get(sep_key))
    # Нормализация data_fields: LIST_ROWS + list_fields (синонимы).
    data_fields = out.get("data_fields")
    if isinstance(data_fields, (list, tuple)):
        alias_fn = {
            u"list_rows": u"LIST_ROWS",
            u"list": u"LIST_ROWS",
            u"список_строк": u"LIST_ROWS",
            u"детализация": u"LIST_ROWS",
            u"detail_rows": u"LIST_ROWS",
            u"detail": u"LIST_ROWS",
        }
        list_keys = (
            u"list_fields",
            u"detail_fields",
            u"detail_columns",
            u"поля_списка",
            u"колонки_списка",
        )
        norm_df = []
        has_list_rows = False
        for item in data_fields:
            if not isinstance(item, dict):
                continue
            field = unicode(item.get(u"field") or u"").strip()
            if field == u"":
                continue
            fn_raw = unicode(item.get(u"function") or u"SUM").strip()
            fn = alias_fn.get(fn_raw.casefold(), fn_raw.upper().replace(u"-", u"_"))
            if fn == u"LIST_ROWS":
                has_list_rows = True
            entry = {u"field": field, u"function": fn}
            list_fields = None
            for lk in list_keys:
                if lk not in item:
                    continue
                val = item.get(lk)
                if isinstance(val, (list, tuple)):
                    list_fields = [
                        unicode(x).strip()
                        for x in val
                        if unicode(x).strip()
                    ]
                else:
                    s = unicode(val or u"").strip()
                    if s:
                        list_fields = [
                            p.strip()
                            for p in s.replace(u";", u",").split(u",")
                            if p.strip()
                        ]
                break
            if list_fields:
                entry[u"list_fields"] = list_fields
            norm_df.append(entry)
        out[u"data_fields"] = norm_df
        if has_list_rows:
            out[u"as_values"] = True
    return out
def _normalize_sort_block(block):
    out = _normalize_block("сортировка", block)
    keys = out.get("keys")
    if keys is None:
        out["keys"] = []
    elif not isinstance(keys, list):
        out["keys"] = []
    norm_keys = []
    for item in out["keys"]:
        if not isinstance(item, dict):
            continue
        col = unicode(item.get("column") or "").strip()
        if col == "":
            continue
        norm_keys.append({"column": col, "desc": bool(item.get("desc"))})
    out["keys"] = norm_keys
    sheet = unicode(out.get("sheet") or "").strip()
    if sheet != "":
        out["sheet"] = sheet
    elif "sheet" in out:
        del out["sheet"]
    return out
def _normalize_colorize_block(block):
    out = _normalize_block("раскрасить_блоки", block)
    cols = out.get("key_columns")
    if cols is None:
        out["key_columns"] = []
    elif isinstance(cols, (list, tuple)):
        out["key_columns"] = [unicode(c).strip() for c in cols if unicode(c).strip()]
    else:
        out["key_columns"] = [unicode(cols).strip()] if unicode(cols).strip() else []
    colors = out.get("colors")
    if colors is None:
        out["colors"] = []
    elif isinstance(colors, (list, tuple)):
        out["colors"] = [unicode(c).strip() for c in colors if unicode(c).strip()]
    else:
        out["colors"] = [unicode(colors).strip()] if unicode(colors).strip() else []
    sort_val = out.get("sort")
    if sort_val in (None, "", "none"):
        out["sort"] = None
    elif unicode(sort_val).casefold() in ("asc", "+", "возр"):
        out["sort"] = "asc"
    elif unicode(sort_val).casefold() in ("desc", "-", "убыв"):
        out["sort"] = "desc"
    else:
        out["sort"] = None
    out["outline_blocks"] = bool(out.get("outline_blocks"))
    oc = out.get("outline_color")
    if oc in (None, "", "null"):
        out["outline_color"] = None
    else:
        out["outline_color"] = unicode(oc).strip()
    sheet = unicode(out.get("sheet") or "").strip()
    if sheet != "":
        out["sheet"] = sheet
    elif "sheet" in out:
        del out["sheet"]
    return out
def _normalize_sheet_list_block(fn_key, block):
    fn_key = normalize_fn_key(fn_key)
    out = _normalize_block(fn_key, block)
    sheets = out.get("sheets")
    if sheets is not None:
        if isinstance(sheets, (list, tuple)):
            norm = [unicode(s).strip() for s in sheets if unicode(s).strip()]
        else:
            single = unicode(sheets).strip()
            norm = [single] if single else []
        if norm:
            out["sheets"] = norm
        else:
            out.pop("sheets", None)
    else:
        out.pop("sheets", None)
    out.pop("sheet", None)
    filt = out.get("filter")
    if filt is None:
        filt = out.get("lambda") or out.get("code") or out.get("predicate")
    filt = unicode(filt or "").strip()
    out.pop("lambda", None)
    out.pop("code", None)
    out.pop("predicate", None)
    if filt != "":
        out["filter"] = filt
    else:
        out.pop("filter", None)
    if sheet_list_allows_convert(fn_key):
        if out.get("convert_existing"):
            out["convert_existing"] = True
        else:
            out.pop("convert_existing", None)
    else:
        out.pop("convert_existing", None)
    if fn_key == "автофильтр":
        action = unicode(
            out.get("action") or out.get("mode") or out.get("op") or "set"
        ).strip().casefold()
        if action in (
            "remove",
            "unset",
            "off",
            "disable",
            "убрать",
            "снять",
            "выкл",
            "-",
            "0",
            "false",
            "нет",
            "no",
        ):
            out["action"] = "remove"
        else:
            out["action"] = "set"
        smart_table = out.get("smart_table")
        if smart_table in (None, ""):
            smart_table = out.get("smart_range")
        if smart_table in (None, ""):
            smart_table = out.get("use_smart_table")
        smart_table = bool(smart_table)
        if smart_table:
            out["smart_table"] = True
            keep = out.get("keep_smart_table")
            if keep in (None, ""):
                keep = True
            out["keep_smart_table"] = bool(keep)
            style = unicode(
                out.get("table_style")
                or out.get("smart_style")
                or "TableStyleLight1"
            ).strip()
            out["table_style"] = style if style else "TableStyleLight1"
        else:
            out.pop("smart_table", None)
            out.pop("keep_smart_table", None)
            out.pop("table_style", None)
        out.pop("mode", None)
        out.pop("op", None)
        out.pop("smart_range", None)
        out.pop("use_smart_table", None)
        out.pop("smart_style", None)
    return out
def _normalize_delete_top_rows_block(block):
    """
    Блок для параметра «удаление_верхних_строк».
    Поля:
      - sheet: имя листа результата
      - n: число строк для удаления (int > 0)
    """
    out = _normalize_block("удаление_верхних_строк", block)
    # sheet
    sheet = unicode(out.get("sheet") or "").strip()
    if sheet != "":
        out["sheet"] = sheet
    else:
        out.pop("sheet", None)
    # n
    n_raw = out.get("n")
    if n_raw in (None, ""):
        n_raw = out.get("rows")
    if n_raw in (None, ""):
        n_raw = out.get("count")
    if n_raw in (None, ""):
        n_raw = out.get("value")
    try:
        n = int(n_raw)
    except Exception:
        n = 0
    if n > 0:
        out["n"] = n
    else:
        out.pop("n", None)
    # cleanup legacy
    out.pop("rows", None)
    out.pop("count", None)
    out.pop("value", None)
    return out


def _normalize_cell_to_column_block(block):
    """Блок «Ячейка_В_Столбец»: col, row (1-based), column, sheet, source_sheet."""
    out = _normalize_block("ячейка_в_столбец", block)
    col_s = unicode(out.get("col") or out.get("src_col") or "").strip()
    if col_s != "":
        out["col"] = col_s
    else:
        out.pop("col", None)
        out.pop("src_col", None)
    try:
        row_n = int(out.get("row", out.get("r", 1)))
    except Exception:
        row_n = 1
    if row_n > 0:
        out["row"] = row_n
    out.pop("r", None)
    column = unicode(out.get("column") or out.get("name") or "").strip()
    if column != "":
        out["column"] = column
    else:
        out.pop("column", None)
        out.pop("name", None)
    sheet = unicode(
        out.get("sheet") or out.get("sheet_name") or out.get("dest_sheet") or ""
    ).strip()
    if sheet != "":
        out["sheet"] = sheet
    else:
        out.pop("sheet", None)
    out.pop("sheet_name", None)
    out.pop("dest_sheet", None)
    source_sheet = unicode(
        out.get("source_sheet") or out.get("src_sheet") or out.get("sourceSheet") or ""
    ).strip()
    if source_sheet != "":
        out["source_sheet"] = source_sheet
    else:
        out.pop("source_sheet", None)
    out.pop("src_sheet", None)
    out.pop("sourceSheet", None)
    return out


def _normalize_add_to_variables_map_block(block):
    """Блок «Добавить_в_карту_переменных»: file, sheet, cell, name."""
    out = _normalize_block("добавить_в_карту_переменных", block)
    file_mask = unicode(
        out.get("file") or out.get("path") or out.get("source") or out.get("source_file") or ""
    ).strip()
    if file_mask != "":
        out["file"] = file_mask
    else:
        out.pop("file", None)
    out.pop("path", None)
    out.pop("source", None)
    out.pop("source_file", None)
    sheet_mask = unicode(
        out.get("sheet") or out.get("source_sheet") or out.get("src_sheet") or ""
    ).strip()
    if sheet_mask != "":
        out["sheet"] = sheet_mask
    else:
        out.pop("sheet", None)
    out.pop("source_sheet", None)
    out.pop("src_sheet", None)
    cell = unicode(out.get("cell") or out.get("addr") or out.get("address") or "").strip()
    if cell != "":
        out["cell"] = cell
    else:
        out.pop("cell", None)
    out.pop("addr", None)
    out.pop("address", None)
    name = unicode(out.get("name") or out.get("key") or out.get("var") or "").strip()
    if name != "":
        out["name"] = name
    else:
        out.pop("name", None)
    out.pop("key", None)
    out.pop("var", None)
    return out


def _normalize_manual_variable_input_block(block):
    """Блок «Ручной_ввод_в_карту_переменных»."""
    out = _normalize_block("ручной_ввод_в_карту_переменных", block)
    name = unicode(out.get("name") or out.get("key") or out.get("var") or "").strip()
    if name != "":
        out["name"] = name
    else:
        out.pop("name", None)
    out.pop("key", None)
    out.pop("var", None)
    vk = unicode(out.get("value_kind") or out.get("kind") or "scalar").strip().casefold()
    if vk in ("list", "список", "array"):
        out["value_kind"] = "list"
    else:
        out["value_kind"] = "scalar"
    out.pop("kind", None)
    dt = unicode(out.get("data_type") or out.get("type") or "text").strip().casefold()
    dt_map = {
        "text": "text",
        "str": "text",
        "string": "text",
        "текст": "text",
        "number": "number",
        "num": "number",
        "float": "number",
        "int": "number",
        "число": "number",
        "date": "date",
        "дата": "date",
        "datetime": "datetime",
        "дата_время": "datetime",
        "дата-время": "datetime",
        "bool": "bool",
        "boolean": "bool",
        "логическое": "bool",
        "логический": "bool",
    }
    out["data_type"] = dt_map.get(dt, "text")
    out.pop("type", None)
    fmt = unicode(out.get("format") or out.get("fmt") or "").strip()
    if fmt != "":
        out["format"] = fmt
    else:
        out.pop("format", None)
    out.pop("fmt", None)
    choices = out.get("choices")

    def _plain_choice_token(raw):
        t = unicode(raw or "").strip()
        if t == "":
            return u""
        try:
            bare, _quoted = unwrap_column_name_token(t)
            bare = unicode(bare or "").strip()
            if bare != "":
                return bare
        except Exception:
            pass
        if len(t) >= 2 and t[0] == t[-1] and t[0] in (u"'", u'"'):
            return t[1:-1].strip()
        return t

    if isinstance(choices, (list, tuple)):
        cleaned = []
        for x in choices:
            t = _plain_choice_token(x)
            if t != "":
                cleaned.append(t)
        if cleaned:
            out["choices"] = cleaned
        else:
            out.pop("choices", None)
    elif choices is not None and unicode(choices).strip() != "":
        parts = []
        try:
            for tok in split_column_name_tokens(unicode(choices)):
                t = _plain_choice_token(tok)
                if t != "":
                    parts.append(t)
        except Exception:
            parts = []
        if not parts:
            for chunk in unicode(choices).replace(";", ",").split(","):
                t = _plain_choice_token(chunk)
                if t != "":
                    parts.append(t)
        if parts:
            out["choices"] = parts
        else:
            out.pop("choices", None)
    else:
        out.pop("choices", None)
    prompt = unicode(out.get("prompt") or out.get("hint") or "").strip()
    if prompt != "":
        out["prompt"] = prompt
    else:
        out.pop("prompt", None)
    out.pop("hint", None)
    title = unicode(out.get("title") or "").strip()
    if title != "":
        out["title"] = title
    else:
        out.pop("title", None)
    if "required" in out:
        raw = out.get("required")
        if isinstance(raw, bool):
            out["required"] = raw
        else:
            s = unicode(raw).strip().casefold()
            out["required"] = s not in ("0", "false", "no", "нет", "n", "off")
    else:
        out["required"] = True

    def _norm_opt_bool(key, *aliases, default=None):
        raw = None
        if key in out:
            raw = out.get(key)
        else:
            for a in aliases:
                if a in out:
                    raw = out.get(a)
                    break
            else:
                if default is None:
                    return
                out[key] = default
                return
        if isinstance(raw, bool):
            out[key] = raw
        else:
            s = unicode(raw).strip().casefold()
            if s in ("",):
                out[key] = default if default is not None else False
            else:
                out[key] = s not in ("0", "false", "no", "нет", "n", "off")
        for a in aliases:
            out.pop(a, None)

    _norm_opt_bool("remember_last", "remember", "persist_last", default=True)
    _norm_opt_bool("encrypt_last", "encrypt_remembered", default=False)
    _norm_opt_bool("combined_dialog", "one_dialog", "single_dialog", default=False)
    if not out.get("remember_last", True):
        out["encrypt_last"] = False
    if "default" in out and out.get("default") is None:
        out.pop("default", None)
    # Без привязки к листу результата (один диалог на весь сбор).
    out.pop("sheet", None)
    out.pop("sheets", None)
    return out


def manual_variable_combined_dialog_enabled(blocks):
    """True, если блоки ручного ввода показываются одним диалогом (флаг на первом блоке)."""
    if not blocks:
        return False
    first = blocks[0]
    if not isinstance(first, dict):
        return False
    raw = first.get("combined_dialog")
    if isinstance(raw, bool):
        return raw
    s = unicode(raw or "").strip().casefold()
    if s in ("",):
        return False
    return s not in ("0", "false", "no", "нет", "n", "off")


def _sheet_block_base(fn_key, sheet_name):
    block = {"v": PARAM_CODEC_VERSION, "fn": normalize_fn_key(fn_key)}
    sheet = unicode(sheet_name or "").strip()
    if sheet != "":
        block["sheet"] = sheet
    return block
def _attach_sheet(block, sheet_name):
    sheet = unicode(sheet_name or "").strip()
    if sheet != "":
        block["sheet"] = sheet
    else:
        block.pop("sheet", None)
    return block
def _normalize_markers_list(value):
    """Маркеры/столбцы: split с учётом '…'; в JSON имена в кавычках."""
    return canonicalize_column_token_list_for_store(value)
def _normalize_gradient_block(block):
    out = _normalize_block("градиент", block)
    out["marker"] = unicode(out.get("marker") or "сумм").strip() or "сумм"
    out["color_min"] = unicode(out.get("color_min") or "зеленый").strip() or "зеленый"
    out["color_max"] = unicode(out.get("color_max") or "красный").strip() or "красный"
    wr = block.get("whole_row")
    if wr is True or wr is False:
        if wr:
            out["whole_row"] = True
    else:
        try:
            from libre_macros_lib import lm_parse_bool_param

            if lm_parse_bool_param(wr, default=False):
                out["whole_row"] = True
        except Exception:
            if wr:
                out["whole_row"] = True
    st = block.get("sort")
    if st is True:
        out["sort"] = True
    elif st is not False and st is not None:
        try:
            from libre_macros_lib import lm_parse_bool_param

            if lm_parse_bool_param(st, default=False):
                out["sort"] = True
        except Exception:
            if st:
                out["sort"] = True
    if out.get("sort"):
        sort_dir = unicode(block.get("sort_dir") or u"по_возр").strip().lower()
        if sort_dir in (u"по_убыв", u"убыв", u"desc", u"-", u"убывание"):
            out["sort_dir"] = u"по_убыв"
        else:
            out["sort_dir"] = u"по_возр"
    return _attach_sheet(out, out.get("sheet"))
def _normalize_threshold_block(block):
    out = _normalize_block("подсветка_по_порогу", block)
    # Столбцы: columns / markers / legacy marker
    columns = _normalize_column_token_list(out.get("columns"))
    if not columns:
        markers = _normalize_markers_list(out.get("markers"))
        if markers:
            columns = list(markers)
    marker = unicode(out.get("marker") or u"").strip()
    if not columns and marker:
        columns = [marker]
    if columns:
        out["columns"] = columns
        out["marker"] = unicode(columns[0])
    else:
        out.pop("columns", None)
        if marker:
            out["marker"] = marker
        else:
            out["marker"] = u"итог"

    def _opt_float(key):
        raw = out.get(key)
        if raw is None or raw is False:
            return None
        s = unicode(raw).strip().replace(u",", u".")
        if s == u"":
            return None
        try:
            return float(s)
        except (TypeError, ValueError):
            return None

    tmin = _opt_float("threshold_min")
    if tmin is None:
        tmin = _opt_float("min")
    tmax = _opt_float("threshold_max")
    if tmax is None:
        tmax = _opt_float("max")
    legacy = out.get("threshold")
    legacy_f = None
    if legacy is not None and legacy is not False:
        try:
            legacy_f = float(unicode(legacy).strip().replace(u",", u"."))
        except (TypeError, ValueError):
            legacy_f = None
    if tmin is None and tmax is None and legacy_f is not None:
        out["threshold"] = legacy_f
        out.pop("threshold_min", None)
        out.pop("threshold_max", None)
    else:
        if tmin is not None:
            out["threshold_min"] = tmin
        else:
            out.pop("threshold_min", None)
        if tmax is not None:
            out["threshold_max"] = tmax
        else:
            out.pop("threshold_max", None)
        out.pop("threshold", None)
    # Алиасы min/max не храним — только threshold_min/max (или legacy threshold).
    out.pop("min", None)
    out.pop("max", None)

    color = unicode(out.get("color") or out.get("fill_color") or u"").strip()
    if color:
        out["color"] = color
    else:
        out.pop("color", None)
    out.pop("fill_color", None)

    style_cols = _normalize_column_token_list(out.get("style_columns"))
    if style_cols:
        out["style_columns"] = style_cols
    else:
        out.pop("style_columns", None)

    fc = unicode(out.get("font_color") or u"").strip()
    auto = out.get("font_color_auto")
    if fc:
        out["font_color"] = fc
        out["font_color_auto"] = False
    else:
        out.pop("font_color", None)
        if auto is False or auto in (0, u"0", u"false", u"нет", u"no", u"-"):
            out["font_color_auto"] = False
        else:
            out["font_color_auto"] = True

    if out.get("bold"):
        out["bold"] = True
    else:
        out.pop("bold", None)
    if out.get("italic"):
        out["italic"] = True
    else:
        out.pop("italic", None)
    if out.get("whole_row"):
        out["whole_row"] = True
    else:
        out.pop("whole_row", None)
    return _attach_sheet(out, out.get("sheet"))
def _normalize_format_block_common(out, block):
    """Общие поля format/columns/markers/sheets для формат_деньги и формат_даты."""
    sheets = out.get("sheets")
    if sheets is not None:
        if isinstance(sheets, (list, tuple)):
            norm = [unicode(s).strip() for s in sheets if unicode(s).strip()]
        else:
            single = unicode(sheets).strip()
            norm = [single] if single else []
        if norm:
            out["sheets"] = norm
        else:
            out.pop("sheets", None)
    else:
        out.pop("sheets", None)
    columns = _normalize_column_token_list(out.get("columns"))
    if columns:
        out["columns"] = columns
    else:
        out.pop("columns", None)
    markers = _normalize_markers_list(out.get("markers"))
    if markers:
        out["markers"] = markers
    else:
        out.pop("markers", None)
    fmt = unicode(out.get("format") or u"").strip().replace(u"\xa0", u" ")
    if fmt:
        out["format"] = fmt
    else:
        out.pop("format", None)
    if out.get("convert_existing"):
        out["convert_existing"] = True
    else:
        out.pop("convert_existing", None)
    parse_formats = unicode(out.get("parse_formats") or u"").strip()
    if parse_formats:
        out["parse_formats"] = parse_formats
    else:
        out.pop("parse_formats", None)
    return out


def _normalize_format_money_block(block):
    out = _normalize_block("формат_деньги", block)
    out = _normalize_format_block_common(out, block)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_format_date_block(block):
    out = _normalize_block("формат_даты", block)
    out = _normalize_format_block_common(out, block)
    return _attach_sheet(out, out.get("sheet"))
def _normalize_delete_columns_block(block):
    out = _normalize_block("удалить_столбцы", block)
    markers = _normalize_markers_list(out.get("markers"))
    if markers:
        out["markers"] = markers
    else:
        out.pop("markers", None)
    return _attach_sheet(out, out.get("sheet"))
def _normalize_grid_block(block):
    out = _normalize_block("сетка", block)
    width = unicode(out.get("width") or "тонкая").strip()
    out["width"] = width if width != "" else "тонкая"
    color = unicode(out.get("color") or "").strip()
    if color != "":
        out["color"] = color
    else:
        out.pop("color", None)
    if out.get("include_header"):
        out["include_header"] = True
    else:
        out.pop("include_header", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_format_columns_column_single(text):
    """Один токен столбца правила → int / 'имя' / all (через canonicalize)."""
    t = unicode(text or u"").strip()
    if t == u"":
        return None
    low = t.casefold()
    if low in (u"all", u"все", u"*"):
        return u"all"
    tok = canonicalize_column_name_token_for_store(t)
    if tok == u"" or tok is None:
        return None
    if isinstance(tok, (str, unicode)) and unicode(tok).strip().casefold() in (
        u"all",
        u"все",
        u"*",
    ):
        return u"all"
    return tok


def _normalize_format_columns_column(col):
    """
    Селектор столбца(ов) правила формат_столбцы.
    Список имён — с '…' (exact); несколько токенов → list (правило на группу).
    """
    if col is None or isinstance(col, bool):
        return None
    if isinstance(col, int):
        return col
    if isinstance(col, (str, unicode)):
        low = unicode(col).strip().casefold()
        if low in (u"all", u"все", u"*"):
            return u"all"
    toks = canonicalize_column_token_list_for_store(col)
    if not toks:
        return None
    items = []
    for tok in toks:
        if isinstance(tok, (str, unicode)) and unicode(tok).strip().casefold() in (
            u"all",
            u"все",
            u"*",
        ):
            return u"all"
        if tok == u"" or tok is None:
            continue
        items.append(tok)
    if not items:
        return None
    if len(items) == 1:
        return items[0]
    return items


def _normalize_format_columns_rule(rule):
    if not isinstance(rule, dict):
        return None
    col_out = _normalize_format_columns_column(rule.get("column"))
    if col_out is None:
        return None
    fmt = unicode(rule.get("format") or "").strip()
    out = {"column": col_out}
    if fmt:
        out["format"] = fmt
    # styling
    fill = unicode(rule.get("fill") or rule.get("bg") or "").strip()
    if fill:
        out["fill"] = fill
    font = unicode(rule.get("font") or rule.get("font_color") or "").strip()
    if font:
        out["font"] = font
    if rule.get("font_auto") is True or rule.get("font_color_auto") is True:
        out["font_auto"] = True
    elif rule.get("font_auto") is False or rule.get("font_color_auto") is False:
        out["font_auto"] = False
    # border
    bw = unicode(rule.get("border_width") or rule.get("border") or "").strip()
    if bw:
        out["border_width"] = bw
    bc = unicode(rule.get("border_color") or "").strip()
    if bc:
        out["border_color"] = bc
    h_align = unicode(rule.get("h_align") or rule.get("align_h") or "").strip().lower()
    if h_align in ("left", "center", "right"):
        out["h_align"] = h_align
    v_align = unicode(rule.get("v_align") or rule.get("align_v") or "").strip().lower()
    if v_align in ("top", "center", "bottom"):
        out["v_align"] = v_align
    if rule.get("header_only") is True:
        out["header_only"] = True
    elif rule.get("include_header") is True:
        out["include_header"] = True
    return out
def _normalize_format_columns_block(block):
    out = _normalize_block("формат_столбцы", block)
    rules_in = out.get("rules")
    norm_rules = []
    if isinstance(rules_in, (list, tuple)):
        for item in rules_in:
            nr = _normalize_format_columns_rule(item)
            if nr is not None:
                norm_rules.append(nr)
    out["rules"] = norm_rules
    if out.get("convert_existing"):
        out["convert_existing"] = True
    else:
        out.pop("convert_existing", None)
    parse_formats = unicode(out.get("parse_formats") or u"").strip()
    if parse_formats:
        out["parse_formats"] = parse_formats
    else:
        out.pop("parse_formats", None)
    return _attach_sheet(out, out.get("sheet"))
def _normalize_row_height_block(block):
    out = _normalize_block("высота_строки", block)
    rows = out.get("rows")
    if rows in (None, "", "all"):
        out["rows"] = "all"
    elif isinstance(rows, (list, tuple)):
        out["rows"] = [int(r) for r in rows if str(r).strip() != ""]
    else:
        text = unicode(rows).strip()
        if text == "" or text.lower() in ("all", "все"):
            out["rows"] = "all"
        else:
            out["rows"] = [int(p.strip()) for p in text.split(DS_LIST) if p.strip()]
    try:
        out["height_mm"] = float(out.get("height_mm"))
    except (TypeError, ValueError):
        out["height_mm"] = 5.0
    return _attach_sheet(out, out.get("sheet"))
def _normalize_column_width_block(block):
    out = _normalize_block("ширина_столбцов", block)
    cols = out.get("columns")
    if cols in (None, "", "all"):
        out["columns"] = "all"
    else:
        toks = canonicalize_column_token_list_for_store(cols)
        if not toks:
            out["columns"] = "all"
        elif len(toks) == 1 and unicode(toks[0]).casefold() in (u"все", u"all", u"*"):
            out["columns"] = "all"
        else:
            cleaned = []
            for t in toks:
                if unicode(t).casefold() in (u"все", u"all", u"*"):
                    continue
                cleaned.append(t)
            out["columns"] = cleaned if cleaned else "all"
    try:
        out["width_mm"] = float(out.get("width_mm"))
    except (TypeError, ValueError):
        out["width_mm"] = 30.0
    return _attach_sheet(out, out.get("sheet"))
def _normalize_indent_block(block):
    out = _normalize_block("отступ", block)
    align = unicode(out.get("h_align") or "left").strip().lower()
    if align not in ("left", "center", "right"):
        align = "left"
    out["h_align"] = align
    try:
        out["steps"] = max(0, int(out.get("steps", 1)))
    except (TypeError, ValueError):
        out["steps"] = 1
    cols = out.get("columns")
    if cols in (None, "", []):
        out.pop("columns", None)
    elif isinstance(cols, (list, tuple)):
        out["columns"] = [unicode(c).strip() for c in cols if unicode(c).strip()]
    else:
        out["columns"] = [p.strip() for p in unicode(cols).split(DS_LIST) if p.strip()]
    return _attach_sheet(out, out.get("sheet"))
def _normalize_font_block(block):
    out = _normalize_block("шрифт", block)
    name = out.get("name")
    if name in (None, u"", u"-", u"—"):
        out.pop("name", None)
    else:
        out["name"] = unicode(name).strip()
    for key in ("size_data", "size_header"):
        val = out.get(key)
        if val in (None, u"", u"-", u"—"):
            out.pop(key, None)
        else:
            try:
                out[key] = float(val)
            except (TypeError, ValueError):
                out.pop(key, None)
    return _attach_sheet(out, out.get("sheet"))
def _normalize_wrap_block(fn_key, block):
    out = _normalize_block(fn_key, block)
    if out.get("wrap") is False:
        out["wrap"] = False
    else:
        out["wrap"] = True
    if out.get("include_header"):
        out["include_header"] = True
    else:
        out.pop("include_header", None)
    return _attach_sheet(out, out.get("sheet"))
def _normalize_header_plus_height_block(block):
    out = _normalize_block("заголовок_плюс_высота", block)
    try:
        out["height_mm"] = float(out.get("height_mm"))
    except (TypeError, ValueError):
        out["height_mm"] = 12.7
    # Минимум 7 мм — защита от схлопывания (и при абсолюте, и при плюсе).
    try:
        if float(out["height_mm"]) < 0:
            out["height_mm"] = 0.0
    except (TypeError, ValueError):
        out["height_mm"] = 12.7
    out["add_height"] = bool(out.get("add_height"))
    hori = unicode(out.get("h_align") or "center").strip().lower() or "center"
    vert = unicode(out.get("v_align") or "center").strip().lower() or "center"
    out["h_align"] = hori
    out["v_align"] = vert
    out["bold"] = bool(out.get("bold"))
    font_name = unicode(out.get("font_name") or u"").strip()
    out["font_name"] = font_name if font_name else u""
    font_size = out.get("font_size")
    if font_size in (u"", u"-", u"—", None):
        out["font_size"] = None
    else:
        try:
            out["font_size"] = float(font_size)
        except (TypeError, ValueError):
            out["font_size"] = None
    out["fill_color"] = unicode(out.get("fill_color") or u"").strip()
    font_color = unicode(out.get("font_color") or u"").strip()
    out["font_color"] = font_color
    if font_color:
        out["font_color_auto"] = bool(out.get("font_color_auto", False))
    else:
        out["font_color_auto"] = bool(out.get("font_color_auto", True))
    return _attach_sheet(out, out.get("sheet"))
def _zebra_role_labels():
    return (
        ("header", "header"),
        ("even", "even"),
        ("odd", "odd"),
    )
def _normalize_zebra_block(block):
    out = _normalize_block("зебра_диапазон", block)
    roles = out.get("roles")
    if isinstance(roles, dict):
        norm_roles = {}
        for role_key, label in _zebra_role_labels():
            spec = roles.get(role_key) or roles.get(label)
            if not isinstance(spec, dict):
                continue
            fill = unicode(spec.get("fill") or "").strip()
            font = unicode(spec.get("font") or "").strip()
            if fill or font:
                norm_roles[role_key] = {"fill": fill, "font": font}
        if norm_roles:
            out["roles"] = norm_roles
        else:
            out.pop("roles", None)
    return _attach_sheet(out, out.get("sheet"))
def _normalize_int_list(value):
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        out = []
        for item in value:
            if isinstance(item, int):
                out.append(item)
            else:
                s = unicode(item).strip()
                if s.isdigit():
                    out.append(int(s))
        return out
    text = unicode(value).strip()
    if text == "":
        return []
    out = []
    for part in re.split(r"[,;]", text):
        part = part.strip()
        if part.isdigit():
            out.append(int(part))
    return out
def _normalize_column_token_list(value):
    """Список столбцов для JSON: кавычки как ввели; цифры → int, буквы A без кавычек."""
    return canonicalize_column_token_list_for_store(value)


def _normalize_concat_columns_block(block):
    out = _normalize_block("конкатенация_столбцов", block)
    out["new_column"] = unicode(out.get("new_column") or "").strip()
    out["separator"] = unicode(out.get("separator") or "")
    # columns: разрешаем индексы (1-based), буквы и заголовки
    out["columns"] = _normalize_column_token_list(out.get("columns"))
    return _attach_sheet(out, out.get("sheet"))


def _normalize_split_by_columns_block(block):
    out = _normalize_block("разделить_по_столбцам", block)
    col = unicode(out.get("column") or "").strip()
    if col != "":
        out["column"] = col
    else:
        out.pop("column", None)
    delim = unicode(out.get("delimiter") or out.get("custom_delimiter") or "").strip()
    if delim != "":
        out["delimiter"] = delim
    else:
        out.pop("delimiter", None)
    pos = unicode(out.get("position") or "_после_").strip()
    if pos == "":
        pos = "_после_"
    out["position"] = pos
    rel = unicode(out.get("relative_column") or "").strip()
    if rel != "":
        out["relative_column"] = rel
    else:
        out.pop("relative_column", None)
    sp = unicode(out.get("source_policy") or "replace").strip().casefold()
    out["source_policy"] = "keep" if sp == "keep" else "replace"
    sm = unicode(out.get("split_mode") or "each").strip().casefold()
    out["split_mode"] = "leftmost" if sm == "leftmost" else "each"
    nt = unicode(out.get("non_text") or "skip").strip().casefold()
    if nt not in ("skip", "empty", "text", "fail"):
        nt = "skip"
    out["non_text"] = nt
    try:
        out["max_columns"] = int(out.get("max_columns") or 0)
    except (TypeError, ValueError):
        out["max_columns"] = 0
    if out["max_columns"] < 0:
        out["max_columns"] = 0
    out["trim_parts"] = bool(out.get("trim_parts", True))
    raw_names = out.get("new_names")
    names = []
    if isinstance(raw_names, (list, tuple)):
        for item in raw_names:
            s = unicode("" if item is None else item).strip()
            if s != "":
                names.append(s)
    elif raw_names is not None:
        s = unicode(raw_names).strip()
        if s != "":
            names = [p.strip() for p in re.split(r"[,;]", s) if p.strip()]
    if names:
        out["new_names"] = names
    else:
        out.pop("new_names", None)
    out.pop("custom_delimiter", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_conditional_branch(branch):
    if not isinstance(branch, dict):
        return {"rules": [], "otherwise": ""}

    out = {}
    rules_in = branch.get("rules")
    rules_out = []
    if isinstance(rules_in, (list, tuple)):
        for rule in rules_in:
            if not isinstance(rule, dict):
                continue
            row = {}
            col = unicode(rule.get("column") or "").strip()
            op = unicode(rule.get("op") or "").strip()
            if col == "" or op == "":
                continue
            row["column"] = col
            row["op"] = op

            if "value_column" in rule and unicode(rule.get("value_column") or "").strip() != "":
                row["value_column"] = unicode(rule.get("value_column") or "").strip()
            elif "value" in rule:
                row["value"] = rule.get("value")

            if "then_column" in rule and unicode(rule.get("then_column") or "").strip() != "":
                row["then_column"] = unicode(rule.get("then_column") or "").strip()
            elif "then" in rule:
                row["then"] = rule.get("then")
            else:
                row["then"] = ""

            rules_out.append(row)
    out["rules"] = rules_out

    ow = branch.get("otherwise")
    if isinstance(ow, dict) and "rules" in ow:
        out["otherwise"] = _normalize_conditional_branch(ow)
    elif "otherwise" in branch:
        out["otherwise"] = ow
    else:
        out["otherwise"] = ""

    ow_col = unicode(branch.get("otherwise_column") or "").strip()
    if ow_col != "":
        out["otherwise_column"] = ow_col
    return out


def _normalize_conditional_column_block(block):
    out = _normalize_block("условный_столбец", block)
    out["new_column"] = unicode(out.get("new_column") or "").strip()
    pos = unicode(out.get("position") or "_конец_").strip()
    if pos == "":
        pos = "_конец_"
    out["position"] = pos
    rel = unicode(out.get("relative_column") or "").strip()
    if rel != "":
        out["relative_column"] = rel
    else:
        out.pop("relative_column", None)
    out["output_policy"] = unicode(out.get("output_policy") or "new").strip().casefold()
    if out["output_policy"] not in ("new", "replace"):
        out["output_policy"] = "new"
    out["trim"] = bool(out.get("trim", True))
    out["case_sensitive"] = bool(out.get("case_sensitive", False))
    cmp_as = unicode(out.get("compare_as") or "auto").strip().casefold()
    if cmp_as not in ("auto", "text", "number", "date"):
        cmp_as = "auto"
    out["compare_as"] = cmp_as
    blank_left = unicode(out.get("blank_left") or "as_empty").strip().casefold()
    if blank_left not in ("as_empty", "skip_rule", "fail"):
        blank_left = "as_empty"
    out["blank_left"] = blank_left
    non_comp = unicode(out.get("non_comparable") or "false").strip().casefold()
    if non_comp not in ("false", "fail", "true"):
        non_comp = "false"
    out["non_comparable"] = non_comp

    if "as_values" in block:
        try:
            from libre_macros_lib import lm_parse_bool_param

            parsed = lm_parse_bool_param(block.get("as_values"), default=None)
            if parsed is True:
                out["as_values"] = True
            elif parsed is False:
                out["as_values"] = False
        except Exception:
            out["as_values"] = bool(block.get("as_values"))
    else:
        out["as_values"] = True

    rules = out.get("rules")
    if isinstance(rules, str):
        parsed = _parse_json_relaxed(rules)
        rules = parsed if isinstance(parsed, list) else []
    if not isinstance(rules, (list, tuple)):
        rules = []

    normalized_branch = _normalize_conditional_branch(
        {
            "rules": rules,
            "otherwise": (
                _parse_json_relaxed(out.get("otherwise"))
                if isinstance(out.get("otherwise"), str)
                and unicode(out.get("otherwise") or "").strip().startswith("{")
                else out.get("otherwise", "")
            ),
            "otherwise_column": out.get("otherwise_column"),
        }
    )
    out["rules"] = normalized_branch.get("rules", [])
    out["otherwise"] = normalized_branch.get("otherwise", "")
    if "otherwise_column" in normalized_branch:
        out["otherwise_column"] = normalized_branch["otherwise_column"]
    else:
        out.pop("otherwise_column", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_lambda_column_block(block):
    """
    Нормализовать JSON-блок `столбец_по_лямбде`.

    Лямбды хранятся как сырой текст (санитизация/компиляция в executor).
    """
    out = _normalize_block("столбец_по_лямбде", block)

    out["expr"] = unicode(out.get("expr") or "").strip()
    out["new_column"] = unicode(out.get("new_column") or "").strip()
    out["new_column_expr"] = unicode(out.get("new_column_expr") or "").strip()

    pos = unicode(out.get("position") or "_конец_").strip()
    if pos == u"":
        pos = "_конец_"
    out["position"] = pos

    rel = unicode(out.get("relative_column") or "").strip()
    if rel != u"":
        out["relative_column"] = rel
    else:
        out.pop("relative_column", None)

    out["output_policy"] = unicode(out.get("output_policy") or "new").strip().casefold()
    if out["output_policy"] not in ("new", "replace"):
        out["output_policy"] = "new"

    out["on_error"] = unicode(out.get("on_error") or "keep").strip().casefold()
    if out["on_error"] not in ("empty", "keep", "fail"):
        out["on_error"] = "keep"

    rt = unicode(out.get("result_type") or "auto").strip().casefold()
    if rt in ("text", u"текст", u"строка"):
        rt = "text"
    elif rt in ("number", u"число", u"numeric", u"числовой"):
        rt = "number"
    else:
        rt = "auto"
    out["result_type"] = rt

    sheet_expr = unicode(out.get("sheet_expr") or "").strip()
    if sheet_expr != u"":
        out["sheet_expr"] = sheet_expr
    else:
        out.pop("sheet_expr", None)

    return _attach_sheet(out, out.get("sheet"))


def _normalize_add_column_block(block):
    """Нормализовать JSON-блок `добавить_столбец`."""
    out = _normalize_block("добавить_столбец", block)
    out["new_column"] = unicode(out.get("new_column") or out.get("name") or "").strip()
    if out["new_column"] == u"":
        out.pop("new_column", None)

    dt = unicode(out.get("data_type") or out.get("type") or u"text").strip().casefold()
    dt_map = {
        u"text": u"text",
        u"str": u"text",
        u"string": u"text",
        u"текст": u"text",
        u"number": u"number",
        u"num": u"number",
        u"numeric": u"number",
        u"число": u"number",
        u"date": u"date",
        u"дата": u"date",
    }
    out["data_type"] = dt_map.get(dt, u"text")

    fmt = unicode(out.get("format") or "").strip()
    if fmt == u"" or fmt.casefold() == u"(глобальный)".casefold():
        out.pop("format", None)
    else:
        out["format"] = fmt

    choices = _normalize_column_token_list(out.get("choices") or out.get("values") or out.get("list"))
    # для списка значений кавычки не нужны — голые токены
    cleaned = []
    for c in choices or ():
        bare, _q = unwrap_column_name_token(c)
        bare = unicode(bare or u"").strip()
        if bare != u"":
            cleaned.append(bare)
    if out["data_type"] == u"text" and cleaned:
        out["choices"] = cleaned
    else:
        out.pop("choices", None)

    def _norm_bool(key, *aliases):
        v = None
        if key in block:
            v = block.get(key)
        else:
            for a in aliases:
                if a in block:
                    v = block.get(a)
                    break
        if v is None:
            return None
        try:
            from libre_macros_lib import lm_parse_bool_param

            return lm_parse_bool_param(v, default=None)
        except Exception:
            return bool(v)

    dv = _norm_bool(
        "data_validation",
        "validation",
        "контроль_данных",
        "validate",
    )
    if out["data_type"] in (u"number", u"date") and dv is True:
        out["data_validation"] = True
    else:
        out.pop("data_validation", None)

    min_v = unicode(out.get("min") or out.get("min_value") or u"").strip()
    max_v = unicode(out.get("max") or out.get("max_value") or u"").strip()
    if out["data_type"] in (u"number", u"date") and out.get("data_validation"):
        if min_v != u"":
            out["min"] = min_v
        else:
            out.pop("min", None)
        if max_v != u"":
            out["max"] = max_v
        else:
            out.pop("max", None)
    else:
        out.pop("min", None)
        out.pop("max", None)

    return _attach_sheet(out, out.get("sheet"))


def _normalize_delete_rows_column_list(value):
    """Список столбцов для удаления строк: имена в '…' (как columns_pick)."""
    return canonicalize_column_token_list_for_store(value)
def _strip_leading_formula_eq(text):
    """Убрать один ведущий '=' (и пробелы) — для IF(...)/хранения без двойного '='."""
    s = unicode(text or "").strip()
    if s.startswith("="):
        s = s[1:].strip()
    return s


def _normalize_apply_formula_block(block):
    out = _normalize_block("применить_формулу", block)
    out["column"] = unicode(out.get("column") or "").strip()
    out["formula"] = unicode(out.get("formula") or "").strip()
    fmt = unicode(out.get("format") or "").strip()
    if fmt == "" or fmt.casefold() == u"(глобальный)".casefold():
        out.pop("format", None)
    else:
        out["format"] = fmt
    if "as_values" in block:
        try:
            from libre_macros_lib import lm_parse_bool_param

            parsed = lm_parse_bool_param(block.get("as_values"), default=None)
            if parsed is True:
                out["as_values"] = True
            elif parsed is False:
                out["as_values"] = False
        except Exception:
            if block.get("as_values"):
                out["as_values"] = True
            else:
                out["as_values"] = False
    return _attach_sheet(out, out.get("sheet"))
def _normalize_delete_rows_block(block):
    out = _normalize_block("удаление_строк", block)
    mode = unicode(out.get("mode") or "columns").strip().casefold()
    if mode == "formula":
        out["mode"] = "formula"
        out["formula"] = _strip_leading_formula_eq(out.get("formula") or "")
        out.pop("columns", None)
    elif mode in ("none", ""):
        out["mode"] = "none"
        out.pop("columns", None)
        out.pop("formula", None)
    else:
        out["mode"] = "columns"
        out["columns"] = _normalize_delete_rows_column_list(out.get("columns"))
        out.pop("formula", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_skip_source_rows_block(block):
    """
    Пропуск_строк_источника: sheet = имя листа *источника* (не результата).
    columns — номера/буквы/заголовки; без sheet — для всех листов файла.
    """
    out = _normalize_block("пропуск_строк_источника", block)
    mode = unicode(out.get("mode") or "columns").strip().casefold()
    if mode == "formula":
        out["mode"] = "formula"
        out["formula"] = _strip_leading_formula_eq(out.get("formula") or "")
        out.pop("columns", None)
    elif mode in ("none", ""):
        out["mode"] = "none"
        out.pop("columns", None)
        out.pop("formula", None)
    else:
        out["mode"] = "columns"
        out["columns"] = _normalize_delete_rows_column_list(out.get("columns"))
        out.pop("formula", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_header_row_block(block):
    """
    Расширенный «Строка_Заголовков»: rows (2-5 / 2:5) и separator (по умолчанию пробел).
    """
    out = _normalize_block("строка_заголовков", block)
    rows_raw = out.get("rows")
    if rows_raw in (None, u""):
        a = out.get("row_from")
        b = out.get("row_to")
        if a not in (None, u"") or b not in (None, u""):
            try:
                from libre_macros_header_lib import format_header_rows_range

                fa = int(a if a not in (None, u"") else b)
                fb = int(b if b not in (None, u"") else a)
                rows_raw = format_header_rows_range(fa, fb)
            except Exception:
                rows_raw = u""
    try:
        from libre_macros_header_lib import (
            format_header_rows_range,
            normalize_header_separator,
            parse_header_rows_range,
        )

        pair = parse_header_rows_range(rows_raw)
        if pair is None:
            out["rows"] = u"1-1"
        else:
            out["rows"] = format_header_rows_range(pair[0], pair[1])
        sep = out.get("separator")
        if sep is None:
            sep = out.get("sep")
        out["separator"] = normalize_header_separator(sep)
    except Exception:
        s = unicode(rows_raw or u"").strip()
        out["rows"] = s if s else u"1-1"
        if "separator" not in out:
            out["separator"] = u" "
    out.pop("sep", None)
    out.pop("row_from", None)
    out.pop("row_to", None)
    out.pop("from", None)
    out.pop("to", None)
    sheet = unicode(out.get("sheet") or "").strip()
    if sheet != "":
        out["sheet"] = sheet
    else:
        out.pop("sheet", None)
    return out


def _normalize_group_agg_fn_code(raw_fn):
    """Подпись/код функции итога → sum|count|…"""
    try:
        from libre_macros_param_wizard_cfg import (
            GROUP_AGG_FN_CODE,
            GROUP_AGG_FN_CHOICES,
        )

        raw = unicode(raw_fn or u"").strip()
        if raw == u"":
            return u"sum"
        code = GROUP_AGG_FN_CODE.get(raw.casefold(), u"")
        if not code:
            code = GROUP_AGG_FN_CODE.get(raw.casefold().replace(u" ", u"_"), u"")
        if not code:
            for c, _lab in GROUP_AGG_FN_CHOICES:
                if c.casefold() == raw.casefold():
                    code = c
                    break
        return code if code else u"sum"
    except Exception:
        fn = unicode(raw_fn or u"sum").strip().casefold()
        if fn in (u"sum", u"count", u"average", u"max", u"min", u"product"):
            return fn
        return u"sum"


def _normalize_group_agg_entry(item):
    """Один уровень итога: {fn, columns}."""
    if not isinstance(item, dict):
        return None
    fn = _normalize_group_agg_fn_code(item.get("fn") or item.get("agg_fn"))
    cols = canonicalize_column_token_list_for_store(
        item.get("columns") if item.get("columns") is not None else item.get("agg_columns")
    )
    if not cols:
        one = item.get("column") or item.get("agg_column")
        if one not in (None, u""):
            cols = canonicalize_column_token_list_for_store(one)
    if not cols:
        return None
    return {u"fn": fn, u"columns": cols}


def _normalize_group_by_column_block(block):
    out = _normalize_block("группировка_по_столбцу", block)
    toks = canonicalize_column_token_list_for_store(out.get("marker"))
    if toks:
        out["marker"] = toks[0]
    else:
        out["marker"] = u"группа"
    aggs = []
    raw_aggs = out.get("aggs")
    if isinstance(raw_aggs, (list, tuple)):
        for item in raw_aggs:
            ent = _normalize_group_agg_entry(item if isinstance(item, dict) else {})
            if ent is not None:
                aggs.append(ent)
    if not aggs:
        # Один уровень из agg_fn + agg_columns (или legacy agg_column).
        fn = _normalize_group_agg_fn_code(out.get("agg_fn"))
        cols = canonicalize_column_token_list_for_store(out.get("agg_columns"))
        if not cols:
            one = out.get("agg_column")
            if one not in (None, u""):
                cols = canonicalize_column_token_list_for_store(one)
        if cols:
            aggs.append({u"fn": fn, u"columns": cols})
        out["agg_fn"] = fn
        out["agg_columns"] = cols
    else:
        # Зеркало первого уровня для совместимости со старым UI/логами.
        out["agg_fn"] = aggs[0]["fn"]
        out["agg_columns"] = list(aggs[0]["columns"])
    out["aggs"] = aggs
    out.pop("agg_column", None)
    try:
        from libre_macros_lib import lm_parse_bool_param

        out["sort"] = bool(lm_parse_bool_param(out.get("sort"), default=True))
        out["outline"] = bool(lm_parse_bool_param(out.get("outline"), default=True))
        out["grand_total"] = bool(
            lm_parse_bool_param(out.get("grand_total"), default=True)
        )
    except Exception:
        out["sort"] = True if out.get("sort") is None else bool(out.get("sort"))
        out["outline"] = True if out.get("outline") is None else bool(out.get("outline"))
        out["grand_total"] = (
            True if out.get("grand_total") is None else bool(out.get("grand_total"))
        )
    return _attach_sheet(out, out.get("sheet"))


def merge_group_by_column_blocks(blocks):
    """
    Склеить подряд идущие блоки группировка_по_столбцу (один marker/sheet)
    в один блок с несколькими уровнями aggs[].
    """
    items = []
    for b in blocks or []:
        if isinstance(b, dict):
            items.append(_normalize_group_by_column_block(b))
    if not items:
        return None
    base = dict(items[0])
    aggs = []
    sort_any = False
    outline_any = False
    grand_any = False
    for b in items:
        sort_any = sort_any or bool(b.get("sort", True))
        outline_any = outline_any or bool(b.get("outline", True))
        grand_any = grand_any or bool(b.get("grand_total", True))
        for a in b.get("aggs") or []:
            if isinstance(a, dict) and a.get("columns"):
                aggs.append(
                    {
                        u"fn": _normalize_group_agg_fn_code(a.get("fn")),
                        u"columns": list(a.get("columns") or []),
                    }
                )
    base["aggs"] = aggs
    if aggs:
        base["agg_fn"] = aggs[0]["fn"]
        base["agg_columns"] = list(aggs[0]["columns"])
    base["sort"] = sort_any
    base["outline"] = outline_any
    base["grand_total"] = grand_any
    return _normalize_group_by_column_block(base)
def _normalize_convert_numeric_block(block):
    out = _normalize_block("преобразовать_числа", block)
    cols = out.get("columns")
    if cols in (None, "", []):
        out["columns"] = []
    else:
        out["columns"] = _normalize_column_token_list(cols)
    return _attach_sheet(out, out.get("sheet"))
def _normalize_text_color_by_value_block(block):
    out = _normalize_block("цвет_текста_по_значению", block)
    m = unicode(out.get("marker") or u"").strip()
    out["marker"] = m if m != u"" else u"статус"
    return _attach_sheet(out, out.get("sheet"))
def _normalize_print_style_block(block):
    out = _normalize_block("стиль_печати", block)
    orient = unicode(out.get("orientation") or "landscape").strip().casefold()
    if orient not in ("landscape", "portrait"):
        orient = "landscape"
    out["orientation"] = orient
    fit = out.get("fit_pages")
    try:
        out["fit_pages"] = int(fit) if fit is not None else 1
    except (TypeError, ValueError):
        out["fit_pages"] = 1
    return _attach_sheet(out, out.get("sheet"))
def _normalize_fill_down_block(block):
    out = _normalize_block("заполнение_вниз", block)
    out["columns"] = _normalize_column_token_list(out.get("columns"))
    return _attach_sheet(out, out.get("sheet"))


def _normalize_fill_up_block(block):
    out = _normalize_block("заполнить_вверх", block)
    out["columns"] = _normalize_column_token_list(out.get("columns"))
    try:
        from libre_macros_lib import lm_parse_bool_param

        out["skip_if_above_not_empty"] = bool(
            lm_parse_bool_param(out.get("skip_if_above_not_empty"), default=True)
        )
        out["drop_trailing_empty"] = bool(
            lm_parse_bool_param(out.get("drop_trailing_empty"), default=False)
        )
    except Exception:
        out["skip_if_above_not_empty"] = True
        out["drop_trailing_empty"] = False
    return _attach_sheet(out, out.get("sheet"))


def _normalize_unpivot_columns_block(block):
    out = _normalize_block("развернуть_столбцы", block)
    out["unpivot_columns"] = _normalize_column_token_list(
        out.get("unpivot_columns") or out.get("columns")
    )
    out["exclude_columns"] = _normalize_column_token_list(out.get("exclude_columns"))
    attr = unicode(out.get("attribute_column") or u"").strip()
    out["attribute_column"] = attr if attr else u"Атрибут"
    val = unicode(out.get("value_column") or u"").strip()
    out["value_column"] = val if val else u"Значение"
    output = unicode(out.get("output") or u"inplace").strip().casefold()
    if output in (u"new_sheet", u"новый_лист", u"лист"):
        out["output"] = u"new_sheet"
    else:
        out["output"] = u"inplace"
    dest = unicode(out.get("dest_sheet") or out.get("dest") or u"").strip()
    if dest:
        out["dest_sheet"] = dest
    else:
        out.pop("dest_sheet", None)
        out.pop("dest", None)
    try:
        from libre_macros_lib import lm_parse_bool_param

        out["drop_empty_rows"] = bool(
            lm_parse_bool_param(out.get("drop_empty_rows"), default=False)
        )
        if "as_values" in out:
            out["as_values"] = bool(
                lm_parse_bool_param(out.get("as_values"), default=True)
            )
        else:
            out["as_values"] = True
    except Exception:
        out["drop_empty_rows"] = bool(out.get("drop_empty_rows"))
        out["as_values"] = True if "as_values" not in out else bool(out.get("as_values"))
    out.pop("skip_header_row", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_fill_down_calculate_rule(rule):
    if not isinstance(rule, dict):
        return None
    col = rule.get("column")
    if col is None:
        return None
    # Список из визарда (накопление) → первый токен; имена в '…'.
    toks = canonicalize_column_token_list_for_store(col)
    if not toks:
        return None
    col_out = toks[0]
    formula = unicode(rule.get("formula") or "").strip()
    if formula == "":
        return None
    out = {"column": col_out, "formula": formula}
    if rule.get("expand_to_right"):
        out["expand_to_right"] = True
    return out
def _normalize_fill_down_calculate_block(block):
    out = _normalize_block("заполнение_вниз_вычислить", block)
    cols = _normalize_column_token_list(out.get("columns"))
    if len(cols) > 0:
        out["columns"] = cols
    else:
        out.pop("columns", None)
    formula = unicode(out.get("formula") or "").strip()
    if formula:
        out["formula"] = formula
    else:
        out.pop("formula", None)
    rules_in = out.get("rules")
    norm_rules = []
    if isinstance(rules_in, (list, tuple)):
        for item in rules_in:
            nr = _normalize_fill_down_calculate_rule(item)
            if nr is not None:
                norm_rules.append(nr)
    out["rules"] = norm_rules
    return _attach_sheet(out, out.get("sheet"))
def _normalize_copy_values_block(block):
    out = _normalize_block("копировать_значения", block)
    if out.get("whole_sheet"):
        out["whole_sheet"] = True
        out["columns"] = []
    else:
        out.pop("whole_sheet", None)
        cols = _normalize_column_token_list(out.get("columns"))
        out["columns"] = cols
        if len(cols) == 0:
            out.pop("columns", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_replace_values_block(block):
    out = _normalize_block("замена_значений", block)
    cols = _normalize_column_token_list(out.get("columns"))
    out["columns"] = cols

    def _norm_str_list(val, keep_empty_quoted=False, keep_quotes=False):
        return split_sep_list_respecting_quotes(
            val, keep_empty_quoted=keep_empty_quoted, keep_quotes=keep_quotes
        )

    find_list = _norm_str_list(
        out.get("find") or out.get("search") or out.get("patterns"),
        keep_quotes=True,
    )
    repl_raw = out.get("replace") if "replace" in out else (out.get("replacements") or out.get("to"))
    repl_list = _norm_str_list(repl_raw, keep_empty_quoted=True, keep_quotes=True)
    # Пустые элементы replace и алиасы (_EMPTY_, …) → _ПУСТО_
    # Кавычки '…' сохраняем (литерал без подстановки <<Переменные…>> при сборе).
    empty_tok = (u"_пусто_", u"_empty_", u"__пусто__")
    fixed_find = []
    for x in find_list or ():
        s = unicode(x) if x is not None else u""
        bare, quoted = unwrap_column_name_token(s)
        if bare.strip().casefold() in empty_tok:
            fixed_find.append(u"_ПУСТО_")
        elif quoted:
            # '' — пропуск; ' '/'  Петр  ' — пробелы сохраняем
            if bare == u"":
                continue
            fixed_find.append(wrap_column_name_token_exact(bare, strip_name=False))
        elif bare.strip() != u"":
            fixed_find.append(bare)
    find_list = fixed_find
    out["find"] = find_list

    fixed_repl = []
    for x in repl_list or ():
        s = unicode(x) if x is not None else u""
        bare, quoted = unwrap_column_name_token(s)
        if bare.strip().casefold() in empty_tok:
            fixed_repl.append(u"_ПУСТО_")
        elif quoted:
            if bare == u"":
                fixed_repl.append(u"_ПУСТО_")
            else:
                fixed_repl.append(wrap_column_name_token_exact(bare, strip_name=False))
        elif bare.strip() == u"":
            fixed_repl.append(u"_ПУСТО_")
        else:
            fixed_repl.append(bare)

    # Режим replace: manual | lambda
    mode_raw = u""
    for _rm_key in (u"replace_mode", u"repl_mode", u"to_mode"):
        if _rm_key in out and unicode(out.get(_rm_key) or u"").strip():
            mode_raw = unicode(out.get(_rm_key) or u"").strip()
            break
    mode_map = {
        u"manual": u"manual",
        u"вручную": u"manual",
        u"ручной": u"manual",
        u"ручной ввод": u"manual",
        u"текст": u"manual",
        u"text": u"manual",
        u"list": u"manual",
        u"lambda": u"lambda",
        u"лямбда": u"lambda",
        u"expr": u"lambda",
        u"выражение": u"lambda",
    }
    replace_mode = mode_map.get(mode_raw.casefold(), u"manual") if mode_raw else u"manual"
    # Авто: есть replace_expr и нет явного replace → lambda
    repl_expr = u""
    for _re_key in (u"replace_expr", u"replace_lambda", u"to_expr", u"to_lambda"):
        if _re_key in out and unicode(out.get(_re_key) or u"").strip():
            repl_expr = unicode(out.get(_re_key) or u"").strip()
            break
    if replace_mode == u"manual" and repl_expr and not fixed_repl and mode_raw == u"":
        replace_mode = u"lambda"
    if replace_mode == u"lambda":
        out[u"replace_mode"] = u"lambda"
        if repl_expr:
            out[u"replace_expr"] = repl_expr
        else:
            out.pop(u"replace_expr", None)
        # список replace не обязателен в lambda-режиме
        if fixed_repl:
            out["replace"] = fixed_repl
        else:
            out.pop("replace", None)
    else:
        out.pop(u"replace_mode", None)
        out.pop(u"replace_expr", None)
        if fixed_repl:
            out["replace"] = fixed_repl
        elif find_list:
            out["replace"] = [u"_ПУСТО_"]
        else:
            out.pop("replace", None)
    for _re_alias in (u"replace_lambda", u"to_expr", u"to_lambda", u"repl_mode", u"to_mode"):
        out.pop(_re_alias, None)

    def _norm_bool(key, *aliases):
        v = None
        if key in block:
            v = block.get(key)
        else:
            for a in aliases:
                if a in block:
                    v = block.get(a)
                    break
        if v is None:
            return
        try:
            from libre_macros_lib import lm_parse_bool_param

            parsed = lm_parse_bool_param(v, default=None)
            if parsed is True:
                out[key] = True
            elif parsed is False:
                out[key] = False
        except Exception:
            out[key] = bool(v)

    _norm_bool("case_insensitive", "без_учета_регистра", "ignore_case")
    _norm_bool("squeeze_spaces", "сжать_пробелы", "trim_spaces")

    row_filter = u""
    for _rf_key in (u"row_filter", u"row_expr", u"row_lambda", u"filter"):
        if _rf_key in out and unicode(out.get(_rf_key) or u"").strip():
            row_filter = unicode(out.get(_rf_key) or u"").strip()
            break
    if row_filter:
        out[u"row_filter"] = row_filter
    else:
        out.pop(u"row_filter", None)
    for _rf_alias in (u"row_expr", u"row_lambda"):
        out.pop(_rf_alias, None)
    if u"filter" in out and row_filter:
        out.pop(u"filter", None)

    match_raw = u""
    for _mp_key in (u"match_pick", u"which", u"occurrence", u"matches"):
        if _mp_key in out and unicode(out.get(_mp_key) or u"").strip():
            match_raw = unicode(out.get(_mp_key) or u"").strip()
            break
    match_map = {
        u"all": u"all",
        u"все": u"all",
        u"*": u"all",
        u"все_совпадения": u"all",
        u"все совпадения": u"all",
        u"first": u"first",
        u"первый": u"first",
        u"первое": u"first",
        u"last": u"last",
        u"последний": u"last",
        u"последнее": u"last",
        u"odd": u"odd",
        u"нечетный": u"odd",
        u"нечётный": u"odd",
        u"нечетные": u"odd",
        u"нечётные": u"odd",
        u"нечётные (1, 3, …)": u"odd",
        u"even": u"even",
        u"четный": u"even",
        u"чётный": u"even",
        u"четные": u"even",
        u"чётные": u"even",
        u"чётные (2, 4, …)": u"even",
        u"lambda": u"lambda",
        u"лямбда": u"lambda",
    }
    match_pick = match_map.get(match_raw.casefold(), u"all") if match_raw else u"all"
    if match_pick != u"all":
        out[u"match_pick"] = match_pick
    else:
        out.pop(u"match_pick", None)
    for _mp_alias in (u"which", u"occurrence", u"matches"):
        out.pop(_mp_alias, None)

    match_expr = u""
    for _me_key in (u"match_expr", u"match_lambda", u"match_filter"):
        if _me_key in out and unicode(out.get(_me_key) or u"").strip():
            match_expr = unicode(out.get(_me_key) or u"").strip()
            break
    if match_expr:
        out[u"match_expr"] = match_expr
    else:
        out.pop(u"match_expr", None)
    for _me_alias in (u"match_lambda", u"match_filter"):
        out.pop(_me_alias, None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_normalize_text_block(block):
    out = _normalize_block("текстовые_операции", block)
    cols = _normalize_column_token_list(out.get("columns"))
    out["columns"] = cols
    if len(cols) == 0:
        out.pop("columns", None)
    try:
        from libre_macros_normalize_lib import coerce_ops_list

        ops = coerce_ops_list(out.get("ops") or out.get("stack") or out.get("pipeline"))
    except Exception:
        ops = out.get("ops") if isinstance(out.get("ops"), list) else []
    out["ops"] = ops
    if len(ops) == 0:
        out.pop("ops", None)
    for key in ("non_text", "on_error", "empty_cells"):
        if key in out and out.get(key) is not None:
            raw = unicode(out.get(key)).strip()
            low = raw.casefold()
            if key == "non_text":
                mapping = {
                    u"пропуск": u"skip",
                    u"как текст": u"coerce",
                    u"как_текст": u"coerce",
                    u"ошибка": u"error",
                    u"skip": u"skip",
                    u"coerce": u"coerce",
                    u"error": u"error",
                }
                out[key] = mapping.get(low, low) or u"skip"
            elif key == "on_error":
                mapping = {
                    u"оставить до ошибки": u"keep",
                    u"оставить": u"keep",
                    u"очистить ячейку": u"empty",
                    u"очистить": u"empty",
                    u"остановить": u"stop",
                    u"keep": u"keep",
                    u"skip_cell": u"keep",
                    u"empty": u"empty",
                    u"stop": u"stop",
                }
                out[key] = mapping.get(low, low) or u"keep"
            else:
                out[key] = raw
    if "as_values" in block:
        try:
            from libre_macros_lib import lm_parse_bool_param

            parsed = lm_parse_bool_param(block.get("as_values"), default=True)
            if parsed is not None:
                out["as_values"] = bool(parsed)
        except Exception:
            out["as_values"] = bool(block.get("as_values"))
    for key in ("key", "key_file"):
        if key in out and out.get(key) is not None:
            out[key] = unicode(out.get(key)).strip()
            if out[key] == u"":
                out.pop(key, None)
    if "key_cell" in out and out.get("key_cell") in (u"", None):
        out.pop("key_cell", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_rename_sheet_block(block):
    out = _normalize_block("переименовать_лист", block)
    out["new_name"] = unicode(out.get("new_name") or "").strip()
    return _attach_sheet(out, out.get("sheet"))

_REORDER_POSITION_ALIASES = {
    "_начало_": "start",
    "_конец_": "end",
    "_перед_": "before",
    "_после_": "after",
    "начало": "start",
    "конец": "end",
    "перед": "before",
    "после": "after",
    "start": "start",
    "end": "end",
    "before": "before",
    "after": "after",
}


def _normalize_reorder_position(raw):
    key = unicode(raw or "").strip().casefold()
    if key in _REORDER_POSITION_ALIASES:
        return _REORDER_POSITION_ALIASES[key]
    if key in ("start", "end", "before", "after"):
        return key
    return ""


def unwrap_column_name_token(token):
    """
    Токен списка столбцов → (имя, quoted).

    Одинарные кавычки вокруг имени — quoted=True (кавычки в JSON/UI не часть имени;
    внутри '' → литерал '). quoted больше не означает «только exact»:
    см. column_header_matches_token ('имя' = exact, 'имя*' = шаблон).

    Также снимаются типографские кавычки ‘…’ «…» и парные ".
    """
    s = unicode(token or u"").strip()
    if len(s) < 2:
        return (s, False)
    # ASCII '…' с экранированием '' внутри.
    if s[0] == u"'" and s[-1] == u"'":
        return (s[1:-1].replace(u"''", u"'"), True)
    pairs = (
        (u"\u2018", u"\u2019"),  # ‘ ’
        (u"\u2019", u"\u2019"),  # ’…’
        (u"\u00AB", u"\u00BB"),  # « »
        (u"\u201C", u"\u201D"),  # “ ”
        (u'"', u'"'),
    )
    for left, right in pairs:
        if s[0] == left and s[-1] == right:
            return (s[1:-1], True)
    return (s, False)


def _clean_column_match_text(text):
    """Та же нормализация пробелов/невидимых, что у заголовков на листе."""
    s = unicode(text or u"").strip()
    if s == u"":
        return u""
    try:
        from libre_macros_header_lib import clean_header_name

        return unicode(clean_header_name(s) or u"").strip()
    except Exception:
        while u"  " in s:
            s = s.replace(u"  ", u" ")
        return s.strip()


def _normalize_column_pattern_stars(pattern):
    """Unicode-звёздочки → ASCII *."""
    p = unicode(pattern or u"")
    for ch in (u"\u2217", u"\uFF0A", u"\u204E", u"\uFE61"):
        if ch in p:
            p = p.replace(ch, u"*")
    return p


def column_header_matches_token(title, token):
    """
    Единая логика отбора столбца по токену (визарды и исполнители).

    - 'name'     — полное совпадение заголовка (без учёта регистра);
    - 'name*' / '*name' / 'na?e' — fnmatch-шаблон внутри кавычек;
    - part       — без кавычек: как '*part*' (подстрока);
    - part* / *x — без кавычек с маской: fnmatch как написано;
    - все / * / all — любой непустой заголовок.

    Кавычки снимаются; пробелы нормализуются как clean_header_name
    (как у заголовков в _lm_pp_header_titles).

    Индексы (1) и буквы (A) без кавычек резолвятся вызывающим кодом до этой функции.
    """
    import fnmatch

    raw = unicode(token or u"").strip()
    text = unicode(title or u"").strip()
    if raw == u"" or text == u"":
        return False
    bare, quoted = unwrap_column_name_token(raw)
    bare = _normalize_column_pattern_stars(_clean_column_match_text(bare))
    text = _clean_column_match_text(text)
    if bare == u"" or text == u"":
        return False
    bare_cf = bare.casefold()
    if bare_cf in (u"все", u"all") or bare == u"*":
        return True
    text_cf = text.casefold()
    has_wild = (u"*" in bare) or (u"?" in bare)
    if quoted:
        if has_wild:
            return fnmatch.fnmatchcase(text_cf, bare_cf)
        return text_cf == bare_cf
    if has_wild:
        return fnmatch.fnmatchcase(text_cf, bare_cf)
    # без кавычек и без маски ≡ '*part*'
    return bare_cf in text_cf


def wrap_column_name_token_exact(name, strip_name=True):
    """Обернуть имя в '…' для точного поиска; ' внутри → ''.

    strip_name=True (по умолчанию) — как для заголовков столбцов.
    strip_name=False — сохранить краевые пробелы (find/replace литералы).
    """
    s = unicode(name or u"")
    if strip_name:
        s = s.strip()
    bare, exact = unwrap_column_name_token(s)
    if exact:
        s = bare
    return u"'" + unicode(s).replace(u"'", u"''") + u"'"


def _is_plain_col_letters_token(text):
    """Латиница A…XYZ (1–3 буквы) — индекс столбца, не имя заголовка."""
    s = unicode(text or u"").strip().upper()
    if s == u"" or len(s) > 3:
        return False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch < u"A" or ch > u"Z":
            return False
        i = i + 1
    return True


def canonicalize_column_name_token_for_store(token):
    """
    Токен столбца для JSON/UI: кавычки сохраняются как ввели; без кавычек — не добавляем.

    - int / цифры (без кавычек) → int (1-based)
    - A…XYZ (1–3 лат. буквы, без кавычек) → голая буква столбца
    - Все / all / * → без кавычек
    - '…' → оставить в '…' (выбор из списка / exact/pattern в кавычках)
    - голое имя / шаблон без кавычек → как есть (подстрока / fnmatch)
    """
    if token is None or isinstance(token, bool):
        return u""
    if isinstance(token, int):
        return int(token)
    if isinstance(token, float):
        try:
            return int(token)
        except (TypeError, ValueError):
            pass
    s = unicode(token).strip()
    if s == u"":
        return u""
    bare, quoted = unwrap_column_name_token(s)
    if bare == u"":
        return u""
    bare_cf = bare.casefold()
    if bare_cf in (u"all", u"все"):
        return u"Все"
    if bare_cf == u"*":
        return u"*"
    if (not quoted) and bare.isdigit():
        try:
            return int(bare)
        except (TypeError, ValueError):
            pass
    # Буквы столбца только без кавычек.
    if (not quoted) and _is_plain_col_letters_token(bare):
        return bare.upper()
    if quoted:
        return wrap_column_name_token_exact(bare)
    return bare


def canonicalize_column_token_list_for_store(value):
    """Список/строка столбцов → список токенов (кавычки только если были)."""
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        seq = value
    else:
        text = unicode(value).strip()
        if text == u"":
            return []
        seq = split_column_name_tokens(text)
    out = []
    for item in seq:
        tok = canonicalize_column_name_token_for_store(item)
        if tok == u"" or tok is None:
            continue
        out.append(tok)
    return out


def _column_name_token_has_sep_outside_quotes(text, sep):
    """True, если sep встречается вне одинарных кавычек."""
    s = unicode(text or u"")
    in_quote = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == u"'":
            if in_quote and i + 1 < len(s) and s[i + 1] == u"'":
                i = i + 2
                continue
            in_quote = not in_quote
            i = i + 1
            continue
        if (not in_quote) and ch == sep:
            return True
        i = i + 1
    return False


def _split_column_name_tokens_by_sep(text, sep):
    """
    Разбор по sep с экранированием \\, и \\;.

    Запятые/«;» внутри (), [] и {} и внутри '…' не считаются разделителями
    (имена вроде 'Фамилия, Имя, Отчество' и шаблоны Summa_[Строка(первая, 2)]).
    Кавычки в токене сохраняются (признак exact-match).
    """
    s = unicode(text or "")
    out = []
    buf = []
    esc = False
    in_quote = False
    depth_paren = 0
    depth_bracket = 0
    depth_brace = 0
    i = 0
    while i < len(s):
        ch = s[i]
        if esc:
            buf.append(ch)
            esc = False
            i = i + 1
            continue
        if (not in_quote) and ch == u"\\":
            esc = True
            i = i + 1
            continue
        if ch == u"'":
            if in_quote and i + 1 < len(s) and s[i + 1] == u"'":
                buf.append(u"'")
                buf.append(u"'")
                i = i + 2
                continue
            in_quote = not in_quote
            buf.append(ch)
            i = i + 1
            continue
        if in_quote:
            buf.append(ch)
            i = i + 1
            continue
        if ch == u"(":
            depth_paren = depth_paren + 1
            buf.append(ch)
            i = i + 1
            continue
        if ch == u")":
            if depth_paren > 0:
                depth_paren = depth_paren - 1
            buf.append(ch)
            i = i + 1
            continue
        if ch == u"[":
            depth_bracket = depth_bracket + 1
            buf.append(ch)
            i = i + 1
            continue
        if ch == u"]":
            if depth_bracket > 0:
                depth_bracket = depth_bracket - 1
            buf.append(ch)
            i = i + 1
            continue
        if ch == u"{":
            depth_brace = depth_brace + 1
            buf.append(ch)
            i = i + 1
            continue
        if ch == u"}":
            if depth_brace > 0:
                depth_brace = depth_brace - 1
            buf.append(ch)
            i = i + 1
            continue
        if (
            ch == sep
            and depth_paren == 0
            and depth_bracket == 0
            and depth_brace == 0
        ):
            t = u"".join(buf).strip()
            if t != "":
                out.append(t)
            buf = []
            i = i + 1
            continue
        buf.append(ch)
        i = i + 1
    t = u"".join(buf).strip()
    if t != "":
        out.append(t)
    return out


def detect_column_name_list_sep(text, known_names=None):
    """
    Разделитель списка имён столбцов.
    «;» — если уже есть вне кавычек, либо одно имя совпадает с known
    и содержит запятую без кавычек (для будущих добавлений).
    """
    s = unicode(text or "").strip()
    if s == "":
        return u","
    if _column_name_token_has_sep_outside_quotes(s, u";"):
        return u";"
    if known_names is not None:
        bare, _exact = unwrap_column_name_token(s)
        bare_cf = bare.casefold()
        for name in known_names or ():
            nm = unicode(name or "").strip()
            if nm != "" and nm.casefold() == bare_cf and u"," in nm:
                # известное имя с запятой без кавычек → дальше через «;»
                name2, exact2 = unwrap_column_name_token(s)
                if not exact2:
                    return u";"
    return u","


def split_column_name_tokens(text, known_names=None):
    """
    Список имён столбцов из строки: «;» если есть вне кавычек, иначе «,».
    known_names: если вся строка = одно известное имя (с запятой) — не дробить.
    Одинарные кавычки защищают запятые и задают exact-match (кавычки сохраняются).
    Поддержка экранирования: \\, и \\; вне кавычек; '' внутри кавычек.
    """
    s = unicode(text or "").strip()
    if s == "":
        return []
    if _column_name_token_has_sep_outside_quotes(s, u";"):
        return _split_column_name_tokens_by_sep(s, u";")
    if known_names is not None:
        bare, exact = unwrap_column_name_token(s)
        bare_cf = bare.casefold()
        for name in known_names or ():
            nm = unicode(name or "").strip()
            if nm != "" and nm.casefold() == bare_cf:
                return [s if exact else nm]
    return _split_column_name_tokens_by_sep(s, u",")


def join_column_name_tokens(tokens, sep=None):
    """
    Склеить имена. Имена с «,»/«;» без кавычек оборачиваются в '…'
    (чтобы не резались при следующем split). sep=None → «,» если все
    «опасные» имена в кавычках, иначе «;».
    """
    parts = []
    for t in tokens or []:
        s = unicode(t or u"").strip()
        if s == u"":
            continue
        bare, exact = unwrap_column_name_token(s)
        if (u"," in bare or u";" in bare) and not exact:
            s = wrap_column_name_token_exact(bare)
        parts.append(s)
    if not parts:
        return u""
    needs_semi = False
    for p in parts:
        bare, exact = unwrap_column_name_token(p)
        if (u"," in bare or u";" in bare) and not exact:
            needs_semi = True
            break
    if sep not in (u",", u";"):
        sep = u";" if needs_semi else u","
    elif sep == u"," and needs_semi:
        sep = u";"
    return sep.join(parts)


def split_sep_list_respecting_quotes(value, keep_empty_quoted=False, keep_quotes=False):
    """
    Список токенов из строки/list с разделителями «,»/«;».

    В одинарных кавычках '…' запятая и «;» не режут токен
    (например: 'word1', word2, ',' → word1, word2, ,).
    По умолчанию кавычки снимаются ('' внутри → ').
    keep_empty_quoted=True — сохранить пустой токен из '' (для replace → _ПУСТО_).
    keep_quotes=True — оставить '…' в результате (литерал без подстановки плейсхолдеров);
    при этом краевые пробелы внутри кавычек сохраняются.
    """
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        raw_parts = []
        for v in value:
            raw_parts.append(unicode(v if v is not None else u""))
    else:
        text = unicode(value or u"").strip()
        if text == u"":
            return []
        raw_parts = split_column_name_tokens(text)
    out = []
    for raw in raw_parts:
        s = unicode(raw or u"").strip()
        if s == u"":
            if keep_empty_quoted:
                # голый пустой элемент списка (не из кавычек) — пропуск
                continue
            continue
        bare, quoted = unwrap_column_name_token(s)
        if quoted:
            if bare == u"" and not keep_empty_quoted:
                continue
            if keep_quotes:
                out.append(wrap_column_name_token_exact(bare, strip_name=False))
            else:
                out.append(bare)
        else:
            if bare == u"":
                continue
            out.append(bare)
    return out


def join_sep_list_respecting_quotes(tokens, sep=u";", empty_as=u"", keep_quotes=False):
    """
    Склеить токены для UI: значения с «,»/«;» — в '…'; пустые → empty_as.
    keep_quotes=True — уже закавыченные токены оставить в '…'.
    Краевые пробелы внутри литерала сохраняются.
    """
    parts = []
    for t in tokens or []:
        if t is None:
            s = u""
        else:
            s = unicode(t)
        if s == u"":
            if empty_as is None:
                continue
            parts.append(unicode(empty_as))
            continue
        bare, quoted = unwrap_column_name_token(s)
        if keep_quotes and quoted:
            parts.append(wrap_column_name_token_exact(bare, strip_name=False))
            continue
        if quoted:
            s = bare
        if u"," in s or u";" in s or s != s.strip():
            parts.append(wrap_column_name_token_exact(s, strip_name=False))
        else:
            parts.append(s)
    if not parts:
        return u""
    use_sep = unicode(sep or u";")
    return (use_sep + u" ").join(parts) if use_sep == u";" else use_sep.join(parts)


def _find_vlookup_extract_alias_sep(text):
    """
    Позиция разделителя переименования extract ВПР вне '…'.
    Сначала «->», иначе «=». Возвращает (pos, len) или (-1, 0).
    """
    s = unicode(text or u"")
    best = (-1, 0)
    in_quote = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == u"'":
            if in_quote and i + 1 < len(s) and s[i + 1] == u"'":
                i = i + 2
                continue
            in_quote = not in_quote
            i = i + 1
            continue
        if not in_quote:
            if i + 1 < len(s) and s[i] == u"-" and s[i + 1] == u">":
                return (i, 2)
            if ch == u"=" and best[0] < 0:
                best = (i, 1)
        i = i + 1
    return best


def parse_vlookup_extract_alias(token):
    """
    Токен «Столбцы (E)» ВПР → (src_token, out_name|None).

    Синтаксис (разделитель «->» или «=»):
      'ФИО'                  — имя на левом как на правом
      ФИО -> Новое_имя       — переименовать на левом листе
      'ФИО' = 'Новое_имя'
      'ФИО -> Новое_имя'     — весь токен в кавычках

    out_name=None — взять заголовок с правого листа.
    src_token сохраняет '…' для exact-поиска столбца справа.
    """
    s = unicode(token or u"").strip()
    if s == u"":
        return (u"", None)
    pos, seplen = _find_vlookup_extract_alias_sep(s)
    if pos < 0:
        bare, exact = unwrap_column_name_token(s)
        if exact:
            pos2, seplen2 = _find_vlookup_extract_alias_sep(bare)
            if pos2 >= 0:
                src = bare[:pos2].strip()
                as_name = bare[pos2 + seplen2 :].strip()
                if as_name != u"":
                    as_bare, _ex = unwrap_column_name_token(as_name)
                    return (src, as_bare if as_bare != u"" else as_name)
        return (s, None)
    src = s[:pos].strip()
    as_name = s[pos + seplen :].strip()
    if as_name == u"":
        return (src, None)
    as_bare, _ex = unwrap_column_name_token(as_name)
    return (src, as_bare if as_bare != u"" else as_name)


def append_column_name_token(text, new_token, sep=None, known_names=None, exact=False):
    """
    Добавить имя столбца к списку без дублей.
    exact=True — обернуть в '…' (выбор из списка заголовков → полное совпадение).
    Если имя содержит «,» — кавычки и/или разделитель «;».
    Возвращает (joined_text, effective_sep).
    """
    token = unicode(new_token or u"").strip()
    raw = unicode(text or u"").strip()
    cur_sep = sep if sep in (u",", u";") else detect_column_name_list_sep(raw, known_names)
    if token == u"":
        return (raw, cur_sep)
    bare, was_exact = unwrap_column_name_token(token)
    if exact or was_exact:
        token = wrap_column_name_token_exact(bare)
        bare, _e = unwrap_column_name_token(token)
    token_cf = bare.casefold()
    if token_cf in (u"all", u"все", u"*"):
        return (u"Все", cur_sep)
    tokens = split_column_name_tokens(raw, known_names)
    if len(tokens) == 1:
        t0, _e0 = unwrap_column_name_token(tokens[0])
        if t0.casefold() in (u"all", u"все", u"*"):
            tokens = []
    for t in tokens:
        tb, _te = unwrap_column_name_token(t)
        if tb.casefold() == token_cf:
            return (join_column_name_tokens(tokens, sep=cur_sep), cur_sep)
    # кавычки защищают запятую — можно оставить «,»; без кавычек → «;»
    if u"," in bare and not (exact or was_exact):
        cur_sep = u";"
    elif any(
        (u"," in unwrap_column_name_token(t)[0] and not unwrap_column_name_token(t)[1])
        for t in tokens
    ):
        cur_sep = u";"
    tokens.append(token)
    return (join_column_name_tokens(tokens, sep=cur_sep), cur_sep)


def expand_rename_columns_mapping_pairs(mappings):
    """
    Развернуть mappings: old/new могут содержать несколько имён через , или ;.
    Пары сопоставляются по позиции (1-е → 1-е, …).
    В old кавычки сохраняются (exact-match); в new — снимаются.
    """
    out = []
    for item in mappings or []:
        if not isinstance(item, dict):
            continue
        olds = split_column_name_tokens(
            item.get("old") or item.get("from") or item.get("source") or ""
        )
        news = split_column_name_tokens(
            item.get("new") or item.get("to") or item.get("dest") or ""
        )
        if not olds or not news:
            continue
        n = min(len(olds), len(news))
        i = 0
        while i < n:
            o = olds[i]
            nv, _nex = unwrap_column_name_token(news[i])
            if o != "" and nv != "":
                out.append({"old": o, "new": nv})
            i = i + 1
    return out


def _parse_column_rename_mappings(value):
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        out = []
        for item in value:
            if isinstance(item, dict):
                old = unicode(
                    item.get("old") or item.get("from") or item.get("source") or ""
                ).strip()
                new = unicode(
                    item.get("new") or item.get("to") or item.get("dest") or ""
                ).strip()
                if old != "" and new != "":
                    out.append({"old": old, "new": new})
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                old = unicode(item[0]).strip()
                new = unicode(item[1]).strip()
                if old != "" and new != "":
                    out.append({"old": old, "new": new})
        return out
    text = unicode(value).strip()
    if text == "":
        return []
    out = []
    pair_sep = u";" if u";" in text and u"=" in text else u","
    for part in re.split(re.escape(pair_sep), text):
        part = part.strip()
        if part == "":
            continue
        if "=" in part:
            old, _, new = part.partition("=")
        elif "->" in part:
            old, _, new = part.partition("->")
        else:
            continue
        old = old.strip()
        new = new.strip()
        if old != "" and new != "":
            out.append({"old": old, "new": new})
    return out


def _normalize_rename_columns_block(block):
    out = _normalize_block("переименовать_столбцы", block)
    mappings = block.get("mappings")
    if mappings is None:
        mappings = block.get("map") or block.get("renames")
    if mappings is None:
        old_name = unicode(block.get("old") or block.get("from") or "").strip()
        new_name = unicode(block.get("new") or block.get("to") or "").strip()
        if old_name != "" and new_name != "":
            mappings = [{"old": old_name, "new": new_name}]
    parsed = _parse_column_rename_mappings(mappings)
    if parsed:
        out["mappings"] = parsed
    else:
        out.pop("mappings", None)
    out.pop("old", None)
    out.pop("new", None)
    out.pop("from", None)
    out.pop("to", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_reorder_columns_block(block):
    out = _normalize_block("переставить_столбцы", block)
    out["columns"] = _normalize_column_token_list(out.get("columns"))
    pos = _normalize_reorder_position(out.get("position"))
    if pos != "":
        out["position"] = pos
    else:
        out.pop("position", None)
    rel = unicode(
        out.get("relative_column") or out.get("relative") or out.get("anchor") or ""
    ).strip()
    if rel != "":
        out["relative_column"] = rel
    else:
        out.pop("relative_column", None)
    out.pop("relative", None)
    out.pop("anchor", None)
    create_new = False
    if "create_new" in out:
        create_new = bool(out.get("create_new"))
    else:
        for key in ("новые", "new_empty", "blank_columns", "create_empty"):
            if key in out:
                create_new = bool(out.get(key))
                break
    for key in ("новые", "new_empty", "blank_columns", "create_empty"):
        out.pop(key, None)
    do_copy = False
    if "copy" in out:
        do_copy = bool(out.get("copy"))
    else:
        mode = unicode(out.get("mode") or "").strip().casefold()
        if mode in (
            "copy",
            "копировать",
            "_копировать_",
            "дублировать",
            "duplicate",
        ):
            do_copy = True
    out.pop("mode", None)
    # create_new важнее copy: пустые столбцы по новым именам.
    if create_new:
        out["create_new"] = True
        out.pop("copy", None)
        do_copy = False
    else:
        out.pop("create_new", None)
        if do_copy:
            out["copy"] = True
        else:
            out.pop("copy", None)
    raw_names = out.get("new_names")
    if raw_names is None:
        raw_names = out.get("copy_names")
    if raw_names is None:
        raw_names = out.get("new_columns")
    names = []
    if isinstance(raw_names, (list, tuple)):
        for item in raw_names:
            names.append(unicode("" if item is None else item).strip())
    elif raw_names is not None and unicode(raw_names).strip() != "":
        for part in re.split(r"[,;]", unicode(raw_names)):
            names.append(part.strip())
    out.pop("copy_names", None)
    out.pop("new_columns", None)
    if (do_copy or create_new) and names:
        out["new_names"] = names
    else:
        out.pop("new_names", None)
    sheet = unicode(out.get("sheet") or "").strip()
    if sheet != "":
        out["sheet"] = sheet
    elif "sheet" in out:
        del out["sheet"]
    return out
def _normalize_value_count_block(block):
    out = _normalize_block("количество_значений", block)
    col = unicode(out.get("new_column") or out.get("column") or u"").strip()
    if col != u"":
        out["new_column"] = col
    else:
        out.pop("new_column", None)
    out.pop("column", None)
    cols = out.get("key_columns")
    if cols is None:
        cols = out.get("columns")
    out["key_columns"] = _normalize_column_token_list(cols)
    out.pop("columns", None)
    if "trim" in out:
        out["trim"] = bool(out.get("trim"))
    else:
        out["trim"] = True
    if "case_sensitive" in out:
        out["case_sensitive"] = bool(out.get("case_sensitive"))
    else:
        out["case_sensitive"] = False
    return _attach_sheet(out, out.get("sheet"))
def _normalize_remove_duplicates_block(block):
    out = _normalize_block("удалить_дубликаты", block)
    out["key_columns"] = _normalize_column_token_list(out.get("key_columns"))
    out.pop("columns", None)
    out["exclude_columns"] = _normalize_column_token_list(out.get("exclude_columns"))
    keep = unicode(out.get("keep") or u"first").strip().casefold()
    if keep not in (u"first", u"last"):
        keep = u"first"
    out["keep"] = keep
    scan = unicode(out.get("scan_order") or u"").strip().casefold()
    if scan in (u"top_down", u"bottom_up"):
        out["scan_order"] = scan
    else:
        out.pop("scan_order", None)
    cmp_as = unicode(out.get("compare_as") or u"text").strip().casefold()
    out["compare_as"] = u"strict" if cmp_as == u"strict" else u"text"
    # Удобная галочка визарда: пропускать строки, у которых суммарный ключ = пусто.
    # Поле сохраняем в JSON явно, а empty_key_policy держим для совместимости исполнения.
    skip_empty_cells = out.get("skip_empty_cells")
    if skip_empty_cells is not None:
        out["skip_empty_cells"] = bool(skip_empty_cells)
        ekp = u"skip" if bool(skip_empty_cells) else u"dedup"
    else:
        ekp = unicode(out.get("empty_key_policy") or u"dedup").strip().casefold()
        out["skip_empty_cells"] = (ekp == u"skip")
    if ekp not in (u"dedup", u"unique", u"skip"):
        ekp = u"dedup"
    out["empty_key_policy"] = ekp
    output = unicode(out.get("output") or u"inplace").strip().casefold()
    if output not in (u"inplace", u"new_sheet", u"offset"):
        output = u"inplace"
    out["output"] = output
    dest_sheet = unicode(
        out.get("dest_sheet") or out.get("dest") or out.get("target_sheet") or u""
    ).strip()
    if dest_sheet != u"":
        out["dest_sheet"] = dest_sheet
    else:
        out.pop("dest_sheet", None)
    out.pop("dest", None)
    out.pop("target_sheet", None)
    dest_cell = unicode(out.get("dest_cell") or u"").strip()
    if dest_cell != u"":
        out["dest_cell"] = dest_cell
    elif "dest_cell" in out:
        del out["dest_cell"]
    if "trim" in out:
        out["trim"] = bool(out.get("trim"))
    else:
        out["trim"] = True
    if "case_sensitive" in out:
        out["case_sensitive"] = bool(out.get("case_sensitive"))
    else:
        out["case_sensitive"] = False
    if "as_values" in out:
        out["as_values"] = bool(out.get("as_values"))
    else:
        out["as_values"] = True
    if "mark_duplicates" in out:
        out["mark_duplicates"] = bool(out.get("mark_duplicates"))
    if "overwrite_overlap" in out:
        out["overwrite_overlap"] = bool(out.get("overwrite_overlap"))
    mark_col = unicode(out.get("mark_column") or u"").strip()
    if mark_col != u"":
        out["mark_column"] = mark_col
    else:
        out.pop("mark_column", None)
    pre_sort = out.get("pre_sort")
    norm_ps = []
    if isinstance(pre_sort, (list, tuple)):
        for item in pre_sort:
            if not isinstance(item, dict):
                continue
            col = unicode(item.get("column") or u"").strip()
            if col == u"":
                continue
            order = unicode(item.get("order") or u"asc").strip().casefold()
            norm_ps.append(
                {
                    "column": col,
                    "order": u"desc"
                    if order in (u"desc", u"-", u"убыв")
                    else u"asc",
                }
            )
    out["pre_sort"] = norm_ps
    return _attach_sheet(out, out.get("sheet"))
def _normalize_activate_sheet_block(block):
    out = _normalize_block("активировать_лист", block)
    target = unicode(
        out.get("target_sheet")
        or out.get("activate_sheet")
        or out.get("sheet_name")
        or out.get("sheet")
        or u""
    ).strip()
    out["target_sheet"] = target
    out.pop("activate_sheet", None)
    out.pop("sheet_name", None)
    tab = unicode(
        out.get("tab_color")
        or out.get("color")
        or out.get("ярлык")
        or out.get("tabColor")
        or u""
    ).strip()
    if tab != u"":
        out["tab_color"] = tab
    else:
        out.pop("tab_color", None)
    out.pop("color", None)
    out.pop("ярлык", None)
    out.pop("tabColor", None)
    return out


def _normalize_copy_sheet_block(block):
    out = _normalize_block("копировать_переместить_лист", block)
    out["source_sheet"] = unicode(
        out.get("source_sheet") or out.get("source") or u""
    ).strip()
    out["dest_sheet"] = unicode(
        out.get("dest_sheet") or out.get("dest") or out.get("target_sheet") or u""
    ).strip()
    cols = out.get("columns")
    if cols is None:
        cols = out.get("markers")  # совместимость: раньше часто думали “удалить маркеры”
    out["columns"] = _normalize_column_token_list(cols)
    out.pop("markers", None)
    if "skip_empty" in out:
        out["skip_empty"] = bool(out.get("skip_empty"))
    sec = out.get("skip_empty_columns")
    if sec is None:
        sec = out.get("skip_empty_cols")
    out["skip_empty_columns"] = _normalize_column_token_list(sec)
    out.pop("skip_empty_cols", None)
    if "header_row" in out and out.get("header_row") not in (None, u""):
        try:
            out["header_row"] = int(out.get("header_row"))
        except (TypeError, ValueError):
            out.pop("header_row", None)
    out.pop("source", None)
    out.pop("dest", None)
    out.pop("target_sheet", None)
    # По умолчанию: copy=false = перемещение.
    # Но для совместимости с legacy-`копировать_лист` (когда в JSON не было ни `copy`,
    # ни `position/anchor_sheet`) считаем, что dest_sheet задан → это именно копия.
    do_copy = False
    legacy_position_missing = (
        unicode(out.get("position") or out.get("place") or out.get("where") or u"").strip()
        == u""
    )
    legacy_anchor_missing = (
        unicode(out.get("anchor_sheet") or out.get("relative_sheet") or out.get("ref_sheet") or out.get("anchor") or u"").strip()
        == u""
    )
    legacy_dest_present = unicode(out.get("dest_sheet") or u"").strip() != u""
    legacy_copy = ("copy" not in out) and legacy_position_missing and legacy_anchor_missing and legacy_dest_present
    if "copy" in out:
        do_copy = bool(out.get("copy"))
    elif "mode" in out:
        mode = unicode(out.get("mode") or u"").strip().casefold()
        do_copy = mode in (u"copy", u"копировать", u"дублировать")
    elif legacy_copy:
        do_copy = True
    out.pop("mode", None)
    if do_copy:
        out["copy"] = True
    else:
        out.pop("copy", None)
    # Позиционирование относительно опорного листа.
    pos = unicode(
        out.get("position") or out.get("place") or out.get("where") or u""
    ).strip().casefold()
    if pos in (u"до", u"перед", u"before", u"_перед_"):
        out["position"] = u"before"
    elif pos in (u"после", u"after", u"_после_"):
        out["position"] = u"after"
    else:
        out.pop("position", None)
    out.pop("place", None)
    out.pop("where", None)
    anchor = unicode(
        out.get("anchor_sheet")
        or out.get("relative_sheet")
        or out.get("ref_sheet")
        or out.get("anchor")
        or u""
    ).strip()
    if anchor != u"":
        out["anchor_sheet"] = anchor
    else:
        out.pop("anchor_sheet", None)
    out.pop("relative_sheet", None)
    out.pop("ref_sheet", None)
    out.pop("anchor", None)
    return _attach_sheet(out, out.get("sheet"))


def _normalize_copy_ranges_block(block):
    out = _normalize_block("копирование_диапазонов", block)
    # Алиасы полей.
    if not unicode(out.get("source_sheet") or u"").strip():
        out["source_sheet"] = unicode(
            out.get("from_sheet") or u""
        ).strip()
    out["source_sheet"] = unicode(out.get("source_sheet") or u"").strip()
    out.pop("from_sheet", None)

    if not unicode(out.get("source_range") or u"").strip():
        out["source_range"] = unicode(out.get("source") or u"").strip()
    out["source_range"] = unicode(out.get("source_range") or u"").strip()
    out.pop("source", None)

    if "source_rows" in out and out.get("source_rows") not in (None, u""):
        rows = out.get("source_rows")
        if isinstance(rows, (list, tuple)):
            out["source_rows"] = [
                unicode(x).strip() for x in rows if unicode(x).strip() != u""
            ]
        else:
            out["source_rows"] = unicode(rows).strip()
    else:
        out.pop("source_rows", None)

    if "source_columns" in out and out.get("source_columns") not in (None, u"", []):
        out["source_columns"] = _normalize_column_token_list(out.get("source_columns"))
    else:
        out.pop("source_columns", None)

    dest_sheets = out.get("dest_sheets")
    if dest_sheets is None:
        dest_sheets = out.get("to_sheets") or out.get("target_sheets")
    names = []
    if isinstance(dest_sheets, (list, tuple)):
        names = [
            unicode(x or u"").strip()
            for x in dest_sheets
            if unicode(x or u"").strip() != u""
        ]
    elif dest_sheets is not None and unicode(dest_sheets).strip() != u"":
        names = [
            p.strip()
            for p in unicode(dest_sheets).replace(u";", u",").split(u",")
            if p.strip() != u""
        ]
    out.pop("to_sheets", None)
    out.pop("target_sheets", None)

    one = unicode(
        out.get("dest_sheet")
        or out.get("to_sheet")
        or out.get("target_sheet")
        or u""
    ).strip()
    out.pop("to_sheet", None)
    out.pop("target_sheet", None)
    if names:
        out["dest_sheets"] = names
        out.pop("dest_sheet", None)
    elif one != u"":
        out["dest_sheet"] = one
        out.pop("dest_sheets", None)
    else:
        out.pop("dest_sheets", None)
        out.pop("dest_sheet", None)

    dest_cell = unicode(
        out.get("dest_cell") or out.get("anchor") or out.get("at") or u"A1"
    ).strip()
    out["dest_cell"] = dest_cell if dest_cell else u"A1"
    out.pop("at", None)
    # anchor уже мог быть dest_cell; не путать с anchor_sheet copy_sheet.
    if "anchor" in out and "dest_cell" in out:
        out.pop("anchor", None)

    mode = unicode(
        out.get("mode") or out.get("paste_mode") or u"replace"
    ).strip().casefold()
    out.pop("paste_mode", None)
    if mode in (u"insert", u"append", u"сдвиг", u"добавление"):
        out["mode"] = u"insert"
    else:
        out["mode"] = u"replace"

    axis = unicode(out.get("insert_axis") or u"auto").strip().casefold()
    if axis in (u"rows", u"row", u"строки", u"строка"):
        out["insert_axis"] = u"rows"
    elif axis in (u"cols", u"columns", u"col", u"столбцы", u"столбец"):
        out["insert_axis"] = u"cols"
    else:
        out["insert_axis"] = u"auto"

    # content / bool-алиасы.
    content = unicode(out.get("content") or u"").strip().casefold()
    as_values = False
    keep_formulas = False
    try:
        from libre_macros_lib import lm_parse_bool_param

        if "as_values" in out:
            as_values = bool(lm_parse_bool_param(out.get("as_values"), default=False))
        if u"только_значения" in out:
            as_values = as_values or bool(
                lm_parse_bool_param(out.get(u"только_значения"), default=False)
            )
        if "keep_formulas" in out:
            keep_formulas = bool(
                lm_parse_bool_param(out.get("keep_formulas"), default=False)
            )
        if u"формулы" in out and not isinstance(out.get(u"формулы"), (unicode, str)):
            keep_formulas = keep_formulas or bool(
                lm_parse_bool_param(out.get(u"формулы"), default=False)
            )
        elif unicode(out.get(u"формулы") or u"").strip().casefold() in (
            u"1",
            u"true",
            u"да",
            u"yes",
        ):
            keep_formulas = True
    except Exception:
        as_values = bool(out.get("as_values"))
        keep_formulas = bool(out.get("keep_formulas"))
    out.pop("as_values", None)
    out.pop(u"только_значения", None)
    out.pop("keep_formulas", None)
    out.pop(u"формулы", None)
    if as_values or content in (u"values", u"value", u"значения"):
        out["content"] = u"values"
    elif keep_formulas or content in (u"formulas", u"formula", u"формулы"):
        out["content"] = u"formulas"
    else:
        out["content"] = u"values"

    for key, default in (
        ("with_formatting", False),
        ("clear_source", False),
        ("create_missing_dest", False),
        ("involve_dest", True),
    ):
        if key not in out:
            out[key] = default
        else:
            try:
                from libre_macros_lib import lm_parse_bool_param

                out[key] = bool(lm_parse_bool_param(out.get(key), default=default))
            except Exception:
                out[key] = bool(out.get(key))

    if "header_row" in out and out.get("header_row") not in (None, u""):
        try:
            out["header_row"] = int(out.get("header_row"))
        except (TypeError, ValueError):
            out.pop("header_row", None)
    else:
        out.pop("header_row", None)

    return _attach_sheet(out, out.get("sheet"))


def _normalize_create_sheet_block(block):
    out = _normalize_block("создать_лист", block)
    mode = unicode(
        out.get("name_mode") or out.get("mode") or u"manual"
    ).strip().casefold()
    if mode in (
        u"expr",
        u"lambda",
        u"compute",
        u"вычислить",
        u"лямбда",
        u"expression",
    ):
        out["name_mode"] = u"expr"
    else:
        out["name_mode"] = u"manual"
    out.pop("mode", None)

    sheet_name = unicode(
        out.get("sheet_name")
        or out.get("name")
        or out.get("dest_sheet")
        or u""
    ).strip()
    out.pop("name", None)
    out.pop("dest_sheet", None)
    expr = unicode(
        out.get("sheet_name_expr")
        or out.get("name_expr")
        or u""
    ).strip()
    out.pop("name_expr", None)
    # UI dual-field name_value → в нужное поле по mode
    name_value = unicode(out.get("name_value") or u"").strip()
    out.pop("name_value", None)
    if out["name_mode"] == u"expr":
        if expr == u"" and name_value != u"":
            expr = name_value
        out["sheet_name_expr"] = expr
        if sheet_name != u"":
            out["sheet_name"] = sheet_name
        else:
            out.pop("sheet_name", None)
    else:
        if sheet_name == u"" and name_value != u"":
            sheet_name = name_value
        out["sheet_name"] = sheet_name
        if expr != u"":
            out["sheet_name_expr"] = expr
        else:
            out.pop("sheet_name_expr", None)

    cols = out.get("columns")
    if cols is None:
        cols = out.get("headers") or out.get("header_columns")
    out.pop("headers", None)
    out.pop("header_columns", None)
    if cols is None or cols == u"" or cols == []:
        out["columns"] = []
    else:
        # Голые имена для записи в ячейки; кавычки только как защита разделителей.
        out["columns"] = split_sep_list_respecting_quotes(cols, keep_quotes=False)

    if "header_row" in out and out.get("header_row") not in (None, u""):
        try:
            hr = int(out.get("header_row"))
            out["header_row"] = hr if hr > 0 else 1
        except (TypeError, ValueError):
            out["header_row"] = 1
    else:
        out["header_row"] = 1

    out["source_sheet"] = unicode(
        out.get("source_sheet") or out.get("from_sheet") or u""
    ).strip()
    out.pop("from_sheet", None)
    out["source_range"] = unicode(
        out.get("source_range") or out.get("source") or u""
    ).strip()
    out.pop("source", None)
    if out["source_sheet"] == u"":
        out.pop("source_sheet", None)
    if out["source_range"] == u"":
        out.pop("source_range", None)

    rows = unicode(
        out.get("rows") or out.get("data_rows") or out.get("manual_rows") or u""
    ).strip()
    out.pop("data_rows", None)
    out.pop("manual_rows", None)
    if rows != u"":
        out["rows"] = rows
    else:
        out.pop("rows", None)

    dm = unicode(out.get("data_mode") or u"").strip().casefold()
    if dm in (u"sheet", u"лист", u"с листа", u"range", u"диапазон"):
        out["data_mode"] = u"sheet"
    elif dm in (
        u"rows",
        u"manual",
        u"вручную",
        u"текст",
        u"text",
        u"values",
        u"строки",
    ):
        out["data_mode"] = u"rows"
    elif dm in (u"none", u"нет", u"empty", u"пусто", u"-"):
        out["data_mode"] = u"none"
    else:
        # Авто по заполненным полям
        if out.get("rows"):
            out["data_mode"] = u"rows"
        elif out.get("source_sheet") or out.get("source_range"):
            out["data_mode"] = u"sheet"
        else:
            out["data_mode"] = u"none"

    if "values_only" not in out:
        out["values_only"] = True
    else:
        try:
            from libre_macros_lib import lm_parse_bool_param

            out["values_only"] = bool(
                lm_parse_bool_param(out.get("values_only"), default=True)
            )
        except Exception:
            out["values_only"] = bool(out.get("values_only"))

    return _attach_sheet(out, out.get("sheet"))


def _normalize_send_mail_block(block):
    out = _normalize_block("отправить_по_почте", block)
    try:
        import libre_macros_mail_lib as ml
    except ImportError:
        ml = None

    def _addr_list(key):
        raw = out.get(key)
        if raw is None:
            out[key] = []
            return
        if ml is not None:
            out[key] = ml.normalize_address_list(raw)
            return
        if isinstance(raw, (list, tuple)):
            out[key] = [unicode(x or u"").strip() for x in raw if unicode(x or u"").strip()]
        else:
            out[key] = [
                p.strip()
                for p in unicode(raw).replace(u";", u",").split(u",")
                if p.strip() != u""
            ]

    for key in ("to", "cc", "bcc"):
        _addr_list(key)

    sheets = out.get("sheets")
    if sheets is None:
        out["sheets"] = []
    elif isinstance(sheets, (list, tuple)):
        out["sheets"] = [
            unicode(x or u"").strip()
            for x in sheets
            if unicode(x or u"").strip() != u""
        ]
    else:
        out["sheets"] = [
            p.strip()
            for p in unicode(sheets).replace(u";", u",").split(u",")
            if p.strip() != u""
        ]

    if ml is not None:
        out["attach"] = ml.normalize_attach_mode(out.get("attach"))
        out["format"] = ml.normalize_format(out.get("format"))
        out["client"] = ml.normalize_client(out.get("client"))
        out["os"] = ml.normalize_os_param(out.get("os"))
    else:
        out["attach"] = unicode(out.get("attach") or u"whole_book").strip().casefold()
        out["format"] = unicode(out.get("format") or u"ods").strip().casefold()
        out["client"] = unicode(out.get("client") or u"auto").strip().casefold()
        out["os"] = unicode(out.get("os") or u"auto").strip().casefold()

    out["filename"] = unicode(out.get("filename") or u"").strip()
    out["to_source_sheet"] = unicode(out.get("to_source_sheet") or u"").strip()
    out["to_source_range"] = unicode(out.get("to_source_range") or u"").strip()
    out["subject"] = unicode(out.get("subject") or u"").strip()
    out["body"] = unicode(out.get("body") or u"").strip()

    if out["to_source_sheet"] == u"":
        out.pop("to_source_sheet", None)
    if out["to_source_range"] == u"":
        out.pop("to_source_range", None)
    if out["filename"] == u"":
        out.pop("filename", None)
    if out["subject"] == u"":
        out.pop("subject", None)
    if out["body"] == u"":
        out.pop("body", None)
    if not out.get("sheets"):
        out.pop("sheets", None)

    if "show_ui" not in out:
        out["show_ui"] = True
    else:
        try:
            from libre_macros_lib import lm_parse_bool_param

            out["show_ui"] = bool(lm_parse_bool_param(out.get("show_ui"), default=True))
        except Exception:
            out["show_ui"] = bool(out.get("show_ui"))

    out.pop("sheet", None)
    return out


def _normalize_merge_sheets_into_one_block(block):
    out = _normalize_block("объединить_листы_в_один", block)
    sheets = out.get("sheets")
    if sheets is None:
        sheets = out.get("source_sheets") or out.get("sheet_list")
    if isinstance(sheets, (list, tuple)):
        out["sheets"] = [
            unicode(x or u"").strip()
            for x in sheets
            if unicode(x or u"").strip() != u""
        ]
    elif sheets is not None:
        out["sheets"] = [
            p.strip()
            for p in unicode(sheets).replace(u";", u",").split(u",")
            if p.strip() != u""
        ]
    else:
        out["sheets"] = []
    out.pop("source_sheets", None)
    out.pop("sheet_list", None)
    out["dest_sheet"] = unicode(
        out.get("dest_sheet")
        or out.get("result_sheet")
        or out.get("target_sheet")
        or out.get("dest")
        or u""
    ).strip()
    out.pop("result_sheet", None)
    out.pop("target_sheet", None)
    out.pop("dest", None)
    if "header_row" in out and out.get("header_row") not in (None, u""):
        try:
            out["header_row"] = int(out.get("header_row"))
        except (TypeError, ValueError):
            out.pop("header_row", None)
    # with_formatting: копировать визуальное оформление ячеек с источников
    wf = None
    for key in (
        "with_formatting",
        "с_форматированием",
        "keep_format",
        "copy_format",
        "keep_formatting",
    ):
        if key in out:
            wf = out.get(key)
            break
    if wf is None:
        out["with_formatting"] = False
    else:
        try:
            from libre_macros_lib import lm_parse_bool_param

            out["with_formatting"] = bool(lm_parse_bool_param(wf, default=False))
        except Exception:
            out["with_formatting"] = bool(wf)
    for key in (
        "с_форматированием",
        "keep_format",
        "copy_format",
        "keep_formatting",
    ):
        out.pop(key, None)
    # columns: пусто = все столбцы по заголовкам источников
    cols = _normalize_column_token_list(out.get("columns"))
    if cols:
        out["columns"] = cols
    else:
        out.pop("columns", None)
    # auto_format: оформление заголовка/шрифта данных по умолчанию
    af = None
    for key in ("auto_format", "автоформат", "autofmt", "default_format"):
        if key in out:
            af = out.get(key)
            break
    if af is None:
        out["auto_format"] = True
    else:
        try:
            from libre_macros_lib import lm_parse_bool_param

            out["auto_format"] = bool(lm_parse_bool_param(af, default=True))
        except Exception:
            out["auto_format"] = bool(af)
    for key in ("автоформат", "autofmt", "default_format"):
        out.pop(key, None)
    return _attach_sheet(out, out.get("sheet"))


_SPLIT_SHEETS_RUN_PHASES = (
    "after_collect",
    "after_postprocess",
    "after_final",
)


def _normalize_split_sheets_block(block):
    out = _normalize_block("разделить_листы", block)
    phase = unicode(out.get("run_phase") or u"after_postprocess").strip().casefold()
    if phase not in _SPLIT_SHEETS_RUN_PHASES:
        phase = u"after_postprocess"
    out["run_phase"] = phase
    src = out.get("source")
    if not isinstance(src, dict):
        src = {}
    out["source"] = {
        "book": u"эта_книга",
        "sheet": unicode(
            src.get("sheet") or out.get("source_sheet") or out.get("sheet") or u""
        ).strip(),
    }
    out.pop("source_sheet", None)

    def _norm_fields(key):
        raw = out.get(key)
        if raw is None:
            out[key] = []
            return
        if isinstance(raw, (list, tuple)):
            out[key] = [
                unicode(x or u"").strip()
                for x in raw
                if unicode(x or u"").strip() != u""
            ]
        else:
            out[key] = [
                p.strip()
                for p in unicode(raw).replace(u";", u",").split(u",")
                if p.strip() != u""
            ]

    _norm_fields("split_books")
    _norm_fields("split_sheets")
    out["book_template"] = unicode(
        out.get("book_template") or out.get("book") or u"эта_книга"
    ).strip()
    out["sheet_template"] = unicode(
        out.get("sheet_template") or out.get("sheet") or u"[Split]"
    ).strip()
    mode = unicode(out.get("split_mode") or u"copy_range").strip().casefold()
    out["split_mode"] = mode if mode in (u"copy_range", u"copy_sheet") else u"copy_range"
    if "header_row" in out and out.get("header_row") not in (None, u""):
        try:
            out["header_row"] = int(out.get("header_row"))
        except (TypeError, ValueError):
            out.pop("header_row", None)
    for key, default in (
        ("pre_sort", True),
        ("copy_header", True),
        ("create_folders", True),
        ("overwrite_sheets", True),
        ("append_if_exists", False),
        ("save_immediately", False),
        ("paste_formats", False),
    ):
        if key not in out:
            out[key] = default
        else:
            try:
                from libre_macros_lib import lm_parse_bool_param

                out[key] = bool(lm_parse_bool_param(out.get(key), default=default))
            except Exception:
                out[key] = bool(out.get(key))
    return out


def _column_tokens_to_row_extra(cols):
    """Список столбцов из блока → extra_args для lm_pp_row_convert_numeric_strings."""
    out = []
    for c in cols or []:
        if isinstance(c, int):
            out.append(c)
            continue
        if isinstance(c, float):
            out.append(c)
            continue
        t = unicode(c).strip()
        if t == "":
            continue
        if t.isdigit():
            out.append(int(t))
            continue
        try:
            if "." in t:
                out.append(float(t.replace(",", ".")))
                continue
        except (TypeError, ValueError):
            pass
        out.append(t)
    return out

def param_decode(fn_key, raw_text):
    """Всегда list[dict] — нормализованные блоки. Только JSON."""
    fn_key = normalize_fn_key(fn_key)
    raw = unicode(raw_text or "").strip()
    if raw == "":
        return []
    if len(raw) >= 2 and raw[0] == u""" and raw[-1] == u""":
        unq_s = unicode(raw[1:-1]).strip()
        if _looks_like_json_payload(unq_s):
            raw = unq_s
    if not _looks_like_json_payload(raw):
        return []
    try:
        items = _parse_json_root(raw)
    except (ValueError, TypeError):
        return []
    return _decode_json_blocks(fn_key, items)


def param_encode(fn_key, blocks):
    """Сериализация блоков в компактный JSON."""
    fn_key = normalize_fn_key(fn_key)
    blocks = blocks or []
    payload = []
    for block in blocks:
        b = _normalize_block_for_fn(fn_key, block)
        payload.append(b)
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def row_extra_args_from_codec(fn_key, raw_text):
    """
    Блоки кодека → extra_args для ROW-исполнителей.

    Возвращает list или None, если fn_key не ROW-кодек.
    """
    fn_key = normalize_fn_key(fn_key)
    if fn_key not in ("преобразовать_числа", "цвет_текста_по_значению"):
        return None
    blocks = param_decode(fn_key, raw_text)
    if fn_key == "цвет_текста_по_значению":
        if len(blocks) == 0:
            return []
        if len(blocks) == 1:
            m = unicode(blocks[0].get("marker") or u"статус").strip()
            return [m or u"статус"]
        return [param_encode(fn_key, blocks)]
    if len(blocks) == 0:
        return []
    if len(blocks) == 1:
        cols = blocks[0].get("columns") or []
        if not cols:
            return []
        return _column_tokens_to_row_extra(cols)
    return [param_encode(fn_key, blocks)]


def sheets_map_is_filter_payload(text):
    """True, если текст — JSON/lambda-фильтр колонки C «Листы», не карта source→target."""
    s = unicode(text or u"").strip()
    if s == u"":
        return False
    lead = s.lstrip()
    if lead.startswith(u"lambda") or lead.startswith(u"def "):
        return True
    if not (s.startswith(u"{") or s.startswith(u"[")):
        return False
    try:
        obj = json.loads(s)
    except Exception:
        return True
    if isinstance(obj, dict):
        if unicode(obj.get("filter") or u"").strip() != u"":
            return True
        if "items" in obj or "map" in obj or "source" in obj or "target" in obj:
            return False
        # неизвестный объект с filter-подобным содержимым
        return "filter" in obj
    if isinstance(obj, list):
        if not obj:
            return False
        first = obj[0]
        if isinstance(first, dict):
            if unicode(first.get("filter") or u"").strip() != u"":
                return True
            if any(k in first for k in ("source", "target", "items", "map")):
                return False
        return False
    return False


def sheets_map_parse(text):
    """
    Разобрать параметр «Листы» (legacy CSV или JSON-карту) →
    (items=[{source, target}], ignore_hidden).
    """
    s = unicode(text or u"").strip()
    items = []
    ignore_hidden = False
    IGNORE = set([u"!hidden", u"ignore_hidden", u"игнорировать_скрытые"])
    ARROW = u"->"

    def _add(src, tgt):
        # JSON может отдать source:1 (число) — храним как "1" для INDEX.
        if isinstance(src, bool):
            src = unicode(src)
        elif isinstance(src, int):
            src = unicode(src)
        elif isinstance(src, float) and src == int(src):
            src = unicode(int(src))
        else:
            src = unicode(src or u"").strip()
        tgt = unicode(tgt or u"").strip()
        if src == u"":
            return
        if tgt == u"":
            tgt = src
        items.append({"source": src, "target": tgt})

    if s == u"":
        return (items, ignore_hidden)

    if (s.startswith(u"{") or s.startswith(u"[")) and (not sheets_map_is_filter_payload(s)):
        try:
            obj = json.loads(s)
        except Exception:
            obj = None
        if isinstance(obj, dict):
            if obj.get("ignore_hidden") is True:
                ignore_hidden = True
            seq = obj.get("items") or obj.get("map") or None
            if seq is None and (obj.get("source") or obj.get("target")):
                seq = [obj]
            for it in seq or []:
                if not isinstance(it, dict):
                    continue
                if it.get("ignore_hidden") is True:
                    ignore_hidden = True
                    continue
                _add(it.get("source") or it.get("from") or it.get("old"),
                     it.get("target") or it.get("to") or it.get("new"))
            return (items, ignore_hidden)
        if isinstance(obj, list):
            for it in obj:
                if not isinstance(it, dict):
                    continue
                if it.get("ignore_hidden") is True:
                    ignore_hidden = True
                    continue
                if "items" in it or "map" in it:
                    if it.get("ignore_hidden") is True:
                        ignore_hidden = True
                    for sub in (it.get("items") or it.get("map") or []):
                        if isinstance(sub, dict):
                            _add(sub.get("source") or sub.get("from") or sub.get("old"),
                                 sub.get("target") or sub.get("to") or sub.get("new"))
                    continue
                _add(it.get("source") or it.get("from") or it.get("old"),
                     it.get("target") or it.get("to") or it.get("new"))
            return (items, ignore_hidden)

    # legacy: a,b->c,!hidden
    normalized = s.replace(u";", u",")
    for chunk in normalized.split(u","):
        t = unicode(chunk or u"").strip()
        if t == u"":
            continue
        if unicode(t).casefold() in IGNORE:
            ignore_hidden = True
            continue
        if ARROW in t:
            left, right = t.split(ARROW, 1)
            _add(left, right)
        else:
            _add(t, t)
    return (items, ignore_hidden)


def sheets_map_to_legacy_tokens(items, ignore_hidden=False):
    """Карта → токены Старое / Старое->Новое / !hidden для iter_sheet_slots."""
    out = []
    for it in items or []:
        src = unicode((it or {}).get("source") or u"").strip()
        tgt = unicode((it or {}).get("target") or u"").strip()
        if src == u"":
            continue
        if tgt != u"" and tgt != src:
            out.append(src + u"->" + tgt)
        else:
            out.append(src)
    if ignore_hidden:
        out.append(u"!hidden")
    return out


def sheets_map_encode_json(items, ignore_hidden=False):
    """Сериализация карты листов в JSON для ячейки параметра «Листы»."""
    payload = {"v": 1, "fn": u"листы", "items": []}
    for it in items or []:
        src = unicode((it or {}).get("source") or u"").strip()
        tgt = unicode((it or {}).get("target") or u"").strip()
        if src == u"":
            continue
        entry = {"source": src}
        if tgt != u"" and tgt != src:
            entry["target"] = tgt
        payload["items"].append(entry)
    if ignore_hidden:
        payload["ignore_hidden"] = True
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def sheets_map_expand_cell_to_specs(text, default_specs=None):
    """
    Ячейка «Листы» → list токенов для classify/iter_sheet_slots.
    JSON-карта разворачивается в legacy-токены; фильтр-JSON не трогаем (→ default).
    """
    if default_specs is None:
        default_specs = []
    s = unicode(text or u"").strip()
    if s == u"":
        return list(default_specs)
    if sheets_map_is_filter_payload(s):
        return list(default_specs)
    items, ignore_hidden = sheets_map_parse(s)
    if (s.startswith(u"{") or s.startswith(u"[")) and items:
        toks = sheets_map_to_legacy_tokens(items, ignore_hidden=ignore_hidden)
        return toks if toks else list(default_specs)
    # legacy CSV уже разобран sheets_map_parse
    if items or ignore_hidden:
        toks = sheets_map_to_legacy_tokens(items, ignore_hidden=ignore_hidden)
        return toks if toks else list(default_specs)
    return list(default_specs)


def columns_map_is_structured_payload(text):
    """True, если ячейка «Столбцы» — JSON-карта pick + rename."""
    s = unicode(text or u"").strip()
    if s == u"":
        return False
    if not (s.startswith(u"{") or s.startswith(u"[")):
        return False
    try:
        obj = json.loads(s)
    except Exception:
        return False
    if isinstance(obj, dict):
        fn = unicode(obj.get("fn") or u"").strip().casefold()
        if fn in (u"листы", u"sheets", u"sheet_map"):
            return False
        if fn in (u"столбцы", u"columns", u"column_map"):
            return True
        if "filter" in obj or obj.get("ignore_hidden") is True:
            return False
        if "pick" in obj:
            return True
        if "items" in obj or "map" in obj:
            return True
        if obj.get("source") or obj.get("target") or obj.get("from") or obj.get("to"):
            return True
        return False
    if isinstance(obj, list) and obj:
        first = obj[0]
        if isinstance(first, dict):
            fn = unicode(first.get("fn") or u"").strip().casefold()
            if fn in (u"столбцы", u"columns", u"column_map"):
                return True
            if any(k in first for k in ("source", "target", "from", "to", "old", "new", "pick", "items")):
                return True
    return False


_COLUMNS_MAP_CONVERT_TRUE = frozenset(
    (u"true", u"yes", u"1", u"да", u"д", u"on", u"y")
)
_COLUMNS_MAP_CONVERT_FALSE = frozenset(
    (u"false", u"no", u"0", u"нет", u"н", u"off", u"n")
)
_COLUMNS_MAP_CONVERT_DEFAULT = frozenset(
    (u"", u"auto", u"default", u"по умолчанию", u"def")
)


def columns_map_normalize_convert_flag(raw):
    """convert_to_numbers в блоке столбца: True / False / None (по умолчанию)."""
    if raw is None:
        return None
    if isinstance(raw, bool):
        return raw
    s = unicode(raw or u"").strip().casefold()
    if s in _COLUMNS_MAP_CONVERT_FALSE:
        return False
    if s in _COLUMNS_MAP_CONVERT_TRUE:
        return True
    if s in _COLUMNS_MAP_CONVERT_DEFAULT:
        return None
    return None


def columns_map_normalize_item(raw):
    """Нормализовать один блок items: source, target, convert_to_numbers?, format?."""
    if not isinstance(raw, dict):
        return None
    src = unicode(
        raw.get("source") or raw.get("from") or raw.get("old") or u""
    ).strip()
    tgt = unicode(
        raw.get("target") or raw.get("to") or raw.get("new") or u""
    ).strip()
    if src == u"":
        return None
    if tgt == u"":
        tgt = src
    out = {"source": src, "target": tgt}
    ctn = columns_map_normalize_convert_flag(raw.get("convert_to_numbers"))
    if ctn is not None:
        out["convert_to_numbers"] = ctn
    fmt = unicode(raw.get("format") or raw.get("number_format") or u"").strip()
    if fmt != u"":
        out["format"] = fmt
    return out


def columns_map_match_item_to_header(header, item, output_headers=False):
    """True, если блок items относится к заголовку (source или target после rename)."""
    hdr = unicode(header or u"").strip()
    if hdr == u"" or not isinstance(item, dict):
        return False
    src = unicode(item.get("source") or u"").strip()
    tgt = unicode(item.get("target") or src or u"").strip()
    if src != u"" and column_header_matches_token(hdr, src):
        return True
    if tgt != u"" and column_header_matches_token(hdr, tgt):
        return True
    _unused = output_headers
    return False


def columns_map_item_for_header(header, items, output_headers=False):
    for it in items or []:
        if columns_map_match_item_to_header(header, it, output_headers=output_headers):
            return it
    return None


def columns_map_headers_for_src_range(sc0, sc1, cols):
    """Индекс столбца в getDataArray-чанке (0-based) → имя заголовка (после rename)."""
    by_src = {}
    for sc, hname in cols or []:
        try:
            by_src[int(sc)] = unicode(hname or u"").strip()
        except (TypeError, ValueError):
            continue
    out = {}
    col = int(sc0)
    end = int(sc1)
    while col <= end:
        if col in by_src:
            out[col - int(sc0)] = by_src[col]
        col += 1
    return out


def columns_map_merge_headers(headers_by_col, extra_by_col):
    """Объединить карты заголовков (extra перекрывает)."""
    merged = {}
    if headers_by_col:
        for col, hdr in headers_by_col.items():
            merged[int(col)] = unicode(hdr or u"").strip()
    if extra_by_col:
        for col, hdr in extra_by_col.items():
            h = unicode(hdr or u"").strip()
            if h != u"":
                merged[int(col)] = h
    return merged


def columns_map_headers_by_index_from_matrix(matrix, hdr_row0):
    out = {}
    if not matrix:
        return out
    hr = int(hdr_row0 if hdr_row0 is not None else 0)
    if hr < 0 or hr >= len(matrix):
        return out
    row = matrix[hr] or []
    c = 0
    while c < len(row):
        out[c] = unicode(row[c] or u"").strip()
        c += 1
    return out


def columns_map_format_storage_kind(fmt):
    """Классификация явного формата: date | number | text | None (авто)."""
    s = unicode(fmt or u"").strip()
    if s == u"":
        return None
    if s == u"@" or s.casefold() in (u"text", u"текст"):
        return u"text"
    up = s.upper()
    if any(
        tok in up
        for tok in (
            u"DD", u"MM", u"YY", u"HH", u"NNN", u"DDD", u"MMMM", u"MMM",
        )
    ):
        return u"date"
    if any(tok in s for tok in (u"ДД", u"ММ", u"ГГ", u"ЧЧ", u"МММ")):
        return u"date"
    if u"#" in s or u"%" in s:
        return u"number"
    if re.match(r"^0([,\\.]\d+)?%?$", s):
        return u"number"
    return None


def columns_map_has_explicit_convert(items):
    """True, если хотя бы один блок «Столбцы» с convert_to_numbers=true."""
    for it in items or ():
        if columns_map_normalize_convert_flag((it or {}).get("convert_to_numbers")) is True:
            return True
    return False


def columns_map_apply_convert_policies( date_cols, num_cols, headers_by_col, items, output_headers=False ):
    """
    Учесть convert_to_numbers (false — исключить; true — принудительно включить),
    явный format.
    Возвращает (date_cols, num_cols, explicit_fmt_by_col).
    """
    date_cols = set(date_cols or ())
    num_cols = set(num_cols or ())
    explicit_fmt = {}
    if not items or not headers_by_col:
        return date_cols, num_cols, explicit_fmt
    for col, hdr in headers_by_col.items():
        it = columns_map_item_for_header(hdr, items, output_headers=output_headers)
        if not it:
            continue
        ctn = columns_map_normalize_convert_flag(it.get("convert_to_numbers"))
        fmt = unicode(it.get("format") or u"").strip()
        if fmt != u"":
            explicit_fmt[col] = fmt
        if ctn is False:
            date_cols.discard(col)
            num_cols.discard(col)
            continue
        if ctn is True:
            kind = columns_map_format_storage_kind(fmt)
            if kind == u"text":
                date_cols.discard(col)
                num_cols.discard(col)
            elif kind == u"date":
                date_cols.add(col)
                num_cols.discard(col)
            elif kind == u"number":
                num_cols.add(col)
                date_cols.discard(col)
            else:
                date_cols.add(col)
                num_cols.add(col)
    return date_cols, num_cols, explicit_fmt


def columns_map_flatten_items(per_file_items):
    """Все блоки столбцов из per_file_col_rename_items (без дублей source)."""
    out = []
    seen = set()
    for group in per_file_items or []:
        for it in group or []:
            norm = columns_map_normalize_item(it)
            if norm is None:
                continue
            key = unicode(norm.get("source") or u"").strip().casefold()
            if key == u"" or key in seen:
                continue
            seen.add(key)
            out.append(norm)
    return out


def columns_map_parse(text):
    """
    Разобрать параметр «Столбцы» (legacy CSV или JSON) →
    (pick_tokens: list[str], items: list[{source, target}]).
    """
    s = unicode(text or u"").strip()
    pick_tokens = []
    items = []
    ARROW = u"->"

    def _add_item(src=None, tgt=None, raw=None):
        if isinstance(raw, dict):
            norm = columns_map_normalize_item(raw)
            if norm:
                items.append(norm)
            return
        src = unicode(src or u"").strip()
        tgt = unicode(tgt or u"").strip()
        if src == u"":
            return
        if tgt == u"":
            tgt = src
        items.append({"source": src, "target": tgt})

    def _add_pick(tok):
        t = unicode(tok or u"").strip()
        if t != u"":
            pick_tokens.append(t)

    if s == u"":
        return (pick_tokens, items)

    if (s.startswith(u"{") or s.startswith(u"[")) and columns_map_is_structured_payload(s):
        try:
            obj = json.loads(s)
        except Exception:
            obj = None
        if isinstance(obj, dict):
            raw_pick = obj.get("pick") or obj.get("filter") or obj.get("include") or u""
            if isinstance(raw_pick, (list, tuple)):
                for x in raw_pick:
                    _add_pick(x)
            else:
                rp = unicode(raw_pick or u"").strip()
                if rp != u"":
                    for chunk in rp.replace(u";", u",").split(u","):
                        _add_pick(chunk)
            seq = obj.get("items") or obj.get("map") or None
            if seq is None and (obj.get("source") or obj.get("target") or obj.get("from") or obj.get("to")):
                seq = [obj]
            for it in seq or []:
                if not isinstance(it, dict):
                    continue
                _add_item(raw=it)
            return (pick_tokens, items)
        if isinstance(obj, list):
            for it in obj:
                if not isinstance(it, dict):
                    continue
                if "items" in it or "map" in it:
                    for sub in (it.get("items") or it.get("map") or []):
                        if isinstance(sub, dict):
                            _add_item(raw=sub)
                    continue
                if it.get("pick") is not None or it.get("filter") is not None:
                    raw_pick = it.get("pick") or it.get("filter") or u""
                    if isinstance(raw_pick, (list, tuple)):
                        for x in raw_pick:
                            _add_pick(x)
                    else:
                        rp = unicode(raw_pick or u"").strip()
                        if rp != u"":
                            for chunk in rp.replace(u";", u",").split(u","):
                                _add_pick(chunk)
                    continue
                _add_item(raw=it)
            return (pick_tokens, items)

    normalized = s.replace(u";", u",")
    for chunk in normalized.split(u","):
        t = unicode(chunk or u"").strip()
        if t == u"":
            continue
        if ARROW in t:
            left, right = t.split(ARROW, 1)
            _add_item(left, right)
        else:
            _add_pick(t)
    return (pick_tokens, items)


def columns_map_encode_json(pick, items):
    """Сериализация карты столбцов в JSON."""
    payload = {"v": 1, "fn": u"столбцы", "items": []}
    pick_s = unicode(pick or u"").strip()
    if pick_s != u"" and pick_s != u"*":
        payload["pick"] = pick_s
    for it in items or []:
        norm = columns_map_normalize_item(it)
        if norm is None:
            continue
        entry = {"source": norm["source"]}
        tgt = unicode(norm.get("target") or u"").strip()
        if tgt != u"" and tgt != norm["source"]:
            entry["target"] = tgt
        ctn = columns_map_normalize_convert_flag(norm.get("convert_to_numbers"))
        if ctn is not None:
            entry["convert_to_numbers"] = ctn
        fmt = unicode(norm.get("format") or u"").strip()
        if fmt != u"":
            entry["format"] = fmt
        payload["items"].append(entry)
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def columns_map_to_legacy_tokens(pick, items):
    """Карта → legacy CSV."""
    out = []
    pick_s = unicode(pick or u"").strip()
    if pick_s != u"" and pick_s != u"*":
        for chunk in pick_s.replace(u";", u",").split(u","):
            t = unicode(chunk or u"").strip()
            if t != u"":
                out.append(t)
    for it in items or []:
        src = unicode((it or {}).get("source") or u"").strip()
        tgt = unicode((it or {}).get("target") or u"").strip()
        if src == u"":
            continue
        if tgt != u"" and tgt != src:
            out.append(src + u"->" + tgt)
        else:
            out.append(src)
    return out


def columns_map_expand_pick_specs(text, default_specs=None):
    """Ячейка «Столбцы» → токены отбора для build_column_list."""
    if default_specs is None:
        default_specs = []
    s = unicode(text or u"").strip()
    if s == u"":
        return list(default_specs)
    if (s.startswith(u"{") or s.startswith(u"[")) and columns_map_is_structured_payload(s):
        pick_tokens, _items = columns_map_parse(s)
        if pick_tokens:
            return list(pick_tokens)
        return list(default_specs)
    pick_tokens, rename_items = columns_map_parse(s)
    if rename_items and not pick_tokens:
        return list(default_specs)
    if pick_tokens:
        return list(pick_tokens)
    return list(default_specs)


def columns_map_rename_items_from_cell(text):
    """Правила переименования из ячейки «Столбцы»."""
    s = unicode(text or u"").strip()
    if s == u"":
        return []
    _pick, items = columns_map_parse(s)
    return list(items or [])


def _columns_map_wildcard_tail(header, source_pat):
    """Хвост для source с * (Заказы_* → «июнь» из «заказы_июнь»)."""
    import fnmatch

    raw = unicode(source_pat or u"").strip()
    hdr = _clean_column_match_text(unicode(header or u""))
    if raw == u"" or hdr == u"":
        return None
    bare, _quoted = unwrap_column_name_token(raw)
    bare = _normalize_column_pattern_stars(_clean_column_match_text(bare))
    if u"*" not in bare:
        return None
    if not column_header_matches_token(hdr, raw):
        return None
    star = bare.index(u"*")
    prefix = bare[:star]
    suffix_pat = bare[star + 1 :]
    if prefix:
        if not hdr.casefold().startswith(prefix.casefold()):
            return None
        tail = hdr[len(prefix) :]
    else:
        tail = hdr
    if suffix_pat:
        if not fnmatch.fnmatchcase(tail.casefold(), suffix_pat.casefold()):
            return None
    return tail


def columns_map_rename_header(header, source_pat, target_pat):
    """Переименовать заголовок: * сохраняет хвост, иначе полная подстановка."""
    hdr = unicode(header or u"").strip()
    src = unicode(source_pat or u"").strip()
    tgt = unicode(target_pat or u"").strip()
    if hdr == u"" or src == u"" or tgt == u"":
        return hdr
    if not column_header_matches_token(hdr, src):
        return hdr
    src_bare, _q = unwrap_column_name_token(src)
    src_bare = _normalize_column_pattern_stars(_clean_column_match_text(src_bare))
    tgt_bare, _q2 = unwrap_column_name_token(tgt)
    tgt_bare = _normalize_column_pattern_stars(_clean_column_match_text(tgt_bare))
    if u"*" in src_bare and u"*" in tgt_bare:
        tail = _columns_map_wildcard_tail(hdr, src)
        if tail is None:
            return hdr
        star = tgt_bare.index(u"*")
        return tgt_bare[:star] + tail + tgt_bare[star + 1 :]
    return tgt_bare


def columns_map_apply_renames(cols, items):
    """Применить rename к [(col_idx, header), …]; конфликт → _01, _02…"""
    if not cols:
        return cols
    rules = []
    for it in items or []:
        if not isinstance(it, dict):
            continue
        src = unicode(it.get("source") or u"").strip()
        tgt = unicode(it.get("target") or u"").strip()
        if src == u"" or tgt == u"":
            continue
        rules.append((src, tgt))
    if not rules:
        return cols
    taken = {}
    out = []
    ci = 0
    while ci < len(cols):
        col_idx, hname = cols[ci]
        ci = ci + 1
        new_name = unicode(hname or u"")
        ri = 0
        while ri < len(rules):
            src, tgt = rules[ri]
            ri = ri + 1
            candidate = columns_map_rename_header(new_name, src, tgt)
            if candidate != new_name:
                new_name = candidate
                break
        key = new_name.casefold()
        if key in taken:
            n = taken[key] + 1
            taken[key] = n
            stem = new_name
            suffix = u"_%02d" % n
            if len(stem) + len(suffix) > 31:
                stem = stem[: max(1, 31 - len(suffix))]
            new_name = stem + suffix
        else:
            taken[key] = 0
        out.append((col_idx, new_name))
    return out

