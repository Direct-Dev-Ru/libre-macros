# -*- coding: utf-8 -*-
"""Примеры JSON для колонки C (ручные шаблоны и docs/13_JSON_PARAMS.md)."""
from __future__ import print_function, unicode_literals

import json

PARAM_CODEC_VERSION = 1


def jblocks(*blocks):
    """Сериализация блоков как в param_encode."""
    out = []
    for b in blocks:
        block = dict(b)
        block.setdefault("v", PARAM_CODEC_VERSION)
        out.append(block)
    return json.dumps(out, ensure_ascii=False, separators=(",", ":"))


def jone(fn, **fields):
    return jblocks({"fn": fn, **fields})


# --- Список листов (sheets) — пусто или нет ключа = все листы ---
SHEET_LIST_FNS = (
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
    "формат_деньги",
    "серые_служебные",
    "полосы_по_пути",
    "чередующиеся_границы",
    "копировать_формат_заголовка",
    "подсветка_по_заголовку",
    "жирный_по_пути",
)

# Готовые строки для генератора шаблонов (ключ → JSON в колонке C)
PRESET = {
    "объединить_листы_в_один": jone(
        "объединить_листы_в_один",
        sheets=["Отчет_Q1", "Отчет_Q2", "Отчет_Q3"],
        dest_sheet="Отчет_Q123",
        header_row=1,
        with_formatting=False,
    ),
    "тонкая_сетка": jone("тонкая_сетка"),
    "толстая_сетка": jone("толстая_сетка"),
    "авто_ширина": jone("авто_ширина"),
    "авто_высота": jone("авто_высота"),
    "перенос": jone("перенос", wrap=True),
    "перенос_и_авто_высота": jone("перенос_и_авто_высота", wrap=True, include_header=True),
    "левое_выравнивание": jone("левое_выравнивание"),
    "закрепить_заголовок": jone("закрепить_заголовок"),
    "автофильтр": jone("автофильтр"),
    "вертикаль_центр": jone("вертикаль_центр"),
    "формат_деньги": jone("формат_деньги", sheets=["Отчет_январь"]),
    "заголовок_плюс_высота": jone(
        "заголовок_плюс_высота",
        height_mm=13,
        h_align="center",
        v_align="center",
        bold=True,
    ),
    "отступ": jone("отступ", h_align="left", steps=1, columns=["Сумма_руб"]),
    "шрифт": jone("шрифт", name="PT Sans", size_data=10, size_header=11),
    "ширина_столбцов": jone("ширина_столбцов", width_mm=30, columns="all"),
    "высота_строки": jone("высота_строки", height_mm=5, rows="all"),
    "зебра_диапазон": jone(
        "зебра_диапазон",
        roles={
            "header": {"fill": "excel_header", "font": "белый"},
            "even": {"fill": "голубой", "font": "синий"},
            "odd": {"fill": "белый", "font": "синий"},
        },
    ),
    "сетка": jone("сетка", width="тонкая", color="серый", include_header=True),
    "формат_даты": jone(
        "формат_даты", markers=["дата", "время"], convert_existing=False
    ),
    "формат_столбцы": jone(
        "формат_столбцы",
        rules=[{"column": "Сумма_руб", "format": "# ##0,00"}],
        convert_existing=False,
    ),
    "подсветка_по_порогу": jone(
        "подсветка_по_порогу",
        marker="Сумма_руб",
        threshold=70000,
        color="светло_красный",
    ),
    "градиент": jone(
        "градиент",
        marker="Сумма_руб",
        color_min="зеленый",
        color_max="красный",
        whole_row=True,
        sort=True,
        sort_dir="по_возр",
    ),
    "удалить_столбцы": jone("удалить_столбцы", markers=["маркер", "ДатаВремя_строка"]),
    "сортировка": jone(
        "сортировка",
        keys=[
            {"column": "ФИО", "desc": False},
            {"column": "Сумма_руб", "desc": True},
        ],
    ),
    "раскрасить_блоки": jone(
        "раскрасить_блоки",
        key_columns=["ФИО"],
        colors=["голубой", "желтый", "зеленый"],
        sort="asc",
        outline_blocks=True,
        outline_color="темно_синий",
    ),
    "конкатенация_столбцов": jone(
        "конкатенация_столбцов",
        new_column="ФИО_Отдел",
        separator=" / ",
        columns=[1, 3],
    ),
    "применить_формулу": jone(
        "применить_формулу",
        column="Бонус_итог",
        formula="=D{row_id}*0.1",
        as_values=False,
    ),
    "удаление_строк": jone("удаление_строк", mode="columns", columns=["маркер"]),
    "группировка_по_столбцу": jone("группировка_по_столбцу", marker="Отдел"),
    "стиль_печати": jone("стиль_печати", orientation="landscape", fit_pages=1),
    "заполнение_вниз": jone("заполнение_вниз", columns=["Отдел"]),
    "заполнение_вниз_вычислить": jone(
        "заполнение_вниз_вычислить",
        rules=[{"column": "E", "formula": "=D{row_id}*2"}],
    ),
    "копировать_значения": jone("копировать_значения", columns=["Сумма_руб", "Сумма_бонус"]),
    "замена_значений": jone(
        "замена_значений",
        columns=["ФИО"],
        find=["ООО"],
        replace=["ОБЩЕСТВО"],
        case_insensitive=True,
        squeeze_spaces=True,
    ),
    "переименовать_лист": jone("переименовать_лист", new_name="Сводка_тест"),
    "переименовать_столбцы": jone(
        "переименовать_столбцы",
        sheet="Заказы",
        mappings=[{"old": "Сумма_руб", "new": "Сумма"}],
    ),
    "переставить_столбцы": jone(
        "переставить_столбцы",
        sheet="Заказы",
        columns=["C", "Сумма_руб"],
        position="_перед_",
        relative_column="Итого",
    ),
    "количество_значений": jone(
        "количество_значений",
        new_column="Количество",
        key_columns=["ФИО", "Отдел"],
        trim=True,
        case_sensitive=False,
    ),
    "удалить_дубликаты": jone(
        "удалить_дубликаты",
        key_columns=["ФИО"],
        keep="first",
        trim=True,
        case_sensitive=False,
    ),
    "копировать_лист": jone(
        "копировать_лист",
        source_sheet="Шаблон",
        dest_sheet="Копия_Шаблон",
    ),
    "преобразовать_числа": jone("преобразовать_числа", columns=["Сумма_руб", "Сумма_бонус"]),
    "цвет_текста_по_значению": jone("цвет_текста_по_значению", marker="статус"),
    "серые_служебные": jone("серые_служебные"),
    "полосы_по_пути": jone("полосы_по_пути"),
    "чередующиеся_границы": jone("чередующиеся_границы"),
    "копировать_формат_заголовка": jone("копировать_формат_заголовка"),
    "подсветка_по_заголовку": jone("подсветка_по_заголовку"),
    "жирный_по_пути": jone("жирный_по_пути"),
    # анкета ↔ таблица (сценарии 14 / 15; источники sources/forms/source_forms.xlsx)
    "анкета_в_таблицу": jblocks(
        {
            "fn": "анкета_в_таблицу",
            "sheet": "Анкета_ФИО",
            "question_column": "A",
            "answer_column": "B",
            "block_mode": "by_repeat_key",
            "block_start_question": "ФИО",
            "output": "new_sheet",
            "dest_sheet": "Wide_из_ФИО",
            "add_block_index": True,
            "as_values": True,
        },
        {
            "fn": "анкета_в_таблицу",
            "sheet": "Анкета_пусто",
            "question_column": "A",
            "answer_column": "B",
            "block_mode": "by_blank_row",
            "blank_rows_to_split": 1,
            "output": "new_sheet",
            "dest_sheet": "Wide_из_пусто",
            "add_block_index": True,
            "as_values": True,
        },
        {
            "fn": "анкета_в_таблицу",
            "sheet": "Анкета_шапка",
            "question_column": "A",
            "answer_column": "B",
            "block_mode": "by_repeat_key",
            "block_start_question": "ФИО",
            "sheet_preamble_row_from": 2,
            "sheet_preamble_row_to": 4,
            "forms_start_row": 6,
            "block_preamble_rows": 2,
            "block_preamble_mode": "qa_pair",
            "output": "new_sheet",
            "dest_sheet": "Wide_из_шапка",
            "add_block_index": True,
            "as_values": True,
        },
    ),
    "таблица_в_анкету": jone(
        "таблица_в_анкету",
        sheet="Wide_анкеты",
        body_columns=["'ФИО'", "'Возраст'", "'Город'", "'Должность'"],
        block_preamble_columns=["'Бланк №'", "'Дата'"],
        skip_columns=["'#Блок'"],
        question_column="A",
        answer_column="B",
        block_separator="blank_row",
        blank_rows_between=1,
        output="new_sheet",
        dest_sheet="Анкета_из_wide",
        as_values=True,
    ),
}

# RANGE-функции для матрицы 12_postprocess_json
RANGE_MATRIX_ORDER = [
    "тонкая_сетка",
    "заголовок_плюс_высота",
    "авто_ширина",
    "перенос",
    "отступ",
    "шрифт",
    "зебра_диапазон",
    "формат_даты",
    "формат_столбцы",
    "подсветка_по_порогу",
    "градиент",
    "сортировка",
    "раскрасить_блоки",
    "ширина_столбцов",
    "стиль_печати",
    "закрепить_заголовок",
    "автофильтр",
    "замена_значений",
    "количество_значений",
    "копировать_лист",
]

ROW_MATRIX_ORDER = [
    ("высота_строки", "высота_строки"),
    ("преобразовать_числа", "преобразовать_числа"),
    ("цвет_текста_по_значению", "цвет_текста_по_значению"),
    ("серые_служебные", "серые_служебные"),
]

# Алиасы имён на листе параметров → ключ PRESET / fn в JSON
SHEET_FN_ALIASES = {
    "зебра": "зебра_диапазон",
    "условное_форматирование": "подсветка_по_порогу",
    "пропуск_пустых_строк": "удаление_строк",
    "раскрасить": "раскрасить_блоки",
}
