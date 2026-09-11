# -*- coding: utf-8 -*-
"""Константы и состояние file_list_macro (AlterOffice 2026: модульный код .py-скрипта вырезается)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.724"
DEFAULT_MAX_DEPTH = 1
DEFAULT_SIZE_UNIT = "b"
DEFAULT_COUNT_PAGES = False
DEFAULT_SORT_MODE = "tree"
DEPTH_INDENT = "    "
SETTINGS_APP_NAME = "libre-macros"
SETTINGS_MODULE_NAME = "file_list_macro"
SIZE_UNIT_OPTIONS = (
    ("b", "б"),
    ("kb", "Кб"),
    ("mb", "Мб"),
    ("gb", "Гб"),
)
SIZE_UNIT_DIVISOR = {
    "b": 1,
    "kb": 1024,
    "mb": 1024 ** 2,
    "gb": 1024 ** 3,
}
# Пресеты сортировки результата: (code, label).
SORT_PRESETS = (
    ("tree", "Дерево (относительный путь)"),
    ("depth_kind_name", "Глубина → тип → имя"),
    ("full_depth_name", "Полный путь → глубина → имя"),
    ("rel_depth_name", "Относительный путь → глубина → имя"),
    ("kind_name", "Тип → имя"),
    ("name", "Имя"),
    ("ext_name", "Расширение → имя"),
    ("parent_name", "Родитель → имя"),
    ("mtime_desc", "Дата изменения (новые сверху)"),
    ("size_desc", "Размер (больше сверху)"),
)
# PDF + Writer + Impress/Draw (Calc не считаем — «страницы печати» ≠ листы).
PAGE_COUNT_EXTS = frozenset(
    (
        ".pdf",
        ".odt",
        ".ott",
        ".fodt",
        ".doc",
        ".docx",
        ".dot",
        ".dotx",
        ".rtf",
        ".odp",
        ".otp",
        ".fodp",
        ".ppt",
        ".pptx",
        ".pot",
        ".potx",
        ".odg",
        ".otg",
        ".fodg",
    )
)
RESULT_HEADER_PAGES = "Страниц"
RESULT_HEADER_BASE = [
    "Тип",
    "Имя",
    "Относительный путь",
    "Родитель",
    "Глубина",
    None,
    "Дата изменения",
    "Расширение",
    "Полный путь",
]
# Оформление заголовка результата: excel_header / белый (как зебра в проекте).
RESULT_HEADER_STYLE = "excel_header/белый"
# После OptimalWidth — запас под иконку автофильтра (доля от текущей ширины).
COLUMN_WIDTH_AUTOFILTER_PAD = 0.15
UI_YIELD_EVERY = 10
WRITE_CHUNK = 300

# Мутабельное состояние UI (живёт в pythonpath — не проходит через AO-фильтр скрипта)
status_indicator = None
yield_counter = 0
