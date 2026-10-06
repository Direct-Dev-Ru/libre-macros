# -*- coding: utf-8 -*-
"""Константы и состояние file_list_macro (AlterOffice 2026: модульный код .py-скрипта вырезается)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.730"
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
# Windows: владелец NTFS (часто совпадает с последним редактором в доменных шарах).
RESULT_HEADER_OWNER = "Owner"
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

# --- DLP (мини): анализ имён/путей по словарю regex ---
DEFAULT_DLP_ANALYZE_NAMES = False
DEFAULT_DLP_ANALYZE_CONTENT = False
DLP_NAME_DICT_FILENAME = "dlp_name_patterns.json"
DLP_CONTENT_DICT_FILENAME = "dlp_content_patterns.json"
DLP_LOG_SHEET_SUFFIX = "_DLP"
DLP_LOG_PATTERN_HEADER = "Паттерн"
# Подсветка строки результата при совпадении (RGB Calc).
DLP_MATCH_ROW_COLOR = 0xFFCCCC
# Словарь по умолчанию: id → regex (флаги можно встроить, напр. (?i)).
DEFAULT_DLP_NAME_PATTERNS = {
    "confidential": r"(?i)конфиденц",
    "trade_secret": r"(?i)коммерческ\w*[\s_\-]*тайн",
    "dsp": r"(?i)(?:^|[_\-\s./\\])дсп(?:[_\-\s./\\]|$)",
    "secret": r"(?i)секретн",
    "personal_data": r"(?i)персональн\w*[\s_\-]*данн",
    "restricted": r"(?i)для[\s_\-]*служебн",
    "internal_only": r"(?i)только[\s_\-]*для[\s_\-]*внутренн",
    "password_hint": r"(?i)(?:парол|password|passwd|pwd)[_\-\s.]*(?:list|файл|file|баз)?",
}

# Мутабельное состояние UI (живёт в pythonpath — не проходит через AO-фильтр скрипта)
status_indicator = None
yield_counter = 0
