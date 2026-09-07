# -*- coding: utf-8 -*-
"""Константы и состояние file_list_macro (AlterOffice 2026: модульный код .py-скрипта вырезается)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.689"
DEFAULT_MAX_DEPTH = 1
DEFAULT_SIZE_UNIT = "b"
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
UI_YIELD_EVERY = 10
WRITE_CHUNK = 300

# Мутабельное состояние UI (живёт в pythonpath — не проходит через AO-фильтр скрипта)
status_indicator = None
yield_counter = 0
