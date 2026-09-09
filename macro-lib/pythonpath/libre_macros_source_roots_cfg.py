# -*- coding: utf-8 -*-
"""
Доверенные корни «Файлы-Источники» — ручные константы.

Меняйте здесь defaults без правки логики в libre_macros_source_roots_lib.py.
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.715"
# Режим по умолчанию: strict | off
#   strict — путь/glob только под разрешёнными корнями
#   off    — проверка выключена (как раньше)
SOURCE_ROOTS_DEFAULT_MODE = u"strict"

# Env / глобальная переменная профиля (доп. корни и режим).
SOURCE_ROOTS_ENV_ROOTS = u"MERGE_ALLOWED_SOURCE_ROOTS"
SOURCE_ROOTS_ENV_MODE = u"MERGE_SOURCE_ROOTS_MODE"
SOURCE_ROOTS_GLOBAL_ROOTS = u"Merge_Allowed_Source_Roots"
SOURCE_ROOTS_GLOBAL_MODE = u"Merge_Source_Roots_Mode"

# Каталог текущей книги параметров тоже доверенный (относительные пути рядом с книгой).
SOURCE_ROOTS_INCLUDE_BOOK_DIR = True

# smb:// / http:// … в strict: False = запрет.
SOURCE_ROOTS_ALLOW_REMOTE = False

# --- Defaults Windows (плейсхолдер {user} = имя ОС-пользователя) ---
# C:\Users\{user}\**\* , S:\{user}\**\* , весь диск H:\
SOURCE_ROOTS_WINDOWS_TEMPLATES = (
    u"C:\\Users\\{user}",
    u"S:\\",
    u"S:\\{user}",
    u"H:\\",
)

# --- Defaults Linux ---
# /home/{user}/**  и  /home/TN/{user}/**
SOURCE_ROOTS_LINUX_TEMPLATES = (
    u"/home/{user}",
    u"/home/TN/{user}",
)

# Доп. фиксированные корни (всегда, любая ОС) — правьте вручную при необходимости.
SOURCE_ROOTS_EXTRA_FIXED = (
    # u"/mnt/data/share",
    # u"D:\\Trusted",
)
