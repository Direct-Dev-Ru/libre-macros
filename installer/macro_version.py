# -*- coding: utf-8 -*-
"""Чтение канонической версии макроса из исходников (без LibreOffice)."""
from __future__ import print_function

import os
import re

MACRO_VERSION_ASSIGN_RE = re.compile(
    r'^MACRO_VERSION\s*=\s*["\']([^"\']+)["\']',
    re.MULTILINE,
)

# Порядок приоритета: version.txt → главная библиотека → точка входа макроса.
CANONICAL_SOURCES = (
    ("macro-lib", "version.txt"),
    ("macro-lib", "collect_workbooks.py"),
    ("macro-lib", "pythonpath", "libre_macros_lib.py"),
    ("macro-lib", "functions_pp.py"),
)


def _read_version_from_py(path):
    try:
        with open(path, encoding="utf-8") as f:
            text = f.read()
    except (IOError, OSError):
        return ""
    m = MACRO_VERSION_ASSIGN_RE.search(text)
    return m.group(1).strip() if m else ""


def _read_version_from_txt(path):
    try:
        with open(path, encoding="utf-8") as f:
            return f.read().strip()
    except (IOError, OSError):
        return ""


def read_canonical_macro_version(repo_root="."):
    """
    Версия из кода макросов.

    Источник по умолчанию — MACRO_VERSION в libre_macros_lib.py;
    далее collect_workbooks.py, functions_pp.py, version.txt.
    """
    root = os.path.abspath(repo_root or ".")
    for parts in CANONICAL_SOURCES:
        path = os.path.join(root, *parts)
        if not os.path.isfile(path):
            continue
        if path.endswith(".txt"):
            version = _read_version_from_txt(path)
        else:
            version = _read_version_from_py(path)
        if version:
            return version
    return "0.0.0"
