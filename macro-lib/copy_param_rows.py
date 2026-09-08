# -*- coding: utf-8 -*-
"""Копирование строк листа параметров сбора."""
MACRO_VERSION = "3.10.705"
from libre_macros_copy_param_rows_lib import copy_param_rows_entry as _entry


def copy_param_rows(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _entry(doc)


CONTEXT_CELL_MENU_KEY = "copy_param_rows"
CONTEXT_CELL_MENU = (
    "submenu#PythonMacros",
    "separator#before",
    "run_function#copy_param_rows;module#copy_param_rows.py;display_name#Копировать строку(и) параметров;order#28",
    "separator#after",
)
g_exportedScripts = (copy_param_rows,)
