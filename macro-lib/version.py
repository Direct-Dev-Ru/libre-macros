# -*- coding: utf-8 -*-
"""
Наименование: version

Краткий About текущей версии пакета макросов (пункт меню «Версия …»).
"""
MACRO_VERSION = "3.10.721"
from libre_macros_version_lib import show_version_entry as _entry


def show_version(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _entry(doc, *args)


g_exportedScripts = (show_version,)
