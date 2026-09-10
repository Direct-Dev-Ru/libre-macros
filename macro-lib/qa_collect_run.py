# -*- coding: utf-8 -*-
"""Headless-точка входа сбора: job JSON → collect_workbooks_for_param_sheet."""
MACRO_VERSION = "3.10.719"
from libre_macros_qa_lib import entry as _entry


def qa_collect_run(*args):
    doc = None
    xsc = None
    try:
        xsc = XSCRIPTCONTEXT
    except Exception:
        xsc = None
    try:
        if xsc is not None:
            doc = xsc.getDocument()
    except Exception:
        doc = None
    return _entry(doc, xsc, *args)


g_exportedScripts = (qa_collect_run,)
