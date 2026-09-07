# -*- coding: utf-8 -*-
"""
Тест запуска внешнего pre.sh из Python AlterOffice / LibreOffice.

Параметры (argv, cwd, timeout) — хардкод в
  pythonpath/libre_macros_test_pre_shell_cfg.py

Запуск из меню макросов:
  test_pre_shell.test_pre_shell
"""
MACRO_VERSION = "3.10.499"
from libre_macros_test_pre_shell_lib import run_test_pre_shell as _entry


def test_pre_shell(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _entry(doc, *args)


g_exportedScripts = (test_pre_shell,)
