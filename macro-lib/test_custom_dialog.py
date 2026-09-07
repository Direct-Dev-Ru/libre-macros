# -*- coding: utf-8 -*-
"""
Тест кастомного диалога (LibreOffice / AlterOffice).

Запуск:
  test_custom_dialog.test_custom_dialog
"""
from libre_macros_test_dialog_lib import run_test_dialog as _run_test_dialog


def test_custom_dialog(*args):
    return _run_test_dialog(*args)


g_exportedScripts = (test_custom_dialog,)
