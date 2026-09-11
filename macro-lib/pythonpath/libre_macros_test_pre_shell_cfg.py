# -*- coding: utf-8 -*-
"""Хардкод параметров тестового макроса test_pre_shell (AO subprocess)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.722"
# Аргументы subprocess.run / Popen — без shell=True.
PRE_SHELL_ARGV = [
    u"/bin/sh",
    u"/home/su/projects/python/libre-macros/macro-lib/samples/pre_shell_test.sh",
]

# Рабочий каталог процесса (пусто → не менять cwd).
PRE_SHELL_CWD = u"/tmp"

# Таймаут секунд (None — без лимита).
PRE_SHELL_TIMEOUT_SEC = 60

# Доп. переменные окружения (поверх os.environ процесса AO).
PRE_SHELL_ENV_EXTRA = {
    u"LM_PRE_SHELL_TEST": u"1",
}

# Показывать MsgBox с результатом.
PRE_SHELL_SHOW_MESSAGE = True
