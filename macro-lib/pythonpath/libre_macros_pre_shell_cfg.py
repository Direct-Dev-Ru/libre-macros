# -*- coding: utf-8 -*-
"""Константы параметра «Предварительный_скрипт» (pre-shell)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.715"
# Имя строки на листе параметров (синхрон с P_MERGE_PRE_SHELL в collect_cfg).
P_MERGE_PRE_SHELL = u"Предварительный_скрипт"

# Ключи JSON в колонке B.
PRE_SHELL_KEY_ARGV = u"argv"
PRE_SHELL_KEY_CWD = u"cwd"
PRE_SHELL_KEY_TIMEOUT = u"timeout"
PRE_SHELL_KEY_ENV = u"env"
# True (по умолчанию): убрать PYTHONHOME/PYTHONPATH и пути AlterOffice/LibreOffice
# из PATH/LD_LIBRARY_PATH — иначе дочерний python ломается (fork_exec TypeError).
PRE_SHELL_KEY_CLEAN_ENV = u"clean_env"

# Значения по умолчанию для пустых полей формы (не подставляются в JSON без ввода).
PRE_SHELL_DEFAULT_TIMEOUT_SEC = 60
PRE_SHELL_DEFAULT_CWD = u""
PRE_SHELL_DEFAULT_CLEAN_ENV = True

# Журнал согласий на запуск (каталог ~/.config/libre-macros/pre_shell_consent/).
PRE_SHELL_CONSENT_SUBDIR = u"pre_shell_consent"

# Диалог ошибки / запрета запуска (крупный красный текст).
PRE_SHELL_ERROR_TEXT_COLOR = 0xB71C1C
PRE_SHELL_ERROR_FONT_HEIGHT = 16
PRE_SHELL_ERROR_DIALOG_TITLE = u"Предварительный_скрипт — невозможно выполнить"
# Диалог «Внимание!!!»: шрифт в окне описания команды (pt).
PRE_SHELL_CONFIRM_BODY_FONT_HEIGHT = 12

# Гейт запуска: OS env + зашифрованная глобальная переменная (см. verify_pre_shell_allow_gate).
PRE_SHELL_ALLOW_ENV_NAME = u"MERGE_ALLOW_PRE_SCRIPT"
PRE_SHELL_ALLOW_GLOBAL_NAME = u"Merge_Allow_Pre_Scripts"


def fold_pre_shell_allow_name(name):
    """
    Нормализация имени глобальной переменной допуска pre-shell:
    регистр игнорируется; «_», «-» и прочие не-буквы/цифры выкидываются.
    Merge_Allow_Pre_Scripts / merge-allow-pre-scripts / MergeAllowPreScripts → одно.
    """
    try:
        unicode
    except NameError:
        _u = str
    else:
        _u = unicode
    s = _u(name or u"")
    chars = []
    i = 0
    while i < len(s):
        ch = s[i]
        i += 1
        try:
            if ch.isalnum():
                chars.append(ch)
        except Exception:
            pass
    folded = u"".join(chars)
    try:
        return folded.casefold()
    except Exception:
        return folded.lower()


PRE_SHELL_ALLOW_GLOBAL_NAME_FOLD = fold_pre_shell_allow_name(PRE_SHELL_ALLOW_GLOBAL_NAME)


def is_pre_shell_allow_global_name(name):
    """True, если имя — канонический Merge_Allow_Pre_Scripts (гибкое написание)."""
    return fold_pre_shell_allow_name(name) == PRE_SHELL_ALLOW_GLOBAL_NAME_FOLD

# Ключи env, значения которых маскируются в журнале.
PRE_SHELL_ENV_SECRET_MARKERS = (
    u"password",
    u"passwd",
    u"secret",
    u"token",
    u"credential",
)

PRE_SHELL_UNSET_ENV_KEYS = (
    u"PYTHONHOME",
    u"PYTHONPATH",
    u"PYTHONSTARTUP",
    u"PYTHONUSERBASE",
    u"PYTHONEXECUTABLE",
    u"PYTHON_EGG_CACHE",
)

# Подстроки путей (casefold) — вычищать из PATH / LD_LIBRARY_PATH.
PRE_SHELL_OFFICE_PATH_MARKERS = (
    u"alteroffice",
    u"libreoffice",
    u"openoffice",
    u"/program/python-core-",
)

# Пример для хинта / пустой формы.
PRE_SHELL_EXAMPLE_JSON = (
    u'{"argv":["/bin/bash","/path/to/run.sh"],'
    u'"cwd":"/path/to/ad-shell","timeout":600,'
    u'"clean_env":true,'
    u'"env":{"AD_BIND_PASSWORD":"…"}}'
)
