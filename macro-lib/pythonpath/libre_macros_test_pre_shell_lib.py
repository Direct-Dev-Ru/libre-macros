# -*- coding: utf-8 -*-
"""
Тест: запуск внешнего pre.sh из встроенного Python AlterOffice / LibreOffice.

Параметры — хардкод в libre_macros_test_pre_shell_cfg.py;
исполнение — через libre_macros_pre_shell_lib (как у параметра сбора).
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.712"
import libre_macros_pre_shell_cfg as _ps_cfg
import libre_macros_test_pre_shell_cfg as _cfg
from libre_macros_pre_shell_lib import (
    clip_pre_shell_text,
    run_pre_shell,
)

try:
    unicode
except NameError:
    unicode = str


def _log(msg):
    try:
        print(u"[test_pre_shell] %s" % msg)
    except Exception:
        pass


def _show_message(doc, title, text):
    if doc is None:
        _log(u"%s: %s" % (title, text))
        return
    try:
        import uno
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(u"com.sun.star.awt.Toolkit", ctx)
    except Exception:
        toolkit = None
    if toolkit is None:
        _log(u"%s: %s" % (title, text))
        return
    parent = None
    try:
        parent = doc.getCurrentController().getFrame().getContainerWindow()
    except Exception:
        parent = None
    if parent is None:
        try:
            parent = toolkit.getDesktopWindow()
        except Exception:
            parent = None
    try:
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
        from com.sun.star.awt.MessageBoxType import INFOBOX
        box = toolkit.createMessageBox(parent, INFOBOX, BUTTONS_OK, title, text)
        box.execute()
    except Exception as err:
        _log(u"MessageBox failed: %s; %s" % (err, text))


def run_test_pre_shell(doc=None, *args):
    """Точка входа: запуск хардкод-pre.sh и отчёт в MsgBox / консоль."""
    _ = args
    spec = {
        _ps_cfg.PRE_SHELL_KEY_ARGV: list(getattr(_cfg, u"PRE_SHELL_ARGV", None) or []),
        _ps_cfg.PRE_SHELL_KEY_CWD: unicode(getattr(_cfg, u"PRE_SHELL_CWD", u"") or u""),
        _ps_cfg.PRE_SHELL_KEY_TIMEOUT: getattr(_cfg, u"PRE_SHELL_TIMEOUT_SEC", None),
        _ps_cfg.PRE_SHELL_KEY_ENV: dict(getattr(_cfg, u"PRE_SHELL_ENV_EXTRA", None) or {}),
    }
    try:
        result = run_pre_shell(spec)
        ok = result[u"returncode"] == 0
        text = (
            u"argv: %s\n"
            u"cwd: %s\n"
            u"returncode: %s\n\n"
            u"--- stdout ---\n%s\n\n"
            u"--- stderr ---\n%s"
        ) % (
            u" ".join(result[u"argv"]),
            result[u"cwd"] or u"(не задан)",
            result[u"returncode"],
            clip_pre_shell_text(result[u"stdout"]) or u"(пусто)",
            clip_pre_shell_text(result[u"stderr"]) or u"(пусто)",
        )
        _log(text.replace(u"\n", u" | "))
        if getattr(_cfg, u"PRE_SHELL_SHOW_MESSAGE", True):
            title = u"test_pre_shell: OK" if ok else u"test_pre_shell: ошибка"
            _show_message(doc, title, text)
        return ok
    except Exception as err:
        msg = u"%s: %s" % (type(err).__name__, err)
        _log(msg)
        if getattr(_cfg, u"PRE_SHELL_SHOW_MESSAGE", True):
            _show_message(doc, u"test_pre_shell: сбой", msg)
        return False
