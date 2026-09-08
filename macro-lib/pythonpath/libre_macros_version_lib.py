# -*- coding: utf-8 -*-
"""Краткий About текущей версии пакета libre-macros."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.708"
try:
    unicode
except NameError:
    unicode = str


def _version_string():
    try:
        from libre_macros_lib import MACRO_VERSION as ver
        return unicode(ver or MACRO_VERSION)
    except Exception:
        return unicode(MACRO_VERSION)


def _show_message(doc, title, text):
    try:
        import uno
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    except Exception:
        toolkit = None
    if toolkit is None:
        try:
            print("%s: %s" % (title, text))
        except Exception:
            pass
        return
    parent = None
    if doc is not None:
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
        try:
            print("version MessageBox failed: %s; %s" % (err, text))
        except Exception:
            pass


def show_version_entry(doc=None, *args):
    """Показать краткий about версии макросов."""
    _ = args
    ver = _version_string()
    text = (
        u"AlterOffice Macros\n\n"
        u"Версия: %s\n\n"
        u"Пакет макросов для AlterOffice Calc.\n"
    ) % ver
    _show_message(doc, u"Версия %s" % ver, text)
    return True
