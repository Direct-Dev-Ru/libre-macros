# -*- coding: utf-8 -*-
"""Снятие / восстановление защиты листов Calc при сборе (режим «Текущие листы»)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.719"
try:
    unicode
except NameError:
    unicode = str


def _log(msg):
    try:
        print("[sheet_protect] %s" % msg)
    except Exception:
        pass


def sheet_is_protected(sheet):
    if sheet is None:
        return False
    try:
        return bool(sheet.isProtected())
    except Exception:
        pass
    try:
        return bool(sheet.getPropertyValue("IsProtected"))
    except Exception:
        return False


def sheet_unprotect(sheet, password=u""):
    if sheet is None or not sheet_is_protected(sheet):
        return True
    pwd = unicode(password if password is not None else u"")
    try:
        sheet.unprotect(pwd)
        if not sheet_is_protected(sheet):
            return True
    except Exception as exc:
        _log("unprotect pwd: %s" % exc)
    try:
        sheet.unprotect(u"")
        return not sheet_is_protected(sheet)
    except Exception as exc:
        _log("unprotect empty: %s" % exc)
    return not sheet_is_protected(sheet)


def sheet_protect(sheet, password=u""):
    if sheet is None:
        return False
    if sheet_is_protected(sheet):
        return True
    pwd = unicode(password if password is not None else u"")
    try:
        sheet.protect(pwd)
        return sheet_is_protected(sheet)
    except Exception as exc:
        _log("protect pwd: %s" % exc)
    try:
        sheet.protect(u"")
        return sheet_is_protected(sheet)
    except Exception:
        return False


class MergeSheetProtectSession(object):
    """
    Запомнить защищённые листы, снять защиту; finish — восстановить.
    Ссылки на листы сохраняются (переименование не ломает restore).
    """

    def __init__(self, password=u""):
        self.password = unicode(password if password is not None else u"")
        self._entries = []
        self._active = False

    def begin_doc(self, doc, sheet_names=None):
        """sheet_names=None — все листы книги; иначе только перечисленные имена."""
        if doc is None:
            return False
        self._entries = []
        try:
            sheets = doc.Sheets
            count = int(sheets.getCount())
        except Exception:
            return False
        want = None
        if sheet_names is not None:
            want = set()
            for nm in sheet_names:
                s = unicode(nm or u"").strip()
                if s:
                    want.add(s)
        i = 0
        while i < count:
            try:
                sh = sheets.getByIndex(i)
            except Exception:
                i = i + 1
                continue
            try:
                name = unicode(sh.Name)
            except Exception:
                name = u""
            if want is not None and name not in want:
                i = i + 1
                continue
            was = sheet_is_protected(sh)
            if was:
                if not sheet_unprotect(sh, self.password):
                    _log(
                        "cannot unprotect sheet %r (pwd_len=%s)"
                        % (name, len(self.password or u""))
                    )
                    return False
            self._entries.append((sh, was))
            i = i + 1
        self._active = True
        return True

    def finish(self):
        if not self._active:
            return
        self._active = False
        i = 0
        while i < len(self._entries):
            sh, was = self._entries[i]
            if was:
                sheet_protect(sh, self.password)
            i = i + 1
        self._entries = []
