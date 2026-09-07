# -*- coding: utf-8 -*-
"""
Диалог ввода пароля AD (маскированный, без сохранения на лист).
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.690"
import unohelper
from com.sun.star.awt import XActionListener

try:
    unicode
except NameError:
    unicode = str

PASSWORD_ECHO_CHAR = 42  # '*'


def _log(msg):
    try:
        print("[ad/dialog] %s" % msg)
    except Exception:
        pass


def _script_context():
    """XSCRIPTCONTEXT: __main__ (AO/LO иногда не кладёт — тогда uno.getComponentContext)."""
    try:
        import __main__

        xsc = getattr(__main__, "XSCRIPTCONTEXT", None)
        if xsc is not None:
            return xsc
    except Exception:
        pass
    try:
        import builtins

        xsc = getattr(builtins, "XSCRIPTCONTEXT", None)
        if xsc is not None:
            return xsc
    except Exception:
        pass
    return None


def _uno_context():
    xsc = _script_context()
    if xsc is not None:
        try:
            return xsc.getComponentContext()
        except Exception:
            pass
    try:
        import uno

        return uno.getComponentContext()
    except Exception:
        return None


def _desktop():
    xsc = _script_context()
    if xsc is not None:
        try:
            desk = xsc.getDesktop()
            if desk is not None:
                return desk
        except Exception:
            pass
    ctx = _uno_context()
    if ctx is None:
        return None
    try:
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    except Exception:
        return None


def _get_toolkit():
    desk = _desktop()
    if desk is not None:
        try:
            tk = desk.getToolkit()
            if tk is not None:
                return tk
        except Exception:
            pass
    ctx = _uno_context()
    if ctx is None:
        return None
    try:
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    except Exception:
        return None


def _dialog_parent(doc=None):
    if doc is not None:
        try:
            return doc.getCurrentController().getFrame().getContainerWindow()
        except Exception:
            pass
    desk = _desktop()
    if desk is not None:
        try:
            frame = desk.getCurrentFrame()
            if frame is not None:
                return frame.getContainerWindow()
        except Exception:
            pass
    return None


def _create_peer(dlg, toolkit, doc=None):
    if dlg is None or toolkit is None:
        return False
    parents = []
    parent_win = _dialog_parent(doc)
    if parent_win is not None:
        parents.append(parent_win)
        try:
            peer = parent_win.getPeer()
            if peer is not None:
                parents.append(peer)
        except Exception:
            pass
    parents.append(None)
    idx = 0
    while idx < len(parents):
        try:
            dlg.createPeer(toolkit, parents[idx])
            return True
        except Exception as err:
            _log("createPeer fail idx=%d: %s" % (idx, err))
        idx += 1
    return False


def prompt_ad_password(doc=None, bind_user=u"", title=None):
    """
    Запросить пароль bind в модальном диалоге (EchoChar='*').
    Возвращает пароль (unicode) или None при отмене. Нигде не сохраняется.
    """
    user_hint = unicode(bind_user or u"").strip()
    if title is None:
        title = u"Пароль Active Directory"

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(
            u"Нет UNO-контекста — запускайте макрос из AlterOffice / LibreOffice Calc."
        )

    if doc is None:
        xsc = _script_context()
        if xsc is not None:
            try:
                doc = xsc.getDocument()
            except Exception:
                doc = None
        if doc is None:
            desk = _desktop()
            if desk is not None:
                try:
                    doc = desk.getCurrentComponent()
                except Exception:
                    doc = None

    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    margin, gap, btn_h, btn_w = 10, 8, 22, 90
    dlg_w = 420
    dm.Title = title
    dm.Width = dlg_w
    dm.PositionX = 120
    dm.PositionY = 120
    content_w = dlg_w - 2 * margin
    y = margin

    hint = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    hint.Name = "HintLabel"
    hint.PositionX = margin
    hint.PositionY = y
    hint.Width = content_w
    hint.Height = 32
    if user_hint:
        hint.Label = u"Пароль для учётной записи:\n%s" % user_hint
    else:
        hint.Label = u"Введите пароль для подключения к Active Directory:"
    dm.insertByName("HintLabel", hint)
    y += 32 + gap

    pwd_label = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    pwd_label.Name = "PwdLabel"
    pwd_label.PositionX = margin
    pwd_label.PositionY = y
    pwd_label.Width = content_w
    pwd_label.Height = 14
    pwd_label.Label = u"Пароль:"
    dm.insertByName("PwdLabel", pwd_label)
    y += 14 + 4

    pwd_edit = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
    pwd_edit.Name = "PasswordEdit"
    pwd_edit.PositionX = margin
    pwd_edit.PositionY = y
    pwd_edit.Width = content_w
    pwd_edit.Height = 24
    pwd_edit.Text = u""
    try:
        pwd_edit.EchoChar = PASSWORD_ECHO_CHAR
    except Exception:
        try:
            pwd_edit.setPropertyValue("EchoChar", PASSWORD_ECHO_CHAR)
        except Exception as err:
            _log("EchoChar not set: %s" % err)
    dm.insertByName("PasswordEdit", pwd_edit)
    y += 24 + gap + 6

    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = "OkBtn"
    ok_btn.Label = u"OK"
    ok_btn.PositionX = dlg_w - margin - 2 * btn_w - gap
    ok_btn.PositionY = y
    ok_btn.Width = btn_w
    ok_btn.Height = btn_h
    ok_btn.DefaultButton = True
    dm.insertByName("OkBtn", ok_btn)

    cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = "CancelBtn"
    cancel_btn.Label = u"Отмена"
    cancel_btn.PositionX = dlg_w - margin - btn_w
    cancel_btn.PositionY = y
    cancel_btn.Width = btn_w
    cancel_btn.Height = btn_h
    dm.insertByName("CancelBtn", cancel_btn)
    dm.Height = y + btn_h + margin

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        raise RuntimeError(u"Не удалось создать peer диалога пароля.")

    state = {u"confirmed": False, u"password": u""}

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "OkBtn":
                try:
                    state[u"password"] = unicode(
                        dialog.getControl("PasswordEdit").getText() or u""
                    )
                except Exception:
                    try:
                        state[u"password"] = unicode(
                            dm.getByName("PasswordEdit").Text or u""
                        )
                    except Exception:
                        state[u"password"] = u""
                state[u"confirmed"] = True
            elif name == "CancelBtn":
                state[u"confirmed"] = False
            else:
                return
            try:
                dialog.endExecute()
            except Exception:
                pass

    handler = _Handler()
    for btn_name in (u"OkBtn", u"CancelBtn"):
        try:
            dialog.getControl(btn_name).addActionListener(handler)
        except Exception as err:
            _log("addActionListener %s: %s" % (btn_name, err))

    try:
        dialog.setVisible(True)
    except Exception:
        pass

    try:
        dialog.execute()
    finally:
        try:
            dialog.dispose()
        except Exception:
            pass

    if not state[u"confirmed"]:
        return None
    return state[u"password"]
