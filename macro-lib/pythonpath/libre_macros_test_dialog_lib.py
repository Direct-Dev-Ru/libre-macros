# -*- coding: utf-8 -*-
"""
Минимальный кастомный диалог для отладки в LibreOffice / AlterOffice.

Паттерн createPeer — как convert_tables_to_ranges / param_wizard.
"""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.718"
import unohelper
from com.sun.star.awt import XActionListener

try:
    unicode
except NameError:
    unicode = str


def _log(msg):
    try:
        print("[test_custom_dialog] %s" % msg)
    except Exception:
        pass


def _get_toolkit():
    try:
        return XSCRIPTCONTEXT.getDesktop().getToolkit()
    except Exception:
        pass
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    except Exception:
        return None


def _dialog_parent(doc=None):
    if doc is None:
        try:
            doc = XSCRIPTCONTEXT.getDocument()
        except Exception:
            doc = None
    if doc is not None:
        try:
            return doc.getCurrentController().getFrame().getContainerWindow()
        except Exception:
            pass
    try:
        frame = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()
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
    if doc is not None:
        try:
            w = doc.getCurrentController().getFrame().getContainerWindow()
            if w is not None and w not in parents:
                parents.append(w)
        except Exception:
            pass
    try:
        frame = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()
        if frame is not None:
            w = frame.getContainerWindow()
            if w is not None and w not in parents:
                parents.append(w)
    except Exception:
        pass
    parents.append(None)
    i = 0
    while i < len(parents):
        try:
            dlg.createPeer(toolkit, parents[i])
            _log("createPeer ok parent_idx=%d" % i)
            return True
        except Exception as err:
            _log("createPeer fail parent_idx=%d: %s" % (i, err))
        i = i + 1
    return False


def _active_cell(doc):
    if doc is None:
        return None
    try:
        controller = doc.getCurrentController()
    except Exception:
        controller = None
    if controller is None:
        return None
    try:
        sel = controller.getSelection()
    except Exception:
        sel = None
    if sel is not None:
        try:
            if sel.supportsService("com.sun.star.table.Cell"):
                return sel
        except Exception:
            pass
        try:
            if sel.supportsService("com.sun.star.table.CellRange"):
                return sel.getCellByPosition(0, 0)
        except Exception:
            pass
    try:
        sheet = controller.getActiveSheet()
        return sheet.getCellByPosition(0, 0)
    except Exception:
        return None


def _insert_into_active_cell(doc, text):
    cell = _active_cell(doc)
    if cell is None:
        _log("insert: active cell not found")
        return False
    value = unicode(text or u"")
    try:
        cell.setString(value)
        _log("insert: setString ok, len=%d" % len(value))
        return True
    except Exception as err:
        _log("insert: setString failed: %s" % err)
        return False


def show_test_dialog(doc=None):
    """
    Кастомный диалог: поле ввода + OK / Отмена.
    OK — вставляет текст в активную ячейку (или A1).
    Возвращает введённый текст или None.
    """
    _log("show_test_dialog: start")
    if doc is None:
        try:
            doc = XSCRIPTCONTEXT.getDocument()
        except Exception:
            doc = None
    toolkit = _get_toolkit()
    if toolkit is None:
        _log("show_test_dialog: toolkit is None")
        return None
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sm = ctx.getServiceManager()
        dm = sm.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx
        )
        m, g, btn_h, btn_w = 10, 8, 22, 90
        dw = 420
        dm.Title = u"Тест кастомного диалога"
        dm.Width = dw
        dm.PositionX = 120
        dm.PositionY = 80
        content_w = dw - 2 * m
        y = m

        hint_m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        hint_m.Name = "HintLabel"
        hint_m.PositionX = m
        hint_m.PositionY = y
        hint_m.Width = content_w
        hint_m.Height = 28
        hint_m.Label = u"Введите текст и нажмите «Вставить в ячейку»:"
        dm.insertByName("HintLabel", hint_m)
        y = y + 28 + g

        edit_m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        edit_m.Name = "TextEdit"
        edit_m.PositionX = m
        edit_m.PositionY = y
        edit_m.Width = content_w
        edit_m.Height = 24
        edit_m.Text = u"Привет из кастомного диалога"
        dm.insertByName("TextEdit", edit_m)
        y = y + 24 + g + 4

        ok_m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        ok_m.Name = "OkBtn"
        ok_m.Label = u"Вставить в ячейку"
        ok_m.PositionX = dw - m - 2 * btn_w - g
        ok_m.PositionY = y
        ok_m.Width = btn_w + 36
        ok_m.Height = btn_h
        dm.insertByName("OkBtn", ok_m)

        cancel_m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        cancel_m.Name = "CancelBtn"
        cancel_m.Label = u"Отмена"
        cancel_m.PositionX = dw - m - btn_w
        cancel_m.PositionY = y
        cancel_m.Width = btn_w
        cancel_m.Height = btn_h
        dm.insertByName("CancelBtn", cancel_m)
        dm.Height = y + btn_h + m

        dialog = sm.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", ctx
        )
        dialog.setModel(dm)
        if not _create_peer(dialog, toolkit, doc=doc):
            _log("show_test_dialog: createPeer failed")
            return None

        state = {"confirmed": False, "text": u""}

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
                        state["text"] = unicode(
                            dialog.getControl("TextEdit").getText() or u""
                        )
                    except Exception:
                        try:
                            state["text"] = unicode(
                                dm.getByName("TextEdit").Text or u""
                            )
                        except Exception:
                            state["text"] = u""
                    state["confirmed"] = True
                elif name == "CancelBtn":
                    state["confirmed"] = False
                else:
                    return
                try:
                    dialog.endExecute()
                except Exception:
                    pass

        handler = _Handler()
        for btn_name in ("OkBtn", "CancelBtn"):
            try:
                dialog.getControl(btn_name).addActionListener(handler)
            except Exception as err:
                _log("addActionListener %s failed: %s" % (btn_name, err))

        try:
            dialog.setVisible(True)
        except Exception:
            pass

        _log("show_test_dialog: execute() begin")
        try:
            dialog.execute()
        except Exception as err:
            _log("show_test_dialog: execute failed: %s" % err)
            try:
                dialog.dispose()
            except Exception:
                pass
            return None
        _log(
            "show_test_dialog: execute done confirmed=%s"
            % bool(state["confirmed"])
        )
        try:
            dialog.dispose()
        except Exception:
            pass

        if not state["confirmed"]:
            return None
        text = state["text"]
        if doc is not None:
            _insert_into_active_cell(doc, text)
        return text
    except Exception as err:
        _log("show_test_dialog failed: %s" % err)
        return None


def run_test_dialog(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        doc = None
    result = show_test_dialog(doc=doc)
    if result is None:
        _log("run_test_dialog: cancelled or failed")
    else:
        _log("run_test_dialog: inserted %r" % result)
    return result
