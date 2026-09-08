# -*- coding: utf-8 -*-
"""
Макрос: таблицы / DatabaseRange → диапазоны или стиль TableStyleLight1.

Диалог — здесь (как merge_ui_dialog_confirm_query в collect_workbooks).
Логика — pythonpath/libre_macros_convert_tables_lib.py (импорт после диалога).
"""
MACRO_VERSION = "3.10.711"
import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO_CANCEL
from com.sun.star.awt.MessageBoxType import QUERYBOX

from libre_macros_convert_tables_cfg import MACRO_VERSION, TARGET_TABLE_STYLE

try:
    from com.sun.star.awt.MessageBoxResults import YES, NO
except Exception:
    YES, NO = 2, 3


def _log(msg):
    try:
        print("[convert_tables] %s" % msg)
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
    """Цепочка родителей как _merge_ui_dlg_create_peer в collect_workbooks."""
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


def _ask_messagebox_fallback(items_count=None):
    try:
        toolkit = _get_toolkit()
        parent = _dialog_parent()
        if items_count is None:
            text = (
                u"Да — преобразовать в диапазоны (с автофильтром)\n"
                u"Нет — применить стиль таблицы %s\n"
                u"Отмена — ничего не делать"
            ) % TARGET_TABLE_STYLE
        else:
            text = (
                u"Найдено объектов: %d.\n\n"
                u"Да — преобразовать в диапазоны (с автофильтром)\n"
                u"Нет — применить стиль таблицы %s\n"
                u"Отмена — ничего не делать"
            ) % (int(items_count), TARGET_TABLE_STYLE)
        box = toolkit.createMessageBox(
            parent, QUERYBOX, BUTTONS_YES_NO_CANCEL,
            u"Таблицы / диапазоны БД", text,
        )
        result = box.execute()
        if result == YES:
            return {"action": "convert"}
        if result == NO:
            return {"action": "style", "style": TARGET_TABLE_STYLE}
        return None
    except Exception as err:
        _log("MessageBox fallback failed: %s" % err)
        return None


def ask_convert_or_style(doc, items_count=None):
    """
    Диалог по образцу collect_workbooks._merge_ui_dialog_confirm_query.
    Без обращения к doc.DatabaseRanges до execute (на AO иначе зависает).
    """
    _log("ask_convert_or_style: start")
    try:
        toolkit = _get_toolkit()
        if toolkit is None:
            return _ask_messagebox_fallback(items_count)
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sm = ctx.getServiceManager()
        dm = sm.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx
        )
        m, g, btn_h, btn_w = 10, 8, 22, 96
        dw = 480
        style_name = TARGET_TABLE_STYLE
        if items_count is None:
            summary = (
                u"Преобразование smart-table / DatabaseRange в этой книге.\n\n"
                u"«В диапазоны» — снять smart-table, оставить автофильтр.\n"
                u"«Стиль %s» — применить стиль таблицы через правку XML на диске.\n"
                u"«Отмена» — ничего не делать."
            ) % style_name
        else:
            summary = (
                u"Найдено объектов DatabaseRange / Table: %d\n\n"
                u"«В диапазоны» — снять smart-table, оставить автофильтр.\n"
                u"«Стиль %s» — применить стиль таблицы через правку XML на диске.\n"
                u"«Отмена» — ничего не делать."
            ) % (int(items_count), style_name)
        dm.Title = u"Таблицы / диапазоны БД (v%s)" % MACRO_VERSION
        dm.Width = dw
        dm.PositionX = 80
        dm.PositionY = 60
        content_w = dw - 2 * m
        text_h = 200
        y = m
        summary_m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        summary_m.Name = "SummaryEdit"
        summary_m.PositionX = m
        summary_m.PositionY = y
        summary_m.Width = content_w
        summary_m.Height = text_h
        summary_m.MultiLine = True
        summary_m.ReadOnly = True
        try:
            summary_m.VScroll = True
        except Exception:
            pass
        dm.insertByName("SummaryEdit", summary_m)
        y = y + text_h + g
        for spec in (
            ("ConvertBtn", u"В диапазоны", m, btn_w + 24),
            ("StyleBtn", u"Стиль %s" % style_name, m + btn_w + 24 + g, btn_w + 40),
            ("CancelBtn", u"Отмена", dw - m - btn_w, btn_w),
        ):
            bm = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
            bm.Name = spec[0]
            bm.Label = spec[1]
            bm.PositionX = spec[2]
            bm.PositionY = y
            bm.Width = spec[3]
            bm.Height = btn_h
            dm.insertByName(spec[0], bm)
        dm.Height = y + btn_h + m

        dialog = sm.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", ctx
        )
        dialog.setModel(dm)
        if not _create_peer(dialog, toolkit, doc=doc):
            _log("createPeer failed → MessageBox")
            return _ask_messagebox_fallback(items_count)
        try:
            dialog.getControl("SummaryEdit").setText(summary)
        except Exception:
            try:
                dm.getByName("SummaryEdit").Text = summary
            except Exception:
                pass

        state = {"action": None}

        class _Handler(unohelper.Base, XActionListener):
            def disposing(self, event):
                pass

            def actionPerformed(self, event):
                try:
                    name = str(event.Source.getModel().Name)
                except Exception:
                    return
                if name == "ConvertBtn":
                    state["action"] = "convert"
                elif name == "StyleBtn":
                    state["action"] = "style"
                elif name == "CancelBtn":
                    state["action"] = None
                else:
                    return
                try:
                    dialog.endExecute()
                except Exception:
                    pass

        handler = _Handler()
        for btn_name in ("ConvertBtn", "StyleBtn", "CancelBtn"):
            try:
                dialog.getControl(btn_name).addActionListener(handler)
            except Exception:
                pass
        try:
            dialog.setVisible(True)
        except Exception:
            pass
        _log("ask_convert_or_style: dialog.execute() begin")
        try:
            dialog.execute()
        except Exception as err:
            _log("ask_convert_or_style: execute failed: %s" % err)
            try:
                dialog.dispose()
            except Exception:
                pass
            return _ask_messagebox_fallback(items_count)
        _log("ask_convert_or_style: execute done action=%s" % state["action"])
        try:
            dialog.dispose()
        except Exception:
            pass
        if state["action"] == "convert":
            return {"action": "convert"}
        if state["action"] == "style":
            return {"action": "style", "style": style_name}
        return None
    except Exception as err:
        _log("ask_convert_or_style failed: %s" % err)
        return _ask_messagebox_fallback(items_count)


def convert_tables_to_ranges(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    if doc is None:
        _log("нет активного документа")
        return None

    # Диалог до любого обращения к doc.DatabaseRanges (иначе execute зависает на AO).
    choice = ask_convert_or_style(doc)
    if choice is None:
        _log("отмена пользователем")
        return None

    from libre_macros_convert_tables_lib import convert_tables_to_ranges as _run_convert
    _run_convert(doc, choice=choice, *args)
    return None


g_exportedScripts = (convert_tables_to_ranges,)
