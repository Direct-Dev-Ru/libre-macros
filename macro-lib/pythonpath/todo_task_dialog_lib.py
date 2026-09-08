# -*- coding: utf-8 -*-
"""Диалог редактирования строки задачи (todo_task_edit)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.709"
import re
import uno
import unohelper
from com.sun.star.awt import XActionListener

import todo_task_cfg as _cfg

try:
    import todo_task_plan_cfg as _pcfg
except Exception:
    _pcfg = None

try:
    from libre_macros_ui_theme import (
        DARK_BLUE_TITLE_BG,
        DARK_BLUE_TITLE_FG,
        SOFT_GRAY_BAR_NAME,
        apply_soft_gray_dark_blue_theme,
        try_paint_titlebar,
    )
except Exception:
    DARK_BLUE_TITLE_BG = 0x0D47A1
    DARK_BLUE_TITLE_FG = 0xFFFFFF
    SOFT_GRAY_BAR_NAME = "SoftGrayTitleBar"

    def apply_soft_gray_dark_blue_theme(dm, title_text=None):
        return False

    def try_paint_titlebar(dlg, title_bg=None, title_fg=None):
        return False

try:
    unicode
except NameError:
    unicode = str


def _todo_dialog_font_pt(doc=None):
    try:
        from todo_task_settings_lib import get_dialog_font_pt

        return int(get_dialog_font_pt(doc=doc))
    except Exception:
        return int(getattr(_cfg, "DEFAULT_DIALOG_FONT_PT", 11))


def _set_model_font_height(model, height, bold=False):
    if model is None:
        return
    try:
        h = int(height)
    except (TypeError, ValueError):
        return
    try:
        fd = model.FontDescriptor
        fd.Height = h
        if bold:
            try:
                fd.Weight = 150
            except Exception:
                pass
        model.FontDescriptor = fd
        return
    except Exception:
        pass
    try:
        model.FontHeight = h
    except Exception:
        pass
    if bold:
        try:
            model.FontWeight = 150
        except Exception:
            pass


def _apply_todo_dialog_fonts(dm, height=None, doc=None):
    """Применить размер шрифта ко всем контролам модели диалога todo."""
    if dm is None:
        return
    if height is None:
        height = _todo_dialog_font_pt(doc=doc)
    try:
        h = int(height)
    except (TypeError, ValueError):
        h = int(getattr(_cfg, "DEFAULT_DIALOG_FONT_PT", 11))
    lo = int(getattr(_cfg, "DIALOG_FONT_PT_MIN", 8))
    hi = int(getattr(_cfg, "DIALOG_FONT_PT_MAX", 24))
    if h < lo:
        h = lo
    if h > hi:
        h = hi
    try:
        names = list(dm.getElementNames())
    except Exception:
        return
    bar = unicode(SOFT_GRAY_BAR_NAME)
    i = 0
    while i < len(names):
        name = names[i]
        i = i + 1
        try:
            ctl = dm.getByName(name)
        except Exception:
            continue
        bold = unicode(name) == bar
        _set_model_font_height(ctl, h, bold=bold)


def _theme_todo_dialog(dm, title_text=None, doc=None):
    """Тема todo + размер шрифта из настроек."""
    try:
        apply_soft_gray_dark_blue_theme(dm, title_text=title_text)
    except Exception:
        pass
    _apply_todo_dialog_fonts(dm, doc=doc)


def _log(msg):
    try:
        from todo_task_settings_lib import get_console_log_enabled
        if not get_console_log_enabled():
            return
    except Exception:
        pass
    try:
        print("[todo_task/dialog] %s" % msg)
    except Exception:
        pass


def _plan_trace(msg):
    """Лог визарда плана в stdout (всегда, для отладки)."""
    try:
        import sys

        sys.stdout.write(u"[todo_task/plan] %s\n" % unicode(msg))
        sys.stdout.flush()
    except Exception:
        try:
            print(u"[todo_task/plan] %s" % msg)
        except Exception:
            pass


def _script_context():
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


def _create_peer(dlg, toolkit, doc=None, parent_dialog=None):
    if dlg is None or toolkit is None:
        return False
    parents = []
    if parent_dialog is not None:
        try:
            peer = parent_dialog.getPeer()
            if peer is not None:
                parents.append(peer)
        except Exception:
            pass
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


def _add_label(dm, name, label, x, y, w, h, align=None, vertical_align=None):
    ctl = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    ctl.Name = name
    ctl.Label = unicode(label)
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    if align is not None:
        try:
            ctl.Align = int(align)
        except Exception:
            pass
    if vertical_align is not None:
        try:
            ctl.VerticalAlign = int(vertical_align)
        except Exception:
            pass
    dm.insertByName(name, ctl)
    return ctl


def _add_edit(dm, name, x, y, w, h, multiline=False, password=False):
    ctl = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    if multiline:
        try:
            ctl.MultiLine = True
        except Exception:
            pass
        try:
            ctl.VScroll = True
        except Exception:
            pass
    if password:
        try:
            ctl.EchoChar = 42
        except Exception:
            pass
    dm.insertByName(name, ctl)
    return ctl


def _add_checkbox(dm, name, label, x, y, w, h):
    ctl = dm.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
    ctl.Name = name
    ctl.Label = unicode(label)
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    dm.insertByName(name, ctl)
    return ctl


def _add_combo(dm, name, x, y, w, h, items):
    ctl = dm.createInstance("com.sun.star.awt.UnoControlComboBoxModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    ctl.Dropdown = True
    try:
        ctl.LineCount = 12
    except Exception:
        pass
    try:
        ctl.Autocomplete = False
    except Exception:
        pass
    try:
        ctl.StringItemList = tuple(unicode(x) for x in (items or []))
    except Exception:
        pass
    dm.insertByName(name, ctl)
    return ctl


def _add_dropdown(dm, name, x, y, w, h, items):
    """Выпадающий список (ListBox+Dropdown) — в AO/LO надёжнее ComboBox."""
    ctl = dm.createInstance("com.sun.star.awt.UnoControlListBoxModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    try:
        ctl.Dropdown = True
    except Exception:
        pass
    try:
        ctl.MultiSelection = False
    except Exception:
        pass
    n_items = len(items or [])
    try:
        ctl.LineCount = max(2, min(8, n_items + 1))
    except Exception:
        pass
    try:
        ctl.ReadOnly = False
    except Exception:
        pass
    try:
        ctl.StringItemList = tuple(unicode(x) for x in (items or []))
    except Exception:
        pass
    dm.insertByName(name, ctl)
    return ctl


def _select_dropdown_item(dialog, dm, name, label, items=None):
    label = unicode(label or u"")
    ctl = None
    try:
        ctl = dialog.getControl(name)
    except Exception:
        ctl = None
    if ctl is not None:
        try:
            ctl.selectItem(label, True)
            return
        except Exception:
            pass
        pos = -1
        src = list(items or [])
        if not src:
            try:
                src = list(ctl.getModel().StringItemList or [])
            except Exception:
                src = []
        i = 0
        while i < len(src):
            if unicode(src[i]) == label:
                pos = i
                break
            i = i + 1
        if pos >= 0:
            try:
                ctl.selectItemPos(pos, True)
                return
            except Exception:
                pass
    try:
        dm.getByName(name).SelectedItems = (label,)
    except Exception:
        pass
    try:
        dm.getByName(name).Text = label
    except Exception:
        pass


def _add_listbox(dm, name, x, y, w, h, items):
    ctl = dm.createInstance("com.sun.star.awt.UnoControlListBoxModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    try:
        ctl.Dropdown = False
    except Exception:
        pass
    try:
        ctl.MultiSelection = False
    except Exception:
        pass
    try:
        # ReadOnly=True в AO/LO мешает выбору пункта для редактирования
        ctl.ReadOnly = False
    except Exception:
        pass
    try:
        ctl.StringItemList = tuple(unicode(x) for x in (items or []))
    except Exception:
        pass
    dm.insertByName(name, ctl)
    return ctl


def _get_control_text(dialog, dm, name):
    try:
        return unicode(dialog.getControl(name).getText() or u"")
    except Exception:
        pass
    try:
        return unicode(dm.getByName(name).Text or u"")
    except Exception:
        return u""


def _set_control_text(dialog, dm, name, text):
    text = unicode(text if text is not None else u"")
    try:
        dialog.getControl(name).setText(text)
        return
    except Exception:
        pass
    try:
        dm.getByName(name).Text = text
    except Exception:
        pass


def _set_control_visible(dialog, dm, name, visible):
    flag = bool(visible)
    try:
        model = dm.getByName(name)
        try:
            model.EnableVisible = flag
        except Exception:
            try:
                model.setPropertyValue("EnableVisible", flag)
            except Exception:
                pass
    except Exception:
        pass
    try:
        ctl = dialog.getControl(name)
        try:
            ctl.setVisible(flag)
        except Exception:
            pass
    except Exception:
        pass


def _add_button(dm, name, label, x, y, w, h, default=False):
    ctl = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ctl.Name = name
    ctl.Label = unicode(label)
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    if default:
        try:
            ctl.DefaultButton = True
        except Exception:
            pass
    try:
        ctl.PushButtonType = 0
    except Exception:
        pass
    dm.insertByName(name, ctl)
    return ctl


def _dialog_warn(doc, text, title=u"Задачи"):
    show_todo_task_message(doc, text, title=title)


def show_todo_task_message(doc=None, text=u"", title=u"Задачи"):
    """Кастомный message box (OK) с тёмно-синей темой. Возвращает True."""
    return _show_todo_task_msg_dialog(
        doc=doc, text=text, title=title, with_cancel=False
    )


def show_todo_task_confirm(doc=None, text=u"", title=u"Задачи"):
    """Кастомный confirm (OK/Отмена) с тёмно-синей темой. True = OK."""
    return _show_todo_task_msg_dialog(
        doc=doc, text=text, title=title, with_cancel=True
    )


def _show_todo_task_msg_dialog(doc=None, text=u"", title=u"Задачи", with_cancel=False):
    text = unicode(text or u"")
    title = unicode(title or u"Задачи")
    ctx = _uno_context()
    toolkit = _get_toolkit()
    if ctx is None or toolkit is None:
        _log("%s: %s" % (title, text))
        return not with_cancel

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_w = int(getattr(_cfg, "MSG_BTN_W", _cfg.EDIT_BTN_W))
    btn_h = int(getattr(_cfg, "MSG_BTN_H", _cfg.EDIT_BTN_H))
    dlg_w = int(getattr(_cfg, "MSG_DIALOG_WIDTH", 420))
    text_h = int(getattr(_cfg, "MSG_TEXT_H", 72))
    content_w = dlg_w - 2 * m

    dm.Title = title
    dm.Width = dlg_w
    dm.PositionX = 140
    dm.PositionY = 160

    y = m
    msg = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    msg.Name = "MsgText"
    msg.Label = text
    msg.PositionX = m
    msg.PositionY = y
    msg.Width = content_w
    msg.Height = text_h
    try:
        msg.MultiLine = True
    except Exception:
        pass
    try:
        msg.Align = 0
    except Exception:
        pass
    try:
        msg.VerticalAlign = 0
    except Exception:
        pass
    dm.insertByName("MsgText", msg)
    y += text_h + gap + 4

    if with_cancel:
        ok_x = dlg_w - m - 2 * btn_w - gap
        cancel_x = dlg_w - m - btn_w
    else:
        ok_x = dlg_w - m - btn_w
        cancel_x = None

    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = "OkBtn"
    ok_btn.Label = unicode(_cfg.BTN_OK_LABEL)
    ok_btn.PositionX = ok_x
    ok_btn.PositionY = y
    ok_btn.Width = btn_w
    ok_btn.Height = btn_h
    try:
        ok_btn.DefaultButton = True
    except Exception:
        pass
    dm.insertByName("OkBtn", ok_btn)

    if with_cancel:
        cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        cancel_btn.Name = "CancelBtn"
        cancel_btn.Label = unicode(_cfg.BTN_CANCEL_LABEL)
        cancel_btn.PositionX = cancel_x
        cancel_btn.PositionY = y
        cancel_btn.Width = btn_w
        cancel_btn.Height = btn_h
        dm.insertByName("CancelBtn", cancel_btn)

    dm.Height = max(
        int(getattr(_cfg, "MSG_DIALOG_MIN_H", 120)),
        y + btn_h + m,
    )

    try:
        _theme_todo_dialog(dm, title_text=title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        _log("%s: %s" % (title, text))
        return not with_cancel
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass

    state = {u"ok": False}

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "OkBtn":
                state[u"ok"] = True
            elif name == "CancelBtn":
                state[u"ok"] = False
            else:
                return
            try:
                dialog.endExecute()
            except Exception:
                pass

    handler = _Handler()
    btn_names = (u"OkBtn", u"CancelBtn") if with_cancel else (u"OkBtn",)
    for btn_name in btn_names:
        try:
            dialog.getControl(btn_name).addActionListener(handler)
        except Exception as err:
            _log("msg listener %s: %s" % (btn_name, err))

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
    return bool(state[u"ok"]) if with_cancel else True


def _set_window_enabled(win, enabled):
    """Включить/выключить окно (setEnable и свойство Enable)."""
    if win is None:
        return False
    on = bool(enabled)
    ok = False
    try:
        if hasattr(win, "setEnable"):
            win.setEnable(on)
            ok = True
    except Exception:
        pass
    try:
        win.Enable = on
        ok = True
    except Exception:
        pass
    return ok


def restore_todo_doc_ui(doc=None):
    """
    Вернуть доступность окна книги после progress/диалогов.
    Без этого после Enable=False книга остаётся «замороженной».
    """
    wins = []
    parent = _dialog_parent(doc)
    if parent is not None:
        wins.append(parent)
        try:
            peer = parent.getPeer()
            if peer is not None and peer not in wins:
                wins.append(peer)
        except Exception:
            pass
    if doc is not None:
        try:
            frame = doc.getCurrentController().getFrame()
            cw = frame.getContainerWindow()
            if cw is not None and cw not in wins:
                wins.append(cw)
            try:
                comp = frame.getComponentWindow()
                if comp is not None and comp not in wins:
                    wins.append(comp)
            except Exception:
                pass
        except Exception:
            pass
    i = 0
    while i < len(wins):
        _set_window_enabled(wins[i], True)
        try:
            if hasattr(wins[i], "setFocus"):
                wins[i].setFocus()
        except Exception:
            pass
        i = i + 1
    if doc is not None:
        try:
            frame = doc.getCurrentController().getFrame()
            frame.activate()
        except Exception:
            pass
        try:
            doc.getCurrentController().getFrame().getContainerWindow().toFront()
        except Exception:
            pass
    toolkit = _get_toolkit()
    try:
        if toolkit is not None and hasattr(toolkit, "processEventsToIdle"):
            toolkit.processEventsToIdle(50)
    except Exception:
        pass


def show_todo_task_progress_dialog(doc=None, title=u"Задачи", text=u"Выполняется…"):
    """
    Окно ожидания (тема проекта): setVisible без execute,
    чтобы можно было обновлять текст/прогресс во время работы.
    Родительское окно НЕ отключаем через Enable=False — иначе книга
    часто остаётся замороженной после закрытия.
    Возвращает объект с update(text, value=None, maximum=None) и close().
    """
    title = unicode(title or u"Задачи")
    text = unicode(text or u"Выполняется…")
    ctx = _uno_context()
    toolkit = _get_toolkit()
    if ctx is None or toolkit is None:
        _log("%s: %s" % (title, text))

        class _NullProg(object):
            def update(self, msg, value=None, maximum=None):
                return None

            def close(self):
                return None

        return _NullProg()

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    dlg_w = int(getattr(_cfg, "MSG_DIALOG_WIDTH", 420))
    content_w = dlg_w - 2 * m
    text_h = 36
    bar_h = 14

    dm.Title = title
    dm.Width = dlg_w
    dm.PositionX = 140
    dm.PositionY = 180
    dm.Closeable = False
    try:
        dm.Moveable = True
    except Exception:
        pass

    y = m
    msg = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    msg.Name = "ProgText"
    msg.Label = text
    msg.PositionX = m
    msg.PositionY = y
    msg.Width = content_w
    msg.Height = text_h
    try:
        msg.MultiLine = True
    except Exception:
        pass
    try:
        msg.VerticalAlign = 1
    except Exception:
        pass
    dm.insertByName("ProgText", msg)
    y += text_h + gap

    bar = dm.createInstance("com.sun.star.awt.UnoControlProgressBarModel")
    bar.Name = "ProgBar"
    bar.PositionX = m
    bar.PositionY = y
    bar.Width = content_w
    bar.Height = bar_h
    bar.ProgressValueMin = 0
    bar.ProgressValueMax = 100
    bar.ProgressValue = 0
    dm.insertByName("ProgBar", bar)
    y += bar_h + m
    dm.Height = max(int(getattr(_cfg, "MSG_DIALOG_MIN_H", 120)), y)

    try:
        _theme_todo_dialog(dm, title_text=title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass

        class _NullProg2(object):
            def update(self, msg, value=None, maximum=None):
                return None

            def close(self):
                return None

        return _NullProg2()

    try:
        try_paint_titlebar(dialog, title_bg=DARK_BLUE_TITLE_BG, title_fg=DARK_BLUE_TITLE_FG)
    except Exception:
        pass

    try:
        dialog.setVisible(True)
    except Exception:
        pass
    try:
        dialog.toFront()
    except Exception:
        pass
    try:
        if hasattr(toolkit, "processEventsToIdle"):
            toolkit.processEventsToIdle(50)
    except Exception:
        pass

    class _Prog(object):
        def __init__(self):
            self._dlg = dialog
            self._toolkit = toolkit
            self._doc = doc
            self._closed = False

        def update(self, msg, value=None, maximum=None):
            if self._closed or self._dlg is None:
                return
            try:
                if msg is not None:
                    mdl = self._dlg.getModel().getByName("ProgText")
                    mdl.Label = unicode(msg)
            except Exception:
                pass
            try:
                bmdl = self._dlg.getModel().getByName("ProgBar")
                if maximum is not None:
                    mx = int(maximum)
                    if mx < 1:
                        mx = 1
                    bmdl.ProgressValueMax = mx
                if value is not None:
                    bmdl.ProgressValue = int(value)
            except Exception:
                pass
            try:
                if hasattr(self._toolkit, "processEventsToIdle"):
                    self._toolkit.processEventsToIdle(30)
            except Exception:
                pass

        def close(self):
            if self._closed:
                return
            self._closed = True
            try:
                self._dlg.setVisible(False)
            except Exception:
                pass
            try:
                self._dlg.dispose()
            except Exception:
                pass
            self._dlg = None
            try:
                restore_todo_doc_ui(self._doc)
            except Exception:
                pass

    return _Prog()


def _get_combo_text(dialog, dm, name, items=None):
    """Текст ComboBox: выбранный пункт или ввод (как в param_wizard)."""
    ctl = None
    try:
        ctl = dialog.getControl(name)
    except Exception:
        pass
    if ctl is not None:
        try:
            pos = ctl.getSelectedItemPos()
            if pos is not None and int(pos) >= 0:
                if items is not None and int(pos) < len(items):
                    return unicode(items[int(pos)]).strip()
                try:
                    item = ctl.getSelectedItem()
                    if item:
                        return unicode(item).strip()
                except Exception:
                    pass
                try:
                    combo_items = ctl.getModel().StringItemList
                    if combo_items is not None and int(pos) < len(combo_items):
                        return unicode(combo_items[int(pos)]).strip()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            t = ctl.getText()
            if t is not None and unicode(t).strip():
                return unicode(t).strip()
        except Exception:
            pass
    return _get_control_text(dialog, dm, name).strip()


def _parse_int_field(text, min_val=None, max_val=None):
    s = unicode(text or u"").strip()
    if not s:
        return None
    m = re.match(r"^(\d+)", s)
    if not m:
        return None
    try:
        val = int(m.group(1))
    except ValueError:
        return None
    if min_val is not None and val < int(min_val):
        return None
    if max_val is not None and val > int(max_val):
        return None
    return val


def _month_index_from_combo(text, selected_pos=None):
    s = unicode(text or u"").strip()
    if s:
        m = re.match(r"^(\d{1,2})", s)
        if m:
            mo = int(m.group(1))
            if 1 <= mo <= 12:
                return mo
        for sep in (u"—", u"–", u"-", u" "):
            if sep in s:
                head = s.split(sep, 1)[0].strip()
                m = re.match(r"^(\d{1,2})$", head)
                if m:
                    mo = int(m.group(1))
                    if 1 <= mo <= 12:
                        return mo
    if selected_pos is not None:
        try:
            pos = int(selected_pos)
            if 0 <= pos <= 11:
                return pos + 1
        except Exception:
            pass
    return None


def _month_combo_text(month):
    items = getattr(_cfg, "MONTH_COMBO_ITEMS", ())
    idx = int(month) - 1
    if 0 <= idx < len(items):
        return unicode(items[idx])
    return u"%02d" % int(month)


def _read_due_parts_from_dialog(dialog, dm):
    month_items = getattr(_cfg, "MONTH_COMBO_ITEMS", ())
    selected_pos = None
    try:
        selected_pos = dialog.getControl("DueMonthCombo").getSelectedItemPos()
    except Exception:
        pass
    day = _parse_int_field(_get_control_text(dialog, dm, "DueDayEdit"), 1, 31)
    year = _parse_int_field(_get_control_text(dialog, dm, "DueYearEdit"), 1900, 2100)
    month = _month_index_from_combo(
        _get_combo_text(dialog, dm, "DueMonthCombo", month_items),
        selected_pos=selected_pos,
    )
    if day is None and year is None and month is None:
        return None
    if day is None or month is None or year is None:
        return u"incomplete"
    try:
        from datetime import date

        date(int(year), int(month), int(day))
    except ValueError:
        return u"invalid"
    return int(day), int(month), int(year)


def _initial_due_parts(fields, use_default_due):
    from todo_task_lib import default_due_parts, parse_due_parts

    parts = fields.get(u"due_parts")
    if parts and len(parts) == 3:
        return tuple(int(x) for x in parts)
    due = unicode(fields.get(u"due") or u"").strip()
    parts = parse_due_parts(due) if due else None
    if parts:
        return parts
    if use_default_due:
        return default_due_parts()
    return None


def _plan_edit_button_label(doc, plan_guid, plan_read_only=False):
    plan_guid = unicode(plan_guid or u"").strip()
    if plan_read_only:
        try:
            from todo_task_plan_lib import count_plan_items

            if plan_guid and count_plan_items(doc, plan_guid) > 0:
                return unicode(
                    getattr(_pcfg, "PLAN_BTN_VIEW", u"Просмотр плана")
                )
        except Exception as exc:
            _log("plan btn label: %s" % exc)
        return unicode(getattr(_pcfg, "PLAN_BTN_NO_PLAN", u"План отсутствует"))
    lbl = unicode(
        getattr(_pcfg, "PLAN_BTN_CREATE", u"Создать план")
        if _pcfg is not None
        else u"Создать план"
    )
    if plan_guid:
        try:
            from todo_task_plan_lib import count_plan_items

            if count_plan_items(doc, plan_guid) > 0:
                lbl = unicode(
                    getattr(_pcfg, "PLAN_BTN_EDIT", u"Редактировать план")
                    if _pcfg is not None
                    else u"Редактировать план"
                )
        except Exception as exc:
            _log("plan btn label: %s" % exc)
    return lbl


def _refresh_plan_buttons(dialog, doc, plan_guid, plan_read_only=False):
    try:
        dialog.getControl("PlanBtn").setLabel(
            _plan_edit_button_label(doc, plan_guid, plan_read_only=plan_read_only)
        )
    except Exception:
        pass
    try:
        has_plan = False
        if unicode(plan_guid or u"").strip():
            from todo_task_plan_lib import count_plan_items

            has_plan = count_plan_items(doc, plan_guid) > 0
        dialog.getControl("PlanBtn").setEnable(
            (not plan_read_only) or has_plan
        )
    except Exception:
        pass
    try:
        dialog.getControl("PlanDelBtn").setEnable(not plan_read_only)
    except Exception:
        pass


def _set_edit_control_enabled(dialog, name, enabled):
    try:
        dialog.getControl(name).setEnable(bool(enabled))
    except Exception:
        pass


def _apply_edit_dialog_access(dialog, dm, access):
    access = access or {}
    read_only = bool(access.get(u"read_only"))
    plan_read_only = bool(access.get(u"plan_read_only"))
    can_unlock = bool(access.get(u"can_unlock_edit"))
    is_merge = bool(access.get(u"is_merge_source"))
    foreign = bool(access.get(u"foreign_creator"))
    for nm in (
        u"NameEdit",
        u"DueDayEdit",
        u"DueMonthCombo",
        u"DueYearEdit",
        u"PriCombo",
        u"SedEdit",
        u"AsgCombo",
        u"StCombo",
        u"CmEdit",
    ):
        _set_edit_control_enabled(dialog, nm, not read_only)
    try:
        dialog.getControl(u"OkBtn").setEnable(not read_only)
    except Exception:
        pass
    hint = u""
    if read_only and is_merge:
        hint = unicode(
            getattr(_cfg, "EDIT_READ_ONLY_MERGE_HINT", u"")
        )
    elif read_only and foreign:
        hint = unicode(
            getattr(_cfg, "EDIT_READ_ONLY_FOREIGN_HINT", u"")
        )
    elif (not read_only) and plan_read_only and foreign:
        hint = unicode(
            getattr(_cfg, "EDIT_FOREIGN_PLAN_HINT", u"")
        )
    try:
        dialog.getControl(u"ReadOnlyLbl").setLabel(hint)
    except Exception:
        pass
    try:
        ub = dialog.getControl(u"UnlockEditBtn")
        ub.setEnable(can_unlock and read_only)
        try:
            ub.setVisible(bool(can_unlock and read_only))
        except Exception:
            pass
    except Exception:
        pass
    plan_guid = unicode(access.get(u"plan_guid") or u"").strip()
    _refresh_plan_buttons(
        dialog,
        access.get(u"doc"),
        plan_guid,
        plan_read_only=plan_read_only,
    )


def _snapshot_edit_dialog_fields(dialog, dm, fields, plan_guid):
    """Снимок полей формы задачи перед отложенным действием (план и т.п.)."""
    due_raw = _read_due_parts_from_dialog(dialog, dm)
    out = {
        u"name": _get_control_text(dialog, dm, "NameEdit").strip(),
        u"priority": _get_control_text(dialog, dm, "PriCombo").strip(),
        u"sed": _get_control_text(dialog, dm, "SedEdit").strip(),
        u"assignee": _get_control_text(dialog, dm, "AsgCombo").strip(),
        u"status": _get_control_text(dialog, dm, "StCombo").strip(),
        u"comment": _get_control_text(dialog, dm, "CmEdit"),
        u"guid": unicode(plan_guid or fields.get(u"guid") or u"").strip(),
    }
    if due_raw in (u"incomplete", u"invalid"):
        if fields.get(u"due_parts"):
            out[u"due_parts"] = fields.get(u"due_parts")
            out[u"due"] = fields.get(u"due") or u""
        else:
            out[u"due"] = u""
    elif due_raw is None:
        out[u"due"] = u""
    else:
        out[u"due_parts"] = due_raw
        out[u"due"] = u"%02d.%02d.%04d" % due_raw
    for k in (u"owner", u"editor", u"created", u"modified"):
        if fields.get(k) is not None:
            out[k] = fields.get(k)
    return out


def _plan_sheet_name(plan_ctx):
    plan_ctx = plan_ctx or {}
    src = unicode(plan_ctx.get(u"sheet_name") or u"").strip()
    if src:
        return src
    try:
        sh = plan_ctx.get(u"sheet")
        if sh is not None:
            return unicode(sh.Name)
    except Exception:
        pass
    return u""


def _merge_draft_from_task_row(draft, sheet, cols, row, sync=None):
    """Подтянуть с листа срок и комментарий после синхронизации с планом."""
    draft = dict(draft or {})
    sync = sync or {}
    if sync.get(u"due_parts"):
        draft[u"due_parts"] = sync.get(u"due_parts")
        draft[u"due"] = sync.get(u"due") or draft.get(u"due") or u""
    if sync.get(u"comment") is not None:
        draft[u"comment"] = sync.get(u"comment")
    if sheet is None or cols is None or row is None:
        return draft
    try:
        from todo_task_lib import read_task_fields

        fresh = read_task_fields(sheet, cols, int(row))
        if fresh.get(u"due_parts"):
            draft[u"due_parts"] = fresh.get(u"due_parts")
            draft[u"due"] = fresh.get(u"due") or u""
        if u"comment" in cols and fresh.get(u"comment") is not None:
            draft[u"comment"] = fresh.get(u"comment")
    except Exception:
        pass
    return draft


def _handle_edit_dialog_defer(defer, draft, doc, plan_ctx, assignees, priorities, statuses):
    """Действия с планом вне модального цикла формы задачи (без вложенного execute)."""
    plan_ctx = plan_ctx or {}
    plan_guid = unicode(
        plan_ctx.get(u"guid") or (draft or {}).get(u"guid") or u""
    ).strip()
    task_nm = unicode((draft or {}).get(u"name") or u"").strip()
    src_sheet = _plan_sheet_name(plan_ctx)

    if defer == u"plan":
        from todo_task_plan_dialog_lib import show_todo_task_plan_wizard

        _plan_trace(u"defer: открытие визарда плана guid=%s" % plan_guid)
        saved, plan_sync = show_todo_task_plan_wizard(
            doc=doc,
            guid=plan_guid,
            task_name=task_nm,
            sheet=plan_ctx.get(u"sheet"),
            cols=plan_ctx.get(u"cols"),
            row=plan_ctx.get(u"row"),
            assignees=assignees,
            parent_dialog=None,
            read_only=bool(plan_ctx.get(u"plan_read_only")),
        )
        _plan_trace(u"defer: визард плана закрыт saved=%s" % bool(saved))
        return _merge_draft_from_task_row(
            draft,
            plan_ctx.get(u"sheet"),
            plan_ctx.get(u"cols"),
            plan_ctx.get(u"row"),
            sync=plan_sync if saved else None,
        )
    if defer == u"del_plan":
        if plan_ctx.get(u"plan_read_only"):
            return draft
        from todo_task_plan_lib import remove_plan_with_archive

        _plan_trace(u"defer: удаление плана guid=%s" % plan_guid)
        ok, msg = remove_plan_with_archive(
            doc,
            plan_guid,
            task_name=task_nm,
            source_sheet_name=src_sheet,
        )
        if msg:
            title = unicode(getattr(_pcfg, "PLAN_BTN_DELETE", u"Удалить план"))
            if ok:
                _dialog_warn(doc, msg, title=title)
            else:
                _dialog_warn(doc, msg)
        return draft
    if defer == u"export_plan":
        from todo_task_plan_lib import export_plan_to_xlsx

        _plan_trace(u"defer: выгрузка плана guid=%s" % plan_guid)
        ok, msg = export_plan_to_xlsx(
            doc,
            plan_guid,
            task_name=task_nm,
            open_after=True,
        )
        if not ok and msg:
            _dialog_warn(
                doc,
                msg,
                title=unicode(getattr(_pcfg, "PLAN_BTN_EXPORT", u"Выгрузить план")),
            )
        elif ok and msg:
            _dialog_warn(
                doc,
                unicode(
                    getattr(_pcfg, "PLAN_EXPORT_DONE", u"План выгружен в файл:\n%s")
                )
                % msg,
                title=unicode(getattr(_pcfg, "PLAN_BTN_EXPORT", u"Выгрузить план")),
            )
        return draft
    return draft


def show_todo_task_edit_dialog( doc=None, fields=None, priorities=None, assignees=None, statuses=None, use_default_due=False, plan_ctx=None, edit_access=None):
    """
    Модальный диалог редактирования полей задачи.
    use_default_due=True — подставить срок (сегодня+4, без выходных).
    Возвращает dict полей или None при отмене.
    """
    from todo_task_lib import validate_due_not_before_today

    fields = fields or {}
    priorities = list(priorities or [])
    assignees = list(assignees or [])
    statuses = list(statuses or [])
    plan_ctx = plan_ctx or {}
    if edit_access is None:
        try:
            from todo_task_lib import build_task_edit_access

            edit_access = build_task_edit_access(
                plan_ctx.get(u"sheet"),
                fields,
                doc=doc,
                use_default_due=use_default_due,
            )
        except Exception:
            edit_access = {
                u"read_only": False,
                u"plan_read_only": False,
                u"can_unlock_edit": False,
            }

    while True:
        outcome = _show_todo_task_edit_dialog_once(
            doc=doc,
            fields=fields,
            priorities=priorities,
            assignees=assignees,
            statuses=statuses,
            use_default_due=use_default_due,
            plan_ctx=plan_ctx,
            edit_access=edit_access,
        )
        defer = outcome.get(u"defer")
        if defer:
            fields = _handle_edit_dialog_defer(
                defer,
                outcome.get(u"draft") or fields,
                doc,
                plan_ctx,
                assignees,
                priorities,
                statuses,
            )
            use_default_due = False
            continue
        if outcome.get(u"confirmed"):
            return outcome.get(u"result")
        return None


def _show_todo_task_edit_dialog_once( doc=None, fields=None, priorities=None, assignees=None, statuses=None, use_default_due=False, plan_ctx=None, edit_access=None):
    """Один показ формы редактирования задачи. defer → повтор снаружи."""
    from todo_task_lib import validate_due_not_before_today

    fields = fields or {}
    priorities = list(priorities or [])
    assignees = list(assignees or [])
    statuses = list(statuses or [])
    plan_ctx = plan_ctx or {}
    edit_access = dict(edit_access or {})
    if not edit_access:
        edit_access = {
            u"read_only": bool(plan_ctx.get(u"read_only")),
            u"plan_read_only": bool(plan_ctx.get(u"plan_read_only")),
            u"can_unlock_edit": bool(plan_ctx.get(u"can_unlock_edit")),
            u"is_merge_source": bool(plan_ctx.get(u"is_merge_source")),
            u"foreign_creator": bool(plan_ctx.get(u"foreign_creator")),
        }

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")

    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)

    dialog_title = (
        _cfg.NEW_TASK_DIALOG_TITLE if use_default_due else _cfg.EDIT_DIALOG_TITLE
    )
    if edit_access.get(u"read_only"):
        dialog_title = unicode(dialog_title) + unicode(
            getattr(_cfg, "EDIT_READ_ONLY_TITLE_SUFFIX", u" (просмотр)")
        )

    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    lbl_h = int(_cfg.EDIT_LABEL_H)
    fld_h = int(_cfg.EDIT_FIELD_H)
    ml_h = int(_cfg.EDIT_MULTILINE_H)
    btn_w = int(_cfg.EDIT_BTN_W)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = int(_cfg.EDIT_DIALOG_WIDTH)
    content_w = dlg_w - 2 * m
    day_w = int(getattr(_cfg, "DUE_DAY_W", 36))
    month_w = int(getattr(_cfg, "DUE_MONTH_W", 130))
    year_w = int(getattr(_cfg, "DUE_YEAR_W", 52))
    month_items = getattr(_cfg, "MONTH_COMBO_ITEMS", ())

    dm.Title = unicode(dialog_title)
    dm.Width = dlg_w
    dm.PositionX = 80
    dm.PositionY = 60

    y = m

    hint_h = 36
    unlock_btn_w = 150
    can_unlock = bool(edit_access.get(u"can_unlock_edit"))
    show_read_hint = bool(edit_access.get(u"read_only")) or (
        bool(edit_access.get(u"foreign_creator"))
        and bool(edit_access.get(u"plan_read_only"))
    )
    if show_read_hint or can_unlock:
        top_row_h = max(hint_h, btn_h)
        if can_unlock:
            _add_button(
                dm,
                "UnlockEditBtn",
                unicode(getattr(_cfg, "EDIT_UNLOCK_BTN", u"Редактировать")),
                m + content_w - unlock_btn_w,
                y,
                unlock_btn_w,
                btn_h,
            )
        if show_read_hint:
            hint_w = content_w - (unlock_btn_w + gap if can_unlock else 0)
            _add_label(dm, "ReadOnlyLbl", u"", m, y, hint_w, top_row_h)
            try:
                dm.getByName(u"ReadOnlyLbl").MultiLine = True
            except Exception:
                pass
        y += top_row_h + 4

    _add_label(dm, "NameLbl", u"Наименование:", m, y, content_w, lbl_h)
    y += lbl_h + 2
    _add_edit(dm, "NameEdit", m, y, content_w, ml_h, multiline=True)
    y += ml_h + gap

    # Срок (слева) и матрица приоритетов (справа) — один блок, пополам.
    col_gap = gap
    half_w = int((content_w - col_gap) // 2)
    if half_w < 1:
        half_w = content_w
    left_x = m
    right_x = m + half_w + col_gap
    due_fields_w = day_w + 4 + month_w + 4 + year_w
    if due_fields_w > half_w:
        # ужать месяц, чтобы тройка дата уместилась в левую половину
        month_w = max(80, half_w - (day_w + 4 + year_w + 4))
        due_fields_w = day_w + 4 + month_w + 4 + year_w

    _add_label(dm, "DueLbl", u"Срок исполнения:", left_x, y, half_w, lbl_h)
    _add_label(dm, "PriLbl", u"Матрица приоритетов:", right_x, y, half_w, lbl_h)
    y += lbl_h + 2

    x = left_x
    _add_label(dm, "DueDayLbl", u"День", x, y, day_w, lbl_h)
    x += day_w + 4
    _add_label(dm, "DueMonthLbl", u"Месяц", x, y, month_w, lbl_h)
    x += month_w + 4
    _add_label(dm, "DueYearLbl", u"Год", x, y, year_w, lbl_h)
    y += lbl_h + 2

    x = left_x
    _add_edit(dm, "DueDayEdit", x, y, day_w, fld_h, multiline=False)
    x += day_w + 4
    _add_combo(dm, "DueMonthCombo", x, y, month_w, fld_h, month_items)
    x += month_w + 4
    _add_edit(dm, "DueYearEdit", x, y, year_w, fld_h, multiline=False)
    _add_combo(dm, "PriCombo", right_x, y, half_w, fld_h, priorities)
    y += fld_h + gap

    _add_label(dm, "SedLbl", u"Реквизиты документа СЭД (при наличии):", m, y, content_w, lbl_h)
    y += lbl_h + 2
    _add_edit(dm, "SedEdit", m, y, content_w, fld_h, multiline=False)
    y += fld_h + gap

    _add_label(dm, "AsgLbl", u"Исполнитель(и):", m, y, content_w, lbl_h)
    y += lbl_h + 2
    _add_combo(dm, "AsgCombo", m, y, content_w, fld_h, assignees)
    y += fld_h + gap

    _add_label(dm, "StLbl", u"Статус исполнения задачи:", m, y, content_w, lbl_h)
    y += lbl_h + 2
    _add_combo(dm, "StCombo", m, y, content_w, fld_h, statuses)
    y += fld_h + gap

    _add_label(dm, "CmLbl", u"Комментарии:", m, y, content_w, lbl_h)
    y += lbl_h + 2
    _add_edit(dm, "CmEdit", m, y, content_w, ml_h, multiline=True)
    y += ml_h + gap

    plan_ctx = plan_ctx or {}
    plan_guid = unicode(
        plan_ctx.get(u"guid") or fields.get(u"guid") or u""
    ).strip()
    plan_read_only = bool(edit_access.get(u"plan_read_only"))
    plan_btn_label = _plan_edit_button_label(
        doc, plan_guid, plan_read_only=plan_read_only
    )
    plan_gap = 4
    w_del = 108
    w_export = 118
    w_plan = max(120, content_w - w_del - w_export - 2 * plan_gap)
    x_plan = m
    x_del = x_plan + w_plan + plan_gap
    x_export = x_del + w_del + plan_gap
    _add_button(
        dm,
        "PlanBtn",
        plan_btn_label,
        x_plan,
        y,
        w_plan,
        btn_h,
    )
    _add_button(
        dm,
        "PlanDelBtn",
        unicode(getattr(_pcfg, "PLAN_BTN_DELETE", u"Удалить план")),
        x_del,
        y,
        w_del,
        btn_h,
    )
    _add_button(
        dm,
        "PlanExportBtn",
        unicode(getattr(_pcfg, "PLAN_BTN_EXPORT", u"Выгрузить план")),
        x_export,
        y,
        w_export,
        btn_h,
    )
    y += btn_h + gap + 4

    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = "OkBtn"
    ok_btn.Label = unicode(_cfg.BTN_OK_LABEL)
    ok_btn.PositionX = dlg_w - m - 2 * btn_w - gap
    ok_btn.PositionY = y
    ok_btn.Width = btn_w
    ok_btn.Height = btn_h
    try:
        ok_btn.DefaultButton = True
    except Exception:
        pass
    dm.insertByName("OkBtn", ok_btn)

    cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = "CancelBtn"
    cancel_btn.Label = unicode(_cfg.BTN_CANCEL_LABEL)
    cancel_btn.PositionX = dlg_w - m - btn_w
    cancel_btn.PositionY = y
    cancel_btn.Width = btn_w
    cancel_btn.Height = btn_h
    dm.insertByName("CancelBtn", cancel_btn)

    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=unicode(dialog_title), doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        raise RuntimeError(u"Не удалось создать peer диалога.")
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass

    initials = {
        u"NameEdit": unicode(fields.get(u"name") or u""),
        u"PriCombo": unicode(fields.get(u"priority") or u""),
        u"SedEdit": unicode(fields.get(u"sed") or u""),
        u"AsgCombo": unicode(fields.get(u"assignee") or u""),
        u"StCombo": unicode(fields.get(u"status") or u""),
        u"CmEdit": unicode(fields.get(u"comment") or u""),
    }
    for name, text in initials.items():
        _set_control_text(dialog, dm, name, text)

    due_parts = _initial_due_parts(fields, use_default_due)
    if due_parts:
        _set_control_text(dialog, dm, "DueDayEdit", u"%02d" % due_parts[0])
        month_text = _month_combo_text(due_parts[1])
        _set_control_text(dialog, dm, "DueMonthCombo", month_text)
        try:
            dialog.getControl("DueMonthCombo").selectItemPos(int(due_parts[1]) - 1, True)
        except Exception:
            pass
        _set_control_text(dialog, dm, "DueYearEdit", unicode(due_parts[2]))

    state = {u"confirmed": False, u"result": None, u"defer": None, u"draft": None}
    access_state = dict(edit_access)
    access_state[u"doc"] = doc
    access_state[u"plan_guid"] = plan_guid

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "UnlockEditBtn":
                if access_state.get(u"foreign_creator"):
                    creator = unicode(fields.get(u"created_by") or u"").strip()
                    if not creator:
                        creator = u"—"
                    warn = unicode(getattr(_cfg, "EDIT_FOREIGN_CREATOR_WARN", u"")) % (
                        creator,
                    )
                    title = unicode(
                        getattr(_cfg, "EDIT_UNLOCK_CONFIRM_TITLE", u"Чужая задача")
                    )
                    if not show_todo_task_confirm(doc=doc, text=warn, title=title):
                        return
                access_state[u"read_only"] = False
                _apply_edit_dialog_access(dialog, dm, access_state)
                return
            if name == "OkBtn":
                if access_state.get(u"read_only"):
                    return
                due_raw = _read_due_parts_from_dialog(dialog, dm)
                if due_raw == u"incomplete":
                    _dialog_warn(
                        doc,
                        u"Заполните день, месяц и год срока исполнения\n"
                        u"или очистите все три поля.",
                    )
                    return
                if due_raw == u"invalid":
                    _dialog_warn(doc, u"Некорректная дата.")
                    return
                if due_raw is not None:
                    ok, err = validate_due_not_before_today(
                        due_raw[0], due_raw[1], due_raw[2]
                    )
                    if not ok:
                        _dialog_warn(doc, err)
                        return
                result = {
                    u"name": _get_control_text(dialog, dm, "NameEdit").strip(),
                    u"priority": _get_control_text(dialog, dm, "PriCombo").strip(),
                    u"sed": _get_control_text(dialog, dm, "SedEdit").strip(),
                    u"assignee": _get_control_text(dialog, dm, "AsgCombo").strip(),
                    u"status": _get_control_text(dialog, dm, "StCombo").strip(),
                    u"comment": _get_control_text(dialog, dm, "CmEdit"),
                }
                if due_raw is None:
                    result[u"due"] = u""
                else:
                    result[u"due_parts"] = due_raw
                    result[u"due"] = u"%02d.%02d.%04d" % due_raw
                state[u"result"] = result
                state[u"confirmed"] = True
            elif name == "PlanBtn":
                if not plan_guid:
                    _dialog_warn(doc, u"Нет GUID задачи — план недоступен.")
                    return
                if plan_read_only:
                    try:
                        from todo_task_plan_lib import count_plan_items

                        if count_plan_items(doc, plan_guid) <= 0:
                            _dialog_warn(doc, u"У задачи нет сохранённого плана.")
                            return
                    except Exception:
                        pass
                state[u"defer"] = u"plan"
                state[u"draft"] = _snapshot_edit_dialog_fields(
                    dialog, dm, fields, plan_guid
                )
                _log(u"PlanBtn: defer open plan wizard guid=%s" % plan_guid)
                try:
                    dialog.endExecute()
                except Exception:
                    pass
                return
            elif name == "PlanDelBtn":
                if not plan_guid:
                    _dialog_warn(doc, u"Нет GUID задачи — план недоступен.")
                    return
                if plan_read_only or access_state.get(u"plan_read_only"):
                    return
                state[u"defer"] = u"del_plan"
                state[u"draft"] = _snapshot_edit_dialog_fields(
                    dialog, dm, fields, plan_guid
                )
                try:
                    dialog.endExecute()
                except Exception:
                    pass
                return
            elif name == "PlanExportBtn":
                if not plan_guid:
                    _dialog_warn(doc, u"Нет GUID задачи — план недоступен.")
                    return
                state[u"defer"] = u"export_plan"
                state[u"draft"] = _snapshot_edit_dialog_fields(
                    dialog, dm, fields, plan_guid
                )
                try:
                    dialog.endExecute()
                except Exception:
                    pass
                return
            elif name == "CancelBtn":
                state[u"confirmed"] = False
                state[u"result"] = None
            else:
                return
            try:
                dialog.endExecute()
            except Exception:
                pass

    handler = _Handler()
    for btn_name in (
        u"OkBtn",
        u"CancelBtn",
        u"PlanBtn",
        u"PlanDelBtn",
        u"PlanExportBtn",
        u"UnlockEditBtn",
    ):
        try:
            dialog.getControl(btn_name).addActionListener(handler)
        except Exception as err:
            _log("addActionListener %s: %s" % (btn_name, err))

    _apply_edit_dialog_access(dialog, dm, access_state)

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

    return {
        u"confirmed": bool(state.get(u"confirmed")),
        u"result": state.get(u"result"),
        u"defer": state.get(u"defer"),
        u"draft": state.get(u"draft"),
    }


def _choice_label(choices, code):
    """Подпись пункта combo по коду; иначе первый пункт."""
    wanted = unicode(code or u"")
    i = 0
    while i < len(choices):
        item_code, item_label = choices[i]
        if unicode(item_code) == wanted:
            return unicode(item_label)
        i = i + 1
    if choices:
        return unicode(choices[0][1])
    return u""


def show_todo_task_sort_dialog(doc=None, initial=None):
    """
    Диалог выбора способа сортировки задач.
    initial — dict {mode, due_direction} из todo_task_settings.
    Возвращает dict или None при отмене.
    """
    try:
        from todo_task_settings_lib import normalize_sort_settings
    except Exception:
        def normalize_sort_settings(obj):
            return {
                u"mode": getattr(_cfg, "SORT_MODE_ASSIGNEE", u"assignee"),
                u"due_direction": getattr(
                    _cfg, "SORT_DUE_DIR_OLD_FIRST", u"old_first"
                ),
            }

    init = normalize_sort_settings(initial)
    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_w = int(_cfg.EDIT_BTN_W)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = int(_cfg.EDIT_DIALOG_WIDTH)
    content_w = dlg_w - 2 * m
    field_h = int(_cfg.EDIT_FIELD_H)
    lbl_h = int(_cfg.EDIT_LABEL_H)

    dm.Title = unicode(getattr(_cfg, "SORT_DIALOG_TITLE", u"Сортировка задач"))
    dm.Width = dlg_w
    dm.PositionX = 140
    dm.PositionY = 120

    y = m
    hint = unicode(
        getattr(
            _cfg,
            "SORT_DIALOG_HINT",
            u"Сортировка целых строк-задач. После сортировки «№ п/п» — 1…N.",
        )
    )
    _add_label(dm, "HintLbl", hint, m, y, content_w, 32)
    y += 36

    mode_choices = getattr(_cfg, "SORT_MODE_CHOICES", ())
    mode_labels = [lbl for _code, lbl in mode_choices]
    _add_label(dm, "ModeLbl", u"Сортировать по:", m, y, content_w, lbl_h)
    y += lbl_h + 2
    _add_combo(dm, "ModeCombo", m, y, content_w, field_h, mode_labels)
    y += field_h + gap

    due_choices = getattr(_cfg, "SORT_DUE_DIR_CHOICES", ())
    due_labels = [lbl for _code, lbl in due_choices]
    _add_label(dm, "DueDirLbl", u"Порядок сроков (для сортировки по дате):", m, y, content_w, lbl_h)
    y += lbl_h + 2
    _add_combo(dm, "DueDirCombo", m, y, content_w, field_h, due_labels)
    y += field_h + gap + 4

    ok_x = dlg_w - m - 2 * btn_w - gap
    cancel_x = dlg_w - m - btn_w
    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = "OkBtn"
    ok_btn.Label = unicode(_cfg.BTN_OK_LABEL)
    ok_btn.PositionX = ok_x
    ok_btn.PositionY = y
    ok_btn.Width = btn_w
    ok_btn.Height = btn_h
    try:
        ok_btn.DefaultButton = True
    except Exception:
        pass
    dm.insertByName("OkBtn", ok_btn)

    cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = "CancelBtn"
    cancel_btn.Label = unicode(_cfg.BTN_CANCEL_LABEL)
    cancel_btn.PositionX = cancel_x
    cancel_btn.PositionY = y
    cancel_btn.Width = btn_w
    cancel_btn.Height = btn_h
    dm.insertByName("CancelBtn", cancel_btn)

    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=dm.Title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        return None
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass

    mode_init = _choice_label(mode_choices, init.get(u"mode"))
    due_init = _choice_label(due_choices, init.get(u"due_direction"))
    try:
        mode_ctl = dialog.getControl("ModeCombo")
        if mode_ctl is not None and mode_init:
            mode_ctl.setText(mode_init)
    except Exception:
        pass
    try:
        due_ctl = dialog.getControl("DueDirCombo")
        if due_ctl is not None and due_init:
            due_ctl.setText(due_init)
    except Exception:
        pass

    def _combo_code(combo_name, choices, default_code):
        label = _get_combo_text(dialog, dm, combo_name)
        ci = 0
        while ci < len(choices):
            code, item_label = choices[ci]
            if label == item_label or label == code:
                return code
            ci = ci + 1
        return default_code

    def _read_spec():
        return normalize_sort_settings(
            {
                u"mode": _combo_code(
                    u"ModeCombo",
                    mode_choices,
                    getattr(_cfg, "SORT_MODE_ASSIGNEE", u"assignee"),
                ),
                u"due_direction": _combo_code(
                    u"DueDirCombo",
                    due_choices,
                    getattr(_cfg, "SORT_DUE_DIR_OLD_FIRST", u"old_first"),
                ),
            }
        )

    state = {u"ok": False, u"result": None}

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "OkBtn":
                state[u"result"] = _read_spec()
                state[u"ok"] = True
            elif name == "CancelBtn":
                state[u"ok"] = False
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
            _log("sort listener %s: %s" % (btn_name, err))

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

    if not state[u"ok"]:
        return None
    return state.get(u"result")


def show_todo_task_colorize_dialog(doc=None, initial=None):
    """
    Диалог параметров раскраски задач.
    initial — dict {all_sheets, sheet_pattern} из todo_task_settings.
    Возвращает dict или None при отмене.
    """
    try:
        from todo_task_settings_lib import normalize_colorize_settings
    except Exception:
        def normalize_colorize_settings(obj):
            return {
                u"all_sheets": False,
                u"sheet_pattern": getattr(
                    _cfg, "COLORIZE_DEFAULT_SHEET_PATTERN", u"Задачи*"
                ),
            }

    init = normalize_colorize_settings(initial)

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_w = int(_cfg.EDIT_BTN_W)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = int(_cfg.EDIT_DIALOG_WIDTH)
    content_w = dlg_w - 2 * m
    field_h = int(_cfg.EDIT_FIELD_H)
    lbl_h = int(_cfg.EDIT_LABEL_H)

    dm.Title = unicode(getattr(_cfg, "COLORIZE_DIALOG_TITLE", u"Раскраска задач"))
    dm.Width = dlg_w
    dm.PositionX = 140
    dm.PositionY = 120

    y = m
    hint = unicode(
        getattr(_cfg, "COLORIZE_DIALOG_HINT", u"")
    )
    _add_label(dm, "HintLbl", hint, m, y, content_w, 36)
    y += 40

    all_lbl = unicode(
        getattr(_cfg, "COLORIZE_ALL_SHEETS_LABEL", u"Применить ко всем листам по шаблону")
    )
    _add_checkbox(dm, "AllSheetsChk", all_lbl, m, y, content_w, 16)
    y += 20 + gap

    pat_lbl = unicode(
        getattr(_cfg, "COLORIZE_PATTERN_LABEL", u"Шаблон имени листа:")
    )
    _add_label(dm, "PatternLbl", pat_lbl, m, y, content_w, lbl_h)
    y += lbl_h + 2
    _add_edit(dm, "PatternEd", m, y, content_w, field_h)
    y += field_h + gap + 4

    ok_x = dlg_w - m - 2 * btn_w - gap
    cancel_x = dlg_w - m - btn_w
    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = "OkBtn"
    ok_btn.Label = unicode(_cfg.BTN_OK_LABEL)
    ok_btn.PositionX = ok_x
    ok_btn.PositionY = y
    ok_btn.Width = btn_w
    ok_btn.Height = btn_h
    try:
        ok_btn.DefaultButton = True
    except Exception:
        pass
    dm.insertByName("OkBtn", ok_btn)

    cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = "CancelBtn"
    cancel_btn.Label = unicode(_cfg.BTN_CANCEL_LABEL)
    cancel_btn.PositionX = cancel_x
    cancel_btn.PositionY = y
    cancel_btn.Width = btn_w
    cancel_btn.Height = btn_h
    dm.insertByName("CancelBtn", cancel_btn)

    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=dm.Title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        return None
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass

    all_on = bool(init.get(u"all_sheets"))
    pattern = unicode(
        init.get(u"sheet_pattern")
        or getattr(_cfg, "COLORIZE_DEFAULT_SHEET_PATTERN", u"Задачи*")
    )
    try:
        dialog.getControl("AllSheetsChk").setState(1 if all_on else 0)
    except Exception:
        pass
    _set_control_text(dialog, dm, "PatternEd", pattern)

    def _sync_pattern_enabled():
        try:
            chk = dialog.getControl("AllSheetsChk")
            ed = dialog.getControl("PatternEd")
            lbl = dialog.getControl("PatternLbl")
            on = bool(chk.getState())
        except Exception:
            return
        try:
            ed.setEnable(on)
        except Exception:
            try:
                ed.Model.Enabled = on
            except Exception:
                pass
        try:
            lbl.Model.Enabled = on
        except Exception:
            pass

    _sync_pattern_enabled()

    state = {u"ok": False, u"result": None}

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "OkBtn":
                all_sheets = False
                try:
                    all_sheets = bool(dialog.getControl("AllSheetsChk").getState())
                except Exception:
                    pass
                pat = _get_control_text(dialog, dm, "PatternEd").strip()
                if not pat:
                    pat = unicode(
                        getattr(_cfg, "COLORIZE_DEFAULT_SHEET_PATTERN", u"Задачи*")
                    )
                state[u"result"] = normalize_colorize_settings(
                    {u"all_sheets": all_sheets, u"sheet_pattern": pat}
                )
                state[u"ok"] = True
            elif name == "CancelBtn":
                state[u"ok"] = False
            elif name == "AllSheetsChk":
                _sync_pattern_enabled()
                return
            else:
                return
            try:
                dialog.endExecute()
            except Exception:
                pass

    handler = _Handler()
    for ctl_name in (u"OkBtn", u"CancelBtn", u"AllSheetsChk"):
        try:
            dialog.getControl(ctl_name).addActionListener(handler)
        except Exception as err:
            _log("colorize listener %s: %s" % (ctl_name, err))

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

    if not state[u"ok"]:
        return None
    return state.get(u"result")



def _settings_book_path_label(book_path, has_private=False, max_len=78):
    path = unicode(book_path or u"").strip()
    prefix = unicode(getattr(_cfg, "SETTINGS_SCOPE_BOOK_LABEL", u"Книга:"))
    if not path:
        return unicode(
            getattr(
                _cfg,
                "SETTINGS_SCOPE_UNSAVED_BOOK",
                u"Книга не сохранена на диск — частные настройки появятся после сохранения файла.",
            )
        )
    shown = path
    extra = u""
    if not has_private:
        extra = u"  (частных ещё нет — показаны глобальные)"
        max_len = int(max_len) - 8
    if len(shown) > int(max_len):
        shown = u"…" + shown[-(int(max_len) - 1) :]
    return u"%s %s%s" % (prefix, shown, extra)


def show_todo_task_settings_dialog( doc=None, initial_general=None, initial_layout=None, header_names=None, initial=None, initial_merge=None, initial_scope=None, global_general=None, global_layout=None, global_merge=None, book_path=None, has_private=False):
    """
    Настройки todo_task со вкладками «Общие» / «Размеры полей».
    Возвращает {general, layout, merge, apply_widths, scope} или None.
    initial — устаревший алиас initial_general.
    initial_scope — private (по умолчанию) или global.
    """
    try:
        from todo_task_settings_lib import (
            normalize_general_settings,
            normalize_layout_settings,
            normalize_merge_settings,
        )
    except Exception:
        def normalize_general_settings(obj):
            return {
                u"protect_password": getattr(
                    _cfg, "TASK_SHEET_PROTECT_PASSWORD", u"task"
                ),
                u"owner_name": u"",
                u"owner_role": getattr(_cfg, "OWNER_ROLE_MANAGER", u"manager"),
                u"console_log": True,
                u"dialog_font_pt": int(
                    getattr(_cfg, "DEFAULT_DIALOG_FONT_PT", 11)
                ),
            }

        def normalize_layout_settings(obj, header_names=None):
            return {
                u"template_sheet": u"Задачи.01",
                u"task_sheet_pattern": u"Задачи*",
                u"summary_sheet_pattern": u"Свод*задач*",
                u"merge_source_sheet": u"задачи_сотрудников",
                u"apply_to_summary": False,
                u"column_widths": [],
            }

        def normalize_merge_settings(obj):
            return {
                u"update_existing_only": False,
                u"check_modified": False,
                u"accumulate_comments": False,
            }

    if initial_general is None:
        initial_general = initial
    init_g = normalize_general_settings(initial_general)
    init_l = normalize_layout_settings(initial_layout, header_names=header_names)
    init_m = normalize_merge_settings(initial_merge)
    glob_g = normalize_general_settings(
        global_general if global_general is not None else init_g
    )
    glob_l = normalize_layout_settings(
        global_layout if global_layout is not None else init_l,
        header_names=header_names,
    )
    glob_m = normalize_merge_settings(
        global_merge if global_merge is not None else init_m
    )
    scope_choices = getattr(_cfg, "SETTINGS_SCOPE_CHOICES", ()) or (
        (u"private", u"Частные (эта книга)"),
        (u"global", u"Глобальные (по умолчанию)"),
    )
    scope_private = unicode(getattr(_cfg, "SETTINGS_SCOPE_PRIVATE", u"private"))
    scope_global = unicode(getattr(_cfg, "SETTINGS_SCOPE_GLOBAL", u"global"))
    start_scope = unicode(initial_scope or scope_private).strip()
    if start_scope != scope_global:
        start_scope = scope_private
    private_pack = {
        u"general": init_g,
        u"layout": init_l,
        u"merge": init_m,
    }
    global_pack = {
        u"general": glob_g,
        u"layout": glob_l,
        u"merge": glob_m,
    }

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_w = int(_cfg.EDIT_BTN_W)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = max(int(_cfg.EDIT_DIALOG_WIDTH), 560)
    content_w = dlg_w - 2 * m
    field_h = int(_cfg.EDIT_FIELD_H)
    lbl_h = int(_cfg.EDIT_LABEL_H)
    tab_h = field_h

    dm.Title = unicode(getattr(_cfg, "SETTINGS_DIALOG_TITLE", u"Настройки задач"))
    dm.Width = dlg_w
    dm.PositionX = 120
    dm.PositionY = 80

    y = m
    _add_label(
        dm,
        "ScopeLbl",
        unicode(getattr(_cfg, "SETTINGS_SCOPE_LABEL", u"Область настроек:")),
        m,
        y,
        content_w,
        lbl_h,
    )
    y += lbl_h + 2
    scope_labels = [lbl for _code, lbl in scope_choices]
    _add_dropdown(dm, "ScopeCombo", m, y, content_w, field_h, scope_labels)
    y += field_h + 2
    _add_label(
        dm,
        "ScopePathLbl",
        _settings_book_path_label(book_path, has_private=has_private),
        m,
        y,
        content_w,
        28,
    )
    y += 30
    tab_w = (content_w - gap) // 2
    _add_button(
        dm,
        "TabGeneralBtn",
        unicode(getattr(_cfg, "SETTINGS_TAB_GENERAL", u"Общие")),
        m,
        y,
        tab_w,
        tab_h,
    )
    _add_button(
        dm,
        "TabLayoutBtn",
        unicode(getattr(_cfg, "SETTINGS_TAB_LAYOUT", u"Размеры полей")),
        m + tab_w + gap,
        y,
        tab_w,
        tab_h,
    )
    y += tab_h + gap
    body_top = y

    # --- вкладка Общие ---
    gy = body_top
    _add_label(dm, "HintLbl", unicode(getattr(_cfg, "SETTINGS_DIALOG_HINT", u"")), m, gy, content_w, 44)
    gy += 48
    _add_label(dm, "PwdLbl", u"Пароль защиты листа:", m, gy, content_w, lbl_h)
    gy += lbl_h + 2
    show_w = 72
    pwd_w = content_w - show_w - gap
    _add_edit(dm, "PwdEd", m, gy, pwd_w, field_h, password=True)
    _add_edit(dm, "PwdPlainEd", m, gy, pwd_w, field_h, password=False)
    try:
        dm.getByName("PwdPlainEd").EnableVisible = False
    except Exception:
        pass
    _add_button(
        dm,
        "ShowPwdBtn",
        unicode(getattr(_cfg, "SETTINGS_SHOW_PWD_LABEL", u"Показать")),
        m + pwd_w + gap,
        gy,
        show_w,
        field_h,
    )
    gy += field_h + gap
    _add_label(dm, "FioLbl", u"ФИО владельца:", m, gy, content_w, lbl_h)
    gy += lbl_h + 2
    _add_edit(dm, "FioEd", m, gy, content_w, field_h)
    gy += field_h + gap
    role_choices = getattr(_cfg, "OWNER_ROLE_CHOICES", ())
    role_labels = [lbl for _code, lbl in role_choices]
    _add_label(dm, "RoleLbl", u"Роль владельца:", m, gy, content_w, lbl_h)
    gy += lbl_h + 2
    _add_combo(dm, "RoleCombo", m, gy, content_w, field_h, role_labels)
    gy += field_h + gap
    _add_label(
        dm,
        "MergeSrcLbl",
        unicode(
            getattr(
                _cfg, "SETTINGS_MERGE_SOURCE_LABEL", u"Лист с задачами сотрудников:"
            )
        ),
        m,
        gy,
        content_w,
        lbl_h,
    )
    gy += lbl_h + 2
    _add_edit(dm, "MergeSrcEd", m, gy, content_w, field_h)
    gy += field_h + gap
    merge_chk_h = 16
    _add_checkbox(
        dm,
        "MergeExistChk",
        unicode(
            getattr(
                _cfg,
                "SETTINGS_MERGE_UPDATE_EXISTING_LABEL",
                u"Слияние: обновлять только существующие",
            )
        ),
        m,
        gy,
        content_w,
        merge_chk_h,
    )
    gy += merge_chk_h + 2
    _add_checkbox(
        dm,
        "MergeDateChk",
        unicode(
            getattr(
                _cfg,
                "SETTINGS_MERGE_CHECK_MODIFIED_LABEL",
                u"Слияние: контролировать дату-время редактирования",
            )
        ),
        m,
        gy,
        content_w,
        merge_chk_h,
    )
    gy += merge_chk_h + 2
    _add_checkbox(
        dm,
        "MergeCmtChk",
        unicode(
            getattr(
                _cfg,
                "SETTINGS_MERGE_ACCUMULATE_COMMENTS_LABEL",
                u"Слияние: накапливать комментарии",
            )
        ),
        m,
        gy,
        content_w,
        merge_chk_h,
    )
    gy += merge_chk_h + gap
    log_lbl = unicode(getattr(_cfg, "SETTINGS_CONSOLE_LOG_LABEL", u"Подробные логи"))
    _add_checkbox(dm, "ConsoleLogChk", log_lbl, m, gy, content_w, 16)
    gy += 20 + gap
    _add_label(
        dm,
        "FontPtLbl",
        unicode(
            getattr(
                _cfg,
                "SETTINGS_DIALOG_FONT_LABEL",
                u"Размер шрифта в диалогах (pt):",
            )
        ),
        m,
        gy,
        content_w,
        lbl_h,
    )
    gy += lbl_h + 2
    _add_edit(dm, "FontPtEd", m, gy, 60, field_h)
    gy += field_h + gap
    general_bottom = gy

    general_names = [
        u"HintLbl",
        u"PwdLbl",
        u"PwdEd",
        u"PwdPlainEd",
        u"ShowPwdBtn",
        u"FioLbl",
        u"FioEd",
        u"RoleLbl",
        u"RoleCombo",
        u"MergeSrcLbl",
        u"MergeSrcEd",
        u"MergeExistChk",
        u"MergeDateChk",
        u"MergeCmtChk",
        u"ConsoleLogChk",
        u"FontPtLbl",
        u"FontPtEd",
    ]

    # --- вкладка Размеры (компактнее: меньше отступы, ширины в 2 колонки) ---
    ly = body_top
    lay_gap = 3
    _add_label(
        dm,
        "LayHintLbl",
        unicode(getattr(_cfg, "SETTINGS_LAYOUT_HINT", u"")),
        m,
        ly,
        content_w,
        28,
    )
    ly += 30
    _add_label(
        dm,
        "TplLbl",
        unicode(getattr(_cfg, "SETTINGS_TEMPLATE_SHEET_LABEL", u"Лист-шаблон:")),
        m,
        ly,
        content_w,
        lbl_h,
    )
    ly += lbl_h + 1
    _add_edit(dm, "TplEd", m, ly, content_w, field_h)
    ly += field_h + lay_gap

    # шаблоны задач / свода в один ряд
    half_pat = (content_w - lay_gap) // 2
    _add_label(
        dm,
        "TaskPatLbl",
        unicode(getattr(_cfg, "SETTINGS_TASK_PATTERN_LABEL", u"Шаблон задач:")),
        m,
        ly,
        half_pat,
        lbl_h,
    )
    _add_label(
        dm,
        "SumPatLbl",
        unicode(getattr(_cfg, "SETTINGS_SUMMARY_PATTERN_LABEL", u"Шаблон свода:")),
        m + half_pat + lay_gap,
        ly,
        half_pat,
        lbl_h,
    )
    ly += lbl_h + 1
    _add_edit(dm, "TaskPatEd", m, ly, half_pat, field_h)
    _add_edit(dm, "SumPatEd", m + half_pat + lay_gap, ly, half_pat, field_h)
    ly += field_h + lay_gap

    apply_sum_lbl = unicode(
        getattr(
            _cfg,
            "SETTINGS_APPLY_TO_SUMMARY_LABEL",
            u"Применять размеры и к листам свода",
        )
    )
    _add_checkbox(dm, "ApplyToSummaryChk", apply_sum_lbl, m, ly, content_w, 16)
    ly += 18 + lay_gap

    _add_label(dm, "WGridLbl", u"Колонка / мм:", m, ly, content_w, lbl_h)
    ly += lbl_h + 1

    width_items = list(init_l.get(u"column_widths") or [])
    # гарантировать имена из header_names
    have = {}
    wi = 0
    while wi < len(width_items):
        have[unicode((width_items[wi] or {}).get(u"name") or u"")] = True
        wi = wi + 1
    hi = 0
    names = list(header_names or [])
    while hi < len(names):
        nm = unicode(names[hi] or u"").strip()
        hi = hi + 1
        if nm and nm not in have:
            width_items.append({u"name": nm, u"width_mm": 25.0})
            have[nm] = True
    cat_title = unicode(getattr(_cfg, "COL_CATEGORY", u"Категория"))
    if cat_title and cat_title not in have:
        width_items.append({u"name": cat_title, u"width_mm": 25.0})
        have[cat_title] = True

    from todo_task_lib import _split_width_items_for_layout_grid

    others, left_meta, right_meta, guid_item, category_item = (
        _split_width_items_for_layout_grid(width_items)
    )
    width_items = (
        list(others)
        + [guid_item, category_item]
        + list(left_meta)
        + list(right_meta)
    )

    max_others = 16
    col_gap = max(lay_gap, 6)
    half_w = (content_w - col_gap) // 2
    mm_w = 48
    name_gap = 4
    name_w = half_w - mm_w - name_gap
    if name_w < 48:
        name_w = 48
        mm_w = max(32, half_w - name_w - name_gap)
    row_step = field_h + 2
    width_edits = []
    grid_top = ly

    def _place_width_field(it, col, row):
        nm = unicode((it or {}).get(u"name") or u"")
        try:
            mm_val = float((it or {}).get(u"width_mm") or 0)
        except (TypeError, ValueError):
            mm_val = 0.0
        x0 = m + col * (half_w + col_gap)
        y0 = grid_top + row * row_step
        idx = len(width_edits)
        lbl_name = u"WName%s" % idx
        ed_name = u"WEd%s" % idx
        show_nm = nm
        if len(show_nm) > 32:
            show_nm = show_nm[:30] + u"…"
        _add_label(
            dm,
            lbl_name,
            show_nm,
            x0,
            y0,
            name_w,
            field_h,
            align=2,
            vertical_align=1,
        )
        _add_edit(dm, ed_name, x0 + name_w + name_gap, y0, mm_w, field_h)
        width_edits.append((nm, ed_name, mm_val))

    oi = 0
    while oi < len(others) and oi < max_others:
        _place_width_field(others[oi], oi % 2, oi // 2)
        oi = oi + 1
    others_rows = (min(len(others), max_others) + 1) // 2
    if others_rows < 0:
        others_rows = 0
    # GUID слева, Категория справа — одна строка
    guid_row = others_rows
    _place_width_field(guid_item, 0, guid_row)
    _place_width_field(category_item, 1, guid_row)
    # мета: даты слева; Создано / Отредактировано справа
    meta_rows = max(len(left_meta), len(right_meta))
    mr = 0
    while mr < meta_rows:
        row = guid_row + 1 + mr
        if mr < len(left_meta):
            _place_width_field(left_meta[mr], 0, row)
        if mr < len(right_meta):
            _place_width_field(right_meta[mr], 1, row)
        mr = mr + 1
    grid_rows = guid_row + 1 + meta_rows
    if grid_rows < 1:
        grid_rows = 0
    ly = grid_top + grid_rows * row_step + lay_gap
    more_n = 0
    if len(others) > max_others:
        more_n = len(others) - max_others
    if more_n > 0:
        _add_label(
            dm,
            "WMoreLbl",
            u"… ещё %s колонок (сохранятся прежние ширины)" % more_n,
            m,
            ly,
            content_w,
            lbl_h,
        )
        ly += lbl_h + 1

    # «Применить» — внизу вместо OK (на вкладке размеров)
    layout_bottom = ly

    layout_names = [
        u"LayHintLbl",
        u"TplLbl",
        u"TplEd",
        u"TaskPatLbl",
        u"TaskPatEd",
        u"SumPatLbl",
        u"SumPatEd",
        u"ApplyToSummaryChk",
        u"WGridLbl",
    ]
    if more_n > 0:
        layout_names.append(u"WMoreLbl")
    wr = 0
    while wr < len(width_edits):
        layout_names.append(u"WName%s" % wr)
        layout_names.append(width_edits[wr][1])
        wr = wr + 1

    y = max(general_bottom, layout_bottom) + 4
    ok_x = dlg_w - m - 2 * btn_w - gap
    cancel_x = dlg_w - m - btn_w
    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = "OkBtn"
    ok_btn.Label = unicode(_cfg.BTN_OK_LABEL)
    ok_btn.PositionX = ok_x
    ok_btn.PositionY = y
    ok_btn.Width = btn_w
    ok_btn.Height = btn_h
    try:
        ok_btn.DefaultButton = True
    except Exception:
        pass
    dm.insertByName("OkBtn", ok_btn)
    # на вкладке размеров — вместо OK (та же позиция)
    apply_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    apply_btn.Name = "ApplyWidthsBtn"
    apply_btn.Label = unicode(
        getattr(_cfg, "SETTINGS_APPLY_WIDTHS_LABEL", u"Применить")
    )
    apply_btn.PositionX = ok_x
    apply_btn.PositionY = y
    apply_btn.Width = btn_w
    apply_btn.Height = btn_h
    try:
        apply_btn.DefaultButton = False
    except Exception:
        pass
    dm.insertByName("ApplyWidthsBtn", apply_btn)
    cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = "CancelBtn"
    cancel_btn.Label = unicode(_cfg.BTN_CANCEL_LABEL)
    cancel_btn.PositionX = cancel_x
    cancel_btn.PositionY = y
    cancel_btn.Width = btn_w
    cancel_btn.Height = btn_h
    dm.insertByName("CancelBtn", cancel_btn)
    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=dm.Title, doc=doc)
    except Exception:
        pass
    try:
        scope_m = dm.getByName("ScopeCombo")
        dm.removeByName("ScopeCombo")
        dm.insertByName("ScopeCombo", scope_m)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        return None
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass

    def _combo_code(combo_name, choices, default_code):
        label = _get_combo_text(dialog, dm, combo_name)
        ci = 0
        while ci < len(choices):
            code, item_label = choices[ci]
            if label == item_label or label == code:
                return code
            ci = ci + 1
        return default_code

    def _password_is_visible():
        try:
            return bool(dm.getByName("PwdPlainEd").EnableVisible)
        except Exception:
            return False

    def _read_password_text():
        if _password_is_visible():
            return _get_control_text(dialog, dm, "PwdPlainEd")
        return _get_control_text(dialog, dm, "PwdEd")

    def _collect_general():
        log_on = True
        try:
            log_on = bool(dialog.getControl("ConsoleLogChk").getState())
        except Exception:
            log_on = True
        font_raw = _get_control_text(dialog, dm, "FontPtEd").strip()
        try:
            from todo_task_settings_lib import normalize_dialog_font_pt

            font_pt = normalize_dialog_font_pt(
                font_raw, getattr(_cfg, "DEFAULT_DIALOG_FONT_PT", 11)
            )
        except Exception:
            font_pt = int(getattr(_cfg, "DEFAULT_DIALOG_FONT_PT", 11))
        return normalize_general_settings(
            {
                u"protect_password": _read_password_text(),
                u"owner_name": _get_control_text(dialog, dm, "FioEd").strip(),
                u"owner_role": _combo_code(
                    u"RoleCombo",
                    role_choices,
                    getattr(_cfg, "OWNER_ROLE_MANAGER", u"manager"),
                ),
                u"console_log": log_on,
                u"dialog_font_pt": font_pt,
            }
        )

    def _collect_layout():
        widths = []
        # сохранённые за пределами видимой сетки
        saved = {}
        pack = global_pack if state.get(u"scope") == scope_global else private_pack
        src_widths = ((pack or {}).get(u"layout") or {}).get(u"column_widths")
        if not src_widths:
            src_widths = width_items
        si = 0
        while si < len(src_widths):
            it = src_widths[si]
            si = si + 1
            nm = unicode((it or {}).get(u"name") or u"")
            if not nm:
                continue
            try:
                saved[nm] = float((it or {}).get(u"width_mm") or 0)
            except (TypeError, ValueError):
                saved[nm] = 0.0
        wi2 = 0
        while wi2 < len(width_edits):
            nm, ed, _old = width_edits[wi2]
            wi2 = wi2 + 1
            raw = _get_control_text(dialog, dm, ed).strip().replace(u",", u".")
            try:
                mm = float(raw) if raw else 0.0
            except (TypeError, ValueError):
                mm = 0.0
            saved[nm] = mm
        for nm in saved:
            widths.append({u"name": nm, u"width_mm": saved[nm]})
        apply_sum = False
        try:
            apply_sum = bool(dialog.getControl("ApplyToSummaryChk").getState())
        except Exception:
            apply_sum = False
        return normalize_layout_settings(
            {
                u"template_sheet": _get_control_text(dialog, dm, "TplEd").strip(),
                u"task_sheet_pattern": _get_control_text(dialog, dm, "TaskPatEd").strip(),
                u"summary_sheet_pattern": _get_control_text(dialog, dm, "SumPatEd").strip(),
                u"merge_source_sheet": _get_control_text(dialog, dm, "MergeSrcEd").strip(),
                u"apply_to_summary": apply_sum,
                u"column_widths": widths,
            },
            header_names=header_names,
        )

    def _collect_merge():
        exist_only = False
        check_mod = False
        acc_cmt = False
        try:
            exist_only = bool(dialog.getControl("MergeExistChk").getState())
        except Exception:
            exist_only = False
        try:
            check_mod = bool(dialog.getControl("MergeDateChk").getState())
        except Exception:
            check_mod = False
        try:
            acc_cmt = bool(dialog.getControl("MergeCmtChk").getState())
        except Exception:
            acc_cmt = False
        return normalize_merge_settings(
            {
                u"update_existing_only": exist_only,
                u"check_modified": check_mod,
                u"accumulate_comments": acc_cmt,
            }
        )

    state = {
        u"ok": False,
        u"result": None,
        u"hide_timer": None,
        u"pwd_visible": False,
        u"tab": u"general",
        u"apply_widths": False,
        u"scope": start_scope,
        u"filling": False,
    }
    show_sec = int(getattr(_cfg, "SETTINGS_SHOW_PWD_SECONDS", 10))
    if show_sec < 1:
        show_sec = 10

    def _fmt_width_val(val):
        try:
            f = float(val)
        except (TypeError, ValueError):
            f = 0.0
        if float(f) == int(f):
            return unicode(int(f))
        return unicode(f)

    def _fill_pack(pack):
        g = normalize_general_settings((pack or {}).get(u"general"))
        lay = normalize_layout_settings(
            (pack or {}).get(u"layout"), header_names=header_names
        )
        mrg = normalize_merge_settings((pack or {}).get(u"merge"))
        init_pwd = unicode(g.get(u"protect_password") or u"")
        _set_control_text(dialog, dm, "PwdEd", init_pwd)
        _set_control_text(dialog, dm, "PwdPlainEd", init_pwd)
        if state.get(u"tab") == u"general":
            _set_control_visible(dialog, dm, "PwdEd", not bool(state.get(u"pwd_visible")))
            _set_control_visible(dialog, dm, "PwdPlainEd", bool(state.get(u"pwd_visible")))
        _set_control_text(dialog, dm, "FioEd", unicode(g.get(u"owner_name") or u""))
        role_init = _choice_label(role_choices, g.get(u"owner_role"))
        try:
            role_ctl = dialog.getControl("RoleCombo")
            if role_ctl is not None and role_init:
                role_ctl.setText(role_init)
        except Exception:
            pass
        try:
            dialog.getControl("ConsoleLogChk").setState(
                1 if g.get(u"console_log") else 0
            )
        except Exception:
            pass
        try:
            from todo_task_settings_lib import normalize_dialog_font_pt

            font_pt = normalize_dialog_font_pt(
                g.get(u"dialog_font_pt"),
                getattr(_cfg, "DEFAULT_DIALOG_FONT_PT", 11),
            )
        except Exception:
            font_pt = int(getattr(_cfg, "DEFAULT_DIALOG_FONT_PT", 11))
        _set_control_text(dialog, dm, "FontPtEd", unicode(int(font_pt)))
        _set_control_text(
            dialog, dm, "TplEd", unicode(lay.get(u"template_sheet") or u"")
        )
        _set_control_text(
            dialog, dm, "TaskPatEd", unicode(lay.get(u"task_sheet_pattern") or u"")
        )
        _set_control_text(
            dialog, dm, "SumPatEd", unicode(lay.get(u"summary_sheet_pattern") or u"")
        )
        _set_control_text(
            dialog, dm, "MergeSrcEd", unicode(lay.get(u"merge_source_sheet") or u"")
        )
        try:
            dialog.getControl("MergeExistChk").setState(
                1 if mrg.get(u"update_existing_only") else 0
            )
        except Exception:
            pass
        try:
            dialog.getControl("MergeDateChk").setState(
                1 if mrg.get(u"check_modified") else 0
            )
        except Exception:
            pass
        try:
            dialog.getControl("MergeCmtChk").setState(
                1 if mrg.get(u"accumulate_comments") else 0
            )
        except Exception:
            pass
        try:
            dialog.getControl("ApplyToSummaryChk").setState(
                1 if lay.get(u"apply_to_summary") else 0
            )
        except Exception:
            pass
        wmap = {}
        wlist = lay.get(u"column_widths") or []
        wi0 = 0
        while wi0 < len(wlist):
            it = wlist[wi0]
            wi0 = wi0 + 1
            nm = unicode((it or {}).get(u"name") or u"")
            if not nm:
                continue
            try:
                wmap[nm] = float((it or {}).get(u"width_mm") or 0)
            except (TypeError, ValueError):
                wmap[nm] = 0.0
        wi = 0
        while wi < len(width_edits):
            nm, ed, old_val = width_edits[wi]
            val = wmap[nm] if nm in wmap else old_val
            _set_control_text(dialog, dm, ed, _fmt_width_val(val))
            wi = wi + 1

    def _set_scope_combo(code):
        label = _choice_label(scope_choices, code)
        _select_dropdown_item(dialog, dm, "ScopeCombo", label, scope_labels)

    def _apply_scope(code, force=False):
        if code != scope_global:
            code = scope_private
        if (not force) and code == state.get(u"scope"):
            return
        state[u"scope"] = code
        pack = global_pack if code == scope_global else private_pack
        state[u"filling"] = True
        try:
            _set_scope_combo(code)
            _fill_pack(pack)
        finally:
            state[u"filling"] = False

    def _scope_from_combo():
        return _combo_code(u"ScopeCombo", scope_choices, scope_private)

    def _set_tab(tab_name):
        state[u"tab"] = tab_name
        show_g = tab_name == u"general"
        show_l = tab_name == u"layout"
        gi = 0
        while gi < len(general_names):
            _set_control_visible(dialog, dm, general_names[gi], show_g)
            gi = gi + 1
        if show_g:
            _set_control_visible(dialog, dm, "PwdPlainEd", bool(state.get(u"pwd_visible")))
            _set_control_visible(dialog, dm, "PwdEd", not bool(state.get(u"pwd_visible")))
        li = 0
        while li < len(layout_names):
            _set_control_visible(dialog, dm, layout_names[li], show_l)
            li = li + 1
        # на «Размеры полей»: Применить вместо OK; на «Общие» — наоборот
        _set_control_visible(dialog, dm, "OkBtn", show_g)
        _set_control_visible(dialog, dm, "ApplyWidthsBtn", show_l)
        try:
            if show_l:
                dialog.getControl("ApplyWidthsBtn").getModel().DefaultButton = True
                dialog.getControl("OkBtn").getModel().DefaultButton = False
            else:
                dialog.getControl("OkBtn").getModel().DefaultButton = True
                dialog.getControl("ApplyWidthsBtn").getModel().DefaultButton = False
        except Exception:
            pass
        try:
            # одинаковая позиция: Применить на месте OK
            ok_m = dm.getByName("OkBtn")
            ap_m = dm.getByName("ApplyWidthsBtn")
            ap_m.PositionX = ok_m.PositionX
            ap_m.PositionY = ok_m.PositionY
            ap_m.Width = ok_m.Width
            ap_m.Height = ok_m.Height
        except Exception:
            pass

    def _cancel_hide_timer():
        tmr = state.get(u"hide_timer")
        if tmr is None:
            return
        try:
            tmr.cancel()
        except Exception:
            pass
        state[u"hide_timer"] = None

    def _hide_password():
        text = _read_password_text()
        _set_control_text(dialog, dm, "PwdEd", text)
        _set_control_text(dialog, dm, "PwdPlainEd", text)
        if state.get(u"tab") == u"general":
            _set_control_visible(dialog, dm, "PwdPlainEd", False)
            _set_control_visible(dialog, dm, "PwdEd", True)
        state[u"pwd_visible"] = False
        state[u"hide_timer"] = None
        try:
            dialog.getControl("ShowPwdBtn").getModel().Label = unicode(
                getattr(_cfg, "SETTINGS_SHOW_PWD_LABEL", u"Показать")
            )
        except Exception:
            pass

    def _hide_password_from_timer():
        try:
            from com.sun.star.awt import XCallback

            cb_svc = sm.createInstanceWithContext(
                "com.sun.star.awt.AsyncCallback", ctx
            )

            class _CB(unohelper.Base, XCallback):
                def notify(self, data):
                    _hide_password()

            cb_svc.addCallback(_CB(), None)
            return
        except Exception:
            pass
        _hide_password()

    def _show_password_briefly():
        import threading

        _cancel_hide_timer()
        text = _read_password_text()
        _set_control_text(dialog, dm, "PwdEd", text)
        _set_control_text(dialog, dm, "PwdPlainEd", text)
        _set_control_visible(dialog, dm, "PwdEd", False)
        _set_control_visible(dialog, dm, "PwdPlainEd", True)
        state[u"pwd_visible"] = True
        try:
            dialog.getControl("ShowPwdBtn").getModel().Label = u"Скрыть"
        except Exception:
            pass
        try:
            tmr = threading.Timer(float(show_sec), _hide_password_from_timer)
            try:
                tmr.daemon = True
            except Exception:
                pass
            tmr.start()
            state[u"hide_timer"] = tmr
        except Exception as err:
            _log("show password timer: %s" % err)

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            _cancel_hide_timer()

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "TabGeneralBtn":
                _set_tab(u"general")
                return
            if name == "TabLayoutBtn":
                _set_tab(u"layout")
                return
            if name == "ShowPwdBtn":
                if state.get(u"pwd_visible"):
                    _cancel_hide_timer()
                    _hide_password()
                else:
                    _show_password_briefly()
                return
            if name == "ApplyWidthsBtn":
                state[u"apply_widths"] = True
                state[u"result"] = {
                    u"general": _collect_general(),
                    u"layout": _collect_layout(),
                    u"merge": _collect_merge(),
                    u"apply_widths": True,
                    u"scope": state.get(u"scope") or scope_private,
                }
                state[u"ok"] = True
                _cancel_hide_timer()
                try:
                    dialog.endExecute()
                except Exception:
                    pass
                return
            if name == "OkBtn":
                _cancel_hide_timer()
                _hide_password()
                state[u"result"] = {
                    u"general": _collect_general(),
                    u"layout": _collect_layout(),
                    u"merge": _collect_merge(),
                    u"apply_widths": False,
                    u"scope": state.get(u"scope") or scope_private,
                }
                state[u"ok"] = True
            elif name == "CancelBtn":
                _cancel_hide_timer()
                state[u"ok"] = False
            else:
                return
            try:
                dialog.endExecute()
            except Exception:
                pass

    handler = _Handler()
    for btn_name in (
        u"OkBtn",
        u"CancelBtn",
        u"ShowPwdBtn",
        u"TabGeneralBtn",
        u"TabLayoutBtn",
        u"ApplyWidthsBtn",
    ):
        try:
            dialog.getControl(btn_name).addActionListener(handler)
        except Exception as err:
            _log("settings listener %s: %s" % (btn_name, err))

    try:
        from com.sun.star.awt import XItemListener

        class _ScopeItem(unohelper.Base, XItemListener):
            def disposing(self, event):
                pass

            def itemStateChanged(self, event):
                if state.get(u"filling"):
                    return
                try:
                    src = event.Source.getModel().Name
                except Exception:
                    src = u""
                if src != u"ScopeCombo" and src != "ScopeCombo":
                    return
                _apply_scope(_scope_from_combo())

        dialog.getControl("ScopeCombo").addItemListener(_ScopeItem())
    except Exception as err:
        _log("settings scope item listener: %s" % err)

    _apply_scope(start_scope, force=True)
    _set_tab(u"general")

    try:
        dialog.setVisible(True)
    except Exception:
        pass
    try:
        dialog.execute()
    finally:
        _cancel_hide_timer()
        try:
            dialog.dispose()
        except Exception:
            pass

    if not state[u"ok"]:
        return None
    return state.get(u"result")


def show_todo_task_refs_dialog(doc=None, sheet=None, dictionaries=None):
    """
    Редактор справочников.
    dictionaries — список dict из discover_ref_dictionaries.
    Возвращает {title: [row_values,...]} или None.
    """
    from todo_task_lib import (
        _parse_ref_item_text,
        _ref_item_display,
        dedupe_ref_items,
        read_ref_dictionary_items,
    )

    dictionaries = list(dictionaries or [])
    if not dictionaries:
        return None

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_w = int(_cfg.EDIT_BTN_W)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = max(int(_cfg.EDIT_DIALOG_WIDTH), 520)
    content_w = dlg_w - 2 * m
    field_h = int(_cfg.EDIT_FIELD_H)
    lbl_h = int(_cfg.EDIT_LABEL_H)

    dm.Title = unicode(getattr(_cfg, "REFS_DIALOG_TITLE", u"Справочники задач"))
    dm.Width = dlg_w
    dm.PositionX = 130
    dm.PositionY = 90

    y = m
    _add_label(
        dm,
        "HintLbl",
        unicode(getattr(_cfg, "REFS_DIALOG_HINT", u"")),
        m,
        y,
        content_w,
        36,
    )
    y += 40
    _add_label(
        dm,
        "DictLbl",
        unicode(getattr(_cfg, "REFS_DICT_LABEL", u"Справочник:")),
        m,
        y,
        content_w,
        lbl_h,
    )
    y += lbl_h + 2
    dict_titles = [unicode(d.get(u"title") or u"") for d in dictionaries]
    _add_combo(dm, "DictCombo", m, y, content_w, field_h, dict_titles)
    y += field_h + gap
    _add_label(
        dm,
        "ItemsLbl",
        unicode(getattr(_cfg, "REFS_ITEMS_LABEL", u"Элементы:")),
        m,
        y,
        content_w,
        lbl_h,
    )
    y += lbl_h + 2
    list_h = 120
    _add_listbox(dm, "ItemsLB", m, y, content_w, list_h, [])
    y += list_h + gap
    _add_label(
        dm,
        "ValLbl",
        unicode(getattr(_cfg, "REFS_VALUE_LABEL", u"Значение:")),
        m,
        y,
        content_w,
        lbl_h,
    )
    y += lbl_h + 2
    _add_edit(dm, "ValEd", m, y, content_w, field_h)
    y += field_h + gap

    act_w = (content_w - 3 * gap) // 4
    if act_w < 70:
        act_w = 70
    _add_button(
        dm,
        "AddBtn",
        unicode(getattr(_cfg, "REFS_BTN_ADD", u"Добавить")),
        m,
        y,
        act_w,
        btn_h,
    )
    _add_button(
        dm,
        "UpdBtn",
        unicode(getattr(_cfg, "REFS_BTN_UPDATE", u"Обновить")),
        m + act_w + gap,
        y,
        act_w,
        btn_h,
    )
    _add_button(
        dm,
        "DelBtn",
        unicode(getattr(_cfg, "REFS_BTN_DELETE", u"Удалить")),
        m + 2 * (act_w + gap),
        y,
        act_w,
        btn_h,
    )
    _add_button(
        dm,
        "DedupBtn",
        unicode(getattr(_cfg, "REFS_BTN_DEDUP", u"Убрать дубликаты")),
        m + 3 * (act_w + gap),
        y,
        max(act_w, content_w - 3 * (act_w + gap)),
        btn_h,
    )
    y += btn_h + gap + 4

    ok_x = dlg_w - m - 2 * btn_w - gap
    cancel_x = dlg_w - m - btn_w
    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = "OkBtn"
    ok_btn.Label = unicode(_cfg.BTN_OK_LABEL)
    ok_btn.PositionX = ok_x
    ok_btn.PositionY = y
    ok_btn.Width = btn_w
    ok_btn.Height = btn_h
    try:
        ok_btn.DefaultButton = True
    except Exception:
        pass
    dm.insertByName("OkBtn", ok_btn)
    cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = "CancelBtn"
    cancel_btn.Label = unicode(_cfg.BTN_CANCEL_LABEL)
    cancel_btn.PositionX = cancel_x
    cancel_btn.PositionY = y
    cancel_btn.Width = btn_w
    cancel_btn.Height = btn_h
    dm.insertByName("CancelBtn", cancel_btn)
    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=dm.Title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        return None
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass

    # working copy of items per dict title
    data = {}
    di = 0
    while di < len(dictionaries):
        d = dictionaries[di]
        di = di + 1
        title = unicode(d.get(u"title") or u"")
        data[title] = dedupe_ref_items(read_ref_dictionary_items(sheet, d))

    state = {u"ok": False, u"cur": dict_titles[0] if dict_titles else u""}

    def _dict_by_title(title):
        i = 0
        while i < len(dictionaries):
            if unicode(dictionaries[i].get(u"title") or u"") == title:
                return dictionaries[i]
            i = i + 1
        return None

    def _ncols(title):
        d = _dict_by_title(title)
        if not d:
            return 1
        return int(d[u"end_col"]) - int(d[u"start_col"]) + 1

    def _reload_list(keep_sel=None):
        title = state.get(u"cur") or u""
        items = data.get(title) or []
        lines = []
        i = 0
        while i < len(items):
            lines.append(_ref_item_display(items[i]))
            i = i + 1
        try:
            lb = dialog.getControl("ItemsLB")
            model = lb.getModel()
            model.StringItemList = tuple(lines)
            if keep_sel is not None and lines:
                idx = int(keep_sel)
                if idx < 0:
                    idx = 0
                if idx >= len(lines):
                    idx = len(lines) - 1
                try:
                    lb.selectItemPos(idx, True)
                except Exception:
                    pass
        except Exception as err:
            _log("refs reload list: %s" % err)

    def _selected_index():
        try:
            lb = dialog.getControl("ItemsLB")
            sel = list(lb.getSelectedItemsPos())
            if sel:
                return int(sel[0])
        except Exception:
            pass
        return -1

    def _load_selected_to_edit():
        title = state.get(u"cur") or _current_title()
        items = data.get(title) or []
        idx = _selected_index()
        if idx < 0 or idx >= len(items):
            return
        _set_control_text(dialog, dm, "ValEd", _ref_item_display(items[idx]))

    def _current_title():
        text = _get_combo_text(dialog, dm, "DictCombo").strip()
        if text in data:
            return text
        # match by label equality
        i = 0
        while i < len(dict_titles):
            if dict_titles[i] == text:
                return dict_titles[i]
            i = i + 1
        return state.get(u"cur") or (dict_titles[0] if dict_titles else u"")

    try:
        dialog.getControl("DictCombo").setText(dict_titles[0])
    except Exception:
        pass
    _reload_list()

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "DictCombo":
                state[u"cur"] = _current_title()
                _set_control_text(dialog, dm, "ValEd", u"")
                _reload_list()
                return
            if name == "AddBtn":
                title = _current_title()
                state[u"cur"] = title
                raw = _get_control_text(dialog, dm, "ValEd")
                vals = _parse_ref_item_text(raw, _ncols(title))
                if not any(unicode(v or u"").strip() for v in vals):
                    return
                items = list(data.get(title) or [])
                items.append(vals)
                data[title] = items
                _set_control_text(dialog, dm, "ValEd", u"")
                _reload_list(keep_sel=len(items) - 1)
                return
            if name == "UpdBtn":
                title = _current_title()
                state[u"cur"] = title
                idx = _selected_index()
                items = list(data.get(title) or [])
                if idx < 0 or idx >= len(items):
                    return
                raw = _get_control_text(dialog, dm, "ValEd")
                vals = _parse_ref_item_text(raw, _ncols(title))
                if not any(unicode(v or u"").strip() for v in vals):
                    return
                items[idx] = vals
                data[title] = items
                _reload_list(keep_sel=idx)
                return
            if name == "DelBtn":
                title = _current_title()
                state[u"cur"] = title
                idx = _selected_index()
                items = list(data.get(title) or [])
                if idx < 0 or idx >= len(items):
                    return
                del items[idx]
                data[title] = items
                _set_control_text(dialog, dm, "ValEd", u"")
                keep = idx if idx < len(items) else (len(items) - 1)
                if keep < 0:
                    keep = None
                _reload_list(keep_sel=keep)
                if keep is not None:
                    _load_selected_to_edit()
                return
            if name == "DedupBtn":
                title = _current_title()
                state[u"cur"] = title
                data[title] = dedupe_ref_items(data.get(title) or [])
                _set_control_text(dialog, dm, "ValEd", u"")
                _reload_list()
                return
            if name == "OkBtn":
                state[u"ok"] = True
                try:
                    dialog.endExecute()
                except Exception:
                    pass
                return
            if name == "CancelBtn":
                state[u"ok"] = False
                try:
                    dialog.endExecute()
                except Exception:
                    pass

    handler = _Handler()
    for n in (u"OkBtn", u"CancelBtn", u"AddBtn", u"UpdBtn", u"DelBtn", u"DedupBtn"):
        try:
            dialog.getControl(n).addActionListener(handler)
        except Exception as err:
            _log("refs listener %s: %s" % (n, err))
    # смена справочника в combo — ItemListener предпочтительнее, но action тоже бывает
    try:
        dialog.getControl("DictCombo").addActionListener(handler)
    except Exception:
        pass

    # poll combo changes via focus-out is hard; add ItemListener if available
    try:
        from com.sun.star.awt import XItemListener

        class _Item(unohelper.Base, XItemListener):
            def disposing(self, event):
                pass

            def itemStateChanged(self, event):
                try:
                    src = event.Source.getModel().Name
                except Exception:
                    src = u""
                if src == u"ItemsLB" or src == "ItemsLB":
                    _load_selected_to_edit()
                    return
                state[u"cur"] = _current_title()
                _set_control_text(dialog, dm, "ValEd", u"")
                _reload_list()

        item_listener = _Item()
        dialog.getControl("DictCombo").addItemListener(item_listener)
        dialog.getControl("ItemsLB").addItemListener(item_listener)
    except Exception as err:
        _log("refs item listener: %s" % err)

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

    if not state[u"ok"]:
        return None
    # final dedupe on save
    out = {}
    for title in data:
        out[title] = dedupe_ref_items(data[title])
    return out


def show_todo_task_move_dialog( doc=None, tasks=None, dest_names=None, dest_choices=None, source_name=u"", default_dest_name=u"" ):
    """
    Диалог перемещения/копирования задач.
    tasks — список dict {name, guid};
    dest_choices — [{name, label}, …] (label с расшифровкой B1);
    dest_names — устаревший список имён листов.
    default_dest_name — имя листа для предустановки combo (напр. по категории).
    Возвращает {dest_name, copy, values_only, add_columns} или None.
    """
    tasks = list(tasks or [])
    default_dest_name = unicode(default_dest_name or u"").strip()
    choices = []
    if dest_choices:
        ci = 0
        while ci < len(dest_choices):
            ch = dest_choices[ci]
            ci = ci + 1
            if not isinstance(ch, dict):
                continue
            nm = unicode(ch.get(u"name") or u"").strip()
            if not nm:
                continue
            lab = unicode(ch.get(u"label") or nm).strip() or nm
            choices.append({u"name": nm, u"label": lab})
    else:
        di = 0
        names = [unicode(x) for x in (dest_names or []) if unicode(x).strip()]
        while di < len(names):
            choices.append({u"name": names[di], u"label": names[di]})
            di = di + 1
    dest_labels = []
    label_to_name = {}
    ci = 0
    while ci < len(choices):
        lab = choices[ci][u"label"]
        dest_labels.append(lab)
        label_to_name[lab] = choices[ci][u"name"]
        # также по имени листа
        label_to_name[choices[ci][u"name"]] = choices[ci][u"name"]
        ci = ci + 1

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_w = int(_cfg.EDIT_BTN_W)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = int(_cfg.EDIT_DIALOG_WIDTH)
    content_w = dlg_w - 2 * m
    field_h = int(_cfg.EDIT_FIELD_H)
    lbl_h = int(_cfg.EDIT_LABEL_H)
    preview_len = int(getattr(_cfg, "MOVE_NAME_PREVIEW_LEN", 72))

    dm.Title = unicode(getattr(_cfg, "MOVE_DIALOG_TITLE", u"Перемещение задач"))
    dm.Width = dlg_w
    dm.PositionX = 140
    dm.PositionY = 90

    y = m
    hint = unicode(getattr(_cfg, "MOVE_DIALOG_HINT", u""))
    src = unicode(source_name or u"").strip()
    if src:
        hint = hint + u"\nИсточник: %s" % src
    _add_label(dm, "HintLbl", hint, m, y, content_w, 40)
    y += 44

    _add_label(
        dm,
        "CountLbl",
        u"Будет обработано задач: %s" % len(tasks),
        m,
        y,
        content_w,
        lbl_h,
    )
    y += lbl_h + 2

    lines = []
    ti = 0
    while ti < len(tasks):
        item = tasks[ti]
        name = unicode((item or {}).get(u"name") or u"").strip()
        if len(name) > preview_len:
            name = name[:preview_len] + u"…"
        guid = unicode((item or {}).get(u"guid") or u"").strip()
        lines.append(u"%s  [%s]" % (name, guid))
        ti = ti + 1
    list_h = 90
    _add_listbox(dm, "TasksLB", m, y, content_w, list_h, lines)
    y += list_h + gap

    _add_label(dm, "DestLbl", u"Лист назначения:", m, y, content_w, lbl_h)
    y += lbl_h + 2
    _add_combo(dm, "DestCombo", m, y, content_w, field_h, dest_labels)
    y += field_h + gap

    _add_checkbox(
        dm,
        "CopyChk",
        u"Копировать (оставить на исходном листе, новый GUID)",
        m,
        y,
        content_w,
        16,
    )
    y += 18
    _add_checkbox(
        dm,
        "ValuesOnlyChk",
        unicode(
            getattr(
                _cfg,
                "MOVE_VALUES_ONLY_LABEL",
                u"Только значения (без форматов и формул)",
            )
        ),
        m,
        y,
        content_w,
        16,
    )
    y += 18
    _add_checkbox(
        dm,
        "AddColsChk",
        unicode(
            getattr(
                _cfg,
                "MOVE_ADD_COLUMNS_LABEL",
                u"Добавить на приёмник недостающие колонки (в конец листа)",
            )
        ),
        m,
        y,
        content_w,
        16,
    )
    y += 20 + gap

    ok_x = dlg_w - m - 2 * btn_w - gap
    cancel_x = dlg_w - m - btn_w
    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = "OkBtn"
    ok_btn.Label = unicode(_cfg.BTN_OK_LABEL)
    ok_btn.PositionX = ok_x
    ok_btn.PositionY = y
    ok_btn.Width = btn_w
    ok_btn.Height = btn_h
    try:
        ok_btn.DefaultButton = True
    except Exception:
        pass
    dm.insertByName("OkBtn", ok_btn)

    cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = "CancelBtn"
    cancel_btn.Label = unicode(_cfg.BTN_CANCEL_LABEL)
    cancel_btn.PositionX = cancel_x
    cancel_btn.PositionY = y
    cancel_btn.Width = btn_w
    cancel_btn.Height = btn_h
    dm.insertByName("CancelBtn", cancel_btn)

    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=dm.Title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        return None
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass

    try:
        dest_ctl = dialog.getControl("DestCombo")
        if dest_ctl is not None and dest_labels:
            preset = u""
            if default_dest_name:
                di = 0
                while di < len(choices):
                    if choices[di][u"name"] == default_dest_name:
                        preset = choices[di][u"label"]
                        break
                    di = di + 1
                if not preset:
                    # case-insensitive
                    dcf = (
                        default_dest_name.casefold()
                        if hasattr(default_dest_name, "casefold")
                        else default_dest_name.lower()
                    )
                    di = 0
                    while di < len(choices):
                        nm = choices[di][u"name"]
                        ncf = nm.casefold() if hasattr(nm, "casefold") else nm.lower()
                        if ncf == dcf:
                            preset = choices[di][u"label"]
                            break
                        di = di + 1
            dest_ctl.setText(preset or dest_labels[0])
    except Exception:
        pass
    try:
        dialog.getControl("CopyChk").setState(0)
    except Exception:
        pass
    try:
        dialog.getControl("ValuesOnlyChk").setState(
            1 if getattr(_cfg, "DEFAULT_MOVE_VALUES_ONLY", True) else 0
        )
    except Exception:
        pass
    try:
        dialog.getControl("AddColsChk").setState(
            1 if getattr(_cfg, "DEFAULT_MOVE_ADD_COLUMNS", False) else 0
        )
    except Exception:
        pass

    state = {u"ok": False, u"result": None}

    def _resolve_dest_name(raw):
        text = unicode(raw or u"").strip()
        if not text:
            return u""
        if text in label_to_name:
            return label_to_name[text]
        # выбран индекс combo
        try:
            ctl = dialog.getControl("DestCombo")
            pos = int(ctl.getSelectedItemPos())
            if 0 <= pos < len(choices):
                return choices[pos][u"name"]
        except Exception:
            pass
        # «Имя (категория)» → Имя
        if u"(" in text and text.endswith(u")"):
            base = text[: text.rfind(u"(")].strip()
            if base in label_to_name:
                return label_to_name[base]
            return base
        return text

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "OkBtn":
                dest_raw = _get_combo_text(
                    dialog, dm, u"DestCombo", items=dest_labels
                ).strip()
                dest = _resolve_dest_name(dest_raw)
                copy_on = False
                try:
                    copy_on = bool(dialog.getControl("CopyChk").getState())
                except Exception:
                    pass
                values_only = True
                try:
                    values_only = bool(dialog.getControl("ValuesOnlyChk").getState())
                except Exception:
                    values_only = True
                add_columns = False
                try:
                    add_columns = bool(dialog.getControl("AddColsChk").getState())
                except Exception:
                    add_columns = False
                state[u"result"] = {
                    u"dest_name": dest,
                    u"copy": copy_on,
                    u"values_only": values_only,
                    u"add_columns": add_columns,
                }
                state[u"ok"] = True
            elif name == "CancelBtn":
                state[u"ok"] = False
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
            _log("move listener %s: %s" % (btn_name, err))

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

    if not state[u"ok"]:
        return None
    return state.get(u"result")


def show_todo_task_move_guid_conflict_dialog( doc=None, dest_sheet_name=u"", task_name=u"", dest_task_name=u"", guid=u"", remaining=1, conflict_index=1, conflict_total=1, existing_sheet_name=u""):
    """
    Диалог при совпадении GUID на листах задач.
    existing_sheet_name — где сейчас лежит совпадение (может отличаться от dest).
    Возвращает {action, apply_all} или None (прервать весь перенос).
    """
    act_overwrite = unicode(
        getattr(_cfg, "MOVE_GUID_ACTION_OVERWRITE", u"overwrite")
    )
    act_merge = unicode(
        getattr(_cfg, "MOVE_GUID_ACTION_MERGE_COMMENTS", u"merge_comments")
    )
    act_new = unicode(getattr(_cfg, "MOVE_GUID_ACTION_NEW_GUID", u"new_guid"))
    preview = int(getattr(_cfg, "MOVE_GUID_NAME_PREVIEW_LEN", 64))

    def _clip(text):
        s = unicode(text or u"").strip()
        if len(s) > preview:
            return s[:preview] + u"…"
        return s

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = int(_cfg.EDIT_DIALOG_WIDTH)
    content_w = dlg_w - 2 * m
    hint_h = int(getattr(_cfg, "MOVE_GUID_CONFLICT_TEXT_H", 120))
    show_apply_all = int(remaining or 1) > 1

    title = unicode(
        getattr(_cfg, "MOVE_GUID_CONFLICT_TITLE", u"Совпадение GUID")
    )
    dm.Title = title
    dm.Width = dlg_w
    dm.PositionX = 140
    dm.PositionY = 120

    y = m
    hint = unicode(getattr(_cfg, "MOVE_GUID_CONFLICT_HINT", u""))
    hint_ctl = _add_label(dm, "HintLbl", hint, m, y, content_w, hint_h)
    try:
        hint_ctl.MultiLine = True
    except Exception:
        pass
    try:
        hint_ctl.Align = 0
    except Exception:
        pass
    try:
        hint_ctl.VerticalAlign = 0
    except Exception:
        pass
    y += hint_h + gap

    details = []
    dest_nm = unicode(dest_sheet_name or u"").strip()
    exist_nm = unicode(existing_sheet_name or u"").strip()
    if dest_nm:
        details.append(u"Лист назначения: %s" % dest_nm)
    if exist_nm:
        if dest_nm:
            dcf = dest_nm.casefold() if hasattr(dest_nm, "casefold") else dest_nm.lower()
            ecf = (
                exist_nm.casefold() if hasattr(exist_nm, "casefold") else exist_nm.lower()
            )
            if ecf != dcf:
                details.append(u"Сейчас на листе: %s" % exist_nm)
            else:
                details.append(u"Уже на листе назначения")
        else:
            details.append(u"Сейчас на листе: %s" % exist_nm)
    details.append(u"Переносимая: %s" % (_clip(task_name) or u"—"))
    dest_tn = _clip(dest_task_name)
    if dest_tn:
        details.append(u"Существующая: %s" % dest_tn)
    g = unicode(guid or u"").strip()
    if g:
        details.append(u"GUID: %s" % g)
    try:
        total = int(conflict_total or 1)
        idx = int(conflict_index or 1)
    except (TypeError, ValueError):
        total = 1
        idx = 1
    if total > 1:
        details.append(u"Совпадение %s из %s" % (idx, total))
    details_h = 16 * len(details) + 4
    if details_h < 36:
        details_h = 36
    det_ctl = _add_label(
        dm, "DetailLbl", u"\n".join(details), m, y, content_w, details_h
    )
    try:
        det_ctl.MultiLine = True
    except Exception:
        pass
    y += details_h + gap

    if show_apply_all:
        _add_checkbox(
            dm,
            "ApplyAllChk",
            unicode(
                getattr(
                    _cfg,
                    "MOVE_GUID_APPLY_ALL_LABEL",
                    u"Для всех оставшихся совпадений",
                )
            ),
            m,
            y,
            content_w,
            16,
        )
        y += 20 + gap

    btn_w = int((content_w - gap) / 2)
    _add_button(
        dm,
        "OverwriteBtn",
        unicode(getattr(_cfg, "MOVE_GUID_BTN_OVERWRITE", u"Перезаписать")),
        m,
        y,
        btn_w,
        btn_h,
        default=True,
    )
    _add_button(
        dm,
        "MergeBtn",
        unicode(getattr(_cfg, "MOVE_GUID_BTN_MERGE", u"Слить комментарии")),
        m + btn_w + gap,
        y,
        btn_w,
        btn_h,
    )
    y += btn_h + gap
    _add_button(
        dm,
        "NewGuidBtn",
        unicode(getattr(_cfg, "MOVE_GUID_BTN_NEW", u"Новая задача")),
        m,
        y,
        btn_w,
        btn_h,
    )
    _add_button(
        dm,
        "AbortBtn",
        unicode(getattr(_cfg, "MOVE_GUID_BTN_ABORT", u"Прервать")),
        m + btn_w + gap,
        y,
        btn_w,
        btn_h,
    )
    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        return None
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass
    if show_apply_all:
        try:
            dialog.getControl("ApplyAllChk").setState(0)
        except Exception:
            pass

    state = {u"ok": False, u"action": u"", u"apply_all": False}

    def _read_apply_all():
        if not show_apply_all:
            return False
        try:
            return bool(dialog.getControl("ApplyAllChk").getState())
        except Exception:
            return False

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "OverwriteBtn":
                state[u"action"] = act_overwrite
                state[u"apply_all"] = _read_apply_all()
                state[u"ok"] = True
            elif name == "MergeBtn":
                state[u"action"] = act_merge
                state[u"apply_all"] = _read_apply_all()
                state[u"ok"] = True
            elif name == "NewGuidBtn":
                state[u"action"] = act_new
                state[u"apply_all"] = _read_apply_all()
                state[u"ok"] = True
            elif name == "AbortBtn":
                state[u"ok"] = False
                state[u"action"] = u""
            else:
                return
            try:
                dialog.endExecute()
            except Exception:
                pass

    handler = _Handler()
    for btn_name in (
        u"OverwriteBtn",
        u"MergeBtn",
        u"NewGuidBtn",
        u"AbortBtn",
    ):
        try:
            dialog.getControl(btn_name).addActionListener(handler)
        except Exception as err:
            _log("guid conflict listener %s: %s" % (btn_name, err))

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

    if not state[u"ok"]:
        return None
    return {
        u"action": state.get(u"action") or u"",
        u"apply_all": bool(state.get(u"apply_all")),
    }


def show_todo_task_merge_confirm_dialog( doc=None, summary=None, initial=None ):
    """
    Подтверждение слияния: сводка + режим комментариев + только значения/форматы.
    summary — dict с полями source_name, total, plan_insert, plan_update, plan_skip.
    Возвращает {accumulate_comments, values_only} или None.
    """
    summary = summary or {}
    initial = initial or {}
    comment_choices = getattr(_cfg, "MERGE_COMMENT_MODE_CHOICES", ()) or (
        (u"accumulate", u"Объединять комментарии"),
        (u"replace", u"Полная замена"),
    )
    copy_choices = getattr(_cfg, "MERGE_COPY_MODE_CHOICES", ()) or (
        (u"values", u"Только значения"),
        (u"formats", u"С форматами"),
    )
    comment_labels = [lbl for _c, lbl in comment_choices]
    copy_labels = [lbl for _c, lbl in copy_choices]
    acc_code = unicode(
        getattr(_cfg, "MERGE_COMMENT_MODE_ACCUMULATE", u"accumulate")
    )
    val_code = unicode(getattr(_cfg, "MERGE_COPY_MODE_VALUES", u"values"))

    want_acc = initial.get(u"accumulate_comments")
    if want_acc is None:
        want_acc = bool(
            getattr(_cfg, "DEFAULT_MERGE_RUN_ACCUMULATE_COMMENTS", True)
        )
    want_vals = initial.get(u"values_only")
    if want_vals is None:
        want_vals = bool(getattr(_cfg, "DEFAULT_MERGE_RUN_VALUES_ONLY", True))

    def _label_for(choices, code, fallback_idx=0):
        i = 0
        while i < len(choices):
            if unicode(choices[i][0]) == unicode(code):
                return choices[i][1]
            i = i + 1
        if choices:
            return choices[fallback_idx][1]
        return u""

    def _code_for(choices, label, default_code):
        lab = unicode(label or u"").strip()
        i = 0
        while i < len(choices):
            if unicode(choices[i][1]) == lab:
                return unicode(choices[i][0])
            i = i + 1
        return unicode(default_code)

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_w = int(_cfg.EDIT_BTN_W)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = int(_cfg.EDIT_DIALOG_WIDTH)
    content_w = dlg_w - 2 * m
    field_h = int(_cfg.EDIT_FIELD_H)
    lbl_h = int(_cfg.EDIT_LABEL_H)

    title = unicode(getattr(_cfg, "MERGE_CONFIRM_TITLE", u"Слияние задач сотрудников"))
    dm.Title = title
    dm.Width = dlg_w
    dm.PositionX = 140
    dm.PositionY = 100

    y = m
    hint = unicode(getattr(_cfg, "MERGE_CONFIRM_HINT", u""))
    hint_ctl = _add_label(dm, "HintLbl", hint, m, y, content_w, 52)
    try:
        hint_ctl.MultiLine = True
    except Exception:
        pass
    y += 56

    src = unicode(summary.get(u"source_name") or u"").strip()
    total = int(summary.get(u"total") or 0)
    plan_ins = int(summary.get(u"plan_insert") or 0)
    plan_upd = int(summary.get(u"plan_update") or 0)
    plan_skip = int(summary.get(u"plan_skip") or 0)
    lines = []
    if src:
        lines.append(u"Источник: «%s»" % src)
    lines.append(u"Задач в своде: %s" % total)
    lines.append(
        u"План: добавить %s, обновить %s, пропустить %s"
        % (plan_ins, plan_upd, plan_skip)
    )
    plan_preview_line = unicode(summary.get(u"plan_preview_line") or u"").strip()
    if plan_preview_line:
        lines.append(plan_preview_line)
    sum_h = 16 * len(lines) + 4
    if sum_h < 40:
        sum_h = 40
    sum_ctl = _add_label(dm, "SumLbl", u"\n".join(lines), m, y, content_w, sum_h)
    try:
        sum_ctl.MultiLine = True
    except Exception:
        pass
    y += sum_h + gap

    _add_label(
        dm,
        "CmtLbl",
        unicode(getattr(_cfg, "MERGE_COMMENT_MODE_LABEL", u"Метод слияния задач:")),
        m,
        y,
        content_w,
        lbl_h,
    )
    y += lbl_h + 2
    _add_combo(dm, "CmtCombo", m, y, content_w, field_h, comment_labels)
    y += field_h + gap

    _add_label(
        dm,
        "CopyLbl",
        unicode(getattr(_cfg, "MERGE_COPY_MODE_LABEL", u"Копирование ячеек:")),
        m,
        y,
        content_w,
        lbl_h,
    )
    y += lbl_h + 2
    _add_combo(dm, "CopyCombo", m, y, content_w, field_h, copy_labels)
    y += field_h + gap + 4

    ok_x = dlg_w - m - 2 * btn_w - gap
    cancel_x = dlg_w - m - btn_w
    _add_button(dm, "OkBtn", unicode(_cfg.BTN_OK_LABEL), ok_x, y, btn_w, btn_h, default=True)
    _add_button(
        dm, "CancelBtn", unicode(_cfg.BTN_CANCEL_LABEL), cancel_x, y, btn_w, btn_h
    )
    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        return None
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass

    cmt_init = _label_for(
        comment_choices, acc_code if want_acc else unicode(
            getattr(_cfg, "MERGE_COMMENT_MODE_REPLACE", u"replace")
        )
    )
    copy_init = _label_for(
        copy_choices,
        val_code
        if want_vals
        else unicode(getattr(_cfg, "MERGE_COPY_MODE_FORMATS", u"formats")),
    )
    try:
        dialog.getControl("CmtCombo").setText(cmt_init)
    except Exception:
        pass
    try:
        dialog.getControl("CopyCombo").setText(copy_init)
    except Exception:
        pass

    state = {u"ok": False, u"result": None}

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "OkBtn":
                cmt_lab = _get_combo_text(
                    dialog, dm, u"CmtCombo", items=comment_labels
                )
                copy_lab = _get_combo_text(
                    dialog, dm, u"CopyCombo", items=copy_labels
                )
                cmt_code = _code_for(comment_choices, cmt_lab, acc_code)
                copy_code = _code_for(copy_choices, copy_lab, val_code)
                state[u"result"] = {
                    u"accumulate_comments": cmt_code == acc_code,
                    u"values_only": copy_code == val_code,
                }
                state[u"ok"] = True
            elif name == "CancelBtn":
                state[u"ok"] = False
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
            _log("merge confirm listener %s: %s" % (btn_name, err))

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

    if not state[u"ok"]:
        return None
    return state.get(u"result")


def show_todo_task_merge_report_dialog( doc=None, summary_text=u"", lines=None ):
    """
    Итоговый отчёт слияния: краткий текст + прокручиваемый список строк.
    """
    summary_text = unicode(summary_text or u"")
    lines = [unicode(x) for x in (lines or [])]
    ctx = _uno_context()
    toolkit = _get_toolkit()
    if ctx is None or toolkit is None:
        _log("merge report: %s\n%s" % (summary_text, u"\n".join(lines[:40])))
        return True

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_w = int(_cfg.EDIT_BTN_W)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = int(_cfg.EDIT_DIALOG_WIDTH)
    content_w = dlg_w - 2 * m
    list_h = int(getattr(_cfg, "MERGE_REPORT_LIST_H", 160))

    title = unicode(getattr(_cfg, "MERGE_REPORT_TITLE", u"Результат слияния"))
    dm.Title = title
    dm.Width = dlg_w
    dm.PositionX = 130
    dm.PositionY = 90

    y = m
    sum_h = 56
    sum_ctl = _add_label(dm, "SumLbl", summary_text, m, y, content_w, sum_h)
    try:
        sum_ctl.MultiLine = True
    except Exception:
        pass
    y += sum_h + gap
    _add_listbox(dm, "ReportLB", m, y, content_w, list_h, lines)
    y += list_h + gap + 4
    ok_x = dlg_w - m - btn_w
    _add_button(
        dm, "OkBtn", unicode(_cfg.BTN_OK_LABEL), ok_x, y, btn_w, btn_h, default=True
    )
    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        return True
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass

    state = {u"ok": True}

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                dialog.endExecute()
            except Exception:
                pass

    handler = _Handler()
    try:
        dialog.getControl("OkBtn").addActionListener(handler)
    except Exception as err:
        _log("merge report listener: %s" % err)

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
    return bool(state[u"ok"])


def show_todo_task_merge_editor_conflict_dialog( doc=None, task_name=u"", guid=u"", sheet_name=u"", owner_editor=u"", employee_editor=u"", owner_modified=u"", employee_modified=u"", remaining=1, conflict_index=1, conflict_total=1):
    """
    Конфликт «Отредактировано»: версия владельца книги vs версия сотрудника.
    Возвращает {action, apply_all} или None (прервать весь merge).
    action: keep_owner | take_employee
    """
    act_keep = unicode(
        getattr(_cfg, "MERGE_EDITOR_KEEP_OWNER", u"keep_owner")
    )
    act_take = unicode(
        getattr(_cfg, "MERGE_EDITOR_TAKE_EMPLOYEE", u"take_employee")
    )
    preview = int(getattr(_cfg, "MERGE_REPORT_NAME_PREVIEW", 48))

    def _clip(text):
        s = unicode(text or u"").strip()
        if len(s) > preview:
            return s[:preview] + u"…"
        return s

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = int(_cfg.EDIT_MARGIN)
    gap = int(_cfg.EDIT_GAP)
    btn_h = int(_cfg.EDIT_BTN_H)
    dlg_w = int(_cfg.EDIT_DIALOG_WIDTH)
    content_w = dlg_w - 2 * m
    hint_h = int(getattr(_cfg, "MERGE_EDITOR_CONFLICT_TEXT_H", 72))
    show_apply_all = int(remaining or 1) > 1

    title = unicode(
        getattr(_cfg, "MERGE_EDITOR_CONFLICT_TITLE", u"Конфликт редактирования")
    )
    dm.Title = title
    dm.Width = dlg_w
    dm.PositionX = 140
    dm.PositionY = 110

    y = m
    hint = unicode(getattr(_cfg, "MERGE_EDITOR_CONFLICT_HINT", u""))
    hint_ctl = _add_label(dm, "HintLbl", hint, m, y, content_w, hint_h)
    try:
        hint_ctl.MultiLine = True
    except Exception:
        pass
    y += hint_h + gap

    details = []
    details.append(u"Задача: %s" % (_clip(task_name) or u"—"))
    g = unicode(guid or u"").strip()
    if g:
        details.append(u"GUID: %s" % g)
    sh = unicode(sheet_name or u"").strip()
    if sh:
        details.append(u"Лист у руководителя: %s" % sh)
    details.append(
        u"Владелец книги: %s" % (_clip(owner_editor) or u"—")
    )
    if owner_modified:
        details.append(u"  дата редакции: %s" % _clip(owner_modified))
    details.append(
        u"Сотрудник (свод): %s" % (_clip(employee_editor) or u"—")
    )
    if employee_modified:
        details.append(u"  дата редакции: %s" % _clip(employee_modified))
    try:
        total = int(conflict_total or 1)
        idx = int(conflict_index or 1)
    except (TypeError, ValueError):
        total = 1
        idx = 1
    if total > 1:
        details.append(u"Конфликт %s из %s" % (idx, total))
    details_h = 16 * len(details) + 4
    if details_h < 48:
        details_h = 48
    det_ctl = _add_label(
        dm, "DetailLbl", u"\n".join(details), m, y, content_w, details_h
    )
    try:
        det_ctl.MultiLine = True
    except Exception:
        pass
    y += details_h + gap

    if show_apply_all:
        _add_checkbox(
            dm,
            "ApplyAllChk",
            unicode(
                getattr(
                    _cfg,
                    "MERGE_EDITOR_APPLY_ALL_LABEL",
                    u"Для всех оставшихся таких конфликтов",
                )
            ),
            m,
            y,
            content_w,
            16,
        )
        y += 20 + gap

    btn_w = int((content_w - gap) / 2)
    _add_button(
        dm,
        "KeepBtn",
        unicode(
            getattr(_cfg, "MERGE_EDITOR_BTN_KEEP", u"Оставить версию владельца")
        ),
        m,
        y,
        btn_w,
        btn_h,
        default=True,
    )
    _add_button(
        dm,
        "TakeBtn",
        unicode(
            getattr(_cfg, "MERGE_EDITOR_BTN_TAKE", u"Взять версию сотрудника")
        ),
        m + btn_w + gap,
        y,
        btn_w,
        btn_h,
    )
    y += btn_h + gap
    _add_button(
        dm,
        "AbortBtn",
        unicode(getattr(_cfg, "MOVE_GUID_BTN_ABORT", u"Прервать")),
        m,
        y,
        btn_w,
        btn_h,
    )
    dm.Height = y + btn_h + m

    try:
        _theme_todo_dialog(dm, title_text=title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc):
        try:
            dialog.dispose()
        except Exception:
            pass
        return None
    try:
        try_paint_titlebar(
            dialog,
            title_bg=DARK_BLUE_TITLE_BG,
            title_fg=DARK_BLUE_TITLE_FG,
        )
    except Exception:
        pass
    if show_apply_all:
        try:
            dialog.getControl("ApplyAllChk").setState(0)
        except Exception:
            pass

    state = {u"ok": False, u"action": u"", u"apply_all": False}

    def _read_apply_all():
        if not show_apply_all:
            return False
        try:
            return bool(dialog.getControl("ApplyAllChk").getState())
        except Exception:
            return False

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "KeepBtn":
                state[u"action"] = act_keep
                state[u"apply_all"] = _read_apply_all()
                state[u"ok"] = True
            elif name == "TakeBtn":
                state[u"action"] = act_take
                state[u"apply_all"] = _read_apply_all()
                state[u"ok"] = True
            elif name == "AbortBtn":
                state[u"ok"] = False
                state[u"action"] = u""
            else:
                return
            try:
                dialog.endExecute()
            except Exception:
                pass

    handler = _Handler()
    for btn_name in (u"KeepBtn", u"TakeBtn", u"AbortBtn"):
        try:
            dialog.getControl(btn_name).addActionListener(handler)
        except Exception as err:
            _log("merge editor conflict listener %s: %s" % (btn_name, err))

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

    if not state[u"ok"]:
        return None
    return {
        u"action": state.get(u"action") or u"",
        u"apply_all": bool(state.get(u"apply_all")),
    }
