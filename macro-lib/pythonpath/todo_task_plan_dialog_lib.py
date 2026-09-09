# -*- coding: utf-8 -*-
"""Полноэкранный визард плана выполнения задачи (ListBox + панель полей; Grid опционален)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.714"
import uno
import unohelper
from com.sun.star.awt import XActionListener, XFocusListener, XItemListener, XMouseListener, XTextListener

import todo_task_plan_cfg as _pcfg

try:
    from libre_macros_ui_theme import (
        DARK_BLUE_TITLE_BG,
        DARK_BLUE_TITLE_FG,
        apply_soft_gray_dark_blue_theme,
        try_paint_titlebar,
    )
except Exception:
    DARK_BLUE_TITLE_BG = 0x0D47A1
    DARK_BLUE_TITLE_FG = 0xFFFFFF

    def apply_soft_gray_dark_blue_theme(dm, title_text=None):
        return False

    def try_paint_titlebar(dlg, title_bg=None, title_fg=None):
        return False

try:
    from todo_task_dialog_lib import (
        _add_button,
        _add_checkbox,
        _add_combo,
        _add_edit,
        _add_label,
        _create_peer,
        _dialog_parent,
        _dialog_warn,
        _get_control_text,
        _get_toolkit,
        _set_control_text,
        _theme_todo_dialog,
        _uno_context,
    )
except Exception:
    raise ImportError(u"todo_task_dialog_lib required for plan wizard")

try:
    unicode
except NameError:
    unicode = str


def _trace(msg):
    """Всегда в stdout (для отладки зависаний визарда плана)."""
    try:
        import sys

        sys.stdout.write(u"[todo_task/plan] %s\n" % unicode(msg))
        sys.stdout.flush()
    except Exception:
        try:
            print(u"[todo_task/plan] %s" % msg)
        except Exception:
            pass


def _log(msg):
    _trace(msg)
    try:
        from todo_task_settings_lib import get_console_log_enabled

        if not get_console_log_enabled():
            return
    except Exception:
        pass


def _screen_work_area(toolkit):
    """Рабочая область экрана — как param_wizard (_wizard_toolkit_work_area)."""
    if toolkit is None:
        return 0, 0, 1024, 768
    for meth in (u"getWorkArea", u"getDisplayWorkArea"):
        if not hasattr(toolkit, meth):
            continue
        try:
            wa = getattr(toolkit, meth)()
            if wa is None:
                continue
            x = int(getattr(wa, u"X", 0) or 0)
            y = int(getattr(wa, u"Y", 0) or 0)
            w = int(getattr(wa, u"Width", 0) or 0)
            h = int(getattr(wa, u"Height", 0) or 0)
            if w > 0 and h > 0:
                return x, y, w, h
        except Exception:
            pass
    try:
        dev = toolkit.createScreenCompatibleDevice(0, "")
        rect = dev.getInfo().WorkArea
        return int(rect.X), int(rect.Y), int(rect.Width), int(rect.Height)
    except Exception:
        return 0, 0, 1024, 768


def _plan_wizard_peer_fill_workarea(dialog, toolkit):
    """Развернуть peer на рабочую область (param_wizard _nested_subdialog_peer_fill_workarea)."""
    sx, sy, sw, sh = _screen_work_area(toolkit)
    if sw <= 0 or sh <= 0 or dialog is None:
        return False
    try:
        from com.sun.star.awt.PosSize import POSSIZE
    except Exception:
        POSSIZE = 15
    try:
        peer = dialog.getPeer()
        if peer is not None:
            peer.setPosSize(int(sx), int(sy), int(sw), int(sh), POSSIZE)
            return True
    except Exception as exc:
        _log("fill workarea: %s" % exc)
    return False


def _plan_wizard_create_peer(dialog, toolkit, doc=None, parent_dialog=None):
    """createPeer как у субвизардов param_wizard + всегда fill workarea."""
    if dialog is None or toolkit is None:
        return False
    try:
        from libre_macros_ui_theme import prepare_dialog_soft_gray

        prepare_dialog_soft_gray(dialog)
    except Exception:
        pass
    parents = []
    parent_doc = _dialog_parent(doc)
    if parent_doc is not None:
        parents.append(parent_doc)
    if parent_dialog is not None:
        try:
            peer = parent_dialog.getPeer()
            if peer is not None and peer not in parents:
                parents.append(peer)
        except Exception:
            pass
    parents.append(None)
    ok = False
    idx = 0
    while idx < len(parents):
        try:
            dialog.createPeer(toolkit, parents[idx])
            ok = True
            break
        except Exception as exc:
            _log("createPeer idx=%s: %s" % (idx, exc))
        idx = idx + 1
    if ok:
        _plan_wizard_peer_fill_workarea(dialog, toolkit)
    return ok


# ---------------------------------------------------------------------------
# Split-форма: ListBox + панель полей (основной режим)
# ---------------------------------------------------------------------------


def _add_plan_listbox(dm, name, x, y, w, h):
    lb = dm.createInstance(u"com.sun.star.awt.UnoControlListBoxModel")
    lb.Name = name
    lb.PositionX = int(x)
    lb.PositionY = int(y)
    lb.Width = int(w)
    lb.Height = int(h)
    try:
        lb.MultiSelection = False
        lb.VScroll = True
    except Exception:
        pass
    try:
        from libre_macros_ui_theme import apply_listbox_font

        apply_listbox_font(lb)
    except Exception:
        pass
    dm.insertByName(name, lb)
    return lb


def _plan_wizard_content_height(sh, dm=None):
    """Высота клиентской области диалога (без полосы темы)."""
    return max(400, int(sh) - _plan_wizard_title_h(dm))


def _plan_list_box_height(list_top, status_y, gap=None):
    """Высота ListBox: строго до строки статуса."""
    if gap is None:
        gap = int(getattr(_pcfg, "PLAN_WIZARD_LIST_BOTTOM_GAP", 6))
    bottom = int(status_y) - int(gap)
    h = bottom - int(list_top)
    min_h = int(getattr(_pcfg, "PLAN_WIZARD_LIST_MIN_H", 80))
    if h < min_h:
        h = min_h
    return h


def _plan_row_list_caption(row, doc=None):
    """Подпись строки в ListBox: номер, название этапа, даты."""
    from todo_task_plan_lib import format_plan_date_field, plan_row_display_name

    max_len = int(getattr(_pcfg, "PLAN_WIZARD_LIST_STAGE_MAX", 60))
    ino = unicode(row.get(u"item_no") or u"").strip()
    depth = max(0, ino.count(u"."))
    indent = u"  " * depth
    prefix = u"★ " if row.get(u"is_summary") else u""
    if row.get(u"is_done"):
        prefix = unicode(getattr(_pcfg, "PLAN_WIZARD_LIST_DONE_PREFIX", u"[v] ")) + prefix
    stage = plan_row_display_name(row).replace(u"\n", u" ")
    if len(stage) > max_len:
        stage = stage[: max_len - 1] + u"…"
    if not stage:
        stage = u"—"
    sp = row.get(u"date_start_parts")
    ep = row.get(u"date_end_parts")
    ds = format_plan_date_field(sp) if sp else u"?"
    de = format_plan_date_field(ep) if ep else u"?"
    dates = u"(%s — %s)" % (ds, de)
    return u"%s%s%-5s  %s  %s" % (indent, prefix, ino, stage, dates)


def _list_selected_index(dialog):
    try:
        lb = dialog.getControl(u"PlanListBox")
        sel = lb.getSelectedItemsPos()
        if sel and len(sel) > 0:
            return int(sel[0])
    except Exception:
        pass
    return -1


def _list_select_index(dialog, idx):
    if idx < 0:
        return
    try:
        dialog.getControl(u"PlanListBox").selectItemPos(int(idx), True)
    except Exception:
        pass


def _list_update_caption(state, idx):
    dialog = state.get(u"dialog")
    rows = state.get(u"rows") or []
    if dialog is None or idx < 0 or idx >= len(rows):
        return
    cap = _plan_row_list_caption(rows[idx], doc=state.get(u"doc"))
    try:
        lb = dialog.getControl(u"PlanListBox")
        seq = list(lb.Model.StringItemList or ())
        if idx < len(seq):
            seq[idx] = cap
            lb.Model.StringItemList = tuple(seq)
    except Exception:
        pass


def _detail_panel_color_for_row(row, doc=None):
    from todo_task_plan_lib import load_holidays, plan_row_status_code

    code = plan_row_status_code(row, holidays=load_holidays(doc))
    done = unicode(getattr(_pcfg, "PLAN_ROW_STATUS_DONE", u"done"))
    overdue = unicode(getattr(_pcfg, "PLAN_ROW_STATUS_OVERDUE", u"overdue"))
    soon = unicode(getattr(_pcfg, "PLAN_ROW_STATUS_SOON", u"soon"))
    if code == done:
        return int(getattr(_pcfg, "PLAN_DETAIL_COLOR_DONE", 0x006400))
    if code == overdue:
        return int(getattr(_pcfg, "PLAN_DETAIL_COLOR_OVERDUE", 0xCC0000))
    if code == soon:
        return int(getattr(_pcfg, "PLAN_DETAIL_COLOR_SOON", 0xE6A800))
    return int(getattr(_pcfg, "PLAN_DETAIL_COLOR_NORMAL", 0x000000))


def _detail_panel_control_names():
    return (
        u"DetailLbl",
        u"ItemLbl",
        u"ItemEd",
        u"SumChk",
        u"DoneChk",
        u"StageLbl",
        u"StageEd",
        u"ResultLbl",
        u"ResultDocEd",
        u"StartLbl",
        u"StartEd",
        u"StartWdLbl",
        u"StartNoSnapChk",
        u"EndLbl",
        u"EndEd",
        u"EndWdLbl",
        u"EndNoSnapChk",
        u"DurLbl",
        u"DurEd",
        u"AsgLbl",
        u"AsgCb",
        u"CritLbl",
        u"CritCb",
    )


def _set_control_enabled(dialog, name, enabled):
    try:
        dialog.getControl(name).setEnable(bool(enabled))
    except Exception:
        pass


def _detail_panel_lock_names():
    """Поля, блокируемые при is_done (кроме ResultDocEd и DoneChk)."""
    return (
        u"ItemEd",
        u"SumChk",
        u"StageEd",
        u"StartEd",
        u"StartNoSnapChk",
        u"EndEd",
        u"EndNoSnapChk",
        u"DurEd",
        u"AsgCb",
        u"CritCb",
    )


def _set_detail_panel_locked(state, row=None):
    dialog = state.get(u"dialog")
    if dialog is None:
        return
    rows = state.get(u"rows") or []
    si = _selected_index(state)
    if row is None and 0 <= si < len(rows):
        row = rows[si]
    row = row or {}
    locked = bool(row.get(u"is_done"))
    for nm in _detail_panel_lock_names():
        _set_control_enabled(dialog, nm, not locked)


def _apply_plan_wizard_read_only(state):
    dialog = state.get(u"dialog")
    dm = state.get(u"dm")
    if dialog is None or not state.get(u"read_only"):
        return
    _set_control_enabled(dialog, u"SaveBtn", False)
    try:
        dialog.getControl(u"CancelBtn").setLabel(
            unicode(getattr(_pcfg, "PLAN_WIZARD_CLOSE_BTN", u"Закрыть"))
        )
    except Exception:
        pass
    try:
        hint = unicode(getattr(_pcfg, "PLAN_WIZARD_VIEW_HINT", u""))
        if hint:
            lbl = dm.getByName(u"HintLbl")
            lbl.Label = hint
            lbl.MultiLine = True
    except Exception:
        pass
    _set_detail_panel_locked(state)


def _apply_detail_panel_colors(state, row=None):
    dialog = state.get(u"dialog")
    dm = state.get(u"dm")
    if dialog is None or dm is None:
        return
    rows = state.get(u"rows") or []
    si = _selected_index(state)
    if row is None and 0 <= si < len(rows):
        row = rows[si]
    color = _detail_panel_color_for_row(row or {}, doc=state.get(u"doc"))
    err_color = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_ERR_COLOR", 0xCC0000))
    for nm in _detail_panel_control_names():
        try:
            dm.getByName(nm).TextColor = int(color)
        except Exception:
            pass
        try:
            dialog.getControl(nm).getModel().TextColor = int(color)
        except Exception:
            pass
    try:
        dm.getByName(u"DetailErrLbl").TextColor = err_color
    except Exception:
        pass


def _detail_set_err(dialog, dm, msg):
    t = unicode(msg or u"")
    err_color = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_ERR_COLOR", 0xCC0000))
    try:
        mdl = dm.getByName(u"DetailErrLbl")
        mdl.Label = t
        mdl.TextColor = err_color
    except Exception:
        pass
    try:
        ctl = dialog.getControl(u"DetailErrLbl")
        ctl.setLabel(t)
        try:
            ctl.getModel().TextColor = err_color
        except Exception:
            pass
    except Exception:
        pass


def _layout_plan_detail_panel(dm, rx, y0, rw):
    """Метка над полем; даты в один ряд; название этапа — сразу под пунктом."""
    gap = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_GAP", 10))
    lbl_h = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_LBL_H", 16))
    row_h = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_ROW_H", 26))
    sec = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_SECTION_GAP", 14))
    stage_h = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_STAGE_H", 80))
    result_h = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_RESULT_H", 72))
    date_w = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_DATE_W", 118))
    wd_w = 32
    no_snap_w = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_NO_SNAP_W", 80))
    err_ind = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_ERR_INDENT", 14))
    err_h = 36
    date_gap = 12
    half = max(160, (int(rw) - date_gap) // 2)
    start2_x = rx + half + date_gap

    def _pos(name, x, y, w, h):
        try:
            c = dm.getByName(name)
            c.PositionX = int(x)
            c.PositionY = int(y)
            c.Width = int(w)
            c.Height = int(h)
        except Exception:
            pass

    y = int(y0)
    _pos(u"DetailLbl", rx, y, rw, 20)
    y += 24

    _pos(u"ItemLbl", rx, y, rw, lbl_h)
    y += lbl_h + gap
    _pos(u"ItemEd", rx, y, 84, row_h)
    _pos(u"SumChk", rx + 92, y, max(120, (rw - 92) // 2 - 4), row_h)
    _pos(u"DoneChk", rx + 92 + max(120, (rw - 92) // 2 - 4) + 8, y, max(120, rw - 92 - max(120, (rw - 92) // 2 - 4) - 8), row_h)
    y += row_h + sec

    _pos(u"StageLbl", rx, y, rw, lbl_h)
    y += lbl_h + gap
    _pos(u"StageEd", rx, y, rw, stage_h)
    try:
        stg = dm.getByName(u"StageEd")
        stg.MultiLine = True
        stg.VScroll = True
    except Exception:
        pass
    y += stage_h + sec

    _pos(u"ResultLbl", rx, y, rw, lbl_h)
    y += lbl_h + gap
    _pos(u"ResultDocEd", rx, y, rw, result_h)
    try:
        rdoc = dm.getByName(u"ResultDocEd")
        rdoc.MultiLine = True
        rdoc.VScroll = True
    except Exception:
        pass
    y += result_h + sec

    _pos(u"StartLbl", rx, y, half, lbl_h)
    _pos(u"EndLbl", rx + half + date_gap, y, half, lbl_h)
    y += lbl_h + gap
    _pos(u"StartEd", rx, y, date_w, row_h)
    _pos(u"StartWdLbl", rx + date_w + 4, y, wd_w, row_h)
    _pos(u"StartNoSnapChk", rx + date_w + wd_w + 8, y, no_snap_w, row_h)
    _pos(u"EndEd", start2_x, y, date_w, row_h)
    _pos(u"EndWdLbl", start2_x + date_w + 4, y, wd_w, row_h)
    _pos(u"EndNoSnapChk", start2_x + date_w + wd_w + 8, y, no_snap_w, row_h)
    y += row_h + sec

    _pos(u"DurLbl", rx, y, rw, lbl_h)
    y += lbl_h + gap
    _pos(u"DurEd", rx, y, 72, row_h)
    y += row_h + sec

    _pos(u"AsgLbl", rx, y, rw, lbl_h)
    y += lbl_h + gap
    _pos(u"AsgCb", rx, y, rw, row_h)
    y += row_h + sec

    _pos(u"CritLbl", rx, y, rw, lbl_h)
    y += lbl_h + gap
    _pos(u"CritCb", rx, y, min(240, rw), row_h)
    y += row_h + sec

    _pos(u"DetailErrLbl", rx + err_ind, y, rw - err_ind, err_h)
    try:
        err = dm.getByName(u"DetailErrLbl")
        err.TextColor = int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_ERR_COLOR", 0xCC0000))
        err.MultiLine = True
        err.Align = 0
    except Exception:
        pass


def _create_plan_detail_panel(dm, rx, ry, rw, assignees, criticalities):
    """Создать контролы правой панели (позиции — _layout_plan_detail_panel)."""
    _add_label(dm, u"DetailLbl", u"Параметры строки", rx, ry, rw, 20)
    _add_label(dm, u"ItemLbl", u"Пункт:", rx, ry, rw, 16)
    _add_edit(dm, u"ItemEd", rx, ry, 84, 26, multiline=False)
    _add_checkbox(dm, u"SumChk", u"Итоговый пункт", rx, ry, 140, 26)
    _add_checkbox(dm, u"DoneChk", u"Выполнено", rx, ry, 120, 26)
    _add_label(dm, u"StageLbl", u"Название этапа:", rx, ry, rw, 16)
    _add_edit(dm, u"StageEd", rx, ry, rw, 80, multiline=True)
    _add_label(dm, u"ResultLbl", u"Результат / отчётный документ:", rx, ry, rw, 16)
    _add_edit(dm, u"ResultDocEd", rx, ry, rw, 72, multiline=True)
    _add_label(dm, u"StartLbl", u"Начало:", rx, ry, 120, 16)
    _add_edit(dm, u"StartEd", rx, ry, 118, 26, multiline=False)
    _add_label(dm, u"StartWdLbl", u"", rx, ry, 36, 26)
    _add_checkbox(
        dm,
        u"StartNoSnapChk",
        unicode(getattr(_pcfg, "PLAN_WIZARD_DETAIL_NO_SNAP_LBL", u"Не корр.")),
        rx,
        ry,
        int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_NO_SNAP_W", 80)),
        26,
    )
    _add_label(dm, u"EndLbl", u"Окончание:", rx, ry, 120, 16)
    _add_edit(dm, u"EndEd", rx, ry, 118, 26, multiline=False)
    _add_label(dm, u"EndWdLbl", u"", rx, ry, 36, 26)
    _add_checkbox(
        dm,
        u"EndNoSnapChk",
        unicode(getattr(_pcfg, "PLAN_WIZARD_DETAIL_NO_SNAP_LBL", u"Не корр.")),
        rx,
        ry,
        int(getattr(_pcfg, "PLAN_WIZARD_DETAIL_NO_SNAP_W", 80)),
        26,
    )
    _add_label(dm, u"DurLbl", u"Длительность, раб. дни:", rx, ry, rw, 16)
    _add_edit(dm, u"DurEd", rx, ry, 72, 26, multiline=False)
    _add_label(dm, u"AsgLbl", u"Ответственные:", rx, ry, rw, 16)
    _add_combo(dm, u"AsgCb", rx, ry, rw, 26, assignees)
    _add_label(dm, u"CritLbl", u"Критичность:", rx, ry, rw, 16)
    _add_combo(dm, u"CritCb", rx, ry, 240, 26, criticalities)
    _add_label(dm, u"DetailErrLbl", u"", rx, ry, rw, 36)
    for lbl_name in (
        u"DetailLbl",
        u"ItemLbl",
        u"StageLbl",
        u"ResultLbl",
        u"StartLbl",
        u"EndLbl",
        u"DurLbl",
        u"AsgLbl",
        u"CritLbl",
        u"StartWdLbl",
        u"EndWdLbl",
    ):
        try:
            dm.getByName(lbl_name).MultiLine = False
        except Exception:
            pass
    _layout_plan_detail_panel(dm, rx, ry, rw)


def _checkbox_state(dialog, name):
    try:
        return bool(dialog.getControl(name).getState())
    except Exception:
        return False


def _set_checkbox_state(dialog, name, checked):
    try:
        dialog.getControl(name).setState(1 if checked else 0)
    except Exception:
        pass


def _clear_detail_panel(state):
    dialog = state.get(u"dialog")
    dm = state.get(u"dm")
    if dialog is None or dm is None:
        return
    for nm in (
        u"ItemEd",
        u"StartEd",
        u"EndEd",
        u"DurEd",
        u"AsgCb",
        u"StageEd",
        u"ResultDocEd",
        u"CritCb",
    ):
        _set_control_text(dialog, dm, nm, u"")
    for nm in (u"StartWdLbl", u"EndWdLbl"):
        _set_fixed_label(dialog, dm, nm, u"")
    _detail_set_err(dialog, dm, u"")
    try:
        dialog.getControl(u"SumChk").setState(0)
        _set_checkbox_state(dialog, u"StartNoSnapChk", False)
        _set_checkbox_state(dialog, u"EndNoSnapChk", False)
        _set_checkbox_state(dialog, u"DoneChk", False)
    except Exception:
        pass


def _load_row_to_panel(state, idx):
    from todo_task_plan_lib import format_plan_date_field, plan_date_weekday_label

    dialog = state.get(u"dialog")
    dm = state.get(u"dm")
    rows = state.get(u"rows") or []
    if dialog is None or dm is None or idx < 0 or idx >= len(rows):
        _clear_detail_panel(state)
        return
    row = rows[idx]
    _set_control_text(dialog, dm, u"ItemEd", unicode(row.get(u"item_no") or u""))
    try:
        dialog.getControl(u"SumChk").setState(1 if row.get(u"is_summary") else 0)
    except Exception:
        pass
    sp = row.get(u"date_start_parts")
    ep = row.get(u"date_end_parts")
    _set_control_text(
        dialog, dm, u"StartEd", format_plan_date_field(sp) if sp else u""
    )
    _set_control_text(
        dialog, dm, u"EndEd", format_plan_date_field(ep) if ep else u""
    )
    _set_fixed_label(
        dialog, dm, u"StartWdLbl", plan_date_weekday_label(sp) if sp else u""
    )
    _set_fixed_label(
        dialog, dm, u"EndWdLbl", plan_date_weekday_label(ep) if ep else u""
    )
    _set_control_text(dialog, dm, u"DurEd", unicode(row.get(u"duration_wd") or u""))
    _set_control_text(dialog, dm, u"AsgCb", unicode(row.get(u"assignees") or u""))
    _set_control_text(dialog, dm, u"StageEd", unicode(row.get(u"stage_name") or u""))
    _set_control_text(
        dialog, dm, u"ResultDocEd", unicode(row.get(u"result_doc") or u"")
    )
    _set_control_text(dialog, dm, u"CritCb", unicode(row.get(u"criticality") or u""))
    _set_checkbox_state(dialog, u"StartNoSnapChk", bool(row.get(u"start_no_snap")))
    _set_checkbox_state(dialog, u"EndNoSnapChk", bool(row.get(u"end_no_snap")))
    _set_checkbox_state(dialog, u"DoneChk", bool(row.get(u"is_done")))
    _detail_set_err(dialog, dm, u"")
    _apply_detail_panel_colors(state, row=row)
    _set_detail_panel_locked(state, row=row)


def _handle_done_checkbox_toggle(state):
    """Подтверждение досрочного выполнения; сохранение и блокировка полей."""
    from todo_task_plan_lib import confirm_early_plan_done

    dialog = state.get(u"dialog")
    dm = state.get(u"dm")
    doc = state.get(u"doc")
    si = _selected_index(state)
    rows = state.get(u"rows") or []
    if dialog is None or si < 0 or si >= len(rows):
        return
    checked = _checkbox_state(dialog, u"DoneChk")
    if checked:
        probe = dict(rows[si])
        probe[u"date_end_parts"] = _parse_date_field(
            _get_control_text(dialog, dm, u"EndEd")
        )
        if not confirm_early_plan_done(doc, probe):
            state[u"form_loading"] = True
            try:
                _set_checkbox_state(dialog, u"DoneChk", False)
            finally:
                state[u"form_loading"] = False
            return
    _save_panel_to_row(state, si, validate_dates=False)
    _set_detail_panel_locked(state)


def _save_panel_to_row(state, idx, validate_dates=False, prefer_duration_field=False):
    from todo_task_plan_lib import item_no_to_sort_key, parent_of_item_no

    dialog = state.get(u"dialog")
    dm = state.get(u"dm")
    doc = state.get(u"doc")
    rows = state.get(u"rows") or []
    if dialog is None or dm is None or idx < 0 or idx >= len(rows):
        return True
    try:
        dur_field = int(
            float(_get_control_text(dialog, dm, u"DurEd").strip() or u"0")
        )
    except (TypeError, ValueError):
        dur_field = 0
    if validate_dates:
        ok, err, sp, ep, wd = _validate_plan_row_dates(
            dialog, dm, doc=doc, update_duration=not prefer_duration_field
        )
        if not ok:
            _detail_set_err(dialog, dm, err)
            if err:
                _dialog_warn(doc, err)
            return False
    else:
        ok, err, sp, ep, wd = _validate_plan_row_dates(
            dialog, dm, doc=doc, update_duration=not prefer_duration_field
        )
        if not ok:
            sp = _parse_date_field(_get_control_text(dialog, dm, u"StartEd"))
            ep = _parse_date_field(_get_control_text(dialog, dm, u"EndEd"))
            wd = None
    row = dict(rows[idx])
    row[u"item_no"] = _get_control_text(dialog, dm, u"ItemEd").strip()
    try:
        row[u"is_summary"] = bool(dialog.getControl(u"SumChk").getState())
    except Exception:
        row[u"is_summary"] = False
    row[u"date_start_parts"] = sp
    row[u"date_end_parts"] = ep
    if prefer_duration_field:
        row[u"duration_wd"] = dur_field
    elif sp and ep and wd is not None:
        row[u"duration_wd"] = int(wd)
    else:
        row[u"duration_wd"] = dur_field
    row[u"assignees"] = _get_control_text(dialog, dm, u"AsgCb").strip()
    row[u"stage_name"] = _get_control_text(dialog, dm, u"StageEd")
    row[u"result_doc"] = _get_control_text(dialog, dm, u"ResultDocEd")
    row[u"criticality"] = _get_control_text(dialog, dm, u"CritCb").strip()
    row[u"start_no_snap"] = _checkbox_state(dialog, u"StartNoSnapChk")
    row[u"end_no_snap"] = _checkbox_state(dialog, u"EndNoSnapChk")
    row[u"is_done"] = _checkbox_state(dialog, u"DoneChk")
    row[u"parent_item"] = parent_of_item_no(row[u"item_no"])
    row[u"sort_key"] = item_no_to_sort_key(row[u"item_no"])
    rows[idx] = row
    state[u"rows"] = rows
    _list_update_caption(state, idx)
    _apply_detail_panel_colors(state, row=row)
    _set_detail_panel_locked(state, row=row)
    return True


def _save_current_row_from_panel(state, validate_dates=False):
    return _save_panel_to_row(state, _selected_index(state), validate_dates=validate_dates)


def _list_refresh(state, keep_index=-1):
    dialog = state.get(u"dialog")
    rows = state.get(u"rows") or []
    doc = state.get(u"doc")
    si = int(keep_index)
    if si < 0:
        si = _selected_index(state)
    labels = [_plan_row_list_caption(r, doc=doc) for r in rows]
    state[u"form_loading"] = True
    try:
        try:
            dialog.getControl(u"PlanListBox").Model.StringItemList = tuple(labels)
        except Exception as exc:
            _trace(u"list refresh: %s" % exc)
        if rows:
            if si < 0:
                si = 0
            if si >= len(rows):
                si = len(rows) - 1
            state[u"selected"] = si
            _list_select_index(dialog, si)
            _load_row_to_panel(state, si)
        else:
            state[u"selected"] = 0
            _clear_detail_panel(state)
    finally:
        state[u"form_loading"] = False
    _update_status(state)


def _sync_selection_from_list(state):
    idx = _list_selected_index(state.get(u"dialog"))
    if idx >= 0:
        state[u"selected"] = idx
    _update_status(state)


def _refresh_view(state, keep_index=-1):
    if state.get(u"use_grid"):
        _grid_refresh_rows(state, keep_index=keep_index)
    else:
        _list_refresh(state, keep_index=keep_index)


def _sync_selection(state):
    if state.get(u"use_grid"):
        _sync_selection_from_grid(state)
    else:
        _sync_selection_from_list(state)


def _relayout_plan_wizard_split(dm, sw, sh, m, gap, toolbar_h, status_h, hint_h=None):
    """Раскладка split-формы на весь экран."""
    title_h = _plan_wizard_title_h(dm)
    total_h = int(sh)
    total_w = int(sw)
    if hint_h is None:
        hint_h = int(getattr(_pcfg, "PLAN_WIZARD_HINT_H", 44))
    hint_top = int(getattr(_pcfg, "PLAN_WIZARD_HINT_TOP_GAP", 8))
    tb_gap = int(getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_GRID_GAP", 8))
    split_gap = int(getattr(_pcfg, "PLAN_WIZARD_SPLIT_GAP", 8))
    footer_btn_w = int(getattr(_pcfg, "PLAN_WIZARD_FOOTER_BTN_W", 110))
    list_lbl_h = int(getattr(_pcfg, "PLAN_WIZARD_LIST_LBL_H", 18))
    try:
        dm.Width = total_w
        dm.Height = total_h
    except Exception:
        pass

    y = title_h + m + hint_top
    try:
        hint = dm.getByName(u"HintLbl")
        hint.PositionX = m
        hint.PositionY = y
        hint.Width = total_w - 2 * m
        hint.Height = hint_h
    except Exception:
        pass
    y += hint_h + gap

    toolbar_y = y
    tb_specs = getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_BTNS_LIST", None) or ()
    x = m
    for spec in tb_specs:
        name = spec[0]
        bw = int(spec[2] if len(spec) > 2 else spec[1])
        try:
            btn = dm.getByName(name)
            btn.PositionX = x
            btn.PositionY = toolbar_y
            btn.Width = bw
            btn.Height = toolbar_h
        except Exception:
            pass
        x += bw + 4
    try:
        save = dm.getByName(u"SaveBtn")
        save.PositionY = toolbar_y
        save.Height = toolbar_h
        save.Width = footer_btn_w
        save.PositionX = total_w - m - 2 * footer_btn_w - gap
    except Exception:
        pass
    try:
        cancel = dm.getByName(u"CancelBtn")
        cancel.PositionY = toolbar_y
        cancel.Height = toolbar_h
        cancel.Width = footer_btn_w
        cancel.PositionX = total_w - m - footer_btn_w
    except Exception:
        pass
    y += toolbar_h + tb_gap

    content_h = _plan_wizard_content_height(total_h, dm)
    try:
        dm.Height = content_h
    except Exception:
        pass
    status_y = content_h - m - status_h
    list_top = y + list_lbl_h + 2
    list_h = _plan_list_box_height(list_top, status_y, gap=gap)

    min_list_w = int(getattr(_pcfg, "PLAN_WIZARD_LIST_MIN_W", 280))
    half = max(min_list_w, (total_w - 2 * m - split_gap) // 2)
    left_w = half
    right_x = m + half + split_gap
    right_w = total_w - right_x - m

    try:
        ll = dm.getByName(u"PlanListLbl")
        ll.PositionX = m
        ll.PositionY = y
        ll.Width = left_w
        ll.Height = list_lbl_h
    except Exception:
        pass
    try:
        lb = dm.getByName(u"PlanListBox")
        lb.PositionX = m
        lb.PositionY = list_top
        lb.Width = left_w
        lb.Height = list_h
    except Exception:
        pass

    _layout_plan_detail_panel(dm, right_x, y, right_w)

    try:
        st = dm.getByName(u"StatusLbl")
        st.PositionX = m
        st.PositionY = status_y
        st.Width = total_w - 2 * m
        st.Height = status_h
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Grid (PLAN_WIZARD_USE_GRID=True) — код сохранён, по умолчанию не используется
# ---------------------------------------------------------------------------


def _grid_create_column(col_model, title, width=90):
    try:
        col = col_model.createColumn()
        col.Title = unicode(title)
        try:
            col.ColumnWidth = int(width)
        except Exception:
            pass
        col_model.addColumn(col)
    except Exception as exc:
        _log("grid column: %s" % exc)


def _grid_setup_columns(col_model):
    _grid_create_column(col_model, u"Пункт", 52)
    _grid_create_column(col_model, u"Итог", 40)
    _grid_create_column(col_model, u"Начало", 118)
    _grid_create_column(col_model, u"Окончание", 118)
    _grid_create_column(col_model, u"Дн.", 36)
    _grid_create_column(col_model, u"Ответственные", 140)
    _grid_create_column(col_model, u"Результат", 180)
    _grid_create_column(col_model, u"Крит.", 80)


def _build_plan_grid_model(dm, px, py, pw, ph, rows, doc=None):
    """Собрать UnoControlGridModel с колонками и preload (как BlocksGrid в param_wizard)."""
    if dm is None:
        raise RuntimeError(u"Нет модели диалога для создания Grid.")
    grid_model = dm.createInstance(u"com.sun.star.awt.grid.UnoControlGridModel")
    grid_model.Name = u"PlanGrid"
    grid_model.PositionX = int(px)
    grid_model.PositionY = int(py)
    grid_model.Width = int(pw)
    grid_model.Height = int(ph)
    try:
        grid_model.ShowColumnHeader = True
        grid_model.ShowRowHeader = False
        grid_model.UseGridLines = True
        grid_model.VScroll = True
        grid_model.HScroll = True
        grid_model.SelectionModel = 1
    except Exception:
        pass
    _grid_setup_columns(grid_model.ColumnModel)
    _grid_preload(grid_model, rows, doc=doc)
    return grid_model


def _grid_nudge_repaint(grid):
    """Пнуть peer Grid для перерисовки после смены данных."""
    if grid is None:
        return
    try:
        from com.sun.star.awt.PosSize import POSSIZE
    except Exception:
        POSSIZE = 15
    try:
        mdl = grid.getModel()
        x, y, w, h = int(mdl.PositionX), int(mdl.PositionY), int(mdl.Width), int(mdl.Height)
        grid.setPosSize(x, y, w, max(40, h - 1), POSSIZE)
        grid.setPosSize(x, y, w, h, POSSIZE)
    except Exception:
        pass
    try:
        peer = grid.getPeer()
        if peer is not None:
            try:
                peer.invalidate(0)
            except Exception:
                pass
    except Exception:
        pass
    try:
        grid.setVisible(False)
        grid.setVisible(True)
    except Exception:
        pass
    try:
        toolkit = _get_toolkit()
        if toolkit is not None:
            toolkit.processEventsToIdle()
    except Exception:
        pass


def _attach_grid_mouse(state, grid):
    """Подключить MouseListener к Grid (после замены контрола)."""

    class _GridMouse(unohelper.Base, XMouseListener):
        def __init__(self, st):
            self.state = st

        def disposing(self, event):
            pass

        def mousePressed(self, event):
            pass

        def mouseReleased(self, event):
            _sync_selection_from_grid(self.state)
            try:
                if int(event.ClickCount or 1) >= 2:
                    _edit_selected(self.state)
            except Exception:
                pass

        def mouseEntered(self, event):
            pass

        def mouseExited(self, event):
            pass

    try:
        grid.addMouseListener(_GridMouse(state))
    except Exception as exc:
        _log("grid mouse: %s" % exc)


def _grid_replace_control(state, rows, doc=None, keep_index=-1):
    """
    Заменить PlanGrid в модели диалога (единственный надёжный refresh в AO).
    """
    dialog = state.get(u"dialog")
    dm = state.get(u"dm")
    if dialog is None or dm is None:
        return False
    try:
        old = dm.getByName(u"PlanGrid")
        px = int(old.PositionX)
        py = int(old.PositionY)
        pw = int(old.Width)
        ph = int(old.Height)
    except Exception as exc:
        _trace(u"grid replace geom: %s" % exc)
        return False
    new_grid = _build_plan_grid_model(dm, px, py, pw, ph, rows, doc=doc)
    old_ctl = None
    try:
        old_ctl = dialog.getControl(u"PlanGrid")
    except Exception:
        pass
    if old_ctl is not None:
        try:
            old_ctl.setVisible(False)
        except Exception:
            pass
        try:
            old_ctl.dispose()
        except Exception as exc:
            _trace(u"grid dispose old: %s" % exc)
    try:
        dm.removeByName(u"PlanGrid")
    except Exception as exc:
        _trace(u"grid removeByName: %s" % exc)
    try:
        dm.insertByName(u"PlanGrid", new_grid)
    except Exception as exc:
        _trace(u"grid insertByName: %s" % exc)
        return False
    grid = None
    try:
        grid = dialog.getControl(u"PlanGrid")
    except Exception as exc:
        _trace(u"grid getControl: %s" % exc)
    if grid is None:
        return False
    try:
        from com.sun.star.awt.PosSize import POSSIZE
    except Exception:
        POSSIZE = 15
    try:
        grid.setPosSize(px, py, pw, ph, POSSIZE)
        grid.setVisible(True)
    except Exception:
        pass
    state[u"grid"] = grid
    state[u"grid_data_model"] = new_grid.GridDataModel
    _attach_grid_mouse(state, grid)
    _grid_nudge_repaint(grid)
    rc = _grid_row_count(state[u"grid_data_model"])
    _trace(u"grid replace: want=%s rc=%s" % (len(rows), rc))
    if keep_index >= 0 and keep_index < len(rows):
        try:
            _grid_select_row(grid, keep_index)
        except Exception:
            pass
    return True


def _grid_force_refresh(state, keep_index=-1):
    """Принудительно перерисовать Grid из state['rows']."""
    rows = state.get(u"rows") or []
    doc = state.get(u"doc")
    si = int(keep_index)
    if si < 0:
        si = _selected_index(state)
    if not _grid_replace_control(state, rows, doc=doc, keep_index=si):
        grid = state.get(u"grid")
        if grid is not None:
            _grid_rebuild_assign(
                grid, rows, doc=doc, state=state, keep_index=si
            )
            _grid_nudge_repaint(grid)
    _update_status(state)


def _grid_refresh_rows(state, keep_index=-1):
    """Обновить Grid после изменения данных (add/del/edit/move/recalc)."""
    _grid_force_refresh(state, keep_index=keep_index)


def _row_grid_tuple(row, doc=None):
    from todo_task_plan_lib import format_plan_date_display

    sp = row.get(u"date_start_parts")
    ep = row.get(u"date_end_parts")
    return (
        unicode(row.get(u"item_no") or u""),
        _pcfg.PLAN_SUMMARY_YES if row.get(u"is_summary") else u"",
        format_plan_date_display(doc, sp) if sp else u"",
        format_plan_date_display(doc, ep) if ep else u"",
        unicode(row.get(u"duration_wd") or u""),
        unicode(row.get(u"assignees") or u""),
        unicode(row.get(u"result_doc") or u""),
        unicode(row.get(u"criticality") or u""),
    )


def _grid_data_model(grid):
    try:
        return grid.getModel().GridDataModel
    except Exception:
        return None


def _grid_row_count(data_model):
    try:
        return int(data_model.getRowCount())
    except Exception:
        return 0


def _grid_index_from_heading(data_model, heading):
    try:
        h = unicode(heading)
    except Exception:
        h = u"%s" % heading
    rc = _grid_row_count(data_model)
    i = 0
    while i < rc:
        try:
            if unicode(data_model.getRowHeading(i)) == h:
                return i
        except Exception:
            pass
        i = i + 1
    return -1


def _grid_select_row(grid, idx):
    if idx < 0:
        return
    try:
        grid.selectRow(int(idx), True)
        return
    except Exception:
        pass
    try:
        grid.goToCell(0, int(idx))
    except Exception:
        pass


def _set_fixed_label(dialog, dm, name, text):
    t = unicode(text or u"")
    try:
        dialog.getControl(name).setLabel(t)
    except Exception:
        pass
    try:
        dm.getByName(name).Label = t
    except Exception:
        pass


def _plan_wizard_title_h(dm=None):
    """Высота внутренней полосы темы (резерв до и после apply_soft_gray)."""
    if dm is not None:
        try:
            from libre_macros_ui_theme import SOFT_GRAY_TITLE_H, soft_gray_theme_applied

            if soft_gray_theme_applied(dm):
                return int(SOFT_GRAY_TITLE_H)
        except Exception:
            pass
    try:
        from libre_macros_ui_theme import SOFT_GRAY_TITLE_H

        return int(SOFT_GRAY_TITLE_H)
    except Exception:
        return 28


def _row_grid_heading(row, index=0):
    """Заголовок строки Grid — индекс (как BlocksGrid в param_wizard)."""
    return str(int(index))


def _plan_toolbar_bottom(dm, fallback=0):
    """Нижняя граница верхнего ряда кнопок (для якоря Grid)."""
    names = (
        u"AddRootBtn",
        u"AddChildBtn",
        u"DelBtn",
        u"RefreshGridBtn",
        u"UpBtn",
        u"DownBtn",
        u"EditBtn",
        u"RecalcBtn",
    )
    bottom = int(fallback)
    for name in names:
        try:
            ctl = dm.getByName(name)
            y = int(ctl.PositionY) + int(ctl.Height)
            if y > bottom:
                bottom = y
        except Exception:
            pass
    return bottom


def _relayout_plan_wizard_full(dm, sw, sh, m, gap, toolbar_h, footer_h, status_h, hint_h=None):
    """Полная раскладка всех контролов под экран sw×sh (после темы)."""
    title_h = _plan_wizard_title_h(dm)
    total_h = int(sh)
    total_w = int(sw)
    if hint_h is None:
        hint_h = int(getattr(_pcfg, "PLAN_WIZARD_HINT_H", 44))
    hint_top = int(getattr(_pcfg, "PLAN_WIZARD_HINT_TOP_GAP", 8))
    tb_grid_gap = int(getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_GRID_GAP", 8))
    grid_shrink = int(getattr(_pcfg, "PLAN_WIZARD_GRID_BOTTOM_RESERVE", 200))
    footer_btn_w = int(getattr(_pcfg, "PLAN_WIZARD_FOOTER_BTN_W", 110))
    try:
        dm.Width = total_w
        dm.Height = total_h
    except Exception:
        pass

    y = title_h + m + hint_top
    try:
        hint = dm.getByName("HintLbl")
        hint.PositionX = m
        hint.PositionY = y
        hint.Width = total_w - 2 * m
        hint.Height = hint_h
    except Exception:
        pass
    y += hint_h + gap

    toolbar_y = y
    tb_specs = getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_BTNS", None) or ()
    x = m
    for spec in tb_specs:
        name = spec[0]
        bw = int(spec[2] if len(spec) > 2 else spec[1])
        try:
            btn = dm.getByName(name)
            btn.PositionX = x
            btn.PositionY = toolbar_y
            btn.Width = bw
            btn.Height = toolbar_h
        except Exception:
            pass
        x += bw + 4
    try:
        save = dm.getByName("SaveBtn")
        save.PositionY = toolbar_y
        save.Height = toolbar_h
        save.Width = footer_btn_w
        save.PositionX = total_w - m - 2 * footer_btn_w - gap
    except Exception:
        pass
    try:
        cancel = dm.getByName("CancelBtn")
        cancel.PositionY = toolbar_y
        cancel.Height = toolbar_h
        cancel.Width = footer_btn_w
        cancel.PositionX = total_w - m - footer_btn_w
    except Exception:
        pass
    y += toolbar_h + tb_grid_gap

    status_y = total_h - m - status_h
    grid_h = status_y - gap - y - grid_shrink
    min_h = int(getattr(_pcfg, "PLAN_WIZARD_GRID_MIN_H", 120))
    if grid_h < min_h:
        grid_h = min_h
    try:
        grid = dm.getByName("PlanGrid")
        grid.PositionX = m
        grid.PositionY = y
        grid.Width = total_w - 2 * m
        grid.Height = grid_h
    except Exception as exc:
        _log("layout grid: %s" % exc)
    try:
        st = dm.getByName("StatusLbl")
        st.PositionX = m
        st.PositionY = status_y
        st.Width = total_w - 2 * m
        st.Height = status_h
    except Exception:
        pass


def _relayout_plan_wizard_controls(dm, m, gap, btn_h, status_h):
    """Устаревший вызов — делегирует в полную раскладку по текущему размеру dm."""
    try:
        sw = int(dm.Width)
        sh = int(dm.Height)
    except Exception:
        sw, sh = 1024, 768
    toolbar_h = int(getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_BTN_H", btn_h))
    footer_h = int(getattr(_pcfg, "PLAN_WIZARD_FOOTER_BTN_H", btn_h))
    _relayout_plan_wizard_full(
        dm, sw, sh, m, gap, toolbar_h, footer_h, status_h
    )


def _fix_plan_wizard_layout(dm, dlg_h, m, gap, btn_h, status_h, sw=None, sh=None):
    """После темы — раскладка на весь экран."""
    title_h = _plan_wizard_title_h(dm)
    if sw is None:
        try:
            sw = int(dm.Width)
        except Exception:
            sw = 1024
    if sh is None:
        sh = int(dlg_h) + title_h
    try:
        dm.Height = int(sh)
    except Exception:
        pass
    toolbar_h = int(getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_BTN_H", btn_h))
    footer_h = int(getattr(_pcfg, "PLAN_WIZARD_FOOTER_BTN_H", btn_h))
    _relayout_plan_wizard_full(
        dm, sw, sh, m, gap, toolbar_h, footer_h, status_h
    )


def _sync_plan_wizard_control_layout(dialog, dm):
    """Синхронизировать все live-контролы с моделью после createPeer."""
    try:
        from com.sun.star.awt.PosSize import POSSIZE
    except Exception:
        POSSIZE = 15
    skip = {u"SoftGrayTitleBar"}
    try:
        names = list(dm.getElementNames())
    except Exception:
        names = []
    for name in names:
        if name in skip:
            continue
        try:
            mdl = dm.getByName(name)
            ctl = dialog.getControl(name)
            ctl.setPosSize(
                int(mdl.PositionX),
                int(mdl.PositionY),
                int(mdl.Width),
                int(mdl.Height),
                POSSIZE,
            )
        except Exception:
            pass


def _create_peer_standalone(dlg, toolkit):
    """Deprecated — используйте _plan_wizard_create_peer."""
    return _plan_wizard_create_peer(dlg, toolkit, doc=None, parent_dialog=None)


def _plan_wizard_fit_screen_peer(dialog, toolkit, sx, sy, sw, sh):
    """Deprecated — fill workarea внутри _plan_wizard_create_peer."""
    return _plan_wizard_peer_fill_workarea(dialog, toolkit)


def _grid_clear_data_model(data_model):
    """Очистить все строки GridDataModel."""
    if data_model is None:
        return
    try:
        ids = list(getattr(data_model, u"RowIds", ()) or ())
    except Exception:
        ids = []
    for rid in ids:
        try:
            data_model.removeRow(rid)
        except Exception:
            pass
    guard = 0
    while _grid_row_count(data_model) > 0 and guard < 5000:
        guard = guard + 1
        rc = _grid_row_count(data_model)
        before = rc
        try:
            data_model.removeRow(data_model.getRowHeading(rc - 1))
        except Exception:
            try:
                data_model.removeRow(str(rc - 1))
            except Exception:
                break
        if _grid_row_count(data_model) >= before:
            break


def _grid_fill_data_model(data_model, rows, doc=None):
    """Заполнить GridDataModel строками rows (заголовок = row_id)."""
    if data_model is None:
        return
    ri = 0
    while ri < len(rows):
        try:
            data_model.addRow(
                _row_grid_heading(rows[ri], ri),
                _row_grid_tuple(rows[ri], doc=doc),
            )
        except Exception as exc:
            _trace(u"grid fill addRow %s: %s" % (ri, exc))
        ri = ri + 1


def _grid_rebuild_assign(grid, rows, doc=None, state=None, keep_index=-1):
    """Полная пересборка GridDataModel (AO: removeRow/update ненадёжны после createPeer)."""
    ctx = _uno_context()
    if ctx is None:
        _trace(u"grid rebuild: no context")
        return
    try:
        sm = ctx.getServiceManager()
        new_dm = sm.createInstanceWithContext(
            u"com.sun.star.awt.grid.DefaultGridDataModel", ctx
        )
    except Exception as exc:
        _trace(u"grid rebuild new: %s" % exc)
        return
    i = 0
    while i < len(rows):
        try:
            new_dm.addRow(str(i), _row_grid_tuple(rows[i], doc=doc))
        except Exception as exc:
            _trace(u"grid rebuild add %s: %s" % (i, exc))
        i = i + 1
    rc = _grid_row_count(new_dm)
    _trace(u"grid rebuild: want=%s rc=%s" % (len(rows), rc))
    try:
        gm = grid.getModel()
        gm.GridDataModel = new_dm
        if state is not None:
            state[u"grid_data_model"] = new_dm
    except Exception as exc:
        _trace(u"grid rebuild assign: %s" % exc)
        return
    if keep_index >= 0 and keep_index < len(rows):
        try:
            _grid_select_row(grid, keep_index)
        except Exception as exc:
            _log("grid select: %s" % exc)
    try:
        toolkit = _get_toolkit()
        if toolkit is not None:
            toolkit.processEventsToIdle()
    except Exception:
        pass


def _grid_sync(grid, rows, doc=None, keep_index=-1, data_model=None, state=None):
    if state is None:
        _trace(u"grid_sync: no state")
        return
    si = int(keep_index)
    if si < 0:
        si = int(state.get(u"selected") or 0)
    _grid_refresh_rows(state, keep_index=si)


def _grid_preload(model, rows, doc=None):
    """Заполнить GridDataModel до createPeer (надёжнее, чем addRow после peer)."""
    try:
        data_model = model.GridDataModel
    except Exception:
        _trace(u"grid_preload: no GridDataModel")
        return
    _grid_clear_data_model(data_model)
    _trace(u"grid_preload start rows=%s" % len(rows))
    _grid_fill_data_model(data_model, rows, doc=doc)
    _trace(u"grid_preload: done rc=%s" % _grid_row_count(data_model))


def _grid_selected_index(grid):
    data_model = _grid_data_model(grid)
    try:
        sel = grid.getSelectedRows()
        if sel and len(sel) > 0:
            raw = sel[0]
            try:
                idx = int(raw)
                if data_model is not None:
                    rc = _grid_row_count(data_model)
                    if 0 <= idx < rc:
                        return idx
                elif idx >= 0:
                    return idx
            except (TypeError, ValueError):
                pass
            if data_model is not None:
                return _grid_index_from_heading(data_model, raw)
    except Exception:
        pass
    return -1


def _parse_date_field(text):
    from todo_task_plan_lib import parse_plan_date_parts

    return parse_plan_date_parts(text)


def _format_date_field_on_blur(dialog, dm, edit_name, wd_lbl_name, doc=None, snap=True):
    from todo_task_plan_lib import (
        _date_to_parts,
        _parts_to_date,
        format_plan_date_field,
        load_holidays,
        parse_plan_date_parts,
        plan_date_weekday_label,
        snap_workday_left,
    )

    parts = parse_plan_date_parts(_get_control_text(dialog, dm, edit_name))
    if parts and snap:
        ds = _parts_to_date(parts)
        if ds:
            ds = snap_workday_left(ds, load_holidays(doc))
            parts = _date_to_parts(ds)
    if parts:
        _set_control_text(dialog, dm, edit_name, format_plan_date_field(parts))
        _set_fixed_label(dialog, dm, wd_lbl_name, plan_date_weekday_label(parts))
    else:
        _set_fixed_label(dialog, dm, wd_lbl_name, u"")


def _row_edit_set_err(dialog, dm, msg):
    t = unicode(msg or u"")
    try:
        dialog.getControl(u"ErrLbl").setLabel(t)
    except Exception:
        pass
    try:
        dm.getByName(u"ErrLbl").Label = t
    except Exception:
        pass


def _validate_plan_row_dates(dialog, dm, doc=None, update_duration=True):
    """Проверка дат строки; при успехе — пересчёт длительности."""
    from todo_task_plan_lib import (
        _parts_to_date,
        load_holidays,
        workdays_between,
    )

    sp = _parse_date_field(_get_control_text(dialog, dm, u"StartEd"))
    ep = _parse_date_field(_get_control_text(dialog, dm, u"EndEd"))
    if sp and not ep:
        return False, u"Укажите дату окончания.", sp, ep, None
    if ep and not sp:
        return False, u"Укажите дату начала.", sp, ep, None
    if sp and ep:
        ds = _parts_to_date(sp)
        de = _parts_to_date(ep)
        if not ds or not de:
            return False, u"Некорректный формат даты.", sp, ep, None
        if de < ds:
            if update_duration:
                _set_control_text(dialog, dm, u"DurEd", u"")
            return (
                False,
                u"Дата окончания не может быть раньше даты начала.",
                sp,
                ep,
                None,
            )
        wd = workdays_between(ds, de, load_holidays(doc))
        if update_duration:
            _set_control_text(dialog, dm, u"DurEd", unicode(wd))
        return True, u"", sp, ep, wd
    if update_duration:
        _set_control_text(dialog, dm, u"DurEd", u"")
    return True, u"", sp, ep, None


def _row_edit_refresh_duration(dialog, dm, doc=None):
    ok, err, sp, ep, wd = _validate_plan_row_dates(
        dialog, dm, doc=doc, update_duration=True
    )
    _row_edit_set_err(dialog, dm, err if not ok else u"")
    if not ok:
        return False, err
    return True, u""


def _row_edit_refresh_dates(dialog, dm, doc=None):
    _format_date_field_on_blur(dialog, dm, u"StartEd", u"StartWdLbl", doc=doc)
    _format_date_field_on_blur(dialog, dm, u"EndEd", u"EndWdLbl", doc=doc)
    return _row_edit_refresh_duration(dialog, dm, doc=doc)


def _show_plan_row_edit_dialog(doc, row, assignees=None, criticalities=None, parent_dialog=None):
    """Диалог правки одной строки плана. Возвращает dict или None."""
    row = dict(row or {})
    assignees = list(assignees or [])
    criticalities = list(criticalities or [])

    ctx = _uno_context()
    if ctx is None:
        return None
    toolkit = _get_toolkit()
    if toolkit is None:
        return None

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = 10
    gap = 6
    row_h = 24
    lbl_w = 118
    dlg_w = 520
    cw = dlg_w - 2 * m
    v_mid = 1

    from todo_task_plan_lib import format_plan_date_field, plan_date_weekday_label

    sp = row.get(u"date_start_parts")
    ep = row.get(u"date_end_parts")
    title = u"Строка плана %s" % (unicode(row.get(u"item_no") or u"").strip() or u"—")
    dm.Title = title
    dm.Width = dlg_w
    dm.PositionX = 120
    dm.PositionY = 100

    date_w = 132
    wd_w = 40

    y = m
    _add_label(dm, u"ItemLbl", u"Пункт:", m, y, lbl_w, row_h, vertical_align=v_mid)
    _add_edit(dm, u"ItemEd", m + lbl_w + 4, y, 72, row_h, multiline=False)
    _add_checkbox(dm, u"SumChk", u"Итоговый пункт", m + lbl_w + 84, y, 150, row_h)
    y += row_h + gap

    _add_label(
        dm, u"StartLbl", u"Начало (ДД.ММ.ГГГГ):", m, y, lbl_w, row_h, vertical_align=v_mid
    )
    _add_edit(dm, u"StartEd", m + lbl_w + 4, y, date_w, row_h, multiline=False)
    _add_label(
        dm,
        u"StartWdLbl",
        u"",
        m + lbl_w + 4 + date_w + 6,
        y,
        wd_w,
        row_h,
        vertical_align=v_mid,
    )
    y += row_h + gap

    _add_label(dm, u"EndLbl", u"Окончание:", m, y, lbl_w, row_h, vertical_align=v_mid)
    _add_edit(dm, u"EndEd", m + lbl_w + 4, y, date_w, row_h, multiline=False)
    _add_label(
        dm,
        u"EndWdLbl",
        u"",
        m + lbl_w + 4 + date_w + 6,
        y,
        wd_w,
        row_h,
        vertical_align=v_mid,
    )
    y += row_h + gap

    _add_label(
        dm, u"DurLbl", u"Длительность, раб. дни:", m, y, lbl_w, row_h, vertical_align=v_mid
    )
    _add_edit(dm, u"DurEd", m + lbl_w + 4, y, 56, row_h, multiline=False)
    y += row_h + gap

    _add_label(
        dm, u"AsgLbl", u"Ответственные:", m, y, lbl_w, row_h, vertical_align=v_mid
    )
    _add_combo(dm, u"AsgCb", m + lbl_w + 4, y, cw - lbl_w - 4, row_h, assignees)
    y += row_h + gap

    _add_label(
        dm, u"ResLbl", u"Результат / документ:", m, y, lbl_w, row_h, vertical_align=v_mid
    )
    _add_edit(dm, u"ResEd", m + lbl_w + 4, y, cw - lbl_w - 4, row_h, multiline=False)
    y += row_h + gap

    _add_label(
        dm, u"CritLbl", u"Критичность:", m, y, lbl_w, row_h, vertical_align=v_mid
    )
    _add_combo(dm, u"CritCb", m + lbl_w + 4, y, 180, row_h, criticalities)
    y += row_h + gap

    for lbl_name in (
        u"ItemLbl",
        u"StartLbl",
        u"EndLbl",
        u"DurLbl",
        u"AsgLbl",
        u"ResLbl",
        u"CritLbl",
        u"StartWdLbl",
        u"EndWdLbl",
    ):
        try:
            dm.getByName(lbl_name).MultiLine = False
        except Exception:
            pass

    _add_label(dm, u"ErrLbl", u"", m, y, cw, 16, vertical_align=v_mid)
    y += 18

    btn_w = 90
    _add_button(dm, u"OkBtn", u"OK", dlg_w - m - 2 * btn_w - gap, y, btn_w, row_h, default=True)
    _add_button(dm, u"CancelBtn", u"Отмена", dlg_w - m - btn_w, y, btn_w, row_h)
    dm.Height = y + row_h + m

    try:
        _theme_todo_dialog(dm, title_text=title, doc=doc)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc, parent_dialog=parent_dialog):
        return None

    _set_control_text(dialog, dm, "ItemEd", unicode(row.get(u"item_no") or u""))
    try:
        dialog.getControl("SumChk").setState(1 if row.get(u"is_summary") else 0)
    except Exception:
        pass
    _set_control_text(
        dialog,
        dm,
        "StartEd",
        format_plan_date_field(sp) if sp else u"",
    )
    _set_control_text(
        dialog,
        dm,
        "EndEd",
        format_plan_date_field(ep) if ep else u"",
    )
    try:
        _set_fixed_label(
            dialog, dm, "StartWdLbl", plan_date_weekday_label(sp) if sp else u""
        )
        _set_fixed_label(
            dialog, dm, "EndWdLbl", plan_date_weekday_label(ep) if ep else u""
        )
    except Exception:
        pass
    _set_control_text(dialog, dm, "DurEd", unicode(row.get(u"duration_wd") or u""))
    _set_control_text(dialog, dm, "AsgCb", unicode(row.get(u"assignees") or u""))
    _set_control_text(dialog, dm, "ResEd", unicode(row.get(u"result_doc") or u""))
    _set_control_text(dialog, dm, "CritCb", unicode(row.get(u"criticality") or u""))

    state = {u"ok": False, u"row": None}

    class _H(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "OkBtn":
                ok_dates, err, sp, ep, wd = _validate_plan_row_dates(
                    dialog, dm, doc=doc, update_duration=True
                )
                if not ok_dates:
                    _row_edit_set_err(dialog, dm, err)
                    _dialog_warn(doc, err)
                    return
                _row_edit_set_err(dialog, dm, u"")
                out = dict(row)
                out[u"item_no"] = _get_control_text(dialog, dm, u"ItemEd").strip()
                try:
                    out[u"is_summary"] = bool(dialog.getControl(u"SumChk").getState())
                except Exception:
                    out[u"is_summary"] = False
                out[u"date_start_parts"] = sp
                out[u"date_end_parts"] = ep
                if sp and ep and wd is not None:
                    out[u"duration_wd"] = int(wd)
                else:
                    try:
                        out[u"duration_wd"] = int(
                            float(
                                _get_control_text(dialog, dm, u"DurEd").strip() or u"0"
                            )
                        )
                    except (TypeError, ValueError):
                        out[u"duration_wd"] = 0
                out[u"assignees"] = _get_control_text(dialog, dm, "AsgCb").strip()
                out[u"result_doc"] = _get_control_text(dialog, dm, "ResEd")
                out[u"criticality"] = _get_control_text(dialog, dm, "CritCb").strip()
                from todo_task_plan_lib import item_no_to_sort_key, parent_of_item_no

                out[u"parent_item"] = parent_of_item_no(out[u"item_no"])
                out[u"sort_key"] = item_no_to_sort_key(out[u"item_no"])
                state[u"row"] = out
                state[u"ok"] = True
            elif name != "CancelBtn":
                return
            try:
                dialog.endExecute()
            except Exception:
                pass

    h = _H()

    class _DateFocus(unohelper.Base, XFocusListener):
        def __init__(self, edit_name, wd_lbl_name):
            self.edit_name = edit_name
            self.wd_lbl_name = wd_lbl_name

        def disposing(self, event):
            pass

        def focusGained(self, event):
            pass

        def focusLost(self, event):
            _row_edit_refresh_dates(dialog, dm, doc=doc)

    for edit_name, wd_name in (("StartEd", "StartWdLbl"), ("EndEd", "EndWdLbl")):
        try:
            dialog.getControl(edit_name).addFocusListener(
                _DateFocus(edit_name, wd_name)
            )
        except Exception:
            pass

    for bn in (u"OkBtn", u"CancelBtn"):
        try:
            dialog.getControl(bn).addActionListener(h)
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
    return state[u"row"]


def _recalc(state, save_panel=False):
    from todo_task_plan_lib import load_holidays, recalc_plan_rows

    if save_panel and not state.get(u"use_grid"):
        _save_current_row_from_panel(state, validate_dates=False)
    si = _selected_index(state)
    rows = state.get(u"rows") or []
    row_id = None
    if 0 <= si < len(rows):
        row_id = rows[si].get(u"row_id")
    state[u"rows"] = recalc_plan_rows(
        rows, holidays=load_holidays(state.get(u"doc"))
    )
    rows = state[u"rows"]
    if row_id:
        ri = 0
        while ri < len(rows):
            if rows[ri].get(u"row_id") == row_id:
                si = ri
                break
            ri = ri + 1
    _refresh_view(state, keep_index=si)
    if 0 <= si < len(rows):
        state[u"selected"] = si


def _panel_refresh_dates(state):
    dialog = state.get(u"dialog")
    dm = state.get(u"dm")
    doc = state.get(u"doc")
    start_snap = not _checkbox_state(dialog, u"StartNoSnapChk")
    end_snap = not _checkbox_state(dialog, u"EndNoSnapChk")
    _format_date_field_on_blur(
        dialog, dm, u"StartEd", u"StartWdLbl", doc=doc, snap=start_snap
    )
    _format_date_field_on_blur(
        dialog, dm, u"EndEd", u"EndWdLbl", doc=doc, snap=end_snap
    )
    ok, err, sp, ep, wd = _validate_plan_row_dates(
        dialog, dm, doc=doc, update_duration=True
    )
    _detail_set_err(dialog, dm, err if not ok else u"")
    return ok, err


def _panel_cascade_duration(state):
    from todo_task_plan_lib import cascade_plan_from_row, load_holidays

    si = _selected_index(state)
    rows = state.get(u"rows") or []
    row_id = rows[si].get(u"row_id") if 0 <= si < len(rows) else None
    if not _save_panel_to_row(
        state, si, validate_dates=False, prefer_duration_field=True
    ):
        return
    rows = state.get(u"rows") or []
    rows = cascade_plan_from_row(
        rows,
        si,
        holidays=load_holidays(state.get(u"doc")),
        row_id=row_id,
    )
    state[u"rows"] = rows
    if row_id:
        ri = 0
        while ri < len(rows):
            if rows[ri].get(u"row_id") == row_id:
                si = ri
                break
            ri = ri + 1
    state[u"selected"] = si
    _list_refresh(state, keep_index=si)


def _panel_cascade_end_date(state):
    from todo_task_plan_lib import cascade_plan_from_end_change, load_holidays

    si = _selected_index(state)
    rows = state.get(u"rows") or []
    row_id = rows[si].get(u"row_id") if 0 <= si < len(rows) else None
    if not _save_panel_to_row(
        state, si, validate_dates=False, prefer_duration_field=False
    ):
        return
    rows = state.get(u"rows") or []
    rows = cascade_plan_from_end_change(
        rows,
        si,
        holidays=load_holidays(state.get(u"doc")),
        row_id=row_id,
    )
    state[u"rows"] = rows
    if row_id:
        ri = 0
        while ri < len(rows):
            if rows[ri].get(u"row_id") == row_id:
                si = ri
                break
            ri = ri + 1
    state[u"selected"] = si
    _list_refresh(state, keep_index=si)


def _update_status(state):
    rows = state.get(u"rows") or []
    si = int(state.get(u"selected") or 0)
    holi = u""
    try:
        from todo_task_plan_lib import holidays_ref_info

        has_ref, n_holi = holidays_ref_info(state.get(u"doc"))
        if has_ref:
            holi = u"; праздников в справочнике: %s" % n_holi
        else:
            holi = u"; справочник «Праздники» не найден"
    except Exception:
        pass
    try:
        state[u"dialog"].getControl("StatusLbl").setLabel(
            u"Строк в плане: %s; выбрана: %s%s"
            % (len(rows), si + 1 if rows and 0 <= si < len(rows) else u"—", holi)
        )
    except Exception:
        pass


def _selected_index(state):
    return int(state.get(u"selected") or 0)


def _sync_selection_from_grid(state):
    idx = _grid_selected_index(state[u"grid"])
    if idx >= 0:
        state[u"selected"] = idx
    _update_status(state)


def _add_root(state):
    if not state.get(u"use_grid"):
        _save_current_row_from_panel(state, validate_dates=False)
    from datetime import date

    from todo_task_lib import default_due_parts
    from todo_task_plan_lib import (
        add_workdays_inclusive,
        item_no_to_sort_key,
        load_holidays,
        next_root_item_no,
        next_workday_after,
        parent_of_item_no,
        snap_workday_left,
        _date_to_parts,
        _new_guid,
        _parts_to_date,
    )

    rows = state[u"rows"]
    holidays = load_holidays(state.get(u"doc"))
    ino = next_root_item_no(rows)
    ds = None
    if rows:
        ep = _parts_to_date(rows[-1].get(u"date_end_parts"))
        if ep:
            ds = next_workday_after(ep, holidays)
    if ds is None:
        parts = default_due_parts()
        try:
            ds = date(parts[2], parts[1], parts[0])
        except Exception:
            ds = date.today()
        ds = snap_workday_left(ds, holidays)
    dur = int(_pcfg.PLAN_DEFAULT_DURATION)
    de = add_workdays_inclusive(ds, dur, holidays)
    rows.append(
        {
            u"guid": state.get(u"guid") or u"",
            u"item_no": ino,
            u"is_summary": True,
            u"parent_item": parent_of_item_no(ino),
            u"sort_key": item_no_to_sort_key(ino),
            u"date_start_parts": _date_to_parts(ds),
            u"date_end_parts": _date_to_parts(de),
            u"duration_wd": dur,
            u"assignees": u"",
            u"stage_name": u"",
            u"result_doc": u"",
            u"criticality": u"",
            u"row_id": _new_guid(),
        }
    )
    state[u"rows"] = rows
    state[u"selected"] = len(rows) - 1
    _recalc(state)


def _add_child(state):
    _sync_selection(state)
    if not state.get(u"use_grid"):
        _save_current_row_from_panel(state, validate_dates=False)
    from todo_task_plan_lib import (
        item_no_to_sort_key,
        next_child_item_no,
        parent_of_item_no,
        plan_root_item_no,
        _new_guid,
    )

    rows = state[u"rows"]
    si = _selected_index(state)
    if si < 0 or si >= len(rows):
        return
    selected_ino = unicode(rows[si].get(u"item_no") or u"").strip()
    if not selected_ino:
        return
    root_ino = plan_root_item_no(selected_ino)
    ri = 0
    while ri < len(rows):
        if unicode(rows[ri].get(u"item_no") or u"").strip() == root_ino:
            rows[ri][u"is_summary"] = True
            break
        ri = ri + 1
    ino = next_child_item_no(rows, root_ino)
    rows.append(
        {
            u"guid": state.get(u"guid") or u"",
            u"item_no": ino,
            u"is_summary": False,
            u"parent_item": root_ino,
            u"sort_key": item_no_to_sort_key(ino),
            u"date_start_parts": None,
            u"date_end_parts": None,
            u"duration_wd": int(_pcfg.PLAN_DEFAULT_DURATION),
            u"assignees": u"",
            u"stage_name": u"",
            u"result_doc": u"",
            u"criticality": u"",
            u"row_id": _new_guid(),
        }
    )
    state[u"rows"] = rows
    state[u"selected"] = len(rows) - 1
    _recalc(state)


def _delete_selected(state):
    _sync_selection(state)
    if not state.get(u"use_grid"):
        _save_current_row_from_panel(state, validate_dates=False)
    rows = state[u"rows"]
    si = _selected_index(state)
    if si < 0 or si >= len(rows):
        return
    from todo_task_plan_lib import _plan_delete_targets

    selected_ino = unicode(rows[si].get(u"item_no") or u"").strip()
    kill = _plan_delete_targets(rows, selected_ino)
    to_del = []
    i = 0
    while i < len(rows):
        ino = unicode(rows[i].get(u"item_no") or u"").strip()
        if ino in kill:
            to_del.append(i)
        i = i + 1
    to_del.sort(reverse=True)
    j = 0
    while j < len(to_del):
        del rows[to_del[j]]
        j = j + 1
    state[u"rows"] = rows
    if rows:
        state[u"selected"] = min(si, len(rows) - 1)
    else:
        state[u"selected"] = 0
    _recalc(state)


def _move_selected(state, delta):
    _sync_selection(state)
    if not state.get(u"use_grid"):
        _save_current_row_from_panel(state, validate_dates=False)
    rows = state[u"rows"]
    si = _selected_index(state)
    ni = si + int(delta)
    if si < 0 or si >= len(rows) or ni < 0 or ni >= len(rows):
        return
    tmp = rows[si]
    rows[si] = rows[ni]
    rows[ni] = tmp
    state[u"selected"] = ni
    state[u"rows"] = rows
    _recalc(state)


def _edit_selected(state):
    _sync_selection(state)
    si = _selected_index(state)
    rows = state[u"rows"]
    if si < 0 or si >= len(rows):
        _dialog_warn(state.get(u"doc"), u"Выберите строку в таблице (клик по строке).")
        return
    prev = dict(rows[si])
    edited = _show_plan_row_edit_dialog(
        state.get(u"doc"),
        rows[si],
        assignees=state.get(u"assignees"),
        criticalities=state.get(u"criticalities"),
        parent_dialog=state.get(u"dialog"),
    )
    if edited is None:
        return
    sp = edited.get(u"date_start_parts")
    ep = edited.get(u"date_end_parts")
    if sp and ep:
        from todo_task_plan_lib import _parts_to_date

        ds = _parts_to_date(sp)
        de = _parts_to_date(ep)
        if ds and de and de < ds:
            _dialog_warn(
                state.get(u"doc"),
                u"Дата окончания не может быть раньше даты начала.",
            )
            return
    if prev.get(u"row_id") and not edited.get(u"row_id"):
        edited[u"row_id"] = prev.get(u"row_id")
    if prev.get(u"guid") and not edited.get(u"guid"):
        edited[u"guid"] = prev.get(u"guid")
    rows[si] = edited
    state[u"rows"] = rows
    state[u"selected"] = si
    _recalc(state)


def show_todo_task_plan_wizard( doc=None, guid=u"", task_name=u"", sheet=None, cols=None, row=None, assignees=None, criticalities=None, parent_dialog=None, read_only=False):
    """Полноэкранный визард плана. По умолчанию — split ListBox; Grid — PLAN_WIZARD_USE_GRID."""
    if getattr(_pcfg, "PLAN_WIZARD_USE_GRID", False):
        return _show_plan_wizard_grid(
            doc=doc,
            guid=guid,
            task_name=task_name,
            sheet=sheet,
            cols=cols,
            row=row,
            assignees=assignees,
            criticalities=criticalities,
            parent_dialog=parent_dialog,
            read_only=read_only,
        )
    return _show_plan_wizard_list(
        doc=doc,
        guid=guid,
        task_name=task_name,
        sheet=sheet,
        cols=cols,
        row=row,
        assignees=assignees,
        criticalities=criticalities,
        parent_dialog=parent_dialog,
        read_only=read_only,
    )


def _show_plan_wizard_list( doc=None, guid=u"", task_name=u"", sheet=None, cols=None, row=None, assignees=None, criticalities=None, parent_dialog=None, read_only=False):
    """Split-форма: ListBox слева, поля строки справа."""
    _trace(u"=== show_todo_task_plan_wizard (list) START parent=%s ===" % (
        u"yes" if parent_dialog is not None else u"no",
    ))
    from todo_task_plan_lib import (
        apply_plan_save_to_task,
        default_plan_rows,
        ensure_plan_ref_blocks,
        load_plan_by_guid,
        save_plan_by_guid,
    )

    guid = unicode(guid or u"").strip()
    if not guid:
        _dialog_warn(doc, u"Нет GUID задачи — план недоступен.")
        return False

    from todo_task_plan_lib import plans_workbook_path
    import todo_task_plan_cfg as _pcfg

    if not plans_workbook_path(doc):
        _dialog_warn(
            doc,
            unicode(
                getattr(
                    _pcfg,
                    "PLAN_STORE_DOC_UNSAVED",
                    u"Сохраните книгу задач на диск.",
                )
            ),
        )
        return False

    ensure_plan_ref_blocks(doc)
    rows = load_plan_by_guid(doc, guid)
    if not rows and not read_only:
        rows = default_plan_rows(doc=doc, guid=guid)
    elif not rows:
        rows = []

    if assignees is None:
        assignees = []
    if criticalities is None:
        from todo_task_plan_lib import load_criticality_choices

        criticalities = load_criticality_choices(doc)

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sx, sy, sw, sh = _screen_work_area(toolkit)
    title_reserve = _plan_wizard_title_h()
    dlg_w = max(640, int(sw))
    screen_h = max(480, int(sh))
    dlg_h = max(400, screen_h - title_reserve)

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialogModel", ctx)
    m = 8
    gap = 4
    toolbar_h = int(getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_BTN_H", 28))
    status_h = 18
    hint_h = int(getattr(_pcfg, "PLAN_WIZARD_HINT_H", 44))
    footer_btn_w = int(getattr(_pcfg, "PLAN_WIZARD_FOOTER_BTN_W", 110))
    split_gap = int(getattr(_pcfg, "PLAN_WIZARD_SPLIT_GAP", 8))
    list_lbl_h = int(getattr(_pcfg, "PLAN_WIZARD_LIST_LBL_H", 18))

    title = u"%s: %s" % (
        unicode(_pcfg.PLAN_WIZARD_TITLE),
        unicode(task_name or guid)[:80],
    )
    if read_only:
        title = title + unicode(
            getattr(_pcfg, "PLAN_WIZARD_VIEW_SUFFIX", u" (просмотр)")
        )
    dm.Title = title
    dm.Width = dlg_w
    dm.Height = dlg_h
    dm.PositionX = int(sx)
    dm.PositionY = int(sy)
    try:
        dm.Sizeable = True
    except Exception:
        pass

    y = m
    _add_label(dm, u"HintLbl", unicode(_pcfg.PLAN_WIZARD_HINT), m, y, dlg_w - 2 * m, hint_h)
    try:
        dm.getByName(u"HintLbl").MultiLine = True
    except Exception:
        pass
    y += hint_h + gap

    toolbar_y = y
    tb_specs = getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_BTNS_LIST", ())
    x = m
    for spec in tb_specs:
        _add_button(dm, spec[0], spec[1], x, toolbar_y, int(spec[2]), toolbar_h)
        x += int(spec[2]) + 4
    _add_button(
        dm, u"SaveBtn", u"Сохранить",
        dlg_w - m - 2 * footer_btn_w - gap, toolbar_y, footer_btn_w, toolbar_h, default=True,
    )
    _add_button(
        dm, u"CancelBtn", u"Отмена",
        dlg_w - m - footer_btn_w, toolbar_y, footer_btn_w, toolbar_h,
    )
    y += toolbar_h + gap

    half = max(int(getattr(_pcfg, "PLAN_WIZARD_LIST_MIN_W", 280)), (dlg_w - 2 * m - split_gap) // 2)
    right_x = m + half + split_gap
    right_w = dlg_w - right_x - m
    content_h = _plan_wizard_content_height(screen_h, dm)
    try:
        dm.Height = content_h
    except Exception:
        pass
    status_y = content_h - m - status_h
    list_top = y + list_lbl_h + 2
    list_h = _plan_list_box_height(list_top, status_y, gap=gap)

    _add_label(dm, u"PlanListLbl", u"Пункты плана:", m, y, half, list_lbl_h)
    _add_plan_listbox(dm, u"PlanListBox", m, y + list_lbl_h + 2, half, list_h)
    _create_plan_detail_panel(dm, right_x, y, right_w, assignees, criticalities)
    _add_label(dm, u"StatusLbl", u"", m, status_y, dlg_w - 2 * m, status_h)

    try:
        _theme_todo_dialog(dm, title_text=title, doc=doc)
    except Exception:
        pass
    _relayout_plan_wizard_split(
        dm, dlg_w, screen_h, m, gap, toolbar_h, status_h, hint_h=hint_h
    )

    dialog = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _plan_wizard_create_peer(
        dialog, toolkit, doc=doc, parent_dialog=parent_dialog
    ):
        if not _create_peer(dialog, toolkit, doc=doc, parent_dialog=parent_dialog):
            raise RuntimeError(u"Не удалось создать peer диалога плана.")
    _relayout_plan_wizard_split(
        dm, dlg_w, screen_h, m, gap, toolbar_h, status_h, hint_h=hint_h
    )
    _sync_plan_wizard_control_layout(dialog, dm)
    try:
        from com.sun.star.awt.PosSize import POSSIZE
        lb = dialog.getControl(u"PlanListBox")
        mdl = dm.getByName(u"PlanListBox")
        lb.setPosSize(
            int(mdl.PositionX),
            int(mdl.PositionY),
            int(mdl.Width),
            int(mdl.Height),
            POSSIZE,
        )
    except Exception:
        pass
    try:
        try_paint_titlebar(
            dialog, title_bg=DARK_BLUE_TITLE_BG, title_fg=DARK_BLUE_TITLE_FG
        )
    except Exception:
        pass

    state = {
        u"dialog": dialog,
        u"dm": dm,
        u"doc": doc,
        u"guid": guid,
        u"rows": rows,
        u"selected": 0,
        u"saved": False,
        u"plan_sync": None,
        u"sheet": sheet,
        u"cols": cols,
        u"row": row,
        u"use_grid": False,
        u"form_loading": False,
        u"assignees": assignees,
        u"criticalities": criticalities,
        u"read_only": bool(read_only),
    }

    class _ListHandler(unohelper.Base, XActionListener, XItemListener, XFocusListener, XTextListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == u"CancelBtn":
                try:
                    dialog.endExecute()
                except Exception:
                    pass
                return
            if name == u"SaveBtn":
                if state.get(u"read_only"):
                    return
                if not _save_current_row_from_panel(state, validate_dates=True):
                    return
                if not save_plan_by_guid(doc, guid, state[u"rows"]):
                    from todo_task_plan_lib import get_last_plan_io_error

                    err = get_last_plan_io_error() or u"Не удалось сохранить план."
                    _dialog_warn(doc, err)
                    return
                if sheet is not None and cols is not None and row is not None:
                    state[u"plan_sync"] = apply_plan_save_to_task(
                        doc, sheet, cols, int(row), load_plan_by_guid(doc, guid)
                    )
                state[u"saved"] = True
                try:
                    dialog.endExecute()
                except Exception:
                    pass
            elif name == u"AddRootBtn":
                _add_root(state)
            elif name == u"AddChildBtn":
                _add_child(state)
            elif name == u"DelBtn":
                _delete_selected(state)
            elif name == u"UpBtn":
                _move_selected(state, -1)
            elif name == u"DownBtn":
                _move_selected(state, 1)
            elif name == u"RecalcBtn":
                _save_current_row_from_panel(state, validate_dates=False)
                _recalc(state)

        def itemStateChanged(self, event):
            if state.get(u"form_loading"):
                return
            try:
                src_name = str(event.Source.getModel().Name)
            except Exception:
                src_name = u""
            if src_name == u"DoneChk":
                _handle_done_checkbox_toggle(state)
                return
            if src_name in (u"StartNoSnapChk", u"EndNoSnapChk", u"SumChk"):
                _save_panel_to_row(state, _selected_index(state), validate_dates=False)
                return
            prev = _selected_index(state)
            if prev >= 0:
                _save_panel_to_row(state, prev, validate_dates=False)
            idx = _list_selected_index(dialog)
            if idx >= 0:
                state[u"selected"] = idx
                state[u"form_loading"] = True
                try:
                    _load_row_to_panel(state, idx)
                finally:
                    state[u"form_loading"] = False
            _update_status(state)

        def focusGained(self, event):
            pass

        def focusLost(self, event):
            if state.get(u"form_loading"):
                return
            try:
                mname = str(event.Source.getModel().Name)
            except Exception:
                return
            rows = state.get(u"rows") or []
            si = _selected_index(state)
            if 0 <= si < len(rows) and rows[si].get(u"is_done"):
                return
            if mname == u"DurEd":
                _panel_cascade_duration(state)
                return
            if mname not in (u"StartEd", u"EndEd"):
                return
            start_snap = not _checkbox_state(dialog, u"StartNoSnapChk")
            end_snap = not _checkbox_state(dialog, u"EndNoSnapChk")
            _format_date_field_on_blur(
                dialog, dm, u"StartEd", u"StartWdLbl", doc=doc, snap=start_snap
            )
            _format_date_field_on_blur(
                dialog, dm, u"EndEd", u"EndWdLbl", doc=doc, snap=end_snap
            )
            ok, err, sp, ep, wd = _validate_plan_row_dates(
                dialog, dm, doc=doc, update_duration=True
            )
            _detail_set_err(dialog, dm, err if not ok else u"")
            if ok:
                if mname == u"EndEd":
                    _panel_cascade_end_date(state)
                else:
                    _save_panel_to_row(state, _selected_index(state), validate_dates=False)

        def textChanged(self, event):
            if state.get(u"form_loading"):
                return
            try:
                mname = str(event.Source.getModel().Name)
            except Exception:
                mname = u""
            rows = state.get(u"rows") or []
            si = _selected_index(state)
            if (
                mname != u"ResultDocEd"
                and 0 <= si < len(rows)
                and rows[si].get(u"is_done")
            ):
                return
            _save_panel_to_row(state, si, validate_dates=False)

    handler = _ListHandler()
    for btn in (
        u"SaveBtn",
        u"CancelBtn",
        u"AddRootBtn",
        u"AddChildBtn",
        u"DelBtn",
        u"UpBtn",
        u"DownBtn",
        u"RecalcBtn",
    ):
        try:
            dialog.getControl(btn).addActionListener(handler)
        except Exception as err:
            _log("listener %s: %s" % (btn, err))
    try:
        dialog.getControl(u"PlanListBox").addItemListener(handler)
    except Exception as err:
        _log("list listener: %s" % err)
    for edit_name in (u"StartEd", u"EndEd", u"DurEd"):
        try:
            dialog.getControl(edit_name).addFocusListener(handler)
        except Exception:
            pass
    for chk_name in (u"SumChk", u"StartNoSnapChk", u"EndNoSnapChk", u"DoneChk"):
        try:
            dialog.getControl(chk_name).addItemListener(handler)
        except Exception:
            pass
    for edit_name in (u"StageEd", u"ResultDocEd", u"ItemEd"):
        try:
            dialog.getControl(edit_name).addTextListener(handler)
        except Exception:
            pass

    _list_refresh(state, keep_index=0)
    _apply_plan_wizard_read_only(state)
    try:
        toolkit.processEventsToIdle()
    except Exception:
        pass

    try:
        dialog.execute()
    finally:
        try:
            dialog.dispose()
        except Exception:
            pass

    _trace(u"=== show_todo_task_plan_wizard (list) END saved=%s ===" % bool(state.get(u"saved")))
    return bool(state.get(u"saved")), state.get(u"plan_sync") or {}


def _show_plan_wizard_grid( doc=None, guid=u"", task_name=u"", sheet=None, cols=None, row=None, assignees=None, criticalities=None, parent_dialog=None, read_only=False):
    """
    Legacy Grid-визард (PLAN_WIZARD_USE_GRID=True). По умолчанию не используется.
    """
    _trace(u"=== show_todo_task_plan_wizard (grid) START parent=%s ===" % (
        u"yes" if parent_dialog is not None else u"no",
    ))
    from todo_task_plan_lib import (
        apply_plan_save_to_task,
        default_plan_rows,
        ensure_plan_ref_blocks,
        load_plan_by_guid,
        save_plan_by_guid,
    )

    guid = unicode(guid or u"").strip()
    if not guid:
        _trace(u"ABORT: нет GUID")
        _dialog_warn(doc, u"Нет GUID задачи — план недоступен.")
        return False

    from todo_task_plan_lib import plans_workbook_path
    import todo_task_plan_cfg as _pcfg

    if not plans_workbook_path(doc):
        _dialog_warn(
            doc,
            unicode(
                getattr(
                    _pcfg,
                    "PLAN_STORE_DOC_UNSAVED",
                    u"Сохраните книгу задач на диск.",
                )
            ),
        )
        return False

    _trace(u"ensure_plan_ref_blocks…")
    ensure_plan_ref_blocks(doc)
    _trace(u"load_plan_by_guid…")
    rows = load_plan_by_guid(doc, guid)
    _trace(u"loaded rows: %s" % len(rows))
    if not rows:
        rows = default_plan_rows(doc=doc, guid=guid)
        _trace(u"default_plan_rows: 1")

    if assignees is None:
        assignees = []
    if criticalities is None:
        from todo_task_plan_lib import load_criticality_choices

        _trace(u"load_criticality_choices…")
        criticalities = load_criticality_choices(doc)

    _trace(u"uno context…")
    ctx = _uno_context()
    if ctx is None:
        _trace(u"ABORT: нет UNO context")
        raise RuntimeError(u"Нет UNO-контекста — запускайте макрос из Calc.")
    toolkit = _get_toolkit()
    if toolkit is None:
        _trace(u"ABORT: нет toolkit")
        raise RuntimeError(u"Не удалось получить awt.Toolkit.")

    sx, sy, sw, sh = _screen_work_area(toolkit)
    _trace(u"screen work area: %sx%s @ %s,%s" % (sw, sh, sx, sy))
    title_reserve = _plan_wizard_title_h()
    dlg_w = max(640, int(sw))
    screen_h = max(480, int(sh))
    dlg_h = max(400, screen_h - title_reserve)

    sm = ctx.getServiceManager()
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = 8
    gap = 4
    toolbar_h = int(getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_BTN_H", 28))
    footer_h = int(getattr(_pcfg, "PLAN_WIZARD_FOOTER_BTN_H", 28))
    btn_h = toolbar_h
    status_h = 18
    hint_h = int(getattr(_pcfg, "PLAN_WIZARD_HINT_H", 44))
    footer_btn_w = int(getattr(_pcfg, "PLAN_WIZARD_FOOTER_BTN_W", 110))

    title = u"%s: %s" % (
        unicode(_pcfg.PLAN_WIZARD_TITLE),
        unicode(task_name or guid)[:80],
    )
    dm.Title = title
    dm.Width = dlg_w
    dm.Height = dlg_h
    dm.PositionX = int(sx)
    dm.PositionY = int(sy)
    try:
        dm.Sizeable = True
    except Exception:
        pass

    y = m
    hint = unicode(_pcfg.PLAN_WIZARD_HINT)
    _add_label(dm, "HintLbl", hint, m, y, dlg_w - 2 * m, hint_h)
    try:
        dm.getByName("HintLbl").MultiLine = True
    except Exception:
        pass
    y += hint_h + gap

    tb_specs = getattr(_pcfg, "PLAN_WIZARD_TOOLBAR_BTNS", ())
    toolbar_y = y
    x = m
    for spec in tb_specs:
        name = spec[0]
        label = spec[1]
        bw = int(spec[2])
        _add_button(dm, name, label, x, toolbar_y, bw, toolbar_h)
        x += bw + 4
    _add_button(
        dm,
        "SaveBtn",
        u"Сохранить",
        dlg_w - m - 2 * footer_btn_w - gap,
        toolbar_y,
        footer_btn_w,
        toolbar_h,
        default=True,
    )
    _add_button(
        dm,
        "CancelBtn",
        u"Отмена",
        dlg_w - m - footer_btn_w,
        toolbar_y,
        footer_btn_w,
        toolbar_h,
    )
    y += toolbar_h + gap

    grid_top = y
    status_y = dlg_h - m - status_h
    grid_h = status_y - gap - grid_top
    if grid_h < 80:
        grid_h = 80

    grid_model = _build_plan_grid_model(
        dm, m, grid_top, dlg_w - 2 * m, grid_h, rows, doc=doc
    )
    grid_data_model = grid_model.GridDataModel
    dm.insertByName(u"PlanGrid", grid_model)

    _add_label(dm, "StatusLbl", u"", m, status_y, dlg_w - 2 * m, status_h)

    try:
        _theme_todo_dialog(dm, title_text=title, doc=doc)
    except Exception:
        pass
    _fix_plan_wizard_layout(
        dm, dlg_h, m, gap, btn_h, status_h, sw=dlg_w, sh=screen_h
    )
    _trace(u"dialog model ready %sx%s" % (dlg_w, screen_h))

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    _trace(u"createPeer (plan wizard fullscreen)…")
    if not _plan_wizard_create_peer(
        dialog, toolkit, doc=doc, parent_dialog=parent_dialog
    ):
        _trace(u"createPeer failed, fallback…")
        if not _create_peer(dialog, toolkit, doc=doc, parent_dialog=parent_dialog):
            _trace(u"ABORT: createPeer failed")
            raise RuntimeError(u"Не удалось создать peer диалога плана.")
    _trace(u"createPeer OK")
    _fix_plan_wizard_layout(
        dm, dlg_h, m, gap, btn_h, status_h, sw=dlg_w, sh=screen_h
    )
    _sync_plan_wizard_control_layout(dialog, dm)
    try:
        try_paint_titlebar(
            dialog, title_bg=DARK_BLUE_TITLE_BG, title_fg=DARK_BLUE_TITLE_FG
        )
    except Exception:
        pass

    grid = dialog.getControl("PlanGrid")
    state = {
        u"dialog": dialog,
        u"dm": dm,
        u"doc": doc,
        u"guid": guid,
        u"rows": rows,
        u"selected": 0,
        u"saved": False,
        u"plan_sync": None,
        u"sheet": sheet,
        u"cols": cols,
        u"row": row,
        u"use_grid": True,
        u"grid": grid,
        u"grid_data_model": grid_data_model,
        u"assignees": assignees,
        u"criticalities": criticalities,
    }

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "SaveBtn":
                if not save_plan_by_guid(doc, guid, state[u"rows"]):
                    from todo_task_plan_lib import get_last_plan_io_error

                    err = get_last_plan_io_error() or u"Не удалось сохранить план."
                    _dialog_warn(doc, err)
                    return
                if sheet is not None and cols is not None and row is not None:
                    fresh = load_plan_by_guid(doc, guid)
                    state[u"plan_sync"] = apply_plan_save_to_task(
                        doc, sheet, cols, int(row), fresh
                    )
                state[u"saved"] = True
                try:
                    dialog.endExecute()
                except Exception:
                    pass
            elif name == "CancelBtn":
                try:
                    dialog.endExecute()
                except Exception:
                    pass
            elif name == "AddRootBtn":
                _add_root(state)
            elif name == "AddChildBtn":
                _sync_selection(state)
                _add_child(state)
            elif name == "DelBtn":
                _delete_selected(state)
            elif name == "RefreshGridBtn":
                _grid_force_refresh(state)
            elif name == "UpBtn":
                _move_selected(state, -1)
            elif name == "DownBtn":
                _move_selected(state, 1)
            elif name == "RecalcBtn":
                if not state.get(u"use_grid"):
                    _save_current_row_from_panel(state, validate_dates=False)
                _recalc(state)
            elif name == "EditBtn":
                _sync_selection(state)
                _edit_selected(state)

    handler = _Handler()
    for btn in (
        u"SaveBtn",
        u"CancelBtn",
        u"AddRootBtn",
        u"AddChildBtn",
        u"DelBtn",
        u"RefreshGridBtn",
        u"UpBtn",
        u"DownBtn",
        u"EditBtn",
        u"RecalcBtn",
    ):
        try:
            dialog.getControl(btn).addActionListener(handler)
        except Exception as err:
            _log("listener %s: %s" % (btn, err))

    _attach_grid_mouse(state, grid)
    _update_status(state)
    try:
        toolkit.processEventsToIdle()
    except Exception:
        pass
    _trace(u"dialog.execute() — ожидание пользователя…")

    try:
        dialog.execute()
    finally:
        try:
            dialog.dispose()
        except Exception:
            pass

    _trace(u"=== show_todo_task_plan_wizard (grid) END saved=%s ===" % bool(state.get(u"saved")))
    return bool(state.get(u"saved")), state.get(u"plan_sync") or {}
