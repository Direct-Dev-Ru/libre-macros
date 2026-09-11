# -*- coding: utf-8 -*-
"""Субвизард условия выполнения шага (колонка D / E)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.725"
try:
    unicode
except NameError:
    unicode = str

import uno
import unohelper
from com.sun.star.awt import XActionListener, XItemListener

import libre_macros_step_condition_cfg as _cfg
import libre_macros_step_condition_lib as _lib


def _dlg_ctl(dlg, name):
    try:
        return dlg.getControl(name)
    except Exception:
        return None


def _set_text(ctl, text):
    if ctl is None:
        return
    try:
        ctl.setText(unicode(text or u""))
    except Exception:
        try:
            ctl.Text = unicode(text or u"")
        except Exception:
            pass


def _get_text(ctl):
    if ctl is None:
        return u""
    try:
        return unicode(ctl.getText() or u"")
    except Exception:
        try:
            return unicode(ctl.Text or u"")
        except Exception:
            return u""


def _set_visible(ctl, vis):
    if ctl is None:
        return
    try:
        ctl.setVisible(bool(vis))
    except Exception:
        try:
            ctl.Model.EnableVisible = bool(vis)
        except Exception:
            pass


def _combo_select_by_data(ctl, code, pairs):
    if ctl is None:
        return
    code_cf = unicode(code or u"").strip().casefold()
    idx = 0
    i = 0
    while i < len(pairs):
        if unicode(pairs[i][0]).casefold() == code_cf:
            idx = i
            break
        i += 1
    try:
        ctl.selectItemPos(idx, True)
    except Exception:
        try:
            ctl.setText(pairs[idx][1])
        except Exception:
            pass


def _combo_selected_code(ctl, pairs, default=u""):
    if ctl is None:
        return default
    try:
        pos = int(ctl.getSelectedItemPos())
    except Exception:
        pos = -1
    if 0 <= pos < len(pairs):
        return pairs[pos][0]
    # ComboBox: текст подписи
    txt = _get_text(ctl).strip().casefold()
    i = 0
    while i < len(pairs):
        if unicode(pairs[i][1]).casefold() == txt or unicode(pairs[i][0]).casefold() == txt:
            return pairs[i][0]
        i += 1
    return default


def _fill_combo(ctl, pairs, blank_first=False):
    if ctl is None:
        return
    labels = []
    if blank_first:
        labels.append(u"—")
    labels.extend([p[1] for p in pairs])
    try:
        ctl.removeItems(0, ctl.getItemCount())
    except Exception:
        pass
    try:
        ctl.addItems(tuple(labels), 0)
    except Exception:
        i = 0
        while i < len(labels):
            try:
                ctl.addItem(labels[i], i)
            except Exception:
                pass
            i += 1


def _example_label_to_expr(label):
    lab = unicode(label or u"").strip().casefold()
    if lab in (u"", u"—", u"-"):
        return u""
    for name, expr in _cfg.LAMBDA_EXAMPLE_CHOICES:
        if unicode(name).casefold() == lab:
            return unicode(expr)
    return u""


def show_step_condition_dialog( parent_dialog=None, initial_text=u"", doc=None, is_plugin=False, toolkit=None, create_peer_fn=None, add_fixed=None, add_edit=None, add_button=None, add_combo=None, wire_help=None, set_sizeable=None, open_lambda_builder_fn=None, load_saved_lambdas_fn=None, save_lambda_fn=None, refresh_saved_labels_fn=None):
    """
    Диалог условия. Возвращает:
      None — Отмена
      '' — Очистить условие
      str — JSON для записи в D/E
    """
    ctx = uno.getComponentContext()
    sm = ctx.getServiceManager()
    if toolkit is None:
        toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)

    dw = int(getattr(_cfg, "DIALOG_W", 580) or 580)
    dh = int(getattr(_cfg, "DIALOG_H", 420) or 420)
    m = int(getattr(_cfg, "DIALOG_MARGIN", 12) or 12)
    row_h = int(getattr(_cfg, "DIALOG_ROW_H", 28) or 28)
    ctl_h = int(getattr(_cfg, "DIALOG_CTL_H", 22) or 22)
    lbl_h = int(getattr(_cfg, "DIALOG_LBL_H", 16) or 16)
    lbl_w = int(getattr(_cfg, "DIALOG_LBL_W", 150) or 150)
    btn_h = int(getattr(_cfg, "DIALOG_BTN_H", 24) or 24)
    btn_w = int(getattr(_cfg, "DIALOG_BTN_W", 100) or 100)
    btn_gap = int(getattr(_cfg, "DIALOG_BTN_GAP", 10) or 10)
    footer_h = int(getattr(_cfg, "DIALOG_FOOTER_H", 36) or 36)
    lambda_edit_h = int(getattr(_cfg, "DIALOG_LAMBDA_EDIT_H", 110) or 110)
    os_users_h = int(getattr(_cfg, "DIALOG_OS_USERS_EDIT_H", 72) or 72)
    hint_h = int(getattr(_cfg, "DIALOG_HINT_H", 36) or 36)

    dm.PositionX = 120
    dm.PositionY = 60
    dm.Width = dw
    dm.Height = dh
    dm.Title = _cfg.DIALOG_TITLE
    if set_sizeable is not None:
        try:
            set_sizeable(dm)
        except Exception:
            pass

    def _fixed(name, text, x, y, w, h, multiline=False):
        if add_fixed is not None:
            add_fixed(dm, name, text, x, y, w, h, multiline=multiline)
            return
        model = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        model.Name = name
        model.PositionX = x
        model.PositionY = y
        model.Width = w
        model.Height = h
        model.Label = unicode(text or u"")
        model.MultiLine = bool(multiline)
        dm.insertByName(name, model)

    def _edit(name, x, y, w, h, multiline=False):
        if add_edit is not None:
            add_edit(dm, name, x, y, w, h, multiline=multiline)
            return
        model = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        model.Name = name
        model.PositionX = x
        model.PositionY = y
        model.Width = w
        model.Height = h
        model.MultiLine = bool(multiline)
        if multiline:
            try:
                model.VScroll = True
            except Exception:
                pass
        dm.insertByName(name, model)

    def _button(name, label, x, y, w, h):
        if add_button is not None:
            add_button(dm, name, label, x, y, w, h)
            return
        model = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        model.Name = name
        model.PositionX = x
        model.PositionY = y
        model.Width = w
        model.Height = h
        model.Label = unicode(label or u"")
        dm.insertByName(name, model)

    def _combo(name, x, y, w, h):
        if add_combo is not None:
            add_combo(dm, name, x, y, w, h)
            return
        model = dm.createInstance("com.sun.star.awt.UnoControlComboBoxModel")
        model.Name = name
        model.PositionX = x
        model.PositionY = y
        model.Width = w
        model.Height = h
        model.Dropdown = True
        try:
            model.LineCount = 12
        except Exception:
            pass
        dm.insertByName(name, model)

    content_w = dw - m * 2
    field_x = m + lbl_w + 8
    field_w = max(180, dw - field_x - m)

    hint = _cfg.DIALOG_HINT_PLUGIN if is_plugin else _cfg.DIALOG_HINT_CODEC
    y = m
    _fixed("HintLbl", hint, m, y, content_w, hint_h, multiline=True)
    y += hint_h + 6

    _fixed("KindLbl", u"Тип условия:", m, y + 2, lbl_w, lbl_h)
    _combo("KindCb", field_x, y, field_w, ctl_h)
    y += row_h + 4
    panel_top = y

    # --- lambda ---
    _fixed("LambdaLbl", u"expr (lambda: …):", m, panel_top, content_w, lbl_h)
    pick_w = 150
    ex_w = 150
    build_w = 110
    save_w = 90
    gap = 6
    tools_y = panel_top + lbl_h + 4
    _combo("LambdaPickCb", m, tools_y, pick_w, ctl_h)
    _button("LambdaBuildBtn", u"Конструктор…", m + pick_w + gap, tools_y, build_w, ctl_h)
    _button("LambdaSaveBtn", u"Сохранить", m + pick_w + gap + build_w + gap, tools_y, save_w, ctl_h)
    _combo("LambdaExampleCb", m + pick_w + gap + build_w + gap + save_w + gap, tools_y, ex_w, ctl_h)
    edit_y = tools_y + ctl_h + 6
    _edit("LambdaEd", m, edit_y, content_w, lambda_edit_h, multiline=True)
    _fixed(
        "LambdaHintLbl",
        u"Сохранённые ← | Конструктор | Сохранить | Примеры →. Helpers: var/env/os_user/glob_var.",
        m,
        edit_y + lambda_edit_h + 4,
        content_w,
        lbl_h,
    )

    # --- variable ---
    _fixed("VarNameLbl", u"Имя переменной:", m, panel_top + 2, lbl_w, lbl_h)
    _edit("VarNameEd", field_x, panel_top, field_w, ctl_h)
    _fixed("VarOpLbl", u"Оператор:", m, panel_top + row_h + 2, lbl_w, lbl_h)
    _combo("VarOpCb", field_x, panel_top + row_h, field_w, ctl_h)
    _fixed("VarValLbl", u"Значение / список:", m, panel_top + row_h * 2 + 2, lbl_w, lbl_h)
    _edit("VarValEd", field_x, panel_top + row_h * 2, field_w, ctl_h)
    _fixed("VarMissLbl", u"Если нет в карте:", m, panel_top + row_h * 3 + 2, lbl_w, lbl_h)
    _combo("VarMissCb", field_x, panel_top + row_h * 3, field_w, ctl_h)

    # --- env ---
    _fixed("EnvNameLbl", u"Переменная среды:", m, panel_top + 2, lbl_w, lbl_h)
    _edit("EnvNameEd", field_x, panel_top, field_w, ctl_h)
    _fixed("EnvOpLbl", u"Оператор:", m, panel_top + row_h + 2, lbl_w, lbl_h)
    _combo("EnvOpCb", field_x, panel_top + row_h, field_w, ctl_h)
    _fixed("EnvRhsLbl", u"Сравнить с:", m, panel_top + row_h * 2 + 2, lbl_w, lbl_h)
    _combo("EnvRhsCb", field_x, panel_top + row_h * 2, field_w, ctl_h)
    _fixed("EnvValLbl", u"Значение / имя:", m, panel_top + row_h * 3 + 2, lbl_w, lbl_h)
    _edit("EnvValEd", field_x, panel_top + row_h * 3, field_w, ctl_h)

    # --- os_user ---
    _fixed("OsOpLbl", u"Оператор:", m, panel_top + 2, lbl_w, lbl_h)
    _combo("OsOpCb", field_x, panel_top, field_w, ctl_h)
    _fixed("OsUsersLbl", u"Пользователи (user1,user2,…):", m, panel_top + row_h + 2, content_w, lbl_h)
    _edit("OsUsersEd", m, panel_top + row_h + lbl_h + 6, content_w, os_users_h, multiline=True)

    # --- footer: Clear | OK | Cancel .......... Help (без JsonChk)
    footer_y = dh - m - footer_h
    clear_w = 140
    _button("ClearBtn", u"Очистить условие", m, footer_y, clear_w, btn_h)
    _button("OkBtn", u"OK", m + clear_w + btn_gap, footer_y, btn_w, btn_h)
    _button("CancelBtn", u"Отмена", m + clear_w + btn_gap + btn_w + btn_gap, footer_y, btn_w, btn_h)
    help_w = 90
    _button("HelpButton", u"Справка", dw - m - help_w, footer_y, help_w, btn_h)
    try:
        ok_m = dm.getByName("OkBtn")
        ok_m.DefaultButton = True
        ok_m.PushButtonType = 1
    except Exception:
        pass
    try:
        dm.getByName("CancelBtn").PushButtonType = 2
    except Exception:
        pass
    try:
        dm.getByName("HelpButton").PushButtonType = 0
    except Exception:
        pass
    try:
        dm.getByName("ClearBtn").PushButtonType = 0
    except Exception:
        pass
    try:
        dm.getByName("LambdaBuildBtn").PushButtonType = 0
        dm.getByName("LambdaSaveBtn").PushButtonType = 0
    except Exception:
        pass

    dlg = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    if create_peer_fn is not None:
        create_peer_fn(dlg, toolkit, doc, parent_dialog, center_in_main_wizard=bool(parent_dialog))
    else:
        dlg.createPeer(toolkit, None)
    if wire_help is not None:
        try:
            wire_help(dlg, doc=doc, help_key=u"условие_выполнения")
        except Exception:
            try:
                wire_help(dlg, doc=doc)
            except Exception:
                pass

    kind_cb = _dlg_ctl(dlg, "KindCb")
    lambda_ed = _dlg_ctl(dlg, "LambdaEd")
    lambda_pick = _dlg_ctl(dlg, "LambdaPickCb")
    lambda_example = _dlg_ctl(dlg, "LambdaExampleCb")
    var_name = _dlg_ctl(dlg, "VarNameEd")
    var_op = _dlg_ctl(dlg, "VarOpCb")
    var_val = _dlg_ctl(dlg, "VarValEd")
    var_miss = _dlg_ctl(dlg, "VarMissCb")
    env_name = _dlg_ctl(dlg, "EnvNameEd")
    env_op = _dlg_ctl(dlg, "EnvOpCb")
    env_rhs = _dlg_ctl(dlg, "EnvRhsCb")
    env_val = _dlg_ctl(dlg, "EnvValEd")
    os_op = _dlg_ctl(dlg, "OsOpCb")
    os_users = _dlg_ctl(dlg, "OsUsersEd")

    _fill_combo(kind_cb, _cfg.KIND_CHOICES)
    _fill_combo(var_op, _cfg.COMPARE_OPS)
    _fill_combo(var_miss, _cfg.MISSING_CHOICES)
    _fill_combo(env_op, _cfg.COMPARE_OPS)
    _fill_combo(env_rhs, _cfg.RHS_KIND_CHOICES)
    _fill_combo(os_op, _cfg.OS_USER_OPS)
    _fill_combo(lambda_example, _cfg.LAMBDA_EXAMPLE_CHOICES, blank_first=True)

    saved_state = [()]
    if load_saved_lambdas_fn is not None:
        try:
            saved_state[0] = list(load_saved_lambdas_fn() or [])
        except Exception:
            saved_state[0] = []

    def _refresh_saved_pick(expr_text=u""):
        labels = (u"—",)
        if refresh_saved_labels_fn is not None:
            try:
                labels = tuple(refresh_saved_labels_fn(saved_state[0]) or (u"—",))
            except Exception:
                labels = (u"—",)
        else:
            names = []
            for item in saved_state[0] or ():
                nm = unicode(item.get(u"name") or u"").strip()
                if nm:
                    names.append(nm)
            labels = (u"—",) + tuple(names)
        try:
            lambda_pick.removeItems(0, lambda_pick.getItemCount())
        except Exception:
            pass
        try:
            lambda_pick.addItems(tuple(labels), 0)
        except Exception:
            pass
        pick_name = u"—"
        ex = unicode(expr_text or u"").strip()
        if ex:
            for item in saved_state[0] or ():
                if unicode(item.get(u"expr") or u"").strip() == ex:
                    pick_name = unicode(item.get(u"name") or u"—")
                    break
        _set_text(lambda_pick, pick_name)

    initial = unicode(initial_text or u"").strip()
    cond = None
    if initial != u"":
        try:
            cond = _lib.decode_step_condition(initial)
        except Exception:
            cond = None
    kind0 = (cond or {}).get(u"kind") or _cfg.KIND_OS_USER
    _combo_select_by_data(kind_cb, kind0, _cfg.KIND_CHOICES)
    if cond:
        if cond.get(u"kind") == _cfg.KIND_LAMBDA:
            _set_text(lambda_ed, cond.get(u"expr") or u"")
        elif cond.get(u"kind") == _cfg.KIND_VARIABLE:
            _set_text(var_name, cond.get(u"name") or u"")
            _combo_select_by_data(var_op, cond.get(u"op") or u"eq", _cfg.COMPARE_OPS)
            val = cond.get(u"value")
            if isinstance(val, (list, tuple)):
                _set_text(var_val, u",".join([unicode(x) for x in val]))
            else:
                _set_text(var_val, val or u"")
            _combo_select_by_data(var_miss, cond.get(u"missing") or u"false", _cfg.MISSING_CHOICES)
        elif cond.get(u"kind") == _cfg.KIND_ENV:
            _set_text(env_name, cond.get(u"env_name") or u"")
            _combo_select_by_data(env_op, cond.get(u"op") or u"eq", _cfg.COMPARE_OPS)
            _combo_select_by_data(env_rhs, cond.get(u"rhs_kind") or u"literal", _cfg.RHS_KIND_CHOICES)
            if (cond.get(u"rhs_kind") or u"literal") == u"literal":
                val = cond.get(u"value")
                if isinstance(val, (list, tuple)):
                    _set_text(env_val, u",".join([unicode(x) for x in val]))
                else:
                    _set_text(env_val, val or u"")
            else:
                _set_text(env_val, cond.get(u"global_name") or u"")
        elif cond.get(u"kind") == _cfg.KIND_OS_USER:
            _combo_select_by_data(os_op, cond.get(u"op") or u"in", _cfg.OS_USER_OPS)
            _set_text(os_users, u",".join(cond.get(u"users") or []))

    _refresh_saved_pick(_get_text(lambda_ed))

    panel_names = {
        _cfg.KIND_LAMBDA: (
            "LambdaLbl",
            "LambdaEd",
            "LambdaPickCb",
            "LambdaBuildBtn",
            "LambdaSaveBtn",
            "LambdaExampleCb",
            "LambdaHintLbl",
        ),
        _cfg.KIND_VARIABLE: (
            "VarNameLbl",
            "VarNameEd",
            "VarOpLbl",
            "VarOpCb",
            "VarValLbl",
            "VarValEd",
            "VarMissLbl",
            "VarMissCb",
        ),
        _cfg.KIND_ENV: (
            "EnvNameLbl",
            "EnvNameEd",
            "EnvOpLbl",
            "EnvOpCb",
            "EnvRhsLbl",
            "EnvRhsCb",
            "EnvValLbl",
            "EnvValEd",
        ),
        _cfg.KIND_OS_USER: (
            "OsOpLbl",
            "OsOpCb",
            "OsUsersLbl",
            "OsUsersEd",
        ),
    }

    def _show_kind(kind):
        for k, names in panel_names.items():
            vis = k == kind
            for nm in names:
                _set_visible(_dlg_ctl(dlg, nm), vis)

    _show_kind(kind0)

    result = {u"action": None, u"text": None}

    def _apply_saved_by_name(name):
        nm = unicode(name or u"").strip()
        if nm in (u"", u"—", u"-"):
            return
        for item in saved_state[0] or ():
            if unicode(item.get(u"name") or u"").strip() == nm:
                _set_text(lambda_ed, item.get(u"expr") or u"")
                return

    def _open_builder():
        if open_lambda_builder_fn is None:
            return
        try:
            ok, expr = open_lambda_builder_fn(
                dlg,
                kind=getattr(_cfg, "LAMBDA_BUILDER_KIND", u"gate"),
                initial_expr=_get_text(lambda_ed),
                doc=doc,
                saved_lambdas=saved_state[0],
            )
        except Exception:
            return
        if ok:
            _set_text(lambda_ed, expr or u"")
            try:
                if load_saved_lambdas_fn is not None:
                    saved_state[0] = list(load_saved_lambdas_fn() or [])
            except Exception:
                pass
            _refresh_saved_pick(_get_text(lambda_ed))

    def _save_current_lambda():
        if save_lambda_fn is None:
            return
        expr = _get_text(lambda_ed).strip()
        if expr == u"":
            return
        try:
            updated = save_lambda_fn(dlg, expr, saved_state[0], doc=doc)
            if updated is not None:
                saved_state[0] = list(updated or [])
        except Exception:
            pass
        _refresh_saved_pick(expr)

    class _H(unohelper.Base, XActionListener, XItemListener):
        def disposing(self, ev):
            pass

        def itemStateChanged(self, ev):
            try:
                src = ev.Source
                name = unicode(src.getModel().Name)
            except Exception:
                name = u""
            if name == u"KindCb":
                try:
                    kind = _combo_selected_code(kind_cb, _cfg.KIND_CHOICES, _cfg.KIND_OS_USER)
                    _show_kind(kind)
                except Exception:
                    pass
                return
            if name == u"LambdaPickCb":
                try:
                    _apply_saved_by_name(_get_text(lambda_pick))
                except Exception:
                    pass
                return
            if name == u"LambdaExampleCb":
                try:
                    expr = _example_label_to_expr(_get_text(lambda_example))
                    if expr:
                        _set_text(lambda_ed, expr)
                        _refresh_saved_pick(expr)
                except Exception:
                    pass

        def actionPerformed(self, ev):
            try:
                cmd = unicode(ev.ActionCommand or u"")
            except Exception:
                cmd = u""
            src = getattr(ev, "Source", None)
            try:
                name = unicode(src.getModel().Name)
            except Exception:
                name = cmd
            if name in (u"LambdaBuildBtn",) or cmd == u"LambdaBuildBtn":
                _open_builder()
                return
            if name in (u"LambdaSaveBtn",) or cmd == u"LambdaSaveBtn":
                _save_current_lambda()
                return
            if name in (u"ClearBtn",) or cmd == u"ClearBtn":
                result[u"action"] = u"clear"
                try:
                    dlg.endExecute()
                except Exception:
                    pass
                return
            if name in (u"OkBtn", u"OkButton", u"ok", u"OK") or cmd in (u"ok", u"OK", u"OkBtn"):
                result[u"action"] = u"ok"
                try:
                    dlg.endExecute()
                except Exception:
                    pass
                return
            if name in (u"CancelBtn", u"CancelButton", u"cancel") or cmd in (
                u"cancel",
                u"CancelBtn",
            ):
                result[u"action"] = u"cancel"
                try:
                    dlg.endExecute()
                except Exception:
                    pass

    h = _H()
    for listen_name in (u"KindCb", u"LambdaPickCb", u"LambdaExampleCb"):
        ctl = _dlg_ctl(dlg, listen_name)
        if ctl is None:
            continue
        try:
            ctl.addItemListener(h)
        except Exception:
            pass
    for btn_name in (
        u"ClearBtn",
        u"OkBtn",
        u"CancelBtn",
        u"LambdaBuildBtn",
        u"LambdaSaveBtn",
    ):
        btn = _dlg_ctl(dlg, btn_name)
        if btn is None:
            continue
        try:
            btn.setActionCommand(btn_name)
            btn.addActionListener(h)
        except Exception:
            pass

    rc = dlg.execute()
    action = result.get(u"action")
    if action == u"clear":
        return u""
    if action == u"cancel" or (action is None and rc != 1):
        return None
    kind = _combo_selected_code(kind_cb, _cfg.KIND_CHOICES, _cfg.KIND_OS_USER)
    obj = {u"v": 1, u"fn": _cfg.STEP_CONDITION_FN, u"kind": kind}
    try:
        if kind == _cfg.KIND_LAMBDA:
            obj[u"expr"] = _get_text(lambda_ed).strip()
        elif kind == _cfg.KIND_VARIABLE:
            obj[u"name"] = _get_text(var_name).strip()
            obj[u"op"] = _combo_selected_code(var_op, _cfg.COMPARE_OPS, u"eq")
            obj[u"value"] = _get_text(var_val).strip()
            obj[u"missing"] = _combo_selected_code(var_miss, _cfg.MISSING_CHOICES, u"false")
            obj[u"compare_as"] = u"text"
            obj[u"case_sensitive"] = False
        elif kind == _cfg.KIND_ENV:
            obj[u"env_name"] = _get_text(env_name).strip()
            obj[u"op"] = _combo_selected_code(env_op, _cfg.COMPARE_OPS, u"eq")
            rhs = _combo_selected_code(env_rhs, _cfg.RHS_KIND_CHOICES, u"literal")
            obj[u"rhs_kind"] = rhs
            if rhs == u"literal":
                obj[u"value"] = _get_text(env_val).strip()
                obj[u"global_name"] = u""
            else:
                obj[u"global_name"] = _get_text(env_val).strip()
                obj[u"value"] = u""
            obj[u"missing"] = u"false"
            obj[u"compare_as"] = u"text"
            obj[u"case_sensitive"] = False
        else:
            obj[u"op"] = _combo_selected_code(os_op, _cfg.OS_USER_OPS, u"in")
            obj[u"users"] = _get_text(os_users).strip()
            obj[u"case_sensitive"] = False
        return _lib.encode_step_condition(obj)
    except Exception as err:
        raise ValueError(u"Условие: %s" % err)
