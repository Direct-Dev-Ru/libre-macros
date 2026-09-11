# -*- coding: utf-8 -*-
"""Субвизард условия выполнения шага (колонка D / E)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.722"
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
        pass


def _combo_selected_code(ctl, pairs, default=u""):
    if ctl is None:
        return default
    try:
        pos = int(ctl.getSelectedItemPos())
    except Exception:
        pos = -1
    if pos < 0 or pos >= len(pairs):
        return default
    return pairs[pos][0]


def _fill_combo(ctl, pairs):
    if ctl is None:
        return
    try:
        ctl.removeItems(0, ctl.getItemCount())
    except Exception:
        pass
    labels = [p[1] for p in pairs]
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


def show_step_condition_dialog( parent_dialog=None, initial_text=u"", doc=None, is_plugin=False, toolkit=None, create_peer_fn=None, add_fixed=None, add_edit=None, add_button=None, add_combo=None, add_footer=None, wire_help=None, set_sizeable=None, inner_height=None):
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
    dw = 520
    dh = 340
    if inner_height is not None:
        try:
            dh = int(inner_height(340))
        except Exception:
            pass
    m = 10
    dm.PositionX = 120
    dm.PositionY = 80
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
        model = dm.createInstance("com.sun.star.awt.UnoControlListBoxModel")
        model.Name = name
        model.PositionX = x
        model.PositionY = y
        model.Width = w
        model.Height = h
        model.Dropdown = True
        dm.insertByName(name, model)

    hint = _cfg.DIALOG_HINT_PLUGIN if is_plugin else _cfg.DIALOG_HINT_CODEC
    y = m
    _fixed("HintLbl", hint, m, y, dw - m * 2, 28, multiline=True)
    y += 32
    _fixed("KindLbl", u"Тип условия:", m, y, 120, 12)
    _combo("KindCb", m + 120, y - 2, 240, 16)
    y += 22
    # lambda
    _fixed("LambdaLbl", u"expr (lambda: …):", m, y, dw - m * 2, 12)
    _edit("LambdaEd", m, y + 14, dw - m * 2, 70, multiline=True)
    # variable
    _fixed("VarNameLbl", u"Имя переменной:", m, y, 140, 12)
    _edit("VarNameEd", m + 140, y - 2, 200, 16)
    _fixed("VarOpLbl", u"Оператор:", m, y + 22, 140, 12)
    _combo("VarOpCb", m + 140, y + 20, 200, 16)
    _fixed("VarValLbl", u"Значение / список:", m, y + 44, 140, 12)
    _edit("VarValEd", m + 140, y + 42, 280, 16)
    _fixed("VarMissLbl", u"Если нет в карте:", m, y + 66, 140, 12)
    _combo("VarMissCb", m + 140, y + 64, 200, 16)
    # env
    _fixed("EnvNameLbl", u"Переменная среды:", m, y, 140, 12)
    _edit("EnvNameEd", m + 140, y - 2, 200, 16)
    _fixed("EnvOpLbl", u"Оператор:", m, y + 22, 140, 12)
    _combo("EnvOpCb", m + 140, y + 20, 200, 16)
    _fixed("EnvRhsLbl", u"Сравнить с:", m, y + 44, 140, 12)
    _combo("EnvRhsCb", m + 140, y + 42, 200, 16)
    _fixed("EnvValLbl", u"Значение / имя:", m, y + 66, 140, 12)
    _edit("EnvValEd", m + 140, y + 64, 280, 16)
    # os_user
    _fixed("OsOpLbl", u"Оператор:", m, y, 140, 12)
    _combo("OsOpCb", m + 140, y - 2, 200, 16)
    _fixed("OsUsersLbl", u"Пользователи (user1,user2,…):", m, y + 22, dw - m * 2, 12)
    _edit("OsUsersEd", m, y + 38, dw - m * 2, 40, multiline=True)

    footer_y = dh - m - 40
    _button("ClearBtn", u"Очистить условие", m, footer_y, 130, 18)
    if add_footer is not None:
        add_footer(dm, m, dh, help_btn=True, dw=dw)
    else:
        _button("OkBtn", u"OK", dw - m - 150, footer_y, 70, 18)
        _button("CancelBtn", u"Отмена", dw - m - 70, footer_y, 70, 18)

    dlg = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    if create_peer_fn is not None:
        create_peer_fn(dlg, toolkit, doc, parent_dialog, center_in_main_wizard=bool(parent_dialog))
    else:
        dlg.createPeer(toolkit, None)
    if wire_help is not None:
        try:
            wire_help(dlg, doc=doc)
        except Exception:
            pass

    kind_cb = _dlg_ctl(dlg, "KindCb")
    lambda_ed = _dlg_ctl(dlg, "LambdaEd")
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

    panel_names = {
        _cfg.KIND_LAMBDA: (
            "LambdaLbl",
            "LambdaEd",
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
                ctl = _dlg_ctl(dlg, nm)
                if ctl is None:
                    continue
                try:
                    ctl.setVisible(vis)
                except Exception:
                    try:
                        ctl.Model.EnableVisible = vis
                    except Exception:
                        pass

    _show_kind(kind0)

    result = {u"action": None, u"text": None}

    class _H(unohelper.Base, XActionListener, XItemListener):
        def disposing(self, ev):
            pass

        def itemStateChanged(self, ev):
            try:
                kind = _combo_selected_code(kind_cb, _cfg.KIND_CHOICES, _cfg.KIND_OS_USER)
                _show_kind(kind)
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
            if name in (u"ClearBtn",) or cmd == u"ClearBtn":
                result[u"action"] = u"clear"
                try:
                    dlg.endExecute()
                except Exception:
                    pass
                return
            if name in (u"OkBtn", u"ok", u"OK") or cmd in (u"ok", u"OK", u"OkBtn"):
                result[u"action"] = u"ok"
                try:
                    dlg.endExecute()
                except Exception:
                    pass
                return
            if name in (u"CancelBtn", u"cancel") or cmd in (u"cancel", u"CancelBtn"):
                result[u"action"] = u"cancel"
                try:
                    dlg.endExecute()
                except Exception:
                    pass

    h = _H()
    try:
        kind_cb.addItemListener(h)
    except Exception:
        pass
    for btn_name in (u"ClearBtn", u"OkBtn", u"CancelBtn"):
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
    # OK — собрать JSON
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
        try:
            from libre_macros_param_wizard_cfg import _show_message  # may fail
        except Exception:
            pass
        raise ValueError(u"Условие: %s" % err)
