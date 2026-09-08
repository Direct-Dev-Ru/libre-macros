# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.703"
"""
Субвизард «Конструктор лямбда-выражения».

Логика пресетов / вставки / проверки + UNO-диалог.
Точка входа: show_lambda_builder_dialog(...).
"""

try:
    unicode
except NameError:
    unicode = str

try:
    import uno
except Exception:
    uno = None

try:
    import unohelper
except Exception:
    unohelper = None

try:
    from com.sun.star.awt import XActionListener, XItemListener
except Exception:
    XActionListener = object
    XItemListener = object


def _u(v):
    return unicode(v or u"")


def lambda_builder_kind_meta(kind=None):
    """Метаданные kind из cfg; fallback для rows_*."""
    try:
        import libre_macros_param_wizard_cfg as _pw_cfg

        meta = getattr(_pw_cfg, "LAMBDA_BUILDER_KIND_META", {}) or {}
        k = _u(kind).strip() or u"rows_value"
        if k in meta:
            return dict(meta[k])
        for code, m in meta.items():
            if _u(m.get(u"label")).casefold() == k.casefold():
                return dict(m)
    except Exception:
        pass
    if _u(kind).strip() in (u"rows_predicate", u"filter", u"predicate"):
        return {
            u"kind": u"rows_predicate",
            u"label": u"Фильтр строк",
            u"signature": u"lambda rows: ",
            u"compile": u"rows",
        }
    return {
        u"kind": u"rows_value",
        u"label": u"Значение ячейки",
        u"signature": u"lambda rows: ",
        u"compile": u"rows",
    }


def lambda_builder_presets_for_kind(kind):
    """Список (label, draft_expr) для kind."""
    try:
        import libre_macros_param_wizard_cfg as _pw_cfg

        bag = getattr(_pw_cfg, "LAMBDA_BUILDER_PRESETS", {}) or {}
        items = bag.get(_normalize_kind(kind)) or ()
        out = []
        for it in items:
            if isinstance(it, (list, tuple)) and len(it) >= 2:
                out.append((_u(it[0]), _u(it[1])))
            elif isinstance(it, dict):
                out.append((_u(it.get(u"label")), _u(it.get(u"expr"))))
        return out
    except Exception:
        return []


def _normalize_kind(kind):
    k = _u(kind).strip().casefold()
    aliases = {
        u"filter": u"rows_predicate",
        u"predicate": u"rows_predicate",
        u"row_filter": u"rows_predicate",
        u"value": u"rows_value",
        u"compute": u"rows_value",
        u"expr": u"rows_value",
        u"column_name": u"name",
        u"sheet": u"sheets",
        u"sheet_list": u"sheets",
    }
    if k in aliases:
        return aliases[k]
    meta = lambda_builder_kind_meta(kind)
    return _u(meta.get(u"kind") or kind).strip() or u"rows_value"


def quote_column_for_rows(header):
    """
    Имя столбца → литерал для rows(...).
    Предпочтение двойных кавычек; экранирование.
    """
    name = _u(header).strip()
    if name.startswith(u"'") and name.endswith(u"'") and len(name) >= 2:
        name = name[1:-1]
    if name.startswith(u'"') and name.endswith(u'"') and len(name) >= 2:
        name = name[1:-1]
    if u'"' in name and u"'" not in name:
        return u"'%s'" % name.replace(u"\\", u"\\\\")
    return u'"%s"' % name.replace(u"\\", u"\\\\").replace(u'"', u'\\"')


def rows_call_snippet(header, suffix=u""):
    return u"rows(%s)%s" % (quote_column_for_rows(header), _u(suffix))


def variable_placeholder(name, quoted=True):
    nm = _u(name).strip()
    if nm.startswith(u"<<") and nm.endswith(u">>"):
        ph = nm
    else:
        if nm.lower().startswith(u"переменные."):
            nm = nm.split(u".", 1)[-1].strip()
        ph = u"<<Переменные.%s>>" % nm
    if quoted:
        return u'"%s"' % ph
    return ph


def list_source_variable_names():
    """Уникальные имена переменных runtime-карты + глобальные (fallback)."""
    names = []
    seen = set()

    def _add(n):
        t = _u(n).strip()
        if t == u"":
            return
        key = t.casefold()
        if key in seen:
            return
        seen.add(key)
        names.append(t)

    try:
        import libre_macros_collect_cfg as _cw_cfg

        order = list(getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_ORDER", None) or [])
        vm = getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_MAP", None) or {}
        for key in order:
            bare = _u(key).split(u"~", 1)[0].strip()
            _add(bare)
        for key in vm.keys():
            bare = _u(key).split(u"~", 1)[0].strip()
            _add(bare)
    except Exception:
        pass
    if not names:
        try:
            from libre_macros_global_settings_lib import global_variables_for_edit

            for item in global_variables_for_edit() or ():
                if isinstance(item, dict):
                    _add(item.get(u"name"))
                else:
                    _add(item)
        except Exception:
            pass
    return names


def find_unquoted_placeholders(expr):
    """
    Найти <<…>> вне строковых литералов '…' / \"…\".
    Возвращает список (start, end, token).
    """
    s = _u(expr)
    out = []
    i = 0
    n = len(s)
    in_s = None  # "'" or '"'
    escape = False
    while i < n:
        ch = s[i]
        if in_s is not None:
            if escape:
                escape = False
            elif ch == u"\\":
                escape = True
            elif ch == in_s:
                in_s = None
            i += 1
            continue
        if ch in (u"'", u'"'):
            in_s = ch
            i += 1
            continue
        if ch == u"<" and i + 1 < n and s[i + 1] == u"<":
            j = s.find(u">>", i + 2)
            if j >= 0:
                token = s[i : j + 2]
                if token.startswith(u"<<") and u"Переменные." in token:
                    out.append((i, j + 2, token))
                i = j + 2
                continue
        i += 1
    return out


def wrap_unquoted_placeholders(expr):
    """Обернуть незакавыченные <<Переменные.…>> в \"…\"."""
    s = _u(expr)
    hits = find_unquoted_placeholders(s)
    if not hits:
        return s, 0
    parts = []
    last = 0
    count = 0
    for a, b, tok in hits:
        parts.append(s[last:a])
        parts.append(u'"%s"' % tok)
        last = b
        count += 1
    parts.append(s[last:])
    return u"".join(parts), count


def ensure_lambda_signature(expr, kind):
    """Если нет lambda/def — предложить обёртку сигнатурой kind."""
    text = _u(expr).strip()
    if text == u"":
        return text
    head = text.lstrip()
    if head.startswith(u"lambda") or head.startswith(u"def "):
        return text
    meta = lambda_builder_kind_meta(kind)
    sig = _u(meta.get(u"signature") or u"lambda rows: ")
    if not sig.endswith(u" "):
        sig = sig + u" "
    return sig + text


def apply_preset_placeholders(draft, column=None, var_name=None):
    """Подмена «Столбец»/«Имя»/A/B в черновике пресета."""
    text = _u(draft)
    col = _u(column).strip()
    if col.startswith(u"'") and col.endswith(u"'") and len(col) >= 2:
        col = col[1:-1]
    if col.startswith(u'"') and col.endswith(u'"') and len(col) >= 2:
        col = col[1:-1]
    var = _u(var_name).strip()
    if col:
        q = quote_column_for_rows(col)
        # rows("Столбец") / rows('Столбец')
        text = text.replace(u'rows("Столбец")', u"rows(%s)" % q)
        text = text.replace(u"rows('Столбец')", u"rows(%s)" % q)
        text = text.replace(u'"Столбец"', q)
        text = text.replace(u"'Столбец'", q)
        # Склеить два столбца: A / B — оба на выбранный, если один столбец
        text = text.replace(u'rows("A")', u"rows(%s)" % q)
        text = text.replace(u"rows('A')", u"rows(%s)" % q)
    if var:
        text = text.replace(u"<<Переменные.Имя>>", u"<<Переменные.%s>>" % var)
        text = text.replace(u'"<<Переменные.Имя>>"', u'"<<Переменные.%s>>"' % var)
    return text


def validate_lambda_expr(expr, kind):
    """
    Проверка: expand → sugar → sanitize → compile.
    Возвращает (ok: bool, message: str).
    """
    text = _u(expr).strip()
    if text == u"":
        return False, u"Пустое выражение"
    text = ensure_lambda_signature(text, kind)
    nk = _normalize_kind(kind)
    try:
        from libre_macros_lambda_column_lib import (
            compile_lambda_column_name,
            compile_lambda_column_sheet_pick,
            compile_rows_lambda,
        )
    except Exception as err:
        return False, u"Нет libre_macros_lambda_column_lib: %s" % err
    try:
        if nk in (u"rows_predicate", u"rows_value"):
            packed = compile_rows_lambda(text)
        elif nk == u"name":
            packed = compile_lambda_column_name(text)
        elif nk == u"sheets":
            packed = compile_lambda_column_sheet_pick(text)
        else:
            packed = compile_rows_lambda(text)
        if packed is None:
            return False, u"Пусто после разбора"
        return True, u"OK: компиляция успешна"
    except Exception as err:
        return False, _u(err)


def insert_into_text(text, insert, sel_min=None, sel_max=None):
    """Вставить insert в позицию курсора / вместо выделения."""
    s = _u(text)
    ins = _u(insert)
    try:
        a = int(sel_min) if sel_min is not None else len(s)
        b = int(sel_max) if sel_max is not None else a
    except Exception:
        a = len(s)
        b = a
    if a < 0:
        a = 0
    if b < a:
        b = a
    if a > len(s):
        a = len(s)
    if b > len(s):
        b = len(s)
    return s[:a] + ins + s[b:], a + len(ins)


def _dlg_add_fixed(dm, name, label, x, y, w, h, multiline=False):
    ctl = dm.createInstance(u"com.sun.star.awt.UnoControlFixedTextModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    ctl.Label = _u(label)
    ctl.MultiLine = bool(multiline)
    try:
        ctl.VerticalAlign = 0  # TOP
    except Exception:
        pass
    dm.insertByName(name, ctl)


def _dlg_add_edit(dm, name, x, y, w, h, multiline=False):
    ctl = dm.createInstance(u"com.sun.star.awt.UnoControlEditModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    ctl.MultiLine = bool(multiline)
    ctl.VScroll = bool(multiline)
    dm.insertByName(name, ctl)


def _dlg_add_combo(dm, name, x, y, w, h):
    ctl = dm.createInstance(u"com.sun.star.awt.UnoControlComboBoxModel")
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
    dm.insertByName(name, ctl)


def _dlg_add_button(dm, name, label, x, y, w, h):
    ctl = dm.createInstance(u"com.sun.star.awt.UnoControlButtonModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    ctl.Label = _u(label)
    dm.insertByName(name, ctl)


def _set_combo(ctl, items, select=u""):
    labels = tuple(_u(x) for x in (items or ()))
    try:
        ctl.removeItems(0, ctl.getItemCount())
    except Exception:
        pass
    if labels:
        try:
            ctl.addItems(labels, 0)
        except Exception:
            pass
    sel = _u(select)
    if sel == u"" and labels:
        sel = labels[0]
    try:
        ctl.setText(sel)
    except Exception:
        pass


def _combo_text(ctl):
    try:
        return _u(ctl.getText())
    except Exception:
        return u""


def _edit_selection(ctl):
    try:
        sel = ctl.getSelection()
        return int(sel.Min), int(sel.Max)
    except Exception:
        return None, None


def _create_peer(dlg, toolkit, doc=None, parent_dialog=None):
    try:
        from param_wizard import _nested_subdialog_create_peer

        return bool(
            _nested_subdialog_create_peer(
                dlg,
                toolkit,
                doc,
                parent_dialog,
                center_in_main_wizard=bool(parent_dialog),
            )
        )
    except Exception:
        pass
    try:
        peer = None
        if parent_dialog is not None:
            try:
                peer = parent_dialog.getPeer()
            except Exception:
                peer = None
        if peer is None and doc is not None:
            try:
                peer = doc.getCurrentController().getFrame().getContainerWindow()
            except Exception:
                peer = None
        dlg.createPeer(toolkit, peer)
        return True
    except Exception:
        return False


def show_lambda_builder_dialog( parent_dialog, kind=u"rows_value", initial_expr=u"", doc=None, sheet_name=u"", headers=None, saved_lambdas=None, title=None, lock_kind=True):
    """
    Субвизард конструктора.

    Возвращает (True, expr) | (False, initial_или_пусто).
    """
    kind0 = _normalize_kind(kind)
    initial = _u(initial_expr)
    headers_list = []
    if isinstance(headers, (list, tuple)):
        for h in headers:
            t = _u(h).strip()
            if t != u"":
                headers_list.append(t)

    try:
        import libre_macros_param_wizard_cfg as _pw_cfg
    except Exception:
        _pw_cfg = None

    kind_choices = []
    if _pw_cfg is not None:
        kind_choices = list(getattr(_pw_cfg, "LAMBDA_BUILDER_KIND_CHOICES", ()) or ())
    if not kind_choices:
        kind_choices = [
            (u"rows_predicate", u"Фильтр строк"),
            (u"rows_value", u"Значение ячейки"),
            (u"name", u"Имя столбца"),
            (u"sheets", u"Список листов"),
        ]
    kind_label_by_code = {c: lab for c, lab in kind_choices}
    kind_code_by_label = {lab.casefold(): c for c, lab in kind_choices}

    sugar_btns = []
    if _pw_cfg is not None:
        sugar_btns = list(getattr(_pw_cfg, "LAMBDA_BUILDER_SUGAR_BUTTONS", ()) or ())
    if not sugar_btns:
        sugar_btns = [
            (u"rows", u"rows(…)"),
            (u"rows_prev", u"rows(…)[-1]"),
            (u"upd_rows", u"upd_rows(…)"),
            (u"upd_seen", u"upd_rows_seen"),
            (u"upd_len", u"upd_rows_len"),
            (u"tr", u"tr(…)"),
            (u"str", u"str(…)"),
            (u"strip", u".strip()"),
            (u"empty", u"пусто"),
            (u"not_empty", u"непусто"),
        ]

    op_btns = []
    if _pw_cfg is not None:
        op_btns = list(getattr(_pw_cfg, "LAMBDA_BUILDER_OP_BUTTONS", ()) or ())
    if not op_btns:
        op_btns = [
            (u"==", u"=="),
            (u"!=", u"!="),
            (u">", u">"),
            (u"<", u"<"),
            (u"in", u"in"),
            (u"not", u"not"),
            (u"and", u"and"),
            (u"or", u"or"),
        ]

    hint_text = u""
    if _pw_cfg is not None:
        hint_text = _u(getattr(_pw_cfg, "LAMBDA_BUILDER_HINT", u"") or u"")
    if hint_text == u"":
        hint_text = (
            u"Sugar: rows(\"Кол\"), rows(...)[-1], upd_rows, tr(…). "
            u"<<Переменные.имя>> — вставляйте в кавычках. "
            u"Сырой текст можно править вручную."
        )

    win_title = _u(title)
    if win_title == u"":
        win_title = u"Конструктор лямбда-выражения"

    # Headless
    try:
        import libre_macros_collect_cfg as _cw_cfg

        if bool(getattr(_cw_cfg, "_MERGE_QA_HEADLESS", False)):
            return False, initial
    except Exception:
        pass

    if uno is None:
        return False, initial

    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(u"com.sun.star.awt.Toolkit", ctx)
    except Exception:
        return False, initial

    dw = 620
    m = 10
    dm = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialogModel", ctx)
    dm.PositionX = 100
    dm.PositionY = 60
    dm.Width = dw
    dm.Height = 520
    dm.Title = win_title
    try:
        from libre_macros_ui_theme import inner_dialog_set_sizeable

        inner_dialog_set_sizeable(dm)
    except Exception:
        try:
            dm.Sizeable = True
        except Exception:
            pass

    y = m
    _dlg_add_fixed(dm, u"HintLbl", hint_text, m, y, dw - 2 * m, 36, multiline=True)
    y += 40

    _dlg_add_fixed(dm, u"KindLbl", u"Режим:", m, y, 48, 14)
    _dlg_add_combo(dm, u"KindCb", m + 50, y - 2, 150, 16)
    _dlg_add_fixed(dm, u"PresetLbl", u"Шаблон:", m + 210, y, 52, 14)
    _dlg_add_combo(dm, u"PresetCb", m + 264, y - 2, dw - m - 264, 16)
    y += 22

    _dlg_add_fixed(dm, u"ColLbl", u"Столбцы:", m, y, 56, 14)
    _dlg_add_combo(dm, u"ColCb", m + 58, y - 2, 200, 16)
    bx = m + 264
    bw = 56
    for bi, (_sid, slab) in enumerate(sugar_btns[:5]):
        _dlg_add_button(dm, u"SugarBtn_%d" % bi, slab, bx + bi * (bw + 2), y - 2, bw, 14)
    y += 20
    bx = m
    for bi, (_sid, slab) in enumerate(sugar_btns[5:]):
        idx = bi + 5
        _dlg_add_button(dm, u"SugarBtn_%d" % idx, slab, bx + bi * (bw + 2), y - 2, bw, 14)
    y += 20

    _dlg_add_fixed(dm, u"OpLbl", u"Операторы:", m, y, 64, 14)
    ox = m + 68
    ow = 36
    for oi, (_oid, olab) in enumerate(op_btns):
        _dlg_add_button(dm, u"OpBtn_%d" % oi, olab, ox + oi * (ow + 2), y - 2, ow, 14)
    y += 20

    _dlg_add_fixed(dm, u"VarLbl", u"Переменные:", m, y, 70, 14)
    _dlg_add_combo(dm, u"VarCb", m + 72, y - 2, 180, 16)
    _dlg_add_button(dm, u"VarQuotedBtn", u"В кавычках", m + 258, y - 2, 80, 14)
    _dlg_add_button(dm, u"VarRawBtn", u"Как есть", m + 342, y - 2, 70, 14)
    y += 22

    _dlg_add_fixed(dm, u"ExprLbl", u"Выражение:", m, y, dw - 2 * m, 14)
    y += 14
    expr_h = 110
    _dlg_add_edit(dm, u"ExprEd", m, y, dw - 2 * m, expr_h, multiline=True)
    y += expr_h + 6

    _dlg_add_button(dm, u"CheckBtn", u"Проверить", m, y, 80, 16)
    _dlg_add_fixed(dm, u"StatusLbl", u"", m + 88, y, dw - m - 88 - m, 16)
    y += 22

    _dlg_add_fixed(dm, u"SavedLbl", u"Сохранённые:", m, y, 80, 14)
    _dlg_add_combo(dm, u"SavedCb", m + 84, y - 2, 200, 16)
    _dlg_add_button(dm, u"SavedMgrBtn", u"Сохранить…", m + 290, y - 2, 80, 14)
    y += 24

    btn_h = 18
    _dlg_add_button(dm, u"OkButton", u"OK", m, y, 80, btn_h)
    _dlg_add_button(dm, u"CancelButton", u"Отмена", m + 88, y, 80, btn_h)
    _dlg_add_button(dm, u"HelpButton", u"Справка", dw - m - 90, y, 90, btn_h)
    try:
        dm.getByName(u"OkButton").DefaultButton = True
        dm.getByName(u"OkButton").PushButtonType = 1
        dm.getByName(u"CancelButton").PushButtonType = 2
    except Exception:
        pass
    dm.Height = y + btn_h + m

    dlg = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    if not _create_peer(dlg, toolkit, doc=doc, parent_dialog=parent_dialog):
        return False, initial

    kind_cb = dlg.getControl(u"KindCb")
    preset_cb = dlg.getControl(u"PresetCb")
    col_cb = dlg.getControl(u"ColCb")
    var_cb = dlg.getControl(u"VarCb")
    expr_ed = dlg.getControl(u"ExprEd")
    status_lbl = dlg.getControl(u"StatusLbl")
    saved_cb = dlg.getControl(u"SavedCb")

    state = {
        u"kind": kind0,
        u"lambdas": list(saved_lambdas or []),
        u"presets": [],
        u"lock_kind": bool(lock_kind),
        u"result": None,
    }

    def _set_status(msg):
        try:
            status_lbl.Model.Label = _u(msg)[:180]
        except Exception:
            try:
                status_lbl.setText(_u(msg)[:180])
            except Exception:
                pass

    def _current_kind():
        if state[u"lock_kind"]:
            return state[u"kind"]
        lab = _combo_text(kind_cb)
        code = kind_code_by_label.get(lab.casefold())
        if code:
            return code
        return state[u"kind"]

    def _selected_column():
        return _combo_text(col_cb).strip()

    def _selected_var():
        return _combo_text(var_cb).strip()

    def _expr_get():
        try:
            return _u(expr_ed.getText())
        except Exception:
            return u""

    def _expr_set(text, caret=None):
        try:
            expr_ed.setText(_u(text))
        except Exception:
            pass
        if caret is not None:
            try:
                from com.sun.star.awt import Selection

                sel = Selection()
                sel.Min = int(caret)
                sel.Max = int(caret)
                expr_ed.setSelection(sel)
            except Exception:
                pass

    def _insert(snippet):
        cur = _expr_get()
        a, b = _edit_selection(expr_ed)
        new_text, caret = insert_into_text(cur, snippet, a, b)
        _expr_set(new_text, caret)

    def _refresh_presets():
        presets = lambda_builder_presets_for_kind(_current_kind())
        state[u"presets"] = presets
        labels = [p[0] for p in presets] or [u"Пустой / свой"]
        _set_combo(preset_cb, labels, labels[0] if labels else u"")

    def _apply_preset_by_label(label):
        lab = _u(label)
        draft = u""
        for plab, pexpr in state[u"presets"]:
            if plab == lab:
                draft = pexpr
                break
        if draft == u"" and lab.casefold().startswith(u"пуст"):
            meta = lambda_builder_kind_meta(_current_kind())
            draft = _u(meta.get(u"signature") or u"lambda rows: ")
        if draft == u"":
            return
        cur = _expr_get().strip()
        if cur and cur != initial.strip():
            # не затираем молча — только если пресет «Пустой / свой» или подтверждение не нужно: заменим если пусто или совпадает с сигнатурой
            meta = lambda_builder_kind_meta(_current_kind())
            sig = _u(meta.get(u"signature") or u"").strip()
            if cur not in (u"", sig) and not lab.casefold().startswith(u"пуст"):
                # мягко: всё равно применяем (визард — черновик); статус
                _set_status(u"Шаблон применён (предыдущий текст заменён)")
        filled = apply_preset_placeholders(
            draft, column=_selected_column(), var_name=_selected_var()
        )
        _expr_set(filled)

    def _sugar_insert(sid):
        col = _selected_column()
        if sid == u"rows":
            if col:
                _insert(rows_call_snippet(col))
            else:
                _insert(u'rows("Столбец")')
            return
        if sid == u"rows_prev":
            if col:
                _insert(rows_call_snippet(col, u"[-1]"))
            else:
                _insert(u'rows("Столбец")[-1]')
            return
        if sid == u"upd_rows":
            if col:
                _insert(u"upd_rows(%s)" % quote_column_for_rows(col))
            else:
                _insert(u'upd_rows("Столбец")')
            return
        if sid == u"upd_seen":
            if col:
                q = quote_column_for_rows(col)
                _insert(u"upd_rows_seen(%s, rows(%s))" % (q, q))
            else:
                _insert(u'upd_rows_seen("Столбец", rows("Столбец"))')
            return
        if sid == u"upd_len":
            _insert(u"upd_rows_len()")
            return
        if sid == u"tr":
            _insert(u"tr(")
            return
        if sid == u"str":
            _insert(u"str(")
            return
        if sid == u"strip":
            _insert(u".strip()")
            return
        if sid == u"empty":
            if col:
                _insert(
                    u'str(rows(%s) or "").strip() == ""' % quote_column_for_rows(col)
                )
            else:
                _insert(u'str(rows("Столбец") or "").strip() == ""')
            return
        if sid == u"not_empty":
            if col:
                _insert(
                    u'str(rows(%s) or "").strip() != ""' % quote_column_for_rows(col)
                )
            else:
                _insert(u'str(rows("Столбец") or "").strip() != ""')
            return

    def _op_insert(oid):
        mapping = {
            u"==": u" == ",
            u"!=": u" != ",
            u">": u" > ",
            u"<": u" < ",
            u"in": u" in ",
            u"not": u"not ",
            u"and": u" and ",
            u"or": u" or ",
        }
        _insert(mapping.get(oid, u" %s " % oid))

    def _refresh_saved_pick(expr_hint=None):
        labels = [u"—"]
        for item in state[u"lambdas"]:
            nm = _u(item.get(u"name")).strip()
            if nm:
                labels.append(nm)
        pick = u"—"
        ex = _u(expr_hint) if expr_hint is not None else _expr_get().strip()
        if ex:
            for item in state[u"lambdas"]:
                if _u(item.get(u"expr")).strip() == ex:
                    pick = _u(item.get(u"name")).strip() or u"—"
                    break
        _set_combo(saved_cb, labels, pick)

    def _msg(text, title_s=u"Конструктор"):
        try:
            from com.sun.star.awt.MessageBoxType import MESSAGEBOX
            from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK

            peer = dlg.getPeer()
            mb = toolkit.createMessageBox(
                peer, MESSAGEBOX, BUTTONS_OK, title_s, _u(text)[:1500]
            )
            mb.execute()
        except Exception:
            _set_status(_u(text)[:120])

    def _ask_yes_no(text, title_s=u"Конструктор"):
        try:
            from com.sun.star.awt.MessageBoxType import MESSAGEBOX
            from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO
            from com.sun.star.awt.MessageBoxResults import YES

            peer = dlg.getPeer()
            mb = toolkit.createMessageBox(
                peer, MESSAGEBOX, BUTTONS_YES_NO, title_s, _u(text)[:1500]
            )
            return mb.execute() == YES
        except Exception:
            return False

    # init combos
    kind_labels = [lab for _c, lab in kind_choices]
    cur_lab = kind_label_by_code.get(kind0, kind_labels[0] if kind_labels else u"")
    _set_combo(kind_cb, kind_labels, cur_lab)
    try:
        kind_cb.setEnable(not bool(lock_kind))
    except Exception:
        pass

    col_items = list(headers_list) if headers_list else []
    if not col_items:
        col_items = [u""]
    _set_combo(col_cb, col_items, col_items[0] if col_items else u"")

    var_names = list_source_variable_names() or [u"Имя"]
    _set_combo(var_cb, var_names, var_names[0])

    _refresh_presets()
    if initial.strip():
        _expr_set(initial)
    else:
        meta = lambda_builder_kind_meta(kind0)
        _expr_set(_u(meta.get(u"signature") or u"lambda rows: "))
    _refresh_saved_pick(initial)

    class _Handler(unohelper.Base if unohelper else object, XActionListener, XItemListener):
        def disposing(self, ev):
            pass

        def itemStateChanged(self, ev):
            try:
                src = ev.Source
            except Exception:
                return
            try:
                if src == kind_cb:
                    new_k = _current_kind()
                    old_k = state[u"kind"]
                    if new_k != old_k:
                        cur = _expr_get().strip()
                        meta_old = lambda_builder_kind_meta(old_k)
                        sig_old = _u(meta_old.get(u"signature") or u"").strip()
                        if cur and cur != sig_old:
                            if not _ask_yes_no(
                                u"Сменить режим? Текст выражения может не "
                                u"соответствовать новой сигнатуре."
                            ):
                                # откат combo
                                _set_combo(
                                    kind_cb,
                                    kind_labels,
                                    kind_label_by_code.get(old_k, cur_lab),
                                )
                                return
                        state[u"kind"] = new_k
                        _refresh_presets()
                elif src == preset_cb:
                    _apply_preset_by_label(_combo_text(preset_cb))
                elif src == saved_cb:
                    nm = _combo_text(saved_cb).strip()
                    if nm in (u"", u"—", u"-"):
                        return
                    for item in state[u"lambdas"]:
                        if _u(item.get(u"name")).strip() == nm:
                            _expr_set(_u(item.get(u"expr")))
                            break
            except Exception as err:
                _set_status(u"Ошибка: %s" % err)

        def actionPerformed(self, ev):
            try:
                cmd = _u(ev.ActionCommand)
            except Exception:
                cmd = u""
            if cmd == u"":
                try:
                    cmd = _u(ev.Source.Model.Name)
                except Exception:
                    pass
            if cmd == u"CheckBtn":
                ok, msg = validate_lambda_expr(_expr_get(), _current_kind())
                _set_status(msg)
                if not ok:
                    _msg(msg, u"Проверка")
                return
            if cmd == u"VarQuotedBtn":
                nm = _selected_var() or u"Имя"
                _insert(variable_placeholder(nm, quoted=True))
                return
            if cmd == u"VarRawBtn":
                if _ask_yes_no(
                    u"Вставить <<Переменные.…>> без кавычек?\n"
                    u"Нужно только для чисел или если курсор уже внутри \"…\"."
                ):
                    nm = _selected_var() or u"Имя"
                    _insert(variable_placeholder(nm, quoted=False))
                return
            if cmd == u"SavedMgrBtn":
                try:
                    from param_wizard import _show_saved_lambda_manager_dialog

                    updated, expr_out = _show_saved_lambda_manager_dialog(
                        dlg, _expr_get(), state[u"lambdas"]
                    )
                    if updated is not None:
                        state[u"lambdas"] = list(updated)
                        if expr_out is not None:
                            _expr_set(_u(expr_out))
                        _refresh_saved_pick(_expr_get())
                except Exception as err:
                    _msg(u"Менеджер сохранённых: %s" % err)
                return
            if cmd == u"HelpButton":
                body = u""
                try:
                    from libre_macros_wizard_help_cfg import lookup_fn_help

                    body = _u(lookup_fn_help(u"лямбда_конструктор") or u"")
                except Exception:
                    body = u""
                if body == u"":
                    body = hint_text
                _msg(body, u"Справка")
                return
            if cmd == u"OkButton":
                text = _expr_get().strip()
                if text == u"":
                    _msg(u"Выражение пустое — укажите лямбду или Отмена.")
                    return
                text = ensure_lambda_signature(text, _current_kind())
                hits = find_unquoted_placeholders(text)
                if hits:
                    if _ask_yes_no(
                        u"Найден плейсхолдер <<Переменные.…>> вне кавычек.\n"
                        u"Обернуть в кавычки автоматически?"
                    ):
                        text, _n = wrap_unquoted_placeholders(text)
                ok, msg = validate_lambda_expr(text, _current_kind())
                if not ok:
                    _msg(u"Нельзя сохранить: %s" % msg, u"Ошибка")
                    _set_status(msg)
                    return
                state[u"result"] = text
                try:
                    dlg.endExecute()
                except Exception:
                    pass
                return
            if cmd == u"CancelButton":
                state[u"result"] = None
                try:
                    dlg.endExecute()
                except Exception:
                    pass
                return
            if cmd.startswith(u"SugarBtn_"):
                try:
                    idx = int(cmd.split(u"_", 1)[1])
                    sid = sugar_btns[idx][0]
                    _sugar_insert(sid)
                except Exception:
                    pass
                return
            if cmd.startswith(u"OpBtn_"):
                try:
                    idx = int(cmd.split(u"_", 1)[1])
                    oid = op_btns[idx][0]
                    _op_insert(oid)
                except Exception:
                    pass
                return

    h = _Handler()
    try:
        kind_cb.addItemListener(h)
        preset_cb.addItemListener(h)
        saved_cb.addItemListener(h)
    except Exception:
        pass
    for name in (
        u"CheckBtn",
        u"VarQuotedBtn",
        u"VarRawBtn",
        u"SavedMgrBtn",
        u"HelpButton",
        u"OkButton",
        u"CancelButton",
    ):
        try:
            ctl = dlg.getControl(name)
            ctl.setActionCommand(name)
            ctl.addActionListener(h)
        except Exception:
            pass
    for bi in range(len(sugar_btns)):
        try:
            ctl = dlg.getControl(u"SugarBtn_%d" % bi)
            ctl.setActionCommand(u"SugarBtn_%d" % bi)
            ctl.addActionListener(h)
        except Exception:
            pass
    for oi in range(len(op_btns)):
        try:
            ctl = dlg.getControl(u"OpBtn_%d" % oi)
            ctl.setActionCommand(u"OpBtn_%d" % oi)
            ctl.addActionListener(h)
        except Exception:
            pass

    try:
        dlg.execute()
    except Exception:
        return False, initial
    finally:
        try:
            dlg.dispose()
        except Exception:
            pass

    if state[u"result"] is not None:
        # вернуть обновлённый список лямбд через атрибут (родитель может прочитать)
        show_lambda_builder_dialog.last_saved_lambdas = list(state[u"lambdas"])
        return True, _u(state[u"result"])
    show_lambda_builder_dialog.last_saved_lambdas = list(state[u"lambdas"])
    return False, initial


show_lambda_builder_dialog.last_saved_lambdas = []
