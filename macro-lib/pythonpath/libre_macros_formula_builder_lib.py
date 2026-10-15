# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.784"
"""
Субвизард «конструктор_формулы» для «применить_формулу».

Слева — ListBox типовых шаблонов; справа — выбор столбца, плейсхолдеры
и функции Calc. OK возвращает текст формулы в поле визарда.

Точка входа: show_formula_builder_dialog(...).
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


def _strip_header_quotes(name):
    s = _u(name).strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in (u"'", u'"'):
        return s[1:-1].strip()
    return s


def apply_formula_preset_placeholders(draft, column=None, column2=None):
    """Подмена [Столбец]/[Столбец2]/{Столбец} в черновике пресета."""
    text = _u(draft)
    col = _strip_header_quotes(column)
    col2 = _strip_header_quotes(column2) or col
    if col:
        text = text.replace(u"[Столбец]", u"[%s]" % col)
        text = text.replace(u"{Столбец}", u"{%s}" % col)
    if col2:
        text = text.replace(u"[Столбец2]", u"[%s]" % col2)
        text = text.replace(u"{Столбец2}", u"{%s}" % col2)
    return text


def column_insert_snippet(header, mode=u"bracket_row"):
    """Фрагмент ссылки на столбец по заголовку."""
    name = _strip_header_quotes(header)
    if name == u"":
        name = u"Столбец"
    mode = _u(mode).strip().casefold() or u"bracket_row"
    if mode == u"bracket":
        return u"[%s]" % name
    if mode == u"curly":
        return u"{%s}" % name
    if mode == u"curly_row":
        return u"{%s}{row_id}" % name
    return u"[%s]{row_id}" % name


def insert_into_text(text, insert, sel_min=None, sel_max=None):
    """Вставка в позицию каретки / выделения. → (new_text, caret)."""
    cur = _u(text)
    snip = _u(insert)
    if sel_min is None or sel_max is None:
        return cur + snip, len(cur) + len(snip)
    a = max(0, min(int(sel_min), len(cur)))
    b = max(a, min(int(sel_max), len(cur)))
    new_text = cur[:a] + snip + cur[b:]
    return new_text, a + len(snip)


def formula_builder_presets():
    try:
        import libre_macros_param_wizard_cfg as _pw_cfg

        items = getattr(_pw_cfg, "FORMULA_BUILDER_PRESETS", ()) or ()
        out = []
        for it in items:
            if isinstance(it, (list, tuple)) and len(it) >= 2:
                out.append((_u(it[0]), _u(it[1])))
            elif isinstance(it, dict):
                out.append((_u(it.get(u"label")), _u(it.get(u"formula") or it.get(u"expr"))))
        return out
    except Exception:
        return [(u"Пустая / своя", u"=")]


def _cfg_tuple(attr, fallback):
    try:
        import libre_macros_param_wizard_cfg as _pw_cfg

        items = getattr(_pw_cfg, attr, None)
        if items:
            return list(items)
    except Exception:
        pass
    return list(fallback)


def _dlg_add_fixed(dm, name, label, x, y, w, h, multiline=False):
    ctl = dm.createInstance(u"com.sun.star.awt.UnoControlFixedTextModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    ctl.Label = _u(label)
    if multiline:
        try:
            ctl.MultiLine = True
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
    if multiline:
        try:
            ctl.MultiLine = True
            ctl.VScroll = True
        except Exception:
            pass
    dm.insertByName(name, ctl)


def _dlg_add_combo(dm, name, x, y, w, h):
    ctl = dm.createInstance(u"com.sun.star.awt.UnoControlComboBoxModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    try:
        ctl.Dropdown = True
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


def _dlg_add_listbox(dm, name, x, y, w, h):
    ctl = dm.createInstance(u"com.sun.star.awt.UnoControlListBoxModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    try:
        ctl.Dropdown = False
        ctl.MultiSelection = False
    except Exception:
        pass
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


def _set_listbox(ctl, items, select_index=0):
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
    try:
        if labels:
            idx = int(select_index)
            if idx < 0 or idx >= len(labels):
                idx = 0
            ctl.selectItemPos(idx, True)
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


def _list_selected_index(ctl):
    try:
        poss = ctl.getSelectedItemsPos()
        if poss is not None and len(poss) > 0:
            return int(poss[0])
    except Exception:
        pass
    try:
        return int(ctl.getSelectedItemPos())
    except Exception:
        return -1


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


def show_formula_builder_dialog( parent_dialog, initial_formula=u"", doc=None, sheet_name=u"", headers=None, title=None):
    """
    Конструктор формулы Calc для «применить_формулу».

    Возвращает (True, formula) | (False, initial_или_пусто).
    """
    initial = _u(initial_formula)
    headers_list = []
    if isinstance(headers, (list, tuple)):
        for h in headers:
            t = _strip_header_quotes(h)
            if t != u"":
                headers_list.append(t)

    try:
        import libre_macros_param_wizard_cfg as _pw_cfg
    except Exception:
        _pw_cfg = None

    hint_text = u""
    if _pw_cfg is not None:
        hint_text = _u(getattr(_pw_cfg, "FORMULA_BUILDER_HINT", u"") or u"")
    if hint_text == u"":
        hint_text = (
            u"Типовые шаблоны слева; столбец и плейсхолдеры справа. "
            u"{row_id}, {col_id}, [Имя]{row_id}, R[-1]C."
        )

    presets = formula_builder_presets()
    token_btns = _cfg_tuple(
        "FORMULA_BUILDER_TOKEN_BUTTONS",
        ((u"row_id", u"{row_id}"), (u"col_id", u"{col_id}"), (u"r1c1", u"R[-1]C")),
    )
    func_btns = _cfg_tuple(
        "FORMULA_BUILDER_FUNC_BUTTONS",
        ((u"if", u"IF("), (u"iferror", u"IFERROR("), (u"sum", u"SUM(")),
    )
    col_insert_btns = _cfg_tuple(
        "FORMULA_BUILDER_COL_INSERT_BUTTONS",
        ((u"bracket_row", u"[…]{row_id}"), (u"bracket", u"[…]")),
    )

    win_title = _u(title) or u"Конструктор формулы"

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

    dw = 640
    m = 10
    left_w = 200
    gap = 8
    right_x = m + left_w + gap
    right_w = dw - right_x - m

    dm = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialogModel", ctx)
    dm.PositionX = 90
    dm.PositionY = 50
    dm.Width = dw
    dm.Height = 480
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
    _dlg_add_fixed(dm, u"HintLbl", hint_text, m, y, dw - 2 * m, 34, multiline=True)
    y += 36

    list_top = y
    list_h = 210
    _dlg_add_fixed(dm, u"PresetLbl", u"Типовые формулы:", m, list_top, left_w, 14)
    _dlg_add_listbox(dm, u"PresetLB", m, list_top + 14, left_w, list_h)

    ry = list_top
    _dlg_add_fixed(dm, u"ColLbl", u"Столбец 1:", right_x, ry, 70, 14)
    _dlg_add_combo(dm, u"ColCb", right_x + 72, ry - 2, right_w - 72, 16)
    ry += 20
    _dlg_add_fixed(dm, u"Col2Lbl", u"Столбец 2:", right_x, ry, 70, 14)
    _dlg_add_combo(dm, u"Col2Cb", right_x + 72, ry - 2, right_w - 72, 16)
    ry += 22

    _dlg_add_fixed(dm, u"ColInsLbl", u"Вставить столбец:", right_x, ry, right_w, 14)
    ry += 14
    bx = right_x
    bw = 78
    for bi, (_cid, clab) in enumerate(col_insert_btns):
        _dlg_add_button(
            dm, u"ColInsBtn_%d" % bi, clab, bx + bi * (bw + 3), ry, bw, 16
        )
    ry += 22

    _dlg_add_fixed(dm, u"TokLbl", u"Плейсхолдеры / знаки:", right_x, ry, right_w, 14)
    ry += 14
    tw = 58
    tx = right_x
    row_i = 0
    for ti, (_tid, tlab) in enumerate(token_btns):
        if ti and ti % 5 == 0:
            ry += 18
            tx = right_x
            row_i += 1
        _dlg_add_button(dm, u"TokBtn_%d" % ti, tlab, tx, ry, tw, 16)
        tx += tw + 3
    ry += 20

    _dlg_add_fixed(dm, u"FnLbl", u"Функции:", right_x, ry, right_w, 14)
    ry += 14
    fx = right_x
    fw = 70
    for fi, (_fid, flab) in enumerate(func_btns):
        if fi and fi % 4 == 0:
            ry += 18
            fx = right_x
        _dlg_add_button(dm, u"FnBtn_%d" % fi, flab, fx, ry, fw, 16)
        fx += fw + 3
    ry += 20

    tools_bottom = ry
    content_bottom = max(list_top + 14 + list_h, tools_bottom)
    y = content_bottom + 8

    _dlg_add_fixed(dm, u"ExprLbl", u"Формула:", m, y, dw - 2 * m, 14)
    y += 14
    expr_h = 84
    _dlg_add_edit(dm, u"ExprEd", m, y, dw - 2 * m, expr_h, multiline=True)
    try:
        from libre_macros_ui_theme import apply_edit_font

        # Крупнее обычного edit визарда — удобнее править длинные формулы.
        apply_edit_font(dm.getByName(u"ExprEd"), height=13)
    except Exception:
        pass
    y += expr_h + 6

    # Статус и кнопки на РАЗНЫХ рядах — иначе FixedText перекрывает OK/Отмена.
    _dlg_add_fixed(dm, u"StatusLbl", u"", m, y, dw - 2 * m, 14)
    y += 18
    btn_h = 22
    btn_w = 90
    btn_gap = 6
    ok_x = dw - m - btn_w * 3 - btn_gap * 2
    _dlg_add_button(dm, u"OkButton", u"OK", ok_x, y, btn_w, btn_h)
    _dlg_add_button(
        dm, u"CancelButton", u"Отмена", ok_x + btn_w + btn_gap, y, btn_w, btn_h
    )
    _dlg_add_button(
        dm, u"HelpButton", u"Справка", ok_x + (btn_w + btn_gap) * 2, y, btn_w, btn_h
    )
    try:
        dm.getByName(u"OkButton").DefaultButton = True
        # Без PushButtonType: закрытие только из обработчика (иначе гонка с overlay/peer).
    except Exception:
        pass
    dm.Height = y + btn_h + m + 4

    try:
        from libre_macros_ui_theme import apply_soft_gray_theme

        apply_soft_gray_theme(dm, title_text=win_title)
    except Exception:
        pass
    # Тема могла пересоздать FontDescriptor — вернуть крупный шрифт формулы.
    try:
        from libre_macros_ui_theme import apply_edit_font

        apply_edit_font(dm.getByName(u"ExprEd"), height=13)
    except Exception:
        pass

    dlg = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    if not _create_peer(dlg, toolkit, doc=doc, parent_dialog=parent_dialog):
        return False, initial

    preset_lb = dlg.getControl(u"PresetLB")
    col_cb = dlg.getControl(u"ColCb")
    col2_cb = dlg.getControl(u"Col2Cb")
    expr_ed = dlg.getControl(u"ExprEd")
    status_lbl = dlg.getControl(u"StatusLbl")
    try:
        from libre_macros_ui_theme import apply_edit_font

        apply_edit_font(expr_ed.Model, height=13)
    except Exception:
        pass

    state = {u"presets": presets, u"result": None, u"applying": False}

    def _set_status(msg):
        try:
            status_lbl.Model.Label = _u(msg)[:160]
        except Exception:
            try:
                status_lbl.setText(_u(msg)[:160])
            except Exception:
                pass

    def _col1():
        return _combo_text(col_cb).strip()

    def _col2():
        return _combo_text(col2_cb).strip()

    def _expr_get():
        try:
            return _u(expr_ed.getText())
        except Exception:
            return u""

    def _focus_expr_end():
        """Фокус в поле формулы, курсор в конец текста."""
        try:
            expr_ed.setFocus()
        except Exception:
            pass
        end = len(_expr_get())
        try:
            from com.sun.star.awt import Selection

            sel = Selection()
            sel.Min = int(end)
            sel.Max = int(end)
            expr_ed.setSelection(sel)
        except Exception:
            pass

    def _expr_set(text, caret=None, focus_end=False):
        try:
            expr_ed.setText(_u(text))
        except Exception:
            pass
        if focus_end:
            _focus_expr_end()
            return
        if caret is not None:
            try:
                expr_ed.setFocus()
            except Exception:
                pass
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
        new_text, _caret = insert_into_text(cur, snippet, a, b)
        _expr_set(new_text, focus_end=True)

    def _apply_preset_index(idx):
        if idx < 0 or idx >= len(state[u"presets"]):
            return
        _lab, draft = state[u"presets"][idx]
        filled = apply_formula_preset_placeholders(draft, column=_col1(), column2=_col2())
        if filled and not filled.startswith(u"=") and not filled.startswith(u"{"):
            # шаблоны без = — оставить как есть (условия удаления и т.п.)
            pass
        _expr_set(filled, focus_end=True)
        _set_status(u"Шаблон: %s" % _lab)

    # init
    col_items = list(headers_list) if headers_list else [u""]
    _set_combo(col_cb, col_items, col_items[0] if col_items else u"")
    col2_default = col_items[1] if len(col_items) > 1 else (col_items[0] if col_items else u"")
    _set_combo(col2_cb, col_items, col2_default)

    preset_labels = [p[0] for p in presets] or [u"Пустая / своя"]
    _set_listbox(preset_lb, preset_labels, 0)

    if initial.strip():
        _expr_set(initial)
        # подсветить совпадающий пресет, если есть
        sel_i = 0
        for i, (_lab, draft) in enumerate(presets):
            if _u(draft).strip() == initial.strip():
                sel_i = i
                break
        try:
            preset_lb.selectItemPos(sel_i, True)
        except Exception:
            pass
    else:
        _apply_preset_index(0)

    class _Handler(unohelper.Base if unohelper else object, XActionListener, XItemListener):
        def disposing(self, ev):
            pass

        def itemStateChanged(self, ev):
            try:
                src = ev.Source
            except Exception:
                return
            try:
                if src == preset_lb and not state[u"applying"]:
                    idx = _list_selected_index(preset_lb)
                    if idx >= 0:
                        state[u"applying"] = True
                        try:
                            _apply_preset_index(idx)
                        finally:
                            state[u"applying"] = False
            except Exception as err:
                _set_status(u"Ошибка: %s" % err)

        def actionPerformed(self, ev):
            try:
                cmd = _u(ev.ActionCommand)
            except Exception:
                try:
                    cmd = _u(ev.Source.Model.Name)
                except Exception:
                    return
            try:
                src_name = u""
                try:
                    src_name = _u(ev.Source.Model.Name)
                except Exception:
                    src_name = cmd

                if src_name == u"OkButton" or cmd == u"OkButton":
                    text = _expr_get().strip()
                    if text == u"":
                        text = u"="
                    state[u"result"] = text
                    try:
                        dlg.endExecute()
                    except Exception:
                        pass
                    return
                if src_name == u"CancelButton" or cmd == u"CancelButton":
                    state[u"result"] = None
                    try:
                        dlg.endExecute()
                    except Exception:
                        pass
                    return
                if src_name == u"HelpButton" or cmd == u"HelpButton":
                    body = u""
                    try:
                        from libre_macros_wizard_help_cfg import lookup_fn_help

                        body = _u(lookup_fn_help(u"конструктор_формулы") or u"")
                        if body == u"":
                            body = _u(lookup_fn_help(u"применить_формулу") or u"")
                    except Exception:
                        body = u""
                    if body == u"":
                        body = hint_text
                    try:
                        from com.sun.star.awt.MessageBoxType import MESSAGEBOX
                        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK

                        peer = dlg.getPeer()
                        mb = toolkit.createMessageBox(
                            peer,
                            MESSAGEBOX,
                            BUTTONS_OK,
                            u"Справка — конструктор формулы",
                            body[:2000],
                        )
                        mb.execute()
                    except Exception:
                        _set_status(u"Справка недоступна")
                    return

                if src_name.startswith(u"TokBtn_"):
                    try:
                        ti = int(src_name.split(u"_", 1)[1])
                    except Exception:
                        return
                    if 0 <= ti < len(token_btns):
                        _insert(_u(token_btns[ti][1]))
                    return
                if src_name.startswith(u"FnBtn_"):
                    try:
                        fi = int(src_name.split(u"_", 1)[1])
                    except Exception:
                        return
                    if 0 <= fi < len(func_btns):
                        _insert(_u(func_btns[fi][1]))
                    return
                if src_name.startswith(u"ColInsBtn_"):
                    try:
                        bi = int(src_name.split(u"_", 1)[1])
                    except Exception:
                        return
                    if 0 <= bi < len(col_insert_btns):
                        mode = _u(col_insert_btns[bi][0])
                        _insert(column_insert_snippet(_col1(), mode=mode))
                    return
            except Exception as err:
                _set_status(u"Ошибка: %s" % err)

    h = _Handler()
    try:
        preset_lb.addItemListener(h)
    except Exception:
        pass
    for name in (
        u"OkButton",
        u"CancelButton",
        u"HelpButton",
    ):
        try:
            ctl = dlg.getControl(name)
            ctl.setActionCommand(name)
            ctl.addActionListener(h)
        except Exception:
            pass
    for ti in range(len(token_btns)):
        try:
            name = u"TokBtn_%d" % ti
            ctl = dlg.getControl(name)
            ctl.setActionCommand(name)
            ctl.addActionListener(h)
        except Exception:
            pass
    for fi in range(len(func_btns)):
        try:
            name = u"FnBtn_%d" % fi
            ctl = dlg.getControl(name)
            ctl.setActionCommand(name)
            ctl.addActionListener(h)
        except Exception:
            pass
    for bi in range(len(col_insert_btns)):
        try:
            name = u"ColInsBtn_%d" % bi
            ctl = dlg.getControl(name)
            ctl.setActionCommand(name)
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

    if state[u"result"] is None:
        return False, initial
    return True, _u(state[u"result"])
