# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.703"
"""
Диалог ручного ввода значения в runtime-карту переменных (жёлтый title).
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
    from com.sun.star.awt import XActionListener
    from com.sun.star.lang import XEventListener
except Exception:
    XActionListener = object
    XEventListener = object


def _log(msg):
    try:
        print("[manual_var] %s" % msg)
    except Exception:
        pass


def _parse_bool(raw, default=True):
    if raw is None:
        return bool(default)
    if isinstance(raw, bool):
        return raw
    s = unicode(raw).strip().casefold()
    if s in (u"",):
        return bool(default)
    if s in (u"1", u"true", u"yes", u"да", u"y", u"on"):
        return True
    if s in (u"0", u"false", u"no", u"нет", u"n", u"off"):
        return False
    return bool(default)


def _choices_list(raw):
    if raw is None:
        return []
    if isinstance(raw, (list, tuple)):
        out = []
        for x in raw:
            t = unicode(x or u"").strip()
            if t != u"":
                out.append(t)
        return out
    text = unicode(raw or u"").strip()
    if text == u"":
        return []
    parts = []
    for chunk in text.replace(u";", u",").split(u","):
        t = chunk.strip()
        if t != u"":
            parts.append(t)
    return parts


def _normalize_list_text(raw):
    """Список → одна строка через ';' (v1 для карты)."""
    if raw is None:
        return u""
    if isinstance(raw, (list, tuple)):
        parts = []
        for x in raw:
            t = unicode(x or u"").strip()
            if t != u"":
                parts.append(t)
        return u";".join(parts)
    text = unicode(raw or u"").replace(u"\r\n", u"\n").replace(u"\r", u"\n")
    parts = []
    for line in text.split(u"\n"):
        for chunk in line.replace(u",", u";").split(u";"):
            t = chunk.strip()
            if t != u"":
                parts.append(t)
    return u";".join(parts)


def _validate_typed_value(text, data_type, fmt_hint=u""):
    """
    Проверка/нормализация скаляра. Возвращает (ok, normalized_text, type_token, err).
    """
    s = unicode(text or u"").strip()
    dt = unicode(data_type or u"text").strip().casefold()
    if dt in (u"", u"text", u"str", u"string", u"текст"):
        return (True, s, u"text", u"")
    if dt in (u"number", u"num", u"float", u"int", u"число"):
        if s == u"":
            return (True, u"", u"number", u"")
        try:
            s2 = s.replace(u" ", u"").replace(u",", u".")
            float(s2)
            return (True, s2, u"number", u"")
        except Exception:
            return (False, s, u"number", u"Ожидается число")
    if dt in (u"bool", u"boolean", u"logical", u"логическое", u"логический"):
        if s == u"":
            return (True, u"", u"bool", u"")
        cf = s.casefold()
        if cf in (u"1", u"true", u"yes", u"да", u"y", u"on"):
            return (True, u"true", u"bool", u"")
        if cf in (u"0", u"false", u"no", u"нет", u"n", u"off"):
            return (True, u"false", u"bool", u"")
        return (False, s, u"bool", u"Ожидается Да/Нет (true/false)")
    if dt in (u"date", u"дата"):
        if s == u"":
            return (True, u"", u"date", u"")
        # Лёгкая проверка: ДД.ММ.ГГГГ или ГГГГ-ММ-ДД
        import re

        if re.match(r"^\d{1,2}\.\d{1,2}\.\d{2,4}$", s) or re.match(
            r"^\d{4}-\d{1,2}-\d{1,2}$", s
        ):
            return (True, s, u"date", u"")
        hint = unicode(fmt_hint or u"").strip()
        msg = u"Ожидается дата"
        if hint != u"":
            msg = u"%s (%s)" % (msg, hint)
        return (False, s, u"date", msg)
    if dt in (u"datetime", u"дата_время", u"дата-время"):
        if s == u"":
            return (True, u"", u"datetime", u"")
        import re

        if re.match(
            r"^\d{1,2}\.\d{1,2}\.\d{2,4}([ T]\d{1,2}:\d{2}(:\d{2})?)?$", s
        ) or re.match(r"^\d{4}-\d{1,2}-\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?$", s):
            return (True, s, u"datetime", u"")
        hint = unicode(fmt_hint or u"").strip()
        msg = u"Ожидается дата-время"
        if hint != u"":
            msg = u"%s (%s)" % (msg, hint)
        return (False, s, u"datetime", msg)
    return (True, s, u"text", u"")


def _dlg_add_fixed(dm, name, label, x, y, w, h, multiline=False):
    ctl = dm.createInstance(u"com.sun.star.awt.UnoControlFixedTextModel")
    ctl.Name = name
    ctl.PositionX = x
    ctl.PositionY = y
    ctl.Width = w
    ctl.Height = h
    ctl.Label = unicode(label or u"")
    ctl.MultiLine = bool(multiline)
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
    dm.insertByName(name, ctl)


def _create_peer(dlg, toolkit, doc=None):
    try:
        peer = None
        if doc is not None:
            try:
                win = doc.getCurrentController().getFrame().getContainerWindow()
                peer = win
            except Exception:
                peer = None
        if peer is None:
            peer = toolkit.getDesktopWindow()
        dlg.createPeer(toolkit, peer)
        return True
    except Exception as err:
        _log(u"createPeer: %s" % err)
        return False


def show_manual_variable_input_dialog(block, doc=None):
    """
    Показать диалог ввода. block — dict JSON-блока.

    Возвращает:
        (True, text, type_token) — OK;
        (False, None, None) — Cancel / ошибка UNO без default.
    """
    if not isinstance(block, dict):
        block = {}
    name = unicode(block.get(u"name") or u"").strip()
    title = unicode(block.get(u"title") or u"").strip()
    if title == u"":
        title = name if name != u"" else u"Ввод переменной"
    prompt = unicode(block.get(u"prompt") or u"").strip()
    fmt_hint = unicode(block.get(u"format") or u"").strip()
    value_kind = unicode(block.get(u"value_kind") or u"scalar").strip().casefold()
    if value_kind in (u"list", u"список", u"array"):
        value_kind = u"list"
    else:
        value_kind = u"scalar"
    data_type = unicode(block.get(u"data_type") or u"text").strip()
    required = _parse_bool(block.get(u"required"), default=True)
    choices = _choices_list(block.get(u"choices"))
    # default: remembered last (если включено) → иначе поле default из JSON
    sheet_name = unicode(
        block.get(u"_param_sheet")
        or block.get(u"param_sheet")
        or u""
    ).strip()
    default_text = u""
    try:
        from libre_macros_manual_var_store_lib import resolve_dialog_default

        default_text = unicode(
            resolve_dialog_default(block, sheet_name=sheet_name) or u""
        )
    except Exception:
        default_raw = block.get(u"default")
        if default_raw is None:
            default_text = u""
        else:
            default_text = unicode(default_raw)
    if value_kind == u"list" and default_text != u"":
        default_text = _normalize_list_text(default_text)

    def _finish_ok(text_val, token_val):
        # list v1: одна строка с «;», type_token=text
        out_text = unicode(text_val or u"")
        out_tok = u"text" if value_kind == u"list" else unicode(token_val or u"text")
        try:
            from libre_macros_manual_var_store_lib import save_dialog_answer

            save_dialog_answer(block, sheet_name, out_text)
        except Exception as err:
            _log(u"save remembered: %s" % err)
        if value_kind == u"list":
            return (True, out_text, u"text")
        return (True, out_text, out_tok)

    def _headless_or_default_reply():
        if default_text != u"" or not required:
            text = default_text
            if value_kind == u"list":
                text = _normalize_list_text(text)
            ok, norm, token, _err = _validate_typed_value(text, data_type, fmt_hint)
            if not ok and required:
                return (False, None, None)
            return _finish_ok(norm if ok else text, token if ok else u"text")
        return (False, None, None)

    # Headless QA / без UNO: default или fail
    try:
        import libre_macros_collect_cfg as _cw_cfg

        if bool(getattr(_cw_cfg, "_MERGE_QA_HEADLESS", False)):
            return _headless_or_default_reply()
    except Exception:
        pass
    if uno is None:
        reply = _headless_or_default_reply()
        if reply[0]:
            return reply
        _log(u"UNO unavailable, required input without default")
        return (False, None, None)

    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(u"com.sun.star.awt.Toolkit", ctx)
    except Exception as err:
        _log(u"toolkit: %s" % err)
        if default_text != u"" or not required:
            text = default_text
            if value_kind == u"list":
                text = _normalize_list_text(text)
            ok, norm, token, _err = _validate_typed_value(text, data_type, fmt_hint)
            return _finish_ok(norm if ok else text, token if ok else u"text")
        return (False, None, None)

    try:
        from libre_macros_ui_theme import (
            MANUAL_INPUT_VALUE_FONT_HEIGHT,
            apply_edit_font,
            inner_dialog_add_ok_cancel_footer,
            inner_dialog_footer_y,
            inner_dialog_set_sizeable,
            prepare_dialog_soft_gray_manual_input,
        )
    except Exception as err:
        _log(u"ui_theme: %s" % err)
        return (False, None, None)

    dm = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialogModel", ctx)
    m = 10
    dw = 420
    # Базовая высота поля: scalar 18 / list 60; +1 pt шрифта → ×1.1 по высоте.
    try:
        font_pt = int(MANUAL_INPUT_VALUE_FONT_HEIGHT)
    except Exception:
        font_pt = 11
    try:
        from libre_macros_ui_theme import DIALOG_EDIT_FONT_HEIGHT as _base_pt
        base_pt = int(_base_pt) if int(_base_pt) > 0 else 10
    except Exception:
        base_pt = 10
    scale = float(font_pt) / float(base_pt)
    use_combo = value_kind == u"scalar" and len(choices) > 0
    multiline = value_kind == u"list"
    edit_h_base = 60 if multiline else 18
    edit_h = max(edit_h_base, int(round(edit_h_base * scale)))
    # Диалог: запас под более высокое поле ввода.
    dh_extra = max(0, edit_h - edit_h_base)
    dh = (220 if value_kind != u"list" else 280) + dh_extra
    dm.PositionX = 120
    dm.PositionY = 90
    dm.Width = dw
    dm.Height = dh
    dm.Title = title
    inner_dialog_set_sizeable(dm)

    y = m
    hint_parts = []
    if prompt != u"":
        hint_parts.append(prompt)
    else:
        hint_parts.append(u"Введите значение переменной «%s»." % (name or u"?"))
    if fmt_hint != u"":
        hint_parts.append(u"Формат: %s" % fmt_hint)
    if value_kind == u"list":
        hint_parts.append(u"Список: через запятую, «;» или с новой строки.")
    hint = u"\n".join(hint_parts)
    _dlg_add_fixed(dm, u"HintLbl", hint, m, y, dw - 2 * m, 48, multiline=True)
    y += 52
    footer_y = inner_dialog_footer_y(dh, m)
    if y + edit_h + 8 > footer_y:
        edit_h = max(edit_h_base, footer_y - y - 8)
    if use_combo:
        _dlg_add_combo(dm, u"ValueCtrl", m, y, dw - 2 * m, edit_h if edit_h >= 18 else 18)
    else:
        _dlg_add_edit(dm, u"ValueCtrl", m, y, dw - 2 * m, edit_h, multiline=multiline)
    try:
        apply_edit_font(dm.getByName(u"ValueCtrl"), font_pt)
    except Exception:
        pass
    _dlg_add_fixed(dm, u"ErrLbl", u"", m, footer_y - 16, dw - 2 * m, 14)
    inner_dialog_add_ok_cancel_footer(
        dm, m, dh, ok_label=u"OK", cancel_label=u"Отмена", help_btn=True, dw=dw,
        help_key=u"ручной_ввод",
    )

    dlg = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    prepare_dialog_soft_gray_manual_input(dlg, title_text=title)
    if not _create_peer(dlg, toolkit, doc=doc):
        if default_text != u"" or not required:
            text = default_text
            if value_kind == u"list":
                text = _normalize_list_text(text)
            ok, norm, token, _err = _validate_typed_value(text, data_type, fmt_hint)
            return _finish_ok(norm if ok else text, token if ok else u"text")
        return (False, None, None)

    try:
        help_btn = dlg.getControl(u"HelpButton")
    except Exception:
        help_btn = None
    if help_btn is not None:
        help_body = u""
        try:
            import libre_macros_param_wizard_cfg as _pw_cfg
            help_body = unicode(
                (_pw_cfg.WIZARD_HINTS or {}).get(u"Ручной_ввод_в_карту_переменных", u"")
                or u""
            )
        except Exception:
            help_body = u""
        if help_body == u"":
            help_body = (
                u"Введите значение переменной для runtime-карты.\n"
                u"Строка параметра выше «Файлы-Источники» — значение доступно "
                u"в путях и масках (<<Переменные.Имя>>)."
            )

        class _HelpListener(XActionListener, XEventListener):
            def disposing(self, ev):
                pass

            def actionPerformed(self, ev):
                try:
                    from com.sun.star.awt.MessageBoxType import MESSAGEBOX
                    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
                    peer = dlg.getPeer()
                    mb = toolkit.createMessageBox(peer, MESSAGEBOX, BUTTONS_OK, u"Справка", help_body[:1500])
                    mb.execute()
                except Exception:
                    pass

        try:
            help_btn.setActionCommand(u"HelpButton")
            help_btn.addActionListener(_HelpListener())
        except Exception:
            pass

    try:
        ctl = dlg.getControl(u"ValueCtrl")
        if use_combo:
            try:
                ctl.removeItems(0, ctl.getItemCount())
            except Exception:
                pass
            try:
                ctl.addItems(tuple(choices), 0)
            except Exception:
                pass
            try:
                ctl.setText(default_text if default_text != u"" else (choices[0] if choices else u""))
            except Exception:
                pass
        else:
            try:
                ctl.setText(default_text)
            except Exception:
                pass
    except Exception:
        pass

    class _OkListener(XActionListener, XEventListener):
        def __init__(self, holder):
            self.holder = holder

        def disposing(self, ev):
            pass

        def actionPerformed(self, ev):
            self.holder[0] = True

    # Цикл валидации: при ошибке не закрываем через OK-логику вручную —
    # UNO OK закрывает диалог; при ошибке показываем снова.
    while True:
        rc = dlg.execute()
        try:
            rc = int(rc)
        except Exception:
            rc = 0
        if rc != 1:
            try:
                dlg.dispose()
            except Exception:
                pass
            return (False, None, None)
        raw_val = u""
        try:
            raw_val = unicode(dlg.getControl(u"ValueCtrl").getText() or u"")
        except Exception:
            raw_val = u""
        if value_kind == u"list":
            text = _normalize_list_text(raw_val)
        else:
            text = unicode(raw_val).strip()
        if text == u"" and required:
            try:
                dlg.getControl(u"ErrLbl").Text = u"Значение обязательно"
            except Exception:
                pass
            # recreate peer cycle: dispose and rebuild is heavy; set Text and re-execute
            # After execute() dialog is closed — need to show again
            try:
                dlg.dispose()
            except Exception:
                pass
            return show_manual_variable_input_dialog(block, doc=doc)
        ok, norm, token, err = _validate_typed_value(text, data_type, fmt_hint)
        if not ok:
            try:
                dlg.dispose()
            except Exception:
                pass
            # повтор с прежним default = введённое
            block2 = dict(block)
            block2[u"default"] = text
            # show error via prompt prefix
            prev_prompt = unicode(block2.get(u"prompt") or u"").strip()
            block2[u"prompt"] = (u"%s\n%s" % (err, prev_prompt)).strip()
            return show_manual_variable_input_dialog(block2, doc=doc)
        try:
            dlg.dispose()
        except Exception:
            pass
        return _finish_ok(norm, token)


def _manual_var_block_default_text(block, sheet_name=u""):
    """Значение по умолчанию для поля блока (remember → default)."""
    default_text = u""
    try:
        from libre_macros_manual_var_store_lib import resolve_dialog_default

        default_text = unicode(
            resolve_dialog_default(block, sheet_name=sheet_name) or u""
        )
    except Exception:
        default_raw = block.get(u"default")
        if default_raw is None:
            default_text = u""
        else:
            default_text = unicode(default_raw)
    value_kind = unicode(block.get(u"value_kind") or u"scalar").strip().casefold()
    if value_kind in (u"list", u"список", u"array"):
        value_kind = u"list"
    else:
        value_kind = u"scalar"
    if value_kind == u"list" and default_text != u"":
        default_text = _normalize_list_text(default_text)
    return default_text


def _manual_var_block_field_meta(block):
    """Метаданные поля блока для UI: label, control kind, defaults."""
    name = unicode(block.get(u"name") or u"").strip()
    title = unicode(block.get(u"title") or u"").strip()
    label = title if title != u"" else name
    if label == u"":
        label = u"?"
    value_kind = unicode(block.get(u"value_kind") or u"scalar").strip().casefold()
    if value_kind in (u"list", u"список", u"array"):
        value_kind = u"list"
    else:
        value_kind = u"scalar"
    data_type = unicode(block.get(u"data_type") or u"text").strip()
    fmt_hint = unicode(block.get(u"format") or u"").strip()
    required = _parse_bool(block.get(u"required"), default=True)
    choices = _choices_list(block.get(u"choices"))
    use_combo = value_kind == u"scalar" and len(choices) > 0
    multiline = value_kind == u"list"
    return {
        u"name": name,
        u"label": label,
        u"value_kind": value_kind,
        u"data_type": data_type,
        u"fmt_hint": fmt_hint,
        u"required": required,
        u"choices": choices,
        u"use_combo": use_combo,
        u"multiline": multiline,
    }


def show_manual_variable_input_combined_dialog(blocks, doc=None):
    """
    Один диалог для нескольких блоков ручного ввода.

    Подсказки (prompt, format) — только от первого блока.
    Возвращает (True, [(name, text, type_token), ...]) или (False, None).
    """
    if not blocks:
        return (True, [])
    norm_blocks = []
    bi = 0
    while bi < len(blocks):
        b = blocks[bi]
        bi = bi + 1
        if not isinstance(b, dict):
            continue
        name = unicode(b.get(u"name") or u"").strip()
        if name == u"":
            continue
        norm_blocks.append(dict(b))
    if not norm_blocks:
        return (True, [])
    if len(norm_blocks) == 1:
        ok, text, token = show_manual_variable_input_dialog(norm_blocks[0], doc=doc)
        if not ok:
            return (False, None)
        return (True, [(unicode(norm_blocks[0].get(u"name") or u"").strip(), text, token)])

    first = norm_blocks[0]
    sheet_name = unicode(
        first.get(u"_param_sheet") or first.get(u"param_sheet") or u""
    ).strip()
    title = unicode(first.get(u"title") or u"").strip()
    if title == u"":
        title = u"Ввод переменных"
    prompt = unicode(first.get(u"prompt") or u"").strip()
    fmt_hint = unicode(first.get(u"format") or u"").strip()

    metas = []
    defaults = []
    mi = 0
    while mi < len(norm_blocks):
        blk = norm_blocks[mi]
        mi = mi + 1
        metas.append(_manual_var_block_field_meta(blk))
        defaults.append(_manual_var_block_default_text(blk, sheet_name=sheet_name))

    def _headless_reply():
        entries = []
        di = 0
        while di < len(norm_blocks):
            blk = norm_blocks[di]
            meta = metas[di]
            text = defaults[di]
            di = di + 1
            if meta[u"required"] and unicode(text).strip() == u"":
                return (False, None)
            ok, norm, token, _err = _validate_typed_value(
                text, meta[u"data_type"], meta[u"fmt_hint"]
            )
            if not ok and meta[u"required"]:
                return (False, None)
            out_text = norm if ok else text
            out_tok = u"text" if meta[u"value_kind"] == u"list" else (token if ok else u"text")
            try:
                from libre_macros_manual_var_store_lib import save_dialog_answer

                save_dialog_answer(blk, sheet_name, out_text)
            except Exception as err:
                _log(u"save remembered combined: %s" % err)
            entries.append((meta[u"name"], out_text, out_tok))
        return (True, entries)

    try:
        import libre_macros_collect_cfg as _cw_cfg

        if bool(getattr(_cw_cfg, "_MERGE_QA_HEADLESS", False)):
            return _headless_reply()
    except Exception:
        pass
    if uno is None:
        reply = _headless_reply()
        if reply[0]:
            return reply
        _log(u"UNO unavailable, required combined input without default")
        return (False, None)

    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(u"com.sun.star.awt.Toolkit", ctx)
    except Exception as err:
        _log(u"toolkit combined: %s" % err)
        return _headless_reply()

    try:
        from libre_macros_ui_theme import (
            MANUAL_INPUT_VALUE_FONT_HEIGHT,
            apply_edit_font,
            inner_dialog_add_ok_cancel_footer,
            inner_dialog_footer_y,
            inner_dialog_set_sizeable,
            prepare_dialog_soft_gray_manual_input,
        )
    except Exception as err:
        _log(u"ui_theme combined: %s" % err)
        return (False, None)

    try:
        font_pt = int(MANUAL_INPUT_VALUE_FONT_HEIGHT)
    except Exception:
        font_pt = 11
    try:
        from libre_macros_ui_theme import DIALOG_EDIT_FONT_HEIGHT as _base_pt

        base_pt = int(_base_pt) if int(_base_pt) > 0 else 10
    except Exception:
        base_pt = 10
    scale = float(font_pt) / float(base_pt)

    m = 10
    dw = 460
    row_gap = 6
    lbl_h = 14
    edit_h_scalar = max(18, int(round(18 * scale)))
    edit_h_list = max(48, int(round(48 * scale)))
    hint_parts = []
    if prompt != u"":
        hint_parts.append(prompt)
    else:
        hint_parts.append(u"Введите значения переменных.")
    if fmt_hint != u"":
        hint_parts.append(u"Формат: %s" % fmt_hint)
    any_list = False
    ai = 0
    while ai < len(metas):
        if metas[ai][u"value_kind"] == u"list":
            any_list = True
            break
        ai = ai + 1
    if any_list:
        hint_parts.append(u"Список: через запятую, «;» или с новой строки.")
    hint = u"\n".join(hint_parts)
    hint_h = 52

    y = m + hint_h + 8
    field_layout = []
    fi = 0
    while fi < len(metas):
        meta = metas[fi]
        edit_h = edit_h_list if meta[u"multiline"] else edit_h_scalar
        field_layout.append((fi, y, edit_h, meta))
        y += lbl_h + 2 + edit_h + row_gap
        fi = fi + 1
    footer_reserve = 56
    dh = max(280, y + footer_reserve + m)

    dm = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialogModel", ctx)
    dm.PositionX = 120
    dm.PositionY = 80
    dm.Width = dw
    dm.Height = dh
    dm.Title = title
    inner_dialog_set_sizeable(dm)

    _dlg_add_fixed(dm, u"HintLbl", hint, m, m, dw - 2 * m, hint_h, multiline=True)
    for fi, fy, edit_h, meta in field_layout:
        ctl_suffix = unicode(fi)
        lbl = meta[u"label"]
        if meta[u"required"]:
            lbl = u"%s:" % lbl
        else:
            lbl = u"%s (необяз.):" % lbl
        _dlg_add_fixed(dm, u"NameLbl_%s" % ctl_suffix, lbl, m, fy, dw - 2 * m, lbl_h)
        ctrl_y = fy + lbl_h + 2
        if meta[u"use_combo"]:
            _dlg_add_combo(dm, u"ValueCtrl_%s" % ctl_suffix, m, ctrl_y, dw - 2 * m, edit_h)
        else:
            _dlg_add_edit(
                dm,
                u"ValueCtrl_%s" % ctl_suffix,
                m,
                ctrl_y,
                dw - 2 * m,
                edit_h,
                multiline=meta[u"multiline"],
            )
        try:
            apply_edit_font(dm.getByName(u"ValueCtrl_%s" % ctl_suffix), font_pt)
        except Exception:
            pass

    footer_y = inner_dialog_footer_y(dh, m)
    _dlg_add_fixed(dm, u"ErrLbl", u"", m, footer_y - 16, dw - 2 * m, 14)
    inner_dialog_add_ok_cancel_footer(
        dm, m, dh, ok_label=u"OK", cancel_label=u"Отмена", help_btn=True, dw=dw,
        help_key=u"ручной_ввод",
    )

    dlg = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    prepare_dialog_soft_gray_manual_input(dlg, title_text=title)
    if not _create_peer(dlg, toolkit, doc=doc):
        return _headless_reply()

    try:
        help_btn = dlg.getControl(u"HelpButton")
    except Exception:
        help_btn = None
    if help_btn is not None:
        help_body = u""
        try:
            import libre_macros_param_wizard_cfg as _pw_cfg

            help_body = unicode(
                (_pw_cfg.WIZARD_HINTS or {}).get(u"Ручной_ввод_в_карту_переменных", u"")
                or u""
            )
        except Exception:
            help_body = u""
        if help_body == u"":
            help_body = (
                u"Введите значения переменных для runtime-карты.\n"
                u"Подсказки сверху — от первого блока в JSON."
            )

        class _HelpListener(XActionListener, XEventListener):
            def disposing(self, ev):
                pass

            def actionPerformed(self, ev):
                try:
                    from com.sun.star.awt.MessageBoxType import MESSAGEBOX
                    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK

                    peer = dlg.getPeer()
                    mb = toolkit.createMessageBox(
                        peer, MESSAGEBOX, BUTTONS_OK, u"Справка", help_body[:1500]
                    )
                    mb.execute()
                except Exception:
                    pass

        try:
            help_btn.setActionCommand(u"HelpButton")
            help_btn.addActionListener(_HelpListener())
        except Exception:
            pass

    fi = 0
    while fi < len(metas):
        meta = metas[fi]
        default_text = defaults[fi]
        ctl_suffix = unicode(fi)
        try:
            ctl = dlg.getControl(u"ValueCtrl_%s" % ctl_suffix)
            if meta[u"use_combo"]:
                try:
                    ctl.removeItems(0, ctl.getItemCount())
                except Exception:
                    pass
                try:
                    ctl.addItems(tuple(meta[u"choices"]), 0)
                except Exception:
                    pass
                try:
                    ctl.setText(
                        default_text
                        if default_text != u""
                        else (meta[u"choices"][0] if meta[u"choices"] else u"")
                    )
                except Exception:
                    pass
            else:
                try:
                    ctl.setText(default_text)
                except Exception:
                    pass
        except Exception:
            pass
        fi = fi + 1

    while True:
        rc = dlg.execute()
        try:
            rc = int(rc)
        except Exception:
            rc = 0
        if rc != 1:
            try:
                dlg.dispose()
            except Exception:
                pass
            return (False, None)

        entries = []
        err_msg = u""
        raw_vals = []
        fi = 0
        while fi < len(norm_blocks):
            blk = norm_blocks[fi]
            meta = metas[fi]
            ctl_suffix = unicode(fi)
            raw_val = u""
            try:
                raw_val = unicode(dlg.getControl(u"ValueCtrl_%s" % ctl_suffix).getText() or u"")
            except Exception:
                raw_val = u""
            raw_vals.append(raw_val)
            if meta[u"value_kind"] == u"list":
                text = _normalize_list_text(raw_val)
            else:
                text = unicode(raw_val).strip()
            if text == u"" and meta[u"required"]:
                err_msg = u"Заполните обязательное поле «%s»" % meta[u"label"]
                break
            ok, norm, token, verr = _validate_typed_value(
                text, meta[u"data_type"], meta[u"fmt_hint"]
            )
            if not ok:
                err_msg = u"%s: %s" % (meta[u"label"], verr or u"неверное значение")
                break
            out_text = norm
            out_tok = u"text" if meta[u"value_kind"] == u"list" else token
            entries.append((meta[u"name"], out_text, out_tok))
            fi = fi + 1

        if err_msg != u"":
            try:
                dlg.dispose()
            except Exception:
                pass
            fi = 0
            while fi < len(norm_blocks):
                blk2 = dict(norm_blocks[fi])
                meta = metas[fi]
                raw_val = raw_vals[fi] if fi < len(raw_vals) else u""
                if meta[u"value_kind"] == u"list":
                    blk2[u"default"] = _normalize_list_text(raw_val)
                else:
                    blk2[u"default"] = unicode(raw_val).strip()
                norm_blocks[fi] = blk2
                fi = fi + 1
            first2 = dict(norm_blocks[0])
            prev_prompt = unicode(first2.get(u"prompt") or u"").strip()
            first2[u"prompt"] = (u"%s\n%s" % (err_msg, prev_prompt)).strip()
            norm_blocks[0] = first2
            return show_manual_variable_input_combined_dialog(norm_blocks, doc=doc)

        fi = 0
        while fi < len(norm_blocks):
            try:
                from libre_macros_manual_var_store_lib import save_dialog_answer

                save_dialog_answer(norm_blocks[fi], sheet_name, entries[fi][1])
            except Exception as err:
                _log(u"save remembered combined[%d]: %s" % (fi, err))
            fi = fi + 1
        try:
            dlg.dispose()
        except Exception:
            pass
        return (True, entries)


def put_manual_variable_into_map(name, text, type_token=u"text"):
    """Записать значение в _MERGE_SOURCE_VARIABLES_MAP с ключами ручной_ввод."""
    try:
        import libre_macros_collect_cfg as _cw_cfg
    except Exception as err:
        _log(u"put cfg: %s" % err)
        return False
    var_name = unicode(name or u"").strip()
    if var_name == u"" or u"~" in var_name:
        return False
    sheet_key = getattr(_cw_cfg, u"MANUAL_VARIABLE_SHEET_KEY", u"ручной_ввод")
    file_key = getattr(_cw_cfg, u"MANUAL_VARIABLE_FILE_KEY", u"ручной_ввод")
    key = u"%s~%s~%s" % (var_name, sheet_key, file_key)
    value_tuple = (unicode(text or u""), unicode(type_token or u"text"))
    _cw_cfg._MERGE_SOURCE_VARIABLES_MAP[key] = value_tuple
    order = getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_ORDER", None)
    if order is None:
        _cw_cfg._MERGE_SOURCE_VARIABLES_ORDER = []
        order = _cw_cfg._MERGE_SOURCE_VARIABLES_ORDER
    if key not in order:
        order.append(key)
    return True
