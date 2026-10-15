# -*- coding: utf-8 -*-
"""Полноэкранный визард «Все параметры»: ListBox слева + детали B/C/D справа."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.780"
import uno
import unohelper
from com.sun.star.awt import XActionListener, XItemListener

import libre_macros_all_params_wizard_cfg as _acfg

try:
    from libre_macros_ui_theme import (
        DARK_BLUE_TITLE_BG,
        DARK_BLUE_TITLE_FG,
        apply_soft_gray_dark_blue_theme,
        prepare_dialog_soft_gray,
        try_paint_titlebar,
    )
except Exception:
    DARK_BLUE_TITLE_BG = 0x0D47A1
    DARK_BLUE_TITLE_FG = 0xFFFFFF

    def apply_soft_gray_dark_blue_theme(dm, title_text=None):
        return False

    def prepare_dialog_soft_gray(dlg, title_text=None, title_bg=None, title_fg=None):
        return False

    def try_paint_titlebar(dlg, title_bg=None, title_fg=None):
        return False

try:
    unicode
except NameError:
    unicode = str


def _ensure_scripts_python_on_path():
    import os
    import sys

    try:
        here = os.path.dirname(os.path.abspath(__file__))
        scripts = os.path.dirname(here)
        if scripts and os.path.isdir(scripts) and scripts not in sys.path:
            sys.path.insert(0, scripts)
    except Exception:
        pass


def _pw():
    _ensure_scripts_python_on_path()
    import param_wizard as pw

    return pw


def _log(msg):
    try:
        import sys

        sys.stdout.write(u"[all_params] %s\n" % unicode(msg))
        sys.stdout.flush()
    except Exception:
        pass


def _script_context():
    """XSCRIPTCONTEXT из entry-макроса (__main__/builtins) — в pythonpath его нет."""
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


def _screen_work_area(toolkit):
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
    return 0, 0, 1024, 768


def _title_h(dm=None):
    try:
        from libre_macros_ui_theme import SOFT_GRAY_TITLE_H

        return int(SOFT_GRAY_TITLE_H or 28)
    except Exception:
        return 28


def _content_height(sh, dm=None):
    return max(400, int(sh) - _title_h(dm))


def _list_box_height(list_top, status_y, gap=6):
    bottom = int(status_y) - int(gap)
    h = bottom - int(list_top)
    min_h = int(getattr(_acfg, "ALL_PARAMS_LIST_MIN_H", 80))
    if h < min_h:
        h = min_h
    return h


def _add_label(dm, name, label, x, y, w, h, multiline=False):
    ctl = dm.createInstance(u"com.sun.star.awt.UnoControlFixedTextModel")
    ctl.Name = name
    ctl.Label = unicode(label or u"")
    ctl.PositionX = int(x)
    ctl.PositionY = int(y)
    ctl.Width = int(w)
    ctl.Height = int(h)
    if multiline:
        try:
            ctl.MultiLine = True
        except Exception:
            pass
    dm.insertByName(name, ctl)
    return ctl


def _add_button(dm, name, label, x, y, w, h, default=False, bold=False, font_h=None):
    btn = dm.createInstance(u"com.sun.star.awt.UnoControlButtonModel")
    btn.Name = name
    btn.Label = unicode(label or u"")
    btn.PositionX = int(x)
    btn.PositionY = int(y)
    btn.Width = int(w)
    btn.Height = int(h)
    if default:
        try:
            btn.DefaultButton = True
        except Exception:
            pass
    if bold or font_h is not None:
        try:
            fd = btn.FontDescriptor
            if font_h is not None:
                fd.Height = int(font_h)
            if bold:
                try:
                    from com.sun.star.awt.FontWeight import BOLD

                    fd.Weight = BOLD
                except Exception:
                    fd.Weight = 150
            btn.FontDescriptor = fd
        except Exception:
            if font_h is not None:
                try:
                    btn.FontHeight = int(font_h)
                except Exception:
                    pass
    dm.insertByName(name, btn)
    return btn


def _add_edit(dm, name, x, y, w, h, multiline=False, readonly=False):
    ed = dm.createInstance(u"com.sun.star.awt.UnoControlEditModel")
    ed.Name = name
    ed.PositionX = int(x)
    ed.PositionY = int(y)
    ed.Width = int(w)
    ed.Height = int(h)
    if multiline:
        try:
            ed.MultiLine = True
            ed.VScroll = True
            ed.HScroll = False
        except Exception:
            pass
    if readonly:
        try:
            ed.ReadOnly = True
        except Exception:
            pass
    try:
        from libre_macros_ui_theme import apply_edit_font

        apply_edit_font(ed)
    except Exception:
        pass
    dm.insertByName(name, ed)
    return ed


def _add_combo(dm, name, x, y, w, h, items=None, editable=True, font_h=None):
    """ComboBox: список + ручной ввод (editable=True)."""
    cb = dm.createInstance(u"com.sun.star.awt.UnoControlComboBoxModel")
    cb.Name = name
    cb.PositionX = int(x)
    cb.PositionY = int(y)
    cb.Width = int(w)
    cb.Height = int(h)
    try:
        cb.Dropdown = True
        cb.LineCount = 12
    except Exception:
        pass
    try:
        cb.ReadOnly = not bool(editable)
    except Exception:
        pass
    if items:
        try:
            cb.StringItemList = tuple(unicode(x) for x in items)
        except Exception:
            pass
    try:
        from libre_macros_ui_theme import apply_edit_font

        apply_edit_font(cb, height=font_h)
    except Exception:
        if font_h is not None:
            try:
                cb.FontHeight = int(font_h)
            except Exception:
                pass
    dm.insertByName(name, cb)
    return cb


def _set_combo_items(dialog, name, items, select_text=None):
    try:
        ctl = dialog.getControl(name)
        mdl = ctl.Model
        mdl.StringItemList = tuple(unicode(x) for x in (items or ()))
    except Exception:
        return
    if select_text is not None:
        _set_text(dialog, name, unicode(select_text or u""))


def _add_listbox(dm, name, x, y, w, h):
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


def _add_checkbox(dm, name, label, x, y, w, h):
    chk = dm.createInstance(u"com.sun.star.awt.UnoControlCheckBoxModel")
    chk.Name = name
    chk.Label = unicode(label or u"")
    chk.PositionX = int(x)
    chk.PositionY = int(y)
    chk.Width = int(w)
    chk.Height = int(h)
    dm.insertByName(name, chk)
    return chk


def _set_text(dialog, name, text):
    try:
        dialog.getControl(name).setText(unicode(text or u""))
    except Exception:
        try:
            dialog.getControl(name).Model.Text = unicode(text or u"")
        except Exception:
            pass


def _get_text(dialog, name):
    try:
        return unicode(dialog.getControl(name).getText() or u"")
    except Exception:
        try:
            return unicode(dialog.getControl(name).Model.Text or u"")
        except Exception:
            return u""


def _set_enabled(dialog, name, enabled):
    try:
        dialog.getControl(name).setEnable(bool(enabled))
    except Exception:
        pass


def _peer_fill_workarea(dialog, toolkit):
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


def _create_peer(dialog, toolkit, doc=None, parent_dialog=None):
    if dialog is None or toolkit is None:
        return False
    try:
        prepare_dialog_soft_gray(dialog)
    except Exception:
        pass
    parents = []
    try:
        if doc is not None:
            peer = doc.getCurrentController().getFrame().getContainerWindow()
            if peer is not None:
                parents.append(peer)
    except Exception:
        pass
    if parent_dialog is not None:
        try:
            peer = parent_dialog.getPeer()
            if peer is not None and peer not in parents:
                parents.append(peer)
        except Exception:
            pass
    parents.append(None)
    ok = False
    for p in parents:
        try:
            dialog.createPeer(toolkit, p)
            ok = True
            break
        except Exception:
            pass
    if ok:
        _peer_fill_workarea(dialog, toolkit)
    return ok


def _entry_caption(entry, index, sheet):
    """Подпись шага в ListBox."""
    pw = _pw()
    max_len = int(getattr(_acfg, "ALL_PARAMS_CAPTION_MAX", 120))
    b_max = int(getattr(_acfg, "ALL_PARAMS_CAPTION_B_MAX", 56))
    name = unicode(entry.get(u"sheet_name") or entry.get(u"name") or u"")
    disabled = bool(entry.get(u"disabled"))
    prefix = unicode(getattr(_acfg, "ALL_PARAMS_DISABLED_PREFIX", u"[откл] ")) if disabled else u""
    detail = u""
    try:
        row = int(entry.get(u"row"))
        b = pw._cell_string(sheet.getCellByPosition(1, row))
        if b:
            detail = b.replace(u"\n", u" ").strip()
            if len(detail) > b_max:
                detail = detail[: b_max - 1] + u"…"
    except Exception:
        detail = u""
    body = u"%03d  %s" % (int(index) + 1, name)
    if detail:
        body = u"%s  |  %s" % (body, detail)
    if len(body) > max_len:
        body = body[: max_len - 1] + u"…"
    return prefix + body


def _list_selected_index(dialog):
    try:
        lb = dialog.getControl(u"StepsListBox")
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
        dialog.getControl(u"StepsListBox").selectItemPos(int(idx), True)
    except Exception:
        pass


def _reload_entries(state):
    pw = _pw()
    sheet = state.get(u"sheet")
    catalog = state.get(u"catalog")
    entries = pw._scan_sheet_param_entries(sheet, catalog)
    state[u"entries"] = entries
    return entries


def _detail_snapshot(dialog):
    """Снимок полей панели для проверки dirty без записи на лист."""
    if dialog is None:
        return None
    disabled = 0
    try:
        disabled = int(dialog.getControl(u"DisabledChk").Model.State or 0)
    except Exception:
        disabled = 0
    return (
        _get_text(dialog, u"NameCombo"),
        _get_text(dialog, u"ValueCombo"),
        _get_text(dialog, u"ExtraEd"),
        disabled,
    )


def _detail_is_dirty(state):
    snap = state.get(u"loaded_snap")
    if snap is None:
        return False
    return _detail_snapshot(state.get(u"dialog")) != snap


def _remember_loaded_snap(state):
    state[u"loaded_snap"] = _detail_snapshot(state.get(u"dialog"))


def _left_panel_width(total_w, m, split_gap):
    """Ширина списка шагов: ровно ALL_PARAMS_LIST_WIDTH_RATIO от доступной ширины."""
    avail = max(200, int(total_w) - 2 * int(m) - int(split_gap))
    ratio = float(getattr(_acfg, "ALL_PARAMS_LIST_WIDTH_RATIO", 0.40) or 0.40)
    if ratio < 0.2:
        ratio = 0.2
    if ratio > 0.6:
        ratio = 0.6
    w = int(round(avail * ratio))
    min_w = int(getattr(_acfg, "ALL_PARAMS_LIST_MIN_W", 220))
    if w < min_w:
        w = min(min_w, avail)
    # справа оставляем место под детали
    if w > avail - 200:
        w = max(min_w, avail - 200)
    return w


def _refresh_list(state, keep_row=None, keep_index=-1):
    dialog = state.get(u"dialog")
    sheet = state.get(u"sheet")
    state[u"form_loading"] = True
    try:
        entries = _reload_entries(state)
        caps = []
        i = 0
        while i < len(entries):
            caps.append(_entry_caption(entries[i], i, sheet))
            i = i + 1
        try:
            lb = dialog.getControl(u"StepsListBox")
            lb.Model.StringItemList = tuple(caps)
        except Exception:
            pass
        pick = -1
        if keep_row is not None:
            i = 0
            while i < len(entries):
                if int(entries[i].get(u"row", -1)) == int(keep_row):
                    pick = i
                    break
                i = i + 1
        elif keep_index >= 0:
            pick = min(int(keep_index), len(entries) - 1) if entries else -1
        elif entries:
            pick = 0
        if pick >= 0:
            _list_select_index(dialog, pick)
            state[u"selected"] = pick
            _load_detail(state, pick)
        else:
            state[u"selected"] = -1
            state[u"loaded_snap"] = None
            _clear_detail(state)
        _update_status(state)
    finally:
        state[u"form_loading"] = False


def _update_status(state):
    dialog = state.get(u"dialog")
    entries = state.get(u"entries") or []
    si = int(state.get(u"selected") or -1)
    msg = u"Шагов: %d" % len(entries)
    if 0 <= si < len(entries):
        e = entries[si]
        msg = u"%s  ·  выбрано: %s (стр. %d)" % (
            msg,
            unicode(e.get(u"name") or u""),
            int(e.get(u"row", 0)) + 1,
        )
    try:
        dialog.getControl(u"StatusLbl").Model.Label = msg
    except Exception:
        pass


def _clear_detail(state):
    dialog = state.get(u"dialog")
    if dialog is None:
        return
    was_loading = bool(state.get(u"form_loading"))
    state[u"form_loading"] = True
    try:
        _set_text(dialog, u"NameCombo", u"")
        _set_text(dialog, u"ValueCombo", u"")
        _set_text(dialog, u"ExtraEd", u"")
        _set_text(dialog, u"CondEd", u"")
        _set_text(dialog, u"DescEd", u"")
        try:
            dialog.getControl(u"DisabledChk").Model.State = 0
        except Exception:
            pass
        _set_enabled(dialog, u"NameCombo", False)
        _set_enabled(dialog, u"ValueCombo", False)
        _set_enabled(dialog, u"ExtraParamsBtn", False)
        _set_enabled(dialog, u"CondBtn", False)
        _set_enabled(dialog, u"ApplyBtn", False)
        _set_enabled(dialog, u"DelBtn", False)
        _set_enabled(dialog, u"CopyBtn", False)
        _set_enabled(dialog, u"UpBtn", False)
        _set_enabled(dialog, u"DownBtn", False)
        state[u"loaded_snap"] = None
    finally:
        state[u"form_loading"] = was_loading


def _row_can_open_col_c(pw, spec, primary):
    """Лёгкая проверка кнопки C без сборки субдиалога."""
    if spec is None:
        return False
    mode = spec.get(u"mode")
    if pw._spec_has_col_c_subdialog(spec):
        return True
    if mode in (u"vlookup", u"merge_sheets", u"split_sheets", u"source_extra"):
        return True
    if mode in (u"postprocess_range", u"postprocess_row"):
        try:
            return bool(pw._param_row_has_subdialog(spec, primary))
        except Exception:
            return False
    name = unicode(spec.get(u"name") or u"").strip()
    known = (
        u"Листы",
        u"Столбцы",
        u"удаление_верхних_строк",
        u"Пропуск_строк_источника",
        u"Предварительный_скрипт",
        u"Строка_Заголовков",
        u"Ячейка_В_Столбец",
        u"Добавить_в_карту_переменных",
        u"Ручной_ввод_в_карту_переменных",
        u"Листы_не_удалять",
        u"Листы_удалить",
        u"автофильтр_результата",
        u"закрепить_заголовок",
    )
    i = 0
    while i < len(known):
        try:
            if pw._param_match(name, known[i]):
                return True
        except Exception:
            pass
        i = i + 1
    return False


def _catalog_name_choices(catalog, current_name=None):
    names = []
    seen = set()
    i = 0
    while catalog is not None and i < len(catalog):
        nm = unicode((catalog[i] or {}).get(u"name") or u"").strip()
        if nm and nm.casefold() not in seen:
            names.append(nm)
            seen.add(nm.casefold())
        i = i + 1
    cur = unicode(current_name or u"").strip()
    if cur:
        # без префикса отключения
        try:
            pw = _pw()
            cur = pw._param_canonical_sheet_name(cur)
        except Exception:
            pass
        if cur and cur.casefold() not in seen:
            names.insert(0, cur)
    return names


def _value_b_choices(spec, primary):
    """Варианты для combo B (из choices спецификации + текущее значение)."""
    choices = []
    seen = set()
    raw = (spec or {}).get(u"choices") or ()
    i = 0
    while i < len(raw):
        item = unicode(raw[i] or u"").strip()
        if item and item.casefold() not in seen:
            choices.append(item)
            seen.add(item.casefold())
        i = i + 1
    cur = unicode(primary or u"").strip()
    if cur and cur.casefold() not in seen:
        choices.insert(0, cur)
    return choices


def _condition_text_for_row(sheet, row):
    pw = _pw()
    fn_b = pw._cell_string(sheet.getCellByPosition(1, int(row)))
    is_plugin = pw._pp_is_plugin_primary(fn_b)
    col = 4 if is_plugin else 3
    return pw._cell_string(sheet.getCellByPosition(col, int(row)))


def _load_detail(state, idx):
    dialog = state.get(u"dialog")
    sheet = state.get(u"sheet")
    entries = state.get(u"entries") or []
    if dialog is None or idx < 0 or idx >= len(entries):
        _clear_detail(state)
        return
    pw = _pw()
    entry = entries[idx]
    spec = entry.get(u"spec") or {}
    row = int(entry.get(u"row"))
    was_loading = bool(state.get(u"form_loading"))
    state[u"form_loading"] = True
    try:
        primary, extra, plugin_ref = pw._read_values_at_row(sheet, spec, row)
        name = unicode(entry.get(u"sheet_name") or entry.get(u"name") or u"")
        try:
            name_show = pw._param_canonical_sheet_name(name)
        except Exception:
            name_show = name
        _set_combo_items(
            dialog,
            u"NameCombo",
            _catalog_name_choices(state.get(u"catalog"), name_show),
            select_text=name_show,
        )
        _set_combo_items(
            dialog,
            u"ValueCombo",
            _value_b_choices(spec, primary),
            select_text=unicode(primary or u""),
        )
        # Колонка C: для постобработки — extra (JSON); для плагина JSON/extra
        # уже в extra; plugin_ref — отдельное поле C в модели чтения.
        mode = spec.get(u"mode")
        if mode in (u"postprocess_range", u"postprocess_row") and pw._pp_is_plugin_primary(primary):
            c_text = unicode(plugin_ref or u"")
            d_extra = unicode(extra or u"")
            if d_extra:
                c_text = c_text + (u"\n---\n" + d_extra if c_text else d_extra)
            _set_text(dialog, u"ExtraEd", c_text)
        else:
            _set_text(dialog, u"ExtraEd", unicode(extra or u""))
        cond = u""
        can_cond = mode in (u"postprocess_range", u"postprocess_row")
        if can_cond:
            cond = _condition_text_for_row(sheet, row)
        _set_text(dialog, u"CondEd", cond)
        hint = unicode(spec.get(u"hint") or u"")
        _set_text(dialog, u"DescEd", hint)
        try:
            dialog.getControl(u"DisabledChk").Model.State = 1 if entry.get(u"disabled") else 0
        except Exception:
            pass
        can_c = _row_can_open_col_c(pw, spec, primary)
        # Для постобработки/ВПР кнопка C всегда активна, если в B есть функция —
        # сам визард откроется через _show_inner_codec_param_dialog.
        if (not can_c) and mode in (
            u"postprocess_range",
            u"postprocess_row",
            u"vlookup",
            u"merge_sheets",
            u"split_sheets",
        ):
            can_c = unicode(primary or u"").strip() != u""
        _set_enabled(dialog, u"NameCombo", True)
        _set_enabled(dialog, u"ValueCombo", True)
        _set_enabled(dialog, u"ExtraParamsBtn", can_c)
        _set_enabled(dialog, u"CondBtn", can_cond and unicode(primary or u"").strip() != u"")
        _set_enabled(dialog, u"ApplyBtn", True)
        _set_enabled(dialog, u"DelBtn", True)
        _set_enabled(dialog, u"CopyBtn", True)
        _set_enabled(dialog, u"UpBtn", idx > 0)
        _set_enabled(dialog, u"DownBtn", idx < len(entries) - 1)
        _remember_loaded_snap(state)
    finally:
        state[u"form_loading"] = was_loading


def _save_detail_to_sheet(state, force=False):
    """Записать B/C (и disabled) из панели в текущую строку листа.
    Без force пишет только если панель реально менялась (иначе навигация тормозит).
    """
    if state.get(u"form_loading"):
        return True
    if (not force) and (not _detail_is_dirty(state)):
        return True
    dialog = state.get(u"dialog")
    sheet = state.get(u"sheet")
    catalog = state.get(u"catalog")
    doc = state.get(u"doc")
    entries = state.get(u"entries") or []
    si = int(state.get(u"selected") or -1)
    if dialog is None or sheet is None or si < 0 or si >= len(entries):
        return False
    pw = _pw()
    entry = entries[si]
    spec = entry.get(u"spec") or {}
    row = int(entry.get(u"row"))
    name_raw = _get_text(dialog, u"NameCombo").strip()
    try:
        name_canon = pw._param_canonical_sheet_name(name_raw) if name_raw else unicode(entry.get(u"name") or u"")
    except Exception:
        name_canon = name_raw or unicode(entry.get(u"name") or u"")
    if name_canon:
        found = pw._find_catalog_spec(catalog, name_canon)
        if found is not None:
            spec = found
            entry[u"spec"] = found
            entry[u"name"] = found.get(u"name") or name_canon
    primary = _get_text(dialog, u"ValueCombo").strip()
    extra = _get_text(dialog, u"ExtraEd").strip()
    plugin_ref = None
    mode = spec.get(u"mode")
    if mode in (u"postprocess_range", u"postprocess_row") and pw._pp_is_plugin_primary(primary):
        # ExtraEd: ref и опционально --- + D extra
        parts = extra.split(u"\n---\n", 1)
        plugin_ref = parts[0].strip()
        extra = parts[1].strip() if len(parts) > 1 else u""
    try:
        disabled = bool(dialog.getControl(u"DisabledChk").Model.State)
    except Exception:
        disabled = bool(entry.get(u"disabled"))
    try:
        if mode == u"vlookup":
            # B не редактируем текстом — только через визард C
            pass
        elif mode in (u"merge_sheets", u"split_sheets"):
            pw.apply_param_value(
                sheet, spec, extra, None, target_row=row, target_col=2, catalog=catalog, doc=doc,
            )
        elif pw._spec_has_col_c_subdialog(spec) and (not pw._spec_col_c_keeps_b(spec)) and mode != u"choice":
            import libre_macros_param_wizard_cfg as _pw_cfg

            pw.apply_param_value(
                sheet, spec, _pw_cfg._PARAM_B_SEE_COLUMN_C, extra,
                target_row=row, catalog=catalog, doc=doc,
            )
        else:
            pw.apply_param_value(
                sheet, spec, primary, extra, target_row=row, catalog=catalog, doc=doc,
                plugin_ref=plugin_ref,
            )
        save_name = unicode(spec.get(u"name") or name_canon or entry.get(u"name") or u"")
        pw._write_param_row_sheet_name(sheet, row, save_name, disabled, catalog=catalog)
    except Exception as exc:
        try:
            pw._show_message(u"Не удалось сохранить строку:\n%s" % exc, dialog=dialog, doc=doc)
        except Exception:
            pass
        return False
    _remember_loaded_snap(state)
    return True


def _suspend_all_params_dialog(state):
    """
    Полноэкранный all_params как owner ломает ComboBox во вложенных визардах
    (списки не открываются) и может прятать диалог C за собой.
    На время субвизарда отключаем owner и не передаём его в createPeer.
    """
    dialog = state.get(u"dialog")
    if dialog is None:
        return None
    try:
        dialog.setEnable(False)
    except Exception:
        pass
    return dialog


def _resume_all_params_dialog(dialog):
    if dialog is None:
        return
    try:
        dialog.setEnable(True)
    except Exception:
        pass
    try:
        dialog.setFocus()
    except Exception:
        pass
    try:
        dialog.toFront()
    except Exception:
        pass


def _open_col_c_wizard(state):
    """
    Визард колонки C — как ColC/PivotParams в основном визарде.
    parent_dialog=None: peer к окну документа, иначе combo/форма не работают.
    """
    dialog = state.get(u"dialog")
    sheet = state.get(u"sheet")
    catalog = state.get(u"catalog")
    doc = state.get(u"doc")
    entries = state.get(u"entries") or []
    si = int(state.get(u"selected") or -1)
    if si < 0 or si >= len(entries):
        return
    _save_detail_to_sheet(state, force=True)
    pw = _pw()
    entry = entries[si]
    spec = entry.get(u"spec") or {}
    row = int(entry.get(u"row"))
    primary = _get_text(dialog, u"ValueCombo").strip()
    initial_text = _get_text(dialog, u"ExtraEd").strip()
    mode = spec.get(u"mode")
    # Не передавать fullscreen all_params как owner.
    nest_parent = None

    host = _suspend_all_params_dialog(state)
    err_msg = None
    info_msg = None
    try:
        # 1) Спецификация с явным col_c_subdialog (Листы, pre_shell, …).
        sub = pw._col_c_subdialog_for_row(sheet, catalog, row)
        if sub is not None and sub.get(u"kind") not in (u"inner_codec", u"pivot", u"plugin"):
            ok = pw._run_subdialog_and_save(
                doc, sheet, catalog, sub, parent_dialog=nest_parent,
            )
            if ok:
                _refresh_list(state, keep_row=row)
            return

        # 2) Постобработка / финал / плагин — как кнопка «Параметры…» в main.
        if mode in (u"postprocess_range", u"postprocess_row") or mode == u"vlookup":
            fn_key = primary
            if fn_key == u"" and sub is not None:
                fn_key = unicode(sub.get(u"fn_key") or u"").strip()
            if fn_key == u"":
                try:
                    fn_key = pw._cell_string(sheet.getCellByPosition(1, row)).strip()
                except Exception:
                    fn_key = u""
            if fn_key == u"":
                info_msg = u"Выберите функцию в поле «Значение (B)»."
                return
            if initial_text == u"":
                try:
                    initial_text = pw._cell_string(sheet.getCellByPosition(2, row))
                except Exception:
                    initial_text = u""
            force_av = False
            try:
                force_av = bool(pw._spec_is_postprocess_xml(spec))
            except Exception:
                force_av = False
            try:
                if pw._pp_extra_needs_json_seed(fn_key, initial_text):
                    initial_text = pw._pp_minimal_json_for_fn(fn_key, force_as_values=force_av)
            except Exception:
                pass
            result = None
            if pw._pp_is_plugin_primary(fn_key):
                plugin_sub = sub if (sub and sub.get(u"kind") == u"plugin") else {
                    u"kind": u"plugin",
                    u"spec": spec,
                    u"row": row,
                    u"initial_ref": initial_text,
                    u"initial_extra": pw._cell_string(sheet.getCellByPosition(3, row)),
                }
                ok = pw._run_subdialog_and_save(
                    doc, sheet, catalog, plugin_sub, parent_dialog=nest_parent,
                )
                if ok:
                    _refresh_list(state, keep_row=row)
                return
            if pw._is_pivot_table_pp_name(fn_key):
                result = pw.show_pivot_table_param_dialog(
                    nest_parent,
                    initial_text,
                    doc=doc,
                    param_sheet=sheet,
                    param_row=row,
                    param_primary=fn_key,
                    force_as_values=force_av,
                )
                if result is not None:
                    try:
                        from libre_macros_param_codec import param_decode, param_encode

                        blocks = param_decode(u"сводная_таблица", result)
                        result = param_encode(u"сводная_таблица", blocks)
                    except Exception:
                        pass
            else:
                result = pw._show_inner_codec_param_dialog(
                    nest_parent, fn_key, initial_text, doc=doc, force_as_values=force_av,
                )
            if result is None:
                return
            _set_text(dialog, u"ExtraEd", unicode(result or u"").strip())
            if _save_detail_to_sheet(state, force=True):
                state[u"changed"] = True
                _refresh_list(state, keep_row=row)
            return

        # 3) Fallback: то, что нашёл _col_c_subdialog_for_row.
        if sub is not None:
            ok = pw._run_subdialog_and_save(
                doc, sheet, catalog, sub, parent_dialog=nest_parent,
            )
            if ok:
                _refresh_list(state, keep_row=row)
            return
        info_msg = u"Для этой строки нет визарда колонки C."
    except Exception as exc:
        err_msg = unicode(exc)
    finally:
        _resume_all_params_dialog(host)
    if info_msg:
        try:
            pw._show_message(info_msg, dialog=dialog, doc=doc, title=u"Параметры C")
        except Exception:
            pass
    if err_msg:
        try:
            pw._show_message(
                u"Ошибка визарда колонки C:\n%s" % err_msg,
                dialog=dialog,
                doc=doc,
                title=u"Параметры C",
            )
        except Exception:
            pass


def _open_col_d_wizard(state):
    dialog = state.get(u"dialog")
    sheet = state.get(u"sheet")
    catalog = state.get(u"catalog")
    doc = state.get(u"doc")
    entries = state.get(u"entries") or []
    si = int(state.get(u"selected") or -1)
    if si < 0 or si >= len(entries):
        return
    _save_detail_to_sheet(state, force=True)
    pw = _pw()
    row = int(entries[si].get(u"row"))
    sub = pw._col_d_condition_subdialog_for_row(sheet, catalog, row)
    if sub is None:
        try:
            pw._show_message(
                u"Условие задаётся для шагов постобработки (колонка D/E).",
                dialog=dialog,
                doc=doc,
                title=u"Условие",
            )
        except Exception:
            pass
        return
    # parent=None: иначе ComboBox условия/итераций не раскрываются поверх fullscreen.
    host = _suspend_all_params_dialog(state)
    err_msg = None
    try:
        ok = pw._run_subdialog_and_save(
            doc, sheet, catalog, sub, parent_dialog=None,
        )
        if ok:
            _refresh_list(state, keep_row=row)
    except Exception as exc:
        err_msg = unicode(exc)
    finally:
        _resume_all_params_dialog(host)
    if err_msg:
        try:
            pw._show_message(
                u"Ошибка визарда условия:\n%s" % err_msg,
                dialog=dialog,
                doc=doc,
                title=u"Условие",
            )
        except Exception:
            pass


def _delete_selected(state):
    dialog = state.get(u"dialog")
    sheet = state.get(u"sheet")
    catalog = state.get(u"catalog")
    doc = state.get(u"doc")
    entries = state.get(u"entries") or []
    si = int(state.get(u"selected") or -1)
    if si < 0 or si >= len(entries):
        return
    pw = _pw()
    entry = entries[si]
    row = int(entry.get(u"row"))
    pname = unicode(entry.get(u"name") or u"")
    if not pw._wp_confirm_delete_param_row(pname, row, doc=doc, dialog=dialog):
        return
    pw._clear_param_sheet_row(sheet, row, catalog, doc)
    next_idx = si if si < len(entries) - 1 else si - 1
    _refresh_list(state, keep_index=max(0, next_idx))


def _copy_selected(state):
    dialog = state.get(u"dialog")
    sheet = state.get(u"sheet")
    catalog = state.get(u"catalog")
    doc = state.get(u"doc")
    entries = state.get(u"entries") or []
    si = int(state.get(u"selected") or -1)
    if si < 0 or si >= len(entries):
        return
    _save_detail_to_sheet(state)
    pw = _pw()
    entry = entries[si]
    spec = entry.get(u"spec") or {}
    if not spec.get(u"multi_row"):
        try:
            pw._show_message(
                u"Параметр «%s» должен быть в одном экземпляре — копирование невозможно."
                % unicode(entry.get(u"name") or u""),
                dialog=dialog,
                doc=doc,
            )
        except Exception:
            pass
        return
    row = int(entry.get(u"row"))
    try:
        from libre_macros_copy_param_rows_lib import (
            _clear_pp_params_on_row,
            _confirm_copy_dialog,
            _copy_end_col,
            _count_free_rows,
            _is_pp_or_final_spec,
            _restore_header_after_copy,
        )
    except Exception as exc:
        try:
            pw._show_message(u"Копирование недоступно: %s" % exc, dialog=dialog, doc=doc)
        except Exception:
            pass
        return
    n_pp = 1 if _is_pp_or_final_spec(spec) else 0
    ok, keep_pp = _confirm_copy_dialog(doc, 1, n_pp)
    if not ok:
        return
    insert_at = row + 1
    end_col = _copy_end_col(pw, sheet, catalog, [row])
    free = _count_free_rows(sheet, insert_at, 1, end_col)
    pw._param_sheet_begin_write(sheet)
    try:
        if free < 1:
            sheet.getRows().insertByIndex(insert_at, 1)
        pw._copy_sheet_row(sheet, row, insert_at, end_col)
        if (not keep_pp) and n_pp > 0:
            _clear_pp_params_on_row(sheet, insert_at, end_col)
        _restore_header_after_copy(pw, sheet, catalog, [row])
    except Exception as exc:
        try:
            pw._show_message(u"Ошибка копирования: %s" % exc, dialog=dialog, doc=doc)
        except Exception:
            pass
        return
    finally:
        try:
            pw._param_sheet_end_write(sheet)
        except Exception:
            pass
    _refresh_list(state, keep_row=insert_at)


def _move_selected(state, direction):
    dialog = state.get(u"dialog")
    sheet = state.get(u"sheet")
    doc = state.get(u"doc")
    entries = state.get(u"entries") or []
    si = int(state.get(u"selected") or -1)
    if si < 0 or si >= len(entries):
        return
    _save_detail_to_sheet(state)
    pw = _pw()
    row = int(entries[si].get(u"row"))
    ok, err, new_start, _new_end, _ts, _te = pw._merge_param_move_rows(
        doc, sheet, row, row, direction, min_row=1,
    )
    if not ok:
        try:
            pw._show_message(err or u"Не удалось переместить.", dialog=dialog, doc=doc)
        except Exception:
            pass
        return
    _refresh_list(state, keep_row=new_start)


def _detail_chrome_height(combo_h, apply_h, lbl_h, gap, sec, btn_row_h):
    """Высота подписей/отступов/кнопок без полей C/D/описание (B — фикс. combo)."""
    # DetailLbl + NameLbl + NameCombo + DisabledChk
    # + ValueLbl + ValueCombo
    # + ExtraLbl + CondLbl + DescLbl
    # + ряд кнопок C/D + Apply (зарезервирован отдельно в avail)
    return (
        (lbl_h + 2) * 2
        + combo_h
        + gap
        + lbl_h
        + sec
        + (lbl_h + 2)
        + combo_h
        + sec
        + (lbl_h + 2)
        + sec
        + btn_row_h
        + gap
        + (lbl_h + 2)
        + sec
        + (lbl_h + 2)
        + gap
        + apply_h
    )


def _detail_field_heights(avail_h):
    """Распределить высоту между C / D / описанием (B — combo; Apply — фикс снизу)."""
    combo_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_COMBO_H", 18))
    apply_h = int(getattr(_acfg, "ALL_PARAMS_APPLY_BTN_H", 36))
    lbl_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_LBL_H", 14))
    gap = int(getattr(_acfg, "ALL_PARAMS_DETAIL_GAP", 4))
    sec = int(getattr(_acfg, "ALL_PARAMS_DETAIL_SECTION_GAP", 6))
    btn_row_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_BTN_ROW_H", 30))
    chrome = _detail_chrome_height(combo_h, apply_h, lbl_h, gap, sec, btn_row_h)
    rest = max(80, int(avail_h) - chrome)
    emin = int(getattr(_acfg, "ALL_PARAMS_DETAIL_EXTRA_MIN_H", 45))
    cmin = int(getattr(_acfg, "ALL_PARAMS_DETAIL_COND_MIN_H", 44))
    dmin = int(getattr(_acfg, "ALL_PARAMS_DETAIL_DESC_MIN_H", 56))
    need = emin + cmin + dmin
    if rest <= need:
        scale = float(rest) / float(need) if need > 0 else 1.0
        eh = max(28, int(emin * scale))
        ch = max(24, int(cmin * scale))
        dh = max(24, int(dmin * scale))
        s = eh + ch + dh
        eh = max(28, eh + (rest - s))
        return eh, ch, dh
    free = rest - need
    return (
        emin + int(free * float(getattr(_acfg, "ALL_PARAMS_DETAIL_EXTRA_SHARE", 0.38))),
        cmin + int(free * float(getattr(_acfg, "ALL_PARAMS_DETAIL_COND_SHARE", 0.28))),
        dmin + int(free * float(getattr(_acfg, "ALL_PARAMS_DETAIL_DESC_SHARE", 0.34))),
    )


def _create_detail_panel(dm, rx, ry, rw):
    combo_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_COMBO_H", 18))
    lbl_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_LBL_H", 14))
    gap = int(getattr(_acfg, "ALL_PARAMS_DETAIL_GAP", 4))
    sec = int(getattr(_acfg, "ALL_PARAMS_DETAIL_SECTION_GAP", 6))
    btn_w = int(getattr(_acfg, "ALL_PARAMS_DETAIL_BTN_W", 130))
    btn_row_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_BTN_ROW_H", 30))
    btn_font = int(getattr(_acfg, "ALL_PARAMS_DETAIL_BTN_FONT_H", 11))
    apply_h = int(getattr(_acfg, "ALL_PARAMS_APPLY_BTN_H", 36))
    apply_w = int(getattr(_acfg, "ALL_PARAMS_APPLY_BTN_W", 200))
    name_font = int(getattr(_acfg, "ALL_PARAMS_NAME_FONT_H", 12))
    value_font = int(getattr(_acfg, "ALL_PARAMS_VALUE_FONT_H", 10))
    extra_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_EXTRA_MIN_H", 45))
    cond_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_COND_MIN_H", 44))
    desc_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_DESC_MIN_H", 56))
    y = int(ry)
    _add_label(dm, u"DetailLbl", u"Параметры шага:", rx, y, rw, lbl_h)
    y += lbl_h + 2
    _add_label(dm, u"NameLbl", u"Имя (A):", rx, y, rw, lbl_h)
    y += lbl_h + 2
    _add_combo(dm, u"NameCombo", rx, y, rw, combo_h, editable=True, font_h=name_font)
    y += combo_h + gap
    _add_checkbox(dm, u"DisabledChk", u"Отключить параметр (__ в колонке A)", rx, y, rw, lbl_h)
    y += lbl_h + sec
    _add_label(dm, u"ValueLbl", u"Значение (B):", rx, y, rw, lbl_h)
    y += lbl_h + 2
    _add_combo(dm, u"ValueCombo", rx, y, rw, combo_h, editable=True, font_h=value_font)
    y += combo_h + sec
    c_btn = unicode(getattr(_acfg, "ALL_PARAMS_BTN_C_LABEL", u"Параметры C…"))
    d_btn = unicode(getattr(_acfg, "ALL_PARAMS_BTN_D_LABEL", u"Параметры D…"))
    _add_label(dm, u"ExtraLbl", u"Колонка C (JSON / доп. аргументы):", rx, y, rw, lbl_h)
    y += lbl_h + 2
    _add_edit(dm, u"ExtraEd", rx, y, rw, extra_h, multiline=True)
    y += extra_h + sec
    # Кнопки C и D — выше, жирный шрифт, между полями C и D.
    _add_button(
        dm, u"ExtraParamsBtn", c_btn, rx, y, btn_w, btn_row_h, bold=True, font_h=btn_font,
    )
    _add_button(
        dm, u"CondBtn", d_btn, rx + btn_w + gap, y, btn_w, btn_row_h, bold=True, font_h=btn_font,
    )
    y += btn_row_h + gap
    _add_label(dm, u"CondLbl", u"Колонка D/E (условие шага):", rx, y, rw, lbl_h)
    y += lbl_h + 2
    _add_edit(dm, u"CondEd", rx, y, rw, cond_h, multiline=True, readonly=True)
    y += cond_h + sec
    _add_label(dm, u"DescLbl", u"Описание:", rx, y, rw, lbl_h)
    y += lbl_h + 2
    _add_edit(dm, u"DescEd", rx, y, rw, desc_h, multiline=True, readonly=True)
    y += desc_h + gap
    _add_button(
        dm,
        u"ApplyBtn",
        u"Применить к строке",
        rx,
        y,
        min(apply_w, rw),
        apply_h,
        bold=True,
        font_h=btn_font,
    )
    return y + apply_h


def _layout_detail_panel(dm, rx, ry, rw, bottom_y):
    """A/B — combo; C −20%; кнопки C/D между C и D; Apply прижата к низу."""
    combo_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_COMBO_H", 18))
    lbl_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_LBL_H", 14))
    gap = int(getattr(_acfg, "ALL_PARAMS_DETAIL_GAP", 4))
    sec = int(getattr(_acfg, "ALL_PARAMS_DETAIL_SECTION_GAP", 6))
    btn_w = int(getattr(_acfg, "ALL_PARAMS_DETAIL_BTN_W", 130))
    btn_row_h = int(getattr(_acfg, "ALL_PARAMS_DETAIL_BTN_ROW_H", 30))
    apply_h = int(getattr(_acfg, "ALL_PARAMS_APPLY_BTN_H", 36))
    apply_w = int(getattr(_acfg, "ALL_PARAMS_APPLY_BTN_W", 200))
    avail = max(120, int(bottom_y) - int(ry))
    extra_h, cond_h, desc_h = _detail_field_heights(avail)
    y = int(ry)
    specs = (
        (u"DetailLbl", rx, y, rw, lbl_h),
    )
    y += lbl_h + 2
    specs += ((u"NameLbl", rx, y, rw, lbl_h),)
    y += lbl_h + 2
    specs += ((u"NameCombo", rx, y, rw, combo_h),)
    y += combo_h + gap
    specs += ((u"DisabledChk", rx, y, rw, lbl_h),)
    y += lbl_h + sec
    specs += ((u"ValueLbl", rx, y, rw, lbl_h),)
    y += lbl_h + 2
    specs += ((u"ValueCombo", rx, y, rw, combo_h),)
    y += combo_h + sec
    specs += ((u"ExtraLbl", rx, y, rw, lbl_h),)
    y += lbl_h + 2
    specs += ((u"ExtraEd", rx, y, rw, extra_h),)
    y += extra_h + sec
    specs += (
        (u"ExtraParamsBtn", rx, y, btn_w, btn_row_h),
        (u"CondBtn", rx + btn_w + gap, y, btn_w, btn_row_h),
    )
    y += btn_row_h + gap
    specs += ((u"CondLbl", rx, y, rw, lbl_h),)
    y += lbl_h + 2
    specs += ((u"CondEd", rx, y, rw, cond_h),)
    y += cond_h + sec
    specs += ((u"DescLbl", rx, y, rw, lbl_h),)
    y += lbl_h + 2
    specs += ((u"DescEd", rx, y, rw, desc_h),)
    # «Применить» — к нижней кромке свободной области.
    apply_y = int(bottom_y) - apply_h
    if apply_y < y + gap:
        apply_y = y + gap
    specs += ((u"ApplyBtn", rx, apply_y, min(apply_w, rw), apply_h),)
    for name, x, yy, w, h in specs:
        try:
            ctl = dm.getByName(name)
            ctl.PositionX = int(x)
            ctl.PositionY = int(yy)
            ctl.Width = int(w)
            ctl.Height = int(h)
        except Exception:
            pass


def _sync_control_layout(dialog, dm):
    """Как todo plan: после createPeer/fill применить PosSize из модели ко всем контролам."""
    if dialog is None or dm is None:
        return
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


def _relayout(dm, sw, sh, m, gap, toolbar_h, status_h, hint_h):
    """
    Раскладка как у todo plan (split): hint + Close сверху;
    слева над списком — Удалить/Скопировать/↑/↓; справа — детали C/D.
    """
    title_h = _title_h(dm)
    total_h = int(sh)
    total_w = int(sw)
    hint_top = int(getattr(_acfg, "ALL_PARAMS_HINT_TOP_GAP", 8))
    list_tb_gap = int(getattr(_acfg, "ALL_PARAMS_LIST_TOOLBAR_GAP", 4))
    split_gap = int(getattr(_acfg, "ALL_PARAMS_SPLIT_GAP", 8))
    footer_btn_w = int(getattr(_acfg, "ALL_PARAMS_FOOTER_BTN_W", 110))
    list_lbl_h = int(getattr(_acfg, "ALL_PARAMS_LIST_LBL_H", 18))
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
        hint.Width = max(40, total_w - 2 * m - footer_btn_w - gap)
        hint.Height = hint_h
    except Exception:
        pass
    try:
        close = dm.getByName(u"CloseBtn")
        close.PositionY = y
        close.Height = toolbar_h
        close.Width = footer_btn_w
        close.PositionX = total_w - m - footer_btn_w
    except Exception:
        pass
    y += max(hint_h, toolbar_h) + gap

    content_h = _content_height(total_h, dm)
    try:
        dm.Height = content_h
    except Exception:
        pass
    status_y = content_h - m - status_h
    left_w = _left_panel_width(total_w, m, split_gap)
    right_x = m + left_w + split_gap
    right_w = total_w - right_x - m

    # Кнопки шагов — строго над списком слева.
    toolbar_y = y
    x = m
    for spec in getattr(_acfg, "ALL_PARAMS_TOOLBAR_BTNS", ()) or ():
        name = spec[0]
        bw = int(spec[2])
        try:
            btn = dm.getByName(name)
            btn.PositionX = x
            btn.PositionY = toolbar_y
            btn.Width = bw
            btn.Height = toolbar_h
        except Exception:
            pass
        x += bw + 4
    y_list = toolbar_y + toolbar_h + list_tb_gap
    try:
        ll = dm.getByName(u"StepsListLbl")
        ll.PositionX = m
        ll.PositionY = y_list
        ll.Width = left_w
        ll.Height = list_lbl_h
    except Exception:
        pass
    list_top = y_list + list_lbl_h + 2
    list_h = _list_box_height(list_top, status_y, gap=gap)
    try:
        lb = dm.getByName(u"StepsListBox")
        lb.PositionX = m
        lb.PositionY = list_top
        lb.Width = left_w
        lb.Height = list_h
    except Exception:
        pass
    # Детали справа — с верхней линии тулбара списка.
    _layout_detail_panel(dm, right_x, toolbar_y, right_w, status_y - gap)
    try:
        st = dm.getByName(u"StatusLbl")
        st.PositionX = m
        st.PositionY = status_y
        st.Width = total_w - 2 * m
        st.Height = status_h
    except Exception:
        pass


def _selected_sheet_row(state):
    entries = state.get(u"entries") or []
    si = int(state.get(u"selected") or -1)
    if 0 <= si < len(entries):
        try:
            return int(entries[si].get(u"row"))
        except Exception:
            return None
    return None


def show_all_params_wizard(doc=None, sheet=None, catalog=None, parent_dialog=None, select_row=None):
    """
    Полноэкранный обзор всех параметров листа (пайплайн).
    Возвращает (changed, selected_row): changed — были правки листа;
    selected_row — номер строки выбранного шага или None.
    """
    pw = _pw()
    if doc is None:
        xsc = _script_context()
        if xsc is not None:
            try:
                doc = xsc.getDocument()
            except Exception:
                doc = None
    if sheet is None:
        try:
            pw._show_message(u"Нет листа параметров для обзора.", doc=doc)
        except Exception:
            pass
        return (False, None)
    if catalog is None:
        catalog = pw._build_wizard_catalog()

    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError(u"Нет UNO-контекста.")
    sm = ctx.getServiceManager()
    toolkit = sm.createInstanceWithContext(u"com.sun.star.awt.Toolkit", ctx)
    sx, sy, sw, sh = _screen_work_area(toolkit)
    title_reserve = _title_h()
    dlg_w = max(640, int(sw))
    screen_h = max(480, int(sh))
    dlg_h = max(400, screen_h - title_reserve)
    m = 8
    gap = 4
    toolbar_h = int(getattr(_acfg, "ALL_PARAMS_TOOLBAR_BTN_H", 28))
    status_h = int(getattr(_acfg, "ALL_PARAMS_STATUS_H", 18))
    hint_h = int(getattr(_acfg, "ALL_PARAMS_HINT_H", 40))
    footer_btn_w = int(getattr(_acfg, "ALL_PARAMS_FOOTER_BTN_W", 110))
    split_gap = int(getattr(_acfg, "ALL_PARAMS_SPLIT_GAP", 8))
    list_lbl_h = int(getattr(_acfg, "ALL_PARAMS_LIST_LBL_H", 18))
    list_tb_gap = int(getattr(_acfg, "ALL_PARAMS_LIST_TOOLBAR_GAP", 4))

    dm = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialogModel", ctx)
    title = unicode(getattr(_acfg, "ALL_PARAMS_WIZARD_TITLE", u"Все параметры"))
    try:
        sn = unicode(sheet.Name)
        if sn:
            title = u"%s — %s" % (title, sn)
    except Exception:
        pass
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
    _add_label(
        dm,
        u"HintLbl",
        unicode(getattr(_acfg, "ALL_PARAMS_WIZARD_HINT", u"")),
        m,
        y,
        max(40, dlg_w - 2 * m - footer_btn_w - gap),
        hint_h,
        multiline=True,
    )
    _add_button(
        dm,
        u"CloseBtn",
        u"Закрыть",
        dlg_w - m - footer_btn_w,
        y,
        footer_btn_w,
        toolbar_h,
        default=True,
    )
    y += max(hint_h, toolbar_h) + gap

    left_w = _left_panel_width(dlg_w, m, split_gap)
    right_x = m + left_w + split_gap
    right_w = dlg_w - right_x - m
    content_h = _content_height(screen_h, dm)
    try:
        dm.Height = content_h
    except Exception:
        pass
    status_y = content_h - m - status_h

    toolbar_y = y
    x = m
    for spec in getattr(_acfg, "ALL_PARAMS_TOOLBAR_BTNS", ()) or ():
        _add_button(dm, spec[0], spec[1], x, toolbar_y, int(spec[2]), toolbar_h)
        x += int(spec[2]) + 4
    y_list = toolbar_y + toolbar_h + list_tb_gap
    list_top = y_list + list_lbl_h + 2
    list_h = _list_box_height(list_top, status_y, gap=gap)
    _add_label(dm, u"StepsListLbl", u"Шаги пайплайна:", m, y_list, left_w, list_lbl_h)
    _add_listbox(dm, u"StepsListBox", m, list_top, left_w, list_h)
    _create_detail_panel(dm, right_x, toolbar_y, right_w)
    _add_label(dm, u"StatusLbl", u"", m, status_y, dlg_w - 2 * m, status_h)

    try:
        apply_soft_gray_dark_blue_theme(dm, title_text=title)
    except Exception:
        pass
    _relayout(dm, dlg_w, screen_h, m, gap, toolbar_h, status_h, hint_h)

    dialog = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    if not _create_peer(dialog, toolkit, doc=doc, parent_dialog=parent_dialog):
        raise RuntimeError(u"Не удалось создать peer диалога «Все параметры».")
    # После fill workarea — пересчитать модель и жёстко синхронизировать peer-контролы
    # (иначе ListBox остаётся на начальной ширине и «съедает» правую панель).
    _relayout(dm, dlg_w, screen_h, m, gap, toolbar_h, status_h, hint_h)
    _sync_control_layout(dialog, dm)
    try:
        try_paint_titlebar(dialog, title_bg=DARK_BLUE_TITLE_BG, title_fg=DARK_BLUE_TITLE_FG)
    except Exception:
        pass

    state = {
        u"dialog": dialog,
        u"dm": dm,
        u"doc": doc,
        u"sheet": sheet,
        u"catalog": catalog,
        u"entries": [],
        u"selected": -1,
        u"changed": False,
        u"form_loading": False,
        u"loaded_snap": None,
    }

    class _Handler(unohelper.Base, XActionListener, XItemListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == u"CloseBtn":
                try:
                    if _detail_is_dirty(state):
                        _save_detail_to_sheet(state, force=True)
                        state[u"changed"] = True
                except Exception:
                    pass
                try:
                    dialog.endExecute()
                except Exception:
                    pass
                return
            if name == u"DelBtn":
                _delete_selected(state)
                state[u"changed"] = True
            elif name == u"CopyBtn":
                _copy_selected(state)
                state[u"changed"] = True
            elif name == u"UpBtn":
                _move_selected(state, -1)
                state[u"changed"] = True
            elif name == u"DownBtn":
                _move_selected(state, 1)
                state[u"changed"] = True
            elif name == u"ExtraParamsBtn":
                _open_col_c_wizard(state)
                state[u"changed"] = True
            elif name == u"CondBtn":
                _open_col_d_wizard(state)
                state[u"changed"] = True
            elif name == u"ApplyBtn":
                if _save_detail_to_sheet(state, force=True):
                    state[u"changed"] = True
                    row = None
                    entries = state.get(u"entries") or []
                    si = int(state.get(u"selected") or -1)
                    if 0 <= si < len(entries):
                        row = int(entries[si].get(u"row"))
                    _refresh_list(state, keep_row=row)

        def itemStateChanged(self, event):
            if state.get(u"form_loading"):
                return
            try:
                src_name = str(event.Source.getModel().Name)
            except Exception:
                src_name = u""
            if src_name == u"DisabledChk":
                return
            if src_name == u"NameCombo":
                # Сменили тип параметра A — обновить список B по новой спецификации.
                try:
                    pw = _pw()
                    name_raw = _get_text(dialog, u"NameCombo").strip()
                    name_canon = pw._param_canonical_sheet_name(name_raw) if name_raw else u""
                    spec = pw._find_catalog_spec(state.get(u"catalog"), name_canon)
                    primary = _get_text(dialog, u"ValueCombo")
                    _set_combo_items(
                        dialog,
                        u"ValueCombo",
                        _value_b_choices(spec or {}, primary),
                        select_text=primary,
                    )
                except Exception:
                    pass
                return
            if src_name == u"ValueCombo":
                return
            idx = _list_selected_index(dialog)
            prev = int(state.get(u"selected") or -1)
            if idx < 0 or idx == prev:
                return
            # Пишем на лист только если поля менялись — иначе Calc зависает.
            if prev >= 0 and _detail_is_dirty(state):
                if _save_detail_to_sheet(state, force=True):
                    state[u"changed"] = True
            state[u"form_loading"] = True
            try:
                state[u"selected"] = idx
                _load_detail(state, idx)
                _update_status(state)
            finally:
                state[u"form_loading"] = False

    handler = _Handler()
    for btn in (
        u"CloseBtn",
        u"DelBtn",
        u"CopyBtn",
        u"UpBtn",
        u"DownBtn",
        u"ExtraParamsBtn",
        u"CondBtn",
        u"ApplyBtn",
    ):
        try:
            dialog.getControl(btn).addActionListener(handler)
        except Exception as err:
            _log("listener %s: %s" % (btn, err))
    try:
        dialog.getControl(u"StepsListBox").addItemListener(handler)
    except Exception as err:
        _log("list listener: %s" % err)
    try:
        dialog.getControl(u"NameCombo").addItemListener(handler)
    except Exception as err:
        _log("name combo listener: %s" % err)

    if select_row is not None:
        _refresh_list(state, keep_row=int(select_row))
    else:
        _refresh_list(state, keep_index=0)

    try:
        toolkit.processEventsToIdle()
    except Exception:
        pass
    # Ещё раз после idle — LO иногда перетирает PosSize при первом показе.
    try:
        _relayout(dm, dlg_w, screen_h, m, gap, toolbar_h, status_h, hint_h)
        _sync_control_layout(dialog, dm)
    except Exception:
        pass
    try:
        dialog.execute()
    finally:
        try:
            dialog.dispose()
        except Exception:
            pass
    return (bool(state.get(u"changed")), _selected_sheet_row(state))
