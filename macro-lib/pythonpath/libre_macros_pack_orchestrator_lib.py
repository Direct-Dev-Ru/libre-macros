# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.728"
"""
Диалог и persist настроек оркестратора collect_pack.

Цикл запуска — в collect_workbooks.collect_pack / helpers.
"""

import json
import os
import sys
import traceback

import uno
import unohelper
from com.sun.star.awt import XActionListener, XItemListener, XMouseListener

try:
    unicode
except NameError:
    unicode = str

import libre_macros_pack_orchestrator_cfg as _cfg


def _u(value):
    if value is None:
        return u""
    try:
        if isinstance(value, unicode):
            return value
    except NameError:
        pass
    try:
        return unicode(value)
    except Exception:
        return str(value)


def _pack_config_base():
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.environ.get(
            "XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config")
        )
    return os.path.join(base, "libre-macros")


def _pack_presets_dir():
    path = os.path.join(_pack_config_base(), _cfg.PACK_ORCH_PRESETS_SUBDIR)
    try:
        os.makedirs(path, exist_ok=True)
    except TypeError:
        if not os.path.isdir(path):
            try:
                os.makedirs(path)
            except OSError:
                pass
    except OSError:
        pass
    return path


def _pack_json_path():
    return os.path.join(_pack_presets_dir(), _cfg.PACK_ORCH_JSON_FILE)


def lm_pack_load_store():
    """Загрузить весь JSON store {v, books}."""
    path = _pack_json_path()
    if not os.path.isfile(path):
        return {u"v": _cfg.PACK_ORCH_JSON_VER, u"books": {}}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except TypeError:
        with open(path, "r") as fh:
            data = json.load(fh)
    except Exception:
        return {u"v": _cfg.PACK_ORCH_JSON_VER, u"books": {}}
    if not isinstance(data, dict):
        return {u"v": _cfg.PACK_ORCH_JSON_VER, u"books": {}}
    books = data.get(u"books")
    if not isinstance(books, dict):
        books = {}
    return {u"v": int(data.get(u"v") or _cfg.PACK_ORCH_JSON_VER), u"books": books}


def lm_pack_save_store(store):
    path = _pack_json_path()
    payload = {
        u"v": _cfg.PACK_ORCH_JSON_VER,
        u"books": (store or {}).get(u"books") or {},
    }
    text = json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True)
    try:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(text)
            fh.write("\n")
    except TypeError:
        with open(path, "w") as fh:
            fh.write(text.encode("utf-8") if isinstance(text, unicode) else text)
            fh.write("\n")
    try:
        from libre_macros_global_settings_lib import notify_config_changed

        notify_config_changed()
    except Exception:
        pass


def lm_pack_load_book_entry(book_path):
    key = _u(book_path or u"").strip()
    if key == u"":
        return None
    store = lm_pack_load_store()
    entry = (store.get(u"books") or {}).get(key)
    if not isinstance(entry, dict):
        return None
    return entry


def lm_pack_save_book_entry(book_path, entry):
    key = _u(book_path or u"").strip()
    if key == u"":
        raise ValueError(_cfg.PACK_ORCH_ERR_UNSAVED)
    store = lm_pack_load_store()
    books = store.get(u"books") or {}
    books[key] = entry
    store[u"books"] = books
    lm_pack_save_store(store)


def lm_pack_default_rows(sheet_names):
    rows = []
    for name in sheet_names or []:
        nm = _u(name).strip()
        if nm != u"":
            rows.append({u"name": nm, u"enabled": True})
    return rows


def lm_pack_merge_saved_rows(sheet_names, saved_entry):
    """Сопоставить сохранённый порядок/enable с актуальными именами листов.

    Возвращает (rows, reset_db, auto_continue).
    auto_continue по умолчанию True (мульти-план без лишних «Запуск»).
    """
    names = [_u(n).strip() for n in (sheet_names or []) if _u(n).strip() != u""]
    if not names:
        return [], False, True
    reset_db = False
    auto_continue = True
    saved_sheets = []
    if isinstance(saved_entry, dict):
        reset_db = bool(saved_entry.get(u"reset_db", False))
        if u"auto_continue" in saved_entry:
            auto_continue = bool(saved_entry.get(u"auto_continue"))
        raw = saved_entry.get(u"sheets") or []
        if isinstance(raw, list):
            for item in raw:
                if not isinstance(item, dict):
                    continue
                nm = _u(item.get(u"name")).strip()
                if nm == u"":
                    continue
                saved_sheets.append(
                    {u"name": nm, u"enabled": bool(item.get(u"enabled", True))}
                )
    known = set(names)
    rows = []
    seen = set()
    for item in saved_sheets:
        nm = item[u"name"]
        if nm in known and nm not in seen:
            rows.append({u"name": nm, u"enabled": bool(item.get(u"enabled", True))})
            seen.add(nm)
    for nm in names:
        if nm not in seen:
            rows.append({u"name": nm, u"enabled": True})
            seen.add(nm)
    return rows, reset_db, auto_continue


def lm_pack_plan_from_rows(rows):
    plan = []
    for item in rows or []:
        if not item.get(u"enabled"):
            continue
        nm = _u(item.get(u"name")).strip()
        if nm != u"":
            plan.append(nm)
    return plan


def _pack_comment_preview(text, max_len=None):
    """Первые max_len символов комментария (одна строка)."""
    if max_len is None:
        max_len = int(getattr(_cfg, "PACK_ORCH_COMMENT_PREVIEW_LEN", 50) or 50)
    s = _u(text or u"").strip()
    if s == u"":
        return u""
    s = s.replace(u"\r\n", u" ").replace(u"\n", u" ").replace(u"\r", u" ")
    while u"  " in s:
        s = s.replace(u"  ", u" ")
    try:
        n = int(max_len)
    except (TypeError, ValueError):
        n = 50
    if n < 1:
        n = 50
    if len(s) <= n:
        return s
    if n == 1:
        return u"…"
    return s[: n - 1] + u"…"


def _pack_sheet_display_label(name, sheet_comments=None):
    """Имя листа + (начало комментария) для подписи в списке."""
    nm = _u(name).strip()
    if nm == u"":
        return u""
    preview = u""
    if sheet_comments:
        raw = sheet_comments.get(nm)
        if raw is None:
            raw = sheet_comments.get(name)
        preview = _pack_comment_preview(raw)
    if preview != u"":
        return u"%s (%s)" % (nm, preview)
    return nm


def _pack_row_label(item, sheet_comments=None):
    """Метка для отладки/совместимости (в UI — CheckBox + TextColor)."""
    mark = _cfg.PACK_ORCH_MARK_ON if item.get(u"enabled") else _cfg.PACK_ORCH_MARK_OFF
    return mark + _u(item.get(u"name"))


def _pack_row_text_color(enabled):
    return int(_cfg.PACK_ORCH_COLOR_ON if enabled else _cfg.PACK_ORCH_COLOR_OFF)


def _pack_dlg_add_fixed(dm, name, x, y, w, h, label=u"", multiline=False):
    m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    m.Name = str(name)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    m.Label = _u(label)
    if multiline:
        try:
            m.MultiLine = True
        except Exception:
            pass
    dm.insertByName(str(name), m)
    return m


def _pack_dlg_add_button(dm, name, label, x, y, w, h, font_height=None):
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.Name = str(name)
    m.Label = _u(label)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    if font_height is not None:
        try:
            fh = int(font_height)
        except Exception:
            fh = 0
        if fh > 0:
            try:
                fd = m.FontDescriptor
                fd.Height = fh
                m.FontDescriptor = fd
            except Exception:
                try:
                    m.FontHeight = fh
                except Exception:
                    pass
    dm.insertByName(str(name), m)
    return m


def _pack_dlg_add_side_button(dm, name, label, x, y, w, h):
    """Кнопка правого вертикального ряда (+1 pt к подписи)."""
    try:
        fh = int(getattr(_cfg, "PACK_ORCH_SIDE_BTN_FONT_HEIGHT", 11) or 11)
    except Exception:
        fh = 11
    return _pack_dlg_add_button(dm, name, label, x, y, w, h, font_height=fh)

def _pack_dlg_add_checkbox(dm, name, label, x, y, w, h, checked=True):
    m = dm.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
    m.Name = str(name)
    m.Label = _u(label)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    try:
        m.State = 1 if checked else 0
    except Exception:
        pass
    _pack_apply_chk_font(m, selected=False)
    dm.insertByName(str(name), m)
    return m


def _pack_apply_chk_font(model, selected=False):
    """Крупнее шрифт подписи листа; при выделении строки — жирный."""
    if model is None:
        return
    try:
        h = int(getattr(_cfg, "PACK_ORCH_CHK_FONT_HEIGHT", 11) or 11)
    except Exception:
        h = 11
    if h < 9:
        h = 9
    weight = 150 if selected else 100
    try:
        fd = model.FontDescriptor
        fd.Height = h
        try:
            fd.Weight = weight
        except Exception:
            pass
        model.FontDescriptor = fd
    except Exception:
        try:
            model.FontHeight = h
        except Exception:
            pass
        try:
            model.FontWeight = weight
        except Exception:
            pass


def _pack_list_h(n_rows):
    """Высота блока задач: по числу строк (без обрезки — иначе чекбоксы налезут на кнопки)."""
    row_h = int(_cfg.PACK_ORCH_ROW_H)
    n = max(1, int(n_rows or 1))
    return int(max(n * row_h, 64))


def _pack_chk_name(idx):
    return "SheetChk%d" % int(idx)


def _pack_parse_chk_index(name):
    s = str(name or "")
    if not s.startswith("SheetChk"):
        return -1
    try:
        return int(s[8:])
    except Exception:
        return -1


def _pack_selected_indices(state):
    sel = state.get(u"sel") or []
    out = []
    for x in sel:
        try:
            i = int(x)
        except Exception:
            continue
        if i >= 0:
            out.append(i)
    return sorted(set(out))


def _pack_set_selection(state, indices, rows_len):
    clean = []
    for i in indices or []:
        try:
            ii = int(i)
        except Exception:
            continue
        if 0 <= ii < int(rows_len):
            clean.append(ii)
    state[u"sel"] = sorted(set(clean)) if clean else ([0] if rows_len > 0 else [])


def _pack_apply_row_visual(dialog, state, idx, item):
    name = _pack_chk_name(idx)
    try:
        model = dialog.getModel().getByName(name)
    except Exception:
        return
    enabled = bool(item.get(u"enabled"))
    try:
        model.Label = _pack_sheet_display_label(
            item.get(u"name"), state.get(u"sheet_comments")
        )
    except Exception:
        pass
    try:
        model.State = 1 if enabled else 0
    except Exception:
        pass
    try:
        model.TextColor = _pack_row_text_color(enabled)
    except Exception:
        pass
    sel = set(_pack_selected_indices(state))
    selected = idx in sel
    try:
        model.BackgroundColor = (
            int(_cfg.PACK_ORCH_COLOR_SEL_BG)
            if selected
            else int(_cfg.PACK_ORCH_COLOR_ROW_BG)
        )
    except Exception:
        pass
    _pack_apply_chk_font(model, selected=selected)


def _pack_refresh_sheet_rows(dialog, state, select_indices=None):
    rows = state.get(u"rows") or []
    if select_indices is not None:
        _pack_set_selection(state, select_indices, len(rows))
    elif not _pack_selected_indices(state) and rows:
        state[u"sel"] = [0]
    state[u"_suppress_item"] = True
    try:
        for i, item in enumerate(rows):
            _pack_apply_row_visual(dialog, state, i, item)
    finally:
        state[u"_suppress_item"] = False


def _pack_read_checkbox_enabled(dialog, idx):
    try:
        return int(dialog.getControl(_pack_chk_name(idx)).getState()) != 0
    except Exception:
        try:
            return int(dialog.getModel().getByName(_pack_chk_name(idx)).State) != 0
        except Exception:
            return True


def _pack_sync_enabled_from_ui(dialog, rows):
    for i in range(len(rows or [])):
        rows[i][u"enabled"] = _pack_read_checkbox_enabled(dialog, i)


def _pack_toggle_indices(rows, indices):
    """Инвертировать enabled у указанных индексов. Возвращает True, если что-то меняли."""
    changed = False
    for idx in indices or []:
        try:
            i = int(idx)
        except Exception:
            continue
        if i < 0 or i >= len(rows):
            continue
        rows[i][u"enabled"] = not bool(rows[i].get(u"enabled"))
        changed = True
    return changed


def _pack_checkbox_state(dialog, name):
    try:
        return int(dialog.getControl(name).getState()) != 0
    except Exception:
        try:
            return int(dialog.getModel().getByName(name).State) != 0
        except Exception:
            return True


def _pack_center_dialog(dm, toolkit, parent_win):
    """
    Как диалог сбора / param_wizard: PositionX/Y = 80/60 относительно parent
    при createPeer(parent). Не центрировать через getPosSize/getWorkArea —
    на двух мониторах окно уезжает в «щель» между экранами.
    """
    try:
        import macro_ui

        if hasattr(macro_ui, "dlg_center_model"):
            macro_ui.dlg_center_model(dm, toolkit, parent_win)
            return
    except Exception:
        pass
    try:
        dm.PositionX = 80
        dm.PositionY = 60
    except Exception:
        pass


def _pack_show_help_dialog(parent_dialog=None, toolkit=None, sm=None, ctx=None, parent_window=None, doc=None):
    """Модальная справка оркестратора (кнопка «Справка»)."""
    title = getattr(_cfg, u"PACK_ORCH_HELP_TITLE", None) or u"Справка — оркестратор сбора"
    text = getattr(_cfg, u"PACK_ORCH_HELP_TEXT", None) or u"Справка недоступна."
    if sm is None or toolkit is None:
        return
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dw = 520
    dh = 360
    m = 10
    dm.PositionX = 100
    dm.PositionY = 80
    dm.Width = dw
    dm.Height = dh
    dm.Title = title
    try:
        dm.Sizeable = True
    except Exception:
        pass
    btn_h = 22
    btn_w = 100
    y_btn = dh - m - btn_h
    edit_h = max(80, y_btn - m - 8)
    ed = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
    ed.Name = "HelpTextEd"
    ed.PositionX = m
    ed.PositionY = m
    ed.Width = dw - m * 2
    ed.Height = edit_h
    try:
        ed.MultiLine = True
        ed.ReadOnly = True
        ed.VScroll = True
    except Exception:
        pass
    try:
        from libre_macros_ui_theme import HELP_TEXT_FONT_HEIGHT, apply_edit_font

        apply_edit_font(ed, HELP_TEXT_FONT_HEIGHT)
    except Exception:
        pass
    dm.insertByName("HelpTextEd", ed)
    close_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    close_btn.Name = "HelpCloseBtn"
    close_btn.Label = u"Закрыть"
    close_btn.PositionX = dw - m - btn_w
    close_btn.PositionY = y_btn
    close_btn.Width = btn_w
    close_btn.Height = btn_h
    try:
        close_btn.PushButtonType = 1
        close_btn.DefaultButton = True
    except Exception:
        pass
    dm.insertByName("HelpCloseBtn", close_btn)
    dlg = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    try:
        peer = parent_window
        if peer is None and parent_dialog is not None:
            try:
                peer = parent_dialog.getPeer()
            except Exception:
                peer = None
        if peer is None:
            peer = toolkit.getDesktopWindow()
        dlg.createPeer(toolkit, peer)
    except Exception:
        try:
            dlg.createPeer(toolkit, toolkit.getDesktopWindow())
        except Exception:
            return
    try:
        ctl = dlg.getControl("HelpTextEd")
        ctl.setText(_u(text))
    except Exception:
        pass
    try:
        dlg.execute()
    finally:
        try:
            dlg.dispose()
        except Exception:
            pass


def lm_pack_show_dialog( doc, sheet_names, book_path=u"", toolkit=None, parent_window=None, show_message=None, ctx=None, sheet_comments=None):
    """
    Диалог оркестратора.

    Возвращает:
      None / {'ok': False} — отмена
      {'ok': True, 'plan': [...], 'reset_db': bool, 'rows': [...]} — запуск
    """
    names = [_u(n).strip() for n in (sheet_names or []) if _u(n).strip() != u""]
    if not names:
        if show_message is not None:
            show_message(_cfg.PACK_ORCH_ERR_NO_SHEETS, doc)
        return {u"ok": False}

    saved = lm_pack_load_book_entry(book_path) if _u(book_path).strip() else None
    rows, _reset_db_unused, auto_continue = lm_pack_merge_saved_rows(names, saved)
    single = len(names) == 1
    if single:
        auto_continue = False

    if ctx is None:
        try:
            ctx = uno.getComponentContext()
        except Exception:
            ctx = None
    if ctx is None:
        if show_message is not None:
            show_message(u"Нет ComponentContext для диалога оркестратора.", doc)
        return {u"ok": False}
    sm = ctx.getServiceManager()
    if toolkit is None:
        try:
            toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        except Exception:
            toolkit = None
    if toolkit is None:
        if show_message is not None:
            show_message(u"Нет Toolkit для диалога оркестратора.", doc)
        return {u"ok": False}

    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    m = _cfg.PACK_DLG_MARGIN
    g = _cfg.PACK_DLG_GAP
    btn_h = _cfg.PACK_DLG_BTN_H
    btn_w = _cfg.PACK_DLG_BTN_W
    row_h = int(_cfg.PACK_ORCH_ROW_H)
    dm.Title = _cfg.PACK_ORCH_TITLE
    dm.Width = _cfg.PACK_DLG_W
    content_w = dm.Width - 2 * m
    y = m

    if single:
        preview = _pack_comment_preview(
            (sheet_comments or {}).get(names[0]) if sheet_comments else u""
        )
        suffix = (u" (%s)" % preview) if preview != u"" else u""
        hint = _cfg.PACK_ORCH_HINT_SINGLE % (names[0], suffix)
        _pack_dlg_add_fixed(
            dm, "HintLbl", m, y, content_w, _cfg.PACK_DLG_HINT_H, hint, multiline=True
        )
        y = y + _cfg.PACK_DLG_HINT_H + g
    else:
        _pack_dlg_add_fixed(
            dm,
            "HintLbl",
            m,
            y,
            content_w,
            _cfg.PACK_DLG_HINT_H,
            _cfg.PACK_ORCH_HINT_MULTI,
            multiline=True,
        )
        y = y + _cfg.PACK_DLG_HINT_H + g
        list_h = _pack_list_h(len(rows))
        list_w = content_w - btn_w - g
        for i, item in enumerate(rows):
            chk = _pack_dlg_add_checkbox(
                dm,
                _pack_chk_name(i),
                _pack_sheet_display_label(item.get(u"name"), sheet_comments),
                m,
                y + i * row_h,
                list_w,
                row_h,
                checked=bool(item.get(u"enabled")),
            )
            try:
                chk.TextColor = _pack_row_text_color(item.get(u"enabled"))
            except Exception:
                pass
            try:
                chk.BackgroundColor = int(_cfg.PACK_ORCH_COLOR_ROW_BG)
            except Exception:
                pass
        bx = m + list_w + g
        by = y
        _pack_dlg_add_side_button(
            dm, "EnableAllBtn", _cfg.PACK_ORCH_BTN_ENABLE_ALL, bx, by, btn_w, btn_h
        )
        by = by + btn_h + g
        _pack_dlg_add_side_button(
            dm, "DisableAllBtn", _cfg.PACK_ORCH_BTN_DISABLE_ALL, bx, by, btn_w, btn_h
        )
        by = by + btn_h + g
        _pack_dlg_add_side_button(
            dm, "ToggleBtn", _cfg.PACK_ORCH_BTN_TOGGLE, bx, by, btn_w, btn_h
        )
        by = by + btn_h + g
        _pack_dlg_add_side_button(
            dm, "UpBtn", _cfg.PACK_ORCH_BTN_UP, bx, by, btn_w, btn_h
        )
        by = by + btn_h + g
        _pack_dlg_add_side_button(
            dm, "DownBtn", _cfg.PACK_ORCH_BTN_DOWN, bx, by, btn_w, btn_h
        )
        side_h = by + btn_h - y
        y = y + max(list_h, side_h) + g

    if not single:
        _pack_dlg_add_checkbox(
            dm,
            "AutoContChk",
            _cfg.PACK_ORCH_AUTO_CONTINUE_LABEL,
            m,
            y,
            content_w,
            _cfg.PACK_DLG_CHECK_H,
            checked=auto_continue,
        )
        y = y + _cfg.PACK_DLG_CHECK_H + g

    save_w = int(getattr(_cfg, "PACK_DLG_SAVE_BTN_W", 122))
    help_w = int(getattr(_cfg, "PACK_DLG_HELP_BTN_W", 68))
    help_label = getattr(_cfg, "PACK_ORCH_BTN_HELP", None) or u"Справка"
    gap_sm = int(getattr(_cfg, "PACK_DLG_FOOTER_GAP_SMALL", g) or g)
    # [Запустить] [gap] [Отмена] …… [Сохранить настройки] [gap] [Справка]
    run_x = m
    cancel_x = run_x + btn_w + gap_sm
    help_x = dm.Width - m - help_w
    save_x = help_x - gap_sm - save_w
    _pack_dlg_add_button(dm, "RunBtn", _cfg.PACK_ORCH_BTN_RUN, run_x, y, btn_w, btn_h)
    _pack_dlg_add_button(
        dm, "CancelBtn", _cfg.PACK_ORCH_BTN_CANCEL, cancel_x, y, btn_w, btn_h
    )
    _pack_dlg_add_button(dm, "SaveBtn", _cfg.PACK_ORCH_BTN_SAVE, save_x, y, save_w, btn_h)
    _pack_dlg_add_button(dm, "HelpBtn", help_label, help_x, y, help_w, btn_h)
    dm.Height = y + btn_h + m

    try:
        from libre_macros_ui_theme import apply_soft_gray_theme

        apply_soft_gray_theme(dm, title_text=dm.Title)
    except Exception:
        pass

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    _pack_center_dialog(dm, toolkit, parent_window)
    try:
        peer = toolkit.getDesktopWindow() if parent_window is None else parent_window
        dialog.createPeer(toolkit, peer)
    except Exception:
        try:
            dialog.createPeer(toolkit, toolkit.getDesktopWindow())
        except Exception as err:
            if show_message is not None:
                show_message(u"Не удалось показать диалог оркестратора:\n%s" % err, doc)
            return {u"ok": False}
    try:
        from libre_macros_ui_theme import try_paint_titlebar

        try_paint_titlebar(dialog)
    except Exception:
        pass

    state = {
        u"rows": list(rows),
        u"result": {u"ok": False},
        u"book_path": _u(book_path).strip(),
        u"sel": [0] if rows else [],
        u"_suppress_item": False,
        u"sheet_comments": dict(sheet_comments or {}),
    }
    if not single:
        _pack_refresh_sheet_rows(dialog, state, [0])

    def _current_entry():
        if not single:
            _pack_sync_enabled_from_ui(dialog, state[u"rows"])
        entry = {
            u"sheets": [
                {u"name": _u(it.get(u"name")), u"enabled": bool(it.get(u"enabled"))}
                for it in state[u"rows"]
            ],
            # Галочка сброса диапазонов убрана из UI; макрос collect_reset_db_ranges — отдельно.
            u"reset_db": False,
        }
        if not single:
            entry[u"auto_continue"] = _pack_checkbox_state(dialog, "AutoContChk")
        return entry

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = str(event.Source.getModel().Name)
            except Exception:
                return
            if name == "CancelBtn":
                state[u"result"] = {u"ok": False}
                try:
                    dialog.endExecute()
                except Exception:
                    pass
                return
            if name == "HelpBtn":
                try:
                    _pack_show_help_dialog(
                        dialog,
                        toolkit=toolkit,
                        sm=sm,
                        ctx=ctx,
                        parent_window=parent_window,
                        doc=doc,
                    )
                except Exception:
                    if show_message is not None:
                        try:
                            show_message(
                                getattr(_cfg, u"PACK_ORCH_HELP_TEXT", u"") or u"Справка недоступна.",
                                doc,
                            )
                        except Exception:
                            pass
                return
            if name == "RunBtn":
                if not single:
                    _pack_sync_enabled_from_ui(dialog, state[u"rows"])
                plan = lm_pack_plan_from_rows(state[u"rows"])
                if not plan:
                    if show_message is not None:
                        show_message(_cfg.PACK_ORCH_ERR_EMPTY_PLAN, doc)
                    return
                auto_cont = False
                if not single:
                    auto_cont = _pack_checkbox_state(dialog, "AutoContChk")
                state[u"result"] = {
                    u"ok": True,
                    u"plan": plan,
                    u"reset_db": False,
                    u"auto_continue": bool(auto_cont),
                    u"rows": list(state[u"rows"]),
                }
                try:
                    dialog.endExecute()
                except Exception:
                    pass
                return
            if name == "SaveBtn":
                path = state[u"book_path"]
                if path == u"":
                    if show_message is not None:
                        show_message(_cfg.PACK_ORCH_ERR_UNSAVED, doc)
                    return
                try:
                    lm_pack_save_book_entry(path, _current_entry())
                    if show_message is not None:
                        show_message(_cfg.PACK_ORCH_SAVE_OK, doc)
                except Exception as err:
                    if show_message is not None:
                        show_message(_cfg.PACK_ORCH_ERR_SAVE_FAILED % err, doc)
                return
            if single:
                return
            if name == "EnableAllBtn":
                for it in state[u"rows"]:
                    it[u"enabled"] = True
                sel = _pack_selected_indices(state) or [0]
                _pack_refresh_sheet_rows(dialog, state, select_indices=sel)
                return
            if name == "DisableAllBtn":
                for it in state[u"rows"]:
                    it[u"enabled"] = False
                sel = _pack_selected_indices(state) or [0]
                _pack_refresh_sheet_rows(dialog, state, select_indices=sel)
                return
            indices = _pack_selected_indices(state)
            if name == "ToggleBtn":
                if not indices:
                    return
                if not _pack_toggle_indices(state[u"rows"], indices):
                    return
                _pack_refresh_sheet_rows(dialog, state, select_indices=indices)
                return
            idx = indices[0] if indices else -1
            if name == "UpBtn":
                if idx <= 0 or idx >= len(state[u"rows"]):
                    return
                rows_l = state[u"rows"]
                rows_l[idx - 1], rows_l[idx] = rows_l[idx], rows_l[idx - 1]
                _pack_refresh_sheet_rows(dialog, state, select_indices=[idx - 1])
                return
            if name == "DownBtn":
                if idx < 0 or idx >= len(state[u"rows"]) - 1:
                    return
                rows_l = state[u"rows"]
                rows_l[idx + 1], rows_l[idx] = rows_l[idx], rows_l[idx + 1]
                _pack_refresh_sheet_rows(dialog, state, select_indices=[idx + 1])
                return

    handler = _Handler()
    btn_names = ["RunBtn", "CancelBtn", "SaveBtn", "HelpBtn"]
    if not single:
        btn_names.extend(
            ["ToggleBtn", "UpBtn", "DownBtn", "EnableAllBtn", "DisableAllBtn"]
        )
    for btn_name in btn_names:
        try:
            dialog.getControl(btn_name).addActionListener(handler)
        except Exception:
            pass

    if not single:
        class _ChkItem(unohelper.Base, XItemListener):
            def disposing(self, event):
                pass

            def itemStateChanged(self, event):
                if state.get(u"_suppress_item"):
                    return
                try:
                    name = str(event.Source.getModel().Name)
                except Exception:
                    return
                idx = _pack_parse_chk_index(name)
                if idx < 0 or idx >= len(state[u"rows"]):
                    return
                enabled = _pack_read_checkbox_enabled(dialog, idx)
                state[u"rows"][idx][u"enabled"] = enabled
                _pack_apply_row_visual(dialog, state, idx, state[u"rows"][idx])

        class _ChkMouse(unohelper.Base, XMouseListener):
            def disposing(self, event):
                pass

            def mousePressed(self, event):
                try:
                    name = str(event.Source.getModel().Name)
                except Exception:
                    return
                idx = _pack_parse_chk_index(name)
                if idx < 0 or idx >= len(state[u"rows"]):
                    return
                try:
                    mods = int(getattr(event, "Modifiers", 0) or 0)
                except Exception:
                    mods = 0
                # KEY_MOD1 / Ctrl ≈ 2 в awt.KeyModifier
                ctrl = bool(mods & 2) or bool(mods & 8)
                if ctrl:
                    sel = set(_pack_selected_indices(state))
                    if idx in sel:
                        sel.discard(idx)
                    else:
                        sel.add(idx)
                    if not sel:
                        sel.add(idx)
                    _pack_set_selection(state, list(sel), len(state[u"rows"]))
                else:
                    _pack_set_selection(state, [idx], len(state[u"rows"]))
                _pack_refresh_sheet_rows(dialog, state)

            def mouseReleased(self, event):
                pass

            def mouseEntered(self, event):
                pass

            def mouseExited(self, event):
                pass

        chk_item = _ChkItem()
        chk_mouse = _ChkMouse()
        for i in range(len(state[u"rows"])):
            try:
                ctl = dialog.getControl(_pack_chk_name(i))
                ctl.addItemListener(chk_item)
                ctl.addMouseListener(chk_mouse)
            except Exception:
                pass
    try:
        dialog.setVisible(True)
    except Exception:
        pass
    try:
        dialog.execute()
    except Exception as err:
        try:
            dialog.dispose()
        except Exception:
            pass
        if show_message is not None:
            show_message(u"Диалог оркестратора: %s" % err, doc)
        return {u"ok": False}
    try:
        dialog.dispose()
    except Exception:
        pass
    return state[u"result"]
