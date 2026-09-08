# -*- coding: utf-8 -*-
"""
Вложенные микродиалоги параметров (ТЗ §3.2–3.4).

Этап 1: реестр INNER_DIALOG_REGISTRY, текстовый/JSON-редактор на базе кодека,
маршрутизация show_inner_param_dialog().
"""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.704"
try:
    unicode
except NameError:
    unicode = str

from libre_macros_param_codec import normalize_fn_key, param_decode, param_encode

# dialog_pattern: custom | sheet_block | sheet_list | codec_editor
# Полный реестр диалогов — в param_wizard._inner_codec_dialog_supported;
# здесь — пилоты и точка входа show_inner_param_dialog → param_wizard.
INNER_DIALOG_REGISTRY = {
    "сводная_таблица": {
        "dialog_pattern": "custom_pivot",
        "title": "Сводная таблица",
    },
    "сортировка": {
        "dialog_pattern": "codec_editor",
        "title": "Сортировка",
        "hint": "Лист | Столбец -> +/− ; несколько блоков через ;",
    },
    "раскрасить_блоки": {
        "dialog_pattern": "codec_editor",
        "title": "Раскрасить блоки",
        "hint": "Лист | ключи -> цвета/+/-!граница",
    },
    "раскрасить": {
        "dialog_pattern": "codec_editor",
        "title": "Раскрасить блоки",
        "alias_of": "раскрасить_блоки",
    },
}


def inner_param_dialog_supported(fn_key):
    try:
        from param_wizard import _inner_codec_dialog_supported

        if _inner_codec_dialog_supported(fn_key):
            return True
    except Exception:
        pass
    key = normalize_fn_key(fn_key)
    if key in INNER_DIALOG_REGISTRY:
        return True
    raw = unicode(fn_key or "").strip()
    return raw.casefold() in INNER_DIALOG_REGISTRY


def _registry_entry(fn_key):
    key = normalize_fn_key(fn_key)
    entry = INNER_DIALOG_REGISTRY.get(key)
    if entry is not None:
        return key, entry
    raw = unicode(fn_key or "").strip()
    entry = INNER_DIALOG_REGISTRY.get(raw)
    if entry is not None:
        alias = entry.get("alias_of")
        if alias:
            return normalize_fn_key(alias), INNER_DIALOG_REGISTRY.get(alias, entry)
        return normalize_fn_key(raw), entry
    return key, None


def pivot_initial_for_dialog(raw_text):
    """JSON из C → строка для show_pivot_table_param_dialog (компактный массив)."""
    blocks = param_decode("сводная_таблица", raw_text)
    if not blocks:
        return u""
    return param_encode("сводная_таблица", blocks)


def pivot_result_to_storage(pivot_json):
    """Объект/массив из диалога сводной → канонический JSON-массив в C."""
    if pivot_json is None:
        return None
    s = unicode(pivot_json).strip()
    if s == "":
        return None
    blocks = param_decode("сводная_таблица", s)
    return param_encode("сводная_таблица", blocks)


def show_inner_param_dialog( fn_key, initial_text=u"", parent_dialog=None, doc=None, *, param_sheet=None, param_row=None):
    """
    Универсальный вход для микродиалога параметров.
    Возвращает строку для колонки C или None.
    """
    try:
        from param_wizard import (
            _inner_codec_dialog_supported,
            _show_inner_codec_param_dialog,
        )

        if _inner_codec_dialog_supported(fn_key):
            return _show_inner_codec_param_dialog(
                parent_dialog, fn_key, initial_text, doc=doc
            )
    except Exception:
        pass
    key, entry = _registry_entry(fn_key)
    if entry is None:
        return None
    pattern = entry.get("dialog_pattern") or "codec_editor"
    if pattern == "custom_pivot":
        return None
    if pattern == "codec_editor":
        return _show_codec_editor_dialog(
            key,
            entry.get("title") or key,
            entry.get("hint") or u"",
            initial_text,
            parent_dialog=parent_dialog,
            doc=doc,
        )
    return None


def _show_codec_editor_dialog( fn_key, title, hint, initial_text, parent_dialog=None, doc=None):
    try:
        import uno
        from com.sun.star.awt.MessageBoxType import MESSAGEBOX
        from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
    except Exception:
        return None

    try:
        from param_wizard import (
            XSCRIPTCONTEXT,
            _dlg_control,
            _nested_subdialog_create_peer,
            _wizard_dlg_add_checkbox,
            _wizard_dlg_add_edit,
            _wizard_dlg_add_fixed,
            _wp_message_box_execute,
        )
    except Exception:
        return None

    initial_text = unicode(initial_text or "")
    as_json_default = True
    try:
        blocks = param_decode(fn_key, initial_text)
        preview = param_encode(fn_key, blocks)
    except Exception as err:
        blocks = []
        preview = initial_text
        parse_warn = unicode(err)
    else:
        parse_warn = None

    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        dm = sm.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx
        )
    except Exception:
        return None

    dw = 520
    dh = 360
    m = 10
    dm.PositionX = 0
    dm.PositionY = 0
    dm.Width = dw
    dm.Height = dh
    dm.Title = title

    y = m
    if hint:
        _wizard_dlg_add_fixed(dm, "HintLbl", hint, m, y, dw - m * 2, 28)
        y += 32

    _wizard_dlg_add_fixed(dm, "MainLbl", "Параметры (C):", m, y, 200, 14)
    y += 18
    edit_h = dh - y - 80
    _wizard_dlg_add_edit(
        dm, "ParamEdit", m, y, dw - m * 2, edit_h, multiline=True, readonly=False
    )
    y += edit_h + 6
    _wizard_dlg_add_checkbox(dm, "JsonChk", "Сохранять как JSON", m, y, 200, 14)
    y += 26
    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = "OkButton"
    ok_btn.Label = "OK"
    try:
        ok_btn.PushButtonType = 1
    except Exception:
        pass
    ok_btn.PositionX = dw - m - 170
    ok_btn.PositionY = y
    ok_btn.Width = 80
    ok_btn.Height = 22
    dm.insertByName("OkButton", ok_btn)
    cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = "CancelButton"
    cancel_btn.Label = "Отмена"
    try:
        cancel_btn.PushButtonType = 2
    except Exception:
        pass
    cancel_btn.PositionX = dw - m - 80
    cancel_btn.PositionY = y
    cancel_btn.Width = 80
    cancel_btn.Height = 22
    dm.insertByName("CancelButton", cancel_btn)

    dlg = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    try:
        ok_peer = _nested_subdialog_create_peer(dlg, toolkit, doc, parent_dialog)
    except Exception:
        ok_peer = False
    if not ok_peer and doc is not None:
        _wp_message_box_execute(
            doc,
            parent_dialog,
            MESSAGEBOX,
            MSG_BUTTONS.BUTTONS_OK,
            title,
            "Не удалось создать окно диалога (createPeer). Попытка продолжить.",
        )

    param_edit = _dlg_control(dlg, "ParamEdit")
    json_chk = _dlg_control(dlg, "JsonChk")
    if param_edit is not None:
        param_edit.setText(preview)
    if json_chk is not None:
        try:
            json_chk.setState(1 if as_json_default else 0)
        except Exception:
            pass
    if parse_warn and doc is not None:
        _wp_message_box_execute(
            doc,
            parent_dialog,
            MESSAGEBOX,
            MSG_BUTTONS.BUTTONS_OK,
            title,
            "Не удалось разобрать текущее значение:\n%s\n\nОткрыт редактор с исходной строкой."
            % parse_warn,
        )

    if dlg.execute() != 1:
        return None

    raw = u""
    if param_edit is not None:
        try:
            raw = param_edit.getText().strip()
        except Exception:
            raw = u""
    use_json = True
    if json_chk is not None:
        try:
            use_json = bool(json_chk.getState())
        except Exception:
            pass
    try:
        blocks = param_decode(fn_key, raw)
        return param_encode(fn_key, blocks)
    except Exception as err:
        _wp_message_box_execute(
            doc,
            parent_dialog,
            MESSAGEBOX,
            MSG_BUTTONS.BUTTONS_OK,
            title,
            "Ошибка параметров:\n%s" % unicode(err),
        )
        return None
