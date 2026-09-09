# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
"""
Копирование строк листа параметров сбора (Параметры_Объединения* / collect_params*).

Уникальные параметры (multi_row=False) не копируются.
Вставка ниже выделения: reuse пустых строк или insertByIndex.
Для Постобработка_* / Финальная_обработка по умолчанию очищается колонка C
(параметры функции); в диалоге подтверждения можно оставить параметры.
"""
MACRO_VERSION = "3.10.715"
try:
    unicode
except NameError:
    unicode = str

import uno
import unohelper
from com.sun.star.awt import XActionListener

COPY_PARAM_ROWS_MAX = 100
_PP_CLEAR_MODES = (
    u"postprocess_range",
    u"postprocess_row",
)


def _ensure_scripts_python_on_path():
    """
    Entry-скрипты (param_wizard.py) лежат в Scripts/python/, а этот модуль —
    в Scripts/python/pythonpath/. Без родителя в sys.path import param_wizard
    даёт ModuleNotFoundError (как при запуске из меню AO).
    """
    import os
    import sys

    try:
        here = os.path.dirname(os.path.abspath(__file__))
        scripts = os.path.dirname(here)
        if scripts and os.path.isdir(scripts) and scripts not in sys.path:
            sys.path.insert(0, scripts)
    except Exception:
        pass


def _import_param_wizard():
    _ensure_scripts_python_on_path()
    import param_wizard as pw

    return pw


def _msg(text, title=u"Параметры"):
    try:
        pw = _import_param_wizard()
        pw._show_message(unicode(text), title=title)
        return
    except Exception:
        pass
    try:
        import msgbox
        mb = msgbox.MsgBox(None)
        mb.addButton("OK")
        mb.show(unicode(text), 0, unicode(title))
    except Exception:
        print(unicode(text))


def _get_selection_rows(doc, min_row=1):
    """Упорядоченный список уникальных 0-based строк выделения (без заголовка)."""
    rows = []
    if doc is None:
        return rows
    try:
        ctrl = doc.getCurrentController()
        sel = ctrl.getSelection()
        if sel is None:
            return rows
        try:
            addrs = sel.getRangeAddresses()
            if addrs is not None and len(addrs) > 0:
                for a in addrs:
                    sr = int(a.StartRow)
                    er = int(a.EndRow)
                    if sr > er:
                        sr, er = er, sr
                    r = sr
                    while r <= er:
                        if r >= int(min_row) and r not in rows:
                            rows.append(r)
                        r = r + 1
                rows.sort()
                return rows
        except Exception:
            pass
        addr = sel.getRangeAddress()
        sr = int(addr.StartRow)
        er = int(addr.EndRow)
        if sr > er:
            sr, er = er, sr
        r = sr
        while r <= er:
            if r >= int(min_row) and r not in rows:
                rows.append(r)
            r = r + 1
        rows.sort()
    except Exception:
        pass
    return rows


def _row_is_empty(sheet, row, end_col):
    c = 0
    while c <= int(end_col):
        try:
            s = sheet.getCellByPosition(c, int(row)).String
            if s is not None and unicode(s).strip() != u"":
                return False
        except Exception:
            pass
        c = c + 1
    return True


def _count_free_rows(sheet, start_row, need, end_col, scan_limit=500):
    free = 0
    r = int(start_row)
    limit = int(start_row) + int(scan_limit)
    while free < int(need) and r < limit:
        if not _row_is_empty(sheet, r, end_col):
            break
        free = free + 1
        r = r + 1
    return free


def _copy_end_col(pw, sheet, catalog, rows):
    """Ширина копирования по данным строк — без раздувания used-area до D–F."""
    end_col = 1
    for r in rows or []:
        try:
            lc = int(pw._param_sheet_row_last_col(sheet, int(r), max_cols=40) or 0)
            if lc > end_col:
                end_col = lc
        except Exception:
            pass
        try:
            mc = int(pw._param_row_max_col(sheet, int(r), catalog) or 0)
            if mc > end_col:
                end_col = mc
        except Exception:
            pass
    return max(end_col, 1)


def _restore_header_after_copy(pw, sheet, catalog, rows):
    """
    Восстановить шапку без clamp до 5 колонок.
    Явный max_col < 5 ещё и очищает ложные «Значение 3..5».
    """
    max_col = _copy_end_col(pw, sheet, catalog, rows)
    try:
        pw._restore_param_sheet_header_row(sheet, max_col)
    except Exception:
        pass


def _is_pp_or_final_spec(spec):
    if not isinstance(spec, dict):
        return False
    mode = unicode(spec.get(u"mode") or u"").strip()
    return mode in _PP_CLEAR_MODES


def _count_pp_final_rows(pw, sheet, catalog, rows):
    n = 0
    for r in rows or []:
        try:
            canon = pw._param_canonical_sheet_name(pw._row_param_name(sheet, r))
            spec = pw._find_catalog_spec(catalog, canon)
            if _is_pp_or_final_spec(spec):
                n = n + 1
        except Exception:
            pass
    return n


def _clear_pp_params_on_row(sheet, row, end_col):
    """Очистить параметры функции (C и правее); A/B (имя и ключ функции) оставить."""
    c = 2
    while c <= int(end_col):
        try:
            cell = sheet.getCellByPosition(c, int(row))
            cell.String = u""
        except Exception:
            pass
        c = c + 1


def _confirm_copy_dialog(doc, n_rows, n_pp_final):
    """
    Подтверждение копирования.
    Возвращает (ok: bool, keep_pp_params: bool).
    keep_pp_params=True — не очищать C у Постобработка_*/Финальная_обработка.
    """
    result = {u"ok": False, u"keep": False}
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(u"com.sun.star.awt.Toolkit", ctx)
        dm = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialogModel", ctx)
    except Exception:
        return (True, False)

    m = 12
    dw = 380
    msg = u"Скопировать %d строк(и) параметров ниже выделения?" % int(n_rows)
    if int(n_pp_final) > 0:
        msg = (
            msg
            + u"\n\nСреди них %d — Постобработка_* / Финальная_обработка."
            % int(n_pp_final)
        )
        msg = msg + u"\nПараметры функции (колонка C) по умолчанию будут очищены."

    y = 10
    lbl = dm.createInstance(u"com.sun.star.awt.UnoControlFixedTextModel")
    lbl.Name = u"MsgLbl"
    lbl.Label = msg
    lbl.PositionX = m
    lbl.PositionY = y
    lbl.Width = dw - m * 2
    lbl.Height = 70 if int(n_pp_final) > 0 else 36
    try:
        lbl.MultiLine = True
    except Exception:
        pass
    dm.insertByName(u"MsgLbl", lbl)
    y = y + int(lbl.Height) + 8

    show_chk = int(n_pp_final) > 0
    if show_chk:
        chk = dm.createInstance(u"com.sun.star.awt.UnoControlCheckBoxModel")
        chk.Name = u"KeepParamsChk"
        chk.Label = u"Не очищать параметры постобработки / финальной обработки"
        chk.PositionX = m
        chk.PositionY = y
        chk.Width = dw - m * 2
        chk.Height = 18
        chk.State = 0
        dm.insertByName(u"KeepParamsChk", chk)
        y = y + 26

    btn_h = 22
    btn_w = 90
    ok_btn = dm.createInstance(u"com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = u"OkBtn"
    ok_btn.Label = u"Копировать"
    ok_btn.PositionX = m
    ok_btn.PositionY = y
    ok_btn.Width = btn_w
    ok_btn.Height = btn_h
    try:
        ok_btn.DefaultButton = True
    except Exception:
        pass
    dm.insertByName(u"OkBtn", ok_btn)

    cancel_btn = dm.createInstance(u"com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = u"CancelBtn"
    cancel_btn.Label = u"Отмена"
    cancel_btn.PositionX = m + btn_w + 10
    cancel_btn.PositionY = y
    cancel_btn.Width = btn_w
    cancel_btn.Height = btn_h
    dm.insertByName(u"CancelBtn", cancel_btn)

    dm.PositionX = 140
    dm.PositionY = 120
    dm.Width = dw
    dm.Height = y + btn_h + m
    dm.Title = u"Копирование строк параметров"

    try:
        import libre_macros_ui_theme as _ui_theme

        _ui_theme.apply_soft_gray_theme(dm, title_text=dm.Title)
    except Exception:
        pass

    dlg = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    parent = None
    try:
        parent = doc.getCurrentController().getFrame().getContainerWindow()
    except Exception:
        parent = None
    try:
        dlg.createPeer(toolkit, parent)
    except Exception:
        try:
            dlg.createPeer(toolkit, None)
        except Exception:
            return (True, False)

    class _Handler(unohelper.Base, XActionListener):
        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            try:
                name = unicode(event.Source.getModel().Name)
            except Exception:
                return
            if name == u"OkBtn":
                result[u"ok"] = True
                if show_chk:
                    try:
                        result[u"keep"] = bool(dlg.getControl(u"KeepParamsChk").getState())
                    except Exception:
                        result[u"keep"] = False
                try:
                    dlg.endExecute()
                except Exception:
                    pass
            elif name == u"CancelBtn":
                result[u"ok"] = False
                try:
                    dlg.endExecute()
                except Exception:
                    pass

    handler = _Handler()
    try:
        dlg.getControl(u"OkBtn").addActionListener(handler)
        dlg.getControl(u"CancelBtn").addActionListener(handler)
    except Exception:
        pass
    try:
        dlg.execute()
    except Exception:
        return (False, False)
    try:
        dlg.dispose()
    except Exception:
        pass
    return (bool(result[u"ok"]), bool(result[u"keep"]))


def copy_param_rows_entry(doc):
    """
    Скопировать выделенные строки параметров и вставить ниже.
    Возвращает None; ошибки — через MsgBox.
    """
    pw = _import_param_wizard()

    if doc is None:
        _msg(u"Нет активного документа.")
        return None
    sheet, sheet_err = pw._active_param_sheet_for_move(doc)
    if sheet is None:
        _msg(sheet_err or u"Активный лист не является листом параметров сбора.")
        return None

    src_rows = _get_selection_rows(doc, min_row=1)
    if not src_rows:
        _msg(u"Нет строк параметров для копирования.")
        return None
    if len(src_rows) > COPY_PARAM_ROWS_MAX:
        _msg(u"Слишком много строк (макс. %d)." % COPY_PARAM_ROWS_MAX)
        return None

    catalog = pw._build_wizard_catalog()
    copy_rows = []
    unique_blocked = []
    for r in src_rows:
        raw_name = pw._row_param_name(sheet, r)
        if raw_name == u"":
            continue
        canon = pw._param_canonical_sheet_name(raw_name)
        spec = pw._find_catalog_spec(catalog, canon)
        if spec is None:
            continue
        if not spec.get("multi_row"):
            unique_blocked.append(canon)
            continue
        copy_rows.append(r)

    if unique_blocked:
        names = u", ".join(unique_blocked[:5])
        if len(unique_blocked) > 5:
            names = names + u", …"
        _msg(
            u"Параметр «%s» должен быть в одном экземпляре — копирование отменено."
            % names
        )
        return None
    if not copy_rows:
        _msg(u"Нет строк параметров для копирования.")
        return None

    n = len(copy_rows)
    n_pp = _count_pp_final_rows(pw, sheet, catalog, copy_rows)
    ok, keep_pp_params = _confirm_copy_dialog(doc, n, n_pp)
    if not ok:
        return None

    insert_at = int(max(copy_rows)) + 1
    end_col = _copy_end_col(pw, sheet, catalog, copy_rows)
    free = _count_free_rows(sheet, insert_at, n, end_col)
    clear_pp = (not keep_pp_params) and n_pp > 0

    unlock = None
    try:
        unlock = pw._param_sheet_begin_write(sheet)
        if free < n:
            try:
                sheet.getRows().insertByIndex(insert_at, n)
            except Exception as err:
                _msg(u"Не удалось вставить строки: %s" % err)
                return None
        i = 0
        while i < n:
            dest_row = insert_at + i
            pw._copy_sheet_row(sheet, copy_rows[i], dest_row, end_col)
            if clear_pp:
                try:
                    canon = pw._param_canonical_sheet_name(
                        pw._row_param_name(sheet, dest_row)
                    )
                    spec = pw._find_catalog_spec(catalog, canon)
                    if _is_pp_or_final_spec(spec):
                        _clear_pp_params_on_row(sheet, dest_row, end_col)
                except Exception:
                    pass
            i = i + 1
        _restore_header_after_copy(pw, sheet, catalog, copy_rows)
        try:
            pw._param_wizard_select_rows(doc, sheet, insert_at, insert_at + n - 1)
        except Exception:
            pass
    except Exception as err:
        _msg(u"Ошибка копирования: %s" % err)
        return None
    finally:
        if unlock is not None:
            try:
                pw._param_sheet_end_write(sheet)
            except Exception:
                pass
    return None
