# -*- coding: utf-8 -*-
"""
Разделение листа на листы/книги по ключевым полям (разделить_листы).
"""

from __future__ import print_function
MACRO_VERSION = "3.10.704"
import os

try:
    import uno
    from com.sun.star.beans import PropertyValue
except Exception:
    uno = None
    PropertyValue = None

import libre_macros_lib as lm

_THIS_BOOK_MARKERS = ("эта_книга", "this_book", "current_book", "current")
_VALID_RUN_PHASES = ("after_collect", "after_postprocess", "after_final")


def _u(text):
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _is_this_book_marker(value):
    try:
        import libre_macros_collect_cfg as _ccfg

        return bool(_ccfg.is_current_book_marker_key(value))
    except Exception:
        pass
    v = _u(value or u"").strip().casefold()
    key = []
    for ch in v:
        if ch in u" \t_-.":
            continue
        if ch == u"ё":
            ch = u"е"
        key.append(ch)
    k = u"".join(key)
    if k in (
        u"этакнига",
        u"thisworkbook",
        u"thisbook",
        u"currentworkbook",
        u"currentbook",
        u"current",
    ):
        return True
    return v in tuple(x.casefold() for x in _THIS_BOOK_MARKERS)


def _make_prop(name, value):
    if PropertyValue is None:
        return None
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _script_context():
    """XSCRIPTCONTEXT из __main__ (в pythonpath-модуле его нет)."""
    try:
        import __main__

        return getattr(__main__, "XSCRIPTCONTEXT", None)
    except Exception:
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


def _desktop(ctx=None):
    """Desktop UNO: XSCRIPTCONTEXT → createInstance(Desktop) → None."""
    if ctx is None:
        xsc = _script_context()
        if xsc is not None:
            try:
                desk = xsc.getDesktop()
                if desk is not None:
                    return desk
            except Exception:
                pass
        ctx = _uno_context()
    if ctx is None:
        return None
    try:
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    except Exception:
        return None


def _split_log(run_context, doc, sheet_name, phase, label, status, note):
    fn = (run_context or {}).get("log_fn")
    if fn is not None:
        try:
            fn(doc, sheet_name or "", phase, label, status, note)
            return
        except Exception:
            pass
    lm._lm_log_postprocess(doc, sheet_name or "", phase, label, status, note)


def _split_notify_sheet_created(run_context, doc, sheet_name, header_row_0, source_sheet):
    reg = (run_context or {}).get("register_pending")
    if reg is not None:
        try:
            reg(sheet_name, int(header_row_0), _u(source_sheet or u""))
        except Exception:
            pass
    append_fn = (run_context or {}).get("append_sheets_info")
    if append_fn is not None:
        try:
            append_fn(doc)
        except Exception:
            pass


def _split_resolve_header_row_0(sheet, header_row_param):
    if header_row_param not in (None, u""):
        try:
            return max(int(header_row_param) - 1, 0)
        except (TypeError, ValueError):
            pass
    try:
        return int(lm._lm_pp_copy_sheet_guess_header_row(sheet))
    except Exception:
        return 0


def _split_header_titles(sheet, header_row_0, sc, ec):
    titles = {}
    c = int(sc)
    while c <= int(ec):
        try:
            titles[c] = lm.cell_text(sheet.getCellByPosition(c, int(header_row_0)))
        except Exception:
            titles[c] = u""
        c = c + 1
    by_name = {}
    for col, title in titles.items():
        t = _u(title or u"").strip()
        if t != u"":
            by_name[t.casefold()] = int(col)
    return titles, by_name


def _split_resolve_field_columns(field_names, by_name):
    cols = []
    missing = []
    for name in field_names or []:
        key = _u(name or u"").strip().casefold()
        if key == u"":
            continue
        col = by_name.get(key)
        if col is None:
            missing.append(_u(name))
        else:
            cols.append(int(col))
    return cols, missing


def _split_row_field_values(sheet, row, field_names, by_name):
    out = {}
    for name in field_names or []:
        key = _u(name or u"").strip()
        if key == u"":
            continue
        col = by_name.get(key.casefold())
        if col is None:
            out[key] = u""
            continue
        try:
            out[key] = lm.cell_text(sheet.getCellByPosition(int(col), int(row)))
        except Exception:
            out[key] = u""
    return out


def _split_row_key_values(sheet, row, field_names, by_name):
    vals = []
    for name in field_names or []:
        key = _u(name or u"").strip()
        if key == u"":
            continue
        col = by_name.get(key.casefold())
        if col is None:
            vals.append(None)
            continue
        try:
            vals.append(lm.cell_text(sheet.getCellByPosition(int(col), int(row))))
        except Exception:
            vals.append(u"")
    return vals


def _split_abs_path(path, base_dir):
    raw = _u(path or u"").strip()
    if raw == u"":
        return u""
    if os.path.isabs(raw):
        return os.path.abspath(raw)
    base = _u(base_dir or u"").strip()
    if base == u"":
        return os.path.abspath(raw)
    return os.path.abspath(os.path.join(base, raw))


def _split_path_is_absolute(text):
    """Абсолютный путь: Unix `/…` или Windows `C:/…`."""
    s = _u(text or u"").replace(u"\\", u"/").strip()
    if s.startswith(u"/"):
        return True
    if len(s) >= 2 and s[1] == u":":
        return True
    return False


def _split_resolve_book_path(template, field_values, base_dir, create_folders):
    if _is_this_book_marker(template):
        return u""
    text = lm._lm_replace_fields(template, field_values)
    text = _u(text or u"").replace(u"\\", u"/").strip()
    if text == u"":
        return u""
    is_abs = _split_path_is_absolute(text)
    parts = []
    for part in text.split(u"/"):
        if part in (u"", u"."):
            continue
        if part == u"..":
            if parts:
                parts.pop()
            continue
        # диск Windows «C:» — не нормализовать как имя файла
        if len(part) == 2 and part[1] == u":":
            parts.append(part)
        else:
            parts.append(lm._lm_normalize_path_component(part, kind="path"))
    if not parts:
        return u""
    if is_abs:
        if len(parts[0]) == 2 and parts[0][1] == u":":
            path = parts[0] + u"/" + u"/".join(parts[1:]) if len(parts) > 1 else (parts[0] + u"/")
        else:
            path = u"/" + u"/".join(parts)
        path = os.path.abspath(path)
    else:
        path = _split_abs_path(u"/".join(parts), base_dir)
    dir_name = os.path.dirname(path)
    if create_folders and dir_name and not os.path.isdir(dir_name):
        try:
            os.makedirs(dir_name)
        except Exception:
            pass
    return path


def _split_resolve_sheet_name(template, field_values):
    text = lm._lm_replace_fields(template, field_values)
    return lm._lm_sanitize_sheet_name(text)


def _split_unique_sheet_name(doc, base_name, exclude=None):
    return lm._lm_pp_unique_sheet_name(doc, base_name, exclude_sheet=exclude)


def _split_store_filter_for_path(path):
    """FilterName для storeToURL по расширению файла."""
    ext = os.path.splitext(_u(path or u""))[1].casefold()
    if ext in (u".xlsx", u".xlsm"):
        return u"Calc MS Excel 2007 XML"
    if ext == u".xls":
        return u"MS Excel 97"
    if ext in (u".ods", u".fods"):
        return u"calc8"
    return u""


def _split_file_stat_note(path):
    """Краткая проверка файла на диске для журнала."""
    p = _u(path or u"")
    if p == u"":
        return u"путь пуст"
    try:
        if os.path.isfile(p):
            return u"есть, %d байт" % int(os.path.getsize(p))
        if os.path.isdir(os.path.dirname(p) or u"."):
            return u"нет файла (каталог есть)"
        return u"нет файла и каталога"
    except Exception as err:
        return u"stat: %s" % _split_err_text(err)


def _split_save_doc(doc, path=None):
    """
    Сохранить книгу. Возвращает (ok, note).
    """
    if doc is None:
        return False, u"doc=None"
    abs_path = os.path.abspath(_u(path)) if path else u""
    if abs_path:
        try:
            url = uno.systemPathToFileUrl(abs_path)
            props = [_make_prop("Overwrite", True)]
            filt = _split_store_filter_for_path(abs_path)
            if filt:
                props.append(_make_prop("FilterName", filt))
            props = tuple([p for p in props if p is not None])
            doc.storeToURL(url, props)
            try:
                if getattr(doc, "setModified", None) is not None:
                    doc.setModified(False)
            except Exception:
                pass
            return True, u"storeToURL %s [%s] → %s" % (
                abs_path,
                filt or u"auto",
                _split_file_stat_note(abs_path),
            )
        except Exception as err:
            note = u"storeToURL %s: %s" % (abs_path, _split_err_text(err))
            try:
                lm._lm_pp_execute_dispatch(doc, ".uno:Save")
                return True, note + u"; fallback .uno:Save"
            except Exception as err2:
                return False, note + u"; Save: %s" % _split_err_text(err2)
    try:
        lm._lm_pp_execute_dispatch(doc, ".uno:Save")
        return True, u".uno:Save (без path)"
    except Exception as err:
        return False, u".uno:Save: %s" % _split_err_text(err)


def _split_close_doc(doc):
    """
    Закрыть книгу UNO (после store), чтобы снять блокировку файла.
    Возвращает (ok, note).
    """
    if doc is None:
        return True, u"doc=None"
    try:
        if getattr(doc, "setModified", None) is not None:
            doc.setModified(False)
    except Exception:
        pass
    try:
        doc.close(True)
        return True, u"close"
    except Exception as err:
        try:
            doc.dispose()
            return True, u"dispose (%s)" % _split_err_text(err)
        except Exception as err2:
            return False, u"close: %s; dispose: %s" % (
                _split_err_text(err),
                _split_err_text(err2),
            )


def _split_save_and_close_doc(doc, path=None):
    """Сохранить на path и закрыть документ. Возвращает (ok, note)."""
    ok, note = _split_save_doc(doc, path)
    if not ok:
        # даже при сбое сохранения пытаемся закрыть, иначе файл останется залоченным
        cok, cnote = _split_close_doc(doc)
        return False, note + u"; " + cnote
    cok, cnote = _split_close_doc(doc)
    if cok:
        return True, note + u"; " + cnote
    return False, note + u"; " + cnote


def _split_open_or_create_book( desktop, path, append_if_exists, create_folders, base_dir, hidden=True ):
    if path in (None, u""):
        return None, u"пустой путь книги", {}
    abs_path = _split_abs_path(path, base_dir)
    info = {
        "path": abs_path,
        "existed": bool(os.path.isfile(abs_path)),
        "action": u"",
    }
    dir_name = os.path.dirname(abs_path)
    if create_folders and dir_name and not os.path.isdir(dir_name):
        try:
            os.makedirs(dir_name)
            info["mkdir"] = dir_name
        except Exception as err:
            return None, u"каталог: %s" % _split_err_text(err), info
    props = []
    if hidden:
        props.append(_make_prop("Hidden", True))
    props = tuple([p for p in props if p is not None])
    if os.path.isfile(abs_path):
        try:
            url = uno.systemPathToFileUrl(abs_path)
            doc = desktop.loadComponentFromURL(url, "_blank", 0, props)
            if doc is None:
                return None, u"не открылась: %s" % abs_path, info
            info["action"] = u"open"
            return doc, u"", info
        except Exception as err:
            return None, u"открытие: %s" % _split_err_text(err), info
    try:
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, props)
        if doc is None:
            return None, u"не создана новая книга", info
        url = uno.systemPathToFileUrl(abs_path)
        store_props = [_make_prop("Overwrite", True)]
        filt = _split_store_filter_for_path(abs_path)
        if filt:
            store_props.append(_make_prop("FilterName", filt))
        store_props = tuple([p for p in store_props if p is not None])
        doc.storeToURL(url, store_props)
        info["action"] = u"create"
        info["after_create"] = _split_file_stat_note(abs_path)
        info["filter"] = filt or u""
        if not os.path.isfile(abs_path):
            return None, u"создание: файл не появился на диске (%s)" % abs_path, info
        return doc, u"", info
    except Exception as err:
        return None, u"создание: %s" % _split_err_text(err), info


def _split_copy_column_widths(src_sheet, dst_sheet, sc, ec):
    c = int(sc)
    while c <= int(ec):
        try:
            dst_sheet.Columns[c].Width = src_sheet.Columns[c].Width
        except Exception:
            pass
        c = c + 1


def _split_copy_row_values(src_sheet, dst_sheet, src_row, dst_row, sc, ec):
    if int(src_row) == int(dst_row) and src_sheet == dst_sheet:
        return
    try:
        src_range = src_sheet.getCellRangeByPosition(int(sc), int(src_row), int(ec), int(src_row))
        data = src_range.getDataArray()
        dst_range = dst_sheet.getCellRangeByPosition(int(sc), int(dst_row), int(ec), int(dst_row))
        dst_range.setDataArray(data)
    except Exception:
        c = int(sc)
        while c <= int(ec):
            try:
                dst_sheet.getCellByPosition(c, int(dst_row)).setString(
                    lm.cell_text(src_sheet.getCellByPosition(c, int(src_row)))
                )
            except Exception:
                pass
            c = c + 1


def _split_copy_header_row(src_sheet, dst_sheet, header_row_0, sc, ec):
    _split_copy_row_values(src_sheet, dst_sheet, int(header_row_0), int(header_row_0), sc, ec)


def _split_err_text(err):
    """Текст исключения; пустой str(err) → имя класса."""
    try:
        msg = _u(err).strip()
    except Exception:
        msg = u""
    if msg:
        return msg
    try:
        return _u(err.__class__.__name__)
    except Exception:
        return u"ошибка"


def _split_insert_contents_formats_props():
    """PropertyValue для .uno:InsertContents — только форматы (Flags=T)."""
    props = []
    for name, value in (
        ("Flags", "T"),
        ("FormulaCommand", 0),
        ("SkipEmptyCells", False),
        ("Transpose", False),
        ("AsLink", False),
        ("MoveMode", 4),
        ("Overwrite", True),
    ):
        p = _make_prop(name, value)
        if p is not None:
            props.append(p)
    return tuple(props)


def _split_row_has_significant_data(sheet, row, sc, ec):
    """True, если в строке есть непустое значение (значащая строка данных)."""
    if sheet is None:
        return False
    col = int(sc)
    while col <= int(ec):
        try:
            cell = sheet.getCellByPosition(int(col), int(row))
        except Exception:
            col = col + 1
            continue
        try:
            t = int(getattr(cell, "Type", 0) or 0)
            # com.sun.star.table.CellContentType: EMPTY=0, VALUE=1, TEXT=2, FORMULA=3
            if t in (1, 2, 3):
                return True
        except Exception:
            pass
        try:
            s = _u(cell.String or u"").strip()
            if s != u"":
                return True
        except Exception:
            pass
        try:
            if float(cell.Value) != 0.0:
                return True
        except Exception:
            pass
        col = col + 1
    return False


def _split_clipboard_copy_range(doc, sheet, sc0, sr0, sc1, sr1):
    """Выделить прямоугольник и Copy в буфер."""
    if doc is None or sheet is None:
        return False
    try:
        rng = sheet.getCellRangeByPosition(int(sc0), int(sr0), int(sc1), int(sr1))
    except Exception:
        return False
    return bool(lm._lm_pp_copy_range_to_clipboard(doc, sheet, rng))


def _split_clipboard_paste_formats_only(doc, sheet, sc0, sr0, sc1, sr1):
    """
    Спецвставка только форматов из буфера на диапазон.
    Сначала InsertContents T на выделенный диапазон; fallback — pasteCellRange HARDATTR|STYLES.
    """
    if doc is None or sheet is None:
        return False
    if int(sr1) < int(sr0) or int(sc1) < int(sc0):
        return False
    try:
        dest_range = sheet.getCellRangeByPosition(int(sc0), int(sr0), int(sc1), int(sr1))
    except Exception:
        return False
    if lm._lm_pp_select_cell_range(doc, sheet, dest_range):
        if lm._lm_pp_execute_dispatch(
            doc, ".uno:InsertContents", _split_insert_contents_formats_props()
        ):
            return True
    try:
        paste_op = 0
        insert_mode = 0
        fmt_flags = 96
        try:
            from com.sun.star.sheet.CellFlags import HARDATTR, STYLES

            fmt_flags = int(HARDATTR) | int(STYLES)
        except Exception:
            pass
        sheet.pasteCellRange(
            dest_range.getRangeAddress(),
            paste_op,
            fmt_flags,
            False,
            False,
            False,
            insert_mode,
        )
        return True
    except Exception:
        return False


def _split_apply_formats_props_column(src_sheet, sample_row, col, dst_sheet, dest_sr, dest_er):
    """
    Скопировать оформление ячейки-образца на весь столбец данных приёмника.
    Надёжнее буфера: свойства ячейки → первая строка → разворот на диапазон столбца.
    """
    try:
        src_cell = src_sheet.getCellByPosition(int(col), int(sample_row))
        dst0 = dst_sheet.getCellByPosition(int(col), int(dest_sr))
    except Exception:
        return False
    try:
        lm._lm_pp_copy_cell_format_same_doc(src_cell, dst0)
    except Exception:
        return False
    if int(dest_er) <= int(dest_sr):
        return True
    try:
        r0 = dst_sheet.getCellRangeByPosition(int(col), int(dest_sr), int(col), int(dest_sr))
        r1 = dst_sheet.getCellRangeByPosition(int(col), int(dest_sr), int(col), int(dest_er))
        lm._lm_pp_copy_cell_range_format(r0, r1)
        return True
    except Exception:
        # Fallback: поячеечно
        row = int(dest_sr) + 1
        while row <= int(dest_er):
            try:
                dst = dst_sheet.getCellByPosition(int(col), int(row))
                lm._lm_pp_copy_cell_format_same_doc(src_cell, dst)
            except Exception:
                pass
            row = row + 1
        return True


def _split_apply_formats_from_first_data_row( src_doc, src_sheet, sample_row, sc, ec, dst_doc, dst_sheet, dest_sr, dest_er ):
    """
    Скопировать оформление первой значащей строки данных источника
    на диапазон данных листа-приёмника.

    Основной путь — свойства ячеек (надёжно). Буфер InsertContents T — доп. попытка.
    """
    if src_sheet is None or dst_sheet is None:
        return False
    if int(dest_er) < int(dest_sr):
        return False
    ok_cols = 0
    col = int(sc)
    while col <= int(ec):
        if _split_apply_formats_props_column(
            src_sheet, sample_row, col, dst_sheet, dest_sr, dest_er
        ):
            ok_cols = ok_cols + 1
        col = col + 1
    if ok_cols > 0:
        return True
    # Запасной путь: буфер (иногда помогает на внешних книгах)
    if src_doc is None or dst_doc is None:
        return False
    if not _split_clipboard_copy_range(
        src_doc, src_sheet, int(sc), int(sample_row), int(ec), int(sample_row)
    ):
        return False
    try:
        lm._lm_ui_yield(force=True)
    except Exception:
        pass
    return _split_clipboard_paste_formats_only(
        dst_doc, dst_sheet, int(sc), int(dest_sr), int(ec), int(dest_er)
    )


def _split_copy_sheet_to_doc(src_doc, src_sheet_name, dst_doc, dst_sheet_name):
    """
    Копия листа: в той же книге — copyByName; между книгами — importSheet + rename.
    Возвращает (ok, err, actual_name).
    """
    src_name = _u(src_sheet_name or u"").strip()
    dst_name = _u(dst_sheet_name or u"").strip()
    if src_doc is None or dst_doc is None:
        return False, u"книга не задана", u""
    if src_name == u"" or dst_name == u"":
        return False, u"пустое имя листа", u""
    actual = _split_unique_sheet_name(dst_doc, dst_name)
    if src_doc is dst_doc:
        try:
            if src_name.casefold() == actual.casefold():
                actual = _split_unique_sheet_name(dst_doc, actual + u"_1")
            pos = int(dst_doc.Sheets.getCount())
            dst_doc.Sheets.copyByName(src_name, actual, pos)
            if lm._lm_pp_get_sheet_ref(dst_doc, actual) is None:
                return False, u"copyByName: лист не появился", u""
            return True, u"", actual
        except Exception as err:
            return False, u"copyByName: %s" % _split_err_text(err), u""
    try:
        t = uno.getTypeByName("com.sun.star.sheet.XSpreadsheets2")
        sheets2 = dst_doc.Sheets.queryInterface(t)
    except Exception as err:
        return False, u"XSpreadsheets2: %s" % _split_err_text(err), u""
    if not sheets2:
        return False, u"XSpreadsheets2 недоступен", u""
    try:
        if not src_doc.Sheets.hasByName(src_name):
            return False, u"нет листа-источника «%s»" % src_name, u""
    except Exception as err:
        return False, u"источник: %s" % _split_err_text(err), u""
    insert_index = int(dst_doc.Sheets.getCount())
    try:
        sheets2.importSheet(src_doc, str(src_name), insert_index)
    except Exception as err:
        return False, u"importSheet: %s" % _split_err_text(err), u""
    new_sheet = lm._lm_pp_get_sheet_ref(dst_doc, src_name)
    if new_sheet is None:
        # иногда importSheet даёт уникальное имя
        try:
            new_sheet = dst_doc.Sheets.getByIndex(insert_index)
        except Exception:
            new_sheet = None
    if new_sheet is None:
        return False, u"importSheet: лист не появился", u""
    imported_name = _u(getattr(new_sheet, "Name", u"") or u"")
    if imported_name.casefold() != actual.casefold():
        try:
            # если actual уже занят — уникализировать ещё раз
            if dst_doc.Sheets.hasByName(actual) and imported_name.casefold() != actual.casefold():
                actual = _split_unique_sheet_name(dst_doc, actual)
            new_sheet.Name = actual
        except Exception as err:
            return False, u"переименование «%s»→«%s»: %s" % (
                imported_name,
                actual,
                _split_err_text(err),
            ), u""
    else:
        actual = imported_name
    if lm._lm_pp_get_sheet_ref(dst_doc, actual) is None:
        return False, u"лист «%s» не найден после копирования" % actual, u""
    return True, u"", actual


def _split_delete_rows_except(sheet, keep_rows):
    if sheet is None:
        return
    keep = set(int(r) for r in (keep_rows or []))
    if len(keep) == 0:
        return
    try:
        _sc, sr, _ec, er = lm._lm_vlookup_sheet_used_area(sheet)
    except Exception:
        return
    r = int(er)
    while r >= int(sr):
        if r not in keep:
            try:
                sheet.Rows.removeByIndex(r, 1)
            except Exception:
                pass
        r = r - 1


def _split_sheet_names(doc):
    names = []
    if doc is None:
        return names
    try:
        n = int(doc.Sheets.getCount())
    except Exception:
        return names
    i = 0
    while i < n:
        try:
            names.append(_u(doc.Sheets.getByIndex(i).Name))
        except Exception:
            pass
        i = i + 1
    return names


def _split_remove_non_created_sheets(doc, keep_names):
    """
    Удалить с книги все листы, кроме созданных при разделении
    (Sheet1 / Лист1 и пр. из private:factory).
    В Calc должен остаться хотя бы один лист.
    Возвращает список удалённых имён.
    """
    removed = []
    if doc is None:
        return removed
    keep_cf = set()
    for n in keep_names or []:
        s = _u(n or u"").strip()
        if s:
            keep_cf.add(s.casefold())
    if not keep_cf:
        return removed
    for name in list(_split_sheet_names(doc)):
        if name.casefold() in keep_cf:
            continue
        try:
            if int(doc.Sheets.getCount()) <= 1:
                break
            doc.Sheets.removeByName(name)
            removed.append(name)
        except Exception:
            pass
    return removed


def _split_effective_header_row(sheet, requested_header_0, copy_header):
    """Строка заголовка после переноса (после copy_sheet+удаления строк шапка обычно вверху)."""
    try:
        _sc, sr, _ec, er = lm._lm_vlookup_sheet_used_area(sheet)
    except Exception:
        return int(requested_header_0 or 0)
    if er < sr:
        return int(requested_header_0 or 0)
    req = int(requested_header_0 or 0)
    if copy_header:
        if req >= int(sr) and req <= int(er):
            return req
        return int(sr)
    return max(int(sr), req)


def _split_clear_database_ranges(doc, keep_anonymous=False):
    """
    Удалить DatabaseRanges книги.
    keep_anonymous=True — оставить __Anonymous_Sheet_DB__* (как у меню Data→Автофильтр).
    """
    removed = []
    if doc is None:
        return removed
    try:
        dbr = doc.DatabaseRanges
    except Exception:
        return removed
    names = []
    try:
        n = int(dbr.getCount())
        i = 0
        while i < n:
            try:
                names.append(_u(dbr.getByIndex(i).Name))
            except Exception:
                pass
            i = i + 1
    except Exception:
        try:
            names = list(dbr.getElementNames())
        except Exception:
            names = []
    for name in names:
        if keep_anonymous and _u(name).startswith(u"__Anonymous_Sheet_DB__"):
            continue
        try:
            db = dbr.getByName(name)
            try:
                db.AutoFilter = False
            except Exception:
                try:
                    db.setPropertyValue("AutoFilter", False)
                except Exception:
                    pass
            dbr.removeByName(name)
            removed.append(name)
        except Exception:
            pass
    return removed


def _split_doc_window_set_visible(doc, visible):
    """Показать/скрыть окно книги."""
    if doc is None:
        return False
    try:
        ctrl = doc.getCurrentController()
        if ctrl is None:
            return False
        frame = ctrl.getFrame()
        if frame is None:
            return False
        win = frame.getContainerWindow()
        if win is None:
            return False
        if hasattr(win, "setVisible"):
            win.setVisible(bool(visible))
        elif hasattr(win, "IsVisible"):
            win.IsVisible = bool(visible)
        else:
            return False
        return True
    except Exception:
        return False


def _split_activate_document(doc):
    """
    Полноценно показать книгу на экране (activate + visible + toFront).
    Нужно для freezeAtPosition и .uno:DataFilterAutoFilter.
    """
    if doc is None:
        return False
    ok = False
    try:
        ctrl = doc.getCurrentController()
    except Exception:
        ctrl = None
    frame = None
    if ctrl is not None:
        try:
            frame = ctrl.getFrame()
        except Exception:
            frame = None
    if frame is not None:
        try:
            frame.activate()
            ok = True
        except Exception:
            pass
        try:
            win = frame.getContainerWindow()
            if win is not None:
                for attr, val in (("IsMinimized", False), ("Minimized", False)):
                    if hasattr(win, attr):
                        try:
                            setattr(win, attr, val)
                        except Exception:
                            pass
                try:
                    win.setVisible(True)
                    ok = True
                except Exception:
                    try:
                        win.IsVisible = True
                        ok = True
                    except Exception:
                        pass
                if hasattr(win, "toFront"):
                    try:
                        win.toFront()
                        ok = True
                    except Exception:
                        pass
        except Exception:
            pass
    try:
        lm._lm_ui_yield(force=True)
    except Exception:
        pass
    return ok


def _split_window_get_bool(win, names):
    for name in names:
        try:
            return bool(getattr(win, name))
        except Exception:
            pass
    return None


def _split_window_set_bool(win, names, value):
    ok = False
    val = bool(value)
    for name in names:
        try:
            setattr(win, name, val)
            ok = True
        except Exception:
            pass
    # Прямые имена XTopWindow2 — надёжнее hasattr/setattr на pyuno
    if "IsMaximized" in names or "Maximized" in names:
        try:
            win.IsMaximized = val
            ok = True
        except Exception:
            pass
        try:
            win.Maximized = val
            ok = True
        except Exception:
            pass
    if "IsMinimized" in names or "Minimized" in names:
        try:
            win.IsMinimized = val
            ok = True
        except Exception:
            pass
        try:
            win.Minimized = val
            ok = True
        except Exception:
            pass
    return ok


def _split_maximize_via_workarea(win):
    """Запасной путь: растянуть окно на рабочую область экрана."""
    if win is None:
        return False
    wa = None
    try:
        toolkit = win.getToolkit()
    except Exception:
        toolkit = None
    if toolkit is not None:
        for meth in ("getWorkArea", "getDisplayWorkArea"):
            if not hasattr(toolkit, meth):
                continue
            try:
                wa = getattr(toolkit, meth)()
                if wa is not None:
                    break
            except Exception:
                wa = None
    if wa is None:
        try:
            ctx = uno.getComponentContext()
            sm = ctx.getServiceManager()
            tk = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
            if tk is not None and hasattr(tk, "getWorkArea"):
                wa = tk.getWorkArea()
        except Exception:
            wa = None
    if wa is None:
        return False
    try:
        from com.sun.star.awt.PosSize import POSSIZE
    except Exception:
        POSSIZE = 15
    try:
        x = int(getattr(wa, "X", 0) or 0)
        y = int(getattr(wa, "Y", 0) or 0)
        w = int(getattr(wa, "Width", 0) or 0)
        h = int(getattr(wa, "Height", 0) or 0)
        if w <= 0 or h <= 0:
            return False
        win.setPosSize(x, y, w, h, POSSIZE)
        return True
    except Exception:
        return False


# Документ, который нужно максимизировать после закрытия UI сбора (диалог прогресса).
_SPLIT_MAXIMIZE_PENDING_DOC = None


def _split_request_maximize_pending(doc):
    """Отложить повторную максимизацию до конца collect (после диалога прогресса)."""
    global _SPLIT_MAXIMIZE_PENDING_DOC
    if doc is not None:
        _SPLIT_MAXIMIZE_PENDING_DOC = doc


def _split_flush_maximize_pending():
    """Выполнить отложенную максимизацию (конец collect_workbooks_run)."""
    global _SPLIT_MAXIMIZE_PENDING_DOC
    doc = _SPLIT_MAXIMIZE_PENDING_DOC
    _SPLIT_MAXIMIZE_PENDING_DOC = None
    if doc is None:
        return False
    return _split_maximize_document(doc)


def _split_maximize_document(doc):
    """
    Максимизировать окно книги (эта_книга после разделения / конец сбора).

    .uno:Maximize на части сборок — toggle; hasattr/setattr по UNO ненадёжны.
    Важно: dispatch Maximize только если IsMaximized явно False.
    Если свойство не читается (None), а окно уже развёрнуто — toggle
    снимет maximize (типичный баг «в конце окно не максимизировано»).

    Порядок: activate → снять minimize → IsMaximized=True → (если явно
    не maximized) dispatch / workarea → yield → повтор IsMaximized.
    """
    if doc is None:
        return False
    _split_activate_document(doc)
    ok = False
    win = None
    try:
        ctrl = doc.getCurrentController()
        frame = ctrl.getFrame() if ctrl is not None else None
        if frame is not None:
            try:
                frame.activate()
            except Exception:
                pass
            win = frame.getContainerWindow()
    except Exception:
        win = None
    if win is None:
        return False
    _split_window_set_bool(win, ("IsMinimized", "Minimized"), False)
    try:
        win.setVisible(True)
        ok = True
    except Exception:
        pass
    if _split_window_set_bool(win, ("IsMaximized", "Maximized"), True):
        ok = True
    already = _split_window_get_bool(win, ("IsMaximized", "Maximized"))
    # Dispatch только при явном False. None = «не знаем» — не трогать toggle.
    if already is False:
        for slot in (".uno:Maximize", ".uno:ShowMaxFrame"):
            try:
                if lm._lm_pp_execute_dispatch(doc, slot):
                    ok = True
                    break
            except Exception:
                pass
        if _split_window_get_bool(win, ("IsMaximized", "Maximized")) is not True:
            if _split_window_set_bool(win, ("IsMaximized", "Maximized"), True):
                ok = True
    if already is False and _split_window_get_bool(win, ("IsMaximized", "Maximized")) is not True:
        # setPosSize выводит из maximize — только если точно не maximized.
        if _split_maximize_via_workarea(win):
            ok = True
            _split_window_set_bool(win, ("IsMaximized", "Maximized"), True)
    try:
        if hasattr(win, "toFront"):
            win.toFront()
    except Exception:
        pass
    try:
        # На части сборок AO/LO флаг maximize меняется, а геометрия окна — нет.
        # Дотягиваем окно до workarea даже при уже выставленном IsMaximized=True.
        if _split_window_get_bool(win, ("IsMaximized", "Maximized")) is True:
            if _split_maximize_via_workarea(win):
                ok = True
                _split_window_set_bool(win, ("IsMaximized", "Maximized"), True)
    except Exception:
        pass
    try:
        lm._lm_ui_yield(force=True)
    except Exception:
        pass
    if _split_window_get_bool(win, ("IsMaximized", "Maximized")) is False:
        if _split_window_set_bool(win, ("IsMaximized", "Maximized"), True):
            ok = True
        elif _split_maximize_via_workarea(win):
            ok = True
            _split_window_set_bool(win, ("IsMaximized", "Maximized"), True)
    return ok


def _split_document_is_maximized(doc):
    """True/False/None — состояние IsMaximized окна книги."""
    if doc is None:
        return None
    try:
        ctrl = doc.getCurrentController()
        frame = ctrl.getFrame() if ctrl is not None else None
        win = frame.getContainerWindow() if frame is not None else None
    except Exception:
        return None
    return _split_window_get_bool(win, ("IsMaximized", "Maximized"))


def _split_restore_document_maximized(doc, was_maximized):
    """
    Вернуть maximize, только если до сбора окно было развёрнуто
    (или состояние неизвестно — тогда безопасный ensure без toggle).
    """
    if doc is None:
        return False
    if was_maximized is False:
        return False
    now = _split_document_is_maximized(doc)
    if now is True:
        return True
    return _split_maximize_document(doc)


def _split_reopen_book_visible(desktop, path):
    """
    Открыть книгу с диска видимой (Hidden=False), как reopen xml-фазы сбора.
    Возвращает (doc, err).
    """
    if desktop is None:
        return None, u"нет Desktop"
    abs_path = os.path.abspath(_u(path or u""))
    if abs_path == u"" or not os.path.isfile(abs_path):
        return None, u"нет файла: %s" % abs_path
    try:
        url = uno.systemPathToFileUrl(abs_path)
    except Exception as err:
        return None, u"url: %s" % _split_err_text(err)
    props = tuple(
        [
            p
            for p in (
                _make_prop("Hidden", False),
                _make_prop("ReadOnly", False),
            )
            if p is not None
        ]
    )
    try:
        # _blank — не подменять активную книгу сбора
        doc = desktop.loadComponentFromURL(url, "_blank", 0, props)
    except Exception as err:
        return None, u"load: %s" % _split_err_text(err)
    if doc is None:
        return None, u"load вернул None: %s" % abs_path
    _split_activate_document(doc)
    try:
        lm._lm_ui_yield(force=True)
    except Exception:
        pass
    return doc, u""


def _split_select_a1(doc, sheet):
    """Курсор на A1 активного листа."""
    if doc is None or sheet is None:
        return False
    try:
        cell = sheet.getCellByPosition(0, 0)
        return bool(lm._lm_pp_select_cell_range(doc, sheet, cell))
    except Exception:
        return False


def _split_apply_menu_autofilter(doc, sheet, sc, hdr, ec, er):
    """
    Автофильтр как в меню (Data → AutoFilter), без AF_/смарт-таблиц.

    Детект и включение через UnnamedDatabaseRanges (не toggle-dispatch).
    Если фильтр уже есть — no-op. Пустые заголовки заполняются.
    """
    if doc is None or sheet is None:
        return False
    try:
        ensure = getattr(lm, "_lm_pp_ensure_sheet_menu_autofilter", None)
        if callable(ensure):
            return bool(ensure(doc, sheet, int(sc), int(hdr), int(ec), int(er)))
    except Exception:
        pass
    # Fallback на старый путь (детект + dispatch).
    try:
        if lm._lm_pp_sheet_has_active_autofilter(doc, sheet):
            return True
    except Exception:
        pass
    try:
        data_range = sheet.getCellRangeByPosition(int(sc), int(hdr), int(ec), int(er))
    except Exception:
        return False

    def _apply():
        lm._lm_pp_fill_empty_header_cells(sheet, int(hdr), int(sc), int(ec))
        if not lm._lm_pp_select_cell_range(doc, sheet, data_range):
            return False
        return bool(lm._lm_pp_execute_dispatch(doc, ".uno:DataFilterAutoFilter"))

    return bool(lm._lm_pp_run_with_suppressed_ui(doc, _apply))


def _split_restore_sheet_menu_autofilter(doc, sheet, header_row_0, copy_header=True):
    """
    Поставить меню-автофильтр на лист (как на листах разделения).
    Без смены оформления заголовка / autofit / freeze.
    """
    if doc is None or sheet is None:
        return u"нет листа"
    try:
        sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(sheet)
    except Exception as err:
        return u"used_area: %s" % _split_err_text(err)
    if er < sr or ec < sc:
        return u"пусто"
    hdr = _split_effective_header_row(sheet, header_row_0, copy_header)
    if hdr < sr:
        hdr = int(sr)
    if hdr > er:
        hdr = int(sr)
    try:
        _split_activate_document(doc)
        ctrl = doc.getCurrentController()
        if ctrl is not None:
            try:
                ctrl.setActiveSheet(sheet)
            except Exception:
                pass
    except Exception:
        pass
    try:
        if _split_apply_menu_autofilter(doc, sheet, sc, hdr, ec, er):
            return u"автофильтр(меню)"
        return u"автофильтр: сбой dispatch"
    except Exception as err:
        return u"автофильтр: %s" % _split_err_text(err)


def _split_apply_basic_sheet_formatting( doc, sheet, header_row_0, copy_header=True, manage_window=True, clear_db=True ):
    """
    Базовое оформление листа после разделения:
    серый жирный заголовок, автоширина, закрепление шапки (нужно видимое окно),
    автофильтр через меню-dispatch (анонимный диапазон, без AF_*).
    Курсор в конце на A1.
    manage_window=False — окно уже показано снаружи (пакетная финализация книги).
    clear_db=False — не трогать DatabaseRanges (уже очищены снаружи).
    """
    if sheet is None:
        return u"нет листа"
    try:
        sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(sheet)
    except Exception as err:
        return u"used_area: %s" % _split_err_text(err)
    if er < sr or ec < sc:
        return u"пусто"
    hdr = _split_effective_header_row(sheet, header_row_0, copy_header)
    if hdr < sr:
        hdr = int(sr)
    if hdr > er:
        hdr = int(sr)
    notes = []
    if manage_window:
        try:
            _split_activate_document(doc)
        except Exception:
            pass
    if clear_db:
        try:
            dropped = _split_clear_database_ranges(doc, keep_anonymous=False)
            try:
                lm._lm_pp_remove_sheet_autofilters(doc, sheet)
            except Exception:
                pass
            if dropped:
                notes.append(u"сняты БД:%d" % len(dropped))
        except Exception as err:
            notes.append(u"БД: %s" % _split_err_text(err))
    try:
        lm._lm_apply_default_header_row_appearance(sheet, int(ec), int(hdr))
        notes.append(u"заголовок")
    except Exception as err:
        notes.append(u"заголовок: %s" % _split_err_text(err))
    try:
        cols = list(range(int(sc), int(ec) + 1))
        lm._lm_pp_autofit_columns(doc, sheet, cols, sr=int(hdr), er=int(er))
        notes.append(u"ширина")
    except Exception as err:
        notes.append(u"ширина: %s" % _split_err_text(err))
    try:
        ctrl = doc.getCurrentController() if doc is not None else None
        if ctrl is not None:
            try:
                ctrl.setActiveSheet(sheet)
            except Exception:
                pass
            try:
                ctrl.FirstVisibleColumn = 0
                ctrl.FirstVisibleRow = 0
            except Exception:
                pass
        lm._lm_pp_freeze_sheet_header(doc, sheet, int(hdr))
        frozen_ok = False
        try:
            if ctrl is not None and hasattr(ctrl, "hasFrozenPanes"):
                frozen_ok = bool(ctrl.hasFrozenPanes())
            else:
                frozen_ok = True
        except Exception:
            frozen_ok = True
        notes.append(u"закрепление" + (u"" if frozen_ok else u"(не подтверждено)"))
    except Exception as err:
        notes.append(u"закрепление: %s" % _split_err_text(err))
    try:
        if _split_apply_menu_autofilter(doc, sheet, sc, hdr, ec, er):
            notes.append(u"автофильтр(меню)")
        else:
            notes.append(u"автофильтр: сбой dispatch")
    except Exception as err:
        notes.append(u"автофильтр: %s" % _split_err_text(err))
    try:
        if _split_select_a1(doc, sheet):
            notes.append(u"A1")
    except Exception as err:
        notes.append(u"A1: %s" % _split_err_text(err))
    # manage_window: окно не прячем — caller закрывает документ после store
    return u", ".join(notes)


class _SplitWriter(object):
    def __init__( self, result_doc, src_sheet, spec, sc, ec, header_row_0, by_name, run_context, desktop, base_dir):
        self.result_doc = result_doc
        self.src_sheet = src_sheet
        self.spec = spec
        self.sc = int(sc)
        self.ec = int(ec)
        self.header_row_0 = int(header_row_0)
        self.by_name = by_name
        self.run_context = run_context or {}
        self.desktop = desktop
        self.base_dir = base_dir
        self.this_book = _is_this_book_marker(spec.get("book_template"))
        self.split_books = [] if self.this_book else list(spec.get("split_books") or [])
        self.split_sheets = list(spec.get("split_sheets") or [])
        self.copy_header = bool(spec.get("copy_header", True))
        self.overwrite_sheets = bool(spec.get("overwrite_sheets", True))
        self.append_if_exists = bool(spec.get("append_if_exists", False))
        self.save_immediately = bool(spec.get("save_immediately", False))
        self.create_folders = bool(spec.get("create_folders", True))
        self.split_mode = _u(spec.get("split_mode") or u"copy_range").casefold()
        self.paste_formats = bool(spec.get("paste_formats", False))
        self.book_template = _u(spec.get("book_template") or u"эта_книга")
        self.sheet_template = _u(spec.get("sheet_template") or u"[Split]")
        self.open_books = {}
        self.current_book_key = None
        self.current_book_doc = result_doc
        self.current_book_path = u""
        self.current_sheet_key = None
        self.current_sheet = None
        self.current_dest_row = None
        self.current_block_rows = []
        # copy_range + paste_formats: границы данных на текущем листе-приёмнике
        self._cr_dest_sr = None
        self._cr_dest_er = None
        # (sheet_name, dest_sr, dest_er) — для повторного paste_formats после оформления
        self._paste_fmt_targets = []
        self.created_sheet_names = []
        self.created_book_paths = []
        # id(doc) / path → set имён листов, созданных разделением
        self.book_created_sheets = {}
        self.book_created_sheets_by_path = {}
        self._log_phase = u"диапазон"
        self._log_label = _u(self.run_context.get("label") or u"разделить_листы")

    def _dbg(self, status, note, sheet_name=None):
        sn = sheet_name
        if sn is None:
            try:
                sn = _u(getattr(self.src_sheet, "Name", u"") or u"")
            except Exception:
                sn = u""
        _split_log(
            self.run_context,
            self.result_doc,
            sn,
            self._log_phase,
            self._log_label,
            status,
            note,
        )

    def _track_created_sheet(self, doc, sheet_name):
        if doc is None:
            return
        name = _u(sheet_name or u"").strip()
        if name == u"":
            return
        key = id(doc)
        if key not in self.book_created_sheets:
            self.book_created_sheets[key] = set()
        self.book_created_sheets[key].add(name)
        path = _u(self.current_book_path or u"").strip()
        if path:
            if path not in self.book_created_sheets_by_path:
                self.book_created_sheets_by_path[path] = set()
            self.book_created_sheets_by_path[path].add(name)
        if name not in self.created_sheet_names:
            self.created_sheet_names.append(name)

    def _keep_sheets_for_doc(self, doc, path=None):
        keep = set(self.book_created_sheets.get(id(doc), set()) or [])
        p = _u(path or self.current_book_path or u"").strip()
        if p and p in self.book_created_sheets_by_path:
            keep |= set(self.book_created_sheets_by_path.get(p) or [])
        return list(keep)

    def _book_key(self, field_values):
        if self.this_book:
            return u"__THIS_BOOK__"
        vals = _split_row_key_values(
            self.src_sheet, self.current_row, self.split_books, self.by_name
        )
        return lm._lm_vlookup_composite_key(vals)

    def _sheet_key(self, field_values):
        vals = _split_row_key_values(
            self.src_sheet, self.current_row, self.split_sheets, self.by_name
        )
        return lm._lm_vlookup_composite_key(vals)

    def _ensure_book(self, book_key, field_values):
        if self.this_book:
            self.current_book_doc = self.result_doc
            self.current_book_path = u""
            self._dbg(u"инфо", u"книга: эта_книга (внешние файлы не создаются)")
            return True, u""
        path = _split_resolve_book_path(
            self.book_template, field_values, self.base_dir, self.create_folders
        )
        self._dbg(
            u"инфо",
            u"книга key=%s template=%s → path=%s fields=%s"
            % (book_key, self.book_template, path or u"(пусто)", field_values),
        )
        if path == u"":
            return False, u"пустой путь книги"
        if book_key in self.open_books:
            entry = self.open_books[book_key]
            if entry.get("doc") is not None and not entry.get("closed"):
                self.current_book_doc = entry.get("doc")
                self.current_book_path = entry.get("path") or path
                self._dbg(u"инфо", u"книга уже открыта: %s" % self.current_book_path)
                return True, u""
            # была закрыта (save_immediately) — откроем снова
            self._dbg(u"инфо", u"книга была закрыта, повторное открытие: %s" % path)
        doc, err, info = _split_open_or_create_book(
            self.desktop,
            path,
            self.append_if_exists,
            self.create_folders,
            self.base_dir,
            hidden=True,
        )
        if doc is None:
            self._dbg(
                u"ошибка",
                u"книга не создана/не открыта: %s | info=%s" % (err or u"книга", info),
            )
            return False, err or u"книга"
        self.open_books[book_key] = {"doc": doc, "path": path, "closed": False}
        self.current_book_doc = doc
        self.current_book_path = path
        # перенести keep-листы с path на новый id(doc)
        prev_keep = set(self.book_created_sheets_by_path.get(path, set()) or [])
        if prev_keep:
            self.book_created_sheets[id(doc)] = set(prev_keep)
        if path not in self.created_book_paths:
            self.created_book_paths.append(path)
        self._dbg(
            u"инфо",
            u"книга %s: %s | на диске: %s"
            % (
                info.get("action") or u"ok",
                path,
                info.get("after_create") or _split_file_stat_note(path),
            ),
        )
        return True, u""

    def _remove_sheet_if_exists(self, doc, name):
        try:
            sheets = doc.Sheets
            if sheets.hasByName(name):
                if self.overwrite_sheets:
                    sheets.removeByName(name)
                    return None
                return sheets.getByName(name)
        except Exception:
            pass
        return None

    def _prepare_sheet_copy_range(self, doc, sheet_name):
        existing = self._remove_sheet_if_exists(doc, sheet_name)
        if existing is not None and not self.overwrite_sheets:
            self.current_sheet = existing
            self.current_dest_row = int(self.header_row_0) + (1 if self.copy_header else 0)
            if self.copy_header:
                try:
                    _ec = self.ec
                    if self.current_dest_row > int(self.header_row_0):
                        last = self.current_dest_row - 1
                        if last >= int(self.header_row_0):
                            doc_rows = last - int(self.header_row_0)
                            if doc_rows > 0:
                                pass
                except Exception:
                    pass
            try:
                _sc, _sr, _ec, er = lm._lm_vlookup_sheet_used_area(existing)
                self.current_dest_row = max(int(er) + 1, self.current_dest_row)
            except Exception:
                pass
            self._cr_dest_sr = int(self.current_dest_row)
            self._cr_dest_er = None
            return True, u""
        actual = _split_unique_sheet_name(doc, sheet_name)
        try:
            pos = int(doc.Sheets.getCount())
            doc.Sheets.insertNewByName(actual, pos)
        except Exception as err:
            return False, u"лист: %s" % err
        dst = lm._lm_pp_get_sheet_ref(doc, actual)
        if dst is None:
            return False, u"лист не создан"
        try:
            dst.PageStyle = self.src_sheet.PageStyle
        except Exception:
            pass
        _split_copy_column_widths(self.src_sheet, dst, self.sc, self.ec)
        if self.copy_header:
            _split_copy_header_row(self.src_sheet, dst, self.header_row_0, self.sc, self.ec)
        self.current_sheet = dst
        self.current_dest_row = int(self.header_row_0) + (1 if self.copy_header else 0)
        self._cr_dest_sr = int(self.current_dest_row)
        self._cr_dest_er = None
        self._track_created_sheet(doc, actual)
        if doc == self.result_doc:
            _split_notify_sheet_created(
                self.run_context,
                doc,
                actual,
                self.header_row_0,
                self.src_sheet.Name,
            )
        return True, u""

    def _prepare_sheet_copy_sheet(self, doc, sheet_name, block_rows):
        existing = self._remove_sheet_if_exists(doc, sheet_name)
        if existing is not None and not self.overwrite_sheets:
            self.current_sheet = existing
            self.current_dest_row = None
            self._track_created_sheet(doc, _u(getattr(existing, "Name", sheet_name) or sheet_name))
            return True, u""
        src_doc = self.result_doc
        src_name = str(self.src_sheet.Name)
        ok, err, actual = _split_copy_sheet_to_doc(src_doc, src_name, doc, sheet_name)
        if not ok:
            return False, u"copy_sheet: %s" % (err or u"сбой")
        self._dbg(
            u"инфо",
            u"copy_sheet «%s» → «%s» в %s (строк блока %d)"
            % (
                src_name,
                actual,
                self.current_book_path or u"эта_книга",
                len(block_rows or []),
            ),
        )
        dst = lm._lm_pp_get_sheet_ref(doc, actual)
        if dst is None:
            return False, u"copy_sheet: лист не найден"
        keep = set(int(r) for r in block_rows)
        if self.copy_header:
            keep.add(int(self.header_row_0))
        _split_delete_rows_except(dst, keep)
        self.current_sheet = dst
        self.current_dest_row = None
        self._track_created_sheet(doc, actual)
        if doc == self.result_doc:
            _split_notify_sheet_created(
                self.run_context,
                doc,
                actual,
                self.header_row_0,
                self.src_sheet.Name,
            )
        return True, u""

    def _flush_block_copy_sheet(self):
        if not self.current_block_rows:
            return True, u""
        field_values = _split_row_field_values(
            self.src_sheet, self.current_block_rows[0], self.split_sheets, self.by_name
        )
        for name in self.split_books:
            field_values[name] = _split_row_field_values(
                self.src_sheet, self.current_block_rows[0], [name], self.by_name
            ).get(name, u"")
        sheet_name = _split_resolve_sheet_name(self.sheet_template, field_values)
        ok, err = self._prepare_sheet_copy_sheet(
            self.current_book_doc, sheet_name, list(self.current_block_rows)
        )
        self.current_block_rows = []
        return ok, err

    def _append_row_copy_range(self, row):
        if self.current_sheet is None or self.current_dest_row is None:
            return False, u"нет листа-приёмника"
        _split_copy_row_values(
            self.src_sheet,
            self.current_sheet,
            int(row),
            int(self.current_dest_row),
            self.sc,
            self.ec,
        )
        self._cr_dest_er = int(self.current_dest_row)
        self.current_dest_row = int(self.current_dest_row) + 1
        return True, u""

    def _sample_data_row_0(self):
        """Первая значащая строка данных на листе-источнике (после заголовка)."""
        start = int(self.header_row_0) + 1
        er = start
        try:
            _sc, sr, _ec, er = lm._lm_vlookup_sheet_used_area(self.src_sheet)
            er = int(er)
            if start < int(sr):
                start = int(sr)
            if self.copy_header and start <= int(self.header_row_0):
                start = int(self.header_row_0) + 1
            if start > er:
                return None
        except Exception:
            er = start + 5000
        row = int(start)
        while row <= int(er):
            if _split_row_has_significant_data(self.src_sheet, row, self.sc, self.ec):
                return int(row)
            row = row + 1
        # Нет непустых — всё же взять первую после заголовка, если она в used area
        if start <= int(er):
            return int(start)
        return None

    def _flush_paste_formats_copy_range(self):
        """
        После заполнения листа (copy_range): оформление 1-й значащей строки источника
        → на диапазон данных приёмника (свойства ячеек; буфер — запасной путь).
        """
        if not self.paste_formats or self.split_mode != u"copy_range":
            return True, u""
        if self.current_sheet is None:
            return True, u""
        if self._cr_dest_sr is None or self._cr_dest_er is None:
            return True, u""
        if int(self._cr_dest_er) < int(self._cr_dest_sr):
            return True, u""
        sample = self._sample_data_row_0()
        if sample is None:
            self._dbg(u"инфо", u"paste_formats: нет строки-образца в источнике")
            return True, u""
        dst_doc = self.current_book_doc or self.result_doc
        name = u""
        try:
            name = _u(getattr(self.current_sheet, "Name", u"") or u"")
        except Exception:
            name = u""
        self._dbg(
            u"инфо",
            u"paste_formats «%s»: образец стр.%d → %d..%d (cols %d..%d)"
            % (
                name,
                int(sample) + 1,
                int(self._cr_dest_sr) + 1,
                int(self._cr_dest_er) + 1,
                int(self.sc) + 1,
                int(self.ec) + 1,
            ),
        )
        ok = _split_apply_formats_from_first_data_row(
            self.result_doc,
            self.src_sheet,
            sample,
            self.sc,
            self.ec,
            dst_doc,
            self.current_sheet,
            self._cr_dest_sr,
            self._cr_dest_er,
        )
        if ok:
            self._dbg(u"ok", u"paste_formats «%s»: форматы применены" % name)
            try:
                self._paste_fmt_targets.append(
                    (name, int(self._cr_dest_sr), int(self._cr_dest_er))
                )
            except Exception:
                pass
        else:
            self._dbg(
                u"предупреждение",
                u"paste_formats «%s»: не удалось вставить форматы" % name,
            )
        return True, u""

    def _reapply_paste_formats_after_formatting(self, doc):
        """
        Повторно наложить форматы после базового оформления листов
        (autofit/заголовок не должны затирать NumberFormat, но буфер мог не сработать —
         здесь ещё раз надёжный props-путь).
        """
        if not self.paste_formats or self.split_mode != u"copy_range":
            return
        if doc is None or self.src_sheet is None:
            return
        sample = self._sample_data_row_0()
        if sample is None:
            return
        seen = set()
        for item in list(self._paste_fmt_targets or []):
            try:
                name, dest_sr, dest_er = item
            except Exception:
                continue
            key = _u(name or u"").casefold()
            if key == u"" or key in seen:
                continue
            seen.add(key)
            sh = lm._lm_pp_get_sheet_ref(doc, name)
            if sh is None:
                continue
            ok = _split_apply_formats_from_first_data_row(
                self.result_doc,
                self.src_sheet,
                sample,
                self.sc,
                self.ec,
                doc,
                sh,
                int(dest_sr),
                int(dest_er),
            )
            self._dbg(
                u"ok" if ok else u"предупреждение",
                u"paste_formats (после оформления) «%s»: %s"
                % (name, u"ok" if ok else u"сбой"),
            )

    def _cleanup_external_book_sheets(self, doc, path):
        """Удалить чужие листы и именованные БД (пока книга ещё Hidden). Возвращает keep."""
        if doc is None or self.this_book:
            return []
        keep = self._keep_sheets_for_doc(doc, path)
        before = _split_sheet_names(doc)
        removed = _split_remove_non_created_sheets(doc, keep)
        after = _split_sheet_names(doc)
        self._dbg(
            u"инфо",
            u"очистка листов %s: было %s; keep=%s; удалены=%s; стало %s"
            % (
                path or u"",
                before,
                keep,
                removed,
                after,
            ),
        )
        try:
            dropped = _split_clear_database_ranges(doc, keep_anonymous=False)
            if dropped:
                self._dbg(u"инфо", u"сняты DatabaseRanges: %s" % u", ".join(dropped))
        except Exception as err:
            self._dbg(u"инфо", u"очистка БД: %s" % _split_err_text(err))
        return keep

    def _apply_visible_book_formatting(self, doc, keep):
        """
        Freeze / меню-автофильтр / A1 — только на полноценно видимой книге.
        """
        if doc is None:
            return
        _split_activate_document(doc)
        first_sheet = None
        for name in keep or []:
            sh = lm._lm_pp_get_sheet_ref(doc, name)
            if sh is None:
                continue
            if first_sheet is None:
                first_sheet = sh
            # каждый лист — снова activate (freeze/dispatch чувствительны к фокусу)
            _split_activate_document(doc)
            note = _split_apply_basic_sheet_formatting(
                doc,
                sh,
                self.header_row_0,
                self.copy_header,
                manage_window=False,
                clear_db=False,
            )
            self._dbg(u"инфо", u"оформление «%s»: %s" % (name, note))
        # эта_книга: очистка DatabaseRanges снимает автофильтр и с источника — вернуть.
        if self.this_book and self.src_sheet is not None:
            try:
                src_name = _u(getattr(self.src_sheet, "Name", u"") or u"")
            except Exception:
                src_name = u""
            if src_name and src_name not in (keep or []):
                _split_activate_document(doc)
                note_src = _split_restore_sheet_menu_autofilter(
                    doc,
                    self.src_sheet,
                    self.header_row_0,
                    self.copy_header,
                )
                self._dbg(
                    u"инфо",
                    u"автофильтр источника «%s»: %s" % (src_name, note_src),
                )
        if first_sheet is not None:
            try:
                ctrl = doc.getCurrentController()
                if ctrl is not None:
                    ctrl.setActiveSheet(first_sheet)
                    try:
                        ctrl.FirstVisibleColumn = 0
                        ctrl.FirstVisibleRow = 0
                    except Exception:
                        pass
            except Exception:
                pass
            _split_select_a1(doc, first_sheet)
            self._dbg(u"инфо", u"активный лист «%s», курсор A1" % first_sheet.Name)

    def _save_and_close_book_entry(self, doc, path, book_key=None):
        """
        Внешняя книга:
        1) очистка листов (Hidden) → store → close
        2) reopen Hidden=False (полноценное окно)
        3) freeze / автофильтр / A1 → store → close
        """
        if doc is None:
            return False, u"doc=None"
        keep = self._cleanup_external_book_sheets(doc, path)
        ok, note = _split_save_and_close_doc(doc, path)
        self._dbg(u"инфо" if ok else u"ошибка", u"сохранение+закрытие (до reopen): %s" % note)
        if book_key is not None and book_key in self.open_books:
            self.open_books[book_key] = {"doc": None, "path": path, "closed": True}
        if doc is self.current_book_doc:
            self.current_book_doc = None
        if not ok:
            return False, note
        desktop = self.desktop
        if desktop is None:
            try:
                desktop = _desktop()
            except Exception:
                desktop = None
        vis_doc, rerr = _split_reopen_book_visible(desktop, path)
        if vis_doc is None:
            self._dbg(u"ошибка", u"reopen видимой книги: %s" % rerr)
            return True, note + u"; reopen: %s" % rerr
        self._dbg(u"инфо", u"книга открыта видимой для оформления: %s" % (path or u""))
        try:
            self._apply_visible_book_formatting(vis_doc, keep)
            self._reapply_paste_formats_after_formatting(vis_doc)
            ok2, note2 = _split_save_and_close_doc(vis_doc, path)
            self._dbg(
                u"инфо" if ok2 else u"ошибка",
                u"сохранение+закрытие (после оформления): %s" % note2,
            )
            if ok2:
                return True, note + u"; reopen+format; " + note2
            return False, note + u"; reopen+format; " + note2
        except Exception as err:
            try:
                _split_close_doc(vis_doc)
            except Exception:
                pass
            self._dbg(u"ошибка", u"оформление видимой книги: %s" % _split_err_text(err))
            return False, note + u"; format: %s" % _split_err_text(err)

    def _save_current_book(self):
        if self.this_book:
            return True, u""
        if not self.current_book_path:
            return False, u"нет пути книги"
        book_key = None
        for k, entry in self.open_books.items():
            if entry.get("doc") is self.current_book_doc:
                book_key = k
                break
        return self._save_and_close_book_entry(
            self.current_book_doc, self.current_book_path, book_key=book_key
        )

    def finish(self):
        if self.split_mode == u"copy_sheet":
            self._flush_block_copy_sheet()
        # эта_книга: оформление на уже видимой книге сбора
        if self.this_book:
            keep = list(self.book_created_sheets.get(id(self.result_doc), set()) or [])
            try:
                dropped = _split_clear_database_ranges(
                    self.result_doc, keep_anonymous=False
                )
                if dropped:
                    self._dbg(
                        u"инфо",
                        u"сняты DatabaseRanges: %s" % u", ".join(dropped),
                    )
            except Exception as err:
                self._dbg(u"инфо", u"очистка БД: %s" % _split_err_text(err))
            self._apply_visible_book_formatting(self.result_doc, keep)
            self._reapply_paste_formats_after_formatting(self.result_doc)
            try:
                _split_request_maximize_pending(self.result_doc)
                if _split_maximize_document(self.result_doc):
                    self._dbg(u"инфо", u"окно книги максимизировано")
                else:
                    self._dbg(u"инфо", u"максимизация: не удалось сразу (будет повтор в конце сбора)")
            except Exception as err:
                self._dbg(u"инфо", u"максимизация: %s" % _split_err_text(err))
                try:
                    _split_request_maximize_pending(self.result_doc)
                except Exception:
                    pass
        save_notes = []
        if self.save_immediately:
            if self.current_book_doc is not None and not self.this_book:
                ok, note = self._save_current_book()
                if note:
                    save_notes.append(note)
        elif not self.this_book:
            for book_key, entry in list(self.open_books.items()):
                path = entry.get("path")
                doc = entry.get("doc")
                if entry.get("closed"):
                    continue
                if path and doc is not None:
                    ok, note = self._save_and_close_book_entry(doc, path, book_key=book_key)
                    save_notes.append(note)
        summary_parts = [
            u"режим=%s" % self.split_mode,
            u"this_book=%s" % self.this_book,
            u"template=%s" % self.book_template,
            u"base_dir=%s" % (self.base_dir or u"(пусто)"),
            u"книг=%d" % len(self.open_books),
            u"листов=%d" % len(self.created_sheet_names),
        ]
        if self.created_book_paths:
            summary_parts.append(
                u"пути: " + u"; ".join(self.created_book_paths[:20])
            )
            for p in self.created_book_paths:
                summary_parts.append(u"%s → %s" % (p, _split_file_stat_note(p)))
        # контроль: незакрытые внешние книги
        still_open = []
        for book_key, entry in self.open_books.items():
            if entry.get("doc") is not None and not entry.get("closed"):
                still_open.append(entry.get("path") or book_key)
        if still_open:
            self._dbg(
                u"предупреждение",
                u"книги ещё открыты (блокировка файла?): %s" % u"; ".join(still_open),
            )
        else:
            if not self.this_book and self.open_books:
                self._dbg(u"инфо", u"все внешние книги сохранены и закрыты")
        self._dbg(u"инфо", u"итог: " + u" | ".join(summary_parts))
        return True, u"; ".join(save_notes) if save_notes else u""

    def process_rows(self, data_rows):
        prev_book_key = None
        prev_sheet_key = None
        field_names_all = list(self.split_books) + list(self.split_sheets)
        self._dbg(
            u"инфо",
            u"старт: строк=%d split_books=%s split_sheets=%s mode=%s paste_formats=%s book_template=%s sheet_template=%s base_dir=%s"
            % (
                len(data_rows or []),
                self.split_books,
                self.split_sheets,
                self.split_mode,
                self.paste_formats,
                self.book_template,
                self.sheet_template,
                self.base_dir or u"(пусто)",
            ),
        )
        for row in data_rows:
            self.current_row = int(row)
            field_values = _split_row_field_values(
                self.src_sheet, self.current_row, field_names_all, self.by_name
            )
            book_key = self._book_key(field_values)
            sheet_key = self._sheet_key(field_values)
            if book_key != prev_book_key:
                if self.split_mode == u"copy_sheet":
                    ok, err = self._flush_block_copy_sheet()
                    if not ok:
                        return False, err
                elif prev_sheet_key is not None:
                    self._flush_paste_formats_copy_range()
                if prev_book_key is not None and self.save_immediately:
                    self._save_current_book()
                ok, err = self._ensure_book(book_key, field_values)
                if not ok:
                    return False, err
                prev_book_key = book_key
                prev_sheet_key = None
            if sheet_key != prev_sheet_key:
                if self.split_mode == u"copy_sheet":
                    ok, err = self._flush_block_copy_sheet()
                    if not ok:
                        return False, err
                    self.current_block_rows = [int(row)]
                else:
                    if prev_sheet_key is not None:
                        self._flush_paste_formats_copy_range()
                    sheet_name = _split_resolve_sheet_name(self.sheet_template, field_values)
                    ok, err = self._prepare_sheet_copy_range(
                        self.current_book_doc, sheet_name
                    )
                    if not ok:
                        return False, err
                    ok, err = self._append_row_copy_range(row)
                    if not ok:
                        return False, err
                prev_sheet_key = sheet_key
            else:
                if self.split_mode == u"copy_sheet":
                    self.current_block_rows.append(int(row))
                else:
                    ok, err = self._append_row_copy_range(row)
                    if not ok:
                        return False, err
        if self.split_mode == u"copy_sheet":
            ok, err = self._flush_block_copy_sheet()
            if not ok:
                return False, err
        else:
            self._flush_paste_formats_copy_range()
        self.finish()
        return True, u""


def _split_pre_sort_sheet(doc, sheet, header_row_0, sc, ec, sr, er, sort_fields, by_name):
    if er < sr:
        return True, u""
    specs = []
    for name in sort_fields or []:
        col = by_name.get(_u(name).strip().casefold())
        if col is None:
            return False, u"сортировка: столбец «%s» не найден" % name
        specs.append((_u(name), True))
    if len(specs) == 0:
        return True, u""
    header_range = sheet.getCellRangeByPosition(int(sc), int(header_row_0), int(ec), int(header_row_0))
    data_range = sheet.getCellRangeByPosition(int(sc), int(sr), int(ec), int(er))
    resolved = []
    for col_token, ascending in [
        (_u(n), True) for n in sort_fields if _u(n).strip() != u""
    ]:
        col_idx = lm._lm_pp_resolve_column_token(
            col_token, sheet, header_range, int(sc), int(ec)
        )
        if col_idx is None:
            return False, u"сортировка: столбец «%s» не найден" % col_token
        resolved.append((int(col_idx), bool(ascending)))
    ok, err = lm._lm_pp_range_sort_dispatch(
        doc,
        sheet,
        int(sc),
        int(sr),
        int(ec),
        int(er),
        resolved,
        int(sc),
        sheet_name=str(getattr(sheet, "Name", "") or ""),
        fn_label=u"разделить_листы",
    )
    return (True, u"") if ok else (False, err or u"сортировка")


def lm_split_sheets_run(doc, spec, run_context=None):
    """
    Разделить лист по split_books / split_sheets.

  run_context:
        base_dir, log_fn(doc, sheet, phase, label, status, note),
        register_pending(name, header_row_0, source_sheet),
        append_sheets_info(doc), desktop
    """
    spec = dict(spec or {})
    phase = _u(spec.get("run_phase") or u"after_postprocess").strip().casefold()
    if phase not in _VALID_RUN_PHASES:
        phase = u"after_postprocess"
    label = _u((run_context or {}).get("label") or u"разделить_листы")
    if doc is None:
        return False, u"книга не задана", []
    src_info = spec.get("source") or {}
    src_sheet_name = _u(src_info.get("sheet") or u"").strip()
    if src_sheet_name == u"":
        _split_log(run_context, doc, u"", u"диапазон", label, u"ошибка", u"не задан source.sheet")
        return False, u"не задан source.sheet", []
    src_sheet = lm._lm_pp_get_sheet_ref(doc, src_sheet_name)
    if src_sheet is None:
        _split_log(
            run_context,
            doc,
            src_sheet_name,
            u"диапазон",
            label,
            u"ошибка",
            u"лист «%s» не найден" % src_sheet_name,
        )
        return False, u"лист не найден", []
    sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(src_sheet)
    if er < sr:
        _split_log(run_context, doc, src_sheet_name, u"диапазон", label, u"пропуск", u"нет данных")
        return False, u"нет данных", []
    header_row_0 = _split_resolve_header_row_0(src_sheet, spec.get("header_row"))
    if header_row_0 > er:
        header_row_0 = int(sr)
    data_start = max(int(header_row_0) + 1, int(sr))
    titles, by_name = _split_header_titles(src_sheet, header_row_0, sc, ec)
    book_fields = [] if _is_this_book_marker(spec.get("book_template")) else list(spec.get("split_books") or [])
    sheet_fields = list(spec.get("split_sheets") or [])
    _cols_b, miss_b = _split_resolve_field_columns(book_fields, by_name)
    _cols_s, miss_s = _split_resolve_field_columns(sheet_fields, by_name)
    missing = miss_b + miss_s
    if missing:
        note = u"столбцы не найдены: %s" % u", ".join(missing)
        _split_log(run_context, doc, src_sheet_name, u"диапазон", label, u"ошибка", note)
        return False, note, []
    sort_fields = list(book_fields) + list(sheet_fields)
    if bool(spec.get("pre_sort", True)) and len(sort_fields) > 0 and data_start <= er:
        ok_sort, err_sort = _split_pre_sort_sheet(
            doc, src_sheet, header_row_0, sc, ec, data_start, er, sort_fields, by_name
        )
        if not ok_sort:
            _split_log(run_context, doc, src_sheet_name, u"диапазон", label, u"ошибка", err_sort)
            return False, err_sort, []
        sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(src_sheet)
        data_start = max(int(header_row_0) + 1, int(sr))
    data_rows = []
    row = int(data_start)
    while row <= int(er):
        data_rows.append(int(row))
        row = row + 1
    if len(data_rows) == 0:
        _split_log(run_context, doc, src_sheet_name, u"диапазон", label, u"пропуск", u"нет строк данных")
        return False, u"нет строк данных", []
    base_dir = _u((run_context or {}).get("base_dir") or u"").strip()
    if base_dir == u"":
        try:
            url = doc.getURL()
            if url:
                base_dir = os.path.dirname(uno.fileUrlToSystemPath(url))
        except Exception:
            base_dir = u""
    desktop = (run_context or {}).get("desktop") or _desktop()
    needs_desktop = not _is_this_book_marker(spec.get("book_template"))
    if desktop is None and needs_desktop:
        note = u"нет Desktop UNO (не удалось создать/открыть книги)"
        _split_log(run_context, doc, src_sheet_name, u"диапазон", label, u"ошибка", note)
        return False, note, []
    writer = _SplitWriter(
        doc,
        src_sheet,
        spec,
        sc,
        ec,
        header_row_0,
        by_name,
        run_context,
        desktop,
        base_dir,
    )
    try:
        ok, note = writer.process_rows(data_rows)
    except Exception as err:
        ok = False
        note = u"%s" % err
    status = u"ok" if ok else u"ошибка"
    if ok:
        books_n = len(getattr(writer, "created_book_paths", []) or [])
        paths = list(getattr(writer, "created_book_paths", []) or [])
        if writer.this_book:
            note = u"лист «%s» → %d лист(ов) в этой книге, строк %d" % (
                src_sheet_name,
                len(writer.created_sheet_names),
                len(data_rows),
            )
        else:
            note = u"лист «%s» → книг %d, листов %d, строк %d" % (
                src_sheet_name,
                books_n,
                len(writer.created_sheet_names),
                len(data_rows),
            )
            if paths:
                note = note + u"; файлы: " + u"; ".join(
                    u"%s (%s)" % (p, _split_file_stat_note(p)) for p in paths[:10]
                )
    _split_log(run_context, doc, src_sheet_name, u"диапазон", label, status, note)
    if ok and writer.this_book:
        try:
            _split_request_maximize_pending(writer.result_doc or doc)
            _split_maximize_document(writer.result_doc or doc)
        except Exception:
            pass
    return ok, note, list(writer.created_sheet_names)
