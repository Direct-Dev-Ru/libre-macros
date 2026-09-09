# -*- coding: utf-8 -*-
from __future__ import print_function
"""
Преобразование диапазонов и сводных (DataPilot) в статические значения.

Используется финальной обработкой (Только_значения) и постобработкой сводных (as_values).
"""
MACRO_VERSION = "3.10.717"
import re
import time

import libre_macros_lib as lm
from com.sun.star.table.CellHoriJustify import CENTER as HORI_CENTER
from com.sun.star.table.CellVertJustify import CENTER as VERT_CENTER


def lm_values_is_lo_default_sheet_name(name):
    """Автоимя листа Calc (SheetN, ЛистN, …)."""
    s = str(name or "").strip()
    if s == "":
        return False
    if re.match(r"^Sheet\d+$", s, re.IGNORECASE):
        return True
    if re.match(r"^Лист\d+$", s, re.IGNORECASE):
        return True
    if re.match(r"^Untitled\s*\d+$", s, re.IGNORECASE):
        return True
    if re.match(r"^Безымянный\s*\d+$", s, re.IGNORECASE):
        return True
    return False


def lm_values_purge_lo_default_sheets(doc, keep_names=None, log_sheet=""):
    """Удалить автолисты SheetN / ЛистN (кроме keep_names)."""
    if doc is None:
        return
    keep = set()
    if keep_names is not None:
        for nm in keep_names:
            s = str(nm or "").strip()
            if s != "":
                keep.add(s)
    to_remove = []
    try:
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            try:
                name = str(sheets.getByIndex(i).Name)
            except Exception:
                i = i + 1
                continue
            if lm_values_is_lo_default_sheet_name(name) and name not in keep:
                if name not in to_remove:
                    to_remove.append(name)
            i = i + 1
    except Exception:
        return
    topic = str(log_sheet or "").strip() or "только_значения"
    for name in to_remove:
        try:
            if doc.Sheets.hasByName(name):
                doc.Sheets.removeByName(name)
                lm._lm_log_postprocess(
                    doc,
                    name,
                    topic,
                    "purge_lo_default",
                    "успех",
                    "удалён автолист",
                )
        except Exception as err:
            lm._lm_log_postprocess(
                doc,
                name,
                topic,
                "purge_lo_default",
                "ошибка",
                "removeByName: %s" % err,
            )


def _lm_values_get_datapilot_tables(sheet):
    if sheet is None:
        return None
    try:
        return sheet.getDataPilotTables()
    except Exception:
        pass
    try:
        from com.sun.star.sheet import XDataPilotTablesSupplier

        supplier = sheet.queryInterface(XDataPilotTablesSupplier)
        if supplier is not None:
            return supplier.getDataPilotTables()
    except Exception:
        pass
    return None


def lm_values_sheet_has_datapilot(sheet):
    dp_tables = _lm_values_get_datapilot_tables(sheet)
    if dp_tables is None:
        return False
    try:
        names = dp_tables.getElementNames()
        return names is not None and len(names) > 0
    except Exception:
        return False


def _lm_values_datapilot_output_bounds(dp_tables):
    if dp_tables is None:
        return None
    try:
        names = dp_tables.getElementNames()
    except Exception:
        return None
    if names is None or len(names) == 0:
        return None
    sc = None
    sr = None
    ec = None
    er = None
    i = 0
    while i < len(names):
        nm = str(names[i])
        try:
            if not dp_tables.hasByName(nm):
                i = i + 1
                continue
            out = dp_tables.getByName(nm).getOutputRange()
            tsc = int(out.StartColumn)
            tsr = int(out.StartRow)
            tec = int(out.EndColumn)
            ter = int(out.EndRow)
            if ter < tsr or tec < tsc:
                i = i + 1
                continue
            if sc is None:
                sc, sr, ec, er = tsc, tsr, tec, ter
            else:
                if tsc < sc:
                    sc = tsc
                if tsr < sr:
                    sr = tsr
                if tec > ec:
                    ec = tec
                if ter > er:
                    er = ter
        except Exception:
            pass
        i = i + 1
    if sc is None:
        return None
    return int(sc), int(sr), int(ec), int(er)


def _lm_values_remove_datapilot_tables(sheet):
    removed = []
    errors = []
    dp_tables = _lm_values_get_datapilot_tables(sheet)
    if dp_tables is None:
        return removed, ["getDataPilotTables недоступен"]
    names = []
    try:
        for nm in dp_tables.getElementNames():
            s = str(nm).strip()
            if s != "":
                names.append(s)
    except Exception:
        pass
    i = len(names) - 1
    while i >= 0:
        nm = names[i]
        try:
            if dp_tables.hasByName(nm):
                dp_tables.removeByName(nm)
                removed.append(nm)
        except Exception as err:
            errors.append("%s: %s" % (nm, err))
        i = i - 1
    return removed, errors


def _lm_values_flatten_cell_range(sheet, sc, sr, ec, er):
    if sheet is None or er < sr or ec < sc:
        return False, "пустой диапазон"
    try:
        cell_range = sheet.getCellRangeByPosition(int(sc), int(sr), int(ec), int(er))
        data = cell_range.getDataArray()
        if not data or len(data) == 0:
            return False, "пустые данные"
        nrow = len(data)
        ncol = len(data[0]) if nrow > 0 else 0
        if ncol <= 0:
            return False, "пустые данные"
        dest_ec = int(sc) + ncol - 1
        dest_er = int(sr) + nrow - 1
        dest_range = sheet.getCellRangeByPosition(int(sc), int(sr), dest_ec, dest_er)
        dest_range.setDataArray(data)
        return True, "%d×%d" % (nrow, ncol)
    except Exception as err:
        return False, str(err)


def lm_values_pivot_inplace(doc, sheet):
    """
    Сводная на листе → статические значения на том же листе.
    getDataArray → remove DataPilot → calculateAll → setDataArray.
    """
    dp_tables = _lm_values_get_datapilot_tables(sheet)
    if dp_tables is None:
        return False, "DataPilot недоступен"
    bounds = _lm_values_datapilot_output_bounds(dp_tables)
    if bounds is None:
        return False, "пустой вывод сводной"
    sc, sr, ec, er = bounds
    try:
        usc, usr, uec, uer = lm._lm_vlookup_sheet_used_area(sheet)
        if uec >= usc and uer >= usr:
            if int(uec) > int(ec):
                ec = int(uec)
            if int(uer) > int(er):
                er = int(uer)
            if int(usc) < int(sc):
                sc = int(usc)
    except Exception:
        pass

    src_range = sheet.getCellRangeByPosition(sc, sr, ec, er)
    try:
        saved_data = src_range.getDataArray()
    except Exception as err:
        return False, "чтение: %s" % err
    if not saved_data or len(saved_data) == 0:
        return False, "пустой вывод"
    nrow = len(saved_data)
    ncol = len(saved_data[0]) if nrow > 0 else 0
    if ncol <= 0:
        return False, "пустой вывод"

    removed, dp_errors = _lm_values_remove_datapilot_tables(sheet)
    if len(removed) == 0 and len(dp_errors) > 0:
        return False, "не удалось удалить сводную: %s" % "; ".join(dp_errors)

    # У DataPilot/автофильтра могут оставаться DatabaseRanges (AF_*, AF_PP_*,
    # __Anonymous_Sheet_DB__*, смарт-таблица с TableStyle). Полностью чистим лист.
    if doc is not None and sheet is not None:
        try:
            rem = getattr(lm, "_lm_pp_remove_all_sheet_database_ranges", None)
            if callable(rem):
                rem(doc, sheet)
            else:
                lm._lm_pp_remove_sheet_autofilters(doc, sheet)
        except Exception:
            pass

    if doc is not None:
        try:
            doc.calculateAll()
        except Exception:
            pass

    dest_ec = int(sc) + int(ncol) - 1
    dest_er = int(sr) + int(nrow) - 1
    try:
        dest_range = sheet.getCellRangeByPosition(int(sc), int(sr), dest_ec, dest_er)
        dest_range.setDataArray(saved_data)
    except Exception as err:
        return False, "запись: %s" % err

    note = "сводная→значения %d×%d" % (nrow, ncol)
    if len(removed) > 0:
        note = note + "; DP: " + ", ".join(removed)
    if len(dp_errors) > 0:
        note = note + "; " + "; ".join(dp_errors)
    return True, note


def lm_values_flatten_used_area(sheet):
    """Вся значащая область листа → только значения."""
    sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(sheet)
    if er < sr or ec < sc:
        return False, "пустой лист"
    ok, size = _lm_values_flatten_cell_range(sheet, sc, sr, ec, er)
    if not ok:
        return False, size
    return True, "значения " + size


def _lm_final_header_plus_height_apply(doc, sheet, data_rng, header_rng, *extra_args):
    _unused = data_rng
    if sheet is None or header_rng is None:
        lm._lm_pp_header_height_dbg(
            "final_apply", "SKIP sheet=%s header_rng=%s" % (sheet, header_rng)
        )
        return
    lm._lm_pp_header_height_dbg(
        "final_apply",
        "sheet=%s extra=%r"
        % (
            getattr(sheet, "Name", sheet),
            extra_args[0] if extra_args else None,
        ),
    )
    cfg = lm._lm_pp_header_height_cfg_from_extra_or_block(doc, sheet, extra_args)
    if cfg is None:
        lm._lm_pp_header_height_dbg("final_apply", "SKIP cfg=None")
        return
    lm._lm_pp_header_height_dbg("final_apply", "cfg=%s" % cfg)
    lm._lm_pp_apply_header_row_height_pad(sheet, header_rng, cfg)


def lm_final_header_plus_height(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Финальная обработка: высота заголовка и выравнивание.

    C — param_decode (как у range «заголовок_плюс_высота»):
    «+13 center-center» (все листы), «лист | +8 center/bottom ; …».
    Пустая C — +12,7 мм, center/center на всех листах.
    """
    _lm_final_apply_range_blocks(
        doc,
        sheets,
        data_ranges,
        header_row_ranges,
        "заголовок_плюс_высота",
        _lm_final_header_plus_height_apply,
        label="заголовок_плюс_высота",
        require_data_range=False,
        extra_args=extra_args,
    )


def _lm_final_apply_grid_preset( doc, sheets, data_ranges, header_row_ranges, extra_args, apply_fn, label, sheet_filter_fn):
    """Общий цикл финальной сетки (тонкая/толстая) с фильтром по листам."""
    try:
        import libre_macros_final_lib as fl
    except ImportError:
        fl = None

    def _log(status, note=""):
        if fl is not None and doc is not None:
            fl._lm_final_log(doc, label, status, note)

    if sheets is None:
        _log("пропуск", "нет листов")
        return

    rest_text = str(extra_args[0]).strip() if extra_args and len(extra_args) > 0 else ""
    data = list(data_ranges or [])
    headers = list(header_row_ranges or [])
    done = []
    skipped = []
    i = 0
    while i < len(sheets):
        sheet = sheets[i]
        data_rng = data[i] if i < len(data) else None
        header_rng = headers[i] if i < len(headers) else None
        i = i + 1
        if sheet is None:
            continue
        try:
            nm = str(getattr(sheet, "Name", "")).strip()
        except Exception:
            nm = ""
        if not sheet_filter_fn(rest_text, nm, doc, sheet):
            continue
        if data_rng is None or header_rng is None:
            skipped.append(nm or "?")
            continue
        apply_fn(doc, sheet, data_rng, header_rng, rest_text)
        done.append(nm)

    if len(done) > 0:
        _log("ok", "%s: %s" % (label, ", ".join(done)))
    elif len(skipped) > 0:
        _log("пропуск", "нет диапазонов: %s" % ", ".join(skipped))
    else:
        _log("пропуск", "ни один лист не обработан")


def lm_final_grid(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Финальная обработка: «сетка» на листах книги (param_decode → lm_pp_range_grid_borders).
    """
    _lm_final_apply_range_blocks(
        doc,
        sheets,
        data_ranges,
        header_row_ranges,
        "сетка",
        lm.lm_pp_range_grid_borders,
        label="сетка",
        extra_args=extra_args,
    )


def lm_final_thin_grid(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """Финальная «тонкая_сетка» — param_decode, SheetListPanel."""
    _lm_final_apply_range_blocks(
        doc,
        sheets,
        data_ranges,
        header_row_ranges,
        "тонкая_сетка",
        lm.lm_pp_range_thin_grid_borders,
        label="тонкая_сетка",
        extra_args=extra_args,
    )


def lm_final_thick_grid(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """Финальная «толстая_сетка» — param_decode, SheetListPanel."""
    _lm_final_apply_range_blocks(
        doc,
        sheets,
        data_ranges,
        header_row_ranges,
        "толстая_сетка",
        lm.lm_pp_range_thick_grid_borders,
        label="толстая_сетка",
        extra_args=extra_args,
    )


def lm_final_vert_center(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Финальная обработка: вертикальное выравнивание по центру на строках данных.

    Строка заголовка (из header_row_ranges) не меняется — чтобы не затирать
    «заголовок_плюс_высота» (left-top и т.д.).
    """
    DEBUG_LOG = False  # включить подробный лог этой функции
    _DBG_SCOPE = "финал"
    _DBG_FN = "lm_final_vert_center"

    def _dbg(sheet, status, note):
        if not DEBUG_LOG:
            return
        try:
            sheet_name = ""
            try:
                sheet_name = str(getattr(sheet, "Name", "") or "")
            except Exception:
                sheet_name = ""
            lm._lm_log_postprocess(doc, sheet_name, _DBG_SCOPE, _DBG_FN, status, str(note))
        except Exception:
            try:
                print("[lm_final_vert_center] %s: %s" % (status, note))
            except Exception:
                pass

    if sheets is None:
        _dbg(None, "skip", "sheets is None")
        return

    dispatch_props = ()
    try:
        from com.sun.star.beans import PropertyValue  # type: ignore

        p = PropertyValue()
        p.Name = "VerticalAlignment"
        p.Value = 2  # CENTER (как в Macro Recorder)
        dispatch_props = (p,)
    except Exception as err:
        _dbg(None, "warn", "PropertyValue setup failed: %s" % err)

    headers = list(header_row_ranges or [])
    si = 0
    while si < len(sheets):
        sheet = sheets[si]
        hr_range = headers[si] if si < len(headers) else None
        si = si + 1
        if sheet is None:
            _dbg(None, "skip", "sheet is None")
            continue
        try:
            sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(sheet)
        except Exception as err:
            _dbg(sheet, "skip", "used_area failed: %s" % err)
            continue
        if er < sr or ec < sc:
            _dbg(sheet, "skip", "empty used area")
            continue

        data_sr = int(sr)
        if hr_range is not None:
            try:
                data_sr = int(hr_range.getRangeAddress().StartRow) + 1
            except Exception:
                data_sr = int(sr) + 1
        else:
            data_sr = int(sr) + 1
        if data_sr < int(sr):
            data_sr = int(sr)
        if int(er) < data_sr:
            _dbg(sheet, "skip", "no data rows below header")
            continue

        _dbg(
            sheet,
            "info",
            "data_area: (%s,%s)-(%s,%s) header skipped" % (sc, data_sr, ec, er),
        )

        try:
            cell_range = sheet.getCellRangeByPosition(
                int(sc), int(data_sr), int(ec), int(er)
            )
        except Exception as err:
            _dbg(sheet, "skip", "getCellRangeByPosition failed: %s" % err)
            continue

        try:
            cell_range.VertJustify = VERT_CENTER
            _dbg(sheet, "ok", "cell_range.VertJustify = CENTER")
        except Exception as err:
            _dbg(sheet, "warn", "VertJustify assign failed: %s" % err)

        if doc is None or len(dispatch_props) == 0:
            continue
        if not lm._lm_pp_select_cell_range(doc, sheet, cell_range):
            _dbg(sheet, "warn", "select range failed")
            continue
        ok_disp = lm._lm_pp_execute_dispatch(doc, ".uno:VerticalAlignment", dispatch_props)
        if ok_disp:
            _dbg(sheet, "ok", "dispatch .uno:VerticalAlignment ok")
        else:
            _dbg(sheet, "warn", "dispatch .uno:VerticalAlignment failed")


def _lm_final_parse_sheet_names_from_spec(sheet_spec):
    return lm._lm_pp_parse_sheet_names_from_spec(sheet_spec)


def _lm_final_sheet_name_matches_target(sheet_name, target_names):
    return lm._lm_pp_sheet_name_matches_target(sheet_name, target_names)


def _lm_final_parse_sort_sheet_assignments(text):
    return lm._lm_pp_parse_sort_sheet_assignments(text)


def lm_final_zebra(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """Финальная «зебра_диапазон» — param_decode → lm_pp_range_zebra_even_rows."""
    _lm_final_apply_range_blocks(
        doc,
        sheets,
        data_ranges,
        header_row_ranges,
        "зебра_диапазон",
        lm.lm_pp_range_zebra_even_rows,
        label="зебра_диапазон",
        extra_args=extra_args,
    )


def lm_final_sort(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """Финальная «сортировка» — param_decode → lm_pp_range_sort_data."""
    _lm_final_apply_range_blocks(
        doc,
        sheets,
        data_ranges,
        header_row_ranges,
        "сортировка",
        lm.lm_pp_range_sort_data,
        label="сортировка",
        extra_args=extra_args,
    )


def lm_final_colorize_blocks(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """Финальная «раскрасить_блоки» — param_decode → lm_pp_range_colorize_data."""
    _lm_final_apply_range_blocks(
        doc,
        sheets,
        data_ranges,
        header_row_ranges,
        "раскрасить_блоки",
        lm.lm_pp_range_colorize_data,
        label="раскрасить_блоки",
        extra_args=extra_args,
    )


lm_final_colorize = lm_final_colorize_blocks


# =============================================================================
# Финал §3.5 B: param_decode → lm_pp_range_* через общий цикл блоков
# =============================================================================


def _lm_final_log_fn(label):
    try:
        import libre_macros_final_lib as fl
    except ImportError:
        fl = None

    def _log(doc, status, note=""):
        if fl is not None and doc is not None:
            fl._lm_final_log(doc, label, status, note)

    return _log


def _lm_final_blocks_for_sheet(fn_key, blocks, sheet_name):
    """Блоки, применимые к листу: конкретные листы приоритетнее «все листы»."""
    try:
        from libre_macros_param_codec import is_sheet_list_fn, normalize_fn_key
    except ImportError:
        return list(blocks or [])
    fn_key = normalize_fn_key(fn_key)
    blocks = blocks or []
    specific = []
    global_blocks = []
    for block in blocks:
        if not isinstance(block, dict):
            continue
        if is_sheet_list_fn(fn_key):
            sheets = block.get("sheets")
            if sheets is None or (
                isinstance(sheets, (list, tuple)) and len(sheets) == 0
            ):
                global_blocks.append(block)
            elif _lm_final_sheet_name_matches_target(sheet_name, sheets):
                specific.append(block)
        else:
            sn = str(block.get("sheet") or "").strip()
            if sn == "":
                global_blocks.append(block)
            elif _lm_final_sheet_name_matches_target(sheet_name, [sn]):
                specific.append(block)
    if len(specific) > 0:
        return specific
    return global_blocks


def _lm_final_block_to_extra_args(fn_key, block):
    try:
        from libre_macros_param_codec import is_sheet_list_fn, normalize_fn_key, param_encode
    except ImportError:
        return []
    fn_key = normalize_fn_key(fn_key)
    if is_sheet_list_fn(fn_key):
        text = param_encode(fn_key, [block])
    else:
        text = param_encode(fn_key, [block])
    text = str(text or "").strip()
    return [text] if text else []


def _lm_final_ranges_from_used(sheet, header_row=0):
    """
    Актуальные data/header диапазоны по used area листа.

    Нужно после шагов вроде «применить_формулу», которые добавляют столбцы
    правее исходного data_range, переданного в финальную фазу.
    """
    if sheet is None:
        return None, None
    try:
        header_row = int(header_row)
    except (TypeError, ValueError):
        header_row = 0
    if header_row < 0:
        header_row = 0
    try:
        end_col, end_row = lm.get_sheet_used_bounds(sheet)
    except Exception:
        end_col, end_row = 0, 0
    if end_col < 0:
        end_col = 0
    data_start = int(header_row) + 1
    try:
        header_rng = sheet.getCellRangeByPosition(0, header_row, end_col, header_row)
        if end_row < data_start:
            data_rng = sheet.getCellRangeByPosition(0, data_start, end_col, data_start)
        else:
            data_rng = sheet.getCellRangeByPosition(0, data_start, end_col, end_row)
        return data_rng, header_rng
    except Exception:
        return None, None


def _lm_final_resolve_header_row_0(doc, header_rng):
    """
    0-based строка заголовка для пересборки used-area в финале.

    «Текущие листы» — из «Строка_Заголовков» (пусто → 0).
    Иначе — StartRow уже переданного header_rng (обычно 0 после сбора).
    """
    try:
        import libre_macros_collect_cfg as _cw_cfg

        if getattr(_cw_cfg, "_MERGE_COLLECT_MODE", None) == getattr(
            _cw_cfg, "MERGE_MODE_INVOLVE", None
        ):
            try:
                import collect_workbooks as _cw

                return int(_cw._merge_involve_header_row_0(doc))
            except Exception:
                pass
    except Exception:
        pass
    try:
        if header_rng is not None:
            return max(0, int(header_rng.getRangeAddress().StartRow))
    except Exception:
        pass
    return 0


def _lm_final_apply_range_blocks( doc, sheets, data_ranges, header_row_ranges, fn_key, range_fn, label=None, uses_extra=True, require_data_range=True, extra_args=()):
    """
    Общий цикл финала §3.5 B: param_decode(C) → блоки → lm_pp_range_* по листам.
    """
    label = label or fn_key
    _log = _lm_final_log_fn(label)
    prev_stage = lm.lm_pp_active_log_stage()
    lm.lm_pp_set_log_stage(u"финальная_обработка")
    t_all = time.time()
    prev_ctx = None
    try:
        if sheets is None:
            _log(doc, "пропуск", "нет листов")
            return
        try:
            from libre_macros_param_codec import normalize_fn_key, param_decode
        except ImportError:
            _log(doc, "пропуск", "нет param_codec")
            return
        fn_key = normalize_fn_key(fn_key)
        raw = str(extra_args[0] if extra_args and len(extra_args) > 0 else "").strip()
        try:
            blocks = param_decode(fn_key, raw)
        except Exception:
            blocks = []
        if len(blocks) == 0 and raw == "":
            blocks = [{"v": 1, "fn": fn_key}]

        data = list(data_ranges or [])
        headers = list(header_row_ranges or [])

        if fn_key == u"разделить_по_столбцам":
            try:
                sheet_names = []
                si = 0
                while si < len(sheets or ()):
                    sh = sheets[si]
                    nm = ""
                    try:
                        nm = str(getattr(sh, "Name", "") or "").strip()
                    except Exception:
                        nm = ""
                    sheet_names.append(nm)
                    si = si + 1
                _log(
                    doc,
                    "инфо",
                    "debug %s: sheets=%d %r data_ranges=%d header_ranges=%d blocks=%d"
                    % (
                        fn_key,
                        len(sheet_names),
                        sheet_names,
                        len(data),
                        len(headers),
                        len(blocks),
                    ),
                )
            except Exception:
                pass

        done = []
        skipped = []
        i = 0
        prev_ctx = lm.lm_pp_active_context()
        while i < len(sheets):
            sheet = sheets[i]
            data_rng = data[i] if i < len(data) else None
            header_rng = headers[i] if i < len(headers) else None
            i = i + 1
            if sheet is None:
                continue
            try:
                nm = str(getattr(sheet, "Name", "")).strip()
            except Exception:
                nm = ""
            sheet_blocks = _lm_final_blocks_for_sheet(fn_key, blocks, nm)
            if len(sheet_blocks) == 0:
                continue
            # Пересобрать диапазоны по used area — иначе столбцы после
            # «применить_формулу» остаются вне ec и формат_столбцы их не находит.
            # Важно: не затирать Строка_Заголовков режима «Текущие листы» нулём.
            if require_data_range:
                hdr0 = _lm_final_resolve_header_row_0(doc, header_rng)
                fresh_data, fresh_hdr = _lm_final_ranges_from_used(
                    sheet, header_row=hdr0
                )
                if fresh_data is not None and fresh_hdr is not None:
                    data_rng, header_rng = fresh_data, fresh_hdr
            if require_data_range and (data_rng is None or header_rng is None):
                if nm and nm not in skipped:
                    skipped.append(nm)
                continue
            ctx = {
                "sheet": sheet,
                "header_row": 0,
                "data_start": 1,
                "end_col": 0,
                "end_row": 0,
            }
            try:
                if header_rng is not None:
                    hdr_addr = header_rng.getRangeAddress()
                    ctx["header_row"] = int(hdr_addr.StartRow)
                    ctx["data_start"] = int(hdr_addr.StartRow) + 1
                    ctx["end_col"] = int(hdr_addr.EndColumn)
                    ctx["end_row"] = int(hdr_addr.EndRow)
                if data_rng is not None:
                    data_addr = data_rng.getRangeAddress()
                    ctx["end_col"] = max(int(ctx.get("end_col", 0)), int(data_addr.EndColumn))
                    ctx["end_row"] = max(int(ctx.get("end_row", 0)), int(data_addr.EndRow))
                    ctx["data_start"] = int(data_addr.StartRow)
            except Exception:
                pass
            try:
                import libre_macros_collect_cfg as _cw_cfg

                ctx["source_variables_map"] = getattr(_cw_cfg, "_MERGE_SOURCE_VARIABLES_MAP", {}) or {}
                ctx["source_variables_order"] = getattr(_cw_cfg, "_MERGE_SOURCE_VARIABLES_ORDER", []) or []
            except Exception:
                if prev_ctx is not None:
                    ctx["source_variables_map"] = prev_ctx.get("source_variables_map") or {}
                    ctx["source_variables_order"] = prev_ctx.get("source_variables_order") or []
            lm.lm_pp_set_active_context(ctx)
            for block in sheet_blocks:
                if fn_key == u"заголовок_плюс_высота":
                    if require_data_range and (data_rng is None or header_rng is None):
                        if nm and nm not in skipped:
                            skipped.append(nm)
                        continue
                    lm.lm_pp_log_fn_begin()
                    try:
                        range_fn(doc, sheet, data_rng, header_rng, block)
                    finally:
                        lm.lm_pp_log_fn_end()
                    continue
                if fn_key == u"применить_формулу":
                    if require_data_range and (data_rng is None or header_rng is None):
                        if nm and nm not in skipped:
                            skipped.append(nm)
                        continue
                    lm.lm_pp_log_fn_begin()
                    try:
                        range_fn(doc, sheet, data_rng, header_rng, block)
                    finally:
                        lm.lm_pp_log_fn_end()
                    continue
                lm.lm_pp_log_fn_begin()
                try:
                    if uses_extra:
                        ea = _lm_final_block_to_extra_args(fn_key, block)
                        if len(ea) > 0:
                            range_fn(doc, sheet, data_rng, header_rng, *ea)
                        else:
                            range_fn(doc, sheet, data_rng, header_rng)
                    else:
                        range_fn(doc, sheet, data_rng, header_rng)
                finally:
                    lm.lm_pp_log_fn_end()
            if nm and nm not in done:
                done.append(nm)
        lm.lm_pp_set_active_context(prev_ctx)

        elapsed_all = time.time() - t_all
        if len(done) > 0:
            _log(
                doc,
                "ok",
                "%s: %s, %.2f с" % (label, ", ".join(done), elapsed_all),
            )
        elif len(skipped) > 0:
            _log(
                doc,
                "пропуск",
                "нет диапазонов: %s, %.2f с" % (", ".join(skipped), elapsed_all),
            )
        else:
            _log(doc, "пропуск", "ни один лист не обработан, %.2f с" % elapsed_all)
    finally:
        try:
            lm.lm_pp_set_active_context(prev_ctx)
        except Exception:
            pass
        lm.lm_pp_set_log_stage(prev_stage)


def _lm_final_make_range( fn_key, py_name, range_attr, uses_extra=True, require_data_range=True):
    def _fn(doc, sheets, data_ranges, header_row_ranges, *extra_args):
        range_fn = getattr(lm, range_attr)
        _lm_final_apply_range_blocks(
            doc,
            sheets,
            data_ranges,
            header_row_ranges,
            fn_key,
            range_fn,
            label=fn_key,
            uses_extra=uses_extra,
            require_data_range=require_data_range,
            extra_args=extra_args,
        )

    _fn.__name__ = "lm_final_" + py_name
    _fn.__doc__ = (
        u"Финал: «%s» — param_decode колонки C, затем %s по листам книги."
        % (fn_key, range_attr)
    )
    return _fn


_LM_FINAL_RANGE_SPEC = (
    ("перенос", "word_wrap", "lm_pp_range_word_wrap"),
    (
        "перенос_и_авто_высота",
        "word_wrap_and_fit",
        "lm_pp_range_word_wrap_and_fit_rows",
    ),
    ("авто_высота", "autofit_rows", "lm_pp_range_autofit_row_heights", False),
    (
        "авто_ширина",
        "autofit_columns",
        "lm_pp_range_autofit_columns_width",
        False,
    ),
    ("левое_выравнивание", "left_align", "lm_pp_range_left_align_data", False),
    (
        "сброс_стрипов",
        "reset_path_stripes",
        "lm_pp_range_init_path_stripe_state",
        False,
        False,
    ),
    ("закрепить_заголовок", "freeze_header", "lm_pp_range_freeze_header", False),
    ("автофильтр", "autofilter", "lm_pp_range_add_filter"),
    ("высота_строки", "row_height", "lm_pp_range_row_height_pad"),
    ("отступ", "indent", "lm_pp_range_increase_indent"),
    ("шрифт", "set_font", "lm_pp_range_set_font"),
    ("ширина_столбцов", "column_width", "lm_pp_range_set_column_widths"),
    ("формат_деньги", "format_money", "lm_pp_range_format_money_columns"),
    ("формат_даты", "format_date", "lm_pp_range_format_date_columns"),
    ("формат_столбцы", "format_columns", "lm_pp_range_format_columns"),
    (
        "подсветка_по_порогу",
        "highlight_threshold",
        "lm_pp_range_highlight_by_threshold",
    ),
    (
        "подкрасить_пороги",
        "highlight_threshold",
        "lm_pp_range_highlight_by_threshold",
    ),
    ("градиент", "color_scale", "lm_pp_range_color_scale_simple"),
    ("удалить_столбцы", "delete_columns", "lm_pp_range_delete_columns"),
    ("конкатенация_столбцов", "concat_columns", "lm_pp_range_concat_columns"),
    ("разделить_по_столбцам", "split_by_columns", "lm_pp_range_split_by_columns"),
    ("условный_столбец", "conditional_column", "lm_pp_range_conditional_column"),
    ("столбец_по_лямбде", "lambda_column", "lm_pp_range_lambda_column"),
    ("добавить_столбец", "add_column", "lm_pp_range_add_column"),
    ("применить_формулу", "apply_formula", "lm_pp_range_apply_formula"),
    ("удаление_строк", "delete_rows", "lm_pp_range_skip_empty_rows"),
    ("группировка_по_столбцу", "group_by_column", "lm_pp_range_group_by_column"),
    (
        "стиль_печати",
        "print_style",
        "lm_pp_range_set_page_style",
        True,
        False,
    ),
    ("заполнение_вниз", "fill_down", "lm_pp_range_fill_down_empty"),
    ("заполнить_вверх", "fill_up", "lm_pp_range_fill_up_empty"),
    ("развернуть_столбцы", "unpivot_columns", "lm_pp_range_unpivot_columns"),
    ("транспонировать_таблицу", "transpose_table", "lm_pp_range_transpose_table"),
    ("анкета_в_таблицу", "form_to_table", "lm_pp_range_form_to_table"),
    ("таблица_в_анкету", "table_to_form", "lm_pp_range_table_to_form"),
    (
        "заполнение_вниз_вычислить",
        "fill_down_calculate",
        "lm_pp_range_fill_down_calculate",
    ),
    ("копировать_значения", "copy_values", "lm_pp_range_copy_values"),
    ("замена_значений", "replace_values", "lm_pp_range_replace_values"),
    ("текстовые_операции", "normalize_text", "lm_pp_range_normalize_text"),
    (
        "переименовать_лист",
        "rename_sheet",
        "lm_pp_range_rename_sheet",
        True,
        False,
    ),
    (
        "переименовать_столбцы",
        "rename_columns",
        "lm_pp_range_rename_columns",
    ),
    (
        "переставить_столбцы",
        "reorder_columns",
        "lm_pp_range_reorder_columns",
    ),
    ("количество_значений", "value_count", "lm_pp_range_value_count"),
    ("удалить_дубликаты", "remove_duplicates", "lm_pp_range_remove_duplicates"),
    ("группировать_строки", "group_by_rows", "lm_pp_range_group_by_rows"),
)

for _spec in _LM_FINAL_RANGE_SPEC:
    _fn_key = _spec[0]
    _py_name = _spec[1]
    _range_attr = _spec[2]
    _uses_extra = _spec[3] if len(_spec) > 3 else True
    _require_dr = _spec[4] if len(_spec) > 4 else True
    globals()["lm_final_" + _py_name] = _lm_final_make_range(
        _fn_key,
        _py_name,
        _range_attr,
        uses_extra=_uses_extra,
        require_data_range=_require_dr,
    )
