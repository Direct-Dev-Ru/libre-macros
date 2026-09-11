# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.724"
"""
развернуть_столбцы (unpivot): широкая таблица → длинная.

Карта: lm_pp_range_unpivot_columns / merge_final_unpivot_columns.
"""

try:
    unicode
except NameError:
    unicode = str


def _u(v):
    return unicode(v if v is not None else u"")


def unpivot_value_is_empty(val):
    if val is None:
        return True
    try:
        if isinstance(val, float) and val != val:  # NaN
            return True
    except Exception:
        pass
    s = _u(val).strip()
    return s == u""


def build_unpivot_matrix( header_row, body_rows, unpivot_rel_indices, keep_rel_indices, attribute_column, value_column, drop_empty_rows=False):
    """
    Построить матрицу [header] + rows.

    header_row / body_rows — списки ячеек одной ширины.
    *_rel_indices — индексы относительно header_row (0-based).
    Возвращает (matrix, dropped_count).
    """
    hdr = list(header_row or [])
    up = [int(i) for i in (unpivot_rel_indices or [])]
    keep = [int(i) for i in (keep_rel_indices or [])]
    attr_name = _u(attribute_column).strip() or u"Атрибут"
    val_name = _u(value_column).strip() or u"Значение"

    out_header = []
    for ki in keep:
        if 0 <= ki < len(hdr):
            out_header.append(hdr[ki])
        else:
            out_header.append(u"")
    out_header.append(attr_name)
    out_header.append(val_name)

    out_rows = []
    dropped = 0
    for row in body_rows or ():
        row = list(row)
        keep_vals = []
        for ki in keep:
            if 0 <= ki < len(row):
                keep_vals.append(row[ki])
            else:
                keep_vals.append(u"")
        for ui in up:
            if 0 <= ui < len(hdr):
                attr = hdr[ui]
            else:
                attr = u""
            if 0 <= ui < len(row):
                val = row[ui]
            else:
                val = u""
            if drop_empty_rows and unpivot_value_is_empty(val):
                dropped += 1
                continue
            out_rows.append(list(keep_vals) + [attr, val])

    return [out_header] + out_rows, dropped


def lm_pp_range_unpivot_columns(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Unpivot Columns (Power Query): выбранные столбцы → attribute + value строки.

    JSON:
      [{"v":1,"fn":"развернуть_столбцы","unpivot_columns":["'Январь'"],
        "attribute_column":"Месяц","value_column":"Сумма","output":"inplace"}]
    """
    from libre_macros_lib import (
        _lm_log_postprocess,
        _lm_pp_calculate_with_auto,
        _lm_pp_clear_sheet_content,
        _lm_pp_copy_range_as_values,
        _lm_pp_get_sheet_ref,
        _lm_pp_param_block_for_sheet,
        _lm_pp_param_skip_sheet,
        _lm_pp_parse_as_values_flag,
        _lm_pp_range_address,
        _lm_pp_resolve_column_token,
        _lm_pp_unique_sheet_name,
        lm_pp_copy_sheet_register_pending,
    )

    fn_label = u"развернуть_столбцы"
    sheet_name = sheet.Name if sheet is not None else u""
    if _lm_pp_param_skip_sheet(fn_label, extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(fn_label, extra_args, doc, sheet)
    if block is None:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нет блока JSON для листа"
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    if sr <= h_sr:
        sr = h_sr + 1
    if er < sr:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нет строк данных"
        )
        return
    if ec < sc:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нет столбцов"
        )
        return

    attr_col = _u(block.get(u"attribute_column") or u"Атрибут").strip() or u"Атрибут"
    val_col = _u(block.get(u"value_column") or u"Значение").strip() or u"Значение"
    output = _u(block.get(u"output") or u"inplace").strip().casefold()
    if output not in (u"inplace", u"new_sheet"):
        output = u"inplace"
    drop_empty = bool(block.get(u"drop_empty_rows"))
    as_values = True
    if u"as_values" in block:
        as_values = bool(_lm_pp_parse_as_values_flag(block))

    if output == u"new_sheet":
        dest_sheet_name = _u(
            block.get(u"dest_sheet") or block.get(u"dest") or u""
        ).strip()
        if dest_sheet_name == u"":
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"output=new_sheet: нужен dest_sheet",
            )
            return

    up_tokens = block.get(u"unpivot_columns") or []
    if not isinstance(up_tokens, (list, tuple)):
        up_tokens = [up_tokens]
    excl_tokens = block.get(u"exclude_columns") or []
    if not isinstance(excl_tokens, (list, tuple)):
        excl_tokens = [excl_tokens]

    up_abs = []
    seen_up = set()
    for tok in up_tokens:
        col = _lm_pp_resolve_column_token(tok, sheet, header_row_range, sc, ec)
        if col is None:
            continue
        if col < sc or col > ec:
            continue
        if col in seen_up:
            continue
        seen_up.add(col)
        up_abs.append(col)

    excl_abs = set()
    for tok in excl_tokens:
        col = _lm_pp_resolve_column_token(tok, sheet, header_row_range, sc, ec)
        if col is None:
            continue
        if sc <= col <= ec:
            excl_abs.add(col)

    if len(up_abs) == 0:
        # Режим «все кроме exclude»
        c = sc
        while c <= ec:
            if c not in excl_abs:
                up_abs.append(c)
            c += 1

    if len(up_abs) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"нет столбцов для разворота (unpivot_columns / exclude)",
        )
        return

    keep_abs = []
    up_set = set(up_abs)
    c = sc
    while c <= ec:
        if c not in up_set:
            keep_abs.append(c)
        c += 1

    if as_values and doc is not None and sheet is not None:
        try:
            mat_rng = sheet.getCellRangeByPosition(int(sc), int(h_sr), int(ec), int(er))
            _lm_pp_calculate_with_auto(doc, sheet)
            _lm_pp_copy_range_as_values(doc, sheet, mat_rng)
        except Exception:
            pass

    try:
        full_rng = sheet.getCellRangeByPosition(int(sc), int(h_sr), int(ec), int(er))
        data = full_rng.getDataArray()
    except Exception as err:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"getDataArray: %s" % err,
        )
        return

    if not data:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"пустой диапазон"
        )
        return

    header = list(data[0])
    body = [list(data[i]) for i in range(1, len(data))]
    n_cols = int(ec) - int(sc) + 1
    # нормализовать ширину строк
    while len(header) < n_cols:
        header.append(u"")
    header = header[:n_cols]
    fixed_body = []
    for row in body:
        r = list(row)
        while len(r) < n_cols:
            r.append(u"")
        fixed_body.append(r[:n_cols])

    up_rel = [c - sc for c in up_abs]
    keep_rel = [c - sc for c in keep_abs]

    matrix, dropped = build_unpivot_matrix(
        header,
        fixed_body,
        up_rel,
        keep_rel,
        attr_col,
        val_col,
        drop_empty_rows=drop_empty,
    )

    in_r = len(fixed_body)
    in_c = len(up_abs)
    out_r = max(0, len(matrix) - 1)
    out_c = len(matrix[0]) if matrix else 0

    if output == u"new_sheet":
        dest_base = _u(block.get(u"dest_sheet") or block.get(u"dest") or u"").strip()
        actual_dest = _lm_pp_unique_sheet_name(doc, dest_base, exclude_sheet=None)
        dest_sh = _lm_pp_get_sheet_ref(doc, actual_dest)
        if dest_sh is None:
            try:
                pos = int(doc.Sheets.getCount())
                doc.Sheets.insertNewByName(actual_dest, pos)
            except Exception as err:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    u"диапазон",
                    fn_label,
                    u"ошибка",
                    u"создание листа «%s»: %s" % (actual_dest, err),
                )
                return
            dest_sh = _lm_pp_get_sheet_ref(doc, actual_dest)
        else:
            _lm_pp_clear_sheet_content(dest_sh)
        try:
            from libre_macros_lib import _lm_pp_remove_duplicates_write_matrix

            ok, werr = _lm_pp_remove_duplicates_write_matrix(doc, dest_sh, 0, 0, matrix)
        except Exception as err:
            ok, werr = False, _u(err)
        if not ok:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                werr or u"запись на лист",
            )
            return
        try:
            lm_pp_copy_sheet_register_pending(
                actual_dest, 0, sheet_name, from_dedup=False
            )
        except Exception:
            pass
        note = (
            u"вход %d×%d unpivot → строк %d, столбцов %d на «%s»"
            % (in_r, in_c, out_r, out_c, actual_dest)
        )
        if drop_empty and dropped:
            note = note + u", отброшено пустых %d" % dropped
        _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
        return

    # inplace: очистить старый прямоугольник и записать новую матрицу
    try:
        old_rng = sheet.getCellRangeByPosition(int(sc), int(h_sr), int(ec), int(er))
        old_rng.clearContents(7)
    except Exception:
        pass
    try:
        from libre_macros_lib import _lm_pp_remove_duplicates_write_matrix

        ok, werr = _lm_pp_remove_duplicates_write_matrix(
            doc, sheet, int(sc), int(h_sr), matrix
        )
    except Exception as err:
        ok, werr = False, _u(err)
    if not ok:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            werr or u"запись inplace",
        )
        return

    # если новая матрица уже/уже по столбцам — подчистить хвост справа
    new_ec = int(sc) + out_c - 1
    new_er = int(h_sr) + len(matrix) - 1
    if new_ec < int(ec):
        try:
            tail = sheet.getCellRangeByPosition(
                new_ec + 1, int(h_sr), int(ec), max(int(er), new_er)
            )
            tail.clearContents(7)
        except Exception:
            pass
    if new_er < int(er):
        try:
            below = sheet.getCellRangeByPosition(
                int(sc), new_er + 1, max(int(ec), new_ec), int(er)
            )
            below.clearContents(7)
        except Exception:
            pass

    try:
        from libre_macros_lib import lm_pp_active_context, _lm_pp_context_refresh

        ctx = lm_pp_active_context()
        if ctx is not None:
            _lm_pp_context_refresh(ctx)
    except Exception:
        pass

    note = (
        u"вход %d строк × %d unpivot → %d строк × %d столбцов, inplace"
        % (in_r, in_c, out_r, out_c)
    )
    if drop_empty and dropped:
        note = note + u", отброшено пустых %d" % dropped
    _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
