# -*- coding: utf-8 -*-
"""
группировать_строки — сжатие таблицы по ключам с агрегатами (Group By).

Контракт JSON: todo/новая функция группировать_строки.md
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.728"
try:
    unicode
except NameError:
    unicode = str

FN_LABEL = u"группировать_строки"
OPS = frozenset((u"sum", u"count", u"min", u"max", u"avg", u"first", u"last"))


def _u(v):
    try:
        return unicode(v if v is not None else u"")
    except Exception:
        return str(v if v is not None else "")


def normalize_op(raw, default=u"sum"):
    s = _u(raw).strip().casefold()
    aliases = {
        u"sum": u"sum",
        u"сумма": u"sum",
        u"count": u"count",
        u"количество": u"count",
        u"кол-во": u"count",
        u"min": u"min",
        u"минимум": u"min",
        u"max": u"max",
        u"максимум": u"max",
        u"avg": u"avg",
        u"average": u"avg",
        u"среднее": u"avg",
        u"first": u"first",
        u"первый": u"first",
        u"last": u"last",
        u"последний": u"last",
    }
    return aliases.get(s, default if s not in OPS else s)


def parse_number(value):
    """Число или None (пусто / не распознано)."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            return float(value)
        except Exception:
            return None
    s = _u(value).strip()
    if s == u"":
        return None
    s = s.replace(u"\u00a0", u"").replace(u" ", u"")
    s = s.replace(u",", u".")
    try:
        return float(s)
    except Exception:
        return None


def is_empty_cell(value):
    if value is None:
        return True
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return False
    return _u(value).strip() == u""


def normalize_key_part(value, key_trim=True, key_case_sensitive=False):
    if value is None:
        s = u""
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            fv = float(value)
            if fv == int(fv):
                return (u"n", int(fv))
            return (u"n", fv)
        except Exception:
            s = _u(value)
    else:
        s = _u(value)
    if key_trim:
        s = s.strip()
    if (not key_case_sensitive) and s != u"":
        try:
            s = s.casefold()
        except Exception:
            s = s.lower()
    return (u"s", s)


def make_group_key(row, key_idxs, key_trim=True, key_case_sensitive=False):
    parts = []
    i = 0
    while i < len(key_idxs):
        ci = key_idxs[i]
        val = row[ci] if ci < len(row) else None
        parts.append(normalize_key_part(val, key_trim=key_trim, key_case_sensitive=key_case_sensitive))
        i = i + 1
    return tuple(parts)


def key_is_empty(gkey):
    """True, если хотя бы одна часть ключа пуста (после нормализации)."""
    i = 0
    while i < len(gkey):
        kind, val = gkey[i]
        if kind == u"n":
            i = i + 1
            continue
        if _u(val).strip() == u"":
            return True
        i = i + 1
    return False


def display_key_values(row, key_idxs):
    """Значения ключа для вывода (как в исходной строке-представителе)."""
    out = []
    i = 0
    while i < len(key_idxs):
        ci = key_idxs[i]
        out.append(row[ci] if ci < len(row) else u"")
        i = i + 1
    return out


def _cmp_minmax(a, b, as_num):
    if as_num:
        na = parse_number(a)
        nb = parse_number(b)
        if na is not None and nb is not None:
            return na < nb
        if na is not None and nb is None:
            return True
        if na is None and nb is not None:
            return False
    sa = _u(a)
    sb = _u(b)
    return sa < sb


def update_agg_state(state, op, value):
    """Обновить аккумулятор агрегата. state — dict."""
    if op == u"count":
        # value=None → COUNT(*); иначе непустые
        if value is None or (not is_empty_cell(value)):
            state["count"] = int(state.get("count") or 0) + 1
        return
    if op == u"sum":
        n = parse_number(value)
        if n is not None:
            state["sum"] = float(state.get("sum") or 0.0) + n
        return
    if op == u"avg":
        n = parse_number(value)
        if n is not None:
            state["sum"] = float(state.get("sum") or 0.0) + n
            state["count"] = int(state.get("count") or 0) + 1
        return
    if op == u"min":
        if is_empty_cell(value):
            return
        if "min" not in state:
            state["min"] = value
            state["min_num"] = parse_number(value) is not None
        else:
            prefer_num = state.get("min_num") and (parse_number(value) is not None)
            if _cmp_minmax(value, state["min"], prefer_num):
                state["min"] = value
                state["min_num"] = parse_number(value) is not None
        return
    if op == u"max":
        if is_empty_cell(value):
            return
        if "max" not in state:
            state["max"] = value
            state["max_num"] = parse_number(value) is not None
        else:
            prefer_num = state.get("max_num") and (parse_number(value) is not None)
            if _cmp_minmax(state["max"], value, prefer_num):
                state["max"] = value
                state["max_num"] = parse_number(value) is not None
        return
    if op == u"first":
        if "first" not in state:
            state["first"] = value
        return
    if op == u"last":
        state["last"] = value


def finalize_agg(state, op):
    if op == u"count":
        return int(state.get("count") or 0)
    if op == u"sum":
        return float(state.get("sum") or 0.0)
    if op == u"avg":
        c = int(state.get("count") or 0)
        if c <= 0:
            return u""
        return float(state.get("sum") or 0.0) / float(c)
    if op == u"min":
        return state.get("min", u"")
    if op == u"max":
        return state.get("max", u"")
    if op == u"first":
        return state.get("first", u"")
    if op == u"last":
        return state.get("last", u"")
    return u""


def normalize_aggregations(block):
    """
    Список {op, column, as} из block.
    Поддержка flat-полей agg_op / agg_column / agg_as.
    """
    out = []
    raw = block.get("aggregations")
    if isinstance(raw, (list, tuple)):
        for item in raw:
            if not isinstance(item, dict):
                continue
            op = normalize_op(item.get("op") or item.get("fn") or item.get("agg_fn"), default=u"sum")
            col = item.get("column")
            if col is None:
                col = item.get("columns")
            if isinstance(col, (list, tuple)):
                col = col[0] if len(col) > 0 else u""
            col = _u(col).strip()
            as_name = _u(item.get("as") or item.get("name") or item.get("alias") or u"").strip()
            if as_name == u"":
                if op == u"count" and col == u"":
                    as_name = u"Количество"
                elif col != u"":
                    as_name = u"%s_%s" % (op, col.strip(u"'"))
                else:
                    as_name = op
            out.append({u"op": op, u"column": col, u"as": as_name})
    if out:
        return out
    # flat wizard fields
    op = normalize_op(
        block.get("agg_op") or block.get("op") or block.get("agg_fn"),
        default=u"sum",
    )
    col = block.get("agg_column") or block.get("column") or u""
    if isinstance(col, (list, tuple)):
        col = col[0] if len(col) > 0 else u""
    col = _u(col).strip()
    as_name = _u(block.get("agg_as") or block.get("as") or u"").strip()
    if as_name == u"":
        if op == u"count" and col == u"":
            as_name = u"Количество"
        elif col != u"":
            as_name = u"%s_%s" % (op, col.strip(u"'"))
        else:
            as_name = op
    out.append({u"op": op, u"column": col, u"as": as_name})
    return out


def group_rows(body_rows, key_idxs, agg_specs, key_trim=True, key_case_sensitive=False, skip_empty_keys=True):
    """
    body_rows — список строк (без заголовка).
    agg_specs — [{op, col_idx|None, as}].
    Возвращает (out_rows, n_groups, n_skipped).
    """
    groups = {}  # key -> {rep_row, aggs: [state,...], order}
    order = []
    n_skipped = 0
    ri = 0
    while ri < len(body_rows):
        row = body_rows[ri]
        gkey = make_group_key(row, key_idxs, key_trim=key_trim, key_case_sensitive=key_case_sensitive)
        if skip_empty_keys and key_is_empty(gkey):
            n_skipped = n_skipped + 1
            ri = ri + 1
            continue
        if gkey not in groups:
            states = []
            ai = 0
            while ai < len(agg_specs):
                states.append({})
                ai = ai + 1
            groups[gkey] = {u"rep": list(row), u"states": states}
            order.append(gkey)
        entry = groups[gkey]
        ai = 0
        while ai < len(agg_specs):
            spec = agg_specs[ai]
            op = spec[u"op"]
            ci = spec.get(u"col_idx")
            if op == u"count" and ci is None:
                val = None
            elif ci is None:
                val = u""
            else:
                val = row[ci] if ci < len(row) else u""
            update_agg_state(entry[u"states"][ai], op, val)
            ai = ai + 1
        ri = ri + 1

    out_rows = []
    oi = 0
    while oi < len(order):
        gkey = order[oi]
        entry = groups[gkey]
        out = display_key_values(entry[u"rep"], key_idxs)
        ai = 0
        while ai < len(agg_specs):
            out.append(finalize_agg(entry[u"states"][ai], agg_specs[ai][u"op"]))
            ai = ai + 1
        out_rows.append(out)
        oi = oi + 1
    return out_rows, len(out_rows), n_skipped


def sort_grouped_rows(rows, n_key_cols):
    """Сортировка по ключам (первые n_key_cols колонок)."""
    def _sk(row):
        parts = []
        i = 0
        while i < n_key_cols and i < len(row):
            v = row[i]
            n = parse_number(v)
            if n is not None:
                parts.append((0, n))
            else:
                parts.append((1, _u(v).casefold() if hasattr(_u(v), "casefold") else _u(v).lower()))
            i = i + 1
        return parts

    try:
        return sorted(rows, key=_sk)
    except Exception:
        return rows

def lm_pp_range_group_by_rows(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Group By: сжать строки по key_columns + aggregations.

    JSON:
      [{"v":1,"fn":"группировать_строки","key_columns":["'Отдел'"],
        "aggregations":[{"op":"sum","column":"'Сумма'","as":"Сумма"}],
        "output":"inplace"}]
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

    fn_label = FN_LABEL
    sheet_name = sheet.Name if sheet is not None else u""
    if _lm_pp_param_skip_sheet(fn_label, extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(fn_label, extra_args, doc, sheet)
    if block is None:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нет блока JSON для листа"
        )
        return
    try:
        from libre_macros_param_codec import _normalize_group_by_rows_block
        block = _normalize_group_by_rows_block(block)
    except Exception:
        pass

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

    key_tokens = block.get(u"key_columns") or []
    if not isinstance(key_tokens, (list, tuple)):
        key_tokens = [key_tokens]
    key_abs = []
    seen_k = set()
    for tok in key_tokens:
        col = _lm_pp_resolve_column_token(tok, sheet, header_row_range, sc, ec)
        if col is None:
            continue
        if col < sc or col > ec:
            continue
        if col in seen_k:
            continue
        seen_k.add(col)
        key_abs.append(col)
    if len(key_abs) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"key_columns: столбцы не найдены (нужен минимум 1)",
        )
        return

    aggs_raw = normalize_aggregations(block)
    agg_specs = []
    for item in aggs_raw:
        op = item.get(u"op") or u"sum"
        col_tok = _u(item.get(u"column") or u"").strip()
        as_name = _u(item.get(u"as") or u"").strip() or op
        col_idx = None
        if op == u"count" and col_tok == u"":
            col_idx = None
        elif col_tok == u"":
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"aggregations: для op=%s нужен column" % op,
            )
            return
        else:
            abs_c = _lm_pp_resolve_column_token(col_tok, sheet, header_row_range, sc, ec)
            if abs_c is None or abs_c < sc or abs_c > ec:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    u"диапазон",
                    fn_label,
                    u"ошибка",
                    u"aggregations: столбец «%s» не найден" % col_tok,
                )
                return
            col_idx = abs_c - sc
        agg_specs.append({u"op": op, u"col_idx": col_idx, u"as": as_name})

    if len(agg_specs) == 0:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"ошибка", u"aggregations пуст"
        )
        return

    output = _u(block.get(u"output") or u"inplace").strip().casefold()
    if output not in (u"inplace", u"new_sheet"):
        output = u"inplace"
    sort_keys = bool(block.get(u"sort_keys"))
    key_trim = True if u"key_trim" not in block else bool(block.get(u"key_trim"))
    key_case_sensitive = bool(block.get(u"key_case_sensitive"))
    skip_empty_keys = True if u"skip_empty_keys" not in block else bool(
        block.get(u"skip_empty_keys")
    )
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

    n_cols = int(ec) - int(sc) + 1
    header = list(data[0])
    while len(header) < n_cols:
        header.append(u"")
    header = header[:n_cols]
    body = []
    for i in range(1, len(data)):
        r = list(data[i])
        while len(r) < n_cols:
            r.append(u"")
        body.append(r[:n_cols])

    key_rel = [c - sc for c in key_abs]
    out_header = []
    for ki in key_rel:
        if 0 <= ki < len(header):
            out_header.append(header[ki])
        else:
            out_header.append(u"")
    for spec in agg_specs:
        out_header.append(spec[u"as"])

    grouped, n_groups, n_skipped = group_rows(
        body,
        key_rel,
        agg_specs,
        key_trim=key_trim,
        key_case_sensitive=key_case_sensitive,
        skip_empty_keys=skip_empty_keys,
    )
    if sort_keys:
        grouped = sort_grouped_rows(grouped, len(key_rel))

    matrix = [out_header] + grouped
    in_r = len(body)
    out_r = len(grouped)
    out_c = len(out_header)

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
            u"строк %d → групп %d, столбцов %d на «%s»"
            % (in_r, n_groups, out_c, actual_dest)
        )
        if skip_empty_keys and n_skipped:
            note = note + u", пропущено пустых ключей %d" % n_skipped
        _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
        return

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

    note = u"строк %d → групп %d, столбцов %d, inplace" % (in_r, n_groups, out_c)
    if skip_empty_keys and n_skipped:
        note = note + u", пропущено пустых ключей %d" % n_skipped
    _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
