# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.703"
"""
анкета_в_таблицу / таблица_в_анкету — вертикальные пары Q/A ↔ wide-таблица.

Карта:
  lm_pp_range_form_to_table / merge_final_form_to_table
  lm_pp_range_table_to_form / merge_final_table_to_form
"""

try:
    unicode
except NameError:
    unicode = str

_MAX_OUT_PAIRS = 1000000
_MAX_WIDE_COLS = 500
_BLOCK_INDEX_NAMES = (u"#блок", u"#Блок", u"block_index", u"блок")
_SOURCE_ROW_NAMES = (
    u"source_start_row",
    u"source_row",
    u"строка_источника",
    u"#строка",
)


def _u(v):
    return unicode(v if v is not None else u"")


def _empty(val):
    if val is None:
        return True
    try:
        if isinstance(val, float) and val != val:
            return True
    except Exception:
        pass
    return _u(val).strip() == u""


def _optional_row_1based(val):
    """Пусто / None → None; иначе int ≥ 1 или raise ValueError."""
    if val is None:
        return None
    s = _u(val).strip()
    if s == u"":
        return None
    n = int(s)
    if n < 1:
        raise ValueError("row < 1")
    return n


def normalize_question_key(text, trim=True, collapse_spaces=True, ignore_case=False, mode=u"none"):
    s = _u(text)
    if trim:
        s = s.strip()
    if collapse_spaces:
        parts = []
        for chunk in s.replace(u"\t", u" ").split(u" "):
            if chunk != u"":
                parts.append(chunk)
        s = u" ".join(parts)
    key = s
    nm = _u(mode).strip().casefold()
    if nm == u"lower":
        key = key.casefold()
    elif nm == u"title":
        try:
            key = key.title()
        except Exception:
            pass
    if ignore_case:
        key = key.casefold()
    return key, s


def resolve_col_ref(spec, headers=None):
    """
    col-ref → 0-based index.
    Accepts: 1-based int, 'A'/'B', header name (exact then substring).
    """
    from libre_macros_lib import col_letters_to_index, is_col_letters, lm_identity_key

    raw = _u(spec).strip()
    if raw == u"":
        return None
    exact = False
    token = raw
    try:
        from libre_macros_param_codec import unwrap_column_name_token, _clean_column_match_text

        token, exact = unwrap_column_name_token(raw)
        token = _clean_column_match_text(token)
    except Exception:
        token = raw
    token = _u(token).strip()
    if token == u"":
        return None
    if not exact:
        try:
            idx = int(token) - 1
            if idx >= 0:
                return idx
        except (TypeError, ValueError):
            pass
        if is_col_letters(token):
            try:
                return int(col_letters_to_index(token))
            except Exception:
                pass
    if headers is not None:
        key = lm_identity_key(token)
        i = 0
        while i < len(headers):
            if lm_identity_key(_u(headers[i])) == key:
                return i
            i += 1
        if not exact:
            try:
                from libre_macros_param_codec import column_header_matches_token

                i = 0
                while i < len(headers):
                    if column_header_matches_token(_u(headers[i]), raw):
                        return i
                    i += 1
            except Exception:
                pass
    return None


def _is_service_header(name, skip_block_index=True, skip_source_row=True):
    k = _u(name).strip().casefold()
    if skip_block_index:
        for n in _BLOCK_INDEX_NAMES:
            if k == n.casefold():
                return True
    if skip_source_row:
        for n in _SOURCE_ROW_NAMES:
            if k == n.casefold():
                return True
    return False


# ---------------------------------------------------------------------------
# Pure: анкета → wide
# ---------------------------------------------------------------------------

def split_form_blocks( pairs, block_mode, block_start_question=u"", blank_rows_to_split=1, questions_per_block=0, q_norm_opts=None):
    """
    pairs: list of dict {q, a, row, blank} where blank=True for fully empty rows.
    Returns list of blocks; each block = list of {q,a,row} (non-blank pairs only).
    """
    opts = q_norm_opts or {}
    mode = _u(block_mode).strip().casefold() or u"by_repeat_key"
    start_q_key = u""
    if _u(block_start_question).strip():
        start_q_key, _disp = normalize_question_key(
            block_start_question,
            trim=opts.get(u"trim", True),
            collapse_spaces=opts.get(u"collapse", True),
            ignore_case=opts.get(u"ignore_case", False),
            mode=opts.get(u"mode", u"none"),
        )
    n_blank_need = max(1, int(blank_rows_to_split or 1))
    qpb = int(questions_per_block or 0)

    blocks = []
    cur = []
    blank_run = 0
    seen_in_cur = set()

    def _flush():
        if cur:
            blocks.append(list(cur))
            cur[:] = []
            seen_in_cur.clear()

    for item in pairs or ():
        if item.get(u"blank"):
            blank_run += 1
            if mode == u"by_blank_row" and blank_run >= n_blank_need and cur:
                _flush()
            continue
        blank_run = 0
        q = item.get(u"q")
        a = item.get(u"a")
        row = item.get(u"row")
        q_key, q_disp = normalize_question_key(
            q,
            trim=opts.get(u"trim", True),
            collapse_spaces=opts.get(u"collapse", True),
            ignore_case=opts.get(u"ignore_case", False),
            mode=opts.get(u"mode", u"none"),
        )
        pair = {u"q": q_disp, u"q_key": q_key, u"a": a, u"row": row}

        if mode == u"by_repeat_key":
            if start_q_key and q_key == start_q_key and cur:
                _flush()
            cur.append(pair)
        elif mode == u"by_unique_cycle":
            if q_key in seen_in_cur and cur:
                _flush()
            cur.append(pair)
            if q_key:
                seen_in_cur.add(q_key)
        elif mode == u"fixed_size":
            cur.append(pair)
            if qpb > 0 and len(cur) >= qpb:
                _flush()
        else:  # by_blank_row — split only on blanks
            cur.append(pair)
    _flush()
    return blocks


def apply_duplicate_question(block_pairs, policy=u"last", concat_sep=u"; "):
    """block_pairs → ordered list of (display_q, q_key, answer) unique by key."""
    pol = _u(policy).strip().casefold() or u"last"
    order = []
    by_key = {}
    for p in block_pairs or ():
        k = p.get(u"q_key") or u""
        if k == u"" and _empty(p.get(u"q")):
            continue
        if k not in by_key:
            order.append(k)
            by_key[k] = {
                u"q": p.get(u"q"),
                u"a": p.get(u"a"),
                u"row": p.get(u"row"),
            }
        else:
            if pol == u"first":
                pass
            elif pol == u"concat":
                prev = by_key[k][u"a"]
                cur = p.get(u"a")
                if _empty(prev):
                    by_key[k][u"a"] = cur
                elif _empty(cur):
                    pass
                else:
                    by_key[k][u"a"] = _u(prev) + _u(concat_sep) + _u(cur)
            elif pol == u"error":
                return None, u"duplicate_question"
            else:  # last
                by_key[k][u"a"] = p.get(u"a")
                by_key[k][u"q"] = p.get(u"q")
    out = []
    for k in order:
        item = by_key[k]
        out.append((item[u"q"], k, item[u"a"]))
    return out, u""


def build_wide_matrix_from_blocks( blocks, sheet_preamble=None, column_order=None, extra_questions=u"append", add_block_index=False, add_source_row=False, preamble_columns_first=True, duplicate_policy=u"last", concat_sep=u"; ", keep_partial=True):
    """
    blocks: list of {body:[{q,q_key,a,row}], preamble:[(name,val)], start_row:int}
    sheet_preamble: ordered list of (name, val)
    Returns (matrix, err)
    """
    sheet_preamble = list(sheet_preamble or [])
    body_order = []
    body_disp = {}
    block_pre_order = []
    block_pre_disp = {}

    parsed_blocks = []
    for bi, blk in enumerate(blocks or ()):
        body_raw = blk.get(u"body") or []
        uniq, err = apply_duplicate_question(
            body_raw, policy=duplicate_policy, concat_sep=concat_sep
        )
        if err == u"duplicate_question":
            if not keep_partial:
                continue
            # skip block on error policy
            continue
        if not uniq and not (blk.get(u"preamble") or sheet_preamble):
            if not keep_partial:
                continue
        for q_disp, q_key, _a in uniq:
            if q_key not in body_disp:
                body_order.append(q_key)
                body_disp[q_key] = q_disp
        for name, _val in blk.get(u"preamble") or ():
            nk, nd = normalize_question_key(name)
            if nk not in block_pre_disp:
                block_pre_order.append(nk)
                block_pre_disp[nk] = nd
        parsed_blocks.append(
            {
                u"body": {k: a for (_q, k, a) in uniq},
                u"preamble": {
                    normalize_question_key(n)[0]: v
                    for (n, v) in (blk.get(u"preamble") or ())
                },
                u"start_row": blk.get(u"start_row"),
                u"index": bi + 1,
            }
        )

    if column_order:
        ordered = []
        seen = set()
        for q in column_order:
            k, d = normalize_question_key(q)
            if k in body_disp and k not in seen:
                ordered.append(k)
                seen.add(k)
            elif k not in body_disp:
                body_disp[k] = d
                ordered.append(k)
                seen.add(k)
        if _u(extra_questions).casefold() != u"drop":
            for k in body_order:
                if k not in seen:
                    ordered.append(k)
        body_order = ordered

    header = []
    if add_block_index:
        header.append(u"#Блок")
    if add_source_row:
        header.append(u"source_start_row")
    sheet_names = [n for (n, _v) in sheet_preamble]
    block_names = [block_pre_disp[k] for k in block_pre_order]
    body_names = [body_disp[k] for k in body_order]
    if preamble_columns_first:
        header.extend(sheet_names)
        header.extend(block_names)
        header.extend(body_names)
    else:
        header.extend(body_names)
        header.extend(sheet_names)
        header.extend(block_names)

    if len(header) > _MAX_WIDE_COLS:
        return None, u"слишком много столбцов (%d > %d)" % (len(header), _MAX_WIDE_COLS)
    if len(header) == 0:
        return None, u"нет столбцов результата"

    rows = []
    for pb in parsed_blocks:
        row = []
        if add_block_index:
            row.append(pb[u"index"])
        if add_source_row:
            row.append(pb.get(u"start_row") or u"")
        sheet_cells = []
        for n, v in sheet_preamble:
            sheet_cells.append(v)
        block_cells = [pb[u"preamble"].get(k, u"") for k in block_pre_order]
        body_cells = [pb[u"body"].get(k, u"") for k in body_order]
        if preamble_columns_first:
            row.extend(sheet_cells)
            row.extend(block_cells)
            row.extend(body_cells)
        else:
            row.extend(body_cells)
            row.extend(sheet_cells)
            row.extend(block_cells)
        rows.append(row)

    return [header] + rows, u""


# ---------------------------------------------------------------------------
# Pure: wide → анкета
# ---------------------------------------------------------------------------

def build_form_pairs_from_wide( header, body_rows, body_rel_indices, sheet_preamble_rel=None, block_preamble_rel=None, block_separator=u"blank_row", blank_rows_between=1, empty_answer=u"write", preamble_columns_first=True, has_header_in_output=True, question_header=u"Вопрос", answer_header=u"Ответ"):
    """
    Returns (pairs, sheet_preamble_pairs) where pairs is list of (q,a) or None for blank row.
    sheet_preamble_pairs written once before blocks.
    """
    sheet_preamble_rel = list(sheet_preamble_rel or [])
    block_preamble_rel = list(block_preamble_rel or [])
    body_rel = list(body_rel_indices or [])
    sep = _u(block_separator).strip().casefold() or u"blank_row"
    n_blank = max(0, int(blank_rows_between or 1))
    if sep == u"none":
        n_blank = 0
    skip_empty = _u(empty_answer).strip().casefold() == u"skip"

    sheet_pairs = []
    if body_rows and sheet_preamble_rel:
        first = list(body_rows[0])
        for ri in sheet_preamble_rel:
            q = header[ri] if 0 <= ri < len(header) else u""
            a = first[ri] if 0 <= ri < len(first) else u""
            sheet_pairs.append((q, a))

    out = []
    if has_header_in_output:
        out.append((_u(question_header) or u"Вопрос", _u(answer_header) or u"Ответ"))

    for bi, row in enumerate(body_rows or ()):
        row = list(row)
        block_pairs = []
        pre_pairs = []
        for ri in block_preamble_rel:
            q = header[ri] if 0 <= ri < len(header) else u""
            a = row[ri] if 0 <= ri < len(row) else u""
            pre_pairs.append((q, a))
        body_pairs = []
        for ri in body_rel:
            q = header[ri] if 0 <= ri < len(header) else u""
            a = row[ri] if 0 <= ri < len(row) else u""
            if skip_empty and _empty(a):
                continue
            body_pairs.append((q, a))
        if preamble_columns_first:
            block_pairs = pre_pairs + body_pairs
        else:
            block_pairs = body_pairs + pre_pairs
        if bi > 0 and n_blank > 0:
            for _ in range(n_blank):
                out.append(None)
        out.extend(block_pairs)

    return out, sheet_pairs


# ---------------------------------------------------------------------------
# UNO helpers
# ---------------------------------------------------------------------------

def _write_matrix(doc, sheet, dest_col, dest_row, matrix):
    from libre_macros_lib import _lm_pp_remove_duplicates_write_matrix

    return _lm_pp_remove_duplicates_write_matrix(doc, sheet, dest_col, dest_row, matrix)


def _ensure_dest_sheet(doc, dest_base, sheet_name, fn_label):
    from libre_macros_lib import (
        _lm_log_postprocess,
        _lm_pp_clear_sheet_content,
        _lm_pp_get_sheet_ref,
        _lm_pp_unique_sheet_name,
    )

    actual = _lm_pp_unique_sheet_name(doc, dest_base, exclude_sheet=None)
    dest_sh = _lm_pp_get_sheet_ref(doc, actual)
    if dest_sh is None:
        try:
            pos = int(doc.Sheets.getCount())
            doc.Sheets.insertNewByName(actual, pos)
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"создание листа «%s»: %s" % (actual, err),
            )
            return None, u""
        dest_sh = _lm_pp_get_sheet_ref(doc, actual)
    else:
        _lm_pp_clear_sheet_content(dest_sh)
    return dest_sh, actual


def _read_cell_value(sheet, col, row):
    try:
        cell = sheet.getCellByPosition(int(col), int(row))
    except Exception:
        return u""
    try:
        f = cell.Formula
        if f is not None and _u(f).strip() != u"" and _u(f).startswith(u"="):
            # prefer displayed/computed
            pass
    except Exception:
        pass
    try:
        t = cell.Type
        # 1=EMPTY, 2=VALUE, 3=TEXT, 4=FORMULA
        if int(t) == 1:
            return u""
        if int(t) == 2:
            return cell.Value
    except Exception:
        pass
    try:
        return cell.String
    except Exception:
        try:
            return cell.Value
        except Exception:
            return u""


def _resolve_qa_cols(block, sheet, header_row_range, sc, ec):
    from libre_macros_lib import _lm_pp_header_titles, _lm_pp_resolve_column_token

    headers = _lm_pp_header_titles(sheet, header_row_range) if header_row_range else None
    q_raw = block.get(u"question_column")
    a_raw = block.get(u"answer_column")
    if q_raw in (None, u"") and a_raw in (None, u""):
        q_raw, a_raw = u"A", u"B"
    q_col = _lm_pp_resolve_column_token(q_raw, sheet, header_row_range, sc, ec)
    if q_col is None:
        q_col = resolve_col_ref(q_raw, headers)
    a_col = _lm_pp_resolve_column_token(a_raw, sheet, header_row_range, sc, ec)
    if a_col is None:
        a_col = resolve_col_ref(a_raw, headers)
    return q_col, a_col


def _q_norm_opts_from_block(block):
    return {
        u"trim": True if u"trim_questions" not in block else bool(block.get(u"trim_questions")),
        u"collapse": True
        if u"collapse_spaces" not in block
        else bool(block.get(u"collapse_spaces")),
        u"ignore_case": bool(block.get(u"ignore_case")),
        u"mode": _u(block.get(u"normalize_question") or u"none"),
    }


def _trim_answer(val, do_trim=True):
    if not do_trim:
        return val
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return val
    s = _u(val)
    return s.strip()


# ---------------------------------------------------------------------------
# lm_pp_range_form_to_table  (анкета_в_таблицу)
# ---------------------------------------------------------------------------

def lm_pp_range_form_to_table(doc, sheet, data_range, header_row_range, *extra_args):
    from libre_macros_lib import (
        _lm_log_postprocess,
        _lm_pp_calculate_with_auto,
        _lm_pp_clear_sheet_content,
        _lm_pp_copy_range_as_values,
        _lm_pp_get_sheet_ref,
        _lm_pp_param_block_for_sheet,
        _lm_pp_param_skip_sheet,
        _lm_pp_parse_as_values_flag,
        _lm_pp_parse_cell_a1_ref,
        _lm_pp_range_address,
        lm_pp_copy_sheet_register_pending,
        lm_pp_active_context,
        _lm_pp_context_refresh,
    )

    fn_label = u"анкета_в_таблицу"
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

    q_col, a_col = _resolve_qa_cols(block, sheet, header_row_range, sc, ec)
    if q_col is None or a_col is None or q_col == a_col:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"нужны разные question_column / answer_column",
        )
        return

    as_values = True
    if u"as_values" in block:
        as_values = bool(_lm_pp_parse_as_values_flag(block))
    if as_values and doc is not None:
        try:
            c0 = min(q_col, a_col)
            c1 = max(q_col, a_col)
            mat_rng = sheet.getCellRangeByPosition(c0, 0, c1, max(er, h_er) + 50)
            _lm_pp_calculate_with_auto(doc, sheet)
            _lm_pp_copy_range_as_values(doc, sheet, mat_rng)
        except Exception:
            pass

    # sheet preamble absolute rows (1-based). Пустые поля визарда → авто по диапазону.
    # sheet_preamble_row_from/to (Шапка листа: с / по)
    # forms_start_row (Строка начала анкет)
    sheet_preamble = []
    data_start_1 = int(max(sr, h_sr + 1 if header_row_range is not None else sr)) + 1
    sp_from_raw = block.get(u"sheet_preamble_row_from")
    sp_to_raw = block.get(u"sheet_preamble_row_to")
    forms_raw = block.get(u"forms_start_row")
    sp_from = None
    sp_to = None
    forms_start_1 = None
    try:
        sp_from = _optional_row_1based(sp_from_raw)
        sp_to = _optional_row_1based(sp_to_raw)
        forms_start_1 = _optional_row_1based(forms_raw)
    except Exception:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"sheet_preamble_row_from/to / forms_start_row "
            u"(Шапка листа: с / по, Строка начала анкет) — целые номера строк ≥ 1",
        )
        return

    if (sp_from is None) ^ (sp_to is None):
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"sheet_preamble_row_from и sheet_preamble_row_to "
            u"(Шапка листа: с / по) задайте оба или оставьте пустыми",
        )
        return

    # Авто: задана только «Строка начала анкет» → шапка листа =
    # строки диапазона данных от начала до forms_start−1
    if sp_from is None and sp_to is None and forms_start_1 is not None:
        if forms_start_1 > data_start_1:
            sp_from = data_start_1
            sp_to = forms_start_1 - 1
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ok",
                u"sheet_preamble_row_from/to (Шапка листа: с / по) авто = %d…%d "
                u"по диапазону до forms_start_row"
                % (sp_from, sp_to),
            )

    if sp_from is not None and sp_to is not None:
        if sp_from > sp_to or sp_from < 1:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"неверный диапазон шапки листа "
                u"(sheet_preamble_row_from/to = Шапка листа: с / по)",
            )
            return
        if forms_start_1 is None:
            forms_start_1 = sp_to + 1
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ok",
                u"forms_start_row (Строка начала анкет) авто = %d" % forms_start_1,
            )
        if forms_start_1 <= sp_to:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"forms_start_row (Строка начала анкет) пересекает шапку листа "
                u"(sheet_preamble_row_to)",
            )
            return
        sp_mode = _u(block.get(u"sheet_preamble_mode") or u"qa_pair").casefold()
        if sp_mode == u"qa_pair":
            r = sp_from - 1
            while r <= sp_to - 1:
                qv = _read_cell_value(sheet, q_col, r)
                av = _read_cell_value(sheet, a_col, r)
                if not _empty(qv):
                    sheet_preamble.append((_u(qv), av))
                r += 1
        elif sp_mode == u"fixed_cells":
            fields = block.get(u"sheet_preamble_fields") or []
            if not fields:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    u"диапазон",
                    fn_label,
                    u"ошибка",
                    u"sheet_preamble_mode=fixed_cells без fields",
                )
                return
            for fld in fields:
                if not isinstance(fld, dict):
                    continue
                cell = _u(fld.get(u"cell") or u"").strip()
                as_name = _u(fld.get(u"as") or cell).strip()
                ref = _lm_pp_parse_cell_a1_ref(cell)
                if ref is None:
                    continue
                sheet_preamble.append(
                    (as_name, _read_cell_value(sheet, ref[0], ref[1]))
                )
        # ignore → empty sheet_preamble
    elif forms_start_1 is None:
        # Нет шапки листа и нет явного старта → диапазон данных сбора
        forms_start_1 = data_start_1

    # stream start
    if forms_start_1 is not None:
        stream_r0 = int(forms_start_1) - 1
    else:
        start_row = block.get(u"start_row")
        if start_row is not None and _u(start_row).strip() != u"":
            try:
                stream_r0 = int(start_row) - 1
            except Exception:
                stream_r0 = max(sr, 0)
        else:
            stream_r0 = max(sr, 0)
            if bool(block.get(u"has_header_in_source")):
                stream_r0 = max(stream_r0, h_sr + 1) if header_row_range else stream_r0 + 1

    end_row = block.get(u"end_row")
    if end_row is not None:
        try:
            stream_r1 = int(end_row) - 1
        except Exception:
            stream_r1 = er
    else:
        # scan used
        stream_r1 = er
        try:
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(False)
            stream_r1 = max(stream_r1, int(cursor.getRangeAddress().EndRow))
        except Exception:
            pass

    if bool(block.get(u"has_header_in_source")):
        # skip first stream row
        stream_r0 = stream_r0 + 1

    # read pairs
    trim_ans = True if u"trim_answers" not in block else bool(block.get(u"trim_answers"))
    pairs = []
    r = stream_r0
    while r <= stream_r1:
        qv = _read_cell_value(sheet, q_col, r)
        av = _read_cell_value(sheet, a_col, r)
        if _empty(qv) and _empty(av):
            pairs.append({u"blank": True, u"row": r + 1})
        elif _empty(qv) and not _empty(av):
            # orphan
            pol = _u(block.get(u"orphan_answer") or u"skip").casefold()
            if pol == u"attach_prev" and pairs and not pairs[-1].get(u"blank"):
                prev = pairs[-1]
                prev[u"a"] = _u(prev.get(u"a")) + u"; " + _u(_trim_answer(av, trim_ans))
            # else skip
        else:
            pairs.append(
                {
                    u"q": qv,
                    u"a": _trim_answer(av, trim_ans),
                    u"row": r + 1,
                    u"blank": False,
                }
            )
        r += 1

    # block preamble height
    bp_rows = 0
    if block.get(u"block_preamble_rows") is not None:
        try:
            bp_rows = max(0, int(block.get(u"block_preamble_rows")))
        except Exception:
            bp_rows = 0
    if block.get(u"block_preamble_row_from") is not None and block.get(
        u"block_preamble_row_to"
    ) is not None:
        try:
            bf = int(block.get(u"block_preamble_row_from"))
            bt = int(block.get(u"block_preamble_row_to"))
            if bt >= bf >= 1:
                bp_rows = bt - bf + 1
        except Exception:
            pass
    bp_until = _u(
        block.get(u"block_preamble_until_question")
        or block.get(u"preamble_until_question")
        or u""
    ).strip()
    bp_mode = _u(
        block.get(u"block_preamble_mode") or block.get(u"preamble_mode") or u"qa_pair"
    ).casefold()

    # If block preamble via until_question — handled after split by peeling
    # First split raw stream ignoring preamble height for by_repeat_key on body key.
    # Strategy: if bp_rows > 0, peel first N non-blank from each provisional block...
    # Better: convert pairs to body-only for split when using until_question.

    opts = _q_norm_opts_from_block(block)
    block_mode = _u(block.get(u"block_mode") or u"by_repeat_key").casefold()
    start_q = _u(block.get(u"block_start_question") or u"")
    if block_mode == u"by_repeat_key" and start_q == u"":
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"by_repeat_key: нужен block_start_question",
        )
        return
    qpb = 0
    try:
        qpb = int(block.get(u"questions_per_block") or 0)
    except Exception:
        qpb = 0
    if block_mode == u"fixed_size" and qpb <= 0:
        try:
            total = int(block.get(u"block_total_rows") or 0)
            qpb = max(0, total - bp_rows)
        except Exception:
            qpb = 0
        if qpb <= 0:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"fixed_size: нужен questions_per_block",
            )
            return

    blank_n = 1
    try:
        blank_n = int(block.get(u"blank_rows_to_split") or 1)
    except Exception:
        blank_n = 1

    # When fixed preamble rows:    blank_n = 1
    try:
        blank_n = int(block.get(u"blank_rows_to_split") or 1)
    except Exception:
        blank_n = 1

    # by_repeat_key + block_preamble_rows: блок = N строк шапки + тело с ключа до следующего ключа
    structured = []
    if block_mode == u"by_repeat_key" and bp_rows > 0 and start_q:
        start_key, _ = normalize_question_key(
            start_q,
            trim=opts[u"trim"],
            collapse_spaces=opts[u"collapse"],
            ignore_case=opts[u"ignore_case"],
            mode=opts[u"mode"],
        )
        key_idxs = []
        pi = 0
        while pi < len(pairs):
            p = pairs[pi]
            if not p.get(u"blank"):
                k, _d = normalize_question_key(
                    p.get(u"q"),
                    trim=opts[u"trim"],
                    collapse_spaces=opts[u"collapse"],
                    ignore_case=opts[u"ignore_case"],
                    mode=opts[u"mode"],
                )
                if k == start_key:
                    key_idxs.append(pi)
            pi += 1
        bi = 0
        while bi < len(key_idxs):
            ki = key_idxs[bi]
            next_ki = key_idxs[bi + 1] if bi + 1 < len(key_idxs) else len(pairs)
            pre_start = max(0, ki - bp_rows)
            if bi > 0:
                pre_start = max(pre_start, key_idxs[bi - 1] + 1)
            pre_pairs = []
            body_pairs = []
            pj = pre_start
            while pj < ki:
                p = pairs[pj]
                if not p.get(u"blank"):
                    pre_pairs.append(p)
                pj += 1
            pj = ki
            while pj < next_ki:
                p = pairs[pj]
                if not p.get(u"blank"):
                    body_pairs.append(p)
                pj += 1
            preamble = []
            if bp_mode == u"qa_pair":
                for p in pre_pairs:
                    k, d = normalize_question_key(
                        p.get(u"q"),
                        trim=opts[u"trim"],
                        collapse_spaces=opts[u"collapse"],
                        ignore_case=opts[u"ignore_case"],
                        mode=opts[u"mode"],
                    )
                    preamble.append((d, p.get(u"a")))
            elif bp_mode == u"fixed_cells":
                fields = block.get(u"block_preamble_fields") or []
                start_row = pre_pairs[0][u"row"] if pre_pairs else (
                    body_pairs[0][u"row"] if body_pairs else None
                )
                if start_row is not None:
                    base = int(start_row) - 1
                    for fld in fields:
                        if not isinstance(fld, dict):
                            continue
                        cell = _u(fld.get(u"cell") or u"").strip()
                        as_name = _u(fld.get(u"as") or cell).strip()
                        ref = _lm_pp_parse_cell_a1_ref(cell)
                        if ref is None:
                            continue
                        preamble.append(
                            (as_name, _read_cell_value(sheet, ref[0], base + ref[1]))
                        )
            body2 = []
            for p in body_pairs:
                k, d = normalize_question_key(
                    p.get(u"q"),
                    trim=opts[u"trim"],
                    collapse_spaces=opts[u"collapse"],
                    ignore_case=opts[u"ignore_case"],
                    mode=opts[u"mode"],
                )
                body2.append(
                    {u"q": d, u"q_key": k, u"a": p.get(u"a"), u"row": p.get(u"row")}
                )
            start_row = None
            if pre_pairs:
                start_row = pre_pairs[0].get(u"row")
            elif body_pairs:
                start_row = body_pairs[0].get(u"row")
            structured.append(
                {u"body": body2, u"preamble": preamble, u"start_row": start_row}
            )
            bi += 1
    else:
        raw_blocks = split_form_blocks(
            pairs,
            block_mode,
            block_start_question=start_q,
            blank_rows_to_split=blank_n,
            questions_per_block=qpb + bp_rows
            if (bp_rows > 0 and block_mode == u"fixed_size")
            else qpb,
            q_norm_opts=opts,
        )

        until_key = u""
        if bp_until:
            until_key, _ = normalize_question_key(
                bp_until,
                trim=opts[u"trim"],
                collapse_spaces=opts[u"collapse"],
                ignore_case=opts[u"ignore_case"],
                mode=opts[u"mode"],
            )

        for rb in raw_blocks:
            start_row = rb[0][u"row"] if rb else None
            preamble = []
            body = list(rb)
            if until_key:
                pre = []
                rest = []
                hit = False
                for p in rb:
                    k, d = normalize_question_key(
                        p.get(u"q"),
                        trim=opts[u"trim"],
                        collapse_spaces=opts[u"collapse"],
                        ignore_case=opts[u"ignore_case"],
                        mode=opts[u"mode"],
                    )
                    p = dict(p)
                    p[u"q_key"] = k
                    p[u"q"] = d
                    if not hit and k == until_key:
                        hit = True
                        rest.append(p)
                    elif not hit:
                        pre.append(p)
                    else:
                        rest.append(p)
                if bp_mode == u"qa_pair":
                    preamble = [(p[u"q"], p[u"a"]) for p in pre]
                body = rest
            elif bp_rows > 0:
                pre = body[:bp_rows]
                body = body[bp_rows:]
                if bp_mode == u"qa_pair":
                    preamble = []
                    for p in pre:
                        k, d = normalize_question_key(
                            p.get(u"q"),
                            trim=opts[u"trim"],
                            collapse_spaces=opts[u"collapse"],
                            ignore_case=opts[u"ignore_case"],
                            mode=opts[u"mode"],
                        )
                        preamble.append((d, p.get(u"a")))
                elif bp_mode == u"fixed_cells":
                    fields = block.get(u"block_preamble_fields") or []
                    if start_row is not None:
                        base = int(start_row) - 1
                        for fld in fields:
                            if not isinstance(fld, dict):
                                continue
                            cell = _u(fld.get(u"cell") or u"").strip()
                            as_name = _u(fld.get(u"as") or cell).strip()
                            ref = _lm_pp_parse_cell_a1_ref(cell)
                            if ref is None:
                                continue
                            preamble.append(
                                (
                                    as_name,
                                    _read_cell_value(sheet, ref[0], base + ref[1]),
                                )
                            )
            body2 = []
            for p in body:
                k, d = normalize_question_key(
                    p.get(u"q"),
                    trim=opts[u"trim"],
                    collapse_spaces=opts[u"collapse"],
                    ignore_case=opts[u"ignore_case"],
                    mode=opts[u"mode"],
                )
                body2.append(
                    {u"q": d, u"q_key": k, u"a": p.get(u"a"), u"row": p.get(u"row")}
                )
            structured.append(
                {u"body": body2, u"preamble": preamble, u"start_row": start_row}
            )

    keep_partial = True if u"keep_partial_blocks" not in block else bool(
        block.get(u"keep_partial_blocks")
    )
    col_order = block.get(u"column_order")
    if isinstance(col_order, str):
        col_order = [x.strip() for x in col_order.replace(u";", u",").split(u",") if x.strip()]
    matrix, err = build_wide_matrix_from_blocks(
        structured,
        sheet_preamble=sheet_preamble,
        column_order=col_order,
        extra_questions=_u(block.get(u"extra_questions") or u"append"),
        add_block_index=bool(block.get(u"add_block_index")),
        add_source_row=bool(block.get(u"add_source_row")),
        preamble_columns_first=True
        if u"preamble_columns_first" not in block
        else bool(block.get(u"preamble_columns_first")),
        duplicate_policy=_u(block.get(u"duplicate_question") or u"last"),
        concat_sep=_u(block.get(u"concat_sep") or u"; "),
        keep_partial=keep_partial,
    )
    if err:
        _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ошибка", err)
        return
    if not matrix or len(matrix) <= 1:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"0 блоков анкет"
        )
        return

    output = _u(block.get(u"output") or u"new_sheet").casefold()
    if output not in (u"new_sheet", u"inplace", u"replace_sheet"):
        output = u"new_sheet"

    if output == u"new_sheet":
        dest_base = _u(block.get(u"dest_sheet") or block.get(u"dest") or u"Анкета_таблица").strip()
        if dest_base == u"":
            dest_base = u"Анкета_таблица"
        dest_sh, actual = _ensure_dest_sheet(doc, dest_base, sheet_name, fn_label)
        if dest_sh is None:
            return
        ok, werr = _write_matrix(doc, dest_sh, 0, 0, matrix)
        if not ok:
            _lm_log_postprocess(
                doc, sheet_name, u"диапазон", fn_label, u"ошибка", werr or u"запись"
            )
            return
        try:
            lm_pp_copy_sheet_register_pending(actual, 0, sheet_name, from_dedup=False)
        except Exception:
            pass
        note = u"блоков %d, столбцов %d → «%s»" % (
            len(matrix) - 1,
            len(matrix[0]),
            actual,
        )
        _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
        return

    if output == u"replace_sheet":
        _lm_pp_clear_sheet_content(sheet)
        ok, werr = _write_matrix(doc, sheet, 0, 0, matrix)
        if not ok:
            _lm_log_postprocess(
                doc, sheet_name, u"диапазон", fn_label, u"ошибка", werr or u"запись"
            )
            return
        note = u"блоков %d, столбцов %d, replace_sheet" % (
            len(matrix) - 1,
            len(matrix[0]),
        )
        _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
        return

    # inplace
    dest_cell = _u(block.get(u"dest_cell") or u"A1")
    anchor = _lm_pp_parse_cell_a1_ref(dest_cell) or (0, 0)
    clear_first = bool(block.get(u"clear_source_first"))
    # overlap check: dest covers source QA columns
    dest_ec = anchor[0] + len(matrix[0]) - 1
    dest_er = anchor[1] + len(matrix) - 1
    src_c0, src_c1 = min(q_col, a_col), max(q_col, a_col)
    overlap = not (dest_ec < src_c0 or anchor[0] > src_c1 or dest_er < stream_r0 or anchor[1] > stream_r1)
    if overlap and not clear_first:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"inplace перекрывает источник — нужен clear_source_first",
        )
        return
    if clear_first:
        try:
            sheet.getCellRangeByPosition(src_c0, stream_r0, src_c1, stream_r1).clearContents(7)
        except Exception:
            pass
    ok, werr = _write_matrix(doc, sheet, anchor[0], anchor[1], matrix)
    if not ok:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"ошибка", werr or u"запись"
        )
        return
    try:
        ctx = lm_pp_active_context()
        if ctx is not None:
            _lm_pp_context_refresh(ctx)
    except Exception:
        pass
    _lm_log_postprocess(
        doc,
        sheet_name,
        u"диапазон",
        fn_label,
        u"ok",
        u"блоков %d, столбцов %d, inplace @ %s"
        % (len(matrix) - 1, len(matrix[0]), dest_cell),
    )


# ---------------------------------------------------------------------------
# lm_pp_range_table_to_form  (таблица_в_анкету)
# ---------------------------------------------------------------------------

def lm_pp_range_table_to_form(doc, sheet, data_range, header_row_range, *extra_args):
    from libre_macros_lib import (
        _lm_log_postprocess,
        _lm_pp_calculate_with_auto,
        _lm_pp_clear_sheet_content,
        _lm_pp_copy_range_as_values,
        _lm_pp_param_block_for_sheet,
        _lm_pp_param_skip_sheet,
        _lm_pp_parse_as_values_flag,
        _lm_pp_parse_cell_a1_ref,
        _lm_pp_range_address,
        _lm_pp_resolve_column_token,
        _lm_pp_header_titles,
        lm_pp_copy_sheet_register_pending,
        lm_pp_active_context,
        _lm_pp_context_refresh,
        col_letters_to_index,
        is_col_letters,
    )

    fn_label = u"таблица_в_анкету"
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

    as_values = True
    if u"as_values" in block:
        as_values = bool(_lm_pp_parse_as_values_flag(block))
    if as_values and doc is not None:
        try:
            mat_rng = sheet.getCellRangeByPosition(int(sc), int(h_sr), int(ec), int(er))
            _lm_pp_calculate_with_auto(doc, sheet)
            _lm_pp_copy_range_as_values(doc, sheet, mat_rng)
        except Exception:
            pass

    try:
        data = sheet.getCellRangeByPosition(int(sc), int(h_sr), int(ec), int(er)).getDataArray()
    except Exception as err:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"ошибка", u"getDataArray: %s" % err
        )
        return
    if not data:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"пустой диапазон"
        )
        return

    header = list(data[0])
    n_cols = int(ec) - int(sc) + 1
    while len(header) < n_cols:
        header.append(u"")
    header = header[:n_cols]
    body = []
    skip_empty_rows = True if u"skip_empty_rows" not in block else bool(
        block.get(u"skip_empty_rows")
    )
    for i in range(1, len(data)):
        row = list(data[i])
        while len(row) < n_cols:
            row.append(u"")
        row = row[:n_cols]
        if skip_empty_rows and all(_empty(x) for x in row):
            continue
        body.append(row)
    if not body:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нет строк данных"
        )
        return

    def _tok_list(key):
        raw = block.get(key) or []
        if not isinstance(raw, (list, tuple)):
            raw = [raw]
        out = []
        for t in raw:
            col = _lm_pp_resolve_column_token(t, sheet, header_row_range, sc, ec)
            if col is None:
                continue
            if sc <= col <= ec:
                out.append(col - sc)
        return out

    skip_rel = set(_tok_list(u"skip_columns"))
    skip_bi = True if u"skip_block_index" not in block else bool(block.get(u"skip_block_index"))
    skip_sr = True if u"skip_source_row" not in block else bool(block.get(u"skip_source_row"))
    i = 0
    while i < len(header):
        if _is_service_header(header[i], skip_bi, skip_sr):
            skip_rel.add(i)
        i += 1

    sheet_pre_rel = _tok_list(u"sheet_preamble_columns")
    block_pre_rel = _tok_list(u"block_preamble_columns")
    body_rel = _tok_list(u"body_columns")
    if not body_rel and block.get(u"column_order"):
        co = block.get(u"column_order")
        if isinstance(co, str):
            co = [x.strip() for x in co.replace(u";", u",").split(u",") if x.strip()]
        for t in co or ():
            col = _lm_pp_resolve_column_token(t, sheet, header_row_range, sc, ec)
            if col is not None and sc <= col <= ec:
                body_rel.append(col - sc)
    reserved = set(sheet_pre_rel) | set(block_pre_rel) | skip_rel
    if not body_rel:
        extra = _u(block.get(u"extra_columns") or u"drop").casefold()
        body_rel = []
        i = 0
        while i < n_cols:
            if i not in reserved:
                body_rel.append(i)
            i += 1
        if extra == u"drop":
            pass
    else:
        # strip preamble cols from body
        body_rel = [i for i in body_rel if i not in reserved]

    if not body_rel and not sheet_pre_rel and not block_pre_rel:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"нет столбцов тела/шапок для анкеты",
        )
        return

    # dest Q/A columns
    q_spec = block.get(u"question_column") or u"A"
    a_spec = block.get(u"answer_column") or u"B"
    q_out = resolve_col_ref(q_spec)
    a_out = resolve_col_ref(a_spec)
    if q_out is None:
        q_out = 0
    if a_out is None:
        a_out = 1
    if q_out == a_out:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"question_column и answer_column совпадают",
        )
        return

    sep = _u(block.get(u"block_separator") or u"blank_row").casefold()
    blank_n = 1
    try:
        blank_n = int(block.get(u"blank_rows_between") or 1)
    except Exception:
        blank_n = 1
    pairs, sheet_pairs = build_form_pairs_from_wide(
        header,
        body,
        body_rel,
        sheet_preamble_rel=sheet_pre_rel,
        block_preamble_rel=block_pre_rel,
        block_separator=sep,
        blank_rows_between=blank_n,
        empty_answer=_u(block.get(u"empty_answer") or u"write"),
        preamble_columns_first=True
        if u"preamble_columns_first" not in block
        else bool(block.get(u"preamble_columns_first")),
        has_header_in_output=(
            True
            if u"has_header_in_output" not in block
            else bool(block.get(u"has_header_in_output"))
        ),
        question_header=_u(block.get(u"question_header") or u"Вопрос"),
        answer_header=_u(block.get(u"answer_header") or u"Ответ"),
    )

    # Build 2-col matrix for stream (sheet preamble + optional blank + forms)
    stream = []
    sp_mode = _u(block.get(u"sheet_preamble_mode") or u"")
    if sheet_pre_rel and sp_mode != u"fixed_cells":
        if not sp_mode:
            sp_mode = u"qa_pair"
    if sheet_pairs and _u(sp_mode or u"qa_pair").casefold() == u"qa_pair":
        for q, a in sheet_pairs:
            stream.append((q, a))
        if True if u"sheet_preamble_blank_after" not in block else bool(
            block.get(u"sheet_preamble_blank_after")
        ):
            stream.append(None)

    # forms_start_row on dest — pad blank rows if needed after dest_cell
    for p in pairs:
        stream.append(p)

    if len([x for x in stream if x is not None]) > _MAX_OUT_PAIRS:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"слишком много пар (> %d)" % _MAX_OUT_PAIRS,
        )
        return

    # matrix with only two logical columns at q_out/a_out — write as contiguous if A/B
    # Build full matrix starting at min(q,a)
    c0 = min(q_out, a_out)
    width = max(q_out, a_out) - c0 + 1
    matrix = []
    for item in stream:
        row = [u""] * width
        if item is None:
            matrix.append(row)
            continue
        q, a = item
        row[q_out - c0] = q
        row[a_out - c0] = a
        matrix.append(row)

    output = _u(block.get(u"output") or u"new_sheet").casefold()
    if output not in (u"new_sheet", u"inplace", u"replace_sheet"):
        output = u"new_sheet"

    dest_cell = _u(block.get(u"dest_cell") or u"A1")
    anchor = _lm_pp_parse_cell_a1_ref(dest_cell) or (0, 0)
    # if question/answer are A/B, prefer writing so that q_out maps to absolute col
    write_col = c0
    write_row = anchor[1]
    # If dest_cell is A1 but q_out is 0, fine. If dest is B1 and q is A — unusual; use dest as top-left of matrix.
    if anchor[0] != 0 or c0 != 0:
        # write matrix at dest_cell as top-left; relative q/a within matrix
        write_col = anchor[0]
        # rebuild matrix relative to dest: columns from 0..width-1 still
        pass

    forms_start = block.get(u"forms_start_row")
    if forms_start is not None:
        try:
            # 1-based dest row for first form block — pad
            fs = int(forms_start) - 1
            # sheet preamble already in stream; if forms_start > current length, pad
            # simpler: set write_row and insert blanks at front of form part — skip for v1 if stream already ordered
            write_row = min(write_row, fs)
        except Exception:
            pass

    if output == u"new_sheet":
        dest_base = _u(block.get(u"dest_sheet") or block.get(u"dest") or u"Анкета").strip()
        if dest_base == u"":
            dest_base = u"Анкета"
        dest_sh, actual = _ensure_dest_sheet(doc, dest_base, sheet_name, fn_label)
        if dest_sh is None:
            return
        # fixed_cells sheet preamble on dest
        if _u(block.get(u"sheet_preamble_mode") or u"").casefold() == u"fixed_cells":
            fields = block.get(u"sheet_preamble_fields") or []
            first = body[0] if body else []
            for fld in fields:
                if not isinstance(fld, dict):
                    continue
                cell = _u(fld.get(u"cell") or u"").strip()
                as_name = _u(fld.get(u"as") or u"").strip()
                ref = _lm_pp_parse_cell_a1_ref(cell)
                if ref is None:
                    continue
                val = u""
                vf = _u(fld.get(u"value_from") or u"").strip()
                if vf.lower().startswith(u"col:"):
                    tok = vf.split(u":", 1)[1].strip()
                    ci = _lm_pp_resolve_column_token(tok, sheet, header_row_range, sc, ec)
                    if ci is not None:
                        val = first[ci - sc] if 0 <= ci - sc < len(first) else u""
                elif u"value" in fld:
                    val = fld.get(u"value")
                try:
                    if as_name:
                        dest_sh.getCellByPosition(q_out, ref[1]).String = as_name
                    dest_sh.getCellByPosition(ref[0], ref[1]).String = _u(val)
                except Exception:
                    pass
        ok, werr = _write_matrix(doc, dest_sh, write_col, write_row, matrix)
        if not ok:
            _lm_log_postprocess(
                doc, sheet_name, u"диапазон", fn_label, u"ошибка", werr or u"запись"
            )
            return
        try:
            lm_pp_copy_sheet_register_pending(actual, 0, sheet_name, from_dedup=False)
        except Exception:
            pass
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ok",
            u"строк wide %d → пар/строк %d на «%s»"
            % (len(body), len(matrix), actual),
        )
        return

    if output == u"replace_sheet":
        _lm_pp_clear_sheet_content(sheet)
        ok, werr = _write_matrix(doc, sheet, write_col, write_row, matrix)
        if not ok:
            _lm_log_postprocess(
                doc, sheet_name, u"диапазон", fn_label, u"ошибка", werr or u"запись"
            )
            return
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ok",
            u"строк wide %d → %d строк, replace_sheet" % (len(body), len(matrix)),
        )
        return

    # inplace
    clear_first = bool(block.get(u"clear_source_first"))
    dest_ec = write_col + width - 1
    dest_er = write_row + len(matrix) - 1
    overlap = not (
        dest_ec < sc or write_col > ec or dest_er < h_sr or write_row > er
    )
    if overlap and not clear_first:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"inplace перекрывает источник — нужен clear_source_first",
        )
        return
    if clear_first:
        try:
            sheet.getCellRangeByPosition(sc, h_sr, ec, er).clearContents(7)
        except Exception:
            pass
    ok, werr = _write_matrix(doc, sheet, write_col, write_row, matrix)
    if not ok:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"ошибка", werr or u"запись"
        )
        return
    try:
        ctx = lm_pp_active_context()
        if ctx is not None:
            _lm_pp_context_refresh(ctx)
    except Exception:
        pass
    _lm_log_postprocess(
        doc,
        sheet_name,
        u"диапазон",
        fn_label,
        u"ok",
        u"строк wide %d → %d строк, inplace" % (len(body), len(matrix)),
    )
