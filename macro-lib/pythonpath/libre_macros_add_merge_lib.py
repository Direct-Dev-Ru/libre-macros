# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.800"
"""
добавить_объединение — вертикальный merge выбранных столбцов по блокам.

Обратная к убрать_объединение: строки не удаляются; значения блока
склеиваются через sep в верхнюю ячейку, затем merge(True).

Режимы блоков:
  • block_rows — постоянная высота;
  • marker — опорный столбец с уже существующими вертикальными merge.

Карта: lm_pp_range_add_merge / merge_pp_range_add_merge.
"""

try:
    unicode
except NameError:
    unicode = str


_FN = u"добавить_объединение"


def _u(v):
    return unicode(v if v is not None else u"")


def _merge_column_block(sheet, col, r0, r1, sep=u" "):
    """
    Собрать значения столбца col в [r0..r1] → верхняя ячейка через sep,
    снять старые merge в этом столбце в диапазоне, merge вертикально.
    Возвращает True если merge выполнен (высота ≥ 2).
    """
    from libre_macros_header_lib import uno_merges_in_rect
    from libre_macros_unmerge_collapse_lib import (
        _cell_string,
        _empty_text,
        _join_tokens,
        _set_cell_string,
        unmerge_rect,
    )

    r0 = int(r0)
    r1 = int(r1)
    col = int(col)
    if r1 <= r0:
        return False

    # Значения с учётом уже существующих merge (только TL)
    merges = uno_merges_in_rect(sheet, col, r0, col, r1)
    owners = {}
    for sc, sr, ec, er in merges:
        rr = max(int(sr), r0)
        while rr <= min(int(er), r1):
            owners[(col, rr)] = (int(sc), int(sr), int(ec), int(er))
            rr += 1

    tokens = []
    r = r0
    while r <= r1:
        owner = owners.get((col, r))
        if owner is not None:
            sc, sr, _ec, _er = owner
            if col == sc and r == sr:
                t = _cell_string(sheet, sc, sr)
                if not _empty_text(t):
                    tokens.append(t)
        else:
            t = _cell_string(sheet, col, r)
            if not _empty_text(t):
                tokens.append(t)
        r += 1

    joined = _join_tokens(tokens, sep)

    # Снять пересекающие merge в этом столбце, чтобы merge(True) не падал
    unmerge_rect(sheet, merges)

    _set_cell_string(sheet, col, r0, joined)
    r = r0 + 1
    while r <= r1:
        _set_cell_string(sheet, col, r, u"")
        r += 1

    try:
        rng = sheet.getCellRangeByPosition(col, r0, col, r1)
        rng.merge(True)
    except Exception:
        return False
    return True


def _resolve_columns(block, sheet, header_row_range, sc, ec):
    """Список 0-based индексов столбцов из columns (tokens)."""
    from libre_macros_lib import (
        _lm_pp_resolve_column_token,
        col_letters_to_index,
        is_col_letters,
    )
    from libre_macros_unmerge_collapse_lib import _u

    raw = block.get(u"columns") or []
    if not isinstance(raw, (list, tuple)):
        raw = [raw]
    cols = []
    seen = set()
    i = 0
    while i < len(raw):
        token = _u(raw[i]).strip()
        i += 1
        if token == u"":
            continue
        col = _lm_pp_resolve_column_token(token, sheet, header_row_range, sc, ec)
        if col is None:
            try:
                if is_col_letters(token):
                    col = col_letters_to_index(token)
                else:
                    col = int(token) - 1
            except Exception:
                col = None
        if col is None or col < 0:
            return None, u"столбец не найден: %s" % token
        col = int(col)
        if col in seen:
            continue
        seen.add(col)
        cols.append(col)
    if not cols:
        return None, u"columns: нужен хотя бы один столбец"
    return cols, None


def lm_pp_range_add_merge(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Вертикально объединить выбранные столбцы по блокам.

    JSON:
      [{"v":1,"fn":"добавить_объединение","columns":["B","C"],"marker":"A","sep":" "}]
      [{"v":1,"fn":"добавить_объединение","columns":["A"],"block_rows":2,"sep":" | "}]
    """
    from libre_macros_lib import (
        _lm_log_postprocess,
        _lm_pp_context_refresh,
        _lm_pp_param_block_for_sheet,
        _lm_pp_param_skip_sheet,
        _lm_pp_range_address,
        _lm_vlookup_sheet_used_area,
        lm_pp_active_context,
    )
    from libre_macros_unmerge_collapse_lib import (
        _find_marker_blocks,
        _fixed_blocks,
        _resolve_marker_col,
        _u,
    )

    fn_label = _FN
    sheet_name = sheet.Name if sheet is not None else u""
    if _lm_pp_param_skip_sheet(fn_label, extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(fn_label, extra_args, doc, sheet)
    if block is None and extra_args and isinstance(extra_args[0], dict):
        block = extra_args[0]
    if block is None:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нет блока JSON для листа"
        )
        return

    try:
        from libre_macros_param_codec import _normalize_add_merge_block

        block = _normalize_add_merge_block(block)
    except Exception:
        pass

    try:
        from libre_macros_value_join_lib import resolve_value_join_delimiter

        sep = resolve_value_join_delimiter(
            block.get(u"sep") if block.get(u"sep") is not None else u" ",
            default=u" ",
        )
    except Exception:
        sep = _u(block.get(u"sep") if block.get(u"sep") is not None else u" ")
        # Fallback: \\n / \\t / \\r → реальные символы
        sep = (
            sep.replace(u"\\\\", u"\0")
            .replace(u"\\n", u"\n")
            .replace(u"\\r", u"\n")
            .replace(u"\\t", u"\t")
            .replace(u"\0", u"\\")
        )

    has_marker = False
    raw_marker = block.get(u"marker")
    if isinstance(raw_marker, (list, tuple)):
        has_marker = any(_u(x).strip() != u"" for x in raw_marker)
    else:
        has_marker = _u(raw_marker or u"").strip() != u""
    has_block_rows = block.get(u"block_rows") is not None

    if has_marker and has_block_rows:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"задайте либо marker, либо block_rows",
        )
        return
    if not has_marker and not has_block_rows:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"нужен marker или block_rows",
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    ua_sc, ua_sr, ua_ec, ua_er = _lm_vlookup_sheet_used_area(sheet)
    if ua_ec >= ua_sc:
        sc = min(int(sc), int(ua_sc))
        ec = max(int(ec), int(ua_ec))
    if ua_er >= ua_sr:
        er = max(int(er), int(ua_er))
    if ec < sc:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нет столбцов"
        )
        return

    cols, cerr = _resolve_columns(block, sheet, header_row_range, sc, ec)
    if cerr:
        _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ошибка", cerr)
        return

    data_r0 = None
    if block.get(u"data_start_row") not in (None, u""):
        try:
            data_r0 = int(block.get(u"data_start_row")) - 1
        except (TypeError, ValueError):
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"data_start_row: целое 1-based",
            )
            return
    if data_r0 is None:
        # После строки заголовка, если data_range начинается с данных
        if sr <= h_sr:
            data_r0 = int(h_sr) + 1
        else:
            data_r0 = int(sr)
        if data_r0 < int(ua_sr) and ua_er >= ua_sr:
            # used может начинаться выше
            pass

    data_r1 = int(er)
    if data_r1 < data_r0:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нет строк данных"
        )
        return

    blocks = []
    if has_marker:
        marker_col, merr = _resolve_marker_col(
            block, sheet, header_row_range, sc, ec
        )
        if merr:
            _lm_log_postprocess(
                doc, sheet_name, u"диапазон", fn_label, u"ошибка", merr
            )
            return
        blocks = _find_marker_blocks(
            sheet, marker_col, data_r0, data_r1, sc, ec
        )
    else:
        try:
            br = int(block.get(u"block_rows"))
        except (TypeError, ValueError):
            br = 0
        if br < 2:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"block_rows: целое ≥ 2",
            )
            return
        blocks = _fixed_blocks(data_r0, data_r1, br)

    merged_n = 0
    bi = 0
    while bi < len(blocks):
        br0, br1 = blocks[bi]
        ci = 0
        while ci < len(cols):
            if _merge_column_block(sheet, cols[ci], br0, br1, sep=sep):
                merged_n += 1
            ci += 1
        bi += 1

    try:
        ctx = lm_pp_active_context()
        _lm_pp_context_refresh(ctx)
    except Exception:
        pass

    note = u"блоков=%d; merge столбцов=%d" % (len(blocks), merged_n)
    _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
