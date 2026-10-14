# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.769"
"""
убрать_объединение — схлопывание хаотичных merge (Word→Calc) в одну физ. строку.

Карта: lm_pp_range_unmerge_collapse / merge_pp_range_unmerge_collapse.
"""

try:
    unicode
except NameError:
    unicode = str


_FN = u"убрать_объединение"


def _u(v):
    return unicode(v if v is not None else u"")


def _empty_text(s):
    return _u(s).strip() == u""


def _cell_string(sheet, c, r):
    try:
        from libre_macros_lib import cell_text

        return cell_text(sheet.getCellByPosition(int(c), int(r)))
    except Exception:
        try:
            return _u(sheet.getCellByPosition(int(c), int(r)).String)
        except Exception:
            return u""


def _set_cell_string(sheet, c, r, text):
    try:
        sheet.getCellByPosition(int(c), int(r)).String = _u(text)
    except Exception:
        pass


def _join_tokens(tokens, sep):
    out = []
    i = 0
    while i < len(tokens):
        t = _u(tokens[i]).strip()
        if t != u"":
            out.append(t)
        i += 1
    return _u(sep).join(out)


def _merge_owner_map(merges, r0, r1, c0, c1):
    """
    (c, r) → (sc, sr, ec, er) для ячеек внутри rect, попавших в merge.
    """
    owners = {}
    for m in merges or ():
        sc, sr, ec, er = m
        rr = max(int(sr), int(r0))
        while rr <= min(int(er), int(r1)):
            cc = max(int(sc), int(c0))
            while cc <= min(int(ec), int(c1)):
                owners[(cc, rr)] = (int(sc), int(sr), int(ec), int(er))
                cc += 1
            rr += 1
    return owners


def _raw_grid(sheet, r0, r1, c0, c1, owners):
    """
    Сетка значений rect: значение только в TL merge; covered → "".
    """
    nrows = int(r1) - int(r0) + 1
    ncols = int(c1) - int(c0) + 1
    grid = []
    ri = 0
    while ri < nrows:
        row = []
        ci = 0
        while ci < ncols:
            c = int(c0) + ci
            r = int(r0) + ri
            owner = owners.get((c, r))
            if owner is None:
                row.append(_cell_string(sheet, c, r))
            else:
                sc, sr, _ec, _er = owner
                if c == sc and r == sr:
                    row.append(_cell_string(sheet, sc, sr))
                else:
                    row.append(u"")
            ci += 1
        grid.append(row)
        ri += 1
    return grid


def _collapse_horizontal_in_grid(grid, owners, r0, c0, c1):
    """
    В каждой строке: горизонтальные merge → склейка влево (rel), правые очистить.
    """
    if not grid:
        return grid
    nrows = len(grid)
    ncols = int(c1) - int(c0) + 1
    done = set()
    ri = 0
    while ri < nrows:
        r = int(r0) + ri
        ci = 0
        while ci < ncols:
            c = int(c0) + ci
            owner = owners.get((c, r))
            if owner is None:
                ci += 1
                continue
            sc, sr, ec, er = owner
            key = (sc, sr, ec, er, r)
            if key in done:
                ci += 1
                continue
            done.add(key)
            if int(ec) <= int(sc):
                ci += 1
                continue
            # Значения слева направо по столбцам merge в этой физ. строке
            # (для 2D merge значение обычно только в TL-строке).
            tokens = []
            cc = int(sc)
            while cc <= int(ec):
                if cc < int(c0) or cc > int(c1):
                    cc += 1
                    continue
                rel_c = cc - int(c0)
                if 0 <= rel_c < ncols:
                    tok = grid[ri][rel_c]
                    if not _empty_text(tok):
                        tokens.append(tok)
                cc += 1
            left_rel = int(sc) - int(c0)
            if 0 <= left_rel < ncols:
                # sep подставим позже при финальной склейке столбца;
                # здесь временно храним list-маркер через спец. join позже —
                # оставляем токены как один промежуточный список в left,
                # остальные чистим. Для простоты: join с \x1e, разберём снаружи.
                grid[ri][left_rel] = tokens
                cc = int(sc) + 1
                while cc <= int(ec):
                    if int(c0) <= cc <= int(c1):
                        grid[ri][cc - int(c0)] = u""
                    cc += 1
            ci += 1
        ri += 1
    # Нормализовать ячейки: list → остаётся list; str без изменений
    return grid


def _normalize_cell_token(val):
    if isinstance(val, (list, tuple)):
        return [ _u(x).strip() for x in val if not _empty_text(x) ]
    s = _u(val).strip()
    if s == u"":
        return []
    return [s]


def collapse_rect_values(grid, sep):
    """
    grid[nrows][ncols] после гориз. схлопывания → одна строка значений.
    """
    if not grid:
        return []
    ncols = len(grid[0]) if grid else 0
    out = []
    c = 0
    while c < ncols:
        tokens = []
        r = 0
        while r < len(grid):
            cell_tokens = _normalize_cell_token(grid[r][c] if c < len(grid[r]) else u"")
            # Горизонтальные токены в одной ячейке — склеить sep сразу
            if len(cell_tokens) > 1:
                joined = _join_tokens(cell_tokens, sep)
                if joined != u"":
                    tokens.append(joined)
            elif len(cell_tokens) == 1:
                tokens.append(cell_tokens[0])
            r += 1
        out.append(_join_tokens(tokens, sep))
        c += 1
    return out


def unmerge_rect(sheet, merges):
    """Снять merge снизу вверх (как apply_flattened_names_uno)."""
    i = len(merges) - 1
    while i >= 0:
        sc, sr, ec, er = merges[i]
        try:
            rng = sheet.getCellRangeByPosition(int(sc), int(sr), int(ec), int(er))
            if rng.IsMerged:
                rng.merge(False)
        except Exception:
            pass
        i -= 1


def collapse_rect(sheet, r0, r1, c0, c1, sep=u" "):
    """
    Схлопнуть прямоугольник: гориз+верт склейка → верхняя строка,
    unmerge, удалить нижние строки. Возвращает число удалённых строк.
    """
    r0 = int(r0)
    r1 = int(r1)
    c0 = int(c0)
    c1 = int(c1)
    if r1 < r0 or c1 < c0:
        return 0
    from libre_macros_header_lib import uno_merges_in_rect
    from libre_macros_lib import _lm_sheet_delete_rows

    merges = uno_merges_in_rect(sheet, c0, r0, c1, r1)
    owners = _merge_owner_map(merges, r0, r1, c0, c1)
    grid = _raw_grid(sheet, r0, r1, c0, c1, owners)
    grid = _collapse_horizontal_in_grid(grid, owners, r0, c0, c1)
    top_row = collapse_rect_values(grid, sep)

    unmerge_rect(sheet, merges)

    # Записать верхнюю строку, очистить нижние
    c = c0
    while c <= c1:
        rel = c - c0
        val = top_row[rel] if rel < len(top_row) else u""
        _set_cell_string(sheet, c, r0, val)
        r = r0 + 1
        while r <= r1:
            _set_cell_string(sheet, c, r, u"")
            r += 1
        c += 1

    deleted = 0
    if r1 > r0:
        count = r1 - r0
        _lm_sheet_delete_rows(sheet, r0 + 1, count)
        deleted = count
    return deleted


def _resolve_anchor_rows(sheet, anchor_text):
    """
    header_anchor → (r0, r1) 0-based или (None, err).
    Одиночная ячейка расширяется до merge.
    """
    text = _u(anchor_text).strip()
    if text == u"":
        return None, u"header_anchor пуст"
    try:
        from libre_macros_copy_ranges_lib import parse_source_range_a1
    except Exception:
        return None, u"parse_source_range_a1 недоступен"
    rect, err = parse_source_range_a1(sheet, text)
    if err:
        return None, err
    sc, sr, ec, er = rect
    if sc == ec and sr == er:
        try:
            cell = sheet.getCellByPosition(int(sc), int(sr))
            if bool(cell.IsMerged):
                cursor = sheet.createCursorByRange(cell)
                cursor.collapseToMergedArea()
                addr = cursor.getRangeAddress()
                sr = int(addr.StartRow)
                er = int(addr.EndRow)
        except Exception:
            pass
    if er < sr:
        sr, er = er, sr
    return (int(sr), int(er)), None


def _find_marker_blocks(sheet, marker_col, data_r0, data_r1, c0, c1):
    """
    Блоки данных по merge в столбце-маркере (высота ≥ 2).
    Возвращает list[(r0, r1)] сверху вниз.
    """
    from libre_macros_header_lib import uno_merges_in_rect

    merges = uno_merges_in_rect(sheet, int(marker_col), int(data_r0), int(marker_col), int(data_r1))
    blocks = []
    seen_rows = set()
    for sc, sr, ec, er in merges:
        if int(sc) != int(marker_col):
            continue
        if int(er) <= int(sr):
            continue
        if int(sr) < int(data_r0) or int(sr) > int(data_r1):
            continue
        top = max(int(sr), int(data_r0))
        bot = min(int(er), int(data_r1))
        if bot <= top:
            continue
        if top in seen_rows:
            continue
        seen_rows.add(top)
        blocks.append((top, bot))
    blocks.sort(key=lambda b: b[0])
    return blocks


def _fixed_blocks(data_r0, data_r1, block_rows):
    """Нарезать [data_r0..data_r1] на сегменты высоты N; хвост ≥2 тоже."""
    n = int(block_rows)
    if n < 2:
        return []
    out = []
    r = int(data_r0)
    end = int(data_r1)
    while r <= end:
        bot = min(r + n - 1, end)
        if bot > r:
            out.append((r, bot))
        r = bot + 1
    return out


def _resolve_marker_col(block, sheet, header_row_range, sc, ec):
    raw = block.get(u"marker")
    if isinstance(raw, (list, tuple)):
        raw = raw[0] if raw else u""
    token = _u(raw).strip()
    if token == u"":
        return None, u"marker пуст"
    from libre_macros_lib import _lm_pp_resolve_column_token

    col = _lm_pp_resolve_column_token(token, sheet, header_row_range, sc, ec)
    if col is None:
        # Буква/индекс без заголовка
        try:
            from libre_macros_lib import col_letters_to_index, is_col_letters

            if is_col_letters(token):
                col = col_letters_to_index(token)
            else:
                col = int(token) - 1
        except Exception:
            col = None
    if col is None or col < 0:
        return None, u"столбец-маркер не найден: %s" % token
    return int(col), None


def lm_pp_range_unmerge_collapse(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Схлопнуть вертикальные (и горизонтальные) объединения Word→Calc.

    JSON:
      [{"v":1,"fn":"убрать_объединение","header_rows":3,"marker":"A","sep":" "}]
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
        from libre_macros_param_codec import _normalize_unmerge_collapse_block

        block = _normalize_unmerge_collapse_block(block)
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
        sep = (
            sep.replace(u"\\\\", u"\0")
            .replace(u"\\n", u"\n")
            .replace(u"\\r", u"\n")
            .replace(u"\\t", u"\t")
            .replace(u"\0", u"\\")
        )

    has_header_rows = u"header_rows" in block and block.get(u"header_rows") is not None
    has_header_anchor = _u(block.get(u"header_anchor") or u"").strip() != u""
    has_marker = False
    raw_marker = block.get(u"marker")
    if isinstance(raw_marker, (list, tuple)):
        has_marker = any(_u(x).strip() != u"" for x in raw_marker)
    else:
        has_marker = _u(raw_marker or u"").strip() != u""
    has_block_rows = block.get(u"block_rows") is not None

    if has_header_rows and has_header_anchor:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"задайте либо header_rows, либо header_anchor",
        )
        return
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
    if not (has_header_rows or has_header_anchor or has_marker or has_block_rows):
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"нужен заголовок и/или режим данных (marker|block_rows)",
        )
        return

    sc, _sr, ec, er = _lm_pp_range_address(data_range)
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

    deleted_total = 0
    header_bottom = -1  # последняя строка заголовка до схлопывания
    do_header = False
    hdr_r0 = 0
    hdr_r1 = -1

    if has_header_anchor:
        rows, aerr = _resolve_anchor_rows(sheet, block.get(u"header_anchor"))
        if aerr:
            _lm_log_postprocess(
                doc, sheet_name, u"диапазон", fn_label, u"ошибка", aerr
            )
            return
        hdr_r0, hdr_r1 = rows
        do_header = hdr_r1 >= hdr_r0
    elif has_header_rows:
        try:
            n = int(block.get(u"header_rows"))
        except (TypeError, ValueError):
            n = -1
        if n < 0:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"header_rows: целое ≥ 0",
            )
            return
        if n == 0:
            do_header = False
        else:
            hdr_r0 = int(ua_sr) if ua_er >= ua_sr else 0
            hdr_r1 = hdr_r0 + n - 1
            do_header = True

    if do_header:
        header_bottom = hdr_r1
        deleted = collapse_rect(sheet, hdr_r0, hdr_r1, sc, ec, sep=sep)
        deleted_total += deleted
        # После схлопывания заголовок занимает 1 строку
        header_bottom = hdr_r0

    # Зона данных
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
        if do_header:
            data_r0 = int(header_bottom) + 1
        else:
            data_r0 = int(ua_sr) if ua_er >= ua_sr else 0

    # Перечитать used после удаления строк заголовка
    _ua_sc2, _ua_sr2, ua_ec2, ua_er2 = _lm_vlookup_sheet_used_area(sheet)
    if ua_ec2 >= sc:
        ec = max(int(ec), int(ua_ec2))
    data_r1 = int(ua_er2) if ua_er2 >= data_r0 else int(data_r0) - 1

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
        if data_r1 >= data_r0:
            blocks = _find_marker_blocks(
                sheet, marker_col, data_r0, data_r1, sc, ec
            )
    elif has_block_rows:
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
        if data_r1 >= data_r0:
            blocks = _fixed_blocks(data_r0, data_r1, br)

    # Снизу вверх
    bi = len(blocks) - 1
    while bi >= 0:
        br0, br1 = blocks[bi]
        deleted = collapse_rect(sheet, br0, br1, sc, ec, sep=sep)
        deleted_total += deleted
        bi -= 1

    try:
        ctx = lm_pp_active_context()
        _lm_pp_context_refresh(ctx)
    except Exception:
        pass

    note = u"удалено строк=%d; блоков данных=%d" % (deleted_total, len(blocks))
    if do_header:
        note = u"заголовок схлопнут; " + note
    _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
