# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.727"
"""
транспонировать_таблицу — строки ↔ столбцы.

Карта: lm_pp_range_transpose_table / merge_final_transpose_table.
"""

try:
    unicode
except NameError:
    unicode = str

_MAX_CELLS = 5000000


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


def transpose_matrix(matrix):
    """Обычное транспонирование: R×C → C×R. Пустой вход → []."""
    if not matrix:
        return []
    rows = [list(r) for r in matrix]
    n_r = len(rows)
    n_c = max(len(r) for r in rows) if rows else 0
    for r in rows:
        while len(r) < n_c:
            r.append(u"")
        if len(r) > n_c:
            del r[n_c:]
    out = []
    c = 0
    while c < n_c:
        col = []
        r = 0
        while r < n_r:
            col.append(rows[r][c])
            r += 1
        out.append(col)
        c += 1
    return out


def build_headers_from_column_matrix(matrix, header_col_rel=0):
    """
    §2.4: отбросить строку 0; метки = колонка k по строкам данных;
    тело = остальные столбцы, строки 1.. → транспонировать.

    Возвращает (out_matrix, err).
    """
    if not matrix or len(matrix) < 2:
        return None, u"headers_from_column: нужно ≥2 строк"
    rows = [list(r) for r in matrix]
    n_r = len(rows)
    n_c = max(len(r) for r in rows)
    if n_c < 2:
        return None, u"headers_from_column: нужно ≥2 столбцов"
    for r in rows:
        while len(r) < n_c:
            r.append(u"")
        if len(r) > n_c:
            del r[n_c:]
    k = int(header_col_rel or 0)
    if k < 0 or k >= n_c:
        return None, u"header_column вне диапазона"
    new_header = []
    r = 1
    while r < n_r:
        new_header.append(rows[r][k])
        r += 1
    body_src = []
    r = 1
    while r < n_r:
        row = []
        c = 0
        while c < n_c:
            if c != k:
                row.append(rows[r][c])
            c += 1
        body_src.append(row)
        r += 1
    # body_src: (R-1) × (C-1) → transpose → (C-1) × (R-1)
    body_t = transpose_matrix(body_src)
    return [new_header] + body_t, u""


def apply_skip_flags(matrix, skip_header_row=False, skip_first_column=False):
    """Вырезать первую строку и/или первый столбец до транспонирования."""
    if not matrix:
        return []
    rows = [list(r) for r in matrix]
    if skip_header_row:
        rows = rows[1:]
    if skip_first_column and rows:
        rows = [r[1:] for r in rows]
    return rows


def parse_result_headers(raw):
    """
    Список имён столбцов результата из JSON/текста.
    Пусто → [].
    """
    names = []
    if raw is None or raw == u"":
        return names
    if isinstance(raw, (list, tuple)):
        for item in raw:
            s = _u(item).strip()
            if s != u"":
                names.append(s)
        return names
    s = _u(raw).strip()
    if s == u"":
        return names
    import re

    for p in re.split(r"[,;]", s):
        p = p.strip()
        if p != u"":
            names.append(p)
    return names


def apply_result_headers(out_matrix, headers):
    """
    Добавить строку заголовков сверху к уже транспонированной матрице.
    Длина подгоняется под ширину результата (pad пустыми / truncate).
    Возвращает (new_matrix, note) — note про длину при несовпадении.
    """
    if not headers:
        return out_matrix, u""
    body = [list(r) for r in (out_matrix or [])]
    n_c = max(len(r) for r in body) if body else len(headers)
    for r in body:
        while len(r) < n_c:
            r.append(u"")
        if len(r) > n_c:
            del r[n_c:]
    hdr = [_u(h) for h in headers]
    note = u""
    if len(hdr) < n_c:
        note = u"result_headers: дополнено пустыми (%d→%d)" % (len(hdr), n_c)
        while len(hdr) < n_c:
            hdr.append(u"")
    elif len(hdr) > n_c:
        note = u"result_headers: обрезано (%d→%d)" % (len(hdr), n_c)
        hdr = hdr[:n_c]
    if n_c == 0:
        n_c = len(hdr)
    return [hdr] + body, note


def validate_transpose_block(block):
    """
    Валидация JSON-блока до UNO.
    Возвращает err string или u"".
    """
    output = _u(block.get(u"output") or u"inplace").casefold()
    if output not in (u"inplace", u"new_sheet", u"offset"):
        output = u"inplace"
    skip_h = bool(block.get(u"skip_header_row"))
    skip_c = bool(block.get(u"skip_first_column"))
    hfc = bool(block.get(u"headers_from_column"))
    if output == u"inplace" and (skip_h or skip_c) and not hfc:
        return (
            u"inplace + skip_header_row/skip_first_column "
            u"без headers_from_column запрещены"
        )
    if hfc and skip_c:
        return u"headers_from_column + skip_first_column запрещены"
    if output == u"new_sheet":
        if _u(block.get(u"dest_sheet") or block.get(u"dest") or u"").strip() == u"":
            return u"output=new_sheet: нужен dest_sheet"
    if output == u"offset":
        if _u(block.get(u"dest_cell") or u"").strip() == u"":
            return u"output=offset: нужен dest_cell"
    return u""


def _parse_a1_range(spec):
    """
    'A1:F20' / 'A1' → (sc, sr, ec, er) 0-based или None.
    """
    from libre_macros_lib import _lm_pp_parse_cell_a1_ref

    raw = _u(spec).strip().replace(u"$", u"")
    if raw == u"":
        return None
    if u":" in raw:
        left, _, right = raw.partition(u":")
        a = _lm_pp_parse_cell_a1_ref(left.strip())
        b = _lm_pp_parse_cell_a1_ref(right.strip())
        if a is None or b is None:
            return None
        sc, sr = a
        ec, er = b
        if sc > ec:
            sc, ec = ec, sc
        if sr > er:
            sr, er = er, sr
        return sc, sr, ec, er
    one = _lm_pp_parse_cell_a1_ref(raw)
    if one is None:
        return None
    return one[0], one[1], one[0], one[1]


def _rects_overlap(a, b):
    """a/b = (sc,sr,ec,er)."""
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def _clear_rect(sheet, sc, sr, ec, er):
    if ec < sc or er < sr:
        return
    try:
        rng = sheet.getCellRangeByPosition(int(sc), int(sr), int(ec), int(er))
        rng.clearContents(1 | 2 | 4 | 8 | 16 | 32 | 64 | 128 | 256 | 512)
    except Exception:
        try:
            r = sr
            while r <= er:
                c = sc
                while c <= ec:
                    cell = sheet.getCellByPosition(int(c), int(r))
                    cell.String = u""
                    c += 1
                r += 1
        except Exception:
            pass


def lm_pp_range_transpose_table(doc, sheet, data_range, header_row_range, *extra_args):
    """RANGE: транспонировать_таблицу."""
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
        _lm_pp_remove_duplicates_write_matrix,
        _lm_pp_resolve_column_token,
        _lm_pp_unique_sheet_name,
        lm_pp_copy_sheet_register_pending,
        col_letters_to_index,
        is_col_letters,
    )

    fn_label = u"транспонировать_таблицу"
    sheet_name = sheet.Name if sheet is not None else u""
    if _lm_pp_param_skip_sheet(fn_label, extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(fn_label, extra_args, doc, sheet)
    if block is None:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нет блока JSON для листа"
        )
        return

    verr = validate_transpose_block(block)
    if verr:
        _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ошибка", verr)
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)

    # Явный range / header_row из блока
    range_spec = _u(block.get(u"range") or u"").strip()
    if range_spec:
        parsed = _parse_a1_range(range_spec)
        if parsed is None:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"range=%s: неверный A1" % range_spec,
            )
            return
        sc, sr, ec, er = parsed
        h_sr = sr
    else:
        # авто: с header_row / data_start; по умолчанию включаем строку заголовка
        header_row = block.get(u"header_row")
        data_start = block.get(u"data_start")
        try:
            if header_row is not None and _u(header_row).strip() != u"":
                h_sr = max(0, int(header_row) - 1)
            if data_start is not None and _u(data_start).strip() != u"":
                sr = max(0, int(data_start) - 1)
            else:
                # транспонируем вместе с заголовком диапазона сбора
                sr = min(int(sr), int(h_sr))
        except Exception:
            sr = min(int(sr), int(h_sr))
        # расширить вниз/вправо по UsedArea
        try:
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(False)
            used_er = int(cursor.getRangeAddress().EndRow)
            used_ec = int(cursor.getRangeAddress().EndColumn)
            if used_er > er:
                er = used_er
            if used_ec > ec:
                ec = used_ec
        except Exception:
            pass

    if er < sr or ec < sc:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"пустой диапазон"
        )
        return

    # columns filter
    col_tokens = block.get(u"columns") or []
    if not isinstance(col_tokens, (list, tuple)):
        col_tokens = [col_tokens] if col_tokens not in (None, u"") else []
    col_abs = []
    if col_tokens:
        hdr_rng = None
        try:
            hdr_rng = sheet.getCellRangeByPosition(int(sc), int(h_sr), int(ec), int(h_sr))
        except Exception:
            hdr_rng = header_row_range
        seen = set()
        for tok in col_tokens:
            col = _lm_pp_resolve_column_token(tok, sheet, hdr_rng, sc, ec)
            if col is None:
                continue
            if sc <= col <= ec and col not in seen:
                seen.add(col)
                col_abs.append(col)
        if not col_abs:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"columns: ни один столбец не найден",
            )
            return

    output = _u(block.get(u"output") or u"inplace").casefold()
    if output not in (u"inplace", u"new_sheet", u"offset"):
        output = u"inplace"
    hfc = bool(block.get(u"headers_from_column"))
    skip_h = bool(block.get(u"skip_header_row"))
    skip_c = bool(block.get(u"skip_first_column"))
    if hfc:
        skip_h = True  # эквивалент

    as_values = True
    if u"as_values" in block:
        as_values = bool(_lm_pp_parse_as_values_flag(block))
    if as_values and doc is not None:
        try:
            mat_rng = sheet.getCellRangeByPosition(int(sc), int(sr), int(ec), int(er))
            _lm_pp_calculate_with_auto(doc, sheet)
            _lm_pp_copy_range_as_values(doc, sheet, mat_rng)
        except Exception:
            pass

    try:
        data = sheet.getCellRangeByPosition(int(sc), int(sr), int(ec), int(er)).getDataArray()
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

    n_cols = int(ec) - int(sc) + 1
    matrix = []
    for row in data:
        r = list(row)
        while len(r) < n_cols:
            r.append(u"")
        matrix.append(r[:n_cols])

    if col_abs:
        rel = [c - sc for c in col_abs]
        matrix = [[row[i] for i in rel] for row in matrix]
        # для overlap/inplace исходный прямоугольник — только выбранные столбцы
        # (минимальный охватывающий)
        sc_src = min(col_abs)
        ec_src = max(col_abs)
    else:
        sc_src, ec_src = sc, ec
    sr_src, er_src = sr, er

    in_r = len(matrix)
    in_c = len(matrix[0]) if matrix else 0
    if in_r * in_c > _MAX_CELLS:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"слишком большой диапазон (%d ячеек > %d)" % (in_r * in_c, _MAX_CELLS),
        )
        return

    # header_column для hfc
    header_col_rel = 0
    if hfc:
        hc = block.get(u"header_column")
        if hc is not None and _u(hc).strip() != u"":
            # индекс в текущей matrix (после columns filter)
            raw = _u(hc).strip()
            resolved = None
            try:
                resolved = int(raw) - 1
            except Exception:
                resolved = None
            if resolved is None and is_col_letters(raw):
                try:
                    abs_c = int(col_letters_to_index(raw))
                    if col_abs:
                        if abs_c in col_abs:
                            resolved = col_abs.index(abs_c)
                    else:
                        resolved = abs_c - sc
                except Exception:
                    resolved = None
            if resolved is None:
                # имя из строки заголовка matrix[0]
                hdr0 = matrix[0] if matrix else []
                i = 0
                while i < len(hdr0):
                    if _u(hdr0[i]).strip().casefold() == raw.casefold():
                        resolved = i
                        break
                    i += 1
            if resolved is None or resolved < 0 or resolved >= in_c:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    u"диапазон",
                    fn_label,
                    u"ошибка",
                    u"header_column=%s не найден в диапазоне" % raw,
                )
                return
            header_col_rel = resolved

    rh_raw = block.get(u"result_headers")
    if rh_raw is None:
        rh_raw = block.get(u"new_headers")
    result_headers = parse_result_headers(rh_raw)
    rh_note = u""

    if hfc:
        if result_headers:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ok",
                u"result_headers игнорируется при headers_from_column",
            )
        out_matrix, herr = build_headers_from_column_matrix(matrix, header_col_rel)
        if herr:
            _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ошибка", herr)
            return
    else:
        trimmed = apply_skip_flags(matrix, skip_h, skip_c)
        if not trimmed or not trimmed[0]:
            _lm_log_postprocess(
                doc, sheet_name, u"диапазон", fn_label, u"пропуск", u"нечего транспонировать"
            )
            return
        out_matrix = transpose_matrix(trimmed)
        if result_headers:
            out_matrix, rh_note = apply_result_headers(out_matrix, result_headers)

    out_r = len(out_matrix)
    out_c = len(out_matrix[0]) if out_matrix else 0
    if out_r * out_c > _MAX_CELLS:
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ошибка",
            u"результат слишком большой (%d)" % (out_r * out_c),
        )
        return

    if bool(block.get(u"with_formatting")):
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ok",
            u"with_formatting: в v1 не переносится, только значения",
        )

    # --- запись ---
    if output == u"new_sheet":
        dest_base = _u(block.get(u"dest_sheet") or block.get(u"dest") or u"").strip()
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
                return
            dest_sh = _lm_pp_get_sheet_ref(doc, actual)
        else:
            _lm_pp_clear_sheet_content(dest_sh)
        ok, werr = _lm_pp_remove_duplicates_write_matrix(doc, dest_sh, 0, 0, out_matrix)
        if not ok:
            _lm_log_postprocess(
                doc, sheet_name, u"диапазон", fn_label, u"ошибка", werr or u"запись"
            )
            return
        if bool(block.get(u"clear_source")):
            _clear_rect(sheet, sc_src, sr_src, ec_src, er_src)
        try:
            lm_pp_copy_sheet_register_pending(actual, 0, sheet_name, from_dedup=False)
        except Exception:
            pass
        note = u"%d×%d → %d×%d на «%s»" % (in_r, in_c, out_r, out_c, actual)
        if hfc:
            note = note + u", headers_from_column"
        elif result_headers:
            note = note + u", result_headers"
            if rh_note:
                note = note + u" (" + rh_note + u")"
        _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
        return

    if output == u"offset":
        dest_cell = _u(block.get(u"dest_cell") or u"").strip()
        anchor = _lm_pp_parse_cell_a1_ref(dest_cell)
        if anchor is None:
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"dest_cell=%s: нужен A1" % dest_cell,
            )
            return
        dc, dr = anchor
        dest_rect = (dc, dr, dc + out_c - 1, dr + out_r - 1)
        src_rect = (sc_src, sr_src, ec_src, er_src)
        if _rects_overlap(src_rect, dest_rect) and not bool(
            block.get(u"overwrite_overlap")
        ):
            _lm_log_postprocess(
                doc,
                sheet_name,
                u"диапазон",
                fn_label,
                u"ошибка",
                u"offset пересекается с источником (overwrite_overlap=false)",
            )
            return
        ok, werr = _lm_pp_remove_duplicates_write_matrix(
            doc, sheet, dc, dr, out_matrix
        )
        if not ok:
            _lm_log_postprocess(
                doc, sheet_name, u"диапазон", fn_label, u"ошибка", werr or u"запись"
            )
            return
        if bool(block.get(u"clear_source")):
            _clear_rect(sheet, sc_src, sr_src, ec_src, er_src)
        note = u"%d×%d → %d×%d @ %s" % (in_r, in_c, out_r, out_c, dest_cell)
        if hfc:
            note = note + u", headers_from_column"
        elif result_headers:
            note = note + u", result_headers"
            if rh_note:
                note = note + u" (" + rh_note + u")"
        _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
        return

    # inplace
    if out_c != (ec_src - sc_src + 1) or out_r != (er_src - sr_src + 1):
        _lm_log_postprocess(
            doc,
            sheet_name,
            u"диапазон",
            fn_label,
            u"ok",
            u"inplace: размер %d×%d → %d×%d (область изменится)"
            % (in_r, in_c, out_r, out_c),
        )
    _clear_rect(sheet, sc_src, sr_src, ec_src, er_src)
    # если новый прямоугольник шире/выше — clear уже покрыл старое; дописать
    new_ec = sc_src + out_c - 1
    new_er = sr_src + out_r - 1
    if new_ec < ec_src or new_er < er_src:
        # хвосты уже очищены _clear_rect
        pass
    ok, werr = _lm_pp_remove_duplicates_write_matrix(
        doc, sheet, sc_src, sr_src, out_matrix
    )
    if not ok:
        _lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_label, u"ошибка", werr or u"запись"
        )
        return
    note = u"%d×%d → %d×%d inplace" % (in_r, in_c, out_r, out_c)
    if hfc:
        note = note + u", headers_from_column"
    elif result_headers:
        note = note + u", result_headers"
        if rh_note:
            note = note + u" (" + rh_note + u")"
    _lm_log_postprocess(doc, sheet_name, u"диапазон", fn_label, u"ok", note)
