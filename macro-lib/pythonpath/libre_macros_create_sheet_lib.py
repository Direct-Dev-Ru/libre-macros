# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
"""
Создание листа результата с заголовками и опциональной вставкой данных.

Ключ: создать_лист (RANGE + FINAL).
Созданный лист регистрируется в pending → следующие шаги pp/final.
"""
MACRO_VERSION = "3.10.716"
try:
    unicode
except NameError:
    unicode = str

import json
import re

import libre_macros_lib as lm

CREATE_SHEET_FN = u"создать_лист"

# Глобальные блоки без sheet/source_sheet — один раз за проход книги.
_CREATE_SHEET_RAN = set()


def clear_create_sheet_ran():
    global _CREATE_SHEET_RAN
    _CREATE_SHEET_RAN = set()


def _bool_param(block, key, default=False):
    if not isinstance(block, dict) or key not in block:
        return bool(default)
    try:
        return bool(lm.lm_parse_bool_param(block.get(key), default=default))
    except Exception:
        return bool(block.get(key))


def _parse_header_row_1based(block, default=1):
    raw = None
    if isinstance(block, dict):
        raw = block.get("header_row")
    if raw is None or raw == u"":
        return int(default)
    try:
        v = int(raw)
    except (TypeError, ValueError):
        return int(default)
    if v <= 0:
        return 1
    return v


def _block_fingerprint(block):
    try:
        return unicode(json.dumps(block, sort_keys=True, ensure_ascii=False))
    except Exception:
        return unicode(repr(block))


def _list_sheet_names(doc):
    names = []
    try:
        sheets = doc.getSheets()
        n = int(sheets.getCount())
    except Exception:
        return names
    i = 0
    while i < n:
        try:
            names.append(unicode(sheets.getByIndex(i).Name))
        except Exception:
            pass
        i = i + 1
    return names


def _normalize_name_mode(block):
    raw = unicode(
        (block or {}).get("name_mode")
        or (block or {}).get("mode")
        or u"manual"
    ).strip().casefold()
    if raw in (
        u"expr",
        u"lambda",
        u"compute",
        u"вычислить",
        u"лямбда",
        u"expression",
    ):
        return u"expr"
    return u"manual"


def _parse_column_headers(block):
    from libre_macros_param_codec import split_sep_list_respecting_quotes

    raw = None
    if isinstance(block, dict):
        raw = block.get("columns")
        if raw is None:
            raw = block.get("headers") or block.get("header_columns")
    if raw is None or raw == u"" or raw == []:
        return []
    # Кавычки снимаем: в ячейки пишем голые имена; ',' внутри '…' сохраняется.
    return [
        unicode(x)
        for x in split_sep_list_respecting_quotes(raw, keep_quotes=False)
        if unicode(x).strip() != u""
    ]


def resolve_sheet_name(doc, block):
    """
    → (name, err). Имя уже финализировано (_lm_pp_finalize_sheet_name).
    """
    if not isinstance(block, dict):
        return u"", u"блок не dict"
    mode = _normalize_name_mode(block)
    raw_name = u""
    if mode == u"expr":
        expr = unicode(
            block.get("sheet_name_expr")
            or block.get("name_expr")
            or block.get("name_value")
            or u""
        ).strip()
        if expr == u"":
            return u"", u"name_mode=expr: нужен sheet_name_expr"
        try:
            from libre_macros_lambda_column_lib import compile_lambda_column_sheet_pick

            compiled = compile_lambda_column_sheet_pick(expr)
            if compiled is None:
                return u"", u"пустая лямбда имени"
            fn_pick, glo_pick = compiled
            sheets = _list_sheet_names(doc)
            glo_pick[u"_sheets"] = sheets
            result = fn_pick(sheets)
        except Exception as err:
            return u"", u"лямбда имени: %s" % err
        if result is None:
            return u"", u"лямбда имени вернула None"
        if isinstance(result, (list, tuple)):
            if not result:
                return u"", u"лямбда имени вернула пустой список"
            raw_name = unicode(result[0] or u"").strip()
        else:
            raw_name = unicode(result or u"").strip()
        if raw_name == u"":
            return u"", u"лямбда имени вернула пустую строку"
    else:
        raw_name = unicode(
            block.get("sheet_name")
            or block.get("name")
            or block.get("dest_sheet")
            or block.get("name_value")
            or u""
        ).strip()
        if raw_name == u"":
            return u"", u"нужен sheet_name"
        try:
            from libre_macros_lambda_column_lib import (
                expand_source_variables_in_lambda_code,
            )

            raw_name = unicode(
                expand_source_variables_in_lambda_code(
                    raw_name, log_fn_key=CREATE_SHEET_FN
                )
                or u""
            ).strip()
        except Exception:
            pass
        if raw_name == u"":
            return u"", u"пустое имя после подстановки переменных"

    def _taken(cand):
        name = unicode(cand or u"").strip()
        if name == u"":
            return True
        try:
            return bool(doc.getSheets().hasByName(name))
        except Exception:
            return True

    final = lm._lm_pp_finalize_sheet_name(raw_name, is_taken=_taken)
    if final == u"":
        return u"", u"пустое имя после нормализации"
    return final, None


def _write_headers(sheet, header_row_1, columns):
    """Записать имена колонок в строку header_row_1 (1-based). → note."""
    if not columns:
        return u"заголовок пуст"
    hr0 = int(header_row_1) - 1
    if hr0 < 0:
        hr0 = 0
    ci = 0
    while ci < len(columns):
        try:
            cell = sheet.getCellByPosition(ci, hr0)
            cell.setString(unicode(columns[ci]))
        except Exception:
            try:
                cell = sheet.getCellByPosition(ci, hr0)
                cell.String = unicode(columns[ci])
            except Exception as err:
                return u"ошибка заголовка col=%d: %s" % (ci + 1, err)
        ci = ci + 1
    return u"заголовок %d кол. в строке %d" % (len(columns), header_row_1)


_PURE_VAR_RE = re.compile(
    r"^\s*<<\s*Переменные\.([^>]+?)\s*>>\s*$", re.IGNORECASE
)


def _normalize_data_mode(block):
    raw = unicode((block or {}).get("data_mode") or u"").strip().casefold()
    if raw in (u"sheet", u"лист", u"с листа", u"range", u"диапазон"):
        return u"sheet"
    if raw in (
        u"rows",
        u"manual",
        u"вручную",
        u"текст",
        u"text",
        u"values",
        u"строки",
    ):
        return u"rows"
    if raw in (u"none", u"нет", u"empty", u"пусто", u"-"):
        return u"none"
    # Авто: rows / sheet / none
    rows = unicode((block or {}).get("rows") or (block or {}).get("data_rows") or u"").strip()
    if rows != u"":
        return u"rows"
    src = unicode(
        (block or {}).get("source_sheet") or (block or {}).get("from_sheet") or u""
    ).strip()
    rng = unicode(
        (block or {}).get("source_range") or (block or {}).get("source") or u""
    ).strip()
    if src != u"" or rng != u"":
        return u"sheet"
    return u"none"


def _split_rows_by_pipe(text):
    """Сегменты строк по «|» вне одинарных кавычек."""
    s = unicode(text or u"")
    if s.strip() == u"":
        return []
    parts = []
    cur = []
    in_quote = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == u"'":
            if in_quote and i + 1 < len(s) and s[i + 1] == u"'":
                cur.append(u"'")
                cur.append(u"'")
                i = i + 2
                continue
            in_quote = not in_quote
            cur.append(ch)
            i = i + 1
            continue
        if ch == u"|" and not in_quote:
            parts.append(u"".join(cur).strip())
            cur = []
            i = i + 1
            continue
        cur.append(ch)
        i = i + 1
    parts.append(u"".join(cur).strip())
    return [p for p in parts if p != u""]


def _lookup_variable_raw(query):
    """
    Сырое значение переменной из карты (без разворота списка).
    → (text, err_or_None).
    """
    q = unicode(query or u"").strip()
    if q == u"":
        return u"", u"пустой запрос"
    # Переиспользуем expander на уникальном маркере — но нужен сырой lookup.
    try:
        import libre_macros_collect_cfg as _cw_cfg

        vm = getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_MAP", {}) or {}
        order = getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_ORDER", []) or []
    except Exception:
        vm = {}
        order = []
    ctx = lm.lm_pp_active_context() or {}
    if ctx.get(u"source_variables_map"):
        vm = ctx.get(u"source_variables_map") or vm
    if ctx.get(u"source_variables_order"):
        order = ctx.get(u"source_variables_order") or order
    if not isinstance(vm, dict) or not vm:
        return u"", u"карта переменных пуста"

    # Простой путь: имя без уточнений.
    name = q
    if u"(" in q and q.endswith(u")"):
        # Форма Name(лист, файл) — отдаём expander'у целиком через template.
        try:
            marker = u"<<Переменные.%s>>" % q
            expanded = lm._lm_pp_expand_source_variables_template(
                marker, for_formula=False, log_fn_key=CREATE_SHEET_FN
            )
            return unicode(expanded or u""), None
        except Exception as err:
            return u"", unicode(err)

    name_cf = name.casefold()
    matches = []
    for key in vm:
        segs = [unicode(p or u"").strip() for p in unicode(key or u"").split(u"~")]
        if not segs:
            continue
        if segs[0].casefold() != name_cf:
            continue
        matches.append(key)
    if not matches:
        return u"", u"Переменные.%s: не найдено" % name

    order_pos = {}
    oi = 0
    while oi < len(order):
        order_pos[unicode(order[oi])] = oi
        oi = oi + 1
    best = matches[0]
    bi = order_pos.get(unicode(best), -1)
    mi = 1
    while mi < len(matches):
        k = matches[mi]
        ki = order_pos.get(unicode(k), -1)
        if ki >= bi:
            best = k
            bi = ki
        mi = mi + 1
    val = vm.get(best)
    if isinstance(val, (list, tuple)) and len(val) > 0:
        return unicode(val[0] or u""), None
    return unicode(val or u""), None


def _split_variable_list_text(text):
    """Список из карты (v1: через ';', на вводе также ',') → части."""
    try:
        from libre_macros_manual_var_dialog_lib import _normalize_list_text

        norm = _normalize_list_text(text)
    except Exception:
        norm = unicode(text or u"").replace(u",", u";")
    if norm == u"":
        return []
    return [p.strip() for p in norm.split(u";") if p.strip() != u""]


def _expand_cell_token(token):
    """
    Токен ячейки → ('lit', str) | ('list', [str,…]).
    В '…' — литерал без подстановки. Чистый <<Переменные.X>> со списком → list.
    """
    from libre_macros_param_codec import unwrap_column_name_token

    raw = unicode(token if token is not None else u"")
    bare, quoted = unwrap_column_name_token(raw.strip())
    if quoted:
        return (u"lit", bare)
    m = _PURE_VAR_RE.match(bare)
    if m is not None:
        query = unicode(m.group(1) or u"").strip()
        txt, err = _lookup_variable_raw(query)
        if err:
            # Оставляем плейсхолдер как текст — проще отлаживать.
            return (u"lit", bare)
        parts = _split_variable_list_text(txt)
        if len(parts) > 1:
            return (u"list", parts)
        if len(parts) == 1:
            return (u"lit", parts[0])
        return (u"lit", txt)
    # Скалярная подстановка плейсхолдеров внутри текста.
    try:
        expanded = lm._lm_pp_expand_source_variables_template(
            bare, for_formula=False, log_fn_key=CREATE_SHEET_FN
        )
        return (u"lit", unicode(expanded if expanded is not None else bare))
    except Exception:
        return (u"lit", bare)


def build_rows_matrix(rows_text):
    """
    Текст rows → матрица list[list[str]].

    Строки — через «|»; ячейки — через «,»/«;» с учётом '…'.
    Чистый <<Переменные.список>> разворачивается в столбец (число строк = длина списка).
    """
    from libre_macros_param_codec import split_sep_list_respecting_quotes

    segments = _split_rows_by_pipe(rows_text)
    matrix = []
    for seg in segments:
        cells = split_sep_list_respecting_quotes(seg, keep_quotes=True)
        if not cells:
            continue
        specs = []
        for cell in cells:
            specs.append(_expand_cell_token(cell))
        list_lens = [len(s[1]) for s in specs if s[0] == u"list"]
        if list_lens:
            n = max(list_lens)
            ri = 0
            while ri < n:
                row = []
                for kind, val in specs:
                    if kind == u"list":
                        row.append(val[ri] if ri < len(val) else u"")
                    else:
                        row.append(val)
                matrix.append(row)
                ri = ri + 1
        else:
            matrix.append([val for _k, val in specs])
    return matrix


def _write_rows_matrix(sheet, header_row_1, matrix):
    """Записать матрицу со строки данных. → (ok, note)."""
    if not matrix:
        return True, u"вручную: 0 строк"
    data_sr = int(header_row_1)
    if data_sr < 1:
        data_sr = 1
    # data_sr — 1-based номер первой строки данных = header_row+1
    # header_row_1 is header; data starts at header_row_1+1
    start0 = int(header_row_1)  # 0-based index of first data row = header_row_1 (1-based header → next)
    # header_row_1=1 → header at row 0, data at row 1 → start0 = 1 = header_row_1
    # Actually: header at (header_row_1 - 1), data at header_row_1 (0-based) = header_row_1 (1-based number of next row) - wait
    # header_row_1 = 1 means row index 0 is header, data starts at index 1.
    start0 = int(header_row_1)  # since header is 1-based, next row 0-based index = header_row_1
    # If header_row_1=1, start0 should be 1. Yes start0 = header_row_1 works: 1-based header number equals 0-based data start index.
    nrows = len(matrix)
    ncols = 0
    for row in matrix:
        if len(row) > ncols:
            ncols = len(row)
    if ncols <= 0:
        return True, u"вручную: пустые строки"
    # Выровнять ширину
    aligned = []
    for row in matrix:
        r = [unicode(c if c is not None else u"") for c in row]
        while len(r) < ncols:
            r.append(u"")
        aligned.append(tuple(r))
    try:
        rng = sheet.getCellRangeByPosition(0, start0, ncols - 1, start0 + nrows - 1)
        rng.setDataArray(tuple(aligned))
    except Exception:
        # Fallback по ячейкам
        ri = 0
        while ri < nrows:
            ci = 0
            while ci < ncols:
                try:
                    cell = sheet.getCellByPosition(ci, start0 + ri)
                    cell.setString(unicode(aligned[ri][ci]))
                except Exception:
                    try:
                        cell = sheet.getCellByPosition(ci, start0 + ri)
                        cell.String = unicode(aligned[ri][ci])
                    except Exception as err:
                        return False, u"запись строки %d: %s" % (ri + 1, err)
                ci = ci + 1
            ri = ri + 1
    return True, u"вручную: %d×%d с строки %d" % (nrows, ncols, start0 + 1)


def _fill_manual_rows(doc, dest_sheet, block, header_row_1):
    _unused = doc
    rows_text = unicode(
        block.get("rows") or block.get("data_rows") or block.get("manual_rows") or u""
    ).strip()
    if rows_text == u"":
        return True, u"вручную: пусто"
    try:
        matrix = build_rows_matrix(rows_text)
    except Exception as err:
        return False, u"разбор rows: %s" % err
    return _write_rows_matrix(dest_sheet, header_row_1, matrix)


def _copy_source_data(doc, dest_sheet, block, header_row_1):
    """
    Вставка источника начиная со строки данных (header_row+1).
    → (ok, note). Если источник не задан — (True, «без источника»).
    """
    mode = _normalize_data_mode(block)
    if mode == u"none":
        return True, u"без данных"
    if mode == u"rows":
        return _fill_manual_rows(doc, dest_sheet, block, header_row_1)

    src_sheet_name = unicode(
        block.get("source_sheet") or block.get("from_sheet") or u""
    ).strip()
    src_range = unicode(
        block.get("source_range") or block.get("source") or u""
    ).strip()
    if src_sheet_name == u"" and src_range == u"":
        return True, u"без источника"
    if src_sheet_name == u"":
        return False, u"указан source_range без source_sheet"
    if src_range == u"":
        return False, u"указан source_sheet без source_range"

    import libre_macros_copy_ranges_lib as cr

    values_only = _bool_param(block, "values_only", True)
    content = u"values" if values_only else u"formulas"
    proxy = {
        "source_sheet": src_sheet_name,
        "source_range": src_range,
        "header_row": header_row_1,
    }
    src_sh, sc, sr, ec, er, err = cr.resolve_source_rect(doc, proxy)
    if err:
        return False, err
    src_rect = (sc, sr, ec, er)
    payload, rerr = cr.read_block(doc, src_sh, src_rect, content)
    if rerr:
        return False, rerr
    data_row_1 = int(header_row_1) + 1
    dest_cell = u"A%d" % data_row_1
    anchor = lm._lm_pp_parse_cell_a1_ref(dest_cell)
    if anchor is None:
        return False, u"якорь данных %s" % dest_cell
    dsc0, dsr0 = anchor
    ok_w, note_w = cr.write_block(
        doc,
        src_sh,
        src_rect,
        dest_sheet,
        dsc0,
        dsr0,
        payload,
        content,
        with_formatting=False,
        allow_clipboard=True,
    )
    if not ok_w and content == u"values":
        ok_w, note_w = cr._write_values(dest_sheet, dsc0, dsr0, payload)
    if not ok_w and content == u"formulas":
        ok_w, note_w = cr._write_formulas_matrix(doc, dest_sheet, dsc0, dsr0, payload)
    if not ok_w:
        return False, note_w or u"вставка данных не удалась"
    h = er - sr + 1
    w = ec - sc + 1
    return True, u"данные %s!%s → %s (%d×%d, %s)" % (
        unicode(getattr(src_sh, "Name", src_sheet_name) or src_sheet_name),
        src_range,
        dest_cell,
        h,
        w,
        content,
    )


def execute_create_sheet_block(doc, block):
    """Один JSON-блок. → (ok, note, dest_name)."""
    if not isinstance(block, dict):
        return False, u"блок не dict", u""

    name, nerr = resolve_sheet_name(doc, block)
    if nerr:
        return False, nerr, u""

    existing = lm._lm_pp_get_sheet_ref(doc, name)
    if existing is not None:
        return False, u"лист «%s» уже есть" % name, u""

    try:
        pos = int(doc.Sheets.getCount())
        doc.Sheets.insertNewByName(unicode(name), pos)
    except Exception as err:
        return False, u"создание «%s»: %s" % (name, err), u""

    dest = lm._lm_pp_get_sheet_ref(doc, name)
    if dest is None:
        return False, u"не удалось открыть созданный «%s»" % name, u""
    actual = unicode(getattr(dest, "Name", name) or name)

    header_row_1 = _parse_header_row_1based(block, default=1)
    columns = _parse_column_headers(block)
    hdr_note = _write_headers(dest, header_row_1, columns)

    ok_data, data_note = _copy_source_data(doc, dest, block, header_row_1)
    if not ok_data:
        # Лист уже создан — считаем частичным успехом, но статус ошибки.
        involve_note = u""
        try:
            lm.lm_pp_copy_sheet_register_pending(
                actual,
                int(header_row_1) - 1,
                source_sheet_name=unicode(
                    block.get("source_sheet") or u""
                ).strip(),
            )
            involve_note = u"; pending"
        except Exception as err:
            involve_note = u"; pending fail:%s" % err
        return (
            False,
            u"создан «%s»; %s; данные: %s%s" % (actual, hdr_note, data_note, involve_note),
            actual,
        )

    try:
        lm.lm_pp_copy_sheet_register_pending(
            actual,
            int(header_row_1) - 1,
            source_sheet_name=unicode(block.get("source_sheet") or u"").strip(),
        )
        involve_note = u"pending"
    except Exception as err:
        involve_note = u"pending fail:%s" % err

    try:
        lm._lm_pp_activate_sheet_by_name(doc, actual)
    except Exception:
        pass

    summary = u"создан «%s»; %s; %s; %s" % (
        actual,
        hdr_note,
        data_note,
        involve_note,
    )
    return True, summary, actual


def run_create_sheet_blocks(doc, blocks):
    """Список блоков. → (status, note)."""
    if not blocks:
        return u"пропуск", u"пустая C"
    ok_n = 0
    err_n = 0
    notes = []
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        ok, note, _dest = execute_create_sheet_block(doc, block)
        if ok:
            ok_n = ok_n + 1
        else:
            err_n = err_n + 1
        notes.append(note)
        bi = bi + 1
    if err_n > 0 and ok_n == 0:
        status = u"ошибка"
    elif ok_n > 0:
        status = u"ok"
    else:
        status = u"пропуск"
    return status, u"; ".join(notes)


def _filter_blocks_for_host(doc, sheet, blocks):
    """
    Контекстные блоки — как есть.
    Глобальные — один раз за проход; если задан source_sheet — на листе-источнике.
    """
    filtered = []
    for b in blocks:
        if not isinstance(b, dict):
            continue
        has_ctx = bool(lm._lm_pp_block_target_sheet_names(b))
        if has_ctx:
            filtered.append(b)
            continue
        src = unicode(
            b.get("source_sheet") or b.get("from_sheet") or u""
        ).strip()
        if src != u"":
            if sheet is None:
                continue
            try:
                host = unicode(sheet.Name)
            except Exception:
                host = u""
            if host.casefold() != src.casefold():
                continue
        fp = _block_fingerprint(b)
        if fp in _CREATE_SHEET_RAN:
            continue
        filtered.append(b)
        _CREATE_SHEET_RAN.add(fp)
    return filtered


def lm_pp_range_create_sheet(doc, sheet, data_range, header_row_range, *extra_args):
    """RANGE-обёртка создать_лист."""
    _unused = (data_range, header_row_range)
    sheet_name = sheet.Name if sheet is not None else u""
    fn_key = CREATE_SHEET_FN
    if lm._lm_pp_param_skip_sheet(fn_key, extra_args, doc, sheet):
        return
    blocks = lm._lm_pp_blocks_for_sheet(fn_key, extra_args, doc, sheet)
    if not blocks:
        lm._lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_key, u"пропуск", u"нет блока JSON для листа"
        )
        return
    filtered = _filter_blocks_for_host(doc, sheet, blocks)
    if not filtered:
        return
    status, note = run_create_sheet_blocks(doc, filtered)
    lm._lm_log_postprocess(doc, sheet_name, u"диапазон", fn_key, status, note)
