# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
"""
Копирование прямоугольника ячеек между листами книги результата.

Ключ: копирование_диапазонов (FINAL + опционально RANGE).
"""
MACRO_VERSION = "3.10.689"
import re

try:
    unicode
except NameError:
    unicode = str

import libre_macros_lib as lm

COPY_RANGES_FN = u"копирование_диапазонов"
_COPY_RANGES_MAX_CELLS = 5000000

_A1_CELL_RE = re.compile(r"^\$?([A-Za-z]+)\$?(\d+)$")
_A1_COL_RE = re.compile(r"^\$?([A-Za-z]+):\$?([A-Za-z]+)$")
_A1_ROW_RE = re.compile(r"^\$?(\d+):\$?(\d+)$")
_A1_RECT_RE = re.compile(
    r"^\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$"
)


def _bool_param(block, key, default=False):
    if not isinstance(block, dict) or key not in block:
        return bool(default)
    try:
        return bool(lm.lm_parse_bool_param(block.get(key), default=default))
    except Exception:
        return bool(block.get(key))


def _parse_header_row_0(block, fallback=0):
    raw = None
    if isinstance(block, dict):
        raw = block.get("header_row")
    if raw is None or raw == u"":
        return int(fallback)
    try:
        v = int(raw)
    except (TypeError, ValueError):
        return int(fallback)
    if v <= 0:
        return 0
    return v - 1


def _a1_of(col0, row0):
    return u"%s%d" % (lm.col_index_to_letters(col0), int(row0) + 1)


def _rect_label(sc, sr, ec, er):
    if sc == ec and sr == er:
        return _a1_of(sc, sr)
    return u"%s:%s" % (_a1_of(sc, sr), _a1_of(ec, er))


def _rects_overlap(a, b):
    sc, sr, ec, er = a
    dsc, dsr, dec, der = b
    if ec < dsc or dec < sc:
        return False
    if er < dsr or der < sr:
        return False
    return True


def _used_area(sheet):
    return lm._lm_vlookup_sheet_used_area(sheet)


def _parse_int_list_continuous(spec, label):
    """
    1-based строки/индексы → непрерывный (lo, hi) включительно.
    Допускает «2:10», «2,5,7» (дыры заполняются), список.
    """
    if spec is None or spec == u"" or spec == []:
        return None, None
    nums = []
    if isinstance(spec, (list, tuple)):
        items = list(spec)
    else:
        text = unicode(spec).strip()
        if text == u"":
            return None, None
        text = text.replace(u";", u",")
        items = []
        for part in text.split(u","):
            part = part.strip()
            if part == u"":
                continue
            if u":" in part:
                a, _, b = part.partition(u":")
                a = a.strip()
                b = b.strip()
                if a.isdigit() and b.isdigit():
                    lo = int(a)
                    hi = int(b)
                    if hi < lo:
                        lo, hi = hi, lo
                    items.extend(range(lo, hi + 1))
                    continue
            items.append(part)
    for item in items:
        try:
            n = int(item)
        except (TypeError, ValueError):
            return None, u"%s: нечисловое значение %r" % (label, item)
        if n < 1:
            return None, u"%s: индекс < 1" % label
        nums.append(n)
    if not nums:
        return None, None
    return (min(nums), max(nums)), None


def _resolve_columns_span(sheet, spec, header_row_0, ua_sc, ua_ec):
    """
    source_columns → непрерывный (sc, ec) 0-based.
    """
    if spec is None or spec == u"" or spec == []:
        return None, None
    if isinstance(spec, (list, tuple)):
        parts = [unicode(x).strip() for x in spec if unicode(x).strip() != u""]
        text = u",".join(parts)
    else:
        text = unicode(spec).strip()
    if text == u"":
        return None, None
    # Буквенный диапазон A:C без кавычек.
    m = _A1_COL_RE.match(text.replace(u"$", u"").replace(u" ", u""))
    if m is not None:
        c0 = lm.col_letters_to_index(m.group(1))
        c1 = lm.col_letters_to_index(m.group(2))
        if c0 < 0 or c1 < 0:
            return None, u"source_columns: некорректный диапазон столбцов"
        if c1 < c0:
            c0, c1 = c1, c0
        return (c0, c1), None
    # Список индексов / букв / заголовков через codec-резолвер.
    try:
        hdr_rng = sheet.getCellRangeByPosition(
            int(ua_sc), int(header_row_0), int(ua_ec), int(header_row_0)
        )
    except Exception:
        hdr_rng = None
    cols = lm._lm_pp_resolve_column_list(
        text, sheet, hdr_rng, int(ua_sc), int(ua_ec)
    )
    if not cols:
        return None, u"source_columns: столбцы не найдены"
    return (min(cols), max(cols)), None


def parse_source_range_a1(sheet, text):
    """
    A1-диапазон → (sc, sr, ec, er) или (None, err).
    Строки/столбцы целиком — в пределах used-area.
    """
    s = unicode(text or u"").strip().replace(u"$", u"").replace(u" ", u"")
    if s == u"":
        return None, u"source_range пуст"
    ua_sc, ua_sr, ua_ec, ua_er = _used_area(sheet)
    m = _A1_RECT_RE.match(s)
    if m is not None:
        sc = lm.col_letters_to_index(m.group(1))
        sr = int(m.group(2)) - 1
        ec = lm.col_letters_to_index(m.group(3))
        er = int(m.group(4)) - 1
        if sc < 0 or ec < 0 or sr < 0 or er < 0:
            return None, u"source_range=%s: некорректный A1" % text
        if ec < sc:
            sc, ec = ec, sc
        if er < sr:
            sr, er = er, sr
        return (sc, sr, ec, er), None
    m = _A1_CELL_RE.match(s)
    if m is not None:
        sc = lm.col_letters_to_index(m.group(1))
        sr = int(m.group(2)) - 1
        if sc < 0 or sr < 0:
            return None, u"source_range=%s: некорректный A1" % text
        return (sc, sr, sc, sr), None
    m = _A1_COL_RE.match(s)
    if m is not None:
        if ua_er < ua_sr:
            return None, u"source_range=%s: used-area пуст" % text
        sc = lm.col_letters_to_index(m.group(1))
        ec = lm.col_letters_to_index(m.group(2))
        if sc < 0 or ec < 0:
            return None, u"source_range=%s: некорректный A1" % text
        if ec < sc:
            sc, ec = ec, sc
        return (sc, ua_sr, ec, ua_er), None
    m = _A1_ROW_RE.match(s)
    if m is not None:
        if ua_ec < ua_sc:
            return None, u"source_range=%s: used-area пуст" % text
        sr = int(m.group(1)) - 1
        er = int(m.group(2)) - 1
        if sr < 0 or er < 0:
            return None, u"source_range=%s: некорректный A1" % text
        if er < sr:
            sr, er = er, sr
        return (ua_sc, sr, ua_ec, er), None
    return None, u"source_range=%s: некорректный A1" % text


def resolve_source_rect(doc, block):
    """
    → (sheet, sc, sr, ec, er, err). err — строка или None.
    """
    if not isinstance(block, dict):
        return None, 0, 0, -1, -1, u"блок не dict"
    source_name = unicode(
        block.get("source_sheet") or block.get("from_sheet") or u""
    ).strip()
    if source_name == u"":
        return None, 0, 0, -1, -1, u"нет source_sheet"
    sheet = lm._lm_pp_get_sheet_ref(doc, source_name)
    if sheet is None:
        return None, 0, 0, -1, -1, u"source_sheet=%s: лист не найден" % source_name

    source_range = unicode(
        block.get("source_range") or block.get("source") or u""
    ).strip()
    if source_range != u"":
        rect, err = parse_source_range_a1(sheet, source_range)
        if err:
            return sheet, 0, 0, -1, -1, err
        sc, sr, ec, er = rect
        n = (ec - sc + 1) * (er - sr + 1)
        if n > _COPY_RANGES_MAX_CELLS:
            return sheet, sc, sr, ec, er, u"слишком большой блок (%d ячеек)" % n
        return sheet, sc, sr, ec, er, None

    ua_sc, ua_sr, ua_ec, ua_er = _used_area(sheet)
    if ua_er < ua_sr or ua_ec < ua_sc:
        return sheet, 0, 0, -1, -1, u"used-area пуст"

    hdr0 = _parse_header_row_0(block, fallback=ua_sr)
    rows_span, rows_err = _parse_int_list_continuous(
        block.get("source_rows"), u"source_rows"
    )
    if rows_err:
        return sheet, 0, 0, -1, -1, rows_err
    cols_span, cols_err = _resolve_columns_span(
        sheet, block.get("source_columns"), hdr0, ua_sc, ua_ec
    )
    if cols_err:
        return sheet, 0, 0, -1, -1, cols_err

    if rows_span is None and cols_span is None:
        return sheet, 0, 0, -1, -1, u"нужен source_range / source_rows / source_columns"

    if rows_span is None:
        sr, er = ua_sr, ua_er
    else:
        sr = rows_span[0] - 1
        er = rows_span[1] - 1
    if cols_span is None:
        sc, ec = ua_sc, ua_ec
    else:
        sc, ec = cols_span

    if er < sr or ec < sc:
        return sheet, 0, 0, -1, -1, u"пустой источник после раскрытия"
    n = (ec - sc + 1) * (er - sr + 1)
    if n > _COPY_RANGES_MAX_CELLS:
        return sheet, sc, sr, ec, er, u"слишком большой блок (%d ячеек)" % n
    return sheet, sc, sr, ec, er, None


def resolve_dest_sheets(doc, block):
    """→ (names_list, err)."""
    if not isinstance(block, dict):
        return [], u"блок не dict"
    names = []
    dest_sheets = block.get("dest_sheets")
    if dest_sheets is None:
        dest_sheets = block.get("to_sheets") or block.get("target_sheets")
    if isinstance(dest_sheets, (list, tuple)):
        for x in dest_sheets:
            s = unicode(x or u"").strip()
            if s != u"":
                names.append(s)
    elif dest_sheets is not None and unicode(dest_sheets).strip() != u"":
        text = unicode(dest_sheets).replace(u";", u",")
        for p in text.split(u","):
            s = p.strip()
            if s != u"":
                names.append(s)
    if not names:
        one = unicode(
            block.get("dest_sheet")
            or block.get("to_sheet")
            or block.get("target_sheet")
            or u""
        ).strip()
        if one != u"":
            names.append(one)
    if not names:
        return [], u"dest_sheets пуст"
    return names, None


def _normalize_mode(block):
    raw = unicode(
        block.get("mode") or block.get("paste_mode") or u"replace"
    ).strip().casefold()
    if raw in (u"insert", u"append", u"сдвиг", u"добавление"):
        return u"insert"
    return u"replace"


def _normalize_content(block):
    if _bool_param(block, "as_values", default=False) or _bool_param(
        block, u"только_значения", default=False
    ):
        return u"values"
    if _bool_param(block, "keep_formulas", default=False) or _bool_param(
        block, u"формулы", default=False
    ):
        return u"formulas"
    raw = unicode(block.get("content") or u"values").strip().casefold()
    if raw in (u"formulas", u"formula", u"формулы", u"keep_formulas"):
        return u"formulas"
    return u"values"


def _normalize_insert_axis(block, height, width, source_only_rows, source_only_cols):
    raw = unicode(block.get("insert_axis") or u"auto").strip().casefold()
    if raw in (u"rows", u"row", u"строки", u"строка"):
        return u"rows"
    if raw in (u"cols", u"columns", u"col", u"столбцы", u"столбец"):
        return u"cols"
    if source_only_rows or height == 1:
        return u"rows"
    if source_only_cols or width == 1:
        return u"cols"
    return u"rows"


def read_block(doc, sheet, rect, content):
    """
    Snapshot источника.
    payload: {content, data, formulas, formula_count, height, width}
    """
    sc, sr, ec, er = rect
    h = er - sr + 1
    w = ec - sc + 1
    payload = {
        "content": content,
        "data": None,
        "formulas": None,
        "formula_count": 0,
        "height": h,
        "width": w,
        "sc": sc,
        "sr": sr,
        "ec": ec,
        "er": er,
    }
    try:
        lm._lm_pp_calculate_with_auto(doc, sheet)
    except Exception:
        pass
    try:
        rng = sheet.getCellRangeByPosition(sc, sr, ec, er)
        payload["data"] = rng.getDataArray()
    except Exception as err:
        return None, u"getDataArray: %s" % err

    if content == u"formulas":
        formulas = []
        fcount = 0
        r = sr
        while r <= er:
            row = []
            c = sc
            while c <= ec:
                cell = sheet.getCellByPosition(c, r)
                ftxt = u""
                try:
                    if lm._lm_pp_cell_is_formula(cell):
                        ftxt = unicode(cell.getFormula() or u"").strip()
                        if ftxt == u"":
                            ftxt = unicode(getattr(cell, "Formula", u"") or u"").strip()
                        if ftxt != u"":
                            fcount = fcount + 1
                except Exception:
                    ftxt = u""
                row.append(ftxt)
                c = c + 1
            formulas.append(tuple(row))
            r = r + 1
        payload["formulas"] = tuple(formulas)
        payload["formula_count"] = fcount
    return payload, None


def prepare_dest(sheet, dest_cell, height, width, mode, insert_axis):
    """→ (sc0, sr0, err). При insert — вставляет строки/столбцы перед якорем."""
    anchor = lm._lm_pp_parse_cell_a1_ref(dest_cell)
    if anchor is None:
        return 0, 0, u"dest_cell=%s: нужен A1 одной ячейки" % dest_cell
    sc0, sr0 = anchor
    if mode != u"insert":
        return sc0, sr0, None
    try:
        if insert_axis == u"cols":
            sheet.getColumns().insertByIndex(int(sc0), int(width))
        else:
            sheet.getRows().insertByIndex(int(sr0), int(height))
    except Exception as err:
        return sc0, sr0, u"insertByIndex: %s" % err
    return sc0, sr0, None


def clear_rect(sheet, rect):
    sc, sr, ec, er = rect
    try:
        rng = sheet.getCellRangeByPosition(sc, sr, ec, er)
        rng.clearContents(7)
    except Exception:
        r = sr
        while r <= er:
            c = sc
            while c <= ec:
                try:
                    sheet.getCellByPosition(c, r).setString(u"")
                except Exception:
                    pass
                c = c + 1
            r = r + 1


def _write_values(sheet, sc0, sr0, payload):
    data = payload.get("data")
    if data is None:
        return False, u"нет data"
    h = int(payload["height"])
    w = int(payload["width"])
    try:
        dest = sheet.getCellRangeByPosition(sc0, sr0, sc0 + w - 1, sr0 + h - 1)
        dest.setDataArray(data)
        return True, u"values"
    except Exception as err:
        return False, u"setDataArray: %s" % err


def _write_formulas_matrix(doc, sheet, sc0, sr0, payload):
    formulas = payload.get("formulas")
    data = payload.get("data")
    h = int(payload["height"])
    w = int(payload["width"])
    if formulas is None:
        return False, u"нет formulas"
    r = 0
    while r < h:
        c = 0
        while c < w:
            cell = sheet.getCellByPosition(sc0 + c, sr0 + r)
            ftxt = u""
            try:
                ftxt = unicode(formulas[r][c] or u"").strip()
            except Exception:
                ftxt = u""
            if ftxt != u"":
                ok = False
                try:
                    ok = lm._lm_pp_try_set_cell_formula(
                        cell, ftxt, use_local=False, allow_formula_error=True
                    )
                except Exception:
                    ok = False
                if not ok:
                    try:
                        lm._lm_pp_set_cell_formula(
                            doc,
                            cell,
                            ftxt,
                            row_0=sr0 + r,
                            col_0=sc0 + c,
                            allow_formula_error=True,
                        )
                    except Exception:
                        pass
            else:
                # Обычное значение из snapshot.
                try:
                    val = data[r][c] if data is not None else u""
                except Exception:
                    val = u""
                try:
                    if isinstance(val, (int, float)) and not isinstance(val, bool):
                        cell.Value = float(val)
                    else:
                        cell.setString(unicode(val if val is not None else u""))
                except Exception:
                    try:
                        cell.setString(unicode(val if val is not None else u""))
                    except Exception:
                        pass
            c = c + 1
        r = r + 1
    return True, u"formulas as text (no ref shift)"


def _clipboard_paste(doc, src_sheet, src_rect, dst_sheet, sc0, sr0, content, with_formatting):
    sc, sr, ec, er = src_rect
    try:
        src_rng = src_sheet.getCellRangeByPosition(sc, sr, ec, er)
    except Exception:
        return False, u"src range"
    if not lm._lm_pp_copy_range_to_clipboard(doc, src_sheet, src_rng):
        return False, u"clipboard copy failed"
    h = er - sr + 1
    w = ec - sc + 1
    try:
        dest_rng = dst_sheet.getCellRangeByPosition(
            sc0, sr0, sc0 + w - 1, sr0 + h - 1
        )
    except Exception:
        return False, u"dest range"
    if with_formatting:
        if lm._lm_pp_clipboard_paste_at_cell(doc, dst_sheet, sc0, sr0):
            return True, u"paste=clipboard+format"
        return False, u"clipboard paste (format) failed"
    if content == u"formulas":
        if lm._lm_pp_paste_cell_range_formulas_only(dst_sheet, dest_rng):
            return True, u"paste=clipboard"
        # InsertContents formulas
        if lm._lm_pp_select_cell_range(doc, dst_sheet, dest_rng):
            if lm._lm_pp_execute_dispatch(
                doc, ".uno:InsertContents", lm._lm_pp_insert_contents_formula_props()
            ):
                return True, u"paste=InsertContents formulas"
        return False, u"clipboard formulas paste failed"
    # values
    if lm._lm_pp_paste_cell_range_values_only(dst_sheet, dest_rng):
        return True, u"paste=clipboard values"
    if lm._lm_pp_select_cell_range(doc, dst_sheet, dest_rng):
        if lm._lm_pp_execute_dispatch(
            doc, ".uno:InsertContents", lm._lm_pp_insert_contents_values_props()
        ):
            return True, u"paste=InsertContents values"
    return False, u"clipboard values paste failed"


def write_block( doc, src_sheet, src_rect, dst_sheet, sc0, sr0, payload, content, with_formatting=False, allow_clipboard=True):
    """→ (ok, note)."""
    notes = []
    # Форматирование / формулы с автосдвигом — через clipboard, если источник жив.
    if allow_clipboard and (with_formatting or content == u"formulas"):
        ok_c, note_c = _clipboard_paste(
            doc,
            src_sheet,
            src_rect,
            dst_sheet,
            sc0,
            sr0,
            content,
            with_formatting,
        )
        if ok_c:
            return True, note_c
        notes.append(note_c)
        if content == u"formulas":
            ok_f, note_f = _write_formulas_matrix(doc, dst_sheet, sc0, sr0, payload)
            if ok_f:
                notes.append(note_f)
                return True, u"; ".join(notes)
            return False, u"; ".join(notes + [note_f])
        # formatting failed → values without format
        notes.append(u"format skipped")
    ok_v, note_v = _write_values(dst_sheet, sc0, sr0, payload)
    if ok_v:
        notes.append(note_v)
        return True, u"; ".join(notes) if notes else note_v
    return False, note_v


def _try_involve_extra(names):
    try:
        import collect_workbooks as cw

        fn = getattr(cw, "_merge_involve_extra_sheet_names", None)
        if callable(fn):
            fn(names)
            return True
    except Exception:
        pass
    return False


def involve_dest_sheet(doc, dest_name, header_row_0, source_name, expanded_cols=0):
    """Регистрация pending + involve. → note."""
    parts = []
    try:
        lm.lm_pp_copy_sheet_register_pending(
            dest_name, int(header_row_0), source_sheet_name=source_name
        )
        parts.append(u"pending")
    except Exception as err:
        parts.append(u"pending fail:%s" % err)
    if _try_involve_extra([dest_name]):
        parts.append(u"involve=%s" % dest_name)
    else:
        parts.append(u"involve=skip")
    if int(expanded_cols or 0) > 0:
        parts.append(u"cols+=%d" % int(expanded_cols))
    return u"; ".join(parts)


def _ensure_dest_sheet(doc, name, create_missing):
    sh = lm._lm_pp_get_sheet_ref(doc, name)
    if sh is not None:
        return sh, unicode(getattr(sh, "Name", name) or name), None
    if not create_missing:
        return None, name, u"лист «%s» не найден" % name
    try:
        pos = int(doc.Sheets.getCount())
        doc.Sheets.insertNewByName(unicode(name), pos)
    except Exception as err:
        return None, name, u"создание «%s»: %s" % (name, err)
    sh = lm._lm_pp_get_sheet_ref(doc, name)
    if sh is None:
        return None, name, u"не удалось создать «%s»" % name
    return sh, unicode(getattr(sh, "Name", name) or name), None


def execute_copy_ranges_block(doc, block):
    """
    Один JSON-блок. → (ok, note).
    """
    if not isinstance(block, dict):
        return False, u"блок не dict"

    src_sheet, sc, sr, ec, er, err = resolve_source_rect(doc, block)
    if err:
        return False, err
    src_rect = (sc, sr, ec, er)
    h = er - sr + 1
    w = ec - sc + 1
    src_name = unicode(getattr(src_sheet, "Name", u"") or u"")

    dest_names, derr = resolve_dest_sheets(doc, block)
    if derr:
        return False, derr

    dest_cell = unicode(
        block.get("dest_cell")
        or block.get("anchor")
        or block.get("at")
        or u"A1"
    ).strip() or u"A1"
    if lm._lm_pp_parse_cell_a1_ref(dest_cell) is None:
        return False, u"dest_cell=%s: нужен A1 одной ячейки" % dest_cell

    mode = _normalize_mode(block)
    content = _normalize_content(block)
    with_formatting = _bool_param(block, "with_formatting", False)
    clear_source = _bool_param(block, "clear_source", False)
    create_missing = _bool_param(block, "create_missing_dest", False)
    involve_dest = _bool_param(block, "involve_dest", True)
    header_row_0 = _parse_header_row_0(block, fallback=0)

    only_rows = bool(
        unicode(block.get("source_rows") or u"").strip()
        or (
            isinstance(block.get("source_rows"), (list, tuple))
            and block.get("source_rows")
        )
    ) and not (
        unicode(block.get("source_columns") or u"").strip()
        or (
            isinstance(block.get("source_columns"), (list, tuple))
            and block.get("source_columns")
        )
    )
    only_cols = bool(
        unicode(block.get("source_columns") or u"").strip()
        or (
            isinstance(block.get("source_columns"), (list, tuple))
            and block.get("source_columns")
        )
    ) and not (
        unicode(block.get("source_rows") or u"").strip()
        or (
            isinstance(block.get("source_rows"), (list, tuple))
            and block.get("source_rows")
        )
    )
    insert_axis = _normalize_insert_axis(block, h, w, only_rows, only_cols)

    # Overlap + clear_source запрет (v1).
    if clear_source:
        for dn in dest_names:
            dsh = lm._lm_pp_get_sheet_ref(doc, dn)
            if dsh is None:
                continue
            if unicode(getattr(dsh, "Name", u"") or u"").casefold() != src_name.casefold():
                continue
            anchor = lm._lm_pp_parse_cell_a1_ref(dest_cell)
            if anchor is None:
                continue
            dsc0, dsr0 = anchor
            # При insert прямоугольник сдвигается — overlap до insert считаем по якорю.
            dest_rect = (dsc0, dsr0, dsc0 + w - 1, dsr0 + h - 1)
            if _rects_overlap(src_rect, dest_rect):
                return False, u"overlap + clear_source запрещены на одном листе"

    payload, rerr = read_block(doc, src_sheet, src_rect, content)
    if rerr:
        return False, rerr

    ok_dests = []
    notes = []
    involve_notes = []
    any_ok = False
    di = 0
    while di < len(dest_names):
        dname = dest_names[di]
        dsh, actual, cerr = _ensure_dest_sheet(doc, dname, create_missing)
        if cerr:
            notes.append(u"%s: %s" % (dname, cerr))
            di = di + 1
            continue
        # Overlap: clipboard с живого источника опасен — пишем из snapshot.
        same_sheet = unicode(getattr(dsh, "Name", u"") or u"").casefold() == src_name.casefold()
        dsc0, dsr0, perr = prepare_dest(
            dsh, dest_cell, h, w, mode, insert_axis
        )
        if perr:
            notes.append(u"%s: %s" % (actual, perr))
            di = di + 1
            continue
        dest_rect = (dsc0, dsr0, dsc0 + w - 1, dsr0 + h - 1)
        overlap = same_sheet and _rects_overlap(src_rect, dest_rect)
        allow_clip = (not overlap) and (not same_sheet or content != u"formulas" or not with_formatting)
        # При overlap всегда из памяти; при formulas+overlap — matrix без сдвига.
        if overlap:
            allow_clip = False
        ok_w, note_w = write_block(
            doc,
            src_sheet,
            src_rect,
            dsh,
            dsc0,
            dsr0,
            payload,
            content,
            with_formatting=with_formatting and allow_clip,
            allow_clipboard=allow_clip,
        )
        if not ok_w and content == u"formulas" and not allow_clip:
            ok_w, note_w = _write_formulas_matrix(doc, dsh, dsc0, dsr0, payload)
        if not ok_w and content == u"values":
            ok_w, note_w = _write_values(dsh, dsc0, dsr0, payload)
        if ok_w:
            any_ok = True
            ok_dests.append(u"%s!%s" % (actual, dest_cell))
            notes.append(note_w)
            if involve_dest:
                # cols+= если вставка/замена правее прежнего used.
                ua_sc, ua_sr, ua_ec, ua_er = _used_area(dsh)
                expanded = 0
                if dsc0 + w - 1 > ua_ec:
                    expanded = (dsc0 + w - 1) - max(ua_ec, -1)
                if mode == u"insert" and insert_axis == u"cols":
                    expanded = max(expanded, w)
                inv = involve_dest_sheet(
                    doc, actual, header_row_0, src_name, expanded_cols=expanded
                )
                involve_notes.append(inv)
            else:
                involve_notes.append(u"involve=skip")
        else:
            notes.append(u"%s: %s" % (actual, note_w))
        di = di + 1

    if clear_source and any_ok:
        clear_rect(src_sheet, src_rect)
        notes.append(u"clear_source")

    src_lbl = u"%s!%s" % (src_name, _rect_label(sc, sr, ec, er))
    dest_lbl = u", ".join(ok_dests) if ok_dests else u"—"
    axis_note = u""
    if mode == u"insert":
        axis_note = u"/%s" % insert_axis
    content_note = content
    if content == u"values" and int(payload.get("formula_count") or 0) > 0:
        content_note = u"content=values; formulas_src=%d → values" % int(
            payload.get("formula_count") or 0
        )
    else:
        content_note = u"content=%s" % content
    involve_lbl = u""
    if involve_dest and ok_dests:
        inv_names = []
        for dn in ok_dests:
            inv_names.append(dn.split(u"!")[0])
        involve_lbl = u"; involve=%s" % u",".join(inv_names)
    elif not involve_dest:
        involve_lbl = u"; involve=skip"

    summary = (
        u"%s → %s; mode=%s%s; %s; %d×%d%s"
        % (
            src_lbl,
            dest_lbl,
            mode,
            axis_note,
            content_note,
            h,
            w,
            involve_lbl,
        )
    )
    extra = [n for n in notes if n and n not in (u"values", u"paste=clipboard")]
    if extra:
        # Краткие предупреждения (clipboard fallback и т.п.)
        warns = []
        for n in extra:
            if u"no ref shift" in n or u"failed" in n or u"format skipped" in n:
                warns.append(n)
        if warns:
            summary = summary + u"; " + u"; ".join(warns[:3])
    return any_ok, summary


def run_copy_ranges_blocks(doc, blocks):
    """
    Список блоков. → (status, note) status: ok|ошибка|пропуск.
    """
    if not blocks:
        return u"пропуск", u"пустая C"
    ok_n = 0
    err_n = 0
    notes = []
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        ok, note = execute_copy_ranges_block(doc, block)
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


def lm_pp_range_copy_ranges(doc, sheet, data_range, header_row_range, *extra_args):
    """
    RANGE-обёртка: тот же core; source_sheet в блоке обязателен (v1).

    Без поля sheet блок выполняется один раз — когда хост = source_sheet
    (иначе при обходе листов копирование дублировалось бы).
    """
    _unused = (data_range, header_row_range)
    sheet_name = sheet.Name if sheet is not None else u""
    fn_key = COPY_RANGES_FN
    if lm._lm_pp_param_skip_sheet(fn_key, extra_args, doc, sheet):
        return
    blocks = lm._lm_pp_blocks_for_sheet(fn_key, extra_args, doc, sheet)
    if not blocks:
        lm._lm_log_postprocess(
            doc, sheet_name, u"диапазон", fn_key, u"пропуск", u"нет блока JSON для листа"
        )
        return
    filtered = []
    for b in blocks:
        if not isinstance(b, dict):
            continue
        has_ctx = bool(lm._lm_pp_block_target_sheet_names(b))
        if has_ctx:
            filtered.append(b)
            continue
        # Глобальный блок: только на листе-источнике.
        src = unicode(
            b.get("source_sheet") or b.get("from_sheet") or u""
        ).strip()
        if src == u"":
            continue
        if sheet is not None and unicode(sheet.Name).casefold() == src.casefold():
            filtered.append(b)
    if not filtered:
        return
    status, note = run_copy_ranges_blocks(doc, filtered)
    lm._lm_log_postprocess(doc, sheet_name, u"диапазон", fn_key, status, note)
