# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.698"
"""
Чтение ODS / XLSX / XLSM (и .xls через внешнюю конвертацию) в матрицу для collect_workbooks.

Потоковое чтение строк + опции values_mode / merged_cells / trim_empty.
"""

try:
    unicode
except NameError:
    unicode = str

BOOK_SOURCE_EXTENSIONS = (u".ods", u".xls", u".xlsx", u".xlsm")


def is_book_source_path(path):
    p = unicode(path or u"").strip().lower()
    if p.endswith(BOOK_SOURCE_EXTENSIONS):
        return True
    q = p.split(u"?", 1)[0]
    return q.endswith(BOOK_SOURCE_EXTENSIONS)


def book_source_file_kind(path):
    p = unicode(path or u"").strip().lower()
    if p.endswith(u".xlsm"):
        return u"xlsm"
    if p.endswith(u".xlsx"):
        return u"xlsx"
    if p.endswith(u".xls"):
        return u"xls"
    if p.endswith(u".ods"):
        return u"ods"
    return u""


def book_source_normalize_local_path(path):
    p = unicode(path or u"").strip()
    if p == u"":
        return p
    try:
        from libre_macros_bundle_paths import normalize_source_path
        return normalize_source_path(p)
    except Exception:
        return p


def book_source_sheet_names(path):
    """Имена листов книги (локальный путь)."""
    kind = book_source_file_kind(path)
    p = book_source_normalize_local_path(path)
    if kind in (u"xlsx", u"xlsm"):
        import openpyxl_bundled  # noqa: F401
        from openpyxl import load_workbook
        wb = load_workbook(p, read_only=True, data_only=True)
        try:
            return [unicode(n) for n in wb.sheetnames]
        finally:
            try:
                wb.close()
            except Exception:
                pass
    if kind == u"ods":
        import odf_bundled  # noqa: F401
        from odf.opendocument import load
        from odf.table import Table
        doc = load(p)
        tables = doc.spreadsheet.getElementsByType(Table)
        out = []
        for table in tables:
            out.append(unicode(table.getAttribute(u"name") or u""))
        return out
    return []


def _xlsx_data_only(values_mode):
    vm = unicode(values_mode or u"cached")
    if vm == u"formulas":
        return False
    return True


def _cell_to_matrix_value(v, values_mode, keep_text=False):
    if v is None:
        return u""
    if keep_text:
        if isinstance(v, bool):
            return unicode(v)
        return unicode(v)
    try:
        import libre_macros_direct_lib as direct
        return direct._direct_coerce_cell_value(v)
    except Exception:
        if isinstance(v, bool):
            return unicode(v)
        if isinstance(v, (int, float)):
            return v
        return unicode(v).strip() if v is not None else u""


def _xlsx_formula_text(v):
    if v is None:
        return u""
    text = unicode(v)
    if text.startswith(u"="):
        return text
    return u""


def _xlsx_pick_cell_value(cached_v, formula_v, values_mode):
    vm = unicode(values_mode or u"cached")
    if vm == u"formulas":
        fo = _xlsx_formula_text(formula_v)
        if fo != u"":
            return fo
        return formula_v if formula_v is not None else u""
    if vm == u"prefer_cached":
        if cached_v is None or (isinstance(cached_v, unicode) and unicode(cached_v).strip() == u""):
            fo = _xlsx_formula_text(formula_v)
            if fo != u"":
                return fo
            if formula_v is not None and formula_v != u"":
                return formula_v
            return u"" if cached_v is None else cached_v
        return cached_v
    return cached_v


def _xlsx_build_merge_fill_map(ws, values_mode=u"cached", ws_formulas=None):
    """col0,row0 -> value для merged_cells=fill."""
    fill = {}
    try:
        ranges = getattr(ws, u"merged_cells", None)
        if ranges is None and ws_formulas is not None:
            ranges = getattr(ws_formulas, u"merged_cells", None)
        if ranges is None:
            return fill
        for mrange in ranges.ranges:
            try:
                min_col = int(mrange.min_col) - 1
                min_row = int(mrange.min_row) - 1
                max_col = int(mrange.max_col) - 1
                max_row = int(mrange.max_row) - 1
            except Exception:
                continue
            try:
                cached_v = ws.cell(row=min_row + 1, column=min_col + 1).value
            except Exception:
                cached_v = u""
            formula_v = None
            if ws_formulas is not None:
                try:
                    formula_v = ws_formulas.cell(row=min_row + 1, column=min_col + 1).value
                except Exception:
                    formula_v = None
            val = _xlsx_pick_cell_value(cached_v, formula_v, values_mode)
            rr = min_row
            while rr <= max_row:
                cc = min_col
                while cc <= max_col:
                    fill[(cc, rr)] = val
                    cc = cc + 1
                rr = rr + 1
    except Exception:
        pass
    return fill


def _ods_cell_formula_text(cell):
    try:
        raw = unicode(cell.getAttribute(u"formula") or u"").strip()
    except Exception:
        return u""
    if raw == u"":
        return u""
    # of:=SUM(A1:A2) / msoxl:=A1+B1
    if u":=" in raw:
        raw = raw.split(u":=", 1)[1]
    if not raw.startswith(u"="):
        raw = u"=" + raw
    return raw


def _ods_cell_by_values_mode(cell, values_mode):
    vm = unicode(values_mode or u"cached")
    try:
        import libre_macros_direct_lib as direct
        cached = direct._ods_cell_value(cell)
    except Exception:
        cached = u""
    fo = _ods_cell_formula_text(cell)
    if vm == u"formulas":
        if fo != u"":
            return fo
        return cached
    if vm == u"prefer_cached":
        empty = cached is None or (isinstance(cached, unicode) and unicode(cached).strip() == u"")
        if empty and fo != u"":
            return fo
        return cached if cached is not None else u""
    return cached if cached is not None else u""


def _ods_expand_row_with_spans(row_elem, end_col0, fill_mode, values_mode=u"cached"):
    from odf.table import TableCell
    out = {}
    col = 0
    for cell in row_elem.getElementsByType(TableCell):
        try:
            repeat = int(cell.getAttribute(u"numbercolumnsrepeated") or 1)
        except Exception:
            repeat = 1
        if repeat < 1:
            repeat = 1
        try:
            cspan = int(cell.getAttribute(u"numbercolumnspanned") or 1)
        except Exception:
            cspan = 1
        try:
            rspan = int(cell.getAttribute(u"numberrowspanned") or 1)
        except Exception:
            rspan = 1
        if cspan < 1:
            cspan = 1
        if rspan < 1:
            rspan = 1
        val = _ods_cell_by_values_mode(cell, values_mode)
        ci = 0
        while ci < repeat:
            cidx = col + ci
            if cidx <= int(end_col0):
                out[cidx] = val
            if fill_mode and (cspan > 1 or rspan > 1):
                cc = 1
                while cc < cspan:
                    if cidx + cc <= int(end_col0):
                        out[cidx + cc] = val
                    cc = cc + 1
            ci = ci + 1
        col = col + repeat
    if not out:
        return []
    width = min(int(end_col0) + 1, max(out.keys()) + 1)
    row = []
    c = 0
    while c < width:
        row.append(_cell_to_matrix_value(out.get(c, u""), values_mode))
        c = c + 1
    return row


def book_source_resolve_bounds(read_path, kind, sheet_index, hdr_row0, start_col0, trim_empty, last_row_col_spec=u""):
    """(end_col0, end_row0) включительно, 0-based."""
    import libre_macros_direct_lib as direct
    hdr_row0 = int(hdr_row0 or 0)
    start_col0 = int(start_col0 or 0)
    if trim_empty:
        try:
            ec, er = direct._direct_resolve_sheet_used_bounds(
                read_path, u"xlsx" if kind in (u"xlsx", u"xlsm", u"xls") else u"ods",
                int(sheet_index), hdr_row0, start_col0, last_row_col_spec=last_row_col_spec,
            )
            return int(ec), int(er)
        except Exception:
            pass
    if kind in (u"xlsx", u"xlsm", u"xls"):
        import openpyxl_bundled  # noqa: F401
        from openpyxl import load_workbook
        wb = load_workbook(read_path, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[int(sheet_index)]
            dim = direct._direct_xlsx_bounds_from_dimension(ws)
            if dim is not None:
                return int(dim[0]), int(dim[1])
            mr = int(getattr(ws, u"max_row", 0) or 0)
            mc = int(getattr(ws, u"max_column", 0) or 0)
            return max(mc - 1, start_col0), max(mr - 1, hdr_row0)
        finally:
            try:
                wb.close()
            except Exception:
                pass
    import odf_bundled  # noqa: F401
    from odf.opendocument import load
    from odf.table import Table, TableRow
    doc = load(read_path)
    tables = doc.spreadsheet.getElementsByType(Table)
    if int(sheet_index) < 0 or int(sheet_index) >= len(tables):
        return start_col0 - 1, hdr_row0
    rows = tables[int(sheet_index)].getElementsByType(TableRow)
    end_row = len(rows) - 1
    end_col = start_col0
    ri = 0
    while ri < len(rows):
        row_map = dict(direct._ods_row_col_values(rows[ri]))
        if row_map:
            mc = max(row_map.keys())
            if mc > end_col:
                end_col = mc
        ri = ri + 1
    return int(end_col), int(end_row)


def _xlsx_collect_rows_range( read_path, sheet_index, first_row0, last_row0, end_col0, values_mode, fill_mode, yield_callback=None):
    """Список строк диапазона XLSX/XLSM."""
    import openpyxl_bundled  # noqa: F401
    from openpyxl import load_workbook
    vm = unicode(values_mode or u"cached")
    need_formulas = vm in (u"formulas", u"prefer_cached")
    use_stream = (not fill_mode) and (not need_formulas)
    out = []
    if use_stream:
        wb = load_workbook(read_path, read_only=True, data_only=_xlsx_data_only(vm))
        try:
            ws = wb.worksheets[int(sheet_index)]
            rr = first_row0
            for row_vals in ws.iter_rows(
                min_row=first_row0 + 1,
                max_row=last_row0 + 1,
                min_col=1,
                max_col=end_col0 + 1,
                values_only=True,
            ):
                if yield_callback and rr % 200 == 0:
                    yield_callback()
                row = []
                ci = 0
                while ci <= end_col0:
                    v = row_vals[ci] if ci < len(row_vals) else None
                    row.append(_cell_to_matrix_value(v, vm))
                    ci = ci + 1
                out.append(row)
                rr = rr + 1
        finally:
            try:
                wb.close()
            except Exception:
                pass
        return out

    wb_cached = None
    wb_formulas = None
    try:
        if vm == u"formulas":
            wb_formulas = load_workbook(read_path, read_only=False, data_only=False)
            ws_cached = wb_formulas.worksheets[int(sheet_index)]
            ws_formulas = ws_cached
        elif vm == u"prefer_cached":
            wb_cached = load_workbook(read_path, read_only=False, data_only=True)
            wb_formulas = load_workbook(read_path, read_only=False, data_only=False)
            ws_cached = wb_cached.worksheets[int(sheet_index)]
            ws_formulas = wb_formulas.worksheets[int(sheet_index)]
        else:
            wb_cached = load_workbook(read_path, read_only=False, data_only=True)
            ws_cached = wb_cached.worksheets[int(sheet_index)]
            ws_formulas = None
        merge_fill = {}
        if fill_mode:
            merge_fill = _xlsx_build_merge_fill_map(
                ws_cached, values_mode=vm, ws_formulas=ws_formulas,
            )
        rr = first_row0
        while rr <= last_row0:
            if yield_callback and rr % 200 == 0:
                yield_callback()
            row = []
            cc = 0
            while cc <= end_col0:
                if fill_mode and (cc, rr) in merge_fill:
                    v = merge_fill.get((cc, rr))
                else:
                    try:
                        cached_v = ws_cached.cell(row=rr + 1, column=cc + 1).value
                    except Exception:
                        cached_v = None
                    formula_v = None
                    if ws_formulas is not None:
                        try:
                            formula_v = ws_formulas.cell(row=rr + 1, column=cc + 1).value
                        except Exception:
                            formula_v = None
                    v = _xlsx_pick_cell_value(cached_v, formula_v, vm)
                row.append(_cell_to_matrix_value(v, vm))
                cc = cc + 1
            out.append(row)
            rr = rr + 1
    finally:
        for wb in (wb_cached, wb_formulas):
            if wb is None:
                continue
            try:
                wb.close()
            except Exception:
                pass
    return out


def _ods_collect_rows_range( read_path, sheet_index, first_row0, last_row0, end_col0, values_mode, fill_mode, yield_callback=None):
    """Список строк диапазона ODS."""
    import odf_bundled  # noqa: F401
    from odf.opendocument import load
    from odf.table import Table, TableRow, TableCell
    doc = load(read_path)
    tables = doc.spreadsheet.getElementsByType(Table)
    if int(sheet_index) < 0 or int(sheet_index) >= len(tables):
        return []
    table_rows = tables[int(sheet_index)].getElementsByType(TableRow)
    out = []
    rr = first_row0
    while rr <= last_row0 and rr < len(table_rows):
        if yield_callback and rr % 200 == 0:
            yield_callback()
        if fill_mode:
            row = _ods_expand_row_with_spans(table_rows[rr], end_col0, True, values_mode)
        else:
            row_elem = table_rows[rr]
            cells = row_elem.getElementsByType(TableCell)
            col = 0
            row_map = {}
            for cell in cells:
                try:
                    repeat = int(cell.getAttribute(u"numbercolumnsrepeated") or 1)
                except Exception:
                    repeat = 1
                if repeat < 1:
                    repeat = 1
                val = _ods_cell_by_values_mode(cell, values_mode)
                ci = 0
                while ci < repeat:
                    cidx = col + ci
                    if cidx <= int(end_col0):
                        row_map[cidx] = val
                    ci = ci + 1
                col = col + repeat
            row = []
            cc = 0
            while cc <= end_col0:
                row.append(_cell_to_matrix_value(row_map.get(cc, u""), values_mode))
                cc = cc + 1
        out.append(row)
        rr = rr + 1
    return out


def book_source_read_probe_rows(read_path, kind, sheet_index, end_col0, end_row0, values_mode, merged_cells):
    """
    Прочитать строки 0..end_row0 для построения матрицы-пробы (trim / merge / cols).
    """
    end_col0 = int(end_col0)
    end_row0 = int(end_row0)
    if end_row0 < 0 or end_col0 < 0:
        return []
    fill_mode = unicode(merged_cells or u"top_left") == u"fill"
    if kind in (u"xlsx", u"xlsm", u"xls"):
        return _xlsx_collect_rows_range(
            read_path, sheet_index, 0, end_row0, end_col0,
            values_mode, fill_mode,
        )
    return _ods_collect_rows_range(
        read_path, sheet_index, 0, end_row0, end_col0,
        values_mode, fill_mode,
    )


def book_source_iter_data_rows( read_path, kind, sheet_index, first_row0, last_row0, end_col0, values_mode, merged_cells, yield_callback=None):
    """
    Генератор строк данных [first_row0 .. last_row0].
    yield_callback() вызывается периодически (UI yield).
    """
    first_row0 = int(first_row0)
    last_row0 = int(last_row0)
    end_col0 = int(end_col0)
    if last_row0 < first_row0:
        return
    fill_mode = unicode(merged_cells or u"top_left") == u"fill"
    vm = unicode(values_mode or u"cached")
    # Потоковый путь только для XLSX без merge-fill и без formula/prefer_cached.
    stream_ok = (
        kind in (u"xlsx", u"xlsm", u"xls")
        and (not fill_mode)
        and vm not in (u"formulas", u"prefer_cached")
    )
    if stream_ok:
        import openpyxl_bundled  # noqa: F401
        from openpyxl import load_workbook
        wb = load_workbook(read_path, read_only=True, data_only=_xlsx_data_only(vm))
        try:
            ws = wb.worksheets[int(sheet_index)]
            rr = first_row0
            for row_vals in ws.iter_rows(
                min_row=first_row0 + 1,
                max_row=last_row0 + 1,
                min_col=1,
                max_col=end_col0 + 1,
                values_only=True,
            ):
                if yield_callback and rr % 200 == 0:
                    yield_callback()
                row = []
                ci = 0
                while ci <= end_col0:
                    v = row_vals[ci] if ci < len(row_vals) else None
                    row.append(_cell_to_matrix_value(v, vm))
                    ci = ci + 1
                yield row
                rr = rr + 1
        finally:
            try:
                wb.close()
            except Exception:
                pass
        return
    # ODS / fill / formulas: один проход чтения диапазона, затем yield по строкам.
    if kind in (u"xlsx", u"xlsm", u"xls"):
        batch = _xlsx_collect_rows_range(
            read_path, sheet_index, first_row0, last_row0, end_col0,
            values_mode, fill_mode, yield_callback=yield_callback,
        )
    else:
        batch = _ods_collect_rows_range(
            read_path, sheet_index, first_row0, last_row0, end_col0,
            values_mode, fill_mode, yield_callback=yield_callback,
        )
    for row in batch or []:
        yield row


def book_source_read_cells(read_path, kind, sheet_index, addrs, values_mode=u"cached"):
    """
    Прочитать набор ячеек одним открытием книги.

    addrs — iterable (col0, row0); возвращает dict[(col0, row0)] = raw value.
    """
    need = []
    seen = set()
    for addr in addrs or []:
        try:
            c0 = int(addr[0])
            r0 = int(addr[1])
        except Exception:
            continue
        if c0 < 0 or r0 < 0:
            continue
        key = (c0, r0)
        if key in seen:
            continue
        seen.add(key)
        need.append(key)
    out = {}
    if not need:
        return out
    max_col = 0
    max_row = 0
    for c0, r0 in need:
        if c0 > max_col:
            max_col = c0
        if r0 > max_row:
            max_row = r0
    # Читаем прямоугольник 0..max_row / 0..max_col и выбираем нужные.
    rows = []
    if kind in (u"xlsx", u"xlsm", u"xls"):
        rows = _xlsx_collect_rows_range(
            read_path, sheet_index, 0, max_row, max_col,
            values_mode, False,
        ) or []
    else:
        rows = _ods_collect_rows_range(
            read_path, sheet_index, 0, max_row, max_col,
            values_mode, False,
        ) or []
    for c0, r0 in need:
        if r0 < 0 or r0 >= len(rows):
            out[(c0, r0)] = None
            continue
        row = rows[r0]
        if c0 < 0 or c0 >= len(row):
            out[(c0, r0)] = None
            continue
        # Сырое значение до _cell_to_matrix_value уже внутри collect;
        # здесь значение уже приведено — для снимка годится.
        out[(c0, r0)] = row[c0]
    return out
