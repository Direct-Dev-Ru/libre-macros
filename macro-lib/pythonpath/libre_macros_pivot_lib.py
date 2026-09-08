# -*- coding: utf-8 -*-
from __future__ import print_function
"""
Библиотека сводных таблиц (DataPilot) для libre-macros.

Подключается из libre_macros_lib.py; публичный API — lm_pp_pivot_* и lm_pp_range_pivot_table.
"""
MACRO_VERSION = "3.10.708"
import json
import re

import uno
from com.sun.star.table.CellVertJustify import CENTER as VERT_CENTER

# Импорт основной lib — только после полной инициализации хелперов (см. libre_macros_lib).
import libre_macros_lib as lm
import libre_macros_values_lib as vl

def _lm_pp_pivot_resolve_output_bounds(pivot_sheet, dp_table=None):
    """
    Границы вывода сводной (0-based): sc, sr, ec, er.

    Предпочтительно getOutputRange() DataPilot; иначе значащая область листа.
    """
    if dp_table is not None:
        try:
            out = dp_table.getOutputRange()
            sc = int(out.StartColumn)
            sr = int(out.StartRow)
            ec = int(out.EndColumn)
            er = int(out.EndRow)
            if er >= sr and ec >= sc:
                return sc, sr, ec, er
        except Exception:
            pass
    return lm._lm_vlookup_sheet_used_area(pivot_sheet)


def _lm_pp_pivot_flatten_range_values(sheet, sc, sr, ec, er):
    """Заменить формулы/ссылки DataPilot на статические значения (без буфера)."""
    if sheet is None or er < sr or ec < sc:
        return False
    cell_range = sheet.getCellRangeByPosition(sc, sr, ec, er)
    try:
        data = cell_range.getDataArray()
        cell_range.setDataArray(data)
        return True
    except Exception:
        pass
    try:
        r = int(sr)
        while r <= int(er):
            c = int(sc)
            while c <= int(ec):
                cell = sheet.getCellByPosition(c, r)
                formula = str(cell.getFormula() or "").strip()
                if formula.startswith("="):
                    try:
                        ctype = cell.getType()
                    except Exception:
                        ctype = None
                    if ctype == 1:
                        val = cell.getValue()
                        cell.setFormula("")
                        cell.setValue(val)
                    else:
                        txt = cell.String
                        cell.setFormula("")
                        cell.setString(txt)
                c = c + 1
            r = r + 1
        return True
    except Exception:
        return False


_LM_PP_PIVOT_MATERIALIZED = []
_LM_PP_PIVOT_AS_VALUES_PENDING = []
_LM_PP_PIVOT_LIST_ROWS_PENDING = {}
_LM_PP_PIVOT_POST_VALUES_FORMAT_ENABLED = False

# Псевдоагрегат «Список_строк» (после as_values подставляется текст детализации).
_LM_PP_PIVOT_LIST_ROWS_FIELD_SEP = " | "
_LM_PP_PIVOT_LIST_ROWS_NAME_SEP = ": "
_LM_PP_PIVOT_LIST_ROWS_RECORD_SEP = "\n"
_LM_PP_PIVOT_LIST_ROWS_MAX_PER_CELL = 200
_LM_PP_PIVOT_LIST_ROWS_HEADER = "Список_строк"
_LM_PP_PIVOT_LIST_ROWS_FN_ALIASES = frozenset(
    {
        "list_rows",
        "list",
        "список_строк",
        "детализация",
        "detail_rows",
        "detail",
    }
)
_LM_PP_PIVOT_LIST_FIELDS_KEYS = (
    "list_fields",
    "detail_fields",
    "detail_columns",
    "поля_списка",
    "колонки_списка",
)


def lm_pp_pivot_set_post_values_format_enabled(enabled):
    """Вкл./выкл. оформление после материализации сводной (флаг задаётся в collect_workbooks)."""
    global _LM_PP_PIVOT_POST_VALUES_FORMAT_ENABLED
    _LM_PP_PIVOT_POST_VALUES_FORMAT_ENABLED = bool(enabled)


def lm_pp_pivot_post_values_format_enabled():
    return bool(_LM_PP_PIVOT_POST_VALUES_FORMAT_ENABLED)


def lm_pp_pivot_clear_materialized_registry():
    global _LM_PP_PIVOT_MATERIALIZED
    _LM_PP_PIVOT_MATERIALIZED = []


def lm_pp_pivot_clear_as_values_pending():
    global _LM_PP_PIVOT_AS_VALUES_PENDING
    global _LM_PP_PIVOT_LIST_ROWS_PENDING
    _LM_PP_PIVOT_AS_VALUES_PENDING = []
    _LM_PP_PIVOT_LIST_ROWS_PENDING = {}


def lm_pp_pivot_list_materialized_sheets():
    return list(_LM_PP_PIVOT_MATERIALIZED)


def lm_pp_pivot_list_as_values_pending():
    return list(_LM_PP_PIVOT_AS_VALUES_PENDING)


def lm_pp_pivot_get_list_rows_meta(sheet_name):
    name = str(sheet_name or "").strip()
    if name == "":
        return None
    return _LM_PP_PIVOT_LIST_ROWS_PENDING.get(name)


def lm_pp_pivot_register_materialized_sheet(sheet_name, header_row=0):
    global _LM_PP_PIVOT_MATERIALIZED
    name = str(sheet_name or "").strip()
    if name == "":
        return
    _LM_PP_PIVOT_MATERIALIZED.append(
        {"name": name, "header_row": int(header_row)}
    )


def lm_pp_pivot_register_as_values_pending(sheet_name, list_rows_meta=None):
    """Лист со сводной (as_values) — материализация отложена до финальной обработки."""
    global _LM_PP_PIVOT_AS_VALUES_PENDING
    global _LM_PP_PIVOT_LIST_ROWS_PENDING
    name = str(sheet_name or "").strip()
    if name == "":
        return
    if name not in _LM_PP_PIVOT_AS_VALUES_PENDING:
        _LM_PP_PIVOT_AS_VALUES_PENDING.append(name)
    if list_rows_meta and list_rows_meta.get("active"):
        _LM_PP_PIVOT_LIST_ROWS_PENDING[name] = dict(list_rows_meta)


def _lm_pp_pivot_guess_header_row(sheet, sc, sr, ec, er, config):
    """
    Строка заголовков данных на материализованной сводной (0-based).

    Первая строка вывода часто — область фильтра Page; заголовки — ниже.
    """
    if len(config.get("filter_fields") or []) > 0:
        hr = int(sr) + 1
        return hr if hr <= er else int(sr)
    if sr >= er:
        return int(sr)
    try:
        n0 = 0
        n1 = 0
        c = int(sc)
        while c <= int(ec):
            if _lm_pp_pivot_cell_has_content(sheet.getCellByPosition(c, sr)):
                n0 = n0 + 1
            if _lm_pp_pivot_cell_has_content(sheet.getCellByPosition(c, sr + 1)):
                n1 = n1 + 1
            c = c + 1
        if n1 > n0 and n0 <= 2:
            return int(sr) + 1
    except Exception:
        pass
    return int(sr)


def _lm_pp_pivot_sheet_names_set(doc):
    """Имена всех листов книги (снимок перед материализацией)."""
    names = set()
    if doc is None:
        return names
    try:
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            names.add(str(sheets.getByIndex(i).Name))
            i = i + 1
    except Exception:
        pass
    return names


def lm_pp_pivot_purge_lo_default_garbage_sheets(doc, keep_names=None):
    """Удалить автолисты SheetN / ЛистN (обёртка libre_macros_values_lib)."""
    vl.lm_values_purge_lo_default_sheets(
        doc, keep_names=keep_names, log_sheet="сводная_таблица"
    )


def _lm_pp_activate_sheet_cell(doc, sheet, col=0, row=0):
    """Активировать лист и выделить ячейку (чтобы не оставаться на пустом SheetN)."""
    if doc is None or sheet is None:
        return
    try:
        controller = doc.getCurrentController()
        if controller is None:
            return
        controller.setActiveSheet(sheet)
        cell = sheet.getCellByPosition(int(col), int(row))
        controller.select(cell)
    except Exception:
        pass


def _lm_pp_pivot_row_is_preamble(sheet, sc, ec, row):
    """Строка над заголовками данных: Filter, пустая, одна метка Page."""
    if sheet is None:
        return True
    filled = 0
    first_txt = ""
    c = int(sc)
    while c <= int(ec):
        cell = sheet.getCellByPosition(c, int(row))
        if _lm_pp_pivot_cell_has_content(cell):
            filled = filled + 1
            if filled == 1:
                first_txt = lm.cell_text(cell).strip().lower()
        c = c + 1
    if filled == 0:
        return True
    if filled == 1 and first_txt in (
        "filter",
        "фильтр",
        "filters",
        "фильтры",
        "page",
        "страница",
    ):
        return True
    return False


def _lm_pp_pivot_detect_data_header_row(sheet, sc, sr, ec, er, config):
    """0-based строка реальных заголовков столбцов (ниже Filter/Page)."""
    if len(config.get("filter_fields") or []) > 0:
        hr = int(sr) + 1
        while hr <= int(er) and _lm_pp_pivot_row_is_preamble(sheet, sc, ec, hr):
            hr = hr + 1
        return hr if hr <= er else int(sr)
    r = int(sr)
    while r <= int(er):
        if not _lm_pp_pivot_row_is_preamble(sheet, sc, ec, r):
            return r
        r = r + 1
    return int(sr)


def _lm_pp_pivot_trim_top_rows(sheet, sc, sr, ec, er, config):
    """
    Удалить служебные строки над заголовками (Filter, пустая…).

    Возвращает (sc, sr, ec, er, header_row) после обрезки.
    """
    header_row = _lm_pp_pivot_detect_data_header_row(
        sheet, sc, sr, ec, er, config
    )
    if header_row <= int(sr):
        return int(sc), int(sr), int(ec), int(er), int(header_row)
    count = int(header_row) - int(sr)
    try:
        sheet.getRows().removeByIndex(int(sr), count)
    except Exception:
        return int(sc), int(sr), int(ec), int(er), int(header_row)
    new_er = int(er) - count
    return int(sc), int(sr), int(ec), new_er, int(sr)


def _lm_pp_pivot_detect_header_row_by_consecutive_cells(sheet, sc, sr, ec, er, min_run=2):
    """
    0-based строка заголовков: первая строка, где подряд заполнены >= min_run ячеек.
    Используется для «табличной» материализации сводной (as_values).
    """
    if sheet is None:
        return int(sr)
    try:
        r = int(sr)
        while r <= int(er):
            run = 0
            c = int(sc)
            while c <= int(ec):
                if _lm_pp_pivot_cell_has_content(sheet.getCellByPosition(c, r)):
                    run = run + 1
                    if run >= int(min_run):
                        return r
                else:
                    run = 0
                c = c + 1
            r = r + 1
    except Exception:
        pass
    return int(sr)


def _lm_pp_pivot_detect_header_row_first_columns(sheet, sc, sr, ec, er, col_count=2):
    """
    0-based строка заголовков: первая строка, где заполнены первые col_count столбцов
    диапазона (sc, sc+1, …). Для сводной — ниже Filter/Data, на строке «Код | Менеджер».
    """
    if sheet is None:
        return int(sr)
    try:
        col_count = max(int(col_count), 1)
        cols = [int(sc) + i for i in range(col_count)]
        if cols[-1] > int(ec):
            return _lm_pp_pivot_detect_header_row_by_consecutive_cells(
                sheet, sc, sr, ec, er, min_run=col_count
            )
        r = int(sr)
        while r <= int(er):
            ok = True
            ci = 0
            while ci < len(cols):
                if not _lm_pp_pivot_cell_has_content(
                    sheet.getCellByPosition(cols[ci], r)
                ):
                    ok = False
                    break
                ci = ci + 1
            if ok:
                return r
            r = r + 1
    except Exception:
        pass
    return int(sr)


def _lm_pp_pivot_apply_materialized_sheet_formatting( doc, sheet, sc, sr, ec, er, header_row ):
    """Устарело: оформление inline в lm_pp_pivot_post_values_format (см. комментарий там)."""
    pass


_LM_PP_PIVOT_HEADER_COLLAPSE_DELIM = "\\"
_LM_PP_PIVOT_HEADER_EMPTY_VALUE = "[пусто]"


def _lm_pp_pivot_header_cell_raw(cell):
    """Текст ячейки шапки (trim)."""
    return str(lm.cell_text(cell) or "").strip()


def _lm_pp_pivot_header_value_is_empty(text):
    """Пустое значение уровня столбцового поля (пусто или «0»)."""
    t = str(text or "").strip()
    return t == "" or t == "0"


def _lm_pp_pivot_header_format_field_value(text):
    """Значение уровня для «Метка: значение»; пусто/0 → [пусто]."""
    if _lm_pp_pivot_header_value_is_empty(text):
        return _LM_PP_PIVOT_HEADER_EMPTY_VALUE
    return " ".join(str(text or "").split())


def _lm_pp_pivot_is_pivot_column_label(text):
    """Текст метки столбцового поля в строке перечня (любое имя, кроме служебного)."""
    t = str(text or "").strip()
    if t == "" or t == "0":
        return False
    low = t.casefold()
    if low in (
        "filter",
        "фильтр",
        "filters",
        "фильтры",
        "page",
        "страница",
    ):
        return False
    return True


def _lm_pp_pivot_count_column_labels_in_row(sheet, sc, ec, row, R):
    """Число меток столбцовых полей в строке (область данных справа от R)."""
    if sheet is None:
        return 0
    data_start = int(sc) + max(int(R), 0)
    count = 0
    c = data_start
    while c <= int(ec):
        if _lm_pp_pivot_is_pivot_column_label(
            _lm_pp_pivot_header_cell_raw(sheet.getCellByPosition(c, int(row)))
        ):
            count = count + 1
        c = c + 1
    return count


def _lm_pp_pivot_forward_fill_row_value(sheet, row, sc_start, col):
    """Значение в (row, col) с протягиванием слева от sc_start."""
    if sheet is None:
        return ""
    carry = ""
    c = int(sc_start)
    while c <= int(col):
        raw = _lm_pp_pivot_header_cell_raw(sheet.getCellByPosition(c, int(row)))
        if raw != "":
            carry = raw
        c = c + 1
    return carry


def _lm_pp_pivot_row_left_cols_empty(sheet, sc, row, R):
    """Первые R столбцов строки пусты."""
    if sheet is None or R <= 0:
        return True
    ci = 0
    while ci < R:
        if _lm_pp_pivot_header_cell_raw(
            sheet.getCellByPosition(int(sc) + ci, int(row))
        ) != "":
            return False
        ci = ci + 1
    return True


def _lm_pp_pivot_detect_column_fields_row( sheet, sc, ec, header_top, header_bottom, R, config=None ):
    """
    Первая строка перечня столбцовых полей сводной (любые имена меток).
    0-based; None если не найдена.
    """
    if sheet is None:
        return None
    data_start = int(sc) + max(int(R), 0)
    if data_start > int(ec):
        return None

    if config is not None:
        cfg_names = config.get("column_fields") or []
        if len(cfg_names) > 0:
            r = int(header_top)
            while r < int(header_bottom):
                if not _lm_pp_pivot_row_left_cols_empty(sheet, sc, r, R):
                    r = r + 1
                    continue
                found = 0
                ni = 0
                while ni < len(cfg_names):
                    name = str(cfg_names[ni] or "").strip()
                    if name != "":
                        c = data_start
                        while c <= int(ec):
                            raw = _lm_pp_pivot_header_cell_raw(
                                sheet.getCellByPosition(c, r)
                            )
                            if raw.casefold() == name.casefold():
                                found = found + 1
                                break
                            c = c + 1
                    ni = ni + 1
                if found > 0:
                    return r
                r = r + 1
            return None

    r = int(header_top)
    while r < int(header_bottom):
        if not _lm_pp_pivot_row_left_cols_empty(sheet, sc, r, R):
            r = r + 1
            continue
        if _lm_pp_pivot_count_column_labels_in_row(sheet, sc, ec, r, R) > 0:
            return r
        r = r + 1
    return None


def _lm_pp_pivot_parse_column_fields(sheet, sc, ec, fields_row, R, config=None):
    """Список (col, name) столбцовых полей слева направо из строки меток."""
    if sheet is None or fields_row is None:
        return []
    data_start = int(sc) + max(int(R), 0)
    fields = []
    cfg_names = []
    if config is not None:
        for nm in config.get("column_fields") or []:
            s = str(nm or "").strip()
            if s != "":
                cfg_names.append(s)

    if len(cfg_names) > 0:
        used_cols = set()
        ni = 0
        while ni < len(cfg_names):
            name = cfg_names[ni]
            c = data_start
            while c <= int(ec):
                raw = _lm_pp_pivot_header_cell_raw(sheet.getCellByPosition(c, fields_row))
                if raw.casefold() == name.casefold() and c not in used_cols:
                    fields.append((c, raw))
                    used_cols.add(c)
                    break
                c = c + 1
            ni = ni + 1
        if len(fields) > 0:
            return fields

    c = data_start
    while c <= int(ec):
        raw = _lm_pp_pivot_header_cell_raw(sheet.getCellByPosition(c, fields_row))
        if _lm_pp_pivot_is_pivot_column_label(raw):
            fields.append((c, raw))
        c = c + 1
    return fields


def _lm_pp_pivot_value_row_for_field(fields_row, field_index, header_bottom):
    """Строка значения для поля field_index: одна строка под перечнём меток."""
    vrow = int(fields_row) + 1 + int(field_index)
    if vrow > int(header_bottom):
        vrow = int(header_bottom)
    return vrow


def _lm_pp_pivot_build_labeled_column_header( sheet, col, fields, fields_row, header_bottom, data_start ):
    """«Метка: значение\\Метка2: значение» — по одному значению на каждую метку перечня."""
    parts = []
    fi = 0
    while fi < len(fields):
        _fcol, fname = fields[fi]
        vrow = _lm_pp_pivot_value_row_for_field(fields_row, fi, header_bottom)
        val = _lm_pp_pivot_forward_fill_row_value(sheet, vrow, data_start, col)
        parts.append(
            "%s: %s"
            % (
                fname,
                _lm_pp_pivot_header_format_field_value(val),
            )
        )
        fi = fi + 1
    return _LM_PP_PIVOT_HEADER_COLLAPSE_DELIM.join(parts)


def _lm_pp_pivot_trim_preamble_rows(sheet, sc, sr, ec, er):
    """Удалить Filter/Page/пустые строки над шапкой. Возвращает (sc, sr, ec, er)."""
    if sheet is None:
        return int(sc), int(sr), int(ec), int(er)
    try:
        r = int(sr)
        top = int(sr)
        while r <= int(er) and _lm_pp_pivot_row_is_preamble(sheet, sc, ec, r):
            r = r + 1
        if r > top:
            count = r - top
            sheet.getRows().removeByIndex(top, count)
            er = int(er) - count
    except Exception:
        pass
    return int(sc), int(sr), int(ec), int(er)


def _lm_pp_pivot_detect_row_field_col_count(sheet, sc, ec, header_top, header_bottom):
    """
    Число левых столбцов строковых полей: подпись только в нижней строке шапки.
    """
    if sheet is None:
        return 0
    R = 0
    c = int(sc)
    while c <= int(ec):
        bottom = _lm_pp_pivot_header_cell_raw(
            sheet.getCellByPosition(c, int(header_bottom))
        )
        if bottom == "":
            break
        upper_nonempty = False
        r = int(header_top)
        while r < int(header_bottom):
            if _lm_pp_pivot_header_cell_raw(
                sheet.getCellByPosition(c, r)
            ) != "":
                upper_nonempty = True
                break
            r = r + 1
        if upper_nonempty:
            break
        R = R + 1
        c = c + 1
    return R


def _lm_pp_pivot_collapse_header_rows(sheet, sc, sr, ec, er, config=None):
    """
    Многоуровневая шапка материализованной сводной → одна строка.

    Строка перечня столбцовых полей (любые имена) задаёт метки;
    под ней — по одному значению на метку: «Имя: значение», пусто/0 → [пусто].
    Возвращает (sc, sr, ec, er, header_row, note).
    """
    if sheet is None or int(er) < int(sr) or int(ec) < int(sc):
        return int(sc), int(sr), int(ec), int(er), int(sr), "пусто"

    header_top = int(sr)
    col_count = 2
    if config is not None:
        row_fields = config.get("row_fields") or []
        if len(row_fields) > 0:
            col_count = max(len(row_fields), 1)
    header_bottom = _lm_pp_pivot_detect_header_row_first_columns(
        sheet, sc, header_top, ec, er, col_count=col_count
    )
    H = int(header_bottom) - header_top + 1
    if H <= 1:
        return int(sc), int(sr), int(ec), int(er), header_top, "шапка 1 строка"

    R = 0
    if config is not None:
        row_fields = config.get("row_fields") or []
        R = len(row_fields)
    if R <= 0:
        R = _lm_pp_pivot_detect_row_field_col_count(
            sheet, sc, ec, header_top, header_bottom
        )

    data_start = int(sc) + max(int(R), 0)
    fields_row = _lm_pp_pivot_detect_column_fields_row(
        sheet, sc, ec, header_top, header_bottom, R, config=config
    )
    column_fields = []
    if fields_row is not None:
        column_fields = _lm_pp_pivot_parse_column_fields(
            sheet, sc, ec, fields_row, R, config=config
        )

    ncol = int(ec) - int(sc) + 1
    new_headers = []
    c_idx = 0
    while c_idx < ncol:
        col = int(sc) + c_idx
        if c_idx < R:
            new_headers.append(
                _lm_pp_pivot_header_cell_raw(
                    sheet.getCellByPosition(col, int(header_bottom))
                )
            )
        elif len(column_fields) > 0 and fields_row is not None:
            new_headers.append(
                _lm_pp_pivot_build_labeled_column_header(
                    sheet,
                    col,
                    column_fields,
                    fields_row,
                    header_bottom,
                    data_start,
                )
            )
        else:
            new_headers.append(
                _lm_pp_pivot_header_format_field_value(
                    _lm_pp_pivot_header_cell_raw(
                        sheet.getCellByPosition(col, int(header_bottom))
                    )
                )
            )
            if new_headers[-1] == _LM_PP_PIVOT_HEADER_EMPTY_VALUE:
                new_headers[-1] = ""
        c_idx = c_idx + 1

    try:
        c = int(sc)
        hi = 0
        while hi < len(new_headers):
            cell = sheet.getCellByPosition(c, header_top)
            try:
                cell.setFormula("")
            except Exception:
                pass
            cell.setString(new_headers[hi])
            c = c + 1
            hi = hi + 1
    except Exception:
        return int(sc), int(sr), int(ec), int(er), header_top, "ошибка записи шапки"

    if H > 1:
        try:
            sheet.getRows().removeByIndex(header_top + 1, H - 1)
            er = int(er) - (H - 1)
        except Exception:
            return int(sc), int(sr), int(ec), int(er), header_top, "ошибка удаления строк шапки"

    return int(sc), int(sr), int(ec), int(er), header_top, "шапка %d→1" % H


_LM_PP_PIVOT_MATERIALIZED_MAX_COL_WIDTH_MM = 50.0


_LM_PP_PIVOT_MATERIALIZED_HEADER_HEIGHT_EXTRA_RATIO = 0.25


_LM_PP_PIVOT_GRAND_TOTAL_MARKERS = frozenset(
    {
        "total result",
        "grand total",
        "total",
        "итого",
        "итог",
        "общий итог",
        "результат",
        "всего",
    }
)


def _lm_pp_pivot_apply_row_bold(sheet, sc, ec, row):
    """Жирный шрифт строки (CharWeight=150)."""
    if sheet is None:
        return
    row = int(row)
    try:
        row_range = sheet.getCellRangeByPosition(int(sc), row, int(ec), row)
        row_range.CharWeight = lm._LM_FMT_CHAR_WEIGHT_BOLD
        return
    except Exception:
        pass
    c = int(sc)
    while c <= int(ec):
        try:
            sheet.getCellByPosition(c, row).CharWeight = lm._LM_FMT_CHAR_WEIGHT_BOLD
        except Exception:
            pass
        c = c + 1


def _lm_pp_pivot_apply_header_bold(sheet, sc, ec, header_row):
    """Жирный шрифт строки заголовка (CharWeight=150, как в заголовок_плюс_высота)."""
    _lm_pp_pivot_apply_row_bold(sheet, sc, ec, header_row)


def _lm_pp_pivot_cell_looks_like_grand_total_label(text):
    """Метка строки «Итого» / Total Result в первом столбце сводной."""
    t = " ".join(str(text or "").split()).casefold()
    if t == "":
        return False
    if t in _LM_PP_PIVOT_GRAND_TOTAL_MARKERS:
        return True
    prefixes = ("total result", "grand total", "общий итог", "итого ")
    pi = 0
    while pi < len(prefixes):
        if t.startswith(prefixes[pi]):
            return True
        pi = pi + 1
    return False


def _lm_pp_pivot_row_looks_like_grand_total(sheet, sc, row):
    """True, если в столбце sc строки row — подпись итоговой строки сводной."""
    if sheet is None:
        return False
    raw = _lm_pp_pivot_header_cell_raw(sheet.getCellByPosition(int(sc), int(row)))
    return _lm_pp_pivot_cell_looks_like_grand_total_label(raw)


def _lm_pp_pivot_detect_row_field_cols_after_collapse( sheet, sc, ec, header_row, config=None ):
    """
    Число столбцов строковых полей после схлопывания шапки.
    Столбцы данных распознаются по «Метка: значение» или '\\' в заголовке.
    """
    if config is not None:
        row_fields = config.get("row_fields") or []
        if len(row_fields) > 0:
            return len(row_fields)
    if sheet is None:
        return 0
    R = 0
    c = int(sc)
    while c <= int(ec):
        t = _lm_pp_pivot_header_cell_raw(
            sheet.getCellByPosition(c, int(header_row))
        )
        if t == "":
            break
        if ":" in t or _LM_PP_PIVOT_HEADER_COLLAPSE_DELIM in t:
            break
        R = R + 1
        c = c + 1
    return R


def _lm_pp_pivot_apply_data_number_format( doc, sheet, sc, ec, er, header_row, row_field_cols ):
    """
    Числовой формат # ##0,00 на столбцы данных (правее строковых полей), одним диапазоном.
    """
    if sheet is None or doc is None:
        return
    data_c0 = int(sc) + max(int(row_field_cols), 0)
    if data_c0 > int(ec) or int(header_row) >= int(er):
        return
    try:
        values_range = sheet.getCellRangeByPosition(
            data_c0, int(header_row) + 1, int(ec), int(er)
        )
        fmt_key = lm._lm_pp_number_format_key(doc, lm._LM_PP_FORMAT_TOTAL_SUM)
        values_range.NumberFormat = fmt_key
    except Exception:
        pass


def _lm_pp_pivot_autofit_header_row_height(sheet, header_row, extra_ratio=None):
    """
    Автоподбор высоты строки заголовка, затем +extra_ratio (по умолчанию 25%).
    """
    if sheet is None:
        return
    try:
        ratio = float(
            extra_ratio
            if extra_ratio is not None
            else _LM_PP_PIVOT_MATERIALIZED_HEADER_HEIGHT_EXTRA_RATIO
        )
    except (TypeError, ValueError):
        ratio = _LM_PP_PIVOT_MATERIALIZED_HEADER_HEIGHT_EXTRA_RATIO
    if ratio < 0.0:
        ratio = 0.0
    try:
        row_obj = sheet.getRows().getByIndex(int(header_row))
        row_obj.OptimalHeight = True
        lm._lm_ui_yield(force=True)
        h = int(row_obj.Height)
        if h < 1:
            h = 0
        row_obj.OptimalHeight = False
        row_obj.Height = int(h + round(h * ratio))
    except Exception:
        pass


def _lm_pp_pivot_freeze_header_row(doc, sheet, header_row):
    """Закрепить строки 0…header_row (freeze panes)."""
    if doc is None or sheet is None:
        return
    try:
        try:
            import libre_macros_lib as _lm
            freeze_fn = getattr(_lm, '_lm_pp_freeze_sheet_header', None)
            if freeze_fn is not None:
                freeze_fn(doc, sheet, int(header_row))
                return
        except Exception:
            pass
        freeze_rows = int(header_row) + 1
        if freeze_rows < 1:
            freeze_rows = 1
        controller = doc.getCurrentController()
        if controller is None:
            return
        try:
            active = controller.getActiveSheet()
            if active is None or active.Name != sheet.Name:
                controller.setActiveSheet(sheet)
        except Exception:
            controller.setActiveSheet(sheet)
        try:
            controller.FirstVisibleColumn = 0
        except Exception:
            pass
        try:
            controller.FirstVisibleRow = 0
        except Exception:
            pass
        controller.freezeAtPosition(0, freeze_rows)
    except Exception:
        pass


def _lm_pp_pivot_cap_column_widths_mm(sheet, sc, ec, max_mm=None):
    """
    Ограничить ширину столбцов сверху (мм): если уже меньше max — не менять.
    Ширина Calc в 1/100 мм (как в lm_pp_range_set_column_widths).
    """
    if sheet is None:
        return
    try:
        limit_mm = float(
            max_mm
            if max_mm is not None
            else _LM_PP_PIVOT_MATERIALIZED_MAX_COL_WIDTH_MM
        )
    except (TypeError, ValueError):
        limit_mm = _LM_PP_PIVOT_MATERIALIZED_MAX_COL_WIDTH_MM
    max_w = int(round(limit_mm * 100.0))
    if max_w < 1:
        return
    try:
        columns = sheet.getColumns()
        c = int(sc)
        while c <= int(ec):
            try:
                col = columns.getByIndex(c)
                w = int(col.Width)
                if w > max_w:
                    col.OptimalWidth = False
                    col.Width = max_w
            except Exception:
                pass
            c = c + 1
    except Exception:
        pass


def _lm_pp_pivot_apply_materialized_format( doc, pivot_sheet, sc, sr, ec, er, header_row, config=None ):
    """
    Оформление материализованной сводной (те же алгоритмы, что у постобработки):
    тонкая_сетка, перенос, вертикаль_центр, центр заголовка, жирный заголовок,
    числовой формат данных, cap ширины, авто_высота (+25% заголовку), жирный итог, freeze.
    """
    if pivot_sheet is None or int(er) < int(sr) or int(ec) < int(sc):
        return
    header_row = int(header_row)
    sc = int(sc)
    ec = int(ec)
    er = int(er)
    row_field_cols = _lm_pp_pivot_detect_row_field_cols_after_collapse(
        pivot_sheet, sc, ec, header_row, config=config
    )
    header_range = pivot_sheet.getCellRangeByPosition(sc, header_row, ec, header_row)
    if header_row < er:
        data_range = pivot_sheet.getCellRangeByPosition(
            sc, header_row + 1, ec, er
        )
    else:
        data_range = header_range

    lm._lm_pp_apply_table_grid(
        data_range,
        doc=doc,
        sheet=pivot_sheet,
        grid_label="тонкая_сетка",
        header_row_range=header_range,
    )

    lm.lm_pp_range_word_wrap(
        doc, pivot_sheet, data_range, header_range, "да", "заголовок"
    )

    if header_row < er:
        try:
            data_range.VertJustify = VERT_CENTER
        except Exception:
            pass
    lm._lm_apply_header_row_alignment(
        pivot_sheet, ec, header_row, hori="center", vert="center"
    )
    _lm_pp_pivot_apply_header_bold(pivot_sheet, sc, ec, header_row)

    _lm_pp_pivot_apply_data_number_format(
        doc, pivot_sheet, sc, ec, er, header_row, row_field_cols
    )

    _lm_pp_pivot_cap_column_widths_mm(pivot_sheet, sc, ec)

    if header_row < er:
        lm.lm_pp_range_autofit_row_heights(
            doc, pivot_sheet, data_range, header_range
        )
    _lm_pp_pivot_autofit_header_row_height(pivot_sheet, header_row)
    if header_row < er and _lm_pp_pivot_row_looks_like_grand_total(
        pivot_sheet, sc, er
    ):
        _lm_pp_pivot_apply_row_bold(pivot_sheet, sc, ec, er)
    _lm_pp_pivot_freeze_header_row(doc, pivot_sheet, header_row)


def _lm_pp_pivot_trim_to_header_row(sheet, sc, sr, ec, er, header_row):
    """Удалить строки выше header_row. Возвращает (sc, sr, ec, er, header_row)."""
    if sheet is None:
        return int(sc), int(sr), int(ec), int(er), int(header_row)
    try:
        header_row = int(header_row)
        sr = int(sr)
        er = int(er)
    except Exception:
        return int(sc), int(sr), int(ec), int(er), int(header_row)
    if header_row <= sr:
        return int(sc), int(sr), int(ec), int(er), int(header_row)
    count = header_row - sr
    try:
        sheet.getRows().removeByIndex(sr, count)
    except Exception:
        return int(sc), int(sr), int(ec), int(er), int(header_row)
    new_er = er - count
    return int(sc), int(sr), int(ec), int(new_er), int(sr)


def lm_pp_pivot_post_values_format(doc, pivot_sheet):
    """
    После «Только_значения» на листе сводной: обрезка Filter/Data и оформление.
    Управляется lm_pp_pivot_set_post_values_format_enabled (флаг в collect_workbooks).
    """
    do_format = lm_pp_pivot_post_values_format_enabled()
    if doc is None or pivot_sheet is None:
        return 0, "нет листа"
    
    sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(pivot_sheet)
    if er < sr or ec < sc:
        return 0, "пустой лист"

    sc, sr, ec, er = _lm_pp_pivot_trim_preamble_rows(
        pivot_sheet, sc, sr, ec, er
    )
    if er < sr or ec < sc:
        return 0, "пустой лист"

    sc, sr, ec, er, header_row, collapse_note = _lm_pp_pivot_collapse_header_rows(
        pivot_sheet, sc, sr, ec, er, config=None
    )

    if not do_format:
        return header_row, "заголовок строка %d; %s; без оформления" % (
            header_row + 1,
            collapse_note,
        )

    # Оформление (включая freeze); затем простой меню-автофильтр (без AF_*/смарт-таблиц).
    if er >= sr and ec >= sc:
        try:
            _lm_pp_pivot_apply_materialized_format(
                doc, pivot_sheet, sc, sr, ec, er, header_row
            )
            try:
                # Ещё раз снять любые DatabaseRanges на листе (остатки DP/TableStyle).
                try:
                    rem = getattr(lm, "_lm_pp_remove_all_sheet_database_ranges", None)
                    if callable(rem):
                        rem(doc, pivot_sheet)
                    else:
                        lm._lm_pp_remove_sheet_autofilters(doc, pivot_sheet)
                except Exception:
                    pass
                from libre_macros_split_lib import _split_activate_document
                from libre_macros_split_lib import _split_apply_menu_autofilter
                try:
                    _split_activate_document(doc)
                    ctrl = doc.getCurrentController() if doc is not None else None
                    if ctrl is not None:
                        try:
                            ctrl.setActiveSheet(pivot_sheet)
                        except Exception:
                            pass
                except Exception:
                    pass
                ok_af = bool(
                    _split_apply_menu_autofilter(
                        doc, pivot_sheet, int(sc), int(header_row), int(ec), int(er)
                    )
                )
                if ok_af:
                    try:
                        sn = str(getattr(pivot_sheet, "Name", "") or "").strip()
                        if sn:
                            import libre_macros_collect_cfg as _cw_cfg_pp

                            done = set(
                                getattr(
                                    _cw_cfg_pp,
                                    "_MERGE_SHEETS_WITH_MENU_AUTOFILTER",
                                    None,
                                )
                                or set()
                            )
                            done.add(sn)
                            _cw_cfg_pp._MERGE_SHEETS_WITH_MENU_AUTOFILTER = done
                    except Exception:
                        pass
            except Exception:
                pass
        except Exception:
            pass
    return header_row, "заголовок строка %d; %s; оформление" % (
        header_row + 1,
        collapse_note,
    )


def _lm_pp_pivot_apply_as_values(doc, pivot_sheet):
    """
    as_values (устар. прямой вызов): inplace + post_values_format + сводная_пп.
    При создании сводной не вызывается — см. lm_pp_pivot_register_as_values_pending.
    """
    if doc is None or pivot_sheet is None:
        return None, "нет листа сводной"
    pivot_name = str(pivot_sheet.Name)
    ok, note = vl.lm_values_pivot_inplace(doc, pivot_sheet)
    if not ok:
        return None, note
    header_row, fmt_note = lm_pp_pivot_post_values_format(doc, pivot_sheet)
    try:
        lr_meta = lm_pp_pivot_get_list_rows_meta(pivot_name)
        if lr_meta and lr_meta.get("active"):
            _filled, lr_note = lm_pp_pivot_apply_list_rows(
                doc, pivot_sheet, meta=lr_meta
            )
            if lr_note:
                fmt_note = "%s; LIST_ROWS: %s" % (fmt_note, lr_note)
    except Exception as err:
        fmt_note = "%s; LIST_ROWS: %s" % (fmt_note, err)
    lm_pp_pivot_register_materialized_sheet(pivot_name, header_row)
    return pivot_sheet, "%s; %s" % (note, fmt_note)


def _lm_pp_pivot_tabular_layout_mode():
    try:
        from com.sun.star.sheet import DataPilotFieldLayoutMode as DFLM

        return int(DFLM.TABULAR_LAYOUT)
    except Exception:
        return 0


def _lm_pp_pivot_set_field_property(field, prop_name, value):
    if field is None:
        return False
    try:
        field.setPropertyValue(prop_name, value)
        return True
    except Exception:
        pass
    try:
        from com.sun.star.beans import XPropertySet

        ps = field.queryInterface(XPropertySet)
        if ps is not None:
            ps.setPropertyValue(prop_name, value)
            return True
    except Exception:
        pass
    return False


def _lm_pp_pivot_set_field_repeat_item_labels(field):
    if field is None:
        return False
    for method_name in ("setRepeatItemLabels",):
        try:
            fn = getattr(field, method_name, None)
            if callable(fn):
                fn(True)
                return True
        except Exception:
            pass
    return _lm_pp_pivot_set_field_property(field, "RepeatItemLabels", True)


def _lm_pp_pivot_set_field_tabular_layout(field):
    """
    Tabular Layout + Repeat Item Labels для поля строки/столбца.

    Без LayoutMode=TABULAR при нескольких DATA LibreOffice оставляет compact-вид
    (столбец «Data» с именами мер в строках).
    """
    if field is None:
        return False
    ok = False
    try:
        info = uno.createUnoStruct("com.sun.star.sheet.DataPilotFieldLayoutInfo")
        info.LayoutMode = _lm_pp_pivot_tabular_layout_mode()
        info.AddEmptyLines = False
        if _lm_pp_pivot_set_field_property(field, "LayoutInfo", info):
            ok = True
    except Exception:
        pass
    if _lm_pp_pivot_set_field_repeat_item_labels(field):
        ok = True
    return ok


def _lm_pp_pivot_get_data_layout_field(dp_obj, fields=None):
    """
    Псевдополе «Data» — управляет расположением нескольких полей данных.
    getDataLayoutField() не зависит от локали UI.
    """
    if dp_obj is not None:
        try:
            dlf = dp_obj.getDataLayoutField()
            if dlf is not None:
                return dlf
        except Exception:
            pass
    if fields is None and dp_obj is not None:
        try:
            fields = dp_obj.getDataPilotFields()
        except Exception:
            fields = None
    if fields is None:
        return None
    for nm in ("Data", "Данные", "data"):
        try:
            return fields.getByName(nm)
        except Exception:
            pass
    try:
        n = int(fields.getCount())
    except Exception:
        return None
    i = 0
    while i < n:
        try:
            f = fields.getByIndex(i)
            name = _lm_pp_pivot_field_display_name(f)
            if name.casefold() in ("data", "данные"):
                return f
        except Exception:
            pass
        i = i + 1
    return None


def _lm_pp_pivot_apply_table_layout(dp_obj, config, header_names=None):
    """
    Табличный макет строк/столбцов (Repeat Item Labels) и при 2+ мерах
    ориентация псевдополя Data в столбцы.

    Вызывается на дескрипторе до insertNewByName и на dp_table после refresh —
    без повторного refresh LO иногда оставляет compact, как при ручном переключении.
    """
    if dp_obj is None or config is None:
        return
    try:
        fields = dp_obj.getDataPilotFields()
    except Exception:
        return
    if fields is None:
        return

    row_fields = config.get("row_fields") or []
    for nm in row_fields:
        name = str(nm or "").strip()
        if name == "":
            continue
        field = _lm_pp_pivot_resolve_field(fields, name, header_names=header_names)
        _lm_pp_pivot_set_field_tabular_layout(field)

    column_fields = config.get("column_fields") or []
    for nm in column_fields:
        name = str(nm or "").strip()
        if name == "":
            continue
        field = _lm_pp_pivot_resolve_field(fields, name, header_names=header_names)
        _lm_pp_pivot_set_field_tabular_layout(field)

    data_count = len(config.get("data_fields") or [])
    if data_count > 1:
        dlf = _lm_pp_pivot_get_data_layout_field(dp_obj, fields=fields)
        if dlf is not None:
            orient = _lm_pp_pivot_orientation("COLUMN")
            _lm_pp_pivot_set_field_property(dlf, "Orientation", orient)


def _lm_pp_pivot_enable_repeat_item_labels( fields, row_fields, header_names=None, column_fields=None ):
    """Совместимость: Tabular Layout + Repeat Item Labels для полей строк и столбцов."""
    if fields is None:
        return
    for nm in row_fields or []:
        name = str(nm or "").strip()
        if name == "":
            continue
        field = _lm_pp_pivot_resolve_field(fields, name, header_names=header_names)
        _lm_pp_pivot_set_field_tabular_layout(field)
    for nm in column_fields or []:
        name = str(nm or "").strip()
        if name == "":
            continue
        field = _lm_pp_pivot_resolve_field(fields, name, header_names=header_names)
        _lm_pp_pivot_set_field_tabular_layout(field)


def _lm_pp_is_pivot_followup_skip_name(fn_name):
    """Шаги, которые не повторяют на материализованных сводных в финальной фазе."""
    key = str(fn_name or "").strip().casefold()
    try:
        import libre_macros_final_lib as fl

        if key in fl._LM_FINAL_SKIP_PIVOT_FOLLOWUP_KEYS:
            return True
    except ImportError:
        if key in (
            "удалить_служебные_листы",
            "удалить_листы",
            "удаление_листов",
            "скрытие_листов",
        ):
            return True
    if key == "сводная_таблица":
        return True
    return False


def _lm_pp_format_pivot_output_sheet(doc, pivot_sheet):
    """
    Оформление листа сводной: выделить значащую область,
    выравнивание по вертикали по центру, автоподбор ширины столбцов.
    """
    if doc is None or pivot_sheet is None:
        return
    try:
        sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(pivot_sheet)
        if er < sr or ec < sc:
            return
        cell_range = pivot_sheet.getCellRangeByPosition(sc, sr, ec, er)
        lm._lm_pp_select_cell_range(doc, pivot_sheet, cell_range)
        try:
            cell_range.setPropertyValue("VertJustify", VERT_CENTER)
        except Exception:
            pass
        col_indices = []
        c = int(sc)
        while c <= int(ec):
            col_indices.append(c)
            c = c + 1
        lm._lm_pp_autofit_columns(doc, pivot_sheet, col_indices, sr=sr, er=er)
    except Exception:
        pass
_LM_PP_PIVOT_ORIENTATION = {
    "HIDDEN": 0,
    "COLUMN": 1,
    "ROW": 2,
    "PAGE": 3,
    "DATA": 4,
}
_LM_PP_PIVOT_FUNCTION = {
    "SUM": 1,
    "COUNT": 2,
    "AVERAGE": 3,
    "MAX": 4,
    "MIN": 5,
    "PRODUCT": 6,
    "STDDEV": 8,
    "STD_DEV": 8,
    "VAR": 10,
}


PIVOT_CONFIG_VERSION = 1


def _lm_pp_parse_json_relaxed(text):
    """
    Разбор JSON из ячейки/диалога: строгий JSON, Python True/False/None, хвостовые запятые.
    """
    s = str(text or "").strip()
    if s == "":
        return None
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        pass
    try:
        import ast

        val = ast.literal_eval(s)
        if isinstance(val, (dict, list)):
            return val
    except Exception:
        pass
    fixed = s
    fixed = re.sub(r"\bTrue\b", "true", fixed)
    fixed = re.sub(r"\bFalse\b", "false", fixed)
    fixed = re.sub(r"\bNone\b", "null", fixed)
    fixed = re.sub(r",\s*}", "}", fixed)
    fixed = re.sub(r",\s*]", "]", fixed)
    return json.loads(fixed)


def lm_pp_pivot_default_config(field_names=None):
    cfg = {
        "v": PIVOT_CONFIG_VERSION,
        "source_sheet": "",
        "header_row": "",
        "data_start": "",
        "data_end": "",
        "sheet_name": "",
        "as_values": False,
        "list_rows_expand": False,
        "list_rows_include_names": True,
        "list_rows_name_sep": _LM_PP_PIVOT_LIST_ROWS_NAME_SEP,
        "list_rows_record_sep": _LM_PP_PIVOT_LIST_ROWS_RECORD_SEP,
        "row_fields": [],
        "column_fields": [],
        "filter_fields": [],
        "data_fields": [],
    }
    if field_names:
        for nm in field_names:
            s = str(nm).strip()
            if s != "" and s.casefold() != "data":
                cfg["data_fields"] = [{"field": s, "function": "SUM"}]
                break
    return cfg


def lm_pp_pivot_config_from_json(text):
    raw = str(text or "").strip()
    if raw == "":
        return lm_pp_pivot_default_config()
    data = _lm_pp_parse_json_relaxed(raw)
    if isinstance(data, list):
        if len(data) == 0:
            return lm_pp_pivot_default_config()
        if not isinstance(data[0], dict):
            raise ValueError("ожидается объект JSON")
        data = data[0]
    if not isinstance(data, dict):
        raise ValueError("ожидается объект JSON")
    cfg = lm_pp_pivot_default_config()
    for key in (
        "source_sheet",
        "header_row",
        "data_start",
        "data_end",
        "sheet_name",
        "as_values",
        "list_rows_expand",
        "list_rows_include_names",
        "list_rows_name_sep",
        "list_rows_record_sep",
        "row_fields",
        "column_fields",
        "filter_fields",
        "data_fields",
    ):
        if key in data:
            cfg[key] = data[key]
    if "v" in data:
        cfg["v"] = data["v"]
    cfg["data_fields"] = _lm_pp_pivot_normalize_data_fields(cfg.get("data_fields"))
    if _lm_pp_pivot_config_has_list_rows(cfg):
        cfg["as_values"] = True
    cfg["list_rows_expand"] = bool(cfg.get("list_rows_expand"))
    if "list_rows_include_names" in data:
        cfg["list_rows_include_names"] = bool(cfg.get("list_rows_include_names"))
    else:
        cfg["list_rows_include_names"] = True
    name_sep = str(cfg.get("list_rows_name_sep") or "").replace("\\n", "\n").replace("\\t", "\t")
    if name_sep == "":
        name_sep = _LM_PP_PIVOT_LIST_ROWS_NAME_SEP
    cfg["list_rows_name_sep"] = name_sep
    rec_sep = str(cfg.get("list_rows_record_sep") if cfg.get("list_rows_record_sep") is not None else "")
    rec_sep = rec_sep.replace("\\n", "\n").replace("\\t", "\t")
    if rec_sep == "" and "list_rows_record_sep" not in data:
        rec_sep = _LM_PP_PIVOT_LIST_ROWS_RECORD_SEP
    elif rec_sep == "" and data.get("list_rows_record_sep") in ("", None):
        # явно пустой — оставим пустым (склеивать без разделителя)
        rec_sep = ""
    if "list_rows_record_sep" not in data:
        cfg["list_rows_record_sep"] = _LM_PP_PIVOT_LIST_ROWS_RECORD_SEP
    else:
        cfg["list_rows_record_sep"] = rec_sep
    return cfg


def _lm_pp_pivot_is_list_rows_fn(name):
    key = str(name or "").strip().casefold().replace("-", "_")
    if key == "list_rows":
        return True
    return key in _LM_PP_PIVOT_LIST_ROWS_FN_ALIASES


def _lm_pp_pivot_normalize_fn_code(name):
    raw = str(name or "SUM").strip()
    if raw == "":
        return "SUM"
    if _lm_pp_pivot_is_list_rows_fn(raw):
        return "LIST_ROWS"
    return raw.upper().replace("-", "_")


def _lm_pp_pivot_extract_list_fields(item):
    if not isinstance(item, dict):
        return []
    for key in _LM_PP_PIVOT_LIST_FIELDS_KEYS:
        if key not in item:
            continue
        val = item.get(key)
        if isinstance(val, (list, tuple)):
            out = []
            for x in val:
                s = str(x or "").strip()
                if s != "" and s not in out:
                    out.append(s)
            return out
        s = str(val or "").strip()
        if s == "":
            return []
        parts = []
        for chunk in s.replace(";", ",").split(","):
            p = chunk.strip()
            if p != "" and p not in parts:
                parts.append(p)
        return parts
    return []


def _lm_pp_pivot_normalize_data_fields(data_fields):
    out = []
    if not isinstance(data_fields, (list, tuple)):
        return out
    for item in data_fields:
        if not isinstance(item, dict):
            continue
        field = str(item.get("field") or "").strip()
        if field == "":
            continue
        fn = _lm_pp_pivot_normalize_fn_code(item.get("function") or "SUM")
        norm = {"field": field, "function": fn}
        list_fields = _lm_pp_pivot_extract_list_fields(item)
        if fn == "LIST_ROWS" or list_fields:
            if list_fields:
                norm["list_fields"] = list_fields
        out.append(norm)
    return out


def _lm_pp_pivot_config_has_list_rows(config):
    for item in (config or {}).get("data_fields") or []:
        if isinstance(item, dict) and _lm_pp_pivot_is_list_rows_fn(
            item.get("function")
        ):
            return True
    return False


def _lm_pp_pivot_pick_count_field(list_fields, key_fields, header_names, fallback_field):
    key_set = set()
    for nm in key_fields or []:
        s = str(nm or "").strip().casefold()
        if s != "":
            key_set.add(s)
    for nm in list_fields or []:
        s = str(nm or "").strip()
        if s != "" and s.casefold() not in key_set:
            return s
    for nm in header_names or []:
        s = str(nm or "").strip()
        if s != "" and s.casefold() not in key_set and s.casefold() != "data":
            return s
    fb = str(fallback_field or "").strip()
    if fb != "":
        return fb
    if list_fields:
        return str(list_fields[0]).strip()
    if header_names:
        return str(header_names[0]).strip()
    return ""


def _lm_pp_pivot_prepare_list_rows(config, header_names=None):
    """
    LIST_ROWS → COUNT для DataPilot + meta для пост-заполнения.
    Возвращает (config_for_datapilot, list_rows_meta|None).
    """
    import copy

    cfg = copy.deepcopy(config or lm_pp_pivot_default_config())
    cfg["data_fields"] = _lm_pp_pivot_normalize_data_fields(cfg.get("data_fields"))
    data_fields = list(cfg.get("data_fields") or [])
    list_items = []
    i = 0
    while i < len(data_fields):
        item = data_fields[i]
        if isinstance(item, dict) and item.get("function") == "LIST_ROWS":
            list_items.append(item)
        i = i + 1
    if len(list_items) == 0:
        return cfg, None

    first = list_items[0]
    list_fields = list(first.get("list_fields") or [])
    if len(list_fields) == 0:
        # Без list_fields детализацию не подставляем — оставляем обычный COUNT.
        new_fields = []
        for item in data_fields:
            if not isinstance(item, dict):
                continue
            if item.get("function") == "LIST_ROWS":
                new_fields.append(
                    {
                        "field": item.get("field"),
                        "function": "COUNT",
                    }
                )
            else:
                new_fields.append(item)
        cfg["data_fields"] = new_fields
        cfg["as_values"] = True
        return cfg, {
            "active": False,
            "reason": "пустой list_fields",
            "count_field": str(first.get("field") or ""),
            "list_fields": [],
        }

    if len(list_items) > 1:
        try:
            print(
                "сводная LIST_ROWS: несколько data-полей — берём первое, остальные игнорируем"
            )
        except Exception:
            pass

    key_fields = []
    for key in ("row_fields", "column_fields", "filter_fields"):
        for nm in cfg.get(key) or []:
            s = str(nm or "").strip()
            if s != "" and s not in key_fields:
                key_fields.append(s)

    count_field = _lm_pp_pivot_pick_count_field(
        list_fields,
        key_fields,
        header_names,
        first.get("field"),
    )
    if count_field == "":
        count_field = str(first.get("field") or "").strip()

    # v1: только одно data-поле — COUNT по счётчику.
    cfg["data_fields"] = [{"field": count_field, "function": "COUNT"}]
    cfg["as_values"] = True

    try:
        max_n = int(
            (config or {}).get("list_rows_max_per_cell")
            or _LM_PP_PIVOT_LIST_ROWS_MAX_PER_CELL
        )
    except Exception:
        max_n = _LM_PP_PIVOT_LIST_ROWS_MAX_PER_CELL
    if max_n < 1:
        max_n = _LM_PP_PIVOT_LIST_ROWS_MAX_PER_CELL

    meta = {
        "active": True,
        "count_field": count_field,
        "list_fields": list_fields,
        "row_fields": list(cfg.get("row_fields") or []),
        "column_fields": list(cfg.get("column_fields") or []),
        "filter_fields": list(cfg.get("filter_fields") or []),
        "source_sheet": str(cfg.get("source_sheet") or "").strip(),
        "header_row": cfg.get("header_row"),
        "data_start": cfg.get("data_start"),
        "data_end": cfg.get("data_end"),
        "max_per_cell": max_n,
        "expand_rows": bool((config or {}).get("list_rows_expand")),
        "include_names": bool(
            True
            if (config or {}).get("list_rows_include_names") is None
            else (config or {}).get("list_rows_include_names")
        ),
        "name_sep": str(
            (config or {}).get("list_rows_name_sep")
            if (config or {}).get("list_rows_name_sep") not in (None, "")
            else _LM_PP_PIVOT_LIST_ROWS_NAME_SEP
        ),
        "record_sep": str(
            (config or {}).get("list_rows_record_sep")
            if (config or {}).get("list_rows_record_sep") is not None
            else _LM_PP_PIVOT_LIST_ROWS_RECORD_SEP
        ),
        "field_sep": _LM_PP_PIVOT_LIST_ROWS_FIELD_SEP,
    }
    return cfg, meta


def lm_pp_pivot_config_to_json(config):
    """JSON только с заданными полями (пустые data_start/data_end/header_row не пишутся)."""
    cfg = config or {}
    out = {"v": cfg.get("v") or PIVOT_CONFIG_VERSION}
    for key in ("source_sheet", "sheet_name"):
        val = str(cfg.get(key) or "").strip()
        if val != "":
            out[key] = val
    for key in ("header_row", "data_start", "data_end"):
        val = str(cfg.get(key) or "").strip()
        if val != "":
            out[key] = val
    if cfg.get("as_values"):
        out["as_values"] = True
    if cfg.get("list_rows_expand"):
        out["list_rows_expand"] = True
    if cfg.get("list_rows_include_names") is False:
        out["list_rows_include_names"] = False
    name_sep = str(cfg.get("list_rows_name_sep") or "")
    if name_sep and name_sep != _LM_PP_PIVOT_LIST_ROWS_NAME_SEP:
        out["list_rows_name_sep"] = name_sep.replace("\n", "\\n").replace("\t", "\\t")
    rec_sep = cfg.get("list_rows_record_sep")
    if rec_sep is not None and str(rec_sep) != _LM_PP_PIVOT_LIST_ROWS_RECORD_SEP:
        out["list_rows_record_sep"] = str(rec_sep).replace("\n", "\\n").replace("\t", "\\t")
    for key in ("row_fields", "column_fields", "filter_fields", "data_fields"):
        val = cfg.get(key)
        if val:
            out[key] = val
    return json.dumps(out, ensure_ascii=False, separators=(",", ":"))


def _lm_pp_pivot_sheet_end_bounds(sheet):
    """Последний столбец/строка used area (0-based); при сбое — (50, 1000)."""
    if sheet is None:
        return 50, 1000
    try:
        cursor = sheet.createCursor()
        if cursor is None:
            return 50, 1000
        cursor.gotoEndOfUsedArea(False)
        addr = cursor.getRangeAddress()
        ec = max(int(addr.EndColumn), 0)
        er = max(int(addr.EndRow), 0)
        return ec, er
    except Exception:
        return 50, 1000


def _lm_pp_pivot_header_row_index(header_row=None):
    """Строка заголовков (0-based): из JSON или 1-я строка листа."""
    hr_spec = str(header_row or "").strip()
    if hr_spec.isdigit():
        return max(int(hr_spec) - 1, 0)
    return 0


def _lm_pp_pivot_parse_corner(spec):
    """
    Угол диапазона: A1, K100 или 2:1 (строка:столбец, 1-based).
    Возвращает (col, row) 0-based или None.
    """
    s = str(spec or "").strip()
    if s == "":
        return None
    m = re.match(r"^([A-Za-z]+)(\d+)$", s, re.IGNORECASE)
    if m:
        return lm.col_letters_to_index(m.group(1)), int(m.group(2)) - 1
    m = re.match(r"^(\d+):(\d+)$", s)
    if m:
        return int(m.group(2)) - 1, int(m.group(1)) - 1
    if s.isdigit():
        return None, int(s) - 1
    return None


def _lm_pp_pivot_resolve_bounds(sheet, header_row=None, data_start=None, data_end=None):
    """
    Границы источника сводной (0-based): sc, sr, ec, er, header_row_idx.
    Пустые data_start/data_end → значащая область листа.
    """
    ua_sc, ua_sr, ua_ec, ua_er = lm._lm_vlookup_sheet_used_area(sheet)
    if ua_er < ua_sr or ua_ec < ua_sc:
        ua_ec, ua_er = _lm_pp_pivot_sheet_end_bounds(sheet)
        ua_sc, ua_sr = 0, 0

    start = _lm_pp_pivot_parse_corner(data_start)
    end = _lm_pp_pivot_parse_corner(data_end)

    if start is not None:
        sc = start[0] if start[0] is not None else ua_sc
        sr = start[1] if start[1] is not None else ua_sr
    else:
        sc, sr = ua_sc, ua_sr

    if end is not None:
        ec = end[0] if end[0] is not None else ua_ec
        er = end[1] if end[1] is not None else ua_er
    else:
        ec, er = ua_ec, ua_er

    if ec < sc:
        ec = max(ua_ec, sc + 15)
    if er < sr:
        er = max(ua_er, sr)

    hr_idx = _lm_pp_pivot_header_row_index(header_row)
    return int(sc), int(sr), int(ec), int(er), int(hr_idx)


def _lm_pp_pivot_read_header_titles(sheet, header_row_idx, sc, ec):
    """Заголовки столбцов в строке header_row_idx от sc до ec."""
    titles = []
    if sheet is None or ec < sc:
        return titles
    try:
        ra = sheet.getCellRangeByPosition(int(sc), int(header_row_idx), int(ec), int(header_row_idx))
        data = ra.getDataArray()
        if data and len(data) > 0:
            row = data[0]
            i = 0
            while i < len(row):
                val = row[i]
                if val is None:
                    titles.append("")
                else:
                    titles.append(str(val).strip())
                i = i + 1
            return titles
    except Exception:
        pass
    col = int(sc)
    while col <= int(ec):
        titles.append(lm.cell_text(sheet.getCellByPosition(col, int(header_row_idx))))
        col = col + 1
    return titles


PIVOT_COLLECT_API = 2


def lm_pp_pivot_collect_field_names( doc, source_sheet_name=None, header_row=None, data_start=None, data_end=None ):
    """
    Заголовки столбцов с листа-источника (source_sheet в JSON).

    header_row — 1-based; data_start/data_end — A1…K100 или пусто (значащая область).
    """
    if doc is None:
        return []
    sh = _lm_pp_pivot_resolve_source_sheet(doc, source_sheet_name)
    if sh is None:
        return []

    names = []
    seen = set()
    try:
        sc, sr, ec, er, hr_idx = _lm_pp_pivot_resolve_bounds(
            sh, header_row=header_row, data_start=data_start, data_end=data_end
        )
        ec = max(ec, sc + 15)
        titles = _lm_pp_pivot_read_header_titles(sh, hr_idx, sc, ec)
        i = 0
        while i < len(titles):
            title = titles[i]
            if title == "":
                title = "Колонка_%d" % (sc + i + 1)
            key = title.casefold()
            if key != "data" and key not in seen:
                seen.add(key)
                names.append(title)
            i = i + 1
    except Exception:
        pass
    return names


def _lm_pp_pivot_get_datapilot_tables(sheet):
    """DataPilotTables без queryInterface (PyUNO ломается на XDataPilotTablesSupplier)."""
    if sheet is None:
        return None
    try:
        return sheet.getDataPilotTables()
    except Exception:
        return None


def _lm_pp_pivot_resolve_source_sheet(doc, source_sheet_name):
    """Первый существующий лист по source_sheet (точное имя, N_префикс, шаблон text_match)."""
    spec = str(source_sheet_name or "").strip()
    if spec == "" or doc is None:
        return None
    parts = [p.strip() for p in re.split(r"[;,|]", spec) if p.strip()]
    if len(parts) == 0:
        return None
    try:
        sheets = doc.getSheets()
    except Exception:
        return None
    for part in parts:
        sh = lm._lm_vlookup_resolve_sheet(doc, part)
        if sh is not None:
            try:
                name = str(sh.Name)
                if not lm._lm_pp_is_param_sheet_name(name):
                    return sh
            except Exception:
                return sh
        i = 0
        while i < sheets.getCount():
            try:
                sh = sheets.getByIndex(i)
                name = str(sh.Name)
                if lm._lm_pp_is_param_sheet_name(name):
                    i = i + 1
                    continue
                if lm.text_match(name, part):
                    return sh
            except Exception:
                pass
            i = i + 1
    return None


def _lm_pp_pivot_source_sheet_matches(config, sheet_name):
    """Пустой source_sheet — все листы; иначе шаблон text_match (точное, *суффикс, префикс*)."""
    spec = str((config or {}).get("source_sheet") or "").strip()
    if spec == "":
        return True
    actual = str(sheet_name or "").strip()
    for part in re.split(r"[;,|]", spec):
        p = part.strip()
        if p != "" and lm.text_match(actual, p):
            return True
    return False


def _lm_pp_is_pivot_table_postprocess_name(fn_name):
    return str(fn_name or "").strip().lower() == "сводная_таблица"


def _lm_pp_pivot_table_extra_args(extra_raw):
    s = str(extra_raw).strip() if extra_raw is not None else ""
    if s == "":
        return []
    return [s]


def _lm_pp_pivot_orientation(name):
    key = str(name or "").strip().upper()
    try:
        from com.sun.star.sheet import DataPilotFieldOrientation as DPO

        return getattr(DPO, key)
    except Exception:
        pass
    try:
        return uno.Enum("com.sun.star.sheet.DataPilotFieldOrientation", key)
    except Exception:
        return _LM_PP_PIVOT_ORIENTATION.get(key, 0)


def _lm_pp_pivot_function(name):
    key = str(name or "SUM").strip().upper().replace("-", "_")
    if _lm_pp_pivot_is_list_rows_fn(key):
        key = "COUNT"
    try:
        from com.sun.star.sheet import GeneralFunction as GF

        if key == "STDDEV":
            key = "STD_DEV"
        return getattr(GF, key)
    except Exception:
        pass
    try:
        if key == "STDDEV":
            key = "STD_DEV"
        return uno.Enum("com.sun.star.sheet.GeneralFunction", key)
    except Exception:
        return _LM_PP_PIVOT_FUNCTION.get(key, 1)


def _lm_pp_pivot_parse_config(extra_args):
    if not extra_args:
        return None
    if len(extra_args) == 1:
        text = str(extra_args[0]).strip()
    else:
        # Ячейка C ошибочно разбита parse_pp_extra_args по запятым — склеить обратно.
        text = ",".join(str(x) for x in extra_args).strip()
    if text == "":
        return None
    return lm_pp_pivot_config_from_json(text)


def _lm_pp_pivot_header_names_from_range(source_sheet, source_range):
    """Заголовки столбцов из первой строки диапазона источника."""
    if source_sheet is None or source_range is None:
        return []
    try:
        sc, sr, ec, er = lm._lm_pp_range_address(source_range)
        if er < sr or ec < sc:
            return []
        return _lm_pp_pivot_read_header_titles(source_sheet, sr, sc, ec)
    except Exception:
        return []


def _lm_pp_pivot_field_display_name(field_obj):
    if field_obj is None:
        return ""
    for attr in ("getName",):
        try:
            fn = getattr(field_obj, attr, None)
            if callable(fn):
                s = str(fn()).strip()
                if s != "":
                    return s
        except Exception:
            pass
    try:
        return str(field_obj.Name).strip()
    except Exception:
        return ""


def _lm_pp_pivot_list_descriptor_fields(fields):
    out = []
    if fields is None:
        return out
    try:
        n = int(fields.getCount())
    except Exception:
        return out
    i = 0
    while i < n:
        try:
            out.append(_lm_pp_pivot_field_display_name(fields.getByIndex(i)))
        except Exception:
            out.append("")
        i = i + 1
    return out


def _lm_pp_pivot_col_index(header_names, field_name):
    key = str(field_name or "").strip().casefold()
    if key == "" or not header_names:
        return -1
    j = 0
    while j < len(header_names):
        if str(header_names[j] or "").strip().casefold() == key:
            return j
        j = j + 1
    return -1


def _lm_pp_pivot_resolve_field(fields, field_name, header_names=None):
    """Поле DataPilot по имени заголовка (getByName или по индексу столбца)."""
    name = str(field_name or "").strip()
    if name == "" or fields is None:
        return None
    try:
        return fields.getByName(name)
    except Exception:
        pass
    key = name.casefold()
    try:
        n = int(fields.getCount())
    except Exception:
        n = 0
    i = 0
    while i < n:
        try:
            f = fields.getByIndex(i)
            fn = _lm_pp_pivot_field_display_name(f)
            if fn.casefold() == key:
                return f
        except Exception:
            pass
        i = i + 1
    if header_names:
        j = 0
        while j < len(header_names):
            if str(header_names[j] or "").strip().casefold() == key:
                try:
                    return fields.getByIndex(j)
                except Exception:
                    pass
            j = j + 1
    return None


def _lm_pp_pivot_sheet_replace_if_old_output(doc, sheet_name):
    """
    Перед повторной сводной: удалить лист с тем же именем, если это старая пустая сводная.
    Иначе lm._lm_pp_unique_sheet_name даёт суффикс _2, _3…
    """
    if doc is None:
        return
    name = str(sheet_name or "").strip()[:31]
    if name == "":
        return
    try:
        sheets = doc.getSheets()
        if not sheets.hasByName(name):
            return
        sh = sheets.getByName(name)
        dp = _lm_pp_pivot_get_datapilot_tables(sh)
        has_dp = dp is not None and int(dp.getCount()) > 0
        sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(sh)
        emptyish = er < sr
        if has_dp or emptyish:
            sheets.removeByName(name)
    except Exception:
        pass


def _lm_pp_pivot_set_field(fields, field_name, orientation_key, func_name=None, header_names=None):
    if fields is None:
        return False
    name = str(field_name or "").strip()
    if name == "" or name.casefold() == "data":
        return False
    field = None
    if header_names:
        col_idx = _lm_pp_pivot_col_index(header_names, name)
        if col_idx >= 0:
            try:
                field = fields.getByIndex(col_idx)
            except Exception:
                field = None
    if field is None:
        field = _lm_pp_pivot_resolve_field(fields, name, header_names=header_names)
    if field is None:
        return False
    orient = _lm_pp_pivot_orientation(orientation_key)
    try:
        field.setPropertyValue("Orientation", orient)
    except Exception:
        try:
            from com.sun.star.beans import XPropertySet

            ps = field.queryInterface(XPropertySet)
            if ps is not None:
                ps.setPropertyValue("Orientation", orient)
            else:
                return False
        except Exception:
            return False
    if func_name is not None and orientation_key == "DATA":
        fn = _lm_pp_pivot_function(func_name)
        try:
            field.setPropertyValue("Function", fn)
        except Exception:
            try:
                field.setPropertyValue("Function", int(fn))
            except Exception:
                pass
    return True


def _lm_pp_pivot_sheet_index(doc, sheet):
    """Индекс листа в книге (0-based); сравнение по Name — PyUNO «==» ненадёжен."""
    if doc is None or sheet is None:
        return -1
    try:
        want = str(sheet.Name)
        sheets = doc.getSheets()
        i = 0
        while i < sheets.getCount():
            sh = sheets.getByIndex(i)
            if str(sh.Name) == want:
                return i
            i = i + 1
    except Exception:
        pass
    return -1


def _lm_pp_pivot_format_addr(addr):
    if addr is None:
        return "?"
    try:
        return "S%d C%d..%d R%d..%d" % (
            int(addr.Sheet),
            int(addr.StartColumn),
            int(addr.EndColumn),
            int(addr.StartRow),
            int(addr.EndRow),
        )
    except Exception:
        return str(addr)


def _lm_pp_pivot_cell_has_content(cell):
    if cell is None:
        return False
    try:
        if int(cell.getType().value) == 0:
            return False
    except Exception:
        pass
    try:
        if str(cell.getString() or "").strip() != "":
            return True
    except Exception:
        pass
    try:
        val = cell.getValue()
        if val is not None and val != 0.0:
            return True
    except Exception:
        pass
    try:
        if str(cell.getFormula() or "").strip() != "":
            return True
    except Exception:
        pass
    return False


def _lm_pp_pivot_output_has_content(pivot_sheet, dp_table):
    if pivot_sheet is None or dp_table is None:
        return False
    try:
        out = dp_table.getOutputRange()
        if out.EndRow < out.StartRow or out.EndColumn < out.StartColumn:
            return False
        r = int(out.StartRow)
        while r <= int(out.EndRow):
            c = int(out.StartColumn)
            while c <= int(out.EndColumn):
                if _lm_pp_pivot_cell_has_content(
                    pivot_sheet.getCellByPosition(c, r)
                ):
                    return True
                c = c + 1
            r = r + 1
    except Exception:
        return False
    return False


def _lm_pp_pivot_create_table(doc, source_sheet, source_range, config):
    """
    Создать сводную таблицу на новом листе.

    config — dict из JSON (диалог «Параметры» / lm_pp_pivot_config_from_json).
    Возвращает (pivot_sheet, table_name, note) или (None, "", error).
    """
    if doc is None or source_sheet is None or source_range is None:
        return None, "", "нет документа или листа-источника"
    if config is None:
        return None, "", "пустая конфигурация"

    header_names = _lm_pp_pivot_header_names_from_range(source_sheet, source_range)
    config, list_rows_meta = _lm_pp_pivot_prepare_list_rows(config, header_names)
    if list_rows_meta is not None and not list_rows_meta.get("active"):
        try:
            print(
                "сводная LIST_ROWS: детализация отключена (%s)"
                % (list_rows_meta.get("reason") or "—")
            )
        except Exception:
            pass
        list_rows_meta = None

    data_fields = config.get("data_fields") or []
    if len(data_fields) == 0:
        return None, "", "не заданы поля данных"

    dp_src = _lm_pp_pivot_get_datapilot_tables(source_sheet)
    if dp_src is None:
        return None, "", "getDataPilotTables недоступен на листе-источнике"

    try:
        dp_desc = dp_src.createDataPilotDescriptor()
        src_addr = source_range.getRangeAddress()
        sh_idx = _lm_pp_pivot_sheet_index(doc, source_sheet)
        if sh_idx < 0:
            return (
                None,
                "",
                "лист-источник «%s» не найден в книге"
                % (getattr(source_sheet, "Name", "?")),
            )
        src_addr.Sheet = sh_idx
        dp_desc.setSourceRange(src_addr)
        try:
            dp_desc.setIgnoreEmptyRows(False)
        except Exception:
            pass
        fields = dp_desc.getDataPilotFields()
    except Exception as err:
        return None, "", "дескриптор: %s" % err

    set_data = 0
    set_row = 0
    set_col = 0

    for nm in config.get("filter_fields") or []:
        _lm_pp_pivot_set_field(
            fields, nm, "PAGE", header_names=header_names
        )
    for nm in config.get("column_fields") or []:
        if _lm_pp_pivot_set_field(
            fields, nm, "COLUMN", header_names=header_names
        ):
            set_col = set_col + 1
    for nm in config.get("row_fields") or []:
        if _lm_pp_pivot_set_field(
            fields, nm, "ROW", header_names=header_names
        ):
            set_row = set_row + 1

    for item in data_fields:
        if isinstance(item, dict):
            if _lm_pp_pivot_set_field(
                fields,
                item.get("field"),
                "DATA",
                item.get("function") or "SUM",
                header_names=header_names,
            ):
                set_data = set_data + 1

    # Tabular + Data→COLUMN только после назначения всех полей данных.
    try:
        _lm_pp_pivot_apply_table_layout(dp_desc, config, header_names=header_names)
    except Exception:
        pass

    if set_data == 0:
        avail = _lm_pp_pivot_list_descriptor_fields(fields)
        hdr = ", ".join(header_names[:20]) if header_names else "—"
        return (
            None,
            "",
            "поля данных не сопоставлены; заголовки [%s]; поля DP [%s]"
            % (hdr, ", ".join(avail[:20])),
        )

    row_wanted = len(config.get("row_fields") or [])
    col_wanted = len(config.get("column_fields") or [])
    if row_wanted > 0 and set_row == 0:
        hdr = ", ".join(header_names[:20]) if header_names else "—"
        return (
            None,
            "",
            "поля строк не сопоставлены; заголовки [%s]; источник %s"
            % (hdr, _lm_pp_pivot_format_addr(src_addr)),
        )
    if col_wanted > 0 and set_col == 0:
        hdr = ", ".join(header_names[:20]) if header_names else "—"
        return (
            None,
            "",
            "поля столбцов не сопоставлены; заголовки [%s]; источник %s"
            % (hdr, _lm_pp_pivot_format_addr(src_addr)),
        )

    base_sheet_name = str(config.get("sheet_name") or "").strip()
    if base_sheet_name == "":
        base_sheet_name = "Сводная"
    _lm_pp_pivot_sheet_replace_if_old_output(doc, base_sheet_name)
    temp_name = lm._lm_pp_unique_sheet_name(doc, base_sheet_name)
    insert_pos = int(doc.Sheets.getCount())
    try:
        doc.Sheets.insertNewByName(temp_name, insert_pos)
    except Exception as err:
        return None, "", "не удалось создать лист: %s" % err
    pivot_sheet = doc.Sheets.getByName(temp_name)
    if pivot_sheet is None:
        return None, "", "лист сводной не найден после создания"

    dp_tables = _lm_pp_pivot_get_datapilot_tables(pivot_sheet)
    if dp_tables is None:
        try:
            doc.Sheets.removeByName(temp_name)
        except Exception:
            pass
        return None, "", "getDataPilotTables недоступен на листе сводной"

    table_name = "Pivot_PP_%d" % (insert_pos + 1)
    suffix = 2
    while True:
        try:
            if not dp_tables.hasByName(table_name):
                break
        except Exception:
            break
        table_name = "Pivot_PP_%d" % (insert_pos + suffix)
        suffix = suffix + 1
        if suffix > 999:
            break

    dest = uno.createUnoStruct("com.sun.star.table.CellAddress")
    dest_idx = _lm_pp_pivot_sheet_index(doc, pivot_sheet)
    if dest_idx < 0:
        try:
            doc.Sheets.removeByName(temp_name)
        except Exception:
            pass
        return None, "", "лист сводной не найден в книге после создания"
    dest.Sheet = dest_idx
    dest.Column = 0
    dest.Row = 0

    try:
        dp_tables.insertNewByName(table_name, dest, dp_desc)
    except Exception as err:
        try:
            doc.Sheets.removeByName(temp_name)
        except Exception:
            pass
        return None, "", "insertNewByName: %s" % err

    try:
        dp_table = dp_tables.getByName(table_name)
        try:
            doc.calculateAll()
        except Exception:
            pass
        dp_table.refresh()
        dp_table.refresh()
    except Exception:
        try:
            fresh = _lm_pp_pivot_get_datapilot_tables(pivot_sheet)
            if fresh is not None:
                dp_table = fresh.getByName(table_name)
                try:
                    doc.calculateAll()
                except Exception:
                    pass
                dp_table.refresh()
        except Exception:
            dp_table = None

    if dp_table is not None:
        try:
            _lm_pp_pivot_apply_table_layout(
                dp_table, config, header_names=header_names
            )
            try:
                doc.calculateAll()
            except Exception:
                pass
            dp_table.refresh()
        except Exception:
            pass

    if dp_table is not None:
        try:
            dp_table.gotoStart()
        except Exception:
            pass
        if not _lm_pp_pivot_output_has_content(pivot_sheet, dp_table):
            try:
                doc.Sheets.removeByName(pivot_sheet.Name)
            except Exception:
                pass
            return (
                None,
                "",
                "сводная без данных; источник %s; лист «%s»; поля данных %d, строк %d"
                % (
                    _lm_pp_pivot_format_addr(src_addr),
                    str(getattr(source_sheet, "Name", "?")),
                    set_data,
                    set_row,
                ),
            )
        try:
            out = dp_table.getOutputRange()
            if out.EndRow < out.StartRow or out.EndColumn < out.StartColumn:
                try:
                    doc.Sheets.removeByName(pivot_sheet.Name)
                except Exception:
                    pass
                return (
                    None,
                    "",
                    "сводная без вывода после refresh; проверьте поля и диапазон",
                )
        except Exception:
            pass

    target_sheet_name = str(config.get("sheet_name") or "").strip()
    if target_sheet_name == "":
        target_sheet_name = "Сводная"

    if target_sheet_name != "" and pivot_sheet.Name != target_sheet_name:
        final_name = lm._lm_pp_unique_sheet_name(doc, target_sheet_name, exclude_sheet=pivot_sheet)
        if final_name != "" and final_name != pivot_sheet.Name:
            try:
                pivot_sheet.Name = final_name
                pivot_sheet = doc.Sheets.getByName(final_name)
            except Exception:
                pass

    _lm_pp_format_pivot_output_sheet(doc, pivot_sheet)

    mat_suffix = ""
    if config.get("as_values"):
        # source_sheet в meta — фактический лист данных (для пост-заполнения LIST_ROWS).
        if list_rows_meta is not None:
            try:
                list_rows_meta["source_sheet"] = str(
                    getattr(source_sheet, "Name", "") or ""
                ).strip() or list_rows_meta.get("source_sheet")
            except Exception:
                pass
        lm_pp_pivot_register_as_values_pending(
            pivot_sheet.Name, list_rows_meta=list_rows_meta
        )
        if list_rows_meta is not None and list_rows_meta.get("active"):
            mat_suffix = "; as_values→финал; LIST_ROWS→COUNT"
        else:
            mat_suffix = "; as_values→финал"

    return (
        pivot_sheet,
        table_name,
        "ok%s; источник %s; поля данных %d, строк %d"
        % (
            mat_suffix,
            _lm_pp_pivot_format_addr(src_addr),
            set_data,
            set_row,
        ),
    )


def lm_pp_pivot_step_applies_to_sheet(extra_raw, sheet_name):
    """Нужно ли выполнять шаг «сводная_таблица» на данном листе (до вызова колбэка)."""
    s = str(extra_raw or "").strip()
    if s == "":
        return True
    try:
        cfg = lm_pp_pivot_config_from_json(s)
    except Exception:
        return True
    if cfg is None:
        return True
    return _lm_pp_pivot_source_sheet_matches(cfg, sheet_name)


def lm_pp_range_pivot_table(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Построить сводную таблицу (DataPilot) по диапазону результата.

    Колонка C — JSON макета (диалог «Параметры» в визарде).
    source_sheet — применить только к листу результата (шаблон как в фильтре «Листы»).
    Сводная создаётся на новом листе; имя листа и флаг as_values — в JSON.

    Карта: «сводная_таблица».
    """
    sheet_name = sheet.Name if sheet is not None else ""
    try:
        config = _lm_pp_pivot_parse_config(extra_args)
    except Exception as err:
        lm._lm_log_postprocess(
            doc,
            sheet_name,
            "сводная_таблица",
            "—",
            "ошибка",
            "JSON: %s" % err,
        )
        return
    if config is None:
        lm._lm_log_postprocess(
            doc,
            sheet_name,
            "сводная_таблица",
            "—",
            "пропуск",
            "пустая колонка C",
        )
        return
    if not _lm_pp_pivot_source_sheet_matches(config, sheet_name):
        lm._lm_log_postprocess(
            doc,
            sheet_name,
            "сводная_таблица",
            "—",
            "пропуск",
            "лист не совпадает с source_sheet (%s)"
            % (config.get("source_sheet") or ""),
        )
        return

    hsc, hsr, hec, her = lm._lm_pp_range_address(header_row_range)
    sc, sr, ec, er = lm._lm_pp_range_address(data_range)
    cfg_start = str(config.get("data_start") or "").strip()
    cfg_end = str(config.get("data_end") or "").strip()
    cfg_hdr = str(config.get("header_row") or "").strip()
    if cfg_start != "" or cfg_end != "" or cfg_hdr != "":
        b_sc, b_sr, b_ec, b_er, b_hr = _lm_pp_pivot_resolve_bounds(
            sheet,
            header_row=config.get("header_row"),
            data_start=config.get("data_start"),
            data_end=config.get("data_end"),
        )
        c0 = b_sc
        c1 = b_ec
        r0 = min(b_sr, b_hr)
        r1 = b_er
    else:
        c0 = min(sc, hsc)
        c1 = max(ec, hec)
        r0 = min(sr, hsr)
        r1 = max(er, her)
    if r1 < r0 or c1 < c0:
        lm._lm_log_postprocess(
            doc,
            sheet_name,
            "сводная_таблица",
            "—",
            "пропуск",
            "нет данных",
        )
        return

    source_range = sheet.getCellRangeByPosition(c0, r0, c1, r1)
    pivot_sheet, table_name, note = _lm_pp_pivot_create_table(
        doc, sheet, source_range, config
    )
    if pivot_sheet is None:
        lm._lm_log_postprocess(
            doc,
            sheet_name,
            "сводная_таблица",
            "—",
            "ошибка",
            note,
        )
        return
    lm._lm_log_postprocess(
        doc,
        pivot_sheet.Name,
        "сводная_таблица",
        str(config.get("sheet_name") or ""),
        "ok",
        "таблица %s; %s" % (table_name, note),
    )


def _lm_pp_pivot_list_rows_cell_text(cell):
    """Отображаемый текст ячейки для сопоставления ключей / детализации."""
    try:
        return str(lm.cell_text(cell) or "").strip()
    except Exception:
        pass
    try:
        return str(cell.getString() or "").strip()
    except Exception:
        return ""


def _lm_pp_pivot_list_rows_format_detail_line( header_names, row_vals, list_fields, include_names=True, name_sep=None, field_sep=None):
    if name_sep is None:
        name_sep = _LM_PP_PIVOT_LIST_ROWS_NAME_SEP
    if field_sep is None:
        field_sep = _LM_PP_PIVOT_LIST_ROWS_FIELD_SEP
    parts = []
    i = 0
    while i < len(list_fields):
        nm = str(list_fields[i] or "").strip()
        i = i + 1
        if nm == "":
            continue
        idx = _lm_pp_pivot_col_index(header_names, nm)
        val = ""
        if idx >= 0 and idx < len(row_vals):
            val = str(row_vals[idx] or "").strip()
        if include_names:
            parts.append("%s%s%s" % (nm, name_sep, val))
        else:
            parts.append(val)
    return field_sep.join(parts)


def _lm_pp_pivot_list_rows_build_index( source_sheet, header_names, sc, data_sr, ec, er, key_fields, list_fields, include_names=True, name_sep=None, field_sep=None):
    """
    Индекс tuple(ключ) → список строк детализации (порядок источника).
    key_fields — имена колонок источника (row+column; filter v1 не включаем).
    """
    key_idxs = []
    for nm in key_fields or []:
        key_idxs.append(_lm_pp_pivot_col_index(header_names, nm))
    index = {}
    r = int(data_sr)
    while r <= int(er):
        row_vals = []
        c = int(sc)
        while c <= int(ec):
            row_vals.append(
                _lm_pp_pivot_list_rows_cell_text(
                    source_sheet.getCellByPosition(c, r)
                )
            )
            c = c + 1
        key_parts = []
        ki = 0
        while ki < len(key_idxs):
            idx = key_idxs[ki]
            ki = ki + 1
            if idx < 0 or idx >= len(row_vals):
                key_parts.append("")
            else:
                key_parts.append(row_vals[idx])
        key = tuple(key_parts)
        line = _lm_pp_pivot_list_rows_format_detail_line(
            header_names,
            row_vals,
            list_fields,
            include_names=include_names,
            name_sep=name_sep,
            field_sep=field_sep,
        )
        if key not in index:
            index[key] = []
        index[key].append(line)
        r = r + 1
    return index


def _lm_pp_pivot_list_rows_column_key_from_header(header_text, column_fields):
    """Ключ столбцов из заголовка data-ячейки после материализации."""
    n = len(column_fields or [])
    if n == 0:
        return ()
    t = str(header_text or "").strip()
    if t == "":
        return tuple([""] * n)
    # «Метка: значение» / несколько через « \ » или « | »
    if ":" in t:
        chunks = []
        for piece in t.replace(_LM_PP_PIVOT_HEADER_COLLAPSE_DELIM, "|").split("|"):
            p = piece.strip()
            if ":" in p:
                chunks.append(p.split(":", 1)[1].strip())
            elif p != "":
                chunks.append(p)
        if len(chunks) == 0:
            chunks = [t.split(":", 1)[-1].strip()]
        while len(chunks) < n:
            chunks.append("")
        return tuple(chunks[:n])
    if n == 1:
        return (t,)
    return tuple(([t] + [""] * (n - 1))[:n])


def _lm_pp_pivot_list_rows_join_lines(lines, max_n, record_sep=None):
    if not lines:
        return ""
    if record_sep is None:
        record_sep = _LM_PP_PIVOT_LIST_ROWS_RECORD_SEP
    max_n = int(max_n or _LM_PP_PIVOT_LIST_ROWS_MAX_PER_CELL)
    if max_n < 1:
        max_n = _LM_PP_PIVOT_LIST_ROWS_MAX_PER_CELL
    if len(lines) <= max_n:
        return record_sep.join(lines)
    shown = lines[:max_n]
    rest = len(lines) - max_n
    shown.append(u"… ещё %d строк" % rest)
    return record_sep.join(shown)


def _lm_pp_pivot_list_rows_read_count(cell):
    """COUNT из ячейки; 0 / пусто → 0."""
    if cell is None:
        return 0
    try:
        s = str(cell.getString() or "").strip()
    except Exception:
        s = ""
    if s != "":
        try:
            # локаль: «1 234» / «1,5»
            cleaned = s.replace("\u00a0", " ").replace(" ", "").replace(",", ".")
            return float(cleaned)
        except Exception:
            pass
    try:
        return float(cell.getValue())
    except Exception:
        return 0


def _lm_pp_pivot_list_rows_set_cell_text(cell, text):
    if cell is None:
        return False
    try:
        cell.setString(text)
        return True
    except Exception:
        pass
    try:
        cell.String = text
        return True
    except Exception:
        return False


def _lm_pp_pivot_list_rows_rename_data_headers( pivot_sheet, data_c0, pec, header_row, count_field="", force_all=False ):
    """Заголовок области данных: «Количество — …» → «Список_строк»."""
    label = _LM_PP_PIVOT_LIST_ROWS_HEADER
    cf = str(count_field or "").strip().casefold()
    c = int(data_c0)
    while c <= int(pec):
        cell = pivot_sheet.getCellByPosition(c, int(header_row))
        hdr = _lm_pp_pivot_list_rows_cell_text(cell)
        if _lm_pp_pivot_cell_looks_like_grand_total_label(hdr):
            c = c + 1
            continue
        low = hdr.casefold()
        rename = bool(force_all)
        if not rename:
            if "количество" in low or "count" in low or "список_строк" in low:
                rename = True
            elif cf != "" and cf in low:
                rename = True
            elif hdr == "":
                rename = True
        if rename:
            _lm_pp_pivot_list_rows_set_cell_text(cell, label)
        c = c + 1


def _lm_pp_pivot_list_rows_write_expand_groups( pivot_sheet, groups, psc, data_c0, pec, max_n, record_sep=None ):
    """
    Разворот групп вниз: ключи только в первой строке группы,
    каждая запись списка — отдельная строка листа.
    groups: [{r, col_lines: {col: [str,...]}}, ...] сверху вниз.
    """
    filled = 0
    i = len(groups) - 1
    while i >= 0:
        g = groups[i]
        i = i - 1
        col_lines = g.get("col_lines") or {}
        n = 0
        for lines in col_lines.values():
            if len(lines) > n:
                n = len(lines)
        if n > int(max_n):
            n = int(max_n)
        if n <= 0:
            continue
        base_r = int(g["r"])
        if n > 1:
            try:
                pivot_sheet.getRows().insertByIndex(base_r + 1, n - 1)
            except Exception:
                try:
                    pivot_sheet.Rows.insertByIndex(base_r + 1, n - 1)
                except Exception:
                    # fallback: одна ячейка
                    for c, lines in col_lines.items():
                        text = _lm_pp_pivot_list_rows_join_lines(
                            lines, max_n, record_sep=record_sep
                        )
                        _lm_pp_pivot_list_rows_set_cell_text(
                            pivot_sheet.getCellByPosition(int(c), base_r), text
                        )
                        filled = filled + 1
                    continue
        li = 0
        while li < n:
            row = base_r + li
            if li > 0:
                c = int(psc)
                while c < int(data_c0):
                    _lm_pp_pivot_list_rows_set_cell_text(
                        pivot_sheet.getCellByPosition(c, row), ""
                    )
                    c = c + 1
            for c, lines in col_lines.items():
                text = ""
                if li < len(lines):
                    text = lines[li]
                    if li == int(max_n) - 1 and len(lines) > int(max_n):
                        text = text  # join already truncated via slice below
                _lm_pp_pivot_list_rows_set_cell_text(
                    pivot_sheet.getCellByPosition(int(c), row), text
                )
                if text != "":
                    filled = filled + 1
            li = li + 1
        # хвост «… ещё N» в последней строке, если обрезали
        for c, lines in col_lines.items():
            if len(lines) > int(max_n):
                rest = len(lines) - int(max_n)
                last = pivot_sheet.getCellByPosition(int(c), base_r + n - 1)
                cur = _lm_pp_pivot_list_rows_cell_text(last)
                if cur != "":
                    _lm_pp_pivot_list_rows_set_cell_text(
                        last, cur + " … ещё %d строк" % rest
                    )
    return filled


def lm_pp_pivot_apply_list_rows(doc, pivot_sheet, meta=None):
    """
    После as_values: заменить COUNT на Список_строк.
    meta.expand_rows — каждая запись на отдельной строке листа;
    иначе — многострочная ячейка.
    Возвращает (filled_count, note).
    """
    if pivot_sheet is None:
        return 0, "нет листа сводной"
    if meta is None:
        meta = lm_pp_pivot_get_list_rows_meta(getattr(pivot_sheet, "Name", ""))
    if not meta or not meta.get("active"):
        return 0, "нет LIST_ROWS meta"

    list_fields = list(meta.get("list_fields") or [])
    if len(list_fields) == 0:
        return 0, "пустой list_fields"

    src_name = str(meta.get("source_sheet") or "").strip()
    source_sheet = _lm_pp_pivot_resolve_source_sheet(doc, src_name)
    if source_sheet is None:
        return 0, "лист-источник «%s» не найден" % (src_name or "—")

    sc0, sr0, ec0, er0, hdr_idx = _lm_pp_pivot_resolve_bounds(
        source_sheet,
        header_row=meta.get("header_row"),
        data_start=meta.get("data_start"),
        data_end=meta.get("data_end"),
    )
    header_names = _lm_pp_pivot_read_header_titles(source_sheet, hdr_idx, sc0, ec0)
    data_sr = max(int(hdr_idx) + 1, int(sr0))
    if er0 < data_sr:
        return 0, "нет строк данных в источнике"

    row_fields = [str(x).strip() for x in (meta.get("row_fields") or []) if str(x).strip()]
    column_fields = [
        str(x).strip() for x in (meta.get("column_fields") or []) if str(x).strip()
    ]
    key_fields = list(row_fields) + list(column_fields)
    if len(meta.get("filter_fields") or []) > 0:
        try:
            print(
                "сводная LIST_ROWS: filter_fields не учтены при выборке источника"
            )
        except Exception:
            pass

    index = _lm_pp_pivot_list_rows_build_index(
        source_sheet,
        header_names,
        sc0,
        data_sr,
        ec0,
        er0,
        key_fields,
        list_fields,
        include_names=bool(meta.get("include_names", True)),
        name_sep=meta.get("name_sep"),
        field_sep=meta.get("field_sep"),
    )

    psc, psr, pec, per = lm._lm_vlookup_sheet_used_area(pivot_sheet)
    if per < psr or pec < psc:
        return 0, "пустой лист сводной"

    config_hint = {
        "row_fields": row_fields,
        "column_fields": column_fields,
    }
    header_row = _lm_pp_pivot_detect_data_header_row(
        pivot_sheet, psc, psr, pec, per, config_hint
    )
    R = _lm_pp_pivot_detect_row_field_cols_after_collapse(
        pivot_sheet, psc, pec, header_row, config=config_hint
    )
    if R <= 0 and len(row_fields) > 0:
        R = len(row_fields)
    data_c0 = int(psc) + max(int(R), 0)
    if data_c0 > int(pec) or int(header_row) >= int(per):
        return 0, "нет области данных на сводной"

    expand_rows = bool(meta.get("expand_rows"))
    max_n = int(meta.get("max_per_cell") or _LM_PP_PIVOT_LIST_ROWS_MAX_PER_CELL)
    count_field = str(meta.get("count_field") or "").strip()
    record_sep = meta.get("record_sep")
    if record_sep is None:
        record_sep = _LM_PP_PIVOT_LIST_ROWS_RECORD_SEP

    groups = []
    warned = 0
    r = int(header_row) + 1
    while r <= int(per):
        if _lm_pp_pivot_row_looks_like_grand_total(pivot_sheet, psc, r):
            r = r + 1
            continue
        row_key_parts = []
        c = int(psc)
        while c < data_c0:
            row_key_parts.append(
                _lm_pp_pivot_list_rows_cell_text(
                    pivot_sheet.getCellByPosition(c, r)
                )
            )
            c = c + 1
        if len(row_key_parts) > 0:
            last = row_key_parts[0]
            i = 1
            while i < len(row_key_parts):
                if row_key_parts[i] == "" and last != "":
                    row_key_parts[i] = last
                elif row_key_parts[i] != "":
                    last = row_key_parts[i]
                i = i + 1

        col_lines = {}
        c = data_c0
        while c <= int(pec):
            cell = pivot_sheet.getCellByPosition(c, r)
            n_count = _lm_pp_pivot_list_rows_read_count(cell)
            if n_count <= 0:
                c = c + 1
                continue
            hdr = _lm_pp_pivot_list_rows_cell_text(
                pivot_sheet.getCellByPosition(c, header_row)
            )
            if _lm_pp_pivot_cell_looks_like_grand_total_label(hdr):
                c = c + 1
                continue
            col_key = _lm_pp_pivot_list_rows_column_key_from_header(
                hdr, column_fields
            )
            full_key = tuple(row_key_parts) + tuple(col_key)
            if len(column_fields) == 0:
                full_key = tuple(row_key_parts)
            lines = list(index.get(full_key) or [])
            if len(lines) == 0:
                warned = warned + 1
                c = c + 1
                continue
            if abs(len(lines) - int(n_count)) > 0 and warned < 5:
                try:
                    print(
                        "сводная LIST_ROWS: COUNT=%s строк источника=%s ключ=%s"
                        % (int(n_count), len(lines), full_key)
                    )
                except Exception:
                    pass
                warned = warned + 1
            col_lines[c] = lines
            c = c + 1

        if len(col_lines) > 0:
            groups.append({"r": r, "col_lines": col_lines})
        r = r + 1

    filled = 0
    if expand_rows:
        filled = _lm_pp_pivot_list_rows_write_expand_groups(
            pivot_sheet, groups, psc, data_c0, pec, max_n, record_sep=record_sep
        )
        mode_note = "разворот строк"
    else:
        gi = 0
        while gi < len(groups):
            g = groups[gi]
            gi = gi + 1
            base_r = int(g["r"])
            for c, lines in (g.get("col_lines") or {}).items():
                text = _lm_pp_pivot_list_rows_join_lines(
                    lines, max_n, record_sep=record_sep
                )
                cell = pivot_sheet.getCellByPosition(int(c), base_r)
                if not _lm_pp_pivot_list_rows_set_cell_text(cell, text):
                    continue
                try:
                    cell.IsTextWrapped = True
                except Exception:
                    pass
                filled = filled + 1
        mode_note = "в ячейке"

    # Заголовок «Список_строк» — после сопоставления ключей столбцов.
    _lm_pp_pivot_list_rows_rename_data_headers(
        pivot_sheet,
        data_c0,
        pec,
        header_row,
        count_field=count_field,
        force_all=(len(column_fields) == 0),
    )

    # Область после возможного insert — перечитать used area.
    psc2, psr2, pec2, per2 = lm._lm_vlookup_sheet_used_area(pivot_sheet)
    if filled > 0 and header_row < per2 and data_c0 <= pec2:
        try:
            data_range = pivot_sheet.getCellRangeByPosition(
                data_c0, int(header_row) + 1, int(pec2), int(per2)
            )
            header_range = pivot_sheet.getCellRangeByPosition(
                int(psc2), int(header_row), int(pec2), int(header_row)
            )
            if not expand_rows:
                try:
                    data_range.IsTextWrapped = True
                except Exception:
                    pass
                try:
                    lm.lm_pp_range_word_wrap(
                        doc, pivot_sheet, data_range, header_range, "да", "заголовок"
                    )
                except Exception:
                    pass
            try:
                lm.lm_pp_range_autofit_row_heights(
                    doc, pivot_sheet, data_range, header_range
                )
            except Exception:
                pass
            # сетка на расширенную область
            if expand_rows:
                try:
                    lm._lm_pp_apply_table_grid(
                        data_range,
                        doc=doc,
                        sheet=pivot_sheet,
                        grid_label="тонкая_сетка",
                        header_row_range=header_range,
                    )
                except Exception:
                    pass
        except Exception:
            pass

    note = "заполнено %d (%s)" % (filled, mode_note)
    if warned > 0:
        note = note + "; предупреждений %d" % warned
    return filled, note
