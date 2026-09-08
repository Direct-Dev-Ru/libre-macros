# -*- coding: utf-8 -*-
from __future__ import print_function
"""
Общая библиотека постобработки LibreOffice Calc для макросов libre-macros.

Функции уровня диапазона: lm_pp_range_*
Функции уровня строки:     lm_pp_row_*
Сводные таблицы:          libre_macros_pivot_lib (re-export lm_pp_pivot_*, lm_pp_range_pivot_table)

"""
MACRO_VERSION = "3.10.701"
# Подробные логи исполнения постобработки (раскраска, границы и т.д.).
LIBRE_MACROS_DEBUG = False
import json
import re
import time

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table.CellHoriJustify import CENTER as HORI_CENTER
from com.sun.star.table.CellHoriJustify import LEFT as HORI_LEFT
from com.sun.star.table.CellHoriJustify import RIGHT as HORI_RIGHT
from com.sun.star.table.CellVertJustify import BOTTOM as VERT_BOTTOM
from com.sun.star.table.CellVertJustify import CENTER as VERT_CENTER
from com.sun.star.table.CellVertJustify import TOP as VERT_TOP

try:
    unicode
except NameError:
    unicode = str

_LM_PP_ACTIVE_CTX = None
_LM_PP_LOG_STAGE = None
_LM_PP_FN_T0 = None
_LM_PP_COPY_ROW_CHUNK = 1000
# Слияние листов с оформлением: Copy/Paste столбца пакетами (не по ячейкам).
_LM_PP_MERGE_FORMAT_CHUNK = 200
_LM_PP_DEFAULT_FONT_SIZE = 11.0
_LM_PP_FREEZE_HEADER = True
# False — не вызывать _lm_finalize_header_row_appearance (см. MERGE_FINALIZE_HEADER).
_LM_MERGE_FINALIZE_HEADER = False
_LM_PP_HEADER_HEIGHT_DEBUG = False
_LM_PP_APPLY_FORMULA_DEBUG = False
_LM_PP_SORT_DEBUG = False
# Режим сортировки: "native" (range.sort + createSortDescriptor, формулы сохраняются) | "data_array"
_LM_PP_SORT_MODE = "native"
# True — один dispatch .uno:SetOptimalColumnWidth по блоку данных (не по умолчанию: нестабильно в макросах).
_LM_PP_AUTOFIT_USE_DISPATCH = False


_LM_CONSOLE_LOG_HOOK = None


def lm_set_console_log_hook(callback):
    """Зеркало консольного debug-лога (collect_workbooks → лист «Сбор_книг_лог»)."""
    global _LM_CONSOLE_LOG_HOOK
    _LM_CONSOLE_LOG_HOOK = callback


_LM_FREEZE_DEFER_HOOK = None


def lm_set_freeze_defer_hook(callback):
    """Отложить freeze panes, пока окно книги скрыто (xml_excel_ods)."""
    global _LM_FREEZE_DEFER_HOOK
    _LM_FREEZE_DEFER_HOOK = callback


def _lm_console_log(line):
    hook = globals().get("_LM_CONSOLE_LOG_HOOK")
    if hook is not None:
        try:
            hook(line)
        except Exception:
            pass


def _lm_dbg(msg):
    if not LIBRE_MACROS_DEBUG:
        return
    try:
        line = "[libre_macros DEBUG] %s" % msg
        print(line)
        _lm_console_log(line)
    except Exception:
        pass


def _lm_pp_header_height_dbg(where, note=""):
    """Отладка «заголовок_плюс_высота» (_LM_PP_HEADER_HEIGHT_DEBUG или LIBRE_MACROS_DEBUG)."""
    if not (_LM_PP_HEADER_HEIGHT_DEBUG or LIBRE_MACROS_DEBUG):
        return
    try:
        parts = ["[заголовок_плюс_высота]", str(where)]
        if note not in ("", None):
            parts.append(str(note))
        print(" ".join(parts))
        _lm_console_log(" ".join(parts))
    except Exception:
        pass


def _lm_pp_apply_formula_dbg(doc, sheet_name, note):
    """Отладка «применить_формулу» → журнал и stdout."""
    _lm_pp_formula_fill_dbg(doc, sheet_name, note, channel="применить_формулу")


def _lm_pp_skip_rows_dbg(doc, sheet_name, note):
    """Отладка «удаление_строк» / «пропуск_пустых_строк»."""
    _lm_pp_formula_fill_dbg(doc, sheet_name, note, channel="удаление_строк")


def _lm_pp_formula_fill_dbg(doc, sheet_name, note, channel="formula_fill"):
    """Отладка протяжки формул (применить_формулу, удаление_строк и т.п.)."""
    if not (LIBRE_MACROS_DEBUG or _LM_PP_APPLY_FORMULA_DEBUG):
        return
    msg = str(note or "")
    ch = str(channel or "formula_fill")
    try:
        line = "[%s] %s" % (ch, msg)
        print(line)
        _lm_console_log(line)
    except Exception:
        pass
    if not LIBRE_MACROS_DEBUG:
        _lm_log_postprocess(doc, sheet_name or "", "диапазон", ch, "дебаг", msg)


def _lm_pp_sort_dbg(where, note=""):
    """Отладка «сортировка» в stdout (и опционально журнал)."""
    if not (_LM_PP_SORT_DEBUG or LIBRE_MACROS_DEBUG):
        return
    try:
        parts = ["[сортировка]", str(where)]
        if note not in ("", None):
            parts.append(str(note))
        print(" ".join(parts))
        _lm_console_log(" ".join(parts))
    except Exception:
        pass


def _lm_pp_merge_sheets_dbg(where, note=""):
    """Отладка «объединить_листы_в_один» → stdout при MERGE_DEBUG."""
    if not LIBRE_MACROS_DEBUG:
        return
    try:
        parts = ["[объединить_листы_в_один]", str(where)]
        if note not in ("", None):
            parts.append(str(note))
        print(" ".join(parts))
        _lm_console_log(" ".join(parts))
    except Exception:
        pass


def _lm_pp_parse_as_values_flag(block):
    """Разбор as_values из блока кодека (true/false, 1/0, да/нет)."""
    if not isinstance(block, dict):
        return False
    if "as_values" not in block:
        return False
    raw = block.get("as_values")
    try:
        parsed = lm_parse_bool_param(raw, default=None)
        if parsed is not None:
            return bool(parsed)
    except ValueError:
        pass
    try:
        return bool(raw)
    except Exception:
        return False


def lm_pp_active_context():
    """Текущий контекст постобработки (_LM_PP_ACTIVE_CTX) или None."""
    return _LM_PP_ACTIVE_CTX


def lm_pp_set_active_context(ctx):
    """Установить контекст (вызывается макросом-хостом перед шагами постобработки)."""
    global _LM_PP_ACTIVE_CTX
    _LM_PP_ACTIVE_CTX = ctx


def lm_pp_set_log_stage(stage):
    """«финальная_обработка» — дочерние lm_pp_range_* пишут в журнал финала."""
    global _LM_PP_LOG_STAGE
    _LM_PP_LOG_STAGE = stage


def lm_pp_active_log_stage():
    return _LM_PP_LOG_STAGE


def lm_pp_log_fn_begin():
    """Начало замера времени для следующей строки журнала."""
    global _LM_PP_FN_T0
    _LM_PP_FN_T0 = time.time()


def lm_pp_log_fn_end():
    global _LM_PP_FN_T0
    _LM_PP_FN_T0 = None


def lm_pp_format_log_note(note=""):
    """Дописать к примечанию время выполнения (сек), если задан lm_pp_log_fn_begin."""
    note = str(note or "")
    if _LM_PP_FN_T0 is None:
        return note
    elapsed = time.time() - _LM_PP_FN_T0
    suffix = "%.2f с" % elapsed
    if note != "":
        return "%s, %s" % (note, suffix)
    return suffix


def _lm_pp_resolve_log_scope(default_scope):
    stage = str(_LM_PP_LOG_STAGE or "").strip().casefold()
    if stage == u"финальная_обработка":
        return u"финальная_обработка"
    return default_scope


def _lm_log_postprocess(doc, target_sheet, scope, fn_label, status, note=""):
    """Журнал постобработки (хост может задать lm_pp_set_log_hook)."""
    resolved_scope = _lm_pp_resolve_log_scope(scope)
    note = str(note or "")
    if resolved_scope == u"финальная_обработка":
        sheet_note = str(target_sheet or "").strip()
        if sheet_note != "":
            note = ("%s: %s" % (sheet_note, note)) if note != "" else sheet_note
        target_sheet = ""
    if _lm_pp_log_hook is not None:
        try:
            _lm_pp_log_hook(doc, target_sheet, resolved_scope, fn_label, status, note)
        except Exception:
            pass


_lm_pp_log_hook = None


def lm_pp_set_log_hook(fn):
    """fn(doc, target_sheet, scope, fn_label, status, note) — опциональный журнал."""
    global _lm_pp_log_hook
    _lm_pp_log_hook = fn


def _lm_pp_foreach_row(sheet, sr, er, row_fn, status_prefix=None):
    """Цикл по строкам без UI (упрощённая версия для автономной lib)."""
    r = sr
    while r <= er:
        row_fn(sheet, r)
        r = r + 1


def _lm_ui_should_update_status(counter, is_last=False):
    return False


def _lm_ui_is_dialog_mode():
    return False


def _lm_ui_status_set(counter, text):
    pass


def _lm_ui_yield(counter=0, force=False):
    pass


def _lm_ui_poll_abort():
    return False


def _lm_ui_detail_set(text):
    pass


def _lm_ui_sync_progress(extra_frac=0.0, file_frac=None):
    pass


def _lm_pp_status_text(prefix, detail=""):
    if detail:
        return "%s: %s" % (prefix, detail)
    return str(prefix)


def _lm_postprocess_callbacks(callback):
    if callback is None:
        return []
    if callable(callback):
        return [callback]
    if isinstance(callback, (list, tuple)):
        result = []
        i = 0
        while i < len(callback):
            item = callback[i]
            if item is not None and callable(item):
                result.append(item)
            i = i + 1
        return result
    return []


def header_at(sheet, row, col):
    t = cell_text(sheet.getCellByPosition(col, row))
    try:
        from libre_macros_header_lib import clean_header_name

        t = clean_header_name(t)
    except Exception:
        pass
    if t != "":
        return t
    return "Колонка_%s" % col_index_to_letters(col)



# --- Зависимости (дубликаты из collect_workbooks) ---

def cell_text(cell):
    """Текст ячейки Calc: getString/String (trim) или Value; пустая → ""."""
    try:
        if hasattr(cell, "getString"):
            s = cell.getString()
            if s is not None:
                s = str(s).strip()
                if s:
                    return s
    except Exception:
        pass
    try:
        s = cell.String
        if s:
            s_stripped = str(s).strip()
            if s_stripped:
                return s_stripped
    except Exception:
        pass
    try:
        v = cell.Value
        if v not in (None, 0.0, 0):
            if v == int(v):
                return str(int(v))
            return str(v)
    except Exception:
        pass
    return ""

def col_letters_to_index(letters):
    """
    Преобразование букв столбца Excel/Calc в 0-based индекс.

    Параметры:
        letters — «A», «AA», «CZ» (регистр не важен).

    Возвращает:
        int — 0 для A, 1 для B, …

    Связь:
        parse_col_spec, build_column_list (col_kind LETTER).
    """
    letters = letters.strip().upper()
    n = 0
    i = 0
    while i < len(letters):
        n = n * 26 + (ord(letters[i]) - ord("A") + 1)
        i = i + 1
    return n - 1


def col_index_to_letters(index):
    """
    0-based индекс столбца → буквы Calc (A, B, …, AA).

    Связь:
        _merge_skip_columns_to_calc_formula, _lm_skip_expand_formula_refs.
    """
    n = int(index) + 1
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters

def is_col_letters(spec):
    """
    Проверка: spec — буквы столбца Calc (латиница A…XYZ, 1–3 символа A–Z).

    Параметры:
        spec — элемент фильтра «Столбцы» или «Начало_Столбец».

    Возвращает:
        bool — False для кириллических имён (ФИО, Отдел) и для строк
        длиннее 3 символов (GUID, NAME… — иначе читаются как огромный индекс).

    Связь:
        classify_column_specs, parse_block_from_strings, build_column_list,
        _lm_pp_resolve_column_token (латиница A…XYZ не ищется подстрокой).
    """
    s = spec.strip().upper()
    if s == "" or len(s) > 3:
        return False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch < "A" or ch > "Z":
            return False
        i = i + 1
    return True


def lm_identity_key(text):
    """Ключ сравнения имён/заголовков: trim + casefold (как merge_identity_key в макросе)."""
    if text is None:
        return ""
    return str(text).strip().casefold()


def text_match(text, pattern):
    """
    Сопоставление имени листа или заголовка столбца с шаблоном.

    Правила:
        «Все» / «*» — любой text;
        иначе при ``*``/``?`` — fnmatch (``Отчет*``, ``*Отчет``, ``*Отчет*``);
        без маски — точное совпадение (без учёта регистра).

    Параметры:
        text — имя листа или заголовок (после strip).
        pattern — элемент фильтра «Листы»/«Столбцы».

    Связь:
        iter_sheet_slots (NAME), build_column_list (NAME).
    """
    import fnmatch

    pattern = str(pattern or "").strip()
    text = str(text or "").strip()
    if lm_identity_key(pattern) == "все" or pattern == "*":
        return True
    if "*" in pattern or "?" in pattern:
        return fnmatch.fnmatchcase(text.casefold(), pattern.casefold())
    return lm_identity_key(text) == lm_identity_key(pattern)


def _lm_pp_header_matches_marker(title, marker):
    """
    Заголовок столбца совпадает с маркером (единая логика column_header_matches_token).

    'name' — exact; 'name*' — шаблон; без кавычек — '*part*' (подстрока) или fnmatch.
    """
    try:
        from libre_macros_param_codec import column_header_matches_token

        return column_header_matches_token(title, marker)
    except Exception:
        pass
    pattern_raw = str(marker or "").strip()
    text = str(title or "").strip()
    if pattern_raw == "" or text == "":
        return False
    return pattern_raw.casefold() in text.casefold()


def _lm_pp_resolve_delete_marker_columns(marker, sheet, header_row_range, end_col):
    """
    Один токен «удалить_столбцы» → список 0-based индексов столбцов.

    Число (1-based), буква A/AA — один столбец (без отсечения по used-area);
    иначе все заголовки по `_lm_pp_match_header_marker`
    ('…' = полное совпадение; без кавычек — подстрока / шаблон *).
    """
    token = str(marker or "").strip()
    if token == "":
        return []
    max_col = int(end_col) if end_col is not None else -1
    try:
        if sheet is not None:
            cnt = int(sheet.getColumns().getCount())
            if cnt > 0:
                max_col = max(max_col, cnt - 1)
    except Exception:
        pass
    if max_col < 0:
        max_col = 1023
    try:
        idx = int(token) - 1
        if 0 <= idx <= max_col:
            return [idx]
        return []
    except (TypeError, ValueError):
        pass
    if is_col_letters(token):
        idx = col_letters_to_index(token)
        if 0 <= idx <= max_col:
            return [idx]
        return []
    if header_row_range is None or sheet is None:
        return []
    titles = _lm_pp_header_titles(sheet, header_row_range)
    h_sc, _h_sr, h_ec, _h_er = _lm_pp_range_address(header_row_range)
    out = []
    c = h_sc
    while c <= h_ec:
        pos = c - h_sc
        title = titles[pos] if pos < len(titles) else ""
        if _lm_pp_match_header_marker(token, title):
            out.append(c)
        c = c + 1
    return out


def _lm_doc_char_locale(doc):
    """Локаль книги (CharLocale) или None."""
    if doc is None:
        return None
    try:
        cl = doc.CharLocale
        if cl is not None:
            return cl
    except Exception:
        pass
    return None

_SKIP_FORMULA_CELL_REF = re.compile(r"([A-Za-z]+)(\d+)")
_SKIP_FORMULA_BRACKET_REF = re.compile(r"\[([^\]]+)\]\{([^}]+)\}")
_SKIP_FORMULA_NAME_ROWID_REF = re.compile(
    r"(?<![A-Za-zА-Яа-яЁё_])([A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9_]*)\{([^}]+)\}"
)
_ROW_ID_BRACE_RE = re.compile(r"\$\{([^}]+)\}|\{([^}]+)\}")
_R1C1_REF_RE = re.compile(
    r"R(\[([+-]?\d+)\]|(\d+))?C(\[([+-]?\d+)\]|(\d+))?",
    re.IGNORECASE,
)
_R1C1_CYR_COL_RE = re.compile(
    r"R(\[([+-]?\d+)\]|(\d+))?([Сс])(\[([+-]?\d+)\]|(\d+))?"
)
_SKIP_HEADER_SCAN_LIMIT = 300


def _lm_split_filter_specs(text):
    """Разбиение списка столбцов/маркеров с учётом '…' и ;/,."""
    s = str(text or "").strip()
    if s == "":
        return []
    try:
        from libre_macros_param_codec import split_column_name_tokens

        parts = split_column_name_tokens(s)
        if parts:
            return parts
    except Exception:
        pass
    parts = []
    normalized = s.replace(";", ",")
    for chunk in normalized.split(","):
        st = chunk.strip()
        if st != "":
            parts.append(st)
    return parts


def parse_skip_row_spec(spec_text):
    """
    Разбор условия пропуска/удаления строк (columns | formula | none).

    Используется «удаление_строк» / «пропуск_пустых_строк» и фильтрами источника.
    """
    text = str(spec_text).strip()
    if text == "":
        return {"mode": "none"}
    if re.search(r"R(?:\[[+-]?\d+\]|\d+)?[CcСс]", text):
        return {"mode": "formula", "formula": text}
    if re.search(r"\[[^\]]+\]", text):
        return {"mode": "formula", "formula": text}
    if re.search(r"\{[^}]*\brow_id\b[^}]*\}", text) or re.search(
        r"\$\{[^}]*\brow_id\b[^}]*\}", text
    ):
        return {"mode": "formula", "formula": text}
    if re.search(
        r"[A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_]*\{[^}]*\brow_id\b[^}]*\}",
        text,
    ):
        return {"mode": "formula", "formula": text}
    if re.search(r"[A-Za-z]+\{", text):
        return {"mode": "formula", "formula": text}
    if re.search(r"[A-Za-z]+\d+", text) and re.search(r"[><=!]", text):
        return {"mode": "formula", "formula": text}
    if re.search(r"[A-Za-z]+\d+\s*[\+\-\*/]", text):
        return {"mode": "formula", "formula": text}
    parts = _lm_split_filter_specs(text)
    if len(parts) == 0:
        return {"mode": "none"}
    return {"mode": "columns", "parts": parts}

def _resolve_skip_column_spec_to_index(sheet, hdr_row, start_col, col_count, spec):
    """
    Столбец для проверки пропуска: номер, буква или имя/шаблон заголовка.

    Токен: 'name' exact; 'name*' шаблон; без кавычек — '*part*' / fnmatch / цифра / буква.
    """
    st_raw = str(spec).strip()
    if st_raw == "":
        return -1
    quoted = False
    try:
        from libre_macros_param_codec import unwrap_column_name_token

        st, quoted = unwrap_column_name_token(st_raw)
        st = str(st).strip()
    except Exception:
        st = st_raw
        quoted = False
    if st == "":
        return -1
    if (not quoted) and st.isdigit():
        return int(st) - 1
    if (not quoted) and is_col_letters(st):
        return col_letters_to_index(st)
    try:
        from libre_macros_param_codec import column_header_matches_token

        match_fn = column_header_matches_token
    except Exception:
        match_fn = None
    scan = int(start_col)
    limit = scan + _SKIP_HEADER_SCAN_LIMIT
    while scan < limit:
        hdr = header_at(sheet, hdr_row, scan)
        if match_fn is not None:
            if match_fn(hdr, st_raw):
                return scan
        elif lm_identity_key(hdr) == lm_identity_key(st):
            return scan
        scan = scan + 1
    return -1


def _resolve_skip_column_indices(sheet, hdr_row, start_col, col_count, parts):
    indices = []
    missing = []
    for part in parts or []:
        idx = _resolve_skip_column_spec_to_index(
            sheet, hdr_row, start_col, col_count, part
        )
        if idx >= 0:
            if idx not in indices:
                indices.append(idx)
        else:
            missing.append(str(part).strip())
    return indices, missing


def _lm_safe_eval_bool(expr):
    """Безопасное вычисление логического выражения после подстановки чисел."""
    s = str(expr).strip()
    if s == "":
        return False
    if not re.match(r"^[\d\.\+\-\*/\%\(\)<>=!\s]+$", s):
        return False
    try:
        return bool(eval(s))
    except Exception:
        return False


def _lm_safe_eval_row_index_expr(expr, row_id_value, row_id_anchor=None):
    """
    Вычислить номер строки Calc (1-based) из выражения в {…} с row_id.

    «row_id» без смещения — row_id_value (текущая строка при построчной проверке).
    В выражениях со смещением (row_id-1, row_id+2) подставляется row_id_anchor
    (первая строка данных) или row_id_value, если якорь не задан.
    """
    s = str(expr).strip()
    if "row_id" not in s:
        return None
    if re.match(r"^\s*row_id\s*$", s):
        return int(row_id_value)
    base = int(row_id_anchor) if row_id_anchor is not None else int(row_id_value)
    s = re.sub(r"\brow_id\b", str(base), s)
    if not re.match(r"^[\d\+\-\*/\%\(\)\s]+$", s):
        return None
    try:
        v = eval(s)
        if isinstance(v, (int, float)):
            return int(v)
    except Exception:
        pass
    return None


def _lm_safe_eval_col_index_expr(expr, col_id_value, col_id_anchor=None):
    """
    Вычислить 0-based индекс столбца из выражения в {…} с col_id.

    «col_id» без смещения — col_id_value (столбец вставки формулы).
    В выражениях со смещением (col_id-1) подставляется col_id_anchor или col_id_value.
    """
    s = str(expr).strip()
    if "col_id" not in s:
        return None
    if re.match(r"^\s*col_id\s*$", s):
        return int(col_id_value)
    base = int(col_id_anchor) if col_id_anchor is not None else int(col_id_value)
    s = re.sub(r"\bcol_id\b", str(base), s)
    if not re.match(r"^[\d\+\-\*/\%\(\)\s]+$", s):
        return None
    try:
        v = eval(s)
        if isinstance(v, (int, float)):
            return int(v)
    except Exception:
        pass
    return None


def _lm_skip_expand_formula_refs( sheet, hdr_row, start_col, col_count, formula_text, doc=None ):
    """
    Подставить буквы столбцов вместо ссылок на заголовок в шаблоне формулы:
    [Имя_Заголовка]{row_id}, Имя_Заголовка{row_id}, голый [Имя] / {Имя}.
    """
    text = str(formula_text)
    uses_r1c1 = _lm_pp_doc_uses_r1c1(doc) if doc is not None else False

    def _header_to_col_letters(name):
        return _lm_pp_header_to_col_token(
            sheet, hdr_row, start_col, col_count, name, uses_r1c1=uses_r1c1
        )

    def _bracket_replace(match):
        letters = _header_to_col_letters(match.group(1).strip())
        if letters is None:
            return match.group(0)
        return letters + "{" + match.group(2) + "}"

    def _name_rowid_replace(match):
        name = match.group(1)
        brace = match.group(2)
        if len(name) == 1 and name.isalpha() and "A" <= name.upper() <= "Z":
            return match.group(0)
        letters = _header_to_col_letters(name)
        if letters is None:
            return match.group(0)
        return letters + "{" + brace + "}"

    def _bare_bracket_replace(match):
        name = match.group(1) if match.group(1) is not None else match.group(2)
        name = str(name or "").strip()
        if name == "":
            return match.group(0)
        letters = _header_to_col_letters(name)
        if letters is None:
            return match.group(0)
        return letters

    text = _SKIP_FORMULA_BRACKET_REF.sub(_bracket_replace, text)
    text = _SKIP_FORMULA_NAME_ROWID_REF.sub(_name_rowid_replace, text)
    text = _SKIP_FORMULA_BARE_BRACKET_REF.sub(_bare_bracket_replace, text)
    return text

def _lm_row_skip_by_empty_columns(sheet, row_0, col_indices):
    """True — все указанные столбцы пусты → пропустить строку."""
    if col_indices is None or len(col_indices) == 0:
        return False
    ci = 0
    while ci < len(col_indices):
        if cell_text(sheet.getCellByPosition(col_indices[ci], row_0)) != "":
            return False
        ci = ci + 1
    return True

def get_sheet_used_bounds(sheet):
    """Последний использованный столбец и строка (0-based) на листе."""
    if sheet is None:
        return 0, 0
    try:
        cursor = sheet.createCursor()
        if cursor is None:
            return 0, 0
        cursor.gotoEndOfUsedArea(False)
        addr = cursor.getRangeAddress()
        return addr.EndColumn, addr.EndRow
    except Exception:
        return 0, 0

def _lm_postprocess_callbacks_from_collect(callback):
    """
    Нормализация спецификации колбэков постобработки в список callable.

    Параметры:
        callback — None, одна функция, list или tuple (None и не-callable пропускаются).

    Возвращает:
        list — пустой список означает «постобработка отключена».

    Связь:
        _merge_postprocess_has_callbacks, _merge_run_postprocess_range/row.
    """
    if callback is None:
        return []
    if callable(callback):
        return [callback]
    if isinstance(callback, (list, tuple)):
        result = []
        i = 0
        while i < len(callback):
            item = callback[i]
            if item is not None and callable(item):
                result.append(item)
            i = i + 1
        return result
    return []

def _lm_pp_context_expand_end_col(ctx, col_index):
    """Расширить end_col в контексте постобработки (новый столбец справа)."""
    if ctx is None:
        return ctx
    try:
        col_index = int(col_index)
    except (TypeError, ValueError):
        return ctx
    if col_index > int(ctx.get("end_col", 0)):
        ctx["end_col"] = col_index
    return ctx


def _lm_pp_context_refresh(ctx):
    """Перечитать end_col/end_row с листа (после изменения структуры столбцов)."""
    if ctx is None:
        return ctx
    sheet = ctx.get("sheet")
    header_row = ctx.get("header_row", 0)
    if sheet is None:
        return ctx
    end_col, end_row = get_sheet_used_bounds(sheet)
    ctx["end_col"] = end_col
    ctx["end_row"] = end_row
    ctx["data_start"] = int(header_row) + 1
    return ctx

_LM_PP_HORI_JUSTIFY_ALIASES = {
    "left": "left",
    "лево": "left",
    "левый": "left",
    "л": "left",
    "center": "center",
    "centre": "center",
    "центр": "center",
    "по центру": "center",
    "ц": "center",
    "right": "right",
    "право": "right",
    "правый": "right",
    "п": "right",
}

_LM_PP_VERT_JUSTIFY_ALIASES = {
    "top": "top",
    "верх": "top",
    "верхний": "top",
    "в": "top",
    "center": "center",
    "centre": "center",
    "центр": "center",
    "по центру": "center",
    "ц": "center",
    "bottom": "bottom",
    "низ": "bottom",
    "нижний": "bottom",
    "н": "bottom",
}

_LM_PP_SHEET_ALL_MARKERS = frozenset({"все", "all", "*"})


def _lm_pp_resolve_hori_justify(name):
    key = _LM_PP_HORI_JUSTIFY_ALIASES.get(str(name or "").strip().casefold(), "center")
    if key == "left":
        return HORI_LEFT
    if key == "right":
        return HORI_RIGHT
    return HORI_CENTER


def _lm_pp_resolve_vert_justify(name):
    key = _LM_PP_VERT_JUSTIFY_ALIASES.get(str(name or "").strip().casefold(), "center")
    if key == "top":
        return VERT_TOP
    if key == "bottom":
        return VERT_BOTTOM
    return VERT_CENTER


def _lm_pp_set_cell_alignment(cell, hori_val, vert_key):
    """Горизонталь и вертикаль: VertJustify + VertJustify2 (LO 7+)."""
    if cell is None:
        return
    try:
        cell.HoriJustify = hori_val
    except Exception:
        pass
    vert_key = str(vert_key or "center").strip().casefold()
    try:
        from com.sun.star.table.CellVertJustify import (
            BOTTOM as VJ_BOTTOM,
            CENTER as VJ_CENTER,
            TOP as VJ_TOP,
        )

        if vert_key == "top":
            cell.VertJustify = VJ_TOP
        elif vert_key == "bottom":
            cell.VertJustify = VJ_BOTTOM
        else:
            cell.VertJustify = VJ_CENTER
    except Exception:
        pass
    try:
        from com.sun.star.table.CellVertJustify2 import (
            BOTTOM as VJ2_BOTTOM,
            CENTER as VJ2_CENTER,
            TOP as VJ2_TOP,
        )

        if vert_key == "top":
            cell.VertJustify2 = VJ2_TOP
        elif vert_key == "bottom":
            cell.VertJustify2 = VJ2_BOTTOM
        else:
            cell.VertJustify2 = VJ2_CENTER
    except Exception:
        pass


def _lm_apply_header_row_alignment(sheet, end_col, header_row=0, hori="center", vert="center"):
    """Заголовок: выравнивание ячеек (горизонталь и вертикаль)."""
    if sheet is None or end_col < 0:
        return
    hori_val = _lm_pp_resolve_hori_justify(hori)
    vert_key = _LM_PP_VERT_JUSTIFY_ALIASES.get(
        str(vert or "").strip().casefold(), "center"
    )
    c = 0
    while c <= end_col:
        try:
            cell = sheet.getCellByPosition(c, header_row)
            _lm_pp_set_cell_alignment(cell, hori_val, vert_key)
        except Exception:
            pass
        c = c + 1

def _lm_apply_header_row_font_size(sheet, end_col, header_row=0):
    """Размер шрифта заголовка (_LM_PP_DEFAULT_FONT_SIZE); по ячейкам — надёжнее, чем XCellRange."""
    if sheet is None or end_col < 0:
        return
    size = float(_LM_PP_DEFAULT_FONT_SIZE)
    c = 0
    while c <= end_col:
        try:
            sheet.getCellByPosition(c, header_row).CharHeight = size
        except Exception:
            pass
        c = c + 1

def _lm_pp_controller_prepare_freeze(controller, sheet):
    """
    Перед freezeAtPosition: активный лист и вид с левого верхнего угла.

    В Calc freezeAtPosition относителен к FirstVisibleRow/Column: после Copy/Paste
    или буфера вид уезжает вниз — без сброса «заморозка 1 строки» цепляется
    к последнему блоку данных, а не к заголовку.
    """
    if controller is None or sheet is None:
        return False
    try:
        active = controller.getActiveSheet()
        if active is None or str(active.Name) != str(sheet.Name):
            controller.setActiveSheet(sheet)
    except Exception:
        try:
            controller.setActiveSheet(sheet)
        except Exception:
            return False
    try:
        controller.FirstVisibleColumn = 0
    except Exception:
        pass
    try:
        controller.FirstVisibleRow = 0
    except Exception:
        pass
    return True


def _lm_freeze_result_header(doc, sheet, header_row=0):
    """
    Закрепить строки заголовка на листе результата (freeze panes).

    header_row — 0-based; закрепляются строки 0…header_row (данные с header_row + 1).
    Отключается параметром «Закрепить_заголовок» (_LM_PP_FREEZE_HEADER).
    """
    if not _LM_PP_FREEZE_HEADER:
        return
    if doc is None or sheet is None:
        return
    try:
        freeze_rows = int(header_row) + 1
        if freeze_rows < 1:
            freeze_rows = 1
        controller = doc.getCurrentController()
        if controller is None:
            return
        if not _lm_pp_controller_prepare_freeze(controller, sheet):
            return
        controller.freezeAtPosition(0, freeze_rows)
    except Exception as err:
        print(
            "Закрепление заголовка листа «%s»: %s"
            % (sheet.Name if sheet is not None else "", err)
        )


# --- Постобработка (из region collect_workbooks) ---

# =============================================================================
# Константы постобработки (_LM_PP_* и связанные _MERGE_FMT_*)
# =============================================================================

# --- Общие для форматирования (используются и в apply_result_sheet_formatting) ---
_LM_FMT_CHAR_WEIGHT_BOLD = 150
# Серый фон заголовка (как _MERGE_FMT_HEADER_BG в collect_workbooks).
_LM_FMT_HEADER_BG = 0xD9D9D9
# Вес шрифта LO: 150 = жирный (lm_pp_row_bold_if_path_contains).

_LM_FMT_SERVICE_FG = 0x9E9E9E
# Цвет текста служебных колонок #Путь / #ДатаВремяФайла (RGB 158,158,158).

_LM_FMT_HEADER_ROW_PAD_MM = 12.7
# Высота заголовка по умолчанию (мм) и минимум (защита от схлопывания).
_LM_FMT_HEADER_ROW_MIN_MM = 7.0
# Дополнительная высота строки заголовков в мм (lm_pp_range_header_row_height_pad; ≈ бывш. 0,5″).

_LM_PP_ALIGN_HINT = (
    "left/center/right (Л/Ц/П, лево/центр/право) и "
    "top/center/bottom (В/Ц/Н, верх/центр/низ); между горизонталью и вертикалью — -, _ или /"
)

# --- Числовые и денежные форматы ---
_LM_PP_FORMAT_MONEY_FMT = "# ##0,00"
# Шаблон NumberFormat для lm_pp_range_format_money_columns (локаль ru/RU).
# Не использовать Excel-стиль #,##0.00 — в ru/RU запятая = десятичный разделитель.

_LM_PP_FORMAT_TOTAL_SUM = "# ##0,00"
# Формат ячеек SUM в строке «Итого» (pp_stroka_itogo; локаль ru/RU).

_LM_PP_DATE_FORMAT = "DD.MM.YYYY"
# Шаблон даты для lm_pp_range_format_date_columns (локаль ru/RU).

_LM_PP_NF_TYPE_DATE = 2
_LM_PP_NF_TYPE_TIME = 4
_LM_PP_NF_TYPE_DATETIME = 6
# Типы com.sun.star.util.NumberFormat для ВПР / дат.

_LM_PP_VLOOKUP_DATE_TITLE_MARKERS = (
    "дата",
    "date",
    "время",
    "time",
    "рожден",
    "трудоустрой",
)

_LM_PP_NUMBER_DECIMALS = 2
# Зарезервировано для будущих колбэков с фиксированной точностью (сейчас не читается).

# --- Построчная геометрия ---
_LM_PP_ROW_HEIGHT_PAD_INCH = 0.08
# Запас высоты одной строки данных (~2 мм), lm_pp_row_height_pad (row-уровень).

_LM_PP_DEFAULT_ROW_HEIGHT_MM = 5.0
# Высота строки по умолчанию (lm_pp_range_row_height_pad), если C пуста.

# --- Состояние «полосы по пути» (между вызовами row-колбэков на одном листе) ---
_LM_PP_PATH_STRIPE_STATE = {}
# Ключи: «имя_листа:path», «имя_листа:stripe». Сброс: lm_pp_range_init_path_stripe_state.

# --- Подсветка и жирность по содержимому ---
_LM_PP_ROW_HIGHLIGHT_HEADER = "итог"
# Подстрока в заголовке столбца для lm_pp_row_highlight_by_header_value.

_LM_PP_ROW_HIGHLIGHT_BG = None
# Фон строки при срабатывании подсветки; None → светло-жёлтый lo_color_rgb(255,242,204).

_LM_PP_ROW_HIGHLIGHT_FG = None
# Цвет текста строки; None — не менять CharColor.

_LM_PP_ROW_BOLD_PATH_SUBSTRING = ""
# Подстрока в #Путь для жирного шрифта; пусто — lm_pp_row_bold_if_path_contains no-op.

_LM_PP_FONT_NAME = "PT Sans"
# Имя шрифта для lm_pp_range_set_font (данные + заголовок).


# =============================================================================
# Утилиты секции постобработки
# =============================================================================

def lo_color_rgb(red, green, blue):
    """
    Собрать цвет Calc/LO в формате 0xRRGGBB для CellBackColor / CharColor2.

    Не колбэк постобработки; используется внутри lm_pp_* и форматирования листа.
    """
    return (int(red) << 16) | (int(green) << 8) | int(blue)


def lo_color_hex(hex_rgb):
    """
    Разбор цвета из HEX (#RRGGBB) или имени из _LM_PP_ZEBRA_COLOR_ALIASES.

    Используется в _lm_pp_zebra_resolve_color_token и подсветке ВПР.
    """
    key = str(hex_rgb).strip().lower()
    if key in _LM_PP_ZEBRA_COLOR_ALIASES:
        t = _LM_PP_ZEBRA_COLOR_ALIASES[key]
        return lo_color_rgb(t[0], t[1], t[2])
    h = key.lstrip("#")
    if len(h) != 6:
        raise ValueError("цвет: ожидается #RRGGBB или имя, получено: %s" % hex_rgb)
    return lo_color_rgb(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


_LM_BOOL_TRUE_TOKENS = frozenset({
    "1", "+", "true", "yes", "y", "да", "д", "вкл", "включено", "on", "истина",
})
_LM_BOOL_FALSE_TOKENS = frozenset({
    "0", "-", "false", "no", "n", "нет", "н", "выкл", "выключено", "off", "ложь",
})


def lm_parse_bool_param(raw, default=True):
    """
    Разбор параметра Да/Нет без учёта регистра.

    Истина: Да, Yes, true, 1, +, истина, вкл, on, …
    Ложь: Нет, No, false, 0, -, ложь, выкл, off, …
    Пусто или None — default.
    """
    if raw is None:
        return default
    s = str(raw).strip().lower()
    if s == "":
        return default
    if s in _LM_BOOL_TRUE_TOKENS:
        return True
    if s in _LM_BOOL_FALSE_TOKENS:
        return False
    raise ValueError("not bool: %r" % raw)


def _lm_sheet_from_range(cell_range):
    """Лист Calc по XCellRange."""
    if cell_range is None:
        return None
    try:
        return cell_range.getSpreadsheet()
    except Exception:
        return None


def _lm_foreach_cell_in_range(sheet, cell_range, cell_fn):
    """Вызвать cell_fn(cell) для каждой ячейки диапазона."""
    if sheet is None or cell_range is None or cell_fn is None:
        return
    sc, sr, ec, er = _lm_pp_range_address(cell_range)
    r = sr
    while r <= er:
        c = sc
        while c <= ec:
            try:
                cell_fn(sheet.getCellByPosition(c, r))
            except Exception:
                pass
            c = c + 1
        r = r + 1


def _lm_cell_has_property(cell, prop_name):
    try:
        info = cell.getPropertySetInfo()
        return info is not None and info.hasPropertyByName(prop_name)
    except Exception:
        return False


def _lm_cell_reset_property_to_default(cell, prop_name):
    if not _lm_cell_has_property(cell, prop_name):
        return
    try:
        cell.setPropertyToDefault(prop_name)
    except Exception:
        pass
    try:
        cell.resetPropertyValue(prop_name)
    except Exception:
        pass


def _lm_apply_default_header_row_appearance(sheet, end_col, header_row=0):
    """
    Оформление заголовка по умолчанию: серый фон, PT Sans, жирный, по центру.
    Как встроенное оформление листа результата без зебры.
    """
    if sheet is None or end_col < 0:
        return
    try:
        header_range = sheet.getCellRangeByPosition(0, header_row, end_col, header_row)
    except Exception:
        header_range = None
    if header_range is not None:
        try:
            header_range.IsCellBackgroundTransparent = False
            header_range.CellBackColor = int(_LM_FMT_HEADER_BG)
        except Exception:
            pass
        try:
            header_range.CharFontName = _LM_PP_FONT_NAME
            header_range.CharWeight = _LM_FMT_CHAR_WEIGHT_BOLD
        except Exception:
            pass
    else:
        c = 0
        while c <= end_col:
            try:
                cell = sheet.getCellByPosition(c, header_row)
                cell.IsCellBackgroundTransparent = False
                cell.CellBackColor = int(_LM_FMT_HEADER_BG)
                cell.CharFontName = _LM_PP_FONT_NAME
                cell.CharWeight = _LM_FMT_CHAR_WEIGHT_BOLD
            except Exception:
                pass
            c = c + 1
    _lm_apply_header_row_alignment(sheet, end_col, header_row)
    _lm_apply_header_row_font_size(sheet, end_col, header_row)


def _lm_finalize_header_row_appearance( sheet, end_col, header_row=0, doc=None, *callback_specs ):
    """
    Финальное оформление заголовка после постобработки / freeze.

    С «зебра» header:заливка/шрифт — как у строк данных (явный CharColor2).
    Без зебры — серый фон по умолчанию, PT Sans, жирный.
    Затем повторно «заголовок_плюс_высота» (если задан в параметрах).
    """
    if not _LM_MERGE_FINALIZE_HEADER:
        _lm_pp_header_height_dbg(
            "finalize_header",
            "SKIP _LM_MERGE_FINALIZE_HEADER=False sheet=%s"
            % (getattr(sheet, "Name", sheet)),
        )
        return
    if sheet is None or end_col < 0:
        return
    header_style = _lm_pp_resolve_zebra_header_style_from_specs(*callback_specs)
    try:
        header_range = sheet.getCellRangeByPosition(0, header_row, end_col, header_row)
    except Exception:
        header_range = None
    if header_style is not None and header_range is not None:
        _lm_pp_zebra_apply_header_row_style(
            sheet, header_range, header_style, header_row, end_col
        )
        _lm_apply_header_row_alignment(sheet, end_col, header_row)
        _lm_apply_header_row_font_size(sheet, end_col, header_row)
    else:
        _lm_apply_default_header_row_appearance(sheet, end_col, header_row)
    _lm_pp_maybe_reapply_header_height_pad(
        doc, sheet, header_row, end_col, *callback_specs
    )


def _lm_pp_apply_merge_dest_header_default(dest_sheet, header_row=0):
    """После объединить_листы_в_один — заголовок приёмника по умолчанию."""
    if dest_sheet is None:
        return
    try:
        end_col, _end_row = get_sheet_used_bounds(dest_sheet)
    except Exception:
        return
    if end_col < 0:
        return
    _lm_apply_default_header_row_appearance(dest_sheet, end_col, int(header_row or 0))


def _lm_set_char_color_cell(cell, color_rgb):
    """
    Цвет шрифта одной ячейки (надёжнее для сохранения в ODF, чем XCellRange).

    CharAutoColor=False; CharColor2 предпочтительнее CharColor (см. комментарий в начале файла).
    """
    if cell is None:
        return
    color_rgb = int(color_rgb)
    try:
        cell.CharAutoColor = False
    except Exception:
        pass
    try:
        cell.CharColor2 = color_rgb
    except Exception:
        pass
    try:
        cell.CharColor = color_rgb
    except Exception:
        pass


def _lm_set_char_color_safe(cell_or_range, color_rgb):
    """
    Установить цвет шрифта с учётом версии LO (CharColor2 или CharColor).

    Для диапазона — по ячейкам (иначе строка заголовков может не сохраниться).

    Параметры:
        cell_or_range — com.sun.star.table.XCell или XCellRange
        color_rgb — int, 0xRRGGBB (см. lo_color_rgb)

    Внутренний хелпер; не в карте MAP.
    """
    color_rgb = int(color_rgb)
    sheet = _lm_sheet_from_range(cell_or_range)
    if sheet is not None:
        def _apply(cell):
            _lm_set_char_color_cell(cell, color_rgb)

        _lm_foreach_cell_in_range(sheet, cell_or_range, _apply)
        return
    _lm_set_char_color_cell(cell_or_range, color_rgb)


def _lm_style_service_data_cell(cell):
    """
    Оформить одну ячейку служебной колонки бледно-серым текстом (_LM_FMT_SERVICE_FG).

    Параметры:
        cell — com.sun.star.table.XCell

    Внутренний хелпер для lm_pp_row_gray_service_columns; не в карте MAP.
    """
    _lm_set_char_color_safe(cell, _LM_FMT_SERVICE_FG)


def _lm_active_service_layout():
    """Раскладка #Путь / #ДатаВремя из cfg сбора (без импорта collect_workbooks)."""
    try:
        import libre_macros_collect_cfg as _cfg

        layout = getattr(_cfg, "_MERGE_ACTIVE_SERVICE_LAYOUT", None)
        if layout:
            return layout
    except Exception:
        pass
    return {u"path_col": 0, u"datetime_col": 1, u"first_data_col": 2}


def _lm_cell_path_value(sheet, row_index):
    layout = _lm_active_service_layout()
    pc = layout.get(u"path_col", 0)
    try:
        pc = int(pc)
    except Exception:
        pc = 0
    if pc < 0:
        return u""
    try:
        return unicode(sheet.getCellByPosition(pc, int(row_index)).String or u"")
    except Exception:
        return u""


# =============================================================================
# Внутренние хелперы постобработки
# =============================================================================

def _lm_pp_range_address(data_range):
    """
    Координаты диапазона в 0-based индексах Calc.

    Параметры:
        data_range — com.sun.star.table.XCellRange

    Возвращает (start_col, start_row, end_col, end_row). Используется всеми lm_pp_*.
    Не в карте MAP.
    """
    addr = data_range.getRangeAddress()
    return addr.StartColumn, addr.StartRow, addr.EndColumn, addr.EndRow


def _lm_pp_apply_column_number_format_chunked( sheet, col_index, sr, er, fmt_key, status_prefix="диапазон" ):
    """
    NumberFormat на столбец пакетами по _LM_PP_COPY_ROW_CHUNK строк (без фриза UI).
    """
    r = sr
    n = 0
    total = er - sr + 1
    while r <= er:
        r_end = r + _LM_PP_COPY_ROW_CHUNK - 1
        if r_end > er:
            r_end = er
        chunk = sheet.getCellRangeByPosition(col_index, r, col_index, r_end)
        chunk.NumberFormat = fmt_key
        n = n + (r_end - r + 1)
        if _lm_ui_should_update_status(n, is_last=(r_end >= er)):
            _lm_ui_status_set(
                n,
                _lm_pp_status_text(
                    status_prefix,
                    "столбец %d, строки %d/%d" % (col_index + 1, r_end + 1, er + 1),
                ),
            )
        _lm_ui_yield(counter=n)
        r = r_end + 1


def _lm_pp_parse_concat_columns_extra(extra_args):
    """
    Разбор rest для «конкатенация_столбцов»:
    имя_столбца [, разделитель] , индекс1 , индекс2 , …

    Индексы столбцов — 1-based как в Calc (1 = A, 2 = B).
    Если сразу после имени идёт целое — разделитель пустой (склейка без разделителя).
    """
    if not extra_args or len(extra_args) < 2:
        return None
    col_name = str(extra_args[0]).strip()
    if col_name == "":
        return None
    rest = extra_args[1:]
    delim = ""
    idx_start = 0
    if len(rest) >= 1:
        first = rest[0]
        if isinstance(first, int):
            idx_start = 0
        elif isinstance(first, str) and first.strip().isdigit():
            idx_start = 0
        else:
            delim = str(first)
            idx_start = 1
    indices = []
    i = idx_start
    while i < len(rest):
        item = rest[i]
        try:
            indices.append(int(item))
        except (TypeError, ValueError):
            s = str(item).strip()
            if s.isdigit():
                indices.append(int(s))
            else:
                return None
        i = i + 1
    if not indices:
        return None
    return col_name, delim, indices


def _lm_pp_find_or_append_header_column(sheet, hname, header_row):
    """
    Найти столбец по тексту заголовка в header_row или вернуть индекс нового справа.

    Не записывает заголовок — только индекс столбца (0-based).
    На пустом листе Calc used bounds часто (0,0) — новый столбец = первая
    пустая ячейка строки заголовков, а не end_col+1 (иначе фантомный столбец A).
    """
    hname = str(hname).strip()
    if hname == "" or sheet is None:
        return -1
    hkey = lm_identity_key(hname)
    try:
        header_row = int(header_row)
    except (TypeError, ValueError):
        header_row = 0
    end_col, _end_row = get_sheet_used_bounds(sheet)
    try:
        end_col = int(end_col)
    except (TypeError, ValueError):
        end_col = 0
    scan_limit = max(end_col, 0) + 64
    tc = 0
    while tc <= max(end_col, 0):
        existing = cell_text(sheet.getCellByPosition(tc, header_row)).strip()
        if existing != "" and lm_identity_key(existing) == hkey:
            return tc
        tc = tc + 1
    tc = 0
    while tc < scan_limit:
        if cell_text(sheet.getCellByPosition(tc, header_row)).strip() == "":
            return tc
        tc = tc + 1
    new_col = max(end_col, 0) + 1
    if new_col >= 300:
        return -1
    return new_col


_LM_PP_CELL_RANGE_FORMAT_PROPS = (
    # Важно: часть оформления (в т.ч. фон) может быть задана стилем ячейки.
    # Копируем стиль вместе с прямыми свойствами диапазона.
    "CellStyle",
    "NumberFormat",
    "IsCellBackgroundTransparent",
    "CellBackColor",
    "CharFontName",
    "CharHeight",
    "CharWeight",
    "CharPosture",
    "CharFontStyleName",
    "CharUnderline",
    "CharAutoColor",
    "CharColor",
    "CharColor2",
    "HoriJustify",
    "VertJustify",
    "VertJustify2",
    "ParaIndent",
    "IsTextWrapped",
    "TopBorder",
    "BottomBorder",
    "LeftBorder",
    "RightBorder",
)


def _lm_pp_copy_cell_range_format(src_range, dst_range):
    """Скопировать визуальные свойства с одного XCellRange на другой (поддиапазон)."""
    if src_range is None or dst_range is None:
        return
    pi = 0
    while pi < len(_LM_PP_CELL_RANGE_FORMAT_PROPS):
        prop = _LM_PP_CELL_RANGE_FORMAT_PROPS[pi]
        try:
            setattr(dst_range, prop, getattr(src_range, prop))
        except Exception:
            pass
        pi = pi + 1


def _lm_pp_copy_column_format_subrange(sheet, src_col, dst_col, top_row, bottom_row):
    """
    Склонировать формат столбца-образца на целевой столбец одним поддиапазоном (не по ячейкам).
    """
    if sheet is None or src_col < 0 or dst_col < 0:
        return
    try:
        top_row = int(top_row)
        bottom_row = int(bottom_row)
    except (TypeError, ValueError):
        return
    if bottom_row < top_row:
        return
    try:
        src_range = sheet.getCellRangeByPosition(src_col, top_row, src_col, bottom_row)
        dst_range = sheet.getCellRangeByPosition(dst_col, top_row, dst_col, bottom_row)
    except Exception:
        return
    _lm_pp_copy_cell_range_format(src_range, dst_range)


def _lm_pp_clone_concat_column_formats( sheet, ref_col, target_col, header_row, data_top, data_bottom ):
    """
    Склонировать формат сборного столбца поддиапазонами: отдельно заголовок и данные.

    Один диапазон на весь столбец в Calc сливает свойства строк — заголовок теряет стиль.
    """
    if sheet is None or ref_col < 0 or target_col < 0:
        return
    try:
        header_row = int(header_row)
        data_top = int(data_top)
        data_bottom = int(data_bottom)
    except (TypeError, ValueError):
        return
    _lm_pp_copy_column_format_subrange(
        sheet, ref_col, target_col, header_row, header_row
    )
    if data_bottom >= data_top:
        _lm_pp_copy_column_format_subrange(
            sheet, ref_col, target_col, data_top, data_bottom
        )


def _lm_pp_parse_formula_column_extra(extra_args):
    """
    Разбор rest для «применить_формулу».

    Новый формат (рекомендуется):  Имя_Столбца=Текст_Формулы
      пример: Итого=E{row_id}+F{row_id}

    Старый формат (совместимость): Имя_Столбца,Текст_Формулы
      пример: Итого,E{row_id}+F{row_id}

    Если в формуле есть запятые (SUM(…)), части после первой склеиваются обратно.
    """
    if not extra_args:
        return None

    # В некоторых хостах rest приходит как один элемент (строка из колонки C).
    # Поддерживаем синтаксис "имя=формула" и "имя,формула".
    if len(extra_args) == 1:
        raw = str(extra_args[0]).strip()
        if raw == "":
            return None
        if "=" in raw:
            left, right = raw.split("=", 1)
            col_name = left.strip()
            formula = right.strip()
            if col_name == "" or formula == "":
                return None
            return col_name, formula
        if "," in raw:
            left, right = raw.split(",", 1)
            col_name = left.strip()
            formula = right.strip()
            if col_name == "" or formula == "":
                return None
            return col_name, formula
        return None

    # Если формула содержит запятые (SUM(...)), хост уже мог разбить rest по запятым.
    # Поддерживаем новый формат "имя=формула" даже при extra_args вида:
    #   ["Сумма=SUM(E{row_id}", "F{row_id})"]
    head = str(extra_args[0]).strip()
    if "=" in head:
        left, right = head.split("=", 1)
        col_name = left.strip()
        if col_name == "":
            return None
        tail = ",".join(str(x) for x in extra_args[1:]).strip()
        formula = (right + ("," + tail if tail else "")).strip()
        if formula == "":
            return None
        return col_name, formula

    col_name = str(extra_args[0]).strip()
    if col_name == "":
        return None
    if len(extra_args) == 2:
        formula = str(extra_args[1]).strip()
    else:
        formula = ",".join(str(x) for x in extra_args[1:]).strip()
    if formula == "":
        return None
    return col_name, formula


def _lm_pp_block_target_sheet_names(block):
    """Имена листов из полей sheet / sheets блока JSON."""
    if not isinstance(block, dict):
        return []
    names = []
    sheet_name = str(block.get("sheet") or "").strip()
    if sheet_name != "":
        names.extend(_lm_pp_parse_sheet_names_from_spec(sheet_name))
    sheets_ref = block.get("sheets")
    if sheets_ref is not None:
        if isinstance(sheets_ref, (list, tuple)):
            for item in sheets_ref:
                s = str(item or "").strip()
                if s != "":
                    names.append(s)
        else:
            s = str(sheets_ref or "").strip()
            if s != "":
                names.append(s)
    return names


def _lm_pp_block_matches_sheet(block, doc, sheet, sheet_name=None):
    """True, если блок относится к текущему листу (sheet / sheets)."""
    names = _lm_pp_block_target_sheet_names(block)
    if not names:
        return False
    if sheet_name is None and sheet is not None:
        try:
            sheet_name = str(sheet.Name).strip()
        except Exception:
            sheet_name = ""
    if doc is not None and sheet is not None:
        return _lm_pp_sheet_ref_list_matches_sheet(doc, sheet, names)
    return _lm_pp_sheet_name_matches_target(sheet_name or "", set(names))


def _lm_pp_blocks_for_sheet(fn_key, extra_args, doc, sheet, sheet_name=None):
    """
    JSON-блоки fn_key для текущего листа.
    Конкретные листы приоритетнее блоков без sheet (глобальных).
    """
    blocks = _lm_pp_param_decode(fn_key, extra_args)
    if len(blocks) == 0:
        return []
    if sheet_name is None and sheet is not None:
        try:
            sheet_name = str(sheet.Name).strip()
        except Exception:
            sheet_name = ""
    specific = []
    global_blocks = []
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        bi = bi + 1
        if not isinstance(block, dict):
            continue
        if _lm_pp_block_target_sheet_names(block):
            if _lm_pp_block_matches_sheet(block, doc, sheet, sheet_name):
                specific.append(block)
        else:
            global_blocks.append(block)
    if len(specific) > 0:
        return specific
    return global_blocks


def _lm_pp_pick_sheet_block_for_sheet(blocks, doc, sheet):
    """Блок с полем sheet: совпадение листа; пустой sheet — fallback для всех."""
    if not blocks:
        return None
    fallback = None
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        bi = bi + 1
        if not isinstance(block, dict):
            continue
        if not _lm_pp_block_target_sheet_names(block):
            fallback = block
            continue
        if _lm_pp_block_matches_sheet(block, doc, sheet):
            return block
    return fallback


def _lm_pp_apply_formula_spec_for_sheet(extra_args, doc, sheet):
    """
    (column, formula, as_values) для текущего листа или None (шаг пропустить).
    JSON/текст кодека или legacy «Столбец=формула» без листа.
    """
    sheet_name = sheet.Name if sheet is not None else ""
    if extra_args and isinstance(extra_args[0], dict):
        _lm_pp_apply_formula_dbg(
            doc,
            sheet_name,
            "extra_args[0]=dict keys=%s as_values=%r"
            % (list(extra_args[0].keys()), extra_args[0].get("as_values")),
        )
        block = _lm_pp_pick_sheet_block_for_sheet([extra_args[0]], doc, sheet)
    else:
        raw = _lm_pp_extra_args_rest_text(extra_args)
        _lm_pp_apply_formula_dbg(
            doc,
            sheet_name,
            "extra_args raw len=%d head=%r"
            % (len(str(raw)), str(raw)[:160]),
        )
        if raw == "":
            return None
        was_json = _lm_pp_payload_is_json(raw)
        from libre_macros_param_codec import param_decode

        blocks = param_decode("применить_формулу", raw)
        _lm_pp_apply_formula_dbg(
            doc,
            sheet_name,
            "param_decode blocks=%d first_as_values=%r"
            % (
                len(blocks),
                blocks[0].get("as_values") if blocks else None,
            ),
        )
        if len(blocks) == 0:
            if was_json:
                return None
            parsed = _lm_pp_parse_formula_column_extra(extra_args)
            if parsed is None:
                return None
            col_name, formula = parsed
            return col_name, formula, False, ""
        block = _lm_pp_pick_sheet_block_for_sheet(blocks, doc, sheet)
    if block is None:
        _lm_pp_apply_formula_dbg(doc, sheet_name, "блок для листа не найден — пропуск")
        return None
    col_name = str(block.get("column") or "").strip()
    formula = str(block.get("formula") or "").strip()
    if col_name == "" or formula == "":
        return None
    as_values = _lm_pp_parse_as_values_flag(block)
    fmt = str(block.get("format") or "").strip()
    if fmt.casefold() == "(глобальный)":
        fmt = ""
    _lm_pp_apply_formula_dbg(
        doc,
        sheet_name,
        "spec: column=%r as_values=%s format=%r (raw=%r)"
        % (col_name, as_values, fmt, block.get("as_values")),
    )
    return col_name, formula, as_values, fmt


def lm_pp_apply_formula_step_applies_to_sheet(extra_raw, sheet_name, doc=None, sheet=None):
    """True, если шаг «применить_формулу» относится к указанному листу."""
    raw = str(extra_raw or "").strip()
    if raw == "":
        return True
    extra_args = [raw]
    target = sheet
    if target is None and doc is not None:
        target = _lm_vlookup_resolve_sheet(doc, sheet_name)
    if target is not None:
        return _lm_pp_apply_formula_spec_for_sheet(extra_args, doc, target) is not None
    try:
        from libre_macros_param_codec import param_decode

        blocks = param_decode("применить_формулу", raw)
    except Exception:
        blocks = []
    if len(blocks) == 0:
        return True
    i = 0
    while i < len(blocks):
        block = blocks[i]
        sheet_ref = str(block.get("sheet") or "").strip()
        if sheet_ref == "":
            return True
        names = _lm_pp_parse_sheet_names_from_spec(sheet_ref)
        if _lm_pp_sheet_name_matches_target(sheet_name, names):
            return True
        m = re.match(r"^(\d+)_(.+)$", str(sheet_name or "").strip())
        if m is not None and _lm_pp_sheet_name_matches_target(m.group(2), names):
            return True
        i += 1
    return False


def _lm_safe_eval_apply_formula_row(expr, row_1based):
    """
    Выражение в {…} с row / row_id → номер строки Calc (1-based).

    Примеры: row, row-1, row_id+2.
    """
    s = str(expr).strip()
    sl = re.sub(r"\brow_id\b", "row", s, flags=re.IGNORECASE)
    if not re.search(r"\brow\b", sl, re.IGNORECASE):
        return None
    if re.match(r"^\s*row\s*$", sl, re.IGNORECASE):
        return int(row_1based)
    base = int(row_1based)
    s2 = re.sub(r"\brow\b", str(base), sl, flags=re.IGNORECASE)
    if not re.match(r"^[\d\+\-\*/\%\(\)\s]+$", s2):
        return None
    try:
        v = eval(s2)
        if isinstance(v, (int, float)):
            return int(v)
    except Exception:
        pass
    return None


def _lm_safe_eval_apply_formula_col(expr, col_0based):
    """
    Выражение в {…} с col / col_id → буквы столбца (A, B, …).

    col — 0-based индекс столбца вставки; col-1 — столбец левее.
    """
    s = str(expr).strip()
    sl = re.sub(r"\bcol_id\b", "col", s, flags=re.IGNORECASE)
    if not re.search(r"\bcol\b", sl, re.IGNORECASE):
        return None
    if re.match(r"^\s*col\s*$", sl, re.IGNORECASE):
        return col_index_to_letters(int(col_0based))
    base = int(col_0based)
    s2 = re.sub(r"\bcol\b", str(base), sl, flags=re.IGNORECASE)
    if not re.match(r"^[\d\+\-\*/\%\(\)\s]+$", s2):
        return None
    try:
        v = eval(s2)
        if isinstance(v, (int, float)):
            ci = int(v)
            if ci < 0:
                return None
            return col_index_to_letters(ci)
    except Exception:
        pass
    return None


def _lm_pp_normalize_r1c1_cyrillic(formula_text):
    """Кириллическая С/с в R1C1-ссылках → латинская C (защита от опечатки)."""

    def _repl(m):
        row_bracket = m.group(2)
        row_abs = m.group(3)
        col_bracket = m.group(6)
        col_abs = m.group(7)
        s = "R"
        if row_bracket is not None:
            s += "[" + row_bracket + "]"
        elif row_abs is not None:
            s += row_abs
        s += "C"
        if col_bracket is not None:
            s += "[" + col_bracket + "]"
        elif col_abs is not None:
            s += col_abs
        return s

    return _R1C1_CYR_COL_RE.sub(_repl, str(formula_text or ""))


def _lm_pp_normalize_apply_formula_template(template, uses_r1c1=False):
    """R[-1]C (синоним) → {col}{row-1}; в R1C1-документе R1C1 не трогаем."""
    s = _lm_pp_normalize_r1c1_cyrillic(template)
    s = str(s).strip()
    if not uses_r1c1:
        s = re.sub(r"R\[-1\]C", "{col}{row-1}", s, flags=re.IGNORECASE)
    return s


_APPLY_FORMULA_QUOTED_BRACKET_REF = re.compile(r"\[\s*\"([^\"]+)\"\s*\]\{([^}]+)\}")
_APPLY_FORMULA_CURLY_HEADER_REF = re.compile(
    r"\{\s*([^{}]+?)\s*\}\s*\{\s*([^{}]*(?:row_id|row)[^{}]*)\s*\}",
    re.IGNORECASE,
)
# Голый [Имя] / ["Имя"] без следующей пары {row…}
_APPLY_FORMULA_BARE_BRACKET_REF = re.compile(r"\[\s*\"([^\"]+)\"\s*\]|\[([^\]]+)\]")
_SKIP_FORMULA_BARE_BRACKET_REF = re.compile(r"\[\s*\"([^\"]+)\"\s*\]|\[([^\]]+)\]")
_APPLY_FORMULA_SOURCE_VAR_RE = re.compile(r"<<\s*Переменные\.([^>]+?)\s*>>", re.IGNORECASE)


def _lm_pp_header_to_col_token(sheet, hdr_row, start_col, col_count, name, uses_r1c1=False):
    """Имя/маркер заголовка → буква столбца (A1) или Cn (R1C1, 1-based)."""
    idx = _resolve_skip_column_spec_to_index(sheet, hdr_row, start_col, col_count, name)
    if idx < 0:
        return None
    if uses_r1c1:
        return "C%d" % (int(idx) + 1)
    return col_index_to_letters(idx)


def _lm_pp_sheet_name_candidates(sheet_name):
    sn = str(sheet_name or "").strip()
    out = []
    if sn == "":
        return out
    out.append(sn.casefold())
    m = re.match(r"^(\d+)_(.+)$", sn)
    if m is not None:
        base = str(m.group(2) or "").strip()
        if base != "":
            out.append(base.casefold())
    return out


def _lm_pp_split_args_csv(text):
    parts = []
    cur = []
    depth = 0
    for ch in str(text or ""):
        if ch == "[":
            depth += 1
        elif ch == "]" and depth > 0:
            depth -= 1
        if ch == "," and depth == 0:
            parts.append("".join(cur).strip())
            cur = []
            continue
        cur.append(ch)
    parts.append("".join(cur).strip())
    return parts


def _lm_pp_split_source_path_label(path_val):
    s = str(path_val or "").strip()
    if s == "":
        return ("", "")
    if "#" in s:
        file_part, sheet_part = s.split("#", 1)
        return (str(file_part or "").strip(), str(sheet_part or "").strip())
    return (s, "")


def _lm_pp_row_value_by_header(sheet, header_row, row_0, name):
    if sheet is None:
        return ""
    hdr_name = str(name or "").strip()
    if hdr_name == "":
        return ""
    try:
        end_col, _end_row = get_sheet_used_bounds(sheet)
    except Exception:
        end_col = 0
    idx = _resolve_skip_column_spec_to_index(sheet, int(header_row), 0, int(end_col) + 1, hdr_name)
    if idx < 0:
        return ""
    try:
        return cell_text(sheet.getCellByPosition(idx, int(row_0)))
    except Exception:
        return ""


def _lm_pp_parse_source_var_query(query, sheet=None, row_0=0, header_row=0):
    raw = str(query or "").strip()
    if raw == "":
        return {"name": "", "sheet": "", "file": "", "mode": "empty"}
    if raw.endswith(")") and "(" in raw:
        head, tail = raw.split("(", 1)
        name = str(head or "").strip()
        args_raw = str(tail[:-1] or "").strip()
        args = _lm_pp_split_args_csv(args_raw) if args_raw != "" else []
        sheet_hint = ""
        file_hint = ""
        path_file = ""
        path_sheet = ""
        ai = 0
        while ai < len(args):
            arg = str(args[ai] or "").strip()
            ai = ai + 1
            if arg == "":
                continue
            if arg.startswith("[") and arg.endswith("]"):
                arg = _lm_pp_row_value_by_header(sheet, header_row, row_0, arg[1:-1])
            if arg == "#Путь":
                arg = _lm_pp_row_value_by_header(sheet, header_row, row_0, "#Путь")
            if ai == 1:
                sheet_hint = str(arg or "").strip()
            elif ai == 2:
                file_hint = str(arg or "").strip()
        if len(args) == 1:
            path_file, path_sheet = _lm_pp_split_source_path_label(sheet_hint)
            if path_file != "" or path_sheet != "":
                file_hint = path_file
                sheet_hint = path_sheet
        return {"name": name, "sheet": sheet_hint, "file": file_hint, "mode": "call"}
    if "~" in raw:
        parts = [str(p or "").strip() for p in raw.split("~") if str(p or "").strip() != ""]
        return {
            "name": parts[0] if len(parts) > 0 else "",
            "sheet": parts[1] if len(parts) > 1 else "",
            "file": parts[2] if len(parts) > 2 else "",
            "mode": "segments",
        }
    return {"name": raw, "sheet": "", "file": "", "mode": "plain"}


def _lm_pp_source_var_matches(seg_value, hint, partial=False):
    value_cf = str(seg_value or "").strip().casefold()
    hint_cf = str(hint or "").strip().casefold()
    if hint_cf == "":
        return True
    if value_cf == hint_cf:
        return True
    if not partial:
        return False
    return (hint_cf in value_cf) or (value_cf in hint_cf)


def _lm_pp_expand_source_variables_template( text, doc=None, sheet=None, row_0=0, header_row=0, log_fn_key="применить_формулу", for_formula=True, sheet_name_hint=None):
    ctx = lm_pp_active_context() or {}
    vm = ctx.get("source_variables_map") or {}
    if not isinstance(vm, dict) or len(vm) == 0:
        return text
    order = ctx.get("source_variables_order") or []
    order_pos = {}
    oi = 0
    while oi < len(order):
        order_pos[str(order[oi])] = oi
        oi = oi + 1
    sheet_name = str(sheet_name_hint or "").strip()
    if sheet_name == "":
        try:
            if sheet is not None:
                sheet_name = str(sheet.Name or "")
        except Exception:
            sheet_name = ""
    sheet_cands = _lm_pp_sheet_name_candidates(sheet_name)

    def _pick_best(matches):
        if not matches:
            return (None, "не найдено")
        if len(matches) > 1 and sheet_cands:
            narrowed = []
            mi = 0
            while mi < len(matches):
                key = matches[mi]
                parts = [str(p or "").strip().casefold() for p in str(key or "").split("~")]
                if len(parts) > 1 and parts[1] in sheet_cands:
                    narrowed.append(key)
                mi = mi + 1
            if len(narrowed) == 1:
                matches = narrowed
            elif len(narrowed) > 1:
                matches = narrowed
        best = matches[0]
        bi = order_pos.get(str(best), -1)
        mi = 1
        while mi < len(matches):
            k = matches[mi]
            ki = order_pos.get(str(k), -1)
            if ki >= bi:
                best = k
                bi = ki
            mi = mi + 1
        val = vm.get(best)
        txt = ""
        if isinstance(val, (list, tuple)) and len(val) > 0:
            txt = str(val[0] or "")
        else:
            txt = str(val or "")
        return (txt, None if len(matches) == 1 else "несколько совпадений, выбран последний")

    def _resolve(query):
        parsed = _lm_pp_parse_source_var_query(query, sheet=sheet, row_0=row_0, header_row=header_row)
        name_hint = str(parsed.get("name") or "").strip()
        sheet_hint = str(parsed.get("sheet") or "").strip()
        file_hint = str(parsed.get("file") or "").strip()
        mode = str(parsed.get("mode") or "")
        if name_hint == "":
            return (None, "пустой запрос")
        exact = []
        partial = []
        for key in vm:
            segs = [str(p or "").strip() for p in str(key or "").split("~")]
            if len(segs) < 1:
                continue
            if segs[0].casefold() != name_hint.casefold():
                continue
            src_sheet = segs[1] if len(segs) > 1 else ""
            src_file = segs[2] if len(segs) > 2 else ""
            if mode == "segments":
                ok = True
                if sheet_hint != "" and src_sheet.casefold() != sheet_hint.casefold():
                    ok = False
                if file_hint != "" and src_file.casefold() != file_hint.casefold():
                    ok = False
                if ok:
                    exact.append(key)
                continue
            exact_ok = _lm_pp_source_var_matches(src_sheet, sheet_hint, partial=False) and _lm_pp_source_var_matches(src_file, file_hint, partial=False)
            if exact_ok:
                exact.append(key)
                continue
            partial_ok = _lm_pp_source_var_matches(src_sheet, sheet_hint, partial=True) and _lm_pp_source_var_matches(src_file, file_hint, partial=True)
            if partial_ok:
                partial.append(key)
        if exact:
            return _pick_best(exact)
        if partial:
            val, warn = _pick_best(partial)
            if val is None:
                return (None, "не найдено")
            if warn:
                return (val, "полное совпадение не найдено; частичный поиск, %s" % warn)
            return (val, "полное совпадение не найдено; использован частичный поиск")
        return (None, "не найдено")

    def _repl(match):
        q = str(match.group(1) or "").strip()
        val, warn = _resolve(q)
        if val is None:
            try:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    log_fn_key,
                    "предупреждение",
                    "Переменные.%s: не найдено совпадений" % q,
                )
            except Exception:
                pass
            return match.group(0)
        if warn:
            try:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    log_fn_key,
                    "предупреждение",
                    "Переменные.%s: %s" % (q, warn),
                )
            except Exception:
                pass
        if for_formula:
            return val.replace('"', '""')
        return val

    return _APPLY_FORMULA_SOURCE_VAR_RE.sub(_repl, str(text or ""))


_RENAME_NAME_BRACKET_RE = re.compile(r"\[([^\]]+)\]")
_RENAME_HEADER_PLACEHOLDER_RE = re.compile(
    r"(?iu)^заголовок\s*\(\s*(\d+)\s*\)$"
)
_RENAME_ROW_PLACEHOLDER_RE = re.compile(
    r"(?iu)^строка\s*\(\s*([^,\)]+)\s*,\s*(\d+)\s*\)$"
)
# [Текущая_Дата_Время_YYYYMMDD] / [Текущая_Дата_YYYY-MM-DD] / [Текущее_Время_HH-mm-SS]
# + EN: Now_… / Today_… / Current_Time_…
_RENAME_NOW_DT_FMT_RE = re.compile(
    r"(?iu)^(?:текущая_дата_время|now)_(.+)$"
)
_RENAME_NOW_DATE_FMT_RE = re.compile(
    r"(?iu)^(?:текущая_дата|today)_(.+)$"
)
_RENAME_NOW_TIME_FMT_RE = re.compile(
    r"(?iu)^(?:текущее_время|current_time)_(.+)$"
)


def _lm_pp_rename_excel_fmt_to_strftime(fmt):
    """
    Excel-подобные токены → strftime: YYYY/YY/MM/DD/HH/mm/SS.
    Порядок замен важен (длинные первыми).
    """
    s = str(fmt or "")
    if s == "":
        return ""
    # Защита литералов: временно маскируем уже занятые %-последовательности не нужны —
    # подставляем уникальные маркеры, затем strftime-коды.
    out = s
    repls = (
        ("YYYY", "\x00Y4\x00"),
        ("YY", "\x00Y2\x00"),
        ("MM", "\x00MO\x00"),
        ("DD", "\x00D2\x00"),
        ("HH", "\x00H2\x00"),
        ("mm", "\x00MI\x00"),
        ("SS", "\x00S2\x00"),
    )
    for tok, mark in repls:
        out = out.replace(tok, mark)
    out = (
        out.replace("\x00Y4\x00", "%Y")
        .replace("\x00Y2\x00", "%y")
        .replace("\x00MO\x00", "%m")
        .replace("\x00D2\x00", "%d")
        .replace("\x00H2\x00", "%H")
        .replace("\x00MI\x00", "%M")
        .replace("\x00S2\x00", "%S")
    )
    return out


def _lm_pp_rename_format_now(fmt):
    """Текущие дата/время по Excel-формату (YYYYMMDD и т.п.)."""
    import datetime

    py_fmt = _lm_pp_rename_excel_fmt_to_strftime(fmt)
    if py_fmt == "":
        return ""
    try:
        return datetime.datetime.now().strftime(py_fmt)
    except Exception:
        return ""


def _lm_pp_rename_doc_basename(doc):
    """Имя файла книги без расширения (пусто, если документ не сохранён)."""
    if doc is None:
        return ""
    try:
        url = str(doc.getURL() or "").strip()
    except Exception:
        return ""
    if url == "":
        return ""
    path = ""
    try:
        path = uno.fileUrlToSystemPath(url)
    except Exception:
        path = url
    try:
        import os

        base = os.path.basename(str(path))
        name, _ext = os.path.splitext(base)
        return str(name or "").strip()
    except Exception:
        return str(path).strip()


def _lm_pp_rename_sheet_index_1(doc, sheet):
    """1-based индекс вкладки в книге или ''."""
    if doc is None or sheet is None:
        return ""
    try:
        want = str(getattr(sheet, "Name", "") or "").strip()
    except Exception:
        want = ""
    if want == "":
        return ""
    try:
        sheets = doc.Sheets
        n = int(sheets.getCount())
    except Exception:
        return ""
    i = 0
    while i < n:
        try:
            nm = str(sheets.getByIndex(i).Name or "").strip()
        except Exception:
            nm = ""
        if nm == want:
            return str(i + 1)
        i = i + 1
    return ""


def _lm_pp_rename_guid8():
    try:
        import uuid

        return uuid.uuid4().hex[:8]
    except Exception:
        return ""


def _lm_pp_rename_template_needs_expand(text):
    s = str(text or "")
    if s == "":
        return False
    if _APPLY_FORMULA_SOURCE_VAR_RE.search(s):
        return True
    if _RENAME_NAME_BRACKET_RE.search(s):
        return True
    return False


def _lm_pp_expand_rename_name_template( template, doc=None, sheet=None, old_name="", header_row=0, used_sr=0, used_er=-1, cell_text_fn=None, log_fn_key="переименовать_лист", sheet_name_hint=None):
    """
    Развернуть шаблон имени листа/столбца:
      <<Переменные.…>>
      [Старое_имя]
      [Заголовок(N)] — N = 1-based индекс столбца, строка заголовка
      [Строка(R, C)] — R = номер / первая / последняя; C = 1-based столбец
      [Текущая_Дата_Время_YYYYMMDD] / [Текущая_Дата] / [Текущее_Время_…]
      [Имя_Книги] / [Имя_Файла] / [Номер_Листа] / [GUID8]
    """
    text = str(template or "")
    if text == "":
        return ""
    if _APPLY_FORMULA_SOURCE_VAR_RE.search(text):
        text = _lm_pp_expand_source_variables_template(
            text,
            doc=doc,
            sheet=sheet,
            log_fn_key=log_fn_key,
            for_formula=False,
            sheet_name_hint=sheet_name_hint,
        )
    if "[" not in text:
        return text

    def _read_cell(row_0, col_0):
        if cell_text_fn is not None:
            try:
                return str(cell_text_fn(int(row_0), int(col_0)) or "")
            except Exception:
                return ""
        if sheet is None:
            return ""
        try:
            return cell_text(sheet.getCellByPosition(int(col_0), int(row_0)))
        except Exception:
            return ""

    def _repl(match):
        inner = str(match.group(1) or "").strip()
        if inner == "":
            return match.group(0)
        # Пробелы → _; регистр формата YYYY/mm сохраняем (не casefold всей строки).
        norm = inner.replace(" ", "_")
        key = norm.casefold()
        if key in ("старое_имя", "old_name", "староеимя"):
            return str(old_name or "")
        if key in ("имя_книги", "имя_файла", "book_name", "file_name", "workbook"):
            return _lm_pp_rename_doc_basename(doc)
        if key in ("номер_листа", "sheet_index", "sheet_no", "sheet_number"):
            return _lm_pp_rename_sheet_index_1(doc, sheet)
        if key in ("guid8", "guid_8", "uuid8"):
            return _lm_pp_rename_guid8()
        # Shorthand дата/время без формата
        if key in ("текущая_дата", "today", "current_date"):
            return _lm_pp_rename_format_now("YYYYMMDD")
        if key in ("текущее_время", "current_time", "now_time"):
            return _lm_pp_rename_format_now("HH-mm-SS")
        if key in ("текущая_дата_время", "now", "current_datetime"):
            return _lm_pp_rename_format_now("YYYYMMDD_HHMMSS")
        # С форматом: сначала дата_время (длиннее), потом дата, потом время
        m_dt = _RENAME_NOW_DT_FMT_RE.match(norm)
        if m_dt is not None:
            fmt = str(m_dt.group(1) or "").strip()
            if fmt != "":
                return _lm_pp_rename_format_now(fmt)
        m_d = _RENAME_NOW_DATE_FMT_RE.match(norm)
        if m_d is not None:
            fmt = str(m_d.group(1) or "").strip()
            # отсечь «время_…» / «Время_…» — это ветка дата_время
            if fmt != "" and not fmt.casefold().startswith("время_"):
                return _lm_pp_rename_format_now(fmt)
        m_t = _RENAME_NOW_TIME_FMT_RE.match(norm)
        if m_t is not None:
            fmt = str(m_t.group(1) or "").strip()
            if fmt != "":
                return _lm_pp_rename_format_now(fmt)
        hm = _RENAME_HEADER_PLACEHOLDER_RE.match(inner)
        if hm is not None:
            col_1 = int(hm.group(1))
            if col_1 < 1:
                return match.group(0)
            return _read_cell(int(header_row), col_1 - 1)
        rm = _RENAME_ROW_PLACEHOLDER_RE.match(inner)
        if rm is not None:
            row_tok = str(rm.group(1) or "").strip().casefold()
            col_1 = int(rm.group(2))
            if col_1 < 1:
                return match.group(0)
            if row_tok in ("последняя", "last", "конец"):
                if int(used_er) < 0:
                    return ""
                row_0 = int(used_er)
            elif row_tok in ("первая", "first", "начало"):
                row_0 = int(used_sr)
            else:
                try:
                    row_1 = int(row_tok)
                except (TypeError, ValueError):
                    return match.group(0)
                if row_1 < 1:
                    return match.group(0)
                row_0 = row_1 - 1
            return _read_cell(row_0, col_1 - 1)
        return match.group(0)

    return _RENAME_NAME_BRACKET_RE.sub(_repl, text)


def _lm_pp_sanitize_name_chars(name, kind="sheet"):
    """Убрать недопустимые символы; длину не резать."""
    s = unicode(name or u"").strip()
    if kind == "sheet":
        out = []
        for ch in s:
            out.append(u"_" if ch in _LM_SHEET_NAME_FORBIDDEN else ch)
        s = u"".join(out).strip(u".'")
        return s
    s = re.sub(r"[\r\n\t]+", u" ", s)
    s = re.sub(r" +", u" ", s).strip()
    return s


def _lm_pp_finalize_sheet_name(name, is_taken=None):
    """
    Нормализация имени листа: запрещённые символы → _; длина ≤ 31.
    Если длиннее 31 или имя занято — stem[:28] + _01, _02, …
    """
    s = _lm_pp_sanitize_name_chars(name, kind="sheet")
    if s == u"":
        return u""

    def _taken(cand):
        if is_taken is None:
            return False
        try:
            return bool(is_taken(cand))
        except Exception:
            return True

    if len(s) <= 31 and not _taken(s):
        return s
    stem = s[:28]
    if stem == u"":
        stem = u"Sheet"
    i = 1
    while i <= 99:
        cand = u"%s_%02d" % (stem, i)
        if len(cand) > 31:
            cand = cand[:31]
        if not _taken(cand):
            return cand
        i = i + 1
    return stem[:31]


def _lm_pp_finalize_column_name(name, is_taken=None):
    """Нормализация имени столбца; при конфликте — суффикс _01, _02, …"""
    s = _lm_pp_sanitize_name_chars(name, kind="column")
    if s == u"":
        return u""

    def _taken(cand):
        if is_taken is None:
            return False
        try:
            return bool(is_taken(cand))
        except Exception:
            return True

    if not _taken(s):
        return s
    i = 1
    while i <= 99:
        cand = u"%s_%02d" % (s, i)
        if not _taken(cand):
            return cand
        i = i + 1
    return s


def _lm_pp_expand_apply_formula_template(template, row_0, col_0, doc=None, sheet=None, header_row=0):
    """
    Развернуть шаблон формулы для одной ячейки (row_0, col_0) 0-based.

    {col}{row-1} в H2 (row_0=1, col=7) → H1;
    {row_id}/{col_id} — совместимость со старыми шаблонами;
    R1C1 (R[-1]C, R[2]C[-1] …) — в A1-документе → A1 для ячейки,
    в R1C1-документе — без изменений (после нормализации кириллической С).
    """
    uses_r1c1 = _lm_pp_doc_uses_r1c1(doc) if doc is not None else False
    body = _lm_pp_normalize_apply_formula_template(template, uses_r1c1=uses_r1c1)
    if _lm_pp_formula_has_r1c1_ref(body):
        if uses_r1c1:
            s = str(body).strip()
            if not s.startswith("="):
                s = "=" + s
            return s
        s = _lm_pp_convert_r1c1_refs_to_a1(body, row_0, col_0)
        if not s.startswith("="):
            s = "=" + s
        return s
    row_1 = int(row_0) + 1
    col_ix = int(col_0)
    s = str(body)
    s = _lm_pp_expand_source_variables_template(s, doc=doc, sheet=sheet, row_0=row_0, header_row=header_row)

    # --- Расширение: ссылки на столбец по заголовку ---
    # Поддержка форм:
    #   [Заголовок]{row} / [Заголовок]{row_id}
    #   ["Заголовок"]{row} / ["Заголовок"]{row_id}
    #   {Заголовок}{row} / {Заголовок}{row_id}  (только если подряд две пары скобок)
    if sheet is not None:
        try:
            end_col, _end_row = get_sheet_used_bounds(sheet)
        except Exception:
            end_col = 0
        try:
            hdr = int(header_row)
        except Exception:
            hdr = 0
        start_col = 0
        col_count = int(end_col) + 1

        def _header_to_col_letters(name):
            return _lm_pp_header_to_col_token(
                sheet, hdr, start_col, col_count, name, uses_r1c1=uses_r1c1
            )

        def _bracket_replace(match):
            letters = _header_to_col_letters(match.group(1).strip())
            if letters is None:
                return match.group(0)
            return letters + "{" + match.group(2) + "}"

        def _quoted_bracket_replace(match):
            letters = _header_to_col_letters(match.group(1).strip())
            if letters is None:
                return match.group(0)
            return letters + "{" + match.group(2) + "}"

        def _curly_pair_replace(match):
            # {Header}{row...} → [Header]{row...} → затем обычная обработка
            header = str(match.group(1) or "").strip()
            brace = str(match.group(2) or "").strip()
            if header == "" or brace == "":
                return match.group(0)
            letters = _header_to_col_letters(header)
            if letters is None:
                return match.group(0)
            return letters + "{" + brace + "}"

        def _bare_bracket_replace(match):
            name = match.group(1) if match.group(1) is not None else match.group(2)
            name = str(name or "").strip()
            if name == "":
                return match.group(0)
            letters = _header_to_col_letters(name)
            if letters is None:
                return match.group(0)
            return letters

        s = _APPLY_FORMULA_CURLY_HEADER_REF.sub(_curly_pair_replace, s)
        s = _SKIP_FORMULA_BRACKET_REF.sub(_bracket_replace, s)
        s = _APPLY_FORMULA_QUOTED_BRACKET_REF.sub(_quoted_bracket_replace, s)
        s = _APPLY_FORMULA_BARE_BRACKET_REF.sub(_bare_bracket_replace, s)

    def _repl_brace(match):
        inner = match.group(1) if match.group(1) is not None else match.group(2)
        if inner is None:
            return match.group(0)
        inner_s = str(inner).strip()
        letters = _lm_safe_eval_apply_formula_col(inner_s, col_ix)
        if letters is not None:
            return letters
        row_num = _lm_safe_eval_apply_formula_row(inner_s, row_1)
        if row_num is not None:
            return str(row_num)
        # Голый {Имя_столбца} → буква / Cn
        if sheet is not None:
            try:
                end_col2, _er2 = get_sheet_used_bounds(sheet)
            except Exception:
                end_col2 = 0
            try:
                hdr2 = int(header_row)
            except Exception:
                hdr2 = 0
            tok = _lm_pp_header_to_col_token(
                sheet,
                hdr2,
                0,
                int(end_col2) + 1,
                inner_s,
                uses_r1c1=uses_r1c1,
            )
            if tok is not None:
                return tok
        return match.group(0)

    s = _ROW_ID_BRACE_RE.sub(_repl_brace, s)
    if not s.startswith("="):
        s = "=" + s
    return s


def _lm_pp_calc_formula_string(formula_text, first_data_row, target_col_0based=None, doc=None):
    """Формула для первой строки данных (совместимость с логом и старыми шаблонами)."""
    col_0 = int(target_col_0based) if target_col_0based is not None else 0
    return _lm_pp_expand_apply_formula_template(formula_text, first_data_row, col_0, doc=doc)


def _lm_pp_formula_has_r1c1_ref(formula_text):
    return _R1C1_REF_RE.search(
        _lm_pp_normalize_r1c1_cyrillic(formula_text)
    ) is not None


def _lm_pp_cell_is_formula_strict(cell):
    """Формула только при CellContentType.FORMULA (3) и непустом getFormula."""
    if cell is None:
        return False
    try:
        if int(cell.getType()) != 3:
            return False
    except Exception:
        return False
    try:
        f = str(cell.getFormula() or "").strip()
        if f == "":
            f = str(cell.Formula or "").strip()
        return f != ""
    except Exception:
        return False


def _lm_pp_cell_formula_written(cell):
    """
    Формула записана: непустой getFormula/Formula (или строгий тип FORMULA).

    В макросах Calc getType() иногда не FORMULA сразу после присвоения Formula —
    для «применить_формулу» достаточно непустого текста формулы.
    """
    if cell is None:
        return False
    try:
        got = str(cell.getFormula() or cell.Formula or "").strip()
        if got.startswith("="):
            return True
    except Exception:
        pass
    return _lm_pp_cell_is_formula_strict(cell)


def _lm_pp_r1c1_ref_to_offset_api(match):
    """R1C1-относительная ссылка → OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())), dr, dc)."""
    row_rel = match.group(2)
    row_abs = match.group(3)
    col_rel = match.group(5)
    col_abs = match.group(6)
    if row_abs is not None or col_abs is not None:
        return "\x00FALLBACK\x00"
    row_off = int(row_rel) if row_rel is not None else 0
    col_off = int(col_rel) if col_rel is not None else 0
    return "OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),%d,%d)" % (row_off, col_off)


def _lm_pp_formula_runtime_api(formula_text):
    """
    Одна API-формула для всех строк: ROW()/COLUMN()/OFFSET — Calc не сворачивает в константу.

    Возвращает (api_text, ok). ok=False — нужна построчная A1-запись (абсолютные R1C1 и т.п.).
    """
    body = str(formula_text or "").strip()
    if body == "":
        return "", False
    if not body.startswith("="):
        body = "=" + body
    core = body[1:]
    if "{row_id" in core.casefold() or "{col_id" in core.casefold():
        return "", False
    if not _lm_pp_formula_has_r1c1_ref(core):
        return _lm_pp_formula_separators_to_api(body), True
    converted = _R1C1_REF_RE.sub(_lm_pp_r1c1_ref_to_offset_api, core)
    if "\x00FALLBACK\x00" in converted:
        return "", False
    return _lm_pp_formula_separators_to_api("=" + converted), True


def _lm_pp_try_set_cell_formula_api(cell, api_body):
    """setFormula (API, запятая) — для runtime-формул с ROW()/OFFSET."""
    text = str(api_body or "").strip()
    if text == "":
        return False
    if not text.startswith("="):
        text = "=" + text
    try:
        cell.setFormula(text)
    except Exception:
        try:
            cell.Formula = text
        except Exception:
            return False
    return _lm_pp_cell_formula_written(cell)


def _lm_pp_formula_separators_to_api(formula_text):
    """Разделители аргументов ; → , вне строк (для cell.setFormula / Formula)."""
    s = str(formula_text)
    out = []
    i = 0
    in_str = False
    str_ch = None
    while i < len(s):
        ch = s[i]
        if in_str:
            out.append(ch)
            if ch == str_ch and (i == 0 or s[i - 1] != "\\"):
                in_str = False
            i += 1
            continue
        if ch in ('"', "'"):
            in_str = True
            str_ch = ch
            out.append(ch)
            i += 1
            continue
        if ch == ";":
            out.append(",")
            i += 1
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _lm_pp_doc_uses_r1c1(doc):
    """True, если в документе включена R1C1-нотация ссылок."""
    if doc is None:
        return False
    try:
        info = doc.getPropertySetInfo()
        if info.hasPropertyByName("FormulaReferenceStyle"):
            return int(doc.getPropertyValue("FormulaReferenceStyle")) != 0
    except Exception:
        pass
    return False


def _lm_pp_r1c1_match_to_a1(match, row_0, col_0):
    """R1C1-фрагмент → A1 для ячейки (row_0, col_0) 0-based."""
    row_0 = int(row_0)
    col_0 = int(col_0)
    row_rel = match.group(2)
    row_abs = match.group(3)
    col_rel = match.group(5)
    col_abs = match.group(6)
    if row_rel is not None:
        tr = (row_0 + 1) + int(row_rel)
    elif row_abs is not None:
        tr = int(row_abs)
    else:
        tr = row_0 + 1
    if col_rel is not None:
        tc = (col_0 + 1) + int(col_rel)
    elif col_abs is not None:
        tc = int(col_abs)
    else:
        tc = col_0 + 1
    if tr < 1 or tc < 1:
        return match.group(0)
    return col_index_to_letters(tc - 1) + str(tr)


def _lm_pp_convert_r1c1_refs_to_a1(formula_text, row_0, col_0):
    def _repl(m):
        return _lm_pp_r1c1_match_to_a1(m, row_0, col_0)

    return _R1C1_REF_RE.sub(_repl, str(formula_text))


def _lm_pp_convert_same_col_above_a1_to_r1c1(formula_text, row_0, col_0):
    """Ссылка на ячейку выше в том же столбце (A1) → R[-1]C для R1C1-документа."""
    letters = col_index_to_letters(int(col_0))
    ref = letters + str(int(row_0))
    return re.sub(
        r"(?<![A-ZА-Я])" + re.escape(ref) + r"(?!\d)",
        "R[-1]C",
        str(formula_text),
        flags=re.IGNORECASE,
    )


def _lm_pp_prepare_formula_for_doc(formula_text, doc, row_0, col_0, preserve_r1c1=False):
    """Привести текст формулы к нотации документа (A1 или R1C1)."""
    body = _lm_pp_normalize_r1c1_cyrillic(formula_text)
    body = str(body).strip()
    if body == "":
        return body
    if row_0 is None or col_0 is None:
        return body
    has_r1c1 = _lm_pp_formula_has_r1c1_ref(body)
    if preserve_r1c1 and has_r1c1:
        return body
    uses_r1c1 = _lm_pp_doc_uses_r1c1(doc)
    if has_r1c1 and not uses_r1c1:
        return _lm_pp_convert_r1c1_refs_to_a1(body, row_0, col_0)
    if uses_r1c1 and not has_r1c1:
        return _lm_pp_convert_same_col_above_a1_to_r1c1(body, row_0, col_0)
    return body


def _lm_pp_cell_has_formula_error(cell):
    if cell is None:
        return True
    try:
        t = cell_text(cell).strip()
        if t.startswith("Err:") or t.startswith("#"):
            return True
    except Exception:
        pass
    try:
        err = int(cell.Error)
        if err != 0:
            return True
    except Exception:
        pass
    return False


_LM_FORMULA_EN_TO_RU = ()
_LM_FORMULA_BARE_TOKENS = frozenset()
try:
    from libre_macros_formula_names_ru import FORMULA_EN_TO_RU as _LM_FORMULA_EN_TO_RU
    from libre_macros_formula_names_ru import FORMULA_BARE_TOKENS as _LM_FORMULA_BARE_TOKENS
except Exception:
    # Запасной минимум, если модуль карты не установлен.
    _LM_FORMULA_EN_TO_RU = (
        ("IFERROR", "ЕСЛИОШИБКА"),
        ("ISBLANK", "ЕПУСТО"),
        ("IF", "ЕСЛИ"),
        ("SUM", "СУММ"),
        ("VLOOKUP", "ВПР"),
        ("LEN", "ДЛСТР"),
        ("AND", "И"),
        ("OR", "ИЛИ"),
        ("NOT", "НЕ"),
        ("TRUE", "ИСТИНА"),
        ("FALSE", "ЛОЖЬ"),
        ("NA", "НД"),
    )
    _LM_FORMULA_BARE_TOKENS = frozenset(("TRUE", "FALSE", "PI", "NA"))


def _lm_pp_formula_yo_variants(name):
    """Варианты с ё/е для русских имён (СЧЁТЕСЛИ / СЧЕТЕСЛИ)."""
    s = str(name or "")
    out = [s]
    if u"ё" in s or u"Ё" in s:
        out.append(s.replace(u"ё", u"е").replace(u"Ё", u"Е"))
    return out


def _lm_pp_formula_funcs_to_local(formula_text):
    """Английские имена функций → русские (для FormulaLocal)."""
    s = str(formula_text or "")
    pairs = sorted(_LM_FORMULA_EN_TO_RU, key=lambda p: len(p[0]), reverse=True)
    bare = _LM_FORMULA_BARE_TOKENS or frozenset()
    for en, ru in pairs:
        if en in bare:
            s = re.sub(r"(?i)\b" + re.escape(en) + r"\b", ru, s)
        else:
            s = re.sub(
                r"(?i)\b" + re.escape(en) + r"(?=\s*\()",
                ru,
                s,
            )
    return s


def _lm_pp_formula_funcs_to_api(formula_text):
    """Русские имена функций → английские (для Formula / setFormula)."""
    s = str(formula_text or "")
    pairs = sorted(_LM_FORMULA_EN_TO_RU, key=lambda p: len(p[1]), reverse=True)
    bare = _LM_FORMULA_BARE_TOKENS or frozenset()
    for en, ru in pairs:
        variants = _lm_pp_formula_yo_variants(ru)
        for ru_v in variants:
            if en in bare:
                s = re.sub(r"(?i)\b" + re.escape(ru_v) + r"\b", en, s)
            else:
                s = re.sub(
                    r"(?i)\b" + re.escape(ru_v) + r"(?=\s*\()",
                    en,
                    s,
                )
    return s


def _lm_pp_put_cell_formula_property(cell, formula_text, allow_formula_error=False):
    """
    Записать формулу только через свойство Formula (как ручной ввод в Calc).

    Без setFormula, FormulaLocal и преобразований нотации.
    """
    if cell is None:
        return False
    text = str(formula_text or "").strip()
    if text == "":
        return False
    if not text.startswith("="):
        text = "=" + text
    candidates = [text]
    alt = _lm_pp_formula_separators_to_api(text)
    if alt != text:
        candidates.append(alt)
    for candidate in candidates:
        written = False
        try:
            cell.setFormula(candidate)
            written = True
        except Exception:
            pass
        if not written:
            try:
                cell.Formula = candidate
            except Exception:
                continue
        if not _lm_pp_cell_formula_written(cell):
            continue
        if not allow_formula_error and _lm_pp_cell_has_formula_error(cell):
            continue
        return True
    return False


def _lm_pp_try_set_cell_formula(cell, text, use_local=False, allow_formula_error=False):
    """Записать формулу: setFormula / Formula / FormulaLocal."""
    if cell is None or text is None or str(text).strip() == "":
        return False
    payload = str(text)
    try:
        if use_local:
            cell.FormulaLocal = payload
        else:
            try:
                cell.setFormula(payload)
            except Exception:
                cell.Formula = payload
    except Exception:
        return False
    if not _lm_pp_cell_formula_written(cell):
        return False
    if not allow_formula_error and _lm_pp_cell_has_formula_error(cell):
        return False
    return True


def _lm_pp_set_cell_formula( doc, cell, formula_text, row_0=None, col_0=None, sheet_name="", preserve_r1c1=False, allow_formula_error=False):
    """
    Записать формулу в ячейку.

    Порядок: setFormula/Formula (англ. имена, «;») → API с «,» → FormulaLocal
    (локализованные имена: ЕСЛИОШИБКА, ДЛСТР и т.д.).
    Русские имена из шаблона сначала переводятся в EN для Formula API —
    иначе ДЛСТР через Formula даёт #ИМЯ? и as_values «замораживает» ошибку.
    """
    raw = str(formula_text or "").strip()
    if raw == "":
        return False
    if not raw.startswith("="):
        raw = "=" + raw

    body = _lm_pp_prepare_formula_for_doc(
        raw, doc, row_0, col_0, preserve_r1c1=preserve_r1c1
    )
    api_body = _lm_pp_formula_funcs_to_api(body)
    uses_r1c1 = _lm_pp_doc_uses_r1c1(doc)
    _lm_pp_apply_formula_dbg(
        doc,
        sheet_name,
        "запись формулы: doc_R1C1=%s → %r api=%r"
        % (uses_r1c1, body[:120], api_body[:120]),
    )

    # 1. API: английские имена, «;» (рус. локаль Calc)
    if _lm_pp_try_set_cell_formula(
        cell, api_body, use_local=False, allow_formula_error=allow_formula_error
    ):
        return True

    # 2. API: запятая (англ. локаль документа)
    api_comma = _lm_pp_formula_separators_to_api(api_body)
    if api_comma != api_body and _lm_pp_try_set_cell_formula(
        cell, api_comma, use_local=False, allow_formula_error=allow_formula_error
    ):
        _lm_pp_apply_formula_dbg(
            doc, sheet_name, "запись через API (,): %r" % (api_comma[:120],)
        )
        return True

    # 3. FormulaLocal: локализованные имена
    local_body = _lm_pp_formula_funcs_to_local(api_body)
    if local_body != api_body and _lm_pp_try_set_cell_formula(
        cell, local_body, use_local=True, allow_formula_error=allow_formula_error
    ):
        _lm_pp_apply_formula_dbg(
            doc, sheet_name, "запись FormulaLocal: %r" % (local_body[:120],)
        )
        return True

    # 4. FormulaLocal как в шаблоне (уже русские имена)
    if body != api_body and _lm_pp_try_set_cell_formula(
        cell, body, use_local=True, allow_formula_error=allow_formula_error
    ):
        return True
    if _lm_pp_try_set_cell_formula(
        cell, body, use_local=True, allow_formula_error=allow_formula_error
    ):
        return True

    # 5. Запас: Formula через свойство (setFormula + Formula, оба разделителя)
    if _lm_pp_put_cell_formula_property(
        cell, api_body, allow_formula_error=allow_formula_error
    ):
        _lm_pp_apply_formula_dbg(
            doc, sheet_name, "запись Formula (relaxed): %r" % (api_body[:120],)
        )
        return True

    if row_0 is not None and col_0 is not None and not preserve_r1c1:
        if uses_r1c1 and body != raw:
            if _lm_pp_try_set_cell_formula(
                cell, raw, use_local=True, allow_formula_error=allow_formula_error
            ):
                _lm_pp_apply_formula_dbg(doc, sheet_name, "запас A1 local: %r" % (raw[:120],))
                return True
        elif not uses_r1c1:
            if _lm_pp_formula_has_r1c1_ref(raw):
                alt = _lm_pp_convert_r1c1_refs_to_a1(raw, row_0, col_0)
            else:
                alt = _lm_pp_convert_same_col_above_a1_to_r1c1(raw, row_0, col_0)
            if alt != body and _lm_pp_try_set_cell_formula(
                cell, alt, use_local=True, allow_formula_error=allow_formula_error
            ):
                _lm_pp_apply_formula_dbg(doc, sheet_name, "запас R1C1 local: %r" % (alt[:120],))
                return True

    return False


def _lm_pp_resolve_derived_column(sheet, data_range, header_row_range, col_name):
    """
    Найти/добавить столбец и координаты для производных range-колбэков (конкатенация, формула).

    Возвращает dict или None.
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    ref_col = ec if ec >= h_sc else h_sc
    target_col = _lm_pp_find_or_append_header_column(sheet, col_name, h_sr)
    if target_col < 0:
        return None
    return {
        "sc": sc,
        "sr": sr,
        "ec": ec,
        "er": er,
        "h_sr": h_sr,
        "ref_col": ref_col,
        "target_col": target_col,
    }


def _lm_pp_finalize_derived_column(sheet, ref_col, target_col, header_row, data_top, doc=None):
    """
    После заполнения производного столбца: стиль заголовка, автоширина, расширение end_col.
    """
    _lm_pp_copy_column_format_subrange(
        sheet, ref_col, target_col, header_row, header_row
    )
    # Для добавленного столбца клонируем и формат данных до конца used area:
    # заливка/шрифт/горизонталь/вертикаль и прочие свойства столбца-донора.
    try:
        _ec_used, er_used = get_sheet_used_bounds(sheet)
    except Exception:
        er_used = int(data_top)
    try:
        data_top = int(data_top)
    except Exception:
        data_top = int(header_row) + 1
    if int(er_used) >= int(data_top):
        _lm_pp_copy_column_format_subrange(
            sheet, ref_col, target_col, int(data_top), int(er_used)
        )
    _lm_pp_autofit_columns(None, sheet, [target_col])
    if doc is not None:
        try:
            _lm_pp_calculate_doc(doc)
        except Exception:
            pass
    ctx = lm_pp_active_context()
    if ctx is not None:
        _lm_pp_context_refresh(ctx)
        _lm_pp_context_expand_end_col(ctx, target_col)


def _lm_pp_select_cell_range(doc, sheet, cell_range):
    """Выделить поддиапазон на листе (нужно для Copy / FillDown через dispatch)."""
    if doc is None or sheet is None or cell_range is None:
        return False
    try:
        controller = doc.getCurrentController()
        if controller is None:
            return False
        try:
            active = controller.getActiveSheet()
            if active is None or active.Name != sheet.Name:
                controller.setActiveSheet(sheet)
        except Exception:
            try:
                controller.setActiveSheet(sheet)
            except Exception:
                pass
        controller.select(cell_range)
        return True
    except Exception:
        return False


def _lm_pp_paste_cell_range_formulas_only(sheet, dest_range):
    """
    Специальная вставка из буфера: только формулы (CellFlags.FORMULA).

    Буфер должен содержать скопированный поддиапазон ячеек (XSheetPastable.pasteCellRange).
    """
    if sheet is None or dest_range is None:
        return False
    paste_op = 0
    insert_mode = 0
    formula_flag = 16
    try:
        from com.sun.star.sheet.PasteOperation import NONE as PASTE_NONE
        from com.sun.star.sheet.CellInsertMode import NONE as INSERT_NONE
        from com.sun.star.sheet.CellFlags import FORMULA as FORMULA_FLAG

        paste_op = PASTE_NONE
        insert_mode = INSERT_NONE
        formula_flag = FORMULA_FLAG
    except Exception:
        pass
    dest_addr = dest_range.getRangeAddress()
    try:
        sheet.pasteCellRange(
            dest_addr,
            paste_op,
            formula_flag,
            False,
            False,
            False,
            insert_mode,
        )
        return True
    except Exception:
        pass
    try:
        from com.sun.star.sheet import XSheetPastable

        pastable = sheet.queryInterface(XSheetPastable)
        if pastable is not None:
            pastable.pasteCellRange(
                dest_addr,
                paste_op,
                formula_flag,
                False,
                False,
                False,
                insert_mode,
            )
            return True
    except Exception:
        pass
    return False


def _lm_pp_paste_cell_range_values_only(sheet, dest_range):
    """
    Специальная вставка из буфера: только значения (без формул и оформления).
    """
    if sheet is None or dest_range is None:
        return False
    paste_op = 0
    insert_mode = 0
    value_flags = 7
    try:
        from com.sun.star.sheet.PasteOperation import NONE as PASTE_NONE
        from com.sun.star.sheet.CellInsertMode import NONE as INSERT_NONE
        from com.sun.star.sheet.CellFlags import VALUE, STRING, DATETIME

        paste_op = PASTE_NONE
        insert_mode = INSERT_NONE
        value_flags = int(VALUE) | int(STRING) | int(DATETIME)
    except Exception:
        pass
    dest_addr = dest_range.getRangeAddress()
    try:
        sheet.pasteCellRange(
            dest_addr,
            paste_op,
            value_flags,
            False,
            False,
            False,
            insert_mode,
        )
        return True
    except Exception:
        pass
    try:
        from com.sun.star.sheet import XSheetPastable

        pastable = sheet.queryInterface(XSheetPastable)
        if pastable is not None:
            pastable.pasteCellRange(
                dest_addr,
                paste_op,
                value_flags,
                False,
                False,
                False,
                insert_mode,
            )
            return True
    except Exception:
        pass
    return False


def _lm_pp_paste_cell_range_values_and_formats(sheet, dest_range):
    """Специальная вставка: значения и оформление (без формул)."""
    if sheet is None or dest_range is None:
        return False
    paste_op = 0
    insert_mode = 0
    flags = 7 | 32 | 64 | 512
    try:
        from com.sun.star.sheet.PasteOperation import NONE as PASTE_NONE
        from com.sun.star.sheet.CellInsertMode import NONE as INSERT_NONE
        from com.sun.star.sheet.CellFlags import (
            VALUE,
            STRING,
            DATETIME,
            HARDATTR,
            STYLES,
            FORMATTED,
        )

        paste_op = PASTE_NONE
        insert_mode = INSERT_NONE
        flags = int(VALUE) | int(STRING) | int(DATETIME) | int(HARDATTR) | int(
            STYLES
        ) | int(FORMATTED)
    except Exception:
        pass
    dest_addr = dest_range.getRangeAddress()

    def _try_paste(target):
        try:
            target.pasteCellRange(
                dest_addr,
                paste_op,
                flags,
                False,
                False,
                False,
                insert_mode,
            )
            return True
        except Exception:
            return False

    if _try_paste(sheet):
        return True
    try:
        from com.sun.star.sheet import XSheetPastable

        pastable = sheet.queryInterface(XSheetPastable)
        if pastable is not None:
            return _try_paste(pastable)
    except Exception:
        pass
    return False


def _lm_pp_select_entire_sheet(doc, sheet):
    """Выделить значащую область листа (надёжнее SelectAll для DataPilot)."""
    if doc is None or sheet is None:
        return False
    sc, sr, ec, er = _lm_vlookup_sheet_used_area(sheet)
    if er >= sr and ec >= sc:
        return _lm_pp_select_cell_range(
            doc, sheet, sheet.getCellRangeByPosition(sc, sr, ec, er)
        )
    try:
        controller = doc.getCurrentController()
        if controller is None:
            return False
        try:
            active = controller.getActiveSheet()
            if active is None or str(active.Name) != str(sheet.Name):
                controller.setActiveSheet(sheet)
        except Exception:
            controller.setActiveSheet(sheet)
        return _lm_pp_execute_dispatch(doc, ".uno:SelectAll")
    except Exception:
        return False


def _lm_pp_copy_range_to_clipboard(doc, sheet, cell_range):
    """Выделить диапазон и скопировать в буфер обмена."""
    if not _lm_pp_select_cell_range(doc, sheet, cell_range):
        return False
    return _lm_pp_execute_dispatch(doc, ".uno:Copy")


def _lm_pp_insert_contents_formula_props():
    """PropertyValue для .uno:InsertContents — вставка только формул (Flags=32)."""
    props = []
    for name, value in (
        ("Flags", 32),
        ("FormulaCommand", 0),
        ("SkipEmptyCells", False),
        ("Transpose", False),
    ):
        prop = PropertyValue()
        prop.Name = name
        prop.Value = value
        props.append(prop)
    return tuple(props)


def _lm_pp_insert_contents_values_props():
    """PropertyValue для .uno:InsertContents — вставка только значений (SVD)."""
    props = []
    for name, value in (
        ("Flags", "SVD"),
        ("FormulaCommand", 0),
        ("SkipEmptyCells", False),
        ("Transpose", False),
        ("AsLink", False),
        ("MoveMode", 4),
        ("Overwrite", True),
    ):
        prop = PropertyValue()
        prop.Name = name
        prop.Value = value
        props.append(prop)
    return tuple(props)


def _lm_pp_clipboard_materialize_range_as_values( doc, sheet, cell_range, sheet_name="", dbg_channel="formula_fill", uno_cleanup=True ):
    """
    Формулы → значения: Copy диапазона → InsertContents (SVD) в тот же диапазон.
    """
    if doc is None or sheet is None or cell_range is None:
        _lm_pp_formula_fill_dbg(
            doc,
            sheet_name,
            "materialize: нет doc/sheet/range",
            channel=dbg_channel,
        )
        return False

    _lm_pp_calculate_with_auto(doc, sheet)

    range_ref = None
    insert_props = None
    try:
        range_ref = cell_range
        if not _lm_pp_select_cell_range(doc, sheet, range_ref):
            _lm_pp_formula_fill_dbg(
                doc, sheet_name, "materialize: select FAIL", channel=dbg_channel
            )
            return False
        if not _lm_pp_execute_dispatch(doc, ".uno:Copy", uno_cleanup=uno_cleanup):
            _lm_pp_formula_fill_dbg(
                doc, sheet_name, "materialize: .uno:Copy FAIL", channel=dbg_channel
            )
            return False
        if not _lm_pp_select_cell_range(doc, sheet, range_ref):
            _lm_pp_formula_fill_dbg(
                doc,
                sheet_name,
                "materialize: re-select FAIL",
                channel=dbg_channel,
            )
            return False
        insert_props = _lm_pp_insert_contents_values_props()
        if not _lm_pp_execute_dispatch(
            doc, ".uno:InsertContents", insert_props, uno_cleanup=uno_cleanup
        ):
            _lm_pp_formula_fill_dbg(
                doc,
                sheet_name,
                "materialize: .uno:InsertContents FAIL",
                channel=dbg_channel,
            )
            return False
        _lm_pp_formula_fill_dbg(
            doc, sheet_name, "materialize: InsertContents SVD OK", channel=dbg_channel
        )
        return True
    finally:
        if uno_cleanup:
            try:
                del range_ref
            except Exception:
                pass
            try:
                del insert_props
            except Exception:
                pass
            _lm_pp_uno_gc(True)


def _lm_pp_clipboard_fill_formulas_down( doc, sheet, col, top_row, bottom_row, sheet_name="", dbg_channel="formula_fill", uno_cleanup=True):
    """
    Протянуть формулу вниз: Copy ячейки-образца → InsertContents (только формулы).

    Проверенный путь для макросов Calc (быстрее copyRange/fillAuto/postрочной записи).
    """
    if sheet is None or bottom_row <= top_row:
        return True
    try:
        col = int(col)
        top_row = int(top_row)
        bottom_row = int(bottom_row)
    except (TypeError, ValueError):
        return False

    if _lm_pp_apply_formula_rows_have_formulas(sheet, col, top_row, bottom_row):
        _lm_pp_formula_fill_dbg(
            doc,
            sheet_name,
            "clipboard_fill: уже заполнено col=%d rows %d..%d"
            % (col + 1, top_row + 1, bottom_row + 1),
            channel=dbg_channel,
        )
        return True

    source_range = None
    dest_range = None
    insert_props = None
    try:
        source_range = sheet.getCellRangeByPosition(col, top_row, col, top_row)
        dest_range = sheet.getCellRangeByPosition(
            col, top_row + 1, col, bottom_row
        )

        if not _lm_pp_select_cell_range(doc, sheet, source_range):
            _lm_pp_formula_fill_dbg(
                doc, sheet_name, "clipboard_fill: select source FAIL", channel=dbg_channel
            )
            return False
        if not _lm_pp_execute_dispatch(doc, ".uno:Copy", uno_cleanup=uno_cleanup):
            _lm_pp_formula_fill_dbg(
                doc, sheet_name, "clipboard_fill: .uno:Copy FAIL", channel=dbg_channel
            )
            return False
        if not _lm_pp_select_cell_range(doc, sheet, dest_range):
            _lm_pp_formula_fill_dbg(
                doc, sheet_name, "clipboard_fill: select dest FAIL", channel=dbg_channel
            )
            return False
        insert_props = _lm_pp_insert_contents_formula_props()
        if not _lm_pp_execute_dispatch(
            doc, ".uno:InsertContents", insert_props, uno_cleanup=uno_cleanup
        ):
            _lm_pp_formula_fill_dbg(
                doc,
                sheet_name,
                "clipboard_fill: .uno:InsertContents FAIL",
                channel=dbg_channel,
            )
            return False

        need = bottom_row - top_row
        got = _lm_pp_count_formula_cells(sheet, col, top_row + 1, bottom_row)
        ok = got >= need
        _lm_pp_formula_fill_dbg(
            doc,
            sheet_name,
            "clipboard_fill: InsertContents col=%d rows %d..%d formulas=%d/%d %s"
            % (
                col + 1,
                top_row + 2,
                bottom_row + 1,
                got,
                need,
                "OK" if ok else "FAIL",
            ),
            channel=dbg_channel,
        )
        return ok
    finally:
        if uno_cleanup:
            try:
                del source_range
            except Exception:
                pass
            try:
                del dest_range
            except Exception:
                pass
            try:
                del insert_props
            except Exception:
                pass
            _lm_pp_uno_gc(True)


def _lm_pp_copy_sheet_to_clipboard(doc, sheet):
    """Скопировать значащую область листа в буфер обмена."""
    if doc is None or sheet is None:
        return False
    sc, sr, ec, er = _lm_vlookup_sheet_used_area(sheet)
    if er >= sr and ec >= sc:
        cell_range = sheet.getCellRangeByPosition(sc, sr, ec, er)
        if _lm_pp_copy_range_to_clipboard(doc, sheet, cell_range):
            return True
    return (
        _lm_pp_select_entire_sheet(doc, sheet)
        and _lm_pp_execute_dispatch(doc, ".uno:Copy")
    )


def _lm_pp_freeze_sheet_header(doc, sheet, header_row_0based):
    """Закрепить строки 0…header_row (без флага MERGE_FREEZE_HEADER)."""
    if doc is None or sheet is None:
        return
    try:
        freeze_rows = int(header_row_0based) + 1
        if freeze_rows < 1:
            freeze_rows = 1
        controller = doc.getCurrentController()
        if controller is None:
            return
        if not _lm_pp_controller_prepare_freeze(controller, sheet):
            return
        controller.freezeAtPosition(0, freeze_rows)
    except Exception:
        pass


def _lm_pp_sheet_autofilter_safe_name(sheet):
    """Безопасный хвост имени DatabaseRange для листа (до 24 символов)."""
    safe = "".join(
        ch if ch.isalnum() or ch in ("_", "-") else "_"
        for ch in (sheet.Name if sheet is not None else u"")
    )
    if len(safe) > 24:
        safe = safe[:24]
    return safe


def _lm_pp_sheet_index(sheet):
    """Индекс листа в книге (Sheet из RangeAddress) или -1."""
    if sheet is None:
        return -1
    try:
        return int(sheet.getCellRangeByPosition(0, 0, 0, 0).getRangeAddress().Sheet)
    except Exception:
        return -1


def _lm_pp_doc_unnamed_database_ranges(doc):
    """
    XUnnamedDatabaseRanges документа (меню Data→AutoFilter хранит состояние здесь).

    Именованные doc.DatabaseRanges — другое хранилище (AF_*/смарт-таблицы).
    """
    if doc is None:
        return None
    try:
        udr = doc.getPropertyValue("UnnamedDatabaseRanges")
        if udr is not None:
            return udr
    except Exception:
        pass
    try:
        return getattr(doc, "UnnamedDatabaseRanges", None)
    except Exception:
        return None


def _lm_pp_db_range_autofilter_on(db):
    """True, если у DatabaseRange включён AutoFilter."""
    if db is None:
        return False
    try:
        return bool(getattr(db, "AutoFilter", False))
    except Exception:
        pass
    try:
        return bool(db.getPropertyValue("AutoFilter"))
    except Exception:
        return False


def _lm_pp_smart_table_keep_sheets():
    """Набор имён листов-исключений из очистки smart-диапазонов."""
    try:
        import libre_macros_collect_cfg as _cw_cfg_pp

        return set(getattr(_cw_cfg_pp, "_MERGE_KEEP_SMART_TABLE_SHEETS", set()) or set())
    except Exception:
        return set()


def _lm_pp_set_sheet_keep_smart_table(sheet_name, enabled=True):
    """Добавить/убрать лист в исключения очистки smart-диапазонов."""
    sn = str(sheet_name or "").strip()
    if sn == "":
        return
    try:
        import libre_macros_collect_cfg as _cw_cfg_pp

        cur = set(getattr(_cw_cfg_pp, "_MERGE_KEEP_SMART_TABLE_SHEETS", set()) or set())
        if enabled:
            cur.add(sn)
        else:
            cur.discard(sn)
        _cw_cfg_pp._MERGE_KEEP_SMART_TABLE_SHEETS = cur
    except Exception:
        pass


def _lm_pp_remember_smart_table_style(sheet_name, style_name):
    """Запомнить стиль smart-таблицы листа (для повторного назначения после reopen)."""
    sn = str(sheet_name or "").strip()
    st = str(style_name or "").strip() or "TableStyleLight1"
    if sn == "":
        return
    try:
        import libre_macros_collect_cfg as _cw_cfg_pp

        cur = dict(getattr(_cw_cfg_pp, "_MERGE_SMART_TABLE_STYLES", {}) or {})
        cur[sn] = st
        _cw_cfg_pp._MERGE_SMART_TABLE_STYLES = cur
        _cw_cfg_pp._MERGE_NEED_SMART_TABLE_REOPEN = True
    except Exception:
        pass


def _lm_pp_take_smart_table_styles():
    """Забрать и очистить карту sheet→style для финального назначения после reopen."""
    try:
        import libre_macros_collect_cfg as _cw_cfg_pp

        cur = dict(getattr(_cw_cfg_pp, "_MERGE_SMART_TABLE_STYLES", {}) or {})
        _cw_cfg_pp._MERGE_SMART_TABLE_STYLES = {}
        return cur
    except Exception:
        return {}


def _lm_pp_find_sheet_database_ranges(doc, sheet):
    """Список DatabaseRange на листе (имя, объект)."""
    found = []
    if doc is None or sheet is None:
        return found
    want = _lm_pp_sheet_index(sheet)
    if want < 0:
        return found
    try:
        dbr = doc.DatabaseRanges
        n = int(dbr.getCount())
    except Exception:
        return found
    i = 0
    while i < n:
        try:
            db = dbr.getByIndex(i)
            addr = db.getDataArea()
            if int(addr.Sheet) == want:
                found.append((str(getattr(db, "Name", "") or ""), db))
        except Exception:
            pass
        i = i + 1
    return found


def _lm_pp_reapply_smart_table_styles(doc, styles_by_sheet=None):
    """
    После reopen повторно назначить TableStyle на smart-диапазоны листов.

    styles_by_sheet: {имя_листа: TableStyleLight1, ...}
    Возвращает список строк-заметок для журнала.
    """
    notes = []
    if doc is None:
        return notes
    styles = dict(styles_by_sheet or {})
    if len(styles) == 0:
        styles = _lm_pp_take_smart_table_styles()
    if len(styles) == 0:
        return notes
    sheets = None
    try:
        sheets = doc.getSheets()
    except Exception:
        sheets = None
    for sn, style_name in styles.items():
        sn = str(sn or "").strip()
        style_name = str(style_name or "").strip() or "TableStyleLight1"
        if sn == "":
            continue
        sheet = None
        try:
            if sheets is not None and sheets.hasByName(sn):
                sheet = sheets.getByName(sn)
        except Exception:
            sheet = None
        if sheet is None:
            notes.append("%s: лист не найден" % sn)
            continue
        applied = 0
        for db_name, db in _lm_pp_find_sheet_database_ranges(doc, sheet):
            name_cf = str(db_name or "").strip().casefold()
            # Предпочитаем AF_PP_ / AF_; иначе любой диапазон с AutoFilter на листе.
            prefer = name_cf.startswith("af_pp_") or name_cf.startswith("af_")
            if not prefer and not _lm_pp_db_range_autofilter_on(db):
                continue
            rng = None
            prep_state = {"header": [], "cols": [], "sc": 0, "sr": 0, "ec": 0, "er": 0}
            try:
                rng = _lm_pp_db_range_to_cell_range(doc, sheet, db)
                if rng is not None:
                    prep_state = _lm_pp_prepare_range_for_smart_table_style(
                        doc, sheet, rng
                    )
            except Exception:
                prep_state = {"header": [], "cols": [], "sc": 0, "sr": 0, "ec": 0, "er": 0}
            if lm_pp_apply_db_range_table_style(db, style_name, True):
                applied = applied + 1
            # Повтор после короткого «пинга» свойства AutoFilter (AO иногда сбрасывает стиль).
            try:
                db.AutoFilter = True
            except Exception:
                try:
                    db.setPropertyValue("AutoFilter", True)
                except Exception:
                    pass
            lm_pp_apply_db_range_table_style(db, style_name, True)
            if rng is not None:
                _lm_pp_finish_range_after_smart_table_style(sheet, rng, prep_state)
        if applied > 0:
            notes.append("%s→%s (%d)" % (sn, style_name, applied))
        else:
            notes.append("%s: диапазон не найден для %s" % (sn, style_name))
    return notes


def _lm_pp_sheet_has_active_autofilter(doc, sheet):
    """
    True, если на листе уже включён AutoFilter.

    1) UnnamedDatabaseRanges.getByTable(sheet) — меню Data→AutoFilter
       (основной путь; не создаёт AF_/смарт-таблицы).
    2) Именованные DatabaseRanges с AutoFilter=True на этом листе (AF_*/AF_PP_*).
    """
    if doc is None or sheet is None:
        return False
    want = _lm_pp_sheet_index(sheet)
    if want < 0:
        return False
    # 1) Безымянный диапазон листа (меню-автофильтр).
    try:
        udr = _lm_pp_doc_unnamed_database_ranges(doc)
        if udr is not None:
            has = False
            try:
                has = bool(udr.hasByTable(int(want)))
            except Exception:
                has = False
            if has:
                try:
                    db = udr.getByTable(int(want))
                    if _lm_pp_db_range_autofilter_on(db):
                        return True
                except Exception:
                    pass
    except Exception:
        pass
    # 2) Именованные AF_* / AF_PP_* / __Anonymous_Sheet_DB__* в DatabaseRanges.
    try:
        dbr = doc.DatabaseRanges
        n = int(dbr.getCount())
    except Exception:
        return False
    i = 0
    while i < n:
        try:
            db = dbr.getByIndex(i)
        except Exception:
            i = i + 1
            continue
        i = i + 1
        if not _lm_pp_db_range_autofilter_on(db):
            continue
        try:
            addr = db.getDataArea()
            if int(addr.Sheet) == want:
                return True
        except Exception:
            pass
    return False


def _lm_pp_disable_sheet_menu_autofilter(doc, sheet):
    """
    Выключить меню-автофильтр текущего листа через UnnamedDatabaseRanges.

    Возвращает True, если операция выполнена без ошибок API.
    """
    if doc is None or sheet is None:
        return False
    want = _lm_pp_sheet_index(sheet)
    if want < 0:
        return False
    try:
        udr = _lm_pp_doc_unnamed_database_ranges(doc)
        if udr is None:
            return False
        if not bool(udr.hasByTable(int(want))):
            return True
        db = udr.getByTable(int(want))
        try:
            db.AutoFilter = False
        except Exception:
            db.setPropertyValue("AutoFilter", False)
        return True
    except Exception:
        return False


def _lm_pp_ensure_sheet_menu_autofilter(doc, sheet, sc, sr, ec, er):
    """
    Включить меню-автофильтр на диапазоне без toggle-dispatch.

    Предпочитает UnnamedDatabaseRanges.setByTable + AutoFilter=True
    (как внутренний механизм Calc, без AF_/TableStyle).
    Если уже включён — no-op True.
    Fallback — .uno:DataFilterAutoFilter (только если детект сказал «нет»).
    """
    if doc is None or sheet is None:
        return False
    try:
        sc = int(sc)
        sr = int(sr)
        ec = int(ec)
        er = int(er)
    except (TypeError, ValueError):
        return False
    if er < sr or ec < sc:
        return False
    if _lm_pp_sheet_has_active_autofilter(doc, sheet):
        return True
    try:
        _lm_pp_fill_empty_header_cells(sheet, sr, sc, ec)
    except Exception:
        pass
    want = _lm_pp_sheet_index(sheet)
    if want < 0:
        return False
    # API без toggle.
    try:
        udr = _lm_pp_doc_unnamed_database_ranges(doc)
        if udr is not None:
            from com.sun.star.table import CellRangeAddress

            addr = CellRangeAddress()
            addr.Sheet = int(want)
            addr.StartColumn = int(sc)
            addr.StartRow = int(sr)
            addr.EndColumn = int(ec)
            addr.EndRow = int(er)
            udr.setByTable(addr)
            db = udr.getByTable(int(want))
            try:
                db.AutoFilter = True
            except Exception:
                db.setPropertyValue("AutoFilter", True)
            if _lm_pp_sheet_has_active_autofilter(doc, sheet):
                return True
    except Exception:
        pass
    # Fallback: dispatch (риск toggle — только если детект выше сказал «нет»).
    try:
        data_range = sheet.getCellRangeByPosition(sc, sr, ec, er)
    except Exception:
        return False

    def _apply():
        if not _lm_pp_select_cell_range(doc, sheet, data_range):
            return False
        return bool(_lm_pp_execute_dispatch(doc, ".uno:DataFilterAutoFilter"))

    return bool(_lm_pp_run_with_suppressed_ui(doc, _apply))


def _lm_pp_remove_all_sheet_database_ranges(doc, sheet):
    """
    Удалить все DatabaseRanges, привязанные к листу
    (включая __Anonymous_Sheet_DB__*, AF_*, AF_PP_* и смарт-таблицы с TableStyle).

    Перед удалением снимает AutoFilter и TableStyleName.
    Возвращает список удалённых имён.
    """
    removed = []
    if doc is None or sheet is None:
        return removed
    keep_smart = False
    try:
        keep_smart = str(getattr(sheet, "Name", "") or "").strip() in _lm_pp_smart_table_keep_sheets()
    except Exception:
        keep_smart = False
    want = _lm_pp_sheet_index(sheet)
    if want < 0:
        return removed
    try:
        dbr = doc.DatabaseRanges
        n = int(dbr.getCount())
    except Exception:
        return removed
    names = []
    i = 0
    while i < n:
        try:
            db = dbr.getByIndex(i)
            addr = db.getDataArea()
            if int(addr.Sheet) == want:
                names.append(str(db.Name))
        except Exception:
            pass
        i = i + 1
    for name in names:
        try:
            db = dbr.getByName(name)
        except Exception:
            continue
        if keep_smart:
            try:
                style_name = str(getattr(db, "TableStyleName", "") or "")
            except Exception:
                try:
                    style_name = str(db.getPropertyValue("TableStyleName") or "")
                except Exception:
                    style_name = ""
            db_name_cf = str(name or "").strip().casefold()
            if style_name.strip() != "" or db_name_cf.startswith("af_pp_"):
                continue
        try:
            db.AutoFilter = False
        except Exception:
            try:
                db.setPropertyValue("AutoFilter", False)
            except Exception:
                pass
        try:
            db.TableStyleName = ""
        except Exception:
            try:
                db.setPropertyValue("TableStyleName", "")
            except Exception:
                pass
        try:
            dbr.removeByName(name)
            removed.append(name)
        except Exception:
            pass
    return removed


def _lm_pp_remove_sheet_autofilters(doc, sheet):
    """
    Снять все именованные автофильтры листа: AF_* (встроенный) и AF_PP_* (постобработка).
    Иначе Calc показывает стрелки и на шапке, и на первой строке данных.
    Меню-фильтр (__Anonymous_Sheet_DB__*) не трогает.
    """
    if doc is None or sheet is None:
        return
    safe = _lm_pp_sheet_autofilter_safe_name(sheet)
    if not safe:
        return
    for db_name in ("AF_" + safe, "AF_PP_" + safe):
        try:
            if db_name in doc.DatabaseRanges:
                doc.DatabaseRanges.removeByName(db_name)
        except Exception:
            pass


def lm_pp_apply_db_range_table_style( db, style_name="TableStyleLight1", row_stripes=True ):
    """
    Дизайн динамического диапазона (AO/LO 26+): TableStyleName + полосы строк.

    Первый стиль галереи «Светлые» — TableStyleLight1 (тёмная шапка, чередование строк).
    На старых сборках без свойства — тихо no-op.
    """
    if db is None:
        return False
    name = str(style_name or "").strip()
    if name == "":
        return False
    applied = False

    def _set(prop, value):
        try:
            setattr(db, prop, value)
            return True
        except Exception:
            pass
        try:
            db.setPropertyValue(prop, value)
            return True
        except Exception:
            return False

    if _set("TableStyleName", name):
        applied = True
    _set("UseRowStripes", bool(row_stripes))
    _set("ContainsHeader", True)
    return applied


# Атрибуты заголовка, которые сохраняем при сбросе формата под smart-стиль.
# Заливку и цвет шрифта НЕ сохраняем — их должен дать TableStyle.
_LM_PP_SMART_HEADER_KEEP_PROPS = (
    "CharFontName",
    "CharHeight",
    "CharWeight",
    "CharPosture",
    "CharUnderline",
    "HoriJustify",
    "VertJustify",
    "VertJustify2",
    "ParaIndent",
    "IsTextWrapped",
)

# С первой строки данных: формат числа/даты + выравнивание столбца.
# Восстанавливаем одним вызовом на весь столбец данных (не по ячейкам).
_LM_PP_SMART_DATA_COL_KEEP_PROPS = (
    "NumberFormat",
    "HoriJustify",
    "VertJustify",
    "VertJustify2",
)

# Свойства, которые сбрасываем на дефолт листа (через donor-ячейку / setPropertyToDefault).
_LM_PP_SMART_RESET_PROPS = (
    "CellStyle",
    "IsCellBackgroundTransparent",
    "CellBackColor",
    "CharColor",
    "CharColor2",
    "CharAutoColor",
    "CharFontName",
    "CharHeight",
    "CharWeight",
    "CharPosture",
    "CharUnderline",
    "HoriJustify",
    "VertJustify",
    "VertJustify2",
    "ParaIndent",
    "IsTextWrapped",
    "NumberFormat",
    "TopBorder",
    "BottomBorder",
    "LeftBorder",
    "RightBorder",
)


def _lm_pp_get_cell_prop(cell, prop_name):
    try:
        return getattr(cell, prop_name)
    except Exception:
        pass
    try:
        return cell.getPropertyValue(prop_name)
    except Exception:
        return None


def _lm_pp_set_cell_prop(cell, prop_name, value):
    if value is None:
        return False
    try:
        setattr(cell, prop_name, value)
        return True
    except Exception:
        pass
    try:
        cell.setPropertyValue(prop_name, value)
        return True
    except Exception:
        return False


def _lm_pp_snapshot_header_keep_attrs(sheet, sc, header_row, ec):
    """Снимок шрифта/выравнивания ячеек строки заголовка (без заливки и цвета)."""
    snaps = []
    if sheet is None:
        return snaps
    try:
        sc = int(sc)
        ec = int(ec)
        header_row = int(header_row)
    except (TypeError, ValueError):
        return snaps
    c = sc
    while c <= ec:
        cell = None
        try:
            cell = sheet.getCellByPosition(c, header_row)
        except Exception:
            cell = None
        snap = {}
        if cell is not None:
            pi = 0
            while pi < len(_LM_PP_SMART_HEADER_KEEP_PROPS):
                prop = _LM_PP_SMART_HEADER_KEEP_PROPS[pi]
                snap[prop] = _lm_pp_get_cell_prop(cell, prop)
                pi = pi + 1
        snaps.append(snap)
        c = c + 1
    return snaps


def _lm_pp_restore_header_keep_attrs(sheet, sc, header_row, ec, snaps):
    """Вернуть сохранённые шрифт/выравнивание заголовка (без заливки и цвета)."""
    if sheet is None or not snaps:
        return
    try:
        sc = int(sc)
        ec = int(ec)
        header_row = int(header_row)
    except (TypeError, ValueError):
        return
    c = sc
    i = 0
    while c <= ec and i < len(snaps):
        cell = None
        try:
            cell = sheet.getCellByPosition(c, header_row)
        except Exception:
            cell = None
        snap = snaps[i] or {}
        if cell is not None:
            pi = 0
            while pi < len(_LM_PP_SMART_HEADER_KEEP_PROPS):
                prop = _LM_PP_SMART_HEADER_KEEP_PROPS[pi]
                if prop in snap:
                    _lm_pp_set_cell_prop(cell, prop, snap.get(prop))
                pi = pi + 1
        c = c + 1
        i = i + 1


def _lm_pp_snapshot_data_col_keep_attrs(sheet, sc, data_row, ec):
    """
    Снимок NumberFormat/выравнивания по первой строке данных.
    Один dict на столбец (для последующего применения на весь столбец данных).
    """
    snaps = []
    if sheet is None:
        return snaps
    try:
        sc = int(sc)
        ec = int(ec)
        data_row = int(data_row)
    except (TypeError, ValueError):
        return snaps
    c = sc
    while c <= ec:
        cell = None
        try:
            cell = sheet.getCellByPosition(c, data_row)
        except Exception:
            cell = None
        snap = {}
        if cell is not None:
            pi = 0
            while pi < len(_LM_PP_SMART_DATA_COL_KEEP_PROPS):
                prop = _LM_PP_SMART_DATA_COL_KEEP_PROPS[pi]
                snap[prop] = _lm_pp_get_cell_prop(cell, prop)
                pi = pi + 1
        snaps.append(snap)
        c = c + 1
    return snaps


def _lm_pp_restore_data_col_keep_attrs(sheet, sc, data_top, ec, data_bottom, snaps):
    """
    Вернуть NumberFormat/выравнивание на столбцы данных одним UNO-вызовом на столбец.
    """
    if sheet is None or not snaps:
        return
    try:
        sc = int(sc)
        ec = int(ec)
        data_top = int(data_top)
        data_bottom = int(data_bottom)
    except (TypeError, ValueError):
        return
    if data_bottom < data_top:
        return
    c = sc
    i = 0
    while c <= ec and i < len(snaps):
        snap = snaps[i] or {}
        col_rng = None
        try:
            col_rng = sheet.getCellRangeByPosition(c, data_top, c, data_bottom)
        except Exception:
            col_rng = None
        if col_rng is not None:
            pi = 0
            while pi < len(_LM_PP_SMART_DATA_COL_KEEP_PROPS):
                prop = _LM_PP_SMART_DATA_COL_KEEP_PROPS[pi]
                if prop in snap and snap.get(prop) is not None:
                    _lm_pp_set_cell_prop(col_rng, prop, snap.get(prop))
                pi = pi + 1
        c = c + 1
        i = i + 1


def _lm_pp_donor_default_cell(sheet, sc, sr, ec, er):
    """
    Ячейка вне диапазона smart-таблицы — как «чистый» формат нового листа.
    Возвращает (cell, col, row) или (None, -1, -1).
    """
    if sheet is None:
        return None, -1, -1
    candidates = (
        (255, 0),
        (500, 0),
        (1023, 0),
        (0, max(int(er) + 100, 2000)),
        (255, max(int(er) + 100, 2000)),
    )
    i = 0
    while i < len(candidates):
        col, row = candidates[i]
        i = i + 1
        if int(sc) <= col <= int(ec) and int(sr) <= row <= int(er):
            continue
        try:
            return sheet.getCellByPosition(int(col), int(row)), int(col), int(row)
        except Exception:
            continue
    return None, -1, -1


def _lm_pp_force_cell_to_sheet_default(cell):
    """Привести ячейку к формату по умолчанию (стиль Default + сброс свойств)."""
    if cell is None:
        return
    try:
        cell.CellStyle = "Default"
    except Exception:
        try:
            cell.setPropertyValue("CellStyle", "Default")
        except Exception:
            pass
    pi = 0
    while pi < len(_LM_PP_SMART_RESET_PROPS):
        _lm_cell_reset_property_to_default(cell, _LM_PP_SMART_RESET_PROPS[pi])
        pi = pi + 1


def _lm_pp_reset_range_formats_to_sheet_default(doc, sheet, cell_range):
    """
    Сбросить формат всего диапазона к умолчанию листа.

    Берём donor-ячейку вне диапазона, принудительно делаем её Default,
    копируем визуальные свойства на диапазон (как «вставить только форматы»).
    """
    if sheet is None or cell_range is None:
        return False
    try:
        addr = cell_range.getRangeAddress()
        sc = int(addr.StartColumn)
        sr = int(addr.StartRow)
        ec = int(addr.EndColumn)
        er = int(addr.EndRow)
    except Exception:
        return False
    donor, dcol, drow = _lm_pp_donor_default_cell(sheet, sc, sr, ec, er)
    if donor is None or dcol < 0:
        return False
    _lm_pp_force_cell_to_sheet_default(donor)
    try:
        donor_rng = sheet.getCellRangeByPosition(dcol, drow, dcol, drow)
        _lm_pp_copy_cell_range_format(donor_rng, cell_range)
    except Exception:
        pass
    # Добивка: явный Default + setPropertyToDefault на самом диапазоне.
    try:
        cell_range.CellStyle = "Default"
    except Exception:
        try:
            cell_range.setPropertyValue("CellStyle", "Default")
        except Exception:
            pass
    pi = 0
    while pi < len(_LM_PP_SMART_RESET_PROPS):
        prop = _LM_PP_SMART_RESET_PROPS[pi]
        try:
            cell_range.setPropertyToDefault(prop)
        except Exception:
            pass
        pi = pi + 1
    return True


def _lm_pp_prepare_range_for_smart_table_style(doc, sheet, cell_range):
    """
    Подготовка диапазона к TableStyle:
      1) снимок шрифта/выравнивания заголовка;
      2) снимок NumberFormat/выравнивания с первой строки данных (по столбцам);
      3) сброс формата всего диапазона на Default листа;
      4) возврат снимков для восстановления после стиля.
    """
    empty = {"header": [], "cols": [], "sc": 0, "sr": 0, "ec": 0, "er": 0}
    if sheet is None or cell_range is None:
        return empty
    try:
        addr = cell_range.getRangeAddress()
        sc = int(addr.StartColumn)
        sr = int(addr.StartRow)
        ec = int(addr.EndColumn)
        er = int(addr.EndRow)
    except Exception:
        return empty
    header_snaps = _lm_pp_snapshot_header_keep_attrs(sheet, sc, sr, ec)
    col_snaps = []
    data_row = sr + 1
    if data_row <= er:
        col_snaps = _lm_pp_snapshot_data_col_keep_attrs(sheet, sc, data_row, ec)
    _lm_pp_reset_range_formats_to_sheet_default(doc, sheet, cell_range)
    return {
        "header": header_snaps,
        "cols": col_snaps,
        "sc": sc,
        "sr": sr,
        "ec": ec,
        "er": er,
    }


def _lm_pp_finish_range_after_smart_table_style(sheet, cell_range, prep_state):
    """
    После TableStyle:
      - вернуть шрифт/выравнивание заголовка (по ячейкам);
      - вернуть NumberFormat/выравнивание столбцов данных одним UNO на столбец.
    """
    if sheet is None or cell_range is None:
        return
    # Совместимость: раньше передавали просто list снимков заголовка.
    if isinstance(prep_state, list):
        try:
            addr = cell_range.getRangeAddress()
            _lm_pp_restore_header_keep_attrs(
                sheet,
                int(addr.StartColumn),
                int(addr.StartRow),
                int(addr.EndColumn),
                prep_state,
            )
        except Exception:
            pass
        return
    if not isinstance(prep_state, dict):
        return
    header_snaps = prep_state.get("header") or []
    col_snaps = prep_state.get("cols") or []
    try:
        sc = int(prep_state.get("sc", 0))
        sr = int(prep_state.get("sr", 0))
        ec = int(prep_state.get("ec", 0))
        er = int(prep_state.get("er", 0))
    except (TypeError, ValueError):
        try:
            addr = cell_range.getRangeAddress()
            sc = int(addr.StartColumn)
            sr = int(addr.StartRow)
            ec = int(addr.EndColumn)
            er = int(addr.EndRow)
        except Exception:
            return
    if header_snaps:
        _lm_pp_restore_header_keep_attrs(sheet, sc, sr, ec, header_snaps)
    if col_snaps and er > sr:
        _lm_pp_restore_data_col_keep_attrs(sheet, sc, sr + 1, ec, er, col_snaps)


def _lm_pp_db_range_to_cell_range(doc, sheet, db):
    """XCellRange по DataArea DatabaseRange (на указанном листе)."""
    if doc is None or sheet is None or db is None:
        return None
    try:
        addr = db.getDataArea()
        return sheet.getCellRangeByPosition(
            int(addr.StartColumn),
            int(addr.StartRow),
            int(addr.EndColumn),
            int(addr.EndRow),
        )
    except Exception:
        return None


def _lm_pp_add_autofilter_range(doc, sheet, data_range):
    """Именованный DatabaseRange + AutoFilter на диапазон (один на лист)."""
    return _lm_pp_add_autofilter_range_with_style(doc, sheet, data_range)


def _lm_pp_add_autofilter_range_with_style(doc, sheet, data_range, style_name=None, row_stripes=True):
    """Именованный AF_PP_ + стиль (экспериментальный smart-диапазон)."""
    if doc is None or sheet is None or data_range is None:
        return False
    prep_state = {"header": [], "cols": [], "sc": 0, "sr": 0, "ec": 0, "er": 0}
    try:
        addr = data_range.getRangeAddress()
        safe = _lm_pp_sheet_autofilter_safe_name(sheet)
        if not safe:
            return False
        db_name = "AF_PP_" + safe
        _lm_pp_remove_sheet_autofilters(doc, sheet)
        # Сброс формата диапазона на Default + сохранить заголовок и форматы столбцов.
        prep_state = _lm_pp_prepare_range_for_smart_table_style(doc, sheet, data_range)
        doc.DatabaseRanges.addNewByName(db_name, addr)
        db = doc.DatabaseRanges.getByName(db_name)
        try:
            db.AutoFilter = True
        except Exception:
            db.setPropertyValue("AutoFilter", True)
        eff_style_name = str(style_name or "").strip()
        eff_row_stripes = bool(row_stripes)
        if eff_style_name == "":
            eff_style_name = "TableStyleLight1"
            try:
                import libre_macros_collect_cfg as _cw_cfg

                eff_style_name = str(
                    getattr(_cw_cfg, "MERGE_RESULT_TABLE_STYLE", eff_style_name) or eff_style_name
                )
                eff_row_stripes = bool(
                    getattr(_cw_cfg, "MERGE_RESULT_TABLE_USE_ROW_STRIPES", True)
                )
            except Exception:
                pass
        lm_pp_apply_db_range_table_style(db, eff_style_name, eff_row_stripes)
        # AO/LO иногда визуально применяет smart-диапазон только после пересчёта/перерисовки.
        try:
            doc.calculateAll()
        except Exception:
            pass
        try:
            db2 = doc.DatabaseRanges.getByName(db_name)
            lm_pp_apply_db_range_table_style(db2, eff_style_name, eff_row_stripes)
        except Exception:
            pass
        _lm_pp_finish_range_after_smart_table_style(sheet, data_range, prep_state)
        try:
            ctrl = doc.getCurrentController()
            if ctrl is not None:
                try:
                    ctrl.setActiveSheet(sheet)
                except Exception:
                    pass
                try:
                    ctrl.select(data_range)
                except Exception:
                    pass
        except Exception:
            pass
        return True
    except Exception:
        try:
            _lm_pp_finish_range_after_smart_table_style(sheet, data_range, prep_state)
        except Exception:
            pass
        return False




def _lm_pp_group_consecutive_columns(cols):
    """Список 0-based столбцов → [(start, end), …] для смежных групп."""
    if cols is None or len(cols) == 0:
        return []
    sorted_cols = sorted(int(c) for c in cols)
    groups = []
    start = sorted_cols[0]
    prev = sorted_cols[0]
    i = 1
    while i < len(sorted_cols):
        if sorted_cols[i] != prev + 1:
            groups.append((start, prev))
            start = sorted_cols[i]
        prev = sorted_cols[i]
        i = i + 1
    groups.append((start, prev))
    return groups


def _lm_pp_copy_range_as_values(doc, sheet, cell_range):
    """Copy → InsertContents (SVD) в тот же диапазон — только значения."""
    sheet_name = sheet.Name if sheet is not None else ""
    return _lm_pp_clipboard_materialize_range_as_values(
        doc,
        sheet,
        cell_range,
        sheet_name,
        dbg_channel="copy_as_values",
    )


def _lm_pp_cell_is_formula(cell):
    """Ячейка содержит формулу (строго CellContentType.FORMULA)."""
    return _lm_pp_cell_is_formula_strict(cell)


def _lm_pp_materialize_cell_as_value(cell):
    """Заменить формулу в ячейке вычисленным значением (без буфера обмена)."""
    if cell is None:
        return False
    if not _lm_pp_cell_is_formula(cell):
        return True
    # Не «замораживать» #ИМЯ?/#VALUE! как будто это нормальный результат.
    if _lm_pp_cell_has_formula_error(cell):
        return False
    try:
        display = str(cell.String) if cell.String is not None else ""
    except Exception:
        display = ""
    try:
        val = cell.Value
    except Exception:
        val = None
    try:
        cell.setFormula("")
    except Exception:
        try:
            cell.Formula = ""
        except Exception:
            return False
    try:
        if display.strip() != "" and not display.strip().startswith("#"):
            cell.String = display
            return True
    except Exception:
        pass
    try:
        if val is not None:
            cell.Value = val
            return True
    except Exception:
        pass
    return False


def _lm_pp_sort_use_native():
    """True — native (range.sort); False — getDataArray/setDataArray."""
    return str(_LM_PP_SORT_MODE or "native").strip().casefold() != "data_array"


def _lm_pp_materialize_range_via_data_array(doc, sheet, cell_range, sheet_name="", label=""):
    """
    Заменить формулы в диапазоне вычисленными значениями за одну операцию.

    getDataArray (с пересчётом) → setDataArray — как в legacy-сортировке.
    """
    if cell_range is None:
        return False
    _lm_pp_calculate_with_auto(doc, sheet)
    try:
        data = cell_range.getDataArray()
        cell_range.setDataArray(data)
        return True
    except Exception as err:
        _lm_pp_apply_formula_dbg(
            doc, sheet_name, "%s: getDataArray/setDataArray FAIL: %s" % (label, err)
        )
        return False


def _lm_pp_materialize_column_as_values(doc, sheet, col, sr, er):
    """Построчно формула → значение (без буфера обмена)."""
    sheet_name = sheet.Name if sheet is not None else ""
    ok_n = 0
    fail_n = 0
    r = int(sr)
    end_r = int(er)
    while r <= end_r:
        try:
            cell = sheet.getCellByPosition(int(col), r)
            if _lm_pp_materialize_cell_as_value(cell):
                ok_n += 1
            else:
                fail_n += 1
        except Exception:
            fail_n += 1
        r += 1
    _lm_pp_apply_formula_dbg(
        doc,
        sheet_name,
        "materialize_column col=%s rows %s..%s: ok=%d fail=%d"
        % (col, sr, er, ok_n, fail_n),
    )
    return fail_n == 0


def _lm_pp_cell_range_fill_auto_down(sheet, col, top_row, bottom_row):
    """
    Протянуть формулу вниз через XCellSeries.fillAuto(TO_BOTTOM, 1).

    Диапазон включает seed-ячейку (top_row) и все строки до bottom_row.
    """
    if sheet is None or bottom_row <= top_row:
        return False
    col = int(col)
    top_row = int(top_row)
    bottom_row = int(bottom_row)
    fill_range = sheet.getCellRangeByPosition(col, top_row, col, bottom_row)
    fill_dir = 0
    try:
        from com.sun.star.sheet.FillDirection import TO_BOTTOM

        fill_dir = TO_BOTTOM
    except Exception:
        pass
    for direction, count in ((fill_dir, 1), (0, 1)):
        try:
            fill_range.fillAuto(direction, count)
            return True
        except Exception:
            pass
    try:
        from com.sun.star.sheet import XCellSeries

        series = fill_range.queryInterface(XCellSeries)
        if series is not None:
            series.fillAuto(fill_dir, 1)
            return True
    except Exception:
        pass
    return False


def _lm_pp_apply_formula_put_first_cell( doc, sheet, formula_text, target_col, sr, sheet_name="", header_row=0 ):
    """Развернуть шаблон для первой строки данных; записать только cell.Formula."""
    col = int(target_col)
    row = int(sr)
    try:
        hdr = int(header_row)
    except (TypeError, ValueError):
        hdr = 0
    first_formula = _lm_pp_expand_apply_formula_template(
        formula_text, row, col, doc=doc, sheet=sheet, header_row=hdr
    )
    _lm_pp_apply_formula_dbg(
        doc,
        sheet_name,
        "первая ячейка row=%d col=%d → %r"
        % (row + 1, col + 1, first_formula[:120]),
    )
    cell = sheet.getCellByPosition(col, row)
    # allow_formula_error=False: иначе ДЛСТР через Formula принимается с #ИМЯ?,
    # а as_values материализует ошибку вместо 0/1.
    ok = _lm_pp_set_cell_formula(
        doc,
        cell,
        first_formula,
        row,
        col,
        sheet_name,
        preserve_r1c1=True,
        allow_formula_error=False,
    )
    if ok:
        try:
            fa = str(cell.getFormula() or cell.Formula or "")[:100]
            _lm_pp_apply_formula_dbg(
                doc, sheet_name, "Formula=%r type=%s" % (fa, cell.getType())
            )
        except Exception:
            pass
    else:
        _lm_pp_apply_formula_dbg(doc, sheet_name, "Formula= FAIL")
    return ok, first_formula


def _lm_pp_formula_uses_source_variables(formula_text):
    return _APPLY_FORMULA_SOURCE_VAR_RE.search(str(formula_text or "")) is not None


def _lm_pp_column_has_formula_errors(sheet, col, sr, er):
    """True, если в столбце есть ячейки-формулы с ошибкой (#ИМЯ? и т.п.)."""
    r = int(sr)
    end_r = int(er)
    col_i = int(col)
    while r <= end_r:
        try:
            cell = sheet.getCellByPosition(col_i, r)
            if _lm_pp_cell_formula_written(cell) and _lm_pp_cell_has_formula_error(cell):
                return True
        except Exception:
            pass
        r += 1
    return False


def _lm_pp_apply_formula_put_formulas_for_as_values( doc, sheet, formula_text, target_col, sr, er, sheet_name="", header_row=0 ):
    """
    as_values=True: записать формулы, пересчитать, материализовать в значения.

    Порядок материализации: getDataArray/setDataArray → clipboard SVD → по ячейкам.
    """
    if _lm_pp_formula_uses_source_variables(formula_text):
        if not _lm_pp_write_formula_column_per_row(
            doc,
            sheet,
            formula_text,
            target_col,
            sr,
            er,
            sheet_name=sheet_name,
            header_row=header_row,
        ):
            return False
    else:
        if not _lm_pp_write_formula_column(
            doc,
            sheet,
            formula_text,
            target_col,
            sr,
            er,
            sheet_name,
            header_row=header_row,
        ):
            return False

    col = int(target_col)
    top = int(sr)
    bottom = int(er)
    try:
        col_range = sheet.getCellRangeByPosition(col, top, col, bottom)
    except Exception:
        return False

    _lm_pp_calculate_with_auto(doc, sheet)
    if _lm_pp_column_has_formula_errors(sheet, col, top, bottom):
        _lm_pp_apply_formula_dbg(
            doc,
            sheet_name,
            "as_values: после записи есть #ИМЯ?/#ЗНАЧ! — материализацию прерываем",
        )
        return False

    if _lm_pp_materialize_range_via_data_array(
        doc, sheet, col_range, sheet_name, label="as_values data_array"
    ):
        try:
            left = _lm_pp_count_formula_cells(sheet, col, top, bottom)
            _lm_pp_apply_formula_dbg(
                doc, sheet_name, "as_values data_array: formula_cells_left=%d" % left
            )
            if left == 0:
                return True
        except Exception:
            return True

    values_ok = _lm_pp_clipboard_materialize_range_as_values(
        doc,
        sheet,
        col_range,
        sheet_name,
        dbg_channel="применить_формулу",
    )
    if values_ok:
        try:
            left = _lm_pp_count_formula_cells(sheet, col, top, bottom)
            _lm_pp_apply_formula_dbg(
                doc, sheet_name, "as_values clipboard: formula_cells_left=%d" % left
            )
            if left == 0:
                return True
        except Exception:
            return True

    if _lm_pp_materialize_column_as_values(doc, sheet, col, top, bottom):
        try:
            left = _lm_pp_count_formula_cells(sheet, col, top, bottom)
            return left == 0
        except Exception:
            return True
    return False


def _lm_pp_apply_formula_rows_have_formulas(sheet, col, sr, er):
    total = int(er) - int(sr) + 1
    if total <= 0:
        return True
    return _lm_pp_count_formula_cells(sheet, int(col), int(sr), int(er)) >= total


def _lm_pp_apply_formula_paste_down_once( doc, sheet, target_col, sr, er, sheet_name="", formula_text=None ):
    """Протянуть формулу из первой ячейки вниз (Copy→InsertContents)."""
    return _lm_pp_fill_formula_down(
        doc,
        sheet,
        target_col,
        sr,
        er,
        sheet_name,
        dbg_channel="применить_формулу",
    )


def _lm_pp_write_formula_column( doc, sheet, formula_text, target_col, sr, er, sheet_name="", header_row=0 ):
    """
    Записать формулу в столбец: шаблон в первой строке данных + Copy→InsertContents вниз.
    """
    ok, _first = _lm_pp_apply_formula_put_first_cell(
        doc,
        sheet,
        formula_text,
        target_col,
        sr,
        sheet_name,
        header_row=header_row,
    )
    if not ok:
        return False
    if int(er) <= int(sr):
        return True
    return _lm_pp_fill_formula_down(
        doc,
        sheet,
        target_col,
        sr,
        er,
        sheet_name,
        dbg_channel="применить_формулу",
    )


def _lm_pp_write_formula_column_per_row(doc, sheet, formula_text, target_col, sr, er, sheet_name="", header_row=0):
    """
    Записать формулу в столбец построчно: шаблон разворачивается отдельно для каждой строки.

    Нужен для <<Переменные...>>, где значения зависят от текущей строки
    (#Путь, [Источник_Лист], [Источник_Файл] и т.п.).
    """
    col = int(target_col)
    top = int(sr)
    bottom = int(er)
    hdr = int(header_row)
    r = top
    written = 0
    while r <= bottom:
        expanded = _lm_pp_expand_apply_formula_template(
            formula_text, r, col, doc=doc, sheet=sheet, header_row=hdr
        )
        cell = sheet.getCellByPosition(col, r)
        ok = _lm_pp_set_cell_formula(
            doc,
            cell,
            expanded,
            r,
            col,
            sheet_name,
            preserve_r1c1=True,
            allow_formula_error=False,
        )
        if not ok:
            _lm_pp_apply_formula_dbg(
                doc,
                sheet_name,
                "postrow FAIL row=%d col=%d formula=%r" % (r + 1, col + 1, expanded[:120]),
            )
            return False
        written += 1
        if (written <= 3) or (r == bottom):
            _lm_pp_apply_formula_dbg(
                doc,
                sheet_name,
                "postrow row=%d col=%d -> %r" % (r + 1, col + 1, expanded[:120]),
            )
        if written % 50 == 0:
            _lm_ui_yield(counter=written)
        r += 1
    _lm_ui_yield(force=True)
    return True


def _lm_pp_count_formula_cells(sheet, col, sr, er):
    n = 0
    r = int(sr)
    end_r = int(er)
    while r <= end_r:
        try:
            # Для «применить_формулу» важно не CellContentType.FORMULA,
            # а наличие записанного текста формулы. В некоторых хостах тип
            # обновляется с задержкой после copyRange/fillAuto.
            if _lm_pp_cell_formula_written(sheet.getCellByPosition(int(col), r)):
                n += 1
        except Exception:
            pass
        r += 1
    return n


_LM_PP_AF_FILL_SERIES_END = 1.0e99


def _lm_pp_af_fill_series(cell_range, direction):
    """
    fillSeries (SIMPLE) для «применить_формулу» и зависимых шагов.

    direction:
        0 — вниз (TO_BOTTOM)
        1 — вправо (TO_RIGHT)
    """
    if cell_range is None:
        return False
    dir_n = int(direction)
    try:
        cell_range.fillSeries(dir_n, 0, 0, 1.0, _LM_PP_AF_FILL_SERIES_END)
        return True
    except Exception:
        pass
    try:
        from com.sun.star.sheet import XCellSeries

        series = cell_range.queryInterface(XCellSeries)
        if series is not None:
            series.fillSeries(dir_n, 0, 0, 1.0, _LM_PP_AF_FILL_SERIES_END)
            return True
    except Exception:
        pass
    return False


def _lm_pp_af_fill_series_down(sheet, col, top_row, bottom_row):
    """Протяжка вниз через fillSeries по диапазону с seed-ячейкой."""
    if sheet is None or int(bottom_row) <= int(top_row):
        return True
    col = int(col)
    top_row = int(top_row)
    bottom_row = int(bottom_row)
    fill_range = sheet.getCellRangeByPosition(col, top_row, col, bottom_row)
    return _lm_pp_af_fill_series(fill_range, 0)


def _lm_pp_fill_formula_down( doc, sheet, col, top_row, bottom_row, sheet_name="", dbg_channel="formula_fill", uno_cleanup=True):
    """
    Протянуть формулу из top_row вниз:
    - основной путь: fillSeries (SIMPLE) как в рабочем макросе;
    - fallback: Copy → InsertContents (только формулы).
    """
    if sheet is None or int(bottom_row) <= int(top_row):
        return True
    try:
        col_i = int(col)
        tr = int(top_row)
        br = int(bottom_row)
    except (TypeError, ValueError):
        return False

    if _lm_pp_apply_formula_rows_have_formulas(sheet, col_i, tr, br):
        _lm_pp_formula_fill_dbg(
            doc,
            sheet_name,
            "fill_down: уже заполнено col=%d rows %d..%d"
            % (col_i + 1, tr + 1, br + 1),
            channel=dbg_channel,
        )
        return True

    if _lm_pp_af_fill_series_down(sheet, col_i, tr, br):
        need = br - tr
        got = _lm_pp_count_formula_cells(sheet, col_i, tr + 1, br)
        if got >= need:
            _lm_pp_formula_fill_dbg(
                doc,
                sheet_name,
                "fillSeries DOWN col=%d rows %d..%d formulas=%d/%d OK"
                % (col_i + 1, tr + 2, br + 1, got, need),
                channel=dbg_channel,
            )
            return True
        _lm_pp_formula_fill_dbg(
            doc,
            sheet_name,
            "fillSeries DOWN partial col=%d rows %d..%d formulas=%d/%d → clipboard"
            % (col_i + 1, tr + 2, br + 1, got, need),
            channel=dbg_channel,
        )
    else:
        _lm_pp_formula_fill_dbg(
            doc,
            sheet_name,
            "fillSeries DOWN FAIL col=%d rows %d..%d → clipboard"
            % (col_i + 1, tr + 2, br + 1),
            channel=dbg_channel,
        )

    return _lm_pp_clipboard_fill_formulas_down(
        doc,
        sheet,
        col_i,
        tr,
        br,
        sheet_name,
        dbg_channel=dbg_channel,
        uno_cleanup=uno_cleanup,
    )


def _lm_pp_fill_down_calculate_normalize_template(template):
    """
    Нормализация шаблона только для «заполнение_вниз_вычислить».
    R{-1}C → R[-1]C (фигурные скобки в R1C1-смещениях).
    """
    s = str(template or "").strip()
    if s == "":
        return s
    s = re.sub(r"R\{([+-]?\d+)\}C", r"R[\1]C", s, flags=re.IGNORECASE)
    s = re.sub(
        r"R(\[([+-]?\d+)\]|(\d+))?C\{([+-]?\d+)\}",
        lambda m: "R%sC[%s]"
        % (
            m.group(1) or "",
            m.group(4),
        ),
        s,
        flags=re.IGNORECASE,
    )
    return s


def _lm_pp_fill_down_calculate_cell_is_empty(sheet, col, row):
    try:
        return cell_text(sheet.getCellByPosition(int(col), int(row))) == ""
    except Exception:
        return True


_LM_FDC_FILL_SERIES_END = 1.0e99


def _lm_pp_fdc_fill_series(cell_range, direction):
    """
    fillSeries (SIMPLE) — как в рабочем макросе formula_fill.
    direction: 0 = вниз (TO_BOTTOM), 1 = вправо (TO_RIGHT).
    """
    if cell_range is None:
        return False
    dir_n = int(direction)
    try:
        cell_range.fillSeries(dir_n, 0, 0, 1.0, _LM_FDC_FILL_SERIES_END)
        return True
    except Exception:
        pass
    try:
        from com.sun.star.sheet import XCellSeries

        series = cell_range.queryInterface(XCellSeries)
        if series is not None:
            series.fillSeries(dir_n, 0, 0, 1.0, _LM_FDC_FILL_SERIES_END)
            return True
    except Exception:
        pass
    return False


def _lm_pp_fdc_fill_series_down(sheet, col, top_row, bottom_row):
    if sheet is None or int(bottom_row) <= int(top_row):
        return True
    col = int(col)
    top_row = int(top_row)
    bottom_row = int(bottom_row)
    fill_range = sheet.getCellRangeByPosition(col, top_row, col, bottom_row)
    return _lm_pp_fdc_fill_series(fill_range, 0)


def _lm_pp_fdc_fill_series_right(sheet, left_col, row, right_col):
    if sheet is None or int(right_col) <= int(left_col):
        return True
    left_col = int(left_col)
    row = int(row)
    right_col = int(right_col)
    fill_range = sheet.getCellRangeByPosition(left_col, row, right_col, row)
    return _lm_pp_fdc_fill_series(fill_range, 1)


def _lm_pp_fdc_fill_series_right_block(sheet, left_col, top_row, right_col, bottom_row):
    """fillSeries вправо по прямоугольнику (formula_fill шаг 2)."""
    if sheet is None:
        return False
    left_col = int(left_col)
    top_row = int(top_row)
    right_col = int(right_col)
    bottom_row = int(bottom_row)
    if right_col <= left_col or bottom_row < top_row:
        return True
    fill_range = sheet.getCellRangeByPosition(
        left_col, top_row, right_col, bottom_row
    )
    return _lm_pp_fdc_fill_series(fill_range, 1)


def _lm_pp_fdc_count_formula_cells_row(sheet, left_col, right_col, row):
    """Сколько ячеек с формулами в строке на отрезке [left_col..right_col]."""
    if sheet is None:
        return 0
    n = 0
    c = int(left_col)
    end_c = int(right_col)
    row = int(row)
    while c <= end_c:
        try:
            if _lm_pp_cell_formula_written(sheet.getCellByPosition(c, row)):
                n += 1
        except Exception:
            pass
        c += 1
    return n


def _lm_pp_fdc_count_formula_cells_rect(sheet, left_col, top_row, right_col, bottom_row):
    """Сколько ячеек с формулами в прямоугольнике (перебор по ячейкам)."""
    if sheet is None:
        return 0
    left_col = int(left_col)
    right_col = int(right_col)
    top_row = int(top_row)
    bottom_row = int(bottom_row)
    if right_col < left_col or bottom_row < top_row:
        return 0
    total = 0
    r = top_row
    while r <= bottom_row:
        total += _lm_pp_fdc_count_formula_cells_row(sheet, left_col, right_col, r)
        r += 1
    return total


def _lm_pp_fdc_right_empty_end_col(sheet, row, start_col, end_col):
    """Последний подряд пустой столбец вправо (0-based), не выше end_col."""
    c = int(start_col)
    limit = int(end_col)
    row = int(row)
    while c <= limit:
        if not _lm_pp_fill_down_calculate_cell_is_empty(sheet, c, row):
            break
        c += 1
    return c - 1


def _lm_pp_fdc_expand_run_to_right( doc, sheet, col, run_start, run_end, end_col, is_formula, template, sheet_name=""):
    """После вертикального пробега — fillSeries вправо по пустым (в пределах ec)."""
    filled = 0
    errors = 0
    col = int(col)
    end_col = int(end_col)
    template_norm = _lm_pp_fill_down_calculate_normalize_template(template)
    dbg = "заполнение_вниз_вычислить"

    if is_formula:
        block_right = col
        r = int(run_start)
        while r <= int(run_end):
            re = _lm_pp_fdc_right_empty_end_col(sheet, r, col + 1, end_col)
            if re > col:
                if block_right == col or re < block_right:
                    block_right = re
            r += 1
        if block_right > col:
            if _lm_pp_fdc_fill_series_right_block(
                sheet, col, run_start, block_right, run_end
            ):
                need = (block_right - col) * (run_end - run_start + 1)
                got = _lm_pp_fdc_count_formula_cells_rect(
                    sheet, col + 1, run_start, block_right, run_end
                )
                if got >= need:
                    _lm_pp_formula_fill_dbg(
                        doc,
                        sheet_name,
                        "fillSeries RIGHT block col %d..%d rows %d..%d (%d/%d)"
                        % (
                            col + 2,
                            block_right + 1,
                            run_start + 1,
                            run_end + 1,
                            got,
                            need,
                        ),
                        channel=dbg,
                    )
                    return (block_right - col) * (run_end - run_start + 1), 0
            _lm_pp_formula_fill_dbg(
                doc,
                sheet_name,
                "fillSeries RIGHT block FAIL/partial, по строкам",
                channel=dbg,
            )

    r = int(run_start)
    while r <= int(run_end):
        right_end = _lm_pp_fdc_right_empty_end_col(sheet, r, col + 1, end_col)
        if right_end <= col:
            r += 1
            continue
        span = right_end - col
        if is_formula:
            try:
                seed = sheet.getCellByPosition(col, r)
            except Exception:
                errors += span
                r += 1
                continue
            if not _lm_pp_cell_formula_written(seed):
                errors += span
            elif _lm_pp_fdc_fill_series_right(sheet, col, r, right_end):
                got = _lm_pp_fdc_count_formula_cells_row(sheet, col + 1, right_end, r)
                if got >= span:
                    filled += span
                else:
                    c = col + 1
                    while c <= right_end:
                        try:
                            if _lm_pp_cell_formula_written(
                                sheet.getCellByPosition(c, r)
                            ):
                                filled += 1
                                c += 1
                                continue
                        except Exception:
                            pass
                        if _lm_pp_set_cell_formula(
                            doc,
                            sheet.getCellByPosition(c, r),
                            template_norm,
                            r,
                            c,
                            sheet_name,
                            preserve_r1c1=False,
                            allow_formula_error=True,
                        ):
                            filled += 1
                        else:
                            errors += 1
                        c += 1
            else:
                c = col + 1
                while c <= right_end:
                    if _lm_pp_set_cell_formula(
                        doc,
                        sheet.getCellByPosition(c, r),
                        template_norm,
                        r,
                        c,
                        sheet_name,
                        preserve_r1c1=False,
                        allow_formula_error=True,
                    ):
                        filled += 1
                    else:
                        errors += 1
                    c += 1
        else:
            c = col + 1
            while c <= right_end:
                try:
                    sheet.getCellByPosition(c, r).String = template_norm
                    filled += 1
                except Exception:
                    errors += 1
                c += 1
        r += 1
    return filled, errors


def _lm_pp_fill_down_calculate_write_run( doc, sheet, col, run_start, run_end, formula_template, sheet_name="", expand_to_right=False, end_col=None):
    """
    Заполнить один непрерывный пробег пустых ячеек: формула в первой + fillSeries вниз.
    expand_to_right — fillSeries вправо (как formula_fill).
    """
    col = int(col)
    run_start = int(run_start)
    run_end = int(run_end)
    template = _lm_pp_fill_down_calculate_normalize_template(formula_template)
    dbg = "заполнение_вниз_вычислить"
    is_formula = template.startswith("=")

    def _finish(filled, errors):
        if expand_to_right and end_col is not None:
            ef, ee = _lm_pp_fdc_expand_run_to_right(
                doc,
                sheet,
                col,
                run_start,
                run_end,
                end_col,
                is_formula,
                template,
                sheet_name,
            )
            return filled + ef, errors + ee
        return filled, errors

    if not is_formula:
        r = run_start
        ok_n = 0
        err_n = 0
        while r <= run_end:
            try:
                sheet.getCellByPosition(col, r).String = template
                ok_n += 1
            except Exception:
                err_n += 1
            r += 1
        return _finish(ok_n, err_n)

    first_formula = _lm_pp_expand_apply_formula_template(
        template, run_start, col, doc=doc
    )
    _lm_pp_formula_fill_dbg(
        doc,
        sheet_name,
        "пробег rows %d..%d col=%d → %r"
        % (run_start + 1, run_end + 1, col + 1, first_formula[:120]),
        channel=dbg,
    )
    if not _lm_pp_set_cell_formula(
        doc,
        sheet.getCellByPosition(col, run_start),
        template,
        run_start,
        col,
        sheet_name,
        preserve_r1c1=False,
        allow_formula_error=True,
    ):
        _lm_pp_formula_fill_dbg(
            doc,
            sheet_name,
            "первая ячейка FAIL row=%d" % (run_start + 1),
            channel=dbg,
        )
        return 0, run_end - run_start + 1

    if run_start >= run_end:
        return _finish(1, 0)

    filled = 1
    errors = 0
    if _lm_pp_fdc_fill_series_down(sheet, col, run_start, run_end):
        need = run_end - run_start
        got = _lm_pp_count_formula_cells(sheet, col, run_start + 1, run_end)
        if got >= need:
            _lm_pp_formula_fill_dbg(
                doc,
                sheet_name,
                "fillSeries DOWN col=%d rows %d..%d (%d/%d)"
                % (col + 1, run_start + 2, run_end + 1, got, need),
                channel=dbg,
            )
            return _finish(run_end - run_start + 1, 0)
        _lm_pp_formula_fill_dbg(
            doc,
            sheet_name,
            "fillSeries DOWN partial col=%d (%d/%d), построчно"
            % (col + 1, got, need),
            channel=dbg,
        )
    else:
        _lm_pp_formula_fill_dbg(
            doc,
            sheet_name,
            "fillSeries DOWN FAIL col=%d rows %d..%d, построчно"
            % (col + 1, run_start + 2, run_end + 1),
            channel=dbg,
        )

    r = run_start + 1
    while r <= run_end:
        try:
            if _lm_pp_cell_formula_written(sheet.getCellByPosition(col, r)):
                filled += 1
                r += 1
                continue
        except Exception:
            pass
        if _lm_pp_set_cell_formula(
            doc,
            sheet.getCellByPosition(col, r),
            template,
            r,
            col,
            sheet_name,
            preserve_r1c1=False,
            allow_formula_error=True,
        ):
            filled += 1
        else:
            errors += 1
        r += 1
    return _finish(filled, errors)


def _lm_pp_fill_down_calculate_empty_runs( doc, sheet, col, first_row, last_row, formula_template, sheet_name="", expand_to_right=False, end_col=None):
    """
    «Заполнение_вниз_вычислить»: только пустые ячейки столбца, пробегами.
    Отдельный путь — без clipboard/dispatch (не затрагивает «применить_формулу»).
    """
    filled = 0
    errors = 0
    col = int(col)
    row = int(first_row)
    end_row = int(last_row)

    while row <= end_row:
        if not _lm_pp_fill_down_calculate_cell_is_empty(sheet, col, row):
            row += 1
            continue

        run_start = row
        while row <= end_row:
            if not _lm_pp_fill_down_calculate_cell_is_empty(sheet, col, row):
                break
            row += 1
        run_end = row - 1

        run_filled, run_errors = _lm_pp_fill_down_calculate_write_run(
            doc,
            sheet,
            col,
            run_start,
            run_end,
            formula_template,
            sheet_name,
            expand_to_right=expand_to_right,
            end_col=end_col,
        )
        filled += run_filled
        errors += run_errors

    return filled, errors


def _lm_pp_autofit_columns(doc, sheet, col_indices, sr=None, er=None):
    """
    Автоподбор ширины столбцов.

    Основной путь: Column.OptimalWidth=True по API (без выделения 0..1048575).
    Опционально (_LM_PP_AUTOFIT_USE_DISPATCH): один dispatch по прямоугольнику
    sr..er и min..max(col_indices), не по каждой колонке отдельно.
    """
    if sheet is None or not col_indices:
        return

    cols = []
    i = 0
    while i < len(col_indices):
        try:
            cols.append(int(col_indices[i]))
        except (TypeError, ValueError):
            pass
        i += 1
    if not cols:
        return

    sc = min(cols)
    ec = max(cols)

    if sr is None or er is None:
        try:
            _ua_sc, u_sr, _ua_ec, u_er = _lm_vlookup_sheet_used_area(sheet)
            if sr is None:
                sr = u_sr
            if er is None:
                er = u_er
        except Exception:
            pass
    try:
        sr = int(sr) if sr is not None else 0
        er = int(er) if er is not None else sr
    except (TypeError, ValueError):
        sr, er = 0, 0
    if er < sr:
        sr, er = 0, max(0, er)

    if doc is not None:
        _lm_pp_calculate_with_auto(doc, sheet)

    try:
        columns = sheet.getColumns()
    except Exception:
        columns = None

    ci = 0
    while ci < len(cols):
        try:
            if columns is not None:
                columns.getByIndex(cols[ci]).OptimalWidth = True
        except Exception:
            pass
        ci += 1

    if not _LM_PP_AUTOFIT_USE_DISPATCH or doc is None or er < sr:
        return

    try:
        block = sheet.getCellRangeByPosition(sc, sr, ec, er)
        if _lm_pp_select_cell_range(doc, sheet, block):
            _lm_pp_execute_dispatch(doc, ".uno:SetOptimalColumnWidth")
    except Exception:
        pass


def _lm_pp_header_titles(sheet, header_row_range):
    """
    Список текстов заголовков по столбцам строки результата.

    Параметры:
        sheet — com.sun.star.sheet.XSpreadsheet
        header_row_range — com.sun.star.table.XCellRange (обычно строка 0)

    Возвращает list[str]. Нужен колбэкам, ищущим столбцы по подстроке в заголовке.
    Не в карте MAP.
    """
    sc, sr, ec, er = _lm_pp_range_address(header_row_range)
    titles = []
    clean_fn = None
    try:
        from libre_macros_header_lib import clean_header_name

        clean_fn = clean_header_name
    except Exception:
        clean_fn = None
    c = sc
    while c <= ec:
        t = cell_text(sheet.getCellByPosition(c, sr))
        if clean_fn is not None:
            try:
                t = clean_fn(t)
            except Exception:
                pass
        titles.append(t)
        c = c + 1
    return titles


def _lm_sheet_delete_columns(sheet, col_index, count=1):
    """
    Удалить столбцы на листе Calc (XTableColumns.removeByIndex).

    Параметры:
        sheet — XSpreadsheet
        col_index — 0-based индекс первого удаляемого столбца
        count — сколько столбцов удалить

    Связь:
        lm_pp_range_delete_columns (у XSpreadsheet нет метода deleteColumns).
    """
    if sheet is None:
        return
    sheet.getColumns().removeByIndex(int(col_index), int(count))


def _lm_sheet_delete_rows(sheet, row_index, count=1):
    """
    Удалить строки на листе Calc (XTableRows.removeByIndex).

    Параметры:
        sheet — XSpreadsheet
        row_index — 0-based индекс первой удаляемой строки
        count — сколько строк удалить

    Связь:
        lm_pp_range_skip_empty_rows.
    """
    if sheet is None:
        return
    sheet.getRows().removeByIndex(int(row_index), int(count))


def _lm_pp_reconstruct_skip_spec_from_extra(extra_args):
    """Собрать текст условия пропуска из rest (колонка C) после parse_pp_extra_args."""
    if not extra_args:
        return ""
    if len(extra_args) == 1:
        return str(extra_args[0]).strip()
    return ",".join(str(x) for x in extra_args).strip()


def _lm_pp_cell_is_true(cell):
    """Истинность ячейки после пересчёта формулы (Calc TRUE / ненулевое число)."""
    if cell is None:
        return False
    try:
        v = cell.Value
        if isinstance(v, (int, float)):
            return v != 0
    except Exception:
        pass
    t = cell_text(cell).strip().upper()
    return t in ("TRUE", "ИСТИНА", "1", "ДА")


# Кириллические токены даты/времени (визард) → коды Calc. Длинные раньше коротких.
_LM_PP_CYRILLIC_DATE_TOKENS = (
    (u"ГГГГ", u"YYYY"),
    (u"ДДДД", u"NNNN"),
    (u"ММММ", u"MMMM"),
    (u"ЧЧ", u"HH"),
    (u"СС", u"SS"),
    (u"ГГ", u"YY"),
    (u"ДДД", u"NNN"),
    (u"МММ", u"MMM"),
    (u"ДД", u"DD"),
    (u"ММ", u"MM"),
    (u"Ч", u"H"),
    (u"Д", u"D"),
    (u"М", u"M"),
    (u"С", u"S"),
    (u"Г", u"Y"),
)


def _lm_pp_translate_cyrillic_date_format(fmt_string):
    """
    ДД.ММ.ГГГГ ЧЧ:ММ:СС → DD.MM.YYYY HH:MM:SS (Calc не понимает кириллические токены).
    Латинские шаблоны и числовые форматы не меняет.
    """
    s = str(fmt_string or "").replace("\xa0", " ")
    if s == "":
        return s
    low = s.lower()
    if not any(ch in low for ch in (u"д", u"м", u"г", u"ч", u"с")):
        return s
    out = []
    i = 0
    n = len(s)
    while i < n:
        matched = False
        for src, dst in _LM_PP_CYRILLIC_DATE_TOKENS:
            sl = len(src)
            if i + sl > n:
                continue
            chunk = s[i : i + sl]
            try:
                same = chunk.casefold() == src.casefold()
            except AttributeError:
                same = chunk.lower() == src.lower()
            if not same:
                continue
            out.append(dst)
            i = i + sl
            matched = True
            break
        if not matched:
            out.append(s[i])
            i = i + 1
    return u"".join(out)


def _lm_pp_normalize_calc_number_format(fmt_string, lang="ru"):
    """
    Excel-шаблон (#,##0.00) → Calc для локали ru/RU (# ##0,00).

    В ru запятая — десятичный разделитель; Excel-тысячные «,» дают
    артефакты вроде 50300,000.00.
    Кириллические токены даты (ДД.ММ.ГГГГ …) → латинские коды Calc.
    """
    s = str(fmt_string or "").strip().replace("\xa0", " ")
    if s == "":
        return s
    s = _lm_pp_translate_cyrillic_date_format(s)
    if str(lang or "ru").lower()[:2] != "ru":
        return s
    if "[" in s or "$" in s or "₽" in s:
        return s
    parts = s.split(";")
    out = []
    pi = 0
    while pi < len(parts):
        p = parts[pi]
        pi = pi + 1
        m = re.search(r"\.(\d+)$", p)
        if m is not None:
            p = p[: m.start()] + "," + m.group(1)
        p = p.replace("#,##", "# ##")
        while "#," in p:
            p = p.replace("#,", "# ")
        out.append(p)
    return ";".join(out)


def _lm_pp_number_format_key(doc, fmt_string, lang="ru", country="RU"):
    """
    Получить или создать ключ NumberFormat в документе для шаблона fmt_string.

    Параметры:
        doc — com.sun.star.sheet.SpreadsheetDocument
        fmt_string — str, шаблон Calc (например _LM_PP_FORMAT_MONEY_FMT)
        lang, country — локаль формата (по умолчанию ru/RU)

    Возвращает int (ключ для cell.NumberFormat). Не в карте MAP.
    """
    fmt_string = _lm_pp_normalize_calc_number_format(fmt_string, lang=lang)
    nf = doc.NumberFormats
    loc = uno.createUnoStruct("com.sun.star.lang.Locale")
    loc.Language = lang
    loc.Country = country
    key = nf.queryKey(fmt_string, loc, False)
    if key == -1:
        cl = _lm_doc_char_locale(doc)
        if cl is not None:
            key = nf.queryKey(fmt_string, cl, False)
    if key == -1:
        key = nf.addNew(fmt_string, loc)
    return key


def _lm_pp_border_line(color=0xB4B4B4, width=10):
    """
    Новая структура BorderLine2 для одной стороны ячейки (не переиспользовать!).

    width — LineWidth в 1/100 мм. LineStyle=0 — сплошная (1 часто «точки», не видно).
    """
    bl = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    bl.Color = int(color)
    bl.LineStyle = 0
    w = int(width)
    bl.LineWidth = w
    bl.InnerLineWidth = w
    bl.OuterLineWidth = w
    return bl


def _lm_pp_set_cell_side_border(cell, side, color, width):
    """Одна сторона ячейки — свой экземпляр BorderLine2."""
    line = _lm_pp_border_line(color, width)
    if side == "top":
        cell.TopBorder = line
    elif side == "bottom":
        cell.BottomBorder = line
    elif side == "left":
        cell.LeftBorder = line
    elif side == "right":
        cell.RightBorder = line


def _lm_pp_log_grid_debug(doc, sheet, grid_label, status, note):
    """Отладочная запись в «Сбор_книг_лог» для постобработки сетки."""
    sheet_name = ""
    if sheet is not None:
        try:
            sheet_name = sheet.Name
        except Exception:
            pass
    _lm_log_postprocess(doc, sheet_name, "сетка", grid_label, status, note)


def _lm_pp_format_range_debug(data_range):
    """Краткое описание диапазона для лога сетки."""
    if data_range is None:
        return "data_range=None"
    try:
        sc, sr, ec, er = _lm_pp_range_address(data_range)
        ncols = ec - sc + 1
        nrows = er - sr + 1
        svc = "?"
        try:
            if data_range.supportsService("com.sun.star.table.CellRange"):
                svc = "CellRange"
        except Exception:
            pass
        return "%d×%d ячеек (%d:%d..%d:%d), %s" % (ncols, nrows, sc, sr, ec, er, svc)
    except Exception as err:
        return "диапазон: %s" % err


def _lm_pp_cell_border_readback(cell):
    """Проверка LineWidth у BottomBorder одной ячейки (для лога)."""
    try:
        bl = cell.BottomBorder
        if bl is None:
            return "BottomBorder=None"
        w = getattr(bl, "LineWidth", None)
        if w is None or int(w) == 0:
            w = getattr(bl, "OuterLineWidth", 0)
        return "LineWidth=%s" % w
    except Exception as err:
        return "readback: %s" % err


def _lm_pp_grid_target_range(sheet, data_range, header_row_range):
    """
    Диапазон для сетки: данные + строка заголовков (если header выше data_range).
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    if header_row_range is not None:
        try:
            h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
            if h_sr < sr:
                sr = h_sr
            if h_sc < sc:
                sc = h_sc
            if h_ec > ec:
                ec = h_ec
        except Exception:
            pass
    return sheet.getCellRangeByPosition(sc, sr, ec, er)


def _lm_pp_apply_cell_border_grid(sheet, data_range, color, width):
    """
    Сетка: применяет границы ко всему диапазону сразу через TableBorder.

    Вместо обработки каждой ячейки отдельно — один вызов на диапазон:
    - Внутренние ячейки: вертикальные/горизонтальные линии через VerticalLine/HorizontalLine
    - Внешние границы: через TopBorder/BottomBorder/LeftBorder/RightBorder

    Возвращает (cells_ok, cells_fail) — при успехе (1, 0), иначе (0, 1).
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    cells_ok = 0
    cells_fail = 0

    try:
        # Получаем диапазон для установки границ
        cell_range = sheet.getCellRangeByPosition(sc, sr, ec, er)
        if cell_range is None:
            return 0, 1

        # Создаём линию для границ
        line = _lm_pp_border_line(color, width)

        # Получаем текущий TableBorder или создаём новый
        try:
            tb = cell_range.TableBorder
        except Exception:
            # Если TableBorder не поддерживается — fallback к построчному методу
            return _lm_pp_apply_cell_border_grid_fallback(sheet, sc, sr, ec, er, color, width)

        # Устанавливаем все границы
        tb.TopBorder = line
        tb.BottomBorder = line
        tb.LeftBorder = line
        tb.RightBorder = line
        tb.HorizontalLine = line
        tb.VerticalLine = line

        # Применяем к диапазону
        cell_range.TableBorder = tb

        # Отдаём управление UI
        _lm_ui_yield(force=True)
        cells_ok = 1
    except Exception:
        cells_fail = 1
        # Fallback к построчному методу при ошибке
        try:
            return _lm_pp_apply_cell_border_grid_fallback(sheet, sc, sr, ec, er, color, width)
        except Exception:
            pass

    return cells_ok, cells_fail


def _lm_pp_apply_cell_border_grid_fallback(sheet, sc, sr, ec, er, color, width):
    """
    Fallback: построчная установка границ если TableBorder не работает.
    Упрощённая версия — только основные границы без перебора каждой ячейки.
    """
    cells_ok = 0
    cells_fail = 0
    line = _lm_pp_border_line(color, width)

    try:
        # Устанавливаем внешние границы всего диапазона одним вызовом
        outer_range = sheet.getCellRangeByPosition(sc, sr, ec, er)
        if outer_range is not None:
            outer_range.TopBorder = line
            outer_range.BottomBorder = line
            outer_range.LeftBorder = line
            outer_range.RightBorder = line
            cells_ok = cells_ok + 1
    except Exception:
        cells_fail = cells_fail + 1

    # Внутренние границы — построчно, но всей строкой сразу
    r = sr
    n = 0
    while r <= er:
        try:
            # Горизонтальная линия под строкой (BottomBorder для всей строки)
            if r < er:  # Не последняя строка
                row_range = sheet.getCellRangeByPosition(sc, r, ec, r)
                if row_range is not None:
                    row_range.BottomBorder = line
        except Exception:
            cells_fail = cells_fail + 1
        r = r + 1
        n = n + 1
        if n % 50 == 0:
            _lm_ui_yield(counter=n)

    # Вертикальные линии — по столбцам
    c = sc
    n = 0
    while c <= ec:
        try:
            if c < ec:  # Не последний столбец
                col_range = sheet.getCellRangeByPosition(c, sr, c, er)
                if col_range is not None:
                    col_range.RightBorder = line
        except Exception:
            cells_fail = cells_fail + 1
        c = c + 1
        n = n + 1
        if n % 50 == 0:
            _lm_ui_yield(counter=n)

    _lm_ui_yield(force=True)
    return cells_ok, cells_fail


def _lm_pp_apply_range_grid_borders( data_range, color=0x666666, width=15, doc=None, sheet=None, grid_label="сетка", header_row_range=None, include_header=False):
    """
    Сетка по ячейкам (TableBorder). По умолчанию только data_range;
    include_header=True — объединить со строкой заголовков.
    """
    if data_range is None:
        _lm_pp_log_grid_debug(doc, sheet, grid_label, "ошибка", "data_range is None")
        return
    if sheet is None:
        try:
            sheet = data_range.getSpreadsheet()
        except Exception:
            sheet = None
    if sheet is None:
        _lm_pp_log_grid_debug(doc, sheet, grid_label, "ошибка", "нет листа (sheet)")
        return

    target = data_range
    if include_header and header_row_range is not None:
        try:
            target = _lm_pp_grid_target_range(sheet, data_range, header_row_range)
        except Exception as err:
            _lm_pp_log_grid_debug(
                doc, sheet, grid_label, "debug", "без заголовка: %s" % err
            )

    _lm_pp_log_grid_debug(
        doc,
        sheet,
        grid_label,
        "debug",
        "старт w=%s color=#%06X: %s"
        % (width, int(color) & 0xFFFFFF, _lm_pp_format_range_debug(target)),
    )

    cells_ok, cells_fail = _lm_pp_apply_cell_border_grid(
        sheet, target, color, width
    )
    sc, sr, ec, er = _lm_pp_range_address(target)
    sample = ""
    try:
        sample = _lm_pp_cell_border_readback(
            sheet.getCellByPosition(sc, sr)
        )
    except Exception:
        pass
    if cells_fail > 0:
        _lm_pp_log_grid_debug(
            doc,
            sheet,
            grid_label,
            "ошибка",
            "ячейки ok=%d fail=%d; %s" % (cells_ok, cells_fail, sample),
        )
    else:
        _lm_pp_log_grid_debug(
            doc,
            sheet,
            grid_label,
            "ok",
            "ячейки %d (строки %d..%d), %s" % (cells_ok, sr, er, sample),
        )


def _lm_pp_apply_table_grid( data_range, color=0x666666, doc=None, sheet=None, grid_label="тонкая_сетка", header_row_range=None):
    """
    Тонкая серая сетка по data_range (границы ячеек).

    Вызывается из lm_pp_range_thin_grid_borders. Не в карте MAP.
    """
    _lm_pp_apply_range_grid_borders(
        data_range,
        color=color,
        width=15,
        doc=doc,
        sheet=sheet,
        grid_label=grid_label,
        header_row_range=header_row_range,
        include_header=True,
    )


def _lm_pp_apply_table_grid_thick( data_range, color=0x333333, width=35, doc=None, sheet=None, grid_label="толстая_сетка", header_row_range=None):
    """
    Толстая серая сетка по data_range (границы ячеек).

    Вызывается из lm_pp_range_thick_grid_borders. Не в карте MAP.
    """
    _lm_pp_apply_range_grid_borders(
        data_range,
        color=color,
        width=width,
        doc=doc,
        sheet=sheet,
        grid_label=grid_label,
        header_row_range=header_row_range,
        include_header=True,
    )


def _lm_pp_pad_row_height(sheet, row_index, add_mm):
    """
    Увеличить высоту строки на add_mm (миллиметры → 1/100 мм в Calc).
    Итог не ниже _LM_FMT_HEADER_ROW_MIN_MM.
    """
    _lm_pp_set_or_pad_row_height_mm(sheet, row_index, add_mm, add_height=True)


def _lm_pp_set_or_pad_row_height_mm(sheet, row_index, height_mm, add_height=False):
    """
    Высота строки: абсолютная (по умолчанию) или плюс к текущей.

    Минимум — _LM_FMT_HEADER_ROW_MIN_MM (защита от схлопывания).
    """
    if sheet is None:
        return
    try:
        mm = float(height_mm)
    except (TypeError, ValueError):
        mm = float(_LM_FMT_HEADER_ROW_PAD_MM)
    min_h = _lm_pp_mm_to_row_height(_LM_FMT_HEADER_ROW_MIN_MM)
    delta_h = _lm_pp_mm_to_row_height(mm)
    row = sheet.getRows().getByIndex(int(row_index))
    try:
        row.IsVisible = True
    except Exception:
        pass
    try:
        row.OptimalHeight = False
    except Exception:
        pass
    if add_height:
        try:
            h = int(row.Height or 0)
        except Exception:
            h = 0
        if h < 1:
            h = 0
        new_h = int(h) + int(delta_h)
        if new_h < min_h:
            new_h = min_h
        row.Height = new_h
    else:
        new_h = int(delta_h)
        if new_h < min_h:
            new_h = min_h
        row.Height = new_h


# =============================================================================
# Постобработка диапазона данных (без строки заголовков)
# Колонка C: только JSON [{"v":1,"fn":"…",…}] — docs/13_JSON_PARAMS.md, визард «Параметры…».
# =============================================================================

def _lm_pp_word_wrap_enabled(extra_args):
    """Устаревший API: только флаг переноса (см. _lm_pp_wrap_settings)."""
    wrap_on, _include = _lm_pp_wrap_settings(extra_args)
    return wrap_on


def _lm_pp_wrap_settings(extra_args):
    """
    JSON колонки C (param_decode): wrap, include_header; пустой блок — перенос вкл., без заголовка.
    """
    if extra_args is None or len(extra_args) == 0:
        return True, False
    if len(extra_args) == 1:
        text = str(extra_args[0]).strip()
    else:
        text = _lm_pp_join_extra_args(extra_args).strip()
    if text == "":
        return True, False
    if _lm_pp_payload_is_json(text):
        from libre_macros_param_codec import param_decode

        blocks = param_decode("перенос", text)
        if blocks:
            b = blocks[0]
            return bool(b.get("wrap", True)), bool(b.get("include_header"))
        return True, False
    tokens = [t.strip() for t in text.split(",") if t.strip()]
    if not tokens:
        return True, False
    try:
        wrap_on = lm_parse_bool_param(tokens[0], default=True)
    except ValueError:
        wrap_on = True
    include = False
    ti = 1
    while ti < len(tokens):
        tok = tokens[ti].strip().lower()
        if tok in ("header", "заголовок", "заг", "hdr"):
            include = True
        ti = ti + 1
    return wrap_on, include


def lm_pp_range_word_wrap(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Включить или выключить перенос текста по словам для блока данных (без заголовка).

    Параметры:
        doc — com.sun.star.sheet.SpreadsheetDocument (не используется)
        sheet — com.sun.star.sheet.XSpreadsheet (не используется)
        data_range — com.sun.star.table.XCellRange, строки данных
        header_row_range — com.sun.star.table.XCellRange, строка заголовков
        *extra_args — rest из колонки C (см. _lm_pp_word_wrap_enabled)

    Карта MERGE_RESULT_POSTPROCESS_RANGE_MAP: «перенос».

    Когда вызывать:
        Перед «авто_высота» или вместе как «перенос_и_авто_высота»; после копирования данных.

    Что меняет:
        data_range.IsTextWrapped; при include_header — и header_row_range.
    """
    _unused = (doc, sheet)
    wrap_on, include_hdr = _lm_pp_wrap_settings(extra_args)
    data_range.IsTextWrapped = wrap_on
    if include_hdr and header_row_range is not None:
        header_row_range.IsTextWrapped = wrap_on


def _lm_pp_header_height_cfg_from_extra_or_block(doc, sheet, extra_args):
    """Блок JSON (dict) или текст/JSON в extra_args[0] → cfg."""
    if extra_args and isinstance(extra_args[0], dict):
        cfg = _lm_pp_header_height_block_to_cfg(extra_args[0])
        _lm_pp_header_height_dbg(
            "cfg_from_block",
            "sheet=%s block=%r cfg=%s"
            % (getattr(sheet, "Name", sheet), extra_args[0], cfg),
        )
        return cfg
    raw = ""
    if extra_args and len(extra_args) > 0:
        raw = str(extra_args[0] or "").strip()
    return _lm_pp_header_height_config_from_extra(doc, sheet, raw)


def _lm_pp_header_height_debug_readback(sheet, row, col=0):
    """Снять фактические свойства ячейки заголовка после apply."""
    if sheet is None:
        return
    try:
        cell = sheet.getCellByPosition(int(col), int(row))
        rh = None
        try:
            rh = sheet.getRows().getByIndex(int(row)).Height
        except Exception:
            pass
        _lm_pp_header_height_dbg(
            "readback",
            "sheet=%s row=%s col=%s row_h=%s CharWeight=%s HoriJustify=%s VertJustify=%s VertJustify2=%s CharFontName=%r CharHeight=%s CellBackColor=%s CharColor=%s CharAutoColor=%s"
            % (
                getattr(sheet, "Name", sheet),
                row,
                col,
                rh,
                getattr(cell, "CharWeight", None),
                getattr(cell, "HoriJustify", None),
                getattr(cell, "VertJustify", None),
                getattr(cell, "VertJustify2", None),
                getattr(cell, "CharFontName", None),
                getattr(cell, "CharHeight", None),
                getattr(cell, "CellBackColor", None),
                getattr(cell, "CharColor", None),
                getattr(cell, "CharAutoColor", None),
            ),
        )
    except Exception as err:
        _lm_pp_header_height_dbg("readback", "err=%s" % err)


def lm_pp_range_header_row_height_pad(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Задать высоту строки заголовков (абсолютно или плюсом к текущей) и выровнять ячейки.

    JSON колонки C:
        [{"v":1,"fn":"заголовок_плюс_высота","height_mm":13,"add_height":false,
          "h_align":"center","v_align":"center"}]
        add_height=false (по умолчанию) — поставить height_mm как есть (мин. 7 мм);
        add_height=true — прибавить height_mm к текущей высоте (итог тоже ≥ 7 мм).

    Карта: «заголовок_плюс_высота».
    """
    _unused = (doc, data_range)
    extra_preview = ""
    try:
        if extra_args:
            extra_preview = str(extra_args[0])[:240]
    except Exception:
        pass
    sheet_name = ""
    try:
        sheet_name = str(sheet.Name)
    except Exception:
        pass
    if _lm_pp_param_skip_sheet("заголовок_плюс_высота", extra_args, doc, sheet):
        return
    _lm_pp_header_height_dbg(
        "lm_pp_range_header_row_height_pad",
        "sheet=%s extra=%r"
        % (
            sheet_name,
            extra_preview if not (extra_args and isinstance(extra_args[0], dict)) else extra_args[0],
        ),
    )
    cfg = _lm_pp_header_height_cfg_from_extra_or_block(doc, sheet, extra_args)
    if cfg is None:
        _lm_pp_header_height_dbg(
            "lm_pp_range_header_row_height_pad",
            "SKIP cfg=None sheet=%s" % sheet_name,
        )
        return
    _lm_pp_header_height_dbg(
        "lm_pp_range_header_row_height_pad",
        "cfg=%s" % cfg,
    )
    _lm_pp_apply_header_row_height_pad(sheet, header_row_range, cfg)


def _lm_pp_sheet_list_skip(fn_key, extra_args, sheet):
    """True — не выполнять шаг на этом листе (фильтр sheets из C/JSON)."""
    return _lm_pp_sheet_list_block_for_sheet(fn_key, extra_args, sheet) is None


def lm_pp_range_autofit_row_heights(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Автоподбор высоты каждой строки внутри data_range (OptimalHeight).

    Параметры:
        doc — SpreadsheetDocument (не используется)
        sheet — XSpreadsheet
        data_range — XCellRange данных
        header_row_range — не используется

    JSON колонки C: [{"v":1,"fn":"авто_высота"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «авто_высота».

    Когда вызывать:
        После «перенос»; не сочетать с «высота_строки» на row без необходимости.

    Что меняет:
        OptimalHeight = True для строк sr…er из data_range.
    """
    if _lm_pp_sheet_list_skip("авто_высота", extra_args, sheet):
        return
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    rows = sheet.getRows()

    def _autofit_row(sheet, r):
        rows.getByIndex(r).OptimalHeight = True

    _lm_pp_foreach_row(sheet, sr, er, _autofit_row, status_prefix="диапазон")


def lm_pp_range_autofit_columns_width(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Автоподбор ширины столбцов диапазона (Column.OptimalWidth по API).

    Параметры:
        doc — SpreadsheetDocument (пересчёт формул перед подбором)
        sheet — XSpreadsheet
        data_range — XCellRange данных
        header_row_range — строка заголовков (учитывается при определении столбцов)

    JSON колонки C: [{"v":1,"fn":"авто_ширина"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «авто_ширина».

    Что меняет:
        OptimalWidth = True для столбцов sc..ec; строки — от заголовка до конца данных.
    """
    if _lm_pp_sheet_list_skip("авто_ширина", extra_args, sheet):
        return
    sc_d, sr_d, ec_d, er_d = _lm_pp_range_address(data_range)
    sc_h, sr_h, ec_h, er_h = _lm_pp_range_address(header_row_range)
    sc = min(sc_d, sc_h)
    ec = max(ec_d, ec_h)
    sr = min(sr_d, sr_h)
    er = max(er_d, er_h)
    if ec < sc:
        return
    col_indices = []
    c = sc
    while c <= ec:
        col_indices.append(c)
        c = c + 1
    _lm_pp_autofit_columns(doc, sheet, col_indices, sr=sr, er=er)


def lm_pp_range_word_wrap_and_fit_rows(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Композит: перенос по словам + автоподбор высоты строк данных.

    Параметры: см. lm_pp_range_word_wrap (*extra_args передаётся в «перенос»).

    JSON колонки C: [{"v":1,"fn":"перенос_и_авто_высота","wrap":true,"include_header":true}].

    Карта: «перенос_и_авто_высота».

    Когда вызывать:
        Вместо пары «перенос» + «авто_высота» одной строкой на листе параметров.

    Что меняет:
        То же, что два колбэка по порядку (см. вызываемые функции).
    """
    lm_pp_range_word_wrap(
        doc, sheet, data_range, header_row_range, *extra_args
    )
    lm_pp_range_autofit_row_heights(doc, sheet, data_range, header_row_range)
    _wrap_on, include_hdr = _lm_pp_wrap_settings(extra_args)
    if include_hdr and header_row_range is not None and _wrap_on:
        try:
            hr = int(header_row_range.getRangeAddress().StartRow)
            sheet.getRows().getByIndex(hr).OptimalHeight = True
        except Exception:
            pass


def _lm_pp_payload_is_json(text):
    """True, если строка — валидный JSON (в т.ч. Calc: внешние кавычки, «»)."""
    s = str(text or "").strip()
    if s == "":
        return False
    try:
        from libre_macros_param_codec import _looks_like_json_payload, _parse_json_root

        if not _looks_like_json_payload(s):
            if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
                inner = s[1:-1].strip()
                if not _looks_like_json_payload(inner):
                    return False
                s = inner
            else:
                return False
        _parse_json_root(s)
        return True
    except (ValueError, TypeError, ImportError):
        return False


def _lm_pp_param_decode(fn_key, extra_args):
    """JSON из C → list[dict] блоков."""
    text = _lm_pp_join_extra_args(extra_args)
    text = str(text or "").strip()
    if text == "":
        return []
    from libre_macros_param_codec import normalize_fn_key, param_decode

    return param_decode(normalize_fn_key(fn_key), text)


def _lm_pp_param_block_for_sheet(fn_key, extra_args, doc, sheet):
    """Блок параметров для текущего листа или None."""
    blocks = _lm_pp_blocks_for_sheet(fn_key, extra_args, doc, sheet)
    if len(blocks) == 0:
        return None
    return blocks[0]


def _lm_pp_param_skip_sheet(fn_key, extra_args, doc, sheet):
    """True — JSON задан, но блок для листа не найден."""
    raw = _lm_pp_join_extra_args(extra_args).strip()
    if raw == "" or not _lm_pp_payload_is_json(raw):
        return False
    return len(_lm_pp_blocks_for_sheet(fn_key, extra_args, doc, sheet)) == 0


def lm_pp_format_step_applies_to_sheet(fn_key, extra_raw, sheet_name, doc=None, sheet=None):
    """True, если шаг «формат_*» относится к указанному листу."""
    raw = str(extra_raw or "").strip()
    if raw == "":
        return True
    if not _lm_pp_payload_is_json(raw):
        return True
    try:
        from libre_macros_param_codec import normalize_fn_key
    except ImportError:
        return True
    fn_key = normalize_fn_key(fn_key)
    if fn_key not in ("формат_деньги", "формат_даты", "формат_столбцы"):
        return True
    target = sheet
    if target is None and doc is not None:
        target = _lm_vlookup_resolve_sheet(doc, sheet_name)
    return len(_lm_pp_blocks_for_sheet(fn_key, [raw], doc, target, sheet_name)) > 0


def _lm_pp_codec_text_args(fn_key, extra_raw):
    """Колонка C → extra_args (сырой JSON одной строкой)."""
    s = str(extra_raw).strip() if extra_raw is not None else ""
    return [s] if s else []


def _lm_pp_grid_config_from_block(block):
    from libre_macros_param_codec import _normalize_grid_block

    block = _normalize_grid_block(block)
    width_token = str(block.get("width") or "тонкая").strip()
    w = _lm_pp_grid_parse_width_token(width_token)
    if w is None:
        try:
            w = float(width_token.replace(",", "."))
        except (TypeError, ValueError):
            w = None
    color = _LM_PP_GRID_DEFAULT_COLOR
    if block.get("color"):
        resolved = _lm_pp_zebra_resolve_color_token(str(block.get("color")))
        if resolved is not None:
            color = resolved
    width = w if w is not None else _LM_PP_GRID_DEFAULT_WIDTH
    return width, color, bool(block.get("include_header"))


def _lm_pp_zebra_styles_from_block(block):
    from libre_macros_param_codec import _normalize_zebra_block

    block = _normalize_zebra_block(block)
    roles = block.get("roles") or {}
    styles = {"header": None, "even": None, "odd": None}
    for role in ("header", "even", "odd"):
        spec = roles.get(role)
        if not isinstance(spec, dict):
            continue
        fill = str(spec.get("fill") or "").strip()
        font = str(spec.get("font") or "").strip()
        if fill or font:
            val = fill
            if font:
                val = "%s/%s" % (fill, font) if fill else font
            styles[role] = _lm_pp_zebra_parse_stripe_spec(val)
    has_data_role = styles["even"] is not None or styles["odd"] is not None
    if not has_data_role and styles["header"] is None:
        styles["odd"] = _lm_pp_zebra_default_even_style()
    elif styles["even"] is None and has_data_role:
        styles["even"] = _lm_pp_zebra_default_even_style()
    return styles


def _lm_pp_zebra_styles_for_sheet(extra_args, doc, sheet):
    block = _lm_pp_param_block_for_sheet("зебра_диапазон", extra_args, doc, sheet)
    if block is not None:
        return _lm_pp_zebra_styles_from_block(block)
    raw = _lm_pp_join_extra_args(extra_args).strip()
    if raw == "":
        return _lm_pp_zebra_parse_rest("")
    if _lm_pp_payload_is_json(raw):
        return {"header": None, "even": None, "odd": _lm_pp_zebra_default_even_style()}
    return _lm_pp_zebra_parse_rest(raw)


def _lm_pp_threshold_params_from_block(block):
    """
    Разобрать блок «подсветка_по_порогу» / «подкрасить_пороги».

    Возвращает dict:
      columns — list[str] маркеров/имён столбцов;
      threshold_min / threshold_max — float|None (диапазон, включительно);
      legacy_gt — float|None (старое «строго больше»);
      fill_color — int LO RGB;
      font_color — int|None;
      font_color_auto — bool;
      bold / italic — bool.
    """
    from libre_macros_param_codec import _normalize_threshold_block

    block = _normalize_threshold_block(block)
    columns = []
    raw_cols = block.get("columns") or []
    if isinstance(raw_cols, (list, tuple)):
        for c in raw_cols:
            s = str(c).strip()
            if s:
                columns.append(s)
    marker = str(block.get("marker") or "").strip()
    if not columns and marker:
        columns = [marker]
    if not columns:
        columns = [_LM_PP_DEFAULT_THRESHOLD_HIGHLIGHT_MARKER]

    fill = None
    color_raw = str(block.get("color") or block.get("fill_color") or "").strip()
    if color_raw:
        fill = _lm_pp_zebra_resolve_color_token(color_raw)
        # Неразрешённый токен — не подставляем цвет по умолчанию (пустая заливка = не менять).

    style_columns = []
    raw_style = block.get("style_columns") or []
    if isinstance(raw_style, (list, tuple)):
        for c in raw_style:
            s = str(c).strip()
            if s:
                style_columns.append(s)

    font_auto = bool(block.get("font_color_auto", True))
    font_color = None
    fc_tok = str(block.get("font_color") or "").strip()
    if fc_tok:
        font_auto = False
        font_color = _lm_pp_zebra_resolve_color_token(fc_tok)

    tmin = block.get("threshold_min")
    tmax = block.get("threshold_max")
    legacy_gt = None
    if tmin is None and tmax is None and block.get("threshold") is not None:
        try:
            legacy_gt = float(block.get("threshold"))
        except (TypeError, ValueError):
            legacy_gt = _LM_PP_DEFAULT_THRESHOLD_HIGHLIGHT_VALUE
    else:
        try:
            tmin = float(tmin) if tmin is not None else None
        except (TypeError, ValueError):
            tmin = None
        try:
            tmax = float(tmax) if tmax is not None else None
        except (TypeError, ValueError):
            tmax = None

    return {
        "columns": columns,
        "threshold_min": tmin,
        "threshold_max": tmax,
        "legacy_gt": legacy_gt,
        "fill_color": fill,
        "font_color": font_color,
        "font_color_auto": font_auto,
        "bold": bool(block.get("bold")),
        "italic": bool(block.get("italic")),
        "whole_row": bool(block.get("whole_row")),
        "style_columns": style_columns,
    }


def _lm_pp_threshold_resolve_marker_columns(sc, ec, titles, needles):
    """0-based индексы столбцов диапазона, совпавших с маркерами needles."""
    target_cols = []
    c = int(sc)
    while c <= int(ec):
        idx = c - int(sc)
        title = titles[idx] if idx < len(titles) else ""
        matched = False
        ni = 0
        while ni < len(needles):
            needle = needles[ni]
            if needle.isdigit():
                try:
                    n = int(needle)
                except Exception:
                    n = -1
                if n > 0 and (c + 1 == n or idx + 1 == n):
                    matched = True
                    break
            elif len(needle) <= 3 and needle.isalpha():
                try:
                    if int(col_letters_to_index(needle)) == c:
                        matched = True
                        break
                except Exception:
                    pass
                if _lm_pp_header_matches_marker(title, needle):
                    matched = True
                    break
            elif _lm_pp_header_matches_marker(title, needle):
                matched = True
                break
            ni = ni + 1
        if matched:
            target_cols.append(c)
        c = c + 1
    return target_cols


def _lm_pp_threshold_paint_columns(spec, target_cols, sc, ec, titles):
    """
    Столбцы, куда применять стиль:
    style_columns — явный список;
    whole_row — sc..ec (вызывающий расширяет ec до used);
    иначе — только target_cols (столбцы порога).
    """
    style_tokens = spec.get("style_columns") or []
    if style_tokens:
        needles = []
        for col_tok in style_tokens:
            n = str(col_tok).strip().lower()
            if n:
                needles.append(n)
        if needles:
            cols = _lm_pp_threshold_resolve_marker_columns(sc, ec, titles, needles)
            if cols:
                return cols
    if bool(spec.get("whole_row")):
        out = []
        cc = int(sc)
        while cc <= int(ec):
            out.append(cc)
            cc = cc + 1
        return out
    return list(target_cols)


# Только то, что подсветка по порогу может выставить (не трогаем рамки, формат числа и т.д.).
_LM_PP_THRESHOLD_RESET_PROPS = (
    "CellBackColor",
    "IsCellBackgroundTransparent",
    "CharColor",
    "CharColor2",
    "CharAutoColor",
    "CharWeight",
    "CharPosture",
)


def _lm_pp_threshold_reset_range_style(cell_or_range):
    """Сброс заливки / цвета шрифта / жирности / курсива перед подсветкой."""
    if cell_or_range is None:
        return
    pi = 0
    while pi < len(_LM_PP_THRESHOLD_RESET_PROPS):
        prop = _LM_PP_THRESHOLD_RESET_PROPS[pi]
        try:
            cell_or_range.setPropertyToDefault(prop)
        except Exception:
            pass
        pi = pi + 1


def _lm_pp_threshold_reset_paint_cols(sheet, paint_cols, runs):
    """Сброс только оформления подсветки на целевых столбцах matching-строк."""
    if sheet is None or not paint_cols or not runs:
        return
    ri = 0
    while ri < len(runs):
        r0, r1 = runs[ri]
        ci = 0
        while ci < len(paint_cols):
            cc = int(paint_cols[ci])
            try:
                rng = sheet.getCellRangeByPosition(cc, int(r0), cc, int(r1))
                _lm_pp_threshold_reset_range_style(rng)
            except Exception:
                pass
            ci = ci + 1
        ri = ri + 1


def _lm_pp_threshold_value_matches(val, spec):
    """True, если числовое значение попадает под пороги spec."""
    try:
        num = float(val)
    except (TypeError, ValueError):
        return False
    legacy = spec.get("legacy_gt")
    if legacy is not None and spec.get("threshold_min") is None and spec.get("threshold_max") is None:
        return num > float(legacy)
    tmin = spec.get("threshold_min")
    tmax = spec.get("threshold_max")
    if tmin is None and tmax is None:
        return False
    if tmin is not None and num < float(tmin):
        return False
    if tmax is not None and num > float(tmax):
        return False
    return True


def _lm_pp_apply_threshold_cell_style(cell_or_range, spec):
    """Заливка + опционально шрифт / жирный / курсив (ячейка, ряд или SheetCellRanges)."""
    fill = spec.get("fill_color")
    if fill is not None:
        try:
            cell_or_range.CellBackColor = int(fill)
        except Exception:
            pass
    if not spec.get("font_color_auto") and spec.get("font_color") is not None:
        try:
            _lm_set_char_color_safe(cell_or_range, int(spec.get("font_color")))
        except Exception:
            pass
    if spec.get("bold"):
        try:
            cell_or_range.CharWeight = _LM_FMT_CHAR_WEIGHT_BOLD
        except Exception:
            pass
    if spec.get("italic"):
        try:
            from com.sun.star.awt.FontSlant import ITALIC

            cell_or_range.CharPosture = ITALIC
        except Exception:
            try:
                cell_or_range.CharPosture = 2
            except Exception:
                pass


def _lm_pp_coalesce_row_runs(rows):
    """Отсортированные 0-based номера строк → список (r0, r1) включительно."""
    out = []
    if not rows:
        return out
    ordered = sorted(set(int(r) for r in rows))
    run_start = ordered[0]
    prev = ordered[0]
    i = 1
    while i < len(ordered):
        cur = ordered[i]
        if cur == prev + 1:
            prev = cur
        else:
            out.append((run_start, prev))
            run_start = cur
            prev = cur
        i = i + 1
    out.append((run_start, prev))
    return out


def _lm_pp_threshold_collect_matching_rows(sheet, target_cols, sr, er, spec):
    """
    Строки data-блока, где хотя бы один целевой столбец попадает под порог.

    Сначала пытаемся прочитать столбец через getDataArray (один UNO на столбец),
    иначе — по ячейкам.

    Возвращает (matching_rows, scan_mode):
      scan_mode — «DataArray» | «ячейки» | «DataArray+ячейки» | «нет столбцов».
    """
    matching = []
    seen = set()
    top = int(sr)
    bottom = int(er)
    used_data_array = False
    used_cells = False
    ti = 0
    while ti < len(target_cols or []):
        col = int(target_cols[ti])
        ti = ti + 1
        data = None
        try:
            col_range = sheet.getCellRangeByPosition(col, top, col, bottom)
            data = col_range.getDataArray()
        except Exception:
            data = None
        if data is not None:
            used_data_array = True
            ri = 0
            while ri < len(data):
                row_tup = data[ri]
                val = row_tup[0] if row_tup else None
                if _lm_pp_threshold_value_matches(val, spec):
                    row_idx = top + ri
                    if row_idx not in seen:
                        seen.add(row_idx)
                        matching.append(row_idx)
                ri = ri + 1
            continue
        used_cells = True
        r = top
        while r <= bottom:
            try:
                val = sheet.getCellByPosition(col, r).Value
            except Exception:
                r = r + 1
                continue
            if _lm_pp_threshold_value_matches(val, spec):
                if r not in seen:
                    seen.add(r)
                    matching.append(r)
            r = r + 1
    matching.sort()
    if used_data_array and used_cells:
        scan_mode = u"DataArray+ячейки"
    elif used_data_array:
        scan_mode = u"DataArray"
    elif used_cells:
        scan_mode = u"ячейки"
    else:
        scan_mode = u"нет столбцов"
    return matching, scan_mode


def _lm_pp_threshold_make_multi_ranges(doc, sheet, runs, paint_cols):
    """
    Непрерывные блоки строк × столбцы подсветки → SheetCellRanges (как Ctrl+выделение).

    Возвращает объект с addRangeAddress или None.
    """
    if doc is None or sheet is None or not runs or not paint_cols:
        return None
    try:
        ranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")
    except Exception:
        ranges = None
    if ranges is None:
        return None
    added = 0
    for r0, r1 in runs:
        ci = 0
        while ci < len(paint_cols):
            col = int(paint_cols[ci])
            ci = ci + 1
            try:
                rng = sheet.getCellRangeByPosition(col, int(r0), col, int(r1))
                addr = rng.getRangeAddress()
                ranges.addRangeAddress(addr, False)
                added = added + 1
            except Exception:
                pass
    if added <= 0:
        return None
    return ranges


def _lm_pp_threshold_apply_batch(doc, sheet, target_cols, sc, sr, er, ec_for_paint, titles, spec):
    """
    Собрать строки по порогу, сбросить стили целевых столбцов, применить оформление.

    ec_for_paint — правая граница диапазона для разрешения style_columns / whole_row
    (при whole_row — расширенная до used bounds).
    """
    sheet_name = ""
    try:
        sheet_name = str(getattr(sheet, "Name", "") or "")
    except Exception:
        sheet_name = ""
    paint_cols = _lm_pp_threshold_paint_columns(
        spec, target_cols, int(sc), int(ec_for_paint), titles
    )
    if not paint_cols:
        return 0
    matching, scan_mode = _lm_pp_threshold_collect_matching_rows(
        sheet, target_cols, sr, er, spec
    )
    if not matching:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "подсветка_по_порогу",
            "инфо",
            "порог: нет строк под условие (скан=%s)" % scan_mode,
        )
        return 0
    runs = _lm_pp_coalesce_row_runs(matching)
    _lm_pp_threshold_reset_paint_cols(sheet, paint_cols, runs)
    multi = _lm_pp_threshold_make_multi_ranges(doc, sheet, runs, paint_cols)
    cols_desc = ",".join(str(int(c) + 1) for c in paint_cols[:8])
    if len(paint_cols) > 8:
        cols_desc = cols_desc + ",…"
    if multi is not None:
        try:
            _lm_pp_apply_threshold_cell_style(multi, spec)
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                "подсветка_по_порогу",
                "ok",
                "порог: стиль=SheetCellRanges (мульти), скан=%s, строк %d, блоков %d, cols [%s]"
                % (scan_mode, len(matching), len(runs), cols_desc),
            )
            return len(matching)
        except Exception as multi_err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                "подсветка_по_порогу",
                "инфо",
                "порог: SheetCellRanges FAIL → запас (%s)" % str(multi_err)[:120],
            )
    runs_ok = 0
    cells_ok = 0
    for r0, r1 in runs:
        ci = 0
        while ci < len(paint_cols):
            col = int(paint_cols[ci])
            ci = ci + 1
            try:
                col_range = sheet.getCellRangeByPosition(col, int(r0), col, int(r1))
                _lm_pp_apply_threshold_cell_style(col_range, spec)
                runs_ok = runs_ok + 1
            except Exception:
                cells_ok = cells_ok + 1
                r = int(r0)
                while r <= int(r1):
                    try:
                        _lm_pp_apply_threshold_cell_style(
                            sheet.getCellByPosition(col, r), spec
                        )
                    except Exception:
                        pass
                    r = r + 1
    if cells_ok == 0:
        style_mode = u"диапазоны по блокам"
    elif runs_ok == 0:
        style_mode = u"поячеечно"
    else:
        style_mode = u"смешанный (блоки+ячейки)"
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        "подсветка_по_порогу",
        "ok",
        "порог: стиль=%s, скан=%s, строк %d, блоков %d (range=%d, cell=%d), cols [%s]"
        % (
            style_mode,
            scan_mode,
            len(matching),
            len(runs),
            runs_ok,
            cells_ok,
            cols_desc,
        ),
    )
    return len(matching)


def _lm_pp_threshold_apply_whole_row_batch(doc, sheet, target_cols, sc, sr, er, paint_ec, spec):
    """Обратная совместимость: делегирует в _lm_pp_threshold_apply_batch."""
    titles = []
    try:
        hr = int(sr) - 1
        if hr >= 0:
            titles = _lm_pp_header_titles(sheet, sheet.getCellRangeByPosition(int(sc), hr, int(paint_ec), hr))
    except Exception:
        titles = []
    return _lm_pp_threshold_apply_batch(
        doc, sheet, target_cols, sc, sr, er, paint_ec, titles, spec
    )


def _lm_pp_font_spec_from_block(block):
    from libre_macros_param_codec import _normalize_font_block

    block = _normalize_font_block(block)
    return {
        "name": block.get("name"),
        "data_size": block.get("size_data"),
        "header_size": block.get("size_header"),
    }


def _lm_pp_dim_mm_from_block(block, kind, default_mm):
    from libre_macros_param_codec import (
        _normalize_column_width_block,
        _normalize_row_height_block,
    )

    if kind == "columns":
        block = _normalize_column_width_block(block)
        mm = block.get("width_mm", default_mm)
        indices_val = block.get("columns")
    else:
        block = _normalize_row_height_block(block)
        mm = block.get("height_mm", default_mm)
        indices_val = block.get("rows")
    if indices_val in (None, "", "all"):
        indices = None
    elif isinstance(indices_val, list):
        indices = indices_val
    else:
        indices = [indices_val]
    return indices, float(mm)


def _lm_pp_print_style_from_block(block):
    from libre_macros_param_codec import _normalize_print_style_block

    block = _normalize_print_style_block(block)
    return block.get("orientation") or "landscape", int(block.get("fit_pages") or 1)


def _lm_pp_row_height_extra_args(extra_raw):
    """Колонка C для «высота_строки» — одна строка (запятая — часть синтаксиса номеров строк)."""
    return _lm_pp_codec_text_args("высота_строки", extra_raw)


def _lm_pp_mm_to_row_height(mm):
    """Миллиметры → единицы высоты строки Calc (1/100 мм)."""
    return int(round(float(mm) * 100.0))


def _lm_pp_parse_mm_number(text):
    """
    Число миллиметров из фрагмента C: «13», «+12,7 мм», «0,5″» (дюймы → мм).
    """
    raw = str(text or "").strip()
    if raw == "":
        raise ValueError("empty mm")
    inch = bool(
        re.search(
            r"(?:″|\"|''|in(?:ch(?:es)?)?|дюйм)\s*$",
            raw,
            re.IGNORECASE,
        )
    )
    num_part = re.sub(r"^\+\s*", "", raw)
    num_part = re.sub(
        r"\s*(?:мм|mm|″|\"|''|in(?:ch(?:es)?)?|дюйм)\s*$",
        "",
        num_part,
        flags=re.IGNORECASE,
    ).strip()
    num_part = num_part.replace(",", ".")
    m = re.match(r"^([\d.]+)", num_part)
    if m is None:
        raise ValueError("no number: %r" % text)
    val = float(m.group(1))
    if inch:
        val = val * 25.4
    return val


def _lm_pp_apply_row_height_mm(sheet, row_index, height_mm):
    row = sheet.getRows().getByIndex(int(row_index))
    if float(height_mm) == 0.0:
        row.IsVisible = False
        return
    row.IsVisible = True
    row.OptimalHeight = False
    row.Height = _lm_pp_mm_to_row_height(height_mm)


def _lm_pp_row_height_target_matches(row_index, row_specs):
    """row_index 0-based; row_specs — номера строк 1-based или None = каждая строка цикла."""
    if not row_specs:
        return True
    return (int(row_index) + 1) in row_specs


def lm_pp_range_row_height_pad(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Задать высоту строк data_range в миллиметрах (см. _lm_pp_parse_row_height_rest).

    Колонка C — как у «ширина_столбцов»:
      пусто — все строки данных, 5 мм;
      «6» — все строки данных, 6 мм;
      «2,3,4 / 6» (или ->, =) — строки 1-based на листе;
      «0» — скрыть строку(и) (IsVisible=False).

    Карта: «высота_строки» (диапазон).
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    _unused = (h_sc, h_ec, h_er, sc)
    data_start = int(h_sr) + 1
    if sr < data_start:
        sr = data_start
    if er < sr:
        return

    row_specs, height_mm = _lm_pp_parse_row_height_rest(extra_args)

    def _apply_row(r):
        if r < 0 or r > er:
            return
        _lm_pp_apply_row_height_mm(sheet, r, height_mm)

    if not row_specs:
        r = sr
        while r <= er:
            _apply_row(r)
            r = r + 1
    else:
        i = 0
        while i < len(row_specs):
            _apply_row(int(row_specs[i]) - 1)
            i = i + 1


# =============================================================================
# Сетка (range) — JSON в колонке C: width, color, include_header
# =============================================================================

_LM_PP_GRID_DEFAULT_WIDTH = 15
_LM_PP_GRID_DEFAULT_COLOR = 0x666666

_LM_PP_GRID_WIDTH_ALIASES = {
    "тонкая": 15,
    "thin": 15,
    "толстая": 35,
    "thick": 35,
    "средняя": 25,
    "medium": 25,
}


def _lm_pp_grid_parse_width_token(token):
    key = str(token or "").strip().lower()
    if key == "":
        return None
    if key in _LM_PP_GRID_WIDTH_ALIASES:
        return _LM_PP_GRID_WIDTH_ALIASES[key]
    try:
        v = int(key)
        if v > 0:
            return v
    except (TypeError, ValueError):
        pass
    return None


_LM_PP_GRID_REST_HINT = (
    "толщина/цвет — тонкая|толстая|число (1/100 мм) / имя цвета или hex; "
    "пустая часть — по умолчанию; пусто = тонкая серая; "
    "по листам: «лист | 25/#CCCCCC ; лист2 : толстая/синий» (разделитель лист/параметры: |, : или /)",
    "Сводная | 25/#CCCCCC ; Orders : толстая/темныйсиний",
)

_LM_PP_GRID_PRESET_REST_HINT = (
    "список листов через запятую или ; (пусто — все листы)",
    "Сводная, Orders",
)


def _lm_pp_is_grid_postprocess_name(fn_name):
    return str(fn_name or "").strip().lower() == "сетка"


def _lm_pp_is_grid_preset_postprocess_name(fn_name):
    key = str(fn_name or "").strip().lower()
    return key in ("тонкая_сетка", "толстая_сетка")


def _lm_pp_is_any_grid_postprocess_name(fn_name):
    key = str(fn_name or "").strip().lower()
    return key in ("сетка", "тонкая_сетка", "толстая_сетка")


def _lm_pp_grid_extra_args(extra_raw):
    return _lm_pp_codec_text_args("сетка", extra_raw)


def _lm_pp_looks_like_grid_rest_only(text):
    """True, если C — только толщина/цвет без префикса листа."""
    raw = str(text or "").strip()
    if raw == "":
        return True
    if "|" in raw or ":" in raw:
        return False
    if re.search(r"[,;]", raw):
        return False
    if "/" not in raw:
        if raw.lower() in _LM_PP_GRID_WIDTH_ALIASES:
            return True
        if _lm_pp_grid_parse_width_token(raw) is not None and re.match(
            r"^\d+$", raw.strip()
        ):
            return True
        if _lm_pp_zebra_resolve_color_token(raw) is not None:
            return True
        return False
    left, _, right = raw.partition("/")
    left = left.strip()
    right = right.strip()
    if left.lower() in _LM_PP_GRID_WIDTH_ALIASES:
        return True
    if _lm_pp_grid_parse_width_token(left) is not None and (
        right.startswith("#")
        or _lm_pp_zebra_resolve_color_token(right) is not None
    ):
        return True
    return False


def _lm_pp_split_grid_slash_assignment(block):
    """Разделить блок «лист / параметры» (если это не просто толщина/цвет)."""
    block = str(block or "").strip()
    if block == "" or "|" in block or ":" in block:
        return None, block
    idx = block.find("/")
    if idx < 0:
        return None, block
    left = block[:idx].strip()
    right = block[idx + 1 :].strip()
    if left == "" or right == "":
        return None, block
    if left.lower() in _LM_PP_GRID_WIDTH_ALIASES:
        return None, block
    w = _lm_pp_grid_parse_width_token(left)
    if w is not None and (
        right.startswith("#") or _lm_pp_zebra_resolve_color_token(right) is not None
    ):
        return None, block
    return left, right


def _lm_pp_block_looks_like_slash_grid_assignment(block):
    sheet_part, param_part = _lm_pp_split_grid_slash_assignment(block)
    return sheet_part is not None and param_part is not None


def _lm_pp_split_grid_assignment_blocks(text):
    """Блоки «лист | параметры» через ; или , (в т.ч. лист / параметры)."""
    text = str(text or "").strip()
    if text == "":
        return []
    if ";" in text:
        blocks = re.split(r"\s*;\s*", text)
    elif re.search(r"(?:\||:)", text) and "," in text:
        blocks = re.split(r",\s*(?=[^,]+(?:\||:))", text)
    elif "," in text:
        parts = [p.strip() for p in text.split(",") if p.strip()]
        if len(parts) > 1 and all(
            _lm_pp_block_looks_like_slash_grid_assignment(p) for p in parts
        ):
            blocks = parts
        else:
            return [text]
    else:
        return [text]
    out = []
    bi = 0
    while bi < len(blocks):
        block = str(blocks[bi]).strip()
        bi = bi + 1
        if block != "":
            out.append(block)
    return out


def _lm_pp_parse_grid_sheet_assignments(text):
    """
    «лист | 25/#CCCCCC ; лист2 : толстая/синий» или «лист / 25/#CCCCCC».
  Пустой список — C без привязки к листам (все листы, один rest).
    """
    text = str(text or "").strip()
    if text == "" or _lm_pp_looks_like_grid_rest_only(text):
        return []
    assignments = _lm_pp_parse_sheet_param_assignments(text)
    if len(assignments) > 0:
        return assignments
    blocks = _lm_pp_split_grid_assignment_blocks(text)
    out = []
    bi = 0
    while bi < len(blocks):
        block = str(blocks[bi]).strip()
        bi = bi + 1
        if block == "":
            continue
        sheet_part, param_part = _lm_pp_split_grid_slash_assignment(block)
        if sheet_part is None:
            continue
        sheet_names = _lm_pp_parse_sheet_names_from_spec(sheet_part)
        if len(sheet_names) == 0 or param_part == "":
            continue
        out.append((sheet_names, param_part))
    return out


def _lm_pp_parse_grid_preset_sheet_names(text):
    """
    Legacy-список листов для «тонкая_сетка» / «толстая_сетка» (не JSON).
    None — все листы; иначе список имён.
    """
    text = str(text or "").strip()
    if text == "":
        return None
    # Legacy assignment-синтаксис («лист | …») — не список имён пресета.
    if "|" in text or (":" in text and not _lm_pp_payload_is_json(text)):
        return []
    if _lm_pp_looks_like_grid_rest_only(text):
        return None
    names = _lm_pp_parse_sheet_names_from_spec(text)
    return names if len(names) > 0 else []


def _lm_pp_grid_config_for_sheet(extra_args, doc, sheet):
    """(width, color, include_header) для листа или None — шаг пропустить."""
    block = _lm_pp_param_block_for_sheet("сетка", extra_args, doc, sheet)
    if block is not None:
        return _lm_pp_grid_config_from_block(block)
    raw = _lm_pp_join_extra_args(extra_args).strip()
    if raw == "":
        return (_LM_PP_GRID_DEFAULT_WIDTH, _LM_PP_GRID_DEFAULT_COLOR, False)
    if _lm_pp_payload_is_json(raw):
        return None
    return None


def _lm_pp_grid_preset_applies_to_sheet( extra_raw, sheet_name, doc=None, sheet=None, fn_key="тонкая_сетка" ):
    """Фильтр «тонкая_сетка» / «толстая_сетка» по списку листов в C (JSON или legacy)."""
    _unused = (doc, sheet)
    s = str(extra_raw or "").strip()
    if s == "":
        return True
    key = str(fn_key or "тонкая_сетка").strip().lower() or "тонкая_сетка"
    if _lm_pp_payload_is_json(s):
        blocks = _lm_pp_param_decode(key, (s,))
        if len(blocks) == 0:
            return True
        nm = str(sheet_name or "").strip()
        bi = 0
        while bi < len(blocks):
            block = blocks[bi]
            bi = bi + 1
            sheets = block.get("sheets")
            if not sheets:
                return True
            if _lm_pp_sheet_name_matches_target(nm, sheets):
                return True
        return False
    names = _lm_pp_parse_grid_preset_sheet_names(s)
    if names is None:
        return True
    if len(names) == 0:
        return False
    return _lm_pp_sheet_name_matches_target(sheet_name, names)


def _lm_pp_custom_grid_step_applies_to_sheet(extra_raw, sheet_name, doc=None, sheet=None):
    """Нужно ли выполнять «сетка» на данном листе."""
    s = str(extra_raw or "").strip()
    if s == "":
        return True
    if not _lm_pp_payload_is_json(s):
        return True
    if sheet is None and doc is not None:
        sheet = _lm_vlookup_resolve_sheet(doc, sheet_name)
    if sheet is not None:
        return _lm_pp_grid_config_for_sheet((s,), doc, sheet) is not None
    blocks = _lm_pp_param_decode("сетка", (s,))
    if len(blocks) == 0:
        return True
    nm = str(sheet_name or "").strip()
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        bi = bi + 1
        sheet_ref = str(block.get("sheet") or "").strip()
        if sheet_ref == "":
            return True
        names = _lm_pp_parse_sheet_names_from_spec(sheet_ref)
        if _lm_pp_sheet_name_matches_target(nm, names):
            return True
    return False


def lm_pp_grid_step_applies_to_sheet(fn_name, extra_raw, sheet_name, doc=None, sheet=None):
    """
    Фильтр шагов «сетка» / «тонкая_сетка» / «толстая_сетка» по листу.
    Вызывается из pipeline как merge_pp_grid_step_applies_to_sheet (шим из lib).
    """
    key = str(fn_name or "").strip().lower()
    if key == "сетка":
        return _lm_pp_custom_grid_step_applies_to_sheet(
            extra_raw, sheet_name, doc, sheet
        )
    if key in ("тонкая_сетка", "толстая_сетка"):
        return _lm_pp_grid_preset_applies_to_sheet(
            extra_raw, sheet_name, doc, sheet, fn_key=key
        )
    return True


def lm_pp_grid_preset_step_applies_to_sheet(extra_raw, sheet_name, doc=None, sheet=None):
    """Нужно ли выполнять «тонкая_сетка» / «толстая_сетка» на данном листе."""
    return _lm_pp_grid_preset_applies_to_sheet(extra_raw, sheet_name, doc, sheet)


def _lm_pp_grid_parse_rest(rest_text):
    """
    Колонка C для «сетка»: «толщина/цвет» (как заливка/шрифт у «зебра»).
    «,заголовок» / «,header» — включить строку заголовков.

    Возвращает (width, color, include_header).
    """
    raw = str(rest_text or "").strip()
    include_header = False
    if "," in raw:
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if parts:
            raw = parts[0]
            pi = 1
            while pi < len(parts):
                tok = parts[pi].strip().casefold()
                if tok in ("header", "заголовок", "заг", "hdr"):
                    include_header = True
                pi = pi + 1
    width = _LM_PP_GRID_DEFAULT_WIDTH
    color = _LM_PP_GRID_DEFAULT_COLOR
    if raw == "":
        return width, color, include_header

    if "/" not in raw:
        w = _lm_pp_grid_parse_width_token(raw)
        if w is not None:
            return w, color, include_header
        c = _lm_pp_zebra_resolve_color_token(raw)
        if c is not None:
            return width, c, include_header
        return width, color, include_header

    thickness_t, _, color_t = raw.partition("/")
    thickness_t = thickness_t.strip()
    color_t = color_t.strip()
    if thickness_t:
        w = _lm_pp_grid_parse_width_token(thickness_t)
        if w is not None:
            width = w
    if color_t:
        c = _lm_pp_zebra_resolve_color_token(color_t)
        if c is not None:
            color = c
    return width, color, include_header


def _lm_pp_grid_rest_text(extra_args):
    return _lm_pp_join_extra_args(extra_args)


def lm_pp_range_grid_borders(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Универсальная сетка по data_range (по умолчанию без заголовка).

    Параметры:
        *extra_args — dict-блок из JSON (param_decode): width, color, include_header.

    JSON колонки C: [{"v":1,"fn":"сетка","width":"тонкая","color":"серый","include_header":true}]
    или с "sheet":"Лист". См. docs/13_JSON_PARAMS.md.

    Карта: «сетка».

    Когда вызывать:
        Вместо «тонкая_сетка» / «толстая_сетка», если нужен свой цвет или толщина.

    Что меняет:
        TableBorder / Top/Bottom/Left/RightBorder ячеек диапазона.
    """
    cfg = _lm_pp_grid_config_for_sheet(extra_args, doc, sheet)
    if cfg is None:
        return
    width, color, include_header = cfg
    _lm_pp_apply_range_grid_borders(
        data_range,
        color=color,
        width=width,
        doc=doc,
        sheet=sheet,
        grid_label="сетка",
        header_row_range=header_row_range,
        include_header=include_header,
    )


def lm_pp_range_thin_grid_borders(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Тонкая серая сетка по data_range и строке заголовков.

    Параметры:
        data_range, header_row_range — объединяются в один блок для линий сетки

    JSON колонки C: [{"v":1,"fn":"толстая_сетка"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «тонкая_сетка».

    Когда вызывать:
        Обычно в _MERGE_CODE_POSTPROCESS_RANGE; до построчной «зебры».

    Что меняет:
        Top/Bottom/Left/RightBorder каждой ячейки (формат ячеек Calc).
    """
    if _lm_pp_sheet_list_skip("тонкая_сетка", extra_args, sheet):
        return
    _lm_pp_apply_table_grid(
        data_range,
        doc=doc,
        sheet=sheet,
        grid_label="тонкая_сетка",
        header_row_range=header_row_range,
    )


def lm_pp_range_thick_grid_borders(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Толстая серая внутренняя сетка по data_range.

    Параметры: см. lm_pp_range_thin_grid_borders.

    JSON колонки C: [{"v":1,"fn":"толстая_сетка"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «толстая_сетка».

    Когда вызывать:
        Альтернатива «тонкая_сетка» (не обе сразу).

    Что меняет:
        Top/Bottom/Left/RightBorder каждой ячейки (формат ячеек Calc).
    """
    if _lm_pp_sheet_list_skip("толстая_сетка", extra_args, sheet):
        return
    _lm_pp_apply_table_grid_thick(
        data_range,
        doc=doc,
        sheet=sheet,
        grid_label="толстая_сетка",
        header_row_range=header_row_range,
    )


def _lm_pp_postprocess_fn_key(name):
    return str(name or "").strip().casefold()


def _lm_pp_is_delete_sheets_postprocess_name(fn_name):
    """Шаги удаления/скрытия листов — только в «Финальная_обработка», не в постобработке диапазона."""
    try:
        import libre_macros_final_lib as fl

        return fl._lm_final_is_pivot_followup_skip_name(fn_name)
    except ImportError:
        key = _lm_pp_postprocess_fn_key(fn_name)
        return key in (
            "удалить_служебные_листы",
            "удалить_листы",
            "удаление_листов",
            "скрытие_листов",
        )


def _lm_pp_is_param_sheet_name(sheet_name):
    """Лист параметров объединения — без импорта collect_workbooks (для диалога сводной)."""
    try:
        from libre_macros_global_settings_lib import is_param_sheet_name as _gs_is_param

        return bool(_gs_is_param(sheet_name))
    except Exception:
        pass
    s = str(sheet_name or "").strip()
    if s == "":
        return False
    for prefix in ("collect_params", "Параметры_Объединения"):
        if s == prefix or s.startswith(prefix + "_") or s.startswith(prefix):
            return True
    return False


def _lm_pp_uno_gc(enabled=True):
    """Сброс ссылок UNO-bridge после dispatch/диапазонов (как в ручном макросе)."""
    if not enabled:
        return
    try:
        import gc

        gc.collect()
    except Exception:
        pass


def _lm_pp_fill_empty_header_cells(sheet, header_row, sc, ec):
    """Пустые ячейки строки заголовков → «Колонка_*» (до автофильтра)."""
    if sheet is None:
        return
    try:
        header_row = int(header_row)
        sc = int(sc)
        ec = int(ec)
    except (TypeError, ValueError):
        return
    if ec < sc:
        return
    col = sc
    while col <= ec:
        try:
            cell = sheet.getCellByPosition(col, header_row)
            if cell_text(cell).strip() == "":
                cell.String = "Колонка_%s" % col_index_to_letters(col)
        except Exception:
            pass
        col = col + 1


def _lm_pp_create_silent_interaction_handler(ctx):
    if ctx is None:
        return None
    try:
        sm = ctx.ServiceManager
        return sm.createInstanceWithContext(
            "com.sun.star.comp.framework.SilentInteractionHandler", ctx
        )
    except Exception:
        return None


def _lm_pp_run_with_suppressed_ui(doc, callback):
    """Выполнить callback без системных диалогов LO (InteractionHandler)."""
    if not callable(callback):
        return None
    if doc is None:
        return callback()
    ctx = None
    frame = None
    old_handler = None
    handler = None
    try:
        try:
            ctx = XSCRIPTCONTEXT.getComponentContext()
        except NameError:
            ctx = uno.getComponentContext()
        handler = _lm_pp_create_silent_interaction_handler(ctx)
        controller = doc.getCurrentController()
        if controller is not None:
            frame = controller.getFrame()
        if frame is not None and handler is not None:
            try:
                old_handler = frame.getPropertyValue("InteractionHandler")
            except Exception:
                old_handler = None
            try:
                frame.setPropertyValue("InteractionHandler", handler)
            except Exception:
                old_handler = None
        return callback()
    finally:
        if frame is not None and handler is not None:
            try:
                frame.setPropertyValue("InteractionHandler", old_handler)
            except Exception:
                pass


def _lm_pp_execute_dispatch(doc, slot, props=(), uno_cleanup=True):
    """
    Выполнить UNO-команду на frame активного документа.

    Параметры:
        doc — SpreadsheetDocument
        slot — например «.uno:IncrementIndent»
        props — кортеж PropertyValue

    Возвращает:
        bool — True, если dispatch не выбросил исключение.
    """
    if doc is None:
        return False
    last_err = None
    ctx = None
    dispatcher = None
    controller = None
    frame = None
    try:
        try:
            ctx = XSCRIPTCONTEXT.getComponentContext()
        except NameError:
            ctx = uno.getComponentContext()
        dispatcher = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", ctx
        )
        controller = doc.getCurrentController()
        if controller is None:
            last_err = "no controller"
            return False
        frame = controller.getFrame()
        if dispatcher is None or frame is None:
            last_err = "no dispatcher/frame"
            return False
        for target in ("_self", ""):
            try:
                dispatcher.executeDispatch(frame, slot, target, 0, props)
                return True
            except Exception as err:
                last_err = err
        return False
    except Exception as err:
        last_err = err
        return False
    finally:
        if uno_cleanup:
            try:
                del dispatcher
            except Exception:
                pass
            try:
                del frame
            except Exception:
                pass
            try:
                del controller
            except Exception:
                pass
            try:
                del ctx
            except Exception:
                pass
            _lm_pp_uno_gc(True)
        if last_err is not None and (LIBRE_MACROS_DEBUG or _LM_PP_SORT_DEBUG):
            try:
                msg = str(last_err)
            except Exception:
                msg = "<unprintable>"
            try:
                print("[dispatch] slot=%s FAIL: %s" % (slot, msg))
            except Exception:
                pass


_LM_PP_INDENT_ALIGN_MAP = {
    "left": "left",
    "лево": "left",
    "левый": "left",
    "center": "center",
    "centre": "center",
    "центр": "center",
    "центрировать": "center",
    "right": "right",
    "право": "right",
    "правый": "right",
}

_LM_PP_INDENT_STEP = 353
# Один шаг кнопки «Увеличить отступ» в Calc (1/100 мм), как в диспетчере объектов.


def _lm_pp_parse_indent_steps_token(token):
    """Число шагов отступа (целое): 1, 2, 3…"""
    try:
        return max(0, int(token))
    except (ValueError, TypeError):
        return None


def _lm_pp_parse_indent_param(*extra_args):
    """
    Rest для «отступ» из колонки C.

    Формат (через запятую):
        [left|center|right] [, шаги] [, селекторы столбцов…]

    Селектор столбца: индекс с 1 (как в Calc), буква (A, C), подстрока заголовка.
    Без селекторов — все столбцы data_range.

    Возвращает:
        (align, steps, col_parts) — align: left|center|right; steps: int; col_parts: list[str].
    """
    align = "left"
    steps = 1
    col_parts = []

    tokens = []
    if extra_args:
        for arg in extra_args:
            for part in str(arg).split(","):
                p = part.strip()
                if p == "" or p.lower() == "отступ":
                    continue
                tokens.append(p)

    if len(tokens) == 0:
        return align, steps, col_parts

    idx = 0
    align_key = tokens[0].lower()
    if align_key in _LM_PP_INDENT_ALIGN_MAP:
        align = _LM_PP_INDENT_ALIGN_MAP[align_key]
        idx = 1

    if idx < len(tokens):
        parsed_steps = _lm_pp_parse_indent_steps_token(tokens[idx])
        if parsed_steps is not None:
            steps = parsed_steps
            idx = idx + 1

    while idx < len(tokens):
        col_parts.append(tokens[idx])
        idx = idx + 1

    return align, steps, col_parts


def _lm_pp_resolve_indent_columns(col_parts, sheet, header_row_range, sc, ec):
    """
    Список 0-based индексов столбцов в пределах data_range для «отступ».

    col_parts пуст — все столбцы sc…ec.
    Каждый селектор разбирается отдельно (A и Сумма — не одна строка «A,Сумма»).
    """
    if col_parts is None or len(col_parts) == 0:
        cols = []
        c = sc
        while c <= ec:
            cols.append(c)
            c = c + 1
        return cols

    result = []
    for part in col_parts:
        parsed = _lm_pp_parse_columns_param(
            str(part).strip(), sheet, header_row_range, sc, ec
        )
        if parsed is None:
            continue
        for col_idx in parsed:
            if sc <= col_idx <= ec and col_idx not in result:
                result.append(col_idx)
    result.sort()
    return result


def _lm_pp_apply_align_on_range(cell_range, align):
    """Выравнивание через HoriJustify (надёжно отображается без dispatch)."""
    from com.sun.star.table.CellHoriJustify import (
        CENTER as HORI_CENTER,
        LEFT as HORI_LEFT,
        RIGHT as HORI_RIGHT,
    )

    justify_map = {
        "left": HORI_LEFT,
        "center": HORI_CENTER,
        "right": HORI_RIGHT,
    }
    if cell_range is None:
        return False
    try:
        cell_range.HoriJustify = justify_map.get(align, HORI_LEFT)
        return True
    except Exception:
        return False


def _lm_pp_apply_indent_on_range(cell_range, steps):
    """
    Отступ ParaIndent на диапазон: ровно steps × 353 (1/100 мм), без накопления.
    """
    if cell_range is None or steps <= 0:
        return False
    try:
        cell_range.ParaIndent = int(steps) * _LM_PP_INDENT_STEP
        return True
    except Exception:
        return False


def _lm_pp_apply_indent_column(col_range, align, steps):
    """Один столбец: выравнивание + отступ (только свойства ячеек, один раз)."""
    if col_range is None:
        return
    _lm_pp_apply_align_on_range(col_range, align)
    if steps > 0:
        _lm_pp_apply_indent_on_range(col_range, steps)


def lm_pp_range_left_align_data(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Выравнивание содержимого data_range по левому краю.

    Параметры:
        data_range — XCellRange данных; header_row_range не меняется

    JSON колонки C: [{"v":1,"fn":"левое_выравнивание"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «левое_выравнивание».

    Когда вызывать:
        Если стандартное центрирование заголовка не должно распространяться на данные.

    Что меняет:
        data_range.HoriJustify = LEFT.
    """
    if _lm_pp_sheet_list_skip("левое_выравнивание", extra_args, sheet):
        return
    from com.sun.star.table.CellHoriJustify import LEFT as HORI_LEFT
    data_range.HoriJustify = HORI_LEFT


def lm_pp_range_increase_indent(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Выравнивание и отступ по столбцам data_range (HoriJustify + ParaIndent).

    Параметры:
        doc — SpreadsheetDocument (для журнала при пропуске)
        sheet — XSpreadsheet
        data_range — XCellRange данных
        header_row_range — для поиска столбцов по подстроке заголовка
        *extra_args — rest из колонки C

    JSON колонки C: [{"v":1,"fn":"отступ","h_align":"left","steps":1,"columns":["Сумма"]}].

    По умолчанию: left, 1, все столбцы. Один шаг = ParaIndent 353.

    Примеры:
        «right,1,Сумма» — вправо, 1 шаг, столбцы с «Сумма» в заголовке
        «center,2» — по центру, 2 шага (706), все столбцы
        «left,1,3,5» — столбцы 3 и 5 (нумерация с 1)

    Карта: «отступ».

    Что меняет:
        HoriJustify и ParaIndent = шаги × 353 на каждый целевой столбец (один раз).
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    if _lm_pp_param_skip_sheet("отступ", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet("отступ", extra_args, doc, sheet)
    if block is not None:
        align = block.get("h_align") or "left"
        try:
            steps = max(0, int(block.get("steps", 1)))
        except (TypeError, ValueError):
            steps = 1
        col_parts = [str(c) for c in (block.get("columns") or [])]
    else:
        align, steps, col_parts = _lm_pp_parse_indent_param(*extra_args)
    target_cols = _lm_pp_resolve_indent_columns(
        col_parts, sheet, header_row_range, sc, ec
    )
    if len(target_cols) == 0:
        spec = ", ".join(col_parts) if col_parts else "(все)"
        _lm_log_postprocess(
            doc,
            sheet.Name if sheet is not None else "",
            "диапазон",
            "отступ",
            "пропуск",
            "столбцы не найдены в диапазоне %d–%d: %s"
            % (sc + 1, ec + 1, spec),
        )
        return

    n = 0
    for col_idx in target_cols:
        col_range = sheet.getCellRangeByPosition(col_idx, sr, col_idx, er)
        _lm_pp_apply_indent_column(col_range, align, steps)
        n = n + 1
        if n % 10 == 0:
            _lm_ui_yield(counter=n)

    _lm_ui_yield(force=True)


def _lm_pp_font_extra_args(extra_raw):
    s = str(extra_raw).strip() if extra_raw is not None else ""
    if s == "":
        return []
    if _lm_pp_payload_is_json(s):
        return _lm_pp_codec_text_args("шрифт", extra_raw)
    if ";" in s:
        return [s]
    return parse_pp_extra_args(extra_raw)


def _lm_pp_font_rest_text(extra_args):
    if extra_args is None or len(extra_args) == 0:
        return ""
    if len(extra_args) == 1:
        return str(extra_args[0]).strip()
    parts = []
    i = 0
    while i < len(extra_args) and i < 3:
        parts.append(str(extra_args[i]).strip())
        i = i + 1
    while len(parts) < 3:
        parts.append("")
    return ";".join(parts)


def _lm_pp_font_optional_token(token):
    s = str(token or "").strip()
    if s == "" or s == "-":
        return None
    return s


def _lm_pp_font_optional_size(token):
    s = _lm_pp_font_optional_token(token)
    if s is None:
        return None
    try:
        return float(s.replace(",", "."))
    except (TypeError, ValueError):
        return None


def _lm_pp_parse_font_rest(rest_text):
    """
    Колонка C: «имя;размер_данных;размер_заголовка» или через запятую.
    «-» или пустое поле — не менять этот атрибут.
    """
    raw = str(rest_text or "").strip()
    if raw == "":
        return {"name": None, "data_size": None, "header_size": None}
    if ";" in raw:
        parts = [p.strip() for p in raw.split(";")]
    else:
        parts = [p.strip() for p in raw.split(",")]
    while len(parts) < 3:
        parts.append("")
    return {
        "name": _lm_pp_font_optional_token(parts[0]),
        "data_size": _lm_pp_font_optional_size(parts[1]),
        "header_size": _lm_pp_font_optional_size(parts[2]),
    }


def lm_pp_range_set_font(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Задать имя и/или размер шрифта для данных и заголовка.

    Параметры:
        data_range, header_row_range — XCellRange
        *extra_args — rest из колонки C (см. _lm_pp_parse_font_rest)

    JSON колонки C: [{"v":1,"fn":"шрифт","name":"PT Sans","size_data":10,"size_header":11}].

    Карта: «шрифт».

    Когда вызывать:
        После копирования, если нужен единый шрифт отличный от стиля книги.

    Что меняет:
        CharFontName / CharHeight на data_range и header_row_range.
    """
    _unused = (doc, sheet)
    if _lm_pp_param_skip_sheet("шрифт", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet("шрифт", extra_args, doc, sheet)
    if block is not None:
        spec = _lm_pp_font_spec_from_block(block)
    elif extra_args is None or len(extra_args) == 0:
        try:
            data_range.CharFontName = _LM_PP_FONT_NAME
            data_range.CharHeight = float(_LM_PP_DEFAULT_FONT_SIZE)
        except Exception:
            pass
        try:
            header_row_range.CharFontName = _LM_PP_FONT_NAME
            header_row_range.CharHeight = float(_LM_PP_DEFAULT_FONT_SIZE)
        except Exception:
            pass
        return
    spec = _lm_pp_parse_font_rest(_lm_pp_font_rest_text(extra_args))
    if spec["name"] is not None:
        try:
            data_range.CharFontName = spec["name"]
        except Exception:
            pass
        try:
            header_row_range.CharFontName = spec["name"]
        except Exception:
            pass
    if spec["data_size"] is not None:
        try:
            data_range.CharHeight = spec["data_size"]
        except Exception:
            pass
    if spec["header_size"] is not None:
        try:
            header_row_range.CharHeight = spec["header_size"]
        except Exception:
            pass


def _lm_pp_apply_format_money_block(doc, sheet, data_range, header_row_range, block):
    """Один JSON-блок «формат_деньги» → конверт/NumberFormat целевых столбцов."""
    if not isinstance(block, dict):
        block = {}
    sheet_name = ""
    try:
        sheet_name = str(getattr(sheet, "Name", "") or "")
    except Exception:
        sheet_name = ""
    convert_existing = bool(block.get("convert_existing"))
    fmt = str(block.get("format") or "").strip()
    if fmt == "":
        fmt = _LM_PP_FORMAT_MONEY_FMT
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    titles = _lm_pp_header_titles(sheet, header_row_range)
    fmt_key = _lm_pp_number_format_key(doc, fmt)
    columns = block.get("columns")
    indices = []
    if columns not in (None, "", []):
        indices = _lm_pp_resolve_format_columns_indices(columns, titles, sc, ec)
    if not indices:
        markers = block.get("markers")
        if markers:
            if isinstance(markers, (list, tuple)):
                marker_list = [str(m).strip() for m in markers if str(m).strip()]
            else:
                try:
                    from libre_macros_param_codec import split_column_name_tokens

                    marker_list = [
                        p for p in split_column_name_tokens(str(markers)) if str(p).strip()
                    ]
                except Exception:
                    marker_list = [
                        p.strip()
                        for p in str(markers).replace(";", ",").split(",")
                        if p.strip()
                    ]
        else:
            marker_list = ["сумм", "руб", "amount", "price", "стоим"]
        c = sc
        while c <= ec:
            idx = c - sc
            title = titles[idx] if idx < len(titles) else ""
            mi = 0
            while mi < len(marker_list):
                if _lm_pp_match_header_marker(marker_list[mi], title):
                    indices.append(c)
                    break
                mi = mi + 1
            c = c + 1
    if not indices:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "формат_деньги",
            "пропуск",
            "столбцы не найдены",
        )
        return
    for col_idx in indices:
        if convert_existing:
            _lm_pp_convert_column_cells_chunked(
                sheet, col_idx, sr, er, "number", status_prefix="диапазон"
            )
        _lm_pp_apply_column_number_format_chunked(
            sheet, col_idx, sr, er, fmt_key, status_prefix="диапазон"
        )


def lm_pp_range_format_money_columns(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Числовой формат денежных/числовых столбцов.

    Параметры:
        doc — SpreadsheetDocument (NumberFormats)
        sheet — XSpreadsheet
        data_range — XCellRange данных (по столбцам)
        header_row_range — XCellRange заголовков

    JSON колонки C:
      [{"v":1,"fn":"формат_деньги","sheet":"Отчет_Q2","columns":[5,6],
        "format":"# ##0,00","convert_existing":true}]
      format Excel-стиля (#,##0.00) автоматически нормализуется в Calc ru (# ##0,00).
      Без columns — по markers или маркерам по умолчанию (сумм/руб/…).
      На лист применяются только блоки с совпадающим sheet (или без sheet).

    Карта: «формат_деньги».
    """
    rest_text = _lm_pp_extra_args_rest_text(extra_args)
    if rest_text.strip() == "":
        blocks = [{}]
    else:
        if _lm_pp_param_skip_sheet("формат_деньги", extra_args, doc, sheet):
            return
        blocks = _lm_pp_blocks_for_sheet("формат_деньги", extra_args, doc, sheet)
        if len(blocks) == 0:
            return
    for block in blocks:
        _lm_pp_apply_format_money_block(
            doc, sheet, data_range, header_row_range, block
        )


def lm_pp_range_hide_service_columns(doc, sheet, data_range, header_row_range):
    """
    Назначение:
        Скрыть служебные столбцы A (#Путь) и B (#ДатаВремяФайла).

    Параметры:
        sheet — XSpreadsheet; doc, data_range, header_row_range не используются

    JSON колонки C: [{"v":1,"fn":"скрыть_служебные"}] без доп. полей или пусто.

    Карта: «скрыть_служебные».

    Когда вызывать:
        После сбора, если путь и mtime не нужны пользователю на экране.

    Что меняет:
        Columns[0].IsVisible и Columns[1].IsVisible = False.
    """
    layout = _lm_active_service_layout()
    for svc_col in (layout["path_col"], layout["datetime_col"]):
        if svc_col >= 0:
            sheet.getColumns().getByIndex(svc_col).IsVisible = False




def lm_pp_range_init_path_stripe_state(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Сбросить глобальное состояние полос по смене #Путь для данного листа.

    Параметры:
        sheet — XSpreadsheet (используется sheet.Name)

    JSON колонки C: [{"v":1,"fn":"левое_выравнивание"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «сброс_стрипов».

    Когда вызывать:
        Обязательно range-шагом ПЕРЕД «полосы_по_пути» в POSTPROCESS_ROW.

    Что меняет:
        Удаляет ключи _LM_PP_PATH_STRIPE_STATE для имени листа (не ячейки).
    """
    if _lm_pp_sheet_list_skip("сброс_стрипов", extra_args, sheet):
        return
    name = sheet.Name
    if name in _LM_PP_PATH_STRIPE_STATE:
        del _LM_PP_PATH_STRIPE_STATE[name]


# =============================================================================
# «Зебра» — общие хелперы (range и row)
# =============================================================================

def _lm_pp_is_zebra_postprocess_name(fn_name):
    key = str(fn_name or "").strip().lower()
    return key == "зебра_диапазон"


def _lm_pp_is_sort_postprocess_name(fn_name):
    key = str(fn_name or "").strip().lower()
    return key == "сортировка"


def _lm_pp_sort_extra_args(extra_raw):
    """Колонка C для «сортировка» — одна строка (запятая — часть синтаксиса ключей)."""
    return _lm_pp_codec_text_args("сортировка", extra_raw)


def _lm_pp_zebra_extra_args(extra_raw):
    """Колонка C для «зебра» — одна строка; JSON → канонический текст."""
    return _lm_pp_codec_text_args("зебра_диапазон", extra_raw)


_LM_PP_ZEBRA_COLOR_ALIASES = {
    "белый": (255, 255, 255),
    "white": (255, 255, 255),
    "черный": (0, 0, 0),
    "чёрный": (0, 0, 0),
    "black": (0, 0, 0),
    "серый": (242, 242, 242),
    "серый_светлый": (242, 242, 242),
    "светлыйсерый": (242, 242, 242),
    "светлосерый": (242, 242, 242),
    "светлый_серый": (242, 242, 242),
    "lightgray": (242, 242, 242),
    "lightgrey": (242, 242, 242),
    "grey": (128, 128, 128),
    "gray": (128, 128, 128),
    "синий": (68, 114, 196),
    "blue": (68, 114, 196),
    "темныйсиний": (31, 78, 121),
    "тёмныйсиний": (31, 78, 121),
    "темный_синий": (31, 78, 121),
    "darkblue": (31, 78, 121),
    "голубой": (217, 234, 247),
    "lightblue": (217, 234, 247),
    "красный": (255, 199, 206),
    "светло_красный": (255, 199, 206),
    "светло-красный": (255, 199, 206),
    "lightred": (255, 199, 206),
    "red": (255, 0, 0),
    "желтый": (255, 235, 156),
    "жёлтый": (255, 235, 156),
    "yellow": (255, 255, 0),
    "зеленый": (198, 239, 206),
    "зелёный": (198, 239, 206),
    "green": (0, 128, 0),
    "оранжевый": (255, 220, 177),
    "orange": (255, 165, 0),
    "фиолетовый": (228, 208, 255),
    "purple": (128, 0, 128),
    "бежевый": (250, 248, 239),
    "кремовый": (255, 248, 220),
    "navy": (0, 0, 128),
    # #1F4E79 — тёмно-синий фон заголовка таблицы Excel
    "excel_header": (31, 78, 121),
}

# Роль в JSON «зебра_диапазон»: roles.header / even / odd
_LM_PP_ZEBRA_ROLE_KEYS = {
    "header": "header",
    "заголовок": "header",
    "even": "even",
    "четные": "even",
    "чётные": "even",
    "odd": "odd",
    "нечетные": "odd",
    "нечётные": "odd",
}

_LM_PP_ZEBRA_ROLE_SEGMENT_SPLIT_RE = re.compile(
    r"\s*[;,]\s*(?=(?:header|заголовок|even|odd|"
    r"четные|чётные|нечетные|нечётные)\s*:)",
    re.IGNORECASE,
)

_LM_PP_ZEBRA_REST_HINT = (
    "роль:заливка/шрифт через «;» или «,» — header/заголовок, even/четные (2,4,6…), odd/нечетные (1,3,5…)",
    "header:blue/white, четные:grey/white",
)

# Row-«зебра»: заголовок красится один раз на лист (ключ — имя листа).
_LM_PP_ZEBRA_HEADER_DONE = {}


def _lm_pp_parse_color_param(color_val, default):
    """
    Разобрать параметр цвета из строки.

    Поддерживает:
        - "R,G,B" — три числа через запятую
        - "0xRRGGBB" или "#RRGGBB" — hex
        - None — вернуть default
    """
    if color_val is None:
        return default

    s = str(color_val).strip()

    # Hex формат
    if s.startswith("0x") or s.startswith("#"):
        try:
            hex_str = s[2:] if s.startswith("0x") else s[1:]
            return int(hex_str, 16)
        except (ValueError, TypeError):
            return default

    # RGB формат через запятую
    parts = s.split(",")
    if len(parts) >= 3:
        try:
            r = int(parts[0].strip())
            g = int(parts[1].strip())
            b = int(parts[2].strip())
            return lo_color_rgb(r, g, b)
        except (ValueError, TypeError):
            return default

    # Пробуем как целое число напрямую
    try:
        return int(s)
    except (ValueError, TypeError):
        return default


def _lm_pp_zebra_resolve_color_token(token):
    if token is None:
        return None
    s = str(token).strip()
    if s == "":
        return None
    key = s.lower()
    if key in _LM_PP_ZEBRA_COLOR_ALIASES:
        rgb = _LM_PP_ZEBRA_COLOR_ALIASES[key]
        return lo_color_rgb(rgb[0], rgb[1], rgb[2])
    try:
        return lo_color_hex(s)
    except (TypeError, ValueError):
        pass
    return _lm_pp_parse_color_param(s, None)


def _lm_pp_zebra_parse_stripe_spec(spec_str):
    spec = str(spec_str or "").strip()
    if spec == "":
        return {"bg": None, "fg": None, "bg_clear": False}
    if "/" not in spec:
        bg = _lm_pp_zebra_resolve_color_token(spec)
        return {"bg": bg, "fg": None, "bg_clear": bg is None}
    fill_t, _, font_t = spec.partition("/")
    fill_t = fill_t.strip()
    font_t = font_t.strip()
    bg = _lm_pp_zebra_resolve_color_token(fill_t) if fill_t else None
    fg = _lm_pp_zebra_resolve_color_token(font_t) if font_t else None
    return {"bg": bg, "fg": fg, "bg_clear": False}


def _lm_pp_zebra_normalize_role_key(key_text):
    key = str(key_text or "").strip().lower().replace("ё", "е")
    return _LM_PP_ZEBRA_ROLE_KEYS.get(key)


def _lm_pp_zebra_default_even_style():
    return {
        "bg": lo_color_rgb(242, 242, 242),
        "fg": None,
        "bg_clear": False,
    }


def _lm_pp_zebra_split_role_segments(raw):
    """Члены «роль:заливка/шрифт» через «;» или «,» (не внутри RGB и не в заливка/шрифт)."""
    text = str(raw or "").strip()
    if text == "":
        return []
    parts = _LM_PP_ZEBRA_ROLE_SEGMENT_SPLIT_RE.split(text)
    return [p.strip() for p in parts if p.strip()]


def _lm_pp_zebra_parse_rest(rest_text):
    """
    Разобрать колонку C для «зебра».

    Формат: «роль:заливка/шрифт» через «;» или «,»
    (роли: header/заголовок, even/четные, odd/нечетные).

    Возвращает dict с ключами header, even, odd (значение — style dict или None).
    """
    raw = str(rest_text or "").strip()
    roles = {"header": None, "even": None, "odd": None}
    if raw == "":
        # Классическая зебра: 1-я, 3-я… строка данных (odd) — светло-серый фон
        roles["odd"] = _lm_pp_zebra_default_even_style()
        return roles

    for segment in _lm_pp_zebra_split_role_segments(raw):
        segment = segment.strip()
        if segment == "" or ":" not in segment:
            continue
        key_part, _, val_part = segment.partition(":")
        role = _lm_pp_zebra_normalize_role_key(key_part)
        if role is None:
            continue
        roles[role] = _lm_pp_zebra_parse_stripe_spec(val_part)

    has_data_role = roles["even"] is not None or roles["odd"] is not None
    if not has_data_role and roles["header"] is None:
        roles["even"] = _lm_pp_zebra_default_even_style()
    elif roles["even"] is None and has_data_role:
        roles["even"] = _lm_pp_zebra_default_even_style()
    return roles


def _lm_pp_zebra_data_row_rel_index(header_row_range, row_index):
    """0-based индекс строки внутри данных (0 = первая строка под заголовком)."""
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    return row_index - (h_sr + 1)


def _lm_pp_zebra_rel_is_even_role(rel):
    """
    even/четные — 2-я, 4-я, 6-я… строка данных (rel=1,3,5…).
    odd/нечетные — 1-я, 3-я, 5-я… (rel=0,2,4…).
    """
    if rel < 0:
        return False
    return rel % 2 == 1


def _lm_pp_zebra_data_row_is_even(header_row_range, row_index):
    return _lm_pp_zebra_rel_is_even_role(
        _lm_pp_zebra_data_row_rel_index(header_row_range, row_index)
    )


def _lm_pp_zebra_apply_header_row_style( sheet, header_row_range, header_style, header_row=0, end_col=-1 ):
    """
    Заголовок из «зебра»: header:заливка/шрифт — тот же механизм, что у строк данных.

    Пример: header:темныйсиний/белый → CellBackColor + CharColor2 (не Automatic).
    """
    _unused = (sheet, header_row, end_col)
    if header_row_range is None or header_style is None:
        return
    _lm_pp_zebra_apply_row_style(header_row_range, header_style)
    try:
        header_row_range.CharFontName = _LM_PP_FONT_NAME
    except Exception:
        pass
    try:
        header_row_range.CharWeight = _LM_FMT_CHAR_WEIGHT_BOLD
    except Exception:
        pass
    try:
        header_row_range.CharHeight = float(_LM_PP_DEFAULT_FONT_SIZE)
    except Exception:
        pass
    if sheet is not None and end_col >= 0:
        try:
            hr = int(header_row)
        except Exception:
            hr = 0
        _lm_apply_header_row_font_size(sheet, end_col, hr)


def _lm_pp_zebra_maybe_apply_header(sheet, header_row_range, header_style):
    if header_style is None or header_row_range is None:
        return
    name = sheet.Name
    if _LM_PP_ZEBRA_HEADER_DONE.get(name):
        return
    _lm_pp_zebra_apply_header_row_style(
        sheet, header_row_range, header_style
    )
    _LM_PP_ZEBRA_HEADER_DONE[name] = True


def _lm_pp_zebra_clear_header_done(sheet):
    name = sheet.Name
    if name in _LM_PP_ZEBRA_HEADER_DONE:
        del _LM_PP_ZEBRA_HEADER_DONE[name]


def _lm_pp_zebra_preapply_headers_from_fns(fns, sheet, header_row_range):
    """Заголовок из row-«зебра» — один раз до цикла по строкам (не ждать 1-ю строку данных)."""
    if sheet is None or header_row_range is None:
        return
    for fn in fns:
        rest = getattr(fn, "_lm_pp_zebra_rest", None)
        if rest is None or str(rest).strip() == "":
            continue
        styles = _lm_pp_zebra_parse_rest(rest)
        header_style = styles.get("header")
        if header_style is not None:
            _lm_pp_zebra_maybe_apply_header(
                sheet, header_row_range, header_style
            )


def _lm_pp_resolve_zebra_header_style_from_specs(*callback_specs):
    """Последний заданный header: из цепочек range/row «зебра» (колонка C)."""
    header_style = None
    si = 0
    while si < len(callback_specs):
        spec = callback_specs[si]
        if spec is not None:
            for fn in _lm_postprocess_callbacks(spec):
                rest = getattr(fn, "_lm_pp_zebra_rest", None)
                if rest is None or str(rest).strip() == "":
                    continue
                styles = _lm_pp_zebra_parse_rest(rest)
                hs = styles.get("header")
                if hs is not None:
                    header_style = hs
        si = si + 1
    return header_style


def _lm_pp_zebra_apply_row_style(row_range, style, odd_default_clear=False):
    """Оформление одной строки зебры: фон диапазона + цвет шрифта (как у данных)."""
    if style is None:
        if odd_default_clear:
            try:
                row_range.IsCellBackgroundTransparent = True
            except Exception:
                pass
        return
    if style.get("bg") is not None:
        row_range.CellBackColor = style["bg"]
    elif style.get("bg_clear") or odd_default_clear:
        try:
            row_range.IsCellBackgroundTransparent = True
        except Exception:
            pass
    if style.get("fg") is not None:
        _lm_set_char_color_safe(row_range, style["fg"])


def _lm_pp_zebra_rest_text(extra_args):
    return _lm_pp_join_extra_args(extra_args)


# =============================================================================
# Постобработка одной строки данных
# =============================================================================

def lm_pp_row_zebra_stripes( doc, sheet, data_row_range, header_row_range, row_index, *extra_args ):
    """
    Назначение:
        «Зебра»: чередование оформления строк данных по чётности row_index.

    Параметры:
        doc — SpreadsheetDocument (не используется)
        sheet — XSpreadsheet (не используется)
        data_row_range — XCellRange одной строки данных
        header_row_range — не используется
        row_index — int, 0-based; чётность — от 1-й строки данных (после заголовка)
        *extra_args — rest из колонки C (одна строка, см. _lm_pp_zebra_parse_rest)

    Карта MERGE_RESULT_POSTPROCESS_ROW_MAP: «зебра».

    Когда вызывать:
        В _MERGE_CODE_POSTPROCESS_ROW; несовместимо визуально с «полосы_по_пути» на том же листе.

    Что меняет:
        CellBackColor и CharColor заголовка (один раз) и data_row_range по чётности строки.
    """
    styles = _lm_pp_zebra_styles_for_sheet(extra_args, doc, sheet)
    _lm_pp_zebra_maybe_apply_header(
        sheet, header_row_range, styles.get("header")
    )
    if _lm_pp_zebra_data_row_is_even(header_row_range, row_index):
        _lm_pp_zebra_apply_row_style(data_row_range, styles.get("even"))
    else:
        odd_style = styles.get("odd")
        _lm_pp_zebra_apply_row_style(
            data_row_range,
            odd_style,
            odd_default_clear=odd_style is None,
        )


def lm_pp_range_zebra_even_rows( doc, sheet, data_range, header_row_range, *extra_args ):
    """
    Назначение:
        «Зебра» через range-постобработку: оформление всех строк данных за один проход.

    Параметры:
        doc — SpreadsheetDocument (не используется)
        sheet — XSpreadsheet
        data_range — XCellRange всех данных (без заголовка)
        header_row_range — не используется
        *extra_args — rest из колонки C (одна строка, см. _lm_pp_zebra_parse_rest)

    Карта MERGE_RESULT_POSTPROCESS_RANGE_MAP: «зебра_диапазон».

    Когда вызывать:
        После копирования данных; быстрее чем row-«зебра».

    Что меняет:
        CellBackColor и CharColor заголовка и каждой строки диапазона данных.
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    styles = _lm_pp_zebra_styles_for_sheet(extra_args, doc, sheet)
    _lm_pp_zebra_clear_header_done(sheet)
    _lm_pp_zebra_maybe_apply_header(
        sheet, header_row_range, styles.get("header")
    )

    row_count = 0
    r = sr
    while r <= er:
        if _lm_pp_zebra_rel_is_even_role(row_count):
            style = styles.get("even")
            odd_default = False
        else:
            style = styles.get("odd")
            odd_default = style is None
        try:
            row_range = sheet.getCellRangeByPosition(sc, r, ec, r)
            if row_range is not None:
                _lm_pp_zebra_apply_row_style(
                    row_range,
                    style,
                    odd_default_clear=odd_default,
                )
        except Exception:
            pass
        row_count = row_count + 1
        r = r + 1
        if row_count % 50 == 0:
            _lm_ui_yield(counter=row_count)
            _lm_ui_status_set(
                row_count,
                _lm_pp_status_text("зебра_диапазон", "строка %d/%d" % (r - sr, er - sr + 1)),
            )
    _lm_ui_yield(force=True)


def lm_pp_row_height_pad(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
    Задать высоту текущей строки данных в мм (см. _lm_pp_parse_row_height_rest).

    Колонка C — как у range «высота_строки»; при списке строк применяется только
    к строкам из списка (номера 1-based на листе).

    Карта: «высота_строки» (построчно).
    """
    _unused = (data_row_range, header_row_range)
    row_specs, height_mm = _lm_pp_parse_row_height_rest(extra_args)
    if not _lm_pp_row_height_target_matches(row_index, row_specs):
        return
    _lm_pp_apply_row_height_mm(sheet, row_index, height_mm)


def lm_pp_row_gray_service_columns(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
    Назначение:
        Бледный текст в ячейках A–B текущей строки данных.

    Параметры:
        sheet — XSpreadsheet
        data_row_range — XCellRange одной строки (координаты столбцов)
        row_index — не используется напрямую (строка из data_row_range)

    JSON колонки C: [{"v":1,"fn":"серые_служебные"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «серые_служебные».

    Когда вызывать:
        Вместе с «зебра» или «полосы_по_пути»; дублирует часть стандартного форматирования.

    Что меняет:
        CharColor служебных ячеек строки (индексы столбцов 0 и 1).
    """
    sc, sr, ec, er = _lm_pp_range_address(data_row_range)
    layout = _lm_active_service_layout()
    try:
        hr_sc, hr_sr, hr_ec, hr_er = _lm_pp_range_address(header_row_range)
        header_row = int(hr_sr)
    except Exception:
        header_row = 0
    svc_pairs = (
        (layout["path_col"], u"#Путь"),
        (layout["datetime_col"], u"#ДатаВремяФайла"),
    )
    for svc_col, svc_hdr in svc_pairs:
        if svc_col < 0 or svc_col < sc or svc_col > ec:
            continue
        try:
            hdr_text = unicode(
                sheet.getCellByPosition(svc_col, header_row).String or u""
            ).strip()
        except Exception:
            hdr_text = u""
        if hdr_text.casefold() != svc_hdr.casefold():
            continue
        _lm_style_service_data_cell(sheet.getCellByPosition(svc_col, sr))


def lm_pp_row_stripe_on_path_change(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
    Назначение:
        Чередовать оттенок фона блока строк при смене значения в колонке #Путь (A).

    Параметры:
        sheet — XSpreadsheet
        data_row_range — XCellRange одной строки
        header_row_range — не используется
        row_index — int, 0-based (чтение пути из A,row_index)

    JSON колонки C: [{"v":1,"fn":"полосы_по_пути"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «полосы_по_пути».

    Когда вызывать:
        После range «сброс_стрипов»; вместо «зебра».

    Что меняет:
        Фон data_row_range; состояние в _LM_PP_PATH_STRIPE_STATE (ключи :path, :stripe).
    """
    path = _lm_cell_path_value(sheet, row_index)
    name = sheet.Name
    prev = _LM_PP_PATH_STRIPE_STATE.get(name + ":path", None)
    # Смена пути — переключить чётность полосы (0/1)
    if prev is None or path != prev:
        stripe = _LM_PP_PATH_STRIPE_STATE.get(name + ":stripe", 0)
        stripe = 1 - stripe
        _LM_PP_PATH_STRIPE_STATE[name + ":stripe"] = stripe
        _LM_PP_PATH_STRIPE_STATE[name + ":path"] = path
    stripe = _LM_PP_PATH_STRIPE_STATE.get(name + ":stripe", 0)
    if stripe == 0:
        data_row_range.IsCellBackgroundTransparent = True
    else:
        data_row_range.CellBackColor = lo_color_rgb(235, 241, 250)


def lm_pp_row_highlight_by_header_value(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
    Назначение:
        Подсветить всю строку, если в «итоговом» столбце (по заголовку) ячейка не пуста.

    Параметры:
        sheet, data_row_range, header_row_range, row_index — см. сигнатуру row-колбэка

    JSON колонки C: [{"v":1,"fn":"подсветка_по_заголовку"}] — подстрока заголовка в коде (_LM_PP_ROW_HIGHLIGHT_HEADER).

    Карта: «подсветка_по_заголовку».

    Когда вызывать:
        Для выделения строк с заполненным итогом; после числовых колбэков.

    Что меняет:
        CellBackColor и/или CharColor всей data_row_range при первом совпадении столбца.
    """
    fg = _LM_PP_ROW_HIGHLIGHT_FG
    bg = _LM_PP_ROW_HIGHLIGHT_BG
    if fg is None and bg is None:
        bg = lo_color_rgb(255, 242, 204)
    sc, sr, ec, er = _lm_pp_range_address(data_row_range)
    titles = _lm_pp_header_titles(sheet, header_row_range)
    needle = str(_LM_PP_ROW_HIGHLIGHT_HEADER).lower()
    c = sc
    while c <= ec:
        idx = c - sc
        title = titles[idx].lower() if idx < len(titles) else ""
        if needle in title:
            val = cell_text(sheet.getCellByPosition(c, sr))
            if val != "":
                if bg is not None:
                    data_row_range.CellBackColor = bg
                if fg is not None:
                    _lm_set_char_color_safe(data_row_range, fg)
                return
        c = c + 1


def lm_pp_row_bold_if_path_contains(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
    Назначение:
        Жирный шрифт строки, если #Путь содержит подстроку _LM_PP_ROW_BOLD_PATH_SUBSTRING.

    Параметры:
        row_index — строка для чтения A; data_row_range — куда пишется CharWeight

    JSON колонки C: [{"v":1,"fn":"жирный_по_пути"}] — маркер в коде, без полей в C.

    Карта: «жирный_по_пути».

    Когда вызывать:
        После изменения _LM_PP_ROW_BOLD_PATH_SUBSTRING в модуле; пустая строка — no-op.

    Что меняет:
        data_row_range.CharWeight = 150 при совпадении пути.
    """
    substring = _LM_PP_ROW_BOLD_PATH_SUBSTRING
    if substring == "":
        return
    path = _lm_cell_path_value(sheet, row_index).lower()
    if substring.lower() in path:
        data_row_range.CharWeight = _LM_FMT_CHAR_WEIGHT_BOLD


def lm_pp_row_convert_numeric_strings(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
    Назначение:
        Преобразовать числа, записанные как текст (в т.ч. с апострофом), в cell.Value + формат.

    Параметры:
        doc — SpreadsheetDocument (NumberFormats)
        sheet — XSpreadsheet
        data_row_range — XCellRange одной строки
        header_row_range — используется для разрешения имён колонок (если extra_args содержит имена)
        row_index — не используется (строка из data_row_range)

    JSON колонки C: [{"v":1,"fn":"преобразовать_числа","columns":["Сумма","Количество"]}].
        - Пусто или не указано → обрабатывать все колонки (как раньше)

    Примеры:
        «преобразовать_числа» — все колонки
        «преобразовать_числа,0,2,5» — только колонки 0, 2 и 5 (индексы от 0)
        «преобразовать_числа,Сумма,Количество» — колонки с заголовками "Сумма" и "Количество"

    Карта: «преобразовать_числа».

    Когда вызывать:
        До «формат_деньги»; на каждой строке (дорого на больших таблицах).

    Что меняет:
        Тип и NumberFormat ячеек строки, похожих на число; даты/время пропускаются.
    """
    sc, sr, ec, er = _lm_pp_range_address(data_row_range)

    raw_joined = _lm_pp_join_extra_args(extra_args)
    if _lm_pp_payload_is_json(raw_joined):
        param = _lm_pp_codec_param_text_for_sheet(
            "преобразовать_числа", extra_args, doc, sheet
        )
        if param is None:
            return
        extra_args = parse_pp_extra_args(param) if param else ()

    # Определяем список колонок для обработки
    # extra_args содержит разобранные аргументы из колонки C (разделённые по запятой)
    target_cols = None
    if extra_args:
        # Объединяем все extra_args через запятую — это был один список в колонке C
        columns = ",".join(str(a) for a in extra_args)
        target_cols = _lm_pp_parse_columns_param(columns, sheet, header_row_range, sc, ec)

    # Для ru_RU: пробел как разделитель тысяч, запятая как разделитель дроби → "72 000,00"
    fmt_key = _lm_pp_number_format_key(doc, "# ##0,00")

    # Если target_cols=None — обрабатываем все колонки, иначе только указанные
    cols_to_process = target_cols if target_cols is not None else range(sc, ec + 1)

    for c in cols_to_process:
        # Проверяем что колонка в допустимых границах
        if c < sc or c > ec:
            continue
        cell = sheet.getCellByPosition(c, sr)
        try:
            s = str(cell.String) if cell.String is not None else ""
        except Exception:
            s = ""
        s = s.strip()
        if s == "":
            continue
        if s.startswith("'"):
            s = s[1:].strip()
        if s == "":
            continue
        # Похоже на дату/время — пропускаем
        if ":" in s:
            continue
        if s.count(".") == 2 and all(p.isdigit() for p in s.split(".")):
            continue
        if s.count("-") == 2 and all(p.isdigit() for p in s.split("-")):
            continue
        # Нормализация числа
        t = s.replace("\u00A0", " ").replace(" ", "")
        if "," in t and "." not in t:
            # десятичная запятая
            t = t.replace(",", ".")
        # Быстрый фильтр: допустимые символы
        ok = True
        for ch in t:
            if not (ch.isdigit() or ch in (".", "-", "+")):
                ok = False
                break
        if not ok:
            continue
        try:
            val = float(t)
        except Exception:
            continue
        # Запись в ячейку как числа
        try:
            cell.String = ""
        except Exception:
            pass
        cell.Value = val
        cell.NumberFormat = fmt_key


def _lm_pp_parse_columns_param(columns, sheet, header_row_range, sc, ec):
    """
    Разобрать параметр columns для lm_pp_row_convert_numeric_strings.

    Параметры:
        columns — строка или список: индексы или имена колонок
        sheet — XSpreadsheet
        header_row_range — XCellRange строки заголовков
        sc, ec — границы диапазона данных (start/end column)

    Возвращает:
        list[int] — список индексов колонок для обработки, или None если не удалось разобрать
    """
    if columns is None:
        return None

    # Если уже список — используем как есть
    if isinstance(columns, (list, tuple)):
        return list(columns)

    # Преобразуем в строку
    s = str(columns).strip()
    if not s:
        return None

    # Убираем квадратные скобки если есть
    if s.startswith("[") and s.endswith("]"):
        s = s[1:-1]
    elif s.startswith("(") and s.endswith(")"):
        s = s[1:-1]

    parts = [p.strip() for p in s.split(",")]
    if not parts:
        return None

    result = []
    header_names = None

    for part in parts:
        if not part:
            continue
        # Пробуем как число (индекс колонки, 1-based от пользователя → 0-based в коде)
        try:
            idx = int(part) - 1
            result.append(idx)
            continue
        except ValueError:
            pass

        # Буквы столбца Calc (A, AA…)
        if is_col_letters(part):
            result.append(col_letters_to_index(part))
            continue

        # Иначе это имя колонки — ищем в заголовках
        if header_names is None and header_row_range is not None:
            header_names = _lm_pp_header_titles(sheet, header_row_range)

        if header_names is not None:
            # Ищем в заголовках (без учёта регистра); все совпадения по подстроке
            part_lower = part.lower()
            for col_idx, hname in enumerate(header_names):
                abs_idx = sc + col_idx
                hlower = hname.lower()
                if hlower == part_lower or part_lower in hlower:
                    if abs_idx not in result:
                        result.append(abs_idx)

    return result if result else None


def _lm_pp_extra_args_rest_text(extra_args):
    if extra_args is None or len(extra_args) == 0:
        return ""
    if len(extra_args) == 1:
        return str(extra_args[0]).strip()
    return ",".join(str(a) for a in extra_args).strip()


def _lm_pp_looks_like_date_string(s):
    s = str(s or "").strip()
    if s == "":
        return False
    if ":" in s:
        return True
    if s.count(".") == 2:
        parts = s.split(".")
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            return True
    if s.count("-") == 2:
        parts = s.split("-")
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            return True
    return False


def _lm_pp_try_parse_number_string(s):
    s = str(s or "").strip()
    if s == "":
        return None
    if s.startswith("'"):
        s = s[1:].strip()
    if s == "" or _lm_pp_looks_like_date_string(s):
        return None
    t = s.replace("\u00A0", " ").replace(" ", "")
    if "," in t and "." not in t:
        t = t.replace(",", ".")
    ok = True
    for ch in t:
        if not (ch.isdigit() or ch in (".", "-", "+")):
            ok = False
            break
    if not ok:
        return None
    try:
        return float(t)
    except Exception:
        return None


def _lm_pp_resolve_parse_formats_raw(block=None):
    """
    Строка parse_formats для convert текста → дата.
    Приоритет: поле блока → MERGE_XML_DATETIME_FORMATS (bound direct options).
    """
    if isinstance(block, dict):
        raw = str(block.get("parse_formats") or "").strip()
        if raw:
            return raw
    try:
        from libre_macros_direct_lib import _direct_options

        raw = str((_direct_options or {}).get("datetime_formats", "") or "").strip()
        if raw:
            return raw
    except Exception:
        pass
    return ""


def _lm_pp_try_parse_date_string(s, formats_raw=None):
    s = str(s or "").strip()
    if s.startswith("'"):
        s = s[1:].strip()
    if s == "":
        return None
    import datetime

    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            base = datetime.datetime(1899, 12, 30)
            delta = dt - base
            return float(delta.days) + float(delta.seconds) / 86400.0
        except Exception:
            pass
    if ":" in s:
        for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
            try:
                dt = datetime.datetime.strptime(s, fmt)
                base = datetime.datetime(1899, 12, 30)
                delta = dt - base
                return float(delta.days) + float(delta.seconds) / 86400.0
            except Exception:
                pass
    try:
        from libre_macros_direct_lib import (
            _direct_datetime_formats_from_raw,
            _direct_try_parse_datetime_text,
        )

        raw = str(formats_raw if formats_raw is not None else "").strip()
        if raw == "" and formats_raw is None:
            raw = _lm_pp_resolve_parse_formats_raw(None)
        if raw:
            allowed = _direct_datetime_formats_from_raw(raw)
            parsed = _direct_try_parse_datetime_text(
                s, allowed_dt=allowed, formats_raw=raw
            )
        else:
            parsed = _direct_try_parse_datetime_text(s)
        if parsed is None:
            return None
        if isinstance(parsed, datetime.datetime):
            dt = parsed
        else:
            dt = datetime.datetime(parsed.year, parsed.month, parsed.day)
        base = datetime.datetime(1899, 12, 30)
        delta = dt - base
        return float(delta.days) + float(delta.seconds) / 86400.0
    except Exception:
        return None


def _lm_pp_format_convert_mode(format_string):
    f = str(format_string or "").strip().lower()
    if f in ("общий", "general", ""):
        return "general"
    if "%" in f:
        return "number"
    for token in ("дд", "dd", "мм", "mm", "гг", "yy", "yyyy", "чч", "hh", "ми", "ss"):
        if token in f:
            return "date"
    if any(ch in f for ch in ("#", "0")):
        return "number"
    return "general"


def _lm_pp_cell_content_is_text(cell):
    try:
        ctype = int(cell.getPropertyValue("CellContentType"))
        return ctype == 2
    except Exception:
        try:
            s = str(cell.String) if cell.String is not None else ""
            return s.strip() != ""
        except Exception:
            return False


def _lm_pp_convert_cell_text_for_format(cell, mode, formats_raw=None):
    if not _lm_pp_cell_content_is_text(cell):
        return "skip"
    try:
        s = str(cell.String) if cell.String is not None else ""
    except Exception:
        s = ""
    s = s.strip()
    if s == "":
        return "skip"
    val = None
    use_mode = mode
    if use_mode == "general":
        val = _lm_pp_try_parse_number_string(s)
        if val is None:
            val = _lm_pp_try_parse_date_string(s, formats_raw=formats_raw)
            use_mode = "date" if val is not None else "general"
    elif use_mode == "date":
        val = _lm_pp_try_parse_date_string(s, formats_raw=formats_raw)
    else:
        val = _lm_pp_try_parse_number_string(s)
    if val is None:
        return "skip"
    try:
        cell.String = ""
    except Exception:
        pass
    try:
        cell.Value = val
    except Exception:
        return "error"
    return "converted"


def _lm_pp_convert_column_cells_chunked( sheet, col_index, sr, er, mode, status_prefix="конверт", formats_raw=None ):
    r = sr
    n = 0
    converted = skipped = errors = 0
    total = er - sr + 1
    while r <= er:
        r_end = r + _LM_PP_COPY_ROW_CHUNK - 1
        if r_end > er:
            r_end = er
        ri = r
        while ri <= r_end:
            cell = sheet.getCellByPosition(col_index, ri)
            res = _lm_pp_convert_cell_text_for_format(
                cell, mode, formats_raw=formats_raw
            )
            if res == "converted":
                converted = converted + 1
            elif res == "error":
                errors = errors + 1
            else:
                skipped = skipped + 1
            ri = ri + 1
        n = n + (r_end - r + 1)
        if _lm_ui_should_update_status(n, is_last=(r_end >= er)):
            _lm_ui_status_set(
                n,
                _lm_pp_status_text(
                    status_prefix,
                    "конверт столбец %d: %d/%d (ok %d, проп %d, ош %d)"
                    % (col_index + 1, r_end + 1, er + 1, converted, skipped, errors),
                ),
            )
        _lm_ui_yield(counter=n)
        r = r_end + 1
    return converted, skipped, errors


def _lm_pp_sheet_list_block_for_sheet(fn_key, extra_args, sheet):
    text = _lm_pp_extra_args_rest_text(extra_args)
    if text == "":
        return {}
    from libre_macros_param_codec import param_decode

    blocks = param_decode(fn_key, text)
    if len(blocks) == 0:
        return {}
    block = blocks[0]
    sheets = block.get("sheets")
    if sheets:
        try:
            sheet_name = str(getattr(sheet, "Name", "") or "")
        except Exception:
            sheet_name = ""
        if not _lm_pp_sheet_name_matches_target(sheet_name, set(sheets)):
            return None
    return block


def _lm_pp_format_columns_blocks_for_sheet(extra_args, sheet, doc=None):
    return _lm_pp_blocks_for_sheet("формат_столбцы", extra_args, doc, sheet)


def _lm_pp_resolve_format_columns_indices(column_selector, titles, sc, ec):
    """Селектор столбца(ов) формат_* → список 0-based индексов в диапазоне."""
    if column_selector is None:
        return []
    if column_selector == u"all":
        return list(range(sc, ec + 1))
    specs = []
    if isinstance(column_selector, (list, tuple)):
        specs = list(column_selector)
    else:
        s = str(column_selector).strip()
        if s == "":
            return []
        low = s.casefold()
        if low in ("all", "все", "*"):
            return list(range(sc, ec + 1))
        try:
            from libre_macros_param_codec import split_column_name_tokens

            parts = split_column_name_tokens(s)
            if parts:
                specs = parts
            elif "," in s:
                specs = [p.strip() for p in s.split(",") if p.strip()]
            else:
                specs = [s]
        except Exception:
            if "," in s:
                specs = [p.strip() for p in s.split(",") if p.strip()]
            else:
                specs = [s]
    out = []
    seen = set()
    for spec in specs:
        if isinstance(spec, int):
            col_idx = int(spec) - 1
            if sc <= col_idx <= ec and col_idx not in seen:
                seen.add(col_idx)
                out.append(col_idx)
            continue
        s = str(spec or "").strip()
        if s == "":
            continue
        low = s.casefold()
        if low in ("all", "все", "*"):
            c = sc
            while c <= ec:
                if c not in seen:
                    seen.add(c)
                    out.append(c)
                c += 1
            continue
        for col_idx in _lm_pp_resolve_format_column_indices(s, titles, sc, ec):
            if col_idx not in seen:
                seen.add(col_idx)
                out.append(col_idx)
    out.sort()
    return out


def _lm_pp_resolve_format_column_index(column_selector, titles, sc, ec):
    """Один токен → первый подходящий индекс (совместимость)."""
    found = _lm_pp_resolve_format_column_indices(column_selector, titles, sc, ec)
    if found:
        return found[0]
    return None


def _lm_pp_resolve_format_column_indices(column_selector, titles, sc, ec):
    """
    Один токен → список индексов столбцов.
    'name' = exact; 'name*' = шаблон; без кавычек = '*part*' / fnmatch;
    без кавычек число/буква = один столбец.
    """
    sel = column_selector
    if isinstance(sel, int):
        col_idx = int(sel) - 1
        if sc <= col_idx <= ec:
            return [col_idx]
        return []
    s_raw = str(sel or "").strip()
    if s_raw == "":
        return []
    quoted = False
    try:
        from libre_macros_param_codec import unwrap_column_name_token

        s, quoted = unwrap_column_name_token(s_raw)
    except Exception:
        s = s_raw
        quoted = False
    s = str(s or "").strip()
    if s == "":
        return []
    if not quoted:
        try:
            col_idx = int(s) - 1
            if sc <= col_idx <= ec:
                return [col_idx]
        except ValueError:
            pass
        if is_col_letters(s):
            col_idx = col_letters_to_index(s)
            if sc <= col_idx <= ec:
                return [col_idx]
            return []
    try:
        from libre_macros_param_codec import column_header_matches_token

        match_fn = column_header_matches_token
    except Exception:
        match_fn = None
    out = []
    c = sc
    while c <= ec:
        idx = c - sc
        title = titles[idx] if idx < len(titles) else ""
        ok = False
        if match_fn is not None:
            ok = match_fn(title, s_raw)
        else:
            ok = str(s).lower() in str(title or "").lower()
        if ok:
            out.append(c)
        c = c + 1
    return out


def _lm_pp_match_header_marker(marker, title):
    """Маркер заголовка vs title: единая логика column_header_matches_token."""
    try:
        from libre_macros_param_codec import column_header_matches_token

        return column_header_matches_token(title, marker)
    except Exception:
        pass
    raw = str(marker or "").strip()
    title_l = str(title or "").lower()
    if raw == "" or title_l == "":
        return False
    return raw.lower() in title_l


def _lm_pp_format_columns_row_range(rule, header_row, data_start_row, data_end_row):
    if bool(rule.get("header_only")):
        return header_row, header_row
    if bool(rule.get("include_header")):
        return header_row, data_end_row
    return data_start_row, data_end_row


_LM_PP_FORMAT_HEADER_COLS = {}


def _lm_pp_track_format_header_cols(sheet, col_indices):
    """Столбцы заголовка, оформленные format_столбцы (include_header/header_only)."""
    try:
        sheet_name = str(getattr(sheet, "Name", "") or "").strip()
    except Exception:
        sheet_name = ""
    if sheet_name == "":
        return
    cols = _LM_PP_FORMAT_HEADER_COLS.setdefault(sheet_name, set())
    for ci in col_indices or []:
        try:
            cols.add(int(ci))
        except (TypeError, ValueError):
            pass


def lm_pp_take_format_header_cols():
    """Снять накопленные столбцы заголовка format_столбцы (для collect_workbooks)."""
    global _LM_PP_FORMAT_HEADER_COLS
    buf = {k: set(v) for k, v in (_LM_PP_FORMAT_HEADER_COLS or {}).items()}
    _LM_PP_FORMAT_HEADER_COLS = {}
    return buf


def _lm_pp_format_columns_auto_font_color(bg_rgb):
    r = (int(bg_rgb) >> 16) & 0xFF
    g = (int(bg_rgb) >> 8) & 0xFF
    b = int(bg_rgb) & 0xFF
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255.0
    if luminance > 0.55:
        return lo_color_rgb(0, 0, 0)
    return lo_color_rgb(255, 255, 255)


def _lm_pp_format_columns_rule_style(rule):
    """Стиль из правила формат_столбцы → dict или None."""
    if not isinstance(rule, dict):
        return None
    fill_tok = rule.get("fill") or rule.get("bg")
    fill_rgb = (
        _lm_pp_zebra_resolve_color_token(fill_tok) if fill_tok not in (None, u"") else None
    )
    font_rgb = None
    if rule.get("font_auto") is True or rule.get("font_color_auto") is True:
        if fill_rgb is not None:
            font_rgb = _lm_pp_format_columns_auto_font_color(fill_rgb)
    else:
        font_tok = rule.get("font") or rule.get("font_color")
        if font_tok not in (None, u""):
            font_rgb = _lm_pp_zebra_resolve_color_token(font_tok)
    h_raw = str(rule.get("h_align") or rule.get("align_h") or u"").strip().lower()
    v_raw = str(rule.get("v_align") or rule.get("align_v") or u"").strip().lower()
    h_align = h_raw if h_raw in (u"left", u"center", u"right") else None
    v_align = v_raw if v_raw in (u"top", u"center", u"bottom") else None
    bw_tok = rule.get("border_width") or rule.get("border")
    bc_tok = rule.get("border_color")
    border_width = None
    border_color = None
    if bw_tok not in (None, u"") or bc_tok not in (None, u""):
        border_width = (
            _lm_pp_grid_parse_width_token(bw_tok or u"тонкая")
            or _LM_PP_GRID_WIDTH_ALIASES[u"тонкая"]
        )
        border_color = (
            _lm_pp_zebra_resolve_color_token(bc_tok or u"серый")
            or _LM_PP_GRID_DEFAULT_COLOR
        )
    if (
        fill_rgb is None
        and font_rgb is None
        and h_align is None
        and v_align is None
        and border_width is None
    ):
        return None
    return {
        u"fill": fill_rgb,
        u"font": font_rgb,
        u"h_align": h_align,
        u"v_align": v_align,
        u"border_width": border_width,
        u"border_color": border_color,
    }


def _lm_pp_apply_format_columns_cell_style(cell, style):
    if cell is None or not style:
        return
    fill_rgb = style.get(u"fill")
    if fill_rgb is not None:
        try:
            cell.IsCellBackgroundTransparent = False
        except Exception:
            pass
        try:
            cell.CellBackColor = int(fill_rgb)
        except Exception:
            pass
    font_rgb = style.get(u"font")
    if font_rgb is not None:
        _lm_set_char_color_cell(cell, font_rgb)
    h_align = style.get(u"h_align")
    v_align = style.get(u"v_align")
    if h_align or v_align:
        hori_val = _lm_pp_resolve_hori_justify(h_align or u"left")
        _lm_pp_set_cell_alignment(cell, hori_val, v_align or u"center")
    border_width = style.get(u"border_width")
    if border_width is not None:
        border_color = style.get(u"border_color") or _LM_PP_GRID_DEFAULT_COLOR
        for side in (u"top", u"bottom", u"left", u"right"):
            _lm_pp_set_cell_side_border(cell, side, border_color, border_width)


def _lm_pp_apply_format_columns_style_chunked( sheet, col_index, sr, er, style, status_prefix=u"формат" ):
    r = sr
    n = 0
    while r <= er:
        r_end = r + _LM_PP_COPY_ROW_CHUNK - 1
        if r_end > er:
            r_end = er
        rr = r
        while rr <= r_end:
            try:
                cell = sheet.getCellByPosition(col_index, rr)
                _lm_pp_apply_format_columns_cell_style(cell, style)
            except Exception:
                pass
            rr += 1
        n = n + (r_end - r + 1)
        if _lm_ui_should_update_status(n, is_last=(r_end >= er)):
            _lm_ui_status_set(
                n,
                _lm_pp_status_text(
                    status_prefix,
                    u"стиль столбца %d, строки %d/%d"
                    % (col_index + 1, r_end + 1, er + 1),
                ),
            )
        _lm_ui_yield(counter=n)
        r = r_end + 1


def _lm_pp_apply_format_column_rule( doc, sheet, data_range, header_row_range, rule, convert_existing=False, formats_raw=None ):
    column_selector = rule.get("column")
    format_string = str(rule.get("format") or "").strip()
    if format_string.lower() in ("общий", "general"):
        format_string = ""
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    _, hr, _, _ = _lm_pp_range_address(header_row_range)
    titles = _lm_pp_header_titles(sheet, header_row_range)
    indices = _lm_pp_resolve_format_columns_indices(column_selector, titles, sc, ec)
    # Столбец мог появиться после «применить_формулу» правее устаревшего ec.
    if not indices and sheet is not None:
        try:
            used_ec, used_er = get_sheet_used_bounds(sheet)
        except Exception:
            used_ec, used_er = ec, er
        if int(used_ec) > int(ec) or int(used_er) > int(er):
            ec2 = max(int(ec), int(used_ec))
            er2 = max(int(er), int(used_er))
            try:
                hdr2 = sheet.getCellRangeByPosition(0, hr, ec2, hr)
                titles2 = _lm_pp_header_titles(sheet, hdr2)
                indices2 = _lm_pp_resolve_format_columns_indices(
                    column_selector, titles2, 0, ec2
                )
                if indices2:
                    indices = indices2
                    sc, ec, er = 0, ec2, er2
                    titles = titles2
            except Exception:
                pass
    if not indices:
        return
    row0, row1 = _lm_pp_format_columns_row_range(rule, hr, sr, er)
    style = _lm_pp_format_columns_rule_style(rule)
    has_fmt = format_string != u""
    if not has_fmt and style is None:
        return
    if (
        bool(rule.get("header_only")) or bool(rule.get("include_header"))
    ) and row0 <= hr <= row1:
        _lm_pp_track_format_header_cols(sheet, indices)
    for col_idx in indices:
        if convert_existing and has_fmt:
            mode = _lm_pp_format_convert_mode(format_string)
            _lm_pp_convert_column_cells_chunked(
                sheet,
                col_idx,
                row0,
                row1,
                mode,
                status_prefix=u"формат",
                formats_raw=formats_raw,
            )
        if has_fmt:
            fmt_key = _lm_pp_number_format_key(doc, format_string)
            _lm_pp_apply_column_number_format_chunked(
                sheet, col_idx, row0, row1, fmt_key, status_prefix=u"формат"
            )
        if style is not None:
            _lm_pp_apply_format_columns_style_chunked(
                sheet, col_idx, row0, row1, style, status_prefix=u"формат"
            )


# =============================================================================
# Дополнительные функции постобработки (расширенный набор)
# =============================================================================

# --- Расширенные функции (константы _LM_PP_DATE_FORMAT / _NUMBER_DECIMALS — выше) ---


def _lm_pp_apply_format_date_block(doc, sheet, data_range, header_row_range, block):
    """Один JSON-блок «формат_даты» → конверт/NumberFormat целевых столбцов."""
    if not isinstance(block, dict):
        block = {}
    convert_existing = bool(block.get("convert_existing"))
    markers = block.get("markers")
    columns = block.get("columns")
    fmt = _LM_PP_DATE_FORMAT
    raw_fmt = str(block.get("format") or "").strip()
    if raw_fmt:
        fmt = raw_fmt
    if markers is None:
        markers = []
    if len(markers) == 0:
        date_markers_use = ("дата", "date", "время", "time")
    elif isinstance(markers, (list, tuple)):
        date_markers_use = tuple(str(m).strip() for m in markers if str(m).strip())
    else:
        try:
            from libre_macros_param_codec import split_column_name_tokens

            date_markers_use = tuple(
                p for p in split_column_name_tokens(str(markers)) if str(p).strip()
            )
        except Exception:
            date_markers_use = tuple(
                p.strip()
                for p in str(markers).replace(";", ",").split(",")
                if p.strip()
            )
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    titles = _lm_pp_header_titles(sheet, header_row_range)
    fmt_key = _lm_pp_number_format_key(doc, fmt)
    indices = []
    if columns not in (None, "", []):
        indices = _lm_pp_resolve_format_columns_indices(columns, titles, sc, ec)
    if not indices:
        c = sc
        while c <= ec:
            idx = c - sc
            title = titles[idx] if idx < len(titles) else ""
            for marker in date_markers_use:
                if _lm_pp_match_header_marker(marker, title):
                    indices.append(c)
                    break
            c = c + 1
    for col_idx in indices:
        if convert_existing:
            parse_raw = _lm_pp_resolve_parse_formats_raw(block)
            _lm_pp_convert_column_cells_chunked(
                sheet,
                col_idx,
                sr,
                er,
                "date",
                status_prefix="диапазон",
                formats_raw=parse_raw if parse_raw else None,
            )
        _lm_pp_apply_column_number_format_chunked(
            sheet, col_idx, sr, er, fmt_key, status_prefix="диапазон"
        )


def lm_pp_range_format_date_columns(doc, sheet, data_range, header_row_range, *date_markers):
    """
    Назначение:
        Применить формат даты к столбцам по markers/columns.

    Параметры:
        doc — SpreadsheetDocument
        sheet — XSpreadsheet
        data_range — XCellRange данных
        header_row_range — XCellRange заголовков
        *date_markers — rest из колонки C (маркеры, листы, `!конверт` / JSON)

    JSON колонки C:
      [{"v":1,"fn":"формат_даты","sheet":"Лист","format":"DD.MM.YYYY",
        "markers":["дата"],"columns":[2],"convert_existing":false,
        "parse_formats":"FIRST_DAY.MMMM.YEAR"}]
      На лист применяются только блоки с совпадающим sheet (или без sheet).
      parse_formats — шаблоны разбора текста (как MERGE_XML_DATETIME_FORMATS),
      в т.ч. month-only: FIRST_DAY/LAST_DAY + MMMM + YEAR/NEXT_YEAR.

    Карта: «формат_даты».
    """
    rest_text = _lm_pp_extra_args_rest_text(date_markers)
    if rest_text.strip() == "":
        blocks = [{}]
    elif not _lm_pp_payload_is_json(rest_text):
        return
    else:
        if _lm_pp_param_skip_sheet("формат_даты", date_markers, doc, sheet):
            return
        blocks = _lm_pp_blocks_for_sheet("формат_даты", date_markers, doc, sheet)
        if len(blocks) == 0:
            return
    for block in blocks:
        _lm_pp_apply_format_date_block(
            doc, sheet, data_range, header_row_range, block
        )


def lm_pp_range_format_columns(doc, sheet, data_range, header_row_range, *rest_args):
    """
    Универсальное форматирование столбцов (карта: «формат_столбцы»).

    C — JSON или текст §6.1: блоки по листам, правила `столбец,формат` через `;`,
    опция `!конверт` / `convert_existing` (§6.1.1).
    """
    sheet_name = ""
    try:
        sheet_name = str(getattr(sheet, "Name", "") or "")
    except Exception:
        sheet_name = ""
    if _lm_pp_param_skip_sheet("формат_столбцы", rest_args, doc, sheet):
        return
    blocks = _lm_pp_format_columns_blocks_for_sheet(rest_args, sheet, doc)
    if len(blocks) == 0:
        return
    for block in blocks:
        convert_existing = bool(block.get("convert_existing"))
        parse_raw = _lm_pp_resolve_parse_formats_raw(block)
        for rule in block.get("rules") or []:
            _lm_pp_apply_format_column_rule(
                doc,
                sheet,
                data_range,
                header_row_range,
                rule,
                convert_existing=convert_existing,
                formats_raw=parse_raw if parse_raw else None,
            )


def lm_pp_range_highlight_by_threshold(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Подсветить ячейки столбцов по порогам (карта: «подсветка_по_порогу» / «подкрасить_пороги»).

    JSON: columns, threshold_min/max (включительно) либо legacy threshold (>);
    color (пусто — заливку не менять), font_color / font_color_auto, bold, italic;
    style_columns — столбцы подсветки; whole_row — вся используемая ширина листа;
    иначе стиль только на столбцах порога. Перед подсветкой — сброс заливки/шрифта/жирности/курсива.
    """
    if _lm_pp_param_skip_sheet("подсветка_по_порогу", extra_args, doc, sheet):
        return
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    titles = _lm_pp_header_titles(sheet, header_row_range)
    block = _lm_pp_param_block_for_sheet("подсветка_по_порогу", extra_args, doc, sheet)
    if block is not None:
        spec = _lm_pp_threshold_params_from_block(block)
    else:
        marker, threshold, highlight_color = _lm_pp_parse_threshold_highlight_rest(extra_args)
        spec = {
            "columns": [str(marker)],
            "threshold_min": None,
            "threshold_max": None,
            "legacy_gt": float(threshold),
            "fill_color": highlight_color,
            "font_color": None,
            "font_color_auto": True,
            "bold": False,
            "italic": False,
            "whole_row": False,
            "style_columns": [],
        }

    needles = []
    for col_tok in spec.get("columns") or []:
        n = str(col_tok).strip().lower()
        if n:
            needles.append(n)
    if not needles:
        return

    target_cols = _lm_pp_threshold_resolve_marker_columns(sc, ec, titles, needles)

    if not target_cols:
        return

    ec_for_paint = int(ec)
    if bool(spec.get("whole_row")):
        try:
            used_ec, _used_er = get_sheet_used_bounds(sheet)
            if int(used_ec) > ec_for_paint:
                ec_for_paint = int(used_ec)
        except Exception:
            pass
    _lm_pp_threshold_apply_batch(
        doc, sheet, target_cols, sc, sr, er, ec_for_paint, titles, spec
    )


_LM_PP_DEFAULT_THRESHOLD_HIGHLIGHT_MARKER = "итог"
_LM_PP_DEFAULT_THRESHOLD_HIGHLIGHT_VALUE = 1000.0
_LM_PP_PARAM_SEPARATORS = ("->", "/", "=")
_LM_PP_THRESHOLD_HIGHLIGHT_SEPARATORS = _LM_PP_PARAM_SEPARATORS


def _lm_pp_default_threshold_highlight_color():
    return lo_color_rgb(255, 199, 206)


def _lm_pp_try_parse_threshold_number(text):
    s = str(text or "").strip()
    if s == "":
        return None
    try:
        return float(s.replace(",", "."))
    except (ValueError, TypeError):
        return None


def _lm_pp_split_pp_param_tokens(raw):
    text = str(raw or "").strip()
    if text == "":
        return []
    norm = text
    used_struct_sep = False
    i = 0
    while i < len(_LM_PP_PARAM_SEPARATORS):
        sep = _LM_PP_PARAM_SEPARATORS[i]
        if sep in norm:
            used_struct_sep = True
            norm = norm.replace(sep, "|")
        i = i + 1
    tokens = [t.strip() for t in norm.split("|") if t.strip()]
    if not used_struct_sep and len(tokens) <= 1 and "," in text:
        tokens = [t.strip() for t in text.split(",") if t.strip()]
    return tokens


def _lm_pp_split_threshold_highlight_tokens(raw):
    return _lm_pp_split_pp_param_tokens(raw)


def _lm_pp_try_consume_rgb_token(tokens, idx):
    """Съесть три подряд 0..255 как R,G,B; иначе None."""
    if idx + 2 >= len(tokens):
        return None, idx
    nums = []
    j = idx
    while j < idx + 3:
        n = _lm_pp_try_parse_threshold_number(tokens[j])
        if n is None or n < 0 or n > 255 or int(n) != n:
            return None, idx
        nums.append(int(n))
        j = j + 1
    return "%d,%d,%d" % (nums[0], nums[1], nums[2]), idx + 3


def _lm_pp_consume_color_token(tokens, idx):
    if idx >= len(tokens):
        return "", idx
    rgb_text, next_idx = _lm_pp_try_consume_rgb_token(tokens, idx)
    if rgb_text is not None:
        return rgb_text, next_idx
    return tokens[idx], idx + 1


def _lm_pp_color_int_to_rgb(color_int):
    c = int(color_int)
    return (c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF


def _lm_pp_interpolate_color(color_min, color_max, norm):
    t = max(0.0, min(1.0, float(norm)))
    r0, g0, b0 = _lm_pp_color_int_to_rgb(color_min)
    r1, g1, b1 = _lm_pp_color_int_to_rgb(color_max)
    return lo_color_rgb(
        int(round(r0 + (r1 - r0) * t)),
        int(round(g0 + (g1 - g0) * t)),
        int(round(b0 + (b1 - b0) * t)),
    )


def _lm_pp_default_color_scale_min():
    return lo_color_rgb(0, 255, 0)


def _lm_pp_default_color_scale_max():
    return lo_color_rgb(255, 0, 0)


def _lm_pp_parse_color_scale_rest(extra_args):
    """
    Колонка C для «градиент»: маркер столбца, цвет min, цвет max.
    Разделители: , / -> =; цвета — имя, #hex, R,G,B.
    """
    raw = ""
    if extra_args is not None and len(extra_args) > 0:
        if len(extra_args) == 1:
            raw = str(extra_args[0]).strip()
        else:
            raw = ",".join(str(a) for a in extra_args).strip()

    marker = "сумм"
    min_color = _lm_pp_default_color_scale_min()
    max_color = _lm_pp_default_color_scale_max()

    tokens = _lm_pp_split_pp_param_tokens(raw)
    if len(tokens) == 0:
        return marker, min_color, max_color

    idx = 0
    if idx < len(tokens):
        marker = tokens[idx]
        idx = idx + 1

    if idx < len(tokens):
        min_token, idx = _lm_pp_consume_color_token(tokens, idx)
        resolved = _lm_pp_zebra_resolve_color_token(min_token)
        if resolved is not None:
            min_color = resolved

    if idx < len(tokens):
        max_token, idx = _lm_pp_consume_color_token(tokens, idx)
        if idx < len(tokens):
            max_token = max_token + "," + ",".join(tokens[idx:])
        resolved = _lm_pp_zebra_resolve_color_token(max_token)
        if resolved is not None:
            max_color = resolved

    return marker, min_color, max_color


def _lm_pp_color_scale_extra_args(extra_raw):
    """Колонка C для «градиент» — JSON как есть или канонический текст."""
    s = str(extra_raw).strip() if extra_raw is not None else ""
    if s == "":
        return []
    if s.startswith("[") or s.startswith("{"):
        return [s]
    return _lm_pp_codec_text_args("градиент", extra_raw)


def _lm_pp_gradient_dbg(where, note=""):
    """Отладка «градиент» в stdout."""
    if not LIBRE_MACROS_DEBUG:
        return
    try:
        parts = ["[градиент]", str(where)]
        if note not in ("", None):
            parts.append(str(note))
        print(" ".join(parts))
    except Exception:
        pass


def _lm_pp_gradient_log(doc, sheet_name, note):
    """Отладка «градиент» в журнал Сбор_книг_лог (видно в LO)."""
    msg = str(note or "")
    _lm_pp_gradient_dbg("journal", msg)
    _lm_log_postprocess(doc, sheet_name or "", "диапазон", "градиент", "дебаг", msg)


def _lm_pp_read_bool_field(block, key, default=False):
    if not isinstance(block, dict):
        return default
    raw = block.get(key)
    if isinstance(raw, bool):
        return raw
    if raw is None:
        return default
    try:
        return bool(lm_parse_bool_param(raw, default=default))
    except (ValueError, TypeError):
        return default


def _lm_pp_color_scale_block_params(block):
    """Параметры градиента из блока кодека."""
    _lm_pp_gradient_dbg(
        "block_params",
        "raw block keys=%s whole_row=%r sort=%r sort_dir=%r"
        % (
            list(block.keys()) if isinstance(block, dict) else type(block),
            block.get("whole_row") if isinstance(block, dict) else None,
            block.get("sort") if isinstance(block, dict) else None,
            block.get("sort_dir") if isinstance(block, dict) else None,
        ),
    )
    marker = str(block.get("marker") or "сумм").strip() or "сумм"
    min_token = str(block.get("color_min") or "зеленый").strip() or "зеленый"
    max_token = str(block.get("color_max") or "красный").strip() or "красный"
    min_color = _lm_pp_zebra_resolve_color_token(min_token)
    if min_color is None:
        min_color = _lm_pp_default_color_scale_min()
    max_color = _lm_pp_zebra_resolve_color_token(max_token)
    if max_color is None:
        max_color = _lm_pp_default_color_scale_max()
    whole_row = _lm_pp_read_bool_field(block, "whole_row", False)
    sort_enabled = _lm_pp_read_bool_field(block, "sort", False)
    sort_ascending = True
    if sort_enabled:
        sort_dir = str(block.get("sort_dir") or "по_возр").strip().lower()
        sort_ascending = sort_dir not in (
            "по_убыв",
            "убыв",
            "desc",
            "-",
            "убывание",
        )
    out = {
        "marker": marker,
        "min_color": min_color,
        "max_color": max_color,
        "whole_row": whole_row,
        "sort": sort_enabled,
        "sort_ascending": sort_ascending,
    }
    _lm_pp_gradient_dbg("block_params", "resolved=%r" % out)
    return out


def _lm_pp_gradient_marker_tokens(marker_raw):
    """Список подстрок-маркеров из поля marker (через , или ;)."""
    raw = str(marker_raw or "").strip()
    if raw == "":
        return ["сумм"]
    parts = []
    for chunk in raw.replace(";", ",").split(","):
        token = chunk.strip()
        if token:
            parts.append(token)
    return parts if parts else ["сумм"]


def _lm_pp_gradient_find_column(titles, sc, ec, marker_raw):
    """(col_index, matched_marker) или (None, None)."""
    markers = _lm_pp_gradient_marker_tokens(marker_raw)
    mi = 0
    while mi < len(markers):
        needle = markers[mi].lower()
        c = sc
        while c <= ec:
            idx = c - sc
            title = titles[idx].lower() if idx < len(titles) else ""
            if needle in title:
                return c, markers[mi]
            c = c + 1
        mi = mi + 1
    return None, None


def _lm_pp_color_scale_legacy_spec(extra_args):
    marker, min_color, max_color = _lm_pp_parse_color_scale_rest(extra_args)
    return {
        "marker": marker,
        "min_color": min_color,
        "max_color": max_color,
        "whole_row": False,
        "sort": False,
        "sort_ascending": True,
    }


def _lm_pp_color_scale_decode_blocks(extra_raw):
    """Блоки «градиент» из JSON или текста C."""
    s = str(extra_raw or "").strip()
    if s == "":
        return []
    from libre_macros_param_codec import param_decode

    return param_decode("градиент", s)


def _lm_pp_color_scale_spec_for_sheet(extra_args, doc, sheet):
    """
    (marker, min_color, max_color) для текущего листа или None (пропуск).
    JSON/текст кодека с полем sheet или legacy без листа.
    """
    sheet_name = ""
    try:
        sheet_name = str(getattr(sheet, "Name", "") or "")
    except Exception:
        pass
    if extra_args and isinstance(extra_args[0], dict):
        _lm_pp_gradient_dbg("spec_for_sheet", "extra_args[0]=dict")
        block = _lm_pp_pick_sheet_block_for_sheet([extra_args[0]], doc, sheet)
        if block is None:
            return None
        return _lm_pp_color_scale_block_params(block)
    raw = _lm_pp_extra_args_rest_text(extra_args)
    json_strict = _lm_pp_payload_is_json(raw) or (
        "градиент" in raw and ("whole_row" in raw or '"sort"' in raw or "sort_dir" in raw)
    )
    _lm_pp_gradient_dbg(
        "spec_for_sheet",
        "raw_head=%r json_strict=%s" % (raw[:200], json_strict),
    )
    if raw == "":
        return None
    blocks = []
    if raw != "":
        try:
            blocks = _lm_pp_color_scale_decode_blocks(raw)
        except Exception as err:
            _lm_pp_gradient_log(
                doc,
                sheet_name,
                "param_decode: %s | C=%r" % (err, raw[:240]),
            )
            if json_strict:
                return None
            blocks = []
    if len(blocks) > 0:
        block = _lm_pp_pick_sheet_block_for_sheet(blocks, doc, sheet)
        if block is None:
            _lm_pp_gradient_log(doc, sheet_name, "блок для листа не найден")
            return None
        return _lm_pp_color_scale_block_params(block)
    if json_strict:
        _lm_pp_gradient_log(doc, sheet_name, "JSON без блоков после decode")
        return None
    return _lm_pp_color_scale_legacy_spec(extra_args)


def lm_pp_color_scale_step_applies_to_sheet(extra_raw, sheet_name, doc=None, sheet=None):
    """Нужно ли выполнять «градиент» на данном листе."""
    s = str(extra_raw or "").strip()
    _lm_pp_gradient_dbg(
        "step_applies",
        "sheet=%r extra_head=%r" % (sheet_name, s[:160]),
    )
    if s == "":
        return False
    blocks = _lm_pp_color_scale_decode_blocks(s)
    _lm_pp_gradient_dbg("step_applies", "blocks=%d" % len(blocks))
    if len(blocks) == 0:
        return True
    has_named = False
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        bi = bi + 1
        sheet_ref = str(block.get("sheet") or "").strip()
        if sheet_ref == "":
            continue
        has_named = True
        names = _lm_pp_parse_sheet_names_from_spec(sheet_ref)
        if doc is not None and sheet is not None:
            if _lm_pp_sheet_ref_list_matches_sheet(doc, sheet, names):
                return True
        elif _lm_pp_sheet_name_matches_target(sheet_name, names):
            return True
    if has_named:
        return False
    return True


def _lm_pp_is_color_scale_postprocess_name(fn_name):
    return str(fn_name or "").strip().lower() == "градиент"


def _lm_pp_parse_threshold_highlight_rest(extra_args):
    block = None
    if extra_args and len(extra_args) == 1 and isinstance(extra_args[0], dict):
        block = extra_args[0]
    else:
        blocks = _lm_pp_param_decode("подсветка_по_порогу", extra_args)
        if len(blocks) == 1:
            block = blocks[0]
    if block is not None:
        # Новый формат возвращает dict — оставляем совместимость тройки для legacy-пути
        # в lm_pp_range_highlight_by_threshold (блок уже обработан выше).
        # Здесь только текст → (marker, threshold, color).
        if isinstance(block, dict) and (
            block.get("threshold_min") is not None
            or block.get("threshold_max") is not None
            or block.get("columns")
            or block.get("font_color")
            or block.get("bold")
            or block.get("italic")
        ):
            # Вернём через промежуточный вызов: caller already prefers block path
            spec = _lm_pp_threshold_params_from_block(block)
            marker = (spec.get("columns") or ["итог"])[0]
            thr = spec.get("legacy_gt")
            if thr is None:
                thr = spec.get("threshold_min")
            if thr is None:
                thr = _LM_PP_DEFAULT_THRESHOLD_HIGHLIGHT_VALUE
            return marker, float(thr), spec.get("fill_color")
        return (
            str(block.get("marker") or "итог").strip() or "итог",
            float(block.get("threshold", 1000.0)),
            _lm_pp_zebra_resolve_color_token(str(block.get("color") or ""))
            or _lm_pp_default_threshold_highlight_color(),
        )
    raw = _lm_pp_join_extra_args(extra_args).strip()

    marker = _LM_PP_DEFAULT_THRESHOLD_HIGHLIGHT_MARKER
    threshold = _LM_PP_DEFAULT_THRESHOLD_HIGHLIGHT_VALUE
    color = _lm_pp_default_threshold_highlight_color()

    tokens = _lm_pp_split_threshold_highlight_tokens(raw)
    if len(tokens) == 0:
        return marker, threshold, color

    idx = 0
    if idx < len(tokens):
        num = _lm_pp_try_parse_threshold_number(tokens[idx])
        if num is None:
            marker = tokens[idx]
            idx = idx + 1

    if idx < len(tokens):
        num = _lm_pp_try_parse_threshold_number(tokens[idx])
        if num is not None:
            threshold = num
            idx = idx + 1

    if idx < len(tokens):
        color_token = ",".join(tokens[idx:])
        resolved = _lm_pp_zebra_resolve_color_token(color_token)
        if resolved is not None:
            color = resolved

    return marker, threshold, color


def _lm_pp_threshold_highlight_extra_args(extra_raw):
    """Колонка C для «подсветка_по_порогу» — одна строка (запятая может быть в RGB)."""
    return _lm_pp_codec_text_args("подсветка_по_порогу", extra_raw)


def _lm_pp_is_threshold_highlight_postprocess_name(fn_name):
    key = str(fn_name or "").strip().lower()
    return key in (
        "подсветка_по_порогу",
        "условное_форматирование",
        "подкрасить_пороги",
    )


# Устаревшее имя функции (обратная совместимость Python-кода).
lm_pp_range_conditional_format_simple = lm_pp_range_highlight_by_threshold


def lm_pp_range_set_column_widths(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Задать ширину столбцов data_range в миллиметрах (см. _lm_pp_parse_column_width_rest).

    Колонка C (JSON):
      columns — "all" / Все, либо список 1-based / букв / имён (в '…' — exact);
      width_mm — мм; 0 — скрыть столбец(ы).

    Карта: «ширина_столбцов».
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    columns = sheet.getColumns()
    col_specs, width_mm = _lm_pp_parse_column_width_rest(extra_args)

    def _apply_col(c):
        if c < 0 or c > ec:
            return
        col = columns.getByIndex(c)
        if float(width_mm) == 0.0:
            col.IsVisible = False
            return
        col.IsVisible = True
        col.OptimalWidth = False
        col.Width = int(round(float(width_mm) * 100))

    if not col_specs:
        c = sc
        while c <= ec:
            _apply_col(c)
            c = c + 1
    else:
        i = 0
        while i < len(col_specs):
            spec = col_specs[i]
            c0 = None
            if isinstance(spec, bool):
                pass
            elif isinstance(spec, int):
                c0 = int(spec) - 1
            elif isinstance(spec, float):
                c0 = int(spec) - 1
            else:
                s = str(spec).strip()
                if s.casefold() in ("all", "все", "*"):
                    c = sc
                    while c <= ec:
                        _apply_col(c)
                        c = c + 1
                    return
                if (not s.startswith("'")) and s.isdigit():
                    c0 = int(s) - 1
                else:
                    c0 = _lm_pp_resolve_column_token(
                        spec, sheet, header_row_range, sc, ec
                    )
            if c0 is not None:
                _apply_col(c0)
            i = i + 1


_LM_PP_DEFAULT_COLUMN_WIDTH_MM = 30.0
_LM_PP_DIM_MM_SEPARATORS = ("->", "/", "=")


def _lm_pp_split_dim_mm_pair(raw):
    """Разделить «номера / значение_мм» по /, -> или =."""
    text = str(raw or "").strip()
    if text == "":
        return "", "", False
    i = 0
    while i < len(_LM_PP_DIM_MM_SEPARATORS):
        sep = _LM_PP_DIM_MM_SEPARATORS[i]
        if sep in text:
            left, right = text.split(sep, 1)
            return left.strip(), right.strip(), True
        i = i + 1
    return text, "", False


def _lm_pp_parse_dim_mm_rest(extra_args, default_mm):
    """
    Общий разбор C для «ширина_столбцов» и «высота_строки»:
      пусто — все элементы диапазона, default_mm;
      «N» — все элементы, N мм;
      «1,2,3 / N» (или ->, =) — элементы 1-based;
      «0» — скрыть (столбец или строка).
    """
    raw = ",".join(str(a) for a in extra_args).strip()
    if len(extra_args) == 1:
        raw = str(extra_args[0]).strip()
    default_mm = float(default_mm)
    if raw == "":
        return None, default_mm

    left, right, has_sep = _lm_pp_split_dim_mm_pair(raw)

    size_mm = default_mm
    if right != "":
        try:
            size_mm = _lm_pp_parse_mm_number(right)
        except ValueError:
            size_mm = default_mm
    elif not has_sep:
        try:
            size_mm = _lm_pp_parse_mm_number(raw)
            return None, size_mm
        except ValueError:
            left = raw

    indices = []
    if left != "":
        for part in left.split(","):
            part = part.strip()
            if part == "":
                continue
            try:
                indices.append(int(part))
            except (ValueError, TypeError):
                pass
    if len(indices) == 0:
        return None, size_mm
    return indices, size_mm


def _lm_pp_parse_column_width_rest(extra_args):
    block = None
    if extra_args and len(extra_args) == 1 and isinstance(extra_args[0], dict):
        block = extra_args[0]
    else:
        block = _lm_pp_param_block_for_sheet(
            "ширина_столбцов", extra_args, None, None
        )
        if block is None:
            blocks = _lm_pp_param_decode("ширина_столбцов", extra_args)
            if len(blocks) == 1:
                block = blocks[0]
    if block is not None:
        return _lm_pp_dim_mm_from_block(
            block, "columns", _LM_PP_DEFAULT_COLUMN_WIDTH_MM
        )
    return _lm_pp_parse_dim_mm_rest(extra_args, _LM_PP_DEFAULT_COLUMN_WIDTH_MM)


def _lm_pp_parse_row_height_rest(extra_args):
    block = None
    if extra_args and len(extra_args) == 1 and isinstance(extra_args[0], dict):
        block = extra_args[0]
    else:
        blocks = _lm_pp_param_decode("высота_строки", extra_args)
        if len(blocks) == 1:
            block = blocks[0]
    if block is not None:
        return _lm_pp_dim_mm_from_block(
            block, "rows", _LM_PP_DEFAULT_ROW_HEIGHT_MM
        )
    return _lm_pp_parse_dim_mm_rest(extra_args, _LM_PP_DEFAULT_ROW_HEIGHT_MM)


def _lm_pp_column_width_extra_args(extra_raw):
    """Колонка C для «ширина_столбцов» — одна строка (запятая — часть синтаксиса столбцов)."""
    return _lm_pp_codec_text_args("ширина_столбцов", extra_raw)


def _lm_pp_is_column_width_postprocess_name(fn_name):
    return str(fn_name or "").strip().lower() == "ширина_столбцов"


def lm_pp_range_freeze_header(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Закрепить строки выше первой строки данных (freeze panes).

    Параметры:
        sheet — XSpreadsheet
        data_range — XCellRange (sr = первая строка данных)

    JSON колонки C: [{"v":1,"fn":"левое_выравнивание"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «закрепить_заголовок».

    Когда вызывать:
        После форматирования заголовка; до прокрутки пользователем.

    Что меняет:
        freezeAtPosition(0, sr) — закреплены строки 0…sr-1 (sr — первая строка данных).
    """
    if _lm_pp_sheet_list_skip("закрепить_заголовок", extra_args, sheet):
        return
    # Надёжнее вычислять строку заголовка по header_row_range:
    # в некоторых режимах прямой обработки границы pp_ctx/data_range могут
    # отличаться, и freezeAtPosition визуально "не срабатывает".
    header_row = None
    try:
        if header_row_range is not None:
            _sc_h, sr_h, _ec_h, _er_h = _lm_pp_range_address(header_row_range)
            header_row = int(sr_h)
    except Exception:
        header_row = None
    if header_row is None:
        _sc, sr, _ec, _er = _lm_pp_range_address(data_range)
        header_row = int(sr) - 1
    if header_row < 0:
        header_row = 0
    try:
        _lm_dbg(
            "lm_pp_range_freeze_header sheet=%s header_row=%d"
            % (getattr(sheet, "Name", ""), int(header_row))
        )
    except Exception:
        pass
    defer_hook = globals().get("_LM_FREEZE_DEFER_HOOK")
    if defer_hook is not None:
        try:
            if defer_hook(doc, sheet, int(header_row)):
                return
        except Exception:
            pass
    _lm_pp_freeze_sheet_header(doc, sheet, int(header_row))


def lm_pp_range_add_filter(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Простой автофильтр на data_range (меню Data→AutoFilter), без AF_PP_*.

    Параметры:
        doc — SpreadsheetDocument
        sheet, data_range — XSpreadsheet, XCellRange

    JSON колонки C: [{"v":1,"fn":"автофильтр"}] или с "sheets":["Лист1"]; без sheets — все листы.

    Карта: «автофильтр».

    Состояние читается из UnnamedDatabaseRanges (не из AF_/смарт-таблиц).
    Повторный запуск не снимает фильтр: ensure через setByTable + AutoFilter=True.
    """
    block = _lm_pp_sheet_list_block_for_sheet("автофильтр", extra_args, sheet)
    if block is None:
        return
    action = str((block or {}).get("action") or "set").strip().casefold()
    if action not in ("set", "remove"):
        action = "set"
    smart_table = bool((block or {}).get("smart_table"))
    keep_smart_table = bool((block or {}).get("keep_smart_table", smart_table))
    table_style = str((block or {}).get("table_style") or "TableStyleLight1").strip()
    if table_style == "":
        table_style = "TableStyleLight1"
    # Диапазон фильтра = заголовок + данные (иначе Calc без шапки не показывает стрелки).
    sc_d, sr_d, ec_d, er_d = _lm_pp_range_address(data_range)
    try:
        sc_h, sr_h, ec_h, er_h = _lm_pp_range_address(header_row_range)
    except Exception:
        sc_h, sr_h, ec_h, er_h = sc_d, sr_d, ec_d, er_d
    sc = min(int(sc_d), int(sc_h))
    sr = min(int(sr_d), int(sr_h))
    ec = max(int(ec_d), int(ec_h))
    er = max(int(er_d), int(er_h))
    if er < sr or ec < sc:
        return
    try:
        sheet.getCellRangeByPosition(sc, sr, ec, er)
    except Exception:
        return

    try:
        # Снять именованные AF_ / AF_PP_ — иначе smart-table / двойные стрелки.
        # UnnamedDatabaseRanges (меню-фильтр) не трогаем.
        _lm_pp_remove_sheet_autofilters(doc, sheet)
        sn = str(getattr(sheet, "Name", "") or "").strip()
        if action == "remove":
            _lm_pp_disable_sheet_menu_autofilter(doc, sheet)
            _lm_pp_set_sheet_keep_smart_table(sn, False)
            try:
                import libre_macros_collect_cfg as _cw_cfg_pp
                done = set(
                    getattr(_cw_cfg_pp, "_MERGE_SHEETS_WITH_MENU_AUTOFILTER", None)
                    or set()
                )
                done.discard(sn)
                _cw_cfg_pp._MERGE_SHEETS_WITH_MENU_AUTOFILTER = done
            except Exception:
                pass
            _lm_log_postprocess(
                doc,
                sheet.Name,
                "автофильтр",
                "меню",
                "ok",
                "снят; строки %d..%d, столбцы %d..%d"
                % (sr + 1, er + 1, sc + 1, ec + 1),
            )
            return
        already = bool(_lm_pp_sheet_has_active_autofilter(doc, sheet))
        ok = False
        if smart_table:
            try:
                rng = sheet.getCellRangeByPosition(sc, sr, ec, er)
                ok = bool(_lm_pp_add_autofilter_range_with_style(
                    doc,
                    sheet,
                    rng,
                    style_name=table_style,
                    row_stripes=True,
                ))
            except Exception:
                ok = False
            if keep_smart_table:
                _lm_pp_set_sheet_keep_smart_table(sn, True)
            else:
                _lm_pp_set_sheet_keep_smart_table(sn, False)
        else:
            _lm_pp_set_sheet_keep_smart_table(sn, False)
            try:
                from libre_macros_split_lib import _split_activate_document
                try:
                    _split_activate_document(doc)
                    ctrl = doc.getCurrentController() if doc is not None else None
                    if ctrl is not None:
                        try:
                            ctrl.setActiveSheet(sheet)
                        except Exception:
                            pass
                except Exception:
                    pass
                ok = bool(_lm_pp_ensure_sheet_menu_autofilter(doc, sheet, sc, sr, ec, er))
            except Exception:
                ok = False
        if ok:
            try:
                if sn:
                    import libre_macros_collect_cfg as _cw_cfg_pp
                    done = set(
                        getattr(_cw_cfg_pp, "_MERGE_SHEETS_WITH_MENU_AUTOFILTER", None)
                        or set()
                    )
                    done.add(sn)
                    _cw_cfg_pp._MERGE_SHEETS_WITH_MENU_AUTOFILTER = done
                    if smart_table:
                        # Стиль/рамка smart-таблицы в AO/LO часто видны только после reopen.
                        _lm_pp_remember_smart_table_style(sn, table_style)
            except Exception:
                pass
            note = "уже был" if already else "включён"
            _lm_log_postprocess(
                doc,
                sheet.Name,
                "автофильтр",
                ("smart" if smart_table else "меню"),
                "ok",
                ("%s%s; строки %d..%d, столбцы %d..%d")
                % (
                    note,
                    (
                        "; стиль=%s keep_smart=%s; reopen в конце"
                        % (table_style, str(keep_smart_table))
                        if smart_table
                        else ""
                    ),
                    sr + 1,
                    er + 1,
                    sc + 1,
                    ec + 1,
                ),
            )
        else:
            _lm_log_postprocess(
                doc,
                sheet.Name,
                "автофильтр",
                ("smart" if smart_table else "меню"),
                "ошибка",
                (
                    "не удалось включить %s-автофильтр; строки %d..%d, столбцы %d..%d"
                    % (("smart" if smart_table else "меню"), sr + 1, er + 1, sc + 1, ec + 1)
                ),
            )
    except Exception as err:
        _lm_log_postprocess(
            doc,
            sheet.Name,
            "автофильтр",
            "—",
            "ошибка",
            str(err),
        )


def lm_pp_range_delete_sheets(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Доступность «удаление_листов» в «Постобработка_Диапазон».

    Внутри делегирует финальную реализацию удаления листов, чтобы логика
    (patterns + optional filter) совпадала с «Финальная_обработка».
    """
    # 1) Сначала проверяем фильтр sheets (чтобы не дергать удаление на всех листах).
    #    Для текущего листа этот блок вернётся только если лист попадает под sheets.
    try:
        block = _lm_pp_sheet_list_block_for_sheet("удаление_листов", extra_args, sheet)
    except Exception:
        block = None
    if block is None:
        return

    # 2) Удаление нужно выполнить один раз за прогон постобработки.
    ctx = lm_pp_active_context()
    if ctx is None:
        ctx = {}
        lm_pp_set_active_context(ctx)
    if ctx.get("_lm_pp_delete_sheets_done"):
        return
    ctx["_lm_pp_delete_sheets_done"] = True
    lm_pp_set_active_context(ctx)

    raw = _lm_pp_extra_args_rest_text(extra_args)
    try:
        import libre_macros_final_lib as fl
    except ImportError:
        return

    # sheets/data_ranges/header_row_ranges не используем — только sheet_specs из C.
    try:
        fl.lm_final_delete_sheets(doc, None, None, None, raw)
    except Exception as err:
        _lm_log_postprocess(doc, getattr(sheet, "Name", ""), "диапазон", "удаление_листов", "ошибка", str(err))


def lm_pp_range_hide_sheets(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Доступность «скрытие_листов» в «Постобработка_Диапазон».
    """
    try:
        block = _lm_pp_sheet_list_block_for_sheet("скрытие_листов", extra_args, sheet)
    except Exception:
        block = None
    if block is None:
        return

    ctx = lm_pp_active_context()
    if ctx is None:
        ctx = {}
        lm_pp_set_active_context(ctx)
    if ctx.get("_lm_pp_hide_sheets_done"):
        return
    ctx["_lm_pp_hide_sheets_done"] = True
    lm_pp_set_active_context(ctx)

    raw = _lm_pp_extra_args_rest_text(extra_args)
    try:
        import libre_macros_final_lib as fl
    except ImportError:
        return

    try:
        fl.lm_final_hide_sheets(doc, None, None, None, raw)
    except Exception as err:
        _lm_log_postprocess(doc, getattr(sheet, "Name", ""), "диапазон", "скрытие_листов", "ошибка", str(err))


def lm_pp_range_set_page_style(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Настроить стиль страницы листа для печати (ориентация, вписывание в страницу).

    Параметры:
        doc — SpreadsheetDocument (StyleFamilies PageStyles)
        sheet — XSpreadsheet (sheet.PageStyle)
        data_range, header_row_range — не используются
        *extra_args — rest из колонки C: orientation, fit_to_pages

    JSON колонки C: [{"v":1,"fn":"стиль_печати","orientation":"landscape","fit_pages":1}].
        *(пусто)* — landscape, 1 (по умолчанию)

    Карта: «стиль_печати».

    Когда вызывать:
        Последним перед выдачей отчёта пользователю; не влияет на экранное отображение ячеек.

    Что меняет:
        Orientation и ScaleToPagesX/Y стиля страницы листа.
    """
    if _lm_pp_param_skip_sheet("стиль_печати", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet("стиль_печати", extra_args, doc, sheet)
    if block is not None:
        orientation, fit_to_pages = _lm_pp_print_style_from_block(block)
    else:
        orientation = "landscape"
        fit_to_pages = 1
        if extra_args:
            orientation = str(extra_args[0]).lower()
        if len(extra_args) > 1:
            try:
                fit_to_pages = int(extra_args[1])
            except (ValueError, TypeError):
                fit_to_pages = 1
    
    try:
        page_style = doc.StyleFamilies.getByName("PageStyles").getByName(sheet.PageStyle)

        if orientation == "landscape":
            page_style.Orientation = 1  # LANDSCAPE
        else:
            page_style.Orientation = 0  # PORTRAIT

        if fit_to_pages > 0:
            page_style.ScaleToPagesX = fit_to_pages
            page_style.ScaleToPagesY = 0  # высота — авто
    except Exception:
        pass


def lm_pp_range_delete_columns( doc, sheet, data_range, header_row_range, *name_markers ):
    """
    Назначение:
        Удалить столбцы по смешанным селекторам: подстрока/шаблон заголовка,
        имя из списка, индекс 1-based или буква столбца (A, AA).

    Параметры:
        doc — SpreadsheetDocument (не используется)
        sheet — XSpreadsheet
        data_range — якорь области (строка заголовков берётся из header_row_range)
        header_row_range — строка заголовков
        *name_markers — rest из C: токены через запятую

    JSON колонки C: [{"v":1,"fn":"удалить_столбцы","markers":["#Путь","3","A","Condit*"]}].
    Шаблон: `Condit*` — заголовок начинается с Condit; `*ция` — оканчивается на «ция».

    Карта: «удалить_столбцы».

    Когда вызывать:
        После копирования. Совпадения собираются по текущим заголовкам,
        удаление — справа налево (индексы не сдвигаются).

    Что меняет:
        sheet.getColumns().removeByIndex для каждого найденного столбца.
    """
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("удалить_столбцы", name_markers, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet("удалить_столбцы", name_markers, doc, sheet)
    if block is not None:
        try:
            from libre_macros_param_codec import _normalize_markers_list

            markers = _normalize_markers_list(block.get("markers"))
        except Exception:
            raw = block.get("markers")
            if isinstance(raw, (list, tuple)):
                markers = [str(m).strip() for m in raw if str(m).strip()]
            else:
                markers = []
                for part in str(raw or "").replace(";", ",").split(","):
                    p = part.strip()
                    if p:
                        markers.append(p)
        if len(markers) == 0:
            markers = ["маркер", "marker"]
    elif len(name_markers) == 0:
        markers = ["маркер", "marker"]
    else:
        markers = []
        for item in name_markers:
            s = str(item).strip()
            if s == "":
                continue
            # «11,12» одним аргументом — разбить
            if ("," in s or ";" in s) and not s.startswith("["):
                for part in s.replace(";", ",").split(","):
                    p = part.strip()
                    if not p:
                        continue
                    dup = False
                    j = 0
                    while j < len(markers):
                        if markers[j].lower() == p.lower():
                            dup = True
                            break
                        j = j + 1
                    if not dup:
                        markers.append(p)
                continue
            dup = False
            j = 0
            while j < len(markers):
                if markers[j].lower() == s.lower():
                    dup = True
                    break
                j = j + 1
            if not dup:
                markers.append(s)

    if len(markers) == 0:
        _lm_log_postprocess(
            doc,
            sheet.Name if sheet is not None else "",
            "удалить_столбцы",
            "—",
            "пропуск",
            "нет маркеров",
        )
        return

    try:
        header_row = _lm_pp_range_address(header_row_range)[1]
    except Exception:
        header_row = 0

    sheet_name = sheet.Name if sheet is not None else ""
    end_col, _end_row = get_sheet_used_bounds(sheet)
    if end_col < 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "удалить_столбцы",
            "—",
            "пропуск",
            "пустой лист",
        )
        return
    hdr_range = sheet.getCellRangeByPosition(0, header_row, end_col, header_row)
    titles = _lm_pp_header_titles(sheet, hdr_range)
    h_sc, _h_sr, _h_ec, _h_er = _lm_pp_range_address(hdr_range)

    to_delete = []
    seen = set()
    mi = 0
    while mi < len(markers):
        cols = _lm_pp_resolve_delete_marker_columns(
            markers[mi], sheet, hdr_range, end_col
        )
        if len(cols) == 0:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "удалить_столбцы",
                markers[mi],
                "пропуск",
                "нет столбца по маркеру «%s»" % markers[mi],
            )
        else:
            for col_idx in cols:
                if col_idx not in seen:
                    seen.add(col_idx)
                    to_delete.append(col_idx)
        mi = mi + 1

    to_delete.sort(reverse=True)
    total_deleted = 0
    di = 0
    while di < len(to_delete):
        col_to_delete = to_delete[di]
        pos = col_to_delete - h_sc
        col_title = titles[pos] if 0 <= pos < len(titles) else ""
        try:
            _lm_sheet_delete_columns(sheet, col_to_delete, 1)
            total_deleted = total_deleted + 1
            _lm_log_postprocess(
                doc,
                sheet_name,
                "удалить_столбцы",
                str(col_to_delete + 1),
                "ok",
                "удалён столбец %d «%s»" % (col_to_delete + 1, col_title),
            )
        except Exception as err:
            msg = "столбец %d «%s»: %s" % (col_to_delete + 1, col_title, err)
            _lm_log_postprocess(
                doc,
                sheet_name,
                "удалить_столбцы",
                str(col_to_delete + 1),
                "ошибка",
                msg,
            )
            print("  ⚠ Не удалось удалить столбец: %s" % msg)
        _lm_ui_yield(force=True)
        di = di + 1

    _lm_log_postprocess(
        doc,
        sheet_name,
        "удалить_столбцы",
        "итог",
        "инфо",
        "удалено столбцов: %d, маркеры: %s"
        % (total_deleted, ", ".join(str(m) for m in markers)),
    )


def _lm_pp_skip_rows_formula_for_columns(col_indices):
    """Формула Calc: 1, если все указанные столбцы пусты, иначе 0."""
    if col_indices is None or len(col_indices) == 0:
        return ""
    parts = []
    ci = 0
    while ci < len(col_indices):
        parts.append("ISBLANK(%s{row_id})" % col_index_to_letters(col_indices[ci]))
        ci = ci + 1
    if len(parts) == 1:
        bool_expr = parts[0]
    else:
        bool_expr = "AND(" + ";".join(parts) + ")"
    # Делать числовой результат, чтобы AutoFilter мог сравнивать как numeric.
    return "IF(%s;1;0)" % bool_expr


def _lm_pp_delete_rows_by_filtered_helper( doc, sheet, hdr_row, data_start, data_end, helper_col, filter_sc, sheet_name="" ):
    """
    Автофильтр по helper_col==TRUE(1) и удаление видимых строк пакетами.
    Возвращает число удалённых строк или -1 при ошибке.
    """
    if doc is None or sheet is None:
        return -1
    db_name = "_lm_pp_skip_rows_filter"
    fd = None
    deleted = 0
    try:
        try:
            if db_name in doc.DatabaseRanges:
                doc.DatabaseRanges.removeByName(db_name)
        except Exception:
            pass
        addr = sheet.getCellRangeByPosition(
            int(filter_sc), int(hdr_row), int(helper_col), int(data_end)
        ).getRangeAddress()
        doc.DatabaseRanges.addNewByName(db_name, addr)
        doc.DatabaseRanges.getByName(db_name).AutoFilter = True

        fd = sheet.createFilterDescriptor(True)
        ff = uno.createUnoStruct("com.sun.star.sheet.TableFilterField")
        ff.Field = int(helper_col) - int(filter_sc)
        try:
            from com.sun.star.sheet.FilterOperator import EQUAL
            ff.Operator = EQUAL
        except Exception:
            ff.Operator = 0
        ff.IsNumeric = True
        ff.NumericValue = 1.0
        try:
            fd.setContainsHeader(True)
        except Exception:
            # В разных версиях UNO/LO API метод может отсутствовать;
            # пробуем свойство/оставляем дефолт.
            try:
                fd.ContainsHeader = True
            except Exception:
                pass
        fd.setFilterFields((ff,))
        sheet.filter(fd)

        visible_rows = []
        row = int(data_start)
        while row <= int(data_end):
            vis = True
            try:
                vis = bool(sheet.getRows().getByIndex(row).IsVisible)
            except Exception:
                vis = True
            if vis:
                visible_rows.append(row)
            row = row + 1

        if len(visible_rows) == 0:
            return 0

        next_log = 200
        i = len(visible_rows) - 1
        while i >= 0:
            run_end = visible_rows[i]
            run_start = run_end
            while i > 0 and visible_rows[i - 1] == run_start - 1:
                i = i - 1
                run_start = visible_rows[i]
            count = run_end - run_start + 1
            _lm_sheet_delete_rows(sheet, run_start, count)
            deleted = deleted + count
            while deleted >= next_log:
                _lm_pp_skip_rows_dbg(
                    doc,
                    sheet_name,
                    "autofilter batch deleted=%d (last group %d..%d x%d)"
                    % (deleted, run_start + 1, run_end + 1, count),
                )
                _lm_ui_yield(force=True)
                next_log = next_log + 200
            i = i - 1
        return deleted
    except Exception as err:
        _lm_pp_skip_rows_dbg(
            doc,
            sheet_name,
            "autofilter delete failed: %s" % (str(err)[:180],),
        )
        return -1
    finally:
        try:
            if fd is not None:
                sheet.removeFilter(fd)
        except Exception:
            pass
        try:
            if db_name in doc.DatabaseRanges:
                doc.DatabaseRanges.removeByName(db_name)
        except Exception:
            pass


def lm_pp_range_skip_empty_rows(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Удалить строки данных на листе результата по тем же правилам, что
        «Пропуск_строк_источника», но после сбора (постобработка диапазона).

    Параметры:
        doc, sheet, data_range, header_row_range — стандартные range
        *extra_args — rest из колонки C (parse_pp_extra_args)

    JSON колонки C: [{"v":1,"fn":"удаление_строк","mode":"columns","columns":["ФИО"]}]
    или mode formula: {"mode":"formula","formula":"E{row_id}+F{row_id}=0"}.
        Режим «формула»: критерий «строка пустая» — формула должна давать ИСТИНА
        именно для пустых (ненужных) строк; такие строки удаляются. Пример:
        «E{row_id}+F{row_id}=0». Во временный столбец справа от data_range
        вставляется формула Calc, пересчёт, удаление строк с ИСТИНА, столбец
        убирается. {row_id} — номер первой строки данных в Calc (1-based).

    Карта: «удаление_строк» (алиас: «пропуск_пустых_строк»).

    Когда вызывать:
        После сбора и «Ячейка_В_Столбец»; до row-постобработки и «удалить_столбцы»,
        если столбцы условия ещё нужны.

    Что меняет:
        Физически удаляет строки листа (removeByIndex); в режиме формулы — кратко
        добавляет и убирает служебный столбец справа от диапазона данных.
    """
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("удаление_строк", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet("удаление_строк", extra_args, doc, sheet)
    # parts из JSON-блока — без join/split по «,» (иначе ломаются 'Фамилия, Имя').
    column_parts = None
    if block is not None:
        mode = str(block.get("mode") or "columns").casefold()
        if mode == "formula":
            spec_text = str(block.get("formula") or "").strip()
        elif mode in ("none", ""):
            spec_text = ""
        else:
            cols = block.get("columns") or []
            column_parts = []
            ci = 0
            while ci < len(cols):
                t = str(cols[ci]).strip()
                if t != "":
                    column_parts.append(t)
                ci = ci + 1
            try:
                from libre_macros_param_codec import join_column_name_tokens

                spec_text = join_column_name_tokens(column_parts)
            except Exception:
                spec_text = ",".join(column_parts)
    else:
        spec_text = _lm_pp_reconstruct_skip_spec_from_extra(extra_args)
    if spec_text == "":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "удаление_строк",
            "—",
            "пропуск",
            "пустая колонка C",
        )
        return

    parsed = parse_skip_row_spec(spec_text)
    if parsed.get("mode") == "none":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "удаление_строк",
            spec_text,
            "пропуск",
            "не распознано условие",
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    hdr_row = h_sr
    col_count = ec - sc + 1
    if er < sr:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "удаление_строк",
            spec_text,
            "пропуск",
            "нет строк данных (end_row=%d)" % er,
        )
        return

    mode = parsed.get("mode")
    col_indices = None
    helper_col = -1
    rowid_col = -1
    formula_text = ""
    used_end_col, _used_end_row = get_sheet_used_bounds(sheet)
    # Служебные столбцы критерия/rowid всегда убирать (раньше оставляли при DEBUG).
    keep_helpers = False

    if mode == "columns":
        parts = column_parts if column_parts is not None else (parsed.get("parts") or [])
        col_indices, _missing_parts = _resolve_skip_column_indices(
            sheet, hdr_row, sc, col_count, parts
        )
        if len(col_indices) == 0:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "удаление_строк",
                spec_text,
                "ошибка",
                "не найдены столбцы «%s»" % spec_text,
            )
            return
        formula_text = _lm_pp_skip_rows_formula_for_columns(col_indices)
        _lm_pp_skip_rows_dbg(
            doc,
            sheet_name,
            "mode=columns cols=%s rows=%d..%d formula=%s"
            % (
                ",".join(str(i + 1) for i in col_indices),
                sr + 1,
                er + 1,
                formula_text,
            ),
        )
    elif mode == "formula":
        formula_text = _lm_skip_expand_formula_refs(
            sheet, hdr_row, sc, col_count, parsed.get("formula", ""), doc=doc
        )
        _lm_pp_skip_rows_dbg(
            doc,
            sheet_name,
            "mode=formula rows=%d..%d шаблон=%r"
            % (sr + 1, er + 1, formula_text[:120]),
        )

    deleted = 0
    if formula_text != "":
        # Без ведущего '=' — иначе IF(=ISBLANK(...);0;1) ломает Calc.
        if formula_text.lstrip().startswith("="):
            formula_text = formula_text.lstrip()[1:].lstrip()
        helper_col = max(int(ec) + 1, int(used_end_col) + 1)
        rowid_col = helper_col + 1

        # Критерий удаления: 0 — удалить, 1 — оставить.
        crit_formula_text = "IF(%s;0;1)" % formula_text
        calc_formula = _lm_pp_calc_formula_string(crit_formula_text, sr, helper_col, doc=doc)
        _lm_pp_skip_rows_dbg(
            doc,
            sheet_name,
            "helper criterion col=%d first=%r" % (helper_col + 1, calc_formula[:120]),
        )
        if not _lm_pp_set_cell_formula(
            doc,
            sheet.getCellByPosition(helper_col, sr),
            calc_formula,
            sr,
            helper_col,
            sheet_name,
        ):
            _lm_log_postprocess(
                doc,
                sheet_name,
                "удаление_строк",
                spec_text,
                "ошибка",
                "не удалось записать helper-критерий",
            )
            return

        row_formula = _lm_pp_calc_formula_string("ROW()", sr, rowid_col, doc=doc)
        _lm_pp_skip_rows_dbg(
            doc,
            sheet_name,
            "helper rowid col=%d first=%r" % (rowid_col + 1, row_formula[:120]),
        )
        if not _lm_pp_set_cell_formula(
            doc,
            sheet.getCellByPosition(rowid_col, sr),
            row_formula,
            sr,
            rowid_col,
            sheet_name,
        ):
            _lm_log_postprocess(
                doc,
                sheet_name,
                "удаление_строк",
                spec_text,
                "ошибка",
                "не удалось записать helper-ROW()",
            )
            return

        if er > sr:
            if not _lm_pp_fill_formula_down(
                doc,
                sheet,
                helper_col,
                sr,
                er,
                sheet_name,
                dbg_channel="удаление_строк",
            ):
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "удаление_строк",
                    spec_text,
                    "ошибка",
                    "не удалось протянуть helper-критерий",
                )
                return
            if not _lm_pp_fill_formula_down(
                doc,
                sheet,
                rowid_col,
                sr,
                er,
                sheet_name,
                dbg_channel="удаление_строк",
            ):
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "удаление_строк",
                    spec_text,
                    "ошибка",
                    "не удалось протянуть helper-ROW()",
                )
                return

        _lm_pp_calculate_doc(doc)
        _lm_pp_skip_rows_dbg(
            doc,
            sheet_name,
            "пересчёт выполнен, материализация helper cols %d/%d"
            % (helper_col + 1, rowid_col + 1),
        )

        # Материализуем helper-колонки (значения вместо формул).
        rr = sr
        while rr <= er:
            try:
                hc = sheet.getCellByPosition(helper_col, rr)
                hv = 1 if _lm_pp_cell_is_true(hc) else 0
                hc.Formula = ""
                hc.Value = hv
            except Exception:
                pass
            try:
                rc = sheet.getCellByPosition(rowid_col, rr)
                rv = rc.Value
                rc.Formula = ""
                rc.Value = float(rv)
            except Exception:
                pass
            rr = rr + 1

        sort_sc = 0
        sort_ec, _sort_er = get_sheet_used_bounds(sheet)
        sort_ec = max(int(sort_ec), int(rowid_col))

        # Сортировка: сначала 0 (удаляемые), потом 1 (оставляемые).
        resolved = [(int(helper_col), True)]
        ok, err = _lm_pp_range_sort_via_native(
            doc,
            sheet,
            sort_sc,
            sr,
            sort_ec,
            er,
            resolved,
            sheet_name=sheet_name,
            fn_label="удаление_строк/критерий",
        )
        if not ok:
            ok, err = _lm_pp_range_sort_via_data_array(
                sheet,
                sort_sc,
                sr,
                sort_ec,
                er,
                resolved,
                sort_sc,
                sheet_name=sheet_name,
                fn_label="удаление_строк/критерий",
            )
        if not ok:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "удаление_строк",
                spec_text,
                "ошибка",
                "не удалось отсортировать по helper-критерию: %s" % err,
            )
            return

        # Граница: первая строка, где helper == 1 (оставляем).
        lo = int(sr)
        hi = int(er)
        first_keep = int(er) + 1
        while lo <= hi:
            mid = (lo + hi) // 2
            is_keep = False
            try:
                is_keep = _lm_pp_cell_is_true(sheet.getCellByPosition(helper_col, mid))
            except Exception:
                is_keep = False
            if is_keep:
                first_keep = mid
                hi = mid - 1
            else:
                lo = mid + 1

        del_count = max(0, int(first_keep) - int(sr))
        _lm_pp_skip_rows_dbg(
            doc,
            sheet_name,
            "sorted boundary first_keep=%d delete_count=%d" % (first_keep + 1, del_count),
        )

        if del_count > 0:
            _lm_sheet_delete_rows(sheet, sr, del_count)
            deleted = del_count
            _lm_pp_skip_rows_dbg(
                doc,
                sheet_name,
                "delete one-shot rows %d..%d (count=%d)"
                % (sr + 1, sr + del_count, del_count),
            )

        # Восстановить исходный порядок по rowid helper (по возрастанию).
        new_er = int(er) - int(deleted)
        if new_er >= int(sr):
            sort_ec2, _ = get_sheet_used_bounds(sheet)
            sort_ec2 = max(int(sort_ec2), int(rowid_col))
            resolved_back = [(int(rowid_col), True)]
            ok2, err2 = _lm_pp_range_sort_via_native(
                doc,
                sheet,
                sort_sc,
                sr,
                sort_ec2,
                new_er,
                resolved_back,
                sheet_name=sheet_name,
                fn_label="удаление_строк/restore_order",
            )
            if not ok2:
                ok2, err2 = _lm_pp_range_sort_via_data_array(
                    sheet,
                    sort_sc,
                    sr,
                    sort_ec2,
                    new_er,
                    resolved_back,
                    sort_sc,
                    sheet_name=sheet_name,
                    fn_label="удаление_строк/restore_order",
                )
            if not ok2:
                _lm_pp_skip_rows_dbg(
                    doc,
                    sheet_name,
                    "restore sort failed: %s" % (err2 or "?"),
                )

    if (helper_col >= 0 or rowid_col >= 0) and (not keep_helpers):
        # В не-debug режиме helper-колонки очищаем.
        try:
            if rowid_col >= 0:
                _lm_sheet_delete_columns(sheet, rowid_col, 1)
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "удаление_строк",
                spec_text,
                "ошибка",
                "не удалён helper-rowid %d: %s" % (rowid_col + 1, err),
            )
        try:
            if helper_col >= 0:
                _lm_sheet_delete_columns(sheet, helper_col, 1)
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "удаление_строк",
                spec_text,
                "ошибка",
                "не удалён helper-критерий %d: %s" % (helper_col + 1, err),
            )

    _lm_log_postprocess(
        doc,
        sheet_name,
        "удаление_строк",
        spec_text,
        "ok",
        "удалено строк: %d (режим %s)" % (deleted, mode),
    )


def lm_pp_range_concat_columns(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Добавить столбец с заданным заголовком; в каждой строке данных —
        конкатенация значений указанных столбцов через разделитель.

    Параметры:
        doc, sheet, data_range, header_row_range — стандартные range
        *extra_args — dict-блок JSON (param_decode)

    JSON колонки C: [{"v":1,"fn":"конкатенация_столбцов","new_column":"ФИО","separator":" / ","columns":[1,3]}].

    Карта: «конкатенация_столбцов».

    Когда вызывать:
        После сбора данных, до удаления служебных столбцов (если они участвуют в склейке).

    Что меняет:
        Новый или существующий столбец с заголовком col_name и текстом в строках данных.
        Формат клонируется поддиапазонами с последнего столбца data_range: отдельно строка
        заголовка и блок данных; ширина — OptimalWidth; end_col контекста расширяется.
    """
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("конкатенация_столбцов", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(
        "конкатенация_столбцов", extra_args, doc, sheet
    )
    parsed = None
    if block is not None:
        col_name = str(block.get("new_column") or "").strip()
        delim = str(block.get("separator") or "")
        indices = block.get("columns") or []
        if col_name and indices:
            parsed = (col_name, delim, indices)
    if parsed is None:
        if extra_args and len(extra_args) == 1 and "," in str(extra_args[0]):
            ea = parse_pp_extra_args(extra_args[0])
        else:
            ea = extra_args
        parsed = _lm_pp_parse_concat_columns_extra(ea)
    if parsed is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "конкатенация_столбцов",
            "ошибка",
            "ожидается C: Имя_Столбца[,Разделитель],индекс1,индекс2,…",
        )
        return
    col_name, delim, indices = parsed
    # columns: поддерживаем индексы (1-based), буквы (A, AA) и заголовки
    resolved = []
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    for tok in indices:
        if tok is None or isinstance(tok, bool):
            continue
        if isinstance(tok, int):
            ci = int(tok) - 1
            if ci >= 0:
                resolved.append(ci)
            continue
        s = str(tok).strip()
        if s == "":
            continue
        if s.isdigit():
            ci = int(s) - 1
            if ci >= 0:
                resolved.append(ci)
            continue
        col_idx = _lm_pp_resolve_column_token(s, sheet, header_row_range, sc, ec)
        if col_idx is not None:
            resolved.append(int(col_idx))
    if not resolved:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "конкатенация_столбцов",
            "ошибка",
            "столбцы не найдены (columns)",
        )
        return
    # unique, keep order
    seen = set()
    resolved2 = []
    for c in resolved:
        if c not in seen:
            seen.add(c)
            resolved2.append(c)
    indices = resolved2
    loc = _lm_pp_resolve_derived_column(
        sheet, data_range, header_row_range, col_name
    )
    if loc is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "конкатенация_столбцов",
            "ошибка",
            "не удалось добавить столбец «%s»" % col_name,
        )
        return
    sr = loc["sr"]
    er = loc["er"]
    h_sr = loc["h_sr"]
    ref_col = loc["ref_col"]
    target_col = loc["target_col"]
    _lm_pp_clone_concat_column_formats(
        sheet, ref_col, target_col, h_sr, sr, er
    )
    sheet.getCellByPosition(target_col, h_sr).String = col_name
    r = sr
    n = 0
    while r <= er:
        parts = []
        for idx in indices:
            # indices уже 0-based (после resolve выше).
            try:
                c = int(idx)
            except (TypeError, ValueError):
                continue
            if c < 0:
                continue
            parts.append(cell_text(sheet.getCellByPosition(c, r)))
        sheet.getCellByPosition(target_col, r).String = delim.join(parts)
        n = n + 1
        if _lm_ui_should_update_status(n, is_last=(r == er)):
            _lm_ui_status_set(
                n,
                _lm_pp_status_text(
                    "диапазон",
                    "конкатенация: строки %d/%d" % (r + 1, er + 1),
                ),
            )
        _lm_ui_yield(counter=n)
        r = r + 1
    _lm_pp_finalize_derived_column(sheet, ref_col, target_col, h_sr, sr)
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        "конкатенация_столбцов",
        "ok",
        "столбец %d «%s», образец %d, индексы %s, разделитель «%s»"
        % (target_col, col_name, ref_col, ",".join(str(i) for i in indices), delim),
    )


def lm_pp_range_split_by_columns(doc, sheet, data_range, header_row_range, *extra_args):
    """Разделить один столбец по разделителю с вставкой новых столбцов."""
    sheet_name = sheet.Name if sheet is not None else ""
    fn_key = "разделить_по_столбцам"
    if _lm_pp_param_skip_sheet(fn_key, extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(fn_key, extra_args, doc, sheet)
    if not isinstance(block, dict):
        _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "ожидается JSON-блок")
        return
    # Защита от ситуаций, когда финальная обёртка передаёт "нецелевой" блок.
    # Тогда применяем explicit sheet/sheets match и не модифицируем лист.
    try:
        target_names = _lm_pp_block_target_sheet_names(block)
        if target_names:
            if not _lm_pp_block_matches_sheet(block, doc, sheet, sheet_name):
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    fn_key,
                    "пропуск",
                    "лист не совпадает с sheet/sheets фильтром",
                )
                return
    except Exception:
        pass
    column = str(block.get("column") or "").strip()
    delimiter = str(block.get("delimiter") or "")
    if column == "" or delimiter == "":
        _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "нужны column и delimiter")
        return
    split_mode = str(block.get("split_mode") or "each").strip().casefold()
    if split_mode not in ("each", "leftmost"):
        split_mode = "each"
    trim_parts = bool(block.get("trim_parts", True))
    source_policy = str(block.get("source_policy") or "replace").strip().casefold()
    if source_policy not in ("replace", "keep"):
        source_policy = "replace"
    non_text = str(block.get("non_text") or "skip").strip().casefold()
    if non_text not in ("skip", "empty", "text", "fail"):
        non_text = "skip"
    max_columns = 0
    try:
        max_columns = int(block.get("max_columns") or 0)
    except (TypeError, ValueError):
        max_columns = 0
    if max_columns < 0:
        _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "max_columns < 0")
        return
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    _h_sc, h_sr, _h_ec, _h_er = _lm_pp_range_address(header_row_range)
    src_col, src_err = _lm_pp_resolve_column_token_unique(column, sheet, header_row_range, sc, ec)
    if src_col is None:
        _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", src_err or "column не найден")
        return
    pos_raw = str(block.get("position") or "_после_").strip().casefold()
    pos_map = {
        "_начало_": "start",
        "_конец_": "end",
        "_перед_": "before",
        "_после_": "after",
        "start": "start",
        "end": "end",
        "before": "before",
        "after": "after",
    }
    position = pos_map.get(pos_raw, "after")
    rel_token = str(block.get("relative_column") or "").strip() or column
    rel_col = src_col
    if position in ("before", "after"):
        rel_col, rel_err = _lm_pp_resolve_column_token_unique(rel_token, sheet, header_row_range, sc, ec)
        if rel_col is None:
            _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", rel_err or "relative_column не найден")
            return
    if position == "start":
        insert_at = sc
    elif position == "end":
        insert_at = ec + 1
    elif position == "before":
        insert_at = rel_col
    else:
        insert_at = rel_col + 1
    rng = sheet.getCellRangeByPosition(sc, sr, ec, er)
    data = rng.getDataArray()
    row_parts = {}
    max_parts = 1
    data_rows = 0
    skipped_non_text = 0

    def _split_value(text_value):
        text = str(text_value or "")
        if split_mode == "leftmost":
            parts = text.split(delimiter, 1)
        else:
            parts = text.split(delimiter)
        if trim_parts:
            parts = [p.strip() for p in parts]
        if max_columns > 0 and len(parts) > max_columns:
            parts = parts[: max_columns - 1] + [delimiter.join(parts[max_columns - 1 :])]
        if len(parts) == 0:
            return [""]
        return parts

    r = sr
    while r <= er:
        row_idx = r - sr
        if r <= h_sr:
            r = r + 1
            continue
        data_rows = data_rows + 1
        cell_val = data[row_idx][src_col - sc]
        is_text = isinstance(cell_val, str)
        if not is_text:
            if non_text == "skip":
                skipped_non_text = skipped_non_text + 1
                r = r + 1
                continue
            if non_text == "empty":
                row_parts[row_idx] = [""]
                r = r + 1
                continue
            if non_text == "fail":
                _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "нетекстовая ячейка в строке %d" % (r + 1))
                return
        parts = _split_value(cell_val)
        row_parts[row_idx] = parts
        if len(parts) > max_parts:
            max_parts = len(parts)
        r = r + 1
    n_new = max(0, max_parts - 1) if source_policy == "replace" else max_parts
    if n_new > 0:
        sheet.getColumns().insertByIndex(insert_at, n_new)
    if insert_at <= src_col:
        src_col = src_col + n_new
    new_ec = ec + n_new
    out_rng = sheet.getCellRangeByPosition(sc, sr, new_ec, er)
    out_data = [list(row) for row in out_rng.getDataArray()]

    expected_rows = int(er) - int(sr) + 1
    expected_cols = int(new_ec) - int(sc) + 1
    if expected_rows < 0:
        expected_rows = 0
    if expected_cols < 0:
        expected_cols = 0
    # UNO иногда отдаёт «рваные» массивы (без хвостовых пустых колонок/строк).
    # Нормализуем, чтобы избежать list assignment out of range.
    if len(out_data) < expected_rows:
        out_data.extend([[""] * expected_cols for _ in range(expected_rows - len(out_data))])
    elif len(out_data) > expected_rows:
        out_data = out_data[:expected_rows]
    for ri in range(len(out_data)):
        row = out_data[ri]
        if len(row) < expected_cols:
            row.extend([""] * (expected_cols - len(row)))
        elif len(row) > expected_cols:
            del row[expected_cols:]

    r = sr
    while r <= er:
        row_idx = r - sr
        if r <= h_sr:
            r = r + 1
            continue
        parts = row_parts.get(row_idx)
        if parts is None:
            r = r + 1
            continue
        if source_policy == "replace":
            padded = list(parts) + [""] * max(0, max_parts - len(parts))
            c0 = int(src_col) - int(sc)
            if 0 <= row_idx < len(out_data) and 0 <= c0 < expected_cols:
                out_data[row_idx][c0] = padded[0]
            i = 1
            while i < max_parts:
                ci = int(insert_at) + int(i) - 1 - int(sc)
                if 0 <= row_idx < len(out_data) and 0 <= ci < expected_cols:
                    out_data[row_idx][ci] = padded[i]
                i = i + 1
        else:
            padded = list(parts) + [""] * max(0, max_parts - len(parts))
            i = 0
            while i < max_parts:
                ci = int(insert_at) + int(i) - int(sc)
                if 0 <= row_idx < len(out_data) and 0 <= ci < expected_cols:
                    out_data[row_idx][ci] = padded[i]
                i = i + 1
        r = r + 1
    out_rng.setDataArray(tuple(tuple(row) for row in out_data))
    new_names = block.get("new_names")
    names = []
    if isinstance(new_names, (list, tuple)):
        names = [str(x).strip() for x in new_names if str(x).strip() != ""]
    base_header = cell_text(sheet.getCellByPosition(src_col, h_sr))
    if base_header == "":
        base_header = "Столбец"
    i = 0
    while i < n_new:
        if i < len(names):
            hdr = names[i]
        else:
            hdr = _lm_pp_allocate_new_header_name(sheet, h_sr, "%s.%d" % (base_header, i + 1))
        sheet.getCellByPosition(insert_at + i, h_sr).String = hdr
        i = i + 1
    status = "ok" if n_new > 0 else "пропуск"
    status_note = "вставка=0 (delimiter не расколол или header/data не совпали)" if n_new == 0 else ""
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        fn_key,
        status,
        "column=%s delimiter=%r max_parts=%d inserted=%d rows=%d skipped_non_text=%d expected_rows=%d expected_cols=%d %s"
        % (
            column,
            delimiter,
            max_parts,
            n_new,
            data_rows,
            skipped_non_text,
            expected_rows,
            expected_cols,
            status_note,
        ),
    )


def lm_pp_range_conditional_column(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Условный столбец по правилам (аналог Power Query Conditional Column).

    Ключ:
        fn=«условный_столбец»

    Оценка:
        правила исполняются в памяти по `getDataArray`, без `=IF(...)` на листе.

    Поддержка v1:
        `otherwise` может быть вложенным conditional-блоком с собственными `rules`
        (рекурсивная ветка ELSE IF).
    """
    sheet_name = sheet.Name if sheet is not None else ""
    fn_key = "условный_столбец"
    if _lm_pp_param_skip_sheet(fn_key, extra_args, doc, sheet):
        return

    block = _lm_pp_param_block_for_sheet(fn_key, extra_args, doc, sheet)
    if not isinstance(block, dict):
        _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "ожидается JSON-блок")
        return

    # Защита от ситуаций, когда финальная обёртка передаёт "нецелевой" блок.
    # Тогда применяем explicit sheet/sheets match и не модифицируем лист.
    try:
        target_names = _lm_pp_block_target_sheet_names(block)
        if target_names:
            if not _lm_pp_block_matches_sheet(block, doc, sheet, sheet_name):
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    fn_key,
                    "пропуск",
                    "лист не совпадает с sheet/sheets фильтром",
                )
                return
    except Exception:
        pass

    new_column = str(block.get("new_column") or "").strip()
    if new_column == "":
        _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "нужен new_column")
        return

    # Глобальные сравнения (v1).
    trim = bool(block.get("trim", True))
    case_sensitive = bool(block.get("case_sensitive", False))
    compare_as = str(block.get("compare_as") or "auto").strip().casefold()
    if compare_as not in ("auto", "text", "number", "date"):
        compare_as = "auto"

    as_values = bool(block.get("as_values", True))
    output_policy = str(block.get("output_policy") or "new").strip().casefold()
    if output_policy not in ("new", "replace"):
        output_policy = "new"

    blank_left = str(block.get("blank_left") or "as_empty").strip().casefold()
    if blank_left not in ("as_empty", "skip_rule", "fail"):
        blank_left = "as_empty"

    non_comparable = str(block.get("non_comparable") or "false").strip().casefold()
    if non_comparable not in ("false", "fail", "true"):
        non_comparable = "false"
    non_comparable_fail = non_comparable == "fail"

    # Диапазоны.
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)

    # Ищем существующий столбец по заголовку.
    target_col = None
    try:
        titles = _lm_pp_header_titles(sheet, header_row_range)
        if isinstance(titles, (list, tuple)):
            key = lm_identity_key(new_column)
            for i, t in enumerate(titles):
                if lm_identity_key(t) == key:
                    target_col = int(h_sc) + int(i)
                    break
    except Exception:
        target_col = None

    # Куда вставлять при отсутствии столбца.
    pos_raw = str(block.get("position") or "_конец_").strip().casefold()
    pos_map = {
        "_начало_": "start",
        "_конец_": "end",
        "_перед_": "before",
        "_после_": "after",
        "start": "start",
        "end": "end",
        "before": "before",
        "after": "after",
    }
    position = pos_map.get(pos_raw, "end")
    rel_token = str(block.get("relative_column") or "").strip()

    inserted = False
    insert_at = None
    if target_col is None:
        if position == "start":
            insert_at = sc
        elif position == "end":
            insert_at = ec + 1
        elif position in ("before", "after"):
            if rel_token == "":
                # Без relative_column ставим "относительно левого края" (sc).
                rel_col = sc
            else:
                rel_col, rel_err = _lm_pp_resolve_column_token_unique(rel_token, sheet, header_row_range, sc, ec)
                if rel_col is None:
                    _lm_log_postprocess(
                        doc,
                        sheet_name,
                        "диапазон",
                        fn_key,
                        "ошибка",
                        "relative_column не найден: %s" % (rel_err or rel_token),
                    )
                    return
            insert_at = int(rel_col) if position == "before" else int(rel_col) + 1
        else:
            insert_at = ec + 1

        # В v1 ожидаем, что вставка внутри/на границе data_range, чтобы не
        # пришлось усложнять маппинг индексов с "дырками".
        if insert_at < sc or insert_at > ec + 1:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_key,
                "ошибка",
                "неподдерживаемая позиция вставки: insert_at=%d sc=%d ec=%d"
                % (insert_at, sc, ec),
            )
            return

        inserted = True
        target_col = insert_at

    # materialize формулы до getDataArray.
    if as_values and doc is not None:
        try:
            _lm_pp_calculate_doc(doc)
        except Exception:
            pass

    # Читаем правила.
    root_branch = {
        "rules": block.get("rules") or [],
        "otherwise": block.get("otherwise", ""),
        "otherwise_column": block.get("otherwise_column"),
    }
    rules_root = root_branch.get("rules") or []
    if not isinstance(rules_root, (list, tuple)):
        _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "rules должен быть массивом")
        return

    def _is_conditional_block(obj):
        return isinstance(obj, dict) and "rules" in obj

    # 1) Сбор токенов колонок из всех веток (иначе/вложенный otherwise).
    col_tokens = set()

    def _collect_col_tokens(branch):
        if not isinstance(branch, dict):
            return
        rr = branch.get("rules") or []
        if isinstance(rr, (list, tuple)):
            for rule in rr:
                if not isinstance(rule, dict):
                    continue
                for k in ("column", "value_column", "then_column"):
                    tok = rule.get(k)
                    if tok is None:
                        continue
                    ts = str(tok).strip()
                    if ts != "":
                        col_tokens.add(ts)
        ow_col = branch.get("otherwise_column")
        if ow_col is not None:
            ts = str(ow_col).strip()
            if ts != "":
                col_tokens.add(ts)
        ow = branch.get("otherwise")
        if _is_conditional_block(ow):
            _collect_col_tokens(ow)

    _collect_col_tokens(root_branch)

    # 2) Разрешаем токены в абсолютные индексы столбцов (до вставки).
    col_old = {}  # token_str -> old_col_abs (0-based)
    for tok in col_tokens:
        resolved, err = _lm_pp_resolve_column_token_unique(tok, sheet, header_row_range, sc, ec)
        if resolved is None:
            _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "столбец не найден: %s" % (err or tok))
            return
        col_old[tok] = int(resolved)

    # 3) Маппинг после вставки.
    def _shift_col_abs(old_col):
        if not inserted:
            return old_col
        return old_col + 1 if old_col >= insert_at else old_col

    token_to_rel = {}
    for tok, old_idx in col_old.items():
        token_to_rel[tok] = int(_shift_col_abs(old_idx)) - int(sc)

    ref_col = int(ec)
    if inserted:
        try:
            sheet.getColumns().insertByIndex(int(insert_at), 1)
        except Exception as e:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_key,
                "ошибка",
                "не удалось вставить столбец: %s" % (e,),
            )
            return
        if int(insert_at) <= int(sc):
            ref_col = int(sc) + 1 if int(sc) < int(ec) else int(sc)
        else:
            ref_col = int(insert_at) - 1

    # 4) Читаем диапазон для setDataArray: должен включать вставленный столбец.
    if inserted:
        read_sc = sc
        read_ec = ec + 1
    else:
        read_sc = sc
        read_ec = max(ec, int(target_col))

    out_rng = sheet.getCellRangeByPosition(read_sc, sr, read_ec, er)
    out_data = [list(row) for row in out_rng.getDataArray()]

    # UNO иногда отдаёт «рваные» массивы (без хвостовых пустых колонок/строк).
    expected_rows = int(er) - int(sr) + 1
    expected_cols = int(read_ec) - int(read_sc) + 1
    if expected_rows < 0:
        expected_rows = 0
    if expected_cols < 0:
        expected_cols = 0
    if len(out_data) < expected_rows:
        out_data.extend([[""] * expected_cols for _ in range(expected_rows - len(out_data))])
    elif len(out_data) > expected_rows:
        out_data = out_data[:expected_rows]
    for ri in range(len(out_data)):
        row = out_data[ri]
        if len(row) < expected_cols:
            row.extend([""] * (expected_cols - len(row)))
        elif len(row) > expected_cols:
            del row[expected_cols:]

    # target_col в режиме вставки — это индекс уже созданного столбца,
    # поэтому не применяем shift для "старых" колонок.
    target_rel = int(target_col) - int(read_sc)
    if target_rel < 0 or target_rel >= expected_cols:
        _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "не удалось определить индекс целевого столбца")
        return

    # 5) Компараторы.
    def _val_to_text(v):
        if v is None:
            return ""
        if isinstance(v, bool):
            # Дружелюбно по RU (фиксируем для PQ-like output).
            return u"Да" if v else u"Нет"
        try:
            if isinstance(v, float):
                if v == int(v):
                    return str(int(v))
                return str(v)
        except Exception:
            pass
        try:
            if isinstance(v, int):
                return str(v)
        except Exception:
            pass
        try:
            s = str(v)
            return s.strip()
        except Exception:
            return ""

    def _maybe_strip(s):
        return s.strip() if trim else s

    def _norm_text(s):
        s = _maybe_strip(s or "")
        if case_sensitive:
            return s
        return s.casefold()

    def _try_parse_number_from_any(v):
        if v is None:
            return None
        if isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            try:
                return float(v)
            except Exception:
                return None
        s = _val_to_text(v)
        return _lm_pp_try_parse_number_string(s)

    def _try_parse_date_from_any(v):
        # UNO date может быть float (serial), тогда просто сравниваем serial.
        if v is None:
            return None
        if isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            try:
                return float(v)
            except Exception:
                return None
        s = _val_to_text(v)
        return _lm_pp_try_parse_date_string(s)

    text_ops = {
        "eq",
        "ne",
        "contains",
        "not_contains",
        "starts_with",
        "not_starts_with",
        "ends_with",
        "not_ends_with",
    }
    blank_ops = {"is_blank", "is_not_blank"}
    numeric_ops = {"gt", "ge", "lt", "le"}

    # 6) Рекурсивная оценка conditional-блока.
    max_depth = 10
    hit_counts_root = [0 for _ in range(len(rules_root))]
    otherwise_hits_root = 0

    def _eval_output_from_literal_or_column(val_or_col_tok, row_vals):
        if val_or_col_tok is None:
            return ""
        if isinstance(val_or_col_tok, str) and val_or_col_tok in token_to_rel:
            v = row_vals[token_to_rel[val_or_col_tok]]
        else:
            v = val_or_col_tok
        if v is None:
            return ""
        if isinstance(v, bool):
            return u"Да" if v else u"Нет"
        return v

    def _eval_rule(rule, row_vals):
        if not isinstance(rule, dict):
            return False
        op = str(rule.get("op") or "").strip().casefold()
        if op == "":
            return False

        # Левый аргумент всегда из column.
        col_tok = str(rule.get("column") or "").strip()
        if col_tok == "" or col_tok not in token_to_rel:
            return False
        left = row_vals[token_to_rel[col_tok]]
        left_text = _norm_text(_val_to_text(left))
        left_blank = (left_text == "")

        if left_blank and op not in blank_ops:
            if blank_left == "skip_rule":
                return False
            if blank_left == "fail":
                raise ValueError("blank_left=fail")

        if op in blank_ops:
            if op == "is_blank":
                return left_blank
            return not left_blank

        # Правый аргумент.
        if "value_column" in rule and rule.get("value_column") not in (None, ""):
            vc_tok = str(rule.get("value_column") or "").strip()
            if vc_tok == "" or vc_tok not in token_to_rel:
                return False
            right = row_vals[token_to_rel[vc_tok]]
        else:
            right = rule.get("value")

        if op in text_ops:
            right_text = _norm_text(_val_to_text(right))
            if op == "eq":
                return left_text == right_text
            if op == "ne":
                return left_text != right_text
            if op == "contains":
                return right_text in left_text
            if op == "not_contains":
                return right_text not in left_text
            if op == "starts_with":
                return left_text.startswith(right_text)
            if op == "not_starts_with":
                return not left_text.startswith(right_text)
            if op == "ends_with":
                return left_text.endswith(right_text)
            if op == "not_ends_with":
                return not left_text.endswith(right_text)
            return False

        if op in numeric_ops:
            # compare_as определяет тип сравнения, fallback при auto — числа->даты->текст.
            if compare_as == "number":
                ln = _try_parse_number_from_any(left)
                rn = _try_parse_number_from_any(right)
                if ln is None or rn is None:
                    return False if not non_comparable_fail else (_lm_pp_throw_non_comparable())
                if op == "gt":
                    return ln > rn
                if op == "ge":
                    return ln >= rn
                if op == "lt":
                    return ln < rn
                if op == "le":
                    return ln <= rn
                return False

            if compare_as == "date":
                ld = _try_parse_date_from_any(left)
                rd = _try_parse_date_from_any(right)
                if ld is None or rd is None:
                    return False if not non_comparable_fail else (_lm_pp_throw_non_comparable())
                if op == "gt":
                    return ld > rd
                if op == "ge":
                    return ld >= rd
                if op == "lt":
                    return ld < rd
                if op == "le":
                    return ld <= rd
                return False

            # compare_as == "auto"
            ln = _try_parse_number_from_any(left)
            rn = _try_parse_number_from_any(right)
            if ln is not None and rn is not None:
                if op == "gt":
                    return ln > rn
                if op == "ge":
                    return ln >= rn
                if op == "lt":
                    return ln < rn
                if op == "le":
                    return ln <= rn

            ld = _try_parse_date_from_any(left)
            rd = _try_parse_date_from_any(right)
            if ld is not None and rd is not None:
                if op == "gt":
                    return ld > rd
                if op == "ge":
                    return ld >= rd
                if op == "lt":
                    return ld < rd
                if op == "le":
                    return ld <= rd

            # fallback: текстовая лексикографика (PQ-like).
            right_text = _norm_text(_val_to_text(right))
            if op == "gt":
                return left_text > right_text
            if op == "ge":
                return left_text >= right_text
            if op == "lt":
                return left_text < right_text
            if op == "le":
                return left_text <= right_text
            return False

        return False

    def _eval_branch(branch, row_vals, depth, root_depth_is_0=True):
        if not isinstance(branch, dict):
            return ""
        if depth > max_depth:
            raise ValueError("conditional nesting depth exceeded")
        rr = branch.get("rules") or []
        if not isinstance(rr, (list, tuple)):
            rr = []
        # Правила.
        for i, rule in enumerate(rr):
            ok = _eval_rule(rule, row_vals)
            if ok:
                if root_depth_is_0 and depth == 0 and i < len(hit_counts_root):
                    hit_counts_root[i] = hit_counts_root[i] + 1
                # then/then_column
                then_col_tok = rule.get("then_column")
                if then_col_tok not in (None, ""):
                    tok = str(then_col_tok).strip()
                    if tok in token_to_rel:
                        v = row_vals[token_to_rel[tok]]
                        return u"Да" if v is True else (u"Нет" if v is False else ("" if v is None else v))
                # literal
                then_val = rule.get("then")
                if then_val is None:
                    return ""
                if isinstance(then_val, bool):
                    return u"Да" if then_val else u"Нет"
                return then_val

        # Ничего не сработало: otherwise.
        ow = branch.get("otherwise", "")
        if _is_conditional_block(ow):
            return _eval_branch(ow, row_vals, depth + 1, root_depth_is_0=root_depth_is_0)

        ow_col_tok = branch.get("otherwise_column")
        if ow_col_tok not in (None, ""):
            tok = str(ow_col_tok).strip()
            if tok in token_to_rel:
                v = row_vals[token_to_rel[tok]]
                return u"Да" if v is True else (u"Нет" if v is False else ("" if v is None else v))

        if ow is None:
            return ""
        if isinstance(ow, bool):
            return u"Да" if ow else u"Нет"
        return ow

    # Внутренняя функция для fail non_comparable.
    def _lm_pp_throw_non_comparable():
        raise ValueError("non_comparable=fail")

    # Важно: python не поднимает вложенную функцию в _eval_rule по порядку определений,
    # поэтому _lm_pp_throw_non_comparable должен быть объявлен до вызова.

    # Перезаписываем ссылки в _eval_rule через замыкание: нужен только факт исключения.
    # (в python это и так работает, но оставляем комментарий для читаемости)
    # no-op
    _lm_pp_throw_non_comparable  # silence linter

    # Оценка по строкам (строку заголовка пропускаем).
    data_top = int(h_sr) + 1
    if data_top < int(sr):
        data_top = int(sr)
    r = 0
    while r < expected_rows:
        abs_row = int(sr) + int(r)
        if abs_row <= int(h_sr):
            r = r + 1
            continue
        row_vals = out_data[r]
        try:
            result_value = _eval_branch(root_branch, row_vals, depth=0)
        except Exception as e:
            _lm_log_postprocess(doc, sheet_name, "диапазон", fn_key, "ошибка", "ошибка сравнения: %s" % (e,))
            return
        out_data[r][target_rel] = result_value
        r = r + 1

    try:
        out_rng.setDataArray(tuple(tuple(row) for row in out_data))
    except Exception as e:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_key,
            "ошибка",
            "не удалось записать данные: %s" % (e,),
        )
        return

    _lm_pp_clone_concat_column_formats(
        sheet, ref_col, int(target_col), int(h_sr), data_top, int(er)
    )
    try:
        sheet.getCellByPosition(int(target_col), int(h_sr)).String = new_column
    except Exception:
        pass
    _lm_pp_finalize_derived_column(
        sheet, ref_col, int(target_col), int(h_sr), data_top, doc
    )

    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        fn_key,
        "ok",
        "rows=%d insert=%s target_col=%d position=%s compare_as=%s hits_root=%s"
        % (
            expected_rows,
            str(inserted),
            int(target_col),
            position,
            compare_as,
            ",".join(str(x) for x in hit_counts_root),
        ),
    )


def lm_pp_range_lambda_column(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Столбец по лямбде: вычисляет значения in-memory и пишет setDataArray.

    JSON:
      [{"v":1,"fn":"столбец_по_лямбде",...,"expr":"lambda rows: ..."}]
    """
    sheet_name = sheet.Name if sheet is not None else ""
    fn_key = "столбец_по_лямбде"

    # Запросы к лямбдам не связаны с _lm_pp_param_skip_sheet (sheet_expr может
    # фильтровать иначе), поэтому декодируем блоки руками.
    try:
        blocks = _lm_pp_param_decode(fn_key, extra_args)
    except Exception:
        blocks = []

    if not blocks:
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)

    # Базовый список заголовков по текущему диапазону (для new_column_expr).
    try:
        base_headers = _lm_pp_header_titles(sheet, header_row_range)
    except Exception:
        base_headers = []

    def _all_sheet_names():
        if doc is None:
            return [sheet_name] if sheet_name else []
        try:
            shs = doc.getSheets()
            out = []
            i = 0
            while i < shs.getCount():
                sh = shs.getByIndex(i)
                try:
                    out.append(str(sh.Name).strip())
                except Exception:
                    pass
                i = i + 1
            return [x for x in out if x]
        except Exception:
            return [sheet_name] if sheet_name else []

    all_names = _all_sheet_names()

    def _select_block_by_sheet_expr(block):
        """True, если текущий лист входит в листы для этого блока."""
        try:
            cur = str(sheet_name or "").strip()
            if cur == "":
                return False
            patterns = _lm_pp_block_target_sheet_names(block)
            candidate = []
            if patterns:
                cand_set = set(patterns or [])
                for nm in all_names:
                    if _lm_pp_sheet_name_matches_target(nm, cand_set):
                        candidate.append(nm)
            else:
                candidate = list(all_names or [])

            # Применить sheet_expr поверх предфильтра.
            sheet_expr = unicode(block.get("sheet_expr") or "").strip()
            if sheet_expr == u"":
                if not candidate:
                    return False
                return cur in set(candidate)

            from libre_macros_lambda_column_lib import compile_lambda_column_sheet_pick

            fn_pick, glo_pick = compile_lambda_column_sheet_pick(sheet_expr)
            res = fn_pick(candidate)

            if isinstance(res, (list, tuple)):
                picked = [str(x).strip() for x in res if str(x).strip() != ""]
            elif isinstance(res, (str, unicode)):
                picked = [str(res).strip()]
            else:
                raise ValueError("sheet_expr вернул не str/list/tuple")

            picked = [x for x in picked if x != u""] if picked else []
            if not picked:
                raise ValueError("sheet_expr вернул пусто")

            # Проверка: только из candidate.
            cand_set2 = set(candidate or [])
            for p in picked:
                if p not in cand_set2:
                    raise ValueError("sheet_expr вернул недопустимое имя: %r" % p)
            return cur in set(picked)
        except Exception:
            return False

    # Основной проход по блокам (возможны множественные шаги).
    processed_any = False
    for block in blocks:
        if not isinstance(block, dict):
            continue
        if not _select_block_by_sheet_expr(block):
            continue
        processed_any = True

        # 1) Параметры блока.
        new_column = str(block.get("new_column") or "").strip()
        new_column_expr = str(block.get("new_column_expr") or "").strip()
        expr_code = str(block.get("expr") or "").strip()

        output_policy = str(block.get("output_policy") or "new").strip().casefold()
        if output_policy not in ("new", "replace"):
            output_policy = "new"

        on_error = str(block.get("on_error") or "keep").strip().casefold()
        if on_error not in ("empty", "keep", "fail"):
            on_error = "keep"

        result_type = str(block.get("result_type") or "auto").strip().casefold()
        if result_type not in ("auto", "text", "number"):
            result_type = "auto"

        # 2) new_column: литерал или lambda sheet/headers.
        if new_column_expr != u"":
            try:
                from libre_macros_lambda_column_lib import compile_lambda_column_name

                fn_name, glo_name = compile_lambda_column_name(new_column_expr)
                try:
                    res_name = fn_name(sheet_name, base_headers)
                except Exception:
                    res_name = fn_name(sheet_name, base_headers)
                new_column = str(res_name or "").strip()
            except Exception as err:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    fn_key,
                    "ошибка",
                    "new_column_expr: %s" % err,
                )
                continue
        else:
            if new_column == u"":
                _lm_log_postprocess(
                    doc, sheet_name, "диапазон", fn_key, "ошибка", "нужен new_column"
                )
                continue

        # 3) Компиляция основной лямбды.
        if expr_code == u"":
            _lm_log_postprocess(
                doc, sheet_name, "диапазон", fn_key, "ошибка", "нужен expr"
            )
            continue
        try:
            from libre_macros_lambda_column_lib import (
                compile_lambda_column_expr,
                make_rows_accessor,
                make_upd_rows_store,
                call_rows_lambda,
            )

            fn_expr_glo = compile_lambda_column_expr(expr_code)
            if fn_expr_glo is None:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    fn_key,
                    "ошибка",
                    "expr пустая/некорректная",
                )
                continue
            fn_expr, glo_expr = fn_expr_glo
        except Exception as err:
            _lm_log_postprocess(
                doc, sheet_name, "диапазон", fn_key, "ошибка", "expr: %s" % err
            )
            continue

        # 4) Разместить столбец: найти/вставить внутри data_range.
        # Заголовки могут "съехать" из-за вставок. Поэтому сканируем минимум
        # на +1 столбец вправо от h_ec.
        try:
            titles_now = []
            scan_end = int(h_ec) + 1
            c = int(h_sc)
            while c <= scan_end:
                try:
                    titles_now.append(cell_text(sheet.getCellByPosition(c, int(h_sr))).strip())
                except Exception:
                    titles_now.append(u"")
                c = c + 1
        except Exception:
            titles_now = list(base_headers or [])

        target_col = None
        try:
            key = lm_identity_key(new_column)
            if isinstance(titles_now, (list, tuple)):
                for i, t in enumerate(titles_now):
                    if lm_identity_key(t) == key:
                        target_col = int(h_sc) + int(i)
                        break
        except Exception:
            target_col = None

        position = str(block.get("position") or "_конец_").strip().casefold()
        pos_map = {
            u"_начало_": u"start",
            u"_конец_": u"end",
            u"_перед_": u"before",
            u"_после_": u"after",
            u"start": u"start",
            u"end": u"end",
            u"before": u"before",
            u"after": u"after",
        }
        position = pos_map.get(position, u"end")

        rel_token = str(block.get("relative_column") or "").strip()

        inserted = False
        insert_at = None
        if target_col is None:
            # Куда вставлять при отсутствии столбца.
            if position == u"start":
                insert_at = int(sc)
            elif position == u"end":
                insert_at = int(ec) + 1
            elif position in (u"before", u"after"):
                if rel_token == u"":
                    rel_col = int(sc)
                    rel_err = u""
                else:
                    rel_col, rel_err = _lm_pp_resolve_column_token_unique(
                        rel_token, sheet, header_row_range, sc, ec
                    )
                    if rel_col is None:
                        _lm_log_postprocess(
                            doc,
                            sheet_name,
                            "диапазон",
                            fn_key,
                            "ошибка",
                            "relative_column не найден: %s" % (rel_err or rel_token),
                        )
                        continue
                insert_at = int(rel_col) if position == u"before" else int(rel_col) + 1
            else:
                insert_at = int(ec) + 1

            # Ограничение: вставка внутри/на границе data_range.
            if insert_at < int(sc) or insert_at > int(ec) + 1:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    fn_key,
                    "ошибка",
                    "неподдерживаемая позиция вставки: %d" % insert_at,
                )
                continue

            inserted = True
            target_col = insert_at
            # Вставить физически.
            try:
                sheet.getColumns().insertByIndex(int(insert_at), 1)
            except Exception as err:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    fn_key,
                    "ошибка",
                    "вставка столбца: %s" % err,
                )
                continue

        # В v1 при output_policy=replace ожидаем существующий столбец.
        if target_col is None and output_policy == u"replace":
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_key,
                "ошибка",
                "столбец не найден для output_policy=replace",
            )
            continue

        # Обновить/написать заголовок.
        try:
            if inserted or (new_column != u"" and output_policy == u"new"):
                # при inserted заголовок точно новый;
                # при replace: если имя совпадает — заголовок не трогаем.
                if not (not inserted and output_policy == u"replace"):
                    sheet.getCellByPosition(int(target_col), int(h_sr)).String = new_column
        except Exception:
            pass

        # Донор формата (соседний столбец по ТЗ).
        ref_col = None
        if int(target_col) >= int(sc) and int(target_col) <= int(ec):
            if inserted:
                if position in (u"end", u"after"):
                    ref_col = int(target_col) - 1
                else:
                    ref_col = int(target_col) + 1
            else:
                # существующий столбец: берём левый, если есть, иначе правый
                if int(target_col) > int(sc):
                    ref_col = int(target_col) - 1
                else:
                    ref_col = int(target_col) + 1
        else:
            # На всякий случай
            ref_col = int(target_col) - 1 if int(target_col) > int(sc) else int(target_col) + 1

        if ref_col is not None and int(ref_col) >= 0 and int(target_col) >= 0:
            try:
                # По ТЗ при output_policy=replace заголовок не перекопируем.
                if inserted or output_policy != u"replace":
                    _lm_pp_copy_column_format_subrange(
                        sheet, ref_col, int(target_col), int(h_sr), int(h_sr)
                    )
                # Для удобства копируем и данные (но number format потом переопределим при необходимости).
                _lm_pp_copy_column_format_subrange(sheet, ref_col, int(target_col), int(sr), int(er))
            except Exception:
                pass

        # 5) Матрица и заголовки для rows/upd_rows.
        if inserted:
            read_sc = int(sc)
            read_ec = int(ec) + 1
        else:
            read_sc = int(sc)
            read_ec = max(int(ec), int(target_col))

        if read_ec < read_sc:
            read_ec = read_sc

        out_rng = sheet.getCellRangeByPosition(read_sc, sr, read_ec, er)
        out_data = [list(row) for row in (out_rng.getDataArray() or [])]

        expected_rows = int(er) - int(sr) + 1
        expected_cols = int(read_ec) - int(read_sc) + 1
        if expected_rows < 0:
            expected_rows = 0
        if expected_cols < 0:
            expected_cols = 0
        if len(out_data) < expected_rows:
            out_data.extend([[""] * expected_cols for _ in range(expected_rows - len(out_data))])
        elif len(out_data) > expected_rows:
            out_data = out_data[:expected_rows]

        # Нормализовать ширину строк.
        for ri in range(len(out_data)):
            if len(out_data[ri]) < expected_cols:
                out_data[ri].extend([""] * (expected_cols - len(out_data[ri])))
            elif len(out_data[ri]) > expected_cols:
                del out_data[ri][expected_cols:]

        # Заголовки под rows/accessor.
        headers = []
        c = int(read_sc)
        while c <= int(read_ec):
            try:
                headers.append(cell_text(sheet.getCellByPosition(c, int(h_sr))).strip())
            except Exception:
                headers.append(u"")
            c = c + 1

        target_rel = int(target_col) - int(read_sc)
        if target_rel < 0 or target_rel >= expected_cols:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_key,
                "ошибка",
                "не удалось определить индекс целевого столбца",
            )
            continue

        rows_accessor = make_rows_accessor(out_data, headers, start_col=read_sc)
        upd_store = make_upd_rows_store()

        # 6) Вычисления.
        import datetime

        def _is_empty_val(v):
            if v is None:
                return True
            if isinstance(v, str):
                return v.strip() == u""
            return False

        def _to_cell_value(val):
            # Сохраняем типы (число/дата) для Calc, но booleans дружелюбно в RU.
            if val is None:
                return u""
            if isinstance(val, bool):
                return u"Да" if val else u"Нет"
            return val

        def _to_cell_value_for_store(val):
            v = _to_cell_value(val)
            # setDataArray: строки/числа/дата-объекты.
            if result_type == u"text":
                # В ТЗ text всегда str.
                if v == u"":
                    return u""
                return unicode(v)
            if result_type == u"number":
                # parse строки → число, иначе пусто.
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    return v
                if v is None or v == u"":
                    return u""
                try:
                    s = unicode(v or u"").strip()
                    if s == u"":
                        return u""
                    s = s.replace(u",", u".")
                    f = float(s)
                    # целые — как int
                    if int(f) == f:
                        return int(f)
                    return f
                except Exception:
                    return u""
            if result_type == u"auto":
                return v if v is not None else u""
            return v

        first_success_val = None
        first_success_set = False
        fmt_kind = None  # 'text'|'number'|'date'

        stop_sheet = False
        error_rows = 0
        processed_rows = 0

        row_i = 0
        while row_i < len(out_data):
            cur_v = out_data[row_i][target_rel] if target_rel < len(out_data[row_i]) else u""
            try:
                v = call_rows_lambda(
                    fn_expr,
                    glo_expr,
                    row_i,
                    extra={u"idx": row_i, u"n": row_i + 1, u"text": cur_v},
                    rows_accessor=rows_accessor,
                    upd_store=upd_store,
                )
                out_data[row_i][target_rel] = _to_cell_value_for_store(v)

                if (not first_success_set) and (not _is_empty_val(v)):
                    first_success_set = True
                    first_success_val = v

                    # Определить тип для форматирования.
                    if result_type == u"auto":
                        if isinstance(v, (datetime.date, datetime.datetime)):
                            fmt_kind = u"date"
                        elif isinstance(v, (int, float)) and not isinstance(v, bool):
                            fmt_kind = u"number"
                        elif isinstance(v, bool):
                            fmt_kind = u"text"
                        else:
                            fmt_kind = u"text"
                    elif result_type == u"number":
                        fmt_kind = u"number"
                    elif result_type == u"text":
                        fmt_kind = u"text"

                processed_rows = processed_rows + 1
            except Exception as err:
                error_rows = error_rows + 1
                if on_error == u"empty":
                    out_data[row_i][target_rel] = u""
                elif on_error == u"keep":
                    out_data[row_i][target_rel] = cur_v
                else:
                    stop_sheet = True
                    break

                processed_rows = processed_rows + 1

            upd_store.commit(row_i)
            row_i = row_i + 1

        if stop_sheet:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_key,
                "fail",
                "on_error=fail, err_rows=%d" % error_rows,
            )
            continue

        # 7) Запись данных значениями (точечно, по одному столбцу).
        try:
            rng = sheet.getCellRangeByPosition(int(target_col), sr, int(target_col), er)
            col_vals = []
            for ri in range(expected_rows):
                try:
                    col_vals.append((out_data[ri][target_rel],))
                except Exception:
                    col_vals.append((u"",))
            rng.setDataArray(tuple(col_vals))
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_key,
                "ошибка",
                "setDataArray: %s" % err,
            )
            continue

        # 8) NumberFormat по first successful value (или result_type forced).
        try:
            if doc is not None and fmt_kind is not None:
                from libre_macros_global_settings_lib import load_global_settings

                gs = load_global_settings()
                if fmt_kind == u"number":
                    fmt_string = str((gs or {}).get("number_format") or u"# ##0,00").strip()
                    fmt_key = _lm_pp_number_format_key(doc, fmt_string)
                elif fmt_kind == u"date":
                    fmt_string = _lm_vlookup_default_date_format_string()
                    fmt_key = _lm_pp_number_format_key(doc, fmt_string)
                else:
                    # Text: Calc '@' (если поддерживается хостом).
                    fmt_key = _lm_pp_number_format_key(doc, u"@")

                _lm_pp_apply_column_number_format_chunked(
                    sheet,
                    int(target_col),
                    sr,
                    er,
                    fmt_key,
                    status_prefix=fn_key,
                )
        except Exception:
            pass

        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_key,
            "ok",
            "new_column=%s target_col=%d inserted=%s rows=%d ok_rows=%d err_rows=%d"
            % (
                new_column,
                int(target_col),
                str(inserted),
                int(er) - int(sr) + 1,
                processed_rows,
                error_rows,
            ),
        )


def _lm_pp_parse_number_bound(text):
    """Строка → float или None."""
    s = str(text or "").strip().replace(" ", "").replace("\xa0", "")
    if s == "":
        return None
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _lm_pp_parse_date_bound(text):
    """Строка даты → datetime.date или None."""
    import datetime

    s = str(text or "").strip()
    if s == "":
        return None
    for fmt in (
        "%d.%m.%Y",
        "%d.%m.%y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y/%m/%d",
        "%d.%m.%Y %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except Exception:
            pass
    try:
        # serial number as string
        n = float(s.replace(",", "."))
        epoch = datetime.date(1899, 12, 30)
        return epoch + datetime.timedelta(days=int(n))
    except Exception:
        return None


def _lm_pp_date_to_calc_serial(d):
    import datetime

    if d is None:
        return None
    if isinstance(d, datetime.datetime):
        d = d.date()
    epoch = datetime.date(1899, 12, 30)
    try:
        return int((d - epoch).days)
    except Exception:
        return None


def _lm_pp_validation_list_formula(items):
    parts = []
    for x in items or []:
        s = str(x or "").strip()
        if s == "":
            continue
        parts.append(s.replace('"', '""'))
    if not parts:
        return None
    return '"' + '";"'.join(parts) + '"'


def _lm_pp_apply_range_validation(sheet, col, sr, er, make_val):
    """Применить TableValidation к столбцу данных [sr..er]."""
    if sheet is None or make_val is None:
        return False
    try:
        col = int(col)
        sr = int(sr)
        er = int(er)
    except Exception:
        return False
    if er < sr:
        return False
    try:
        rng = sheet.getCellRangeByPosition(col, sr, col, er)
    except Exception:
        return False
    try:
        val = rng.Validation
    except Exception:
        val = None
    if val is None:
        try:
            val = uno.createUnoStruct("com.sun.star.sheet.TableValidation")
        except Exception:
            try:
                val = uno.createUnoStruct("com.sun.star.table.TableValidation")
            except Exception:
                return False
    try:
        make_val(val)
    except Exception:
        return False
    try:
        val.IgnoreBlankCells = True
    except Exception:
        try:
            val.IgnoreBlank = True
        except Exception:
            pass
    try:
        val.ShowErrorMessage = True
        val.ErrorMessage = u"Значение вне допустимого диапазона"
        val.ErrorTitle = u"Контроль данных"
    except Exception:
        pass
    try:
        rng.setPropertyValue("Validation", val)
        return True
    except Exception:
        try:
            rng.Validation = val
            return True
        except Exception:
            return False


def _lm_pp_set_column_list_validation(sheet, col, sr, er, items):
    formula = _lm_pp_validation_list_formula(items)
    if formula is None:
        return False
    try:
        from com.sun.star.sheet.ValidationType import LIST as VALIDATION_LIST
    except Exception:
        VALIDATION_LIST = 6
    try:
        from com.sun.star.sheet.ConditionOperator import EQUAL
    except Exception:
        EQUAL = 0

    def _make(val):
        val.Type = VALIDATION_LIST
        try:
            val.setOperator(EQUAL)
        except Exception:
            val.Operator = EQUAL
        try:
            val.setFormula1(str(formula))
        except Exception:
            val.Formula1 = str(formula)
        try:
            from com.sun.star.sheet.TableValidationVisibility import SORTEDASCENDING

            val.ShowList = SORTEDASCENDING
        except Exception:
            try:
                val.ShowList = 2
            except Exception:
                pass

    return _lm_pp_apply_range_validation(sheet, col, sr, er, _make)


def _lm_pp_set_column_decimal_validation(sheet, col, sr, er, min_v, max_v):
    try:
        from com.sun.star.sheet.ValidationType import DECIMAL as VALIDATION_DECIMAL
    except Exception:
        VALIDATION_DECIMAL = 2
    try:
        from com.sun.star.sheet.ConditionOperator import BETWEEN, GREATER_EQUAL, LESS_EQUAL
    except Exception:
        BETWEEN, GREATER_EQUAL, LESS_EQUAL = 1, 3, 5
    has_min = min_v is not None
    has_max = max_v is not None
    if not has_min and not has_max:
        return False

    def _make(val):
        val.Type = VALIDATION_DECIMAL
        if has_min and has_max:
            op = BETWEEN
            f1 = str(min_v)
            f2 = str(max_v)
        elif has_min:
            op = GREATER_EQUAL
            f1 = str(min_v)
            f2 = ""
        else:
            op = LESS_EQUAL
            f1 = str(max_v)
            f2 = ""
        try:
            val.setOperator(op)
        except Exception:
            val.Operator = op
        try:
            val.setFormula1(f1)
        except Exception:
            val.Formula1 = f1
        if f2 != "":
            try:
                val.setFormula2(f2)
            except Exception:
                val.Formula2 = f2

    return _lm_pp_apply_range_validation(sheet, col, sr, er, _make)


def _lm_pp_set_column_date_validation(sheet, col, sr, er, min_d, max_d):
    try:
        from com.sun.star.sheet.ValidationType import DATE as VALIDATION_DATE
    except Exception:
        VALIDATION_DATE = 3
    try:
        from com.sun.star.sheet.ConditionOperator import BETWEEN, GREATER_EQUAL, LESS_EQUAL
    except Exception:
        BETWEEN, GREATER_EQUAL, LESS_EQUAL = 1, 3, 5
    smin = _lm_pp_date_to_calc_serial(min_d)
    smax = _lm_pp_date_to_calc_serial(max_d)
    has_min = smin is not None
    has_max = smax is not None
    if not has_min and not has_max:
        return False

    def _make(val):
        val.Type = VALIDATION_DATE
        if has_min and has_max:
            op = BETWEEN
            f1 = str(smin)
            f2 = str(smax)
        elif has_min:
            op = GREATER_EQUAL
            f1 = str(smin)
            f2 = ""
        else:
            op = LESS_EQUAL
            f1 = str(smax)
            f2 = ""
        try:
            val.setOperator(op)
        except Exception:
            val.Operator = op
        try:
            val.setFormula1(f1)
        except Exception:
            val.Formula1 = f1
        if f2 != "":
            try:
                val.setFormula2(f2)
            except Exception:
                val.Formula2 = f2

    return _lm_pp_apply_range_validation(sheet, col, sr, er, _make)


def lm_pp_range_add_column(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Добавить пустой столбец: имя, тип (text/number/date), формат, список/мин–макс.

    JSON: [{"v":1,"fn":"добавить_столбец","new_column":"Статус","data_type":"text",
            "choices":["А","Б"],"format":"@"}]
    """
    fn_key = "добавить_столбец"
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet(fn_key, extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(fn_key, extra_args, doc, sheet)
    if block is None:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_key, "пропуск", "нет блока JSON для листа"
        )
        return
    new_column = str(block.get("new_column") or block.get("name") or "").strip()
    if new_column == "":
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_key, "ошибка", "не задано new_column"
        )
        return
    data_type = str(block.get("data_type") or "text").strip().casefold()
    if data_type not in ("text", "number", "date"):
        data_type = "text"
    fmt = str(block.get("format") or "").strip()
    choices = block.get("choices") or block.get("values") or []
    if not isinstance(choices, (list, tuple)):
        choices = [choices]
    choice_list = []
    for c in choices:
        s = str(c or "").strip()
        if s != "":
            choice_list.append(s)
    data_validation = bool(
        block.get("data_validation")
        or block.get("validation")
        or block.get("контроль_данных")
    )
    min_raw = str(block.get("min") or block.get("min_value") or "").strip()
    max_raw = str(block.get("max") or block.get("max_value") or "").strip()

    h_sr = sr
    try:
        if header_row_range is not None:
            _a, h_sr, _b, _c = _lm_pp_range_address(header_row_range)
    except Exception:
        h_sr = max(0, int(sr) - 1)

    existing = _lm_pp_find_or_append_header_column(sheet, new_column, h_sr)
    if existing < 0:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_key, "ошибка", "не удалось найти место столбца"
        )
        return
    # Уже есть заголовок с этим именем?
    existed = False
    try:
        existed = (
            lm_identity_key(cell_text(sheet.getCellByPosition(existing, h_sr)).strip())
            == lm_identity_key(new_column)
        )
    except Exception:
        existed = False
    target_col = existing
    inserted = False
    if not existed:
        # Пустая ячейка справа — пишем заголовок; иначе вставка в конец used.
        try:
            cur = cell_text(sheet.getCellByPosition(existing, h_sr)).strip()
        except Exception:
            cur = ""
        if cur == "":
            target_col = existing
            try:
                sheet.getCellByPosition(target_col, h_sr).String = new_column
            except Exception:
                pass
            inserted = True
        else:
            insert_at = int(ec) + 1
            try:
                sheet.getColumns().insertByIndex(insert_at, 1)
                target_col = insert_at
                sheet.getCellByPosition(target_col, h_sr).String = new_column
                inserted = True
            except Exception as err:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    fn_key,
                    "ошибка",
                    "вставка столбца: %s" % err,
                )
                return
    else:
        target_col = existing

    ref_col = max(0, int(sc))
    try:
        _lm_pp_finalize_derived_column(
            sheet, ref_col, target_col, h_sr, sr, doc=doc
        )
    except Exception:
        pass

    # NumberFormat по типу / format
    if doc is not None:
        try:
            if fmt == "":
                if data_type == "number":
                    try:
                        from libre_macros_global_settings_lib import load_global_settings

                        gs = load_global_settings()
                        fmt = str((gs or {}).get("number_format") or "# ##0,00").strip()
                    except Exception:
                        fmt = "# ##0,00"
                elif data_type == "date":
                    try:
                        fmt = _lm_vlookup_default_date_format_string()
                    except Exception:
                        fmt = "DD.MM.YYYY"
                else:
                    fmt = "@"
            fmt_key = _lm_pp_number_format_key(doc, fmt)
            data_top = int(sr)
            data_bot = int(er) if int(er) >= data_top else data_top
            # заголовок тоже форматируем для даты/числа? обычно только данные
            _lm_pp_apply_column_number_format_chunked(
                sheet, int(target_col), data_top, data_bot, fmt_key, status_prefix=fn_key
            )
        except Exception:
            pass

    # Контроль данных
    dv_ok = False
    if data_type == "text" and choice_list:
        dv_ok = _lm_pp_set_column_list_validation(
            sheet, target_col, sr, er if er >= sr else sr, choice_list
        )
    elif data_type == "number" and data_validation:
        vmin = _lm_pp_parse_number_bound(min_raw)
        vmax = _lm_pp_parse_number_bound(max_raw)
        dv_ok = _lm_pp_set_column_decimal_validation(
            sheet, target_col, sr, er if er >= sr else sr, vmin, vmax
        )
    elif data_type == "date" and data_validation:
        dmin = _lm_pp_parse_date_bound(min_raw)
        dmax = _lm_pp_parse_date_bound(max_raw)
        dv_ok = _lm_pp_set_column_date_validation(
            sheet, target_col, sr, er if er >= sr else sr, dmin, dmax
        )

    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        fn_key,
        "ok",
        "new_column=%s type=%s col=%d inserted=%s validation=%s"
        % (new_column, data_type, int(target_col), str(inserted), str(dv_ok)),
    )


def lm_pp_range_apply_formula(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Добавить столбец с формулами Calc: шаблон в первой строке данных,
        затем заполнение вниз с автосмещением ссылок.

    Параметры:
        doc, sheet, data_range, header_row_range — стандартные range
        *extra_args — dict-блок JSON (param_decode)

    JSON колонки C: [{"v":1,"fn":"применить_формулу","column":"Бонус","formula":"=D{row_id}*0.1","as_values":false}]
    Плейсхолдеры: {row_id}, {col_id}, R[-1]C. См. docs/13_JSON_PARAMS.md.

    Карта: «применить_формулу».

    Когда вызывать:
        После сбора данных и «преобразовать_числа», если в формуле участвуют числовые столбцы.

    Что меняет:
        Новый столбец: заголовок; Formula в первой строке + fillAuto вниз;
        при as_values — Copy→InsertContents (SVD) по столбцу (значения без формул);
        клон формата, OptimalWidth, расширение end_col; очистка undo.
    """
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("применить_формулу", extra_args, doc, sheet):
        return
    parsed = _lm_pp_apply_formula_spec_for_sheet(extra_args, doc, sheet)
    if parsed is None:
        return
    if len(parsed) >= 4:
        col_name, formula_text, as_values, fmt = parsed[0], parsed[1], parsed[2], parsed[3]
    else:
        col_name, formula_text, as_values = parsed[0], parsed[1], parsed[2]
        fmt = ""
    fmt = str(fmt or "").strip()
    loc = _lm_pp_resolve_derived_column(
        sheet, data_range, header_row_range, col_name
    )
    if loc is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "применить_формулу",
            "ошибка",
            "не удалось добавить столбец «%s»" % col_name,
        )
        return
    sr = loc["sr"]
    er = loc["er"]
    h_sr = loc["h_sr"]
    ref_col = loc["ref_col"]
    target_col = loc["target_col"]
    # Явный format из JSON — приоритет; иначе клон с соседнего столбца.
    if fmt == "":
        _lm_pp_clone_concat_column_formats(
            sheet, ref_col, target_col, h_sr, sr, er
        )
    sheet.getCellByPosition(target_col, h_sr).String = col_name
    if as_values:
        # Для as_values не зависим от fillAuto/Copy: пишем формулы построчно и сразу материализуем.
        if not _lm_pp_apply_formula_put_formulas_for_as_values(
            doc,
            sheet,
            formula_text,
            target_col,
            sr,
            er,
            sheet_name,
            header_row=h_sr,
        ):
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                "применить_формулу",
                "ошибка",
                "не удалось записать/материализовать значения в столбец «%s»" % col_name,
            )
            return
    else:
        if _lm_pp_formula_uses_source_variables(formula_text):
            ok = _lm_pp_write_formula_column_per_row(
                doc,
                sheet,
                formula_text,
                target_col,
                sr,
                er,
                sheet_name=sheet_name,
                header_row=h_sr,
            )
        else:
            ok = _lm_pp_write_formula_column(
                doc,
                sheet,
                formula_text,
                target_col,
                sr,
                er,
                sheet_name,
                header_row=h_sr,
            )
        if not ok:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                "применить_формулу",
                "ошибка",
                "не удалось записать формулы в столбец «%s»" % col_name,
            )
            return
    total_rows = int(er) - int(sr) + 1
    formula_cells = _lm_pp_count_formula_cells(sheet, target_col, sr, er)
    _lm_pp_apply_formula_dbg(
        doc,
        sheet_name,
        "formula_cells=%d/%d rows %s..%s"
        % (formula_cells, total_rows, sr, er),
    )
    try:
        probe = sheet.getCellByPosition(target_col, sr)
        _lm_pp_apply_formula_dbg(
            doc,
            sheet_name,
            "проверка row %s: type=%s formula=%r"
            % (
                sr,
                probe.getType(),
                str(probe.getFormula() or probe.Formula or "")[:100],
            ),
        )
    except Exception:
        pass
    if (not as_values) and formula_cells < total_rows:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "применить_формулу",
            "ошибка",
            "формулы в столбце «%s»: %d/%d строк"
            % (col_name, formula_cells, total_rows),
        )
        return
    try:
        hdr_row = int(loc.get("h_sr", 0))
    except Exception:
        hdr_row = 0
    calc_formula = _lm_pp_expand_apply_formula_template(
        formula_text, sr, target_col, doc=doc, sheet=sheet, header_row=hdr_row
    )
    try:
        sample_local = str(
            sheet.getCellByPosition(target_col, sr).FormulaLocal or ""
        ).strip()
    except Exception:
        sample_local = ""
    if sample_local:
        _lm_pp_apply_formula_dbg(
            doc,
            sheet_name,
            "образец FormulaLocal row %s: %r" % (sr, sample_local[:120]),
        )
    if as_values:
        _lm_pp_apply_formula_dbg(doc, sheet_name, "as_values=ON (materialized)")
    else:
        _lm_pp_apply_formula_dbg(doc, sheet_name, "as_values=OFF")
    _lm_pp_undo_clear(doc, "применить_формулу")
    _lm_pp_finalize_derived_column(sheet, ref_col, target_col, h_sr, sr, doc=doc)
    # NumberFormat после finalize: иначе клон с соседа (NumberFormat в props) затирает явный format.
    if fmt != "" and doc is not None and er >= sr:
        try:
            fmt_key = _lm_pp_number_format_key(doc, fmt)
            _lm_pp_apply_column_number_format_chunked(
                sheet, target_col, sr, er, fmt_key, status_prefix="применить_формулу"
            )
        except Exception as fmt_err:
            _lm_pp_apply_formula_dbg(
                doc, sheet_name, "format apply FAIL: %s" % fmt_err
            )
    _lm_pp_autofit_columns(doc, sheet, [target_col], sr=h_sr, er=er)
    note_fmt = (", format «%s»" % fmt) if fmt else ""
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        "применить_формулу",
        "ok",
        "столбец %d «%s», образец %d, формула «%s»%s%s"
        % (
            target_col,
            col_name,
            ref_col,
            calc_formula,
            ", как значение" if as_values else "",
            note_fmt,
        ),
    )


def lm_pp_range_color_scale_simple(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Градиент фона числового столбца от min к max (карта: «градиент»).

    Колонка C — маркер столбца, цвет min, цвет max:
      «сумм, зеленый, красный»;
      «сумм -> зеленый -> красный»;
      «amount / #00FF00 / #FF0000».
    JSON: marker, color_min, color_max, sheet, whole_row, sort, sort_dir.
    Пусто — сумм, зелёный (min) → красный (max), как раньше.
    """
    sheet_name = sheet.Name if sheet is not None else ""
    raw_head = _lm_pp_extra_args_rest_text(extra_args)[:300]
    _lm_pp_gradient_log(
        doc,
        sheet_name,
        "start v=%s C=%r" % (MACRO_VERSION, raw_head),
    )
    # JSON не конвертируем в legacy-текст: block_to_executor_param_text
    # кодирует только marker/color_min/color_max и теряет whole_row/sort.
    spec = _lm_pp_color_scale_spec_for_sheet(extra_args, doc, sheet)
    if spec is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "градиент",
            "инфо",
            "пропуск: параметры не для этого листа",
        )
        return
    value_column_marker = spec["marker"]
    color_min = spec["min_color"]
    color_max = spec["max_color"]
    whole_row = bool(spec.get("whole_row"))
    sort_enabled = bool(spec.get("sort"))
    sort_ascending = bool(spec.get("sort_ascending", True))
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    paint_sc = sc
    paint_ec = ec
    if whole_row:
        paint_sc = 0
        used_ec, _used_er = get_sheet_used_bounds(sheet)
        if used_ec > paint_ec:
            paint_ec = used_ec
    _lm_pp_gradient_log(
        doc,
        sheet_name,
        "spec marker=%r whole_row=%s sort=%s sort_asc=%s paint_cols=%d..%d rows=%d..%d"
        % (
            value_column_marker,
            whole_row,
            sort_enabled,
            sort_ascending,
            paint_sc,
            paint_ec,
            sr,
            er,
        ),
    )
    titles = _lm_pp_header_titles(sheet, header_row_range)
    target_col, matched_marker = _lm_pp_gradient_find_column(
        titles, sc, ec, value_column_marker
    )

    if target_col is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "градиент",
            "инфо",
            "столбец не найден по маркеру «%s»" % value_column_marker,
        )
        return

    if sort_enabled:
        _lm_pp_gradient_dbg("sort", "before sort_args=%r" % (matched_marker,))
        sort_args = (
            "%s -> %s" % (matched_marker, "+" if sort_ascending else "-"),
        )
        sort_ok, sort_err = _lm_pp_range_sort_data_core(
            doc,
            sheet,
            data_range,
            header_row_range,
            sort_args,
            fn_label="градиент",
        )
        if not sort_ok:
            _lm_pp_gradient_dbg("sort", "FAILED err=%r" % (sort_err,))
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                "градиент",
                "ошибка",
                "сортировка не выполнена: %s" % (sort_err or "—"),
            )
            return
        _lm_pp_gradient_dbg("sort", "ok")

    values = []
    r = sr
    n = 0
    while r <= er:
        cell = sheet.getCellByPosition(target_col, r)
        try:
            val = float(cell.Value)
            values.append((r, val))
        except (ValueError, TypeError):
            pass
        n = n + 1
        if _lm_ui_should_update_status(n, is_last=(r == er)):
            _lm_ui_status_set(
                n,
                _lm_pp_status_text("диапазон", "градиент: чтение %d/%d" % (r + 1, er + 1)),
            )
        _lm_ui_yield(counter=n)
        r = r + 1

    if not values:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "градиент",
            "инфо",
            "нет числовых значений в столбце «%s»" % matched_marker,
        )
        return

    min_val = min(v[1] for v in values)
    max_val = max(v[1] for v in values)
    val_range = max_val - min_val if max_val != min_val else 1

    vi = 0
    while vi < len(values):
        row_idx, val = values[vi]
        norm = (val - min_val) / val_range if val_range != 0 else 0.5
        bg = _lm_pp_interpolate_color(color_min, color_max, norm)
        if whole_row:
            if vi == 0:
                _lm_pp_gradient_dbg(
                    "paint",
                    "whole_row cols %d..%d row=%d" % (paint_sc, paint_ec, row_idx),
                )
            cc = paint_sc
            while cc <= paint_ec:
                sheet.getCellByPosition(cc, row_idx).CellBackColor = bg
                cc = cc + 1
        else:
            if vi == 0:
                _lm_pp_gradient_dbg(
                    "paint",
                    "single col=%d row=%d (whole_row=False)" % (target_col, row_idx),
                )
            sheet.getCellByPosition(target_col, row_idx).CellBackColor = bg
        vi = vi + 1
        if _lm_ui_should_update_status(vi, is_last=(vi >= len(values))):
            _lm_ui_status_set(
                vi,
                _lm_pp_status_text(
                    "диапазон", "градиент: раскраска %d/%d" % (vi, len(values))
                ),
            )
        _lm_ui_yield(counter=vi)
    extras = []
    if whole_row:
        extras.append("вся строка")
    if sort_enabled:
        extras.append("сорт. %s" % ("возр" if sort_ascending else "убыв"))
    extra_note = (", " + ", ".join(extras)) if extras else ""
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        "градиент",
        "ok",
        "маркер «%s», столбец %d, ячеек %d%s"
        % (matched_marker, target_col, len(values), extra_note),
    )


def lm_pp_row_copy_format_from_header(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
    Назначение:
        Скопировать оформление ячеек заголовка в текущую строку данных (по столбцам).

    Параметры:
        sheet — XSpreadsheet
        data_row_range — XCellRange одной строки
        header_row_range — XCellRange строки заголовков
        row_index — не используется (строка из data_row_range)

    JSON колонки C: [{"v":1,"fn":"скрыть_служебные"}] без доп. полей или пусто.

    Карта: «копировать_формат_заголовка».

    Когда вызывать:
        Если данные должны визуально повторять заголовок; конфликтует с «зебра».

    Что меняет:
        Шрифт, заливку, выравнивание каждой ячейки строки (не значения).
    """
    sc, sr, ec, er = _lm_pp_range_address(data_row_range)
    _hsc, hsr, _hec, _her = _lm_pp_range_address(header_row_range)

    c = sc
    while c <= ec:
        source_cell = sheet.getCellByPosition(c, hsr)
        target_cell = sheet.getCellByPosition(c, sr)

        # Побочно копируем только свойства формата, не Formula/String
        target_cell.CharFontName = source_cell.CharFontName
        target_cell.CharHeight = source_cell.CharHeight
        target_cell.CharWeight = source_cell.CharWeight
        target_cell.CellBackColor = source_cell.CellBackColor
        _lm_set_char_color_safe(target_cell, source_cell.CharColor)
        target_cell.HoriJustify = source_cell.HoriJustify
        target_cell.VertJustify = source_cell.VertJustify
        
        c = c + 1


def lm_pp_row_alternate_borders(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
    Назначение:
        Тонкая верхняя граница у нечётных строк данных (визуальное разделение).

    Параметры:
        data_row_range — XCellRange одной строки
        row_index — int (чётные строки пропускаются)

    JSON колонки C: [{"v":1,"fn":"скрыть_служебные"}] без доп. полей или пусто.

    Карта: «чередующиеся_границы».

    Когда вызывать:
        Вместе с «тонкая_сетка» или вместо «зебра»; на каждой строке.

    Что меняет:
        TableBorder.TopLine у диапазона строки (нечётные row_index).
    """
    if row_index % 2 == 0:
            return

    sc, sr, ec, er = _lm_pp_range_address(data_row_range)
    line = _lm_pp_border_line(color=0xD9D9D9, width=5)

    # Только верхняя линия строки — лёгкий разделитель между блоками
    tb = uno.createUnoStruct("com.sun.star.table.TableBorder")
    tb.IsTopLineValid = True
    tb.TopLine = line
    
    row_range = sheet.getCellRangeByPosition(sc, sr, ec, sr)
    row_range.TableBorder = tb


def lm_pp_row_text_color_by_value(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):
    """
    Назначение:
        Покрасить текст всей строки по значению в столбце «статус» (или другом по заголовку).

    Параметры:
        sheet, data_row_range, header_row_range, row_index — стандартные row
        *extra_args — dict-блок JSON (param_decode)

    JSON колонки C: [{"v":1,"fn":"цвет_текста_по_значению","marker":"статус"}].

    Карта: «цвет_текста_по_значению».

    Когда вызывать:
        После заполнения статусов; на каждой строке.

    Что меняет:
        CharColor всей data_row_range при вхождении ключа color_map в текст ячейки столбца.
    """
    if _lm_pp_param_skip_sheet("цвет_текста_по_значению", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(
        "цвет_текста_по_значению", extra_args, doc, sheet
    )
    if block is not None:
        header_marker = str(block.get("marker") or "статус")
    else:
        header_marker = extra_args[0] if len(extra_args) > 0 else "статус"
    # Карта цветов (встроенная)
    color_map = {
        "ок": (0, 128, 0),
        "ok": (0, 128, 0),
        "да": (0, 128, 0),
        "yes": (0, 128, 0),
        "ошибка": (255, 0, 0),
        "error": (255, 0, 0),
        "нет": (255, 0, 0),
        "no": (255, 0, 0),
    }
    
    sc, sr, ec, er = _lm_pp_range_address(data_row_range)
    titles = _lm_pp_header_titles(sheet, header_row_range)
    needle = str(header_marker).lower()
    
    # Находим столбец
    target_col = None
    c = sc
    while c <= ec:
        idx = c - sc
        title = titles[idx].lower() if idx < len(titles) else ""
        if needle in title:
            target_col = c
            break
        c = c + 1
    
    if target_col is None:
            return

    cell_value = cell_text(sheet.getCellByPosition(target_col, sr)).lower()

    # Первое вхождение ключа из color_map — цвет на всю строку
    for key, rgb in color_map.items():
        if key in cell_value:
            # Применяем ко всей строке
            row_range = sheet.getCellRangeByPosition(sc, sr, ec, sr)
            _lm_set_char_color_safe(row_range, lo_color_rgb(*rgb))
            break


def _lm_pp_outline_group_rows(sheet, start_row, end_row, sc, ec):
    """Outline-группа строк [start_row..end_row] через XSheetOutline / OutlineLevel."""
    if sheet is None or end_row < start_row:
        return False, "пустой диапазон"
    # 1) XSheetOutline.group + createUnoStruct (CellRangeAddress() в Python часто «пустой»).
    try:
        addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
        addr.Sheet = int(_lm_pp_sheet_index(sheet))
        if addr.Sheet < 0:
            addr.Sheet = 0
        addr.StartColumn = int(sc)
        addr.StartRow = int(start_row)
        addr.EndColumn = int(ec)
        addr.EndRow = int(end_row)
        try:
            from com.sun.star.table import TableOrientation

            sheet.group(addr, TableOrientation.ROWS)
            return True, ""
        except Exception:
            sheet.group(addr, 1)  # ROWS
            return True, ""
    except Exception as err1:
        err_group = unicode(err1)
    else:
        err_group = u""
    # 2) Fallback: уровень outline на строках (даёт плюсики слева).
    try:
        rows = sheet.getRows()
        r = int(start_row)
        while r <= int(end_row):
            rows.getByIndex(r).OutlineLevel = 1
            r = r + 1
        return True, "OutlineLevel"
    except Exception as err2:
        return False, u"group=%s; OutlineLevel=%s" % (err_group, unicode(err2))


def _lm_pp_outline_auto_from_formulas(sheet, sc, sr, ec, er):
    """autoOutline по формулам SUBTOTAL/ИТОГ — как у ручных промежуточных итогов."""
    if sheet is None or er < sr:
        return False
    try:
        addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
        addr.Sheet = int(_lm_pp_sheet_index(sheet))
        if addr.Sheet < 0:
            addr.Sheet = 0
        addr.StartColumn = int(sc)
        addr.StartRow = int(sr)
        addr.EndColumn = int(ec)
        addr.EndRow = int(er)
        sheet.autoOutline(addr)
        return True
    except Exception:
        return False


def _lm_pp_group_agg_subtotal_num(agg_fn):
    """Номер функции SUBTOTAL (Calc/Excel) для промежуточного итога."""
    try:
        from libre_macros_param_wizard_cfg import GROUP_AGG_SUBTOTAL_NUM

        return int(GROUP_AGG_SUBTOTAL_NUM.get(str(agg_fn or "sum").casefold(), 9))
    except Exception:
        m = {
            "average": 1,
            "count": 3,
            "max": 4,
            "min": 5,
            "product": 6,
            "sum": 9,
        }
        return int(m.get(str(agg_fn or "sum").casefold(), 9))


def _lm_pp_group_subtotal_formula(agg_fn, col_0, row_start_0, row_end_0):
    """
    Формула промежуточного итога в стиле ручного Data→Subtotals (RU):
      =ИТОГ(9;$D$2:$D$9)
    Разделитель «;», абсолютные ссылки — иначе Err:509 в русской локали.
    """
    letters = col_index_to_letters(int(col_0))
    n = _lm_pp_group_agg_subtotal_num(agg_fn)
    r1 = int(row_start_0) + 1
    r2 = int(row_end_0) + 1
    # FormulaLocal (русские имена) — как у пользователя вручную.
    return u"=ИТОГ(%d;$%s$%d:$%s$%d)" % (n, letters, r1, letters, r2)


def _lm_pp_group_write_subtotal_formula(doc, cell, agg_fn, col_0, row_start_0, row_end_0, sheet_name=""):
    """Записать ИТОГ/SUBTOTAL; сначала FormulaLocal с «;», затем EN-запас."""
    local_f = _lm_pp_group_subtotal_formula(agg_fn, col_0, row_start_0, row_end_0)
    if _lm_pp_try_set_cell_formula(
        cell, local_f, use_local=True, allow_formula_error=False
    ):
        return True
    letters = col_index_to_letters(int(col_0))
    n = _lm_pp_group_agg_subtotal_num(agg_fn)
    r1 = int(row_start_0) + 1
    r2 = int(row_end_0) + 1
    # API: английское имя, «;» (типично для RU-дока через Formula).
    api_semi = u"=SUBTOTAL(%d;$%s$%d:$%s$%d)" % (n, letters, r1, letters, r2)
    if _lm_pp_set_cell_formula(
        doc,
        cell,
        api_semi,
        row_0=int(row_end_0) + 1 if row_end_0 is not None else None,
        col_0=col_0,
        sheet_name=sheet_name,
        allow_formula_error=False,
    ):
        return True
    # Запас: запятая (EN-локаль).
    api_comma = u"=SUBTOTAL(%d,$%s$%d:$%s$%d)" % (n, letters, r1, letters, r2)
    return bool(
        _lm_pp_set_cell_formula(
            doc,
            cell,
            api_comma,
            row_0=int(row_end_0) + 1 if row_end_0 is not None else None,
            col_0=col_0,
            sheet_name=sheet_name,
            allow_formula_error=False,
        )
    )


def _lm_pp_group_agg_fn_label(agg_fn):
    try:
        from libre_macros_param_wizard_cfg import GROUP_AGG_FN_LABEL

        return unicode(GROUP_AGG_FN_LABEL.get(str(agg_fn or "sum").casefold(), u"Сумма"))
    except Exception:
        return u"Сумма"


def _lm_pp_group_agg_grand_label(agg_fn):
    try:
        from libre_macros_param_wizard_cfg import GROUP_AGG_GRAND_LABEL

        return unicode(
            GROUP_AGG_GRAND_LABEL.get(
                str(agg_fn or "sum").casefold(), u"Общая сумма"
            )
        )
    except Exception:
        return u"Общая сумма"


def _lm_pp_group_is_total_row_label(text):
    """True, если ячейка — подпись нашего промежуточного/общего итога (не данные)."""
    s = unicode(text or u"").strip()
    if s == u"":
        return False
    low = s.casefold()
    if low.startswith(u"общ") or low.startswith(u"grand") or low.startswith(u"итого"):
        return True
    labels = [
        u"сумма",
        u"количество",
        u"среднее",
        u"максимум",
        u"минимум",
        u"произведение",
    ]
    try:
        from libre_macros_param_wizard_cfg import GROUP_AGG_FN_LABEL, GROUP_AGG_GRAND_LABEL

        for lab in list(GROUP_AGG_FN_LABEL.values()) + list(GROUP_AGG_GRAND_LABEL.values()):
            t = unicode(lab or u"").strip()
            if t:
                labels.append(t.casefold())
    except Exception:
        pass
    for lab in labels:
        if low == lab or low.endswith(u" " + lab):
            return True
    return False


def _lm_pp_group_style_total_row(sheet, sc, ec, row):
    """Жирный + подчёркнутый для строки промежуточного/общего итога."""
    try:
        rng = sheet.getCellRangeByPosition(int(sc), int(row), int(ec), int(row))
        rng.CharWeight = _LM_FMT_CHAR_WEIGHT_BOLD
        try:
            rng.CharUnderline = 1  # SINGLE
        except Exception:
            pass
    except Exception:
        pass


def _lm_pp_group_resolve_agg_levels(block, sheet, header_row_range, sc, ec):
    """
    Список уровней итога: [{fn, col_indices, label}, …].
    Один уровень может писать формулы в несколько столбцов.
    """
    levels = []
    aggs = block.get("aggs") if isinstance(block, dict) else None
    if not isinstance(aggs, (list, tuple)) or not aggs:
        aggs = [
            {
                u"fn": block.get("agg_fn") or u"sum",
                u"columns": block.get("agg_columns") or [],
            }
        ]
    for item in aggs:
        if not isinstance(item, dict):
            continue
        fn = unicode(item.get("fn") or u"sum").strip().casefold() or u"sum"
        toks = item.get("columns") or []
        if not isinstance(toks, (list, tuple)):
            toks = [toks]
        cols = []
        for tok in toks:
            ci = _lm_pp_resolve_column_token(tok, sheet, header_row_range, sc, ec)
            if ci is not None and ci not in cols:
                cols.append(ci)
        if not cols:
            continue
        levels.append(
            {
                u"fn": fn,
                u"cols": cols,
                u"label": _lm_pp_group_agg_fn_label(fn),
                u"grand_label": _lm_pp_group_agg_grand_label(fn),
            }
        )
    return levels


def lm_pp_range_group_by_column(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Промежуточные итоги по смене значения в столбце + опциональный outline.
        Несколько уровней (count+sum и т.п.) — поле aggs[] или склейка подряд
        идущих шагов группировка_по_столбцу с одним marker.

    JSON:
      [{"v":1,"fn":"группировка_по_столбцу","marker":"'Отдел'",
        "aggs":[{"fn":"count","columns":["'Сумма_руб'"]},
                {"fn":"sum","columns":["'Сумма_руб'"]}],
        "sort":true,"outline":true,"grand_total":true}]
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    sheet_name = ""
    try:
        sheet_name = str(sheet.Name) if sheet is not None else ""
    except Exception:
        pass

    if _lm_pp_param_skip_sheet("группировка_по_столбцу", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(
        "группировка_по_столбцу", extra_args, doc, sheet
    )
    if block is None and extra_args and isinstance(extra_args[0], dict):
        block = extra_args[0]
    if block is None:
        block = {
            "marker": extra_args[0] if len(extra_args) > 0 else "группа",
            "agg_fn": "sum",
            "agg_columns": [],
            "sort": True,
            "outline": True,
            "grand_total": True,
        }
    try:
        from libre_macros_param_codec import _normalize_group_by_column_block

        block = _normalize_group_by_column_block(block)
    except Exception:
        pass

    group_column_marker = block.get("marker")
    if isinstance(group_column_marker, (list, tuple)):
        group_column_marker = (
            group_column_marker[0] if group_column_marker else "группа"
        )
    if group_column_marker in (None, ""):
        group_column_marker = "группа"

    group_col = _lm_pp_resolve_column_token(
        group_column_marker, sheet, header_row_range, sc, ec
    )
    if group_col is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "группировка_по_столбцу",
            unicode(group_column_marker),
            "ошибка",
            "столбец группировки не найден",
        )
        return

    do_sort = bool(block.get("sort", True))
    do_outline = bool(block.get("outline", True))
    do_grand = bool(block.get("grand_total", True))
    levels = _lm_pp_group_resolve_agg_levels(
        block, sheet, header_row_range, sc, ec
    )
    do_subtotals = len(levels) > 0
    n_levels = len(levels)

    if er < sr:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "группировка_по_столбцу",
            unicode(group_column_marker),
            "пропуск",
            "нет строк данных",
        )
        return

    # Повторный прогон / грязный лист: убрать уже вставленные строки итогов,
    # иначе сортировка поднимает «Общее…» наверх и формулы плывут.
    removed_old = 0
    r_del = int(er)
    while r_del >= int(sr):
        try:
            lab = cell_text(sheet.getCellByPosition(group_col, r_del))
        except Exception:
            lab = u""
        if _lm_pp_group_is_total_row_label(lab):
            try:
                sheet.getRows().removeByIndex(r_del, 1)
                removed_old = removed_old + 1
                er = int(er) - 1
            except Exception:
                pass
        r_del = r_del - 1
    if er < sr:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "группировка_по_столбцу",
            unicode(group_column_marker),
            "пропуск",
            "нет строк данных после очистки итогов",
        )
        return

    if do_sort:
        ok_sort, err_sort = _lm_pp_range_sort_dispatch(
            doc,
            sheet,
            sc,
            sr,
            ec,
            er,
            [(group_col, True)],
            sc,
            sheet_name=sheet_name,
            fn_label="группировка_по_столбцу/sort",
        )
        if not ok_sort:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "группировка_по_столбцу",
                unicode(group_column_marker),
                "ошибка",
                "сортировка: %s" % (err_sort or "не выполнена"),
            )
            return

    groups = []
    prev_value = None
    group_start = None
    last_data_row = None
    r = sr
    while r <= er:
        current_value = cell_text(sheet.getCellByPosition(group_col, r))
        # Уже вставленные строки итогов не считаем данными (защита от второго прохода).
        if _lm_pp_group_is_total_row_label(current_value):
            r = r + 1
            continue
        if prev_value is not None and current_value != prev_value:
            groups.append((group_start, last_data_row, prev_value))
            group_start = r
        if prev_value is None:
            group_start = r
        prev_value = current_value
        last_data_row = r
        r = r + 1
    if prev_value is not None and group_start is not None and last_data_row is not None:
        groups.append((group_start, last_data_row, prev_value))

    if not groups:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "группировка_по_столбцу",
            unicode(group_column_marker),
            "пропуск",
            "нет групп",
        )
        return

    inserted = 0
    formula_ok = 0
    formula_fail = 0
    outlined = 0
    outline_err = u""
    # Снизу вверх: вставка N строк итога после группы.
    gi = len(groups) - 1
    while gi >= 0:
        g_start, g_end, g_val = groups[gi]
        g_name = g_val if unicode(g_val).strip() != u"" else u"(пусто)"
        if do_subtotals and n_levels > 0:
            insert_at = int(g_end) + 1
            try:
                sheet.getRows().insertByIndex(insert_at, n_levels)
            except Exception as err:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "группировка_по_столбцу",
                    unicode(group_column_marker),
                    "ошибка",
                    "insertRows: %s" % err,
                )
                return
            inserted = inserted + n_levels
            li = 0
            while li < n_levels:
                row = insert_at + li
                level = levels[li]
                label = u"%s %s" % (g_name, level[u"label"])
                try:
                    lab_cell = sheet.getCellByPosition(group_col, row)
                    lab_cell.String = label
                except Exception:
                    pass
                for ac in level[u"cols"]:
                    cell = sheet.getCellByPosition(ac, row)
                    if _lm_pp_group_write_subtotal_formula(
                        doc,
                        cell,
                        level[u"fn"],
                        ac,
                        g_start,
                        g_end,
                        sheet_name=sheet_name,
                    ):
                        formula_ok = formula_ok + 1
                    else:
                        formula_fail = formula_fail + 1
                _lm_pp_group_style_total_row(sheet, sc, ec, row)
                li = li + 1

        if do_outline and g_end >= g_start:
            ok_ol, err_ol = _lm_pp_outline_group_rows(sheet, g_start, g_end, sc, ec)
            if ok_ol:
                outlined = outlined + 1
            elif not outline_err and err_ol:
                outline_err = unicode(err_ol)
        gi = gi - 1

    # Общие итоги внизу (SUBTOTAL игнорирует вложенные ИТОГ).
    if do_subtotals and do_grand and n_levels > 0:
        try:
            after_groups_er = int(er) + int(inserted)
        except Exception:
            after_groups_er = er
        grand_at = after_groups_er + 1
        try:
            sheet.getRows().insertByIndex(grand_at, n_levels)
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "группировка_по_столбцу",
                unicode(group_column_marker),
                "ошибка",
                "insertRows grand: %s" % err,
            )
            return
        inserted = inserted + n_levels
        li = 0
        while li < n_levels:
            row = grand_at + li
            level = levels[li]
            try:
                lab_cell = sheet.getCellByPosition(group_col, row)
                lab_cell.String = level[u"grand_label"]
            except Exception:
                pass
            for ac in level[u"cols"]:
                cell = sheet.getCellByPosition(ac, row)
                # Весь блок данных+групповые итоги; ИТОГ пропускает вложенные ИТОГ.
                if _lm_pp_group_write_subtotal_formula(
                    doc,
                    cell,
                    level[u"fn"],
                    ac,
                    sr,
                    after_groups_er,
                    sheet_name=sheet_name,
                ):
                    formula_ok = formula_ok + 1
                else:
                    formula_fail = formula_fail + 1
            _lm_pp_group_style_total_row(sheet, sc, ec, row)
            li = li + 1

    if do_outline and do_subtotals and inserted > 0:
        try:
            new_er = int(er) + int(inserted)
        except Exception:
            new_er = er
        if _lm_pp_outline_auto_from_formulas(sheet, sc, sr, ec, new_er):
            if outlined == 0:
                outlined = len(groups)

    agg_note = u"+".join([unicode(lv[u"fn"]) for lv in levels]) if levels else u"—"
    note_parts = [
        "групп=%d" % len(groups),
        "уровней=%d" % n_levels,
        "итогов=%d" % inserted,
        "формул_ok=%d" % formula_ok,
        "outline=%d" % outlined,
        "agg=%s" % agg_note,
    ]
    if removed_old:
        note_parts.append("снято_старых=%d" % removed_old)
    if formula_fail:
        note_parts.append("формул_err=%d" % formula_fail)
    if outline_err and outlined == 0:
        note_parts.append("outline_err=%s" % outline_err[:80])
    if not do_subtotals:
        note_parts.append("без_итогов (пустые agg_columns)")
    _lm_log_postprocess(
        doc,
        sheet_name,
        "группировка_по_столбцу",
        unicode(group_column_marker),
        "ok" if formula_fail == 0 else "ошибка",
        "; ".join(note_parts),
    )


# заполнить вниз

def _lm_pp_join_extra_args(extra_args):
    """Склеить rest из колонки C в одну строку."""
    if extra_args is None:
        return ""
    parts = []
    i = 0
    while i < len(extra_args):
        s = str(extra_args[i]).strip()
        if s != "":
            parts.append(s)
        i = i + 1
    return ", ".join(parts)


def _lm_pp_split_column_tokens(spec_text):
    """Разбить C на токены столбцов по запятой или точке с запятой (с учётом '…')."""
    text = str(spec_text or "").strip()
    if text == "":
        return []
    try:
        from libre_macros_param_codec import split_column_name_tokens
        parts = split_column_name_tokens(text)
        if parts:
            return parts
    except Exception:
        pass
    raw = re.split(r"[,;]", text)
    out = []
    i = 0
    while i < len(raw):
        t = str(raw[i]).strip()
        if t != "":
            out.append(t)
        i = i + 1
    return out


def _lm_pp_resolve_column_token(token, sheet, header_row_range, sc, ec):
    """
    Один токен → 0-based индекс столбца в пределах листа.
    Порядок: номер (1-based) → точный заголовок → буква A/AA…XYZ →
    подстрока в заголовке.

    Токен в '…' — только полное совпадение заголовка (без букв/подстроки).

    Точное имя заголовка важнее букв (иначе ID/SUM читаются как индекс).
    Латиница A…XYZ (1–3 буквы A–Z, is_col_letters) не участвует в поиске
    по подстроке — иначе «D»/«d» цепляет «Base64_from_D».
    """
    token_raw = str(token).strip()
    if token_raw == "":
        return None
    exact = False
    try:
        from libre_macros_param_codec import unwrap_column_name_token
        from libre_macros_param_codec import _clean_column_match_text
        token, exact = unwrap_column_name_token(token_raw)
        token = _clean_column_match_text(token)
    except Exception:
        token = token_raw
        exact = False
        try:
            from libre_macros_param_codec import _clean_column_match_text
            token = _clean_column_match_text(token)
        except Exception:
            token = str(token).strip()
    token = str(token).strip()
    if token == "":
        return None
    if not exact:
        try:
            idx = int(token) - 1
            if idx >= 0:
                return idx
        except (TypeError, ValueError):
            pass
    titles = _lm_pp_header_titles(sheet, header_row_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    key = lm_identity_key(token)
    # Нормализация для частого кейса: в параметрах имя задают как «Код_позиции»,
    # а в заголовке на листе — «Код позиции» (или с дефисами/точками).
    try:
        key2 = re.sub(r"[\W_]+", "", str(token or ""), flags=re.UNICODE).strip().casefold()
    except Exception:
        key2 = ""
    letter_token = is_col_letters(token)

    # 1) Точное совпадение заголовка (до букв и до подстроки).
    # titles уже через clean_header_name — token тоже очищен выше.
    c = h_sc
    while c <= h_ec:
        pos = c - h_sc
        title = titles[pos] if pos < len(titles) else ""
        tkey = lm_identity_key(title)
        if key != "" and tkey == key:
            return c
        if (not exact) and key2:
            try:
                tkey2 = re.sub(
                    r"[\W_]+", "", str(title or ""), flags=re.UNICODE
                ).strip().casefold()
            except Exception:
                tkey2 = ""
            if tkey2 == key2:
                return c
        c = c + 1

    if exact:
        return None

    # 2) Буквы столбца Calc A…XYZ — до нечёткого поиска по заголовку.
    if letter_token:
        return col_letters_to_index(token)

    # 3) Подстрока / вхождение в заголовок (имена вроде «Код» → «Код_позиции»).
    # Ключи-латиница A…XYZ из подстрочного поиска исключаем.
    sub_key = key if (key and not is_col_letters(key)) else ""
    sub_key2 = key2 if (key2 and not is_col_letters(key2)) else ""
    if sub_key == "" and sub_key2 == "":
        return None
    c = h_sc
    while c <= h_ec:
        pos = c - h_sc
        title = titles[pos] if pos < len(titles) else ""
        tkey = lm_identity_key(title)
        if sub_key != "" and sub_key in tkey:
            return c
        if sub_key2:
            try:
                tkey2 = re.sub(
                    r"[\W_]+", "", str(title or ""), flags=re.UNICODE
                ).strip().casefold()
            except Exception:
                tkey2 = ""
            if tkey2 and sub_key2 in tkey2:
                return c
        c = c + 1
    return None


def _lm_pp_resolve_column_list(spec_text, sheet, header_row_range, sc, ec):
    """Список уникальных 0-based индексов столбцов из C."""
    tokens = _lm_pp_split_column_tokens(spec_text)
    cols = []
    i = 0
    while i < len(tokens):
        col = _lm_pp_resolve_column_token(
            tokens[i], sheet, header_row_range, sc, ec
        )
        if col is not None and col not in cols:
            cols.append(col)
        i = i + 1
    cols.sort()
    return cols


def _lm_pp_resolve_column_matches(token, sheet, header_row_range, sc, ec):
    """
    Все 0-based индексы столбцов, совпадающих с токеном (слева направо).
    Цифра — один столбец; * в токене — шаблон text_match по заголовку;
    токен в '…' — только полное совпадение заголовка;
    иначе — через _lm_pp_resolve_column_token
    (точный заголовок → буквы A…XYZ → подстрока без латиницы A…XYZ).
    """
    token_raw = str(token).strip()
    if token_raw == "":
        return []
    exact = False
    try:
        from libre_macros_param_codec import unwrap_column_name_token
        from libre_macros_param_codec import _clean_column_match_text
        from libre_macros_param_codec import column_header_matches_token
        token, exact = unwrap_column_name_token(token_raw)
        token = _clean_column_match_text(token)
    except Exception:
        column_header_matches_token = None
        token = token_raw
        exact = False
        try:
            from libre_macros_param_codec import _clean_column_match_text
            token = _clean_column_match_text(token)
        except Exception:
            token = str(token).strip()
    token = str(token).strip()
    if token == "":
        return []
    if not exact:
        try:
            idx = int(token) - 1
            if idx >= 0 and idx <= ec:
                return [idx]
        except (TypeError, ValueError):
            pass
    titles = _lm_pp_header_titles(sheet, header_row_range)
    h_sc, _h_sr, h_ec, _h_er = _lm_pp_range_address(header_row_range)
    if (not exact) and "*" in token:
        out = []
        c = h_sc
        while c <= h_ec:
            pos = c - h_sc
            title = titles[pos] if pos < len(titles) else ""
            if _lm_pp_header_matches_marker(title, token):
                if c not in out:
                    out.append(c)
            c = c + 1
        return out
    if exact:
        out = []
        c = h_sc
        while c <= h_ec:
            pos = c - h_sc
            title = titles[pos] if pos < len(titles) else ""
            matched = False
            if column_header_matches_token is not None:
                try:
                    matched = column_header_matches_token(title, token_raw)
                except Exception:
                    matched = False
            if not matched:
                key = lm_identity_key(token)
                matched = key != "" and lm_identity_key(title) == key
            if matched and c not in out:
                out.append(c)
            c = c + 1
        return out
    one = _lm_pp_resolve_column_token(token_raw, sheet, header_row_range, sc, ec)
    if one is not None:
        return [int(one)]
    return []


def _lm_pp_resolve_column_token_unique(token, sheet, header_row_range, sc, ec):
    """Один столбец или (None, сообщение об ошибке)."""
    matches = _lm_pp_resolve_column_matches(
        token, sheet, header_row_range, sc, ec
    )
    label = str(token or "").strip()
    try:
        from libre_macros_param_codec import unwrap_column_name_token
        from libre_macros_param_codec import _clean_column_match_text
        bare, _q = unwrap_column_name_token(label)
        bare = _clean_column_match_text(bare)
        if bare != "":
            label = bare
    except Exception:
        pass
    if len(matches) == 0:
        return None, "столбец «%s» не найден" % label
    if len(matches) > 1:
        return None, "столбец «%s» не уникален (%d совпадений)" % (
            label,
            len(matches),
        )
    return matches[0], None


def _lm_pp_resolve_move_column_indices(tokens, sheet, header_row_range, sc, ec):
    """Список столбцов для переноса: порядок токенов сохраняется, дубликаты убираются."""
    out = []
    seen = set()
    i = 0
    while i < len(tokens or []):
        tok = tokens[i]
        if isinstance(tok, int):
            idx = int(tok) - 1 if int(tok) > 0 else int(tok)
        elif isinstance(tok, float):
            idx = int(tok) - 1
        else:
            s = str(tok).strip()
            if s == "":
                i = i + 1
                continue
            if s.isdigit():
                idx = int(s) - 1
            else:
                matches = _lm_pp_resolve_column_matches(
                    s, sheet, header_row_range, sc, ec
                )
                j = 0
                while j < len(matches):
                    m = matches[j]
                    if m not in seen:
                        seen.add(m)
                        out.append(m)
                    j = j + 1
                i = i + 1
                continue
        if idx >= 0 and idx <= ec and idx not in seen:
            seen.add(idx)
            out.append(idx)
        i = i + 1
    return out


def _lm_pp_build_column_reorder(all_cols, moving_cols, position, relative_col):
    """Построить новый порядок 0-based индексов столбцов (перенос)."""
    moving_set = set(moving_cols)
    rest = [c for c in all_cols if c not in moving_set]
    if position == "start":
        insert_at = 0
    elif position == "end":
        insert_at = len(rest)
    elif position in ("before", "after"):
        if relative_col is None:
            return None, "не задан опорный столбец"
        if relative_col not in rest:
            return None, "опорный столбец %d не в оставшихся" % (relative_col + 1)
        insert_at = rest.index(relative_col)
        if position == "after":
            insert_at = insert_at + 1
    else:
        return None, "неизвестная position: %s" % position
    return rest[:insert_at] + list(moving_cols) + rest[insert_at:], None


def _lm_pp_build_column_copy(all_cols, moving_cols, position, relative_col):
    """
    Порядок индексов с дубликатами (копии столбцов).

    Возвращает (new_order, insert_at, err). Длина new_order = len(all_cols) + len(moving).
    """
    rest = list(all_cols)
    if position == "start":
        insert_at = 0
    elif position == "end":
        insert_at = len(rest)
    elif position in ("before", "after"):
        if relative_col is None:
            return None, None, "не задан опорный столбец"
        if relative_col not in rest:
            return None, None, "опорный столбец %d не в диапазоне" % (relative_col + 1)
        insert_at = rest.index(relative_col)
        if position == "after":
            insert_at = insert_at + 1
    else:
        return None, None, "неизвестная position: %s" % position
    new_order = rest[:insert_at] + list(moving_cols) + rest[insert_at:]
    return new_order, insert_at, None


def _lm_pp_block_is_copy_columns(block):
    """True, если переставить_столбцы в режиме копирования."""
    if block is None:
        return False
    if _lm_pp_block_is_create_new_columns(block):
        return False
    if "copy" in block:
        try:
            return bool(block.get("copy"))
        except Exception:
            pass
    mode = str(block.get("mode") or "").strip().casefold()
    if mode in (
        "copy",
        "копировать",
        "_копировать_",
        "дублировать",
        "duplicate",
    ):
        return True
    return False


def _lm_pp_block_is_create_new_columns(block):
    """True — создать пустые столбцы по new_names (флаг «новые»)."""
    if block is None:
        return False
    if "create_new" in block:
        try:
            return bool(block.get("create_new"))
        except Exception:
            return False
    for key in ("новые", "new_empty", "blank_columns", "create_empty"):
        if key in block:
            try:
                return bool(block.get(key))
            except Exception:
                return False
    return False


def _lm_pp_parse_copy_new_names(block):
    """Список новых имён для копий (параллельно columns); пустые элементы — авто."""
    raw = None
    if block is not None:
        raw = block.get("new_names")
        if raw is None:
            raw = block.get("copy_names")
        if raw is None:
            raw = block.get("new_columns")
    if raw is None or raw == "":
        return []
    if isinstance(raw, (list, tuple)):
        return [str(x).strip() for x in raw]
    text = str(raw).strip()
    if text == "":
        return []
    out = []
    for p in re.split(r"[,;]", text):
        out.append(p.strip())
    return out


def _lm_pp_resolve_insert_at(all_cols, position, relative_col):
    """Индекс вставки блока столбцов (как в режиме copy без moving)."""
    rest = list(all_cols)
    if position == "start":
        return 0, None
    if position == "end":
        return len(rest), None
    if position in ("before", "after"):
        if relative_col is None:
            return None, "не задан опорный столбец"
        if relative_col not in rest:
            return None, "опорный столбец %d не в диапазоне" % (relative_col + 1)
        insert_at = rest.index(relative_col)
        if position == "after":
            insert_at = insert_at + 1
        return insert_at, None
    return None, "неизвестная position: %s" % position


def _lm_pp_insert_empty_columns_uno(sheet, insert_at, names, header_row):
    """
    Вставить пустые столбцы и задать заголовки.
    names — список заголовков (пустой элемент → Столбец_N).
    """
    if sheet is None or not names:
        return False, insert_at
    try:
        insert_at = int(insert_at)
        header_row = int(header_row)
    except (TypeError, ValueError):
        return False, insert_at
    n = len(names)
    if n <= 0 or insert_at < 0:
        return False, insert_at
    try:
        sheet.getColumns().insertByIndex(insert_at, n)
    except Exception:
        return False, insert_at
    i = 0
    while i < n:
        title = str(names[i] or "").strip()
        if title == "":
            title = "Столбец_%d" % (i + 1)
        try:
            sheet.getCellByPosition(insert_at + i, header_row).String = title
        except Exception:
            pass
        i = i + 1
    return True, insert_at


def _lm_pp_insert_contents_formats_props():
    """PropertyValue для .uno:InsertContents — только форматы (Flags=T)."""
    props = []
    for name, value in (
        ("Flags", "T"),
        ("FormulaCommand", 0),
        ("SkipEmptyCells", False),
        ("Transpose", False),
        ("AsLink", False),
        ("MoveMode", 4),
        ("Overwrite", True),
    ):
        prop = PropertyValue()
        prop.Name = name
        prop.Value = value
        props.append(prop)
    return tuple(props)


def _lm_pp_copy_column_width(sheet, src_col, dst_col):
    """Скопировать Width столбца (отключить OptimalWidth)."""
    if sheet is None or src_col < 0 or dst_col < 0:
        return
    try:
        cols = sheet.getColumns()
        src = cols.getByIndex(int(src_col))
        dst = cols.getByIndex(int(dst_col))
        try:
            dst.OptimalWidth = False
        except Exception:
            pass
        dst.Width = int(src.Width)
    except Exception:
        pass


def _lm_pp_copy_column_values_array(sheet, src_col, dst_col, top_row, bottom_row):
    """Скопировать значения столбца через DataArray (без формул)."""
    if sheet is None or src_col < 0 or dst_col < 0:
        return False
    if int(bottom_row) < int(top_row):
        return False
    try:
        src = sheet.getCellRangeByPosition(
            int(src_col), int(top_row), int(src_col), int(bottom_row)
        )
        data = src.getDataArray()
        dst = sheet.getCellRangeByPosition(
            int(dst_col), int(top_row), int(dst_col), int(bottom_row)
        )
        dst.setDataArray(data)
        return True
    except Exception:
        return False


def _lm_pp_clipboard_paste_column_values_then_formats( doc, sheet, src_col, dst_col, top_row, bottom_row ):
    """
    Copy столбца → InsertContents SVD (значения) → снова Copy → InsertContents T (форматы).

    Возвращает True, если хотя бы значения вставлены.
    """
    if doc is None or sheet is None:
        return False
    try:
        src_range = sheet.getCellRangeByPosition(
            int(src_col), int(top_row), int(src_col), int(bottom_row)
        )
        dst_range = sheet.getCellRangeByPosition(
            int(dst_col), int(top_row), int(dst_col), int(bottom_row)
        )
    except Exception:
        return False
    if not _lm_pp_copy_range_to_clipboard(doc, sheet, src_range):
        return False
    if not _lm_pp_select_cell_range(doc, sheet, dst_range):
        return False
    ok_v = _lm_pp_execute_dispatch(
        doc, ".uno:InsertContents", _lm_pp_insert_contents_values_props()
    )
    if not ok_v:
        return False
    # Повторный Copy — буфер после SVD иногда «пустеет»/меняется.
    if _lm_pp_copy_range_to_clipboard(doc, sheet, src_range):
        if _lm_pp_select_cell_range(doc, sheet, dst_range):
            _lm_pp_execute_dispatch(
                doc, ".uno:InsertContents", _lm_pp_insert_contents_formats_props()
            )
    return True


def _lm_pp_copy_one_column_full( doc, sheet, src_col, dst_col, top_row, bottom_row, header_row ):
    """
    Полная копия столбца: значения (+форматы через буфер или API) и ширина.

    Сначала clipboard (значения→форматы); при сбое — DataArray + clone format.
    """
    pasted = False
    if doc is not None:
        try:
            pasted = _lm_pp_clipboard_paste_column_values_then_formats(
                doc, sheet, src_col, dst_col, top_row, bottom_row
            )
        except Exception:
            pasted = False
    if not pasted:
        if not _lm_pp_copy_column_values_array(
            sheet, src_col, dst_col, top_row, bottom_row
        ):
            return False
        try:
            _lm_pp_clone_concat_column_formats(
                sheet, src_col, dst_col, header_row, header_row + 1, bottom_row
            )
        except Exception:
            pass
    _lm_pp_copy_column_width(sheet, src_col, dst_col)
    return True


def _lm_pp_copy_columns_insert_uno( doc, sheet, moving, insert_at, end_row, header_row ):
    """
    Режим copy: вставить пустые столбцы (insertByIndex), затем скопировать
    значения/форматы/ширину с исходных. Существующие столбцы не переписываются
    целиком (в отличие от DataArray-reorder) — форматы соседних не ломаются.

    moving — индексы ДО вставки; insert_at — куда вставить блок копий.
    Возвращает (ok, insert_at).
    """
    if sheet is None or not moving:
        return False, insert_at
    try:
        insert_at = int(insert_at)
        end_row = int(end_row)
        header_row = int(header_row)
    except (TypeError, ValueError):
        return False, insert_at
    n = len(moving)
    if n <= 0 or insert_at < 0:
        return False, insert_at
    top_row = min(header_row, end_row)
    bottom_row = max(header_row, end_row)
    try:
        sheet.getColumns().insertByIndex(insert_at, n)
    except Exception:
        return False, insert_at

    def _map_src(old):
        if int(old) >= insert_at:
            return int(old) + n
        return int(old)

    i = 0
    while i < n:
        src = _map_src(moving[i])
        dst = insert_at + i
        if not _lm_pp_copy_one_column_full(
            doc, sheet, src, dst, top_row, bottom_row, header_row
        ):
            return False, insert_at
        i = i + 1
    return True, insert_at


def _lm_pp_apply_copied_column_headers( sheet, header_row, insert_at, moving, new_names, src_titles ):
    """Переименовать заголовки вставленных копий (insert_at .. insert_at+n-1)."""
    if sheet is None or insert_at is None:
        return
    n = len(moving)
    if n <= 0:
        return
    try:
        header_row = int(header_row)
    except (TypeError, ValueError):
        header_row = 0
    names = list(new_names or [])
    titles = list(src_titles or [])
    i = 0
    while i < n:
        desired = names[i] if i < len(names) else ""
        desired = str(desired or "").strip()
        if desired == "":
            base = titles[i] if i < len(titles) else ""
            base = str(base or "").strip()
            if base == "":
                base = "Столбец_%d" % (moving[i] + 1)
            desired = base
        name = _lm_pp_allocate_new_header_name(sheet, header_row, desired)
        try:
            sheet.getCellByPosition(insert_at + i, header_row).String = name
        except Exception:
            pass
        i = i + 1


def _lm_pp_apply_column_reorder_uno(sheet, end_col, end_row, new_order):
    """Переставить столбцы 0..end_col на строках 0..end_row через DataArray.

    new_order той же длины, что end_col+1 (режим переноса). Для copy —
    см. _lm_pp_copy_columns_insert_uno.
    """
    if not new_order:
        return False
    ncol = end_col + 1
    ncol_new = len(new_order)
    if ncol_new != ncol:
        return False
    if new_order == list(range(ncol)):
        return True
    rng = sheet.getCellRangeByPosition(0, 0, end_col, end_row)
    data = rng.getDataArray()
    nrow = len(data)
    new_data = []
    ri = 0
    while ri < nrow:
        row = list(data[ri]) if ri < len(data) else []
        while len(row) < ncol:
            row.append(None)
        new_row = []
        ci = 0
        while ci < len(new_order):
            old = new_order[ci]
            new_row.append(row[old] if old < len(row) else None)
            ci = ci + 1
        new_data.append(tuple(new_row))
        ri = ri + 1
    rng.setDataArray(tuple(new_data))
    return True


# сортировка строк данных


def _lm_pp_parse_sheet_names_from_spec(sheet_spec):
    """Имена листов из фрагмента C (через запятую или ;)."""
    names = []
    text = str(sheet_spec or "").replace(";", ",")
    for part in text.split(","):
        s = part.strip()
        if s != "":
            names.append(s)
    return names


def _lm_pp_sheet_name_matches_target(sheet_name, target_names):
    """Совпадение имени листа с поддержкой wildcard и префикса N_."""
    nm = str(sheet_name or "").strip()
    if nm == "":
        return False
    for t in target_names or []:
        pat = str(t or "").strip()
        if pat == "":
            continue
        if text_match(nm, pat):
            return True
    # Поддержка листов результата вида "N_Имя" (частый кейс в «Копирование листов»):
    # если в фильтре указан "Имя", то должно совпасть и с "1_Имя", "2_Имя", ...
    m = re.match(r"^\d+_(.+)$", nm)
    if m is not None:
        base = str(m.group(1) or "").strip()
        if base != "":
            for t in target_names or []:
                pat = str(t or "").strip()
                if pat == "":
                    continue
                if text_match(base, pat):
                    return True
    return False


def _lm_pp_parse_sort_sheet_assignments(text):
    """«сортировка»: «Лист1,Лист2 | ключи ; Лист3: ключи»."""
    return _lm_pp_parse_sheet_param_assignments(text)


def _lm_pp_split_sheet_param_blocks(text):
    """Блоки «лист | параметры» через «;» или «,» (запятая не внутри «+число»)."""
    text = str(text or "").strip()
    if text == "":
        return []
    if ";" in text:
        # «;» между листами — только если дальше есть «|» или «:» (новый лист).
        # Иначе «;font=…;fill=…» в заголовок_плюс_высота режется как блоки.
        blocks = re.split(r"\s*;\s*(?=[^;]+(?:\||:))", text)
    elif re.search(r"(?:\||:)", text) and "," in text:
        blocks = re.split(r",\s*(?=[^,]+(?:\||:))", text)
    else:
        return [text]
    out = []
    bi = 0
    while bi < len(blocks):
        block = str(blocks[bi]).strip()
        bi = bi + 1
        if block != "":
            out.append(block)
    return out


def _lm_pp_parse_sheet_param_assignments(text):
    """
    «лист | параметры ; …» — блоки через «;» или «,», лист/параметры — «|» или «:».
    Общий разбор для сортировки, заголовок_плюс_высота и аналогов.
    """
    text = str(text or "").strip()
    if text == "":
        return []
    out = []
    blocks = _lm_pp_split_sheet_param_blocks(text)
    bi = 0
    while bi < len(blocks):
        block = str(blocks[bi]).strip()
        bi = bi + 1
        if block == "":
            continue
        sep_pos = -1
        sep_len = 0
        for sep in ("|", ":"):
            idx = block.find(sep)
            if idx >= 0 and (sep_pos < 0 or idx < sep_pos):
                sep_pos = idx
                sep_len = len(sep)
        if sep_pos < 0:
            continue
        sheet_part = block[:sep_pos].strip()
        param_part = block[sep_pos + sep_len :].strip()
        if sheet_part == "" or param_part == "":
            continue
        sheet_names = _lm_pp_parse_sheet_names_from_spec(sheet_part)
        if len(sheet_names) == 0:
            continue
        out.append((sheet_names, param_part))
    return out


def _lm_pp_sort_text_from_extra_args(extra_args):
    if extra_args is None or len(extra_args) == 0:
        return ""
    if len(extra_args) == 1:
        return str(extra_args[0]).strip()
    return _lm_pp_join_extra_args(extra_args)


def _lm_pp_sort_keys_codec_text(keys):
    chunks = []
    for item in keys or []:
        if not isinstance(item, dict):
            continue
        col = str(item.get("column") or "").strip()
        if col == "":
            continue
        if item.get("desc"):
            chunks.append("%s -> -" % col)
        else:
            chunks.append("%s -> +" % col)
    return ", ".join(chunks)


def _lm_pp_sort_extra_args_for_sheet(extra_args, sheet_name):
    """
    Ключи сортировки для текущего листа (JSON blocks или legacy-текст).
    """
    text = _lm_pp_sort_text_from_extra_args(extra_args)
    if text == "":
        return extra_args
    if _lm_pp_payload_is_json(text):
        blocks = _lm_pp_param_decode("сортировка", extra_args)
        if len(blocks) == 0:
            return extra_args
        nm = str(sheet_name or "").strip()
        has_named = False
        bi = 0
        while bi < len(blocks):
            block = blocks[bi]
            bi = bi + 1
            sheet_ref = str(block.get("sheet") or "").strip()
            if sheet_ref != "":
                has_named = True
                names = _lm_pp_parse_sheet_names_from_spec(sheet_ref)
                matched = _lm_pp_sheet_name_matches_target(nm, names)
                if not matched:
                    m = re.match(r"^(\d+)_(.+)$", nm)
                    matched = bool(
                        m and _lm_pp_sheet_name_matches_target(m.group(2), names)
                    )
                if not matched:
                    continue
            elif has_named:
                continue
            keys_text = _lm_pp_sort_keys_codec_text(block.get("keys") or [])
            return [keys_text] if keys_text else extra_args
        return None if has_named else extra_args
    assignments = _lm_pp_parse_sort_sheet_assignments(text)
    if len(assignments) == 0:
        return extra_args
    nm = str(sheet_name or "").strip()
    ai = 0
    while ai < len(assignments):
        sheet_names, sort_part = assignments[ai]
        if _lm_pp_sheet_name_matches_target(nm, sheet_names):
            return [sort_part]
        ai = ai + 1
    return None


def lm_pp_sort_step_applies_to_sheet(extra_raw, sheet_name):
    """Нужно ли выполнять «сортировка» на данном листе (до вызова колбэка)."""
    s = str(extra_raw or "").strip()
    nm = str(sheet_name or "").strip()
    _lm_pp_sort_dbg(
        "step_applies",
        "sheet=%r extra_head=%r" % (nm, s[:160]),
    )
    if s == "":
        _lm_pp_sort_dbg("step_applies", "empty C -> True")
        return True
    if _lm_pp_payload_is_json(s):
        from libre_macros_param_codec import param_decode

        blocks = param_decode("сортировка", s)
        _lm_pp_sort_dbg("step_applies", "JSON blocks=%d" % len(blocks))
        if len(blocks) == 0:
            return True
        has_named = False
        bi = 0
        while bi < len(blocks):
            block = blocks[bi]
            bi = bi + 1
            sheet_ref = str(block.get("sheet") or "").strip()
            if sheet_ref == "":
                continue
            has_named = True
            names = _lm_pp_parse_sheet_names_from_spec(sheet_ref)
            if _lm_pp_sheet_name_matches_target(nm, names):
                _lm_pp_sort_dbg("step_applies", "JSON match sheet_ref=%r" % sheet_ref)
                return True
            m = re.match(r"^(\d+)_(.+)$", nm)
            if m is not None and _lm_pp_sheet_name_matches_target(m.group(2), names):
                _lm_pp_sort_dbg("step_applies", "JSON match suffix=%r" % m.group(2))
                return True
        if has_named:
            _lm_pp_sort_dbg("step_applies", "JSON no block for sheet -> False")
            return False
        _lm_pp_sort_dbg("step_applies", "JSON no sheet fields -> True")
        return True
    assignments = _lm_pp_parse_sort_sheet_assignments(s)
    _lm_pp_sort_dbg("step_applies", "legacy assignments=%d" % len(assignments))
    if len(assignments) == 0:
        return True
    ai = 0
    while ai < len(assignments):
        sheet_names, _sort_part = assignments[ai]
        if _lm_pp_sheet_name_matches_target(nm, sheet_names):
            _lm_pp_sort_dbg("step_applies", "legacy match sheets=%r" % (sheet_names,))
            return True
        ai = ai + 1
    _lm_pp_sort_dbg("step_applies", "legacy no match -> False")
    return False


_LM_PP_HEADER_HEIGHT_MM_RE = re.compile(
    r"^(?:\+\s*)?([\d]+[.,]\d+|\d+)\s*(?:мм|mm|″|\"|''|in(?:ch(?:es)?)?|дюйм)?\s*(.*)$",
    re.IGNORECASE,
)

_LM_PP_HEADER_HEIGHT_EXT_KEYS = frozenset(
    {
        "font",
        "шрифт",
        "fill",
        "заливка",
        "bg",
        "fontcolor",
        "font_color",
        "цвет",
        "fontcolor_auto",
        "font_color_auto",
        "авто_цвет",
    }
)


def _lm_pp_split_header_height_main_and_extensions(text):
    """Основная часть (+мм align bold) и хвост font=/fill=/fontcolor= через , или ;."""
    text = str(text or "").strip()
    if text == "":
        return "", []
    pos = 0
    while pos < len(text):
        ch = text[pos]
        if ch in ",;":
            rest = text[pos + 1 :].lstrip()
            key = rest.split("=", 1)[0].strip().casefold().replace("ё", "е")
            if key in _LM_PP_HEADER_HEIGHT_EXT_KEYS and "=" in rest:
                main = text[:pos].strip()
                ext_parts = [p.strip() for p in re.split(r"[,;]", rest) if p.strip()]
                return main, ext_parts
        pos = pos + 1
    return text, []


_LM_PP_HEADER_HEIGHT_BOLD_TOKENS = frozenset(
    {"bold", "жирный", "b", "+bold"}
)


def _lm_pp_parse_header_height_bold_suffix(text):
    """Суффикс «bold» / «жирный» в хвосте параметра."""
    text = str(text or "").strip()
    if text == "":
        return False, text
    parts = text.split()
    if len(parts) >= 2 and parts[-2].casefold() == "+" and parts[-1].casefold() == "bold":
        return True, " ".join(parts[:-2]).strip()
    if parts and parts[-1].casefold() in _LM_PP_HEADER_HEIGHT_BOLD_TOKENS:
        return True, " ".join(parts[:-1]).strip()
    return False, text


def _lm_pp_parse_header_height_align_pair(rest):
    """Разбор «center-center», «center/bottom», «left top»."""
    rest = str(rest or "").strip()
    if rest == "":
        return ("center", "center")
    for sep in ("-", "_", "/"):
        if sep in rest:
            parts = rest.split(sep, 1)
            if len(parts) == 2:
                h = parts[0].strip()
                v = parts[1].strip()
                if h != "" and v != "":
                    return (h, v)
    parts = rest.split(None, 1)
    if len(parts) == 2:
        return (parts[0].strip(), parts[1].strip())
    return (rest, "center")


def _lm_pp_parse_header_height_extensions(ext_parts):
    """Разбор «;font:PT Sans/11;fill:excel_header;fontcolor:white»."""
    out = {
        "font_name": None,
        "font_size": None,
        "fill_color": None,
        "font_color": None,
        "font_color_auto": True,
    }
    if not ext_parts:
        return out
    ei = 0
    while ei < len(ext_parts):
        part = str(ext_parts[ei] or "").strip()
        ei = ei + 1
        if part == "":
            continue
        if "=" in part:
            key, _, val = part.partition("=")
        elif ":" in part:
            key, _, val = part.partition(":")
        else:
            continue
        key = key.strip().casefold().replace("ё", "е")
        val = val.strip()
        if key in ("font", "шрифт"):
            if "/" in val:
                name, size_t = val.split("/", 1)
                out["font_name"] = name.strip() or None
                try:
                    out["font_size"] = float(str(size_t).replace(",", "."))
                except (TypeError, ValueError):
                    pass
            else:
                out["font_name"] = val or None
        elif key in ("fill", "заливка", "bg"):
            out["fill_color"] = val or None
        elif key in ("fontcolor", "font_color", "цвет", "fg", "шрифт_цвет"):
            if val.casefold() in ("auto", "авто", "automatic"):
                out["font_color"] = None
                out["font_color_auto"] = True
            elif val != "":
                out["font_color"] = val
                out["font_color_auto"] = False
        elif key in ("fontcolor_auto", "font_color_auto", "авто_цвет"):
            out["font_color_auto"] = val.casefold() not in (
                "0",
                "false",
                "no",
                "нет",
                "-",
            )
    return out


def _lm_pp_header_height_cfg_defaults():
    return {
        "pad_mm": float(_LM_FMT_HEADER_ROW_PAD_MM),
        "add_height": False,
        "hori": "center",
        "vert": "center",
        "bold": False,
        "font_name": None,
        "font_size": None,
        "fill_color": None,
        "font_color": None,
        "font_color_auto": True,
    }


def _lm_pp_parse_header_height_params(text):
    """
    «+13 center-center bold;font:PT Sans/11;fill:excel_header;fontcolor:white».
    Пусто — +12,7 мм, center/center.
    """
    text = str(text or "").strip()
    cfg = _lm_pp_header_height_cfg_defaults()
    if text == "":
        return cfg
    main, ext_parts = _lm_pp_split_header_height_main_and_extensions(text)
    cfg.update(_lm_pp_parse_header_height_extensions(ext_parts))
    m = _LM_PP_HEADER_HEIGHT_MM_RE.match(main)
    if m is not None:
        try:
            align_rest = str(m.group(2) or "").strip()
            size_token = main[: len(main) - len(align_rest)].strip()
            cfg["pad_mm"] = _lm_pp_parse_mm_number(size_token)
        except Exception:
            align_rest = str(m.group(2) or "").strip()
        bold, align_text = _lm_pp_parse_header_height_bold_suffix(align_rest)
        hori, vert = _lm_pp_parse_header_height_align_pair(align_text)
        cfg["bold"] = bold
        cfg["hori"] = hori
        cfg["vert"] = vert
    else:
        bold, align_text = _lm_pp_parse_header_height_bold_suffix(main)
        hori, vert = _lm_pp_parse_header_height_align_pair(align_text)
        cfg["bold"] = bold
        cfg["hori"] = hori
        cfg["vert"] = vert
    return cfg


def _lm_pp_sheet_ref_list_matches_sheet(doc, sheet, target_names):
    """Лист совпадает с именем, 1-based индексом или маркером «Все»."""
    if sheet is None:
        return False
    nm = ""
    try:
        nm = str(sheet.Name).strip()
    except Exception:
        pass
    ti = 0
    while ti < len(target_names):
        ts = str(target_names[ti]).strip()
        ti = ti + 1
        if ts.casefold() in _LM_PP_SHEET_ALL_MARKERS:
            return True
        if _lm_pp_sheet_name_matches_target(nm, (ts,)):
            return True
        if ts.isdigit() and doc is not None:
            resolved = _lm_vlookup_resolve_sheet(doc, ts)
            if resolved is not None:
                try:
                    if str(resolved.Name).strip() == nm:
                        return True
                except Exception:
                    pass
    return False


def _lm_pp_header_height_block_to_cfg(block):
    """JSON/блок кодека → cfg для _lm_pp_apply_header_row_height_pad."""
    if not isinstance(block, dict):
        return None
    try:
        from libre_macros_param_codec import _normalize_header_plus_height_block

        block = _normalize_header_plus_height_block(block)
    except Exception:
        pass
    font_name = str(block.get("font_name") or "").strip()
    font_size = block.get("font_size")
    fill_color = str(block.get("fill_color") or "").strip()
    font_color = str(block.get("font_color") or "").strip()
    return {
        "pad_mm": block.get("height_mm", _LM_FMT_HEADER_ROW_PAD_MM),
        "add_height": bool(block.get("add_height")),
        "hori": block.get("h_align", "center"),
        "vert": block.get("v_align", "center"),
        "bold": bool(block.get("bold")),
        "font_name": font_name or None,
        "font_size": font_size,
        "fill_color": fill_color or None,
        "font_color": font_color or None,
        "font_color_auto": bool(block.get("font_color_auto", True)),
    }


def _lm_pp_header_height_pick_block_for_sheet(blocks, doc, sheet):
    return _lm_pp_pick_sheet_block_for_sheet(blocks, doc, sheet)


def _lm_pp_header_height_config_from_extra(doc, sheet, extra_raw):
    """Разбор C (текст или JSON) в cfg для текущего листа."""
    raw = str(extra_raw or "").strip()
    sheet_name = ""
    try:
        sheet_name = str(sheet.Name) if sheet is not None else ""
    except Exception:
        pass
    _lm_pp_header_height_dbg(
        "config_from_extra",
        "sheet=%s raw=%r" % (sheet_name, raw[:240]),
    )
    if raw == "":
        cfg = _lm_pp_parse_header_height_params("")
        _lm_pp_header_height_dbg("config_from_extra", "empty -> default %s" % cfg)
        return cfg
    if _lm_pp_payload_is_json(raw):
        from libre_macros_param_codec import param_decode

        blocks = param_decode("заголовок_плюс_высота", raw)
        block = _lm_pp_header_height_pick_block_for_sheet(blocks, doc, sheet)
        if block is None:
            _lm_pp_header_height_dbg(
                "config_from_extra", "JSON no block for sheet=%s" % sheet_name
            )
            return None
        cfg = _lm_pp_header_height_block_to_cfg(block)
        _lm_pp_header_height_dbg("config_from_extra", "JSON cfg=%s" % cfg)
        return cfg
    assignments = _lm_pp_parse_sheet_param_assignments(raw)
    if len(assignments) == 0:
        cfg = _lm_pp_parse_header_height_params(raw)
        _lm_pp_header_height_dbg("config_from_extra", "text cfg=%s" % cfg)
        return cfg
    ai = 0
    while ai < len(assignments):
        sheet_names, param_text = assignments[ai]
        if _lm_pp_sheet_ref_list_matches_sheet(doc, sheet, sheet_names):
            cfg = _lm_pp_parse_header_height_params(param_text)
            _lm_pp_header_height_dbg(
                "config_from_extra",
                "assignment sheets=%s cfg=%s" % (sheet_names, cfg),
            )
            return cfg
        ai = ai + 1
    _lm_pp_header_height_dbg(
        "config_from_extra", "no assignment for sheet=%s" % sheet_name
    )
    return None


def _lm_pp_resolve_header_height_config_from_specs(doc, sheet, *callback_specs):
    """Последний подходящий блок «заголовок_плюс_высота» из цепочки постобработки."""
    cfg = None
    si = 0
    while si < len(callback_specs):
        spec = callback_specs[si]
        if spec is not None:
            for fn in _lm_postprocess_callbacks(spec):
                extra = getattr(fn, "_lm_pp_header_height_extra", None)
                if extra is None or str(extra).strip() == "":
                    continue
                c = _lm_pp_header_height_config_from_extra(doc, sheet, extra)
                if c is not None:
                    cfg = c
        si = si + 1
    return cfg


def _lm_pp_maybe_reapply_header_height_pad( doc, sheet, header_row, end_col, *callback_specs ):
    """Повторно применить заголовок_плюс_высота после финализации (сбрасывает дефолты)."""
    if sheet is None or end_col < 0:
        return
    _lm_pp_header_height_dbg(
        "maybe_reapply",
        "sheet=%s header_row=%s end_col=%s"
        % (getattr(sheet, "Name", sheet), header_row, end_col),
    )
    cfg = _lm_pp_resolve_header_height_config_from_specs(doc, sheet, *callback_specs)
    if cfg is None:
        _lm_pp_header_height_dbg("maybe_reapply", "SKIP cfg=None")
        return
    try:
        header_range = sheet.getCellRangeByPosition(0, header_row, end_col, header_row)
    except Exception:
        header_range = None
    if header_range is None:
        _lm_pp_header_height_dbg("maybe_reapply", "SKIP header_range=None")
        return
    _lm_pp_header_height_dbg("maybe_reapply", "cfg=%s" % cfg)
    _lm_pp_apply_header_row_height_pad(sheet, header_range, cfg)


def _lm_pp_header_height_text_from_extra_args(extra_args):
    return _lm_pp_sort_text_from_extra_args(extra_args)


def _lm_pp_header_height_config_for_sheet(extra_args, doc, sheet):
    """
    Параметры заголовок_плюс_высота для листа.
    None — лист не указан в C (пропуск). Пустая C — значения по умолчанию.
    """
    text = _lm_pp_header_height_text_from_extra_args(extra_args)
    return _lm_pp_header_height_config_from_extra(doc, sheet, text)


def _lm_pp_apply_header_row_cell_style(cell, cfg, fill_rgb=None, font_color_rgb=None):
    if cell is None or cfg is None:
        return
    font_name = cfg.get("font_name")
    if font_name:
        try:
            cell.CharFontName = str(font_name)
        except Exception:
            pass
    font_size = cfg.get("font_size")
    if font_size is not None:
        try:
            cell.CharHeight = float(font_size)
        except Exception:
            pass
    if cfg.get("bold"):
        try:
            cell.CharWeight = _LM_FMT_CHAR_WEIGHT_BOLD
        except Exception:
            pass
    else:
        try:
            from com.sun.star.awt.FontWeight import NORMAL

            cell.CharWeight = NORMAL
        except Exception:
            pass
    if fill_rgb is not None:
        try:
            cell.IsCellBackgroundTransparent = False
            cell.CellBackColor = int(fill_rgb)
        except Exception:
            pass
    font_color_auto = bool(cfg.get("font_color_auto", True))
    if font_color_rgb is not None:
        font_color_auto = False
    if font_color_auto:
        try:
            cell.CharAutoColor = True
        except Exception:
            pass
        if fill_rgb is not None:
            _lm_set_char_color_cell(
                cell, _lm_pp_colorize_auto_text_color(fill_rgb)
            )
    elif font_color_rgb is not None:
        _lm_set_char_color_cell(cell, font_color_rgb)


def _lm_pp_apply_header_row_height_pad(sheet, header_row_range, cfg):
    """Высота заголовка, выравнивание, шрифт и цвета по cfg."""
    if sheet is None or header_row_range is None or cfg is None:
        _lm_pp_header_height_dbg(
            "apply",
            "SKIP sheet=%s range=%s cfg=%s"
            % (sheet, header_row_range, cfg),
        )
        return
    pad_mm = cfg.get("pad_mm", _LM_FMT_HEADER_ROW_PAD_MM)
    add_height = bool(cfg.get("add_height"))
    hori = cfg.get("hori", "center")
    vert = cfg.get("vert", "center")
    _sc, sr, ec, _er = _lm_pp_range_address(header_row_range)
    _lm_pp_header_height_dbg(
        "apply",
        "sheet=%s row=%s cols=0..%s pad_mm=%s add_height=%s hori=%s vert=%s cfg=%s"
        % (getattr(sheet, "Name", sheet), sr, ec, pad_mm, add_height, hori, vert, cfg),
    )
    _lm_pp_set_or_pad_row_height_mm(sheet, sr, pad_mm, add_height=add_height)
    _lm_apply_header_row_alignment(sheet, ec, sr, hori, vert)
    fill_rgb = None
    fill_token = cfg.get("fill_color")
    if fill_token:
        fill_rgb = _lm_pp_zebra_resolve_color_token(fill_token)
    font_color_rgb = None
    font_token = cfg.get("font_color")
    if font_token:
        font_color_rgb = _lm_pp_zebra_resolve_color_token(font_token)
    c = 0
    while c <= ec:
        try:
            cell = sheet.getCellByPosition(c, sr)
            _lm_pp_apply_header_row_cell_style(
                cell, cfg, fill_rgb=fill_rgb, font_color_rgb=font_color_rgb
            )
        except Exception:
            pass
        c = c + 1
    _lm_pp_header_height_debug_readback(sheet, sr, 0)
    _lm_pp_header_height_dbg(
        "apply",
        "done sheet=%s row=%s fill_rgb=%s font_color_rgb=%s"
        % (
            getattr(sheet, "Name", sheet),
            sr,
            fill_rgb,
            font_color_rgb,
        ),
    )


def lm_pp_header_height_step_applies_to_sheet(extra_raw, sheet_name, doc=None, sheet=None):
    """Нужно ли выполнять «заголовок_плюс_высота» на данном листе."""
    s = str(extra_raw or "").strip()
    if s == "":
        applies = True
    else:
        assignments = _lm_pp_parse_sheet_param_assignments(s)
        if len(assignments) == 0:
            applies = True
        else:
            if sheet is None and doc is not None:
                sheet = _lm_vlookup_resolve_sheet(doc, sheet_name)
            if sheet is not None:
                applies = (
                    _lm_pp_header_height_config_for_sheet((s,), doc, sheet)
                    is not None
                )
            else:
                nm = str(sheet_name or "").strip()
                applies = False
                ai = 0
                while ai < len(assignments):
                    sheet_names, _param = assignments[ai]
                    if _lm_pp_sheet_name_matches_target(nm, sheet_names):
                        applies = True
                        break
                    sj = 0
                    while sj < len(sheet_names):
                        if (
                            str(sheet_names[sj]).strip().casefold()
                            in _LM_PP_SHEET_ALL_MARKERS
                        ):
                            applies = True
                            break
                        sj = sj + 1
                    if applies:
                        break
                    ai = ai + 1
    _lm_pp_header_height_dbg(
        "step_applies",
        "sheet=%s extra=%r -> %s"
        % (sheet_name, s[:120], applies),
    )
    return applies


_LM_PP_SORT_DIRECTION_DESC = frozenset({
    "-",
    "desc",
    "d",
    "descending",
    "убыв",
    "убывание",
    "убыванию",
})


def _lm_pp_parse_sort_direction(token):
    """True — по возрастанию (+ / asc / возр); False — по убыванию."""
    t = str(token or "").strip().casefold()
    if t in _LM_PP_SORT_DIRECTION_DESC:
        return False
    return True


def _lm_pp_parse_sort_block(block):
    """
    Один блок C: «колонка -> +», «Сумма/-», «1=desc», «Дата - убыв».
    Разделитель колонка/направление: ->, /, = или пробелы вокруг «-».
    """
    block = str(block or "").strip()
    if block == "":
        return None
    col_part = block
    dir_part = ""
    for sep in ("->", "/", "="):
        idx = block.find(sep)
        if idx >= 0:
            col_part = block[:idx]
            dir_part = block[idx + len(sep) :]
            break
    else:
        m = re.match(r"^(.+?)\s*-\s*(.+)$", block)
        if m is not None:
            right = str(m.group(2)).strip().casefold()
            if (
                right in _LM_PP_SORT_DIRECTION_DESC
                or right in ("+", "-")
                or right in ("asc", "desc", "возр", "убыв")
            ):
                col_part = m.group(1)
                dir_part = m.group(2)
        elif len(block) >= 2 and block[-1] in "+-":
            col_part = block[:-1]
            dir_part = block[-1]
    col_token = str(col_part).strip()
    if col_token == "":
        return None
    return (col_token, _lm_pp_parse_sort_direction(dir_part))


def _lm_pp_parse_sort_specs(extra_args):
    """Список (токен_столбца, ascending) из колонки C."""
    if extra_args is None or len(extra_args) == 0:
        return []
    if len(extra_args) == 1:
        text = str(extra_args[0]).strip()
    else:
        text = _lm_pp_join_extra_args(extra_args)
    blocks = _lm_pp_split_column_tokens(text)
    specs = []
    i = 0
    while i < len(blocks):
        parsed = _lm_pp_parse_sort_block(blocks[i])
        if parsed is not None:
            specs.append(parsed)
        i = i + 1
    return specs


def _lm_pp_sort_normalize_value(val):
    """Ключ сравнения: (ранг_типа, значение) — пустые, текст, числа."""
    if val is None:
        return (0, "")
    if isinstance(val, bool):
        return (2, 1 if val else 0)
    if isinstance(val, (int, float)):
        try:
            return (2, float(val))
        except (TypeError, ValueError):
            pass
    s = str(val).strip()
    if s == "":
        return (0, "")
    norm = s
    for sep in (" ", "\t", "\xa0", "\u202f", "\u2009"):
        norm = norm.replace(sep, "")
    norm = norm.replace(",", ".")
    try:
        return (2, float(norm))
    except ValueError:
        return (1, s.casefold())


def _lm_pp_sort_row_list(row_list, resolved_specs, start_col):
    """
    Стабильная многоуровневая сортировка строк (tuple из getDataArray).
    resolved_specs: [(col_index_0based, ascending), …] — порядок важен.
    """
    if row_list is None or len(row_list) == 0:
        return row_list
    if resolved_specs is None or len(resolved_specs) == 0:
        return row_list
    sorted_rows = list(row_list)
    ki = len(resolved_specs) - 1
    while ki >= 0:
        col_idx, ascending = resolved_specs[ki]
        rel = int(col_idx) - int(start_col)

        def _row_key(row, rel_col=rel):
            if rel_col < 0 or rel_col >= len(row):
                return _lm_pp_sort_normalize_value(None)
            return _lm_pp_sort_normalize_value(row[rel_col])

        sorted_rows.sort(key=_row_key, reverse=not ascending)
        ki = ki - 1
    return sorted_rows


def _lm_pp_make_table_sort_field(field, ascending=True):
    """TableSortField через uno.createUnoStruct (как в test_sort.py)."""
    sort_field = uno.createUnoStruct("com.sun.star.table.TableSortField")
    sort_field.Field = int(field)
    sort_field.IsAscending = bool(ascending)
    return sort_field


def _lm_pp_sort_fields_uno_any(sort_fields):
    """Значение SortFields для дескриптора сортировки."""
    try:
        return uno.Any("[]com.sun.star.table.TableSortField", tuple(sort_fields))
    except Exception:
        return tuple(sort_fields)


def _lm_pp_patch_sort_descriptor(desc, sort_fields, contains_header):
    """Подставить SortFields и ContainsHeader в дескриптор createSortDescriptor()."""
    patched = []
    i = 0
    while i < len(desc):
        prop = desc[i]
        if prop.Name == "SortFields":
            prop.Value = _lm_pp_sort_fields_uno_any(sort_fields)
        elif prop.Name == "ContainsHeader":
            prop.Value = bool(contains_header)
        patched.append(prop)
        i += 1
    names = [p.Name for p in patched]
    if "SortFields" not in names:
        prop_fields = PropertyValue()
        prop_fields.Name = "SortFields"
        prop_fields.Value = _lm_pp_sort_fields_uno_any(sort_fields)
        patched.append(prop_fields)
    if "ContainsHeader" not in names:
        prop_header = PropertyValue()
        prop_header.Name = "ContainsHeader"
        prop_header.Value = bool(contains_header)
        patched.append(prop_header)
    return tuple(patched)


def _lm_pp_range_sort_native_via_descriptor( block, resolved, sc, ec, contains_header=False, sheet_name="", fn_label="сортировка" ):
    """
    Нативная сортировка: range.sort(createSortDescriptor()) — без queryInterface.

    Подход проверен в macro-lib/test_sort.py.
    """
    if block is None:
        return False, "нет диапазона"
    if not hasattr(block, "sort"):
        return False, "нет метода sort()"
    if not hasattr(block, "createSortDescriptor"):
        return False, "нет createSortDescriptor()"

    fields = []
    i = 0
    while i < len(resolved):
        col_idx, ascending = resolved[i]
        rel = int(col_idx) - int(sc)
        if rel < 0 or rel > int(ec) - int(sc):
            return False, "столбец %d вне диапазона" % (int(col_idx) + 1)
        fields.append(_lm_pp_make_table_sort_field(rel, ascending))
        i += 1

    try:
        sort_desc = _lm_pp_patch_sort_descriptor(
            block.createSortDescriptor(), fields, contains_header
        )
        block.sort(sort_desc)
    except Exception as err:
        return False, "sort: %s" % err

    _lm_pp_sort_dbg(
        "native_descriptor",
        "sheet=%r fn=%s keys=%d ContainsHeader=%s"
        % (sheet_name, fn_label, len(fields), contains_header),
    )
    return True, ""


def _lm_pp_range_sort_via_data_array( sheet, sc, sr, ec, er, resolved, start_col, sheet_name="", fn_label="сортировка" ):
    """Legacy: сортировка через getDataArray / setDataArray (формулы → значения)."""
    try:
        block = sheet.getCellRangeByPosition(sc, sr, ec, er)
        rows = block.getDataArray()
    except Exception as err:
        return False, "getDataArray: %s" % err

    if rows is None or len(rows) == 0:
        return False, "пусто"

    row_list = []
    ri = 0
    while ri < len(rows):
        row_list.append(tuple(rows[ri]))
        ri = ri + 1
    sorted_rows = _lm_pp_sort_row_list(row_list, resolved, start_col)
    try:
        block.setDataArray(tuple(sorted_rows))
    except Exception as err:
        return False, "setDataArray: %s" % err
    _lm_pp_sort_dbg(
        "data_array",
        "sheet=%r rows=%d keys=%d" % (sheet_name, len(sorted_rows), len(resolved)),
    )
    return True, ""


def _lm_pp_range_sort_via_native( doc, sheet, sc, sr, ec, er, resolved, sheet_name="", fn_label="сортировка" ):
    """
    Нативная сортировка Calc (формулы сохраняются):

    1) range.sort(createSortDescriptor()) — как test_sort.py;
    2) запасной путь: .uno:DataSort через dispatch (Macro Recorder).
    """
    try:
        block = sheet.getCellRangeByPosition(sc, sr, ec, er)
    except Exception as err:
        return False, "range: %s" % err

    ok, err = _lm_pp_range_sort_native_via_descriptor(
        block,
        resolved,
        sc,
        ec,
        contains_header=False,
        sheet_name=sheet_name,
        fn_label=fn_label,
    )
    if ok:
        return True, ""

    _lm_pp_sort_dbg("native_descriptor_fail", err or "—")

    try:
        if doc is None:
            return False, err or "sort FAIL"
        if not _lm_pp_select_cell_range(doc, sheet, block):
            return False, "%s (select FAIL)" % (err or "sort FAIL")

        props = []

        def _pv(name, value):
            p = PropertyValue()
            p.Name = name
            p.Value = value
            return p

        props.append(_pv("ByRows", True))
        props.append(_pv("HasHeader", False))
        props.append(_pv("CaseSensitive", False))
        props.append(_pv("NaturalSort", False))
        props.append(_pv("IncludeAttribs", True))
        props.append(_pv("UserDefIndex", 0))
        ki = 0
        while ki < len(resolved) and ki < 3:
            col_idx, asc = resolved[ki]
            rel_1based = int(col_idx) - int(sc) + 1
            props.append(_pv("Col%d" % (ki + 1), int(rel_1based)))
            props.append(_pv("Ascending%d" % (ki + 1), bool(asc)))
            ki += 1
        props.append(_pv("IncludeComments", False))
        props.append(_pv("IncludeImages", True))
        if not _lm_pp_execute_dispatch(doc, ".uno:DataSort", tuple(props)):
            return False, "%s (DataSort FAIL)" % (err or "sort FAIL")
        _lm_pp_sort_dbg(
            "native_dispatch",
            "sheet=%r rows=%d..%d keys=%d"
            % (sheet_name, int(sr) + 1, int(er) + 1, len(resolved)),
        )
        return True, ""
    except Exception as exc:
        return False, "%s (DataSort EXC: %s)" % (err or "sort FAIL", exc)


def _lm_pp_range_sort_dispatch( doc, sheet, sc, sr, ec, er, resolved, start_col, sheet_name="", fn_label="сортировка" ):
    """Единая точка: native (range.sort) или data_array — по _LM_PP_SORT_MODE."""
    use_native = _lm_pp_sort_use_native()
    _lm_pp_sort_dbg(
        "mode",
        "%s sheet=%r fn=%s"
        % ("native" if use_native else "data_array", sheet_name, fn_label),
    )
    if use_native:
        ok, err = _lm_pp_range_sort_via_native(
            doc,
            sheet,
            sc,
            sr,
            ec,
            er,
            resolved,
            sheet_name=sheet_name,
            fn_label=fn_label,
        )
        if ok:
            return True, ""
        # Если native (range.sort / DataSort) не сработал — откат на data_array.
        _lm_pp_sort_dbg(
            "native_fail",
            "fallback to data_array: %s" % (err or "—"),
        )
        return _lm_pp_range_sort_via_data_array(
            sheet,
            sc,
            sr,
            ec,
            er,
            resolved,
            start_col,
            sheet_name=sheet_name,
            fn_label=fn_label,
        )
    return _lm_pp_range_sort_via_data_array(
        sheet,
        sc,
        sr,
        ec,
        er,
        resolved,
        start_col,
        sheet_name=sheet_name,
        fn_label=fn_label,
    )


def _lm_pp_range_sort_data_core( doc, sheet, data_range, header_row_range, extra_args, fn_label="сортировка" ):
    """Общая логика сортировки диапазона данных (range и финал)."""
    sheet_name = ""
    try:
        sheet_name = str(sheet.Name) if sheet is not None else ""
    except Exception:
        pass
    specs = _lm_pp_parse_sort_specs(extra_args)
    if len(specs) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            fn_label,
            _lm_pp_join_extra_args(extra_args),
            "пропуск",
            "пустая или неверная колонка C",
        )
        return False, "пустая C"

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    if er < sr:
        _lm_log_postprocess(
            doc,
            sheet_name,
            fn_label,
            _lm_pp_join_extra_args(extra_args),
            "пропуск",
            "нет строк данных",
        )
        return False, "нет строк"

    resolved = []
    i = 0
    while i < len(specs):
        col_token, ascending = specs[i]
        col_idx = _lm_pp_resolve_column_token(
            col_token, sheet, header_row_range, sc, ec
        )
        if col_idx is None:
            _lm_log_postprocess(
                doc,
                sheet_name,
                fn_label,
                col_token,
                "ошибка",
                "столбец не найден: %s" % col_token,
            )
            return False, "столбец не найден"
        resolved.append((col_idx, ascending))
        i = i + 1

    use_native = _lm_pp_sort_use_native()
    ok, err = _lm_pp_range_sort_dispatch(
        doc,
        sheet,
        sc,
        sr,
        ec,
        er,
        resolved,
        sc,
        sheet_name=sheet_name,
        fn_label=fn_label,
    )
    if not ok:
        _lm_log_postprocess(
            doc,
            sheet_name,
            fn_label,
            "—",
            "ошибка",
            err or "сортировка не выполнена",
        )
        return False, err or "ошибка"

    key_desc = []
    j = 0
    while j < len(specs):
        tok, asc = specs[j]
        key_desc.append("%s %s" % (tok, "+" if asc else "-"))
        j = j + 1
    mode_tag = "native" if use_native else "data_array"
    _lm_log_postprocess(
        doc,
        sheet_name,
        fn_label,
        _lm_pp_join_extra_args(extra_args),
        "ok",
        "строк %d, ключи: %s (%s)"
        % (int(er) - int(sr) + 1, ", ".join(key_desc), mode_tag),
    )
    return True, ""


def lm_pp_range_sort_data(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Отсортировать строки данных диапазона (заголовок не трогается).

    JSON колонки C: [{"v":1,"fn":"сортировка","keys":[{"column":"ФИО","desc":false},{"column":"Сумма","desc":true}]}]
    или с "sheet":"Лист". См. docs/13_JSON_PARAMS.md.

    Карта: «сортировка».

    Режим (_LM_PP_SORT_MODE):
        native (по умолчанию) — range.sort + createSortDescriptor, формулы сохраняются;
        data_array — getDataArray/setDataArray (формулы → значения).

    Что меняет:
        Переставляет строки данных (native или data_array).
    """
    sheet_name = ""
    try:
        sheet_name = str(sheet.Name) if sheet is not None else ""
    except Exception:
        pass
    _lm_pp_sort_dbg(
        "lm_pp_range_sort_data",
        "start sheet=%r extra=%r"
        % (sheet_name, _lm_pp_sort_text_from_extra_args(extra_args)[:200]),
    )
    if _lm_pp_param_skip_sheet("сортировка", extra_args, doc, sheet):
        _lm_pp_sort_dbg("param_skip", "skip sheet=%r" % sheet_name)
        return
    resolved = _lm_pp_sort_extra_args_for_sheet(extra_args, sheet_name)
    _lm_pp_sort_dbg(
        "resolve_for_sheet",
        "resolved=%r" % (_lm_pp_sort_text_from_extra_args(resolved) if resolved else None),
    )
    if resolved is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "сортировка",
            _lm_pp_sort_text_from_extra_args(extra_args),
            "пропуск",
            "лист не указан в C",
        )
        return
    _lm_pp_range_sort_data_core(
        doc, sheet, data_range, header_row_range, resolved, fn_label="сортировка"
    )


# =============================================================================
# «Раскрасить» — группировка по ключу столбцов, опциональная сортировка, заливка
# =============================================================================


def _lm_pp_is_colorize_postprocess_name(fn_name):
    try:
        from libre_macros_param_codec import normalize_fn_key

        return normalize_fn_key(fn_name) == "раскрасить_блоки"
    except ImportError:
        key = str(fn_name or "").strip().lower()
        return key in ("раскрасить", "раскрасить_блоки")


def _lm_pp_colorize_extra_args(extra_raw):
    """Колонка C для «раскрасить_блоки» — одна строка; JSON → канонический текст."""
    return _lm_pp_codec_text_args("раскрасить_блоки", extra_raw)


def _lm_pp_split_colorize_key_and_style(param_part):
    """
    «столбцы ключа -> оформление» — разделитель ->, = или пробелы вокруг «-».
    """
    text = str(param_part or "").strip()
    if text == "":
        return "", ""
    for sep in ("->", "="):
        idx = text.find(sep)
        if idx >= 0:
            return text[:idx].strip(), text[idx + len(sep) :].strip()
    m = re.search(r"\s+-\s+", text)
    if m is not None:
        return text[: m.start()].strip(), text[m.end() :].strip()
    return text, ""


_LM_PP_COLORIZE_OUTLINE_FLAG_RE = re.compile(
    r"!\s*(?:граница|outline)(?:\s*:\s*(?P<color>[^\s;]+))?\s*$",
    re.IGNORECASE,
)


def _lm_pp_colorize_block_for_sheet_parse(block):
    """Убрать «!граница:цвет» перед поиском разделителя листа «:»."""
    text = str(block or "").strip()
    m = _LM_PP_COLORIZE_OUTLINE_FLAG_RE.search(text)
    if m is not None:
        return text[: m.start()].strip()
    return text


def _lm_pp_parse_colorize_assignments(text):
    """
    Блоки «лист | ключ -> цвета ; …» или «ключ -> цвета» (все листы).
    Несколько блоков — только через «;» (запятая внутри ключей/цветов не режет блок).
    """
    text = str(text or "").strip()
    if text == "":
        return []
    if ";" in text:
        blocks = [b.strip() for b in re.split(r"\s*;\s*", text) if b.strip()]
    else:
        blocks = [text]
    out = []
    bi = 0
    while bi < len(blocks):
        block = str(blocks[bi]).strip()
        bi = bi + 1
        if block == "":
            continue
        block_parse = _lm_pp_colorize_block_for_sheet_parse(block)
        sep_pos = -1
        sep_len = 0
        for sep in ("|", ":"):
            idx = block_parse.find(sep)
            if idx >= 0 and (sep_pos < 0 or idx < sep_pos):
                sep_pos = idx
                sep_len = len(sep)
        if sep_pos < 0:
            key_part, style_part = _lm_pp_split_colorize_key_and_style(block)
            if key_part != "":
                out.append((None, key_part, style_part))
            continue
        sheet_part = block[:sep_pos].strip()
        rest = block[sep_pos + sep_len :].strip()
        key_part, style_part = _lm_pp_split_colorize_key_and_style(rest)
        if key_part == "":
            continue
        sheet_names = _lm_pp_parse_sheet_names_from_spec(sheet_part)
        if len(sheet_names) == 0:
            continue
        out.append((sheet_names, key_part, style_part))
    return out


def _lm_pp_colorize_split_style_and_sort(style_part):
    """
    «оформление» и опционально «/направление» в конце.
    Без суффикса — сортировка не выполняется; /+, /asc, /возр — по возрастанию;
    /-, /desc, /убыв — по убыванию (как у «сортировка»).
    """
    text = str(style_part or "").strip()
    if text == "":
        return "", None
    m = re.search(
        r"\s*/(?P<dir>\+|−|-|asc|desc|d|descending|возр|возрастание|возрастанию|убыв|убывание|убыванию)\s*$",
        text,
        re.I,
    )
    if m is None:
        return text, None
    colors_part = text[: m.start()].strip()
    sort_asc = _lm_pp_parse_sort_direction(m.group("dir"))
    return colors_part, sort_asc


def _lm_pp_colorize_parse_colors(style_part):
    """Оформление: один цвет или список через «%» / «,»."""
    style_part = str(style_part or "").strip()
    if style_part == "":
        return []
    if "%" in style_part:
        tokens = [t.strip() for t in style_part.split("%") if t.strip()]
    elif "," in style_part:
        tokens = [t.strip() for t in style_part.split(",") if t.strip()]
    else:
        tokens = [style_part]
    out = []
    i = 0
    while i < len(tokens):
        c = _lm_pp_zebra_resolve_color_token(tokens[i])
        if c is not None:
            out.append(c)
        i = i + 1
    return out


def _lm_pp_colorize_strip_outline_flags(style_part):
    text = str(style_part or "").strip()
    m = _LM_PP_COLORIZE_OUTLINE_FLAG_RE.search(text)
    if m is None:
        return text, False, None
    before = text[: m.start()].strip()
    color = m.group("color")
    if color is not None:
        color = str(color).strip() or None
    return before, True, color


def _lm_pp_colorize_make_config(key_part, style_part):
    style_part, outline_blocks, outline_color = _lm_pp_colorize_strip_outline_flags(
        style_part
    )
    colors_text, sort_asc = _lm_pp_colorize_split_style_and_sort(style_part)
    return {
        "key_tokens": _lm_pp_split_column_tokens(key_part),
        "colors": _lm_pp_colorize_parse_colors(colors_text),
        "sort": sort_asc,
        "outline_blocks": outline_blocks,
        "outline_color": outline_color,
    }


def _lm_pp_colorize_resolve_outline_color(color_token):
    if color_token is None or str(color_token).strip() == "":
        return 0x1F4E79
    resolved = _lm_pp_zebra_resolve_color_token(str(color_token).strip())
    if resolved is not None:
        return int(resolved)
    return 0x1F4E79


def _lm_pp_colorize_apply_block_outline(sheet, sc, ec, row_first, row_last, color, width):
    """Толстая внешняя рамка блока (§6.1 раскрасить_блоки, outline_blocks)."""
    if sheet is None or row_last < row_first:
        return
    c = int(sc)
    while c <= int(ec):
        try:
            _lm_pp_set_cell_side_border(
                sheet.getCellByPosition(c, int(row_first)), "top", color, width
            )
            _lm_pp_set_cell_side_border(
                sheet.getCellByPosition(c, int(row_last)), "bottom", color, width
            )
        except Exception:
            pass
        c = c + 1
    r = int(row_first)
    while r <= int(row_last):
        try:
            _lm_pp_set_cell_side_border(
                sheet.getCellByPosition(int(sc), r), "left", color, width
            )
            _lm_pp_set_cell_side_border(
                sheet.getCellByPosition(int(ec), r), "right", color, width
            )
        except Exception:
            pass
        r = r + 1


def _lm_pp_colorize_resolve_color_list(colors):
    """Имена/hex → int RGB (как _lm_pp_colorize_parse_colors для JSON)."""
    out = []
    i = 0
    while i < len(colors):
        c = _lm_pp_zebra_resolve_color_token(colors[i])
        if c is not None:
            out.append(int(c))
        i = i + 1
    return out


def _lm_pp_colorize_config_from_block(block):
    from libre_macros_param_codec import _normalize_colorize_block

    block = _normalize_colorize_block(block)
    sort_val = block.get("sort")
    sort_asc = None
    if sort_val == "asc":
        sort_asc = True
    elif sort_val == "desc":
        sort_asc = False
    return {
        "key_tokens": list(block.get("key_columns") or []),
        "colors": _lm_pp_colorize_resolve_color_list(block.get("colors") or []),
        "sort": sort_asc,
        "outline_blocks": bool(block.get("outline_blocks")),
        "outline_color": block.get("outline_color"),
    }


def _lm_pp_colorize_config_for_sheet(extra_args, sheet_name):
    """
    Параметры «раскрасить» для листа.
    None — лист не указан в C (пропуск); пустая C — пропуск.
    """
    text = _lm_pp_sort_text_from_extra_args(extra_args)
    _lm_dbg("colorize_config_for_sheet sheet=%r text=%r" % (sheet_name, text))
    if text == "":
        return None
    if _lm_pp_payload_is_json(text):
        blocks = _lm_pp_param_decode("раскрасить_блоки", extra_args)
        if len(blocks) == 0:
            return None
        nm = str(sheet_name or "").strip()
        has_named = False
        bi = 0
        while bi < len(blocks):
            block = blocks[bi]
            bi = bi + 1
            sheet_ref = str(block.get("sheet") or "").strip()
            if sheet_ref == "":
                cfg = _lm_pp_colorize_config_from_block(block)
                _lm_pp_colorize_config_for_sheet_log_result(cfg)
                return cfg
            has_named = True
            if _lm_pp_sheet_name_matches_target(
                nm, _lm_pp_parse_sheet_names_from_spec(sheet_ref)
            ):
                cfg = _lm_pp_colorize_config_from_block(block)
                _lm_pp_colorize_config_for_sheet_log_result(cfg)
                return cfg
        if has_named:
            _lm_pp_colorize_config_for_sheet_log_result(None)
            return None
        cfg = _lm_pp_colorize_config_from_block(blocks[0])
        _lm_pp_colorize_config_for_sheet_log_result(cfg)
        return cfg
    assignments = _lm_pp_parse_colorize_assignments(text)
    if len(assignments) == 0:
        key_part, style_part = _lm_pp_split_colorize_key_and_style(text)
        if key_part == "":
            _lm_pp_colorize_config_for_sheet_log_result(None)
            return None
        cfg = _lm_pp_colorize_make_config(key_part, style_part)
        _lm_pp_colorize_config_for_sheet_log_result(cfg)
        return cfg
    nm = str(sheet_name or "").strip()
    has_named = False
    ai = 0
    while ai < len(assignments):
        sheet_names, key_part, style_part = assignments[ai]
        ai = ai + 1
        if sheet_names is None:
            cfg = _lm_pp_colorize_make_config(key_part, style_part)
            _lm_pp_colorize_config_for_sheet_log_result(cfg)
            return cfg
        has_named = True
        if _lm_pp_sheet_name_matches_target(nm, sheet_names):
            cfg = _lm_pp_colorize_make_config(key_part, style_part)
            _lm_pp_colorize_config_for_sheet_log_result(cfg)
            return cfg
    if has_named:
        _lm_pp_colorize_config_for_sheet_log_result(None)
        return None
    _lm_pp_colorize_config_for_sheet_log_result(None)
    return None


def _lm_pp_colorize_config_for_sheet_log_result(cfg):
    if not LIBRE_MACROS_DEBUG:
        return
    if cfg is None:
        _lm_dbg("colorize_config result=None")
        return
    colors = cfg.get("colors") or []
    resolved = []
    for c in colors:
        if isinstance(c, str):
            rc = _lm_pp_zebra_resolve_color_token(c)
        else:
            rc = c
        resolved.append("%s->%s" % (c, rc))
    _lm_dbg(
        "colorize_config keys=%s colors=%s resolved=%s outline=%s outline_color=%r"
        % (
            cfg.get("key_tokens"),
            colors,
            resolved,
            cfg.get("outline_blocks"),
            cfg.get("outline_color"),
        )
    )


def lm_pp_colorize_step_applies_to_sheet(extra_raw, sheet_name):
    """Нужно ли выполнять «раскрасить» на данном листе."""
    s = str(extra_raw or "").strip()
    if s == "":
        return False
    cfg = _lm_pp_colorize_config_for_sheet((s,), sheet_name)
    return cfg is not None


def _lm_pp_colorize_trim_key_text(val):
    """Трим ключа: все символы, считающиеся пробельными в Unicode."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, (int, float)):
        try:
            if isinstance(val, float) and val == int(val):
                return str(int(val))
        except (TypeError, ValueError, OverflowError):
            pass
        return str(val)
    return str(val).strip()


def _lm_pp_colorize_rgb_components(color):
    if color is None:
        return 0, 0, 0
    if isinstance(color, str):
        resolved = _lm_pp_zebra_resolve_color_token(color)
        if resolved is None:
            return 0, 0, 0
        color = int(resolved)
    else:
        try:
            color = int(color)
        except (TypeError, ValueError):
            return 0, 0, 0
    return (color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF


def _lm_pp_colorize_blend_rgb(c1, c2, t):
    r1, g1, b1 = _lm_pp_colorize_rgb_components(c1)
    r2, g2, b2 = _lm_pp_colorize_rgb_components(c2)
    t = max(0.0, min(1.0, float(t)))
    return lo_color_rgb(
        int(r1 + (r2 - r1) * t),
        int(g1 + (g2 - g1) * t),
        int(b1 + (b2 - b1) * t),
    )


def _lm_pp_colorize_auto_text_color(bg_color):
    r, g, b = _lm_pp_colorize_rgb_components(bg_color)
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255.0
    if luminance > 0.55:
        return lo_color_rgb(0, 0, 0)
    return lo_color_rgb(255, 255, 255)


def _lm_pp_colorize_shade_single(base_color, group_index, total_groups):
    white = lo_color_rgb(255, 255, 255)
    if total_groups <= 1:
        t = 0.42
    else:
        t = 0.14 + 0.72 * (float(group_index) / float(total_groups - 1))
    return _lm_pp_colorize_blend_rgb(white, base_color, t)


def _lm_pp_colorize_bg_for_group(colors, group_index, total_groups):
    if colors is None or len(colors) == 0:
        colors = [lo_color_rgb(217, 234, 247)]
    if len(colors) == 1:
        return _lm_pp_colorize_shade_single(colors[0], group_index, total_groups)
    return colors[group_index % len(colors)]


def _lm_pp_colorize_set_block_char_color(cell_range, color_rgb):
    """
    Цвет шрифта на всём блоке одним UNO-свойством (не по ячейкам).

    Для раскраски данных построчный CharColor даёт минуты на тысячах строк;
    здесь достаточно свойства диапазона. Fallback — _lm_set_char_color_safe.
    """
    if cell_range is None:
        return
    color_rgb = int(color_rgb)
    try:
        cell_range.CharAutoColor = False
    except Exception:
        pass
    ok = False
    try:
        cell_range.CharColor2 = color_rgb
        ok = True
    except Exception:
        pass
    try:
        cell_range.CharColor = color_rgb
        ok = True
    except Exception:
        pass
    if not ok:
        _lm_set_char_color_safe(cell_range, color_rgb)


def _lm_pp_colorize_build_formula(col_indices, row_1based):
    parts = []
    ci = 0
    while ci < len(col_indices):
        letter = col_index_to_letters(int(col_indices[ci]))
        parts.append("TRIM(%s%d)" % (letter, int(row_1based)))
        ci = ci + 1
    if len(parts) == 0:
        return '""'
    if len(parts) == 1:
        return parts[0]
    expr = parts[0]
    pi = 1
    while pi < len(parts):
        expr = "(%s)&(%s)" % (expr, parts[pi])
        pi = pi + 1
    return expr


def _lm_pp_calculate_doc(doc):
    """
    Пересчитать формулы книги (нужно перед getDataArray и «вставить как значение»).

    На время пересчёта временно включает авторасчёт, если макрос сбора его выключил
    (иначе calculateAll() может оставить в ячейках 0 до конца этапа 2).
    """
    _lm_pp_calculate_with_auto(doc, sheet=None)


def _lm_pp_calculate_with_auto(doc, sheet=None):
    """Включить авторасчёт при необходимости, пересчитать лист/книгу, восстановить режим."""
    if doc is None:
        return
    auto_was = None
    try:
        try:
            auto_was = bool(doc.isAutomaticCalculationEnabled())
        except Exception:
            try:
                auto_was = bool(doc.getPropertyValue("IsAutomaticCalculation"))
            except Exception:
                auto_was = None
        if auto_was is False:
            try:
                doc.enableAutomaticCalculation(True)
            except Exception:
                pass
        if sheet is not None:
            try:
                from com.sun.star.sheet import XCalculatable

                calc = sheet.queryInterface(XCalculatable)
                if calc is not None:
                    calc.calculate()
                else:
                    doc.calculateAll()
            except Exception:
                doc.calculateAll()
        else:
            doc.calculateAll()
    except Exception:
        pass
    finally:
        if auto_was is False:
            try:
                doc.enableAutomaticCalculation(False)
            except Exception:
                pass


def _lm_pp_undo_begin_batch(doc):
    """Заблокировать запись undo на время пакетной операции (без clipboard/dispatch)."""
    state = {"um": None, "locked": False}
    if doc is None:
        return state
    try:
        from com.sun.star.document import XUndoManagerSupplier

        sup = doc.queryInterface(XUndoManagerSupplier)
        if sup is None:
            return state
        um = sup.getUndoManager()
        if um is not None:
            um.lock()
            state["um"] = um
            state["locked"] = True
    except Exception:
        pass
    return state


def _lm_pp_undo_end_batch(doc, state, where="", clear=True):
    um = None
    if isinstance(state, dict):
        um = state.get("um")
    if um is not None and state.get("locked"):
        try:
            um.unlock()
        except Exception:
            pass
    if clear:
        _lm_pp_undo_clear(doc, where)


def _lm_pp_autocalc_suspend(doc):
    """Выключить авторасчёт на время пакетной записи; вернуть прежнее состояние или None."""
    if doc is None:
        return None
    try:
        auto_was = bool(doc.isAutomaticCalculationEnabled())
    except Exception:
        try:
            auto_was = bool(doc.getPropertyValue("IsAutomaticCalculation"))
        except Exception:
            return None
    if auto_was:
        try:
            doc.enableAutomaticCalculation(False)
        except Exception:
            pass
    return auto_was


def _lm_pp_autocalc_restore(doc, auto_was):
    if doc is None or auto_was is None:
        return
    try:
        doc.enableAutomaticCalculation(bool(auto_was))
    except Exception:
        pass


def _lm_pp_undo_clear(doc, where=""):
    """Очистить стеки Undo/Redo после Copy/Paste формул."""
    if doc is None:
        return
    um = None
    try:
        from com.sun.star.document import XUndoManagerSupplier

        sup = doc.queryInterface(XUndoManagerSupplier)
        if sup is None:
            return
        um = sup.getUndoManager()
        um.clear()
        if where:
            _lm_pp_apply_formula_dbg(doc, "", "undo_clear: %s" % where)
    except Exception:
        if um is not None:
            try:
                um.reset()
                if where:
                    _lm_pp_apply_formula_dbg(doc, "", "undo_clear reset: %s" % where)
            except Exception:
                pass


def _lm_pp_colorize_calculate_doc(doc):
    _lm_pp_calculate_doc(doc)


def _lm_pp_colorize_sort_key_text(val):
    s = _lm_pp_colorize_trim_key_text(val)
    return s.casefold()


def _lm_pp_colorize_row_list_from_block(block, rel_key):
    """Строки блока + ключ сортировки из служебного столбца (для раскрасить_блоки)."""
    row_list = []
    if block is None:
        return row_list
    ri = 0
    while ri < len(block):
        row = block[ri]
        if not isinstance(row, tuple):
            row = (row,)
        row = tuple(row)
        sort_key = ""
        if 0 <= int(rel_key) < len(row):
            sort_key = _lm_pp_colorize_sort_key_text(row[rel_key])
        row_list.append((sort_key, row))
        ri += 1
    return row_list


def _lm_pp_range_colorize_data_core( doc, sheet, data_range, header_row_range, extra_args, fn_label="раскрасить" ):
    """Служебная колонка-ключ, опциональная сортировка, раскраска блоков, удаление колонки."""
    sheet_name = ""
    try:
        sheet_name = str(sheet.Name) if sheet is not None else ""
    except Exception:
        pass
    cfg = _lm_pp_colorize_config_for_sheet(extra_args, sheet_name)
    if cfg is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            fn_label,
            _lm_pp_sort_text_from_extra_args(extra_args),
            "пропуск",
            "пустая C или лист не указан",
        )
        return False, "пропуск"

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    if er < sr:
        _lm_log_postprocess(
            doc, sheet_name, fn_label, "—", "пропуск", "нет строк данных"
        )
        return False, "нет строк"

    key_cols = []
    ti = 0
    while ti < len(cfg.get("key_tokens") or []):
        col_idx = _lm_pp_resolve_column_token(
            cfg["key_tokens"][ti], sheet, header_row_range, sc, ec
        )
        if col_idx is None:
            _lm_log_postprocess(
                doc,
                sheet_name,
                fn_label,
                cfg["key_tokens"][ti],
                "ошибка",
                "столбец ключа не найден: %s" % cfg["key_tokens"][ti],
            )
            return False, "столбец не найден"
        if col_idx not in key_cols:
            key_cols.append(col_idx)
        ti = ti + 1

    colors = cfg.get("colors") or []
    temp_col = int(ec) + 1
    first_formula = "=" + _lm_pp_colorize_build_formula(key_cols, int(sr) + 1)
    if not _lm_pp_set_cell_formula(
        doc,
        sheet.getCellByPosition(temp_col, sr),
        first_formula,
        sr,
        temp_col,
        sheet_name,
        allow_formula_error=True,
    ):
        _lm_log_postprocess(
            doc,
            sheet_name,
            fn_label,
            "—",
            "ошибка",
            "не удалось записать формулу ключа «%s»" % first_formula[:120],
        )
        return False, "формула"
    if int(er) > int(sr):
        if not _lm_pp_fill_formula_down(
            doc,
            sheet,
            temp_col,
            sr,
            er,
            sheet_name,
            dbg_channel="раскрасить_блоки",
        ):
            _lm_log_postprocess(
                doc,
                sheet_name,
                fn_label,
                "—",
                "ошибка",
                "не удалось протянуть формулу ключа (столбец %d)" % (temp_col + 1),
            )
            return False, "протяжка"
    _lm_pp_colorize_calculate_doc(doc)

    _lm_pp_colorize_calculate_doc(doc)

    row_list = None
    total_groups = 1
    try:
        try:
            block = sheet.getCellRangeByPosition(
                sc, sr, temp_col, er
            ).getDataArray()
        except Exception as err:
            _lm_log_postprocess(
                doc, sheet_name, fn_label, "—", "ошибка", "getDataArray: %s" % err
            )
            return False, str(err)

        if block is None or len(block) == 0:
            _lm_log_postprocess(
                doc, sheet_name, fn_label, "—", "пропуск", "пустой диапазон"
            )
            return False, "пусто"

        row_list = _lm_pp_colorize_row_list_from_block(block, temp_col - sc)

        sort_asc = cfg.get("sort")
        if sort_asc is not None:
            if _lm_pp_sort_use_native():
                ok, err = _lm_pp_range_sort_dispatch(
                    doc,
                    sheet,
                    sc,
                    sr,
                    temp_col,
                    er,
                    [(temp_col, bool(sort_asc))],
                    sc,
                    sheet_name=sheet_name,
                    fn_label=fn_label,
                )
                if not ok:
                    _lm_log_postprocess(
                        doc,
                        sheet_name,
                        fn_label,
                        "—",
                        "ошибка",
                        "сортировка (native): %s" % (err or "—"),
                    )
                    return False, err or "сортировка"
                try:
                    block = sheet.getCellRangeByPosition(
                        sc, sr, temp_col, er
                    ).getDataArray()
                except Exception as err:
                    _lm_log_postprocess(
                        doc,
                        sheet_name,
                        fn_label,
                        "—",
                        "ошибка",
                        "getDataArray после sort: %s" % err,
                    )
                    return False, str(err)
                row_list = _lm_pp_colorize_row_list_from_block(block, temp_col - sc)
            else:
                row_list.sort(key=lambda item: item[0], reverse=not sort_asc)
                sorted_rows = []
                si = 0
                while si < len(row_list):
                    sorted_rows.append(row_list[si][1])
                    si = si + 1
                try:
                    sheet.getCellRangeByPosition(sc, sr, temp_col, er).setDataArray(
                        tuple(sorted_rows)
                    )
                except Exception as err:
                    _lm_log_postprocess(
                        doc, sheet_name, fn_label, "—", "ошибка", "setDataArray: %s" % err
                    )
                    return False, str(err)

        distinct_keys = []
        prev_key = None
        gi = 0
        while gi < len(row_list):
            k = row_list[gi][0]
            if k != prev_key:
                distinct_keys.append(k)
                prev_key = k
            gi = gi + 1
        total_groups = len(distinct_keys)
        if total_groups == 0:
            total_groups = 1

        key_to_group = {}
        g = 0
        while g < len(distinct_keys):
            key_to_group[distinct_keys[g]] = g
            g = g + 1

        # Заливка целыми непрерывными блоками одинакового ключа (не построчно):
        # иначе на ~4k×70 UNO-вызовов уходят минуты; CharColor — на весь блок.
        r = int(sr)
        pi = 0
        n_rows = len(row_list)
        while pi < n_rows:
            k = row_list[pi][0]
            group_index = key_to_group.get(k, 0)
            bg = _lm_pp_colorize_bg_for_group(colors, group_index, total_groups)
            fg = _lm_pp_colorize_auto_text_color(bg)
            pj = pi + 1
            while pj < n_rows and row_list[pj][0] == k:
                pj = pj + 1
            block_start = r
            block_end = r + (pj - pi) - 1
            try:
                block_range = sheet.getCellRangeByPosition(sc, block_start, ec, block_end)
                block_range.CellBackColor = bg
                _lm_pp_colorize_set_block_char_color(block_range, fg)
            except Exception:
                pass
            if _lm_ui_should_update_status(pj, is_last=(pj >= n_rows)):
                _lm_ui_status_set(
                    pj,
                    _lm_pp_status_text(
                        fn_label, "строка %d/%d" % (pj, n_rows)
                    ),
                )
            _lm_ui_yield(counter=pj, force=(pj >= n_rows))
            r = block_end + 1
            pi = pj

        if cfg.get("outline_blocks"):
            outline_w = 35
            outline_clr = _lm_pp_colorize_resolve_outline_color(
                cfg.get("outline_color")
            )
            gi = 0
            while gi < len(row_list):
                k = row_list[gi][0]
                block_start = int(sr) + gi
                gj = gi + 1
                while gj < len(row_list) and row_list[gj][0] == k:
                    gj = gj + 1
                block_end = int(sr) + gj - 1
                _lm_pp_colorize_apply_block_outline(
                    sheet, sc, ec, block_start, block_end, outline_clr, outline_w
                )
                gi = gj
    finally:
        _lm_sheet_delete_columns(sheet, temp_col, 1)

    _lm_log_postprocess(
        doc,
        sheet_name,
        fn_label,
        _lm_pp_sort_text_from_extra_args(extra_args),
        "ok",
        "строк %d, групп %d, столбцы %s"
        % (
            len(row_list) if row_list is not None else 0,
            total_groups,
            ",".join(str(c + 1) for c in key_cols),
        ),
    )
    return True, ""


def lm_pp_range_colorize_data(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Назначение:
        Раскрасить строки данных блоками по смене составного ключа столбцов.

    JSON колонки C: [{"v":1,"fn":"раскрасить_блоки","key_columns":["ФИО"],"colors":["голубой","желтый"],"sort":"asc","outline_blocks":true}]
    или с "sheet":"Лист". См. docs/13_JSON_PARAMS.md.

    Карта: «раскрасить».

    Алгоритм:
        Временная колонка с формулой TRIM-конкатенации ключей → при указании
        направления сортировка → заливка блоков → удаление колонки.
    """
    sheet_name = ""
    try:
        sheet_name = str(sheet.Name) if sheet is not None else ""
    except Exception:
        pass
    if _lm_pp_param_skip_sheet("раскрасить_блоки", extra_args, doc, sheet):
        return
    _lm_pp_range_colorize_data_core(
        doc, sheet, data_range, header_row_range, extra_args, fn_label="раскрасить"
    )


def _lm_pp_copy_cell_value_down(src_cell, dst_cell):
    """Скопировать значение и оформление из src в dst (для fill-down)."""
    if src_cell is None or dst_cell is None:
        return False

    def _cell_is_empty(cell):
        # Пустая ячейка в Calc: нет формулы, нет строки, Value==0.
        try:
            f = cell.Formula
            if f is not None and str(f).strip() != "":
                return False
        except Exception:
            pass
        try:
            s = cell.String
            if s is not None and str(s).strip() != "":
                return False
        except Exception:
            pass
        try:
            v = cell.Value
            if v not in (None, 0.0, 0):
                return False
        except Exception:
            pass
        return True

    if not _cell_is_empty(dst_cell):
        return False
    if _cell_is_empty(src_cell):
        return False

    def _copy_format(src, dst):
        for attr in (
            "NumberFormat",
            "CharFontName",
            "CharHeight",
            "CharWeight",
            "CellBackColor",
            "HoriJustify",
            "VertJustify",
            "IsTextWrapped",
            "TableBorder",
            "CellProtection",
        ):
            try:
                setattr(dst, attr, getattr(src, attr))
            except Exception:
                pass
        try:
            _lm_set_char_color_safe(dst, getattr(src, "CharColor", None))
        except Exception:
            pass
    try:
        dst_cell.String = src_cell.String
        _copy_format(src_cell, dst_cell)
        return True
    except Exception:
        try:
            dst_cell.Value = src_cell.Value
            _copy_format(src_cell, dst_cell)
            return True
        except Exception:
            return False


def _lm_pp_sheet_same(sheet_a, sheet_b):
    """Сравнение листов по имени (в PyUNO «==» по объектам ненадёжен)."""
    if sheet_a is None or sheet_b is None:
        return sheet_a is sheet_b
    try:
        return str(sheet_a.Name) == str(sheet_b.Name)
    except Exception:
        try:
            return sheet_a == sheet_b
        except Exception:
            return False


def _lm_pp_unique_sheet_name(doc, base_name, exclude_sheet=None):
    """
    Уникальное имя листа в книге (макс. 31 символ).

    exclude_sheet — текущий лист: его имя не считается занятым.
    При конфликте добавляется суффикс _2, _3, …
    """
    base = str(base_name or "").strip()
    if base == "":
        return ""
    max_len = 31
    exclude_name = ""
    if exclude_sheet is not None:
        try:
            exclude_name = str(exclude_sheet.Name).strip()
        except Exception:
            exclude_name = ""

    def _taken(name):
        if doc is None:
            return True
        try:
            sheets = doc.getSheets()
            if not sheets.hasByName(name):
                return False
            if exclude_sheet is None and exclude_name == "":
                return True
            if exclude_name != "" and name == exclude_name:
                return False
            try:
                other = sheets.getByName(name)
                if _lm_pp_sheet_same(other, exclude_sheet):
                    return False
            except Exception:
                pass
            return True
        except Exception:
            return True

    if not _taken(base):
        return base[:max_len]

    suffix = 2
    while suffix <= 999:
        tail = "_%d" % suffix
        head_len = max_len - len(tail)
        if head_len < 1:
            head_len = 1
        candidate = base[:head_len] + tail
        if not _taken(candidate):
            return candidate
        suffix = suffix + 1
    return base[:max_len]


def _lm_pp_get_sheet_ref(doc, spec):
    """Лист книги по имени (без учёта регистра) или 1-based индексу."""
    s = str(spec or "").strip()
    if s == "" or doc is None:
        return None
    try:
        if s.isdigit():
            idx = int(s) - 1
            if idx >= 0:
                return doc.Sheets.getByIndex(idx)
    except Exception:
        pass
    try:
        sheets = doc.Sheets
        if sheets.hasByName(s):
            return sheets.getByName(s)
    except Exception:
        pass
    key = s.casefold()
    try:
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            sh = sheets.getByIndex(i)
            if str(sh.Name).casefold() == key:
                return sh
            i = i + 1
    except Exception:
        pass
    return None


def _lm_pp_allocate_new_header_name(sheet, header_row, base_name):
    """
    Имя заголовка для нового столбца: base, base_2, base_3, … пока нет в строке заголовков.
    """
    base = str(base_name or "").strip()
    if base == "":
        base = "_Столбец"
    try:
        header_row = int(header_row)
    except (TypeError, ValueError):
        header_row = 0
    end_col, _end_row = get_sheet_used_bounds(sheet)
    candidate = base
    n = 2
    while True:
        hkey = lm_identity_key(candidate)
        tc = 0
        while tc <= end_col:
            if lm_identity_key(cell_text(sheet.getCellByPosition(tc, header_row))) == hkey:
                candidate = "%s_%d" % (base, n)
                n = n + 1
                break
            tc = tc + 1
        else:
            return candidate
        if n > 200:
            return candidate


def _lm_pp_value_count_header_name(block, sheet, header_row):
    """Имя столбца из JSON или служебное _Количество_значений[_N]."""
    col = str(block.get("new_column") or block.get("column") or "").strip()
    if col != "":
        return col
    return _lm_pp_allocate_new_header_name(sheet, header_row, "_Количество_значений")


def _lm_pp_value_count_key_from_parts(parts, trim=True, case_sensitive=False):
    """Составной ключ строки для «количество_значений»."""
    sep = "\x1f"
    out = []
    i = 0
    while i < len(parts):
        s = str(parts[i] if parts[i] is not None else "")
        if trim:
            s = s.strip()
        if not case_sensitive:
            s = s.casefold()
        out.append(s)
        i = i + 1
    return sep.join(out)


def _lm_pp_value_count_row_key(sheet, row_index, key_cols, trim, case_sensitive):
    parts = []
    i = 0
    while i < len(key_cols):
        col = key_cols[i]
        parts.append(cell_text(sheet.getCellByPosition(col, row_index)))
        i = i + 1
    return _lm_pp_value_count_key_from_parts(parts, trim, case_sensitive)


def _lm_pp_copy_sheet_guess_header_row(sheet):
    """Строка заголовков на скопированном листе (0-based)."""
    if sheet is None:
        return 0
    try:
        cursor = sheet.createCursor()
        if cursor is not None:
            cursor.gotoStartOfUsedArea(False)
            return int(cursor.getRangeAddress().StartRow)
    except Exception:
        pass
    return 0


_LM_PP_COPY_SHEET_PENDING = []
_LM_FINAL_ACTIVATE_SHEET_NAME = ""


def lm_final_activate_sheet_clear():
    """Сброс имени листа из финального «активировать_лист» (начало сбора)."""
    global _LM_FINAL_ACTIVATE_SHEET_NAME
    _LM_FINAL_ACTIVATE_SHEET_NAME = ""


def lm_final_activate_sheet_remember(sheet_name):
    """Запомнить лист после успешного «активировать_лист» (чтобы не перебить в конце сбора)."""
    global _LM_FINAL_ACTIVATE_SHEET_NAME
    nm = str(sheet_name or "").strip()
    if nm != "":
        _LM_FINAL_ACTIVATE_SHEET_NAME = nm


def lm_final_activate_sheet_name():
    """Имя листа из последнего успешного «активировать_лист» или пусто."""
    return str(_LM_FINAL_ACTIVATE_SHEET_NAME or "").strip()


def lm_pp_copy_sheet_clear_pending():
    """Сброс очереди листов после «копировать_лист» (начало сбора)."""
    global _LM_PP_COPY_SHEET_PENDING
    _LM_PP_COPY_SHEET_PENDING = []
    try:
        from libre_macros_create_sheet_lib import clear_create_sheet_ran

        clear_create_sheet_ran()
    except Exception:
        pass


def lm_pp_copy_sheet_remove_pending_names(names):
    """Убрать из очереди указанные имена листов (без сброса остальных)."""
    global _LM_PP_COPY_SHEET_PENDING
    drop = set()
    for n in names or []:
        s = str(n or "").strip()
        if s != "":
            drop.add(s)
    if not drop:
        return 0
    kept = []
    removed = 0
    for entry in list(_LM_PP_COPY_SHEET_PENDING or []):
        name = str((entry or {}).get("name") or "").strip()
        if name in drop:
            removed = removed + 1
            continue
        kept.append(entry)
    _LM_PP_COPY_SHEET_PENDING = kept
    return removed


def lm_pp_copy_sheet_register_pending( dest_sheet_name, header_row=0, source_sheet_name="", from_dedup=False ):
    """Зарегистрировать лист-приёмник для постобработки после листов из источников."""
    global _LM_PP_COPY_SHEET_PENDING
    name = str(dest_sheet_name or "").strip()
    if name == "":
        return
    i = 0
    while i < len(_LM_PP_COPY_SHEET_PENDING):
        if str(_LM_PP_COPY_SHEET_PENDING[i].get("name") or "") == name:
            return
        i = i + 1
    _LM_PP_COPY_SHEET_PENDING.append(
        {
            "name": name,
            "header_row": int(header_row),
            "source": str(source_sheet_name or "").strip(),
            "from_dedup": bool(from_dedup),
        }
    )


def lm_pp_copy_sheet_take_pending():
    """Снять и очистить очередь скопированных листов."""
    global _LM_PP_COPY_SHEET_PENDING
    out = list(_LM_PP_COPY_SHEET_PENDING)
    _LM_PP_COPY_SHEET_PENDING = []
    return out


def lm_pp_copy_sheet_list_pending():
    """Копия очереди без сброса."""
    return list(_LM_PP_COPY_SHEET_PENDING)


def _lm_pp_col_letters(col_index):
    n = int(col_index)
    if n < 0:
        n = 0
    s = ""
    while True:
        s = chr(65 + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def _lm_pp_clear_sheet_content(sheet):
    if sheet is None:
        return
    try:
        cursor = sheet.createCursor()
        cursor.gotoStartOfUsedArea(False)
        cursor.gotoEndOfUsedArea(True)
        sheet.getCellRangeByCursor(cursor).clearContents(7)
        return
    except Exception:
        pass
    end_col, end_row = get_sheet_used_bounds(sheet)
    r = 0
    while r <= end_row:
        c = 0
        while c <= end_col:
            try:
                cell = sheet.getCellByPosition(c, r)
                cell.setString("")
            except Exception:
                pass
            c = c + 1
        r = r + 1


def _lm_pp_copy_cell_same_doc(src_cell, dst_cell):
    """Значение + NumberFormat (без визуального оформления)."""
    _lm_pp_copy_cell_same_doc_ex(src_cell, dst_cell, with_formatting=False)


def _lm_pp_copy_cell_format_same_doc(src_cell, dst_cell):
    """
    Только оформление ячейки в той же книге (NumberFormat + визуал).

    Надёжнее буфера обмена при объединении нескольких листов подряд.
    """
    if src_cell is None or dst_cell is None:
        return
    try:
        dst_cell.NumberFormat = src_cell.NumberFormat
    except Exception:
        pass
    try:
        transparent = bool(src_cell.IsCellBackgroundTransparent)
        dst_cell.IsCellBackgroundTransparent = transparent
        if not transparent:
            dst_cell.CellBackColor = int(src_cell.CellBackColor)
    except Exception:
        try:
            dst_cell.CellBackColor = int(src_cell.CellBackColor)
            dst_cell.IsCellBackgroundTransparent = False
        except Exception:
            pass
    for attr in (
        "CharFontName",
        "CharHeight",
        "CharWeight",
        "CharPosture",
        "HoriJustify",
        "VertJustify",
        "IsTextWrapped",
        "TopBorder",
        "BottomBorder",
        "LeftBorder",
        "RightBorder",
    ):
        try:
            setattr(dst_cell, attr, getattr(src_cell, attr))
        except Exception:
            pass
    # Цвет шрифта: CharAutoColor + CharColor2 (как при сборе)
    auto = True
    try:
        auto = bool(getattr(src_cell, "CharAutoColor", True))
        dst_cell.CharAutoColor = auto
    except Exception:
        auto = True
    if not auto:
        color_rgb = None
        try:
            color_rgb = int(getattr(src_cell, "CharColor2", None))
        except Exception:
            color_rgb = None
        if color_rgb is None:
            try:
                color_rgb = int(getattr(src_cell, "CharColor", None))
            except Exception:
                color_rgb = None
        if color_rgb is not None:
            try:
                _lm_set_char_color_cell(dst_cell, color_rgb)
            except Exception:
                try:
                    _lm_set_char_color_safe(dst_cell, color_rgb)
                except Exception:
                    pass


def _lm_pp_copy_cell_same_doc_ex(src_cell, dst_cell, with_formatting=False):
    """
    Копия ячейки в той же книге.

    always: значение/формула + NumberFormat.
    with_formatting: фон, шрифт, цвет, выравнивание, границы, перенос.
    """
    if src_cell is None or dst_cell is None:
        return
    wrote = False
    try:
        f = src_cell.Formula
        if f is not None and str(f).strip() != "":
            dst_cell.Formula = f
            wrote = True
    except Exception:
        pass
    if not wrote:
        try:
            # Числа/даты надёжнее через Value (String ломает NumberFormat).
            ctype = None
            try:
                ctype = int(src_cell.getType())
            except Exception:
                ctype = None
            # VALUE=1, FORMULA=3
            if ctype == 1:
                dst_cell.Value = src_cell.Value
                wrote = True
            else:
                s = src_cell.String
                if s is not None and str(s).strip() != "":
                    dst_cell.String = s
                    wrote = True
        except Exception:
            pass
    if not wrote:
        try:
            dst_cell.Value = src_cell.Value
            wrote = True
        except Exception:
            try:
                dst_cell.String = ""
            except Exception:
                pass
    if with_formatting:
        _lm_pp_copy_cell_format_same_doc(src_cell, dst_cell)
    else:
        try:
            dst_cell.NumberFormat = src_cell.NumberFormat
        except Exception:
            pass


def _lm_pp_guess_merge_header_row(sheet, default=0):
    """
    Строка заголовков (0-based) для объединить_листы_в_один.

    Первая строка с ≥2 непустыми ячейками; иначе — StartRow используемой области.
    """
    end_col, end_row = get_sheet_used_bounds(sheet)
    if end_row < 0 or end_col < 0:
        return default
    try:
        cursor = sheet.createCursor()
        if cursor is not None:
            cursor.gotoStartOfUsedArea(False)
            default = int(cursor.getRangeAddress().StartRow)
    except Exception:
        pass
    r = 0
    while r <= min(end_row, 50):
        cnt = 0
        c = 0
        while c <= end_col:
            if cell_text(sheet.getCellByPosition(c, r)).strip() != "":
                cnt = cnt + 1
                if cnt >= 2:
                    return r
            c = c + 1
        r = r + 1
    if default < 0:
        return 0
    return default


def _lm_pp_parse_merge_header_row(sheet, header_row_param):
    """
    header_row из JSON — номер строки в Calc (1-based), как «Строка_Заголовков».

    Пусто — автоопределение (_lm_pp_guess_merge_header_row).
    """
    if header_row_param is None or str(header_row_param).strip() == "":
        return _lm_pp_guess_merge_header_row(sheet)
    try:
        hr = int(header_row_param) - 1
    except (TypeError, ValueError):
        return _lm_pp_guess_merge_header_row(sheet)
    if hr < 0:
        hr = 0
    return hr


def _lm_pp_sheet_columns_by_header(sheet, header_row):
    end_col, end_row = get_sheet_used_bounds(sheet)
    try:
        header_row = int(header_row)
    except (TypeError, ValueError):
        header_row = 0
    start_col = 0
    while start_col <= end_col:
        if cell_text(sheet.getCellByPosition(start_col, header_row)).strip() != "":
            break
        start_col = start_col + 1
    if start_col > end_col:
        return [], end_row
    cols = []
    c = start_col
    while c <= end_col:
        t = cell_text(sheet.getCellByPosition(c, header_row)).strip()
        if t == "":
            t = "Колонка_%s" % _lm_pp_col_letters(c)
        cols.append((c, t))
        c = c + 1
    return cols, end_row


def _lm_pp_row_empty_in_first_col(sheet, row, first_col):
    return cell_text(sheet.getCellByPosition(first_col, row)).strip() == ""


def _lm_pp_doc_sheet_names(doc):
    names = []
    if doc is None:
        return names
    try:
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            sh = sheets.getByIndex(i)
            if sh is not None:
                names.append(str(sh.Name))
            i = i + 1
    except Exception:
        pass
    return names


def _lm_pp_resolve_merge_source_sheet_names(doc, sheets_spec):
    if isinstance(sheets_spec, (list, tuple)):
        parts = [str(x or "").strip() for x in sheets_spec]
    else:
        parts = []
        for chunk in str(sheets_spec or "").replace(";", ",").split(","):
            st = chunk.strip()
            if st != "":
                parts.append(st)
    parts = [p for p in parts if p != ""]
    if not parts:
        return []
    all_names = _lm_pp_doc_sheet_names(doc)
    out = []
    seen = set()
    pi = 0
    while pi < len(parts):
        part = parts[pi]
        pi = pi + 1
        if "*" in part or "?" in part:
            ni = 0
            while ni < len(all_names):
                nm = all_names[ni]
                if text_match(nm, part):
                    key = lm_identity_key(nm)
                    if key not in seen:
                        seen.add(key)
                        out.append(nm)
                ni = ni + 1
            continue
        sh = _lm_pp_get_sheet_ref(doc, part)
        if sh is None:
            continue
        nm = str(sh.Name)
        key = lm_identity_key(nm)
        if key not in seen:
            seen.add(key)
            out.append(nm)
    return out


def _lm_pp_iter_consecutive_row_chunks(rows):
    """Разбить список 0-based индексов строк на подряд идущие фрагменты."""
    if not rows:
        return
    start = 0
    i = 1
    while i <= len(rows):
        if i == len(rows) or rows[i] != rows[i - 1] + 1:
            yield rows[start:i], start
            start = i
        i = i + 1


def _lm_pp_clipboard_paste_at_cell(doc, dst_sheet, tc0, tr0):
    """
    Вставка из буфера в ячейку-якорь (как `_merge_clipboard_paste_at` в сборе).

    Сначала .uno:Paste (полный перенос стилей), затем paste/pasteCellRange
    на CellAddress с полным набором CellFlags.
    """
    if doc is None or dst_sheet is None:
        return False
    try:
        dst_cell = dst_sheet.getCellByPosition(tc0, tr0)
    except Exception:
        return False
    if not _lm_pp_select_cell_range(doc, dst_sheet, dst_cell):
        return False
    if _lm_pp_execute_dispatch(doc, ".uno:Paste"):
        return True
    paste_op = 0
    insert_mode = 0
    n_flags = 65535
    try:
        from com.sun.star.sheet.PasteOperation import NONE as PASTE_NONE
        from com.sun.star.sheet.CellInsertMode import NONE as INSERT_NONE
        from com.sun.star.sheet.CellFlags import (
            VALUE,
            DATETIME,
            STRING,
            ANNOTATION,
            FORMULA,
            HARDATTR,
            STYLES,
            OBJECTS,
            EDITATTR,
            FORMULA_FORMAT,
            FORMATTED,
        )

        paste_op = PASTE_NONE
        insert_mode = INSERT_NONE
        n_flags = (
            int(VALUE)
            | int(DATETIME)
            | int(STRING)
            | int(ANNOTATION)
            | int(FORMULA)
            | int(HARDATTR)
            | int(STYLES)
            | int(OBJECTS)
            | int(EDITATTR)
            | int(FORMULA_FORMAT)
            | int(FORMATTED)
        )
    except Exception:
        pass
    try:
        addr = dst_cell.getCellAddress()
        dst_sheet.paste(addr)
        return True
    except Exception:
        pass
    try:
        dest_range = dst_sheet.getCellRangeByPosition(tc0, tr0, tc0, tr0)
        dest_addr = dest_range.getRangeAddress()
        dst_sheet.pasteCellRange(
            dest_addr,
            paste_op,
            n_flags,
            True,
            False,
            False,
            insert_mode,
        )
        return True
    except Exception:
        pass
    return False


def _lm_pp_clipboard_copy_paste_range( doc, src_sheet, sc0, sr0, sc1, sr1, dst_sheet, tc0, tr0 ):
    """
    Скопировать прямоугольник источника через буфер в ячейку приёмника
    (как сбор «буфер по заголовкам»: Copy → yield → Paste).
    """
    if doc is None or src_sheet is None or dst_sheet is None:
        return False
    if sr1 < sr0 or sc1 < sc0:
        return False
    try:
        src_range = src_sheet.getCellRangeByPosition(sc0, sr0, sc1, sr1)
    except Exception:
        return False
    if not _lm_pp_copy_range_to_clipboard(doc, src_sheet, src_range):
        return False
    _lm_ui_yield(force=True)
    return _lm_pp_clipboard_paste_at_cell(doc, dst_sheet, tc0, tr0)


def _lm_pp_merge_transferable_paste_column( doc, src, sc, sr0, sr1, dest, tc, tr0 ):
    """
    Та же книга: контроллер getTransferable → insertTransferable
    (часто стабильнее системного буфера на 2+ источнике).
    """
    if doc is None or src is None or dest is None:
        return False
    if sr1 < sr0:
        return False
    try:
        controller = doc.getCurrentController()
        if controller is None:
            return False
        src_range = src.getCellRangeByPosition(sc, sr0, sc, sr1)
        dst_cell = dest.getCellByPosition(tc, tr0)
        try:
            if controller.getActiveSheet().Name != src.Name:
                controller.setActiveSheet(src)
        except Exception:
            controller.setActiveSheet(src)
        controller.select(src_range)
        transferable = controller.getTransferable()
        if transferable is None:
            return False
        try:
            if controller.getActiveSheet().Name != dest.Name:
                controller.setActiveSheet(dest)
        except Exception:
            controller.setActiveSheet(dest)
        controller.select(dst_cell)
        controller.insertTransferable(transferable)
        return True
    except Exception:
        return False


def _lm_pp_merge_formats_ok_sample(src, sc, sr0, dest, tc, tr0):
    """
    Проба: стили реально попали на dest (CellStyle и/или HARDATTR фона).
    Нужна, чтобы не останавливаться на «успешном» Paste только со значениями.
    """
    try:
        scell = src.getCellByPosition(sc, sr0)
        dcell = dest.getCellByPosition(tc, tr0)
    except Exception:
        return False
    try:
        ss = str(scell.CellStyle or "").strip()
        ds = str(dcell.CellStyle or "").strip()
        if ss != "" and ss != ds:
            low = ss.casefold()
            if low not in ("default", "default style", "обычный"):
                return False
    except Exception:
        pass
    try:
        if not bool(scell.IsCellBackgroundTransparent):
            if bool(dcell.IsCellBackgroundTransparent):
                return False
            if int(scell.CellBackColor) != int(dcell.CellBackColor):
                return False
    except Exception:
        pass
    try:
        if int(scell.CharColor) != int(dcell.CharColor):
            return False
    except Exception:
        pass
    return True


def _lm_pp_insert_contents_values_and_formats_props():
    """InsertContents: значения + форматы (SVDT), без формул."""
    props = []
    for name, value in (
        ("Flags", "SVDT"),
        ("FormulaCommand", 0),
        ("SkipEmptyCells", False),
        ("Transpose", False),
        ("AsLink", False),
        ("MoveMode", 4),
        ("Overwrite", True),
    ):
        prop = PropertyValue()
        prop.Name = name
        prop.Value = value
        props.append(prop)
    return tuple(props)


def _lm_pp_merge_paste_column_subchunk_formats( doc, src, sc, sr0, sr1, dest, tc, tr0 ):
    """
    Один столбец-подчанок со стилями (без поячеечного fallback):

      1) Copy → .uno:Paste / paste (как сбор «буфер по заголовкам»);
      2) Transferable той же книги;
      3) Copy → InsertContents SVDT.

    Если API вернул True, но проба стилей не прошла — пробуем следующий способ.
    """
    if doc is None or src is None or dest is None:
        return False
    try:
        n = int(sr1) - int(sr0) + 1
    except (TypeError, ValueError):
        return False
    if n <= 0:
        return False

    if _lm_pp_clipboard_copy_paste_range(
        doc, src, sc, sr0, sc, sr1, dest, tc, tr0
    ):
        if _lm_pp_merge_formats_ok_sample(src, sc, sr0, dest, tc, tr0):
            return True

    if _lm_pp_merge_transferable_paste_column(
        doc, src, sc, sr0, sr1, dest, tc, tr0
    ):
        if _lm_pp_merge_formats_ok_sample(src, sc, sr0, dest, tc, tr0):
            return True

    try:
        src_range = src.getCellRangeByPosition(sc, sr0, sc, sr1)
    except Exception:
        return False
    if not _lm_pp_copy_range_to_clipboard(doc, src, src_range):
        return False
    _lm_ui_yield(force=True)
    try:
        dst_cell = dest.getCellByPosition(tc, tr0)
    except Exception:
        return False
    if not _lm_pp_select_cell_range(doc, dest, dst_cell):
        return False
    if not _lm_pp_execute_dispatch(
        doc,
        ".uno:InsertContents",
        _lm_pp_insert_contents_values_and_formats_props(),
    ):
        return False
    return _lm_pp_merge_formats_ok_sample(src, sc, sr0, dest, tc, tr0)


def _lm_pp_merge_copy_column_chunk( doc, src, sc, sr0, sr1, dest, tc, tr0, with_formatting=False ):
    """
    Один столбец-чанк в dest.

    with_formatting=True — постолбцово чанками по _LM_PP_MERGE_FORMAT_CHUNK
      через Transferable / Copy+Paste (как сбор); без поячеечного fallback.
    False — значения + NumberFormat; визуал потом оформит дефолт сбора.
    """
    try:
        n = int(sr1) - int(sr0) + 1
    except (TypeError, ValueError):
        return False
    if n <= 0:
        return False

    if with_formatting:
        chunk = int(_LM_PP_MERGE_FORMAT_CHUNK) or 200
        r = int(sr0)
        while r <= int(sr1):
            r1 = r + chunk - 1
            if r1 > int(sr1):
                r1 = int(sr1)
            tr = int(tr0) + (r - int(sr0))
            ok = _lm_pp_merge_paste_column_subchunk_formats(
                doc, src, sc, r, r1, dest, tc, tr
            )
            if not ok:
                _lm_pp_merge_sheets_dbg(
                    "format",
                    "Paste FAIL col=%d rows %d..%d (без поячеечного fallback)"
                    % (sc, r, r1),
                )
            r = r1 + 1
            _lm_ui_yield(force=True)
        return True

    values_ok = False
    try:
        src_range = src.getCellRangeByPosition(sc, sr0, sc, sr1)
        data = src_range.getDataArray()
        dest_range = dest.getCellRangeByPosition(tc, tr0, tc, tr0 + n - 1)
        dest_range.setDataArray(data)
        values_ok = True
    except Exception:
        values_ok = False

    if values_ok:
        r = sr0
        off = 0
        while r <= sr1:
            try:
                dest.getCellByPosition(tc, tr0 + off).NumberFormat = src.getCellByPosition(
                    sc, r
                ).NumberFormat
            except Exception:
                pass
            r = r + 1
            off = off + 1
        return True
    r = sr0
    off = 0
    while r <= sr1:
        _lm_pp_copy_cell_same_doc_ex(
            src.getCellByPosition(sc, r),
            dest.getCellByPosition(tc, tr0 + off),
            with_formatting=False,
        )
        r = r + 1
        off = off + 1
    return False


def _lm_pp_merge_apply_default_data_format(dest, header_row=0):
    """Шрифт данных на приёмнике — как встроенное оформление при сборе."""
    if dest is None:
        return
    try:
        end_col, end_row = get_sheet_used_bounds(dest)
    except Exception:
        return
    hr = int(header_row or 0)
    if end_col < 0 or end_row <= hr:
        return
    try:
        data_range = dest.getCellRangeByPosition(0, hr + 1, end_col, end_row)
    except Exception:
        return
    if data_range is None:
        return
    try:
        data_range.CharFontName = _LM_PP_FONT_NAME
    except Exception:
        pass
    try:
        data_range.CharHeight = float(_LM_PP_DEFAULT_FONT_SIZE)
    except Exception:
        pass
    try:
        data_range.CharWeight = 100
    except Exception:
        pass


def _lm_pp_filter_merge_cols_by_names(cols, column_filter):
    """Оставить столбцы по именам/1-based номерам; пустой filter = все."""
    if not column_filter:
        return cols
    wanted_names = set()
    wanted_nums = set()
    for t in column_filter:
        if isinstance(t, int):
            wanted_nums.add(int(t))
            continue
        if isinstance(t, float):
            try:
                wanted_nums.add(int(t))
            except (TypeError, ValueError):
                pass
            continue
        s = str(t or "").strip()
        if s == "":
            continue
        if s.isdigit():
            wanted_nums.add(int(s))
        else:
            wanted_names.add(s.lower())
    if not wanted_names and not wanted_nums:
        return cols
    out = []
    for sc, hname in cols or []:
        keep = False
        if str(hname or "").strip().lower() in wanted_names:
            keep = True
        elif (int(sc) + 1) in wanted_nums:
            keep = True
        if keep:
            out.append((sc, hname))
    return out

def _lm_pp_merge_sheets_into_one_core( doc, source_names, dest_name, header_row=None, with_formatting=False, columns=None, auto_format=True):
    """
    Слияние листов «как сбор на один лист / буфер по заголовкам»:
      - столбцы по имени заголовка (find_or_append);
      - данные чанками по столбцу;
      - пустые строки по первому столбцу пропускаются (уплотнение).
      - with_formatting: Transferable/Copy+Paste столбца чанками (как сбор);
        иначе — NumberFormat (+ дефолтный шрифт данных, если auto_format);
      - columns: необязательный фильтр имён/номеров; пусто = все столбцы;
      - auto_format: оформление данных по умолчанию (если не with_formatting).
    """
    column_filter = list(columns or [])
    _lm_pp_merge_sheets_dbg(
        "begin",
        "sources=%s dest=%s header_row=%r with_formatting=%s columns=%r auto_format=%s"
        % (
            ",".join(str(s) for s in source_names),
            dest_name,
            header_row,
            with_formatting,
            column_filter,
            auto_format,
        ),
    )
    sources = []
    si = 0
    while si < len(source_names):
        sh = _lm_pp_get_sheet_ref(doc, source_names[si])
        if sh is None:
            _lm_pp_merge_sheets_dbg("error", "лист «%s» не найден" % source_names[si])
            return False, "лист «%s» не найден" % source_names[si], ""
        sources.append(sh)
        si = si + 1
    if not sources:
        _lm_pp_merge_sheets_dbg("error", "нет листов-источников")
        return False, "нет листов-источников", ""

    hdr = _lm_pp_parse_merge_header_row(sources[0], header_row)
    _lm_pp_merge_sheets_dbg("header", "0-based row=%d (1-based=%d)" % (hdr, hdr + 1))

    dest = _lm_pp_get_sheet_ref(doc, dest_name)
    actual_dest = str(dest_name).strip()
    if dest is None:
        try:
            pos = int(doc.Sheets.getCount())
            doc.Sheets.insertNewByName(actual_dest, pos)
            _lm_pp_merge_sheets_dbg("dest", "создан лист «%s» pos=%d" % (actual_dest, pos))
        except Exception as err:
            _lm_pp_merge_sheets_dbg("error", "создание «%s»: %s" % (actual_dest, err))
            return False, str(err), ""
        dest = _lm_pp_get_sheet_ref(doc, actual_dest)
        if dest is None:
            _lm_pp_merge_sheets_dbg("error", "не удалось получить «%s»" % actual_dest)
            return False, "не удалось создать «%s»" % actual_dest, ""
    else:
        actual_dest = str(dest.Name)
        _lm_pp_merge_sheets_dbg("dest", "очистка существующего «%s»" % actual_dest)
        _lm_pp_clear_sheet_content(dest)

    out_data_row = hdr + 1
    total_rows = 0
    si = 0
    while si < len(sources):
        src = sources[si]
        src_name = str(src.Name)
        all_cols, end_row = _lm_pp_sheet_columns_by_header(src, hdr)
        if len(all_cols) == 0:
            _lm_pp_merge_sheets_dbg("source", "«%s» пропуск — нет столбцов" % src_name)
            si = si + 1
            continue
        first_col = all_cols[0][0]
        cols = _lm_pp_filter_merge_cols_by_names(all_cols, column_filter)
        col_names = ",".join(str(c[1]) for c in cols[:12])
        if len(cols) > 12:
            col_names += ",…"
        _lm_pp_merge_sheets_dbg(
            "source",
            "«%s» cols=%d/%d end_row=%d headers=[%s]"
            % (src_name, len(cols), len(all_cols), end_row, col_names),
        )
        if len(cols) == 0:
            _lm_pp_merge_sheets_dbg(
                "source", "«%s» пропуск — нет столбцов после фильтра" % src_name
            )
            si = si + 1
            continue
        # Заголовки только текстом — оформление заголовка всегда по умолчанию (снаружи).
        ci = 0
        while ci < len(cols):
            sc, hname = cols[ci]
            tc = _lm_pp_find_or_append_header_column(dest, hname, hdr)
            if tc >= 0 and cell_text(dest.getCellByPosition(tc, hdr)).strip() == "":
                try:
                    dest.getCellByPosition(tc, hdr).String = hname
                except Exception:
                    pass
            ci = ci + 1
        data_rows = []
        r = hdr + 1
        while r <= end_row:
            if not _lm_pp_row_empty_in_first_col(src, r, first_col):
                data_rows.append(r)
            r = r + 1
        sheet_rows = len(data_rows)
        if sheet_rows == 0:
            _lm_pp_merge_sheets_dbg("source", "«%s» нет строк данных" % src_name)
            si = si + 1
            continue
        for chunk, chunk_off in _lm_pp_iter_consecutive_row_chunks(data_rows):
            sr0 = chunk[0]
            sr1 = chunk[-1]
            tr0 = out_data_row + chunk_off
            ci = 0
            while ci < len(cols):
                sc, hname = cols[ci]
                tc = _lm_pp_find_or_append_header_column(dest, hname, hdr)
                if tc >= 0:
                    _lm_pp_merge_copy_column_chunk(
                        doc,
                        src,
                        sc,
                        sr0,
                        sr1,
                        dest,
                        tc,
                        tr0,
                        with_formatting=bool(with_formatting),
                    )
                ci = ci + 1
                _lm_ui_yield(force=True)
        out_data_row = out_data_row + sheet_rows
        total_rows = total_rows + sheet_rows
        _lm_pp_merge_sheets_dbg(
            "source",
            "«%s» скопировано строк=%d → dest row до %d (буфер по заголовкам)"
            % (src_name, sheet_rows, out_data_row - 1),
        )
        si = si + 1
    if (not with_formatting) and auto_format:
        _lm_pp_merge_apply_default_data_format(dest, hdr)
    note = "объединено %d лист(ов), %d строк → «%s»%s" % (
        len(sources),
        total_rows,
        actual_dest,
        " +формат" if with_formatting else "",
    )
    _lm_pp_merge_sheets_dbg("done", note)
    return (
        True,
        note,
        actual_dest,
    )


def _lm_pp_sheet_is_visible(sh):
    """Лист видим в UI (не скрыт)."""
    if sh is None:
        return True
    try:
        return bool(sh.getPropertyValue("IsVisible"))
    except Exception:
        pass
    try:
        from com.sun.star.sheet.SheetVisibility import VISIBLE

        state = sh.getPropertyValue("SheetState")
        return int(state) == int(VISIBLE)
    except Exception:
        pass
    try:
        return bool(sh.IsVisible)
    except Exception:
        return True


def _lm_pp_sheet_set_visible(sh, visible=True):
    """Показать или скрыть лист (IsVisible / SheetState)."""
    if sh is None:
        return False
    want = bool(visible)
    try:
        sh.setPropertyValue("IsVisible", want)
        return True
    except Exception:
        pass
    try:
        from com.sun.star.sheet.SheetVisibility import HIDDEN, VISIBLE

        sh.setPropertyValue("SheetState", VISIBLE if want else HIDDEN)
        return True
    except Exception:
        pass
    try:
        sh.IsVisible = want
        return True
    except Exception:
        return False


def _lm_pp_activate_sheet_by_name(doc, sheet_name):
    """Сделать лист активным в UI (после копирования/перемещения)."""
    if doc is None:
        return False
    nm = str(sheet_name or "").strip()
    if nm == "":
        return False
    sh = _lm_pp_get_sheet_ref(doc, nm)
    if sh is None:
        return False
    try:
        controller = doc.getCurrentController()
        if controller is None:
            return False
        active = controller.getActiveSheet()
        if active is not None and str(active.Name) == str(sh.Name):
            return True
        controller.setActiveSheet(sh)
        return True
    except Exception:
        try:
            doc.getCurrentController().setActiveSheet(sh)
            return True
        except Exception:
            return False


def _lm_pp_sheet_tab_color_token_is_clear(token):
    """True — сбросить цвет ярлычка к умолчанию ОС/темы."""
    t = str(token or "").strip().casefold().replace("ё", "е")
    if t == "":
        return False
    return t in (
        "нет",
        "no",
        "none",
        "default",
        "auto",
        "авто",
        "сброс",
        "reset",
        "clear",
        "умолчание",
        "по_умолчанию",
        "по умолчанию",
    )


def _lm_pp_set_sheet_tab_color(sheet, color_token):
    """
    Задать цвет ярлычка листа (UNO TabColor).

    color_token: имя из zebra-алиасов / #RRGGBB;
    «нет»/default/сброс — TabColor=-1 (системный);
    пусто — не менять.

    Возвращает (ok, note_or_empty).
    """
    if sheet is None:
        return False, "нет листа"
    raw = str(color_token or "").strip()
    if raw == "":
        return True, ""
    if _lm_pp_sheet_tab_color_token_is_clear(raw):
        color_int = -1
        note = "ярлык: сброс"
    else:
        color_int = _lm_pp_zebra_resolve_color_token(raw)
        if color_int is None:
            return False, "ярлык: неизвестный цвет «%s»" % raw
        try:
            color_int = int(color_int) & 0xFFFFFF
        except Exception:
            return False, "ярлык: цвет «%s»" % raw
        note = "ярлык: #%06X" % color_int
    try:
        sheet.TabColor = color_int
        return True, note
    except Exception:
        pass
    try:
        sheet.setPropertyValue("TabColor", color_int)
        return True, note
    except Exception as err:
        return False, "ярлык: %s" % err


def _lm_pp_activate_sheet_focus_a1(doc, sheet_name, tab_color=None):
    """
    Активировать лист; если скрыт — показать; выделить ячейку A1.
    tab_color — опционально цвет ярлычка (имя/#hex/сброс).
    Возвращает (ok, note).
    """
    if doc is None:
        return False, "книга не задана"
    nm = str(sheet_name or "").strip()
    if nm == "":
        return False, "пустое имя листа"
    sh = _lm_pp_get_sheet_ref(doc, nm)
    if sh is None:
        return False, "лист «%s» не найден" % nm
    actual = str(getattr(sh, "Name", "") or "").strip()
    shown = False
    if not _lm_pp_sheet_is_visible(sh):
        if not _lm_pp_sheet_set_visible(sh, True):
            return False, "не удалось показать «%s»" % actual
        shown = True
    tab_note = ""
    raw_tab = str(tab_color or "").strip()
    if raw_tab != "":
        ok_tab, tab_note = _lm_pp_set_sheet_tab_color(sh, raw_tab)
        if not ok_tab:
            return False, tab_note or ("ярлык «%s»" % actual)
    try:
        controller = doc.getCurrentController()
        if controller is None:
            return False, "нет контроллера"
        try:
            active = controller.getActiveSheet()
            if active is None or str(active.Name) != str(sh.Name):
                controller.setActiveSheet(sh)
        except Exception:
            controller.setActiveSheet(sh)
        cell = sh.getCellByPosition(0, 0)
        controller.select(cell)
        note = "«%s»" % actual
        if shown:
            note = note + ", показан"
        if tab_note:
            note = note + ", " + tab_note
        note = note + ", A1"
        return True, note
    except Exception as err:
        return False, str(err)


def _lm_pp_copy_sheet_named(doc, source_sheet, dest_sheet):
    """Скопировать вкладку в той же книге (Sheets.copyByName). Возвращает (ok, note, dest_name)."""
    src_name = str(source_sheet or "").strip()
    dst_name = str(dest_sheet or "").strip()
    if doc is None:
        return False, "книга не задана", ""
    if src_name == "" or dst_name == "":
        return False, "пустое имя источника или приёмника", ""
    src_sh = _lm_pp_get_sheet_ref(doc, src_name)
    if src_sh is None:
        return False, "лист-источник «%s» не найден" % src_name, ""
    actual_src = str(src_sh.Name)
    target = _lm_pp_unique_sheet_name(doc, dst_name, exclude_sheet=None)
    if target == "":
        return False, "пустое имя копии", ""
    try:
        if actual_src.casefold() == target.casefold():
            return True, "«%s» уже существует" % target, target
    except Exception:
        pass
    try:
        pos = int(doc.Sheets.getCount())
        doc.Sheets.copyByName(actual_src, target, pos)
        return True, "«%s» → «%s»" % (actual_src, target), target
    except Exception as err:
        return False, str(err), ""


def _lm_pp_sheet_name_index(doc, sheet_name):
    """0-based индекс листа по имени (без учёта регистра), иначе -1."""
    if doc is None:
        return -1
    nm = str(sheet_name or "").strip()
    if nm == "":
        return -1
    ref = _lm_pp_get_sheet_ref(doc, nm)
    if ref is None:
        return -1
    actual = str(getattr(ref, "Name", "") or "").strip()
    try:
        i = 0
        while i < int(doc.Sheets.getCount()):
            try:
                sh = doc.Sheets.getByIndex(i)
                if str(getattr(sh, "Name", "") or "").strip().casefold() == actual.casefold():
                    return i
            except Exception:
                pass
            i = i + 1
    except Exception:
        return -1
    return -1


def _lm_pp_reposition_sheet(doc, moving_sheet_name, position=None, anchor_sheet_name=None):
    """
    Переставить лист относительно опорного.

    position: before|after|None. Возвращает (ok, note).
    """
    if doc is None:
        return False, "книга не задана"
    moving_ref = _lm_pp_get_sheet_ref(doc, moving_sheet_name)
    if moving_ref is None:
        return False, "лист «%s» не найден" % (moving_sheet_name or "")
    moving_name = str(getattr(moving_ref, "Name", "") or "").strip()
    if moving_name == "":
        return False, "пустое имя листа"
    pos_key = str(position or "").strip().casefold()
    if pos_key not in ("before", "after", "до", "перед", "_перед_", "после", "_после_"):
        return True, "позиция не задана (оставлен как есть)"
    anchor_ref = _lm_pp_get_sheet_ref(doc, anchor_sheet_name)
    if anchor_ref is None:
        return False, "опорный лист «%s» не найден" % (anchor_sheet_name or "")
    anchor_name = str(getattr(anchor_ref, "Name", "") or "").strip()
    if anchor_name == "":
        return False, "пустое имя опорного листа"
    if moving_name.casefold() == anchor_name.casefold():
        return True, "опорный лист совпадает с перемещаемым"
    moving_idx = _lm_pp_sheet_name_index(doc, moving_name)
    anchor_idx = _lm_pp_sheet_name_index(doc, anchor_name)
    if moving_idx < 0 or anchor_idx < 0:
        return False, "не удалось вычислить позицию листа"
    before = pos_key in ("before", "до", "перед", "_перед_")
    target_idx = anchor_idx if before else (anchor_idx + 1)
    if moving_idx < target_idx:
        target_idx = target_idx - 1
    if target_idx < 0:
        target_idx = 0
    try:
        max_idx = int(doc.Sheets.getCount()) - 1
    except Exception:
        max_idx = target_idx
    if target_idx > max_idx:
        target_idx = max_idx
    if moving_idx == target_idx:
        return True, "лист уже на нужной позиции"
    try:
        doc.Sheets.moveByName(moving_name, int(target_idx))
        rel = "перед" if before else "после"
        return True, "«%s» %s «%s»" % (moving_name, rel, anchor_name)
    except Exception as err:
        return False, str(err)


def lm_pp_range_value_count(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Столбец с числом строк в диапазоне с тем же составным ключом.

    JSON колонки C:
    [{"v":1,"fn":"количество_значений","new_column":"Количество","key_columns":["ФИО","Отдел"],
      "trim":true,"case_sensitive":false}]
    """
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("количество_значений", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet("количество_значений", extra_args, doc, sheet)
    if block is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "количество_значений",
            "пропуск",
            "нет блока JSON для листа",
        )
        return
    key_columns = block.get("key_columns") or block.get("columns") or []
    trim = True
    if "trim" in block:
        trim = bool(block.get("trim"))
    case_sensitive = bool(block.get("case_sensitive"))

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    new_col = _lm_pp_value_count_header_name(block, sheet, h_sr)
    if sr <= h_sr:
        sr = h_sr + 1
    if er < sr:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", "количество_значений", "пропуск", "нет строк данных"
        )
        return

    if isinstance(key_columns, (list, tuple)):
        spec_text = ",".join(str(c) for c in key_columns if str(c).strip() != "")
    else:
        spec_text = str(key_columns or "").strip()
    key_cols = _lm_pp_resolve_column_list(spec_text, sheet, header_row_range, sc, ec)
    if len(key_cols) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "количество_значений",
            "ошибка",
            "не заданы столбцы ключа (key_columns)",
        )
        return

    counts = {}
    row = sr
    while row <= er:
        key = _lm_pp_value_count_row_key(sheet, row, key_cols, trim, case_sensitive)
        counts[key] = counts.get(key, 0) + 1
        row = row + 1

    loc = _lm_pp_resolve_derived_column(sheet, data_range, header_row_range, new_col)
    if loc is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "количество_значений",
            "ошибка",
            "не удалось добавить столбец «%s»" % new_col,
        )
        return
    target_col = loc["target_col"]
    ref_col = loc["ref_col"]
    h_sr = loc["h_sr"]
    sr = loc["sr"]
    er = loc["er"]
    _lm_pp_clone_concat_column_formats(
        sheet, ref_col, target_col, h_sr, sr, er
    )
    sheet.getCellByPosition(target_col, h_sr).String = new_col
    row = sr
    while row <= er:
        key = _lm_pp_value_count_row_key(sheet, row, key_cols, trim, case_sensitive)
        cnt = counts.get(key, 0)
        cell = sheet.getCellByPosition(target_col, row)
        try:
            cell.Value = int(cnt)
        except Exception:
            cell.Value = cnt
        row = row + 1

    _lm_pp_finalize_derived_column(
        sheet, ref_col, target_col, h_sr, sr
    )
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        "количество_значений",
        "ok",
        "столбец «%s», ключ %d пол(ей), строк %d"
        % (new_col, len(key_cols), er - sr + 1),
    )


_LM_PP_REMOVE_DUPLICATES_MAX_ROWS = 500000
_LM_PP_REMOVE_DUPLICATES_FN = "удалить_дубликаты"


def _lm_pp_parse_cell_a1_ref(cell_ref):
    """A1-адрес → (col0, row0) или None."""
    s = str(cell_ref or "").strip()
    if s == "":
        return None
    m = re.match(r"^\$?([A-Za-z]+)\$?(\d+)$", s)
    if m is None:
        return None
    col_s = str(m.group(1) or "").strip()
    row_s = str(m.group(2) or "").strip()
    try:
        src_col = col_letters_to_index(col_s)
        row_0 = int(row_s) - 1
    except Exception:
        return None
    if src_col < 0 or row_0 < 0:
        return None
    return (src_col, row_0)


def _lm_pp_remove_duplicates_resolve_key_cols(block, sheet, header_row_range, sc, ec):
    """Столбцы ключа (0-based). Ошибка — (None, сообщение)."""
    key_columns = block.get("key_columns") or block.get("columns") or []
    exclude_columns = block.get("exclude_columns") or []
    has_key = False
    if isinstance(key_columns, (list, tuple)):
        has_key = any(str(c).strip() != "" for c in key_columns)
    else:
        has_key = str(key_columns or "").strip() != ""
    has_excl = False
    if isinstance(exclude_columns, (list, tuple)):
        has_excl = any(str(c).strip() != "" for c in exclude_columns)
    else:
        has_excl = str(exclude_columns or "").strip() != ""
    if has_key and has_excl:
        return None, "key_columns и exclude_columns заданы одновременно"
    if has_excl:
        if isinstance(exclude_columns, (list, tuple)):
            spec_text = ",".join(str(c) for c in exclude_columns if str(c).strip() != "")
        else:
            spec_text = str(exclude_columns or "").strip()
        excl_cols = _lm_pp_resolve_column_list(
            spec_text, sheet, header_row_range, sc, ec
        )
        if len(excl_cols) == 0 and spec_text != "":
            return None, "exclude_columns: столбцы не найдены"
        excl_set = set(excl_cols)
        key_cols = []
        c = sc
        while c <= ec:
            if c not in excl_set:
                key_cols.append(c)
            c = c + 1
        if len(key_cols) == 0:
            return None, "нет столбцов ключа после exclude_columns"
        return key_cols, None
    if has_key:
        if isinstance(key_columns, (list, tuple)):
            spec_text = ",".join(str(c) for c in key_columns if str(c).strip() != "")
        else:
            spec_text = str(key_columns or "").strip()
        key_cols = _lm_pp_resolve_column_list(
            spec_text, sheet, header_row_range, sc, ec
        )
        if len(key_cols) == 0:
            return None, "key_columns: столбцы не найдены"
        return key_cols, None
    key_cols = []
    c = sc
    while c <= ec:
        key_cols.append(c)
        c = c + 1
    return key_cols, None


def _lm_pp_remove_duplicates_part_empty(val, trim=True):
    if val is None:
        return True
    if isinstance(val, (int, float)):
        return False
    s = str(val)
    if trim:
        s = s.strip()
    return s == ""


def _lm_pp_remove_duplicates_strict_part(val, trim=True, case_sensitive=False):
    if _lm_pp_remove_duplicates_part_empty(val, trim=trim):
        return ("e", "")
    if isinstance(val, bool):
        return ("b", bool(val))
    if isinstance(val, (int, float)):
        try:
            return ("n", float(val))
        except (TypeError, ValueError):
            return ("n", 0.0)
    s = str(val)
    if trim:
        s = s.strip()
    if not case_sensitive:
        s = s.casefold()
    return ("t", s)


def _lm_pp_remove_duplicates_array_row_key( row, rel_cols, trim, case_sensitive, compare_as, row_idx, empty_key_policy):
    parts = []
    i = 0
    while i < len(rel_cols):
        rel = rel_cols[i]
        if rel >= 0 and rel < len(row):
            parts.append(row[rel])
        else:
            parts.append(None)
        i = i + 1
    all_empty = True
    j = 0
    while j < len(parts):
        if not _lm_pp_remove_duplicates_part_empty(parts[j], trim=trim):
            all_empty = False
            break
        j = j + 1
    if all_empty:
        pol = str(empty_key_policy or "dedup").strip().casefold()
        if pol == "unique":
            return ("__empty_unique__", int(row_idx))
        if pol == "skip":
            return ("__empty_skip__", int(row_idx))
    if str(compare_as or "text").strip().casefold() == "strict":
        out = []
        k = 0
        while k < len(parts):
            out.append(
                _lm_pp_remove_duplicates_strict_part(
                    parts[k], trim=trim, case_sensitive=case_sensitive
                )
            )
            k = k + 1
        return tuple(out)
    str_parts = []
    k = 0
    while k < len(parts):
        s = "" if parts[k] is None else str(parts[k])
        if trim:
            s = s.strip()
        if not case_sensitive:
            s = s.casefold()
        str_parts.append(s)
        k = k + 1
    return _lm_pp_value_count_key_from_parts(str_parts, trim=False, case_sensitive=True)


def _lm_pp_remove_duplicates_resolve_pre_sort(pre_sort, sheet, header_row_range, sc, ec):
    """pre_sort → [(col_index, ascending), …] или None."""
    resolved = []
    i = 0
    while i < len(pre_sort or []):
        item = pre_sort[i]
        if not isinstance(item, dict):
            i = i + 1
            continue
        col_token = str(item.get("column") or "").strip()
        if col_token == "":
            i = i + 1
            continue
        col_idx = _lm_pp_resolve_column_token(
            col_token, sheet, header_row_range, sc, ec
        )
        if col_idx is None:
            return None
        order = str(item.get("order") or "asc").strip().casefold()
        ascending = order not in ("desc", "-", "убыв")
        resolved.append((col_idx, ascending))
        i = i + 1
    return resolved


def _lm_pp_remove_duplicates_ranges_overlap(sc1, sr1, ec1, er1, sc2, sr2, ec2, er2):
    return not (ec1 < sc2 or ec2 < sc1 or er1 < sr2 or er2 < sr1)


def _lm_pp_remove_duplicates_pick_deletes( body, rel_cols, keep, scan_order, trim, case_sensitive, compare_as, empty_key_policy):
    n = len(body)
    if n == 0:
        return [], 0
    keep_v = str(keep or "first").strip().casefold()
    if keep_v not in ("first", "last"):
        keep_v = "first"
    scan = str(scan_order or "").strip().casefold()
    if scan not in ("top_down", "bottom_up"):
        scan = "top_down" if keep_v == "first" else "bottom_up"
    order = list(range(n))
    if scan == "bottom_up":
        order.reverse()
    seen = set()
    to_delete = []
    unique_count = 0
    i = 0
    while i < len(order):
        idx = order[i]
        key = _lm_pp_remove_duplicates_array_row_key(
            body[idx],
            rel_cols,
            trim,
            case_sensitive,
            compare_as,
            idx,
            empty_key_policy,
        )
        pol = str(empty_key_policy or "dedup").strip().casefold()
        if pol == "skip" and isinstance(key, tuple) and len(key) == 2 and key[0] == "__empty_skip__":
            i = i + 1
            continue
        if key in seen:
            to_delete.append(idx)
        else:
            seen.add(key)
            unique_count = unique_count + 1
        i = i + 1
    return to_delete, unique_count


def _lm_pp_remove_duplicates_delete_rows_batch(sheet, row_indices):
    if sheet is None or not row_indices:
        return 0
    rows = sorted(set(int(r) for r in row_indices))
    groups = []
    i = len(rows) - 1
    while i >= 0:
        end = rows[i]
        start = end
        while i > 0 and rows[i - 1] == start - 1:
            i = i - 1
            start = rows[i]
        groups.append((start, end - start + 1))
        i = i - 1
    deleted = 0
    gi = 0
    while gi < len(groups):
        start, count = groups[gi]
        _lm_sheet_delete_rows(sheet, start, count)
        deleted = deleted + count
        gi = gi + 1
        if deleted % 200 == 0:
            _lm_ui_yield(force=True)
    return deleted


def _lm_pp_remove_duplicates_apply_dest_formats( doc, src_sheet, dest_sheet, sc, ec, h_sr, sr, n_data_rows ):
    """
    Оформление листа результата dedup (output=new_sheet):
    заголовок — как на источнике; строки данных — формат первой строки данных,
    одной спецвставкой форматов (InsertContents T) на весь блок.
    """
    if src_sheet is None or dest_sheet is None:
        return False
    try:
        sc = int(sc)
        ec = int(ec)
        h_sr = int(h_sr)
        sr = int(sr)
        n_data_rows = int(n_data_rows)
    except (TypeError, ValueError):
        return False
    if ec < sc:
        return False
    dest_ec = ec - sc
    ok = False
    try:
        src_hdr = src_sheet.getCellRangeByPosition(sc, h_sr, ec, h_sr)
        dst_hdr = dest_sheet.getCellRangeByPosition(0, 0, dest_ec, 0)
        _lm_pp_copy_cell_range_format(src_hdr, dst_hdr)
        ok = True
    except Exception:
        pass
    if n_data_rows <= 0:
        return ok
    try:
        from libre_macros_split_lib import (
            _split_clipboard_copy_range,
            _split_clipboard_paste_formats_only,
        )
        if _split_clipboard_copy_range(doc, src_sheet, sc, sr, ec, sr):
            if _split_clipboard_paste_formats_only(
                doc, dest_sheet, 0, 1, dest_ec, n_data_rows
            ):
                return True
    except Exception:
        pass
    try:
        from libre_macros_split_lib import _split_apply_formats_from_first_data_row
        if _split_apply_formats_from_first_data_row(
            doc, src_sheet, sr, sc, ec, doc, dest_sheet, 1, n_data_rows
        ):
            ok = True
    except Exception:
        pass
    return ok


def _lm_pp_remove_duplicates_expand_offset_sheet_range( sheet, h_sr, h_ec, data_ec, dest_col, dest_row, out_ec, header_row_values=None ):
    """
    После output=offset на том же листе: расширить end_col постобработки.

    Для новых столбцов справа от текущего диапазона — проверить строку заголовков;
    пустые ячейки дописать синтетическими «Колонка_*» (как при сборе без заголовка).
    Промежуток между концом data_range и dest_cell (например F–G при выводе в H1) тоже
    заполняется в строке заголовков листа (h_sr).
    """
    if sheet is None:
        return
    try:
        h_sr = int(h_sr)
        h_ec = int(h_ec)
        data_ec = int(data_ec)
        dest_col = int(dest_col)
        dest_row = int(dest_row)
        out_ec = int(out_ec)
    except (TypeError, ValueError):
        return
    if out_ec < dest_col:
        return
    prev_end = max(h_ec, data_ec)
    if out_ec <= prev_end:
        ctx = lm_pp_active_context()
        if ctx is not None:
            _lm_pp_context_expand_end_col(ctx, prev_end)
        return
    col = prev_end + 1
    while col <= out_ec:
        hdr_row = dest_row if col >= dest_col else h_sr
        hdr_cell = sheet.getCellByPosition(col, hdr_row)
        existing = cell_text(hdr_cell).strip()
        if existing == "":
            name = ""
            if header_row_values is not None and col >= dest_col:
                idx = col - dest_col
                if idx >= 0 and idx < len(header_row_values):
                    name = str(header_row_values[idx] or "").strip()
            if name == "":
                name = "Колонка_%s" % col_index_to_letters(col)
            try:
                hdr_cell.String = name
            except Exception:
                pass
        col = col + 1
    ctx = lm_pp_active_context()
    if ctx is not None:
        _lm_pp_context_refresh(ctx)
        _lm_pp_context_expand_end_col(ctx, out_ec)


def _lm_pp_remove_duplicates_write_matrix(doc, dest_sheet, dest_col, dest_row, matrix):
    if dest_sheet is None or not matrix:
        return False, "пустой блок"
    nrows = len(matrix)
    ncols = len(matrix[0]) if nrows > 0 else 0
    if ncols <= 0:
        return False, "нет столбцов"
    er = dest_row + nrows - 1
    ec = dest_col + ncols - 1
    try:
        rng = dest_sheet.getCellRangeByPosition(dest_col, dest_row, ec, er)
        rng.setDataArray(tuple(tuple(row) for row in matrix))
        return True, ""
    except Exception as err:
        return False, str(err)


def lm_pp_range_remove_duplicates(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Удалить повторяющиеся строки по ключу (или пометить / записать на другой лист).

    JSON колонки C:
    [{"v":1,"fn":"удалить_дубликаты","key_columns":["ФИО"],"keep":"first"}]
    """
    fn_label = _LM_PP_REMOVE_DUPLICATES_FN
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet(fn_label, extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(fn_label, extra_args, doc, sheet)
    if block is None:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_label, "пропуск", "нет блока JSON для листа"
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    if sr <= h_sr:
        sr = h_sr + 1
    if er < sr:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_label, "пропуск", "нет строк данных"
        )
        return

    row_count = int(er) - int(sr) + 1
    if row_count > _LM_PP_REMOVE_DUPLICATES_MAX_ROWS:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_label,
            "ошибка",
            "строк %d > лимит %d"
            % (row_count, _LM_PP_REMOVE_DUPLICATES_MAX_ROWS),
        )
        return

    key_cols, key_err = _lm_pp_remove_duplicates_resolve_key_cols(
        block, sheet, header_row_range, sc, ec
    )
    if key_err:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_label, "ошибка", key_err
        )
        return

    trim = True
    if "trim" in block:
        trim = bool(block.get("trim"))
    case_sensitive = bool(block.get("case_sensitive"))
    compare_as = str(block.get("compare_as") or "text").strip().casefold()
    if compare_as not in ("text", "strict"):
        compare_as = "text"
    empty_key_policy = str(block.get("empty_key_policy") or "dedup").strip().casefold()
    if empty_key_policy not in ("dedup", "unique", "skip"):
        empty_key_policy = "dedup"
    keep = str(block.get("keep") or "first").strip().casefold()
    scan_order = str(block.get("scan_order") or "").strip().casefold()
    output = str(block.get("output") or "inplace").strip().casefold()
    if output not in ("inplace", "new_sheet", "offset"):
        output = "inplace"
    mark_duplicates = bool(block.get("mark_duplicates"))
    as_values = True
    if "as_values" in block:
        as_values = bool(block.get("as_values"))
    pre_sort = block.get("pre_sort") or []

    if output == "new_sheet":
        dest_sheet_name = str(block.get("dest_sheet") or block.get("dest") or "").strip()
        if dest_sheet_name == "":
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_label,
                "ошибка",
                "output=new_sheet: нужен dest_sheet",
            )
            return

    dest_cell = str(block.get("dest_cell") or "A1").strip()
    if output == "offset" and dest_cell == "":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_label,
            "ошибка",
            "output=offset: нужен dest_cell",
        )
        return

    if as_values and doc is not None and sheet is not None:
        try:
            mat_rng = sheet.getCellRangeByPosition(int(sc), int(h_sr), int(ec), int(er))
            _lm_pp_calculate_with_auto(doc, sheet)
            _lm_pp_copy_range_as_values(doc, sheet, mat_rng)
        except Exception:
            pass

    if pre_sort and output == "inplace":
        resolved_sort = _lm_pp_remove_duplicates_resolve_pre_sort(
            pre_sort, sheet, header_row_range, sc, ec
        )
        if resolved_sort is None:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_label,
                "ошибка",
                "pre_sort: столбец не найден",
            )
            return
        if len(resolved_sort) > 0:
            _lm_pp_range_sort_dispatch(
                doc,
                sheet,
                sc,
                sr,
                ec,
                er,
                resolved_sort,
                sc,
                sheet_name=sheet_name,
                fn_label=fn_label,
            )

    try:
        full_rng = sheet.getCellRangeByPosition(int(sc), int(h_sr), int(ec), int(er))
        data = full_rng.getDataArray()
    except Exception as err:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_label,
            "ошибка",
            "getDataArray: %s" % err,
        )
        return

    header = list(data[0]) if len(data) > 0 else []
    body = []
    ri = 1
    while ri < len(data):
        body.append(list(data[ri]))
        ri = ri + 1

    if pre_sort and output in ("new_sheet", "offset"):
        resolved_sort = _lm_pp_remove_duplicates_resolve_pre_sort(
            pre_sort, sheet, header_row_range, sc, ec
        )
        if resolved_sort is None:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_label,
                "ошибка",
                "pre_sort: столбец не найден",
            )
            return
        if len(resolved_sort) > 0:
            body = _lm_pp_sort_row_list(body, resolved_sort, sc)

    rel_cols = []
    ki = 0
    while ki < len(key_cols):
        rel_cols.append(int(key_cols[ki]) - int(sc))
        ki = ki + 1

    del_body_idx, unique_keys = _lm_pp_remove_duplicates_pick_deletes(
        body,
        rel_cols,
        keep,
        scan_order,
        trim,
        case_sensitive,
        compare_as,
        empty_key_policy,
    )
    delete_set = set(del_body_idx)
    deleted_count = len(del_body_idx)

    if mark_duplicates:
        mark_name = str(block.get("mark_column") or "Дубликаты").strip()
        if mark_name == "":
            mark_name = "Дубликаты"
        loc = _lm_pp_resolve_derived_column(
            sheet, data_range, header_row_range, mark_name
        )
        if loc is None:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_label,
                "ошибка",
                "не удалось добавить столбец «%s»" % mark_name,
            )
            return
        mark_col = loc["target_col"]
        ref_col = loc["ref_col"]
        h_sr = loc["h_sr"]
        _lm_pp_clone_concat_column_formats(
            sheet, ref_col, mark_col, h_sr, sr, er
        )
        sheet.getCellByPosition(mark_col, h_sr).String = mark_name
        bi = 0
        while bi < len(body):
            if bi in delete_set:
                try:
                    sheet.getCellByPosition(mark_col, sr + bi).String = "Да"
                except Exception:
                    pass
            bi = bi + 1
        _lm_pp_finalize_derived_column(
            sheet, ref_col, mark_col, h_sr, sr, doc=doc
        )
        note = (
            "маркировка «%s»: строк %d, помечено %d, уникальных ключей %d"
            % (mark_name, row_count, deleted_count, unique_keys)
        )
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_label, "ok", note
        )
        return

    if output == "inplace":
        if deleted_count == 0:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_label,
                "ok",
                "строк %d, удалено 0, дубликатов не найдено, уникальных ключей %d"
                % (row_count, unique_keys),
            )
            return
        sheet_rows = [int(sr) + int(i) for i in del_body_idx]
        removed = _lm_pp_remove_duplicates_delete_rows_batch(sheet, sheet_rows)
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_label,
            "ok",
            "строк было %d, удалено %d, осталось %d, уникальных ключей %d, output=inplace"
            % (row_count, removed, row_count - removed, unique_keys),
        )
        return

    survivor_body = []
    si = 0
    while si < len(body):
        if si not in delete_set:
            survivor_body.append(body[si])
        si = si + 1
    out_matrix = [header] + survivor_body

    if output == "new_sheet":
        dest_base = str(block.get("dest_sheet") or block.get("dest") or "").strip()
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
                    "диапазон",
                    fn_label,
                    "ошибка",
                    "создание листа «%s»: %s" % (actual_dest, err),
                )
                return
            dest_sh = _lm_pp_get_sheet_ref(doc, actual_dest)
        else:
            _lm_pp_clear_sheet_content(dest_sh)
        ok, werr = _lm_pp_remove_duplicates_write_matrix(
            doc, dest_sh, 0, 0, out_matrix
        )
        if not ok:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_label,
                "ошибка",
                werr or "запись на лист",
            )
            return
        _lm_pp_remove_duplicates_apply_dest_formats(
            doc, sheet, dest_sh, sc, ec, h_sr, sr, len(survivor_body)
        )
        lm_pp_copy_sheet_register_pending(
            actual_dest, 0, sheet_name, from_dedup=True
        )
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_label,
            "ok",
            "строк было %d, записано %d на «%s», удалено %d, уникальных ключей %d"
            % (
                row_count,
                len(survivor_body),
                actual_dest,
                deleted_count,
                unique_keys,
            ),
        )
        return

    # offset
    anchor = _lm_pp_parse_cell_a1_ref(dest_cell)
    if anchor is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_label,
            "ошибка",
            "неверный dest_cell «%s»" % dest_cell,
        )
        return
    dest_col, dest_row = anchor
    out_rows = len(out_matrix)
    out_cols = len(header)
    out_er = dest_row + out_rows - 1
    out_ec = dest_col + out_cols - 1
    overwrite = bool(block.get("overwrite_overlap"))
    if not overwrite and _lm_pp_remove_duplicates_ranges_overlap(
        sc, sr, ec, er, dest_col, dest_row, out_ec, out_er
    ):
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_label,
            "ошибка",
            "вывод пересекается с исходным диапазоном (overwrite_overlap=false)",
        )
        return
    ok, werr = _lm_pp_remove_duplicates_write_matrix(
        doc, sheet, dest_col, dest_row, out_matrix
    )
    if not ok:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_label,
            "ошибка",
            werr or "запись offset",
        )
        return
    _lm_pp_remove_duplicates_expand_offset_sheet_range(
        sheet,
        h_sr,
        h_ec,
        ec,
        dest_col,
        dest_row,
        out_ec,
        header_row_values=header,
    )
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        fn_label,
        "ok",
        "строк было %d, записано %d с %s, удалено из выборки %d, уникальных ключей %d, output=offset"
        % (
            row_count,
            len(survivor_body),
            dest_cell,
            deleted_count,
            unique_keys,
        ),
    )


def lm_pp_range_copy_sheet(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Копировать/переместить лист в этой же книге.

    JSON колонки C:
    [{"v":1,"fn":"копировать_переместить_лист","source_sheet":"Шаблон","copy":true,"dest_sheet":"Копия_Шаблон"}]
    Поле sheet — на каком шаге постобработки выполнить копирование (пусто — на всех).
    copy=false (по умолчанию): перемещение source_sheet относительно anchor_sheet (before/after);
    при перемещении прочие поля копирования игнорируются.
    """
    _unused = (data_range, header_row_range)
    sheet_name = sheet.Name if sheet is not None else ""
    fn_key = "копировать_переместить_лист"
    if _lm_pp_param_skip_sheet(fn_key, extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(fn_key, extra_args, doc, sheet)
    if block is None:
        # Обратная совместимость со старым именем.
        block = _lm_pp_param_block_for_sheet("копировать_лист", extra_args, doc, sheet)
    if block is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_key,
            "пропуск",
            "нет блока JSON для листа",
        )
        return
    source = str(
        block.get("source_sheet") or block.get("source") or ""
    ).strip()
    do_copy = False
    try:
        do_copy = bool(lm_parse_bool_param(block.get("copy"), default=False))
    except Exception:
        do_copy = bool(block.get("copy"))
    position = str(
        block.get("position") or block.get("place") or block.get("where") or ""
    ).strip()
    anchor_sheet = str(
        block.get("anchor_sheet")
        or block.get("relative_sheet")
        or block.get("ref_sheet")
        or block.get("anchor")
        or ""
    ).strip()
    if source == "":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_key,
            "ошибка",
            "нужен source_sheet",
        )
        return
    if not do_copy:
        ok, note = _lm_pp_reposition_sheet(
            doc, source, position=position, anchor_sheet_name=anchor_sheet
        )
        if ok:
            _lm_pp_activate_sheet_by_name(doc, source)
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_key,
            "ok" if ok else "ошибка",
            note,
        )
        return
    dest = str(
        block.get("dest_sheet") or block.get("dest") or block.get("target_sheet") or ""
    ).strip()
    if dest == "":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_key,
            "ошибка",
            "для copy=true нужен dest_sheet",
        )
        return
    ok, note, dest_name = _lm_pp_copy_sheet_named(doc, source, dest)
    if ok and dest_name != "" and position != "" and anchor_sheet != "":
        ok2, note2 = _lm_pp_reposition_sheet(
            doc, dest_name, position=position, anchor_sheet_name=anchor_sheet
        )
        if not ok2:
            ok = False
            note = "%s; позиция: %s" % (note, note2)
        else:
            note = "%s; %s" % (note, note2)
    if ok and dest_name != "":
        dest_sh = _lm_pp_get_sheet_ref(doc, dest_name)
        if dest_sh is not None:
            hdr = _lm_pp_copy_sheet_guess_header_row(dest_sh)
            lm_pp_copy_sheet_register_pending(dest_name, hdr, source)
        _lm_pp_activate_sheet_by_name(doc, dest_name)
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        fn_key,
        "ok" if ok else "ошибка",
        note,
    )


def lm_pp_range_merge_sheets_into_one(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Объединить несколько листов книги в один (как сбор «буфер по заголовкам»).

    JSON:
    [{"v":1,"fn":"объединить_листы_в_один","sheets":["Merge_1","Merge_2"],
      "dest_sheet":"Сводная","header_row":2,"sheet":"Merge_1","with_formatting":false}]

    - sheets — источники; sheet — фильтр контекста (когда выполнять шаг);
    - dest_sheet — приёмник (создаётся или очищается);
    - with_formatting — true: Transferable/Copy+Paste столбца чанками; false: как при сборе;
    - столбцы стыкуются по имени заголовка; заголовок всегда по умолчанию.
    header_row — номер строки заголовков в Calc (1-based); пусто — авто.
  """
    _unused = (data_range, header_row_range)
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("объединить_листы_в_один", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(
        "объединить_листы_в_один", extra_args, doc, sheet
    )
    if block is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "объединить_листы_в_один",
            "пропуск",
            "нет блока JSON для листа",
        )
        return
    sheets_spec = (
        block.get("sheets")
        or block.get("source_sheets")
        or block.get("sheet_list")
        or []
    )
    dest = str(
        block.get("dest_sheet")
        or block.get("result_sheet")
        or block.get("target_sheet")
        or block.get("dest")
        or ""
    ).strip()
    if dest == "":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "объединить_листы_в_один",
            "ошибка",
            "нужен dest_sheet",
        )
        return
    source_names = _lm_pp_resolve_merge_source_sheet_names(doc, sheets_spec)
    _lm_pp_merge_sheets_dbg(
        "resolve",
        "ctx=%s spec=%r → [%s]"
        % (sheet_name, sheets_spec, ",".join(source_names)),
    )
    if not source_names:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "объединить_листы_в_один",
            "ошибка",
            "нет листов по sheets",
        )
        return
    hdr_param = block.get("header_row")
    first_src = _lm_pp_get_sheet_ref(doc, source_names[0])
    hdr0 = _lm_pp_parse_merge_header_row(first_src, hdr_param)
    with_formatting = False
    try:
        with_formatting = bool(
            lm_parse_bool_param(block.get("with_formatting"), default=False)
        )
    except Exception:
        with_formatting = bool(block.get("with_formatting"))
    auto_format = True
    try:
        auto_format = bool(
            lm_parse_bool_param(block.get("auto_format"), default=True)
        )
    except Exception:
        if "auto_format" in block:
            auto_format = bool(block.get("auto_format"))
    columns = list(block.get("columns") or [])
    ok, note, dest_name = _lm_pp_merge_sheets_into_one_core(
        doc,
        source_names,
        dest,
        header_row=hdr_param,
        with_formatting=with_formatting,
        columns=columns,
        auto_format=auto_format,
    )
    if ok and dest_name != "":
        dest_sh = _lm_pp_get_sheet_ref(doc, dest_name)
        if dest_sh is not None:
            if auto_format:
                _lm_pp_apply_merge_dest_header_default(dest_sh, hdr0)
            lm_pp_copy_sheet_register_pending(dest_name, hdr0, sheet_name)
            _lm_pp_merge_sheets_dbg(
                "pipeline",
                "«%s»%s в очереди постобработки (alias=%s)"
                % (
                    dest_name,
                    " оформлен заголовок +" if auto_format else "",
                    sheet_name,
                ),
            )
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        "объединить_листы_в_один",
        "ok" if ok else "ошибка",
        note,
    )


def lm_pp_range_rename_sheet(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Переименовать текущий лист результата.

    Колонка C / new_name: новое имя; плейсхолдеры:
      <<Переменные.…>>, [Старое_имя], [Заголовок(N)], [Строка(R, C)],
      [Текущая_Дата(_Время)_FMT], [Текущее_Время_FMT], shorthand без FMT,
      [Имя_Книги]/[Имя_Файла], [Номер_Листа], [GUID8].
      Формат даты: YYYY/YY/MM/DD/HH/mm/SS (минуты — mm).
    После подстановки — нормализация (запрещённые символы, ≤31, иначе 28+_01).

    Карта: «переименовать_лист».
    """
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("переименовать_лист", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet("переименовать_лист", extra_args, doc, sheet)
    if block is not None:
        new_name = str(block.get("new_name") or "").strip()
    else:
        new_name = _lm_pp_join_extra_args(extra_args).strip()
    raw_name = new_name
    if new_name == "":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_лист",
            "—",
            "пропуск",
            "пустая колонка C",
        )
        return
    if sheet is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_лист",
            new_name,
            "ошибка",
            "лист не задан",
        )
        return

    ctx = lm_pp_active_context() or {}
    header_row = 0
    try:
        header_row = int(ctx.get("header_row") or 0)
    except (TypeError, ValueError):
        header_row = 0
    if header_row_range is not None:
        try:
            _hsc, hsr, _hec, _her = _lm_pp_range_address(header_row_range)
            header_row = int(hsr)
        except Exception:
            pass
    data_sr = header_row + 1
    data_er = data_sr
    if data_range is not None:
        try:
            _dsc, dsr, _dec, der = _lm_pp_range_address(data_range)
            data_sr = int(dsr)
            data_er = int(der)
        except Exception:
            pass
    else:
        _sc, used_sr, _ec, used_er = _lm_vlookup_sheet_used_area(sheet)
        data_sr = max(int(header_row) + 1, int(used_sr))
        data_er = max(int(data_sr), int(used_er))
    new_name = _lm_pp_expand_rename_name_template(
        new_name,
        doc=doc,
        sheet=sheet,
        old_name=sheet_name,
        header_row=header_row,
        used_sr=data_sr,
        used_er=data_er,
        log_fn_key="переименовать_лист",
    ).strip()

    exclude_name = str(sheet_name or "").strip()

    def _sheet_taken(cand):
        name = str(cand or "").strip()
        if name == "":
            return True
        if exclude_name != "" and name.casefold() == exclude_name.casefold():
            return False
        try:
            sheets = doc.getSheets()
            if not sheets.hasByName(name):
                return False
            other = sheets.getByName(name)
            if _lm_pp_sheet_same(other, sheet):
                return False
            return True
        except Exception:
            return True

    target = _lm_pp_finalize_sheet_name(new_name, is_taken=_sheet_taken)
    if target == "":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_лист",
            raw_name or "—",
            "пропуск",
            "пустое имя после подстановки/нормализации",
        )
        return
    if target == sheet_name:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_лист",
            raw_name,
            "ok",
            "имя без изменений",
        )
        return
    try:
        sheet.Name = target
        note = "«%s» → «%s»" % (sheet_name, target)
        if raw_name != target:
            note = note + " (шаблон «%s»)" % raw_name
        _lm_log_postprocess(
            doc,
            target,
            "переименовать_лист",
            raw_name,
            "ok",
            note,
        )
    except Exception as err:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_лист",
            raw_name,
            "ошибка",
            str(err),
        )




def lm_pp_range_reorder_columns(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Переставить или скопировать группу столбцов: columns + position
    (_начало_|_конец_|_перед_|_после_).

    JSON: [{"v":1,"fn":"переставить_столбцы","sheet":"Заказы","columns":["C","Сумма"],
    "position":"before","relative_column":"Итого"}].
    Копия: "copy":true, "new_names":["Сумма_копия"] — исходные столбцы остаются,
    копии вставляются в position; new_names — новые заголовки (пусто — авто _2).
    Новые пустые: "create_new":true, "new_names":["A","B"] — создать столбцы
    по именам и вставить по position (columns не нужны).
    Блок без sheet или sheet=Все — пропуск (только конкретный лист).
    Опорный столбец (для before/after) должен быть единственным; при переносе
    если он в columns — исключается из переноса (при копировании — нет).

    Карта: «переставить_столбцы».
    """
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("переставить_столбцы", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(
        "переставить_столбцы", extra_args, doc, sheet
    )
    if block is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переставить_столбцы",
            "—",
            "пропуск",
            "нет блока для листа",
        )
        return
    sheet_filter = str(block.get("sheet") or "").strip()
    if sheet_filter == "" or lm_identity_key(sheet_filter) == "все":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переставить_столбцы",
            "—",
            "пропуск",
            "нужен конкретный лист в sheet (не «Все»)",
        )
        return
    tokens = block.get("columns") or []
    do_create_new = _lm_pp_block_is_create_new_columns(block)
    do_copy = _lm_pp_block_is_copy_columns(block)
    position = str(block.get("position") or "").strip().casefold()
    try:
        from libre_macros_param_codec import _normalize_reorder_position

        position = _normalize_reorder_position(position)
    except Exception:
        pass
    if position == "":
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переставить_столбцы",
            "—",
            "ошибка",
            "не задана position",
        )
        return
    if do_create_new:
        new_names = _lm_pp_parse_copy_new_names(block)
        # Пустые токены тоже считаются (место под авто-имя).
        if not new_names:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переставить_столбцы",
                "—",
                "ошибка",
                "для «новые» нужны имена в new_names",
            )
            return
    elif not tokens:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переставить_столбцы",
            "—",
            "ошибка",
            "не заданы columns",
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    end_col, end_row = get_sheet_used_bounds(sheet)
    if end_col < 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переставить_столбцы",
            "—",
            "пропуск",
            "нет данных на листе",
        )
        return
    if end_row < 0:
        end_row = er
    else:
        end_row = max(end_row, er)

    hr = 0
    try:
        _h_sc, hr, _h_ec, _h_er = _lm_pp_range_address(header_row_range)
    except Exception:
        hr = 0

    relative_col = None
    rel_token = str(block.get("relative_column") or "").strip()
    if position in ("before", "after"):
        if rel_token == "":
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переставить_столбцы",
                position,
                "ошибка",
                "для before/after нужен relative_column",
            )
            return
        relative_col, rel_err = _lm_pp_resolve_column_token_unique(
            rel_token, sheet, header_row_range, 0, end_col
        )
        if rel_err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переставить_столбцы",
                rel_token,
                "ошибка",
                rel_err,
            )
            return

    if do_create_new:
        all_cols = list(range(end_col + 1))
        insert_at, plan_err = _lm_pp_resolve_insert_at(
            all_cols, position, relative_col
        )
        if plan_err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переставить_столбцы",
                position,
                "ошибка",
                plan_err,
            )
            return
        try:
            ok, insert_at = _lm_pp_insert_empty_columns_uno(
                sheet, insert_at, new_names, hr
            )
            note = "новые пустые ×%d → %s" % (len(new_names), position)
            if rel_token:
                note = note + ", опора «%s»" % rel_token
            if ok:
                ctx = lm_pp_active_context()
                if ctx is not None:
                    _lm_pp_context_refresh(ctx)
                    _lm_pp_context_expand_end_col(
                        ctx, end_col + len(new_names)
                    )
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переставить_столбцы",
                ",".join(new_names),
                "ok" if ok else "ошибка",
                note,
            )
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переставить_столбцы",
                ",".join(new_names),
                "ошибка",
                str(err),
            )
        return

    moving = _lm_pp_resolve_move_column_indices(
        tokens, sheet, header_row_range, 0, end_col
    )
    if LIBRE_MACROS_DEBUG:
        print(
            "  [переставить_столбцы] лист=%s tokens=%s → moving=%s position=%s copy=%s"
            % (sheet_name, tokens, moving, position, do_copy)
        )
    if not moving:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переставить_столбцы",
            ",".join(str(t) for t in tokens),
            "ошибка",
            "столбцы для переноса не найдены",
        )
        return

    if position in ("before", "after"):
        if (not do_copy) and relative_col in moving:
            moving = [c for c in moving if c != relative_col]
            if LIBRE_MACROS_DEBUG:
                print(
                    "  [переставить_столбцы] опорный %d исключён из переноса"
                    % relative_col
                )
        if not moving:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переставить_столбцы",
                rel_token,
                "пропуск",
                "после исключения опорного не осталось столбцов",
            )
            return

    src_titles = []
    if do_copy:
        mi = 0
        while mi < len(moving):
            try:
                src_titles.append(
                    cell_text(sheet.getCellByPosition(moving[mi], hr))
                )
            except Exception:
                src_titles.append("")
            mi = mi + 1

    all_cols = list(range(end_col + 1))
    insert_at = None
    if do_copy:
        _new_order, insert_at, plan_err = _lm_pp_build_column_copy(
            all_cols, moving, position, relative_col
        )
    else:
        new_order, plan_err = _lm_pp_build_column_reorder(
            all_cols, moving, position, relative_col
        )
    if plan_err:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переставить_столбцы",
            position,
            "ошибка",
            plan_err,
        )
        return
    if LIBRE_MACROS_DEBUG:
        print(
            "  [переставить_столбцы] insert_at=%s moving=%s copy=%s"
            % (insert_at, moving, do_copy)
        )
    try:
        if do_copy:
            ok, insert_at = _lm_pp_copy_columns_insert_uno(
                doc, sheet, moving, insert_at, end_row, hr
            )
            if ok and insert_at is not None:
                new_names = _lm_pp_parse_copy_new_names(block)
                _lm_pp_apply_copied_column_headers(
                    sheet, hr, insert_at, moving, new_names, src_titles
                )
                # расширить диапазон для следующих шагов постобработки
                ctx = lm_pp_active_context()
                if ctx is not None:
                    _lm_pp_context_refresh(ctx)
                    _lm_pp_context_expand_end_col(
                        ctx, end_col + len(moving)
                    )
        else:
            ok = _lm_pp_apply_column_reorder_uno(
                sheet, end_col, end_row, new_order
            )
        action = "копия" if do_copy else "перенос"
        moved_labels = []
        mi = 0
        while mi < len(moving):
            c = moving[mi]
            try:
                moved_labels.append(
                    "%s(%d)" % (col_index_to_letters(c), c + 1)
                )
            except Exception:
                moved_labels.append(str(c + 1))
            mi = mi + 1
        note = "%s %s → %s" % (
            action,
            ",".join(moved_labels),
            position,
        )
        if rel_token:
            note = note + ", опора «%s»" % rel_token
        if do_copy and ok:
            note = note + ", +%d кол." % len(moving)
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переставить_столбцы",
            ",".join(str(t) for t in tokens),
            "ok" if ok else "ошибка",
            note,
        )
    except Exception as err:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переставить_столбцы",
            ",".join(str(t) for t in tokens),
            "ошибка",
            str(err),
        )


def lm_pp_range_rename_columns(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Переименовать заголовки столбцов: mappings [{old, new}, …] или «Старое=Новое, …».

    В new поддерживаются плейсхолдеры:
      <<Переменные.…>>, [Старое_имя], [Заголовок(N)], [Строка(R, C)].

    JSON: [{"v":1,"fn":"переименовать_столбцы","sheet":"Заказы",
    "mappings":[{"old":"Сумма","new":"Итого_сумма"}]}].

    Карта: «переименовать_столбцы».
    """
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("переименовать_столбцы", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(
        "переименовать_столбцы", extra_args, doc, sheet
    )
    if block is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_столбцы",
            "—",
            "пропуск",
            "нет блока для листа",
        )
        return
    mappings = block.get("mappings") or []
    if not mappings:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_столбцы",
            "—",
            "пропуск",
            "пустые mappings",
        )
        return
    try:
        from libre_macros_param_codec import expand_rename_columns_mapping_pairs
        mappings = expand_rename_columns_mapping_pairs(mappings)
    except Exception:
        pass
    if not mappings:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_столбцы",
            "—",
            "пропуск",
            "пустые mappings",
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    h_sc, h_sr, h_ec, h_er = _lm_pp_range_address(header_row_range)
    hdr_row = h_sr
    end_col, _end_row = get_sheet_used_bounds(sheet)
    if end_col < 0:
        end_col = ec
    data_sr = int(sr)
    data_er = int(er)

    existing_headers = {}
    c = 0
    while c <= end_col:
        try:
            title = cell_text(sheet.getCellByPosition(c, hdr_row))
        except Exception:
            title = ""
        key = str(title or "").strip().casefold()
        if key != "":
            existing_headers[key] = existing_headers.get(key, 0) + 1
        c = c + 1

    renamed = 0
    mi = 0
    while mi < len(mappings):
        item = mappings[mi]
        if not isinstance(item, dict):
            mi = mi + 1
            continue
        old_name = str(item.get("old") or "").strip()
        new_tmpl = str(item.get("new") or "").strip()
        if old_name == "" or new_tmpl == "":
            mi = mi + 1
            continue
        col_idx, col_err = _lm_pp_resolve_column_token_unique(
            old_name, sheet, header_row_range, sc, end_col
        )
        if col_err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переименовать_столбцы",
                old_name,
                "ошибка",
                col_err,
            )
            mi = mi + 1
            continue
        try:
            actual_old = cell_text(sheet.getCellByPosition(col_idx, hdr_row))
        except Exception:
            actual_old = old_name
        if str(actual_old or "").strip() == "":
            actual_old = old_name
        expanded = _lm_pp_expand_rename_name_template(
            new_tmpl,
            doc=doc,
            sheet=sheet,
            old_name=actual_old,
            header_row=hdr_row,
            used_sr=data_sr,
            used_er=data_er,
            log_fn_key="переименовать_столбцы",
        ).strip()
        old_key = str(actual_old or "").strip().casefold()

        def _col_taken(cand, _old_key=old_key, _headers=existing_headers):
            key = str(cand or "").strip().casefold()
            if key == "":
                return True
            if key == _old_key:
                return False
            return _headers.get(key, 0) > 0

        new_name = _lm_pp_finalize_column_name(expanded, is_taken=_col_taken)
        if new_name == "":
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переименовать_столбцы",
                old_name,
                "пропуск",
                "пустое имя после подстановки (шаблон «%s»)" % new_tmpl,
            )
            mi = mi + 1
            continue
        try:
            sheet.getCellByPosition(col_idx, hdr_row).String = new_name
            if old_key != "" and existing_headers.get(old_key, 0) > 0:
                existing_headers[old_key] = existing_headers.get(old_key, 1) - 1
                if existing_headers[old_key] <= 0:
                    existing_headers.pop(old_key, None)
            new_key = new_name.casefold()
            existing_headers[new_key] = existing_headers.get(new_key, 0) + 1
            renamed = renamed + 1
            if LIBRE_MACROS_DEBUG:
                print(
                    "  [переименовать_столбцы] %d «%s» → «%s»"
                    % (col_idx + 1, old_name, new_name)
                )
            note = "столбец %d → «%s»" % (col_idx + 1, new_name)
            if new_tmpl != new_name:
                note = note + " (шаблон «%s»)" % new_tmpl
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переименовать_столбцы",
                old_name,
                "ok",
                note,
            )
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "переименовать_столбцы",
                old_name,
                "ошибка",
                str(err),
            )
        mi = mi + 1
    if renamed == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_столбцы",
            "итог",
            "пропуск",
            "ни один столбец не переименован",
        )
    else:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "переименовать_столбцы",
            "итог",
            "ok",
            "переименовано: %d" % renamed,
        )


def lm_pp_range_copy_values(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Скопировать столбцы в буфер и вставить как значения
    (формулы заменяются результатом, оформление не меняется).

    Колонка C: столбцы через , или ; — номер (1-based), буква A или имя заголовка.
    Пример: «Сумма,Количество» или «3,E,Итого».
    Пустая C — вся значащая область листа (как выделение в перекрестье заголовков).

    Карта: «копировать_значения».
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("копировать_значения", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet("копировать_значения", extra_args, doc, sheet)
    if block is not None and block.get("whole_sheet"):
        spec_text = ""
    elif block is not None:
        cols = block.get("columns") or []
        spec_text = ",".join(str(c) for c in cols)
    else:
        spec_text = _lm_pp_join_extra_args(extra_args)
    if spec_text == "":
        ua_sc, ua_sr, ua_ec, ua_er = _lm_vlookup_sheet_used_area(sheet)
        if ua_er >= ua_sr:
            c0, r0, c1, r1 = ua_sc, ua_sr, ua_ec, ua_er
        else:
            hsc, hsr, hec, her = _lm_pp_range_address(header_row_range)
            c0 = min(sc, hsc)
            c1 = max(ec, hec)
            r0 = min(sr, hsr)
            r1 = max(er, her)
        if r1 < r0 or c1 < c0:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "копировать_значения",
                "—",
                "пропуск",
                "нет данных на листе",
            )
            return
        cell_range = sheet.getCellRangeByPosition(c0, r0, c1, r1)
        ok = _lm_pp_copy_range_as_values(doc, sheet, cell_range)
        _lm_log_postprocess(
            doc,
            sheet_name,
            "копировать_значения",
            "—",
            "ok" if ok else "ошибка",
            "весь лист (%d×%d)" % (c1 - c0 + 1, r1 - r0 + 1),
        )
        return

    cols = _lm_pp_resolve_column_list(spec_text, sheet, header_row_range, sc, ec)
    if len(cols) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "копировать_значения",
            spec_text,
            "ошибка",
            "не найдены столбцы",
        )
        return

    if er < sr:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "копировать_значения",
            spec_text,
            "пропуск",
            "нет строк данных",
        )
        return

    groups = _lm_pp_group_consecutive_columns(cols)
    ok_count = 0
    fail_count = 0
    gi = 0
    while gi < len(groups):
        c0, c1 = groups[gi]
        cell_range = sheet.getCellRangeByPosition(c0, sr, c1, er)
        if _lm_pp_copy_range_as_values(doc, sheet, cell_range):
            ok_count = ok_count + 1
        else:
            fail_count = fail_count + 1
        gi = gi + 1
        _lm_ui_yield(counter=gi)

    status = "ok" if fail_count == 0 else "ошибка"
    _lm_log_postprocess(
        doc,
        sheet_name,
        "копировать_значения",
        spec_text,
        status,
        "групп столбцов: %d, успех: %d, ошибка: %d"
        % (len(groups), ok_count, fail_count),
    )


LM_PP_REPLACE_EMPTY = "_ПУСТО_"
LM_PP_REPLACE_UPPER = "_РЕГИСТР_ВВЕРХ_"
LM_PP_REPLACE_LOWER = "_РЕГИСТР_ВНИЗ_"
_LM_PP_REPLACE_EMPTY_FIND = object()
_LM_PP_REPLACE_REPEAT_RE = re.compile(r"^_ПОВТОР\*(\d+)_$", re.IGNORECASE)
_LM_PP_REPLACE_EMPTY_ALIASES = frozenset(
    ("", "_пусто_", "_empty_", "__пусто__")
)


def _lm_pp_replace_values_is_empty_token(token):
    try:
        from libre_macros_param_codec import unwrap_column_name_token

        bare, _q = unwrap_column_name_token(token)
    except Exception:
        bare = str(token or "")
    raw = str(bare if bare is not None else "")
    key = raw.strip().casefold()
    if key in ("_пусто_", "_empty_", "__пусто__"):
        return True
    # только действительно пустой шаблон; пробелы (' ') — обычный поиск
    return raw == ""


def _lm_pp_replace_values_expand_var_token( token, doc=None, sheet=None, sheet_name_hint=None):
    """
    Токен find/replace: снять '…'; <<Переменные.…>> подставить только вне кавычек.
    В одинарных кавычках плейсхолдеры не трогаем (литерал).
    """
    s = str(token if token is not None else "")
    try:
        from libre_macros_param_codec import unwrap_column_name_token

        bare, quoted = unwrap_column_name_token(s)
    except Exception:
        bare, quoted = s, False
    if quoted:
        return str(bare)
    if "<<" not in bare:
        return str(bare)
    prev_ctx = lm_pp_active_context()
    ctx = dict(prev_ctx or {})
    if not ctx.get("source_variables_map"):
        try:
            import libre_macros_collect_cfg as _cw_cfg

            ctx["source_variables_map"] = (
                getattr(_cw_cfg, "_MERGE_SOURCE_VARIABLES_MAP", {}) or {}
            )
            ctx["source_variables_order"] = (
                getattr(_cw_cfg, "_MERGE_SOURCE_VARIABLES_ORDER", []) or []
            )
        except Exception:
            pass
    lm_pp_set_active_context(ctx)
    try:
        return str(
            _lm_pp_expand_source_variables_template(
                bare,
                doc=doc,
                sheet=sheet,
                log_fn_key="замена_значений",
                for_formula=False,
                sheet_name_hint=sheet_name_hint,
            )
            or ""
        )
    except Exception:
        return str(bare)
    finally:
        lm_pp_set_active_context(prev_ctx)


def _lm_pp_replace_values_expand_list( tokens, doc=None, sheet=None, sheet_name_hint=None):
    out = []
    for t in tokens or []:
        out.append(
            _lm_pp_replace_values_expand_var_token(
                t, doc=doc, sheet=sheet, sheet_name_hint=sheet_name_hint
            )
        )
    return out


def _lm_pp_replace_values_compile_pattern(pattern, case_insensitive=False):
    """Шаблон поиска: * — любая подстрока; _ПУСТО_ — пустая ячейка; иначе regex."""
    if _lm_pp_replace_values_is_empty_token(pattern):
        return _LM_PP_REPLACE_EMPTY_FIND
    p = str(pattern or "")
    if p == "":
        return None
    rx_parts = []
    i = 0
    while i < len(p):
        if p[i] == "*":
            rx_parts.append(".*")
            i += 1
        else:
            j = i
            while j < len(p) and p[j] != "*":
                j += 1
            rx_parts.append(re.escape(p[i:j]))
            i = j
    flags = re.IGNORECASE if case_insensitive else 0
    try:
        return re.compile("".join(rx_parts), flags)
    except Exception:
        return None


def _lm_pp_replace_values_resolve_repl(repl_token, matched_text):
    """Подставить replace-токен; краевые пробелы литерала сохраняются."""
    raw = str(repl_token if repl_token is not None else "")
    # спецтокены (_ПУСТО_, _РЕГИСТР_*, _ПОВТОР*N_) — по обрезанной форме
    t = raw.strip()
    if _lm_pp_replace_values_is_empty_token(raw):
        return ""
    tc = t.casefold()
    if tc == LM_PP_REPLACE_UPPER.casefold():
        return str(matched_text or "").upper()
    if tc == LM_PP_REPLACE_LOWER.casefold():
        return str(matched_text or "").lower()
    m = _LM_PP_REPLACE_REPEAT_RE.match(t)
    if m is not None:
        try:
            n = int(m.group(1))
        except Exception:
            n = 1
        if n < 0:
            n = 0
        return str(matched_text or "") * n
    return raw


def _lm_pp_replace_values_parse_replace_mode(block):
    if not isinstance(block, dict):
        return "manual"
    raw = ""
    for key in ("replace_mode", "repl_mode", "to_mode"):
        v = block.get(key)
        if v is None:
            continue
        s = str(v).strip()
        if s != "":
            raw = s
            break
    t = raw.casefold().replace("ё", "е")
    if t in (
        "lambda",
        "лямбда",
        "expr",
        "выражение",
    ):
        return "lambda"
    if t in (
        "manual",
        "вручную",
        "ручной",
        "ручной ввод",
        "текст",
        "text",
        "list",
        "",
    ):
        # авто по наличию expr
        if t == "" and _lm_pp_replace_values_replace_expr_text(block):
            has_list = False
            for k in ("replace", "replacements", "to"):
                if k in block and block.get(k) not in (None, "", [], ()):
                    has_list = True
                    break
            if not has_list:
                return "lambda"
        return "manual"
    return "manual"


def _lm_pp_replace_values_replace_expr_text(block):
    if not isinstance(block, dict):
        return ""
    for key in ("replace_expr", "replace_lambda", "to_expr", "to_lambda"):
        v = block.get(key)
        if v is None:
            continue
        s = str(v).strip()
        if s != "":
            return s
    return ""


def _lm_pp_replace_values_coerce_repl_result(result):
    """Результат лямбды замены → список строк (одно значение или кортеж/список)."""
    if result is None:
        return []
    if isinstance(result, (list, tuple)):
        out = []
        for x in result:
            if x is None:
                out.append("")
            elif isinstance(x, bool):
                out.append("true" if x else "false")
            else:
                out.append(str(x))
        return out
    if isinstance(result, bool):
        return ["true" if result else "false"]
    return [str(result)]


def _lm_pp_replace_values_align_find_repl(find_list, repl_list):
    """
    Согласовать длины find/replace.

    1 replace → broadcast на все find (как раньше).
    N=N → ок.
    N < find → обрезать find, warning.
    N > find → обрезать replace, warning.

    → (find_out, repl_out, warn_or_None)
    """
    finds = list(find_list or [])
    repls = list(repl_list or [])
    nf = len(finds)
    nr = len(repls)
    if nf == 0:
        return [], [], None
    if nr == 0:
        return finds, [LM_PP_REPLACE_EMPTY], None
    if nr == 1:
        return finds, repls, None
    if nr == nf:
        return finds, repls, None
    if nr < nf:
        return (
            finds[:nr],
            repls,
            "replace короче find: применено %d из %d поисков (остальные find пропущены)"
            % (nr, nf),
        )
    return (
        finds,
        repls[:nf],
        "replace длиннее find: использовано %d из %d значений (лишние replace пропущены)"
        % (nf, nr),
    )


def _lm_pp_replace_values_eval_replace_expr( expr_text, rows_accessor=None, upd_store=None, row_id=0):
    """
    Скомпилировать и один раз вызвать лямбду замены.
    → (repl_list, error_or_None)
    """
    code = str(expr_text or "").strip()
    if code == "":
        return None, "replace_mode=lambda: нужен replace_expr (lambda rows: …)"
    from libre_macros_lambda_column_lib import (
        call_rows_lambda,
        compile_rows_lambda,
        make_upd_rows_store,
    )

    try:
        packed = compile_rows_lambda(code)
    except Exception as err:
        return None, "лямбда замены: %s" % err
    if packed is None:
        return None, "replace_mode=lambda: пустой replace_expr"
    fn, glo = packed
    store = upd_store if upd_store is not None else make_upd_rows_store()
    try:
        result = call_rows_lambda(
            fn,
            glo,
            int(row_id or 0),
            rows_accessor=rows_accessor,
            upd_store=store,
        )
    except Exception as err:
        return None, "лямбда замены (eval): %s" % err
    return _lm_pp_replace_values_coerce_repl_result(result), None


def _lm_pp_replace_values_transform_text( s, find_norm, repl_norm, case_insensitive=False, pats=None ):
    """Применить все пары find/replace к строке s. pats — предкомпилированные regex."""
    out = str(s or "")
    i = 0
    while i < len(find_norm):
        f = find_norm[i]
        if f == "":
            i += 1
            continue
        if len(repl_norm) == 1:
            repl_tok = repl_norm[0]
        elif i < len(repl_norm):
            repl_tok = repl_norm[i]
        else:
            break
        if _lm_pp_replace_values_is_empty_token(f):
            if str(out or "").strip() == "":
                out = _lm_pp_replace_values_resolve_repl(repl_tok, "")
            i += 1
            continue
        pat = None
        if pats is not None and i < len(pats):
            pat = pats[i]
        else:
            pat = _lm_pp_replace_values_compile_pattern(f, case_insensitive)
        if pat is None:
            i += 1
            continue
        if pat is _LM_PP_REPLACE_EMPTY_FIND:
            if str(out or "").strip() == "":
                out = _lm_pp_replace_values_resolve_repl(repl_tok, "")
            i += 1
            continue

        def _sub(match, tok=repl_tok):
            return _lm_pp_replace_values_resolve_repl(tok, match.group(0))

        try:
            out = pat.sub(_sub, out)
        except Exception:
            pass
        i += 1
    return out


def _lm_pp_replace_values_compile_all(find_norm, case_insensitive=False):
    pats = []
    i = 0
    while i < len(find_norm):
        f = find_norm[i]
        if f == "":
            pats.append(None)
        else:
            pats.append(
                _lm_pp_replace_values_compile_pattern(f, case_insensitive)
            )
        i += 1
    return pats


def _lm_pp_replace_values_has_hit(s, pats):
    i = 0
    while i < len(pats):
        pat = pats[i]
        i += 1
        if pat is _LM_PP_REPLACE_EMPTY_FIND:
            if str(s or "").strip() == "":
                return True
            continue
        if pat is None:
            continue
        try:
            if pat.search(s):
                return True
        except Exception:
            pass
    return False


def _lm_pp_replace_values_pattern_hits(s, find_norm, case_insensitive=False):
    """Список bool — было ли совпадение каждого шаблона в s."""
    hits = []
    i = 0
    while i < len(find_norm):
        f = find_norm[i]
        if f == "":
            hits.append(False)
            i += 1
            continue
        pat = _lm_pp_replace_values_compile_pattern(f, case_insensitive)
        hit = False
        if pat is _LM_PP_REPLACE_EMPTY_FIND:
            hit = str(s or "").strip() == ""
        elif pat is not None:
            try:
                hit = bool(pat.search(s))
            except Exception:
                hit = False
        hits.append(hit)
        i += 1
    return hits


def _lm_pp_replace_values_split_list(value):
    """
    Привести поле списка (find/replace) к list[str].
    Допускает строку с разделителями ',' / ';' или list/tuple.
    В одинарных кавычках '…' запятая/«;» не режут токен (',' → «,»).
    Кавычки сохраняются ('<<Переменные.X>>' — литерал без подстановки).
    """
    try:
        from libre_macros_param_codec import split_sep_list_respecting_quotes

        return split_sep_list_respecting_quotes(
            value, keep_empty_quoted=True, keep_quotes=True
        )
    except Exception:
        if value is None:
            return []
        if isinstance(value, (list, tuple)):
            out = []
            for v in value:
                s = str(v).strip()
                if s != "":
                    out.append(s)
            return out
        s = str(value).strip()
        if s == "":
            return []
        parts = []
        for p in re.split(r"[,;]", s):
            p = p.strip()
            if p != "":
                parts.append(p)
        return parts


def _lm_pp_replace_values_normalize_text(s, squeeze_spaces=False):
    if s is None:
        return ""
    t = str(s)
    if squeeze_spaces:
        t = re.sub(r"\s+", " ", t.strip())
    # squeeze_spaces=false — не strip: краевые пробелы в find/ячейке значимы
    return t


def _lm_pp_replace_values_should_skip_column(sheet, col, sr, er):
    """
    Определить по первой значащей строке, можно ли обрабатывать столбец как текстовый.
    Пропуск: числовые/даты/логические — приближенно по типу VALUE либо формуле без текста.
    """
    if sheet is None:
        return False, ""
    col = int(col)
    r = int(sr)
    end_r = int(er)
    while r <= end_r:
        try:
            cell = sheet.getCellByPosition(col, r)
        except Exception:
            r += 1
            continue
        try:
            if cell_text(cell) == "":
                r += 1
                continue
        except Exception:
            r += 1
            continue
        try:
            t = int(cell.getType())
        except Exception:
            # если тип не получить — считаем текстовым
            return False, ""
        # CellContentType: EMPTY=0, VALUE=1, TEXT=2, FORMULA=3
        if t == 1:
            return True, "тип VALUE (число/дата/логическое)"
        if t == 3:
            # Формула: если текст результата пустой — вероятно число/дата/логическое
            try:
                s = str(cell.getString() or "").strip()
            except Exception:
                s = ""
            if s == "":
                return True, "формула без текста (вероятно число/дата/логическое)"
        return False, ""
    return False, ""


_LM_PP_REPLACE_MATCH_PICK_ALIASES = {
    "all": "all",
    "все": "all",
    "*": "all",
    "все_совпадения": "all",
    "first": "first",
    "первый": "first",
    "первое": "first",
    "first_match": "first",
    "last": "last",
    "последний": "last",
    "последнее": "last",
    "last_match": "last",
    "odd": "odd",
    "нечетный": "odd",
    "нечётный": "odd",
    "нечетные": "odd",
    "нечётные": "odd",
    "even": "even",
    "четный": "even",
    "чётный": "even",
    "четные": "even",
    "чётные": "even",
    "lambda": "lambda",
    "лямбда": "lambda",
    "expr": "lambda",
}


def _lm_pp_replace_values_parse_match_pick(block):
    if not isinstance(block, dict):
        return "all"
    raw = None
    for key in ("match_pick", "which", "occurrence", "matches"):
        if key in block and block.get(key) not in (None, ""):
            raw = block.get(key)
            break
    t = str(raw or "all").strip().casefold()
    if t == "":
        return "all"
    return _LM_PP_REPLACE_MATCH_PICK_ALIASES.get(t, "all")


def _lm_pp_replace_values_row_filter_text(block):
    if not isinstance(block, dict):
        return ""
    for key in ("row_filter", "row_expr", "row_lambda", "filter"):
        v = block.get(key)
        if v is None:
            continue
        s = str(v).strip()
        if s != "":
            return s
    return ""


def _lm_pp_replace_values_match_expr_text(block):
    if not isinstance(block, dict):
        return ""
    for key in ("match_expr", "match_lambda", "match_filter"):
        v = block.get(key)
        if v is None:
            continue
        s = str(v).strip()
        if s != "":
            return s
    return ""


def _lm_pp_replace_values_pick_indices(n, mode, pred=None):
    """Индексы кандидатов (0-based). odd/even — 1-based среди совпадений."""
    try:
        n = int(n)
    except Exception:
        n = 0
    if n <= 0:
        return []
    mode = str(mode or "all").strip().casefold()
    if mode == "first":
        return [0]
    if mode == "last":
        return [n - 1]
    if mode == "odd":
        out = []
        i = 0
        while i < n:
            if (i + 1) % 2 == 1:
                out.append(i)
            i += 1
        return out
    if mode == "even":
        out = []
        i = 0
        while i < n:
            if (i + 1) % 2 == 0:
                out.append(i)
            i += 1
        return out
    if mode == "lambda":
        out = []
        i = 0
        while i < n:
            ok = False
            if pred is not None:
                try:
                    ok = bool(pred(i, n))
                except Exception:
                    ok = False
            if ok:
                out.append(i)
            i += 1
        return out
    return list(range(n))


def _lm_pp_replace_values_is_blank_matrix_cell(v):
    if v is None:
        return True
    if isinstance(v, (str, unicode)):
        return str(v).strip() == ""
    return False


def _lm_pp_replace_values_is_non_text_value(v):
    if isinstance(v, bool):
        return True
    if isinstance(v, (int, float)):
        return True
    try:
        import datetime as _dt

        if isinstance(v, (_dt.date, _dt.datetime, _dt.time)):
            return True
    except Exception:
        pass
    return False


def _lm_pp_replace_values_skip_column_matrix(matrix, local_col):
    """True — столбец не текстовый (первая значащая ячейка число/дата/bool)."""
    local_col = int(local_col)
    ri = 0
    n = len(matrix) if matrix is not None else 0
    while ri < n:
        row = matrix[ri]
        try:
            v = row[local_col] if local_col < len(row) else None
        except Exception:
            v = None
        if _lm_pp_replace_values_is_blank_matrix_cell(v):
            ri += 1
            continue
        return _lm_pp_replace_values_is_non_text_value(v)
    return False


def lm_pp_replace_values_prepare_filters(block):
    """
    Скомпилировать row_filter / match_expr.

    Возвращает dict: error, row_fn, row_glo, match_pick, match_fn, match_glo.
    """
    from libre_macros_lambda_column_lib import compile_rows_lambda

    row_text = _lm_pp_replace_values_row_filter_text(block)
    match_pick = _lm_pp_replace_values_parse_match_pick(block)
    match_text = _lm_pp_replace_values_match_expr_text(block)
    row_fn = None
    row_glo = None
    match_fn = None
    match_glo = None
    if row_text:
        try:
            packed = compile_rows_lambda(row_text)
        except Exception as err:
            return {
                "error": "фильтр строк: %s" % err,
                "row_fn": None,
                "row_glo": None,
                "match_pick": match_pick,
                "match_fn": None,
                "match_glo": None,
            }
        if packed:
            row_fn, row_glo = packed
    if match_pick == "lambda":
        if match_text == "":
            return {
                "error": "match_pick=lambda: нужен match_expr (lambda rows: …)",
                "row_fn": row_fn,
                "row_glo": row_glo,
                "match_pick": match_pick,
                "match_fn": None,
                "match_glo": None,
            }
        try:
            packed = compile_rows_lambda(match_text)
        except Exception as err:
            return {
                "error": "лямбда отбора совпадений: %s" % err,
                "row_fn": row_fn,
                "row_glo": row_glo,
                "match_pick": match_pick,
                "match_fn": None,
                "match_glo": None,
            }
        if packed:
            match_fn, match_glo = packed
        else:
            return {
                "error": "match_pick=lambda: пустой match_expr",
                "row_fn": row_fn,
                "row_glo": row_glo,
                "match_pick": match_pick,
                "match_fn": None,
                "match_glo": None,
            }
    return {
        "error": None,
        "row_fn": row_fn,
        "row_glo": row_glo,
        "match_pick": match_pick,
        "match_fn": match_fn,
        "match_glo": match_glo,
    }


def lm_pp_replace_values_apply_matrix( matrix, headers, col_abs_indices, find_norm, repl_norm, case_insensitive=False, squeeze_spaces=False, start_col=0, row_fn=None, row_glo=None, match_pick="all", match_fn=None, match_glo=None):
    """
    Замена в памяти по матрице data-строк. Один проход, без UNO.

    matrix — list[list] (мутабельно). headers[i] ↔ столбец start_col+i.
    col_abs_indices — абсолютные 0-based столбцы листа.
    """
    from libre_macros_lambda_column_lib import call_rows_lambda, make_rows_accessor, make_upd_rows_store

    if matrix is None:
        matrix = []
    n_rows = len(matrix)
    start_col = int(start_col)
    accessor = make_rows_accessor(matrix, headers, start_col=start_col)
    upd_store = make_upd_rows_store()
    if row_glo is not None:
        row_glo["_lm_rows"] = accessor
    if match_glo is not None:
        match_glo["_lm_rows"] = accessor

    row_ok = [True] * n_rows
    if row_fn is not None:
        ri = 0
        while ri < n_rows:
            try:
                row_ok[ri] = bool(
                    call_rows_lambda(
                        row_fn, row_glo, ri,
                        rows_accessor=accessor, upd_store=upd_store,
                    )
                )
            except Exception:
                row_ok[ri] = False
            if row_ok[ri]:
                upd_store.commit(ri)
            ri += 1

    pats = _lm_pp_replace_values_compile_all(find_norm, case_insensitive)
    skipped_cols = 0
    changed_cols = 0
    changed_cells = 0
    ci = 0
    n_cols = len(col_abs_indices) if col_abs_indices is not None else 0
    while ci < n_cols:
        abs_c = int(col_abs_indices[ci])
        ci += 1
        local = abs_c - start_col
        if local < 0:
            skipped_cols += 1
            continue
        if _lm_pp_replace_values_skip_column_matrix(matrix, local):
            skipped_cols += 1
            continue

        candidates = []
        ri = 0
        while ri < n_rows:
            if not row_ok[ri]:
                ri += 1
                continue
            row = matrix[ri]
            try:
                v = row[local] if local < len(row) else None
            except Exception:
                v = None
            if _lm_pp_replace_values_is_non_text_value(v):
                ri += 1
                continue
            s0 = _lm_pp_replace_values_normalize_text(
                v, squeeze_spaces=squeeze_spaces
            )
            if _lm_pp_replace_values_has_hit(s0, pats):
                candidates.append(ri)
            ri += 1

        def _match_pred(idx, n, _cands=candidates, _local=local, _upd=upd_store):
            if match_fn is None or match_glo is None:
                return False
            ri2 = _cands[idx]
            row2 = matrix[ri2]
            try:
                v2 = row2[_local] if _local < len(row2) else None
            except Exception:
                v2 = None
            text = _lm_pp_replace_values_normalize_text(
                v2, squeeze_spaces=squeeze_spaces
            )
            extra = {"idx": idx, "n": n, "text": text, "i": idx + 1}
            try:
                return bool(
                    call_rows_lambda(
                        match_fn, match_glo, ri2, extra=extra,
                        rows_accessor=accessor, upd_store=_upd,
                    )
                )
            except Exception:
                return False

        pick = _lm_pp_replace_values_pick_indices(
            len(candidates),
            match_pick,
            pred=_match_pred if match_pick == "lambda" else None,
        )
        selected = {}
        pi = 0
        while pi < len(pick):
            selected[int(pick[pi])] = True
            pi += 1

        col_changed = False
        k = 0
        while k < len(candidates):
            if k not in selected:
                k += 1
                continue
            ri = candidates[k]
            row = matrix[ri]
            try:
                v = row[local] if local < len(row) else None
            except Exception:
                v = None
            s0 = _lm_pp_replace_values_normalize_text(
                v, squeeze_spaces=squeeze_spaces
            )
            s = _lm_pp_replace_values_transform_text(
                s0,
                find_norm,
                repl_norm,
                case_insensitive=case_insensitive,
                pats=pats,
            )
            if s != s0:
                while local >= len(row):
                    row.append("")
                row[local] = s
                changed_cells += 1
                col_changed = True
            k += 1
        if col_changed:
            changed_cols += 1

    return {
        "changed_cells": changed_cells,
        "changed_cols": changed_cols,
        "skipped_cols": skipped_cols,
        "total_cols": n_cols,
    }


def lm_pp_range_replace_values(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Замена подстрок в текстовых столбцах.

    JSON колонки C (пример):
      [{"v":1,"fn":"замена_значений","columns":["ФИО","Комментарий"],
        "find":["ООО","*MTP"], "replace":["ОБЩЕСТВО","_ПУСТО_"],
        "case_insensitive":true, "squeeze_spaces":true}]

    Правила:
    - replace: либо 1 значение (для всех find), либо столько же, сколько find;
    - _ПУСТО_ / _EMPTY_ / пустое replace — стирание найденного фрагмента;
    - _ПУСТО_ в find — пустая ячейка (замена всего содержимого);
    - _ПОВТОР*N_, _РЕГИСТР_ВВЕРХ_, _РЕГИСТР_ВНИЗ_ — служебные замены;
    - в find символ * — любое число символов слева/справа;
    - <<Переменные.…>> в find/replace — из runtime-карты; в '…' плейсхолдеры не трогаем;
    - столбец пропускается, если первая значащая ячейка число/дата/bool;
    - row_filter: lambda rows: … (sugar rows("Col") / rows("Col")[-1]);
    - match_pick: all|first|last|odd|even|lambda (+ match_expr); odd/even — 1-based
      среди совпавших строк столбца. Внутри выбранной ячейки — все вхождения find.
    Чтение: 1–2 getDataArray на прямоугольник; запись: один setDataArray.
    """
    sc, sr, ec, er = _lm_pp_range_address(data_range)
    sheet_name = sheet.Name if sheet is not None else ""
    if _lm_pp_param_skip_sheet("замена_значений", extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet("замена_значений", extra_args, doc, sheet)
    if block is None:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "замена_значений",
            "пропуск",
            "нет блока JSON для листа",
        )
        return

    cols_in = block.get("columns") or []
    if isinstance(cols_in, (list, tuple)):
        spec_text = ",".join(str(c) for c in cols_in if str(c).strip() != "")
    else:
        spec_text = str(cols_in or "").strip()

    find_list = _lm_pp_replace_values_split_list(
        block.get("find") or block.get("search") or block.get("patterns")
    )
    find_list = _lm_pp_replace_values_expand_list(
        find_list, doc=doc, sheet=sheet, sheet_name_hint=sheet_name
    )
    if len(find_list) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "замена_значений",
            "ошибка",
            "не заданы подстроки поиска (find)",
        )
        return

    case_insensitive = bool(
        block.get("case_insensitive")
        or block.get("без_учета_регистра")
        or block.get("ignore_case")
    )
    squeeze_spaces = bool(
        block.get("squeeze_spaces")
        or block.get("сжать_пробелы")
        or block.get("trim_spaces")
    )

    prepared = lm_pp_replace_values_prepare_filters(block)
    if prepared.get("error"):
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "замена_значений",
            "ошибка",
            prepared.get("error"),
        )
        return

    if er < sr:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "замена_значений",
            "пропуск",
            "нет строк данных",
        )
        return

    try:
        h_sc, h_sr, h_ec, _h_er = _lm_pp_range_address(header_row_range)
    except Exception:
        h_sc, h_sr, h_ec = sc, max(0, int(sr) - 1), ec

    combined = (
        int(h_sr) + 1 == int(sr)
        and int(h_sc) == int(sc)
        and int(h_ec) == int(ec)
    )
    header_vals = []
    matrix = []
    write_rng = None
    try:
        if combined:
            rng = sheet.getCellRangeByPosition(int(sc), int(h_sr), int(ec), int(er))
            arr = rng.getDataArray()
            if arr:
                header_vals = list(arr[0])
                matrix = [list(r) for r in arr[1:]]
            write_rng = sheet.getCellRangeByPosition(int(sc), int(sr), int(ec), int(er))
        else:
            h_rng = sheet.getCellRangeByPosition(
                int(h_sc), int(h_sr), int(h_ec), int(h_sr)
            )
            h_arr = h_rng.getDataArray()
            if h_arr:
                header_vals = list(h_arr[0])
            write_rng = sheet.getCellRangeByPosition(int(sc), int(sr), int(ec), int(er))
            arr = write_rng.getDataArray()
            if arr:
                matrix = [list(r) for r in arr]
    except Exception as err:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "замена_значений",
            "ошибка",
            "getDataArray: %s" % err,
        )
        return

    from libre_macros_lambda_column_lib import (
        make_rows_accessor,
        make_upd_rows_store,
        resolve_column_token_from_titles,
    )

    titles_resolve = []
    hi = 0
    while hi < len(header_vals):
        hv = header_vals[hi]
        titles_resolve.append("" if hv is None else str(hv))
        hi += 1
    width = int(ec) - int(sc) + 1
    if width < 0:
        width = 0
    titles = [""] * width
    hi = 0
    while hi < len(header_vals):
        abs_c = int(h_sc) + hi
        loc = abs_c - int(sc)
        if 0 <= loc < width:
            hv = header_vals[hi]
            titles[loc] = "" if hv is None else str(hv)
        hi += 1

    tokens = _lm_pp_split_column_tokens(spec_text)
    cols = []
    ti = 0
    while ti < len(tokens):
        c = resolve_column_token_from_titles(
            tokens[ti], titles_resolve, int(h_sc)
        )
        if c is not None and int(sc) <= int(c) <= int(ec) and int(c) not in cols:
            cols.append(int(c))
        ti += 1
    cols.sort()
    if len(cols) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "замена_значений",
            "ошибка",
            "не найдены столбцы (columns): %s" % (spec_text or "—",),
        )
        return

    # replace: вручную или лямбда (один вызов → значение/кортеж)
    replace_mode = _lm_pp_replace_values_parse_replace_mode(block)
    repl_list = []
    if replace_mode == "lambda":
        rows_accessor = make_rows_accessor(matrix, titles, start_col=int(sc))
        repl_list, rerr = _lm_pp_replace_values_eval_replace_expr(
            _lm_pp_replace_values_replace_expr_text(block),
            rows_accessor=rows_accessor,
            upd_store=make_upd_rows_store(),
            row_id=0,
        )
        if rerr:
            _lm_log_postprocess(
                doc, sheet_name, "диапазон", "замена_значений", "ошибка", rerr
            )
            return
    else:
        repl_list = _lm_pp_replace_values_split_list(
            block.get("replace")
            if "replace" in block
            else (block.get("replacements") or block.get("to"))
        )
        repl_list = _lm_pp_replace_values_expand_list(
            repl_list, doc=doc, sheet=sheet, sheet_name_hint=sheet_name
        )
    find_list, repl_list, align_warn = _lm_pp_replace_values_align_find_repl(
        find_list, repl_list
    )
    if align_warn:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "замена_значений",
            "предупреждение",
            align_warn,
        )
    if len(find_list) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            "замена_значений",
            "ошибка",
            "после согласования find/replace не осталось пар",
        )
        return

    find_norm = [
        _lm_pp_replace_values_normalize_text(x, squeeze_spaces=squeeze_spaces)
        for x in find_list
    ]
    repl_norm = [
        _lm_pp_replace_values_normalize_text(x, squeeze_spaces=squeeze_spaces)
        for x in repl_list
    ]
    for ri in range(len(repl_norm)):
        if _lm_pp_replace_values_is_empty_token(repl_norm[ri]):
            repl_norm[ri] = LM_PP_REPLACE_EMPTY

    stats = lm_pp_replace_values_apply_matrix(
        matrix,
        titles,
        cols,
        find_norm,
        repl_norm,
        case_insensitive=case_insensitive,
        squeeze_spaces=squeeze_spaces,
        start_col=int(sc),
        row_fn=prepared.get("row_fn"),
        row_glo=prepared.get("row_glo"),
        match_pick=prepared.get("match_pick") or "all",
        match_fn=prepared.get("match_fn"),
        match_glo=prepared.get("match_glo"),
    )
    changed_cells = int(stats.get("changed_cells") or 0)
    changed_cols = int(stats.get("changed_cols") or 0)
    skipped_cols = int(stats.get("skipped_cols") or 0)
    total_cols = int(stats.get("total_cols") or len(cols))

    if changed_cells > 0 and write_rng is not None:
        out_arr = []
        ri = 0
        while ri < len(matrix):
            out_arr.append(tuple(matrix[ri]))
            ri += 1
        try:
            write_rng.setDataArray(tuple(out_arr))
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                "замена_значений",
                "ошибка",
                "setDataArray: %s" % err,
            )
            return

    note = "столбцов: %d, пропуск: %d, изменено столбцов: %d, изменено ячеек: %d" % (
        total_cols,
        skipped_cols,
        changed_cols,
        changed_cells,
    )
    pick = prepared.get("match_pick") or "all"
    if pick != "all":
        note = "%s, отбор: %s" % (note, pick)
    if prepared.get("row_fn") is not None:
        note = "%s, фильтр строк" % note
    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        "замена_значений",
        "ok",
        note,
    )


def lm_pp_range_normalize_text(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Текстовые операции в столбцах: стек ops (trim, case, омоглифы, base64, шифр…).

    JSON колонки C (пример):
      [{"v":1,"fn":"текстовые_операции","columns":["ФИО"],
        "ops":[{"op":"collapse_ws"},{"op":"lang_homoglyphs"},{"op":"case","mode":"proper"}]}]

    Per-op ``row_filter``: all | non_empty | empty | lambda (+ row_filter_lambda).
    Лямбда: ``lambda rows: rows("Столбец")`` / ``rows("Столбец")[-1]`` (соседняя строка).
    ``upd_rows("Столбец")`` / ``upd_rows_seen(...)`` — уже обработанные строки выше текущей.
    Legacy: ``lambda row: row.get("Наименование")`` (dict заголовков + ``row_id``).
    """
    from libre_macros_normalize_lib import (
        NormalizeOpError,
        apply_ops,
        build_row_dict,
        cell_text_for_ops,
        coerce_ops_list,
        compile_ops_compute_fns,
        compile_ops_row_filters,
        needs_crypto_key,
        resolve_crypto_key,
    )

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    sheet_name = sheet.Name if sheet is not None else ""
    fn_key = "текстовые_операции"
    if _lm_pp_param_skip_sheet(fn_key, extra_args, doc, sheet):
        return
    block = _lm_pp_param_block_for_sheet(fn_key, extra_args, doc, sheet)
    if block is None:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_key, "пропуск", "нет блока JSON для листа"
        )
        return

    cols_in = block.get("columns") or []
    if isinstance(cols_in, (list, tuple)):
        spec_text = ",".join(str(c) for c in cols_in if str(c).strip() != "")
    else:
        spec_text = str(cols_in or "").strip()
    cols = _lm_pp_resolve_column_list(spec_text, sheet, header_row_range, sc, ec)
    if len(cols) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "диапазон",
            fn_key,
            "ошибка",
            "не найдены столбцы (columns): %s" % (spec_text or "—",),
        )
        return

    ops = coerce_ops_list(block.get("ops"))
    if len(ops) == 0:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_key, "ошибка", "пустой стек ops"
        )
        return

    if er < sr:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_key, "пропуск", "нет строк данных"
        )
        return

    non_text = str(block.get("non_text") or "skip").strip() or "skip"
    on_error = str(block.get("on_error") or "keep").strip() or "keep"
    empty_cells = str(block.get("empty_cells") or "skip").strip() or "skip"
    as_values = block.get("as_values")
    if as_values is None:
        as_values = True
    else:
        try:
            as_values = bool(lm_parse_bool_param(as_values, default=True))
        except Exception:
            as_values = bool(as_values)
    # getDataArray уже отдаёт вычисленные значения; флаг сохранён для совместимости JSON/визарда.
    _ = as_values

    crypto_key = None
    if needs_crypto_key(ops):
        try:
            crypto_key = resolve_crypto_key(block, doc=doc, default_sheet=sheet)
        except NormalizeOpError as err:
            _lm_log_postprocess(
                doc, sheet_name, "диапазон", fn_key, "ошибка", str(err)
            )
            return
        if crypto_key in (None, ""):
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_key,
                "ошибка",
                "для encrypt/decrypt нужен key / key_cell / key_file",
            )
            return

    try:
        row_filters = compile_ops_row_filters(ops)
    except Exception as err:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_key, "ошибка", "фильтр строк: %s" % err
        )
        return
    try:
        compute_fns = compile_ops_compute_fns(ops)
    except Exception as err:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_key, "ошибка", "вычислить: %s" % err
        )
        return

    end_col_used, _er_used = get_sheet_used_bounds(sheet)
    row_end_col = int(ec)
    if end_col_used is not None and int(end_col_used) > row_end_col:
        row_end_col = int(end_col_used)
    if cols:
        row_end_col = max(row_end_col, max(int(c) for c in cols))
    headers = []
    try:
        _h_sc, h_sr, _h_ec, _h_er = _lm_pp_range_address(header_row_range)
        hr = int(h_sr)
        headers = list(
            _lm_pp_header_titles(
                sheet, sheet.getCellRangeByPosition(0, hr, row_end_col, hr)
            )
        )
    except Exception:
        headers = []
        try:
            hr = int(_lm_pp_range_address(header_row_range)[1])
        except Exception:
            hr = max(0, int(sr) - 1)
        c = 0
        while c <= row_end_col:
            try:
                headers.append(cell_text(sheet.getCellByPosition(c, hr)))
            except Exception:
                headers.append("")
            c = c + 1
    try:
        data_arr = sheet.getCellRangeByPosition(0, sr, row_end_col, er).getDataArray()
    except Exception as err:
        _lm_log_postprocess(
            doc, sheet_name, "диапазон", fn_key, "ошибка", "getDataArray: %s" % err
        )
        return

    try:
        data_arr = [list(r) for r in (data_arr or [])]
    except Exception:
        pass

    from libre_macros_lambda_column_lib import make_rows_accessor, make_upd_rows_store

    rows_accessor = make_rows_accessor(data_arr, headers, start_col=0)

    changed_cells = 0
    skipped_cells = 0
    error_cells = 0
    changed_cols = 0

    for col in cols:
        col = int(col)
        skip_col, skip_reason = _lm_pp_replace_values_should_skip_column(
            sheet, col, sr, er
        )
        # не режем столбец целиком из-за non_text=skip — пустые строки решает row_filter
        if skip_col and non_text == "skip":
            skipped_cells += 1
            _lm_log_postprocess(
                doc,
                sheet_name,
                "диапазон",
                fn_key,
                "пропуск",
                "столбец %d: %s" % (col + 1, skip_reason),
            )
            continue
        out_rows = []
        col_changed = False
        stop_sheet = False
        upd_store = make_upd_rows_store()
        row_i = 0
        while row_i < len(data_arr):
            row_tuple = data_arr[row_i]
            try:
                v = row_tuple[col] if col < len(row_tuple) else ""
            except Exception:
                v = ""
            row_dict = build_row_dict(headers, row_tuple, row_i)
            try:
                text0 = cell_text_for_ops(
                    v, non_text=non_text, ops=ops, empty_cells=empty_cells
                )
            except NormalizeOpError:
                error_cells += 1
                if on_error == "stop":
                    stop_sheet = True
                    out_rows.append((v,))
                    break
                if on_error == "empty":
                    out_rows.append(("",))
                    col_changed = True
                    changed_cells += 1
                else:
                    out_rows.append((v,))
                    skipped_cells += 1
                row_i = row_i + 1
                continue
            if text0 is None:
                skipped_cells += 1
                out_rows.append((v,))
                row_i = row_i + 1
                continue
            try:
                new_text, err, _n = apply_ops(
                    text0,
                    ops,
                    key=crypto_key,
                    on_error=on_error,
                    cell_value=v,
                    row=row_dict,
                    row_filters=row_filters,
                    row_id=row_i,
                    rows_accessor=rows_accessor,
                    upd_store=upd_store,
                    compute_fns=compute_fns,
                )
            except NormalizeOpError as err:
                error_cells += 1
                _lm_log_postprocess(
                    doc, sheet_name, "диапазон", fn_key, "ошибка", str(err)
                )
                stop_sheet = True
                out_rows.append((v,))
                break
            if err is not None:
                error_cells += 1
            if new_text != text0:
                col_changed = True
                changed_cells += 1
                out_rows.append((new_text,))
                try:
                    if row_i < len(data_arr) and col < len(data_arr[row_i]):
                        data_arr[row_i][col] = new_text
                except Exception:
                    pass
            else:
                out_rows.append((v,))
            upd_store.commit(row_i)
            row_i = row_i + 1
        if stop_sheet:
            if col_changed:
                try:
                    rng = sheet.getCellRangeByPosition(col, sr, col, er)
                    need = er - sr + 1
                    while len(out_rows) < need:
                        out_rows.append(("",))
                    rng.setDataArray(tuple(out_rows[:need]))
                except Exception:
                    pass
            break
        if col_changed:
            try:
                rng = sheet.getCellRangeByPosition(col, sr, col, er)
                need = er - sr + 1
                while len(out_rows) < need:
                    out_rows.append(("",))
                rng.setDataArray(tuple(out_rows[:need]))
                changed_cols += 1
            except Exception as err:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "диапазон",
                    fn_key,
                    "ошибка",
                    "col=%s setDataArray: %s" % (col + 1, err),
                )

    _lm_log_postprocess(
        doc,
        sheet_name,
        "диапазон",
        fn_key,
        "ok",
        "ops=%d, столбцов=%d, изм.столбцов=%d, изм.ячеек=%d, пропуск=%d, ошибок=%d"
        % (
            len(ops),
            len(cols),
            changed_cols,
            changed_cells,
            skipped_cells,
            error_cells,
        ),
    )


def lm_pp_range_fill_down_empty(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Заполнение вниз: в указанных столбцах пустые ячейки получают значение из строки выше.

    Rest (колонка C): столбцы через , или ; — номер (1-based), буква или имя заголовка.
    Пример: «ФИО;Отдел;3».

    Карта: «заполнение_вниз».
    """
    sheet_name = sheet.Name if sheet is not None else ""
    spec_text = _lm_pp_join_extra_args(extra_args)
    if spec_text == "":
        _lm_log_postprocess(
            doc, sheet_name, "заполнение_вниз", "—", "пропуск", "пустая колонка C"
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    if er <= sr:
        _lm_log_postprocess(
            doc, sheet_name, "заполнение_вниз", spec_text, "пропуск", "нет строк данных"
        )
        return

    # json-first: param_decode без фолбэка на legacy при валидном JSON.
    col_tokens = None
    if _lm_pp_payload_is_json(spec_text):
        from libre_macros_param_codec import param_decode

        try:
            blocks = param_decode("заполнение_вниз", spec_text)
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "заполнение_вниз",
                spec_text,
                "ошибка",
                "json: %s" % err,
            )
            return
        out = []
        for b in blocks:
            if not isinstance(b, dict):
                continue
            target_sheet = str(b.get("sheet") or "").strip()
            if target_sheet and not _lm_pp_sheet_name_matches_target(
                sheet_name, [target_sheet]
            ):
                continue
            targets = b.get("sheets")
            if isinstance(targets, (list, tuple)) and len(targets) > 0:
                if not _lm_pp_sheet_name_matches_target(
                    sheet_name, [str(x) for x in targets if str(x).strip() != ""]
                ):
                    continue
            cols_raw = b.get("columns")
            if cols_raw is None:
                cols_raw = b.get("column")
            if isinstance(cols_raw, (list, tuple)):
                for t in cols_raw:
                    tt = str(t or "").strip()
                    if tt != "":
                        out.append(tt)
            elif isinstance(cols_raw, str):
                for t in _lm_pp_split_column_tokens(cols_raw):
                    if str(t).strip() != "":
                        out.append(str(t).strip())
        if len(out) == 0:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "заполнение_вниз",
                spec_text,
                "пропуск",
                "лист не в scope JSON",
            )
            return
        col_tokens = out

    if col_tokens is not None:
        cols = []
        for tok in col_tokens:
            col = _lm_pp_resolve_column_token(tok, sheet, header_row_range, sc, ec)
            if col is not None and col not in cols:
                cols.append(col)
        cols.sort()
    else:
        cols = _lm_pp_resolve_column_list(spec_text, sheet, header_row_range, sc, ec)
    if len(cols) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "заполнение_вниз",
            spec_text,
            "ошибка",
            "не найдены столбцы",
        )
        return

    def _cell_is_empty(cell):
        # Пустая ячейка в Calc: нет формулы, нет строки, Value==0.
        if cell is None:
            return True
        try:
            f = cell.Formula
            if f is not None and str(f).strip() != "":
                return False
        except Exception:
            pass
        try:
            s = cell.String
            if s is not None and str(s).strip() != "":
                return False
        except Exception:
            pass
        try:
            v = cell.Value
            if v not in (None, 0.0, 0):
                return False
        except Exception:
            pass
        return True

    # Новый алгоритм:
    # - идём по указанным колонкам;
    # - при первом пустом в строке смотрим «сколько пусто вправо»;
    # - копируем диапазон из строки выше и вставляем на пустые места целиком (со всеми атрибутами).
    filled = 0
    ci = 0
    while ci < len(cols):
        col = cols[ci]
        row = sr + 1
        while row <= er:
            try:
                cur = sheet.getCellByPosition(col, row)
            except Exception:
                cur = None
            if not _cell_is_empty(cur):
                row = row + 1
                continue

            # Считаем, сколько подряд пустых ячеек вправо, но не выходим за правую границу данных.
            end_col = col
            while end_col <= ec:
                try:
                    ccell = sheet.getCellByPosition(end_col, row)
                except Exception:
                    break
                if not _cell_is_empty(ccell):
                    break
                end_col = end_col + 1
            end_col = end_col - 1
            if end_col < col:
                row = row + 1
                continue

            # Копируем диапазон из предыдущей строки в текущую (с форматами/формулами).
            try:
                src_range = sheet.getCellRangeByPosition(col, row - 1, end_col, row - 1)
                dest_cell = sheet.getCellByPosition(col, row)
                sheet.copyRange(dest_cell.CellAddress, src_range.RangeAddress)
                filled = filled + (end_col - col + 1)
            except Exception:
                # Фолбэк: постолбцово (старое поведение).
                try:
                    c2 = col
                    while c2 <= end_col:
                        if _lm_pp_copy_cell_value_down(
                            sheet.getCellByPosition(c2, row - 1),
                            sheet.getCellByPosition(c2, row),
                        ):
                            filled = filled + 1
                        c2 = c2 + 1
                except Exception:
                    pass

            row = row + 1
        ci = ci + 1

    _lm_log_postprocess(
        doc,
        sheet_name,
        "заполнение_вниз",
        spec_text,
        "ok",
        "столбцов: %d, заполнено ячеек: %d" % (len(cols), filled),
    )


def lm_pp_range_fill_up_empty(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Заполнить вверх: пустые ячейки получают значение из ближайшей непустой снизу.

    JSON C: columns (columns_pick); skip_if_above_not_empty (default true);
    drop_trailing_empty (v1: false — верх без источника остаются пустыми).

    Карта: «заполнить_вверх» (алиасы fill_up, заполнение_вверх).
    """
    sheet_name = sheet.Name if sheet is not None else ""
    spec_text = _lm_pp_join_extra_args(extra_args)
    if spec_text == "":
        _lm_log_postprocess(
            doc, sheet_name, "заполнить_вверх", "—", "пропуск", "пустая колонка C"
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    if er <= sr:
        _lm_log_postprocess(
            doc, sheet_name, "заполнить_вверх", spec_text, "пропуск", "нет строк данных"
        )
        return

    col_tokens = None
    if _lm_pp_payload_is_json(spec_text):
        from libre_macros_param_codec import param_decode

        try:
            blocks = param_decode("заполнить_вверх", spec_text)
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "заполнить_вверх",
                spec_text,
                "ошибка",
                "json: %s" % err,
            )
            return
        out = []
        for b in blocks:
            if not isinstance(b, dict):
                continue
            target_sheet = str(b.get("sheet") or "").strip()
            if target_sheet and not _lm_pp_sheet_name_matches_target(
                sheet_name, [target_sheet]
            ):
                continue
            targets = b.get("sheets")
            if isinstance(targets, (list, tuple)) and len(targets) > 0:
                if not _lm_pp_sheet_name_matches_target(
                    sheet_name, [str(x) for x in targets if str(x).strip() != ""]
                ):
                    continue
            cols_raw = b.get("columns")
            if cols_raw is None:
                cols_raw = b.get("column")
            if isinstance(cols_raw, (list, tuple)):
                for t in cols_raw:
                    tt = str(t or "").strip()
                    if tt != "":
                        out.append(tt)
            elif isinstance(cols_raw, str):
                for t in _lm_pp_split_column_tokens(cols_raw):
                    if str(t).strip() != "":
                        out.append(str(t).strip())
        if len(out) == 0:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "заполнить_вверх",
                spec_text,
                "пропуск",
                "лист не в scope JSON",
            )
            return
        col_tokens = out

    if col_tokens is not None:
        cols = []
        for tok in col_tokens:
            col = _lm_pp_resolve_column_token(tok, sheet, header_row_range, sc, ec)
            if col is not None and col not in cols:
                cols.append(col)
        cols.sort()
    else:
        cols = _lm_pp_resolve_column_list(spec_text, sheet, header_row_range, sc, ec)
    if len(cols) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "заполнить_вверх",
            spec_text,
            "ошибка",
            "не найдены столбцы",
        )
        return

    def _cell_is_empty(cell):
        if cell is None:
            return True
        try:
            f = cell.Formula
            if f is not None and str(f).strip() != "":
                return False
        except Exception:
            pass
        try:
            s = cell.String
            if s is not None and str(s).strip() != "":
                return False
        except Exception:
            pass
        try:
            v = cell.Value
            if v not in (None, 0.0, 0):
                return False
        except Exception:
            pass
        return True

    # Снизу вверх: last_seen — ближайшее непустое снизу; пустые заполняем им.
    # Включаем sr (первая строка данных): иначе верхняя пустая ячейка не обрабатывается.
    filled = 0
    per_col = []
    ci = 0
    while ci < len(cols):
        col = cols[ci]
        col_filled = 0
        last_src_row = -1
        row = er
        while row >= sr:
            try:
                cur = sheet.getCellByPosition(col, row)
            except Exception:
                cur = None
            if not _cell_is_empty(cur):
                last_src_row = row
                row = row - 1
                continue
            # skip_if_above_not_empty=true (default): только пустые — уже здесь.
            if last_src_row < 0:
                row = row - 1
                continue
            try:
                src_range = sheet.getCellRangeByPosition(
                    col, last_src_row, col, last_src_row
                )
                dest_cell = sheet.getCellByPosition(col, row)
                sheet.copyRange(dest_cell.CellAddress, src_range.RangeAddress)
                col_filled = col_filled + 1
                filled = filled + 1
            except Exception:
                try:
                    src_cell = sheet.getCellByPosition(col, last_src_row)
                    if _lm_pp_copy_cell_value_down(src_cell, cur):
                        col_filled = col_filled + 1
                        filled = filled + 1
                except Exception:
                    pass
            row = row - 1
        per_col.append("%d:%d" % (col + 1, col_filled))
        ci = ci + 1

    note = "столбцов: %d, заполнено ячеек: %d" % (len(cols), filled)
    if per_col:
        note = note + "; по столбцам: " + ",".join(per_col)
    _lm_log_postprocess(
        doc,
        sheet_name,
        "заполнить_вверх",
        spec_text,
        "ok",
        note,
    )


def _lm_pp_parse_fill_down_formula_pairs(spec_text):
    """
    Разобрать параметр C для fill_down_calculate.
    Формат: столбец/формула пары через , или ;
    Примеры: ФИО/=A{row_id}&" "&B{row_id};3/=C{row_id-1}*2
    Возвращает список кортежей (column_token, formula_template).
    """
    text = str(spec_text or "").strip()
    if text == "":
        return []
    raw = re.split(r"[,;]", text)
    pairs = []
    for part in raw:
        part = part.strip()
        if part == "":
            continue
        if "/" not in part:
            continue
        idx = part.find("/")
        col_token = part[:idx].strip()
        formula = part[idx + 1:].strip()
        if col_token and formula:
            pairs.append((col_token, formula))
    return pairs


def _lm_pp_fill_down_calculate_resolve_block_pairs( block, sheet, header_row_range, sc, ec ):
    """
    Собрать задания для блока JSON.

    columns — столбцы, в которых ищутся пустые ячейки (пропуски).
    rules[].formula — формула для столбца (сопоставление rule.column по индексу).
    formula — общая формула для всех columns, если для столбца нет rule.
    Без columns — столбцы берутся из rules (обратная совместимость).
    """
    columns_in = block.get("columns") if isinstance(block, dict) else None
    has_columns = isinstance(columns_in, (list, tuple)) and len(columns_in) > 0
    rules = (block.get("rules") or []) if isinstance(block, dict) else []
    default_formula = str((block or {}).get("formula") or "").strip()

    formula_by_col = {}
    expand_by_col = {}
    ri = 0
    while ri < len(rules):
        r = rules[ri]
        ri += 1
        if not isinstance(r, dict):
            continue
        col_t = r.get("column")
        formula = str(r.get("formula") or "").strip()
        if col_t is None or formula == "":
            continue
        idx = _lm_pp_resolve_column_token(col_t, sheet, header_row_range, sc, ec)
        if idx is not None:
            formula_by_col[int(idx)] = formula
            if r.get("expand_to_right"):
                expand_by_col[int(idx)] = True

    pairs = []
    if has_columns:
        ti = 0
        while ti < len(columns_in):
            col_t = columns_in[ti]
            ti += 1
            idx = _lm_pp_resolve_column_token(col_t, sheet, header_row_range, sc, ec)
            if idx is None:
                pairs.append((col_t, None, None, False))
                continue
            formula = formula_by_col.get(int(idx))
            if formula is None and default_formula:
                formula = default_formula
            if formula is None and len(rules) == 1 and isinstance(rules[0], dict):
                formula = str(rules[0].get("formula") or "").strip()
            expand = bool(expand_by_col.get(int(idx), False))
            if formula is None and len(rules) == 1 and isinstance(rules[0], dict):
                expand = expand or bool(rules[0].get("expand_to_right"))
            if formula:
                pairs.append((col_t, formula, int(idx), expand))
            else:
                pairs.append((col_t, None, int(idx), expand))
        return pairs

    ri = 0
    while ri < len(rules):
        r = rules[ri]
        ri += 1
        if not isinstance(r, dict):
            continue
        col_t = r.get("column")
        formula = str(r.get("formula") or "").strip()
        if col_t is None or formula == "":
            continue
        idx = _lm_pp_resolve_column_token(col_t, sheet, header_row_range, sc, ec)
        expand = bool(r.get("expand_to_right"))
        pairs.append((col_t, formula, idx, expand))
    return pairs


def _lm_pp_format_fill_formula(template, row_1based):
    """
    Подставить номера строк в шаблон формулы.
    {row_id} / {row} → row_1based
    {row_id-1} / {row-1} → row_1based - 1
    """
    if template is None:
        return ""
    result = str(template)
    result = result.replace("{row_id-1}", str(row_1based - 1))
    result = result.replace("{row-1}", str(row_1based - 1))
    result = result.replace("{row_id}", str(row_1based))
    result = result.replace("{row}", str(row_1based))
    return result


def lm_pp_range_fill_down_calculate(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Заполнение вниз с вычислением: в указанных столбцах пустые ячейки получают
    значение по формуле/константе из параметра C.

    Алгоритм (отдельный от «применить_формулу»):
    - по каждому столбцу из columns (или из rules, если columns нет)
      найти пустые ячейки в диапазоне данных;
    - записать формулу (rules[].formula или общее formula блока);
    - протянуть fillAuto вниз по непрерывному пробегу пустых ячеек.

    JSON:
    - columns — столбцы с пропусками (1, A, имя заголовка);
    - rules — [{column, formula, expand_to_right?}, …];
    - formula — одна формула на все columns, если нет rule для столбца.

    Legacy-текст: пары столбец/формула через , или ;

    На время работы: undo.lock(), авторасчёт off → пересчёт → undo.clear().

    Карта: «заполнение_вниз_вычислить».
    """
    sheet_name = sheet.Name if sheet is not None else ""
    spec_text = _lm_pp_join_extra_args(extra_args)
    if spec_text == "":
        _lm_log_postprocess(
            doc, sheet_name, "заполнение_вниз_вычислить", "—", "пропуск", "пустая колонка C"
        )
        return

    sc, sr, ec, er = _lm_pp_range_address(data_range)
    if er <= sr:
        _lm_log_postprocess(
            doc, sheet_name, "заполнение_вниз_вычислить", spec_text, "пропуск", "нет строк данных"
        )
        return

    # json-first: если JSON распарсился — работаем строго по JSON (без фолбэков на legacy).
    pairs = []
    json_sheet_matched = False
    if _lm_pp_payload_is_json(spec_text):
        try:
            from libre_macros_param_codec import param_decode

            blocks = param_decode("заполнение_вниз_вычислить", spec_text)
        except Exception as err:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "заполнение_вниз_вычислить",
                spec_text,
                "ошибка",
                "json: %s" % err,
            )
            return
        for b in blocks or []:
            if not isinstance(b, dict):
                continue
            bs = str(b.get("sheet") or "").strip()
            if bs and not _lm_pp_sheet_name_matches_target(sheet_name, {bs}):
                continue
            targets = b.get("sheets")
            if isinstance(targets, (list, tuple)) and len(targets) > 0:
                if not _lm_pp_sheet_name_matches_target(
                    sheet_name, [str(x) for x in targets if str(x).strip() != ""]
                ):
                    continue
            json_sheet_matched = True
            pairs.extend(
                _lm_pp_fill_down_calculate_resolve_block_pairs(
                    b, sheet, header_row_range, sc, ec
                )
            )
        if not json_sheet_matched:
            _lm_log_postprocess(
                doc,
                sheet_name,
                "заполнение_вниз_вычислить",
                spec_text,
                "пропуск",
                "лист не в scope JSON",
            )
            return
    else:
        legacy = _lm_pp_parse_fill_down_formula_pairs(spec_text)
        li = 0
        while li < len(legacy):
            col_token, formula_template = legacy[li]
            li += 1
            pairs.append((col_token, formula_template, None, False))
    if len(pairs) == 0:
        _lm_log_postprocess(
            doc,
            sheet_name,
            "заполнение_вниз_вычислить",
            spec_text,
            "ошибка",
            "нет столбцов columns или rules",
        )
        return

    filled = 0
    errors = 0
    undo_state = _lm_pp_undo_begin_batch(doc)
    auto_was = _lm_pp_autocalc_suspend(doc)
    try:
        ci = 0
        while ci < len(pairs):
            col_token, formula_template, col_resolved, expand_to_right = pairs[ci]
            ci += 1
            if formula_template is None or str(formula_template).strip() == "":
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "заполнение_вниз_вычислить",
                    spec_text,
                    "ошибка",
                    "нет формулы для столбца: %s" % col_token,
                )
                errors += 1
                continue
            if col_resolved is not None:
                col = int(col_resolved)
            else:
                col = _lm_pp_resolve_column_token(
                    col_token, sheet, header_row_range, sc, ec
                )
            if col is None:
                _lm_log_postprocess(
                    doc,
                    sheet_name,
                    "заполнение_вниз_вычислить",
                    spec_text,
                    "ошибка",
                    "столбец не найден: %s" % col_token,
                )
                errors += 1
                continue

            row = sr + 1
            run_filled, run_errors = _lm_pp_fill_down_calculate_empty_runs(
                doc,
                sheet,
                col,
                row,
                er,
                formula_template,
                sheet_name,
                expand_to_right=expand_to_right,
                end_col=ec,
            )
            filled += run_filled
            errors += run_errors
    finally:
        _lm_pp_autocalc_restore(doc, auto_was)
        _lm_pp_calculate_doc(doc)
        _lm_pp_undo_end_batch(
            doc, undo_state, "заполнение_вниз_вычислить", clear=True
        )

    _lm_log_postprocess(
        doc,
        sheet_name,
        "заполнение_вниз_вычислить",
        spec_text,
        "ok" if errors == 0 else "ок с ошибками",
        "пар: %d, заполнено: %d, ошибок: %d" % (len(pairs), filled, errors),
    )


# =============================================================================
# ВПР (Smart VLOOKUP / LEFT OUTER JOIN) — TZ_VLOOKUP2.md §7
# =============================================================================

_LM_VLOOKUP_KEY_EMPTY = "__ПУСТО__"
_LM_VLOOKUP_KEYMAP_CHUNK = 200
_LM_VLOOKUP_UI_YIELD_EVERY = 50
_LM_VLOOKUP_TRIM_EDGE_CHARS = frozenset(
    u" \t\n\r\v\f\u00a0\u1680\u2000\u2001\u2002\u2003\u2004\u2005\u2006"
    u"\u2007\u2008\u2009\u200a\u2028\u2029\u202f\u205f\u3000\ufeff"
)


def _lm_vlookup_trim_key_part(value):
    """Обрезка краёв ключа: пробелы, «похожие» символы и переводы строк (§5.2 trim_keys)."""
    if value is None:
        return _LM_VLOOKUP_KEY_EMPTY
    if isinstance(value, float):
        try:
            if value == int(value):
                return str(int(value))
        except (TypeError, ValueError, OverflowError):
            pass
    if isinstance(value, int):
        return str(value)
    s = str(value)
    if s == "":
        return _LM_VLOOKUP_KEY_EMPTY
    i = 0
    n = len(s)
    while i < n and s[i] in _LM_VLOOKUP_TRIM_EDGE_CHARS:
        i = i + 1
    j = n
    while j > i and s[j - 1] in _LM_VLOOKUP_TRIM_EDGE_CHARS:
        j = j - 1
    s = s[i:j]
    return _LM_VLOOKUP_KEY_EMPTY if s == "" else s


def _lm_vlookup_key_part(value, trim=False):
    """Нормализация части составного ключа (пусто → __ПУСТО__)."""
    if trim:
        return _lm_vlookup_trim_key_part(value).casefold()
    if value is None:
        return _LM_VLOOKUP_KEY_EMPTY
    if isinstance(value, str):
        s = value.strip()
        return _LM_VLOOKUP_KEY_EMPTY if s == "" else s.casefold()
    if isinstance(value, float):
        try:
            if value == int(value):
                return str(int(value))
        except (TypeError, ValueError, OverflowError):
            pass
    if isinstance(value, int):
        return str(value)
    # Для любых прочих типов: строковое представление + casefold для регистронезависимости.
    try:
        return str(value).casefold()
    except Exception:
        return str(value)


def _lm_vlookup_composite_key(values, trim=False):
    return "|".join(_lm_vlookup_key_part(v, trim=trim) for v in values)


_LM_REPLACE_FIELD_RE = None


def _lm_replace_fields(template, field_values):
    """Подстановка [ИмяКолонки] в шаблоне значениями из dict (заголовок → значение)."""
    global _LM_REPLACE_FIELD_RE
    import re

    if _LM_REPLACE_FIELD_RE is None:
        _LM_REPLACE_FIELD_RE = re.compile(r"\[([^\]]+)\]")
    s = unicode(template or u"")
    if s == u"":
        return u""
    values = field_values or {}
    lookup = {}
    for key, val in values.items():
        k = unicode(key or u"").strip()
        if k == u"":
            continue
        lookup[k.casefold()] = val

    def _repl(match):
        name = unicode(match.group(1) or u"").strip()
        if name in values:
            v = values.get(name)
            return unicode(v if v is not None else u"")
        cf = name.casefold()
        if cf in lookup:
            v = lookup.get(cf)
            return unicode(v if v is not None else u"")
        return match.group(0)

    return _LM_REPLACE_FIELD_RE.sub(_repl, s)


_LM_SHEET_NAME_FORBIDDEN = u":[]?*/\\"
_LM_WIN_PATH_FORBIDDEN = u'<>:"/\\|?*'
_LM_WIN_RESERVED = frozenset(
    [
        "CON",
        "PRN",
        "AUX",
        "NUL",
    ]
    + ["COM%d" % i for i in range(1, 10)]
    + ["LPT%d" % i for i in range(1, 10)]
)


def _lm_normalize_path_component(name, kind="path"):
    """
    Очистка компоненты пути или имени листа от недопустимых символов.
    kind: "path" | "sheet"
    """
    import os

    s = unicode(name or u"").strip()
    if kind == "sheet":
        out = []
        for ch in s:
            out.append(u"_" if ch in _LM_SHEET_NAME_FORBIDDEN else ch)
        s = u"".join(out).strip(u".'")
        if s == u"":
            s = u"Sheet"
        if len(s) > 31:
            s = s[:31]
        return s
    if os.name == "nt":
        out = []
        for ch in s:
            out.append(u"_" if ch in _LM_WIN_PATH_FORBIDDEN else ch)
        s = u"".join(out).strip(u". ")
        base = s
        if u"." in base:
            base = base.split(u".", 1)[0]
        if base.casefold() in _LM_WIN_RESERVED:
            s = u"_" + s
    else:
        s = s.replace(u"/", u"_").strip()
    return s if s != u"" else u"_"


def _lm_sanitize_sheet_name(base):
    """Имя листа Calc/Excel (≤31 символ)."""
    return _lm_normalize_path_component(base, kind="sheet")


def _lm_vlookup_sheet_used_area(sheet):
    """Значащая область листа (0-based): sc, sr, ec, er."""
    if sheet is None:
        return 0, 0, -1, -1
    try:
        cursor = sheet.createCursor()
        if cursor is None:
            return 0, 0, -1, -1
        cursor.gotoStartOfUsedArea(False)
        start_addr = cursor.getRangeAddress()
        cursor.gotoEndOfUsedArea(False)
        end_addr = cursor.getRangeAddress()
        return (
            start_addr.StartColumn,
            start_addr.StartRow,
            end_addr.EndColumn,
            end_addr.EndRow,
        )
    except Exception:
        return 0, 0, -1, -1


def _lm_vlookup_resolve_sheet(doc, sheet_ref):
    """Имя или 1-based индекс листа → XSpreadsheet."""
    if doc is None or sheet_ref is None:
        return None
    ref = str(sheet_ref).strip()
    if ref == "":
        return None
    sheets = doc.getSheets()
    if ref.isdigit():
        idx = int(ref) - 1
        if 0 <= idx < sheets.getCount():
            return sheets.getByIndex(idx)
        return None
    try:
        if sheets.hasByName(ref):
            return sheets.getByName(ref)
    except Exception:
        pass
    key = ref.casefold()
    matches = []
    i = 0
    while i < sheets.getCount():
        try:
            sh = sheets.getByIndex(i)
            name = str(sh.Name).strip()
            if name.casefold() == key:
                return sh
            m = re.match(r"^(\d+)_(.+)$", name)
            if m is not None and m.group(2).strip().casefold() == key:
                matches.append((int(m.group(1)), sh))
        except Exception:
            pass
        i = i + 1
    if len(matches) == 0:
        return None
    matches.sort(key=lambda item: item[0])
    return matches[0][1]


def _lm_vlookup_resolve_bounds( sheet, start_row_1, start_col_1, header_row_1, range_a1=None, end_row_1=None, end_col_1=None ):
    """
    Границы таблицы на листе (0-based).

    Если задан range_a1 — границы из A1 (заголовок = первая строка диапазона).
    Иначе: пустые параметры → used area; заголовок = первая строка области
    или Начало_Строка, если Строка_Заголовков не задана.
    """
    range_txt = str(range_a1 or "").strip()
    if range_txt != "":
        try:
            from libre_macros_copy_ranges_lib import parse_source_range_a1
        except Exception:
            parse_source_range_a1 = None
        if parse_source_range_a1 is None:
            raise ValueError("не удалось разобрать range=%s" % range_txt)
        rect, err = parse_source_range_a1(sheet, range_txt)
        if err:
            raise ValueError(err)
        sc, sr, ec, er = rect
        if er < sr or ec < sc:
            raise ValueError("пустой диапазон «%s»" % range_txt)
        return {
            "start_col": sc,
            "header_row": sr,
            "data_start": sr + 1,
            "end_col": ec,
            "end_row": er,
        }

    ua_sc, ua_sr, ua_ec, ua_er = _lm_vlookup_sheet_used_area(sheet)
    if ua_er < ua_sr:
        raise ValueError("лист «%s» пуст" % (sheet.Name if sheet is not None else "?"))

    if start_col_1 is None:
        start_col = ua_sc
    else:
        start_col = int(start_col_1) - 1
        if start_col < 0:
            raise ValueError("некорректный начальный столбец: %s" % start_col_1)

    if header_row_1 is not None:
        header_row = int(header_row_1) - 1
    elif start_row_1 is not None:
        header_row = int(start_row_1) - 1
    else:
        header_row = ua_sr

    if start_row_1 is not None:
        data_start = int(start_row_1) - 1
    else:
        data_start = header_row + 1

    if data_start < 0:
        raise ValueError("некорректная начальная строка данных")

    if end_col_1 is not None:
        end_col = int(end_col_1) - 1
    else:
        end_col = ua_ec
    if end_row_1 is not None:
        end_row = int(end_row_1) - 1
    else:
        end_row = ua_er

    return {
        "start_col": start_col,
        "header_row": header_row,
        "data_start": data_start,
        "end_col": end_col,
        "end_row": end_row,
    }


def _lm_vlookup_header_range(sheet, bounds):
    ec = bounds["end_col"]
    if ec < bounds["start_col"]:
        ec = bounds["start_col"]
    hr = bounds["header_row"]
    return sheet.getCellRangeByPosition(bounds["start_col"], hr, ec, hr)


def _lm_vlookup_resolve_key_columns(sheet, bounds, key_tokens):
    """Список пар (токен, 0-based индекс столбца) для ключей."""
    hdr = _lm_vlookup_header_range(sheet, bounds)
    sc = bounds["start_col"]
    ec = bounds["end_col"]
    cols = []
    i = 0
    while i < len(key_tokens):
        token = key_tokens[i]
        col = _lm_pp_resolve_column_token(token, sheet, hdr, sc, ec)
        if col is None:
            raise ValueError("ключ «%s» не найден на листе «%s»" % (token, sheet.Name))
        cols.append(col)
        i = i + 1
    return cols


def _lm_vlookup_resolve_extract_columns(sheet, bounds, col_tokens):
    """
    Список (0-based индекс, имя на левом листе) для столбцов извлечения.

    Токены: «ФИО» / «'ФИО'» или «ФИО -> Новое» / «'ФИО' = Новое»
    (см. parse_vlookup_extract_alias).
    """
    hdr = _lm_vlookup_header_range(sheet, bounds)
    sc = bounds["start_col"]
    ec = bounds["end_col"]
    hr = bounds["header_row"]
    try:
        from libre_macros_param_codec import parse_vlookup_extract_alias
    except Exception:
        parse_vlookup_extract_alias = None
    out = []
    i = 0
    while i < len(col_tokens):
        token = col_tokens[i]
        src_token = token
        as_name = None
        if parse_vlookup_extract_alias is not None:
            src_token, as_name = parse_vlookup_extract_alias(token)
        col = _lm_pp_resolve_column_token(src_token, sheet, hdr, sc, ec)
        if col is None:
            raise ValueError(
                "столбец «%s» не найден на листе «%s»" % (src_token, sheet.Name)
            )
        if as_name is not None and str(as_name).strip() != "":
            title = str(as_name).strip()
        else:
            title = cell_text(sheet.getCellByPosition(col, hr)).strip()
            if title == "":
                title = "Столбец_%d" % (col + 1)
        out.append((col, title))
        i = i + 1
    return out


def _lm_vlookup_row_key_values(sheet, row, key_cols):
    vals = []
    i = 0
    while i < len(key_cols):
        vals.append(cell_text(sheet.getCellByPosition(key_cols[i], row)))
        i = i + 1
    return vals


def _lm_vlookup_not_found_fill_is_empty_cell(fill_text):
    """Заполнитель «не найдено» — оставить ячейку пустой (_ПУСТО_, _EMPTY_, __ПУСТО__)."""
    s = str(fill_text or "").strip()
    if s == "":
        return False
    key = s.casefold().replace(" ", "")
    return key in ("_пусто_", "__пусто__", "_empty_")


def _lm_vlookup_is_na_fill(text):
    """True, если заполнитель означает ошибку «нет данных» (#N/A / #Н/Д / =NA())."""
    s = str(text or "").strip()
    if s == "":
        return False
    key = s.casefold().replace(" ", "")
    return key in (
        "#н/д", "#n/a", "#na", "н/д", "n/a", "na()", "=na()", "na",
    )


def _lm_vlookup_apply_fill_value(cell, fill_text):
    """Записать в ячейку значение-заполнитель (строка, число или =NA())."""
    if cell is None:
        return
    s = str(fill_text if fill_text is not None else "").strip()
    if _lm_vlookup_not_found_fill_is_empty_cell(s):
        try:
            cell.setString("")
        except Exception:
            try:
                cell.String = ""
            except Exception:
                pass
        return
    if s == "":
        try:
            cell.String = ""
        except Exception:
            pass
        return
    if _lm_vlookup_is_na_fill(s):
        try:
            cell.setFormula("=NA()")
            return
        except Exception:
            pass
    try:
        if s.isdigit() or (s.startswith("-") and s[1:].isdigit()):
            cell.Value = float(int(s))
            return
    except Exception:
        pass
    try:
        if "." in s or "," in s:
            cell.Value = float(s.replace(",", "."))
            return
    except Exception:
        pass
    try:
        cell.String = s
    except Exception:
        pass


def _lm_vlookup_set_cell_value(cell, value):
    if cell is None:
        return
    if value is None or (isinstance(value, str) and value.strip() == ""):
        try:
            cell.String = ""
        except Exception:
            pass
        return
    try:
        if isinstance(value, (int, float)):
            cell.Value = float(value)
            return
    except Exception:
        pass
    try:
        cell.String = str(value)
    except Exception:
        pass


def _lm_vlookup_highlight_style_active(highlight_style):
    if highlight_style is None:
        return False
    return (
        highlight_style.get("bg") is not None
        or highlight_style.get("fg") is not None
    )


def _lm_vlookup_apply_cell_highlight(cell, highlight_style):
    if cell is None or not _lm_vlookup_highlight_style_active(highlight_style):
        return
    bg = highlight_style.get("bg")
    fg = highlight_style.get("fg")
    try:
        if bg is not None:
            cell.CellBackColor = bg
    except Exception:
        pass
    try:
        if fg is not None:
            cell.CharColor = fg
    except Exception:
        pass


def _lm_vlookup_parse_highlight_style(color_text):
    """Заливка/шрифт подтянутых ячеек; без цвета в F — None (без подсветки по умолчанию)."""
    raw = str(color_text or "").strip()
    if raw == "":
        return None
    if ":" in raw and not raw.lstrip().startswith("#"):
        _lbl, _sep, raw = raw.partition(":")
        raw = raw.strip()
    if raw == "":
        return None
    style = _lm_pp_zebra_parse_stripe_spec(raw)
    if style.get("bg") is None and style.get("fg") is None:
        return None
    return style


def _lm_vlookup_parse_highlight_color(color_text):
    style = _lm_vlookup_parse_highlight_style(color_text)
    if style is None:
        return None
    return style.get("bg")


def _lm_vlookup_build_keymap(right_sheet, bounds, right_key_cols, extract_specs, trim_keys=False):
    """
    KeyMap: составной ключ → список dict {base_title: value}.
    Правая таблица читается чанками (_LM_VLOOKUP_KEYMAP_CHUNK строк).
    """
    keymap = {}
    sc = bounds["start_col"]
    ec = bounds["end_col"]
    row = bounds["data_start"]
    end_row = bounds["end_row"]
    if row > end_row:
        return keymap

    while row <= end_row:
        chunk_end = min(row + _LM_VLOOKUP_KEYMAP_CHUNK - 1, end_row)
        try:
            chunk_range = right_sheet.getCellRangeByPosition(sc, row, ec, chunk_end)
            data = chunk_range.getDataArray()
        except Exception:
            data = ()

        ri = 0
        while ri < len(data):
            row_vals = data[ri]
            if not isinstance(row_vals, tuple):
                row_vals = (row_vals,)
            key_parts = []
            ki = 0
            while ki < len(right_key_cols):
                kc = right_key_cols[ki]
                rel = kc - sc
                if 0 <= rel < len(row_vals):
                    key_parts.append(row_vals[rel])
                else:
                    key_parts.append(None)
                ki = ki + 1
            key = _lm_vlookup_composite_key(key_parts, trim=trim_keys)
            record = {}
            ei = 0
            while ei < len(extract_specs):
                ec_idx, base_title = extract_specs[ei]
                rel = ec_idx - sc
                if 0 <= rel < len(row_vals):
                    record[base_title] = row_vals[rel]
                else:
                    record[base_title] = None
                ei = ei + 1
            if key not in keymap:
                keymap[key] = []
            keymap[key].append(record)
            ri = ri + 1

        data = None
        row = chunk_end + 1

    return keymap


def _lm_vlookup_title_looks_like_date(title):
    t = str(title or "").casefold()
    if t == "":
        return False
    for m in _LM_PP_VLOOKUP_DATE_TITLE_MARKERS:
        if m in t:
            return True
    return False


def _lm_pp_number_format_type(doc, fmt_key):
    if doc is None or fmt_key is None:
        return None
    try:
        return int(doc.NumberFormats.getType(int(fmt_key)))
    except Exception:
        return None


def _lm_pp_number_format_is_date_time(doc, fmt_key):
    t = _lm_pp_number_format_type(doc, fmt_key)
    if t is None:
        return False
    return t in (
        _LM_PP_NF_TYPE_DATE,
        _LM_PP_NF_TYPE_TIME,
        _LM_PP_NF_TYPE_DATETIME,
    )


def _lm_pp_number_format_is_general(doc, fmt_key):
    if fmt_key is None:
        return True
    try:
        k = int(fmt_key)
    except (TypeError, ValueError):
        return True
    if k <= 0:
        return True
    if doc is None:
        return False
    try:
        info = doc.NumberFormats.getByKey(k)
        fs = str(getattr(info, "FormatString", "") or "").strip().casefold()
    except Exception:
        return False
    if fs == "" or fs in ("general", "стандартный", "стандарт", "standard"):
        return True
    return False


def _lm_vlookup_default_date_format_string():
    """Целевой формат даты как при сборе (глобальные настройки → DD.MM.YYYY)."""
    try:
        from libre_macros_global_settings_lib import (
            DEFAULT_DATE_FORMAT,
            load_global_settings,
        )

        gs = load_global_settings()
        raw = str((gs or {}).get("default_date_format") or "").strip()
        if raw:
            return raw
        return str(DEFAULT_DATE_FORMAT or _LM_PP_DATE_FORMAT)
    except Exception:
        return _LM_PP_DATE_FORMAT


def _lm_vlookup_sample_number_format(sheet, col, top_row, bottom_row):
    """Ключ NumberFormat из первой непустой ячейки столбца в диапазоне строк."""
    if sheet is None or col is None or int(col) < 0:
        return None
    try:
        r0 = int(top_row)
        r1 = int(bottom_row)
    except (TypeError, ValueError):
        return None
    if r1 < r0:
        return None
    r = r0
    while r <= r1 and r <= r0 + 32:
        try:
            cell = sheet.getCellByPosition(int(col), r)
            txt = cell_text(cell).strip()
            if txt == "" and not (
                isinstance(getattr(cell, "Value", None), (int, float))
                and float(cell.Value) != 0
            ):
                r = r + 1
                continue
            return int(cell.NumberFormat)
        except Exception:
            pass
        r = r + 1
    try:
        return int(sheet.getCellByPosition(int(col), r0).NumberFormat)
    except Exception:
        return None


def _lm_vlookup_apply_collect_defaults_to_column( doc, left_sheet, new_col, header_row, data_end, col_title, right_sheet=None, right_col=None, right_bounds=None, template_col=-1):
    """
    Оформление нового столбца ВПР как при сборе:
      — заголовок: серый / жирный / PT Sans;
      — данные: NumberFormat с правого столбца (или дата по маркеру заголовка);
      — шрифт данных: PT Sans + размер по умолчанию.
    """
    if left_sheet is None or new_col is None or int(new_col) < 0:
        return
    hr = int(header_row or 0)
    try:
        data_end = int(data_end)
    except (TypeError, ValueError):
        data_end = hr
    data_top = hr + 1
    # Заголовок — как у листа результата
    try:
        hcell = left_sheet.getCellByPosition(int(new_col), hr)
        hcell.IsCellBackgroundTransparent = False
        hcell.CellBackColor = int(_LM_FMT_HEADER_BG)
        hcell.CharFontName = _LM_PP_FONT_NAME
        hcell.CharWeight = _LM_FMT_CHAR_WEIGHT_BOLD
        hcell.CharHeight = float(_LM_PP_DEFAULT_FONT_SIZE)
        _lm_pp_set_cell_alignment(hcell, HORI_CENTER, "center")
    except Exception:
        pass

    fmt_key = None
    if right_sheet is not None and right_col is not None and int(right_col) >= 0:
        rb = right_bounds or {}
        src_hr = int(rb.get("header_row", 0) or 0)
        src_ds = int(rb.get("data_start", src_hr + 1) or (src_hr + 1))
        src_er = int(rb.get("end_row", src_ds) or src_ds)
        fmt_key = _lm_vlookup_sample_number_format(
            right_sheet, right_col, src_ds, src_er
        )
        if fmt_key is None:
            fmt_key = _lm_vlookup_sample_number_format(
                right_sheet, right_col, src_hr, src_hr
            )

    want_date = _lm_vlookup_title_looks_like_date(col_title)
    if want_date or _lm_pp_number_format_is_date_time(doc, fmt_key):
        if fmt_key is None or _lm_pp_number_format_is_general(doc, fmt_key):
            fmt_key = _lm_pp_number_format_key(
                doc, _lm_vlookup_default_date_format_string()
            )
        elif want_date and not _lm_pp_number_format_is_date_time(doc, fmt_key):
            # Заголовок «дата*», но справа не date-type — как при сборе.
            fmt_key = _lm_pp_number_format_key(
                doc, _lm_vlookup_default_date_format_string()
            )

    if fmt_key is None and template_col is not None and int(template_col) >= 0:
        fmt_key = _lm_vlookup_sample_number_format(
            left_sheet, template_col, data_top, max(data_top, data_end)
        )

    if data_end >= data_top:
        try:
            data_range = left_sheet.getCellRangeByPosition(
                int(new_col), data_top, int(new_col), data_end
            )
        except Exception:
            data_range = None
        if data_range is not None:
            if fmt_key is not None:
                try:
                    data_range.NumberFormat = int(fmt_key)
                except Exception:
                    pass
            try:
                data_range.CharFontName = _LM_PP_FONT_NAME
            except Exception:
                pass
            try:
                data_range.CharHeight = float(_LM_PP_DEFAULT_FONT_SIZE)
            except Exception:
                pass
            try:
                data_range.CharWeight = 100
            except Exception:
                pass


def _lm_vlookup_seed_key_col_map_from_headers(sheet, bounds, extract_specs, run_suffix, key_col_map):
    """Заполнить key_col_map уже существующими столбцами выхода (full join, append-pass)."""
    if sheet is None or not extract_specs:
        return
    hr = int(bounds.get("header_row", 0) or 0)
    sc = int(bounds.get("start_col", 0) or 0)
    ec = int(bounds.get("end_col", 0) or 0)
    ei = 0
    while ei < len(extract_specs):
        _ec_idx, base_title = extract_specs[ei]
        col_title = base_title + run_suffix
        if col_title in key_col_map:
            ei = ei + 1
            continue
        candidates = [col_title]
        if base_title != col_title:
            candidates.append(base_title)
        found_col = None
        ci = 0
        while ci < len(candidates) and found_col is None:
            cand = candidates[ci]
            hkey = lm_identity_key(cand)
            c = sc
            while c <= ec:
                try:
                    htxt = cell_text(sheet.getCellByPosition(c, hr)).strip()
                except Exception:
                    htxt = ""
                if htxt != "" and lm_identity_key(htxt) == hkey:
                    found_col = c
                    break
                c = c + 1
            ci = ci + 1
        if found_col is not None:
            key_col_map[col_title] = found_col
        ei = ei + 1


def _lm_vlookup_ensure_output_column( sheet, key_col_map, col_title, bounds, template_col, highlight_style=None, doc=None, right_sheet=None, right_col=None, right_bounds=None):
    """Добавить столбец справа от таблицы, если его ещё нет в KeyColMap."""
    if col_title in key_col_map:
        return key_col_map[col_title]

    # Коллизии заголовков при последовательных ВПР:
    # если col_title уже есть на левом листе, то новый добавляемый столбец
    # должен получить уникальное имя (col, col_2, col_3, ...),
    # а существующий заголовок(и) не трогаем.
    hr = bounds.get("header_row", 0)
    try:
        # Важно: используем общий аллокатор, чтобы корректно учитывать
        # все уже существующие заголовки на строке hr (а не только bounds).
        header_to_write = _lm_pp_allocate_new_header_name(sheet, hr, col_title)
    except Exception:
        header_to_write = col_title

    new_col = bounds["end_col"] + 1
    if template_col < bounds["start_col"]:
        template_col = bounds["end_col"]
    sheet.getColumns().insertByIndex(new_col, 1)
    data_end = max(bounds["end_row"], bounds["data_start"])
    header_cell = sheet.getCellByPosition(new_col, hr)
    header_cell.String = header_to_write
    # Оформление как при сборе: заголовок + NumberFormat/шрифт данных (с правого столбца).
    _lm_vlookup_apply_collect_defaults_to_column(
        doc,
        sheet,
        new_col,
        hr,
        data_end,
        col_title,
        right_sheet=right_sheet,
        right_col=right_col,
        right_bounds=right_bounds,
        template_col=template_col,
    )
    _lm_vlookup_apply_cell_highlight(header_cell, highlight_style)
    bounds["end_col"] = new_col
    key_col_map[col_title] = new_col
    # Автоширина по диапазону данных с учётом заголовка
    try:
        _lm_pp_autofit_columns(None, sheet, [new_col], sr=hr, er=data_end)
    except Exception:
        pass
    return new_col


def _lm_vlookup_ensure_all_output_columns( sheet, key_col_map, extract_specs, run_suffix, bounds, template_col, highlight_style=None, doc=None, right_sheet=None, right_bounds=None):
    """Создать все столбцы извлечения до обхода строк."""
    ei = 0
    while ei < len(extract_specs):
        ec_idx, base_title = extract_specs[ei]
        col_title = base_title + run_suffix
        template_col = _lm_vlookup_ensure_output_column(
            sheet,
            key_col_map,
            col_title,
            bounds,
            template_col,
            highlight_style,
            doc=doc,
            right_sheet=right_sheet,
            right_col=ec_idx,
            right_bounds=right_bounds,
        )
        ei = ei + 1
    return template_col


def _lm_vlookup_copy_left_row(sheet, src_row, dest_row, start_col, end_col):
    """Скопировать значения левой строки (getDataArray → setDataArray)."""
    if dest_row == src_row:
        return
    try:
        src_range = sheet.getCellRangeByPosition(start_col, src_row, end_col, src_row)
        row_data = src_range.getDataArray()
        if row_data is None or len(row_data) == 0:
            return
        dest_range = sheet.getCellRangeByPosition(start_col, dest_row, end_col, dest_row)
        dest_range.setDataArray(row_data)
    except Exception:
        pass


def _lm_vlookup_multiline_part(value):
    """Фрагмент значения для склейки через перенос строки."""
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        try:
            fv = float(value)
            if fv == int(fv):
                return str(int(fv))
        except Exception:
            pass
        return str(value)
    return str(value)


def _lm_vlookup_set_cell_multiline_value(cell, text):
    if cell is None:
        return
    try:
        cell.String = str(text if text is not None else "")
    except Exception:
        pass
    try:
        cell.setPropertyValue("IsTextWrapped", True)
    except Exception:
        pass


def _lm_vlookup_fill_multi_match_row( sheet, row, key_col_map, extract_specs, run_suffix, matches, highlight_style ):
    """Несколько совпадений справа → одна ячейка на столбец, значения через \\n."""
    ei = 0
    while ei < len(extract_specs):
        _ec_idx, base_title = extract_specs[ei]
        col_title = base_title + run_suffix
        out_col = key_col_map.get(col_title)
        if out_col is None:
            ei = ei + 1
            continue
        parts = []
        mi = 0
        while mi < len(matches):
            parts.append(_lm_vlookup_multiline_part(matches[mi].get(base_title)))
            mi = mi + 1
        cell = sheet.getCellByPosition(out_col, row)
        _lm_vlookup_set_cell_multiline_value(cell, "\n".join(parts))
        _lm_vlookup_apply_cell_highlight(cell, highlight_style)
        ei = ei + 1


def _lm_vlookup_fill_extract_row( sheet, row, key_col_map, extract_specs, run_suffix, match_record, highlight_style ):
    ei = 0
    while ei < len(extract_specs):
        _ec_idx, base_title = extract_specs[ei]
        col_title = base_title + run_suffix
        out_col = key_col_map.get(col_title)
        if out_col is None:
            ei = ei + 1
            continue
        cell = sheet.getCellByPosition(out_col, row)
        _lm_vlookup_set_cell_value(cell, match_record.get(base_title))
        _lm_vlookup_apply_cell_highlight(cell, highlight_style)
        ei = ei + 1


def _lm_vlookup_fill_not_found_row( sheet, row, key_col_map, extract_specs, run_suffix, fill_text, highlight_style ):
    ei = 0
    while ei < len(extract_specs):
        _ec_idx, base_title = extract_specs[ei]
        col_title = base_title + run_suffix
        out_col = key_col_map.get(col_title)
        if out_col is None:
            ei = ei + 1
            continue
        cell = sheet.getCellByPosition(out_col, row)
        _lm_vlookup_apply_fill_value(cell, fill_text)
        _lm_vlookup_apply_cell_highlight(cell, highlight_style)
        ei = ei + 1


def _lm_vlookup_normalize_join_type(raw, default="left"):
    s = str(raw or "").strip().casefold()
    if s in ("", "left", "левое", "left_join", "left outer", "left_outer"):
        return "left"
    if s in (
        "inner",
        "внутреннее",
        "inner_join",
        "inner join",
        "внутреннее_соединение",
        "внутреннее соединение",
    ):
        return "inner"
    if s in (
        "full",
        "полное",
        "full_outer",
        "full outer",
        "полное_соединение",
        "полное соединение",
    ):
        return "full"
    return default


def _lm_vlookup_normalize_multi_match(raw, default="all"):
    s = str(raw or "").strip().casefold()
    if s in ("", "all", "все", "*", "all_rows", "все_совпадения"):
        return "all"
    if s in ("first", "первый", "первая", "первое", "first_row", "1"):
        return "first"
    if s in ("last", "последний", "последняя", "последнее", "last_row", "-1"):
        return "last"
    return default


def _lm_vlookup_select_matches(matches, mode):
    if not matches:
        return matches
    m = _lm_vlookup_normalize_multi_match(mode, default="all")
    if m == "first":
        return [matches[0]]
    if m == "last":
        return [matches[len(matches) - 1]]
    return matches


def lm_vlookup_join_sheets(doc, left_sheet, spec):
    """
    LEFT / INNER / FULL JOIN двух листов книги результата.
    При join_type=full — два прохода + union строк только справа.
    При join_type=inner — как left, но строки без совпадения удаляются.

    Параметры:
        doc — книга результата.
        left_sheet — левый лист (XSpreadsheet).
        spec — dict из parse_vlookup_spec + run_suffix (_R1, _R2, …).

    Возвращает:
        dict: rows_added, rows_processed, columns_added, keymap_size.
    """
    if doc is None or left_sheet is None or spec is None:
        raise ValueError("ВПР: не заданы doc, лист или spec")
    if not spec.get("_vlookup_skip_full_dispatch"):
        jt = _lm_vlookup_normalize_join_type(spec.get("join_type"), default="left")
        if jt == "full":
            return _lm_vlookup_join_full(doc, left_sheet, spec)
    return _lm_vlookup_join_sheets_left(doc, left_sheet, spec)


def _lm_vlookup_mirror_spec(spec):
    """Зеркало для правого прохода full join."""
    out = dict(spec)
    out["left_sheet"] = spec.get("right_sheet")
    out["right_sheet"] = spec.get("left_sheet")
    out["left_range"] = spec.get("right_range")
    out["right_range"] = spec.get("left_range")
    out["left_start_row"] = spec.get("right_start_row")
    out["left_start_col"] = spec.get("right_start_col")
    out["left_header_row"] = spec.get("right_header_row")
    out["left_end_row"] = spec.get("right_end_row")
    out["left_end_col"] = spec.get("right_end_col")
    out["right_start_row"] = spec.get("left_start_row")
    out["right_start_col"] = spec.get("left_start_col")
    out["right_header_row"] = spec.get("left_header_row")
    out["right_end_row"] = spec.get("left_end_row")
    out["right_end_col"] = spec.get("left_end_col")
    out["join_keys"] = list(spec.get("join_keys_right") or spec.get("join_keys") or [])
    out["join_keys_left"] = list(spec.get("join_keys_right") or spec.get("join_keys") or [])
    out["join_keys_right"] = list(spec.get("join_keys_left") or spec.get("join_keys") or [])
    out["extract_cols"] = list(spec.get("extract_cols_right") or [])
    out["join_type"] = "left"
    out["_vlookup_skip_full_dispatch"] = True
    return out


def _lm_vlookup_merge_stats(a, b):
    out = dict(a or {})
    if not b:
        return out
    for k in (
        "rows_added",
        "rows_processed",
        "rows_matched",
        "rows_not_found",
        "rows_multi",
        "columns_added",
        "keymap_size",
    ):
        out[k] = int(out.get(k, 0) or 0) + int(b.get(k, 0) or 0)
    return out


def _lm_vlookup_append_right_only_rows(doc, left_sheet, right_sheet, spec):
    """Добавить на левый лист строки правой таблицы без ключа слева (full join)."""
    left_key_tokens = list(spec.get("join_keys_left") or spec.get("join_keys") or [])
    right_key_tokens = list(spec.get("join_keys_right") or spec.get("join_keys") or [])
    col_tokens = list(spec.get("extract_cols") or [])
    if len(left_key_tokens) == 0 or len(right_key_tokens) == 0 or len(col_tokens) == 0:
        return {"rows_added": 0, "rows_processed": 0}

    left_bounds = _lm_vlookup_resolve_bounds(
        left_sheet,
        spec.get("left_start_row"),
        spec.get("left_start_col"),
        spec.get("left_header_row"),
        range_a1=spec.get("left_range"),
        end_row_1=spec.get("left_end_row"),
        end_col_1=spec.get("left_end_col"),
    )
    # После left-pass used-area слева могла вырасти — расширяем end_row.
    ua_sc, ua_sr, ua_ec, ua_er = _lm_vlookup_sheet_used_area(left_sheet)
    if ua_er > left_bounds["end_row"]:
        left_bounds = dict(left_bounds)
        left_bounds["end_row"] = ua_er
        left_bounds["end_col"] = max(left_bounds["end_col"], ua_ec)

    right_bounds = _lm_vlookup_resolve_bounds(
        right_sheet,
        spec.get("right_start_row"),
        spec.get("right_start_col"),
        spec.get("right_header_row"),
        range_a1=spec.get("right_range"),
        end_row_1=spec.get("right_end_row"),
        end_col_1=spec.get("right_end_col"),
    )
    left_key_cols = _lm_vlookup_resolve_key_columns(left_sheet, left_bounds, left_key_tokens)
    right_key_cols = _lm_vlookup_resolve_key_columns(right_sheet, right_bounds, right_key_tokens)
    extract_specs = _lm_vlookup_resolve_extract_columns(right_sheet, right_bounds, col_tokens)
    trim_keys = bool(spec.get("trim_keys", False))
    if spec.get("run_suffix") is not None:
        run_suffix = str(spec.get("run_suffix"))
    else:
        run_suffix = "_R1"
    hc_raw = spec.get("highlight_color")
    if hc_raw is None or str(hc_raw).strip() == "" or str(hc_raw).strip().casefold() == u"нет":
        highlight_style = None
    else:
        try:
            highlight_style = _lm_vlookup_parse_highlight_style(hc_raw)
        except Exception:
            highlight_style = None

    left_key_set = set()
    row = left_bounds["data_start"]
    while row <= left_bounds["end_row"]:
        key_vals = _lm_vlookup_row_key_values(left_sheet, row, left_key_cols)
        key = _lm_vlookup_composite_key(key_vals, trim=trim_keys)
        if key is not None and key != "":
            left_key_set.add(key)
        row = row + 1

    key_col_map = {}
    seeded = spec.get("output_key_col_map")
    if isinstance(seeded, dict):
        for _k, _v in seeded.items():
            try:
                key_col_map[str(_k)] = int(_v)
            except Exception:
                pass
    _lm_vlookup_seed_key_col_map_from_headers(
        left_sheet, left_bounds, extract_specs, run_suffix, key_col_map
    )
    template_col = left_bounds["end_col"]
    template_col = _lm_vlookup_ensure_all_output_columns(
        left_sheet,
        key_col_map,
        extract_specs,
        run_suffix,
        left_bounds,
        template_col,
        highlight_style,
        doc=doc,
        right_sheet=right_sheet,
        right_bounds=right_bounds,
    )

    keymap = _lm_vlookup_build_keymap(
        right_sheet, right_bounds, right_key_cols, extract_specs, trim_keys=trim_keys
    )
    append_at = left_bounds["end_row"] + 1
    rows_added = 0
    rows_processed = 0
    rrow = right_bounds["data_start"]
    while rrow <= right_bounds["end_row"]:
        key_vals = _lm_vlookup_row_key_values(right_sheet, rrow, right_key_cols)
        key = _lm_vlookup_composite_key(key_vals, trim=trim_keys)
        rows_processed = rows_processed + 1
        if key is None or key == "" or key in left_key_set:
            rrow = rrow + 1
            continue
        matches = keymap.get(key) or []
        if len(matches) == 0:
            rrow = rrow + 1
            continue
        # Одна строка на ключ (первый матч) — union без размножения мульти.
        ki = 0
        while ki < len(left_key_cols) and ki < len(key_vals):
            try:
                left_sheet.getCellByPosition(left_key_cols[ki], append_at).setString(
                    str(key_vals[ki] if key_vals[ki] is not None else "")
                )
            except Exception:
                pass
            ki = ki + 1
        _lm_vlookup_fill_extract_row(
            left_sheet,
            append_at,
            key_col_map,
            extract_specs,
            run_suffix,
            matches[0],
            highlight_style,
        )
        left_key_set.add(key)
        append_at = append_at + 1
        rows_added = rows_added + 1
        rrow = rrow + 1
    return {
        "rows_added": rows_added,
        "rows_processed": rows_processed,
        "rows_matched": rows_added,
        "rows_not_found": 0,
        "rows_multi": 0,
        "columns_added": len(key_col_map),
        "keymap_size": len(keymap),
    }


def _lm_vlookup_join_full(doc, left_sheet, spec):
    """Полное соединение: left-pass + right-pass + union строк только справа."""
    import time

    t0 = time.time()
    left_spec = dict(spec)
    left_spec["join_type"] = "left"
    left_spec["_vlookup_skip_full_dispatch"] = True
    r1 = _lm_vlookup_join_sheets_left(doc, left_sheet, left_spec)

    right_sheet = _lm_vlookup_resolve_sheet(doc, spec.get("right_sheet"))
    if right_sheet is None:
        raise ValueError("правый лист «%s» не найден" % spec.get("right_sheet", "?"))
    mirror = _lm_vlookup_mirror_spec(spec)
    r2 = _lm_vlookup_join_sheets_left(doc, right_sheet, mirror)
    append_spec = dict(left_spec)
    if isinstance(r1.get("output_key_col_map"), dict):
        append_spec["output_key_col_map"] = dict(r1.get("output_key_col_map"))
    r3 = _lm_vlookup_append_right_only_rows(doc, left_sheet, right_sheet, append_spec)
    out = _lm_vlookup_merge_stats(_lm_vlookup_merge_stats(r1, r2), r3)
    elapsed = time.time() - t0
    left_name = left_sheet.Name if left_sheet is not None else "?"
    right_name = right_sheet.Name if right_sheet is not None else "?"
    _lm_log_postprocess(
        doc,
        left_name,
        "впр",
        "%s ↔ %s" % (left_name, right_name),
        "ok",
        "full_join; L+R+union; обработано=%d; добавлено_строк=%d; столбцов=%d; сек=%.2f"
        % (
            int(out.get("rows_processed", 0) or 0),
            int(out.get("rows_added", 0) or 0),
            int(out.get("columns_added", 0) or 0),
            elapsed,
        ),
    )
    out["join_type"] = "full"
    out["elapsed_sec"] = elapsed
    return out


def _lm_vlookup_join_sheets_left(doc, left_sheet, spec):
    """
    LEFT OUTER JOIN двух листов книги результата (TZ_VLOOKUP2.md §7).

    Параметры:
        doc — книга результата.
        left_sheet — левый лист (XSpreadsheet).
        spec — dict из parse_vlookup_spec + run_suffix (_R1, _R2, …).

    Возвращает:
        dict: rows_added, rows_processed, columns_added, keymap_size.
    """
    if doc is None or left_sheet is None or spec is None:
        raise ValueError("ВПР: не заданы doc, лист или spec")

    right_sheet = _lm_vlookup_resolve_sheet(doc, spec.get("right_sheet"))
    if right_sheet is None:
        raise ValueError(
            "правый лист «%s» не найден" % spec.get("right_sheet", "?")
        )

    if spec.get("run_suffix") is not None:
        run_suffix = str(spec.get("run_suffix"))
    else:
        run_suffix = "_R1"
    left_key_tokens = list(
        spec.get("join_keys_left") or spec.get("join_keys") or []
    )
    right_key_tokens = list(
        spec.get("join_keys_right") or spec.get("join_keys") or []
    )
    col_tokens = list(spec.get("extract_cols") or [])
    if len(left_key_tokens) == 0:
        raise ValueError("не указаны ключи левой таблицы")
    if len(right_key_tokens) == 0:
        raise ValueError("не указаны ключи правой таблицы")
    if len(col_tokens) == 0:
        raise ValueError("не указаны столбцы для извлечения")

    left_bounds = _lm_vlookup_resolve_bounds(
        left_sheet,
        spec.get("left_start_row"),
        spec.get("left_start_col"),
        spec.get("left_header_row"),
        range_a1=spec.get("left_range"),
        end_row_1=spec.get("left_end_row"),
        end_col_1=spec.get("left_end_col"),
    )
    right_bounds = _lm_vlookup_resolve_bounds(
        right_sheet,
        spec.get("right_start_row"),
        spec.get("right_start_col"),
        spec.get("right_header_row"),
        range_a1=spec.get("right_range"),
        end_row_1=spec.get("right_end_row"),
        end_col_1=spec.get("right_end_col"),
    )

    left_key_cols = _lm_vlookup_resolve_key_columns(
        left_sheet, left_bounds, left_key_tokens
    )
    right_key_cols = _lm_vlookup_resolve_key_columns(
        right_sheet, right_bounds, right_key_tokens
    )
    extract_specs = _lm_vlookup_resolve_extract_columns(
        right_sheet, right_bounds, col_tokens
    )

    # Во время сбора авто-пересчёт выключен: getDataArray() без calculateAll()
    # вернёт 0 для ячеек с формулами (например, столбец из «применить_формулу»).
    _lm_pp_calculate_doc(doc)

    hc_raw = spec.get("highlight_color")
    if hc_raw is None or str(hc_raw).strip() == "":
        highlight_style = None
    elif str(hc_raw).strip().casefold() == u"нет":
        highlight_style = None
    else:
        try:
            if lm_parse_bool_param(hc_raw, default=None) is False:
                highlight_style = None
            else:
                highlight_style = _lm_vlookup_parse_highlight_style(hc_raw)
        except ValueError:
            highlight_style = _lm_vlookup_parse_highlight_style(hc_raw)
    fill_duplicates = spec.get("fill_duplicates", True)
    if fill_duplicates is None:
        fill_duplicates = True
    else:
        try:
            fill_duplicates = bool(fill_duplicates)
        except Exception:
            fill_duplicates = True
    multi_match_one_cell = spec.get("multi_match_one_cell", False)
    if multi_match_one_cell is None:
        multi_match_one_cell = False
    else:
        try:
            multi_match_one_cell = bool(multi_match_one_cell)
        except Exception:
            multi_match_one_cell = False
    multi_match_mode = _lm_vlookup_normalize_multi_match(spec.get("multi_match"), default="all")
    is_inner = _lm_vlookup_normalize_join_type(spec.get("join_type"), default="left") == "inner"
    not_found_fill = spec.get("not_found_fill")
    if not_found_fill is None or str(not_found_fill).strip() == "":
        not_found_fill = "#Н/Д"

    trim_keys = spec.get("trim_keys", False)
    if trim_keys is None:
        trim_keys = False
    else:
        try:
            trim_keys = bool(trim_keys)
        except Exception:
            trim_keys = False

    left_name = left_sheet.Name if left_sheet is not None else "?"
    right_name = right_sheet.Name if right_sheet is not None else "?"
    extract_map = []
    ei0 = 0
    while ei0 < len(extract_specs):
        ec_idx, out_title = extract_specs[ei0]
        tok0 = col_tokens[ei0] if ei0 < len(col_tokens) else "?"
        extract_map.append("%s→«%s»(R%d)" % (tok0, out_title, int(ec_idx) + 1))
        ei0 = ei0 + 1
    _lm_log_postprocess(
        doc,
        left_name,
        "впр",
        "%s → %s" % (left_name, right_name),
        "старт",
        "ключи L=[%s] R=[%s]; extract=[%s]; suffix=«%s»; trim=%s; fill_dup=%s; multi=%s; multi_match=%s; join=%s; nf=«%s»; L(hdr=%d data=%d..%d col=%d..%d) R(hdr=%d data=%d..%d col=%d..%d)"
        % (
            ",".join([str(t) for t in left_key_tokens]),
            ",".join([str(t) for t in right_key_tokens]),
            "; ".join(extract_map) if extract_map else "—",
            run_suffix,
            u"да" if trim_keys else u"нет",
            u"да" if fill_duplicates else u"нет",
            u"да" if multi_match_one_cell else u"нет",
            multi_match_mode,
            u"inner" if is_inner else u"left",
            not_found_fill,
            left_bounds["header_row"] + 1,
            left_bounds["data_start"] + 1,
            left_bounds["end_row"] + 1,
            left_bounds["start_col"] + 1,
            left_bounds["end_col"] + 1,
            right_bounds["header_row"] + 1,
            right_bounds["data_start"] + 1,
            right_bounds["end_row"] + 1,
            right_bounds["start_col"] + 1,
            right_bounds["end_col"] + 1,
        ),
    )

    keymap = _lm_vlookup_build_keymap(
        right_sheet, right_bounds, right_key_cols, extract_specs, trim_keys=trim_keys
    )
    _lm_log_postprocess(
        doc,
        left_name,
        "впр",
        "%s → %s" % (left_name, right_name),
        "инфо",
        "keymap=%d ключей; столбцы выхода: %s"
        % (
            len(keymap),
            ",".join([t + run_suffix for (_c, t) in extract_specs]) if extract_specs else "—",
        ),
    )

    key_col_map = {}
    template_col = left_bounds["end_col"]
    template_col = _lm_vlookup_ensure_all_output_columns(
        left_sheet,
        key_col_map,
        extract_specs,
        run_suffix,
        left_bounds,
        template_col,
        highlight_style,
        doc=doc,
        right_sheet=right_sheet,
        right_bounds=right_bounds,
    )
    rows_added = 0
    rows_processed = 0
    rows_matched = 0
    rows_not_found = 0
    rows_multi = 0
    rows_deleted = 0
    unmatched_rows = []
    counter = 0

    row = left_bounds["data_start"]
    max_row = left_bounds["end_row"]
    if row > max_row:
        _lm_log_postprocess(
            doc,
            left_name,
            "впр",
            "%s → %s" % (left_name, right_name),
            "ok",
            "нет строк данных слева; keymap=%d; столбцов=%d"
            % (len(keymap), len(key_col_map)),
        )
        return {
            "rows_added": 0,
            "rows_processed": 0,
            "rows_matched": 0,
            "rows_not_found": 0,
            "rows_multi": 0,
            "columns_added": len(key_col_map),
            "keymap_size": len(keymap),
            "output_key_col_map": dict(key_col_map),
        }

    while row <= max_row:
        key_vals = _lm_vlookup_row_key_values(left_sheet, row, left_key_cols)
        key = _lm_vlookup_composite_key(key_vals, trim=trim_keys)
        matches = keymap.get(key)
        if matches is not None and len(matches) > 0:
            matches = _lm_vlookup_select_matches(matches, multi_match_mode)

        if matches is not None and len(matches) > 0:
            n_match = len(matches)
            rows_matched = rows_matched + 1
            if n_match > 1:
                rows_multi = rows_multi + 1
            if multi_match_one_cell and n_match > 1:
                _lm_vlookup_fill_multi_match_row(
                    left_sheet,
                    row,
                    key_col_map,
                    extract_specs,
                    run_suffix,
                    matches,
                    highlight_style,
                )
                row = row + 1
            else:
                if n_match > 1:
                    n_extra = n_match - 1
                    insert_at = row + 1
                    left_sheet.getRows().insertByIndex(insert_at, n_extra)
                    if fill_duplicates:
                        copy_end = left_bounds["end_col"]
                        di = 0
                        while di < n_extra:
                            _lm_vlookup_copy_left_row(
                                left_sheet, row, insert_at + di, left_bounds["start_col"], copy_end
                            )
                            di = di + 1
                    max_row = max_row + n_extra
                    rows_added = rows_added + n_extra

                mi = 0
                while mi < n_match:
                    target_row = row + mi
                    _lm_vlookup_fill_extract_row(
                        left_sheet,
                        target_row,
                        key_col_map,
                        extract_specs,
                        run_suffix,
                        matches[mi],
                        highlight_style,
                    )
                    mi = mi + 1
                row = row + n_match
        else:
            rows_not_found = rows_not_found + 1
            if is_inner:
                unmatched_rows.append(row)
            else:
                _lm_vlookup_fill_not_found_row(
                    left_sheet,
                    row,
                    key_col_map,
                    extract_specs,
                    run_suffix,
                    not_found_fill,
                    highlight_style,
                )
            row = row + 1

        rows_processed = rows_processed + 1
        counter = counter + 1
        if counter % _LM_VLOOKUP_UI_YIELD_EVERY == 0:
            _lm_ui_yield(counter=counter)

    if is_inner and len(unmatched_rows) > 0:
        ui = len(unmatched_rows) - 1
        while ui >= 0:
            try:
                _lm_sheet_delete_rows(left_sheet, unmatched_rows[ui], 1)
                rows_deleted = rows_deleted + 1
            except Exception:
                pass
            ui = ui - 1

    out_cols = []
    for _k, _v in key_col_map.items():
        out_cols.append("%s@%d" % (_k, int(_v) + 1))
    _lm_log_postprocess(
        doc,
        left_name,
        "впр",
        "%s → %s" % (left_name, right_name),
        "ok",
        "обработано=%d; совпало=%d; не_найдено=%d; удалено=%d; мульти=%d; добавлено_строк=%d; keymap=%d; столбцы=[%s]"
        % (
            rows_processed,
            rows_matched,
            rows_not_found,
            rows_deleted,
            rows_multi,
            rows_added,
            len(keymap),
            ",".join(out_cols) if out_cols else "—",
        ),
    )
    return {
        "rows_added": rows_added,
        "rows_processed": rows_processed,
        "rows_matched": rows_matched,
        "rows_not_found": rows_not_found,
        "rows_deleted": rows_deleted,
        "rows_multi": rows_multi,
        "columns_added": len(key_col_map),
        "keymap_size": len(keymap),
        "output_key_col_map": dict(key_col_map),
        "join_type": "inner" if is_inner else "left",
    }


# endregion Вспомогательные функции постобработки для collect_workbooks.


# Имена для регистрации в макросах-хостах (ключ merge_pp_* → lm_pp_*).

# region Сводные таблицы (libre_macros_pivot_lib)
from libre_macros_pivot_lib import (  # noqa: E402
    PIVOT_COLLECT_API,
    PIVOT_CONFIG_VERSION,
    _LM_PP_PIVOT_FUNCTION,
    _LM_PP_PIVOT_MATERIALIZED,
    _LM_PP_PIVOT_ORIENTATION,
    _lm_pp_format_pivot_output_sheet,
    _lm_pp_is_pivot_followup_skip_name,
    _lm_pp_is_pivot_table_postprocess_name,
    _lm_pp_pivot_cell_has_content,
    _lm_pp_pivot_col_index,
    _lm_pp_pivot_create_table,
    _lm_pp_pivot_detect_data_header_row,
    _lm_pp_pivot_detect_header_row_by_consecutive_cells,
    _lm_pp_pivot_enable_repeat_item_labels,
    _lm_pp_pivot_field_display_name,
    _lm_pp_pivot_flatten_range_values,
    _lm_pp_pivot_format_addr,
    _lm_pp_pivot_get_datapilot_tables,
    _lm_pp_pivot_guess_header_row,
    _lm_pp_pivot_header_names_from_range,
    _lm_pp_pivot_header_row_index,
    _lm_pp_pivot_list_descriptor_fields,
    _lm_pp_pivot_orientation,
    _lm_pp_pivot_output_has_content,
    _lm_pp_pivot_parse_config,
    _lm_pp_pivot_parse_corner,
    _lm_pp_pivot_read_header_titles,
    _lm_pp_pivot_resolve_bounds,
    _lm_pp_pivot_resolve_field,
    _lm_pp_pivot_resolve_output_bounds,
    _lm_pp_pivot_resolve_source_sheet,
    _lm_pp_pivot_row_is_preamble,
    _lm_pp_pivot_set_field,
    _lm_pp_pivot_sheet_end_bounds,
    _lm_pp_pivot_sheet_index,
    _lm_pp_pivot_sheet_names_set,
    _lm_pp_pivot_sheet_replace_if_old_output,
    _lm_pp_pivot_source_sheet_matches,
    _lm_pp_pivot_table_extra_args,
    _lm_pp_pivot_trim_to_header_row,
    _lm_pp_pivot_trim_top_rows,
    _lm_pp_pivot_function,
    lm_pp_pivot_purge_lo_default_garbage_sheets,
    lm_pp_pivot_clear_materialized_registry,
    lm_pp_pivot_clear_as_values_pending,
    lm_pp_pivot_collect_field_names,
    lm_pp_pivot_config_from_json,
    lm_pp_pivot_config_to_json,
    lm_pp_pivot_default_config,
    lm_pp_pivot_get_list_rows_meta,
    lm_pp_pivot_apply_list_rows,
    lm_pp_pivot_list_as_values_pending,
    lm_pp_pivot_list_materialized_sheets,
    lm_pp_pivot_post_values_format,
    lm_pp_pivot_register_as_values_pending,
    lm_pp_pivot_register_materialized_sheet,
    lm_pp_pivot_set_post_values_format_enabled,
    lm_pp_pivot_step_applies_to_sheet,
    lm_pp_range_pivot_table,
)
# endregion Сводные таблицы

# region Финальная обработка (libre_macros_final_lib)
from libre_macros_final_lib import (  # noqa: E402
    LM_FINAL_PUBLIC_NAMES,
    _LM_FINAL_SKIP_PIVOT_FOLLOWUP_KEYS,
    _lm_final_is_pivot_followup_skip_name,
    lm_final_apply_pending_pivot_as_values,
    lm_final_copy_sheet,
    lm_final_copy_ranges,
    lm_final_create_sheet,
    lm_final_send_mail,
    lm_final_activate_sheet,
    lm_final_delete_sheets,
    lm_final_hide_sheets,
    lm_final_merge_sheets_into_one,
    lm_final_values_only,
)


def lm_pp_range_copy_ranges(doc, sheet, data_range, header_row_range, *extra_args):
    """RANGE: копирование_диапазонов (lazy import — без цикла с copy_ranges_lib)."""
    from libre_macros_copy_ranges_lib import lm_pp_range_copy_ranges as _fn

    return _fn(doc, sheet, data_range, header_row_range, *extra_args)


def lm_pp_range_create_sheet(doc, sheet, data_range, header_row_range, *extra_args):
    """RANGE: создать_лист (lazy import)."""
    from libre_macros_create_sheet_lib import lm_pp_range_create_sheet as _fn

    return _fn(doc, sheet, data_range, header_row_range, *extra_args)


def lm_pp_range_unpivot_columns(doc, sheet, data_range, header_row_range, *extra_args):
    """RANGE: развернуть_столбцы / unpivot (lazy import)."""
    from libre_macros_unpivot_lib import lm_pp_range_unpivot_columns as _fn

    return _fn(doc, sheet, data_range, header_row_range, *extra_args)


def lm_pp_range_transpose_table(doc, sheet, data_range, header_row_range, *extra_args):
    """RANGE: транспонировать_таблицу (lazy import)."""
    from libre_macros_transpose_lib import lm_pp_range_transpose_table as _fn

    return _fn(doc, sheet, data_range, header_row_range, *extra_args)


def lm_pp_range_form_to_table(doc, sheet, data_range, header_row_range, *extra_args):
    """RANGE: анкета_в_таблицу (lazy import)."""
    from libre_macros_form_table_lib import lm_pp_range_form_to_table as _fn

    return _fn(doc, sheet, data_range, header_row_range, *extra_args)


def lm_pp_range_table_to_form(doc, sheet, data_range, header_row_range, *extra_args):
    """RANGE: таблица_в_анкету (lazy import)."""
    from libre_macros_form_table_lib import lm_pp_range_table_to_form as _fn

    return _fn(doc, sheet, data_range, header_row_range, *extra_args)


# endregion Финальная обработка

# region Финальная обработка (libre_macros_values_lib)
from libre_macros_values_lib import (  # noqa: E402
    lm_final_apply_formula,
    lm_final_autofilter,
    lm_final_autofit_columns,
    lm_final_autofit_rows,
    lm_final_color_scale,
    lm_final_colorize_blocks,
    lm_final_colorize,
    lm_final_column_width,
    lm_final_concat_columns,
    lm_final_conditional_column,
    lm_final_lambda_column,
    lm_final_add_column,
    lm_final_split_by_columns,
    lm_final_copy_values,
    lm_final_replace_values,
    lm_final_normalize_text,
    lm_final_delete_columns,
    lm_final_delete_rows,
    lm_final_fill_down,
    lm_final_fill_up,
    lm_final_unpivot_columns,
    lm_final_transpose_table,
    lm_final_form_to_table,
    lm_final_table_to_form,
    lm_final_fill_down_calculate,
    lm_final_format_columns,
    lm_final_format_date,
    lm_final_format_money,
    lm_final_freeze_header,
    lm_final_grid,
    lm_final_group_by_column,
    lm_final_header_plus_height,
    lm_final_highlight_threshold,
    lm_final_indent,
    lm_final_left_align,
    lm_final_print_style,
    lm_final_rename_sheet,
    lm_final_rename_columns,
    lm_final_reorder_columns,
    lm_final_remove_duplicates,
    lm_final_reset_path_stripes,
    lm_final_row_height,
    lm_final_set_font,
    lm_final_sort,
    lm_final_thick_grid,
    lm_final_thin_grid,
    lm_final_value_count,
    lm_final_vert_center,
    lm_final_word_wrap,
    lm_final_word_wrap_and_fit,
    lm_final_zebra,
)
# endregion Финальная обработка (values)

LM_PP_PUBLIC_NAMES = (
    "lm_pp_range_add_filter",
    "lm_pp_range_apply_formula",
    "lm_pp_range_autofit_columns_width",
    "lm_pp_range_autofit_row_heights",
    "lm_pp_range_color_scale_simple",
    "lm_pp_range_concat_columns",
    "lm_pp_range_colorize_data",
    "lm_pp_colorize_step_applies_to_sheet",
    "lm_pp_format_step_applies_to_sheet",
    "lm_pp_range_highlight_by_threshold",
    "lm_pp_range_conditional_format_simple",
    "lm_pp_range_delete_columns",
    "lm_pp_range_delete_sheets",
    "lm_pp_range_format_columns",
    "lm_pp_range_format_date_columns",
    "lm_pp_range_format_money_columns",
    "lm_pp_range_freeze_header",
    "lm_pp_range_grid_borders",
    "lm_pp_range_group_by_column",
    "lm_pp_range_header_row_height_pad",
    "lm_pp_range_hide_service_columns",
    "lm_pp_range_hide_sheets",
    "lm_pp_range_increase_indent",
    "lm_pp_range_init_path_stripe_state",
    "lm_pp_range_left_align_data",
    "lm_pp_range_row_height_pad",
    "lm_pp_range_set_column_widths",
    "lm_pp_range_set_font",
    "lm_pp_range_set_page_style",
    "lm_pp_range_conditional_column",
    "lm_pp_range_lambda_column",
    "lm_pp_range_add_column",
    "lm_pp_range_split_by_columns",
    "lm_pp_range_skip_empty_rows",
    "lm_pp_range_thick_grid_borders",
    "lm_pp_range_thin_grid_borders",
    "lm_pp_range_word_wrap",
    "lm_pp_range_word_wrap_and_fit_rows",
    "lm_pp_range_zebra_even_rows",
    "lm_pp_row_alternate_borders",
    "lm_pp_row_bold_if_path_contains",
    "lm_pp_row_convert_numeric_strings",
    "lm_pp_row_copy_format_from_header",
    "lm_pp_row_gray_service_columns",
    "lm_pp_row_height_pad",
    "lm_pp_row_highlight_by_header_value",
    "lm_pp_row_stripe_on_path_change",
    "lm_pp_row_text_color_by_value",
    "lm_pp_row_zebra_stripes",
    "lm_pp_range_fill_down_empty",
    "lm_pp_range_fill_up_empty",
    "lm_pp_range_unpivot_columns",
    "lm_pp_range_transpose_table",
    "lm_pp_range_form_to_table",
    "lm_pp_range_table_to_form",
    "lm_pp_range_copy_values",
    "lm_pp_range_fill_down_calculate",
    "lm_pp_range_replace_values",
    "lm_pp_range_normalize_text",
    "lm_pp_range_rename_sheet",
    "lm_pp_range_rename_columns",
    "lm_pp_range_reorder_columns",
    "lm_pp_range_value_count",
    "lm_pp_range_remove_duplicates",
    "lm_pp_range_copy_sheet",
    "lm_pp_range_copy_ranges",
    "lm_pp_range_create_sheet",
    "lm_pp_range_merge_sheets_into_one",
    "lm_pp_copy_sheet_clear_pending",
    "lm_pp_copy_sheet_remove_pending_names",
    "lm_pp_copy_sheet_register_pending",
    "lm_pp_copy_sheet_take_pending",
    "lm_pp_copy_sheet_list_pending",
    "lm_pp_range_pivot_table",
    "lm_pp_range_sort_data",
    "lm_pp_sort_step_applies_to_sheet",
    "lm_pp_color_scale_step_applies_to_sheet",
    "lm_pp_apply_formula_step_applies_to_sheet",
    "lm_pp_header_height_step_applies_to_sheet",
    "lm_pp_grid_step_applies_to_sheet",
    "lm_pp_grid_preset_step_applies_to_sheet",
    "lm_vlookup_join_sheets"
)
