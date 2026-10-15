# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
"""
Финальная обработка «подготовка_писем»: AText (Writer) из данных ACell (Calc).

Один выходной файл на блок строк (block_rows или marker-merge).
Плейсхолдеры: <<Столбец>>, <<Столбец>>{k}, <<n>>, <<row>>{k}.
"""
MACRO_VERSION = "3.10.780"
import os
import re
import subprocess
import sys

try:
    unicode
except NameError:
    unicode = str  # noqa: A001

try:
    import uno
except ImportError:
    uno = None

_PLACEHOLDER_RE = re.compile(r"<<([^>]+?)>>(?:\{(\d+)\})?")
_FILENAME_BAD_RE = re.compile(r'[\\/:*?"<>|]+')
_ATEXT_FORMATS = frozenset((u"odt", u"pdf", u"docx"))
_MISSING_MODES = frozenset((u"empty", u"keep", u"fail"))
_DEFAULT_MAX_FILES = 500
_MAX_FILENAME_LEN = 180

_FILTER_BY_EXT = {
    u".odt": u"writer8",
    u".ott": u"writer8",
    u".pdf": u"writer_pdf_Export",
    u".docx": u"MS Word 2007 XML",
}


def _u(v):
    try:
        return unicode(v or u"")
    except Exception:
        try:
            return str(v or "")
        except Exception:
            return u""


def _make_prop(name, value):
    p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    p.Name = str(name)
    p.Value = value
    return p


def _path_to_url(path):
    return uno.systemPathToFileUrl(os.path.abspath(str(path)))


def _script_context():
    try:
        import __main__

        return getattr(__main__, "XSCRIPTCONTEXT", None)
    except Exception:
        return None


def _get_desktop():
    xsc = _script_context()
    if xsc is not None:
        try:
            return xsc.getDesktop()
        except Exception:
            pass
    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    except Exception:
        return None


def _expand_vars(text, doc=None, sheet=None):
    raw = _u(text)
    if raw == u"" or u"<<Переменные." not in raw:
        return raw
    try:
        from libre_macros_lib import _lm_pp_expand_source_variables_template

        return _lm_pp_expand_source_variables_template(
            raw,
            doc=doc,
            sheet=sheet,
            log_fn_key=u"подготовка_писем",
            for_formula=False,
        )
    except Exception:
        return raw


def _strip_col_quotes(name):
    s = _u(name).strip()
    if len(s) >= 2 and s[0] == u"'" and s[-1] == u"'":
        return s[1:-1].strip()
    return s


def _sanitize_filename(name):
    s = _FILENAME_BAD_RE.sub(u"_", _u(name).strip())
    s = s.replace(u"\n", u" ").replace(u"\r", u" ").replace(u"\t", u" ")
    while u"  " in s:
        s = s.replace(u"  ", u" ")
    s = s.strip(u" ._")
    if s == u"":
        s = u"letter"
    if len(s) > _MAX_FILENAME_LEN:
        s = s[:_MAX_FILENAME_LEN].rstrip(u" ._")
    return s or u"letter"


def _ensure_dir(path):
    path = os.path.abspath(_u(path).strip())
    if path == u"":
        return u"", u"out_dir пуст"
    try:
        if not os.path.isdir(path):
            os.makedirs(path)
    except Exception as err:
        return u"", u"не создать out_dir: %s" % err
    return path, None


def _open_folder(path):
    path = os.path.abspath(_u(path))
    if path == u"" or not os.path.isdir(path):
        return
    try:
        if sys.platform.startswith(u"win"):
            os.startfile(path)  # noqa: B606
            return
        if sys.platform == u"darwin":
            subprocess.Popen([u"open", path])
            return
        subprocess.Popen([u"xdg-open", path])
    except Exception:
        pass


def _filter_name_for_path(path):
    ext = os.path.splitext(_u(path))[1].lower()
    return _FILTER_BY_EXT.get(ext, u"writer8")


def _load_template(path):
    desktop = _get_desktop()
    if desktop is None:
        return None, u"нет Desktop"
    props = (
        _make_prop(u"Hidden", True),
        _make_prop(u"AsTemplate", True),
    )
    try:
        doc = desktop.loadComponentFromURL(_path_to_url(path), u"_blank", 0, props)
    except Exception as err:
        return None, u"открыть шаблон: %s" % err
    if doc is None:
        return None, u"шаблон не загружен"
    return doc, None


def _close_doc(doc):
    if doc is None:
        return
    try:
        doc.close(True)
    except Exception:
        try:
            doc.dispose()
        except Exception:
            pass


def _store_doc(doc, path, overwrite=True):
    path = os.path.abspath(_u(path))
    if overwrite and os.path.isfile(path):
        try:
            os.unlink(path)
        except Exception:
            pass
    props = (
        _make_prop(u"FilterName", _filter_name_for_path(path)),
        _make_prop(u"Overwrite", bool(overwrite)),
    )
    try:
        doc.storeToURL(_path_to_url(path), props)
        return True, None
    except Exception as err:
        return False, u"storeToURL: %s" % err


def _replace_all(doc, search, replace):
    if search is None or search == u"":
        return
    try:
        desc = doc.createReplaceDescriptor()
        desc.SearchString = _u(search)
        desc.ReplaceString = _u(replace)
        try:
            desc.SearchCaseSensitive = True
        except Exception:
            pass
        try:
            desc.SearchRegularExpression = False
        except Exception:
            pass
        doc.replaceAll(desc)
    except Exception:
        pass


def _sync_user_fields(doc, values_row1):
    """User Field «Имя» ← значение строки 1 блока (без индекса)."""
    if doc is None or not values_row1:
        return
    try:
        masters = doc.getTextFieldMasters()
    except Exception:
        return
    try:
        names = list(masters.getElementNames() or ())
    except Exception:
        names = []
    for full in names:
        short = _u(full)
        # com.sun.star.text.fieldmaster.User.ФИО
        if u"." in short:
            short = short.rsplit(u".", 1)[-1]
        key = None
        from libre_macros_lib import lm_identity_key

        for hdr, val in values_row1.items():
            if lm_identity_key(hdr) == lm_identity_key(short):
                key = hdr
                break
        if key is None:
            continue
        try:
            master = masters.getByName(full)
            master.Content = _u(values_row1.get(key))
            try:
                master.update()
            except Exception:
                pass
        except Exception:
            pass
    try:
        enum = doc.getTextFields().createEnumeration()
        while enum.hasMoreElements():
            field = enum.nextElement()
            try:
                field.update()
            except Exception:
                pass
    except Exception:
        pass


def _used_last_row(sheet, header_row0):
    """0-based последняя строка used area; fallback header_row0."""
    try:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        last = int(cursor.getRangeAddress().EndRow)
        if last >= int(header_row0):
            return last
    except Exception:
        pass
    return int(header_row0)


def _used_last_col(sheet, header_row0):
    try:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        return int(cursor.getRangeAddress().EndColumn)
    except Exception:
        pass
    # fallback: scan header row
    last = 0
    try:
        from libre_macros_lib import cell_text

        c = 0
        while c < 256:
            if cell_text(sheet.getCellByPosition(c, int(header_row0))).strip() != u"":
                last = c
            c = c + 1
    except Exception:
        pass
    return last


def _read_headers(sheet, header_row0, c0, c1):
    from libre_macros_header_lib import clean_header_name
    from libre_macros_lib import cell_text

    headers = []
    c = int(c0)
    while c <= int(c1):
        raw = cell_text(sheet.getCellByPosition(c, int(header_row0)))
        headers.append(clean_header_name(raw))
        c = c + 1
    return headers


def _resolve_sheet(doc, name):
    name = _u(name).strip()
    if doc is None or name == u"":
        return None
    try:
        sheets = doc.getSheets()
        if sheets.hasByName(name):
            return sheets.getByName(name)
    except Exception:
        pass
    # casefold match
    try:
        from libre_macros_lib import lm_identity_key

        sheets = doc.getSheets()
        i = 0
        n = int(sheets.getCount())
        want = lm_identity_key(name)
        while i < n:
            sh = sheets.getByIndex(i)
            if lm_identity_key(sh.Name) == want:
                return sh
            i = i + 1
    except Exception:
        pass
    return None


def _cell_value(sheet, col, row0):
    from libre_macros_lib import cell_text

    try:
        return cell_text(sheet.getCellByPosition(int(col), int(row0)))
    except Exception:
        return u""


def _build_header_index(headers):
    from libre_macros_lib import lm_identity_key

    idx = {}
    i = 0
    while i < len(headers):
        h = headers[i]
        if h != u"":
            k = lm_identity_key(h)
            if k not in idx:
                idx[k] = i
        i = i + 1
    return idx


def _allowed_columns_set(columns_tokens, headers, header_idx):
    """None = все столбцы; иначе set identity keys."""
    if not columns_tokens:
        return None
    from libre_macros_lib import lm_identity_key
    from libre_macros_param_codec import column_header_matches_token

    allowed = set()
    for tok in columns_tokens:
        t = _u(tok).strip()
        if t == u"":
            continue
        i = 0
        while i < len(headers):
            if column_header_matches_token(headers[i], t):
                allowed.add(lm_identity_key(headers[i]))
            i = i + 1
        # буквы/индексы
        if lm_identity_key(t) not in allowed:
            try:
                from libre_macros_lib import col_letters_to_index, is_col_letters

                if is_col_letters(_strip_col_quotes(t)):
                    ci = col_letters_to_index(_strip_col_quotes(t))
                    if 0 <= ci < len(headers) and headers[ci] != u"":
                        allowed.add(lm_identity_key(headers[ci]))
                else:
                    n = int(_strip_col_quotes(t))
                    ci = n - 1
                    if 0 <= ci < len(headers) and headers[ci] != u"":
                        allowed.add(lm_identity_key(headers[ci]))
            except Exception:
                pass
    return allowed


def _row_matrix(sheet, headers, row0s, c0):
    """Список dict[header]=value для каждой физ. строки."""
    out = []
    for r0 in row0s:
        d = {}
        i = 0
        while i < len(headers):
            h = headers[i]
            if h != u"":
                d[h] = _cell_value(sheet, int(c0) + i, r0)
            i = i + 1
        out.append(d)
    return out


def _apply_row_filter(sheet, headers, row0s, c0, row_filter_code):
    code = _u(row_filter_code).strip()
    if code == u"":
        return list(row0s), None
    try:
        from libre_macros_lambda_column_lib import (
            call_rows_lambda,
            compile_rows_lambda,
            make_rows_accessor,
        )
    except Exception as err:
        return None, u"row_filter import: %s" % err
    try:
        compiled = compile_rows_lambda(code)
    except Exception as err:
        return None, u"row_filter: %s" % err
    if compiled is None:
        return list(row0s), None
    fn, glo = compiled
    matrix = []
    for r0 in row0s:
        row = []
        i = 0
        while i < len(headers):
            row.append(_cell_value(sheet, int(c0) + i, r0))
            i = i + 1
        matrix.append(row)
    accessor = make_rows_accessor(matrix, headers, start_col=0)
    kept = []
    ri = 0
    while ri < len(row0s):
        try:
            ok = call_rows_lambda(fn, glo, ri, rows_accessor=accessor)
        except Exception:
            ok = False
        if ok:
            kept.append(row0s[ri])
        ri = ri + 1
    return kept, None


def _chunk_rows(row0s, block_rows):
    n = max(1, int(block_rows))
    blocks = []
    i = 0
    while i < len(row0s):
        blocks.append(list(row0s[i : i + n]))
        i = i + n
    return blocks


def _marker_blocks(sheet, marker_token, data_r0, data_r1, header_row_range, c0, c1):
    from libre_macros_unmerge_collapse_lib import _find_marker_blocks, _resolve_marker_col

    block = {u"marker": marker_token}
    col, err = _resolve_marker_col(block, sheet, header_row_range, c0, c1)
    if err:
        return None, err
    spans = _find_marker_blocks(sheet, col, data_r0, data_r1, c0, c1)
    # Также одиночные строки без merge между блоками? TZ: высота = merge ≥2.
    # Строки вне merge не входят — как у add_merge.
    out = []
    for top, bot in spans:
        rows = list(range(int(top), int(bot) + 1))
        if rows:
            out.append(rows)
    return out, None


def _lookup_value(block_rows_dicts, headers, name, k, missing_mode, warnings):
    """k 1-based; name — имя столбца или спец. n/row обрабатываются снаружи."""
    from libre_macros_lib import lm_identity_key

    clean = _strip_col_quotes(name)
    if k < 1 or k > len(block_rows_dicts):
        warnings.append(u"<<%s>>{%d} вне блока" % (clean, k))
        if missing_mode == u"keep":
            return None  # сигнал keep
        if missing_mode == u"fail":
            raise ValueError(u"индекс {%d} вне блока (размер %d)" % (k, len(block_rows_dicts)))
        return u""
    row = block_rows_dicts[k - 1]
    want = lm_identity_key(clean)
    for h, v in row.items():
        if lm_identity_key(h) == want:
            return _u(v)
    warnings.append(u"столбец не найден: %s" % clean)
    if missing_mode == u"keep":
        return None
    if missing_mode == u"fail":
        raise ValueError(u"столбец не найден: %s" % clean)
    return u""


def _substitute_text(template, block_rows_dicts, file_n, phys_rows, missing_mode):
    """
    Замена <<Name>> / <<Name>>{k} / <<n>> / <<row>>{k}.
    phys_rows — 1-based физ. номера строк блока.
    Возвращает (text, warnings).
    """
    warnings = []
    text = _u(template)

    def _repl(m):
        raw_name = _u(m.group(1)).strip()
        k_s = m.group(2)
        k = int(k_s) if k_s else 1
        name_cf = raw_name.casefold()
        if name_cf in (u"n", u"№", u"no", u"num"):
            return unicode(int(file_n))
        if name_cf == u"row":
            if k < 1 or k > len(phys_rows):
                warnings.append(u"<<row>>{%d} вне блока" % k)
                if missing_mode == u"keep":
                    return m.group(0)
                if missing_mode == u"fail":
                    raise ValueError(u"<<row>>{%d} вне блока" % k)
                return u""
            return unicode(int(phys_rows[k - 1]))
        val = _lookup_value(
            block_rows_dicts, None, raw_name, k, missing_mode, warnings
        )
        if val is None:
            return m.group(0)
        return _u(val)

    # Сначала явные {k}, потом bare — regex уже различает.
    # Для keep: не трогаем совпадение.
    # replaceAll в Writer — построчно через наш regex для имени файла;
    # для документа соберём пары search→replace.
    return _PLACEHOLDER_RE.sub(_repl, text), warnings


def _collect_replacements(block_rows_dicts, file_n, phys_rows, missing_mode, allowed_keys):
    """
    Список (search, replace) для Writer: сначала <<Name>>{k}, потом <<Name>>.
    """
    from libre_macros_lib import lm_identity_key

    warnings = []
    pairs = []
    seen = set()

    def _add(search, value):
        if search in seen:
            return
        seen.add(search)
        pairs.append((search, value))

    # спец
    _add(u"<<n>>", unicode(int(file_n)))
    _add(u"<<N>>", unicode(int(file_n)))
    ki = 1
    while ki <= len(phys_rows):
        _add(u"<<row>>{%d}" % ki, unicode(int(phys_rows[ki - 1])))
        ki = ki + 1
    if phys_rows:
        _add(u"<<row>>", unicode(int(phys_rows[0])))

    # столбцы: для каждой строки блока
    # Собираем имена из первой строки + все ключи
    all_headers = []
    seen_h = set()
    for row in block_rows_dicts:
        for h in row.keys():
            ik = lm_identity_key(h)
            if ik and ik not in seen_h:
                seen_h.add(ik)
                all_headers.append(h)

    for h in all_headers:
        ik = lm_identity_key(h)
        if allowed_keys is not None and ik not in allowed_keys:
            continue
        k = 1
        while k <= max(1, len(block_rows_dicts)):
            val = _lookup_value(block_rows_dicts, None, h, k, missing_mode, warnings)
            if val is None:
                # keep — не добавляем пару
                pass
            else:
                _add(u"<<%s>>{%d}" % (h, k), _u(val))
                # также вариант с кавычками в имени
                _add(u"<<'%s'>>{%d}" % (h, k), _u(val))
            k = k + 1
        # bare = {1}
        val1 = _lookup_value(block_rows_dicts, None, h, 1, missing_mode, warnings)
        if val1 is not None:
            _add(u"<<%s>>" % h, _u(val1))
            _add(u"<<'%s'>>" % h, _u(val1))

    # Индексированные первыми (длиннее), затем bare
    pairs.sort(key=lambda p: (-len(p[0]), p[0]))
    return pairs, warnings


def _fill_document(doc, pairs, values_row1):
    for search, replace in pairs:
        _replace_all(doc, search, replace)
    _sync_user_fields(doc, values_row1 or {})


def normalize_atext_letters_block(block):
    """Нормализация JSON-блока (вызывается из param_codec)."""
    from libre_macros_param_codec import (
        _attach_sheet,
        _normalize_block,
        _normalize_column_token_list,
        canonicalize_column_token_list_for_store,
    )
    from libre_macros_lib import lm_parse_bool_param

    out = _normalize_block(u"подготовка_писем", block)
    out[u"template"] = _u(out.get(u"template")).strip()
    # Лист данных: слева в визарде → sheet; legacy source_sheet тоже ок.
    src = _u(out.get(u"sheet") or out.get(u"source_sheet")).strip()
    out[u"source_sheet"] = src
    out[u"out_dir"] = _u(out.get(u"out_dir")).strip()
    out[u"filename"] = _u(out.get(u"filename")).strip()
    if out[u"filename"] == u"":
        out[u"filename"] = u"Письмо_<<n>>"

    fmt = _u(out.get(u"format") or u"odt").strip().casefold()
    if fmt not in _ATEXT_FORMATS:
        fmt = u"odt"
    out[u"format"] = fmt

    miss = _u(out.get(u"missing_column") or u"empty").strip().casefold()
    if miss not in _MISSING_MODES:
        miss = u"empty"
    out[u"missing_column"] = miss

    for key, default in (
        (u"header_row", 1),
        (u"data_start_row", None),
        (u"data_end_row", None),
        (u"block_rows", None),
        (u"max_files", _DEFAULT_MAX_FILES),
    ):
        raw = out.get(key)
        if raw in (None, u""):
            if default is None:
                out.pop(key, None)
            else:
                out[key] = default
            continue
        try:
            out[key] = int(raw)
        except (TypeError, ValueError):
            if default is None:
                out.pop(key, None)
            else:
                out[key] = default

    # XOR marker / block_rows (как у добавить_объединение: marker побеждает)
    marker_toks = canonicalize_column_token_list_for_store(out.get(u"marker"))
    if marker_toks:
        out[u"marker"] = marker_toks[0]
        out.pop(u"block_rows", None)
    else:
        out.pop(u"marker", None)
        br = out.get(u"block_rows")
        if br is None:
            out[u"block_rows"] = 1
        else:
            try:
                out[u"block_rows"] = max(1, int(br))
            except (TypeError, ValueError):
                out[u"block_rows"] = 1

    cols = _normalize_column_token_list(out.get(u"columns"))
    if cols:
        out[u"columns"] = cols
    else:
        out.pop(u"columns", None)

    out[u"row_filter"] = _u(out.get(u"row_filter")).strip()
    if out[u"row_filter"] == u"":
        out.pop(u"row_filter", None)

    try:
        out[u"open_folder"] = bool(
            lm_parse_bool_param(out.get(u"open_folder"), default=False)
        )
    except Exception:
        out[u"open_folder"] = False
    try:
        out[u"overwrite"] = bool(
            lm_parse_bool_param(out.get(u"overwrite"), default=True)
        )
    except Exception:
        out[u"overwrite"] = True

    mf = out.get(u"max_files")
    try:
        out[u"max_files"] = max(1, int(mf if mf is not None else _DEFAULT_MAX_FILES))
    except (TypeError, ValueError):
        out[u"max_files"] = _DEFAULT_MAX_FILES

    return _attach_sheet(out, src)


def expand_atext_letters_blocks_for_sheets(blocks, sheets):
    """
    Блоки с явным sheet/source_sheet — один раз на этот лист.
    Блок без листа («одинаково на всех») — размножить на каждый лист из sheets
    (как у RANGE-финала); без листов слева это даёт лишние прогоны — визард
    требует указать листы слева.
    """
    blocks = blocks or []
    sheet_names = []
    si = 0
    while si < len(sheets or ()):
        sh = sheets[si]
        si = si + 1
        try:
            nm = _u(getattr(sh, u"Name", u"")).strip()
        except Exception:
            nm = u""
        if nm != u"":
            sheet_names.append(nm)

    out = []
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        bi = bi + 1
        if not isinstance(block, dict):
            continue
        src = _u(block.get(u"sheet") or block.get(u"source_sheet")).strip()
        if src != u"":
            b2 = dict(block)
            b2[u"sheet"] = src
            b2[u"source_sheet"] = src
            out.append(b2)
            continue
        if not sheet_names:
            b2 = dict(block)
            out.append(b2)
            continue
        for nm in sheet_names:
            b2 = dict(block)
            b2[u"sheet"] = nm
            b2[u"source_sheet"] = nm
            out.append(b2)
    return out


def run_atext_letters_block(doc, block):
    """
    Выполнить один JSON-блок. Возвращает (status, note):
      status: ok | ошибка | пропуск
    """
    if uno is None:
        return u"ошибка", u"нужен UNO (LibreOffice/AlterOffice)"

    try:
        block = normalize_atext_letters_block(block or {})
    except Exception as err:
        return u"ошибка", _u(err)

    template = _expand_vars(block.get(u"template"), doc=doc)
    out_dir = _expand_vars(block.get(u"out_dir"), doc=doc)
    filename_tpl = _expand_vars(block.get(u"filename"), doc=doc)
    source_sheet = _u(block.get(u"source_sheet")).strip()
    fmt = block.get(u"format") or u"odt"
    missing_mode = block.get(u"missing_column") or u"empty"
    overwrite = bool(block.get(u"overwrite", True))
    max_files = int(block.get(u"max_files") or _DEFAULT_MAX_FILES)
    open_folder = bool(block.get(u"open_folder", False))

    if template == u"":
        return u"ошибка", u"нужен template"
    if not os.path.isfile(template):
        return u"ошибка", u"шаблон не найден: %s" % template
    if source_sheet == u"":
        return u"ошибка", u"нужен лист данных (слева в визарде)"
    if out_dir == u"":
        return u"ошибка", u"нужен out_dir"

    out_dir, derr = _ensure_dir(out_dir)
    if derr:
        return u"ошибка", derr

    sheet = _resolve_sheet(doc, source_sheet)
    if sheet is None:
        return u"ошибка", u"лист не найден: %s" % source_sheet

    header_row = int(block.get(u"header_row") or 1)
    if header_row < 1:
        header_row = 1
    header_row0 = header_row - 1

    data_start = block.get(u"data_start_row")
    if data_start in (None, u""):
        data_start0 = header_row0 + 1
    else:
        data_start0 = max(0, int(data_start) - 1)

    last_used = _used_last_row(sheet, header_row0)
    data_end = block.get(u"data_end_row")
    if data_end in (None, u""):
        data_end0 = last_used
    else:
        data_end0 = max(data_start0, int(data_end) - 1)

    if data_end0 < data_start0:
        return u"пропуск", u"нет строк данных"

    c0 = 0
    c1 = _used_last_col(sheet, header_row0)
    headers = _read_headers(sheet, header_row0, c0, c1)
    if not any(h != u"" for h in headers):
        return u"ошибка", u"пустые заголовки"

    # header_row_range для resolve marker
    try:
        header_row_range = sheet.getCellRangeByPosition(c0, header_row0, c1, header_row0)
    except Exception:
        header_row_range = None

    all_rows = list(range(int(data_start0), int(data_end0) + 1))
    filtered, ferr = _apply_row_filter(
        sheet, headers, all_rows, c0, block.get(u"row_filter")
    )
    if ferr:
        return u"ошибка", ferr

    marker = _u(block.get(u"marker")).strip()
    if marker != u"":
        raw_blocks, merr = _marker_blocks(
            sheet, marker, data_start0, data_end0, header_row_range, c0, c1
        )
        if merr:
            return u"ошибка", merr
        filt_set = set(filtered)
        blocks = []
        for rows in raw_blocks or []:
            kept = [r for r in rows if r in filt_set]
            if kept:
                blocks.append(kept)
    else:
        br = int(block.get(u"block_rows") or 1)
        blocks = _chunk_rows(filtered, br)

    if not blocks:
        return u"пропуск", u"нет блоков после фильтра"

    allowed = _allowed_columns_set(block.get(u"columns"), headers, None)

    ok_n = 0
    err_n = 0
    warn_notes = []
    file_paths = []
    bi = 0
    while bi < len(blocks) and ok_n + err_n < max_files:
        row0s = blocks[bi]
        file_n = bi + 1
        phys_rows = [r + 1 for r in row0s]
        try:
            matrix = _row_matrix(sheet, headers, row0s, c0)
            pairs, warns = _collect_replacements(
                matrix, file_n, phys_rows, missing_mode, allowed
            )
            for w in warns:
                if w not in warn_notes:
                    warn_notes.append(w)

            fname_raw, fwarns = _substitute_text(
                filename_tpl, matrix, file_n, phys_rows, missing_mode
            )
            for w in fwarns:
                if w not in warn_notes:
                    warn_notes.append(w)
            stem = _sanitize_filename(fname_raw)
            if not stem.lower().endswith(u".%s" % fmt):
                out_path = os.path.join(out_dir, u"%s.%s" % (stem, fmt))
            else:
                out_path = os.path.join(out_dir, stem)

            if (not overwrite) and os.path.isfile(out_path):
                err_n = err_n + 1
                warn_notes.append(u"файл существует: %s" % os.path.basename(out_path))
                bi = bi + 1
                continue

            wdoc, oerr = _load_template(template)
            if oerr:
                err_n = err_n + 1
                warn_notes.append(u"блок %d: %s" % (file_n, oerr))
                bi = bi + 1
                continue
            try:
                values_row1 = matrix[0] if matrix else {}
                if allowed is not None:
                    from libre_macros_lib import lm_identity_key

                    values_row1 = {
                        h: v
                        for h, v in values_row1.items()
                        if lm_identity_key(h) in allowed
                    }
                _fill_document(wdoc, pairs, values_row1)
                sok, serr = _store_doc(wdoc, out_path, overwrite=overwrite)
                if not sok:
                    err_n = err_n + 1
                    warn_notes.append(u"блок %d: %s" % (file_n, serr))
                else:
                    ok_n = ok_n + 1
                    file_paths.append(out_path)
            finally:
                _close_doc(wdoc)
        except Exception as err:
            err_n = err_n + 1
            warn_notes.append(u"блок %d: %s" % (file_n, err))
            if missing_mode == u"fail":
                break
        bi = bi + 1

    if open_folder and ok_n > 0:
        _open_folder(out_dir)

    note_parts = [u"файлов=%d" % ok_n, u"ошибок=%d" % err_n, u"dir=%s" % out_dir]
    if warn_notes:
        note_parts.append(u"; ".join(warn_notes[:8]))
        if len(warn_notes) > 8:
            note_parts.append(u"…ещё %d" % (len(warn_notes) - 8))
    note = u"; ".join(note_parts)
    if err_n > 0 and ok_n == 0:
        return u"ошибка", note
    if ok_n > 0:
        return u"ok", note
    return u"пропуск", note


def run_atext_letters_blocks(doc, blocks):
    if not blocks:
        return u"пропуск", u"пустая C"
    notes = []
    ok_n = 0
    err_n = 0
    skip_n = 0
    bi = 0
    while bi < len(blocks):
        st, note = run_atext_letters_block(doc, blocks[bi])
        notes.append(u"блок %d: %s" % (bi + 1, note))
        if st == u"ok":
            ok_n = ok_n + 1
        elif st == u"ошибка":
            err_n = err_n + 1
        else:
            skip_n = skip_n + 1
        bi = bi + 1
    if err_n > 0 and ok_n == 0:
        return u"ошибка", u"; ".join(notes)
    if ok_n > 0:
        return u"ok", u"; ".join(notes)
    return u"пропуск", u"; ".join(notes)
