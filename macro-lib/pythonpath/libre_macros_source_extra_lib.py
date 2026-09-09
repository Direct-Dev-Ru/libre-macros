# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.717"
"""
Доп. параметры источника (книга ODS / XLS / XLSX / XLSM): prepare_source_in_memory и субключи.

CSV-ключи — в libre_macros_csv_source_lib.py; здесь только book-extra.
"""

try:
    unicode
except NameError:
    unicode = str

import os

SOURCE_EXTRA_DEFAULT_CHUNK_ROWS = 50000

# Имя глобальной переменной / fallback-литерал ключа для lm1: паролей в JSON источника.
ENCRYPT_PASSWORDS_DEFAULT = u"encrypt_passwords_default"

_PREPARE_IN_MEMORY_KEYS = (
    u"prepare_source_in_memory",
    u"подготовка_источника_в_памяти",
    u"prepare_in_memory",
    u"source_in_memory",
    u"матрица_в_памяти",
)

_VALUES_MODE_KEYS = (u"values_mode", u"cell_values", u"значения")
_MERGED_CELLS_KEYS = (u"merged_cells", u"merges", u"объединения")
_TRIM_EMPTY_KEYS = (u"trim_empty", u"trim_trailing_empty", u"обрезать_пустые")
_LAST_ROW_COLUMN_KEYS = (
    u"last_row_column",
    u"last_data_row_column",
    u"analysis_column",
    u"колонка_анализа_последней_строки",
    u"колонка_последней_строки",
    u"значащая_колонка",
)
_CONVERT_TYPES_KEYS = (
    u"convert_types",
    u"types",
    u"типы",
    u"convert_numbers_dates",
    u"convert_where",
)
_CHUNK_ROWS_KEYS = (
    u"chunk_rows",
    u"memory_chunk_rows",
    u"max_chunk_rows",
    u"лимит_строк_чанка",
)
_EXCLUDE_MASK_KEYS = (
    u"exclude_mask",
    u"MERGE_SOURCE_EXCLUDE_MASK",
    u"source_exclude_mask",
    u"исключить",
    u"маски_исключения",
)
_FILE_PICK_KEYS = (
    u"file_pick",
    u"доп_фильтр_файлов",
    u"source_file_pick",
    u"pick_files",
)
_FILE_PICK_SCOPE_KEYS = (
    u"file_pick_scope",
    u"область_фильтра",
    u"file_pick_area",
)
_FILE_PICK_LAMBDA_KEYS = (
    u"file_pick_lambda",
    u"лямбда_фильтра_файлов",
    u"file_lambda",
)

_FILE_PICK_CANON = frozenset(
    (
        u"all",
        u"newest_ctime",
        u"oldest_ctime",
        u"newest_mtime",
        u"oldest_mtime",
        u"lambda",
    )
)
_FILE_PICK_SCOPE_CANON = frozenset((u"all", u"per_subdir"))
_FILE_PICK_ALIASES = {
    u"все": u"all",
    u"*": u"all",
    u"новый_по_дате_создания": u"newest_ctime",
    u"старый_по_дате_создания": u"oldest_ctime",
    u"новый_по_дате_модификации": u"newest_mtime",
    u"старый_по_дате_модификации": u"oldest_mtime",
    u"лямбда": u"lambda",
    u"lambda_filter": u"lambda",
}
_FILE_PICK_SCOPE_ALIASES = {
    u"все": u"all",
    u"все_подкаталоги": u"all",
    u"вместе": u"all",
    u"подкаталог": u"per_subdir",
    u"в_подкаталоге": u"per_subdir",
    u"per_dir": u"per_subdir",
    u"each_subdir": u"per_subdir",
}
_SKIP_MISSING_SHEETS_KEYS = (
    u"skip_missing_sheets",
    u"пропускать_отсутствующие_листы",
    u"allow_missing_sheets",
    u"не_проверять_листы",
    u"skip_sheet_check",
)
_SHEET_UNPROTECT_PASSWORD_KEYS = (
    u"sheet_unprotect_password",
    u"пароль_снятия_защиты",
    u"unprotect_password",
    u"sheet_protect_password",
)

_VALUES_MODE_CANON = frozenset((u"cached", u"formulas", u"prefer_cached"))
_MERGED_CELLS_CANON = frozenset((u"top_left", u"fill"))
_CONVERT_TYPES_CANON = frozenset((u"in_memory", u"uno", u"no", u"keep_text"))


def parse_exclude_mask_list(raw):
    """Разбор масок исключения: строка через ,/; или list → list[str]."""
    if raw is None:
        return []
    if isinstance(raw, (list, tuple)):
        out = []
        i = 0
        while i < len(raw):
            t = unicode(raw[i] or u"").strip()
            if t != u"":
                out.append(t)
            i = i + 1
        return out
    text = unicode(raw or u"").strip()
    if text == u"":
        return []
    parts = []
    for chunk in text.replace(u";", u",").split(u","):
        t = chunk.strip()
        if t != u"":
            parts.append(t)
    return parts


def source_extra_exclude_mask_text(extra):
    """Текст масок исключения из dict доп. параметров (для UI / JSON)."""
    raw = _extra_key(extra, _EXCLUDE_MASK_KEYS)
    if raw is None:
        return u""
    if isinstance(raw, (list, tuple)):
        return u",".join(parse_exclude_mask_list(raw))
    return unicode(raw or u"").strip()


def source_extra_exclude_masks(extra):
    """Список масок исключения из dict доп. параметров (пусто → [])."""
    return parse_exclude_mask_list(source_extra_exclude_mask_text(extra))


def source_extra_file_pick(extra):
    """Режим доп. фильтра файлов: all|newest_*|oldest_*|lambda."""
    raw = _extra_key(extra, _FILE_PICK_KEYS)
    if raw is None:
        return u"all"
    return _normalize_choice(raw, _FILE_PICK_CANON, u"all", aliases=_FILE_PICK_ALIASES)


def source_extra_file_pick_scope(extra):
    """Область фильтра: all | per_subdir."""
    raw = _extra_key(extra, _FILE_PICK_SCOPE_KEYS)
    if raw is None:
        return u"all"
    return _normalize_choice(
        raw, _FILE_PICK_SCOPE_CANON, u"all", aliases=_FILE_PICK_SCOPE_ALIASES
    )


def source_extra_file_pick_lambda(extra):
    """Текст лямбды доп. фильтра файлов (пусто → '')."""
    raw = _extra_key(extra, _FILE_PICK_LAMBDA_KEYS)
    if raw is None:
        return u""
    return unicode(raw or u"").strip()


def _file_stat_times(path, log_fn=None):
    """(mtime, ctime) unix; при ошибке (0, 0). ctime недоступен → mtime + журнал."""
    try:
        st = os.stat(path)
        mtime = float(getattr(st, "st_mtime", 0) or 0)
        ctime = float(getattr(st, "st_ctime", 0) or 0)
        if ctime <= 0:
            ctime = mtime
            if log_fn and mtime > 0:
                try:
                    log_fn(
                        u"ctime недоступен для %s — fallback mtime"
                        % os.path.basename(path)
                    )
                except Exception:
                    pass
        return (mtime, ctime)
    except Exception:
        return (0.0, 0.0)


def _file_pick_group_key(path):
    try:
        return os.path.dirname(os.path.abspath(path))
    except Exception:
        return u"."


def _pick_extreme_path(paths, mode, log_fn=None):
    """Один путь с экстремумом mtime/ctime; пустой список → []."""
    if not paths:
        return []
    want_newest = mode.startswith(u"newest_")
    use_ctime = mode.endswith(u"ctime")
    best = None
    best_t = None
    for p in paths:
        mtime, ctime = _file_stat_times(p, log_fn=log_fn if use_ctime else None)
        t = ctime if use_ctime else mtime
        if best is None:
            best = p
            best_t = t
            continue
        if want_newest:
            if t > best_t or (t == best_t and unicode(p) < unicode(best)):
                best = p
                best_t = t
        else:
            if t < best_t or (t == best_t and unicode(p) < unicode(best)):
                best = p
                best_t = t
    return [best] if best is not None else []


def apply_source_extra_file_pick(paths, extra, log_fn=None):
    """
    После exclude_mask: сузить список файлов по file_pick / scope / lambda.

    date-режимы → 1 файл на scope; lambda → N файлов (filter).
    """
    items = []
    for p in paths or []:
        s = unicode(p or u"").strip()
        if s != u"":
            items.append(s)
    mode = source_extra_file_pick(extra)
    if mode == u"all" or not items:
        return items
    scope = source_extra_file_pick_scope(extra)
    if mode == u"lambda":
        code = source_extra_file_pick_lambda(extra)
        if code == u"":
            return items
        try:
            from libre_macros_sheet_filter_lib import compile_file_pick_lambda

            fn = compile_file_pick_lambda(code)
        except Exception as err:
            if log_fn:
                try:
                    log_fn(u"file_pick lambda compile: %s" % err)
                except Exception:
                    pass
            else:
                _log(u"file_pick lambda compile: %s" % err)
            return items
        if fn is None:
            return items

        def _filter_group(group):
            out = []
            for p in group:
                try:
                    mtime, ctime = _file_stat_times(p, log_fn=log_fn)
                    size = 0
                    try:
                        size = int(os.path.getsize(p))
                    except Exception:
                        size = 0
                    ctx = {
                        u"path": p,
                        u"путь": p,
                        u"name": os.path.basename(p),
                        u"имя": os.path.basename(p),
                        u"dir": os.path.dirname(p),
                        u"каталог": os.path.dirname(p),
                        u"mtime": mtime,
                        u"ctime": ctime,
                        u"size": size,
                        u"размер": size,
                    }
                    if fn(ctx):
                        out.append(p)
                except Exception:
                    pass
            return out

        if scope == u"per_subdir":
            groups = {}
            order_keys = []
            for p in items:
                gk = _file_pick_group_key(p)
                if gk not in groups:
                    groups[gk] = []
                    order_keys.append(gk)
                groups[gk].append(p)
            result = []
            for gk in order_keys:
                result.extend(_filter_group(groups[gk]))
            return result
        return _filter_group(items)

    # date modes
    if scope == u"per_subdir":
        groups = {}
        order_keys = []
        for p in items:
            gk = _file_pick_group_key(p)
            if gk not in groups:
                groups[gk] = []
                order_keys.append(gk)
            groups[gk].append(p)
        result = []
        for gk in order_keys:
            result.extend(_pick_extreme_path(groups[gk], mode, log_fn=log_fn))
        return result
    return _pick_extreme_path(items, mode, log_fn=log_fn)


def source_extra_skip_missing_sheets(extra):
    """
    True — в режиме «Текущие листы» / Эта_книга не останавливать сбор,
    если в «Листы» указаны имена/шаблоны без совпадений в книге.
    Отсутствующие пропускаются; совпавшие вовлекаются.
    """
    raw = _extra_key(extra, _SKIP_MISSING_SHEETS_KEYS)
    if raw is None:
        return False
    return _parse_source_extra_bool(raw, default=False)


def resolve_encrypt_passwords_key():
    """
    Ключ для lm1: паролей источника (sheet_unprotect_password).

    Если в глобальных переменных задана переменная с именем
    encrypt_passwords_default — берём её значение; иначе литерал
    ENCRYPT_PASSWORDS_DEFAULT.
    """
    name = ENCRYPT_PASSWORDS_DEFAULT
    try:
        name_cf = name.casefold()
    except Exception:
        name_cf = name.lower()
    try:
        from libre_macros_global_settings_lib import resolved_global_variables

        for item in resolved_global_variables() or []:
            nm = unicode(item.get(u"name") or u"").strip()
            if nm == u"":
                continue
            try:
                nm_cf = nm.casefold()
            except Exception:
                nm_cf = nm.lower()
            if nm_cf != name_cf:
                continue
            val = unicode(item.get(u"value") if item.get(u"value") is not None else u"")
            if val != u"":
                return val
            break
    except Exception as exc:
        _log("resolve_encrypt_passwords_key: %s" % exc)
    return ENCRYPT_PASSWORDS_DEFAULT


def _decrypt_source_extra_secret(stored):
    """Расшифровка lm1:-blob; без префикса — legacy plain text."""
    stored = unicode(stored or u"").strip()
    if stored == u"":
        return u""
    try:
        from libre_macros_normalize_lib import CRYPTO_PREFIX, _has_tag_prefix, op_decrypt

        if not _has_tag_prefix(stored, CRYPTO_PREFIX):
            return stored
        keys = []
        primary = resolve_encrypt_passwords_key()
        keys.append(primary)
        if primary != ENCRYPT_PASSWORDS_DEFAULT:
            keys.append(ENCRYPT_PASSWORDS_DEFAULT)
        last_err = None
        for key in keys:
            try:
                return op_decrypt(stored, key=key)
            except Exception as exc:
                last_err = exc
                continue
        _log("decrypt sheet_unprotect_password failed: %s" % last_err)
        return u""
    except Exception as exc:
        _log("decrypt import/parse failed: %s" % exc)
    return stored


def _log(msg):
    try:
        print("[source_extra] %s" % msg)
    except Exception:
        pass


def encrypt_source_extra_password_for_store(plain):
    """Зашифровать пароль для записи в JSON (sheet_unprotect_password)."""
    plain = unicode(plain or u"").strip()
    if plain == u"":
        return u""
    try:
        from libre_macros_normalize_lib import op_encrypt

        return op_encrypt(plain, key=resolve_encrypt_passwords_key())
    except Exception:
        return plain


def source_extra_sheet_unprotect_password(extra):
    """Пароль снятия защиты листов (режим «Текущие листы»); пусто — без пароля."""
    raw = _extra_key(extra, _SHEET_UNPROTECT_PASSWORD_KEYS)
    if raw is None:
        return u""
    return _decrypt_source_extra_secret(raw)


def _parse_source_extra_bool(raw, default=False):
    try:
        from libre_macros_csv_source_lib import _parse_source_extra_bool as _csv_bool
        return _csv_bool(raw, default=default)
    except Exception:
        if raw is None:
            return bool(default)
        if isinstance(raw, bool):
            return bool(raw)
        if isinstance(raw, (int, float)) and (not isinstance(raw, bool)):
            return bool(int(raw))
        s = unicode(raw or u"").strip().lower()
        if s in (u"1", u"true", u"yes", u"да", u"on"):
            return True
        if s in (u"0", u"false", u"no", u"нет", u"off"):
            return False
        return bool(default)


def _extra_key(extra, keys):
    if not isinstance(extra, dict):
        return None
    for key in keys:
        if key in extra:
            return extra.get(key)
    return None


def _normalize_choice(raw, canon_set, default, aliases=None):
    if raw is None:
        return default
    text = unicode(raw or u"").strip().lower()
    if text == u"":
        return default
    if aliases and text in aliases:
        text = aliases[text]
    if text in canon_set:
        return text
    return default


def source_extra_prepare_in_memory(extra):
    raw = _extra_key(extra, _PREPARE_IN_MEMORY_KEYS)
    if raw is None:
        return False
    return _parse_source_extra_bool(raw, default=False)


def source_extra_values_mode(extra):
    raw = _extra_key(extra, _VALUES_MODE_KEYS)
    return _normalize_choice(raw, _VALUES_MODE_CANON, u"cached")


def source_extra_merged_cells(extra):
    raw = _extra_key(extra, _MERGED_CELLS_KEYS)
    return _normalize_choice(raw, _MERGED_CELLS_CANON, u"top_left")


def source_extra_trim_empty(extra, prepare_active=False):
    raw = _extra_key(extra, _TRIM_EMPTY_KEYS)
    if raw is None:
        return bool(prepare_active)
    return _parse_source_extra_bool(raw, default=bool(prepare_active))


def source_extra_last_row_column(extra):
    """Колонка анализа последней значащей строки: 2 / B / Имя_заголовка."""
    raw = _extra_key(extra, _LAST_ROW_COLUMN_KEYS)
    if raw is None:
        return u""
    return unicode(raw or u"").strip()


def source_extra_convert_types(extra):
    raw = _extra_key(extra, _CONVERT_TYPES_KEYS)
    aliases = {
        u"yes": u"in_memory",
        u"true": u"in_memory",
        u"да": u"in_memory",
        u"auto": u"in_memory",
    }
    return _normalize_choice(raw, _CONVERT_TYPES_CANON, u"in_memory", aliases=aliases)


def source_extra_chunk_rows(extra):
    raw = _extra_key(extra, _CHUNK_ROWS_KEYS)
    if raw is None:
        return int(SOURCE_EXTRA_DEFAULT_CHUNK_ROWS)
    try:
        n = int(raw)
    except (TypeError, ValueError):
        return int(SOURCE_EXTRA_DEFAULT_CHUNK_ROWS)
    if n < 1:
        return int(SOURCE_EXTRA_DEFAULT_CHUNK_ROWS)
    return int(n)


def _encode_file_pick_fields(extra, payload):
    """Добавить file_pick / scope / lambda в payload, если не default."""
    if not isinstance(extra, dict):
        return
    mode = source_extra_file_pick(extra)
    if mode != u"all":
        payload[u"file_pick"] = mode
        scope = source_extra_file_pick_scope(extra)
        if scope != u"all":
            payload[u"file_pick_scope"] = scope
        if mode == u"lambda":
            lam = source_extra_file_pick_lambda(extra)
            if lam != u"":
                payload[u"file_pick_lambda"] = lam
    else:
        # явный all не пишем; но если в исходном dict были ключи — сохраняем не-default scope/lambda только при mode!=all
        pass


def encode_source_extra_current_book_json(extra):
    """Сериализация доп. параметров для источника «Эта_книга» (skip_missing, пароль)."""
    import json

    if not isinstance(extra, dict):
        extra = {}
    skip_miss = source_extra_skip_missing_sheets(extra)
    unprot_pwd = source_extra_sheet_unprotect_password(extra)
    last_row_col = source_extra_last_row_column(extra)
    payload = {}
    if skip_miss:
        payload[u"skip_missing_sheets"] = True
    if last_row_col != u"":
        payload[u"last_row_column"] = last_row_col
    if unprot_pwd != u"":
        payload[u"sheet_unprotect_password"] = encrypt_source_extra_password_for_store(unprot_pwd)
    _encode_file_pick_fields(extra, payload)
    if not payload:
        return u""
    try:
        return unicode(json.dumps(payload, ensure_ascii=False, separators=(u",", u":")))
    except Exception:
        return u""


def encode_source_extra_book_json(extra, prepare_active=False):
    """Сериализация book-extra dict в JSON (только ключи книги)."""
    import json

    if not isinstance(extra, dict):
        extra = {}
    prep = source_extra_prepare_in_memory(extra)
    excl = source_extra_exclude_mask_text(extra)
    skip_miss = source_extra_skip_missing_sheets(extra)
    unprot_pwd = source_extra_sheet_unprotect_password(extra)
    if not prep and not prepare_active:
        payload = {}
        if prep is False and _extra_key(extra, _PREPARE_IN_MEMORY_KEYS) is not None:
            payload[u"prepare_source_in_memory"] = False
        if excl != u"":
            payload[u"exclude_mask"] = excl
        last_row_col = source_extra_last_row_column(extra)
        if last_row_col != u"":
            payload[u"last_row_column"] = last_row_col
        if skip_miss:
            payload[u"skip_missing_sheets"] = True
        if unprot_pwd != u"":
            payload[u"sheet_unprotect_password"] = encrypt_source_extra_password_for_store(
                unprot_pwd
            )
        _encode_file_pick_fields(extra, payload)
        if not payload:
            return u""
        try:
            return unicode(json.dumps(payload, ensure_ascii=False, separators=(u",", u":")))
        except Exception:
            return u""
    payload = {u"prepare_source_in_memory": bool(prep)}
    if prep:
        payload[u"values_mode"] = source_extra_values_mode(extra)
        payload[u"merged_cells"] = source_extra_merged_cells(extra)
        payload[u"trim_empty"] = source_extra_trim_empty(extra, prepare_active=True)
        last_row_col = source_extra_last_row_column(extra)
        if last_row_col != u"":
            payload[u"last_row_column"] = last_row_col
        payload[u"convert_types"] = source_extra_convert_types(extra)
        payload[u"chunk_rows"] = source_extra_chunk_rows(extra)
    if excl != u"":
        payload[u"exclude_mask"] = excl
    elif not prep:
        last_row_col = source_extra_last_row_column(extra)
        if last_row_col != u"":
            payload[u"last_row_column"] = last_row_col
    if skip_miss:
        payload[u"skip_missing_sheets"] = True
    if unprot_pwd != u"":
        payload[u"sheet_unprotect_password"] = encrypt_source_extra_password_for_store(unprot_pwd)
    _encode_file_pick_fields(extra, payload)
    try:
        return unicode(json.dumps(payload, ensure_ascii=False, separators=(u",", u":")))
    except Exception:
        return u""


def encode_source_extra_xml_json(extra):
    """Сериализация доп. параметров XML-источника: exclude_mask + file_pick*."""
    import json

    if not isinstance(extra, dict):
        extra = {}
    payload = {}
    excl = source_extra_exclude_mask_text(extra)
    if excl != u"":
        payload[u"exclude_mask"] = excl
    _encode_file_pick_fields(extra, payload)
    if not payload:
        return u""
    try:
        return unicode(json.dumps(payload, ensure_ascii=False, separators=(u",", u":")))
    except Exception:
        return u""
