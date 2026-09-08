# -*- coding: utf-8 -*-
"""Настройки группы макросов todo_task (JSON в ~/.config/libre-macros/todo_task/)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.711"
import json
import os
import sys

try:
    unicode
except NameError:
    unicode = str

try:
    unichr
except NameError:
    unichr = chr

TODO_TASK_SETTINGS_FILE = u"todo_task_settings.json"
TODO_TASK_SETTINGS_SUBDIR = u"todo_task"
TODO_TASK_SETTINGS_JSON_VER = 2
SCOPE_PRIVATE = u"private"
SCOPE_GLOBAL = u"global"

DEFAULT_COLORIZE_ALL_SHEETS = False
DEFAULT_COLORIZE_SHEET_PATTERN = u"Задачи*"
DEFAULT_SORT_MODE = u"assignee"
DEFAULT_SORT_DUE_DIRECTION = u"old_first"
DEFAULT_PROTECT_PASSWORD = u"task"
DEFAULT_OWNER_ROLE = u"manager"
DEFAULT_CONSOLE_LOG = True
DEFAULT_DIALOG_FONT_PT = 11
DIALOG_FONT_PT_MIN = 8
DIALOG_FONT_PT_MAX = 24
DEFAULT_LAYOUT_TEMPLATE_SHEET = u"Задачи.01"
DEFAULT_TASK_SHEET_PATTERN = u"Задачи*"
DEFAULT_SUMMARY_SHEET_PATTERN = u"Свод*задач*"
DEFAULT_MERGE_SOURCE_SHEET = u"задачи_сотрудников"
DEFAULT_MERGE_UPDATE_EXISTING_ONLY = False
DEFAULT_MERGE_CHECK_MODIFIED = False
DEFAULT_MERGE_ACCUMULATE_COMMENTS = False
DEFAULT_APPLY_WIDTHS_TO_SUMMARY = False
_PROTECT_ENC_PREFIX = u"x1:"

# Кэш флага логов: None = ещё не читали; (book_path, enabled).
_CONSOLE_LOG_CACHE = None
# Кэш JSON-хранилища: (path, mtime, store). Сбрасывается при записи и при смене mtime.
_STORE_CACHE = None


def _as_bool(value, default=True):
    if value is None:
        return bool(default)
    if isinstance(value, bool):
        return value
    try:
        if isinstance(value, (int, float)):
            return bool(int(value))
    except Exception:
        pass
    text = unicode(value).strip()
    if not text:
        return bool(default)
    low = text.casefold() if hasattr(text, "casefold") else text.lower()
    if low in (u"1", u"true", u"yes", u"y", u"on", u"да", u"д"):
        return True
    if low in (u"0", u"false", u"no", u"n", u"off", u"нет", u"н"):
        return False
    return bool(default)


def invalidate_console_log_cache():
    global _CONSOLE_LOG_CACHE
    _CONSOLE_LOG_CACHE = None


def invalidate_todo_task_settings_cache():
    """Сбросить кэш JSON и флага логов (после макроса настроек / записи на диск)."""
    global _STORE_CACHE
    _STORE_CACHE = None
    invalidate_console_log_cache()


def get_console_log_enabled(settings=None, doc=None):
    """Подробные логи todo_task в консоль (по умолчанию включены)."""
    global _CONSOLE_LOG_CACHE
    path = u""
    try:
        path = get_doc_book_path(doc)
    except Exception:
        path = u""
    use_cache = settings is None
    if use_cache and _CONSOLE_LOG_CACHE is not None:
        try:
            cached_path, cached_on = _CONSOLE_LOG_CACHE
            if cached_path == path:
                return bool(cached_on)
        except Exception:
            pass
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    gen = (settings or {}).get(u"general") or {}
    enabled = _as_bool(gen.get(u"console_log"), DEFAULT_CONSOLE_LOG)
    if use_cache:
        _CONSOLE_LOG_CACHE = (path, bool(enabled))
    return bool(enabled)


def normalize_dialog_font_pt(value, default=None):
    """Размер шрифта диалогов (pt), clamp в допустимый диапазон."""
    if default is None:
        default = DEFAULT_DIALOG_FONT_PT
    try:
        pt = int(round(float(value)))
    except (TypeError, ValueError):
        pt = int(default)
    if pt < int(DIALOG_FONT_PT_MIN):
        pt = int(DIALOG_FONT_PT_MIN)
    if pt > int(DIALOG_FONT_PT_MAX):
        pt = int(DIALOG_FONT_PT_MAX)
    return pt


def get_dialog_font_pt(settings=None, doc=None):
    """Размер шрифта GUI серии todo_task (pt)."""
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    gen = (settings or {}).get(u"general") or {}
    return normalize_dialog_font_pt(
        gen.get(u"dialog_font_pt"), DEFAULT_DIALOG_FONT_PT
    )


def _config_base():
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.environ.get(
            "XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config")
        )
    return os.path.join(base, "libre-macros")


def todo_task_settings_dir():
    path = os.path.join(_config_base(), TODO_TASK_SETTINGS_SUBDIR)
    try:
        os.makedirs(path, exist_ok=True)
    except TypeError:
        if not os.path.isdir(path):
            try:
                os.makedirs(path)
            except OSError:
                pass
    except OSError:
        pass
    return path


def todo_task_settings_path():
    return os.path.join(todo_task_settings_dir(), TODO_TASK_SETTINGS_FILE)


def _current_doc(doc=None):
    if doc is not None:
        return doc
    xsc = None
    try:
        import __main__

        xsc = getattr(__main__, "XSCRIPTCONTEXT", None)
    except Exception:
        xsc = None
    if xsc is None:
        try:
            import builtins

            xsc = getattr(builtins, "XSCRIPTCONTEXT", None)
        except Exception:
            xsc = None
    if xsc is not None:
        try:
            return xsc.getDocument()
        except Exception:
            pass
    return None


def _file_url_to_system_path(url):
    text = unicode(url or u"").strip()
    if not text:
        return u""
    low = text[:5].lower() if len(text) >= 5 else text.lower()
    if low != u"file:":
        return text
    try:
        import uno

        converted = uno.fileUrlToSystemPath(text)
        if converted:
            return unicode(converted)
    except Exception:
        pass
    rest = text[5:]
    if rest.startswith(u"///"):
        rest = rest[3:]
        if not (len(rest) >= 2 and rest[1] == u":"):
            rest = u"/" + rest
    elif rest.startswith(u"//"):
        rest = rest[2:]
    try:
        if sys.version_info[0] >= 3:
            from urllib.parse import unquote
        else:
            from urllib import unquote

        rest = unquote(rest)
    except Exception:
        pass
    return unicode(rest)


def normalize_book_path(path_or_url):
    """Абсолютный системный путь книги (ключ by_path)."""
    text = unicode(path_or_url or u"").strip()
    if not text:
        return u""
    low = text.lower()
    if low.startswith(u"private:"):
        return u""
    text = _file_url_to_system_path(text)
    text = unicode(text or u"").strip()
    if not text:
        return u""
    try:
        text = os.path.abspath(os.path.normpath(text))
    except Exception:
        pass
    if os.name == "nt":
        text = text.replace(u"/", u"\\")
    return unicode(text)


def get_doc_book_path(doc=None):
    """Системный путь текущей книги или пустая строка (не сохранена)."""
    doc = _current_doc(doc)
    if doc is None:
        return u""
    url = u""
    try:
        url = unicode(doc.getURL() if hasattr(doc, "getURL") else getattr(doc, "URL", u"") or u"").strip()
    except Exception:
        url = u""
    if not url:
        return u""
    return normalize_book_path(url)


def _path_key_match(left, right):
    a = unicode(left or u"")
    b = unicode(right or u"")
    if not a or not b:
        return False
    if a == b:
        return True
    if os.name == "nt":
        al = a.casefold() if hasattr(a, "casefold") else a.lower()
        bl = b.casefold() if hasattr(b, "casefold") else b.lower()
        return al == bl
    return False


def _find_private_key(by_path, path):
    if not path or not isinstance(by_path, dict):
        return None
    if path in by_path:
        return path
    for key in by_path:
        if _path_key_match(key, path):
            return key
    return None


def _looks_like_settings_slot(obj):
    if not isinstance(obj, dict):
        return False
    for key in (u"colorize", u"sort", u"general", u"layout", u"merge"):
        if key in obj:
            return True
    return False


def _looks_like_settings_store(obj):
    if not isinstance(obj, dict):
        return False
    if isinstance(obj.get(u"by_path"), dict):
        return True
    glob = obj.get(u"global")
    return isinstance(glob, dict) and (
        _looks_like_settings_slot(glob) or u"version" in glob
    )


def default_todo_task_store():
    return {
        u"version": TODO_TASK_SETTINGS_JSON_VER,
        u"global": default_todo_task_settings(),
        u"by_path": {},
    }


def normalize_todo_task_store(obj):
    store = default_todo_task_store()
    if not isinstance(obj, dict):
        return store
    if _looks_like_settings_store(obj):
        store[u"global"] = normalize_todo_task_settings(obj.get(u"global"))
        raw_paths = obj.get(u"by_path")
        if isinstance(raw_paths, dict):
            by_path = {}
            for key, val in raw_paths.items():
                nk = normalize_book_path(key)
                if not nk:
                    nk = unicode(key or u"").strip()
                if not nk:
                    continue
                by_path[nk] = normalize_todo_task_settings(val)
            store[u"by_path"] = by_path
        return store
    if _looks_like_settings_slot(obj):
        store[u"global"] = normalize_todo_task_settings(obj)
        return store
    return store


def load_todo_task_store():
    path = todo_task_settings_path()
    mtime = None
    try:
        if os.path.isfile(path):
            mtime = os.path.getmtime(path)
    except Exception:
        mtime = None
    global _STORE_CACHE
    if _STORE_CACHE is not None:
        try:
            cpath, cmtime, cstore = _STORE_CACHE
            if cpath == path and cmtime == mtime and cstore is not None:
                return cstore
        except Exception:
            pass
    if not os.path.isfile(path):
        store = default_todo_task_store()
        _STORE_CACHE = (path, mtime, store)
        return store
    try:
        with open(path, "rb") as f:
            raw = f.read()
        if not raw:
            store = default_todo_task_store()
            _STORE_CACHE = (path, mtime, store)
            return store
        data = json.loads(raw.decode("utf-8"))
    except Exception as err:
        print(u"todo_task_settings: read %s: %s" % (path, err))
        store = default_todo_task_store()
        _STORE_CACHE = (path, mtime, store)
        return store
    store = normalize_todo_task_store(data)
    _STORE_CACHE = (path, mtime, store)
    return store


def has_private_todo_task_settings(doc=None, book_path=None):
    path = unicode(book_path or u"").strip() or get_doc_book_path(doc)
    if not path:
        return False
    store = load_todo_task_store()
    return _find_private_key(store.get(u"by_path"), path) is not None


def _settings_codec_key():
    """Ключ XOR; в исходнике не хранится одной строкой."""
    # mask 0xA5, payload = key.encode xor mask
    blob = (
        0xCD, 0xD7, 0xC0, 0xCB, 0xFA, 0xC8, 0xC0, 0xCB, 0xDC, 0xC4,
        0xFA, 0xD3, 0xDF, 0xC9, 0xCA, 0xC8, 0xC4, 0xC0, 0xD6, 0xCD,
        0xFA, 0x97, 0x95, 0x97, 0x93,
    )
    mask = 0xA5
    chars = []
    i = 0
    while i < len(blob):
        chars.append(unichr(int(blob[i]) ^ mask))
        i = i + 1
    return u"".join(chars)


def _byte_at(buf, i):
    v = buf[i]
    if isinstance(v, int):
        return int(v) & 0xFF
    return ord(v) & 0xFF


def _xor_utf8(text, key):
    data = unicode(text or u"").encode("utf-8")
    key_b = unicode(key or u"").encode("utf-8")
    if not key_b:
        return data
    out = bytearray(len(data))
    i = 0
    klen = len(key_b)
    while i < len(data):
        out[i] = _byte_at(data, i) ^ _byte_at(key_b, i % klen)
        i = i + 1
    return bytes(out)


def _bytes_to_hex(raw):
    try:
        import binascii

        hx = binascii.hexlify(raw)
        if not isinstance(hx, str):
            hx = hx.decode("ascii")
        return unicode(hx)
    except Exception:
        pass
    parts = []
    i = 0
    while i < len(raw):
        parts.append(u"%02x" % (raw[i] if isinstance(raw[i], int) else ord(raw[i])))
        i = i + 1
    return u"".join(parts)


def _hex_to_bytes(text):
    s = unicode(text or u"").strip()
    try:
        import binascii

        raw = binascii.unhexlify(s.encode("ascii"))
        return raw
    except Exception:
        pass
    out = bytearray()
    i = 0
    while i + 1 < len(s):
        try:
            out.append(int(s[i : i + 2], 16))
        except Exception:
            return b""
        i = i + 2
    return bytes(out)


def encode_protect_password(plain):
    text = unicode(plain if plain is not None else u"")
    if text == u"":
        return u""
    raw = _xor_utf8(text, _settings_codec_key())
    return _PROTECT_ENC_PREFIX + _bytes_to_hex(raw)


def decode_protect_password(stored):
    s = unicode(stored if stored is not None else u"")
    if s == u"":
        return u""
    if not s.startswith(_PROTECT_ENC_PREFIX):
        return s
    raw = _hex_to_bytes(s[len(_PROTECT_ENC_PREFIX) :])
    if not raw:
        return u""
    key = _settings_codec_key().encode("utf-8")
    out = bytearray(len(raw))
    i = 0
    klen = len(key)
    while i < len(raw):
        out[i] = _byte_at(raw, i) ^ _byte_at(key, i % klen)
        i = i + 1
    try:
        return bytes(out).decode("utf-8")
    except Exception:
        return u""


def normalize_colorize_settings(obj):
    """Параметры макроса «раскрасить» (todo_task_colorize)."""
    out = {
        u"all_sheets": bool(DEFAULT_COLORIZE_ALL_SHEETS),
        u"sheet_pattern": unicode(DEFAULT_COLORIZE_SHEET_PATTERN),
    }
    if not isinstance(obj, dict):
        return out
    if obj.get("all_sheets") in (True, 1, u"1", u"true", u"yes", u"да"):
        out[u"all_sheets"] = True
    else:
        out[u"all_sheets"] = False
    pat = unicode(obj.get("sheet_pattern") or u"").strip()
    if pat != u"":
        out[u"sheet_pattern"] = pat
    return out


def _sort_choice_codes(attr_name, fallback):
    try:
        import todo_task_cfg as _cfg

        items = getattr(_cfg, attr_name, ()) or ()
        out = []
        i = 0
        while i < len(items):
            code = unicode(items[i][0] or u"").strip()
            if code:
                out.append(code)
            i = i + 1
        if out:
            return tuple(out)
    except Exception:
        pass
    return fallback


def normalize_sort_settings(obj):
    """Параметры макроса «сортировать» (todo_task_sort)."""
    mode_codes = _sort_choice_codes(
        "SORT_MODE_CHOICES",
        (u"assignee", u"due", u"overdue", u"priority"),
    )
    due_codes = _sort_choice_codes(
        "SORT_DUE_DIR_CHOICES",
        (u"old_first", u"new_first"),
    )
    out = {
        u"mode": unicode(DEFAULT_SORT_MODE),
        u"due_direction": unicode(DEFAULT_SORT_DUE_DIRECTION),
    }
    if mode_codes:
        out[u"mode"] = mode_codes[0]
    if due_codes:
        out[u"due_direction"] = due_codes[0]
    if not isinstance(obj, dict):
        return out
    mode = unicode(obj.get("mode") or u"").strip()
    if mode in mode_codes:
        out[u"mode"] = mode
    due = unicode(obj.get("due_direction") or u"").strip()
    if due in due_codes:
        out[u"due_direction"] = due
    return out


def _owner_role_codes():
    return _sort_choice_codes(
        "OWNER_ROLE_CHOICES",
        (u"manager", u"employee"),
    )


def normalize_merge_settings(obj):
    """Опции слияния с сводом сотрудников (todo_task_merge)."""
    out = {
        u"update_existing_only": bool(DEFAULT_MERGE_UPDATE_EXISTING_ONLY),
        u"check_modified": bool(DEFAULT_MERGE_CHECK_MODIFIED),
        u"accumulate_comments": bool(DEFAULT_MERGE_ACCUMULATE_COMMENTS),
    }
    if not isinstance(obj, dict):
        return out
    out[u"update_existing_only"] = _as_bool(
        obj.get(u"update_existing_only"), DEFAULT_MERGE_UPDATE_EXISTING_ONLY
    )
    out[u"check_modified"] = _as_bool(
        obj.get(u"check_modified"), DEFAULT_MERGE_CHECK_MODIFIED
    )
    out[u"accumulate_comments"] = _as_bool(
        obj.get(u"accumulate_comments"), DEFAULT_MERGE_ACCUMULATE_COMMENTS
    )
    return out


def normalize_layout_settings(obj, header_names=None):
    """
    Размеры колонок и шаблоны имён листов.
    column_widths — список {name, width_mm}; name — текст заголовка.
    apply_to_summary — если False, ширины только на листы задач.
    """
    out = {
        u"template_sheet": unicode(DEFAULT_LAYOUT_TEMPLATE_SHEET),
        u"task_sheet_pattern": unicode(DEFAULT_TASK_SHEET_PATTERN),
        u"summary_sheet_pattern": unicode(DEFAULT_SUMMARY_SHEET_PATTERN),
        u"merge_source_sheet": unicode(DEFAULT_MERGE_SOURCE_SHEET),
        u"apply_to_summary": bool(DEFAULT_APPLY_WIDTHS_TO_SUMMARY),
        u"column_widths": [],
    }
    if isinstance(obj, dict):
        ts = unicode(obj.get(u"template_sheet") or u"").strip()
        if ts:
            out[u"template_sheet"] = ts
        tp = unicode(obj.get(u"task_sheet_pattern") or u"").strip()
        if tp:
            out[u"task_sheet_pattern"] = tp
        sp = unicode(obj.get(u"summary_sheet_pattern") or u"").strip()
        if sp:
            out[u"summary_sheet_pattern"] = sp
        ms = unicode(
            obj.get(u"merge_source_sheet") or obj.get(u"merge_source") or u""
        ).strip()
        if ms:
            out[u"merge_source_sheet"] = ms
        if u"apply_to_summary" in obj:
            out[u"apply_to_summary"] = _as_bool(
                obj.get(u"apply_to_summary"), DEFAULT_APPLY_WIDTHS_TO_SUMMARY
            )
        raw_list = obj.get(u"column_widths")
        if raw_list is None and isinstance(obj.get(u"column_widths_mm"), dict):
            raw_list = []
            for k, v in obj.get(u"column_widths_mm").items():
                raw_list.append({u"name": k, u"width_mm": v})
        if isinstance(raw_list, list):
            i = 0
            while i < len(raw_list):
                item = raw_list[i]
                i = i + 1
                if not isinstance(item, dict):
                    continue
                name = unicode(item.get(u"name") or u"").strip()
                if not name:
                    continue
                try:
                    w = float(item.get(u"width_mm"))
                except (TypeError, ValueError):
                    w = 0.0
                out[u"column_widths"].append({u"name": name, u"width_mm": w})
    # дополнить именами с листа-шаблона, если переданы
    if header_names:
        have = {}
        for it in out[u"column_widths"]:
            have[unicode(it.get(u"name") or u"")] = True
        hi = 0
        names = list(header_names or [])
        while hi < len(names):
            nm = unicode(names[hi] or u"").strip()
            hi = hi + 1
            if not nm or nm in have:
                continue
            out[u"column_widths"].append({u"name": nm, u"width_mm": 0.0})
            have[nm] = True
    return out


def layout_width_map(layout):
    """name -> width_mm."""
    out = {}
    lay = normalize_layout_settings(layout)
    i = 0
    items = lay.get(u"column_widths") or []
    while i < len(items):
        it = items[i]
        i = i + 1
        name = unicode((it or {}).get(u"name") or u"").strip()
        if not name:
            continue
        try:
            out[name] = float((it or {}).get(u"width_mm"))
        except (TypeError, ValueError):
            out[name] = 0.0
    return out


def normalize_general_settings(obj):
    """Общие параметры серии (пароль защиты, владелец)."""
    roles = _owner_role_codes()
    out = {
        u"protect_password": unicode(DEFAULT_PROTECT_PASSWORD),
        u"protect_password_prev": u"",
        u"owner_name": u"",
        u"owner_role": unicode(DEFAULT_OWNER_ROLE),
        u"console_log": bool(DEFAULT_CONSOLE_LOG),
        u"dialog_font_pt": int(DEFAULT_DIALOG_FONT_PT),
    }
    if roles:
        out[u"owner_role"] = roles[0]
    if not isinstance(obj, dict):
        return out
    if u"protect_password_enc" in obj:
        out[u"protect_password"] = decode_protect_password(
            obj.get(u"protect_password_enc")
        )
    elif u"protect_password" in obj:
        out[u"protect_password"] = unicode(obj.get(u"protect_password") or u"")
    if u"protect_password_prev_enc" in obj:
        out[u"protect_password_prev"] = decode_protect_password(
            obj.get(u"protect_password_prev_enc")
        )
    elif u"protect_password_prev" in obj:
        out[u"protect_password_prev"] = unicode(obj.get(u"protect_password_prev") or u"")
    if u"console_log" in obj:
        out[u"console_log"] = _as_bool(obj.get(u"console_log"), DEFAULT_CONSOLE_LOG)
    if u"dialog_font_pt" in obj or u"dialog_font" in obj or u"ui_font_pt" in obj:
        raw_pt = obj.get(u"dialog_font_pt")
        if raw_pt is None:
            raw_pt = obj.get(u"dialog_font")
        if raw_pt is None:
            raw_pt = obj.get(u"ui_font_pt")
        out[u"dialog_font_pt"] = normalize_dialog_font_pt(
            raw_pt, DEFAULT_DIALOG_FONT_PT
        )
    name = unicode(obj.get("owner_name") or obj.get("fio_owner") or u"").strip()
    out[u"owner_name"] = name
    role = unicode(obj.get("owner_role") or u"").strip()
    role_cf = role.casefold() if hasattr(role, "casefold") else role.lower()
    if role in roles:
        out[u"owner_role"] = role
    elif role_cf in (u"руководитель", u"manager", u"boss"):
        out[u"owner_role"] = u"manager"
    elif role_cf in (u"сотрудник", u"employee", u"worker"):
        out[u"owner_role"] = u"employee"
    return out


def default_todo_task_settings():
    return {
        u"version": TODO_TASK_SETTINGS_JSON_VER,
        u"colorize": normalize_colorize_settings(None),
        u"sort": normalize_sort_settings(None),
        u"general": normalize_general_settings(None),
        u"layout": normalize_layout_settings(None),
        u"merge": normalize_merge_settings(None),
    }


def normalize_todo_task_settings(obj):
    defaults = default_todo_task_settings()
    if obj is None or not isinstance(obj, dict):
        return dict(defaults)
    out = dict(defaults)
    try:
        ver = int(obj.get("version", TODO_TASK_SETTINGS_JSON_VER))
    except (TypeError, ValueError):
        ver = TODO_TASK_SETTINGS_JSON_VER
    out[u"version"] = ver
    out[u"colorize"] = normalize_colorize_settings(obj.get("colorize"))
    out[u"sort"] = normalize_sort_settings(obj.get("sort"))
    gen = obj.get("general")
    if gen is None:
        gen = obj
    out[u"general"] = normalize_general_settings(gen)
    out[u"layout"] = normalize_layout_settings(obj.get("layout"))
    merge = obj.get("merge")
    out[u"merge"] = normalize_merge_settings(
        merge if isinstance(merge, dict) else None
    )
    return out


def load_todo_task_settings(doc=None, scope=None, book_path=None):
    """
    Настройки слота: private (книга) или global.
    scope=None — частные, если есть запись для пути, иначе глобальные.
    scope='private' — частные или копия глобальных (без записи на диск).
    scope='global' — только глобальная запись.
    """
    store = load_todo_task_store()
    glob = store.get(u"global") or default_todo_task_settings()
    wanted = unicode(scope or u"").strip()
    if wanted == SCOPE_GLOBAL:
        return normalize_todo_task_settings(glob)
    path = unicode(book_path or u"").strip() or get_doc_book_path(doc)
    key = _find_private_key(store.get(u"by_path"), path)
    if key:
        return normalize_todo_task_settings((store.get(u"by_path") or {}).get(key))
    return normalize_todo_task_settings(glob)


def _general_for_store(general):
    gen = normalize_general_settings(general)
    plain = unicode(gen.get(u"protect_password") if gen.get(u"protect_password") is not None else u"")
    prev = unicode(
        gen.get(u"protect_password_prev") if gen.get(u"protect_password_prev") is not None else u""
    )
    out = {
        u"protect_password_enc": encode_protect_password(plain),
        u"owner_name": unicode(gen.get(u"owner_name") or u""),
        u"owner_role": unicode(gen.get(u"owner_role") or DEFAULT_OWNER_ROLE),
        u"console_log": bool(gen.get(u"console_log")),
        u"dialog_font_pt": normalize_dialog_font_pt(
            gen.get(u"dialog_font_pt"), DEFAULT_DIALOG_FONT_PT
        ),
    }
    if prev != u"":
        out[u"protect_password_prev_enc"] = encode_protect_password(prev)
    return out


def _layout_for_store(layout):
    lay = normalize_layout_settings(layout)
    widths = []
    i = 0
    items = lay.get(u"column_widths") or []
    while i < len(items):
        it = items[i]
        i = i + 1
        name = unicode((it or {}).get(u"name") or u"").strip()
        if not name:
            continue
        try:
            w = float((it or {}).get(u"width_mm"))
        except (TypeError, ValueError):
            w = 0.0
        widths.append({u"name": name, u"width_mm": w})
    return {
        u"template_sheet": unicode(lay.get(u"template_sheet") or DEFAULT_LAYOUT_TEMPLATE_SHEET),
        u"task_sheet_pattern": unicode(
            lay.get(u"task_sheet_pattern") or DEFAULT_TASK_SHEET_PATTERN
        ),
        u"summary_sheet_pattern": unicode(
            lay.get(u"summary_sheet_pattern") or DEFAULT_SUMMARY_SHEET_PATTERN
        ),
        u"merge_source_sheet": unicode(
            lay.get(u"merge_source_sheet") or DEFAULT_MERGE_SOURCE_SHEET
        ),
        u"apply_to_summary": bool(lay.get(u"apply_to_summary")),
        u"column_widths": widths,
    }


def _slot_for_store(settings):
    payload = normalize_todo_task_settings(settings)
    payload[u"general"] = _general_for_store(payload.get(u"general"))
    payload[u"layout"] = _layout_for_store(payload.get(u"layout"))
    payload[u"merge"] = normalize_merge_settings(payload.get(u"merge"))
    return payload


def _notify_settings_saved(settings):
    try:
        from libre_macros_global_settings_lib import notify_config_changed

        notify_config_changed()
    except Exception:
        pass
    invalidate_todo_task_settings_cache()


def save_todo_task_store(store):
    payload = normalize_todo_task_store(store)
    disk = {
        u"version": TODO_TASK_SETTINGS_JSON_VER,
        u"global": _slot_for_store(payload.get(u"global")),
        u"by_path": {},
    }
    by_path = payload.get(u"by_path") or {}
    for key, val in by_path.items():
        nk = normalize_book_path(key) or unicode(key or u"").strip()
        if not nk:
            continue
        disk[u"by_path"][nk] = _slot_for_store(val)
    path = todo_task_settings_path()
    try:
        todo_task_settings_dir()
        with open(path, "wb") as f:
            f.write(
                json.dumps(disk, ensure_ascii=False, indent=2).encode("utf-8")
            )
    except Exception as err:
        return u"Не удалось сохранить настройки todo_task:\n%s" % err
    invalidate_todo_task_settings_cache()
    return None


def save_todo_task_settings(settings, doc=None, scope=None, book_path=None):
    """
    Записать слот настроек.
    scope='global' — в глобальную запись.
    scope='private' — в запись пути книги (если пути нет — в глобальные).
    scope=None — обновить частные, если для пути уже есть запись, иначе глобальные.
    """
    store = load_todo_task_store()
    slot = normalize_todo_task_settings(settings)
    wanted = unicode(scope or u"").strip()
    path = unicode(book_path or u"").strip() or get_doc_book_path(doc)
    if wanted == SCOPE_GLOBAL:
        write_global = True
    elif wanted == SCOPE_PRIVATE:
        write_global = not path
    else:
        write_global = (not path) or (
            _find_private_key(store.get(u"by_path"), path) is None
        )
    if write_global:
        store[u"global"] = slot
    else:
        by_path = store.get(u"by_path")
        if not isinstance(by_path, dict):
            by_path = {}
            store[u"by_path"] = by_path
        key = _find_private_key(by_path, path) or path
        by_path[key] = slot
    err = save_todo_task_store(store)
    if err:
        return err
    _notify_settings_saved(slot)
    return None


def get_protect_password(settings=None, doc=None):
    """Пароль защиты листа (plaintext). Пустая строка — защита без пароля."""
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    gen = (settings or {}).get(u"general") or {}
    if u"protect_password" in gen:
        return unicode(gen.get(u"protect_password") if gen.get(u"protect_password") is not None else u"")
    return unicode(DEFAULT_PROTECT_PASSWORD)


def get_owner_display_name(settings=None, doc=None):
    """ФИО владельца из настроек; если пусто — пользователь ОС."""
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    gen = (settings or {}).get(u"general") or {}
    name = unicode(gen.get(u"owner_name") or u"").strip()
    if name:
        return name
    try:
        import getpass

        name = unicode(getpass.getuser() or u"").strip()
        if name:
            return name
    except Exception:
        pass
    for key in ("USERNAME", "USER", "LOGNAME"):
        try:
            name = unicode(os.environ.get(key) or u"").strip()
        except Exception:
            name = u""
        if name:
            return name
    return u""


def get_owner_role(settings=None, doc=None):
    """Роль владельца: manager / employee (из настроек)."""
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    gen = (settings or {}).get(u"general") or {}
    role = unicode(gen.get(u"owner_role") or DEFAULT_OWNER_ROLE).strip()
    if role not in (u"manager", u"employee"):
        role = unicode(DEFAULT_OWNER_ROLE)
    return role


def get_merge_source_sheet(settings=None, doc=None):
    """Имя листа свода задач сотрудников в текущей книге (todo_task_merge)."""
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    lay = (settings or {}).get(u"layout") or {}
    name = unicode(lay.get(u"merge_source_sheet") or u"").strip()
    if name:
        return name
    return unicode(DEFAULT_MERGE_SOURCE_SHEET)


def get_merge_options(settings=None, doc=None):
    """Опции слияния: только существующие, контроль даты, накопление комментариев."""
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    return normalize_merge_settings((settings or {}).get(u"merge"))


def get_default_assignee(settings=None, doc=None):
    """
    Исполнитель по умолчанию для новой/копии задачи:
    владелец — только если роль «сотрудник»; для «руководитель» — пусто.
    """
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    role = get_owner_role(settings, doc=doc)
    try:
        import todo_task_cfg as _tcfg

        manager = unicode(
            getattr(_tcfg, "OWNER_ROLE_MANAGER", u"manager")
        )
    except Exception:
        manager = u"manager"
    if role == manager:
        return u""
    return unicode(get_owner_display_name(settings, doc=doc) or u"")


def get_unprotect_passwords(settings=None, extra=None, doc=None):
    """Пароли для снятия защиты: актуальный, предыдущий, default task, пустой."""
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    gen = (settings or {}).get(u"general") or {}
    out = []

    def _add(pwd):
        text = unicode(pwd if pwd is not None else u"")
        if text not in out:
            out.append(text)

    _add(get_protect_password(settings, doc=doc))
    prev = unicode(gen.get(u"protect_password_prev") or u"")
    if prev != u"":
        _add(prev)
    extra_list = extra if extra is not None else []
    ei = 0
    while ei < len(extra_list):
        _add(extra_list[ei])
        ei = ei + 1
    fallback = unicode(DEFAULT_PROTECT_PASSWORD)
    if fallback != u"":
        _add(fallback)
    _add(u"")
    return out
