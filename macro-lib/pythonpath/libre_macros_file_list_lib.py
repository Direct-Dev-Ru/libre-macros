# -*- coding: utf-8 -*-
"""Вспомогательная логика file_list_macro: Owner (Windows) и поиск по именам/содержимому."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.783"
import csv
import hashlib
import json
import os
import re
import sys

import libre_macros_file_list_cfg as _cfg

try:
    unicode
except NameError:
    unicode = str  # noqa: A001


def _u(value):
    if value is None:
        return u""
    if isinstance(value, bytes):
        for enc in ("utf-8", sys.getfilesystemencoding() or "utf-8", "cp1251"):
            if not enc:
                continue
            try:
                return value.decode(enc)
            except Exception:
                pass
        return value.decode("utf-8", errors="replace")
    return unicode(value)


def is_windows_platform():
    return sys.platform == "win32" or os.name == "nt"


def _owner_from_win32security(path):
    import win32security  # type: ignore

    sd = win32security.GetFileSecurity(
        path, win32security.OWNER_SECURITY_INFORMATION
    )
    sid = sd.GetSecurityDescriptorOwner()
    name, domain, _sid_type = win32security.LookupAccountSid(None, sid)
    name = _u(name).strip()
    domain = _u(domain).strip()
    if domain and name:
        return u"%s\\%s" % (domain, name)
    return name or domain


def _owner_from_ctypes(path):
    import ctypes
    from ctypes import wintypes

    OWNER_SECURITY_INFORMATION = 0x00000001
    SE_FILE_OBJECT = 1

    advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    GetNamedSecurityInfoW = advapi32.GetNamedSecurityInfoW
    GetNamedSecurityInfoW.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        ctypes.POINTER(wintypes.LPVOID),
        ctypes.POINTER(wintypes.LPVOID),
        ctypes.POINTER(wintypes.LPVOID),
        ctypes.POINTER(wintypes.LPVOID),
        ctypes.POINTER(wintypes.LPVOID),
    ]
    GetNamedSecurityInfoW.restype = wintypes.DWORD

    LookupAccountSidW = advapi32.LookupAccountSidW
    LookupAccountSidW.argtypes = [
        wintypes.LPCWSTR,
        wintypes.LPVOID,
        wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
        wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
        ctypes.POINTER(wintypes.DWORD),
    ]
    LookupAccountSidW.restype = wintypes.BOOL

    LocalFree = kernel32.LocalFree
    LocalFree.argtypes = [wintypes.LPVOID]
    LocalFree.restype = wintypes.LPVOID

    owner_sid = wintypes.LPVOID()
    sd_ptr = wintypes.LPVOID()
    rc = GetNamedSecurityInfoW(
        _u(path),
        SE_FILE_OBJECT,
        OWNER_SECURITY_INFORMATION,
        ctypes.byref(owner_sid),
        None,
        None,
        None,
        ctypes.byref(sd_ptr),
    )
    if rc != 0 or not owner_sid:
        if sd_ptr:
            LocalFree(sd_ptr)
        return u""

    name_len = wintypes.DWORD(0)
    domain_len = wintypes.DWORD(0)
    sid_use = wintypes.DWORD(0)
    LookupAccountSidW(
        None,
        owner_sid,
        None,
        ctypes.byref(name_len),
        None,
        ctypes.byref(domain_len),
        ctypes.byref(sid_use),
    )
    if name_len.value == 0:
        if sd_ptr:
            LocalFree(sd_ptr)
        return u""

    name_buf = ctypes.create_unicode_buffer(name_len.value)
    domain_buf = ctypes.create_unicode_buffer(domain_len.value)
    ok = LookupAccountSidW(
        None,
        owner_sid,
        name_buf,
        ctypes.byref(name_len),
        domain_buf,
        ctypes.byref(domain_len),
        ctypes.byref(sid_use),
    )
    if sd_ptr:
        LocalFree(sd_ptr)
    if not ok:
        return u""
    name = _u(name_buf.value).strip()
    domain = _u(domain_buf.value).strip()
    if domain and name:
        return u"%s\\%s" % (domain, name)
    return name or domain


def get_file_owner(path):
    """
    Владелец файла/папки на Windows (NTFS Owner).
    На Linux/macOS — пустая строка (колонка Owner не добавляется).
    """
    if not is_windows_platform():
        return u""
    full = _u(path)
    if not full:
        return u""
    try:
        return _owner_from_win32security(full)
    except Exception:
        pass
    try:
        return _owner_from_ctypes(full)
    except Exception as err:
        print("Owner: %s — %s" % (full, err))
        return u""


def default_dlp_dict_uses_sheet():
    return bool(getattr(_cfg, "DEFAULT_DLP_DICT_USE_SHEET", False))


def coerce_saved_dlp_dict_ref(value, settings_dir, kind=u"name"):
    """
    Для standalone (DEFAULT_DLP_DICT_USE_SHEET): пустое или дефолтный JSON
    → имя листа; иначе как сохранено (файл или лист).
    kind: 'name' | 'content'.
    """
    text = _u(value).strip()
    if not default_dlp_dict_uses_sheet():
        return text
    if is_dlp_dict_sheet_ref(text):
        return text
    if text == "":
        return default_dlp_pattern_sheet_name()
    base = os.path.basename(text).casefold()
    if kind == u"content":
        default_name = _u(
            getattr(_cfg, "DLP_CONTENT_DICT_FILENAME", "search_content_patterns.json")
        )
    else:
        default_name = _u(
            getattr(_cfg, "DLP_NAME_DICT_FILENAME", "search_name_patterns.json")
        )
    legacy = (
        default_name.casefold(),
        u"dlp_name_patterns.json",
        u"dlp_content_patterns.json",
        u"search_name_patterns.json",
        u"search_content_patterns.json",
    )
    if base in legacy:
        return default_dlp_pattern_sheet_name()
    # полный путь к дефолтному файлу в settings_dir
    try:
        expected = os.path.normpath(os.path.join(_u(settings_dir), default_name))
        if os.path.normpath(text) == expected:
            return default_dlp_pattern_sheet_name()
    except Exception:
        pass
    return text


def default_dlp_name_dict_path(settings_dir):
    if default_dlp_dict_uses_sheet():
        return default_dlp_pattern_sheet_name()
    name = getattr(_cfg, "DLP_NAME_DICT_FILENAME", "dlp_name_patterns.json")
    return os.path.join(_u(settings_dir), name)


def default_dlp_content_dict_path(settings_dir):
    if default_dlp_dict_uses_sheet():
        return default_dlp_pattern_sheet_name()
    name = getattr(_cfg, "DLP_CONTENT_DICT_FILENAME", "dlp_content_patterns.json")
    return os.path.join(_u(settings_dir), name)


def is_dlp_dict_sheet_ref(value):
    """
    Одно слово в поле словаря → имя листа текущей книги (не путь к файлу).
    Путь: слэши, диск «:», пробелы, расширения .json/.csv/…
    """
    text = _u(value).strip()
    if not text:
        return False
    if any(ch in text for ch in (u"/", u"\\", u":")):
        return False
    if u" " in text or u"\t" in text:
        return False
    if any(ch in text for ch in (u"[", u"]", u"*", u"?", u"'")):
        return False
    low = text.casefold()
    for ext in (u".json", u".jsn", u".csv", u".txt", u".tsv"):
        if low.endswith(ext):
            return False
    # Calc: имя листа ≤ 31 символ
    if len(text) > 31:
        return False
    return True


def default_dlp_pattern_sheet_name():
    name = _u(
        getattr(_cfg, "DEFAULT_DLP_PATTERN_SHEET_NAME", u"поиск_шаблоны")
    ).strip()
    return name or u"поиск_шаблоны"


def resolve_dlp_pattern_sheet_name(value, result_sheet_name=u""):
    """
    Имя листа для генерации/словаря.
    Некорректное/пустое → DEFAULT_DLP_PATTERN_SHEET_NAME.
    Если совпадает с листом результата — суффикс «_шаблоны».
    Возвращает (name, used_default_bool).
    """
    text = _u(value).strip()
    used_default = False
    if is_dlp_dict_sheet_ref(text):
        name = text
    else:
        name = default_dlp_pattern_sheet_name()
        used_default = True
    result = _u(result_sheet_name).strip()
    if result and name.casefold() == result.casefold():
        base = default_dlp_pattern_sheet_name()
        if base.casefold() == result.casefold():
            name = u"поиск_шаблоны"
            if name.casefold() == result.casefold():
                name = u"поиск_шаблоны_1"
        else:
            name = base
        used_default = True
    if len(name) > 31:
        name = name[:31]
    return name, used_default


def default_dlp_pattern_sheet_rows():
    """Строки листа-словаря по умолчанию: (имя, паттерн, описание)."""
    raw = getattr(_cfg, "DEFAULT_DLP_PATTERN_SHEET_ROWS", None) or ()
    out = []
    for item in raw:
        if not item or len(item) < 2:
            continue
        name = _u(item[0]).strip()
        pattern = _u(item[1]).strip()
        desc = _u(item[2]).strip() if len(item) > 2 else u""
        if name and pattern:
            out.append((name, pattern, desc))
    if out:
        return out
    # fallback: компактный DEFAULT_DLP_NAME_PATTERNS
    for key, pattern in (getattr(_cfg, "DEFAULT_DLP_NAME_PATTERNS", None) or {}).items():
        k = _u(key).strip()
        p = _u(pattern).strip()
        if k and p:
            out.append((k, p, u""))
    return out


def default_dlp_pattern_sheet_headers():
    headers = getattr(_cfg, "DLP_PATTERN_SHEET_HEADERS", None) or (
        u"имя",
        u"паттерн",
        u"описание",
    )
    return tuple(_u(h) for h in headers)


def dlp_pattern_sheet_font_name():
    """Шрифт листа паттернов (standalone-safe: через _cfg, не _fl_state)."""
    name = _u(getattr(_cfg, "DLP_PATTERN_SHEET_FONT_NAME", u"PT Sans")).strip()
    return name or u"PT Sans"


def default_dlp_pattern_sheet_help_lines():
    raw = getattr(_cfg, "DLP_PATTERN_SHEET_HELP_LINES", None) or ()
    return tuple(_u(line) for line in raw)


def dlp_pattern_sheet_help_placement():
    """(col_0based, start_row_0based) для справки на листе паттернов."""
    try:
        col = int(getattr(_cfg, "DLP_PATTERN_SHEET_HELP_COL", 4))
    except (TypeError, ValueError):
        col = 4
    try:
        row = int(getattr(_cfg, "DLP_PATTERN_SHEET_HELP_START_ROW", 1))
    except (TypeError, ValueError):
        row = 1
    if col < 0:
        col = 4
    if row < 0:
        row = 1
    return col, row


def patterns_dict_from_sheet_rows(rows):
    """Список (имя, паттерн[, описание]) или ячеек → dict id→regex (через coerce)."""
    out = {}
    first = True
    for row in rows or ():
        if not row or len(row) < 2:
            continue
        pid = _u(row[0]).strip()
        preg = _u(row[1]).strip()
        if not pid and not preg:
            continue
        if first:
            first = False
            if _csv_header_row(pid, preg):
                continue
        if pid and preg:
            out[pid] = coerce_dlp_pattern(preg)
    return out


def default_dlp_name_patterns():
    raw = getattr(_cfg, "DEFAULT_DLP_NAME_PATTERNS", None) or {}
    out = {}
    for key, pattern in raw.items():
        k = _u(key).strip()
        p = _u(pattern).strip()
        if k and p:
            out[k] = p
    return out


def ensure_default_dlp_name_dict(settings_dir):
    """Создать JSON-словарь имён по умолчанию, если файла ещё нет."""
    path = default_dlp_name_dict_path(settings_dir)
    if is_dlp_dict_sheet_ref(path):
        return path
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        if not os.path.isfile(path):
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(
                    default_dlp_name_patterns(),
                    handle,
                    ensure_ascii=False,
                    indent=2,
                )
    except Exception as err:
        print("Поиск: не удалось записать словарь по умолчанию: %s" % err)
    return path


def coerce_dlp_pattern(pattern):
    """
    Простое слово → поиск подфразы без учёта регистра: (?i)слово.
    Уже regex (есть метасимволы или флаги (?…)) — как есть.
    «(?i)» — inline ignore-case в Python re.
    """
    p = _u(pattern).strip()
    if not p:
        return p
    # Метасимволы regex / inline-флаги — оставляем как есть (re в теле def: AO-safe).
    if p.startswith(u"(?") or re.search(r"[.\\^$*+?{}\[\]|()\\\\]", p):
        return p
    return u"(?i)" + re.escape(p)


def normalize_dlp_patterns(raw):
    """dict id→regex или list[{id,pattern}|{name,pattern}] → dict."""
    out = {}
    if isinstance(raw, dict):
        # формат {"patterns": {...}} или сразу id→regex
        if "patterns" in raw and isinstance(raw.get("patterns"), dict):
            raw = raw.get("patterns")
        for key, pattern in raw.items():
            if key in ("version", "v", "comment", "description"):
                continue
            if isinstance(pattern, dict):
                pid = _u(pattern.get("id") or pattern.get("name") or key).strip()
                preg = _u(pattern.get("pattern") or pattern.get("regex") or "").strip()
            else:
                pid = _u(key).strip()
                preg = _u(pattern).strip()
            if pid and preg:
                out[pid] = coerce_dlp_pattern(preg)
        return out
    if isinstance(raw, (list, tuple)):
        for item in raw:
            if not isinstance(item, dict):
                continue
            pid = _u(item.get("id") or item.get("name") or "").strip()
            preg = _u(item.get("pattern") or item.get("regex") or "").strip()
            if pid and preg:
                out[pid] = coerce_dlp_pattern(preg)
    return out


def _csv_sniff_delimiter(sample):
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        return dialect.delimiter
    except Exception:
        pass
    # Excel RU часто «;»
    if sample.count(";") >= sample.count(","):
        return ";"
    return ","


def _csv_header_row(name_cell, pattern_cell):
    n = _u(name_cell).strip().casefold()
    p = _u(pattern_cell).strip().casefold()
    name_headers = (
        u"name",
        u"id",
        u"имя",
        u"название",
        u"pattern_name",
        u"ключ",
    )
    pattern_headers = (
        u"pattern",
        u"regex",
        u"паттерн",
        u"шаблон",
        u"выражение",
    )
    return n in name_headers or p in pattern_headers


def load_dlp_patterns_from_csv(path):
    """
    CSV: колонка 1 — имя, колонка 2 — паттерн/слово.
    Разделитель , или ; (автоопределение). Простое слово → (?i)слово.
    """
    path = _u(path)
    out = {}
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        sample = handle.read(4096)
        handle.seek(0)
        delim = _csv_sniff_delimiter(sample)
        reader = csv.reader(handle, delimiter=delim)
        first = True
        for row in reader:
            if not row or len(row) < 2:
                continue
            pid = _u(row[0]).strip()
            preg = _u(row[1]).strip()
            if not pid and not preg:
                continue
            if first:
                first = False
                if _csv_header_row(pid, preg):
                    continue
            if pid and preg:
                out[pid] = coerce_dlp_pattern(preg)
    return out


def load_dlp_patterns_from_json(path):
    with open(path, "r", encoding="utf-8-sig") as handle:
        data = json.load(handle)
    return normalize_dlp_patterns(data)


def load_dlp_patterns(dict_path, fallback_defaults=True):
    """
    Загрузить словарь из JSON или CSV (имя, паттерн).
    При ошибке/пустом пути — DEFAULT_DLP_NAME_PATTERNS (если fallback_defaults).
    """
    path = _u(dict_path).strip()
    if path and os.path.isfile(path):
        ext = os.path.splitext(path)[1].lower()
        patterns = {}
        err_msg = None
        try:
            if ext == ".csv":
                patterns = load_dlp_patterns_from_csv(path)
            elif ext in (".json", ".jsn"):
                patterns = load_dlp_patterns_from_json(path)
            else:
                # расширение неизвестно — сначала JSON, затем CSV
                try:
                    patterns = load_dlp_patterns_from_json(path)
                except Exception:
                    patterns = load_dlp_patterns_from_csv(path)
        except Exception as err:
            err_msg = err
            patterns = {}
        if patterns:
            return patterns, path
        if err_msg is not None:
            print("Поиск: не удалось прочитать «%s»: %s" % (path, err_msg))
        else:
            print("Поиск: файл «%s» не содержит паттернов" % path)
    if fallback_defaults:
        # defaults тоже прогоняем через coerce (на случай простых слов в cfg)
        return normalize_dlp_patterns(default_dlp_name_patterns()), path
    return {}, path


def compile_dlp_patterns(patterns):
    """Список (id, compiled_re); битые regex пропускаются с логом."""
    compiled = []
    if not isinstance(patterns, dict):
        return compiled
    for pid, preg in patterns.items():
        try:
            compiled.append((_u(pid), re.compile(_u(preg))))
        except re.error as err:
            print("Поиск: неверный regex «%s»: %s" % (pid, err))
    return compiled


def dlp_haystack_for_entry(entry):
    """Текст для проверки имён: имя + относительный + полный путь."""
    if not isinstance(entry, dict):
        return u""
    parts = [
        _u(entry.get("name")),
        _u(entry.get("rel_path")),
        _u(entry.get("parent")),
        _u(entry.get("full_path")),
    ]
    return u"\n".join([p for p in parts if p])


def match_text_dlp_patterns(text, compiled_patterns):
    """Проверить произвольный текст по всему словарю → список id."""
    hay = _u(text)
    if not hay or not compiled_patterns:
        return []
    hits = []
    for pid, cre in compiled_patterns:
        try:
            if cre.search(hay):
                hits.append(pid)
        except Exception:
            pass
    return hits


def match_entry_dlp_patterns(entry, compiled_patterns):
    """
    Проверить запись по всему словарю.
    Возвращает список id паттернов с совпадением (порядок словаря).
    """
    return match_text_dlp_patterns(dlp_haystack_for_entry(entry), compiled_patterns)


def _norm_ext(path_or_ext):
    text = _u(path_or_ext).strip().lower()
    if not text:
        return u""
    # путь / имя.файла → расширение; голый «.docx» / «docx» — как есть
    if u"/" in text or u"\\" in text or (u"." in text[1:] and not text.startswith(u".")):
        text = os.path.splitext(text)[1]
    if text and not text.startswith(u"."):
        text = u"." + text
    return text


def dlp_ext_in_set(path, ext_set):
    return _norm_ext(path) in (ext_set or frozenset())


def is_dlp_text_file(path):
    return dlp_ext_in_set(path, getattr(_cfg, "DLP_TEXT_EXTS", frozenset()))


def is_dlp_word_file(path):
    return dlp_ext_in_set(path, getattr(_cfg, "DLP_WORD_EXTS", frozenset()))


def is_dlp_excel_file(path):
    return dlp_ext_in_set(path, getattr(_cfg, "DLP_EXCEL_EXTS", frozenset()))


def _bom_encoding(raw):
    if raw.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    if raw.startswith(b"\xff\xfe"):
        return "utf-16-le"
    if raw.startswith(b"\xfe\xff"):
        return "utf-16-be"
    if raw.startswith(b"\xff\xfe\x00\x00"):
        return "utf-32-le"
    if raw.startswith(b"\x00\x00\xfe\xff"):
        return "utf-32-be"
    return None


def _decode_score(text):
    """Грубая оценка «похоже на текст» (больше — лучше)."""
    if not text:
        return -1
    sample = text[:4000]
    if not sample:
        return -1
    bad = 0
    good = 0
    for ch in sample:
        o = ord(ch)
        if ch in ("\x00", "\ufffd"):
            bad += 3
        elif o < 32 and ch not in ("\n", "\r", "\t"):
            bad += 1
        else:
            good += 1
    return good - bad * 4


def detect_and_decode_bytes(raw):
    """
    Каскад кодировок: BOM → utf-8 → cp1251 → cp866 → utf-16 → latin-1.
    Возвращает (text, encoding_label).
    """
    if not raw:
        return u"", u""
    bom_enc = _bom_encoding(raw)
    candidates = []
    if bom_enc:
        candidates.append(bom_enc)
    for enc in ("utf-8", "cp1251", "cp866", "utf-16-le", "utf-16-be", "latin-1"):
        if enc not in candidates:
            candidates.append(enc)
    best_text = u""
    best_enc = u""
    best_score = -10**9
    for enc in candidates:
        try:
            if enc == "utf-8":
                text = raw.decode("utf-8")
            else:
                text = raw.decode(enc)
        except Exception:
            continue
        score = _decode_score(text)
        # BOM / utf-8 без ошибок — бонус
        if enc == bom_enc:
            score += 50
        if enc == "utf-8" and bom_enc is None:
            score += 10
        if score > best_score:
            best_score = score
            best_text = text
            best_enc = enc
    return best_text, best_enc


def read_text_file_for_dlp(path, max_bytes=None, max_chars=None):
    """
    Прочитать текстовый файл с каскадом кодировок.
    Возвращает (text, encoding_label); при ошибке («», «»).
    """
    full = _u(path)
    if not full or not os.path.isfile(full):
        return u"", u""
    if max_bytes is None:
        max_bytes = int(getattr(_cfg, "DLP_CONTENT_MAX_BYTES", 2 * 1024 * 1024))
    if max_chars is None:
        max_chars = int(getattr(_cfg, "DLP_CONTENT_MAX_CHARS", 500000))
    try:
        with open(full, "rb") as handle:
            raw = handle.read(int(max_bytes) + 1)
    except Exception as err:
        print("Поиск (чтение): «%s»: %s" % (full, err))
        return u"", u""
    if not raw:
        return u"", u""
    truncated = len(raw) > int(max_bytes)
    if truncated:
        raw = raw[: int(max_bytes)]
    text, enc = detect_and_decode_bytes(raw)
    if max_chars and len(text) > int(max_chars):
        text = text[: int(max_chars)]
    label = enc or u"unknown"
    if truncated:
        label = label + u"+trunc"
    return text, label


def ensure_default_dlp_content_dict(settings_dir):
    """Создать JSON-словарь содержимого по умолчанию, если файла ещё нет."""
    path = default_dlp_content_dict_path(settings_dir)
    if is_dlp_dict_sheet_ref(path):
        return path
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        if not os.path.isfile(path):
            raw = getattr(_cfg, "DEFAULT_DLP_CONTENT_PATTERNS", None)
            if not raw:
                raw = default_dlp_name_patterns()
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(
                    normalize_dlp_patterns(raw),
                    handle,
                    ensure_ascii=False,
                    indent=2,
                )
    except Exception as err:
        print("Поиск: не удалось записать словарь содержимого: %s" % err)
    return path


def dlp_content_scope_for_path(path, want_text, want_word, want_excel):
    """
    Какой экстрактор использовать: 'text' | 'word' | 'excel' | None.
    Word/Excel предпочтительнее text при пересечении расширений (например .rtf).
    """
    if want_word and is_dlp_word_file(path):
        return u"word"
    if want_excel and is_dlp_excel_file(path):
        return u"excel"
    if want_text and is_dlp_text_file(path):
        return u"text"
    return None


def dlp_log_sheet_name(result_sheet_name):
    base = _u(result_sheet_name).strip() or u"Результат"
    suffix = _u(getattr(_cfg, "DLP_LOG_SHEET_SUFFIX", u"_DLP"))
    name = base + suffix
    # Calc: имя листа ≤ 31 символ
    if len(name) > 31:
        keep = 31 - len(suffix)
        if keep < 1:
            keep = 1
        name = base[:keep] + suffix
        name = name[:31]
    return name


def highlight_sheet_rows(sheet, row_indices, ncols, color=None):
    """Подкрасить строки (1-based индексы листа) цветом DLP_MATCH_ROW_COLOR."""
    if sheet is None or not row_indices or ncols < 1:
        return 0
    fill = color
    if fill is None:
        fill = int(getattr(_cfg, "DLP_MATCH_ROW_COLOR", 0xFFCCCC))
    ec = int(ncols) - 1
    done = 0
    for row in row_indices:
        try:
            r = int(row)
        except (TypeError, ValueError):
            continue
        if r < 0:
            continue
        try:
            rng = sheet.getCellRangeByPosition(0, r, ec, r)
            rng.CellBackColor = int(fill)
            done += 1
        except Exception:
            try:
                c = 0
                while c < ncols:
                    sheet.getCellByPosition(c, r).CellBackColor = int(fill)
                    c += 1
                done += 1
            except Exception:
                pass
    return done


def sha256_max_bytes_from_gb(max_gb):
    """Гб (float/str) → лимит в байтах; по умолчанию 1 Гб."""
    try:
        gb = float(max_gb)
    except (TypeError, ValueError):
        gb = float(getattr(_cfg, "DEFAULT_SHA256_MAX_GB", 1.0) or 1.0)
    if gb <= 0:
        gb = float(getattr(_cfg, "DEFAULT_SHA256_MAX_GB", 1.0) or 1.0)
    return int(gb * (1024 ** 3))


def file_sha256_hex(path, max_bytes=None, chunk_size=None):
    """
    SHA-256 файла побайтно (чанками).
    Возвращает hex-строку или u'' (папка / слишком большой / ошибка).
    """
    full = _u(path)
    if not full or not os.path.isfile(full):
        return u""
    if max_bytes is None:
        max_bytes = sha256_max_bytes_from_gb(
            getattr(_cfg, "DEFAULT_SHA256_MAX_GB", 1.0)
        )
    if chunk_size is None:
        chunk_size = int(getattr(_cfg, "SHA256_READ_CHUNK", 1024 * 1024))
    try:
        size = int(os.path.getsize(full))
    except Exception:
        return u""
    if size > int(max_bytes):
        return u""
    digest = hashlib.sha256()
    try:
        with open(full, "rb") as handle:
            while True:
                block = handle.read(int(chunk_size))
                if not block:
                    break
                digest.update(block)
    except Exception as err:
        print("SHA256 «%s»: %s" % (full, err))
        return u""
    return _u(digest.hexdigest())


def collect_entry_sha256_digests(entries, max_bytes=None, progress_cb=None):
    """
    Для каждой записи: hex SHA256 или ''.
    progress_cb(done, total, name) — опционально.
    Возвращает (digests list, hash_to_entry_indices dict).
    """
    if max_bytes is None:
        max_bytes = sha256_max_bytes_from_gb(
            getattr(_cfg, "DEFAULT_SHA256_MAX_GB", 1.0)
        )
    digests = []
    hash_map = {}
    total = len(entries or [])
    for idx, entry in enumerate(entries or []):
        dig = u""
        if entry and entry.get("kind") == "файл":
            dig = file_sha256_hex(entry.get("full_path") or "", max_bytes=max_bytes)
        digests.append(dig)
        if dig:
            bucket = hash_map.get(dig)
            if bucket is None:
                hash_map[dig] = [idx]
            else:
                bucket.append(idx)
        if progress_cb is not None:
            try:
                progress_cb(
                    idx + 1,
                    total,
                    _u((entry or {}).get("name") or ""),
                )
            except Exception:
                pass
    return digests, hash_map


def dup_marks_from_sha256_map(digests, hash_map):
    """
    По digests и карте hash→indices:
    для каждой строки (is_dup 'Да'|'', dup_n int|None).
    Первый в группе дублей — 1, второй — 2, …
    """
    out = []
    for idx, dig in enumerate(digests or []):
        if not dig:
            out.append((u"", None))
            continue
        group = hash_map.get(dig) or []
        if len(group) < 2:
            out.append((u"", None))
            continue
        try:
            order = group.index(idx) + 1
        except ValueError:
            order = 1
        out.append((u"Да", order))
    return out


def parse_path_masks(mask_text):
    """
    Маски путей: через запятую или `;`. Пусто → [].
    Единственный плейсхолдер — `*` (любой фрагмент, в т.ч. пустой).
    """
    text = _u(mask_text).strip()
    if text == u"":
        return []
    parts = []
    for chunk in text.replace(u",", u";").split(u";"):
        item = chunk.strip()
        if item != u"" and item not in parts:
            parts.append(item)
    return parts


def _normalize_path_for_mask(path_text):
    return _u(path_text).replace(u"\\", u"/")


def path_matches_star_pattern(path_text, pattern):
    """
    Совпадение пути с маской: только `*` — плейсхолдер (как .* в regex).
    Остальные символы — литералы. Сравнение без учёта регистра на Windows.
    """
    text = _normalize_path_for_mask(path_text)
    pat = _normalize_path_for_mask(pattern).strip()
    if not pat:
        return False
    # Собрать regex: * → .*, прочее экранировать.
    chunks = []
    i = 0
    n = len(pat)
    while i < n:
        ch = pat[i]
        if ch == u"*":
            chunks.append(u".*")
            i += 1
            continue
        chunks.append(re.escape(ch))
        i += 1
    flags = re.IGNORECASE if sys.platform == "win32" else 0
    try:
        rx = re.compile(u"^" + u"".join(chunks) + u"$", flags)
    except re.error:
        return False
    return rx.match(text) is not None


def path_matches_any_star_mask(path_text, masks):
    if not masks:
        return False
    text = _normalize_path_for_mask(path_text)
    # Доп. вариант с ведущим «/»: чтобы «*/.*» ловил и «.git» на корне
    # (как «/.git» — комбинация «/.» в пути).
    variants = [text]
    if text and not text.startswith(u"/"):
        variants.append(u"/" + text)
    for pattern in masks:
        for cand in variants:
            if path_matches_star_pattern(cand, pattern):
                return True
    return False


def entry_matches_include_masks(name, rel_path, full_path, masks):
    """
    Маски включения: пусто → все.
    Проверка по имени файла, относительному и полному пути (* по всему пути).
    """
    if not masks:
        return True
    candidates = (
        _u(name),
        _normalize_path_for_mask(rel_path),
        _normalize_path_for_mask(full_path),
    )
    for cand in candidates:
        if cand and path_matches_any_star_mask(cand, masks):
            return True
    return False


def entry_matches_exclude_masks(name, rel_path, full_path, masks):
    """Маски исключения: пусто → ничего не исключаем. Иначе OR по маскам."""
    if not masks:
        return False
    candidates = (
        _u(name),
        _normalize_path_for_mask(rel_path),
        _normalize_path_for_mask(full_path),
    )
    for cand in candidates:
        if cand and path_matches_any_star_mask(cand, masks):
            return True
    return False
