# -*- coding: utf-8 -*-
"""Вспомогательная логика file_list_macro: Owner (Windows) и мини-DLP по именам."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.730"
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


def default_dlp_name_dict_path(settings_dir):
    name = getattr(_cfg, "DLP_NAME_DICT_FILENAME", "dlp_name_patterns.json")
    return os.path.join(_u(settings_dir), name)


def default_dlp_content_dict_path(settings_dir):
    name = getattr(_cfg, "DLP_CONTENT_DICT_FILENAME", "dlp_content_patterns.json")
    return os.path.join(_u(settings_dir), name)


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
        print("DLP: не удалось записать словарь по умолчанию: %s" % err)
    return path


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
                out[pid] = preg
        return out
    if isinstance(raw, (list, tuple)):
        for item in raw:
            if not isinstance(item, dict):
                continue
            pid = _u(item.get("id") or item.get("name") or "").strip()
            preg = _u(item.get("pattern") or item.get("regex") or "").strip()
            if pid and preg:
                out[pid] = preg
    return out


def load_dlp_patterns(dict_path, fallback_defaults=True):
    """
    Загрузить словарь regex из JSON.
    При ошибке/пустом пути — DEFAULT_DLP_NAME_PATTERNS (если fallback_defaults).
    """
    path = _u(dict_path).strip()
    if path and os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
            patterns = normalize_dlp_patterns(data)
            if patterns:
                return patterns, path
            print("DLP: файл «%s» не содержит паттернов" % path)
        except Exception as err:
            print("DLP: не удалось прочитать «%s»: %s" % (path, err))
    if fallback_defaults:
        return default_dlp_name_patterns(), path
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
            print("DLP: неверный regex «%s»: %s" % (pid, err))
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


def match_entry_dlp_patterns(entry, compiled_patterns):
    """
    Проверить запись по всему словарю.
    Возвращает список id паттернов с совпадением (порядок словаря).
    """
    text = dlp_haystack_for_entry(entry)
    if not text or not compiled_patterns:
        return []
    hits = []
    for pid, cre in compiled_patterns:
        try:
            if cre.search(text):
                hits.append(pid)
        except Exception:
            pass
    return hits


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
    """Подкрасить строки (0-based индексы листа) цветом DLP_MATCH_ROW_COLOR."""
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
