# -*- coding: utf-8 -*-
"""
Seed runtime-карты переменных: окружение (__ENV_*) и сведения об ОС (__OS_*).

Пишется в сегмент «Глобальные_переменные» при старте сбора
(_merge_seed_globals / merge_run_context_begin).
"""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.728"
try:
    unicode
except NameError:
    unicode = str

import os
import platform
import socket
import sys
import tempfile

try:
    import getpass
except ImportError:
    getpass = None

# Префиксы имён в runtime-карте / <<Переменные.…>>
ENV_VAR_PREFIX = u"__ENV_"
OS_VAR_PREFIX = u"__OS_"

# cell_ref для журнала / отладки
CELL_REF_ENV = u"env"
CELL_REF_OS = u"os"


def _u(v):
    if v is None:
        return u""
    try:
        return unicode(v)
    except Exception:
        return u"%s" % v


def sanitize_runtime_var_name(raw, prefix=u""):
    """
    Имя без '~' (разделитель ключа карты). Остальное — как есть.
    Пустой результат после очистки → ''.
    """
    name = _u(raw).strip()
    if name == u"":
        return u""
    if u"~" in name:
        name = name.replace(u"~", u"_")
    pref = _u(prefix)
    if pref and not name.startswith(pref):
        name = pref + name
    if u"~" in name:
        name = name.replace(u"~", u"_")
    return name


def iter_env_variable_entries():
    """
    Список (name, text) для всех переменных среды.
    name = __ENV_<имя_из_os.environ> (с заменой '~').
    """
    out = []
    try:
        items = list(os.environ.items())
    except Exception:
        return out
    i = 0
    while i < len(items):
        key, val = items[i]
        i += 1
        nm = sanitize_runtime_var_name(key, ENV_VAR_PREFIX)
        if nm == u"":
            continue
        out.append((nm, _u(val)))
    return out


def _os_username():
    try:
        from libre_macros_source_roots_lib import os_username

        return _u(os_username()).strip()
    except Exception:
        pass
    if getpass is not None:
        try:
            return _u(getpass.getuser() or u"").strip()
        except Exception:
            pass
    for key in (u"USER", u"USERNAME", u"LOGNAME"):
        try:
            v = _u(os.environ.get(key) or u"").strip()
        except Exception:
            v = u""
        if v:
            return v
    return u""


def _memory_bytes():
    """(total, available) в байтах или (None, None)."""
    total = None
    avail = None
    try:
        import psutil

        vm = psutil.virtual_memory()
        total = int(vm.total)
        avail = int(vm.available)
        return total, avail
    except Exception:
        pass
    if sys.platform.startswith(u"linux"):
        try:
            f = open(u"/proc/meminfo", u"r")
            try:
                for line in f:
                    if line.startswith(u"MemTotal:"):
                        total = int(line.split()[1]) * 1024
                    elif line.startswith(u"MemAvailable:"):
                        avail = int(line.split()[1]) * 1024
            finally:
                f.close()
        except Exception:
            pass
    if total is None and sys.platform == u"win32":
        try:
            import ctypes

            class _MEM(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong),
                    ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong),
                    ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]

            st = _MEM()
            st.dwLength = ctypes.sizeof(_MEM)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(st)):
                total = int(st.ullTotalPhys)
                avail = int(st.ullAvailPhys)
        except Exception:
            pass
    return total, avail


def _safe_platform(attr, default=u""):
    try:
        fn = getattr(platform, attr, None)
        if callable(fn):
            return _u(fn() or default)
    except Exception:
        pass
    return _u(default)


def iter_os_variable_entries():
    """
    Список (name, text) со сведениями об ОС / runtime.
    name = __OS_<КЛЮЧ>.
    """
    pairs = []

    def add(suffix, value):
        nm = sanitize_runtime_var_name(suffix, OS_VAR_PREFIX)
        if nm == u"":
            return
        pairs.append((nm, _u(value)))

    user = _os_username()
    add(u"USER", user)
    add(u"USERNAME", user)
    try:
        home = os.path.expanduser(u"~")
    except Exception:
        home = u""
    add(u"HOME", home)
    try:
        cwd = os.getcwd()
    except Exception:
        cwd = u""
    add(u"CWD", cwd)
    try:
        tmp = tempfile.gettempdir()
    except Exception:
        tmp = u""
    add(u"TEMP", tmp)
    add(u"TMPDIR", tmp)

    add(u"PLATFORM", sys.platform)
    add(u"SYSTEM", _safe_platform(u"system"))
    add(u"RELEASE", _safe_platform(u"release"))
    add(u"VERSION", _safe_platform(u"version"))
    add(u"MACHINE", _safe_platform(u"machine"))
    add(u"PROCESSOR", _safe_platform(u"processor"))
    arch = u""
    try:
        bits, linkage = platform.architecture()
        arch = (u"%s %s" % (_u(bits), _u(linkage))).strip()
    except Exception:
        arch = _safe_platform(u"machine")
    add(u"ARCHITECTURE", arch)
    add(u"NODE", _safe_platform(u"node"))
    try:
        host = socket.gethostname()
    except Exception:
        host = u""
    add(u"HOSTNAME", host)

    add(u"PYTHON_VERSION", _u(sys.version.split()[0] if sys.version else u""))
    add(u"PYTHON_IMPLEMENTATION", _safe_platform(u"python_implementation"))
    try:
        add(u"PID", os.getpid())
    except Exception:
        pass
    try:
        add(u"CPU_COUNT", os.cpu_count() or u"")
    except Exception:
        pass
    try:
        add(u"SEP", os.sep)
        add(u"PATHSEP", os.pathsep)
        add(u"LINESSEP", repr(os.linesep))
    except Exception:
        pass

    if hasattr(os, u"getuid"):
        try:
            add(u"UID", os.getuid())
        except Exception:
            pass
    if hasattr(os, u"getgid"):
        try:
            add(u"GID", os.getgid())
        except Exception:
            pass

    total, avail = _memory_bytes()
    if total is not None:
        add(u"MEMORY_TOTAL", total)
        add(u"MEMORY_TOTAL_MB", int(total // (1024 * 1024)))
    if avail is not None:
        add(u"MEMORY_AVAILABLE", avail)
        add(u"MEMORY_AVAILABLE_MB", int(avail // (1024 * 1024)))

    # Краткая сводка одной строкой (удобно в журнале / отладке).
    summary = u"%s | %s | user=%s | home=%s" % (
        _safe_platform(u"system") or sys.platform,
        _safe_platform(u"machine"),
        user,
        home,
    )
    add(u"SUMMARY", summary)
    return pairs


def seed_env_os_into_variables_map(put_fn, sheet_key=None, file_key=None):
    """
    Записать __ENV_* и __OS_* через put_fn(name, sheet, file, (text, 'text'), ...).

    put_fn — обычно _merge_source_variables_put из collect_workbooks.
    """
    if put_fn is None or not callable(put_fn):
        return 0
    try:
        from libre_macros_global_settings_lib import (
            GLOBAL_VARIABLES_FILE_KEY,
            GLOBAL_VARIABLES_SHEET_KEY,
        )

        sk = sheet_key if sheet_key is not None else GLOBAL_VARIABLES_SHEET_KEY
        fk = file_key if file_key is not None else GLOBAL_VARIABLES_FILE_KEY
    except Exception:
        sk = sheet_key if sheet_key is not None else u"Глобальные_переменные"
        fk = file_key if file_key is not None else u"Глобальные_переменные"

    n = 0
    for name, text in iter_env_variable_entries():
        try:
            put_fn(
                name,
                sk,
                fk,
                (text, u"text"),
                report=None,
                cell_ref=CELL_REF_ENV,
            )
            n += 1
        except Exception:
            pass
    for name, text in iter_os_variable_entries():
        try:
            put_fn(
                name,
                sk,
                fk,
                (text, u"text"),
                report=None,
                cell_ref=CELL_REF_OS,
            )
            n += 1
        except Exception:
            pass
    return n
