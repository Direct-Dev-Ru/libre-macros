# -*- coding: utf-8 -*-
"""
Проверка «Файлы-Источники» против доверенных корней (strict / off).

Корни = defaults из libre_macros_source_roots_cfg
      + env MERGE_ALLOWED_SOURCE_ROOTS
      + глобальная Merge_Allowed_Source_Roots
      (+ каталог книги, если включено в cfg).
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.720"
import os
import re
import sys

try:
    unicode
except NameError:
    unicode = str

try:
    import libre_macros_source_roots_cfg as _cfg
except Exception:
    _cfg = None


def _u(v):
    try:
        return unicode(v or u"")
    except Exception:
        return str(v or "")


def _log(msg):
    try:
        print(u"[source_roots] %s" % msg)
    except Exception:
        pass


def _is_windows():
    try:
        return os.name == "nt" or sys.platform.startswith("win")
    except Exception:
        return False


def _path_mod():
    """os.path на текущей ОС; ntpath если принудительно Windows-логика."""
    if _is_windows():
        try:
            import ntpath

            return ntpath
        except Exception:
            pass
    return os.path


def _fold_key(s):
    t = _u(s).strip()
    try:
        return t.casefold()
    except Exception:
        return t.lower()


def os_username():
    """Имя пользователя ОС для шаблонов {user}."""
    for key in ("USERNAME", "USER", "LOGNAME"):
        try:
            v = os.environ.get(key)
        except Exception:
            v = None
        if v and _u(v).strip():
            return _u(v).strip()
    try:
        import getpass

        v = getpass.getuser()
        if v and _u(v).strip():
            return _u(v).strip()
    except Exception:
        pass
    return u"user"


def split_roots_text(text):
    """
    Разбор списка корней из env/глобальной.

    Разделители: перевод строки, `;`, и `:` только если дальше абсолютный
    путь (``/…`` или ``X:\\…``) — чтобы не ломать ``C:\\Users``.
    """
    raw = _u(text).replace(u"\r\n", u"\n").replace(u"\r", u"\n")
    if raw.strip() == u"":
        return []
    chunks = []
    for part in re.split(r"[;\n]+", raw):
        part = _u(part).strip()
        if part == u"":
            continue
        # `:`` перед /abs или X:\ — разделитель списка, не диск Windows
        for s in re.split(r":(?=/|[A-Za-z]:[\\/])", part):
            s = _u(s).strip().strip(u'"').strip(u"'")
            if s != u"":
                chunks.append(s)
    return chunks


def normalize_root_path(path):
    """abspath + normpath; пусто → ''. На Windows корень диска остаётся X:\\ ."""
    p = _u(path).strip().strip(u'"').strip(u"'")
    if p == u"":
        return u""
    try:
        p = os.path.expanduser(p)
    except Exception:
        pass
    try:
        p = os.path.expandvars(p)
    except Exception:
        pass
    pm = _path_mod()
    # Windows: сохранить «H:\» (весь том). posixpath.splitdrive не видит букву диска.
    if _is_windows():
        drive, tail = pm.splitdrive(p.replace(u"/", u"\\"))
        if drive and len(drive) == 2 and drive[1] == u":":
            t = _u(tail).lstrip(u"\\/")
            if t == u"":
                return _u(drive) + u"\\"
    try:
        if _is_windows() and pm is not os.path:
            # Не вызывать posix abspath для Windows-путей (склеивает с cwd Linux).
            p = pm.normpath(p.replace(u"/", u"\\"))
        else:
            p = os.path.abspath(p)
            p = os.path.normpath(p)
    except Exception:
        try:
            p = pm.normpath(p)
        except Exception:
            pass
    p = _u(p)
    if _is_windows():
        drive, tail = pm.splitdrive(p.replace(u"/", u"\\"))
        if drive and len(drive) == 2 and drive[1] == u":":
            t = _u(tail).replace(u"/", u"\\")
            if t in (u"", u"\\", u"."):
                return _u(drive) + u"\\"
            if t.startswith(u"\\"):
                return _u(drive) + t
            return _u(drive) + u"\\" + t
    return p


def real_path_safe(path):
    """realpath если возможно, иначе abspath/normpath."""
    p = normalize_root_path(path)
    if p == u"":
        return u""
    try:
        if os.path.exists(p):
            return _u(os.path.normpath(os.path.realpath(p)))
    except Exception:
        pass
    try:
        if _is_windows():
            return normalize_root_path(p)
        return _u(os.path.normpath(os.path.abspath(p)))
    except Exception:
        return p


def _norm_key(path):
    """Ключ без realpath (сохраняет букву диска H:\\ при сравнении с UNC)."""
    p = normalize_root_path(path)
    if p == u"":
        return u""
    if _is_windows():
        return _fold_key(p)
    return p


def _path_key(path):
    """Ключ сравнения: realpath + casefold на Windows."""
    p = real_path_safe(path)
    if p == u"":
        return u""
    if _is_windows():
        return _fold_key(p)
    return p


def _keys_under(pk, rk):
    """pk лежит под rk (оба уже ключи сравнения)."""
    if pk == u"" or rk == u"":
        return False
    if pk == rk:
        return True
    if _is_windows():
        if rk.endswith(u"\\") or rk.endswith(u"/"):
            return pk.startswith(rk)
        return pk.startswith(rk + u"\\") or pk.startswith(rk + u"/")
    sep = os.sep
    if rk.endswith(sep):
        return pk.startswith(rk)
    return pk.startswith(rk + sep)


def path_is_under_root(path, root):
    """
    True, если path == root или path лежит строго внутри root.

    Сравниваем и «буквенную» форму (H:\\…), и realpath (\\\\server\\share\\…):
    иначе корень H:\\ после realpath → UNC, а префикс маски остаётся H:\\… —
    и startswith ломается.
    """
    if _keys_under(_norm_key(path), _norm_key(root)):
        return True
    if _keys_under(_path_key(path), _path_key(root)):
        return True
    if _is_windows():
        if _keys_under(_path_key(path), _norm_key(root)):
            return True
        # Буква диска корня vs буква пути (весь том H: доверен).
        try:
            pm = _path_mod()
            r_drive, r_tail = pm.splitdrive(normalize_root_path(root).replace(u"/", u"\\"))
            p_drive, _p_tail = pm.splitdrive(normalize_root_path(path).replace(u"/", u"\\"))
            if r_drive and p_drive and _fold_key(r_drive) == _fold_key(p_drive):
                rt = _u(r_tail).replace(u"/", u"\\").rstrip(u"\\")
                if rt == u"" or rt == u".":
                    return True
        except Exception:
            pass
    return False


def path_is_under_any_root(path, roots):
    for root in roots or []:
        if path_is_under_root(path, root):
            return True
    return False


def glob_literal_prefix_dir(pattern):
    """
    Каталог-префикс маски до первого сегмента с * ? [.
    Для ``C:\\Users\\me\\data\\*.xlsx`` → ``C:\\Users\\me\\data``.
    Для ``C:\\Users\\me\\**\\*.xlsx`` → ``C:\\Users\\me``.
    Для ``H:\\SOME1\\SOME2\\*.xlsx`` → ``H:\\SOME1\\SOME2``.
    """
    raw = _u(pattern).strip().strip(u'"').strip(u"'")
    if raw == u"":
        return u""
    p = raw
    try:
        p = os.path.expanduser(p)
    except Exception:
        pass
    try:
        p = os.path.expandvars(p)
    except Exception:
        pass
    p = _u(p)
    pm = _path_mod()
    if _is_windows():
        p = p.replace(u"/", u"\\")
        drive, tail = pm.splitdrive(p)
        sep = u"\\"
    else:
        drive, tail = os.path.splitdrive(p)
        sep = os.sep
        if not drive and (p.startswith(u"/") or (os.altsep and p.startswith(os.altsep))):
            drive = u""
            tail = p
    parts = []
    raw_parts = _u(tail).replace(u"/", sep).split(sep)
    i = 0
    while i < len(raw_parts):
        part = raw_parts[i]
        i = i + 1
        if part == u"":
            continue
        if part == u"**" or any(c in part for c in u"*?["):
            break
        parts.append(part)
    if not parts:
        if drive:
            return normalize_root_path(drive + sep)
        if p.startswith(os.sep) or (os.altsep and p.startswith(os.altsep)):
            return normalize_root_path(os.sep)
        return u""
    joined = sep.join(parts)
    if drive:
        out = drive + sep + joined
    elif p.startswith(os.sep) or (os.altsep and p.startswith(os.altsep)):
        out = sep + joined
    else:
        out = joined
    return normalize_root_path(out)


def default_os_roots(username=None):
    """Корни из шаблонов cfg для текущей ОС."""
    user = _u(username or os_username()).strip() or u"user"
    templates = ()
    if _cfg is not None:
        if _is_windows():
            templates = getattr(_cfg, u"SOURCE_ROOTS_WINDOWS_TEMPLATES", ()) or ()
        else:
            templates = getattr(_cfg, u"SOURCE_ROOTS_LINUX_TEMPLATES", ()) or ()
    else:
        if _is_windows():
            templates = (
                u"C:\\Users\\{user}",
                u"S:\\{user}",
                u"H:\\",
            )
        else:
            templates = (u"/home/{user}", u"/home/TN/{user}")
    out = []
    for tmpl in templates:
        try:
            path = _u(tmpl).format(user=user)
        except Exception:
            path = _u(tmpl).replace(u"{user}", user)
        path = normalize_root_path(path)
        if path != u"":
            out.append(path)
    return out


def fixed_extra_roots():
    if _cfg is None:
        return []
    out = []
    for p in getattr(_cfg, u"SOURCE_ROOTS_EXTRA_FIXED", ()) or ():
        np = normalize_root_path(p)
        if np != u"":
            out.append(np)
    return out


def lookup_global_text(global_name, variables_map=None):
    """Текст глобальной переменной из runtime-карты (сегмент Глобальные_переменные)."""
    try:
        from libre_macros_allow_gate_lib import lookup_global_from_runtime_map

        found, value = lookup_global_from_runtime_map(global_name, variables_map)
        if found:
            return _u(value)
    except Exception as err:
        _log(u"global lookup %s: %s" % (global_name, err))
    return u""


def env_text(env_name):
    name = _u(env_name).strip()
    if name == u"":
        return u""
    try:
        raw = os.environ.get(name)
    except Exception:
        raw = None
    if raw is None:
        return u""
    return _u(raw).strip()


def resolve_mode(variables_map=None):
    """
    Итоговый режим: off | strict.

    Приоритет: env MERGE_SOURCE_ROOTS_MODE → глобальная Merge_Source_Roots_Mode → cfg default.
    """
    env_name = getattr(_cfg, u"SOURCE_ROOTS_ENV_MODE", u"MERGE_SOURCE_ROOTS_MODE") if _cfg else u"MERGE_SOURCE_ROOTS_MODE"
    g_name = getattr(_cfg, u"SOURCE_ROOTS_GLOBAL_MODE", u"Merge_Source_Roots_Mode") if _cfg else u"Merge_Source_Roots_Mode"
    default = getattr(_cfg, u"SOURCE_ROOTS_DEFAULT_MODE", u"strict") if _cfg else u"strict"

    raw = env_text(env_name)
    if raw == u"":
        raw = lookup_global_text(g_name, variables_map)
    if raw == u"":
        raw = _u(default)
    key = _fold_key(raw).replace(u"-", u"").replace(u"_", u"")
    if key in (u"off", u"0", u"false", u"no", u"disable", u"disabled", u"выкл", u"выключено"):
        return u"off"
    return u"strict"


def collect_allowed_roots(variables_map=None, book_path=None):
    """
    Полный список доверенных корней (нормированные пути, без дублей).
    """
    roots = []
    seen = set()

    def _add(path):
        np = normalize_root_path(path)
        if np == u"":
            return
        key = _path_key(np)
        if key == u"" or key in seen:
            return
        seen.add(key)
        roots.append(np)

    for p in default_os_roots():
        _add(p)
    for p in fixed_extra_roots():
        _add(p)

    env_name = getattr(_cfg, u"SOURCE_ROOTS_ENV_ROOTS", u"MERGE_ALLOWED_SOURCE_ROOTS") if _cfg else u"MERGE_ALLOWED_SOURCE_ROOTS"
    g_name = getattr(_cfg, u"SOURCE_ROOTS_GLOBAL_ROOTS", u"Merge_Allowed_Source_Roots") if _cfg else u"Merge_Allowed_Source_Roots"
    for p in split_roots_text(env_text(env_name)):
        _add(p)
    for p in split_roots_text(lookup_global_text(g_name, variables_map)):
        _add(p)

    include_book = True
    if _cfg is not None:
        include_book = bool(getattr(_cfg, u"SOURCE_ROOTS_INCLUDE_BOOK_DIR", True))
    if include_book and book_path:
        bp = _u(book_path).strip()
        if bp and (not _is_marker(bp)):
            try:
                d = os.path.dirname(real_path_safe(bp) or bp)
            except Exception:
                d = u""
            if d:
                _add(d)
    return roots


def _is_marker(path):
    try:
        from libre_macros_collect_cfg import merge_normalize_current_book_marker_key, _MERGE_CURRENT_BOOK_MARKER_KEYS

        key = merge_normalize_current_book_marker_key(path)
        return key in _MERGE_CURRENT_BOOK_MARKER_KEYS
    except Exception:
        p = _fold_key(path).replace(u"_", u"").replace(u" ", u"").replace(u"-", u"").replace(u".", u"")
        return p in (
            u"этакнига",
            u"thisworkbook",
            u"thisbook",
            u"currentbook",
            u"currentworkbook",
            u"current",
        )


def allow_remote():
    if _cfg is None:
        return False
    return bool(getattr(_cfg, u"SOURCE_ROOTS_ALLOW_REMOTE", False))


def is_remote_url(path):
    try:
        from libre_macros_bundle_paths import is_remote_url as _fn

        return bool(_fn(path))
    except Exception:
        p = _u(path).strip()
        if p == u"" or p.lower().startswith("file://"):
            return False
        return u"://" in p


def check_source_path(path, roots, mode=u"strict", is_current_book=False):
    """
    Проверка одного пути (не glob).

    Возвращает (ok: bool, error_or_empty: unicode).
    """
    if _fold_key(mode) == u"off":
        return True, u""
    if is_current_book or _is_marker(path):
        return True, u""
    p = _u(path).strip()
    if p == u"":
        return True, u""
    if is_remote_url(p):
        if allow_remote():
            return True, u""
        return False, (
            u"Удалённый URL источника запрещён режимом доверенных корней (strict):\n%s\n"
            u"Разрешите remote в libre_macros_source_roots_cfg.SOURCE_ROOTS_ALLOW_REMOTE "
            u"или укажите локальный путь."
            % p
        )
    if path_is_under_any_root(p, roots):
        return True, u""
    roots_txt = u"\n".join(u"  • %s" % r for r in (roots or [])[:20])
    if len(roots or []) > 20:
        roots_txt += u"\n  • …"
    return False, (
        u"Путь источника вне доверенных корней:\n%s\n\n"
        u"Разрешённые корни:\n%s\n\n"
        u"Добавьте корень в глобальную «%s» или env %s "
        u"(либо режим off: %s / %s)."
        % (
            p,
            roots_txt or u"  (пусто)",
            getattr(_cfg, u"SOURCE_ROOTS_GLOBAL_ROOTS", u"Merge_Allowed_Source_Roots") if _cfg else u"Merge_Allowed_Source_Roots",
            getattr(_cfg, u"SOURCE_ROOTS_ENV_ROOTS", u"MERGE_ALLOWED_SOURCE_ROOTS") if _cfg else u"MERGE_ALLOWED_SOURCE_ROOTS",
            getattr(_cfg, u"SOURCE_ROOTS_ENV_MODE", u"MERGE_SOURCE_ROOTS_MODE") if _cfg else u"MERGE_SOURCE_ROOTS_MODE",
            getattr(_cfg, u"SOURCE_ROOTS_GLOBAL_MODE", u"Merge_Source_Roots_Mode") if _cfg else u"Merge_Source_Roots_Mode",
        )
    )


def check_glob_pattern(pattern, roots, mode=u"strict"):
    """
    Перед разворотом glob: literal-prefix каталог должен быть под корнем.
    Иначе отказ без обхода ФС.
    """
    if _fold_key(mode) == u"off":
        return True, u""
    p = _u(pattern).strip()
    if p == u"":
        return True, u""
    if is_remote_url(p):
        return check_source_path(p, roots, mode=mode)
    prefix = glob_literal_prefix_dir(p)
    if prefix == u"":
        return False, (
            u"Маска источника слишком широкая (нет доверенного префикса-каталога):\n%s\n"
            u"Укажите путь вида «корень\\папка\\*.xlsx», где корень в списке доверенных."
            % p
        )
    if path_is_under_any_root(prefix, roots):
        return True, u""
    # Если префикс сам является одним из корней — ок (уже покрыто).
    # Иначе — запрет.
    roots_txt = u"\n".join(u"  • %s" % r for r in (roots or [])[:20])
    return False, (
        u"Маска источника вне доверенных корней (префикс «%s»):\n%s\n\n"
        u"Разрешённые корни:\n%s"
        % (prefix, p, roots_txt or u"  (пусто)")
    )


def validate_expanded_sources(files, is_current_book_flags=None, book_path=None, variables_map=None):
    """
    Проверка списка после expand_source_files.

    Возвращает unicode-ошибку или пустую строку / None.
    """
    mode = resolve_mode(variables_map)
    if mode == u"off":
        return None
    roots = collect_allowed_roots(variables_map=variables_map, book_path=book_path)
    flags = is_current_book_flags or []
    errors = []
    i = 0
    while i < len(files or []):
        f = files[i]
        is_curr = False
        if i < len(flags):
            is_curr = bool(flags[i])
        ok, err = check_source_path(f, roots, mode=mode, is_current_book=is_curr)
        if not ok and err:
            errors.append(err)
        i += 1
    if not errors:
        return None
    # Одна сводка без дублей полного текста корней на каждый файл
    denied = []
    i = 0
    while i < len(files or []):
        f = files[i]
        is_curr = bool(flags[i]) if i < len(flags) else False
        if is_curr or _is_marker(f):
            i += 1
            continue
        if is_remote_url(f) and (not allow_remote()):
            denied.append(_u(f))
        elif not path_is_under_any_root(f, roots):
            denied.append(_u(f))
        i += 1
    if not denied:
        return errors[0]
    roots_txt = u"\n".join(u"  • %s" % r for r in roots[:24])
    if len(roots) > 24:
        roots_txt += u"\n  • …"
    lines = u"\n".join(u"  • %s" % d for d in denied[:40])
    if len(denied) > 40:
        lines += u"\n  • … (+%d)" % (len(denied) - 40)
    g_roots = getattr(_cfg, u"SOURCE_ROOTS_GLOBAL_ROOTS", u"Merge_Allowed_Source_Roots") if _cfg else u"Merge_Allowed_Source_Roots"
    e_roots = getattr(_cfg, u"SOURCE_ROOTS_ENV_ROOTS", u"MERGE_ALLOWED_SOURCE_ROOTS") if _cfg else u"MERGE_ALLOWED_SOURCE_ROOTS"
    e_mode = getattr(_cfg, u"SOURCE_ROOTS_ENV_MODE", u"MERGE_SOURCE_ROOTS_MODE") if _cfg else u"MERGE_SOURCE_ROOTS_MODE"
    g_mode = getattr(_cfg, u"SOURCE_ROOTS_GLOBAL_MODE", u"Merge_Source_Roots_Mode") if _cfg else u"Merge_Source_Roots_Mode"
    return (
        u"Источники вне доверенных корней (режим strict):\n%s\n\n"
        u"Разрешённые корни:\n%s\n\n"
        u"Добавить корень: глобальная «%s» или env %s.\n"
        u"Отключить проверку: %s=off или глобальная «%s»=off."
        % (lines, roots_txt or u"  (пусто)", g_roots, e_roots, e_mode, g_mode)
    )


def guard_glob_or_path(resolved_path, is_glob, book_path=None, variables_map=None):
    """
    Проверка до разворота / добавления файла.
    Возвращает (ok, error).
    """
    mode = resolve_mode(variables_map)
    if mode == u"off":
        return True, u""
    roots = collect_allowed_roots(variables_map=variables_map, book_path=book_path)
    if is_glob:
        return check_glob_pattern(resolved_path, roots, mode=mode)
    return check_source_path(resolved_path, roots, mode=mode, is_current_book=False)
