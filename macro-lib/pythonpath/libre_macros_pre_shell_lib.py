# -*- coding: utf-8 -*-
"""
Предварительный скрипт перед сбором источников (subprocess без shell=True).

JSON в колонке B параметра «Предварительный_скрипт»:
  {"argv":["/bin/bash","/path/to/run.sh"],"cwd":"/tmp","timeout":60,
   "clean_env":true,"env":{"K":"V"}}
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.690"
import datetime
import json
import os
import re
import subprocess
import sys

import libre_macros_pre_shell_cfg as _cfg

try:
    import getpass
except ImportError:
    getpass = None

try:
    import uno
    import unohelper
    from com.sun.star.awt import XActionListener
except Exception:
    uno = None
    unohelper = None
    XActionListener = None

try:
    unicode
except NameError:
    unicode = str


def _log(msg):
    try:
        print(u"[pre_shell] %s" % msg)
    except Exception:
        pass


def clip_pre_shell_text(text, limit=1200):
    s = unicode(text or u"")
    if len(s) <= limit:
        return s
    return s[:limit] + u"\n… (обрезано)"


def parse_pre_shell_dict(raw):
    """
    Разбор JSON параметра → dict или None (пустая ячейка / невалидно / нет argv).

    Возвращает нормализованный dict с ключами argv, cwd, timeout, env, clean_env
    либо None, если запускать нечего.
    """
    if raw is None:
        return None
    if isinstance(raw, dict):
        obj = raw
    else:
        text = unicode(raw).strip()
        if text == u"":
            return None
        try:
            obj = json.loads(text)
        except Exception:
            return None
        if not isinstance(obj, dict):
            return None
    argv_raw = obj.get(_cfg.PRE_SHELL_KEY_ARGV)
    if argv_raw is None:
        argv_raw = obj.get(u"args")
    argv = _normalize_argv(argv_raw)
    if not argv:
        return None
    cwd = unicode(obj.get(_cfg.PRE_SHELL_KEY_CWD) or u"").strip()
    timeout = _normalize_timeout(obj.get(_cfg.PRE_SHELL_KEY_TIMEOUT))
    env = _normalize_env(obj.get(_cfg.PRE_SHELL_KEY_ENV))
    clean_env = _normalize_bool(
        obj.get(_cfg.PRE_SHELL_KEY_CLEAN_ENV),
        default=_cfg.PRE_SHELL_DEFAULT_CLEAN_ENV,
    )
    return {
        _cfg.PRE_SHELL_KEY_ARGV: argv,
        _cfg.PRE_SHELL_KEY_CWD: cwd,
        _cfg.PRE_SHELL_KEY_TIMEOUT: timeout,
        _cfg.PRE_SHELL_KEY_ENV: env,
        _cfg.PRE_SHELL_KEY_CLEAN_ENV: clean_env,
    }


def encode_pre_shell_json(spec):
    """Сериализация dict → компактный JSON (ensure_ascii=False)."""
    if not spec:
        return u""
    argv = _normalize_argv(spec.get(_cfg.PRE_SHELL_KEY_ARGV))
    if not argv:
        return u""
    payload = {_cfg.PRE_SHELL_KEY_ARGV: argv}
    cwd = unicode(spec.get(_cfg.PRE_SHELL_KEY_CWD) or u"").strip()
    if cwd:
        payload[_cfg.PRE_SHELL_KEY_CWD] = cwd
    timeout = _normalize_timeout(spec.get(_cfg.PRE_SHELL_KEY_TIMEOUT))
    if timeout is not None:
        payload[_cfg.PRE_SHELL_KEY_TIMEOUT] = timeout
    env = _normalize_env(spec.get(_cfg.PRE_SHELL_KEY_ENV))
    if env:
        payload[_cfg.PRE_SHELL_KEY_ENV] = env
    clean_env = _normalize_bool(
        spec.get(_cfg.PRE_SHELL_KEY_CLEAN_ENV),
        default=_cfg.PRE_SHELL_DEFAULT_CLEAN_ENV,
    )
    # Явно пишем, чтобы в книге было видно поведение по умолчанию.
    payload[_cfg.PRE_SHELL_KEY_CLEAN_ENV] = bool(clean_env)
    try:
        return unicode(json.dumps(payload, ensure_ascii=False, separators=(u",", u":")))
    except Exception:
        try:
            return unicode(json.dumps(payload, ensure_ascii=False))
        except Exception:
            return u""


def _normalize_argv(raw):
    if raw is None:
        return []
    if isinstance(raw, (list, tuple)):
        out = []
        i = 0
        while i < len(raw):
            s = unicode(raw[i]).strip()
            if s != u"":
                out.append(s)
            i = i + 1
        return out
    text = unicode(raw).strip()
    if text == u"":
        return []
    # Одна строка: разобрать как JSON-массив или по строкам.
    if text.startswith(u"["):
        try:
            arr = json.loads(text)
            return _normalize_argv(arr)
        except Exception:
            pass
    lines = text.replace(u"\r\n", u"\n").replace(u"\r", u"\n").split(u"\n")
    out = []
    i = 0
    while i < len(lines):
        s = unicode(lines[i]).strip()
        if s != u"":
            out.append(s)
        i = i + 1
    return out


def _normalize_timeout(raw):
    if raw is None or raw is False:
        return None
    if isinstance(raw, bool):
        return None
    text = unicode(raw).strip()
    if text == u"" or text.lower() in (u"null", u"none", u"-"):
        return None
    try:
        val = float(text) if not isinstance(raw, (int, float)) else float(raw)
    except Exception:
        return None
    if val <= 0:
        return None
    if val == int(val):
        return int(val)
    return val


def _normalize_env(raw):
    out = {}
    if raw is None:
        return out
    if isinstance(raw, dict):
        for key in raw:
            k = unicode(key).strip()
            if k == u"":
                continue
            out[k] = unicode(raw[key])
        return out
    text = unicode(raw).strip()
    if text == u"":
        return out
    if text.startswith(u"{"):
        try:
            obj = json.loads(text)
            if isinstance(obj, dict):
                return _normalize_env(obj)
        except Exception:
            pass
    lines = text.replace(u"\r\n", u"\n").replace(u"\r", u"\n").split(u"\n")
    i = 0
    while i < len(lines):
        line = unicode(lines[i]).strip()
        i = i + 1
        if line == u"" or line.startswith(u"#"):
            continue
        if u"=" not in line:
            continue
        key, val = line.split(u"=", 1)
        key = key.strip()
        if key == u"":
            continue
        out[key] = val
    return out


def _normalize_bool(raw, default=True):
    if raw is None:
        return bool(default)
    if isinstance(raw, bool):
        return raw
    text = unicode(raw).strip().lower()
    if text == u"":
        return bool(default)
    if text in (u"1", u"true", u"yes", u"on", u"y", u"да"):
        return True
    if text in (u"0", u"false", u"no", u"off", u"n", u"нет"):
        return False
    return bool(default)


def argv_lines_to_text(argv):
    return u"\n".join([unicode(x) for x in (argv or [])])


def env_dict_to_text(env):
    if not env:
        return u""
    keys = sorted(list(env.keys()))
    lines = []
    i = 0
    while i < len(keys):
        k = keys[i]
        lines.append(u"%s=%s" % (k, env[k]))
        i = i + 1
    return u"\n".join(lines)


def _path_looks_like_office(part):
    p = unicode(part or u"").casefold()
    if p == u"":
        return False
    markers = getattr(_cfg, u"PRE_SHELL_OFFICE_PATH_MARKERS", ()) or ()
    i = 0
    while i < len(markers):
        if unicode(markers[i]).casefold() in p:
            return True
        i = i + 1
    return False


def _sanitize_path_value(raw):
    """Убрать из PATH/LD_LIBRARY_PATH сегменты AlterOffice/LibreOffice."""
    text = unicode(raw or u"")
    if text == u"":
        return u""
    sep = u";" if (u";" in text and u":" not in text) else u":"
    parts = text.split(sep)
    kept = []
    i = 0
    while i < len(parts):
        part = parts[i]
        i = i + 1
        if part == u"":
            continue
        if _path_looks_like_office(part):
            continue
        kept.append(part)
    return sep.join(kept)


def sanitize_env_from_office(env):
    """
    Очистить окружение AO/LO, чтобы дочерний /usr/bin/python3 и ldapsearch
    не подхватывали встроенный python-core (иначе TypeError fork_exec).
    """
    if env is None:
        env = {}
    out = dict(env)
    keys = getattr(_cfg, u"PRE_SHELL_UNSET_ENV_KEYS", ()) or ()
    i = 0
    while i < len(keys):
        out.pop(unicode(keys[i]), None)
        i = i + 1
    if u"PATH" in out:
        out[u"PATH"] = _sanitize_path_value(out.get(u"PATH"))
    if u"LD_LIBRARY_PATH" in out:
        cleaned = _sanitize_path_value(out.get(u"LD_LIBRARY_PATH"))
        if cleaned:
            out[u"LD_LIBRARY_PATH"] = cleaned
        else:
            out.pop(u"LD_LIBRARY_PATH", None)
    # На всякий случай не тащим UNO bootstrap в обычные CLI-скрипты.
    for extra_key in (u"URE_BOOTSTRAP", u"UNO_PATH", u"STAR_PROFILEHOME"):
        out.pop(extra_key, None)
    return out


def build_env_for_run(extra_env, clean_env=True):
    env = os.environ.copy()
    if clean_env:
        env = sanitize_env_from_office(env)
    extra = extra_env or {}
    for key in extra:
        env[unicode(key)] = unicode(extra[key])
    return env


def pre_shell_consent_log_dir():
    """~/.config/libre-macros/pre_shell_consent/"""
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.environ.get(
            "XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config")
        )
    path = os.path.join(base, "libre-macros", _cfg.PRE_SHELL_CONSENT_SUBDIR)
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


def _pre_shell_os_user():
    user = u""
    try:
        if getpass is not None:
            user = unicode(getpass.getuser() or u"")
    except Exception:
        pass
    if user == u"":
        user = unicode(os.environ.get("USER") or os.environ.get("USERNAME") or u"unknown")
    return user


def _pre_shell_safe_filename_part(text, default=u"unknown"):
    s = unicode(text or u"").strip()
    if s == u"":
        s = unicode(default)
    s = re.sub(r"[^\w\-.@]+", u"_", s, flags=re.UNICODE)
    s = s.strip(u"._")
    if s == u"":
        s = unicode(default)
    return s[:80]


def _pre_shell_doc_book_basename(doc):
    if doc is None:
        return u"book"
    try:
        url = unicode(doc.getPropertyValue(u"URL") or u"")
    except Exception:
        url = u""
    if url.startswith(u"file://"):
        try:
            import uno as _uno
            path = _uno.fileUrlToSystemPath(url)
            return _pre_shell_safe_filename_part(os.path.basename(path), u"book")
        except Exception:
            pass
    try:
        title = unicode(doc.getPropertyValue(u"Title") or u"")
        if title.strip():
            return _pre_shell_safe_filename_part(title, u"book")
    except Exception:
        pass
    return u"book"


def _pre_shell_mask_env_for_log(env):
    out = {}
    if not env:
        return out
    markers = getattr(_cfg, u"PRE_SHELL_ENV_SECRET_MARKERS", ()) or ()
    for key in env:
        k = unicode(key)
        kl = k.casefold()
        secret = False
        i = 0
        while i < len(markers):
            if unicode(markers[i]).casefold() in kl:
                secret = True
                break
            i = i + 1
        out[k] = u"***" if secret else unicode(env[key])
    return out


def format_pre_shell_run_description(spec):
    """Краткое описание команды для диалога."""
    if not spec:
        return u"(не задано)"
    argv = spec.get(_cfg.PRE_SHELL_KEY_ARGV) or []
    lines = [u" ".join([unicode(x) for x in argv])]
    cwd = unicode(spec.get(_cfg.PRE_SHELL_KEY_CWD) or u"").strip()
    if cwd:
        lines.append(u"Каталог: %s" % cwd)
    timeout = spec.get(_cfg.PRE_SHELL_KEY_TIMEOUT)
    if timeout is not None:
        lines.append(u"Таймаут: %s сек" % timeout)
    env = spec.get(_cfg.PRE_SHELL_KEY_ENV) or {}
    if env:
        masked = _pre_shell_mask_env_for_log(env)
        keys = sorted(list(masked.keys()))
        parts = []
        i = 0
        while i < len(keys):
            parts.append(u"%s=%s" % (keys[i], masked[keys[i]]))
            i = i + 1
        lines.append(u"env: %s" % u", ".join(parts))
    clean_env = spec.get(_cfg.PRE_SHELL_KEY_CLEAN_ENV)
    if clean_env is not False:
        lines.append(u"Очистка env AO/LO: да")
    return u"\n".join(lines)


def build_pre_shell_sheet_summary(doc, table, param_sheet_name=u""):
    """Сводка параметров листа (без паролей)."""
    lines = []
    if param_sheet_name:
        lines.append(u"Лист параметров: %s" % unicode(param_sheet_name))
    if doc is not None:
        lines.append(u"Книга: %s" % _pre_shell_doc_book_basename(doc))
    if not table:
        return u"\n".join(lines)

    def _table_values(key):
        key_norm = unicode(key or u"").strip().casefold().replace(u"_", u"").replace(u"-", u"")
        for pname in table:
            pn = unicode(pname or u"").strip().casefold().replace(u"_", u"").replace(u"-", u"")
            if pn == key_norm:
                return table[pname]
        return None

    def _first(key):
        row = _table_values(key)
        if not row:
            return u""
        i = 0
        while i < len(row):
            s = unicode(row[i] or u"").strip()
            if s != u"":
                return s
            i = i + 1
        return u""

    mode = _first(u"Режим")
    if mode:
        lines.append(u"Режим: %s" % mode)
    files_row = _table_values(u"Файлы-Источники")
    if files_row:
        shown = []
        i = 0
        while i < len(files_row) and i < 6:
            s = unicode(files_row[i] or u"").strip()
            if s:
                shown.append(s)
            i = i + 1
        if shown:
            lines.append(u"Файлы-Источники: %s" % u"; ".join(shown))
    pre_raw = _first(_cfg.P_MERGE_PRE_SHELL)
    if pre_raw:
        spec = parse_pre_shell_dict(pre_raw)
        if spec:
            lines.append(u"Предварительный_скрипт:")
            lines.append(format_pre_shell_run_description(spec))
    return u"\n".join(lines)


def log_pre_shell_consent_event(doc, spec, table, param_sheet_name, confirmed, macro_version=u""):
    """Запись согласия/отказа в ~/.config/libre-macros/pre_shell_consent/."""
    now = datetime.datetime.now()
    user = _pre_shell_os_user()
    book = _pre_shell_doc_book_basename(doc)
    fname = u"%s_%s_%s_%s.json" % (
        now.strftime("%Y%m%d_%H%M%S"),
        _pre_shell_safe_filename_part(user, u"user"),
        book,
        u"ok" if confirmed else u"cancel",
    )
    path = os.path.join(pre_shell_consent_log_dir(), fname)
    payload = {
        u"timestamp": now.strftime("%Y-%m-%dT%H:%M:%S"),
        u"user": user,
        u"book_basename": book,
        u"param_sheet": unicode(param_sheet_name or u""),
        u"confirmed": bool(confirmed),
        u"macro_version": unicode(macro_version or u""),
        u"pre_shell": {},
        u"param_sheet_summary": build_pre_shell_sheet_summary(doc, table, param_sheet_name),
    }
    if spec:
        payload[u"pre_shell"] = {
            _cfg.PRE_SHELL_KEY_ARGV: list(spec.get(_cfg.PRE_SHELL_KEY_ARGV) or []),
            _cfg.PRE_SHELL_KEY_CWD: unicode(spec.get(_cfg.PRE_SHELL_KEY_CWD) or u""),
            _cfg.PRE_SHELL_KEY_TIMEOUT: spec.get(_cfg.PRE_SHELL_KEY_TIMEOUT),
            _cfg.PRE_SHELL_KEY_CLEAN_ENV: spec.get(_cfg.PRE_SHELL_KEY_CLEAN_ENV),
            _cfg.PRE_SHELL_KEY_ENV: _pre_shell_mask_env_for_log(spec.get(_cfg.PRE_SHELL_KEY_ENV)),
        }
    try:
        with open(path, "wb") as f:
            f.write(json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"))
        _log(u"consent log: %s" % path)
        try:
            from libre_macros_global_settings_lib import notify_config_changed

            notify_config_changed()
        except Exception:
            pass
        return path
    except Exception as err:
        _log(u"consent log failed: %s" % err)
        return None


def _pre_shell_dialog_create_peer(dlg, toolkit, doc=None):
    parents = []
    if doc is not None:
        try:
            w = doc.getCurrentController().getFrame().getContainerWindow()
            if w is not None:
                parents.append(w)
                try:
                    peer = w.getPeer()
                    if peer is not None:
                        parents.append(peer)
                except Exception:
                    pass
        except Exception:
            pass
    parents.append(None)
    i = 0
    while i < len(parents):
        try:
            dlg.createPeer(toolkit, parents[i])
            return True
        except Exception:
            pass
        i = i + 1
    return False


def _pre_shell_dlg_add_fixed(dm, name, label, x, y, w, h, multiline=False):
    m = dm.createInstance(u"com.sun.star.awt.UnoControlFixedTextModel")
    m.Name = unicode(name)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    m.Label = unicode(label)
    if multiline:
        m.MultiLine = True
    dm.insertByName(unicode(name), m)
    return m


def _pre_shell_dlg_add_edit(dm, name, x, y, w, h, multiline=False, readonly=False):
    edit = dm.createInstance(u"com.sun.star.awt.UnoControlEditModel")
    edit.Name = unicode(name)
    edit.PositionX = int(x)
    edit.PositionY = int(y)
    edit.Width = int(w)
    edit.Height = int(h)
    if multiline:
        edit.MultiLine = True
        try:
            edit.VScroll = True
            edit.HScroll = False
        except Exception:
            pass
    if readonly:
        edit.ReadOnly = True
    dm.insertByName(unicode(name), edit)
    return edit


def confirm_pre_shell_before_run(doc, spec, table=None, param_sheet_name=u"", macro_version=u""):
    """
    Диалог «Внимание!!!» перед запуском предварительного скрипта.
    Возвращает True — продолжить; False — отмена (с записью в журнал).
    """
    if spec is None:
        return True
    if uno is None or XActionListener is None:
        _log(u"confirm: UNO недоступен, пропуск диалога")
        return True
    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(u"com.sun.star.awt.Toolkit", ctx)
    except Exception as err:
        _log(u"confirm: toolkit error: %s" % err)
        return True
    try:
        from libre_macros_ui_theme import (
            inner_dialog_add_ok_cancel_footer,
            inner_dialog_footer_y,
            inner_dialog_set_sizeable,
            prepare_dialog_soft_gray_warning,
        )
        import libre_macros_param_wizard_cfg as _pw_cfg
    except Exception as err:
        _log(u"confirm: ui_theme import error: %s" % err)
        return True

    run_desc = format_pre_shell_run_description(spec)
    sheet_summary = build_pre_shell_sheet_summary(doc, table, param_sheet_name)
    body = (
        u"Будет выполнена внешняя команда от вашего имени.\n\n"
        u"Команда:\n%s\n\n"
        u"Продолжайте только если полностью уверены в безопасности скрипта\n"
        u"и доверяете его автору.\n\n"
        u"--- параметры листа ---\n%s"
    ) % (run_desc, sheet_summary)

    dm = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialogModel", ctx)
    m = 10
    dw = 560
    try:
        dh = int(_pw_cfg._DLG_H)
    except Exception:
        dh = 420
    dm.PositionX = 90
    dm.PositionY = 70
    dm.Width = dw
    dm.Height = dh
    dm.Title = u"Внимание!!! — предварительный скрипт"
    inner_dialog_set_sizeable(dm)

    y = m
    hint = (
        u"Будет выполнена внешняя команда от вашего имени. "
        u"Продолжайте только если доверяете скрипту и понимаете риски."
    )
    _pre_shell_dlg_add_fixed(dm, u"HintLbl", hint, m, y, dw - 2 * m, 36, multiline=True)
    y += 40
    edit_h = inner_dialog_footer_y(dh, m) - y - 8
    if edit_h < 120:
        edit_h = 120
    _pre_shell_dlg_add_edit(
        dm, u"BodyEd", m, y, dw - 2 * m, edit_h, multiline=True, readonly=True,
    )
    inner_dialog_add_ok_cancel_footer(
        dm, m, dh, ok_label=u"Продолжить", cancel_label=u"Отмена",
    )

    dlg = sm.createInstanceWithContext(u"com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    prepare_dialog_soft_gray_warning(dlg, title_text=u"Внимание!!!")
    if not _pre_shell_dialog_create_peer(dlg, toolkit, doc=doc):
        _log(u"confirm: createPeer failed")
        return True
    try:
        dlg.getControl(u"BodyEd").setText(body)
    except Exception:
        pass

    rc = dlg.execute()
    try:
        rc = int(rc)
    except Exception:
        rc = 0
    confirmed = (rc == 1)
    log_pre_shell_consent_event(
        doc, spec, table, param_sheet_name, confirmed, macro_version=macro_version,
    )
    return confirmed


def run_pre_shell(spec):
    """
    Запуск argv из spec. Возвращает dict: argv, cwd, returncode, stdout, stderr.
    Бросает RuntimeError/IOError/ValueError при таймауте и прочих сбоях запуска.
    """
    if not spec:
        raise ValueError(u"Спецификация предварительного скрипта пуста")
    argv = list(spec.get(_cfg.PRE_SHELL_KEY_ARGV) or [])
    if not argv:
        raise ValueError(u"argv пуст")
    argv = [unicode(x) for x in argv]
    cwd = unicode(spec.get(_cfg.PRE_SHELL_KEY_CWD) or u"").strip() or None
    timeout = spec.get(_cfg.PRE_SHELL_KEY_TIMEOUT)
    if timeout is not None:
        try:
            timeout = float(timeout)
        except Exception:
            timeout = None
    if cwd is not None and (not os.path.isdir(cwd)):
        raise IOError(u"Рабочий каталог не найден: %s" % cwd)
    exe = argv[0]
    if (u"/" in exe or (len(exe) >= 2 and exe[1] == u":")) and (not os.path.isfile(exe)):
        raise IOError(u"Исполняемый файл не найден: %s" % exe)

    _log(u"argv=%r cwd=%r timeout=%r" % (argv, cwd, timeout))
    clean_env = _normalize_bool(
        spec.get(_cfg.PRE_SHELL_KEY_CLEAN_ENV),
        default=_cfg.PRE_SHELL_DEFAULT_CLEAN_ENV,
    )
    kwargs = {
        u"args": argv,
        u"cwd": cwd,
        u"env": build_env_for_run(spec.get(_cfg.PRE_SHELL_KEY_ENV), clean_env=clean_env),
        u"stdout": subprocess.PIPE,
        u"stderr": subprocess.PIPE,
        u"shell": False,
    }
    if sys.version_info[0] >= 3:
        kwargs[u"universal_newlines"] = True
    try:
        if timeout is not None and hasattr(subprocess, u"run"):
            completed = subprocess.run(timeout=timeout, **kwargs)
            returncode = completed.returncode
            stdout = completed.stdout or u""
            stderr = completed.stderr or u""
        else:
            proc = subprocess.Popen(**kwargs)
            stdout, stderr = proc.communicate()
            returncode = proc.returncode
            if stdout is None:
                stdout = u""
            if stderr is None:
                stderr = u""
            if not isinstance(stdout, unicode):
                stdout = stdout.decode(u"utf-8", u"replace")
            if not isinstance(stderr, unicode):
                stderr = stderr.decode(u"utf-8", u"replace")
    except Exception as err:
        if type(err).__name__ == u"TimeoutExpired":
            raise RuntimeError(
                u"Таймаут предварительного скрипта (%s сек): %s" % (timeout, u" ".join(argv))
            )
        raise

    return {
        u"argv": argv,
        u"cwd": cwd or u"",
        u"returncode": int(returncode if returncode is not None else -1),
        u"stdout": unicode(stdout),
        u"stderr": unicode(stderr),
    }


def format_pre_shell_failure(prefix, result=None, err=None):
    """Текст ошибки для MessageBox / лога сбора."""
    parts = [unicode(prefix or u"Предварительный_скрипт")]
    if err is not None:
        parts.append(u"%s: %s" % (type(err).__name__, err))
    if result is not None:
        parts.append(u"argv: %s" % u" ".join(result.get(u"argv") or []))
        cwd = result.get(u"cwd") or u""
        if cwd:
            parts.append(u"cwd: %s" % cwd)
        parts.append(u"returncode: %s" % result.get(u"returncode"))
        out = clip_pre_shell_text(result.get(u"stdout"))
        err_t = clip_pre_shell_text(result.get(u"stderr"))
        if out:
            parts.append(u"stdout:\n%s" % out)
        if err_t:
            parts.append(u"stderr:\n%s" % err_t)
    return u"\n".join(parts)


def run_pre_shell_from_param_raw(raw):
    """
    Высокоуровневый запуск из текста ячейки.

    Возвращает:
      (None, None) — параметр пуст, пропускаем;
      (None, result) — успех (returncode==0);
      (err_text, result_or_None) — ошибка, сбор прервать.
    """
    spec = parse_pre_shell_dict(raw)
    if spec is None:
        return (None, None)
    try:
        result = run_pre_shell(spec)
    except Exception as err:
        return (format_pre_shell_failure(u"Сбой предварительного скрипта", err=err), None)
    if int(result.get(u"returncode") or 0) != 0:
        return (
            format_pre_shell_failure(
                u"Предварительный скрипт завершился с ошибкой",
                result=result,
            ),
            result,
        )
    return (None, result)
