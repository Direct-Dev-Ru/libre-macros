# -*- coding: utf-8 -*-
"""
Финальная отправка по почте: вложение во временный файл + compose UI (без SMTP).

Используется из lm_final_send_mail (libre_macros_final_lib.py).
"""
from __future__ import print_function

MACRO_VERSION = "3.10.701"
import datetime
import os
import re
import subprocess
import sys

try:
    unicode  # type: ignore[name-defined]
except NameError:  # pragma: no cover
    unicode = str

import uno
from com.sun.star.beans import PropertyValue

import libre_macros_lib as lm

_ATTACH_ALIASES = {
    u"whole_book": u"whole_book",
    u"книга": u"whole_book",
    u"вся_книга": u"whole_book",
    u"book": u"whole_book",
    u"sheets": u"sheets",
    u"листы": u"sheets",
    u"sheet": u"sheets",
    u"none": u"none",
    u"нет": u"none",
    u"без": u"none",
    u"без_вложения": u"none",
    u"self": u"self",
    u"current": u"self",
    u"book_file": u"self",
    u"сама_книга": u"self",
    u"текущая_книга": u"self",
    u"файл_книги": u"self",
}

_FORMAT_ALIASES = {
    u"ods": u"ods",
    u"xlsx": u"xlsx",
    u"xls": u"xlsx",
    u"excel": u"xlsx",
}

_CLIENT_ALIASES = {
    u"auto": u"auto",
    u"system": u"system",
    u"системный": u"system",
    u"uno": u"system",
    u"outlook": u"outlook",
    u"thunderbird": u"thunderbird",
    u"mapi": u"mapi",
}

_OS_ALIASES = {
    u"auto": u"auto",
    u"windows": u"windows",
    u"win": u"windows",
    u"linux": u"linux",
    u"macos": u"macos",
    u"mac": u"macos",
    u"darwin": u"macos",
}

_EMAIL_RE = re.compile(
    r"^[^@\s]+@[^@\s]+\.[^@\s]+$", re.UNICODE
)
_FIO_RE = re.compile(
    r"^[А-Яа-яЁёA-Za-z][А-Яа-яЁёA-Za-z\-]*"
    r"(?:\s+[А-Яа-яЁёA-Za-z][А-Яа-яЁёA-Za-z\-]*){1,3}$",
    re.UNICODE,
)

_TEMP_PREFIX = u"libre_macros_mail_"
# Пути вложений намеренно не удаляем — их читает почтовый клиент.
_MAIL_KEPT_PATHS = []
# Документы, удерживающие lock на файле вложения (закрываем перед compose).
_MAIL_HOLD_DOCS = []
_ATTACH_PREPARE_TIMEOUT_S = 10.0
_ATTACH_STABLE_STEP_S = 0.1
_ATTACH_STABLE_CHECKS = 3
_ATTACH_READY_DELAY_S = 0.5


def _u(text):
    if text is None:
        return u""
    if isinstance(text, unicode):
        return text
    try:
        return unicode(text)
    except Exception:
        return unicode(str(text))


def _err_text(err):
    try:
        return _u(err)
    except Exception:
        return u"ошибка"


def _make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _file_stat_note(path):
    p = os.path.abspath(_u(path or u""))
    if p == u"":
        return u"путь пуст"
    try:
        if os.path.isfile(p):
            return u"есть, %d байт" % int(os.path.getsize(p))
        if os.path.isdir(os.path.dirname(p) or u"."):
            return u"нет файла (каталог есть)"
        return u"нет файла и каталога"
    except Exception as err:
        return u"stat: %s" % _err_text(err)


def _wait_attachment_file(path, timeout_s=5.0, step_s=0.05):
    import time

    deadline = time.time() + float(timeout_s)
    note = u""
    while True:
        note = _file_stat_note(path)
        if note.startswith(u"есть,"):
            return True, note
        if time.time() >= deadline:
            return False, note
        time.sleep(step_s)


def _verify_attachment_readable(path):
    p = os.path.abspath(_u(path or u""))
    if p == u"":
        return False, u"путь пуст"
    if not os.path.isfile(p):
        return False, u"нет файла"
    try:
        sz = int(os.path.getsize(p))
    except Exception as err:
        return False, u"stat: %s" % _err_text(err)
    if sz <= 0:
        return False, u"файл пуст"
    try:
        f = open(p, "rb")
        try:
            chunk = f.read(min(4096, sz))
        finally:
            f.close()
        if len(chunk) <= 0:
            return False, u"не читается"
        return True, u"читается, %d байт" % sz
    except Exception as err:
        return False, u"чтение: %s" % _err_text(err)


def _wait_attachment_stable(path, timeout_s=None, step_s=None, stable_checks=None):
    import time

    if timeout_s is None:
        timeout_s = _ATTACH_PREPARE_TIMEOUT_S
    if step_s is None:
        step_s = _ATTACH_STABLE_STEP_S
    if stable_checks is None:
        stable_checks = _ATTACH_STABLE_CHECKS
    p = os.path.abspath(_u(path or u""))
    deadline = time.time() + float(timeout_s)
    stable = 0
    last_size = -1
    while time.time() < deadline:
        if not os.path.isfile(p):
            stable = 0
            last_size = -1
            time.sleep(step_s)
            continue
        try:
            sz = int(os.path.getsize(p))
        except Exception:
            stable = 0
            last_size = -1
            time.sleep(step_s)
            continue
        if sz > 0 and sz == last_size:
            stable = stable + 1
            if stable >= int(stable_checks):
                ok, rnote = _verify_attachment_readable(p)
                if ok:
                    return True, u"стабилен, %d байт; %s" % (sz, rnote)
                stable = 0
        else:
            stable = 0
        last_size = sz
        time.sleep(step_s)
    return False, _file_stat_note(p)


def _prepare_attachment_for_client(path):
    """
    Закрыть hold-документы LO, дождаться стабильного файла на диске, пауза.
    """
    import time

    _mail_release_holds()
    ready, stat = _wait_attachment_stable(path)
    if not ready:
        return False, stat
    time.sleep(_ATTACH_READY_DELAY_S)
    ok, rnote = _verify_attachment_readable(path)
    if not ok:
        return False, rnote
    return True, stat + u"; пауза %.1fс; %s" % (_ATTACH_READY_DELAY_S, rnote)


def _mail_keep_attachment_path(path):
    p = os.path.abspath(_u(path or u""))
    if p == u"":
        return
    if p not in _MAIL_KEPT_PATHS:
        _MAIL_KEPT_PATHS.append(p)


def _mail_release_holds():
    while _MAIL_HOLD_DOCS:
        doc = _MAIL_HOLD_DOCS.pop()
        _close_doc_quiet(doc)


def _get_desktop():
    ctx = uno.getComponentContext()
    sm = ctx.getServiceManager()
    return sm.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def detect_os():
    if sys.platform.startswith("win"):
        return u"windows"
    if sys.platform == "darwin":
        return u"macos"
    return u"linux"


def _norm_token_map(raw, aliases, default, allowed):
    key = _u(raw or u"").strip().casefold()
    if key == u"":
        key = _u(default).casefold()
    code = aliases.get(key, key)
    if code not in allowed:
        code = default
    return code


def normalize_attach_mode(raw):
    return _norm_token_map(
        raw,
        _ATTACH_ALIASES,
        u"whole_book",
        (u"whole_book", u"sheets", u"none", u"self"),
    )


def normalize_format(raw):
    return _norm_token_map(
        raw,
        _FORMAT_ALIASES,
        u"ods",
        (u"ods", u"xlsx"),
    )


def normalize_client(raw):
    return _norm_token_map(
        raw,
        _CLIENT_ALIASES,
        u"auto",
        (u"auto", u"system", u"outlook", u"thunderbird", u"mapi"),
    )


def normalize_os_param(raw):
    return _norm_token_map(
        raw,
        _OS_ALIASES,
        u"auto",
        (u"auto", u"windows", u"linux", u"macos"),
    )


def _normalize_string_list(raw):
    if raw is None:
        return []
    if isinstance(raw, (list, tuple)):
        parts = [_u(x).strip() for x in raw]
    else:
        text = _u(raw).strip()
        if text == u"":
            return []
        parts = []
        for chunk in re.split(r"[,;]", text):
            st = _u(chunk).strip()
            if st != u"":
                parts.append(st)
    out = []
    for p in parts:
        if p != u"":
            out.append(p)
    return out


def _collapse_spaces(text):
    return re.sub(r"\s+", u" ", _u(text).strip())


def _recipient_key(item):
    s = _collapse_spaces(item)
    if u"@" in s:
        return (u"email", s.casefold())
    return (u"fio", s.casefold())


def _is_valid_email(token):
    return bool(_EMAIL_RE.match(_u(token).strip()))


def _is_valid_fio(token):
    return bool(_FIO_RE.match(_collapse_spaces(token)))


def _dedup_recipients(items):
    out = []
    seen = set()
    for item in items or []:
        s = _u(item).strip()
        if s == u"":
            continue
        key = _recipient_key(s)
        if key in seen:
            continue
        seen.add(key)
        out.append(_collapse_spaces(s) if key[0] == u"fio" else s)
    return out


def _validate_to_source_pair(block):
    sheet = _u(block.get("to_source_sheet") or u"").strip()
    rng = _u(block.get("to_source_range") or u"").strip()
    if (sheet == u"") ^ (rng == u""):
        if sheet == u"":
            return u"to_source_range=%s: нужен to_source_sheet" % rng
        return u"to_source_sheet=%s: нужен to_source_range" % sheet
    return None


def normalize_address_list(raw):
  """Нормализовать to/cc/bcc: list[str] из list или строки с , / ;."""
  return _normalize_string_list(raw)


def collect_to_recipients(doc, block):
    """
    Собрать список получателей to из JSON и опционального диапазона.

    Возвращает (to_list, meta, err).
    meta: json_n, range_n, rejected, rejected_samples, source_cells.
    """
    meta = {
        u"json_n": 0,
        u"range_n": 0,
        u"rejected": 0,
        u"rejected_samples": [],
        u"source_cells": 0,
    }
    pair_err = _validate_to_source_pair(block)
    if pair_err:
        return [], meta, pair_err

    base = _normalize_string_list(block.get("to"))
    meta[u"json_n"] = len(base)
    extra = []
    sheet_name = _u(block.get("to_source_sheet") or u"").strip()
    range_txt = _u(block.get("to_source_range") or u"").strip()
    if sheet_name != u"" and range_txt != u"":
        sheet = lm._lm_pp_get_sheet_ref(doc, sheet_name)
        if sheet is None:
            return [], meta, u"to_source_sheet=%s: лист не найден" % sheet_name
        try:
            from libre_macros_copy_ranges_lib import parse_source_range_a1
        except ImportError:
            return [], meta, u"copy_ranges_lib недоступен"
        rect, perr = parse_source_range_a1(sheet, range_txt)
        if rect is None:
            return [], meta, u"to_source_range=%s: %s" % (range_txt, perr)
        sc, sr, ec, er = rect
        r = int(sr)
        while r <= int(er):
            c = int(sc)
            while c <= int(ec):
                try:
                    cell = sheet.getCellByPosition(c, r)
                    val = lm.cell_text(cell).strip()
                except Exception:
                    val = u""
                if val != u"":
                    meta[u"source_cells"] = meta[u"source_cells"] + 1
                    for tok in re.split(r"[,;]", val):
                        t = _u(tok).strip()
                        if t == u"":
                            continue
                        if _is_valid_email(t) or _is_valid_fio(t):
                            extra.append(t)
                        else:
                            meta[u"rejected"] = meta[u"rejected"] + 1
                            if len(meta[u"rejected_samples"]) < 3:
                                meta[u"rejected_samples"].append(t)
                c = c + 1
            r = r + 1
        meta[u"range_n"] = len(extra)

    combined = _dedup_recipients(base + extra)
    return combined, meta, None


def resolve_client(block, real_os):
    """
    Выбрать стратегию клиента.

    Возвращает (client_code, warnings).
    """
    warnings = []
    client = normalize_client(block.get("client"))
    os_param = normalize_os_param(block.get("os"))
    if os_param != u"auto" and os_param != real_os:
        warnings.append(
            u"os=%s ≠ реальная %s — выполнение по %s"
            % (os_param, real_os, real_os)
        )
    if client == u"outlook" and real_os != u"windows":
        return None, warnings + [u"Outlook только Windows"]
    if client == u"auto":
        if real_os == u"windows":
            client = u"system"
        else:
            client = u"system"
    if client == u"mapi":
        client = u"system"
    return client, warnings


def _book_filename_stem(doc):
    """Имя открытой книги без расширения; пусто если не сохранена."""
    try:
        if doc is not None and getattr(doc, "URL", None) is not None:
            url = _u(doc.URL or u"")
            if url != u"" and not url.startswith(u"private:"):
                base = os.path.splitext(os.path.basename(uno.fileUrlToSystemPath(url)))[0]
                base = _u(base).strip()
                if base != u"":
                    return base
    except Exception:
        pass
    try:
        if doc is not None and getattr(doc, "getTitle", None) is not None:
            title = _u(doc.getTitle() or u"").strip()
            if title != u"":
                return os.path.splitext(title)[0] or title
    except Exception:
        pass
    return u""


def _strip_known_attachment_ext(name):
    stem = _u(name or u"").strip()
    low = stem.casefold()
    for ext in (u".ods", u".xlsx", u".xls", u".fods", u".ots"):
        if low.endswith(ext):
            return stem[: -len(ext)]
    return stem


def _expand_filename_template(template, doc):
    """
    Базовое имя вложения (без расширения и без суффикса даты).

    Если filename задан — использовать как есть (без плейсхолдеров).
    Если пусто — имя исходной книги; иначе «Результат».
    """
    text = _strip_known_attachment_ext(template)
    if text == u"":
        text = _book_filename_stem(doc)
    if text == u"":
        text = u"Результат"
    return lm._lm_normalize_path_component(text, kind="path")


def _store_filter_for_format(fmt):
    if fmt == u"xlsx":
        return u"Calc MS Excel 2007 XML"
    return u"calc8"


def _extension_for_format(fmt):
    return u".xlsx" if fmt == u"xlsx" else u".ods"


def _profile_mail_temp_dir():
    """
    Каталог временных вложений в профиле пользователя: …/alter-macros/mail.

    Не /tmp и не ~/.cache — snap-Thunderbird их не читает.
    """
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        if not base:
            base = os.path.join(os.path.expanduser(u"~"), u"AppData", u"Local")
        path = os.path.join(base, u"alter-macros", u"mail")
    elif sys.platform == "darwin":
        base = os.path.join(
            os.path.expanduser(u"~"),
            u"Library",
            u"Application Support",
        )
        path = os.path.join(base, u"alter-macros", u"mail")
    else:
        path = os.path.join(os.path.expanduser(u"~"), u"alter-macros", u"mail")
    try:
        if not os.path.isdir(path):
            os.makedirs(path)
    except Exception:
        pass
    return path


def _temp_attachment_path(fmt, filename_stem, doc=None):
    """
    Путь вложения: {имя}_{YYYYMMDD_HHMMSS}.{ext}

    Имя — из параметра filename или имени книги; префикс libre_macros_mail_ не добавляется.
    """
    _unused = doc
    stamp = datetime.datetime.now().strftime(u"%Y%m%d_%H%M%S")
    stem = _u(filename_stem or u"").strip() or u"Результат"
    stem = _strip_known_attachment_ext(stem)
    stem = lm._lm_normalize_path_component(stem, kind="path")
    if stem == u"" or stem == u"_":
        stem = u"Результат"
    ext = _extension_for_format(fmt)
    name = u"%s_%s%s" % (stem, stamp, ext)
    return os.path.join(_profile_mail_temp_dir(), name)


def _store_doc_to_path(doc, path, fmt):
    abs_path = os.path.abspath(_u(path))
    url = uno.systemPathToFileUrl(abs_path)
    props = [_make_prop("Overwrite", True)]
    filt = _store_filter_for_format(fmt)
    if filt:
        props.append(_make_prop("FilterName", filt))
    props = tuple(props)
    doc.storeToURL(url, props)
    ready, stat = _wait_attachment_file(abs_path)
    if not ready:
        raise Exception(u"файл вложения не появился: %s" % stat)
    _mail_keep_attachment_path(abs_path)
    return abs_path, url


def _is_default_sheet_name(name):
    n = _u(name or u"").strip()
    if n == u"":
        return True
    try:
        import libre_macros_values_lib as vl

        if vl.lm_values_is_lo_default_sheet_name(n):
            return True
    except Exception:
        pass
    return bool(re.match(r"^(Sheet|Лист)\d+$", n, re.IGNORECASE))


def _import_sheet_to_doc(src_doc, dst_doc, src_name):
    try:
        t = uno.getTypeByName("com.sun.star.sheet.XSpreadsheets2")
        sheets2 = dst_doc.Sheets.queryInterface(t)
        if sheets2 is None:
            return False, u"XSpreadsheets2 недоступен"
        insert_index = int(dst_doc.Sheets.getCount())
        sheets2.importSheet(src_doc, _u(src_name), insert_index)
        return True, u""
    except Exception as err:
        return False, _err_text(err)


def _close_doc_quiet(doc):
    if doc is None:
        return
    try:
        if getattr(doc, "setModified", None) is not None:
            doc.setModified(False)
    except Exception:
        pass
    try:
        doc.close(True)
    except Exception:
        try:
            doc.dispose()
        except Exception:
            pass


def _build_attachment_sheets(doc, block, fmt, filename_stem):
    sheets_spec = block.get("sheets") or []
    names = lm._lm_pp_resolve_merge_source_sheet_names(doc, sheets_spec)
    if not names:
        missing = _normalize_string_list(sheets_spec)
        if missing:
            return None, None, u"", u"листы не найдены: %s" % u", ".join(missing)
        return None, None, u"", u"attach=sheets: пустой список листов"
    path = _temp_attachment_path(fmt, filename_stem, doc)
    desktop = _get_desktop()
    temp_doc = desktop.loadComponentFromURL(
        "private:factory/scalc", "_blank", 0, (_make_prop("Hidden", True),)
    )
    if temp_doc is None:
        return None, None, u"", u"не удалось создать временную книгу"
    try:
        imported = []
        ni = 0
        while ni < len(names):
            ok, note = _import_sheet_to_doc(doc, temp_doc, names[ni])
            if ok:
                imported.append(names[ni])
            else:
                return None, None, u"", u"importSheet %s: %s" % (names[ni], note)
            ni = ni + 1
        if not imported:
            return None, None, u"", u"ни один лист не импортирован"
        try:
            to_remove = []
            i = 0
            while i < int(temp_doc.Sheets.getCount()):
                sh = temp_doc.Sheets.getByIndex(i)
                nm = _u(getattr(sh, "Name", u""))
                if _is_default_sheet_name(nm) and nm not in imported:
                    to_remove.append(nm)
                i = i + 1
            for nm in to_remove:
                try:
                    temp_doc.Sheets.removeByName(nm)
                except Exception:
                    pass
        except Exception:
            pass
        abs_path, url = _store_doc_to_path(temp_doc, path, fmt)
        note = u"sheets=%d; dir=%s; file=%s; %s" % (
            len(imported),
            _profile_mail_temp_dir(),
            os.path.basename(abs_path),
            _file_stat_note(abs_path),
        )
        return abs_path, url, note, None
    except Exception as err:
        return None, None, u"", u"листы: %s" % _err_text(err)
    finally:
        _close_doc_quiet(temp_doc)


def _doc_saved_path(doc):
    url = _u(getattr(doc, "URL", None) or u"").strip()
    if url == u"" or url.startswith(u"private:"):
        return None, u"книга не сохранена на диск — сохраните файл (Ctrl+S)"
    try:
        path = os.path.abspath(uno.fileUrlToSystemPath(url))
    except Exception as err:
        return None, u"URL книги: %s" % _err_text(err)
    if not os.path.isfile(path):
        return None, u"файл книги не найден: %s" % path
    return path, None


def _build_attachment_self(doc):
    try:
        modified = False
        try:
            modified = bool(doc.isModified())
        except Exception:
            modified = False
        if modified:
            doc.store()
    except Exception as err:
        return None, None, u"", u"store: %s" % _err_text(err)
    path, perr = _doc_saved_path(doc)
    if perr:
        return None, None, u"", perr
    ready, stat = _wait_attachment_stable(path)
    if not ready:
        return None, None, u"", u"вложение self: %s" % stat
    url = uno.systemPathToFileUrl(path)
    _mail_keep_attachment_path(path)
    note = u"attach=self; file=%s; %s" % (os.path.basename(path), stat)
    return path, url, note, None


def build_attachment(doc, block):
    """
    Подготовить файл вложения.

    Возвращает (path, url, note, err).
    """
    attach = normalize_attach_mode(block.get("attach"))
    fmt = normalize_format(block.get("format"))
    filename = _expand_filename_template(block.get("filename"), doc)
    if attach == u"none":
        return None, None, u"attach=none", None
    if attach == u"self":
        return _build_attachment_self(doc)
    if attach == u"whole_book":
        path = _temp_attachment_path(fmt, filename, doc)
        try:
            abs_path, url = _store_doc_to_path(doc, path, fmt)
            note = u"attach=whole_book; dir=%s; file=%s; %s" % (
                _profile_mail_temp_dir(),
                os.path.basename(abs_path),
                _file_stat_note(abs_path),
            )
            return abs_path, url, note, None
        except Exception as err:
            return None, None, u"", _err_text(err)
    return _build_attachment_sheets(doc, block, fmt, filename)


def _join_recipients(items):
    if not items:
        return u""
    return u";".join(_u(x) for x in items if _u(x).strip() != u"")


def _uno_mail_service_name(real_os, client):
    if real_os == u"windows":
        return u"com.sun.star.system.SimpleSystemMail"
    return u"com.sun.star.system.SimpleCommandMail"


def _open_compose_uno(real_os, to, cc, bcc, subject, body, attachment_url):
    ctx = uno.getComponentContext()
    sm = ctx.getServiceManager()
    svc = _uno_mail_service_name(real_os, None)
    try:
        mail = sm.createInstanceWithContext(svc, ctx)
    except Exception as err:
        return False, u"%s недоступен: %s" % (svc, _err_text(err))
    if mail is None:
        return False, u"%s недоступен" % svc
    try:
        msg = mail.createSimpleMailMessage()
        if msg is None:
            return False, u"createSimpleMailMessage failed"
        to_s = _join_recipients(to)
        if to_s != u"":
            msg.setRecipient(to_s)
        cc_s = _join_recipients(cc)
        if cc_s != u"":
            msg.setCopyRecipient(cc_s)
        bcc_s = _join_recipients(bcc)
        if bcc_s != u"":
            msg.setBlindCopyRecipient(bcc_s)
        msg.setSubject(_u(subject))
        msg.setBody(_u(body))
        if attachment_url:
            msg.setAttachement((attachment_url,))
        mail.sendSimpleMailMessage(msg, 0)
        return True, svc
    except Exception as err:
        return False, u"%s: %s" % (svc, _err_text(err))


def _shell_quote_arg(text):
    s = _u(text)
    if sys.platform.startswith("win"):
        if u'"' in s:
            return u'"%s"' % s.replace(u'"', u'\\"')
        return u'"%s"' % s
    return u"'" + s.replace(u"'", u"'\\''") + u"'"


def _outlook_mailto_arg(to, subject, body):
    try:
        from urllib import quote as urlquote
    except ImportError:
        from urllib.parse import quote as urlquote

    to_s = _join_recipients(to)
    mailto = to_s
    if subject:
        mailto = mailto + u"&subject=" + urlquote(_u(subject))
    if body:
        mailto = mailto + u"&body=" + urlquote(_u(body))
    return mailto


def _find_outlook_exe():
    if not sys.platform.startswith("win"):
        return None
    try:
        import winreg

        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            key = None
            try:
                key = winreg.OpenKey(
                    root,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
                )
                val, _typ = winreg.QueryValueEx(key, None)
                val = _u(val or u"").strip()
                if val != u"" and os.path.isfile(val):
                    return val
            except Exception:
                pass
            finally:
                if key is not None:
                    try:
                        winreg.CloseKey(key)
                    except Exception:
                        pass
    except ImportError:
        pass
    roots = []
    for env in (u"ProgramFiles", u"ProgramFiles(x86)"):
        base = os.environ.get(env)
        if base:
            roots.append(base)
    rels = (
        r"Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"Microsoft Office\Office16\OUTLOOK.EXE",
        r"Microsoft Office\root\Office15\OUTLOOK.EXE",
        r"Microsoft Office\Office15\OUTLOOK.EXE",
        r"Microsoft Office\root\Office14\OUTLOOK.EXE",
        r"Microsoft Office\Office14\OUTLOOK.EXE",
    )
    for base in roots:
        for rel in rels:
            path = os.path.join(base, rel)
            if os.path.isfile(path):
                return path
    return None


def _open_compose_outlook_shell(to, subject, body, attachment_path):
    outlook = _find_outlook_exe()
    if outlook is None:
        return False, u"OUTLOOK.EXE не найден (Office / App Paths)"
    mailto = _outlook_mailto_arg(to, subject, body)
    args = [outlook]
    if attachment_path:
        args.extend([u"/a", _u(attachment_path)])
    if mailto:
        args.extend([u"/m", mailto])
    try:
        subprocess.Popen(args)
        return True, u"outlook: %s" % outlook
    except Exception as err:
        return False, u"outlook %s: %s" % (outlook, _err_text(err))


def _which_executable(name):
    path_env = _u(os.environ.get("PATH", u""))
    for part in path_env.split(os.pathsep):
        part = part.strip()
        if part == u"":
            continue
        candidate = os.path.join(part, name)
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate
    return None


def _find_thunderbird_compose_bin():
    """
    Клиент compose: AlterOffice amail или Thunderbird.

    amail — обёртка TB в AlterOffice (/usr/bin/alteroffice-amail).
    """
    candidates = (
        u"/usr/bin/alteroffice-amail",
        u"/usr/local/bin/alteroffice-amail",
        u"/snap/bin/thunderbird",
        u"/usr/bin/thunderbird",
    )
    ci = 0
    while ci < len(candidates):
        path = candidates[ci]
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return path
        ci = ci + 1
    found = _which_executable(u"alteroffice-amail")
    if found:
        return found
    found = _which_executable(u"thunderbird")
    if found:
        return found
    return None


def _thunderbird_attachment_token(path):
    p = os.path.abspath(_u(path or u""))
    if p == u"":
        return u""
    # Thunderbird -compose: attachment — file:// URL (не голый путь).
    return uno.systemPathToFileUrl(p)


def _open_compose_thunderbird_shell(to, cc, subject, body, attachment_path, prep_note=u""):
    parts = []
    to_s = _join_recipients(to)
    if to_s:
        parts.append(u"to='%s'" % to_s.replace(u"'", u"\\'"))
    cc_s = _join_recipients(cc)
    if cc_s:
        parts.append(u"cc='%s'" % cc_s.replace(u"'", u"\\'"))
    if subject:
        parts.append(u"subject='%s'" % _u(subject).replace(u"'", u"\\'"))
    if body:
        parts.append(u"body='%s'" % _u(body).replace(u"'", u"\\'"))
    attach_tok = u""
    if attachment_path:
        attach_tok = _thunderbird_attachment_token(attachment_path)
        if attach_tok != u"":
            parts.append(u"attachment='%s'" % attach_tok.replace(u"'", u"\\'"))
    compose = u",".join(parts)
    mail_bin = _find_thunderbird_compose_bin()
    if mail_bin is None:
        return False, u"alteroffice-amail/thunderbird не найден"
    try:
        subprocess.Popen([mail_bin, u"-compose", compose])
        note = u"compose: %s" % mail_bin
        if attach_tok:
            note = note + u"; attach=" + attach_tok
        if prep_note:
            note = prep_note + u"; " + note
        return True, note
    except Exception as err:
        return False, u"%s: %s" % (mail_bin, _err_text(err))


def open_compose(block, to, cc, bcc, subject, body, attachment_path, attachment_url, client, real_os):
    """
    Открыть окно compose. Возвращает (ok, note).
    """
    client = normalize_client(client)
    path = os.path.abspath(_u(attachment_path or u""))
    has_attach = path != u""
    prep_note = u""

    if has_attach:
        ready, prep_note = _prepare_attachment_for_client(path)
        if not ready:
            return False, u"вложение: %s" % prep_note

    if client == u"outlook":
        return _open_compose_outlook_shell(to, subject, body, path or None)

    if client == u"thunderbird":
        return _open_compose_thunderbird_shell(
            to, cc, subject, body, path or None, prep_note=prep_note
        )

    # auto / system / mapi
    if has_attach and real_os in (u"linux", u"macos"):
        # SimpleCommandMail на Linux/macOS забирает вложение до открытия клиента.
        return _open_compose_thunderbird_shell(
            to, cc, subject, body, path, prep_note=prep_note
        )

    if has_attach and real_os == u"windows":
        ok, note = _open_compose_uno(
            real_os, to, cc, bcc, subject, body, attachment_url
        )
        if ok:
            return True, note
        ok2, note2 = _open_compose_outlook_shell(to, subject, body, path)
        if ok2:
            return True, note + u"; " + note2
        return False, note + u"; " + note2

    ok, note = _open_compose_uno(
        real_os,
        to,
        cc,
        bcc,
        subject,
        body,
        attachment_url if has_attach else None,
    )
    if ok:
        return True, note
    if real_os == u"windows":
        ok2, note2 = _open_compose_outlook_shell(to, subject, body, path or None)
        if ok2:
            return True, note + u"; " + note2
        return False, note + u"; " + note2
    ok2, note2 = _open_compose_thunderbird_shell(
        to, cc, subject, body, path or None, prep_note=prep_note
    )
    if ok2:
        return True, note + u"; " + note2
    return False, note + u"; " + note2


def format_to_log_note(to_list, meta, attach_note, client_note, real_os, client):
    parts = [u"%s/%s" % (real_os, client)]
    if attach_note:
        parts.append(attach_note)
    to_n = len(to_list or [])
    json_n = int(meta.get(u"json_n") or 0)
    range_n = int(meta.get(u"range_n") or 0)
    rejected = int(meta.get(u"rejected") or 0)
    if range_n > 0 or rejected > 0:
        parts.append(
            u"to=%d (json=%d, range=%d, rejected=%d)"
            % (to_n, json_n, range_n, rejected)
        )
    else:
        parts.append(u"to=%d" % to_n)
    if client_note:
        parts.append(client_note)
    if to_n == 0 and int(meta.get(u"source_cells") or 0) > 0:
        parts.append(u"to_source: 0 valid")
    return u"; ".join(parts)
