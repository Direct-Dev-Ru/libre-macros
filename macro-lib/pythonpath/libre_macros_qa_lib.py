# -*- coding: utf-8 -*-
"""Headless-прогон collect_workbooks: job JSON → тихий сбор → storeAsURL."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.707"
import json
import os
import sys
import traceback

try:
    unicode
except NameError:
    unicode = str


def _as_text(raw):
    if raw is None:
        return u""
    try:
        return unicode(raw)
    except Exception:
        return str(raw)


def _load_job(args):
    path = u""
    if args:
        first = args[0]
        if first is not None and _as_text(first).strip() not in (u"", u"None"):
            path = _as_text(first).strip()
    if path == u"":
        path = _as_text(os.environ.get("LM_QA_JOB") or u"").strip()
    if path == u"" or not os.path.isfile(path):
        raise ValueError(u"qa_collect_run: нет файла job JSON: %s" % path)
    with open(path, "rb") as fh:
        raw = fh.read().decode("utf-8")
    job = json.loads(raw)
    if not isinstance(job, dict):
        raise ValueError(u"qa_collect_run: job должен быть объектом JSON")
    job["_job_path"] = path
    return job


def _store_filter(path):
    ext = os.path.splitext(_as_text(path))[1].lower()
    if ext in (u".xlsx", u".xltx", u".xlsm"):
        return u"Calc MS Excel 2007 XML"
    if ext in (u".xls", u".xlt"):
        return u"MS Excel 97"
    return u"calc8"


def _file_url(path):
    import uno

    path = os.path.abspath(_as_text(path))
    return uno.systemPathToFileUrl(path)


def _make_prop(name, value):
    from com.sun.star.beans import PropertyValue

    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _doc_needs_disk_save(doc):
    """Untitled / шаблон .ots — xml_excel_ods требует .ods/.xlsx на диске."""
    try:
        url = _as_text(doc.getURL() if hasattr(doc, "getURL") else getattr(doc, "URL", u""))
    except Exception:
        url = u""
    low = url.strip().lower()
    if low == u"" or low.startswith(u"private:"):
        return True
    for ext in (u".ots", u".xltx", u".xltm", u".xlt", u".xlt"):
        if low.endswith(ext):
            return True
    try:
        import uno

        path = uno.fileUrlToSystemPath(url)
        ext = os.path.splitext(path)[1].lower()
        if ext not in (u".ods", u".xlsx", u".xlsm"):
            return True
        if not os.path.isfile(path):
            return True
    except Exception:
        return True
    return False


def _store_doc(doc, save_path):
    import uno

    save_path = os.path.abspath(_as_text(save_path))
    parent = os.path.dirname(save_path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    url = _file_url(save_path)
    props = (
        _make_prop(u"FilterName", _store_filter(save_path)),
        _make_prop(u"Overwrite", True),
    )
    if getattr(doc, "storeAsURL", None) is not None:
        doc.storeAsURL(url, props)
    else:
        doc.storeToURL(url, props)
    return save_path


def _resolve_param_sheet(cw, doc, wanted):
    wanted = _as_text(wanted).strip()
    names = []
    try:
        names = list(cw.list_param_sheet_names(doc, tab_order=True) or [])
    except Exception:
        names = []
    if wanted:
        for n in names:
            if n == wanted:
                return n
        for n in names:
            if _as_text(n).casefold() == wanted.casefold():
                return n
        raise ValueError(u"Лист параметров «%s» не найден. Есть: %s" % (wanted, u", ".join(names)))
    if len(names) == 1:
        return names[0]
    active = u""
    try:
        active = _as_text(cw.get_active_param_sheet_name(doc) or u"").strip()
    except Exception:
        active = u""
    if active:
        return active
    if names:
        return names[0]
    raise ValueError(u"В книге нет листа параметров (collect_params / Параметры_Объединения)")


def _import_collect_workbooks():
    """
    collect_workbooks.py — сосед Scripts/python/, этот модуль — в pythonpath/.
    UNO-import часто не видит Scripts/python в sys.path → грузим по абсолютному пути.
    """
    name = "collect_workbooks"
    if name in sys.modules:
        return sys.modules[name]
    _pp = os.path.dirname(os.path.abspath(__file__))
    _scripts = os.path.dirname(_pp)
    if _scripts and os.path.isdir(_scripts) and _scripts not in sys.path:
        sys.path.insert(0, _scripts)
    path = os.path.join(_scripts, "collect_workbooks.py")
    if not os.path.isfile(path):
        raise ImportError(u"не найден %s" % path)
    try:
        import importlib.util

        spec = importlib.util.spec_from_file_location(name, path)
        if spec is None or spec.loader is None:
            raise ImportError(u"spec_from_file_location failed: %s" % path)
        mod = importlib.util.module_from_spec(spec)
        sys.modules[name] = mod
        spec.loader.exec_module(mod)
        return mod
    except Exception:
        import collect_workbooks as cw

        return cw


def _desktop_spreadsheet():
    """Запасной поиск открытой книги Calc, если XSCRIPTCONTEXT без документа."""
    try:
        import uno

        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        cur = desktop.getCurrentComponent()
        if cur is not None and hasattr(cur, "getSheets"):
            return cur
        comps = desktop.getComponents()
        if comps is None:
            return None
        enum = comps.createEnumeration()
        while enum.hasMoreElements():
            c = enum.nextElement()
            if c is not None and hasattr(c, "getSheets"):
                return c
    except Exception:
        return None
    return None


def _inject_xscriptcontext(cw, xsc):
    """collect_workbooks при importlib не получает XSCRIPTCONTEXT от pythonscript."""
    if cw is None or xsc is None:
        return
    try:
        cw.XSCRIPTCONTEXT = xsc
    except Exception:
        pass
    try:
        import libre_macros_collect_cfg as _cfg

        _cfg.XSCRIPTCONTEXT = xsc
    except Exception:
        pass


def _diag_log_rows(cw, doc, max_rows=12):
    rows = []
    try:
        name = getattr(cw, "_cw_cfg", None)
        log_name = u"Сбор_книг_лог"
        if name is not None:
            log_name = getattr(cw._cw_cfg, "MERGE_LOG_SHEET_NAME", log_name) or log_name
        sh = cw.get_sheet(doc, log_name)
        if sh is None:
            return rows
        r = 0
        while r < max_rows:
            line = []
            c = 0
            while c < 7:
                try:
                    line.append(cw.cell_text(sh.getCellByPosition(c, r)))
                except Exception:
                    line.append(u"")
                c += 1
            if any(_as_text(x).strip() for x in line):
                rows.append(line)
            r += 1
    except Exception as err:
        rows.append([u"diag_log_error", _as_text(err)])
    return rows


def entry(doc, xsc=None, *args):
    """Точка входа макроса qa_collect_run: doc/XSCRIPTCONTEXT + job.json."""
    # Совместимость: entry(doc, job_path) без xsc
    if xsc is not None and not hasattr(xsc, "getDocument") and (
        isinstance(xsc, (str, bytes)) or _as_text(xsc).endswith(u".json")
    ):
        args = (xsc,) + tuple(args)
        xsc = None

    cw = _import_collect_workbooks()
    _inject_xscriptcontext(cw, xsc)

    job = _load_job(args)
    save_path = _as_text(job.get("save_path") or u"").strip()
    if save_path == u"":
        raise ValueError(u"qa_collect_run: в job нет save_path")
    if doc is None and xsc is not None:
        try:
            doc = xsc.getDocument()
        except Exception:
            doc = None
    if doc is None:
        try:
            doc = cw.get_run_target_doc()
        except Exception:
            doc = None
    if doc is None:
        doc = _desktop_spreadsheet()
    if doc is None:
        raise ValueError(u"qa_collect_run: нет открытой книги")

    cw._cw_set("_MERGE_QA_HEADLESS", True)
    try:
        cw._cw_set("_MERGE_LAST_COLLECT_OK", False)
    except Exception:
        pass
    sheet_name = _resolve_param_sheet(cw, doc, job.get("param_sheet"))
    ok = False
    err_text = u""
    started_ok = False
    parse_err = u""
    try:
        cw.set_run_target_doc(doc)
        if _doc_needs_disk_save(doc):
            _store_doc(doc, save_path)
        try:
            parse_err, _settings = cw.parse_collect_settings(doc, validate_sources=True)
            parse_err = _as_text(parse_err or u"")
        except Exception as exc:
            parse_err = u"parse exception: %s" % _as_text(exc)
        started_ok = cw.collect_workbooks_for_param_sheet(
            sheet_name,
            suppress_pre_clear=bool(job.get("suppress_pre_clear", False)),
            skip_consent=True,
            skip_confirm=True,
            quiet_finish=True,
        )
        ok = bool(cw._cw_get("_MERGE_LAST_COLLECT_OK", False))
        if not ok:
            ok = bool(started_ok)
        result_doc = None
        try:
            result_doc = cw.get_run_target_doc()
        except Exception:
            result_doc = None
        if result_doc is None:
            result_doc = doc
        stored = _store_doc(result_doc, save_path)
        job_out = {
            "ok": bool(ok),
            "save_path": stored,
            "param_sheet": sheet_name,
            "started_ok": bool(started_ok),
            "last_ok": bool(cw._cw_get("_MERGE_LAST_COLLECT_OK", False)),
            "parse_err": parse_err,
            "has_xsc": bool(xsc is not None),
            "log_rows": _diag_log_rows(cw, result_doc),
        }
        marker = save_path + u".qa.json"
        with open(marker, "w", encoding="utf-8") as fh:
            json.dump(job_out, fh, ensure_ascii=False, indent=2)
        return (ok, stored)
    except Exception as exc:
        err_text = _as_text(exc)
        traceback.print_exc()
        marker = save_path + u".qa.json"
        try:
            parent = os.path.dirname(os.path.abspath(save_path))
            if parent and not os.path.isdir(parent):
                os.makedirs(parent)
            with open(marker, "w", encoding="utf-8") as fh:
                json.dump(
                    {
                        "ok": False,
                        "error": err_text,
                        "param_sheet": sheet_name,
                        "parse_err": parse_err,
                        "has_xsc": bool(xsc is not None),
                        "log_rows": _diag_log_rows(cw, doc),
                    },
                    fh,
                    ensure_ascii=False,
                    indent=2,
                )
        except Exception:
            pass
        raise
    finally:
        try:
            cw._cw_set("_MERGE_QA_HEADLESS", False)
        except Exception:
            pass
