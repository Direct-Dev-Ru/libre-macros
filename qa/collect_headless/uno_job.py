# -*- coding: utf-8 -*-
"""Клиент UNO: открыть книгу, вызвать qa_collect_run, выйти.

Запускать интерпретатором офиса (program/python), не системным venv.
"""
from __future__ import print_function, unicode_literals

import json
import os
import sys
import time
import traceback

try:
    unicode
except NameError:
    unicode = str

SCRIPT_URI = (
    "vnd.sun.star.script:qa_collect_run.py$qa_collect_run?language=Python&location=user"
)


def _as_text(raw):
    if raw is None:
        return u""
    try:
        return unicode(raw)
    except Exception:
        return str(raw)


def _load_job(path):
    with open(path, "rb") as fh:
        return json.loads(fh.read().decode("utf-8"))


def connect(host, port, timeout=40):
    import uno
    from com.sun.star.connection import NoConnectException

    url = "uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % (host, port)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local
    )
    deadline = time.time() + float(timeout)
    last = None
    while time.time() < deadline:
        try:
            return resolver.resolve(url)
        except NoConnectException as err:
            last = err
            time.sleep(0.3)
        except Exception as err:
            last = err
            time.sleep(0.3)
    raise RuntimeError("UNO connect failed: %s" % last)


def _prop(name, value):
    from com.sun.star.beans import PropertyValue

    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def file_url(path):
    import uno

    return uno.systemPathToFileUrl(os.path.abspath(path))


def load_doc(desktop, path, as_template=True):
    ext = os.path.splitext(path)[1].lower()
    # Hidden=True иногда ломает запись на листы / текущий документ для макросов
    props = [
        _prop("Hidden", False),
        _prop("ReadOnly", False),
        _prop("MacroExecutionMode", 4),
    ]
    if as_template or ext in (".ots", ".xltx", ".xlt", ".xltm"):
        props.append(_prop("AsTemplate", True))
    url = file_url(path)
    doc = desktop.loadComponentFromURL(url, "_default", 0, tuple(props))
    if doc is None:
        raise RuntimeError("Не удалось открыть %s" % path)
    return doc


def close_other_docs(desktop, keep):
    try:
        enum = desktop.getComponents().createEnumeration()
    except Exception:
        return
    while enum.hasMoreElements():
        part = enum.nextElement()
        if part == keep:
            continue
        try:
            part.close(True)
        except Exception:
            try:
                part.close(False)
            except Exception:
                pass


def invoke_macro(ctx, job_path, doc=None):
    factory = ctx.getValueByName(
        "/singletons/com.sun.star.script.provider.theMasterScriptProviderFactory"
    )
    if factory is None:
        raise RuntimeError("нет MasterScriptProviderFactory")
    # Привязка к документу → XSCRIPTCONTEXT.getDocument() в макросе не None
    provider = factory.createScriptProvider(doc if doc is not None else "")
    script = provider.getScript(SCRIPT_URI)
    out = script.invoke((_as_text(job_path),), (), ())
    return out


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv:
        print("usage: uno_job.py JOB.json", file=sys.stderr)
        return 2
    job_path = os.path.abspath(argv[0])
    job = _load_job(job_path)
    host = _as_text(job.get("host") or "127.0.0.1")
    port = int(job.get("port") or 2002)
    doc_path = _as_text(job.get("doc_path") or u"").strip()
    if not doc_path:
        print("job без doc_path", file=sys.stderr)
        return 2
    ctx = connect(host, port, timeout=int(job.get("connect_timeout") or 40))
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = None
    try:
        doc = load_doc(desktop, doc_path, as_template=bool(job.get("as_template", True)))
        close_other_docs(desktop, doc)
        invoke_macro(ctx, job_path, doc=doc)
        marker = _as_text(job.get("save_path") or u"") + u".qa.json"
        ok = False
        if marker and os.path.isfile(marker):
            try:
                with open(marker, "rb") as fh:
                    data = json.loads(fh.read().decode("utf-8"))
                ok = bool(data.get("ok"))
            except Exception:
                ok = False
        return 0 if ok else 1
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        if doc is not None:
            try:
                doc.close(True)
            except Exception:
                try:
                    doc.close(False)
                except Exception:
                    pass


if __name__ == "__main__":
    sys.exit(main())
