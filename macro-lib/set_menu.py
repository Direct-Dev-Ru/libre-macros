# -*- coding: utf-8 -*-
"""
Наименование: set_menu

Описание: простая установка встроенных menubar.xml, cell.xml
и кастомного тулбара в профили AlterOffice (Linux / Windows).

Точка входа: set_menu()
"""
MACRO_VERSION = "3.10.706"
import os
import sys
import tempfile
import zipfile
import uno


def _set_menu_add_path(path):
    if path and os.path.isdir(path) and path not in sys.path:
        sys.path.insert(0, path)
        return True
    return False


def _set_menu_add_user_pythonpath():
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        path_sub = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.util.PathSubstitution", ctx)
        user_url = path_sub.getSubstituteVariableValue("user")
        user_dir = uno.fileUrlToSystemPath(user_url)
        return _set_menu_add_path(
            os.path.join(user_dir, "Scripts", "python", "pythonpath")
        )
    except Exception:
        return False


def _set_menu_script_dir():
    src = ""
    try:
        src = str(__file__ or "")
    except Exception:
        src = ""
    if not src:
        return ""
    norm = src.replace("\\", "/")
    idx = norm.rfind("/")
    if idx <= 0:
        return ""
    return src[:idx]


def _set_menu_add_file_pythonpath(script_dir):
    if not script_dir:
        return False
    try:
        raw = str(script_dir)
        if raw.startswith("vnd.sun.star.tdoc:"):
            return False
        if raw.startswith("file:"):
            local = uno.fileUrlToSystemPath(script_dir)
        else:
            local = script_dir
        added = _set_menu_add_path(os.path.join(local, "pythonpath"))
        added = _set_menu_add_path(local) or added
        return added
    except Exception:
        return False


def _set_menu_copy_uri_to_dir(sfa, uri, dest_dir, name):
    local = os.path.join(dest_dir, name)
    sfa.copy(uri, uno.systemPathToFileUrl(local))
    return os.path.isfile(local)


def _set_menu_add_tdoc_pythonpath(script_dir):
    """AO грузит документный макрос с tdoc: — pythonpath рядом не в sys.path."""
    names = (
        "libre_macros_set_menu_lib.py",
        "libre_macros_set_menu_cfg.py",
        "libre_macros_set_menu_images_blob.py",
    )
    if not script_dir:
        return False
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sfa = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.ucb.SimpleFileAccess", ctx)
        src_dirs = (script_dir + "/pythonpath", script_dir)
        found = {}
        for name in names:
            for folder in src_dirs:
                uri = folder + "/" + name
                try:
                    if sfa.exists(uri):
                        found[name] = uri
                        break
                except Exception:
                    continue
        if names[0] not in found or names[1] not in found or names[2] not in found:
            return False
        dest = tempfile.mkdtemp(prefix="lm_set_menu_pp_")
        for name, uri in found.items():
            _set_menu_copy_uri_to_dir(sfa, uri, dest, name)
        return _set_menu_add_path(dest)
    except Exception:
        return False


def _set_menu_add_ods_pythonpath():
    """Запасной путь: lib/cfg/blob лежат внутри MacroInstaller_bundled.ods."""
    names = (
        "libre_macros_set_menu_lib.py",
        "libre_macros_set_menu_cfg.py",
        "libre_macros_set_menu_images_blob.py",
    )
    prefixes = ("Scripts/python/pythonpath/", "Scripts/python/")
    try:
        doc = XSCRIPTCONTEXT.getDocument()
        url = str(doc.getURL() or "")
        if not url.lower().startswith("file:"):
            return False
        ods_path = uno.fileUrlToSystemPath(url)
        if not os.path.isfile(ods_path):
            return False
        payloads = {}
        with zipfile.ZipFile(ods_path, "r") as zf:
            znames = set(zf.namelist())
            for name in names:
                for prefix in prefixes:
                    arc = prefix + name
                    if arc in znames:
                        payloads[name] = zf.read(arc)
                        break
        if names[0] not in payloads or names[1] not in payloads or names[2] not in payloads:
            return False
        dest = tempfile.mkdtemp(prefix="lm_set_menu_pp_")
        for name, payload in payloads.items():
            with open(os.path.join(dest, name), "wb") as fh:
                fh.write(payload)
        return _set_menu_add_path(dest)
    except Exception:
        return False


def _set_menu_prepare_imports():
    _set_menu_add_user_pythonpath()
    script_dir = _set_menu_script_dir()
    _set_menu_add_file_pythonpath(script_dir)
    try:
        from libre_macros_set_menu_lib import set_menu_entry as _entry
        return _entry
    except Exception:
        sys.modules.pop("libre_macros_set_menu_lib", None)
        sys.modules.pop("libre_macros_set_menu_cfg", None)
        sys.modules.pop("libre_macros_set_menu_images_blob", None)
    _set_menu_add_tdoc_pythonpath(script_dir)
    _set_menu_add_ods_pythonpath()
    from libre_macros_set_menu_lib import set_menu_entry as _entry
    return _entry


def set_menu(*args):
    _entry = _set_menu_prepare_imports()
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _entry(doc, *args)


g_exportedScripts = (set_menu,)
