# -*- coding: utf-8 -*-
"""Простая установка menubar.xml, cell.xml, тулбара и иконок команд в профиль AlterOffice."""
from __future__ import print_function

import os
import re
import sys

import uno

from libre_macros_set_menu_cfg import (
    CELL_FILENAME,
    CELL_XML,
    LC_IMAGELIST_FILENAME,
    LC_USERIMAGES_PNG_FILENAME,
    MENUBAR_FILENAME,
    MENUBAR_XML,
    TOOLBAR_FILENAME,
    TOOLBAR_RESOURCE,
    TOOLBAR_UINAME,
    TOOLBAR_XML,
    MACRO_VERSION as CFG_MACRO_VERSION,
)

MACRO_VERSION = "3.10.689"
def _menu_xml_with_version(xml, version=None):
    """Подставить MACRO_VERSION в подписи PythonMacros / Версия (__LM_VER__)."""
    ver = str(version or MACRO_VERSION or CFG_MACRO_VERSION or "").strip()
    if not ver:
        ver = "0"
    try:
        return str(xml).replace("__LM_VER__", ver)
    except Exception:
        return xml


def patch_menu_xml_version_labels(xml_text, version=None):
    """
    Обновить в уже установленном menubar/cell.xml подписи:
      PythonMacros (x.y.z)
      Версия x.y.z
    и при отсутствии — добавить пункт version.py$show_version.
    Возвращает (new_xml, changed_bool).
    """
    ver = str(version or MACRO_VERSION or CFG_MACRO_VERSION or "").strip() or "0"
    xml = str(xml_text or "")
    if xml == "":
        return xml, False
    orig = xml

    def _sub_title(pattern, repl):
        nonlocal xml
        xml2, n = re.subn(pattern, repl, xml, count=1, flags=re.IGNORECASE)
        if n:
            xml = xml2

    _sub_title(
        r'(menu:id="vnd\.openoffice\.org:CustomMenu1"\s+menu:label=")PythonMacros(?:\s*\([^"]*\))?"',
        r'\1PythonMacros (%s)"' % ver,
    )
    _sub_title(
        r'(menu:id="PythonMacros"\s+menu:label=")PythonMacros(?:\s*\([^"]*\))?"',
        r'\1PythonMacros (%s)"' % ver,
    )
    # Уже есть пункт Версия → обновить label
    xml2, n = re.subn(
        r'(version\.py\$show_version\?language=Python&amp;location=user"\s+menu:label=")[^"]*"',
        r'\1Версия %s"' % ver,
        xml,
        flags=re.IGNORECASE,
    )
    if n:
        xml = xml2
    elif "version.py$show_version" not in xml:
        # Добавить separator + пункт перед закрытием menupopup PythonMacros / CustomMenu1
        item = (
            '<menu:menuseparator/>\n'
            '   <menu:menuitem menu:id="vnd.sun.star.script:version.py$show_version'
            '?language=Python&amp;location=user" menu:label="Версия %s"/>'
        ) % ver
        inserted = False
        for pat in (
            r'(menu:id="vnd\.openoffice\.org:CustomMenu1"[^>]*>\s*<menu:menupopup>)(.*?)(</menu:menupopup>)',
            r'(menu:id="PythonMacros"[^>]*>\s*<menu:menupopup>)(.*?)(</menu:menupopup>)',
        ):
            m = re.search(pat, xml, flags=re.IGNORECASE | re.DOTALL)
            if not m:
                continue
            body = m.group(2).rstrip()
            xml = (
                xml[: m.start(2)]
                + body
                + "\n   "
                + item
                + "\n  "
                + xml[m.start(3) :]
            )
            inserted = True
            break
        _ = inserted
    # placeholder из референса
    xml = xml.replace("__LM_VER__", ver)
    return xml, xml != orig


def patch_installed_menu_files(user_dir, version=None):
    """
    Если у пользователя уже есть menubar.xml / cell.xml — обновить подписи версии.
    Возвращает list[str] заметок для отчёта.
    """
    notes = []
    if not user_dir or not os.path.isdir(user_dir):
        return notes
    ver = str(version or MACRO_VERSION or "").strip() or "0"
    candidates = (
        ("menubar.xml", ("config", "aoffice.cfg", "modules", "acell", "menubar", MENUBAR_FILENAME)),
        ("cell.xml", ("config", "aoffice.cfg", "modules", "acell", "popupmenu", CELL_FILENAME)),
        ("menubar.xml", ("config", "soffice.cfg", "modules", "scalc", "menubar", MENUBAR_FILENAME)),
        ("cell.xml", ("config", "soffice.cfg", "modules", "scalc", "popupmenu", CELL_FILENAME)),
    )
    seen = set()
    for kind, parts in candidates:
        path = _join_user(user_dir, parts)
        if path in seen:
            continue
        seen.add(path)
        if not os.path.isfile(path):
            continue
        try:
            with open(path, "r", encoding="utf-8") as fh:
                old = fh.read()
        except Exception as err:
            notes.append("ERR read %s: %s" % (path, err))
            continue
        new, changed = patch_menu_xml_version_labels(old, ver)
        if not changed:
            notes.append("skip %s (без PythonMacros / без изменений)" % path)
            continue
        try:
            _write_text(path, new)
            notes.append("OK  %s → PythonMacros/Версия %s" % (path, ver))
        except Exception as err:
            notes.append("ERR write %s: %s" % (path, err))
    return notes


def _log(msg):
    try:
        print("[set_menu] %s" % msg)
    except Exception:
        pass


def _ao_rel_dirs():
    return (
        ("config", "aoffice.cfg", "modules", "acell", "menubar"),
        ("config", "aoffice.cfg", "modules", "acell", "popupmenu"),
        ("config", "aoffice.cfg", "modules", "acell", "toolbar"),
        ("config", "aoffice.cfg", "modules", "acell", "images"),
    )


def _join_user(user_dir, parts):
    return os.path.abspath(os.path.join(user_dir, *parts))


def _is_alteroffice_user(user_dir):
    if not user_dir:
        return False
    low = str(user_dir).replace("\\", "/").lower()
    if "alteroffice" in low:
        return True
    cfg = os.path.join(user_dir, "config", "aoffice.cfg")
    return os.path.isdir(cfg)


def _profile_version_label(user_dir):
    if not user_dir:
        return "?"
    parts = os.path.normpath(user_dir).split(os.sep)
    for i in range(len(parts) - 1, -1, -1):
        if parts[i] == "user" and i > 0:
            return parts[i - 1]
    m = re.search(r"[\\/](\d+)[\\/]user$", user_dir.replace("\\", "/"))
    if m:
        return m.group(1)
    return "?"


def _user_dir_from_path(path):
    """.../user/Scripts/python[/...] → .../user."""
    if not path:
        return None
    d = os.path.abspath(path)
    if os.path.isfile(d):
        d = os.path.dirname(d)
    # .../Scripts/python/pythonpath → .../Scripts/python
    if os.path.basename(d) == "pythonpath":
        d = os.path.dirname(d)
    # .../Scripts/python → .../Scripts → .../user
    if os.path.basename(d) == "python":
        scripts = os.path.dirname(d)
        if os.path.basename(scripts) == "Scripts":
            user_dir = os.path.dirname(scripts)
            if os.path.basename(user_dir) == "user":
                return user_dir
    return None


def _user_dir_from_macro_lib():
    try:
        from_file = _user_dir_from_path(__file__)
        if from_file is not None:
            return from_file
    except Exception:
        pass
    try:
        for p in sys.path:
            if not p:
                continue
            if os.path.isfile(os.path.join(p, "set_menu.py")):
                found = _user_dir_from_path(p)
                if found is not None:
                    return found
    except Exception:
        pass
    return None


def _get_desktop():
    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    except Exception:
        return None


def _get_toolkit():
    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    except Exception:
        return None


def _resolve_doc(doc=None):
    if doc is not None:
        try:
            if doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                return doc
        except Exception:
            pass
    desktop = _get_desktop()
    if desktop is None:
        return None
    try:
        cur = desktop.getCurrentComponent()
        if cur is not None and cur.supportsService(
            "com.sun.star.sheet.SpreadsheetDocument"
        ):
            return cur
    except Exception:
        pass
    return None


def _dialog_parent(doc):
    if doc is not None:
        try:
            return doc.getCurrentController().getFrame().getContainerWindow()
        except Exception:
            pass
    try:
        desktop = _get_desktop()
        if desktop is not None:
            return desktop.getCurrentFrame().getContainerWindow()
    except Exception:
        pass
    toolkit = _get_toolkit()
    if toolkit is not None:
        try:
            return toolkit.getDesktopWindow()
        except Exception:
            pass
    return None


def _show_message(doc, title, text, is_error=False):
    """Как в convert_tables: простой MessageBox INFOBOX/ERRORBOX."""
    toolkit = _get_toolkit()
    if toolkit is None:
        _log("%s: %s" % (title, text))
        return
    parent = _dialog_parent(doc)
    if parent is None:
        try:
            parent = toolkit.getDesktopWindow()
        except Exception:
            parent = None
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX

    box_type = ERRORBOX if is_error else INFOBOX
    try:
        _log("MessageBox.execute() begin")
        box = toolkit.createMessageBox(parent, box_type, BUTTONS_OK, title, text)
        box.execute()
        _log("MessageBox.execute() done")
    except Exception as err:
        _log("MessageBox failed: %s; %s: %s" % (err, title, text))


def _uno_user_profile_dir():
    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        ps = sm.createInstanceWithContext("com.sun.star.util.PathSettings", ctx)
        url = ps.getPropertyValue("UserConfig")
        if url in (None, ""):
            return None
        from urllib.parse import unquote, urlparse

        parsed = urlparse(str(url))
        if parsed.scheme != "file":
            return None
        path = unquote(parsed.path)
        if os.path.basename(path) == "user":
            return os.path.abspath(path)
    except Exception:
        pass
    return None


def _iter_candidate_user_roots():
    seen = set()

    def add(path):
        if not path:
            return
        ap = os.path.abspath(path)
        if ap in seen:
            return
        seen.add(ap)
        if os.path.isdir(ap):
            yield ap

    for p in add(_user_dir_from_macro_lib()):
        yield p
    uno = _uno_user_profile_dir()
    if uno:
        yield uno

    home = os.path.expanduser("~")
    for ver in ("5", "4", "3"):
        p = os.path.join(home, ".config", "alteroffice", ver, "user")
        if os.path.isdir(p):
            yield os.path.abspath(p)

    appdata = os.environ.get("APPDATA")
    if appdata:
        for office in ("AlterOffice3", "AlterOffice"):
            for ver in ("5", "4", "3"):
                p = os.path.join(appdata, office, ver, "user")
                if os.path.isdir(p):
                    yield os.path.abspath(p)

    roaming = os.path.join(home, "AppData", "Roaming")
    if os.path.isdir(roaming):
        for office in ("AlterOffice3", "AlterOffice"):
            for ver in ("5", "4", "3"):
                p = os.path.join(roaming, office, ver, "user")
                if os.path.isdir(p):
                    yield os.path.abspath(p)


def find_active_alteroffice_target():
    """
    Найти один активный профиль AlterOffice.
    Приоритет: каталог макросов → UNO UserConfig → первый существующий AO-профиль.
    """
    macro_user = _user_dir_from_macro_lib()
    if macro_user and _is_alteroffice_user(macro_user):
        _log("profile from macro lib: %s" % macro_user)
        return _build_target(macro_user, "каталог макросов Scripts/python")

    uno_user = _uno_user_profile_dir()
    if uno_user and _is_alteroffice_user(uno_user):
        _log("profile from UNO UserConfig: %s" % uno_user)
        return _build_target(uno_user, "UNO UserConfig")

    for user_dir in _iter_candidate_user_roots():
        if _is_alteroffice_user(user_dir):
            _log("profile from scan: %s" % user_dir)
            return _build_target(user_dir, "найденный профиль AlterOffice")

    return None


def _build_target(user_dir, source_label):
    menubar_rel, popup_rel, toolbar_rel, images_rel = _ao_rel_dirs()
    toolbar_dir = _join_user(user_dir, toolbar_rel)
    images_dir = _join_user(user_dir, images_rel)
    return {
        "user": os.path.abspath(user_dir),
        "version": _profile_version_label(user_dir),
        "source": source_label,
        "menubar_dir": _join_user(user_dir, menubar_rel),
        "popup_dir": _join_user(user_dir, popup_rel),
        "toolbar_dir": toolbar_dir,
        "images_dir": images_dir,
        "menubar_path": _join_user(user_dir, menubar_rel + (MENUBAR_FILENAME,)),
        "cell_path": _join_user(user_dir, popup_rel + (CELL_FILENAME,)),
        "toolbar_path": _resolve_toolbar_path(toolbar_dir),
        "imagelist_path": os.path.join(images_dir, LC_IMAGELIST_FILENAME),
        "userimages_path": os.path.join(images_dir, "Bitmaps", LC_USERIMAGES_PNG_FILENAME),
    }


def _resolve_toolbar_path(toolbar_dir):
    """Существующий custom toolbar с нашим uiname, иначе канонический файл."""
    canonical = os.path.join(toolbar_dir, TOOLBAR_FILENAME)
    needle = 'toolbar:uiname="%s"' % TOOLBAR_UINAME
    if not os.path.isdir(toolbar_dir):
        return canonical
    try:
        names = os.listdir(toolbar_dir)
    except Exception:
        return canonical
    for name in names:
        if not str(name).startswith("custom_toolbar_") or not str(name).endswith(".xml"):
            continue
        path = os.path.join(toolbar_dir, name)
        try:
            with open(path, "r", encoding="utf-8") as fh:
                txt = fh.read()
        except Exception:
            continue
        if needle in txt:
            return path
    return canonical


def _show_custom_toolbar(doc):
    """Показать тулбар в текущем окне (после записи XML)."""
    resource = "private:resource/toolbar/%s" % TOOLBAR_RESOURCE
    try:
        frame = doc.getCurrentController().getFrame()
        lm = frame.LayoutManager
    except Exception as err:
        _log("toolbar LayoutManager: %s" % err)
        return False
    try:
        lm.createElement(resource)
    except Exception:
        pass
    try:
        lm.showElement(resource)
        _log("toolbar shown: %s" % resource)
        return True
    except Exception as err:
        _log("toolbar showElement: %s" % err)
        return False


def install_toolbar_to_user(user_dir):
    """
    Записать кастомный тулбар и иконки команд в профиль
    (для установщика ODS после копирования макросов).
    Возвращает list[str] строк отчёта.
    """
    notes = []
    if not user_dir or not os.path.isdir(user_dir):
        return notes
    if not _is_alteroffice_user(user_dir):
        notes.append("тулбар: пропущен (не профиль AlterOffice): %s" % user_dir)
        return notes
    _unused1, _unused2, toolbar_rel, _images_rel = _ao_rel_dirs()
    toolbar_dir = _join_user(user_dir, toolbar_rel)
    path = _resolve_toolbar_path(toolbar_dir)
    try:
        _write_text(path, TOOLBAR_XML)
        notes.append("OK  toolbar %s" % path)
    except Exception as err:
        notes.append("ERR toolbar %s: %s" % (path, err))
    notes.extend(install_toolbar_images_to_user(user_dir) or [])
    return notes


def install_toolbar_images_to_user(user_dir):
    """Записать lc_imagelist.xml + Bitmaps/lc_userimages.png в aoffice.cfg/…/images/."""
    notes = []
    if not user_dir or not os.path.isdir(user_dir):
        return notes
    if not _is_alteroffice_user(user_dir):
        return notes
    target = _build_target(user_dir, "install")
    try:
        from libre_macros_set_menu_images_blob import write_lc_imagelist_xml
        write_lc_imagelist_xml(target["imagelist_path"])
        notes.append("OK  imagelist %s" % target["imagelist_path"])
    except Exception as err:
        notes.append("ERR imagelist %s: %s" % (target["imagelist_path"], err))
    try:
        from libre_macros_set_menu_images_blob import write_lc_userimages_png
        write_lc_userimages_png(target["userimages_path"])
        notes.append("OK  userimages %s" % target["userimages_path"])
    except Exception as err:
        notes.append("ERR userimages %s: %s" % (target.get("userimages_path"), err))
    return notes


def _write_text(path, content):
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(content)
        if not content.endswith("\n"):
            fh.write("\n")


def _copy_to_target(target):
    results = []
    menubar = _menu_xml_with_version(MENUBAR_XML)
    cell = _menu_xml_with_version(CELL_XML)
    for kind, path, content in (
        ("menubar.xml", target["menubar_path"], menubar),
        ("cell.xml", target["cell_path"], cell),
        (TOOLBAR_FILENAME, target["toolbar_path"], TOOLBAR_XML),
    ):
        _log("write %s -> %s" % (kind, path))
        try:
            _write_text(path, content)
            results.append(("ok", kind, path, None))
        except Exception as err:
            results.append(("err", kind, path, err))
    for kind, path, writer_name in (
        (LC_IMAGELIST_FILENAME, target["imagelist_path"], "write_lc_imagelist_xml"),
        (LC_USERIMAGES_PNG_FILENAME, target["userimages_path"], "write_lc_userimages_png"),
    ):
        _log("write %s -> %s" % (kind, path))
        try:
            from libre_macros_set_menu_images_blob import (
                write_lc_imagelist_xml,
                write_lc_userimages_png,
            )
            writer = write_lc_imagelist_xml if writer_name.endswith("xml") else write_lc_userimages_png
            writer(path)
            results.append(("ok", kind, path, None))
        except Exception as err:
            results.append(("err", kind, path, err))
    return results


def set_menu_entry(doc=None, *args):
    """Точка входа: как convert_tables — doc из entry-скрипта (XSCRIPTCONTEXT)."""
    _ = args
    active = _resolve_doc(doc)
    _log("START")
    target = find_active_alteroffice_target()
    if target is None:
        msg = (
            "Профиль AlterOffice не найден.\n\n"
            "Ожидается каталог вида:\n"
            "  ~/.config/alteroffice/5/user/...\n"
            "или\n"
            "  %APPDATA%\\AlterOffice3\\5\\user\\...\n\n"
            "Сначала установите макросы через install.sh."
        )
        _log("profile not found")
        _show_message(active, "Установка меню", msg, is_error=True)
        return False

    results = _copy_to_target(target)
    ok = 0
    err = 0
    lines = [
        "Версия: %s" % MACRO_VERSION,
        "Профиль AlterOffice: версия %s" % target["version"],
        "Источник определения: %s" % target["source"],
        "Каталог user:",
        "  %s" % target["user"],
        "",
        "Записано:",
    ]
    for status, kind, path, error in results:
        if status == "ok":
            ok = ok + 1
            lines.append("  OK  %s" % path)
        else:
            err = err + 1
            lines.append("  ERR %s" % path)
            lines.append("      %s" % error)
    lines.append("")
    if err == 0 and ok >= 5:
        lines.append("Готово: menubar.xml, cell.xml, тулбар и иконки установлены.")
    elif err == 0:
        lines.append("Готово: записано файлов %d." % ok)
    else:
        lines.append("Ошибок: %d из %d." % (err, len(results)))
    shown = False
    if err == 0 and active is not None:
        shown = _show_custom_toolbar(active)
    lines.append("")
    if shown:
        lines.append("Тулбар «%s» показан в текущем окне." % TOOLBAR_UINAME)
        lines.append("Если пункты меню/иконки не обновились — перезапустите Calc.")
    else:
        lines.append("Перезапустите Calc для отображения меню, тулбара и иконок.")
        lines.append("Вид → Панели инструментов → %s" % TOOLBAR_UINAME)
    report = "\n".join(lines)
    _log("DONE profile=%s ok=%d err=%d" % (target["user"], ok, err))
    _show_message(active, "Установка меню", report, is_error=(err > 0))
    return err == 0 and ok > 0
