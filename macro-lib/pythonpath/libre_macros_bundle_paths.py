# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.710"
"""
Пути к bundled test/sources — та же логика, что в installer/build_ods_bundled.py.

Windows: S:\\My Documents\\alter_office_macros\\test\\collect_workbooks\\sources\\…
Linux:   $HOME/Documents/alter_office_macros/test/collect_workbooks/sources/…
"""

import os
import sys

try:
    unicode
except NameError:
    unicode = str

WINDOWS_BUNDLE_ROOT = r"S:\My Documents\alter_office_macros"
BUNDLE_DIR_NAME = "alter_office_macros"
TEST_SOURCES_PARTS = ("test", "collect_workbooks", "sources")


def is_windows_os():
    return os.name == "nt" or str(sys.platform).lower().startswith("win")


def system_path(raw):
    """
    file:// URL или обычный путь → нативный абсолютный путь.
    В LibreOffice __file__ часто приходит как file:///home/user/…
    """
    text = unicode(raw or u"").strip()
    if text == u"":
        return u""
    low = text.lower()
    if low.startswith("file://"):
        try:
            import uno

            return os.path.normpath(uno.fileUrlToSystemPath(text))
        except Exception:
            rest = text[7:]
            if rest.startswith("//"):
                rest = rest[1:]
            elif not rest.startswith("/"):
                rest = "/" + rest
            return os.path.normpath(rest)
    text = os.path.expanduser(text)
    text = os.path.expandvars(text)
    return os.path.normpath(os.path.abspath(text))


def is_file_url(raw):
    """True для file://… (локальный файл в нотации LO)."""
    return unicode(raw or u"").strip().lower().startswith("file://")


def is_remote_url(raw):
    """True для smb://, http://, …; file:// — False (см. normalize_source_path)."""
    text = unicode(raw or u"").strip()
    if text == u"" or is_file_url(text):
        return False
    return "://" in text


def normalize_source_path(raw):
    """
    Путь источника для сбора: file:// → системный путь;
    smb:// и прочие URL — без изменений; обычный путь — strip.
    """
    text = unicode(raw or u"").strip()
    if text == u"" or is_remote_url(text):
        return text
    if is_file_url(text):
        return system_path(text)
    return text


def get_home_dir():
    try:
        import uno

        xsc = None
        try:
            import __main__

            xsc = getattr(__main__, "XSCRIPTCONTEXT", None)
        except Exception:
            pass
        if xsc is not None:
            ctx = xsc.getComponentContext()
            path_sub = ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.util.PathSubstitution", ctx
            )
            home_url = path_sub.getSubstituteVariableValue("$(home)")
            home = system_path(home_url)
            if home:
                return home
    except Exception:
        pass
    for key in ("HOME", "USERPROFILE"):
        val = os.environ.get(key)
        if val:
            home = system_path(val)
            if home:
                return home
    return system_path(u"~")


def resolve_documents_dir():
    """Каталог «Документы» пользователя (Windows / Linux, EN и RU имена)."""
    home = get_home_dir()
    candidates = []
    xdg_docs = os.environ.get("XDG_DOCUMENTS_DIR")
    if xdg_docs:
        candidates.append(system_path(xdg_docs))
    for name in (
        "Documents",
        "My Documents",
        "Документы",
        "Мои документы",
        "Мои Документы",
        "Documentos",
    ):
        candidates.append(os.path.join(home, name))
    seen = set()
    for path in candidates:
        if not path:
            continue
        path = os.path.normpath(os.path.abspath(path))
        if path in seen:
            continue
        seen.add(path)
        if os.path.isdir(path):
            return path
    dest = os.path.join(home, "Documents")
    try:
        os.makedirs(dest, exist_ok=True)
    except Exception:
        pass
    return dest


def resolve_bundle_install_dir():
    """
    Корень установки test/ и docs/ для bundled-пакета.
    Windows: S:\\My Documents\\alter_office_macros (или %USERPROFILE%\\My Documents\\…).
    Linux: $HOME/Documents/alter_office_macros.
    """
    if is_windows_os():
        candidates = [
            WINDOWS_BUNDLE_ROOT,
            os.path.join(get_home_dir(), "My Documents", BUNDLE_DIR_NAME),
            os.path.join(get_home_dir(), "Documents", BUNDLE_DIR_NAME),
        ]
        seen = set()
        for path in candidates:
            if not path:
                continue
            path = os.path.normpath(path)
            if path in seen:
                continue
            seen.add(path)
            return path
        return os.path.normpath(WINDOWS_BUNDLE_ROOT)
    return os.path.join(resolve_documents_dir(), BUNDLE_DIR_NAME)


def _macro_lib_dir_from_anchor(anchor_file):
    anchor = system_path(anchor_file or __file__)
    macro_lib_dir = os.path.dirname(anchor)
    if os.path.basename(macro_lib_dir) == "pythonpath":
        macro_lib_dir = os.path.dirname(macro_lib_dir)
    return macro_lib_dir


def repo_dev_sources_dir(anchor_file=None):
    """
    Sources в git-репозитории (только если каталог реально содержит тестовые файлы).
    Не путать с ложным Scripts/test/… под каталогом макросов LibreOffice.
    """
    macro_lib_dir = _macro_lib_dir_from_anchor(anchor_file)
    repo_root = os.path.dirname(macro_lib_dir)
    dev = os.path.join(repo_root, "test", "collect_workbooks", "sources")
    dev = os.path.normpath(dev)
    for marker in ("source_01.xlsx", "source_pivot.xlsx"):
        if os.path.isfile(os.path.join(dev, marker)):
            return dev
    return None


def resolve_test_sources_dir(anchor_file=None):
    """
    Каталог sources для примеров пресетов — bundled-установка:
    $HOME/Documents/alter_office_macros/test/collect_workbooks/sources (Linux)
    или S:\\My Documents\\alter_office_macros\\… (Windows).
    """
    return os.path.normpath(
        os.path.join(resolve_bundle_install_dir(), *TEST_SOURCES_PARTS)
    )


def resolve_example_source_path(anchor_file, *parts):
    root = resolve_test_sources_dir(anchor_file)
    out = root
    i = 0
    while i < len(parts):
        out = os.path.join(out, unicode(parts[i]))
        i = i + 1
    return os.path.normpath(out)
