# -*- coding: utf-8 -*-
"""
Сборка самодостаточного MacroInstaller_bundled.ods:
внутри ODS — ZIP с .py из macro-lib; install_macros распаковывает и копирует
всё по назначению без диалогов (итоговая сводка); install_macros_dialogs — с диалогами.

Имя выходного файла по умолчанию: MacroInstaller_bundled_<версия>_<YYYYMMDD>.ods.
test.zip / docs.zip: установка в …/alter_office_macros/ (Windows: S:\\My Documents\\…).
Документация docs.zip — из каталога docs/ репозитория (.md, .docx, .pdf, figs/).

Обычный установщик: build_ods.py → MacroInstaller.ods
"""
import argparse
import fnmatch
import io
import os
import re
import subprocess
import sys
import zipfile
from datetime import datetime
import xml.etree.ElementTree as ET
import xml.sax.saxutils

from build_ods import (
    MANIFEST_NS,
    DEFAULT_DEBUG_SNAPSHOT,
    _patch_version_cell,
    _resolve_macro_version,
    disable_debug_flags_file,
    sync_macro_version_file,
    write_debug_flags_snapshot,
)
from ods_modify_password import (
    protect_ods_modify_password,
    resolve_installer_password,
)

# ==============================================================================
# КОД ВСТРОЕННОГО МАКРОСА (INSTALLER_BUNDLED.PY)
# ==============================================================================
INSTALLER_PY_CODE = r'''# -*- coding: utf-8 -*-
import uno
import unohelper
import os
import io
import shutil
import re
import zipfile
import tempfile
import traceback
from com.sun.star.awt import XActionListener

BUNDLED_ZIP_PATH = "bundled/macro-lib.zip"
BUNDLED_TEST_ZIP_PATH = "bundled/test.zip"
BUNDLED_DOCS_ZIP_PATH = "bundled/docs.zip"
WINDOWS_BUNDLE_ROOT = r"S:\My Documents\alter_office_macros"

# Всегда копируются в Scripts/python (без g_exportedScripts).
ALWAYS_INSTALL_PY = ("functions_pp.py", "functions_final.py")

# Последняя техническая причина ошибки при распаковке встроенного archive.
# Нужна, чтобы в LibreOffice было понятно, что именно пошло не так.
_BUNDLED_MACROS_ERROR = ""

# ------------------------------------------------------------------------------
# 1. Вспомогательные функции
# ------------------------------------------------------------------------------
def _is_windows_os():
    return os.name == "nt"


def _installer_dialog_parent():
    """Родитель диалогов — как param_wizard._wizard_dialog_parent."""
    try:
        frame = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()
        if frame is not None:
            return frame.getContainerWindow()
    except Exception:
        pass
    return None


def _installer_doc_parent_window():
    """Окно документа — как param_wizard._wp_doc_parent_window."""
    try:
        doc = XSCRIPTCONTEXT.getDocument()
        if doc is not None:
            return doc.getCurrentController().getFrame().getContainerWindow()
    except Exception:
        pass
    return _installer_dialog_parent()


def _installer_center_dialog_model(model, toolkit, parent_win=None):
    """Центрирование относительно окна документа или рабочей области."""
    if model is None or toolkit is None:
        return
    try:
        w = int(model.Width)
        h = int(model.Height)
        if parent_win is not None:
            ps = parent_win.getPosSize()
            model.PositionX = int(ps.X + max(0, (int(ps.Width) - w) / 2))
            model.PositionY = int(ps.Y + max(0, (int(ps.Height) - h) / 2))
            return
        wa = toolkit.getWorkArea()
        model.PositionX = int(wa.X + max(0, (int(wa.Width) - w) / 2))
        model.PositionY = int(wa.Y + max(0, (int(wa.Height) - h) / 2))
    except Exception:
        pass


def _installer_toolkit():
    try:
        return XSCRIPTCONTEXT.getDesktop().getToolkit()
    except Exception:
        pass
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        return ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx
        )
    except Exception:
        return None


def _installer_message_box(title, text, box_type=0, buttons=1, dialog=None):
    """MessageBox с перебором родителей — как param_wizard._wp_message_box_execute."""
    toolkit = _installer_toolkit()
    if toolkit is None:
        return None
    parents = []
    parent_doc = _installer_doc_parent_window()
    if parent_doc is not None:
        parents.append(parent_doc)
    parent_frame = _installer_dialog_parent()
    if parent_frame is not None and parent_frame not in parents:
        parents.append(parent_frame)
    if dialog is not None:
        try:
            peer = dialog.getPeer()
            if peer is not None and peer not in parents:
                parents.append(peer)
        except Exception:
            pass
    title_s = str(title)
    text_s = str(text)
    for parent in parents:
        try:
            box = toolkit.createMessageBox(parent, box_type, buttons, title_s, text_s)
            return box.execute()
        except Exception:
            pass
    try:
        box = toolkit.createMessageBox(None, box_type, buttons, title_s, text_s)
        return box.execute()
    except Exception:
        return None


def get_home_dir():
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        path_sub = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.util.PathSubstitution", ctx)
        home_url = path_sub.getSubstituteVariableValue("$(home)")
        return uno.fileUrlToSystemPath(home_url)
    except Exception:
        return os.environ.get("HOME") or os.environ.get("USERPROFILE") or os.path.expanduser("~")


def normalize_installer_path(raw):
    """%HOME% / $HOME → домашний каталог; слэши под текущую ОС."""
    text = str(raw or "").strip()
    if not text:
        return ""
    home = get_home_dir()
    text = re.sub(r"(?i)%HOME%", home, text)
    text = re.sub(r"(?i)\$HOME\b", home, text)
    if os.name != "nt":
        text = text.replace("\\", "/")
    text = os.path.expanduser(text)
    text = os.path.expandvars(text)
    return os.path.normpath(os.path.abspath(text))


def read_default_source_dir():
    """Путь по умолчанию: лист 1, строка 2, колонка A (ячейка A2)."""
    try:
        doc = XSCRIPTCONTEXT.getDocument()
        if doc is None:
            return ""
        sheets = doc.getSheets()
        sheet = None
        if sheets.hasByName("Sheet1"):
            sheet = sheets.getByName("Sheet1")
        elif sheets.getCount() > 0:
            sheet = sheets.getByIndex(0)
        if sheet is None:
            return ""
        cell = sheet.getCellByPosition(0, 1)
        raw = cell.String if cell.String else ""
        if not str(raw).strip():
            return ""
        return normalize_installer_path(raw)
    except Exception:
        return ""


def get_document_ods_path():
    """Путь к .ods на диске (для чтения вложенного ZIP)."""
    try:
        doc = XSCRIPTCONTEXT.getDocument()
        if doc is None:
            return ""
        url = str(doc.getURL() or "")
        if not url.lower().startswith("file:"):
            return ""
        path = uno.fileUrlToSystemPath(url)
        path = os.path.normpath(path)
        if path and os.path.isfile(path):
            return path
    except Exception:
        pass
    return ""


def extract_bundled_macros():
    """Распаковать bundled/macro-lib.zip из пакета документа во временную папку."""
    bundled_zip_path = "bundled/macro-lib.zip"
    error_details = ""
    ods_path = get_document_ods_path()
    if not ods_path or not os.path.isfile(ods_path):
        error_details = (
            "Не удалось определить путь к текущему ODS на диске "
            "(doc URL не file: или файл не найден)."
        )
        return None, error_details

    def _unpack_from(zip_path):
        # 1) Достаем вложенный macro-lib.zip из ODS (как обычный zip).
        try:
            with zipfile.ZipFile(zip_path, "r") as outer:
                names = outer.namelist()
                if bundled_zip_path not in names:
                    raise KeyError(
                        "В ODS нет файла '%s'. Внутри ODS есть: %s"
                        % (bundled_zip_path, ", ".join(names[:40]))
                    )
                payload = outer.read(bundled_zip_path)
        except Exception as e:
            raise RuntimeError("Ошибка чтения вложенного ZIP из ODS: " + str(e))

        # 2) Распаковываем вложенный zip в temp/.
        temp_dir = tempfile.mkdtemp(prefix="libre_macros_bundled_")
        try:
            with zipfile.ZipFile(io.BytesIO(payload), "r") as inner:
                inner.extractall(temp_dir)
            return temp_dir
        except Exception as e:
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass
            raise RuntimeError("Ошибка распаковки вложенного ZIP: " + str(e))

    try:
        return _unpack_from(ods_path), ""
    except Exception:
        error_details = traceback.format_exc()
    tmp_ods = None
    try:
        fd, tmp_ods = tempfile.mkstemp(prefix="libre_macros_ods_", suffix=".ods")
        os.close(fd)
        shutil.copy2(ods_path, tmp_ods)
        return _unpack_from(tmp_ods), ""
    except Exception:
        error_details = traceback.format_exc()
        return None, error_details
    finally:
        if tmp_ods:
            try:
                os.unlink(tmp_ods)
            except Exception:
                pass


def resolve_documents_dir():
    """
    Каталог «Документы» пользователя (Windows / Linux, EN и RU имена).
    Сначала — подстановки LibreOffice и XDG, затем типичные имена под $HOME.
    """
    home = get_home_dir()
    candidates = []

    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        path_sub = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.util.PathSubstitution", ctx)
        for rel in ("/Documents", "/My Documents", "/Документы", "/Мои документы", "/Мои Документы"):
            try:
                url = path_sub.substituteVariables("$(home)" + rel, True)
                if url:
                    candidates.append(uno.fileUrlToSystemPath(url))
            except Exception:
                pass
    except Exception:
        pass

    xdg_docs = os.environ.get("XDG_DOCUMENTS_DIR")
    if xdg_docs:
        candidates.append(xdg_docs)

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
    os.makedirs(dest, exist_ok=True)
    return dest


def resolve_bundle_install_dir():
    """
    Корень установки test/ и docs/ для bundled-пакета.
    Windows: S:\\My Documents\\alter_office_macros (или %USERPROFILE%\\My Documents\\…).
    Linux: «Документы»/alter_office_macros.
    """
    if _is_windows_os():
        windows_bundle_root = r"S:\My Documents\alter_office_macros"
        candidates = [
            windows_bundle_root,
            os.path.join(get_home_dir(), "My Documents", "alter_office_macros"),
            os.path.join(get_home_dir(), "Documents", "alter_office_macros"),
        ]
        seen = set()
        for path in candidates:
            if not path:
                continue
            path = os.path.normpath(path)
            if path in seen:
                continue
            seen.add(path)
            try:
                os.makedirs(path, exist_ok=True)
                return path
            except Exception:
                continue
        try:
            os.makedirs(windows_bundle_root, exist_ok=True)
        except Exception:
            pass
        return os.path.normpath(windows_bundle_root)
    docs_dir = resolve_documents_dir()
    dest = os.path.join(docs_dir, "alter_office_macros")
    os.makedirs(dest, exist_ok=True)
    return dest


def extract_bundled_tests():
    """
    Распаковать bundled/test.zip в …/alter_office_macros/test/.
    Возвращает путь к корню test/ или None.
    """
    ods_path = get_document_ods_path()
    if not ods_path or not os.path.isfile(ods_path):
        return None
    bundled_test_zip_path = "bundled/test.zip"
    try:
        with zipfile.ZipFile(ods_path, "r") as outer:
            if bundled_test_zip_path not in outer.namelist():
                return None
            payload = outer.read(bundled_test_zip_path)
        install_root = resolve_bundle_install_dir()
        with zipfile.ZipFile(io.BytesIO(payload), "r") as inner:
            inner.extractall(install_root)
        root = os.path.join(install_root, "test")
        if os.path.isdir(root):
            return root
        for walk_root, _dirs, _files in os.walk(install_root):
            if os.path.basename(walk_root) == "test":
                return walk_root
        return root
    except Exception:
        return None


def extract_bundled_docs():
    """
    Распаковать bundled/docs.zip в …/alter_office_macros/docs/.
    Возвращает путь к корню docs/ или None.
    """
    ods_path = get_document_ods_path()
    if not ods_path or not os.path.isfile(ods_path):
        return None
    bundled_docs_zip_path = "bundled/docs.zip"
    try:
        with zipfile.ZipFile(ods_path, "r") as outer:
            if bundled_docs_zip_path not in outer.namelist():
                return None
            payload = outer.read(bundled_docs_zip_path)
        install_root = resolve_bundle_install_dir()
        with zipfile.ZipFile(io.BytesIO(payload), "r") as inner:
            inner.extractall(install_root)
        root = os.path.join(install_root, "docs")
        if os.path.isdir(root):
            return root
        for walk_root, _dirs, _files in os.walk(install_root):
            if os.path.basename(walk_root) == "docs":
                return walk_root
        return root
    except Exception:
        return None


def resolve_source_dir():
    """
    Источник макросов: встроенный архив (без диалога) или выбор папки.
    Возвращает (путь, is_temp_dir).
    """
    bundled_dir, _err = extract_bundled_macros()
    if bundled_dir:
        return bundled_dir, True
    default_dir = read_default_source_dir()
    picked = pick_folder(
        "Выберите папку, откуда копировать макросы",
        default_dir,
    )
    if not picked:
        return None, False
    return picked, False


def get_user_scripts_path():
    ctx = XSCRIPTCONTEXT.getComponentContext()
    path_sub = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.util.PathSubstitution", ctx)
    user_url = path_sub.getSubstituteVariableValue("user")
    user_dir = uno.fileUrlToSystemPath(user_url)
    return os.path.join(user_dir, "Scripts", "python")


def pick_folder(title="Выберите папку с макросами", initial_dir=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    picker = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.ui.dialogs.FolderPicker", ctx)
    picker.setTitle(title)
    start = str(initial_dir or "").strip()
    if start:
        display = start if os.path.isdir(start) else os.path.dirname(start)
        if display and os.path.isdir(display):
            try:
                picker.setDisplayDirectory(uno.systemPathToFileUrl(display))
            except Exception:
                pass
    try:
        parent = _installer_dialog_parent()
        if parent is not None and hasattr(picker, "setParent"):
            picker.setParent(parent)
    except Exception:
        pass
    if picker.execute() == 1:
        return uno.fileUrlToSystemPath(picker.getDirectory())
    return None


def show_message(text, title="Информация", dialog=None):
    _installer_message_box(title, text, box_type=0, buttons=1, dialog=dialog)


def ask_yes_no(text, title="Вопрос", dialog=None):
    """QueryBox Да/Нет; True — перезаписать."""
    try:
        from com.sun.star.awt.MessageBoxType import QUERYBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO
        from com.sun.star.awt.MessageBoxResults import YES
    except Exception:
        QUERYBOX = 3
        BUTTONS_YES_NO = 4
        YES = 2
    result = _installer_message_box(
        title, text, box_type=QUERYBOX, buttons=BUTTONS_YES_NO, dialog=dialog
    )
    return result == YES


def _helper_py_details(filename):
    if filename == "functions_pp.py":
        return (
            "📌 Наименование: пользовательские функции постобработки\n\n"
            "📝 Описание: шаблоны для колонки B (functions_pp.py#имя_функции) "
            "на листе «Параметры_Объединения»."
        )
    if filename == "functions_final.py":
        return (
            "📌 Наименование: пользовательские функции финальной обработки\n\n"
            "📝 Описание: шаблоны для «Финальная_обработка» "
            "(functions_final.py#имя_функции)."
        )
    return "Вспомогательный модуль Python (без точки входа g_exportedScripts)."


def _ensure_always_install_in_files_info(files_info, source_dir):
    """Добавить functions_pp.py / functions_final.py в список установки."""
    for filename in ("functions_pp.py", "functions_final.py"):
        if filename in files_info:
            continue
        filepath = os.path.join(source_dir, filename)
        if os.path.isfile(filepath):
            files_info[filename] = {
                "macros": ["(модуль)"],
                "details": _helper_py_details(filename),
            }


def _collect_existing_install_paths(dest_dir, filenames, src_pythonpath):
    """Файлы, которые уже есть в каталоге макросов LibreOffice."""
    existing = []
    for filename in filenames:
        dst = os.path.join(dest_dir, filename)
        if os.path.isfile(dst):
            existing.append(filename)
    if os.path.isdir(src_pythonpath):
        dest_pythonpath = os.path.join(dest_dir, "pythonpath")
        for item in os.listdir(src_pythonpath):
            s = os.path.join(src_pythonpath, item)
            if os.path.isdir(s):
                continue
            d = os.path.join(dest_pythonpath, item)
            if os.path.isfile(d):
                existing.append("pythonpath/" + item)
    return existing


def _list_py_files_to_install(source_dir):
    """Все .py из пакета для установки (как при упаковке macro-lib.zip)."""
    names = []
    for filename in sorted(os.listdir(source_dir)):
        if (
            filename.endswith(".py")
            and not filename.startswith("test")
            and filename not in ("installer.py", "installer_bundled.py")
        ):
            names.append(filename)
    return names


def _copy_macros_to_dest(source_dir, selected_files, dest_dir):
    """Копирование выбранных .py и pythonpath/. Возвращает (copied_files, pythonpath_copied)."""
    os.makedirs(dest_dir, exist_ok=True)
    copied_files = []
    for filename in selected_files:
        src = os.path.join(source_dir, filename)
        if not os.path.isfile(src):
            continue
        dst = os.path.join(dest_dir, filename)
        shutil.copy2(src, dst)
        copied_files.append(filename)

    pythonpath_copied = False
    src_pythonpath = os.path.join(source_dir, "pythonpath")
    if os.path.isdir(src_pythonpath):
        dest_pythonpath = os.path.join(dest_dir, "pythonpath")
        os.makedirs(dest_pythonpath, exist_ok=True)
        for item in os.listdir(src_pythonpath):
            s = os.path.join(src_pythonpath, item)
            d = os.path.join(dest_pythonpath, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, dirs_exist_ok=True)
            else:
                shutil.copy2(s, d)
        pythonpath_copied = True
    return copied_files, pythonpath_copied


def _read_package_macro_version(source_dir):
    """Версия из распакованного пакета (version.txt / MACRO_VERSION)."""
    if source_dir and os.path.isdir(source_dir):
        vt = os.path.join(source_dir, "version.txt")
        if os.path.isfile(vt):
            try:
                with open(vt, encoding="utf-8") as fh:
                    ver = fh.read().strip()
                if ver:
                    return ver
            except Exception:
                pass
        for rel in (
            os.path.join("pythonpath", "libre_macros_lib.py"),
            "collect_workbooks.py",
            "version.py",
        ):
            path = os.path.join(source_dir, rel)
            if not os.path.isfile(path):
                continue
            try:
                with open(path, encoding="utf-8") as fh:
                    text = fh.read(4000)
                m = re.search(r'MACRO_VERSION\s*=\s*["\']([^"\']+)["\']', text)
                if m:
                    return m.group(1).strip()
            except Exception:
                pass
    return ""


def _user_dir_from_scripts_path(dest_dir):
    """…/user/Scripts/python → …/user."""
    d = os.path.abspath(dest_dir or "")
    if os.path.basename(d) == "python":
        scripts = os.path.dirname(d)
        if os.path.basename(scripts) == "Scripts":
            user_dir = os.path.dirname(scripts)
            if os.path.basename(user_dir) == "user":
                return user_dir
    return None


def _install_profile_toolbar(dest_scripts_dir):
    """После копирования макросов: тулбар + иконки в aoffice.cfg/acell/."""
    notes = []
    user_dir = _user_dir_from_scripts_path(dest_scripts_dir)
    if not user_dir:
        return notes
    try:
        import sys
        pp = os.path.join(dest_scripts_dir, "pythonpath")
        if pp and os.path.isdir(pp) and pp not in sys.path:
            sys.path.insert(0, pp)
        from libre_macros_set_menu_lib import install_toolbar_to_user
        notes.extend(install_toolbar_to_user(user_dir) or [])
    except Exception as err:
        notes.append("тулбар/иконки: не установлены (%s)" % err)
    return notes


def _patch_menu_xml_version_labels(xml_text, version):
    """Обновить/добавить подписи PythonMacros (ver) и Версия ver в XML меню."""
    ver = str(version or "").strip() or "0"
    xml = str(xml_text or "")
    if xml == "":
        return xml, False
    orig = xml

    def _sub_once(pattern, repl):
        nonlocal xml
        xml2, n = re.subn(pattern, repl, xml, count=1, flags=re.IGNORECASE)
        if n:
            xml = xml2

    _sub_once(
        r'(menu:id="vnd\.openoffice\.org:CustomMenu1"\s+menu:label=")PythonMacros(?:\s*\([^"]*\))?"',
        r'\1PythonMacros (%s)"' % ver,
    )
    _sub_once(
        r'(menu:id="PythonMacros"\s+menu:label=")PythonMacros(?:\s*\([^"]*\))?"',
        r'\1PythonMacros (%s)"' % ver,
    )
    xml2, n = re.subn(
        r'(version\.py\$show_version\?language=Python&amp;location=user"\s+menu:label=")[^"]*"',
        r'\1Версия %s"' % ver,
        xml,
        flags=re.IGNORECASE,
    )
    if n:
        xml = xml2
    elif "version.py$show_version" not in xml:
        item = (
            '<menu:menuseparator/>\n'
            '   <menu:menuitem menu:id="vnd.sun.star.script:version.py$show_version'
            '?language=Python&amp;location=user" menu:label="Версия %s"/>'
        ) % ver
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
            break
    xml = xml.replace("__LM_VER__", ver)
    return xml, xml != orig


def _patch_user_menu_versions(dest_scripts_dir, version):
    """
    Если у пользователя уже есть кастомные menubar.xml / cell.xml —
    обновить подписи PythonMacros и Версия (AO и при наличии LO).
    """
    notes = []
    user_dir = _user_dir_from_scripts_path(dest_scripts_dir)
    if not user_dir:
        return notes
    ver = str(version or "").strip()
    if not ver:
        return notes
    rels = (
        ("config", "aoffice.cfg", "modules", "acell", "menubar", "menubar.xml"),
        ("config", "aoffice.cfg", "modules", "acell", "popupmenu", "cell.xml"),
        ("config", "soffice.cfg", "modules", "scalc", "menubar", "menubar.xml"),
        ("config", "soffice.cfg", "modules", "scalc", "popupmenu", "cell.xml"),
    )
    seen = set()
    for parts in rels:
        path = os.path.join(user_dir, *parts)
        if path in seen or not os.path.isfile(path):
            continue
        seen.add(path)
        try:
            with open(path, "r", encoding="utf-8") as fh:
                old = fh.read()
        except Exception as err:
            notes.append("меню: ошибка чтения %s (%s)" % (path, err))
            continue
        if "PythonMacros" not in old and "version.py$show_version" not in old:
            continue
        new, changed = _patch_menu_xml_version_labels(old, ver)
        if not changed:
            continue
        try:
            parent = os.path.dirname(path)
            if parent and not os.path.isdir(parent):
                os.makedirs(parent, exist_ok=True)
            with open(path, "w", encoding="utf-8", newline="\n") as fh:
                fh.write(new)
                if not new.endswith("\n"):
                    fh.write("\n")
            notes.append("меню: обновлены подписи версии в %s" % path)
        except Exception as err:
            notes.append("меню: ошибка записи %s (%s)" % (path, err))
    return notes


def _build_install_report(copied_files, dest_dir, temp_source=False, pythonpath_copied=False):
    report = "Успешно установлено " + str(len(copied_files)) + " файл(ов):\n"
    report += "\n".join("  • " + f for f in copied_files)
    if temp_source:
        report += "\n\nИсточник: встроенный архив документа."
    if pythonpath_copied:
        report += "\n\nТакже скопировано содержимое папки 'pythonpath'."
    report += "\n\nПуть назначения:\n" + dest_dir
    return report

# ------------------------------------------------------------------------------
# 2. Парсинг Python-файлов
# ------------------------------------------------------------------------------
def extract_module_docstring(content):
    try:
        import ast
        tree = ast.parse(content)
        doc = ast.get_docstring(tree)
        if doc:
            return doc.strip()
    except Exception:
        pass
    return ""


def extract_header_comments(content):
    lines = content.splitlines()
    comments = []
    started = False
    for line in lines:
        stripped = line.strip()
        if not started:
            if stripped.startswith("#") and "coding" not in stripped.lower():
                started = True
                comments.append(stripped.lstrip("#").strip())
            elif stripped == "" or stripped.startswith("# -*-"):
                continue
            else:
                break
        else:
            if stripped.startswith("#"):
                comments.append(stripped.lstrip("#").strip())
            elif stripped == "":
                if comments:
                    break
            else:
                break
    return "\n".join(comments).strip()


def field_from_doc(docstring, field_pattern):
    if not docstring:
        return None
    m = re.search(
        r"(?is)" + field_pattern + r"\s*:\s*(.*?)(?=\n\s*(?:наименование|описание|назначение)\b|\Z)",
        docstring,
    )
    if m:
        val = m.group(1).strip()
        if val:
            return val
    return None


def format_macros_list(macros):
    if not macros:
        return "  • не определены"
    return "\n".join("  • " + m for m in macros)


def analyze_python_file(filepath):
    result = {"macros": [], "details": "Описание отсутствует."}
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
    except UnicodeDecodeError:
        try:
            with open(filepath, "r", encoding="cp1251") as f:
                content = f.read()
        except Exception:
            result["details"] = "⚠️ Не удалось прочитать файл (неподдерживаемая кодировка)."
            return result
    except Exception:
        result["details"] = "⚠️ Ошибка доступа к файлу."
        return result

    try:
        match_macros = re.search(r"g_exportedScripts\s*=\s*\((.*?)\)", content, re.DOTALL)
        if match_macros:
            macros_str = match_macros.group(1)
            result["macros"] = [m.strip().strip("\"'") for m in macros_str.split(",") if m.strip()]

        docstring = extract_module_docstring(content)
        header = extract_header_comments(content) if not docstring else ""

        name = field_from_doc(docstring, r"наименование(?:\s+макроса)?")
        desc = field_from_doc(docstring, r"описание(?:\s+макроса)?")
        purpose = field_from_doc(docstring, r"назначение(?:\s+(?:макроса|назначение))?")

        details_text = []
        if name:
            details_text.append("📌 Наименование: " + name)
        if desc:
            details_text.append("📝 Описание: " + desc)
        if purpose:
            details_text.append("🎯 Назначение: " + purpose)

        if details_text:
            result["details"] = "\n\n".join(details_text)
        elif docstring:
            result["details"] = "📝 Описание (докстринг модуля):\n\n" + docstring
        elif header:
            result["details"] = "📝 Описание (комментарии в начале файла):\n\n" + header
        else:
            result["details"] = (
                "⚠️ Стандартное описание не найдено.\n\n"
                "В файле обнаружены макросы:\n"
                + format_macros_list(result["macros"])
                + "\n\n"
                "💡 Совет: добавьте в начало файла докстринг с полями\n"
                "   «Наименование:», «Описание:» и «Назначение:»"
            )
    except Exception:
        result["details"] = "⚠️ Произошла ошибка при анализе содержимого файла."

    return result

# ------------------------------------------------------------------------------
# 3. Обработчик кнопки "Подробнее"
# ------------------------------------------------------------------------------
class DetailsListener(unohelper.Base, XActionListener):
    def __init__(self, ctx, details_text):
        self.ctx = ctx
        self.details_text = details_text

    def actionPerformed(self, event):
        show_details_popup(self.ctx, self.details_text)

    def disposing(self, event):
        pass


def show_details_popup(ctx, details_text):
    safe_text = str(details_text) if details_text else "Описание отсутствует."
    try:
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        parent_win = _installer_dialog_parent()

        model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        model.Width = 480
        model.Height = 320
        model.Title = "Подробная информация о макросе"
        _installer_center_dialog_model(model, toolkit, parent_win)

        txt_model = model.createInstance("com.sun.star.awt.UnoControlEditModel")
        if txt_model is None:
            raise RuntimeError("UnoControlEditModel недоступен")
        txt_model.Name = "txt_info"
        txt_model.Text = safe_text
        txt_model.PositionX = 10
        txt_model.PositionY = 10
        txt_model.Width = 460
        txt_model.Height = 260
        txt_model.ReadOnly = True
        txt_model.MultiLine = True
        try:
            txt_model.VScroll = True
            txt_model.HScroll = False
        except Exception:
            pass
        model.insertByName("txt_info", txt_model)

        btn_close = model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        if btn_close is None:
            raise RuntimeError("UnoControlButtonModel недоступен")
        btn_close.Name = "btn_close"
        btn_close.Label = "Закрыть"
        btn_close.PositionX = 180
        btn_close.PositionY = 280
        btn_close.Width = 100
        btn_close.Height = 25
        btn_close.PushButtonType = 1
        model.insertByName("btn_close", btn_close)

        dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        dialog.setModel(model)
        dialog.createPeer(toolkit, parent_win)
        dialog.execute()

    except Exception as e:
        fallback_text = (
            safe_text
            + "\n\n(Техническая заметка: не удалось отобразить расширенное окно.\n"
            + "Ошибка: "
            + str(e)
            + ")"
        )
        show_message(fallback_text, "Информация о макросе")

# ------------------------------------------------------------------------------
# 4. Основной диалог выбора
# ------------------------------------------------------------------------------
def show_selection_dialog(files_info, from_bundled=False):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent_win = _installer_dialog_parent()

    dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialog_model.Width = 450
    title = "Выберите макросы для установки"
    if from_bundled:
        title = "Выберите макросы (встроенный архив)"
    dialog_model.Title = title

    chk_names = {}
    y_pos = 10
    idx = 0
    for filename, data in files_info.items():
        macros_list = ", ".join(data["macros"])
        ctl_name = "chk_%d" % idx
        chk_names[ctl_name] = filename

        chk_model = dialog_model.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
        chk_model.Name = ctl_name
        chk_model.Label = filename + "  [" + macros_list + "]"
        chk_model.PositionX = 10
        chk_model.PositionY = y_pos
        chk_model.Width = 280
        chk_model.Height = 20
        chk_model.State = 1
        dialog_model.insertByName(ctl_name, chk_model)

        btn_info = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn_info.Name = "btn_info_%d" % idx
        btn_info.Label = "Подробнее..."
        btn_info.PositionX = 300
        btn_info.PositionY = y_pos
        btn_info.Width = 130
        btn_info.Height = 20
        btn_info.PushButtonType = 0
        dialog_model.insertByName(btn_info.Name, btn_info)

        y_pos += 30
        idx += 1

    dialog_model.Height = min(max(150, y_pos + 45), 700)
    _installer_center_dialog_model(dialog_model, toolkit, parent_win)

    btn_ok = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
    btn_ok.Name = "btn_ok"
    btn_ok.Label = "Установить выбранные"
    btn_ok.PositionX = 120
    btn_ok.PositionY = dialog_model.Height - 35
    btn_ok.Width = 130
    btn_ok.Height = 25
    btn_ok.PushButtonType = 1
    dialog_model.insertByName("btn_ok", btn_ok)

    btn_cancel = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
    btn_cancel.Name = "btn_cancel"
    btn_cancel.Label = "Отмена"
    btn_cancel.PositionX = 260
    btn_cancel.PositionY = dialog_model.Height - 35
    btn_cancel.Width = 130
    btn_cancel.Height = 25
    btn_cancel.PushButtonType = 2
    dialog_model.insertByName("btn_cancel", btn_cancel)

    dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dialog_model)
    dialog.createPeer(toolkit, parent_win)

    idx = 0
    for filename, data in files_info.items():
        listener = DetailsListener(ctx, data["details"])
        dialog.getControl("btn_info_%d" % idx).addActionListener(listener)
        idx += 1

    if dialog.execute() == 1:
        selected_files = []
        for ctl_name, filename in chk_names.items():
            try:
                if dialog.getControl(ctl_name).getState() == 1:
                    selected_files.append(filename)
            except Exception:
                pass
        return selected_files, dialog
    return [], dialog

# ------------------------------------------------------------------------------
# 5. Установка: install_macros (без диалогов) и install_macros_dialogs
# ------------------------------------------------------------------------------
def _append_bundle_extras(report):
    test_dir = extract_bundled_tests()
    if test_dir:
        report += "\n\nТестовые данные распакованы в:\n" + test_dir
    docs_dir = extract_bundled_docs()
    if docs_dir:
        report += "\n\nДокументация распакована в:\n" + docs_dir
    return report


def install_macros(*args):
    """Установка всего встроенного пакета без диалогов; итог — сводка."""
    source_dir = None
    details = ""
    try:
        source_dir, details = extract_bundled_macros()
        if not source_dir or not os.path.isdir(source_dir):
            details = str(details or "").strip()
            show_message(
                "Не удалось распаковать встроенный архив.\n"
                "Сохраните ODS на диск и повторите.\n"
                "Для установки с диалогами: install_macros_dialogs."
                + ("\n\nТехническая причина:\n" + details if details else ""),
                "Ошибка",
            )
            return

        selected_files = _list_py_files_to_install(source_dir)
        if not selected_files:
            show_message(
                "Во встроенном архиве нет .py для установки.",
                "Внимание",
            )
            return

        dest_dir = get_user_scripts_path()
        copied_files, pythonpath_copied = _copy_macros_to_dest(
            source_dir, selected_files, dest_dir
        )
        report = _build_install_report(
            copied_files, dest_dir, temp_source=True, pythonpath_copied=pythonpath_copied
        )
        menu_notes = _patch_user_menu_versions(
            dest_dir, _read_package_macro_version(source_dir)
        )
        menu_notes.extend(_install_profile_toolbar(dest_dir))
        if menu_notes:
            report += "\n\n" + "\n".join(menu_notes)
        report = _append_bundle_extras(report)
        show_message(report, "Установка завершена")

    except Exception as e:
        show_message(
            "Критическая ошибка:\n" + str(e) + "\n\n" + traceback.format_exc(),
            "Ошибка",
        )
    finally:
        if source_dir and os.path.isdir(source_dir):
            shutil.rmtree(source_dir, ignore_errors=True)


def install_macros_dialogs(*args):
    """Установка с диалогами выбора (как раньше install_macros)."""
    source_dir = None
    temp_source = False
    try:
        source_dir, temp_source = resolve_source_dir()
        if not source_dir:
            show_message("Операция отменена пользователем.")
            return
        if not os.path.isdir(source_dir):
            show_message("Папка не найдена:\n" + source_dir)
            return

        files_info = {}
        for filename in os.listdir(source_dir):
            if (
                filename.endswith(".py")
                and not filename.startswith("test")
                and filename not in ("installer.py", "installer_bundled.py")
            ):
                filepath = os.path.join(source_dir, filename)
                analysis = analyze_python_file(filepath)
                if analysis["macros"]:
                    files_info[filename] = analysis

        if not files_info:
            show_message(
                "Не найдено ни одного .py с g_exportedScripts = (...).\n"
                + ("Встроенный архив пуст или повреждён." if temp_source else "Проверьте выбранную папку."),
                "Внимание",
            )
            return

        _ensure_always_install_in_files_info(files_info, source_dir)

        selected_files, sel_dialog = show_selection_dialog(files_info, from_bundled=temp_source)
        if not selected_files:
            show_message(
                "Установка отменена (ничего не выбрано или закрыт диалог).",
                "Установка",
                dialog=sel_dialog,
            )
            return

        for helper in ("functions_pp.py", "functions_final.py"):
            if helper not in selected_files and helper in files_info:
                selected_files.append(helper)

        dest_dir = get_user_scripts_path()
        src_pythonpath = os.path.join(source_dir, "pythonpath")
        existing = _collect_existing_install_paths(
            dest_dir, selected_files, src_pythonpath
        )
        if len(existing) > 0:
            preview = existing[:12]
            more = len(existing) - len(preview)
            lines = "\n".join("  • " + x for x in preview)
            if more > 0:
                lines = lines + "\n  … и ещё " + str(more)
            if not ask_yes_no(
                "Уже установлены файлы макросов:\n" + lines + "\n\nПерезаписать?",
                "Переустановка",
                dialog=sel_dialog,
            ):
                show_message(
                    "Установка отменена. Существующие файлы не изменены.",
                    dialog=sel_dialog,
                )
                return

        copied_files, pythonpath_copied = _copy_macros_to_dest(
            source_dir, selected_files, dest_dir
        )
        report = _build_install_report(
            copied_files, dest_dir, temp_source=temp_source, pythonpath_copied=pythonpath_copied
        )
        menu_notes = _patch_user_menu_versions(
            dest_dir, _read_package_macro_version(source_dir)
        )
        menu_notes.extend(_install_profile_toolbar(dest_dir))
        if menu_notes:
            report += "\n\n" + "\n".join(menu_notes)
        report = _append_bundle_extras(report)
        show_message(report, "Установка завершена", dialog=sel_dialog)

    except Exception as e:
        show_message(
            "Критическая ошибка:\n" + str(e) + "\n\n" + traceback.format_exc(),
            "Ошибка",
        )
    finally:
        if temp_source and source_dir and os.path.isdir(source_dir):
            shutil.rmtree(source_dir, ignore_errors=True)


g_exportedScripts = (install_macros, install_macros_dialogs,)
'''

SCRIPT_FILE = "Scripts/python/installer_bundled.py"
SET_MENU_SCRIPT_FILE = "Scripts/python/set_menu.py"
SET_MENU_PYTHONPATH_DIR = "Scripts/python/pythonpath/"
SET_MENU_LIB_SCRIPT_FILE = "Scripts/python/pythonpath/libre_macros_set_menu_lib.py"
SET_MENU_CFG_SCRIPT_FILE = "Scripts/python/pythonpath/libre_macros_set_menu_cfg.py"
SET_MENU_IMAGES_BLOB_SCRIPT_FILE = (
    "Scripts/python/pythonpath/libre_macros_set_menu_images_blob.py"
)
BUNDLED_DIR = "bundled"
BUNDLED_ZIP = "bundled/macro-lib.zip"
BUNDLED_TEST_ZIP = "bundled/test.zip"
BUNDLED_DOCS_ZIP = "bundled/docs.zip"
BUNDLED_SOURCE_LABEL = "_Встроенный_Пакет_"

DEFAULT_MACRO_LIB = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "macro-lib")
)
SET_MENU_DOC_MODULES = (
    ("set_menu.py", SET_MENU_SCRIPT_FILE),
    (os.path.join("pythonpath", "libre_macros_set_menu_lib.py"), SET_MENU_LIB_SCRIPT_FILE),
    (os.path.join("pythonpath", "libre_macros_set_menu_cfg.py"), SET_MENU_CFG_SCRIPT_FILE),
    (
        os.path.join("pythonpath", "libre_macros_set_menu_images_blob.py"),
        SET_MENU_IMAGES_BLOB_SCRIPT_FILE,
    ),
)
DEFAULT_REPO_ROOT = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
)
DEFAULT_TEST_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "test")
)
DEFAULT_DOCS_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "docs")
)
DEFAULT_BUNDLE_EXCLUDE_PATHS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "bundle_exclude_paths.txt",
)

DOCS_ZIP_EXTENSIONS = frozenset({
    ".md",
    ".pdf",
    ".docx",
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".webp",
    ".svg",
})

ALWAYS_INSTALL_PY = ("functions_pp.py", "functions_final.py")

WINDOWS_BUNDLE_ROOT = r"S:\My Documents\alter_office_macros"
WINDOWS_BUNDLE_TEST_ROOT = WINDOWS_BUNDLE_ROOT + r"\test"

BUNDLE_PATH_PATCH_EXTS = frozenset({
    ".xml",
    ".xltx",
    ".xlsx",
    ".xlsm",
    ".ods",
    ".ots",
    ".md",
    ".py",
    ".txt",
    ".json",
    ".csv",
})

BUNDLE_OFFICE_ZIP_EXTS = frozenset({".xltx", ".xlsx", ".xlsm", ".ods", ".ots"})

DEFAULT_BUNDLED_OUTPUT_STEM = "MacroInstaller_bundled.ods"

def bundled_output_filename(version, build_date=None):
    """Имя ODS: MacroInstaller_bundled_<версия>_<YYYYMMDD>.ods"""
    if build_date is None:
        build_date = datetime.now().strftime("%Y%m%d")
    safe_ver = re.sub(r"[^\w.\-]+", "_", str(version).strip()) or "0"
    return "MacroInstaller_bundled_%s_%s.ods" % (safe_ver, build_date)


def resolve_bundled_output_path(output_path, version, build_date=None):
    if output_path in (None, "", DEFAULT_BUNDLED_OUTPUT_STEM):
        return bundled_output_filename(version, build_date=build_date)
    return output_path


def _posix_path(path):
    return os.path.abspath(path).replace("\\", "/").rstrip("/")


def _linux_test_path_to_windows(linux_path, test_posix, windows_test_root):
    p = str(linux_path).replace("\\", "/")
    if p == test_posix:
        return windows_test_root
    prefix = test_posix + "/"
    if p.startswith(prefix):
        rel = p[len(prefix) :].replace("/", "\\")
        return windows_test_root + "\\" + rel
    return linux_path


def rewrite_test_bundle_paths_in_text(text, test_dir_abs, windows_test_root):
    """Замена путей dev test/ на Windows-путь bundled-установки."""
    if not text:
        return text
    test_posix = _posix_path(test_dir_abs)
    patterns = [test_posix]
    test_native = os.path.normpath(test_dir_abs)
    if test_native.replace("\\", "/") != test_posix:
        patterns.append(test_native)

    def _replace_pattern(src, pat):
        if pat == test_posix:
            regex = re.escape(pat) + r"(?:/[^\"'<>\s]*)*"
        else:
            regex = re.escape(pat) + r"(?:\\[^\"'<>\s]*)*"

        def _repl(match):
            return _linux_test_path_to_windows(
                match.group(0), test_posix, windows_test_root
            )

        return re.sub(regex, _repl, src)

    out = text
    for pat in patterns:
        out = _replace_pattern(out, pat)
    return out


def rewrite_test_bundle_paths_in_bytes(data, test_dir_abs, windows_test_root):
    for enc in ("utf-8", "cp1251"):
        try:
            text = data.decode(enc)
        except UnicodeDecodeError:
            continue
        new_text = rewrite_test_bundle_paths_in_text(
            text, test_dir_abs, windows_test_root
        )
        if new_text != text:
            return new_text.encode(enc)
        return data
    return data


def _patch_office_zip_bytes(data, test_dir_abs, windows_test_root):
    in_buf = io.BytesIO(data)
    out_buf = io.BytesIO()
    with zipfile.ZipFile(in_buf, "r") as zin:
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                payload = zin.read(item.filename)
                lower = item.filename.lower()
                if lower.endswith((".xml", ".rels", ".vml")) or "/xml" in lower:
                    payload = rewrite_test_bundle_paths_in_bytes(
                        payload, test_dir_abs, windows_test_root
                    )
                zout.writestr(item, payload)
    return out_buf.getvalue()


def _prepare_test_file_bytes(full_path, test_dir_abs, windows_bundle_paths, windows_test_root):
    with open(full_path, "rb") as f:
        data = f.read()
    if not windows_bundle_paths:
        return data
    ext = os.path.splitext(full_path)[1].lower()
    if ext in BUNDLE_OFFICE_ZIP_EXTS:
        return _patch_office_zip_bytes(data, test_dir_abs, windows_test_root)
    if ext in BUNDLE_PATH_PATCH_EXTS:
        return rewrite_test_bundle_paths_in_bytes(
            data, test_dir_abs, windows_test_root
        )
    return data


TEST_ZIP_SKIP_DIR_NAMES = frozenset({"__pycache__", ".git"})
TEST_ZIP_SKIP_FILE_PREFIXES = (".~lock",)
TEST_ZIP_SKIP_REL_DIRS = frozenset({
    "collect_workbooks/workbooks/results",
})

ZIP_EXCLUDE_PY = frozenset({
    "generate_py2_version.py",
    "generate_no_docstrings.py",
    "collect_workbooks_no_docstrings.py",
    "collect_workbooks_py2.py",
    "pivot_table_wizard.py",
    "merge_param_presets.py",
    "installer.py",
    "installer_bundled.py",
})


def _repo_rel_posix(path, repo_root):
    try:
        rel = os.path.relpath(os.path.abspath(path), os.path.abspath(repo_root))
    except Exception:
        rel = str(path)
    return rel.replace("\\", "/")


def load_bundle_exclude_patterns(path):
    patterns = []
    if not path:
        return patterns
    if not os.path.isfile(path):
        return patterns
    with open(path, "r", encoding="utf-8") as fh:
        for line in fh:
            raw = line.strip()
            if not raw or raw.startswith("#"):
                continue
            patterns.append(raw.replace("\\", "/"))
    return patterns


def _bundle_path_excluded(rel_posix, patterns):
    if not patterns:
        return False
    rel = str(rel_posix or "").replace("\\", "/").lstrip("./")
    for pat in patterns:
        p = pat.lstrip("./")
        if fnmatch.fnmatch(rel, p):
            return True
        if fnmatch.fnmatch("/" + rel, p):
            return True
    return False

def _manifest_package_entries(include_tests=False, include_docs=False):
    entries = [
        ("Scripts/", "application/binary"),
        ("Scripts/python/", "application/binary"),
        (SCRIPT_FILE, "application/binary"),
        (SET_MENU_SCRIPT_FILE, "application/binary"),
        (SET_MENU_PYTHONPATH_DIR, "application/binary"),
        (SET_MENU_LIB_SCRIPT_FILE, "application/binary"),
        (SET_MENU_CFG_SCRIPT_FILE, "application/binary"),
        (SET_MENU_IMAGES_BLOB_SCRIPT_FILE, "application/binary"),
        (BUNDLED_DIR + "/", "application/binary"),
        (BUNDLED_ZIP, "application/binary"),
    ]
    if include_tests:
        entries.append((BUNDLED_TEST_ZIP, "application/binary"))
    if include_docs:
        entries.append((BUNDLED_DOCS_ZIP, "application/binary"))
    return tuple(entries)


def _ensure_manifest_entries(manifest_data, extra_entries):
    root = ET.fromstring(manifest_data)
    existing = {}
    for entry in root.findall(f"{{{MANIFEST_NS}}}file-entry"):
        path = entry.get(f"{{{MANIFEST_NS}}}full-path")
        if path:
            existing[path] = entry
    for full_path, media_type in extra_entries:
        entry = existing.get(full_path)
        if entry is None:
            entry = ET.SubElement(root, f"{{{MANIFEST_NS}}}file-entry")
            entry.set(f"{{{MANIFEST_NS}}}full-path", full_path)
        entry.set(f"{{{MANIFEST_NS}}}media-type", media_type)
    ET.register_namespace("manifest", MANIFEST_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_manifest_rdf(rdf_data, parts):
    text = rdf_data.decode("utf-8")
    block_parts = []
    for part_path in parts:
        if part_path in text:
            continue
        block_parts.append(
            f'  <rdf:Description rdf:about="{part_path}">\n'
            '    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/>\n'
            "  </rdf:Description>\n"
            '  <rdf:Description rdf:about="">\n'
            '    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" '
            f'rdf:resource="{part_path}"/>\n'
            "  </rdf:Description>\n"
        )
    if not block_parts:
        return rdf_data
    return text.replace("</rdf:RDF>", "".join(block_parts) + "</rdf:RDF>").encode("utf-8")


def _patch_form_button_event(xml_text, form_name, script_href):
    if script_href in xml_text:
        return xml_text
    event_block = (
        "<office:event-listeners>"
        '<script:event-listener script:language="ooo:script" '
        'script:event-name="form:performaction" '
        'xlink:href="%s" '
        'xlink:type="simple"/>'
        "</office:event-listeners>" % script_href
    )
    pattern = re.compile(
        r'(<form:button\b[^>]*form:name="%s"[^>]*>)(.*?)(</form:button>)'
        % re.escape(form_name),
        re.DOTALL,
    )

    def _replacer(match):
        open_tag, body, close_tag = match.group(1), match.group(2), match.group(3)
        props_close = body.find("</form:properties>")
        if props_close == -1:
            body = event_block + body
        else:
            pos = props_close + len("</form:properties>")
            body = body[:pos] + event_block + body[pos:]
        return open_tag + body + close_tag

    new_text, count = pattern.subn(_replacer, xml_text, count=1)
    if count:
        return new_text
    return xml_text


def _ensure_menu_button(content_data):
    text = content_data.decode("utf-8")
    if 'form:name="MenuInstallButton"' in text:
        return content_data

    button_inserted = False
    button_pattern = re.compile(
        r'(<form:button\b[^>]*form:name="InstallButton"[^>]*xml:id=")control1("'
        r'[^>]*form:id=")control1("'
        r'[^>]*form:label=")[^"]*("'
        r'[^>]*>.*?</form:button>)',
        re.DOTALL,
    )

    def _button_replacer(match):
        nonlocal button_inserted
        button_inserted = True
        original = match.group(0)
        extra = (
            original
            .replace('form:name="InstallButton"', 'form:name="MenuInstallButton"', 1)
            .replace('xml:id="control1"', 'xml:id="control2"', 1)
            .replace('form:id="control1"', 'form:id="control2"', 1)
            .replace('form:label="Запуск диалога установки"', 'form:label="Установить меню"', 1)
        )
        return original + extra

    text = button_pattern.sub(_button_replacer, text, count=1)
    if not button_inserted:
        return content_data

    draw_inserted = False
    draw_pattern = re.compile(
        r'(<draw:control\b[^>]*draw:name=")Control 1("'
        r'[^>]*svg:y=")0mm("'
        r'[^>]*draw:control=")control1("'
        r'[^>]*/>)',
        re.DOTALL,
    )

    def _draw_replacer(match):
        nonlocal draw_inserted
        draw_inserted = True
        original = match.group(0)
        extra = (
            match.group(1)
            + "Control 2"
            + match.group(2)
            + "16mm"
            + match.group(3)
            + "control2"
            + match.group(4)
        )
        return original + extra

    text = draw_pattern.sub(_draw_replacer, text, count=1)
    if not draw_inserted:
        return content_data

    return text.encode("utf-8")


def _patch_button_events(content_data):
    text = content_data.decode("utf-8")
    text = _patch_form_button_event(
        text,
        "InstallButton",
        "vnd.sun.star.script:installer_bundled.py$install_macros?language=Python&amp;location=document",
    )
    text = _patch_form_button_event(
        text,
        "MenuInstallButton",
        "vnd.sun.star.script:set_menu.py$set_menu?language=Python&amp;location=document",
    )
    return text.encode("utf-8")


def _patch_bundled_source_cell(content_data):
    """Ячейка A2 (строка ro2, колонка 1) — метка встроенного пакета."""
    text = content_data.decode("utf-8")
    label_xml = xml.sax.saxutils.escape(BUNDLED_SOURCE_LABEL)
    pattern = re.compile(
        r"(<table:table-row table:style-name=\"ro2\"[^>]*>)"
        r"(<table:table-cell\b)([^>]*>\s*<text:p>.*?</text:p>\s*</table:table-cell>)",
        re.DOTALL,
    )

    def _replacer(match):
        row_open, tag_open, tag_attrs = match.group(1), match.group(2), match.group(3)
        style_match = re.search(r'table:style-name="([^"]*)"', tag_attrs)
        style = style_match.group(1) if style_match else "ce3"
        filled_a2 = (
            f'<table:table-cell table:style-name="{style}" office:value-type="string" '
            f'calcext:value-type="string"><text:p>{label_xml}</text:p></table:table-cell>'
        )
        return row_open + filled_a2

    new_text, count = pattern.subn(_replacer, text, count=1)
    if count:
        return new_text.encode("utf-8")
    return content_data


def _read_document_macro_modules(macro_lib_dir):
    payloads = {}
    for rel_src, arcname in SET_MENU_DOC_MODULES:
        src = os.path.join(macro_lib_dir, rel_src)
        if not os.path.isfile(src):
            raise FileNotFoundError(f"Не найден документный модуль для кнопки меню: {src}")
        with open(src, "rb") as fh:
            payloads[arcname] = fh.read()
    return payloads


def _is_test_py_name(name):
    base = os.path.basename(str(name or ""))
    return base.startswith("test") and base.endswith(".py")


def _iter_packable_py_files(macro_lib_dir, exclude):
    macro_lib_dir = os.path.abspath(macro_lib_dir)
    for name in sorted(os.listdir(macro_lib_dir)):
        if not name.endswith(".py") or name in exclude or _is_test_py_name(name):
            continue
        full = os.path.join(macro_lib_dir, name)
        if os.path.isfile(full):
            yield full
    pypp = os.path.join(macro_lib_dir, "pythonpath")
    if os.path.isdir(pypp):
        for root, _dirs, files in os.walk(pypp):
            for fn in files:
                if not fn.endswith(".py") or _is_test_py_name(fn):
                    continue
                yield os.path.join(root, fn)


def sync_macro_version_tree(macro_lib_dir, version, exclude_py=None):
    """Перед упаковкой: синхронизировать MACRO_VERSION во всех .py пакета.

    version обычно из _resolve_macro_version (без --version — из libre_macros_lib.py).
    """
    exclude = set(exclude_py or ZIP_EXCLUDE_PY)
    updated = []
    for path in _iter_packable_py_files(macro_lib_dir, exclude):
        if sync_macro_version_file(path, version):
            updated.append(path)
    return updated


def disable_debug_flags_tree(macro_lib_dir, exclude_py=None, snapshot_path=None):
    """Перед упаковкой: *_DEBUG / DEBUG_LOG / ключи *DEBUG в dict → False; снимок для restore."""
    exclude = set(exclude_py or ZIP_EXCLUDE_PY)
    macro_lib_dir = os.path.abspath(macro_lib_dir)
    updated = []
    snapshot_files = {}
    for path in _iter_packable_py_files(macro_lib_dir, exclude):
        changed, module, dict_keys = disable_debug_flags_file(path)
        if changed:
            updated.append(path)
            rel = os.path.relpath(path, macro_lib_dir).replace("\\", "/")
            snapshot_files[rel] = {"module": module, "dict": dict_keys}
    snap = snapshot_path if snapshot_path is not None else DEFAULT_DEBUG_SNAPSHOT
    if snapshot_files:
        write_debug_flags_snapshot(snap, macro_lib_dir, snapshot_files)
    elif os.path.isfile(snap):
        os.remove(snap)
    return updated


def _flatten_aoffice_defs_tree(macro_lib_dir):
    """Свернуть многострочные def для совместимости с AlterOffice Temp/Scripts."""
    flatten_py = os.path.join(macro_lib_dir, "aoffice_flatten_defs.py")
    if not os.path.isfile(flatten_py):
        return
    import subprocess

    subprocess.check_call(
        [sys.executable, flatten_py, "--in-place", "--tree", macro_lib_dir]
    )


def _aoffice_compat_check_tree(macro_lib_dir):
    """Симуляция фильтра AO 2026: compile/exec entry Scripts/*.py."""
    check_py = os.path.join(macro_lib_dir, "aoffice_compat_check.py")
    if not os.path.isfile(check_py):
        return
    import subprocess

    env = dict(os.environ)
    pp = os.path.join(macro_lib_dir, "pythonpath")
    prev = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = pp + (os.pathsep + prev if prev else "")
    subprocess.check_call(
        [sys.executable, check_py, "--tree", macro_lib_dir],
        env=env,
    )


def build_macros_zip(macro_lib_dir, exclude_py=None, exclude_patterns=None, repo_root=None):
    """ZIP с .py из macro-lib (корень + pythonpath/) и version.txt."""
    macro_lib_dir = os.path.abspath(macro_lib_dir)
    if not os.path.isdir(macro_lib_dir):
        raise FileNotFoundError(f"Каталог macro-lib не найден: {macro_lib_dir}")

    _flatten_aoffice_defs_tree(macro_lib_dir)
    _aoffice_compat_check_tree(macro_lib_dir)

    exclude = set(exclude_py or ZIP_EXCLUDE_PY)
    repo_base = repo_root or DEFAULT_REPO_ROOT
    bundle_exclude = list(exclude_patterns or [])
    buf = io.BytesIO()
    file_count = 0

    for helper in ALWAYS_INSTALL_PY:
        helper_path = os.path.join(macro_lib_dir, helper)
        if not os.path.isfile(helper_path):
            raise FileNotFoundError(
                f"Обязательный модуль не найден: {helper_path}"
            )

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in sorted(os.listdir(macro_lib_dir)):
            if not name.endswith(".py") or name in exclude or _is_test_py_name(name):
                continue
            full = os.path.join(macro_lib_dir, name)
            if os.path.isfile(full):
                rel_repo = _repo_rel_posix(full, repo_base)
                if _bundle_path_excluded(rel_repo, bundle_exclude):
                    continue
                zout.write(full, arcname=name)
                file_count += 1

        version_path = os.path.join(macro_lib_dir, "version.txt")
        if os.path.isfile(version_path):
            zout.write(version_path, arcname="version.txt")
            file_count += 1

        pypp = os.path.join(macro_lib_dir, "pythonpath")
        if os.path.isdir(pypp):
            for root, _dirs, files in os.walk(pypp):
                for fn in files:
                    if not fn.endswith(".py") or _is_test_py_name(fn):
                        continue
                    full = os.path.join(root, fn)
                    rel_repo = _repo_rel_posix(full, repo_base)
                    if _bundle_path_excluded(rel_repo, bundle_exclude):
                        continue
                    rel = os.path.relpath(full, macro_lib_dir).replace("\\", "/")
                    zout.write(full, arcname=rel)
                    file_count += 1

    if file_count == 0:
        raise RuntimeError(f"В {macro_lib_dir} не найдено .py для упаковки")

    return buf.getvalue(), file_count


def _should_skip_test_path(rel_posix):
    rel_posix = rel_posix.replace("\\", "/")
    for skip_dir in TEST_ZIP_SKIP_REL_DIRS:
        if rel_posix == skip_dir or rel_posix.startswith(skip_dir + "/"):
            return True
    return False


def build_tests_zip(
    test_dir,
    windows_bundle_paths=True,
    windows_test_root=None,
    exclude_patterns=None,
    repo_root=None,
):
    """ZIP каталога test/ (arcname test/…) для распаковки в alter_office_macros/."""
    test_dir = os.path.abspath(test_dir)
    if not os.path.isdir(test_dir):
        raise FileNotFoundError(f"Каталог test не найден: {test_dir}")

    win_root = windows_test_root or WINDOWS_BUNDLE_TEST_ROOT
    repo_base = repo_root or DEFAULT_REPO_ROOT
    bundle_exclude = list(exclude_patterns or [])
    test_root_name = os.path.basename(test_dir.rstrip(os.sep))
    parent = os.path.dirname(test_dir)
    buf = io.BytesIO()
    file_count = 0
    patched_count = 0

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(test_dir):
            dirs[:] = [
                d for d in dirs
                if d not in TEST_ZIP_SKIP_DIR_NAMES and not d.startswith(".")
            ]
            for fn in files:
                if fn.startswith(TEST_ZIP_SKIP_FILE_PREFIXES):
                    continue
                full = os.path.join(root, fn)
                rel_repo = _repo_rel_posix(full, repo_base)
                if _bundle_path_excluded(rel_repo, bundle_exclude):
                    continue
                rel = os.path.relpath(full, parent).replace("\\", "/")
                if _should_skip_test_path(rel[len(test_root_name) + 1:] if rel.startswith(test_root_name + "/") else rel):
                    continue
                raw = _prepare_test_file_bytes(
                    full, test_dir, windows_bundle_paths, win_root
                )
                zout.writestr(rel, raw)
                file_count += 1
                if windows_bundle_paths:
                    with open(full, "rb") as f:
                        if f.read() != raw:
                            patched_count = patched_count + 1

    if file_count == 0:
        raise RuntimeError(f"В {test_dir} не найдено файлов для упаковки")

    return buf.getvalue(), file_count, patched_count


def build_docs_docx(docs_dir=None):
    """
    Сгенерировать docs/docx/*.docx из Markdown (pandoc).

    Вызывается перед упаковкой docs.zip. При ошибке сборка ODS продолжается
    (в zip попадут только уже существующие .docx).
    """
    docs_dir = os.path.abspath(docs_dir or DEFAULT_DOCS_DIR)
    script = os.path.join(docs_dir, "build_docx.py")
    if not os.path.isfile(script):
        print(f"⚠️  build_docx.py не найден: {script}")
        return False
    try:
        proc = subprocess.run(
            [sys.executable, script, "--all", "--docs-root", docs_dir],
            capture_output=True,
            text=True,
            timeout=600,
        )
    except Exception as exc:
        print(f"⚠️  build_docx.py: {exc}")
        return False
    if proc.stdout.strip():
        print(proc.stdout.strip())
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()
        print(f"⚠️  build_docx.py завершился с кодом {proc.returncode}: {err}")
        return False
    return True


def build_docs_zip(docs_dir=None, exclude_patterns=None, repo_root=None):
    """
    ZIP документации из каталога docs/ проекта для распаковки в «Документы».

    Структура в архиве повторяет исходный каталог: docs/00_INDEX.md, docs/docx/…, docs/figs/…
    Включаются .md, .docx, .pdf и иллюстрации (.png, .jpg, …).
    """
    docs_dir = os.path.abspath(docs_dir or DEFAULT_DOCS_DIR)
    repo_base = repo_root or DEFAULT_REPO_ROOT
    bundle_exclude = list(exclude_patterns or [])
    if not os.path.isdir(docs_dir):
        raise FileNotFoundError(f"Каталог документации не найден: {docs_dir}")

    buf = io.BytesIO()
    file_count = 0

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(docs_dir):
            dirs[:] = [
                d for d in dirs
                if d not in TEST_ZIP_SKIP_DIR_NAMES and not d.startswith(".")
            ]
            for fn in files:
                if fn.startswith(TEST_ZIP_SKIP_FILE_PREFIXES):
                    continue
                ext = os.path.splitext(fn)[1].lower()
                if ext not in DOCS_ZIP_EXTENSIONS:
                    continue
                full = os.path.join(root, fn)
                rel_repo = _repo_rel_posix(full, repo_base)
                if _bundle_path_excluded(rel_repo, bundle_exclude):
                    continue
                rel = os.path.relpath(full, docs_dir).replace("\\", "/")
                zout.write(full, arcname="docs/" + rel)
                file_count += 1

    if file_count == 0:
        raise RuntimeError(
            f"Не найдено файлов документации ({', '.join(sorted(DOCS_ZIP_EXTENSIONS))}) "
            f"в {docs_dir}"
        )

    return buf.getvalue(), file_count


def build_ods_bundled(
    template_path="template.ods",
    output_path=None,
    macro_lib_dir=None,
    version=None,
    protect=True,
    password=None,
    with_tests=False,
    test_dir=None,
    with_docs=True,
    docs_dir=None,
    build_docx=True,
    build_date=None,
    windows_test_root=None,
    windows_bundle_paths=True,
    exclude_paths_file=DEFAULT_BUNDLE_EXCLUDE_PATHS_FILE,
):
    macro_lib_dir = macro_lib_dir or DEFAULT_MACRO_LIB
    if not os.path.exists(template_path):
        print(
            f"❌ Ошибка: не найден шаблон '{template_path}' "
            "(создайте template.ods в LibreOffice)."
        )
        return False

    version_path = os.path.join(os.path.abspath(macro_lib_dir), "version.txt")
    macro_version = _resolve_macro_version(version, version_path=version_path)
    output_path = resolve_bundled_output_path(
        output_path, macro_version, build_date=build_date
    )
    exclude_patterns = load_bundle_exclude_patterns(exclude_paths_file)
    version_updated = sync_macro_version_tree(macro_lib_dir, macro_version)
    debug_disabled = disable_debug_flags_tree(macro_lib_dir)
    document_macro_files = _read_document_macro_modules(macro_lib_dir)
    bundled_zip, zip_count = build_macros_zip(
        macro_lib_dir,
        exclude_patterns=exclude_patterns,
        repo_root=DEFAULT_REPO_ROOT,
    )

    tests_zip = None
    tests_zip_count = 0
    tests_patched_count = 0
    if with_tests:
        tests_dir = test_dir or DEFAULT_TEST_DIR
        tests_zip, tests_zip_count, tests_patched_count = build_tests_zip(
            tests_dir,
            windows_bundle_paths=windows_bundle_paths,
            windows_test_root=windows_test_root,
            exclude_patterns=exclude_patterns,
            repo_root=DEFAULT_REPO_ROOT,
        )

    docs_zip = None
    docs_zip_count = 0
    if with_docs:
        docs_root = docs_dir or DEFAULT_DOCS_DIR
        if build_docx:
            build_docs_docx(docs_dir=docs_root)
        docs_zip, docs_zip_count = build_docs_zip(
            docs_dir=docs_root,
            exclude_patterns=exclude_patterns,
            repo_root=DEFAULT_REPO_ROOT,
        )

    with zipfile.ZipFile(template_path, "r") as zin:
        manifest_data = zin.read("META-INF/manifest.xml")

    manifest_entries = _manifest_package_entries(
        include_tests=with_tests,
        include_docs=with_docs,
    )
    updated_manifest = _ensure_manifest_entries(manifest_data, manifest_entries)
    rdf_parts = [
        SCRIPT_FILE,
        SET_MENU_SCRIPT_FILE,
        SET_MENU_LIB_SCRIPT_FILE,
        SET_MENU_CFG_SCRIPT_FILE,
        SET_MENU_IMAGES_BLOB_SCRIPT_FILE,
        BUNDLED_ZIP,
    ]
    if with_tests:
        rdf_parts.append(BUNDLED_TEST_ZIP)
    if with_docs:
        rdf_parts.append(BUNDLED_DOCS_ZIP)

    with zipfile.ZipFile(template_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "META-INF/manifest.xml":
                    data = updated_manifest
                elif item.filename == "manifest.rdf":
                    data = _patch_manifest_rdf(data, tuple(rdf_parts))
                elif item.filename == "content.xml":
                    data = _patch_bundled_source_cell(data)
                    data = _patch_version_cell(data, macro_version)
                    data = _ensure_menu_button(data)
                    data = _patch_button_events(data)
                if item.filename == "mimetype":
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)

            zout.writestr("Scripts/", b"")
            zout.writestr("Scripts/python/", b"")
            zout.writestr(SET_MENU_PYTHONPATH_DIR, b"")
            zout.writestr(SCRIPT_FILE, INSTALLER_PY_CODE.encode("utf-8"))
            for arcname, payload in document_macro_files.items():
                zout.writestr(arcname, payload)
            zout.writestr(BUNDLED_DIR + "/", b"")
            zout.writestr(BUNDLED_ZIP, bundled_zip)
            if with_tests and tests_zip is not None:
                zout.writestr(BUNDLED_TEST_ZIP, tests_zip)
            if with_docs and docs_zip is not None:
                zout.writestr(BUNDLED_DOCS_ZIP, docs_zip)

    zip_kb = len(bundled_zip) / 1024
    print(f"✅ Успешно создан: {os.path.abspath(output_path)}")
    print(f"   Версия: {macro_version}, в архиве файлов: {zip_count}, размер ZIP: {zip_kb:.1f} КБ")
    if with_tests and tests_zip is not None:
        tests_kb = len(tests_zip) / 1024
        print(
            f"   Тесты: {tests_zip_count} файлов, "
            f"путей Windows: {tests_patched_count}, "
            f"размер test.zip: {tests_kb:.1f} КБ"
        )
        print(f"   Целевой test: {windows_test_root or WINDOWS_BUNDLE_TEST_ROOT}")
    if with_docs and docs_zip is not None:
        docs_kb = len(docs_zip) / 1024
        print(f"   Документация: {docs_zip_count} файлов, размер docs.zip: {docs_kb:.1f} КБ")
    if version_updated:
        print(f"   MACRO_VERSION обновлён в {len(version_updated)} файле(ах)")
    if debug_disabled:
        print(f"   DEBUG-флаги сброшены в False в {len(debug_disabled)} файле(ах)")
        print(f"   Восстановить: python {os.path.join(os.path.dirname(__file__), 'restore_debug_flags.py')}")

    if protect:
        modify_password = resolve_installer_password(password)
        protect_ods_modify_password(output_path, modify_password)
        print("   Защита от изменений: включена (пароль в installer_password.txt)")

    return True


def main():
    parser = argparse.ArgumentParser(
        description="Сборка MacroInstaller_bundled.ods со встроенным архивом macro-lib",
    )
    parser.add_argument(
        "--template",
        default="template.ods",
        help="Шаблон ODS (по умолчанию: template.ods)",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_BUNDLED_OUTPUT_STEM,
        help="Выходной ODS (по умолчанию: MacroInstaller_bundled_<версия>_<дата>.ods)",
    )
    parser.add_argument(
        "--macro-lib-dir",
        default=DEFAULT_MACRO_LIB,
        help=f"Каталог с макросами (по умолчанию: {DEFAULT_MACRO_LIB})",
    )
    parser.add_argument(
        "--version",
        dest="macro_version",
        metavar="VER",
        help="Версия: записать в macro-lib/version.txt и в ячейку B2 листа",
    )
    parser.add_argument(
        "--no-protect",
        action="store_true",
        help="Не ставить пароль на изменение ODS",
    )
    parser.add_argument(
        "--password",
        metavar="PWD",
        help="Пароль защиты от изменений (иначе env / installer_password.txt / по умолчанию)",
    )
    parser.add_argument(
        "--with-tests",
        action="store_true",
        help="Вложить test.zip; распаковка в …/alter_office_macros/test",
    )
    parser.add_argument(
        "--test-dir",
        default=DEFAULT_TEST_DIR,
        help=f"Каталог test для упаковки (по умолчанию: {DEFAULT_TEST_DIR})",
    )
    parser.add_argument(
        "--no-docs",
        action="store_true",
        help="Не включать docs.zip (.md + .docx + .pdf) в ODS",
    )
    parser.add_argument(
        "--skip-docx-build",
        action="store_true",
        help="Не вызывать docs/build_docx.py перед упаковкой (только существующие .docx)",
    )
    parser.add_argument(
        "--docs-dir",
        default=DEFAULT_DOCS_DIR,
        help=f"Каталог документации для docs.zip (по умолчанию: {DEFAULT_DOCS_DIR})",
    )
    parser.add_argument(
        "--build-date",
        metavar="YYYYMMDD",
        help="Дата в имени файла (по умолчанию — сегодня)",
    )
    parser.add_argument(
        "--windows-test-root",
        default=WINDOWS_BUNDLE_TEST_ROOT,
        help=f"Корень test в шаблонах Windows (по умолчанию: {WINDOWS_BUNDLE_TEST_ROOT})",
    )
    parser.add_argument(
        "--no-windows-test-paths",
        action="store_true",
        help="Не переписывать пути test/ на Windows при упаковке test.zip",
    )
    parser.add_argument(
        "--exclude-paths-file",
        default=DEFAULT_BUNDLE_EXCLUDE_PATHS_FILE,
        help="Файл с glob-исключениями путей для macro-lib/test/docs",
    )
    args = parser.parse_args()
    with_docs = not args.no_docs
    ok = build_ods_bundled(
        template_path=args.template,
        output_path=args.output,
        macro_lib_dir=args.macro_lib_dir,
        version=args.macro_version,
        protect=not args.no_protect,
        password=args.password,
        with_tests=args.with_tests,
        test_dir=args.test_dir,
        with_docs=with_docs,
        docs_dir=args.docs_dir,
        build_docx=not args.skip_docx_build,
        build_date=args.build_date,
        windows_test_root=args.windows_test_root,
        windows_bundle_paths=not args.no_windows_test_paths,
        exclude_paths_file=args.exclude_paths_file,
    )
    raise SystemExit(0 if ok else 1)


if __name__ == "__main__":
    main()
