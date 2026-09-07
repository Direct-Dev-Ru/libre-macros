# -*- coding: utf-8 -*-
import argparse
import json
import zipfile
import os
import re
import sys
from datetime import datetime, timezone
import xml.etree.ElementTree as ET
import xml.sax.saxutils

_INSTALLER_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DEBUG_SNAPSHOT = os.path.join(_INSTALLER_DIR, ".bundled_debug_snapshot.json")

# ==============================================================================
# КОД ВСТРОЕННОГО МАКРОСА (INSTALLER.PY)
# ==============================================================================
INSTALLER_PY_CODE = r'''# -*- coding: utf-8 -*-
import uno
import unohelper
import os
import shutil
import re
import traceback
from com.sun.star.awt import XActionListener

# Всегда копируются в Scripts/python (без g_exportedScripts).
ALWAYS_INSTALL_PY = ("functions_pp.py", "functions_final.py")

# ------------------------------------------------------------------------------
# 1. Вспомогательные функции
# ------------------------------------------------------------------------------
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
    if picker.execute() == 1:
        return uno.fileUrlToSystemPath(picker.getDirectory())
    return None

def show_message(text, title="Информация"):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    frame = desktop.getCurrentFrame()
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent_win = frame.getContainerWindow() if frame else toolkit.getDesktopWindow()
    # 0 = info icon, 1 = OK button
    box = toolkit.createMessageBox(parent_win, 0, 1, title, str(text))
    box.execute()

# ------------------------------------------------------------------------------
# 2. Парсинг Python-файлов
# ------------------------------------------------------------------------------
def extract_module_docstring(content):
    """Докстринг модуля (не первый docstring внутри функций)."""
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
    """Комментарии # в начале файла, если нет модульного докстринга."""
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
    """Значение поля до следующего «Наименование/Описание/Назначение»."""
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
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        try:
            with open(filepath, 'r', encoding='cp1251') as f:
                content = f.read()
        except Exception:
            result["details"] = "⚠️ Не удалось прочитать файл (неподдерживаемая кодировка)."
            return result
    except Exception:
        result["details"] = "⚠️ Ошибка доступа к файлу."
        return result

    try:
        match_macros = re.search(r'g_exportedScripts\s*=\s*\((.*?)\)', content, re.DOTALL)
        if match_macros:
            macros_str = match_macros.group(1)
            result["macros"] = [m.strip().strip('"\'') for m in macros_str.split(',') if m.strip()]

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
    for filename in ALWAYS_INSTALL_PY:
        if filename in files_info:
            continue
        filepath = os.path.join(source_dir, filename)
        if os.path.isfile(filepath):
            files_info[filename] = {
                "macros": ["(модуль)"],
                "details": _helper_py_details(filename),
            }

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
    """Показывает окно с информацией. Имеет защиту от сбоев UI."""
    safe_text = str(details_text) if details_text else "Описание отсутствует."
    try:
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        parent_win = None
        try:
            frame = desktop.getCurrentFrame()
            if frame is not None:
                parent_win = frame.getContainerWindow()
        except Exception:
            pass
        if parent_win is None:
            parent_win = toolkit.getDesktopWindow()

        model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        model.Width = 480
        model.Height = 320
        model.Title = "Подробная информация о макросе"

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
def show_selection_dialog(files_info):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    frame = desktop.getCurrentFrame()
    parent_win = frame.getContainerWindow() if frame else toolkit.getDesktopWindow()

    dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialog_model.Width = 450
    dialog_model.Height = min(100 + len(files_info) * 30, 450)
    dialog_model.Title = "Выберите макросы для установки"

    dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dialog_model)
    if parent_win:
        dialog.createPeer(toolkit, parent_win)

    y_pos = 10
    for filename, data in files_info.items():
        macros_list = ", ".join(data["macros"])
        
        chk_model = dialog_model.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
        chk_model.Name = "chk_" + filename
        chk_model.Label = f"{filename}  [{macros_list}]"
        chk_model.PositionX = 10
        chk_model.PositionY = y_pos
        chk_model.Width = 280
        chk_model.Height = 20
        chk_model.State = 1
        dialog_model.insertByName("chk_" + filename, chk_model)
        
        btn_info = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn_info.Name = "btn_info_" + filename
        btn_info.Label = "Подробнее..."
        btn_info.PositionX = 300
        btn_info.PositionY = y_pos
        btn_info.Width = 130
        btn_info.Height = 20
        btn_info.PushButtonType = 0
        dialog_model.insertByName("btn_info_" + filename, btn_info)
        
        listener = DetailsListener(ctx, data["details"])
        dialog.getControl("btn_info_" + filename).addActionListener(listener)
        
        y_pos += 30

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

    if dialog.execute() == 1:
        selected_files = []
        for filename in files_info.keys():
            if dialog.getControl("chk_" + filename).getState() == 1:
                selected_files.append(filename)
        return selected_files
    return []

# ------------------------------------------------------------------------------
# 5. Главная функция установки
# ------------------------------------------------------------------------------
def install_macros(*args):
    try:
        default_dir = read_default_source_dir()
        source_dir = pick_folder(
            "Выберите папку, откуда копировать макросы",
            default_dir,
        )
        if not source_dir:
            show_message("Операция отменена пользователем.")
            return
        if not os.path.isdir(source_dir):
            show_message(f"Папка не найдена:\n{source_dir}")
            return

        files_info = {}
        for filename in os.listdir(source_dir):
            if filename.endswith(".py") and filename != "installer.py":
                filepath = os.path.join(source_dir, filename)
                analysis = analyze_python_file(filepath)
                if analysis["macros"]: 
                    files_info[filename] = analysis

        if not files_info:
            show_message("В выбранной папке не найдено ни одного .py файла\nс объявлением g_exportedScripts = (...)", "Внимание")
            return

        _ensure_always_install_in_files_info(files_info, source_dir)

        selected_files = show_selection_dialog(files_info)
        if not selected_files:
            show_message("Ни один макрос не выбран. Установка отменена.")
            return

        for helper in ALWAYS_INSTALL_PY:
            if helper not in selected_files and helper in files_info:
                selected_files.append(helper)

        dest_dir = get_user_scripts_path()
        os.makedirs(dest_dir, exist_ok=True)
        
        copied_files = []
        for filename in selected_files:
            src = os.path.join(source_dir, filename)
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

        report = f"✅ Успешно установлено {len(copied_files)} макрос(ов):\n"
        report += "\n".join(f"  • {f}" for f in copied_files)
        if pythonpath_copied:
            report += "\n\n📁 Также скопировано содержимое папки 'pythonpath'."
        report += f"\n\n📂 Путь назначения:\n{dest_dir}"
        
        show_message(report, "Установка завершена")
        
    except Exception as e:
        show_message("Критическая ошибка:\n" + str(e) + "\n\n" + traceback.format_exc(), "Ошибка")

g_exportedScripts = (install_macros,)
'''

MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
SCRIPT_FILE = "Scripts/python/installer.py"
INSTALL_MACRO_HREF = (
    "vnd.sun.star.script:installer.py$install_macros"
    "?language=Python&location=document"
)

MANIFEST_SCRIPT_ENTRIES = (
    ("Scripts/", "application/binary"),
    ("Scripts/python/", "application/binary"),
    (SCRIPT_FILE, "application/binary"),
)
VERSION_FILE = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "macro-lib", "version.txt")
)


def _read_macro_version(version_path=VERSION_FILE):
    with open(version_path, encoding="utf-8") as f:
        return f.read().strip()


def _read_canonical_macro_version():
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    installer_dir = os.path.dirname(os.path.abspath(__file__))
    if installer_dir not in sys.path:
        sys.path.insert(0, installer_dir)
    from macro_version import read_canonical_macro_version

    return read_canonical_macro_version(repo)


def _write_macro_version(version, version_path=VERSION_FILE):
    text = str(version).strip()
    if not text:
        raise ValueError("Версия не может быть пустой")
    parent = os.path.dirname(version_path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(version_path, "w", encoding="utf-8") as f:
        f.write(text + "\n")
    return text


def _resolve_macro_version(version_arg=None, version_path=VERSION_FILE):
    if version_arg is not None and str(version_arg).strip():
        return _write_macro_version(version_arg, version_path)
    try:
        canonical = _read_canonical_macro_version()
        if canonical and canonical != "0.0.0":
            return _write_macro_version(canonical, version_path)
    except Exception:
        pass
    return _read_macro_version(version_path)


MACRO_VERSION_ASSIGN_RE = re.compile(
    r'^MACRO_VERSION\s*=\s*(["\']).*?\1\s*$',
    re.MULTILINE,
)
CODING_LINE_RE = re.compile(
    r"^#\s*(?:-\*-)?\s*coding[:=]\s*[\w.-]+\s*(?:-\*-)?\s*$",
    re.IGNORECASE,
)
MODULE_DOCSTRING_RE = re.compile(
    r'^\s*(?:"""[\s\S]*?"""|\'\'\'[\s\S]*?\'\'\')\s*',
    re.MULTILINE,
)
FUTURE_IMPORT_RE = re.compile(
    r"^from __future__ import [^\n]+\n",
    re.MULTILINE,
)


def _macro_version_insert_pos(content):
    """Позиция вставки MACRO_VERSION: после coding, докстринга и from __future__."""
    pos = 0
    lines = content.splitlines(True)
    if lines and CODING_LINE_RE.match(lines[0].strip()):
        pos = len(lines[0])
    rest = content[pos:]
    match = MODULE_DOCSTRING_RE.match(rest)
    if match:
        pos += match.end()
        rest = content[pos:]
    while True:
        match = FUTURE_IMPORT_RE.match(rest)
        if not match:
            break
        pos += match.end()
        rest = content[pos:]
    return pos


def sync_macro_version_in_source(content, version):
    """Обновить или добавить MACRO_VERSION = \"…\" в тексте .py."""
    version = str(version).strip()
    new_line = f'MACRO_VERSION = "{version}"'
    if MACRO_VERSION_ASSIGN_RE.search(content):
        new_content = MACRO_VERSION_ASSIGN_RE.sub(new_line, content, count=1)
        return new_content, new_content != content
    pos = _macro_version_insert_pos(content)
    insert = new_line + "\n"
    if pos < len(content) and not content[pos:].startswith("\n"):
        insert += "\n"
    new_content = content[:pos] + insert + content[pos:]
    return new_content, True


def sync_macro_version_file(path, version):
    try:
        with open(path, encoding="utf-8") as f:
            content = f.read()
    except UnicodeDecodeError:
        with open(path, encoding="cp1251") as f:
            content = f.read()
        encoding = "cp1251"
    else:
        encoding = "utf-8"
    new_content, changed = sync_macro_version_in_source(content, version)
    if changed:
        with open(path, "w", encoding=encoding) as f:
            f.write(new_content)
    return changed


# Сброс отладочных флагов перед release/bundled (см. build_ods_bundled.sync_macro_version_tree).
DEBUG_MODULE_FLAG_RE = re.compile(
    r"^(\s*)([A-Z_]*DEBUG|DEBUG_LOG)\s*=\s*True\b",
    re.MULTILINE,
)
DEBUG_DICT_FLAG_RE = re.compile(
    r'^(\s*"[^"]*DEBUG"\s*:\s*)True\b',
    re.MULTILINE,
)


def collect_debug_true_flags(content):
    """Список module-level *_DEBUG / DEBUG_LOG и ключей *DEBUG в dict, равных True."""
    module = [m.group(2) for m in DEBUG_MODULE_FLAG_RE.finditer(content)]
    dict_keys = []
    for m in DEBUG_DICT_FLAG_RE.finditer(content):
        key_m = re.search(r'"([^"]*)"', m.group(0))
        if key_m:
            dict_keys.append(key_m.group(1))
    return module, dict_keys


def disable_debug_flags_in_source(content):
    """Выключить module-level *_DEBUG / DEBUG_LOG и ключи *DEBUG в dict literals."""
    module, dict_keys = collect_debug_true_flags(content)
    new_content = DEBUG_MODULE_FLAG_RE.sub(r"\1\2 = False", content)
    new_content = DEBUG_DICT_FLAG_RE.sub(r"\1False", new_content)
    changed = new_content != content
    return new_content, changed, module, dict_keys


def _read_py_source(path):
    try:
        with open(path, encoding="utf-8") as f:
            return f.read(), "utf-8"
    except UnicodeDecodeError:
        with open(path, encoding="cp1251") as f:
            return f.read(), "cp1251"


def disable_debug_flags_file(path):
    content, encoding = _read_py_source(path)
    new_content, changed, _module, _dict_keys = disable_debug_flags_in_source(content)
    if changed:
        with open(path, "w", encoding=encoding) as f:
            f.write(new_content)
    return changed, _module, _dict_keys


def restore_debug_flags_in_source(content, module_flags, dict_keys):
    """Вернуть в True только те флаги, что были сняты при сборке бандла."""
    new_content = content
    for name in module_flags or ():
        new_content = re.sub(
            rf"^(\s*){re.escape(name)}\s*=\s*False\b",
            rf"\1{name} = True",
            new_content,
            count=1,
            flags=re.MULTILINE,
        )
    for key in dict_keys or ():
        new_content = re.sub(
            rf'^(\s*"{re.escape(key)}"\s*:\s*)False\b',
            r"\1True",
            new_content,
            count=1,
            flags=re.MULTILINE,
        )
    return new_content, new_content != content


def restore_debug_flags_file(path, module_flags, dict_keys):
    if not module_flags and not dict_keys:
        return False
    content, encoding = _read_py_source(path)
    new_content, changed = restore_debug_flags_in_source(
        content, module_flags, dict_keys
    )
    if changed:
        with open(path, "w", encoding=encoding) as f:
            f.write(new_content)
    return changed


def write_debug_flags_snapshot(snapshot_path, macro_lib_dir, files_data):
    payload = {
        "created": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "macro_lib_dir": os.path.abspath(macro_lib_dir),
        "files": files_data,
    }
    with open(snapshot_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.write("\n")


def restore_debug_flags_tree(
    macro_lib_dir=None,
    snapshot_path=None,
    remove_snapshot=False,
):
    snapshot_path = snapshot_path or DEFAULT_DEBUG_SNAPSHOT
    if not os.path.isfile(snapshot_path):
        raise FileNotFoundError(
            f"Снимок DEBUG-флагов не найден: {snapshot_path}"
        )
    with open(snapshot_path, encoding="utf-8") as f:
        data = json.load(f)
    files = data.get("files") or {}
    if not files:
        return []
    base_dir = os.path.abspath(macro_lib_dir or data.get("macro_lib_dir") or "")
    if not base_dir:
        raise ValueError("macro_lib_dir не задан и отсутствует в снимке")
    restored = []
    for rel, flags in sorted(files.items()):
        path = os.path.join(base_dir, rel.replace("/", os.sep))
        if restore_debug_flags_file(
            path,
            flags.get("module") or [],
            flags.get("dict") or [],
        ):
            restored.append(path)
    if remove_snapshot and restored:
        os.remove(snapshot_path)
    return restored


DEBUG_MODULE_FLAG_ANY_RE = re.compile(
    r"^(\s*)([A-Z_]*DEBUG|DEBUG_LOG)\s*=\s*(?:True|False)\b",
    re.MULTILINE,
)
DEBUG_DICT_FLAG_ANY_RE = re.compile(
    r'^(\s*"[^"]*DEBUG"\s*:\s*)(?:True|False)\b',
    re.MULTILINE,
)


def enable_debug_flags_in_source(content):
    """Принудительно включить все module-level *_DEBUG / DEBUG_LOG и ключи *DEBUG в dict."""
    new_content = DEBUG_MODULE_FLAG_ANY_RE.sub(r"\1\2 = True", content)
    new_content = DEBUG_DICT_FLAG_ANY_RE.sub(r"\1True", new_content)
    return new_content, new_content != content


def enable_debug_flags_file(path):
    content, encoding = _read_py_source(path)
    new_content, changed = enable_debug_flags_in_source(content)
    if changed:
        with open(path, "w", encoding=encoding) as f:
            f.write(new_content)
    return changed


def enable_debug_flags_tree(macro_lib_dir, exclude_py=None):
    """Включить DEBUG-флаги во всех .py macro-lib (независимо от текущего значения)."""
    from build_ods_bundled import ZIP_EXCLUDE_PY, _iter_packable_py_files

    macro_lib_dir = os.path.abspath(macro_lib_dir)
    exclude = set(exclude_py or ZIP_EXCLUDE_PY)
    updated = []
    for path in _iter_packable_py_files(macro_lib_dir, exclude):
        if enable_debug_flags_file(path):
            updated.append(path)
    return updated


def _ensure_manifest_script_entries(manifest_data):
    root = ET.fromstring(manifest_data)
    existing = {}
    for entry in root.findall(f"{{{MANIFEST_NS}}}file-entry"):
        path = entry.get(f"{{{MANIFEST_NS}}}full-path")
        if path:
            existing[path] = entry
    for full_path, media_type in MANIFEST_SCRIPT_ENTRIES:
        entry = existing.get(full_path)
        if entry is None:
            entry = ET.SubElement(root, f"{{{MANIFEST_NS}}}file-entry")
            entry.set(f"{{{MANIFEST_NS}}}full-path", full_path)
        entry.set(f"{{{MANIFEST_NS}}}media-type", media_type)
    ET.register_namespace("manifest", MANIFEST_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_manifest_rdf(rdf_data):
    text = rdf_data.decode("utf-8")
    if SCRIPT_FILE in text:
        return rdf_data
    block = (
        '  <rdf:Description rdf:about="Scripts/python/installer.py">\n'
        '    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/>\n'
        "  </rdf:Description>\n"
        '  <rdf:Description rdf:about="">\n'
        '    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" '
        'rdf:resource="Scripts/python/installer.py"/>\n'
        "  </rdf:Description>\n"
    )
    return text.replace("</rdf:RDF>", block + "</rdf:RDF>").encode("utf-8")


def _patch_version_cell(content_data, version):
    """Ячейка B2 (строка ro2, вторая колонка) — версия из macro-lib/version.txt."""
    text = content_data.decode("utf-8")
    ver_xml = xml.sax.saxutils.escape(version.strip())
    pattern = re.compile(
        r"(<table:table-row table:style-name=\"ro2\"[^>]*>"
        r'<table:table-cell\b[^>]*>\s*<text:p>.*?</text:p>\s*</table:table-cell>\s*)'
        r"<table:table-cell\b([^>]*/>)",
        re.DOTALL,
    )

    def _replacer(match):
        prefix, tag_attrs = match.group(1), match.group(2)
        style_match = re.search(r'table:style-name="([^"]*)"', tag_attrs)
        style = style_match.group(1) if style_match else "ce3"
        filled = (
            f'<table:table-cell table:style-name="{style}" office:value-type="string" '
            f'calcext:value-type="string"><text:p>{ver_xml}</text:p></table:table-cell>'
        )
        return prefix + filled

    new_text, count = pattern.subn(_replacer, text, count=1)
    if count:
        return new_text.encode("utf-8")
    return content_data


def _patch_install_button_event(content_data):
    text = content_data.decode("utf-8")
    if "installer.py$install_macros" in text:
        return content_data
    event_block = (
        "<office:event-listeners>"
        '<script:event-listener script:language="ooo:script" '
        'script:event-name="form:performaction" '
        'xlink:href="vnd.sun.star.script:installer.py$install_macros?language=Python&amp;location=document" '
        'xlink:type="simple"/>'
        "</office:event-listeners>"
    )
    pattern = re.compile(
        r'(<form:button\b[^>]*form:name="InstallButton"[^>]*>)(.*?)(</form:button>)',
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

    new_text, count = pattern.subn(_replacer, text, count=1)
    if count:
        return new_text.encode("utf-8")
    return content_data


def build_ods(template_path="template.ods", output_path="MacroInstaller.ods", version=None):
    if not os.path.exists(template_path):
        print(f"❌ Ошибка: Найдите файл '{template_path}' (создайте пустой ODS в LibreOffice и сохраните его под этим именем).")
        return False

    macro_version = _resolve_macro_version(version)

    with zipfile.ZipFile(template_path, 'r') as zin:
        manifest_data = zin.read('META-INF/manifest.xml')

    updated_manifest = _ensure_manifest_script_entries(manifest_data)

    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'META-INF/manifest.xml':
                    data = updated_manifest
                elif item.filename == 'manifest.rdf':
                    data = _patch_manifest_rdf(data)
                elif item.filename == 'content.xml':
                    data = _patch_version_cell(data, macro_version)
                    data = _patch_install_button_event(data)
                if item.filename == 'mimetype':
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)
            zout.writestr('Scripts/', b'')
            zout.writestr('Scripts/python/', b'')
            zout.writestr(SCRIPT_FILE, INSTALLER_PY_CODE.encode('utf-8'))

    print(f"✅ Успешно создан: {os.path.abspath(output_path)} (версия {macro_version})")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Сборка MacroInstaller.ods с встроенным макросом установки",
    )
    parser.add_argument(
        "--template",
        default="template.ods",
        help="Шаблон ODS (по умолчанию: template.ods)",
    )
    parser.add_argument(
        "--output",
        default="MacroInstaller.ods",
        help="Выходной файл (по умолчанию: MacroInstaller.ods)",
    )
    parser.add_argument(
        "--version",
        dest="macro_version",
        metavar="VER",
        help="Версия: записать в macro-lib/version.txt и в ячейку B2 листа",
    )
    args = parser.parse_args()
    ok = build_ods(
        template_path=args.template,
        output_path=args.output,
        version=args.macro_version,
    )
    raise SystemExit(0 if ok else 1)


if __name__ == "__main__":
    main()