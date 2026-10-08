# -*- coding: utf-8 -*-
"""
Собрать standalone file_list_macro.ods:

- Один Scripts/python/file_list_macro.py (cfg+lib+help вшиты) —
  document-pythonpath в AlterOffice/LO для tdoc: не работает.
- лист «Старт» с кнопками «Сбор файлов» и «Справка»
- docs/file_list_user_guide.odt вшивается as-is (править вручную в Writer)

Запуск:

  python3 standalone/file_list_macros/build_file_list_ods.py
"""
from __future__ import print_function

import base64
import os
import re
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
import xml.sax.saxutils

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, "..", ".."))
MACRO_LIB = os.path.join(REPO, "macro-lib")
PYTHONPATH_SRC = os.path.join(MACRO_LIB, "pythonpath")
TEMPLATE_ODS = os.path.join(REPO, "installer", "template.ods")
DOCS_DIR = os.path.join(HERE, "docs")
ODT_GUIDE = os.path.join(DOCS_DIR, "file_list_user_guide.odt")
ODS_NAME = "file_list_macro.ods"
SCRIPT_NAME = "file_list_macro.py"
START_SHEET = "Старт"
MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"

RUN_HREF = (
    "vnd.sun.star.script:file_list_macro.py$analyze_files_dialog"
    "?language=Python&amp;location=document"
)
HELP_HREF = (
    "vnd.sun.star.script:file_list_macro.py$show_file_list_help"
    "?language=Python&amp;location=document"
)

# Имена из cfg → как в entry после `from … import X as _X`
_CFG_ALIASES = (
    ("COLUMN_WIDTH_AUTOFILTER_PAD", "_COLUMN_WIDTH_AUTOFILTER_PAD"),
    ("DEFAULT_COUNT_PAGES", "_DEFAULT_COUNT_PAGES"),
    ("DEFAULT_COUNT_SHA256", "_DEFAULT_COUNT_SHA256"),
    ("DEFAULT_DLP_ANALYZE_CONTENT", "_DEFAULT_DLP_ANALYZE_CONTENT"),
    ("DEFAULT_DLP_ANALYZE_NAMES", "_DEFAULT_DLP_ANALYZE_NAMES"),
    ("DEFAULT_DLP_CONTENT_EXCEL", "_DEFAULT_DLP_CONTENT_EXCEL"),
    ("DEFAULT_DLP_CONTENT_TEXT", "_DEFAULT_DLP_CONTENT_TEXT"),
    ("DEFAULT_DLP_CONTENT_WORD", "_DEFAULT_DLP_CONTENT_WORD"),
    ("DEFAULT_MAX_DEPTH", "_DEFAULT_MAX_DEPTH"),
    ("DEFAULT_RESULT_SHEET_NAME", "_DEFAULT_RESULT_SHEET_NAME"),
    ("DEFAULT_SHA256_MAX_GB", "_DEFAULT_SHA256_MAX_GB"),
    ("DEFAULT_SIZE_UNIT", "_DEFAULT_SIZE_UNIT"),
    ("DEFAULT_SORT_MODE", "_DEFAULT_SORT_MODE"),
    ("DEPTH_INDENT", "_DEPTH_INDENT"),
    ("DLP_CONTENT_MAX_CHARS", "_DLP_CONTENT_MAX_CHARS"),
    ("DLP_LOG_PATTERN_HEADER", "_DLP_LOG_PATTERN_HEADER"),
    ("DLP_LOG_SOURCE_HEADER", "_DLP_LOG_SOURCE_HEADER"),
    ("DLP_SOURCE_CONTENT", "_DLP_SOURCE_CONTENT"),
    ("DLP_SOURCE_NAME", "_DLP_SOURCE_NAME"),
    ("PAGE_COUNT_EXTS", "_PAGE_COUNT_EXTS"),
    ("RESULT_HEADER_BASE", "_RESULT_HEADER_BASE"),
    ("RESULT_HEADER_DUP", "_RESULT_HEADER_DUP"),
    ("RESULT_HEADER_DUP_N", "_RESULT_HEADER_DUP_N"),
    ("RESULT_HEADER_OWNER", "_RESULT_HEADER_OWNER"),
    ("RESULT_HEADER_PAGES", "_RESULT_HEADER_PAGES"),
    ("RESULT_HEADER_SHA256", "_RESULT_HEADER_SHA256"),
    ("RESULT_HEADER_STYLE", "_RESULT_HEADER_STYLE"),
    ("SETTINGS_APP_NAME", "_SETTINGS_APP_NAME"),
    ("SETTINGS_MODULE_NAME", "_SETTINGS_MODULE_NAME"),
    ("SIZE_UNIT_DIVISOR", "_SIZE_UNIT_DIVISOR"),
    ("SIZE_UNIT_OPTIONS", "_SIZE_UNIT_OPTIONS"),
    ("SORT_PRESETS", "_SORT_PRESETS"),
    ("UI_YIELD_EVERY", "FILE_LIST_UI_YIELD_EVERY"),
    ("WRITE_CHUNK", "FILE_LIST_WRITE_CHUNK"),
)


def _esc(text):
    return xml.sax.saxutils.escape(str(text))


def _read_text(path):
    with open(path, "r", encoding="utf-8") as fh:
        return fh.read()


def _read_bytes(path):
    with open(path, "rb") as fh:
        return fh.read()


def _strip_leading_noise(src):
    """Убрать shebang/coding/future/MACRO_VERSION/модульный docstring."""
    lines = src.replace("\r\n", "\n").split("\n")
    out = []
    i = 0
    n = len(lines)
    while i < n:
        s = lines[i].strip()
        if i < 8 and (
            s.startswith("#")
            or s.startswith("from __future__")
            or s.startswith("MACRO_VERSION")
            or s == ""
        ):
            i += 1
            continue
        if s.startswith('"""') or s.startswith("'''"):
            q = s[:3]
            if s.count(q) >= 2 and len(s) > 3:
                i += 1
                continue
            i += 1
            while i < n and q not in lines[i]:
                i += 1
            if i < n:
                i += 1
            continue
        break
    while i < n:
        out.append(lines[i])
        i += 1
    return "\n".join(out)


def _cfg_assignments_for_class(cfg_src):
    """Тело класса _C: присваивания из cfg (без import/future)."""
    text = _strip_leading_noise(cfg_src)
    body = []
    for line in text.split("\n"):
        st = line.strip()
        if not st:
            body.append("")
            continue
        if st.startswith("from ") or st.startswith("import "):
            continue
        if st.startswith("#") and "Мутабельное" in st:
            body.append("    " + st)
            continue
        body.append("    " + line)
    return "\n".join(body)


def _build_cfg_class(cfg_src, odt_bytes):
    b64 = base64.b64encode(odt_bytes).decode("ascii")
    # Каждая строка с отступом — иначе AO-фильтр обрывает class (_C).
    chunks = []
    for i in range(0, len(b64), 76):
        part = b64[i : i + 76]
        chunks.append('        "%s"' % part)
    assigns = _cfg_assignments_for_class(cfg_src)
    return (
        "class _C(object):\n"
        "    # --- из libre_macros_file_list_cfg (AO: только тело class/def живёт) ---\n"
        "%s\n"
        "    HELP_ODT_B64 = (\n"
        "%s\n"
        "    )\n"
        % (assigns, "\n".join(chunks))
    )


def _prepare_lib_source(lib_src):
    text = _strip_leading_noise(lib_src)
    lines = []
    skip_unicode_shim = 0
    for line in text.split("\n"):
        st = line.strip()
        if st.startswith("import libre_macros_file_list_cfg"):
            continue
        if st.startswith("from __future__"):
            continue
        # модульный try/except unicode — AO вырежет; заменим def ниже
        if st == "try:" and skip_unicode_shim == 0:
            # look-ahead handled by state
            pass
        if re.match(r"^try:\s*$", line) and skip_unicode_shim == 0:
            skip_unicode_shim = 1
            continue
        if skip_unicode_shim == 1 and st == "unicode":
            skip_unicode_shim = 2
            continue
        if skip_unicode_shim == 2 and st.startswith("except NameError"):
            skip_unicode_shim = 3
            continue
        if skip_unicode_shim == 3 and ("unicode = str" in st or st.startswith("unicode =")):
            skip_unicode_shim = 0
            continue
        if skip_unicode_shim:
            skip_unicode_shim = 0
        lines.append(line)
    text = "\n".join(lines)
    # _cfg.X и getattr(_cfg, ...) → _C (document bundle, без pythonpath)
    text = re.sub(r"\b_cfg\b", "_C", text)
    helper = (
        "def _py_unicode(value):\n"
        "    try:\n"
        "        return unicode(value)\n"
        "    except NameError:\n"
        "        return str(value)\n\n"
    )
    # вызовы unicode(...) → _py_unicode(...); не трогать слова в комментариях грубо
    text = re.sub(r"\bunicode\(", "_py_unicode(", text)
    return helper + text


_UI_CONST_NAMES = (
    "SOFT_GRAY_FORM_BG",
    "SOFT_GRAY_TITLE_BG",
    "SOFT_GRAY_TITLE_FG",
    "SOFT_GRAY_LABEL_FG",
    "SOFT_GRAY_FIELD_BG",
    "SOFT_GRAY_BTN_BG",
    "SOFT_GRAY_TITLE_H",
    "SOFT_GRAY_BAR_NAME",
    "WARNING_TITLE_BG",
    "WARNING_TITLE_FG",
    "ORANGE_TITLE_BG",
    "ORANGE_TITLE_FG",
    "MANUAL_INPUT_TITLE_BG",
    "MANUAL_INPUT_TITLE_FG",
    "DARK_BLUE_TITLE_BG",
    "DARK_BLUE_TITLE_FG",
    "GREEN_TITLE_BG",
    "GREEN_TITLE_FG",
    "INNER_BTN_W",
    "INNER_BTN_H",
    "INNER_BTN_GAP",
    "INNER_CHK_H",
    "LISTBOX_FONT_HEIGHT",
    "DIALOG_EDIT_FONT_HEIGHT",
    "MANUAL_INPUT_VALUE_FONT_HEIGHT",
    "HELP_TEXT_FONT_HEIGHT",
    "PLUGIN_WARN_TEXT_FONT_HEIGHT",
    "TOKEN_BTN_FONT_HEIGHT",
    "TOKEN_BTN_LABELS",
)


def _prepare_ui_theme_source(ui_src):
    """
    Константы ui_theme → class _Ui (AO сохраняет тело class).
    Функции — на уровне модуля, ссылки на константы через _Ui.*.
    """
    text = _strip_leading_noise(ui_src)
    text = re.sub(
        r"try:\n\s*unicode\nexcept NameError:\n\s*unicode = str\n*",
        "",
        text,
    )
    lines = text.split("\n")
    const_lines = []
    func_lines = []
    in_func = False
    for line in lines:
        if line.startswith("def ") or line.startswith("async def "):
            in_func = True
            func_lines.append(line)
            continue
        if in_func:
            if line.strip() and not (line.startswith(" ") or line.startswith("\t")):
                in_func = False
            else:
                func_lines.append(line)
                continue
        st = line.strip()
        if not st:
            const_lines.append("")
            continue
        if st.startswith("from ") or st.startswith("import "):
            continue
        const_lines.append("    " + line)

    func_text = "\n".join(func_lines)
    for name in sorted(_UI_CONST_NAMES, key=len, reverse=True):
        func_text = re.sub(r"\b%s\b" % re.escape(name), "_Ui.%s" % name, func_text)
    func_text = re.sub(r"\bunicode\(", "_py_unicode(", func_text)

    class_body = "\n".join(const_lines).rstrip() + "\n"
    return (
        "class _Ui(object):\n"
        "    # --- libre_macros_ui_theme constants ---\n"
        "%s\n"
        "%s\n"
        % (class_body, func_text)
    )


def _prepare_entry_source(entry_src):
    text = entry_src
    # убрать import cfg/lib
    text = re.sub(
        r"^from libre_macros_file_list_cfg import[^\n]*\n",
        "",
        text,
        flags=re.M,
    )
    text = re.sub(
        r"^import libre_macros_file_list_cfg as _fl_state\n",
        "",
        text,
        flags=re.M,
    )
    text = re.sub(
        r"^import libre_macros_file_list_lib as _fl_lib\n",
        "",
        text,
        flags=re.M,
    )
    text = re.sub(
        r"\ng_exportedScripts\s*=\s*\([^)]*\)\s*\n?",
        "\n",
        text,
        flags=re.M,
    )
    # alias → _C.
    for cfg_name, alias in _CFG_ALIASES:
        text = re.sub(r"\b%s\b" % re.escape(alias), "_C.%s" % cfg_name, text)
    text = text.replace("_fl_state.", "_C.")
    text = text.replace("_fl_lib.", "")
    # ui imports inside functions — функции уже в бандле; константы из _Ui
    text = text.replace(
        "from libre_macros_ui_theme import apply_soft_gray_green_theme\n",
        "pass  # apply_soft_gray_green_theme bundled\n",
    )
    text = text.replace(
        "from libre_macros_ui_theme import GREEN_TITLE_BG, GREEN_TITLE_FG, try_paint_titlebar\n",
        "GREEN_TITLE_BG, GREEN_TITLE_FG = _Ui.GREEN_TITLE_BG, _Ui.GREEN_TITLE_FG\n",
    )
    # help append
    help_fn = '''
def write_file_list_help_odt(path):
    import base64 as _b64
    data = _b64.b64decode(_C.HELP_ODT_B64.encode("ascii"))
    path = os.path.abspath(str(path))
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    with open(path, "wb") as handle:
        handle.write(data)
    return path


def show_file_list_help(*args):
    """Открыть вшитую справку (ODT) во Writer / Atext."""
    try:
        import tempfile as _tempfile
        path = os.path.join(_tempfile.gettempdir(), "file_list_user_guide.odt")
        write_file_list_help_odt(path)
    except Exception as exc:
        show_message("Не удалось подготовить файл справки:\\n%s" % exc)
        return False
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        desktop.loadComponentFromURL(url, "_blank", 0, ())
        return True
    except Exception as exc:
        show_message("Не удалось открыть справку:\\n%s" % exc)
        return False


g_exportedScripts = (get_file_list, analyze_files_dialog, show_file_list_help)
'''
    if "def show_file_list_help" not in text:
        text = text.rstrip() + "\n" + help_fn
    return text


def _flatten_source(text):
    flatten = os.path.join(MACRO_LIB, "aoffice_flatten_defs.py")
    if not os.path.isfile(flatten):
        return text
    fd, tmp = tempfile.mkstemp(prefix="file_list_entry_", suffix=".py")
    os.close(fd)
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(text)
        import subprocess

        subprocess.check_call([sys.executable, flatten, "--in-place", tmp])
        return _read_text(tmp)
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


def prepare_entry_script(odt_bytes):
    """Один AO-совместимый файл: class _C + lib defs + entry defs."""
    cfg_path = os.path.join(PYTHONPATH_SRC, "libre_macros_file_list_cfg.py")
    lib_path = os.path.join(PYTHONPATH_SRC, "libre_macros_file_list_lib.py")
    entry_path = os.path.join(MACRO_LIB, SCRIPT_NAME)
    for p in (cfg_path, lib_path, entry_path):
        if not os.path.isfile(p):
            raise IOError("Нет исходника: %s" % p)

    ui_path = os.path.join(PYTHONPATH_SRC, "libre_macros_ui_theme.py")
    if not os.path.isfile(ui_path):
        raise IOError("Нет исходника: %s" % ui_path)

    cfg_class = _build_cfg_class(_read_text(cfg_path), odt_bytes)
    ui_body = _prepare_ui_theme_source(_read_text(ui_path))
    lib_body = _prepare_lib_source(_read_text(lib_path))
    entry_body = _prepare_entry_source(_read_text(entry_path))
    # убрать coding/MACRO из entry — оставим один заголовок
    entry_body = re.sub(r"^# -\*- coding: utf-8 -\*-\n", "", entry_body)
    entry_body = re.sub(r'^MACRO_VERSION = "[^"]*"\n', "", entry_body)

    # _py_unicode нужен и ui_theme, и lib — выносим в начало
    unicode_helper = (
        "def _py_unicode(value):\n"
        "    try:\n"
        "        return unicode(value)\n"
        "    except NameError:\n"
        "        return str(value)\n"
    )
    # убрать дубликат helper из lib
    lib_body = re.sub(
        r"def _py_unicode\(value\):\n"
        r"    try:\n"
        r"        return unicode\(value\)\n"
        r"    except NameError:\n"
        r"        return str\(value\)\n\n?",
        "",
        lib_body,
        count=1,
    )

    parts = [
        "# -*- coding: utf-8 -*-",
        '"""Standalone file_list_macro (один файл для AlterOffice document scripts)."""',
        "",
        unicode_helper,
        "",
        cfg_class,
        "",
        "# --- libre_macros_ui_theme ---",
        ui_body,
        "",
        "# --- libre_macros_file_list_lib ---",
        lib_body,
        "",
        "# --- file_list_macro entry ---",
        entry_body,
        "",
    ]
    text = "\n".join(parts)
    text = _flatten_source(text)
    return text


def _p(text):
    return "<text:p>%s</text:p>" % _esc(text)


def _str_cell(style, text):
    return (
        '<table:table-cell table:style-name="%s" office:value-type="string" '
        'calcext:value-type="string">%s</table:table-cell>'
        % (style, _p(text))
    )


def _empty_cell(style=None):
    if style:
        return '<table:table-cell table:style-name="%s"/>' % style
    return "<table:table-cell/>"


def _content_xml():
    rows = []
    rows.append('<table:table-row table:style-name="roHead">')
    rows.append(_str_cell("ceHead", "Сбор данных о файлах и папках"))
    rows.append(
        '<table:table-cell table:style-name="ceBtn">'
        '<draw:control draw:z-index="0" draw:name="RunButtonShape" '
        'draw:style-name="grRun" draw:text-style-name="PBtn" '
        'svg:width="42mm" svg:height="9mm" svg:x="2mm" svg:y="1mm" '
        'draw:control="control1"/>'
        "</table:table-cell>"
    )
    rows.append(
        '<table:table-cell table:style-name="ceBtn">'
        '<draw:control draw:z-index="1" draw:name="HelpButtonShape" '
        'draw:style-name="grHelp" draw:text-style-name="PBtn" '
        'svg:width="32mm" svg:height="9mm" svg:x="2mm" svg:y="1mm" '
        'draw:control="control2"/>'
        "</table:table-cell>"
    )
    rows.append("</table:table-row>")

    notes = [
        "Кнопка «Сбор файлов» открывает диалог обхода каталога.",
        "Результат пишется на выбранный лист; при DLP — ещё лист «*_DLP».",
        "«Справка» открывает вшитый ODT (тот же, что в папке docs/).",
        "На Windows добавляется колонка Owner. Мини-DLP — вкладка DLP в диалоге.",
        "Пересборка книги: python3 standalone/file_list_macros/build_file_list_ods.py",
    ]
    for note in notes:
        rows.append('<table:table-row table:style-name="roData">')
        rows.append(_str_cell("ceNote", note))
        rows.append(_empty_cell())
        rows.append(_empty_cell())
        rows.append("</table:table-row>")

    ns = (
        'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" '
        'xmlns:css3t="http://www.w3.org/TR/css3-text/" '
        'xmlns:grddl="http://www.w3.org/2003/g/data-view#" '
        'xmlns:xhtml="http://www.w3.org/1999/xhtml" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
        'xmlns:xforms="http://www.w3.org/2002/xforms" '
        'xmlns:dom="http://www.w3.org/2001/xml-events" '
        'xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" '
        'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" '
        'xmlns:math="http://www.w3.org/1998/Math/MathML" '
        'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:ooo="http://openoffice.org/2004/office" '
        'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
        'xmlns:ooow="http://openoffice.org/2004/writer" '
        'xmlns:xlink="http://www.w3.org/1999/xlink" '
        'xmlns:drawooo="http://openoffice.org/2010/draw" '
        'xmlns:oooc="http://openoffice.org/2004/calc" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" '
        'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
        'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" '
        'xmlns:tableooo="http://openoffice.org/2009/table" '
        'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
        'xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" '
        'xmlns:rpt="http://openoffice.org/2005/report" '
        'xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" '
        'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" '
        'xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" '
        'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
        'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" '
        'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" '
        'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" '
        'xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" '
        'office:version="1.3"'
    )
    styles = """
<office:automatic-styles>
  <style:style style:name="coA" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="150mm"/>
  </style:style>
  <style:style style:name="coB" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="48mm"/>
  </style:style>
  <style:style style:name="coC" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="38mm"/>
  </style:style>
  <style:style style:name="roHead" style:family="table-row">
    <style:table-row-properties style:row-height="12mm" fo:break-before="auto" style:use-optimal-row-height="false"/>
  </style:style>
  <style:style style:name="roData" style:family="table-row">
    <style:table-row-properties style:row-height="8mm" fo:break-before="auto" style:use-optimal-row-height="false"/>
  </style:style>
  <style:style style:name="ta1" style:family="table" style:master-page-name="Default">
    <style:table-properties table:display="true" style:writing-mode="lr-tb"/>
  </style:style>
  <style:style style:name="ceHead" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties fo:background-color="#2E7D32" style:text-align-source="fix" fo:wrap-option="wrap" style:vertical-align="middle"/>
    <style:paragraph-properties fo:text-align="left"/>
    <style:text-properties fo:color="#ffffff" fo:font-size="14pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="ceBtn" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties style:vertical-align="middle"/>
  </style:style>
  <style:style style:name="ceNote" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties style:text-align-source="fix" fo:wrap-option="wrap" style:vertical-align="middle"/>
    <style:paragraph-properties fo:text-align="left"/>
    <style:text-properties fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="grRun" style:family="graphic" style:parent-style-name="Default">
    <style:graphic-properties fo:background-color="#2E7D32"/>
    <style:text-properties fo:color="#ffffff" fo:font-size="11pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="grHelp" style:family="graphic" style:parent-style-name="Default">
    <style:graphic-properties fo:background-color="#546E7A"/>
    <style:text-properties fo:color="#ffffff" fo:font-size="11pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="PBtn" style:family="paragraph">
    <style:paragraph-properties fo:text-align="center"/>
    <style:text-properties fo:color="#ffffff" fo:font-size="11pt" fo:font-weight="bold"/>
  </style:style>
</office:automatic-styles>
"""
    form = (
        '<office:forms form:automatic-focus="false" form:apply-design-mode="false">'
        '<form:form form:name="Form" form:apply-filter="true" form:command-type="table" '
        'form:control-implementation="ooo:com.sun.star.form.component.Form" office:target-frame="">'
        "<form:properties>"
        '<form:property form:property-name="PropertyChangeNotificationEnabled" '
        'office:value-type="boolean" office:boolean-value="true"/>'
        "</form:properties>"
        '<form:button form:name="RunButton" '
        'form:control-implementation="ooo:com.sun.star.form.component.CommandButton" '
        'xml:id="control1" form:id="control1" form:label="Сбор файлов" '
        'form:printable="false" office:target-frame="" form:default-button="true" '
        'form:delay-for-repeat="PT0.050000000S" form:image-position="center">'
        "<form:properties>"
        '<form:property form:property-name="DefaultControl" office:value-type="string" '
        'office:string-value="com.sun.star.form.control.CommandButton"/>'
        "</form:properties>"
        "<office:event-listeners>"
        '<script:event-listener script:language="ooo:script" '
        'script:event-name="form:performaction" '
        'xlink:href="%s" xlink:type="simple"/>'
        "</office:event-listeners>"
        "</form:button>"
        '<form:button form:name="HelpButton" '
        'form:control-implementation="ooo:com.sun.star.form.component.CommandButton" '
        'xml:id="control2" form:id="control2" form:label="Справка" '
        'form:printable="false" office:target-frame="" '
        'form:delay-for-repeat="PT0.050000000S" form:image-position="center">'
        "<form:properties>"
        '<form:property form:property-name="DefaultControl" office:value-type="string" '
        'office:string-value="com.sun.star.form.control.CommandButton"/>'
        "</form:properties>"
        "<office:event-listeners>"
        '<script:event-listener script:language="ooo:script" '
        'script:event-name="form:performaction" '
        'xlink:href="%s" xlink:type="simple"/>'
        "</office:event-listeners>"
        "</form:button>"
        "</form:form></office:forms>"
        % (RUN_HREF, HELP_HREF)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        "<office:document-content %s>"
        "<office:scripts/>"
        "<office:font-face-decls>"
        '<style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" '
        'style:font-family-generic="swiss" style:font-pitch="variable"/>'
        '<style:font-face style:name="PT Sans" svg:font-family="&apos;PT Sans&apos;" '
        'style:font-family-generic="swiss" style:font-pitch="variable"/>'
        "</office:font-face-decls>"
        "%s"
        "<office:body><office:spreadsheet>"
        '<table:calculation-settings table:automatic-find-labels="false" '
        'table:use-regular-expressions="false" table:use-wildcards="true"/>'
        '<table:table table:name="%s" table:style-name="ta1">'
        "%s"
        '<table:table-column table:style-name="coA"/>'
        '<table:table-column table:style-name="coB"/>'
        '<table:table-column table:style-name="coC"/>'
        "%s"
        "</table:table>"
        "<table:named-expressions/>"
        "</office:spreadsheet></office:body></office:document-content>"
        % (ns, styles, START_SHEET, form, "".join(rows))
    )


def _meta_xml():
    return """<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 office:version="1.3">
 <office:meta>
  <dc:title>Сбор файлов и папок</dc:title>
  <dc:description>Standalone file_list_macro: обход каталога, DLP по именам, Owner на Windows.</dc:description>
  <meta:generator>libre-macros/standalone/file_list_macros/build_file_list_ods.py</meta:generator>
 </office:meta>
</office:document-meta>
"""


def _zip_write(zout, name, data, stored=False):
    if not isinstance(data, bytes):
        data = data.encode("utf-8")
    info = zipfile.ZipInfo(name)
    info.compress_type = zipfile.ZIP_STORED if stored else zipfile.ZIP_DEFLATED
    zout.writestr(info, data)


def _ensure_manifest_entries(manifest_data, file_paths):
    root = ET.fromstring(manifest_data)
    existing = {}
    for entry in root.findall("{%s}file-entry" % MANIFEST_NS):
        path = entry.get("{%s}full-path" % MANIFEST_NS)
        if path:
            existing[path] = entry
    for full_path, media_type in file_paths:
        entry = existing.get(full_path)
        if entry is None:
            entry = ET.SubElement(root, "{%s}file-entry" % MANIFEST_NS)
            entry.set("{%s}full-path" % MANIFEST_NS, full_path)
        entry.set("{%s}media-type" % MANIFEST_NS, media_type)
    ET.register_namespace("manifest", MANIFEST_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_manifest_rdf(rdf_data, parts):
    text = rdf_data.decode("utf-8")
    blocks = []
    for part in parts:
        if part in text:
            continue
        blocks.append(
            '  <rdf:Description rdf:about="%s">\n'
            '    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/>\n'
            "  </rdf:Description>\n"
            '  <rdf:Description rdf:about="">\n'
            '    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" '
            'rdf:resource="%s"/>\n'
            "  </rdf:Description>\n"
            % (part, part)
        )
    if not blocks:
        return rdf_data
    return text.replace("</rdf:RDF>", "".join(blocks) + "</rdf:RDF>").encode("utf-8")


def build_ods(output_path=None):
    output_path = output_path or os.path.join(HERE, ODS_NAME)
    if not os.path.isfile(TEMPLATE_ODS):
        raise IOError("Не найден шаблон ODS: %s" % TEMPLATE_ODS)
    if not os.path.isfile(ODT_GUIDE):
        raise IOError("Нет файла справки ODT: %s" % ODT_GUIDE)

    odt_bytes = _read_bytes(ODT_GUIDE)

    entry_src = prepare_entry_script(odt_bytes)
    script_rel = "Scripts/python/" + SCRIPT_NAME

    staged = os.path.join(HERE, "_staged_python")
    if os.path.isdir(staged):
        shutil.rmtree(staged)
    os.makedirs(staged)
    with open(os.path.join(staged, SCRIPT_NAME), "w", encoding="utf-8") as fh:
        fh.write(entry_src)

    manifest_extra = [
        ("Scripts/", "application/binary"),
        ("Scripts/python/", "application/binary"),
        (script_rel, "application/binary"),
    ]
    rdf_parts = [script_rel]

    content = _content_xml().encode("utf-8")
    with zipfile.ZipFile(TEMPLATE_ODS, "r") as zin:
        manifest = _ensure_manifest_entries(
            zin.read("META-INF/manifest.xml"), manifest_extra
        )
        with zipfile.ZipFile(output_path, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "content.xml":
                    data = content
                elif item.filename == "META-INF/manifest.xml":
                    data = manifest
                elif item.filename == "manifest.rdf":
                    data = _patch_manifest_rdf(data, rdf_parts)
                elif item.filename == "meta.xml":
                    data = _meta_xml().encode("utf-8")
                elif item.filename == "settings.xml":
                    data = data.replace(
                        b"Installer_Libre_Macro", START_SHEET.encode("utf-8")
                    )
                if item.filename == "mimetype":
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)
            _zip_write(zout, "Scripts/", b"")
            _zip_write(zout, "Scripts/python/", b"")
            _zip_write(zout, script_rel, entry_src.encode("utf-8"))
    return output_path, len(odt_bytes), 1


def main():
    path, odt_len, n_py = build_ods()
    print("OK:", path)
    print("docs odt (embedded as-is):", ODT_GUIDE, "(%d bytes)" % odt_len)
    print("python modules in ODS:", n_py, "(single-file bundle)")
    print("staged copy:", os.path.join(HERE, "_staged_python"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
