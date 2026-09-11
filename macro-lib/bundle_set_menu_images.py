from __future__ import print_function
MACRO_VERSION = "3.10.724"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пересобрать pythonpath/libre_macros_set_menu_images_blob.py из set_menu_assets/.

Источники:
  set_menu_assets/lc_imagelist.xml
  set_menu_assets/Bitmaps/lc_userimages.png

XML и PNG оба в blob (не в cfg) — однострочный литерал с \\n ломается при правках.
В cfg остаются только имена файлов.
"""
import base64
import re
import sys
from pathlib import Path

MACRO_DIR = Path(__file__).resolve().parent
ASSETS = MACRO_DIR / "set_menu_assets"
XML_SRC = ASSETS / "lc_imagelist.xml"
PNG_SRC = ASSETS / "Bitmaps" / "lc_userimages.png"
BLOB_OUT = MACRO_DIR / "pythonpath" / "libre_macros_set_menu_images_blob.py"
CFG_OUT = MACRO_DIR / "pythonpath" / "libre_macros_set_menu_cfg.py"
VERSION_TXT = MACRO_DIR / "version.txt"


def _read_version():
    try:
        return VERSION_TXT.read_text(encoding="utf-8").strip().splitlines()[0].strip()
    except Exception:
        return "0"


def _triple_quote_xml(text):
    """XML → литерал '''...''' (без ''' внутри)."""
    s = text.replace("\r\n", "\n").replace("\r", "\n")
    if "'''" in s:
        raise SystemExit("XML содержит ''' — нельзя triple-single")
    if not s.endswith("\n"):
        s = s + "\n"
    return "'''%s'''" % s


def _write_blob(xml_text, png_bytes, version):
    b64 = base64.b64encode(png_bytes).decode("ascii")
    wrapped = "\n".join(b64[i : i + 76] for i in range(0, len(b64), 76))
    xml_lit = _triple_quote_xml(xml_text)
    BLOB_OUT.write_text(
        "# -*- coding: utf-8 -*-\n"
        '"""Вшитые lc_imagelist.xml + lc_userimages.png для set_menu.\n'
        "Пересборка: python3 macro-lib/bundle_set_menu_images.py\n"
        '"""\n'
        "from __future__ import print_function, unicode_literals\n"
        'MACRO_VERSION = "%s"\n\n'
        "import base64\n"
        "import os\n\n"
        'LC_IMAGELIST_FILENAME = "lc_imagelist.xml"\n'
        'LC_USERIMAGES_PNG_FILENAME = "lc_userimages.png"\n'
        'LC_USERIMAGES_PNG_NAME = LC_USERIMAGES_PNG_FILENAME\n\n'
        "LC_IMAGELIST_XML = %s\n\n"
        "# binary length %d\n"
        "LC_USERIMAGES_PNG_B64 = (\n"
        '"""%s\n"""\n'
        ")\n\n\n"
        "def lc_imagelist_xml_text():\n"
        "    return LC_IMAGELIST_XML\n\n\n"
        "def lc_userimages_png_bytes():\n"
        '    return base64.b64decode(LC_USERIMAGES_PNG_B64.encode("ascii"))\n\n\n'
        "def write_lc_imagelist_xml(path):\n"
        "    text = lc_imagelist_xml_text()\n"
        "    path = os.path.abspath(str(path))\n"
        "    parent = os.path.dirname(path)\n"
        "    if parent and not os.path.isdir(parent):\n"
        "        os.makedirs(parent)\n"
        '    with open(path, "w", encoding="utf-8", newline="\\n") as f:\n'
        "        f.write(text)\n"
        "        if not text.endswith(\"\\n\"):\n"
        "            f.write(\"\\n\")\n"
        "    return path\n\n\n"
        "def write_lc_userimages_png(path):\n"
        "    data = lc_userimages_png_bytes()\n"
        "    path = os.path.abspath(str(path))\n"
        "    parent = os.path.dirname(path)\n"
        "    if parent and not os.path.isdir(parent):\n"
        "        os.makedirs(parent)\n"
        '    with open(path, "wb") as f:\n'
        "        f.write(data)\n"
        "    return path\n"
        % (version, xml_lit, len(png_bytes), wrapped),
        encoding="utf-8",
    )


def _patch_cfg_remove_xml():
    """В cfg оставить имена файлов; убрать сломанный/дублирующий LC_IMAGELIST_XML."""
    if not CFG_OUT.is_file():
        raise SystemExit("нет cfg: %s" % CFG_OUT)
    cfg = CFG_OUT.read_text(encoding="utf-8")
    # вырезать старый многострочный/однострочный LC_IMAGELIST_XML = ...
    cfg2, n = re.subn(
        r"^LC_IMAGELIST_XML\s*=\s*(?:'''[\s\S]*?'''|'[^']*'|\"[^\"]*\")\s*\n?",
        "",
        cfg,
        count=1,
        flags=re.M,
    )
    if n:
        cfg = cfg2
    else:
        # сломанный вариант: открывающая кавычка и сырой XML до одиночной "'"
        cfg2, n = re.subn(
            r"^LC_IMAGELIST_XML\s*=\s*'[^']*$\n(?:.*\n)*?^'\s*\n?",
            "",
            cfg,
            count=1,
            flags=re.M,
        )
        if n:
            cfg = cfg2
    # гарантировать имена файлов после TOOLBAR_RESOURCE
    if "LC_IMAGELIST_FILENAME" not in cfg:
        m = re.search(r"^TOOLBAR_RESOURCE\s*=.*\n", cfg, flags=re.M)
        if not m:
            raise SystemExit("не найден TOOLBAR_RESOURCE в cfg")
        block = (
            "\n"
            'LC_IMAGELIST_FILENAME = "lc_imagelist.xml"\n'
            'LC_USERIMAGES_PNG_FILENAME = "lc_userimages.png"\n'
        )
        cfg = cfg[: m.end()] + block + cfg[m.end() :]
    CFG_OUT.write_text(cfg, encoding="utf-8")


def main():
    if not XML_SRC.is_file():
        print("Нет файла:", XML_SRC, file=sys.stderr)
        return 1
    if not PNG_SRC.is_file():
        print("Нет файла:", PNG_SRC, file=sys.stderr)
        return 1
    version = _read_version()
    xml_text = XML_SRC.read_text(encoding="utf-8")
    png_bytes = PNG_SRC.read_bytes()
    _write_blob(xml_text, png_bytes, version)
    _patch_cfg_remove_xml()
    # проверка синтаксиса
    import ast

    ast.parse(BLOB_OUT.read_text(encoding="utf-8"))
    ast.parse(CFG_OUT.read_text(encoding="utf-8"))
    print("OK:", BLOB_OUT, "(xml+png)")
    print("OK:", CFG_OUT, "(без LC_IMAGELIST_XML)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
