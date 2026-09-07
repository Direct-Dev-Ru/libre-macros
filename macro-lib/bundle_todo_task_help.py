from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.690"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Собрать pythonpath/todo_task_help_blob.py из todo_task_help_ru.md."""
import base64
import io
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

try:
    unicode
except NameError:
    unicode = str

def _macro_dir():
    return Path(__file__).resolve().parent


def _src_path():
    return _macro_dir() / "todo_task_help_ru.md"


def _out_path():
    return _macro_dir() / "pythonpath" / "todo_task_help_blob.py"


def _version():
    path = _macro_dir() / "version.txt"
    try:
        return path.read_text(encoding="utf-8").strip() or "0.0.0"
    except Exception:
        return "0.0.0"


def _inline_xml(text):
    """Простой **жирный** внутри абзаца."""
    s = unicode(text or u"")
    out = []
    i = 0
    n = len(s)
    while i < n:
        if s.startswith("**", i):
            j = s.find("**", i + 2)
            if j > i:
                out.append(
                    '<text:span text:style-name="TBold">%s</text:span>'
                    % escape(s[i + 2 : j])
                )
                i = j + 2
                continue
        out.append(escape(s[i]))
        i += 1
    return "".join(out)


def _parse_blocks(md):
    blocks = []
    for raw in md.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw.rstrip()
        if not line.strip():
            blocks.append(("blank", ""))
            continue
        if line.startswith("### "):
            blocks.append(("h3", line[4:].strip()))
        elif line.startswith("## "):
            blocks.append(("h2", line[3:].strip()))
        elif line.startswith("# "):
            blocks.append(("h1", line[2:].strip()))
        elif line.startswith("- "):
            blocks.append(("li", line[2:].strip()))
        else:
            blocks.append(("p", line.strip()))
    return blocks


def _content_xml(blocks):
    paras = []
    first_h1 = True
    for kind, text in blocks:
        if kind == "blank":
            paras.append('<text:p text:style-name="PGap"/>')
            continue
        body = _inline_xml(text)
        if kind == "h1":
            style = "PTitle" if first_h1 else "PH1"
            first_h1 = False
            paras.append('<text:p text:style-name="%s">%s</text:p>' % (style, body))
        elif kind == "h2":
            paras.append('<text:p text:style-name="PH2">%s</text:p>' % body)
        elif kind == "h3":
            paras.append('<text:p text:style-name="PH3">%s</text:p>' % body)
        elif kind == "li":
            paras.append(
                '<text:p text:style-name="PBullet"><text:s text:c="1"/>•  %s</text:p>'
                % body
            )
        else:
            paras.append('<text:p text:style-name="PBody">%s</text:p>' % body)
    inner = "\n".join("      " + p for p in paras)
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        "<office:document-content",
        ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
        ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
        ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
        ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
        ' office:version="1.2">',
        "  <office:automatic-styles>",
        '    <style:style style:name="PTitle" style:family="paragraph">',
        '      <style:paragraph-properties fo:margin-bottom="0.35cm" fo:margin-top="0cm"/>',
        '      <style:text-properties fo:font-size="20pt" fo:font-weight="bold" fo:color="#0D47A1"/>',
        "    </style:style>",
        '    <style:style style:name="PH1" style:family="paragraph">',
        '      <style:paragraph-properties fo:margin-top="0.55cm" fo:margin-bottom="0.18cm"/>',
        '      <style:text-properties fo:font-size="16pt" fo:font-weight="bold" fo:color="#0D47A1"/>',
        "    </style:style>",
        '    <style:style style:name="PH2" style:family="paragraph">',
        '      <style:paragraph-properties fo:margin-top="0.38cm" fo:margin-bottom="0.12cm"/>',
        '      <style:text-properties fo:font-size="13pt" fo:font-weight="bold" fo:color="#1565C0"/>',
        "    </style:style>",
        '    <style:style style:name="PH3" style:family="paragraph">',
        '      <style:paragraph-properties fo:margin-top="0.28cm" fo:margin-bottom="0.08cm"/>',
        '      <style:text-properties fo:font-size="12pt" fo:font-weight="bold" fo:color="#1976D2"/>',
        "    </style:style>",
        '    <style:style style:name="PBody" style:family="paragraph">',
        '      <style:paragraph-properties fo:margin-bottom="0.12cm" fo:line-height="120%"/>',
        '      <style:text-properties fo:font-size="11pt"/>',
        "    </style:style>",
        '    <style:style style:name="PBullet" style:family="paragraph">',
        '      <style:paragraph-properties fo:margin-left="0.45cm" fo:margin-bottom="0.06cm" fo:line-height="120%"/>',
        '      <style:text-properties fo:font-size="11pt"/>',
        "    </style:style>",
        '    <style:style style:name="PGap" style:family="paragraph">',
        '      <style:paragraph-properties fo:margin-bottom="0.08cm"/>',
        '      <style:text-properties fo:font-size="6pt"/>',
        "    </style:style>",
        '    <style:style style:name="TBold" style:family="text">',
        '      <style:text-properties fo:font-weight="bold"/>',
        "    </style:style>",
        "  </office:automatic-styles>",
        "  <office:body>",
        "    <office:text>",
        inner,
        "    </office:text>",
        "  </office:body>",
        "</office:document-content>",
        "",
    ]
    return "\n".join(parts)


def _styles_xml():
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        "<office:document-styles",
        ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
        ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
        ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
        ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
        ' office:version="1.2">',
        "  <office:styles>",
        '    <style:default-style style:family="paragraph">',
        '      <style:paragraph-properties fo:line-height="120%"/>',
        '      <style:text-properties fo:font-size="11pt" style:font-name="Liberation Serif"',
        '       fo:language="ru" fo:country="RU"/>',
        "    </style:default-style>",
        "  </office:styles>",
        "  <office:automatic-styles>",
        '    <style:page-layout style:name="pm1">',
        '      <style:page-layout-properties fo:page-width="21.0cm" fo:page-height="29.7cm"',
        '       fo:margin-top="1.6cm" fo:margin-bottom="1.6cm"',
        '       fo:margin-left="1.8cm" fo:margin-right="1.8cm"/>',
        "    </style:page-layout>",
        "  </office:automatic-styles>",
        "  <office:master-styles>",
        '    <style:master-page style:name="Standard" style:page-layout-name="pm1"/>',
        "  </office:master-styles>",
        "</office:document-styles>",
        "",
    ]
    return "\n".join(parts)


def _meta_xml():
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        "<office:document-meta",
        ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
        ' xmlns:dc="http://purl.org/dc/elements/1.1/"',
        ' xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
        ' office:version="1.2">',
        "  <office:meta>",
        "    <dc:title>Справка по макросам задач</dc:title>",
        "    <meta:generator>libre-macros todo_task_help</meta:generator>",
        "  </office:meta>",
        "</office:document-meta>",
        "",
    ]
    return "\n".join(parts)


def _settings_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        "<office:document-settings"
        ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
        ' office:version="1.2"><office:settings/></office:document-settings>'
    )


def _manifest_xml():
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">',
        '  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>',
        '  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>',
        '  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>',
        '  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>',
        '  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>',
        "</manifest:manifest>",
        "",
    ]
    return "\n".join(parts)


def build_odt_bytes(md_text):
    blocks = _parse_blocks(md_text)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        info = zipfile.ZipInfo("mimetype")
        info.compress_type = zipfile.ZIP_STORED
        z.writestr(info, "application/vnd.oasis.opendocument.text")
        z.writestr("content.xml", _content_xml(blocks).encode("utf-8"))
        z.writestr("styles.xml", _styles_xml().encode("utf-8"))
        z.writestr("meta.xml", _meta_xml().encode("utf-8"))
        z.writestr("settings.xml", _settings_xml().encode("utf-8"))
        z.writestr("META-INF/manifest.xml", _manifest_xml().encode("utf-8"))
    return buf.getvalue()


def write_blob_module(data, version):
    b64 = base64.b64encode(data).decode("ascii")
    wrapped = "\n".join(b64[i : i + 76] for i in range(0, len(b64), 76))
    _out_path().write_text(
        "# -*- coding: utf-8 -*-\n"
        '"""Вшитый todo_task_help_ru.odt (base64). Пересборка: python3 macro-lib/bundle_todo_task_help.py"""\n'
        "from __future__ import print_function, unicode_literals\n"
        'MACRO_VERSION = "%s"\n'
        "import base64\n"
        "import os\n\n"
        'TODO_TASK_HELP_ODT_NAME = "todo_task_help_ru.odt"\n\n'
        "# binary length %d\n"
        "TODO_TASK_HELP_ODT_B64 = (\n"
        '"""%s\n"""\n'
        ")\n\n\n"
        "def todo_task_help_odt_bytes():\n"
        '    return base64.b64decode(TODO_TASK_HELP_ODT_B64.encode("ascii"))\n\n\n'
        "def write_todo_task_help_odt(path):\n"
        "    data = todo_task_help_odt_bytes()\n"
        "    path = os.path.abspath(str(path))\n"
        "    parent = os.path.dirname(path)\n"
        "    if parent and not os.path.isdir(parent):\n"
        "        os.makedirs(parent)\n"
        '    with open(path, "wb") as f:\n'
        "        f.write(data)\n"
        "    return path\n"
        % (version, len(data), wrapped),
        encoding="utf-8",
    )


def main():
    src = _src_path()
    if not src.is_file():
        print("Нет файла:", src, file=sys.stderr)
        return 1
    md = src.read_text(encoding="utf-8")
    data = build_odt_bytes(md)
    write_blob_module(data, _version())
    print("OK:", _out_path(), "(%d bytes)" % len(data))
    return 0


if __name__ == "__main__":
    sys.exit(main())
