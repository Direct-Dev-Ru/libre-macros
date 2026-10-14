#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Источник + шаблоны AText для сценария 20_atext_letters.

  sources/atext_letters/source_recipients.xlsx
  sources/atext_letters/template_letter.odt   — <<ФИО>> / <<Адрес>> / <<Сумма>>
  sources/atext_letters/template_pair.odt     — <<ФИО>>{1}/<<ФИО>>{2} …

Сценарий: 20_atext_letters.xltx
"""
from __future__ import print_function, unicode_literals

import os
import zipfile

try:
    unicode
except NameError:
    unicode = str  # noqa: A001

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.join(ROOT, "test", "collect_workbooks", "sources", "atext_letters")
OUT_XLSX = os.path.join(OUT_DIR, "source_recipients.xlsx")
TPL_LETTER = os.path.join(OUT_DIR, "template_letter.odt")
TPL_PAIR = os.path.join(OUT_DIR, "template_pair.odt")

_MACRO_PYTHONPATH = os.path.join(ROOT, "macro-lib", "pythonpath")


def _wb():
    try:
        import openpyxl
        from openpyxl import Workbook
        return Workbook
    except ImportError:
        import sys

        if _MACRO_PYTHONPATH not in sys.path:
            sys.path.insert(0, _MACRO_PYTHONPATH)
        import openpyxl_bundled  # noqa: F401
        from openpyxl import Workbook

        return Workbook


def _odt_content(paragraphs):
    """Минимальный content.xml с текстовыми абзацами."""
    body_parts = []
    for p in paragraphs:
        # Escape XML specials minimally
        t = (
            unicode(p)
            .replace(u"&", u"&amp;")
            .replace(u"<", u"&lt;")
            .replace(u">", u"&gt;")
        )
        # Restore placeholders: we escaped << >> — put them back as literal text
        # Actually << becomes &lt;&lt; which is correct in XML text nodes.
        body_parts.append(
            u'<text:p text:style-name="Standard">%s</text:p>' % t
        )
    return (
        u'<?xml version="1.0" encoding="UTF-8"?>\n'
        u'<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        u'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        u'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
        u'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
        u'office:version="1.2">\n'
        u"<office:scripts/>\n"
        u"<office:font-face-decls/>\n"
        u"<office:automatic-styles/>\n"
        u"<office:body>\n"
        u"<office:text>\n"
        + u"\n".join(body_parts)
        + u"\n</office:text>\n"
        u"</office:body>\n"
        u"</office:document-content>\n"
    )


def _odt_styles():
    return (
        u'<?xml version="1.0" encoding="UTF-8"?>\n'
        u'<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        u'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
        u'office:version="1.2">\n'
        u"<office:styles/>\n"
        u"<office:automatic-styles/>\n"
        u"<office:master-styles/>\n"
        u"</office:document-styles>\n"
    )


def _odt_meta():
    return (
        u'<?xml version="1.0" encoding="UTF-8"?>\n'
        u'<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        u'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" '
        u'office:version="1.2">\n'
        u"<office:meta>"
        u"<meta:generator>libre-macros generate_atext_letters_sources</meta:generator>"
        u"</office:meta>\n"
        u"</office:document-meta>\n"
    )


def _odt_manifest():
    return (
        u'<?xml version="1.0" encoding="UTF-8"?>\n'
        u'<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" '
        u'manifest:version="1.2">\n'
        u'<manifest:file-entry manifest:full-path="/" '
        u'manifest:media-type="application/vnd.oasis.opendocument.text"/>\n'
        u'<manifest:file-entry manifest:full-path="content.xml" '
        u'manifest:media-type="text/xml"/>\n'
        u'<manifest:file-entry manifest:full-path="styles.xml" '
        u'manifest:media-type="text/xml"/>\n'
        u'<manifest:file-entry manifest:full-path="meta.xml" '
        u'manifest:media-type="text/xml"/>\n'
        u"</manifest:manifest>\n"
    )


def write_odt(path, paragraphs):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    # ODT: mimetype must be first, stored (no compression)
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr(
            "mimetype",
            "application/vnd.oasis.opendocument.text",
            compress_type=zipfile.ZIP_STORED,
        )
        zf.writestr("META-INF/manifest.xml", _odt_manifest().encode("utf-8"))
        zf.writestr("content.xml", _odt_content(paragraphs).encode("utf-8"))
        zf.writestr("styles.xml", _odt_styles().encode("utf-8"))
        zf.writestr("meta.xml", _odt_meta().encode("utf-8"))
    print("  %s" % path)


def generate_atext_letters_sources():
    Workbook = _wb()
    os.makedirs(OUT_DIR, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = u"Получатели"
    ws.append([u"ФИО", u"Адрес", u"Сумма"])
    rows = [
        (u"Иванов", u"ул. Ленина, 1", 1000),
        (u"Петров", u"пр. Мира, 2", 2000),
        (u"Сидоров", u"пер. Садовый, 3", 3000),
        (u"Козлов", u"ул. Речная, 4", 4000),
        (u"", u"пусто_фио", 0),  # для row_filter
    ]
    for r in rows:
        ws.append(list(r))
    wb.save(OUT_XLSX)
    print("  %s" % OUT_XLSX)

    write_odt(
        TPL_LETTER,
        [
            u"Уведомление",
            u"Уважаемый(ая) <<ФИО>>!",
            u"Адрес: <<Адрес>>",
            u"Сумма: <<Сумма>>",
            u"Номер письма: <<n>>",
        ],
    )
    write_odt(
        TPL_PAIR,
        [
            u"Пара адресатов",
            u"Основной: <<ФИО>>{1}",
            u"Дополнительный: <<ФИО>>{2}",
            u"Адрес 1: <<Адрес>>{1}",
            u"Адрес 2: <<Адрес>>{2}",
            u"Сумма 1: <<Сумма>>{1}",
            u"Сумма 2: <<Сумма>>{2}",
            u"Файл № <<n>>, строки <<row>>{1}-<<row>>{2}",
        ],
    )

    readme = os.path.join(OUT_DIR, "README.md")
    with open(readme, "w", encoding="utf-8") as f:
        f.write(
            u"# atext_letters\n\n"
            u"- `source_recipients.xlsx` — лист «Получатели» (ФИО, Адрес, Сумма).\n"
            u"- `template_letter.odt` — `block_rows=1`, плейсхолдеры `<<ФИО>>` и т.д.\n"
            u"- `template_pair.odt` — `block_rows=2`, `<<…>>{1}` / `{2}`.\n\n"
            u"Сценарий: `20_atext_letters.xltx`\n\n"
            u"Пересоздать:\n"
            u"```bash\n"
            u"python test/collect_workbooks/generate_atext_letters_sources.py\n"
            u"```\n"
        )
    print("  %s" % readme)
    return OUT_XLSX


if __name__ == "__main__":
    generate_atext_letters_sources()
