# -*- coding: utf-8 -*-
"""Собрать test/file_list/search_corpus/ — файлы всех типов поиска (часть с паттернами, часть чистые)."""
from __future__ import print_function

import os
import shutil
import subprocess
import sys
import zipfile
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "search_corpus")
PATTERNS_CSV = os.path.join(ROOT, "dlp_name_patterns_example.csv")

HIT_PHRASES = (
    u"Конфиденциально: внутренний отчёт.",
    u"Document marked confidential for review.",
    u"Содержит коммерческую тайну предприятия.",
    u"This is a trade secret excerpt.",
    u"Пометка ДСП — не для внешней рассылки.",
    u"Секретный раздел приложения №1.",
    u"Strictly secret draft notes.",
    u"Персональные данные сотрудников (фрагмент).",
    u"Personal data sample / PII checklist.",
    u"Для служебного пользования.",
    u"Только для внутреннего использования.",
    u"Файл password_list и перечень паролей.",
    u"NDA agreement summary.",
    u"Гриф секретности на титуле.",
    u"Материалы ограниченного распространения.",
    u"Сведения составляют банковскую тайну.",
    u"Сведения составляют врачебную тайну.",
    u"Описание ноу-хау процесса.",
    u"Инсайдерская информация по сделке.",
    u"Сведения составляют государственную тайну.",
)

CLEAN_PHRASES = (
    u"Обычный квартальный отчёт без пометок.",
    u"Список задач на неделю и график встреч.",
    u"Публичный прайс-лист и описание услуг.",
    u"Учебный пример таблицы продаж.",
    u"Черновик презентации для открытого вебинара.",
)

TEXT_EXTS = (
    ".txt",
    ".csv",
    ".tsv",
    ".json",
    ".xml",
    ".log",
    ".md",
    ".html",
    ".htm",
    ".ini",
    ".cfg",
    ".yml",
    ".yaml",
    ".sql",
    ".py",
    ".js",
    ".css",
    ".properties",
    ".rtf",
)
WORD_MODERN = (".docx", ".docm", ".dotx", ".dotm", ".odt", ".ott", ".fodt", ".rtf")
EXCEL_MODERN = (".xlsx", ".xlsm", ".xltx", ".xltm", ".ods", ".ots", ".fods")
# Legacy через LO convert (если доступен soffice).
WORD_LEGACY = (".doc", ".dot")
EXCEL_LEGACY = (".xls", ".xlt", ".xlsb")


def _w(path, data, mode="wb"):
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    if "b" in mode:
        payload = data if isinstance(data, bytes) else data.encode("utf-8")
        with open(path, mode) as handle:
            handle.write(payload)
    else:
        payload = data.decode("utf-8") if isinstance(data, bytes) else data
        with open(path, mode, encoding="utf-8") as handle:
            handle.write(payload)


def _text_body(hit, idx):
    pool = HIT_PHRASES if hit else CLEAN_PHRASES
    return pool[idx % len(pool)]


def write_plain(path, text, encoding="utf-8"):
    raw = text if isinstance(text, bytes) else text.encode(encoding)
    _w(path, raw)


def write_rtf(path, text):
    # Минимальный RTF с unicode escapes для кириллицы.
    parts = [r"{\rtf1\ansi\deff0\uc1 "]
    for ch in text:
        o = ord(ch)
        if ch == "\\":
            parts.append(r"\\")
        elif ch == "{":
            parts.append(r"\{")
        elif ch == "}":
            parts.append(r"\}")
        elif o < 128:
            parts.append(ch)
        else:
            parts.append(r"\u%d?" % (o if o < 32768 else o - 65536))
    parts.append("}")
    _w(path, "".join(parts), mode="w")


def write_docx_like(path, text, content_type_override=None):
    """Минимальный OOXML Word (docx/docm/dotx/dotm)."""
    body = escape(text)
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body><w:p><w:r><w:t>%s</w:t></w:r></w:p></w:body></w:document>"
    ) % body
    ct = content_type_override or (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)


def write_xlsx_like(path, text, content_type_override=None):
    shared = escape(text)
    ct = content_type_override or (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml"/>'
        "</Relationships>"
    )
    sheet = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>'
    )
    sst = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">'
        "<si><t>%s</t></si></sst>"
    ) % shared
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet)
        zf.writestr("xl/sharedStrings.xml", sst)


def write_odt(path, text, template=False):
    content = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">'
        "<office:body><office:text><text:p>%s</text:p></office:text></office:body>"
        "</office:document-content>"
    ) % escape(text)
    mime = "application/vnd.oasis.opendocument.text-template" if template else (
        "application/vnd.oasis.opendocument.text"
    )
    meta = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'office:version="1.2"><office:meta/></office:document-meta>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'office:version="1.2"/>'
    )
    manifest = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">'
        '<manifest:file-entry manifest:full-path="/" manifest:media-type="%s"/>'
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>'
        '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>'
        "</manifest:manifest>"
    ) % mime
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("mimetype", mime, compress_type=zipfile.ZIP_STORED)
        zf.writestr("content.xml", content)
        zf.writestr("styles.xml", styles)
        zf.writestr("meta.xml", meta)
        zf.writestr("META-INF/manifest.xml", manifest)


def write_ods(path, text, template=False):
    content = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
        'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">'
        "<office:body><office:spreadsheet><table:table table:name=\"Sheet1\">"
        "<table:table-row><table:table-cell office:value-type=\"string\">"
        "<text:p>%s</text:p></table:table-cell></table:table-row>"
        "</table:table></office:spreadsheet></office:body></office:document-content>"
    ) % escape(text)
    mime = (
        "application/vnd.oasis.opendocument.spreadsheet-template"
        if template
        else "application/vnd.oasis.opendocument.spreadsheet"
    )
    meta = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'office:version="1.2"><office:meta/></office:document-meta>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'office:version="1.2"/>'
    )
    manifest = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">'
        '<manifest:file-entry manifest:full-path="/" manifest:media-type="%s"/>'
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>'
        '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>'
        "</manifest:manifest>"
    ) % mime
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("mimetype", mime, compress_type=zipfile.ZIP_STORED)
        zf.writestr("content.xml", content)
        zf.writestr("styles.xml", styles)
        zf.writestr("meta.xml", meta)
        zf.writestr("META-INF/manifest.xml", manifest)


def write_fodt(path, text):
    body = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        'office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.2">'
        "<office:body><office:text><text:p>%s</text:p></office:text></office:body>"
        "</office:document>"
    ) % escape(text)
    _w(path, body)


def write_fods(path, text):
    body = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
        'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        'office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.2">'
        "<office:body><office:spreadsheet><table:table table:name=\"Sheet1\">"
        "<table:table-row><table:table-cell office:value-type=\"string\">"
        "<text:p>%s</text:p></table:table-cell></table:table-row>"
        "</table:table></office:spreadsheet></office:body></office:document>"
    ) % escape(text)
    _w(path, body)


def write_text_by_ext(path, ext, text, idx):
    if ext == ".rtf":
        write_rtf(path, text)
        return
    if ext == ".json":
        write_plain(path, u'{"note":"%s","n":%d}\n' % (text.replace('"', "'"), idx))
        return
    if ext == ".xml":
        write_plain(path, u'<?xml version="1.0" encoding="UTF-8"?><note>%s</note>\n' % escape(text))
        return
    if ext == ".csv":
        write_plain(path, u"col1;col2\n%s;%d\n" % (text.replace(";", ","), idx))
        return
    if ext == ".tsv":
        write_plain(path, u"col1\tcol2\n%s\t%d\n" % (text.replace("\t", " "), idx))
        return
    if ext in (".yml", ".yaml"):
        write_plain(path, u"note: |\n  %s\nidx: %d\n" % (text, idx))
        return
    if ext == ".html" or ext == ".htm":
        write_plain(path, u"<html><body><p>%s</p></body></html>\n" % escape(text))
        return
    if ext == ".ini" or ext == ".cfg" or ext == ".properties":
        write_plain(path, u"[main]\nnote=%s\nidx=%d\n" % (text, idx))
        return
    if ext == ".sql":
        write_plain(path, u"-- %s\nSELECT %d;\n" % (text, idx))
        return
    if ext == ".py":
        write_plain(path, u"# %s\nVALUE = %d\n" % (text, idx))
        return
    if ext == ".js":
        write_plain(path, u"// %s\nconst n = %d;\n" % (text, idx))
        return
    if ext == ".css":
        write_plain(path, u"/* %s */\n.n { order: %d; }\n" % (text, idx))
        return
    if ext == ".md":
        write_plain(path, u"# Note\n\n%s\n" % text)
        return
    if ext == ".log":
        write_plain(path, u"INFO idx=%d %s\n" % (idx, text))
        return
    write_plain(path, text + u"\n")


_WORD_CT = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    ".docm": "application/vnd.ms-word.document.macroEnabled.main+xml",
    ".dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    ".dotm": "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
}
_EXCEL_CT = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    ".xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    ".xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}


def _ct_xml(main_part, main_type, extras=()):
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="%s" ContentType="%s"/>' % (main_part, main_type),
    ]
    parts.extend(extras)
    parts.append("</Types>")
    return "".join(parts)


def write_office(path, ext, text):
    if ext in (".docx", ".docm", ".dotx", ".dotm"):
        extras = ()
        ct = _ct_xml("/word/document.xml", _WORD_CT[ext], extras)
        write_docx_like(path, text, content_type_override=ct)
        return
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        extras = (
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
            '<Override PartName="/xl/sharedStrings.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
        )
        ct = _ct_xml("/xl/workbook.xml", _EXCEL_CT[ext], extras)
        write_xlsx_like(path, text, content_type_override=ct)
        return
    if ext == ".odt":
        write_odt(path, text, template=False)
        return
    if ext == ".ott":
        write_odt(path, text, template=True)
        return
    if ext == ".ods":
        write_ods(path, text, template=False)
        return
    if ext == ".ots":
        write_ods(path, text, template=True)
        return
    if ext == ".fodt":
        write_fodt(path, text)
        return
    if ext == ".fods":
        write_fods(path, text)
        return
    if ext == ".rtf":
        write_rtf(path, text)
        return
    raise ValueError("unsupported office ext %s" % ext)


def find_soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None


def convert_legacy(soffice, src, out_dir, target_ext):
    """target_ext без точки, напр. doc / xls."""
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--convert-to",
        target_ext,
        "--outdir",
        out_dir,
        src,
    ]
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
        return True
    except Exception as err:
        print("convert %s → %s: %s" % (src, target_ext, err))
        return False


def make_bucket(bucket, hit_content, hit_name, encodings_extra=False):
    idx = 0
    # Текстовые
    for ext in TEXT_EXTS:
        text = _text_body(hit_content, idx)
        stem = u"конфиденц_отчёт" if hit_name else u"обычный_отчёт"
        if hit_name and ext == ".txt":
            name = u"%s_%02d%s" % (stem, idx, ext)
        elif hit_name:
            name = u"secret_note_%02d%s" % (idx, ext) if idx % 2 else u"для_служебного_%02d%s" % (idx, ext)
        else:
            name = u"note_%02d%s" % (idx, ext)
        path = os.path.join(OUT, bucket, u"text", name)
        write_text_by_ext(path, ext, text, idx)
        if encodings_extra and ext == ".txt" and hit_content:
            write_plain(
                os.path.join(OUT, bucket, u"text", u"cp1251_конфиденц_%02d.txt" % idx),
                text,
                encoding="cp1251",
            )
        idx += 1

    # Word / Excel modern
    for ext in WORD_MODERN:
        if ext == ".rtf":
            continue  # уже в TEXT_EXTS
        text = _text_body(hit_content, idx)
        if hit_name:
            name = u"гриф_секретн_%02d%s" % (idx, ext)
        else:
            name = u"doc_%02d%s" % (idx, ext)
        write_office(os.path.join(OUT, bucket, u"word", name), ext, text)
        idx += 1

    for ext in EXCEL_MODERN:
        text = _text_body(hit_content, idx)
        if hit_name:
            name = u"пароли_список_%02d%s" % (idx, ext)
        else:
            name = u"sheet_%02d%s" % (idx, ext)
        write_office(os.path.join(OUT, bucket, u"excel", name), ext, text)
        idx += 1

    return idx


def make_nested(hit_content):
    text_hit = _text_body(True, 0)
    text_clean = _text_body(False, 0)
    base = os.path.join(OUT, u"nested", u"level_a", u"level_b")
    write_text_by_ext(os.path.join(base, u"deep_конфиденц.txt"), ".txt", text_hit, 0)
    write_text_by_ext(os.path.join(base, u"deep_clean.txt"), ".txt", text_clean, 1)
    write_office(os.path.join(base, u"deep_trade_secret.docx"), ".docx", text_hit)
    write_office(os.path.join(base, u"deep_clean.xlsx"), ".xlsx", text_clean)
    write_office(os.path.join(base, u"deep_dsp.odt"), ".odt", u"Пометка ДСП внутри вложенной папки.")
    write_plain(
        os.path.join(OUT, u"nested", u"level_a", u"readme_clean.md"),
        u"# Вложенная папка\n\nОбычное описание структуры.\n",
    )


def make_legacy(soffice):
    if not soffice:
        print("soffice не найден — legacy .doc/.xls/.xlsb пропущены")
        return
    tmp = os.path.join(OUT, u"_convert_src")
    if os.path.isdir(tmp):
        shutil.rmtree(tmp)
    os.makedirs(tmp)
    hit = HIT_PHRASES[0]
    clean = CLEAN_PHRASES[0]
    write_odt(os.path.join(tmp, "hit_word.odt"), hit)
    write_odt(os.path.join(tmp, "clean_word.odt"), clean)
    write_ods(os.path.join(tmp, "hit_excel.ods"), hit)
    write_ods(os.path.join(tmp, "clean_excel.ods"), clean)

    legacy_dir = os.path.join(OUT, u"legacy")
    os.makedirs(legacy_dir, exist_ok=True)

    pairs = [
        ("hit_word.odt", "doc", u"hit_конфиденц.doc"),
        ("hit_word.odt", "dot", u"hit_шаблон_секретн.dot"),
        ("clean_word.odt", "doc", u"clean_report.doc"),
        ("clean_word.odt", "dot", u"clean_template.dot"),
        ("hit_excel.ods", "xls", u"hit_пароли.xls"),
        ("hit_excel.ods", "xlt", u"hit_шаблон_дсп.xlt"),
        ("clean_excel.ods", "xls", u"clean_sheet.xls"),
        ("clean_excel.ods", "xlt", u"clean_template.xlt"),
    ]
    for src_name, fmt, out_name in pairs:
        src = os.path.join(tmp, src_name)
        convert_legacy(soffice, src, legacy_dir, fmt)
        # LO пишет basename.ext — переименуем
        produced = os.path.join(legacy_dir, os.path.splitext(src_name)[0] + "." + fmt)
        target = os.path.join(legacy_dir, out_name)
        if os.path.isfile(produced):
            if os.path.abspath(produced) != os.path.abspath(target):
                if os.path.isfile(target):
                    os.remove(target)
                os.rename(produced, target)
        else:
            print("нет файла после convert:", produced)

    # xlsb: у многих сборок LO нет export filter — кладём OLE-копию .xls как stub.
    hit_xls = os.path.join(legacy_dir, u"hit_пароли.xls")
    clean_xls = os.path.join(legacy_dir, u"clean_sheet.xls")
    if os.path.isfile(hit_xls):
        shutil.copy2(hit_xls, os.path.join(legacy_dir, u"hit_keystore.xlsb"))
    if os.path.isfile(clean_xls):
        shutil.copy2(clean_xls, os.path.join(legacy_dir, u"clean_book.xlsb"))

    shutil.rmtree(tmp, ignore_errors=True)


def write_readme():
    text = u"""# search_corpus — тестовые файлы для вкладки «Поиск»

Словарь паттернов: `../dlp_name_patterns_example.csv`.

| Папка | Имена | Содержимое |
|-------|-------|------------|
| `hit_name/` | с маркерами (конфиденц, secret, гриф…) | чистое |
| `hit_content/` | нейтральные | с маркерами из словаря |
| `hit_both/` | с маркерами | с маркерами |
| `clean/` | нейтральные | чистое |
| `nested/` | смешанные, вложенные папки | смешанное |
| `legacy/` | .doc/.dot/.xls/.xlt/.xlsb (через LibreOffice) | hit/clean |

Пересоздать: `python3 test/file_list/generate_search_corpus.py`
"""
    _w(os.path.join(OUT, "README.md"), text)


def main():
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT)
    make_bucket(u"hit_name", hit_content=False, hit_name=True)
    make_bucket(u"hit_content", hit_content=True, hit_name=False, encodings_extra=True)
    make_bucket(u"hit_both", hit_content=True, hit_name=True)
    make_bucket(u"clean", hit_content=False, hit_name=False)
    make_nested(True)
    make_legacy(find_soffice())
    # копия словаря рядом для удобства ручного теста
    shutil.copy2(PATTERNS_CSV, os.path.join(OUT, "patterns_example.csv"))
    write_readme()
    # статистика
    n = 0
    for dirpath, _dirnames, filenames in os.walk(OUT):
        n += len(filenames)
    print("OK: %s (%d files)" % (OUT, n))
    return 0


if __name__ == "__main__":
    sys.exit(main())
