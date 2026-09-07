# -*- coding: utf-8 -*-
"""Собрать import_xml.ods: вшитый макрос, кнопка запуска, инструкция."""
from __future__ import print_function

import os
import sys
import zipfile
import xml.etree.ElementTree as ET
import xml.sax.saxutils

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

SCRIPT_NAME = "import_xml.py"
ODS_NAME = "import_xml.ods"
SAMPLE_XML = "sample_employees.xml"
MACRO_HREF = (
    "vnd.sun.star.script:import_xml.py$import_xml"
    "?language=Python&amp;location=document"
)
DATA_SHEET = "Данные"
HELP_SHEET = "Инструкция"
TEMPLATE_ODS = os.path.normpath(os.path.join(HERE, "..", "installer", "template.ods"))
MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"


def _esc(text):
    return xml.sax.saxutils.escape(str(text))


def help_rows():
    return [
        ("Макрос import_xml",),
        ("Импорт XML в активную книгу: плоская таблица, как в сборе книг (collect_workbooks).",),
        ("",),
        ("Как запустить",),
        ("1. Разрешите макросы, если LibreOffice / AlterOffice спросит.",),
        ("2. Нажмите кнопку «Импорт XML» на листе «Данные» (или Сервис → Макрос).",),
        ("3. Укажите путь к .xml (рядом лежит sample_employees.xml) и при необходимости имя листа.",),
        ("4. OK — данные выгрузятся и оформятся (заголовок, автоширина, закрепление, TableStyleLight1).",),
        ("",),
        ("Поля диалога",),
        ("Файл XML — путь к импортируемому файлу; кнопка «Обзор…».",),
        ("Лист выгрузки — имя листа этой книги. Пусто: новый лист; если текущий абсолютно пуст — на него.",),
        ("Точки входа — имена Worksheet (SpreadsheetML) или xpath-теги; через запятую. Пусто — весь документ.",),
        ("Текст → число/дата — по умолчанию выкл.; включите, чтобы распознавать числа и даты в столбцах.",),
        ("",),
        ("Форматы XML",),
        ("• SpreadsheetML (Office 2003 XML) — листы = Worksheet/@Name.",),
        ("• Произвольный XML — иерархический flatten; колонки = цепочки тегов (как Employee.Name).",),
        ("• Encoding берётся из объявления <?xml … encoding=…?> (utf-8, windows-1251, …).",),
        ("",),
        ("Оформление по умолчанию (как лист результата сбора)",),
        ("Шрифт PT Sans 11; серый жирный заголовок; автоширина + запас; закрепление строки 1;",),
        ("smart-table / автофильтр со стилем TableStyleLight1 (если сборка AO/LO поддерживает).",),
        ("Конвертация текста в число/дату — только если включена опция в диалоге.",),
    ]


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
    data_body = []
    data_body.append('<table:table-row table:style-name="roHead">')
    data_body.append(_str_cell("ceHead", "Импорт XML"))
    data_body.append(
        '<table:table-cell table:style-name="ceBtn">'
        '<draw:control draw:z-index="0" draw:name="RunButtonShape" '
        'draw:style-name="gr1" draw:text-style-name="P1" '
        'svg:width="42mm" svg:height="9mm" svg:x="2mm" svg:y="1mm" '
        'draw:control="control1"/>'
        "</table:table-cell>"
    )
    data_body.append(_empty_cell())
    data_body.append("</table:table-row>")
    data_body.append('<table:table-row table:style-name="roData">')
    data_body.append(
        _str_cell(
            "ceNote",
            "Нажмите «Импорт XML». Пустой лист — данные появятся здесь, если в диалоге не указать другое имя.",
        )
    )
    data_body.append(_empty_cell())
    data_body.append(_empty_cell())
    data_body.append("</table:table-row>")

    help_body = []
    for i, row in enumerate(help_rows()):
        style = "roHelpTitle" if i == 0 else "roHelp"
        cell_style = "ceHelpTitle" if i == 0 else "ceHelp"
        help_body.append('<table:table-row table:style-name="%s">' % style)
        help_body.append(_str_cell(cell_style, row[0]))
        help_body.append("</table:table-row>")

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
    <style:table-column-properties fo:break-before="auto" style:column-width="140mm"/>
  </style:style>
  <style:style style:name="coB" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="50mm"/>
  </style:style>
  <style:style style:name="coHelp" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="240mm"/>
  </style:style>
  <style:style style:name="roHead" style:family="table-row">
    <style:table-row-properties style:row-height="12mm" fo:break-before="auto" style:use-optimal-row-height="false"/>
  </style:style>
  <style:style style:name="roData" style:family="table-row">
    <style:table-row-properties style:row-height="14mm" fo:break-before="auto" style:use-optimal-row-height="false"/>
  </style:style>
  <style:style style:name="roHelpTitle" style:family="table-row">
    <style:table-row-properties style:row-height="10mm" fo:break-before="auto"/>
  </style:style>
  <style:style style:name="roHelp" style:family="table-row">
    <style:table-row-properties style:row-height="7mm" fo:break-before="auto"/>
  </style:style>
  <style:style style:name="ta1" style:family="table" style:master-page-name="Default">
    <style:table-properties table:display="true" style:writing-mode="lr-tb"/>
  </style:style>
  <style:style style:name="ceHead" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties fo:background-color="#55308d" style:text-align-source="fix" fo:wrap-option="wrap" style:vertical-align="middle"/>
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
  <style:style style:name="ceHelpTitle" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties style:vertical-align="middle"/>
    <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="ceHelp" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties fo:wrap-option="wrap" style:vertical-align="middle"/>
    <style:text-properties fo:font-size="11pt"/>
  </style:style>
  <style:style style:name="gr1" style:family="graphic" style:parent-style-name="Default">
    <style:graphic-properties fo:background-color="#55308d"/>
    <style:text-properties fo:color="#ffffff" fo:font-size="12pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="P1" style:family="paragraph">
    <style:paragraph-properties fo:text-align="center"/>
    <style:text-properties fo:color="#ffffff" fo:font-size="12pt" fo:font-weight="bold"/>
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
        'xml:id="control1" form:id="control1" form:label="Импорт XML" '
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
        "</form:button></form:form></office:forms>"
        % MACRO_HREF
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
        "%s"
        "</table:table>"
        '<table:table table:name="%s" table:style-name="ta1">'
        '<table:table-column table:style-name="coHelp"/>'
        "%s"
        "</table:table>"
        "<table:named-expressions/>"
        "</office:spreadsheet></office:body></office:document-content>"
        % (ns, styles, DATA_SHEET, form, "".join(data_body), HELP_SHEET, "".join(help_body))
    )


def _meta_xml():
    return """<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 office:version="1.3">
 <office:meta>
  <dc:title>Импорт XML</dc:title>
  <dc:description>Самодостаточный макрос import_xml: XML → лист Calc.</dc:description>
  <meta:generator>libre-macros/standalone/build_import_xml_ods.py</meta:generator>
 </office:meta>
</office:document-meta>
"""


def _zip_write(zout, name, data, stored=False):
    if not isinstance(data, bytes):
        data = data.encode("utf-8")
    info = zipfile.ZipInfo(name)
    info.compress_type = zipfile.ZIP_STORED if stored else zipfile.ZIP_DEFLATED
    zout.writestr(info, data)


def _ensure_manifest_script_entries(manifest_data):
    root = ET.fromstring(manifest_data)
    existing = {}
    for entry in root.findall("{%s}file-entry" % MANIFEST_NS):
        path = entry.get("{%s}full-path" % MANIFEST_NS)
        if path:
            existing[path] = entry
    extra = (
        ("Scripts/", "application/binary"),
        ("Scripts/python/", "application/binary"),
        ("Scripts/python/" + SCRIPT_NAME, "application/binary"),
    )
    for full_path, media_type in extra:
        entry = existing.get(full_path)
        if entry is None:
            entry = ET.SubElement(root, "{%s}file-entry" % MANIFEST_NS)
            entry.set("{%s}full-path" % MANIFEST_NS, full_path)
        entry.set("{%s}media-type" % MANIFEST_NS, media_type)
    ET.register_namespace("manifest", MANIFEST_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_manifest_rdf(rdf_data):
    text = rdf_data.decode("utf-8")
    part = "Scripts/python/" + SCRIPT_NAME
    if part in text:
        return rdf_data
    block = (
        '  <rdf:Description rdf:about="%s">\n'
        '    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/>\n'
        "  </rdf:Description>\n"
        '  <rdf:Description rdf:about="">\n'
        '    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" '
        'rdf:resource="%s"/>\n'
        "  </rdf:Description>\n"
        % (part, part)
    )
    return text.replace("</rdf:RDF>", block + "</rdf:RDF>").encode("utf-8")


def build_ods(output_path=None):
    output_path = output_path or os.path.join(HERE, ODS_NAME)
    if not os.path.isfile(TEMPLATE_ODS):
        raise IOError("Не найден шаблон ODS: %s" % TEMPLATE_ODS)
    script_path = os.path.join(HERE, SCRIPT_NAME)
    with open(script_path, "r", encoding="utf-8") as fh:
        script_src = fh.read()
    content = _content_xml().encode("utf-8")
    with zipfile.ZipFile(TEMPLATE_ODS, "r") as zin:
        manifest = _ensure_manifest_script_entries(zin.read("META-INF/manifest.xml"))
        with zipfile.ZipFile(output_path, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "content.xml":
                    data = content
                elif item.filename == "META-INF/manifest.xml":
                    data = manifest
                elif item.filename == "manifest.rdf":
                    data = _patch_manifest_rdf(data)
                elif item.filename == "meta.xml":
                    data = _meta_xml().encode("utf-8")
                elif item.filename == "settings.xml":
                    data = data.replace(b"Installer_Libre_Macro", DATA_SHEET.encode("utf-8"))
                if item.filename == "mimetype":
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)
            _zip_write(zout, "Scripts/", b"")
            _zip_write(zout, "Scripts/python/", b"")
            _zip_write(zout, "Scripts/python/" + SCRIPT_NAME, script_src)
    return output_path


def main():
    path = build_ods()
    sample = os.path.join(HERE, SAMPLE_XML)
    print("OK:", path)
    print("sample XML:", sample if os.path.isfile(sample) else "(нет)")
    print("script:", os.path.join(HERE, SCRIPT_NAME))
    return 0


if __name__ == "__main__":
    sys.exit(main())
