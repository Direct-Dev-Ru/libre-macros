# -*- coding: utf-8 -*-
"""Собрать ts_ownership_months.ods: вшитый макрос, лист Sample с кнопкой и примерами."""
from __future__ import print_function

import os
import sys
import zipfile
import xml.etree.ElementTree as ET
import xml.sax.saxutils

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from ts_ownership_months import count_months, parse_date, parse_quarter  # noqa: E402

SCRIPT_NAME = "ts_ownership_months.py"
ODS_NAME = "ts_ownership_months.ods"
MACRO_HREF = (
    "vnd.sun.star.script:ts_ownership_months.py$ts_ownership_months"
    "?language=Python&amp;location=document"
)
SAMPLE_SHEET = "Sample"
HELP_SHEET = "Инструкция"
TEMPLATE_ODS = os.path.normpath(os.path.join(HERE, "..", "installer", "template.ods"))
MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"


def _esc(text):
    return xml.sax.saxutils.escape(str(text))


def _expected_result(disposal_text, dereg_text, quarter_text):
    disposal = parse_date(disposal_text)
    if disposal is None:
        return "ошибка"
    parsed = parse_quarter(quarter_text)
    if parsed is None:
        return "ошибка"
    quarter, year = parsed
    dereg = parse_date(dereg_text) if dereg_text else None
    return count_months(disposal, dereg, quarter, year, verbose=False)


def sample_rows():
    """Строки данных: выбытие, снятие, квартал, сценарий. E заполняет макрос."""
    rows = [
        (
            "06.02.2026",
            "27.06.2026",
            "2.2026",
            "Пример из описания: выбытие до квартала, снятие 27.06 — все 3 месяца Q2",
        ),
        (
            "20.03.2026",
            "19.05.2026",
            "2.2026",
            "Пример из описания: снятие 19.05 — апрель и май да, июнь нет (снятие ≤ 15.06)",
        ),
        (
            "29.01.2026",
            "20.04.2026",
            "2.2026",
            "Пример из описания: снятие 20.04 — только апрель (снятие > 15.04, но ≤ 15.05)",
        ),
        (
            "01.01.2026",
            "",
            "2.2026",
            "Снятия нет: выбытие до квартала — правило 1 на все месяцы Q2",
        ),
        (
            "15.04.2026",
            "",
            "2.2026",
            "Граница: выбытие ровно 15.04 без снятия — апрель ещё считается",
        ),
        (
            "16.04.2026",
            "",
            "2.2026",
            "Правило 3: выбытие 16.04 — апрель не считается, май и июнь да",
        ),
        (
            "16.06.2026",
            "",
            "2.2026",
            "Выбытие после 15 июня — ни одного месяца квартала",
        ),
        (
            "15.05.2026",
            "15.05.2026",
            "2.2026",
            "Оба 15.05: апрель — выбытие позже 15.04; май и июнь — правило 2",
        ),
        (
            "15.05.2026",
            "16.05.2026",
            "2.2026",
            "Снятие на день позже 15-го: май считается, июнь нет",
        ),
        (
            "01.04.2026",
            "16.04.2026",
            "2.2026",
            "Снятие сразу после 15 апреля — только апрель",
        ),
        (
            "01.04.2026",
            "10.04.2026",
            "2.2026",
            "Снятие до 15 апреля включительно — ни одного месяца (правило 2)",
        ),
        (
            "20.05.2026",
            "",
            "2.2026",
            "Выбытие 20.05 без снятия — апрель и май нет (правило 3), июнь да",
        ),
        (
            "10.01.2026",
            "20.02.2026",
            "1.2026",
            "I квартал: январь да, февраль да (снятие 20.02 > 15.02), март нет",
        ),
        (
            "16.01.2026",
            "28.03.2026",
            "1.2026",
            "Выбытие 16.01: январь по правилу 3 нет, февраль и март да",
        ),
        (
            "29.02.2024",
            "",
            "1.2024",
            "Високосный год: выбытие 29.02 — январь и февраль нет, март да",
        ),
        (
            "01.07.2026",
            "10.08.2026",
            "3.2026",
            "III квартал: июль да, август нет (снятие 10.08 ≤ 15.08), сентябрь нет",
        ),
        (
            "20.09.2025",
            "20.12.2025",
            "4.2025",
            "IV квартал: выбытие до квартала, снятие 20.12 > 15.12 — все 3 месяца",
        ),
        (
            "15.10.2025",
            "15.11.2025",
            "4.2025",
            "Границы 15-го: октябрь да (снятие 15.11 > 15.10), ноябрь и декабрь нет",
        ),
        (
            "01.04.2026",
            "30.06.2026",
            "2,2026",
            "Квартал с запятой 2,2026 — разбор как 2.2026, все месяцы Q2",
        ),
        (
            "не дата",
            "01.06.2026",
            "2.2026",
            "Ошибка разбора даты выбытия → в E должно быть «ошибка»",
        ),
        (
            "01.04.2026",
            "01.07.2026",
            "5.2026",
            "Ошибка: квартал 5 не существует → в E должно быть «ошибка»",
        ),
        (
            "32.01.2026",
            "01.02.2026",
            "1.2026",
            "Ошибка: несуществующая дата 32.01.2026 → «ошибка»",
        ),
    ]
    out = []
    for disposal, dereg, quarter, note in rows:
        out.append((disposal, dereg, quarter, note, _expected_result(disposal, dereg, quarter)))
    return out


def help_rows():
    return [
        ("Макрос ts_ownership_months",),
        ("Считает, сколько месяцев квартала засчитывается по датам выбытия и снятия с учёта.",),
        ("",),
        ("Как запустить",),
        ("1. Оставайтесь на листе Sample (макрос читает активный лист).",),
        ("2. Разрешите макросы, если LibreOffice / AlterOffice спросит.",),
        ("3. Нажмите кнопку «Посчитать месяцы» справа от таблицы.",),
        ("4. Результат пишется в колонку E. Журнал — новый лист «Лог_ТС_месяцы».",),
        ("Колонка F — ожидаемое значение, макрос её не трогает (для сверки).",),
        ("",),
        ("Колонки на Sample",),
        ("A — дата выбытия ТС, текст ДД.ММ.ГГГГ (обязательна; пустая A = конец данных).",),
        ("B — дата снятия с учёта, тот же формат или пусто.",),
        ("C — квартал и год, например 2.2026 (допускается запятая: 2,2026).",),
        ("D — комментарий, макрос не использует.",),
        ("E — число месяцев или текст «ошибка».",),
        ("",),
        ("Правила на каждый месяц M квартала (граница — 15-е число)",),
        ("1. Выбытие ≤ 15.M и (снятия нет или снятие > 15.M) — месяц считается.",),
        ("2. Выбытие ≤ 15.M и снятие ≤ 15.M — месяц не считается.",),
        ("3. Выбытие > 15.M — месяц не считается.",),
        ("",),
        ("Даты в образце записаны текстом ДД.ММ.ГГГГ: макрос читает getString().",),
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


def _content_xml(rows):
    headers = (
        "Дата выбытия ТС",
        "Дата снятия с учёта",
        "Квартал и год",
        "Сценарий",
        "Месяцев (E)",
        "Ожидаемо",
        "",
    )
    sample_body = []
    sample_body.append('<table:table-row table:style-name="roHead">')
    for i, title in enumerate(headers):
        if i == 6:
            sample_body.append(
                '<table:table-cell table:style-name="ceBtn">'
                '<draw:control draw:z-index="0" draw:name="RunButtonShape" '
                'draw:style-name="gr1" draw:text-style-name="P1" '
                'svg:width="48mm" svg:height="9mm" svg:x="2mm" svg:y="1mm" '
                'draw:control="control1"/>'
                "</table:table-cell>"
            )
        elif title:
            sample_body.append(_str_cell("ceHead", title))
        else:
            sample_body.append(_empty_cell("ceHead"))
    sample_body.append("</table:table-row>")

    for disposal, dereg, quarter, note, expected in rows:
        sample_body.append('<table:table-row table:style-name="roData">')
        sample_body.append(_str_cell("ceText", disposal))
        if dereg:
            sample_body.append(_str_cell("ceText", dereg))
        else:
            sample_body.append(_empty_cell("ceText"))
        sample_body.append(_str_cell("ceText", quarter))
        sample_body.append(_str_cell("ceNote", note))
        sample_body.append(_empty_cell("ceResult"))
        sample_body.append(_str_cell("ceExpect", expected))
        sample_body.append(_empty_cell())
        sample_body.append("</table:table-row>")

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
    <style:table-column-properties fo:break-before="auto" style:column-width="38mm"/>
  </style:style>
  <style:style style:name="coB" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="40mm"/>
  </style:style>
  <style:style style:name="coC" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="28mm"/>
  </style:style>
  <style:style style:name="coD" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="118mm"/>
  </style:style>
  <style:style style:name="coE" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="28mm"/>
  </style:style>
  <style:style style:name="coG" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="54mm"/>
  </style:style>
  <style:style style:name="coHelp" style:family="table-column">
    <style:table-column-properties fo:break-before="auto" style:column-width="220mm"/>
  </style:style>
  <style:style style:name="roHead" style:family="table-row">
    <style:table-row-properties style:row-height="11mm" fo:break-before="auto" style:use-optimal-row-height="false"/>
  </style:style>
  <style:style style:name="roData" style:family="table-row">
    <style:table-row-properties style:row-height="9mm" fo:break-before="auto" style:use-optimal-row-height="false"/>
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
    <style:paragraph-properties fo:text-align="center"/>
    <style:text-properties fo:color="#ffffff" fo:font-size="10pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="ceBtn" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties style:vertical-align="middle"/>
  </style:style>
  <style:style style:name="ceText" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties style:text-align-source="fix" fo:border="0.5pt solid #808080" style:vertical-align="middle"/>
    <style:paragraph-properties fo:text-align="center"/>
    <style:text-properties fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="ceNote" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties style:text-align-source="fix" fo:border="0.5pt solid #808080" fo:wrap-option="wrap" style:vertical-align="middle"/>
    <style:paragraph-properties fo:text-align="left"/>
    <style:text-properties fo:font-size="9pt"/>
  </style:style>
  <style:style style:name="ceResult" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties fo:background-color="#fff3cd" style:text-align-source="fix" fo:border="0.5pt solid #808080" style:vertical-align="middle"/>
    <style:paragraph-properties fo:text-align="center"/>
    <style:text-properties fo:font-size="11pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="ceExpect" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties fo:background-color="#e8f5e9" style:text-align-source="fix" fo:border="0.5pt solid #808080" style:vertical-align="middle"/>
    <style:paragraph-properties fo:text-align="center"/>
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
        'xml:id="control1" form:id="control1" form:label="Посчитать месяцы" '
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
        '<table:table-column table:style-name="coC"/>'
        '<table:table-column table:style-name="coD"/>'
        '<table:table-column table:style-name="coE" table:number-columns-repeated="2"/>'
        '<table:table-column table:style-name="coG"/>'
        "%s"
        "</table:table>"
        '<table:table table:name="%s" table:style-name="ta1">'
        '<table:table-column table:style-name="coHelp"/>'
        "%s"
        "</table:table>"
        "<table:named-expressions/>"
        "</office:spreadsheet></office:body></office:document-content>"
        % (ns, styles, SAMPLE_SHEET, form, "".join(sample_body), HELP_SHEET, "".join(help_body))
    )


def _meta_xml():
    return """<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 office:version="1.3">
 <office:meta>
  <dc:title>Месяцы владения ТС</dc:title>
  <dc:description>Образец для макроса ts_ownership_months: таблица дат выбытия и снятия с учёта.</dc:description>
  <meta:generator>libre-macros-work/standalone/build_ts_ownership_ods.py</meta:generator>
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
    rows = sample_rows()
    content = _content_xml(rows).encode("utf-8")
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
                    data = data.replace(b"Installer_Libre_Macro", SAMPLE_SHEET.encode("utf-8"))
                if item.filename == "mimetype":
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)
            _zip_write(zout, "Scripts/", b"")
            _zip_write(zout, "Scripts/python/", b"")
            _zip_write(zout, "Scripts/python/" + SCRIPT_NAME, script_src)
    return output_path, rows


def main():
    path, rows = build_ods()
    print("OK:", path)
    print("строк образца:", len(rows))
    for i, row in enumerate(rows, start=2):
        print("  %2d  A=%s  B=%s  C=%s  ожидаемо=%s" % (i, row[0], row[1] or "(пусто)", row[2], row[4]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
