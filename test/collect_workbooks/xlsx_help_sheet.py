# -*- coding: utf-8 -*-
"""
Добавление листа «Справка_макроса» в .xlsx/.xltx без пересохранения openpyxl.

Сохраняет кнопку «Обзор» на листе параметров (патч ZIP, как в старом xlsx_templates).
"""
from __future__ import print_function

import io
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"main": MAIN_NS}
ET.register_namespace("", MAIN_NS)

BASE_LINE_HEIGHT_PT = 14.0
ROW_HEIGHT_FACTOR = 1.20

HELP_SHEET_COL_WIDTH = 118.0
HELP_TAB_COLOR = "FF70AD47"
HELP_SHEET_XML = "xl/worksheets/sheet2.xml"
HELP_SHEET_RELS_XML = "xl/worksheets/_rels/sheet2.xml.rels"
EMPTY_SHEET_RELS = (
    b'<?xml version="1.0" encoding="UTF-8"?>'
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
)

STYLE_HELP_TITLE = 7
STYLE_HELP_H1 = 8
STYLE_HELP_H2 = 9
STYLE_HELP_H3 = 10
STYLE_HELP_EXAMPLE = 11
STYLE_HELP_BODY = 12


def _col_ref(col_index):
    n = col_index
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _cell_address(row, col_index):
    return "%s%d" % (_col_ref(col_index), row)


def _inline_cell(ref, text, style=None):
    cell = ET.Element("{%s}c" % MAIN_NS, {"r": ref})
    if style is not None:
        cell.set("s", str(style))
    if text is None or str(text).strip() == "":
        if style is not None:
            return cell
        return None
    cell.set("t", "inlineStr")
    is_elem = ET.SubElement(cell, "{%s}is" % MAIN_NS)
    t_elem = ET.SubElement(is_elem, "{%s}t" % MAIN_NS)
    t_elem.text = str(text)
    return cell


def _find_row(sheet_data, row_num):
    for row in sheet_data.findall("main:row", NS):
        if row.get("r") == str(row_num):
            return row
    return None


def _cell_element_text(cell):
    for is_elem in cell.findall("main:is", NS):
        parts = []
        for t_elem in is_elem.findall("main:t", NS):
            if t_elem.text:
                parts.append(t_elem.text)
        if parts:
            return "".join(parts)
    v = cell.find("main:v", NS)
    if v is not None and v.text is not None:
        return v.text
    return ""


def _estimate_text_lines(text, width_chars):
    if text is None or str(text).strip() == "":
        return 1
    s = str(text).replace("\r\n", "\n").replace("\r", "\n")
    max_chars = max(8.0, float(width_chars) * 0.85)
    lines = 0
    for part in s.split("\n"):
        part_len = len(part)
        lines = lines + max(1, int((part_len + max_chars - 1) // max_chars))
    return max(1, lines)


def classify_help_line(text):
    t = str(text).rstrip()
    if t.strip() == "":
        return STYLE_HELP_BODY
    s = t.strip()
    if len(s) >= 8 and set(s) <= set("=- "):
        return STYLE_HELP_H1
    if s.startswith("--- "):
        return STYLE_HELP_H1
    if s.startswith("• "):
        return STYLE_HELP_H2
    if s.startswith("  •") or s.startswith("    •"):
        return STYLE_HELP_H3
    low = s.lower()
    if low.startswith("пример") or "пример c:" in low or "примеры:" in low:
        return STYLE_HELP_EXAMPLE
    if low.startswith("    c:") or low.startswith("  c:"):
        return STYLE_HELP_EXAMPLE
    if s.endswith(":") and not s.startswith(" ") and len(s) < 90:
        return STYLE_HELP_H2
    if not s.startswith(" ") and len(s) < 72:
        if s[0].isupper() and "—" not in s and "." not in s:
            return STYLE_HELP_H1
    return STYLE_HELP_BODY


def _patch_styles_for_help_sheet(styles_xml):
    if isinstance(styles_xml, bytes):
        text = styles_xml.decode("utf-8")
    else:
        text = styles_xml
    root = ET.fromstring(text)
    fonts = root.find("main:fonts", NS)
    cell_xfs = root.find("main:cellXfs", NS)
    if fonts is None or cell_xfs is None:
        return styles_xml

    def _add_font(bold=False, italic=False, sz=10, rgb="FF000000"):
        f = ET.Element("{%s}font" % MAIN_NS)
        if bold:
            ET.SubElement(f, "{%s}b" % MAIN_NS, {"val": "true"})
        if italic:
            ET.SubElement(f, "{%s}i" % MAIN_NS, {"val": "true"})
        ET.SubElement(f, "{%s}sz" % MAIN_NS, {"val": str(sz)})
        ET.SubElement(f, "{%s}color" % MAIN_NS, {"rgb": rgb})
        ET.SubElement(f, "{%s}name" % MAIN_NS, {"val": "PT Sans"})
        fonts.append(f)
        return len(fonts.findall("main:font", NS)) - 1

    f_title = _add_font(bold=True, sz=14, rgb="FF1F4E79")
    f_h1 = _add_font(bold=True, sz=12, rgb="FF2E5090")
    f_h2 = _add_font(bold=True, sz=11, rgb="FF375623")
    f_h3 = _add_font(bold=True, sz=10, rgb="FF7030A0")
    f_ex = _add_font(italic=True, sz=10, rgb="FF595959")
    f_body = _add_font(sz=10, rgb="FF333333")
    fonts.set("count", str(len(fonts.findall("main:font", NS))))

    def _add_xf(font_id):
        xf = ET.Element(
            "{%s}xf" % MAIN_NS,
            {
                "numFmtId": "164",
                "fontId": str(font_id),
                "fillId": "0",
                "borderId": "0",
                "xfId": "0",
                "applyFont": "true",
                "applyAlignment": "true",
            },
        )
        ET.SubElement(
            xf,
            "{%s}alignment" % MAIN_NS,
            {
                "horizontal": "left",
                "vertical": "top",
                "wrapText": "true",
            },
        )
        cell_xfs.append(xf)

    _add_xf(f_title)
    _add_xf(f_h1)
    _add_xf(f_h2)
    _add_xf(f_h3)
    _add_xf(f_ex)
    _add_xf(f_body)
    cell_xfs.set("count", str(len(cell_xfs.findall("main:xf", NS))))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _help_sheet_row(row_num, text, style):
    row = ET.Element(
        "{%s}row" % MAIN_NS,
        {
            "r": str(row_num),
            "customFormat": "false",
            "hidden": "false",
            "customHeight": "false",
            "outlineLevel": "0",
            "collapsed": "false",
        },
    )
    cell = _inline_cell(_cell_address(row_num, 1), text, style)
    if cell is not None:
        row.append(cell)
    return row


def _apply_help_row_heights(sheet_data, last_row, col_width):
    row_num = 1
    while row_num <= last_row:
        row = _find_row(sheet_data, row_num)
        if row is None:
            row_num = row_num + 1
            continue
        text = ""
        for cell in row.findall("main:c", NS):
            text = _cell_element_text(cell)
            break
        max_lines = _estimate_text_lines(text, col_width)
        height = BASE_LINE_HEIGHT_PT * max_lines * ROW_HEIGHT_FACTOR
        row.set("ht", "%.2f" % height)
        row.set("customHeight", "true")
        row_num = row_num + 1


def build_help_sheet_xml(title, lines):
    ws = ET.Element("{%s}worksheet" % MAIN_NS)
    sheet_pr = ET.SubElement(ws, "{%s}sheetPr" % MAIN_NS)
    ET.SubElement(sheet_pr, "{%s}tabColor" % MAIN_NS, {"rgb": HELP_TAB_COLOR})
    cols = ET.SubElement(ws, "{%s}cols" % MAIN_NS)
    ET.SubElement(
        cols,
        "{%s}col" % MAIN_NS,
        {
            "min": "1",
            "max": "1",
            "width": "%.2f" % HELP_SHEET_COL_WIDTH,
            "customWidth": "1",
        },
    )
    sheet_data = ET.SubElement(ws, "{%s}sheetData" % MAIN_NS)
    sheet_data.append(_help_sheet_row(1, title, STYLE_HELP_TITLE))
    row_num = 2
    for line in lines:
        sheet_data.append(
            _help_sheet_row(row_num, line, classify_help_line(line))
        )
        row_num = row_num + 1
    last_row = max(1, row_num - 1)
    ET.SubElement(
        ws,
        "{%s}dimension" % MAIN_NS,
        {"ref": "A1:A%d" % last_row},
    )
    _apply_help_row_heights(sheet_data, last_row, HELP_SHEET_COL_WIDTH)
    sheet_views = ET.SubElement(ws, "{%s}sheetViews" % MAIN_NS)
    ET.SubElement(
        sheet_views,
        "{%s}sheetView" % MAIN_NS,
        {"workbookViewId": "0"},
    )
    ET.SubElement(
        ws,
        "{%s}sheetFormatPr" % MAIN_NS,
        {
            "defaultRowHeight": "15",
            "customHeight": "false",
            "outlineLevelRow": "0",
            "outlineLevelCol": "0",
        },
    )
    return ET.tostring(ws, encoding="utf-8", xml_declaration=True)


def _next_workbook_rel_id(rels_root):
    max_id = 0
    for rel in rels_root:
        rid = rel.get("Id", "")
        m = re.match(r"^rId(\d+)$", rid)
        if m:
            max_id = max(max_id, int(m.group(1)))
    return "rId%d" % (max_id + 1)


def _patch_workbook_add_help_sheet(workbook_xml, sheet_name, rel_id):
    root = ET.fromstring(workbook_xml)
    sheets = root.find("main:sheets", NS)
    if sheets is None:
        sheets = ET.SubElement(root, "{%s}sheets" % MAIN_NS)
    for sh in list(sheets.findall("main:sheet", NS)):
        if sh.get("name") == sheet_name:
            sheets.remove(sh)
    sheet_id = 2
    for sh in sheets.findall("main:sheet", NS):
        try:
            sid = int(sh.get("sheetId", "0"))
            if sid >= sheet_id:
                sheet_id = sid + 1
        except Exception:
            pass
    new_sheet = ET.Element(
        "{%s}sheet" % MAIN_NS,
        {
            "name": sheet_name,
            "sheetId": str(sheet_id),
            "state": "visible",
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": rel_id,
        },
    )
    if len(sheets.findall("main:sheet", NS)) == 0:
        sheets.append(new_sheet)
    else:
        sheets.insert(1, new_sheet)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_workbook_rels_add_sheet(rels_xml, rel_id, target):
    root = ET.fromstring(rels_xml)
    for rel in list(root):
        if rel.get("Target") == target:
            root.remove(rel)
    ET.SubElement(
        root,
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
        {
            "Id": rel_id,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "Target": target,
        },
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_content_types_add_help_sheet(ct_xml):
    text = ct_xml.decode("utf-8") if isinstance(ct_xml, bytes) else ct_xml
    overrides = (
        '<Override PartName="/xl/worksheets/sheet2.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '<Override PartName="/xl/worksheets/_rels/sheet2.xml.rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    )
    for ov in overrides:
        if ov not in text:
            text = text.replace("</Types>", ov + "</Types>")
    return text.encode("utf-8")


def append_help_sheet(dest_path, sheet_name, title, lines):
    """
    Добавить лист справки в .xlsx/.xltx (второй после «Параметры_Объединения»).
    """
    dest_path = Path(dest_path)
    with zipfile.ZipFile(dest_path, "r") as zin:
        names = zin.namelist()
        files = {n: zin.read(n) for n in names}

    rel_id = "rId5"
    if "xl/_rels/workbook.xml.rels" in files:
        rel_root = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
        rel_id = _next_workbook_rel_id(rel_root)
        for rel in rel_root:
            if rel.get("Target") == "worksheets/sheet2.xml":
                rel_id = rel.get("Id", rel_id)
                break

    if "xl/styles.xml" in files:
        files["xl/styles.xml"] = _patch_styles_for_help_sheet(files["xl/styles.xml"])
    files[HELP_SHEET_XML] = build_help_sheet_xml(title, lines)
    files[HELP_SHEET_RELS_XML] = EMPTY_SHEET_RELS
    files["xl/workbook.xml"] = _patch_workbook_add_help_sheet(
        files["xl/workbook.xml"], sheet_name, rel_id
    )
    files["xl/_rels/workbook.xml.rels"] = _patch_workbook_rels_add_sheet(
        files["xl/_rels/workbook.xml.rels"], rel_id, "worksheets/sheet2.xml"
    )
    if "[Content_Types].xml" in files:
        files["[Content_Types].xml"] = _patch_content_types_add_help_sheet(
            files["[Content_Types].xml"]
        )

    out_names = list(names)
    for extra in (HELP_SHEET_XML, HELP_SHEET_RELS_XML):
        if extra not in out_names:
            out_names.append(extra)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zout:
        for name in out_names:
            zout.writestr(name, files[name])
    dest_path.write_bytes(buf.getvalue())
    return dest_path
