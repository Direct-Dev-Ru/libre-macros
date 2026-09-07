# -*- coding: utf-8 -*-
"""
DatabaseRange / Excel Table: преобразовать в диапазон или сменить table-style.

В Calc таблицы Excel и именованные «диапазоны БД» хранятся в doc.DatabaseRanges.
«Преобразовать в диапазон» = двойной reopen + UNO + patch XML
  (.ods: plain database-range (автофильтр) + named-range;
   .xlsx: autoFilter на листах, workbook.xml definedNames, удалить xl/tables).
«Применить стиль» = двойной reopen + патч XML в ZIP
  (1: save/close/open; 2: save/close/patch/open;
   .ods: content.xml+styles.xml; .xlsx: xl/tables).
"""
from __future__ import print_function
MACRO_VERSION = "3.10.691"
import os
import re
import tempfile
import zipfile

import uno

from libre_macros_convert_tables_cfg import MACRO_VERSION, TARGET_TABLE_STYLE

try:
    unicode
except NameError:
    unicode = str

_TARGET_TABLE_STYLE = TARGET_TABLE_STYLE

_DB_RANGE_BLOCK_RE = re.compile(
    r"(<table:database-range\b[^>]*>)(.*?)(</table:database-range>)",
    re.DOTALL,
)
# Self-closing: <table:database-range .../>
_DB_RANGE_SELF_RE = re.compile(r"<table:database-range\b([^>]*?)/>")
# <table:style .../> или <table:style ...></table:style>
_TABLE_STYLE_TAG_RE = re.compile(
    r"<table:style\b[^>]*?(?:/>|>\s*</table:style>)",
    re.DOTALL,
)


def _log(msg):
    try:
        print("[convert_tables] %s" % msg)
    except Exception:
        pass


def _table_style_xml(style_name):
    """Эталонный элемент table:style внутри database-range."""
    return (
        '<table:style table:style-name="%s" table:first-column="false" '
        'table:last-column="false" table:odd-rows="false" '
        'table:odd-columns="false" table:font="true"/>'
        % str(style_name)
    )


def _ensure_style_in_db_range_body(body, style_elem):
    """
    В теле database-range: заменить существующий table:style или создать новый.
    Возвращает (new_body, 'replaced'|'created').
    """
    if _TABLE_STYLE_TAG_RE.search(body):
        return _TABLE_STYLE_TAG_RE.sub(style_elem, body, count=1), "replaced"
    # Создать перед закрывающим тегом (после columns и пр.).
    stripped = body.rstrip()
    # Сохранить завершающий перевод строки тела, если был.
    suffix = body[len(stripped) :]
    if stripped and not stripped.endswith("\n"):
        insert = "\n                    " + style_elem
    elif stripped:
        insert = "                    " + style_elem
    else:
        insert = style_elem
    return stripped + insert + suffix, "created"


def _patch_content_xml_table_styles(xml_text, style_name):
    """
    Во всех table:database-range выставить/создать table:style по эталону.

    Если <table:style> нет — создаём; если есть — заменяем на эталон.
    Self-closing <table:database-range .../> разворачиваем и добавляем style.

    Возвращает (new_xml, count_patched).
    """
    elem = _table_style_xml(style_name)
    counter = [0]
    created = [0]
    replaced = [0]

    # Self-closing → пустой открытый блок; стиль добавит проход ниже.
    def _expand_self(match):
        attrs = match.group(1).rstrip()
        return "<table:database-range%s></table:database-range>" % attrs

    text = _DB_RANGE_SELF_RE.sub(_expand_self, xml_text)

    def _repl_block(match):
        open_tag = match.group(1)
        body = match.group(2)
        close_tag = match.group(3)
        body, kind = _ensure_style_in_db_range_body(body, elem)
        counter[0] = counter[0] + 1
        if kind == "created":
            created[0] = created[0] + 1
        else:
            replaced[0] = replaced[0] + 1
        return open_tag + body + close_tag

    text = _DB_RANGE_BLOCK_RE.sub(_repl_block, text)
    _log(
        "patch styles: total=%d created=%d replaced=%d"
        % (counter[0], created[0], replaced[0])
    )
    return text, counter[0]


_DB_COLUMNS_BLOCK_RE = re.compile(
    r"<table:columns\b.*?</table:columns>\s*", re.DOTALL
)
_SMART_TABLE_ATTR_RE = re.compile(
    r'\s+table:(?:is-smart-table|autoexpand)="[^"]*"', re.IGNORECASE
)


def _strip_smart_table_open_tag(open_tag):
    """Убрать признаки smart-table из открывающего тега database-range."""
    tag = _SMART_TABLE_ATTR_RE.sub("", open_tag)
    return tag


def _collect_ods_db_range_items(db_ranges_xml):
    """
    Собрать (name, target, base, cell_range) из блока database-ranges.
    Поддерживает и <table:database-range .../>, и блочную smart-table разметку.
    """
    items = []
    seen = set()
    for m in _DB_RANGE_BLOCK_RE.finditer(db_ranges_xml):
        name, target = _parse_ods_db_range_attrs(m.group(1))
        if not name or not target or name in seen:
            continue
        base, cell_range = _ods_target_to_named_addrs(target)
        items.append((name, target, base, cell_range))
        seen.add(name)
    for m in _DB_RANGE_SELF_RE.finditer(db_ranges_xml):
        name, target = _parse_ods_db_range_attrs(
            "<table:database-range%s>" % m.group(1)
        )
        if not name or not target or name in seen:
            continue
        base, cell_range = _ods_target_to_named_addrs(target)
        items.append((name, target, base, cell_range))
        seen.add(name)
    return items


def _ods_plain_database_range_xml(name, target):
    """Обычный диапазон БД с автофильтром (без smart-table)."""
    return (
        '<table:database-range table:name="%s" '
        'table:target-range-address="%s" table:display-filter-buttons="true"/>'
        % (name, target)
    )


def _insert_ods_plain_database_ranges(xml_text, db_items):
    """Вставить plain database-ranges после named-expressions."""
    if not db_items:
        return xml_text
    lines = [
        _ods_plain_database_range_xml(name, target)
        for name, target, _base, _cell_range in db_items
    ]
    block = (
        "<table:database-ranges>"
        + "".join(lines)
        + "</table:database-ranges>"
    )
    ne_close = re.search(r"</table:named-expressions>", xml_text)
    if ne_close:
        pos = ne_close.end()
        return xml_text[:pos] + block + xml_text[pos:]
    ins = xml_text.rfind("</table:table>")
    if ins >= 0:
        pos = xml_text.find(">", ins) + 1
        return xml_text[:pos] + block + xml_text[pos:]
    return xml_text + block


def _patch_ods_content_convert_to_plain_ranges(xml_text):
    """
    Smart-table database-ranges → named-range + plain database-range с автофильтром.

    1) Удалить smart-table блок database-ranges
    2) Восстановить table:named-range в named-expressions
    3) Создать plain self-closing database-range (display-filter-buttons)

    Возвращает (new_xml, count_converted, block_removed).
    """
    block_m = re.search(
        r"<table:database-ranges\b.*?</table:database-ranges>",
        xml_text,
        re.DOTALL,
    )
    if not block_m:
        _log("patch ODS convert: database-ranges не найден")
        return xml_text, 0, False

    db_items = _collect_ods_db_range_items(block_m.group(0))
    text = _DB_RANGES_BLOCK_RE.sub("", xml_text)
    if not db_items:
        _log(
            "patch ODS convert: database-ranges без распознанных database-range"
        )
        return text, 0, True

    pending = {
        name: (base, cell_range) for name, _target, base, cell_range in db_items
    }
    text, fixed = _patch_ods_named_expressions(text, pending, db_items)
    text = _insert_ods_plain_database_ranges(text, db_items)
    _log(
        "patch ODS convert: %d plain database-range + named-range (%d fixed)"
        % (len(db_items), fixed)
    )
    return text, len(db_items), True


_DB_RANGES_BLOCK_RE = re.compile(
    r"<table:database-ranges\b.*?</table:database-ranges>\s*", re.DOTALL
)
_DB_RANGE_NAME_ATTR_RE = re.compile(r'\btable:name="([^"]+)"')
_DB_RANGE_TARGET_ATTR_RE = re.compile(
    r'\btable:target-range-address="([^"]+)"'
)
_NAMED_EXPRESSIONS_BLOCK_RE = re.compile(
    r"(<table:named-expressions\b[^>]*>)(.*?)(</table:named-expressions>)",
    re.DOTALL,
)
_NAMED_CHILD_TAG_RE = re.compile(
    r"<table:named-(?:range|expression)\b[^>]*/>"
)


def _parse_ods_db_range_attrs(open_tag):
    name_m = _DB_RANGE_NAME_ATTR_RE.search(open_tag)
    target_m = _DB_RANGE_TARGET_ATTR_RE.search(open_tag)
    name = name_m.group(1) if name_m else ""
    target = target_m.group(1) if target_m else ""
    return name, target


def _ods_target_to_named_addrs(target):
    """
    Отчет_Q2.A1:Отчет_Q2.G801 →
    base $Отчет_Q2.$A$1, range $Отчет_Q2.$A$1:.$G$801
    """
    start, end = str(target).split(":", 1)
    sheet, c1 = start.rsplit(".", 1)
    _, c2 = end.rsplit(".", 1)

    def _abs_cell(cell):
        m = re.match(r"([A-Za-z]+)(\d+)", cell)
        if not m:
            return cell
        return "$%s$%s" % (m.group(1).upper(), m.group(2))

    a1, a2 = _abs_cell(c1), _abs_cell(c2)
    base = "$%s.%s" % (sheet, a1)
    cell_range = "$%s.%s:.%s" % (sheet, a1, a2)
    return base, cell_range


def _ods_named_range_xml(name, base, cell_range):
    return (
        '<table:named-range table:name="%s" '
        'table:base-cell-address="%s" table:cell-range-address="%s"/>'
        % (name, base, cell_range)
    )


def _patch_ods_named_expressions(xml_text, pending, db_items):
    """
    Обновить/добавить table:named-range по pending; вернуть (xml, fixed_count).
    pending: name -> (base, cell_range), мутируется (pop при замене).
    """
    ne_m = _NAMED_EXPRESSIONS_BLOCK_RE.search(xml_text)
    fixed = 0
    new_lines = []

    if ne_m:
        open_tag, body, close_tag = ne_m.group(1), ne_m.group(2), ne_m.group(3)
        for child in _NAMED_CHILD_TAG_RE.findall(body):
            name_m = _DB_RANGE_NAME_ATTR_RE.search(child)
            name = name_m.group(1) if name_m else ""
            is_broken = (
                "table:named-expression" in child
                and 'table:expression="#REF!"' in child
            )
            if name and name in pending and (
                is_broken or "table:named-expression" in child
            ):
                base, cell_range = pending.pop(name)
                new_lines.append(_ods_named_range_xml(name, base, cell_range))
                fixed = fixed + 1
            elif name and name in pending and "table:named-range" in child:
                base, cell_range = pending.pop(name)
                new_lines.append(_ods_named_range_xml(name, base, cell_range))
                fixed = fixed + 1
            else:
                new_lines.append(child.strip())
        for name, _target, base, cell_range in db_items:
            if name in pending:
                new_lines.append(_ods_named_range_xml(name, base, cell_range))
                pending.pop(name)
                fixed = fixed + 1
        inner = "\n                ".join(new_lines)
        new_block = (
            open_tag + "\n                " + inner + "\n            " + close_tag
        )
        xml_text = (
            xml_text[: ne_m.start()] + new_block + xml_text[ne_m.end() :]
        )
        return xml_text, fixed

    lines = [
        _ods_named_range_xml(n, b, cr) for n, _t, b, cr in db_items
    ]
    block_xml = (
        "            <table:named-expressions>\n                "
        + "\n                ".join(lines)
        + "\n            </table:named-expressions>\n"
    )
    ins = xml_text.rfind("</table:table>")
    if ins >= 0:
        ins = xml_text.find(">", ins) + 1
        xml_text = xml_text[:ins] + "\n" + block_xml + xml_text[ins:]
    else:
        xml_text = xml_text + block_xml
    return xml_text, len(db_items)


def _patch_ods_convert_archive(path):
    """
    content.xml: удалить database-ranges, восстановить named-range в named-expressions.
    Возвращает (count, detail_lines).
    """
    details = []
    with zipfile.ZipFile(path, "r") as zf:
        raw = zf.read("content.xml")
    bom, text_xml = _decode_xml_bytes(raw)
    new_xml, n, removed = _patch_ods_content_convert_to_plain_ranges(text_xml)
    if not removed:
        details.append("в content.xml не найдено table:database-ranges")
        return 0, details
    _rewrite_zip_members(
        path, {"content.xml": _encode_xml_text(bom, new_xml)}
    )
    if n:
        details.append(
            "content.xml: %d plain database-range (автофильтр) + named-range"
            % n
        )
    else:
        details.append(
            "content.xml: удалён database-ranges (диапазоны не распознаны)"
        )
    return n, details


_XLSX_TABLE_REF_RE = re.compile(
    r'<table\b[^>]*\bref="([^"]+)"', re.IGNORECASE
)
_XLSX_TABLE_NAME_RE = re.compile(
    r'<table\b[^>]*\bname="([^"]+)"', re.IGNORECASE
)
_XLSX_TABLE_AUTOFILTER_RE = re.compile(
    r'<autoFilter\b[^>]*\bref="([^"]+)"', re.IGNORECASE
)
_XLSX_TABLE_PARTS_RE = re.compile(
    r"<tableParts\b.*?</tableParts>\s*", re.DOTALL
)
_XLSX_WS_AUTOFILTER_RE = re.compile(
    r"<autoFilter\b[^>]*(?:/>|>\s*</autoFilter>)\s*", re.DOTALL
)
_XLSX_TABLE_REL_RE = re.compile(
    r'<Relationship\b[^>]*Type="[^"]*relationships/table"[^>]*/>\s*',
    re.IGNORECASE,
)
_XLSX_SHEET_RELS_MEMBER_RE = re.compile(
    r"^xl/worksheets/_rels/(sheet\d+)\.xml\.rels$", re.IGNORECASE
)


def _parse_xlsx_table_refs(xml_text):
    """Из xl/tables/tableN.xml: (table_ref, autofilter_ref)."""
    ref_m = _XLSX_TABLE_REF_RE.search(xml_text)
    af_m = _XLSX_TABLE_AUTOFILTER_RE.search(xml_text)
    table_ref = ref_m.group(1) if ref_m else ""
    af_ref = af_m.group(1) if af_m else table_ref
    return table_ref, af_ref


def _patch_xlsx_worksheet_remove_table(xml_text, autofilter_ref):
    """
    Убрать tableParts, добавить autoFilter на лист (если нет), filterMode=true.
    """
    text = _XLSX_TABLE_PARTS_RE.sub("", xml_text)
    af = str(autofilter_ref or "").strip()
    if af and not _XLSX_WS_AUTOFILTER_RE.search(text):
        elem = '<autoFilter ref="%s"/>' % af
        close = re.search(r"</worksheet\s*>", text)
        if close is not None:
            text = text[: close.start()] + elem + text[close.start() :]
    if re.search(r"<sheetPr\b", text):
        if re.search(r'\bfilterMode="false"', text):
            text = re.sub(
                r'(\<sheetPr\b[^>]*\bfilterMode=")false(")',
                r"\1true\2",
                text,
                count=1,
            )
        elif "filterMode=" not in text.split("</sheetPr>", 1)[0]:
            text = re.sub(
                r"(<sheetPr\b)",
                r'\1 filterMode="true"',
                text,
                count=1,
            )
    return text


def _patch_xlsx_sheet_rels_remove_table(rels_text):
    """Удалить Relationship типа table из xl/worksheets/_rels/sheetN.xml.rels."""
    return _XLSX_TABLE_REL_RE.sub("", rels_text)


def _patch_content_types_remove_tables(content_types_text, table_members):
    """Убрать Override для удаляемых xl/tables/tableN.xml."""
    omit_parts = set()
    for name in table_members or ():
        part = "/" + str(name).lstrip("/")
        omit_parts.add(part.lower())
    if not omit_parts:
        return content_types_text

    def _drop_override(match):
        part = match.group(1)
        if part.lower() in omit_parts:
            return ""
        return match.group(0)

    return re.sub(
        r'<Override\b[^>]*\bPartName="([^"]+)"[^>]*/>\s*',
        _drop_override,
        content_types_text,
    )


def _build_xlsx_table_to_sheet(zf):
    """
    table_member -> (sheet_member, rels_member).
    По Target="../tables/tableN.xml" в rels листа.
    """
    mapping = {}
    for name in zf.namelist():
        m = _XLSX_SHEET_RELS_MEMBER_RE.match(name)
        if not m:
            continue
        sheet_stem = m.group(1)
        rels_text = zf.read(name).decode("utf-8")
        for tm in re.finditer(
            r'Target="\.\./tables/(table\d+\.xml)"', rels_text, re.IGNORECASE
        ):
            table_member = "xl/tables/" + tm.group(1)
            mapping[table_member] = (
                "xl/worksheets/%s.xml" % sheet_stem,
                name,
            )
    return mapping


def _build_xlsx_sheet_index(zf):
    """
    xl/worksheets/sheetN.xml -> {name, local_sheet_id}.
    local_sheet_id — 0-based индекс в <sheets> workbook.xml.
    """
    wb_text = zf.read("xl/workbook.xml").decode("utf-8")
    rels_text = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_to_member = {}
    for m in re.finditer(
        r'<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"',
        rels_text,
        re.IGNORECASE,
    ):
        target = m.group(2).lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        rid_to_member[m.group(1)] = target

    index = {}
    sheet_idx = 0
    for m in re.finditer(
        r'<sheet\b[^>]*\bname="([^"]+)"[^>]*\br:id="([^"]+)"',
        wb_text,
        re.IGNORECASE,
    ):
        sheet_name = m.group(1)
        member = rid_to_member.get(m.group(2), "")
        if member.startswith("xl/worksheets/"):
            index[member] = {
                "name": sheet_name,
                "local_sheet_id": sheet_idx,
            }
        sheet_idx = sheet_idx + 1
    return index


def _xlsx_ref_to_defined_name(sheet_name, ref):
    """A1:G801 + Отчет_Q1 → Отчет_Q1!$A$1:$G$801"""
    ref = str(ref or "").strip()
    if not ref or ":" not in ref:
        return ""

    def _abs_cell(cell):
        m = re.match(r"([A-Za-z]+)(\d+)", cell)
        if not m:
            return cell
        return "$%s$%s" % (m.group(1).upper(), m.group(2))

    c1, c2 = ref.split(":", 1)
    return "%s!%s:%s" % (sheet_name, _abs_cell(c1), _abs_cell(c2))


def _patch_xlsx_defined_name_value(workbook_text, name, new_value, local_sheet_id=None):
    """
    Заменить текст существующего definedName или вернуть XML нового элемента.
    local_sheet_id — только для _xlnm._FilterDatabase (атрибут localSheetId).
    """
    if local_sheet_id is None:
        pat = re.compile(
            r'(<definedName\b(?=[^>]*\bname="%s")[^>]*>)[^<]*(</definedName>)'
            % re.escape(name),
            re.DOTALL | re.IGNORECASE,
        )
    else:
        pat = re.compile(
            r'(<definedName\b(?=[^>]*\bname="%s")(?=[^>]*\blocalSheetId="%d")[^>]*>)'
            r"[^<]*(</definedName>)"
            % (re.escape(name), int(local_sheet_id)),
            re.DOTALL | re.IGNORECASE,
        )

    if pat.search(workbook_text):
        return pat.sub(r"\1" + new_value + r"\2", workbook_text, count=1), False

    if local_sheet_id is None:
        elem = (
            '<definedName function="false" hidden="false" name="%s" '
            'vbProcedure="false">%s</definedName>'
            % (name, new_value)
        )
    else:
        elem = (
            '<definedName function="false" hidden="true" localSheetId="%d" '
            'name="%s" vbProcedure="false">%s</definedName>'
            % (int(local_sheet_id), name, new_value)
        )
    return workbook_text, elem


def _patch_xlsx_workbook_defined_names(workbook_text, filter_entries):
    """
    Восстановить AF_* и _xlnm._FilterDatabase с #REF! по диапазонам таблиц.
    filter_entries: [{name, sheet_name, local_sheet_id, ref}, ...]
    """
    if not filter_entries:
        return workbook_text, 0

    text = workbook_text
    new_elems = []
    fixed = 0

    for item in filter_entries:
        name = item.get("name") or ""
        sheet_name = item.get("sheet_name") or ""
        ref = item.get("ref") or ""
        local_id = item.get("local_sheet_id")
        excel_ref = _xlsx_ref_to_defined_name(sheet_name, ref)
        if not name or not excel_ref:
            continue

        text, elem = _patch_xlsx_defined_name_value(text, name, excel_ref)
        if elem:
            new_elems.append(elem)
        else:
            fixed = fixed + 1

        filter_name = "_xlnm._FilterDatabase"
        text, elem = _patch_xlsx_defined_name_value(
            text, filter_name, excel_ref, local_sheet_id=local_id
        )
        if elem:
            new_elems.append(elem)
        else:
            fixed = fixed + 1

    if new_elems:
        dn_open = re.search(r"<definedNames\b[^>]*>", text)
        if dn_open:
            insert_at = dn_open.end()
            chunk = "\n        " + "\n        ".join(new_elems)
            text = text[:insert_at] + chunk + text[insert_at:]
            fixed = fixed + len(new_elems)
        else:
            block = (
                "    <definedNames>\n        "
                + "\n        ".join(new_elems)
                + "\n    </definedNames>\n"
            )
            sheets_close = re.search(r"</sheets>", text)
            if sheets_close:
                pos = sheets_close.end()
                text = text[:pos] + "\n" + block + text[pos:]
            else:
                text = text + block
            fixed = fixed + len(new_elems)

    return text, fixed


def _patch_xlsx_convert_archive(path):
    """
    XLSX convert: tableN.xml → autoFilter на листе, удалить xl/tables/*.
    Возвращает (count, detail_lines).
    """
    details = []
    updates = {}
    omit = set()
    sheet_work = {}
    filter_entries = []

    with zipfile.ZipFile(path, "r") as zf:
        names = zf.namelist()
        table_members = sorted(
            n for n in names if _XLSX_TABLE_MEMBER_RE.match(n)
        )
        if not table_members:
            details.append("в архиве нет xl/tables/table*.xml")
            return 0, details

        table_to_sheet = _build_xlsx_table_to_sheet(zf)
        sheet_index = _build_xlsx_sheet_index(zf)
        ct_raw = zf.read("[Content_Types].xml")
        wb_raw = zf.read("xl/workbook.xml")

        for table_member in table_members:
            if table_member not in table_to_sheet:
                details.append("%s: лист не найден, пропуск" % table_member)
                continue
            sheet_member, rels_member = table_to_sheet[table_member]
            bom, table_xml = _decode_xml_bytes(zf.read(table_member))
            _ = bom
            table_ref, af_ref = _parse_xlsx_table_refs(table_xml)
            if not table_ref:
                details.append("%s: нет ref, пропуск" % table_member)
                continue
            name_m = _XLSX_TABLE_NAME_RE.search(table_xml)
            table_name = name_m.group(1) if name_m else ""
            sheet_info = sheet_index.get(sheet_member) or {}
            sheet_title = sheet_info.get("name") or ""
            local_id = sheet_info.get("local_sheet_id")
            if table_name and sheet_title:
                filter_entries.append(
                    {
                        "name": table_name,
                        "sheet_name": sheet_title,
                        "local_sheet_id": local_id,
                        "ref": af_ref or table_ref,
                    }
                )
            info = sheet_work.setdefault(
                sheet_member,
                {"rels_member": rels_member, "af_refs": [], "tables": []},
            )
            if info["rels_member"] != rels_member:
                details.append(
                    "%s: разные rels для одного листа, пропуск" % table_member
                )
                continue
            info["af_refs"].append(af_ref or table_ref)
            info["tables"].append(table_member)
            omit.add(table_member)

        for sheet_member, info in sorted(sheet_work.items()):
            rels_member = info["rels_member"]
            af_ref = info["af_refs"][-1]
            ws_bom, ws_text = _decode_xml_bytes(zf.read(sheet_member))
            new_ws = _patch_xlsx_worksheet_remove_table(ws_text, af_ref)
            updates[sheet_member] = _encode_xml_text(ws_bom, new_ws)

            rel_bom, rel_text = _decode_xml_bytes(zf.read(rels_member))
            new_rel = _patch_xlsx_sheet_rels_remove_table(rel_text)
            updates[rels_member] = _encode_xml_text(rel_bom, new_rel)

            details.append(
                "%s: table→autoFilter (%s), удалено %d table*.xml"
                % (sheet_member, af_ref, len(info["tables"]))
            )

        ct_bom, ct_text = _decode_xml_bytes(ct_raw)
        new_ct = _patch_content_types_remove_tables(ct_text, omit)
        updates["[Content_Types].xml"] = _encode_xml_text(ct_bom, new_ct)

        wb_bom, wb_text = _decode_xml_bytes(wb_raw)
        new_wb, n_names = _patch_xlsx_workbook_defined_names(
            wb_text, filter_entries
        )
        if n_names:
            updates["xl/workbook.xml"] = _encode_xml_text(wb_bom, new_wb)
            details.append(
                "workbook.xml: восстановлено %d definedName (AF + FilterDatabase)"
                % n_names
            )

    if not omit:
        return 0, details

    _rewrite_zip_members(path, updates, omit)
    n = len(omit)
    _log("patch XLSX convert: removed %d table xml, sheets=%d" % (n, len(sheet_work)))
    details.insert(
        0,
        "XLSX: %d таблиц → autoFilter на листах, xl/tables удалены" % n,
    )
    return n, details


def _rewrite_zip_members(zip_path, updates=None, omit=None):
    """
    Перезаписать и/или удалить файлы внутри ZIP (ODS/XLSX).
    updates: dict member_name -> bytes; omit: set имён для исключения из архива.
    """
    updates = dict(updates or {})
    omit = set(omit or ())
    if not updates and not omit:
        return
    zip_path = os.path.abspath(str(zip_path))
    fd, tmp_path = tempfile.mkstemp(
        suffix=os.path.splitext(zip_path)[1] or ".zip",
        dir=os.path.dirname(zip_path) or None,
    )
    os.close(fd)
    try:
        with zipfile.ZipFile(zip_path, "r") as zin:
            names = zin.namelist()
            missing = [n for n in updates if n not in names]
            if missing:
                raise ValueError("в архиве нет: %s" % ", ".join(missing))
            with zipfile.ZipFile(tmp_path, "w") as zout:
                if "mimetype" in names:
                    info = zin.getinfo("mimetype")
                    zi = zipfile.ZipInfo("mimetype")
                    zi.compress_type = zipfile.ZIP_STORED
                    zi.date_time = info.date_time
                    data = updates.get("mimetype", zin.read("mimetype"))
                    zout.writestr(zi, data)
                for info in zin.infolist():
                    name = info.filename
                    if name == "mimetype":
                        continue
                    if name in omit:
                        continue
                    data = updates[name] if name in updates else zin.read(name)
                    zi = zipfile.ZipInfo(filename=name)
                    zi.compress_type = info.compress_type
                    zi.date_time = info.date_time
                    try:
                        zi.external_attr = info.external_attr
                    except Exception:
                        pass
                    zout.writestr(zi, data)
        os.replace(tmp_path, zip_path)
    except Exception:
        try:
            if os.path.isfile(tmp_path):
                os.unlink(tmp_path)
        except Exception:
            pass
        raise


def _rewrite_zip_member(zip_path, member_name, new_data):
    """Перезаписать один файл внутри ZIP."""
    _rewrite_zip_members(zip_path, {member_name: new_data})


def _get_doc_url(doc):
    for getter in (
        lambda: doc.URL,
        lambda: doc.getURL(),
        lambda: doc.getLocation(),
    ):
        try:
            u = getter()
            if u:
                return str(u)
        except Exception:
            pass
    return ""


def _url_to_path(url):
    try:
        return os.path.abspath(uno.fileUrlToSystemPath(str(url)))
    except Exception:
        return ""


def _path_to_url(path):
    return uno.systemPathToFileUrl(os.path.abspath(str(path)))


def _store_doc(doc):
    try:
        doc.store()
        return True
    except Exception as err:
        _log("store: %s" % err)
        url = _get_doc_url(doc)
        if not url:
            return False
        try:
            doc.storeToURL(url, ())
            return True
        except Exception as err2:
            _log("storeToURL: %s" % err2)
            return False


def _make_prop(name, value):
    p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    p.Name = str(name)
    p.Value = value
    return p


def _store_filter_name_for_path(path):
    ext = os.path.splitext(str(path or ""))[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return "Calc MS Excel 2007 XML"
    if ext == ".xls":
        return "MS Excel 97"
    return "calc8"


def _store_doc_to_path(doc, path, overwrite=True):
    path = os.path.abspath(str(path or "").strip())
    if doc is None or path == "":
        return False
    url = _path_to_url(path)
    props = (
        _make_prop("FilterName", _store_filter_name_for_path(path)),
        _make_prop("Overwrite", bool(overwrite)),
    )
    try:
        doc.storeToURL(url, props)
        return True
    except Exception as err:
        _log("storeToURL(%s): %s" % (os.path.basename(path), err))
        return False


def _close_doc(doc):
    try:
        doc.close(True)
        return True
    except Exception as err:
        _log("close: %s" % err)
        return False


def _open_doc(path):
    desktop = _get_desktop()
    if desktop is None:
        return None
    try:
        return desktop.loadComponentFromURL(_path_to_url(path), "_blank", 0, ())
    except Exception as err:
        _log("open: %s" % err)
        return None


def _script_context():
    try:
        import __main__
        return getattr(__main__, "XSCRIPTCONTEXT", None)
    except Exception:
        return None


def _get_uno_context():
    xsc = _script_context()
    if xsc is not None:
        try:
            return xsc.getComponentContext()
        except Exception:
            pass
    try:
        return uno.getComponentContext()
    except Exception:
        return None


def _get_desktop():
    xsc = _script_context()
    if xsc is not None:
        try:
            return xsc.getDesktop()
        except Exception:
            pass
    try:
        ctx = _get_uno_context()
        if ctx is None:
            return None
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    except Exception:
        return None


def _get_toolkit():
    xsc = _script_context()
    if xsc is not None:
        try:
            tk = xsc.getDesktop().getToolkit()
            if tk is not None:
                return tk
        except Exception:
            pass
    try:
        ctx = _get_uno_context()
        if ctx is None:
            return None
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    except Exception:
        return None


def _resolve_doc(doc=None):
    if doc is not None:
        try:
            if doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                return doc
        except Exception:
            pass
    desktop = _get_desktop()
    if desktop is None:
        return None
    try:
        cur = desktop.getCurrentComponent()
        if cur is not None and cur.supportsService(
            "com.sun.star.sheet.SpreadsheetDocument"
        ):
            return cur
    except Exception:
        pass
    return None


def _dialog_parent(doc):
    xsc = _script_context()
    if doc is None and xsc is not None:
        try:
            doc = xsc.getDocument()
        except Exception:
            doc = None
    if doc is not None:
        try:
            return doc.getCurrentController().getFrame().getContainerWindow()
        except Exception:
            pass
    if xsc is not None:
        try:
            frame = xsc.getDesktop().getCurrentFrame()
            if frame is not None:
                return frame.getContainerWindow()
        except Exception:
            pass
    try:
        desktop = _get_desktop()
        if desktop is not None:
            return desktop.getCurrentFrame().getContainerWindow()
    except Exception:
        pass
    toolkit = _get_toolkit()
    if toolkit is not None:
        try:
            return toolkit.getDesktopWindow()
        except Exception:
            pass
    return None


def _show_message(doc, title, text, is_error=False):
    toolkit = _get_toolkit()
    if toolkit is None:
        _log("%s: %s" % (title, text))
        return
    parent = _dialog_parent(doc)
    if parent is None:
        try:
            parent = toolkit.getDesktopWindow()
        except Exception:
            parent = None
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX

    box_type = ERRORBOX if is_error else INFOBOX
    try:
        box = toolkit.createMessageBox(parent, box_type, BUTTONS_OK, title, text)
        box.execute()
    except Exception as err:
        _log("MessageBox failed: %s; %s: %s" % (err, title, text))


def _get_active_sheet(doc):
    try:
        return doc.getCurrentController().getActiveSheet()
    except Exception:
        return None


def _get_active_sheet_name(doc):
    sheet = _get_active_sheet(doc)
    try:
        return str(sheet.Name) if sheet is not None else ""
    except Exception:
        return ""


def _bounce_to_first_and_back_by_name(doc, launch_sheet_name):
    """
    Если активный лист не первый — перейти на первый, затем вернуться на launch_sheet_name.
    """
    if doc is None:
        return
    try:
        controller = doc.getCurrentController()
        sheets = doc.getSheets()
        if controller is None or sheets is None or sheets.getCount() < 1:
            return
        first = sheets.getByIndex(0)
        active = controller.getActiveSheet()
        launch = None
        if launch_sheet_name:
            try:
                launch = sheets.getByName(str(launch_sheet_name))
            except Exception:
                launch = None
        if launch is None:
            launch = active
        try:
            if active is not None and first is not None and active.Name == first.Name:
                return
        except Exception:
            pass
        controller.setActiveSheet(first)
        if launch is not None:
            controller.setActiveSheet(launch)
        _log("bounce sheet: first → launch")
    except Exception as err:
        _log("bounce sheet: %s" % err)


def _bounce_to_first_and_back(doc, launch_sheet):
    """
    Если активный лист не первый — перейти на первый, затем вернуться на launch_sheet.
    Нужно для обновления UI после снятия/создания DatabaseRange.
    """
    if doc is None:
        return
    try:
        controller = doc.getCurrentController()
        sheets = doc.getSheets()
        if controller is None or sheets is None or sheets.getCount() < 1:
            return
        first = sheets.getByIndex(0)
        active = controller.getActiveSheet()
        if launch_sheet is None:
            launch_sheet = active
        try:
            if active is not None and first is not None and active.Name == first.Name:
                return
        except Exception:
            pass
        controller.setActiveSheet(first)
        if launch_sheet is not None:
            controller.setActiveSheet(launch_sheet)
        _log("bounce sheet: first → launch")
    except Exception as err:
        _log("bounce sheet: %s" % err)


_TABLE_STYLE_INFO_RE = re.compile(
    r"<tableStyleInfo\b[^>]*?(?:/>|>\s*</tableStyleInfo>)",
    re.DOTALL,
)
_XLSX_TABLE_MEMBER_RE = re.compile(
    r"^xl/tables/table\d+\.xml$", re.IGNORECASE
)


def _table_style_info_xml(style_name):
    """Эталон tableStyleInfo для XLSX (xl/tables/tableN.xml)."""
    return (
        '<tableStyleInfo name="%s" showFirstColumn="0" showLastColumn="0" '
        'showRowStripes="0" showColumnStripes="0"/>'
        % str(style_name)
    )


def _patch_xlsx_table_xml(xml_text, style_name):
    """
    В одном xl/tables/tableN.xml выставить/создать tableStyleInfo.
    Возвращает (new_xml, 'replaced'|'created'|None).
    """
    elem = _table_style_info_xml(style_name)
    if _TABLE_STYLE_INFO_RE.search(xml_text):
        return _TABLE_STYLE_INFO_RE.sub(elem, xml_text, count=1), "replaced"
    # Вставить перед </table>
    close = re.search(r"</table\s*>", xml_text)
    if close is None:
        return xml_text, None
    before = xml_text[: close.start()].rstrip()
    after = xml_text[close.start() :]
    if before.endswith(">"):
        insert = "\n    " + elem + "\n"
    else:
        insert = elem
    return before + insert + after, "created"


def _decode_xml_bytes(raw):
    bom = b""
    if raw.startswith(b"\xef\xbb\xbf"):
        bom = raw[:3]
        raw = raw[3:]
    return bom, raw.decode("utf-8")


def _encode_xml_text(bom, text):
    return bom + text.encode("utf-8")


_SMART_TABLE_DEFAULT_NAME_RE = re.compile(
    r'(<style:default-style\b(?=[^>]*\bstyle:family="smart-table")[^>]*?\bstyle:name=")([^"]*)(")',
    re.DOTALL,
)


def _patch_styles_xml_smart_table_default(xml_text, style_name):
    """
    В styles.xml: style:default-style family=smart-table → style:name=эталон.
    Возвращает (new_xml, changed: bool).
    """
    style_name = str(style_name)
    new_text, n = _SMART_TABLE_DEFAULT_NAME_RE.subn(
        r"\g<1>%s\3" % style_name, xml_text, count=1
    )
    if n:
        return new_text, True
    # Нет default smart-table — вставить в начало office:styles.
    m = re.search(r"<office:styles\b[^>]*>", xml_text)
    if m is None:
        return xml_text, False
    insert = (
        '\n        <style:default-style style:family="smart-table" '
        'style:name="%s"/>' % style_name
    )
    pos = m.end()
    return xml_text[:pos] + insert + xml_text[pos:], True


def _sheet_name_from_ods_target(target):
    """Merge.A1:Merge.G10 / $Merge.$A$1 / 'Лист 1'.A1 / $'Лист'.$A$1 → имя листа."""
    start = str(target or "").split(":", 1)[0].strip()
    if "." not in start:
        return ""
    sheet, _cell = start.rsplit(".", 1)
    sheet = sheet.strip()
    # Calc часто пишет $Merge.$A$1 или $'Сводная'.$A$1
    if sheet.startswith("$"):
        sheet = sheet[1:]
    if len(sheet) >= 2 and sheet[0] == "'" and sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    return sheet


def _styles_by_sheet_lookup(styles_by_sheet):
    """{sheet: style} → lookup casefold → style; primary = самый частый стиль."""
    lookup = {}
    counts = {}
    for sn, st in dict(styles_by_sheet or {}).items():
        sn = str(sn or "").strip()
        st = str(st or "").strip()
        if sn == "" or st == "":
            continue
        lookup[sn.casefold()] = st
        counts[st] = int(counts.get(st, 0)) + 1
    primary = ""
    best = 0
    for st, n in counts.items():
        if int(n) > best:
            best = int(n)
            primary = st
    if primary == "" and lookup:
        primary = list(lookup.values())[0]
    return lookup, primary


def _patch_content_xml_table_styles_for_sheets(xml_text, styles_by_sheet):
    """
    В database-range на указанных листах выставить table:style.
    Возвращает (new_xml, count_patched, unmatched_sheets).
    """
    lookup, _primary = _styles_by_sheet_lookup(styles_by_sheet)
    if not lookup:
        return xml_text, 0, []
    counter = [0]
    seen_sheets = set()
    patched_sheets = set()

    def _expand_self(match):
        attrs = match.group(1).rstrip()
        return "<table:database-range%s></table:database-range>" % attrs

    text = _DB_RANGE_SELF_RE.sub(_expand_self, xml_text)

    def _repl_block(match):
        open_tag = match.group(1)
        body = match.group(2)
        close_tag = match.group(3)
        _name, target = _parse_ods_db_range_attrs(open_tag)
        sheet = _sheet_name_from_ods_target(target)
        sheet_cf = str(sheet).casefold() if sheet else ""
        if sheet_cf:
            seen_sheets.add(sheet_cf)
        style_name = lookup.get(sheet_cf) if sheet_cf else None
        if not style_name:
            return match.group(0)
        elem = _table_style_xml(style_name)
        body, _kind = _ensure_style_in_db_range_body(body, elem)
        counter[0] = counter[0] + 1
        patched_sheets.add(sheet_cf)
        return open_tag + body + close_tag

    text = _DB_RANGE_BLOCK_RE.sub(_repl_block, text)
    unmatched = [
        sn for sn in lookup.keys() if sn not in patched_sheets
    ]
    _log(
        "patch styles by sheet: total=%d unmatched=%s seen=%s"
        % (counter[0], unmatched, sorted(seen_sheets))
    )
    return text, counter[0], unmatched


def _patch_ods_archive(path, style_name, styles_by_sheet=None):
    """
    Правка content.xml (+ styles.xml default smart-table) в .ods.
    style_name — единый стиль; styles_by_sheet — точечно по листам.
    Возвращает (applied_count, detail_lines).
    """
    details = []
    updates = {}
    lookup, primary = _styles_by_sheet_lookup(styles_by_sheet)
    eff_default = str(style_name or primary or "").strip() or _TARGET_TABLE_STYLE
    unique_styles = sorted(set(lookup.values())) if lookup else []
    multi_style = len(unique_styles) > 1
    with zipfile.ZipFile(path, "r") as zf:
        names = zf.namelist()
        raw_content = zf.read("content.xml")
        raw_styles = zf.read("styles.xml") if "styles.xml" in names else None

    bom_c, text_content = _decode_xml_bytes(raw_content)
    if lookup:
        new_content, n, unmatched = _patch_content_xml_table_styles_for_sheets(
            text_content, styles_by_sheet
        )
        # НЕ fallback на «один стиль всем»: это как раз ломало разные стили.
        if n == 0:
            details.append(
                "content.xml: по листам 0 database-range "
                "(карта=%s; проверьте имена листов в target-range)"
                % sorted(lookup.keys())
            )
        else:
            details.append(
                "content.xml: стиль по листам в %d database-range (%s)"
                % (
                    n,
                    ", ".join(
                        "%s→%s" % (sn, lookup[sn]) for sn in sorted(lookup.keys())
                    ),
                )
            )
            if unmatched:
                details.append(
                    "content.xml: листы без диапазона: %s" % ", ".join(unmatched)
                )
    else:
        new_content, n = _patch_content_xml_table_styles(text_content, eff_default)
        if n == 0:
            details.append("в content.xml не найдено table:database-range")
        else:
            details.append(
                "content.xml: стиль %s в %d database-range" % (eff_default, n)
            )
    if n > 0:
        updates["content.xml"] = _encode_xml_text(bom_c, new_content)
        _log("patched ODS database-range: %d" % n)

    if raw_styles is not None:
        # При разных стилях по листам НЕ трогаем default smart-table —
        # иначе AO визуально тянет один стиль на все таблицы.
        if multi_style:
            details.append(
                "styles.xml: default smart-table не меняем "
                "(несколько стилей: %s)" % ", ".join(unique_styles)
            )
        else:
            bom_s, text_styles = _decode_xml_bytes(raw_styles)
            new_styles, changed = _patch_styles_xml_smart_table_default(
                text_styles, eff_default
            )
            if changed:
                updates["styles.xml"] = _encode_xml_text(bom_s, new_styles)
                details.append(
                    "styles.xml: default smart-table → %s" % eff_default
                )
                _log("patched ODS styles.xml smart-table default → %s" % eff_default)
            else:
                details.append("styles.xml: default smart-table не найден")
    else:
        details.append("styles.xml отсутствует в архиве")

    if updates:
        _rewrite_zip_members(path, updates)
    return int(n), details


def patch_table_styles_on_closed_file(path, styles_by_sheet=None, style_name=None):
    """
    Патч стиля smart-таблиц в закрытом файле на диске (.ods / .xlsx / .xlsm).

    styles_by_sheet: {имя_листа: TableStyleDark1, ...}
    style_name: единый стиль, если карта пуста / fallback.

    UNO TableStyleName после reopen в AO/LO часто не меняет визуал —
    нужен именно XML (content.xml table:style + styles.xml default).

    Возвращает (applied_count, detail_lines).
    """
    path = os.path.abspath(str(path or ""))
    details = []
    if path == "" or not os.path.isfile(path):
        return 0, ["файл не найден: %s" % path]
    lookup, primary = _styles_by_sheet_lookup(styles_by_sheet)
    eff = str(style_name or primary or "").strip() or _TARGET_TABLE_STYLE
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".ods":
            return _patch_ods_archive(
                path, eff, styles_by_sheet=styles_by_sheet if lookup else None
            )
        if ext in (".xlsx", ".xlsm"):
            return _patch_xlsx_archive(
                path, eff, styles_by_sheet=styles_by_sheet if lookup else None
            )
        details.append("формат не поддерживается для патча стиля: %s" % ext)
        return 0, details
    except Exception as err:
        return 0, ["ошибка патча XML: %s" % err]


def _patch_xlsx_archive(path, style_name, styles_by_sheet=None):
    """
    Правка xl/tables/tableN.xml в .xlsx/.xlsm.
    styles_by_sheet — стиль по имени листа; иначе style_name на все tables.
    Возвращает (applied_count, detail_lines).
    """
    details = []
    updates = {}
    applied = 0
    created = 0
    replaced = 0
    lookup, primary = _styles_by_sheet_lookup(styles_by_sheet)
    eff = str(style_name or primary or "").strip() or _TARGET_TABLE_STYLE
    with zipfile.ZipFile(path, "r") as zf:
        members = [
            n for n in zf.namelist() if _XLSX_TABLE_MEMBER_RE.match(n)
        ]
        if not members:
            details.append("в архиве нет xl/tables/table*.xml")
            return 0, details
        members = sorted(members)
        table_to_sheet = _build_xlsx_table_to_sheet(zf) if lookup else {}
        sheet_index = _build_xlsx_sheet_index(zf) if lookup else {}
        for name in members:
            style_for_table = eff
            sheet_title = ""
            if lookup:
                mapped = table_to_sheet.get(name)
                if not mapped:
                    details.append("%s: лист не найден (rels), пропуск" % name)
                    continue
                sheet_member = mapped[0]
                sheet_info = sheet_index.get(sheet_member) or {}
                sheet_title = str(sheet_info.get("name") or "").strip()
                style_for_table = lookup.get(sheet_title.casefold())
                if not style_for_table:
                    details.append(
                        "%s: лист %r не в карте стилей, пропуск"
                        % (name, sheet_title)
                    )
                    continue
            raw = zf.read(name)
            bom, text_xml = _decode_xml_bytes(raw)
            new_xml, kind = _patch_xlsx_table_xml(text_xml, style_for_table)
            if kind is None:
                details.append("%s: нет </table>, пропуск" % name)
                continue
            updates[name] = _encode_xml_text(bom, new_xml)
            applied = applied + 1
            if kind == "created":
                created = created + 1
            else:
                replaced = replaced + 1
            if sheet_title:
                details.append(
                    "%s (%s) → %s" % (name, sheet_title, style_for_table)
                )
    if updates:
        _rewrite_zip_members(path, updates)
    if lookup:
        details.append(
            "xl/tables: по листам %d таблиц (создано %d, заменено %d)"
            % (applied, created, replaced)
        )
    else:
        details.append(
            "xl/tables: стиль %s в %d таблиц (создано %d, заменено %d)"
            % (eff, applied, created, replaced)
        )
    _log(
        "patched XLSX tables: total=%d created=%d replaced=%d"
        % (applied, created, replaced)
    )
    return applied, details


def _save_close_reopen(doc, path, label=""):
    """
    Сохранить → закрыть → открыть.
    Возвращает (new_doc_or_None, ok: bool, error_text).
    """
    tag = label or "save/close/reopen"
    _log("%s: save" % tag)
    if not _store_doc(doc):
        return None, False, u"не удалось сохранить файл"
    _log("%s: close" % tag)
    if not _close_doc(doc):
        return None, False, u"не удалось закрыть файл после сохранения"
    _log("%s: open" % tag)
    new_doc = _open_doc(path)
    if new_doc is None:
        return None, False, u"не удалось открыть файл"
    return new_doc, True, u""


def convert_database_ranges_with_double_reopen(doc, launch_sheet_name=None):
    """
    Преобразовать DatabaseRange → диапазоны с двойным reopen:

      1) save → close → open
      2) снятие table/range + AutoFilter (UNO) + bounce
      3) save/export → close → patch ZIP → open
         .ods: export temp .xlsx → patch xlsx → reopen temp → save back as .ods
         .xlsx: worksheets + удаление xl/tables

    Возвращает (stats_dict, new_doc_or_None).
    """
    result = {
        "converted": 0,
        "filters": 0,
        "errors": 0,
        "details": [],
        "skipped": 0,
        "reopened": False,
    }
    if doc is None:
        result["errors"] = 1
        result["details"].append("нет активной книги Calc")
        return result, None

    url = _get_doc_url(doc)
    path = _url_to_path(url) if url else ""
    if not path or not os.path.isfile(path):
        items = collect_database_ranges(doc)
        conv = convert_database_ranges_to_plain_autofilter(doc, items)
        result.update(conv)
        _bounce_to_first_and_back_by_name(doc, launch_sheet_name)
        result["details"].append(
            u"двойной reopen пропущен (нет пути файла на диске)"
        )
        return result, doc

    doc, ok, err = _save_close_reopen(doc, path, "convert flush")
    if not ok:
        result["errors"] = 1
        result["details"].append(u"предварительный reopen: %s" % err)
        return result, doc
    result["details"].append(u"предварительный save/close/open — OK")

    items = collect_database_ranges(doc)
    _log("convert after flush: DatabaseRange=%d" % len(items))
    conv = convert_database_ranges_to_plain_autofilter(doc, items)
    for key in ("converted", "filters", "errors", "skipped"):
        result[key] = int(conv.get(key) or 0)
    conv_details = conv.get("details") or []
    if conv_details:
        result["details"].extend(conv_details)

    _bounce_to_first_and_back_by_name(doc, launch_sheet_name)

    ext = os.path.splitext(path)[1].lower()
    if ext == ".ods":
        tmp_fd = None
        tmp_xlsx = ""
        try:
            tmp_fd, tmp_xlsx = tempfile.mkstemp(suffix=".xlsx")
            os.close(tmp_fd)
            tmp_fd = None
            _log("convert (ods): export temp xlsx")
            if not _store_doc_to_path(doc, tmp_xlsx):
                result["errors"] = int(result["errors"]) + 1
                result["details"].append(
                    "не удалось сохранить временный XLSX перед патчем"
                )
                return result, doc
            result["details"].append(u"экспорт во временный XLSX — OK")
            _log("convert (ods): close source before xlsx patch")
            if not _close_doc(doc):
                result["errors"] = int(result["errors"]) + 1
                result["details"].append(
                    "не удалось закрыть исходный ODS перед патчем XLSX"
                )
                return result, None
            try:
                n_xml, xml_details = _patch_xlsx_convert_archive(tmp_xlsx)
                result["details"].extend(xml_details)
                if n_xml:
                    result["converted"] = max(int(result["converted"]), int(n_xml))
            except Exception as err:
                result["errors"] = int(result["errors"]) + 1
                result["details"].append("ошибка правки временного XLSX: %s" % err)
                _log("temp XLSX convert patch failed: %s" % err)
            _log("convert (ods): open temp xlsx after patch")
            temp_doc = _open_doc(tmp_xlsx)
            if temp_doc is None:
                result["errors"] = int(result["errors"]) + 1
                result["details"].append(
                    "не удалось открыть временный XLSX после патча"
                )
                return result, None
            _log("convert (ods): store patched temp xlsx back to ods")
            if not _store_doc_to_path(temp_doc, path):
                result["errors"] = int(result["errors"]) + 1
                result["details"].append(
                    "не удалось сохранить пропатченный XLSX обратно в ODS"
                )
                try:
                    _close_doc(temp_doc)
                except Exception:
                    pass
                return result, temp_doc
            if not _close_doc(temp_doc):
                result["errors"] = int(result["errors"]) + 1
                result["details"].append(
                    "не удалось закрыть временный XLSX после сохранения в ODS"
                )
                return result, None
            _log("convert final (ods): open source")
            doc = _open_doc(path)
            if doc is None:
                result["errors"] = int(result["errors"]) + 1
                result["details"].append(
                    "не удалось открыть ODS после сохранения из временного XLSX"
                )
            else:
                result["reopened"] = True
                result["details"].append(
                    u"финальный export temp xlsx/patch/open/save as ods/open — OK"
                )
            return result, doc
        finally:
            if tmp_fd is not None:
                try:
                    os.close(tmp_fd)
                except Exception:
                    pass
            if tmp_xlsx:
                try:
                    if os.path.isfile(tmp_xlsx):
                        os.unlink(tmp_xlsx)
                except Exception:
                    pass

    _log("convert: save before final reopen")
    if not _store_doc(doc):
        result["errors"] = int(result["errors"]) + 1
        result["details"].append("не удалось сохранить перед финальным reopen")
        return result, doc
    if not _close_doc(doc):
        result["errors"] = int(result["errors"]) + 1
        result["details"].append("не удалось закрыть перед финальным reopen")
        return result, None

    if ext in (".xlsx", ".xlsm"):
        try:
            n_xml, xml_details = _patch_xlsx_convert_archive(path)
            result["details"].extend(xml_details)
            if n_xml:
                result["converted"] = max(int(result["converted"]), int(n_xml))
        except Exception as err:
            result["errors"] = int(result["errors"]) + 1
            result["details"].append("ошибка правки XLSX: %s" % err)
            _log("XLSX convert patch failed: %s" % err)

    _log("convert final: open")
    doc = _open_doc(path)
    if doc is None:
        result["errors"] = int(result["errors"]) + 1
        result["details"].append("не удалось открыть файл после convert")
    else:
        result["reopened"] = True
        if ext in (".xlsx", ".xlsm"):
            result["details"].append(
                u"финальный save/close/patch worksheets+tables/open — OK"
            )
        else:
            result["details"].append(u"финальный save/close/open — OK")
    return result, doc


def apply_table_style_via_xml(doc, style_name=None):
    """
    Применить стиль таблицы через XML на диске (двойной reopen):

      1) save → close → open  (прогон без патча, сброс кэша формата)
      2) save → close → патч ZIP → open

    .ods: content.xml + styles.xml; .xlsx: xl/tables/tableN.xml

    Возвращает (stats_dict, new_doc_or_None).
    После первого успешного close старый doc недействителен.
    """
    if style_name is None:
        style_name = _TARGET_TABLE_STYLE
    style_name = str(style_name or "").strip() or _TARGET_TABLE_STYLE
    result = {
        "applied": 0,
        "errors": 0,
        "details": [],
        "skipped": 0,
        "style": style_name,
        "reopened": False,
    }
    if doc is None:
        result["errors"] = 1
        result["details"].append("нет активной книги Calc")
        return result, None

    url = _get_doc_url(doc)
    path = _url_to_path(url) if url else ""
    if not path or not os.path.isfile(path):
        result["errors"] = 1
        result["details"].append(
            "нужен сохранённый файл на диске (сохраните книгу и повторите)"
        )
        return result, doc

    ext = os.path.splitext(path)[1].lower()
    if ext == ".ods":
        kind = "ods"
    elif ext in (".xlsx", ".xlsm"):
        kind = "xlsx"
    else:
        result["errors"] = 1
        result["details"].append(
            "правка стиля через XML: .ods / .xlsx / .xlsm (сейчас: %s)"
            % os.path.basename(path)
        )
        return result, doc

    # 1) Предварительный save → close → open (без патча).
    doc, ok, err = _save_close_reopen(
        doc, path, "style via XML (%s) flush" % kind
    )
    if not ok:
        result["errors"] = 1
        result["details"].append(u"предварительный reopen: %s" % err)
        return result, doc
    result["details"].append(u"предварительный save/close/open — OK")

    # 2) save → close → патч → open.
    _log("style via XML (%s): save before patch" % kind)
    if not _store_doc(doc):
        result["errors"] = 1
        result["details"].append("не удалось сохранить перед патчем")
        return result, doc

    _log("style via XML (%s): close before patch" % kind)
    if not _close_doc(doc):
        result["errors"] = 1
        result["details"].append("не удалось закрыть перед патчем")
        return result, None

    try:
        if kind == "ods":
            n, details = _patch_ods_archive(path, style_name)
        else:
            n, details = _patch_xlsx_archive(path, style_name)
        result["applied"] = int(n)
        result["details"].extend(details)
        if n == 0:
            result["errors"] = 1
    except Exception as err:
        result["errors"] = 1
        result["details"].append("ошибка правки XML: %s" % err)
        _log("patch failed: %s" % err)

    _log("style via XML (%s): reopen after patch" % kind)
    new_doc = _open_doc(path)
    if new_doc is None:
        result["errors"] = result["errors"] + 1
        result["details"].append("не удалось открыть файл после патча")
    else:
        result["reopened"] = True
        result["details"].append(u"после патча save/close/patch/open — OK")
    return result, new_doc


# Совместимость со старым именем.
apply_table_style_via_ods_xml = apply_table_style_via_xml


def _copy_cell_range_address(addr):
    """CellRangeAddress → новый struct (не держим ссылку на живой объект)."""
    out = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    out.Sheet = int(addr.Sheet)
    out.StartColumn = int(addr.StartColumn)
    out.StartRow = int(addr.StartRow)
    out.EndColumn = int(addr.EndColumn)
    out.EndRow = int(addr.EndRow)
    return out


def _addr_label(doc, addr):
    try:
        sheet = doc.getSheets().getByIndex(int(addr.Sheet))
        sheet_name = sheet.Name
    except Exception:
        sheet_name = "Sheet%d" % (int(addr.Sheet) + 1)
    try:
        rng = doc.getSheets().getByIndex(int(addr.Sheet)).getCellRangeByPosition(
            int(addr.StartColumn),
            int(addr.StartRow),
            int(addr.EndColumn),
            int(addr.EndRow),
        )
        return "%s!%s" % (sheet_name, rng.AbsoluteName.split(".")[-1])
    except Exception:
        return "%s!R%dC%d:R%dC%d" % (
            sheet_name,
            int(addr.StartRow) + 1,
            int(addr.StartColumn) + 1,
            int(addr.EndRow) + 1,
            int(addr.EndColumn) + 1,
        )


def _prop_bool(obj, name, default=False):
    try:
        return bool(getattr(obj, name))
    except Exception:
        try:
            return bool(obj.getPropertyValue(name))
        except Exception:
            return default


def _prop_str(obj, name, default=""):
    try:
        v = getattr(obj, name)
        if v is None:
            return default
        return str(v)
    except Exception:
        try:
            v = obj.getPropertyValue(name)
            if v is None:
                return default
            return str(v)
        except Exception:
            return default


def _clear_table_style(db):
    """Снять признаки Excel Table / table style, если свойство есть."""
    for name, value in (
        ("TableStyleName", ""),
        ("UseRowStripes", False),
        ("UseColStripes", False),
        ("UseFirstColumnFormatting", False),
        ("UseLastColumnFormatting", False),
    ):
        try:
            setattr(db, name, value)
        except Exception:
            try:
                db.setPropertyValue(name, value)
            except Exception:
                pass


def _disable_autofilter(db):
    try:
        db.AutoFilter = False
    except Exception:
        try:
            db.setPropertyValue("AutoFilter", False)
        except Exception:
            pass


def count_database_ranges(doc):
    """Быстрый подсчёт без обхода свойств (для диалога до полного collect)."""
    try:
        names = doc.DatabaseRanges.getElementNames()
        if names is None:
            return 0
        return len(tuple(names))
    except Exception:
        return 0


def collect_database_ranges(doc):
    """
    Список dict: name, addr, had_autofilter, table_style, looks_like_table.
    """
    items = []
    try:
        db_ranges = doc.DatabaseRanges
    except Exception:
        return items
    try:
        names = list(db_ranges.getElementNames())
    except Exception:
        return items
    for name in names:
        try:
            db = db_ranges.getByName(name)
        except Exception:
            continue
        try:
            addr = _copy_cell_range_address(db.getDataArea())
        except Exception:
            continue
        had_af = _prop_bool(db, "AutoFilter", False)
        style = _prop_str(db, "TableStyleName", "")
        looks_like_table = style != ""
        # Excel Table / диапазон БД с заголовками — тоже кандидат.
        if not looks_like_table:
            # Имена Excel-таблиц обычно не анонимные __Anonymous_Sheet_DB__*
            n = str(name or "")
            if n and not n.startswith("__Anonymous_Sheet_DB__"):
                looks_like_table = True
        items.append(
            {
                "name": str(name),
                "addr": addr,
                "had_autofilter": had_af,
                "table_style": style,
                "looks_like_table": looks_like_table,
            }
        )
    return items


def convert_database_ranges_to_plain_autofilter(doc, items=None):
    """
    Снять DatabaseRange (table/range) и заново поставить AutoFilter на тот же адрес.

    Возвращает dict: converted, filters, errors, details (list of str).
    """
    result = {
        "converted": 0,
        "filters": 0,
        "errors": 0,
        "details": [],
        "skipped": 0,
    }
    if doc is None:
        result["errors"] = 1
        result["details"].append("нет активной книги Calc")
        return result
    try:
        db_ranges = doc.DatabaseRanges
    except Exception as err:
        result["errors"] = 1
        result["details"].append("DatabaseRanges недоступны: %s" % err)
        return result

    if items is None:
        items = collect_database_ranges(doc)
    if not items:
        result["details"].append("в книге нет DatabaseRange / Table")
        return result

    # Сначала собираем снимок и снимаем объекты (нельзя менять коллекцию вперемешку
    # с ленивым доступом по индексу).
    snapshot = list(items)
    for item in snapshot:
        name = item["name"]
        try:
            if not db_ranges.hasByName(name):
                result["skipped"] = result["skipped"] + 1
                continue
        except Exception:
            result["skipped"] = result["skipped"] + 1
            continue
        try:
            db = db_ranges.getByName(name)
            _disable_autofilter(db)
            _clear_table_style(db)
        except Exception:
            pass
        try:
            db_ranges.removeByName(name)
            result["converted"] = result["converted"] + 1
            result["details"].append(
                "снят: %s (%s)" % (name, _addr_label(doc, item["addr"]))
            )
        except Exception as err:
            result["errors"] = result["errors"] + 1
            result["details"].append("ошибка снятия %s: %s" % (name, err))
            _log("removeByName(%s): %s" % (name, err))

    # Повторно ставим AutoFilter как на обычный диапазон (без table style).
    for item in snapshot:
        name = item["name"]
        addr = item["addr"]
        new_name = name
        # Если имя занято (не должно) — суффикс.
        try:
            if db_ranges.hasByName(new_name):
                new_name = "AF_%s" % name
                suffix = 1
                while db_ranges.hasByName(new_name):
                    new_name = "AF_%s_%d" % (name, suffix)
                    suffix = suffix + 1
        except Exception:
            pass
        try:
            db_ranges.addNewByName(new_name, addr)
            db = db_ranges.getByName(new_name)
            _clear_table_style(db)
            try:
                db.AutoFilter = True
            except Exception:
                db.setPropertyValue("AutoFilter", True)
            result["filters"] = result["filters"] + 1
            result["details"].append(
                "автофильтр: %s (%s)" % (new_name, _addr_label(doc, addr))
            )
        except Exception as err:
            result["errors"] = result["errors"] + 1
            result["details"].append(
                "ошибка AutoFilter %s: %s" % (name, err)
            )
            _log("add AutoFilter(%s): %s" % (name, err))

    return result


def _summary_lines(action_label, stats, count_keys):
    lines = [
        u"Версия: %s" % MACRO_VERSION,
        u"Действие: %s" % action_label,
    ]
    for key, label in count_keys:
        if key in stats:
            lines.append(u"%s: %d" % (label, int(stats[key] or 0)))
    details = stats.get("details") or []
    if details:
        lines.append(u"")
        for line in details[:12]:
            lines.append(str(line))
        if len(details) > 12:
            lines.append(u"… ещё %d" % (len(details) - 12))
    return lines


def convert_tables_to_ranges(doc=None, choice=None, *args):
    """
    Точка входа: преобразовать DatabaseRange/Table в диапазоны
    или применить эталонный стиль TableStyleLight1.

    choice передаётся из entry-скрипта (диалог в convert_tables_to_ranges.py).

    Convert: двойной reopen + UNO + bounce + patch ZIP.
    Style: двойной reopen + патч ZIP (ODS/XLSX).
    """
    _ = args  # макрос может передать event
    active = _resolve_doc(doc)
    title = u"Таблицы / диапазоны БД"
    if active is None:
        _show_message(
            None,
            title,
            u"Откройте книгу Calc и повторите запуск.",
            is_error=True,
        )
        return None

    launch_sheet_name = _get_active_sheet_name(active)
    items = collect_database_ranges(active)
    _log("найдено DatabaseRange: %d" % len(items))
    if not items:
        _show_message(
            active,
            title,
            u"Версия: %s\n\nВ книге не найдено объектов DatabaseRange / Table."
            % MACRO_VERSION,
        )
        return {"converted": 0, "filters": 0, "errors": 0, "details": [], "skipped": 0}

    if choice is None:
        _log("отмена: choice не передан")
        return None

    is_error = False
    msg_doc = active
    if choice.get("action") == "style":
        style_name = choice.get("style") or _TARGET_TABLE_STYLE
        stats, msg_doc = apply_table_style_via_xml(active, style_name)
        lines = _summary_lines(
            u"применить стиль «%s» (через XML)" % style_name,
            stats,
            (
                ("applied", u"Обновлено database-range"),
                ("skipped", u"Пропущено"),
                ("errors", u"Ошибок"),
            ),
        )
        if stats.get("reopened"):
            lines.append(u"")
            lines.append(
                u"Двойной reopen: save/close/open → save/close/patch/open."
            )
        is_error = bool(stats["errors"]) and int(stats["applied"]) == 0
    else:
        stats, msg_doc = convert_database_ranges_with_double_reopen(
            active, launch_sheet_name
        )
        lines = _summary_lines(
            u"преобразовать в диапазоны",
            stats,
            (
                ("converted", u"Снято table/range"),
                ("filters", u"Автофильтр поставлен"),
                ("skipped", u"Пропущено"),
                ("errors", u"Ошибок"),
            ),
        )
        if stats.get("reopened"):
            lines.append(u"")
            lines.append(
                u"Двойной reopen: save/close/open → convert → save/close/open "
                u"(ODS: + patch content.xml)."
            )
        is_error = bool(stats["errors"]) and int(stats["filters"]) == 0

    _show_message(msg_doc, title, u"\n".join(lines), is_error=is_error)
    return stats
