# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.704"
"""
Чтение XML-источников: плоская таблица (как при открытии XML в Excel).

Поддержка:
  - SpreadsheetML (Office 2003 XML) — «Листы» = имя Worksheet
  - произвольный XML: иерархический flatten; «Листы» = точка входа (xpath-теги)
  - имена колонок — цепочка тегов: ПерТрНал.СведТСПер.НачислТС.НалБаза
  - encoding из объявления XML (windows-1251, utf-8, …)
"""

import os
import re
import tempfile

try:
    unicode
except NameError:
    unicode = str

try:
    from collections import OrderedDict
except ImportError:
    OrderedDict = dict

try:
    from defusedxml import ElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET


_RE_XML_ENCODING = re.compile(
    br'<\?xml[^>]*encoding=["\']([^"\']+)["\']', re.IGNORECASE
)

_XML_SHEET_ALL_TOKENS = frozenset(
    (u"все", u"all", u"*", u"")
)


def is_xml_source_path(path):
    p = unicode(path or u"").strip().lower()
    if p.endswith(u".xml"):
        return True
    m = re.match(r"^[^:]+://.+", p)
    if m is not None:
        q = p.split(u"?", 1)[0]
        return q.endswith(u".xml")
    return False


def normalize_xml_sheet_entry_specs(sheet_specs):
    """
    «Листы» для generic XML → точки входа (xpath-теги).
    Пусто / «Все» → None (весь документ с корня).
    """
    if not sheet_specs:
        return None
    out = []
    seen = set()
    i = 0
    while i < len(sheet_specs):
        raw = unicode(sheet_specs[i] or u"").strip()
        i = i + 1
        if raw == u"":
            continue
        low = raw.lower()
        if low in _XML_SHEET_ALL_TOKENS:
            return None
        if low in seen:
            continue
        seen.add(low)
        out.append(raw)
    if not out:
        return None
    return out


def _local_tag(tag):
    t = unicode(tag or u"")
    if u"}" in t:
        return t.rsplit(u"}", 1)[-1]
    return t


def _cell_text(elem):
    if elem is None:
        return u""
    parts = []
    if elem.text:
        parts.append(unicode(elem.text))
    for child in list(elem):
        if child.tail:
            parts.append(unicode(child.tail))
    return u"".join(parts).strip()


def _read_xml_bytes(path):
    with open(path, u"rb") as f:
        return f.read()


def _decode_xml_bytes(raw):
    enc = u"utf-8"
    m = _RE_XML_ENCODING.search(raw[:400])
    if m:
        enc = m.group(1).decode(u"ascii", u"replace").strip() or enc
    try:
        return raw.decode(enc)
    except (LookupError, UnicodeDecodeError):
        pass
    for fallback in (u"utf-8", u"cp1251", u"latin-1"):
        try:
            return raw.decode(fallback)
        except UnicodeDecodeError:
            continue
    return raw.decode(u"utf-8", errors=u"replace")


def _parse_xml_root(path):
    """Прочитать XML с учётом encoding из объявления."""
    raw = _read_xml_bytes(path)
    text = _decode_xml_bytes(raw)
    return ET.fromstring(text)


def _build_parent_map(root):
    pmap = {root: None}
    for parent in root.iter():
        for child in list(parent):
            pmap[child] = parent
    return pmap


def _parse_tag_path(spec):
    """«ПерТрНал» или «ТрНалог/ПерТрНал» → список локальных имён."""
    text = unicode(spec or u"").strip()
    if text == u"":
        return []
    for sep in (u"/", u"\\", u">", u"."):
        if sep in text:
            parts = []
            for piece in text.split(sep):
                piece = piece.strip()
                if piece != u"":
                    parts.append(piece)
            return parts
    return [text]


def _entry_sheet_title(spec, tag_path):
    if spec and unicode(spec).strip() != u"":
        title = unicode(spec).strip()
    elif tag_path:
        title = tag_path[-1]
    else:
        title = u"XML"
    title = title.replace(u"/", u"_").replace(u"\\", u"_").replace(u".", u"_")
    return title[:31] if title else u"XML"


def _find_all_by_local_name(root, name):
    matches = []
    if _local_tag(root.tag) == name:
        matches.append(root)
    for elem in root.iter():
        if elem is root:
            continue
        if _local_tag(elem.tag) == name:
            matches.append(elem)
    return matches


def _find_elements_by_tag_path(root, parts):
    if not parts:
        return [root]
    if len(parts) == 1:
        return _find_all_by_local_name(root, parts[0])

    matches = []

    def walk(elem, depth):
        tag = _local_tag(elem.tag)
        if tag == parts[depth]:
            if depth == len(parts) - 1:
                matches.append(elem)
            else:
                for child in list(elem):
                    walk(child, depth + 1)
        if depth == 0:
            for child in list(elem):
                walk(child, 0)

    walk(root, 0)
    return matches


def _path_prefix(path_tags):
    tags = [unicode(t) for t in (path_tags or []) if unicode(t or u"").strip() != u""]
    return u".".join(tags)


def _contributes_to_path(elem):
    if elem.attrib:
        return True
    for child in list(elem):
        if child.attrib:
            continue
        if len(list(child)) == 0 and _local_tag(child.tag) != u"":
            return True
    return False


def _elem_fields_at_path(elem, path_tags):
    """Поля элемента с полным путём: …Тег.Атрибут / …Тег.ТекстРебёнок."""
    tag = _local_tag(elem.tag)
    out = OrderedDict()
    if tag == u"" or not _contributes_to_path(elem):
        return out, list(path_tags or [])
    elem_path = list(path_tags or []) + [tag]
    prefix = _path_prefix(elem_path)
    for key, val in (elem.attrib or {}).items():
        lk = _local_tag(key)
        if lk == u"":
            continue
        out[u"%s.%s" % (prefix, lk)] = unicode(val)
    for child in list(elem):
        if child.attrib:
            continue
        if len(list(child)) == 0:
            ck = _local_tag(child.tag)
            if ck == u"":
                continue
            out[u"%s.%s" % (prefix, ck)] = _cell_text(child)
    return out, elem_path


def _attr_leaf_fields(row_elem, path_tags, row_tag):
    """Вложенный attribute-only: …Родитель.Строка.Лист.Атрибут."""
    out = OrderedDict()
    if row_tag == u"":
        return out
    row_path = list(path_tags or []) + [row_tag]
    for child in list(row_elem):
        ct = _local_tag(child.tag)
        if ct == u"":
            continue
        if child.attrib and len(list(child)) == 0:
            child_path = row_path + [ct]
            prefix = _path_prefix(child_path)
            for key, val in child.attrib.items():
                lk = _local_tag(key)
                if lk == u"":
                    continue
                out[u"%s.%s" % (prefix, lk)] = unicode(val)
    return out


def _collect_rows_from_entry(root, entry_elem, pmap, include_ancestors=False):
    if include_ancestors:
        ctx = _ancestor_fields(entry_elem, root, pmap)
        path_tags = []
        cur = pmap.get(entry_elem)
        chain = []
        while cur is not None and cur is not entry_elem:
            t = _local_tag(cur.tag)
            if t != u"" and _contributes_to_path(cur):
                _f, chain = _elem_fields_at_path(cur, chain)
            if cur is root:
                break
            cur = pmap.get(cur)
        path_tags = chain
    else:
        ctx = OrderedDict()
        path_tags = []
    rows = _collect_hierarchical_rows(entry_elem, ctx, path_tags, pmap)
    if not rows and ctx:
        rows = [ctx]
    return rows


def _group_by_tag(children):
    groups = OrderedDict()
    for ch in children:
        t = _local_tag(ch.tag)
        if t == u"":
            continue
        groups.setdefault(t, []).append(ch)
    return groups


def _siblings_carry_records(siblings, path_tags):
    for s in siblings:
        st = _local_tag(s.tag)
        if s.attrib:
            return True
        fields, _ep = _elem_fields_at_path(s, path_tags)
        if fields:
            return True
        if _attr_leaf_fields(s, path_tags, st):
            return True
    return False


def _has_nested_record_groups(elem, path_tags):
    for _tag, sibs in _group_by_tag(list(elem)).items():
        if len(sibs) >= 2 and _siblings_carry_records(sibs, path_tags):
            return True
        for s in sibs:
            _f, sp = _elem_fields_at_path(s, path_tags)
            if _has_nested_record_groups(s, sp):
                return True
    return False


def _ancestor_fields(elem, root, pmap):
    """Предки с полным путём (только режим «Все» / без точки входа)."""
    ctx = OrderedDict()
    chain = []
    cur = pmap.get(elem) if pmap else None
    stack = []
    while cur is not None:
        stack.append(cur)
        if cur is root:
            break
        cur = pmap.get(cur)
    stack.reverse()
    for node in stack:
        tag = _local_tag(node.tag)
        if tag == u"" or not _contributes_to_path(node):
            continue
        fields, chain = _elem_fields_at_path(node, chain)
        ctx.update(fields)
    return ctx


def _is_attr_leaf_element(elem):
    tag = _local_tag(elem.tag)
    return tag != u"" and bool(elem.attrib) and len(list(elem)) == 0


def _collect_hierarchical_rows(elem, inherited, path_tags, pmap=None):
    rows = []
    path_tags = list(path_tags or [])
    ctx = OrderedDict(inherited)
    tag = _local_tag(elem.tag)
    elem_path = list(path_tags)
    if tag != u"" and _contributes_to_path(elem):
        fields, elem_path = _elem_fields_at_path(elem, path_tags)
        ctx.update(fields)

    children = [c for c in list(elem) if _local_tag(c.tag) != u""]
    complex_children = []
    attr_leaf_children = []
    for ch in children:
        if _is_attr_leaf_element(ch):
            attr_leaf_children.append(ch)
        else:
            complex_children.append(ch)
    for _ct, sibs in _group_by_tag(attr_leaf_children).items():
        if len(sibs) >= 2:
            complex_children.extend(sibs)

    if not complex_children:
        if tag != u"":
            row = OrderedDict(ctx)
            leaf = _attr_leaf_fields(elem, path_tags, tag)
            if leaf or _elem_fields_at_path(elem, path_tags)[0]:
                row.update(leaf)
                rows.append(row)
        return rows

    for ct, sibs in _group_by_tag(complex_children).items():
        if not _siblings_carry_records(sibs, elem_path):
            for s in sibs:
                sub = OrderedDict(ctx)
                sf, sp = _elem_fields_at_path(s, elem_path)
                sub.update(sf)
                rows.extend(_collect_hierarchical_rows(s, sub, sp, pmap))
            continue

        emit_here = len(sibs) >= 2
        if not emit_here and len(sibs) == 1:
            emit_here = not _has_nested_record_groups(sibs[0], elem_path)

        if emit_here:
            for s in sibs:
                row = OrderedDict(ctx)
                sf, _sp = _elem_fields_at_path(s, elem_path)
                row.update(sf)
                row.update(_attr_leaf_fields(s, elem_path, ct))
                rows.append(row)
        else:
            for s in sibs:
                sub = OrderedDict(ctx)
                sf, sp = _elem_fields_at_path(s, elem_path)
                sub.update(sf)
                rows.extend(_collect_hierarchical_rows(s, sub, sp, pmap))
    return rows


def _rows_to_sheet(rows, sheet_name):
    col_order = []
    seen = set()
    for rd in rows:
        for k in rd.keys():
            if k not in seen:
                seen.add(k)
                col_order.append(k)
    col_order.sort(key=lambda name: (unicode(name).count(u"."), unicode(name).casefold()))
    headers = list(col_order)
    body = []
    for rd in rows:
        body.append([unicode(rd.get(h, u"")) for h in headers])
    return (unicode(sheet_name or u"XML"), headers, body)


def _parse_hierarchical_xml_root(root, entry_specs=None):
    pmap = _build_parent_map(root)
    specs = entry_specs
    if specs:
        sheets = []
        si = 0
        while si < len(specs):
            spec = specs[si]
            parts = _parse_tag_path(spec)
            si = si + 1
            if not parts:
                continue
            matches = _find_elements_by_tag_path(root, parts)
            if not matches:
                continue
            all_rows = []
            mi = 0
            while mi < len(matches):
                all_rows.extend(_collect_rows_from_entry(root, matches[mi], pmap))
                mi = mi + 1
            if not all_rows:
                continue
            sheets.append(_rows_to_sheet(all_rows, _entry_sheet_title(spec, parts)))
        return sheets

    rows = _collect_hierarchical_rows(root, OrderedDict(), [], pmap)
    if not rows:
        return []
    sheet_name = _local_tag(root.tag) or u"XML"
    return [_rows_to_sheet(rows, sheet_name)]


def _spreadsheetml_ns(root):
    tag = unicode(getattr(root, "tag", u"") or u"")
    if u"spreadsheet" in tag.lower():
        m = re.match(r"^\{([^}]+)\}", tag)
        if m:
            return m.group(1)
    for key, val in (getattr(root, "attrib", None) or {}).items():
        if key.endswith(u"xmlns") or key == u"xmlns":
            if u"spreadsheet" in unicode(val):
                return unicode(val)
    return u"urn:schemas-microsoft-com:office:spreadsheet"


def _filter_spreadsheetml_sheets(sheets, entry_specs):
    if not entry_specs:
        return sheets
    wanted = set()
    i = 0
    while i < len(entry_specs):
        spec = unicode(entry_specs[i] or u"").strip()
        i = i + 1
        if spec == u"":
            continue
        wanted.add(spec.casefold())
        parts = _parse_tag_path(spec)
        if parts:
            wanted.add(parts[-1].casefold())
    if not wanted:
        return sheets
    out = []
    for name, headers, body in sheets:
        if unicode(name or u"").strip().casefold() in wanted:
            out.append((name, headers, body))
    return out if out else sheets


def _parse_spreadsheetml_root(root, entry_specs=None):
    ns = _spreadsheetml_ns(root)
    ws_tag = u"{%s}Worksheet" % ns
    table_tag = u"{%s}Table" % ns
    row_tag = u"{%s}Row" % ns
    cell_tag = u"{%s}Cell" % ns
    data_tag = u"{%s}Data" % ns
    sheets = []
    for ws in root.iter(ws_tag):
        name = ws.attrib.get(u"{%s}Name" % ns) or ws.attrib.get(u"Name") or u"Sheet1"
        table = ws.find(table_tag)
        if table is None:
            for t in ws.iter(table_tag):
                table = t
                break
        if table is None:
            continue
        matrix = []
        for row_el in table.findall(row_tag):
            if row_el is None:
                continue
            row_vals = []
            col_idx = 0
            for cell in row_el.findall(cell_tag):
                idx_attr = cell.attrib.get(u"{%s}Index" % ns) or cell.attrib.get(u"Index")
                if idx_attr:
                    try:
                        want = int(idx_attr) - 1
                        while col_idx < want:
                            row_vals.append(u"")
                            col_idx += 1
                    except (TypeError, ValueError):
                        pass
                data_el = cell.find(data_tag)
                val = _cell_text(data_el if data_el is not None else cell)
                row_vals.append(val)
                col_idx += 1
            if any(unicode(x).strip() != u"" for x in row_vals):
                matrix.append(row_vals)
        if not matrix:
            continue
        width = max((len(r) for r in matrix), default=0)
        norm = []
        for r in matrix:
            rr = list(r)
            while len(rr) < width:
                rr.append(u"")
            norm.append(rr)
        headers = norm[0]
        body = norm[1:] if len(norm) > 1 else []
        sheets.append((unicode(name), headers, body))
    return _filter_spreadsheetml_sheets(sheets, entry_specs)


def _parse_spreadsheetml(path, entry_specs=None):
    return _parse_spreadsheetml_root(_parse_xml_root(path), entry_specs)


def _parse_hierarchical_xml(path, entry_specs=None):
    root = _parse_xml_root(path)
    return _parse_hierarchical_xml_root(root, entry_specs)


def xml_source_read_sheets(path, sheet_specs=None):
    """
    Прочитать XML → [(sheet_name, headers, rows), …].
    sheet_specs — значения «Листы» (точки входа xpath для generic XML).
    """
    p = unicode(path or u"").strip()
    if p == u"" or not os.path.isfile(p):
        return []
    entry_specs = normalize_xml_sheet_entry_specs(sheet_specs)
    try:
        root = _parse_xml_root(p)
    except Exception:
        return []
    try:
        sheets = _parse_spreadsheetml_root(root, entry_specs)
        if sheets:
            return sheets
    except Exception:
        pass
    try:
        return _parse_hierarchical_xml_root(root, entry_specs)
    except Exception:
        return []


def xml_source_to_temp_xlsx(path, sheet_specs=None, prefix=u"libre_macros_xml_"):
    """
    Конвертировать локальный .xml во временный .xlsx (openpyxl).
    sheet_specs — «Листы» (точки входа). Возвращает путь или None.
    """
    sheets = xml_source_read_sheets(path, sheet_specs=sheet_specs)
    if not sheets:
        return None
    try:
        import openpyxl_bundled  # noqa: F401
        from openpyxl import Workbook
    except ImportError:
        return None
    fd, out_path = tempfile.mkstemp(prefix=prefix, suffix=u".xlsx")
    os.close(fd)
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    for name, headers, rows in sheets:
        title = unicode(name or u"Sheet")[:31]
        for ch in u"[]:*/?\\":
            title = title.replace(ch, u"_")
        if title == u"":
            title = u"Sheet1"
        ws = wb.create_sheet(title=title)
        if headers:
            ws.append([unicode(h) for h in headers])
        for row in rows:
            ws.append([unicode(c) if c is not None else u"" for c in row])
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.number_format = u"@"
    if len(wb.sheetnames) == 0:
        try:
            os.remove(out_path)
        except Exception:
            pass
        return None
    try:
        wb.save(out_path)
    except Exception:
        try:
            os.remove(out_path)
        except Exception:
            pass
        return None
    return out_path
