# -*- coding: utf-8 -*-
"""
Макрос: импорт XML в активную книгу LibreOffice / AlterOffice Calc.

Принципы разбора — как в collect_workbooks (libre_macros_xml_source_lib):
  - SpreadsheetML (Office 2003 XML)
  - произвольный XML → плоская таблица (цепочки тегов в заголовках)
  - encoding из объявления XML

Диалог: путь к .xml, имя листа выгрузки (пусто = новый / текущий пустой),
опция «Текст → число/дата» (по умолчанию выкл.).
Таблица оформляется как лист результата сбора по умолчанию:
  шрифт PT Sans 11, серый жирный заголовок, автоширина, закрепление,
  smart-table TableStyleLight1 (если поддерживается).

Установка:
  1. Открыть import_xml.ods и нажать «Импорт XML».
  2. Либо скопировать import_xml.py в Scripts/python/ профиля LO.
"""

import os
import re
import collections
import xml.etree.ElementTree as ET


class _Store:
    lines = []
    text_type = str
    try:
        text_type = unicode  # noqa: F821  (Py2)
    except NameError:
        pass
    OrderedDict = getattr(collections, "OrderedDict", dict)
    RE_XML_ENCODING = re.compile(
        br'<\?xml[^>]*encoding=["\']([^"\']+)["\']', re.IGNORECASE
    )
    XML_SHEET_ALL_TOKENS = frozenset((u"все", u"all", u"*", u""))
    FONT_NAME = "PT Sans"
    FONT_SIZE = 11.0
    HEADER_BG = 0xD9D9D9
    CHAR_WEIGHT_BOLD = 150
    COL_PAD_INCH = 0.5
    COL_MAX_INCH = 3.0
    TABLE_STYLE = "TableStyleLight1"
    USE_ROW_STRIPES = True
    CONVERT_TO_NUMBERS = False
    FREEZE_HEADER = True


def unicode(value):
    """Совместимость Py2/Py3 внутри макроса (AO filter)."""
    return _Store.text_type(value)


def _log(msg):
    _Store.lines.append(msg)
    try:
        print("[import_xml] %s" % msg)
    except Exception:
        pass


def _reset_log():
    del _Store.lines[:]


# --- разбор XML (встроено из libre_macros_xml_source_lib) -------------------

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
        if low in _Store.XML_SHEET_ALL_TOKENS:
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
    m = _Store.RE_XML_ENCODING.search(raw[:400])
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
    out = _Store.OrderedDict()
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
    out = _Store.OrderedDict()
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
        ctx = _Store.OrderedDict()
        path_tags = []
    rows = _collect_hierarchical_rows(entry_elem, ctx, path_tags, pmap)
    if not rows and ctx:
        rows = [ctx]
    return rows


def _group_by_tag(children):
    groups = _Store.OrderedDict()
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
    ctx = _Store.OrderedDict()
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
    ctx = _Store.OrderedDict(inherited)
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
            row = _Store.OrderedDict(ctx)
            leaf = _attr_leaf_fields(elem, path_tags, tag)
            if leaf or _elem_fields_at_path(elem, path_tags)[0]:
                row.update(leaf)
                rows.append(row)
        return rows

    for ct, sibs in _group_by_tag(complex_children).items():
        if not _siblings_carry_records(sibs, elem_path):
            for s in sibs:
                sub = _Store.OrderedDict(ctx)
                sf, sp = _elem_fields_at_path(s, elem_path)
                sub.update(sf)
                rows.extend(_collect_hierarchical_rows(s, sub, sp, pmap))
            continue

        emit_here = len(sibs) >= 2
        if not emit_here and len(sibs) == 1:
            emit_here = not _has_nested_record_groups(sibs[0], elem_path)

        if emit_here:
            for s in sibs:
                row = _Store.OrderedDict(ctx)
                sf, _sp = _elem_fields_at_path(s, elem_path)
                row.update(sf)
                row.update(_attr_leaf_fields(s, elem_path, ct))
                rows.append(row)
        else:
            for s in sibs:
                sub = _Store.OrderedDict(ctx)
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

    rows = _collect_hierarchical_rows(root, _Store.OrderedDict(), [], pmap)
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



# --- диалог / LibreOffice ---------------------------------------------------

def _get_document():
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except NameError:
        pass
    if doc is None:
        try:
            import __main__
            xsc = getattr(__main__, "XSCRIPTCONTEXT", None)
            if xsc is not None:
                doc = xsc.getDocument()
        except Exception:
            pass
    return doc


def _msgbox(doc, text, title=u"Импорт XML"):
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sm = ctx.getServiceManager()
        tk = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        parent = None
        try:
            if doc is not None:
                parent = doc.getCurrentController().getFrame().getContainerWindow()
        except Exception:
            parent = None
        box = tk.createMessageBox(parent, 1, 1, title, unicode(text))
        box.execute()
    except Exception:
        _log(unicode(text))


def _dlg_add_fixed(dm, name, label, x, y, w, h, multiline=False):
    lbl = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    lbl.Name = name
    lbl.Label = label
    lbl.PositionX = x
    lbl.PositionY = y
    lbl.Width = w
    lbl.Height = h
    if multiline:
        try:
            lbl.MultiLine = True
        except Exception:
            pass
    dm.insertByName(name, lbl)
    return lbl


def _dlg_add_edit(dm, name, x, y, w, h):
    edit = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
    edit.Name = name
    edit.PositionX = x
    edit.PositionY = y
    edit.Width = w
    edit.Height = h
    dm.insertByName(name, edit)
    return edit


def _dlg_add_button(dm, name, label, x, y, w, h, push=0):
    btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    btn.Name = name
    btn.Label = label
    btn.PositionX = x
    btn.PositionY = y
    btn.Width = w
    btn.Height = h
    if push:
        try:
            btn.PushButtonType = push
        except Exception:
            pass
    if push == 1:
        try:
            btn.DefaultButton = True
        except Exception:
            pass
    dm.insertByName(name, btn)
    return btn


def _dlg_add_check(dm, name, label, x, y, w, h, state=0):
    chk = dm.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
    chk.Name = name
    chk.Label = label
    chk.PositionX = x
    chk.PositionY = y
    chk.Width = w
    chk.Height = h
    try:
        chk.State = int(state)
    except Exception:
        pass
    dm.insertByName(name, chk)
    return chk


def _pick_xml_file(initial_path=u""):
    try:
        import uno
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sm = ctx.getServiceManager()
        picker = sm.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", ctx)
        picker.setTitle(u"Выбор XML-файла")
        try:
            picker.appendFilter(u"XML (*.xml)", u"*.xml")
            picker.appendFilter(u"Все файлы (*.*)", u"*.*")
            picker.setCurrentFilter(u"XML (*.xml)")
        except Exception:
            pass
        start = unicode(initial_path or u"").strip()
        if start:
            try:
                start = os.path.abspath(os.path.expanduser(os.path.expandvars(start)))
            except Exception:
                pass
            d = start if os.path.isdir(start) else os.path.dirname(start)
            if d and os.path.isdir(d):
                try:
                    picker.setDisplayDirectory(uno.systemPathToFileUrl(d))
                except Exception:
                    pass
            if os.path.isfile(start):
                try:
                    picker.setDefaultName(os.path.basename(start))
                except Exception:
                    pass
        if int(picker.execute()) != 1:
            return None
        urls = picker.getFiles()
        if not urls:
            return None
        return os.path.abspath(uno.fileUrlToSystemPath(urls[0]))
    except Exception as err:
        _log(u"FilePicker: %s" % err)
        return None


def _show_import_dialog(doc, default_xml=u"", default_sheet=u""):
    """Диалог: путь XML + имя листа (+ опционально точки входа). Возвращает dict или None."""
    ctx = XSCRIPTCONTEXT.getComponentContext()
    sm = ctx.getServiceManager()
    toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dw = 460
    m = 10
    btn_h = 16
    dm.PositionX = 120
    dm.PositionY = 80
    dm.Width = dw
    dm.Title = u"Импорт XML"
    y = m
    hint = (
        u"Разбор как в сборе книг: SpreadsheetML или иерархический flatten.\n"
        u"Пустой «Лист» — новый лист; если текущий абсолютно пуст — выгрузка на него.\n"
        u"«Точки входа» — имена Worksheet / xpath-теги через запятую (пусто = весь документ)."
    )
    _dlg_add_fixed(dm, "HintLbl", hint, m, y, dw - m * 2, 42, multiline=True)
    y += 46
    _dlg_add_fixed(dm, "PathLbl", u"Файл XML:", m, y, dw - m * 2, 12)
    y += 14
    path_w = dw - m * 2 - 70
    _dlg_add_edit(dm, "PathEd", m, y, path_w, 16)
    _dlg_add_button(dm, "BrowseBtn", u"Обзор…", m + path_w + 6, y, 64, 16)
    y += 22
    _dlg_add_fixed(dm, "SheetLbl", u"Лист выгрузки (пусто = авто):", m, y, dw - m * 2, 12)
    y += 14
    _dlg_add_edit(dm, "SheetEd", m, y, dw - m * 2, 16)
    y += 22
    _dlg_add_fixed(dm, "EntryLbl", u"Точки входа / листы XML (необязательно):", m, y, dw - m * 2, 12)
    y += 14
    _dlg_add_edit(dm, "EntryEd", m, y, dw - m * 2, 16)
    y += 22
    _dlg_add_check(
        dm,
        "ConvertChk",
        u"Текст → число/дата (как MERGE_XML_CONVERT_TO_NUMBERS)",
        m,
        y,
        dw - m * 2,
        14,
        state=0,
    )
    y += 22
    dh = y + btn_h + m + 8
    dm.Height = dh
    _dlg_add_button(dm, "OkButton", u"OK", m, dh - m - btn_h, 80, btn_h, push=1)
    _dlg_add_button(dm, "CancelButton", u"Отмена", m + 88, dh - m - btn_h, 80, btn_h, push=2)
    dlg = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dlg.setModel(dm)
    try:
        parent = None
        if doc is not None:
            parent = doc.getCurrentController().getFrame().getContainerWindow()
        dlg.createPeer(toolkit, parent)
    except Exception:
        dlg.createPeer(toolkit, None)
    try:
        dlg.getControl("PathEd").setText(unicode(default_xml or u""))
    except Exception:
        pass
    try:
        dlg.getControl("SheetEd").setText(unicode(default_sheet or u""))
    except Exception:
        pass
    try:
        dlg.getControl("ConvertChk").setState(0)
    except Exception:
        pass

    # Browse listener
    try:
        from com.sun.star.awt import XActionListener
        import unohelper

        class _BrowseListener(unohelper.Base, XActionListener):
            def actionPerformed(self, ev):
                cur = u""
                try:
                    cur = unicode(dlg.getControl("PathEd").getText() or u"")
                except Exception:
                    cur = u""
                picked = _pick_xml_file(cur)
                if picked:
                    try:
                        dlg.getControl("PathEd").setText(picked)
                    except Exception:
                        pass

            def disposing(self, ev):
                pass

        dlg.getControl("BrowseBtn").addActionListener(_BrowseListener())
    except Exception as err:
        _log(u"browse listener: %s" % err)

    if int(dlg.execute()) != 1:
        try:
            dlg.dispose()
        except Exception:
            pass
        return None
    path = u""
    sheet = u""
    entries = u""
    convert_to_numbers = False
    try:
        path = unicode(dlg.getControl("PathEd").getText() or u"").strip()
    except Exception:
        pass
    try:
        sheet = unicode(dlg.getControl("SheetEd").getText() or u"").strip()
    except Exception:
        pass
    try:
        entries = unicode(dlg.getControl("EntryEd").getText() or u"").strip()
    except Exception:
        pass
    try:
        convert_to_numbers = bool(int(dlg.getControl("ConvertChk").getState()) == 1)
    except Exception:
        convert_to_numbers = False
    try:
        dlg.dispose()
    except Exception:
        pass
    specs = None
    if entries:
        parts = []
        for piece in re.split(r"[,;]", entries):
            piece = piece.strip()
            if piece:
                parts.append(piece)
        if parts:
            specs = parts
    return {
        "path": path,
        "sheet": sheet,
        "sheet_specs": specs,
        "convert_to_numbers": convert_to_numbers,
    }


def _sheet_is_absolutely_empty(sheet):
    """True, если на листе нет ни одной непустой ячейки в used-area."""
    if sheet is None:
        return True
    try:
        cursor = sheet.createCursor()
        cursor.gotoStartOfUsedArea(False)
        cursor.gotoEndOfUsedArea(True)
        addr = cursor.getRangeAddress()
        sc = int(addr.StartColumn)
        sr = int(addr.StartRow)
        ec = int(addr.EndColumn)
        er = int(addr.EndRow)
    except Exception:
        return True
    r = sr
    while r <= er:
        c = sc
        while c <= ec:
            try:
                cell = sheet.getCellByPosition(c, r)
                if unicode(cell.getString() or u"").strip() != u"":
                    return False
                try:
                    if float(cell.getValue()) != 0.0:
                        # пустая числовая тоже 0; смотрим тип
                        t = int(cell.getType())
                        # 0 EMPTY, 1 VALUE, 2 TEXT, 3 FORMULA — приблизительно
                        if t != 0:
                            return False
                except Exception:
                    pass
            except Exception:
                pass
            c += 1
        r += 1
    return True


def _unique_sheet_name(doc, base):
    name = unicode(base or u"XML").strip() or u"XML"
    for ch in u"[]:*/?\\":
        name = name.replace(ch, u"_")
    name = name[:31]
    sheets = doc.getSheets()
    if not sheets.hasByName(name):
        return name
    i = 2
    while i < 1000:
        cand = (u"%s_%d" % (name[:27], i))[:31]
        if not sheets.hasByName(cand):
            return cand
        i += 1
    return (name[:20] + u"_new")[:31]


def _resolve_target_sheet(doc, preferred_name, xml_sheet_name, is_first):
    """
    Выбрать/создать лист выгрузки.
    preferred_name — из диалога (может быть пустым).
    """
    sheets = doc.getSheets()
    active = None
    try:
        active = doc.getCurrentController().getActiveSheet()
    except Exception:
        active = None

    if preferred_name:
        if is_first:
            if sheets.hasByName(preferred_name):
                return sheets.getByName(preferred_name), preferred_name, False
            new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
            sheets.insertByName(preferred_name, new_sheet)
            return new_sheet, preferred_name, True
        # последующие таблицы XML при заданном имени
        name = _unique_sheet_name(doc, u"%s_%s" % (preferred_name, xml_sheet_name or u"XML"))
        new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
        sheets.insertByName(name, new_sheet)
        return new_sheet, name, True

    # имя не указано
    if is_first and active is not None and _sheet_is_absolutely_empty(active):
        return active, active.getName(), False

    base = unicode(xml_sheet_name or u"XML").strip() or u"XML"
    name = _unique_sheet_name(doc, base)
    new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
    sheets.insertByName(name, new_sheet)
    return new_sheet, name, True


def _write_matrix(sheet, headers, rows):
    """Записать заголовки + строки через setDataArray."""
    width = len(headers) if headers else 0
    if width <= 0:
        return 0, 0
    matrix = []
    matrix.append([unicode(h) if h is not None else u"" for h in headers])
    for row in rows or []:
        rr = [unicode(c) if c is not None else u"" for c in row]
        while len(rr) < width:
            rr.append(u"")
        if len(rr) > width:
            rr = rr[:width]
        matrix.append(rr)
    height = len(matrix)
    rng = sheet.getCellRangeByPosition(0, 0, width - 1, height - 1)
    rng.setDataArray(tuple(tuple(r) for r in matrix))
    return width - 1, height - 1


def _inch_to_hmm(inches):
    return int(round(float(inches) * 2540.0))


def _try_parse_number(text):
    s = unicode(text or u"").strip()
    if s == u"":
        return None
    # убрать пробелы-разделители тысяч
    s2 = s.replace(u"\u00a0", u"").replace(u" ", u"")
    if re.match(r"^-?\d{1,3}(\.\d{3})+(,\d+)?$", s2):
        s2 = s2.replace(u".", u"").replace(u",", u".")
    elif u"," in s2 and u"." not in s2:
        s2 = s2.replace(u",", u".")
    try:
        return float(s2)
    except (TypeError, ValueError):
        return None


def _try_parse_date(text):
    s = unicode(text or u"").strip()
    if not s:
        return None
    # DD.MM.YYYY / DD.MM.YY
    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$", s)
    if not m:
        m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        else:
            return None
    else:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
    try:
        from datetime import date
        return date(y, mo, d)
    except Exception:
        return None


def _uno_date(d):
    """python date → UNO Date struct / serial via setValue as days since 1899-12-30."""
    try:
        from datetime import date
        epoch = date(1899, 12, 30)
        return float((d - epoch).days)
    except Exception:
        return None


def _convert_numbers_dates(sheet, end_col, end_row, header_row=0, enabled=False):
    """Как MERGE_XML_CONVERT_TO_NUMBERS: текст → число/дата при ≥3 удачных в столбце."""
    if not enabled:
        return
    if end_row <= header_row or end_col < 0:
        return
    c = 0
    while c <= end_col:
        nums = 0
        dates = 0
        total = 0
        parsed = []
        r = header_row + 1
        while r <= end_row:
            cell = sheet.getCellByPosition(c, r)
            txt = unicode(cell.getString() or u"").strip()
            if txt == u"":
                parsed.append((r, None, None))
                r += 1
                continue
            total += 1
            dt = _try_parse_date(txt)
            if dt is not None:
                dates += 1
                parsed.append((r, "date", dt))
            else:
                num = _try_parse_number(txt)
                if num is not None:
                    nums += 1
                    parsed.append((r, "num", num))
                else:
                    parsed.append((r, None, None))
            r += 1
        kind = None
        if dates >= 3 and dates >= nums:
            kind = "date"
        elif nums >= 3:
            kind = "num"
        if kind:
            for r, k, val in parsed:
                if k != kind or val is None:
                    continue
                cell = sheet.getCellByPosition(c, r)
                if kind == "num":
                    cell.setValue(float(val))
                else:
                    serial = _uno_date(val)
                    if serial is not None:
                        cell.setValue(serial)
                        try:
                            cell.NumberFormat = 37  # fallback date-ish; LO may remap
                        except Exception:
                            pass
        c += 1


def _apply_default_formatting(doc, sheet, end_col, end_row, header_row=0):
    """Оформление как apply_result_sheet_formatting (встроенные шаги без постобработки)."""
    if sheet is None or end_col < 0 or end_row < header_row:
        return
    # шрифт данных
    try:
        if end_row > header_row:
            data_rng = sheet.getCellRangeByPosition(0, header_row + 1, end_col, end_row)
        else:
            data_rng = sheet.getCellRangeByPosition(0, header_row, end_col, end_row)
        data_rng.CharFontName = _Store.FONT_NAME
        data_rng.CharHeight = float(_Store.FONT_SIZE)
    except Exception:
        pass
    # заголовок
    try:
        header_rng = sheet.getCellRangeByPosition(0, header_row, end_col, header_row)
        header_rng.IsCellBackgroundTransparent = False
        header_rng.CellBackColor = int(_Store.HEADER_BG)
        header_rng.CharFontName = _Store.FONT_NAME
        header_rng.CharHeight = float(_Store.FONT_SIZE)
        header_rng.CharWeight = int(_Store.CHAR_WEIGHT_BOLD)
        try:
            header_rng.HoriJustify = 2  # CENTER
            header_rng.VertJustify = 2
        except Exception:
            pass
        try:
            header_rng.IsTextWrapped = True
        except Exception:
            pass
    except Exception:
        pass
    # автоширина + pad/cap
    try:
        columns = sheet.getColumns()
        pad_w = _inch_to_hmm(_Store.COL_PAD_INCH)
        max_w = _inch_to_hmm(_Store.COL_MAX_INCH)
        c = 0
        while c <= end_col:
            col = columns.getByIndex(c)
            col.OptimalWidth = True
            w = int(col.Width) + pad_w
            if w > max_w:
                w = max_w
            col.OptimalWidth = False
            col.Width = w
            c += 1
    except Exception as err:
        _log(u"autofit: %s" % err)
    # smart-table / автофильтр TableStyleLight1
    try:
        data_range = sheet.getCellRangeByPosition(0, header_row, end_col, end_row)
        _apply_smart_table(doc, sheet, data_range)
    except Exception as err:
        _log(u"smart-table: %s" % err)
    # закрепить заголовок
    if _Store.FREEZE_HEADER:
        try:
            ctrl = doc.getCurrentController()
            ctrl.setActiveSheet(sheet)
            ctrl.freezeAtPosition(0, header_row + 1)
        except Exception as err:
            _log(u"freeze: %s" % err)


def _safe_db_name(sheet):
    raw = unicode(getattr(sheet, "Name", u"") or u"Sheet")
    out = []
    for ch in raw:
        if ch.isalnum() or ch in u"_":
            out.append(ch)
        else:
            out.append(u"_")
    name = u"".join(out).strip(u"_") or u"Sheet"
    return (u"AF_XML_" + name)[:60]


def _apply_smart_table(doc, sheet, data_range):
    """DatabaseRange + AutoFilter + TableStyleLight1 (как в основном макросе)."""
    if doc is None or sheet is None or data_range is None:
        return False
    addr = data_range.getRangeAddress()
    db_name = _safe_db_name(sheet)
    # удалить старый с тем же именем
    try:
        dbs = doc.DatabaseRanges
        if dbs.hasByName(db_name):
            dbs.removeByName(db_name)
    except Exception:
        pass
    try:
        doc.DatabaseRanges.addNewByName(db_name, addr)
        db = doc.DatabaseRanges.getByName(db_name)
    except Exception as err:
        _log(u"DatabaseRanges: %s" % err)
        return False
    try:
        db.AutoFilter = True
    except Exception:
        try:
            db.setPropertyValue("AutoFilter", True)
        except Exception:
            pass

    def _set(prop, value):
        try:
            setattr(db, prop, value)
            return True
        except Exception:
            pass
        try:
            db.setPropertyValue(prop, value)
            return True
        except Exception:
            return False

    _set("TableStyleName", _Store.TABLE_STYLE)
    _set("UseRowStripes", bool(_Store.USE_ROW_STRIPES))
    _set("ContainsHeader", True)
    try:
        doc.calculateAll()
    except Exception:
        pass
    return True


def _default_sample_xml_near_doc(doc):
    """Рядом с .ods / в standalone — sample_employees.xml."""
    try:
        import uno
        url = unicode(doc.getURL() or u"")
        if url:
            path = uno.fileUrlToSystemPath(url)
            folder = os.path.dirname(path)
            cand = os.path.join(folder, "sample_employees.xml")
            if os.path.isfile(cand):
                return cand
    except Exception:
        pass
    here = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else u""
    if here:
        cand = os.path.join(here, "sample_employees.xml")
        if os.path.isfile(cand):
            return cand
    return u""


def import_xml(*args):
    """Точка входа: диалог → разбор XML → лист → оформление."""
    _reset_log()
    doc = _get_document()
    if doc is None:
        _log(u"ошибка: документ Calc не найден")
        return
    default_xml = _default_sample_xml_near_doc(doc)
    params = _show_import_dialog(doc, default_xml=default_xml)
    if params is None:
        _log(u"отмена")
        return
    path = unicode(params.get("path") or u"").strip()
    if not path:
        _msgbox(doc, u"Укажите путь к XML-файлу.")
        return
    path = os.path.abspath(os.path.expanduser(os.path.expandvars(path)))
    if not os.path.isfile(path):
        _msgbox(doc, u"Файл не найден:\n%s" % path)
        return
    if not is_xml_source_path(path):
        _log(u"предупреждение: расширение не .xml — пробуем разобрать")
    _log(u"=== импорт: %s ===" % path)
    sheets_data = xml_source_read_sheets(path, sheet_specs=params.get("sheet_specs"))
    if not sheets_data:
        _msgbox(doc, u"Не удалось разобрать XML или нет табличных данных:\n%s" % path)
        return
    preferred = unicode(params.get("sheet") or u"").strip()
    convert_to_numbers = bool(params.get("convert_to_numbers"))
    _log(u"convert_to_numbers=%s" % convert_to_numbers)
    written = []
    i = 0
    while i < len(sheets_data):
        xml_name, headers, rows = sheets_data[i]
        sheet, name, created = _resolve_target_sheet(doc, preferred, xml_name, is_first=(i == 0))
        _log(u"лист «%s» (xml=%s, создан=%s, строк=%d)" % (name, xml_name, created, len(rows or [])))
        # очистить целевой лист при повторном использовании существующего
        if not created:
            try:
                cursor = sheet.createCursor()
                cursor.gotoStartOfUsedArea(False)
                cursor.gotoEndOfUsedArea(True)
                cursor.clearContents(1 | 2 | 4 | 8 | 16 | 32 | 64 | 128 | 256 | 512)
            except Exception:
                pass
            # снять старый AF
            try:
                db_name = _safe_db_name(sheet)
                if doc.DatabaseRanges.hasByName(db_name):
                    doc.DatabaseRanges.removeByName(db_name)
            except Exception:
                pass
        end_col, end_row = _write_matrix(sheet, headers, rows)
        _convert_numbers_dates(
            sheet, end_col, end_row, header_row=0, enabled=convert_to_numbers
        )
        _apply_default_formatting(doc, sheet, end_col, end_row, header_row=0)
        written.append(name)
        try:
            doc.getCurrentController().setActiveSheet(sheet)
        except Exception:
            pass
        i += 1
    _msgbox(
        doc,
        u"Импорт завершён.\nФайл: %s\nЛисты: %s" % (path, u", ".join(written)),
    )
    _log(u"готово: %s" % u", ".join(written))


g_exportedScripts = (import_xml,)
