# -*- coding: utf-8 -*-
from __future__ import print_function
"""
Наименование: context_menu

Описание: установка пунктов меню макросов Calc: контекстное меню ячейки (cell.xml)
и меню «Сервис» (menubar.xml → .uno:ToolsMenu).

Назначение: сканирует соседние .py макросы с CONTEXT_CELL_MENU, показывает диалог выбора
и записывает пункты в cell.xml и menubar.xml профиля LibreOffice / AlterOffice.
Точка входа: setContextMenu().
"""

MACRO_VERSION = "3.10.696"
import ast
import os
import sys
import xml.etree.ElementTree as ET

import uno
import unohelper
from com.sun.star.awt import XActionListener

# Константы в pythonpath: AlterOffice 2026 вырезает модульные присваивания в .py-скрипте.
from libre_macros_context_menu_cfg import DEFAULT_CELL_POPUP_XML as _DEFAULT_CELL_POPUP_XML, DLG_BTN_H as _DLG_BTN_H, DLG_BTN_W as _DLG_BTN_W, DLG_GAP as _DLG_GAP, DLG_MARGIN as _DLG_MARGIN, DLG_ROW_H as _DLG_ROW_H, DLG_W as _DLG_W, MENU_NS, REL_MENUBAR as _REL_MENUBAR, REL_MENUBAR_AO as _REL_MENUBAR_AO, REL_POPUP as _REL_POPUP, REL_POPUP_AO as _REL_POPUP_AO, SKIP_SCAN as _SKIP_SCAN, TARGET_CELL, TARGET_TOOLS, TOOLS_MENU_ID as _TOOLS_MENU_ID
import libre_macros_context_menu_cfg as _cm_cfg


def _menu_tag(local):
    return "{%s}%s" % (MENU_NS, local)


def _init_menu_tags():
    if _cm_cfg.tag_menu is None:
        _cm_cfg.tag_menu = _menu_tag("menu")
        _cm_cfg.tag_menuitem = _menu_tag("menuitem")
        _cm_cfg.tag_menupopup = _menu_tag("menupopup")
        _cm_cfg.tag_menuseparator = _menu_tag("menuseparator")
        _cm_cfg.tag_id = _menu_tag("id")
        _cm_cfg.tag_label = _menu_tag("label")


def _submenu_menu_id(level):
    """id подменю = явный id из spec или текст label (как в cell.xml: id=\"Python\")."""
    sid = level.get("id")
    if sid not in (None, ""):
        return str(sid)
    return str(level.get("value", "Menu"))


def parse_context_level(spec):
    parts = {}
    chunks = str(spec or "").split(";")
    head = chunks[0] if chunks else ""
    if "#" in head:
        kind, val = head.split("#", 1)
        parts["kind"] = kind.strip().lower()
        parts["value"] = val.strip()
    else:
        parts["kind"] = str(spec or "").strip().lower()
        parts["value"] = ""
    ci = 1
    while ci < len(chunks):
        piece = chunks[ci]
        if "=" in piece:
            k, v = piece.split("=", 1)
        elif "#" in piece:
            k, v = piece.split("#", 1)
        else:
            ci = ci + 1
            continue
        parts[k.strip().lower()] = v.strip()
        ci = ci + 1
    return parts


def parse_context_menu_spec(spec_seq):
    levels = []
    if spec_seq is None:
        return levels
    if isinstance(spec_seq, dict):
        keys = sorted(spec_seq.keys(), key=lambda k: int(k) if str(k).isdigit() else k)
        raw = [spec_seq[k] for k in keys]
    else:
        raw = list(spec_seq)
    i = 0
    while i < len(raw):
        levels.append(parse_context_level(raw[i]))
        i = i + 1
    return levels


def _script_url(macro_file, function_name):
    mf = str(macro_file or "").strip()
    if mf != "" and not mf.endswith(".py"):
        mf = mf + ".py"
    fn = str(function_name or "").strip()
    return "vnd.sun.star.script:%s$%s?language=Python&location=user" % (mf, fn)


def _split_menu_levels(levels):
    """Разделить цепочку submenu и остальные уровни (run_function, separator)."""
    submenu_levels = []
    content_levels = []
    i = 0
    while i < len(levels):
        lv = levels[i]
        kind = str(lv.get("kind", "")).strip().lower()
        if kind == "submenu":
            submenu_levels.append(lv)
        else:
            content_levels.append(lv)
        i = i + 1
    return submenu_levels, content_levels


def _is_run_function_level(level):
    kind = str(level.get("kind", "")).strip().lower()
    return kind in ("run_function", "run", "function")


def _is_separator_level(level):
    kind = str(level.get("kind", "")).strip().lower()
    return kind in ("separator", "sep", "menuseparator")


def _level_order(level, default=10000):
    """
    Явный порядок пункта в подменю: order#N / order=N.
    Меньше число — выше в меню. Без order — default (обычно 10000 = в конце).
    Возвращает int или None, если default is None и order не задан.
    """
    if level is None:
        if default is None:
            return None
        return int(default)
    raw = level.get("order")
    if raw in (None, ""):
        if default is None:
            return None
        return int(default)
    try:
        return int(str(raw).strip())
    except Exception:
        if default is None:
            return None
        return int(default)


def _separator_placement(level):
    """
    Позиция разделителя относительно соседних пунктов:
    before/top — перед следующим run_function;
    after/bottom — после предыдущего run_function.
    """
    raw = str(level.get("value") or level.get("position") or "before").strip().lower()
    if raw in ("after", "bottom", "post"):
        return "after"
    return "before"


def _menupopup_last_is_separator(menupopup):
    _init_menu_tags()
    if menupopup is None:
        return False
    children = list(menupopup)
    if len(children) == 0:
        return False
    return children[-1].tag == _cm_cfg.tag_menuseparator


def _append_menu_separator(menupopup):
    """Добавить menu:menuseparator, если последний элемент ещё не разделитель."""
    if menupopup is None:
        return False
    if _menupopup_last_is_separator(menupopup):
        return False
    _init_menu_tags()
    menupopup.append(ET.Element(_cm_cfg.tag_menuseparator))
    return True


def _build_menu_item(run_level, macro_file):
    if run_level is None:
        return None, None
    fn = run_level.get("value", run_level.get("function", ""))
    mod = run_level.get("module", macro_file)
    title = run_level.get("display_name", fn)
    url = _script_url(mod, fn)
    _init_menu_tags()
    item = ET.Element(_cm_cfg.tag_menuitem)
    item.set(_cm_cfg.tag_id, url)
    item.set(_cm_cfg.tag_label, str(title))
    return item, url


def _get_menupopup(menu_el):
    _init_menu_tags()
    popup = menu_el.find(_cm_cfg.tag_menupopup)
    if popup is None:
        popup = ET.SubElement(menu_el, _cm_cfg.tag_menupopup)
    return popup


def _find_child_submenu(container_popup, level):
    """Найти существующее подменю по id или label."""
    _init_menu_tags()
    label = str(level.get("value", "Menu"))
    menu_id = _submenu_menu_id(level)
    for child in list(container_popup):
        if child.tag != _cm_cfg.tag_menu:
            continue
        cid = child.get(_cm_cfg.tag_id, "")
        clabel = child.get(_cm_cfg.tag_label, "")
        if cid == menu_id or clabel == label or cid == label:
            return child
    return None


def _ensure_submenu_path(root_popup, submenu_levels):
    """
    Пройти/создать цепочку подменю внутри root_popup.
    Возвращает menupopup, куда добавлять пункт макроса.
    """
    _init_menu_tags()
    current = root_popup
    i = 0
    while i < len(submenu_levels):
        lv = submenu_levels[i]
        label = str(lv.get("value", "Menu"))
        menu_id = _submenu_menu_id(lv)
        found = _find_child_submenu(current, lv)
        if found is None:
            found = ET.Element(_cm_cfg.tag_menu)
            found.set(_cm_cfg.tag_id, menu_id)
            found.set(_cm_cfg.tag_label, label)
            ET.SubElement(found, _cm_cfg.tag_menupopup)
            current.append(found)
        current = _get_menupopup(found)
        i = i + 1
    return current


def _find_menuitem_by_url(root, script_url):
    _init_menu_tags()
    for el in root.iter():
        if el.tag == _cm_cfg.tag_menuitem and el.get(_cm_cfg.tag_id, "") == script_url:
            return el
    return None


def _item_in_menupopup(menupopup, script_url):
    _init_menu_tags()
    for child in menupopup:
        if child.tag == _cm_cfg.tag_menuitem and child.get(_cm_cfg.tag_id, "") == script_url:
            return child
    return None


def _remove_menuitem_by_url(root, script_url):
    _init_menu_tags()
    removed = False
    for parent in root.iter():
        if parent.tag != _cm_cfg.tag_menupopup:
            continue
        for child in list(parent):
            if child.tag == _cm_cfg.tag_menuitem and child.get(_cm_cfg.tag_id, "") == script_url:
                parent.remove(child)
                removed = True
    return removed


def _submenu_matches_level(menu_el, level):
    """Совпадение элемента menu с уровнем submenu из CONTEXT_CELL_MENU."""
    _init_menu_tags()
    if menu_el is None or level is None or menu_el.tag != _cm_cfg.tag_menu:
        return False
    menu_id = _submenu_menu_id(level)
    label = str(level.get("value", "Menu"))
    cid = menu_el.get(_cm_cfg.tag_id, "")
    clabel = menu_el.get(_cm_cfg.tag_label, "")
    return cid == menu_id or clabel == label or cid == label


def _first_submenu_level(spec_seq):
    """Первый уровень submenu (верхний пункт в cell.xml)."""
    levels = parse_context_menu_spec(spec_seq)
    submenu_levels, _content = _split_menu_levels(levels)
    if len(submenu_levels) == 0:
        return None
    return submenu_levels[0]


def _submenu_level_key(level):
    if level is None:
        return None
    return (_submenu_menu_id(level), str(level.get("value", "")))


def collect_managed_top_submenus(entries):
    """
    Уникальные подменю верхнего уровня из всех CONTEXT_CELL_MENU в каталоге.

    Используется для полной очистки перед переустановкой выбранных пунктов.
    """
    seen = set()
    out = []
    i = 0
    while i < len(entries or []):
        ent = entries[i]
        i = i + 1
        lv = _first_submenu_level(ent.get("spec"))
        if lv is None:
            continue
        key = _submenu_level_key(lv)
        if key in seen:
            continue
        seen.add(key)
        out.append(lv)
    return out


def _remove_top_level_submenus(root_popup, submenu_levels):
    """Удалить подменю верхнего уровня из корневого menupopup."""
    if root_popup is None or not submenu_levels:
        return 0
    _init_menu_tags()
    removed = 0
    for child in list(root_popup):
        if child.tag != _cm_cfg.tag_menu:
            continue
        bi = 0
        while bi < len(submenu_levels):
            if _submenu_matches_level(child, submenu_levels[bi]):
                root_popup.remove(child)
                removed = removed + 1
                break
            bi = bi + 1
    return removed


def _write_menu_xml_tree(path, tree):
    try:
        ET.register_namespace("menu", MENU_NS)
    except Exception:
        pass
    if hasattr(ET, "indent"):
        ET.indent(tree, space="  ")
    tree.write(path, encoding="UTF-8", xml_declaration=True)


def _write_cell_popup_tree(path, tree):
    """Совместимость с тестами и старым кодом."""
    _write_menu_xml_tree(path, tree)


def _menu_target_label(target_kind):
    if target_kind == TARGET_TOOLS:
        return "menubar.xml (Сервис)"
    return "cell.xml (контекстное меню)"


def _find_menu_by_id(root, menu_id):
    _init_menu_tags()
    if root is None or menu_id in (None, ""):
        return None
    for menu_el in root.iter(_cm_cfg.tag_menu):
        if menu_el.get(_cm_cfg.tag_id) == menu_id:
            return menu_el
    return None


def _default_cell_popup_tree():
    """Шаблон cell.xml (uno-id), если в профиле ещё нет контекстного меню."""
    _init_menu_tags()
    try:
        root = ET.fromstring(_DEFAULT_CELL_POPUP_XML.encode("utf-8"))
    except Exception:
        root = ET.Element(_menu_tag("menupopup"))
    return ET.ElementTree(root), root


def _guess_office_kind(user_dir=None):
    """AlterOffice (aoffice.cfg/acell) или LibreOffice (soffice.cfg/scalc)."""
    try:
        lib = _macro_lib_dir() or ""
        low = lib.replace("\\", "/").lower()
        if "alteroffice" in low or "aoffice" in low:
            return "ao"
    except Exception:
        pass
    if user_dir:
        low = str(user_dir).replace("\\", "/").lower()
        if "alteroffice" in low:
            return "ao"
    return "lo"


def _menu_rel_paths(user_dir, target_kind):
    """Относительные пути cell.xml / menubar.xml с учётом LO и AlterOffice (aoffice.cfg/acell)."""
    rels = []
    prefer_ao = _guess_office_kind(user_dir) == "ao"
    if user_dir:
        cfg = os.path.join(user_dir, "config")
        has_ao = os.path.isdir(os.path.join(cfg, "aoffice.cfg"))
        has_lo = os.path.isdir(os.path.join(cfg, "soffice.cfg"))
        if target_kind == TARGET_CELL:
            if has_ao or prefer_ao:
                rels.append(_REL_POPUP_AO)
            if has_lo or (not prefer_ao and not has_ao):
                rels.append(_REL_POPUP)
        else:
            if has_ao or prefer_ao:
                rels.append(_REL_MENUBAR_AO)
            if has_lo or (not prefer_ao and not has_ao):
                rels.append(_REL_MENUBAR)
    if not rels:
        rels = [
            (_REL_POPUP_AO if prefer_ao else _REL_POPUP)
            if target_kind == TARGET_CELL
            else (_REL_MENUBAR_AO if prefer_ao else _REL_MENUBAR)
        ]
    return rels


def _windows_profile_user_roots():
    """%APPDATA%\\LibreOffice\\N\\user и AlterOffice (Windows)."""
    roots = []
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return roots
    for office_name in ("LibreOffice", "AlterOffice", "AlterOffice3"):
        for ver in ("24", "7", "5", "4", "3"):
            p = os.path.join(appdata, office_name, ver, "user")
            if os.path.isdir(p):
                roots.append(p)
    return roots


def _read_menu_container(path, target_kind=TARGET_CELL):
    """Прочитать XML и вернуть (tree, menupopup-контейнер для установки пунктов)."""
    if not os.path.isfile(path):
        if target_kind == TARGET_CELL:
            return _default_cell_popup_tree()
        return None, None
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        _init_menu_tags()
        if target_kind == TARGET_TOOLS:
            tools_menu = _find_menu_by_id(root, _TOOLS_MENU_ID)
            if tools_menu is None:
                return tree, None
            popup = tools_menu.find(_cm_cfg.tag_menupopup)
            if popup is None:
                popup = ET.SubElement(tools_menu, _cm_cfg.tag_menupopup)
            return tree, popup
        if root.tag == _menu_tag("menupopup"):
            return tree, root
        popup = root.find(".//" + _menu_tag("menupopup"))
        if popup is not None:
            return tree, popup
    except Exception:
        pass
    return None, None


def _read_popup_root(path):
    return _read_menu_container(path, TARGET_CELL)


def purge_managed_submenus_from_menu_xml(path, submenu_levels, target_kind=TARGET_CELL):
    """
    Удалить управляемые подменю верхнего уровня из cell.xml или menubar.xml.

    Возвращает (число удалённых, файл изменён).
    """
    if not submenu_levels:
        return 0, False
    menu_xml = _cell_xml_label(path)
    tree, popup = _read_menu_container(path, target_kind)
    if popup is None:
        if target_kind == TARGET_TOOLS:
            return 0, False
        if not os.path.isfile(path):
            return 0, False
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            popup = ET.Element(_menu_tag("menupopup"))
            tree = ET.ElementTree(popup)
            _write_menu_xml_tree(path, tree)
        except Exception:
            return 0, False
        return 0, False
    removed = _remove_top_level_submenus(popup, submenu_levels)
    if removed == 0:
        return 0, False
    try:
        _write_menu_xml_tree(path, tree)
        return removed, True
    except Exception:
        return 0, False


def purge_managed_submenus_from_cell_xml(path, submenu_levels):
    """Совместимость: очистка только cell.xml."""
    return purge_managed_submenus_from_menu_xml(path, submenu_levels, TARGET_CELL)


def purge_managed_submenus_all_paths(submenu_levels, paths=None, target_kind=TARGET_CELL):
    """Очистить подменю во всех найденных XML целевого типа. Возвращает список сообщений."""
    msgs = []
    menu_paths = paths if paths is not None else menu_config_paths(target_kind)
    target_label = _menu_target_label(target_kind)
    if len(menu_paths) == 0:
        return ["%s не найден" % target_label]
    pi = 0
    while pi < len(menu_paths):
        path = menu_paths[pi]
        removed, saved = purge_managed_submenus_from_menu_xml(
            path, submenu_levels, target_kind
        )
        if saved:
            msgs.append(
                "%s [%s]: удалено подменю %d"
                % (_cell_xml_label(path), target_label, removed)
            )
        pi = pi + 1
    if len(msgs) == 0:
        labels = [str(lv.get("value", "")) for lv in submenu_levels]
        msgs.append(
            "%s — подменю не найдено (%s)"
            % (
                target_label,
                ", ".join(labels) if labels else "—",
            )
        )
    return msgs


def _submenu_path_label(submenu_levels):
    parts = [str(lv.get("value", "")) for lv in submenu_levels]
    return " → ".join([p for p in parts if p != ""])


def _entry_display_name(spec, macro_file, key):
    levels = parse_context_menu_spec(spec)
    labels = []
    pending_before = False
    i = 0
    while i < len(levels):
        lv = levels[i]
        if _is_separator_level(lv):
            pending_before = _separator_placement(lv) == "before"
            i = i + 1
            continue
        if _is_run_function_level(lv):
            dn = lv.get("display_name")
            if dn not in (None, ""):
                labels.append(str(dn))
            pending_before = False
        i = i + 1
    if len(labels) > 0:
        if len(labels) == 1:
            return labels[0]
        return ", ".join(labels)
    if key not in (None, ""):
        return str(key)
    return os.path.splitext(os.path.basename(macro_file))[0]


def _submenu_path_key(submenu_levels):
    """Ключ цепочки подменю для группировки пунктов из разных .py."""
    parts = []
    i = 0
    while i < len(submenu_levels):
        lv = submenu_levels[i]
        parts.append("%s\0%s" % (_submenu_menu_id(lv), str(lv.get("value", ""))))
        i = i + 1
    return "\n".join(parts)


def _expand_content_to_ordered_units(content_levels, macro_file, source_index=0):
    """
    Развернуть content_levels в упорядоченные единицы установки.

    Каждая единица: dict с type ('item'|'separator'), order, tie, level, macro_file.
    separator#before без order привязывается к следующему run_function (tie=-1);
    separator#after без order — к предыдущему (tie=+1).
    Явный order# у separator задаёт его место независимо.
    """
    units = []
    pending_before = False
    seq = 0
    i = 0
    while i < len(content_levels):
        lv = content_levels[i]
        if _is_separator_level(lv):
            place = _separator_placement(lv)
            sep_order = _level_order(lv, default=None)
            if sep_order is not None:
                units.append(
                    {
                        "type": "separator",
                        "order": int(sep_order),
                        "tie": -1 if place == "before" else 1,
                        "seq": seq,
                        "source_index": int(source_index),
                        "macro_file": macro_file,
                        "level": lv,
                    }
                )
                seq = seq + 1
            elif place == "before":
                pending_before = True
            else:
                j = len(units) - 1
                while j >= 0:
                    if units[j].get("type") == "item":
                        units.append(
                            {
                                "type": "separator",
                                "order": int(units[j]["order"]),
                                "tie": 1,
                                "seq": seq,
                                "source_index": int(source_index),
                                "macro_file": macro_file,
                                "level": lv,
                            }
                        )
                        seq = seq + 1
                        break
                    j = j - 1
            i = i + 1
            continue
        if not _is_run_function_level(lv):
            i = i + 1
            continue
        item_order = _level_order(lv, default=10000)
        if pending_before:
            units.append(
                {
                    "type": "separator",
                    "order": int(item_order),
                    "tie": -1,
                    "seq": seq,
                    "source_index": int(source_index),
                    "macro_file": macro_file,
                    "level": None,
                }
            )
            seq = seq + 1
            pending_before = False
        units.append(
            {
                "type": "item",
                "order": int(item_order),
                "tie": 0,
                "seq": seq,
                "source_index": int(source_index),
                "macro_file": macro_file,
                "level": lv,
            }
        )
        seq = seq + 1
        i = i + 1
    return units


def _sort_menu_units(units):
    """Сортировка: order ↑, tie (sep before / item / sep after), source_index, seq."""
    def key_fn(u):
        return (
            int(u.get("order", 10000)),
            int(u.get("tie", 0)),
            int(u.get("source_index", 0)),
            int(u.get("seq", 0)),
            str(u.get("macro_file", "")),
        )

    return sorted(list(units), key=key_fn)


def _install_ordered_units(target_popup, units):
    """Записать отсортированные единицы (item/separator) в menupopup."""
    installed = 0
    skipped = 0
    messages = []
    i = 0
    while i < len(units):
        u = units[i]
        if u.get("type") == "separator":
            if _append_menu_separator(target_popup):
                installed = installed + 1
            i = i + 1
            continue
        lv = u.get("level")
        macro_file = u.get("macro_file")
        menu_item, script_url = _build_menu_item(lv, macro_file)
        if menu_item is None or script_url is None:
            i = i + 1
            continue
        item_title = str(lv.get("display_name", lv.get("value", "")))
        if _item_in_menupopup(target_popup, script_url) is not None:
            skipped = skipped + 1
            messages.append("%s: уже в меню" % item_title)
        else:
            migrated = _find_menuitem_by_url(target_popup, script_url) is not None
            if migrated:
                _remove_menuitem_by_url(target_popup, script_url)
            target_popup.append(menu_item)
            installed = installed + 1
            messages.append("%s: добавлено (order=%s)" % (item_title, u.get("order")))
        i = i + 1
    return installed, skipped, messages


def _install_menu_content_levels(target_popup, content_levels, macro_file):
    """
    Установить в menupopup пункты run_function и separator из CONTEXT_CELL_MENU.

    Порядок внутри подменю: явный order#N у run_function / separator
    (меньше — выше). Без order — в конце (10000), с сохранением порядка в файле.

    separator#top / #before — перед следующим пунктом;
    separator#bottom / #after — после предыдущего пункта.
    Разделители не дублируются подряд.
    """
    units = _expand_content_to_ordered_units(content_levels, macro_file, source_index=0)
    units = _sort_menu_units(units)
    return _install_ordered_units(target_popup, units)


def collect_menu_install_groups(entries):
    """
    Собрать пункты выбранных макросов, сгруппированные по цепочке submenu.

    Возвращает список dict:
      submenu_levels, branch, units (уже отсортированы по order).
    """
    groups = {}
    group_order = []
    si = 0
    while si < len(entries):
        ent = entries[si]
        levels = parse_context_menu_spec(ent.get("spec"))
        submenu_levels, content_levels = _split_menu_levels(levels)
        run_levels = [lv for lv in content_levels if _is_run_function_level(lv)]
        if len(run_levels) == 0:
            si = si + 1
            continue
        key = _submenu_path_key(submenu_levels)
        if key not in groups:
            groups[key] = {
                "submenu_levels": submenu_levels,
                "branch": _submenu_path_label(submenu_levels),
                "units": [],
            }
            group_order.append(key)
        units = _expand_content_to_ordered_units(
            content_levels, ent.get("file"), source_index=si
        )
        groups[key]["units"].extend(units)
        si = si + 1
    result = []
    gi = 0
    while gi < len(group_order):
        g = groups[group_order[gi]]
        g["units"] = _sort_menu_units(g["units"])
        result.append(g)
        gi = gi + 1
    return result


def install_menu_groups(groups, target_kind=TARGET_CELL, paths=None):
    """Установить группы пунктов (уже с order) во все профили target_kind."""
    installed = 0
    skipped = 0
    messages = []
    target_label = _menu_target_label(target_kind)
    menu_paths = paths if paths is not None else menu_config_paths(target_kind)
    if len(menu_paths) == 0:
        return 0, 0, [
            "Не найден %s. Каталог макросов: %s"
            % (target_label, os.path.abspath(_macro_lib_dir()))
        ]
    if groups is None or len(groups) == 0:
        return 0, 0, ["Нет пунктов для установки."]

    for path in menu_paths:
        menu_xml = _cell_xml_label(path)
        tree, popup = _read_menu_container(path, target_kind)
        try:
            if popup is None:
                if target_kind == TARGET_TOOLS:
                    messages.append(
                        "%s: меню .uno:ToolsMenu не найдено" % menu_xml
                    )
                    continue
                os.makedirs(os.path.dirname(path), exist_ok=True)
                tree, popup = _default_cell_popup_tree()
            path_inst = 0
            path_skip = 0
            gi = 0
            while gi < len(groups):
                g = groups[gi]
                target_popup = _ensure_submenu_path(popup, g["submenu_levels"])
                inst_n, skip_n, item_msgs = _install_ordered_units(
                    target_popup, g["units"]
                )
                path_inst = path_inst + inst_n
                path_skip = path_skip + skip_n
                branch = g.get("branch") or "меню"
                if inst_n > 0 or skip_n > 0:
                    messages.append(
                        "%s [%s / %s]: записано %d, уже было %d"
                        % (menu_xml, target_label, branch, inst_n, skip_n)
                    )
                    messages.extend(item_msgs)
                gi = gi + 1
            installed = installed + path_inst
            skipped = skipped + path_skip
            if path_inst > 0:
                _write_menu_xml_tree(path, tree)
            elif path_inst == 0 and path_skip == 0:
                messages.append(
                    "%s [%s]: нет новых пунктов" % (menu_xml, target_label)
                )
        except Exception as err:
            messages.append("%s [%s]: ошибка — %s" % (menu_xml, target_label, err))
    return installed, skipped, messages


def _macro_lib_dir():
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        if os.path.isdir(here):
            return here
    except Exception:
        pass
    try:
        import sys
        for p in sys.path:
            if p and os.path.isdir(p) and os.path.isfile(os.path.join(p, "context_menu.py")):
                return p
    except Exception:
        pass
    # Windows: проверяем пути AlterOffice/LibreOffice через APPDATA
    if os.name == 'nt' or sys.platform == 'win32':
        appdata = os.environ.get('APPDATA')
        if appdata:
            # AlterOffice на Windows
            for ao_name in ('AlterOffice', 'AlterOffice3'):
                for ver in ('5', '4', '3'):
                    p = os.path.join(appdata, ao_name, ver, 'user', 'Scripts', 'python')
                    if os.path.isdir(p):
                        return p
            # LibreOffice на Windows
            for ver in ('4', '24', '7', '3', '5'):
                p = os.path.join(appdata, 'LibreOffice', ver, 'user', 'Scripts', 'python')
                if os.path.isdir(p):
                    return p
    # Linux/Mac пути по умолчанию
    home = os.path.expanduser("~")
    # AlterOffice Linux
    for ver in ('5', '4', '3'):
        p = os.path.join(home, '.config', 'alteroffice', ver, 'user', 'Scripts', 'python')
        if os.path.isdir(p):
            return p
    # LibreOffice Linux
    return os.path.join(home, '.config', 'libreoffice', '4', 'user', 'Scripts', 'python')


def _user_dir_from_macro_lib(lib_dir=None):
    """
    Каталог user профиля LO из пути Scripts/python.
    .../user/Scripts/python → .../user
    """
    d = os.path.abspath(lib_dir if lib_dir is not None else _macro_lib_dir())
    if os.path.basename(d) == "python":
        scripts = os.path.dirname(d)
        if os.path.basename(scripts) == "Scripts":
            user_dir = os.path.dirname(scripts)
            if os.path.basename(user_dir) == "user":
                return user_dir
    return None


def _uno_user_profile_dir():
    """Профиль user через PathSettings (при запуске из LibreOffice)."""
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sm = ctx.getServiceManager()
        ps = sm.createInstanceWithContext("com.sun.star.util.PathSettings", ctx)
        url = ps.getPropertyValue("UserConfig")
        if url is None or str(url) == "":
            return None
        from urllib.parse import unquote, urlparse

        parsed = urlparse(str(url))
        if parsed.scheme == "file":
            path = unquote(parsed.path)
            if path.endswith("/user") or path.endswith("\\user"):
                return path
            if os.path.basename(path) == "user":
                return path
    except Exception:
        pass
    return None


def parse_context_menu_from_source(path):
    """
    Прочитать CONTEXT_CELL_MENU и CONTEXT_CELL_MENU_KEY из .py (ast, без import).

    Возвращает dict или None.
    """
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            src = fh.read()
    except Exception:
        return None
    try:
        tree = ast.parse(src, filename=path)
    except Exception:
        return None
    key = None
    spec = None
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if not isinstance(target, ast.Name):
                continue
            name = target.id
            try:
                val = ast.literal_eval(node.value)
            except Exception:
                continue
            if name == "CONTEXT_CELL_MENU_KEY":
                key = val
            elif name == "CONTEXT_CELL_MENU":
                spec = val
    if spec is None:
        return None
    macro_file = os.path.basename(path)
    if key in (None, ""):
        key = os.path.splitext(macro_file)[0]
    return {
        "file": macro_file,
        "path": path,
        "key": str(key),
        "spec": spec,
        "label": _entry_display_name(spec, macro_file, key),
    }


def discover_context_menu_macros(lib_dir=None):
    """Все .py в каталоге макросов с CONTEXT_CELL_MENU."""
    base = lib_dir if lib_dir is not None else _macro_lib_dir()
    entries = []
    try:
        names = sorted(os.listdir(base))
    except Exception:
        return entries
    for name in names:
        if not name.endswith(".py"):
            continue
        if name in _SKIP_SCAN:
            continue
        path = os.path.join(base, name)
        if not os.path.isfile(path):
            continue
        item = parse_context_menu_from_source(path)
        if item is not None:
            entries.append(item)
    return entries


def iter_user_profile_roots():
    """Каталог user внутри профиля LO/AlterOffice (Scripts/python, config/soffice.cfg/...)."""
    home = os.path.expanduser("~")
    seen = set()
    roots = []

    def add(p):
        ap = os.path.abspath(p)
        if ap not in seen and os.path.isdir(ap):
            seen.add(ap)
            roots.append(ap)

    for name in ("libreoffice", "alteroffice", "openoffice"):
        for ver in ("4", "24", "7", "3", "5"):
            add(os.path.join(home, ".config", name, ver, "user"))
    add(
        os.path.join(
            home,
            ".var",
            "app",
            "org.libreoffice.LibreOffice",
            "config",
            "libreoffice",
            "4",
            "user",
        )
    )
    add(
        os.path.join(
            home,
            ".var",
            "app",
            "org.documentfoundation.LibreOffice",
            "config",
            "libreoffice",
            "4",
            "user",
        )
    )
    cfg = os.path.join(home, ".config")
    if os.path.isdir(cfg):
        try:
            for entry in os.listdir(cfg):
                if entry in ("libreoffice", "alteroffice", "openoffice"):
                    continue
                base = os.path.join(cfg, entry)
                if not os.path.isdir(base):
                    continue
                for ver in ("4", "24"):
                    add(os.path.join(base, ver, "user"))
        except Exception:
            pass
    for p in _windows_profile_user_roots():
        add(p)
    return roots


def _iter_alteroffice_popup_paths():
    """Целевые пути cell.xml для AlterOffice (aoffice.cfg/acell/popupmenu)."""
    paths = []
    seen = set()
    for user_dir in iter_user_profile_roots() + _windows_profile_user_roots():
        for rel in _menu_rel_paths(user_dir, TARGET_CELL):
            if rel == _REL_POPUP:
                continue
            p = os.path.abspath(os.path.join(user_dir, *rel))
            if p not in seen:
                seen.add(p)
                paths.append(p)
    return paths


def _iter_alteroffice_menubar_paths():
    r"""
    Специфические пути AlterOffice для menubar.xml.
    Linux: ~/.config/alteroffice/5/user/config/aoffice.cfg/modules/acell/menubar/menubar.xml
    """
    paths = []
    home = os.path.expanduser("~")

    for ver in ("5", "4", "3"):
        p = os.path.join(
            home,
            ".config",
            "alteroffice",
            ver,
            "user",
            "config",
            "aoffice.cfg",
            "modules",
            "acell",
            "menubar",
            "menubar.xml",
        )
        if os.path.isfile(p):
            paths.append(p)

    appdata = os.environ.get("APPDATA")
    if appdata:
        for name in ("AlterOffice", "AlterOffice3"):
            for ver in ("4", "5", "3"):
                p = os.path.join(
                    appdata,
                    name,
                    ver,
                    "user",
                    "config",
                    "aoffice.cfg",
                    "modules",
                    "acell",
                    "menubar",
                    "menubar.xml",
                )
                if os.path.isfile(p):
                    paths.append(p)

    roaming = os.path.join(home, "AppData", "Roaming")
    if os.path.isdir(roaming):
        for name in ("AlterOffice", "AlterOffice3"):
            for ver in ("4", "5", "3"):
                p = os.path.join(
                    roaming,
                    name,
                    ver,
                    "user",
                    "config",
                    "aoffice.cfg",
                    "modules",
                    "acell",
                    "menubar",
                    "menubar.xml",
                )
                if os.path.isfile(p) and p not in paths:
                    paths.append(p)

    return paths


def menu_config_paths(target_kind=TARGET_CELL):
    """Пути cell.xml или menubar.xml в профилях LibreOffice / AlterOffice."""
    paths = []
    seen = set()

    def add_path(menu_xml):
        if menu_xml not in seen:
            seen.add(menu_xml)
            paths.append(menu_xml)

    def add_root(user_dir):
        if user_dir in (None, ""):
            return
        for rel in _menu_rel_paths(user_dir, target_kind):
            add_path(os.path.abspath(os.path.join(user_dir, *rel)))

    add_root(_user_dir_from_macro_lib())
    add_root(_uno_user_profile_dir())
    for root in iter_user_profile_roots():
        add_root(root)
    for root in _windows_profile_user_roots():
        add_root(root)

    if target_kind == TARGET_CELL:
        for ao_path in _iter_alteroffice_popup_paths():
            add_path(ao_path)

    return paths


def cell_popup_paths():
    """cell.xml в том же профиле, куда установлены макросы (Scripts/python)."""
    return menu_config_paths(TARGET_CELL)


def tools_menu_paths():
    """menubar.xml (меню Сервис) в профиле пользователя."""
    return menu_config_paths(TARGET_TOOLS)


def _cell_xml_label(path):
    return os.path.abspath(path)


def install_context_menu(spec_seq, macro_file, entry_key=None, target_kind=TARGET_CELL, paths=None):
    _unused = entry_key
    levels = parse_context_menu_spec(spec_seq)
    submenu_levels, content_levels = _split_menu_levels(levels)
    run_levels = [lv for lv in content_levels if _is_run_function_level(lv)]
    if len(run_levels) == 0:
        return 0, 0, [
            "%s: в CONTEXT_CELL_MENU нет run_function." % macro_file
        ]

    branch = _submenu_path_label(submenu_levels)
    target_label = _menu_target_label(target_kind)

    installed = 0
    skipped = 0
    messages = []
    menu_paths = paths if paths is not None else menu_config_paths(target_kind)
    if len(menu_paths) == 0:
        return 0, 0, [
            "Не найден %s. Каталог макросов: %s"
            % (target_label, os.path.abspath(_macro_lib_dir()))
        ]

    for path in menu_paths:
        menu_xml = _cell_xml_label(path)
        tree, popup = _read_menu_container(path, target_kind)
        try:
            if popup is None:
                if target_kind == TARGET_TOOLS:
                    messages.append(
                        "%s — %s: меню .uno:ToolsMenu не найдено"
                        % (menu_xml, macro_file)
                    )
                    continue
                os.makedirs(os.path.dirname(path), exist_ok=True)
                tree, popup = _default_cell_popup_tree()
            target_popup = _ensure_submenu_path(popup, submenu_levels)
            inst_n, skip_n, item_msgs = _install_menu_content_levels(
                target_popup, content_levels, macro_file
            )
            installed = installed + inst_n
            skipped = skipped + skip_n
            if inst_n == 0 and skip_n == 0:
                messages.append(
                    "%s — %s [%s]: нет новых пунктов (%s)"
                    % (menu_xml, macro_file, target_label, branch or "меню")
                )
            elif inst_n == 0:
                messages.append(
                    "%s — %s [%s]: уже в %s"
                    % (menu_xml, macro_file, target_label, branch or "меню")
                )
            else:
                messages.append(
                    "%s — %s [%s]: записано %d (%s)"
                    % (menu_xml, macro_file, target_label, inst_n, branch or "меню")
                )
            messages.extend(item_msgs)
            if inst_n > 0:
                _write_menu_xml_tree(path, tree)
        except Exception as err:
            messages.append(
                "%s — %s [%s]: ошибка — %s"
                % (menu_xml, macro_file, target_label, err)
            )
    return installed, skipped, messages


def install_context_cell_menu(spec_seq, macro_file, entry_key=None):
    """Совместимость: установка только в cell.xml."""
    return install_context_menu(
        spec_seq, macro_file, entry_key=entry_key, target_kind=TARGET_CELL
    )


def _show_message(text, title="Контекстное меню"):
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        sm = ctx.getServiceManager()
        parent = _dialog_parent()
        toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        from com.sun.star.awt.MessageBoxType import MESSAGEBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK

        dlg = toolkit.createMessageBox(parent, MESSAGEBOX, BUTTONS_OK, title, str(text))
        dlg.execute()
    except Exception:
        print("%s\n%s" % (title, text))


def _dialog_parent():
    try:
        doc = XSCRIPTCONTEXT.getDocument()
        if doc is not None:
            return doc.getCurrentController().getFrame().getContainerWindow()
    except Exception:
        pass
    try:
        return XSCRIPTCONTEXT.getDesktop().getCurrentFrame().getContainerWindow()
    except Exception:
        return None


def _ctl_name(event):
    try:
        src = event.Source
        if hasattr(src, "getModel"):
            return str(src.getModel().Name)
        return str(src.Model.Name)
    except Exception:
        return ""


def _pick_macros_dialog(entries):
    """
    Диалог с флажками: какие макросы установить.

    Возвращает список выбранных entries или None при отмене.
    """
    if len(entries) == 0:
        return None
    ctx = XSCRIPTCONTEXT.getComponentContext()
    sm = ctx.getServiceManager()
    toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    dm = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dm.Title = "Меню макросов Calc"
    dm.Width = _DLG_W
    n = len(entries)
    list_h = min(n * _DLG_ROW_H, 220)
    btn_y = _DLG_MARGIN + 36 + list_h + _DLG_GAP
    dm.Height = btn_y + _DLG_BTN_H + _DLG_MARGIN

    hint = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    hint.PositionX = _DLG_MARGIN
    hint.PositionY = _DLG_MARGIN
    hint.Width = _DLG_W - 2 * _DLG_MARGIN
    hint.Height = 28
    hint.Label = (
        "Отметьте макросы для контекстного меню ячейки и меню «Сервис»:"
    )
    hint.Name = "HintLabel"
    dm.insertByName("HintLabel", hint)

    checks = []
    i = 0
    while i < n:
        m = dm.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
        m.PositionX = _DLG_MARGIN
        m.PositionY = _DLG_MARGIN + 32 + i * _DLG_ROW_H
        m.Width = _DLG_W - 2 * _DLG_MARGIN
        m.Height = _DLG_ROW_H
        m.Label = "%s  (%s)" % (entries[i]["label"], entries[i]["file"])
        m.Name = "Check_%d" % i
        m.State = 1
        dm.insertByName(m.Name, m)
        checks.append(m.Name)
        i = i + 1

    bx = _DLG_W - _DLG_MARGIN - _DLG_BTN_W
    cancel_m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_m.PositionX = bx
    cancel_m.PositionY = btn_y
    cancel_m.Width = _DLG_BTN_W
    cancel_m.Height = _DLG_BTN_H
    cancel_m.Label = "Отмена"
    cancel_m.Name = "CancelButton"
    dm.insertByName("CancelButton", cancel_m)

    ok_m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_m.PositionX = bx - _DLG_GAP - _DLG_BTN_W
    ok_m.PositionY = btn_y
    ok_m.Width = _DLG_BTN_W
    ok_m.Height = _DLG_BTN_H
    ok_m.Label = "Установить"
    ok_m.Name = "OkButton"
    dm.insertByName("OkButton", ok_m)

    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(dm)
    dialog.createPeer(toolkit, _dialog_parent())

    class _Handler(unohelper.Base, XActionListener):
        def __init__(self, dlg):
            unohelper.Base.__init__(self)
            self.dialog = dlg
            self.confirmed = False

        def disposing(self, event):
            pass

        def actionPerformed(self, event):
            name = _ctl_name(event)
            if name == "OkButton":
                self.confirmed = True
                self.dialog.endExecute()
            elif name == "CancelButton":
                self.confirmed = False
                self.dialog.endExecute()

    handler = _Handler(dialog)
    ok_ctl = dialog.getControl("OkButton")
    cancel_ctl = dialog.getControl("CancelButton")
    if ok_ctl is not None:
        ok_ctl.addActionListener(handler)
    if cancel_ctl is not None:
        cancel_ctl.addActionListener(handler)

    dialog.execute()

    if not handler.confirmed:
        return None

    selected = []
    i = 0
    while i < n:
        try:
            ctl = dialog.getControl("Check_%d" % i)
            if ctl is not None and int(ctl.getModel().State) == 1:
                selected.append(entries[i])
        except Exception:
            pass
        i = i + 1
    return selected


def setContextMenu():
    """
    Сканировать соседние .py, выбрать макросы в диалоге, записать cell.xml и menubar.xml.

    Точка входа: Макросы → Python → context_menu.setContextMenu

    Порядок пунктов внутри одного подменю задаётся атрибутом order#N
    у run_function (и при необходимости у separator) в CONTEXT_CELL_MENU.
    Пункты из разных .py с одной цепочкой submenu#… сортируются вместе.
    """
    lib_dir = _macro_lib_dir()
    entries = discover_context_menu_macros(lib_dir)
    if len(entries) == 0:
        _show_message(
            "В каталоге макросов не найдено CONTEXT_CELL_MENU:\n%s\n\n"
            "Добавьте в .py макроса, например:\n"
            "CONTEXT_CELL_MENU = (\n"
            '  "submenu#Python",\n'
            '  "run_function#my_func;module#my.py;display_name#Подпись;order#10",\n'
            ")" % lib_dir,
            "Меню макросов",
        )
        return False

    selected = _pick_macros_dialog(entries)
    if selected is None:
        return False
    if len(selected) == 0:
        _show_message("Ничего не выбрано.", "Меню макросов")
        return False

    managed_submenus = collect_managed_top_submenus(entries)
    groups = collect_menu_install_groups(selected)
    total_inst = 0
    total_skip = 0
    all_msgs = [
        "Каталог макросов: %s" % os.path.abspath(lib_dir),
        "",
    ]
    submenu_labels = [str(lv.get("value", "")) for lv in managed_submenus]
    all_msgs.append(
        "Очистка подменю верхнего уровня: %s"
        % (", ".join(submenu_labels) if submenu_labels else "—")
    )
    all_msgs.append("")
    if len(groups) > 0:
        all_msgs.append("Порядок пунктов (order#):")
        gi = 0
        while gi < len(groups):
            g = groups[gi]
            all_msgs.append("  [%s]" % (g.get("branch") or "меню"))
            ui = 0
            while ui < len(g["units"]):
                u = g["units"][ui]
                if u.get("type") == "separator":
                    all_msgs.append(
                        "    order=%s  — разделитель" % u.get("order")
                    )
                else:
                    lv = u.get("level") or {}
                    title = lv.get("display_name", lv.get("value", "?"))
                    all_msgs.append(
                        "    order=%s  %s (%s)"
                        % (u.get("order"), title, u.get("macro_file"))
                    )
                ui = ui + 1
            gi = gi + 1
        all_msgs.append("")

    for target_kind in (TARGET_CELL, TARGET_TOOLS):
        target_label = _menu_target_label(target_kind)
        menu_paths = menu_config_paths(target_kind)
        all_msgs.append("Файлы %s:" % target_label)
        if len(menu_paths) == 0:
            all_msgs.append("  (профили LibreOffice / AlterOffice не найдены)")
        else:
            pi = 0
            while pi < len(menu_paths):
                all_msgs.append("  %s" % menu_paths[pi])
                pi = pi + 1
        all_msgs.append("")
        purge_msgs = purge_managed_submenus_all_paths(
            managed_submenus, paths=menu_paths, target_kind=target_kind
        )
        pi = 0
        while pi < len(purge_msgs):
            all_msgs.append("  %s" % purge_msgs[pi])
            pi = pi + 1
        all_msgs.append("")
        inst, skip, msgs = install_menu_groups(
            groups, target_kind=target_kind, paths=menu_paths
        )
        total_inst = total_inst + inst
        total_skip = total_skip + skip
        all_msgs.extend(msgs)
        all_msgs.append("")

    all_msgs.append("Записано (суммарно): %d" % total_inst)
    all_msgs.append("Уже было (суммарно): %d" % total_skip)
    if total_inst == 0 and total_skip == 0:
        all_msgs.append("")
        all_msgs.append(
            "Ни один файл меню не изменён. Проверьте пути выше — "
            "должны быть каталоги .../user/config/soffice.cfg/..."
        )
    all_msgs.append("")
    all_msgs.append("Перезапустите Calc для отображения меню.")
    _show_message("\n".join(all_msgs), "Меню макросов")
    return total_inst > 0 or total_skip > 0


g_exportedScripts = (setContextMenu,)
