# -*- coding: utf-8 -*-
"""
Доступ к Active Directory через bundled ldap3 (SIMPLE / NTLM).

Типовые операции: подключение, поиск пользователей в OU, список групп (memberOf).
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.693"
from libre_macros_ad_ssl_stub import install_ssl_stub_if_needed, ssl_is_stub
from libre_macros_ad_winldap import pick_ad_backend

install_ssl_stub_if_needed()

from libre_macros_ad_cfg import (
    AUTH_CHOICES,
    CONFIG_COL_AUTH,
    CONFIG_COL_OU,
    CONFIG_COL_PORT,
    CONFIG_COL_SERVER,
    CONFIG_COL_SSL,
    CONFIG_COL_USER,
    CONFIG_ROW,
    CONFIG_SHEET,
    DEFAULT_AUTH,
    DEFAULT_PAGE_SIZE,
    DEFAULT_PORT,
    DEFAULT_SSL_PORT,
    OUTPUT_DATA_START_ROW,
    OUTPUT_HEADER_ROW,
    OUTPUT_HEADERS,
    USER_LDAP_ATTRIBUTES,
    USER_SEARCH_FILTER,
)

try:
    unicode
except NameError:
    unicode = str

def _log(msg):
    try:
        print("[ad] %s" % msg)
    except Exception:
        pass


def _text(value):
    if value is None:
        return u""
    if isinstance(value, (list, tuple)):
        parts = []
        idx = 0
        while idx < len(value):
            parts.append(_text(value[idx]))
            idx = idx + 1
        return u"; ".join([p for p in parts if p])
    try:
        return unicode(value).strip()
    except Exception:
        return str(value).strip()


def _truthy(value):
    text = _text(value).lower()
    return text in (u"1", u"true", u"yes", u"да", u"y", u"on", u"ssl", u"ldaps")


def _ldap3_imports():
    import ldap3_bundled  # noqa: F401 — регистрирует import hook
    from ldap3 import ALL, NTLM, SIMPLE, SUBTREE, Connection, Server

    return {
        u"ALL": ALL,
        u"NTLM": NTLM,
        u"SIMPLE": SIMPLE,
        u"SUBTREE": SUBTREE,
        u"Connection": Connection,
        u"Server": Server,
    }


def _auth_mode(name):
    text = _text(name).upper()
    l3 = _ldap3_imports()
    if text in (u"NTLM", u"NTLMV2"):
        return l3[u"NTLM"]
    if text in (u"SIMPLE", u""):
        return l3[u"SIMPLE"]
    allowed = u", ".join(AUTH_CHOICES)
    raise ValueError(u"auth должен быть SIMPLE или NTLM, получено: %s (допустимо: %s)" % (name, allowed))


def cn_from_dn(dn):
    """Извлечь CN из distinguishedName (первая RDN-компонента)."""
    text = _text(dn)
    if text == u"":
        return u""
    part = text.split(u",", 1)[0].strip()
    if part.upper().startswith(u"CN="):
        return part[3:].replace(u"\\,", u",")
    return text


def format_member_of(values):
    names = []
    seen = set()
    if not values:
        return u""
    if not isinstance(values, (list, tuple)):
        values = (values,)
    idx = 0
    while idx < len(values):
        cn = cn_from_dn(values[idx])
        if cn and cn not in seen:
            seen.add(cn)
            names.append(cn)
        idx = idx + 1
    return u"; ".join(names)


def resolve_port(use_ssl, port):
    if port is None:
        return DEFAULT_SSL_PORT if use_ssl else DEFAULT_PORT
    try:
        port_num = int(port)
    except (TypeError, ValueError):
        raise ValueError(u"port должен быть числом, получено: %s" % port)
    if port_num <= 0:
        return DEFAULT_SSL_PORT if use_ssl else DEFAULT_PORT
    return port_num


def _ensure_plain_ldap_if_no_ssl(use_ssl_flag):
    if not use_ssl_flag:
        return
    if pick_ad_backend(True) == u"winldap":
        return
    if ssl_is_stub():
        raise ValueError(
            u"LDAPS недоступен через ldap3: встроенный Python без ssl. "
            u"На Windows установите pythonpath/libre_macros_ad_winldap.py (WinLDAP) "
            u"или укажите SSL=0 и порт 389."
        )


def connect_ad(server, bind_user, bind_password, auth=None, use_ssl=False, port=None, receive_timeout=30):
    l3 = _ldap3_imports()
    SIMPLE = l3[u"SIMPLE"]
    NTLM = l3[u"NTLM"]
    ALL = l3[u"ALL"]
    Server = l3[u"Server"]
    Connection = l3[u"Connection"]
    if auth is None:
        auth = SIMPLE
    host = _text(server)
    if host == u"":
        raise ValueError(u"Не задан сервер LDAP/AD (DC).")
    user = _text(bind_user)
    if user == u"":
        raise ValueError(u"Не задан пользователь для bind.")
    password = bind_password if bind_password is not None else u""
    auth_mode = auth if auth in (SIMPLE, NTLM) else _auth_mode(auth)
    use_ssl_flag = bool(use_ssl)
    _ensure_plain_ldap_if_no_ssl(use_ssl_flag)
    port_num = resolve_port(use_ssl_flag, port)

    srv = Server(
        host,
        port=port_num,
        use_ssl=use_ssl_flag,
        get_info=ALL,
        connect_timeout=int(receive_timeout),
    )
    conn = Connection(
        srv,
        user=user,
        password=password,
        authentication=auth_mode,
        auto_bind=True,
        receive_timeout=int(receive_timeout),
    )
    _log(
        u"bind OK: %s:%s ssl=%s auth=%s user=%s"
        % (host, port_num, use_ssl_flag, auth_mode, user)
    )
    return conn


def iter_users_in_ou(conn, ou_dn, attributes=None, search_filter=None, page_size=None):
    l3 = _ldap3_imports()
    SUBTREE = l3[u"SUBTREE"]
    base = _text(ou_dn)
    if base == u"":
        raise ValueError(u"Не задана OU (search base).")
    attrs = list(attributes or USER_LDAP_ATTRIBUTES)
    filt = _text(search_filter or USER_SEARCH_FILTER)
    size = int(page_size or DEFAULT_PAGE_SIZE)
    if size <= 0:
        size = DEFAULT_PAGE_SIZE

    generator = conn.extend.standard.paged_search(
        search_base=base,
        search_filter=filt,
        search_scope=SUBTREE,
        attributes=attrs,
        paged_size=size,
        generator=True,
    )
    for item in generator:
        if not isinstance(item, dict):
            continue
        if item.get("type") != "searchResEntry":
            continue
        yield item


def user_record_from_entry(entry):
    attrs = entry.get("attributes") or {}
    return {
        u"sAMAccountName": _text(attrs.get("sAMAccountName")),
        u"displayName": _text(attrs.get("displayName")),
        u"mail": _text(attrs.get("mail")),
        u"title": _text(attrs.get("title")),
        u"department": _text(attrs.get("department")),
        u"telephoneNumber": _text(attrs.get("telephoneNumber")),
        u"userPrincipalName": _text(attrs.get("userPrincipalName")),
        u"groups": format_member_of(attrs.get("memberOf")),
        u"distinguishedName": _text(attrs.get("distinguishedName") or entry.get("dn")),
    }


def search_users_in_ou(server, bind_user, bind_password, ou_dn, auth=DEFAULT_AUTH, use_ssl=False, port=None, page_size=None):
    backend = pick_ad_backend(use_ssl)
    _log(u"backend=%s ssl=%s server=%s" % (backend, use_ssl, server))
    if backend == u"winldap":
        from libre_macros_ad_winldap import search_users_winldap

        entries = search_users_winldap(
            server,
            bind_user,
            bind_password,
            ou_dn,
            use_ssl=use_ssl,
            port=port,
            page_size=page_size,
        )
        rows = [user_record_from_entry(entry) for entry in entries]
        rows.sort(key=lambda row: (row.get(u"displayName") or row.get(u"sAMAccountName") or u"").lower())
        return rows

    conn = connect_ad(
        server,
        bind_user,
        bind_password,
        auth=auth,
        use_ssl=use_ssl,
        port=port,
    )
    rows = []
    try:
        for entry in iter_users_in_ou(conn, ou_dn, page_size=page_size):
            rows.append(user_record_from_entry(entry))
    finally:
        try:
            conn.unbind()
        except Exception:
            pass
    rows.sort(key=lambda row: (row.get(u"displayName") or row.get(u"sAMAccountName") or u"").lower())
    return rows


def search_groups(conn, base_dn, name_filter=u"*", page_size=None):
    l3 = _ldap3_imports()
    SUBTREE = l3[u"SUBTREE"]
    """Список групп в base_dn (objectClass=group)."""
    base = _text(base_dn)
    if base == u"":
        raise ValueError(u"Не задан base_dn для поиска групп.")
    safe = _text(name_filter).replace(u"(", u"").replace(u")", u"").replace(u"\\", u"")
    if safe in (u"", u"*"):
        filt = u"(objectClass=group)"
    else:
        filt = u"(&(objectClass=group)(cn=%s))" % safe
    size = int(page_size or DEFAULT_PAGE_SIZE)
    if size <= 0:
        size = DEFAULT_PAGE_SIZE

    rows = []
    generator = conn.extend.standard.paged_search(
        search_base=base,
        search_filter=filt,
        search_scope=SUBTREE,
        attributes=("cn", "distinguishedName", "groupType", "description"),
        paged_size=size,
        generator=True,
    )
    for item in generator:
        if not isinstance(item, dict) or item.get("type") != "searchResEntry":
            continue
        attrs = item.get("attributes") or {}
        group_type = attrs.get("groupType")
        is_security = False
        try:
            is_security = bool(int(group_type) & 0x80000000)
        except Exception:
            is_security = False
        rows.append(
            {
                u"cn": _text(attrs.get("cn")),
                u"distinguishedName": _text(attrs.get("distinguishedName") or item.get("dn")),
                u"description": _text(attrs.get("description")),
                u"security": is_security,
            }
        )
    rows.sort(key=lambda row: (row.get(u"cn") or u"").lower())
    return rows


def _cell_string(cell):
    if cell is None:
        return u""
    try:
        val = cell.String
    except Exception:
        val = u""
    if val is None:
        return u""
    return unicode(val).strip()


def _read_config_from_sheet(doc, sheet_name):
    name = _text(sheet_name or CONFIG_SHEET)
    sheets = doc.getSheets()
    if sheets.hasByName(name):
        sheet = sheets.getByName(name)
    else:
        sheet = doc.getSheets().getByIndex(0)
    row = CONFIG_ROW
    return {
        u"server": _cell_string(sheet.getCellByPosition(CONFIG_COL_SERVER, row)),
        u"bind_user": _cell_string(sheet.getCellByPosition(CONFIG_COL_USER, row)),
        u"auth": _cell_string(sheet.getCellByPosition(CONFIG_COL_AUTH, row)) or DEFAULT_AUTH,
        u"use_ssl": _truthy(sheet.getCellByPosition(CONFIG_COL_SSL, row)),
        u"port": _cell_string(sheet.getCellByPosition(CONFIG_COL_PORT, row)),
        u"ou_dn": _cell_string(sheet.getCellByPosition(CONFIG_COL_OU, row)),
        u"sheet": sheet,
    }


def _resolve_connection_params(ou_dn, server, bind_user, bind_password, auth, use_ssl, port, sheet_name, doc):
    cfg = _read_config_from_sheet(doc, sheet_name)
    resolved = {
        u"ou_dn": _text(ou_dn) or cfg.get(u"ou_dn"),
        u"server": _text(server) or cfg.get(u"server"),
        u"bind_user": _text(bind_user) or cfg.get(u"bind_user"),
        u"bind_password": None,
        u"auth": _text(auth) or cfg.get(u"auth") or DEFAULT_AUTH,
        u"use_ssl": use_ssl if use_ssl is not None else cfg.get(u"use_ssl"),
        u"port": port if port not in (None, u"") else cfg.get(u"port"),
        u"sheet": cfg.get(u"sheet"),
    }
    if resolved[u"ou_dn"] == u"":
        raise ValueError(
            u"Не задана OU. Передайте параметр ou_dn или заполните ячейку G2 листа «%s»."
            % CONFIG_SHEET
        )
    from libre_macros_ad_dn_lib import normalize_base_dn

    resolved[u"ou_dn"] = normalize_base_dn(resolved[u"ou_dn"], server=resolved[u"server"])
    if bind_password not in (None, u""):
        resolved[u"bind_password"] = bind_password
    return resolved


def _prompt_bind_password(doc, bind_user):
    from libre_macros_ad_dialog_lib import prompt_ad_password

    password = prompt_ad_password(doc=doc, bind_user=bind_user)
    if password is None:
        return None
    return password


def write_users_to_sheet(sheet, rows, header_row=None, data_start_row=None):
    header_row = OUTPUT_HEADER_ROW if header_row is None else int(header_row)
    data_start_row = OUTPUT_DATA_START_ROW if data_start_row is None else int(data_start_row)

    col = 0
    while col < len(OUTPUT_HEADERS):
        sheet.getCellByPosition(col, header_row).String = OUTPUT_HEADERS[col]
        col = col + 1

    row_idx = data_start_row
    rec_idx = 0
    while rec_idx < len(rows):
        rec = rows[rec_idx]
        values = (
            rec.get(u"sAMAccountName"),
            rec.get(u"displayName"),
            rec.get(u"mail"),
            rec.get(u"title"),
            rec.get(u"department"),
            rec.get(u"telephoneNumber"),
            rec.get(u"userPrincipalName"),
            rec.get(u"groups"),
        )
        col = 0
        while col < len(values):
            sheet.getCellByPosition(col, row_idx).String = _text(values[col])
            col = col + 1
        row_idx = row_idx + 1
        rec_idx = rec_idx + 1
    return len(rows)


def _show_message(title, text, doc=None, error=False):
    import uno
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    if doc is None:
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.getCurrentComponent()
    parent = getattr(
        getattr(getattr(doc, "CurrentController", None), "Frame", None),
        "ContainerWindow",
        None,
    )
    if parent is None:
        parent = toolkit.getDesktopWindow()
    box_type = ERRORBOX if error else INFOBOX
    box = toolkit.createMessageBox(parent, box_type, BUTTONS_OK, title, text)
    box.execute()


def _current_doc():
    import uno

    try:
        import __main__

        xsc = getattr(__main__, "XSCRIPTCONTEXT", None)
        if xsc is not None:
            return xsc.getDocument()
    except Exception:
        pass
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop.getCurrentComponent()


def ad_list_users(ou_dn=u"", server=u"", bind_user=u"", bind_password=None, auth=u"", use_ssl=None, port=None, sheet_name=u""):
    """
    Вывести пользователей OU на лист: основные атрибуты и группы (memberOf).

    Параметры подключения — аргументы макроса или строка 2 листа «AD»:
      A2 сервер, B2 пользователь, D2 auth, E2 SSL, F2 порт, G2 OU.
    Пароль запрашивается в диалоге (маскированный ввод), на лист не сохраняется.
    """
    doc = _current_doc()
    if doc is None or not hasattr(doc, "getSheets"):
        raise RuntimeError(u"Нужна открытая книга Calc.")

    params = _resolve_connection_params(
        ou_dn, server, bind_user, bind_password, auth, use_ssl, port, sheet_name, doc
    )
    password = params[u"bind_password"]
    if password is None:
        password = _prompt_bind_password(doc, params[u"bind_user"])
    if password is None:
        _log(u"cancelled: password dialog")
        return 0
    try:
        rows = search_users_in_ou(
            params[u"server"],
            params[u"bind_user"],
            password,
            params[u"ou_dn"],
            auth=params[u"auth"],
            use_ssl=params[u"use_ssl"],
            port=params[u"port"],
        )
        count = write_users_to_sheet(params[u"sheet"], rows)
        msg = u"OU: %s\nНайдено пользователей: %d" % (params[u"ou_dn"], count)
        _log(msg)
        _show_message(u"Active Directory", msg, doc=doc)
        return count
    finally:
        password = None


def ad_search_users(ou_dn, server, bind_user, bind_password, auth=DEFAULT_AUTH, use_ssl=False, port=None):
    """Поиск пользователей без записи на лист — для вызова из других макросов."""
    return search_users_in_ou(
        server,
        bind_user,
        bind_password,
        ou_dn,
        auth=auth,
        use_ssl=use_ssl,
        port=port,
    )


def ad_list_groups(base_dn, server, bind_user, bind_password, auth=DEFAULT_AUTH, use_ssl=False, port=None, name_filter=u"*"):
    """Список групп в base_dn (в т.ч. признак security по groupType)."""
    conn = connect_ad(
        server,
        bind_user,
        bind_password,
        auth=auth,
        use_ssl=use_ssl,
        port=port,
    )
    try:
        return search_groups(conn, base_dn, name_filter=name_filter)
    finally:
        try:
            conn.unbind()
        except Exception:
            pass
