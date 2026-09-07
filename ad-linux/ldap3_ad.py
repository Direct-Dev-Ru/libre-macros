# -*- coding: utf-8 -*-
"""
Общая логика AD через bundled ldap3 (внешний Python, Linux).

Импорт из локальной копии ad-linux/bundled/ldap3_bundled.py — каталог автономен.
"""
from __future__ import annotations

import sys
from pathlib import Path

_BUNDLED_DIR = Path(__file__).resolve().parent / "bundled"
if not (_BUNDLED_DIR / "ldap3_bundled.py").is_file():
    raise ImportError(
        "Нет ad-linux/bundled/ldap3_bundled.py — выполните: python3 ad-linux/sync_bundles.py"
    )
if str(_BUNDLED_DIR) not in sys.path:
    sys.path.insert(0, str(_BUNDLED_DIR))

import ldap3_bundled  # noqa: F401,E402 — регистрирует import hook
from ldap3 import ALL, NTLM, SIMPLE, SUBTREE, Connection, Server  # noqa: E402

DEFAULT_PORT = 389
DEFAULT_SSL_PORT = 636
DEFAULT_PAGE_SIZE = 500
USER_SEARCH_FILTER = "(&(objectClass=user)(objectCategory=person))"
COMPUTER_SEARCH_FILTER = "(&(objectCategory=computer)(objectClass=computer))"
USER_LDAP_ATTRIBUTES = (
    "sAMAccountName",
    "displayName",
    "mail",
    "title",
    "department",
    "telephoneNumber",
    "userPrincipalName",
    "memberOf",
    "distinguishedName",
)
OUTPUT_COLUMNS = (
    "sAMAccountName",
    "displayName",
    "mail",
    "title",
    "department",
    "telephoneNumber",
    "userPrincipalName",
    "groups",
    "distinguishedName",
)
COMPUTER_LDAP_ATTRIBUTES = (
    "name",
    "sAMAccountName",
    "dNSHostName",
    "operatingSystem",
    "operatingSystemVersion",
    "description",
    "distinguishedName",
    "lastLogonTimestamp",
    "userAccountControl",
    "whenCreated",
    "whenChanged",
)
OUTPUT_COMPUTER_COLUMNS = (
    "name",
    "sAMAccountName",
    "dNSHostName",
    "operatingSystem",
    "operatingSystemVersion",
    "description",
    "enabled",
    "lastLogon",
    "whenCreated",
    "whenChanged",
    "distinguishedName",
)


def cn_from_dn(dn: str) -> str:
    text = (dn or "").strip()
    if not text:
        return ""
    part = text.split(",", 1)[0].strip()
    if part.upper().startswith("CN="):
        return part[3:].replace("\\,", ",")
    return text


def format_member_of(values) -> str:
    if not values:
        return ""
    if not isinstance(values, (list, tuple)):
        values = (values,)
    names = []
    seen = set()
    for item in values:
        cn = cn_from_dn(str(item))
        if cn and cn not in seen:
            seen.add(cn)
            names.append(cn)
    return "; ".join(names)


def resolve_auth(name: str):
    text = (name or "").strip().upper()
    if text in ("NTLM", "NTLMV2"):
        return NTLM
    if text in ("SIMPLE", ""):
        return SIMPLE
    raise ValueError("auth: ожидается NTLM или SIMPLE, получено %r" % name)


def resolve_port(use_ssl: bool, port) -> int:
    if port is None or str(port).strip() == "":
        return DEFAULT_SSL_PORT if use_ssl else DEFAULT_PORT
    port_num = int(port)
    if port_num <= 0:
        return DEFAULT_SSL_PORT if use_ssl else DEFAULT_PORT
    return port_num


def _bind_error_hint(auth: str, bind_user: str) -> str:
    auth_u = (auth or "").strip().upper()
    user = bind_user or ""
    lines = [
        "Bind отклонён AD (invalidCredentials). Проверьте:",
        "  • пароль (getpass / -p); в shell логин в одинарных кавычках: 'DOMAIN\\user'",
    ]
    if auth_u in ("NTLM", "NTLMV2"):
        lines.append("  • NTLM: логин DOMAIN\\user (не просто ivanov)")
        if "@" in user:
            lines.append("  • для user@domain.com попробуйте: -a SIMPLE")
    else:
        lines.append("  • SIMPLE: UPN user@domain.com или полный DN; DOMAIN\\user для SIMPLE не подходит")
        if "\\" in user and "@" not in user:
            lines.append("  • для DOMAIN\\user попробуйте: -a NTLM")
    lines.append("  • учётка не заблокирована; есть право чтения OU")
    lines.append("  • на Linux надёжнее: -a SIMPLE -u reader@corp.local")
    return "\n".join(lines)


def connect(server: str, user: str, password: str, auth: str = "NTLM", use_ssl: bool = True, port=None, timeout: int = 30):
    host = (server or "").strip()
    if not host:
        raise ValueError("Не задан сервер LDAP/AD (DC).")
    bind_user = (user or "").strip()
    if not bind_user:
        raise ValueError("Не задан пользователь для bind.")
    auth_mode = resolve_auth(auth)
    use_ssl_flag = bool(use_ssl)
    port_num = resolve_port(use_ssl_flag, port)

    srv = Server(
        host,
        port=port_num,
        use_ssl=use_ssl_flag,
        get_info=ALL,
        connect_timeout=int(timeout),
    )
    try:
        conn = Connection(
            srv,
            user=bind_user,
            password=password or "",
            authentication=auth_mode,
            auto_bind=True,
            receive_timeout=int(timeout),
        )
    except Exception as err:
        text = str(err).lower()
        if "invalidcredentials" in text or "bind" in text:
            raise ValueError("%s\n\n%s" % (err, _bind_error_hint(auth, bind_user))) from err
        raise
    return conn


def _attr_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return "; ".join(_attr_text(v) for v in value if v is not None)
    return str(value).strip()


def _format_ad_datetime(value) -> str:
    if value is None or value == "":
        return ""
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat(sep=" ", timespec="seconds")
        except TypeError:
            return value.isoformat()
    return _attr_text(value)


def _enabled_from_uac(uac) -> str:
    try:
        val = int(uac)
    except (TypeError, ValueError):
        return ""
    return "1" if not (val & 0x0002) else "0"


def iter_entries(
    conn,
    ou_dn: str,
    search_filter: str,
    attributes,
    page_size: int = DEFAULT_PAGE_SIZE,
    server: str | None = None,
    dc_parts: int | None = None,
):
    from dn_format import normalize_base_dn

    base = normalize_base_dn((ou_dn or "").strip(), server=server, dc_parts=dc_parts)
    if not base:
        raise ValueError("Не задана OU (search base).")
    attrs = list(attributes)
    filt = (search_filter or "").strip()
    size = max(1, int(page_size or DEFAULT_PAGE_SIZE))

    generator = conn.extend.standard.paged_search(
        search_base=base,
        search_filter=filt,
        search_scope=SUBTREE,
        attributes=attrs,
        paged_size=size,
        generator=True,
    )
    for item in generator:
        if isinstance(item, dict) and item.get("type") == "searchResEntry":
            yield item


def iter_users(conn, ou_dn: str, search_filter: str | None = None, attributes=None, page_size: int = DEFAULT_PAGE_SIZE):
    return iter_entries(
        conn,
        ou_dn,
        search_filter or USER_SEARCH_FILTER,
        attributes or USER_LDAP_ATTRIBUTES,
        page_size=page_size,
    )


def iter_computers(conn, ou_dn: str, search_filter: str | None = None, attributes=None, page_size: int = DEFAULT_PAGE_SIZE):
    return iter_entries(
        conn,
        ou_dn,
        search_filter or COMPUTER_SEARCH_FILTER,
        attributes or COMPUTER_LDAP_ATTRIBUTES,
        page_size=page_size,
    )


def entry_to_row(entry: dict) -> dict:
    attrs = entry.get("attributes") or {}
    return {
        "sAMAccountName": _attr_text(attrs.get("sAMAccountName")),
        "displayName": _attr_text(attrs.get("displayName")),
        "mail": _attr_text(attrs.get("mail")),
        "title": _attr_text(attrs.get("title")),
        "department": _attr_text(attrs.get("department")),
        "telephoneNumber": _attr_text(attrs.get("telephoneNumber")),
        "userPrincipalName": _attr_text(attrs.get("userPrincipalName")),
        "groups": format_member_of(attrs.get("memberOf")),
        "distinguishedName": _attr_text(attrs.get("distinguishedName") or entry.get("dn")),
    }


def computer_entry_to_row(entry: dict) -> dict:
    attrs = entry.get("attributes") or {}
    return {
        "name": _attr_text(attrs.get("name")),
        "sAMAccountName": _attr_text(attrs.get("sAMAccountName")),
        "dNSHostName": _attr_text(attrs.get("dNSHostName")),
        "operatingSystem": _attr_text(attrs.get("operatingSystem")),
        "operatingSystemVersion": _attr_text(attrs.get("operatingSystemVersion")),
        "description": _attr_text(attrs.get("description")),
        "enabled": _enabled_from_uac(attrs.get("userAccountControl")),
        "lastLogon": _format_ad_datetime(attrs.get("lastLogonTimestamp")),
        "whenCreated": _format_ad_datetime(attrs.get("whenCreated")),
        "whenChanged": _format_ad_datetime(attrs.get("whenChanged")),
        "distinguishedName": _attr_text(attrs.get("distinguishedName") or entry.get("dn")),
    }


def _search_entries(
    server,
    user,
    password,
    ou_dn,
    auth,
    use_ssl,
    port,
    page_size,
    search_filter,
    attributes,
    row_mapper,
    sort_key,
    dc_parts=None,
):
    conn = connect(server, user, password, auth=auth, use_ssl=use_ssl, port=port)
    rows = []
    try:
        for entry in iter_entries(
            conn, ou_dn, search_filter, attributes, page_size=page_size, server=server, dc_parts=dc_parts
        ):
            rows.append(row_mapper(entry))
    finally:
        try:
            conn.unbind()
        except Exception:
            pass
    rows.sort(key=sort_key)
    return rows


def search_users(
    server: str,
    user: str,
    password: str,
    ou_dn: str,
    auth: str = "NTLM",
    use_ssl: bool = True,
    port=None,
    page_size: int = DEFAULT_PAGE_SIZE,
    search_filter: str | None = None,
    dc_parts: int | None = None,
):
    filt = (search_filter or USER_SEARCH_FILTER).strip()
    return _search_entries(
        server,
        user,
        password,
        ou_dn,
        auth,
        use_ssl,
        port,
        page_size,
        filt,
        USER_LDAP_ATTRIBUTES,
        entry_to_row,
        lambda r: (r.get("displayName") or r.get("sAMAccountName") or "").lower(),
        dc_parts=dc_parts,
    )


def search_computers(
    server: str,
    user: str,
    password: str,
    ou_dn: str,
    auth: str = "NTLM",
    use_ssl: bool = True,
    port=None,
    page_size: int = DEFAULT_PAGE_SIZE,
    search_filter: str | None = None,
    dc_parts: int | None = None,
):
    filt = (search_filter or COMPUTER_SEARCH_FILTER).strip()
    return _search_entries(
        server,
        user,
        password,
        ou_dn,
        auth,
        use_ssl,
        port,
        page_size,
        filt,
        COMPUTER_LDAP_ATTRIBUTES,
        computer_entry_to_row,
        lambda r: (r.get("dNSHostName") or r.get("name") or r.get("sAMAccountName") or "").lower(),
        dc_parts=dc_parts,
    )
