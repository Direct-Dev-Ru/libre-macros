# -*- coding: utf-8 -*-
"""
WinLDAP (Wldap32.dll) backend для Active Directory на Windows.

Зачем
-----
Embedded Python AlterOffice/LibreOffice на Windows не содержит рабочий модуль ssl,
поэтому bundled ldap3 не поддерживает LDAPS (636). WinLDAP использует Schannel ОС
и не зависит от Python ssl.

Архитecture / поток данных
--------------------------
::

    ad_list_users (entry)
            |
            v
    libre_macros_ad_lib.search_users_in_ou()
            |
            v
    pick_ad_backend(use_ssl)
            |
            +-- "winldap" (Windows + wldap32 + ssl/LDAPS или ssl-stub)
            |       connect_winldap -> WinLdapSession.bind_simple
            |       -> paged_search -> [{dn, attributes}] -> user_record_from_entry
            |
            +-- "ldap3" (Linux, Windows с рабочим ssl, plain LDAP 389)
                    connect_ad -> iter_users_in_ou

Публичный API
-------------
- is_winldap_available() -> bool
- pick_ad_backend(use_ssl) -> "winldap" | "ldap3"
- connect_winldap(server, port, use_ssl) -> WinLdapSession
- WinLdapSession.bind_simple(user, password)
- WinLdapSession.paged_search(base, filter, attributes, page_size)
- search_users_winldap(...) -> list[{dn, attributes, type}]

Ограничения v1
--------------
- Windows + wldap32.dll только.
- Bind: SIMPLE (user@domain или DN). NTLM — через LDAP_AUTH_NEGOTIATE (TODO).
- LDAPS: ldap_sslinitW(host, port, 1).
- StartTLS: не реализован.
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.691"
import sys

try:
    unicode
except NameError:
    unicode = str

LDAP_SCOPE_SUBTREE = 2
LDAP_OPT_PROTOCOL_VERSION = 0x0011
LDAP_OPT_REFERRALS = 0x0008
LDAP_VERSION3 = 3
LDAP_OPT_OFF = 0
LDAP_AUTH_SIMPLE = 0x0080
LDAP_SUCCESS = 0

_WINLDAP = None
_WINLDAP_ERR = None


def _log(msg):
    try:
        print("[ad/winldap] %s" % msg)
    except Exception:
        pass


def _text(value):
    if value is None:
        return u""
    try:
        return unicode(value).strip()
    except Exception:
        return str(value).strip()


def is_windows():
    return sys.platform.startswith("win")


def is_winldap_available():
    global _WINLDAP_ERR
    if not is_windows():
        return False
    try:
        _load_wldap32()
        return True
    except Exception as err:
        _WINLDAP_ERR = err
        return False


def winldap_unavailable_reason():
    if not is_windows():
        return u"WinLDAP доступен только на Windows"
    return _text(_WINLDAP_ERR)


def pick_ad_backend(use_ssl):
    """
    winldap: Windows + wldap32, если нужен LDAPS или Python без ssl (AO stub).
    ldap3: иначе.
    """
    use_ssl_flag = bool(use_ssl)
    if not is_winldap_available():
        return u"ldap3"
    try:
        from libre_macros_ad_ssl_stub import ssl_is_stub
    except ImportError:
        ssl_is_stub = lambda: False  # noqa: E731
    if use_ssl_flag or ssl_is_stub():
        return u"winldap"
    return u"ldap3"


def _load_wldap32():
    global _WINLDAP
    if _WINLDAP is not None:
        return _WINLDAP

    import ctypes
    from ctypes import POINTER, Structure, byref, c_char_p, c_int, c_ubyte, c_ulong, c_void_p, wintypes

    dll = ctypes.WinDLL("wldap32")
    ULONG = c_ulong

    class berval(Structure):
        _fields_ = [("bv_len", c_ulong), ("bv_val", c_char_p)]

    class LDAPControlW(Structure):
        _fields_ = [
            ("ldctl_oid", wintypes.LPWSTR),
            ("ldctl_value", berval),
            ("ldctl_iscritical", c_ubyte),
        ]

    PLDAP = c_void_p
    LDAPMessage = c_void_p
    PLDAPControlW = POINTER(LDAPControlW)
    PPLDAPControlW = POINTER(PLDAPControlW)

    dll.ldap_initW.argtypes = [wintypes.LPWSTR, ULONG]
    dll.ldap_initW.restype = PLDAP
    dll.ldap_sslinitW.argtypes = [wintypes.LPWSTR, ULONG, c_int]
    dll.ldap_sslinitW.restype = PLDAP
    dll.ldap_connect.argtypes = [PLDAP, c_void_p]
    dll.ldap_connect.restype = ULONG
    dll.ldap_set_optionW.argtypes = [PLDAP, c_int, c_void_p]
    dll.ldap_set_optionW.restype = ULONG
    dll.ldap_bind_sW.argtypes = [PLDAP, wintypes.LPWSTR, wintypes.LPWSTR, ULONG]
    dll.ldap_bind_sW.restype = ULONG
    dll.ldap_unbind.argtypes = [PLDAP]
    dll.ldap_unbind.restype = ULONG
    dll.ldap_get_last_error.argtypes = []
    dll.ldap_get_last_error.restype = ULONG
    dll.ldap_err2stringW.argtypes = [ULONG]
    dll.ldap_err2stringW.restype = wintypes.LPWSTR
    dll.ldap_first_entry.argtypes = [PLDAP, LDAPMessage]
    dll.ldap_first_entry.restype = c_void_p
    dll.ldap_next_entry.argtypes = [PLDAP, c_void_p]
    dll.ldap_next_entry.restype = c_void_p
    dll.ldap_get_dnW.argtypes = [PLDAP, c_void_p]
    dll.ldap_get_dnW.restype = wintypes.LPWSTR
    dll.ldap_memfreeW.argtypes = [c_void_p]
    dll.ldap_memfreeW.restype = None
    dll.ldap_first_attributeW.argtypes = [PLDAP, c_void_p, POINTER(c_char_p)]
    dll.ldap_first_attributeW.restype = wintypes.LPWSTR
    dll.ldap_next_attributeW.argtypes = [PLDAP, c_void_p, POINTER(c_char_p)]
    dll.ldap_next_attributeW.restype = wintypes.LPWSTR
    dll.ldap_get_values_lenW.argtypes = [PLDAP, c_void_p, wintypes.LPWSTR]
    dll.ldap_get_values_lenW.restype = POINTER(berval)
    dll.ldap_value_free_len.argtypes = [POINTER(berval)]
    dll.ldap_value_free_len.restype = ULONG
    dll.ldap_msgfree.argtypes = [LDAPMessage]
    dll.ldap_msgfree.restype = ULONG
    dll.ldap_create_page_controlW.argtypes = [
        PLDAP, ULONG, POINTER(berval), c_ubyte, POINTER(PLDAPControlW),
    ]
    dll.ldap_create_page_controlW.restype = ULONG
    dll.ldap_parse_page_controlW.argtypes = [
        PLDAP, PPLDAPControlW, POINTER(POINTER(berval)), POINTER(ULONG),
    ]
    dll.ldap_parse_page_controlW.restype = ULONG
    dll.ldap_control_freeW.argtypes = [PLDAPControlW]
    dll.ldap_control_freeW.restype = ULONG
    dll.ldap_parse_resultW.argtypes = [
        PLDAP, LDAPMessage, POINTER(ULONG), POINTER(wintypes.LPWSTR),
        POINTER(wintypes.LPWSTR), POINTER(POINTER(wintypes.LPWSTR)),
        POINTER(PPLDAPControlW), c_ubyte,
    ]
    dll.ldap_parse_resultW.restype = ULONG
    dll.ldap_search_ext_sW.argtypes = [
        PLDAP, wintypes.LPWSTR, ULONG, wintypes.LPWSTR, POINTER(wintypes.LPWSTR),
        ULONG, PPLDAPControlW, PPLDAPControlW, c_void_p, ULONG, POINTER(LDAPMessage),
    ]
    dll.ldap_search_ext_sW.restype = ULONG
    dll.ldap_search_sW.argtypes = [
        PLDAP, wintypes.LPWSTR, ULONG, wintypes.LPWSTR, POINTER(wintypes.LPWSTR),
        ULONG, POINTER(LDAPMessage),
    ]
    dll.ldap_search_sW.restype = ULONG
    dll.ldap_get_next_pageW.argtypes = [PLDAP, LDAPMessage, ULONG, POINTER(LDAPMessage)]
    dll.ldap_get_next_pageW.restype = ULONG
    dll.ldap_get_paged_countW.argtypes = [PLDAP, LDAPMessage, POINTER(ULONG)]
    dll.ldap_get_paged_countW.restype = ULONG

    _WINLDAP = {
        "dll": dll,
        "berval": berval,
        "LDAPControlW": LDAPControlW,
        "PLDAPControlW": PLDAPControlW,
        "PPLDAPControlW": PPLDAPControlW,
        "LDAPMessage": LDAPMessage,
        "byref": byref,
        "c_ulong": c_ulong,
        "c_ubyte": c_ubyte,
        "c_void_p": c_void_p,
        "wintypes": wintypes,
        "POINTER": POINTER,
    }
    return _WINLDAP


class WinLdapError(Exception):
    def __init__(self, code, message, operation=u""):
        self.code = int(code)
        self.operation = _text(operation)
        text = _text(message)
        if self.operation:
            super(WinLdapError, self).__init__(
                u"WinLDAP %s: %s (0x%x)" % (self.operation, text, self.code & 0xFFFFFFFF)
            )
        else:
            super(WinLdapError, self).__init__(u"WinLDAP: %s (0x%x)" % (text, self.code & 0xFFFFFFFF))


def _ldap_error(api, code, operation):
    if code == LDAP_SUCCESS:
        return
    try:
        msg = api["dll"].ldap_err2stringW(code)
        text = msg if msg else u"LDAP error"
    except Exception:
        text = u"LDAP error %s" % code
    raise WinLdapError(code, text, operation)


def _decode_ldap_bytes(raw, length):
    import ctypes

    if not raw or length == 0:
        return u""
    buf = ctypes.string_at(raw, length)
    for enc in ("utf-8", "cp1251", "latin-1"):
        try:
            return buf.decode(enc)
        except Exception:
            continue
    return buf.decode("latin-1", "replace")


class WinLdapSession(object):
    def __init__(self, ld, host, port, use_ssl):
        self._ld = ld
        self.host = _text(host)
        self.port = int(port)
        self.use_ssl = bool(use_ssl)

    def bind_simple(self, user, password):
        api = _load_wldap32()
        code = api["dll"].ldap_bind_sW(
            self._ld, _text(user), password if password is not None else u"", LDAP_AUTH_SIMPLE
        )
        _ldap_error(api, code, u"bind")
        _log(u"bind OK user=%s host=%s:%s ssl=%s" % (user, self.host, self.port, self.use_ssl))

    def unbind(self):
        if not self._ld:
            return
        api = _load_wldap32()
        try:
            api["dll"].ldap_unbind(self._ld)
        finally:
            self._ld = None

    def paged_search(self, base_dn, search_filter, attributes, page_size=500):
        """
        ldap_search_ext_sW + page control (AD OID 1.2.840.113556.1.4.319).
        Yields {"dn", "attributes", "type": "searchResEntry"}.
        """
        api = _load_wldap32()
        base = _text(base_dn)
        filt = _text(search_filter)
        attr_names = [_text(a) for a in (attributes or []) if _text(a)]

        attr_array = (api["wintypes"].LPWSTR * (len(attr_names) + 1))()
        i = 0
        while i < len(attr_names):
            attr_array[i] = attr_names[i]
            i += 1

        cookie_blob = api["berval"]()
        cookie_blob.bv_len = 0
        cookie_blob.bv_val = None
        page_sz = max(1, int(page_size or 500))

        while True:
            ctrl = api["PLDAPControlW"]()
            cookie_ptr = api["byref"](cookie_blob) if cookie_blob.bv_len else None
            code = api["dll"].ldap_create_page_controlW(
                self._ld, page_sz, cookie_ptr, 1, api["byref"](ctrl)
            )
            _ldap_error(api, code, u"create_page_control")

            server_ctrls = (api["PLDAPControlW"] * 2)()
            server_ctrls[0] = ctrl
            result = api["LDAPMessage"]()
            code = api["dll"].ldap_search_ext_sW(
                self._ld,
                base,
                LDAP_SCOPE_SUBTREE,
                filt,
                attr_array,
                0,
                server_ctrls,
                None,
                None,
                0,
                api["byref"](result),
            )
            try:
                _ldap_error(api, code, u"search")
                for entry in self._iter_entries(api, result):
                    yield entry
                cookie_out = api["POINTER"](api["berval"])()
                total = api["c_ulong"](0)
                resp_ctrls = api["PPLDAPControlW"]()
                api["dll"].ldap_parse_resultW(
                    self._ld, result, None, None, None, None, api["byref"](resp_ctrls), 0
                )
                if resp_ctrls:
                    parse_code = api["dll"].ldap_parse_page_controlW(
                        self._ld, resp_ctrls, api["byref"](cookie_out), api["byref"](total)
                    )
                    if parse_code == LDAP_SUCCESS and cookie_out and cookie_out[0]:
                        cookie_blob = cookie_out[0][0]
                        if cookie_blob.bv_len:
                            continue
                break
            finally:
                if result:
                    api["dll"].ldap_msgfree(result)
                try:
                    api["dll"].ldap_control_freeW(ctrl)
                except Exception:
                    pass
            break

    def search_simple(self, base_dn, search_filter, attributes):
        """Один запрос без paging (fallback, лимит сервера ~1000)."""
        api = _load_wldap32()
        attr_names = [_text(a) for a in (attributes or []) if _text(a)]
        attr_array = (api["wintypes"].LPWSTR * (len(attr_names) + 1))()
        i = 0
        while i < len(attr_names):
            attr_array[i] = attr_names[i]
            i += 1
        result = api["LDAPMessage"]()
        code = api["dll"].ldap_search_sW(
            self._ld,
            _text(base_dn),
            LDAP_SCOPE_SUBTREE,
            _text(search_filter),
            attr_array,
            0,
            api["byref"](result),
        )
        try:
            _ldap_error(api, code, u"search")
            for entry in self._iter_entries(api, result):
                yield entry
        finally:
            if result:
                api["dll"].ldap_msgfree(result)

    def _iter_entries(self, api, message):
        entry = api["dll"].ldap_first_entry(self._ld, message)
        while entry:
            dn_ptr = api["dll"].ldap_get_dnW(self._ld, entry)
            dn = dn_ptr if dn_ptr else u""
            try:
                yield {
                    u"dn": dn,
                    u"attributes": self._read_attributes(api, entry),
                    u"type": u"searchResEntry",
                }
            finally:
                if dn_ptr:
                    api["dll"].ldap_memfreeW(dn_ptr)
            entry = api["dll"].ldap_next_entry(self._ld, entry)

    def _read_attributes(self, api, entry):
        out = {}
        ber_ptr = api["c_char_p"]()
        attr = api["dll"].ldap_first_attributeW(self._ld, entry, api["byref"](ber_ptr))
        while attr:
            name = _text(attr)
            vals = api["dll"].ldap_get_values_lenW(self._ld, entry, attr)
            if vals:
                items = []
                vi = 0
                while vals[vi].bv_val is not None:
                    items.append(_decode_ldap_bytes(vals[vi].bv_val, vals[vi].bv_len))
                    vi += 1
                if len(items) == 1:
                    out[name] = items[0]
                elif len(items) > 1:
                    out[name] = items
                api["dll"].ldap_value_free_len(vals)
            api["dll"].ldap_memfreeW(attr)
            attr = api["dll"].ldap_next_attributeW(self._ld, entry, api["byref"](ber_ptr))
        return out


def connect_winldap(server, port, use_ssl=False, timeout_sec=30):
    if not is_winldap_available():
        raise WinLdapError(0, winldap_unavailable_reason() or u"wldap32 недоступен", u"connect")

    import ctypes

    api = _load_wldap32()
    host = _text(server)
    if host == u"":
        raise ValueError(u"Не задан сервер LDAP/AD (DC).")
    port_num = int(port)
    use_ssl_flag = bool(use_ssl)

    if use_ssl_flag:
        ld = api["dll"].ldap_sslinitW(host, port_num, 1)
    else:
        ld = api["dll"].ldap_initW(host, port_num)
    if not ld:
        _ldap_error(api, api["dll"].ldap_get_last_error(), u"init")

    version = api["c_ulong"](LDAP_VERSION3)
    api["dll"].ldap_set_optionW(ld, LDAP_OPT_PROTOCOL_VERSION, api["byref"](version))
    off = api["c_ulong"](LDAP_OPT_OFF)
    api["dll"].ldap_set_optionW(ld, LDAP_OPT_REFERRALS, api["byref"](off))

    code = api["dll"].ldap_connect(ld, None)
    _ldap_error(api, code, u"connect")
    _log(u"connected %s:%s ssl=%s" % (host, port_num, use_ssl_flag))
    return WinLdapSession(ld, host, port_num, use_ssl_flag)


def search_users_winldap( server, bind_user, bind_password, ou_dn, use_ssl=False, port=None, page_size=500, search_filter=None, attributes=None):
    from libre_macros_ad_cfg import (
        DEFAULT_PORT,
        DEFAULT_SSL_PORT,
        USER_LDAP_ATTRIBUTES,
        USER_SEARCH_FILTER,
    )

    use_ssl_flag = bool(use_ssl)
    if port in (None, u""):
        port_num = DEFAULT_SSL_PORT if use_ssl_flag else DEFAULT_PORT
    else:
        port_num = int(port)

    session = connect_winldap(server, port_num, use_ssl=use_ssl_flag)
    filt = _text(search_filter or USER_SEARCH_FILTER)
    attrs = tuple(attributes or USER_LDAP_ATTRIBUTES)
    entries = []
    try:
        session.bind_simple(bind_user, bind_password)
        try:
            for item in session.paged_search(ou_dn, filt, attrs, page_size=page_size):
                entries.append(item)
        except WinLdapError as err:
            _log(u"paged_search failed (%s), fallback search_simple" % err)
            for item in session.search_simple(ou_dn, filt, attrs):
                entries.append(item)
    finally:
        session.unbind()
    return entries
