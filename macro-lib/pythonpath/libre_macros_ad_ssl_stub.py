# -*- coding: utf-8 -*-
"""
Заглушка ssl для embedded Python LibreOffice / AlterOffice без TLS.

ldap3 импортирует ldap3.core.tls при загрузке; без ssl модуль ldap3 не стартует.
Для обычного LDAP (порт 389, use_ssl=False) достаточно stub с константами CERT_*.
LDAPS/start_tls при stub вызовет понятную ошибку из ldap3 или connect_ad.
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.704"
def _ssl_usable(mod):
    if mod is None:
        return False
    for name in ("CERT_NONE", "CERT_OPTIONAL", "CERT_REQUIRED", "wrap_socket"):
        if not hasattr(mod, name):
            return False
    if getattr(mod, "_lm_ad_stub", False):
        return False
    return True


def _build_ssl_stub():
    import types

    ssl = types.ModuleType("ssl")
    ssl._lm_ad_stub = True
    ssl.CERT_NONE = 0
    ssl.CERT_OPTIONAL = 1
    ssl.CERT_REQUIRED = 2
    ssl.PROTOCOL_SSLv23 = 2
    ssl.SSLError = type("SSLError", (Exception,), {})

    def _reject_ssl(*_args, **_kwargs):
        raise ImportError(
            "SSL not supported in this Python interpreter (LibreOffice/AlterOffice stub)"
        )

    ssl.wrap_socket = _reject_ssl
    ssl.create_default_context = _reject_ssl

    class _Purpose(object):
        SERVER_AUTH = None

    ssl.Purpose = _Purpose
    ssl.SSLContext = type(
        "SSLContext",
        (object,),
        {"__init__": lambda self, *a, **k: None},
    )
    return ssl


def install_ssl_stub_if_needed():
    """
    Подменить ssl заглушкой, если в интерпретаторе LO/AO модуль отсутствует или неполный.
    Возвращает True, если установлена заглушка (только plain LDAP).
    """
    import sys

    existing = sys.modules.get("ssl")
    if _ssl_usable(existing):
        return False

    try:
        import ssl as probe

        if _ssl_usable(probe):
            return False
    except Exception:
        pass

    stub = _build_ssl_stub()
    sys.modules["ssl"] = stub
    return True


def ssl_is_stub():
    import sys

    mod = sys.modules.get("ssl")
    return bool(getattr(mod, "_lm_ad_stub", False))
