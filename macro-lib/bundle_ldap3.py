#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка плоского ldap3_bundled.py для macro-lib/pythonpath.

Запуск:
  python macro-lib/bundle_ldap3.py
"""

from __future__ import annotations

MACRO_VERSION = "3.10.697"
import sys
from pathlib import Path

from bundle_python_lib import BundleSpec, main_cli

SPEC = BundleSpec(
    packages=("pyasn1", "ldap3"),
    bundle_tag="ldap3_bundled",
    title="ldap3 + pyasn1 (LDAP/Active Directory, SIMPLE/NTLM) для LibreOffice pythonpath.",
    build_cmd="python macro-lib/bundle_ldap3.py",
    import_hint="import ldap3_bundled",
    dist_names={"pyasn1": "pyasn1", "ldap3": "ldap3"},
)

_TLS_IMPORT_NEEDLE = """try:
    # noinspection PyUnresolvedReferences
    import ssl
except ImportError:
    if log_enabled(ERROR):
        log(ERROR, 'SSL not supported in this Python interpreter')
    raise LDAPSSLNotSupportedError('SSL not supported in this Python interpreter')"""

_TLS_IMPORT_REPL = """try:
    # noinspection PyUnresolvedReferences
    import ssl
except ImportError:
    ssl = None

if ssl is None:
    import sys as _sys
    import types as _types
    ssl = _types.ModuleType('ssl')
    ssl._lm_ad_stub = True
    ssl.CERT_NONE = 0
    ssl.CERT_OPTIONAL = 1
    ssl.CERT_REQUIRED = 2
    ssl.PROTOCOL_SSLv23 = 2
    ssl.SSLError = type('SSLError', (Exception,), {})
    def _lm_ad_ssl_reject(*_a, **_k):
        raise LDAPSSLNotSupportedError('SSL not supported in this Python interpreter')
    ssl.wrap_socket = _lm_ad_ssl_reject
    ssl.create_default_context = _lm_ad_ssl_reject
    class _LmAdPurpose(object):
        SERVER_AUTH = None
    ssl.Purpose = _LmAdPurpose
    ssl.SSLContext = type('SSLContext', (object,), {'__init__': lambda self, *a, **k: None})
    _sys.modules['ssl'] = ssl
    if log_enabled(BASIC):
        log(BASIC, 'ssl unavailable — plain LDAP only (LO/AO stub)')"""


def _patch_ldap3_bundled_tls(path: Path) -> bool:
    """Не падать при import ldap3, если в LO/AO нет ssl (plain LDAP 389)."""
    text = path.read_text(encoding="utf-8")
    if "_lm_ad_stub" in text and "_lm_ad_ssl_reject" in text:
        return False
    if _TLS_IMPORT_NEEDLE not in text:
        raise SystemExit(
            "bundle_ldap3: не найден фрагмент tls.py для патча ssl (версия ldap3?)"
        )
    path.write_text(text.replace(_TLS_IMPORT_NEEDLE, _TLS_IMPORT_REPL, 1), encoding="utf-8")
    return True


if __name__ == "__main__":
    macro_dir = Path(__file__).resolve().parent
    out = macro_dir / "pythonpath" / "ldap3_bundled.py"
    rc = main_cli(
        SPEC,
        default_output=out,
        description="Собрать ldap3_bundled.py",
    )
    if rc == 0 and out.is_file():
        if _patch_ldap3_bundled_tls(out):
            print("ldap3_bundled: tls.py patched for LO/AO without ssl")
        else:
            print("ldap3_bundled: tls.py ssl patch already applied")
    sys.exit(rc)
