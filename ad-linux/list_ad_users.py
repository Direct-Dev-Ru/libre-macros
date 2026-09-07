#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CLI: пользователи AD из OU → CSV (ldap3_bundled, SSL/NTLM)."""
from __future__ import annotations

import argparse
import sys

from ad_cli_common import (
    add_connection_args,
    log_stderr,
    resolve_connection_params,
    write_csv_rows,
)
from ldap3_ad import OUTPUT_COLUMNS, search_users


def _parse_args(argv):
    parser = argparse.ArgumentParser(
        description="Список пользователей Active Directory из OU → CSV."
    )
    add_connection_args(parser)
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = _parse_args(argv or sys.argv[1:])
    password = None
    try:
        server, user, password, ou = resolve_connection_params(args)
        log_stderr(
            "Подключение: %s port=%s ssl=%s auth=%s ou=%s"
            % (server, args.port or ("636" if args.ssl else "389"), args.ssl, args.auth, ou),
            args.quiet,
        )
        rows = search_users(
            server,
            user,
            password,
            ou,
            auth=args.auth,
            use_ssl=args.ssl,
            port=args.port,
            page_size=args.page_size,
            dc_parts=getattr(args, "dc_parts", None),
        )
        write_csv_rows(rows, OUTPUT_COLUMNS, args.output, args.delimiter)
        log_stderr("Найдено пользователей: %d" % len(rows), args.quiet)
        return 0
    except KeyboardInterrupt:
        print("\nОтменено.", file=sys.stderr)
        return 130
    except SystemExit:
        raise
    except Exception as err:
        print("Ошибка: %s" % err, file=sys.stderr)
        return 1
    finally:
        password = None


if __name__ == "__main__":
    sys.exit(main())
