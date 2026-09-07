# -*- coding: utf-8 -*-
"""Общие функции CLI для ad-linux (параметры, getpass, CSV)."""
from __future__ import annotations

import argparse
import csv
import getpass
import sys
from pathlib import Path

DEFAULT_DELIMITER = ","


from dn_format import normalize_base_dn


def add_connection_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("-s", "--server", help="Хост контроллера домена (DC)")
    parser.add_argument("-u", "--user", help="Пользователь bind (DOMAIN\\user или UPN)")
    parser.add_argument("-p", "--password", help="Пароль (лучше getpass, не указывать в командной строке)")
    parser.add_argument(
        "-o",
        "--ou",
        help="Base DN: OU=…,DC=… или путь Sales/Users/corp/local или name1/name2/dc1/c/dc2",
    )
    parser.add_argument(
        "--dc-parts",
        type=int,
        default=None,
        metavar="N",
        help="Сколько последних сегментов пути — DC (auto по -s, иначе 2; для dc1/c/dc2 → 3)",
    )
    parser.add_argument(
        "-a",
        "--auth",
        choices=("NTLM", "SIMPLE"),
        default="NTLM",
        help="Метод bind (по умолчанию NTLM)",
    )
    parser.add_argument(
        "--ssl",
        dest="ssl",
        action="store_true",
        default=True,
        help="LDAPS (по умолчанию включено)",
    )
    parser.add_argument(
        "--no-ssl",
        dest="ssl",
        action="store_false",
        help="Plain LDAP (порт 389)",
    )
    parser.add_argument(
        "-P",
        "--port",
        type=int,
        default=None,
        help="Порт (по умолчанию 636 при --ssl, иначе 389)",
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=500,
        help="Размер страницы paged search",
    )
    parser.add_argument(
        "-f",
        "--output",
        type=Path,
        help="Файл результата CSV. Без параметра — stdout",
    )
    parser.add_argument(
        "--delimiter",
        default=DEFAULT_DELIMITER,
        help="Разделитель CSV (по умолчанию запятая)",
    )
    parser.add_argument(
        "-i",
        "--interactive",
        action="store_true",
        help="Спросить в консоли все поля",
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="Не выводить служебные сообщения в stderr",
    )


def prompt(label: str, default: str = "") -> str:
    hint = " [%s]" % default if default else ""
    try:
        value = input("%s%s: " % (label, hint)).strip()
    except EOFError:
        value = ""
    if not value and default:
        return default
    return value


def resolve_connection_params(args) -> tuple[str, str, str, str]:
    server = (args.server or "").strip()
    user = (args.user or "").strip()
    ou = (args.ou or "").strip()
    password = args.password

    need_prompt = args.interactive or not (server and user and ou)
    if need_prompt and sys.stdin.isatty():
        if args.interactive or not server:
            server = prompt("Сервер DC", server)
        if args.interactive or not user:
            user = prompt("Пользователь (DOMAIN\\user или UPN)", user)
        if args.interactive or not ou:
            ou = prompt("OU (DN или Sales/Users/corp/local)", ou)
        if args.interactive:
            auth_in = prompt("Auth (NTLM/SIMPLE)", args.auth)
            if auth_in:
                args.auth = auth_in.upper()
            ssl_in = prompt("SSL/LDAPS? (y/n)", "y" if args.ssl else "n").lower()
            if ssl_in in ("n", "no", "0", "нет"):
                args.ssl = False
            elif ssl_in in ("y", "yes", "1", "да", ""):
                args.ssl = True

    if not server:
        raise SystemExit("Не задан --server")
    if not user:
        raise SystemExit("Не задан --user")
    if not ou:
        raise SystemExit("Не задан --ou")

    ou = normalize_base_dn(ou, server=server, dc_parts=getattr(args, "dc_parts", None))

    if password is None:
        if not sys.stdin.isatty():
            raise SystemExit("Не задан --password (stdin не TTY, getpass недоступен)")
        password = getpass.getpass("Пароль: ")

    return server, user, password, ou


def log_stderr(msg: str, quiet: bool) -> None:
    if not quiet:
        print(msg, file=sys.stderr)


def write_csv_rows(rows, columns: tuple[str, ...], output: Path | None, delimiter: str) -> None:
    headers = list(columns)
    if output:
        with output.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(
                fh,
                fieldnames=headers,
                delimiter=delimiter,
                extrasaction="ignore",
                lineterminator="\n",
            )
            writer.writeheader()
            writer.writerows(rows)
    else:
        writer = csv.DictWriter(
            sys.stdout,
            fieldnames=headers,
            delimiter=delimiter,
            extrasaction="ignore",
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(rows)
