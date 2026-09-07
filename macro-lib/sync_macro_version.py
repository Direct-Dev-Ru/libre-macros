from __future__ import print_function
MACRO_VERSION = "3.10.681"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Синхронизация MACRO_VERSION во всех .py пакета macro-lib с macro-lib/version.txt.

Запуск из корня репозитория:
  python macro-lib/sync_macro_version.py
  python macro-lib/sync_macro_version.py --check   # только проверка, без записи
"""


import argparse
import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
MACRO_LIB = os.path.join(ROOT, "macro-lib")
VERSION_TXT = os.path.join(MACRO_LIB, "version.txt")

sys.path.insert(0, os.path.join(ROOT, "installer"))
from build_ods_bundled import (  # noqa: E402
    ZIP_EXCLUDE_PY,
    _iter_packable_py_files,
    sync_macro_version_tree,
)


def read_version_txt():
    try:
        with open(VERSION_TXT, encoding="utf-8") as f:
            return f.read().strip()
    except (IOError, OSError):
        return ""


def main():
    parser = argparse.ArgumentParser(
        description="Синхронизировать MACRO_VERSION в macro-lib с version.txt"
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Только показать рассинхрон, не записывать файлы",
    )
    parser.add_argument(
        "--version",
        metavar="X.Y.Z",
        help="Версия (по умолчанию — содержимое macro-lib/version.txt)",
    )
    args = parser.parse_args()

    version = (args.version or read_version_txt()).strip()
    if not version:
        print("Ошибка: не задана версия и version.txt пуст.", file=sys.stderr)
        return 1

    if args.check:
        import re

        pat = re.compile(r'^MACRO_VERSION\s*=\s*["\']([^"\']+)["\']', re.M)
        exclude = set(ZIP_EXCLUDE_PY)
        exclude.add("sync_macro_version.py")
        bad = []
        for path in _iter_packable_py_files(MACRO_LIB, exclude):
            try:
                with open(path, encoding="utf-8") as f:
                    text = f.read()
            except (IOError, OSError, UnicodeDecodeError):
                continue
            m = pat.search(text)
            if m and m.group(1) != version:
                bad.append((os.path.relpath(path, ROOT), m.group(1)))
        if not bad:
            print("OK: все MACRO_VERSION = %s" % version)
            return 0
        print("Рассинхрон с version.txt (%s):" % version)
        for rel, v in sorted(bad):
            print("  %s → %s" % (rel, v))
        return 1

    if not args.version:
        pass  # version уже из version.txt
    else:
        with open(VERSION_TXT, "w", encoding="utf-8") as f:
            f.write(version + "\n")
        print("version.txt → %s" % version)

    updated = sync_macro_version_tree(
        MACRO_LIB,
        version,
        exclude_py=set(ZIP_EXCLUDE_PY) | {"sync_macro_version.py"},
    )
    if updated:
        print("Обновлено файлов: %d" % len(updated))
        for path in sorted(updated):
            print("  %s" % os.path.relpath(path, ROOT))
    else:
        print("Все MACRO_VERSION уже = %s" % version)
    return 0


if __name__ == "__main__":
    sys.exit(main())
