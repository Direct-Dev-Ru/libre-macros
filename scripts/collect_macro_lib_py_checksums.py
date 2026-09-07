#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Собрать список .py файлов macro-lib с относительными путями и SHA-256.

Пути — относительно каталога macro-lib.
По умолчанию пишет JSON рядом со скриптом.

Запуск из корня репозитория или из scripts/:
  python3 scripts/collect_macro_lib_py_checksums.py
  python3 scripts/collect_macro_lib_py_checksums.py -o /tmp/py_files.json
  python3 scripts/collect_macro_lib_py_checksums.py --root /path/to/macro-lib
"""

from __future__ import print_function

import argparse
import hashlib
import json
import os
import sys
from datetime import datetime, timezone

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DEFAULT_ROOT = os.path.join(REPO_ROOT, "macro-lib")
DEFAULT_OUTPUT = os.path.join(SCRIPT_DIR, "macro_lib_py_checksums.json")

# Каталоги, которые не обходим.
_SKIP_DIR_NAMES = frozenset(
    {
        "__pycache__",
        ".git",
        ".venv",
        "venv",
        "node_modules",
        ".tox",
        ".mypy_cache",
        ".pytest_cache",
    }
)


def _sha256_file(path, chunk_size=1024 * 1024):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _rel_posix(root, path):
    rel = os.path.relpath(path, root)
    return rel.replace(os.sep, "/")


def iter_py_files(root):
    """Все .py под root (без __pycache__ и типичного мусора)."""
    root = os.path.abspath(root)
    out = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in _SKIP_DIR_NAMES and not d.startswith(".")]
        for name in filenames:
            if name.startswith("."):
                continue
            # .py везде; .txt только в pythonpath (например terms_of_use_ru.txt)
            lower = name.casefold()
            if lower.endswith(".py"):
                pass
            elif lower.endswith(".txt"):
                rel_dir = os.path.relpath(dirpath, root).replace(os.sep, "/")
                if rel_dir != "pythonpath" and not rel_dir.startswith("pythonpath/"):
                    continue
            else:
                continue
            full = os.path.join(dirpath, name)
            if not os.path.isfile(full):
                continue
            out.append(full)
    out.sort(key=lambda p: _rel_posix(root, p).casefold())
    return out


def collect(root):
    root = os.path.abspath(root)
    if not os.path.isdir(root):
        raise SystemExit("каталог не найден: %s" % root)
    files = []
    for full in iter_py_files(root):
        rel = _rel_posix(root, full)
        try:
            size = os.path.getsize(full)
        except OSError as err:
            print("пропуск %s: %s" % (rel, err), file=sys.stderr)
            continue
        try:
            digest = _sha256_file(full)
        except OSError as err:
            print("пропуск %s: %s" % (rel, err), file=sys.stderr)
            continue
        files.append(
            {
                "path": rel,
                "sha256": digest,
                "size": size,
            }
        )
    return {
        "root": root,
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "algo": "sha256",
        "count": len(files),
        "files": files,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Список .py в macro-lib с относительными путями и SHA-256 → JSON"
    )
    parser.add_argument(
        "--root",
        default=DEFAULT_ROOT,
        help="корень обхода (по умолчанию: …/macro-lib)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_OUTPUT,
        help="путь к JSON (по умолчанию: scripts/macro_lib_py_checksums.json)",
    )
    parser.add_argument(
        "--indent",
        type=int,
        default=2,
        help="отступ JSON (0 — компактно)",
    )
    args = parser.parse_args(argv)

    data = collect(args.root)
    out_path = os.path.abspath(args.output)
    out_dir = os.path.dirname(out_path)
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir)

    indent = None if int(args.indent) <= 0 else int(args.indent)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=indent)
        f.write("\n")

    print("файлов: %d" % data["count"])
    print("записано: %s" % out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
