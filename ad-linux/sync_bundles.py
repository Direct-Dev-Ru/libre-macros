#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скопировать bundled-зависимости из macro-lib в ad-linux/bundled/.

Запуск из корня репозитория:
  python3 ad-linux/sync_bundles.py

Или после пересборки бандла:
  python3 macro-lib/bundle_ldap3.py && python3 ad-linux/sync_bundles.py
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "macro-lib" / "pythonpath"
DST = Path(__file__).resolve().parent / "bundled"

BUNDLES = (
    "ldap3_bundled.py",
)


def main() -> int:
    DST.mkdir(parents=True, exist_ok=True)
    copied = 0
    for name in BUNDLES:
        src = SRC / name
        dst = DST / name
        if not src.is_file():
            print("Нет исходника: %s (сначала python3 macro-lib/bundle_ldap3.py)" % src, file=sys.stderr)
            return 1
        shutil.copy2(src, dst)
        size_kb = dst.stat().st_size // 1024
        print("OK: %s -> %s (%d Кб)" % (src, dst, size_kb))
        copied += 1
    print("Скопировано файлов: %d" % copied)
    return 0


if __name__ == "__main__":
    sys.exit(main())
