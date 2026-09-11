#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка плоского segno_bundled.py для macro-lib/pythonpath.

Запуск:
  python macro-lib/bundle_segno.py
"""

from __future__ import annotations

MACRO_VERSION = "3.10.726"
import sys
from pathlib import Path

from bundle_python_lib import BundleSpec, main_cli

SPEC = BundleSpec(
    packages=("segno",),
    bundle_tag="segno_bundled",
    title="segno (QR-коды) в одном файле для LibreOffice pythonpath.",
    build_cmd="python macro-lib/bundle_segno.py",
    import_hint="import segno_bundled",
)

if __name__ == "__main__":
    macro_dir = Path(__file__).resolve().parent
    sys.exit(
        main_cli(
            SPEC,
            default_output=macro_dir / "pythonpath" / "segno_bundled.py",
            description="Собрать segno_bundled.py",
        )
    )
