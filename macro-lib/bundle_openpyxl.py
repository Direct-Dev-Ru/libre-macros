MACRO_VERSION = "3.10.703"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка плоского openpyxl_bundled.py для macro-lib/pythonpath.

Запуск:
  python macro-lib/bundle_openpyxl.py
"""

from __future__ import annotations

import sys
from pathlib import Path

from bundle_python_lib import BundleSpec, main_cli

SPEC = BundleSpec(
    packages=("et_xmlfile", "openpyxl"),
    bundle_tag="openpyxl_bundled",
    title="openpyxl + et_xmlfile в одном файле для LibreOffice pythonpath.",
    build_cmd="python macro-lib/bundle_openpyxl.py",
    import_hint="import openpyxl_bundled",
)

if __name__ == "__main__":
    macro_dir = Path(__file__).resolve().parent
    sys.exit(
        main_cli(
            SPEC,
            default_output=macro_dir / "pythonpath" / "openpyxl_bundled.py",
            description="Собрать openpyxl_bundled.py",
        )
    )
