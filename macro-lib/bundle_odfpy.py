#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка плоского odf_bundled.py для macro-lib/pythonpath.

Требуется в .venv:
  pip install odfpy

Запуск:
  python macro-lib/bundle_odfpy.py

После сборки патчит cnv_family: LO/AO style:family=smart-table|pivot-table.
"""

from __future__ import annotations

import sys
from pathlib import Path

from bundle_python_lib import BundleSpec, main_cli

MACRO_VERSION = "3.10.701"
SPEC = BundleSpec(
    packages=("defusedxml", "odf"),
    bundle_tag="odf_bundled",
    title="odfpy (import odf) + defusedxml в одном файле для LibreOffice pythonpath.",
    build_cmd="python macro-lib/bundle_odfpy.py",
    import_hint="import odf_bundled",
    dist_names={"odf": "odfpy"},
)


def _patch_odf_bundled_family(path: Path) -> bool:
    """Разрешить LO/AO style:family smart-table / pivot-table в cnv_family."""
    text = path.read_text(encoding="utf-8")
    if '"smart-table", "pivot-table"' in text or '"smart-table", "pivot-table"' in text:
        # уже есть
        if "smart-table" in text and "cnv_family" in text:
            return False
    needle = '"drawing-page", "chart"):'
    repl = '"drawing-page", "chart", "smart-table", "pivot-table"):'
    if needle not in text:
        raise SystemExit(
            "bundle_odfpy: не найден фрагмент cnv_family для патча smart-table"
        )
    if "smart-table" in text.split("def cnv_family", 1)[-1].split("def ", 1)[0]:
        return False
    path.write_text(text.replace(needle, repl, 1), encoding="utf-8")
    return True


if __name__ == "__main__":
    macro_dir = Path(__file__).resolve().parent
    out = macro_dir / "pythonpath" / "odf_bundled.py"
    rc = main_cli(
        SPEC,
        default_output=out,
        description="Собрать odf_bundled.py",
    )
    if rc == 0 and out.is_file():
        if _patch_odf_bundled_family(out):
            print("odf_bundled: cnv_family += smart-table, pivot-table")
        else:
            print("odf_bundled: cnv_family already patched")
    sys.exit(rc)
