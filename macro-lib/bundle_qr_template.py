from __future__ import print_function
MACRO_VERSION = "3.10.688"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Пересобрать pythonpath/libre_macros_qr_template_blob.py из generate_qr_template.ods."""
import base64
import sys
from pathlib import Path

MACRO_DIR = Path(__file__).resolve().parent
ODS = MACRO_DIR / "generate_qr_template.ods"
OUT = MACRO_DIR / "pythonpath" / "libre_macros_qr_template_blob.py"


def main():
    if not ODS.is_file():
        print("Нет файла:", ODS, file=sys.stderr)
        return 1
    data = ODS.read_bytes()
    b64 = base64.b64encode(data).decode("ascii")
    wrapped = "\n".join(b64[i : i + 76] for i in range(0, len(b64), 76))
    OUT.write_text(
        "# -*- coding: utf-8 -*-\n"
        '"""Вшитый generate_qr_template.ods (base64). Пересборка: python3 macro-lib/bundle_qr_template.py"""\n'
        "from __future__ import print_function, unicode_literals\n\n"
        "import base64\n"
        "import os\n\n"
        'QR_TEMPLATE_ODS_NAME = "generate_qr_template.ods"\n'
        'QR_TEMPLATE_DEFAULT_SHEET = u"Лист1"\n\n'
        "# binary length %d\n"
        "QR_TEMPLATE_ODS_B64 = (\n"
        '"""%s\n"""\n'
        ")\n\n\n"
        "def qr_template_ods_bytes():\n"
        '    return base64.b64decode(QR_TEMPLATE_ODS_B64.encode("ascii"))\n\n\n'
        "def write_qr_template_ods(path):\n"
        "    data = qr_template_ods_bytes()\n"
        "    path = os.path.abspath(str(path))\n"
        "    parent = os.path.dirname(path)\n"
        "    if parent and not os.path.isdir(parent):\n"
        "        os.makedirs(parent)\n"
        '    with open(path, "wb") as f:\n'
        "        f.write(data)\n"
        "    return path\n"
        % (len(data), wrapped),
        encoding="utf-8",
    )
    print("OK:", OUT, "(%d bytes)" % len(data))
    return 0


if __name__ == "__main__":
    sys.exit(main())
