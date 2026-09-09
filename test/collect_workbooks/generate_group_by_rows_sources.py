#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Источник для ручной проверки группировать_строки.

sources/group_by_rows/source_group_by_rows.xlsx
  Продажи — Отдел / Месяц / Сумма / ID (+ пустой ключ)

Сценарий: 17_group_by_rows.xltx
"""
from __future__ import print_function, unicode_literals

import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
TEST_DIR = os.path.join(ROOT, "test", "collect_workbooks")
OUT_DIR = os.path.join(TEST_DIR, "sources", "group_by_rows")
OUT_XLSX = os.path.join(OUT_DIR, "source_group_by_rows.xlsx")

try:
    from openpyxl import Workbook
except ImportError:
    sys.path.insert(0, os.path.join(ROOT, "macro-lib", "pythonpath"))
    import openpyxl_bundled  # noqa: F401
    from openpyxl import Workbook


def generate_group_by_rows_sources(out_path=None):
    out_path = out_path or OUT_XLSX
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(u"Продажи")
    ws.append([u"Отдел", u"Месяц", u"Сумма", u"ID", u"Комментарий"])
    rows = [
        (u"МСУЗ", u"Январь", 100, u"A1", u"ок"),
        (u"МСУЗ", u"Январь", 50, u"A2", u""),
        (u"мсуз", u"Январь", 10, u"A3", u"регистр"),  # same key if casefold
        (u"МСУЗ", u"Февраль", 200, u"B1", u""),
        (u"ИТ", u"Январь", 30, u"C1", u""),
        (u"ИТ", u"Январь", 70, u"C2", u""),
        (u"  ИТ  ", u"Январь", 5, u"C3", u"trim"),
        (u"", u"Март", 999, u"X1", u"пустой ключ"),
        (None, u"Март", 1, u"X2", u"пустой ключ 2"),
        (u"Склад", u"Февраль", 40, u"D1", u""),
        (u"Склад", u"Февраль", u"нечисло", u"D2", u"skip sum"),
        (u"Склад", u"Февраль", 60, u"D3", u""),
    ]
    for r in rows:
        ws.append(list(r))

    wb.save(out_path)
    readme = os.path.join(os.path.dirname(out_path), "README.md")
    with open(readme, "w", encoding="utf-8") as f:
        f.write(
            u"# Источник группировать_строки\n\n"
            u"`source_group_by_rows.xlsx`:\n\n"
            u"- **Продажи** — ключи Отдел+Месяц, сумма/count; есть пустые ключи и "
            u"регистр/пробелы для проверки `key_trim` / `key_case_sensitive` / "
            u"`skip_empty_keys`.\n\n"
            u"Ожидание (inplace, ключ Отдел+Месяц, sum Сумма + count(*), "
            u"skip_empty_keys=true, key_case_sensitive=false, key_trim=true):\n\n"
            u"| Отдел | Месяц | Сумма | Количество |\n"
            u"|-------|-------|-------|------------|\n"
            u"| МСУЗ | Январь | 160 | 3 |\n"
            u"| МСУЗ | Февраль | 200 | 1 |\n"
            u"| ИТ | Январь | 105 | 3 |\n"
            u"| Склад | Февраль | 100 | 3 |\n\n"
            u"Сценарий: `17_group_by_rows.xltx`\n"
        )
    print("  %s" % out_path)
    return out_path


def main():
    print("Источники group_by_rows:")
    generate_group_by_rows_sources()


if __name__ == "__main__":
    main()
