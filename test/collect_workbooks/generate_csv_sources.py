#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор CSV-источников для ручного теста collect_workbooks (Доп_Параметры_Источника).

Уровни объёма: tiny → small → medium → large → xlarge.

Запуск:
  python test/collect_workbooks/generate_csv_sources.py
  python test/collect_workbooks/generate_csv_sources.py --tier medium

Создаёт sources/csv/source_csv_{tier}.csv (+ README.md).
"""
from __future__ import print_function, unicode_literals

import argparse
import csv
import io
import os
import random
import sys
from datetime import datetime, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT_DIR = os.path.join(BASE_DIR, "sources", "csv")

# rows — число строк данных (без заголовка)
_VOLUME = {
    "tiny": {"rows": 25, "seed": 11, "delimiter": ";", "encoding": "cp1251"},
    "small": {"rows": 250, "seed": 22, "delimiter": ",", "encoding": "cp1251"},
    "medium": {"rows": 2500, "seed": 33, "delimiter": ";", "encoding": "cp1251"},
    "large": {"rows": 25000, "seed": 44, "delimiter": ";", "encoding": "cp1251"},
    "xlarge": {"rows": 100000, "seed": 55, "delimiter": ";", "encoding": "cp1251"},
}

HEADERS = [
    "ФИО",
    "Табельный_номер",
    "Отдел",
    "Сумма_руб",
    "Сумма_бонус",
    "маркер",
    "ДатаВремя_строка",
    "Филиал",
]

DEPARTMENTS = ("Продажи", "IT", "Склад", "Финансы", "HR")
BRANCHES = ("Северный", "Южный", "Центральный", "Восточный")
SURNAMES = (
    "Иванов",
    "Петров",
    "Сидоров",
    "Козлова",
    "Новикова",
    "Смирнов",
    "Кузнецов",
    "Попова",
)


def _row(rng, index, base_sum):
    dep = DEPARTMENTS[index % len(DEPARTMENTS)]
    branch = BRANCHES[index % len(BRANCHES)]
    fam = SURNAMES[index % len(SURNAMES)]
    day = datetime(2024, 1, 1) + timedelta(days=(index % 365), hours=(index % 24))
    return [
        "%s_%05d" % (fam, index),
        "T%06d" % index,
        dep,
        base_sum + (index % 97) * 100,
        500 + (index % 40) * 25,
        "x" if index % 17 == 0 else "",
        day.strftime("%d.%m.%Y %H:%M:%S"),
        branch,
    ]


def write_csv(path, tier_name, cfg):
    rows_n = int(cfg["rows"])
    delim = cfg["delimiter"]
    enc = cfg["encoding"]
    rng = random.Random(int(cfg["seed"]))
    base_sum = 10000 + int(cfg["seed"]) * 100
    # csv в Py3 пишет str; newline='' обязателен
    kwargs = {"encoding": enc, "newline": ""}
    try:
        f = open(path, "w", **kwargs)
    except TypeError:
        # Py2
        import codecs

        f = codecs.open(path, "w", encoding=enc)
    try:
        writer = csv.writer(f, delimiter=delim, lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
        writer.writerow(HEADERS)
        i = 1
        while i <= rows_n:
            writer.writerow(_row(rng, i, base_sum))
            i += 1
    finally:
        f.close()
    size = os.path.getsize(path)
    print(
        "OK %s: %s rows, delim=%r, %d bytes → %s"
        % (tier_name, rows_n, delim, size, path)
    )


def write_readme(out_dir):
    lines = [
        "# CSV-источники для collect_workbooks",
        "",
        "Сгенерировано: `python test/collect_workbooks/generate_csv_sources.py`",
        "",
        "Кодировка файлов: **Windows-1251** (как default в `Доп_Параметры_Источника`).",
        "",
        "| Файл | Строк данных | Разделитель | Кодировка | Назначение |",
        "|------|-------------:|:-----------:|:---------:|------------|",
    ]
    for name, cfg in _VOLUME.items():
        fname = "source_csv_%s.csv" % name
        lines.append(
            "| `%s` | %d | `%s` | `%s` | объём **%s** |"
            % (fname, cfg["rows"], cfg["delimiter"], cfg["encoding"], name)
        )
    lines.extend(
        [
            "",
            "## Параметры на листе",
            "",
            "```",
            "Файлы-Источники | …/sources/csv/source_csv_tiny.csv | …/source_csv_small.csv | …",
            'Доп_Параметры_Источника | {"delimiter":";","encoding":"windows-1251"} | {"delimiter":",","encoding":"windows-1251"} | …',
            "```",
            "",
            "Пустая `Доп_Параметры_Источника` → разделитель `;`, кодировка `windows-1251`.",
            "Для UTF-8-файлов укажите `\"encoding\":\"utf-8\"` (или `utf-8-sig`).",
            "",
            "Колонки: " + ", ".join(HEADERS) + ".",
            "",
        ]
    )
    path = os.path.join(out_dir, "README.md")
    with io.open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("README → %s" % path)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Generate CSV sources (tiny…xlarge)")
    parser.add_argument(
        "--tier",
        choices=list(_VOLUME.keys()) + ["all"],
        default="all",
        help="Какой объём сгенерировать (по умолчанию all)",
    )
    parser.add_argument(
        "--out-dir",
        default=DEFAULT_OUT_DIR,
        help="Каталог вывода (по умолчанию sources/csv)",
    )
    args = parser.parse_args(argv)
    out_dir = os.path.abspath(args.out_dir)
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    tiers = list(_VOLUME.keys()) if args.tier == "all" else [args.tier]
    for name in tiers:
        path = os.path.join(out_dir, "source_csv_%s.csv" % name)
        write_csv(path, name, _VOLUME[name])
    write_readme(out_dir)
    return 0


if __name__ == "__main__":
    sys.exit(main())
