#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор XLSX-источников для ручного теста ВПР (collect_workbooks).

Два типа соответствия ключей в правой таблице:
  1to1 — один ключ → одна строка справочника
  1toN — один ключ → несколько строк (дубли для проверки «Заполнять_дубли» / мульти)

Уровни объёма: low / middle / high (разное число строк и составной ключ).

Запуск:
  python test/collect_workbooks/generate_vlookup_sources.py
  python test/collect_workbooks/generate_vlookup_sources.py --cardinality 1to1 --tier low

Создаёт sources/vlookup/vlookup_{1to1|1toN}_{low|middle|high}.xlsx
"""
from __future__ import print_function, unicode_literals

import argparse
import json
import random
import sys
import time
from datetime import datetime, timedelta

try:
    from openpyxl import Workbook
except ImportError:
    print("Требуется openpyxl: pip install openpyxl", file=sys.stderr)
    sys.exit(1)

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT_DIR = os.path.join(BASE_DIR, "sources", "vlookup")

LEFT_SHEET = "Заказы"
RIGHT_SHEET = "Справочник_цен"

# Базовые параметры объёма (dup задаётся отдельно для 1to1 / 1toN)
_VOLUME = {
    "low": {"rows_left": 180, "rows_right": 140, "key_cols": 1, "seed": 101},
    "middle": {"rows_left": 850, "rows_right": 420, "key_cols": 2, "seed": 202},
    "high": {"rows_left": 3000, "rows_right": 1100, "key_cols": 3, "seed": 303},
}

_DUP = {
    "1to1": {"dup_min": 1, "dup_max": 1},
    "1toN": {"dup_min": 2, "dup_max": 4},
}

VLOOKUP_TIERS = {}
for _card in ("1to1", "1toN"):
    for _vol, _cfg in _VOLUME.items():
        _name = "%s_%s" % (_card, _vol)
        VLOOKUP_TIERS[_name] = dict(_cfg)
        VLOOKUP_TIERS[_name].update(_DUP[_card])
        VLOOKUP_TIERS[_name]["cardinality"] = _card

REGIONS = ("Москва", "СПб", "Казань", "Новосибирск", "Екатеринбург", "Краснодар")
CURRENCIES = ("RUB", "USD", "EUR")
SUPPLIERS = ("Поставщик_A", "Поставщик_B", "Поставщик_C", "Поставщик_D")
WAREHOUSES = ("Склад_01", "Склад_02", "Склад_03", "Склад_04", "Склад_05")
MANAGERS = ("Иванов", "Петров", "Сидоров", "Козлова", "Новикова")

EXTRACT_COLUMNS = ("Цена", "Валюта", "Поставщик")


def left_headers(key_cols):
    keys = ["Регион", "Код_товара", "Склад"][:key_cols]
    tail = [
        "Код_позиции",
        "Количество",
        "Сумма_база",
        "Дата_заказа",
        "Менеджер",
        "Номер_заказа",
        "Статус",
        "Комментарий",
    ]
    if key_cols == 1:
        return ["Код_позиции"] + tail[1:]
    out = list(keys)
    for h in tail:
        if h not in out:
            out.append(h)
    while len(out) < 8:
        out.append("Поле_%d" % (len(out) + 1))
    if key_cols >= 2 and len(out) < 11:
        out.extend(["Резерв_%d" % (len(out) - 10)] * (11 - len(out)))
    if key_cols >= 3 and len(out) < 14:
        out.extend(["Доп_%d" % i for i in range(1, 15 - len(out) + 1)])
    return out


def right_headers(key_cols):
    keys = ["Регион", "Код_товара", "Склад"][:key_cols]
    if key_cols == 1:
        return [
            "Код_позиции",
            "Цена",
            "Валюта",
            "Поставщик",
            "Дата_цены",
            "Срок_дней",
            "Комментарий",
        ]
    base = keys + [
        "Цена",
        "Валюта",
        "Поставщик",
        "Дата_цены",
        "Срок_дней",
        "Коэф_наценки",
        "Комментарий",
    ]
    while len(base) < 10:
        base.insert(-1, "Атрибут_%d" % (len(base) - 6))
    if key_cols >= 3:
        while len(base) < 12:
            base.insert(-1, "Спр_%d" % (len(base) - 7))
    return base


def key_field_names(key_cols):
    if key_cols == 1:
        return ["Код_позиции"]
    return ["Регион", "Код_товара", "Склад"][:key_cols]


def tier_filename(tier):
    return "vlookup_%s.xlsx" % tier


def _make_key_tuple(rng, key_cols, index):
    if key_cols == 1:
        return ("POS-%06d" % (100000 + index),)
    region = REGIONS[index % len(REGIONS)]
    product = "SKU-%05d" % (index % 5000 + 1)
    if key_cols == 2:
        return (region, product)
    warehouse = WAREHOUSES[index % len(WAREHOUSES)]
    return (region, product, warehouse)


def build_right_catalog(rng, key_cols, unique_keys, dup_min, dup_max):
    catalog = []
    ki = 0
    while ki < unique_keys:
        key = _make_key_tuple(rng, key_cols, ki)
        if dup_min == dup_max == 1:
            dup_n = 1
        elif ki % 3 != 2:
            dup_n = rng.randint(dup_min, dup_max)
        else:
            dup_n = 1
        di = 0
        while di < dup_n:
            base_price = 100 + (ki % 97) * 10 + di * 5
            catalog.append(
                {
                    "key": key,
                    "Цена": base_price + rng.randint(0, 50),
                    "Валюта": CURRENCIES[di % len(CURRENCIES)],
                    "Поставщик": SUPPLIERS[(ki + di) % len(SUPPLIERS)],
                    "Дата_цены": datetime(2024, 6, 1) + timedelta(days=ki % 180),
                    "Срок_дней": 7 + (ki % 30),
                    "Коэф_наценки": round(1.0 + (di * 0.05), 2),
                    "Комментарий": "ref/%d/%d" % (ki + 1, di + 1),
                }
            )
            di += 1
        ki += 1
    return catalog


def catalog_to_rows(catalog, key_cols, headers):
    key_names = key_field_names(key_cols)
    rows = []
    for rec in catalog:
        row = []
        for h in headers:
            if h in key_names:
                idx = key_names.index(h)
                row.append(rec["key"][idx])
            elif h in rec:
                row.append(rec[h])
            elif h.startswith(("Атрибут_", "Спр_", "Доп_", "Поле_", "Резерв_")):
                row.append("")
            else:
                row.append("")
        rows.append(row)
    return rows


def build_left_rows(rng, key_cols, headers, row_count, catalog):
    key_names = key_field_names(key_cols)
    catalog_keys = [rec["key"] for rec in catalog]
    unique_catalog_keys = []
    seen = set()
    for k in catalog_keys:
        if k not in seen:
            seen.add(k)
            unique_catalog_keys.append(k)

    rows = []
    i = 0
    while i < row_count:
        if i % 17 == 0:
            key = _make_key_tuple(rng, key_cols, 900000 + i)
        elif i % 5 == 0 and unique_catalog_keys:
            key = unique_catalog_keys[i % len(unique_catalog_keys)]
        else:
            key = catalog_keys[i % len(catalog_keys)]

        row_map = {}
        ki = 0
        while ki < len(key_names):
            row_map[key_names[ki]] = key[ki]
            ki += 1
        if key_cols == 1:
            row_map["Код_позиции"] = key[0]
        row_map["Количество"] = (i % 50) + 1
        row_map["Сумма_база"] = 1000 + (i % 200) * 17
        row_map["Дата_заказа"] = datetime(2024, 3, 1) + timedelta(days=i % 90)
        row_map["Менеджер"] = MANAGERS[i % len(MANAGERS)]
        row_map["Номер_заказа"] = "ORD-%07d" % (i + 1)
        row_map["Статус"] = "Новый" if i % 4 else "В работе"
        row_map["Комментарий"] = "ord/%d" % (i + 1)

        out = []
        for h in headers:
            out.append(row_map.get(h, ""))
        rows.append(out)
        i += 1
    return rows


def write_sheet(ws, headers, rows):
    ws.append(list(headers))
    for row in rows:
        ws.append(row)


def write_tier_file(out_dir, tier, cfg):
    rng = random.Random(cfg.get("seed", 0))
    key_cols = int(cfg["key_cols"])
    lh = left_headers(key_cols)
    rh = right_headers(key_cols)
    dup_min = int(cfg["dup_min"])
    dup_max = int(cfg["dup_max"])

    unique_keys = max(40, int(cfg["rows_right"]) // max(dup_max, 1))
    catalog = build_right_catalog(rng, key_cols, unique_keys, dup_min, dup_max)
    right_rows = catalog_to_rows(catalog, key_cols, rh)
    target_right = int(cfg["rows_right"])
    if len(right_rows) > target_right:
        right_rows = right_rows[:target_right]
    while len(right_rows) < target_right:
        src = dict(catalog[len(right_rows) % len(catalog)])
        src["Цена"] = int(src["Цена"]) + len(right_rows)
        src["Комментарий"] = "ref/extra/%d" % (len(right_rows) + 1)
        right_rows.append(catalog_to_rows([src], key_cols, rh)[0])

    left_rows = build_left_rows(rng, key_cols, lh, int(cfg["rows_left"]), catalog)

    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, tier_filename(tier))
    t0 = time.time()

    wb = Workbook(write_only=True)
    ws_left = wb.create_sheet(LEFT_SHEET)
    write_sheet(ws_left, lh, left_rows)
    ws_right = wb.create_sheet(RIGHT_SHEET)
    write_sheet(ws_right, rh, right_rows)
    wb.save(path)

    key_counts = {}
    for rec in catalog:
        k = rec["key"]
        key_counts[k] = key_counts.get(k, 0) + 1
    multi_keys = sum(1 for c in key_counts.values() if c > 1)
    print(
        "  %s — %s, левый %d×%d, правый %d×%d, ключ [%s], ключей с дублями: %d, %.1f с"
        % (
            os.path.basename(path),
            cfg.get("cardinality", "?"),
            len(left_rows),
            len(lh),
            len(right_rows),
            len(rh),
            ",".join(key_field_names(key_cols)),
            multi_keys,
            time.time() - t0,
        )
    )
    return path


def vlookup_json_for_tier(tier):
    """JSON для колонки C строки ВПР (§5.2 ТЗ)."""
    cfg = VLOOKUP_TIERS[tier]
    keys = key_field_names(int(cfg["key_cols"]))
    fill_dup = cfg.get("cardinality") == "1toN"
    block = {
        "v": 1,
        "fn": "впр",
        "left": {
            "sheet": LEFT_SHEET,
            "start_row": 2,
            "start_col": 1,
            "header_row": 1,
        },
        "right": {
            "sheet": RIGHT_SHEET,
            "start_row": 2,
            "start_col": 1,
            "header_row": 1,
        },
        "keys_same": True,
        "keys_left": keys,
        "keys_right": keys,
        "extract_columns": list(EXTRACT_COLUMNS),
        "highlight_color": "#FFF2CC",
        "fill_duplicates": fill_dup,
        "not_found_fill": "#Н/Д",
        "multi_match_one_cell": False,
        "trim_keys": False,
        "column_suffix": True,
    }
    return json.dumps([block], ensure_ascii=False, separators=(",", ":"))


def write_readme(out_dir):
    lines = [
        "# Источники для теста ВПР",
        "",
        "Файлы `vlookup_{1to1|1toN}_{low|middle|high}.xlsx` — две вкладки:",
        "",
        "- **Заказы** — левая таблица (без цен из справочника)",
        "- **Справочник_цен** — правая таблица",
        "",
        "Строка 1 — заголовки, данные с строки 2.",
        "",
        "## Типы соответствия",
        "",
        "| Префикс | Правый лист | Зачем |",
        "|---------|-------------|--------|",
        "| **1to1** | один ключ → одна строка | базовый ВПР, без дублей |",
        "| **1toN** | один ключ → 2–4 строки | дубли, «Заполнять_дубли», мульти |",
        "",
        "## Объёмы",
        "",
        "| Уровень | Строк (левый) | Строк (правый) | Ключ |",
        "|---------|---------------|----------------|------|",
        "| low | ~180 | ~140 | Код_позиции |",
        "| middle | ~850 | ~420 | Регион, Код_товара |",
        "| high | ~3000 | ~1100 | Регион, Код_товара, Склад |",
        "",
        "Сценарии: `workbooks/11_vlookup_{tier}.xltx` — режим «Копирование листов»,",
        "ВПР (JSON в C) между шагами постобработки.",
        "",
        "Генерация:",
        "",
        "```bash",
        "python test/collect_workbooks/generate_vlookup_sources.py",
        "./test/collect_workbooks/generate.sh",
        "```",
    ]
    readme = os.path.join(out_dir, "README.md")
    with open(readme, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("  readme:", readme)


def generate_all_tiers(out_dir=None, tier=None, seed=None, cardinality=None):
    out = out_dir or DEFAULT_OUT_DIR
    tiers = sorted(VLOOKUP_TIERS.keys())
    if tier:
        tiers = [t for t in tiers if t.endswith("_%s" % tier) or t == tier]
    if cardinality:
        tiers = [t for t in tiers if t.startswith(cardinality + "_")]
    paths = []
    for name in tiers:
        cfg = dict(VLOOKUP_TIERS[name])
        if seed is not None:
            cfg["seed"] = seed
        paths.append(write_tier_file(out, name, cfg))
    write_readme(out)
    return paths


def main(argv=None):
    ap = argparse.ArgumentParser(description="Генератор XLSX для теста ВПР")
    ap.add_argument(
        "--tier",
        choices=sorted(_VOLUME.keys()),
        default=None,
        help="уровень объёма (low/middle/high); по умолчанию — все",
    )
    ap.add_argument(
        "--cardinality",
        choices=("1to1", "1toN"),
        default=None,
        help="тип соответствия ключей справа",
    )
    ap.add_argument("--seed", type=int, default=None)
    ap.add_argument("--out-dir", default=DEFAULT_OUT_DIR)
    args = ap.parse_args(argv)
    print("Генерация источников ВПР →", os.path.abspath(args.out_dir))
    generate_all_tiers(args.out_dir, args.tier, args.seed, args.cardinality)
    print("Готово.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
