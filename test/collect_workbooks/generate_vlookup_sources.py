#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор XLSX-источников для ручного теста ВПР (collect_workbooks).

Два типа соответствия ключей в правой таблице:
  1to1 — один ключ → одна строка справочника
  1toN — один ключ → несколько строк (дубли для проверки «Заполнять_дубли» / мульти)

Уровни объёма: low / middle / high (разное число строк и составной ключ).

Заказы: несколько неповторяющихся товаров в одном Номер_заказа;
Количество — случайное, в т.ч. дробное; товары повторяются между заказами.
Справочник_цен: ~450 позиций ассортимента + колонка Наименование.

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

# Базовый ассортимент товаров (уникальные SKU / POS).
ASSORTMENT_SIZE = 450

# Базовые параметры объёма (dup задаётся отдельно для 1to1 / 1toN)
_VOLUME = {
    "low": {"rows_left": 180, "rows_right": 450, "key_cols": 1, "seed": 101, "assortment": 450},
    "middle": {"rows_left": 850, "rows_right": 520, "key_cols": 2, "seed": 202, "assortment": 450},
    "high": {"rows_left": 3000, "rows_right": 1200, "key_cols": 3, "seed": 303, "assortment": 450},
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

PRODUCT_KINDS = (
    "Кабель",
    "Разъём",
    "Модуль",
    "Датчик",
    "Блок питания",
    "Контроллер",
    "Крепёж",
    "Корпус",
    "Фильтр",
    "Переходник",
    "Антенна",
    "Реле",
)
PRODUCT_ATTRS = (
    "медный",
    "оптоволоконный",
    "экранированный",
    "промышленный",
    "бытовой",
    "влагозащитный",
    "миниатюрный",
    "усиленный",
)

# Колонки извлечения для JSON ВПР в сценариях (после ключей справа).
EXTRACT_COLUMNS = ("Наименование", "Цена", "Валюта")


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
            "Наименование",
            "Цена",
            "Валюта",
            "Поставщик",
            "Дата_цены",
            "Срок_дней",
            "Комментарий",
        ]
    base = keys + [
        "Наименование",
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


def product_sku(index):
    """Уникальный код товара в ассортименте (1-based)."""
    return "SKU-%05d" % (int(index) + 1)


def product_pos(index):
    """Код позиции для одноколоночного ключа."""
    return "POS-%06d" % (100000 + int(index))


def product_name(index):
    """Человекочитаемое наименование товара."""
    i = int(index)
    kind = PRODUCT_KINDS[i % len(PRODUCT_KINDS)]
    attr = PRODUCT_ATTRS[(i // len(PRODUCT_KINDS)) % len(PRODUCT_ATTRS)]
    return "%s %s %s" % (kind, attr, product_sku(i))


def _make_key_tuple(key_cols, assort_index, region_offset=0):
    """Ключ справочника/заказа по индексу ассортимента 0..assortment-1."""
    i = int(assort_index)
    if key_cols == 1:
        return (product_pos(i),)
    region = REGIONS[(i + int(region_offset)) % len(REGIONS)]
    product = product_sku(i)
    if key_cols == 2:
        return (region, product)
    warehouse = WAREHOUSES[i % len(WAREHOUSES)]
    return (region, product, warehouse)


def _random_qty(rng):
    """Случайное количество: целое или дробное (1 знак)."""
    if rng.random() < 0.45:
        return rng.randint(1, 120)
    # дробное: 0.1 … 99.9, не ноль
    q = round(rng.uniform(0.1, 99.9), 1)
    if q == int(q):
        q = q + 0.5
    return q


def build_right_catalog(rng, key_cols, unique_keys, dup_min, dup_max):
    catalog = []
    ki = 0
    while ki < unique_keys:
        key = _make_key_tuple(key_cols, ki)
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
                    "Наименование": product_name(ki),
                    "Цена": base_price + rng.randint(0, 50),
                    "Валюта": CURRENCIES[di % len(CURRENCIES)],
                    "Поставщик": SUPPLIERS[(ki + di) % len(SUPPLIERS)],
                    "Дата_цены": datetime(2024, 6, 1) + timedelta(days=ki % 180),
                    "Срок_дней": 7 + (ki % 30),
                    "Коэф_наценки": round(1.0 + (di * 0.05), 2),
                    "Комментарий": "ref/%d/%d" % (ki + 1, di + 1),
                    "assort_index": ki,
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


def _unique_catalog_entries(catalog):
    """Один представитель на уникальный ключ (для выбора товаров в заказ)."""
    unique = []
    seen = set()
    for rec in catalog:
        k = rec["key"]
        if k in seen:
            continue
        seen.add(k)
        unique.append(rec)
    return unique


def build_left_rows(rng, key_cols, headers, row_count, catalog):
    """
    Строки заказов: несколько неповторяющихся товаров в одном Номер_заказа.
    Количество — вразнобой и дробное; товары повторяются между заказами.
    Часть строк — ключи вне справочника (для #Н/Д).
    """
    key_names = key_field_names(key_cols)
    unique_recs = _unique_catalog_entries(catalog)
    if not unique_recs:
        return []

    rows = []
    order_no = 1
    line_i = 0
    while len(rows) < row_count:
        # 1–6 позиций в заказе (не больше уникальных товаров в справочнике).
        max_items = min(6, len(unique_recs))
        n_items = rng.randint(1, max_items) if max_items > 0 else 1
        if len(rows) + n_items > row_count:
            n_items = row_count - len(rows)

        # Неповторяющиеся товары внутри заказа.
        picked = rng.sample(unique_recs, n_items)
        order_id = "ORD-%07d" % order_no
        order_comment = "ord/%d" % order_no
        order_date = datetime(2024, 3, 1) + timedelta(days=(order_no - 1) % 90)
        manager = MANAGERS[(order_no - 1) % len(MANAGERS)]
        status = "Новый" if order_no % 4 else "В работе"

        ji = 0
        while ji < len(picked):
            rec = picked[ji]
            # ~6% строк — ключ вне справочника (для проверки «не найдено»).
            if line_i % 17 == 0:
                key = _make_key_tuple(key_cols, ASSORTMENT_SIZE + line_i, region_offset=line_i)
            else:
                key = rec["key"]

            row_map = {}
            ki = 0
            while ki < len(key_names):
                row_map[key_names[ki]] = key[ki]
                ki += 1
            if key_cols == 1:
                row_map["Код_позиции"] = key[0]
            elif "Код_позиции" in headers:
                # Для составного ключа — код товара (удобно смотреть глазами).
                row_map["Код_позиции"] = key[1] if len(key) > 1 else ""

            row_map["Количество"] = _random_qty(rng)
            row_map["Сумма_база"] = round(1000 + rng.uniform(0, 5000), 2)
            row_map["Дата_заказа"] = order_date
            row_map["Менеджер"] = manager
            row_map["Номер_заказа"] = order_id
            row_map["Статус"] = status
            row_map["Комментарий"] = order_comment

            out = []
            for h in headers:
                out.append(row_map.get(h, ""))
            rows.append(out)
            line_i += 1
            ji += 1

        order_no += 1
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

    unique_keys = int(cfg.get("assortment", ASSORTMENT_SIZE))
    if unique_keys < 1:
        unique_keys = ASSORTMENT_SIZE
    catalog = build_right_catalog(rng, key_cols, unique_keys, dup_min, dup_max)
    kept_unique = []
    seen_keys = set()
    for rec in catalog:
        k = rec["key"]
        if k in seen_keys:
            continue
        seen_keys.add(k)
        kept_unique.append(rec)

    if dup_max <= 1:
        # 1to1: ровно ассортимент, без дублей ключей.
        right_rows = catalog_to_rows(kept_unique, key_cols, rh)
    else:
        # 1toN: полный каталог с дублями (ассортимент × 2–4).
        right_rows = catalog_to_rows(catalog, key_cols, rh)
        target_right = max(int(cfg["rows_right"]), len(right_rows))
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

    # Статистика заказов слева
    order_col = None
    if "Номер_заказа" in lh:
        order_col = lh.index("Номер_заказа")
    order_sizes = {}
    if order_col is not None:
        for row in left_rows:
            oid = row[order_col]
            order_sizes[oid] = order_sizes.get(oid, 0) + 1
    multi_orders = sum(1 for c in order_sizes.values() if c > 1)

    print(
        "  %s — %s, левый %d×%d, правый %d×%d, ключ [%s], ассортимент=%d, "
        "ключей с дублями: %d, заказов с >1 позицией: %d, %.1f с"
        % (
            os.path.basename(path),
            cfg.get("cardinality", "?"),
            len(left_rows),
            len(lh),
            len(right_rows),
            len(rh),
            ",".join(key_field_names(key_cols)),
            unique_keys,
            multi_keys,
            multi_orders,
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
        "- **Справочник_цен** — правая таблица (есть **Наименование**)",
        "",
        "Строка 1 — заголовки, данные с строки 2.",
        "",
        "## Содержимое Заказов",
        "",
        "- Несколько **неповторяющихся** товаров в одном `Номер_заказа`",
        "  (тот же `Комментарий` = `ord/N` на всех строках заказа).",
        "- `Количество` — случайное, в том числе **дробное**.",
        "- Товары **повторяются между заказами**; не весь ассортимент обязан",
        "  попасть в заказы.",
        "",
        "## Справочник",
        "",
        "- Ассортимент ≈ **450** позиций (`SKU-*****` / `POS-******`) + наименование.",
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
        "| low | ~180 | ≥450 (1to1) / с дублями | Код_позиции |",
        "| middle | ~850 | ≥450 + запас | Регион, Код_товара |",
        "| high | ~3000 | ≥450 + запас | Регион, Код_товара, Склад |",
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
