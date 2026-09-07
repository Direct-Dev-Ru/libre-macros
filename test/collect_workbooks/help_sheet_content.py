# -*- coding: utf-8 -*-
"""Содержимое листа «Справка_макроса» для тестовых шаблонов."""
from __future__ import print_function, unicode_literals

import ast
import os
import re

HELP_SHEET_NAME = "Справка_макроса"

_COLLECT_PY = "collect_workbooks.py"


def _macro_py_path(root):
    return os.path.join(root, "macro-lib", _COLLECT_PY)


def load_macro_help_lines(root):
    path = _macro_py_path(root)
    if not os.path.isfile(path):
        return ["(не найден файл макроса: %s)" % path]
    with open(path, "r", encoding="utf-8") as f:
        source = f.read()
    doc = ast.get_docstring(ast.parse(source))
    if doc is None or doc.strip() == "":
        return ["(в макросе нет описания)"]
    return doc.splitlines()


def _read_collect_source(root):
    path = _macro_py_path(root)
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def _parse_map_spec_keys(text, spec_name):
    """Ключи B из кортежа _PP_RANGE_MAP_SPEC = (("key", "merge_…"), …)."""
    m = re.search(
        r"%s\s*=\s*\((.*?)\n\)\s*\n" % re.escape(spec_name),
        text,
        re.DOTALL,
    )
    if not m:
        return []
    body = m.group(1)
    return re.findall(r'\(\s*"([^"]+)"\s*,', body)


def _parse_rest_hints_dict(text, dict_name):
    """
    Пары (описание, пример) из MERGE_*_REST_HINTS без import collect_workbooks.
    Пропускает алиасы вида MERGE_*_REST_HINTS["…"].
    """
    m = re.search(
        r"%s\s*=\s*\{" % re.escape(dict_name),
        text,
    )
    if not m:
        return {}
    start = m.end()
    depth = 1
    i = start
    while i < len(text) and depth > 0:
        ch = text[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
        i += 1
    block = text[start : i - 1]
    hints = {}
    for key_m in re.finditer(r'"([^"]+)":\s*\(', block):
        key = key_m.group(1)
        pos = key_m.end()
        if block[pos : pos + 20].strip().startswith("MERGE_"):
            continue
        depth_p = 1
        j = pos + 1
        while j < len(block) and depth_p > 0:
            if block[j] == "(":
                depth_p += 1
            elif block[j] == ")":
                depth_p -= 1
            j += 1
        inner = block[pos:j]
        parts = re.findall(r'"((?:[^"\\]|\\.)*)"', inner, re.DOTALL)
        if len(parts) >= 1:
            desc = parts[0].replace("\\n", "\n")
            example = parts[1].replace("\\n", "\n") if len(parts) > 1 else ""
            hints[key] = (desc, example)
    return hints


def _postprocess_rest_hint_lines(name, hints_map):
    info = hints_map.get(name)
    if info is None:
        return ["    C: (подсказка не задана)"]
    desc, example = info
    lines = ["    C: %s" % desc]
    if example != "":
        lines.append("    Пример C: %s" % example)
    return lines


def _load_pp_hints(root):
    text = _read_collect_source(root)
    range_keys = _parse_map_spec_keys(text, "_PP_RANGE_MAP_SPEC")
    row_keys = _parse_map_spec_keys(text, "_PP_ROW_MAP_SPEC")
    final_keys = _parse_map_spec_keys(text, "_FINAL_MAP_SPEC")
    range_hints = _parse_rest_hints_dict(text, "MERGE_POSTPROCESS_RANGE_REST_HINTS")
    row_hints = _parse_rest_hints_dict(text, "MERGE_POSTPROCESS_ROW_REST_HINTS")
    final_hints = _parse_rest_hints_dict(text, "MERGE_FINAL_PROCESSING_REST_HINTS")
    return range_hints, row_hints, final_hints, range_keys, row_keys, final_keys


def build_help_scenario_lines(scenarios):
    """scenarios: [(filename, comment), ...]"""
    lines = []
    lines.append("")
    lines.append("Тестовые сценарии collect_workbooks")
    lines.append("=" * 40)
    lines.append("")
    for fn, comment in scenarios:
        lines.append("%s — %s" % (fn, comment))
    lines.append("")
    lines.append("Подробнее: docs/12_MANUAL_TESTING.md, docs/13_JSON_PARAMS.md")
    lines.append("Генератор: test/collect_workbooks/generate.sh")
    return lines


# Пример для справки (отдельный A-параметр объединить_листы_в_один)
PRESET_MERGE_EXAMPLE = (
    '[{"v":1,"fn":"объединить_листы_в_один","sheets":["Отчет_Q1","Отчет_Q2","Отчет_Q3"],'
    '"dest_sheet":"Отчет_Q123","header_row":1,"with_formatting":false}]'
)


def build_docs_pointer_lines():
    return [
        "",
        "Документация проекта (docs/)",
        "=" * 40,
        "00_INDEX.md — оглавление",
        "04_DATA_COLLECTION.md — параметры сбора",
        "05_POSTPROCESS.md — постобработка RANGE/ROW",
        "06_PARAM_WIZARD.md — визард параметров",
        "13_JSON_PARAMS.md — JSON в колонке C",
        "15_PARAM_REFERENCE.md — реестр параметров A (в т.ч. объединить_листы_в_один)",
        "",
    ]


def build_pipeline_params_lines():
    """Отдельные строки A pipeline (не ключи B Постобработка_*)."""
    return [
        "Отдельные параметры pipeline (колонка A)",
        "=" * 40,
        "",
        "Как «ВПР»: имя в A, JSON в C. Не выбирать как функцию Постобработка_Диапазон.",
        "",
        "• объединить_листы_в_один — A=объединить_листы_в_один; C=JSON",
        "    sheets, dest_sheet, header_row (1-based на листах результата), with_formatting",
        "    Пример C: %s" % PRESET_MERGE_EXAMPLE,
        "• разделить_листы — A=разделить_листы; C=JSON (заглушка)",
        "• ВПР — A=ВПР; B–F или JSON в C",
        "",
        "Порядок: блоки range ↔ merge/split/vlookup ↔ range; row только после всех range/splitter.",
        "",
    ]


def build_postprocess_catalog_lines(root):
    (
        range_hints,
        row_hints,
        final_hints,
        range_keys,
        row_keys,
        final_keys,
    ) = _load_pp_hints(root)
    # Primary UX: объединить — отдельный A-параметр; в FINAL map остаётся legacy.
    final_keys_show = [
        k for k in final_keys if k.casefold() != "объединить_листы_в_один".casefold()
    ]
    lines = []
    lines.extend(build_pipeline_params_lines())
    lines.append("Постобработка и финальная обработка")
    lines.append("=" * 40)
    lines.append("")
    lines.append("Колонка C — только JSON (массив блоков). Справочник: docs/13_JSON_PARAMS.md")
    lines.append("")
    lines.append("--- Диапазон (RANGE), %d функций ---" % len(range_keys))
    for name in range_keys:
        lines.append("• %s" % name)
        lines.extend(_postprocess_rest_hint_lines(name, range_hints))
    lines.append("")
    lines.append("--- Строка (ROW), %d функций ---" % len(row_keys))
    for name in row_keys:
        lines.append("• %s" % name)
        lines.extend(_postprocess_rest_hint_lines(name, row_hints))
    lines.append("")
    lines.append(
        "--- Финальная обработка, %d функций (без legacy «объединить_листы_в_один») ---"
        % len(final_keys_show)
    )
    for name in final_keys_show:
        lines.append("• %s" % name)
        lines.extend(_postprocess_rest_hint_lines(name, final_hints))
    if len(final_keys_show) < len(final_keys):
        lines.append("")
        lines.append(
            "Примечание: ключ «объединить_листы_в_один» в финальной карте — legacy; "
            "рекомендуется отдельная строка A (см. выше)."
        )
    return lines


def collect_help_sheet_content(root, scenarios, macro_version):
    lines = []
    lines.extend(load_macro_help_lines(root))
    lines.extend(build_docs_pointer_lines())
    lines.extend(build_help_scenario_lines(scenarios))
    lines.append("")
    lines.extend(build_postprocess_catalog_lines(root))
    lines.append("")
    lines.append("Версия макроса при генерации шаблона: %s" % macro_version)
    return lines


def help_sheet_title(macro_version):
    return "collect_workbooks — справка (версия %s)" % macro_version
