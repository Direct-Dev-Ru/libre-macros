#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор источников и книг сценариев для ручной проверки collect_workbooks в Calc.

Колонка C для постобработки — только JSON (см. docs/13_JSON_PARAMS.md).

Использование:
  python test/collect_workbooks/generate_manual_workbooks.py
  python test/collect_workbooks/generate_manual_workbooks.py --odf   # + workbooks/ODF/*.ots
"""
from __future__ import print_function, unicode_literals

import argparse
import datetime
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
TEST_DIR = os.path.join(ROOT, "test", "collect_workbooks")
SOURCES = os.path.join(TEST_DIR, "sources")
WORKBOOKS = os.path.join(TEST_DIR, "workbooks")
ODF_DIR = os.path.join(WORKBOOKS, "ODF")
START_TEMPLATE = os.path.join(TEST_DIR, "start_param_file.xlsx")
MANUAL_CHECK = os.path.join(WORKBOOKS, "MANUAL_CHECK.md")
INSTALLER_DIR = os.path.join(ROOT, "installer")

_MACRO_PYTHONPATH = os.path.join(ROOT, "macro-lib", "pythonpath")
if _MACRO_PYTHONPATH not in sys.path:
    sys.path.insert(0, _MACRO_PYTHONPATH)

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
except ImportError:
    try:
        import openpyxl_bundled  # noqa: F401 — bundled для dev без pip
        import openpyxl
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    except ImportError:
        print("Требуется openpyxl: pip install openpyxl", file=sys.stderr)
        sys.exit(1)

sys.path.insert(0, TEST_DIR)
sys.path.insert(0, INSTALLER_DIR)
from manual_json_presets import PRESET, RANGE_MATRIX_ORDER, ROW_MATRIX_ORDER, SHEET_FN_ALIASES
from generate_vlookup_sources import (
    VLOOKUP_TIERS,
    tier_filename,
    vlookup_json_for_tier,
    generate_all_tiers as generate_vlookup_sources,
)
from generate_form_sources import (
    FORMS_XLSX,
    generate_form_sources,
)
from generate_transpose_sources import (
    OUT_XLSX as TRANSPOSE_XLSX,
    generate_transpose_sources,
)
from help_sheet_content import (
    HELP_SHEET_NAME,
    collect_help_sheet_content,
    help_sheet_title,
)
from xlsx_help_sheet import append_help_sheet
from macro_version import read_canonical_macro_version

P_RANGE = "Постобработка_Диапазон"
P_ROW = "Постобработка_Строка"
P_CTC = "Ячейка_В_Столбец"
P_VLOOKUP = "ВПР"
P_FILES = "Файлы-Источники"

ODF_TEMPLATE_MIME = "application/vnd.oasis.opendocument.spreadsheet-template"
ODF_SPREADSHEET_MIME = "application/vnd.oasis.opendocument.spreadsheet"


def macro_version():
    return read_canonical_macro_version(ROOT)


def abs_source(name):
    return os.path.join(SOURCES, name)


def abs_large_source(name):
    return os.path.join(SOURCES, "large", name)


def abs_vlookup_source(tier):
    return os.path.join(SOURCES, "vlookup", tier_filename(tier))


def abs_form_source():
    return FORMS_XLSX


def abs_transpose_source():
    return TRANSPOSE_XLSX


def vlookup_row(tier):
    """Строка ВПР: JSON в колонке C (B–F пустые)."""
    return (P_VLOOKUP, ("", vlookup_json_for_tier(tier), "", "", ""))


PIVOT_ROWS_PER_SHEET = 800
PIVOT_EMPLOYEE_CATALOG_SIZE = 120

# Сверхбольшой источник: журналы серверов (ручные тесты производительности / ВПР).
LARGE_SERVER_LOG_ROWS = 200_001
LARGE_SERVER_LOG_BATCH = 10_000
LARGE_SERVER_CATALOG_SIZE = 32


def _fmt_text_number_ru(value, decimals=2):
    """Число как текст с пробелами-разделителями и запятой (для тестов convert)."""
    try:
        num = float(value)
    except (TypeError, ValueError):
        return str(value or "")
    sign = "-" if num < 0 else ""
    num = abs(num)
    if decimals <= 0:
        whole = int(round(num))
        s = str(whole)
    else:
        whole = int(num)
        frac = int(round((num - whole) * (10 ** decimals)))
        if frac >= 10 ** decimals:
            whole += 1
            frac = 0
        s = str(whole)
    parts = []
    while len(s) > 3:
        parts.insert(0, s[-3:])
        s = s[:-3]
    parts.insert(0, s)
    grouped = " ".join(parts)
    if decimals <= 0:
        return sign + grouped
    return "%s%s,%s" % (sign, grouped, str(frac).zfill(decimals))


def _build_pivot_employee_catalog(size=PIVOT_EMPLOYEE_CATALOG_SIZE):
    """Справочник сотрудников для source_pivot.xlsx."""
    surnames = (
        "Иванов", "Петров", "Сидоров", "Козлов", "Новиков", "Морозов", "Волков",
        "Соколов", "Лебедев", "Кузнецов", "Попов", "Васильев", "Фёдоров", "Михайлов",
        "Алексеев", "Романов", "Орлов", "Андреев", "Макаров", "Никитин",
    )
    names_m = (
        "Алексей", "Дмитрий", "Сергей", "Андрей", "Иван", "Павел", "Николай", "Максим",
    )
    names_f = (
        "Мария", "Анна", "Елена", "Ольга", "Наталья", "Татьяна", "Ирина", "Светлана",
    )
    patronymics_m = (
        "Иванович", "Петрович", "Сергеевич", "Андреевич", "Дмитриевич", "Николаевич",
    )
    patronymics_f = (
        "Ивановна", "Петровна", "Сергеевна", "Андреевна", "Дмитриевна", "Николаевна",
    )
    positions = (
        "Менеджер", "Инженер", "Аналитик", "Бухгалтер", "Кладовщик", "Экономист",
        "Специалист", "Руководитель группы", "Разработчик", "Оператор",
    )
    catalog = []
    for i in range(1, size + 1):
        code = "Сотр_%04d" % i
        female = i % 3 == 0
        surname = surnames[(i - 1) % len(surnames)]
        if female:
            name = names_f[(i - 1) % len(names_f)]
            patronymic = patronymics_f[(i - 1) % len(patronymics_f)]
        else:
            name = names_m[(i - 1) % len(names_m)]
            patronymic = patronymics_m[(i - 1) % len(patronymics_m)]
        birth_year = 1975 + (i % 25)
        birth_month = (i % 12) + 1
        birth_day = (i % 27) + 1
        hire_year = 2010 + (i % 14)
        hire_month = ((i * 2) % 12) + 1
        hire_day = min((i % 28) + 1, 28)
        birth = datetime.date(birth_year, birth_month, birth_day)
        hire = datetime.date(hire_year, hire_month, hire_day)
        catalog.append(
            {
                "code": code,
                "row": [
                    code,
                    surname,
                    name,
                    patronymic,
                    positions[(i - 1) % len(positions)],
                    birth.strftime("%d.%m.%Y"),
                    hire.strftime("%d.%m.%Y"),
                ],
            }
        )
    return catalog


def _pivot_employee_code_varied(catalog, row_i, sheet_q):
    """Код в колонке ФИО: один сотрудник встречается несколько раз, длина серий переменная."""
    n = len(catalog)
    streak = 1 + ((row_i * sheet_q * 3) % 4)
    block = row_i // streak
    idx = (block * 17 + row_i * 11 + sheet_q * 29) % n
    if row_i % 23 == 0:
        idx = (row_i // 23 + sheet_q * 5) % n
    return catalog[idx]["code"]


def _build_server_catalog(size=LARGE_SERVER_CATALOG_SIZE):
    """Справочник серверов для ВПР (4-й лист source_server_logs.xlsx)."""
    regions = (
        ("MSK", "Москва", "ЦОД Юг, ул. Дата 12"),
        ("SPB", "Санкт-Петербург", "ЦОД Север, пр. Облачный 3"),
        ("EKB", "Екатеринбург", "ЦОД Урал, ш. Серверное 7"),
        ("NSK", "Новосибирск", "ЦОД Сибирь, ул. Вычислительная 2"),
    )
    roles = (
        ("WEB", "nginx", "Веб-фронт"),
        ("DB", "postgresql", "База данных"),
        ("API", "gateway", "API-шлюз"),
        ("CACHE", "redis", "Кэш"),
        ("MQ", "rabbitmq", "Очередь сообщений"),
        ("MON", "prometheus", "Мониторинг"),
    )
    rows = []
    for i in range(1, size + 1):
        reg = regions[(i - 1) % len(regions)]
        role = roles[(i - 1) % len(roles)]
        code = "SRV-%s-%02d" % (role[0], ((i - 1) % 99) + 1)
        host = "%s-%02d.%s.local" % (role[1], i, reg[0].lower())
        rows.append(
            {
                "code": code,
                "row": [
                    code,
                    host,
                    reg[1],
                    reg[2],
                    reg[0],
                    role[2],
                    role[1],
                ],
            }
        )
    return rows


def _large_server_log_configs(catalog):
    """Три листа журналов: по одному «основному» серверу на лист."""
    picks = [
        ("Лог_web_01", catalog[0]["code"], "nginx", "HTTP"),
        ("Лог_db_01", catalog[5]["code"], "postgresql", "SQL"),
        ("Лог_api_01", catalog[10]["code"], "gateway", "REST"),
    ]
    return picks


def _large_log_message(proto, service, row_i, level):
    templates = {
        "HTTP": (
            "GET /api/v1/orders/%d status=%s",
            "POST /api/v1/auth latency spike",
            "PUT /api/v1/profile field=email",
            "DELETE /api/v1/session cleanup",
        ),
        "SQL": (
            "slow query orders_scan rows=%d",
            "connection pool acquire wait_ms=%d",
            "vacuum analyze public.events",
            "deadlock detected txn=%d",
        ),
        "REST": (
            "upstream timeout service=billing",
            "rate limit client bucket=%d",
            "json parse error field=amount",
            "circuit breaker open peer=payments",
        ),
    }
    pool = templates.get(proto, templates["HTTP"])
    msg = pool[row_i % len(pool)]
    if "%d" in msg and "%s" in msg:
        msg = msg % ((row_i % 9000) + 100, level)
    elif "%d" in msg:
        msg = msg % ((row_i % 9000) + 100)
    elif "%s" in msg:
        msg = msg % (200 if level != "ERROR" else 500)
    return "%s [%s] %s" % (service, level, msg)


def _append_large_log_sheet(wb, sheet_name, banner, server_id, service, proto, n_rows):
    """Один лист журнала (write_only, пакетная запись)."""
    hdr = [
        "Время",
        "Код_сервера",
        "Уровень",
        "Служба",
        "Сообщение",
        "HTTP_код",
        "Длительность_мс",
        "IP_клиента",
    ]
    levels = ("INFO", "WARN", "ERROR", "DEBUG")
    ws = wb.create_sheet(sheet_name)
    ws.append([banner] + [None] * (len(hdr) - 1))
    ws.append(hdr)
    base_ts = datetime.datetime(2025, 6, 1, 0, 0, 0)
    batch = []
    ri = 1
    while ri <= n_rows:
        level = levels[ri % len(levels)]
        http_code = [200, 201, 204, 301, 400, 404, 500, 502][ri % 8]
        duration = 5 + (ri % 4000)
        ip = "10.%d.%d.%d" % ((ri % 200) + 1, (ri * 3) % 250, (ri * 7) % 250)
        ts = base_ts + datetime.timedelta(seconds=ri)
        batch.append(
            [
                ts.strftime("%Y-%m-%d %H:%M:%S"),
                server_id,
                level,
                service,
                _large_log_message(proto, service, ri, level),
                http_code,
                duration,
                ip,
            ]
        )
        if len(batch) >= LARGE_SERVER_LOG_BATCH:
            for row in batch:
                ws.append(row)
            batch = []
            if ri % 50_000 == 0:
                print("    %s: %d / %d строк" % (sheet_name, ri, n_rows))
        ri += 1
    for row in batch:
        ws.append(row)
    print("    %s: готово (%d строк данных)" % (sheet_name, n_rows))


def generate_large_server_sources():
    """
    sources/large/source_server_logs.xlsx:
      3 листа × LARGE_SERVER_LOG_ROWS строк (журналы серверов);
      лист «Справочник_серверов» — код, хост, площадка (для ВПР и пр.).
    """
    os.makedirs(os.path.join(SOURCES, "large"), exist_ok=True)
    path = abs_large_source("source_server_logs.xlsx")
    catalog = _build_server_catalog()
    log_hdr = [
        "Код_сервера",
        "Имя_хоста",
        "Площадка",
        "Адрес_площадки",
        "Регион",
        "Назначение",
        "Служба",
    ]
    print(
        "Большой источник %s: %d листов × %d строк + справочник (%d серверов)…"
        % (
            os.path.basename(path),
            3,
            LARGE_SERVER_LOG_ROWS,
            len(catalog),
        )
    )
    t0 = datetime.datetime.now()
    wb = Workbook(write_only=True)
    for sheet_name, srv_id, service, proto in _large_server_log_configs(catalog):
        _append_large_log_sheet(
            wb,
            sheet_name,
            "Журнал %s / %s" % (srv_id, service),
            srv_id,
            service,
            proto,
            LARGE_SERVER_LOG_ROWS,
        )
    ws_cat = wb.create_sheet("Справочник_серверов")
    ws_cat.append(["Справочник серверов и площадок"])
    ws_cat.append(log_hdr)
    for entry in catalog:
        ws_cat.append(entry["row"])
    wb.save(path)
    elapsed = (datetime.datetime.now() - t0).total_seconds()
    print(
        "  %s (%.1f с, ~%d строк данных на 3 листах)"
        % (path, elapsed, LARGE_SERVER_LOG_ROWS * 3)
    )


def scenario_vlookup(tier):
    """Копирование листов + постобработка + ВПР + постобработка."""
    cfg = VLOOKUP_TIERS[tier]
    card = cfg.get("cardinality", "?")
    keys = ",".join(
        ["Код_позиции"]
        if int(cfg["key_cols"]) == 1
        else ["Регион", "Код_товара", "Склад"][: int(cfg["key_cols"])]
    )
    comment = (
        "ВПР %s / %s: ~%d строк заказов, ключ [%s]. "
        "Режим «Копирование листов», ВПР между постобработкой диапазона."
        % (card, tier.split("_", 1)[-1], int(cfg["rows_left"]), keys)
    )
    rows = [
        ("Файлы-Источники", (abs_vlookup_source(tier),)),
        ("Листы", ("Все",)),
    ] + block_rows(start_row="2", header_row="1") + [
        ("Режим", ("Копирование листов",)),
        pp_range("тонкая_сетка"),
        pp_range("авто_ширина"),
        vlookup_row(tier),
        pp_range("заголовок_плюс_высота"),
        pp_range("зебра", "зебра_диапазон"),
    ]
    return comment, rows


def pp_range(fn, json_key=None):
    key = json_key or SHEET_FN_ALIASES.get(fn, fn)
    return (P_RANGE, fn, PRESET.get(key, ""))


def pp_row(fn, json_key=None):
    key = json_key or SHEET_FN_ALIASES.get(fn, fn)
    return (P_ROW, fn, PRESET.get(key, ""))


def block_rows(start_row="2", start_col="1", row_count="", col_count="", header_row="2"):
    return [
        ("Начало_Строка", (start_row,)),
        ("Начало_Столбец", (start_col,)),
        ("Число_Строк", (row_count,) if row_count else ("",)),
        ("Число_Столбцов", (col_count,) if col_count else ("",)),
        ("Строка_Заголовков", (header_row,)),
    ]


def _param_openpyxl_styles():
    thin = Side(style="thin", color="FFB8B8B8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font = Font(name="PT Sans", size=11)
    font_bold = Font(name="PT Sans", size=11, bold=True)
    fill_header = PatternFill("solid", fgColor="DCE4EE")
    fill_name = PatternFill("solid", fgColor="F0F1F3")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_top = Alignment(horizontal="center", vertical="top", wrap_text=True)
    return {
        "border": border,
        "font": font,
        "font_bold": font_bold,
        "fill_header": fill_header,
        "fill_name": fill_name,
        "align_center": align_center,
        "align_top": align_top,
    }


def format_param_sheet_openpyxl(ws, last_row, value_cols=3):
    """
    Оформление листа параметров как в param_wizard (reformat_param_sheet_table).
    """
    if last_row < 1:
        return
    st = _param_openpyxl_styles()
    max_col = 1 + int(value_cols)
    widths = (35, 22, 22, 22, 22, 22)
    for ci in range(max_col):
        col_letter = openpyxl.utils.get_column_letter(ci + 1)
        ws.column_dimensions[col_letter].width = widths[ci] if ci < len(widths) else 22
    for r in range(1, last_row + 1):
        ws.row_dimensions[r].height = 18 if r == 1 else 16
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            cell.border = st["border"]
            if r == 1:
                cell.font = st["font_bold"]
                cell.fill = st["fill_header"]
                cell.alignment = st["align_center"]
            elif c == 1:
                cell.font = st["font_bold"]
                cell.fill = st["fill_name"]
                cell.alignment = st["align_center"]
            else:
                cell.font = st["font"]
                cell.alignment = st["align_top"]


def _template_has_files_row(ws):
    """В start_param_file.xlsx строка 4 — пустая «Файлы-Источники»."""
    val = ws.cell(4, 1).value
    return val is not None and str(val).strip() == P_FILES


def _write_param_row(ws, row_idx, name, rest, value_cols):
    ws.cell(row_idx, 1).value = name
    if (
        len(rest) == 1
        and isinstance(rest[0], (list, tuple))
        and not isinstance(rest[0], (str, bytes))
    ):
        vals = tuple(rest[0])
    else:
        vals = rest
    for ci in range(value_cols):
        ws.cell(row_idx, 2 + ci).value = vals[ci] if ci < len(vals) else None


def write_workbook(filename, comment, body_rows, value_cols=3):
    """Сохранить .xltx из start_param_file + строки параметров (с 4-й, если там «Файлы-Источники»)."""
    if not os.path.isfile(START_TEMPLATE):
        raise SystemExit("Нет шаблона: %s" % START_TEMPLATE)
    wb = openpyxl.load_workbook(START_TEMPLATE)
    ws = wb.active
    max_row = ws.max_row
    if max_row >= 5:
        ws.delete_rows(5, max_row - 4)
    ws.cell(2, 2).value = comment
    ws.cell(3, 2).value = macro_version()
    headers = ["Параметр"] + ["Значение %d" % (i + 1) for i in range(value_cols)]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c).value = h
    items = list(body_rows)
    row_idx = 5
    if items and _template_has_files_row(ws):
        first = items[0]
        if isinstance(first, (list, tuple)) and len(first) >= 1 and first[0] == P_FILES:
            _write_param_row(ws, 4, first[0], first[1:], value_cols)
            items = items[1:]
    for item in items:
        if isinstance(item, (list, tuple)) and len(item) >= 1:
            name = item[0]
            rest = item[1:] if len(item) > 1 else ()
        else:
            continue
        _write_param_row(ws, row_idx, name, rest, value_cols)
        row_idx += 1
    last_data_row = row_idx - 1
    format_param_sheet_openpyxl(ws, last_data_row, value_cols=value_cols)
    for _ in range(3):
        row_idx += 1
    out = os.path.join(WORKBOOKS, filename)
    os.makedirs(WORKBOOKS, exist_ok=True)
    if filename.lower().endswith((".xltx", ".xltm")):
        wb.template = True
    wb.save(out)
    append_help_sheet(
        out,
        HELP_SHEET_NAME,
        help_sheet_title(macro_version()),
        collect_help_sheet_content(
            ROOT,
            [(fn, comment) for fn, comment, _, _ in SCENARIOS],
            macro_version(),
        ),
    )
    return out


def generate_sources():
    """Минимальные источники source_01..03 и source_products (если пересоздаём)."""
    os.makedirs(SOURCES, exist_ok=True)

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    def _convert(src_path, out_ext, filter_name):
        """
        Конвертация через headless LibreOffice.
        out_ext: "xls" | "ods" | ...
        filter_name: например "xls" или "ods"
        """
        if not soffice:
            return None
        outdir = os.path.dirname(src_path)
        cmd = [
            soffice,
            "--headless",
            "--convert-to",
            "%s:%s" % (out_ext, filter_name) if filter_name else out_ext,
            "--outdir",
            outdir,
            src_path,
        ]
        subprocess.call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        stem = os.path.splitext(os.path.basename(src_path))[0]
        cand = os.path.join(outdir, stem + "." + out_ext)
        return cand if os.path.isfile(cand) else None

    def sheet_book(path, sheets_data):
        wb = Workbook()
        wb.remove(wb.active)
        for title, meta, headers, rows in sheets_data:
            ws = wb.create_sheet(title)
            ws.append([meta] + [None] * (len(headers) - 1))
            ws.append(headers)
            for r in rows:
                ws.append(r)
        wb.save(path)

    s1_data = []
    for title, tag, fam, ts, base in [
        ("Отчет_январь", "s01/январь", "Фам_s01", "Ts01", 71100),
        ("Отчет_февраль", "s01/февраль", "Фам_s01", "Ts01", 72100),
        ("Отчет_март", "s03/март", "Фам_s03", "Ts03", 73100),
        ("Отчет_апрель", "s01/апрель", "Фам_s01", "Ts01", 74100),
        ("Отчет_май", "s02/май", "Фам_s02", "Ts02", 75100),
        ("Отчет_июнь", "s03/июнь", "Фам_s03", "Ts03", 76100),
    ]:
        hdr = ["ФИО", "Табельный_номер", "Отдел", "Сумма_руб", "Сумма_бонус", "маркер", "ДатаВремя_строка"]
        rows = []
        for i in range(1, 5 if title != "Отчет_январь" else 13):
            rows.append(
                [
                    "%s_%02d" % (fam, i),
                    "%s%02d" % (ts, i),
                    ["Продажи", "IT", "Склад"][i % 3],
                    base + i * 100,
                    1000 + i * 50,
                    "x" if i == 4 else "",
                    "01.01.2024 09:00:00",
                ]
            )
        s1_data.append(("Филиал Северный", hdr, rows))
        s1_data[-1] = (title, "Филиал Северный", hdr, rows)
    sheet_book(
        abs_source("source_01.xlsx"),
        [(t, m, h, r) for t, m, h, r in s1_data],
    )
    # .xlsm: сохранение тем же openpyxl (без VBA, но формат читается openpyxl и Calc)
    try:
        wb_xlsx = openpyxl.load_workbook(abs_source("source_01.xlsx"))
        wb_xlsx.save(abs_source("source_01.xlsm"))
    except Exception:
        pass

    s2_data = []
    for title in ["Отчет_январь", "Отчет_февраль", "Отчет_май", "Сводка2024", "Архив", "Данные2024", "Прочее"]:
        hdr = ["Отдел", "ФИО", "Табельный_номер", "Сумма_руб", "Сумма_бонус"]
        rows = [
            ["Продажи", "Фам_s02_01", "Ts0201", 72100, 1050],
            ["IT", "Фам_s02_02", "Ts0202", 72200, 1100],
            ["Склад", "Фам_s02_03", "Ts0203", 72300, 1150],
            ["Продажи", "Фам_s02_04", "Ts0204", 72400, 1200],
        ]
        s2_data.append((title, "Филиал Южный", hdr, rows))
    sheet_book(abs_source("source_02.xlsx"), s2_data)
    # .xls: только через LibreOffice (если доступен)
    xls_out = _convert(abs_source("source_02.xlsx"), "xls", "MS Excel 97")
    if not xls_out:
        print("Предупреждение: не удалось создать source_02.xls (нужен soffice/libreoffice)")

    s3_data = []
    for title in ["Отчет_январь", "Отчет_март", "Отчет_июнь"]:
        hdr = ["ФИО", "Отдел", "Сумма_руб", "ДатаВремя_строка", "ДатаВремя_native"]
        rows = []
        for i in range(1, 5):
            rows.append(
                [
                    "Фам_s03_%02d" % i,
                    ["Продажи", "IT", "Склад"][i % 3],
                    73100 + i * 100,
                    "0%d.01.2024 09:00:00" % i,
                    datetime.datetime(2024, 1, i, 9, 0),
                ]
            )
        s3_data.append((title, "Филиал Восточный", hdr, rows))
    sheet_book(abs_source("source_03.xlsx"), s3_data)
    # .ods: только через LibreOffice (если доступен)
    ods_out = _convert(abs_source("source_03.xlsx"), "ods", "calc8")
    if not ods_out:
        print("Предупреждение: не удалось создать source_03.ods (нужен soffice/libreoffice)")

    prod = []
    for title, grp in [("Электроника", "Эл"), ("Одежда", "Од"), ("Продукты", "Пр")]:
        hdr = ["Товар", "Цена", "Количество"]
        rows = [["%s_%d" % (grp, i), 100 * i, i] for i in range(1, 4)]
        prod.append((title, "Группа %s" % title, hdr, rows))
    sheet_book(abs_source("source_products.xlsx"), prod)

    catalog = _build_pivot_employee_catalog()
    ref_hdr = [
        "Код_сотрудника",
        "Фамилия",
        "Имя",
        "Отчество",
        "Должность",
        "Дата_рождения",
        "Дата_трудоустройства",
    ]
    pivot_sheets = [
        (
            "Справочник_сотрудников",
            "Справочник персонала",
            ref_hdr,
            [e["row"] for e in catalog],
        ),
    ]
    depts = ["Продажи", "IT", "Склад", "Маркетинг"]
    for q, title in enumerate(["Отчет_Q1", "Отчет_Q2", "Отчет_Q3"], 1):
        hdr = ["ФИО", "Отдел", "Сумма_руб", "Сумма_бонус"]
        rows = []
        for i in range(1, PIVOT_ROWS_PER_SHEET + 1):
            code = _pivot_employee_code_varied(catalog, i, q)
            code_num = int(code.split("_")[-1])
            sum_main = 40000 + q * 5000 + i * 300
            sum_bonus = 800 + i * 40
            rows.append(
                [
                    code,
                    depts[(code_num + i + q) % len(depts)],
                    _fmt_text_number_ru(sum_main),
                    _fmt_text_number_ru(sum_bonus),
                ]
            )
        pivot_sheets.append((title, "Квартал %d" % q, hdr, rows))
    sheet_book(abs_source("source_pivot.xlsx"), pivot_sheets)

    print(
        "Источники: source_01..03 (+ source_01.xlsm/source_02.xls/source_03.ods при возможности), "
        "source_products.xlsx, source_pivot.xlsx (справочник + 3 листа по %d строк)"
        % PIVOT_ROWS_PER_SHEET
    )


def generate_xml_sources():
    """Разнотипные XML-источники для сбора и ВПР (Excel-like flatten)."""
    xml_dir = os.path.join(SOURCES, "xml")
    os.makedirs(xml_dir, exist_ok=True)

    generic_path = os.path.join(xml_dir, "source_catalog.xml")
    with open(generic_path, "w", encoding="utf-8") as f:
        f.write(
            u"""<?xml version="1.0" encoding="UTF-8"?>
<Catalog>
  <Product id="P001">
    <Name>Товар_A</Name>
    <Price>1200</Price>
    <Qty>3</Qty>
  </Product>
  <Product id="P002">
    <Name>Товар_B</Name>
    <Price>2400</Price>
    <Qty>1</Qty>
  </Product>
  <Product id="P003">
    <Name>Товар_C</Name>
    <Price>800</Price>
    <Qty>7</Qty>
  </Product>
</Catalog>
"""
        )

    employees_path = os.path.join(xml_dir, "source_employees.xml")
    with open(employees_path, "w", encoding="utf-8") as f:
        f.write(
            u"""<?xml version="1.0" encoding="UTF-8"?>
<Employees>
  <Employee dept="IT" branch="Северный">
    <Name>Фам_xml_01</Name>
    <Salary>71200</Salary>
  </Employee>
  <Employee dept="Склад" branch="Северный">
    <Name>Фам_xml_02</Name>
    <Salary>72300</Salary>
  </Employee>
  <Employee dept="Продажи" branch="Южный">
    <Name>Фам_xml_03</Name>
    <Salary>73400</Salary>
  </Employee>
  <Employee dept="IT" branch="Южный">
    <Name>Фам_xml_04</Name>
    <Salary>74500</Salary>
  </Employee>
</Employees>
"""
        )

    ss_path = os.path.join(xml_dir, "source_spreadsheetml.xml")
    with open(ss_path, "w", encoding="utf-8") as f:
        f.write(
            u"""<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Worksheet ss:Name="Отчет_январь">
  <Table>
   <Row>
    <Cell><Data ss:Type="String">ФИО</Data></Cell>
    <Cell><Data ss:Type="String">Отдел</Data></Cell>
    <Cell><Data ss:Type="String">Сумма_руб</Data></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="String">Фам_x01</Data></Cell>
    <Cell><Data ss:Type="String">IT</Data></Cell>
    <Cell><Data ss:Type="Number">71100</Data></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="String">Фам_x02</Data></Cell>
    <Cell><Data ss:Type="String">Склад</Data></Cell>
    <Cell><Data ss:Type="Number">71200</Data></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="String">Фам_x03</Data></Cell>
    <Cell><Data ss:Type="String">Продажи</Data></Cell>
    <Cell><Data ss:Type="Number">71300</Data></Cell>
   </Row>
  </Table>
 </Worksheet>
</Workbook>
"""
        )

    nested_path = os.path.join(xml_dir, "source_orders_nested.xml")
    with open(nested_path, "w", encoding="utf-8") as f:
        f.write(
            u"""<?xml version="1.0" encoding="UTF-8"?>
<Orders>
  <Order order_id="O-100">
    <Customer>Клиент_A</Customer>
    <Line><Sku>SKU-1</Sku><Amount>500</Amount></Line>
    <Line><Sku>SKU-2</Sku><Amount>300</Amount></Line>
  </Order>
  <Order order_id="O-101">
    <Customer>Клиент_B</Customer>
    <Line><Sku>SKU-3</Sku><Amount>900</Amount></Line>
  </Order>
</Orders>
"""
        )

    print(
        "XML-источники: %s (catalog, employees, spreadsheetml, orders_nested)"
        % xml_dir
    )


def scenario_01():
    return [
        ("Файлы-Источники", (abs_source("source_01.xlsx"), abs_source("source_02.xlsx"), abs_source("source_03.xlsx"))),
        ("Листы", ("Отчет*",)),
    ] + block_rows(start_row="3", header_row="2") + [
        ("Режим", ("На один лист",)),
        # Headless QA: буфер обмена в --headless пуст; По_API пишет get/setDataArray.
        ("Способ_переноса", ("По_API",)),
        ("Источник_в_первой_колонке", ("Коротко",)),
        ("Дата_источника_во_второй__колонке", ("Не_выводить",)),
        (P_CTC, ("A:1 -> Филиал",)),
        pp_range("тонкая_сетка"),
        pp_range("заголовок_плюс_высота"),
        pp_range("авто_ширина"),
    ]


def scenario_02():
    return [
        ("Файлы-Источники", (abs_source("source_01.xlsx"), abs_source("source_02.xlsx"))),
        ("Листы", ("Все",)),
    ] + block_rows(start_row="3", header_row="2") + [
        ("Режим", ("На разные листы",)),
        (P_CTC, ("A:1 -> Филиал", "A:1 -> Филиал")),
        pp_range("перенос"),
        pp_range("автофильтр"),
    ]


def scenario_03():
    return [
        ("Файлы-Источники", (abs_source("source_01.xlsx"), abs_source("source_02.xlsx"), abs_source("source_03.xlsx"))),
        ("Листы", ("Отчет_январь",)),
        ("Столбцы", ("ФИО;Табельный_номер;Отдел",)),
    ] + block_rows() + [
        ("Режим", ("На один лист",)),
        ("Источник_в_первой_колонке", ("Коротко",)),
        (P_CTC, ("A:1 -> Филиал", "", "")),
    ]


def scenario_04():
    return [
        ("Файлы-Источники", (abs_source("source_01.xlsx"), abs_source("source_02.xlsx"))),
        ("Листы", ("*2024", "Отчет*")),
        ("Столбцы", ("Сумма_руб;ФИО;Отдел",)),
    ] + block_rows() + [("Режим", ("На один лист",))]


def scenario_05():
    return [
        ("Файлы-Источники", (abs_source("source_01.xlsx"), abs_source("source_03.xlsx"))),
        ("Листы", ("Отчет_март",)),
        ("Столбцы", ("Все",)),
    ] + block_rows(col_count="6") + [("Режим", ("На один лист",))]


def scenario_06():
    return [
        ("Файлы-Источники", (abs_source("source_01.xlsx"), abs_source("source_03.xlsx"))),
        ("Листы", ("1",)),
    ] + block_rows() + [("Режим", ("На один лист",))]


def scenario_07():
    return [
        ("Файлы-Источники", (abs_source("source_01.xlsx"), abs_source("source_02.xlsx"))),
        ("Листы", ("Отчет_январь",)),
        ("Столбцы", ("1", "3", "6")),
    ] + block_rows(start_col="1") + [("Режим", ("На один лист",))]


def scenario_08():
    return [
        ("Файлы-Источники", (abs_source("source_01.xlsx"), abs_source("source_02.xlsx"))),
        ("Листы", ("Отчет_январь",)),
        ("Столбцы", ("A", "C", "F")),
    ] + block_rows(start_col="A") + [("Режим", ("На один лист",))]


def scenario_09():
    return [
        ("Файлы-Источники", (abs_source("source_01.xlsx"), abs_source("source_03.xlsx"))),
        ("Листы", ("Отчет_январь",)),
    ] + block_rows(row_count="3") + [("Режим", ("На один лист",))]


def scenario_10():
    return [
        ("Файлы-Источники", (abs_source("source_products.xlsx"),)),
        ("Листы", ("Электроника", "Одежда", "Продукты")),
    ] + block_rows() + [
        ("Режим", ("На один лист",)),
        ("Источник_в_первой_колонке", ("Коротко",)),
        (P_CTC, ("A:1 -> Группа_Товаров",)),
    ]


def scenario_11_direct_xml_excel_ods():
    return [
        (
            "Файлы-Источники",
            (
                abs_source("source_01.xlsm"),
                abs_source("source_02.xls"),
                abs_source("source_03.ods"),
            ),
        ),
        ("Листы", ("Отчет*",)),
    ] + block_rows() + [
        ("Режим", ("На разные листы",)),
        ("Способ_переноса", ("xml_excel_ods",)),
        ("Источник_в_первой_колонке", ("Коротко",)),
        ("Дата_источника_во_второй__колонке", ("Длинно",)),
        pp_range("тонкая_сетка"),
        pp_range("авто_ширина"),
    ]


def scenario_12_postprocess_json():
    """Матрица постобработки с JSON — ручная проверка всех codec-функций RANGE/ROW."""
    rows = [
        ("Файлы-Источники", (abs_source("source_01.xlsx"),)),
        ("Листы", ("Отчет_январь",)),
    ] + block_rows() + [
        ("Режим", ("На один лист",)),
        ("Источник_в_первой_колонке", ("Коротко",)),
        (P_CTC, ("A:1 -> Филиал",)),
    ]
    for fn in RANGE_MATRIX_ORDER:
        rows.append(pp_range(fn))
    for sheet_fn, preset_key in ROW_MATRIX_ORDER:
        rows.append(pp_row(sheet_fn, preset_key))
    return rows


def scenario_13_merge_sheets_pivot():
    """
    На разные листы (source_pivot) → объединить_листы_в_один (A) → сводная.

    «объединить_листы_в_один» — отдельная строка A, JSON в C (не Постобработка_Диапазон).
    """
    import json

    dest = "Отчет_Q123"
    pivot_sheet = "Сводная_Отдел"
    j_pivot = json.dumps(
        [
            {
                "v": 1,
                "fn": "сводная_таблица",
                "source_sheet": dest,
                "sheet_name": pivot_sheet,
                "as_values": True,
                "header_row": "1",
                "row_fields": ["Отдел"],
                "column_fields": ["Квартал"],
                "filter_fields": [],
                "data_fields": [{"field": "Сумма_руб", "function": "SUM"}],
            }
        ],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return [
        ("Файлы-Источники", (abs_source("source_pivot.xlsx"),)),
        ("Листы", ("Отчет*",)),
    ] + block_rows() + [
        ("Режим", ("На разные листы",)),
        ("Источник_в_первой_колонке", ("Не_выводить",)),
        ("Дата_источника_во_второй__колонке", ("Не_выводить",)),
        (P_CTC, ("A:1 -> Квартал",)),
        ("объединить_листы_в_один", "", PRESET["объединить_листы_в_один"]),
        (P_RANGE, "сводная_таблица", j_pivot),
        (P_RANGE, "тонкая_сетка", PRESET.get("тонкая_сетка", "")),
        (P_RANGE, "авто_ширина", PRESET.get("авто_ширина", "")),
    ]


def scenario_14_form_to_table():
    """Копирование листов Q/A → анкета_в_таблицу (ФИО / пустая строка / шапка)."""
    return [
        ("Файлы-Источники", (abs_form_source(),)),
        ("Листы", ("Анкета_ФИО", "Анкета_пусто", "Анкета_шапка")),
    ] + block_rows(start_row="2", header_row="1") + [
        ("Режим", ("Копирование листов",)),
        pp_range("анкета_в_таблицу"),
        pp_range("тонкая_сетка"),
        pp_range("авто_ширина"),
    ]


def scenario_15_table_to_form():
    """Копирование wide → таблица_в_анкету."""
    return [
        ("Файлы-Источники", (abs_form_source(),)),
        ("Листы", ("Wide_анкеты",)),
    ] + block_rows(start_row="2", header_row="1") + [
        ("Режим", ("Копирование листов",)),
        pp_range("таблица_в_анкету"),
        pp_range("тонкая_сетка"),
        pp_range("авто_ширина"),
    ]


def scenario_16_transpose_table():
    """Копирование → транспонировать_таблицу (обычный + headers_from_column + result_headers)."""
    import json

    j_plain = json.dumps(
        [
            {
                "v": 1,
                "fn": "транспонировать_таблицу",
                "sheet": "Матрица",
                "output": "new_sheet",
                "dest_sheet": "Матрица_T",
                "as_values": True,
            }
        ],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    j_hfc = json.dumps(
        [
            {
                "v": 1,
                "fn": "транспонировать_таблицу",
                "sheet": "Метки_колонка",
                "output": "new_sheet",
                "dest_sheet": "Метки_T",
                "headers_from_column": True,
                "as_values": True,
            }
        ],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    j_custom = json.dumps(
        [
            {
                "v": 1,
                "fn": "транспонировать_таблицу",
                "sheet": "Матрица",
                "output": "new_sheet",
                "dest_sheet": "Матрица_H",
                "headers_from_column": False,
                "result_headers": ["Метка", "Строка1", "Строка2"],
                "as_values": True,
            }
        ],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return [
        ("Файлы-Источники", (abs_transpose_source(),)),
        ("Листы", ("Матрица", "Метки_колонка")),
    ] + block_rows(start_row="2", header_row="1") + [
        ("Режим", ("Копирование листов",)),
        (P_RANGE, "транспонировать_таблицу", j_plain),
        (P_RANGE, "транспонировать_таблицу", j_hfc),
        (P_RANGE, "транспонировать_таблицу", j_custom),
        pp_range("тонкая_сетка"),
        pp_range("авто_ширина"),
    ]


VLOOKUP_SCENARIOS = [
    ("11_vlookup_1to1_low.xltx",) + scenario_vlookup("1to1_low") + (5,),
    ("11_vlookup_1to1_middle.xltx",) + scenario_vlookup("1to1_middle") + (5,),
    ("11_vlookup_1to1_high.xltx",) + scenario_vlookup("1to1_high") + (5,),
    ("11_vlookup_1toN_low.xltx",) + scenario_vlookup("1toN_low") + (5,),
    ("11_vlookup_1toN_middle.xltx",) + scenario_vlookup("1toN_middle") + (5,),
    ("11_vlookup_1toN_high.xltx",) + scenario_vlookup("1toN_high") + (5,),
]

SCENARIOS = [
    ("00_all_parameters.xltx", "Справочник параметров сбора (без постобработки).", scenario_01()[:8], 3),
    ("01_merge_one_sheet.xltx", "Сценарий 01: На один лист, JSON-постобработка (сетка, заголовок, зебра).", scenario_01(), 3),
    ("02_merge_many_sheets.xltx", "Сценарий 02: На разные листы.", scenario_02(), 2),
    ("03_columns_filter.xltx", "Сценарий 03: Фильтр столбцов.", scenario_03(), 3),
    ("04_sheets_wildcard.xltx", "Сценарий 04: Шаблоны листов и столбцов.", scenario_04(), 2),
    ("05_all_columns.xltx", "Сценарий 05: Все столбцы.", scenario_05(), 2),
    ("06_sheets_by_index.xltx", "Сценарий 06: Листы по номеру.", scenario_06(), 2),
    ("07_columns_by_index.xltx", "Сценарий 07: Столбцы по номерам.", scenario_07(), 3),
    ("08_columns_by_letters.xltx", "Сценарий 08: Столбцы по буквам.", scenario_08(), 3),
    ("09_fixed_row_count.xltx", "Сценарий 09: Фиксированное число строк.", scenario_09(), 2),
    ("10_product_groups.xltx", "Сценарий 10: Каталог товаров.", scenario_10(), 3),
    (
        "11_direct_xml_excel_ods.xltx",
        "Сценарий 11: Прямой перенос через файлы (Способ_переноса=xml_excel_ods), источники .xlsm/.xls/.ods.",
        scenario_11_direct_xml_excel_ods(),
        3,
    ),
    ("12_postprocess_json.xltx", "Матрица JSON-постобработки: все основные RANGE/ROW функции.", scenario_12_postprocess_json(), 3),
    (
        "13_merge_sheets_pivot.xltx",
        "Сценарий 13: На разные листы → объединить_листы_в_один (строка A) → сводная Отдел×Квартал.",
        scenario_13_merge_sheets_pivot(),
        3,
    ),
    (
        "14_form_to_table.xltx",
        "Сценарий 14: Копирование анкет (Q/A) → анкета_в_таблицу "
        "(by_repeat_key / by_blank_row / шапка листа+блока).",
        scenario_14_form_to_table(),
        3,
    ),
    (
        "15_table_to_form.xltx",
        "Сценарий 15: Копирование wide → таблица_в_анкету (тело + шапка блока).",
        scenario_15_table_to_form(),
        3,
    ),
    (
        "16_transpose_table.xltx",
        "Сценарий 16: транспонировать_таблицу (обычный + headers_from_column).",
        scenario_16_transpose_table(),
        3,
    ),
] + VLOOKUP_SCENARIOS


def write_manual_check():
    lines = [
        "# Ручная сверка collect_workbooks",
        "",
        "Сгенерировано: %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Версия макроса: %s" % macro_version(),
        "",
        "Колонка C для постобработки — **JSON**. Справочник полей: [docs/13_JSON_PARAMS.md](../../docs/13_JSON_PARAMS.md).",
        "Пайплайн xml_excel_ods: [docs/17_POSTPROCESS_XML.md](../../docs/17_POSTPROCESS_XML.md).",
        "",
        "## Сценарии",
        "",
    ]
    for fn, comment, _, _ in SCENARIOS:
        lines.append("- **%s** — %s" % (fn, comment))
    lines.extend(
        [
            "",
            "## 12_postprocess_json.xltx",
            "",
            "Откройте в Calc, запустите макрос, проверьте визуально каждую строку постобработки.",
            "Сверьте журнал «Сбор_книг_лог» — нет ошибок json/пропусков для листа «Отчет_январь».",
            "",
            "### RANGE (колонка B → ожидание)",
        ]
    )
    for fn in RANGE_MATRIX_ORDER:
        lines.append("- `%s`" % fn)
    lines.append("")
    lines.append("### ROW")
    for sheet_fn, _ in ROW_MATRIX_ORDER:
        lines.append("- `%s`" % sheet_fn)
    lines.extend(
        [
            "",
            "## ВПР (11_vlookup_*.xltx)",
            "",
            "Режим «Копирование листов», источник `sources/vlookup/vlookup_{1to1|1toN}_{low|middle|high}.xlsx`.",
            "",
            "| Шаблон | Тип | Объём |",
            "|--------|-----|-------|",
            "| 11_vlookup_1to1_low | 1:1 | ~180 строк, ключ Код_позиции |",
            "| 11_vlookup_1to1_middle | 1:1 | ~850 строк, ключ Регион+Код_товара |",
            "| 11_vlookup_1to1_high | 1:1 | ~3000 строк, ключ 3 поля |",
            "| 11_vlookup_1toN_low | 1:N | дубли справа, low |",
            "| 11_vlookup_1toN_middle | 1:N | дубли справа, middle |",
            "| 11_vlookup_1toN_high | 1:N | дубли справа, high |",
            "",
            "Проверить: подтянуты Цена/Валюта/Поставщик, подсветка #FFF2CC, дубли только в 1toN.",
            "",
            "## xml_excel_ods (11_direct_xml_excel_ods.xltx)",
            "",
            "1. Книга результата сохранена на диск.",
            "2. Direct-фаза может закрыть книгу — нормально.",
            "3. После показа окна — UNO-оформление и Постобработка_Диапазон/Строка (строка состояния).",
            "4. Copy + xml + MERGE_XML_CONVERT_TO_NUMBERS=Да: текстовые числа (1 000 000,00, '1 000 000.00) и даты (с апострофом).",
            "5. Версия в логе ≥ 3.10.79 (отложенная UNO); ≥ 3.10.80 (даты при copy); ≥ 3.10.81 (числа с разрядами при copy).",
            "",
            "## 13_merge_sheets_pivot.xltx",
            "",
            "На разные листы (`source_pivot.xlsx`) → строка A **`объединить_листы_в_один`** "
            "(JSON → `Отчет_Q123`, не ключ B Постобработка_Диапазон) → сводная Отдел×Квартал.",
            "Ожидание: лист `Отчет_Q123`, лист `Сводная_Отдел`; в логе нет ошибок merge_sheets.",
            "",
            "## 14_form_to_table.xltx / 15_table_to_form.xltx",
            "",
            "Источник `sources/forms/source_forms.xlsx`, режим «Копирование листов».",
            "",
            "| Шаблон | Листы | Ожидание |",
            "|--------|-------|----------|",
            "| 14_form_to_table | Анкета_ФИО, Анкета_пусто, Анкета_шапка | "
            "листы `Wide_из_ФИО` (3 строки), `Wide_из_пусто` (2), `Wide_из_шапка` "
            "(2 + столбцы шапки листа/блока) |",
            "| 15_table_to_form | Wide_анкеты | лист `Анкета_из_wide`: пары Q/A, "
            "блоки через пустую строку, шапка блока Бланк/Дата |",
            "",
            "В логе: `анкета_в_таблицу` / `таблица_в_анкету` — ok, без ошибок JSON.",
        ]
    )
    lines.append("")
    with open(MANUAL_CHECK, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def _odf_mimetype(path):
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read("mimetype").decode("ascii", errors="replace").strip()


def _ensure_ots_template(path):
    """ODS с расширением .ots открывается как файл; нужен MIME шаблона ODF."""
    try:
        mime = _odf_mimetype(path)
    except (IOError, OSError, zipfile.BadZipFile, KeyError):
        return False
    if mime == ODF_TEMPLATE_MIME:
        return True
    if mime != ODF_SPREADSHEET_MIME:
        return False

    entries = []
    with zipfile.ZipFile(path, "r") as zin:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "META-INF/manifest.xml":
                data = data.replace(
                    (ODF_SPREADSHEET_MIME + '"').encode("ascii"),
                    (ODF_TEMPLATE_MIME + '"').encode("ascii"),
                )
            entries.append((info.filename, data))

    fd, tmp = tempfile.mkstemp(suffix=".ots", dir=os.path.dirname(path) or ".")
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w") as zout:
            zi = zipfile.ZipInfo("mimetype")
            zi.compress_type = zipfile.ZIP_STORED
            zout.writestr(zi, (ODF_TEMPLATE_MIME + "\n").encode("ascii"))
            for name, data in entries:
                if name == "mimetype":
                    continue
                zout.writestr(name, data)
        shutil.move(tmp, path)
    except Exception:
        if os.path.isfile(tmp):
            os.remove(tmp)
        raise
    return True


def convert_odf():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice не найден — пропуск конвертации ODF")
        return
    os.makedirs(ODF_DIR, exist_ok=True)
    convert_filters = ("ots:calc8_template", "ots")
    for fn, _, _, _ in SCENARIOS:
        src = os.path.join(WORKBOOKS, fn)
        if not os.path.isfile(src):
            continue
        stem = fn.replace(".xltx", "")
        out = os.path.join(ODF_DIR, stem + ".ots")
        wrong_ods = os.path.join(ODF_DIR, stem + ".ods")
        for stale in (out, wrong_ods):
            if os.path.isfile(stale):
                os.remove(stale)

        converted = None
        for conv in convert_filters:
            cmd = [
                soffice,
                "--headless",
                "--convert-to",
                conv,
                "--outdir",
                ODF_DIR,
                src,
            ]
            subprocess.call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if os.path.isfile(out):
                converted = out
                break
            if os.path.isfile(wrong_ods):
                os.rename(wrong_ods, out)
                converted = out
                break

        if not converted:
            print("  пропуск OTS: %s (конвертация не удалась)" % fn)
            continue
        if not _ensure_ots_template(converted):
            print("  предупреждение: %s — не удалось проверить MIME шаблона" % os.path.basename(converted))
    print("ODF: %s" % ODF_DIR)


def main():
    ap = argparse.ArgumentParser(description="Генератор ручных шаблонов collect_workbooks (JSON)")
    ap.add_argument("--skip-sources", action="store_true", help="Не пересоздавать sources/*.xlsx")
    ap.add_argument(
        "--skip-large",
        action="store_true",
        help="Не создавать sources/large/source_server_logs.xlsx (3×200k+ строк, долго)",
    )
    ap.add_argument("--skip-vlookup", action="store_true", help="Не пересоздавать sources/vlookup/*.xlsx")
    ap.add_argument("--skip-forms", action="store_true", help="Не пересоздавать sources/forms/source_forms.xlsx")
    ap.add_argument(
        "--skip-transpose",
        action="store_true",
        help="Не пересоздавать sources/transpose/source_transpose.xlsx",
    )
    ap.add_argument("--odf", action="store_true", help="Конвертировать в workbooks/ODF/*.ots")
    args = ap.parse_args()
    if not args.skip_sources:
        generate_sources()
        generate_xml_sources()
    if not args.skip_large:
        generate_large_server_sources()
    elif not args.skip_sources:
        print("Пропуск большого источника (--skip-large)")
    if not args.skip_vlookup:
        print("Источники ВПР:")
        generate_vlookup_sources()
    if not args.skip_forms:
        print("Источники анкет:")
        generate_form_sources()
    if not args.skip_transpose:
        print("Источники transpose:")
        generate_transpose_sources()
    for fn, comment, body, cols in SCENARIOS:
        path = write_workbook(fn, comment, body, value_cols=cols)
        print("  %s" % path)
    write_manual_check()
    print("  %s" % MANUAL_CHECK)
    if args.odf:
        convert_odf()
    print("Готово. См. docs/12_MANUAL_TESTING.md")


if __name__ == "__main__":
    main()
