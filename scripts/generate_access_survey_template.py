#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор шаблона XLSX для опроса подразделений по срокам
предоставления доступов к ИС и ресурсам.

Запуск:
  python scripts/generate_access_survey_template.py
  python scripts/generate_access_survey_template.py --layout form
  python scripts/generate_access_survey_template.py -o /path/to/опрос_доступы.xlsx --layout form

Требуется: openpyxl (pip install openpyxl)
"""

from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT = SCRIPT_DIR / "feedback_form_access_survey.xlsx"

# --- Справочники (редактируйте под организацию) ---

DEPARTMENTS = [
    "Филиал Омск",
    "Филиал Казань",
    "Филиал Самара",
]

CUSTOMERS = [
    "— (собственные системы)",
    "Заказчик Омск 1",
    "Заказчик Казань 1",
    "Заказчик Самара 1",
    "Заказчик Омск 2",
    "Заказчик Казань 2",
    "Заказчик Самара 2",
    "Заказчик Омск 3",
    "Заказчик Казань 3",
    "Заказчик Самара 3",
]

REQUEST_TYPES_OWN = [
    "Доступ к файловым ресурсам собственным",    
    "СЭД (расширенный доступ)",    
    "Удаленный доступ к ресурсам ККС",
]

REQUEST_TYPES_CUSTOMER = [
    "Бухгалтерская Учетная система Заказчика",
    "Электронная система хранения документов Заказчика",    
    "Доступ к файловым ресурсам Заказчика",    
    "Прочие ресурсы Заказчика",
]

BASIC_PACKAGE_ITEMS = [
    "Доменная учётная запись",
    "Электронная почта",
    "Базовый доступ к СЭД",
]

YES_NO = ["Да", "Нет", "Не применимо"]

BOTTLENECK_STAGES = [
    "Этап 1 — подготовка и запуск в СЭД",
    "Этап 2 — согласование и замечания в СЭД",
    "Этап 3 — исполнение в КАСПП (КОИТ)",
    "Ожидание зависимой заявки / УЗ",
    "Ожидание действий заказчика",
    "Ожидание действий подразделения",
    "Не выявлено",
]

FEATURE_CATEGORIES = [
    "Маршрут согласования СЭД",
    "Шаблон заявки",
    "Заказчик — регламент",
    "Заказчик — сроки реакции",
    "Зависимость от других заявок",
    "КАСПП / КОИТ",
    "КИС Доступ / ИБ",
    "Кадровый контур (приказ)",
    "Прочее",
]

OWN_CUSTOMER = ["Собственная", "Заказчика"]

# Поля листов: (подпись, ключ валидации или None)
# Ключи: departments, customers, yes_no, bottleneck, own_customer, categories
DEPARTMENT_FIELDS = [
    ("ID подразделения", None),
    ("Наименование подразделения", "departments"),
    ("Контактное лицо (ФИО)", None),
    ("Должность", None),
    ("E-mail / телефон", None),
    ("Рабочий график для расчёта (если не 09–18 пн–пт)", None),
    ("Среднее число новых сотрудников в месяц", None),
    ("Доля заявок на системы заказчиков, %", None),
    ("Используемый шаблон/маршрут СЭД для доступов", None),
    ("Типовые согласующие (роли)", None),
    ("Комментарий по особенностям процесса", None),
]

CASE_FIELDS = [
    ("ID кейса", None),
    ("ID подразделения", "departments"),
    ("Тип кейса (новый сотрудник / перевод / подрядчик / иное)", None),
    ("Дата точки X (приём в организацию)", None),
    ("Определение точки X в кейсе", None),
    ("Приказ подписан: от X, раб.ч", None),
    ("Ознакомление с регламентами (КИС Доступ, ИБ): от X, раб.ч", None),
    ("Базовый пакет полностью выдан: от X, раб.ч", None),
    ("Доменная УЗ: от X, раб.ч", None),
    ("Почта: от X, раб.ч", None),
    ("Базовый СЭД: от X, раб.ч", None),
    ("Задержка базового пакета, раб.ч (если была)", None),
    ("Причина задержки базового пакета", None),
    ("Число доп. заявок в кейсе", None),
    ("Комментарий к кейсу", None),
]

REQUEST_FIELDS = [
    ("ID заявки", None),
    ("ID кейса", None),
    ("ID подразделения", "departments"),
    ("Собственная / Заказчика", "own_customer"),
    ("Заказчик", "customers"),
    ("Система / ресурс", None),
    ("Тип заявки (из справочника или свой)", None),
    ("Критичность для начала работы", None),
    ("Зависит от (ID заявки или «Базовый пакет: …»)", None),
    ("Можно подать до выдачи базового пакета?", "yes_no"),
    ("Факт старта подготовки: от X, раб.ч", None),
    ("Э1: длит. подготовки, раб.ч", None),
    ("Э1: ожидание, раб.ч", None),
    ("Э1: запуск в СЭД: от X, раб.ч", None),
    ("Э2: длит. согласования, раб.ч", None),
    ("Э2: число циклов замечаний", None),
    ("Э2: ожидание согласующих, раб.ч", None),
    ("Э2: завершение согласования: от X, раб.ч", None),
    ("Э3: длит. передачи в КАСПП, раб.ч", None),
    ("Э3: длит. исполнения КОИТ, раб.ч", None),
    ("Э3: ожидание в КАСПП, раб.ч", None),
    ("Э3: доступ выдан: от X, раб.ч", None),
    ("Полный цикл заявки, раб.ч", None),
    ("Ожидание зависимости, раб.ч", None),
    ("Узкое место (этап)", "bottleneck"),
    ("Ответственная сторона задержки", None),
    ("Особенности / примечания", None),
]

FEATURE_FIELDS = [
    ("ID записи", None),
    ("ID подразделения", "departments"),
    ("Заказчик (если применимо)", "customers"),
    ("Категория", "categories"),
    ("Описание особенности / проблемы", None),
    ("Как сейчас обходят", None),
    ("Предложение по улучшению", None),
    ("Влияние на срок (оценка, раб.ч)", None),
]

# --- Оформление ---

FILL_HEADER = PatternFill("solid", fgColor="FF2F5496")
FILL_SECTION = PatternFill("solid", fgColor="FFD9E1F2")

FONT_HEADER = Font(bold=True, color="FFFFFFFF", size=11)
FONT_SECTION = Font(bold=True, size=11)
FONT_NORMAL = Font(size=10)

ALIGN_TOP = Alignment(vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

THIN = Side(style="thin", color="FFB4B4B4")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_header_row(ws: Worksheet, row: int, col_count: int) -> None:
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_ALL


def _style_data_area(ws: Worksheet, start_row: int, end_row: int, col_count: int) -> None:
    for r in range(start_row, end_row + 1):
        for c in range(1, col_count + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = FONT_NORMAL
            cell.alignment = ALIGN_TOP
            cell.border = BORDER_ALL


def _set_col_widths(ws: Worksheet, widths: dict[int, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def _write_table(
    ws: Worksheet,
    headers: list[str],
    start_row: int = 1,
    sample_rows: int = 0,
) -> int:
    col_count = len(headers)
    for col, title in enumerate(headers, start=1):
        ws.cell(row=start_row, column=col, value=title)
    _style_header_row(ws, start_row, col_count)
    first_data = start_row + 1
    last_data = first_data + sample_rows - 1 if sample_rows else first_data
    if sample_rows:
        _style_data_area(ws, first_data, last_data, col_count)
    return last_data


def _add_list_validation(
    ws: Worksheet,
    sqref: str,
    options: list[str],
    *,
    allow_blank: bool = True,
) -> None:
    if not options:
        return
    # Excel ограничивает inline-список ~255 символов; длинные — через скрытый лист.
    formula = '"' + ",".join(options) + '"'
    if len(formula) > 250:
        return
    dv = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=allow_blank,
        showErrorMessage=True,
        errorTitle="Выбор из списка",
        error="Выберите значение из выпадающего списка.",
    )
    ws.add_data_validation(dv)
    dv.add(sqref)


def _add_range_validation(ws: Worksheet, sqref: str, ref_sheet: str, ref_range: str) -> None:
    dv = DataValidation(
        type="list",
        formula1=f"='{ref_sheet}'!{ref_range}",
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Выбор из списка",
        error="Выберите значение из справочника.",
    )
    ws.add_data_validation(dv)
    dv.add(sqref)


def _build_ref_sheet(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet("Справочники")
    ws.sheet_state = "hidden"

    lists = [
        ("A", "Подразделения", DEPARTMENTS),
        ("C", "Заказчики", CUSTOMERS),
        ("E", "Типы_собственные", REQUEST_TYPES_OWN),
        ("G", "Типы_заказчика", REQUEST_TYPES_CUSTOMER),
        ("I", "Да_Нет", YES_NO),
        ("K", "Узкие_места", BOTTLENECK_STAGES),
        ("M", "Базовый_пакет", BASIC_PACKAGE_ITEMS),
    ]

    for col_letter, title, items in lists:
        col = ord(col_letter) - 64
        ws.cell(row=1, column=col, value=title).font = FONT_SECTION
        for i, item in enumerate(items, start=2):
            ws.cell(row=i, column=col, value=item)

    return ws


def _ref_range(col_letter: str, count: int) -> str:
    return f"${col_letter}$2:${col_letter}${count + 1}"


def _validation_spec(key: str) -> tuple[str, str] | tuple[str, list[str]]:
    """Справочник на листе «Справочники» или inline-список."""
    ref_map = {
        "departments": ("Справочники", _ref_range("A", len(DEPARTMENTS))),
        "customers": ("Справочники", _ref_range("C", len(CUSTOMERS))),
        "yes_no": ("Справочники", _ref_range("I", len(YES_NO))),
        "bottleneck": ("Справочники", _ref_range("K", len(BOTTLENECK_STAGES))),
    }
    inline_map = {
        "own_customer": OWN_CUSTOMER,
        "categories": FEATURE_CATEGORIES,
    }
    if key in ref_map:
        return ref_map[key]
    if key in inline_map:
        return ("__inline__", inline_map[key])
    raise KeyError(key)


def _apply_validation_key(ws: Worksheet, sqref: str, key: str) -> None:
    spec = _validation_spec(key)
    if spec[0] == "__inline__":
        _add_list_validation(ws, sqref, spec[1])
    else:
        _add_range_validation(ws, sqref, spec[0], spec[1])


def _fields_to_headers(fields: list[tuple[str, str | None]]) -> list[str]:
    return [label for label, _ in fields]


def _write_vertical_form_block(
    ws: Worksheet,
    row: int,
    fields: list[tuple[str, str | None]],
    *,
    block_title: str | None = None,
    value_rows: dict[str, list[int]] | None = None,
) -> int:
    if value_rows is None:
        value_rows = {}

    if block_title:
        title_cell = ws.cell(row=row, column=1, value=block_title)
        title_cell.font = FONT_SECTION
        title_cell.fill = FILL_SECTION
        title_cell.alignment = ALIGN_TOP
        title_cell.border = BORDER_ALL
        value_cell = ws.cell(row=row, column=2, value="")
        value_cell.fill = FILL_SECTION
        value_cell.border = BORDER_ALL
        row += 1

    ws.cell(row=row, column=1, value="Поле")
    ws.cell(row=row, column=2, value="Значение")
    _style_header_row(ws, row, 2)
    row += 1

    for label, val_key in fields:
        label_cell = ws.cell(row=row, column=1, value=label)
        label_cell.font = FONT_NORMAL
        label_cell.alignment = ALIGN_TOP
        label_cell.border = BORDER_ALL
        value_cell = ws.cell(row=row, column=2, value="")
        value_cell.font = FONT_NORMAL
        value_cell.alignment = ALIGN_TOP
        value_cell.border = BORDER_ALL
        if val_key:
            value_rows.setdefault(val_key, []).append(row)
        row += 1

    return row


def _write_vertical_form_sheet(
    wb: Workbook,
    sheet_name: str,
    fields: list[tuple[str, str | None]],
    *,
    block_count: int,
    block_title_template: str,
) -> Worksheet:
    ws = wb.create_sheet(sheet_name)
    _set_col_widths(ws, {1: 48, 2: 42})

    value_rows: dict[str, list[int]] = {}
    row = 1
    for block_idx in range(1, block_count + 1):
        row = _write_vertical_form_block(
            ws,
            row,
            fields,
            block_title=block_title_template.format(n=block_idx),
            value_rows=value_rows,
        )
        row += 1

    for key, rows in value_rows.items():
        for r in rows:
            _apply_validation_key(ws, f"B{r}", key)

    ws.freeze_panes = "A2"
    return ws


def _build_instruction_sheet(wb: Workbook, *, layout: str) -> None:
    ws = wb.active
    ws.title = "Инструкция"
    ws.column_dimensions["A"].width = 110

    lines = [
        ("Опрос: узкие места при предоставлении доступов к ИС", FONT_SECTION),
        ("", FONT_NORMAL),
        (
            "Цель — собрать фактические сроки и особенности работы с заявками "
            "в подразделениях и у разных заказчиков, выявить задержки по этапам жизненного цикла.",
            FONT_NORMAL,
        ),
        ("", FONT_NORMAL),
        ("Контекст процесса", FONT_SECTION),
        (
            "• При трудоустройстве после подписания приказа и ознакомления с регламентами "
            "(отметка сотрудника ИБ в КИС Доступ) выдаётся базовый пакет: доменная УЗ, "
            "электронная почта, базовый доступ к СЭД.\n"
            "• Остальные доступы — по заявкам (собственные ИС и ресурсы заказчиков).\n"
            "• Жизненный цикл заявки:\n"
            "  1) подготовка и запуск в СЭД;\n"
            "  2) согласование в СЭД и устранение замечаний;\n"
            "  3) передача из СЭД на исполнение в КАСПП корпоративному оператору ИТ (КОИТ).",
            FONT_NORMAL,
        ),
        ("", FONT_NORMAL),
        ("Точка отсчёта X", FONT_SECTION),
        (
            "X — момент приёма сотрудника в организацию (дата из приказа о приёме / фактический "
            "первый рабочий день — укажите в листе «Кейсы» единообразно для всех строк кейса).\n"
            "Все поля «от X, раб.ч» — накопленное рабочее время от точки X до наступления события "
            "(с учётом только рабочих часов: пн–пт, 09:00–18:00, без праздников; при другом "
            "графике подразделения — укажите в листе «Подразделения»).",
            FONT_NORMAL,
        ),
        ("", FONT_NORMAL),
        ("Как заполнять", FONT_SECTION),
        (
            (
                "1. Лист «Подразделения» — одна анкета на отвечающее подразделение.\n"
                "2. Лист «Кейсы» — типовые ситуации приёма сотрудника, дата точки X, сроки базового пакета.\n"
                "3. Лист «Заявки» — каждый блок = одна заявка на доступ; укажите зависимости, "
                "длительности этапов в рабочих часах и моменты от X.\n"
                "4. Лист «Особенности» — качественные наблюдения по шаблонам, согласующим и заказчикам."
                if layout == "form"
                else
                "1. Лист «Подразделения» — одна строка на отвечающее подразделение.\n"
                "2. Лист «Кейсы» — типовые ситуации приёма сотрудника (новый / перевод / подрядчик и т.д.), "
                "дата точки X, сроки базового пакета.\n"
                "3. Лист «Заявки» — каждая строка = одна заявка на доступ в рамках кейса; "
                "укажите зависимости, длительности этапов в рабочих часах и моменты от X.\n"
                "4. Лист «Особенности» — качественные наблюдения: шаблоны, согласующие, "
                "специфика заказчиков, типовые причины возвратов."
            ),
            FONT_NORMAL,
        ),
        ("", FONT_NORMAL),
        ("Поля длительности (раб.ч)", FONT_SECTION),
        (
            "«Длит. этапа, раб.ч» — чистое время этапа (без параллельного ожидания вне этапа). "
            "«Ожидание, раб.ч» — простой внутри этапа (очередь, отсутствие согласующего и т.п.).\n"
            "Если точные даты неизвестны — оцените по последним 3–5 аналогичным случаям (мин / тип / макс).",
            FONT_NORMAL,
        ),
        ("", FONT_NORMAL),
        ("Зависимости заявок", FONT_SECTION),
        (
            "Если запуск заявки возможен только после другой заявки или появления доменной УЗ — "
            "укажите ID зависимой заявки или «Базовый пакет: доменная УЗ» в соответствующих колонках. "
            "Отдельно зафиксируйте задержку из-за зависимости в «Ожидание зависимости, раб.ч».",
            FONT_NORMAL,
        ),
    ]

    for i, (text, font) in enumerate(lines, start=1):
        cell = ws.cell(row=i, column=1, value=text)
        cell.font = font
        cell.alignment = ALIGN_TOP
        if i == 1:
            cell.fill = FILL_SECTION


def _build_department_sheet(wb: Workbook, *, layout: str) -> Worksheet:
    if layout == "form":
        return _write_vertical_form_sheet(
            wb,
            "Подразделения",
            DEPARTMENT_FIELDS,
            block_count=5,
            block_title_template="Подразделение №{n}",
        )

    headers = _fields_to_headers(DEPARTMENT_FIELDS)
    ws = wb.create_sheet("Подразделения")
    widths = {
        1: 14, 2: 28, 3: 22, 4: 18, 5: 22, 6: 26, 7: 16, 8: 14, 9: 30, 10: 24, 11: 36,
    }
    _set_col_widths(ws, widths)
    last = _write_table(ws, headers, sample_rows=5)
    _add_range_validation(ws, f"B2:B{last}", "Справочники", _ref_range("A", len(DEPARTMENTS)))
    ws.freeze_panes = "A2"
    return ws


def _build_cases_sheet(wb: Workbook, *, layout: str) -> Worksheet:
    if layout == "form":
        return _write_vertical_form_sheet(
            wb,
            "Кейсы",
            CASE_FIELDS,
            block_count=10,
            block_title_template="Кейс №{n}",
        )

    headers = _fields_to_headers(CASE_FIELDS)
    ws = wb.create_sheet("Кейсы")
    widths = {
        1: 10, 2: 14, 3: 28, 4: 14, 5: 24, 6: 14, 7: 14, 8: 14, 9: 12, 10: 12,
        11: 12, 12: 14, 13: 28, 14: 12, 15: 32,
    }
    _set_col_widths(ws, widths)
    last = _write_table(ws, headers, sample_rows=10)
    _add_range_validation(ws, f"B2:B{last}", "Справочники", _ref_range("A", len(DEPARTMENTS)))
    ws.freeze_panes = "A2"
    return ws


def _build_requests_sheet(wb: Workbook, *, layout: str) -> Worksheet:
    if layout == "form":
        return _write_vertical_form_sheet(
            wb,
            "Заявки",
            REQUEST_FIELDS,
            block_count=30,
            block_title_template="Заявка №{n}",
        )

    headers = _fields_to_headers(REQUEST_FIELDS)
    ws = wb.create_sheet("Заявки")
    widths = {
        1: 10, 2: 10, 3: 14, 4: 14, 5: 22, 6: 22, 7: 28, 8: 14, 9: 26, 10: 14,
        11: 14, 12: 12, 13: 12, 14: 14, 15: 12, 16: 12, 17: 14, 18: 14,
        19: 12, 20: 12, 21: 12, 22: 14, 23: 12, 24: 14, 25: 22, 26: 22, 27: 32,
    }
    _set_col_widths(ws, widths)
    last = _write_table(ws, headers, sample_rows=30)

    ref = "Справочники"
    _add_range_validation(ws, f"C2:C{last}", ref, _ref_range("A", len(DEPARTMENTS)))
    _add_list_validation(ws, f"D2:D{last}", OWN_CUSTOMER)
    _add_range_validation(ws, f"E2:E{last}", ref, _ref_range("C", len(CUSTOMERS)))
    _add_range_validation(ws, f"J2:J{last}", ref, _ref_range("I", len(YES_NO)))
    _add_range_validation(ws, f"Y2:Y{last}", ref, _ref_range("K", len(BOTTLENECK_STAGES)))

    ws.freeze_panes = "A2"
    return ws


def _build_features_sheet(wb: Workbook, *, layout: str) -> Worksheet:
    if layout == "form":
        return _write_vertical_form_sheet(
            wb,
            "Особенности",
            FEATURE_FIELDS,
            block_count=15,
            block_title_template="Запись №{n}",
        )

    headers = _fields_to_headers(FEATURE_FIELDS)
    ws = wb.create_sheet("Особенности")
    widths = {1: 10, 2: 14, 3: 22, 4: 22, 5: 40, 6: 30, 7: 30, 8: 14}
    _set_col_widths(ws, widths)
    last = _write_table(ws, headers, sample_rows=15)
    _add_range_validation(ws, f"B2:B{last}", "Справочники", _ref_range("A", len(DEPARTMENTS)))
    _add_range_validation(ws, f"C2:C{last}", "Справочники", _ref_range("C", len(CUSTOMERS)))
    _add_list_validation(ws, f"D2:D{last}", FEATURE_CATEGORIES)
    ws.freeze_panes = "A2"
    return ws


def _build_summary_sheet(wb: Workbook, *, layout: str) -> None:
    ws = wb.create_sheet("Сводка_для_анализа")
    ws["A1"] = "Агрегирующие метрики (заполняется аналитиком после сбора опросов)"
    ws["A1"].font = FONT_SECTION
    ws["A1"].fill = FILL_SECTION

    if layout == "form":
        source_hint = "вертикальные блоки на листах"
    else:
        source_hint = "табличные строки на листах"

    rows = [
        ("Метрика", "Формула / источник", "Значение"),
        ("Медиана: базовый пакет, раб.ч от X", f"«Кейсы», поле «Базовый пакет полностью выдан» ({source_hint})", ""),
        ("Медиана: полный цикл заявки, раб.ч", f"«Заявки», поле «Полный цикл заявки» ({source_hint})", ""),
        ("Медиана: Э1 подготовка, раб.ч", f"«Заявки», поле «Э1: длит. подготовки» ({source_hint})", ""),
        ("Медиана: Э2 согласование, раб.ч", f"«Заявки», поле «Э2: длит. согласования» ({source_hint})", ""),
        ("Медиана: Э3 КАСПП+КОИТ, раб.ч", f"Сумма полей Э3 на листе «Заявки» ({source_hint})", ""),
        ("Доля заявок с узким местом Э2", f"«Заявки», поле «Узкое место» ({source_hint})", ""),
        ("Среднее число циклов замечаний", f"«Заявки», поле «Э2: число циклов замечаний» ({source_hint})", ""),
        ("Среднее ожидание зависимостей, раб.ч", f"«Заявки», поле «Ожидание зависимости» ({source_hint})", ""),
    ]
    for i, row in enumerate(rows, start=3):
        for j, val in enumerate(row, start=1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.border = BORDER_ALL
            cell.alignment = ALIGN_TOP
            if i == 3:
                cell.fill = FILL_HEADER
                cell.font = FONT_HEADER
    _set_col_widths(ws, {1: 36, 2: 36, 3: 18})


def build_workbook(*, layout: str = "table") -> Workbook:
    if layout not in ("table", "form"):
        raise ValueError(f"Неизвестный layout: {layout!r} (ожидается 'table' или 'form')")

    wb = Workbook()
    _build_ref_sheet(wb)
    _build_instruction_sheet(wb, layout=layout)
    _build_department_sheet(wb, layout=layout)
    _build_cases_sheet(wb, layout=layout)
    _build_requests_sheet(wb, layout=layout)
    _build_features_sheet(wb, layout=layout)
    _build_summary_sheet(wb, layout=layout)
    return wb


def main() -> None:
    parser = argparse.ArgumentParser(description="Сгенерировать шаблон опросного листа по доступам (XLSX).")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Путь к выходному файлу (по умолчанию: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--layout",
        choices=("table", "form"),
        default="table",
        help="table — горизонтальная таблица (по умолчанию); form — вертикальная анкета (поле | значение)",
    )
    args = parser.parse_args()

    output: Path = args.output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    wb = build_workbook(layout=args.layout)
    wb.save(output)
    layout_label = "таблица" if args.layout == "table" else "анкета (вертикальные блоки)"
    print(f"Шаблон сохранён: {output}")
    print(f"Формат: {layout_label}")
    print("Листы: Инструкция, Подразделения, Кейсы, Заявки, Особенности, Сводка_для_анализа, Справочники (скрыт)")
    print("Перед рассылкой отредактируйте справочники в начале скрипта или на скрытом листе «Справочники».")


if __name__ == "__main__":
    main()
