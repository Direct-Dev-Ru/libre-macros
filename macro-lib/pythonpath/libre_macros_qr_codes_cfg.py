# -*- coding: utf-8 -*-
"""Константы макроса generate_qr_codes."""

MACRO_VERSION = "3.10.716"
SHEET_GENERATOR = "Генератор_QR"
SHEET_PRINT_TEMPLATE = "QR_Print"

COL_INV = 0
COL_NAME = 1
COL_SN1 = 2
COL_SN2 = 3
COL_SN3 = 4
COL_BAR0 = 5
COL_BAR1 = 6
COL_BAR2 = 7
COL_LEN = 8
COL_COUNTRY = 9
COL_MFR = 10
COL_PRODUCT = 11
COL_PRODUCT2 = 12
COL_SHK0 = 13
COL_SHK1 = 14
COL_SHK2 = 15
COL_SHK1_S1 = 16
COL_SHK1_S2 = 17
COL_SHK2_S1 = 18
COL_SHK2_S2 = 19
COL_SHK3_S1 = 20
COL_SHK3_S2 = 21

DATA_START_ROW = 1
HEADER_ROW = 0

# Адреса ячеек карточки QR_Print (по умолчанию, A1-нотация).
# QR слева (A1:A3), справа: инв.номер / название / серийный.
DEFAULT_MAP_INV = u"B1"
DEFAULT_MAP_NAME = u"B2"
DEFAULT_MAP_SERIAL = u"B3"
DEFAULT_MAP_QR = u"A1"
# Пусто / 0 / <0 → автофит по ячейке; >0 → сторона квадрата QR в мм.
DEFAULT_MAP_SIZE_MM = u""

# Совместимость: 0-based из defaults (A1→0,0; B1→1,0; B2→1,1; B3→1,2).
QR_CELL_COL = 0
QR_CELL_ROW = 0
INV_CELL_COL = 1
INV_CELL_ROW = 0
NAME_CELL_COL = 1
NAME_CELL_ROW = 1
SN_CELL_COL = 1
SN_CELL_ROW = 2

# Шаблон QR_Print: размеры в мм.
QR_PRINT_COL_A_MM = 28
QR_PRINT_COL_B_MM = 30
QR_PRINT_ROW1_MM = 7
QR_PRINT_ROW2_MM = 18
QR_PRINT_ROW3_MM = 7

# Бумага карточки задаётся вшитым generate_qr_template.ods (не через UNO).
# Пересборка blob: python3 macro-lib/bundle_qr_template.py

# Рамки A1:B3: толщина линии в 1/100 мм.
QR_BORDER_OUTER_HMM = 50
QR_BORDER_INNER_HMM = 10
QR_BORDER_COLOR = 0x000000

# Шрифт подписей на карточке.
QR_FONT_NAME = u"PT Sans"
QR_FONT_INV_PT = 10
QR_FONT_NAME_PT = 8
QR_FONT_SERIAL_PT = 10

# JSON настроек (маппинг + шаблон + PDF) — папка как у пресетов визарда.
QR_MAPPING_SUBDIR = u"collect_workbooks"
QR_MAPPING_FILE = u"qr_print_mapping.json"
QR_MAPPING_JSON_VER = 3

DEFAULT_TEMPLATE_PATH = u""
DEFAULT_EXPORT_PDF = False
DEFAULT_PDF_DIR = u""

# Столбцы bar0..2 на листе «Генератор_QR» (источник строки для QR).
DEFAULT_BAR_COL0 = u"F"
DEFAULT_BAR_COL1 = u"G"
DEFAULT_BAR_COL2 = u"H"

# Перезапись bar на листе; если False — берётся значение с листа, пусто → BAR_EMPTY_PLACEHOLDER.
DEFAULT_OVERWRITE_BAR_CODES = True
BAR_EMPTY_PLACEHOLDER = u"_EMPTY_"

QR_MARGIN_MM = 2
QR_SCALE = 8
QR_BORDER = 2

# Оформление автосозданного «Генератор_QR».
GEN_HEADER_BG = 0x5C2D91
GEN_HEADER_FG = 0xFFFFFF

