# -*- coding: utf-8 -*-
"""Константы и состояние collect_workbooks (AlterOffice 2026)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.710"
import re

try:
    unicode
except NameError:
    unicode = str

try:
    unichr
except NameError:
    unichr = chr

_MERGE_RESULT_POSTPROCESS_RANGE = None
_MERGE_RESULT_POSTPROCESS_ROW = None
_MERGE_PROCESSING_PIPELINE = None
_MERGE_FINAL_PROCESSING = None
_MERGE_DELETE_TOP_ROWS_MAP = None
_MERGE_DELETE_TOP_ROWS_PER_FILE = None
_MERGE_DELETE_TOP_ROWS_DONE = set()
_MERGE_RESULT_SKIPPED_TOP_ROWS = {}
_MERGE_COPY_SHEETS_BLOCK_BY_NAME = {}
_MERGE_COPY_SHEETS_FILE_INDEX = {}
# Листы CSV: не гонять UNO/direct текст→число/дата
# (уже в матрице, либо copy_whole_sheet — Calc распознал при открытии).
_MERGE_CSV_PRECONVERTED_SHEETS = set()
_MERGE_DIRECT_XML_ACTIVE = False
_MERGE_DIRECT_XML_FORMULAS_IN_FILE = frozenset()
_MERGE_DIRECT_XML_FILE_VISUAL_SHEETS = frozenset()
# Имена листов книги результата на момент старта этапа сбора (run_collect_merge).
# Для них не применять встроенное оформление по умолчанию.
_MERGE_COLLECT_PREEXISTING_SHEETS = frozenset()
# Режим последнего run_collect_merge (для collect_reset_db_ranges и т.п.).
_MERGE_COLLECT_MODE = None
# Имена листов, вовлечённых в режиме «Текущие листы» (финальная обработка только по ним).
_MERGE_INVOLVED_SHEET_NAMES = frozenset()
# Sentinel для filter_fn в iter_sheet_slots: «не передан» → взять global.
# Явный None = без фильтра (не подставлять устаревший global).
# Живёт в cfg: в Scripts/python/*.py модульные присваивания AO 2026 вырезает.
_MERGE_FILTER_FN_UNSET = object()
# Листы, на которых постобработка «автофильтр» уже поставила меню-фильтр
# (чтобы MERGE_RESULT_AUTOFILTER не переключил .uno:DataFilterAutoFilter обратно).
_MERGE_SHEETS_WITH_MENU_AUTOFILTER = set()
# Листы-исключения из очистки smart-диапазонов (экспериментальный режим автофильтра).
_MERGE_KEEP_SMART_TABLE_SHEETS = set()
# True — после финальной обработки нужен save/close/reopen,
# чтобы smart-диапазон + TableStyle проявились в UI (AO/LO).
_MERGE_NEED_SMART_TABLE_REOPEN = False
# sheet_name -> table_style: стили smart-таблиц для повторного назначения после reopen.
_MERGE_SMART_TABLE_STYLES = {}
_MERGE_FORMAT_HEADER_COLS = {}
_MERGE_XML_DEFERRED_UNO_JOB = None
# False — после direct открыть файл результата и сделать UNO-оформление/ПП.
MERGE_XML_SKIP_UNO_PHASE = False
# True — early-relocate: якорь → системный temp до direct; запись в исходный путь
# (для .ods: якорь temp xlsx → OS-копия рядом как .xlsx; финал storeAsURL → .ods;
#  якорь закрывается, sidecar xlsx не удаляется).
MERGE_XML_NO_REPLACE = True
# Устарело для early-relocate (финал всегда open исходного пути).
MERGE_XML_SKIP_NO_REPLACE_FINISH = True
MERGE_XML_OPEN_RESULT_AFTER_FINISH = True
# True — после gut закрыть окно якоря (temp); False — откат (оставить пустую книгу открытой).
MERGE_XML_CLOSE_ANCHOR_AFTER_GUT = True
_MERGE_XML_FINISH_HANDLED = False
_MERGE_XML_POSTPROCESS_PIPELINE = []
_MERGE_PP_ROW_LOGGED = set()
_MERGE_PP_ACTIVE_CTX = None
_MERGE_PP_PIPELINE_STEP_DONE = set()
_MERGE_PP_CURRENT_FN_LABEL = None
MERGE_PIVOT_POST_VALUES_FORMAT = True
MERGE_PARAM_SHEET_NAME = "collect_params"
MERGE_PARAM_SHEET_MAX_ROWS = 200
MERGE_HELP_SHEET_NAME = "Справка_макроса"
MERGE_SERVICE_SHEET_NAME = "Индексы_Источников"
MERGE_LOG_SHEET_BASE = "Сбор_книг_лог"
MERGE_LOG_SHEET_NAME = MERGE_LOG_SHEET_BASE  # активный журнал запуска (оркестратор может переопределить)
# Скрытый лист списков валидации визарда параметров (param_wizard _PARAM_DV_SHEET_NAME).
MERGE_WIZARD_LISTS_SHEET_NAME = "_Визард_Списки"
MERGE_PATH_HEADER = "#Путь"
MERGE_DATETIME_HEADER = "#ДатаВремяФайла"
# Сообщение на пустом листе результата («на разные листы», MERGE_DO_EMPTY_SHEETS).
MERGE_EMPTY_SHEET_MESSAGE = "данных для переноса не обнаружено"
MERGE_SERVICE_DATETIME_FMT_LONG = "DD.MM.YYYY HH:MM:SS"
MERGE_SERVICE_DATETIME_FMT_SHORT = "DD.MM.YYYY"
_MERGE_SERVICE_DT_NF_KEYS = {}
MERGE_FIRST_DATA_COL = 2
_MERGE_ACTIVE_SERVICE_LAYOUT = None
_MERGE_ACTIVE_DATA_TRANSFER_MODE = None
# На время одной книги-источника: По_API, если среди слотов сбора есть скрытый лист.
_MERGE_FILE_TRANSFER_OVERRIDE = None
_MERGE_ROW_SOURCE_BY_SHEET = {}
_MERGE_ADD_TO_VARIABLES_PER_FILE = []
_MERGE_PER_FILE_SOURCE_EXTRA = []
_MERGE_SOURCE_VARIABLES_MAP = {}
_MERGE_SOURCE_VARIABLES_ORDER = []
MERGE_SHEET_INDEX_TARGET_PREFIX = "Merge_"
MERGE_ONE_SHEET_TARGET_NAME = "Merge"
# Легаси-список (точное lower-совпадение); распознавание — через
# merge_normalize_current_book_marker_key + _MERGE_CURRENT_BOOK_MARKER_KEYS.
_MERGE_CURRENT_BOOK_MARKERS = ("эта_книга", "this_book", "current_book", "current")
# Каноны после нормализации (без регистра/пробелов/_/-/.).
_MERGE_CURRENT_BOOK_MARKER_KEYS = frozenset(
    (
        "этакнига",
        "thisworkbook",
        "thisbook",
        "currentworkbook",
        "currentbook",
        "current",
    )
)
MERGE_CURRENT_BOOK_MARKER = "Эта_книга"


def merge_normalize_current_book_marker_key(path):
    """
    Ключ маркера текущей книги: lower, ё→е, без пробелов/_/-/.

    Примеры → «этакнига» / «thisworkbook» / «current».
    """
    try:
        s = str(path or "").strip().lower()
    except Exception:
        return ""
    if not s:
        return ""
    try:
        s = s.replace("ё", "е")
    except Exception:
        pass
    out = []
    for ch in s:
        if ch in " \t_-.":
            continue
        out.append(ch)
    return "".join(out)


def is_current_book_marker_key(path):
    """True, если path — маркер текущей книги (гибкая нормализация)."""
    key = merge_normalize_current_book_marker_key(path)
    if not key:
        return False
    return key in _MERGE_CURRENT_BOOK_MARKER_KEYS
MERGE_SHEETS_UNDELETE_DEFAULT = [
    MERGE_HELP_SHEET_NAME,
    MERGE_LOG_SHEET_NAME,
    MERGE_SERVICE_SHEET_NAME,
    MERGE_WIZARD_LISTS_SHEET_NAME,
]
MERGE_SHEETS_UNDELETE_KEEP_ALL = object()
_MERGE_PP_DELETE_SHEETS_DEFAULT = (
    MERGE_HELP_SHEET_NAME,
    MERGE_LOG_SHEET_NAME,
    MERGE_SERVICE_SHEET_NAME,
)
_MERGE_RUN_PARAM_SHEET = None
# Headless QA (qa/collect_headless): без MessageBox / confirm UI.
_MERGE_QA_HEADLESS = False
_MERGE_LO_COLOR_AUTO = -1
_MERGE_LO_COLOR_NAMES = {
    "navy": (0, 0, 128),
    "white": (255, 255, 255),
    "gray": (158, 158, 158),
    "grey": (158, 158, 158),
    "darkblue": (0, 0, 139),
    "excel_header": (31, 78, 121),  # #1F4E79 — тёмно-синий заголовок Excel
}
_MERGE_DELETE_TOP_ROWS_SEPS = ("->", "=", "\\")
# Источник истины — LM_SANITIZE_BUILTIN_BLOCKLIST (совпадает с AST-запретами).
try:
    from libre_macros_sanitize_lib import LM_SANITIZE_BUILTIN_BLOCKLIST as _MERGE_PP_EVAL_BUILTIN_BLOCKLIST
except Exception:
    _MERGE_PP_EVAL_BUILTIN_BLOCKLIST = frozenset(
        {
            "__import__",
            "eval",
            "exec",
            "execfile",
            "compile",
            "open",
            "input",
            "raw_input",
            "breakpoint",
            "help",
            "exit",
            "quit",
            "copyright",
            "credits",
            "license",
            "globals",
            "locals",
            "vars",
            "dir",
            "getattr",
            "setattr",
            "delattr",
            "hasattr",
            "memoryview",
            "bytearray",
            "buffer",
            "file",
            "reload",
            "__builtins__",
        }
    )
_MERGE_PP_EVAL_BUILTINS_CACHE = None
_MERGE_PP_FILE_FN_CACHE = {}
_MERGE_FINAL_FILE_FN_CACHE = {}
_MERGE_MACRO_SCRIPT_DIR_CACHE = None
_VLOOKUP_KNOWN_CELL_LABELS = frozenset(
    (
        u"левая_таблица",
        u"правая_таблица",
        u"ключ_сравнения",
        u"ключ_сравнения_левый",
        u"ключ_сравнения_правый",
        u"столбцы_для_извлечения",
        u"цвет_добавленных_данных",
        u"заполнять_дубли",
        u"заполнитель_для_не_найдено",
    )
)
VLOOKUP_DEFAULT_NOT_FOUND_FILL = u"#Н/Д"
VLOOKUP_HIGHLIGHT_COLOR_NONE = u"Нет"
VLOOKUP_NOT_FOUND_EMPTY = u"_ПУСТО_"
VLOOKUP_NOT_FOUND_EMPTY_ALIASES = (u"_ПУСТО_", u"__ПУСТО__", u"_EMPTY_")
VLOOKUP_PARAM_COLUMN_HEADERS = (
    u"Левая_Таблица",
    u"Правая_Таблица",
    u"Ключ_сравнения",
    u"Столбцы_для_извлечения",
    u"Цвет_добавленных_данных",
)
def lo_color_rgb(red, green, blue):
    """
    Сборка цвета LibreOffice Calc в формате 0xRRGGBB.

    Параметры:
        red, green, blue — компоненты 0…255.

    Возвращает:
        int — CellBackColor / CharColor2 (не BGR).

    Связь:
        lo_color_hex, lo_color_named; _MERGE_FMT_* и merge_pp_* в секции постобработки.
    """
    return (int(red) << 16) | (int(green) << 8) | int(blue)
def lo_color_hex(hex_rgb):
    """
    Разбор цвета из HEX-строки или имени из _MERGE_LO_COLOR_NAMES.

    Параметры:
        hex_rgb — '#RRGGBB', 'RRGGBB' или ключ словаря (navy, excel_header, …).

    Возвращает:
        int — 0xRRGGBB для свойств ячейки.

    Исключения:
        ValueError — неверный формат HEX.

    Связь:
        lo_color_rgb, lo_color_named.
    """
    key = str(hex_rgb).strip().lower()
    if key in _MERGE_LO_COLOR_NAMES:
        t = _MERGE_LO_COLOR_NAMES[key]
        return lo_color_rgb(t[0], t[1], t[2])
    h = key.lstrip("#")
    if len(h) != 6:
        raise ValueError("цвет: ожидается #RRGGBB или имя, получено: %s" % hex_rgb)
    return lo_color_rgb(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
def lo_color_named(name):
    """
    Удобный доступ к именованным цветам через lo_color_hex.

    Параметры:
        name — ключ _MERGE_LO_COLOR_NAMES или HEX.

    Возвращает:
        int — 0xRRGGBB.

    Связь:
        lo_color_hex; оформление заголовков и merge_pp_range_set_font.
    """
    return lo_color_hex(name)
_MERGE_FMT_HEADER_BG = lo_color_rgb(217, 217, 217)
_MERGE_FMT_HEADER_FG = None
_MERGE_FMT_SERVICE_FG = lo_color_rgb(158, 158, 158)
_MERGE_FMT_CHAR_WEIGHT_BOLD = 150
_MERGE_FONT_SIZE = 11
# Дизайн динамического диапазона (AO/LO 26+): первый стиль «Светлые».
MERGE_RESULT_TABLE_STYLE = "TableStyleLight1"
MERGE_RESULT_TABLE_USE_ROW_STRIPES = True

MERGE_FLAG_DEFAULTS = {
    "MERGE_RESULT_AUTOFILTER": False,
    "MERGE_FREEZE_HEADER": True,
    # Ограничение применения по именам листов результата:
    # None — применять ко всем листам, [] — не применять ни на одном.
    "MERGE_RESULT_AUTOFILTER_SHEETS": None,
    "MERGE_FREEZE_HEADER_SHEETS": None,
    "MERGE_PROGRESS_UI_MODE": "dialog",
    "MERGE_DEBUG": False,
    "MERGE_DEBUG_CONSOLE_LOG_TO_WORKSHEET": False,
    # True — создавать лист «Сбор_книг_лог»; False — только консоль (MERGE_DEBUG и т.д.).
    "MERGE_CREATE_LOG_SHEET": True,
    # False — не вызывать _merge_finalize_header_row_appearance (стили только из постобработки).
    "MERGE_FINALIZE_HEADER": False,
    # False (по умолчанию): удалять «лишние» листы на этапе 1 без запроса.
    # True: спросить подтверждение (maybe_clear_sheets) и удалить при согласии.
    "MERGE_CONFIRM_DELETION_OF_SHEETS": False,
    # xml_excel_ods: преобразовывать текст «700», «1 000 000,00», «01.02.2024» в число/дату при чтении файла.
    "MERGE_XML_CONVERT_TO_NUMBERS": True,
    # xml_excel_ods: допустимые форматы даты/времени при преобразовании текста → date/datetime.
    # Пусто/не задано — набор по умолчанию (см. direct_lib).
    "MERGE_XML_DATETIME_FORMATS": "",
    # Маски «Файлы-Источники»: искать совпадения и во всех подпапках стартовой директории.
    "MERGE_SOURCE_VIEW_SUBFOLDERS": True,
    # Доп. маски исключения (через запятую), напр. --* ; файлы с именем ~* пропускаются всегда.
    "MERGE_SOURCE_EXCLUDE_MASK": "--*",
    # xml_excel_ods + диалог: True — закрыть диалог прогресса сразу после прямой фазы.
    "MERGE_UI_AUTO_CLOSE_AFTER_DIRECT": False,
    # «Копирование листов»: автообрезка по Строка_Заголовков / Начало_Строка.
    # False — не обрезать; явное «удаление_верхних_строк» по-прежнему работает.
    "MERGE_COPY_AUTO_TRIM": True,
    # Буфер_по_заголовкам / Буфер_по_позиции: после вставки значений
    # дополнительно специальная вставка форматов (HARDATTR|STYLES / Flags=T).
    # False (по умолчанию) — только значения (InsertContents SVD).
    "MERGE_PASTE_FORMATS": False,
    # «На разные листы»: при 0 строк к переносу всё равно создать лист результата
    # с заголовком и строкой «данных для переноса не обнаружено».
    # False — пустые целевые листы не оставлять (удалить после сбора).
    "MERGE_DO_EMPTY_SHEETS": True,
    # True — этап 2 через executeDispatch (отдельный вызов макроса).
    # False (по умолчанию) — прямой collect_workbooks_run(): в AlterOffice 2026
    # executeDispatch(collect_workbooks_run) часто зависает после «Этап 1 OK».
    "MERGE_SCHEDULE_STAGE2_VIA_DISPATCH": False,
}
MERGE_RESULT_AUTOFILTER = MERGE_FLAG_DEFAULTS["MERGE_RESULT_AUTOFILTER"]
MERGE_FREEZE_HEADER = MERGE_FLAG_DEFAULTS["MERGE_FREEZE_HEADER"]
MERGE_RESULT_AUTOFILTER_SHEETS = MERGE_FLAG_DEFAULTS["MERGE_RESULT_AUTOFILTER_SHEETS"]
MERGE_FREEZE_HEADER_SHEETS = MERGE_FLAG_DEFAULTS["MERGE_FREEZE_HEADER_SHEETS"]
MERGE_PROGRESS_UI_MODE = MERGE_FLAG_DEFAULTS["MERGE_PROGRESS_UI_MODE"]
MERGE_DEBUG = MERGE_FLAG_DEFAULTS["MERGE_DEBUG"]
MERGE_DEBUG_CONSOLE_LOG_TO_WORKSHEET = MERGE_FLAG_DEFAULTS[
    "MERGE_DEBUG_CONSOLE_LOG_TO_WORKSHEET"
]
MERGE_FINALIZE_HEADER = MERGE_FLAG_DEFAULTS["MERGE_FINALIZE_HEADER"]
MERGE_COPY_AUTO_TRIM = MERGE_FLAG_DEFAULTS["MERGE_COPY_AUTO_TRIM"]
MERGE_PASTE_FORMATS = MERGE_FLAG_DEFAULTS["MERGE_PASTE_FORMATS"]
MERGE_DO_EMPTY_SHEETS = MERGE_FLAG_DEFAULTS["MERGE_DO_EMPTY_SHEETS"]
MERGE_SCHEDULE_STAGE2_VIA_DISPATCH = MERGE_FLAG_DEFAULTS[
    "MERGE_SCHEDULE_STAGE2_VIA_DISPATCH"
]
MERGE_XML_CONVERT_TO_NUMBERS = MERGE_FLAG_DEFAULTS["MERGE_XML_CONVERT_TO_NUMBERS"]
MERGE_XML_DATETIME_FORMATS = MERGE_FLAG_DEFAULTS["MERGE_XML_DATETIME_FORMATS"]
MERGE_SOURCE_VIEW_SUBFOLDERS = MERGE_FLAG_DEFAULTS["MERGE_SOURCE_VIEW_SUBFOLDERS"]
MERGE_SOURCE_EXCLUDE_MASK = MERGE_FLAG_DEFAULTS["MERGE_SOURCE_EXCLUDE_MASK"]
MERGE_PROGRESS_UI_STATUS = "status"
MERGE_PROGRESS_UI_DIALOG = "dialog"
_MERGE_FMT_MAX_COL_WIDTH_INCH = 3.0
_MERGE_FMT_COL_WIDTH_PAD_INCH = 0.5
_MERGE_COPY_DATETIME_FMT = "DD.MM.YYYY"
_MERGE_NF_TYPE_DATE = 2
_MERGE_NF_TYPE_TIME = 4
_MERGE_FMT_HEADER_ROW_PAD_INCH = 0.5
P_MERGE_VERSION = "Версия"
P_MERGE_COMMENT = "Комментарий"
P_MERGE_FILES = "Файлы-Источники"
P_MERGE_SOURCE_EXTRA = "Доп_Параметры_Источника"
P_MERGE_PRE_SHELL = "Предварительный_скрипт"
P_MERGE_MODE = "Режим"
P_MERGE_SHEETS = "Листы"
P_MERGE_COLUMNS = "Столбцы"
P_MERGE_START_ROW = "Начало_Строка"
P_MERGE_START_COL = "Начало_Столбец"
P_MERGE_ROW_COUNT = "Число_Строк"
P_MERGE_COL_COUNT = "Число_Столбцов"
P_MERGE_HEADER_ROW = "Строка_Заголовков"
P_MERGE_PATH_DISPLAY = "Источник_в_первой_колонке"
P_MERGE_DATETIME_DISPLAY = "Дата_источника_во_второй_колонке"
P_MERGE_DATETIME_DISPLAY_LEGACY = "Дата_источника_во_второй__колонке"
P_MERGE_SHEETS_UNDELETE = "Листы_не_удалять"
P_MERGE_SHEETS_FORCE_DELETE = "Листы_удалить"
P_MERGE_CELL_TO_COLUMN = "Ячейка_В_Столбец"
P_MERGE_ADD_TO_VARIABLES_MAP = "Добавить_в_карту_переменных"
P_MERGE_MANUAL_VARIABLE_INPUT = "Ручной_ввод_в_карту_переменных"
# Служебные сегменты ключа карты для ручного ввода: name~ручной_ввод~ручной_ввод
MANUAL_VARIABLE_FILE_KEY = "ручной_ввод"
MANUAL_VARIABLE_SHEET_KEY = "ручной_ввод"
# Ответы ручного ввода текущего прогона: list of (name, text, type_token)
_MERGE_MANUAL_VARIABLE_ANSWERS = []
# Кэш диалогов в рамках одного запуска (parse вызывается на этапе 1 и 2).
# dict: cache_key -> (text, type_token)
_MERGE_MANUAL_VARIABLE_CACHE = {}
# True: этап 1 / pack уже показал ручной ввод; этап 2 не спрашивает снова.
_MERGE_MANUAL_INPUT_DONE = False
P_MERGE_SKIP_SOURCE_ROWS = "Пропуск_строк_источника"
P_MERGE_DELETE_TOP_ROWS = "удаление_верхних_строк"
P_MERGE_COPY_AUTO_TRIM = "Автообрезка_листов"
P_MERGE_PASTE_FORMATS = "MERGE_PASTE_FORMATS"
P_MERGE_RESULT_AUTOFILTER = "Автофильтр_результата"
P_MERGE_FREEZE_HEADER = "Закрепить_заголовок"
P_MERGE_PROGRESS_UI = "Отображение_прогресса"
P_MERGE_DEBUG = "MERGE_DEBUG"
P_MERGE_DEBUG_CONSOLE_LOG = "MERGE_DEBUG_CONSOLE_LOG_TO_WORKSHEET"
P_MERGE_CREATE_LOG_SHEET = "MERGE_CREATE_LOG_SHEET"
P_MERGE_DO_EMPTY_SHEETS = "MERGE_DO_EMPTY_SHEETS"
P_MERGE_XML_CONVERT_TO_NUMBERS = "MERGE_XML_CONVERT_TO_NUMBERS"
P_MERGE_XML_DATETIME_FORMATS = "MERGE_XML_DATETIME_FORMATS"
P_MERGE_SOURCE_VIEW_SUBFOLDERS = "MERGE_SOURCE_VIEW_SUBFOLDERS"
P_MERGE_SOURCE_EXCLUDE_MASK = "MERGE_SOURCE_EXCLUDE_MASK"
P_MERGE_XML_RESULT_DIR = "Папка_результата"
P_MERGE_POSTPROCESS_RANGE = "Постобработка_Диапазон"
P_MERGE_POSTPROCESS_ROW = "Постобработка_Строка"
P_MERGE_POSTPROCESS_XML = "Постобработка_xml"
MERGE_PLUGIN_FUNCTION_KEY = "функция_плагин"
P_MERGE_FINAL_PROCESSING = "Финальная_обработка"
P_MERGE_VLOOKUP = "ВПР"
P_MERGE_SHEETS_INTO_ONE = "объединить_листы_в_один"
P_SPLIT_SHEETS = "разделить_листы"
P_MERGE_DATA_TRANSFER = "Способ_переноса"
MERGE_TRANSFER_API = "по_api"
MERGE_TRANSFER_CLIPBOARD_HEADERS = "буфер_по_заголовкам"
MERGE_TRANSFER_CLIPBOARD_POSITION = "буфер_по_позиции"
MERGE_TRANSFER_XML_EXCEL_ODS = "xml_excel_ods"
# Пустая C у «Режим» / пустой «Способ_переноса» → этот режим.
MERGE_TRANSFER_DEFAULT = MERGE_TRANSFER_API
MERGE_MODE_ONE = "на один лист"
MERGE_MODE_MANY = "на разные листы"
MERGE_MODE_COPY = "копирование листов"
MERGE_MODE_INVOLVE = "текущие листы"
# Компактные ключи синонимов «Текущие листы» (без пробелов/_/-); legacy: вовлечьлисты.
_MERGE_MODE_INVOLVE_KEYS = frozenset(
    (
        "текущиелисты",
        "currentsheets",
        "вовлечьлисты",
        "involvesheets",
        "passthrough",
        "безкопирования",
        "толькообработка",
    )
)
_MERGE_OPEN_SOURCES = []
_MERGE_XML_TEMP_PATHS = []
_MERGE_NUMBER_FORMAT_CACHE = {}
_MERGE_RUN_TARGET_DOC = None
_MERGE_LOG_NEXT_ROW = 1
_MERGE_MACRO_RUN_URLS = (
# URL макроса collect_workbooks_run для schedule_collect_run (этап 1 → этап 2).
    "vnd.sun.star.script:collect_workbooks.py$collect_workbooks_run"
    "?language=Python&location=user",
    "vnd.sun.star.script:collect_workbooks.collect_workbooks_run"
    "?language=Python&location=user",
)
MERGE_ADVANCED_DEFAULTS = {
    "MERGE_UI_YIELD_EVERY_ROWS": 10,
    "MERGE_UI_STATUS_EVERY_ROWS": 300,
    "MERGE_COPY_ROW_CHUNK": 1000,
    "MERGE_BATCH_COPY_MIN_ROWS": 30,
    "MERGE_ROW_POSTPROCESS_MAX_ROWS": 200000,
    "MERGE_FORMATTING_MAX_ROWS": 200000,
    "MERGE_FONT_SIZE": _MERGE_FONT_SIZE,
}
MERGE_ADVANCED_PARAM_NAMES = tuple(sorted(MERGE_ADVANCED_DEFAULTS.keys()))
MERGE_ADVANCED_PARAM_MIN = {
    "MERGE_UI_YIELD_EVERY_ROWS": 1,
    "MERGE_UI_STATUS_EVERY_ROWS": 1,
    "MERGE_COPY_ROW_CHUNK": 1,
    "MERGE_BATCH_COPY_MIN_ROWS": 0,
    "MERGE_ROW_POSTPROCESS_MAX_ROWS": 0,
    "MERGE_FORMATTING_MAX_ROWS": 0,
    "MERGE_FONT_SIZE": 1,
}
MERGE_UI_YIELD_EVERY_ROWS = MERGE_ADVANCED_DEFAULTS["MERGE_UI_YIELD_EVERY_ROWS"]
MERGE_UI_STATUS_EVERY_ROWS = MERGE_ADVANCED_DEFAULTS["MERGE_UI_STATUS_EVERY_ROWS"]
MERGE_COPY_ROW_CHUNK = MERGE_ADVANCED_DEFAULTS["MERGE_COPY_ROW_CHUNK"]
MERGE_BATCH_COPY_MIN_ROWS = MERGE_ADVANCED_DEFAULTS["MERGE_BATCH_COPY_MIN_ROWS"]
MERGE_ROW_POSTPROCESS_MAX_ROWS = MERGE_ADVANCED_DEFAULTS["MERGE_ROW_POSTPROCESS_MAX_ROWS"]
MERGE_FORMATTING_MAX_ROWS = MERGE_ADVANCED_DEFAULTS["MERGE_FORMATTING_MAX_ROWS"]
MERGE_FONT_SIZE = MERGE_ADVANCED_DEFAULTS["MERGE_FONT_SIZE"]
_merge_ui_yield_counter = 0
_merge_ui_status_indicator = None
_merge_ui_progress_dialog = None
_merge_ui_progress_label = None
_merge_ui_progress_detail_label = None
_merge_ui_progress_bar = None
_merge_ui_progress_abort_btn = None
_merge_ui_progress_finish_btn = None
_merge_ui_dialog_handler = None
_merge_ui_confirm_handler = None
_merge_ui_abort_requested = False
_merge_ui_active_settings = None
_merge_ui_active_report = None
_merge_ui_anchor_doc = None
_merge_ui_anchor_frame = None
_merge_ui_anchor_frame_name = "LibreMacrosProgress"
_merge_xml_freeze_pending = None
_merge_ui_progress_scale = 1000
_merge_ui_phase = ""
_merge_ui_phase_detail = ""
_merge_ui_file_index = 0
_merge_ui_file_total = 0
_merge_ui_fname = ""
_merge_ui_sheet_name = ""
_merge_ui_log_lines = []
_MERGE_UI_LOG_MAX_LINES = 400
_MERGE_UI_DIALOG_COPY_CHUNK = 50
# Буфер UNO: мелкий чанк, чтобы чаще обновлять диалог. По_API использует MERGE_COPY_ROW_CHUNK.
_MERGE_UI_DIALOG_YIELD_EVERY_ROWS = 1
_MERGE_UI_DIALOG_STATUS_EVERY_ROWS = 10
_MERGE_UI_DIALOG_EVENT_MS = 250
_merge_debug_t0 = None
_MERGE_DEBUG_LOG_DOC = None
_MERGE_CONSOLE_LOG_BUFFER = []
LOG_HEADERS = (
    "Время",
    "Источник",
    "Лист",
    "Столбец",
    "Строк",
    "Статус",
    "Примечание",
)
DEBUG_LOG_HEADERS = (
    "Время",
    "Сообщение",
)
_MERGE_LOG_DATA_COL_COUNT = 6
_MERGE_LOG_NOTE_COL_INDEX = 6
_MERGE_LOG_NOTE_WIDTH_INCH = 7.0
_MERGE_LOG_AUTOFILTER_NAME = "AF_SborKnigLog"  # legacy: снять при оформлении (меню-автофильтр)
PARAM_DISABLED_PREFIX = "__"
_SKIP_FORMULA_CELL_REF = re.compile(r"([A-Za-z]+)(\d+)")
_SKIP_FORMULA_BRACKET_REF = re.compile(r"\[([^\]]+)\]\{([^}]+)\}")
_SKIP_FORMULA_NAME_ROWID_REF = re.compile(
    r"(?<![A-Za-zА-Яа-яЁё_])([A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9_]*)\{([^}]+)\}"
)
_ROW_ID_BRACE_RE = re.compile(r"\$\{([^}]+)\}|\{([^}]+)\}")
_SKIP_HEADER_SCAN_LIMIT = 300
_MERGE_SKIP_FILTER_DB = "__merge_skip_filter__"
_MERGE_COPY_SHEET_LAST_ERROR = ""
macros_lib = {}
_PP_SHIMS_INSTALLED = False
_PP_MAPS_READY = False
_MERGE_PP_FONT_NAME = "PT Sans"
_MERGE_PP_FORMAT_TOTAL_SUM = "# ##0,00"
_MERGE_PP_GRID_REST_HINT = (
    "JSON: [{\"v\":1,\"fn\":\"сетка\",\"width\":\"тонкая\",\"color\":\"серый\",\"include_header\":true}] "
    "или с \"sheet\":\"Лист\"",
    '[{"v":1,"fn":"сетка","width":25,"color":"#CCCCCC","sheet":"Сводная"}]',
)
_MERGE_PP_GRID_PRESET_REST_HINT = (
    "JSON: [{\"v\":1,\"fn\":\"тонкая_сетка\"}] или с \"sheets\":[\"Лист1\",\"Лист2\"]",
    '[{"v":1,"fn":"тонкая_сетка","sheets":["Сводная","Orders"]}]',
)
_MERGE_PP_ZEBRA_REST_HINT = (
    "JSON: roles header/even/odd с fill и font",
    '[{"v":1,"fn":"зебра_диапазон","roles":{"header":{"fill":"excel_header","font":"белый"},"even":{"fill":"голубой","font":"синий"}}}]',
)
_LM_PP_REQUIRED_BINDINGS = (
    ("merge_pp_range_sort_data", "lm_pp_range_sort_data"),
    ("merge_pp_range_delete_sheets", "lm_pp_range_delete_sheets"),
    ("merge_pp_range_hide_sheets", "lm_pp_range_hide_sheets"),
    ("merge_final_sort", "lm_final_sort"),
    ("merge_pp_range_colorize_data", "lm_pp_range_colorize_data"),
    ("merge_final_colorize", "lm_final_colorize_blocks"),
    ("merge_pp_range_merge_sheets_into_one", "lm_pp_range_merge_sheets_into_one"),
    ("merge_final_merge_sheets_into_one", "lm_final_merge_sheets_into_one"),
    ("merge_final_remove_duplicates", "lm_final_remove_duplicates"),
    ("merge_pp_range_conditional_column", "lm_pp_range_conditional_column"),
    ("merge_final_conditional_column", "lm_final_conditional_column"),
    ("merge_final_activate_sheet", "lm_final_activate_sheet"),
    ("merge_final_send_mail", "lm_final_send_mail"),
)
MERGE_POSTPROCESS_RANGE_REST_HINTS = {
    "сетка": _MERGE_PP_GRID_REST_HINT,
    "тонкая_сетка": _MERGE_PP_GRID_PRESET_REST_HINT,
    "толстая_сетка": _MERGE_PP_GRID_PRESET_REST_HINT,
    "заголовок_плюс_высота": (
        "JSON: height_mm (абсолют по умолчанию, мин. 7 мм), add_height=true — плюс к текущей; "
        "h_align/v_align, шрифт/цвета. Пусто — 12,7 мм center/center на всех листах.",
        '[{"v":1,"fn":"заголовок_плюс_высота","height_mm":13,"add_height":false,"h_align":"center","v_align":"center"}]',
    ),
    "перенос": ("пусто = включить; No/Нет — выключить перенос", "Нет"),
    "перенос_и_авто_высота": ("как у «перенос»: пусто = включить; No/Нет — выключить", "Нет"),
    "авто_высота": ("автоподбор высоты строк данных (OptimalHeight)", ""),
    "авто_ширина": ("автоподбор ширины столбцов диапазона (OptimalWidth)", ""),
    "высота_строки": (
        "пусто — все строки данных, 5 мм; «6» — все строки; «2,3 / 6» (или ->, =); «0» — скрыть",
        "2,3,4 / 6",
    ),
    "левое_выравнивание": ("не используется", ""),
    "отступ": (
        "[left|center|right], число_шагов [, столбцы]; пусто = left, 1, все",
        "right,1,Сумма",
    ),
    "шрифт": (
        "имя;размер_данных;размер_заголовка (или через запятую); «-» = не менять; пусто — MERGE_FONT_SIZE",
        "PT Sans;11;11",
    ),
    "зебра_диапазон": _MERGE_PP_ZEBRA_REST_HINT,
    "сортировка": (
        "«лист | ключи» или «лист: ключи»; несколько листов — «Л1,Л2|…»; "
        "задания через ;. Без листа — ключи для всех листов. "
        "Ключ: колонка -> +/- (разделитель ->, /, = или « - »; блоки через , или ;)",
        "Сводная | Сумма -> -, Дата -> + ; Orders | Сумма -> -",
    ),
    "раскрасить_блоки": (
        "«лист | столбцы_ключа -> цвета» или «лист: столбцы = цвета»; блоки через ; или ,; "
        "без листа — для текущего листа результата. "
        "Столбцы: индекс (1=A), буква или заголовок. "
        "Цвета: имя или #hex; несколько — через %; один цвет — оттенки по группам. "
        "Сортировка: /+ или /asc — по возрастанию, /- или /убыв — по убыванию; без суффикса — без сортировки; "
        "!граница — контур блоков",
        "Сводная | Отдел,Город -> голубой%белый/+!граница ; Orders | A -> #E2EFDA",
    ),
    "раскрасить": (
        "алиас «раскрасить_блоки»",
        "Сводная | Отдел -> голубой",
    ),
    "формат_деньги": (
        "sheet, format (пусто = # ##0,00; Excel #,##0.00 → Calc), columns, markers",
        '[{"v":1,"fn":"формат_деньги","sheet":"Отчет","columns":[5,6],"format":"# ##0,00"}]',
    ),
    "формат_даты": (
        "лист/sheets, format (пусто = глобальный), columns, markers",
        '{"sheet":"Отчет","format":"DD.MM.YYYY","markers":"дата"}',
    ),
    "формат_столбцы": (
        "столбец: имя, номер, буква, список через запятую, Все/*; формат — шаблон (пусто = не менять)",
        "ФИО, Отдел",
    ),
    "подсветка_по_порогу": (
        "маркер/столбцы, порог или min/max; JSON: color (пусто=без заливки), style_columns, whole_row, font_color, bold, italic",
        '{"columns":["\'Итого\'"],"threshold_min":0,"threshold_max":1000,"style_columns":["\'Статус\'"],"font_color":"darkblue","bold":true}',
    ),
    "подкрасить_пороги": (
        "алиас «подсветка_по_порогу»",
        '{"columns":["\'Итого\'"],"threshold_min":0,"threshold_max":1000,"style_columns":["\'Статус\'"],"font_color":"darkblue"}',
    ),
    "градиент": (
        "маркер столбца, цвет min, цвет max; JSON: whole_row, sort, sort_dir (по_возр/по_убыв)",
        "сумм -> зеленый -> красный",
    ),
    "удалить_столбцы": (
        "markers: подстроки/имена заголовка, индексы 1-based, буквы A/C; смесь через запятую",
        "#Путь,3,A,Коммент*",
    ),
    "конкатенация_столбцов": (
        "имя_нового_столбца, разделитель, индексы столбцов 1-based (1=A); без разделителя — сразу индексы",
        "Полный_Код,|,1,3,5",
    ),
    "разделить_по_столбцам": (
        "JSON: column, delimiter, position, relative_column, source_policy, split_mode, max_columns",
        '[{"v":1,"fn":"разделить_по_столбцам","column":"ФИО","delimiter":",","position":"_после_","source_policy":"replace"}]',
    ),
    "условный_столбец": (
        "JSON: new_column + rules[] + otherwise; value/value_column и then/then_column; otherwise может содержать вложенный rules[]",
        '[{"v":1,"fn":"условный_столбец","new_column":"Категория","rules":[{"column":"Сумма","op":"ge","value":1000,"then":"Средний"}],"otherwise":"Мелкий"}]',
    ),
    "добавить_столбец": (
        "JSON: new_column, data_type (text|number|date), format; text→choices[]; "
        "number/date→data_validation + min/max",
        '[{"v":1,"fn":"добавить_столбец","new_column":"Статус","data_type":"text","choices":["Активен","Архив"]}]',
    ),
    "столбец_по_лямбде": (
        "JSON: new_column или new_column_expr (lambda sheet, headers), expr (lambda rows...) + upd_rows/*; позиции как у условного_столбца.",
        '[{"v":1,"fn":"столбец_по_лямбде","new_column":"ИНН_копия","expr":"lambda rows: rows(\\"ИНН\\")","position":"_конец_","on_error":"keep","result_type":"auto","output_policy":"new"}]',
    ),
    "применить_формулу": (
        "имя_столбца=формула Calc; {row_id} — первая строка данных (1-based); "
        "{row_id-1} — смещение; as_values — только Диапазон/Финал (в Постобработка_xml игнорируется)",
        "Итого=E{row_id}+F{row_id}",
    ),
    "удаление_строк": (
        "столбцы (1,A,ФИО) или формула «пустая строка»: ИСТИНА → удалить (E{row_id}+F{row_id}=0)",
        "1,A,ФИО",
    ),
    "сброс_стрипов": ("не используется", ""),
    "группировка_по_столбцу": (
        "JSON: marker + agg_fn/agg_columns (или aggs[]); подряд два шага с одним marker "
        "склеиваются в уровни итогов + grand_total; кнопка «Параметры…»",
        '[{"v":1,"fn":"группировка_по_столбцу","marker":"\'Отдел\'","agg_fn":"count","agg_columns":["\'Сумма_руб\'"]},'
        '{"v":1,"fn":"группировка_по_столбцу","marker":"\'Отдел\'","agg_fn":"sum","agg_columns":["\'Сумма_руб\'"]}]',
    ),
    "закрепить_заголовок": (
        "JSON: sheets — список листов; без sheets — все листы результата",
        '[{"v":1,"fn":"закрепить_заголовок","sheets":["Заказы"]}]',
    ),
    "ширина_столбцов": (
        "JSON: columns (all / список заголовков в '…') + width_mm; 0 мм — скрыть; "
        "кнопка «Параметры…»",
        '[{"v":1,"fn":"ширина_столбцов","columns":"all","width_mm":30}]',
    ),
    "автофильтр": (
        "JSON: sheets — список листов; без sheets — все листы. "
        "В xml_excel_ods — UNO-добивка после reopen (как закрепить_заголовок)",
        '[{"v":1,"fn":"автофильтр","sheets":["Заказы"]}]',
    ),
    "стиль_печати": (
        "ориентация (landscape/portrait), вписать в N страниц по ширине",
        "landscape,1",
    ),
    "заполнение_вниз": (
        "столбцы: номер (1-based), буква A или имя заголовка; через , или ;",
        "ФИО;Отдел;3",
    ),
    "заполнить_вверх": (
        "столбцы из списка заголовков ('ФИО'); пустые → значение из ближайшей снизу",
        '[{"v":1,"fn":"заполнить_вверх","columns":["\'ФИО\'","\'Отдел\'"]}]',
    ),
    "развернуть_столбцы": (
        "Unpivot: unpivot_columns (или exclude_columns); attribute_column + value_column; "
        "output=inplace|new_sheet; drop_empty_rows",
        '[{"v":1,"fn":"развернуть_столбцы","unpivot_columns":["\'Январь\'","\'Февраль\'"],'
        '"attribute_column":"Месяц","value_column":"Сумма","drop_empty_rows":true}]',
    ),
    "транспонировать_таблицу": (
        "Транспонирование: output=inplace|new_sheet|offset; range/header_row; "
        "headers_from_column + header_column; result_headers; dest_sheet/dest_cell",
        '[{"v":1,"fn":"транспонировать_таблицу","output":"new_sheet","dest_sheet":"Матрица_T",'
        '"result_headers":["Метка","Янв","Фев"]}]',
    ),
    "анкета_в_таблицу": (
        "Пары вопрос/ответ → wide: question_column/answer_column, block_mode, "
        "block_start_question; output=new_sheet|inplace|replace_sheet",
        '[{"v":1,"fn":"анкета_в_таблицу","question_column":"A","answer_column":"B",'
        '"block_mode":"by_repeat_key","block_start_question":"ФИО","dest_sheet":"Анкеты_wide"}]',
    ),
    "таблица_в_анкету": (
        "Wide → пары вопрос/ответ: body_columns; question_header/answer_header "
        "(по умолчанию Вопрос/Ответ); has_header_in_output; block_separator; output",
        '[{"v":1,"fn":"таблица_в_анкету","body_columns":["\'ФИО\'","\'Возраст\'"],'
        '"question_header":"Вопрос","answer_header":"Ответ","has_header_in_output":true,'
        '"block_separator":"blank_row","dest_sheet":"Анкеты"}]',
    ),
    "заполнение_вниз_вычислить": (
        "JSON: columns + rules (column, formula, expand_to_right); кнопка «Параметры…»",
        '[{"v":1,"fn":"заполнение_вниз_вычислить","columns":["A"],"rules":[{"column":"A","formula":"=A{row-1}"}]}]',
    ),
    "копировать_значения": (
        "столбцы через , или ; — номер (1-based), буква A или имя заголовка; пусто — весь лист; Copy → вставка значений",
        "Сумма,Количество",
    ),
    "замена_значений": (
        "JSON: columns+find; replace[] или replace_mode=lambda+replace_expr (кортеж); "
        "нехватка replace → частичная замена + warning; флаги, row_filter, match_pick",
        '[{"v":1,"fn":"замена_значений","columns":["ФИО"],"find":["ООО"],"replace":["<<Переменные.Филиал>>"],"case_insensitive":true,"squeeze_spaces":true}]',
    ),
    "текстовые_операции": (
        "JSON: columns + ops[] (trim/collapse_ws/case/lang_homoglyphs/substring/…); "
        "row_filter per-op; non_text/on_error; key|key_cell|key_file. "
        "Кнопка «Параметры…» / «Справка».",
        '[{"v":1,"fn":"текстовые_операции","columns":["ФИО"],"ops":[{"op":"collapse_ws"},{"op":"lang_homoglyphs"},{"op":"case","mode":"proper"}]}]',
    ),
    "переименовать_лист": (
        "новое имя: <<Переменные…>> / [Старое_имя] / [Заголовок(N)] / [Строка(R,C)] / "
        "[Текущая_Дата_Время_YYYYMMDD] / [Имя_Книги] / [Номер_Листа] / [GUID8]; "
        "формат даты YYYY/YY/MM/DD/HH/mm/SS; нормализация ≤31",
        "Свод_[Текущая_Дата_Время_YYYYMMDD]",
    ),
    "переименовать_столбцы": (
        "JSON: sheet + mappings; old в 'кавычках' = полное совпадение (запятые внутри не режут); "
        "в new — те же плейсхолдеры, что у переименовать_лист "
        "(<<Переменные…>>, [Старое_имя], дата/время, [Имя_Книги], …)",
        '[{"v":1,"fn":"переименовать_столбцы","sheet":"Заказы",'
        '"mappings":[{"old":"\'Сумма\'","new":"<<Переменные.Филиал>>_[Старое_имя]"}]}]',
    ),
    "переставить_столбцы": (
        "JSON: sheet (обязателен), columns[], position (_начало_|_конец_|_перед_|_после_), "
        "relative_column для _перед_/_после_; copy=true + new_names[] — копии; "
        "create_new=true + new_names[] — пустые столбцы по «Куда»; "
        "кнопка «Параметры…»",
        '[{"v":1,"fn":"переставить_столбцы","sheet":"Заказы","columns":["C","Сумма"],'
        '"position":"_перед_","relative_column":"Итого","copy":true,'
        '"new_names":["C_копия","Сумма_копия"]}]',
    ),
    "количество_значений": (
        "JSON: new_column, key_columns[], trim (по умолч. true), case_sensitive (по умолч. false)",
        '[{"v":1,"fn":"количество_значений","new_column":"Количество","key_columns":["ФИО","Отдел"]}]',
    ),
    "удалить_дубликаты": (
        "JSON: key_columns[], keep (first|last), pre_sort, output (inplace|new_sheet|offset), "
        "mark_duplicates — как dedup Excel по выбранным столбцам",
        '[{"v":1,"fn":"удалить_дубликаты","key_columns":["ФИО"],"keep":"first"}]',
    ),
    "копировать_переместить_лист": (
        "JSON: source_sheet + copy=false|true; copy=false (по умолч.) — переместить "
        "лист before/after anchor_sheet; copy=true — копия в dest_sheet",
        '[{"v":1,"fn":"копировать_переместить_лист","source_sheet":"Шаблон","copy":false,'
        '"position":"after","anchor_sheet":"Merge"}]',
    ),
    "копирование_диапазонов": (
        "JSON: source_sheet + source_range (A1 / B:B / 2:5) или source_rows/source_columns; "
        "dest_sheet(s) + dest_cell; mode=replace|insert; content=values|formulas; "
        "involve_dest (по умолч. true)",
        '[{"v":1,"fn":"копирование_диапазонов","source_sheet":"Сводная",'
        '"source_range":"A1:D20","dest_sheet":"Отчет","dest_cell":"B3","mode":"replace"}]',
    ),
    "создать_лист": (
        "JSON: name_mode=manual|expr + sheet_name/sheet_name_expr; columns; header_row; "
        "data_mode=none|sheet|rows; sheet→source_sheet+range; "
        "rows→строки через |, <<Переменные.список>> разворачивается в столбец",
        '[{"v":1,"fn":"создать_лист","name_mode":"manual","sheet_name":"Отчет",'
        '"columns":["Код","Сумма"],"header_row":1,"data_mode":"rows",'
        '"rows":"<<Переменные.коды>>, константа"}]',
    ),
    "отправить_по_почте": (
        "финал: вложение во временный файл + compose UI; attach whole_book|sheets|none; "
        "filename — имя+_ГГГГММДД_ЧЧММСС (пусто — имя книги); "
        "to/cc/bcc; to_source_sheet+to_source_range; client auto|system|outlook|thunderbird",
        '[{"v":1,"fn":"отправить_по_почте","attach":"whole_book","to":["boss@corp.local"],'
        '"filename":"Отчет","subject":"Отчет"}]',
    ),
    "сводная_таблица": (
        "JSON макета — кнопка «Параметры»; source_sheet — лист результата; "
        "в xml_excel_ods — статичный отчёт в файле (как as_values), не DataPilot",
        '[{"v":1,"fn":"сводная_таблица","sheet_name":"Сводная","row_fields":["Отдел"],'
        '"column_fields":[],"data_fields":[{"field":"Сумма","function":"SUM"}]}]',
    ),
    MERGE_PLUGIN_FUNCTION_KEY: (
        "C — JSON: ref/code; allow_env/allow_global опционально "
        "(пусто = MERGE_ALLOW_PLUGINS / Merge_Allow_Plugins); extra, sheet",
        '[{"v":1,"fn":"функция_плагин","ref":"functions_pp.py#pp_range_zagotovka"}]',
    ),
}
MERGE_POSTPROCESS_ROW_REST_HINTS = {
    "высота_строки": (
        "пусто — 5 мм на строку; «6» — 6 мм; «2,3 / 6» (или ->, =) — только эти строки (1-based); «0» — скрыть",
        "2,3,4 / 6",
    ),
    "серые_служебные": ("не используется", ""),
    "полосы_по_пути": ("не используется", ""),
    "чередующиеся_границы": ("не используется", ""),
    "копировать_формат_заголовка": ("не используется", ""),
    "подсветка_по_заголовку": ("не используется", ""),
    "цвет_текста_по_значению": (
        "подстрока заголовка столбца-источника",
        "статус",
    ),
    "жирный_по_пути": ("не используется (подстрока в коде макроса)", ""),
    "преобразовать_числа": ("колонки (опционально): индексы 0,2,5 или имена Сумма,Количество", ""),
    "удаление_листов": (
        "имена/шаблоны (*Отчет*); опционально filter=lambdaname→bool (санация); пусто — служебные по умолчанию; "
        "1=первый; _Первый_; _Последний_; _Последний_мусорный_",
        '[{"v":1,"fn":"удаление_листов","sheets":["*Отчет*"],"filter":"lambda name: sheet_age_days(name) is not None and sheet_age_days(name) <= 3"}]',
    ),
    "скрытие_листов": (
        "имена/шаблоны; опционально filter (как у удаления); пусто — служебные по умолчанию; "
        "1=первый; _Первый_; _Последний_; _Последний_мусорный_",
        '[{"v":1,"fn":"скрытие_листов","sheets":["Сбор_книг_лог*"]}]',
    ),
    MERGE_PLUGIN_FUNCTION_KEY: (
        "C — JSON: ref/code; allow_env/allow_global опционально "
        "(пусто = MERGE_ALLOW_PLUGINS / Merge_Allow_Plugins); extra, sheet",
        '[{"v":1,"fn":"функция_плагин","ref":"functions_pp.py#pp_row_zagotovka"}]',
    ),
}
MERGE_RESULT_POSTPROCESS_RANGE_MAP = {}
MERGE_RESULT_POSTPROCESS_XML_MAP = {}
_PP_XML_MAP_SPEC = (
    ("пропуск_пустых_строк", "merge_pp_range_skip_empty_rows"),
    ("удаление_строк", "merge_pp_range_skip_empty_rows"),
    ("сетка", "merge_pp_range_grid_borders"),
    ("тонкая_сетка", "merge_pp_range_thin_grid_borders"),
    ("толстая_сетка", "merge_pp_range_thick_grid_borders"),
    ("заголовок_плюс_высота", "merge_pp_range_header_row_height_pad"),
    ("перенос", "merge_pp_range_word_wrap"),
    ("перенос_и_авто_высота", "merge_pp_range_word_wrap_and_fit_rows"),
    ("авто_высота", "merge_pp_range_autofit_row_heights"),
    ("авто_ширина", "merge_pp_range_autofit_columns_width"),
    ("высота_строки", "merge_pp_range_row_height_pad"),
    ("левое_выравнивание", "merge_pp_range_left_align_data"),
    ("отступ", "merge_pp_range_increase_indent"),
    ("шрифт", "merge_pp_range_set_font"),
    ("зебра_диапазон", "merge_pp_range_zebra_even_rows"),
    ("формат_деньги", "merge_pp_range_format_money_columns"),
    ("формат_даты", "merge_pp_range_format_date_columns"),
    ("формат_столбцы", "merge_pp_range_format_columns"),
    ("применить_формулу", "merge_pp_range_apply_formula"),
    ("закрепить_заголовок", "merge_pp_range_freeze_header"),
    ("автофильтр", "merge_pp_range_add_filter"),
    ("удалить_столбцы", "merge_pp_range_delete_columns"),
    ("конкатенация_столбцов", "merge_pp_range_concat_columns"),
    ("разделить_по_столбцам", "merge_pp_range_split_by_columns"),
    ("условный_столбец", "merge_pp_range_conditional_column"),
    ("замена_значений", "merge_pp_range_replace_values"),
    ("текстовые_операции", "merge_pp_range_normalize_text"),
    ("заполнение_вниз", "merge_pp_range_fill_down_empty"),
    ("заполнить_вверх", "merge_pp_range_fill_up_empty"),
    ("fill_up", "merge_pp_range_fill_up_empty"),
    ("заполнение_вверх", "merge_pp_range_fill_up_empty"),
    ("заполнение_вниз_вычислить", "merge_pp_range_fill_down_calculate"),
    ("сортировка", "merge_pp_range_sort_data"),
    ("переименовать_лист", "merge_pp_range_rename_sheet"),
    ("переименовать_столбцы", "merge_pp_range_rename_columns"),
    ("переставить_столбцы", "merge_pp_range_reorder_columns"),
    ("копировать_переместить_лист", "merge_pp_range_copy_sheet"),
    ("сводная_таблица", "merge_pp_range_pivot_table"),
    ("градиент", "merge_pp_range_color_scale_simple"),
    ("ширина_столбцов", "merge_pp_range_set_column_widths"),
    ("раскрасить_блоки", "merge_pp_range_colorize_data"),
)
MERGE_POSTPROCESS_XML_REST_HINTS = {
    key: MERGE_POSTPROCESS_RANGE_REST_HINTS[key]
    for key in (
        "пропуск_пустых_строк",
        "удаление_строк",
        "сетка",
        "тонкая_сетка",
        "толстая_сетка",
        "заголовок_плюс_высота",
        "перенос",
        "перенос_и_авто_высота",
        "авто_высота",
        "авто_ширина",
        "высота_строки",
        "левое_выравнивание",
        "отступ",
        "шрифт",
        "зебра_диапазон",
        "формат_деньги",
        "формат_даты",
        "формат_столбцы",
        "применить_формулу",
        "закрепить_заголовок",
        "автофильтр",
        "удалить_столбцы",
        "конкатенация_столбцов",
        "разделить_по_столбцам",
        "условный_столбец",
        "замена_значений",
        "текстовые_операции",
        "заполнение_вниз",
        "заполнить_вверх",
        "развернуть_столбцы",
        "транспонировать_таблицу",
        "анкета_в_таблицу",
        "таблица_в_анкету",
        "заполнение_вниз_вычислить",
        "сортировка",
        "переименовать_лист",
        "переименовать_столбцы",
        "переставить_столбцы",
        "копировать_переместить_лист",
        "сводная_таблица",
        "удаление_листов",
        "скрытие_листов",
        "градиент",
        "ширина_столбцов",
        "раскрасить_блоки",
    )
    if key in MERGE_POSTPROCESS_RANGE_REST_HINTS
}
# В файле (прямой доступ) as_values для формул кривой — всегда игнор, даже из JSON.
MERGE_POSTPROCESS_XML_REST_HINTS["применить_формулу"] = (
    "имя_столбца=формула Calc; as_values в Постобработка_xml запрещён "
    "(формулы остаются в файле; галка в визарде недоступна)",
    '[{"v":1,"fn":"применить_формулу","column":"Итого","formula":"=E{row_id}+F{row_id}"}]',
)
_PP_RANGE_MAP_SPEC = (
    ("сетка", "merge_pp_range_grid_borders"),
    ("тонкая_сетка", "merge_pp_range_thin_grid_borders"),
    ("толстая_сетка", "merge_pp_range_thick_grid_borders"),
    ("заголовок_плюс_высота", "merge_pp_range_header_row_height_pad"),
    ("перенос", "merge_pp_range_word_wrap"),
    ("перенос_и_авто_высота", "merge_pp_range_word_wrap_and_fit_rows"),
    ("авто_высота", "merge_pp_range_autofit_row_heights"),
    ("авто_ширина", "merge_pp_range_autofit_columns_width"),
    ("высота_строки", "merge_pp_range_row_height_pad"),
    ("левое_выравнивание", "merge_pp_range_left_align_data"),
    ("отступ", "merge_pp_range_increase_indent"),
    ("шрифт", "merge_pp_range_set_font"),
    ("зебра_диапазон", "merge_pp_range_zebra_even_rows"),
    ("формат_деньги", "merge_pp_range_format_money_columns"),
    ("формат_даты", "merge_pp_range_format_date_columns"),
    ("формат_столбцы", "merge_pp_range_format_columns"),
    ("подсветка_по_порогу", "merge_pp_range_highlight_by_threshold"),
    ("условное_форматирование", "merge_pp_range_highlight_by_threshold"),
    ("подкрасить_пороги", "merge_pp_range_highlight_by_threshold"),
    ("градиент", "merge_pp_range_color_scale_simple"),
    ("удалить_столбцы", "merge_pp_range_delete_columns"),
    ("конкатенация_столбцов", "merge_pp_range_concat_columns"),
    ("разделить_по_столбцам", "merge_pp_range_split_by_columns"),
    ("условный_столбец", "merge_pp_range_conditional_column"),
    ("столбец_по_лямбде", "merge_pp_range_lambda_column"),
    ("добавить_столбец", "merge_pp_range_add_column"),
    ("применить_формулу", "merge_pp_range_apply_formula"),
    ("удаление_строк", "merge_pp_range_skip_empty_rows"),
    ("пропуск_пустых_строк", "merge_pp_range_skip_empty_rows"),
    ("сброс_стрипов", "merge_pp_range_init_path_stripe_state"),
    ("группировка_по_столбцу", "merge_pp_range_group_by_column"),
    ("закрепить_заголовок", "merge_pp_range_freeze_header"),
    ("ширина_столбцов", "merge_pp_range_set_column_widths"),
    ("автофильтр", "merge_pp_range_add_filter"),
    ("стиль_печати", "merge_pp_range_set_page_style"),
    ("заполнение_вниз", "merge_pp_range_fill_down_empty"),
    ("заполнить_вверх", "merge_pp_range_fill_up_empty"),
    ("fill_up", "merge_pp_range_fill_up_empty"),
    ("заполнение_вверх", "merge_pp_range_fill_up_empty"),
    ("развернуть_столбцы", "merge_pp_range_unpivot_columns"),
    ("unpivot", "merge_pp_range_unpivot_columns"),
    ("unpivot_columns", "merge_pp_range_unpivot_columns"),
    ("транспонировать_таблицу", "merge_pp_range_transpose_table"),
    ("transpose", "merge_pp_range_transpose_table"),
    ("transpose_table", "merge_pp_range_transpose_table"),
    ("анкета_в_таблицу", "merge_pp_range_form_to_table"),
    ("form_to_table", "merge_pp_range_form_to_table"),
    ("таблица_в_анкету", "merge_pp_range_table_to_form"),
    ("table_to_form", "merge_pp_range_table_to_form"),
    ("заполнение_вниз_вычислить", "merge_pp_range_fill_down_calculate"),
    ("копировать_значения", "merge_pp_range_copy_values"),
    ("замена_значений", "merge_pp_range_replace_values"),
    ("текстовые_операции", "merge_pp_range_normalize_text"),
    ("переименовать_лист", "merge_pp_range_rename_sheet"),
    ("переименовать_столбцы", "merge_pp_range_rename_columns"),
    ("переставить_столбцы", "merge_pp_range_reorder_columns"),
    ("количество_значений", "merge_pp_range_value_count"),
    ("удалить_дубликаты", "merge_pp_range_remove_duplicates"),
    ("копировать_переместить_лист", "merge_pp_range_copy_sheet"),
    ("копирование_диапазонов", "merge_pp_range_copy_ranges"),
    ("копировать_диапазон", "merge_pp_range_copy_ranges"),
    ("copy_ranges", "merge_pp_range_copy_ranges"),
    ("copy_range", "merge_pp_range_copy_ranges"),
    ("вставить_диапазон", "merge_pp_range_copy_ranges"),
    ("создать_лист", "merge_pp_range_create_sheet"),
    ("сводная_таблица", "merge_pp_range_pivot_table"),
    ("удаление_листов", "merge_pp_range_delete_sheets"),
    ("скрытие_листов", "merge_pp_range_hide_sheets"),
    ("сортировка", "merge_pp_range_sort_data"),
    ("раскрасить_блоки", "merge_pp_range_colorize_data"),
)
_PP_ROW_MAP_SPEC = (
    ("высота_строки", "merge_pp_row_height_pad"),
    ("серые_служебные", "merge_pp_row_gray_service_columns"),
    ("полосы_по_пути", "merge_pp_row_stripe_on_path_change"),
    ("чередующиеся_границы", "merge_pp_row_alternate_borders"),
    ("копировать_формат_заголовка", "merge_pp_row_copy_format_from_header"),
    ("подсветка_по_заголовку", "merge_pp_row_highlight_by_header_value"),
    ("цвет_текста_по_значению", "merge_pp_row_text_color_by_value"),
    ("жирный_по_пути", "merge_pp_row_bold_if_path_contains"),
    ("преобразовать_числа", "merge_pp_row_convert_numeric_strings"),
)
MERGE_RESULT_POSTPROCESS_ROW_MAP = {}
MERGE_FINAL_PROCESSING_REST_HINTS = {
    "удаление_листов": (
        "имена/шаблоны (*Отчет*); опционально filter=lambda name:… (санация); пусто — "
        "Справка_макроса, Сбор_книг_лог, Индексы_Источников; "
        "1=первый; _Первый_; _Последний_; _Последний_мусорный_",
        '[{"v":1,"fn":"удаление_листов","sheets":["*Отчет*"],'
        '"filter":"lambda name: sheet_age_days(name) is not None and sheet_age_days(name) <= 3"}]',
    ),
    "Удаление_листов": (
        "алиас «удаление_листов»",
        "",
    ),
    "скрытие_листов": (
        "имена/шаблоны; опционально filter (как у удаления); пусто — "
        "Справка_макроса, Сбор_книг_лог, Индексы_Источников; "
        "1=первый; _Первый_; _Последний_; _Последний_мусорный_",
        '[{"v":1,"fn":"скрытие_листов","sheets":["Сбор_книг_лог*"]}]',
    ),
    "Скрытие_листов": (
        "алиас «скрытие_листов»",
        "",
    ),
    "удалить_служебные_листы": (
        "устарело — перенесите в «Финальная_обработка» как Удаление_листов",
        "",
    ),
    "удалить_листы": (
        "устарело — перенесите в «Финальная_обработка» как Удаление_листов",
        "1_Справочник_цен",
    ),
    "Удалить_листы": (
        "алиас «удалить_листы»",
        "",
    ),
    "только_значения": (
        "имена листов через запятую или ; — преобразовать в значения; "
        "на листе со сводной (DataPilot) — inplace без нового листа",
        "Сводная,1_Заказы",
    ),
    "Только_значения": (
        "алиас «только_значения»",
        "Сводная",
    ),
    "сводная_таблица": (
        "как в постобработке: строит DataPilot на новом листе; "
        "JSON — кнопка «Параметры»; as_values=true — материализовать перед следующими шагами.",
        '[{"v":1,"fn":"сводная_таблица","source_sheet":"Merge","header_row":"","data_start":"","data_end":"","sheet_name":"","as_values":true,"list_rows_expand":true,"row_fields":["Регион","Отдел"],"column_fields":[],"filter_fields":[],"data_fields":[{"field":"Табельный_номер","function":"LIST_ROWS","list_fields":["Табельный_номер","Должность","Филиал","Комментарий"]}]}]',
    ),
    "заголовок_плюс_высота": (
        "как в постобработке: height_mm абсолют (мин. 7 мм), add_height=true — плюсовать; пусто — 12,7 мм",
        '[{"v":1,"fn":"заголовок_плюс_высота","height_mm":13,"add_height":false,"sheet":"Сводная"}]',
    ),
    "вертикаль_центр": (
        "вертикальное выравнивание по центру на строках данных (заголовок не трогается); "
        "C — список листов (имена через запятую или ;); пусто — пропуск",
        "Сводная_1",
    ),
    "зебра_диапазон": (
        "имя листа (или несколько через , или ;) | роль:заливка/шрифт — как у постобработки «зебра_диапазон»",
        "Итог|header:excel_header/белый; even:голубой/синий; odd:белый/синий",
    ),
    "сортировка": (
        "финал: «лист | ключи» или «лист: ключи»; несколько листов — «Л1,Л2|…»; "
        "задания через ; (как у зебра_диапазон, но вместо цветов — ключи сортировки)",
        "Сводная | Сумма -> -, Дата -> + ; Orders | Сумма -> -",
    ),
    "раскрасить_блоки": (
        "финал: «лист | столбцы_ключа -> цвета»; несколько листов — «Л1,Л2|…»; "
        "блоки через ; или ,. Цвета: имя или #hex; несколько — через %. "
        "Сортировка: /+ или /asc — по возрастанию, /- или /убыв — по убыванию; без суффикса — без сортировки; "
        "!граница — контур блоков",
        "Сводная | Отдел -> голубой%белый/+!граница ; Orders | A, B -> #E2EFDA/-",
    ),
    "раскрасить": (
        "алиас «раскрасить_блоки»",
        "Сводная | Отдел -> голубой",
    ),
    "сетка": (
        "как в постобработке: «толщина/цвет» (все листы) или «лист | 25/#CCCCCC ; …»; "
        "пусто — тонкая серая на всех",
        "Сводная | 25/#CCCCCC ; Orders : толстая/темныйсиний",
    ),
    "тонкая_сетка": _MERGE_PP_GRID_PRESET_REST_HINT,
    "толстая_сетка": _MERGE_PP_GRID_PRESET_REST_HINT,
    "перенос": MERGE_POSTPROCESS_RANGE_REST_HINTS["перенос"],
    "перенос_и_авто_высота": MERGE_POSTPROCESS_RANGE_REST_HINTS[
        "перенос_и_авто_высота"
    ],
    "авто_высота": MERGE_POSTPROCESS_RANGE_REST_HINTS["авто_высота"],
    "авто_ширина": MERGE_POSTPROCESS_RANGE_REST_HINTS["авто_ширина"],
    "левое_выравнивание": MERGE_POSTPROCESS_RANGE_REST_HINTS["левое_выравнивание"],
    "сброс_стрипов": MERGE_POSTPROCESS_RANGE_REST_HINTS["сброс_стрипов"],
    "закрепить_заголовок": MERGE_POSTPROCESS_RANGE_REST_HINTS["закрепить_заголовок"],
    "автофильтр": MERGE_POSTPROCESS_RANGE_REST_HINTS["автофильтр"],
    "высота_строки": MERGE_POSTPROCESS_RANGE_REST_HINTS["высота_строки"],
    "отступ": MERGE_POSTPROCESS_RANGE_REST_HINTS["отступ"],
    "шрифт": MERGE_POSTPROCESS_RANGE_REST_HINTS["шрифт"],
    "ширина_столбцов": MERGE_POSTPROCESS_RANGE_REST_HINTS["ширина_столбцов"],
    "формат_деньги": MERGE_POSTPROCESS_RANGE_REST_HINTS["формат_деньги"],
    "формат_даты": MERGE_POSTPROCESS_RANGE_REST_HINTS["формат_даты"],
    "формат_столбцы": MERGE_POSTPROCESS_RANGE_REST_HINTS["формат_столбцы"],
    "подсветка_по_порогу": MERGE_POSTPROCESS_RANGE_REST_HINTS["подсветка_по_порогу"],
    "условное_форматирование": (
        "алиас «подсветка_по_порогу»",
        MERGE_POSTPROCESS_RANGE_REST_HINTS["подсветка_по_порогу"][1],
    ),
    "подкрасить_пороги": (
        "алиас «подсветка_по_порогу»",
        MERGE_POSTPROCESS_RANGE_REST_HINTS["подсветка_по_порогу"][1],
    ),
    "градиент": MERGE_POSTPROCESS_RANGE_REST_HINTS["градиент"],
    "удалить_столбцы": MERGE_POSTPROCESS_RANGE_REST_HINTS["удалить_столбцы"],
    "конкатенация_столбцов": MERGE_POSTPROCESS_RANGE_REST_HINTS["конкатенация_столбцов"],
    "разделить_по_столбцам": MERGE_POSTPROCESS_RANGE_REST_HINTS["разделить_по_столбцам"],
    "условный_столбец": MERGE_POSTPROCESS_RANGE_REST_HINTS["условный_столбец"],
    "применить_формулу": MERGE_POSTPROCESS_RANGE_REST_HINTS["применить_формулу"],
    "удаление_строк": MERGE_POSTPROCESS_RANGE_REST_HINTS["удаление_строк"],
    "пропуск_пустых_строк": (
        "алиас «удаление_строк»",
        MERGE_POSTPROCESS_RANGE_REST_HINTS["удаление_строк"][1],
    ),
    "группировка_по_столбцу": MERGE_POSTPROCESS_RANGE_REST_HINTS["группировка_по_столбцу"],
    "стиль_печати": MERGE_POSTPROCESS_RANGE_REST_HINTS["стиль_печати"],
    "заполнение_вниз": MERGE_POSTPROCESS_RANGE_REST_HINTS["заполнение_вниз"],
    "заполнить_вверх": MERGE_POSTPROCESS_RANGE_REST_HINTS["заполнить_вверх"],
    "развернуть_столбцы": MERGE_POSTPROCESS_RANGE_REST_HINTS["развернуть_столбцы"],
    "транспонировать_таблицу": MERGE_POSTPROCESS_RANGE_REST_HINTS["транспонировать_таблицу"],
    "анкета_в_таблицу": MERGE_POSTPROCESS_RANGE_REST_HINTS["анкета_в_таблицу"],
    "таблица_в_анкету": MERGE_POSTPROCESS_RANGE_REST_HINTS["таблица_в_анкету"],
    "заполнение_вниз_вычислить": MERGE_POSTPROCESS_RANGE_REST_HINTS[
        "заполнение_вниз_вычислить"
    ],
    "копировать_значения": MERGE_POSTPROCESS_RANGE_REST_HINTS["копировать_значения"],
    "замена_значений": MERGE_POSTPROCESS_RANGE_REST_HINTS["замена_значений"],
    "текстовые_операции": MERGE_POSTPROCESS_RANGE_REST_HINTS["текстовые_операции"],
    "переименовать_лист": MERGE_POSTPROCESS_RANGE_REST_HINTS["переименовать_лист"],
    "переименовать_столбцы": MERGE_POSTPROCESS_RANGE_REST_HINTS["переименовать_столбцы"],
    "переставить_столбцы": (
        "JSON: sheet (обязателен), columns[], position (_начало_|_конец_|_перед_|_после_), "
        "relative_column для _перед_/_после_; copy=true + new_names[] — копии; "
        "create_new=true + new_names[] — пустые столбцы по «Куда»; "
        "кнопка «Параметры…»",
        '[{"v":1,"fn":"переставить_столбцы","sheet":"Заказы","columns":["C","Сумма"],'
        '"position":"_перед_","relative_column":"Итого","copy":true,'
        '"new_names":["C_копия","Сумма_копия"]}]',
    ),
    "количество_значений": MERGE_POSTPROCESS_RANGE_REST_HINTS["количество_значений"],
    "удалить_дубликаты": MERGE_POSTPROCESS_RANGE_REST_HINTS["удалить_дубликаты"],
    "копировать_переместить_лист": MERGE_POSTPROCESS_RANGE_REST_HINTS["копировать_переместить_лист"],
    "копирование_диапазонов": MERGE_POSTPROCESS_RANGE_REST_HINTS["копирование_диапазонов"],
    "копировать_диапазон": (
        "алиас «копирование_диапазонов»",
        MERGE_POSTPROCESS_RANGE_REST_HINTS["копирование_диапазонов"][1],
    ),
    "создать_лист": MERGE_POSTPROCESS_RANGE_REST_HINTS["создать_лист"],
    "отправить_по_почте": MERGE_POSTPROCESS_RANGE_REST_HINTS["отправить_по_почте"],
    "email": (
        "алиас «отправить_по_почте»",
        MERGE_POSTPROCESS_RANGE_REST_HINTS["отправить_по_почте"][1],
    ),
    "почта": (
        "алиас «отправить_по_почте»",
        MERGE_POSTPROCESS_RANGE_REST_HINTS["отправить_по_почте"][1],
    ),
    "активировать_лист": (
        "JSON: target_sheet — сделать лист активным; скрытый показывается; курсор в A1; "
        "tab_color — цвет ярлычка (имя/#hex; нет/сброс — умолчание)",
        '[{"v":1,"fn":"активировать_лист","target_sheet":"Merge","tab_color":"голубой"}]',
    ),
    "копировать_лист": (
        "алиас «копировать_переместить_лист»",
        MERGE_POSTPROCESS_RANGE_REST_HINTS["копировать_переместить_лист"][1],
    ),
    MERGE_PLUGIN_FUNCTION_KEY: (
        "C — JSON: ref/code; allow_env/allow_global опционально "
        "(пусто = MERGE_ALLOW_PLUGINS / Merge_Allow_Plugins); extra, sheet",
        '[{"v":1,"fn":"функция_плагин","ref":"functions_final.py#pp_range_zagotovka"}]',
    ),
}
_FINAL_MAP_SPEC = (
    ("удаление_листов", "merge_final_delete_sheets"),
    ("скрытие_листов", "merge_final_hide_sheets"),
    ("только_значения", "merge_final_values_only"),
    ("заголовок_плюс_высота", "merge_final_header_plus_height"),
    ("вертикаль_центр", "merge_final_vert_center"),
    ("зебра_диапазон", "merge_final_zebra"),
    ("сортировка", "merge_final_sort"),
    ("раскрасить_блоки", "merge_final_colorize"),
    ("сетка", "merge_final_grid"),
    ("тонкая_сетка", "merge_final_thin_grid"),
    ("толстая_сетка", "merge_final_thick_grid"),
    ("удалить_служебные_листы", "merge_final_delete_sheets"),
    ("удалить_листы", "merge_final_delete_sheets"),
    ("перенос", "merge_final_word_wrap"),
    ("перенос_и_авто_высота", "merge_final_word_wrap_and_fit"),
    ("авто_высота", "merge_final_autofit_rows"),
    ("авто_ширина", "merge_final_autofit_columns"),
    ("левое_выравнивание", "merge_final_left_align"),
    ("сброс_стрипов", "merge_final_reset_path_stripes"),
    ("закрепить_заголовок", "merge_final_freeze_header"),
    ("автофильтр", "merge_final_autofilter"),
    ("высота_строки", "merge_final_row_height"),
    ("отступ", "merge_final_indent"),
    ("шрифт", "merge_final_set_font"),
    ("ширина_столбцов", "merge_final_column_width"),
    ("формат_деньги", "merge_final_format_money"),
    ("формат_даты", "merge_final_format_date"),
    ("формат_столбцы", "merge_final_format_columns"),
    ("подсветка_по_порогу", "merge_final_highlight_threshold"),
    ("условное_форматирование", "merge_final_highlight_threshold"),
    ("подкрасить_пороги", "merge_final_highlight_threshold"),
    ("градиент", "merge_final_color_scale"),
    ("удалить_столбцы", "merge_final_delete_columns"),
    ("конкатенация_столбцов", "merge_final_concat_columns"),
    ("разделить_по_столбцам", "merge_final_split_by_columns"),
    ("условный_столбец", "merge_final_conditional_column"),
    ("столбец_по_лямбде", "merge_final_lambda_column"),
    ("добавить_столбец", "merge_final_add_column"),
    ("применить_формулу", "merge_final_apply_formula"),
    ("удаление_строк", "merge_final_delete_rows"),
    ("пропуск_пустых_строк", "merge_final_delete_rows"),
    ("группировка_по_столбцу", "merge_final_group_by_column"),
    ("стиль_печати", "merge_final_print_style"),
    ("заполнение_вниз", "merge_final_fill_down"),
    ("заполнить_вверх", "merge_final_fill_up"),
    ("fill_up", "merge_final_fill_up"),
    ("заполнение_вверх", "merge_final_fill_up"),
    ("развернуть_столбцы", "merge_final_unpivot_columns"),
    ("unpivot", "merge_final_unpivot_columns"),
    ("unpivot_columns", "merge_final_unpivot_columns"),
    ("транспонировать_таблицу", "merge_final_transpose_table"),
    ("transpose", "merge_final_transpose_table"),
    ("transpose_table", "merge_final_transpose_table"),
    ("анкета_в_таблицу", "merge_final_form_to_table"),
    ("form_to_table", "merge_final_form_to_table"),
    ("таблица_в_анкету", "merge_final_table_to_form"),
    ("table_to_form", "merge_final_table_to_form"),
    ("заполнение_вниз_вычислить", "merge_final_fill_down_calculate"),
    ("копировать_значения", "merge_final_copy_values"),
    ("замена_значений", "merge_final_replace_values"),
    ("текстовые_операции", "merge_final_normalize_text"),
    ("переименовать_лист", "merge_final_rename_sheet"),
    ("переименовать_столбцы", "merge_final_rename_columns"),
    ("переставить_столбцы", "merge_final_reorder_columns"),
    ("количество_значений", "merge_final_value_count"),
    ("удалить_дубликаты", "merge_final_remove_duplicates"),
    ("копировать_переместить_лист", "merge_final_copy_sheet"),
    ("активировать_лист", "merge_final_activate_sheet"),
    ("объединить_листы_в_один", "merge_final_merge_sheets_into_one"),
    ("копирование_диапазонов", "merge_final_copy_ranges"),
    ("копировать_диапазон", "merge_final_copy_ranges"),
    ("copy_ranges", "merge_final_copy_ranges"),
    ("copy_range", "merge_final_copy_ranges"),
    ("вставить_диапазон", "merge_final_copy_ranges"),
    ("создать_лист", "merge_final_create_sheet"),
    ("отправить_по_почте", "merge_final_send_mail"),
    ("email", "merge_final_send_mail"),
    ("mailto", "merge_final_send_mail"),
    ("send_mail", "merge_final_send_mail"),
    ("почта", "merge_final_send_mail"),
    ("сводная_таблица", "merge_final_pivot_table"),
)
MERGE_FINAL_PROCESSING_MAP = {}
_MERGE_CODE_POSTPROCESS_ROW = ()
_MERGE_CODE_POSTPROCESS_RANGE = ()
