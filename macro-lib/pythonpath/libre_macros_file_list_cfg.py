# -*- coding: utf-8 -*-
"""Константы и состояние file_list_macro (AlterOffice 2026: модульный код .py-скрипта вырезается)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.779"
DEFAULT_MAX_DEPTH = 1
DEFAULT_SIZE_UNIT = "b"
DEFAULT_COUNT_PAGES = False
DEFAULT_COUNT_SHA256 = False
# Макс. размер файла для SHA256 (Гб); больше — пропуск.
DEFAULT_SHA256_MAX_GB = 1.0
DEFAULT_SORT_MODE = "tree"
# Имя листа результата, если в диалоге поле пустое (не активный лист).
DEFAULT_RESULT_SHEET_NAME = "file_list_result"
DEPTH_INDENT = "    "
SETTINGS_APP_NAME = "libre-macros"
SETTINGS_MODULE_NAME = "file_list_macro"
SIZE_UNIT_OPTIONS = (
    ("b", "б"),
    ("kb", "Кб"),
    ("mb", "Мб"),
    ("gb", "Гб"),
)
SIZE_UNIT_DIVISOR = {
    "b": 1,
    "kb": 1024,
    "mb": 1024 ** 2,
    "gb": 1024 ** 3,
}
# Пресеты сортировки результата: (code, label).
SORT_PRESETS = (
    ("tree", "Дерево (относительный путь)"),
    ("depth_kind_name", "Глубина → тип → имя"),
    ("full_depth_name", "Полный путь → глубина → имя"),
    ("rel_depth_name", "Относительный путь → глубина → имя"),
    ("kind_name", "Тип → имя"),
    ("name", "Имя"),
    ("ext_name", "Расширение → имя"),
    ("parent_name", "Родитель → имя"),
    ("mtime_desc", "Дата изменения (новые сверху)"),
    ("size_desc", "Размер (больше сверху)"),
)
# PDF + Writer + Impress/Draw (Calc не считаем — «страницы печати» ≠ листы).
PAGE_COUNT_EXTS = frozenset(
    (
        ".pdf",
        ".odt",
        ".ott",
        ".fodt",
        ".doc",
        ".docx",
        ".dot",
        ".dotx",
        ".rtf",
        ".odp",
        ".otp",
        ".fodp",
        ".ppt",
        ".pptx",
        ".pot",
        ".potx",
        ".odg",
        ".otg",
        ".fodg",
    )
)
RESULT_HEADER_PAGES = "Страниц"
RESULT_HEADER_SHA256 = "SHA256"
RESULT_HEADER_DUP = "Дубль"
RESULT_HEADER_DUP_N = "№ дубля"
# Чтение файла для SHA256 порциями.
SHA256_READ_CHUNK = 1024 * 1024
# Windows: владелец NTFS (часто совпадает с последним редактором в доменных шарах).
RESULT_HEADER_OWNER = "Owner"
RESULT_HEADER_BASE = [
    "Тип",
    "Имя",
    "Относительный путь",
    "Родитель",
    "Глубина",
    None,
    "Дата изменения",
    "Расширение",
    "Полный путь",
]
# Оформление заголовка результата: excel_header / белый (как зебра в проекте).
RESULT_HEADER_STYLE = "excel_header/белый"
# После OptimalWidth — запас под иконку автофильтра (доля от текущей ширины).
COLUMN_WIDTH_AUTOFILTER_PAD = 0.15
UI_YIELD_EVERY = 10
WRITE_CHUNK = 300

# --- Поиск: анализ имён/путей и содержимого по словарю regex ---
DEFAULT_DLP_ANALYZE_NAMES = False
DEFAULT_DLP_ANALYZE_CONTENT = False
DEFAULT_DLP_CONTENT_TEXT = True
DEFAULT_DLP_CONTENT_WORD = True
DEFAULT_DLP_CONTENT_EXCEL = True
DLP_NAME_DICT_FILENAME = "search_name_patterns.json"
DLP_CONTENT_DICT_FILENAME = "search_content_patterns.json"
DLP_LOG_SHEET_SUFFIX = "_Поиск"
DLP_LOG_PATTERN_HEADER = "Паттерн"
DLP_LOG_SOURCE_HEADER = "Источник"
DLP_SOURCE_NAME = "имя"
DLP_SOURCE_CONTENT = "содержимое"
# Подсветка строки результата при совпадении (RGB Calc).
DLP_MATCH_ROW_COLOR = 0xFFCCCC
# Лимиты чтения содержимого (байты с диска / символы после decode).
DLP_CONTENT_MAX_BYTES = 2 * 1024 * 1024
DLP_CONTENT_MAX_CHARS = 500000
# Текстовые расширения (фаза A).
DLP_TEXT_EXTS = frozenset(
    (
        ".txt",
        ".csv",
        ".tsv",
        ".json",
        ".xml",
        ".log",
        ".md",
        ".html",
        ".htm",
        ".ini",
        ".cfg",
        ".yml",
        ".yaml",
        ".sql",
        ".py",
        ".js",
        ".css",
        ".properties",
        ".rtf",
    )
)
# Word / AText (+ шаблоны, legacy).
DLP_WORD_EXTS = frozenset(
    (
        ".doc",
        ".docx",
        ".docm",
        ".dot",
        ".dotx",
        ".dotm",
        ".odt",
        ".ott",
        ".fodt",
        ".rtf",
    )
)
# Excel / ACell (+ шаблоны, legacy).
DLP_EXCEL_EXTS = frozenset(
    (
        ".xls",
        ".xlsx",
        ".xlsm",
        ".xlsb",
        ".xlt",
        ".xltx",
        ".xltm",
        ".ods",
        ".ots",
        ".fods",
    )
)
# Словарь по умолчанию: id → regex (флаги можно встроить, напр. (?i)).
DEFAULT_DLP_NAME_PATTERNS = {
    "confidential": r"(?i)конфиденц",
    "trade_secret": r"(?i)коммерческ\w*[\s_\-]*тайн",
    "dsp": r"(?i)(?:^|[_\-\s./\\])дсп(?:[_\-\s./\\]|$)",
    "secret": r"(?i)секретн",
    "personal_data": r"(?i)персональн\w*[\s_\-]*данн",
    "restricted": r"(?i)для[\s_\-]*служебн",
    "internal_only": r"(?i)только[\s_\-]*для[\s_\-]*внутренн",
    "password_hint": r"(?i)(?:парол|password|passwd|pwd)[_\-\s.]*(?:list|файл|file|баз)?",
}
DEFAULT_DLP_CONTENT_PATTERNS = dict(DEFAULT_DLP_NAME_PATTERNS)

# Лист книги как словарь паттернов: одно слово в поле пути → имя листа.
# Если при генерации имя некорректно/пусто — это имя по умолчанию.
DEFAULT_DLP_PATTERN_SHEET_NAME = u"поиск_шаблоны"
# True (standalone ODS): поля словарей по умолчанию = имя листа, не JSON-файл.
DEFAULT_DLP_DICT_USE_SHEET = False
# Шрифт заголовка листа паттернов.
DLP_PATTERN_SHEET_FONT_NAME = u"PT Sans"
# Справка на сгенерированном листе: колонка E, со строки 2 (по одной фразе в ячейке).
DLP_PATTERN_SHEET_HELP_COL = 4  # E
DLP_PATTERN_SHEET_HELP_START_ROW = 1  # строка 2 (0-based)
DLP_PATTERN_SHEET_HELP_LINES = (
    u"Как заполнять словарь паттернов",
    u"",
    u"Колонки A–C: имя | паттерн | описание (описание только для вас).",
    u"Поиск смотрит колонку B (паттерн) по имени/пути или содержимому файла.",
    u"",
    u"Простое слово — достаточно куска:",
    u"  конфиденц  → найдёт «конфиденциально», «Конфиденц_…» и т.п. везде",
    u"  (без учёта регистра; не нужно писать * или (?i) вручную).",
    u"",
    u"Сложные случаи — regex Python (как есть):",
    u"  (?i)trade[\\s_\\-]*secret",
    u"  (?i)(?:^|[_\\-\\s./\\\\])дсп(?:[_\\-\\s./\\\\]|$)",
    u"Метасимволы . \\ ^ $ * + ? { } [ ] | ( ) — значит «это regex».",
    u"",
    u"Имя (A) — короткий id для лога *_Поиск; паттерн (B) — что искать.",
    u"Пустые строки и строки без A или B пропускаются.",
    u"Не называйте этот лист так же, как лист результата списка файлов.",
)
# Заголовки и строки по умолчанию — как test/file_list/dlp_name_patterns_example.csv
DLP_PATTERN_SHEET_HEADERS = (u"имя", u"паттерн", u"описание")
DEFAULT_DLP_PATTERN_SHEET_ROWS = (
    (u"confidential", u"конфиденц", u"простое слово → (?i)конфиденц (подфраза, без регистра)"),
    (u"confidential_en", u"confidential", u"EN: confidential"),
    (u"trade_secret", u"(?i)коммерческ\\w*[\\s_\\-]*тайн", u"regex: коммерческая тайна / коммерческой_тайны"),
    (u"trade_secret_en", u"(?i)trade[\\s_\\-]*secret", u"EN: trade secret"),
    (u"dsp", u"(?i)(?:^|[_\\-\\s./\\\\])дсп(?:[_\\-\\s./\\\\]|$)", u"ДСП как отдельный токен в имени/пути"),
    (u"secret", u"секретн", u"секретный / секретные / …"),
    (u"secret_en", u"secret", u"EN: секретный / секретные / …"),
    (u"secret_strict", u"(?i)(?:^|[_\\-\\s./\\\\])строго[\\s_\\-]*секретн", u"строго секретно"),
    (u"personal_data", u"(?i)персональн\\w*[\\s_\\-]*данн", u"персональные данные"),
    (u"personal_data_en", u"(?i)personal[\\s_\\-]*data|PII|GDPR", u"EN/аббревиатуры ПДн"),
    (u"restricted", u"для служебн", u"простое слово / начало фразы «для служебного…»"),
    (u"restricted_full", u"(?i)для[\\s_\\-]*служебн\\w*[\\s_\\-]*пользован", u"для служебного пользования"),
    (u"internal_only", u"только для внутренн", u"только для внутреннего…"),
    (u"internal_en", u"(?i)internal[\\s_\\-]*(?:use|only)|for[\\s_\\-]*internal", u"EN: internal use / for internal"),
    (u"password_hint", u"(?i)(?:парол|password|passwd|pwd)[_\\-\\s.]*(?:list|файл|file|баз)?", u"пароли / password_list"),
    (u"passwords_plain", u"парол", u"простое слово «парол» (пароль/паролей…)"),
    (u"nda", u"nda", u"NDA / nda в имени"),
    (u"гриф", u"гриф", u"гриф секретности в имени"),
    (u"ограниченного_распространения", u"ограниченн", u"ограниченного распространения"),
    (u"банковская_тайна", u"(?i)банковск\\w*[\\s_\\-]*тайн", u"банковская тайна"),
    (u"врачебная_тайна", u"(?i)врачебн\\w*[\\s_\\-]*тайн", u"врачебная тайна"),
    (u"ноу_хау", u"(?i)ноу[\\s_\\-]*хау|know[\\s_\\-]*how", u"ноу-хау / know-how"),
    (u"инсайд", u"инсайдер", u"инсайдерская информация"),
    (u"гостайна", u"(?i)гос(?:ударственн\\w*)?[\\s_\\-]*тайн", u"гостайна / государственная тайна"),
    (u"шифр", u"шифр", u"шифр / шифрование в имени"),
    (u"keystore", u"keystore|key[\\s_\\-]*store|\\.pem$|\\.p12$", u"ключи / pem / p12 (regex)"),
    (u"credentials", u"credential|secret[\\s_\\-]*key|api[\\s_\\-]*key", u"credentials / api key"),
    (u"backup_secrets", u"(?i)backup.*(парол|password|secret)|(парол|password|secret).*backup", u"бэкапы с секретами"),
    (u"пдн", u"пдн", u"аббревиатура ПДн"),
    (u"ппдн", u"ппдн", u"персональные / персональные производные"),
    (u"тк", u"тк[\\s_\\-]*рф|трудов\\w*[\\s_\\-]*кодекс", u"осторожно: короткое «тк» — лучше уточнять regex"),
    (u"договор_нда", u"(?i)договор[\\s_\\-]*(?:нда|nda)|nda[\\s_\\-]*agreement", u"договор NDA"),
    (u"черновик_секретно", u"(?i)черновик[\\s_\\-].*(секрет|конфиденц)|(секрет|конфиденц).*черновик", u"черновик + секрет/конфиденц"),
)

# Мутабельное состояние UI (живёт в pythonpath — не проходит через AO-фильтр скрипта)
status_indicator = None
yield_counter = 0
