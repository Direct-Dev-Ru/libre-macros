# -*- coding: utf-8 -*-
"""Константы макросов управления строками таблицы задач (todo_task_*)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.718"
# Имена колонок заголовка (лист задач). Сопоставление — нормализованное.
COL_NUM = u"№ п/п"
COL_NAME = u"Наименование"
COL_DUE = u"Срок исполнения"
COL_PRIORITY = u"Матрица приоритетов"
COL_SED = u"Реквизиты документа СЭД (при наличии)"
COL_ASSIGNEE = u"Исполнитель(и)"
COL_STATUS = u"Статус исполнения задачи"
COL_COMMENT = u"Комментарии"
COL_GUID = u"GUID Строки"
COL_CREATED = u"Дата и время создания задачи"
COL_MODIFIED = u"Дата и Время Редактирования"
# Автор создания / последнего изменения (ФИО владельца или пользователь ОС).
COL_CREATED_BY = u"Создано"
COL_MODIFIED_BY = u"Отредактировано"
# На своде задач (и при необходимости на листах задач).
COL_CATEGORY = u"Категория"

# Архив удалённых задач (скрытый лист).
DELETED_SHEET_NAME = u"__Удаленные_Задачи"
DELETED_COL_DELETED_AT = u"Дата и время удаления"
DELETED_COL_SOURCE_SHEET = u"Лист источника"

# Защита листа задач: снять перед правками, восстановить / включить после.
# Листы с именем на «Задачи» — защита всегда ставится в конце макроса.
# Пароль по умолчанию; пустой в настройках = защита без пароля.
TASK_SHEET_PROTECT_PASSWORD = u"task"
TASK_SHEET_NAME_PREFIX = u"Задачи"

# При копировании: не переносить статус «Выполнено» (без учёта регистра).
COPY_CLEAR_DONE_STATUS = u"выполнено"

# После автоподбора высоты строк — добавить к высоте (мм).
AUTOFIT_ROW_HEIGHT_EXTRA_MM = 4

# Раскраска просрочки (todo_task_colorize) — только в памяти, без столбцов/формул.
# Коды: 100 выполнено, 10 просрочено, 1 скоро (<3 дней), 0 обычная строка.
COLORIZE_CODE_DONE = 100
COLORIZE_CODE_OVERDUE = 10
COLORIZE_CODE_SOON = 1
COLORIZE_CODE_NORMAL = 0

# Ячейка «1 — ВАЖНО и СРОЧНО» в колонке приоритетов (todo_task_colorize).
COLORIZE_PRIORITY_URGENT_KEY = u"1-важноисрочно"
COLORIZE_PRIORITY_URGENT_FILL = 0xFFFF00
COLORIZE_PRIORITY_URGENT_BORDER_COLOR = 0xFF0000
COLORIZE_PRIORITY_URGENT_BORDER_WIDTH = 25

HEADER_ROW = 0  # 0-based fallback
# Искать строку заголовков с 1-й по 20-ю включительно (0-based: 0..19).
HEADER_SCAN_MAX_ROW = 19

# Лист справочников: строка 1 — заголовки списков, данные со 2-й.
REF_SHEET_NAME = u"__Справочники_задачи"
REF_LIST_HEADER_ROW = 0  # 0-based
REF_LIST_DATA_START_ROW = 1  # 0-based = лист.строка 2
REF_PRIORITY_COL = 0  # A
REF_PRIORITY_ROWS = (1, 4)  # A2:A5 inclusive, 0-based
REF_ASSIGNEE_COL = 2  # C
REF_STATUS_COL = 5  # F — «Статусы исполнения»
REF_STATUS_TITLE = u"Статусы исполнения"

# Форматы: копируем с последней задачи, затем сбрасываем жирность/курсив/заливку.
FORMAT_COPY_PROPS = (
    "CellStyle",
    "NumberFormat",
    "CharFontName",
    "CharHeight",
    "CharColor",
    "CharColor2",
    "CharAutoColor",
    "HoriJustify",
    "VertJustify",
    "VertJustify2",
    "ParaIndent",
    "IsTextWrapped",
    "TopBorder",
    "BottomBorder",
    "LeftBorder",
    "RightBorder",
)

CHAR_WEIGHT_NORMAL = 100
CHAR_POSTURE_NONE = 0

# Диалог редактирования
EDIT_DIALOG_TITLE = u"Редактирование задачи"
NEW_TASK_DIALOG_TITLE = u"Новая задача"
BTN_OK_LABEL = u"~OK"
BTN_CANCEL_LABEL = u"O~тмена"
EDIT_DIALOG_WIDTH = 520
EDIT_MULTILINE_H = 72
EDIT_FIELD_H = 22
EDIT_LABEL_H = 14
EDIT_GAP = 6
EDIT_MARGIN = 10
EDIT_BTN_W = 90
EDIT_BTN_H = 22
DUE_DAY_W = 36
DUE_MONTH_W = 130
DUE_YEAR_W = 52
DUE_DATE_OFFSET_DAYS = 4

# Кастомные message / confirm (вместо системных MessageBox)
MSG_DIALOG_WIDTH = 420
MSG_DIALOG_MIN_H = 120
MSG_TEXT_H = 72
MSG_BTN_W = 90
MSG_BTN_H = 22

# Формат ячейки «Срок исполнения» после сохранения (пример: 13.08.2026 Чт).
DUE_CELL_NUMBER_FORMAT = u"DD.MM.YYYY DDD"

# Дата и время создания / редактирования задачи.
TIMESTAMP_CELL_NUMBER_FORMAT = u"DD.MM.YYYY HH:MM:SS"

# Диалог сортировки задач (todo_task_sort).
SORT_DIALOG_TITLE = u"Сортировка задач"
SORT_DIALOG_HINT = (
    u"Сортировка целых строк-задач (все столбцы листа, включая «Категория»).\n"
    u"После сортировки колонка «№ п/п» перенумеровывается 1…N."
)
SORT_MODE_ASSIGNEE = u"assignee"
SORT_MODE_DUE = u"due"
SORT_MODE_OVERDUE = u"overdue"
SORT_MODE_PRIORITY = u"priority"
SORT_MODE_CHOICES = (
    (SORT_MODE_ASSIGNEE, u"По исполнителю (А → Я)"),
    (SORT_MODE_DUE, u"По сроку исполнения"),
    (SORT_MODE_OVERDUE, u"По степени просрочки"),
    (SORT_MODE_PRIORITY, u"По степени важности"),
)
SORT_DUE_DIR_OLD_FIRST = u"old_first"
SORT_DUE_DIR_NEW_FIRST = u"new_first"
SORT_DUE_DIR_CHOICES = (
    (SORT_DUE_DIR_OLD_FIRST, u"Сначала старые"),
    (SORT_DUE_DIR_NEW_FIRST, u"Сначала новые"),
)

# Диалог раскраски задач (todo_task_colorize).
COLORIZE_DIALOG_TITLE = u"Раскраска задач"
COLORIZE_DIALOG_HINT = (
    u"Раскрасить строки-задачи по просрочке и важности.\n"
    u"По умолчанию — только активный лист (если на нём есть таблица задач)."
)
COLORIZE_ALL_SHEETS_LABEL = u"Применить ко всем листам по шаблону имени"
COLORIZE_PATTERN_LABEL = u"Шаблон имени листа (* и ?):"
COLORIZE_DEFAULT_SHEET_PATTERN = u"Задачи*"

# Общие настройки серии todo_task (todo_task_settings).
SETTINGS_DIALOG_TITLE = u"Настройки задач"
SETTINGS_DIALOG_HINT = (
    u"Параметры макросов задач для выбранной области.\n"
    u"Частные — эта книга; если записи ещё нет, подставляются глобальные.\n"
    u"Пустой пароль — защита листа без пароля. По умолчанию: task."
)
SETTINGS_SCOPE_LABEL = u"Область настроек:"
SETTINGS_SCOPE_PRIVATE = u"private"
SETTINGS_SCOPE_GLOBAL = u"global"
SETTINGS_SCOPE_CHOICES = (
    (SETTINGS_SCOPE_PRIVATE, u"Частные (эта книга)"),
    (SETTINGS_SCOPE_GLOBAL, u"Глобальные (по умолчанию)"),
)
SETTINGS_SCOPE_BOOK_LABEL = u"Книга:"
SETTINGS_SCOPE_UNSAVED_BOOK = u"Книга не сохранена на диск — частные настройки появятся после сохранения файла."
SETTINGS_SHOW_PWD_LABEL = u"Показать"
SETTINGS_SHOW_PWD_SECONDS = 10
SETTINGS_CONSOLE_LOG_LABEL = u"Подробные логи в консоль (этапы и время)"
DEFAULT_CONSOLE_LOG = True
SETTINGS_DIALOG_FONT_LABEL = u"Размер шрифта в диалогах (pt):"
DEFAULT_DIALOG_FONT_PT = 11
DIALOG_FONT_PT_MIN = 8
DIALOG_FONT_PT_MAX = 24
SETTINGS_TAB_GENERAL = u"Общие"
SETTINGS_TAB_LAYOUT = u"Размеры полей"
SETTINGS_APPLY_WIDTHS_LABEL = u"Применить"
SETTINGS_LAYOUT_HINT = (
    u"Ширина в мм (0 или меньше — скрыть). По умолчанию — только листы задач. "
    u"Галка — также листы свода по шаблону ниже."
)
SETTINGS_APPLY_TO_SUMMARY_LABEL = (
    u"Применять размеры и к листам свода (по шаблону)"
)
DEFAULT_APPLY_WIDTHS_TO_SUMMARY = False
SETTINGS_TASK_PATTERN_LABEL = u"Шаблон листов задач (* и ?):"
SETTINGS_SUMMARY_PATTERN_LABEL = u"Шаблон листов свода (* и ?):"
SETTINGS_TEMPLATE_SHEET_LABEL = u"Лист-шаблон заголовков:"
DEFAULT_LAYOUT_TEMPLATE_SHEET = u"Задачи.01"
DEFAULT_TASK_SHEET_PATTERN = u"Задачи*"
DEFAULT_SUMMARY_SHEET_PATTERN = u"Свод*задач*"
# Лист свода задач сотрудников в текущей книге (todo_task_merge).
DEFAULT_MERGE_SOURCE_SHEET = u"задачи_сотрудников"
SETTINGS_MERGE_SOURCE_LABEL = u"Лист с задачами сотрудников:"
DEFAULT_MERGE_UPDATE_EXISTING_ONLY = False
DEFAULT_MERGE_CHECK_MODIFIED = False
DEFAULT_MERGE_ACCUMULATE_COMMENTS = False
SETTINGS_MERGE_UPDATE_EXISTING_LABEL = (
    u"Слияние: обновлять только существующие"
)
SETTINGS_MERGE_CHECK_MODIFIED_LABEL = (
    u"Слияние: контролировать дату-время редактирования"
)
SETTINGS_MERGE_ACCUMULATE_COMMENTS_LABEL = (
    u"Слияние: накапливать комментарии"
)

# Справка по макросам todo_task (todo_task_help).
HELP_DIALOG_TITLE = u"Справка по задачам"
HELP_TEMP_FILE_NAME = u"todo_task_help_ru.odt"
HELP_OPEN_FAILED_MSG = u"Не удалось открыть справку:\n%s"

# Слияние с сводом (todo_task_merge).
MERGE_DIALOG_PROGRESS_TITLE = u"Слияние задач"
MERGE_CONFIRM_TITLE = u"Слияние задач сотрудников"
MERGE_CONFIRM_HINT = (
    u"Перед слиянием проверьте сводку и параметры.\n"
    u"Комментарии: объединять (новые сверху) или полностью заменить.\n"
    u"Копирование «только значения»: числа и даты — как значения (формат берётся с листа приёмника), "
    u"без заливки и обрамления источника."
)
MERGE_COMMENT_MODE_LABEL = u"Метод слияния задач:"
MERGE_COMMENT_MODE_ACCUMULATE = u"accumulate"
MERGE_COMMENT_MODE_REPLACE = u"replace"
MERGE_COMMENT_MODE_CHOICES = (
    (MERGE_COMMENT_MODE_ACCUMULATE, u"Объединять комментарии"),
    (MERGE_COMMENT_MODE_REPLACE, u"Полная замена"),
)
MERGE_COPY_MODE_LABEL = u"Копирование ячеек:"
MERGE_COPY_MODE_VALUES = u"values"
MERGE_COPY_MODE_FORMATS = u"formats"
MERGE_COPY_MODE_CHOICES = (
    (MERGE_COPY_MODE_VALUES, u"Только значения"),
    (MERGE_COPY_MODE_FORMATS, u"С форматами"),
)
DEFAULT_MERGE_RUN_ACCUMULATE_COMMENTS = True
DEFAULT_MERGE_RUN_VALUES_ONLY = True
MERGE_REPORT_TITLE = u"Результат слияния"
MERGE_REPORT_LIST_H = 160
MERGE_REPORT_NAME_PREVIEW = 48
MERGE_EDITOR_CONFLICT_TITLE = u"Конфликт редактирования"
MERGE_EDITOR_CONFLICT_HINT = (
    u"Задача с этим GUID уже есть в книге и отредактирована владельцем файла.\n"
    u"В своде сотрудника указан другой редактор («Отредактировано»).\n"
    u"Чья версия задачи имеет приоритет?"
)
MERGE_EDITOR_KEEP_OWNER = u"keep_owner"
MERGE_EDITOR_TAKE_EMPLOYEE = u"take_employee"
MERGE_EDITOR_BTN_KEEP = u"Оставить версию владельца"
MERGE_EDITOR_BTN_TAKE = u"Взять версию сотрудника"
MERGE_EDITOR_APPLY_ALL_LABEL = u"Для всех оставшихся таких конфликтов"
MERGE_EDITOR_CONFLICT_TEXT_H = 72
MERGE_ROLE_DENIED_MSG = (
    u"Слияние задач с сводом сотрудников доступно только при роли "
    u"владельца «руководитель» (настройки задач)."
)
MERGE_SOURCE_MISSING_MSG = u"В книге нет листа «%s» (свод задач сотрудников)."
MERGE_NO_TASK_SHEETS_MSG = u"Не найдены листы задач по шаблону «%s»."
MERGE_NO_CATEGORY_MSG = (
    u"На листе «%s» нет колонки «%s» — не удалось определить целевой лист."
)
MERGE_PLAN_PREVIEW_NO_SOURCES = (
    u"Планы: в книге нет листов __План_задач_… для переноса."
)
MERGE_PLAN_PREVIEW_FOUND = (
    u"Планы: листов в книге к переносу — %s; задач свода без листа — %s."
)
MERGE_PLAN_REPORT_HEADER = u"--- Перенос планов ---"
MERGE_PLAN_REPORT_NONE = u"Листы планов в книге не найдены (нечего переносить)."
MERGE_PLAN_REPORT_MOVED = u"Перенесено листов планов: %s"
MERGE_PLAN_REPORT_SKIPPED = u"Без плана у задач: %s"
MERGE_PLAN_REPORT_ERRORS = u"Ошибки переноса планов: %s"
EDIT_FOREIGN_CREATOR_WARN = (
    u"Задача создана пользователем «%s» (не совпадает с ФИО владельца в настройках).\n"
    u"План задачи остаётся только для просмотра.\n"
    u"При слиянии с сводом сотрудников правки задачи могут быть перезаписаны.\n\n"
    u"Продолжить редактирование?"
)
EDIT_UNLOCK_CONFIRM_TITLE = u"Чужая задача"
EDIT_READ_ONLY_TITLE_SUFFIX = u" (просмотр)"
EDIT_READ_ONLY_MERGE_HINT = (
    u"Лист свода сотрудников — задача и план только для просмотра."
)
EDIT_READ_ONLY_FOREIGN_HINT = (
    u"Задача перенесена от другого автора — открыта для просмотра.\n"
    u"Нажмите «Редактировать», чтобы изменить поля задачи (план остаётся только для просмотра)."
)
EDIT_UNLOCK_BTN = u"Редактировать"
EDIT_FOREIGN_PLAN_HINT = (
    u"План чужой задачи — только просмотр (сохранение отключено)."
)

# Справочники (todo_task_refs).
REFS_DIALOG_TITLE = u"Справочники задач"
REFS_DIALOG_HINT = (
    u"Выберите элемент в списке — он появится в поле внизу.\n"
    u"«Обновить» записывает правку, «Удалить» — выбранный элемент."
)
REFS_DICT_LABEL = u"Справочник:"
REFS_ITEMS_LABEL = u"Элементы:"
REFS_VALUE_LABEL = u"Значение (для нескольких полей — через | ):"
REFS_BTN_ADD = u"Добавить"
REFS_BTN_UPDATE = u"Обновить"
REFS_BTN_DELETE = u"Удалить"
REFS_BTN_DEDUP = u"Убрать дубликаты"

OWNER_ROLE_MANAGER = u"manager"
OWNER_ROLE_EMPLOYEE = u"employee"
OWNER_ROLE_CHOICES = (
    (OWNER_ROLE_MANAGER, u"руководитель"),
    (OWNER_ROLE_EMPLOYEE, u"сотрудник"),
)

# Диалог перемещения задач (todo_task_move).
MOVE_DIALOG_TITLE = u"Перемещение задач"
MOVE_DIALOG_HINT = (
    u"Переместить выбранные задачи на другой лист «Задачи…».\n"
    u"Колонки сопоставляются по заголовкам (без учёта регистра и пробелов по краям)."
)
MOVE_VALUES_ONLY_LABEL = u"Только значения (без заливки/обрамления; числа и даты — значением)"
MOVE_ADD_COLUMNS_LABEL = (
    u"Добавить на приёмник недостающие колонки (в конец листа)"
)
DEFAULT_MOVE_VALUES_ONLY = True
DEFAULT_MOVE_ADD_COLUMNS = False
MOVE_NAME_PREVIEW_LEN = 72
MOVE_TASKS_LIST_FONT_PT = None  # None — брать из general.dialog_font_pt
MOVE_SHEET_PATTERN = u"Задачи*"

# Конфликт GUID при перемещении/копировании на лист назначения.
MOVE_GUID_ACTION_OVERWRITE = u"overwrite"
MOVE_GUID_ACTION_MERGE_COMMENTS = u"merge_comments"
MOVE_GUID_ACTION_NEW_GUID = u"new_guid"
MOVE_GUID_CONFLICT_TITLE = u"Совпадение GUID"
MOVE_GUID_CONFLICT_HINT = (
    u"Такой GUID уже есть на одном из листов задач.\n"
    u"Если он не на листе назначения — при «Перезаписать» / «Слить комментарии» "
    u"сначала перенесём существующую задачу на приёмник, затем применим выбранное.\n"
    u"Перезаписать — заменить строку на приёмнике.\n"
    u"Слить комментарии — поля обновить, комментарии накопить блоками "
    u"(как при объединении задач сотрудников).\n"
    u"Новая задача — добавить строку с новым GUID (существующую не трогаем)."
)
MOVE_GUID_BTN_OVERWRITE = u"Перезаписать"
MOVE_GUID_BTN_MERGE = u"Слить комментарии"
MOVE_GUID_BTN_NEW = u"Новая задача"
MOVE_GUID_BTN_ABORT = u"Прервать"
MOVE_GUID_APPLY_ALL_LABEL = u"Для всех оставшихся совпадений"
MOVE_GUID_CONFLICT_TEXT_H = 120
MOVE_GUID_NAME_PREVIEW_LEN = 64

MONTH_COMBO_ITEMS = (
    u"01 — Январь",
    u"02 — Февраль",
    u"03 — Март",
    u"04 — Апрель",
    u"05 — Май",
    u"06 — Июнь",
    u"07 — Июль",
    u"08 — Август",
    u"09 — Сентябрь",
    u"10 — Октябрь",
    u"11 — Ноябрь",
    u"12 — Декабрь",
)
