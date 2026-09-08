# -*- coding: utf-8 -*-
"""Константы плана выполнения задачи (todo_task_plan_*)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.711"
PLAN_SHEET_NAME = u"__План_задач"
PLAN_SHEET_PREFIX = u"__План_задач_"
PLAN_SHEET_NAME_MAX = 31
PLAN_ARCHIVE_SHEET_NAME = u"__Архив_план"
# Внешнее хранилище планов (рядом с книгой задач): каталог + .xlsx.
PLAN_STORE_DIR_PREFIX = u"Plans_for_"
PLAN_STORE_FILE_PREFIX = u"Plans_for_"
PLAN_STORE_FILE_EXT = u".xlsx"
PLAN_STORE_PLACEHOLDER_SHEET = u"_index"
PLAN_STORE_DOC_UNSAVED = (
    u"Сохраните книгу задач на диск — планы хранятся в соседнем каталоге Plans_for_… ."
)
PLAN_STORE_IO_ERROR = u"Ошибка доступа к файлу планов:\n%s"

PLAN_HEADER_ROW = 0
PLAN_DATA_START_ROW = 1

PLAN_COL_GUID = u"GUID задачи"
PLAN_COL_ITEM = u"Пункт"
PLAN_COL_SUMMARY = u"Итоговый пункт"
PLAN_COL_START = u"Дата начала"
PLAN_COL_END = u"Дата окончания"
PLAN_COL_START_NO_SNAP = u"Не корр. начало"
PLAN_COL_END_NO_SNAP = u"Не корр. окончание"
PLAN_COL_DURATION = u"Длительность, раб. дни"
PLAN_COL_ASSIGNEES = u"Ответственные"
PLAN_COL_STAGE = u"Название этапа"
PLAN_COL_RESULT = u"Результат / отчётный документ"
PLAN_COL_CRITICALITY = u"Критичность"
PLAN_COL_DONE = u"Выполнено"
PLAN_COL_SORT = u"Порядок"
PLAN_COL_PARENT = u"Родитель"
PLAN_COL_ROW_ID = u"ID строки плана"
PLAN_COL_CREATED = u"Создано"
PLAN_COL_MODIFIED = u"Изменено"

PLAN_ARCHIVE_COL_AT = u"Дата и время архивации"
PLAN_ARCHIVE_COL_SHEET = u"Лист источника"
PLAN_ARCHIVE_COL_TASK = u"Наименование задачи"

PLAN_COLUMN_SPECS = (
    (u"guid", PLAN_COL_GUID),
    (u"item_no", PLAN_COL_ITEM),
    (u"is_summary", PLAN_COL_SUMMARY),
    (u"date_start", PLAN_COL_START),
    (u"date_end", PLAN_COL_END),
    (u"start_no_snap", PLAN_COL_START_NO_SNAP),
    (u"end_no_snap", PLAN_COL_END_NO_SNAP),
    (u"duration_wd", PLAN_COL_DURATION),
    (u"assignees", PLAN_COL_ASSIGNEES),
    (u"stage_name", PLAN_COL_STAGE),
    (u"result_doc", PLAN_COL_RESULT),
    (u"criticality", PLAN_COL_CRITICALITY),
    (u"is_done", PLAN_COL_DONE),
    (u"sort_key", PLAN_COL_SORT),
    (u"parent_item", PLAN_COL_PARENT),
    (u"row_id", PLAN_COL_ROW_ID),
    (u"created_at", PLAN_COL_CREATED),
    (u"modified_at", PLAN_COL_MODIFIED),
)

PLAN_SUMMARY_YES = u"Да"
PLAN_SUMMARY_VALUES = (PLAN_SUMMARY_YES, u"")

REF_HOLIDAYS_TITLE = u"Праздники"
REF_CRITICALITY_TITLE = u"Критичность плана"
REF_CRITICALITY_DEFAULTS = (u"Низкая", u"Средняя", u"Высокая")

PLAN_BTN_CREATE = u"Создать план"
PLAN_BTN_EDIT = u"Редактировать план"
PLAN_BTN_VIEW = u"Просмотр плана"
PLAN_BTN_NO_PLAN = u"План отсутствует"
PLAN_BTN_DELETE = u"Удалить план"
PLAN_BTN_EXPORT = u"Выгрузить план"
PLAN_WIZARD_VIEW_SUFFIX = u" (просмотр)"
PLAN_WIZARD_VIEW_HINT = (
    u"Чужой план — режим просмотра: данные можно редактировать и копировать, "
    u"но «Сохранить» отключено — изменения не будут записаны."
)
PLAN_WIZARD_CLOSE_BTN = u"Закрыть"

PLAN_DELETE_CONFIRM = (
    u"Удалить план задачи «%s»?\n"
    u"Строки плана будут перенесены на лист «%s» в файле планов, затем удалены."
)
PLAN_DELETE_DONE = u"План удалён."
PLAN_DELETE_NO_PLAN = u"У задачи нет сохранённого плана."
PLAN_EXPORT_NO_PLAN = u"У задачи нет сохранённого плана для выгрузки."
PLAN_EXPORT_TITLE = u"Сохранить план как Excel"
PLAN_EXPORT_DONE = u"План выгружен в файл:\n%s"
PLAN_EXPORT_DOC_TITLE = u"План выполнения"
PLAN_EXPORT_SHEET = u"План"
PLAN_EXPORT_FILE_PREFIX = u"План"

# Порядок колонок выгрузки: после «Пункт» — этап и «Выполнено»; «Итоговый пункт» не выгружается.
PLAN_EXPORT_COLUMNS = (
    (u"item_no", PLAN_COL_ITEM, 10),
    (u"stage_name", PLAN_COL_STAGE, 28),
    (u"is_done", PLAN_COL_DONE, 10),
    (u"date_start", PLAN_COL_START, 14),
    (u"date_end", PLAN_COL_END, 14),
    (u"duration_wd", PLAN_COL_DURATION, 10),
    (u"assignees", PLAN_COL_ASSIGNEES, 22),
    (u"result_doc", PLAN_COL_RESULT, 34),
    (u"criticality", PLAN_COL_CRITICALITY, 12),
)
PLAN_EXPORT_FONT_NAME = u"PT Sans"
PLAN_EXPORT_FONT_HDR_SIZE = 12
PLAN_EXPORT_FONT_BODY_SIZE = 11
PLAN_EXPORT_FONT_TITLE_SIZE = 12
PLAN_EXPORT_DATE_FORMAT = u"DD.MM.YYYY"
PLAN_EXPORT_DURATION_FORMAT = u"0"
# Заливка строк: выполнено=зелёный; просрочено=красный; скоро=оранжевый; иначе без заливки.
PLAN_EXPORT_FILL_OVERDUE = u"FFCDD2"
PLAN_EXPORT_FILL_SOON = u"FFE0B2"
PLAN_EXPORT_FILL_DONE = u"C8E6C9"
PLAN_EXPORT_FILL_NORMAL = u""

PLAN_WIZARD_TITLE = u"План выполнения"
PLAN_WIZARD_HINT = (
    u"Ввод даты: ДД.ММ.ГГГГ, 01012026 или 01,01,26 — лишние символы игнорируются.\n"
    u"После ввода дата подгоняется к ближайшему рабочему дню (сб/вс и «Праздники» из справочника).\n"
    u"Флажок «Не корр.» справа от даты — сохранить дату как введена.\n"
    u"При смене длительности или даты окончания пересчитываются следующие подэтапы того же родителя.\n"
    u"★ — итоговый пункт в списке слева."
)
PLAN_WIZARD_DETAIL_RESULT_H = 72
PLAN_WIZARD_DETAIL_NO_SNAP_LBL = u"Не корр."
PLAN_DONE_EARLY_CONFIRM_TITLE = u"Выполнение подэтапа"
PLAN_DONE_EARLY_CONFIRM_MSG = (
    u"Текущая дата раньше даты окончания подэтапа (%s).\n"
    u"Подтвердите досрочное выполнение."
)
PLAN_WIZARD_DETAIL_NO_SNAP_W = 80
PLAN_WIZARD_VISIBLE_ROWS = 10
PLAN_WIZARD_SCREEN_RATIO = 1.0
PLAN_WIZARD_SCREEN_MARGIN = 0
PLAN_WIZARD_TOOLBAR_GRID_GAP = 8
PLAN_WIZARD_TOOLBAR_BTN_H = 28
PLAN_WIZARD_FOOTER_BTN_H = 28
PLAN_WIZARD_FOOTER_BTN_W = 110
PLAN_WIZARD_HINT_H = 44
PLAN_WIZARD_HINT_TOP_GAP = 14
PLAN_WIZARD_GRID_BOTTOM_RESERVE = 0
PLAN_WIZARD_GRID_MIN_H = 120
PLAN_WIZARD_TOOLBAR_BTNS = (
    (u"AddRootBtn", u"+ Пункт", 84),
    (u"AddChildBtn", u"+ Подпункт", 100),
    (u"DelBtn", u"Удалить", 84),
    (u"RefreshGridBtn", u"↻ Таблица", 92),
    (u"UpBtn", u"↑", 36),
    (u"DownBtn", u"↓", 36),
    (u"EditBtn", u"Правка строки…", 132),
    (u"RecalcBtn", u"Пересчитать", 104),
)
# Split-форма (ListBox слева + поля справа) — основной режим; Grid отключён.
PLAN_WIZARD_USE_GRID = False
PLAN_WIZARD_LIST_STAGE_MAX = 60
PLAN_WIZARD_SPLIT_GAP = 8
PLAN_WIZARD_LIST_LBL_H = 18
PLAN_WIZARD_DETAIL_ROW_H = 26
PLAN_WIZARD_DETAIL_LBL_H = 16
PLAN_WIZARD_DETAIL_GAP = 10
PLAN_WIZARD_DETAIL_SECTION_GAP = 14
PLAN_WIZARD_DETAIL_STAGE_H = 80
PLAN_WIZARD_DETAIL_DATE_W = 118
PLAN_WIZARD_DETAIL_ERR_INDENT = 14
PLAN_WIZARD_DETAIL_ERR_COLOR = 0xCC0000
PLAN_WIZARD_LIST_DONE_PREFIX = u"[v] "
PLAN_ROW_STATUS_DONE = u"done"
PLAN_ROW_STATUS_OVERDUE = u"overdue"
PLAN_ROW_STATUS_SOON = u"soon"
PLAN_ROW_STATUS_NORMAL = u"normal"
PLAN_ROW_SOON_WORKDAYS = 2
PLAN_DETAIL_COLOR_NORMAL = 0x000000
PLAN_DETAIL_COLOR_DONE = 0x006400
PLAN_DETAIL_COLOR_OVERDUE = 0xCC0000
PLAN_DETAIL_COLOR_SOON = 0xE6A800
PLAN_WIZARD_LIST_MIN_W = 280
PLAN_WIZARD_LIST_MIN_H = 80
PLAN_WIZARD_LIST_BOTTOM_GAP = 6
PLAN_WIZARD_TOOLBAR_BTNS_LIST = (
    (u"AddRootBtn", u"+ Пункт", 84),
    (u"AddChildBtn", u"+ Подпункт", 100),
    (u"DelBtn", u"Удалить", 84),
    (u"UpBtn", u"↑", 36),
    (u"DownBtn", u"↓", 36),
    (u"RecalcBtn", u"Пересчитать", 104),
)
PLAN_DEFAULT_DURATION = 1
PLAN_AUTO_CHAIN_DATES = True

PLAN_SYNC_DUE_TITLE = u"Срок задачи"
PLAN_SYNC_DUE_MSG = (
    u"Установить срок исполнения задачи по дате окончания плана — %s?"
)
PLAN_SYNC_DUE_COMMENT = (
    u"Дата окончания исполнения задачи изменена согласно плану выполнения."
)

PLAN_DATE_DISPLAY_FMT = u"DD.MM.YYYY DDD"
