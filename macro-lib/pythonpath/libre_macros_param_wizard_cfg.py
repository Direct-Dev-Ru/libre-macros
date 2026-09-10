# -*- coding: utf-8 -*-
"""Константы и состояние param_wizard (AlterOffice 2026)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.718"
import re

try:
    unicode
except NameError:
    unicode = str

try:
    unichr
except NameError:
    unichr = chr

PARAM_WIZARD_DEBUG = False
PARAM_WIZARD_TIMING = True
_PW_TIME_DEPTH = 0
_UI_VSCALE = 0.8
# Мягкая серая палитра для диалога «применить_формулу» (не сливается со светлой темой).
# Цвета UNO: 0x00RRGGBB.
# Системный TitleBar (GTK/WM) через UNO обычно не красится — тёмный «титл» рисуем
# полосой SoftGrayTitleBar внутри окна.
_SOFT_GRAY_FORM_BG = 0xE6E6E6
_SOFT_GRAY_TITLE_BG = 0x4A4A4A
_SOFT_GRAY_TITLE_FG = 0xF5F5F5
_SOFT_GRAY_LABEL_FG = 0x2A2A2A
_SOFT_GRAY_FIELD_BG = 0xF4F4F4
_SOFT_GRAY_BTN_BG = 0xD4D4D4
_SOFT_GRAY_TITLE_H = 28
def _ui_v(n):
    return max(1, int(round(float(n) * _UI_VSCALE)))
_DLG_W = 560
_M = 10
_FIELD_W = _DLG_W - _M * 2
_BROWSE_W = 92
_CTRL_H = _ui_v(22)
_LBL_H = _ui_v(18)
_GAP_S = _ui_v(6)
_GAP_L = _ui_v(10)
_BTN_H_MAIN = max(1, int(round(_ui_v(28) * 0.7 * 2)))
_TAB_Y = _ui_v(10)
_SHEET_COMBO_W = 168
_NEW_MODE_Y = _TAB_Y + _CTRL_H + _GAP_S
_MODE_ROW_H = _CTRL_H
_MODE_GAP = _GAP_S
_MODE_BTN_LABEL_EDIT = u"Переключить на Новый параметр"
_MODE_BTN_LABEL_NEW = u"Переключить в Редактирование"
_MODE_BTN_W = 260
_MODE_BTN_X = _M + _FIELD_W - _MODE_BTN_W
_MODE_LBL_W = _MODE_BTN_X - _M - _MODE_GAP
_PAGE_TOP = _NEW_MODE_Y + _MODE_ROW_H + _GAP_L
_PARAM_DISABLED_CHK_Y = _PAGE_TOP + _ui_v(44) + _GAP_S
_PARAM_DISABLED_CHK_H = _LBL_H
_VALUE_LBL_Y = _PAGE_TOP + _ui_v(68)
_VALUE_LBL_TALL_H = _ui_v(32)
_VALUE_CTRL_Y = _PAGE_TOP + _ui_v(86)
_VALUE_CTRL_Y_TALL = _VALUE_LBL_Y + _VALUE_LBL_TALL_H + _GAP_S
_VALUE_CTRL_H = _CTRL_H
_VALUE_BTNS_Y = _VALUE_CTRL_Y + _VALUE_CTRL_H + _GAP_S
# Список колонок «Доп_Параметры_Источника» + кнопка субвизарда
_SOURCE_EXTRA_LIST_H = _ui_v(72)
_SOURCE_EXTRA_LIST_Y = _VALUE_CTRL_Y
_SOURCE_EXTRA_BTN_W = 100
_SOURCE_EXTRA_BTN_Y = _SOURCE_EXTRA_LIST_Y + _SOURCE_EXTRA_LIST_H + _GAP_S
_FILES_BTN_H = _CTRL_H
_FILES_BTN_W = 128
_FILES_BTN_GAP = _ui_v(8)
_EXTRA_LBL_Y = _VALUE_BTNS_Y + _FILES_BTN_H + _GAP_L
_EXTRA_EDIT_Y = _EXTRA_LBL_Y + _LBL_H
_EXTRA_EDIT_H = _CTRL_H
_HIDDEN_CTRL_Y = -5000
_TRANSFER_PAIR_LBL_Y = _VALUE_BTNS_Y
_TRANSFER_PAIR_CTRL_Y = _TRANSFER_PAIR_LBL_Y + _LBL_H
_TRANSFER_WIZARD_DEFAULT = u"По_API"
_DESC_Y = _EXTRA_EDIT_Y + _EXTRA_EDIT_H + _ui_v(26)
_DESC_H = max(1, int(round(_ui_v(190) * 0.7)))
_REFRESH_SHEET_CHK_Y = _DESC_Y + _DESC_H + _ui_v(12)
_REFRESH_SHEET_CHK_H = _LBL_H
_BTN_Y = _REFRESH_SHEET_CHK_Y + _REFRESH_SHEET_CHK_H + _GAP_S
_ACTION_BTN_W = 90
_ACTION_BTN_GAP = _ui_v(8)
_OK_BTN_X = _DLG_W - _M - _ACTION_BTN_W * 2 - _ACTION_BTN_GAP
_CANCEL_BTN_X = _OK_BTN_X + _ACTION_BTN_W + _ACTION_BTN_GAP
_DELETE_BTN_X = _M + _ACTION_BTN_W + _ACTION_BTN_GAP
_RESET_GLOBAL_BTN_W = 150
_RESET_GLOBAL_BTN_X = _M + _ACTION_BTN_W + _ACTION_BTN_GAP
# Ряд футера: Обновить | Защитить | Скрыть
_FOOTER_CHK_GAP = _ui_v(6)
_PROTECT_SHEET_CHK_W = 108
_HIDE_SHEET_CHK_W = 96
_REFRESH_SHEET_CHK_W = max(
    160,
    _FIELD_W - _PROTECT_SHEET_CHK_W - _HIDE_SHEET_CHK_W - 2 * _FOOTER_CHK_GAP,
)
_PROTECT_SHEET_CHK_X = _M + _REFRESH_SHEET_CHK_W + _FOOTER_CHK_GAP
_HIDE_SHEET_CHK_X = _PROTECT_SHEET_CHK_X + _PROTECT_SHEET_CHK_W + _FOOTER_CHK_GAP
# Пустой пароль — защита от случайного редактирования (Tools → Protect Sheet).
_PARAM_SHEET_PROTECT_PASSWORD = u""
# Вложенный write-guard: глубина снятия защиты и флаг «была защита».
_PARAM_SHEET_WRITE_GUARD = {}
_PARAM_SHEET_WRITE_WAS_PROT = {}
_DLG_H = _BTN_Y + _BTN_H_MAIN + _M
# Вкладка «Глобальные»: флаги как в оркестраторе collect_pack.
GLOBAL_FLAG_MARK_ON = u"[v] "
GLOBAL_FLAG_MARK_OFF = u"[x] "
# Визард «Столбцы»: преобразование текста → число/дата (код, подпись)
_COLUMNS_MAP_CONVERT_CHOICES = (
    (u"", u"по умолчанию"),
    (u"true", u"да — преобразовывать"),
    (u"false", u"нет — оставить текст"),
)
_COLUMNS_MAP_CONVERT_CODE = {lbl: code for code, lbl in _COLUMNS_MAP_CONVERT_CHOICES}
_COLUMNS_MAP_CONVERT_LABEL = {code: lbl for code, lbl in _COLUMNS_MAP_CONVERT_CHOICES}

# Пресеты формата в визарде «Столбцы» (дата + число + текст)
_COLUMNS_MAP_FORMAT_BLANK = u"(авто)"
_COLUMNS_MAP_FORMAT_TEXT = u"@"
try:
    from libre_macros_global_settings_lib import (
        DATE_FORMAT_CHOICES,
        calc_number_format_choices,
    )

    _COLUMNS_MAP_FORMAT_PRESETS = (
        (_COLUMNS_MAP_FORMAT_BLANK, u""),
        (_COLUMNS_MAP_FORMAT_TEXT, u"@"),
    ) + tuple((fmt, fmt) for fmt in DATE_FORMAT_CHOICES) + tuple(
        (fmt, fmt) for fmt in calc_number_format_choices()
        if fmt not in DATE_FORMAT_CHOICES
    )
except Exception:
    _COLUMNS_MAP_FORMAT_PRESETS = (
        (_COLUMNS_MAP_FORMAT_BLANK, u""),
        (_COLUMNS_MAP_FORMAT_TEXT, u"@"),
        (u"DD.MM.YYYY", u"DD.MM.YYYY"),
        (u"DD.MM.YYYY DDD", u"DD.MM.YYYY DDD"),
        (u"# ##0,00", u"# ##0,00"),
    )
_COLUMNS_MAP_FORMAT_LABEL_TO_CODE = {
    lbl: code for code, lbl in _COLUMNS_MAP_FORMAT_PRESETS
}
_COLUMNS_MAP_FORMAT_CODE_TO_LABEL = {
    code: lbl for code, lbl in _COLUMNS_MAP_FORMAT_PRESETS if code != u""
}
_COLUMNS_MAP_FORMAT_CODE_TO_LABEL[u""] = _COLUMNS_MAP_FORMAT_BLANK

GLOBAL_BOOL_FLAG_SPECS = (
    (u"MERGE_XML_CONVERT_TO_NUMBERS", u"MERGE_XML_CONVERT_TO_NUMBERS"),
    (u"MERGE_SOURCE_VIEW_SUBFOLDERS", u"MERGE_SOURCE_VIEW_SUBFOLDERS"),
    (u"MERGE_DEBUG", u"MERGE_DEBUG"),
    (u"MERGE_DEBUG_CONSOLE_LOG_TO_WORKSHEET", u"MERGE_DEBUG_CONSOLE_LOG_TO_WORKSHEET"),
    (u"MERGE_CREATE_LOG_SHEET", u"MERGE_CREATE_LOG_SHEET"),
    (u"MERGE_DO_EMPTY_SHEETS", u"MERGE_DO_EMPTY_SHEETS"),
    (u"Автообрезка_листов", u"MERGE_COPY_AUTO_TRIM"),
    (u"MERGE_PASTE_FORMATS", u"MERGE_PASTE_FORMATS"),
    (u"Автофильтр_результата", u"MERGE_RESULT_AUTOFILTER"),
    (u"Закрепить_заголовок", u"MERGE_FREEZE_HEADER"),
    (u"MERGE_CONFIRM_DELETION_OF_SHEETS", u"MERGE_CONFIRM_DELETION_OF_SHEETS"),
)
_GLOBAL_FMT_GAP = _ui_v(8)
_GLOBAL_HALF_W = (_FIELD_W - _GLOBAL_FMT_GAP) // 2
_GLOBAL_NUM_X = _M + _GLOBAL_HALF_W + _GLOBAL_FMT_GAP
_GLOBAL_FLAGS_LB_H = _ui_v(180)
_GLOBAL_FLAG_HINT_H = _ui_v(56)
_GLOBAL_MIRROR_BROWSE_W = 28
_GLOBAL_VARS_BTN_W = 180
_GLOBAL_FLAGS_BTN_W = 160
_GLOBAL_FLAGS_TOGGLE_BTN_W = 72
_GLOBAL_FLAGS_DLG_LIST_H = _ui_v(180)
_FALLBACK_PP_RANGE_CHOICES = [
    u"авто_высота", u"авто_ширина", u"автофильтр", u"высота_строки", u"градиент", u"группировка_по_столбцу",
    u"заголовок_плюс_высота", u"закрепить_заголовок", u"зебра_диапазон", u"сортировка", u"раскрасить_блоки",
    u"конкатенация_столбцов", u"разделить_по_столбцам", u"условный_столбец",
    u"добавить_столбец",
    u"левое_выравнивание", u"отступ", u"перенос",
    u"перенос_и_авто_высота", u"применить_формулу", u"удаление_строк",
    u"сброс_стрипов", u"сетка", u"стиль_печати", u"толстая_сетка",
    u"тонкая_сетка", u"удалить_столбцы",
    u"подсветка_по_порогу", u"подкрасить_пороги", u"формат_даты", u"формат_деньги",
    u"формат_столбцы",     u"ширина_столбцов", u"шрифт", u"заполнение_вниз",
    u"заполнить_вверх",
    u"развернуть_столбцы",
    u"транспонировать_таблицу",
    u"анкета_в_таблицу",
    u"таблица_в_анкету",
    u"заполнение_вниз_вычислить", u"копировать_значения", u"замена_значений", u"текстовые_операции", u"переименовать_лист",
    u"переименовать_столбцы", u"переставить_столбцы",
    u"количество_значений", u"удалить_дубликаты", u"группировать_строки", u"копировать_переместить_лист",
    u"копирование_диапазонов", u"создать_лист",
    u"сводная_таблица",
    u"удаление_листов", u"скрытие_листов",
]
_FALLBACK_PP_ROW_CHOICES = [
    u"высота_строки", u"серые_служебные", u"полосы_по_пути",
    u"чередующиеся_границы", u"копировать_формат_заголовка",
    u"подсветка_по_заголовку", u"цвет_текста_по_значению", u"жирный_по_пути",
    u"преобразовать_числа",
]
_FALLBACK_PP_XML_CHOICES = [
    u"сетка", u"тонкая_сетка", u"толстая_сетка", u"заголовок_плюс_высота",
    u"перенос", u"перенос_и_авто_высота", u"авто_высота", u"авто_ширина",
    u"высота_строки", u"левое_выравнивание", u"отступ", u"шрифт",
    u"зебра_диапазон", u"формат_деньги", u"формат_даты", u"формат_столбцы",
    u"применить_формулу", u"удалить_столбцы", u"конкатенация_столбцов",
    u"разделить_по_столбцам", u"условный_столбец",
    u"замена_значений", u"текстовые_операции", u"заполнение_вниз", u"заполнить_вверх", u"развернуть_столбцы", u"транспонировать_таблицу", u"анкета_в_таблицу", u"таблица_в_анкету", u"сортировка",
    u"переименовать_лист", u"переименовать_столбцы", u"переставить_столбцы",
    u"копировать_переместить_лист", u"закрепить_заголовок",
    u"автофильтр", u"сводная_таблица",
    u"удаление_листов", u"скрытие_листов",
]
_FALLBACK_PP_PLUGIN_REST = (
    u"JSON в C: ref/code, allow_env/allow_global (пусто = MERGE_ALLOW_PLUGINS / "
    u"Merge_Allow_Plugins), extra, sheet; допуск: env ↔ зашифрованная глобальная",
    u'[{"v":1,"fn":"функция_плагин","ref":"functions_pp.py#pp_range_zagotovka","sheet":"Сводка"}]',
)
_FALLBACK_FINAL_CHOICES = [
    u"удаление_листов", u"скрытие_листов", u"только_значения",
    u"сводная_таблица",
    u"заголовок_плюс_высота", u"вертикаль_центр", u"зебра_диапазон", u"сортировка", u"раскрасить_блоки",
    u"сетка", u"тонкая_сетка", u"толстая_сетка",
    u"удалить_столбцы", u"конкатенация_столбцов", u"разделить_по_столбцам",
    u"условный_столбец", u"добавить_столбец", u"применить_формулу",
    u"количество_значений", u"удалить_дубликаты", u"группировать_строки", u"копировать_переместить_лист",
    u"активировать_лист", u"копирование_диапазонов", u"создать_лист",
    u"отправить_по_почте",
    u"заполнение_вниз", u"заполнить_вверх",
    u"развернуть_столбцы",
    u"транспонировать_таблицу",
    u"анкета_в_таблицу",
    u"таблица_в_анкету",
    u"замена_значений", u"текстовые_операции",
    u"переименовать_лист", u"переименовать_столбцы", u"переставить_столбцы",
]
PLUGIN_FUNCTION_KEY = u"функция_плагин"
_PLUGIN_D_LBL_Y = _EXTRA_EDIT_Y + _EXTRA_EDIT_H + _GAP_S
_PLUGIN_D_EDIT_Y = _PLUGIN_D_LBL_Y + _LBL_H
_PLUGIN_DESC_OFFSET = _LBL_H + _EXTRA_EDIT_H + _GAP_S
_DESC_H_PLUGIN = max(_ui_v(72), _DESC_H - _PLUGIN_DESC_OFFSET)
_FALLBACK_FINAL_REST_HINTS = {
    PLUGIN_FUNCTION_KEY: (
        u"JSON в C: ref/code, allow_env/allow_global (пусто = MERGE_ALLOW_PLUGINS / "
        u"Merge_Allow_Plugins), extra, sheet; допуск: env ↔ зашифрованная глобальная",
        u'[{"v":1,"fn":"функция_плагин","ref":"functions_final.py#pp_range_zagotovka"}]',
    ),
    u"удаление_листов": (
        u"имена/шаблоны (*Отчет*); опционально filter — лямбда name→bool (санация);\n"
        u"пусто — служебные по умолчанию; _Первый_; _Последний_; _Последний_мусорный_",
        u'[{"v":1,"fn":"удаление_листов","sheets":["*Отчет*"],'
        u'"filter":"lambda name: sheet_age_days(name) is not None and sheet_age_days(name) <= 3"}]',
    ),
    u"скрытие_листов": (
        u"имена/шаблоны; опционально filter (как у удаления_листов); пусто — служебные;\n"
        u"_Первый_; _Последний_; _Последний_мусорный_",
        u'[{"v":1,"fn":"скрытие_листов","sheets":["Сбор_книг_лог*"]}]',
    ),
    u"только_значения": (
        u"имена листов через запятую или ; — только значения; "
        u"лист со сводной — inplace (без нового листа)",
        u"Сводная",
    ),
    u"сводная_таблица": (
        u"JSON (кнопка «Параметры»). source_sheet — только для указанного листа результата (* — шаблон). "
        u"as_values=true — материализовать перед следующими финальными шагами.",
        u"{\"v\":1,\"source_sheet\":\"Merge\",\"header_row\":\"\",\"data_start\":\"\",\"data_end\":\"\",\"sheet_name\":\"\",\"as_values\":true,\"list_rows_expand\":true,\"row_fields\":[\"Регион\",\"Отдел\"],\"column_fields\":[],\"filter_fields\":[],\"data_fields\":[{\"field\":\"Табельный_номер\",\"function\":\"LIST_ROWS\",\"list_fields\":[\"Табельный_номер\",\"Должность\",\"Филиал\",\"Комментарий\"]}]}",
    ),
    u"заголовок_плюс_высота": (
        u"как в постобработке: height_mm абсолют (мин. 7 мм), add_height — плюсовать; пусто — 12,7 мм center/center",
        u'[{"v":1,"fn":"заголовок_плюс_высота","height_mm":13,"sheet":"Сводная"}]',
    ),
    u"вертикаль_центр": (
        u"вертикальное выравнивание по центру на строках данных (заголовок не трогается); "
        u"C — список листов; пусто — пропуск",
        u"Сводная_1",
    ),
    u"зебра_диапазон": (
        u"имя листа (или несколько через , или ;) | роль:заливка/шрифт — как у постобработки «зебра_диапазон»",
        u"Итог|header:excel_header/белый; even:голубой/синий; odd:белый/синий",
    ),
    u"сортировка": (
        u"«лист | ключи» или «лист: ключи»; задания через ; (как в постобработке)",
        u"Сводная | Сумма -> -, Дата -> + ; Orders | Сумма -> -",
    ),
    u"раскрасить_блоки": (
        u"«лист | столбцы_ключа -> цвета»; блоки через ;. "
        u"Цвета: имя или #hex; несколько — через %. "
        u"Сортировка: /+ или /asc — по возрастанию, /- или /убыв — по убыванию; "
        u"без суффикса — без сортировки; !граница — контур блоков",
        u"Сводная | Отдел -> голубой%белый/+ ; Orders | A, B -> #E2EFDA/-",
    ),
    u"раскрасить": (
        u"алиас «раскрасить_блоки»",
        u"Сводная | Отдел -> голубой",
    ),
    u"сетка": (
        u"как в постобработке: «толщина/цвет» или «лист | 25/#CCCCCC ; …»; пусто — тонкая серая на всех",
        u"Сводная | 25/#CCCCCC ; Orders : толстая/темныйсиний",
    ),
    u"тонкая_сетка": (
        u"список листов через запятую или ; (пусто — все листы)",
        u"Сводная, Orders",
    ),
    u"толстая_сетка": (
        u"список листов через запятую или ; (пусто — все листы)",
        u"Сводная, Orders",
    ),
    u"количество_значений": (
        u"JSON: new_column, key_columns, trim, case_sensitive — как в постобработке",
        u'[{"v":1,"fn":"количество_значений","new_column":"Количество","key_columns":["ФИО"]}]',
    ),
    u"удалить_дубликаты": (
        u"JSON: key_columns, keep, output (на месте|новый лист|смещение|маркировка); "
        u"dest_cell — левый верх результата на этом же листе (A1)",
        u'[{"v":1,"fn":"удалить_дубликаты","key_columns":["ФИО"],"keep":"first"}]',
    ),
    u"группировать_строки": (
        u"JSON: key_columns + aggregations[{op,column,as}]; output inplace|new_sheet; "
        u"сжимает строки до уникальных ключей (не путать с группировка_по_столбцу / SUBTOTAL)",
        u'[{"v":1,"fn":"группировать_строки","key_columns":["\'Отдел\'"],'
        u'"aggregations":[{"op":"sum","column":"\'Сумма\'","as":"Сумма"}],"output":"inplace"}]',
    ),
    u"копировать_переместить_лист": (
        u"JSON: source_sheet + copy (по умолчанию false=перемещение).\n"
        u"copy=false: переместить лист перед/после anchor_sheet; copy=true: копия в dest_sheet.\n"
        u"Для ODS (xml_excel_ods): при copy=true можно задать columns[] + skip_empty=true.",
        u'[{"v":1,"fn":"копировать_переместить_лист","source_sheet":"Шаблон","copy":false,"position":"after","anchor_sheet":"Merge"}]',
    ),
    u"активировать_лист": (
        u"JSON: target_sheet — лист сделать активным; если скрыт — показать; курсор в A1;\n"
        u"tab_color — опционально цвет ярлычка (имя/#hex; «нет»/сброс — умолчание)",
        u'[{"v":1,"fn":"активировать_лист","target_sheet":"Merge","tab_color":"голубой"}]',
    ),
    u"копирование_диапазонов": (
        u"JSON: source_sheet + source_range (A1:D10 / B:B / 2:5) или rows/columns; "
        u"dest_sheet(s)+dest_cell; mode=replace|insert; content=values|formulas; "
        u"после копирования приёмник вовлекается в скоп следующих шагов",
        u'[{"v":1,"fn":"копирование_диапазонов","source_sheet":"Сводная",'
        u'"source_range":"A1:D20","dest_sheet":"Отчет","dest_cell":"B3"}]',
    ),
    u"создать_лист": (
        u"JSON: name_mode=manual|expr + sheet_name / sheet_name_expr;\n"
        u"columns; header_row; data_mode=none|sheet|rows;\n"
        u"sheet: source_sheet+source_range+values_only;\n"
        u"rows: строки через |, ячейки через ,/; ; <<Переменные.список>> → столбец",
        u'[{"v":1,"fn":"создать_лист","name_mode":"manual","sheet_name":"Отчет",'
        u'"columns":["Код","Сумма"],"header_row":1,"data_mode":"rows",'
        u'"rows":"<<Переменные.коды>>, 0"}]',
    ),
    u"отправить_по_почте": (
        u"Письмо не уходит само — откроется окно клиента. attach: whole_book|sheets|none;\n"
        u"to — адреса через ; ; to_source_sheet+to_source_range — дочитать получателей из диапазона;\n"
        u"filename — как указано + суффикс _ГГГГММДД_ЧЧММСС; пусто — имя книги",
        u'[{"v":1,"fn":"отправить_по_почте","attach":"whole_book","to":["boss@corp.local"],'
        u'"subject":"Отчет"}]',
    ),
    u"email": (
        u"алиас «отправить_по_почте»",
        u'[{"v":1,"fn":"отправить_по_почте","attach":"whole_book","to":["a@b.c"]}]',
    ),
    u"почта": (
        u"алиас «отправить_по_почте»",
        u'[{"v":1,"fn":"отправить_по_почте","attach":"whole_book"}]',
    ),
    u"удалить_столбцы": (
        u"markers: подстроки/имена, индексы 1-based, буквы A/C; смесь через запятую",
        u"#Путь,3,A,Коммент*",
    ),
    u"конкатенация_столбцов": (
        u"имя_нового_столбца, разделитель, индексы 1-based; без разделителя — индексы",
        u"Полный_Код,|,1,3,5",
    ),
    u"разделить_по_столбцам": (
        u"JSON: column, delimiter, position, relative_column, source_policy; кнопка «Параметры…»",
        u'[{"v":1,"fn":"разделить_по_столбцам","column":"ФИО","delimiter":",","position":"_после_"}]',
    ),
    u"условный_столбец": (
        u"JSON: new_column + rules[] + otherwise; кнопка «Параметры…» (конструктор правил/ELSE)",
        u'[{"v":1,"fn":"условный_столбец","new_column":"Категория","rules":[{"column":"Сумма","op":"ge","value":1000,"then":"Средний"}],"otherwise":"Мелкий"}]',
    ),
    u"добавить_столбец": (
        u"JSON: new_column, data_type (text|number|date), format; text→choices[]; "
        u"number/date→data_validation + min/max; лист — SameAll / Sheet",
        u'[{"v":1,"fn":"добавить_столбец","new_column":"Статус","data_type":"text","choices":["Активен","Архив"]}]',
    ),
    u"применить_формулу": (
        u"имя_столбца=формула Calc; {row_id} — первая строка данных",
        u"Итого=E{row_id}+F{row_id}; карта: <<Переменные.Филиал>> / <<Переменные.Филиал([Источник_Лист],[Источник_Файл])>> / <<Переменные.Филиал(#Путь)>>",
    ),
    u"замена_значений": (
        u"JSON: columns+find; replace[] или replace_mode=lambda+replace_expr (кортеж); "
        u"нехватка replace → частичная замена + warning; row_filter/match_pick",
        u'[{"v":1,"fn":"замена_значений","columns":["ФИО"],"find":["ООО"],'
        u'"replace_mode":"lambda","replace_expr":"lambda rows: \\"<<Переменные.Филиал>>\\""}]',
    ),
    u"переименовать_лист": (
        u"новое имя: <<Переменные…>> / [Старое_имя] / [Заголовок(N)] / [Строка(R,C)] / "
        u"[Текущая_Дата_Время_YYYYMMDD] / [Имя_Книги] / [Номер_Листа] / [GUID8]; "
        u"формат даты YYYY/YY/MM/DD/HH/mm/SS; нормализация ≤31",
        u"Свод_[Текущая_Дата_Время_YYYYMMDD]",
    ),
    u"переименовать_столбцы": (
        u"JSON: sheet + mappings; old в 'кавычках' = полное совпадение; "
        u"в new — те же плейсхолдеры, что у переименовать_лист "
        u"(<<Переменные…>>, [Старое_имя], дата/время, [Имя_Книги], …)",
        u'[{"v":1,"fn":"переименовать_столбцы","sheet":"Заказы",'
        u'"mappings":[{"old":"\'Сумма\'","new":"<<Переменные.Филиал>>_[Старое_имя]"}]}]',
    ),
    u"переставить_столбцы": (
        u"JSON: sheet (обязателен), columns, position (_начало_|_конец_|_перед_|_после_), "
        u"relative_column для _перед_/_после_; copy=true + new_names — копии; "
        u"create_new=true + new_names — пустые столбцы по «Куда»",
        u'[{"v":1,"fn":"переставить_столбцы","sheet":"Заказы","columns":["C"],'
        u'"position":"_конец_","copy":true,"new_names":["C_копия"]}]',
    ),
}
_VLOOKUP_COL_HEADERS = (
    u"Левая_Таблица",
    u"Правая_Таблица",
    u"Ключ_сравнения",
    u"Столбцы_для_извлечения",
    u"Цвет_добавленных_данных",
)
_FALLBACK_PP_RANGE_REST_HINTS = {
    u"сетка": (
        u"толщина/цвет — тонкая|толстая|число / имя цвета или hex; "
        u"по листам: «лист | 25/#CCCCCC ; …» (разделитель |, : или /); пусто = тонкая серая",
        u"Сводная | 25/#CCCCCC ; Orders : толстая/темныйсиний",
    ),
    u"тонкая_сетка": (
        u"список листов через запятую или ; (пусто — все листы)",
        u"Сводная, Orders",
    ),
    u"толстая_сетка": (
        u"список листов через запятую или ; (пусто — все листы)",
        u"Сводная, Orders",
    ),
    u"заголовок_плюс_высота": (
        u"height_mm — высота в мм (по умолчанию абсолют, мин. 7 мм); "
        u"add_height=true — прибавить к текущей; выравнивание left/center/right и top/center/bottom",
        u'[{"v":1,"fn":"заголовок_плюс_высота","height_mm":13,"add_height":false,"h_align":"center","v_align":"center"}]',
    ),
    u"перенос": (u"пусто = включить; No/Нет — выключить перенос", u"Нет"),
    u"перенос_и_авто_высота": (u"как у «перенос»: пусто = включить; No/Нет — выключить", u"Нет"),
    u"авто_высота": (u"автоподбор высоты строк данных (OptimalHeight)", u""),
    u"авто_ширина": (u"автоподбор ширины столбцов диапазона (OptimalWidth)", u""),
    u"высота_строки": (
        u"пусто — все строки данных, 5 мм; «6» — все строки; «2,3 / 6» (или ->, =); «0» — скрыть; "
        u"«6 мм» тоже допускается",
        u"2,3,4 / 6",
    ),
    u"левое_выравнивание": (u"не используется", u""),
    u"отступ": (
        u"[left|center|right], число_шагов [, столбцы]; пусто = left, 1, все",
        u"right,1,Сумма",
    ),
    u"шрифт": (
        u"имя;размер_данных;размер_заголовка (или через запятую); «-» = не менять",
        u"Arial;11;12",
    ),
    u"зебра_диапазон": (
        u"роль:заливка/шрифт через «;» или «,» — header/заголовок, even/четные, odd/нечетные",
        u"header:excel_header/белый, even:голубой/синий, odd:белый/синий",
    ),
    u"сортировка": (
        u"«лист | ключи» или «лист: ключи»; задания через ;. "
        u"Без листа — ключи для всех листов. Ключ: колонка -> +/-",
        u"Сводная | Сумма -> -, Дата -> + ; Orders | Сумма -> -",
    ),
    u"раскрасить_блоки": (
        u"«лист | столбцы_ключа -> цвета»; блоки через ;. "
        u"Без листа — для текущего листа. Цвета: имя или #hex; несколько — через %. "
        u"Сортировка: /+ или /asc — по возрастанию, /- или /убыв — по убыванию; "
        u"без суффикса — без сортировки; !граница — контур блоков",
        u"Отдел, Город -> голубой%белый%серый/+!граница",
    ),
    u"раскрасить": (
        u"алиас «раскрасить_блоки»",
        u"Отдел -> голубой",
    ),
    u"формат_деньги": (
        u"лист, формат (пусто = глобальный), столбцы, маркеры заголовка",
        u'{"sheet":"Отчет","format":"# ##0,00","columns":"Сумма,3"}',
    ),
    u"формат_даты": (
        u"лист, формат (пусто = глобальный), столбцы, маркеры заголовка",
        u'{"sheet":"Отчет","format":"DD.MM.YYYY","markers":"дата,время"}',
    ),
    u"формат_столбцы": (
        u"столбец: имя, номер, буква, список через запятую, Все/*; формат — шаблон (пусто = не менять)",
        u"ФИО, Отдел,# ##0,00",
    ),
    u"подсветка_по_порогу": (
        u"столбцы, порог min/max; JSON: color (пусто=без заливки), style_columns, whole_row, font_color, bold, italic",
        u'{"columns":["\'Итого\'"],"threshold_min":0,"threshold_max":1000,"style_columns":["\'Статус\'"],"font_color":"darkblue","bold":true}',
    ),
    u"подкрасить_пороги": (
        u"алиас «подсветка_по_порогу»",
        u'{"columns":["\'Итого\'"],"threshold_min":0,"threshold_max":1000,"style_columns":["\'Статус\'"],"font_color":"darkblue"}',
    ),
    u"градиент": (
        u"маркер столбца, цвет min, цвет max; JSON: whole_row, sort, sort_dir (по_возр/по_убыв)",
        u"сумм -> зеленый -> красный",
    ),
    u"удалить_столбцы": (
        u"markers: подстроки/имена, индексы 1-based, буквы A/C; смесь через запятую",
        u"#Путь,3,A,Коммент*",
    ),
    u"конкатенация_столбцов": (
        u"имя_нового_столбца, разделитель, индексы 1-based (1=A); без разделителя — индексы",
        u"Полный_Код,|,1,3,5",
    ),
    u"разделить_по_столбцам": (
        u"JSON: column, delimiter, position, relative_column, source_policy; кнопка «Параметры…»",
        u'[{"v":1,"fn":"разделить_по_столбцам","column":"ФИО","delimiter":",","position":"_после_"}]',
    ),
    u"условный_столбец": (
        u"JSON: new_column + rules[] + otherwise; кнопка «Параметры…» (конструктор правил/ELSE)",
        u'[{"v":1,"fn":"условный_столбец","new_column":"Категория","rules":[{"column":"Сумма","op":"ge","value":1000,"then":"Средний"}],"otherwise":"Мелкий"}]',
    ),
    u"добавить_столбец": (
        u"JSON: new_column, data_type, format; text→choices; number/date→data_validation+min/max",
        u'[{"v":1,"fn":"добавить_столбец","new_column":"Сумма","data_type":"number","format":"# ##0,00","data_validation":true,"min":"0","max":"1000000"}]',
    ),
    u"применить_формулу": (
        u"имя_столбца=формула Calc; {row_id} — первая строка данных (1-based); "
        u"{row_id-1} — смещение; as_values — только Диапазон/Финал",
        u"Итого=E{row_id}+F{row_id}",
    ),
    u"удаление_строк": (
        u"столбцы из списка ('ФИО') или формула пустой строки: ИСТИНА → удалить",
        u"'Руководитель'",
    ),
    u"сброс_стрипов": (u"не используется", u""),
    u"группировка_по_столбцу": (
        u"JSON: marker (столбец группировки), agg_fn + agg_columns (промежуточные итоги), "
        u"sort/outline; кнопка «Параметры…»",
        u'[{"v":1,"fn":"группировка_по_столбцу","marker":"\'Отдел\'","agg_fn":"sum","agg_columns":["\'Сумма\'"],"sort":true,"outline":true,"sheet":"МСУЗ"}]',
    ),
    u"закрепить_заголовок": (u"не используется", u""),
    u"ширина_столбцов": (
        u"JSON: columns (Все / список заголовков в '…') + width_mm; 0 мм — скрыть; "
        u"кнопка «Параметры…»",
        u'[{"v":1,"fn":"ширина_столбцов","columns":"all","width_mm":30}]',
    ),
    u"автофильтр": (
        u"JSON: sheets — список листов; action=set|remove; "
        u"smart_table=true + table_style — экспериментальный smart-диапазон; "
        u"keep_smart_table=true — исключение листа из очистки smart-таблиц.",
        u'[{"v":1,"fn":"автофильтр","sheets":["Заказы"],"action":"set"}]',
    ),
    u"стиль_печати": (
        u"ориентация (landscape/portrait), вписать в N страниц по ширине",
        u"landscape,1",
    ),
    u"заполнение_вниз": (
        u"столбцы из списка заголовков ('ФИО'); индексы/буквы тоже ок",
        u"'ФИО';'Отдел'",
    ),
    u"заполнить_вверх": (
        u"столбцы ('ФИО'); пустые → значение из ближайшей непустой снизу",
        u'[{"v":1,"fn":"заполнить_вверх","columns":["\'ФИО\'","\'Отдел\'"]}]',
    ),
    u"развернуть_столбцы": (
        u"Unpivot: unpivot_columns или exclude_columns; attribute/value; inplace|new_sheet",
        u'[{"v":1,"fn":"развернуть_столбцы","unpivot_columns":["\'Январь\'","\'Февраль\'"],'
        u'"attribute_column":"Месяц","value_column":"Сумма","drop_empty_rows":true}]',
    ),
    u"транспонировать_таблицу": (
        u"Транспонирование: inplace|new_sheet|offset; headers_from_column; result_headers",
        u'[{"v":1,"fn":"транспонировать_таблицу","output":"new_sheet","dest_sheet":"Матрица_T","result_headers":["A","B"]}]',
    ),
    u"анкета_в_таблицу": (
        u"Пары Q/A → wide: block_mode + block_start_question; new_sheet",
        u'[{"v":1,"fn":"анкета_в_таблицу","question_column":"A","answer_column":"B",'
        u'"block_mode":"by_repeat_key","block_start_question":"ФИО","dest_sheet":"Анкеты_wide"}]',
    ),
    u"таблица_в_анкету": (
        u"Wide → пары Q/A: body_columns; question_header/answer_header (Вопрос/Ответ); "
        u"has_header_in_output; block_separator",
        u'[{"v":1,"fn":"таблица_в_анкету","body_columns":["\'ФИО\'","\'Возраст\'"],'
        u'"question_header":"Вопрос","answer_header":"Ответ","has_header_in_output":true,'
        u'"block_separator":"blank_row","dest_sheet":"Анкеты"}]',
    ),
    u"заполнение_вниз_вычислить": (
        u"JSON: columns + rules (column, formula, expand_to_right); кнопка «Параметры…»",
        u'[{"v":1,"fn":"заполнение_вниз_вычислить","columns":["A"],"rules":[{"column":"A","formula":"=A{row-1}"}]}]',
    ),
    u"копировать_значения": (
        u"столбцы из списка ('Сумма') или «весь лист»; пусто — весь лист",
        u"'Сумма','Количество'",
    ),
    u"замена_значений": (
        u"JSON: columns+find; replace[] или replace_mode=lambda+replace_expr; "
        u"<<Переменные…>>; флаги, row_filter, match_pick; кнопка «Параметры…»",
        u'[{"v":1,"fn":"замена_значений","columns":["ФИО"],"find":["ООО"],"replace":["<<Переменные.Филиал>>"],"case_insensitive":true,"squeeze_spaces":true}]',
    ),
    u"текстовые_операции": (
        u"JSON: columns + ops[] (trim/collapse_ws/case/lang_homoglyphs/substring/clear/"
        u"reverse/base64/sha256/guid/encrypt…); row_filter per-op; non_text/on_error; "
        u"ключ шифра key|key_cell|key_file. Кнопка «Параметры…» / «Справка».",
        u'[{"v":1,"fn":"текстовые_операции","columns":["ФИО"],"ops":[{"op":"collapse_ws"},{"op":"case","mode":"proper"}]}]',
    ),
    u"переименовать_лист": (
        u"новое имя: <<Переменные…>> / [Старое_имя] / [Заголовок(N)] / [Строка(R,C)] / "
        u"[Текущая_Дата_Время_YYYYMMDD] / [Имя_Книги] / [Номер_Листа] / [GUID8]; "
        u"формат даты YYYY/YY/MM/DD/HH/mm/SS; нормализация ≤31",
        u"Свод_[Текущая_Дата_Время_YYYYMMDD]",
    ),
    u"переименовать_столбцы": (
        u"JSON: sheet + mappings; old в 'кавычках' = полное совпадение; "
        u"в new — те же плейсхолдеры, что у переименовать_лист "
        u"(<<Переменные…>>, [Старое_имя], дата/время, [Имя_Книги], …)",
        u'[{"v":1,"fn":"переименовать_столбцы","sheet":"Заказы",'
        u'"mappings":[{"old":"\'Сумма\'","new":"<<Переменные.Филиал>>_[Старое_имя]"}]}]',
    ),
    u"переставить_столбцы": (
        u"JSON: sheet (обязателен), columns, position (_начало_|_конец_|_перед_|_после_), "
        u"relative_column для _перед_/_после_; copy=true + new_names — копии; "
        u"create_new=true + new_names — пустые столбцы по «Куда»",
        u'[{"v":1,"fn":"переставить_столбцы","sheet":"Заказы","columns":["C"],'
        u'"position":"_конец_","copy":true,"new_names":["C_копия"]}]',
    ),
    u"количество_значений": (
        u"JSON: new_column, key_columns, trim, case_sensitive — см. визард «Параметры…»",
        u'[{"v":1,"fn":"количество_значений","new_column":"Количество","key_columns":["ФИО"]}]',
    ),
    u"удалить_дубликаты": (
        u"JSON: key_columns, keep, output (на месте|новый лист|смещение|маркировка), "
        u"dest_cell — якорь A1 на этом же листе для «смещение»",
        u'[{"v":1,"fn":"удалить_дубликаты","key_columns":["ФИО"],"keep":"first"}]',
    ),
    u"группировать_строки": (
        u"JSON: key_columns + aggregations[{op,column,as}]; output inplace|new_sheet — "
        u"сжатие строк (Group By); не путать с группировка_по_столбцу",
        u'[{"v":1,"fn":"группировать_строки","key_columns":["\'Отдел\'"],'
        u'"aggregations":[{"op":"sum","column":"\'Сумма\'","as":"Сумма"}],"output":"inplace"}]',
    ),
    u"копировать_переместить_лист": (
        u"JSON: source_sheet + copy=false|true; false — перемещение before/after anchor_sheet, true — копия в dest_sheet",
        u'[{"v":1,"fn":"копировать_переместить_лист","source_sheet":"Шаблон","copy":false,"position":"before","anchor_sheet":"Merge"}]',
    ),
    u"копирование_диапазонов": (
        u"JSON: source_sheet + source_range / rows / columns → dest_sheet(s) + dest_cell; "
        u"mode=replace|insert; content=values|formulas",
        u'[{"v":1,"fn":"копирование_диапазонов","source_sheet":"Сводная",'
        u'"source_range":"A1:D20","dest_sheet":"Отчет","dest_cell":"B3"}]',
    ),
    u"создать_лист": (
        u"JSON: name_mode + sheet_name/expr; columns; header_row; "
        u"data_mode=none|sheet|rows (sheet: source_*; rows: | и <<Переменные.список>>)",
        u'[{"v":1,"fn":"создать_лист","name_mode":"manual","sheet_name":"Отчет",'
        u'"columns":["Код","Сумма"],"header_row":1,"data_mode":"rows",'
        u'"rows":"<<Переменные.коды>>, 0"}]',
    ),
    u"удаление_листов": (
        u"имена/шаблоны (*Отчет*); опционально filter=lambda name:… (санация); пусто — служебные по умолчанию; "
        u" _Первый_; _Последний_; _Последний_мусорный_",
        u'[{"v":1,"fn":"удаление_листов","sheets":["*Отчет*"],'
        u'"filter":"lambda name: sheet_age_days(name) is not None and sheet_age_days(name) <= 3"}]',
    ),
    u"скрытие_листов": (
        u"имена/шаблоны; опционально filter (как у удаления); пусто — служебные по умолчанию; "
        u"_Первый_; _Последний_; _Последний_мусорный_",
        u'[{"v":1,"fn":"скрытие_листов","sheets":["Сбор_книг_лог*"]}]',
    ),
    u"сводная_таблица": (
        u"JSON (кнопка «Параметры»). source_sheet — только для указанного листа результата (* — шаблон).",
        u"{\"v\":1,\"source_sheet\":\"Merge\",\"sheet_name\":\"Сводная\",\"as_values\":false,\"row_fields\":[\"Отдел\"],\"column_fields\":[],\"filter_fields\":[],\"data_fields\":[{\"field\":\"Сумма\",\"function\":\"SUM\"}]}",
    ),
    PLUGIN_FUNCTION_KEY: _FALLBACK_PP_PLUGIN_REST,
}
_FALLBACK_PP_XML_REST_HINTS = {
    key: _FALLBACK_PP_RANGE_REST_HINTS[key]
    for key in _FALLBACK_PP_XML_CHOICES
    if key in _FALLBACK_PP_RANGE_REST_HINTS
}
_FALLBACK_PP_XML_REST_HINTS[u"применить_формулу"] = (
    u"имя_столбца=формула Calc; as_values в Постобработка_xml запрещён "
    u"(формулы остаются в файле; галка в визарде недоступна)",
    u'[{"v":1,"fn":"применить_формулу","column":"Итого","formula":"=E{row_id}+F{row_id}"}]',
)
_FALLBACK_PP_ROW_REST_HINTS = {
    u"высота_строки": (
        u"пусто — 5 мм на строку; «6» — 6 мм; «2,3 / 6» (или ->, =) — только эти строки (1-based); «0» — скрыть; "
        u"«6 мм» тоже допускается",
        u"2,3,4 / 6",
    ),
    u"серые_служебные": (u"не используется", u""),
    u"полосы_по_пути": (u"не используется", u""),
    u"чередующиеся_границы": (u"не используется", u""),
    u"копировать_формат_заголовка": (u"не используется", u""),
    u"подсветка_по_заголовку": (u"не используется", u""),
    u"цвет_текста_по_значению": (u"подстрока заголовка столбца-источника", u"статус"),
    u"жирный_по_пути": (u"не используется (подстрока в коде макроса)", u""),
    u"преобразовать_числа": (
        u"колонки (опционально): индексы 0,2,5 или имена Сумма,Количество",
        u"",
    ),
    PLUGIN_FUNCTION_KEY: (
        u"C — functions_pp.py#имя или Python lambda/def; D — доп. аргументы через запятую",
        u"functions_pp.py#pp_row_zagotovka",
    ),
}
MERGE_PARAM_SHEET_NAME = u"collect_params"
MERGE_PARAM_SHEET_MAX_ROWS = 200
P_MERGE_COMMENT = u"Комментарий"
P_MERGE_VERSION = u"Версия"
_PARAM_DV_SHEET_NAME = u"_Визард_Списки"
_PARAM_DV_INLINE_MAX = 248
_PARAM_DV_LIST_CACHE = {}
_PARAM_DV_NEXT_COL = 0
_COLLECT = None
_LM_LIB = None
_PP_PLUGIN_SIG_RANGE = (
    "doc",
    "sheet",
    "data_range",
    "header_row_range",
)
_PP_PLUGIN_SIG_ROW = (
    "doc",
    "sheet",
    "data_row_range",
    "header_row_range",
    "row_index",
)
_PP_PLUGIN_SIG_FINAL = (
    "doc",
    "sheets",
    "data_ranges",
    "header_row_ranges",
)
VALUE_HINTS = {
    u"На один лист": (
        u"Все данные собираются на один лист результата (обычно «Merge»).\n"
        u"Доступны служебные колонки #Путь и #ДатаВремяФайла."
    ),
    u"На разные листы": (
        u"Каждый источник/лист — отдельная вкладка (_N в имени).\n"
        u"Служебный лист «Индексы_Источников»."
    ),
    u"Копирование листов": (
        u"Целые листы копируются как вкладки (как «Копировать лист»).\n"
        u"Без встроенного оформления; постобработка по блоку параметров\n"
        u"(Начало_Строка, Начало_Столбец, Число_Строк, Число_Столбцов, Строка_Заголовков)\n"
        u"или по значащей области листа, если границы не заданы.\n"
        u"При конфликте имён: N_имя_листа (N — порядковый номер совпадения имени листа).\n"
        u"Способ переноса: по умолчанию По_API. Для xml_excel_ods выберите в поле "
        u"«Способ_переноса (C)» в визарде или задайте отдельную строку «Способ_переноса».\n"
        u"Служебные #Путь / #ДатаВремяФайла не добавляются."
    ),
    u"Текущие листы": (
        u"Без копирования и объединения: вкладки текущей книги (эта_книга / ThisWorkbook)\n"
        u"сразу идут в постобработку и финальную обработку.\n"
        u"В «Файлы-Источники» — только маркер текущей книги "
        u"(эта_книга, Эта-книга, ThisWorkbook, This_Workbook, …).\n"
        u"«Листы» — какие вкладки обработать. Предварительное удаление листов отключено.\n"
        u"Способ_переноса и столбцы/диапазоны сбора не используются."
    ),
    u"Вовлечь листы": (
        u"Legacy-имя режима «Текущие листы». Без копирования: вкладки эта_книга → постобработка/финал."
    ),
    u"По_API": (
        u"Перенос через UNO API: getDataArray/setDataArray прямоугольниками. "
        u"Поведение по умолчанию, если способ не задан."
    ),
    u"Буфер_по_заголовкам": (
        u"Копирование через буфер обмена по столбцам с матчингом заголовков. "
        u"Если в источнике есть скрытый лист сбора — автоматически По_API."
    ),
    u"Буфер_по_позиции": (
        u"Копирование прямоугольника через буфер без матчинга; заголовки — при первой записи."
    ),
    u"Не_выводить": u"Служебная колонка не добавляется.",
    u"Коротко": u"Короткий вариант (путь или дата без времени).",
    u"Длинно": u"Полный вариант (полный путь или дата+время).",
    u"Да": u"Включить (автофильтр, закрепление заголовка). Допустимо: +/-, true/false, yes/no, истина/ложь.",
    u"Нет": u"Отключить. Регистр не важен.",
    u"Статусная_строка": u"Прогресс в строке состояния Calc (по умолчанию).",
    u"Диалог": u"Отдельное окно с полосой прогресса на этапе 2.",
}
WIZARD_HINTS = {
    P_MERGE_COMMENT: (
        u"Описание сценария для человека (макрос не читает).\n"
        u"Пример: «Слияние отчётов филиалов за январь, один лист Merge»."
    ),
    u"Файлы-Источники": (
        u"Пути к книгам-источникам: B — первый файл, C — второй, D — третий…\n"
        u"Форматы: .xlsx, .xlsm, .xls, .ods, .xml, .csv.\n"
        u"Допустимы маски: /data/отчёт_*.xlsx , *source_* , *source_*.csv (регистронезависимо).\n"
        u"Для .xls/.ods/.csv используйте *source_* (маска *.xlsx не включает их).\n"
        u"Плейсхолдеры карты: <<Переменные.Имя>> в пути или маске\n"
        u"  (…/<<Переменные.Филиал>>_one_sheet.xlsx или *<<Переменные.Код>>*).\n"
        u"  Подстановка до разворота glob из runtime-карты (глобальные +\n"
        u"  «Ручной_ввод_в_карту_переменных» ВЫШЕ этой строки).\n"
        u"  В collect_pack проверка наличия файлов при плейсхолдерах/ручном\n"
        u"  вводе выше — после диалога при реальном сборе, не в prevalidate.\n"
        u"MERGE_SOURCE_VIEW_SUBFOLDERS=Да — рекурсивно во всех подпапках стартовой директории.\n"
        u"MERGE_SOURCE_EXCLUDE_MASK — доп. исключения (по умолчанию --*); имена ~* пропускаются всегда.\n"
        u"Маркер текущей книги: эта_книга / Эта-книга / этакнига / ThisWorkbook / This_Workbook / current…\n"
        u"Визард добавляет путь в следующий свободный столбец строки.\n"
        u"Для .csv см. «Доп_Параметры_Источника» (разделитель и кодировка).\n"
        u"Примеры:\n"
        u"  /home/user/source_01.xlsx\n"
        u"  /home/user/sources/*source_*.xlsx\n"
        u"  /home/user/sources/<<Переменные.Префикс>>_one_sheet.xlsx\n"
        u"  /home/user/data/*.csv\n"
        u"  Эта_книга"
    ),
    u"Доп_Параметры_Источника": (
        u"JSON-настройки на каждый столбец «Файлы-Источники» (B, C, D…).\n"
        u"Пустая ячейка у источника → берётся JSON из B (значение по умолчанию).\n"
        u"Если вся строка пуста — для CSV: разделитель «;», кодировка Windows-1251.\n"
        u"Для .csv:\n"
        u'  {"delimiter":";","encoding":"windows-1251","open_as_xlsx":true,"copy_whole_sheet":false,"exclude_mask":"--*,*.tmp"}\n'
        u"encoding: windows-1251 (по умолчанию), utf-8, utf-8-sig, utf-16le, cp866, koi8-r…\n"
        u"open_as_xlsx: true (по умолчанию) — CSV конвертируется во временный .xlsx\n"
        u"  и дальше идёт как обычная книга (листы по имени файла). false — матрица в памяти.\n"
        u"  При MERGE_DEBUG временный xlsx не удаляется; путь пишется в журнал.\n"
        u"copy_whole_sheet: true — в режиме «Копирование листов» открыть CSV в Calc\n"
        u"  и перенести лист целиком (importSheet), без матрицы/setDataArray.\n"
        u"  При true отдельная конвертация чисел/дат не делается (Calc уже распознал).\n"
        u"exclude_mask — маски исключения имён файлов через запятую (*);\n"
        u"  складываются с параметром листа MERGE_SOURCE_EXCLUDE_MASK (если задан);\n"
        u"  действуют и для glob (*.xlsx), и для явно указанных путей.\n"
        u"file_pick — доп. фильтр после маски/exclude: all (умолч.), newest_ctime / oldest_ctime,\n"
        u"  newest_mtime / oldest_mtime (один файл), lambda (фильтр N файлов).\n"
        u"  file_pick_scope: all | per_subdir (экстремум/лямбда в каждом подкаталоге).\n"
        u"  file_pick_lambda: выражение с path/имя/mtime/ctime/size и <<Переменные.…>>.\n"
        u"  Дата создания: по возможности ОС; на Linux ctime — смена метаданных (fallback mtime).\n"
        u"Для .xml — exclude_mask и file_pick (см. визард «Параметры…»).\n"
        u"Для .ods/.xls/.xlsx — prepare_source_in_memory, exclude_mask, file_pick,\n"
        u"  skip_missing_sheets, sheet_unprotect_password, last_row_column\n"
        u"  (см. визард «Параметры…»).\n"
        u"  sheet_unprotect_password хранится зашифрованным (lm1:). "
        u"Ключ: глобальная переменная encrypt_passwords_default, иначе литерал с тем же именем.\n"
        u"last_row_column: колонка анализа последней значащей строки (например B, 2,\n"
        u"  Код_позиции) — сначала нижняя граница ищется как обычно, затем уточняется\n"
        u"  по этой колонке, чтобы не тянуть хвост из предзаполненных столбцов.\n"
        u"Для Эта_книга — skip_missing_sheets, sheet_unprotect_password и last_row_column\n"
        u"  (пропуск отсутствующих листов из «Листы»; пароль снятия защиты листов;\n"
        u"  уточнение последней строки по выбранной колонке).\n"
        u"В визарде выберите столбец источника и нажмите «Параметры…».\n"
        u"Запуск с курсором в C/D… сразу пишет JSON в эту колонку."
    ),
    u"Предварительный_скрипт": (
        u"Внешняя команда перед сбором источников (этап 2).\n"
        u"Визард по умолчанию: B = «смотри в С», JSON в колонке C\n"
        u"(кнопка «Параметры…»). Пустая C — скрипт не запускается.\n"
        u"Legacy: JSON целиком в B (без маркера) ещё читается.\n"
        u"Допуск (обязательно): переменная среды MERGE_ALLOW_PRE_SCRIPT\n"
        u"  и зашифрованная глобальная Merge_Allow_Pre_Scripts\n"
        u"  (вкладка «Глобальные» → «Глобальные переменные»; имя в любых\n"
        u"  регистрах / с _ или - / слитно; значение всегда lm1:).\n"
        u"Значения среды и глобальной переменной должны совпадать; иначе сбор\n"
        u"прерывается до диалога подтверждения. Ручной ввод / значения из\n"
        u"файлов-источников для допуска не используются.\n"
        u"Перед запуском — диалог «Внимание!!!»; согласие/отказ пишется в\n"
        u"~/.config/libre-macros/pre_shell_consent/ (дата, пользователь, книга).\n"
        u"При ненулевом коде возврата или таймауте сбор прерывается.\n"
        u"Ключи JSON: argv, cwd, timeout, env, clean_env (по умолчанию true).\n"
        u"clean_env=true снимает PYTHONHOME/PYTHONPATH и пути AlterOffice из PATH.\n"
        u"argv — список аргументов без shell, напр.:\n"
        u'  ["/usr/bin/python3","/path/ad_export.py","--config","/path/ad_export.ini"]\n'
        u"Если задан pre-скрипт, проверка наличия файлов-источников — после его отработки.\n"
        u"Пример C:\n"
        u'  {"argv":["/bin/bash","/path/run.sh"],"cwd":"/path","timeout":600,'
        u'"clean_env":true,"env":{"AD_BIND_PASSWORD":"…"}}\n'
        u"В визарде — кнопка «Параметры…» у колонки C или курсор в B/C."
    ),
    u"Листы": (
        u"Фильтр листов на каждый файл (B,C,D… — как у «Файлы-Источники»).\n"
        u"Колонка B — шаблоны через запятую (text_match): МТ*, Отчет*, Все, *, 1.\n"
        u"Или визард соответствия: источник → результат (JSON по умолчанию):\n"
        u'  {"v":1,"fn":"листы","items":[{"source":"Заказы","target":"Заказы_1"}]}\n'
        u"Legacy: Заказы->Заказы_1,Отчет*,!hidden\n"
        u"Колонка C (кнопка «Параметры…») — опциональный JSON-фильтр (как у «Листы_не_удалять»):\n"
        u'  {"v":1,"filter":"lambda name: sheet_age_days(name) is not None and sheet_age_days(name) <= 3"}\n'
        u"Лямбда получает имя листа из отобранных по B; True — включить в обработку.\n"
        u"Если C — не JSON/lambda, это шаблоны/карта для 2-го файла (как раньше); D — для 3-го и т.д.\n"
        u"В режиме «Текущие листы» — какие вкладки текущей книги обработать (без копирования).\n"
        u"Имя с цифрой в начале («1_Merge») — это имя листа, не номер позиции.\n"
        u"Номер (1, 2…) — лист по порядку вкладок слева направо.\n"
        u"Примеры B:\n"
        u"  Отчет*     — листы, имя начинается с «Отчет»\n"
        u"  *2024      — заканчиваются на 2024\n"
        u"  1          — первый лист книги\n"
        u"  Все        — все листы"
    ),
    u"Столбцы": (
        u"Фильтр и переименование столбцов (B,C,D… по файлам). Клик по B/C/D — визард.\n"
        u"Сверху — общий шаблон отбора (пусто или * = все столбцы).\n"
        u"Ниже — блоки по столбцам: исходное имя → итоговое, преобразование в число/дату, формат.\n"
        u"«Преобразование» = нет — столбец остаётся текстом даже при MERGE_XML_CONVERT_TO_NUMBERS=Да.\n"
        u"«да» — принудительно (без порога 3 строк); по формату — дата или число.\n"
        u"Формат — выпадающий список или свой NumberFormat Calc.\n"
        u"С * в обоих именах хвост сохраняется: Заказы_* → Сводные_заказы_*, "
        u"заказы_июнь → Сводные_заказы_июнь.\n"
        u"Без * — полная подстановка; при конфликте — суффикс _01, _02…\n"
        u"Legacy: ФИО;Табельный_номер;Старое->Новое. JSON по умолчанию."
    ),
    u"Начало_Строка": (
        u"Первая строка данных (1-based). Пусто, 0 или отрицательное — 1.\n"
        u"Сбор строк начинается с этой строки.\n"
        u"Примеры: 2   1"
    ),
    u"Начало_Столбец": (
        u"Первый столбец блока: номер (1) или буква (A).\n"
        u"Примеры: 1   A   B"
    ),
    u"Число_Строк": (
        u"Сколько строк данных копировать. Пусто — до пустой ячейки в «Начало_Столбец».\n"
        u"Примеры: (пусто)   10   100"
    ),
    u"Число_Столбцов": (
        u"Лимит столбцов при режиме «Все». Пусто — до пустого заголовка.\n"
        u"Примеры: (пусто)   6   12"
    ),
    u"Строка_Заголовков": (
        u"Строка с именами столбцов (1-based). Пусто — как «Начало_Строка».\n"
        u"0 или отрицательное — заголовка нет (Колонка_A, Колонка_B…).\n"
        u"Простой режим: цифра в B (как раньше).\n"
        u"Расширенный: B = «смотри в C», JSON в C (кнопка «Параметры…») — "
        u"многоэтажная шапка, диапазон строк 2-5 / 2:5 и разделитель.\n"
        u"Примеры: 2   1   0"
    ),
    u"Режим": (
        u"Способ объединения данных на листе результата.\n"
        u"Колонка B — режим сбора (выпадающий список).\n"
        u"Колонка C — способ переноса (выпадающий список; по умолчанию По_API).\n"
        u"«Текущие листы» — без копирования: существующие вкладки эта_книга → постобработка/финал;\n"
        u"предварительное удаление листов не выполняется; C при этом не используется.\n"
        u"Отдельная строка «Способ_переноса» имеет приоритет над C."
    ),
    u"Источник_в_первой_колонке": (
        u"Служебная колонка «#Путь».\n"
        u"Не_выводить / Коротко / Длинно.\n"
        u"Только для «На один лист» и «На разные листы»; в «Копирование листов» / «Текущие листы» не используется."
    ),
    u"Дата_источника_во_второй_колонке": (
        u"Служебная колонка «#ДатаВремяФайла».\n"
        u"Не_выводить / Коротко (дата) / Длинно (дата+время).\n"
        u"Только для «На один лист» и «На разные листы»; в «Копирование листов» / «Текущие листы» не используется."
    ),
    u"Дата_источника_во_второй__колонке": (
        u"(устаревшее имя параметра — используйте «Дата_источника_во_второй_колонке»)\n"
        u"Служебная колонка «#ДатаВремяФайла»."
    ),
    u"Способ_переноса": (
        u"Как переносить данные источника в результат.\n"
        u"Обычно задаётся в колонке C строки «Режим»; отдельная строка с этим именем "
        u"сохранена для совместимости и имеет приоритет над C.\n"
        u"Для «Копирование листов» поле C в визарде доступно так же, как для других режимов.\n"
        u"В режиме «Текущие листы» не используется (переноса нет).\n"
        u"По_API (по умолчанию) — UNO getDataArray/setDataArray (пакеты) + построчно при необходимости.\n"
        u"Буфер_по_заголовкам — .uno:Copy + спецвставка «только значения» по столбцам; "
        u"столбец результата по заголовку. Форматы — только при MERGE_PASTE_FORMATS=Да.\n"
        u"Если среди листов сбора книги-источника есть скрытый — буфер для этой книги "
        u"автоматически заменяется на По_API.\n"
        u"Буфер_по_позиции — копирование прямоугольника данных в буфер без матчинга заголовков "
        u"(тоже только значения; форматы — MERGE_PASTE_FORMATS); "
        u"строка заголовков копируется один раз при первой записи на лист.\n"
        u"xml_excel_ods — прямой перенос через работу с файлами (.xlsx/.xlsm/.xls/.ods): "
        u"книга результата сохраняется и закрывается, файл результата обновляется на диске и открывается снова."
    ),
    u"MERGE_PASTE_FORMATS": (
        u"Только для «Способ_переноса» = Буфер_по_заголовкам / Буфер_по_позиции.\n"
        u"Нет (по умолчанию) — спецвставка только значений (InsertContents SVD).\n"
        u"Да — поверх вставленных значений ещё спецвставка форматов (InsertContents T / HARDATTR|STYLES)."
    ),
    u"Папка_результата": (
        u"Каталог для файла результата в режиме xml_excel_ods без подмены открытой книги.\n"
        u"Абсолютный путь или относительно папки книги с параметрами.\n"
        u"Если не задано — используется системная временная папка.\n"
        u"Туда копируется книга как «имя-result.ext»; исходная книга остаётся открытой.\n"
        u"В конце показывается путь к готовому файлу."
    ),
    u"MERGE_XML_CONVERT_TO_NUMBERS": (
        u"Да (по умолчанию) — текст «700», «1 000 000,00», «01.02.2024» превращается в число/дату в результате.\n"
        u"При «Способ_переноса» = xml_excel_ods — при чтении источников из файла; "
        u"для «Копирование листов» — отдельный проход по скопированным листам.\n"
        u"При По_API / Буфер_по_заголовкам / Буфер_по_позиции — проход по листу результата после копирования.\n"
        u"Столбец обрабатывается, если ≥3 строк данных распознаны как дата/время или число "
        u"(пробел/апостроф между разрядами, дробная , или .).\n"
        u"Нет — сохранять как текст (в т.ч. даты-строки)."
    ),
    u"MERGE_XML_DATETIME_FORMATS": (
        u"Список допустимых форматов даты/времени для преобразования текста → date/datetime "
        u"(через запятую). Используется при MERGE_XML_CONVERT_TO_NUMBERS=Да "
        u"(xml_excel_ods и UNO-сбор).\n"
        u"Если пусто — базовые форматы (точка/дефис, ISO) плюс дополнения по локали LibreOffice "
        u"(ru: DD/MM, 01.фев.2024, 1 февраля 2024; en-US: MM/DD).\n"
        u"Допускаются кириллические токены: ДД.ММ.ГГГГ, ДД/ММ/ГГГГ, ДД.МММ.ГГГГ и т.д.\n"
        u"Month-only (ячейка «Январь»/«янв.»/January) — шаблон день.MMMM.год:\n"
        u"  день: 01 / FIRST_DAY / LAST_DAY; месяц: MMMM/ММММ/MMM; год: YEAR / NEXT_YEAR / 2027.\n"
        u"  Голый MMMM без дня и года — не конструктор.\n"
        u"Примеры:\n"
        u"  DD.MM.YYYY, DD.MM.YY, YYYY-MM-DD, YY-MM-DD\n"
        u"  DD.MM.YYYY HH:MM(:SS), DD.MM.YYYYTHH:MM(:SS)\n"
        u"  DD/MM/YYYY, DD.MMM.YYYY, D MMMM YYYY, MM/DD/YYYY\n"
        u"  FIRST_DAY.MMMM.YEAR, LAST_DAY.MMMM.NEXT_YEAR, 01.ММММ.2027"
    ),
    u"MERGE_SOURCE_VIEW_SUBFOLDERS": (
        u"Да (по умолчанию) — при маске в «Файлы-Источники» искать совпадения "
        u"в стартовой папке и во всех её подпапках.\n"
        u"Нет — только в каталоге из пути маски (без рекурсии).\n"
        u"Синонимы Да/Нет: yes/no, 1/0, true/false (без учёта регистра)."
    ),
    u"MERGE_SOURCE_EXCLUDE_MASK": (
        u"Дополнительные маски исключения для имён файлов (через запятую), "
        u"после основной маски отбора.\n"
        u"По умолчанию: --*\n"
        u"Файлы с именем, начинающимся на ~, пропускаются всегда.\n"
        u"Совпадение регистронезависимое.\n"
        u"Те же маски можно задать в «Доп_Параметры_Источника» (поле exclude_mask)\n"
        u"для столбца источника — они складываются с этим параметром.\n"
        u"Пример: --*,*.tmp,~$*"
    ),
    u"Листы_не_удалять": (
        u"Листы, не удаляемые при очистке книги результата перед сбором.\n"
        u"Колонка B — шаблоны через запятую (text_match): *Отчет*, Отчет*, Справка_макроса.\n"
        u"Пусто — список по умолчанию из макроса.\n"
        u"«*» / «Все» / «All» без фильтра в C — не удалять никакие листы (без запроса).\n"
        u"«*» / «Все» + filter в C — оставить только тех, для кого лямбда True;\n"
        u"остальных удалить; если есть что удалять — всегда диалог со списком.\n"
        u"Колонка C (кнопка «Параметры…») — опциональный JSON:\n"
        u'  {"v":1,"filter":"lambda name: sheet_age_days(name) is not None and sheet_age_days(name) <= 3"}\n'
        u"Лямбда получает имя листа из отобранных по B; True — сохранить. Код санируется.\n"
        u"Листы collect_params* / Параметры_Объединения* и _Визард_Списки защищены всегда.\n"
        u"В режиме «Текущие листы» предварительная очистка всегда отключена."
    ),
    u"Листы_удалить": (
        u"Безусловное удаление указанных листов перед сбором (в любых режимах).\n"
        u"Колонка B — шаблоны через запятую (text_match / плейсхолдеры): *Отчет*, Сводная*, Мусор.\n"
        u"Пусто — параметр не действует (ничего не удаляет).\n"
        u"«*» / «Все» / «All» — удалить все листы, кроме защищённых.\n"
        u"Колонка C (кнопка «Параметры…») — опциональный JSON-фильтр:\n"
        u'  {"v":1,"filter":"lambda name: sheet_age_days(name, \\"YYMMDD-HH-mm\\") is not None and sheet_age_days(name, \\"YYMMDD-HH-mm\\") > 3"}\n'
        u"sheet_age_days(name[, fmt]): fmt — токены YYYY/YY/MM/DD/HH/mm/SS (ищем в имени).\n"
        u"Без fmt — legacy YYYYMMDD. Лямбда: True — удалить; False — пропустить.\n"
        u"Не удаляются только: листы collect_params* / Параметры_Объединения* и листы-источники\n"
        u"(если источник — эта книга). Остальные совпавшие имена удаляются без запроса.\n"
        u"Работает и в режиме «Текущие листы», и при подавлении обычной очистки."
    ),
    u"Пропуск_строк_источника": (
        u"Не копировать «пустые» строки при сборе.\n"
        u"Колонка B — «смотри в С»; правила — JSON в C (кнопка «Параметры…»).\n"
        u"sheet — имя листа *источника* (не результата); пусто = все листы файла.\n"
        u"columns — номера/буквы/заголовки: строка копируется, если хотя бы один не пуст.\n"
        u"Файл источника не сохраняется (изменения только в скрытой копии).\n"
        u'Пример: [{"v":1,"fn":"пропуск_строк_источника","sheet":"Заказы","columns":["1"]}]'
    ),
    u"Автофильтр_результата": (
        u"Автофильтр на листе результата после сбора и постобработки.\n"
        u"Нет (по умолчанию) — не ставить автоматически.\n"
        u"Да — меню-автофильтр на листах результата (с учётом списка в C).\n"
        u"Колонка B — Да/Нет; список листов — JSON в колонке C (кнопка «Параметры…»).\n"
        u"Пустой JSON / все листы — на каждом листе результата.\n"
        u"Если в постобработке/финале уже есть «автофильтр» на лист — "
        u"параметр результата этот лист не трогает (не переключает фильтр повторно)."
    ),
    u"Закрепить_заголовок": (
        u"Закрепить строку заголовков (freeze panes) на листе результата.\n"
        u"Колонка B — «смотри в С»; настройки листов — JSON в колонке C (кнопка «Параметры…»).\n"
        u"Пустой JSON / все листы — на всех листах результата."
    ),
    u"Отображение_прогресса": (
        u"Как показывать прогресс на этапе 2.\n"
        u"Диалог (по умолчанию) или Статусная_строка.\n"
        u"Диалог: сводка + «Запуск», затем прогресс + «Прервать»."
    ),
    u"MERGE_DEBUG": (
        u"Подробный отладочный вывод в консоль (этап, память, таймер).\n"
        u"Да — включает debug-флаги во всех библиотеках макроса.\n"
        u"По умолчанию Нет."
    ),
    u"MERGE_DEBUG_CONSOLE_LOG_TO_WORKSHEET": (
        u"Писать на лист «Сбор_книг_лог» всё, что идёт в консоль (Время | Сообщение),\n"
        u"вместо обычного журнала сбора. Рекомендуется вместе с MERGE_DEBUG=Да.\n"
        u"По умолчанию Нет."
    ),
    u"MERGE_CREATE_LOG_SHEET": (
        u"Создавать лист журнала «Сбор_книг_лог» (и «Сбор_книг_лог_N» в оркестраторе).\n"
        u"Нет — журнал только в консоль; вывод MERGE_DEBUG в консоль не отключается.\n"
        u"По умолчанию Да."
    ),
    u"MERGE_DO_EMPTY_SHEETS": (
        u"Режим «на разные листы»: если по целевому листу к переносу 0 строк\n"
        u"(во всех источниках нет данных) — всё равно создать лист результата,\n"
        u"вставить заголовок и в первой data-строке в столбце A написать:\n"
        u"«данных для переноса не обнаружено».\n"
        u"Нет — такие пустые листы не оставлять (удалить после сбора).\n"
        u"По умолчанию Да. Можно задать глобально (вкладка «Глобальные») и\n"
        u"перекрыть на листе параметров (Да/Нет в B; пусто = global/default)."
    ),
    u"Ячейка_В_Столбец": (
        u"Вертикальный список (имя в A повторяется). B,C,D… — по файлам.\n"
        u"Legacy: «столбец:строка -> имя_колонки» (1-based).\n"
        u"JSON в колонке C (кнопка «Параметры…»):\n"
        u'  [{"col":"A","row":1,"column":"Филиал","sheet":"Merge",'
        u'"source_sheet":"Отчет*"}]\n'
        u"sheet — лист результата (пусто = все); source_sheet — фильтр листа "
        u"источника (*, список через запятую; пусто = все)."
    ),
    u"Добавить_в_карту_переменных": (
        u"Вертикальный список (имя в A повторяется). B — «смотри в С», JSON в C.\n"
        u"Читает ячейку источника и кладёт значение в runtime-карту переменных.\n"
        u"Поля блока: file, sheet, cell(A1), name.\n"
        u'Пример: [{"v":1,"fn":"добавить_в_карту_переменных","file":"*","sheet":"Заказы","cell":"A1","name":"Филиал"}]'
    ),
    u"Ручной_ввод_в_карту_переменных": (
        u"Вертикальный список. B — «смотри в С», JSON в C (кнопка «Параметры…»).\n"
        u"Визард: несколько переменных через «+ Блок» (одна строка, C = массив).\n"
        u"«Показать все блоки в одном диалоге» — при сборе одно окно, поля друг под другом;\n"
        u"подсказки (prompt, format) — только от первого блока. Иначе — по диалогу на имя.\n"
        u"Ключ: Имя~ручной_ввод~ручной_ввод. Плейсхолдер: <<Переменные.Имя>>.\n"
        u"Строка выше «Файлы-Источники» — диалог до разворота масок (можно в фильтре путей).\n"
        u"Поля: name, value_kind, data_type, format, choices, prompt, title, required, default,\n"
        u"remember_last (да), encrypt_last (шифр lm1: как пароли источника),\n"
        u"combined_dialog (на первом блоке) — один диалог для всех переменных.\n"
        u"Последние значения: ~/.config/libre-macros/collect_workbooks/<лист_параметров>.json\n"
        u'Пример: [{"v":1,"fn":"ручной_ввод_в_карту_переменных","combined_dialog":true,"name":"Филиал","choices":["Север","Юг"],"required":true},'
        u'{"v":1,"fn":"ручной_ввод_в_карту_переменных","name":"Источник_Дата","prompt":"YYYYMMDD","required":true}]'
    ),
    u"удаление_верхних_строк": (
        u"После сбора, до Постобработка_xml: удалить N верхних строк на листе результата.\n"
        u"Колонка B — «смотри в С»; правила — JSON в C (кнопка «Параметры…»).\n"
        u"Имя листа: как после копирования («1_Заказы») или без префикса («Заказы») — оба варианта.\n"
        u"Порядок относительно переименовать_лист не важен: удаление всегда раньше rename.\n"
        u"При явном правиле для листа автообрезка (Автообрезка_листов) для него не выполняется."
    ),
    u"Автообрезка_листов": (
        u"Только режим «Копирование листов».\n"
        u"Да (по умолчанию) — после копии обрезать строки до «Строка_Заголовков» "
        u"и промежуток до «Начало_Строка» (шапка оказывается в строке 1).\n"
        u"Нет — копировать лист целиком без автообрезки.\n"
        u"Явное «удаление_верхних_строк» работает независимо от этого флага."
    ),
    u"MERGE_CONFIRM_DELETION_OF_SHEETS": (
        u"Запрашивать подтверждение перед удалением «лишних» листов на этапе 1 сбора.\n"
        u"Нет (по умолчанию) — удалять без запроса; Да — диалог подтверждения."
    ),
    u"Постобработка_Диапазон": (
        u"Вертикальный список шагов постобработки диапазона (имя в A повторяется).\n"
        u"B — ключ функции из списка или «функция_плагин».\n"
        u"C — только JSON (кнопка «Параметры…»); D — rest плагина, если не в JSON.\n"
        u"Порядок шагов — сверху вниз на листе параметров.\n"
        u"Примеры C:\n"
        u'  [{"v":1,"fn":"тонкая_сетка"}]\n'
        u'  [{"v":1,"fn":"функция_плагин","ref":"functions_pp.py#pp_range_zagotovka","extra":"42","sheet":"Сводка"}]\n'
        u'  [{"v":1,"fn":"заголовок_плюс_высота","height_mm":13,"h_align":"center","v_align":"center"}]'
    ),
    u"Постобработка_Строка": (
        u"Вертикальный список шагов постобработки строки (имя в A повторяется).\n"
        u"B — ключ функции из списка или «функция_плагин».\n"
        u"C — JSON; D — rest плагина при необходимости.\n"
        u"Выполняется после всех шагов «Постобработка_Диапазон».\n"
        u'Пример: [{"v":1,"fn":"преобразовать_числа"}]'
    ),
    u"Постобработка_xml": (
        u"Постобработка в файле — только для «Способ_переноса» = xml_excel_ods.\n"
        u"Вертикальный список (имя в A повторяется); B — функция из отдельного каталога;\n"
        u"C — JSON (кнопка «Параметры…»). Выполняется в файле после сбора, до reopen UNO.\n"
        u"Часть шагов (авто_ширина/высота, заголовок_плюс_высота, автофильтр, "
        u"закрепить_заголовок) — UNO-добивка после открытия книги.\n"
        u"«применить_формулу»: as_values недоступен (формулы остаются в файле).\n"
        u"Обычная «Постобработка_Диапазон» при xml_excel_ods — после открытия книги (UNO).\n"
        u'Пример: [{"v":1,"fn":"применить_формулу","column":"Итого","formula":"=D{row_id}*0.1"}]'
    ),
    u"ВПР": (
        u"JOIN двух листов книги результата (между шагами Постобработка_Диапазон).\n"
        u"Основной путь — JSON в C; поддиалог задаёт left/right, join_type,\n"
        u"ключи, extract_columns, multi_match, цвет, not_found_fill, Дубли/Мульти…\n"
        u"В поддиалоге нажмите «Справка» — полное описание всех полей с примерами.\n"
        u"Кратко: left/inner/full; range A1 или used-area; _ПУСТО_ = пустая ячейка.\n"
        u"См. docs/06_PARAM_WIZARD.md, docs/13_JSON_PARAMS.md."
    ),
    u"Финальная_обработка": (
        u"Шаги по всей книге перед диалогом «Завершить».\n"
        u"B — ключ функции (удаление_листов, скрытие_листов, функция_плагин…; "
        u"регистр в B не важен).\n"
        u"C — JSON; D — rest плагина, если не в JSON.\n"
        u"Примеры C:\n"
        u'  [{"v":1,"fn":"удаление_листов"}]\n'
        u'  [{"v":1,"fn":"функция_плагин","ref":"functions_final.py#pp_range_zagotovka","sheet":"Сводка"}]\n'
        u'  [{"v":1,"fn":"заголовок_плюс_высота","height_mm":13,"sheet":"Сводная"}]'
    ),
}
P_MERGE_POSTPROCESS_RANGE = u"Постобработка_Диапазон"
P_MERGE_POSTPROCESS_ROW = u"Постобработка_Строка"
P_MERGE_POSTPROCESS_XML = u"Постобработка_xml"
P_MERGE_CELL_TO_COLUMN = u"Ячейка_В_Столбец"
P_MERGE_ADD_TO_VARIABLES_MAP = u"Добавить_в_карту_переменных"
P_MERGE_MANUAL_VARIABLE_INPUT = u"Ручной_ввод_в_карту_переменных"
_SCENARIO_01_COMMENT = (
    u"Нагрузочный тест ВПР (уровень low); копирование листов «Заказы» и "
    u"«Справочник_цен» из vlookup_data_low_old.xlsx (~180 строк, ключ "
    u"Код_позиции). UNO: Буфер_по_заголовкам + Постобработка_Диапазон "
    u"(переименовать_лист → ВПР → оформление)."
)
_WP_CREATE_DEFAULT_PRESET_NAME = u"— Шаблон по умолчанию (UNO+ВПР) —"
_WP_NEW_PARAM_SHEET_LABEL = u"＋ Новый лист…"
_WP_NEW_PARAM_SHEET_TITLE = u"Новый лист параметров"
_WP_NEW_PARAM_SHEET_PROMPT = (
    u"Имя нового листа параметров (префикс подставится, если его нет):"
)
_WP_PRESETS_FILE = u"merge_param_presets.json"
_WP_PRESETS_SUBDIR = u"collect_workbooks"
_WP_PRESETS_JSON_VER = 2
# В JSON пресетов/экспорта вместо абсолютного префикса листа параметров.
_WP_PRESET_SHEET_PREFIX_TOKEN = u"[prefix]"
# Примеры при пустом файле: только UNO-сценарии (без xml_excel_ods-клонов).
_WP_EXAMPLE_PRESET_COUNT = 16
_WP_SCENARIO_COUNT = 16
_PARAM_VALUE_COLUMNS = 5
_PARAM_FONT_NAME = u"PT Sans"
_PARAM_FONT_SIZE = 11
_PARAM_FILL_HEADER = 0xDCE4EE
_PARAM_FILL_NAME = 0xF0F1F3
_PARAM_BORDER_COLOR = 0xB4B4B4
_PARAM_FILL_DISABLED = 0x666666
_PARAM_B_SEE_COLUMN_C = u"смотри в С"
# Субвизард Доп_Параметры_Источника: (код разделителя, подпись)
_SOURCE_EXTRA_DELIM_CHOICES = (
    (u";", u"Точка с запятой (;)"),
    (u",", u"Запятая (,)"),
    (u"\\t", u"Табуляция"),
    (u"|", u"Вертикальная черта (|)"),
)
# (код encoding для JSON, подпись)
_SOURCE_EXTRA_ENCODING_CHOICES = (
    (u"windows-1251", u"Windows-1251 (кириллица, по умолчанию)"),
    (u"utf-8", u"UTF-8"),
    (u"utf-8-sig", u"UTF-8 с BOM"),
    (u"utf-16", u"UTF-16 (BOM)"),
    (u"utf-16le", u"UTF-16LE (BOM)"),
    (u"utf-16be", u"UTF-16BE (BOM)"),
    (u"windows-1252", u"Windows-1252 (латиница)"),
    (u"cp866", u"CP866 (DOS)"),
    (u"koi8-r", u"KOI8-R"),
    (u"iso-8859-5", u"ISO-8859-5"),
)
# CSV → temp XLSX (open_as_xlsx); по умолчанию «да»
_SOURCE_EXTRA_OPEN_AS_XLSX_CHOICES = (
    (u"нет", u"Нет — CSV матрица в памяти"),
    (u"да", u"Да — CSV → temp XLSX (по умолчанию)"),
)
# «При копировании переносить всем листом» (copy_whole_sheet)
_SOURCE_EXTRA_WHOLE_SHEET_CHOICES = (
    (u"нет", u"Нет — матрица в памяти → setDataArray (по умолчанию)"),
    (u"да", u"Да — открыть CSV в Calc и скопировать лист целиком"),
)
# Субвизард book-extra (ODS / XLS / XLSX / XLSM)
try:
    from libre_macros_source_extra_lib import SOURCE_EXTRA_DEFAULT_CHUNK_ROWS
except Exception:
    SOURCE_EXTRA_DEFAULT_CHUNK_ROWS = 50000
_SOURCE_EXTRA_PREPARE_IN_MEMORY_CHOICES = (
    (u"нет", u"Нет — открыть в Calc (UNO, по умолчанию)"),
    (u"да", u"Да — бандл → матрица → setDataArray"),
)
_SOURCE_EXTRA_VALUES_MODE_CHOICES = (
    (u"cached", u"Кэш / вычисленные (по умолчанию)"),
    (u"formulas", u"Текст формул"),
    (u"prefer_cached", u"Кэш, иначе формула"),
)
_SOURCE_EXTRA_MERGED_CELLS_CHOICES = (
    (u"top_left", u"Только левая верхняя (как Calc)"),
    (u"fill", u"Размножить на всё объединение"),
)
_SOURCE_EXTRA_TRIM_EMPTY_CHOICES = (
    (u"да", u"Да — обрезать пустые края (по умолчанию)"),
    (u"нет", u"Нет — весь used-range"),
)
_SOURCE_EXTRA_CONVERT_TYPES_CHOICES = (
    (u"in_memory", u"В памяти (по умолчанию)"),
    (u"uno", u"Оставить для UNO"),
    (u"no", u"Не конвертировать"),
    (u"keep_text", u"Всё текстом в памяти"),
)
_SOURCE_EXTRA_CHUNK_ROWS_CHOICES = (
    (u"10000", u"10 000"),
    (u"25000", u"25 000"),
    (u"50000", u"50 000 (по умолчанию)"),
    (u"100000", u"100 000"),
)
_SOURCE_EXTRA_SKIP_MISSING_SHEETS_CHOICES = (
    (u"нет", u"Нет — ошибка, если лист из «Листы» не найден (по умолчанию)"),
    (u"да", u"Да — пропустить отсутствующие, вовлечь найденные"),
)
_SOURCE_EXTRA_FILE_PICK_CHOICES = (
    (u"all", u"Все (умолчание)"),
    (u"newest_ctime", u"Новый по дате создания"),
    (u"oldest_ctime", u"Старый по дате создания"),
    (u"newest_mtime", u"Новый по дате модификации"),
    (u"oldest_mtime", u"Старый по дате модификации"),
    (u"lambda", u"Лямбда-фильтр"),
)
_SOURCE_EXTRA_FILE_PICK_SCOPE_CHOICES = (
    (u"all", u"Все подкаталоги вместе"),
    (u"per_subdir", u"В рамках подкаталога"),
)

# Справка субвизарда «Доп_Параметры_Источника» (кнопка «Справка»).
_SOURCE_EXTRA_HELP_TITLE = u"Справка — доп. параметры источника"

# Общий заголовок кнопки «Справка» во визардах / подвизардах.
_WIZARD_HELP_BTN_LABEL = u"Справка"
_WIZARD_HELP_DIALOG_TITLE = u"Справка"
_WIZARD_HELP_EMPTY = (
    u"Справка для этого диалога пока не заполнена.\n"
    u"См. docs/06_PARAM_WIZARD.md и docs/15_PARAM_REFERENCE.md."
)
_WIZARD_MAIN_HELP_TITLE = u"Справка — визард параметров"
_WIZARD_MAIN_HELP_TEXT = (
    u"Визард параметров листа «Параметры_Объединения*».\n"
    u"\n"
    u"• Выберите параметр в списке — справа описание и поля значения.\n"
    u"• «Параметры…» / субвизарды — форма JSON для колонки C (и аналогов).\n"
    u"• Во всех субвизардах кнопка «Справка» — подробности полей.\n"
    u"• «Файлы-Источники»: пути, маски, <<Переменные.Имя>> из карты.\n"
    u"  Ручной ввод выше строки файлов — значение доступно в путях.\n"
    u"• Пресеты — сохранить/применить набор строк параметров.\n"
    u"• Глобальные — переменные и флаги MERGE_* на пользователя.\n"
    u"\n"
    u"Подробнее: docs/06_PARAM_WIZARD.md, docs/04_DATA_COLLECTION.md,\n"
    u"docs/15_PARAM_REFERENCE.md."
)

_SOURCE_EXTRA_HELP_FILE_PICK = (
    u"══════════════════════════════════════\n"
    u"ДОП. ФИЛЬТР ФАЙЛОВ (file_pick)\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u"Порядок отбора файлов:\n"
    u"  1) путь/маска из «Файлы-Источники» (+ recurse при необходимости);\n"
    u"  2) маски исключения (exclude_mask + MERGE_SOURCE_EXCLUDE_MASK);\n"
    u"  3) доп. фильтр file_pick / file_pick_scope / file_pick_lambda;\n"
    u"  4) дальше — чтение CSV / XML / книги как обычно.\n"
    u"\n"
    u"Режимы «Доп. фильтр файлов»:\n"
    u"  • Все (all) — ничего не сужать (умолчание). Область и лямбда игнорируются.\n"
    u"  • Новый/старый по дате создания (newest_ctime / oldest_ctime) —\n"
    u"    оставить ОДИН файл с экстремумом ctime (см. оговорку ОС ниже).\n"
    u"  • Новый/старый по дате модификации (newest_mtime / oldest_mtime) —\n"
    u"    оставить ОДИН файл с экстремумом mtime.\n"
    u"  • Лямбда-фильтр (lambda) — оставить N файлов, для которых выражение\n"
    u"    истинно (это фильтр списка, а не «выбор одного победителя»).\n"
    u"\n"
    u"Область фильтра (file_pick_scope), если режим не «Все»:\n"
    u"  • Все подкаталоги вместе (all) — экстремум / лямбда по всему списку.\n"
    u"  • В рамках подкаталога (per_subdir) — группировка по родительской\n"
    u"    папке файла; в каждой группе свой экстремум (date) или свой прогон\n"
    u"    лямбды (lambda). Файлы прямо в корне источника — отдельная группа.\n"
    u"\n"
    u"Дата создания / ctime:\n"
    u"  На Windows обычно доступна дата создания файла.\n"
    u"  На Linux st_ctime — смена метаданных inode, не «дата создания».\n"
    u"  Если ctime недоступен — используется mtime, в журнал пишется пометка.\n"
    u"  Для надёжного «самого нового файла» предпочитайте newest_mtime.\n"
    u"\n"
    u"══════════════════════════════════════\n"
    u"ЛЯМБДА-ФИЛЬТР (file_pick_lambda) — подробно\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u"Когда активен: только при режиме «Лямбда-фильтр».\n"
    u"Результат: truthy → файл остаётся; иначе отбрасывается.\n"
    u"При per_subdir лямбда применяется ВНУТРИ каждой группы папки\n"
    u"(фильтрует список группы, не выбирает одного «победителя»).\n"
    u"\n"
    u"Синтаксис — как у других лямбд проекта. Предпочтительные формы:\n"
    u"\n"
    u"  1) Один аргумент-словарь (рекомендуется):\n"
    u"     lambda f: \"отчёт\" in (f.get(\"имя\") or \"\").casefold()\n"
    u"\n"
    u"  2) Именованные поля / **kw (тоже поддерживается):\n"
    u"     lambda path, имя, mtime, **kw: mtime > 0 and имя.endswith(\".csv\")\n"
    u"\n"
    u"  3) def fn(f): ...  (тело с return True/False)\n"
    u"\n"
    u"Поля контекста файла (ключи словаря f и/или имена параметров):\n"
    u"  путь / path     — полный путь к файлу\n"
    u"  имя / name      — имя файла (basename)\n"
    u"  каталог / dir   — каталог (dirname)\n"
    u"  mtime           — unix-время модификации (float/int)\n"
    u"  ctime           — unix-время «создания»/ctime\n"
    u"  размер / size   — размер в байтах\n"
    u"\n"
    u"Доступ к полям:\n"
    u"  f[\"имя\"], f.get(\"имя\"), f.get(\"name\") — равнозначные алиасы.\n"
    u"  Сравнивайте строки через casefold()/lower() для регистра.\n"
    u"\n"
    u"Подстановка переменных карты ДО вычисления:\n"
    u"  <<Переменные.Имя>>  — текст из runtime-карты переменных.\n"
    u"  В том числе значения из «Ручной_ввод_в_карту_переменных»,\n"
    u"  если диалог ручного ввода уже отработал (параметр стоит ВЫШЕ\n"
    u"  строки «Файлы-Источники» — тогда плейсхолдер доступен при развороте).\n"
    u"  Пример:\n"
    u"    lambda f: \"<<Переменные.Филиал>>\".casefold() in (f.get(\"имя\") or \"\").casefold()\n"
    u"  (после подстановки станет, например: lambda f: \"север\".casefold() in …)\n"
    u"\n"
    u"Примеры:\n"
    u"  # имя содержит «отчёт» без учёта регистра\n"
    u"  lambda f: \"отчёт\" in (f.get(\"имя\") or \"\").casefold()\n"
    u"\n"
    u"  # только CSV новее условной даты (unix time)\n"
    u"  lambda f: (f.get(\"имя\") or \"\").endswith(\".csv\") and (f.get(\"mtime\") or 0) > 1700000000\n"
    u"\n"
    u"  # файлы крупнее 1 МБ в папках, где путь содержит «архив»\n"
    u"  lambda f: (f.get(\"размер\") or 0) > 1048576 and \"архив\" in (f.get(\"путь\") or \"\").casefold()\n"
    u"\n"
    u"  # отбор по подстроке из ручного ввода / карты\n"
    u"  lambda f: \"<<Переменные.Ключ>>\" in (f.get(\"имя\") or \"\")\n"
    u"\n"
    u"Ограничения:\n"
    u"  • код проходит санацию (libre_macros_sanitize_lib) — без import,\n"
    u"    открытых файлов, сетевых вызовов и опасных builtins;\n"
    u"  • не смешивайте date-режим и lambda в одном прогоне: либо один\n"
    u"    победитель по дате, либо фильтр N файлов лямбдой;\n"
    u"  • пустой список после фильтра → как при пустой маске: запись в журнал,\n"
    u"    без «падения» всего сбора (если так принято для данного запуска).\n"
)

_SOURCE_EXTRA_HELP_BOOK = (
    u"Справка: доп. параметры источника — КНИГА (.ods / .xls / .xlsx / .xlsm)\n"
    u"\n"
    u"JSON пишется в колонку выбранного источника (B/C/D…) строки\n"
    u"«Доп_Параметры_Источника». Пустая ячейка у столбца → берётся JSON из B.\n"
    u"\n"
    u"—— Подготовка источника в памяти ——\n"
    u"prepare_source_in_memory: Да — книга разбирается во временную копию\n"
    u"в памяти (удобно для значений/объединённых ячеек/обрезки). Нет —\n"
    u"обычное открытие источника.\n"
    u"\n"
    u"—— Пароль снятия защиты ——\n"
    u"sheet_unprotect_password: если листы защищены — пароль для снятия\n"
    u"защиты при чтении. Пусто — без пароля. В JSON хранится зашифрованным\n"
    u"(lm1:); ключ — глобальная переменная encrypt_passwords_default.\n"
    u"\n"
    u"—— Значения ячеек ——\n"
    u"values_mode: cached / … — как читать значения при подготовке в памяти\n"
    u"(см. подписи combo в форме).\n"
    u"\n"
    u"—— Колонка последней строки ——\n"
    u"last_row_column: буква (B), номер (2) или имя заголовка. Сначала\n"
    u"нижняя граница ищется как обычно, затем уточняется по этой колонке,\n"
    u"чтобы не тянуть «хвост» из предзаполненных столбцов.\n"
    u"\n"
    u"—— Объединённые ячейки ——\n"
    u"merged_cells: как разворачивать merge при подготовке (top_left и др.).\n"
    u"\n"
    u"—— Обрезать пустые края ——\n"
    u"trim_empty: убрать пустые края области данных при подготовке.\n"
    u"\n"
    u"—— Числа и даты ——\n"
    u"convert_types: конвертация чисел/дат при подготовке в памяти.\n"
    u"\n"
    u"—— Строк в чанке памяти ——\n"
    u"chunk_rows: размер чанка при разборе больших книг (по умолчанию 50000).\n"
    u"\n"
    u"—— Пропускать отсутствующие листы ——\n"
    u"skip_missing_sheets: для режима «Текущие листы» / вовлечения —\n"
    u"пропустить листы из «Листы», которых нет в источнике, вместо ошибки.\n"
    u"\n"
    u"—— Маски исключения файлов ——\n"
    u"exclude_mask: через запятую, с «*» (например --*,*.tmp). Складываются\n"
    u"с параметром листа MERGE_SOURCE_EXCLUDE_MASK. Действуют и для glob,\n"
    u"и для явно указанных путей.\n"
    u"\n"
    + _SOURCE_EXTRA_HELP_FILE_PICK
)

_SOURCE_EXTRA_HELP_CSV = (
    u"Справка: доп. параметры источника — CSV\n"
    u"\n"
    u"JSON пишется в колонку выбранного источника строки\n"
    u"«Доп_Параметры_Источника». Пустая ячейка у столбца → JSON из B.\n"
    u"Если вся строка пуста — разделитель «;», кодировка Windows-1251.\n"
    u"\n"
    u"—— Разделитель ——\n"
    u"delimiter: ; , \\t | — символ разделения полей CSV.\n"
    u"\n"
    u"—— Кодировка ——\n"
    u"encoding: windows-1251 (умолч.), utf-8, utf-8-sig, utf-16le, cp866,\n"
    u"koi8-r и др. Неверная кодировка даёт «кракозябры» в заголовках/данных.\n"
    u"\n"
    u"—— Копировать лист целиком ——\n"
    u"copy_whole_sheet: true — в режиме «Копирование листов» открыть CSV в Calc\n"
    u"и перенести лист целиком (importSheet), без матрицы/setDataArray.\n"
    u"При true отдельная конвертация чисел/дат не делается (Calc уже распознал).\n"
    u"\n"
    u"—— Открыть как XLSX (теневая конвертация) ——\n"
    u"open_as_xlsx: true (умолч.) — CSV конвертируется во временный .xlsx и\n"
    u"дальше идёт как обычная книга (листы по имени файла). false — матрица\n"
    u"в памяти. При MERGE_DEBUG временный xlsx может не удаляться; путь — в журнал.\n"
    u"\n"
    u"—— Маски исключения файлов ——\n"
    u"exclude_mask: через запятую, с «*». Складываются с MERGE_SOURCE_EXCLUDE_MASK.\n"
    u"\n"
    u"—— Колонка анализа последней строки ——\n"
    u"last_row_column: буква / номер / имя заголовка — уточнить нижнюю границу\n"
    u"данных по выбранной колонке.\n"
    u"\n"
    + _SOURCE_EXTRA_HELP_FILE_PICK
)

_SOURCE_EXTRA_HELP_XML = (
    u"Справка: доп. параметры источника — XML\n"
    u"\n"
    u"Для XML-источников в этом визарде задаются только отбор файлов:\n"
    u"маски исключения и доп. фильтр (даты / лямбда).\n"
    u"Разбор XML и маппинг полей — по остальным параметрам сбора.\n"
    u"\n"
    u"—— Маски исключения файлов ——\n"
    u"exclude_mask: через запятую, с «*». Складываются с MERGE_SOURCE_EXCLUDE_MASK.\n"
    u"Действуют после include-маски пути из «Файлы-Источники».\n"
    u"\n"
    + _SOURCE_EXTRA_HELP_FILE_PICK
)

_SOURCE_EXTRA_HELP_CURRENT_BOOK = (
    u"Справка: доп. параметры — Эта_книга\n"
    u"\n"
    u"Источник уже открыт (текущая книга). Подготовка «в памяти» и exclude/\n"
    u"file_pick здесь не применяются — книга одна и уже загружена.\n"
    u"\n"
    u"—— Пропускать отсутствующие листы ——\n"
    u"skip_missing_sheets: если в «Листы» указаны имена, которых нет в книге —\n"
    u"пропустить их вместо ошибки (режим «Текущие листы»).\n"
    u"\n"
    u"—— Пароль снятия защиты ——\n"
    u"sheet_unprotect_password: пароль для снятия защиты листов при чтении.\n"
    u"Пусто — без пароля. В JSON хранится зашифрованным (lm1:).\n"
    u"\n"
    u"—— Колонка анализа последней строки ——\n"
    u"last_row_column: буква (B), номер (2) или имя заголовка — уточнить\n"
    u"нижнюю границу данных по выбранной колонке.\n"
)

MANUAL_VAR_VALUE_KIND_CHOICES = (
    (u"scalar", u"Одно значение"),
    (u"list", u"Список"),
)
MANUAL_VAR_DATA_TYPE_CHOICES = (
    (u"text", u"Текст"),
    (u"number", u"Число"),
    (u"date", u"Дата"),
    (u"datetime", u"Дата-время"),
    (u"bool", u"Логическое"),
)
MANUAL_VAR_VALUE_KIND_LABEL = {c: l for c, l in MANUAL_VAR_VALUE_KIND_CHOICES}
MANUAL_VAR_VALUE_KIND_CODE = {l.casefold(): c for c, l in MANUAL_VAR_VALUE_KIND_CHOICES}
for _vk_code, _vk_lab in MANUAL_VAR_VALUE_KIND_CHOICES:
    MANUAL_VAR_VALUE_KIND_CODE[_vk_code.casefold()] = _vk_code
MANUAL_VAR_DATA_TYPE_LABEL = {c: l for c, l in MANUAL_VAR_DATA_TYPE_CHOICES}
MANUAL_VAR_DATA_TYPE_CODE = {l.casefold(): c for c, l in MANUAL_VAR_DATA_TYPE_CHOICES}
for _dt_code, _dt_lab in MANUAL_VAR_DATA_TYPE_CHOICES:
    MANUAL_VAR_DATA_TYPE_CODE[_dt_code.casefold()] = _dt_code
ADD_COLUMN_DATA_TYPE_CHOICES = (
    (u"text", u"Текст"),
    (u"number", u"Число"),
    (u"date", u"Дата"),
)
ADD_COLUMN_DATA_TYPE_LABEL = {c: l for c, l in ADD_COLUMN_DATA_TYPE_CHOICES}
ADD_COLUMN_DATA_TYPE_CODE = {l.casefold(): c for c, l in ADD_COLUMN_DATA_TYPE_CHOICES}
for _ac_code, _ac_lab in ADD_COLUMN_DATA_TYPE_CHOICES:
    ADD_COLUMN_DATA_TYPE_CODE[_ac_code.casefold()] = _ac_code
_PARAM_COL_WIDTHS = (8500, 9500, 9500, 9500, 9500, 9500)
_DESC_GENERAL_INFO_SEPARATOR = u"---------Общая информация ---------"
_PICK_SOURCE = None
_INNER_BTN_W = 80
# Высота OK/Отмена/Справка в футере субвизардов (+40% к прежним 16).
_INNER_BTN_H = 22
_INNER_BTN_GAP = 10
_INNER_CHK_H = 14
_INNER_FOOTER_RESERVE = _INNER_BTN_H + 4 + _INNER_CHK_H + 6
_INNER_GRID_DH = 420
_INNER_GRID_CONTROLS_H = 56
_INNER_GRID_LIST_GAP = 8
_INNER_LBL_W = 86
_INNER_LBL_GAP = 6
_INNER_FIELD_ROW_H = 22
_INNER_BLOCK_ALL_SHEETS_LABEL = u"(все_листы)"
_SHEET_LIST_PANEL_FNS = (
    u"тонкая_сетка",
    u"толстая_сетка",
    u"авто_высота",
    u"авто_ширина",
    u"левое_выравнивание",
    u"сброс_стрипов",
    u"закрепить_заголовок",
    u"автофильтр",
    u"вертикаль_центр",
    u"только_значения",
    u"удаление_листов",
    u"скрытие_листов",
    u"серые_служебные",
    u"полосы_по_пути",
    u"чередующиеся_границы",
    u"копировать_формат_заголовка",
    u"подсветка_по_заголовку",
    u"жирный_по_пути",
)
_SHEET_LIST_PANEL_FN_CF = frozenset(x.casefold() for x in _SHEET_LIST_PANEL_FNS)
_SHEET_LIST_SPECIAL_FNS = (u"удаление_листов", u"скрытие_листов")
_SHEET_LIST_SPECIAL_TOKENS = (
    u"_Первый_",
    u"_Последний_",
    u"_Последний_мусорный_",
)
_HEADER_PLUS_HEIGHT_HORI_CHOICES = (u"left", u"center", u"right")
_HEADER_PLUS_HEIGHT_VERT_CHOICES = (u"top", u"center", u"bottom")
_SHEET_BLOCK_FIELD_GAP = 5
def _header_plus_height_color_choices():
    """Именованные пресеты + hex для combo заливки/шрифта."""
    out = [u""]
    seen = set()
    try:
        from libre_macros_lib import _LM_PP_ZEBRA_COLOR_ALIASES

        keys = sorted(_LM_PP_ZEBRA_COLOR_ALIASES.keys(), key=lambda x: unicode(x).casefold())
        ki = 0
        while ki < len(keys):
            key = unicode(keys[ki])
            rgb = _LM_PP_ZEBRA_COLOR_ALIASES[key]
            sig = tuple(rgb)
            if sig not in seen:
                seen.add(sig)
                out.append(key)
            ki = ki + 1
    except Exception:
        pass
    for hx in (u"#FFFFFF", u"#F2F2F2", u"#1F4E79", u"#4472C4", u"#FFC7CE"):
        if hx not in out:
            out.append(hx)
    return tuple(out)

DEDUP_KEEP_CHOICES = (
    (u"first", u"Первую"),
    (u"last", u"Последнюю"),
)
DEDUP_KEEP_LABEL = {code: label for code, label in DEDUP_KEEP_CHOICES}
DEDUP_KEEP_CODE = {label.casefold(): code for code, label in DEDUP_KEEP_CHOICES}
for _dedup_keep_code, _dedup_keep_label in DEDUP_KEEP_CHOICES:
    DEDUP_KEEP_CODE[_dedup_keep_code.casefold()] = _dedup_keep_code

# замена_значений: отбор среди совпавших строк столбца (odd/even — 1-based).
REPLACE_MATCH_PICK_CHOICES = (
    (u"all", u"Все совпадения"),
    (u"first", u"Первое"),
    (u"last", u"Последнее"),
    (u"odd", u"Нечётные (1, 3, …)"),
    (u"even", u"Чётные (2, 4, …)"),
    (u"lambda", u"Лямбда"),
)
REPLACE_MATCH_PICK_LABEL = {code: label for code, label in REPLACE_MATCH_PICK_CHOICES}
REPLACE_MATCH_PICK_CODE = {
    label.casefold(): code for code, label in REPLACE_MATCH_PICK_CHOICES
}
for _rmp_code, _rmp_label in REPLACE_MATCH_PICK_CHOICES:
    REPLACE_MATCH_PICK_CODE[_rmp_code.casefold()] = _rmp_code
REPLACE_MATCH_PICK_CODE[u"все"] = u"all"
REPLACE_MATCH_PICK_CODE[u"первый"] = u"first"
REPLACE_MATCH_PICK_CODE[u"первое"] = u"first"
REPLACE_MATCH_PICK_CODE[u"последний"] = u"last"
REPLACE_MATCH_PICK_CODE[u"последнее"] = u"last"
REPLACE_MATCH_PICK_CODE[u"нечетный"] = u"odd"
REPLACE_MATCH_PICK_CODE[u"нечётный"] = u"odd"
REPLACE_MATCH_PICK_CODE[u"нечетные"] = u"odd"
REPLACE_MATCH_PICK_CODE[u"нечётные"] = u"odd"
REPLACE_MATCH_PICK_CODE[u"четный"] = u"even"
REPLACE_MATCH_PICK_CODE[u"чётный"] = u"even"
REPLACE_MATCH_PICK_CODE[u"четные"] = u"even"
REPLACE_MATCH_PICK_CODE[u"чётные"] = u"even"
REPLACE_MATCH_PICK_CODE[u"лямбда"] = u"lambda"

# замена_значений: режим поля «Заменить»
REPLACE_MODE_CHOICES = (
    (u"manual", u"Вручную"),
    (u"lambda", u"Лямбда"),
)
REPLACE_MODE_LABEL = {code: label for code, label in REPLACE_MODE_CHOICES}
REPLACE_MODE_CODE = {label.casefold(): code for code, label in REPLACE_MODE_CHOICES}
for _rpl_mode_code, _rpl_mode_label in REPLACE_MODE_CHOICES:
    REPLACE_MODE_CODE[_rpl_mode_code.casefold()] = _rpl_mode_code
REPLACE_MODE_CODE[u"вручную"] = u"manual"
REPLACE_MODE_CODE[u"ручной"] = u"manual"
REPLACE_MODE_CODE[u"ручной ввод"] = u"manual"
REPLACE_MODE_CODE[u"текст"] = u"manual"
REPLACE_MODE_CODE[u"лямбда"] = u"lambda"
REPLACE_MODE_CODE[u"expr"] = u"lambda"
REPLACE_MODE_CODE[u"выражение"] = u"lambda"

COPY_RANGES_MODE_CHOICES = (
    (u"replace", u"Замена"),
    (u"insert", u"Вставка"),
)
COPY_RANGES_MODE_LABEL = {code: label for code, label in COPY_RANGES_MODE_CHOICES}
COPY_RANGES_MODE_CODE = {label.casefold(): code for code, label in COPY_RANGES_MODE_CHOICES}
for _crm_code, _crm_label in COPY_RANGES_MODE_CHOICES:
    COPY_RANGES_MODE_CODE[_crm_code.casefold()] = _crm_code
COPY_RANGES_MODE_CODE[u"overwrite"] = u"replace"
COPY_RANGES_MODE_CODE[u"замена"] = u"replace"
COPY_RANGES_MODE_CODE[u"сдвиг"] = u"insert"
COPY_RANGES_MODE_CODE[u"добавление"] = u"insert"
COPY_RANGES_MODE_CODE[u"append"] = u"insert"
COPY_RANGES_MODE_CODE[u"вставка со сдвигом"] = u"insert"
COPY_RANGES_MODE_CODE[u"вставка"] = u"insert"

COPY_RANGES_AXIS_CHOICES = (
    (u"auto", u"Авто"),
    (u"rows", u"Строки"),
    (u"cols", u"Столбцы"),
)
COPY_RANGES_AXIS_LABEL = {code: label for code, label in COPY_RANGES_AXIS_CHOICES}
COPY_RANGES_AXIS_CODE = {label.casefold(): code for code, label in COPY_RANGES_AXIS_CHOICES}
for _cra_code, _cra_label in COPY_RANGES_AXIS_CHOICES:
    COPY_RANGES_AXIS_CODE[_cra_code.casefold()] = _cra_code
COPY_RANGES_AXIS_CODE[u"row"] = u"rows"
COPY_RANGES_AXIS_CODE[u"строка"] = u"rows"
COPY_RANGES_AXIS_CODE[u"columns"] = u"cols"
COPY_RANGES_AXIS_CODE[u"col"] = u"cols"
COPY_RANGES_AXIS_CODE[u"столбец"] = u"cols"

COPY_RANGES_CONTENT_CHOICES = (
    (u"values", u"Значения"),
    (u"formulas", u"Формулы"),
)
COPY_RANGES_CONTENT_LABEL = {code: label for code, label in COPY_RANGES_CONTENT_CHOICES}
COPY_RANGES_CONTENT_CODE = {
    label.casefold(): code for code, label in COPY_RANGES_CONTENT_CHOICES
}
for _crc_code, _crc_label in COPY_RANGES_CONTENT_CHOICES:
    COPY_RANGES_CONTENT_CODE[_crc_code.casefold()] = _crc_code
COPY_RANGES_CONTENT_CODE[u"value"] = u"values"
COPY_RANGES_CONTENT_CODE[u"значения"] = u"values"
COPY_RANGES_CONTENT_CODE[u"значения (без формул)"] = u"values"
COPY_RANGES_CONTENT_CODE[u"formula"] = u"formulas"
COPY_RANGES_CONTENT_CODE[u"формулы"] = u"formulas"
COPY_RANGES_CONTENT_CODE[u"as_values"] = u"values"
COPY_RANGES_CONTENT_CODE[u"keep_formulas"] = u"formulas"

SEND_MAIL_ATTACH_CHOICES = (
    (u"whole_book", u"Всю книгу"),
    (u"self", u"Файл книги (тест)"),
    (u"sheets", u"Выбранные листы"),
    (u"none", u"Без вложения"),
)
SEND_MAIL_ATTACH_LABEL = {code: label for code, label in SEND_MAIL_ATTACH_CHOICES}
SEND_MAIL_ATTACH_CODE = {label.casefold(): code for code, label in SEND_MAIL_ATTACH_CHOICES}
for _sma_code, _sma_label in SEND_MAIL_ATTACH_CHOICES:
    SEND_MAIL_ATTACH_CODE[_sma_code.casefold()] = _sma_code
SEND_MAIL_ATTACH_CODE[u"книга"] = u"whole_book"
SEND_MAIL_ATTACH_CODE[u"вся_книга"] = u"whole_book"
SEND_MAIL_ATTACH_CODE[u"листы"] = u"sheets"
SEND_MAIL_ATTACH_CODE[u"сама_книга"] = u"self"
SEND_MAIL_ATTACH_CODE[u"текущая_книга"] = u"self"
SEND_MAIL_ATTACH_CODE[u"файл_книги"] = u"self"

SEND_MAIL_FORMAT_CHOICES = (
    (u"ods", u"ODS"),
    (u"xlsx", u"XLSX"),
)
SEND_MAIL_FORMAT_LABEL = {code: label for code, label in SEND_MAIL_FORMAT_CHOICES}
SEND_MAIL_FORMAT_CODE = {label.casefold(): code for code, label in SEND_MAIL_FORMAT_CHOICES}
for _smf_code, _smf_label in SEND_MAIL_FORMAT_CHOICES:
    SEND_MAIL_FORMAT_CODE[_smf_code.casefold()] = _smf_code

SEND_MAIL_OS_CHOICES = (
    (u"auto", u"Авто"),
    (u"windows", u"Windows"),
    (u"linux", u"Linux"),
    (u"macos", u"macOS"),
)
SEND_MAIL_OS_LABEL = {code: label for code, label in SEND_MAIL_OS_CHOICES}
SEND_MAIL_OS_CODE = {label.casefold(): code for code, label in SEND_MAIL_OS_CHOICES}
for _smo_code, _smo_label in SEND_MAIL_OS_CHOICES:
    SEND_MAIL_OS_CODE[_smo_code.casefold()] = _smo_code

SEND_MAIL_CLIENT_CHOICES = (
    (u"auto", u"Авто"),
    (u"system", u"Системный (UNO)"),
    (u"outlook", u"Outlook"),
    (u"thunderbird", u"Thunderbird"),
)
SEND_MAIL_CLIENT_LABEL = {code: label for code, label in SEND_MAIL_CLIENT_CHOICES}
SEND_MAIL_CLIENT_CODE = {label.casefold(): code for code, label in SEND_MAIL_CLIENT_CHOICES}
for _smc_code, _smc_label in SEND_MAIL_CLIENT_CHOICES:
    SEND_MAIL_CLIENT_CODE[_smc_code.casefold()] = _smc_code
SEND_MAIL_CLIENT_CODE[u"системный"] = u"system"
SEND_MAIL_CLIENT_CODE[u"uno"] = u"system"
SEND_MAIL_CLIENT_CODE[u"mapi"] = u"system"

CREATE_SHEET_NAME_MODE_CHOICES = (
    (u"manual", u"Ввести вручную"),
    (u"expr", u"Вычислить"),
)
CREATE_SHEET_NAME_MODE_LABEL = {
    code: label for code, label in CREATE_SHEET_NAME_MODE_CHOICES
}
CREATE_SHEET_NAME_MODE_CODE = {
    label.casefold(): code for code, label in CREATE_SHEET_NAME_MODE_CHOICES
}
for _csn_code, _csn_label in CREATE_SHEET_NAME_MODE_CHOICES:
    CREATE_SHEET_NAME_MODE_CODE[_csn_code.casefold()] = _csn_code
CREATE_SHEET_NAME_MODE_CODE[u"lambda"] = u"expr"
CREATE_SHEET_NAME_MODE_CODE[u"compute"] = u"expr"
CREATE_SHEET_NAME_MODE_CODE[u"вычислить"] = u"expr"
CREATE_SHEET_NAME_MODE_CODE[u"лямбда"] = u"expr"
CREATE_SHEET_NAME_MODE_CODE[u"вручную"] = u"manual"
CREATE_SHEET_NAME_MODE_CODE[u"ввести вручную"] = u"manual"

CREATE_SHEET_DATA_MODE_CHOICES = (
    (u"none", u"Нет"),
    (u"sheet", u"С листа"),
    (u"rows", u"Вручную"),
)
CREATE_SHEET_DATA_MODE_LABEL = {
    code: label for code, label in CREATE_SHEET_DATA_MODE_CHOICES
}
CREATE_SHEET_DATA_MODE_CODE = {
    label.casefold(): code for code, label in CREATE_SHEET_DATA_MODE_CHOICES
}
for _csd_code, _csd_label in CREATE_SHEET_DATA_MODE_CHOICES:
    CREATE_SHEET_DATA_MODE_CODE[_csd_code.casefold()] = _csd_code
CREATE_SHEET_DATA_MODE_CODE[u"лист"] = u"sheet"
CREATE_SHEET_DATA_MODE_CODE[u"с листа"] = u"sheet"
CREATE_SHEET_DATA_MODE_CODE[u"range"] = u"sheet"
CREATE_SHEET_DATA_MODE_CODE[u"диапазон"] = u"sheet"
CREATE_SHEET_DATA_MODE_CODE[u"manual"] = u"rows"
CREATE_SHEET_DATA_MODE_CODE[u"вручную"] = u"rows"
CREATE_SHEET_DATA_MODE_CODE[u"текст"] = u"rows"
CREATE_SHEET_DATA_MODE_CODE[u"строки"] = u"rows"
CREATE_SHEET_DATA_MODE_CODE[u"пусто"] = u"none"
CREATE_SHEET_DATA_MODE_CODE[u"empty"] = u"none"

# Промежуточные итоги (группировка_по_столбцу): код → подпись; SUBTOTAL-номера Calc/Excel.
GROUP_AGG_FN_CHOICES = (
    (u"sum", u"Сумма"),
    (u"count", u"Количество"),
    (u"average", u"Среднее"),
    (u"max", u"Максимум"),
    (u"min", u"Минимум"),
    (u"product", u"Произведение"),
)
GROUP_AGG_FN_LABEL = {code: label for code, label in GROUP_AGG_FN_CHOICES}
GROUP_AGG_FN_CODE = {label.casefold(): code for code, label in GROUP_AGG_FN_CHOICES}
for _gaf_code, _gaf_label in GROUP_AGG_FN_CHOICES:
    GROUP_AGG_FN_CODE[_gaf_code.casefold()] = _gaf_code
GROUP_AGG_FN_CODE[u"сумма"] = u"sum"
GROUP_AGG_FN_CODE[u"кол-во"] = u"count"
GROUP_AGG_FN_CODE[u"количество"] = u"count"
GROUP_AGG_FN_CODE[u"среднее"] = u"average"
GROUP_AGG_FN_CODE[u"avg"] = u"average"
GROUP_AGG_FN_CODE[u"максимум"] = u"max"
GROUP_AGG_FN_CODE[u"минимум"] = u"min"
GROUP_AGG_FN_CODE[u"произведение"] = u"product"
GROUP_AGG_SUBTOTAL_NUM = {
    u"average": 1,
    u"count": 3,
    u"max": 4,
    u"min": 5,
    u"product": 6,
    u"sum": 9,
}
GROUP_AGG_GRAND_LABEL = {
    u"sum": u"Общая сумма",
    u"count": u"Общее количество",
    u"average": u"Общее среднее",
    u"max": u"Общий максимум",
    u"min": u"Общий минимум",
    u"product": u"Общее произведение",
}

DEDUP_OUTPUT_CHOICES = (
    (u"inplace", u"На месте (удалить)"),
    (u"new_sheet", u"Новый лист"),
    (u"offset", u"Смещение на этом листе"),
    (u"mark", u"Маркировка (столбец)"),
)
DEDUP_OUTPUT_LABEL = {code: label for code, label in DEDUP_OUTPUT_CHOICES}
DEDUP_OUTPUT_CODE = {label.casefold(): code for code, label in DEDUP_OUTPUT_CHOICES}
for _dedup_out_code, _dedup_out_label in DEDUP_OUTPUT_CHOICES:
    DEDUP_OUTPUT_CODE[_dedup_out_code.casefold()] = _dedup_out_code

UNPIVOT_OUTPUT_CHOICES = (
    (u"inplace", u"На месте"),
    (u"new_sheet", u"Новый лист"),
)
UNPIVOT_OUTPUT_LABEL = {code: label for code, label in UNPIVOT_OUTPUT_CHOICES}
UNPIVOT_OUTPUT_CODE = {label.casefold(): code for code, label in UNPIVOT_OUTPUT_CHOICES}
for _up_out_code, _up_out_label in UNPIVOT_OUTPUT_CHOICES:
    UNPIVOT_OUTPUT_CODE[_up_out_code.casefold()] = _up_out_code

GROUP_BY_ROWS_OP_CHOICES = (
    (u"sum", u"Сумма"),
    (u"count", u"Количество"),
    (u"min", u"Минимум"),
    (u"max", u"Максимум"),
    (u"avg", u"Среднее"),
    (u"first", u"Первое"),
    (u"last", u"Последнее"),
)
GROUP_BY_ROWS_OP_LABEL = {code: label for code, label in GROUP_BY_ROWS_OP_CHOICES}
GROUP_BY_ROWS_OP_CODE = {label.casefold(): code for code, label in GROUP_BY_ROWS_OP_CHOICES}
for _gbr_op_code, _gbr_op_label in GROUP_BY_ROWS_OP_CHOICES:
    GROUP_BY_ROWS_OP_CODE[_gbr_op_code.casefold()] = _gbr_op_code
GROUP_BY_ROWS_OP_CODE[u"сумма"] = u"sum"
GROUP_BY_ROWS_OP_CODE[u"кол-во"] = u"count"
GROUP_BY_ROWS_OP_CODE[u"количество"] = u"count"
GROUP_BY_ROWS_OP_CODE[u"минимум"] = u"min"
GROUP_BY_ROWS_OP_CODE[u"максимум"] = u"max"
GROUP_BY_ROWS_OP_CODE[u"среднее"] = u"avg"
GROUP_BY_ROWS_OP_CODE[u"average"] = u"avg"
GROUP_BY_ROWS_OP_CODE[u"первый"] = u"first"
GROUP_BY_ROWS_OP_CODE[u"последний"] = u"last"

GROUP_BY_ROWS_OUTPUT_CHOICES = UNPIVOT_OUTPUT_CHOICES
GROUP_BY_ROWS_OUTPUT_LABEL = UNPIVOT_OUTPUT_LABEL
GROUP_BY_ROWS_OUTPUT_CODE = UNPIVOT_OUTPUT_CODE

TRANSPOSE_OUTPUT_CHOICES = (
    (u"inplace", u"На месте"),
    (u"new_sheet", u"Новый лист"),
    (u"offset", u"Рядом (offset)"),
)
TRANSPOSE_OUTPUT_LABEL = {code: label for code, label in TRANSPOSE_OUTPUT_CHOICES}
TRANSPOSE_OUTPUT_CODE = {label.casefold(): code for code, label in TRANSPOSE_OUTPUT_CHOICES}
for _tr_out_code, _tr_out_label in TRANSPOSE_OUTPUT_CHOICES:
    TRANSPOSE_OUTPUT_CODE[_tr_out_code.casefold()] = _tr_out_code

FORM_TABLE_OUTPUT_CHOICES = (
    (u"new_sheet", u"Новый лист"),
    (u"inplace", u"На месте"),
    (u"replace_sheet", u"Заменить лист"),
)
FORM_TABLE_OUTPUT_LABEL = {code: label for code, label in FORM_TABLE_OUTPUT_CHOICES}
FORM_TABLE_OUTPUT_CODE = {label.casefold(): code for code, label in FORM_TABLE_OUTPUT_CHOICES}
for _ft_out_code, _ft_out_label in FORM_TABLE_OUTPUT_CHOICES:
    FORM_TABLE_OUTPUT_CODE[_ft_out_code.casefold()] = _ft_out_code

FORM_BLOCK_MODE_CHOICES = (
    (u"by_repeat_key", u"По ключевому вопросу"),
    (u"by_blank_row", u"По пустой строке"),
    (u"fixed_size", u"Фиксированное число вопросов"),
    (u"by_unique_cycle", u"По циклу уникальных вопросов"),
)
FORM_BLOCK_MODE_LABEL = {code: label for code, label in FORM_BLOCK_MODE_CHOICES}
FORM_BLOCK_MODE_CODE = {label.casefold(): code for code, label in FORM_BLOCK_MODE_CHOICES}
for _bm_code, _bm_label in FORM_BLOCK_MODE_CHOICES:
    FORM_BLOCK_MODE_CODE[_bm_code.casefold()] = _bm_code

FORM_BLOCK_SEP_CHOICES = (
    (u"blank_row", u"Пустая строка между блоками"),
    (u"none", u"Без разделителя"),
    (u"repeat_key", u"Старт по ключевому вопросу"),
)
FORM_BLOCK_SEP_LABEL = {code: label for code, label in FORM_BLOCK_SEP_CHOICES}
FORM_BLOCK_SEP_CODE = {label.casefold(): code for code, label in FORM_BLOCK_SEP_CHOICES}
for _bs_code, _bs_label in FORM_BLOCK_SEP_CHOICES:
    FORM_BLOCK_SEP_CODE[_bs_code.casefold()] = _bs_code

HEADER_SEP_CHOICES = (
    (u" ", u"пробел"),
    (u"_", u"подчёркивание _"),
    (u"-", u"дефис -"),
    (u"/", u"слэш /"),
    (u"", u"без разделителя"),
)
HEADER_SEP_LABEL = {code: label for code, label in HEADER_SEP_CHOICES}
HEADER_SEP_CODE = {label.casefold(): code for code, label in HEADER_SEP_CHOICES}
for _hdr_sep_code, _hdr_sep_label in HEADER_SEP_CHOICES:
    HEADER_SEP_CODE[_hdr_sep_code.casefold()] = _hdr_sep_code

_SHEET_BLOCK_FORM_SCHEMAS = {
    u"удаление_верхних_строк": {
        "title": u"Удаление верхних строк",
        "hint": (
            u"До Постобработка_xml удалить N верхних строк на листах результата.\n"
            u"Лист: «1_Заказы» или «Заказы» (префикс N_ учитывается).\n"
            u"Удаление всегда до переименовать_лист — порядок строк параметров не важен."
        ),
        "hint_lines": 4,
        "fields": (
            {"id": "n", "label": u"N верхних строк", "default": u"1"},
        ),
    },
    u"пропуск_строк_источника": {
        "title": u"Пропуск строк источника",
        "hint": (
            u"При сборе не копировать строки, где все указанные столбцы пусты.\n"
            u"Лист — имя листа *источника* (как в книге-источнике), не результата.\n"
            u"Пустой лист = правило для всех листов файла.\n"
            u"Столбцы: номера (1=A), буквы или имена заголовков через запятую.\n"
            u"Файл источника на диск не сохраняется."
        ),
        "hint_lines": 5,
        "fields": (
            {
                "id": "columns",
                "label": u"Столбцы (непустота)",
                "default": u"1",
            },
        ),
    },
    u"строка_заголовков": {
        "title": u"Многоэтажный заголовок",
        "hint": (
            u"Собрать имена столбцов из нескольких строк шапки с объединениями.\n"
            u"Диапазон строк: 2-5 или 2:5 (1-based). Итоговые имена пишутся "
            u"в нижней строке диапазона через разделитель (по умолчанию пробел).\n"
            u"Пустой итог → Колонка_A/B/…; повторы → Имя_1, Имя_2, …\n"
            u"Исходные файлы не изменяются (сборка в памяти)."
        ),
        "hint_lines": 5,
        "fields": (
            {
                "id": "rows",
                "label": u"Строки заголовка (2-5 или 2:5)",
                "default": u"1-2",
            },
            {
                "id": "separator",
                "label": u"Разделитель",
                "type": "combo",
                "choices": tuple(lab for _code, lab in HEADER_SEP_CHOICES),
                "default": u"пробел",
            },
        ),
    },
    u"Ячейка_В_Столбец": {
        "title": u"Ячейка → столбец",
        "hint": (
            u"Ячейка на листе источника → новый столбец в результате.\n"
            u"Строка — 1-based; столбец — буква (A) или номер (1).\n"
            u"sheet — лист результата (пусто = все); source_sheet — фильтр листа "
            u"источника (Отчет*, список через запятую)."
        ),
        "hint_lines": 4,
        "fields": (
            {"id": "col", "label": u"Столбец", "default": u"A"},
            {"id": "row", "label": u"Строка", "default": u"1"},
            {"id": "column", "label": u"Имя колонки", "default": u""},
            {"id": "sheet", "label": u"Лист результата", "default": u""},
            {"id": "source_sheet", "label": u"Лист источника", "default": u""},
        ),
    },
    u"Добавить_в_карту_переменных": {
        "title": u"Добавить в карту переменных",
        "hint": (
            u"Записать значение ячейки источника в runtime-карту переменных.\n"
            u"cell — A1-адрес на листе источника, name — имя переменной.\n"
            u"file/sheet — маски источника (*, все, список через запятую)."
        ),
        "hint_lines": 3,
        "fields": (
            {"id": "file", "label": u"Маска файла", "default": u"*"},
            {"id": "sheet", "label": u"Маска листа", "default": u"*"},
            {"id": "cell", "label": u"Ячейка (A1)", "default": u"A1"},
            {"id": "name", "label": u"Имя переменной", "default": u""},
        ),
    },
    u"Ручной_ввод_в_карту_переменных": {
        "title": u"Ручной ввод в карту переменных",
        "hint": (
            u"Несколько переменных — «+ Блок» слева; одна строка параметра, C = JSON-массив.\n"
            u"«Один диалог» — при сборе все переменные в одном окне (подсказки от первого блока).\n"
            u"Иначе — по одному диалогу на имя (ключ Имя~ручной_ввод~ручной_ввод).\n"
            u"«Запоминать» — последнее значение в ~/.config/libre-macros/collect_workbooks/<лист>.json;\n"
            u"«Шифровать» — lm1: тем же ключом, что пароли в Доп_Параметры_Источника."
        ),
        "hint_lines": 5,
        "blocks_without_sheet": True,
        "dialog_fit_content": True,
        "fields": (
            {"id": "name", "label": u"Имя переменной (Филиал)", "default": u""},
            {
                "id": "value_kind",
                "label": u"Вид значения (Одно значение / Список)",
                "type": "combo",
                "choices": tuple(lab for _c, lab in MANUAL_VAR_VALUE_KIND_CHOICES),
                "default": u"Одно значение",
            },
            {
                "id": "data_type",
                "label": u"Тип данных (Текст / Число / Дата)",
                "type": "combo",
                "choices": tuple(lab for _c, lab in MANUAL_VAR_DATA_TYPE_CHOICES),
                "default": u"Текст",
            },
            {"id": "format", "label": u"Формат ввода (ДД.ММ.ГГГГ, #0.00)", "default": u""},
            {
                "id": "choices",
                "label": u"Список выбора (Север, Юг, Центр)",
                "type": "columns_pick",
                "choices_from_headers": False,
                "plain_tokens": True,
                "default": u"",
            },
            {"id": "prompt", "label": u"Подсказка в диалоге (Выберите филиал…)", "default": u""},
            {"id": "title", "label": u"Заголовок диалога (Филиал)", "default": u""},
            {
                "id": "required",
                "label": u"Обязательно",
                "type": "bool",
                "default": True,
            },
            {"id": "default", "label": u"Значение по умолчанию (Привет)", "default": u""},
            {
                "id": "remember_last",
                "label": u"Запоминать последнее",
                "type": "bool",
                "default": True,
                "row_group": u"remember",
                "row_group_slot": 0,
            },
            {
                "id": "encrypt_last",
                "label": u"Шифровать",
                "type": "bool",
                "default": False,
                "row_group": u"remember",
                "row_group_slot": 1,
            },
        ),
    },
    u"градиент": {
        "title": u"Градиент",
        "hint": (
            u"Маркер заголовка столбца (выбор из списка листа, накопление через запятую), "
            u"цвета min и max (пресеты или свой текст). "
            u"«Вся строка» — заливка всех столбцов диапазона. "
            u"«Сортировка» — перед градиентом по найденному столбцу-маркеру."
        ),
        "hint_lines": 4,
        "fields": (
            {
                "id": "marker",
                "label": u"Маркер столбца",
                "type": "columns_pick",
                "default": u"сумм",
            },
            {
                "id": "color_min",
                "label": u"Цвет min",
                "type": "combo",
                "choices": _header_plus_height_color_choices(),
                "default": u"зеленый",
            },
            {
                "id": "color_max",
                "label": u"Цвет max",
                "type": "combo",
                "choices": _header_plus_height_color_choices(),
                "default": u"красный",
            },
            {
                "id": "whole_row",
                "label": u"Вся строка",
                "type": "bool",
                "default": False,
            },
            {
                "id": "sort",
                "label": u"Сортировка",
                "type": "bool",
                "default": False,
            },
            {
                "id": "sort_dir",
                "label": u"Направление",
                "type": "combo",
                # UI как в диалоге «сортировка»: + / -
                "choices": (u"+", u"-"),
                "default": u"+",
            },
        ),
    },
    u"подсветка_по_порогу": {
        "title": u"Подсветка по порогу",
        "hint": (
            u"Столбцы — по ним проверяется порог (накопление, × / ××).\n"
            u"Маски * и ? в именах столбцов (например Сумма_*).\n"
            u"Пороги min/max: подкраска при min ≤ значение ≤ max (пустая сторона — без границы).\n"
            u"Столбцы подсветки — куда применить стиль (пусто: только столбцы порога; "
            u"«вся строка» — вся используемая ширина листа).\n"
            u"Цвет заливки — пресет (пусто — не менять заливку). Цвет шрифта — палитра; "
            u"«Автоматический» — не менять цвет шрифта.\n"
            u"Перед подсветкой на целевых ячейках сбрасываются только заливка, "
            u"цвет шрифта, жирность и курсив (рамки и формат числа не трогаются)."
        ),
        "hint_lines": 5,
        "fields": (
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "default": u"",
            },
            {
                "id": "threshold_min",
                "label": u"Порог min",
                "default": u"",
                "row_group": u"threshold",
                "row_group_slot": 0,
            },
            {
                "id": "threshold_max",
                "label": u"Порог max",
                "default": u"",
                "row_group": u"threshold",
                "row_group_slot": 1,
            },
            {
                "id": "color",
                "label": u"Цвет заливки",
                "type": "combo",
                "choices": _header_plus_height_color_choices(),
                "default": u"",
            },
            {
                "id": "font_color",
                "label": u"Цвет шрифта",
                "type": "combo",
                "choices": _header_plus_height_color_choices(),
                "default": u"",
                "with_auto_check": u"font_color_auto",
            },
            {
                "id": "font_color_auto",
                "label": u"Автоматический цвет шрифта",
                "type": "bool",
                "default": True,
                "pair_after": u"font_color",
            },
            {
                "id": "style_columns",
                "label": u"Столбцы подсветки",
                "type": "columns_pick",
                "default": u"",
                "choices_from_headers": True,
            },
            {
                "id": "whole_row",
                "label": u"Применить для всей строки",
                "type": "bool",
                "default": False,
            },
            {
                "id": "bold",
                "label": u"Жирный шрифт",
                "type": "bool",
                "default": False,
                "row_group": u"font_style",
                "row_group_slot": 0,
            },
            {
                "id": "italic",
                "label": u"Курсив",
                "type": "bool",
                "default": False,
                "row_group": u"font_style",
                "row_group_slot": 1,
            },
        ),
    },
    u"условное_форматирование": {
        "alias_of": u"подсветка_по_порогу",
    },
    u"подкрасить_пороги": {
        "alias_of": u"подсветка_по_порогу",
        "title": u"Подкрасить пороги",
    },
    u"формат_даты": {
        "title": u"Формат даты",
        "hint": (
            u"Формат: пусто — из глобальных настроек.\n"
            u"Столбцы — накопительный выбор из заголовков (из списка → 'кавычки', × / ××);\n"
            u"  'Имя' — полное совпадение; 'Имя*' / '*Имя' — шаблон; без кавычек — подстрока (*часть*).\n"
            u"Индексы 1-based и буквы A допустимы. Пусто — авто по типу ячеек "
            u"(при конвертации — ещё по 3 верхним строкам).\n"
            u"Маркеры — то же правило отбора, если столбцы не заданы.\n"
            u"parse_formats — разбор текста при «Конвертировать» (как MERGE_XML_DATETIME_FORMATS):\n"
            u"  FIRST_DAY.MMMM.YEAR, LAST_DAY.MMMM.NEXT_YEAR, 01.ММММ.2027;\n"
            u"  пусто — взять MERGE_XML_DATETIME_FORMATS, если задан при сборе."
        ),
        "hint_lines": 8,
        "fields": (
            {
                "id": "format",
                "label": u"Формат даты",
                "type": "combo",
                "choices": (),
                "default": u"",
            },
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "markers",
                "label": u"Маркеры заголовка",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "parse_formats",
                "label": u"Разбор текста (parse_formats)",
                "type": "edit",
                "default": u"",
                "hint": (
                    u"Шаблоны через запятую. Month-only: FIRST_DAY.MMMM.YEAR, "
                    u"LAST_DAY.MMMM.NEXT_YEAR, 01.MMMM.2027"
                ),
            },
        ),
        "convert_existing": True,
        "convert_label": u"Конвертировать текст в дату",
    },
    u"формат_деньги": {
        "title": u"Формат денег / чисел",
        "hint": (
            u"Формат: пусто — числовой формат из глобальных настроек.\n"
            u"Столбцы — накопительный выбор из заголовков (из списка → 'кавычки', × / ××);\n"
            u"  'Имя' — полное совпадение; 'Имя*' / '*Имя' — шаблон; без кавычек — подстрока (*часть*).\n"
            u"Индексы и буквы A допустимы. Пусто — автоопределение числовых столбцов "
            u"(при конвертации — ещё текстовые числа в 3 верхних строках).\n"
            u"Маркеры — то же правило отбора, если столбцы не заданы."
        ),
        "hint_lines": 6,
        "fields": (
            {
                "id": "format",
                "label": u"Числовой формат",
                "type": "combo",
                "choices": (),
                "default": u"",
            },
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "markers",
                "label": u"Маркеры заголовка",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
        ),
        "convert_existing": True,
        "convert_label": u"Конвертировать текст в число",
    },
    u"удалить_столбцы": {
        "title": u"Удалить столбцы",
        "hint": (
            u"Столбцы для удаления (через запятую): имена из списка, подстроки заголовка "
            u"(или шаблон *), индексы 1-based и буквы (A, C). Выбор из списка накапливается."
        ),
        "hint_lines": 3,
        "fields": (
            {
                "id": "markers",
                "label": u"Столбцы / маркеры",
                "type": "columns_pick",
                "default": u"",
            },
        ),
    },
    u"сетка": {
        "title": u"Сетка",
        "hint": (
            u"Толщина (тонкая/толстая/число) и цвет рамки. Пример: толстая/#666666. "
            u"«Включая заголовок» — сетка и на строке заголовков (по умолчанию выкл.)."
        ),
        "hint_lines": 3,
        "fields": (
            {"id": "width", "label": u"Толщина", "default": u"тонкая"},
            {"id": "color", "label": u"Цвет", "default": u""},
            {
                "id": "include_header",
                "label": u"Включая заголовок",
                "type": "bool",
                "default": False,
            },
        ),
    },
    u"высота_строки": {
        "title": u"Высота строки",
        "hint": u"Строки: all или 2,3,4; высота в мм (0 = скрыть). Пустой лист — текущий лист шага.",
        "fields": (
            {"id": "rows", "label": u"Строки (all / список)", "default": u"all"},
            {"id": "height_mm", "label": u"Высота, мм", "default": u"5"},
        ),
    },
    u"отступ": {
        "title": u"Отступ",
        "hint": u"Выравнивание left/center/right, шаги, столбцы через запятую (пусто — все).",
        "fields": (
            {"id": "h_align", "label": u"Выравнивание", "default": u"left"},
            {"id": "steps", "label": u"Шаги", "default": u"1"},
            {"id": "columns", "label": u"Столбцы", "default": u""},
        ),
    },
    u"шрифт": {
        "title": u"Шрифт",
        "hint": u"Имя; размер данных; размер заголовка. «—» или пусто — не менять.",
        "fields": (
            {"id": "name", "label": u"Имя", "default": u""},
            {"id": "size_data", "label": u"Размер данных", "default": u""},
            {"id": "size_header", "label": u"Размер заголовка", "default": u""},
        ),
    },
    u"ширина_столбцов": {
        "title": u"Ширина столбцов",
        "hint": (
            u"Столбцы — выбор из заголовков ('кавычки', накопление, × / ××) "
            u"или «Все»; индексы 1-based и буквы A тоже допустимы.\n"
            u"Ширина в мм; 0 = скрыть столбец(ы)."
        ),
        "hint_lines": 3,
        "fields": (
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"all",
            },
            {"id": "width_mm", "label": u"Ширина, мм", "default": u"30"},
        ),
    },
    u"перенос": {
        "title": u"Перенос",
        "hint": u"Перенос текста в данных; опционально включая строку заголовков.",
        "fields": (
            {"id": "wrap", "label": u"Перенос", "type": "bool", "default": True},
            {
                "id": "include_header",
                "label": u"Включая заголовок",
                "type": "bool",
                "default": False,
            },
        ),
    },
    u"перенос_и_авто_высота": {
        "alias_of": u"перенос",
        "title": u"Перенос и авто-высота",
        "hint": u"Перенос + автоподбор высоты строк; «Включая заголовок» — и для заголовка.",
    },
    u"заголовок_плюс_высота": {
        "title": u"Заголовок + высота",
        "hint": (
            u"Высота заголовка в мм (по умолчанию ставится как указано; мин. 7 мм). "
            u"Галочка «Плюсовать» — прибавить к текущей высоте. "
            u"Гориз.: left/center/right; верт.: top/center/bottom. "
            u"Цвета — имя или #hex. «Автоматически» для цвета шрифта — контраст к заливке."
        ),
        "hint_lines": 5,
        "fields": (
            {
                "id": "height_mm",
                "label": u"Высота, мм",
                "default": u"12.7",
                "with_bool_check": u"add_height",
                "edit_frac": 0.30,
            },
            {
                "id": "add_height",
                "label": u"Плюсовать",
                "type": "bool",
                "default": False,
                "pair_after": u"height_mm",
            },
            {
                "id": "h_align",
                "label": u"Гориз.",
                "type": "combo",
                "choices": _HEADER_PLUS_HEIGHT_HORI_CHOICES,
                "default": u"center",
            },
            {
                "id": "v_align",
                "label": u"Верт.",
                "type": "combo",
                "choices": _HEADER_PLUS_HEIGHT_VERT_CHOICES,
                "default": u"center",
            },
            {"id": "font_name", "label": u"Шрифт", "default": u""},
            {"id": "font_size", "label": u"Размер, pt", "default": u""},
            {
                "id": "fill_color",
                "label": u"Заливка",
                "type": "combo",
                "choices": _header_plus_height_color_choices(),
                "default": u"",
            },
            {
                "id": "font_color",
                "label": u"Цвет шрифта",
                "type": "combo",
                "choices": _header_plus_height_color_choices(),
                "default": u"",
            },
            {
                "id": "font_color_auto",
                "label": u"Автоматически (цвет шрифта)",
                "type": "bool",
                "default": True,
            },
            {
                "id": "bold",
                "label": u"Жирный шрифт",
                "type": "bool",
                "default": False,
            },
        ),
    },
    u"зебра_диапазон": {
        "title": u"Зебра диапазон",
        "hint": u"Роли header/even/odd — заливка/шрифт (как в C: header:excel_header/белый).",
        "fields": (
            {"id": "header", "label": u"Заголовок (заливка/шрифт)", "default": u""},
            {"id": "even", "label": u"Чётные", "default": u""},
            {"id": "odd", "label": u"Нечётные", "default": u""},
        ),
    },
    u"конкатенация_столбцов": {
        "title": u"Конкатенация столбцов",
        "hint": (
            u"Имя нового столбца, разделитель (пусто = склейка без разделителя).\n"
            u"Столбцы — выбор из заголовков листа (накопление через запятую; имена с «,» "
            u"в 'кавычках'; × / ××). Допустимы также индексы 1-based и буквы (A, C)."
        ),
        "hint_lines": 3,
        "fields": (
            {"id": "new_column", "label": u"Имя столбца", "default": u"Полный_Код"},
            {"id": "separator", "label": u"Разделитель", "default": u"|"},
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"1,3,5",
            },
        ),
    },
    u"разделить_по_столбцам": {
        "title": u"Разделить столбец",
        "hint": (
            u"Аналог Power Query «Разделить столбец по разделителю».\n"
            u"Выберите исходный столбец, разделитель и место вставки новых столбцов.\n"
            u"В режиме «Заменить» первая часть остаётся в исходном столбце."
        ),
        "hint_lines": 3,
        "dialog_fit_content": True,
        "fields": (
            {
                "id": "column",
                "label": u"Столбец",
                "type": "combo",
                "choices_from_headers": True,
                "default": u"",
            },
            {"id": "delimiter", "label": u"Разделитель", "default": u","},
            {
                "id": "position",
                "label": u"Куда вставить части",
                "type": "combo",
                "choices": (u"_начало_", u"_конец_", u"_перед_", u"_после_"),
                "default": u"_после_",
            },
            {
                "id": "relative_column",
                "label": u"Опорный столбец",
                "type": "combo",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "source_policy",
                "label": u"Исходный столбец",
                "type": "combo",
                "choices": (u"replace", u"keep"),
                "default": u"replace",
            },
            {
                "id": "split_mode",
                "label": u"Режим split",
                "type": "combo",
                "choices": (u"each", u"leftmost"),
                "default": u"each",
            },
            {"id": "max_columns", "label": u"Макс. частей (0=авто)", "default": u"0"},
            {"id": "new_names", "label": u"Имена новых столбцов", "default": u""},
            {"id": "trim_parts", "label": u"Обрезать пробелы", "type": "bool", "default": True},
            {
                "id": "non_text",
                "label": u"Нетекст",
                "type": "combo",
                "choices": (u"skip", u"empty", u"text", u"fail"),
                "default": u"skip",
            },
        ),
    },
    u"условный_столбец": {
        "title": u"Условный столбец",
        "hint": (
            u"Аналог Power Query «Conditional Column».\n"
            u"Правила и ELSE задаются конструктором (кнопка «Правила...»)."
        ),
        "hint_lines": 2,
        "dialog_fit_content": True,
        "fields": (
            {"id": "new_column", "label": u"Имя нового столбца", "default": u"Категория"},
            {
                "id": "position",
                "label": u"Позиция",
                "type": "combo",
                "choices": (u"_начало_", u"_конец_", u"_перед_", u"_после_"),
                "default": u"_конец_",
            },
            {
                "id": "relative_column",
                "label": u"Опорный столбец",
                "type": "combo",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "compare_as",
                "label": u"Сравнивать как",
                "type": "combo",
                "choices": (u"auto", u"text", u"number", u"date"),
                "default": u"auto",
            },
            {"id": "trim", "label": u"Обрезать пробелы", "type": "bool", "default": True},
            {
                "id": "case_sensitive",
                "label": u"Учитывать регистр",
                "type": "bool",
                "default": False,
            },
        ),
    },
    u"столбец_по_лямбде": {
        "title": u"Столбец по лямбде (Python)",
        "hint": (
            u"Custom Column: каждое значение — результат Python-лямбды над строкой.\n"
            u"Sugar: rows(\"Колонка\"), rows(...)[-1], upd_rows(\"Колонка\").\n"
            u"<<Переменные.имя>> — подстановка из карты до компиляции лямбды.\n"
            u"Имя столбца: new_column (литерал) или new_column_expr (лямбда sheet, headers).\n"
            u"Запись — значения (не формулы). Формат данных — по result_type/first value."
        ),
        "hint_lines": 6,
        "dialog_fit_content": True,
        "fields": (
            {"id": "new_column", "label": u"Имя нового столбца", "default": u"Новый_столбец"},
            {
                "id": "new_column_expr",
                "label": u"Лямбда имени столбца (опционально)",
                "type": "lambda_pick",
                "default": u"",
                "hint": u'Пример: lambda sheet, headers: "Копия_" + sheet',
            },
            {
                "id": "position",
                "label": u"Позиция",
                "type": "combo",
                "choices": (u"_начало_", u"_конец_", u"_перед_", u"_после_"),
                "default": u"_конец_",
            },
            {
                "id": "relative_column",
                "label": u"Опорный столбец",
                "type": "combo",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "expr",
                "label": u"Лямбда значений expr (обязательна)",
                "type": "lambda_pick",
                "default": u'lambda rows: rows("Колонка")',
                "hint": (
                    u'Пример: lambda rows: rows("Сумма") != upd_rows("Сумма") if upd_rows_len() else True'
                ),
            },
            {
                "id": "on_error",
                "label": u"При ошибке строки",
                "type": "combo",
                "choices": (u"keep", u"empty", u"fail"),
                "default": u"empty",
            },
            {
                "id": "result_type",
                "label": u"Тип результата",
                "type": "combo",
                "choices": (u"auto", u"text", u"number"),
                "default": u"auto",
            },
            {
                "id": "output_policy",
                "label": u"Если столбец найден",
                "type": "combo",
                "choices": (u"new", u"replace"),
                "default": u"new",
            },
        ),
    },
    u"добавить_столбец": {
        "title": u"Добавить столбец",
        "hint": (
            u"Пустой столбец в конец таблицы на выбранном листе (или на всех — «Одинаково…»).\n"
            u"Имя — одно. Тип: Текст / Число / Дата.\n"
            u"Текст — список возможных значений (выпадающий список на столбце).\n"
            u"Число/Дата — формат; опционально контроль данных и мин/макс.\n"
            u"Если столбец с таким именем уже есть — формат и контроль применяются к нему."
        ),
        "hint_lines": 5,
        "dialog_fit_content": True,
        "fields": (
            {"id": "new_column", "label": u"Имя нового столбца", "default": u"Новый"},
            {
                "id": "data_type",
                "label": u"Тип данных",
                "type": "combo",
                "choices": tuple(lab for _c, lab in ADD_COLUMN_DATA_TYPE_CHOICES),
                "default": u"Текст",
            },
            {
                "id": "format",
                "label": u"Формат",
                "type": "combo",
                "choices": (),
                "default": u"",
            },
            {
                "id": "choices",
                "label": u"Список значений (текст)",
                "type": "columns_pick",
                "choices_from_headers": False,
                "plain_tokens": True,
                "default": u"",
            },
            {
                "id": "data_validation",
                "label": u"Контроль данных",
                "type": "bool",
                "default": False,
            },
            {"id": "min", "label": u"Минимум", "default": u""},
            {"id": "max", "label": u"Максимум", "default": u""},
        ),
    },
    u"применить_формулу": {
        "title": u"Применить формулу",
        "hint": (
            u"Столбец результата и формула Calc. "
            u"{row_id} — строка данных (1-based), {col_id} — столбец вставки (A, B, …). "
            u"[Имя] / {Имя} — столбец по заголовку. "
            u"Формат — NumberFormat всего столбца (пусто — с первой строки / соседнего). "
            u"Лист в блоке — только для него; пусто — все листы шага.\n"
            u"«Как значение» в Постобработка_xml недоступно (прямой доступ).\n"
            u"Карта переменных: <<Переменные.имя>> или <<Переменные.имя_переменной([лист],[файл])>>.\n"
            u"[лист]/[файл] могут быть константами или ссылками на колонки текущей строки: [Источник_Лист], [Источник_Файл].\n"
            u"<<Переменные.имя_переменной(#Путь)>>: берём из служебной колонки #Путь и парсим как file#sheet.\n"
            u"Поиск: сначала полное совпадение, затем частичное по сегментам листа/файла."
        ),
        "hint_lines": 7,
        "fields": (
            {"id": "column", "label": u"Столбец", "default": u"Итого"},
            {
                "id": "formula",
                "label": u"Формула",
                "default": u"=IFERROR({col_id}{row_id-1}+1;1)",
            },
            {
                "id": "format",
                "label": u"Формат столбца",
                "type": "combo",
                "choices": (),
                "default": u"",
            },
            {
                "id": "as_values",
                "label": u"Как значение (после протяжки)",
                "type": "bool",
                "default": False,
            },
        ),
    },
    u"удаление_строк": {
        "title": u"Удаление строк",
        "hint": (
            u"Режим «columns»: удалить строки, где указанные столбцы пусты "
            u"(выбор из заголовков, 'кавычки', × / ××; ручной ввод через , или ;).\n"
            u"Режим «formula»: формула пустой строки — ИСТИНА → удалить "
            u"({row_id}, {col}{row-1}, R[-1]C — как в «применить_формулу»)."
        ),
        "hint_lines": 4,
        "fields": (
            {
                "id": "mode",
                "label": u"Режим",
                "type": "combo",
                "choices": (u"columns", u"formula"),
                "default": u"columns",
            },
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"1,A,ФИО",
            },
            {"id": "formula", "label": u"Формула", "default": u""},
        ),
    },
    u"пропуск_пустых_строк": {
        "alias_of": u"удаление_строк",
        "title": u"Пропуск пустых строк",
    },
    u"группировка_по_столбцу": {
        "title": u"Группировка / промежуточные итоги",
        "hint": (
            u"Столбец группировки — из заголовков ('кавычки').\n"
            u"Промежуточные итоги: функция + столбцы (строка «Отдел Количество» / «Отдел Сумма»).\n"
            u"Несколько функций по одному столбцу — подряд две строки группировка_по_столбцу "
            u"(склеиваются в один проход: count+sum и общие итоги внизу).\n"
            u"«Сортировка» / «Outline» — как у Data→Subtotals."
        ),
        "hint_lines": 6,
        "fields": (
            {
                "id": "marker",
                "label": u"Столбец группировки",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"Отдел",
            },
            {
                "id": "agg_fn",
                "label": u"Функция итога",
                "type": "combo",
                "choices": tuple(label for _c, label in GROUP_AGG_FN_CHOICES),
                "default": u"Сумма",
            },
            {
                "id": "agg_columns",
                "label": u"Столбцы итога",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"Сумма",
            },
            {
                "id": "sort",
                "label": u"Сортировка",
                "type": "bool",
                "default": True,
            },
            {
                "id": "outline",
                "label": u"Outline (свернуть)",
                "type": "bool",
                "default": True,
            },
        ),
    },
    u"стиль_печати": {
        "title": u"Стиль печати",
        "hint": u"landscape/portrait и число страниц по ширине (0 = не вписывать).",
        "fields": (
            {"id": "orientation", "label": u"Ориентация", "default": u"landscape"},
            {"id": "fit_pages", "label": u"Вписать в страниц", "default": u"1"},
        ),
    },
    u"заполнение_вниз": {
        "title": u"Заполнение вниз",
        "hint": (
            u"Столбцы для протяжки вниз — выбор из заголовков листа "
            u"(накопление, 'кавычки', × / ××; ручной ввод через , или ;).\n"
            u"Допустимы также индексы 1-based и буквы (A, C)."
        ),
        "hint_lines": 3,
        "fields": (
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"ФИО,Отдел",
            },
        ),
    },
    u"заполнить_вверх": {
        "title": u"Заполнить вверх",
        "hint": (
            u"Пустые ячейки заполняются значением ближайшей непустой снизу "
            u"(Fill Up). Столбцы — выбор из заголовков ('кавычки', × / ××).\n"
            u"Верхние пустоты без источника снизу остаются пустыми."
        ),
        "hint_lines": 3,
        "fields": (
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
                "hint": u"Накопление из списка; × / ××",
            },
            {
                "id": "skip_if_above_not_empty",
                "label": u"Только пустые ячейки",
                "type": "bool",
                "default": True,
            },
            {
                "id": "drop_trailing_empty",
                "label": u"Удалять верх без источника (v2)",
                "type": "bool",
                "default": False,
            },
        ),
    },
    u"fill_up": {
        "alias_of": u"заполнить_вверх",
    },
    u"заполнение_вверх": {
        "alias_of": u"заполнить_вверх",
    },
    u"развернуть_столбцы": {
        "title": u"Развернуть столбцы (Unpivot)",
        "hint": (
            u"Широкая таблица → длинная: каждый выбранный столбец превращается в блок строк;\n"
            u"имя столбца → колонка «Атрибут», значение ячейки → «Значение».\n"
            u"Либо укажите unpivot_columns, либо оставьте пустым и задайте exclude_columns "
            u"(развернуть всё, кроме…). × / ×× — убрать последний / очистить."
        ),
        "hint_lines": 4,
        "fields": (
            {
                "id": "unpivot_columns",
                "label": u"Столбцы для разворота",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
                "hint": u"Пусто + exclude = все кроме…",
            },
            {
                "id": "exclude_columns",
                "label": u"Исключить (оставить как ключи)",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "attribute_column",
                "label": u"Колонка атрибута (имя)",
                "default": u"Месяц",
            },
            {
                "id": "value_column",
                "label": u"Колонка значения (имя)",
                "default": u"Сумма",
            },
            {
                "id": "output",
                "label": u"Куда писать",
                "type": "combo",
                "choices": tuple(label for _code, label in UNPIVOT_OUTPUT_CHOICES),
                "default": u"На месте",
            },
            {
                "id": "dest_sheet",
                "label": u"Лист назначения (для нового листа)",
                "default": u"Unpivot",
            },
            {
                "id": "drop_empty_rows",
                "label": u"Не создавать строки с пустым значением",
                "type": "bool",
                "default": False,
            },
            {
                "id": "as_values",
                "label": u"Как значения (формулы)",
                "type": "bool",
                "default": True,
            },
        ),
    },
    u"unpivot": {
        "alias_of": u"развернуть_столбцы",
    },
    u"unpivot_columns": {
        "alias_of": u"развернуть_столбцы",
    },
    u"транспонировать_таблицу": {
        "title": u"Транспонировать таблицу",
        "hint": (
            u"Строки ↔ столбцы. Куда писать: на месте / новый лист / рядом.\n"
            u"Заголовки: из колонки меток ИЛИ вручную («Имена столбцов результата»).\n"
            u"Подробности, примеры и ограничения — кнопка «Справка»."
        ),
        "hint_lines": 3,
        "compact_fields": True,
        "fields": (
            {
                "id": "output",
                "label": u"Куда писать",
                "type": "combo",
                "choices": tuple(lab for _c, lab in TRANSPOSE_OUTPUT_CHOICES),
                "default": u"На месте",
            },
            {
                "id": "dest_sheet",
                "label": u"Лист назначения (new_sheet)",
                "default": u"Матрица_T",
            },
            {
                "id": "dest_cell",
                "label": u"Якорь записи (offset), A1",
                "default": u"",
            },
            {
                "id": "range",
                "label": u"Диапазон (пусто = авто)",
                "default": u"",
            },
            {
                "id": "header_row",
                "label": u"Строка заголовка (1…)",
                "default": u"1",
            },
            {
                "id": "columns",
                "label": u"Столбцы (пусто = все)",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "headers_from_column",
                "label": u"Заголовки из колонки",
                "type": "bool",
                "default": False,
            },
            {
                "id": "header_column",
                "label": u"Колонка меток (пусто = первая)",
                "default": u"",
            },
            {
                "id": "result_headers",
                "label": u"Имена столбцов результата",
                "default": u"",
                "hint": (
                    u"Через запятую или «;». Добавляются сверху после транспонирования. "
                    u"Пусто = без доп. шапки. Не используется при «Заголовки из колонки»."
                ),
            },
            {
                "id": "skip_header_row",
                "label": u"Без первой строки (не inplace)",
                "type": "bool",
                "default": False,
            },
            {
                "id": "skip_first_column",
                "label": u"Без первого столбца (не inplace)",
                "type": "bool",
                "default": False,
            },
            {
                "id": "as_values",
                "label": u"Как значения (формулы)",
                "type": "bool",
                "default": True,
            },
        ),
    },
    u"transpose": {"alias_of": u"транспонировать_таблицу"},
    u"transpose_table": {"alias_of": u"транспонировать_таблицу"},
    u"анкета_в_таблицу": {
        "title": u"Анкета → таблица",
        "hint": (
            u"Вертикальные пары «вопрос → ответ» → широкая таблица (строка = анкета).\n"
            u"Нарезка: ключевой вопрос / пустая строка / fixed / цикл. "
            u"Опционально шапка листа (from–to + forms_start_row) и шапка блока."
        ),
        "hint_lines": 2,
        "compact_fields": True,
        "fields": (
            {"id": "question_column", "label": u"Колонка вопросов", "default": u"A"},
            {"id": "answer_column", "label": u"Колонка ответов", "default": u"B"},
            {
                "id": "block_mode",
                "label": u"Нарезка блоков",
                "type": "combo",
                "choices": tuple(lab for _c, lab in FORM_BLOCK_MODE_CHOICES),
                "default": u"По ключевому вопросу",
            },
            {
                "id": "block_start_question",
                "label": u"Ключевой вопрос (старт блока)",
                "default": u"ФИО",
            },
            {
                "id": "blank_rows_to_split",
                "label": u"Пустых строк для разделения",
                "default": u"1",
            },
            {
                "id": "questions_per_block",
                "label": u"Вопросов в блоке (fixed)",
                "default": u"",
            },
            {
                "id": "sheet_preamble_row_from",
                "label": u"Шапка листа: с",
                "default": u"",
                "row_group": u"sheet_preamble_rows",
                "row_group_slot": 0,
                "hint": u"sheet_preamble_row_from; пусто = авто/нет",
            },
            {
                "id": "sheet_preamble_row_to",
                "label": u"по",
                "default": u"",
                "row_group": u"sheet_preamble_rows",
                "row_group_slot": 1,
                "hint": u"sheet_preamble_row_to; задайте оба или оба пусто",
            },
            {
                "id": "forms_start_row",
                "label": u"Строка начала анкет",
                "default": u"",
                "row_group": u"forms_block_preamble",
                "row_group_slot": 0,
                "hint": u"forms_start_row; пусто = начало диапазона данных",
            },
            {
                "id": "block_preamble_rows",
                "label": u"Строк шапки блока",
                "default": u"0",
                "row_group": u"forms_block_preamble",
                "row_group_slot": 1,
                "hint": u"block_preamble_rows",
            },
            {
                "id": "output",
                "label": u"Куда писать",
                "type": "combo",
                "choices": tuple(lab for _c, lab in FORM_TABLE_OUTPUT_CHOICES),
                "default": u"Новый лист",
                "row_group": u"output_dest",
                "row_group_slot": 0,
                "row_group_frac": 0.48,
            },
            {
                "id": "dest_sheet",
                "label": u"Лист назначения",
                "default": u"Анкеты_wide",
                "row_group": u"output_dest",
                "row_group_slot": 1,
            },
            {
                "id": "add_block_index",
                "label": u"Колонка #Блок",
                "type": "bool",
                "default": True,
                "row_group": u"form_flags",
                "row_group_slot": 0,
            },
            {
                "id": "add_source_row",
                "label": u"source_start_row",
                "type": "bool",
                "default": False,
                "row_group": u"form_flags",
                "row_group_slot": 1,
            },
            {
                "id": "as_values",
                "label": u"Как значения (формулы)",
                "type": "bool",
                "default": True,
            },
        ),
    },
    u"form_to_table": {"alias_of": u"анкета_в_таблицу"},
    u"таблица_в_анкету": {
        "title": u"Таблица → анкета",
        "hint": (
            u"Широкая таблица → вертикальные пары вопрос/ответ (блоки вниз).\n"
            u"Столбцы тела — columns_pick; шапка листа / шапка блока — отдельно.\n"
            u"Первая строка выхода: заголовки колонок (по умолчанию Вопрос / Ответ)."
        ),
        "hint_lines": 3,
        "compact_fields": True,
        "fields": (
            {
                "id": "body_columns",
                "label": u"Столбцы тела (вопросы)",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
                "hint": u"Пусто = все, кроме skip/preamble",
            },
            {
                "id": "skip_columns",
                "label": u"Исключить столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "sheet_preamble_columns",
                "label": u"Шапка листа (один раз)",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "block_preamble_columns",
                "label": u"Шапка блока (каждый)",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "question_column",
                "label": u"Колонка вопросов (выход)",
                "default": u"A",
                "row_group": u"qa_out_cols",
                "row_group_slot": 0,
            },
            {
                "id": "answer_column",
                "label": u"Колонка ответов (выход)",
                "default": u"B",
                "row_group": u"qa_out_cols",
                "row_group_slot": 1,
            },
            {
                "id": "question_header",
                "label": u"Заголовок колонки вопросов",
                "default": u"Вопрос",
                "row_group": u"qa_headers",
                "row_group_slot": 0,
                "hint": u"question_header",
            },
            {
                "id": "answer_header",
                "label": u"Заголовок колонки ответов",
                "default": u"Ответ",
                "row_group": u"qa_headers",
                "row_group_slot": 1,
                "hint": u"answer_header",
            },
            {
                "id": "has_header_in_output",
                "label": u"Писать строку заголовков",
                "type": "bool",
                "default": True,
                "hint": u"has_header_in_output",
            },
            {
                "id": "block_separator",
                "label": u"Между блоками",
                "type": "combo",
                "choices": tuple(lab for _c, lab in FORM_BLOCK_SEP_CHOICES),
                "default": u"Пустая строка между блоками",
            },
            {
                "id": "block_start_question",
                "label": u"Ключ старта (для repeat_key)",
                "default": u"",
            },
            {
                "id": "output",
                "label": u"Куда писать",
                "type": "combo",
                "choices": tuple(lab for _c, lab in FORM_TABLE_OUTPUT_CHOICES),
                "default": u"Новый лист",
                "row_group": u"t2f_output_dest",
                "row_group_slot": 0,
                "row_group_frac": 0.48,
            },
            {
                "id": "dest_sheet",
                "label": u"Лист назначения",
                "default": u"Анкеты",
                "row_group": u"t2f_output_dest",
                "row_group_slot": 1,
            },
            {
                "id": "as_values",
                "label": u"Как значения (формулы)",
                "type": "bool",
                "default": True,
            },
        ),
    },
    u"table_to_form": {"alias_of": u"таблица_в_анкету"},
    u"копировать_значения": {
        "title": u"Копировать значения",
        "hint": (
            u"Столбцы для копирования значений — выбор из заголовков "
            u"('кавычки', накопление, × / ××).\n"
            u"«Весь лист» — без списка столбцов."
        ),
        "hint_lines": 3,
        "fields": (
            {"id": "whole_sheet", "label": u"Весь лист", "type": "bool", "default": False},
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"Сумма",
            },
        ),
    },
    u"замена_значений": {
        "title": u"Замена значений",
        "hint": (
            u"Подстроки в текстовых столбцах. find через «;»/«,»; replace — вручную или лямбда (кортеж).\n"
            u"Replace короче find → частичная замена + warning. '…' не дробится; * и _ПУСТО_.\n"
            u"<<Переменные.имя>> в find/replace/лямбдах (в '…' — литерал). Фильтр строк / match_pick — лямбды."
        ),
        "hint_lines": 3,
        "fields": (
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "default": u"ФИО",
            },
            {"id": "find", "label": u"Найти (; или '…')", "default": u"ООО",
             "hint": (
                 u"Несколько через ; . Запятую и пробелы по краям — в кавычках: "
                 u"','; ' Петр '. Без кавычек голая «,» — разделитель."
             )},
            {
                "id": "replace_mode",
                "label": u"Как",
                "type": "combo",
                "choices": tuple(label for _c, label in REPLACE_MODE_CHOICES),
                "default": u"Вручную",
                "row_group": u"replace_line",
                "row_group_slot": 0,
                "row_group_frac": 0.15,
            },
            {
                "id": "replace",
                "label": u"Заменить (1 или N; '…')",
                "default": u"ОБЩЕСТВО",
                "hint": u"Одно значение на все find, или N через ; . Короче find — частичная замена + warning.",
                "row_group": u"replace_line",
                "row_group_slot": 1,
            },
            {
                "id": "replace_expr",
                "label": u"Лямбда замены",
                "type": "lambda_pick",
                "default": u"",
                "hint": (
                    u"Только при «Лямбда». lambda rows: \"X\" или "
                    u'lambda rows: ("a", "b", _ПУСТО_) — кортеж под find. '
                    u"sugar: rows/upd_rows/tr; <<Переменные…>>."
                ),
            },
            {
                "id": "case_insensitive",
                "label": u"Без учета регистра",
                "type": "bool",
                "default": True,
                "row_group": u"replace_flags",
                "row_group_slot": 0,
            },
            {
                "id": "squeeze_spaces",
                "label": u"Сжать пробелы",
                "type": "bool",
                "default": True,
                "row_group": u"replace_flags",
                "row_group_slot": 1,
            },
            {
                "id": "row_filter",
                "label": u"Фильтр строк (лямбда)",
                "type": "lambda_pick",
                "default": u"",
                "hint": (
                    u"Пусто — искать во всех строках. Иначе lambda rows: … "
                    u"(rows(\"Статус\"), rows(\"Кол\")[-1], rows(1) / rows(\"A\")). "
                    u"True — строка участвует в поиске."
                ),
            },
            {
                "id": "match_pick",
                "label": u"Если несколько совпадений",
                "type": "combo",
                "choices": tuple(label for _code, label in REPLACE_MATCH_PICK_CHOICES),
                "default": u"Все совпадения",
            },
            {
                "id": "match_expr",
                "label": u"Лямбда отбора совпадений",
                "type": "lambda_pick",
                "default": u"",
                "hint": (
                    u"Только при отборе «Лямбда». lambda rows: … плюс idx (0-based), "
                    u"n, i (1-based), text ячейки. Пример: lambda rows: idx == 0"
                ),
            },
        ),
    },
    u"текстовые_операции": {
        "title": u"Текстовые операции",
        "hint": (
            u"Стек ops по столбцам (кнопка «Справка» — полный каталог).\n"
            u"ops: trim, collapse_ws, substring, clear, case, lang_homoglyphs, reverse,\n"
            u"base64, sha256, guid, encrypt/decrypt; row_filter per-op; non_text/on_error.\n"
            u"Ключ шифра: key | key_cell | key_file (или <<Переменные.secret_key>>)."
        ),
        "hint_lines": 4,
        "fields": (
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"ФИО",
            },
            {
                "id": "non_text",
                "label": u"Нетекст",
                "type": "combo",
                "choices": (u"пропуск", u"как текст", u"ошибка"),
                "default": u"пропуск",
            },
            {
                "id": "on_error",
                "label": u"При ошибке op",
                "type": "combo",
                "choices": (u"оставить до ошибки", u"очистить ячейку", u"остановить"),
                "default": u"оставить до ошибки",
            },
        ),
    },
    u"переименовать_лист": {
        "title": u"Переименовать лист",
        "hint": (
            u"Новое имя; лист слева — исходное имя (пусто = текущий лист шага).\n"
            u"Плейсхолдеры: <<Переменные.имя>>, [Старое_имя], [Заголовок(N)], "
            u"[Строка(номер|первая|последняя, C)]; "
            u"[Текущая_Дата_Время_YYYYMMDD], [Текущая_Дата_YYYY-MM-DD], "
            u"[Текущее_Время_HH-mm-SS] (или без формата); "
            u"[Имя_Книги], [Номер_Листа], [GUID8]. "
            u"Токены: YYYY/YY/MM/DD/HH/mm/SS (минуты — mm).\n"
            u"После подстановки: убрать недопустимые символы; длина ≤31, иначе 28+_01."
        ),
        "hint_lines": 5,
        "fields": (
            {
                "id": "new_name",
                "label": u"Новое имя",
                "default": u"<<Переменные.Филиал>>_[Старое_имя]",
            },
        ),
    },
    u"переименовать_столбцы": {
        "title": u"Переименовать столбцы",
        "hint": (
            u"Старое имя — выбор из заголовков (в 'кавычках' = полное совпадение; "
            u"запятые внутри кавычек не режут список). × — убрать последнее, ×× — очистить.\n"
            u"Без кавычек — как раньше: точное имя, иначе подстрока.\n"
            u"Новое имя — шаблон как у переименовать_лист: <<Переменные…>>, [Старое_имя], "
            u"[Заголовок(N)], [Строка(R, C)], дата/время, [Имя_Книги], [GUID8].\n"
            u"Пары сопоставляются по позиции: 1-е старое → 1-е новое и т.д."
        ),
        "hint_lines": 5,
        "fields": (
            {
                "id": "old",
                "label": u"Старое наименование",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"Сумма",
            },
            {
                "id": "new",
                "label": u"Новое наименование",
                "type": "columns_pick",
                "default": u"<<Переменные.Филиал>>_[Старое_имя]",
            },
        ),
    },
    u"переставить_столбцы": {
        "title": u"Переставить столбцы",
        "hint": (
            u"Лист слева обязателен (не «Все»). Столбцы — накопительный выбор из заголовков. "
            u"«Копировать» — оставить исходные и вставить копии; "
            u"«новые» — создать пустые столбцы по «Новые имена копий» и вставить по «Куда». "
            u"Опорный столбец — из списка или вручную (без накопления); "
            u"при переносе если он в списке — исключается."
        ),
        "fields": (
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "default": u"C,Сумма",
            },
            {
                "id": "position",
                "label": u"Куда",
                "type": "combo",
                "choices": (
                    u"_начало_",
                    u"_конец_",
                    u"_перед_",
                    u"_после_",
                ),
                "default": u"_конец_",
            },
            {
                "id": "relative_column",
                "label": u"Опорный столбец",
                "type": "combo",
                "choices_from_headers": True,
                "default": u"Итого",
            },
            {
                "id": "copy",
                "label": u"Копировать",
                "type": "bool",
                "default": False,
                "row_group": u"copy_new",
                "row_group_slot": 0,
            },
            {
                "id": "create_new",
                "label": u"новые",
                "type": "bool",
                "default": False,
                "row_group": u"copy_new",
                "row_group_slot": 1,
            },
            {
                "id": "new_names",
                "label": u"Новые имена копий",
                "default": u"",
            },
        ),
    },
    u"количество_значений": {
        "title": u"Количество значений",
        "hint": (
            u"Новый столбец — сколько строк в диапазоне с тем же ключом "
            u"(столбцы ключа через запятую). По умолчанию trim вкл., регистр не важен."
        ),
        "fields": (
            {"id": "new_column", "label": u"Имя столбца", "default": u"Количество"},
            {
                "id": "key_columns",
                "label": u"Столбцы ключа",
                "type": "columns_pick",
                "default": u"ФИО,Отдел",
            },
            {"id": "trim", "label": u"Обрезать пробелы", "type": "bool", "default": True},
            {
                "id": "case_sensitive",
                "label": u"Учитывать регистр",
                "type": "bool",
                "default": False,
            },
        ),
    },
    u"удалить_дубликаты": {
        "title": u"Удалить дубликаты",
        "hint": (
            u"Ключ — столбцы key_columns (как у «количество_значений»). keep + pre_sort — "
            u"какую строку оставить при повторе ключа.\n"
            u"«На месте» — удалить повторы в текущем диапазоне данных.\n"
            u"«Новый лист» — записать результат на dest_sheet; исходный лист не меняется.\n"
            u"«Смещение на этом листе» — записать блок на этом же листе с ячейки dest_cell "
            u"(A1 = левый верх заголовка, данные ниже); исходные строки не удаляются.\n"
            u"«Маркировка» — не удалять; добавить столбец mark_column, в повторах — «Да»."
        ),
        "hint_lines": 6,
        "form_bottom_pad": 28,
        "dialog_fit_content": True,
        "fields": (
            {
                "id": "key_columns",
                "label": u"Столбцы ключа",
                "type": "columns_pick",
                "default": u"ФИО",
            },
            {
                "id": "exclude_columns",
                "label": u"Исключить столбцы (опц.)",
                "type": "columns_pick",
                "default": u"",
            },
            {
                "id": "skip_empty_cells",
                "label": u"Пропускать пустые ключи",
                "type": "bool",
                "default": True,
            },
            {
                "id": "keep",
                "label": u"Оставить",
                "type": "combo",
                "choices": tuple(label for _code, label in DEDUP_KEEP_CHOICES),
                "default": u"Первую",
            },
            {"id": "pre_sort", "label": u"Сортировка до dedup (JSON)", "default": u""},
            {
                "id": "output",
                "label": u"Режим вывода",
                "type": "combo",
                "choices": tuple(label for _code, label in DEDUP_OUTPUT_CHOICES),
                "default": u"На месте (удалить)",
            },
            {
                "id": "dest_sheet",
                "label": u"Лист результата (новый лист)",
                "default": u"Без_дубликатов",
            },
            {
                "id": "dest_cell",
                "label": u"Якорь A1",
                "default": u"A1",
                "row_group": u"dest_mark",
                "row_group_slot": 0,
            },
            {
                "id": "mark_column",
                "label": u"Столбец маркера",
                "default": u"Дубликаты",
                "row_group": u"dest_mark",
                "row_group_slot": 1,
            },
            {"id": "trim", "label": u"Обрезать пробелы", "type": "bool", "default": True},
            {
                "id": "case_sensitive",
                "label": u"Учитывать регистр",
                "type": "bool",
                "default": False,
            },
        ),
    },
    u"группировать_строки": {
        "title": u"Группировать строки (Group By)",
        "hint": (
            u"Сжимает таблицу до уникальных ключей и считает агрегаты (SUM/COUNT/…).\n"
            u"Не путать с «группировка_по_столбцу» (SUBTOTAL/outline — строки не удаляются).\n"
            u"Столбцы ключа — из заголовков ('кавычки'). Несколько агрегатов — в JSON "
            u"поле aggregations[] (визард задаёт первый; остальные допишите в C).\n"
            u"count без столбца = COUNT(*). «Новый лист» — dest_sheet с A1."
        ),
        "hint_lines": 6,
        "form_bottom_pad": 28,
        "dialog_fit_content": True,
        "fields": (
            {
                "id": "key_columns",
                "label": u"Столбцы ключа",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"Отдел",
            },
            {
                "id": "agg_op",
                "label": u"Агрегат (op)",
                "type": "combo",
                "choices": tuple(label for _code, label in GROUP_BY_ROWS_OP_CHOICES),
                "default": u"Сумма",
            },
            {
                "id": "agg_column",
                "label": u"Столбец агрегата",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"Сумма",
            },
            {
                "id": "agg_as",
                "label": u"Имя столбца результата",
                "default": u"Сумма",
            },
            {
                "id": "output",
                "label": u"Режим вывода",
                "type": "combo",
                "choices": tuple(label for _code, label in GROUP_BY_ROWS_OUTPUT_CHOICES),
                "default": u"На месте",
            },
            {
                "id": "dest_sheet",
                "label": u"Лист результата (новый лист)",
                "default": u"Свод",
            },
            {
                "id": "sort_keys",
                "label": u"Сортировать по ключам",
                "type": "bool",
                "default": False,
            },
            {
                "id": "key_trim",
                "label": u"Обрезать пробелы ключа",
                "type": "bool",
                "default": True,
            },
            {
                "id": "key_case_sensitive",
                "label": u"Учитывать регистр ключа",
                "type": "bool",
                "default": False,
            },
            {
                "id": "skip_empty_keys",
                "label": u"Пропускать пустые ключи",
                "type": "bool",
                "default": True,
            },
        ),
    },
    u"копировать_переместить_лист": {
        "title": u"Копировать/переместить лист",
        "hint": (
            u"Лист слева — когда выполнять шаг.\n"
            u"Галочка «Копировать» снята — перемещение перед/после опорного листа; "
            u"включена — копия в «Имя копии»."
        ),
        "hint_lines": 3,
        "dialog_fit_content": True,
        "fields": (
            {
                "id": "source_sheet",
                "label": u"Лист-источник",
                "type": "combo",
                "choices_from_sheets": True,
                "default": u"",
                "with_auto_check": u"copy",
            },
            {
                "id": "copy",
                "label": u"Копировать",
                "type": "bool",
                "default": False,
                "pair_after": u"source_sheet",
            },
            {"id": "dest_sheet", "label": u"Имя копии", "default": u"Копия_Шаблон"},
            {
                "id": "position",
                "label": u"Позиция",
                "type": "combo",
                "choices": (u"_перед_", u"_после_"),
                "default": u"_после_",
                "row_group": u"anchor_pos",
                "row_group_slot": 0,
            },
            {
                "id": "anchor_sheet",
                "label": u"Опорный лист",
                "type": "combo",
                "choices_from_sheets": True,
                "default": u"",
                "row_group": u"anchor_pos",
                "row_group_slot": 1,
            },
            {"id": "columns", "label": u"Столбцы (опц.)", "default": u""},
            {
                "id": "skip_empty",
                "label": u"Пропускать пустые строки",
                "type": "bool",
                "default": False,
            },
            {
                "id": "skip_empty_columns",
                "label": u"Колонки для пустоты (опц.)",
                "default": u"",
            },
            {"id": "header_row", "label": u"Строка заголовков", "default": u"1"},
        ),
    },
    u"активировать_лист": {
        "title": u"Активировать лист",
        "hint": (
            u"Выберите лист книги — он станет активным.\n"
            u"Если лист скрыт — будет показан. Курсор ставится в A1.\n"
            u"Цвет ярлычка — опционально (имя или #hex; «нет» — сброс)."
        ),
        "hint_lines": 3,
        "dialog_fit_content": True,
        "fields": (
            {
                "id": "target_sheet",
                "label": u"Лист",
                "type": "combo",
                "choices_from_sheets": True,
                "default": u"",
            },
            {
                "id": "tab_color",
                "label": u"Цвет ярлычка",
                "type": "combo",
                "choices": (u"", u"нет") + tuple(
                    c for c in _header_plus_height_color_choices() if c
                ),
                "default": u"",
                "hint": u"Пусто — не менять; «нет» — сброс цвета",
            },
        ),
    },
    u"копирование_диапазонов": {
        "title": u"Копирование диапазонов",
        "hint": (
            u"Источник → приёмник(и): Замена поверх якоря или Вставка со сдвигом.\n"
            u"По умолчанию — значения; «Формулы» — с автосдвигом ссылок Calc. "
            u"Приёмник вовлекается в следующие шаги (галочку можно снять)."
        ),
        "hint_lines": 2,
        "dialog_fit_content": True,
        "fields": (
            {
                "id": "source_sheet",
                "label": u"Лист-источник",
                "type": "combo",
                "choices_from_sheets": True,
                "default": u"",
            },
            {
                "id": "source_range",
                "label": u"Диапазон A1",
                "default": u"",
                "hint": u"A1:D10, B:B, 2:5",
                "row_group": u"src_range_rows",
                "row_group_slot": 0,
            },
            {
                "id": "source_rows",
                "label": u"Строки (альт.)",
                "default": u"",
                "hint": u"2:50",
                "row_group": u"src_range_rows",
                "row_group_slot": 1,
            },
            {
                "id": "source_columns",
                "label": u"Столбцы (альт.)",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
            {
                "id": "dest_sheets",
                "label": u"Листы-приёмники",
                "default": u"",
                "hint": u"через , или ;",
            },
            {
                "id": "dest_cell",
                "label": u"Якорь A1",
                "default": u"A1",
                "row_group": u"dest_mode",
                "row_group_slot": 0,
            },
            {
                "id": "mode",
                "label": u"Режим",
                "type": "combo",
                "choices": tuple(label for _c, label in COPY_RANGES_MODE_CHOICES),
                "default": u"Замена",
                "row_group": u"dest_mode",
                "row_group_slot": 1,
            },
            {
                "id": "insert_axis",
                "label": u"Ось сдвига",
                "type": "combo",
                "choices": tuple(label for _c, label in COPY_RANGES_AXIS_CHOICES),
                "default": u"Авто",
                "row_group": u"axis_content",
                "row_group_slot": 0,
            },
            {
                "id": "content",
                "label": u"Содержимое",
                "type": "combo",
                "choices": tuple(label for _c, label in COPY_RANGES_CONTENT_CHOICES),
                "default": u"Значения",
                "row_group": u"axis_content",
                "row_group_slot": 1,
            },
            {
                "id": "with_formatting",
                "label": u"С оформлением",
                "type": "bool",
                "default": False,
                "row_group": u"flags_fmt_clear",
                "row_group_slot": 0,
            },
            {
                "id": "clear_source",
                "label": u"Очистить источник",
                "type": "bool",
                "default": False,
                "row_group": u"flags_fmt_clear",
                "row_group_slot": 1,
            },
            {
                "id": "create_missing_dest",
                "label": u"Создать лист, если нет",
                "type": "bool",
                "default": False,
                "row_group": u"flags_create_involve",
                "row_group_slot": 0,
            },
            {
                "id": "involve_dest",
                "label": u"Вовлечь в следующие шаги",
                "type": "bool",
                "default": True,
                "row_group": u"flags_create_involve",
                "row_group_slot": 1,
            },
            {
                "id": "header_row",
                "label": u"Строка заголовков",
                "default": u"1",
            },
        ),
    },
    u"создать_лист": {
        "title": u"Создать лист",
        "hint": (
            u"Имя — вручную или лямбда sheets. Заголовки через ,/; ; в '…' — разделители в имени.\n"
            u"Данные: С листа / Вручную (| между строками; <<Переменные.список>> → столбец)."
        ),
        "hint_lines": 2,
        "dialog_fit_content": True,
        "fields": (
            {
                "id": "name_mode",
                "label": u"Имя листа",
                "type": "combo",
                "choices": tuple(label for _c, label in CREATE_SHEET_NAME_MODE_CHOICES),
                "default": u"Ввести вручную",
                "row_group": u"name_hdr",
                "row_group_slot": 0,
            },
            {
                "id": "header_row",
                "label": u"Строка заголовка",
                "default": u"1",
                "hint": u"Данные — со следующей",
                "row_group": u"name_hdr",
                "row_group_slot": 1,
            },
            {
                "id": "sheet_name",
                "label": u"Имя вручную",
                "type": "lambda_pick",
                "default": u"Новый_лист",
                "hint": u'При «Вычислить»: lambda sheets: "Отчет_" + str(len(sheets))',
            },
            {
                "id": "columns",
                "label": u"Имена колонок",
                "default": u"",
                "hint": u"Код, Сумма; 'A, B' — запятая в имени",
            },
            {
                "id": "data_mode",
                "label": u"Источник данных",
                "type": "combo",
                "choices": tuple(label for _c, label in CREATE_SHEET_DATA_MODE_CHOICES),
                "default": u"Нет",
                "row_group": u"data_src",
                "row_group_slot": 0,
            },
            {
                "id": "source_sheet",
                "label": u"Лист-источник",
                "type": "combo",
                "choices_from_sheets": True,
                "default": u"",
                "row_group": u"data_src",
                "row_group_slot": 1,
            },
            {
                "id": "source_range",
                "label": u"Диапазон A1",
                "default": u"",
                "hint": u"A2:D100",
                "with_bool_check": u"values_only",
                "edit_frac": 0.45,
            },
            {
                "id": "values_only",
                "label": u"Только значения",
                "type": "bool",
                "default": True,
                "pair_after": u"source_range",
            },
            {
                "id": "rows",
                "label": u"Строки вручную",
                "default": u"",
                "tall": True,
                "hint": (
                    u"a, 'b, c', <<Переменные.x>> | d, e | "
                    u"<<Переменные.список>>, константа"
                ),
            },
        ),
    },
    u"отправить_по_почте": {
        "title": u"Отправить по почте",
        "hint": (
            u"Письмо не уходит само — откроется окно клиента. "
            u"Имя файла: как указано + суффикс даты/времени; пусто — имя книги."
        ),
        "hint_lines": 1,
        "blocks_without_sheet": True,
        "single_form": True,
        "dialog_fit_content": True,
        "compact_fields": True,
        "fields": (
            {
                "id": "attach",
                "label": u"Что вложить",
                "type": "combo",
                "choices": tuple(label for _c, label in SEND_MAIL_ATTACH_CHOICES),
                "default": u"Всю книгу",
                "row_group": u"attach_sheets",
                "row_group_slot": 0,
                "row_group_frac": 0.42,
            },
            {
                "id": "sheets",
                "label": u"Листы",
                "default": u"",
                "hint": u"При «Выбранные листы»",
                "row_group": u"attach_sheets",
                "row_group_slot": 1,
            },
            {
                "id": "format",
                "label": u"Формат",
                "type": "combo",
                "choices": tuple(label for _c, label in SEND_MAIL_FORMAT_CHOICES),
                "default": u"ODS",
                "row_group": u"fmt_name",
                "row_group_slot": 0,
                "row_group_frac": 0.28,
            },
            {
                "id": "filename",
                "label": u"Имя файла",
                "default": u"",
                "hint": u"Без расширения; + _ГГГГММДД_ЧЧММСС",
                "row_group": u"fmt_name",
                "row_group_slot": 1,
            },
            {
                "id": "to",
                "label": u"Кому",
                "default": u"",
                "hint": u"Адреса через ;",
            },
            {
                "id": "to_source_sheet",
                "label": u"Лист для to",
                "type": "combo",
                "choices_from_sheets": True,
                "default": u"",
                "row_group": u"to_src",
                "row_group_slot": 0,
            },
            {
                "id": "to_source_range",
                "label": u"Диапазон to",
                "default": u"",
                "hint": u"B2:B200",
                "row_group": u"to_src",
                "row_group_slot": 1,
            },
            {
                "id": "cc",
                "label": u"Копия",
                "default": u"",
                "row_group": u"cc_bcc",
                "row_group_slot": 0,
            },
            {
                "id": "bcc",
                "label": u"Скрытая",
                "default": u"",
                "row_group": u"cc_bcc",
                "row_group_slot": 1,
            },
            {
                "id": "subject",
                "label": u"Тема",
                "default": u"",
            },
            {
                "id": "body",
                "label": u"Текст",
                "default": u"",
            },
            {
                "id": "os",
                "label": u"ОС",
                "type": "combo",
                "choices": tuple(label for _c, label in SEND_MAIL_OS_CHOICES),
                "default": u"Авто",
                "row_group": u"os_client",
                "row_group_slot": 0,
            },
            {
                "id": "client",
                "label": u"Клиент",
                "type": "combo",
                "choices": tuple(label for _c, label in SEND_MAIL_CLIENT_CHOICES),
                "default": u"Авто",
                "row_group": u"os_client",
                "row_group_slot": 1,
            },
        ),
    },
    u"email": {
        "alias_of": u"отправить_по_почте",
    },
    u"mailto": {
        "alias_of": u"отправить_по_почте",
    },
    u"send_mail": {
        "alias_of": u"отправить_по_почте",
    },
    u"почта": {
        "alias_of": u"отправить_по_почте",
    },
    u"копировать_лист": {
        "alias_of": u"копировать_переместить_лист",
    },
    u"копировать_диапазон": {
        "alias_of": u"копирование_диапазонов",
    },
    u"copy_ranges": {
        "alias_of": u"копирование_диапазонов",
    },
    u"copy_range": {
        "alias_of": u"копирование_диапазонов",
    },
    u"вставить_диапазон": {
        "alias_of": u"копирование_диапазонов",
    },
    u"преобразовать_числа": {
        "title": u"Преобразовать числа",
        "hint": (
            u"Столбцы — выбор из заголовков ('кавычки', накопление, × / ××); "
            u"индексы 1-based и буквы A допустимы. Пусто — все столбцы строки."
        ),
        "hint_lines": 3,
        "fields": (
            {
                "id": "columns",
                "label": u"Столбцы",
                "type": "columns_pick",
                "choices_from_headers": True,
                "default": u"",
            },
        ),
    },
    u"цвет_текста_по_значению": {
        "title": u"Цвет текста по значению",
        "hint": u"Подстрока в заголовке столбца-источника (по умолчанию «статус»).",
        "fields": (
            {"id": "marker", "label": u"Маркер столбца", "default": u"статус"},
        ),
    },
}
_INNER_CODEC_DIALOG_REGISTRY = {
    u"сортировка": {
        "title": u"Сортировка",
        "hint": u"Примеры: «Сводная | Сумма -> - , Дата -> + ; Orders | Сумма -> -» или без листа «Сумма -> -»",
    },
    u"текстовые_операции": {
        "title": u"Текстовые операции",
        "hint": (
            u"Стек ops + столбцы; row_filter per-op; нетекст/ошибки по-русски. "
            u"Кнопка «Справка» — полный каталог ops и примеры JSON."
        ),
    },
    u"нормализация_текста": {
        "title": u"Текстовые операции",
        "hint": u"Legacy-алиас «текстовые_операции».",
        "alias_of": u"текстовые_операции",
    },
    u"раскрасить": {
        "title": u"Раскрасить блоки",
        "hint": u"Примеры: «Сводная | Отдел -> голубой%белый/+!граница ; Orders | A,B -> #E2EFDA/-»",
        "alias_of": u"раскрасить_блоки",
    },
    u"раскрасить_блоки": {
        "title": u"Раскрасить блоки",
        "hint": u"Примеры: «Сводная | Отдел -> голубой%белый/+!граница»",
    },
}

# Каталог ops / подписи для визуального конструктора текстовые_операции
NORMALIZE_OP_CHOICES = (
    (u"trim", u"Обрезать пробелы по краям"),
    (u"collapse_ws", u"Убрать лишние пробелы"),
    (u"substring", u"Подстрока"),
    (u"clear", u"Очистка"),
    (u"case", u"Регистр"),
    (u"compute", u"Вычислить"),
    (u"lang_homoglyphs", u"Омоглифы RU/EN"),
    (u"reverse", u"Перевернуть текст"),
    (u"base64_encode", u"Base64: кодировать"),
    (u"base64_decode", u"Base64: декодировать"),
    (u"sha256", u"SHA-256"),
    (u"guid", u"GUID"),
    (u"encrypt", u"Зашифровать"),
    (u"decrypt", u"Расшифровать"),
)
NORMALIZE_OP_LABEL = {code: label for code, label in NORMALIZE_OP_CHOICES}
NORMALIZE_OP_CODE = {label.casefold(): code for code, label in NORMALIZE_OP_CHOICES}
for _nop_code, _nop_label in NORMALIZE_OP_CHOICES:
    NORMALIZE_OP_CODE[_nop_code.casefold()] = _nop_code
NORMALIZE_OP_CODE[u"вычислить"] = u"compute"
NORMALIZE_OP_CODE[u"calc"] = u"compute"
NORMALIZE_OP_CODE[u"calculate"] = u"compute"
NORMALIZE_OP_CODE[u"expr"] = u"compute"
NORMALIZE_CASE_MODES = (
    (u"lower", u"нижний"),
    (u"upper", u"верхний"),
    (u"proper", u"как имена (Proper)"),
    (u"sentence", u"предложение"),
    (u"toggle", u"инверсия регистра"),
)
NORMALIZE_CASE_LABEL = {code: label for code, label in NORMALIZE_CASE_MODES}
NORMALIZE_CASE_CODE = {label.casefold(): code for code, label in NORMALIZE_CASE_MODES}
NORMALIZE_NON_TEXT_CHOICES = (
    (u"skip", u"пропуск"),
    (u"coerce", u"как текст"),
    (u"error", u"ошибка"),
)
NORMALIZE_NON_TEXT_LABEL = {code: label for code, label in NORMALIZE_NON_TEXT_CHOICES}
NORMALIZE_NON_TEXT_CODE = {label.casefold(): code for code, label in NORMALIZE_NON_TEXT_CHOICES}
NORMALIZE_ON_ERROR_CHOICES = (
    (u"keep", u"оставить до ошибки"),
    (u"empty", u"очистить ячейку"),
    (u"stop", u"остановить"),
)
NORMALIZE_ON_ERROR_LABEL = {code: label for code, label in NORMALIZE_ON_ERROR_CHOICES}
NORMALIZE_ON_ERROR_CODE = {label.casefold(): code for code, label in NORMALIZE_ON_ERROR_CHOICES}
NORMALIZE_ROW_FILTER_CHOICES = (
    (u"all", u"все"),
    (u"non_empty", u"непустые"),
    (u"empty", u"пустые"),
    (u"lambda", u"лямбда выражение"),
)
NORMALIZE_ROW_FILTER_LABEL = {code: label for code, label in NORMALIZE_ROW_FILTER_CHOICES}
NORMALIZE_ROW_FILTER_CODE = {label.casefold(): code for code, label in NORMALIZE_ROW_FILTER_CHOICES}
NORMALIZE_ROW_FILTER_LAMBDA_HINT = (
    u'lambda rows: str(rows("Столбец") or "").strip() != ""'
)
NORMALIZE_ROW_FILTER_LAMBDA_HINT_UPD = (
    u'lambda rows: not upd_rows_seen("Столбец", rows("Столбец"))'
)
NORMALIZE_COMPUTE_LAMBDA_HINT = (
    u'lambda rows: tr(rows("Столбец") or "-").strip()'
)

_VLOOKUP_CELL_HINTS = (
    u"Левая таблица join: Лист|Начало_строки|Начало_столбца|Строка_заголовков.\n"
    u"Имя листа — как в источнике; при копировании подставится префикс N_.",
    u"Правая (справочная) таблица — тот же формат через |.",
    u"Общий ключ: Ключ_сравнения: A,B\n"
    u"Разные: Ключ_сравнения_левый: … # Ключ_сравнения_правый: …",
    u"Столбцы правой таблицы для подтягивания (через , или ;).\n"
    u"Переименование на левом: 'ФИО -> Новое_имя' или 'ФИО = Новое_имя';\n"
    u"просто 'ФИО' — то же имя, что справа.",
    u"Подсветка: #FFF2CC, серый, заливка/шрифт (как у зебры); «Нет» — без подсветки.\n"
    u"Заполнять_дубли: Да/Нет (всегда в ячейке).\n"
    u"Мульти_совпадение_в_одну_ячейку: Да/Нет — все совпадения в одну ячейку через перенос строки.\n"
    u"Заполнитель_для_не_найдено: по умолчанию #Н/Д (#N/A, =NA()); _ПУСТО_ или _EMPTY_ — пустая ячейка.\n"
    u"Булево: +/-, 1/0, true/false, yes/no, истина/ложь (регистр не важен).",
)

# Полная справка поддиалога ВПР (кнопка «Справка»).
_VLOOKUP_HELP_TITLE = u"Справка — параметры ВПР"
_VLOOKUP_HELP_TEXT = (
    u"══════════════════════════════════════\n"
    u"ВПР — соединение двух листов книги результата\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u"По смыслу это JOIN: по ключам сопоставить строки левой и правой\n"
    u"таблицы и перенести выбранные столбцы. Выполняется в pipeline\n"
    u"между шагами «Постобработка_Диапазон»; «Постобработка_Строка» —\n"
    u"только после всех ВПР. Предпочтительная запись — JSON в колонке C.\n"
    u"\n"
    u"══════════════════════════════════════\n"
    u"ЛЕВАЯ / ПРАВАЯ ТАБЛИЦА\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u"• Лист — имя листа в книге результата (после сбора уже может быть\n"
    u"  с префиксом N_, напр. 1_Справочник_цен). Выберите из списка.\n"
    u"• Лев.диапазон / Прав.диапазон — опционально в нотации A1:\n"
    u"    C2:F40   — прямоугольник (заголовок = первая строка диапазона);\n"
    u"    1:200    — строки 1…200, все столбцы used-area по ширине;\n"
    u"    D:H      — столбцы D…H, строки по used-area.\n"
    u"  Пустой диапазон — вся used-area листа (как раньше).\n"
    u"  При сохранении JSON пишутся derived start_row / start_col / header_row.\n"
    u"\n"
    u"══════════════════════════════════════\n"
    u"ТИП СОЕДИНЕНИЯ (join_type)\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u"• Левое (left) — по умолчанию. Все строки левой таблицы;\n"
    u"  справа подтягиваются совпадения. Нет пары → «значение для не найдено».\n"
    u"• Внутреннее (inner) — только строки, у которых есть ключ справа.\n"
    u"  Строки слева без пары удаляются с левого листа.\n"
    u"• Полное (full) — два прохода:\n"
    u"  1) как left: колонки ← из правой на левую (extract_columns);\n"
    u"  2) зеркало: колонки → из левой на правую (extract_columns_right);\n"
    u"  плюс на левый лист добавляются строки, которые есть только справа.\n"
    u"  Поле «Колонки в Прав. из Лев. (→)» видно только при full.\n"
    u"\n"
    u"══════════════════════════════════════\n"
    u"КЛЮЧИ\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u"• «Одинаковые ключи» — один список для обеих сторон\n"
    u"  (JSON: keys_same=true, достаточно keys_left).\n"
    u"• Иначе заполните «Ключи (общие)» = левые и «Ключи (прав.)» = правые\n"
    u"  в одном порядке (составной ключ: несколько столбцов).\n"
    u"• Токены столбцов: из списка заголовков → обычно 'Имя' (точное\n"
    u"  совпадение); без кавычек — подстрока; можно буквы/номера (A, 1).\n"
    u"  Кнопки × / ×× — убрать последний токен / очистить.\n"
    u"\n"
    u"Пример: ключ «Код_позиции» на обоих листах →\n"
    u"  keys_left / keys_right: [\"'Код_позиции'\"]\n"
    u"\n"
    u"══════════════════════════════════════\n"
    u"КОЛОНКИ ДЛЯ ИЗВЛЕЧЕНИЯ\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u"• «Колонки в Лев. из Прав. (←)» (extract_columns) — что взять\n"
    u"  с правой таблицы и дописать справа на левый лист.\n"
    u"• Режим колонок (extract_mode):\n"
    u"    Новые колонки (new) — по умолчанию: создать столбцы справа;\n"
    u"    Заменить значения (replace) — писать в СУЩЕСТВУЮЩИЕ колонки\n"
    u"      слева (суффикс _R не применяется); первый матч очищает ячейку,\n"
    u"      дальше — склейка через перевод строки; нет колонки → как new;\n"
    u"    Объединить значения (merge) — как replace, но первый матч\n"
    u"      не стирает текущее значение слева, а дописывает через \\n.\n"
    u"• «Колонки в Прав. из Лев. (→)» (extract_columns_right) — только\n"
    u"  при join_type=full: зеркальный перенос на правый лист.\n"
    u"• Переименование на приёмнике:\n"
    u"    'ФИО'              — то же имя, что у источника;\n"
    u"    'ФИО -> Новое'     — или 'ФИО = Новое' — имя колонки на левом;\n"
    u"    B / 2              — столбец по букве/номеру.\n"
    u"\n"
    u"Пример:\n"
    u"  extract_columns: [\"'Наименование'\", \"'Цена -> Цена_справ'\"]\n"
    u"\n"
    u"══════════════════════════════════════\n"
    u"МНОЖЕСТВЕННЫЕ СОВПАДЕНИЯ (multi_match)\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u"Если справа несколько строк с одним ключом:\n"
    u"• Все (all) — по умолчанию; дальше влияют «Дубли» и «Мульти»;\n"
    u"• первый (first) — только первая строка справа;\n"
    u"• последний (last) — только последняя.\n"
    u"\n"
    u"══════════════════════════════════════\n"
    u"ИТОГОВЫЕ ПАРАМЕТРЫ\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u"• Цвет для новых колонок — подсветка добавленных столбцов\n"
    u"  (hex #FFF2CC, пресеты голубой/серый/…; «Нет» / пусто — без цвета).\n"
    u"• Значение для не найдено — что писать, если ключа нет справа\n"
    u"  (left/full). По умолчанию #Н/Д. Спецзначения:\n"
    u"    _ПУСТО_ / _EMPTY_ — пустая ячейка (не #N/A).\n"
    u"• Дубли (fill_duplicates) — при multi_match=all и нескольких\n"
    u"  совпадениях: Да — заполнять каждую «дублирующую» строку;\n"
    u"  Нет — только первую (остальные — заполнитель/пусто по правилам).\n"
    u"• Мульти (multi_match_one_cell) — все совпадения в ОДНУ ячейку\n"
    u"  через перевод строки (вместо нескольких строк результата).\n"
    u"• СЖ_ПРОБЕЛЫ (trim_keys) — обрезать пробелы у значений ключа\n"
    u"  перед сравнением (слева и справа).\n"
    u"• Суффикс _R (column_suffix) — к именам новых колонок с правой\n"
    u"  добавлять суффикс, чтобы не конфликтовать с уже существующими\n"
    u"  на левом (по умолчанию Да). В режимах replace/merge для уже\n"
    u"  найденных колонок слева суффикс игнорируется.\n"
    u"• Кол-во Совпадений (match_count) — после колонок извлечения\n"
    u"  добавить столбец «Кол-во Совпадений» с числом матчей справа\n"
    u"  (0 / 1 / N). В режиме «Дубли» одно и то же N на каждой строке\n"
    u"  дубля; в «Мульти» — одно значение в ячейке. По умолчанию Выкл.\n"
    u"\n"
    u"══════════════════════════════════════\n"
    u"ПРИМЕР JSON (колонка C)\n"
    u"══════════════════════════════════════\n"
    u"\n"
    u'[{\n'
    u'  "v": 1,\n'
    u'  "fn": "впр",\n'
    u'  "left":  {"sheet": "Заказы", "range": "A1:G200"},\n'
    u'  "right": {"sheet": "Справочник_цен", "range": "A1:D50"},\n'
    u'  "join_type": "left",\n'
    u'  "keys_same": true,\n'
    u'  "keys_left": ["\'Код_позиции\'"],\n'
    u'  "keys_right": ["\'Код_позиции\'"],\n'
    u'  "extract_columns": ["\'Наименование\'", "\'Цена\'"],\n'
    u'  "highlight_color": "голубой",\n'
    u'  "fill_duplicates": true,\n'
    u'  "not_found_fill": "#Н/Д",\n'
    u'  "multi_match": "all",\n'
    u'  "multi_match_one_cell": false,\n'
    u'  "trim_keys": false,\n'
    u'  "column_suffix": true\n'
    u'}]\n'
    u"\n"
    u"inner — только совпадения; full — добавьте extract_columns_right\n"
    u"и join_type: \"full\". Legacy B–F ещё читается; range / full / inner /\n"
    u"multi_match — только через JSON.\n"
    u"\n"
    u"Подробнее: docs/06_PARAM_WIZARD.md §6 (vlookup), docs/13_JSON_PARAMS.md."
)
_VLOOKUP_COLOR_PRESETS = (
    u"#FFFF00",
    u"#FFF2CC",
    u"excel_highlight",
    u"серый",
    u"голубой",
    u"голубой/синий",
    u"белый/черный",
)
_VLOOKUP_COLOR_PRESET_EMPTY = u"— выберите пресет —"
_VLOOKUP_JOIN_TYPE_CHOICES = (
    (u"left", u"Левое соединение"),
    (u"inner", u"Внутреннее соединение"),
    (u"full", u"Полное соединение"),
)
_VLOOKUP_JOIN_TYPE_LABEL = {code: label for code, label in _VLOOKUP_JOIN_TYPE_CHOICES}
_VLOOKUP_JOIN_TYPE_CODE = {label.casefold(): code for code, label in _VLOOKUP_JOIN_TYPE_CHOICES}
for _alias in (u"левое", u"left_join", u"left outer", u"left_outer"):
    _VLOOKUP_JOIN_TYPE_CODE[_alias] = u"left"
for _alias in (
    u"внутреннее",
    u"inner_join",
    u"inner join",
    u"внутреннее_соединение",
    u"внутреннее соединение",
):
    _VLOOKUP_JOIN_TYPE_CODE[_alias] = u"inner"
for _alias in (
    u"полное",
    u"full_outer",
    u"full outer",
    u"полное_соединение",
    u"полное соединение",
):
    _VLOOKUP_JOIN_TYPE_CODE[_alias] = u"full"
_VLOOKUP_MULTI_MATCH_CHOICES = (
    (u"all", u"Все"),
    (u"first", u"первый"),
    (u"last", u"последний"),
)
_VLOOKUP_MULTI_MATCH_LABEL = {code: label for code, label in _VLOOKUP_MULTI_MATCH_CHOICES}
_VLOOKUP_MULTI_MATCH_CODE = {label.casefold(): code for code, label in _VLOOKUP_MULTI_MATCH_CHOICES}
for _alias in (u"все", u"*", u"all_rows", u"все_совпадения"):
    _VLOOKUP_MULTI_MATCH_CODE[_alias] = u"all"
for _alias in (u"первая", u"first_row", u"первое"):
    _VLOOKUP_MULTI_MATCH_CODE[_alias] = u"first"
for _alias in (u"последняя", u"last_row", u"последнее"):
    _VLOOKUP_MULTI_MATCH_CODE[_alias] = u"last"
# Режим записи колонок из правой таблицы в левую (JSON: extract_mode).
_VLOOKUP_EXTRACT_MODE_CHOICES = (
    (u"new", u"Новые колонки"),
    (u"replace", u"Заменить значения"),
    (u"merge", u"Объединить значения"),
)
_VLOOKUP_EXTRACT_MODE_LABEL = {code: label for code, label in _VLOOKUP_EXTRACT_MODE_CHOICES}
_VLOOKUP_EXTRACT_MODE_CODE = {label.casefold(): code for code, label in _VLOOKUP_EXTRACT_MODE_CHOICES}
for _alias in (u"новые", u"new_columns", u"новые_колонки", u"создать"):
    _VLOOKUP_EXTRACT_MODE_CODE[_alias] = u"new"
for _alias in (u"заменить", u"replace_values", u"замена", u"overwrite"):
    _VLOOKUP_EXTRACT_MODE_CODE[_alias] = u"replace"
for _alias in (u"объединить", u"merge_values", u"append", u"склеить"):
    _VLOOKUP_EXTRACT_MODE_CODE[_alias] = u"merge"
# Имя колонки счётчика совпадений (JSON: match_count=true).
_VLOOKUP_MATCH_COUNT_COLUMN = u"Кол-во Совпадений"
_VLOOKUP_NOT_FOUND_FROM_F_RE = re.compile(
    r"Заполнитель_для_не_найдено\s*:\s*([^;\r\n]+)",
    re.IGNORECASE | re.UNICODE,
)
_PIVOT_DATA_FUNCTIONS = (
    (u"SUM", u"Сумма"),
    (u"COUNT", u"Количество"),
    (u"AVERAGE", u"Среднее"),
    (u"MAX", u"Максимум"),
    (u"MIN", u"Минимум"),
    (u"PRODUCT", u"Произведение"),
    (u"STDDEV", u"Станд. отклонение"),
    (u"VAR", u"Дисперсия"),
    (u"LIST_ROWS", u"Список_строк"),
)
_PIVOT_FN_LABEL = {code: label for code, label in _PIVOT_DATA_FUNCTIONS}
_PIVOT_FN_CODE = {label.casefold(): code for code, label in _PIVOT_DATA_FUNCTIONS}
# Синонимы псевдофункции Список_строк (не UNO GeneralFunction).
for _alias in (
    u"list_rows",
    u"list",
    u"детализация",
    u"detail_rows",
    u"detail",
):
    _PIVOT_FN_CODE[_alias] = u"LIST_ROWS"
_PIVOT_LIST_ROWS_FIELD_SEP = u" | "
_PIVOT_LIST_ROWS_RECORD_SEP = u"\n"
_PIVOT_LIST_ROWS_MAX_PER_CELL = 200
_PIVOT_LIST_FIELDS_KEYS = (
    u"list_fields",
    u"detail_fields",
    u"detail_columns",
    u"поля_списка",
    u"колонки_списка",
)
_PIVOT_REFRESH_BUSY = {}

_WP_SCENARIOS = None

# --- Конструктор лямбда-выражения (субвизард) ---
LAMBDA_BUILDER_HINT = (
    u"Sugar: rows(\"Кол\"), rows(...)[-1], upd_rows / upd_rows_seen / upd_rows_len, tr(…).\n"
    u"<<Переменные.имя>> — вставляйте в кавычках (кнопка «В кавычках»).\n"
    u"Сырой текст правится в редакторе; «Проверить» — санация + compile до OK."
)
LAMBDA_BUILDER_KIND_CHOICES = (
    (u"rows_predicate", u"Фильтр строк"),
    (u"rows_value", u"Значение ячейки"),
    (u"name", u"Имя столбца"),
    (u"sheets", u"Список листов"),
)
LAMBDA_BUILDER_KIND_META = {
    u"rows_predicate": {
        u"kind": u"rows_predicate",
        u"label": u"Фильтр строк",
        u"signature": u"lambda rows: ",
        u"compile": u"rows",
    },
    u"rows_value": {
        u"kind": u"rows_value",
        u"label": u"Значение ячейки",
        u"signature": u"lambda rows: ",
        u"compile": u"rows",
    },
    u"name": {
        u"kind": u"name",
        u"label": u"Имя столбца",
        u"signature": u"lambda sheet, headers: ",
        u"compile": u"name",
    },
    u"sheets": {
        u"kind": u"sheets",
        u"label": u"Список листов",
        u"signature": u"lambda sheets: ",
        u"compile": u"sheets",
    },
}
LAMBDA_BUILDER_PRESETS = {
    u"rows_predicate": (
        (u"Столбец непустой", u'lambda rows: str(rows("Столбец") or "").strip() != ""'),
        (u"Столбец пустой", u'lambda rows: str(rows("Столбец") or "").strip() == ""'),
        (
            u"Значение = переменная",
            u'lambda rows: tr(rows("Столбец")) == "<<Переменные.Имя>>"',
        ),
        (u"Смена группы", u'lambda rows: rows("Столбец") != rows("Столбец")[-1]'),
        (
            u"Уникальность (seen)",
            u'lambda rows: not upd_rows_seen("Столбец", rows("Столбец"))',
        ),
        (u"Первая строка", u"lambda rows: idx == 0"),
        (u"Пустой / свой", u"lambda rows: "),
    ),
    u"rows_value": (
        (u"Как есть (столбец)", u'lambda rows: tr(rows("Столбец"))'),
        (u"Trim + fallback", u'lambda rows: tr(rows("Столбец") or "-").strip()'),
        (u"Константа из переменной", u'lambda rows: "<<Переменные.Имя>>"'),
        (
            u"Склеить два столбца",
            u'lambda rows: tr(rows("A") or "") + " " + tr(rows("B") or "")',
        ),
        (u"Пустой / свой", u"lambda rows: "),
    ),
    u"name": (
        (u"Префикс + sheet", u'lambda sheet, headers: "Копия_" + sheet'),
        (u"Первый заголовок", u"lambda sheet, headers: headers[0] if headers else sheet"),
        (u"Пустой / свой", u"lambda sheet, headers: "),
    ),
    u"sheets": (
        (u"Все имена", u"lambda sheets: list(sheets)"),
        (u"Первый лист", u"lambda sheets: sheets[0] if sheets else \"\""),
        (u"Пустой / свой", u"lambda sheets: "),
    ),
}
LAMBDA_BUILDER_SUGAR_BUTTONS = (
    (u"rows", u"rows(…)"),
    (u"rows_prev", u"[-1]"),
    (u"upd_rows", u"upd_rows"),
    (u"upd_seen", u"seen"),
    (u"upd_len", u"len()"),
    (u"tr", u"tr(…)"),
    (u"str", u"str(…)"),
    (u"strip", u".strip()"),
    (u"empty", u"пусто"),
    (u"not_empty", u"непусто"),
)
LAMBDA_BUILDER_OP_BUTTONS = (
    (u"==", u"=="),
    (u"!=", u"!="),
    (u">", u">"),
    (u"<", u"<"),
    (u"in", u"in"),
    (u"not", u"not"),
    (u"and", u"and"),
    (u"or", u"or"),
)
