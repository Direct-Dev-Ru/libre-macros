# -*- coding: utf-8 -*-
"""Константы полноэкранного визарда «Все параметры» (пайплайн листа)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.789"
ALL_PARAMS_BTN_LABEL = u"Все параметры"
ALL_PARAMS_WIZARD_TITLE = u"Все параметры — пайплайн обработки"
ALL_PARAMS_WIZARD_HINT = (
    u"Сверху — меню шагов (как Power Query). "
    u"Слева — пайплайн; Добавить / Удалить / Скопировать / ↑ / ↓. "
    u"Справа — A/B/C/D; «Параметры C…» / «Параметры D…»."
)

ALL_PARAMS_BTN_C_LABEL = u"Параметры C…"
ALL_PARAMS_BTN_D_LABEL = u"Параметры D…"

# Все кнопки — одна высота, жирный шрифт.
ALL_PARAMS_BTN_H = 30
ALL_PARAMS_BTN_FONT_H = 11
ALL_PARAMS_TOOLBAR_BTN_H = ALL_PARAMS_BTN_H
ALL_PARAMS_HINT_H = 22
ALL_PARAMS_HINT_TOP_GAP = 2
ALL_PARAMS_TOOLBAR_GAP = 4
ALL_PARAMS_SPLIT_GAP = 8
ALL_PARAMS_LIST_LBL_H = 16
ALL_PARAMS_LIST_WIDTH_RATIO = 0.40
ALL_PARAMS_LIST_MIN_W = 220
ALL_PARAMS_LIST_MIN_H = 60
ALL_PARAMS_LIST_BOTTOM_GAP = 4
ALL_PARAMS_STATUS_H = 16
ALL_PARAMS_FOOTER_BTN_W = 110
ALL_PARAMS_DETAIL_BTN_W = 130
ALL_PARAMS_LIST_TOOLBAR_GAP = 3
ALL_PARAMS_DETAIL_ROW_H = 20
# Combo A/B — не уменьшать (требование UX).
ALL_PARAMS_DETAIL_COMBO_H = 28
ALL_PARAMS_DETAIL_LBL_H = 12
ALL_PARAMS_DETAIL_GAP = 3
ALL_PARAMS_DETAIL_SECTION_GAP = 4
ALL_PARAMS_DETAIL_BTN_ROW_H = ALL_PARAMS_BTN_H
ALL_PARAMS_DETAIL_BTN_FONT_H = ALL_PARAMS_BTN_FONT_H
ALL_PARAMS_APPLY_BTN_H = ALL_PARAMS_BTN_H
ALL_PARAMS_APPLY_BTN_W = 200
# Шрифт combo A/B (pt).
ALL_PARAMS_NAME_FONT_H = 12
ALL_PARAMS_VALUE_FONT_H = 11
# C/D/описание сжаты, чтобы освободить ~25% сверху под PQ-меню.
ALL_PARAMS_DETAIL_EXTRA_MIN_H = 28
ALL_PARAMS_DETAIL_COND_MIN_H = 26
ALL_PARAMS_DETAIL_DESC_MIN_H = 32
ALL_PARAMS_DETAIL_EXTRA_SHARE = 0.40
ALL_PARAMS_DETAIL_COND_SHARE = 0.28
ALL_PARAMS_DETAIL_DESC_SHARE = 0.32
ALL_PARAMS_CAPTION_MAX = 120
ALL_PARAMS_CAPTION_B_MAX = 56

# Верхняя лента «как Power Query»: ~25% высоты окна.
ALL_PARAMS_PQ_BAND_RATIO = 0.25
ALL_PARAMS_PQ_BAND_MIN_H = 96
ALL_PARAMS_PQ_BAND_MAX_H = 106
ALL_PARAMS_PQ_TAB_H = 26
ALL_PARAMS_PQ_ACT_H = 26
ALL_PARAMS_PQ_ACT_ROWS = 2
ALL_PARAMS_PQ_ACT_COLS = 6
ALL_PARAMS_PQ_GAP = 4
ALL_PARAMS_PQ_TAB_GAP = 3
ALL_PARAMS_PQ_TAB_W = 250
ALL_PARAMS_PQ_ACT_GAP = 3
ALL_PARAMS_PQ_FONT_H = 10
ALL_PARAMS_PQ_ACTIVE_BG = 0x1565C0
ALL_PARAMS_PQ_ACTIVE_FG = 0xFFFFFF
ALL_PARAMS_PQ_IDLE_BG = 0xE3EAF2
ALL_PARAMS_PQ_IDLE_FG = 0x1A237E
ALL_PARAMS_PQ_DEFAULT_TAB = u"home"
# Имя A для шагов из PQ-меню (если есть в каталоге).
ALL_PARAMS_PQ_STEP_NAME = u"Постобработка_Диапазон"
# Вкладки и действия: (id, подпись) / (fn_key, подпись кнопки).
ALL_PARAMS_PQ_TABS = (
    (u"home", u"Основные операции"),
    (u"transform", u"Преобразование данных"),
    (u"add_col", u"Добавление столбцов"),
    (u"combine", u"Объединить данные"),
    (u"manage", u"Управление и оформление"),
)
ALL_PARAMS_PQ_ACTIONS = {
    u"home": (
        (u"удалить_дубликаты", u"Удалить дубликаты"),
        (u"сортировка", u"Сортировка"),
        (u"удаление_строк", u"Удалить строки"),
        (u"удалить_столбцы", u"Удалить столбцы"),
        (u"переставить_столбцы", u"Переставить столбцы"),
        (u"переименовать_столбцы", u"Переименовать столбцы"),
        (u"заполнение_вниз", u"Заполнить вниз"),
        (u"заполнить_вверх", u"Заполнить вверх"),
        (u"автофильтр", u"Автофильтр"),
        (u"закрепить_заголовок", u"Закрепить заголовок"),
        (u"ширина_столбцов", u"Ширина столбцов"),
        (u"авто_высота", u"Автовысота"),
    ),
    u"transform": (
        (u"замена_значений", u"Замена значений"),
        (u"текстовые_операции", u"Текст"),
        (u"разделить_по_столбцам", u"Разделить столбец"),
        (u"конкатенация_столбцов", u"Объединить столбцы"),
        (u"формат_даты", u"Формат даты"),
        (u"формат_деньги", u"Формат числа"),
        (u"формат_столбцы", u"Формат столбцов"),
        (u"транспонировать_таблицу", u"Транспонировать"),
        (u"развернуть_столбцы", u"Развернуть"),
        (u"убрать_объединение", u"Убрать объединение"),
        (u"добавить_объединение", u"Объединить ячейки"),
        (u"зебра_диапазон", u"Зебра"),
    ),
    u"add_col": (
        (u"условный_столбец", u"Условный столбец"),
        (u"добавить_столбец", u"Добавить столбец"),
        (u"применить_формулу", u"Формула"),
        (u"конкатенация_столбцов", u"Склеить столбцы"),
        (u"количество_значений", u"Количество значений"),
        (u"подсветка_по_порогу", u"Подсветка по порогу"),
        (u"градиент", u"Градиент"),
        (u"копировать_значения", u"Копировать значения"),
    ),
    u"combine": (
        (u"ВПР", u"ВПР"),
        (u"группировать_строки", u"Группировать строки"),
        (u"группировка_по_столбцу", u"Группировка по столбцу"),
        (u"сводная_таблица", u"Сводная таблица"),
        (u"копирование_диапазонов", u"Копировать диапазон"),
        (u"анкета_в_таблицу", u"Анкета → таблица"),
        (u"таблица_в_анкету", u"Таблица → анкета"),
    ),
    u"manage": (
        (u"создать_лист", u"Создать лист"),
        (u"переименовать_лист", u"Переименовать лист"),
        (u"копировать_переместить_лист", u"Копировать лист"),
        (u"удаление_листов", u"Удалить листы"),
        (u"скрытие_листов", u"Скрыть листы"),
        (u"стиль_печати", u"Стиль печати"),
        (u"сетка", u"Сетка"),
        (u"тонкая_сетка", u"Тонкая сетка"),
        (u"толстая_сетка", u"Толстая сетка"),
        (u"шрифт", u"Шрифт"),
        (u"перенос", u"Перенос"),
        (u"левое_выравнивание", u"Выравнивание"),
    ),
}

ALL_PARAMS_TOOLBAR_BTNS = (
    (u"AddBtn", u"Добавить", 90),
    (u"DelBtn", u"Удалить", 84),
    (u"CopyBtn", u"Скопировать", 100),
    (u"UpBtn", u"↑", 36),
    (u"DownBtn", u"↓", 36),
)

# Имя по умолчанию для новой строки (если есть в каталоге и доступно).
ALL_PARAMS_ADD_DEFAULT_NAME = u"Постобработка"

ALL_PARAMS_DISABLED_PREFIX = u"[откл] "
