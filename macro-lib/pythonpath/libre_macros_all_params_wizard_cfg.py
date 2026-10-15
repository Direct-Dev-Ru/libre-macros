# -*- coding: utf-8 -*-
"""Константы полноэкранного визарда «Все параметры» (пайплайн листа)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.784"
ALL_PARAMS_BTN_LABEL = u"Все параметры"
ALL_PARAMS_WIZARD_TITLE = u"Все параметры — пайплайн обработки"
ALL_PARAMS_WIZARD_HINT = (
    u"Слева — все шаги листа параметров (порядок сверху вниз). "
    u"Сверху списка: Добавить / Удалить / Скопировать / ↑ / ↓. "
    u"«Добавить» — новая строка сразу после выделенной. "
    u"Справа — поля A/B/C/D; кнопки «Параметры C…» / «Параметры D…» — между C и D."
)

ALL_PARAMS_BTN_C_LABEL = u"Параметры C…"
ALL_PARAMS_BTN_D_LABEL = u"Параметры D…"

ALL_PARAMS_TOOLBAR_BTN_H = 28
ALL_PARAMS_HINT_H = 28
ALL_PARAMS_HINT_TOP_GAP = 4
ALL_PARAMS_TOOLBAR_GAP = 6
ALL_PARAMS_SPLIT_GAP = 8
ALL_PARAMS_LIST_LBL_H = 18
ALL_PARAMS_LIST_WIDTH_RATIO = 0.40
ALL_PARAMS_LIST_MIN_W = 220
ALL_PARAMS_LIST_MIN_H = 80
ALL_PARAMS_LIST_BOTTOM_GAP = 6
ALL_PARAMS_STATUS_H = 18
ALL_PARAMS_FOOTER_BTN_W = 110
ALL_PARAMS_DETAIL_BTN_W = 130
ALL_PARAMS_LIST_TOOLBAR_GAP = 4
ALL_PARAMS_DETAIL_ROW_H = 20
ALL_PARAMS_DETAIL_COMBO_H = 18
ALL_PARAMS_DETAIL_LBL_H = 14
ALL_PARAMS_DETAIL_GAP = 4
ALL_PARAMS_DETAIL_SECTION_GAP = 6
# Кнопки C/D между полями — выше + жирный шрифт.
ALL_PARAMS_DETAIL_BTN_ROW_H = 30
ALL_PARAMS_DETAIL_BTN_FONT_H = 11
# Нижняя «Применить» — крупная, прижата к низу панели.
ALL_PARAMS_APPLY_BTN_H = 36
ALL_PARAMS_APPLY_BTN_W = 200
# Шрифт combo A (pt).
ALL_PARAMS_NAME_FONT_H = 12
ALL_PARAMS_VALUE_FONT_H = 10
# C на ~20% ниже прежнего минимума/доли; место уходит вниз (D/описание/Apply).
ALL_PARAMS_DETAIL_EXTRA_MIN_H = 45
ALL_PARAMS_DETAIL_COND_MIN_H = 44
ALL_PARAMS_DETAIL_DESC_MIN_H = 56
ALL_PARAMS_DETAIL_EXTRA_SHARE = 0.38
ALL_PARAMS_DETAIL_COND_SHARE = 0.28
ALL_PARAMS_DETAIL_DESC_SHARE = 0.34
ALL_PARAMS_CAPTION_MAX = 120
ALL_PARAMS_CAPTION_B_MAX = 56

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
