# -*- coding: utf-8 -*-
"""Константы полноэкранного визарда «Все параметры» (пайплайн листа)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.777"
ALL_PARAMS_BTN_LABEL = u"Все параметры"
ALL_PARAMS_WIZARD_TITLE = u"Все параметры — пайплайн обработки"
ALL_PARAMS_WIZARD_HINT = (
    u"Слева — все шаги листа параметров (порядок сверху вниз). "
    u"Сверху списка: Удалить / Скопировать / ↑ / ↓. "
    u"Справа — поля B/C/D выбранной строки; "
    u"кнопки «Параметры C…» и «Параметры D…» открывают визарды колонок C и D."
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
ALL_PARAMS_DETAIL_BTN_W = 118
ALL_PARAMS_LIST_TOOLBAR_GAP = 4
ALL_PARAMS_DETAIL_ROW_H = 22
ALL_PARAMS_DETAIL_LBL_H = 14
ALL_PARAMS_DETAIL_GAP = 4
ALL_PARAMS_DETAIL_SECTION_GAP = 6
# Минимальные высоты полей деталей (итог распределяется по доступной высоте).
ALL_PARAMS_DETAIL_VALUE_MIN_H = 40
ALL_PARAMS_DETAIL_EXTRA_MIN_H = 56
ALL_PARAMS_DETAIL_COND_MIN_H = 36
ALL_PARAMS_DETAIL_DESC_MIN_H = 36
# Доли остатка высоты: B / C / D / описание.
ALL_PARAMS_DETAIL_VALUE_SHARE = 0.28
ALL_PARAMS_DETAIL_EXTRA_SHARE = 0.40
ALL_PARAMS_DETAIL_COND_SHARE = 0.16
ALL_PARAMS_DETAIL_DESC_SHARE = 0.16
ALL_PARAMS_CAPTION_MAX = 48
ALL_PARAMS_CAPTION_B_MAX = 24

ALL_PARAMS_TOOLBAR_BTNS = (
    (u"DelBtn", u"Удалить", 84),
    (u"CopyBtn", u"Скопировать", 100),
    (u"UpBtn", u"↑", 36),
    (u"DownBtn", u"↓", 36),
)

ALL_PARAMS_DISABLED_PREFIX = u"[откл] "
