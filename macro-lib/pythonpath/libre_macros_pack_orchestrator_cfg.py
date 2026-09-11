# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.726"
"""Константы диалога оркестратора collect_pack."""

try:
    unicode
except NameError:
    unicode = str

PACK_ORCH_JSON_FILE = u"collect_pack_orchestrator.json"
PACK_ORCH_JSON_VER = 1
PACK_ORCH_PRESETS_SUBDIR = u"collect_workbooks"

PACK_ORCH_TITLE = u"Оркестратор сбора (collect_pack)"
PACK_ORCH_HINT_MULTI = (
    u"Отметьте листы параметров и задайте порядок запуска.\n"
    u"Зелёный — включён, серый — выключен; клик по галочке; "
    u"клик по подписи — выбор; Ctrl — мультивыбор; кнопки справа."
)
PACK_ORCH_HINT_SINGLE = (
    u"Будет запущен сбор с параметрами листа «%s»%s."
)
PACK_ORCH_COMMENT_PREVIEW_LEN = 50
PACK_ORCH_AUTO_CONTINUE_LABEL = (
    u"Автопродолжение шагов (без подтверждения «Запуск» между шагами)"
)
PACK_ORCH_BTN_RUN = u"Запустить"
PACK_ORCH_BTN_CANCEL = u"Отмена"
PACK_ORCH_BTN_SAVE = u"Сохранить настройки"
PACK_ORCH_BTN_HELP = u"Справка"
PACK_ORCH_BTN_TOGGLE = u"Вкл/Выкл"
PACK_ORCH_BTN_UP = u"Выше"
PACK_ORCH_BTN_DOWN = u"Ниже"
PACK_ORCH_BTN_ENABLE_ALL = u"Установить все"
PACK_ORCH_BTN_DISABLE_ALL = u"Снять все"
# Совместимость со старыми метками ListBox (больше не рисуются в UI).
PACK_ORCH_MARK_ON = u"[v] "
PACK_ORCH_MARK_OFF = u"[x] "
# Цвета подписей задач (UNO 0x00RRGGBB): включена / выключена / фон выбора.
PACK_ORCH_COLOR_ON = 0x1B7A3D
PACK_ORCH_COLOR_OFF = 0x808080
PACK_ORCH_COLOR_SEL_BG = 0xD0E4F7
PACK_ORCH_COLOR_ROW_BG = 0xF4F4F4
# высота строки списка — см. PACK_ORCH_ROW_H ниже (с учётом шрифта чекбокса)

PACK_ORCH_ERR_NO_SHEETS = (
    u"В активной книге нет листа параметров "
    u"(collect_params… или Параметры_Объединения…).\n"
    u"Переключитесь на нужную книгу и запустите макрос снова."
)
PACK_ORCH_ERR_EMPTY_PLAN = u"Выберите хотя бы один лист параметров."
PACK_ORCH_ERR_UNSAVED = (
    u"Сначала сохраните книгу на диск — настройки привязываются к пути файла."
)
PACK_ORCH_ERR_SAVE_FAILED = u"Не удалось сохранить настройки:\n%s"
PACK_ORCH_SAVE_OK = u"Настройки оркестратора сохранены."
PACK_ORCH_ERR_PREVALIDATE = u"Ошибка параметров на листе «%s»:\n%s"

PACK_ORCH_HELP_TITLE = u"Справка — оркестратор сбора"
PACK_ORCH_HELP_TEXT = (
    u"Оркестратор collect_pack последовательно запускает сбор по выбранным\n"
    u"листам Параметры_Объединения*.\n"
    u"\n"
    u"• Сохранить — запомнить план и флаги для этой книги.\n"
    u"• Запустить — prevalidate параметров, затем сбор по плану.\n"
    u"• Ручной ввод / <<Переменные.…>> в «Файлы-Источники»:\n"
    u"  проверка наличия файлов выполняется при реальном сборе\n"
    u"  (после диалога ввода), а не на этапе prevalidate.\n"
    u"• Несколько листов — журнал Сбор_книг_лог_1, _2, …\n"
    u"• Сброс DatabaseRange — в конце, если галочка и был успех.\n"
    u"\n"
    u"Подробнее: docs/18_ORCHESTRATOR.md"
)

PACK_DLG_W = 480
PACK_DLG_MARGIN = 10
PACK_DLG_GAP = 8
PACK_DLG_BTN_H = 22
PACK_DLG_BTN_W = 100
# Нижний ряд: [Запустить] [gap] [Отмена] … [Сохранить] [gap] [Справка]
PACK_DLG_SAVE_BTN_W = 122
PACK_DLG_HELP_BTN_W = 68
PACK_DLG_FOOTER_GAP_SMALL = 8
PACK_DLG_LIST_H = 160
PACK_DLG_HINT_H = 40
PACK_DLG_CHECK_H = 18
# Чекбоксы списка листов: крупнее; при выделении — жирный.
PACK_ORCH_CHK_FONT_HEIGHT = 11
PACK_ORCH_ROW_H = 18
# Вертикальный ряд справа: +1 pt к типичным ~10 pt подписей кнопок.
PACK_ORCH_SIDE_BTN_FONT_HEIGHT = 11