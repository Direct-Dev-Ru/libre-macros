# -*- coding: utf-8 -*-
"""Константы условия выполнения шага постобработки / финала (колонка D / E)."""
from __future__ import unicode_literals
MACRO_VERSION = "3.10.728"
try:
    unicode
except NameError:
    unicode = str

STEP_CONDITION_FN = u"условие_выполнения"
STEP_CONDITION_VERSION = 1

# Колонки 0-based на листе параметров
STEP_CONDITION_COL_CODEC = 3  # D
STEP_CONDITION_COL_PLUGIN = 4  # E

KIND_LAMBDA = u"lambda"
KIND_VARIABLE = u"variable"
KIND_ENV = u"env"
KIND_OS_USER = u"os_user"

KIND_CHOICES = (
    (KIND_LAMBDA, u"Лямбда"),
    (KIND_VARIABLE, u"Переменная runtime"),
    (KIND_ENV, u"Переменная среды"),
    (KIND_OS_USER, u"Пользователь ОС"),
)

KIND_LABEL = dict(KIND_CHOICES)
KIND_CODE = dict((lab, code) for code, lab in KIND_CHOICES)

COMPARE_OPS = (
    (u"eq", u"равно"),
    (u"ne", u"не равно"),
    (u"contains", u"содержит"),
    (u"not_contains", u"не содержит"),
    (u"starts_with", u"начинается с"),
    (u"ends_with", u"заканчивается на"),
    (u"gt", u"больше"),
    (u"gte", u"больше или равно"),
    (u"lt", u"меньше"),
    (u"lte", u"меньше или равно"),
    (u"in", u"входит в список"),
    (u"not_in", u"не входит в список"),
)

COMPARE_OP_LABEL = dict(COMPARE_OPS)
COMPARE_OP_CODE = dict((lab, code) for code, lab in COMPARE_OPS)

OS_USER_OPS = (
    (u"in", u"входит в список"),
    (u"not_in", u"не входит в список"),
)

RHS_KIND_CHOICES = (
    (u"literal", u"значение"),
    (u"global", u"глобальная переменная"),
    (u"variable", u"переменная карты"),
)

RHS_KIND_LABEL = dict(RHS_KIND_CHOICES)
RHS_KIND_CODE = dict((lab, code) for code, lab in RHS_KIND_CHOICES)

MISSING_CHOICES = (
    (u"false", u"считать ложным"),
    (u"true", u"считать истинным"),
    (u"fail", u"ошибка шага"),
)

MISSING_LABEL = dict(MISSING_CHOICES)
MISSING_CODE = dict((lab, code) for code, lab in MISSING_CHOICES)

COMPARE_AS_CHOICES = (
    (u"text", u"текст"),
    (u"number", u"число"),
    (u"auto", u"авто"),
)

DIALOG_TITLE = u"Условие выполнения шага"
DIALOG_HINT_CODEC = (
    u"Пустое условие = шаг всегда выполняется.\n"
    u"Сейчас редактируется колонка D."
)
DIALOG_HINT_PLUGIN = (
    u"Пустое условие = шаг всегда выполняется.\n"
    u"Для плагина условие хранится в колонке E (D = extra)."
)

# Размеры субвизарда (не тянуть высоту главного визарда).
DIALOG_W = 580
DIALOG_H = 420
DIALOG_MARGIN = 12
DIALOG_HINT_H = 36
DIALOG_ROW_H = 28
DIALOG_CTL_H = 22
DIALOG_LBL_H = 16
DIALOG_LBL_W = 150
DIALOG_BTN_H = 24
DIALOG_BTN_W = 100
DIALOG_BTN_GAP = 10
DIALOG_FOOTER_H = 36
DIALOG_LAMBDA_EDIT_H = 110
DIALOG_OS_USERS_EDIT_H = 72

# Примеры для kind=lambda (вставка в поле expr).
LAMBDA_EXAMPLE_CHOICES = (
    (
        u"Переменная равна",
        u'lambda: str(<<Переменные.Филиал>> or "") == "МСУЗ"',
    ),
    (
        u"Переменная в списке",
        u'lambda: "<<Переменные.Регион>>" in ("Север", "Юг")',
    ),
    (
        u"Переменная непустая",
        u'lambda: str(<<Переменные.Код>> or "").strip() != ""',
    ),
    (
        u"Пользователь ОС",
        u'lambda: os_user().casefold() in ("ivanov", "petrov")',
    ),
    (
        u"Переменная среды",
        u'lambda: env("MERGE_PROFILE").casefold() == "prod"',
    ),
    (
        u"Комбо: код + пользователь",
        u'lambda: str(<<Переменные.Код>> or "").strip() != "" and os_user().casefold() in ("ivanov", "petrov")',
    ),
)

LAMBDA_BUILDER_KIND = u"gate"
