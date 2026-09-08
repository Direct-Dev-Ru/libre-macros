# -*- coding: utf-8 -*-
"""Константы AD/LDAP: AO 2026 вырезает модульные присваивания в Scripts/*.py."""

MACRO_VERSION = "3.10.700"
DEFAULT_PORT = 389
DEFAULT_SSL_PORT = 636
DEFAULT_AUTH = "SIMPLE"
DEFAULT_PAGE_SIZE = 500

CONFIG_SHEET = u"AD"
CONFIG_ROW = 1
CONFIG_COL_SERVER = 0
CONFIG_COL_USER = 1
# C2 не используется — пароль только в диалоге (libre_macros_ad_dialog_lib).
CONFIG_COL_PASSWORD = 2
CONFIG_COL_AUTH = 3
CONFIG_COL_SSL = 4
CONFIG_COL_PORT = 5
CONFIG_COL_OU = 6

OUTPUT_HEADER_ROW = 3
OUTPUT_DATA_START_ROW = 4

OUTPUT_HEADERS = (
    u"Логин",
    u"ФИО",
    u"Email",
    u"Должность",
    u"Подразделение",
    u"Телефон",
    u"UPN",
    u"Группы безопасности",
)

USER_LDAP_ATTRIBUTES = (
    "sAMAccountName",
    "displayName",
    "mail",
    "title",
    "department",
    "telephoneNumber",
    "userPrincipalName",
    "memberOf",
)

USER_SEARCH_FILTER = u"(&(objectClass=user)(objectCategory=person))"

AUTH_CHOICES = (u"SIMPLE", u"NTLM")
