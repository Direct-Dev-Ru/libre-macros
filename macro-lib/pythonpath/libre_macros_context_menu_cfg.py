# -*- coding: utf-8 -*-
"""Константы context_menu (AlterOffice 2026: модульный код .py-скрипта вырезается)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.715"
MENU_NS = "http://openoffice.org/2001/menu"
REL_POPUP = ("config", "soffice.cfg", "modules", "scalc", "popupmenu", "cell.xml")
REL_POPUP_AO = ("config", "aoffice.cfg", "modules", "acell", "popupmenu", "cell.xml")
REL_MENUBAR = ("config", "soffice.cfg", "modules", "scalc", "menubar", "menubar.xml")
REL_MENUBAR_AO = ("config", "aoffice.cfg", "modules", "acell", "menubar", "menubar.xml")
DEFAULT_CELL_POPUP_XML = (
    u'<?xml version="1.0" encoding="UTF-8"?>\n'
    u'<menu:menupopup xmlns:menu="http://openoffice.org/2001/menu">\n'
    u'  <menu:menuseparator />\n'
    u'  <menu:menuitem menu:id=".uno:Cut" />\n'
    u'  <menu:menuitem menu:id=".uno:Copy" />\n'
    u'  <menu:menuitem menu:id=".uno:Paste" />\n'
    u'  <menu:menu menu:id=".uno:PasteSpecialMenu">\n'
    u'    <menu:menupopup>\n'
    u'      <menu:menuitem menu:id=".uno:PasteUnformatted" />\n'
    u'      <menu:menuseparator />\n'
    u'      <menu:menuitem menu:id=".uno:PasteOnlyText" />\n'
    u'      <menu:menuitem menu:id=".uno:PasteOnlyValue" />\n'
    u'      <menu:menuitem menu:id=".uno:PasteOnlyFormula" />\n'
    u'      <menu:menuseparator />\n'
    u'      <menu:menuitem menu:id=".uno:PasteTransposed" />\n'
    u'      <menu:menuseparator />\n'
    u'      <menu:menuitem menu:id=".uno:PasteSpecial" />\n'
    u'    </menu:menupopup>\n'
    u'  </menu:menu>\n'
    u'  <menu:menuseparator />\n'
    u'  <menu:menuitem menu:id=".uno:DataSelect" />\n'
    u'  <menu:menuitem menu:id=".uno:CurrentValidation" />\n'
    u'  <menu:menuitem menu:id=".uno:DefineCurrentName" />\n'
    u'  <menu:menuseparator />\n'
    u'  <menu:menuitem menu:id=".uno:InsertCell" />\n'
    u'  <menu:menuitem menu:id=".uno:DeleteCell" />\n'
    u'  <menu:menuitem menu:id=".uno:Delete" />\n'
    u'  <menu:menuitem menu:id=".uno:MergeCells" />\n'
    u'  <menu:menuitem menu:id=".uno:SplitCell" />\n'
    u'  <menu:menuseparator />\n'
    u'  <menu:menuitem menu:id=".uno:FormatPaintbrush" />\n'
    u'  <menu:menuitem menu:id=".uno:ResetAttributes" />\n'
    u'  <menu:menu menu:id=".uno:FormatStylesMenu">\n'
    u'    <menu:menupopup>\n'
    u'      <menu:menuitem menu:id=".uno:EditStyle" />\n'
    u'      <menu:menuseparator />\n'
    u'      <menu:menuitem menu:id=".uno:DefaultCellStylesmenu" menu:style="radio" />\n'
    u'      <menu:menuitem menu:id=".uno:Accent1CellStyles" menu:style="radio" />\n'
    u'      <menu:menuitem menu:id=".uno:Accent2CellStyles" menu:style="radio" />\n'
    u'      <menu:menuitem menu:id=".uno:Accent3CellStyles" menu:style="radio" />\n'
    u'      <menu:menuseparator />\n'
    u'      <menu:menuitem menu:id=".uno:BadCellStyles" menu:style="radio" />\n'
    u'      <menu:menuitem menu:id=".uno:ErrorCellStyles" menu:style="radio" />\n'
    u'      <menu:menuitem menu:id=".uno:GoodCellStyles" menu:style="radio" />\n'
    u'      <menu:menuitem menu:id=".uno:NeutralCellStyles" menu:style="radio" />\n'
    u'      <menu:menuitem menu:id=".uno:WarningCellStyles" menu:style="radio" />\n'
    u'      <menu:menuseparator />\n'
    u'      <menu:menuitem menu:id=".uno:FootnoteCellStyles" menu:style="radio" />\n'
    u'      <menu:menuitem menu:id=".uno:NoteCellStyles" menu:style="radio" />\n'
    u'    </menu:menupopup>\n'
    u'  </menu:menu>\n'
    u'  <menu:menuseparator />\n'
    u'  <menu:menuitem menu:id=".uno:InsertAnnotation" />\n'
    u'  <menu:menuitem menu:id=".uno:EditAnnotation" />\n'
    u'  <menu:menuitem menu:id=".uno:DeleteNote" />\n'
    u'  <menu:menuitem menu:id=".uno:ShowNote" />\n'
    u'  <menu:menuitem menu:id=".uno:HideNote" />\n'
    u'  <menu:menuseparator />\n'
    u'  <menu:menu menu:id=".uno:FormatSparklineMenu">\n'
    u'    <menu:menupopup />\n'
    u'  </menu:menu>\n'
    u'  <menu:menuseparator />\n'
    u'  <menu:menuitem menu:id=".uno:CurrentConditionalFormatDialog" />\n'
    u'  <menu:menuitem menu:id=".uno:CurrentConditionalFormatManagerDialog" />\n'
    u'  <menu:menuitem menu:id=".uno:FormatCellDialog" />\n'
    u'</menu:menupopup>'
)
TOOLS_MENU_ID = ".uno:ToolsMenu"
TARGET_CELL = "cell"
TARGET_TOOLS = "tools"
SKIP_SCAN = frozenset(
    [
        "context_menu.py",
        "macro_ui.py",
        "generate_no_docstrings.py",
        "generate_py2_version.py",
        "collect_workbooks_no_docstrings.py",
        "__init__.py",
    ]
)
DLG_W = 460
DLG_MARGIN = 10
DLG_GAP = 8
DLG_BTN_W = 96
DLG_BTN_H = 22
DLG_ROW_H = 22

# Кэш XML-тегов (мутабельно)
tag_menu = None
tag_menuitem = None
tag_menupopup = None
tag_menuseparator = None
tag_id = None
tag_label = None
