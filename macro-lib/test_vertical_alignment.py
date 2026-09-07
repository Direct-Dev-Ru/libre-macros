# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.499"
"""
Тестовый макрос: вертикальное выравнивание по центру на значащем диапазоне листа.

Запуск из LibreOffice (Scripts/python):
  test_vertical_alignment.set_vertical_center_test
"""

import uno

import libre_macros_lib as lm

try:
    unicode
except NameError:
    unicode = str


def _pv(name, value):
    """PropertyValue (UNO)."""
    p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    p.Name = name
    p.Value = value
    return p


def _get_doc():
    """Текущий документ."""
    try:
        doc = XSCRIPTCONTEXT.getDocument()
        if doc is not None:
            return doc
    except Exception:
        pass
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        return desktop.getCurrentComponent()
    except Exception:
        return None


def set_vertical_center_test(use_dispatch=True):
    """
    Тест: вертикальный центр на значащем диапазоне активного листа.

    use_dispatch: дополнительно выполнить .uno:VerticalAlignment (как Macro Recorder)
    """
    doc = _get_doc()
    if doc is None:
        print("Нет активного документа")
        return

    controller = None
    try:
        controller = doc.getCurrentController()
    except Exception:
        controller = None
    if controller is None:
        print("Нет контроллера документа")
        return

    try:
        sheet = controller.ActiveSheet
    except Exception:
        sheet = None
    if sheet is None:
        print("Не удалось получить активный лист")
        return

    sheet_name = ""
    try:
        sheet_name = unicode(getattr(sheet, "Name", "") or "")
    except Exception:
        sheet_name = ""

    try:
        sc, sr, ec, er = lm._lm_vlookup_sheet_used_area(sheet)
    except Exception as err:
        print("Лист %s: used_area ошибка: %s" % (sheet_name, err))
        return
    if er < sr or ec < sc:
        print("Лист %s: пустой значащий диапазон" % sheet_name)
        return

    print("Лист %s: used_area (%s,%s)-(%s,%s)" % (sheet_name, sc, sr, ec, er))

    try:
        cell_range = sheet.getCellRangeByPosition(int(sc), int(sr), int(ec), int(er))
    except Exception as err:
        print("Лист %s: getCellRangeByPosition ошибка: %s" % (sheet_name, err))
        return

    try:
        cell_range.VertJustify = lm.VERT_CENTER
        print("Лист %s: cell_range.VertJustify = CENTER (ok)" % sheet_name)
    except Exception as err:
        print("Лист %s: VertJustify assign error: %s" % (sheet_name, err))

    try:
        c0 = cell_range.getCellByPosition(0, 0)
        v = getattr(c0, "VertJustify", None)
        print("Лист %s: readback top-left VertJustify=%s" % (sheet_name, v))
    except Exception as err:
        print("Лист %s: readback error: %s" % (sheet_name, err))

    if use_dispatch:
        if not lm._lm_pp_select_cell_range(doc, sheet, cell_range):
            print("Лист %s: select range failed" % sheet_name)
            return
        ok = lm._lm_pp_execute_dispatch(
            doc, ".uno:VerticalAlignment", (_pv("VerticalAlignment", 2),)
        )
        print(
            "Лист %s: dispatch .uno:VerticalAlignment => %s"
            % (sheet_name, "ok" if ok else "fail")
        )

        try:
            c0 = cell_range.getCellByPosition(0, 0)
            v = getattr(c0, "VertJustify", None)
            print("Лист %s: readback after dispatch VertJustify=%s" % (sheet_name, v))
        except Exception as err:
            print("Лист %s: readback after dispatch error: %s" % (sheet_name, err))


g_exportedScripts = (set_vertical_center_test,)
