# -*- coding: utf-8 -*-
MACRO_VERSION = "3.10.499"
"""
Проверка bundled odfpy из macro-lib/pythonpath/odf_bundled.py.

Создаёт временный .ods, читает его обратно и выводит результат на активный лист.
"""
import os
import tempfile

import uno
import odf_bundled  # noqa: F401 — регистрирует import hook
ODF_VERSION = "1.4.1"
from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.table import Table, TableCell, TableRow
from odf.text import P


def _log(msg):
    try:
        print("[test_odf] %s" % msg)
    except Exception:
        pass


def _show_info(title, text):
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    parent = getattr(
        getattr(getattr(doc, "CurrentController", None), "Frame", None),
        "ContainerWindow",
        None,
    )
    if parent is None:
        parent = toolkit.getDesktopWindow()
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    from com.sun.star.awt.MessageBoxType import INFOBOX

    box = toolkit.createMessageBox(parent, INFOBOX, BUTTONS_OK, title, text)
    box.execute()


def _cell_text(cell):
    parts = cell.getElementsByType(P)
    if not parts:
        return ""
    node = parts[0].firstChild
    return "" if node is None else str(node.data)


def test_odf_bundle():
    _log("старт, odfpy %s" % ODF_VERSION)

    fd, ods_path = tempfile.mkstemp(prefix="libre_macros_odf_", suffix=".ods")
    os.close(fd)

    try:
        doc = OpenDocumentSpreadsheet()
        table = Table(name="Проверка")
        doc.spreadsheet.addElement(table)

        header = TableRow()
        header.addElement(_make_cell("Ключ"))
        header.addElement(_make_cell("Значение"))
        table.addElement(header)

        row = TableRow()
        row.addElement(_make_cell("alpha"))
        row.addElement(_make_cell("123"))
        table.addElement(row)

        doc.save(ods_path)
        _log("записано: %s" % ods_path)

        doc2 = load(ods_path)
        sheet = doc2.spreadsheet.getElementsByType(Table)[0]
        rows = sheet.getElementsByType(TableRow)
        values = []
        for row_item in rows:
            values.append(tuple(_cell_text(cell) for cell in row_item.getElementsByType(TableCell)))
        _log("прочитано строк: %d" % len(values))

        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        lo_doc = desktop.getCurrentComponent()
        lo_sheet = lo_doc.getCurrentController().getActiveSheet()

        out = (
            ("odfpy", ODF_VERSION),
            ("Файл", ods_path),
            ("Строк прочитано", len(values)),
            ("A2", values[1][0] if len(values) > 1 else None),
            ("B2", values[1][1] if len(values) > 1 else None),
        )
        lo_sheet.getCellRangeByPosition(0, 0, 1, len(out) - 1).setDataArray(out)

        _show_info(
            "odf bundled",
            "odfpy %s работает.\n\n"
            "Временный файл:\n%s\n\n"
            "На лист записана сводка (колонки A:B, с первой строки)."
            % (ODF_VERSION, ods_path),
        )
    finally:
        try:
            os.unlink(ods_path)
        except Exception:
            pass

    _log("конец")


def _make_cell(text):
    cell = TableCell()
    cell.addElement(P(text=str(text)))
    return cell


g_exportedScripts = (test_odf_bundle,)
