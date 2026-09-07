# -*- coding: utf-8 -*-
"""
Проверка bundled segno: текст из A2 → QR-картинка в B2 на активном листе.

Запуск из меню макросов LibreOffice / AlterOffice: test_segno.test_segno_qr.
Перед запуском введите строку в ячейку A2 активного листа.
"""
from __future__ import annotations

import os
import tempfile

import uno

import segno_bundled  # noqa: F401 — регистрирует import hook
import segno


def _log(msg):
    try:
        print("[test_segno] %s" % msg)
    except Exception:
        pass


def _show_message(title, text, is_error=False):
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
    from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX

    box_type = ERRORBOX if is_error else INFOBOX
    box = toolkit.createMessageBox(parent, box_type, BUTTONS_OK, title, text)
    box.execute()


def _cell_text(cell):
    try:
        t = cell.getString()
        if t is not None and str(t).strip() != "":
            return str(t)
    except Exception:
        pass
    try:
        v = cell.getValue()
        if v is not None:
            return str(v)
    except Exception:
        pass
    return ""


def _cell_rect(sheet, col, row):
    """
    Прямоугольник ячейки в 1/100 мм: (x, y, width, height).
    У XCell нет Position/Size — берём у диапазона или суммируем колонки/строки.
    """
    from com.sun.star.awt import Point, Size

    rng = sheet.getCellRangeByPosition(col, row, col, row)
    pos = None
    size = None
    try:
        pos = rng.Position
    except Exception:
        pass
    if pos is None:
        try:
            pos = rng.getPropertyValue("Position")
        except Exception:
            pass
    try:
        size = rng.Size
    except Exception:
        pass
    if size is None:
        try:
            size = rng.getPropertyValue("Size")
        except Exception:
            pass

    cols = sheet.getColumns()
    rows = sheet.getRows()
    if pos is None:
        x = 0
        y = 0
        c = 0
        while c < col:
            x = x + int(cols.getByIndex(c).Width)
            c = c + 1
        r = 0
        while r < row:
            y = y + int(rows.getByIndex(r).Height)
            r = r + 1
        pos = Point(x, y)
    if size is None:
        size = Size(int(cols.getByIndex(col).Width), int(rows.getByIndex(row).Height))
    return (int(pos.X), int(pos.Y), int(size.Width), int(size.Height))


def _fit_centered_in_cell(cell_x, cell_y, cell_w, cell_h, margin_mm=1.5, aspect=1.0):
    """
    Размер и позиция фигуры внутри ячейки с отступом, по центру.
    aspect = width/height исходника (для QR = 1).
    Возвращает (x, y, w, h) в 1/100 мм.
    """
    margin = max(0, int(round(margin_mm * 100)))
    inner_w = max(100, cell_w - 2 * margin)
    inner_h = max(100, cell_h - 2 * margin)
    if aspect <= 0:
        aspect = 1.0
    # вписать прямоугольник aspect в inner_*
    if inner_w / float(aspect) <= inner_h:
        w = inner_w
        h = max(100, int(round(w / float(aspect))))
    else:
        h = inner_h
        w = max(100, int(round(h * float(aspect))))
    x = cell_x + (cell_w - w) // 2
    y = cell_y + (cell_h - h) // 2
    return (x, y, w, h)


def _is_graphic_shape(shape):
    """True для вставленных картинок (GraphicObjectShape)."""
    try:
        st = str(shape.getShapeType() or "")
        if "GraphicObject" in st:
            return True
    except Exception:
        pass
    try:
        if shape.supportsService("com.sun.star.drawing.GraphicObjectShape"):
            return True
    except Exception:
        pass
    return False


def _clear_document_pictures(doc):
    """Удалить все картинки со всех листов книги. Возвращает число удалённых."""
    if doc is None or not hasattr(doc, "getSheets"):
        return 0
    removed = 0
    sheets = doc.getSheets()
    si = 0
    while si < sheets.getCount():
        sheet = sheets.getByIndex(si)
        try:
            draw_page = sheet.getDrawPage()
        except Exception:
            si = si + 1
            continue
        j = draw_page.getCount() - 1
        while j >= 0:
            try:
                shape = draw_page.getByIndex(j)
                if _is_graphic_shape(shape):
                    draw_page.remove(shape)
                    removed = removed + 1
            except Exception:
                pass
            j = j - 1
        si = si + 1
    return removed


def _insert_png_at_cell(doc, sheet, png_path, col, row, margin_mm=1.5):
    """Вставить PNG в ячейку: autofit с отступом, по центру по H и V."""
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    gp = smgr.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", ctx)
    from com.sun.star.beans import PropertyValue
    from com.sun.star.awt import Point, Size

    url = uno.systemPathToFileUrl(os.path.abspath(png_path))
    prop = PropertyValue()
    prop.Name = "URL"
    prop.Value = url
    graphic = gp.queryGraphic((prop,))
    if graphic is None:
        raise RuntimeError("GraphicProvider не загрузил файл: %s" % png_path)

    aspect = 1.0
    try:
        gsz = graphic.Size100thMM
        if gsz is not None and int(gsz.Height) > 0:
            aspect = float(gsz.Width) / float(gsz.Height)
    except Exception:
        try:
            gsz = graphic.SizePixel
            if gsz is not None and int(gsz.Height) > 0:
                aspect = float(gsz.Width) / float(gsz.Height)
        except Exception:
            pass

    cx, cy, cw, ch = _cell_rect(sheet, col, row)
    x, y, w, h = _fit_centered_in_cell(cx, cy, cw, ch, margin_mm=margin_mm, aspect=aspect)

    shape = doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
    shape.Graphic = graphic
    shape.setPosition(Point(x, y))
    shape.setSize(Size(w, h))
    try:
        shape.Anchor = sheet.getCellByPosition(col, row)
    except Exception:
        pass
    sheet.getDrawPage().add(shape)
    return shape


def test_segno_qr(*args):
    _unused = args
    _log("старт, segno %s" % getattr(segno, "__version__", "?"))

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    if doc is None or not hasattr(doc, "getSheets"):
        _show_message("segno", "Откройте книгу Calc.", is_error=True)
        return

    sheet = doc.getCurrentController().getActiveSheet()
    src = sheet.getCellByPosition(0, 1)  # A2
    dst = sheet.getCellByPosition(1, 1)  # B2
    text = _cell_text(src).strip()
    if text == "":
        _show_message(
            "segno",
            "Ячейка A2 пуста.\nВведите текст и запустите макрос снова.",
            is_error=True,
        )
        return

    n_cleared = _clear_document_pictures(doc)
    _log("удалено картинок: %d" % n_cleared)

    fd, png_path = tempfile.mkstemp(prefix="libre_macros_segno_", suffix=".png")
    os.close(fd)
    try:
        qr = segno.make(text, error="m")
        qr.save(png_path, scale=8, border=2)
        _log("QR сохранён: %s (%d байт)" % (png_path, os.path.getsize(png_path)))
        _insert_png_at_cell(doc, sheet, png_path, 1, 1, margin_mm=1.5)
        dst.setString("")
        note = ""
        if n_cleared:
            note = "\nУдалено старых картинок: %d." % n_cleared
        _show_message(
            "segno bundled",
            "segno %s: QR из A2 вставлен в B2 (по центру, autofit).%s\n\nТекст:\n%s"
            % (getattr(segno, "__version__", "?"), note, text[:200]),
        )
    except Exception as err:
        _log("ошибка: %s" % err)
        _show_message("segno", "Ошибка: %s" % err, is_error=True)
    finally:
        try:
            os.unlink(png_path)
        except Exception:
            pass

    _log("конец")


g_exportedScripts = (test_segno_qr,)
