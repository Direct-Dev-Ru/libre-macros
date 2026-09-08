# -*- coding: utf-8 -*-
"""
Общие константы и построители AWT-диалогов для Python-макросов LibreOffice Calc.
"""

MACRO_VERSION = "3.10.706"
DLG_MARGIN = 10
DLG_GAP = 8
DLG_BTN_W = 96
DLG_BTN_H = 22
DLG_INSET = 8
DLG_LABEL_TOP = 16
# Как param_wizard (create_param_wizard_dialog): фиксированный отступ от окна родителя.
DLG_WIZARD_POSITION_X = 80
DLG_WIZARD_POSITION_Y = 60


def dlg_inner_rect(x, y, w, h):
    """Прямоугольник содержимого внутри GroupBox."""
    return (
        int(x) + DLG_INSET,
        int(y) + DLG_LABEL_TOP,
        int(w) - 2 * DLG_INSET,
        int(h) - DLG_LABEL_TOP - DLG_INSET,
    )


def dlg_add_group(dm, name, x, y, w, h, label):
    m = dm.createInstance("com.sun.star.awt.UnoControlGroupBoxModel")
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    m.Label = str(label or "")
    m.Name = str(name)
    dm.insertByName(str(name), m)
    return m


def dlg_add_edit(dm, name, x, y, w, h, multiline=False, readonly=False):
    m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    if multiline:
        m.MultiLine = True
    if readonly:
        m.ReadOnly = True
    try:
        if multiline:
            m.VScroll = True
    except Exception:
        pass
    try:
        m.Border = 1
    except Exception:
        pass
    m.Name = str(name)
    dm.insertByName(str(name), m)
    return m


def dlg_add_framed_edit( dm, frame_name, edit_name, x, y, w, h, frame_label, multiline=True, readonly=True ):
    """Поле ввода внутри GroupBox (обрамление с подписью)."""
    dlg_add_group(dm, frame_name, x, y, w, h, frame_label)
    ix, iy, iw, ih = dlg_inner_rect(x, y, w, h)
    return dlg_add_edit(dm, edit_name, ix, iy, iw, ih, multiline=multiline, readonly=readonly)


def dlg_add_button(dm, name, label, x, y, w=None, h=None):
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w if w is not None else DLG_BTN_W)
    m.Height = int(h if h is not None else DLG_BTN_H)
    m.Label = str(label)
    m.Name = str(name)
    dm.insertByName(str(name), m)
    return m


def dlg_add_progress(dm, name, x, y, w, h, max_value):
    m = dm.createInstance("com.sun.star.awt.UnoControlProgressBarModel")
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    m.ProgressValueMin = 0
    mv = int(max_value) if int(max_value) > 0 else 1000
    m.ProgressValueMax = mv
    m.Name = str(name)
    dm.insertByName(str(name), m)
    return m


def dlg_center_model(dm, toolkit, parent_win=None):
    """
    Позиционирование диалога как в param_wizard (PositionX/Y = 80/60).

    Координаты задаются относительно окна parent при createPeer(parent).
    getWorkArea() не используем — на двух мониторах даёт «щель» между экранами.
    """
    if dm is None:
        return
    dm.PositionX = DLG_WIZARD_POSITION_X
    dm.PositionY = DLG_WIZARD_POSITION_Y


def dlg_center_on_desktop(dm, toolkit):
    """Центр рабочей области (удобно, если parent неизвестен)."""
    dlg_center_model(dm, toolkit, None)


def dlg_place_buttons_right(dm, dialog_width, y, buttons, margin=None, gap=None):
    """
    Кнопки в один ряд у правого края.

    buttons — [(name, label), …] слева направо; размещаются справа налево.
    Возвращает y + высота кнопки.
    """
    m = margin if margin is not None else DLG_MARGIN
    g = gap if gap is not None else DLG_GAP
    x = int(dialog_width) - m - DLG_BTN_W
    for name, label in reversed(buttons):
        dlg_add_button(dm, name, label, x, y)
        x = x - DLG_BTN_W - g
    return int(y) + DLG_BTN_H
