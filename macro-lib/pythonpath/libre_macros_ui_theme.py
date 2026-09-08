# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.711"
"""
Общая soft-gray тема диалогов (визард параметров, collect_workbooks, ВПР и т.п.).

Системный TitleBar (GTK/WM) через UNO обычно не красится — рисуем SoftGrayTitleBar.
"""

try:
    unicode
except NameError:
    unicode = str

# UNO 0x00RRGGBB
SOFT_GRAY_FORM_BG = 0xE6E6E6
SOFT_GRAY_TITLE_BG = 0x4A4A4A
SOFT_GRAY_TITLE_FG = 0xF5F5F5
SOFT_GRAY_LABEL_FG = 0x2A2A2A
SOFT_GRAY_FIELD_BG = 0xF4F4F4
SOFT_GRAY_BTN_BG = 0xD4D4D4
SOFT_GRAY_TITLE_H = 28
SOFT_GRAY_BAR_NAME = "SoftGrayTitleBar"
# Алая полоса-титл (предупреждения безопасности и т.п.)
WARNING_TITLE_BG = 0xB71C1C
WARNING_TITLE_FG = 0xFFFFFF
# Оранжевая полоса-титл (предупреждение о плагинах)
ORANGE_TITLE_BG = 0xEF6C00
ORANGE_TITLE_FG = 0xFFFFFF
# Жёлтая полоса-титл (ручной ввод в карту переменных)
MANUAL_INPUT_TITLE_BG = 0xF9A825
MANUAL_INPUT_TITLE_FG = 0x212121
# Тёмно-синяя полоса-титл (макросы задач / todo_task)
DARK_BLUE_TITLE_BG = 0x0D47A1
DARK_BLUE_TITLE_FG = 0xFFFFFF
# Размеры футера микродиалогов (как в libre_macros_param_wizard_cfg)
INNER_BTN_W = 80
INNER_BTN_H = 16
INNER_BTN_GAP = 10
INNER_CHK_H = 14
# ListBox: типичный диалог LO ~10 pt + 2 pt для меток [v]/[x] и длинных подписей.
LISTBOX_FONT_HEIGHT = 12
# Типичный Edit в диалоге LO ~10 pt; +1 pt для поля ввода / справки.
DIALOG_EDIT_FONT_HEIGHT = 10
MANUAL_INPUT_VALUE_FONT_HEIGHT = DIALOG_EDIT_FONT_HEIGHT + 1
HELP_TEXT_FONT_HEIGHT = DIALOG_EDIT_FONT_HEIGHT + 1
# Предупреждение о плагинах: крупнее основной текст.
PLUGIN_WARN_TEXT_FONT_HEIGHT = DIALOG_EDIT_FONT_HEIGHT + 2
# Компактные кнопки «+» / «×» / «××» в визардах (columns_pick, ВПР, …).
TOKEN_BTN_FONT_HEIGHT = 16
TOKEN_BTN_LABELS = (u"+", u"×", u"××", u"x", u"xx", u"X", u"XX")


def apply_listbox_font(model, height=None):
    """Увеличить шрифт ListBox (UNO FontDescriptor.Height, pt)."""
    if model is None:
        return
    h = int(height if height is not None else LISTBOX_FONT_HEIGHT)
    try:
        fd = model.FontDescriptor
        fd.Height = h
        model.FontDescriptor = fd
        return
    except Exception:
        pass
    try:
        model.FontHeight = h
    except Exception:
        pass


def apply_edit_font(model, height=None):
    """Шрифт Edit/MultiLineEdit (FontDescriptor.Height, pt)."""
    if model is None:
        return
    h = int(height if height is not None else DIALOG_EDIT_FONT_HEIGHT)
    try:
        fd = model.FontDescriptor
        fd.Height = h
        model.FontDescriptor = fd
        return
    except Exception:
        pass
    try:
        model.FontHeight = h
    except Exception:
        pass


def is_token_btn_label(label):
    s = unicode(label or u"").strip()
    if s in TOKEN_BTN_LABELS:
        return True
    # ASCII x / xx как замена × / ××
    low = s.casefold()
    return low in (u"+", u"x", u"xx")


def apply_token_button_font(model, height=None):
    """Крупнее шрифт на кнопках + / × / ××."""
    if model is None:
        return
    h = int(height if height is not None else TOKEN_BTN_FONT_HEIGHT)
    try:
        fd = model.FontDescriptor
        fd.Height = h
        try:
            fd.Weight = 150
        except Exception:
            pass
        model.FontDescriptor = fd
        return
    except Exception:
        pass
    try:
        model.FontHeight = h
    except Exception:
        pass
    try:
        model.FontWeight = 150
    except Exception:
        pass


def soft_gray_theme_applied(dm):
    if dm is None:
        return False
    try:
        return SOFT_GRAY_BAR_NAME in list(dm.getElementNames())
    except Exception:
        return False


def inner_dialog_footer_y(dh, m=10, btn_h=None, chk_h=None):
    btn_h = int(btn_h if btn_h is not None else INNER_BTN_H)
    chk_h = int(chk_h if chk_h is not None else INNER_CHK_H)
    return int(dh) - int(m) - btn_h - 4 - chk_h


def inner_dialog_set_sizeable(dm):
    try:
        dm.Sizeable = True
    except Exception:
        pass


def inner_dialog_add_ok_cancel_footer( dm, m, dh, ok_label=u"OK", cancel_label=u"Отмена", ok_name="OkButton", cancel_name="CancelButton", help_btn=False, dw=None, help_key=None):
    """OK и Отмена в одну строку слева (как футер визарда, без чекбокса JSON)."""
    y_btn = inner_dialog_footer_y(dh, m)
    ok_x = int(m)
    cancel_x = ok_x + INNER_BTN_W + INNER_BTN_GAP
    ok_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_btn.Name = unicode(ok_name)
    ok_btn.Label = unicode(ok_label)
    try:
        ok_btn.DefaultButton = True
    except Exception:
        pass
    try:
        ok_btn.PushButtonType = 1
    except Exception:
        pass
    ok_btn.PositionX = ok_x
    ok_btn.PositionY = y_btn
    ok_btn.Width = INNER_BTN_W
    ok_btn.Height = INNER_BTN_H
    dm.insertByName(unicode(ok_name), ok_btn)
    cancel_btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.Name = unicode(cancel_name)
    cancel_btn.Label = unicode(cancel_label)
    try:
        cancel_btn.PushButtonType = 2
    except Exception:
        pass
    cancel_btn.PositionX = cancel_x
    cancel_btn.PositionY = y_btn
    cancel_btn.Width = INNER_BTN_W
    cancel_btn.Height = INNER_BTN_H
    dm.insertByName(unicode(cancel_name), cancel_btn)
    if help_btn:
        help_w = 90
        try:
            dlg_w = int(dw) if dw is not None else int(getattr(dm, "Width", 0) or 0)
        except Exception:
            dlg_w = 0
        if dlg_w > 0:
            help_x = dlg_w - int(m) - help_w
        else:
            help_x = cancel_x + INNER_BTN_W + INNER_BTN_GAP
        help_btn_m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        help_btn_m.Name = "HelpButton"
        help_btn_m.Label = u"Справка"
        try:
            help_btn_m.PushButtonType = 0
        except Exception:
            pass
        help_btn_m.PositionX = help_x
        help_btn_m.PositionY = y_btn
        help_btn_m.Width = help_w
        help_btn_m.Height = INNER_BTN_H
        if help_key:
            try:
                help_btn_m.HelpText = unicode(help_key)
            except Exception:
                pass
        try:
            dm.insertByName("HelpButton", help_btn_m)
        except Exception:
            pass


def apply_soft_gray_theme(dm, title_text=None, title_bg=None, title_fg=None):
    """
    Серый фон формы + полоса-титл внутри диалога.
    Сдвигает контролы вниз. Идемпотентно (повторный вызов — no-op).
    title_bg/title_fg — переопределение цветов титла (по умолчанию тёмно-серый).
    """
    if dm is None:
        return False
    if soft_gray_theme_applied(dm):
        return True
    form_bg = int(SOFT_GRAY_FORM_BG)
    title_bg = int(title_bg if title_bg is not None else SOFT_GRAY_TITLE_BG)
    title_fg = int(title_fg if title_fg is not None else SOFT_GRAY_TITLE_FG)
    label_fg = int(SOFT_GRAY_LABEL_FG)
    field_bg = int(SOFT_GRAY_FIELD_BG)
    btn_bg = int(SOFT_GRAY_BTN_BG)
    title_h = int(SOFT_GRAY_TITLE_H)
    try:
        dm.BackgroundColor = form_bg
    except Exception:
        pass
    names = []
    try:
        names = list(dm.getElementNames())
    except Exception:
        names = []
    for name in names:
        try:
            ctl = dm.getByName(name)
        except Exception:
            continue
        try:
            ctl.PositionY = int(ctl.PositionY) + title_h
        except Exception:
            pass
        kind = u""
        try:
            if ctl.supportsService("com.sun.star.awt.UnoControlFixedTextModel"):
                kind = u"fixed"
            elif ctl.supportsService("com.sun.star.awt.UnoControlEditModel"):
                kind = u"edit"
            elif ctl.supportsService("com.sun.star.awt.UnoControlListBoxModel"):
                kind = u"list"
            elif ctl.supportsService("com.sun.star.awt.UnoControlComboBoxModel"):
                kind = u"combo"
            elif ctl.supportsService("com.sun.star.awt.UnoControlCheckBoxModel"):
                kind = u"check"
            elif ctl.supportsService("com.sun.star.awt.UnoControlButtonModel"):
                kind = u"button"
            elif ctl.supportsService("com.sun.star.awt.UnoControlProgressBarModel"):
                kind = u"progress"
            elif ctl.supportsService("com.sun.star.awt.UnoControlGroupBoxModel"):
                kind = u"group"
        except Exception:
            ncf = unicode(name or u"").casefold()
            if ncf.endswith(u"lbl") or u"hint" in ncf or u"label" in ncf:
                kind = u"fixed"
            elif ncf.endswith(u"ed") or ncf.endswith(u"edit"):
                kind = u"edit"
            elif ncf.endswith(u"lb") or u"list" in ncf:
                kind = u"list"
            elif ncf.endswith(u"cb") or u"combo" in ncf:
                kind = u"combo"
            elif ncf.endswith(u"chk"):
                kind = u"check"
            elif ncf.endswith(u"btn") or u"button" in ncf:
                kind = u"button"
            elif u"progress" in ncf:
                kind = u"progress"
            elif u"frame" in ncf or u"group" in ncf:
                kind = u"group"
        try:
            if kind == u"fixed":
                ctl.BackgroundColor = form_bg
                ctl.TextColor = label_fg
            elif kind in (u"edit", u"list", u"combo"):
                ctl.BackgroundColor = field_bg
                ctl.TextColor = label_fg
                if kind == u"list":
                    apply_listbox_font(ctl)
            elif kind == u"check":
                # Строки задач collect_pack красятся отдельно (зелёный/серый + фон выбора).
                if str(name or "").startswith("SheetChk"):
                    continue
                ctl.BackgroundColor = form_bg
                ctl.TextColor = label_fg
            elif kind == u"button":
                ctl.BackgroundColor = btn_bg
                ctl.TextColor = label_fg
                try:
                    if is_token_btn_label(getattr(ctl, "Label", u"")):
                        apply_token_button_font(ctl)
                except Exception:
                    pass
            elif kind in (u"progress", u"group"):
                ctl.BackgroundColor = form_bg
        except Exception:
            pass
    try:
        dm.Height = int(dm.Height) + title_h
    except Exception:
        pass
    caption = unicode(title_text or getattr(dm, "Title", u"") or u"").strip()
    if caption == u"":
        caption = u"Диалог"
    try:
        bar = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        bar.Name = SOFT_GRAY_BAR_NAME
        bar.Label = u"  " + caption
        bar.PositionX = 0
        bar.PositionY = 0
        bar.Width = int(dm.Width)
        bar.Height = title_h
        bar.BackgroundColor = title_bg
        bar.TextColor = title_fg
        try:
            bar.Align = 0
        except Exception:
            pass
        try:
            bar.VerticalAlign = 1
        except Exception:
            pass
        try:
            bar.Border = 0
        except Exception:
            pass
        try:
            fd = bar.FontDescriptor
            fd.Weight = 150
            try:
                fd.Height = 11
            except Exception:
                pass
            bar.FontDescriptor = fd
        except Exception:
            pass
        dm.insertByName(SOFT_GRAY_BAR_NAME, bar)
    except Exception:
        pass
    try:
        dm.Title = caption
    except Exception:
        pass
    return True


def try_paint_titlebar(dlg, title_bg=None, title_fg=None):
    """Попытка покрасить OS titlebar через StyleSettings (на GTK часто no-op)."""
    if dlg is None:
        return False
    title_bg = int(title_bg if title_bg is not None else SOFT_GRAY_TITLE_BG)
    title_fg = int(title_fg if title_fg is not None else SOFT_GRAY_TITLE_FG)
    form_bg = int(SOFT_GRAY_FORM_BG)
    painted = False
    peers = []
    try:
        peers.append(dlg.getPeer())
    except Exception:
        pass
    try:
        peers.append(dlg.getPeer().getContainerWindow())
    except Exception:
        pass
    try:
        import uno
        ctx = uno.getComponentContext()
        toolkit = ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx
        )
        if toolkit is not None:
            peers.append(toolkit)
    except Exception:
        pass
    attrs = (
        ("TitleBarColor", title_bg),
        ("TitleBarTextColor", title_fg),
        ("InactiveTitleBarColor", title_bg),
        ("InactiveTitleBarTextColor", title_fg),
        ("DialogColor", form_bg),
        ("WorkspaceColor", form_bg),
        ("FaceColor", form_bg),
        ("WindowColor", form_bg),
    )
    for peer in peers:
        if peer is None:
            continue
        ss = None
        try:
            ss = peer.StyleSettings
        except Exception:
            try:
                ss = peer.getStyleSettings()
            except Exception:
                ss = None
        if ss is None:
            continue
        for attr, val in attrs:
            try:
                setattr(ss, attr, val)
                painted = True
            except Exception:
                try:
                    ss.setPropertyValue(attr, val)
                    painted = True
                except Exception:
                    pass
    return painted


def apply_soft_gray_warning_theme(dm, title_text=None):
    """Soft-gray форма с алой полосой-титлом."""
    return apply_soft_gray_theme(
        dm,
        title_text=title_text,
        title_bg=WARNING_TITLE_BG,
        title_fg=WARNING_TITLE_FG,
    )


def apply_soft_gray_orange_theme(dm, title_text=None):
    """Soft-gray форма с оранжевой полосой-титлом (плагины и т.п.)."""
    return apply_soft_gray_theme(
        dm,
        title_text=title_text,
        title_bg=ORANGE_TITLE_BG,
        title_fg=ORANGE_TITLE_FG,
    )


def apply_soft_gray_manual_input_theme(dm, title_text=None):
    """Soft-gray форма с жёлтой полосой-титлом (ручной ввод переменной)."""
    return apply_soft_gray_theme(
        dm,
        title_text=title_text,
        title_bg=MANUAL_INPUT_TITLE_BG,
        title_fg=MANUAL_INPUT_TITLE_FG,
    )


def apply_soft_gray_dark_blue_theme(dm, title_text=None):
    """Soft-gray форма с тёмно-синей полосой-титлом (как у сбора, другой цвет)."""
    return apply_soft_gray_theme(
        dm,
        title_text=title_text,
        title_bg=DARK_BLUE_TITLE_BG,
        title_fg=DARK_BLUE_TITLE_FG,
    )


def prepare_dialog_soft_gray(dlg, title_text=None, title_bg=None, title_fg=None):
    """Применить тему к модели диалога (до или после setModel, до createPeer)."""
    if dlg is None:
        return False
    try:
        dm = dlg.getModel()
    except Exception:
        return False
    caption = title_text
    if caption is None:
        try:
            caption = getattr(dm, "Title", None)
        except Exception:
            caption = None
    return apply_soft_gray_theme(
        dm, title_text=caption, title_bg=title_bg, title_fg=title_fg,
    )


def prepare_dialog_soft_gray_warning(dlg, title_text=None):
    """Soft-gray форма с алой полосой-титлом."""
    return prepare_dialog_soft_gray(
        dlg,
        title_text=title_text,
        title_bg=WARNING_TITLE_BG,
        title_fg=WARNING_TITLE_FG,
    )


def prepare_dialog_soft_gray_orange(dlg, title_text=None):
    """Soft-gray форма с оранжевой полосой-титлом."""
    return prepare_dialog_soft_gray(
        dlg,
        title_text=title_text,
        title_bg=ORANGE_TITLE_BG,
        title_fg=ORANGE_TITLE_FG,
    )


def prepare_dialog_soft_gray_manual_input(dlg, title_text=None):
    """Soft-gray форма с жёлтой полосой-титлом (ручной ввод переменной)."""
    return prepare_dialog_soft_gray(
        dlg,
        title_text=title_text,
        title_bg=MANUAL_INPUT_TITLE_BG,
        title_fg=MANUAL_INPUT_TITLE_FG,
    )


def prepare_dialog_soft_gray_dark_blue(dlg, title_text=None):
    """Soft-gray форма с тёмно-синей полосой-титлом."""
    return prepare_dialog_soft_gray(
        dlg,
        title_text=title_text,
        title_bg=DARK_BLUE_TITLE_BG,
        title_fg=DARK_BLUE_TITLE_FG,
    )
