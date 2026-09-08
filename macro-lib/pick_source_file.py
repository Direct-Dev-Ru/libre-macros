# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.702"
# Выбор файла для строки «Файлы-Источники»
# Кнопка на листе -> макрос pick_source_file.select_source_file

import os
import uno

# Константы в pythonpath: AlterOffice 2026 вырезает модульные присваивания в .py-скрипте.
from libre_macros_pick_source_cfg import HEADER_ROW, PARAM_NAME, RANGE_FORMAT_PROPS as _RANGE_FORMAT_PROPS


def _cell_string(cell):
    if cell is None:
        return u""
    s = cell.String
    if s is None:
        return u""
    return str(s).strip()


def _value_header_label(col_index):
    """Столбец B (1) -> «Значение 1»."""
    return u"Значение %d" % int(col_index)


def _is_value_header(text):
    t = str(text).strip().lower()
    return t.startswith(u"значение")


def _sheet_last_used_row(sheet):
    try:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        return int(cursor.RangeAddress.EndRow)
    except Exception:
        return 500


def _copy_cell_range_format(src_range, dst_range):
    """Скопировать оформление с одного XCellRange на другой (значения не трогаем)."""
    if src_range is None or dst_range is None:
        return
    pi = 0
    while pi < len(_RANGE_FORMAT_PROPS):
        prop = _RANGE_FORMAT_PROPS[pi]
        try:
            setattr(dst_range, prop, getattr(src_range, prop))
        except Exception:
            pass
        pi = pi + 1


def _copy_row_format(sheet, src_col, dst_col, row):
    try:
        src_range = sheet.getCellRangeByPosition(src_col, row, src_col, row)
        dst_range = sheet.getCellRangeByPosition(dst_col, row, dst_col, row)
    except Exception:
        return
    _copy_cell_range_format(src_range, dst_range)


def _save_column_values(sheet, col, first_row, last_row):
    values = []
    row = first_row
    while row <= last_row:
        values.append(_cell_string(sheet.getCellByPosition(col, row)))
        row = row + 1
    return values


def _restore_column_values(sheet, col, first_row, values):
    row = first_row
    idx = 0
    while idx < len(values):
        sheet.getCellByPosition(col, row).String = values[idx]
        row = row + 1
        idx = idx + 1


def _copy_column_width(sheet, src_col, dst_col):
    try:
        cols = sheet.getColumns()
        src_obj = cols.getByIndex(src_col)
        dst_obj = cols.getByIndex(dst_col)
        dst_obj.Width = src_obj.Width
    except Exception:
        pass
    try:
        cols = sheet.getColumns()
        dst_obj = cols.getByIndex(dst_col)
        src_obj = cols.getByIndex(src_col)
        dst_obj.setPropertyValue("Width", src_obj.getPropertyValue("Width"))
    except Exception:
        pass


def _get_dispatch_helper():
    ctx = XSCRIPTCONTEXT.getComponentContext()
    return ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", ctx
    )


def _ensure_active_sheet(controller, sheet):
    try:
        active = controller.getActiveSheet()
        if active is None or active.Name != sheet.Name:
            controller.setActiveSheet(sheet)
    except Exception:
        try:
            controller.setActiveSheet(sheet)
        except Exception:
            pass


def _focus_cell(doc, sheet, col, row):
    """Выделить ячейку после записи пути (фокус на новое значение)."""
    if doc is None or sheet is None:
        return
    try:
        controller = doc.getCurrentController()
        if controller is None:
            return
        _ensure_active_sheet(controller, sheet)
        cell_range = sheet.getCellRangeByPosition(col, row, col, row)
        controller.select(cell_range)
    except Exception:
        pass


def _paste_format_flags():
    """CellFlags: HARDATTR | STYLES (без VALUE/STRING)."""
    try:
        from com.sun.star.sheet.CellFlags import HARDATTR, STYLES

        return int(HARDATTR) | int(STYLES)
    except Exception:
        return 96


def _copy_column_formats_dispatch(doc, sheet, src_col, dst_col, first_row, last_row):
    """
    Копирование оформления столбца через буфер обмена Calc.
    Надёжнее прямого присвоения свойств ячейкам в LO.
    """
    if doc is None or sheet is None:
        return False
    if src_col < 0 or dst_col < 0 or last_row < first_row:
        return False
    controller = doc.getCurrentController()
    if controller is None:
        return False
    try:
        _ensure_active_sheet(controller, sheet)
        src_range = sheet.getCellRangeByPosition(
            src_col, first_row, src_col, last_row
        )
        dst_range = sheet.getCellRangeByPosition(
            dst_col, first_row, dst_col, last_row
        )
        dispatcher = _get_dispatch_helper()
        frame = controller.getFrame()
        if dispatcher is None or frame is None:
            return False

        saved_values = _save_column_values(sheet, dst_col, first_row, last_row)

        controller.select(src_range)
        dispatcher.executeDispatch(frame, ".uno:Copy", "", 0, ())

        flags = _paste_format_flags()
        dest_addr = dst_range.getRangeAddress()
        pasted = False
        try:
            sheet.pasteCellRange(
                dest_addr,
                0,
                flags,
                False,
                False,
                False,
                0,
            )
            pasted = True
        except Exception:
            pass
        if not pasted:
            try:
                from com.sun.star.sheet import XSheetPastable

                pastable = sheet.queryInterface(XSheetPastable)
                if pastable is not None:
                    pastable.pasteCellRange(
                        dest_addr,
                        0,
                        flags,
                        False,
                        False,
                        False,
                        0,
                    )
                    pasted = True
            except Exception:
                pass

        _restore_column_values(sheet, dst_col, first_row, saved_values)
        return pasted
    except Exception:
        return False


def _copy_column_formats(sheet, src_col, dst_col, first_row, last_row):
    """Форматы столбца: построчно через диапазоны + ширина колонки."""
    if src_col < 0 or dst_col < 0 or last_row < first_row:
        return
    row = first_row
    while row <= last_row:
        _copy_row_format(sheet, src_col, dst_col, row)
        row = row + 1
    _copy_column_width(sheet, src_col, dst_col)


def _prepare_value_column(doc, sheet, col):
    """
    Если в строке заголовка нет «Значение N» — оформление с предыдущего столбца
    и подпись в первой строке.
    """
    if col < 1:
        return
    hdr_cell = sheet.getCellByPosition(col, HEADER_ROW)
    hdr = _cell_string(hdr_cell)
    if hdr != u"" and _is_value_header(hdr):
        return
    prev_col = col - 1
    last_row = _sheet_last_used_row(sheet)
    if last_row < HEADER_ROW:
        last_row = HEADER_ROW
    if not _copy_column_formats_dispatch(
        doc, sheet, prev_col, col, HEADER_ROW, last_row
    ):
        _copy_column_formats(sheet, prev_col, col, HEADER_ROW, last_row)
    hdr_cell.String = _value_header_label(col)


def _last_path_in_row(sheet, row):
    """Последнее непустое значение в строке (столбцы B, C, D…)."""
    last = u""
    col = 1
    while col < 50:
        b = sheet.getCellByPosition(col, row).String
        if b is None:
            b = u""
        b = str(b).strip()
        if b == u"":
            break
        last = b
        col = col + 1
    return last


def _picker_initial(path_str):
    """
    Старт FilePicker из последнего пути: (каталог, имя_файла_в_поле_ввода).
    Каталог без имени файла; имя — basename последнего значения (маска *.xlsx тоже).
    """
    home = os.path.expanduser("~")
    if path_str is None or str(path_str).strip() == u"":
        return home, u""
    p = os.path.normpath(str(path_str).strip())
    name = os.path.basename(p)
    if name == u"" or name == u".":
        name = u""

    d = os.path.dirname(p)
    if d == u"":
        d = home
    elif not os.path.isdir(d):
        abs_d = os.path.dirname(os.path.abspath(p))
        if abs_d != u"" and os.path.isdir(abs_d):
            d = abs_d
        elif os.path.isdir(os.path.abspath(p)):
            d = os.path.abspath(p)
            name = u""
        else:
            d = home

    if os.path.isfile(p):
        d = os.path.dirname(os.path.abspath(p))
        name = os.path.basename(p)

    return d, name


def _dir_to_file_url(dir_path):
    """Каталог в file:// URL для setDisplayDirectory (не системный путь)."""
    if dir_path is None or str(dir_path).strip() == u"":
        dir_path = os.path.expanduser("~")
    abs_dir = os.path.abspath(str(dir_path).strip())
    if not os.path.isdir(abs_dir):
        return None
    return uno.systemPathToFileUrl(abs_dir)


def _get_selection_cell(doc):
    try:
        ctrl = doc.getCurrentController()
        if ctrl is None:
            return None, None
        sel = ctrl.getSelection()
        if sel is None:
            return None, None
        addr = sel.getRangeAddress()
        return int(addr.StartColumn), int(addr.StartRow)
    except Exception:
        return None, None


def _find_files_param_row(sheet, start_row=None):
    """Строка «Файлы-Источники» на листе (0-based)."""
    if start_row is not None and int(start_row) >= 1:
        name = _cell_string(sheet.getCellByPosition(0, int(start_row)))
        if _param_match(name, PARAM_NAME):
            return int(start_row)
    r = 1
    while r < 500:
        a = _cell_string(sheet.getCellByPosition(0, r))
        if r == 1 and a.lower() == u"параметр":
            r = r + 1
            continue
        if a == u"":
            break
        if a.lower() == PARAM_NAME.lower():
            return r
        r = r + 1
    return -1


def _param_match(name, key):
    return str(name).strip().lower() == str(key).strip().lower()


def _apply_picker_initial(picker, start_dir, start_name):
    """Начальная папка и имя файла; при ошибке диалог откроется без них."""
    dir_url = _dir_to_file_url(start_dir)
    if dir_url is not None:
        try:
            picker.setDisplayDirectory(dir_url)
        except Exception:
            pass
    if start_name != u"":
        try:
            picker.setDefaultName(start_name)
        except Exception:
            pass


def _doc_system_path(doc):
    try:
        url = doc.getURL()
        if url:
            return os.path.normpath(
                os.path.abspath(uno.fileUrlToSystemPath(url))
            )
    except Exception:
        pass
    return ""


def _normalize_source_path_for_cell(doc, path):
    try:
        import collect_workbooks as cw
        return cw.normalize_source_path_cell_value(path, _doc_system_path(doc))
    except Exception:
        return path


def select_source_file(*args):
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.getCurrentController().getActiveSheet()

    cur_col, cur_row = _get_selection_cell(doc)
    row = _find_files_param_row(sheet, cur_row)

    if row < 0:
        print(u"Строка «%s» не найдена на листе %s" % (PARAM_NAME, sheet.Name))
        return

    target_col = None
    initial_path = u""
    if (
        cur_col is not None
        and cur_row is not None
        and int(cur_row) == int(row)
        and int(cur_col) >= 1
    ):
        cur_val = _cell_string(sheet.getCellByPosition(int(cur_col), int(row)))
        if cur_val != u"":
            target_col = int(cur_col)
            initial_path = cur_val

    ctx = XSCRIPTCONTEXT.getComponentContext()
    sm = ctx.getServiceManager()
    picker = sm.createInstanceWithContext(
        "com.sun.star.ui.dialogs.FilePicker", ctx
    )
    if target_col is not None:
        start_dir, start_name = _picker_initial(initial_path)
    else:
        start_dir, start_name = _picker_initial(_last_path_in_row(sheet, row))
    _apply_picker_initial(picker, start_dir, start_name)
    if picker.execute() != 1:
        return

    urls = picker.getFiles()
    if urls is None or len(urls) == 0:
        return

    path = uno.fileUrlToSystemPath(urls[0])
    path = os.path.abspath(path)
    path = _normalize_source_path_for_cell(doc, path)

    if target_col is not None:
        _prepare_value_column(doc, sheet, target_col)
        sheet.getCellByPosition(target_col, row).String = path
        _focus_cell(doc, sheet, target_col, row)
        print(u"Обновлено: %s" % path)
        return

    col = 1
    while col < 50:
        if _cell_string(sheet.getCellByPosition(col, row)) == u"":
            _prepare_value_column(doc, sheet, col)
            sheet.getCellByPosition(col, row).String = path
            _focus_cell(doc, sheet, col, row)
            print(u"Записано: %s" % path)
            return
        col = col + 1

    print(u"Нет свободных ячеек в строке «%s»" % PARAM_NAME)


CONTEXT_CELL_MENU_KEY = "pick_source_file"
CONTEXT_CELL_MENU = (
    "submenu#PythonMacros",
    # "submenu#MergeWorkbooks",
    "run_function#select_source_file;module#pick_source_file.py;display_name#Выбрать источник;order#60",
)




g_exportedScripts = (select_source_file,)
