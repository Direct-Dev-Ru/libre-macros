# -*- coding: utf-8 -*-
MACRO_VERSION = "3.10.499"
import uno
from com.sun.star.beans import PropertyValue


def _log(msg):
    try:
        print("[test_sort] %s" % msg)
    except Exception:
        pass


def _dump_range(range_to_sort, label):
    try:
        data = range_to_sort.getDataArray()
        _log("%s (%d строк):" % (label, len(data)))
        ri = 0
        while ri < len(data):
            _log("  [%d] %r" % (ri, data[ri]))
            ri += 1
    except Exception as err:
        _log("%s: getDataArray FAIL: %s" % (label, err))


def _col_values(range_to_sort, col_index, skip_header=True):
    data = range_to_sort.getDataArray()
    start = 1 if skip_header else 0
    out = []
    ri = start
    while ri < len(data):
        row = data[ri]
        out.append(row[col_index] if col_index < len(row) else None)
        ri += 1
    return out


def _describe_sort_fields(sort_fields):
    _log("SortFields (%d):" % len(sort_fields))
    i = 0
    while i < len(sort_fields):
        fld = sort_fields[i]
        parts = ["Field=%s" % getattr(fld, "Field", "?")]
        parts.append("IsAscending=%s" % getattr(fld, "IsAscending", "?"))
        if hasattr(fld, "FieldType"):
            parts.append("FieldType=%s" % getattr(fld, "FieldType", "?"))
        _log("  [%d] %s" % (i, ", ".join(parts)))
        i += 1


def _make_sort_field(field, ascending=True, field_type=None):
    sort_field = uno.createUnoStruct("com.sun.star.table.TableSortField")
    sort_field.Field = int(field)
    sort_field.IsAscending = bool(ascending)
    if field_type is not None:
        sort_field.FieldType = field_type
    return sort_field


def _sort_fields_value(sort_fields):
    try:
        return uno.Any("[]com.sun.star.table.TableSortField", sort_fields)
    except Exception:
        return sort_fields


def _patch_sort_descriptor(desc, sort_fields, contains_header):
    patched = []
    i = 0
    while i < len(desc):
        prop = desc[i]
        if prop.Name == "SortFields":
            prop.Value = _sort_fields_value(sort_fields)
        elif prop.Name == "ContainsHeader":
            prop.Value = contains_header
        patched.append(prop)
        i += 1

    names = [p.Name for p in patched]
    if "SortFields" not in names:
        prop_fields = PropertyValue()
        prop_fields.Name = "SortFields"
        prop_fields.Value = _sort_fields_value(sort_fields)
        patched.append(prop_fields)
    if "ContainsHeader" not in names:
        prop_header = PropertyValue()
        prop_header.Name = "ContainsHeader"
        prop_header.Value = contains_header
        patched.append(prop_header)

    _log("дескриптор: ContainsHeader=%s, SortFields=%d полей" % (
        contains_header, len(sort_fields),
    ))
    return tuple(patched)


def _sort_range_native(range_to_sort, sort_fields, contains_header=True):
    """
    XSortable.sort + createSortDescriptor (DevGuide — Spreadsheet Sorting).
    https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Sorting
    """
    if not hasattr(range_to_sort, "sort"):
        raise RuntimeError("у диапазона нет метода sort()")

    _describe_sort_fields(sort_fields)
    sort_desc = _patch_sort_descriptor(
        range_to_sort.createSortDescriptor(),
        sort_fields,
        contains_header,
    )
    range_to_sort.sort(sort_desc)


def _sort_range(sheet, range_name, sort_fields, contains_header=True):
    _log("=== сортировка %s ===" % range_name)
    range_to_sort = sheet.getCellRangeByName(range_name)
    _dump_range(range_to_sort, "ДО " + range_name)
    _sort_range_native(range_to_sort, sort_fields, contains_header=contains_header)
    _dump_range(range_to_sort, "ПОСЛЕ " + range_name)
    _log("=== конец %s ===" % range_name)


def _write_text_numbers(sheet, left, top, right, bottom, text_col, data):
    block = sheet.getCellRangeByPosition(left, top, right, bottom)
    block.setDataArray(data)

    row = top + 1
    while row <= bottom:
        rel_row = row - top
        value = data[rel_row][text_col - left]
        sheet.getCellByPosition(text_col, row).setString(str(value))
        row += 1

    _log("столбец %s: значения записаны как текст (setString)" % chr(65 + text_col))


def fill_and_sort_data():
    _log("старт")
    _log(
        "DevGuide: в Calc FieldType (NUMERIC/ALPHANUMERIC) игнорируется — "
        "сортировка идёт по фактическому типу ячейки."
    )

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    sheet = doc.getCurrentController().getActiveSheet()
    _log("лист: %s" % sheet.getName())

    # --- Пример 1: A1:C10 — текст, два ключа ---
    data1 = (
        ("Отдел", "Фамилия", "Зарплата"),
        ("Бухгалтерия", "Иванов", 50000),
        ("IT", "Петров", 80000),
        ("Бухгалтерия", "Сидоров", 45000),
        ("IT", "Алексеев", 90000),
        ("HR", "Васильев", 60000),
        ("IT", "Смирнов", 75000),
        ("HR", "Кузнецов", 55000),
        ("Бухгалтерия", "Попов", 48000),
        ("IT", "Соколов", 85000),
    )
    sheet.getCellRangeByPosition(0, 0, 2, 9).setDataArray(data1)
    _sort_range(sheet, "A1:C10", (
        _make_sort_field(0, True),
        _make_sort_field(1, True),
    ))

    # --- Пример 2a: E1:G5 — «числа» как текст, FieldType=NUMERIC не помогает ---
    data2 = (
        ("Код", "Товар", "Кол-во"),
        ("B", "Яблоки", "100"),
        ("C", "Груши", "20"),
        ("A", "Апельсины", "9"),
        ("D", "Сливы", "450"),
    )
    _write_text_numbers(sheet, 4, 0, 6, 4, text_col=6, data=data2)

    from com.sun.star.table.TableSortFieldType import NUMERIC

    _sort_range(sheet, "E1:G5", (_make_sort_field(2, True, NUMERIC),))
    qty_text = _col_values(sheet.getCellRangeByName("E1:G5"), 2)
    _log(
        "пример 2a (текст): Кол-во после сортировки = %r" % qty_text
    )
    _log(
        "ожидание как ТЕКСТ: ['100', '20', '450', '9']; "
        "как ЧИСЛА Calc не умеет для setString-ячеек"
    )

    # --- Пример 2b: I1:K5 — те же данные, но настоящие числа в ячейках ---
    data2_num = (
        ("Код", "Товар", "Кол-во"),
        ("B", "Яблоки", 100),
        ("C", "Груши", 20),
        ("A", "Апельсины", 9),
        ("D", "Сливы", 450),
    )
    sheet.getCellRangeByPosition(8, 0, 10, 4).setDataArray(data2_num)
    _sort_range(sheet, "I1:K5", (_make_sort_field(2, True, NUMERIC),))
    qty_num = _col_values(sheet.getCellRangeByName("I1:K5"), 2)
    _log("пример 2b (числа): Кол-во после сортировки = %r" % qty_num)
    _log("ожидание как ЧИСЛА: [9, 20, 100, 450]")

    from com.sun.star.awt.MessageBoxType import INFOBOX
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK

    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent = getattr(
        getattr(getattr(doc, "CurrentController", None), "Frame", None),
        "ContainerWindow",
        None,
    )
    if parent is None:
        parent = toolkit.getDesktopWindow()

    msg_box = toolkit.createMessageBox(
        parent,
        INFOBOX,
        BUTTONS_OK,
        "Макрос выполнен",
        "A1:C10 — сортировка по отделу и фамилии.\n"
        "E1:G5 — текстовые «числа» → сортировка как текст (100, 20, 450, 9).\n"
        "I1:K5 — настоящие числа → сортировка как числа (9, 20, 100, 450).\n"
        "Подробности в консоли.",
    )
    msg_box.execute()
    _log("конец")


def formula_fill():
    import traceback

    def log(msg):
        print("[DEBUG] %s" % msg)

    log("=== formula_fill: старт ===")

    try:
        doc = XSCRIPTCONTEXT.getDocument()
        sheet = doc.getCurrentController().getActiveSheet()

        target_col = 9          # J
        formula_row = 1         # строка 2
        right_count = 5         # K, L, M, N, O
        end_col = target_col + right_count   # O = 14

        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(True)
        end_row = cursor.getRangeAddress().EndRow

        formula = sheet.getCellByPosition(target_col, formula_row).getFormula()
        log("Лист: %r" % sheet.getName())
        log("J2 formula: %r" % formula)
        log("Строки: %d..%d, столбцы: %d..%d (J..O)"
            % (formula_row + 1, end_row + 1, target_col + 1, end_col + 1))

        if not formula or end_row <= formula_row:
            log("выход")
            return

        # --- шаг 1: вниз по J ---
        col_range = sheet.getCellRangeByPosition(
            target_col, formula_row, target_col, end_row
        )
        log("шаг 1: fillSeries TO_BOTTOM(0) на J%d:J%d"
            % (formula_row + 1, end_row + 1))
        col_range.fillSeries(0, 0, 0, 1.0, 1.0e99)   # вниз, SIMPLE

        c = sheet.getCellByPosition(target_col, 2)
        log("после шага 1 — J3: formula=%r" % c.getFormula())

        # --- шаг 2: вправо на 5 ячеек на каждой строке ---
        block = sheet.getCellRangeByPosition(
            target_col, formula_row, end_col, end_row
        )
        log("шаг 2: fillSeries TO_RIGHT(1) на J%d:O%d"
            % (formula_row + 1, end_row + 1))
        block.fillSeries(1, 0, 0, 1.0, 1.0e99)       # вправо, SIMPLE

        # проверка
        for r in (1, 2, 10, end_row):
            parts = []
            col = target_col
            while col <= end_col:
                cell = sheet.getCellByPosition(col, r)
                col_letter = chr(ord('A') + col) if col < 26 else "c%d" % col
                parts.append("%s=%r" % (col_letter, cell.getFormula()[:20]))
                col += 1
            log("строка %d: %s" % (r + 1, " | ".join(parts)))

        log("=== успех ===")

    except Exception:
        log("ИСКЛЮЧЕНИЕ:\n%s" % traceback.format_exc())
    finally:
        log("=== конец ===")




g_exportedScripts = (fill_and_sort_data, formula_fill)
