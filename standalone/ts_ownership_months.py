# -*- coding: utf-8 -*-
"""
Макрос: подсчёт месяцев владения ТС в квартале (LibreOffice Calc).

Раскладка на активном листе (строка 1 — заголовки, данные со 2-й):
  A — дата выбытия ТС          (ДД.ММ.ГГГГ, напр. 29.01.2026)
  B — дата снятия с учёта       (может быть пустой)
  C — квартал и год             (напр. 2.2026 = II квартал 2026)
  D — не используется (в образце — описание сценария)
  E — результат (число месяцев)

Примеры (колонки A / B / C → E):
  06.02.2026 / 27.06.2026 / 2.2026 → 3
  20.03.2026 / 19.05.2026 / 2.2026 → 2
  29.01.2026 / 20.04.2026 / 2.2026 → 1

Установка:
  1. Открыть ts_ownership_months.ods (макрос вшит в документ) и нажать кнопку
     «Посчитать месяцы» на листе Sample.
  2. Либо скопировать этот файл в каталог Scripts/python/ профиля LibreOffice
     и запустить: Сервис → Выполнить макрос → Python → ts_ownership_months.

Подробный лог пишется в консоль (print) и на новый лист «Лог_ТС_месяцы».
"""

from datetime import date


# Константы и журнал — внутри class: так их не вырезает фильтр AlterOffice 2026.
class _Store:
    lines = []
    QUARTER_MONTHS = {
        1: (1, 2, 3),
        2: (4, 5, 6),
        3: (7, 8, 9),
        4: (10, 11, 12),
    }
    MONTH_NAMES = {
        1: "январь",
        2: "февраль",
        3: "март",
        4: "апрель",
        5: "май",
        6: "июнь",
        7: "июль",
        8: "август",
        9: "сентябрь",
        10: "октябрь",
        11: "ноябрь",
        12: "декабрь",
    }


def _log(msg):
    """Строка в память + print в консоль терминала / stdout LO."""
    _Store.lines.append(msg)
    try:
        print("[ts_ownership] %s" % msg)
    except Exception:
        pass


def _reset_log():
    del _Store.lines[:]


# --- разбор ячеек -----------------------------------------------------------

def _fmt_date(value):
    if value is None:
        return "(нет)"
    return value.strftime("%d.%m.%Y")


def parse_date(text):
    """Строка ДД.ММ.ГГГГ → date. Пустая строка → None. Ошибка → None."""
    text = (text or "").strip()
    if not text:
        return None
    parts = text.split(".")
    if len(parts) != 3:
        return None
    try:
        day = int(parts[0])
        month = int(parts[1])
        year = int(parts[2])
        return date(year, month, day)
    except (ValueError, TypeError):
        return None


def parse_quarter(text):
    """Строка «2.2026» → (квартал, год). Ошибка → None."""
    text = (text or "").strip().replace(",", ".")
    if not text or "." not in text:
        return None
    parts = text.split(".", 1)
    try:
        quarter = int(parts[0])
        year = int(parts[1])
    except (ValueError, TypeError):
        return None
    if quarter < 1 or quarter > 4:
        return None
    return quarter, year


# --- логика подсчёта --------------------------------------------------------

def count_months(disposal, deregistration, quarter, year, verbose=False):
    """
    Сколько месяцев квартала засчитывается по правилам:
      1) выбытие до 15.M включительно и (снятие после 15.M или снятия нет) → считаем;
      2) выбытие до 15.M и снятие до 15.M включительно → не считаем;
      3) выбытие после 15.M → не считаем.
    verbose=True — подробные строки в _log.
    """
    months = _Store.QUARTER_MONTHS[quarter]
    total = 0
    cutoff_fmt = "%d.%m.%Y"

    for month in months:
        cutoff = date(year, month, 15)
        cutoff_text = cutoff.strftime(cutoff_fmt)
        month_title = "%s (%s)" % (_Store.MONTH_NAMES.get(month, str(month)), cutoff_text)

        # правило 3: выбытие после 15-го числа месяца M
        if disposal > cutoff:
            if verbose:
                _log(
                    "    %s: НЕТ — правило 3: выбытие %s позже 15-го (%s)"
                    % (month_title, _fmt_date(disposal), cutoff_text)
                )
            continue

        # пустое снятие — смотрим только на выбытие (правило 1 без снятия)
        if deregistration is None:
            total += 1
            if verbose:
                _log(
                    "    %s: ДА — правило 1: снятия нет, выбытие %s не позже 15-го (%s)"
                    % (month_title, _fmt_date(disposal), cutoff_text)
                )
            continue

        # правило 2: снятие до 15.M включительно
        if deregistration <= cutoff:
            if verbose:
                _log(
                    "    %s: НЕТ — правило 2: выбытие %s не позже 15-го (%s), "
                    "снятие %s не позже 15-го (%s)"
                    % (
                        month_title,
                        _fmt_date(disposal),
                        cutoff_text,
                        _fmt_date(deregistration),
                        cutoff_text,
                    )
                )
            continue

        # правило 1: выбытие до 15.M и снятие после 15.M
        total += 1
        if verbose:
            _log(
                "    %s: ДА — правило 1: выбытие %s не позже 15-го (%s), "
                "снятие %s позже 15-го (%s)"
                % (
                    month_title,
                    _fmt_date(disposal),
                    cutoff_text,
                    _fmt_date(deregistration),
                    cutoff_text,
                )
            )

    return total


# --- LibreOffice ------------------------------------------------------------

def _cell_text(sheet, col, row):
    """Текст ячейки (col и row — с нуля)."""
    return sheet.getCellByPosition(col, row).getString().strip()


def _write_result(sheet, row, value):
    """Записать число или текст «ошибка» в колонку E."""
    cell = sheet.getCellByPosition(4, row)
    if isinstance(value, str):
        cell.setString(value)
    else:
        cell.setValue(float(value))


def _write_log_sheet(doc, lines):
    """Создать новый лист и записать туда все строки журнала."""
    if not lines:
        return None

    base_name = "Лог_ТС_месяцы"
    sheets = doc.getSheets()
    sheet_name = base_name
    suffix = 1
    while sheets.hasByName(sheet_name):
        suffix += 1
        sheet_name = "%s_%d" % (base_name, suffix)

    new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
    sheets.insertByName(sheet_name, new_sheet)

    row = 0
    while row < len(lines):
        new_sheet.getCellByPosition(0, row).setString(lines[row])
        row += 1

    try:
        new_sheet.getColumns().getByIndex(0).Width = 18000
    except Exception:
        pass

    return sheet_name


def ts_ownership_months(*args):
    """Точка входа: обходит строки активного листа, пишет результат в E."""
    try:
        import uno
    except ImportError:
        return

    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except NameError:
        pass
    if doc is None:
        try:
            import __main__
            xsc = getattr(__main__, "XSCRIPTCONTEXT", None)
            if xsc is not None:
                doc = xsc.getDocument()
        except Exception:
            pass
    if doc is None:
        _log("ошибка: документ Calc не найден")
        return

    _reset_log()
    sheet = doc.getCurrentController().getActiveSheet()
    sheet_name = sheet.getName()

    _log("=== старт: лист «%s» ===" % sheet_name)
    _log("строка 1 — заголовки; данные со 2-й; останов при пустой A")

    row = 1  # строка 2 в Calc (0-based row 1)
    processed = 0
    errors = 0

    while True:
        calc_row = row + 1
        disposal_text = _cell_text(sheet, 0, row)
        if not disposal_text:
            _log("--- конец данных: строка %d, колонка A пуста ---" % calc_row)
            break

        dereg_text = _cell_text(sheet, 1, row)
        quarter_text = _cell_text(sheet, 2, row)

        _log("")
        _log("--- строка %d ---" % calc_row)
        _log("  A выбытие: «%s»" % disposal_text)
        _log("  B снятие:  «%s»" % (dereg_text if dereg_text else "(пусто)"))
        _log("  C квартал: «%s»" % quarter_text)

        disposal = parse_date(disposal_text)
        if disposal is None:
            _log("  итог: ОШИБКА — не удалось разобрать дату выбытия")
            _write_result(sheet, row, "ошибка")
            errors += 1
            row += 1
            continue

        parsed = parse_quarter(quarter_text)
        if parsed is None:
            _log("  итог: ОШИБКА — не удалось разобрать квартал (ожидается вид 2.2026)")
            _write_result(sheet, row, "ошибка")
            errors += 1
            row += 1
            continue

        quarter, year = parsed
        deregistration = parse_date(dereg_text) if dereg_text else None

        _log("  разбор: выбытие=%s, снятие=%s, Q%d %d" % (
            _fmt_date(disposal),
            _fmt_date(deregistration),
            quarter,
            year,
        ))
        _log("  месяцы квартала:")

        result = count_months(disposal, deregistration, quarter, year, verbose=True)
        _write_result(sheet, row, result)
        _log("  итог: %d мес. → записано в E%d" % (result, calc_row))
        processed += 1
        row += 1

    _log("")
    _log("=== сводка ===")
    _log("обработано строк: %d" % processed)
    _log("ошибок: %d" % errors)
    _log("всего строк журнала: %d" % len(_Store.lines))

    log_sheet_name = _write_log_sheet(doc, _Store.lines)
    if log_sheet_name:
        _log("журнал записан на лист «%s»" % log_sheet_name)


g_exportedScripts = (ts_ownership_months,)
