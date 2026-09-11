# -*- coding: utf-8 -*-
"""
Расчёт штрихкодов и генерация карточек с QR-кодами (segno bundled).
"""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.724"
import json
import os
import re
import sys
import tempfile
from datetime import datetime

import uno
import unohelper
from com.sun.star.awt import XActionListener, XItemListener

import segno_bundled  # noqa: F401 — регистрирует import hook
import segno

from libre_macros_qr_codes_cfg import (
    COL_INV,
    COL_NAME,
    COL_SN1,
    COL_SN2,
    COL_SN3,
    DATA_START_ROW,
    DEFAULT_MAP_INV,
    DEFAULT_MAP_NAME,
    DEFAULT_MAP_QR,
    DEFAULT_MAP_SERIAL,
    DEFAULT_MAP_SIZE_MM,
    DEFAULT_BAR_COL0,
    DEFAULT_BAR_COL1,
    DEFAULT_BAR_COL2,
    DEFAULT_EXPORT_PDF,
    DEFAULT_OVERWRITE_BAR_CODES,
    BAR_EMPTY_PLACEHOLDER,
    DEFAULT_PDF_DIR,
    DEFAULT_TEMPLATE_PATH,
    GEN_HEADER_BG,
    GEN_HEADER_FG,
    HEADER_ROW,
    QR_BORDER,
    QR_BORDER_COLOR,
    QR_BORDER_INNER_HMM,
    QR_BORDER_OUTER_HMM,
    QR_FONT_INV_PT,
    QR_FONT_NAME,
    QR_FONT_NAME_PT,
    QR_FONT_SERIAL_PT,
    QR_MAPPING_FILE,
    QR_MAPPING_JSON_VER,
    QR_MAPPING_SUBDIR,
    QR_MARGIN_MM,
    QR_PRINT_COL_A_MM,
    QR_PRINT_COL_B_MM,
    QR_PRINT_ROW1_MM,
    QR_PRINT_ROW2_MM,
    QR_PRINT_ROW3_MM,
    QR_SCALE,
    SHEET_GENERATOR,
    SHEET_PRINT_TEMPLATE,
)
from libre_macros_qr_template_blob import (
    QR_TEMPLATE_DEFAULT_SHEET,
    write_qr_template_ods,
)

try:
    unicode
except NameError:
    unicode = str

try:
    unichr
except NameError:
    unichr = chr

_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
_A1_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


def _log(msg):
    try:
        print("[generate_qr_codes] %s" % msg)
    except Exception:
        pass


def _default_mapping():
    return {
        "inv": unicode(DEFAULT_MAP_INV),
        "name": unicode(DEFAULT_MAP_NAME),
        "serial": unicode(DEFAULT_MAP_SERIAL),
        "qr": unicode(DEFAULT_MAP_QR),
        "size_mm": unicode(DEFAULT_MAP_SIZE_MM or u""),
    }


def _parse_size_mm(raw):
    """
    Размер стороны QR в мм.
    Пусто / 0 / <0 → None (автофит); >0 → float мм.
    Некорректный текст → ValueError.
    """
    s = unicode(raw if raw is not None else u"").strip().replace(u",", u".")
    if s == u"":
        return None
    try:
        val = float(s)
    except ValueError:
        raise ValueError(u"Некорректный размер QR (мм): %s" % s)
    if val <= 0:
        return None
    return val


def _format_size_mm(val):
    if val is None:
        return u""
    try:
        f = float(val)
    except (TypeError, ValueError):
        return u""
    if f <= 0:
        return u""
    if abs(f - int(f)) < 1e-9:
        return unicode(int(f))
    return unicode(f)


def _config_base():
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.environ.get(
            "XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config")
        )
    return os.path.join(base, "libre-macros")


def _mapping_dir():
    path = os.path.join(_config_base(), QR_MAPPING_SUBDIR)
    try:
        os.makedirs(path, exist_ok=True)
    except TypeError:
        if not os.path.isdir(path):
            try:
                os.makedirs(path)
            except OSError:
                pass
    except OSError:
        pass
    return path


def _mapping_path():
    return os.path.join(_mapping_dir(), QR_MAPPING_FILE)


def _normalize_mapping(obj):
    defaults = _default_mapping()
    if not isinstance(obj, dict):
        return defaults
    out = dict(defaults)
    for key in ("inv", "name", "serial", "qr"):
        val = unicode(obj.get(key, u"") or u"").strip().upper()
        if val != u"":
            out[key] = val
    if "size_mm" in obj:
        try:
            out["size_mm"] = _format_size_mm(_parse_size_mm(obj.get("size_mm")))
        except ValueError:
            out["size_mm"] = unicode(obj.get("size_mm") or u"").strip()
    return out


def _default_bar_cols():
    return {
        "bar0": unicode(DEFAULT_BAR_COL0),
        "bar1": unicode(DEFAULT_BAR_COL1),
        "bar2": unicode(DEFAULT_BAR_COL2),
    }


def _normalize_bar_cols(obj):
    out = _default_bar_cols()
    if not isinstance(obj, dict):
        return out
    for key in ("bar0", "bar1", "bar2"):
        val = unicode(obj.get(key, u"") or u"").strip().upper()
        if val != u"":
            out[key] = val
    return out


def _validate_bar_cols(bar_cols):
    bar_cols = _normalize_bar_cols(bar_cols)
    labels = {
        "bar0": u"bar0 (СН1)",
        "bar1": u"bar1 (СН2)",
        "bar2": u"bar2 (СН3)",
    }
    for key in ("bar0", "bar1", "bar2"):
        if _col_letters_to_index(bar_cols.get(key)) is None:
            return (
                None,
                u"Некорректный столбец «%s»: %s"
                % (labels.get(key, key), bar_cols.get(key) or u""),
            )
    return bar_cols, None


def _default_settings():
    return {
        "mapping": _default_mapping(),
        "bar_cols": _default_bar_cols(),
        "overwrite_bar_codes": bool(DEFAULT_OVERWRITE_BAR_CODES),
        "template_path": unicode(DEFAULT_TEMPLATE_PATH or u""),
        "export_pdf": bool(DEFAULT_EXPORT_PDF),
        "pdf_dir": unicode(DEFAULT_PDF_DIR or u""),
    }


def _as_bool(val):
    if isinstance(val, bool):
        return val
    s = unicode(val if val is not None else u"").strip().lower()
    return s in (u"1", u"true", u"yes", u"да", u"on")


def _normalize_settings(obj):
    out = _default_settings()
    if not isinstance(obj, dict):
        return out
    if isinstance(obj.get("mapping"), dict):
        out["mapping"] = _normalize_mapping(obj.get("mapping"))
    elif any(k in obj for k in ("inv", "name", "serial", "qr", "size_mm")):
        # Старый файл: корень = mapping.
        out["mapping"] = _normalize_mapping(obj)
    if "template_path" in obj:
        out["template_path"] = unicode(obj.get("template_path") or u"").strip()
    if "pdf_dir" in obj:
        out["pdf_dir"] = unicode(obj.get("pdf_dir") or u"").strip()
    if "export_pdf" in obj:
        out["export_pdf"] = _as_bool(obj.get("export_pdf"))
    if "bar_cols" in obj:
        out["bar_cols"] = _normalize_bar_cols(obj.get("bar_cols"))
    if "overwrite_bar_codes" in obj:
        out["overwrite_bar_codes"] = _as_bool(obj.get("overwrite_bar_codes"))
    return out


def _load_settings():
    path = _mapping_path()
    if not os.path.isfile(path):
        return _default_settings()
    try:
        with open(path, "rb") as f:
            raw = f.read()
        data = json.loads(raw.decode("utf-8"))
    except Exception as err:
        _log("settings read failed: %s" % err)
        return _default_settings()
    return _normalize_settings(data)


def _save_settings(settings):
    s = _normalize_settings(settings if isinstance(settings, dict) else {})
    payload = {
        "version": int(QR_MAPPING_JSON_VER),
        "mapping": s["mapping"],
        "bar_cols": s.get("bar_cols") or _default_bar_cols(),
        "overwrite_bar_codes": bool(s.get("overwrite_bar_codes")),
        "template_path": s.get("template_path") or u"",
        "export_pdf": bool(s.get("export_pdf")),
        "pdf_dir": s.get("pdf_dir") or u"",
    }
    path = _mapping_path()
    try:
        _mapping_dir()
        text = json.dumps(payload, ensure_ascii=False, indent=2)
        if not isinstance(text, bytes):
            text = text.encode("utf-8")
        with open(path, "wb") as f:
            f.write(text)
        _log("settings saved: %s" % path)
        return None
    except Exception as err:
        _log("settings save failed: %s" % err)
        return unicode(err)


def _load_mapping():
    return _load_settings()["mapping"]


def _save_mapping(mapping):
    s = _load_settings()
    s["mapping"] = _normalize_mapping(mapping)
    return _save_settings(s)


def _col_letters_to_index(letters):
    letters = unicode(letters or u"").strip().upper()
    n = 0
    i = 0
    while i < len(letters):
        ch = letters[i]
        if ch < u"A" or ch > u"Z":
            return None
        n = n * 26 + (ord(ch) - 64)
        i = i + 1
    if n <= 0:
        return None
    return n - 1


def _parse_a1(addr):
    """A1 → (col0, row0) или None."""
    s = unicode(addr or u"").strip().upper().replace(u"$", u"")
    m = _A1_RE.match(s)
    if not m:
        return None
    col = _col_letters_to_index(m.group(1))
    if col is None:
        return None
    try:
        row = int(m.group(2)) - 1
    except ValueError:
        return None
    if row < 0:
        return None
    return (col, row)


def _validate_mapping(mapping):
    mapping = _normalize_mapping(mapping)
    labels = {
        "inv": u"Инвентарный номер",
        "name": u"Название оборудования",
        "serial": u"Серийный номер",
        "qr": u"QR код",
    }
    for key in ("inv", "name", "serial", "qr"):
        if _parse_a1(mapping.get(key)) is None:
            return (
                None,
                u"Некорректный адрес «%s»: %s"
                % (labels.get(key, key), mapping.get(key) or u""),
            )
    try:
        size_mm = _parse_size_mm(mapping.get("size_mm"))
    except ValueError as err:
        return None, unicode(err)
    mapping["size_mm"] = _format_size_mm(size_mm)
    return mapping, None


def _script_context():
    try:
        import __main__

        return getattr(__main__, "XSCRIPTCONTEXT", None)
    except Exception:
        return None


def _uno_context():
    xsc = _script_context()
    if xsc is not None:
        try:
            return xsc.getComponentContext()
        except Exception:
            pass
    try:
        return uno.getComponentContext()
    except Exception:
        return None


def _desktop():
    xsc = _script_context()
    if xsc is not None:
        try:
            return xsc.getDesktop()
        except Exception:
            pass
    ctx = _uno_context()
    if ctx is None:
        return None
    try:
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    except Exception:
        return None


def _active_document():
    xsc = _script_context()
    if xsc is not None:
        try:
            doc = xsc.getDocument()
            if doc is not None:
                return doc
        except Exception:
            pass
    desktop = _desktop()
    if desktop is None:
        return None
    try:
        doc = desktop.getCurrentComponent()
        if doc is not None and hasattr(doc, "getSheets"):
            return doc
    except Exception:
        pass
    return None


def _doc_is_readonly(doc):
    """True, если книга открыта только для чтения (нельзя писать в ячейки)."""
    if doc is None:
        return False
    try:
        if hasattr(doc, "isReadonly") and bool(doc.isReadonly()):
            return True
    except Exception:
        pass
    try:
        if bool(getattr(doc, "IsReadOnly", False)):
            return True
    except Exception:
        pass
    return False


def cell_text(cell):
    """Текст ячейки Calc: getString/String (trim) или Value; пустая → \"\"."""
    try:
        if hasattr(cell, "getString"):
            s = cell.getString()
            if s is not None:
                s = unicode(s).strip()
                if s:
                    return s
    except Exception:
        pass
    try:
        s = cell.String
        if s:
            s_stripped = unicode(s).strip()
            if s_stripped:
                return s_stripped
    except Exception:
        pass
    try:
        v = cell.Value
        if v not in (None, 0.0, 0):
            if v == int(v):
                return unicode(int(v))
            return unicode(v)
    except Exception:
        pass
    return u""


def _substr(text, start, length):
    """ПСТР/MID Excel: start с 1."""
    s = unicode(text or u"")
    if start < 1 or length <= 0:
        return u""
    i = int(start) - 1
    return s[i : i + int(length)]


def _is_error(val):
    return unicode(val or u"").strip().casefold() == u"ошибка"


def _excel_char_add(serial, pos, offset):
    ch = _substr(serial, pos, 1)
    if ch == u"":
        return u""
    try:
        n = int(ch)
    except ValueError:
        try:
            n = int(float(ch))
        except ValueError:
            n = ord(ch)
    return unichr(int(n) + int(offset))


def _first_digit_val(serial):
    ch = _substr(serial, 1, 1)
    if ch.isdigit():
        return int(ch)
    try:
        return int(float(ch))
    except ValueError:
        return None


def calc_length(inv):
    return len(unicode(inv or u""))


def calc_country(inv, length):
    a = unicode(inv or u"")
    if a == u"":
        return u"ошибка"
    if _substr(a, 1, 1) == u"A":
        mid = u"4" if length == 9 else u"5"
        return u"3" + mid + u"0"
    if length in (12, 11, 10):
        if _substr(a, 1, 2) == u"МЦ":
            return u"200"
        if _substr(a, 1, 1) == u"C":
            return u"100"
        if _substr(a, 1, 1) == u"A":
            return u"300"
        return u"666"
    if length == 8:
        return u"400"
    if length == 19:
        return u"6" + _substr(a, 19, 1) + u"0"
    return u"900"


def calc_manufacturer(inv, length):
    """Изготовитель (колонка K); сравнение длины — по I, не по E."""
    a = unicode(inv or u"")
    if _substr(a, 1, 1) == u"A":
        return u"00" + _substr(a, 3, 2)
    if length in (12, 11, 10):
        if _substr(a, 1, 2) == u"МЦ":
            prefix = (
                u"0" + _substr(a, 7, 1) if length == 12 else u"00"
            )
            return prefix + _substr(a, 4, 2)
        if _substr(a, 1, 1) == u"C":
            return u"00" + _substr(a, 3, 2)
        return u""
    if length == 8:
        return u"00" + _substr(a, 1, 2)
    if length == 19:
        return u"00" + _substr(a, 1, 2)
    return u"00"


def calc_product(inv, length):
    a = unicode(inv or u"")
    if _substr(a, 1, 1) == u"A":
        return _substr(a, 6, 6)
    if length in (12, 11, 10):
        return _substr(a, 7 + (length - 11), 6)
    if length in (8, 19):
        return _substr(a, 4, 5)
    return u""


def calc_product2(serial3):
    s = unicode(serial3 or u"").strip()
    if s == u"":
        return u""
    try:
        return unicode(int(float(s)) + 100000)
    except ValueError:
        return u""


def calc_shk(country, manufacturer, inv, type_digit):
    """шK0/1/2: LEFT(страна,2) + тип(0/1/2) + изготовитель + последние 5 инв."""
    a = unicode(inv or u"")
    ln = len(a)
    tail = _substr(a, ln - 4, 5) if ln >= 5 else a
    digit = unicode(type_digit or u"")[:1]
    if digit == u"" or not digit.isdigit():
        digit = u"0"
    return (
        unicode(country or u"")[:2]
        + digit
        + unicode(manufacturer or u"")
        + tail
    )


def ean_check_digit(core12):
    s = unicode(core12 or u"")
    if len(s) != 12 or not s.isdigit():
        return u""
    total = 0
    i = 0
    while i < 12:
        d = int(s[i])
        total = total + d * (1 if i % 2 == 0 else 3)
        i = i + 1
    rem = total % 10
    check = (10 - rem) % 10
    return unicode(check)


def calc_bar(shk, country, manufacturer):
    if _is_error(country) or _is_error(manufacturer):
        return u""
    core = unicode(shk or u"")
    if len(core) != 12 or not core.isdigit():
        return u""
    digit = ean_check_digit(core)
    if digit == u"":
        return u""
    return core + digit


def calc_shk_s1(serial):
    s = unicode(serial or u"")
    if s == u"":
        return u""
    parts = [u"*"]
    pos = 8
    while pos <= 13:
        parts.append(_excel_char_add(s, pos, 97))
        pos = pos + 1
    parts.append(u"+")
    return u"".join(parts)


def calc_shk_s2(serial, tovar2):
    s = unicode(serial or u"")
    if s == u"":
        return u""
    fv = _first_digit_val(s)
    if fv is None:
        return u""
    try:
        parts = [_substr(s, 1, 1)]
        parts.append(_excel_char_add(s, 2, 65))
        parts.append(_excel_char_add(s, 3, 75 if fv > 3 else 65))
        parts.append(_excel_char_add(s, 4, 75 if fv in (1, 2, 3, 5, 6, 9) else 65))
        parts.append(_excel_char_add(s, 5, 75 if fv in (2, 3, 6, 7, 8) else 65))
        parts.append(_excel_char_add(s, 6, 75 if fv in (1, 3, 4, 8, 9) else 65))
        parts.append(_excel_char_add(s, 7, 75 if fv in (1, 2, 4, 5, 7) else 65))
        parts.append(unicode(tovar2 or u""))
        return u"".join(parts)
    except Exception:
        return u""


def _set_cell_text(sheet, col, row, value):
    try:
        sheet.getCellByPosition(int(col), int(row)).setString(unicode(value or u""))
    except Exception:
        pass


def _compute_row(inv, name, sn1, sn2, sn3):
    length = calc_length(inv)
    country = calc_country(inv, length)
    manufacturer = calc_manufacturer(inv, length)
    product = calc_product(inv, length)
    product2 = calc_product2(sn3)
    shk0 = calc_shk(country, manufacturer, inv, u"0")
    shk1 = calc_shk(country, manufacturer, inv, u"1")
    shk2 = calc_shk(country, manufacturer, inv, u"2")
    # bar только для непустых серийных номеров
    bar0 = calc_bar(shk0, country, manufacturer) if sn1 != u"" else u""
    bar1 = calc_bar(shk1, country, manufacturer) if sn2 != u"" else u""
    bar2 = calc_bar(shk2, country, manufacturer) if sn3 != u"" else u""
    return {
        "inv": inv,
        "name": name,
        "sn1": sn1,
        "sn2": sn2,
        "sn3": sn3,
        "length": length,
        "country": country,
        "manufacturer": manufacturer,
        "product": product,
        "product2": product2,
        "shk0": shk0,
        "shk1": shk1,
        "shk2": shk2,
        "bar0": bar0,
        "bar1": bar1,
        "bar2": bar2,
        "shk1_s1": calc_shk_s1(sn1),
        "shk1_s2": calc_shk_s2(sn1, product2),
        "shk2_s1": calc_shk_s1(sn2),
        "shk2_s2": calc_shk_s2(sn2, product2),
        "shk3_s1": calc_shk_s1(sn3),
        "shk3_s2": calc_shk_s2(sn3, product2),
    }


def _write_bars_to_sheet(sheet, row, data, bar_cols=None):
    """На лист пишем bar0..bar2 в настроенные столбцы."""
    bc, _err = _validate_bar_cols(bar_cols or _default_bar_cols())
    cols = (
        _col_letters_to_index(bc["bar0"]),
        _col_letters_to_index(bc["bar1"]),
        _col_letters_to_index(bc["bar2"]),
    )
    keys = ("bar0", "bar1", "bar2")
    i = 0
    while i < 3:
        if cols[i] is not None:
            _set_cell_text(sheet, cols[i], row, data.get(keys[i]) or u"")
        i = i + 1


def _write_all_bars(sheet, rows, bar_cols=None):
    for data in rows:
        row = data.get("row")
        if row is None:
            continue
        _write_bars_to_sheet(sheet, row, data, bar_cols=bar_cols)


def _read_bar_cell(sheet, col_idx, row, fallback=u"", use_computed_fallback=True):
    if col_idx is None:
        if use_computed_fallback:
            return fallback or u""
        return BAR_EMPTY_PLACEHOLDER
    try:
        val = cell_text(sheet.getCellByPosition(int(col_idx), int(row))).strip()
    except Exception:
        val = u""
    if val != u"":
        return val
    if use_computed_fallback:
        return fallback or u""
    return BAR_EMPTY_PLACEHOLDER


def _refresh_cards_from_sheet(sheet, rows, bar_cols=None, overwrite_bar_codes=True):
    """Строки QR — из настроенных столбцов листа."""
    bc, err = _validate_bar_cols(bar_cols or _default_bar_cols())
    if err is not None:
        bc = _default_bar_cols()
    cols = (
        _col_letters_to_index(bc["bar0"]),
        _col_letters_to_index(bc["bar1"]),
        _col_letters_to_index(bc["bar2"]),
    )
    use_fallback = bool(overwrite_bar_codes)
    for data in rows:
        row = data.get("row")
        if row is None:
            continue
        cards = []
        sn1 = data.get("sn1") or u""
        sn2 = data.get("sn2") or u""
        sn3 = data.get("sn3") or u""
        if sn1 != u"":
            cards.append(
                (
                    sn1,
                    _read_bar_cell(
                        sheet,
                        cols[0],
                        row,
                        data.get("bar0"),
                        use_computed_fallback=use_fallback,
                    ),
                )
            )
        if sn2 != u"":
            cards.append(
                (
                    sn2,
                    _read_bar_cell(
                        sheet,
                        cols[1],
                        row,
                        data.get("bar1"),
                        use_computed_fallback=use_fallback,
                    ),
                )
            )
        if sn3 != u"":
            cards.append(
                (
                    sn3,
                    _read_bar_cell(
                        sheet,
                        cols[2],
                        row,
                        data.get("bar2"),
                        use_computed_fallback=use_fallback,
                    ),
                )
            )
        data["cards"] = cards


def _sheet_last_data_row(sheet):
    row = DATA_START_ROW
    last = DATA_START_ROW - 1
    while True:
        inv = cell_text(sheet.getCellByPosition(COL_INV, row)).strip()
        if inv == u"":
            break
        last = row
        row = row + 1
        if row > DATA_START_ROW + 50000:
            break
    return last


def _sheet_index(doc, sheet):
    if doc is None or sheet is None:
        return None
    try:
        return int(sheet.RangeAddress.Sheet)
    except Exception:
        pass
    try:
        names = doc.Sheets.getElementNames()
        want = unicode(sheet.Name)
        i = 0
        while i < len(names):
            if unicode(names[i]) == want:
                return i
            i = i + 1
    except Exception:
        pass
    return None


def _selection_addresses(sel):
    if sel is None:
        return []
    try:
        addrs = sel.getRangeAddresses()
        if addrs is not None and len(addrs) > 0:
            return list(addrs)
    except Exception:
        pass
    try:
        return [sel.getRangeAddress()]
    except Exception:
        return []


def _is_single_cell_selection(addrs):
    if addrs is None or len(addrs) != 1:
        return False
    a = addrs[0]
    try:
        return (
            int(a.StartRow) == int(a.EndRow)
            and int(a.StartColumn) == int(a.EndColumn)
        )
    except Exception:
        return False


def _selection_row_set(doc, gen_sheet):
    """
    Строки текущего выделения на листе Генератор_QR, либо None (= весь диапазон).
    Одна ячейка выделением не считается → None.
    """
    if doc is None or gen_sheet is None:
        return None
    try:
        ctrl = doc.getCurrentController()
        if ctrl is None:
            return None
        sel = ctrl.getSelection()
    except Exception:
        return None
    addrs = _selection_addresses(sel)
    if not addrs:
        return None
    if _is_single_cell_selection(addrs):
        _log("выделение: одна ячейка → весь диапазон")
        return None
    sheet_idx = _sheet_index(doc, gen_sheet)
    rows = set()
    i = 0
    while i < len(addrs):
        a = addrs[i]
        try:
            if sheet_idx is not None and int(a.Sheet) != int(sheet_idx):
                i = i + 1
                continue
            sr = int(a.StartRow)
            er = int(a.EndRow)
            if sr > er:
                sr, er = er, sr
            r = sr
            while r <= er:
                if r >= DATA_START_ROW:
                    rows.add(r)
                r = r + 1
        except Exception:
            pass
        i = i + 1
    if not rows:
        _log("выделение: нет строк на листе «%s» → весь диапазон" % SHEET_GENERATOR)
        return None
    _log("выделение: строк=%d" % len(rows))
    return rows


def _collect_rows(sheet, row_filter=None):
    last_row = _sheet_last_data_row(sheet)
    rows = []
    row = DATA_START_ROW
    while row <= last_row:
        if row_filter is not None and row not in row_filter:
            row = row + 1
            continue
        inv = cell_text(sheet.getCellByPosition(COL_INV, row)).strip()
        if inv == u"":
            row = row + 1
            continue
        name = cell_text(sheet.getCellByPosition(COL_NAME, row)).strip()
        sn1 = cell_text(sheet.getCellByPosition(COL_SN1, row)).strip()
        sn2 = cell_text(sheet.getCellByPosition(COL_SN2, row)).strip()
        sn3 = cell_text(sheet.getCellByPosition(COL_SN3, row)).strip()
        data = _compute_row(inv, name, sn1, sn2, sn3)
        data["row"] = row
        cards = []
        if sn1 != u"":
            cards.append((sn1, data.get("bar0") or u""))
        if sn2 != u"":
            cards.append((sn2, data.get("bar1") or u""))
        if sn3 != u"":
            cards.append((sn3, data.get("bar2") or u""))
        data["cards"] = cards
        rows.append(data)
        row = row + 1
    return rows


def _count_cards(rows):
    total = 0
    for item in rows:
        total = total + len(item.get("cards") or [])
    return total


def _build_summary(rows, selection_limited=False):
    total_cards = _count_cards(rows)
    total_rows = len(rows)
    lines = [
        u"Будет сформировано карточек: %d" % total_cards,
        u"Строк данных: %d" % total_rows,
    ]
    if selection_limited:
        lines.append(u"Режим: только выделенные строки")
    else:
        lines.append(u"Режим: весь диапазон")
    lines.extend(
        [
            u"",
            u"Шаблон: «%s»" % SHEET_PRINT_TEMPLATE,
            u"Источник: «%s»" % SHEET_GENERATOR,
            u"",
        ]
    )
    if total_cards == 0:
        lines.append(u"Нет строк с непустыми серийными номерами.")
    else:
        shown = 0
        for item in rows:
            inv = item.get("inv") or u""
            for sn, bar in item.get("cards") or []:
                lines.append(u"• %s / %s → %s" % (inv, sn, bar or u"(пустой bar)"))
                shown = shown + 1
                if shown >= 25:
                    rest = total_cards - shown
                    if rest > 0:
                        lines.append(u"… и ещё %d карточек" % rest)
                    return u"\n".join(lines)
    return u"\n".join(lines)


def _get_toolkit():
    desktop = _desktop()
    if desktop is not None:
        try:
            tk = desktop.getToolkit()
            if tk is not None:
                return tk
        except Exception:
            pass
    ctx = _uno_context()
    if ctx is None:
        return None
    try:
        sm = ctx.getServiceManager()
        return sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    except Exception:
        return None


def _dialog_parent(doc=None):
    if doc is None:
        doc = _active_document()
    if doc is not None:
        try:
            return doc.getCurrentController().getFrame().getContainerWindow()
        except Exception:
            pass
    desktop = _desktop()
    if desktop is not None:
        try:
            frame = desktop.getCurrentFrame()
            if frame is not None:
                return frame.getContainerWindow()
        except Exception:
            pass
    toolkit = _get_toolkit()
    if toolkit is not None:
        try:
            return toolkit.getDesktopWindow()
        except Exception:
            pass
    return None


def _create_peer(dlg, toolkit, doc=None):
    if dlg is None or toolkit is None:
        return False
    parents = []
    parent_win = _dialog_parent(doc)
    if parent_win is not None:
        parents.append(parent_win)
        try:
            peer = parent_win.getPeer()
            if peer is not None:
                parents.append(peer)
        except Exception:
            pass
    parents.append(None)
    i = 0
    while i < len(parents):
        try:
            dlg.createPeer(toolkit, parents[i])
            return True
        except Exception:
            pass
        i = i + 1
    return False


def _ask_yes_no_messagebox(doc, summary_text):
    try:
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK_CANCEL
        from com.sun.star.awt.MessageBoxType import QUERYBOX

        toolkit = _get_toolkit()
        if toolkit is None:
            return None
        parent = _dialog_parent(doc)
        box = toolkit.createMessageBox(
            parent,
            QUERYBOX,
            BUTTONS_OK_CANCEL,
            u"Генерация QR-карточек",
            unicode(summary_text),
        )
        try:
            from com.sun.star.awt.MessageBoxResults import OK

            ok = int(box.execute()) == int(OK)
        except Exception:
            ok = int(box.execute()) == 1
        if not ok:
            return None
        return _load_settings()
    except Exception as err:
        _log("MessageBox fallback failed: %s" % err)
        return None


def _dlg_add_fixed(dm, name, label, x, y, w, h):
    m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    m.Name = name
    m.Label = unicode(label)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    try:
        m.Align = 0
    except Exception:
        pass
    try:
        m.VerticalAlign = 1
    except Exception:
        pass
    dm.insertByName(name, m)
    return m


def _dlg_add_edit(dm, name, x, y, w, h, text=u""):
    m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
    m.Name = name
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    try:
        m.Text = unicode(text or u"")
    except Exception:
        pass
    dm.insertByName(name, m)
    return m


def _dlg_add_button(dm, name, label, x, y, w, h):
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.Name = name
    m.Label = unicode(label)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    dm.insertByName(name, m)
    return m


def _dlg_get_text(dialog, name):
    try:
        return unicode(dialog.getControl(name).getText() or u"").strip()
    except Exception:
        try:
            return unicode(dialog.getModel().getByName(name).Text or u"").strip()
        except Exception:
            return u""


def _dlg_set_text(dialog, name, text):
    try:
        dialog.getControl(name).setText(unicode(text or u""))
        return
    except Exception:
        pass
    try:
        dialog.getModel().getByName(name).Text = unicode(text or u"")
    except Exception:
        pass


def _dlg_read_mapping(dialog):
    return {
        "inv": _dlg_get_text(dialog, "MapInvEd"),
        "name": _dlg_get_text(dialog, "MapNameEd"),
        "serial": _dlg_get_text(dialog, "MapSerialEd"),
        "qr": _dlg_get_text(dialog, "MapQrEd"),
        "size_mm": _dlg_get_text(dialog, "MapSizeEd"),
    }


def _dlg_read_bar_cols(dialog):
    return {
        "bar0": _dlg_get_text(dialog, "Bar0Ed"),
        "bar1": _dlg_get_text(dialog, "Bar1Ed"),
        "bar2": _dlg_get_text(dialog, "Bar2Ed"),
    }


def _normalize_user_path(raw):
    s = unicode(raw if raw is not None else u"").strip()
    if s == u"":
        return u""
    if s.lower().startswith(u"file://"):
        try:
            s = uno.fileUrlToSystemPath(s)
        except Exception:
            pass
    s = os.path.expanduser(os.path.expandvars(s))
    if s == u"":
        return u""
    return os.path.normpath(os.path.abspath(s))


def _path_to_file_url(path):
    if not path:
        return None
    try:
        return uno.systemPathToFileUrl(os.path.abspath(path))
    except Exception:
        return None


def _dlg_add_checkbox(dm, name, label, x, y, w, h, checked=False):
    m = dm.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
    m.Name = name
    m.Label = unicode(label)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    try:
        m.State = 1 if checked else 0
    except Exception:
        pass
    dm.insertByName(name, m)
    return m


def _dlg_get_checkbox(dialog, name):
    try:
        st = int(dialog.getControl(name).getState())
        return st == 1
    except Exception:
        try:
            st = int(dialog.getModel().getByName(name).State)
            return st == 1
        except Exception:
            return False


def _dlg_set_checkbox(dialog, name, checked):
    val = 1 if checked else 0
    try:
        dialog.getControl(name).setState(val)
        return
    except Exception:
        pass
    try:
        dialog.getModel().getByName(name).State = val
    except Exception:
        pass


def _dlg_set_enabled(dialog, name, enabled):
    try:
        dialog.getControl(name).setEnabled(bool(enabled))
        return
    except Exception:
        pass
    try:
        dialog.getModel().getByName(name).Enabled = bool(enabled)
    except Exception:
        pass


def _dlg_apply_pdf_fields(dialog, export_pdf):
    for name in ("PdfDirLbl", "PdfDirEd", "PdfDirBtn"):
        _dlg_set_enabled(dialog, name, bool(export_pdf))


def _dlg_apply_readonly_bar_lock(dialog, readonly):
    """
    Книга только для чтения: снять «Перезаписать bar-коды» и заблокировать галочку.
    """
    if not readonly:
        return
    _dlg_set_checkbox(dialog, "OverwriteBarChk", False)
    _dlg_set_enabled(dialog, "OverwriteBarChk", False)


def _dlg_apply_settings(dialog, settings, readonly=False):
    """Заполнить поля диалога из dict настроек."""
    s = _normalize_settings(settings if isinstance(settings, dict) else {})
    mp = s.get("mapping") or _default_mapping()
    for ed_name, key in (
        ("MapInvEd", "inv"),
        ("MapNameEd", "name"),
        ("MapSerialEd", "serial"),
        ("MapQrEd", "qr"),
        ("MapSizeEd", "size_mm"),
    ):
        _dlg_set_text(dialog, ed_name, mp.get(key))
    bc = s.get("bar_cols") or _default_bar_cols()
    for ed_name, key in (
        ("Bar0Ed", "bar0"),
        ("Bar1Ed", "bar1"),
        ("Bar2Ed", "bar2"),
    ):
        _dlg_set_text(dialog, ed_name, bc.get(key))
    _dlg_set_text(dialog, "TemplateEd", s.get("template_path"))
    _dlg_set_checkbox(dialog, "ExportPdfChk", bool(s.get("export_pdf")))
    overwrite = bool(s.get("overwrite_bar_codes")) and not bool(readonly)
    _dlg_set_checkbox(dialog, "OverwriteBarChk", overwrite)
    _dlg_set_text(dialog, "PdfDirEd", s.get("pdf_dir"))
    _dlg_apply_pdf_fields(dialog, bool(s.get("export_pdf")))
    _dlg_apply_readonly_bar_lock(dialog, readonly)


def _dlg_read_form_settings(dialog, overwrite_bar_codes=None):
    """
    Считать настройки с формы.
    overwrite_bar_codes — если задан, подставить вместо галочки
    (нужно при ReadOnly, чтобы не затирать сохранённое предпочтение).
    """
    raw_map = _dlg_read_mapping(dialog)
    ok_map, err = _validate_mapping(raw_map)
    if err is not None:
        return None, err
    ok_bar, err = _validate_bar_cols(_dlg_read_bar_cols(dialog))
    if err is not None:
        return None, err
    template_path = _normalize_user_path(_dlg_get_text(dialog, "TemplateEd"))
    export_pdf = _dlg_get_checkbox(dialog, "ExportPdfChk")
    pdf_dir = _normalize_user_path(_dlg_get_text(dialog, "PdfDirEd"))
    if template_path != u"" and not os.path.isfile(template_path):
        return None, u"Файл шаблона не найден:\n%s" % template_path
    if export_pdf:
        if pdf_dir == u"":
            return None, u"Укажите каталог для сохранения PDF."
        if not os.path.isdir(pdf_dir):
            try:
                os.makedirs(pdf_dir)
            except Exception:
                return None, u"Каталог PDF недоступен:\n%s" % pdf_dir
    if overwrite_bar_codes is None:
        overwrite = _dlg_get_checkbox(dialog, "OverwriteBarChk")
    else:
        overwrite = bool(overwrite_bar_codes)
    return {
        "mapping": ok_map,
        "bar_cols": ok_bar,
        "overwrite_bar_codes": overwrite,
        "template_path": template_path,
        "export_pdf": bool(export_pdf),
        "pdf_dir": pdf_dir,
    }, None


def _pick_template_file(initial_path=None):
    ctx = _uno_context()
    if ctx is None:
        return None
    try:
        from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE

        sm = ctx.getServiceManager()
        picker = sm.createInstanceWithContext(
            "com.sun.star.ui.dialogs.FilePicker", ctx
        )
        picker.initialize((FILEOPEN_SIMPLE,))
        try:
            picker.appendFilter(u"Книги Calc", "*.ods;*.ots;*.xlsx;*.xlsm")
            picker.appendFilter(u"Все файлы", "*.*")
        except Exception:
            pass
        start = _normalize_user_path(initial_path)
        if start == u"" or not os.path.isfile(start):
            start = _normalize_user_path(os.path.expanduser("~"))
        start_url = _path_to_file_url(start if os.path.isfile(start) else os.path.dirname(start) if start else None)
        if start_url is not None:
            try:
                picker.setDisplayDirectory(start_url)
            except Exception:
                pass
            if os.path.isfile(start):
                try:
                    picker.setDefaultName(os.path.basename(start))
                except Exception:
                    pass
        if int(picker.execute()) != 1:
            return None
        urls = picker.getFiles()
        if not urls:
            return None
        return _normalize_user_path(uno.fileUrlToSystemPath(urls[0]))
    except Exception as err:
        _log("template picker failed: %s" % err)
        return None


def _pick_pdf_folder(initial_dir=None):
    ctx = _uno_context()
    if ctx is None:
        return None
    try:
        sm = ctx.getServiceManager()
        picker = sm.createInstanceWithContext(
            "com.sun.star.ui.dialogs.FolderPicker", ctx
        )
        start = _normalize_user_path(initial_dir)
        if start == u"" or not os.path.isdir(start):
            start = os.path.expanduser("~")
        start_url = _path_to_file_url(start)
        if start_url is not None:
            try:
                picker.setDisplayDirectory(start_url)
            except Exception:
                pass
        if int(picker.execute()) != 1:
            return None
        return _normalize_user_path(uno.fileUrlToSystemPath(picker.getDirectory()))
    except Exception as err:
        _log("folder picker failed: %s" % err)
        return None


def _collect_bars_from_rows(rows):
    bars = []
    for item in rows or []:
        for _sn, bar in item.get("cards") or []:
            b = unicode(bar or u"").strip()
            if b != u"":
                bars.append(b)
    return bars


def _safe_filename_part(text):
    s = unicode(text or u"").strip()
    if s == u"":
        return u""
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', u"_", s)
    s = s.strip(u". ")
    return s


def _make_pdf_filename(first_bar, last_bar):
    fb = _safe_filename_part(first_bar) or u"start"
    lb = _safe_filename_part(last_bar) or fb
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return u"%s-%s_%s.pdf" % (fb, lb, stamp)


def _export_doc_to_pdf(doc, path):
    if doc is None or not path:
        return False
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        try:
            os.makedirs(folder)
        except Exception:
            return False
    try:
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        props = (
            _make_prop("FilterName", "calc_pdf_Export"),
            _make_prop("Overwrite", True),
        )
        doc.storeToURL(url, props)
        _log("pdf: %s" % path)
        return True
    except Exception as err:
        _log("pdf export failed: %s" % err)
        return False


def _export_cards_pdf(doc, rows, pdf_dir):
    bars = _collect_bars_from_rows(rows)
    if not bars:
        return None, u"Нет штрихкодов для имени PDF."
    pdf_dir = _normalize_user_path(pdf_dir)
    if pdf_dir == u"":
        return None, u"Не указан каталог PDF."
    if not os.path.isdir(pdf_dir):
        try:
            os.makedirs(pdf_dir)
        except Exception as err:
            return None, u"Каталог PDF недоступен: %s" % err
    fname = _make_pdf_filename(bars[0], bars[-1])
    path = os.path.join(pdf_dir, fname)
    if not _export_doc_to_pdf(doc, path):
        return None, u"Не удалось экспортировать PDF."
    return path, None

def _ask_start_dialog(doc, summary_text):
    """
    Диалог: слева сводка и пути, справа маппинг ячеек карточки.
    Возвращает dict настроек при Старт, None при Отмена.
    """
    try:
        toolkit = _get_toolkit()
        if toolkit is None:
            _log("dialog: toolkit недоступен → MessageBox")
            return _ask_yes_no_messagebox(doc, summary_text)
        ctx = _uno_context()
        if ctx is None:
            _log("dialog: uno context недоступен → MessageBox")
            return _ask_yes_no_messagebox(doc, summary_text)
        sm = ctx.getServiceManager()
        dm = sm.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx
        )
        settings = _load_settings()
        mapping = settings.get("mapping") or _default_mapping()
        bar_cols = settings.get("bar_cols") or _default_bar_cols()
        doc_readonly = _doc_is_readonly(doc)
        saved_overwrite = bool(settings.get("overwrite_bar_codes"))
        m, g, btn_h, btn_w = 8, 6, 18, 96
        left_w = 300
        map_w = 108
        map_ed_w = 48
        bar_w = 104
        bar_ed_w = 36
        gap_mid = 8
        gap_bar = 4
        dw = m + left_w + gap_mid + map_w + gap_bar + bar_w + m
        dm.Title = u"Генерация QR-карточек"
        dm.Width = dw
        dm.PositionX = 60
        dm.PositionY = 50
        text_h = 132
        y = m
        browse_w = 72
        row_h = 14
        ed_h = 16
        field_gap = 3

        summary_m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        summary_m.Name = "SummaryEdit"
        summary_m.PositionX = m
        summary_m.PositionY = y
        summary_m.Width = left_w
        summary_m.Height = text_h
        summary_m.MultiLine = True
        summary_m.ReadOnly = True
        try:
            summary_m.VScroll = True
        except Exception:
            pass
        dm.insertByName("SummaryEdit", summary_m)

        ly = y + text_h + g
        _dlg_add_fixed(
            dm,
            "TemplateLbl",
            u"Шаблон книги (пусто = вшитый ODS 58×32 мм)",
            m,
            ly,
            left_w,
            14,
        )
        ly = ly + 14
        tpl_ed_w = left_w - browse_w - 4
        _dlg_add_edit(
            dm,
            "TemplateEd",
            m,
            ly,
            tpl_ed_w,
            ed_h,
            settings.get("template_path") or u"",
        )
        _dlg_add_button(
            dm, "TemplateBtn", u"Обзор…", m + tpl_ed_w + 4, ly - 1, browse_w, btn_h
        )
        ly = ly + ed_h + field_gap

        _dlg_add_checkbox(
            dm,
            "ExportPdfChk",
            u"Экспорт в PDF",
            m,
            ly,
            left_w,
            16,
            checked=bool(settings.get("export_pdf")),
        )
        ly = ly + 18
        _dlg_add_fixed(
            dm, "PdfDirLbl", u"Каталог PDF", m, ly, left_w, 14
        )
        ly = ly + 14
        _dlg_add_edit(
            dm,
            "PdfDirEd",
            m,
            ly,
            tpl_ed_w,
            16,
            settings.get("pdf_dir") or u"",
        )
        _dlg_add_button(
            dm, "PdfDirBtn", u"Обзор…", m + tpl_ed_w + 4, ly - 1, browse_w, btn_h
        )
        pdf_browse_y = ly - 1
        left_bottom = ly + 16

        map_x = m + left_w + gap_mid
        _dlg_add_fixed(
            dm, "MapTitleLbl", u"Ячейки карточки QR_Print", map_x, y, map_w, 14
        )
        my = y + 16
        field_specs = (
            ("MapInvLbl", u"Инвентарный номер", "MapInvEd", mapping.get("inv")),
            ("MapNameLbl", u"Название оборудования", "MapNameEd", mapping.get("name")),
            ("MapSerialLbl", u"Серийный номер", "MapSerialEd", mapping.get("serial")),
            ("MapQrLbl", u"QR код", "MapQrEd", mapping.get("qr")),
            (
                "MapSizeLbl",
                u"Размер QR, мм (пусто = автофит)",
                "MapSizeEd",
                mapping.get("size_mm") or u"",
            ),
        )
        for lbl_name, lbl_text, ed_name, ed_val in field_specs:
            _dlg_add_fixed(dm, lbl_name, lbl_text, map_x, my, map_w, row_h)
            my = my + row_h
            _dlg_add_edit(dm, ed_name, map_x, my, map_ed_w, ed_h, ed_val)
            my = my + ed_h + field_gap

        map_hint_y = my
        map_save_y = pdf_browse_y
        map_status_y = map_save_y + btn_h + 2

        bar_x = map_x + map_w + gap_bar
        _dlg_add_fixed(
            dm,
            "BarTitleLbl",
            u"bar → QR",
            bar_x,
            y,
            bar_w,
            14,
        )
        by = y + 16
        bar_specs = (
            ("Bar0Lbl", u"СН1", "Bar0Ed", bar_cols.get("bar0")),
            ("Bar1Lbl", u"СН2", "Bar1Ed", bar_cols.get("bar1")),
            ("Bar2Lbl", u"СН3", "Bar2Ed", bar_cols.get("bar2")),
        )
        for lbl_name, lbl_text, ed_name, ed_val in bar_specs:
            _dlg_add_fixed(dm, lbl_name, lbl_text, bar_x, by, 28, row_h)
            _dlg_add_edit(
                dm, ed_name, bar_x + 30, by + row_h, bar_ed_w, ed_h, ed_val
            )
            by = by + row_h + ed_h + field_gap
        by = by + 4
        overwrite_checked = bool(settings.get("overwrite_bar_codes")) and not doc_readonly
        _dlg_add_checkbox(
            dm,
            "OverwriteBarChk",
            u"Перезаписать bar-коды",
            bar_x,
            by,
            bar_w,
            16,
            checked=overwrite_checked,
        )
        if doc_readonly:
            try:
                dm.getByName("OverwriteBarChk").Enabled = False
            except Exception:
                pass
        bar_bottom = by + 18

        _dlg_add_fixed(
            dm,
            "MapHintLbl",
            u"A1-адреса. QR — в т.ч. merge.",
            map_x,
            map_hint_y,
            map_w,
            14,
        )
        save_btn_w = btn_w + 10
        default_btn_w = btn_w + 18
        _dlg_add_button(
            dm, "SaveMapBtn", u"Сохранить", map_x, map_save_y, save_btn_w, btn_h
        )
        _dlg_add_button(
            dm,
            "DefaultMapBtn",
            u"По умолчанию",
            map_x + save_btn_w + 4,
            map_save_y,
            default_btn_w,
            btn_h,
        )
        status_h = 18
        _dlg_add_fixed(
            dm, "MapStatusLbl", u"", map_x, map_status_y, map_w, status_h
        )
        map_bottom = map_status_y + status_h

        y_btn = max(left_bottom, map_bottom, bar_bottom) + g
        _dlg_add_button(dm, "StartBtn", u"Старт", m, y_btn, btn_w + 24, btn_h)
        _dlg_add_button(
            dm, "CancelBtn", u"Отмена", dw - m - btn_w, y_btn, btn_w, btn_h
        )
        dm.Height = y_btn + btn_h + m
        try:
            from libre_macros_ui_theme import apply_soft_gray_theme, try_paint_titlebar

            apply_soft_gray_theme(dm, title_text=dm.Title)
        except Exception:
            pass
        dialog = sm.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", ctx
        )
        dialog.setModel(dm)
        if not _create_peer(dialog, toolkit, doc=doc):
            _log("dialog: createPeer failed → MessageBox")
            return _ask_yes_no_messagebox(doc, summary_text)
        try:
            from libre_macros_ui_theme import try_paint_titlebar

            try_paint_titlebar(dialog)
        except Exception:
            pass
        try:
            dialog.getControl("CancelBtn").getPeer().setZOrder(0)
        except Exception:
            try:
                peer = dialog.getControl("CancelBtn").getPeer()
                if hasattr(peer, "toFront"):
                    peer.toFront()
            except Exception:
                pass
        try:
            dialog.getControl("OverwriteBarChk").getPeer().setZOrder(0)
        except Exception:
            try:
                peer = dialog.getControl("OverwriteBarChk").getPeer()
                if hasattr(peer, "toFront"):
                    peer.toFront()
            except Exception:
                pass
        try:
            dialog.getControl("SummaryEdit").setText(unicode(summary_text))
        except Exception:
            pass
        _dlg_apply_pdf_fields(dialog, _dlg_get_checkbox(dialog, "ExportPdfChk"))
        _dlg_apply_readonly_bar_lock(dialog, doc_readonly)
        state = {"start": False, "settings": None}

        class _ItemHandler(unohelper.Base, XItemListener):
            def disposing(self, event):
                pass

            def itemStateChanged(self, event):
                try:
                    name = str(event.Source.getModel().Name)
                except Exception:
                    return
                if name == "ExportPdfChk":
                    _dlg_apply_pdf_fields(dialog, _dlg_get_checkbox(dialog, name))

        class _Handler(unohelper.Base, XActionListener):
            def disposing(self, event):
                pass

            def actionPerformed(self, event):
                try:
                    name = str(event.Source.getModel().Name)
                except Exception:
                    return
                if name == "TemplateBtn":
                    picked = _pick_template_file(_dlg_get_text(dialog, "TemplateEd"))
                    if picked:
                        _dlg_set_text(dialog, "TemplateEd", picked)
                    return
                if name == "PdfDirBtn":
                    picked = _pick_pdf_folder(_dlg_get_text(dialog, "PdfDirEd"))
                    if picked:
                        _dlg_set_text(dialog, "PdfDirEd", picked)
                    return
                if name == "SaveMapBtn":
                    # При ReadOnly не затираем сохранённое предпочтение overwrite.
                    preserve = saved_overwrite if doc_readonly else None
                    form, err = _dlg_read_form_settings(
                        dialog, overwrite_bar_codes=preserve
                    )
                    if err is not None:
                        _dlg_set_text(dialog, "MapStatusLbl", err)
                        return
                    save_err = _save_settings(form)
                    if save_err is not None:
                        _dlg_set_text(
                            dialog, "MapStatusLbl", u"Ошибка: %s" % save_err
                        )
                    else:
                        _dlg_apply_settings(dialog, form, readonly=doc_readonly)
                        _dlg_set_text(
                            dialog, "MapStatusLbl", u"Сохранено."
                        )
                    return
                if name == "DefaultMapBtn":
                    defaults = _default_settings()
                    _dlg_apply_settings(dialog, defaults, readonly=doc_readonly)
                    save_err = _save_settings(defaults)
                    if save_err is not None:
                        _dlg_set_text(
                            dialog, "MapStatusLbl", u"Ошибка: %s" % save_err
                        )
                    else:
                        _dlg_set_text(
                            dialog, "MapStatusLbl", u"Сброшено к значениям по умолчанию."
                        )
                    return
                if name == "StartBtn":
                    form, err = _dlg_read_form_settings(dialog)
                    if err is not None:
                        _dlg_set_text(dialog, "MapStatusLbl", err)
                        return
                    if doc_readonly:
                        form["overwrite_bar_codes"] = False
                    state["start"] = True
                    state["settings"] = form
                try:
                    dialog.endExecute()
                except Exception:
                    pass

        handler = _Handler()
        item_handler = _ItemHandler()
        for btn_name in (
            "StartBtn",
            "CancelBtn",
            "SaveMapBtn",
            "DefaultMapBtn",
            "TemplateBtn",
            "PdfDirBtn",
        ):
            try:
                dialog.getControl(btn_name).addActionListener(handler)
            except Exception:
                pass
        try:
            dialog.getControl("ExportPdfChk").addItemListener(item_handler)
        except Exception:
            pass
        try:
            dialog.setVisible(True)
        except Exception:
            pass
        dialog.execute()
        try:
            dialog.dispose()
        except Exception:
            pass
        if state["start"]:
            return state["settings"] or _load_settings()
        return None
    except Exception as err:
        _log("dialog failed: %s" % err)
        return None


def _show_message(title, text, is_error=False):
    try:
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
        from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX

        toolkit = _get_toolkit()
        if toolkit is None:
            _log("%s: %s" % (title, text))
            return
        parent = _dialog_parent()
        box_type = ERRORBOX if is_error else INFOBOX
        box = toolkit.createMessageBox(
            parent, box_type, BUTTONS_OK, unicode(title), unicode(text)
        )
        box.execute()
    except Exception as err:
        _log("message failed: %s — %s: %s" % (err, title, text))


def _activate_document(doc):
    """Вывести книгу на передний план."""
    if doc is None:
        return False
    frame = None
    try:
        frame = doc.getCurrentController().getFrame()
    except Exception:
        pass
    if frame is None:
        return False
    ok = False
    try:
        win = frame.getContainerWindow()
        if win is not None:
            try:
                win.setVisible(True)
            except Exception:
                pass
            if hasattr(win, "toFront"):
                try:
                    win.toFront()
                    ok = True
                except Exception:
                    pass
    except Exception:
        pass
    try:
        frame.activate()
        ok = True
    except Exception:
        pass
    return ok


def _make_prop(name, value):
    from com.sun.star.beans import PropertyValue

    p = PropertyValue()
    p.Name = str(name)
    p.Value = value
    return p


def _close_doc_discard(doc):
    if doc is None:
        return
    try:
        doc.close(True)
    except Exception:
        try:
            doc.dispose()
        except Exception:
            pass


def _temp_qr_result_path():
    """Путь во временной папке ОС (Linux/Windows) через tempfile.gettempdir()."""
    fd, path = tempfile.mkstemp(
        prefix="libre_macros_qr_cards_",
        suffix=".ods",
        dir=tempfile.gettempdir(),
    )
    try:
        os.close(fd)
    except Exception:
        pass
    return path


def _save_doc_to_path(doc, path):
    if doc is None or not path:
        return False
    try:
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        props = (
            _make_prop("FilterName", "calc8"),
            _make_prop("Overwrite", True),
        )
        doc.storeToURL(url, props)
        _log("сохранено: %s" % path)
        return True
    except Exception as err:
        _log("storeToURL failed: %s" % err)
        return False


def _open_doc_visible(path):
    desktop = _desktop()
    if desktop is None:
        return None
    try:
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        return desktop.loadComponentFromURL(url, "_blank", 0, ())
    except Exception as err:
        _log("open failed: %s" % err)
        return None


def _reveal_result_from_hidden(src_doc, hidden_doc):
    """
    Скрытую книгу сохранить во temp, закрыть, открыть видимо, активировать.
    Возвращает открытый документ или None.
    """
    if hidden_doc is None:
        return None
    path = _temp_qr_result_path()
    if not _save_doc_to_path(hidden_doc, path):
        _show_message(
            u"QR-коды",
            u"Не удалось сохранить результат во временный файл:\n%s" % path,
            is_error=True,
        )
        _close_doc_discard(hidden_doc)
        _activate_document(src_doc)
        return None
    _close_doc_discard(hidden_doc)
    opened = _open_doc_visible(path)
    if opened is None:
        _show_message(
            u"QR-коды",
            u"Файл сохранён, но не удалось открыть:\n%s" % path,
            is_error=True,
        )
        _activate_document(src_doc)
        return None
    _activate_document(opened)
    _log("результат открыт: %s" % path)
    return opened


def _show_done_dialog(src_doc, result_doc, text):
    """
    Финальный soft-gray диалог «Готово».
    Книга результата Hidden до OK; по OK — save в temp → open → activate.
    """
    _activate_document(src_doc)
    try:
        toolkit = _get_toolkit()
        ctx = _uno_context()
        if toolkit is None or ctx is None:
            _show_message(u"QR-коды", text)
            return _reveal_result_from_hidden(src_doc, result_doc)
        sm = ctx.getServiceManager()
        dm = sm.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx
        )
        m, g, btn_h, btn_w = 8, 6, 18, 96
        dw = 420
        dm.Title = u"QR-коды"
        dm.Width = dw
        dm.PositionX = 100
        dm.PositionY = 80
        content_w = dw - 2 * m
        text_h = 72
        y = m
        summary_m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
        summary_m.Name = "DoneEdit"
        summary_m.PositionX = m
        summary_m.PositionY = y
        summary_m.Width = content_w
        summary_m.Height = text_h
        summary_m.MultiLine = True
        summary_m.ReadOnly = True
        try:
            summary_m.VScroll = True
        except Exception:
            pass
        dm.insertByName("DoneEdit", summary_m)
        y = y + text_h + g
        _dlg_add_button(
            dm, "OkBtn", u"OK", (dw - btn_w) // 2, y, btn_w, btn_h
        )
        dm.Height = y + btn_h + m
        try:
            from libre_macros_ui_theme import apply_soft_gray_theme, try_paint_titlebar

            apply_soft_gray_theme(dm, title_text=dm.Title)
        except Exception:
            pass
        dialog = sm.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", ctx
        )
        dialog.setModel(dm)
        if not _create_peer(dialog, toolkit, doc=src_doc):
            _show_message(u"QR-коды", text)
            return _reveal_result_from_hidden(src_doc, result_doc)
        try:
            from libre_macros_ui_theme import try_paint_titlebar

            try_paint_titlebar(dialog)
        except Exception:
            pass
        try:
            dialog.getControl("DoneEdit").setText(unicode(text))
        except Exception:
            pass

        class _Handler(unohelper.Base, XActionListener):
            def disposing(self, event):
                pass

            def actionPerformed(self, event):
                try:
                    dialog.endExecute()
                except Exception:
                    pass

        try:
            dialog.getControl("OkBtn").addActionListener(_Handler())
        except Exception:
            pass
        try:
            dialog.setVisible(True)
        except Exception:
            pass
        dialog.execute()
        try:
            dialog.dispose()
        except Exception:
            pass
    except Exception as err:
        _log("done dialog failed: %s" % err)
        _show_message(u"QR-коды", text)
    return _reveal_result_from_hidden(src_doc, result_doc)


def _get_sheet(doc, name):
    if doc is None:
        return None
    target = unicode(name or u"")
    try:
        return doc.Sheets.getByName(target)
    except Exception:
        pass
    try:
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            sh = sheets.getByIndex(i)
            if unicode(sh.Name) == target:
                return sh
            i = i + 1
    except Exception:
        pass
    try:
        want = target.casefold()
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            sh = sheets.getByIndex(i)
            if unicode(sh.Name).casefold() == want:
                return sh
            i = i + 1
    except Exception:
        pass
    return None


def _sanitize_sheet_name(base):
    name = _INVALID_SHEET_CHARS.sub(u"_", unicode(base or u"").strip())
    name = name.strip(u".'")
    if name == u"":
        name = u"QR"
    if len(name) > 31:
        name = name[:31]
    return name


def _unique_sheet_name(doc, base):
    name = _sanitize_sheet_name(base)
    if not doc.Sheets.hasByName(name):
        return name
    n = 2
    while n < 10000:
        suffix = u"_%d" % n
        candidate = (name[: 31 - len(suffix)] + suffix)[:31]
        if not doc.Sheets.hasByName(candidate):
            return candidate
        n = n + 1
    return u"QR_%d" % n


def _import_sheet_between_docs(src_doc, src_name, dst_doc, dst_name):
    try:
        t = uno.getTypeByName("com.sun.star.sheet.XSpreadsheets2")
        sheets2 = dst_doc.Sheets.queryInterface(t)
        if not sheets2:
            return None
        insert_index = int(dst_doc.Sheets.getCount())
        sheets2.importSheet(src_doc, unicode(src_name), insert_index)
        sheet = _get_sheet(dst_doc, src_name)
        if sheet is None:
            return None
        if unicode(dst_name) != unicode(src_name):
            sheet.Name = unicode(dst_name)
            sheet = _get_sheet(dst_doc, dst_name)
        return sheet
    except Exception as err:
        _log("importSheet failed: %s" % err)
        return None


def _copy_sheet_in_doc(doc, src_name, dst_name):
    try:
        pos = int(doc.Sheets.getCount())
        doc.Sheets.copyByName(unicode(src_name), unicode(dst_name), pos)
        return _get_sheet(doc, dst_name)
    except Exception as err:
        _log("copyByName failed: %s" % err)
        return None


def _remove_other_sheets(doc, keep_names):
    keep = set(unicode(n) for n in keep_names)
    names = []
    try:
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            names.append(sheets.getByIndex(i).Name)
            i = i + 1
    except Exception:
        return
    for name in names:
        if name in keep:
            continue
        try:
            doc.Sheets.removeByName(name)
        except Exception:
            pass


def _merged_range_address(sheet, col, row):
    """
    Адрес диапазона с учётом объединения ячеек.
    Возвращает (sc, sr, ec, er) 0-based или (col, row, col, row).
    """
    sc = int(col)
    sr = int(row)
    ec = sc
    er = sr
    try:
        cell = sheet.getCellByPosition(sc, sr)
        cursor = sheet.createCursorByRange(cell)
        cursor.collapseToMergedArea()
        addr = cursor.getRangeAddress()
        sc = int(addr.StartColumn)
        sr = int(addr.StartRow)
        ec = int(addr.EndColumn)
        er = int(addr.EndRow)
    except Exception:
        pass
    return (sc, sr, ec, er)


def _sum_col_widths(sheet, sc, ec):
    cols = sheet.getColumns()
    total = 0
    c = int(sc)
    while c <= int(ec):
        try:
            total = total + int(cols.getByIndex(c).Width)
        except Exception:
            pass
        c = c + 1
    return max(100, total)


def _sum_row_heights(sheet, sr, er):
    rows = sheet.getRows()
    total = 0
    r = int(sr)
    while r <= int(er):
        try:
            total = total + int(rows.getByIndex(r).Height)
        except Exception:
            pass
        r = r + 1
    return max(100, total)


def _cell_rect(sheet, col, row):
    """
    Прямоугольник ячейки/объединённого диапазона в 1/100 мм: (x, y, width, height).
    """
    from com.sun.star.awt import Point, Size

    sc, sr, ec, er = _merged_range_address(sheet, col, row)
    rng = sheet.getCellRangeByPosition(sc, sr, ec, er)
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
        while c < sc:
            x = x + int(cols.getByIndex(c).Width)
            c = c + 1
        r = 0
        while r < sr:
            y = y + int(rows.getByIndex(r).Height)
            r = r + 1
        pos = Point(x, y)
    if size is None:
        size = Size(_sum_col_widths(sheet, sc, ec), _sum_row_heights(sheet, sr, er))
    else:
        # На AO/LO Size одиночной ячейки иногда остаётся — пересчитать по merge.
        if sc != ec or sr != er:
            w = _sum_col_widths(sheet, sc, ec)
            h = _sum_row_heights(sheet, sr, er)
            if int(size.Width) < w or int(size.Height) < h:
                size = Size(w, h)
    return (int(pos.X), int(pos.Y), int(size.Width), int(size.Height))


def _fit_centered_in_cell(cell_x, cell_y, cell_w, cell_h, margin_mm=0, aspect=1.0):
    margin = max(0, int(round(margin_mm * 100)))
    inner_w = max(100, cell_w - 2 * margin)
    inner_h = max(100, cell_h - 2 * margin)
    if aspect <= 0:
        aspect = 1.0
    if inner_w / float(aspect) <= inner_h:
        w = inner_w
        h = max(100, int(round(w / float(aspect))))
    else:
        h = inner_h
        w = max(100, int(round(h * float(aspect))))
    x = cell_x + (cell_w - w) // 2
    y = cell_y + (cell_h - h) // 2
    return (x, y, w, h)


def _clear_sheet_pictures(sheet):
    try:
        draw_page = sheet.getDrawPage()
    except Exception:
        return
    j = draw_page.getCount() - 1
    while j >= 0:
        try:
            shape = draw_page.getByIndex(j)
            st = str(shape.getShapeType() or "")
            if "GraphicObject" in st:
                draw_page.remove(shape)
            elif shape.supportsService("com.sun.star.drawing.GraphicObjectShape"):
                draw_page.remove(shape)
        except Exception:
            pass
        j = j - 1


def _insert_png_at_cell(doc, sheet, png_path, col, row, margin_mm=0, size_mm=None):
    ctx = _uno_context()
    if ctx is None:
        raise RuntimeError("UNO context недоступен")
    smgr = ctx.getServiceManager()
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
        pass

    cx, cy, cw, ch = _cell_rect(sheet, col, row)
    if size_mm is not None and float(size_mm) > 0:
        side = max(100, int(round(float(size_mm) * 100)))
        x = cx + (cw - side) // 2
        y = cy + (ch - side) // 2
        w = side
        h = side
    else:
        x, y, w, h = _fit_centered_in_cell(
            cx, cy, cw, ch, margin_mm=margin_mm, aspect=aspect
        )

    shape = doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
    shape.Graphic = graphic
    shape.setPosition(Point(x, y))
    shape.setSize(Size(w, h))
    try:
        shape.Anchor = sheet.getCellByPosition(col, row)
    except Exception:
        pass
    sheet.getDrawPage().add(shape)


def _clear_hf_content(page_style, prop_name):
    """Очистить Left/Center/Right текст в XHeaderFooterContent."""
    try:
        hfc = page_style.getPropertyValue(prop_name)
    except Exception:
        try:
            hfc = getattr(page_style, prop_name, None)
        except Exception:
            hfc = None
    if hfc is None:
        return
    changed = False
    for part in ("LeftText", "CenterText", "RightText"):
        try:
            txt = getattr(hfc, part, None)
            if txt is not None:
                txt.String = u""
                changed = True
        except Exception:
            pass
    if not changed:
        return
    try:
        page_style.setPropertyValue(prop_name, hfc)
    except Exception:
        try:
            setattr(page_style, prop_name, hfc)
        except Exception:
            pass


def _clear_sheet_headers_footers(doc, sheet):
    """Выключить и очистить все колонтитулы печати листа."""
    if doc is None or sheet is None:
        return
    try:
        style_name = unicode(sheet.PageStyle or u"")
    except Exception:
        style_name = u""
    if style_name == u"":
        return
    try:
        page_style = doc.StyleFamilies.getByName("PageStyles").getByName(style_name)
    except Exception:
        return
    for prop, val in (
        ("HeaderIsOn", False),
        ("FooterIsOn", False),
        ("FirstPageHeaderIsOn", False),
        ("FirstPageFooterIsOn", False),
    ):
        try:
            page_style.setPropertyValue(prop, val)
        except Exception:
            try:
                setattr(page_style, prop, val)
            except Exception:
                pass
    for prop in (
        "LeftPageHeaderContent",
        "CenterPageHeaderContent",
        "RightPageHeaderContent",
        "LeftPageFooterContent",
        "CenterPageFooterContent",
        "RightPageFooterContent",
        "FirstPageHeaderContent",
        "FirstPageFooterContent",
    ):
        _clear_hf_content(page_style, prop)
    for prop in (
        "HeaderText",
        "HeaderTextLeft",
        "HeaderTextRight",
        "FooterText",
        "FooterTextLeft",
        "FooterTextRight",
    ):
        try:
            page_style.setPropertyValue(prop, u"")
        except Exception:
            try:
                setattr(page_style, prop, u"")
            except Exception:
                pass


def _fill_card_sheet(doc, sheet, inv, name, serial, bar_text, mapping=None):
    mapping, err = _validate_mapping(mapping or _default_mapping())
    if err is not None:
        raise ValueError(err)
    _clear_sheet_headers_footers(doc, sheet)
    inv_pos = _parse_a1(mapping["inv"])
    name_pos = _parse_a1(mapping["name"])
    sn_pos = _parse_a1(mapping["serial"])
    qr_pos = _parse_a1(mapping["qr"])
    size_mm = _parse_size_mm(mapping.get("size_mm"))
    _set_cell_text(sheet, inv_pos[0], inv_pos[1], inv)
    _set_cell_text(sheet, name_pos[0], name_pos[1], name)
    _set_cell_text(sheet, sn_pos[0], sn_pos[1], serial)
    try:
        _set_cell_alignment(
            sheet.getCellByPosition(inv_pos[0], inv_pos[1]), u"center", u"center"
        )
        _set_cell_alignment(
            sheet.getCellByPosition(name_pos[0], name_pos[1]), u"center", u"top"
        )
        _set_cell_alignment(
            sheet.getCellByPosition(sn_pos[0], sn_pos[1]), u"center", u"center"
        )
        _set_cell_alignment(
            sheet.getCellByPosition(qr_pos[0], qr_pos[1]), u"center", u"center"
        )
        sheet.getCellByPosition(name_pos[0], name_pos[1]).IsTextWrapped = True
        _set_cell_font(
            sheet.getCellByPosition(inv_pos[0], inv_pos[1]), QR_FONT_INV_PT, True
        )
        _set_cell_font(
            sheet.getCellByPosition(name_pos[0], name_pos[1]), QR_FONT_NAME_PT, True
        )
        _set_cell_font(
            sheet.getCellByPosition(sn_pos[0], sn_pos[1]), QR_FONT_SERIAL_PT, True
        )
    except Exception:
        pass
    try:
        _apply_card_table_borders(sheet)
        _apply_qr_print_page_style(doc, sheet)
    except Exception:
        pass
    _clear_sheet_pictures(sheet)
    bar = unicode(bar_text or u"").strip()
    if bar == u"":
        return
    fd, png_path = tempfile.mkstemp(prefix="libre_macros_qr_", suffix=".png")
    os.close(fd)
    try:
        qr = segno.make_qr(bar, error="m")
        qr.save(png_path, scale=QR_SCALE, border=QR_BORDER)
        _insert_png_at_cell(
            doc,
            sheet,
            png_path,
            qr_pos[0],
            qr_pos[1],
            margin_mm=QR_MARGIN_MM,
            size_mm=size_mm,
        )
    finally:
        try:
            os.unlink(png_path)
        except Exception:
            pass


def _create_result_workbook(src_doc, template_path=None):
    """
    Книга результата: внешний ODS или вшитый generate_qr_template.ods (58×32, поля 0).
    Без UNO-правок PageStyle (на AO зависает).
    """
    desktop = _desktop()
    if desktop is None:
        return None
    tpl_path = _normalize_user_path(template_path)
    temp_tpl = None
    if tpl_path == u"" or not os.path.isfile(tpl_path):
        fd, tpl_path = tempfile.mkstemp(
            prefix="libre_macros_qr_tpl_", suffix=".ods"
        )
        temp_tpl = tpl_path
        try:
            os.close(fd)
        except Exception:
            pass
        try:
            write_qr_template_ods(tpl_path)
        except Exception as err:
            _log("write template blob: %s" % err)
            try:
                os.unlink(tpl_path)
            except Exception:
                pass
            return None
    else:
        _log("шаблон книги: %s" % tpl_path)
    try:
        props = (_make_prop("Hidden", True),)
    except Exception:
        props = ()
    try:
        url = uno.systemPathToFileUrl(os.path.abspath(tpl_path))
        new_doc = desktop.loadComponentFromURL(url, "_blank", 0, props)
    except Exception as err:
        _log("open template: %s" % err)
        new_doc = None
    if temp_tpl:
        try:
            os.unlink(temp_tpl)
        except Exception:
            pass
    if new_doc is None:
        return None
    _activate_document(src_doc)
    template_name = u"__QR_TEMPLATE__"
    if _import_sheet_between_docs(
        src_doc, SHEET_PRINT_TEMPLATE, new_doc, template_name
    ) is None:
        _close_doc_discard(new_doc)
        _activate_document(src_doc)
        return None
    # Назначить стиль страницы из заготовки (Default = 58×32), без изменения свойств.
    try:
        sh = _get_sheet(new_doc, template_name)
        if sh is not None:
            sh.PageStyle = u"Default"
    except Exception:
        pass
    _activate_document(src_doc)
    return new_doc


def _generate_cards(src_doc, rows, mapping=None, template_path=None):
    mapping, err = _validate_mapping(mapping or _load_mapping())
    if err is not None:
        raise ValueError(err)
    new_doc = _create_result_workbook(src_doc, template_path=template_path)
    if new_doc is None:
        return None, 0
    template_name = u"__QR_TEMPLATE__"
    created = 0
    for item in rows:
        inv = item.get("inv") or u""
        name = item.get("name") or u""
        for serial, bar in item.get("cards") or []:
            base = u"%s_%s" % (inv, serial)
            sheet_name = _unique_sheet_name(new_doc, base)
            sheet = _copy_sheet_in_doc(new_doc, template_name, sheet_name)
            if sheet is None:
                continue
            try:
                sheet.PageStyle = u"Default"
            except Exception:
                pass
            _fill_card_sheet(
                new_doc, sheet, inv, name, serial, bar, mapping=mapping
            )
            created = created + 1
    # Удалить заготовку карточки и дефолтный лист шаблона ODS.
    for drop in (template_name, unicode(QR_TEMPLATE_DEFAULT_SHEET), u"Sheet1"):
        try:
            if new_doc.Sheets.hasByName(drop) and new_doc.Sheets.getCount() > 1:
                new_doc.Sheets.removeByName(drop)
        except Exception:
            pass
    return new_doc, created


def _mm_to_hmm(mm):
    return int(round(float(mm) * 100.0))


def _insert_sheet(doc, name):
    sheets = doc.Sheets
    target = unicode(name)
    if sheets.hasByName(target):
        return sheets.getByName(target)
    sheets.insertNewByName(target, sheets.getCount())
    return sheets.getByName(target)


def _style_header_cell(cell):
    try:
        cell.CellBackColor = int(GEN_HEADER_BG)
    except Exception:
        pass
    try:
        cell.CharColor = int(GEN_HEADER_FG)
    except Exception:
        pass
    try:
        cell.CharWeight = 150
    except Exception:
        pass


def _set_cell_alignment(cell, hori="center", vert="center"):
    """hori: left|center|right; vert: top|center|bottom."""
    if cell is None:
        return
    hori_key = unicode(hori or u"center").strip().casefold()
    vert_key = unicode(vert or u"center").strip().casefold()
    try:
        from com.sun.star.table.CellHoriJustify import (
            CENTER as HJ_CENTER,
            LEFT as HJ_LEFT,
            RIGHT as HJ_RIGHT,
        )

        if hori_key == u"left":
            cell.HoriJustify = HJ_LEFT
        elif hori_key == u"right":
            cell.HoriJustify = HJ_RIGHT
        else:
            cell.HoriJustify = HJ_CENTER
    except Exception:
        try:
            cell.HoriJustify = 2 if hori_key == u"center" else (1 if hori_key == u"left" else 3)
        except Exception:
            pass
    try:
        from com.sun.star.table.CellVertJustify import (
            BOTTOM as VJ_BOTTOM,
            CENTER as VJ_CENTER,
            TOP as VJ_TOP,
        )

        if vert_key == u"top":
            cell.VertJustify = VJ_TOP
        elif vert_key == u"bottom":
            cell.VertJustify = VJ_BOTTOM
        else:
            cell.VertJustify = VJ_CENTER
    except Exception:
        pass
    try:
        from com.sun.star.table.CellVertJustify2 import (
            BOTTOM as VJ2_BOTTOM,
            CENTER as VJ2_CENTER,
            TOP as VJ2_TOP,
        )

        if vert_key == u"top":
            cell.VertJustify2 = VJ2_TOP
        elif vert_key == u"bottom":
            cell.VertJustify2 = VJ2_BOTTOM
        else:
            cell.VertJustify2 = VJ2_CENTER
    except Exception:
        pass


def _create_generator_sheet(doc):
    """Создать «Генератор_QR» с заголовками и фейковыми примерами."""
    sheet = _insert_sheet(doc, SHEET_GENERATOR)
    headers = (
        u"Инвентарный номер",
        u"Название оборудования",
        u"Серийный номер 1",
        u"Серийный номер 2",
        u"Серийный номер 3",
        u"bar0",
        u"bar1",
        u"bar2",
    )
    col = 0
    while col < len(headers):
        cell = sheet.getCellByPosition(col, HEADER_ROW)
        cell.setString(headers[col])
        _style_header_cell(cell)
        col = col + 1

    # Фейковые данные: структура инв.номеров как у реальных (C/…, МЦ/…, A…, 8 символов).
    samples = (
        (
            u"C/08/00123",
            u"Коммутатор DemoSwitch X24 — учебный образец для проверки макета печати",
            u"SNDEMO0001",
            u"",
            u"",
        ),
        (
            u"МЦ/01/100200",
            u"Источник бесперебойного питания DemoUPS-500VA, мощность 0,3 кВт, "
            u"вход 160-275 В, выход 220 В (фейковый пример)",
            u"40010010001",
            u"40010010002",
            u"SNDEMO0099",
        ),
        (
            u"A090011122",
            u"Сервер хранения DemoNAS-8bay (тестовая запись)",
            u"NASDEMO7788",
            u"",
            u"",
        ),
        (
            u"AB12CD34",
            u"Маршрутизатор DemoRouter R1 с двумя серийными номерами",
            u"RTXDEMO001",
            u"RTXDEMO002",
            u"",
        ),
        (
            u"C/11/00999",
            u"Точка доступа DemoAP-WiFi6, потолок/стена, PoE",
            u"APDEMO5500",
            u"",
            u"",
        ),
    )
    r = DATA_START_ROW
    for row_vals in samples:
        c = 0
        while c < len(row_vals):
            sheet.getCellByPosition(c, r).setString(row_vals[c])
            c = c + 1
        r = r + 1

    # Ширина всех значимых колонок 50 мм.
    try:
        i = 0
        while i < len(headers):
            sheet.getColumns().getByIndex(i).Width = _mm_to_hmm(50)
            i = i + 1
    except Exception:
        pass

    # Перенос по словам в значащих колонках A–E на 200 строк (заголовок + данные).
    try:
        rng = sheet.getCellRangeByPosition(0, 0, 4, 199)
        rng.IsTextWrapped = True
    except Exception as err:
        _log("word wrap: %s" % err)
    _log("создан лист «%s» с фейковыми примерами" % SHEET_GENERATOR)
    return sheet


def _create_print_template_sheet(doc):
    """
    Создать «QR_Print»: A=28мм, B=30мм; строки 7/18/7 мм; A1:A3 под QR.
    Рамка A1:B3 — толстый контур, тонкая внутренняя сетка.
    Шрифт PT Sans жирный: B1/B3 10pt, B2 8pt.
    """
    sheet = _insert_sheet(doc, SHEET_PRINT_TEMPLATE)
    _apply_qr_print_layout(doc, sheet, with_hints=True)
    _log("создан шаблон «%s»" % SHEET_PRINT_TEMPLATE)
    return sheet


def _border_line(color, width_hmm):
    bl = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    bl.Color = int(color)
    bl.LineStyle = 0
    w = int(width_hmm)
    bl.LineWidth = w
    bl.InnerLineWidth = w
    bl.OuterLineWidth = w
    return bl


def _apply_card_table_borders(sheet):
    """A1:B3 — толстый контур, тонкие внутренние линии."""
    try:
        rng = sheet.getCellRangeByPosition(0, 0, 1, 2)
        thick = _border_line(QR_BORDER_COLOR, QR_BORDER_OUTER_HMM)
        thin = _border_line(QR_BORDER_COLOR, QR_BORDER_INNER_HMM)
        try:
            tb = rng.TableBorder2
        except Exception:
            tb = rng.TableBorder
        tb.TopLine = thick
        tb.BottomLine = thick
        tb.LeftLine = thick
        tb.RightLine = thick
        tb.HorizontalLine = thin
        tb.VerticalLine = thin
        try:
            tb.IsTopLineValid = True
            tb.IsBottomLineValid = True
            tb.IsLeftLineValid = True
            tb.IsRightLineValid = True
            tb.IsHorizontalLineValid = True
            tb.IsVerticalLineValid = True
        except Exception:
            pass
        try:
            rng.TableBorder2 = tb
        except Exception:
            rng.TableBorder = tb
    except Exception as err:
        _log("borders: %s" % err)


def _set_cell_font(cell, size_pt, bold=True, font_name=None):
    if cell is None:
        return
    try:
        cell.CharFontName = unicode(font_name or QR_FONT_NAME)
    except Exception:
        pass
    try:
        cell.CharHeight = float(size_pt)
    except Exception:
        pass
    if bold:
        try:
            cell.CharWeight = 150
        except Exception:
            pass


def _ensure_qr_page_style(doc):
    """
    Больше не создаём/не правим PageStyle через UNO (зависания AO).
    Бумага 58×32 берётся из вшитого generate_qr_template.ods (стиль Default).
    """
    _unused = doc
    return u"Default"


def _apply_qr_print_page_style(doc, sheet):
    """Только назначить Default из книги-шаблона; без setPropertyValue размера."""
    if sheet is None:
        return
    try:
        styles = doc.StyleFamilies.getByName("PageStyles")
        if styles.hasByName(u"Default"):
            sheet.PageStyle = u"Default"
    except Exception as err:
        _log("assign PageStyle Default: %s" % err)
    _clear_sheet_headers_footers(doc, sheet)


def _apply_qr_print_layout(doc, sheet, with_hints=False):
    """Размеры, merge, рамка, выравнивание, шрифт, стиль страницы."""
    try:
        sheet.getColumns().getByIndex(0).Width = _mm_to_hmm(QR_PRINT_COL_A_MM)
        sheet.getColumns().getByIndex(1).Width = _mm_to_hmm(QR_PRINT_COL_B_MM)
    except Exception as err:
        _log("ширина колонок QR_Print: %s" % err)
    try:
        rows = sheet.getRows()
        rows.getByIndex(0).Height = _mm_to_hmm(QR_PRINT_ROW1_MM)
        rows.getByIndex(1).Height = _mm_to_hmm(QR_PRINT_ROW2_MM)
        rows.getByIndex(2).Height = _mm_to_hmm(QR_PRINT_ROW3_MM)
    except Exception as err:
        _log("высота строк QR_Print: %s" % err)
    try:
        sheet.getCellRangeByPosition(0, 0, 0, 2).merge(True)
    except Exception as err:
        _log("merge A1:A3: %s" % err)
    if with_hints:
        try:
            sheet.getCellByPosition(0, 0).setString(u"QR")
            sheet.getCellByPosition(1, 0).setString(u"Инв. номер")
            sheet.getCellByPosition(1, 1).setString(u"Название")
            sheet.getCellByPosition(1, 2).setString(u"Серийный №")
        except Exception:
            pass
    try:
        _set_cell_alignment(sheet.getCellByPosition(0, 0), u"center", u"center")
        _set_cell_alignment(sheet.getCellByPosition(1, 0), u"center", u"center")
        _set_cell_alignment(sheet.getCellByPosition(1, 1), u"center", u"top")
        _set_cell_alignment(sheet.getCellByPosition(1, 2), u"center", u"center")
        sheet.getCellByPosition(1, 1).IsTextWrapped = True
    except Exception as err:
        _log("alignment QR_Print: %s" % err)
    try:
        _set_cell_font(sheet.getCellByPosition(1, 0), QR_FONT_INV_PT, True)
        _set_cell_font(sheet.getCellByPosition(1, 1), QR_FONT_NAME_PT, True)
        _set_cell_font(sheet.getCellByPosition(1, 2), QR_FONT_SERIAL_PT, True)
    except Exception as err:
        _log("font QR_Print: %s" % err)
    _apply_card_table_borders(sheet)
    _apply_qr_print_page_style(doc, sheet)


def _ensure_required_sheets(doc):
    """
    Создать отсутствующие Генератор_QR / QR_Print.
    Возвращает (gen_sheet, created_list).
    """
    created = []
    gen_sheet = _get_sheet(doc, SHEET_GENERATOR)
    if gen_sheet is None:
        gen_sheet = _create_generator_sheet(doc)
        created.append(SHEET_GENERATOR)
    if _get_sheet(doc, SHEET_PRINT_TEMPLATE) is None:
        _create_print_template_sheet(doc)
        created.append(SHEET_PRINT_TEMPLATE)
    return gen_sheet, created


def _show_setup_message(doc, created):
    lines = [u"Созданы листы: %s" % u", ".join(u"«%s»" % n for n in created), u""]
    if SHEET_GENERATOR in created:
        lines.extend(
            [
                u"Лист «%s»:" % SHEET_GENERATOR,
                u"• колонка A — инвентарный номер;",
                u"• B — название оборудования;",
                u"• C, D, E — серийные номера (1–3, можно пустые);",
                u"• F–H (bar0..2) заполняются макросом автоматически.",
                u"Замените примеры своими данными (со 2-й строки).",
                u"",
            ]
        )
    if SHEET_PRINT_TEMPLATE in created:
        lines.extend(
            [
                u"Лист «%s» — макет карточки:" % SHEET_PRINT_TEMPLATE,
                u"• A=28 мм, B=30 мм; строки 7 / 18 / 7 мм;",
                u"• A1:A3 — QR; B1 инв.номер; B2 название; B3 серийный;",
                u"• печать: книга из шаблона 58×32 мм (вшитый ODS).",
                u"",
            ]
        )
    lines.append(u"Заполните данные и запустите макрос снова.")
    _show_message(u"QR-коды — подготовка", u"\n".join(lines))
    try:
        gen = _get_sheet(doc, SHEET_GENERATOR)
        if gen is not None:
            doc.getCurrentController().setActiveSheet(gen)
    except Exception:
        pass


def generate_qr_codes(*args):
    _unused = args
    _log("старт, segno %s" % getattr(segno, "__version__", "?"))
    try:
        doc = _active_document()
        if doc is None or not hasattr(doc, "getSheets"):
            _log("нет активной книги Calc")
            _show_message(u"QR-коды", u"Откройте книгу Calc.", is_error=True)
            return

        doc_readonly = _doc_is_readonly(doc)
        if doc_readonly:
            _log("книга только для чтения — запись bar0..2 будет пропущена")

        # Если листов нет и книга ReadOnly — создать нельзя.
        if doc_readonly:
            missing = []
            if _get_sheet(doc, SHEET_GENERATOR) is None:
                missing.append(SHEET_GENERATOR)
            if _get_sheet(doc, SHEET_PRINT_TEMPLATE) is None:
                missing.append(SHEET_PRINT_TEMPLATE)
            if missing:
                _show_message(
                    u"QR-коды",
                    u"Книга открыта только для чтения.\n"
                    u"Нельзя создать листы: %s.\n"
                    u"Откройте файл для записи или скопируйте листы заранее."
                    % u", ".join(u"«%s»" % n for n in missing),
                    is_error=True,
                )
                return

        gen_sheet, created = _ensure_required_sheets(doc)
        if created:
            _log("созданы листы: %s" % u", ".join(created))
            _show_setup_message(doc, created)
            return
        if gen_sheet is None:
            _show_message(
                u"QR-коды",
                u"Не найден лист «%s»." % SHEET_GENERATOR,
                is_error=True,
            )
            return

        row_filter = _selection_row_set(doc, gen_sheet)
        selection_limited = row_filter is not None
        _log(
            "чтение данных и расчёт bar0..2… (%s)"
            % (u"выделение" if selection_limited else u"весь диапазон")
        )
        rows = _collect_rows(gen_sheet, row_filter=row_filter)
        total_cards = _count_cards(rows)
        summary = _build_summary(rows, selection_limited=selection_limited)
        _log("строк=%d карточек=%d" % (len(rows), total_cards))
        if total_cards == 0:
            if selection_limited:
                summary = (
                    u"В выделении нет строк с серийными номерами.\n\n" + summary
                )
            _show_message(u"QR-коды", summary, is_error=True)
            return
        settings = _ask_start_dialog(doc, summary)
        if settings is None:
            _log("отмена пользователем")
            return
        mapping = settings.get("mapping") or _load_mapping()
        bar_cols = settings.get("bar_cols") or _default_bar_cols()
        overwrite_bar_codes = bool(settings.get("overwrite_bar_codes"))
        # ReadOnly: в книгу не пишем, но для QR берём рассчитанные bar.
        use_computed_bars = overwrite_bar_codes
        if doc_readonly:
            overwrite_bar_codes = False
            use_computed_bars = True
        template_path = settings.get("template_path") or u""
        export_pdf = bool(settings.get("export_pdf"))
        pdf_dir = settings.get("pdf_dir") or u""
        _log(
            "маппинг: inv=%s name=%s serial=%s qr=%s size_mm=%s"
            % (
                mapping.get("inv"),
                mapping.get("name"),
                mapping.get("serial"),
                mapping.get("qr"),
                mapping.get("size_mm") or u"(автофит)",
            )
        )
        _log(
            "bar_cols=%s/%s/%s overwrite=%s computed=%s шаблон=%s pdf=%s dir=%s"
            % (
                bar_cols.get("bar0"),
                bar_cols.get("bar1"),
                bar_cols.get("bar2"),
                overwrite_bar_codes,
                use_computed_bars,
                template_path or u"(вшитый)",
                export_pdf,
                pdf_dir or u"(нет)",
            )
        )

        if overwrite_bar_codes:
            _log("запись bar0..2 в настроенные столбцы…")
            _write_all_bars(gen_sheet, rows, bar_cols)
        elif doc_readonly:
            _log("bar0..2: ReadOnly — расчёт без записи в книгу")
        else:
            _log("bar0..2: чтение с листа без перезаписи")
        _refresh_cards_from_sheet(
            gen_sheet,
            rows,
            bar_cols,
            overwrite_bar_codes=use_computed_bars,
        )

        _log("генерация карточек…")
        new_doc, created_cards = _generate_cards(
            doc,
            rows,
            mapping=mapping,
            template_path=template_path,
        )
        if new_doc is None:
            _show_message(
                u"QR-коды",
                u"Не удалось создать книгу с карточками.",
                is_error=True,
            )
            return
        pdf_note = u""
        if export_pdf:
            pdf_path, pdf_err = _export_cards_pdf(new_doc, rows, pdf_dir)
            if pdf_path:
                pdf_note = u"\nPDF: %s" % pdf_path
            elif pdf_err:
                pdf_note = u"\nPDF: %s" % pdf_err
        _activate_document(doc)
        done_text = u"Готово: создано карточек %d из %d.%s" % (
            created_cards,
            total_cards,
            pdf_note,
        )
        _show_done_dialog(doc, new_doc, done_text)
        _log("конец, карточек=%d" % created_cards)
    except Exception as err:
        _log("ошибка: %s" % err)
        _show_message(u"QR-коды", u"Ошибка: %s" % err, is_error=True)
