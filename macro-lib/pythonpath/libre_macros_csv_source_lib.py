# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.722"
"""
Чтение CSV-источников для collect_workbooks.

UNO: FilterName «Text - txt - csv (StarCalc)» + FilterOptions.
Direct (xml_excel_ods): конвертация во временный .xlsx (openpyxl).

Доп. параметры (JSON): delimiter, encoding (по умолчанию windows-1251),
copy_whole_sheet (режим «Копирование листов»: открыть в Calc и importSheet),
open_as_xlsx (по умолчанию true — CSV → временный .xlsx, тот же путь что у книг).
"""

import csv
import io
import os
import re
import tempfile

try:
    unicode
except NameError:
    unicode = str

try:
    unichr
except NameError:
    unichr = chr

# Разделитель по умолчанию (RU/EU CSV в Excel/Calc).
DEFAULT_CSV_DELIMITER = u";"
# Кодировка по умолчанию — типичный экспорт Excel/1С на Windows (кириллица).
DEFAULT_CSV_ENCODING = u"windows-1251"

_CSV_FILTER_NAME = u"Text - txt - csv (StarCalc)"

# Код → (python codec, LibreOffice rtl_TextEncoding)
# LO: MS_1251=34, UTF-8=76, MS_1252=36, IBM_866=8? — см. rtl/textenc.h
_ENCODING_TABLE = (
    (u"windows-1251", u"cp1251", 34),
    (u"cp1251", u"cp1251", 34),
    (u"win1251", u"cp1251", 34),
    (u"utf-8", u"utf-8", 76),
    (u"utf8", u"utf-8", 76),
    (u"utf-8-sig", u"utf-8-sig", 76),
    # UTF-16 семейство: для direct-режима используем python codec;
    # для UNO charset оставляем UTF-8 fallback (76), а для copy_whole_sheet
    # делаем preconvert CSV -> temp XLSX (см. collect_workbooks.py).
    (u"utf-16", u"utf-16", 76),
    (u"utf16", u"utf-16", 76),
    (u"utf-16le", u"utf-16le", 76),
    (u"utf16le", u"utf-16le", 76),
    (u"utf-16be", u"utf-16be", 76),
    (u"utf16be", u"utf-16be", 76),
    (u"windows-1252", u"cp1252", 36),
    (u"cp1252", u"cp1252", 36),
    (u"iso-8859-5", u"iso-8859-5", 16),
    (u"koi8-r", u"koi8-r", 78),
    (u"cp866", u"cp866", 8),
    (u"ibm866", u"cp866", 8),
    (u"system", u"cp1251", 0),
)


def is_csv_source_path(path):
    p = unicode(path or u"").strip().lower()
    if p.endswith(u".csv"):
        return True
    m = re.match(r"^[^:]+://.+", p)
    if m is not None:
        q = p.split(u"?", 1)[0]
        return q.endswith(u".csv")
    return False


def normalize_csv_delimiter(raw):
    """
    Нормализация разделителя из JSON/ячейки.

    Поддерживает: «;», «,», «|», «\\t»/tab, ASCII-код (число или строка цифр).
    Пусто → DEFAULT_CSV_DELIMITER.
    """
    if raw is None:
        return DEFAULT_CSV_DELIMITER
    if isinstance(raw, bool):
        return DEFAULT_CSV_DELIMITER
    if isinstance(raw, int):
        if 1 <= int(raw) <= 255:
            return unichr(int(raw))
        return DEFAULT_CSV_DELIMITER
    text_raw = unicode(raw)
    # В JSON "\\t" после json.loads становится реальным TAB-символом.
    # Нельзя сначала strip(), иначе "\t" превратится в пустую строку.
    if text_raw == u"\t":
        return u"\t"
    text = text_raw.strip()
    if text == u"":
        return DEFAULT_CSV_DELIMITER
    low = text.lower()
    if low in (u"\\t", u"tab", u"tabs", u"таб", u"табуляция"):
        return u"\t"
    if low in (u"semicolon", u"точка_с_запятой", u"точка с запятой"):
        return u";"
    if low in (u"comma", u"запятая"):
        return u","
    if low in (u"pipe", u"vertical", u"|"):
        return u"|"
    if text.isdigit():
        try:
            code = int(text)
            if 1 <= code <= 255:
                return unichr(code)
        except (TypeError, ValueError):
            pass
    if text == u"\\t":
        return u"\t"
    return text[0]


def normalize_csv_encoding(raw):
    """
    Каноническое имя кодировки для JSON/ячейки.

    Пусто / неизвестно → DEFAULT_CSV_ENCODING (windows-1251).
    """
    if raw is None:
        return DEFAULT_CSV_ENCODING
    text = unicode(raw).strip()
    if text == u"":
        return DEFAULT_CSV_ENCODING
    low = text.lower().replace(u"_", u"-")
    aliases = {
        u"windows1251": u"windows-1251",
        u"win-1251": u"windows-1251",
        u"1251": u"windows-1251",
        u"ansi": u"windows-1251",
        u"кириллица": u"windows-1251",
        u"cyrillic": u"windows-1251",
        u"unicode": u"utf-8",
        u"utf": u"utf-8",
        u"utf16": u"utf-16",
        u"utf16le": u"utf-16le",
        u"utf16be": u"utf-16be",
    }
    if low in aliases:
        low = aliases[low]
    i = 0
    while i < len(_ENCODING_TABLE):
        code, _py, _lo = _ENCODING_TABLE[i]
        if low == code or low == _py:
            # Канон для JSON: windows-1251 / utf-8 / …
            if code in (u"cp1251", u"win1251"):
                return u"windows-1251"
            if code in (u"utf8",):
                return u"utf-8"
            if code in (u"cp1252",):
                return u"windows-1252"
            if code in (u"ibm866",):
                return u"cp866"
            if code == u"system":
                return DEFAULT_CSV_ENCODING
            return code
        i = i + 1
    return DEFAULT_CSV_ENCODING


def csv_encoding_python_name(encoding=None):
    """Имя codec для codecs.decode / open(encoding=…)."""
    canon = normalize_csv_encoding(encoding)
    i = 0
    while i < len(_ENCODING_TABLE):
        code, py_name, _lo = _ENCODING_TABLE[i]
        if canon == code or canon == py_name:
            return py_name
        i = i + 1
    return u"cp1251"


def csv_encoding_uno_charset(encoding=None):
    """Токен charset для FilterOptions CSV (rtl_TextEncoding)."""
    canon = normalize_csv_encoding(encoding)
    i = 0
    while i < len(_ENCODING_TABLE):
        code, _py, lo_code = _ENCODING_TABLE[i]
        if canon == code or canon == _py:
            return int(lo_code)
        i = i + 1
    return 34


def parse_source_extra_dict(raw):
    """
    Разобрать JSON доп. параметров источника.

    Возвращает dict или None (пусто / не JSON).
    """
    import json

    text = unicode(raw or u"").strip()
    if text == u"":
        return None
    try:
        obj = json.loads(text)
    except Exception:
        return None
    if not isinstance(obj, dict):
        return None
    return obj


def source_extra_delimiter(extra):
    """Разделитель из dict доп. параметров (с default)."""
    if not isinstance(extra, dict):
        return DEFAULT_CSV_DELIMITER
    if u"delimiter" in extra:
        return normalize_csv_delimiter(extra.get(u"delimiter"))
    if u"разделитель" in extra:
        return normalize_csv_delimiter(extra.get(u"разделитель"))
    return DEFAULT_CSV_DELIMITER


def source_extra_encoding(extra):
    """Кодировка из dict доп. параметров (с default windows-1251)."""
    if not isinstance(extra, dict):
        return DEFAULT_CSV_ENCODING
    if u"encoding" in extra:
        return normalize_csv_encoding(extra.get(u"encoding"))
    if u"кодировка" in extra:
        return normalize_csv_encoding(extra.get(u"кодировка"))
    if u"charset" in extra:
        return normalize_csv_encoding(extra.get(u"charset"))
    return DEFAULT_CSV_ENCODING


def _parse_source_extra_bool(raw, default=False):
    """Разбор bool из JSON доп. параметров (Да/Нет, true/false, 1/0)."""
    if raw is None:
        return bool(default)
    if isinstance(raw, bool):
        return bool(raw)
    if isinstance(raw, (int, float)) and (not isinstance(raw, bool)):
        return bool(int(raw))
    s = unicode(raw or u"").strip().lower()
    if s == u"":
        return bool(default)
    if s in (
        u"1", u"+", u"true", u"yes", u"y", u"да", u"д",
        u"вкл", u"включено", u"on", u"истина",
    ):
        return True
    if s in (
        u"0", u"-", u"false", u"no", u"n", u"нет", u"н",
        u"выкл", u"выключено", u"off", u"ложь",
    ):
        return False
    return bool(default)


def source_extra_copy_whole_sheet(extra):
    """
    «При копировании переносить всем листом».

    True → в режиме «Копирование листов» открыть CSV в Calc и importSheet
    (не матрица → setDataArray). Отдельная конвертация чисел/дат не нужна —
    Calc распознаёт типы при открытии. По умолчанию False.
    """
    if not isinstance(extra, dict):
        return False
    for key in (
        u"copy_whole_sheet",
        u"переносить_листом",
        u"переносить_всем_листом",
        u"copy_as_sheet",
        u"whole_sheet",
    ):
        if key in extra:
            return _parse_source_extra_bool(extra.get(key), default=False)
    return False


def csv_logical_sheet_title(path):
    """Имя логического листа CSV (stem файла, безопасно для Calc / openpyxl)."""
    base = os.path.basename(unicode(path or u"").strip())
    title = base
    if title.lower().endswith(u".csv"):
        title = title[:-4]
    title = unicode(title or u"").strip()
    if title == u"":
        title = u"CSV"
    for ch in u"[]:*/?\\":
        title = title.replace(ch, u"_")
    if len(title) > 31:
        title = title[:31]
    return title


def source_extra_open_as_xlsx(extra):
    """
    Теневая схема CSV → temp XLSX.

    По умолчанию True: тот же путь, что у обычных xlsx-источников
    (листы / опции / xml_excel_ods). Явное false — матрица в памяти.
    """
    if not isinstance(extra, dict):
        return True
    for key in (
        u"open_as_xlsx",
        u"csv_open_as_xlsx",
        u"preconvert_xlsx",
        u"csv_preconvert_xlsx",
        u"shadow_xlsx",
        u"теневой_xlsx",
    ):
        if key in extra:
            return _parse_source_extra_bool(extra.get(key), default=True)
    return True


def encode_source_extra_json(extra):
    """Сериализация dict доп. параметров в JSON-строку для ячейки."""
    import json

    if not isinstance(extra, dict) or not extra:
        return u""
    payload = {}
    if u"delimiter" in extra:
        delim = normalize_csv_delimiter(extra.get(u"delimiter"))
        if delim == u"\t":
            payload[u"delimiter"] = u"\\t"
        else:
            payload[u"delimiter"] = delim
    if u"encoding" in extra or u"кодировка" in extra or u"charset" in extra:
        enc_raw = extra.get(u"encoding")
        if enc_raw is None:
            enc_raw = extra.get(u"кодировка")
        if enc_raw is None:
            enc_raw = extra.get(u"charset")
        payload[u"encoding"] = normalize_csv_encoding(enc_raw)
    elif u"delimiter" in payload:
        # Всегда пишем encoding явно, чтобы на листе было видно значение.
        payload[u"encoding"] = DEFAULT_CSV_ENCODING
    # copy_whole_sheet — пишем явно (и при False), если ключ был в форме/исходном JSON
    has_whole = False
    for key in (
        u"copy_whole_sheet",
        u"переносить_листом",
        u"переносить_всем_листом",
        u"copy_as_sheet",
        u"whole_sheet",
    ):
        if key in extra:
            has_whole = True
            break
    if has_whole or source_extra_copy_whole_sheet(extra):
        payload[u"copy_whole_sheet"] = bool(source_extra_copy_whole_sheet(extra))
    has_open_as_xlsx = False
    for key in (
        u"open_as_xlsx",
        u"csv_open_as_xlsx",
        u"preconvert_xlsx",
        u"csv_preconvert_xlsx",
        u"shadow_xlsx",
        u"теневой_xlsx",
    ):
        if key in extra:
            has_open_as_xlsx = True
            break
    if has_open_as_xlsx:
        payload[u"open_as_xlsx"] = bool(source_extra_open_as_xlsx(extra))
    elif u"delimiter" in payload or u"encoding" in payload:
        payload[u"open_as_xlsx"] = True
    try:
        from libre_macros_source_extra_lib import source_extra_exclude_mask_text

        excl = source_extra_exclude_mask_text(extra)
    except Exception:
        excl = u""
        for key in (
            u"exclude_mask",
            u"MERGE_SOURCE_EXCLUDE_MASK",
            u"source_exclude_mask",
            u"исключить",
            u"маски_исключения",
        ):
            if key in extra and unicode(extra.get(key) or u"").strip() != u"":
                excl = unicode(extra.get(key) or u"").strip()
                break
    if excl != u"":
        payload[u"exclude_mask"] = excl
    try:
        from libre_macros_source_extra_lib import source_extra_last_row_column

        last_row_col = source_extra_last_row_column(extra)
    except Exception:
        last_row_col = unicode(extra.get(u"last_row_column") or u"").strip()
    if last_row_col != u"":
        payload[u"last_row_column"] = last_row_col
    try:
        from libre_macros_source_extra_lib import (
            source_extra_file_pick,
            source_extra_file_pick_lambda,
            source_extra_file_pick_scope,
        )

        fp = source_extra_file_pick(extra)
        if fp != u"all":
            payload[u"file_pick"] = fp
            scope = source_extra_file_pick_scope(extra)
            if scope != u"all":
                payload[u"file_pick_scope"] = scope
            if fp == u"lambda":
                lam = source_extra_file_pick_lambda(extra)
                if lam != u"":
                    payload[u"file_pick_lambda"] = lam
    except Exception:
        pass
    if not payload:
        return u""
    try:
        return unicode(json.dumps(payload, ensure_ascii=False, separators=(u",", u":")))
    except Exception:
        return u""


def csv_uno_filter_name():
    return _CSV_FILTER_NAME


def csv_uno_filter_options(delimiter=None, encoding=None):
    """
    FilterOptions для Text - txt - csv (StarCalc).

    Формат: fieldSep,textDelim,charset,firstLine,colFormats[,…]
    charset — rtl_TextEncoding (34=Windows-1251, 76=UTF-8, 0=system).
    """
    delim = normalize_csv_delimiter(delimiter)
    charset = csv_encoding_uno_charset(encoding)
    try:
        sep_code = ord(delim[0])
    except Exception:
        sep_code = ord(DEFAULT_CSV_DELIMITER)
    return u"%d,34,%d,1,,0,true,true" % (int(sep_code), int(charset))


def _decode_csv_bytes(raw, encoding=None):
    """Декодировать байты CSV в unicode с учётом выбранной кодировки."""
    primary = csv_encoding_python_name(encoding)
    tried = []
    for enc in (primary, u"utf-8-sig", u"utf-8", u"cp1251", u"latin-1"):
        if enc in tried:
            continue
        tried.append(enc)
        try:
            text = raw.decode(enc)
            # На BOM-файлах utf-16le/utf-16be возможен U+FEFF в начале.
            if text.startswith(u"\ufeff"):
                return text[1:]
            return text
        except (LookupError, UnicodeDecodeError):
            continue
    text = raw.decode(u"cp1251", errors=u"replace")
    if text.startswith(u"\ufeff"):
        return text[1:]
    return text


def csv_encoding_needs_uno_preconvert(encoding=None):
    """
    Для некоторых кодировок UNO CSV FilterOptions даёт нестабильный результат.
    В таких случаях надёжнее конвертировать CSV во временный XLSX.
    """
    canon = normalize_csv_encoding(encoding)
    return canon in (u"utf-16", u"utf-16le", u"utf-16be")


def csv_source_read_rows(path, delimiter=None, encoding=None):
    """
    Прочитать CSV → (headers, rows).

    Первая строка — заголовки (если файл не пуст).
    """
    p = unicode(path or u"").strip()
    if p == u"" or not os.path.isfile(p):
        return ([], [])
    delim = normalize_csv_delimiter(delimiter)
    with open(p, u"rb") as f:
        raw = f.read()
    text = _decode_csv_bytes(raw, encoding=encoding)
    if text == u"":
        return ([], [])
    use_text = True
    try:
        if isinstance(u"", bytes):
            use_text = False
    except Exception:
        use_text = True
    if use_text:
        try:
            stream = io.StringIO(text)
        except TypeError:
            stream = io.StringIO(unicode(text))
        reader = csv.reader(stream, delimiter=delim)
    else:
        # Py2: перекодируем в unicode-строки через text
        try:
            stream = io.StringIO(text)
        except Exception:
            stream = io.BytesIO(text.encode(u"utf-8"))
            delim_b = delim.encode(u"utf-8") if isinstance(delim, unicode) else delim
            reader = csv.reader(stream, delimiter=delim_b)
            rows = []
            for row in reader:
                out_row = []
                for cell in row:
                    if cell is None:
                        out_row.append(u"")
                    elif isinstance(cell, bytes):
                        out_row.append(cell.decode(u"utf-8", u"replace"))
                    else:
                        out_row.append(unicode(cell))
                rows.append(out_row)
            if not rows:
                return ([], [])
            return (rows[0], rows[1:])
        reader = csv.reader(stream, delimiter=delim)
    rows = []
    for row in reader:
        out_row = []
        for cell in row:
            if cell is None:
                out_row.append(u"")
            elif isinstance(cell, bytes):
                try:
                    out_row.append(cell.decode(u"utf-8"))
                except Exception:
                    out_row.append(cell.decode(u"cp1251", u"replace"))
            else:
                out_row.append(unicode(cell))
        rows.append(out_row)
    if not rows:
        return ([], [])
    return (rows[0], rows[1:])


def csv_source_to_temp_xlsx(path, delimiter=None, encoding=None, prefix=u"libre_macros_csv_"):
    """
    Конвертировать локальный .csv во временный .xlsx.
    Возвращает путь или None.
    """
    headers, data_rows = csv_source_read_rows(path, delimiter=delimiter, encoding=encoding)
    if not headers and not data_rows:
        return None
    try:
        import openpyxl_bundled  # noqa: F401
        from openpyxl import Workbook
    except ImportError:
        return None
    fd, out_path = tempfile.mkstemp(prefix=prefix, suffix=u".xlsx")
    os.close(fd)
    wb = Workbook()
    ws = wb.active
    try:
        ws.title = csv_logical_sheet_title(path)
    except Exception:
        try:
            ws.title = u"CSV"
        except Exception:
            pass
    if headers:
        ws.append([unicode(h) for h in headers])
    for row in data_rows:
        ws.append([unicode(c) if c is not None else u"" for c in row])
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row or 1, min_col=1, max_col=max(1, ws.max_column or 1)):
        for cell in row:
            cell.number_format = u"@"
    try:
        wb.save(out_path)
    except Exception:
        try:
            os.remove(out_path)
        except Exception:
            pass
        return None
    return out_path
