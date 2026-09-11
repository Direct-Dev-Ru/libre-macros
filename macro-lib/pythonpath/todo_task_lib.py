# -*- coding: utf-8 -*-
"""
Логика макросов todo_task_*: добавление / удаление / копирование / правка строк задач.

Признак задачи — непустая колонка «Наименование».
Порядковые номера в «№ п/п» могут быть предзаполнены; после операций проверяется
(и при необходимости чинится) непрерывная нумерация 1..N по строкам-задачам.
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.721"
import fnmatch
import re
import time
import uuid

import uno
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, WARNINGBOX, QUERYBOX

try:
    from com.sun.star.awt.MessageBoxResults import OK as _MB_OK
except Exception:
    _MB_OK = 1

import todo_task_cfg as _cfg

try:
    unicode
except NameError:
    unicode = str


def _console_log_enabled():
    try:
        from todo_task_settings_lib import get_console_log_enabled

        return bool(get_console_log_enabled(doc=_get_doc()))
    except Exception:
        return True


def _log(msg):
    if not _console_log_enabled():
        return
    try:
        print("[todo_task] %s" % msg)
    except Exception:
        pass


def _ms(seconds):
    try:
        return "%.0fms" % (float(seconds) * 1000.0)
    except Exception:
        return "?ms"


class _StageTimer(object):
    """Log macro stages with timings."""

    def __init__(self, name):
        self.name = unicode(name or u"op")
        self.t0 = time.time()
        self.last = self.t0
        _log("%s: START" % self.name)

    def step(self, label, detail=u""):
        now = time.time()
        extra = u""
        if detail:
            extra = u" | %s" % unicode(detail)
        _log(
            "%s: %s (%s step, %s total)%s"
            % (
                self.name,
                unicode(label),
                _ms(now - self.last),
                _ms(now - self.t0),
                extra,
            )
        )
        self.last = now

    def done(self, label=u"DONE", detail=u""):
        now = time.time()
        extra = u""
        if detail:
            extra = u" | %s" % unicode(detail)
        _log(
            "%s: %s (%s total)%s"
            % (self.name, unicode(label), _ms(now - self.t0), extra)
        )


def _script_context():
    try:
        import __main__

        xsc = getattr(__main__, "XSCRIPTCONTEXT", None)
        if xsc is not None:
            return xsc
    except Exception:
        pass
    try:
        import builtins

        xsc = getattr(builtins, "XSCRIPTCONTEXT", None)
        if xsc is not None:
            return xsc
    except Exception:
        pass
    return None


def _get_doc(doc=None):
    if doc is not None:
        return doc
    xsc = _script_context()
    if xsc is not None:
        try:
            return xsc.getDocument()
        except Exception:
            pass
    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        desk = sm.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        return desk.getCurrentComponent()
    except Exception:
        return None


def _msg(doc, text, title=u"Задачи", box_type=MESSAGEBOX, buttons=BUTTONS_OK):
    """Кастомное сообщение с тёмно-синей темой (fallback — системный MessageBox)."""
    text = unicode(text or u"")
    title = unicode(title or u"Задачи")
    try:
        from todo_task_dialog_lib import show_todo_task_message

        show_todo_task_message(doc, text, title=title)
        return _MB_OK
    except Exception as err:
        _log("custom msg fail: %s" % err)
    try:
        parent = doc.getCurrentController().getFrame().getContainerWindow()
        toolkit = parent.getToolkit()
        box = toolkit.createMessageBox(parent, box_type, buttons, title, text)
        return box.execute()
    except Exception:
        pass
    try:
        _log("%s: %s" % (title, text))
    except Exception:
        pass
    return _MB_OK


def _confirm(doc, text, title=u"Задачи"):
    try:
        from todo_task_dialog_lib import show_todo_task_confirm

        return bool(show_todo_task_confirm(doc, text, title=title))
    except Exception as err:
        _log("custom confirm fail: %s" % err)
    res = _msg(doc, text, title=title, box_type=QUERYBOX, buttons=BUTTONS_OK_CANCEL)
    return res == _MB_OK


def _cell_text(cell):
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
            s = unicode(s).strip()
            if s:
                return s
    except Exception:
        pass
    try:
        v = cell.Value
        if v not in (None,):
            if float(v) == 0.0:
                # пустая числовая ячейка часто даёт 0 — для № п/п 0 не считаем
                pass
            elif float(v) == int(float(v)):
                return unicode(int(float(v)))
            else:
                return unicode(v)
    except Exception:
        pass
    return u""


def _set_cell_text(cell, text):
    text = unicode(text if text is not None else u"")
    try:
        cell.setString(text)
        return
    except Exception:
        pass
    try:
        cell.String = text
    except Exception:
        pass


def parse_due_parts(text):
    """(day, month, year) из «13.08.2026 Чт» или None."""
    s = unicode(text or u"").strip()
    m = re.match(r"^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", s)
    if not m:
        return None
    d = int(m.group(1))
    mo = int(m.group(2))
    y = int(m.group(3))
    if not (1 <= d <= 31 and 1 <= mo <= 12 and 1900 <= y <= 2100):
        return None
    try:
        from datetime import date

        date(y, mo, d)
    except ValueError:
        return None
    return d, mo, y


def default_due_parts():
    """Сегодня + N дней; если суббота/воскресенье — сдвиг на понедельник."""
    try:
        from datetime import date, timedelta

        d = date.today() + timedelta(days=int(_cfg.DUE_DATE_OFFSET_DAYS))
        while d.weekday() >= 5:
            d += timedelta(days=1)
        return int(d.day), int(d.month), int(d.year)
    except Exception:
        return 1, 1, 2026


def validate_due_not_before_today(day, month, year):
    try:
        from datetime import date

        d = date(int(year), int(month), int(day))
    except (TypeError, ValueError):
        return False, u"Некорректная дата."
    if d < date.today():
        return False, u"Срок исполнения не может быть раньше сегодняшнего дня."
    return True, u""


_due_fmt_key_cache = {}
_timestamp_fmt_key_cache = {}


def _get_timestamp_number_format_key(doc):
    if doc is None:
        return None
    doc_id = id(doc)
    if doc_id in _timestamp_fmt_key_cache:
        return _timestamp_fmt_key_cache[doc_id]
    fmt = unicode(
        getattr(_cfg, "TIMESTAMP_CELL_NUMBER_FORMAT", u"DD.MM.YYYY HH:MM:SS")
    )
    try:
        nf = doc.NumberFormats
        loc = uno.createUnoStruct("com.sun.star.lang.Locale")
        loc.Language = u"ru"
        loc.Country = u"RU"
        key = nf.queryKey(fmt, loc, True)
        if key == -1:
            key = nf.addNew(fmt, loc)
        _timestamp_fmt_key_cache[doc_id] = int(key)
        return int(key)
    except Exception:
        return None


def set_cell_timestamp(doc, cell, when=None):
    """Записать дату-время в ячейку (формат DD.MM.YYYY HH:MM:SS)."""
    if cell is None:
        return
    try:
        from datetime import datetime, date

        if when is None:
            when = datetime.now()
        elif not isinstance(when, datetime):
            when = datetime.now()
        epoch = date(1899, 12, 30)
        days = (when.date() - epoch).days
        frac = (
            float(when.hour) * 3600.0
            + float(when.minute) * 60.0
            + float(when.second)
        ) / 86400.0
        cell.Value = float(days) + frac
        fmt_key = _get_timestamp_number_format_key(doc)
        if fmt_key is not None:
            cell.NumberFormat = fmt_key
    except Exception:
        try:
            from datetime import datetime

            when = when or datetime.now()
            _set_cell_text(
                cell,
                when.strftime(u"%d.%m.%Y %H:%M:%S"),
            )
        except Exception:
            pass


def _stamp_task_timestamps(doc, sheet, cols, row, is_new=False):
    """Обновить колонки создания / редактирования после сохранения задачи."""
    if u"modified" in cols:
        set_cell_timestamp(
            doc, sheet.getCellByPosition(cols[u"modified"], row)
        )
    if is_new and u"created" in cols:
        set_cell_timestamp(
            doc, sheet.getCellByPosition(cols[u"created"], row)
        )


def _get_owner_display_name(settings=None, doc=None):
    try:
        from todo_task_settings_lib import get_owner_display_name

        return unicode(get_owner_display_name(settings, doc=doc) or u"")
    except Exception:
        pass
    try:
        import getpass

        return unicode(getpass.getuser() or u"")
    except Exception:
        return u""


def _creator_matches_owner(created_by, doc=None):
    creator = unicode(created_by or u"").strip()
    owner = _get_owner_display_name(doc=doc)
    if not creator or not owner:
        return True
    c_cf = creator.casefold() if hasattr(creator, "casefold") else creator.lower()
    o_cf = owner.casefold() if hasattr(owner, "casefold") else owner.lower()
    return c_cf == o_cf


def is_foreign_creator_task(fields, doc=None):
    return not _creator_matches_owner((fields or {}).get(u"created_by"), doc=doc)


def build_task_edit_access(sheet, fields, doc=None, use_default_due=False):
    """Флаги доступа к форме задачи и плану (просмотр / редактирование)."""
    if use_default_due:
        return {
            u"read_only": False,
            u"plan_read_only": False,
            u"can_unlock_edit": False,
            u"is_merge_source": False,
            u"foreign_creator": False,
        }
    is_merge = _is_merge_source_sheet(sheet, doc=doc)
    foreign = is_foreign_creator_task(fields, doc=doc)
    if is_merge:
        return {
            u"read_only": True,
            u"plan_read_only": True,
            u"can_unlock_edit": False,
            u"is_merge_source": True,
            u"foreign_creator": foreign,
        }
    if foreign:
        return {
            u"read_only": True,
            u"plan_read_only": True,
            u"can_unlock_edit": True,
            u"is_merge_source": False,
            u"foreign_creator": True,
        }
    return {
        u"read_only": False,
        u"plan_read_only": False,
        u"can_unlock_edit": False,
        u"is_merge_source": False,
        u"foreign_creator": False,
    }


def _owner_cell_is_blank(cell):
    """Пусто / числовой 0 / строка «0» — считать незаполненным автором."""
    try:
        if hasattr(cell, "getString"):
            s = unicode(cell.getString() or u"").strip()
        else:
            s = unicode(getattr(cell, "String", u"") or u"").strip()
    except Exception:
        s = u""
    if s and s != u"0":
        return False
    try:
        v = float(cell.Value)
        if v == 0.0:
            return True
    except Exception:
        pass
    return not s


def _stamp_task_owner(sheet, cols, row, is_new=False, settings=None, doc=None):
    """Записать ФИО владельца (или пользователя ОС) в «Создано» / «Отредактировано»."""
    owner = _get_owner_display_name(settings, doc=doc)
    if is_new and u"created_by" in cols:
        _set_cell_text(sheet.getCellByPosition(cols[u"created_by"], row), owner)
    if u"modified_by" in cols:
        _set_cell_text(sheet.getCellByPosition(cols[u"modified_by"], row), owner)


def _clear_task_meta_cells(sheet, cols, row):
    """Очистить GUID, метки времени и автора (отмена новой / копии)."""
    for key in (u"guid", u"created", u"modified", u"created_by", u"modified_by"):
        if key in cols:
            _clear_cell_value(sheet.getCellByPosition(cols[key], row))


def _get_due_number_format_key(doc):
    if doc is None:
        return None
    doc_id = id(doc)
    if doc_id in _due_fmt_key_cache:
        return _due_fmt_key_cache[doc_id]
    fmt = unicode(getattr(_cfg, "DUE_CELL_NUMBER_FORMAT", u"DD.MM.YYYY DDD"))
    try:
        nf = doc.NumberFormats
        loc = uno.createUnoStruct("com.sun.star.lang.Locale")
        loc.Language = u"ru"
        loc.Country = u"RU"
        key = nf.queryKey(fmt, loc, True)
        if key == -1:
            key = nf.addNew(fmt, loc)
        _due_fmt_key_cache[doc_id] = int(key)
        return int(key)
    except Exception:
        return None


def set_cell_due_date(doc, cell, day, month, year):
    """Записать дату в ячейку с форматом DD.MM.YYYY DDD."""
    if cell is None:
        return
    if day is None or month is None or year is None:
        _set_cell_text(cell, u"")
        return
    try:
        from datetime import date

        d = date(int(year), int(month), int(day))
        epoch = date(1899, 12, 30)
        cell.Value = float((d - epoch).days)
        fmt_key = _get_due_number_format_key(doc)
        if fmt_key is not None:
            cell.NumberFormat = fmt_key
    except Exception:
        _set_cell_text(
            cell, u"%02d.%02d.%04d" % (int(day), int(month), int(year))
        )


def _clear_cell_value(cell):
    """Очистить ячейку без setValue(0): иначе getString() даёт «0» и проверки пустоты ломаются."""
    try:
        cell.setFormula(u"")
    except Exception:
        pass
    try:
        if hasattr(cell, "setString"):
            cell.setString(u"")
    except Exception:
        pass
    try:
        cell.String = u""
    except Exception:
        pass


def _norm_header(text):
    s = unicode(text or u"").strip().lower().replace(u"ё", u"е")
    # латинская C / c рядом с «рок» → кириллическая С (опечатка в заголовке)
    s = s.replace(u"cрок", u"срок")
    s = re.sub(r"\s+", u" ", s)
    return s


def _is_done_task_status(text):
    """Статус «Выполнено» — при копировании не переносится."""
    s = _norm_header(text)
    if not s:
        return False
    target = _norm_header(getattr(_cfg, "COPY_CLEAR_DONE_STATUS", u"выполнено"))
    return s == target


def _prepare_copy_fields(copy_fields, doc=None):
    """Поля для копии: срок и статус «Выполнено» не переносятся; исполнитель — по роли."""
    fields = dict(copy_fields or {})
    if _is_done_task_status(fields.get(u"status")):
        fields[u"status"] = u""
    fields[u"due"] = u""
    fields[u"due_parts"] = None
    fields[u"assignee"] = _default_assignee_for_new_task(doc=doc)
    return fields


def _edit_task_row(doc, sheet, cols, row, use_default_due=False):
    """Диалог правки строки; True если сохранено.

    Autofit — только изменённая строка, один раз в конце (не весь лист).
    """
    from todo_task_dialog_lib import show_todo_task_edit_dialog

    fields = read_task_fields(sheet, cols, row)
    if use_default_due:
        # новая / копия: исполнитель = владелец, если роль не «руководитель»
        fields[u"assignee"] = _default_assignee_for_new_task(doc=doc)
    edit_access = build_task_edit_access(
        sheet, fields, doc=doc, use_default_due=use_default_due
    )
    data_start = int(getattr(_cfg, "REF_LIST_DATA_START_ROW", 1))
    pri_from, pri_to = _cfg.REF_PRIORITY_ROWS
    priorities = read_ref_list(doc, _cfg.REF_PRIORITY_COL, pri_from, pri_to)
    assignees = read_ref_list(doc, _cfg.REF_ASSIGNEE_COL, data_start, None)
    # владелец в списке исполнителей (для удобства выбора), без записи в справочник
    def_asg = unicode(fields.get(u"assignee") or u"").strip()
    if def_asg:
        key = def_asg.casefold() if hasattr(def_asg, "casefold") else def_asg.lower()
        have = False
        for a in assignees:
            ak = a.casefold() if hasattr(a, "casefold") else a.lower()
            if ak == key:
                have = True
                break
        if not have:
            assignees = [def_asg] + list(assignees)
    ref_statuses = read_status_ref_list(doc)
    sheet_statuses = unique_status_values(sheet, cols)
    statuses = merge_status_choices(ref_statuses, sheet_statuses)

    result = show_todo_task_edit_dialog(
        doc=doc,
        fields=fields,
        priorities=priorities,
        assignees=assignees,
        statuses=statuses,
        use_default_due=use_default_due,
        edit_access=edit_access,
        plan_ctx={
            u"sheet": sheet,
            u"cols": cols,
            u"row": row,
            u"guid": fields.get(u"guid") or u"",
            u"sheet_name": _sheet_display_name(sheet),
            u"read_only": edit_access.get(u"read_only"),
            u"plan_read_only": edit_access.get(u"plan_read_only"),
            u"can_unlock_edit": edit_access.get(u"can_unlock_edit"),
            u"is_merge_source": edit_access.get(u"is_merge_source"),
            u"foreign_creator": edit_access.get(u"foreign_creator"),
        },
    )
    if result is None:
        # Новая / копия: при отмене очистить GUID и метки времени.
        if use_default_due:
            gid = unicode(fields.get(u"guid") or u"").strip()
            _clear_task_meta_cells(sheet, cols, row)
            if gid:
                try:
                    from todo_task_plan_lib import delete_plan_by_guid

                    delete_plan_by_guid(doc, gid)
                except Exception as exc:
                    _log("delete plan on cancel: %s" % exc)
            _autofit_task_rows(sheet, [row])
        return False

    write_task_fields(sheet, cols, row, result, clear_due=False, doc=doc)
    # исполнителей в справочник не добавляем
    _stamp_task_timestamps(doc, sheet, cols, row, is_new=use_default_due)
    _stamp_task_owner(sheet, cols, row, is_new=use_default_due, doc=doc)
    ok, note = check_and_fix_task_numbers(sheet, cols, fix=True)
    _autofit_task_rows(sheet, [row])
    if note:
        _msg(doc, u"Задача сохранена.\n\n%s" % note)
    return True


def _active_sheet(doc):
    try:
        return doc.getCurrentController().getActiveSheet()
    except Exception:
        pass
    try:
        return doc.Sheets.getByIndex(0)
    except Exception:
        return None


def _cursor_row(doc, sheet):
    try:
        sel = doc.getCurrentController().getSelection()
        if sel is None:
            return None
        if hasattr(sel, "getRangeAddress"):
            addr = sel.getRangeAddress()
            return int(addr.StartRow)
        if hasattr(sel, "getCount") and sel.getCount() > 0:
            rng = sel.getByIndex(0)
            addr = rng.getRangeAddress()
            return int(addr.StartRow)
    except Exception:
        pass
    return None


def _sheet_index(doc, sheet):
    try:
        return int(sheet.RangeAddress.Sheet)
    except Exception:
        pass
    try:
        want = unicode(getattr(sheet, "Name", u"") or u"")
        sheets = doc.Sheets
        n = int(sheets.getCount())
        i = 0
        while i < n:
            sh = sheets.getByIndex(i)
            if unicode(getattr(sh, "Name", u"") or u"") == want:
                return i
            i = i + 1
    except Exception:
        pass
    return -1


def _iter_selected_rows(doc, sheet):
    """Уникальные номера строк из текущего выделения (мультиселект)."""
    out = []
    seen = {}
    try:
        sel = doc.getCurrentController().getSelection()
    except Exception:
        return out
    if sel is None:
        return out
    ranges = []
    try:
        if hasattr(sel, "getRangeAddress"):
            ranges.append(sel.getRangeAddress())
        elif hasattr(sel, "getCount"):
            i = 0
            n = int(sel.getCount())
            while i < n:
                try:
                    ranges.append(sel.getByIndex(i).getRangeAddress())
                except Exception:
                    pass
                i = i + 1
    except Exception:
        return out
    want = _sheet_index(doc, sheet)
    ri = 0
    while ri < len(ranges):
        addr = ranges[ri]
        ri = ri + 1
        try:
            if want >= 0 and int(addr.Sheet) != want:
                continue
            r = int(addr.StartRow)
            end_r = int(addr.EndRow)
        except Exception:
            continue
        while r <= end_r:
            if r not in seen:
                seen[r] = True
                out.append(r)
            r = r + 1
    out.sort()
    return out


def _iter_selected_task_rows(doc, sheet, cols):
    rows = _iter_selected_rows(doc, sheet)
    tasks = []
    i = 0
    while i < len(rows):
        row = rows[i]
        if _is_task_row(sheet, cols, row):
            tasks.append(row)
        i = i + 1
    if tasks:
        return tasks
    cur = _cursor_row(doc, sheet)
    if cur is not None and _is_task_row(sheet, cols, cur):
        return [cur]
    return []


def _used_last_col(sheet, min_col=0):
    """Правая занятая колонка used area (0-based)."""
    last = max(0, int(min_col))
    try:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        last = max(last, int(cursor.getRangeAddress().EndColumn))
    except Exception:
        pass
    return last


def _used_last_row(sheet, min_cols=1):
    """Нижняя занятая строка по колонке A (и запас по нескольким колонкам)."""
    last = 0
    cols = max(1, int(min_cols))
    c = 0
    while c < cols:
        try:
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(False)
            addr = cursor.getRangeAddress()
            last = max(last, int(addr.EndRow))
            break
        except Exception:
            pass
        c = c + 1
    # fallback: просканировать первые колонки
    if last <= 0:
        r = 1
        while r < 5000:
            empty = True
            cc = 0
            while cc < max(cols, 9):
                try:
                    if _cell_text(sheet.getCellByPosition(cc, r)):
                        empty = False
                        last = r
                        break
                except Exception:
                    pass
                cc = cc + 1
            if empty and r > last + 50:
                break
            r = r + 1
    return int(last)


def _header_row(cols):
    try:
        return int(cols.get(u"_header_row", _cfg.HEADER_ROW))
    except Exception:
        return int(_cfg.HEADER_ROW)


def _scan_header_row(sheet, wanted_norm, last_col, row):
    """Разобрать одну строку как потенциальный заголовок. None — не хватает обязательных."""
    found = {}
    c = 0
    while c <= last_col:
        raw = _cell_text(sheet.getCellByPosition(c, row))
        if not raw:
            c = c + 1
            continue
        key = wanted_norm.get(_norm_header(raw))
        if key is not None and key not in found:
            found[key] = c
        c = c + 1
    required = (u"num", u"name", u"guid")
    for req in required:
        if req not in found:
            return None
    found[u"_header_row"] = int(row)
    return found


def _find_headers(sheet):
    """
    Найти индексы колонок по заголовкам: строки 1..20 (0-based 0..HEADER_SCAN_MAX_ROW).
    Возвращает dict key->col (+ _header_row) или None.
    """
    wanted = {
        u"num": _cfg.COL_NUM,
        u"name": _cfg.COL_NAME,
        u"due": _cfg.COL_DUE,
        u"priority": _cfg.COL_PRIORITY,
        u"sed": _cfg.COL_SED,
        u"assignee": _cfg.COL_ASSIGNEE,
        u"status": _cfg.COL_STATUS,
        u"comment": _cfg.COL_COMMENT,
        u"guid": _cfg.COL_GUID,
        u"created": _cfg.COL_CREATED,
        u"modified": _cfg.COL_MODIFIED,
        u"created_by": getattr(_cfg, "COL_CREATED_BY", u"Создано"),
        u"modified_by": getattr(_cfg, "COL_MODIFIED_BY", u"Отредактировано"),
        u"category": getattr(_cfg, "COL_CATEGORY", u"Категория"),
    }
    wanted_norm = {}
    for k, title in wanted.items():
        wanted_norm[_norm_header(title)] = k

    aliases = {
        _norm_header(u"Cрок исполнения"): u"due",
        _norm_header(u"N п/п"): u"num",
        _norm_header(u"No п/п"): u"num",
        _norm_header(u"GUID"): u"guid",
        _norm_header(u"GUID строки"): u"guid",
        _norm_header(u"дата и время создания"): u"created",
        _norm_header(u"дата и время редактирования"): u"modified",
        _norm_header(u"Создано"): u"created_by",
        _norm_header(u"Отредактировано"): u"modified_by",
        _norm_header(u"категория"): u"category",
    }
    for a, k in aliases.items():
        wanted_norm[a] = k

    last_col = 40
    try:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        last_col = max(last_col, int(cursor.getRangeAddress().EndColumn) + 2)
    except Exception:
        pass

    max_row = int(getattr(_cfg, "HEADER_SCAN_MAX_ROW", 19))
    row = 0
    while row <= max_row:
        found = _scan_header_row(sheet, wanted_norm, last_col, row)
        if found is not None:
            return found
        row = row + 1
    return None


def _is_task_row(sheet, cols, row):
    """Задача = заполненное «Наименование»."""
    if row <= _header_row(cols):
        return False
    return bool(_cell_text(sheet.getCellByPosition(cols[u"name"], row)))


def _is_managed_row(sheet, cols, row):
    """Управляемая строка = есть наименование и/или GUID."""
    if row <= _header_row(cols):
        return False
    if _cell_text(sheet.getCellByPosition(cols[u"name"], row)):
        return True
    if u"guid" in cols and _cell_text(sheet.getCellByPosition(cols[u"guid"], row)):
        return True
    return False


def _iter_task_rows(sheet, cols, last_row=None):
    if last_row is None:
        last_row = _used_last_row(sheet, min_cols=cols.get(u"guid", 8) + 1)
    rows = []
    r = _header_row(cols) + 1
    while r <= last_row:
        if _is_task_row(sheet, cols, r):
            rows.append(r)
        r = r + 1
    return rows


def _new_guid():
    return unicode(uuid.uuid4())


def _copy_cell_format(src, dst):
    props = _cfg.FORMAT_COPY_PROPS
    i = 0
    while i < len(props):
        prop = props[i]
        try:
            setattr(dst, prop, getattr(src, prop))
        except Exception:
            pass
        i = i + 1
    # сброс жирности, курсива, заливки
    try:
        dst.CharWeight = int(_cfg.CHAR_WEIGHT_NORMAL)
    except Exception:
        pass
    try:
        from com.sun.star.awt.FontSlant import NONE as _FS_NONE

        dst.CharPosture = _FS_NONE
    except Exception:
        try:
            dst.CharPosture = int(_cfg.CHAR_POSTURE_NONE)
        except Exception:
            pass
    try:
        dst.IsCellBackgroundTransparent = True
    except Exception:
        pass
    try:
        dst.CellBackColor = -1
    except Exception:
        pass


def _copy_row_formats(sheet, cols, src_row, dst_row):
    if src_row is None or dst_row is None or src_row == dst_row:
        return
    max_col = 0
    for v in cols.values():
        try:
            max_col = max(max_col, int(v))
        except Exception:
            pass
    c = 0
    while c <= max_col:
        try:
            src = sheet.getCellByPosition(c, src_row)
            dst = sheet.getCellByPosition(c, dst_row)
            _copy_cell_format(src, dst)
        except Exception:
            pass
        c = c + 1


def _parse_int_num(text):
    s = unicode(text or u"").strip()
    if not s:
        return None
    m = re.match(r"^(\d+)", s)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _write_row_numbers(sheet, col_idx, rows, start_num=1):
    """Write 1..N into column via setDataArray for contiguous blocks."""
    t0 = time.time()
    rows = list(rows or [])
    n = len(rows)
    if n <= 0:
        return
    col_idx = int(col_idx)
    batches = 0
    fallback_cells = 0
    i = 0
    while i < n:
        j = i + 1
        while j < n and int(rows[j]) == int(rows[j - 1]) + 1:
            j = j + 1
        data = []
        k = i
        while k < j:
            data.append((unicode(int(start_num) + k),))
            k = k + 1
        try:
            rng = sheet.getCellRangeByPosition(
                col_idx, int(rows[i]), col_idx, int(rows[j - 1])
            )
            rng.setDataArray(tuple(data))
            batches = batches + 1
        except Exception as exc:
            _log("renumber DataArray fail %s..%s: %s" % (rows[i], rows[j - 1], exc))
            k = i
            while k < j:
                try:
                    _set_cell_text(
                        sheet.getCellByPosition(col_idx, int(rows[k])),
                        unicode(int(start_num) + k),
                    )
                except Exception:
                    pass
                fallback_cells = fallback_cells + 1
                k = k + 1
        i = j
    _log(
        "renumber write: rows=%s batches=%s fallback_cells=%s (%s)"
        % (n, batches, fallback_cells, _ms(time.time() - t0))
    )


def check_and_fix_task_numbers(sheet, cols, fix=True):
    """
    Проверить нумерацию задач в колонке № п/п.
    Если fix=True — выставить 1..N по порядку строк-задач.
    Возвращает (ok: bool, message: unicode).
    """
    tm = _StageTimer("check_and_fix_numbers")
    tasks = _iter_task_rows(sheet, cols)
    tm.step("iter_task_rows", "n=%s" % len(tasks))
    if not tasks:
        tm.done("empty")
        return True, u""

    problems = []
    expected = 1
    any_filled = False
    for row in tasks:
        raw = _cell_text(sheet.getCellByPosition(cols[u"num"], row))
        num = _parse_int_num(raw)
        if raw:
            any_filled = True
        if num is None:
            problems.append(u"строка %s: нет номера" % (row + 1))
        elif num != expected:
            problems.append(
                u"строка %s: ожидали %s, сейчас %s" % (row + 1, expected, num)
            )
        expected = expected + 1

    if not problems:
        tm.done("ok")
        return True, u""

    if fix and any_filled:
        _write_row_numbers(sheet, cols[u"num"], tasks, start_num=1)
        tm.done("fixed_existing")
        return False, u"Нумерация задач была сбита и восстановлена:\n" + u"\n".join(
            problems[:8]
        )

    if fix and not any_filled:
        _write_row_numbers(sheet, cols[u"num"], tasks, start_num=1)
        tm.done("filled_missing")
        return False, u"Заполнены отсутствующие порядковые номера задач."

    tm.done("problems_nofix")
    return False, u"Сбита нумерация задач:\n" + u"\n".join(problems[:8])


def _ensure_row_number(sheet, cols, row, expected_num):
    cell = sheet.getCellByPosition(cols[u"num"], row)
    raw = _cell_text(cell)
    if not raw:
        _set_cell_text(cell, unicode(expected_num))
        return
    num = _parse_int_num(raw)
    if num != expected_num:
        _set_cell_text(cell, unicode(expected_num))


def _get_sheet_by_name(doc, name):
    name = unicode(name or u"")
    try:
        sheets = doc.Sheets
        if sheets.hasByName(name):
            return sheets.getByName(name)
    except Exception:
        pass
    # мягкий поиск без учёта регистра
    try:
        sheets = doc.Sheets
        n = sheets.getCount()
        i = 0
        target = _norm_header(name)
        while i < n:
            sh = sheets.getByIndex(i)
            try:
                if _norm_header(sh.getName()) == target:
                    return sh
            except Exception:
                pass
            i = i + 1
    except Exception:
        pass
    return None


def read_ref_list(doc, col_idx, row_from, row_to=None):
    """Список непустых значений из листа справочников (данные, без строки заголовка)."""
    sheet = _get_sheet_by_name(doc, _cfg.REF_SHEET_NAME)
    if sheet is None:
        return []
    data_start = int(getattr(_cfg, "REF_LIST_DATA_START_ROW", 1))
    r = max(int(row_from), data_start)
    out = []
    seen = set()
    limit = int(row_to) if row_to is not None else r + 5000
    while r <= limit:
        val = _cell_text(sheet.getCellByPosition(int(col_idx), r))
        if not val:
            if row_to is None and r > data_start:
                if out:
                    break
            r = r + 1
            continue
        key = val.casefold() if hasattr(val, "casefold") else val.lower()
        if key not in seen:
            seen.add(key)
            out.append(val)
        r = r + 1
        if row_to is not None and r > int(row_to):
            break
    return out


def append_assignee_to_ref(doc, value):
    """Устарело: исполнителей в справочник больше не пишем (оставлено no-op)."""
    return False


def unique_status_values(sheet, cols):
    out = []
    seen = set()
    for row in _iter_task_rows(sheet, cols):
        val = _cell_text(sheet.getCellByPosition(cols[u"status"], row))
        if not val:
            continue
        key = val.casefold() if hasattr(val, "casefold") else val.lower()
        if key not in seen:
            seen.add(key)
            out.append(val)
    return out


def read_status_ref_list(doc):
    """Статусы из справочника (колонка F / «Статусы исполнения»)."""
    data_start = int(getattr(_cfg, "REF_LIST_DATA_START_ROW", 1))
    col = int(getattr(_cfg, "REF_STATUS_COL", 5))
    # если заголовок блока найден через discover — предпочесть его start_col
    try:
        ref_sheet = _get_sheet_by_name(doc, _cfg.REF_SHEET_NAME)
        if ref_sheet is not None:
            want = _norm_header(
                getattr(_cfg, "REF_STATUS_TITLE", u"Статусы исполнения")
            )
            dicts = discover_ref_dictionaries(ref_sheet)
            di = 0
            while di < len(dicts):
                d = dicts[di]
                di = di + 1
                title = _norm_header(d.get(u"title") or u"")
                if title == want or want in title:
                    col = int(d.get(u"start_col", col))
                    break
    except Exception:
        pass
    return read_ref_list(doc, col, data_start, None)


def merge_status_choices(ref_statuses, sheet_statuses):
    """Справочник сначала, затем уникальные статусы с листа задач."""
    out = []
    seen = set()
    for src in (list(ref_statuses or []), list(sheet_statuses or [])):
        i = 0
        while i < len(src):
            val = unicode(src[i] or u"").strip()
            i = i + 1
            if not val:
                continue
            key = val.casefold() if hasattr(val, "casefold") else val.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(val)
    return out


def _default_assignee_for_new_task(settings=None, doc=None):
    try:
        from todo_task_settings_lib import get_default_assignee

        return unicode(get_default_assignee(settings, doc=doc) or u"")
    except Exception:
        return u""


def _read_due_from_cell(cell):
    """Дата из ячейки: строка «ДД.ММ.ГГГГ …» или числовое Value."""
    parts = parse_due_parts(_cell_text(cell))
    if parts:
        return parts
    try:
        v = cell.Value
        if v in (None, 0.0, 0):
            return None
        from datetime import date, timedelta

        d = date(1899, 12, 30) + timedelta(days=int(float(v)))
        return int(d.day), int(d.month), int(d.year)
    except Exception:
        return None


def read_task_fields(sheet, cols, row):
    def g(key):
        if key not in cols:
            return u""
        return _cell_text(sheet.getCellByPosition(cols[key], row))

    due_parts = None
    due_text = u""
    if u"due" in cols:
        cell = sheet.getCellByPosition(cols[u"due"], row)
        due_parts = _read_due_from_cell(cell)
        if due_parts:
            due_text = u"%02d.%02d.%04d" % due_parts
        else:
            due_text = _cell_text(cell)

    return {
        u"name": g(u"name"),
        u"due": due_text,
        u"due_parts": due_parts,
        u"priority": g(u"priority"),
        u"sed": g(u"sed"),
        u"assignee": g(u"assignee"),
        u"status": g(u"status"),
        u"comment": g(u"comment"),
        u"guid": g(u"guid"),
        u"created": g(u"created"),
        u"modified": g(u"modified"),
        u"created_by": g(u"created_by"),
        u"modified_by": g(u"modified_by"),
        u"category": g(u"category"),
        u"num": g(u"num"),
    }


def _table_max_col(cols):
    max_col = 0
    for key, val in cols.items():
        if unicode(key).startswith(u"_"):
            continue
        try:
            max_col = max(max_col, int(val))
        except Exception:
            pass
    return max_col


def _sort_copy_end_col(sheet, cols):
    """Правый край копирования строки при сортировке: used area, не уже известных колонок.

    На своде задач последняя колонка «Категория» не входит в карту заголовков —
    её тоже нужно переносить вместе со строкой.
    """
    last = max(_table_max_col(cols), _used_last_col(sheet))
    try:
        hdr = _header_row(cols)
    except Exception:
        return last
    c = last + 1
    empty_run = 0
    while c < last + 24 and empty_run < 4:
        try:
            filled = bool(_cell_text(sheet.getCellByPosition(c, hdr)))
        except Exception:
            filled = False
        if filled:
            last = c
            empty_run = 0
        else:
            empty_run = empty_run + 1
        c = c + 1
    return last


def _last_numbered_row(sheet, cols):
    """Последняя строка с заполненным «№ п/п» (не заголовок)."""
    hdr = _header_row(cols)
    last_used = _used_last_row(sheet, min_cols=_table_max_col(cols) + 1)
    found = None
    r = hdr + 1
    while r <= last_used:
        raw = _cell_text(sheet.getCellByPosition(cols[u"num"], r))
        if raw and _parse_int_num(raw) is not None:
            found = r
        r = r + 1
    return found


def _iter_numbered_rows(sheet, cols, last_row=None):
    if last_row is None:
        last_row = _used_last_row(sheet, min_cols=_table_max_col(cols) + 1)
    rows = []
    r = _header_row(cols) + 1
    while r <= last_row:
        raw = _cell_text(sheet.getCellByPosition(cols[u"num"], r))
        if raw and _parse_int_num(raw) is not None:
            rows.append(r)
        r = r + 1
    return rows


def renumber_all_slots(sheet, cols):
    """Ordinal numbers 1..N for all rows with No."""
    rows = _iter_numbered_rows(sheet, cols)
    _log("renumber_all_slots: rows=%s" % len(rows))
    _write_row_numbers(sheet, cols[u"num"], rows, start_num=1)


def _restore_numbered_row_after_delete(sheet, cols):
    """
    После удаления: в конец (сразу после последней строки с № п/п) — только форматы
    с этой строки; значений нет, кроме нового порядкового номера. Затем 1..N.
    Если номеров не осталось (убрали единственную задачу) — пустая строка
    сразу под заголовком с № 1.
    """
    last_num_row = _last_numbered_row(sheet, cols)
    if last_num_row is None:
        hdr = _header_row(cols)
        new_row = int(hdr) + 1
        try:
            sheet.getRows().insertByIndex(new_row, 1)
        except Exception as exc:
            _log("insert empty row after header: %s" % exc)
            return
        try:
            _copy_row_formats(sheet, cols, hdr, new_row)
        except Exception as exc:
            _log("format empty row from header: %s" % exc)
        if u"num" in cols:
            _set_cell_text(sheet.getCellByPosition(cols[u"num"], new_row), u"1")
        renumber_all_slots(sheet, cols)
        return
    new_row = int(last_num_row) + 1
    try:
        sheet.getRows().insertByIndex(new_row, 1)
    except Exception as exc:
        _log("insert template row: %s" % exc)
        renumber_all_slots(sheet, cols)
        return
    # Только форматы (бордеры и пр.), без значений; жирность/курсив/заливка сбрасываются.
    _copy_row_formats(sheet, cols, last_num_row, new_row)
    # Порядковый номер — единственное значение; остальное пусто.
    numbered = _iter_numbered_rows(sheet, cols)
    next_num = len(numbered) + 1
    _set_cell_text(sheet.getCellByPosition(cols[u"num"], new_row), unicode(next_num))
    renumber_all_slots(sheet, cols)


def write_task_fields(sheet, cols, row, fields, clear_due=False, doc=None):
    mapping = (
        (u"name", u"name"),
        (u"due", u"due"),
        (u"priority", u"priority"),
        (u"sed", u"sed"),
        (u"assignee", u"assignee"),
        (u"status", u"status"),
        (u"comment", u"comment"),
    )
    for src, col_key in mapping:
        if col_key not in cols:
            continue
        if clear_due and col_key == u"due":
            _clear_cell_value(sheet.getCellByPosition(cols[col_key], row))
            continue
        if col_key == u"due":
            if src not in fields and u"due_parts" not in fields:
                continue
            cell = sheet.getCellByPosition(cols[col_key], row)
            parts = fields.get(u"due_parts")
            if parts:
                set_cell_due_date(doc, cell, parts[0], parts[1], parts[2])
            elif clear_due or not fields.get(u"due"):
                _clear_cell_value(cell)
            else:
                p = parse_due_parts(fields.get(u"due"))
                if p and doc is not None:
                    set_cell_due_date(doc, cell, p[0], p[1], p[2])
                else:
                    _set_cell_text(cell, unicode(fields.get(u"due") or u""))
            continue
        if src in fields:
            _set_cell_text(
                sheet.getCellByPosition(cols[col_key], row),
                unicode(fields.get(src) or u""),
            )
    for wrap_key in (u"name", u"comment"):
        if wrap_key not in cols:
            continue
        try:
            sheet.getCellByPosition(cols[wrap_key], row).IsTextWrapped = True
        except Exception:
            pass


def _mm_to_row_height(mm):
    """Миллиметры → единицы высоты строки Calc (1/100 мм)."""
    return int(round(float(mm) * 100.0))


def _autofit_row(sheet, row):
    """Автоподбор высоты одной строки + доп. отступ в мм."""
    try:
        row_obj = sheet.getRows().getByIndex(int(row))
    except Exception:
        return
    try:
        row_obj.OptimalHeight = False
    except Exception:
        pass
    try:
        row_obj.OptimalHeight = True
    except Exception:
        pass
    extra_mm = float(getattr(_cfg, "AUTOFIT_ROW_HEIGHT_EXTRA_MM", 4))
    if extra_mm <= 0:
        return
    try:
        row_obj.OptimalHeight = False
    except Exception:
        pass
    try:
        h = int(row_obj.Height or 0)
    except Exception:
        h = 0
    try:
        row_obj.Height = int(h) + _mm_to_row_height(extra_mm)
    except Exception:
        pass


def _autofit_task_rows(sheet, rows):
    """Autofit только указанные строки (0-based)."""
    if sheet is None:
        return
    uniq = []
    seen = set()
    i = 0
    rows = list(rows or [])
    while i < len(rows):
        try:
            r = int(rows[i])
        except Exception:
            i = i + 1
            continue
        if r < 0 or r in seen:
            i = i + 1
            continue
        seen.add(r)
        uniq.append(r)
        i = i + 1
    if not uniq:
        return
    tm = _StageTimer("autofit_rows")
    j = 0
    while j < len(uniq):
        _autofit_row(sheet, uniq[j])
        j = j + 1
    tm.done("DONE", "rows=%s count=%s" % (uniq[:8], len(uniq)))


def _autofit_sheet_row_heights(sheet, cols):
    """Autofit row heights for task table."""
    if sheet is None or not cols:
        return
    tm = _StageTimer("autofit_rows")
    hdr = _header_row(cols)
    last = _used_last_row(sheet, min_cols=_table_max_col(cols) + 1)
    last_num = _last_numbered_row(sheet, cols)
    if last_num is not None:
        last = max(int(last), int(last_num))
    if last < hdr:
        tm.done("skip")
        return
    r = int(hdr)
    while r <= int(last):
        _autofit_row(sheet, r)
        r = r + 1
    tm.done("DONE", "rows=%s..%s count=%s" % (hdr, last, int(last) - int(hdr) + 1))


def _hide_sheet(sh):
    if sh is None:
        return False
    try:
        sh.setPropertyValue("IsVisible", False)
        return True
    except Exception:
        pass
    try:
        from com.sun.star.sheet.SheetVisibility import HIDDEN

        sh.setPropertyValue("SheetState", HIDDEN)
        return True
    except Exception:
        pass
    try:
        sh.IsVisible = False
        return True
    except Exception:
        return False


def _ensure_deleted_tasks_sheet(doc):
    """Скрытый лист архива удалённых задач; создать при необходимости."""
    name = unicode(getattr(_cfg, "DELETED_SHEET_NAME", u"__Удаленные_Задачи"))
    sheets = doc.Sheets
    if sheets.hasByName(name):
        sh = sheets.getByName(name)
    else:
        pos = int(sheets.getCount())
        sheets.insertNewByName(name, pos)
        sh = sheets.getByName(name)
    _hide_sheet(sh)
    return sh


def _deleted_sheet_next_data_row(deleted_sheet, hdr_row):
    last = _used_last_row(deleted_sheet, min_cols=3)
    r = int(hdr_row) + 1
    if r > last:
        return r
    # первая пустая строка после заголовка
    while r <= last + 1:
        if not _cell_text(deleted_sheet.getCellByPosition(0, r)):
            if not _cell_text(deleted_sheet.getCellByPosition(1, r)):
                return r
        r = r + 1
    return last + 1


def _cell_content_type(cell):
    """CellContentType: 0 пусто, 1 число, 2 текст, 3 формула."""
    if cell is None:
        return 0
    try:
        return int(cell.getPropertyValue("CellContentType"))
    except Exception:
        pass
    try:
        f = unicode(
            cell.getFormula() if hasattr(cell, "getFormula") else cell.Formula or u""
        ).strip()
        if f:
            return 3
    except Exception:
        pass
    try:
        s = unicode(
            cell.getString() if hasattr(cell, "getString") else cell.String or u""
        ).strip()
        if s:
            return 2
    except Exception:
        pass
    try:
        if float(cell.Value) != 0.0:
            return 1
    except Exception:
        pass
    return 0


def _cell_has_formula(cell):
    try:
        f = unicode(
            cell.getFormula() if hasattr(cell, "getFormula") else cell.Formula or u""
        ).strip()
        return bool(f)
    except Exception:
        return False


def _formula_result_is_text(cell):
    """Текстовый результат формулы (не число/дата)."""
    try:
        s = unicode(
            cell.getString() if hasattr(cell, "getString") else cell.String or u""
        ).strip()
        if not s:
            return False
        if s.startswith(u"#"):
            return True
        try:
            if float(cell.Value) != 0.0:
                return False
        except Exception:
            pass
        return True
    except Exception:
        return False


def _cell_is_effectively_empty(cell):
    ctype = _cell_content_type(cell)
    if ctype == 0:
        return True
    if ctype == 1:
        try:
            return float(cell.Value) == 0.0
        except Exception:
            return True
    if ctype == 2:
        try:
            return not unicode(
                cell.getString() if hasattr(cell, "getString") else cell.String or u""
            ).strip()
        except Exception:
            return True
    if ctype == 3:
        if not _cell_has_formula(cell):
            return True
        if _formula_result_is_text(cell):
            try:
                return not unicode(
                    cell.getString() if hasattr(cell, "getString") else cell.String or u""
                ).strip()
            except Exception:
                return True
        try:
            v = float(cell.Value)
            if v != 0.0:
                return False
        except Exception:
            pass
        try:
            return not unicode(
                cell.getString() if hasattr(cell, "getString") else cell.String or u""
            ).strip()
        except Exception:
            return True
    return True


def _copy_cell_value_only(src, dst):
    """
    Скопировать значение ячейки без оформления источника.
    Числа и даты — через Value (отображение по формату приёмника); текст — через String.
    """
    if src is None or dst is None:
        return
    dest_nf = None
    try:
        dest_nf = int(dst.NumberFormat)
    except Exception:
        pass
    if _cell_is_effectively_empty(src):
        _clear_cell_value(dst)
        return
    _clear_cell_value(dst)
    ctype = _cell_content_type(src)
    copied = False
    if ctype == 2:
        try:
            text = unicode(
                src.getString() if hasattr(src, "getString") else src.String or u""
            )
            if hasattr(dst, "setString"):
                dst.setString(text)
            else:
                dst.String = text
            copied = True
        except Exception:
            pass
    elif ctype == 1:
        try:
            dst.Value = src.Value
            copied = True
        except Exception:
            pass
    elif ctype == 3:
        if _formula_result_is_text(src):
            try:
                text = unicode(
                    src.getString() if hasattr(src, "getString") else src.String or u""
                )
                if hasattr(dst, "setString"):
                    dst.setString(text)
                else:
                    dst.String = text
                copied = True
            except Exception:
                pass
        else:
            try:
                dst.Value = src.Value
                copied = True
            except Exception:
                pass
    if not copied:
        try:
            dst.Value = src.Value
            copied = True
        except Exception:
            _set_cell_text(dst, _cell_text(src))
    if dest_nf is not None:
        try:
            dst.NumberFormat = dest_nf
        except Exception:
            pass


def _copy_row_values_only(sheet, src_row, dst_sheet, dst_row, col_from, col_to):
    """Строка на архивный лист — только значения, без форматирования."""
    if col_to < col_from:
        return
    c = int(col_from)
    while c <= int(col_to):
        try:
            _copy_cell_value_only(
                sheet.getCellByPosition(c, src_row),
                dst_sheet.getCellByPosition(c, dst_row),
            )
        except Exception:
            pass
        c = c + 1


def _now_timestamp_text():
    try:
        from datetime import datetime

        return datetime.now().strftime(u"%d.%m.%Y %H:%M:%S")
    except Exception:
        return u""


def _ensure_deleted_sheet_headers(source_sheet, deleted_sheet, hdr_row, max_col):
    meta_deleted_col = int(max_col) + 1
    meta_source_col = int(max_col) + 2
    if _cell_text(deleted_sheet.getCellByPosition(0, hdr_row)):
        return meta_deleted_col, meta_source_col
    _copy_row_values_only(source_sheet, hdr_row, deleted_sheet, hdr_row, 0, max_col)
    _set_cell_text(
        deleted_sheet.getCellByPosition(meta_deleted_col, hdr_row),
        unicode(getattr(_cfg, "DELETED_COL_DELETED_AT", u"Дата и время удаления")),
    )
    _set_cell_text(
        deleted_sheet.getCellByPosition(meta_source_col, hdr_row),
        unicode(getattr(_cfg, "DELETED_COL_SOURCE_SHEET", u"Лист источника")),
    )
    return meta_deleted_col, meta_source_col


def _archive_deleted_task(doc, source_sheet, cols, row):
    """Скопировать удалённую строку на скрытый лист архива (только значения)."""
    try:
        deleted_sheet = _ensure_deleted_tasks_sheet(doc)
        hdr = _header_row(cols)
        max_col = _table_max_col(cols)
        meta_deleted_col, meta_source_col = _ensure_deleted_sheet_headers(
            source_sheet, deleted_sheet, hdr, max_col
        )
        dest_row = _deleted_sheet_next_data_row(deleted_sheet, hdr)
        _copy_row_values_only(source_sheet, row, deleted_sheet, dest_row, 0, max_col)
        _set_cell_text(
            deleted_sheet.getCellByPosition(meta_deleted_col, dest_row),
            _now_timestamp_text(),
        )
        try:
            sheet_name = unicode(source_sheet.Name)
        except Exception:
            sheet_name = u""
        _set_cell_text(
            deleted_sheet.getCellByPosition(meta_source_col, dest_row),
            sheet_name,
        )
        return True
    except Exception as exc:
        _log("archive deleted task: %s" % exc)
        return False


def _task_sheet_protect_password(doc=None):
    try:
        from todo_task_settings_lib import get_protect_password

        return unicode(get_protect_password(doc=doc) or u"")
    except Exception:
        return unicode(getattr(_cfg, "TASK_SHEET_PROTECT_PASSWORD", u"task") or u"")


def _is_named_task_sheet(sheet, doc=None):
    """Лист задач по префиксу «Задачи», кроме листа задач сотрудников из настроек."""
    try:
        name = unicode(getattr(sheet, "Name", u"") or u"").strip()
    except Exception:
        return False
    if not name:
        return False
    if _is_merge_source_sheet(sheet, doc=doc, name=name):
        return False
    prefix = unicode(getattr(_cfg, "TASK_SHEET_NAME_PREFIX", u"Задачи"))
    if not prefix:
        return False
    if hasattr(name, "casefold"):
        return name.casefold().startswith(prefix.casefold())
    return name.lower().startswith(prefix.lower())


def _merge_source_sheet_name(doc=None):
    """Имя листа задач сотрудников из настроек (не защищать, не искать GUID)."""
    try:
        from todo_task_settings_lib import get_merge_source_sheet

        name = unicode(get_merge_source_sheet(doc=doc) or u"").strip()
        if name:
            return name
    except Exception as exc:
        _log("merge source sheet name: %s" % exc)
    return unicode(
        getattr(_cfg, "DEFAULT_MERGE_SOURCE_SHEET", u"задачи_сотрудников") or u""
    )


def _is_merge_source_sheet(sheet=None, doc=None, name=None):
    if name is None:
        name = _sheet_display_name(sheet) if sheet is not None else u""
    name = unicode(name or u"").strip()
    src = _merge_source_sheet_name(doc)
    if not name or not src:
        return False
    if hasattr(name, "casefold"):
        return name.casefold() == src.casefold()
    return name.lower() == src.lower()


def _protect_unlock_error_text():
    return u"Не удалось снять защиту листа задач."


def _task_sheet_is_protected(sheet):
    if sheet is None:
        return False
    try:
        return bool(sheet.isProtected())
    except Exception:
        pass
    try:
        return bool(sheet.getPropertyValue("IsProtected"))
    except Exception:
        return False


def _task_sheet_unprotect_passwords(extra=None, doc=None):
    try:
        from todo_task_settings_lib import get_unprotect_passwords

        return list(get_unprotect_passwords(extra=extra, doc=doc) or [])
    except Exception:
        pwd = _task_sheet_protect_password(doc=doc)
        out = []
        if pwd not in out:
            out.append(pwd)
        fallback = unicode(getattr(_cfg, "TASK_SHEET_PROTECT_PASSWORD", u"task") or u"")
        if fallback and fallback not in out:
            out.append(fallback)
        if u"" not in out:
            out.append(u"")
        return out


def _task_sheet_unprotect(sheet, extra_passwords=None, doc=None):
    if sheet is None or not _task_sheet_is_protected(sheet):
        return True
    candidates = _task_sheet_unprotect_passwords(extra=extra_passwords, doc=doc)
    i = 0
    while i < len(candidates):
        try:
            sheet.unprotect(candidates[i])
            if not _task_sheet_is_protected(sheet):
                return True
        except Exception as exc:
            _log("unprotect task sheet: %s" % exc)
        i = i + 1
    return False


def _task_sheet_protect(sheet, password=None, doc=None):
    if sheet is None:
        return False
    if _is_merge_source_sheet(sheet, doc=doc):
        return True
    if _task_sheet_is_protected(sheet):
        return True
    if password is None:
        pwd = _task_sheet_protect_password(doc=doc)
    else:
        pwd = unicode(password if password is not None else u"")
    try:
        sheet.protect(pwd)
        return _task_sheet_is_protected(sheet)
    except Exception as exc:
        _log("protect task sheet: %s" % exc)
        return False


def _sheet_display_name(sheet):
    try:
        return unicode(getattr(sheet, "Name", u"") or u"")
    except Exception:
        return u"?"


def _reapply_protection_passwords(doc, old_pwd, new_pwd):
    """Re-protect all Задачи sheets and any already protected sheets."""
    ok_names = []
    fail_names = []
    if doc is None:
        return ok_names, fail_names
    try:
        sheets = doc.Sheets
        n = int(sheets.getCount())
    except Exception:
        return ok_names, fail_names
    extras = [old_pwd, new_pwd]
    i = 0
    while i < n:
        sh = None
        try:
            sh = sheets.getByIndex(i)
        except Exception:
            i = i + 1
            continue
        i = i + 1
        name = _sheet_display_name(sh)
        if _is_merge_source_sheet(sh, doc=doc, name=name):
            continue
        is_task = _is_named_task_sheet(sh, doc=doc)
        was_prot = _task_sheet_is_protected(sh)
        if not is_task and not was_prot:
            continue
        if was_prot:
            if not _task_sheet_unprotect(sh, extra_passwords=extras, doc=doc):
                fail_names.append(name)
                continue
        if not _task_sheet_protect(sh, password=new_pwd, doc=doc):
            if was_prot:
                _task_sheet_protect(sh, password=old_pwd, doc=doc)
            fail_names.append(name)
            continue
        ok_names.append(name)
    return ok_names, fail_names


class _TaskSheetProtectSession(object):
    """Снять защиту перед правками.

    Лист «Задачи…» — в finish защита ставится всегда (даже если не стояла).
    Лист задач сотрудников из настроек — защиту не ставим.
    Остальные листы — только восстановить, если защита уже была.
    """

    def __init__(self, sheet, doc=None):
        self.sheet = sheet
        self.doc = doc
        if self.doc is None:
            self.doc = _get_doc()
        self._was_protected = False
        self._active = False

    def begin(self):
        t0 = time.time()
        if self.sheet is None:
            return False
        self._was_protected = _task_sheet_is_protected(self.sheet)
        _log("protect.begin: sheet=%s was_protected=%s" % (_sheet_display_name(self.sheet), self._was_protected))
        if self._was_protected and not _task_sheet_unprotect(self.sheet, doc=self.doc):
            _log("protect.begin: FAIL unprotect (%s)" % _ms(time.time() - t0))
            return False
        self._active = True
        _log("protect.begin: ok (%s)" % _ms(time.time() - t0))
        return True

    def finish(self):
        if not self._active:
            return
        self._active = False
        t0 = time.time()
        if _is_merge_source_sheet(self.sheet, doc=self.doc):
            _log(
                "protect.finish: skip merge source %s"
                % _sheet_display_name(self.sheet)
            )
            return
        if _is_named_task_sheet(self.sheet, doc=self.doc) or self._was_protected:
            _task_sheet_protect(self.sheet, doc=self.doc)
            _log("protect.finish: sheet=%s (%s)" % (_sheet_display_name(self.sheet), _ms(time.time() - t0)))


def _prepare_sheet(doc):
    if doc is None:
        return None, None, u"Нет открытого документа Calc."
    try:
        if not doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            return None, None, u"Макрос работает только в Calc."
    except Exception:
        pass
    sheet = _active_sheet(doc)
    if sheet is None:
        return None, None, u"Не удалось получить активный лист."
    cols = _find_headers(sheet)
    if cols is None:
        return (
            None,
            None,
            u"На активном листе в строках 1–20 не найдены обязательные колонки:\n"
            u"«%s», «%s», «%s»."
            % (_cfg.COL_NUM, _cfg.COL_NAME, _cfg.COL_GUID),
        )
    missing_owner = []
    if u"created_by" not in cols:
        missing_owner.append(
            getattr(_cfg, "COL_CREATED_BY", u"Создано")
        )
    if u"modified_by" not in cols:
        missing_owner.append(
            getattr(_cfg, "COL_MODIFIED_BY", u"Отредактировано")
        )
    if missing_owner:
        return (
            None,
            None,
            u"На листе нет колонок автора: %s.\n"
            u"Добавьте их вручную после «%s» или нажмите «Применить размеры» "
            u"в настройках задач (вкладка «Размеры полей»)."
            % (
                u", ".join([u"«%s»" % x for x in missing_owner]),
                _cfg.COL_GUID,
            ),
        )
    return sheet, cols, u""


def _select_task_cell(doc, sheet, cols, row):
    try:
        cell = sheet.getCellByPosition(cols[u"name"], row)
        doc.getCurrentController().select(cell)
    except Exception:
        pass


def _place_new_task_row(sheet, cols, copy_fields=None, doc=None):
    """
    Новая строка после последней задачи; если задач нет — первая после заголовка.
    Стили заголовка не копируются. copy_fields — поля исходной задачи или None.
    Без autofit/check_numbers — их делает _edit_task_row один раз после диалога.
    Возвращает new_row.
    """
    tasks = _iter_task_rows(sheet, cols)
    if tasks:
        last_task = tasks[-1]
        new_row = last_task + 1
        src_fmt = last_task
        next_num = len(tasks) + 1
        _copy_row_formats(sheet, cols, src_fmt, new_row)
    else:
        new_row = _header_row(cols) + 1
        next_num = 1

    if copy_fields:
        fields = _prepare_copy_fields(copy_fields, doc=doc)
        write_task_fields(sheet, cols, new_row, fields, clear_due=True, doc=None)
        if u"due" in cols:
            _clear_cell_value(sheet.getCellByPosition(cols[u"due"], new_row))
        if u"status" in cols and not fields.get(u"status"):
            _clear_cell_value(sheet.getCellByPosition(cols[u"status"], new_row))

    _set_cell_text(sheet.getCellByPosition(cols[u"guid"], new_row), _new_guid())
    _ensure_row_number(sheet, cols, new_row, next_num)
    return new_row


def todo_task_add(doc=None):
    """Добавить строку задачи сразу после последней задачи (или после заголовка)."""
    tm = _StageTimer("todo_task_add")
    doc = _get_doc(doc)
    sheet, cols, err = _prepare_sheet(doc)
    if err:
        _msg(doc, err, box_type=WARNINGBOX)
        return False

    guard = _TaskSheetProtectSession(sheet, doc=doc)
    if not guard.begin():
        _msg(
            doc,
            _protect_unlock_error_text(),
            box_type=WARNINGBOX,
        )
        return False
    try:
        new_row = _place_new_task_row(sheet, cols, copy_fields=None, doc=doc)
        _select_task_cell(doc, sheet, cols, new_row)
        _edit_task_row(doc, sheet, cols, new_row, use_default_due=True)
        return True
    finally:
        guard.finish()


def todo_task_delete(doc=None):
    """Удалить текущую строку-задачу."""
    tm = _StageTimer("todo_task_delete")
    doc = _get_doc(doc)
    sheet, cols, err = _prepare_sheet(doc)
    if err:
        _msg(doc, err, box_type=WARNINGBOX)
        return False

    row = _cursor_row(doc, sheet)
    if row is None:
        _msg(doc, u"Выберите ячейку в строке задачи.", box_type=WARNINGBOX)
        return False
    if not _is_managed_row(sheet, cols, row):
        _msg(
            doc,
            u"Строка %s не является задачей (нет «Наименования» и GUID)."
            % (row + 1),
            box_type=WARNINGBOX,
        )
        return False

    fields = read_task_fields(sheet, cols, row)
    label = unicode(fields.get(u"name") or fields.get(u"guid") or u"")
    if len(label) > 80:
        label = label[:77] + u"…"
    if not _confirm(
        doc,
        u"Удалить строку %s?\n%s" % (row + 1, label),
        title=u"Удаление задачи",
    ):
        return False

    if not _archive_deleted_task(doc, sheet, cols, row):
        _msg(
            doc,
            u"Не удалось сохранить задачу в архив «%s»."
            % getattr(_cfg, "DELETED_SHEET_NAME", u"__Удаленные_Задачи"),
            box_type=WARNINGBOX,
        )
        return False

    task_guid = unicode(fields.get(u"guid") or u"").strip()
    if task_guid and not _is_merge_source_sheet(sheet, doc=doc):
        try:
            from todo_task_plan_lib import archive_plan_by_guid, delete_plan_by_guid

            try:
                src_name = unicode(sheet.Name)
            except Exception:
                src_name = u""
            if not archive_plan_by_guid(
                doc,
                task_guid,
                source_sheet_name=src_name,
                task_name=unicode(fields.get(u"name") or u""),
            ):
                _msg(
                    doc,
                    u"Не удалось архивировать план задачи.",
                    box_type=WARNINGBOX,
                )
                return False
            if not delete_plan_by_guid(doc, task_guid):
                _msg(
                    doc,
                    u"Не удалось удалить план задачи.",
                    box_type=WARNINGBOX,
                )
                return False
        except Exception as exc:
            _msg(
                doc,
                u"Ошибка архива плана: %s" % exc,
                box_type=WARNINGBOX,
            )
            return False

    guard = _TaskSheetProtectSession(sheet, doc=doc)
    if not guard.begin():
        _msg(
            doc,
            _protect_unlock_error_text(),
            box_type=WARNINGBOX,
        )
        return False
    try:
        try:
            sheet.getRows().removeByIndex(int(row), 1)
        except Exception as exc:
            _msg(doc, u"Не удалось удалить строку: %s" % exc, box_type=WARNINGBOX)
            return False

        _restore_numbered_row_after_delete(sheet, cols)
        # Полный autofit листа не нужен: меняется удаляемая зона + новая пустая строка в конце.
        last_num = _last_numbered_row(sheet, cols)
        fit_rows = []
        if last_num is not None:
            fit_rows.append(int(last_num))
        if int(row) not in fit_rows:
            # строка на месте удаления (бывший следующий ряд сдвинулся вверх)
            fit_rows.append(int(row))
        _autofit_task_rows(sheet, fit_rows)
        _msg(
            doc,
            u"Строка задачи удалена. В конец добавлена пустая строка (форматы + № п/п).",
        )
        return True
    finally:
        guard.finish()


def todo_task_copy(doc=None):
    """Скопировать текущую задачу; если задач нет — как добавление новой строки."""
    tm = _StageTimer("todo_task_copy")
    doc = _get_doc(doc)
    sheet, cols, err = _prepare_sheet(doc)
    if err:
        _msg(doc, err, box_type=WARNINGBOX)
        return False

    tasks = _iter_task_rows(sheet, cols)
    copy_fields = None
    if tasks:
        row = _cursor_row(doc, sheet)
        if row is None or not _is_managed_row(sheet, cols, row):
            _msg(
                doc,
                u"Выберите ячейку в строке задачи для копирования.",
                box_type=WARNINGBOX,
            )
            return False
        copy_fields = read_task_fields(sheet, cols, row)

    guard = _TaskSheetProtectSession(sheet, doc=doc)
    if not guard.begin():
        _msg(
            doc,
            _protect_unlock_error_text(),
            box_type=WARNINGBOX,
        )
        return False
    try:
        new_row = _place_new_task_row(sheet, cols, copy_fields=copy_fields, doc=doc)
        _select_task_cell(doc, sheet, cols, new_row)
        _edit_task_row(doc, sheet, cols, new_row, use_default_due=True)
        return True
    finally:
        guard.finish()


def todo_task_edit(doc=None):
    """Открыть диалог редактирования текущей строки-задачи."""
    tm = _StageTimer("todo_task_edit")
    doc = _get_doc(doc)
    sheet, cols, err = _prepare_sheet(doc)
    if err:
        _msg(doc, err, box_type=WARNINGBOX)
        return False

    row = _cursor_row(doc, sheet)
    if row is None or not _is_managed_row(sheet, cols, row):
        _msg(
            doc,
            u"Выберите ячейку в строке задачи для редактирования.",
            box_type=WARNINGBOX,
        )
        return False

    guard = _TaskSheetProtectSession(sheet, doc=doc)
    if not guard.begin():
        _msg(
            doc,
            _protect_unlock_error_text(),
            box_type=WARNINGBOX,
        )
        return False
    try:
        return _edit_task_row(doc, sheet, cols, row)
    finally:
        guard.finish()


def _normalize_priority_match_key(text):
    """Ключ сравнения приоритета: нижний регистр, без пробелов, тире унифицированы."""
    s = unicode(text or u"").casefold()
    for ch in (u"\r", u"\n", u"\t", u" ", u"\u00a0"):
        s = s.replace(ch, u"")
    for ch in (u"–", u"—", u"−", u"‐", u"‑", u"‒", u"―"):
        s = s.replace(ch, u"-")
    return s


def _is_urgent_priority_text(text):
    key = getattr(_cfg, "COLORIZE_PRIORITY_URGENT_KEY", u"1-важноисрочно")
    return _normalize_priority_match_key(text) == unicode(key).casefold()


def _apply_priority_urgent_cell_style(cell):
    if cell is None:
        return
    fill = int(getattr(_cfg, "COLORIZE_PRIORITY_URGENT_FILL", 0xFFFF00))
    border_color = int(getattr(_cfg, "COLORIZE_PRIORITY_URGENT_BORDER_COLOR", 0xFF0000))
    border_width = int(getattr(_cfg, "COLORIZE_PRIORITY_URGENT_BORDER_WIDTH", 25))
    try:
        cell.CellBackColor = fill
    except Exception:
        pass
    try:
        from libre_macros_lib import _lm_pp_set_cell_side_border

        for side in (u"top", u"bottom", u"left", u"right"):
            _lm_pp_set_cell_side_border(cell, side, border_color, border_width)
    except Exception:
        pass


def _copy_sheet_row(sheet, src_row, dest_row, end_col):
    """Скопировать строку (значения + форматы) через copyRange."""
    src = sheet.getCellRangeByPosition(0, int(src_row), int(end_col), int(src_row))
    dest = sheet.getCellByPosition(0, int(dest_row))
    sheet.copyRange(dest.CellAddress, src.RangeAddress)


def _sort_key_assignee(fields):
    text = unicode(fields.get(u"assignee") or u"").strip()
    if not text:
        return (1, u"")
    key = text.casefold() if hasattr(text, "casefold") else text.lower()
    return (0, key)


def _sort_key_due(fields, new_first=False):
    parts = fields.get(u"due_parts")
    if not parts:
        return (1, 0)
    try:
        from datetime import date

        d = date(int(parts[2]), int(parts[1]), int(parts[0]))
        ord_val = int(d.toordinal())
    except Exception:
        return (1, 0)
    if new_first:
        return (0, -ord_val)
    return (0, ord_val)


def _sort_key_overdue(fields):
    status = fields.get(u"status") or u""
    due_parts = fields.get(u"due_parts")
    if _is_done_task_status(status):
        return (4, 0, 0)
    if not due_parts:
        return (3, 0, 0)
    try:
        from datetime import date

        d = date(int(due_parts[2]), int(due_parts[1]), int(due_parts[0]))
        today = date.today()
        delta = (d - today).days
        if delta < 0:
            return (0, 0, -delta)
        if delta < 3:
            return (1, delta, 0)
        return (2, delta, 0)
    except Exception:
        return (3, 0, 0)


def _sort_key_priority(fields):
    text = unicode(fields.get(u"priority") or u"").strip()
    num = _parse_int_num(text)
    if num is None:
        return (1, 9999, text.casefold() if hasattr(text, "casefold") else text.lower())
    return (0, int(num), u"")


def _task_sort_key(fields, mode, due_direction):
    mode = unicode(mode or u"").strip()
    if mode == getattr(_cfg, "SORT_MODE_DUE", u"due"):
        new_first = unicode(due_direction or u"") == getattr(
            _cfg, "SORT_DUE_DIR_NEW_FIRST", u"new_first"
        )
        return _sort_key_due(fields, new_first=new_first)
    if mode == getattr(_cfg, "SORT_MODE_OVERDUE", u"overdue"):
        return _sort_key_overdue(fields)
    if mode == getattr(_cfg, "SORT_MODE_PRIORITY", u"priority"):
        return _sort_key_priority(fields)
    return _sort_key_assignee(fields)


def _renumber_task_rows(sheet, cols, task_rows):
    """Rewrite No. column as 1..N for current task row order."""
    _write_row_numbers(sheet, cols[u"num"], task_rows, start_num=1)


def _reorder_task_rows(sheet, cols, sorted_src_indices, task_rows):
    """Reorder task rows: sorted_src_indices[i] -> task_rows[i]."""
    n = len(task_rows)
    if n <= 1:
        return
    tm = _StageTimer("reorder_rows")
    max_col = _sort_copy_end_col(sheet, cols)
    tm.step("copy_end_col", "max_col=%s" % max_col)
    temp_base = _used_last_row(sheet, min_cols=max_col + 1) + 2
    rows = sheet.getRows()
    try:
        rows.insertByIndex(temp_base, n)
    except Exception as exc:
        raise RuntimeError(u"sort buffer insert failed: %s" % exc)
    tm.step("insert_temp", "n=%s temp_base=%s" % (n, temp_base))
    j = 0
    while j < n:
        _copy_sheet_row(sheet, task_rows[sorted_src_indices[j]], temp_base + j, max_col)
        j = j + 1
    tm.step("copy_to_temp", "rows=%s" % n)
    j = 0
    while j < n:
        _copy_sheet_row(sheet, temp_base + j, task_rows[j], max_col)
        j = j + 1
    tm.step("copy_back", "rows=%s" % n)
    try:
        rows.removeByIndex(temp_base, n)
    except Exception:
        pass
    tm.done("DONE")


def _sort_tasks(sheet, cols, mode, due_direction):
    tm = _StageTimer("sort_tasks")
    task_rows = _iter_task_rows(sheet, cols)
    tm.step("iter_task_rows", "n=%s" % len(task_rows))
    if len(task_rows) < 2:
        tm.done("skip")
        return 0, u""
    snapshots = []
    i = 0
    while i < len(task_rows):
        row = task_rows[i]
        fields = read_task_fields(sheet, cols, row)
        key = _task_sort_key(fields, mode, due_direction)
        snapshots.append((key, i))
        i = i + 1
    snapshots.sort(key=lambda item: item[0])
    tm.step("build_keys_sort", "n=%s" % len(snapshots))
    sorted_indices = []
    i = 0
    while i < len(snapshots):
        sorted_indices.append(snapshots[i][1])
        i = i + 1
    _reorder_task_rows(sheet, cols, sorted_indices, task_rows)
    tm.step("reorder")
    _renumber_task_rows(sheet, cols, task_rows)
    tm.step("renumber")
    tm.done("DONE", "tasks=%s" % len(task_rows))
    return len(task_rows), u""


def todo_task_sort(doc=None):
    """Отсортировать строки-задачи (диалог выбора критерия)."""
    tm = _StageTimer("todo_task_sort")
    doc = _get_doc(doc)
    sheet, cols, err = _prepare_sheet(doc)
    if err:
        _msg(doc, err, box_type=WARNINGBOX)
        return False

    from todo_task_dialog_lib import show_todo_task_sort_dialog
    from todo_task_settings_lib import (
        load_todo_task_settings,
        normalize_sort_settings,
        save_todo_task_settings,
    )

    settings = load_todo_task_settings(doc=doc)
    try:
        spec = show_todo_task_sort_dialog(
            doc=doc,
            initial=settings.get(u"sort"),
        )
    except Exception as exc:
        _msg(doc, u"Диалог сортировки: %s" % exc, box_type=WARNINGBOX)
        return False
    if spec is None:
        tm.done("cancel")
        return False

    spec = normalize_sort_settings(spec)
    settings[u"sort"] = spec
    save_err = save_todo_task_settings(settings, doc=doc)
    if save_err:
        _log(save_err)

    tasks = _iter_task_rows(sheet, cols)
    if len(tasks) < 2:
        _msg(doc, u"На листе меньше двух задач — сортировка не нужна.", box_type=WARNINGBOX)
        tm.done("too_few")
        return False

    guard = _TaskSheetProtectSession(sheet, doc=doc)
    if not guard.begin():
        _msg(
            doc,
            _protect_unlock_error_text(),
            box_type=WARNINGBOX,
        )
        tm.done("protect_fail")
        return False
    try:
        n, note = _sort_tasks(
            sheet,
            cols,
            spec.get(u"mode"),
            spec.get(u"due_direction"),
        )
        tm.step("sort_tasks", "n=%s" % n)
        _autofit_sheet_row_heights(sheet, cols)
        tm.step("autofit")
        try:
            apply_todo_column_widths(doc, settings, sheets=[sheet])
            tm.step("column_widths")
        except Exception as werr:
            _log("sort apply widths: %s" % werr)
        mode_labels = dict(getattr(_cfg, "SORT_MODE_CHOICES", ()))
        mode_code = unicode(spec.get(u"mode") or u"")
        mode_label = mode_labels.get(mode_code, mode_code)
        _msg(doc, u"Отсортировано задач: %s\n(%s)" % (n, mode_label))
        tm.done("DONE", "tasks=%s" % n)
        return True
    except Exception as exc:
        _msg(doc, u"Ошибка сортировки: %s" % exc, box_type=WARNINGBOX)
        tm.done("error")
        return False
    finally:
        guard.finish()


def _compute_overdue_code(status_text, due_parts):
    """
    Логика «Просрочки» в памяти (без формул в листе):
    100 — выполнено; 10 — срок прошёл; 1 — до срока < 3 дней; 0 — иначе; None — нет срока.
    """
    if _is_done_task_status(status_text):
        return int(getattr(_cfg, "COLORIZE_CODE_DONE", 100))
    if not due_parts:
        return None
    try:
        from datetime import date

        d = date(int(due_parts[2]), int(due_parts[1]), int(due_parts[0]))
        today = date.today()
        if d < today:
            return int(getattr(_cfg, "COLORIZE_CODE_OVERDUE", 10))
        if (d - today).days < 3:
            return int(getattr(_cfg, "COLORIZE_CODE_SOON", 1))
        return int(getattr(_cfg, "COLORIZE_CODE_NORMAL", 0))
    except Exception:
        return None


def _todo_task_colorize_style(code):
    """Стиль строки по коду просрочки (заливка / жирность)."""
    try:
        from libre_macros_lib import _lm_pp_zebra_resolve_color_token
    except Exception:

        def _lm_pp_zebra_resolve_color_token(name):
            return None

    def _fill(token):
        c = _lm_pp_zebra_resolve_color_token(token)
        if c is not None:
            return int(c)
        if unicode(token).casefold() in (u"white", u"белый", u"ffffff"):
            return 0xFFFFFF
        return None

    code = int(code) if code is not None else 0
    done = int(getattr(_cfg, "COLORIZE_CODE_DONE", 100))
    overdue = int(getattr(_cfg, "COLORIZE_CODE_OVERDUE", 10))
    soon = int(getattr(_cfg, "COLORIZE_CODE_SOON", 1))
    if code == done:
        return {
            u"fill_color": _fill(u"зеленый"),
            u"font_color_auto": True,
            u"bold": False,
            u"italic": False,
        }
    if code == overdue:
        return {
            u"fill_color": _fill(u"светло_красный"),
            u"font_color_auto": True,
            u"bold": True,
            u"italic": False,
        }
    if code == soon:
        return {
            u"fill_color": _fill(u"желтый"),
            u"font_color_auto": True,
            u"bold": False,
            u"italic": False,
        }
    return {
        u"fill_color": _fill(u"white"),
        u"font_color_auto": True,
        u"bold": False,
        u"italic": False,
    }


def _apply_colorize_row_style(sheet, row, max_col, spec):
    if sheet is None or spec is None:
        return
    try:
        from libre_macros_lib import _lm_pp_apply_threshold_cell_style
    except Exception:
        return
    try:
        row_range = sheet.getCellRangeByPosition(0, int(row), int(max_col), int(row))
        _lm_pp_apply_threshold_cell_style(row_range, spec)
        return
    except Exception:
        pass
    c = 0
    while c <= int(max_col):
        try:
            _lm_pp_apply_threshold_cell_style(
                sheet.getCellByPosition(c, int(row)), spec
            )
        except Exception:
            pass
        c = c + 1


def _sheet_name_matches_pattern(sheet_name, pattern):
    name = unicode(sheet_name or u"").strip()
    pat = unicode(pattern or u"").strip()
    if pat == u"":
        return True
    try:
        return fnmatch.fnmatchcase(name, pat)
    except Exception:
        return name == pat


def _iter_sheets_matching_pattern(doc, pattern, exclude_merge_source=True):
    out = []
    if doc is None:
        return out
    try:
        sheets = doc.Sheets
        n = int(sheets.getCount())
    except Exception:
        return out
    i = 0
    while i < n:
        try:
            sh = sheets.getByIndex(i)
            nm = unicode(getattr(sh, "Name", u"") or u"")
            if _sheet_name_matches_pattern(nm, pattern):
                if exclude_merge_source and _is_merge_source_sheet(
                    sh, doc=doc, name=nm
                ):
                    i = i + 1
                    continue
                out.append(sh)
        except Exception:
            pass
        i = i + 1
    return out


def _colorize_tasks_on_sheet(doc, sheet):
    """
    Раскрасить задачи на одном листе.
    Возвращает (число задач, сообщение об ошибке или пусто).
    """
    cols = _find_headers(sheet)
    if cols is None:
        return 0, u"нет таблицы задач"
    if u"due" not in cols or u"status" not in cols:
        return (
            0,
            u"нет колонок «%s» / «%s»"
            % (_cfg.COL_DUE, _cfg.COL_STATUS),
        )

    guard = _TaskSheetProtectSession(sheet, doc=doc)
    if not guard.begin():
        return 0, u"не удалось снять защиту листа"
    try:
        tm = _StageTimer("colorize_sheet:%s" % _sheet_display_name(sheet))
        check_and_fix_task_numbers(sheet, cols, fix=True)
        tm.step("check_numbers")
        max_col = _table_max_col(cols)
        tasks = _iter_task_rows(sheet, cols)
        i = 0
        while i < len(tasks):
            row = tasks[i]
            status = _cell_text(
                sheet.getCellByPosition(cols[u"status"], row)
            )
            due_parts = _read_due_from_cell(
                sheet.getCellByPosition(cols[u"due"], row)
            )
            code = _compute_overdue_code(status, due_parts)
            if code is None:
                style = _todo_task_colorize_style(
                    int(getattr(_cfg, "COLORIZE_CODE_NORMAL", 0))
                )
            else:
                style = _todo_task_colorize_style(code)
            _apply_colorize_row_style(sheet, row, max_col, style)
            if u"priority" in cols:
                pri_cell = sheet.getCellByPosition(cols[u"priority"], row)
                if (
                    _is_urgent_priority_text(_cell_text(pri_cell))
                    and not _is_done_task_status(status)
                ):
                    _apply_priority_urgent_cell_style(pri_cell)
            i = i + 1
        tm.step("colorize_loop", "tasks=%s" % len(tasks))
        _autofit_sheet_row_heights(sheet, cols)
        tm.done("DONE")
        return len(tasks), u""
    except Exception as exc:
        return 0, unicode(exc)
    finally:
        guard.finish()


def todo_task_colorize(doc=None):
    """
    Раскрасить строки-задачи по просрочке (диалог: текущий лист или все по шаблону).
    """
    tm = _StageTimer("todo_task_colorize")
    doc = _get_doc(doc)
    if doc is None:
        _msg(doc, u"Нет открытого документа Calc.", box_type=WARNINGBOX)
        return False
    try:
        if not doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            _msg(doc, u"Макрос работает только в Calc.", box_type=WARNINGBOX)
            return False
    except Exception:
        pass

    from todo_task_dialog_lib import show_todo_task_colorize_dialog
    from todo_task_settings_lib import (
        load_todo_task_settings,
        normalize_colorize_settings,
        save_todo_task_settings,
    )

    settings = load_todo_task_settings(doc=doc)
    try:
        spec = show_todo_task_colorize_dialog(
            doc=doc,
            initial=settings.get(u"colorize"),
        )
    except Exception as exc:
        _msg(doc, u"Диалог раскраски: %s" % exc, box_type=WARNINGBOX)
        return False
    if spec is None:
        return False

    spec = normalize_colorize_settings(spec)
    settings[u"colorize"] = spec
    save_err = save_todo_task_settings(settings, doc=doc)
    if save_err:
        _log(save_err)

    sheets = []
    if spec.get(u"all_sheets"):
        pattern = unicode(spec.get(u"sheet_pattern") or u"Задачи*")
        sheets = _iter_sheets_matching_pattern(doc, pattern)
        if not sheets:
            _msg(
                doc,
                u"Нет листов, подходящих под шаблон «%s»." % pattern,
                box_type=WARNINGBOX,
            )
            return False
    else:
        sheet = _active_sheet(doc)
        if sheet is None:
            _msg(doc, u"Не удалось получить активный лист.", box_type=WARNINGBOX)
            return False
        cols = _find_headers(sheet)
        if cols is None:
            _msg(
                doc,
                u"На активном листе в строках 1–20 не найдены обязательные колонки:\n"
                u"«%s», «%s», «%s»."
                % (_cfg.COL_NUM, _cfg.COL_NAME, _cfg.COL_GUID),
                box_type=WARNINGBOX,
            )
            return False
        sheets = [sheet]

    total_tasks = 0
    ok_sheets = 0
    skipped = []
    si = 0
    while si < len(sheets):
        sheet = sheets[si]
        sheet_name = u"?"
        try:
            sheet_name = unicode(getattr(sheet, "Name", u"?") or u"?")
        except Exception:
            pass
        n, note = _colorize_tasks_on_sheet(doc, sheet)
        if n > 0:
            total_tasks = total_tasks + n
            ok_sheets = ok_sheets + 1
        elif note:
            skipped.append(u"%s — %s" % (sheet_name, note))
        si = si + 1

    if ok_sheets == 0:
        if skipped:
            _msg(
                doc,
                u"Раскраска не выполнена:\n" + u"\n".join(skipped[:12]),
                box_type=WARNINGBOX,
            )
        else:
            _msg(doc, u"На выбранных листах нет задач.", box_type=WARNINGBOX)
        return False

    if spec.get(u"all_sheets"):
        msg = u"Раскрашено задач: %s\nЛистов: %s" % (total_tasks, ok_sheets)
        if skipped:
            msg = msg + u"\n\nПропущено:\n" + u"\n".join(skipped[:8])
        _msg(doc, msg)
    else:
        _msg(doc, u"Раскрашено задач: %s" % total_tasks)
    return True


def _iter_named_task_sheets(doc, exclude_name=None):
    """Листы, имя которых начинается с «Задачи», кроме exclude_name."""
    out = []
    if doc is None:
        return out
    excl = unicode(exclude_name or u"").strip()
    if hasattr(excl, "casefold"):
        excl_cf = excl.casefold()
    else:
        excl_cf = excl.lower()
    try:
        sheets = doc.Sheets
        n = int(sheets.getCount())
    except Exception:
        return out
    i = 0
    while i < n:
        try:
            sh = sheets.getByIndex(i)
            if not _is_named_task_sheet(sh, doc=doc):
                i = i + 1
                continue
            nm = unicode(getattr(sh, "Name", u"") or u"")
            nm_cf = nm.casefold() if hasattr(nm, "casefold") else nm.lower()
            if excl and nm_cf == excl_cf:
                i = i + 1
                continue
            out.append(sh)
        except Exception:
            pass
        i = i + 1
    return out


def _copy_cell_content(src, dst, values_only=True):
    """Скопировать ячейку: значение без оформления источника или полностью с форматами."""
    if src is None or dst is None:
        return
    if values_only:
        _copy_cell_value_only(src, dst)
        return
    props = getattr(_cfg, "FORMAT_COPY_PROPS", ())
    pi = 0
    while pi < len(props):
        prop = props[pi]
        try:
            setattr(dst, prop, getattr(src, prop))
        except Exception:
            pass
        pi = pi + 1
    try:
        dst.setFormula(src.getFormula())
        return
    except Exception:
        pass
    _copy_cell_value_only(src, dst)


def _header_norm_first_col(items):
    """norm(title) → первый индекс колонки."""
    out = {}
    i = 0
    while i < len(items):
        col, raw = items[i]
        i = i + 1
        n = _norm_header(raw)
        if n and n not in out:
            out[n] = int(col)
    return out


def _ensure_move_header_pairs( src_sheet, dest_sheet, dest_cols, add_columns=False, values_only=True, src_cols=None ):
    """
    Пары (src_col, dest_col) по заголовкам (trim + lower).
    add_columns=False — только существующие колонки приёмника.
    add_columns=True — недостающие заголовки дописываются в конец шапки приёмника.
    Возвращает (dest_cols, pairs).
    """
    src_items = _list_sheet_header_titles(src_sheet)
    dest_items = _list_sheet_header_titles(dest_sheet)
    dest_by_norm = _header_norm_first_col(dest_items)
    dest_hdr = _header_row(dest_cols)
    last_dest = -1
    di = 0
    while di < len(dest_items):
        col = int(dest_items[di][0])
        if col > last_dest:
            last_dest = col
        di = di + 1
    if last_dest < 0:
        last_dest = max(_table_max_col(dest_cols), 0)

    pairs = []
    used_dest = {}
    missing = []
    si = 0
    while si < len(src_items):
        src_col, raw = src_items[si]
        si = si + 1
        n = _norm_header(raw)
        if not n:
            continue
        if n in dest_by_norm:
            dest_col = int(dest_by_norm[n])
            if dest_col in used_dest:
                continue
            used_dest[dest_col] = True
            pairs.append((int(src_col), dest_col))
        else:
            missing.append((int(src_col), unicode(raw or u"")))

    if add_columns and missing:
        insert_at = int(last_dest) + 1
        n_ins = len(missing)
        try:
            dest_sheet.getColumns().insertByIndex(insert_at, n_ins)
        except Exception as exc:
            _log("move add columns: %s" % exc)
        mi = 0
        while mi < n_ins:
            src_col, raw = missing[mi]
            dest_col = insert_at + mi
            try:
                dest_hdr_cell = dest_sheet.getCellByPosition(dest_col, dest_hdr)
                src_hdr = dest_hdr
                if src_cols is not None:
                    src_hdr = _header_row(src_cols)
                src_hdr_cell = src_sheet.getCellByPosition(src_col, src_hdr)
                _copy_cell_content(
                    src_hdr_cell, dest_hdr_cell, values_only=values_only
                )
                if not _cell_text(dest_hdr_cell):
                    _set_cell_text(dest_hdr_cell, raw)
            except Exception as exc:
                _log("move write header %s: %s" % (raw, exc))
            n = _norm_header(raw)
            if n and n not in dest_by_norm:
                dest_by_norm[n] = dest_col
            if dest_col not in used_dest:
                used_dest[dest_col] = True
                pairs.append((src_col, dest_col))
            mi = mi + 1
        refreshed = _find_headers(dest_sheet)
        if refreshed is not None:
            dest_cols = refreshed
    return dest_cols, pairs


def _copy_move_row_by_pairs( src_sheet, src_row, dest_sheet, dest_row, pairs, values_only=True, skip_dest_cols=None ):
    """Перенести строку задачи поячеечно по парам колонок."""
    src_row = int(src_row)
    dest_row = int(dest_row)
    skip = {}
    sc = list(skip_dest_cols or [])
    si = 0
    while si < len(sc):
        try:
            skip[int(sc[si])] = True
        except (TypeError, ValueError):
            pass
        si = si + 1
    i = 0
    while i < len(pairs):
        src_col, dest_col = pairs[i]
        i = i + 1
        try:
            dc = int(dest_col)
        except (TypeError, ValueError):
            continue
        if dc in skip:
            continue
        try:
            _copy_cell_content(
                src_sheet.getCellByPosition(int(src_col), src_row),
                dest_sheet.getCellByPosition(dc, dest_row),
                values_only=values_only,
            )
        except Exception as exc:
            _log("move cell %s->%s: %s" % (src_col, dest_col, exc))
    return True


def _guid_key(text):
    g = unicode(text or u"").strip()
    if not g:
        return u""
    return g.casefold() if hasattr(g, "casefold") else g.lower()


def _sheet_name_key(name):
    n = unicode(name or u"").strip()
    if not n:
        return u""
    return n.casefold() if hasattr(n, "casefold") else n.lower()


def _task_sheet_pattern_for_doc(doc):
    """Шаблон листов задач из настроек (layout.task_sheet_pattern)."""
    try:
        from todo_task_settings_lib import (
            load_todo_task_settings,
            normalize_layout_settings,
        )

        settings = load_todo_task_settings(doc=doc)
        lay = normalize_layout_settings((settings or {}).get(u"layout"))
        pat = unicode(lay.get(u"task_sheet_pattern") or u"").strip()
        if pat:
            return pat
    except Exception as exc:
        _log("task sheet pattern: %s" % exc)
    return unicode(getattr(_cfg, "MOVE_SHEET_PATTERN", u"Задачи*") or u"Задачи*")


def _iter_move_dest_sheets(doc, exclude_name=None, pattern=None):
    """Листы задач по шаблону настроек (с таблицей), кроме exclude_name."""
    out = []
    if pattern is None:
        pattern = _task_sheet_pattern_for_doc(doc)
    excl = _sheet_name_key(exclude_name)
    sheets = _iter_sheets_matching_pattern(doc, pattern)
    i = 0
    while i < len(sheets):
        sh = sheets[i]
        i = i + 1
        try:
            nm = unicode(getattr(sh, "Name", u"") or u"")
        except Exception:
            continue
        if excl and _sheet_name_key(nm) == excl:
            continue
        if _is_merge_source_sheet(sh, doc=doc, name=nm):
            continue
        if _find_headers(sh) is None:
            continue
        out.append(sh)
    return out


def _default_dest_name_by_category(src_sheet, src_cols, task_rows, dest_ready):
    """
    Если у выбранных задач есть «Категория» и среди приёмников есть лист
    с той же категорией в B1 — вернуть имя этого листа.
    """
    if not dest_ready or src_cols is None or u"category" not in src_cols:
        return u""
    cat = u""
    i = 0
    while i < len(task_rows):
        fields = read_task_fields(src_sheet, src_cols, int(task_rows[i]))
        i = i + 1
        cat = unicode(fields.get(u"category") or u"").strip()
        if cat:
            break
    if not cat:
        return u""
    cat_cf = _sheet_name_key(cat)
    di = 0
    while di < len(dest_ready):
        sh = dest_ready[di]
        di = di + 1
        b1 = unicode(_sheet_category_from_b1(sh) or u"").strip()
        if b1 and _sheet_name_key(b1) == cat_cf:
            try:
                return unicode(sh.Name)
            except Exception:
                return u""
    return u""


def _collect_move_guid_hits(doc, pattern, exclude_sheet_name, prefer_sheet_name):
    """
    guid(casefold) -> {sheet, cols, row, sheet_name, task_name}.
    exclude — лист-источник переноса; prefer — лист назначения (если GUID там, берём его).
    """
    index = {}
    excl = _sheet_name_key(exclude_sheet_name)
    prefer = _sheet_name_key(prefer_sheet_name)
    sheets = _iter_sheets_matching_pattern(doc, pattern)
    # сначала prefer, потом остальные — чтобы совпадение на приёмнике победило
    ordered = []
    rest = []
    si = 0
    while si < len(sheets):
        sh = sheets[si]
        si = si + 1
        try:
            nm = unicode(getattr(sh, "Name", u"") or u"")
        except Exception:
            continue
        nk = _sheet_name_key(nm)
        if excl and nk == excl:
            continue
        if _is_merge_source_sheet(sh, doc=doc, name=nm):
            continue
        if prefer and nk == prefer:
            ordered.append(sh)
        else:
            rest.append(sh)
    ordered.extend(rest)
    oi = 0
    while oi < len(ordered):
        sh = ordered[oi]
        oi = oi + 1
        cols = _find_headers(sh)
        if cols is None or u"guid" not in cols:
            continue
        try:
            sheet_name = unicode(getattr(sh, "Name", u"") or u"")
        except Exception:
            sheet_name = u""
        rows = _iter_task_rows(sh, cols)
        ri = 0
        while ri < len(rows):
            row = int(rows[ri])
            ri = ri + 1
            guid = _cell_text(sh.getCellByPosition(cols[u"guid"], row)).strip()
            key = _guid_key(guid)
            if not key or key in index:
                continue
            fields = read_task_fields(sh, cols, row)
            index[key] = {
                u"sheet": sh,
                u"cols": cols,
                u"row": row,
                u"sheet_name": sheet_name,
                u"task_name": fields.get(u"name") or u"",
            }
    return index


def _merge_move_row_comments( src_sheet, src_row, dest_sheet, dest_row, src_cols, dest_cols ):
    """Накопить комментарии как при объединении задач сотрудников."""
    dest_cmt = _comment_cell(dest_sheet, dest_cols, dest_row)
    src_cmt = _comment_cell(src_sheet, src_cols, src_row)
    local_cmt = _cell_text(dest_cmt) if dest_cmt is not None else u""
    incoming_cmt = _cell_text(src_cmt) if src_cmt is not None else u""
    inc_stamp = _timestamp_display_from_cell(
        _modified_cell(src_sheet, src_cols, src_row)
    )
    loc_stamp = _timestamp_display_from_cell(
        _modified_cell(dest_sheet, dest_cols, dest_row)
    )
    merged = _accumulate_comment_text(
        local_cmt, incoming_cmt, inc_stamp, loc_stamp
    )
    if dest_cmt is not None:
        _set_cell_text(dest_cmt, merged)
    return True


def _resolve_move_guid_conflicts( doc, src_sheet, src_cols, dest_name, task_rows, guid_hits ):
    """
    Спросить, что делать при совпадении GUID на любом листе задач.
    Возвращает (decisions, aborted).
    decisions: src_row -> {action, hit} (hit — запись из guid_hits или None).
    """
    from todo_task_dialog_lib import show_todo_task_move_guid_conflict_dialog

    conflicts = []
    i = 0
    while i < len(task_rows):
        src_row = int(task_rows[i])
        i = i + 1
        if src_cols is None or u"guid" not in src_cols:
            continue
        guid = _cell_text(
            src_sheet.getCellByPosition(src_cols[u"guid"], src_row)
        ).strip()
        key = _guid_key(guid)
        if not key:
            continue
        hit = guid_hits.get(key)
        if hit is None:
            continue
        src_fields = read_task_fields(src_sheet, src_cols, src_row)
        conflicts.append(
            {
                u"src_row": src_row,
                u"guid": guid,
                u"name": src_fields.get(u"name") or u"",
                u"hit": hit,
            }
        )
    decisions = {}
    apply_all = None
    ci = 0
    while ci < len(conflicts):
        item = conflicts[ci]
        remaining = len(conflicts) - ci
        hit = item[u"hit"]
        if apply_all:
            decisions[item[u"src_row"]] = {
                u"action": apply_all,
                u"hit": hit,
            }
            ci = ci + 1
            continue
        spec = show_todo_task_move_guid_conflict_dialog(
            doc=doc,
            dest_sheet_name=dest_name,
            task_name=item[u"name"],
            dest_task_name=hit.get(u"task_name") or u"",
            guid=item[u"guid"],
            remaining=remaining,
            conflict_index=ci + 1,
            conflict_total=len(conflicts),
            existing_sheet_name=hit.get(u"sheet_name") or u"",
        )
        if spec is None:
            return {}, True
        action = unicode(spec.get(u"action") or u"").strip()
        if not action:
            return {}, True
        decisions[item[u"src_row"]] = {u"action": action, u"hit": hit}
        if spec.get(u"apply_all"):
            apply_all = action
        ci = ci + 1
    return decisions, False


def _relocate_existing_task_to_dest( hit, dest_sheet, dest_cols, values_only=True, add_columns=False ):
    """
    Перенести существующую задачу с другого листа на приёмник (append + удалить исходную).
    Возвращает (dest_row, dest_cols) или (None, dest_cols) при ошибке.
    """
    other = hit.get(u"sheet")
    other_cols = hit.get(u"cols")
    other_row = int(hit.get(u"row"))
    if other is None or other_cols is None:
        return None, dest_cols
    dest_cols, pairs = _ensure_move_header_pairs(
        other,
        dest_sheet,
        dest_cols,
        add_columns=add_columns,
        values_only=values_only,
        src_cols=other_cols,
    )
    if not pairs:
        return None, dest_cols
    dest_row = _next_task_append_row(dest_sheet, dest_cols)
    _copy_move_row_by_pairs(
        other,
        other_row,
        dest_sheet,
        dest_row,
        pairs,
        values_only=values_only,
    )
    try:
        other.getRows().removeByIndex(other_row, 1)
    except Exception as exc:
        _log("relocate delete row: %s" % exc)
        return None, dest_cols
    try:
        _restore_numbered_row_after_delete(other, other_cols)
    except Exception:
        pass
    left = _iter_task_rows(other, other_cols)
    if left:
        _renumber_task_rows(other, other_cols, left)
    return int(dest_row), dest_cols


def _adjust_hit_rows_after_delete(decisions, sheet_name, deleted_row):
    """После удаления строки на чужом листе сдвинуть номера в оставшихся hit."""
    sk = _sheet_name_key(sheet_name)
    for _sr, dec in (decisions or {}).items():
        if not isinstance(dec, dict):
            continue
        hit = dec.get(u"hit")
        if not hit:
            continue
        if _sheet_name_key(hit.get(u"sheet_name")) != sk:
            continue
        try:
            r = int(hit.get(u"row"))
        except (TypeError, ValueError):
            continue
        if r > int(deleted_row):
            hit[u"row"] = r - 1


def _next_task_append_row(sheet, cols):
    tasks = _iter_task_rows(sheet, cols)
    if tasks:
        return int(tasks[-1]) + 1
    return _header_row(cols) + 1


def _mm_to_column_width(mm):
    """Миллиметры → единицы ширины колонки Calc (1/100 мм)."""
    try:
        return int(round(float(mm) * 100.0))
    except Exception:
        return 0


def _list_sheet_header_titles(sheet):
    """Упорядоченный список заголовков (строка шапки задач или первая заполненная)."""
    cols = _find_headers(sheet)
    hdr = None
    last = 40
    if cols is not None:
        hdr = _header_row(cols)
        last = max(_table_max_col(cols), _used_last_col(sheet)) + 8
    else:
        max_row = int(getattr(_cfg, "HEADER_SCAN_MAX_ROW", 19))
        try:
            last = max(last, _used_last_col(sheet) + 8)
        except Exception:
            pass
        r = 0
        while r <= max_row:
            filled = 0
            c = 0
            while c <= last:
                if _cell_text(sheet.getCellByPosition(c, r)).strip():
                    filled = filled + 1
                c = c + 1
            if filled >= 2:
                hdr = r
                break
            r = r + 1
    if hdr is None:
        return []
    items = []
    c = 0
    while c <= last:
        raw = _cell_text(sheet.getCellByPosition(c, hdr)).strip()
        if raw:
            items.append((c, raw))
        c = c + 1
    return items


def _get_sheet_by_name_ci(doc, name):
    name = unicode(name or u"").strip()
    if not name:
        return None
    sh = _get_sheet_by_name(doc, name)
    if sh is not None:
        return sh
    try:
        sheets = doc.Sheets
        n = int(sheets.getCount())
    except Exception:
        return None
    target = name.casefold() if hasattr(name, "casefold") else name.lower()
    i = 0
    while i < n:
        try:
            sh = sheets.getByIndex(i)
            nm = unicode(getattr(sh, "Name", u"") or u"")
            nm_cf = nm.casefold() if hasattr(nm, "casefold") else nm.lower()
            if nm_cf == target:
                return sh
        except Exception:
            pass
        i = i + 1
    return None


def _resolve_layout_template_sheet(doc, layout=None):
    from todo_task_settings_lib import normalize_layout_settings

    lay = normalize_layout_settings(layout)
    name = unicode(lay.get(u"template_sheet") or u"Задачи.01")
    sh = _get_sheet_by_name_ci(doc, name)
    if sh is not None and _find_headers(sh) is not None:
        return sh
    pat = unicode(lay.get(u"task_sheet_pattern") or u"Задачи*")
    sheets = _iter_sheets_matching_pattern(doc, pat)
    i = 0
    while i < len(sheets):
        if _find_headers(sheets[i]) is not None:
            return sheets[i]
        i = i + 1
    return None


def collect_layout_header_names(doc, layout=None):
    specs = collect_layout_column_specs(doc, layout)
    out = []
    i = 0
    while i < len(specs):
        out.append(specs[i][u"name"])
        i = i + 1
    return out


def _read_column_width_mm(sheet, col_idx):
    try:
        col = sheet.getColumns().getByIndex(int(col_idx))
        try:
            if hasattr(col, "IsVisible") and not bool(col.IsVisible):
                return 0.0
        except Exception:
            pass
        return float(col.Width) / 100.0
    except Exception:
        return 25.0


def collect_layout_column_specs(doc, layout=None):
    """Список {name, width_mm} с листа-шаблона (текущие ширины как default)."""
    from todo_task_settings_lib import normalize_layout_settings

    lay = normalize_layout_settings(layout)
    sh = _resolve_layout_template_sheet(doc, layout)
    out = []
    have = {}
    if sh is not None:
        pairs = _list_sheet_header_titles(sh)
        i = 0
        while i < len(pairs):
            col_idx, title = pairs[i]
            i = i + 1
            out.append(
                {
                    u"name": title,
                    u"width_mm": _read_column_width_mm(sh, col_idx),
                }
            )
            have[title] = True
    # всегда предлагать размеры для колонок автора
    for title in (
        getattr(_cfg, "COL_CREATED_BY", u"Создано"),
        getattr(_cfg, "COL_MODIFIED_BY", u"Отредактировано"),
    ):
        if title not in have:
            out.append({u"name": title, u"width_mm": 25.0})
            have[title] = True
    # «Категория» — со свода задач (на листах задач часто нет)
    cat_title = getattr(_cfg, "COL_CATEGORY", u"Категория")
    if cat_title not in have:
        cat_mm = 25.0
        try:
            pat = unicode(lay.get(u"summary_sheet_pattern") or u"Свод*задач*")
            matched = _iter_sheets_matching_pattern(doc, pat)
            mi = 0
            while mi < len(matched):
                ssh = matched[mi]
                mi = mi + 1
                pairs = _list_sheet_header_titles(ssh)
                pi = 0
                while pi < len(pairs):
                    col_idx, title = pairs[pi]
                    pi = pi + 1
                    if unicode(title).strip() == cat_title:
                        cat_mm = _read_column_width_mm(ssh, col_idx)
                        break
                else:
                    continue
                break
        except Exception as exc:
            _log("category width from summary: %s" % exc)
        out.append({u"name": cat_title, u"width_mm": cat_mm})
        have[cat_title] = True
    return out


def _layout_width_meta_titles():
    """Спец-поля сетки размеров: даты слева, автор справа; Категория — напротив GUID."""
    return {
        u"left": (
            unicode(_cfg.COL_CREATED),
            unicode(_cfg.COL_MODIFIED),
        ),
        u"right": (
            unicode(getattr(_cfg, "COL_CREATED_BY", u"Создано")),
            unicode(getattr(_cfg, "COL_MODIFIED_BY", u"Отредактировано")),
        ),
        u"guid": unicode(_cfg.COL_GUID),
        u"category": unicode(getattr(_cfg, "COL_CATEGORY", u"Категория")),
    }


def _split_width_items_for_layout_grid(width_items):
    """
    Обычные колонки + пара GUID|Категория + мета-блок
    (даты слева, Создано/Отредактировано справа).
    Возвращает (others, left_meta, right_meta, guid_item, category_item).
    """
    meta = _layout_width_meta_titles()
    left_names = list(meta[u"left"])
    right_names = list(meta[u"right"])
    guid_name = unicode(meta[u"guid"])
    cat_name = unicode(meta[u"category"])
    special = {guid_name: True, cat_name: True}
    i = 0
    while i < len(left_names):
        special[left_names[i]] = True
        i = i + 1
    i = 0
    while i < len(right_names):
        special[right_names[i]] = True
        i = i + 1

    by_name = {}
    others = []
    wi = 0
    items = list(width_items or [])
    while wi < len(items):
        it = items[wi]
        wi = wi + 1
        nm = unicode((it or {}).get(u"name") or u"").strip()
        if not nm:
            continue
        if nm in special:
            by_name[nm] = it
        else:
            others.append(it)

    def _item(name, default_mm=25.0):
        if name in by_name:
            return by_name[name]
        return {u"name": name, u"width_mm": float(default_mm)}

    left_meta = []
    i = 0
    while i < len(left_names):
        left_meta.append(_item(left_names[i]))
        i = i + 1
    right_meta = []
    i = 0
    while i < len(right_names):
        right_meta.append(_item(right_names[i]))
        i = i + 1
    guid_item = _item(guid_name)
    category_item = _item(cat_name)
    return others, left_meta, right_meta, guid_item, category_item


def _copy_sheet_column(sheet, src_col, dest_col, last_row):
    """Скопировать столбец (значения + форматы) через copyRange."""
    src_col = int(src_col)
    dest_col = int(dest_col)
    last_row = max(int(last_row), 0)
    try:
        src = sheet.getCellRangeByPosition(src_col, 0, src_col, last_row)
        dest = sheet.getCellByPosition(dest_col, 0)
        sheet.copyRange(dest.CellAddress, src.RangeAddress)
        return True
    except Exception as exc:
        _log("copy column %s->%s: %s" % (src_col, dest_col, exc))
        return False


def _fill_owner_meta_for_guid_rows(sheet, cols, settings=None, force=False, doc=None):
    """Заполнить «Создано»/«Отредактировано» для строк с GUID (пусто или 0)."""
    if cols is None or u"guid" not in cols:
        return 0
    owner = _get_owner_display_name(settings, doc=doc)
    if not owner:
        _log("fill owner: empty owner name (settings + OS)")
    hdr = _header_row(cols)
    last = _used_last_row(sheet, min_cols=cols.get(u"guid", 0) + 1)
    filled = 0
    r = hdr + 1
    while r <= last:
        if _cell_text(sheet.getCellByPosition(cols[u"guid"], r)).strip():
            for key in (u"created_by", u"modified_by"):
                if key not in cols:
                    continue
                cell = sheet.getCellByPosition(cols[key], r)
                if force or _owner_cell_is_blank(cell):
                    _set_cell_text(cell, owner)
                    filled = filled + 1
        r = r + 1
    return filled


def _ensure_owner_meta_columns(sheet, fill_owner=True, settings=None, doc=None):
    """
    Гарантировать колонки «Создано» / «Отредактировано» сразу после GUID.
    Если нет — вставить пустые столбцы, заголовки, заполнить ФИО владельца для строк с GUID.
    Возвращает (cols, created_flag, error).
    """
    cols = _find_headers(sheet)
    if cols is None:
        return None, False, u"нет таблицы задач"
    if u"guid" not in cols:
        return None, False, u"нет колонки GUID"
    need = []
    if u"created_by" not in cols:
        need.append(
            (
                u"created_by",
                getattr(_cfg, "COL_CREATED_BY", u"Создано"),
            )
        )
    if u"modified_by" not in cols:
        need.append(
            (
                u"modified_by",
                getattr(_cfg, "COL_MODIFIED_BY", u"Отредактировано"),
            )
        )
    created = False
    if need:
        hdr = _header_row(cols)
        guid_col = int(cols[u"guid"])
        insert_at = guid_col + 1
        n_ins = len(need)
        try:
            sheet.getColumns().insertByIndex(insert_at, n_ins)
        except Exception as exc:
            return None, False, u"не удалось вставить столбцы: %s" % exc

        # Пустые столбцы (без copy GUID — иначе после очистки остаются «0»)
        i = 0
        while i < n_ins:
            dest_col = insert_at + i
            key, title = need[i]
            _set_cell_text(sheet.getCellByPosition(dest_col, hdr), title)
            i = i + 1
        created = True
        cols = _find_headers(sheet)
        if cols is None:
            return None, True, u"после вставки не удалось найти заголовки"

    if fill_owner and cols is not None:
        _fill_owner_meta_for_guid_rows(
            sheet, cols, settings=settings, force=created, doc=doc
        )
    return cols, created, u""


def _set_sheet_column_width_mm(sheet, col_idx, width_mm):
    try:
        cols = sheet.getColumns()
        col = cols.getByIndex(int(col_idx))
    except Exception as exc:
        _log("column get %s: %s" % (col_idx, exc))
        return False
    try:
        w = float(width_mm)
    except (TypeError, ValueError):
        w = 0.0
    if w <= 0:
        try:
            col.IsVisible = False
        except Exception:
            pass
        try:
            col.Width = 0
        except Exception:
            pass
        return True
    try:
        col.IsVisible = True
    except Exception:
        pass
    try:
        col.Width = _mm_to_column_width(w)
        return True
    except Exception as exc:
        _log("column width %s: %s" % (col_idx, exc))
        return False


def apply_todo_column_widths(doc, settings=None, sheets=None, progress=None):
    """
    Выставить ширины колонок по настройкам layout на листы задач.
    Листы свода (по шаблону) — только если layout.apply_to_summary.
    На листах задач при отсутствии «Создано»/«Отредактировано» — создать после GUID.
    progress — объект с методами update(text, value=None, maximum=None) / close().
    Возвращает (ok_count, fail_count, detail).
    """
    from todo_task_settings_lib import (
        layout_width_map,
        load_todo_task_settings,
        normalize_layout_settings,
    )

    tm = _StageTimer("apply_column_widths")
    if settings is None:
        settings = load_todo_task_settings(doc=doc)
    lay = normalize_layout_settings((settings or {}).get(u"layout"))
    wmap = layout_width_map(lay)
    if not wmap:
        tm.done("empty_map")
        return 0, 0, u"Нет заданных ширин колонок."

    task_pat = unicode(lay.get(u"task_sheet_pattern") or u"Задачи*")
    summary_pat = unicode(lay.get(u"summary_sheet_pattern") or u"Свод*задач*")
    apply_to_summary = bool(lay.get(u"apply_to_summary"))
    explicit_sheets = sheets is not None
    if sheets is None:
        sheets = []
        seen = set()
        pats = [task_pat]
        if apply_to_summary:
            pats.append(summary_pat)
        pi = 0
        while pi < len(pats):
            pat = pats[pi]
            pi = pi + 1
            matched = _iter_sheets_matching_pattern(doc, pat)
            mi = 0
            while mi < len(matched):
                sh = matched[mi]
                mi = mi + 1
                try:
                    nm = unicode(getattr(sh, "Name", u"") or u"")
                except Exception:
                    nm = u"?"
                if nm in seen:
                    continue
                seen.add(nm)
                sheets.append(sh)

    total = len(sheets)
    if progress is not None:
        try:
            progress.update(
                u"Подготовка… листов: %s" % total, value=0, maximum=max(total, 1)
            )
        except Exception:
            pass

    ok_n = 0
    fail_n = 0
    created_n = 0
    si = 0
    while si < len(sheets):
        sh = sheets[si]
        si = si + 1
        name = _sheet_display_name(sh)
        is_summary = _sheet_name_matches_pattern(name, summary_pat)
        if is_summary and not apply_to_summary and not explicit_sheets:
            continue
        if progress is not None:
            try:
                progress.update(
                    u"Лист %s / %s: %s" % (si, total, name),
                    value=si - 1,
                    maximum=max(total, 1),
                )
            except Exception:
                pass
        is_task_sheet = (
            not _is_merge_source_sheet(sh, doc=doc, name=name)
            and (
                _sheet_name_matches_pattern(name, task_pat)
                or _is_named_task_sheet(sh, doc=doc)
            )
        )
        guard = _TaskSheetProtectSession(sh, doc=doc)
        if not guard.begin():
            fail_n = fail_n + 1
            _log("apply widths: protect fail %s" % name)
            continue
        try:
            if is_task_sheet and _find_headers(sh) is not None:
                _cols, created, err = _ensure_owner_meta_columns(
                    sh, fill_owner=True, settings=settings, doc=doc
                )
                if err and _cols is None:
                    fail_n = fail_n + 1
                    _log("ensure owner cols %s: %s" % (name, err))
                    continue
                if created:
                    created_n = created_n + 1
            headers = _list_sheet_header_titles(sh)
            if not headers:
                fail_n = fail_n + 1
                _log("apply widths: skip %s (no headers)" % name)
                continue
            hi = 0
            while hi < len(headers):
                col_idx, title = headers[hi]
                hi = hi + 1
                if title not in wmap:
                    tkey = title.strip()
                    found = None
                    for k in wmap:
                        if unicode(k).strip() == tkey:
                            found = k
                            break
                    if found is None:
                        continue
                    mm = wmap[found]
                else:
                    mm = wmap[title]
                _set_sheet_column_width_mm(sh, col_idx, mm)
            ok_n = ok_n + 1
        finally:
            guard.finish()
    if progress is not None:
        try:
            progress.update(
                u"Готово: %s лист(ов)" % ok_n, value=total, maximum=max(total, 1)
            )
        except Exception:
            pass
    tm.done(
        "DONE",
        "ok=%s fail=%s created_owner_cols=%s keys=%s"
        % (ok_n, fail_n, created_n, len(wmap)),
    )
    detail = u"Листов обновлено: %s, ошибок: %s." % (ok_n, fail_n)
    if created_n:
        detail = detail + u" Колонки «Создано»/«Отредактировано» добавлены на %s лист(ах)." % created_n
    return ok_n, fail_n, detail


def discover_ref_dictionaries(sheet):
    """
    Справочники: с колонки A, блоки через ≥1 пустой столбец в строке заголовка.
    Возвращает список dict: start_col, end_col, titles, title.
    """
    out = []
    if sheet is None:
        return out
    hdr = int(getattr(_cfg, "REF_LIST_HEADER_ROW", 0))
    last_col = 40
    try:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        last_col = max(last_col, int(cursor.getRangeAddress().EndColumn) + 2)
    except Exception:
        pass
    c = 0
    while c <= last_col:
        title = _cell_text(sheet.getCellByPosition(c, hdr)).strip()
        if not title:
            c = c + 1
            continue
        start = c
        titles = [title]
        c = c + 1
        while c <= last_col:
            t2 = _cell_text(sheet.getCellByPosition(c, hdr)).strip()
            if not t2:
                break
            titles.append(t2)
            c = c + 1
        end = start + len(titles) - 1
        label = u" / ".join(titles)
        out.append(
            {
                u"start_col": start,
                u"end_col": end,
                u"titles": titles,
                u"title": label,
            }
        )
        # пропуск пустых столбцов-разделителей
        while c <= last_col and not _cell_text(sheet.getCellByPosition(c, hdr)).strip():
            c = c + 1
    return out


def _ref_row_values(sheet, start_col, end_col, row):
    vals = []
    c = int(start_col)
    while c <= int(end_col):
        vals.append(_ref_cell_text(sheet.getCellByPosition(c, row)))
        c = c + 1
    return vals


def _ref_cell_text(cell):
    """Текст ячейки справочника: числовой 0 / «0» считаем пустым (артефакт очистки)."""
    try:
        if hasattr(cell, "getString"):
            s = unicode(cell.getString() or u"").strip()
        else:
            s = unicode(getattr(cell, "String", u"") or u"").strip()
    except Exception:
        s = u""
    if s and s != u"0":
        return s
    try:
        v = float(cell.Value)
        if v == 0.0:
            return u""
    except Exception:
        pass
    if s == u"0":
        return u""
    return s


def _ref_row_is_empty(vals):
    i = 0
    while i < len(vals):
        s = unicode(vals[i] or u"").strip()
        if s and s != u"0":
            return False
        i = i + 1
    return True


def _ref_normalize_item(vals, ncols=None):
    """Нормализовать строку справочника; пустые/«0» → ''."""
    if not isinstance(vals, (list, tuple)):
        vals = [unicode(vals or u"").strip()]
    out = []
    i = 0
    while i < len(vals):
        s = unicode(vals[i] or u"").strip()
        if s == u"0":
            s = u""
        out.append(s)
        i = i + 1
    if ncols is not None:
        ncols = int(ncols)
        while len(out) < ncols:
            out.append(u"")
        if len(out) > ncols:
            out = out[:ncols]
    return out


def _ref_item_key(vals):
    parts = []
    i = 0
    while i < len(vals):
        p = unicode(vals[i] or u"").strip()
        if p == u"0":
            p = u""
        parts.append(p.casefold() if hasattr(p, "casefold") else p.lower())
        i = i + 1
    return u"|".join(parts)


def _ref_item_display(vals):
    cleaned = []
    i = 0
    vals = list(vals or [])
    while i < len(vals):
        s = unicode(vals[i] or u"").strip()
        if s == u"0":
            s = u""
        cleaned.append(s)
        i = i + 1
    # для одноколоночного не показывать лишние « | »
    while cleaned and not cleaned[-1]:
        cleaned.pop()
    if not cleaned:
        return u""
    return u" | ".join(cleaned)


def _parse_ref_item_text(text, ncols):
    text = unicode(text or u"").strip()
    if ncols <= 1:
        return [text]
    parts = [p.strip() for p in text.split(u"|")]
    while len(parts) < ncols:
        parts.append(u"")
    if len(parts) > ncols:
        parts = parts[: ncols - 1] + [u"|".join(parts[ncols - 1 :])]
    return parts


def read_ref_dictionary_items(sheet, ref_dict):
    """Список элементов справочника (каждый — list значений по столбцам)."""
    out = []
    if sheet is None or not ref_dict:
        return out
    start = int(ref_dict[u"start_col"])
    end = int(ref_dict[u"end_col"])
    data_start = int(getattr(_cfg, "REF_LIST_DATA_START_ROW", 1))
    last = _used_last_row(sheet, min_cols=end + 1)
    r = data_start
    empty_run = 0
    while r <= max(last, data_start) + 50:
        vals = _ref_row_values(sheet, start, end, r)
        if _ref_row_is_empty(vals):
            empty_run = empty_run + 1
            if empty_run >= 3 and r > last:
                break
            r = r + 1
            continue
        empty_run = 0
        out.append(vals)
        r = r + 1
    return out


def write_ref_dictionary_items(sheet, ref_dict, items):
    """Перезаписать данные справочника (заголовок не трогаем)."""
    if sheet is None or not ref_dict:
        return False
    start = int(ref_dict[u"start_col"])
    end = int(ref_dict[u"end_col"])
    ncols = end - start + 1
    data_start = int(getattr(_cfg, "REF_LIST_DATA_START_ROW", 1))
    # нормализовать и отбросить пустые/нулевые строки
    clean_items = []
    raw_items = list(items or [])
    ri = 0
    while ri < len(raw_items):
        vals = raw_items[ri]
        ri = ri + 1
        if not isinstance(vals, (list, tuple)):
            vals = _parse_ref_item_text(unicode(vals), ncols)
        vals = _ref_normalize_item(vals, ncols=ncols)
        if _ref_row_is_empty(vals):
            continue
        clean_items.append(vals)
    # очистить старые данные (+ запас под хвосты «0»)
    old = read_ref_dictionary_items(sheet, ref_dict)
    clear_to = data_start + max(len(old), len(clean_items), 0) + 20
    try:
        last = _used_last_row(sheet, min_cols=end + 1)
        if last > clear_to:
            clear_to = last + 2
    except Exception:
        pass
    r = data_start
    while r <= clear_to:
        c = start
        while c <= end:
            _clear_cell_value(sheet.getCellByPosition(c, r))
            c = c + 1
        r = r + 1
    # записать
    r = data_start
    i = 0
    while i < len(clean_items):
        vals = clean_items[i]
        i = i + 1
        c = start
        vi = 0
        while c <= end:
            _set_cell_text(sheet.getCellByPosition(c, r), unicode(vals[vi] or u""))
            c = c + 1
            vi = vi + 1
        r = r + 1
    return True


def dedupe_ref_items(items):
    """Убрать дубликаты (trim + без учёта регистра), сохранить первый."""
    out = []
    seen = set()
    i = 0
    items = list(items or [])
    while i < len(items):
        vals = items[i]
        i = i + 1
        if not isinstance(vals, (list, tuple)):
            vals = [unicode(vals or u"").strip()]
        cleaned = _ref_normalize_item(vals)
        if _ref_row_is_empty(cleaned):
            continue
        key = _ref_item_key(cleaned)
        if key in seen:
            continue
        seen.add(key)
        out.append(cleaned)
    return out


def _merge_layout_header_specs(layout, col_specs):
    from todo_task_settings_lib import normalize_layout_settings as _nlay

    lay0 = _nlay(layout)
    have = {}
    for it in lay0.get(u"column_widths") or []:
        have[unicode((it or {}).get(u"name") or u"")] = True
    merged = list(lay0.get(u"column_widths") or [])
    si = 0
    while si < len(col_specs or []):
        sp = col_specs[si]
        si = si + 1
        nm = unicode((sp or {}).get(u"name") or u"")
        if nm and nm not in have:
            merged.append(
                {
                    u"name": nm,
                    u"width_mm": float((sp or {}).get(u"width_mm") or 25.0),
                }
            )
            have[nm] = True
    return _nlay(
        {
            u"template_sheet": lay0.get(u"template_sheet"),
            u"task_sheet_pattern": lay0.get(u"task_sheet_pattern"),
            u"summary_sheet_pattern": lay0.get(u"summary_sheet_pattern"),
            u"merge_source_sheet": lay0.get(u"merge_source_sheet"),
            u"apply_to_summary": lay0.get(u"apply_to_summary"),
            u"column_widths": merged,
        }
    )


def todo_task_settings(doc=None):
    """Визард параметров серии todo_task (частные по книге / глобальные)."""
    tm = _StageTimer("todo_task_settings")
    doc = _get_doc(doc)
    from todo_task_dialog_lib import show_todo_task_settings_dialog
    from todo_task_settings_lib import (
        SCOPE_GLOBAL,
        SCOPE_PRIVATE,
        get_doc_book_path,
        get_protect_password,
        has_private_todo_task_settings,
        load_todo_task_settings,
        normalize_general_settings,
        normalize_layout_settings,
        normalize_merge_settings,
        save_todo_task_settings,
    )

    book_path = get_doc_book_path(doc)
    has_private = has_private_todo_task_settings(doc=doc, book_path=book_path)
    global_settings = load_todo_task_settings(doc=doc, scope=SCOPE_GLOBAL)
    settings = load_todo_task_settings(doc=doc, scope=SCOPE_PRIVATE)
    header_names = collect_layout_header_names(doc, settings.get(u"layout"))
    col_specs = collect_layout_column_specs(doc, settings.get(u"layout"))
    try:
        settings[u"layout"] = _merge_layout_header_specs(
            settings.get(u"layout"), col_specs
        )
        global_settings[u"layout"] = _merge_layout_header_specs(
            global_settings.get(u"layout"), col_specs
        )
    except Exception as exc:
        _log("layout merge specs: %s" % exc)
    try:
        result = show_todo_task_settings_dialog(
            doc=doc,
            initial_general=settings.get(u"general"),
            initial_layout=settings.get(u"layout"),
            initial_merge=settings.get(u"merge"),
            header_names=header_names,
            initial_scope=SCOPE_PRIVATE,
            global_general=global_settings.get(u"general"),
            global_layout=global_settings.get(u"layout"),
            global_merge=global_settings.get(u"merge"),
            book_path=book_path,
            has_private=has_private,
        )
    except Exception as exc:
        _msg(doc, u"Диалог настроек: %s" % exc, box_type=WARNINGBOX)
        tm.done("dialog_fail")
        return False
    if result is None:
        tm.done("cancel")
        return False

    gen_spec = normalize_general_settings(result.get(u"general"))
    lay_spec = normalize_layout_settings(
        result.get(u"layout"), header_names=header_names
    )
    merge_spec = normalize_merge_settings(result.get(u"merge"))
    scope = unicode(result.get(u"scope") or SCOPE_PRIVATE).strip()
    if scope != SCOPE_GLOBAL:
        scope = SCOPE_PRIVATE
    saved_as_global_fallback = False
    if scope == SCOPE_PRIVATE and not book_path:
        scope = SCOPE_GLOBAL
        saved_as_global_fallback = True
    old_src = global_settings if scope == SCOPE_GLOBAL else settings
    old_pwd = get_protect_password(old_src)
    new_pwd = unicode(
        gen_spec.get(u"protect_password")
        if gen_spec.get(u"protect_password") is not None
        else u""
    )
    pwd_changed = new_pwd != unicode(old_pwd if old_pwd is not None else u"")
    if pwd_changed:
        gen_spec[u"protect_password_prev"] = unicode(
            old_pwd if old_pwd is not None else u""
        )
    else:
        prev = u""
        try:
            prev = unicode(
                (old_src.get(u"general") or {}).get(u"protect_password_prev") or u""
            )
        except Exception:
            prev = u""
        gen_spec[u"protect_password_prev"] = prev

    target = dict(old_src)
    target[u"general"] = gen_spec
    target[u"layout"] = lay_spec
    target[u"merge"] = merge_spec
    save_err = save_todo_task_settings(target, doc=doc, scope=scope)
    if save_err:
        _msg(doc, save_err, box_type=WARNINGBOX)
        tm.done("save_fail")
        return False
    try:
        from todo_task_settings_lib import invalidate_todo_task_settings_cache

        invalidate_todo_task_settings_cache()
    except Exception:
        pass
    tm.step("save")

    if scope == SCOPE_GLOBAL:
        lines = [u"Глобальные настройки задач сохранены."]
    else:
        lines = [u"Частные настройки задач сохранены для этой книги."]
    if saved_as_global_fallback:
        lines.append(
            u"Книга не сохранена на диск — записано в глобальные настройки."
        )
    if pwd_changed:
        ok_names, fail_names = _reapply_protection_passwords(doc, old_pwd, new_pwd)
        tm.step(
            "reapply_protect",
            "ok=%s fail=%s" % (len(ok_names), len(fail_names)),
        )
        if ok_names:
            lines.append(
                u"Пароль защиты обновлён (%s): %s"
                % (len(ok_names), u", ".join(ok_names[:12]))
            )
        else:
            lines.append(u"Листов для смены пароля защиты не найдено.")
        if fail_names:
            lines.append(
                u"Не удалось сменить пароль: %s" % u", ".join(fail_names[:12])
            )
    lines.append(
        u"Логи в консоль: %s"
        % (u"вкл" if gen_spec.get(u"console_log") else u"выкл")
    )
    lines.append(
        u"Слияние: только существующие %s; дата редакции %s; накопление комментариев %s"
        % (
            u"вкл" if merge_spec.get(u"update_existing_only") else u"выкл",
            u"вкл" if merge_spec.get(u"check_modified") else u"выкл",
            u"вкл" if merge_spec.get(u"accumulate_comments") else u"выкл",
        )
    )

    if result.get(u"apply_widths"):
        progress = None
        try:
            from todo_task_dialog_lib import (
                restore_todo_doc_ui,
                show_todo_task_progress_dialog,
            )

            progress = show_todo_task_progress_dialog(
                doc,
                title=u"Размеры полей",
                text=u"Применение размеров колонок…",
            )
        except Exception as err:
            _log("progress dialog: %s" % err)
            progress = None
        try:
            ok_n, fail_n, detail = apply_todo_column_widths(
                doc, target, progress=progress
            )
        finally:
            if progress is not None:
                try:
                    progress.close()
                except Exception:
                    pass
            try:
                from todo_task_dialog_lib import restore_todo_doc_ui

                restore_todo_doc_ui(doc)
            except Exception:
                pass
        tm.step("apply_widths", "ok=%s fail=%s" % (ok_n, fail_n))
        lines.append(detail)

    _msg(doc, u"\n".join(lines))
    try:
        from todo_task_dialog_lib import restore_todo_doc_ui

        restore_todo_doc_ui(doc)
    except Exception:
        pass
    tm.done("DONE")
    return True


def todo_task_refs(doc=None):
    """Редактирование справочников на листе __Справочники_задачи."""
    tm = _StageTimer("todo_task_refs")
    doc = _get_doc(doc)
    if doc is None:
        _msg(doc, u"Нет открытого документа Calc.", box_type=WARNINGBOX)
        return False
    sheet = _get_sheet_by_name(doc, getattr(_cfg, "REF_SHEET_NAME", u"__Справочники_задачи"))
    if sheet is None:
        _msg(
            doc,
            u"Лист «%s» не найден."
            % getattr(_cfg, "REF_SHEET_NAME", u"__Справочники_задачи"),
            box_type=WARNINGBOX,
        )
        tm.done("no_sheet")
        return False
    dicts = discover_ref_dictionaries(sheet)
    if not dicts:
        _msg(doc, u"На листе справочников не найдено ни одного справочника.", box_type=WARNINGBOX)
        tm.done("empty")
        return False

    from todo_task_dialog_lib import show_todo_task_refs_dialog

    guard = _TaskSheetProtectSession(sheet, doc=doc)
    if not guard.begin():
        _msg(doc, _protect_unlock_error_text(), box_type=WARNINGBOX)
        tm.done("protect_fail")
        return False
    try:
        result = show_todo_task_refs_dialog(doc=doc, sheet=sheet, dictionaries=dicts)
        if result is None:
            tm.done("cancel")
            return False
        # result: {title -> items}
        changed = 0
        di = 0
        while di < len(dicts):
            d = dicts[di]
            di = di + 1
            title = d.get(u"title")
            if title not in result:
                continue
            items = dedupe_ref_items(result.get(title) or [])
            write_ref_dictionary_items(sheet, d, items)
            changed = changed + 1
        tm.done("DONE", "dicts=%s" % changed)
        _msg(doc, u"Справочники сохранены (%s)." % changed)
        return True
    except Exception as exc:
        _msg(doc, u"Ошибка справочников: %s" % exc, box_type=WARNINGBOX)
        tm.done("error")
        return False
    finally:
        guard.finish()



def _sheet_category_from_b1(sheet):
    """Расшифровка листа: значение категории из B1."""
    if sheet is None:
        return u""
    try:
        return _cell_text(sheet.getCellByPosition(1, 0)).strip()
    except Exception:
        return u""


def _dest_sheet_choice_label(sheet):
    """Подпись листа для combo перемещения: «Имя (категория из B1)»."""
    try:
        name = unicode(getattr(sheet, "Name", u"") or u"").strip()
    except Exception:
        name = u""
    if not name:
        return u""
    cat = _sheet_category_from_b1(sheet)
    if cat:
        return u"%s (%s)" % (name, cat)
    return name


def todo_task_move(doc=None):
    """Переместить или скопировать выбранные задачи на другой лист «Задачи…»."""
    tm = _StageTimer("todo_task_move")
    doc = _get_doc(doc)
    sheet, cols, err = _prepare_sheet(doc)
    if err:
        _msg(doc, err, box_type=WARNINGBOX)
        return False

    try:
        src_name = unicode(getattr(sheet, "Name", u"") or u"")
    except Exception:
        src_name = u""

    task_pat = _task_sheet_pattern_for_doc(doc)
    dest_ready = _iter_move_dest_sheets(doc, exclude_name=src_name, pattern=task_pat)
    if not dest_ready:
        _msg(
            doc,
            u"Нет других листов задач (шаблон «%s») с таблицей для перемещения."
            % task_pat,
            box_type=WARNINGBOX,
        )
        return False

    task_rows = _iter_selected_task_rows(doc, sheet, cols)
    if not task_rows:
        _msg(
            doc,
            u"Выделите одну или несколько строк-задач (мультиселект).",
            box_type=WARNINGBOX,
        )
        return False

    previews = []
    pi = 0
    while pi < len(task_rows):
        fields = read_task_fields(sheet, cols, task_rows[pi])
        previews.append(
            {
                u"name": fields.get(u"name") or u"",
                u"guid": fields.get(u"guid") or u"",
                u"row": task_rows[pi],
            }
        )
        pi = pi + 1

    dest_choices = []
    di = 0
    while di < len(dest_ready):
        sh = dest_ready[di]
        di = di + 1
        try:
            nm = unicode(sh.Name)
        except Exception:
            continue
        dest_choices.append(
            {
                u"name": nm,
                u"label": _dest_sheet_choice_label(sh),
            }
        )

    default_dest = _default_dest_name_by_category(
        sheet, cols, task_rows, dest_ready
    )

    from todo_task_dialog_lib import show_todo_task_move_dialog

    try:
        spec = show_todo_task_move_dialog(
            doc=doc,
            tasks=previews,
            dest_choices=dest_choices,
            source_name=src_name,
            default_dest_name=default_dest,
        )
    except Exception as exc:
        _msg(doc, u"Диалог перемещения: %s" % exc, box_type=WARNINGBOX)
        return False
    if spec is None:
        return False

    dest_name = unicode(spec.get(u"dest_name") or u"").strip()
    copy_mode = bool(spec.get(u"copy"))
    values_only = spec.get(u"values_only")
    if values_only is None:
        values_only = bool(getattr(_cfg, "DEFAULT_MOVE_VALUES_ONLY", True))
    else:
        values_only = bool(values_only)
    add_columns = spec.get(u"add_columns")
    if add_columns is None:
        add_columns = bool(getattr(_cfg, "DEFAULT_MOVE_ADD_COLUMNS", False))
    else:
        add_columns = bool(add_columns)
    if not dest_name:
        _msg(doc, u"Укажите лист назначения.", box_type=WARNINGBOX)
        return False
    dest_cf = _sheet_name_key(dest_name)
    src_cf = _sheet_name_key(src_name)
    if dest_cf == src_cf:
        _msg(doc, u"Лист назначения совпадает с источником.", box_type=WARNINGBOX)
        return False

    dest_sheet = None
    di = 0
    while di < len(dest_ready):
        try:
            nm = unicode(dest_ready[di].Name)
            if nm == dest_name or _sheet_name_key(nm) == dest_cf:
                dest_sheet = dest_ready[di]
                dest_name = nm
                break
        except Exception:
            pass
        di = di + 1
    if dest_sheet is None:
        _msg(doc, u"Лист назначения «%s» не найден." % dest_name, box_type=WARNINGBOX)
        return False

    dest_cols = _find_headers(dest_sheet)
    if dest_cols is None:
        _msg(
            doc,
            u"На листе «%s» нет таблицы задач." % dest_name,
            box_type=WARNINGBOX,
        )
        return False

    act_overwrite = unicode(
        getattr(_cfg, "MOVE_GUID_ACTION_OVERWRITE", u"overwrite")
    )
    act_merge = unicode(
        getattr(_cfg, "MOVE_GUID_ACTION_MERGE_COMMENTS", u"merge_comments")
    )
    act_new = unicode(getattr(_cfg, "MOVE_GUID_ACTION_NEW_GUID", u"new_guid"))

    guid_hits = _collect_move_guid_hits(
        doc,
        task_pat,
        exclude_sheet_name=src_name,
        prefer_sheet_name=dest_name,
    )
    try:
        decisions, aborted = _resolve_move_guid_conflicts(
            doc,
            sheet,
            cols,
            dest_name,
            task_rows,
            guid_hits,
        )
    except Exception as exc:
        _msg(doc, u"Диалог совпадения GUID: %s" % exc, box_type=WARNINGBOX)
        return False
    if aborted:
        return False

    # листы чужих совпадений, с которых может понадобиться перенос
    extra_sheets = {}
    for _sr, dec in decisions.items():
        if not isinstance(dec, dict):
            continue
        hit = dec.get(u"hit") or {}
        action = unicode(dec.get(u"action") or u"")
        if action not in (act_overwrite, act_merge):
            continue
        hs = hit.get(u"sheet")
        hn = unicode(hit.get(u"sheet_name") or u"")
        if hs is None:
            continue
        if _sheet_name_key(hn) == dest_cf or _sheet_name_key(hn) == src_cf:
            continue
        extra_sheets[_sheet_name_key(hn)] = hs

    src_guard = _TaskSheetProtectSession(sheet, doc=doc)
    dest_guard = _TaskSheetProtectSession(dest_sheet, doc=doc)
    extra_guards = []
    if not src_guard.begin():
        _msg(doc, _protect_unlock_error_text(), box_type=WARNINGBOX)
        return False
    if not dest_guard.begin():
        src_guard.finish()
        _msg(doc, _protect_unlock_error_text(), box_type=WARNINGBOX)
        return False
    for _ek, esh in extra_sheets.items():
        g = _TaskSheetProtectSession(esh, doc=doc)
        if not g.begin():
            dest_guard.finish()
            src_guard.finish()
            ei = 0
            while ei < len(extra_guards):
                try:
                    extra_guards[ei].finish()
                except Exception:
                    pass
                ei = ei + 1
            _msg(doc, _protect_unlock_error_text(), box_type=WARNINGBOX)
            return False
        extra_guards.append(g)
    try:
        dest_cols, header_pairs = _ensure_move_header_pairs(
            sheet,
            dest_sheet,
            dest_cols,
            add_columns=add_columns,
            values_only=values_only,
            src_cols=cols,
        )
        if not header_pairs:
            _msg(
                doc,
                u"Нет совпадающих колонок по заголовкам между «%s» и «%s»."
                % (src_name, dest_name),
                box_type=WARNINGBOX,
            )
            return False
        copied = 0
        n_overwrite = 0
        n_merge = 0
        n_new = 0
        n_relocated = 0
        dest_fit_rows = []
        ci = 0
        while ci < len(task_rows):
            src_row = int(task_rows[ci])
            ci = ci + 1
            dec = decisions.get(src_row)
            action = u""
            hit = None
            if isinstance(dec, dict):
                action = unicode(dec.get(u"action") or u"")
                hit = dec.get(u"hit")
            elif dec:
                action = unicode(dec)

            dest_row = None
            skip_comment = False
            stamp_as_new = False
            assign_new_guid = False

            if action in (act_overwrite, act_merge) and hit is not None:
                hit_name = unicode(hit.get(u"sheet_name") or u"")
                if _sheet_name_key(hit_name) != dest_cf:
                    old_row = int(hit.get(u"row"))
                    new_row, dest_cols = _relocate_existing_task_to_dest(
                        hit,
                        dest_sheet,
                        dest_cols,
                        values_only=values_only,
                        add_columns=add_columns,
                    )
                    if new_row is None:
                        _msg(
                            doc,
                            u"Не удалось перенести существующую задачу "
                            u"с листа «%s» на «%s»."
                            % (hit_name, dest_name),
                            box_type=WARNINGBOX,
                        )
                        return False
                    _adjust_hit_rows_after_delete(decisions, hit_name, old_row)
                    hit[u"sheet"] = dest_sheet
                    hit[u"cols"] = dest_cols
                    hit[u"row"] = int(new_row)
                    hit[u"sheet_name"] = dest_name
                    n_relocated = n_relocated + 1
                    # пары источник→приёмник могли устареть после add_columns
                    dest_cols, header_pairs = _ensure_move_header_pairs(
                        sheet,
                        dest_sheet,
                        dest_cols,
                        add_columns=False,
                        values_only=values_only,
                        src_cols=cols,
                    )
                dest_row = int(hit.get(u"row"))
                if action == act_overwrite:
                    n_overwrite = n_overwrite + 1
                else:
                    skip_comment = True
                    n_merge = n_merge + 1
            elif action == act_new:
                dest_row = _next_task_append_row(dest_sheet, dest_cols)
                assign_new_guid = True
                stamp_as_new = True
                n_new = n_new + 1
            else:
                dest_row = _next_task_append_row(dest_sheet, dest_cols)
                if copy_mode:
                    assign_new_guid = True
                    stamp_as_new = True
            if dest_row is None:
                dest_row = _next_task_append_row(dest_sheet, dest_cols)
                if copy_mode or action == act_new:
                    assign_new_guid = True
                    stamp_as_new = True

            skip_cols = []
            if skip_comment:
                _merge_move_row_comments(
                    sheet,
                    src_row,
                    dest_sheet,
                    dest_row,
                    cols,
                    dest_cols,
                )
                cmt_cell = _comment_cell(dest_sheet, dest_cols, dest_row)
                cmt_col = None
                if cmt_cell is not None:
                    try:
                        cmt_col = int(cmt_cell.CellAddress.Column)
                    except Exception:
                        try:
                            cmt_col = int(cmt_cell.getCellAddress().Column)
                        except Exception:
                            cmt_col = None
                if cmt_col is None and dest_cols is not None and u"comment" in dest_cols:
                    cmt_col = dest_cols[u"comment"]
                if cmt_col is not None:
                    skip_cols.append(cmt_col)
            _copy_move_row_by_pairs(
                sheet,
                src_row,
                dest_sheet,
                dest_row,
                header_pairs,
                values_only=values_only,
                skip_dest_cols=skip_cols,
            )
            if assign_new_guid and dest_cols is not None and u"guid" in dest_cols:
                new_g = _new_guid()
                _set_cell_text(
                    dest_sheet.getCellByPosition(dest_cols[u"guid"], dest_row),
                    new_g,
                )
            if stamp_as_new:
                _stamp_task_timestamps(
                    doc, dest_sheet, dest_cols, dest_row, is_new=True
                )
                _stamp_task_owner(
                    dest_sheet, dest_cols, dest_row, is_new=True, doc=doc
                )
            dest_fit_rows.append(int(dest_row))
            copied = copied + 1

        if not copy_mode:
            rev = list(task_rows)
            rev.sort(reverse=True)
            ri = 0
            while ri < len(rev):
                try:
                    sheet.getRows().removeByIndex(int(rev[ri]), 1)
                except Exception as exc:
                    _msg(
                        doc,
                        u"Не удалось удалить строку %s: %s" % (int(rev[ri]) + 1, exc),
                        box_type=WARNINGBOX,
                    )
                    return False
                ri = ri + 1
            _restore_numbered_row_after_delete(sheet, cols)
            src_left = _iter_task_rows(sheet, cols)
            if src_left:
                _renumber_task_rows(sheet, cols, src_left)

        dest_left = _iter_task_rows(dest_sheet, dest_cols)
        _renumber_task_rows(dest_sheet, dest_cols, dest_left)
        _autofit_task_rows(dest_sheet, dest_fit_rows)

        extra = []
        if n_relocated:
            extra.append(u"подтянуто с других листов: %s" % n_relocated)
        if n_overwrite:
            extra.append(u"перезаписано: %s" % n_overwrite)
        if n_merge:
            extra.append(u"слияние комментариев: %s" % n_merge)
        if n_new:
            extra.append(u"новые GUID: %s" % n_new)
        extra_s = u""
        if extra:
            extra_s = u"\n" + u"; ".join(extra)
        kept = n_overwrite + n_merge
        if copy_mode:
            guid_note = u"новые GUID"
            if kept and kept == copied:
                guid_note = u"GUID совпавших сохранены"
            elif kept:
                guid_note = u"новые GUID, кроме совпавших"
            _msg(
                doc,
                u"Скопировано задач: %s\n«%s» → «%s» (%s).%s"
                % (copied, src_name, dest_name, guid_note, extra_s),
            )
        else:
            guid_note = u"GUID сохранены"
            if n_new and n_new == copied:
                guid_note = u"новые GUID"
            elif n_new:
                guid_note = u"GUID сохранены, кроме новых"
            _msg(
                doc,
                u"Перемещено задач: %s\n«%s» → «%s» (%s).%s"
                % (copied, src_name, dest_name, guid_note, extra_s),
            )
        return True
    except Exception as exc:
        _msg(doc, u"Ошибка перемещения: %s" % exc, box_type=WARNINGBOX)
        return False
    finally:
        ei = 0
        while ei < len(extra_guards):
            try:
                extra_guards[ei].finish()
            except Exception:
                pass
            ei = ei + 1
        dest_guard.finish()
        src_guard.finish()


def _todo_task_help_temp_path():
    """Временный путь для встроенного ODT-файла справки."""
    import os
    import tempfile

    fd, path = tempfile.mkstemp(
        prefix="libre_macros_todo_task_help_",
        suffix=".odt",
        dir=tempfile.gettempdir(),
    )
    try:
        os.close(fd)
    except Exception:
        pass
    return path


def _open_path_in_writer(path):
    """Открыть файл в Writer / Atext."""
    import os

    import uno

    from todo_task_dialog_lib import _desktop

    path = os.path.abspath(unicode(path or u""))
    if not path or not os.path.isfile(path):
        return False, u"Файл не найден: %s" % path
    desktop = _desktop()
    if desktop is None:
        return False, u"Не удалось получить Desktop."
    try:
        url = uno.systemPathToFileUrl(path)
        desktop.loadComponentFromURL(url, "_blank", 0, ())
        return True, u""
    except Exception as exc:
        return False, unicode(exc)


def todo_task_help(doc=None):
    """Открыть встроенную справку по макросам серии todo_task в Writer / Atext."""
    tm = _StageTimer("todo_task_help")
    try:
        from todo_task_help_blob import write_todo_task_help_odt
    except Exception as exc:
        _msg(doc, u"Не удалось загрузить встроенную справку:\n%s" % exc, box_type=WARNINGBOX)
        tm.done("load_blob_fail")
        return False
    try:
        path = _todo_task_help_temp_path()
        write_todo_task_help_odt(path)
    except Exception as exc:
        _msg(doc, u"Не удалось подготовить файл справки:\n%s" % exc, box_type=WARNINGBOX)
        tm.done("write_blob_fail")
        return False
    ok, err = _open_path_in_writer(path)
    if not ok:
        _msg(
            doc,
            getattr(_cfg, "HELP_OPEN_FAILED_MSG", u"Не удалось открыть справку:\n%s")
            % err,
            box_type=WARNINGBOX,
        )
        tm.done("error")
        return False
    tm.done("ok")
    return True


def _header_col_map_for_merge(sheet):
    """norm_header(title) -> col_idx."""
    out = {}
    items = _list_sheet_header_titles(sheet)
    i = 0
    while i < len(items):
        col_idx, title = items[i]
        i = i + 1
        key = _norm_header(title)
        if key and key not in out:
            out[key] = int(col_idx)
    return out


def _copy_task_row_by_headers( src_sheet, src_row, dest_sheet, dest_row, skip_num=True, skip_keys=None, skip_dest_cols=None, values_only=True ):
    """Скопировать строку по совпадающим заголовкам (значения или с форматами)."""
    src_map = _header_col_map_for_merge(src_sheet)
    dest_map = _header_col_map_for_merge(dest_sheet)
    skip_norm = {}
    if skip_num:
        skip_norm[_norm_header(_cfg.COL_NUM)] = True
    sk = list(skip_keys or [])
    si = 0
    while si < len(sk):
        title = unicode(sk[si] or u"")
        si = si + 1
        if title:
            skip_norm[_norm_header(title)] = True
    skip_idx = {}
    sc = list(skip_dest_cols or [])
    ci = 0
    while ci < len(sc):
        try:
            skip_idx[int(sc[ci])] = True
        except (TypeError, ValueError):
            pass
        ci = ci + 1
    for key, src_col in src_map.items():
        if key in skip_norm:
            continue
        if key not in dest_map:
            continue
        dest_col = dest_map[key]
        if int(dest_col) in skip_idx:
            continue
        try:
            src_cell = src_sheet.getCellByPosition(int(src_col), int(src_row))
            dest_cell = dest_sheet.getCellByPosition(int(dest_col), int(dest_row))
            if values_only:
                _copy_cell_content(src_cell, dest_cell, values_only=True)
            else:
                try:
                    src_rng = src_sheet.getCellRangeByPosition(
                        int(src_col), int(src_row), int(src_col), int(src_row)
                    )
                    dest_sheet.copyRange(dest_cell.CellAddress, src_rng.RangeAddress)
                except Exception:
                    _copy_cell_content(src_cell, dest_cell, values_only=False)
        except Exception:
            try:
                _copy_cell_value_only(
                    src_sheet.getCellByPosition(int(src_col), int(src_row)),
                    dest_sheet.getCellByPosition(int(dest_col), int(dest_row)),
                )
            except Exception:
                pass


def _datetime_serial_from_parts(year, month, day, hour=0, minute=0, second=0):
    try:
        from datetime import date

        d = date(int(year), int(month), int(day))
    except (TypeError, ValueError):
        return None
    epoch = date(1899, 12, 30)
    days = (d - epoch).days
    frac = (
        float(int(hour)) * 3600.0
        + float(int(minute)) * 60.0
        + float(int(second))
    ) / 86400.0
    return float(days) + frac


def _parse_datetime_serial_from_text(text):
    s = unicode(text or u"").strip()
    if not s:
        return None
    m = re.match(
        r"^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})"
        r"(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?",
        s,
    )
    if not m:
        return None
    return _datetime_serial_from_parts(
        int(m.group(3)),
        int(m.group(2)),
        int(m.group(1)),
        int(m.group(4) or 0),
        int(m.group(5) or 0),
        int(m.group(6) or 0),
    )


def _cell_datetime_serial(cell):
    if cell is None:
        return None
    try:
        v = float(cell.Value)
        if v != 0.0:
            return v
    except Exception:
        pass
    try:
        text = _cell_text(cell)
    except Exception:
        text = u""
    return _parse_datetime_serial_from_text(text)


def _serial_to_seconds(serial):
    try:
        return int(round(float(serial) * 86400.0))
    except (TypeError, ValueError):
        return None


def _format_serial_timestamp(serial):
    if serial is None:
        return u""
    try:
        from datetime import datetime, timedelta

        total = _serial_to_seconds(serial)
        if total is None:
            return u""
        dt = datetime(1899, 12, 30) + timedelta(seconds=int(total))
        return dt.strftime(u"%d.%m.%Y %H:%M:%S")
    except Exception:
        return u""


def _timestamp_display_from_cell(cell):
    if cell is None:
        return u""
    text = u""
    try:
        if hasattr(cell, "getString"):
            text = unicode(cell.getString() or u"").strip()
        else:
            text = unicode(cell.String or u"").strip()
    except Exception:
        text = _cell_text(cell)
    if text:
        m = re.match(
            r"^(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4}"
            r"(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)",
            text,
        )
        if m:
            return m.group(1)
        parsed = _parse_datetime_serial_from_text(text)
        if parsed is not None:
            return _format_serial_timestamp(parsed)
        return text
    return _format_serial_timestamp(_cell_datetime_serial(cell))


def _keyed_cell(sheet, cols, row, key, title):
    if sheet is None or row is None:
        return None
    if cols and key in cols:
        try:
            return sheet.getCellByPosition(cols[key], int(row))
        except Exception:
            return None
    try:
        hdr = _header_col_map_for_merge(sheet)
        nk = _norm_header(title)
        if nk in hdr:
            return sheet.getCellByPosition(hdr[nk], int(row))
    except Exception:
        pass
    return None


def _modified_cell(sheet, cols, row):
    return _keyed_cell(sheet, cols, row, u"modified", _cfg.COL_MODIFIED)


def _comment_cell(sheet, cols, row):
    return _keyed_cell(sheet, cols, row, u"comment", _cfg.COL_COMMENT)


def _modified_by_cell(sheet, cols, row):
    return _keyed_cell(
        sheet,
        cols,
        row,
        u"modified_by",
        getattr(_cfg, "COL_MODIFIED_BY", u"Отредактировано"),
    )


def _modified_by_text(sheet, cols, row):
    cell = _modified_by_cell(sheet, cols, row)
    if cell is None:
        return u""
    return _cell_text(cell).strip()


def _person_name_key(name):
    return _sheet_name_key(name)


def _is_merge_editor_conflict(local_by, employee_by, owner_name):
    """
    Конфликт: локально задачу правил владелец книги, в своде другой редактор.
    """
    loc = unicode(local_by or u"").strip()
    emp = unicode(employee_by or u"").strip()
    own = unicode(owner_name or u"").strip()
    if not loc or not emp:
        return False
    if _person_name_key(loc) == _person_name_key(emp):
        return False
    if own:
        return _person_name_key(loc) == _person_name_key(own)
    return True


def _count_merge_editor_conflicts( src_sheet, src_cols, src_rows, local_index, owner_name ):
    n = 0
    ri = 0
    while ri < len(src_rows):
        src_row = int(src_rows[ri])
        ri = ri + 1
        if u"guid" not in src_cols:
            continue
        guid = _cell_text(
            src_sheet.getCellByPosition(src_cols[u"guid"], src_row)
        ).strip()
        if not guid:
            continue
        g_cf = guid.casefold() if hasattr(guid, "casefold") else guid.lower()
        if g_cf not in local_index:
            continue
        hit = local_index[g_cf]
        dest_sheet = hit.get(u"sheet")
        dest_cols = hit.get(u"cols")
        if dest_sheet is None or dest_cols is None:
            continue
        local_by = _modified_by_text(dest_sheet, dest_cols, int(hit[u"row"]))
        emp_by = _modified_by_text(src_sheet, src_cols, src_row)
        if _is_merge_editor_conflict(local_by, emp_by, owner_name):
            n = n + 1
    return n


def _incoming_modified_is_newer(src_sheet, src_cols, src_row, dest_sheet, dest_cols, dest_row):
    src_s = _cell_datetime_serial(_modified_cell(src_sheet, src_cols, src_row))
    dest_s = _cell_datetime_serial(
        _modified_cell(dest_sheet, dest_cols, dest_row)
    )
    if src_s is None:
        return False
    if dest_s is None:
        return True
    src_sec = _serial_to_seconds(src_s)
    dest_sec = _serial_to_seconds(dest_s)
    if src_sec is None or dest_sec is None:
        return False
    return src_sec > dest_sec


def _split_comment_stamp_blocks(text):
    """Разобрать комментарий на блоки ====штамп==== / преамбула без штампа."""
    raw = unicode(text or u"").strip()
    if not raw:
        return []
    matches = list(
        re.finditer(
            r"====([^=\n]+)====\n?(.*?)(?=\n====|\Z)",
            raw,
            re.DOTALL,
        )
    )
    if not matches:
        return [{u"stamp": u"", u"body": raw}]
    out = []
    pre = raw[: matches[0].start()].strip()
    if pre:
        out.append({u"stamp": u"", u"body": pre})
    i = 0
    while i < len(matches):
        m = matches[i]
        i = i + 1
        out.append(
            {
                u"stamp": unicode(m.group(1) or u"").strip(),
                u"body": unicode(m.group(2) or u"").strip(),
            }
        )
    return out


def _format_comment_stamp_blocks(blocks):
    parts = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        i = i + 1
        stamp = unicode((b or {}).get(u"stamp") or u"").strip()
        body = unicode((b or {}).get(u"body") or u"").strip()
        if not stamp and not body:
            continue
        if stamp:
            parts.append(u"====%s====\n%s" % (stamp, body))
        else:
            parts.append(body)
    return u"\n\n".join(parts)


def _comment_stamp_block_key(block):
    b = block or {}
    return (
        unicode(b.get(u"stamp") or u"").strip(),
        unicode(b.get(u"body") or u"").strip(),
    )


def _now_comment_stamp():
    try:
        from datetime import datetime

        return datetime.now().strftime(u"%d.%m.%Y %H:%M:%S")
    except Exception:
        return u""


def _accumulate_comment_text( local_text, incoming_text, incoming_stamp, local_stamp ):
    """Накопить комментарии: сверху новые блоки, ниже более ранние."""
    loc = unicode(local_text or u"").rstrip()
    inc = unicode(incoming_text or u"").strip()
    if not inc:
        return loc
    loc_n = loc.strip()
    if loc_n == inc:
        return loc if loc else inc
    if loc_n and inc in loc_n:
        return loc
    if loc_n and inc.startswith(loc_n):
        extra = inc[len(loc_n) :].strip()
        if not extra:
            return loc
        inc = extra
    stamp = unicode(incoming_stamp or u"").strip()
    if not stamp:
        stamp = _now_comment_stamp()
    loc_blocks = _split_comment_stamp_blocks(loc)
    if len(loc_blocks) == 1 and not loc_blocks[0].get(u"stamp"):
        lstamp = unicode(local_stamp or u"").strip()
        if lstamp:
            loc_blocks[0][u"stamp"] = lstamp
    inc_blocks = _split_comment_stamp_blocks(inc)
    if len(inc_blocks) == 1 and not inc_blocks[0].get(u"stamp"):
        inc_blocks[0][u"stamp"] = stamp
    have = {}
    have_body = {}
    li = 0
    while li < len(loc_blocks):
        b = loc_blocks[li]
        li = li + 1
        have[_comment_stamp_block_key(b)] = True
        body = unicode(b.get(u"body") or u"").strip()
        if body:
            have_body[body] = True
    new_blocks = []
    ii = 0
    while ii < len(inc_blocks):
        b = inc_blocks[ii]
        ii = ii + 1
        key = _comment_stamp_block_key(b)
        body = unicode(b.get(u"body") or u"").strip()
        if key in have:
            continue
        if body and body in have_body:
            continue
        new_blocks.append(b)
        have[key] = True
        if body:
            have_body[body] = True
    if not new_blocks:
        return loc
    return _format_comment_stamp_blocks(new_blocks + loc_blocks)


def _collect_local_tasks_index(doc, task_pattern):
    """
    Индекс локальных задач: guid (casefold) -> {sheet, cols, row, sheet_name, category_b1}.
    category_by_sheet: имя листа -> категория из B1.
    """
    index = {}
    category_by_sheet = {}
    sheets = _iter_sheets_matching_pattern(doc, task_pattern)
    si = 0
    while si < len(sheets):
        sh = sheets[si]
        si = si + 1
        cols = _find_headers(sh)
        if cols is None:
            continue
        name = _sheet_display_name(sh)
        if _is_merge_source_sheet(sh, doc=doc, name=name):
            continue
        if not (
            _is_named_task_sheet(sh, doc=doc)
            or _sheet_name_matches_pattern(name, task_pattern)
        ):
            continue
        cat_b1 = _sheet_category_from_b1(sh)
        category_by_sheet[name] = cat_b1
        rows = _iter_task_rows(sh, cols)
        ri = 0
        while ri < len(rows):
            row = int(rows[ri])
            ri = ri + 1
            if u"guid" not in cols:
                continue
            guid = _cell_text(sh.getCellByPosition(cols[u"guid"], row)).strip()
            if not guid:
                continue
            g_cf = guid.casefold() if hasattr(guid, "casefold") else guid.lower()
            index[g_cf] = {
                u"sheet": sh,
                u"cols": cols,
                u"row": row,
                u"sheet_name": name,
                u"category_b1": cat_b1,
            }
    return index, category_by_sheet


def _category_to_sheet_lookup(category_by_sheet):
    """Категория (B1, casefold) -> имя листа задач."""
    out = {}
    for sheet_name, cat in (category_by_sheet or {}).items():
        cat = unicode(cat or u"").strip()
        if not cat:
            continue
        key = cat.casefold() if hasattr(cat, "casefold") else cat.lower()
        out[key] = unicode(sheet_name)
    return out


def _summary_row_category(sheet, cols, row):
    if u"category" in cols:
        return _cell_text(sheet.getCellByPosition(cols[u"category"], row)).strip()
    cat_title = getattr(_cfg, "COL_CATEGORY", u"Категория")
    hdr_map = _header_col_map_for_merge(sheet)
    key = _norm_header(cat_title)
    if key in hdr_map:
        return _cell_text(
            sheet.getCellByPosition(hdr_map[key], row)
        ).strip()
    return u""


def _resolve_dest_sheet_for_category(doc, category, cat_lookup):
    cat = unicode(category or u"").strip()
    if not cat:
        return None, u""
    key = cat.casefold() if hasattr(cat, "casefold") else cat.lower()
    dest_name = cat_lookup.get(key)
    if not dest_name:
        return None, u""
    sh = _get_sheet_by_name(doc, dest_name)
    if sh is None:
        return None, dest_name
    return sh, dest_name


def _insert_task_from_summary( src_sheet, src_row, dest_sheet, dest_cols, doc=None, values_only=True ):
    """Новая задача на листе «Задачи…» из строки свода (GUID сохраняется)."""
    tasks_before = _iter_task_rows(dest_sheet, dest_cols)
    dest_row = _next_task_append_row(dest_sheet, dest_cols)
    if tasks_before:
        _copy_row_formats(dest_sheet, dest_cols, tasks_before[-1], dest_row)
    else:
        hdr = _header_row(dest_cols)
        template = hdr + 1
        try:
            if template <= _used_last_row(dest_sheet):
                _copy_row_formats(dest_sheet, dest_cols, template, dest_row)
        except Exception:
            pass
    _copy_task_row_by_headers(
        src_sheet,
        src_row,
        dest_sheet,
        dest_row,
        skip_num=True,
        values_only=values_only,
    )
    next_num = len(tasks_before) + 1
    _ensure_row_number(dest_sheet, dest_cols, dest_row, next_num)
    _autofit_task_rows(dest_sheet, [dest_row])
    return dest_row


def _update_task_from_summary( src_sheet, src_row, dest_sheet, dest_cols, dest_row, skip_keys=None, skip_dest_cols=None, values_only=True ):
    """Обновить существующую задачу (кроме «№ п/п»)."""
    _copy_task_row_by_headers(
        src_sheet,
        src_row,
        dest_sheet,
        int(dest_row),
        skip_num=True,
        skip_keys=skip_keys,
        skip_dest_cols=skip_dest_cols,
        values_only=values_only,
    )
    _autofit_task_rows(dest_sheet, [int(dest_row)])


def _merge_task_name_preview(sheet, cols, row, maxlen=None):
    if maxlen is None:
        maxlen = int(getattr(_cfg, "MERGE_REPORT_NAME_PREVIEW", 48))
    name = u""
    try:
        if cols is not None and u"name" in cols:
            name = _cell_text(sheet.getCellByPosition(cols[u"name"], int(row))).strip()
    except Exception:
        name = u""
    if len(name) > int(maxlen):
        return name[: int(maxlen)] + u"…"
    return name or u"—"


def _plan_merge_counts( src_sheet, src_cols, src_rows, local_index, category_by_sheet, cat_lookup, doc, update_existing_only, check_modified ):
    """Сухой прогон: сколько добавить / обновить / пропустить."""
    plan_ins = 0
    plan_upd = 0
    plan_skip = 0
    ri = 0
    while ri < len(src_rows):
        src_row = int(src_rows[ri])
        ri = ri + 1
        guid = _cell_text(
            src_sheet.getCellByPosition(src_cols[u"guid"], src_row)
        ).strip()
        if not guid:
            plan_skip = plan_skip + 1
            continue
        g_cf = guid.casefold() if hasattr(guid, "casefold") else guid.lower()
        # уже есть у руководителя — обновлять по GUID на текущем листе
        if g_cf in local_index:
            hit = local_index[g_cf]
            dest_sheet = hit.get(u"sheet")
            dest_cols = hit.get(u"cols")
            if dest_sheet is None or dest_cols is None:
                plan_skip = plan_skip + 1
                continue
            if check_modified:
                comment_newer = _incoming_modified_is_newer(
                    src_sheet,
                    src_cols,
                    src_row,
                    dest_sheet,
                    dest_cols,
                    int(hit[u"row"]),
                )
                if not comment_newer:
                    plan_skip = plan_skip + 1
                    continue
            plan_upd = plan_upd + 1
            continue
        # новой задачи нет — лист по категории свода
        category = _summary_row_category(src_sheet, src_cols, src_row)
        dest_sheet, dest_name = _resolve_dest_sheet_for_category(
            doc, category, cat_lookup
        )
        if dest_sheet is None:
            plan_skip = plan_skip + 1
            continue
        if _find_headers(dest_sheet) is None:
            plan_skip = plan_skip + 1
            continue
        if update_existing_only:
            plan_skip = plan_skip + 1
            continue
        plan_ins = plan_ins + 1
    return plan_ins, plan_upd, plan_skip


def _format_merge_report_line(kind, guid, name, sheet_name, reason=u""):
    g = unicode(guid or u"").strip() or u"—"
    n = unicode(name or u"").strip() or u"—"
    sh = unicode(sheet_name or u"").strip() or u"—"
    r = unicode(reason or u"").strip()
    if kind == u"insert":
        if r:
            return u"+ %s | %s → %s (%s)" % (g, n, sh, r)
        return u"+ %s | %s → %s" % (g, n, sh)
    if kind == u"update":
        if r:
            return u"~ %s | %s → %s (%s)" % (g, n, sh, r)
        return u"~ %s | %s → %s" % (g, n, sh)
    if not r:
        r = u"пропуск"
    return u"− %s | %s | %s: %s" % (g, n, sh, r)


def todo_task_merge(doc=None):
    """
    Слияние задач с листа свода сотрудников на листы «Задачи…».
    Существующие — обновление по GUID на текущем листе руководителя;
    новые — на лист по категории из свода.
    Доступно только при роли владельца «руководитель».
    """
    tm = _StageTimer("todo_task_merge")
    doc = _get_doc(doc)
    if doc is None:
        _msg(doc, u"Нет открытого документа Calc.", box_type=WARNINGBOX)
        return False

    try:
        from todo_task_settings_lib import (
            get_merge_options,
            get_merge_source_sheet,
            get_owner_role,
            load_todo_task_settings,
            normalize_layout_settings,
        )
    except Exception as exc:
        _msg(doc, u"Настройки задач: %s" % exc, box_type=WARNINGBOX)
        return False

    settings = load_todo_task_settings(doc=doc)
    if get_owner_role(settings) != getattr(_cfg, "OWNER_ROLE_MANAGER", u"manager"):
        _msg(
            doc,
            getattr(
                _cfg,
                "MERGE_ROLE_DENIED_MSG",
                u"Слияние доступно только руководителю.",
            ),
            box_type=WARNINGBOX,
        )
        tm.done("role_denied")
        return False

    lay = normalize_layout_settings((settings or {}).get(u"layout"))
    task_pat = unicode(lay.get(u"task_sheet_pattern") or u"Задачи*")
    merge_sheet_name = get_merge_source_sheet(settings)
    merge_opts = get_merge_options(settings)
    update_existing_only = bool(merge_opts.get(u"update_existing_only"))
    check_modified = bool(merge_opts.get(u"check_modified"))
    accumulate_comments = bool(merge_opts.get(u"accumulate_comments"))

    src_sheet = _get_sheet_by_name(doc, merge_sheet_name)
    if src_sheet is None:
        _msg(
            doc,
            getattr(_cfg, "MERGE_SOURCE_MISSING_MSG", u"Нет листа «%s».")
            % merge_sheet_name,
            box_type=WARNINGBOX,
        )
        tm.done("no_source")
        return False

    src_cols = _find_headers(src_sheet)
    if src_cols is None or u"guid" not in src_cols:
        _msg(
            doc,
            u"На листе «%s» нет таблицы задач (колонка GUID)."
            % merge_sheet_name,
            box_type=WARNINGBOX,
        )
        tm.done("no_src_table")
        return False

    local_index, category_by_sheet = _collect_local_tasks_index(doc, task_pat)
    if not category_by_sheet:
        _msg(
            doc,
            getattr(_cfg, "MERGE_NO_TASK_SHEETS_MSG", u"Нет листов задач.")
            % task_pat,
            box_type=WARNINGBOX,
        )
        tm.done("no_dest_sheets")
        return False

    cat_lookup = _category_to_sheet_lookup(category_by_sheet)
    src_rows = _iter_task_rows(src_sheet, src_cols)
    if not src_rows:
        _msg(doc, u"На листе «%s» нет задач." % merge_sheet_name, box_type=WARNINGBOX)
        tm.done("empty_source")
        return False

    plan_ins, plan_upd, plan_skip = _plan_merge_counts(
        src_sheet,
        src_cols,
        src_rows,
        local_index,
        category_by_sheet,
        cat_lookup,
        doc,
        update_existing_only,
        check_modified,
    )

    merge_guids = []
    ri = 0
    while ri < len(src_rows):
        g = _cell_text(
            src_sheet.getCellByPosition(src_cols[u"guid"], int(src_rows[ri]))
        ).strip()
        if g:
            merge_guids.append(g)
        ri = ri + 1

    plan_preview = {}
    plan_preview_line = u""
    try:
        from todo_task_plan_lib import (
            format_merge_plan_preview_line,
            preview_plan_transfer_for_merge,
        )

        plan_preview = preview_plan_transfer_for_merge(doc, merge_guids)
        plan_preview_line = format_merge_plan_preview_line(plan_preview)
    except Exception as exc:
        _log("merge plan preview: %s" % exc)
        plan_preview_line = u"Планы: не удалось проверить листы в книге."

    from todo_task_dialog_lib import (
        show_todo_task_merge_confirm_dialog,
        show_todo_task_merge_editor_conflict_dialog,
        show_todo_task_merge_report_dialog,
        show_todo_task_progress_dialog,
    )

    try:
        run_spec = show_todo_task_merge_confirm_dialog(
            doc=doc,
            summary={
                u"source_name": merge_sheet_name,
                u"total": len(src_rows),
                u"plan_insert": plan_ins,
                u"plan_update": plan_upd,
                u"plan_skip": plan_skip,
                u"plan_preview_line": plan_preview_line,
            },
            initial={
                u"accumulate_comments": bool(
                    getattr(
                        _cfg, "DEFAULT_MERGE_RUN_ACCUMULATE_COMMENTS", True
                    )
                ),
                u"values_only": bool(
                    getattr(_cfg, "DEFAULT_MERGE_RUN_VALUES_ONLY", True)
                ),
            },
        )
    except Exception as exc:
        _msg(doc, u"Диалог слияния: %s" % exc, box_type=WARNINGBOX)
        tm.done("confirm_fail")
        return False
    if run_spec is None:
        tm.done("cancel")
        return False

    accumulate_comments = bool(run_spec.get(u"accumulate_comments"))
    values_only = bool(run_spec.get(u"values_only"))
    if run_spec.get(u"values_only") is None:
        values_only = True

    owner_name = _get_owner_display_name(settings, doc=doc)
    act_keep = unicode(
        getattr(_cfg, "MERGE_EDITOR_KEEP_OWNER", u"keep_owner")
    )
    act_take = unicode(
        getattr(_cfg, "MERGE_EDITOR_TAKE_EMPLOYEE", u"take_employee")
    )
    editor_conflict_total = _count_merge_editor_conflicts(
        src_sheet, src_cols, src_rows, local_index, owner_name
    )
    editor_apply_all = None
    editor_conflict_i = 0
    merge_aborted = False

    progress = None
    try:
        progress = show_todo_task_progress_dialog(
            doc=doc,
            title=getattr(_cfg, "MERGE_DIALOG_PROGRESS_TITLE", u"Слияние задач"),
            text=u"Подготовка…",
        )
    except Exception as err:
        _log("merge progress: %s" % err)

    inserted = 0
    updated = 0
    skipped = 0
    report_lines = []
    merged_guids = []
    touched_sheets = {}
    total = len(src_rows)

    ri = 0
    while ri < total:
        src_row = int(src_rows[ri])
        ri = ri + 1
        if progress is not None:
            try:
                progress.update(
                    u"Строка %s / %s свода…" % (ri, total),
                    value=ri,
                    maximum=max(total, 1),
                )
            except Exception:
                pass

        task_name = _merge_task_name_preview(src_sheet, src_cols, src_row)
        guid = _cell_text(
            src_sheet.getCellByPosition(src_cols[u"guid"], src_row)
        ).strip()
        if not guid:
            skipped = skipped + 1
            report_lines.append(
                _format_merge_report_line(
                    u"skip", u"", task_name, u"—", u"нет GUID"
                )
            )
            continue
        g_cf = guid.casefold() if hasattr(guid, "casefold") else guid.lower()

        # Обновление: только по GUID, на том листе где задача уже у руководителя.
        if g_cf in local_index:
            hit = local_index[g_cf]
            dest_sheet = hit.get(u"sheet")
            dest_cols = hit.get(u"cols")
            dest_row = int(hit.get(u"row"))
            dest_key = unicode(hit.get(u"sheet_name") or u"")
            if dest_sheet is None or dest_cols is None or not dest_key:
                skipped = skipped + 1
                report_lines.append(
                    _format_merge_report_line(
                        u"skip",
                        guid,
                        task_name,
                        dest_key or u"—",
                        u"повреждён индекс GUID",
                    )
                )
                continue
            if dest_key not in touched_sheets:
                guard = _TaskSheetProtectSession(dest_sheet, doc=doc)
                if not guard.begin():
                    _msg(doc, _protect_unlock_error_text(), box_type=WARNINGBOX)
                    if progress is not None:
                        progress.close()
                    tm.done("protect_fail")
                    return False
                touched_sheets[dest_key] = {
                    u"sheet": dest_sheet,
                    u"cols": dest_cols,
                    u"guard": guard,
                    u"fit_rows": [],
                }
            bucket = touched_sheets[dest_key]
            dest_cols = bucket[u"cols"]

            local_by = _modified_by_text(dest_sheet, dest_cols, dest_row)
            emp_by = _modified_by_text(src_sheet, src_cols, src_row)
            if _is_merge_editor_conflict(local_by, emp_by, owner_name):
                editor_conflict_i = editor_conflict_i + 1
                action = editor_apply_all
                if not action:
                    remaining = max(
                        int(editor_conflict_total) - editor_conflict_i + 1, 1
                    )
                    try:
                        spec = show_todo_task_merge_editor_conflict_dialog(
                            doc=doc,
                            task_name=task_name,
                            guid=guid,
                            sheet_name=dest_key,
                            owner_editor=local_by or owner_name,
                            employee_editor=emp_by,
                            owner_modified=_timestamp_display_from_cell(
                                _modified_cell(
                                    dest_sheet, dest_cols, dest_row
                                )
                            ),
                            employee_modified=_timestamp_display_from_cell(
                                _modified_cell(
                                    src_sheet, src_cols, src_row
                                )
                            ),
                            remaining=remaining,
                            conflict_index=editor_conflict_i,
                            conflict_total=max(int(editor_conflict_total), 1),
                        )
                    except Exception as exc:
                        _log("merge editor conflict dialog: %s" % exc)
                        spec = None
                    if spec is None:
                        merge_aborted = True
                        report_lines.append(
                            _format_merge_report_line(
                                u"skip",
                                guid,
                                task_name,
                                dest_key,
                                u"слияние прервано (конфликт редакторов)",
                            )
                        )
                        skipped = skipped + 1
                        break
                    action = unicode(spec.get(u"action") or u"").strip()
                    if not action:
                        merge_aborted = True
                        report_lines.append(
                            _format_merge_report_line(
                                u"skip",
                                guid,
                                task_name,
                                dest_key,
                                u"слияние прервано (конфликт редакторов)",
                            )
                        )
                        skipped = skipped + 1
                        break
                    if spec.get(u"apply_all"):
                        editor_apply_all = action
                if action == act_keep:
                    skipped = skipped + 1
                    report_lines.append(
                        _format_merge_report_line(
                            u"skip",
                            guid,
                            task_name,
                            dest_key,
                            u"оставлена версия владельца",
                        )
                    )
                    _log(
                        "merge keep owner guid=%s local_by=%s emp_by=%s"
                        % (guid, local_by, emp_by)
                    )
                    continue
                if action == act_take:
                    # Полная замена всех колонок из свода; форматы — как в confirm.
                    _update_task_from_summary(
                        src_sheet,
                        src_row,
                        dest_sheet,
                        dest_cols,
                        dest_row,
                        skip_keys=[],
                        skip_dest_cols=[],
                        values_only=values_only,
                    )
                    bucket[u"fit_rows"].append(dest_row)
                    updated = updated + 1
                    merged_guids.append(guid)
                    report_lines.append(
                        _format_merge_report_line(
                            u"update",
                            guid,
                            task_name,
                            dest_key,
                            u"версия сотрудника (конфликт редакторов)",
                        )
                    )
                    _log(
                        "merge take employee guid=%s local_by=%s emp_by=%s"
                        % (guid, local_by, emp_by)
                    )
                    continue
                skipped = skipped + 1
                report_lines.append(
                    _format_merge_report_line(
                        u"skip",
                        guid,
                        task_name,
                        dest_key,
                        u"неизвестный выбор конфликта",
                    )
                )
                continue

            comment_newer = _incoming_modified_is_newer(
                src_sheet,
                src_cols,
                src_row,
                dest_sheet,
                dest_cols,
                dest_row,
            )
            if check_modified and not comment_newer:
                skipped = skipped + 1
                report_lines.append(
                    _format_merge_report_line(
                        u"skip",
                        guid,
                        task_name,
                        dest_key,
                        u"дата редакции свода не новее",
                    )
                )
                _log(
                    "merge skip guid=%s: incoming modified is not newer"
                    % guid
                )
                continue
            skip_keys = []
            skip_dest_cols = []
            local_cmt = u""
            incoming_cmt = u""
            inc_stamp = u""
            loc_stamp = u""
            if accumulate_comments:
                skip_keys.append(
                    getattr(_cfg, "COL_COMMENT", u"Комментарии")
                )
                if u"comment" in dest_cols:
                    skip_dest_cols.append(dest_cols[u"comment"])
                dest_cmt_cell = _comment_cell(
                    dest_sheet, dest_cols, dest_row
                )
                src_cmt_cell = _comment_cell(
                    src_sheet, src_cols, src_row
                )
                local_cmt = (
                    _cell_text(dest_cmt_cell)
                    if dest_cmt_cell is not None
                    else u""
                )
                incoming_cmt = (
                    _cell_text(src_cmt_cell)
                    if src_cmt_cell is not None
                    else u""
                )
                inc_stamp = _timestamp_display_from_cell(
                    _modified_cell(src_sheet, src_cols, src_row)
                )
                loc_stamp = _timestamp_display_from_cell(
                    _modified_cell(dest_sheet, dest_cols, dest_row)
                )
            _update_task_from_summary(
                src_sheet,
                src_row,
                dest_sheet,
                dest_cols,
                dest_row,
                skip_keys=skip_keys,
                skip_dest_cols=skip_dest_cols,
                values_only=values_only,
            )
            if accumulate_comments and comment_newer:
                dest_cmt_cell = _comment_cell(
                    dest_sheet, dest_cols, dest_row
                )
                if dest_cmt_cell is not None:
                    merged_cmt = _accumulate_comment_text(
                        local_cmt,
                        incoming_cmt,
                        inc_stamp,
                        loc_stamp,
                    )
                    _set_cell_text(dest_cmt_cell, merged_cmt)
            bucket[u"fit_rows"].append(dest_row)
            updated = updated + 1
            merged_guids.append(guid)
            report_lines.append(
                _format_merge_report_line(
                    u"update", guid, task_name, dest_key
                )
            )
            continue

        # Новая задача — лист по категории из свода.
        category = _summary_row_category(src_sheet, src_cols, src_row)
        dest_sheet, dest_name = _resolve_dest_sheet_for_category(
            doc, category, cat_lookup
        )
        if dest_sheet is None:
            skipped = skipped + 1
            if category:
                reason = u"нет листа для категории «%s»" % category
            else:
                reason = u"пустая категория"
            report_lines.append(
                _format_merge_report_line(
                    u"skip", guid, task_name, u"—", reason
                )
            )
            if dest_name:
                _log("merge skip guid=%s: no sheet for category %s" % (guid, category))
            else:
                _log(
                    "merge skip guid=%s: unknown category «%s»"
                    % (guid, category)
                )
            continue

        dest_cols = _find_headers(dest_sheet)
        if dest_cols is None:
            skipped = skipped + 1
            report_lines.append(
                _format_merge_report_line(
                    u"skip",
                    guid,
                    task_name,
                    dest_name or u"—",
                    u"нет таблицы задач",
                )
            )
            continue

        try:
            dest_key = unicode(dest_sheet.Name)
        except Exception:
            dest_key = dest_name or u"?"

        if dest_key not in touched_sheets:
            guard = _TaskSheetProtectSession(dest_sheet, doc=doc)
            if not guard.begin():
                _msg(doc, _protect_unlock_error_text(), box_type=WARNINGBOX)
                if progress is not None:
                    progress.close()
                tm.done("protect_fail")
                return False
            touched_sheets[dest_key] = {
                u"sheet": dest_sheet,
                u"cols": dest_cols,
                u"guard": guard,
                u"fit_rows": [],
            }

        bucket = touched_sheets[dest_key]
        dest_cols = bucket[u"cols"]

        if update_existing_only:
            skipped = skipped + 1
            report_lines.append(
                _format_merge_report_line(
                    u"skip",
                    guid,
                    task_name,
                    dest_key,
                    u"только существующие — GUID не найден",
                )
            )
            _log(
                "merge skip guid=%s: not on local sheets (existing only)"
                % guid
            )
            continue
        dest_row = _insert_task_from_summary(
            src_sheet,
            src_row,
            dest_sheet,
            dest_cols,
            doc=doc,
            values_only=values_only,
        )
        bucket[u"fit_rows"].append(int(dest_row))
        local_index[g_cf] = {
            u"sheet": dest_sheet,
            u"cols": dest_cols,
            u"row": int(dest_row),
            u"sheet_name": dest_key,
            u"category_b1": category_by_sheet.get(dest_key, u""),
        }
        inserted = inserted + 1
        merged_guids.append(guid)
        report_lines.append(
            _format_merge_report_line(
                u"insert", guid, task_name, dest_key
            )
        )

    for dest_key in touched_sheets:
        bucket = touched_sheets[dest_key]
        sh = bucket[u"sheet"]
        cols = bucket[u"cols"]
        left = _iter_task_rows(sh, cols)
        _renumber_task_rows(sh, cols, left)

    for dest_key in touched_sheets:
        try:
            touched_sheets[dest_key][u"guard"].finish()
        except Exception:
            pass

    plan_xfer = {}
    if not merge_aborted:
        try:
            from todo_task_plan_lib import transfer_employee_plans_on_merge

            plan_xfer = transfer_employee_plans_on_merge(doc, merge_guids)
        except Exception as exc:
            plan_xfer = {u"errors": [unicode(exc)]}
            _log("merge plan transfer: %s" % exc)

    if progress is not None:
        try:
            progress.close()
        except Exception:
            pass

    tm.done(
        "ABORT" if merge_aborted else "DONE",
        "ins=%s upd=%s skip=%s" % (inserted, updated, skipped),
    )
    cmt_mode = (
        u"объединение комментариев"
        if accumulate_comments
        else u"полная замена комментариев"
    )
    copy_mode = (
        u"только значения" if values_only else u"с форматами"
    )
    if merge_aborted:
        head = u"Слияние прервано (конфликт редакторов).\n"
    else:
        head = u"Слияние завершено.\n"
    summary_text = head + (
        u"Добавлено: %s; обновлено: %s; пропущено: %s\n"
        u"Источник: «%s»\n"
        u"%s; %s"
        % (
            inserted,
            updated,
            skipped,
            merge_sheet_name,
            cmt_mode,
            copy_mode,
        )
    )
    if plan_preview_line:
        summary_text = summary_text + u"\n" + plan_preview_line
    if plan_xfer:
        moved_n = len(plan_xfer.get(u"moved") or [])
        miss_n = len(plan_xfer.get(u"missing") or [])
        err_n = len(plan_xfer.get(u"errors") or [])
        if not moved_n and not err_n:
            summary_text = summary_text + u"\n" + unicode(
                getattr(_cfg, "MERGE_PLAN_REPORT_NONE", u"")
            )
        else:
            summary_text = summary_text + u"\n" + unicode(
                getattr(_cfg, "MERGE_PLAN_REPORT_MOVED", u"")
            ) % moved_n
            if miss_n:
                summary_text = summary_text + u"; " + unicode(
                    getattr(_cfg, "MERGE_PLAN_REPORT_SKIPPED", u"")
                ) % miss_n
            if err_n:
                summary_text = summary_text + u"; " + unicode(
                    getattr(_cfg, "MERGE_PLAN_REPORT_ERRORS", u"")
                ) % err_n
        report_lines.append(unicode(getattr(_cfg, "MERGE_PLAN_REPORT_HEADER", u"")))
        if moved_n:
            report_lines.extend(plan_xfer.get(u"lines") or [])
        elif not moved_n and not err_n:
            report_lines.append(
                unicode(getattr(_cfg, "MERGE_PLAN_REPORT_NONE", u""))
            )
        ei = 0
        while ei < err_n:
            report_lines.append(u"! " + unicode(plan_xfer[u"errors"][ei]))
            ei = ei + 1
    try:
        show_todo_task_merge_report_dialog(
            doc=doc, summary_text=summary_text, lines=report_lines
        )
    except Exception as exc:
        _msg(doc, summary_text + u"\n\n(список: %s)" % exc)
    return inserted > 0 or updated > 0
