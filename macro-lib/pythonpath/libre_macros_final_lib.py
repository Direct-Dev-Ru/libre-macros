# -*- coding: utf-8 -*-
from __future__ import print_function
"""
Финальная обработка книги после сбора (удаление/скрытие листов, пользовательские шаги).

Подключается из libre_macros_lib.py; публичный API — lm_final_*.
"""
MACRO_VERSION = "3.10.689"
import re

import libre_macros_lib as lm
import libre_macros_values_lib as vl

_LM_FINAL_DEFAULT_DELETE_SHEETS = (
    "Справка_макроса",
    "Сбор_книг_лог",
    "Индексы_Источников",
)

_LM_FINAL_BUILTIN_KEYS = frozenset(
    ("удаление_листов", "скрытие_листов")
)
_LM_FINAL_LEGACY_DELETE_KEYS = frozenset(
    ("удалить_служебные_листы", "удалить_листы")
)
_LM_FINAL_SKIP_PIVOT_FOLLOWUP_KEYS = (
    _LM_FINAL_BUILTIN_KEYS | _LM_FINAL_LEGACY_DELETE_KEYS
)

_LM_FINAL_SHEET_SPEC_FIRST = frozenset(("_первый_",))
_LM_FINAL_SHEET_SPEC_LAST = frozenset(("_последний_",))
_LM_FINAL_SHEET_SPEC_LAST_GARBAGE = frozenset(
    ("_последний_мусорный_", "_мусорный_последний_")
)


def _lm_final_postprocess_fn_key(name):
    return str(name or "").strip().casefold()


def _lm_final_is_builtin_name(fn_name):
    return _lm_final_postprocess_fn_key(fn_name) in _LM_FINAL_BUILTIN_KEYS


def _lm_final_is_pivot_followup_skip_name(fn_name):
    return _lm_final_postprocess_fn_key(fn_name) in _LM_FINAL_SKIP_PIVOT_FOLLOWUP_KEYS


def _lm_final_log(doc, fn_label, status, note=""):
    lm._lm_log_postprocess(
        doc,
        "",
        "финальная_обработка",
        str(fn_label),
        str(status),
        str(note) if note != "" else "",
    )


def _lm_final_collect_delete_protected_sheet_names(doc):
    """
    Листы, которые нельзя удалять служебными именами _Первый_ / _Последний_.
    Явное имя в колонке C по-прежнему удаляет лист безусловно.
    """
    protected = set()
    if doc is None:
        return protected
    try:
        import libre_macros_pivot_lib as pl

        for entry in pl.lm_pp_pivot_list_materialized_sheets():
            if isinstance(entry, dict):
                nm = entry.get("name")
            else:
                nm = entry
            s = str(nm or "").strip()
            if s != "":
                protected.add(s)
        for nm in pl.lm_pp_pivot_list_as_values_pending():
            s = str(nm or "").strip()
            if s != "":
                protected.add(s)
    except Exception:
        pass
    try:
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            try:
                name = str(sheets.getByIndex(i).Name)
                if lm._lm_pp_is_param_sheet_name(name):
                    protected.add(name)
            except Exception:
                pass
            i = i + 1
    except Exception:
        pass
    try:
        import libre_macros_collect_cfg as cw

        wiz = str(getattr(cw, "MERGE_WIZARD_LISTS_SHEET_NAME", "") or "").strip()
        if wiz != "":
            protected.add(wiz)
    except Exception:
        protected.add("_Визард_Списки")
    return protected


def _lm_final_is_materialization_garbage_sheet_name(name):
    n = str(name or "").strip()
    if n == "":
        return False
    try:
        if vl.lm_values_is_lo_default_sheet_name(n):
            return True
    except Exception:
        if re.match(r"^Sheet\d+$", n, re.IGNORECASE):
            return True
        if re.match(r"^Лист\d+$", n, re.IGNORECASE):
            return True
    if n.startswith("__pp_mat") or n.startswith("_to_delete_"):
        return True
    return False


def _lm_final_list_sheet_names(doc):
    names = []
    if doc is None:
        return names
    try:
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            try:
                names.append(str(sheets.getByIndex(i).Name))
            except Exception:
                pass
            i = i + 1
    except Exception:
        pass
    return names


def _lm_final_delete_one_sheet(doc, actual):
    """
    Удалить лист по имени (hasByName + поиск без учёта регистра).
    Возвращает (deleted_name, error_text).
    """
    target = str(actual or "").strip()
    if target == "" or doc is None:
        return None, "пустое имя листа"
    try:
        sheets = doc.Sheets
        if sheets.hasByName(target):
            sheets.removeByName(target)
            return target, ""
    except Exception as err:
        return None, "removeByName(%s): %s" % (target, err)
    key = target.casefold()
    try:
        sheets = doc.Sheets
        i = 0
        while i < sheets.getCount():
            try:
                name = str(sheets.getByIndex(i).Name)
            except Exception:
                i = i + 1
                continue
            if name.casefold() == key:
                try:
                    sheets.removeByName(name)
                    return name, ""
                except Exception as err:
                    return None, "removeByName(%s): %s" % (name, err)
            i = i + 1
    except Exception as err:
        return None, str(err)
    all_names = _lm_final_list_sheet_names(doc)
    return None, "лист «%s» не найден; в книге (%d): %s" % (
        target,
        len(all_names),
        ", ".join(all_names),
    )


def _lm_final_spec_is_sheet_pattern(spec):
    """Шаблон имени листа: *, ? или «Все» (не спецтокены _Первый_ и т.п.)."""
    s = str(spec or "").strip()
    if s == "":
        return False
    key = s.casefold()
    if key in _LM_FINAL_SHEET_SPEC_FIRST:
        return False
    if key in _LM_FINAL_SHEET_SPEC_LAST:
        return False
    if key in _LM_FINAL_SHEET_SPEC_LAST_GARBAGE:
        return False
    if re.fullmatch(r"\d+", s):
        return False
    if lm.lm_identity_key(s) == "все" or s == "*":
        return True
    return ("*" in s) or ("?" in s)


def _lm_final_expand_sheet_pattern(doc, pattern, protected=None):
    """
    Раскрыть шаблон имени листа в список фактических имён (text_match).

    Защищённые листы (параметры, визард, сводные…) при шаблоне пропускаются —
    даже при force-удалении: «*» не должен сносить Параметры_Объединения.
    """
    s = str(pattern or "").strip()
    if s == "" or doc is None:
        return []
    if protected is None:
        protected = _lm_final_collect_delete_protected_sheet_names(doc)
    out = []
    for name in _lm_final_list_sheet_names(doc):
        if name in protected:
            continue
        if lm.text_match(name, s):
            if name not in out:
                out.append(name)
    return out


def _lm_final_resolve_named_sheet_spec(doc, spec, protected=None):
    """
    Резолвер имён для «Удаление_листов» / «Скрытие_листов».

    _Первый_ / _Последний_ — с конца/начала, пропуская защищённые листы.
    _Последний_мусорный_ — последний SheetN / __pp_mat / _to_delete_.
    Шаблоны (Сбор_книг_лог*, *лог) — см. _lm_final_expand_sheet_pattern /
    _lm_final_resolve_sheet_names_from_specs (здесь возвращается первое совпадение).
    """
    s = str(spec or "").strip()
    if s == "" or doc is None:
        return None
    if protected is None:
        protected = _lm_final_collect_delete_protected_sheet_names(doc)
    key = s.casefold()

    try:
        sheets = doc.Sheets
        count = int(sheets.getCount())
    except Exception:
        return None

    if key in _LM_FINAL_SHEET_SPEC_LAST_GARBAGE:
        i = count - 1
        while i >= 0:
            try:
                name = str(sheets.getByIndex(i).Name)
                if name not in protected and _lm_final_is_materialization_garbage_sheet_name(
                    name
                ):
                    return name
            except Exception:
                pass
            i = i - 1
        return None

    if key in _LM_FINAL_SHEET_SPEC_LAST:
        i = count - 1
        while i >= 0:
            try:
                name = str(sheets.getByIndex(i).Name)
                if name not in protected:
                    return name
            except Exception:
                pass
            i = i - 1
        return None

    if key in _LM_FINAL_SHEET_SPEC_FIRST:
        i = 0
        while i < count:
            try:
                name = str(sheets.getByIndex(i).Name)
                if name not in protected:
                    return name
            except Exception:
                pass
            i = i + 1
        return None

    if re.fullmatch(r"\d+", s):
        idx = int(s) - 1
        try:
            if idx >= 0 and idx < count:
                return str(sheets.getByIndex(idx).Name)
        except Exception:
            return None
        return None

    if _lm_final_spec_is_sheet_pattern(s):
        matched = _lm_final_expand_sheet_pattern(doc, s, protected=protected)
        if matched:
            return matched[0]
        return None

    i = 0
    while i < count:
        try:
            name = str(sheets.getByIndex(i).Name)
            if name.casefold() == key:
                return name
        except Exception:
            pass
        i = i + 1
    return None


def _lm_final_resolve_sheet_name_spec(doc, spec):
    """Имя листа, 1-based индекс, _Первый_ / _Последний_ → фактическое имя (без защиты)."""
    s = str(spec or "").strip()
    if s == "" or doc is None:
        return None
    try:
        sheets = doc.getSheets()
    except Exception:
        return None
    key = s.casefold()
    if key in _LM_FINAL_SHEET_SPEC_FIRST:
        try:
            if sheets.getCount() > 0:
                return str(sheets.getByIndex(0).Name)
        except Exception:
            return None
        return None
    if key in _LM_FINAL_SHEET_SPEC_LAST:
        try:
            cnt = int(sheets.getCount())
            if cnt > 0:
                return str(sheets.getByIndex(cnt - 1).Name)
        except Exception:
            return None
        return None
    if re.fullmatch(r"\d+", s):
        idx = int(s) - 1
        try:
            if idx >= 0 and idx < sheets.getCount():
                return str(sheets.getByIndex(idx).Name)
        except Exception:
            return None
        return None
    i = 0
    while i < sheets.getCount():
        try:
            name = str(sheets.getByIndex(i).Name)
            if name.casefold() == key:
                return name
        except Exception:
            pass
        i = i + 1
    return None


def _lm_final_sheet_specs_from_codec(fn_key, extra_args):
    """Имена листов из JSON sheets[] или legacy-текста C."""
    args = list(extra_args or ())
    if len(args) == 1:
        raw = str(args[0] or "").strip()
        if raw.startswith("[") or raw.startswith("{"):
            try:
                from libre_macros_param_codec import normalize_fn_key, param_decode

                blocks = param_decode(normalize_fn_key(fn_key), raw)
                specs = []
                bi = 0
                while bi < len(blocks):
                    b = blocks[bi]
                    bi = bi + 1
                    sheets = b.get("sheets") or []
                    si = 0
                    while si < len(sheets):
                        sn = str(sheets[si]).strip()
                        if sn != "" and sn not in specs:
                            specs.append(sn)
                        si = si + 1
                if len(specs) > 0:
                    return specs
            except Exception:
                pass
    return _lm_final_parse_sheet_specs(*args)


def _lm_final_sheet_filter_code_from_codec(fn_key, extra_args):
    """Опциональный filter/lambda из JSON блока fn_key."""
    args = list(extra_args or ())
    if len(args) != 1:
        return None
    raw = str(args[0] or "").strip()
    if not (raw.startswith("[") or raw.startswith("{")):
        return None
    try:
        from libre_macros_param_codec import normalize_fn_key, param_decode
        from libre_macros_sheet_filter_lib import parse_sheet_filter_code

        blocks = param_decode(normalize_fn_key(fn_key), raw)
        bi = 0
        while bi < len(blocks):
            code = parse_sheet_filter_code(blocks[bi])
            bi = bi + 1
            if code:
                return code
    except Exception:
        pass
    return None


def _lm_final_apply_optional_name_filter(names, filter_code):
    """Отфильтровать список имён лямбдой (с санацией). (names, err_msg)."""
    if not filter_code:
        return list(names or []), None
    try:
        from libre_macros_sheet_filter_lib import (
            apply_sheet_name_filter,
            compile_sheet_name_filter,
        )

        fn = compile_sheet_name_filter(filter_code)
        return apply_sheet_name_filter(names, fn), None
    except Exception as err:
        return None, "filter: %s" % err


def _lm_final_parse_sheet_specs(*specs):
    out = []
    for item in specs:
        for part in re.split(r"[,;]", str(item or "")):
            p = part.strip()
            if p != "" and p not in out:
                out.append(p)
    return out


def _lm_final_delete_sheets_by_names(doc, names, skip_param_sheets=True, force=False):
    deleted = []
    missing = []
    errors = []
    if force:
        skip_param_sheets = False
    i = len(names) - 1
    while i >= 0:
        raw = names[i]
        if force:
            actual = _lm_final_resolve_named_sheet_spec(doc, raw)
        else:
            actual = _lm_final_resolve_sheet_name_spec(doc, raw) or raw
        if actual is None:
            missing.append(raw)
            errors.append(
                "«%s»: не удалось сопоставить имя; в книге: %s"
                % (raw, ", ".join(_lm_final_list_sheet_names(doc)))
            )
            i = i - 1
            continue
        if skip_param_sheets and lm._lm_pp_is_param_sheet_name(actual):
            i = i - 1
            continue
        try:
            import libre_macros_collect_cfg as cw

            wiz = str(getattr(cw, "MERGE_WIZARD_LISTS_SHEET_NAME", "") or "").strip()
        except Exception:
            wiz = "_Визард_Списки"
        if skip_param_sheets and wiz and str(actual).strip().casefold() == wiz.casefold():
            i = i - 1
            continue
        removed_name, err = _lm_final_delete_one_sheet(doc, actual)
        if removed_name:
            deleted.append(removed_name)
        else:
            missing.append(raw)
            if err:
                errors.append(err)
        i = i - 1
    return deleted, missing, errors


def _lm_final_resolve_sheet_names_from_specs(doc, specs, force=False):
    """
    Раскрыть specs в список фактических имён листов.

    Поддерживает шаблоны text_match: «Сбор_книг_лог*», «*лог», «?тчет».
    Один шаблон может дать несколько имён.
    """
    names = []
    protected = _lm_final_collect_delete_protected_sheet_names(doc)
    for spec in specs:
        s = str(spec or "").strip()
        if s == "":
            continue
        if _lm_final_spec_is_sheet_pattern(s):
            expanded = _lm_final_expand_sheet_pattern(doc, s, protected=protected)
            if len(expanded) == 0:
                if s not in names:
                    names.append(s)
                continue
            for n in expanded:
                if n not in names:
                    names.append(n)
            continue
        if force:
            resolved = _lm_final_resolve_named_sheet_spec(
                doc, s, protected=protected
            )
        else:
            resolved = _lm_final_resolve_sheet_name_spec(doc, s) or s
        if resolved is not None:
            if resolved not in names:
                names.append(resolved)
        elif s not in names:
            names.append(s)
    return names


def _lm_final_hide_sheets_by_names(doc, names):
    hidden = []
    missing = []
    errors = []
    i = 0
    while i < len(names):
        name = str(names[i] or "").strip()
        if name == "":
            i = i + 1
            continue
        try:
            sheets = doc.Sheets
            if sheets.hasByName(name):
                sh = sheets.getByName(name)
                sh.setPropertyValue("IsVisible", False)
                hidden.append(name)
            else:
                missing.append(name)
        except Exception as err:
            errors.append("%s: %s" % (name, err))
        i = i + 1
    return hidden, missing, errors


def _lm_final_format_sheet_op_note(deleted_or_hidden, missing, errors, empty_msg):
    note_parts = []
    if deleted_or_hidden:
        note_parts.append("обработано: " + ", ".join(deleted_or_hidden))
    if missing:
        note_parts.append("нет: " + ", ".join(missing))
    if errors:
        note_parts.append("; ".join(errors))
    return "; ".join(note_parts) if note_parts else empty_msg


def lm_final_values_only(doc, sheets, data_ranges, header_row_ranges, *sheet_specs):
    """
    Преобразовать ячейки в значения на указанных листах (колонка C).

    C — имена листов через запятую или «;» (допускаются _Первый_, номер 1…).
    Лист со сводной (DataPilot) — inplace: снять объект, оставить значения вывода.
    Обычный лист — getDataArray/setDataArray по used area.
    """
    label = "Только_значения"
    specs = _lm_final_sheet_specs_from_codec("только_значения", sheet_specs)
    if len(specs) == 0:
        _lm_final_log(doc, label, "пропуск", "укажите имена листов в C")
        return

    ok_parts = []
    err_parts = []
    si = 0
    while si < len(specs):
        spec = specs[si]
        actual = _lm_final_resolve_named_sheet_spec(doc, spec)
        if actual is None:
            err_parts.append(
                "«%s»: лист не найден (в книге: %s)"
                % (spec, ", ".join(_lm_final_list_sheet_names(doc)))
            )
            si = si + 1
            continue
        try:
            if not doc.Sheets.hasByName(actual):
                err_parts.append("«%s»: нет в книге" % actual)
                si = si + 1
                continue
            sheet = doc.Sheets.getByName(actual)
        except Exception as err:
            err_parts.append("«%s»: %s" % (actual, err))
            si = si + 1
            continue

        if vl.lm_values_sheet_has_datapilot(sheet):
            ok, note = vl.lm_values_pivot_inplace(doc, sheet)
        else:
            ok, note = vl.lm_values_flatten_used_area(sheet)

        if ok:
            ok_parts.append("%s (%s)" % (actual, note))
        else:
            err_parts.append("%s: %s" % (actual, note))
        si = si + 1

    note_parts = []
    if len(ok_parts) > 0:
        note_parts.append("обработано: " + "; ".join(ok_parts))
    if len(err_parts) > 0:
        note_parts.append("; ".join(err_parts))
    note = "; ".join(note_parts) if len(note_parts) > 0 else "ничего не сделано"
    if len(err_parts) > 0:
        status = "ошибка"
    elif len(ok_parts) > 0:
        status = "ok"
    else:
        status = "пропуск"
    _lm_final_log(doc, label, status, note)


def lm_final_apply_pending_pivot_as_values(doc):
    """
    Неявная «Только_значения» для листов сводных с as_values (отложенная материализация).

    Вызывается в начале финальной фазы сбора, до шагов листа «Финальная_обработка».
    Для каждого листа: inplace-значения → обрезка Filter/Data → оформление.
    Регистрация для сводная_пп не выполняется.
    """
    label = "Только_значения"
    try:
        import libre_macros_pivot_lib as pl
    except ImportError:
        return

    pending = pl.lm_pp_pivot_list_as_values_pending()
    if doc is None or len(pending) == 0:
        return

    ok_parts = []
    err_parts = []
    pi = 0
    while pi < len(pending):
        name = str(pending[pi] or "").strip()
        pi = pi + 1
        if name == "":
            continue
        try:
            if not doc.Sheets.hasByName(name):
                err_parts.append("«%s»: лист не найден" % name)
                continue
            sheet = doc.Sheets.getByName(name)
        except Exception as err:
            err_parts.append("«%s»: %s" % (name, err))
            continue

        if not vl.lm_values_sheet_has_datapilot(sheet):
            err_parts.append("«%s»: нет DataPilot" % name)
            continue

        ok, note = vl.lm_values_pivot_inplace(doc, sheet)
        if not ok:
            err_parts.append("%s: %s" % (name, note))
            continue

        header_row = 0
        try:
            header_row, fmt_note = pl.lm_pp_pivot_post_values_format(doc, sheet)
            ok_parts.append("%s (%s; %s)" % (name, note, fmt_note))
        except Exception as err:
            ok_parts.append("%s (%s)" % (name, note))
            err_parts.append("%s: оформление: %s" % (name, err))

        try:
            lr_meta = pl.lm_pp_pivot_get_list_rows_meta(name)
            if lr_meta and lr_meta.get("active"):
                filled, lr_note = pl.lm_pp_pivot_apply_list_rows(
                    doc, sheet, meta=lr_meta
                )
                if filled > 0:
                    ok_parts.append("%s LIST_ROWS: %s" % (name, lr_note))
                elif lr_note and lr_note not in ("нет LIST_ROWS meta",):
                    err_parts.append("%s LIST_ROWS: %s" % (name, lr_note))
        except Exception as err:
            err_parts.append("%s LIST_ROWS: %s" % (name, err))

        # Нужно для финального стека в режиме «Текущие листы»: после снятия
        # DataPilot лист иначе выпадает из _MERGE_INVOLVED_SHEET_NAMES.
        try:
            pl.lm_pp_pivot_register_materialized_sheet(name, header_row)
        except Exception:
            pass

    pl.lm_pp_pivot_clear_as_values_pending()

    note_parts = []
    if len(ok_parts) > 0:
        note_parts.append("сводные: " + "; ".join(ok_parts))
    if len(err_parts) > 0:
        note_parts.append("; ".join(err_parts))
    note = "; ".join(note_parts) if len(note_parts) > 0 else "ничего не сделано"
    if len(err_parts) > 0 and len(ok_parts) == 0:
        status = "ошибка"
    elif len(ok_parts) > 0:
        status = "ok"
    else:
        status = "пропуск"
    _lm_final_log(doc, label + " (сводные)", status, note)


def lm_final_delete_sheets(doc, sheets, data_ranges, header_row_ranges, *sheet_specs):
    """
    Удаление листов книги (финальная фаза).

    JSON: [{"v":1,"fn":"удаление_листов","sheets":["Отчет*"],
    "filter":"lambda name: sheet_age_days(name) is not None and sheet_age_days(name) <= 3"}]
    """
    label = "Удаление_листов"
    specs = _lm_final_sheet_specs_from_codec("удаление_листов", sheet_specs)
    filter_code = _lm_final_sheet_filter_code_from_codec("удаление_листов", sheet_specs)
    if len(specs) == 0:
        names = list(_LM_FINAL_DEFAULT_DELETE_SHEETS)
        deleted, missing, errors = _lm_final_delete_sheets_by_names(doc, names)
    else:
        names = _lm_final_resolve_sheet_names_from_specs(doc, specs, force=True)
        names, ferr = _lm_final_apply_optional_name_filter(names, filter_code)
        if ferr:
            _lm_final_log(doc, label, "ошибка", ferr)
            return
        if len(names) == 0:
            _lm_final_log(doc, label, "пропуск", "пустой список листов после filter")
            return
        deleted, missing, errors = _lm_final_delete_sheets_by_names(
            doc, names, skip_param_sheets=False, force=True
        )
    note = _lm_final_format_sheet_op_note(
        deleted, missing, errors, "ничего не удалено"
    )
    if errors or missing:
        status = "ошибка"
    elif deleted:
        status = "ok"
    else:
        status = "пропуск"
    _lm_final_log(doc, label, status, note)


def lm_final_hide_sheets(doc, sheets, data_ranges, header_row_ranges, *sheet_specs):
    """
    Скрытие листов книги (финальная фаза).

    Как удаление: шаблоны + опциональный filter (лямбда с санацией).
    """
    label = "Скрытие_листов"
    specs = _lm_final_sheet_specs_from_codec("скрытие_листов", sheet_specs)
    filter_code = _lm_final_sheet_filter_code_from_codec("скрытие_листов", sheet_specs)
    if len(specs) == 0:
        names = list(_LM_FINAL_DEFAULT_DELETE_SHEETS)
    else:
        names = _lm_final_resolve_sheet_names_from_specs(doc, specs, force=True)
        names, ferr = _lm_final_apply_optional_name_filter(names, filter_code)
        if ferr:
            _lm_final_log(doc, label, "ошибка", ferr)
            return
        if len(names) == 0:
            _lm_final_log(doc, label, "пропуск", "пустой список листов после filter")
            return
    hidden, missing, errors = _lm_final_hide_sheets_by_names(doc, names)
    note = _lm_final_format_sheet_op_note(
        hidden, missing, errors, "ничего не скрыто"
    )
    _lm_final_log(doc, label, "ok" if hidden else "пропуск", note)


def lm_final_copy_sheet(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Финальная фаза: копирование вкладок по блокам JSON в C.

    [{"v":1,"fn":"копировать_лист","source_sheet":"Шаблон","dest_sheet":"Копия"}]
    """
    # В кодек-модуле «копировать_лист» является алиасом и нормализуется
    # до «копировать_переместить_лист».
    label = "копировать_лист"
    blocks = lm._lm_pp_param_decode(label, extra_args)
    if len(blocks) == 0:
        _lm_final_log(doc, label, "пропуск", "пустая C")
        return
    ok_n = 0
    err_n = 0
    notes = []
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        source = str(
            block.get("source_sheet") or block.get("source") or ""
        ).strip()
        copy_mode = bool(block.get("copy"))
        dest = str(
            block.get("dest_sheet") or block.get("dest") or block.get("target_sheet") or ""
        ).strip()
        if source == "":
            err_n = err_n + 1
            notes.append("блок %d: нет source_sheet" % (bi + 1))
        elif copy_mode:
            if dest == "":
                err_n = err_n + 1
                notes.append("блок %d: нет dest_sheet" % (bi + 1))
            else:
                ok, note, dest_actual = lm._lm_pp_copy_sheet_named(doc, source, dest)
                if ok:
                    ok_n = ok_n + 1
                    if dest_actual != "":
                        lm._lm_pp_activate_sheet_by_name(doc, dest_actual)
                else:
                    err_n = err_n + 1
                notes.append(note)
        else:
            position = str(block.get("position") or u"").strip()
            anchor_sheet = str(block.get("anchor_sheet") or u"").strip()
            ok, note = lm._lm_pp_reposition_sheet(
                doc, source, position=position, anchor_sheet_name=anchor_sheet
            )
            if ok:
                ok_n = ok_n + 1
                lm._lm_pp_activate_sheet_by_name(doc, source)
            else:
                err_n = err_n + 1
            notes.append(note)
        bi = bi + 1
    if err_n > 0 and ok_n == 0:
        status = "ошибка"
    elif ok_n > 0:
        status = "ok"
    else:
        status = "пропуск"
    _lm_final_log(doc, label, status, "; ".join(notes))


def lm_final_activate_sheet(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Финальная фаза: активировать указанный лист.

    [{"v":1,"fn":"активировать_лист","target_sheet":"Merge","tab_color":"голубой"}]
    Скрытый лист показывается; выделение — ячейка A1.
    tab_color — опционально цвет ярлычка (имя/#hex; «нет»/сброс — умолчание).
    """
    label = "активировать_лист"
    blocks = lm._lm_pp_param_decode(label, extra_args)
    if len(blocks) == 0:
        _lm_final_log(doc, label, "пропуск", "пустая C")
        return
    ok_n = 0
    err_n = 0
    notes = []
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        target = str(
            block.get("target_sheet")
            or block.get("activate_sheet")
            or block.get("sheet_name")
            or block.get("sheet")
            or ""
        ).strip()
        tab_color = str(
            block.get("tab_color")
            or block.get("color")
            or block.get("ярлык")
            or block.get("tabColor")
            or ""
        ).strip()
        if target == "":
            err_n = err_n + 1
            notes.append("блок %d: нет target_sheet" % (bi + 1))
        else:
            ok, note = lm._lm_pp_activate_sheet_focus_a1(
                doc, target, tab_color=tab_color or None
            )
            if ok:
                ok_n = ok_n + 1
                try:
                    # Фактическое имя (с учётом регистра на вкладке).
                    sh = lm._lm_pp_get_sheet_ref(doc, target)
                    actual = ""
                    if sh is not None:
                        actual = str(getattr(sh, "Name", "") or "").strip()
                    lm.lm_final_activate_sheet_remember(actual or target)
                except Exception:
                    try:
                        lm.lm_final_activate_sheet_remember(target)
                    except Exception:
                        pass
            else:
                err_n = err_n + 1
            notes.append(note)
        bi = bi + 1
    if err_n > 0 and ok_n == 0:
        status = "ошибка"
    elif ok_n > 0:
        status = "ok"
    else:
        status = "пропуск"
    _lm_final_log(doc, label, status, "; ".join(notes))


def lm_final_merge_sheets_into_one(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Финальная фаза: объединение листов в один по блокам JSON.

    [{"v":1,"fn":"объединить_листы_в_один","sheets":["A","B"],"dest_sheet":"Сводная"}]
    """
    label = "объединить_листы_в_один"
    blocks = lm._lm_pp_param_decode(label, extra_args)
    if len(blocks) == 0:
        _lm_final_log(doc, label, "пропуск", "пустая C")
        return
    ok_n = 0
    err_n = 0
    notes = []
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        lm._lm_pp_merge_sheets_dbg(
            "final",
            "блок %d/%d dest=%s sheets=%r"
            % (
                bi + 1,
                len(blocks),
                block.get("dest_sheet") or block.get("result_sheet") or "",
                block.get("sheets") or block.get("source_sheets") or [],
            ),
        )
        sheets_spec = (
            block.get("sheets")
            or block.get("source_sheets")
            or block.get("sheet_list")
            or []
        )
        dest = str(
            block.get("dest_sheet")
            or block.get("result_sheet")
            or block.get("target_sheet")
            or ""
        ).strip()
        if dest == "":
            err_n = err_n + 1
            notes.append("блок %d: нет dest_sheet" % (bi + 1))
        else:
            source_names = lm._lm_pp_resolve_merge_source_sheet_names(doc, sheets_spec)
            if not source_names:
                err_n = err_n + 1
                notes.append("блок %d: нет листов" % (bi + 1))
            else:
                hdr_param = block.get("header_row")
                first_src = lm._lm_pp_get_sheet_ref(doc, source_names[0])
                hdr0 = lm._lm_pp_parse_merge_header_row(first_src, hdr_param)
                with_formatting = False
                try:
                    with_formatting = bool(
                        lm.lm_parse_bool_param(
                            block.get("with_formatting"), default=False
                        )
                    )
                except Exception:
                    with_formatting = bool(block.get("with_formatting"))
                auto_format = True
                try:
                    auto_format = bool(
                        lm.lm_parse_bool_param(
                            block.get("auto_format"), default=True
                        )
                    )
                except Exception:
                    if "auto_format" in block:
                        auto_format = bool(block.get("auto_format"))
                columns = list(block.get("columns") or [])
                ok, note, dest_name = lm._lm_pp_merge_sheets_into_one_core(
                    doc,
                    source_names,
                    dest,
                    header_row=hdr_param,
                    with_formatting=with_formatting,
                    columns=columns,
                    auto_format=auto_format,
                )
                if ok:
                    ok_n = ok_n + 1
                    if dest_name != "":
                        dest_sh = lm._lm_pp_get_sheet_ref(doc, dest_name)
                        if dest_sh is not None and auto_format:
                            lm._lm_pp_apply_merge_dest_header_default(dest_sh, hdr0)
                        if dest_sh is not None:
                            alias = str(
                                block.get("sheet") or source_names[0] or ""
                            ).strip()
                            lm.lm_pp_copy_sheet_register_pending(
                                dest_name, hdr0, alias
                            )
                else:
                    err_n = err_n + 1
                notes.append(note)
        bi = bi + 1
    if err_n > 0 and ok_n == 0:
        status = "ошибка"
    elif ok_n > 0:
        status = "ok"
    else:
        status = "пропуск"
    _lm_final_log(doc, label, status, "; ".join(notes))


def lm_final_copy_ranges(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Финальная фаза: копирование диапазона ячеек между листами книги.

    [{"v":1,"fn":"копирование_диапазонов","source_sheet":"Сводная",
      "source_range":"A1:D20","dest_sheet":"Отчет","dest_cell":"B3"}]
    """
    _unused = (sheets, data_ranges, header_row_ranges)
    label = "копирование_диапазонов"
    blocks = lm._lm_pp_param_decode(label, extra_args)
    if len(blocks) == 0:
        _lm_final_log(doc, label, "пропуск", "пустая C")
        return
    import libre_macros_copy_ranges_lib as cr

    status, note = cr.run_copy_ranges_blocks(doc, blocks)
    _lm_final_log(doc, label, status, note)


def lm_final_create_sheet(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Финальная фаза: создать лист с заголовками и опциональными данными.

    [{"v":1,"fn":"создать_лист","name_mode":"manual","sheet_name":"Отчет",
      "columns":["A","B"],"header_row":1,"source_sheet":"Merge","source_range":"A2:B10"}]
    """
    _unused = (sheets, data_ranges, header_row_ranges)
    label = "создать_лист"
    blocks = lm._lm_pp_param_decode(label, extra_args)
    if len(blocks) == 0:
        _lm_final_log(doc, label, "пропуск", "пустая C")
        return
    import libre_macros_create_sheet_lib as cs

    status, note = cs.run_create_sheet_blocks(doc, blocks)
    _lm_final_log(doc, label, status, note)


def lm_final_send_mail(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Финальная фаза: подготовить вложение и открыть compose в почтовом клиенте.

    [{"v":1,"fn":"отправить_по_почте","attach":"whole_book","to":["boss@corp.local"],
      "subject":"Отчет"}]
    """
    _unused = (sheets, data_ranges, header_row_ranges)
    label = "отправить_по_почте"
    blocks = lm._lm_pp_param_decode(label, extra_args)
    if len(blocks) == 0:
        _lm_final_log(doc, label, "пропуск", "пустая C")
        return
    import libre_macros_mail_lib as ml

    ok_n = 0
    err_n = 0
    notes = []
    real_os = ml.detect_os()
    bi = 0
    while bi < len(blocks):
        block = blocks[bi]
        try:
            show_ui = bool(
                lm.lm_parse_bool_param(block.get("show_ui"), default=True)
            )
        except Exception:
            show_ui = block.get("show_ui", True) is not False
        if not show_ui:
            err_n = err_n + 1
            notes.append("блок %d: show_ui=false не поддерживается" % (bi + 1))
            bi = bi + 1
            continue
        client, cwarnings = ml.resolve_client(block, real_os)
        if client is None:
            err_n = err_n + 1
            notes.append(
                "блок %d: %s" % (bi + 1, "; ".join(cwarnings) if cwarnings else "client")
            )
            bi = bi + 1
            continue
        to_list, rmeta, rerr = ml.collect_to_recipients(doc, block)
        if rerr:
            err_n = err_n + 1
            notes.append("блок %d: %s" % (bi + 1, rerr))
            bi = bi + 1
            continue
        cc = ml.normalize_address_list(block.get("cc"))
        bcc = ml.normalize_address_list(block.get("bcc"))
        subject = str(block.get("subject") or "").strip()
        body = str(block.get("body") or "").strip()
        path, url, anote, aerr = ml.build_attachment(doc, block)
        if aerr:
            err_n = err_n + 1
            notes.append("блок %d: %s" % (bi + 1, aerr))
            bi = bi + 1
            continue
        ok, cnote = ml.open_compose(
            block,
            to_list,
            cc,
            bcc,
            subject,
            body,
            path,
            url,
            client,
            real_os,
        )
        warn_txt = "; ".join(cwarnings) if cwarnings else ""
        log_note = ml.format_to_log_note(
            to_list, rmeta, anote, cnote if ok else cnote, real_os, client
        )
        if warn_txt:
            log_note = warn_txt + "; " + log_note
        if ok:
            ok_n = ok_n + 1
            notes.append("блок %d: %s" % (bi + 1, log_note))
        else:
            err_n = err_n + 1
            notes.append("блок %d: %s" % (bi + 1, log_note))
        bi = bi + 1
    if err_n > 0 and ok_n == 0:
        status = "ошибка"
    elif ok_n > 0:
        status = "ok"
    else:
        status = "пропуск"
    _lm_final_log(doc, label, status, "; ".join(notes))


LM_FINAL_PUBLIC_NAMES = (
    "lm_final_delete_sheets",
    "lm_final_hide_sheets",
    "lm_final_values_only",
    "lm_final_copy_sheet",
    "lm_final_activate_sheet",
    "lm_final_merge_sheets_into_one",
    "lm_final_copy_ranges",
    "lm_final_create_sheet",
    "lm_final_send_mail",
)
