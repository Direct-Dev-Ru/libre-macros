# -*- coding: utf-8 -*-
"""Условие выполнения шага постобработки / финальной обработки (колонка D или E)."""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.723"
try:
    unicode
except NameError:
    unicode = str

import json
import re

import libre_macros_step_condition_cfg as _cfg


class StepConditionError(ValueError):
    """Невалидное условие или ошибка оценки (fail step)."""


def _u(v):
    if v is None:
        return u""
    try:
        return unicode(v)
    except Exception:
        return u"%s" % v


def is_plugin_fn_key(fn_name):
    key = _u(fn_name).strip().casefold()
    if key == u"":
        return False
    if key in (u"функция_плагин", u"plugin", u"plugin_function", u"функция-плагин"):
        return True
    try:
        import libre_macros_collect_cfg as _cw_cfg

        pk = _u(getattr(_cw_cfg, u"MERGE_PLUGIN_FUNCTION_KEY", u"функция_плагин")).strip().casefold()
        return key == pk
    except Exception:
        return False


def condition_col_for_fn(fn_name):
    """0-based индекс колонки условия: D=3 для codec, E=4 для плагина."""
    if is_plugin_fn_key(fn_name):
        return int(_cfg.STEP_CONDITION_COL_PLUGIN)
    return int(_cfg.STEP_CONDITION_COL_CODEC)


def pick_condition_raw(fn_name, d_raw=u"", e_raw=u""):
    """Текст условия из D (codec) или E (plugin)."""
    if is_plugin_fn_key(fn_name):
        return _u(e_raw).strip()
    return _u(d_raw).strip()


def _parse_users_list(raw):
    if isinstance(raw, (list, tuple)):
        out = []
        for item in raw:
            s = _u(item).strip()
            if s != u"":
                out.append(s)
        return out
    text = _u(raw).strip()
    if text == u"":
        return []
    parts = re.split(r"[,;]+", text)
    out = []
    for p in parts:
        s = _u(p).strip().strip(u'"').strip(u"'")
        if s != u"":
            out.append(s)
    return out


def _parse_value_list(raw):
    return _parse_users_list(raw)


def normalize_step_condition(obj):
    """Нормализовать dict условия. Пустое / None → None."""
    if obj is None:
        return None
    if isinstance(obj, (list, tuple)):
        if len(obj) == 0:
            return None
        obj = obj[0]
    if not isinstance(obj, dict):
        raise StepConditionError(u"условие: ожидается JSON-объект")
    kind = _u(obj.get(u"kind") or u"").strip().casefold()
    if kind == u"":
        raise StepConditionError(u"условие: не задан kind")
    aliases = {
        u"лямбда": _cfg.KIND_LAMBDA,
        u"переменная": _cfg.KIND_VARIABLE,
        u"var": _cfg.KIND_VARIABLE,
        u"runtime": _cfg.KIND_VARIABLE,
        u"среда": _cfg.KIND_ENV,
        u"env_var": _cfg.KIND_ENV,
        u"пользователь": _cfg.KIND_OS_USER,
        u"os_username": _cfg.KIND_OS_USER,
        u"user": _cfg.KIND_OS_USER,
    }
    kind = aliases.get(kind, kind)
    if kind not in (
        _cfg.KIND_LAMBDA,
        _cfg.KIND_VARIABLE,
        _cfg.KIND_ENV,
        _cfg.KIND_OS_USER,
    ):
        raise StepConditionError(u"условие: неизвестный kind «%s»" % kind)
    out = {
        u"v": int(_cfg.STEP_CONDITION_VERSION),
        u"fn": _cfg.STEP_CONDITION_FN,
        u"kind": kind,
    }
    if kind == _cfg.KIND_LAMBDA:
        expr = _u(obj.get(u"expr") or obj.get(u"lambda") or u"").strip()
        if expr == u"":
            raise StepConditionError(u"условие lambda: пустой expr")
        out[u"expr"] = expr
        return out
    if kind == _cfg.KIND_VARIABLE:
        name = _u(obj.get(u"name") or u"").strip()
        if name == u"":
            raise StepConditionError(u"условие variable: пустой name")
        if u"~" in name:
            raise StepConditionError(u"условие variable: «~» в имени запрещена")
        out[u"name"] = name
        out[u"file"] = _u(obj.get(u"file") or u"").strip()
        out[u"sheet"] = _u(obj.get(u"sheet") or u"").strip()
        op = _u(obj.get(u"op") or u"eq").strip().casefold() or u"eq"
        if op not in _cfg.COMPARE_OP_LABEL:
            raise StepConditionError(u"условие variable: неизвестный op «%s»" % op)
        out[u"op"] = op
        if op in (u"in", u"not_in"):
            out[u"value"] = _parse_value_list(obj.get(u"value"))
        else:
            out[u"value"] = _u(obj.get(u"value") if u"value" in obj else u"")
        cas = _u(obj.get(u"compare_as") or u"text").strip().casefold() or u"text"
        if cas not in (u"text", u"number", u"auto"):
            cas = u"text"
        out[u"compare_as"] = cas
        out[u"case_sensitive"] = bool(obj.get(u"case_sensitive", False))
        missing = _u(obj.get(u"missing") or u"false").strip().casefold() or u"false"
        if missing not in (u"false", u"true", u"fail"):
            missing = u"false"
        out[u"missing"] = missing
        return out
    if kind == _cfg.KIND_ENV:
        env_name = _u(obj.get(u"env_name") or obj.get(u"name") or u"").strip()
        if env_name == u"":
            raise StepConditionError(u"условие env: пустой env_name")
        out[u"env_name"] = env_name
        op = _u(obj.get(u"op") or u"eq").strip().casefold() or u"eq"
        if op not in _cfg.COMPARE_OP_LABEL:
            raise StepConditionError(u"условие env: неизвестный op «%s»" % op)
        out[u"op"] = op
        rhs_kind = _u(obj.get(u"rhs_kind") or u"literal").strip().casefold() or u"literal"
        if rhs_kind not in (u"literal", u"global", u"variable"):
            rhs_kind = u"literal"
        out[u"rhs_kind"] = rhs_kind
        if rhs_kind == u"literal":
            if op in (u"in", u"not_in"):
                out[u"value"] = _parse_value_list(obj.get(u"value"))
            else:
                out[u"value"] = _u(obj.get(u"value") if u"value" in obj else u"")
            out[u"global_name"] = u""
        else:
            gname = _u(obj.get(u"global_name") or obj.get(u"rhs_name") or u"").strip()
            if gname == u"":
                raise StepConditionError(u"условие env: пустой global_name")
            out[u"global_name"] = gname
            out[u"value"] = u""
        cas = _u(obj.get(u"compare_as") or u"text").strip().casefold() or u"text"
        if cas not in (u"text", u"number", u"auto"):
            cas = u"text"
        out[u"compare_as"] = cas
        out[u"case_sensitive"] = bool(obj.get(u"case_sensitive", False))
        missing = _u(obj.get(u"missing") or u"false").strip().casefold() or u"false"
        if missing not in (u"false", u"true", u"fail"):
            missing = u"false"
        out[u"missing"] = missing
        return out
    # os_user
    op = _u(obj.get(u"op") or u"in").strip().casefold() or u"in"
    if op not in (u"in", u"not_in"):
        raise StepConditionError(u"условие os_user: op только in / not_in")
    users = _parse_users_list(obj.get(u"users") if u"users" in obj else obj.get(u"value"))
    if len(users) == 0:
        raise StepConditionError(u"условие os_user: пустой список users")
    out[u"op"] = op
    out[u"users"] = users
    out[u"case_sensitive"] = bool(obj.get(u"case_sensitive", False))
    return out


def decode_step_condition(raw_text):
    """
    Разобрать текст ячейки условия.
    Пусто → None.
    Невалидный JSON / kind → StepConditionError.
    """
    text = _u(raw_text).strip()
    if text == u"":
        return None
    try:
        data = json.loads(text)
    except Exception as err:
        raise StepConditionError(u"условие: невалидный JSON (%s)" % err)
    return normalize_step_condition(data)


def encode_step_condition(obj):
    """Сериализация dict условия в компактный JSON; None/пусто → «»."""
    if obj is None:
        return u""
    norm = normalize_step_condition(obj)
    if norm is None:
        return u""
    return json.dumps(norm, ensure_ascii=False, separators=(u",", u":"))


def _text_match_mask(value, mask):
    m = _u(mask).strip()
    if m == u"" or m == u"*" or m.casefold() in (u"все", u"all"):
        return True
    v = _u(value)
    if u"*" in m or u"?" in m:
        try:
            import fnmatch

            return fnmatch.fnmatchcase(v.casefold(), m.casefold())
        except Exception:
            pass
    return v.casefold() == m.casefold()


def lookup_runtime_variable(name, variables_map=None, variables_order=None, file_mask=u"", sheet_mask=u""):
    """
    Найти значение в runtime-карте по логическому имени (+ опц. маски file/sheet).
    Возвращает (found, text).
    """
    want = _u(name).strip()
    if want == u"":
        return False, u""
    try:
        import libre_macros_collect_cfg as _cw_cfg

        vm = variables_map
        if vm is None:
            vm = getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_MAP", None) or {}
        order = variables_order
        if order is None:
            order = getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_ORDER", None) or []
    except Exception:
        vm = variables_map if isinstance(variables_map, dict) else {}
        order = variables_order or []
    if not isinstance(vm, dict):
        vm = {}
    matches = []
    for key in vm:
        parts = _u(key).split(u"~")
        if len(parts) < 1:
            continue
        if parts[0].casefold() != want.casefold():
            continue
        src_sheet = parts[1] if len(parts) > 1 else u""
        src_file = parts[2] if len(parts) > 2 else u""
        if not _text_match_mask(src_sheet, sheet_mask):
            continue
        if not _text_match_mask(src_file, file_mask):
            continue
        matches.append(key)
    if not matches:
        return False, u""
    order_pos = {}
    oi = 0
    while oi < len(order or []):
        order_pos[_u(order[oi])] = oi
        oi += 1
    best = matches[0]
    bi = order_pos.get(_u(best), -1)
    mi = 1
    while mi < len(matches):
        k = matches[mi]
        ki = order_pos.get(_u(k), -1)
        if ki >= bi:
            best = k
            bi = ki
        mi += 1
    raw = vm.get(best)
    if isinstance(raw, (list, tuple)):
        text = _u(raw[0] if len(raw) > 0 else u"")
    else:
        text = _u(raw)
    return True, text


def _as_number(text):
    s = _u(text).strip().replace(u" ", u"").replace(u",", u".")
    if s == u"":
        return None
    try:
        return float(s)
    except Exception:
        return None


def _compare_values(left, right, op, compare_as=u"text", case_sensitive=False):
    op = _u(op).strip().casefold()
    cas = _u(compare_as).strip().casefold() or u"text"
    lt = _u(left)
    if op in (u"in", u"not_in"):
        items = right if isinstance(right, (list, tuple)) else _parse_value_list(right)
        if not case_sensitive:
            lt_cf = lt.strip().casefold()
            hit = False
            for it in items:
                if _u(it).strip().casefold() == lt_cf:
                    hit = True
                    break
        else:
            hit = lt.strip() in [_u(it).strip() for it in items]
        return (not hit) if op == u"not_in" else hit
    rt = _u(right)
    use_num = cas == u"number"
    if cas == u"auto":
        ln = _as_number(lt)
        rn = _as_number(rt)
        use_num = ln is not None and rn is not None
    if use_num and op in (u"eq", u"ne", u"gt", u"gte", u"lt", u"lte"):
        ln = _as_number(lt)
        rn = _as_number(rt)
        if ln is None or rn is None:
            use_num = False
        else:
            if op == u"eq":
                return ln == rn
            if op == u"ne":
                return ln != rn
            if op == u"gt":
                return ln > rn
            if op == u"gte":
                return ln >= rn
            if op == u"lt":
                return ln < rn
            if op == u"lte":
                return ln <= rn
    a = lt if case_sensitive else lt.casefold()
    b = rt if case_sensitive else rt.casefold()
    if op == u"eq":
        return a.strip() == b.strip()
    if op == u"ne":
        return a.strip() != b.strip()
    if op == u"contains":
        return b in a
    if op == u"not_contains":
        return b not in a
    if op == u"starts_with":
        return a.startswith(b)
    if op == u"ends_with":
        return a.endswith(b)
    if op == u"gt":
        return a > b
    if op == u"gte":
        return a >= b
    if op == u"lt":
        return a < b
    if op == u"lte":
        return a <= b
    raise StepConditionError(u"неизвестный op «%s»" % op)


def _missing_result(policy, kind_label):
    p = _u(policy).strip().casefold() or u"false"
    if p == u"true":
        return True, u"%s missing→True" % kind_label
    if p == u"fail":
        raise StepConditionError(u"%s: значение не найдено (missing=fail)" % kind_label)
    return False, u"%s missing→False" % kind_label


def _eval_lambda(expr):
    from libre_macros_lambda_column_lib import expand_source_variables_in_lambda_code
    from libre_macros_sanitize_lib import (
        LM_SANITIZE_POLICY_STANDARD,
        lm_sanitize_code_or_raise,
        lm_sanitize_filtered_builtins,
    )
    from libre_macros_source_roots_lib import env_text, lookup_global_text, os_username

    code = expand_source_variables_in_lambda_code(
        _u(expr).strip(), log_fn_key=u"условие_выполнения"
    )
    if code == u"":
        raise StepConditionError(u"условие lambda: пустой expr после expand")
    stripped = code.lstrip()
    if not (stripped.startswith(u"lambda") or stripped.startswith(u"def ")):
        code = u"lambda: " + code
    lm_sanitize_code_or_raise(code, policy=LM_SANITIZE_POLICY_STANDARD)

    def tr(v):
        if v is None:
            return u""
        return _u(v)

    def var(name, file_mask=u"", sheet_mask=u""):
        found, text = lookup_runtime_variable(name, file_mask=file_mask, sheet_mask=sheet_mask)
        return text if found else u""

    def env(name):
        return env_text(name)

    def os_user():
        return os_username()

    def glob_var(name):
        return lookup_global_text(name)

    glo = {
        u"__builtins__": lm_sanitize_filtered_builtins(),
        u"tr": tr,
        u"str": str,
        u"unicode": unicode,
        u"int": int,
        u"float": float,
        u"bool": bool,
        u"len": len,
        u"abs": abs,
        u"min": min,
        u"max": max,
        u"round": round,
        u"var": var,
        u"env": env,
        u"os_user": os_user,
        u"os_username": os_user,
        u"glob_var": glob_var,
        u"lookup_global": glob_var,
    }
    try:
        if code.lstrip().startswith(u"lambda"):
            fn = eval(code, glo, {})
        else:
            loc = {}
            exec(code, glo, loc)
            fn = None
            for v in loc.values():
                if callable(v):
                    fn = v
                    break
    except Exception as err:
        raise StepConditionError(u"условие lambda: компиляция: %s" % err)
    if fn is None or not callable(fn):
        raise StepConditionError(u"условие lambda: нужен callable")
    try:
        result = fn()
    except TypeError:
        # lambda rows: … — без rows не вызываем
        raise StepConditionError(
            u"условие lambda: выражение требует аргументы (rows не поддерживается в gate шага)"
        )
    except Exception as err:
        raise StepConditionError(u"условие lambda: выполнение: %s" % err)
    return bool(result), u"lambda"


def _eval_variable(cond):
    found, text = lookup_runtime_variable(
        cond.get(u"name"),
        file_mask=cond.get(u"file") or u"",
        sheet_mask=cond.get(u"sheet") or u"",
    )
    if not found:
        return _missing_result(cond.get(u"missing"), u"variable")
    ok = _compare_values(
        text,
        cond.get(u"value"),
        cond.get(u"op") or u"eq",
        compare_as=cond.get(u"compare_as") or u"text",
        case_sensitive=bool(cond.get(u"case_sensitive")),
    )
    return ok, u"variable %s %s" % (cond.get(u"name"), cond.get(u"op"))


def _eval_env(cond):
    from libre_macros_source_roots_lib import env_text, lookup_global_text

    env_name = cond.get(u"env_name") or u""
    left = env_text(env_name)
    if left == u"" and env_name:
        # отличить отсутствующую env от пустой строки: os.environ.get
        import os

        try:
            missing_env = os.environ.get(_u(env_name)) is None
        except Exception:
            missing_env = True
        if missing_env:
            return _missing_result(cond.get(u"missing"), u"env")
    rhs_kind = _u(cond.get(u"rhs_kind") or u"literal").casefold()
    if rhs_kind == u"literal":
        right = cond.get(u"value")
    else:
        gname = cond.get(u"global_name") or u""
        if rhs_kind == u"global":
            right = lookup_global_text(gname)
            if right == u"":
                found, right = lookup_runtime_variable(gname)
                if not found:
                    return _missing_result(cond.get(u"missing"), u"env.rhs")
        else:
            found, right = lookup_runtime_variable(gname)
            if not found:
                right = lookup_global_text(gname)
                if right == u"":
                    return _missing_result(cond.get(u"missing"), u"env.rhs")
    ok = _compare_values(
        left,
        right,
        cond.get(u"op") or u"eq",
        compare_as=cond.get(u"compare_as") or u"text",
        case_sensitive=bool(cond.get(u"case_sensitive")),
    )
    return ok, u"env %s %s" % (env_name, cond.get(u"op"))


def _eval_os_user(cond):
    from libre_macros_source_roots_lib import os_username

    user = os_username()
    ok = _compare_values(
        user,
        cond.get(u"users") or [],
        cond.get(u"op") or u"in",
        compare_as=u"text",
        case_sensitive=bool(cond.get(u"case_sensitive")),
    )
    return ok, u"os_user %s [%s]" % (cond.get(u"op"), u",".join(cond.get(u"users") or []))


def eval_step_condition(raw_or_obj, variables_map=None, variables_order=None):
    """
    Оценить условие.

    Возвращает (ok: bool, summary: unicode).
    Пустое условие → (True, «нет условия»).
    Невалидное / ошибка оценки → StepConditionError.
    """
    if raw_or_obj is None or (isinstance(raw_or_obj, (str, unicode)) and _u(raw_or_obj).strip() == u""):
        return True, u"нет условия"
    if isinstance(raw_or_obj, dict):
        cond = normalize_step_condition(raw_or_obj)
    else:
        cond = decode_step_condition(raw_or_obj)
    if cond is None:
        return True, u"нет условия"
    kind = cond.get(u"kind")
    # variables_map override for tests
    if variables_map is not None:
        try:
            import libre_macros_collect_cfg as _cw_cfg

            prev_map = getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_MAP", None)
            prev_order = getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_ORDER", None)
            _cw_cfg._MERGE_SOURCE_VARIABLES_MAP = variables_map
            if variables_order is not None:
                _cw_cfg._MERGE_SOURCE_VARIABLES_ORDER = variables_order
        except Exception:
            prev_map = None
            prev_order = None
    else:
        prev_map = None
        prev_order = None
    try:
        if kind == _cfg.KIND_LAMBDA:
            return _eval_lambda(cond.get(u"expr"))
        if kind == _cfg.KIND_VARIABLE:
            return _eval_variable(cond)
        if kind == _cfg.KIND_ENV:
            return _eval_env(cond)
        if kind == _cfg.KIND_OS_USER:
            return _eval_os_user(cond)
        raise StepConditionError(u"условие: неизвестный kind «%s»" % kind)
    finally:
        if prev_map is not None:
            try:
                import libre_macros_collect_cfg as _cw_cfg

                _cw_cfg._MERGE_SOURCE_VARIABLES_MAP = prev_map
                if prev_order is not None:
                    _cw_cfg._MERGE_SOURCE_VARIABLES_ORDER = prev_order
            except Exception:
                pass


def condition_summary(cond_or_raw):
    """Краткая строка для журнала / UI."""
    try:
        if isinstance(cond_or_raw, dict):
            cond = normalize_step_condition(cond_or_raw)
        else:
            cond = decode_step_condition(cond_or_raw)
    except Exception as err:
        return u"ошибка: %s" % err
    if cond is None:
        return u""
    kind = cond.get(u"kind")
    if kind == _cfg.KIND_LAMBDA:
        return u"lambda"
    if kind == _cfg.KIND_VARIABLE:
        return u"variable:%s %s" % (cond.get(u"name"), cond.get(u"op"))
    if kind == _cfg.KIND_ENV:
        return u"env:%s %s" % (cond.get(u"env_name"), cond.get(u"op"))
    if kind == _cfg.KIND_OS_USER:
        return u"os_user %s [%s]" % (cond.get(u"op"), u",".join(cond.get(u"users") or [])[:5])
    return _u(kind)
