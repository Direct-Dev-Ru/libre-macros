# -*- coding: utf-8 -*-
"""
Sugar лямбд `rows("Col")` / `rows("Col")[-1]` → `_lm_rows(_ri, col, offset)`.
Sugar `upd_rows("Col")` / `upd_rows("Col")[-1]` → `_lm_upd_rows(_ri, col, offset)`.

`rows` — исходная/текущая матрица (смещение по строкам листа).
`upd_rows` — уже обработанные строки **выше текущей** (накопление при
последовательном проходе сверху вниз; значения после применения ops).

Перед sugar/санацией/compile: `<<Переменные.…>>` из runtime-карты
(`expand_source_variables_in_lambda_code`).

Общий helper для `замена_значений`, `текстовые_операции`, `столбец_по_лямбде`.
"""
from __future__ import print_function
MACRO_VERSION = "3.10.717"
import ast
import datetime
import re

try:
    unicode  # type: ignore[name-defined]
except NameError:  # pragma: no cover
    unicode = str


def _ast_const_int(node):
    """Целая константа из AST (Py2 Num / Py3 Constant, унарный +/-)."""
    if node is None:
        return None
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
        v = _ast_const_int(node.operand)
        if v is None:
            return None
        return -int(v)
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.UAdd):
        return _ast_const_int(node.operand)
    if hasattr(ast, "Constant") and isinstance(node, ast.Constant):
        val = node.value
        if isinstance(val, bool) or not isinstance(val, int):
            return None
        return int(val)
    if hasattr(ast, "Num") and isinstance(node, ast.Num):
        try:
            n = node.n
        except Exception:
            return None
        if isinstance(n, bool) or not isinstance(n, int):
            return None
        return int(n)
    return None


def _subscript_offset(node):
    """rows(...)[N] → N или None, если индекс не целочисленный литерал."""
    sl = getattr(node, "slice", None)
    if sl is None:
        return None
    if hasattr(ast, "Index") and isinstance(sl, ast.Index):
        sl = sl.value
    if isinstance(sl, ast.Slice) or (
        hasattr(ast, "ExtSlice") and isinstance(sl, ast.ExtSlice)
    ):
        return None
    return _ast_const_int(sl)


def _is_sugar_call(node, sugar_name):
    if not isinstance(node, ast.Call):
        return False
    func = node.func
    if not isinstance(func, ast.Name):
        return False
    if unicode(func.id) != unicode(sugar_name):
        return False
    args = list(getattr(node, "args", None) or [])
    return len(args) >= 1


def _is_rows_call(node):
    return _is_sugar_call(node, u"rows")


def _is_upd_rows_call(node):
    return _is_sugar_call(node, u"upd_rows")


def _make_name_load(ident):
    return ast.Name(id=str(ident), ctx=ast.Load())


def _make_int_node(value):
    value = int(value)
    if hasattr(ast, "Constant"):
        return ast.Constant(value)
    return ast.Num(n=value)


def _copy_call_extras(src, dst_kwargs):
    fields = getattr(ast.Call, "_fields", ())
    if "starargs" in fields:
        dst_kwargs["starargs"] = getattr(src, "starargs", None)
    if "kwargs" in fields:
        dst_kwargs["kwargs"] = getattr(src, "kwargs", None)


def _make_lm_accessor_call(rows_call, target_name, offset):
    arg0 = rows_call.args[0]
    kw = dict(
        func=_make_name_load(target_name),
        args=[
            _make_name_load("_ri"),
            arg0,
            _make_int_node(offset),
        ],
        keywords=list(getattr(rows_call, "keywords", None) or []),
    )
    _copy_call_extras(rows_call, kw)
    return ast.Call(**kw)


def _make_lm_rows_call(rows_call, offset):
    return _make_lm_accessor_call(rows_call, u"_lm_rows", offset)


def _make_lm_upd_rows_call(rows_call, offset):
    return _make_lm_accessor_call(rows_call, u"_lm_upd_rows", offset)


class _LambdaSugarRewriter(ast.NodeTransformer):
    """rows/upd_rows(arg)[n] / rows/upd_rows(arg) → _lm_*(_ri, arg, n|0)."""

    def _rewrite_subscript(self, node, is_call, make_call):
        inner = node.value
        if is_call(inner):
            offset = _subscript_offset(node)
            if offset is not None:
                visited_args = [self.visit(a) for a in inner.args]
                kw = dict(
                    func=inner.func,
                    args=visited_args,
                    keywords=[
                        self.visit(k) for k in (getattr(inner, "keywords", None) or [])
                    ],
                )
                _copy_call_extras(inner, kw)
                inner_vis = ast.Call(**kw)
                return make_call(inner_vis, offset)
        return None

    def visit_Subscript(self, node):
        for is_call, make_call in (
            (_is_upd_rows_call, _make_lm_upd_rows_call),
            (_is_rows_call, _make_lm_rows_call),
        ):
            repl = self._rewrite_subscript(node, is_call, make_call)
            if repl is not None:
                return repl
        return self.generic_visit(node)

    def visit_Call(self, node):
        node = self.generic_visit(node)
        if _is_upd_rows_call(node):
            return _make_lm_upd_rows_call(node, 0)
        if _is_rows_call(node):
            return _make_lm_rows_call(node, 0)
        return node


_ROWSSugarRewriter = _LambdaSugarRewriter

_UPD_ROWS_INDEX_RE = re.compile(
    r"\bupd_rows\s*\(([^)]*)\)\s*\[\s*([+-]?\d+)\s*\]"
)
_UPD_ROWS_BARE_RE = re.compile(r"\bupd_rows\s*\(([^)]*)\)")
_ROWS_INDEX_RE = re.compile(
    r"\brows\s*\(([^)]*)\)\s*\[\s*([+-]?\d+)\s*\]"
)
_ROWS_BARE_RE = re.compile(r"\brows\s*\(([^)]*)\)")


def _rewrite_lambda_sugar_regex(code):
    text = unicode(code or u"")

    def _upd_idx(m):
        return u"_lm_upd_rows(_ri, %s, %s)" % (m.group(1).strip(), m.group(2))

    text = _UPD_ROWS_INDEX_RE.sub(_upd_idx, text)

    def _upd_bare(m):
        return u"_lm_upd_rows(_ri, %s, 0)" % m.group(1).strip()

    text = _UPD_ROWS_BARE_RE.sub(_upd_bare, text)

    def _idx(m):
        return u"_lm_rows(_ri, %s, %s)" % (m.group(1).strip(), m.group(2))

    text = _ROWS_INDEX_RE.sub(_idx, text)

    def _bare(m):
        return u"_lm_rows(_ri, %s, 0)" % m.group(1).strip()

    return _ROWS_BARE_RE.sub(_bare, text)


def _rewrite_rows_sugar_regex(code):
    return _rewrite_lambda_sugar_regex(code)


def _ast_to_source(tree, fallback_code):
    if tree is None:
        return unicode(fallback_code or u"")
    try:
        ast.fix_missing_locations(tree)
    except Exception:
        pass
    if hasattr(ast, "unparse"):
        try:
            return unicode(ast.unparse(tree)).strip()
        except Exception:
            pass
    return _rewrite_lambda_sugar_regex(fallback_code)


def lambda_uses_rows_sugar(code):
    """True, если в коде есть sugar rows(...) или upd_rows(...)."""
    text = unicode(code or u"")
    if text.strip() == u"":
        return False
    if u"upd_rows(" in text or u"rows(" in text:
        return True
    try:
        rewritten = rewrite_lambda_column_sugar(text, kind=u"rows")
        return u"_lm_rows" in rewritten or u"_lm_upd_rows" in rewritten
    except Exception:
        return False


_SOURCE_VAR_PLACEHOLDER_RE = re.compile(r"<<\s*Переменные\.([^>]+?)\s*>>", re.IGNORECASE)


def expand_source_variables_in_lambda_code( code_text, log_fn_key=u"лямбда", sheet_name_hint=None):
    """
    Подставить <<Переменные.…>> из runtime-карты до sugar / санации / compile.

    Значение вставляется как текст (for_formula=False): в лямбде обычно
    уже в кавычках, напр. rows(\"X\") == \"<<Переменные.Owner>>\".
    Без плейсхолдеров / без карты — исходный код без изменений.
    """
    raw = unicode(code_text or u"")
    if raw == u"" or _SOURCE_VAR_PLACEHOLDER_RE.search(raw) is None:
        return raw
    try:
        from libre_macros_lib import (
            _lm_pp_expand_source_variables_template,
            lm_pp_active_context,
            lm_pp_set_active_context,
        )
        import libre_macros_collect_cfg as _cw_cfg
    except Exception:
        return raw
    prev_ctx = lm_pp_active_context()
    ctx = dict(prev_ctx or {})
    # Runtime-карта сбора — источник истины (seed + ручной ввод / put).
    ctx[u"source_variables_map"] = (
        getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_MAP", {}) or {}
    )
    ctx[u"source_variables_order"] = (
        getattr(_cw_cfg, u"_MERGE_SOURCE_VARIABLES_ORDER", []) or []
    )
    sheet = ctx.get(u"sheet")
    doc = ctx.get(u"doc")
    try:
        header_row = int(ctx.get(u"header_row") or 0)
    except Exception:
        header_row = 0
    try:
        data_start = int(ctx.get(u"data_start") or (header_row + 1))
    except Exception:
        data_start = header_row + 1
    lm_pp_set_active_context(ctx)
    try:
        return unicode(
            _lm_pp_expand_source_variables_template(
                raw,
                doc=doc,
                sheet=sheet,
                row_0=data_start,
                header_row=header_row,
                log_fn_key=log_fn_key,
                for_formula=False,
                sheet_name_hint=sheet_name_hint,
            )
            or u""
        )
    except Exception:
        return raw
    finally:
        lm_pp_set_active_context(prev_ctx)


def rewrite_lambda_column_sugar(code, kind=u"rows"):
    """
    Трансформация sugar → обычный Python.

    kind:
      rows  — rows("Col") / upd_rows("Col") и смещения [...]
      name / sheets — без rows-sugar (вернуть как есть)
    """
    text = unicode(code or u"").strip()
    if text == u"":
        return text
    kind = unicode(kind or u"rows").strip().casefold()
    if kind != u"rows":
        return text
    stripped = text.lstrip()
    mode = "eval" if stripped.startswith(u"lambda") else "exec"
    try:
        tree = ast.parse(text, mode=mode)
        new_tree = _LambdaSugarRewriter().visit(tree)
        return _ast_to_source(new_tree, text)
    except Exception:
        return _rewrite_lambda_sugar_regex(text)


def rows_lambda_eval_namespace(extra=None):
    """Namespace eval: _lm_rows/_ri + re/datetime/len/str/int/… (без UNO)."""

    def tr(v):
        if v is None:
            return u""
        return unicode(v)

    glo = {
        u"__builtins__": None,
        u"_lm_rows": None,
        u"_ri": 0,
        u"row_id": 0,
        u"datetime": datetime,
        u"date": datetime.date,
        u"timedelta": datetime.timedelta,
        u"re": re,
        u"len": len,
        u"str": str,
        u"tr": tr,
        u"int": int,
        u"float": float,
        u"round": round,
        u"min": min,
        u"max": max,
        u"abs": abs,
        u"bool": bool,
        u"unicode": unicode,
    }
    if extra:
        glo.update(extra)
    return glo


def compile_rows_lambda(code_text, extra_namespace=None, policy=None):
    """
    Санация + компиляция callable.

    Пустой код → None.
    Успех → (fn, glo). glo нужно обновлять: _ri, _lm_rows, idx, n, text.
    """
    raw = expand_source_variables_in_lambda_code(
        unicode(code_text or u"").strip(), log_fn_key=u"лямбда"
    )
    if raw == u"":
        return None
    rewritten = rewrite_lambda_column_sugar(raw, kind=u"rows")
    from libre_macros_sanitize_lib import (
        LM_SANITIZE_POLICY_STANDARD,
        lm_sanitize_code,
        lm_sanitize_code_or_raise,
        lm_sanitize_filtered_builtins,
        lm_sanitize_format_issues,
    )

    if policy is None:
        policy = LM_SANITIZE_POLICY_STANDARD
    try:
        lm_sanitize_code_or_raise(rewritten, policy=policy)
    except Exception as err:
        try:
            res = lm_sanitize_code(rewritten, policy=policy)
            detail = lm_sanitize_format_issues(res.issues)
            raise ValueError(u"санация лямбды rows: %s" % detail)
        except ValueError:
            raise
        except Exception:
            raise ValueError(u"санация лямбды rows: %s" % err)

    builtins_map = lm_sanitize_filtered_builtins()
    glo = rows_lambda_eval_namespace(extra_namespace)
    glo[u"__builtins__"] = builtins_map
    try:
        if rewritten.lstrip().startswith(u"lambda"):
            fn = eval(rewritten, glo, {})
        else:
            loc = {}
            exec(rewritten, glo, loc)
            fn = None
            if u"fn" in loc and callable(loc.get(u"fn")):
                fn = loc[u"fn"]
            else:
                for v in loc.values():
                    if callable(v):
                        fn = v
                        break
    except Exception as err:
        raise ValueError(u"компиляция лямбды rows: %s" % err)
    if fn is None or not callable(fn):
        raise ValueError(u"нужен lambda rows: … или def …")
    return fn, glo


compile_lambda_column_expr = compile_rows_lambda


def _lambda_column_eval_namespace(extra=None):
    """Namespace eval: безопасные builtins + datetime/re/лен/строки."""
    glo = {
        u"__builtins__": None,
        u"datetime": datetime,
        u"date": datetime.date,
        u"timedelta": datetime.timedelta,
        u"re": re,
        u"len": len,
        u"str": str,
        u"int": int,
        u"float": float,
        u"round": round,
        u"min": min,
        u"max": max,
        u"abs": abs,
        u"bool": bool,
        u"unicode": unicode,
    }
    if extra:
        glo.update(extra)
    return glo


def compile_lambda_column_name(code_text, extra_namespace=None, policy=None):
    """
    compile: callable(sheet:str, headers:list[str]) -> str

    Пустой код → None. Ошибка санации/синтаксиса → ValueError.
    """
    raw = expand_source_variables_in_lambda_code(
        unicode(code_text or u"").strip(), log_fn_key=u"лямбда_имя"
    )
    if raw == u"":
        return None

    rewritten = rewrite_lambda_column_sugar(raw, kind=u"name")
    from libre_macros_sanitize_lib import (
        LM_SANITIZE_POLICY_STANDARD,
        lm_sanitize_code_or_raise,
        lm_sanitize_filtered_builtins,
        lm_sanitize_format_issues,
        lm_sanitize_code,
    )

    if policy is None:
        policy = LM_SANITIZE_POLICY_STANDARD
    try:
        lm_sanitize_code_or_raise(rewritten, policy=policy)
    except Exception as err:
        # Стараться показать детали санитайзера.
        try:
            res = lm_sanitize_code(rewritten, policy=policy)
            detail = lm_sanitize_format_issues(res.issues)
            raise ValueError(u"санация лямбды column name: %s" % detail)
        except ValueError:
            raise
        except Exception:
            raise ValueError(u"санация лямбды column name: %s" % err)

    builtins_map = lm_sanitize_filtered_builtins()
    glo = _lambda_column_eval_namespace(extra_namespace)
    glo[u"__builtins__"] = builtins_map

    try:
        if rewritten.lstrip().startswith(u"lambda"):
            fn = eval(rewritten, glo, {})
        else:
            loc = {}
            exec(rewritten, glo, loc)
            fn = None
            if u"fn" in loc and callable(loc.get(u"fn")):
                fn = loc[u"fn"]
            else:
                for v in loc.values():
                    if callable(v):
                        fn = v
                        break
    except Exception as err:
        raise ValueError(u"компиляция лямбды column name: %s" % err)

    if fn is None or not callable(fn):
        raise ValueError(u"лямбда column name: нужен lambda sheet, headers: … или def …")
    return fn, glo


def compile_lambda_column_sheet_pick(code_text, extra_namespace=None, policy=None):
    """
    compile: callable(sheets:list[str]) -> (str | list/tuple[str])

    Пустой код → None. Ошибка санации/синтаксиса → ValueError.
    """
    raw = expand_source_variables_in_lambda_code(
        unicode(code_text or u"").strip(), log_fn_key=u"лямбда_листы"
    )
    if raw == u"":
        return None

    rewritten = rewrite_lambda_column_sugar(raw, kind=u"sheets")
    from libre_macros_sanitize_lib import (
        LM_SANITIZE_POLICY_STANDARD,
        lm_sanitize_code_or_raise,
        lm_sanitize_filtered_builtins,
        lm_sanitize_format_issues,
        lm_sanitize_code,
    )

    if policy is None:
        policy = LM_SANITIZE_POLICY_STANDARD
    try:
        lm_sanitize_code_or_raise(rewritten, policy=policy)
    except Exception as err:
        try:
            res = lm_sanitize_code(rewritten, policy=policy)
            detail = lm_sanitize_format_issues(res.issues)
            raise ValueError(u"санация лямбды column sheet pick: %s" % detail)
        except ValueError:
            raise
        except Exception:
            raise ValueError(u"санация лямбды column sheet pick: %s" % err)

    builtins_map = lm_sanitize_filtered_builtins()
    glo = _lambda_column_eval_namespace(extra_namespace)
    glo[u"__builtins__"] = builtins_map

    try:
        if rewritten.lstrip().startswith(u"lambda"):
            fn = eval(rewritten, glo, {})
        else:
            loc = {}
            exec(rewritten, glo, loc)
            fn = None
            if u"fn" in loc and callable(loc.get(u"fn")):
                fn = loc[u"fn"]
            else:
                for v in loc.values():
                    if callable(v):
                        fn = v
                        break
    except Exception as err:
        raise ValueError(u"компиляция лямбды column sheet pick: %s" % err)

    if fn is None or not callable(fn):
        raise ValueError(u"лямбда sheet pick: нужен lambda sheets: … или def …")
    return fn, glo


def attach_upd_rows_helpers(glo, upd_store, rows_accessor):
    """Подключить _lm_upd_rows и upd_rows_len/seen/count к namespace лямбды."""
    if glo is None or upd_store is None or rows_accessor is None:
        return
    upd_acc = UpdRowsAccessor(upd_store, rows_accessor)
    glo[u"_lm_upd_rows"] = upd_acc
    glo[u"_lm_upd_store"] = upd_store

    def _norm_val(v):
        return unicode(v or u"").strip().casefold()

    def upd_rows_len():
        ri = int(glo.get(u"_ri") or 0)
        return len(upd_store.indices_before(ri))

    def upd_rows_seen(col, value):
        ri = int(glo.get(u"_ri") or 0)
        want = _norm_val(value)
        if want == u"":
            return False
        for prev in upd_store.indices_before(ri):
            if _norm_val(rows_accessor(prev, col, 0)) == want:
                return True
        return False

    def upd_rows_count(col, value):
        ri = int(glo.get(u"_ri") or 0)
        want = _norm_val(value)
        n = 0
        for prev in upd_store.indices_before(ri):
            if _norm_val(rows_accessor(prev, col, 0)) == want:
                n += 1
        return n

    glo[u"upd_rows_len"] = upd_rows_len
    glo[u"upd_rows_seen"] = upd_rows_seen
    glo[u"upd_rows_count"] = upd_rows_count


def call_rows_lambda(fn, glo, ri, extra=None, rows_accessor=None, upd_store=None):
    """Вызвать скомпилированную лямбду для строки ri. extra → idx/n/text."""
    if fn is None:
        return None
    if rows_accessor is not None:
        glo[u"_lm_rows"] = rows_accessor
    if upd_store is not None and rows_accessor is not None:
        attach_upd_rows_helpers(glo, upd_store, rows_accessor)
    glo[u"_ri"] = int(ri)
    glo[u"row_id"] = int(ri)
    if extra:
        for k, v in extra.items():
            glo[k] = v
    nargs = 1
    try:
        nargs = int(fn.__code__.co_argcount)
    except Exception:
        nargs = 1
    dummy = None
    if nargs <= 0:
        return fn()
    if nargs == 1:
        return fn(dummy)
    if nargs == 2:
        return fn(dummy, glo.get(u"idx", 0))
    return fn(dummy, glo.get(u"idx", 0), glo.get(u"n", 0))


def resolve_column_token_from_titles(token, titles, start_col=0):
    """
    Токен столбца → абсолютный 0-based индекс (как _lm_pp_resolve_column_token).
    titles[i] соответствует столбцу start_col+i.

    'name' exact; 'name*' шаблон; без кавычек — '*part*' / цифра / буква.
    """
    token_raw = unicode(token).strip() if token is not None else u""
    if token_raw == u"":
        return None
    quoted = False
    try:
        from libre_macros_param_codec import unwrap_column_name_token

        token, quoted = unwrap_column_name_token(token_raw)
    except Exception:
        token = token_raw
        quoted = False
    token = unicode(token or u"").strip()
    if token == u"":
        return None

    try:
        from libre_macros_lib import col_letters_to_index, is_col_letters
    except Exception:
        col_letters_to_index = None
        is_col_letters = None

    titles = list(titles or [])
    h_sc = int(start_col)
    h_ec = h_sc + len(titles) - 1

    if not quoted:
        try:
            idx = int(token) - 1
            if idx >= 0:
                return idx
        except (TypeError, ValueError):
            pass
        if is_col_letters is not None and is_col_letters(token) and col_letters_to_index is not None:
            return col_letters_to_index(token)

    try:
        from libre_macros_param_codec import column_header_matches_token

        match_fn = column_header_matches_token
    except Exception:
        match_fn = None

    c = h_sc
    while c <= h_ec:
        pos = c - h_sc
        title = titles[pos] if pos < len(titles) else u""
        if match_fn is not None:
            if match_fn(title, token_raw):
                return c
        else:
            if unicode(title or u"").casefold() == token.casefold():
                return c
            if (not quoted) and token.casefold() in unicode(title or u"").casefold():
                return c
        c += 1
    return None


def _cell_value_for_rows(v):
    if v is None:
        return u""
    return v


class RowsAccessor(object):
    """_lm_rows(ri, col_token, offset) над матрицей data-строк в памяти."""

    def __init__(self, matrix, headers, start_col=0):
        self.matrix = matrix if matrix is not None else []
        self.headers = list(headers or [])
        self.start_col = int(start_col)
        self.n_rows = len(self.matrix)
        self._col_cache = {}

    def resolve_col(self, col_token):
        cache_key = col_token
        if cache_key in self._col_cache:
            return self._col_cache[cache_key]
        abs_c = resolve_column_token_from_titles(
            col_token, self.headers, self.start_col
        )
        self._col_cache[cache_key] = abs_c
        return abs_c

    def __call__(self, ri, col_token, offset=0):
        try:
            off = int(offset or 0)
        except Exception:
            off = 0
        try:
            target = int(ri) + off
        except Exception:
            return None
        if target < 0 or target >= self.n_rows:
            return None
        abs_c = self.resolve_col(col_token)
        if abs_c is None:
            return None
        local = int(abs_c) - self.start_col
        row = self.matrix[target]
        try:
            width = len(row)
        except Exception:
            width = 0
        if local < 0 or local >= width:
            return None
        try:
            v = row[local]
        except Exception:
            v = None
        return _cell_value_for_rows(v)


def make_rows_accessor(matrix, headers, start_col=0):
    return RowsAccessor(matrix, headers, start_col=start_col)


class UpdRowsStore(object):
    """
    Накопитель индексов строк, уже обработанных при последовательном проходе.

    commit(ri) вызывается после успешной обработки строки ri.
    Для строки current_ri доступны только commit-ы с индексом < current_ri.
    """

    def __init__(self):
        self._committed = []

    def reset(self):
        self._committed = []

    def commit(self, ri):
        self._committed.append(int(ri))

    def indices_before(self, current_ri):
        cur = int(current_ri)
        return [x for x in self._committed if int(x) < cur]

    def resolve_target_ri(self, current_ri, offset):
        """
        offset 0 — последняя обработанная строка выше current_ri;
        offset -1 — предпоследняя и т.д. (аналог rows, но по upd_rows).
        """
        avail = self.indices_before(current_ri)
        if not avail:
            return None
        try:
            off = int(offset or 0)
        except Exception:
            off = 0
        pos = len(avail) - 1 + off
        if pos < 0 or pos >= len(avail):
            return None
        return avail[pos]


class UpdRowsAccessor(object):
    """_lm_upd_rows(current_ri, col_token, offset) — значения из upd_rows."""

    def __init__(self, store, rows_accessor):
        self.store = store
        self.rows_accessor = rows_accessor

    def __call__(self, current_ri, col_token, offset=0):
        target_ri = self.store.resolve_target_ri(current_ri, offset)
        if target_ri is None:
            return None
        return self.rows_accessor(target_ri, col_token, 0)


def make_upd_rows_store():
    return UpdRowsStore()


# Примеры для визарда / документации
LAMBDA_ROWS_SUGAR_EXAMPLES = (
    u'lambda rows: str(rows("Столбец") or "").strip() != ""',
    u'lambda rows: rows("Статус") != rows("Статус")[-1]',
    u'lambda rows: not upd_rows_seen("ИНН", rows("ИНН"))',
    u'lambda rows: upd_rows_len() == 0 or rows("Сумма") != upd_rows("Сумма")',
    u'lambda rows: upd_rows_count("Отдел", rows("Отдел")) == 0',
)
