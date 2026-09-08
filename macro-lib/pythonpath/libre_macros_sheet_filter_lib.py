# -*- coding: utf-8 -*-
"""
Фильтр имён листов: шаблоны text_match + опциональная лямбда (с санацией).

Используется в «Листы», «Листы_не_удалять», «Листы_удалить»
и «удаление_листов» / «скрытие_листов».
"""
from __future__ import print_function
MACRO_VERSION = "3.10.706"
import datetime
import re

try:
    unicode  # type: ignore[name-defined]
except NameError:  # pragma: no cover
    unicode = str

# Токены формата (как у переименовать_лист): длинные первыми.
_DATE_FMT_TOKEN_RE = re.compile(r"(YYYY|YY|MM|DD|HH|mm|SS)")
_DATE_FMT_TOKEN_MAP = {
    u"YYYY": (u"%Y", r"\d{4}"),
    u"YY": (u"%y", r"\d{2}"),
    u"MM": (u"%m", r"\d{2}"),
    u"DD": (u"%d", r"\d{2}"),
    u"HH": (u"%H", r"\d{2}"),
    u"mm": (u"%M", r"\d{2}"),
    u"SS": (u"%S", r"\d{2}"),
}


def _is_text(value):
    return isinstance(value, (str, unicode))


def _fmt_to_strptime_and_regex(fmt):
    """
    YYMMDD-HH-mm → ("%y%m%d-%H-%M", r"\\d{2}\\d{2}\\d{2}-\\d{2}-\\d{2}").
    Литералы экранируются для regex; в strptime «%» удваивается.
    """
    text = unicode(fmt or u"").strip()
    if text == u"":
        raise ValueError("пустой формат")
    sp_parts = []
    rx_parts = []
    for part in _DATE_FMT_TOKEN_RE.split(text):
        if part == u"":
            continue
        mapped = _DATE_FMT_TOKEN_MAP.get(part)
        if mapped is not None:
            sp_parts.append(mapped[0])
            rx_parts.append(mapped[1])
            continue
        sp_parts.append(part.replace(u"%", u"%%"))
        rx_parts.append(re.escape(part))
    sp = u"".join(sp_parts)
    rx = u"".join(rx_parts)
    if sp == u"" or rx == u"":
        raise ValueError("формат без токенов даты")
    return sp, rx


def sheet_name_date(name, fmt=None):
    """
    Дата из имени листа.

    Без fmt — legacy: YYYYMMDD в хвосте или после «_» (Отчет_20260725).
    С fmt — искать фрагмент по шаблону токенов YYYY/YY/MM/DD/HH/mm/SS
    в любом месте имени, например:
      sheet_name_date("Свод_Задач__260813-10-41", "YYMMDD-HH-mm") → date(2026, 8, 13)

    Время в формате учитывается при разборе, возвращается только date.
    None — даты в имени нет / не разобралась.
    """
    s = unicode(name or u"").strip()
    if s == u"":
        return None
    fmt_text = unicode(fmt or u"").strip()
    if fmt_text == u"":
        m = re.search(r"(\d{8})\s*$", s)
        if m is None:
            m = re.search(r"_(\d{8})(?:_|$)", s)
        if m is None:
            return None
        raw = m.group(1)
        try:
            y = int(raw[0:4])
            mo = int(raw[4:6])
            d = int(raw[6:8])
            return datetime.date(y, mo, d)
        except (TypeError, ValueError):
            return None
    try:
        sp, rx = _fmt_to_strptime_and_regex(fmt_text)
    except Exception:
        return None
    m = re.search(rx, s)
    if m is None:
        return None
    raw = m.group(0)
    try:
        dt = datetime.datetime.strptime(raw, sp)
        return dt.date()
    except (TypeError, ValueError):
        return None


def sheet_age_days(name, fmt=None, today=None):
    """
    Сколько полных дней прошло от даты в имени до today (по умолчанию date.today()).
    None — даты в имени нет.

    fmt — шаблон даты/времени в имени (как у sheet_name_date).
    Вызов: sheet_age_days(name, "YYMMDD-HH-mm") или sheet_age_days(name, fmt="…").

    Обратная совместимость: sheet_age_days(name, some_date) — второй аргумент-дата
    трактуется как today (не fmt).

    Пример (Листы_удалить, старше 3 суток):
      lambda name: sheet_age_days(name, "YYMMDD-HH-mm") is not None and sheet_age_days(name, "YYMMDD-HH-mm") > 3
    """
    # sheet_age_days(name, date(...)) — legacy второй аргумент = today
    if fmt is not None and not _is_text(fmt) and today is None:
        today = fmt
        fmt = None
    d = sheet_name_date(name, fmt=fmt)
    if d is None:
        return None
    if today is None:
        today = datetime.date.today()
    elif isinstance(today, datetime.datetime):
        today = today.date()
    try:
        return int((today - d).days)
    except Exception:
        return None


def parse_sheet_filter_code(raw):
    """
    Извлечь текст лямбды/def из JSON C, dict-блока или голой строки.

    Поддержка:
      {"v":1,"filter":"lambda name: …"}
      {"filter":"…"}
      [{"v":1,"filter":"…"}]
      lambda name: …
    """
    if isinstance(raw, dict):
        for key in ("filter", "lambda", "code", "predicate"):
            v = raw.get(key)
            if v is not None and unicode(v).strip() != u"":
                return unicode(v).strip()
        return u""
    text = unicode(raw or u"").strip()
    if text == u"":
        return u""
    if text.startswith(u"{") or text.startswith(u"["):
        try:
            import json

            data = json.loads(text)
        except Exception:
            return text
        if isinstance(data, list) and data:
            data = data[0]
        if isinstance(data, dict):
            return parse_sheet_filter_code(data)
    return text


def sheet_filter_eval_namespace():
    """Имена, доступные в лямбде фильтра имени листа (без import в коде)."""
    return {
        "datetime": datetime,
        "date": datetime.date,
        "timedelta": datetime.timedelta,
        "re": re,
        "sheet_name_date": sheet_name_date,
        "sheet_age_days": sheet_age_days,
    }


def compile_sheet_name_filter(code_text, policy=None):
    """
    Санация + компиляция callable(name) -> bool.

    Пустой код → None. Ошибка санации/синтаксиса → исключение.
    """
    code = unicode(code_text or u"").strip()
    if code == u"":
        return None
    try:
        from libre_macros_lambda_column_lib import expand_source_variables_in_lambda_code

        code = expand_source_variables_in_lambda_code(code, log_fn_key=u"фильтр_листов")
    except Exception:
        pass
    from libre_macros_sanitize_lib import (
        LM_SANITIZE_POLICY_STANDARD,
        lm_sanitize_code_or_raise,
        lm_sanitize_filtered_builtins,
        lm_sanitize_format_issues,
    )

    if policy is None:
        policy = LM_SANITIZE_POLICY_STANDARD
    try:
        lm_sanitize_code_or_raise(code, policy=policy)
    except Exception as err:
        # пробросить с понятным текстом
        try:
            from libre_macros_sanitize_lib import lm_sanitize_code

            res = lm_sanitize_code(code, policy=policy)
            detail = lm_sanitize_format_issues(res.issues)
            raise ValueError(u"санация фильтра листов: %s" % detail)
        except ValueError:
            raise
        except Exception:
            raise ValueError(u"санация фильтра листов: %s" % err)

    builtins_map = lm_sanitize_filtered_builtins()
    glo = {"__builtins__": builtins_map}
    glo.update(sheet_filter_eval_namespace())
    try:
        if code.lstrip().startswith(u"lambda"):
            fn = eval(code, glo, {})
        else:
            loc = {}
            exec(code, glo, loc)
            fn = None
            if "fn" in loc and callable(loc["fn"]):
                fn = loc["fn"]
            else:
                for v in loc.values():
                    if callable(v):
                        fn = v
                        break
    except Exception as err:
        raise ValueError(u"компиляция фильтра листов: %s" % err)
    if fn is None or not callable(fn):
        raise ValueError(u"фильтр листов: нужен lambda name: … или def …")
    return fn


def file_pick_eval_namespace():
    """Имена для лямбды доп. фильтра файлов."""
    ns = sheet_filter_eval_namespace()
    return ns


def compile_file_pick_lambda(code_text, policy=None):
    """
    Компиляция callable(ctx_dict) -> bool для file_pick=lambda.

    Поддерживается:
      lambda f: f['имя'].endswith('.csv')
      lambda path, name, mtime, **kw: mtime > 0
      def fn(f): return True
    Также одноаргументная лямбда от dict или от path (str) — см. обёртку.
    """
    code = unicode(code_text or u"").strip()
    if code == u"":
        return None
    try:
        from libre_macros_lambda_column_lib import expand_source_variables_in_lambda_code

        code = expand_source_variables_in_lambda_code(code, log_fn_key=u"file_pick")
    except Exception:
        pass
    from libre_macros_sanitize_lib import (
        LM_SANITIZE_POLICY_STANDARD,
        lm_sanitize_code_or_raise,
        lm_sanitize_filtered_builtins,
        lm_sanitize_format_issues,
    )

    if policy is None:
        policy = LM_SANITIZE_POLICY_STANDARD
    try:
        lm_sanitize_code_or_raise(code, policy=policy)
    except Exception as err:
        try:
            from libre_macros_sanitize_lib import lm_sanitize_code

            res = lm_sanitize_code(code, policy=policy)
            detail = lm_sanitize_format_issues(res.issues)
            raise ValueError(u"санация file_pick: %s" % detail)
        except ValueError:
            raise
        except Exception:
            raise ValueError(u"санация file_pick: %s" % err)

    builtins_map = lm_sanitize_filtered_builtins()
    glo = {"__builtins__": builtins_map}
    glo.update(file_pick_eval_namespace())
    try:
        if code.lstrip().startswith(u"lambda"):
            raw_fn = eval(code, glo, {})
        else:
            loc = {}
            exec(code, glo, loc)
            raw_fn = None
            if "fn" in loc and callable(loc["fn"]):
                raw_fn = loc["fn"]
            else:
                for v in loc.values():
                    if callable(v):
                        raw_fn = v
                        break
    except Exception as err:
        raise ValueError(u"компиляция file_pick: %s" % err)
    if raw_fn is None or not callable(raw_fn):
        raise ValueError(u"file_pick: нужен lambda … или def …")

    def _wrapped(ctx):
        if not isinstance(ctx, dict):
            ctx = {u"path": unicode(ctx or u""), u"путь": unicode(ctx or u"")}
        try:
            return bool(raw_fn(ctx))
        except TypeError:
            pass
        try:
            return bool(
                raw_fn(
                    ctx.get(u"path"),
                    name=ctx.get(u"name"),
                    mtime=ctx.get(u"mtime"),
                    ctime=ctx.get(u"ctime"),
                    size=ctx.get(u"size"),
                )
            )
        except TypeError:
            pass
        try:
            return bool(raw_fn(ctx.get(u"path") or ctx.get(u"путь") or u""))
        except Exception:
            return False

    return _wrapped


def apply_sheet_name_filter(names, filter_fn):
    """Оставить имена, для которых filter_fn(name) истинно. None filter → как есть."""
    if filter_fn is None:
        return list(names or [])
    out = []
    for name in names or []:
        n = unicode(name or u"").strip()
        if n == u"":
            continue
        try:
            if filter_fn(n):
                out.append(n)
        except Exception:
            pass
    return out


def sheet_name_matches_patterns(name, patterns, text_match_fn=None):
    """True, если имя листа совпадает с любым шаблоном (text_match)."""
    n = unicode(name or u"").strip()
    if n == u"":
        return False
    if text_match_fn is None:
        try:
            from libre_macros_lib import text_match as text_match_fn
        except Exception:
            return False
    for pat in patterns or []:
        p = unicode(pat or u"").strip()
        if p == u"":
            continue
        try:
            if text_match_fn(n, p):
                return True
        except Exception:
            pass
    return False


# Пример для визарда / документации
SHEET_FILTER_EXAMPLE_LAST_3_DAYS = (
    u'lambda name: sheet_age_days(name) is not None and sheet_age_days(name) <= 3'
)
SHEET_FILTER_EXAMPLE_OLDER_3_DAYS_YYMMDD_HH_MM = (
    u'lambda name: sheet_age_days(name, "YYMMDD-HH-mm") is not None '
    u'and sheet_age_days(name, "YYMMDD-HH-mm") > 3'
)
