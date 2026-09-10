# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.719"
"""
Ядро стека текстовых операций (ops) для PP ``текстовые_операции``
(legacy-алиас: ``нормализация_текста``).

Только stdlib: trim / case / омоглифы / reverse / base64 / lm1-шифр.
Не импортировать из user-eval — sanitize запрещает base64/crypto в плагинах.
"""

import base64
import hashlib
import hmac
import os
import re
import struct
import uuid

try:
    unicode
except NameError:
    unicode = str

FN_KEY = u"текстовые_операции"
CRYPTO_PREFIX = u"lm1:"
BASE64_PREFIX = u"base64:"
SHA256_PREFIX = u"sha256:"

_HOMOGLYPH_BASE = (
    (u"a", u"а"),
    (u"c", u"с"),
    (u"e", u"е"),
    (u"o", u"о"),
    (u"p", u"р"),
    (u"x", u"х"),
    (u"y", u"у"),
)
_HOMOGLYPH_EXTENDED = _HOMOGLYPH_BASE + (
    (u"k", u"к"),
    (u"m", u"м"),
    (u"t", u"т"),
    (u"b", u"в"),
    (u"h", u"н"),
)

_RE_WS = re.compile(r"\s+", re.UNICODE)
_RE_LATIN = re.compile(r"[A-Za-z]", re.UNICODE)
_RE_CYR = re.compile(r"[А-Яа-яЁё]", re.UNICODE)

_ROW_FILTER_ALIASES = {
    u"all": u"all",
    u"все": u"all",
    u"*": u"all",
    u"non_empty": u"non_empty",
    u"nonempty": u"non_empty",
    u"непустые": u"non_empty",
    u"не пустые": u"non_empty",
    u"empty": u"empty",
    u"пустые": u"empty",
    u"lambda": u"lambda",
    u"лямбда": u"lambda",
    u"лямбда выражение": u"lambda",
    u"expr": u"lambda",
}


def normalize_row_filter(raw):
    """Код фильтра строк op: all | non_empty | empty | lambda."""
    key = _u(raw or u"all").strip().casefold()
    if key == u"":
        return u"all"
    mapped = _ROW_FILTER_ALIASES.get(key)
    if mapped:
        return mapped
    if key in (u"all", u"non_empty", u"empty", u"lambda"):
        return key
    return u"all"


class NormalizeOpError(Exception):
    """Ошибка одного op (base64/decrypt и т.п.)."""


def _u(value):
    if value is None:
        return u""
    try:
        if isinstance(value, unicode):
            return value
    except NameError:
        pass
    try:
        return unicode(value)
    except Exception:
        return str(value)


def _has_tag_prefix(text, prefix):
    return _u(text).startswith(prefix)


def _strip_tag_prefix(text, prefix):
    s = _u(text)
    if s.startswith(prefix):
        return s[len(prefix) :]
    return s


def normalize_op_name(raw):
    name = _u(raw).strip().casefold()
    aliases = {
        u"squeeze_spaces": u"collapse_ws",
        u"squeeze": u"collapse_ws",
        u"сжать_пробелы": u"collapse_ws",
        u"capitalize": u"proper",
        u"lang": u"lang_homoglyphs",
        u"homoglyphs": u"lang_homoglyphs",
        u"омоглифы": u"lang_homoglyphs",
        u"b64_encode": u"base64_encode",
        u"b64_decode": u"base64_decode",
        u"sha": u"sha256",
        u"hash": u"sha256",
        u"шифровать": u"encrypt",
        u"расшифровать": u"decrypt",
        u"gen_guid": u"guid",
        u"generate_guid": u"guid",
        u"uuid": u"guid",
        u"uuid4": u"guid",
        u"substring": u"substring",
        u"substr": u"substring",
        u"slice": u"substring",
        u"подстрока": u"substring",
        u"clear": u"clear",
        u"очистка": u"clear",
        u"очистить": u"clear",
        u"регистр": u"case",
        u"обрезать пробелы по краям": u"trim",
        u"убрать лишние пробелы": u"collapse_ws",
        u"омоглифы ru/en": u"lang_homoglyphs",
        u"перевернуть текст": u"reverse",
        u"base64: кодировать": u"base64_encode",
        u"base64: декодировать": u"base64_decode",
        u"sha-256": u"sha256",
        u"guid": u"guid",
        u"зашифровать": u"encrypt",
        u"расшифровать": u"decrypt",
        u"вычислить": u"compute",
        u"compute": u"compute",
        u"calc": u"compute",
        u"calculate": u"compute",
    }
    return aliases.get(name, name)


def parse_op_text(raw):
    """
    Разобрать текст операции из визарда (в т.ч. ручной ввод):
      trim
      Подстрока(1,5)
      substring(2,0)
      очистка
    Возвращает dict {op, ...} или None.
    """
    text = _u(raw).strip()
    if text == u"":
        return None
    m = re.match(r"^([^(]+)\((.*)\)$", text)
    if m is None:
        op = normalize_op_name(text)
        if op == u"":
            return None
        return {u"op": op}
    name = normalize_op_name(m.group(1))
    args = _u(m.group(2)).strip()
    if name == u"substring":
        parts = [p.strip() for p in re.split(r"[;,]", args) if p.strip() != u""]
        if len(parts) < 2:
            raise ValueError(u"Подстрока: нужно 2 числа, пример Подстрока(1,5)")
        try:
            start = int(parts[0])
            length = int(parts[1])
        except Exception:
            raise ValueError(u"Подстрока: параметры должны быть целыми числами")
        return {u"op": u"substring", u"start": int(start), u"length": int(length)}
    # Для остальных ops с (...) сохраняем только имя операции.
    if name == u"":
        return None
    return {u"op": name}


def coerce_ops_list(raw_ops):
    """Строка / список → список dict {op,...}. Строка-коротыш → {op: name}."""
    if raw_ops is None:
        return []
    if isinstance(raw_ops, (bytes, bytearray)):
        try:
            raw_ops = raw_ops.decode("utf-8")
        except Exception:
            raw_ops = _u(raw_ops)
    if isinstance(raw_ops, unicode) or type(raw_ops).__name__ == "str":
        text = _u(raw_ops).strip()
        if text == u"":
            return []
        try:
            import json

            parsed = json.loads(text)
            return coerce_ops_list(parsed)
        except Exception:
            parts = [p.strip() for p in re.split(r"[;\n]+", text) if p.strip()]
            return [{"op": normalize_op_name(p)} for p in parts]
    if not isinstance(raw_ops, (list, tuple)):
        return []
    out = []
    for item in raw_ops:
        if isinstance(item, (unicode, str)) or type(item).__name__ == "str":
            row = parse_op_text(item) or {}
            name = normalize_op_name(row.get(u"op"))
            if name:
                row[u"op"] = name
                if name == u"guid":
                    row[u"row_filter"] = u"empty"
                else:
                    row[u"row_filter"] = u"all"
                out.append(row)
            continue
        if not isinstance(item, dict):
            continue
        op = normalize_op_name(item.get("op") or item.get("name") or item.get("fn"))
        if not op:
            continue
        row = dict(item)
        row["op"] = op
        rf_raw = row.get(u"row_filter")
        if rf_raw is None:
            rf_raw = row.get(u"rows")
        code = u""
        for key in (
            u"row_filter_lambda",
            u"filter_lambda",
            u"filter",
            u"lambda",
            u"predicate",
        ):
            v = row.get(key)
            if v is not None and _u(v).strip() != u"":
                code = _u(v).strip()
                break
        if rf_raw is not None and _u(rf_raw).strip() != u"":
            row[u"row_filter"] = normalize_row_filter(rf_raw)
        elif code:
            row[u"row_filter"] = u"lambda"
        elif op == u"guid":
            row[u"row_filter"] = u"empty"
        else:
            row[u"row_filter"] = u"all"
        if code:
            row[u"row_filter_lambda"] = code
        else:
            row.pop(u"row_filter_lambda", None)
        for drop in (
            u"rows",
            u"filter_lambda",
            u"filter",
            u"lambda",
            u"predicate",
        ):
            row.pop(drop, None)
        if row.get(u"row_filter") != u"lambda":
            row.pop(u"row_filter_lambda", None)
        if op == u"compute":
            expr = u""
            for ek in (u"expr", u"compute_expr", u"value_expr", u"compute_lambda"):
                ev = row.get(ek)
                if ev is not None and _u(ev).strip() != u"":
                    expr = _u(ev).strip()
                    break
            if expr:
                row[u"expr"] = expr
            else:
                row.pop(u"expr", None)
            for ek in (u"compute_expr", u"value_expr", u"compute_lambda"):
                row.pop(ek, None)
        else:
            row.pop(u"expr", None)
            row.pop(u"compute_expr", None)
            row.pop(u"value_expr", None)
            row.pop(u"compute_lambda", None)
        if row.get(u"row_filter") == u"all":
            # all — дефолт, можно не хранить, но визарду удобнее явное поле
            pass
        out.append(row)
    return out


def op_trim(text, _op=None):
    return _u(text).strip()


def op_collapse_ws(text, op=None):
    op = op or {}
    to_space = _u(op.get("to_space") if op.get("to_space") is not None else u" ")
    if to_space == u"":
        to_space = u" "
    t = _u(text).strip()
    return _RE_WS.sub(to_space, t)


def op_case(text, op=None):
    op = op or {}
    mode = _u(op.get("mode") or op.get("case") or u"").strip().casefold()
    if mode in (u"proper", u"capitalize"):
        mode = u"proper"
    t = _u(text)
    if mode == u"lower":
        return t.lower()
    if mode == u"upper":
        return t.upper()
    if mode == u"toggle":
        chars = []
        for ch in t:
            up = ch.upper()
            low = ch.lower()
            if ch == low and ch != up:
                chars.append(up)
            elif ch == up and ch != low:
                chars.append(low)
            else:
                chars.append(ch)
        return u"".join(chars)
    if mode == u"proper":
        parts = []
        for word in re.split(r"(\s+)", t):
            if not word or word.isspace():
                parts.append(word)
                continue
            parts.append(word[:1].upper() + word[1:].lower())
        return u"".join(parts)
    if mode == u"sentence":
        low = t.lower()
        out = []
        cap_next = True
        for ch in low:
            if cap_next and ch.isalpha():
                out.append(ch.upper())
                cap_next = False
            else:
                out.append(ch)
                if ch in u".?!":
                    cap_next = True
        return u"".join(out)
    raise NormalizeOpError(u"case: нужен mode (lower|upper|proper|sentence|toggle)")


def _is_latin(ch):
    return bool(_RE_LATIN.match(ch))


def _is_cyr(ch):
    return bool(_RE_CYR.match(ch))


def _detect_lang(text):
    n_lat = 0
    n_cyr = 0
    for ch in _u(text):
        if _is_latin(ch):
            n_lat += 1
        elif _is_cyr(ch):
            n_cyr += 1
    if n_lat == 0 and n_cyr == 0:
        return None, 0, 0
    if n_cyr >= n_lat:
        return u"ru", n_cyr, n_lat
    return u"en", n_cyr, n_lat


def _pair_maps(extended=False):
    pairs = _HOMOGLYPH_EXTENDED if extended else _HOMOGLYPH_BASE
    lat_to_cyr = {}
    cyr_to_lat = {}
    for lat, cyr in pairs:
        lat_to_cyr[lat] = cyr
        lat_to_cyr[lat.upper()] = cyr.upper()
        cyr_to_lat[cyr] = lat
        cyr_to_lat[cyr.upper()] = lat.upper()
    return lat_to_cyr, cyr_to_lat


def _alphabet_runs(text, is_foreign):
    runs = []
    start = None
    s = _u(text)
    for i, ch in enumerate(s):
        if is_foreign(ch):
            if start is None:
                start = i
        else:
            if start is not None:
                runs.append((start, i))
                start = None
    if start is not None:
        runs.append((start, len(s)))
    return runs


def op_lang_homoglyphs(text, op=None):
    op = op or {}
    t = _u(text)
    if t == u"":
        return t
    force = _u(op.get("force_lang") or u"").strip().casefold()
    if force in (u"ru", u"en"):
        lang = force
    else:
        lang, _n_cyr, _n_lat = _detect_lang(t)
        if lang is None:
            return t
    extended = bool(op.get("extended_pairs") or op.get("extended"))
    protect = op.get("protect_latin_runs")
    if protect is None:
        protect = True
    else:
        protect = bool(protect)
    lat_to_cyr, cyr_to_lat = _pair_maps(extended=extended)
    if lang == u"ru":
        mapping = lat_to_cyr
        is_foreign = _is_latin
    else:
        mapping = cyr_to_lat
        is_foreign = _is_cyr
    runs = _alphabet_runs(t, is_foreign)
    protected = set()
    if protect:
        for a, b in runs:
            if (b - a) >= 2:
                for i in range(a, b):
                    protected.add(i)
    chars = list(t)
    for i, ch in enumerate(chars):
        if i in protected:
            continue
        if not is_foreign(ch):
            continue
        repl = mapping.get(ch)
        if repl is None:
            continue
        chars[i] = repl
    return u"".join(chars)


def op_reverse(text, _op=None):
    return _u(text)[::-1]


def op_base64_encode(text, op=None):
    s = _u(text)
    if _has_tag_prefix(s, BASE64_PREFIX):
        return s
    op = op or {}
    raw = s.encode("utf-8")
    if bool(op.get("urlsafe")):
        enc = base64.urlsafe_b64encode(raw)
    else:
        enc = base64.b64encode(raw)
    if isinstance(enc, memoryview):
        enc = enc.tobytes()
    if isinstance(enc, bytes):
        enc = enc.decode("ascii")
    else:
        enc = _u(enc)
    return BASE64_PREFIX + enc


def op_base64_decode(text, op=None):
    op = op or {}
    s = _u(text).strip()
    s = _strip_tag_prefix(s, BASE64_PREFIX)
    s = _RE_WS.sub(u"", s)
    if s == u"":
        return u""
    urlsafe = _u(op.get("urlsafe") or u"auto").strip().casefold()
    raw = s.encode("ascii", errors="ignore")
    err = None
    data = None
    if urlsafe in (u"1", u"true", u"yes", u"urlsafe"):
        trials = [base64.urlsafe_b64decode]
    elif urlsafe in (u"0", u"false", u"no", u"standard"):
        trials = [base64.b64decode]
    else:
        trials = [base64.b64decode, base64.urlsafe_b64decode]
    for fn in trials:
        try:
            data = fn(raw)
            err = None
            break
        except Exception as e:
            err = e
            data = None
    if data is None:
        raise NormalizeOpError(u"base64_decode: %s" % err)
    try:
        return data.decode("utf-8")
    except Exception:
        return data.decode("utf-8", errors="replace")


def op_sha256(text, op=None):
    """SHA-256 от UTF-8 текста. По умолчанию hex; format=base64 — Base64 digest."""
    s = _u(text)
    if _has_tag_prefix(s, SHA256_PREFIX):
        return s
    op = op or {}
    digest = hashlib.sha256(s.encode("utf-8")).digest()
    fmt = _u(op.get("format") or op.get("encoding") or u"hex").strip().casefold()
    if fmt in (u"base64", u"b64"):
        enc = base64.b64encode(digest)
        if isinstance(enc, bytes):
            body = enc.decode("ascii")
        else:
            body = _u(enc)
    else:
        body = digest.hex() if hasattr(digest, "hex") else digest.encode("hex")
    return SHA256_PREFIX + body


def _derive_key(password):
    return hashlib.sha256(_u(password).encode("utf-8")).digest()


def _keystream(key, nonce, length):
    out = bytearray()
    counter = 0
    while len(out) < length:
        block = hashlib.sha256(key + nonce + struct.pack(">I", counter)).digest()
        out.extend(block)
        counter += 1
    return bytes(out[:length])


def op_encrypt(text, op=None, key=None):
    s = _u(text)
    if _has_tag_prefix(s, CRYPTO_PREFIX):
        return s
    op = op or {}
    password = key if key is not None else op.get("key")
    password = _u(password)
    if password == u"":
        raise NormalizeOpError(u"encrypt: пустой ключ")
    key_b = _derive_key(password)
    nonce = os.urandom(16)
    data = s.encode("utf-8")
    ks = _keystream(key_b, nonce, len(data))
    ct = bytes(a ^ b for a, b in zip(bytearray(data), bytearray(ks)))
    mac = hmac.new(key_b, nonce + ct, hashlib.sha256).digest()
    blob = base64.b64encode(nonce + ct + mac)
    if isinstance(blob, bytes):
        blob = blob.decode("ascii")
    return CRYPTO_PREFIX + _u(blob)


def op_guid(text, op=None):
    """UUID v4; не перезаписывает непустой текст (фильтр строк — row_filter в op)."""
    op = op or {}
    s = _u(text).strip()
    if s != u"":
        return s
    fmt = _u(op.get("format") or u"lower").strip().casefold()
    value = unicode(uuid.uuid4())
    if fmt == u"upper":
        return value.upper()
    if fmt in (u"braces", u"brace", u"{}"):
        return u"{%s}" % value
    return value


def op_substring(text, op=None):
    """
    Подстрока(start,length), где start — 1-based позиция первого символа.
    length=0 → пустая строка.
    """
    op = op or {}
    s = _u(text)
    try:
        start = int(op.get("start", 1))
    except Exception:
        start = 1
    try:
        length = int(op.get("length", op.get("len", 0)))
    except Exception:
        length = 0
    if start < 1:
        start = 1
    if length <= 0:
        return u""
    i0 = int(start) - 1
    if i0 >= len(s):
        return u""
    return s[i0 : i0 + int(length)]


def op_clear(_text, _op=None):
    """Очистка: всегда вернуть пустую строку."""
    return u""


def op_decrypt(text, op=None, key=None):
    op = op or {}
    password = key if key is not None else op.get("key")
    password = _u(password)
    if password == u"":
        raise NormalizeOpError(u"decrypt: пустой ключ")
    s = _u(text).strip()
    if not _has_tag_prefix(s, CRYPTO_PREFIX):
        return s
    s = _strip_tag_prefix(s, CRYPTO_PREFIX)
    try:
        raw = base64.b64decode(_RE_WS.sub(u"", s).encode("ascii"))
    except Exception as err:
        raise NormalizeOpError(u"decrypt: битый blob (%s)" % err)
    if len(raw) < 16 + 32:
        raise NormalizeOpError(u"decrypt: слишком короткий blob")
    nonce = raw[:16]
    mac = raw[-32:]
    ct = raw[16:-32]
    key_b = _derive_key(password)
    expect = hmac.new(key_b, nonce + ct, hashlib.sha256).digest()
    if not hmac.compare_digest(expect, mac):
        raise NormalizeOpError(u"decrypt: неверный ключ или повреждённые данные")
    ks = _keystream(key_b, nonce, len(ct))
    data = bytes(a ^ b for a, b in zip(bytearray(ct), bytearray(ks)))
    try:
        return data.decode("utf-8")
    except Exception:
        return data.decode("utf-8", errors="replace")


_OP_HANDLERS = {
    u"trim": op_trim,
    u"collapse_ws": op_collapse_ws,
    u"case": op_case,
    u"lang_homoglyphs": op_lang_homoglyphs,
    u"reverse": op_reverse,
    u"base64_encode": op_base64_encode,
    u"base64_decode": op_base64_decode,
    u"sha256": op_sha256,
    u"guid": op_guid,
    u"substring": op_substring,
    u"clear": op_clear,
    u"encrypt": op_encrypt,
    u"decrypt": op_decrypt,
}


def known_ops():
    return sorted(list(_OP_HANDLERS.keys()) + [u"compute"])


def op_row_filter_code(op):
    """Текст лямбды из op (row_filter_lambda / filter / lambda)."""
    op = op or {}
    for key in (u"row_filter_lambda", u"filter_lambda", u"filter", u"lambda", u"predicate"):
        v = op.get(key)
        if v is not None and _u(v).strip() != u"":
            return _u(v).strip()
    return u""


def _row_filter_uses_rows_sugar(code):
    """True, если лямбда использует sugar rows/upd_rows."""
    try:
        from libre_macros_lambda_column_lib import lambda_uses_rows_sugar

        return bool(lambda_uses_rows_sugar(code))
    except Exception:
        text = _u(code or u"").strip()
        return u"rows(" in text or u"upd_rows(" in text


def compile_row_filter_lambda(code_text, policy=None):
    """
    Санация + компиляция фильтра строк.

    Возвращает (kind, fn, glo) или None.
    kind:
      rows — sugar rows(\"Col\") / rows(\"Col\")[-1] (glo + accessor);
      row  — legacy lambda row: row.get(…).
    Пустой код → None. Ошибка → ValueError.
    """
    code = _u(code_text or u"").strip()
    if code == u"":
        return None
    if _row_filter_uses_rows_sugar(code):
        from libre_macros_lambda_column_lib import compile_rows_lambda

        packed = compile_rows_lambda(code, policy=policy)
        if packed is None:
            return None
        fn, glo = packed
        return (u"rows", fn, glo)
    try:
        from libre_macros_lambda_column_lib import expand_source_variables_in_lambda_code

        code = expand_source_variables_in_lambda_code(code, log_fn_key=u"фильтр_строк")
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
            raise ValueError(u"санация фильтра строк: %s" % detail)
        except ValueError:
            raise
        except Exception:
            raise ValueError(u"санация фильтра строк: %s" % err)

    builtins_map = lm_sanitize_filtered_builtins()
    glo = {u"__builtins__": builtins_map}
    try:
        import datetime as _dt
        import re as _re

        glo.update(
            {
                u"datetime": _dt,
                u"date": _dt.date,
                u"re": _re,
                u"unicode": unicode,
                u"str": str,
                u"int": int,
                u"float": float,
                u"bool": bool,
                u"len": len,
            }
        )
    except Exception:
        pass
    try:
        if code.lstrip().startswith(u"lambda"):
            fn = eval(code, glo, {})
        else:
            loc = {}
            exec(code, glo, loc)
            fn = None
            if u"fn" in loc and callable(loc[u"fn"]):
                fn = loc[u"fn"]
            else:
                for v in loc.values():
                    if callable(v):
                        fn = v
                        break
    except Exception as err:
        raise ValueError(u"компиляция фильтра строк: %s" % err)
    if fn is None or not callable(fn):
        raise ValueError(u"фильтр строк: нужен lambda row: … / lambda rows: … или def …")
    return (u"row", fn, None)


def compile_ops_row_filters(ops, policy=None):
    """
    Список (mode, fn_or_None) параллельно coerce_ops_list(ops).
    Для mode!=lambda fn всегда None.
    """
    out = []
    for op in coerce_ops_list(ops):
        mode = normalize_row_filter(op.get(u"row_filter") or op.get(u"rows") or u"all")
        fn = None
        if mode == u"lambda":
            fn = compile_row_filter_lambda(op_row_filter_code(op), policy=policy)
        out.append((mode, fn))
    return out


def op_compute_expr_code(op):
    """Текст лямбды вычисления из op."""
    if not isinstance(op, dict):
        return u""
    for key in (u"expr", u"compute_expr", u"value_expr", u"compute_lambda"):
        v = op.get(key)
        if v is not None and _u(v).strip() != u"":
            return _u(v).strip()
    return u""


def compile_compute_lambda(code_text, policy=None):
    """
    Санация + компиляция лямбды значения (как столбец_по_лямбде / фильтр rows).
    → (fn, glo) или None. Ошибка → ValueError.
    """
    code = _u(code_text or u"").strip()
    if code == u"":
        return None
    from libre_macros_lambda_column_lib import compile_rows_lambda

    return compile_rows_lambda(code, policy=policy)


def compile_ops_compute_fns(ops, policy=None):
    """Список (fn, glo)|None параллельно coerce_ops_list(ops)."""
    out = []
    for op in coerce_ops_list(ops):
        if normalize_op_name(op.get(u"op")) != u"compute":
            out.append(None)
            continue
        code = op_compute_expr_code(op)
        if code == u"":
            raise ValueError(u"вычислить: нужна лямбда expr (lambda rows: …)")
        packed = compile_compute_lambda(code, policy=policy)
        if packed is None:
            raise ValueError(u"вычислить: пустая лямбда expr")
        out.append(packed)
    return out


def _eval_compute_packed(packed, row_id=None, rows_accessor=None, upd_store=None, text=u""):
    if packed is None:
        return text
    fn, glo = packed
    from libre_macros_lambda_column_lib import call_rows_lambda

    ri = 0 if row_id is None else int(row_id)
    try:
        glo[u"text"] = _u(text)
        glo[u"idx"] = ri
        glo[u"n"] = ri
    except Exception:
        pass
    result = call_rows_lambda(
        fn,
        glo,
        ri,
        rows_accessor=rows_accessor,
        upd_store=upd_store,
    )
    if result is None:
        return u""
    if isinstance(result, bool):
        return u"true" if result else u"false"
    return _u(result)


def row_filter_allows(mode, fn, cell_value=None, row=None, row_id=None, rows_accessor=None, upd_store=None):
    """True — применить op к ячейке."""
    mode = normalize_row_filter(mode)
    if mode == u"all":
        return True
    if mode == u"non_empty":
        return not cell_is_blank(cell_value)
    if mode == u"empty":
        return cell_is_blank(cell_value)
    if mode == u"lambda":
        if fn is None:
            return True
        kind = u"row"
        inner = fn
        glo = None
        if isinstance(fn, (tuple, list)) and len(fn) >= 2:
            if fn[0] in (u"row", u"rows"):
                kind = fn[0]
                inner = fn[1]
                glo = fn[2] if len(fn) > 2 else None
            elif callable(fn[0]):
                kind = u"rows"
                inner = fn[0]
                glo = fn[1]
        try:
            if kind == u"rows" and callable(inner):
                from libre_macros_lambda_column_lib import call_rows_lambda

                if glo is not None and rows_accessor is not None:
                    glo[u"_lm_rows"] = rows_accessor
                if upd_store is not None and rows_accessor is not None:
                    from libre_macros_lambda_column_lib import attach_upd_rows_helpers

                    attach_upd_rows_helpers(glo, upd_store, rows_accessor)
                ri = row_id
                if ri is None and isinstance(row, dict):
                    try:
                        ri = int(row.get(u"row_id") or 0)
                    except Exception:
                        ri = 0
                if ri is None:
                    ri = 0
                return bool(call_rows_lambda(inner, glo if glo is not None else {}, ri))
            if callable(inner):
                return bool(inner(row if isinstance(row, dict) else {}))
        except Exception:
            return False
        return False
    return True


# Индекс столбца в row.get: цифры 1… или 1–3 латинские буквы ВЕРХНЕГО регистра (A…ZZZ).
_RE_ROW_COL_DIGITS = re.compile(r"^[1-9]\d*$")
_RE_ROW_COL_LETTERS = re.compile(r"^[A-Z]{1,3}$")


def _letters_to_col_index_0(letters):
    """A→0, B→1, … (только A–Z)."""
    letters = _u(letters).strip().upper()
    n = 0
    i = 0
    while i < len(letters):
        ch = letters[i]
        if ch < u"A" or ch > u"Z":
            return None
        n = n * 26 + (ord(ch) - ord(u"A") + 1)
        i = i + 1
    if n <= 0:
        return None
    return n - 1


def resolve_row_col_index_alias(key):
    """
    Если key похож на индекс столбца — вернуть «Col_N» (1-based), иначе None.

    Цифры: «1», «2», … → Col_1, Col_2.
    Буквы ВЕРХНЕГО регистра: «A»…«ZZZ» (1–3 латиницы) → Col_*.
    Строчные / смешанный регистр / кириллица — не индекс.
    """
    if isinstance(key, bool):
        return None
    if isinstance(key, int):
        if key >= 1:
            return u"Col_%d" % int(key)
        return None
    k = _u(key).strip()
    if k == u"":
        return None
    if _RE_ROW_COL_DIGITS.match(k):
        return u"Col_%d" % int(k)
    if _RE_ROW_COL_LETTERS.match(k):
        idx0 = _letters_to_col_index_0(k)
        if idx0 is None:
            return None
        return u"Col_%d" % (idx0 + 1)
    return None


class FilterRowDict(dict):
    """
    Словарь строки для λ-фильтра.

    row.get(key): сначала точный ключ заголовка / Col_N / row_id;
    если нет и key — индекс («2», «B», 2) — значение столбца по номеру/букве.
    Так «B» в заголовке не конфликтует с буквенным индексом B: заголовок побеждает.
    """

    def _lookup_key(self, key):
        if dict.__contains__(self, key):
            return key
        try:
            k = _u(key)
        except Exception:
            k = None
        if k is not None and k != key and dict.__contains__(self, k):
            return k
        alias = resolve_row_col_index_alias(key)
        if alias is not None and dict.__contains__(self, alias):
            return alias
        return None

    def get(self, key, default=None):
        found = self._lookup_key(key)
        if found is None:
            return default
        return dict.__getitem__(self, found)

    def __getitem__(self, key):
        found = self._lookup_key(key)
        if found is None:
            raise KeyError(key)
        return dict.__getitem__(self, found)

    def __contains__(self, key):
        return self._lookup_key(key) is not None


def build_row_dict(headers, values, row_id):
    """
    Словарь строки для лямбды фильтра: ключи = заголовки (или Col_N),
    плюс row_id — 0-based индекс в диапазоне данных.
    Всегда есть алиасы Col_1…Col_N (1-based номер столбца слева, A=1).
    get(«2»)/get(«B») — индекс столбца, если нет одноимённого заголовка.
    Типы values — как из getDataArray / openpyxl (str/int/float/None/…).
    """
    row = FilterRowDict()
    row[u"row_id"] = int(row_id)
    n = max(len(headers or ()), len(values or ()))
    i = 0
    while i < n:
        title = u""
        if headers is not None and i < len(headers):
            title = _u(headers[i]).strip()
        if title == u"":
            title = u"Col_%d" % (i + 1)
        val = None
        if values is not None and i < len(values):
            val = values[i]
        if title.casefold() == u"row_id":
            title = u"Col_row_id"
        row[title] = val
        col_alias = u"Col_%d" % (i + 1)
        if col_alias.casefold() != title.casefold():
            row[col_alias] = val
        i = i + 1
    return row


# Примеры для визарда / проверки фильтрации:
ROW_FILTER_LAMBDA_EXAMPLE = (
    u'lambda rows: (rows("Сумма") or 0) > 1000'
)
ROW_FILTER_LAMBDA_EXAMPLE_EVEN = (
    u"lambda rows: int(row_id or 0) % 2 == 0"
)
ROW_FILTER_LAMBDA_EXAMPLE_NONEMPTY_COL = (
    u'lambda rows: str(rows("Наименование МТР_2") or "").strip() != ""'
)
ROW_FILTER_LAMBDA_EXAMPLE_COL_INDEX = (
    u'lambda rows: str(rows("B") or "").strip() != ""'
)
ROW_FILTER_LAMBDA_EXAMPLE_PREV_ROW = (
    u'lambda rows: rows("Статус") != rows("Статус")[-1]'
)
ROW_FILTER_LAMBDA_EXAMPLE_UPD_ROWS = (
    u'lambda rows: not upd_rows_seen("ИНН", rows("ИНН"))'
)
ROW_FILTER_LAMBDA_EXAMPLE_UPD_PREV = (
    u'lambda rows: upd_rows_len() == 0 or rows("Сумма") != upd_rows("Сумма")'
)


def apply_ops( text, ops, key=None, on_error=u"keep", cell_value=None, row=None, row_filters=None, row_id=None, rows_accessor=None, upd_store=None, compute_fns=None ):
    """
    Применить стек ops к тексту.
    Возвращает (result_text, error_or_None, ops_applied_count).

    row_filters — список (mode, fn) из compile_ops_row_filters; если None —
    фильтр читается из каждого op (лямбда компилируется на лету — медленнее).
    compute_fns — список (fn, glo)|None из compile_ops_compute_fns; если None —
    лямбды compute компилируются на лету.
    cell_value — исходное значение ячейки (для non_empty/empty).
    row — dict строки (+ row_id) для legacy-лямбды lambda row: ….
    row_id / rows_accessor — для sugar lambda rows: rows(\"Col\").
    upd_store — накопитель уже обработанных строк (upd_rows / upd_rows_seen).
    """
    cur = _u(text)
    applied = 0
    policy = _u(on_error or u"keep").strip().casefold()
    if policy == u"skip_cell":
        policy = u"keep"
    ops_list = coerce_ops_list(ops)
    if row_filters is None:
        try:
            row_filters = compile_ops_row_filters(ops_list)
        except Exception:
            row_filters = [(u"all", None)] * len(ops_list)
    if compute_fns is None:
        try:
            compute_fns = compile_ops_compute_fns(ops_list)
        except Exception:
            compute_fns = [None] * len(ops_list)
    i = 0
    while i < len(ops_list):
        raw = ops_list[i]
        mode, fn = (u"all", None)
        if i < len(row_filters):
            mode, fn = row_filters[i]
        else:
            mode = normalize_row_filter(raw.get(u"row_filter") or u"all")
        if not row_filter_allows(
            mode,
            fn,
            cell_value=cell_value,
            row=row,
            row_id=row_id,
            rows_accessor=rows_accessor,
            upd_store=upd_store,
        ):
            i = i + 1
            continue
        op_name = normalize_op_name(raw.get("op"))
        try:
            if op_name == u"compute":
                packed = None
                if compute_fns is not None and i < len(compute_fns):
                    packed = compute_fns[i]
                if packed is None:
                    packed = compile_compute_lambda(op_compute_expr_code(raw))
                if packed is None:
                    raise NormalizeOpError(u"вычислить: нужна лямбда expr (lambda rows: …)")
                cur = _eval_compute_packed(
                    packed,
                    row_id=row_id,
                    rows_accessor=rows_accessor,
                    upd_store=upd_store,
                    text=cur,
                )
                applied += 1
            else:
                handler = _OP_HANDLERS.get(op_name)
                if handler is None:
                    err = NormalizeOpError(u"неизвестный op: %s" % op_name)
                    if policy == u"stop":
                        raise err
                    if policy == u"empty":
                        return u"", err, applied
                    return cur, err, applied
                if op_name in (u"encrypt", u"decrypt"):
                    cur = handler(cur, raw, key=key)
                else:
                    cur = handler(cur, raw)
                applied += 1
        except NormalizeOpError as err:
            if policy == u"stop":
                raise
            if policy == u"empty":
                return u"", err, applied
            return cur, err, applied
        except Exception as err:
            wrapped = NormalizeOpError(u"%s: %s" % (op_name, err))
            if policy == u"stop":
                raise wrapped
            if policy == u"empty":
                return u"", wrapped, applied
            return cur, wrapped, applied
        i = i + 1
    return cur, None, applied


def parse_key_cell_ref(ref):
    """«Лист!A1» / {sheet, cell} → (sheet_or_None, a1)."""
    if isinstance(ref, dict):
        sheet = _u(ref.get("sheet") or ref.get("source_sheet") or u"").strip()
        cell = _u(ref.get("cell") or ref.get("addr") or u"").strip()
        if cell == u"":
            col = ref.get("col") or ref.get("column")
            row = ref.get("row")
            if col is not None and row is not None:
                try:
                    from libre_macros_lib import _lm_pp_col_index_to_letters

                    c0 = int(col)
                    if c0 >= 1:
                        c0 = c0 - 1
                    cell = u"%s%d" % (_lm_pp_col_index_to_letters(c0), int(row))
                except Exception:
                    cell = u""
        return (sheet or None), cell
    text = _u(ref).strip()
    if text == u"":
        return None, u""
    if u"!" in text:
        left, right = text.split(u"!", 1)
        left = left.strip().strip(u"'").strip(u'"')
        return left, right.strip()
    return None, text


def expand_crypto_key_variables(text, sheet_name_hint=None):
    """
    Подставить <<Переменные.…>> в ключе шифра (key / key_cell / key_file)
    из runtime-карты сбора (в т.ч. глобальные переменные).
    """
    s = _u(text)
    if s == u"" or u"<<" not in s:
        return s
    try:
        import libre_macros_collect_cfg as _cw_cfg
        import libre_macros_lib as lm
    except Exception:
        return s
    prev = lm.lm_pp_active_context()
    ctx = dict(prev or {})
    ctx["source_variables_map"] = getattr(_cw_cfg, "_MERGE_SOURCE_VARIABLES_MAP", {}) or {}
    ctx["source_variables_order"] = getattr(_cw_cfg, "_MERGE_SOURCE_VARIABLES_ORDER", []) or []
    lm.lm_pp_set_active_context(ctx)
    try:
        out = lm._lm_pp_expand_source_variables_template(
            s,
            for_formula=False,
            log_fn_key=u"текстовые_операции",
            sheet_name_hint=_u(sheet_name_hint or u""),
        )
    finally:
        lm.lm_pp_set_active_context(prev)
    return _u(out)


def resolve_crypto_key(block, doc=None, default_sheet=None, read_file=True):
    """Из блока: key | key_cell | key_file → строка ключа (с <<Переменные…>>)."""
    block = block or {}
    channel = set()
    lit = None
    cell_ref = None
    file_path = None
    pairs = [
        (u"key", block.get("key")),
        (u"key_cell", block.get("key_cell")),
        (u"key_file", block.get("key_file")),
    ]
    for op in coerce_ops_list(block.get("ops")):
        if op.get("op") not in (u"encrypt", u"decrypt"):
            continue
        pairs.extend(
            [
                (u"key", op.get("key")),
                (u"key_cell", op.get("key_cell")),
                (u"key_file", op.get("key_file")),
            ]
        )
    for name, value in pairs:
        if value in (None, u"", {}):
            continue
        channel.add(name)
        if name == u"key":
            lit = value
        elif name == u"key_cell":
            cell_ref = value
        elif name == u"key_file":
            file_path = value

    if len(channel) > 1:
        raise NormalizeOpError(
            u"ключ: задайте только один из key / key_cell / key_file"
        )
    if not channel:
        return None

    sheet_hint = u""
    try:
        if default_sheet is not None:
            sheet_hint = _u(getattr(default_sheet, "Name", u"") or u"")
    except Exception:
        sheet_hint = u""

    if u"key" in channel:
        key = expand_crypto_key_variables(lit, sheet_name_hint=sheet_hint)
        if u"<<" in key and u"Переменные" in key:
            raise NormalizeOpError(
                u"key: не удалось подставить <<Переменные…>> (нет в карте): %s" % key
            )
        return key
    if u"key_file" in channel:
        path = expand_crypto_key_variables(file_path, sheet_name_hint=sheet_hint).strip()
        if u"<<" in path and u"Переменные" in path:
            raise NormalizeOpError(
                u"key_file: не удалось подставить <<Переменные…>>: %s" % path
            )
        if not read_file:
            return path
        if path == u"" or not os.path.isfile(path):
            raise NormalizeOpError(u"key_file не найден: %s" % path)
        try:
            with open(path, "r", encoding="utf-8") as fh:
                line = fh.readline()
        except TypeError:
            with open(path, "r") as fh:
                line = fh.readline()
                if isinstance(line, bytes):
                    line = line.decode("utf-8", errors="replace")
        except Exception as err:
            raise NormalizeOpError(u"key_file: %s" % err)
        return _u(line).strip(u"\r\n")
    sheet_name, a1 = parse_key_cell_ref(cell_ref)
    if not a1:
        raise NormalizeOpError(u"key_cell: пустая ссылка")
    if doc is None:
        raise NormalizeOpError(u"key_cell: нет документа для чтения")
    try:
        sheets = doc.Sheets
        if sheet_name:
            sh = sheets.getByName(sheet_name)
        elif default_sheet is not None:
            sh = default_sheet
        else:
            sh = sheets.getByIndex(0)
        cell = sh.getCellRangeByName(a1)
        try:
            from libre_macros_lib import cell_text

            raw = _u(cell_text(cell)).strip()
        except Exception:
            try:
                raw = _u(cell.getString()).strip()
            except Exception:
                raw = _u(cell.getValue())
        return expand_crypto_key_variables(raw, sheet_name_hint=sheet_hint)
    except NormalizeOpError:
        raise
    except Exception as err:
        raise NormalizeOpError(u"key_cell «%s»: %s" % (_u(cell_ref), err))


def needs_crypto_key(ops):
    for op in coerce_ops_list(ops):
        if op.get("op") in (u"encrypt", u"decrypt"):
            return True
    return False


def cell_is_blank(value):
    """True для None / пустой строки (пробелы = пусто). Числа/bool — не blank."""
    if value is None:
        return True
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return False
    try:
        return _u(value).strip() == u""
    except Exception:
        return False


def cell_value_to_text(value, non_text=u"skip"):
    """
    Привести значение ячейки к тексту или None (пропуск).
    non_text: skip | coerce | error
    """
    policy = _u(non_text or u"skip").strip().casefold()
    if value is None:
        return u"" if policy == u"coerce" else None
    if isinstance(value, bool):
        if policy == u"skip":
            return None
        if policy == u"error":
            raise NormalizeOpError(u"нетекст (bool)")
        return u"TRUE" if value else u"FALSE"
    if isinstance(value, (int, float)):
        if policy == u"skip":
            return None
        if policy == u"error":
            raise NormalizeOpError(u"нетекст (число)")
        # coerce: как str без лишних .0
        try:
            if isinstance(value, float) and value == int(value):
                return _u(int(value))
        except Exception:
            pass
        return _u(value)
    s = _u(value)
    return s


def cell_text_for_ops(value, non_text=u"skip", ops=None, empty_cells=u"skip"):
    """
    Текст для стека ops или None (пропуск ячейки).

    Пустые ячейки → ``""`` (обработку режет row_filter каждого op).
    Числа/даты/bool — по ``non_text``.
    """
    _ = ops, empty_cells
    if cell_is_blank(value):
        return u""
    return cell_value_to_text(value, non_text=non_text)
