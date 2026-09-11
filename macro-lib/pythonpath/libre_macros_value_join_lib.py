# -*- coding: utf-8 -*-
"""Общий разделитель склейки значений в одной ячейке (ВПР / копирование_диапазонов)."""
from __future__ import unicode_literals
MACRO_VERSION = "3.10.722"
try:
    unicode
except NameError:
    unicode = str


def resolve_value_join_delimiter(raw, default=u"\n"):
    """
    Текст из JSON/combo → реальный разделитель.
    Пусто → default (\\n). Поддерживает экранирование \\n \\t \\r.
    """
    s = unicode(raw if raw is not None else u"").strip()
    if s == u"":
        return unicode(default if default is not None else u"\n")
    # Уже содержит реальный перевод строки — как есть.
    if u"\n" in s or u"\r" in s or u"\t" in s:
        return s.replace(u"\r\n", u"\n").replace(u"\r", u"\n")
    out = []
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == u"\\" and i + 1 < len(s):
            nxt = s[i + 1]
            if nxt == u"n":
                out.append(u"\n")
                i += 2
                continue
            if nxt == u"t":
                out.append(u"\t")
                i += 2
                continue
            if nxt == u"r":
                out.append(u"\n")
                i += 2
                continue
            if nxt == u"\\":
                out.append(u"\\")
                i += 2
                continue
        out.append(ch)
        i += 1
    return u"".join(out)


def escape_value_join_delimiter(delim):
    """Реальный разделитель → строка для JSON/combo (\\n вместо перевода строки)."""
    s = unicode(delim if delim is not None else u"")
    s = s.replace(u"\\", u"\\\\")
    s = s.replace(u"\r\n", u"\\n").replace(u"\n", u"\\n").replace(u"\r", u"\\n")
    s = s.replace(u"\t", u"\\t")
    return s


def normalize_value_join_delimiter_for_store(raw):
    """Канон для JSON: экранированная форма; пусто → не писать / пустая строка."""
    s = unicode(raw if raw is not None else u"").strip()
    if s == u"":
        return u""
    # Если пользователь выбрал пресет с реальным \\n в тексте — экранируем.
    resolved = resolve_value_join_delimiter(s, default=s)
    return escape_value_join_delimiter(resolved)


def join_cell_value_parts(existing, parts, delimiter=None):
    """
    Склеить existing + parts через delimiter (default \\n).
    Пустые parts пропускаются; existing пустой (strip) — не включается.
    """
    delim = resolve_value_join_delimiter(delimiter, default=u"\n")
    chunks = []
    ex = unicode(existing if existing is not None else u"")
    if ex.strip() != u"":
        chunks.append(ex)
    i = 0
    while i < len(parts or []):
        p = parts[i]
        if p is None:
            i += 1
            continue
        if isinstance(p, bool):
            ps = unicode(p)
        elif isinstance(p, (int, float)):
            try:
                if float(p) == int(p):
                    ps = unicode(int(p))
                else:
                    ps = unicode(p)
            except Exception:
                ps = unicode(p)
        else:
            ps = unicode(p)
        if ps != u"":
            chunks.append(ps)
        i += 1
    if not chunks:
        return u""
    return delim.join(chunks)


def cell_is_empty_for_fill(cell_or_text):
    """Пустая ячейка для режима fill_empty."""
    if cell_or_text is None:
        return True
    # UNO cell
    try:
        if hasattr(cell_or_text, "getType") or hasattr(cell_or_text, "String"):
            try:
                t = int(cell_or_text.getType())
            except Exception:
                t = -1
            # EMPTY=0 in Calc
            if t == 0:
                return True
            try:
                s = unicode(cell_or_text.String or u"").strip()
            except Exception:
                s = u""
            if s != u"":
                return False
            try:
                # VALUE may still be 0 — считаем непустым только если type=VALUE
                if t == 1:
                    return False
            except Exception:
                pass
            return True
    except Exception:
        pass
    return unicode(cell_or_text).strip() == u""
