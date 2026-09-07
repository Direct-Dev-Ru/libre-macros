# -*- coding: utf-8 -*-
"""Преобразование пути OU в LDAP DN (общее для макроса Calc)."""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.691"
try:
    unicode
except NameError:
    unicode = str


def domain_labels_from_server(server):
    host = unicode(server or u"").strip()
    if host == u"" or u"." not in host:
        return []
    labels = host.split(u".")
    if len(labels) >= 3 and labels[0].lower().startswith(u"dc"):
        return labels[1:]
    return labels


def normalize_base_dn(value, server=None, dc_parts=None):
    text = unicode(value or u"").strip()
    if text == u"":
        return text
    if u"=" in text:
        return text
    if u"/" not in text:
        return text

    parts = [p.strip() for p in text.split(u"/") if p.strip()]
    if not parts:
        return text

    dc_count = dc_parts
    if dc_count is None:
        domain_labels = domain_labels_from_server(server)
        if domain_labels and len(parts) >= len(domain_labels):
            tail = [p.lower() for p in parts[-len(domain_labels) :]]
            ref = [unicode(d).lower() for d in domain_labels]
            if tail == ref:
                dc_count = len(domain_labels)
        if dc_count is None:
            dc_count = 2

    if len(parts) < dc_count:
        raise ValueError(
            u"Путь «%s»: нужно минимум %d сегмент(ов) для DC"
            % (text, dc_count)
        )

    dc_segments = parts[-dc_count:]
    ou_segments = parts[:-dc_count]

    rdn = []
    idx = len(ou_segments) - 1
    while idx >= 0:
        rdn.append(u"OU=%s" % ou_segments[idx])
        idx = idx - 1
    for dc in dc_segments:
        rdn.append(u"DC=%s" % dc)
    return u",".join(rdn)
