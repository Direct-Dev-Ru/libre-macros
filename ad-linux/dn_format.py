# -*- coding: utf-8 -*-
"""
Преобразование пути OU в LDAP DN.

Поддерживаемые форматы -o / ou_dn:
  OU=Child,OU=Parent,DC=corp,DC=local     (классический DN, без изменений)
  Parent/Child/corp/local                 (OU слева направо, домен в конце)
  name1/name2/dc1/c/dc2                   (--dc-parts 3 для трёх DC)
"""
from __future__ import annotations


def domain_labels_from_server(server: str | None) -> list[str]:
    """corp.local или dc01.corp.local → ['corp', 'local']."""
    host = (server or "").strip()
    if not host or "." not in host:
        return []
    labels = host.split(".")
    if len(labels) >= 3 and labels[0].lower().startswith("dc"):
        return labels[1:]
    return labels


def normalize_base_dn(value: str | None, server: str | None = None, dc_parts: int | None = None) -> str:
    text = (value or "").strip()
    if not text:
        return text
    if "=" in text:
        return text
    if "/" not in text:
        return text

    parts = [p.strip() for p in text.split("/") if p.strip()]
    if not parts:
        return text

    dc_count = dc_parts
    if dc_count is None:
        domain_labels = domain_labels_from_server(server)
        if domain_labels and len(parts) >= len(domain_labels):
            tail = [p.lower() for p in parts[-len(domain_labels) :]]
            if tail == [d.lower() for d in domain_labels]:
                dc_count = len(domain_labels)
        if dc_count is None:
            dc_count = 2

    if len(parts) < dc_count:
        raise ValueError(
            "Путь «%s»: нужно минимум %d сегмент(ов) для DC (или укажите --dc-parts)"
            % (text, dc_count)
        )

    dc_segments = parts[-dc_count:]
    ou_segments = parts[:-dc_count]

    rdn: list[str] = []
    idx = len(ou_segments) - 1
    while idx >= 0:
        rdn.append("OU=%s" % ou_segments[idx])
        idx -= 1
    for dc in dc_segments:
        rdn.append("DC=%s" % dc)
    return ",".join(rdn)
