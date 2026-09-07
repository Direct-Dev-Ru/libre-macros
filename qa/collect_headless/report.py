# -*- coding: utf-8 -*-
"""Отчёт headless-прогона (Markdown + JSON)."""
from __future__ import print_function, unicode_literals

import json
import os
from datetime import datetime

try:
    unicode
except NameError:
    unicode = str


def _as_text(raw):
    if raw is None:
        return u""
    try:
        return unicode(raw)
    except Exception:
        return str(raw)


def summarize_case(case_result):
    checks = case_result.get("checks") or []
    fail = sum(1 for c in checks if c.get("status") == "fail")
    skip = sum(1 for c in checks if c.get("status") == "skip")
    passed = sum(1 for c in checks if c.get("status") == "pass")
    run_ok = bool(case_result.get("ok"))
    status = u"pass"
    if case_result.get("timeout"):
        status = u"fail"
    elif not run_ok:
        status = u"fail"
    elif fail:
        status = u"fail"
    return {
        "status": status,
        "pass": passed,
        "fail": fail,
        "skip": skip,
    }


def write_reports(results_dir, payload):
    os.makedirs(results_dir, exist_ok=True)
    json_path = os.path.join(results_dir, "report.json")
    md_path = payload.get("report_path") or os.path.join(results_dir, "report.md")
    parent = os.path.dirname(os.path.abspath(md_path))
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
    lines = [
        u"# Headless-прогон collect_workbooks",
        u"",
        u"Время: %s" % payload.get("started_at", u""),
        u"Версия макросов: %s" % payload.get("macro_version", u""),
        u"Офис: `%s`" % payload.get("office_bin", u""),
        u"",
        u"| Case | Статус | Checks pass/fail/skip | Файл |",
        u"|------|--------|----------------------|------|",
    ]
    for case in payload.get("cases") or []:
        sm = summarize_case(case)
        rel = case.get("save_path") or u""
        lines.append(
            u"| %s | **%s** | %s/%s/%s | `%s` |"
            % (
                case.get("id"),
                sm["status"],
                sm["pass"],
                sm["fail"],
                sm["skip"],
                rel,
            )
        )
    lines.append(u"")
    for case in payload.get("cases") or []:
        lines.append(u"## %s" % case.get("id"))
        if case.get("error"):
            lines.append(u"")
            lines.append(u"Ошибка запуска: `%s`" % _as_text(case.get("error")))
        lines.append(u"")
        lines.append(u"- шаблон: `%s`" % (case.get("template") or u""))
        lines.append(u"- результат: `%s`" % (case.get("save_path") or u""))
        lines.append(u"- макрос ok: %s" % case.get("ok"))
        if case.get("log_excerpt"):
            lines.append(u"")
            lines.append(u"```")
            lines.append(_as_text(case.get("log_excerpt"))[:2000])
            lines.append(u"```")
        checks = case.get("checks") or []
        if checks:
            lines.append(u"")
            lines.append(u"| type | status | detail |")
            lines.append(u"|------|--------|--------|")
            for ch in checks:
                detail = _as_text(ch.get("detail")).replace(u"|", u"\\|")
                lines.append(
                    u"| %s | %s | %s |"
                    % (ch.get("type"), ch.get("status"), detail[:180])
                )
        lines.append(u"")
    with open(md_path, "w", encoding="utf-8") as fh:
        fh.write(u"\n".join(lines) + u"\n")
    return json_path, md_path


def now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
