# -*- coding: utf-8 -*-
"""CLI headless-прогона collect_workbooks.

  python qa/collect_headless/run.py --manifest qa/collect_headless/suites/default.yaml
  python qa/collect_headless/run.py --templates-dir DIR --results-dir DIR
"""
from __future__ import print_function, unicode_literals

import argparse
import json
import os
import shutil
import sys
import traceback

HERE = os.path.abspath(os.path.dirname(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from checks import default_checks, run_checks  # noqa: E402
from office import (  # noqa: E402
    find_office_bin,
    find_office_python,
    install_macros_to_profile,
    repo_root,
    start_office,
    start_xvfb_if_needed,
    stop_proc,
    wait_port,
    write_low_macro_security,
)
from patch_sources import patch_workbook_copy  # noqa: E402
from report import now_iso, summarize_case, write_reports  # noqa: E402

try:
    unicode
except NameError:
    unicode = str

TEMPLATE_EXTS = (".ots", ".ods", ".xltx", ".xlsx", ".xlsm")


def _as_text(raw):
    if raw is None:
        return u""
    try:
        return unicode(raw)
    except Exception:
        return str(raw)


def read_macro_version(root):
    path = os.path.join(root, "macro-lib", "version.txt")
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except Exception:
        return u""


def load_manifest(path):
    path = os.path.abspath(path)
    with open(path, "r", encoding="utf-8") as fh:
        text = fh.read()
    stripped = text.lstrip()
    if stripped.startswith(u"{") or path.lower().endswith(".json"):
        data = json.loads(text)
    else:
        try:
            import yaml  # type: ignore
        except ImportError:
            try:
                data = json.loads(text)
            except Exception:
                raise SystemExit(
                    "Манифест YAML: установите PyYAML или сохраните файл как JSON.\n%s"
                    % path
                )
        else:
            data = yaml.safe_load(text)
    if not isinstance(data, dict):
        raise SystemExit("Манифест должен быть объектом")
    data["_manifest_path"] = path
    data["_manifest_dir"] = os.path.dirname(path)
    return data


def _abs_from(base, rel):
    rel = _as_text(rel).strip()
    if rel == u"":
        return u""
    if os.path.isabs(rel):
        return rel
    return os.path.abspath(os.path.join(base, rel))


def glob_cases(templates_dir, pattern, results_dir):
    import fnmatch

    cases = []
    names = sorted(os.listdir(templates_dir))
    for name in names:
        if not fnmatch.fnmatch(name, pattern):
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext not in TEMPLATE_EXTS:
            continue
        stem = os.path.splitext(name)[0]
        save_ext = ".ods" if ext in (".ots", ".ods") else ".xlsx"
        cases.append(
            {
                "id": stem,
                "template": name,
                "save_as": stem + save_ext,
                "checks": default_checks(),
            }
        )
    return cases


def resolve_cases(manifest, root, args):
    mdir = manifest.get("_manifest_dir") or root
    templates_dir = args.templates_dir or manifest.get("templates_dir") or ""
    templates_dir = _abs_from(root, templates_dir) if templates_dir else u""
    if args.templates_dir:
        templates_dir = os.path.abspath(args.templates_dir)
    sources_dir = args.sources_dir or manifest.get("sources_dir") or ""
    sources_dir = _abs_from(root, sources_dir) if sources_dir else os.path.join(
        root, "test", "collect_workbooks", "sources"
    )
    results_dir = args.results_dir or manifest.get("results_dir") or ""
    results_dir = _abs_from(root, results_dir) if results_dir else os.path.join(
        root, "test", "collect_workbooks", "workbooks", "results", "headless"
    )
    cases = list(manifest.get("cases") or [])
    if args.templates_dir and not cases:
        cases = glob_cases(templates_dir, args.glob, results_dir)
    elif not cases and templates_dir:
        cases = glob_cases(templates_dir, args.glob, results_dir)
    if args.only:
        want = set(args.only)
        cases = [c for c in cases if _as_text(c.get("id")) in want]
    for case in cases:
        if case.get("checks_file"):
            cf = _abs_from(mdir, case["checks_file"])
            extra = load_manifest(cf)
            if isinstance(extra, dict) and extra.get("checks"):
                case["checks"] = extra["checks"]
            elif isinstance(extra, list):
                case["checks"] = extra
        if not case.get("checks"):
            case["checks"] = list(manifest.get("checks") or default_checks())
    return {
        "templates_dir": templates_dir,
        "sources_dir": sources_dir,
        "results_dir": results_dir,
        "cases": cases,
        "office": args.office or manifest.get("office") or "auto",
        "timeout_sec": int(args.timeout or manifest.get("timeout_sec") or 300),
        "copy_template": bool(manifest.get("copy_template", True)),
        "param_sheet": manifest.get("param_sheet") or u"",
        "report_path": _abs_from(root, manifest.get("report_path") or u"")
        or os.path.join(results_dir, "report.md"),
        "host": os.environ.get("LM_QA_HOST") or "127.0.0.1",
        "port": int(os.environ.get("LM_QA_PORT") or manifest.get("port") or 2002),
    }


def log_excerpt(save_path):
    try:
        from checks import _log_blob, load_workbook_grid

        grids = load_workbook_grid(save_path)
        blob = _log_blob(grids)
        lines = blob.splitlines()
        return u"\n".join(lines[-25:])
    except Exception:
        return u""


def run_one_case(cfg, case, office_bin, office_python, profile_dir, keep_office, office_proc):
    root = repo_root()
    templates_dir = cfg["templates_dir"]
    tmpl_name = _as_text(case.get("template") or u"")
    tmpl_path = tmpl_name
    if tmpl_name and not os.path.isabs(tmpl_name):
        tmpl_path = os.path.join(templates_dir, tmpl_name)
    if not os.path.isfile(tmpl_path):
        raise RuntimeError("нет шаблона: %s" % tmpl_path)
    results_dir = cfg["results_dir"]
    run_dir = os.path.join(results_dir, "_run", _as_text(case.get("id")))
    if os.path.isdir(run_dir):
        shutil.rmtree(run_dir, ignore_errors=True)
    os.makedirs(run_dir, exist_ok=True)
    ext = os.path.splitext(tmpl_path)[1]
    work_copy = os.path.join(run_dir, "template" + ext)
    save_as = _as_text(case.get("save_as") or (case.get("id") + ".ods"))
    save_path = os.path.join(results_dir, save_as)
    xml_out = os.path.join(run_dir, "xml_result")
    os.makedirs(xml_out, exist_ok=True)
    patch_workbook_copy(tmpl_path, work_copy, cfg["sources_dir"], result_dir=xml_out)
    job_path = os.path.join(run_dir, "job.json")
    job = {
        "doc_path": os.path.abspath(work_copy),
        "save_path": os.path.abspath(save_path),
        "param_sheet": _as_text(case.get("param_sheet") or cfg.get("param_sheet") or u""),
        "host": cfg["host"],
        "port": cfg["port"],
        "as_template": True,
        "skip_consent": True,
        "skip_confirm": True,
        "quiet_finish": True,
        "suppress_pre_clear": bool(case.get("suppress_pre_clear", False)),
    }
    with open(job_path, "w", encoding="utf-8") as fh:
        json.dump(job, fh, ensure_ascii=False, indent=2)
    uno_job = os.path.join(HERE, "uno_job.py")
    env = os.environ.copy()
    env["LM_QA_JOB"] = job_path
    program_dir = os.path.dirname(os.path.abspath(office_bin))
    import subprocess

    proc = office_proc
    started_here = False
    if proc is None:
        proc, _cmd = start_office(
            office_bin, profile_dir, host=cfg["host"], port=cfg["port"]
        )
        started_here = True
        wait_port(cfg["host"], cfg["port"], timeout_sec=min(60, cfg["timeout_sec"]))
    try:
        completed = subprocess.run(
            [office_python, uno_job, job_path],
            timeout=cfg["timeout_sec"],
            env=env,
            cwd=program_dir,
        )
        timeout = False
        rc = int(completed.returncode or 0)
    except subprocess.TimeoutExpired:
        timeout = True
        rc = 1
        stop_proc(proc)
        proc = None
    marker = save_path + ".qa.json"
    ok = False
    err = u""
    if os.path.isfile(marker):
        try:
            with open(marker, "r", encoding="utf-8") as fh:
                meta = json.load(fh)
            ok = bool(meta.get("ok"))
            err = _as_text(meta.get("error") or u"")
            if not ok and not err:
                err = _as_text(meta.get("parse_err") or u"")
        except Exception as exc:
            err = _as_text(exc)
    if timeout:
        err = err or u"timeout %ss" % cfg["timeout_sec"]
        ok = False
    elif rc != 0 and not ok:
        err = err or u"uno_job exit %s" % rc
    checks_out = []
    if os.path.isfile(save_path) and not cfg.get("no_checks"):
        checks_out = run_checks(
            save_path,
            case.get("checks") or default_checks(),
            version_hint=read_macro_version(root),
        )
    result = {
        "id": case.get("id"),
        "template": tmpl_path,
        "save_path": save_path,
        "ok": ok,
        "timeout": timeout,
        "error": err,
        "checks": checks_out,
        "log_excerpt": log_excerpt(save_path) if os.path.isfile(save_path) else u"",
        "uno_exit": rc,
    }
    if started_here and not keep_office:
        stop_proc(proc)
        proc = None
    return result, proc


def parse_args(argv):
    p = argparse.ArgumentParser(description="Headless-прогон collect_workbooks")
    p.add_argument("--manifest", help="YAML/JSON манифест сценариев")
    p.add_argument("--templates-dir", help="Каталог шаблонов (.ots/.ods/.xltx)")
    p.add_argument("--results-dir", help="Куда сохранять результаты")
    p.add_argument("--sources-dir", help="Каталог sources/ для подмены путей")
    p.add_argument("--glob", default="*.ots", help="Маска файлов без манифеста")
    p.add_argument("--only", nargs="+", help="Только эти id")
    p.add_argument("--office", default="", help="acell | soffice | auto")
    p.add_argument("--timeout", type=int, default=0, help="Секунд на case")
    p.add_argument("--keep-office", action="store_true", help="Один процесс офиса на все case")
    p.add_argument("--no-install", action="store_true", help="Не копировать макросы в QA-профиль")
    p.add_argument("--no-checks", action="store_true", help="Только прогон и сохранение")
    p.add_argument(
        "--profile",
        default="",
        help="Каталог UserInstallation (по умолчанию qa/collect_headless/.profile)",
    )
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(sys.argv[1:] if argv is None else argv)
    root = repo_root()
    if not args.manifest and not args.templates_dir:
        args.manifest = os.path.join(HERE, "suites", "default.yaml")
    manifest = {}
    if args.manifest:
        if not os.path.isfile(args.manifest):
            print("Нет манифеста: %s" % args.manifest, file=sys.stderr)
            return 2
        manifest = load_manifest(args.manifest)
    cfg = resolve_cases(manifest, root, args)
    cfg["no_checks"] = bool(args.no_checks)
    if not cfg["cases"]:
        print("Нет сценариев (проверьте --templates-dir / cases в манифесте).", file=sys.stderr)
        return 2
    if not cfg["templates_dir"] or not os.path.isdir(cfg["templates_dir"]):
        print("Нет каталога шаблонов: %s" % cfg["templates_dir"], file=sys.stderr)
        return 2
    os.makedirs(cfg["results_dir"], exist_ok=True)
    profile = args.profile or os.environ.get("LM_QA_USER_INSTALL") or os.path.join(
        HERE, ".profile"
    )
    if str(profile).startswith("file:"):
        print("LM_QA_USER_INSTALL — путь ФС, не file:// URL", file=sys.stderr)
        return 2
    profile = os.path.abspath(profile)
    xvfb = None
    office_proc = None
    try:
        xvfb = start_xvfb_if_needed()
        office_bin = find_office_bin(cfg["office"])
        office_python = find_office_python(office_bin)
        if not args.no_install:
            dest, n = install_macros_to_profile(profile)
            print("Макросы → %s (%d файлов)" % (dest, n))
        else:
            write_low_macro_security(profile)
        print("Офис: %s" % office_bin)
        print("Python UNO: %s" % office_python)
        print("Профиль: %s" % profile)
        if args.keep_office:
            office_proc, _cmd = start_office(
                office_bin, profile, host=cfg["host"], port=cfg["port"]
            )
            wait_port(cfg["host"], cfg["port"], timeout_sec=60)
        case_results = []
        for case in cfg["cases"]:
            cid = case.get("id")
            print(u"— %s" % cid)
            try:
                one, office_proc = run_one_case(
                    cfg,
                    case,
                    office_bin,
                    office_python,
                    profile,
                    keep_office=args.keep_office,
                    office_proc=office_proc if args.keep_office else None,
                )
            except Exception as exc:
                traceback.print_exc()
                one = {
                    "id": cid,
                    "template": case.get("template"),
                    "save_path": u"",
                    "ok": False,
                    "timeout": False,
                    "error": _as_text(exc),
                    "checks": [],
                    "log_excerpt": u"",
                }
            sm = summarize_case(one)
            print(u"  %s  checks %s/%s/%s" % (sm["status"], sm["pass"], sm["fail"], sm["skip"]))
            if one.get("error"):
                print(u"  %s" % one["error"])
            case_results.append(one)
        payload = {
            "started_at": now_iso(),
            "macro_version": read_macro_version(root),
            "office_bin": office_bin,
            "office_python": office_python,
            "profile": profile,
            "templates_dir": cfg["templates_dir"],
            "results_dir": cfg["results_dir"],
            "report_path": cfg["report_path"],
            "cases": case_results,
        }
        json_path, md_path = write_reports(cfg["results_dir"], payload)
        print("Отчёт: %s" % md_path)
        print("JSON:  %s" % json_path)
        office_fail = any(not c.get("ok") or c.get("timeout") for c in case_results)
        check_fail = any(
            ch.get("status") == "fail" for c in case_results for ch in (c.get("checks") or [])
        )
        if office_fail and not any(os.path.isfile(_as_text(c.get("save_path"))) for c in case_results):
            return 2
        return 1 if (office_fail or check_fail) else 0
    except RuntimeError as exc:
        print(_as_text(exc), file=sys.stderr)
        return 2
    finally:
        if office_proc is not None:
            stop_proc(office_proc)
        if xvfb is not None:
            stop_proc(xvfb)


if __name__ == "__main__":
    sys.exit(main())
