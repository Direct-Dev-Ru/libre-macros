# -*- coding: utf-8 -*-
"""Поиск ACell/soffice, изолированный профиль, Xvfb, UNO-сокет."""
from __future__ import print_function, unicode_literals

import os
import shutil
import socket
import subprocess
import sys
import time

try:
    unicode
except NameError:
    unicode = str

KNOWN_OFFICE_BINS = (
    "acell",
    "soffice",
    "soffice.bin",
    "libreoffice",
)

KNOWN_OFFICE_PATHS = (
    "/usr/lib/alteroffice/program/acell",
    "/usr/lib/alteroffice/program/soffice",
    "/opt/alteroffice/program/acell",
    "/opt/alteroffice/program/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/usr/lib/libreoffice/program/soffice.bin",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
)


def repo_root():
    here = os.path.abspath(os.path.dirname(__file__))
    return os.path.abspath(os.path.join(here, "..", ".."))


def file_url_for_path(path):
    path = os.path.abspath(path)
    if os.name == "nt":
        drive, rest = os.path.splitdrive(path)
        rest = rest.replace("\\", "/")
        return "file:///" + drive.replace(":", "|") + rest
    return "file://" + path


def find_office_bin(prefer="auto"):
    env = os.environ.get("LM_QA_OFFICE") or ""
    if env.strip():
        if os.path.isfile(env) and os.access(env, os.X_OK):
            return os.path.abspath(env)
        found = shutil.which(env)
        if found:
            return os.path.abspath(found)
        raise RuntimeError("LM_QA_OFFICE не найден: %s" % env)
    names = list(KNOWN_OFFICE_BINS)
    prefer = (prefer or "auto").strip().lower()
    if prefer == "acell":
        names = ["acell"] + [n for n in names if n != "acell"]
    elif prefer in ("soffice", "libreoffice"):
        names = ["soffice", "soffice.bin", "libreoffice"] + [
            n for n in names if n not in ("soffice", "soffice.bin", "libreoffice")
        ]
    for name in names:
        found = shutil.which(name)
        if found:
            return os.path.abspath(found)
    for path in KNOWN_OFFICE_PATHS:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return os.path.abspath(path)
    raise RuntimeError(
        "Не найден ACell/LibreOffice (acell/soffice). Задайте LM_QA_OFFICE."
    )


def find_office_python(office_bin):
    env = os.environ.get("LM_QA_OFFICE_PYTHON") or ""
    if env.strip() and os.path.isfile(env):
        return os.path.abspath(env)
    program_dir = os.path.dirname(os.path.abspath(office_bin))
    candidates = [
        os.path.join(program_dir, "python"),
        os.path.join(program_dir, "python.exe"),
        os.path.join(os.path.dirname(program_dir), "program", "python"),
        "/usr/lib/libreoffice/program/python",
        "/usr/lib/alteroffice/program/python",
        "/opt/libreoffice/program/python",
    ]
    real = os.path.realpath(office_bin)
    real_dir = os.path.dirname(real)
    if real_dir not in (program_dir,):
        candidates.insert(0, os.path.join(real_dir, "python"))
    for path in candidates:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return path
    return sys.executable


def wait_port(host, port, timeout_sec=45):
    deadline = time.time() + float(timeout_sec)
    last_err = None
    while time.time() < deadline:
        try:
            sock = socket.create_connection((host, int(port)), timeout=1.0)
            sock.close()
            return True
        except Exception as err:
            last_err = err
            time.sleep(0.25)
    raise RuntimeError("UNO-сокет %s:%s не открылся: %s" % (host, port, last_err))


def port_free(host, port):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.settimeout(0.3)
        return sock.connect_ex((host, int(port))) != 0
    finally:
        try:
            sock.close()
        except Exception:
            pass


def start_xvfb_if_needed():
    if os.environ.get("LM_QA_DISPLAY"):
        os.environ["DISPLAY"] = os.environ["LM_QA_DISPLAY"]
        return None
    if os.environ.get("DISPLAY"):
        return None
    xvfb = shutil.which("Xvfb")
    if not xvfb:
        print(
            "Предупреждение: нет DISPLAY и нет Xvfb — AWT может зависнуть.",
            file=sys.stderr,
        )
        return None
    display = os.environ.get("LM_QA_XVFB_DISPLAY") or ":91"
    proc = subprocess.Popen(
        [xvfb, display, "-screen", "0", "1280x1024x24", "-nolisten", "tcp"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    os.environ["DISPLAY"] = display
    time.sleep(0.4)
    return proc


def write_low_macro_security(profile_dir):
    """Разрешить Python-макросы в свежем профиле (MacroSecurityLevel=0)."""
    user = os.path.join(profile_dir, "user")
    os.makedirs(user, exist_ok=True)
    path = os.path.join(user, "registrymodifications.xcu")
    xml = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                old = fh.read()
            if "MacroSecurityLevel" in old:
                return path
        except Exception:
            pass
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(xml)
    return path


def install_macros_to_profile(profile_dir, macro_lib=None):
    """Скопировать entry *.py + pythonpath/ в profile/user/Scripts/python (без вопросов)."""
    root = repo_root()
    if macro_lib is None:
        macro_lib = os.path.join(root, "macro-lib")
    dest = os.path.join(profile_dir, "user", "Scripts", "python")
    pp_dest = os.path.join(dest, "pythonpath")
    os.makedirs(pp_dest, exist_ok=True)
    skip = {
        "aoffice_flatten_defs.py",
        "aoffice_compat_check.py",
        "sync_macro_version.py",
        "bundle_odfpy.py",
        "bundle_openpyxl.py",
        "bundle_python_lib.py",
        "bundle_ldap3.py",
        "bundle_segno.py",
        "bundle_qr_template.py",
        "bundle_set_menu_images.py",
        "bundle_todo_task_help.py",
        "generate_no_docstrings.py",
        "generate_py2_version.py",
        "collect_workbooks_no_docstrings.py",
        "collect_workbooks_py2.py",
    }
    copied = 0
    for name in os.listdir(macro_lib):
        if not name.endswith(".py") or name in skip:
            continue
        if name.startswith("test_"):
            continue
        src = os.path.join(macro_lib, name)
        if not os.path.isfile(src):
            continue
        shutil.copy2(src, os.path.join(dest, name))
        copied += 1
    pp_src = os.path.join(macro_lib, "pythonpath")
    if os.path.isdir(pp_src):
        for name in os.listdir(pp_src):
            if not name.endswith(".py"):
                continue
            shutil.copy2(os.path.join(pp_src, name), os.path.join(pp_dest, name))
            copied += 1
    entry = os.path.join(dest, "qa_collect_run.py")
    if not os.path.isfile(entry):
        raise RuntimeError("После копирования нет qa_collect_run.py в %s" % dest)
    write_low_macro_security(profile_dir)
    return dest, copied


def start_office(office_bin, profile_dir, host="127.0.0.1", port=2002, extra_args=None):
    os.makedirs(profile_dir, exist_ok=True)
    accept = "socket,host=%s,port=%d;urp;StarOffice.ComponentContext" % (host, int(port))
    user_install = "-env:UserInstallation=%s" % file_url_for_path(profile_dir)
    cmd = [
        office_bin,
        "--headless",
        "--invisible",
        "--nologo",
        "--norestore",
        "--nofirststartwizard",
        "--nolockcheck",
        "--nodefault",
        user_install,
        "--accept=%s" % accept,
    ]
    if extra_args:
        cmd.extend(list(extra_args))
    env = os.environ.copy()
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        env=env,
        cwd=os.path.dirname(office_bin) or None,
    )
    return proc, cmd


def stop_proc(proc, timeout=8):
    if proc is None:
        return
    if proc.poll() is not None:
        return
    try:
        proc.terminate()
        proc.wait(timeout=timeout)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass
