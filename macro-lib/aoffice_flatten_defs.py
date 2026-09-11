# -*- coding: utf-8 -*-
"""
Свёртка многострочных сигнатур def/async def в одну строку.

AlterOffice 2026 (pythonscript / getModuleByUrl) ломает исходник Scripts/*.py:
если заголовок функции разнесён на несколько строк, строка «):» без отступа
сбрасывает режим «внутри функции» и тело/следующий def склеиваются.

Также опасны многострочные «from x import (» — см. aoffice_compat_check.py
и .cursor/rules/alteroffice-2026-pythonscript.mdc.

Использование:
  python3 aoffice_flatten_defs.py macro-lib/file_list_macro.py
  python3 aoffice_flatten_defs.py --in-place macro-lib/*.py
  python3 aoffice_flatten_defs.py --tree macro-lib
"""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.725"
import argparse
import io
import os
import py_compile
import re
import sys
import tokenize

try:
    unicode
except NameError:
    unicode = str


def _collapse_signature_lines(lines, start, end):
    """start/end — 0-based индексы строк сигнатуры (включительно)."""
    first = lines[start]
    indent = first[: len(first) - len(first.lstrip())]
    chunks = []
    li = start
    while li <= end:
        chunks.append(lines[li].strip())
        li += 1
    merged = u" ".join(x for x in chunks if x)
    merged = re.sub(r",\s*\)", u")", merged)
    if not merged.endswith(u"\n"):
        merged += u"\n"
    if not merged.startswith(indent):
        merged = indent + merged.lstrip()
    return merged


def _find_multiline_def_spans(source):
    """Вернуть список (start_line, end_line) 0-based для многострочных сигнатур."""
    lines = source.splitlines(keepends=True)
    if not lines:
        return []
    try:
        tokens = list(tokenize.generate_tokens(io.StringIO(source).readline))
    except tokenize.TokenError:
        return []

    spans = []
    i = 0
    n = len(tokens)
    while i < n:
        async_line = None
        if tokens[i].type == tokenize.NAME and tokens[i].string == u"async":
            if i + 1 < n and tokens[i + 1].string == u"def":
                async_line = tokens[i].start[0] - 1
                i += 2
            else:
                i += 1
                continue
        elif tokens[i].type == tokenize.NAME and tokens[i].string == u"def":
            i += 1
        else:
            i += 1
            continue

        def_line = async_line if async_line is not None else tokens[i - 1].start[0] - 1
        # пропуск имени функции и декораторов уже позади; i на NAME после def
        while i < n and tokens[i].string != u"(":
            i += 1
        if i >= n:
            break

        depth = 0
        sig_end_line = None
        while i < n:
            t = tokens[i]
            if t.string == u"(":
                depth += 1
            elif t.string == u")":
                depth -= 1
                if depth == 0:
                    j = i + 1
                    while j < n and tokens[j].type in (
                        tokenize.NL,
                        tokenize.NEWLINE,
                        tokenize.INDENT,
                        tokenize.DEDENT,
                        tokenize.COMMENT,
                    ):
                        j += 1
                    if j < n and tokens[j].string == u"->":
                        while j < n and tokens[j].string != u":":
                            j += 1
                    if j < n and tokens[j].string == u":":
                        sig_end_line = tokens[j].start[0] - 1
                    i = j
                    break
            i += 1

        if sig_end_line is not None and sig_end_line > def_line:
            spans.append((def_line, sig_end_line))
        i += 1
    return spans


def polish_def_trailing_commas(source):
    """Убрать висячую запятую перед ) в однострочных def (после свёртки)."""
    out = []
    for line in source.splitlines(keepends=True):
        stripped = line.lstrip()
        if stripped.startswith(u"def ") or stripped.startswith(u"async def "):
            line = re.sub(r",\s*\)\s*:", u"):", line)
            line = re.sub(r",\s*\)\s*->", u") ->", line)
        out.append(line)
    return u"".join(out)


def flatten_multiline_defs(source):
    lines = source.splitlines(keepends=True)
    spans = _find_multiline_def_spans(source)
    if not spans:
        polished = polish_def_trailing_commas(source)
        return polished, 0
    for start, end in sorted(spans, reverse=True):
        lines[start : end + 1] = [_collapse_signature_lines(lines, start, end)]
    result = polish_def_trailing_commas(u"".join(lines))
    return result, len(spans)


def process_file(path, in_place=False, check_only=False, polish_only=False):
    path = os.path.abspath(path)
    with open(path, "rb") as f:
        raw = f.read()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("utf-8", errors="replace")

    if polish_only:
        new_text = polish_def_trailing_commas(text)
        changed = new_text != text
        count = 1 if changed else 0
    else:
        new_text, count = flatten_multiline_defs(text)
        changed = count > 0 or new_text != text

    if not changed:
        return 0, False

    if check_only:
        return count, True

    if in_place:
        with open(path, "wb") as f:
            f.write(new_text.encode("utf-8"))
        try:
            py_compile.compile(path, doraise=True)
        except py_compile.PyCompileError as err:
            raise SystemExit(u"После свёртки синтаксическая ошибка в %s:\n%s" % (path, err))
    return count, True


SKIP_DIR_NAMES = frozenset([u"py2", u"__pycache__", u".git"])
SKIP_FILE_NAMES = frozenset(
    [
        u"aoffice_flatten_defs.py",
        u"sync_macro_version.py",
        u"bundle_odfpy.py",
        u"bundle_openpyxl.py",
        u"bundle_python_lib.py",
        u"collect_workbooks_py2.py",
    ]
)
SKIP_FILE_PREFIXES = (u"test_", u"generate_")


def _should_process_file(path):
    name = os.path.basename(path)
    if name in SKIP_FILE_NAMES:
        return False
    if name.startswith(SKIP_FILE_PREFIXES):
        return False
    return name.endswith(u".py")


def iter_py_files(tree, skip_dirs=None):
    skip_dirs = set(skip_dirs or ())
    skip_dirs.update(SKIP_DIR_NAMES)
    for root, dirs, files in os.walk(tree):
        dirs[:] = [d for d in dirs if d not in skip_dirs]
        for name in files:
            fp = os.path.join(root, name)
            if _should_process_file(fp):
                yield fp


def main(argv=None):
    parser = argparse.ArgumentParser(
        description=u"Свернуть многострочные def/async def (совместимость AlterOffice)."
    )
    parser.add_argument(u"paths", nargs=u"*", help=u"Файлы .py или каталоги")
    parser.add_argument(
        u"--tree",
        metavar=u"DIR",
        help=u"Обработать все .py в каталоге (кроме py2/)",
    )
    parser.add_argument(
        u"--in-place",
        action=u"store_true",
        help=u"Перезаписать файлы на месте",
    )
    parser.add_argument(
        u"--polish-only",
        action=u"store_true",
        help=u"Только убрать висячие запятые перед ) в def (без свёртки)",
    )
    parser.add_argument(
        u"--check",
        action=u"store_true",
        help=u"Только проверить, есть ли многострочные сигнатуры",
    )
    args = parser.parse_args(argv)

    targets = list(args.paths or [])
    if args.tree:
        targets.append(args.tree)
    if not targets:
        parser.error(u"укажите файлы или --tree DIR")

    if not args.in_place and not args.check:
        parser.error(u"нужен --in-place или --check")
    if args.polish_only and args.check:
        parser.error(u"--polish-only несовместим с --check")

    total_files = 0
    total_spans = 0
    for path in targets:
        if os.path.isdir(path):
            files = list(iter_py_files(path))
        else:
            files = [path]
        for fp in files:
            count, changed = process_file(
                fp,
                in_place=args.in_place,
                check_only=args.check,
                polish_only=args.polish_only,
            )
            if changed:
                total_files += 1
                total_spans += count
                action = u"CHECK" if args.check else u"OK"
                print(u"[%s] %s: %d сигнатур" % (action, fp, count))

    if args.check and total_spans:
        print(
            u"\nНайдено многострочных сигнатур: %d в %d файлах"
            % (total_spans, total_files),
            file=sys.stderr,
        )
        return 1
    if args.in_place:
        print(u"Готово: %d сигнатур в %d файлах" % (total_spans, total_files))
    return 0


if __name__ == u"__main__":
    sys.exit(main())
