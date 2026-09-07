#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конвертация Markdown из docs/ в Word (.docx) через pandoc.

Зависимости (один из вариантов):
  pip install -r docs/requirements-build.txt   # pandoc внутри (pypandoc-binary)
  sudo apt install pandoc                        # системный pandoc + pip install pypandoc

Примеры:
  python3 docs/build_docx.py --all
  python3 docs/build_docx.py 02_INSTALL_USER.md
  python3 docs/build_docx.py technical_documentation/01_TECHNICAL_REQUIREMENTS.md
  python3 docs/build_docx.py --all --output-dir /tmp/docs-docx
"""
from __future__ import print_function, unicode_literals

import argparse
import os
import shutil
import sys

DOCS_ROOT = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(DOCS_ROOT, "docx")

SKIP_DIR_NAMES = {".git", "__pycache__", "docx"}
SKIP_FILE_PREFIXES = (".~lock.",)


def _ensure_pandoc(use_system=False):
    """Вернуть модуль pypandoc с рабочим бинарником pandoc."""
    try:
        import pypandoc
    except ImportError:
        print(
            "Нужен pandoc: pip install -r docs/requirements-build.txt\n"
            "или: sudo apt install pandoc && pip install pypandoc",
            file=sys.stderr,
        )
        return None

    if use_system:
        if shutil.which("pandoc"):
            return pypandoc
        print("Флаг --system-pandoc: pandoc не найден в PATH.", file=sys.stderr)
        return None

    try:
        pypandoc.get_pandoc_path()
    except OSError:
        try:
            import pypandoc_binary  # noqa: F401
        except ImportError:
            pass
        try:
            pypandoc.download_pandoc()
        except Exception as exc:
            if shutil.which("pandoc"):
                return pypandoc
            print(
                "Не удалось найти pandoc.\n"
                "  pip install -r docs/requirements-build.txt\n"
                "  или: sudo apt install pandoc\n"
                "Ошибка: %s" % exc,
                file=sys.stderr,
            )
            return None
    return pypandoc


def _iter_markdown_files(docs_root):
    for root, dirs, files in os.walk(docs_root):
        dirs[:] = [
            d
            for d in dirs
            if d not in SKIP_DIR_NAMES and not d.startswith(".")
        ]
        for fn in files:
            if not fn.endswith(".md"):
                continue
            if fn.startswith(SKIP_FILE_PREFIXES):
                continue
            yield os.path.join(root, fn)


def _rel_doc_path(path, docs_root):
    path = os.path.abspath(path)
    docs_root = os.path.abspath(docs_root)
    if not path.startswith(docs_root + os.sep) and path != docs_root:
        raise ValueError("Файл вне каталога docs: %s" % path)
    return os.path.relpath(path, docs_root)


def _output_path(md_path, docs_root, out_dir):
    rel = _rel_doc_path(md_path, docs_root)
    base, _ext = os.path.splitext(rel)
    return os.path.join(out_dir, base + ".docx")


def _pandoc_extra_args(md_path, docs_root):
    md_dir = os.path.dirname(os.path.abspath(md_path))
    # figs/… в MD относительно docs/; локальные картинки — от каталога файла.
    resource_paths = [docs_root, md_dir]
    seen = set()
    ordered = []
    for p in resource_paths:
        p = os.path.abspath(p)
        if p not in seen:
            seen.add(p)
            ordered.append(p)
    return [
        "--from=markdown",
        "--to=docx",
        "--standalone",
        "--resource-path=%s" % os.pathsep.join(ordered),
    ]


def convert_file(pypandoc, md_path, docs_root, out_dir, dry_run=False):
    md_path = os.path.abspath(md_path)
    if not os.path.isfile(md_path):
        raise FileNotFoundError(md_path)

    out_path = _output_path(md_path, docs_root, out_dir)
    if dry_run:
        print("%s -> %s" % (_rel_doc_path(md_path, docs_root), out_path))
        return out_path

    out_parent = os.path.dirname(out_path)
    if out_parent and not os.path.isdir(out_parent):
        os.makedirs(out_parent)

    extra = _pandoc_extra_args(md_path, docs_root)
    pypandoc.convert_file(md_path, "docx", outputfile=out_path, extra_args=extra)
    return out_path


def _resolve_inputs(paths, docs_root, all_docs):
    if all_docs:
        return sorted(_iter_markdown_files(docs_root))

    if not paths:
        return []

    resolved = []
    for raw in paths:
        p = raw
        if not os.path.isabs(p):
            p = os.path.join(docs_root, p)
        p = os.path.abspath(p)
        if os.path.isdir(p):
            for md in _iter_markdown_files(p):
                resolved.append(md)
        else:
            resolved.append(p)
    # сохранить порядок, убрать дубликаты
    out = []
    seen = set()
    for p in resolved:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Собрать Word (.docx) из Markdown в docs/ (pandoc)."
    )
    parser.add_argument(
        "paths",
        nargs="*",
        help="Файлы или каталоги .md (относительно docs/). Без аргументов — см. --all.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Конвертировать все .md в docs/ (кроме docx/ и служебных каталогов).",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUT,
        help="Каталог для .docx (по умолчанию: docs/docx/).",
    )
    parser.add_argument(
        "--docs-root",
        default=DOCS_ROOT,
        help="Корень документации (по умолчанию: каталог этого скрипта).",
    )
    parser.add_argument(
        "--system-pandoc",
        action="store_true",
        help="Использовать pandoc из PATH, не bundled.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Только показать пары вход → выход.",
    )
    args = parser.parse_args(argv)

    docs_root = os.path.abspath(args.docs_root)
    out_dir = os.path.abspath(args.output_dir)

    inputs = _resolve_inputs(args.paths, docs_root, args.all)
    if not inputs:
        parser.error("Укажите файлы или --all.")

    pypandoc = None if args.dry_run else _ensure_pandoc(use_system=args.system_pandoc)
    if not args.dry_run and pypandoc is None:
        return 1

    ok = 0
    failed = 0
    for md in inputs:
        rel = _rel_doc_path(md, docs_root)
        try:
            out = convert_file(
                pypandoc, md, docs_root, out_dir, dry_run=args.dry_run
            )
            if args.dry_run:
                ok += 1
            else:
                size = os.path.getsize(out)
                print("OK  %s  (%d bytes)" % (rel, size))
                ok += 1
        except Exception as exc:
            print("ERR %s: %s" % (rel, exc), file=sys.stderr)
            failed += 1

    if failed:
        print(
            "Готово: %d успешно, %d ошибок." % (ok, failed),
            file=sys.stderr,
        )
        return 1
    print("Готово: %d файл(ов)." % ok)
    return 0


if __name__ == "__main__":
    sys.exit(main())
