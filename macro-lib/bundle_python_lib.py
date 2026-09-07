#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Общая сборка плоских bundled-модулей для macro-lib/pythonpath."""

from __future__ import annotations

MACRO_VERSION = "3.10.690"
import argparse
import ast
import sys
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class BundleSpec:
    packages: tuple[str, ...]
    bundle_tag: str
    title: str
    build_cmd: str
    import_hint: str
    dist_names: dict[str, str] = field(default_factory=dict)


def _module_name(site_packages: Path, py_path: Path) -> str:
    rel = py_path.relative_to(site_packages)
    parts = list(rel.parts)
    if parts[-1] == "__init__.py":
        parts = parts[:-1]
    else:
        parts[-1] = parts[-1][:-3]
    return ".".join(parts)


def _is_package(py_path: Path) -> bool:
    return py_path.name == "__init__.py"


def quote_source(source: str) -> str:
    """Встроить исходник как raw-многострочный литерал Python."""
    for delim in ("r\"\"\"", "r'''"):
        quote = delim[1:]
        if quote not in source:
            return delim + "\n" + source + "\n" + quote
    return repr(source)


def collect_modules(site_packages: Path, packages: tuple[str, ...]) -> dict:
    modules = {}
    for package in packages:
        root = site_packages / package
        if not root.is_dir():
            raise FileNotFoundError("Не найден пакет: %s" % root)
        for py_path in sorted(root.rglob("*.py")):
            if "__pycache__" in py_path.parts:
                continue
            name = _module_name(site_packages, py_path)
            source = py_path.read_text(encoding="utf-8")
            ast.parse(source, filename=str(py_path))
            modules[name] = {
                "source": source,
                "is_package": _is_package(py_path),
            }
    return modules


def read_versions(
    site_packages: Path,
    packages: tuple[str, ...],
    dist_names: dict[str, str],
) -> str:
    parts = []
    for package in packages:
        dist_name = dist_names.get(package, package)
        dist_dirs = sorted(site_packages.glob("%s-*.dist-info" % dist_name))
        if not dist_dirs:
            parts.append("%s=?" % package)
            continue
        name = dist_dirs[-1].name
        version = name.split("-", 1)[1].rsplit(".dist-info", 1)[0]
        parts.append("%s %s" % (package, version))
    return ", ".join(parts)


def format_sources_block(packages: tuple[str, ...], modules: dict) -> str:
    lines = ["_PREFIXES = %r" % (packages,)]
    lines.append("_SOURCES = {")
    for name in sorted(modules):
        item = modules[name]
        lines.append("    %r: {" % name)
        lines.append("        %r: %s," % ("is_package", item["is_package"]))
        lines.append("        %r: %s," % ("source", quote_source(item["source"])))
        lines.append("    },")
    lines.append("}")
    lines.append("")
    return "\n".join(lines)


def render_bundle(spec: BundleSpec, versions: str, sources_block: str) -> str:
    header = '''\
# -*- coding: utf-8 -*-
"""
{title}

Версии: {versions}
Сборка: {build_cmd}

Перед import целевых пакетов выполните:
  {import_hint}
"""
from __future__ import annotations

import importlib.abc
import importlib.util
import sys

_FINDER = None
_BUNDLE_TAG = "{tag}"

'''.format(
        title=spec.title,
        versions=versions,
        build_cmd=spec.build_cmd,
        import_hint=spec.import_hint,
        tag=spec.bundle_tag,
    )
    footer = '''

class _BundledLoader(importlib.abc.Loader):
    def __init__(self, fullname, source, is_package):
        self.fullname = fullname
        self.source = source
        self.is_package = is_package

    def create_module(self, spec):
        return None

    def exec_module(self, module):
        module_path = self.fullname.replace(".", "/") + ".py"
        module.__file__ = "<%s:%s>" % (_BUNDLE_TAG, module_path)
        code = compile(
            self.source,
            module.__file__,
            "exec",
            dont_inherit=True,
        )
        exec(code, module.__dict__)
        if self.is_package:
            package_path = self.fullname.replace(".", "/")
            module.__path__ = ["<%s:%s>" % (_BUNDLE_TAG, package_path)]


class _BundledFinder(importlib.abc.MetaPathFinder):
    def __init__(self, prefixes):
        self._prefixes = prefixes

    def _owns(self, fullname):
        for prefix in self._prefixes:
            if fullname == prefix or fullname.startswith(prefix + "."):
                return True
        return False

    def find_spec(self, fullname, path, target=None):
        if not self._owns(fullname):
            return None
        item = _SOURCES.get(fullname)
        if item is None:
            return None
        loader = _BundledLoader(fullname, item["source"], item["is_package"])
        return importlib.util.spec_from_loader(
            fullname,
            loader,
            is_package=item["is_package"],
        )


def install():
    """Зарегистрировать bundled-пакеты в sys.meta_path."""
    global _FINDER
    if _FINDER is not None:
        return
    _FINDER = _BundledFinder(_PREFIXES)
    sys.meta_path.insert(0, _FINDER)


install()

'''.format(tag=spec.bundle_tag)
    return header + sources_block + footer


def generate_bundle(
    site_packages: Path,
    output_path: Path,
    spec: BundleSpec,
) -> tuple[Path, int, int]:
    modules = collect_modules(site_packages, spec.packages)
    versions = read_versions(site_packages, spec.packages, spec.dist_names)
    body = render_bundle(
        spec,
        versions,
        format_sources_block(spec.packages, modules),
    )
    ast.parse(body, filename=str(output_path))
    output_path.write_text(body, encoding="utf-8")
    return output_path, len(modules), len(body)


def main_cli(spec: BundleSpec, default_output: Path, description: str) -> int:
    macro_dir = Path(__file__).resolve().parent
    repo_root = macro_dir.parent
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--site-packages",
        type=Path,
        default=repo_root / ".venv" / "lib" / "python3.12" / "site-packages",
    )
    parser.add_argument("-o", "--output", type=Path, default=default_output)
    args = parser.parse_args()
    if not args.site_packages.is_dir():
        print("Нет site-packages: %s" % args.site_packages, file=sys.stderr)
        return 1
    try:
        out, module_count, size = generate_bundle(
            args.site_packages,
            args.output,
            spec,
        )
    except Exception as err:
        print("Ошибка: %s" % err, file=sys.stderr)
        return 1
    print("Записано: %s" % out)
    print("Модулей: %d, размер: %d Кб" % (module_count, size // 1024))
    return 0
