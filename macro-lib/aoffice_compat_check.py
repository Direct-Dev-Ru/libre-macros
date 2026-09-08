# -*- coding: utf-8 -*-
"""
Проверка совместимости Scripts/python/*.py с фильтром AlterOffice 2026.

AO перед compile выкидывает всё, кроме import/from и тел def/class
(см. test/pythonscript.py → getModuleByUrl). Симулируем фильтр и exec.

Дополнительно (--warn-assigns): предупреждения, если тело функций ссылается
на модульные присваивания скрипта — они исчезнут при загрузке из меню.

Вне LibreOffice (venv без `uno`) подставляются лёгкие stub-модули
uno / unohelper / com.sun.star.* — достаточно для import-time проверки.

Использование:
  PYTHONPATH=macro-lib/pythonpath python3 macro-lib/aoffice_compat_check.py
  python3 macro-lib/aoffice_compat_check.py --tree macro-lib --warn-assigns
"""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.710"
import argparse
import ast
import importlib.abc
import importlib.machinery
import os
import re
import sys
import types

try:
    unicode
except NameError:
    unicode = str


SKIP_FILE_NAMES = frozenset(
    [
        u"aoffice_flatten_defs.py",
        u"aoffice_compat_check.py",
        u"sync_macro_version.py",
        u"bundle_odfpy.py",
        u"bundle_openpyxl.py",
        u"bundle_python_lib.py",
        u"collect_workbooks_py2.py",
        u"generate_py2_version.py",
        u"generate_no_docstrings.py",
        u"collect_workbooks_no_docstrings.py",
        u"_extract_aoffice_cfg.py",
        u"_run_collect_extract.py",
        u"_fix_cw_get.py",
    ]
)

# Метаданные для сканеров (context_menu читает AST с диска) — ок, что вырежутся при exec.
HARMLESS_STRIPPED = frozenset(
    [
        u"MACRO_VERSION",
        u"g_exportedScripts",
        u"CONTEXT_CELL_MENU",
        u"CONTEXT_CELL_MENU_KEY",
        u"CONTEXT_TOOLS_MENU",
        u"CONTEXT_TOOLS_MENU_KEY",
    ]
)

_UNO_STUBS_READY = False


class _UnoAttrStub(object):
    """Любой атрибут/вызов — ещё один stub (PropertyValue, CENTER, …)."""

    def __init__(self, name=u"Stub"):
        self._name = name

    def __getattr__(self, item):
        if item.startswith(u"__") and item.endswith(u"__"):
            raise AttributeError(item)
        return _UnoAttrStub(u"%s.%s" % (self._name, item))

    def __call__(self, *args, **kwargs):
        return _UnoAttrStub(self._name)

    def __iter__(self):
        return iter(())

    def __bool__(self):
        return True

    __nonzero__ = __bool__


class _UnoStubLoader(importlib.abc.Loader):
    def create_module(self, spec):
        return sys.modules.get(spec.name)

    def exec_module(self, module):
        return None


def _make_uno_stub_module(fullname, is_pkg=False):
    mod = types.ModuleType(fullname)
    mod.__file__ = u"<aoffice-uno-stub>"
    if is_pkg:
        mod.__path__ = []
        mod.__package__ = fullname
    else:
        mod.__package__ = fullname.rpartition(u".")[0]

    def _getattr(name):
        if name.startswith(u"__") and name.endswith(u"__"):
            raise AttributeError(name)
        return _UnoAttrStub(u"%s.%s" % (fullname, name))

    mod.__getattr__ = _getattr  # type: ignore[attr-defined]
    return mod


class _UnoStubFinder(importlib.abc.MetaPathFinder):
    """uno / unohelper / com.* — stub, если модуль ещё не в sys.modules."""

    def find_spec(self, fullname, path, target=None):
        if fullname not in (u"uno", u"unohelper") and fullname != u"com" and (
            not fullname.startswith(u"com.")
        ):
            return None
        if fullname in sys.modules:
            return importlib.machinery.ModuleSpec(
                fullname,
                _UnoStubLoader(),
                is_package=hasattr(sys.modules[fullname], u"__path__"),
                origin=u"<aoffice-uno-stub>",
            )
        is_pkg = fullname == u"com" or fullname.startswith(u"com.")
        mod = _make_uno_stub_module(fullname, is_pkg=is_pkg)
        if fullname == u"uno":
            mod.fileUrlToSystemPath = lambda u: u  # type: ignore[attr-defined]
            mod.systemPathToFileUrl = lambda p: p  # type: ignore[attr-defined]
            mod.getComponentContext = lambda: _UnoAttrStub(u"getComponentContext")  # type: ignore[attr-defined]
        if fullname == u"unohelper":

            class Base(object):
                pass

            mod.Base = Base  # type: ignore[attr-defined]
        sys.modules[fullname] = mod
        return importlib.machinery.ModuleSpec(
            fullname,
            _UnoStubLoader(),
            is_package=is_pkg,
            origin=u"<aoffice-uno-stub>",
        )


def ensure_uno_stubs():
    """Подставить stub-модули, если `import uno` недоступен (venv без LO)."""
    global _UNO_STUBS_READY
    if _UNO_STUBS_READY:
        return
    try:
        import uno  # noqa: F401
        import unohelper  # noqa: F401

        _UNO_STUBS_READY = True
        return
    except ImportError:
        pass
    here = os.path.dirname(os.path.abspath(__file__))
    mocks = os.path.abspath(os.path.join(here, u"..", u"mocks"))
    if os.path.isdir(mocks) and mocks not in sys.path:
        sys.path.insert(0, mocks)
    # Повторная попытка: mocks/unohelper.py есть, uno — обычно нет.
    try:
        import uno  # noqa: F401
        import unohelper  # noqa: F401

        _UNO_STUBS_READY = True
        return
    except ImportError:
        pass
    already = any(isinstance(f, _UnoStubFinder) for f in sys.meta_path)
    if not already:
        sys.meta_path.insert(0, _UnoStubFinder())
    import uno  # noqa: F401
    import unohelper  # noqa: F401

    _UNO_STUBS_READY = True


def ao_filter_source(src):
    """Копия логики AlterOffice 2026 getModuleByUrl (построчный фильтр)."""
    out = []
    is_inside = False
    for line in src.splitlines(keepends=True):
        if line.startswith(("from", "import")):
            if u"IDE_utils" not in line:
                out.append(line)
        elif line.startswith(("def", "class")):
            is_inside = True
            out.append(line)
        elif is_inside and line.startswith(u" "):
            out.append(line)
        elif line.strip():
            is_inside = False
    return u"".join(out)


def find_exported_names(src):
    m = re.search(r"g_exportedScripts\s*=\s*\(([^)]*)\)", src, re.S)
    if not m:
        return []
    return re.findall(r"([A-Za-z_][A-Za-z0-9_]*)", m.group(1))


def module_level_assigns(tree):
    names = set()
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name):
                    names.add(t.id)
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            names.add(node.target.id)
    return names


def stripped_refs_in_functions(tree):
    """Имена модульных присваиваний, на которые ссылаются тела/defaults def."""
    assigns = module_level_assigns(tree) - HARMLESS_STRIPPED
    used = set()
    for node in tree.body:
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            continue
        for n in ast.walk(node):
            if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load) and n.id in assigns:
                used.add(n.id)
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            defaults = list(node.args.defaults)
            kw = getattr(node.args, u"kw_defaults", None) or []
            for d in defaults + list(kw):
                if d is None:
                    continue
                for n in ast.walk(d):
                    if isinstance(n, ast.Name) and n.id in assigns:
                        used.add(n.id)
    return sorted(used)


def check_file(path, pythonpath):
    path = os.path.abspath(path)
    with open(path, "rb") as f:
        raw = f.read()
    try:
        src = raw.decode("utf-8")
    except UnicodeDecodeError:
        src = raw.decode("utf-8", errors="replace")

    filt = ao_filter_source(src)
    exports = find_exported_names(src)
    try:
        code = compile(filt, path, "exec")
    except SyntaxError as err:
        return False, u"SyntaxError line %s: %s" % (err.lineno, err.msg), exports, [], []

    if pythonpath and pythonpath not in sys.path:
        sys.path.insert(0, pythonpath)

    ensure_uno_stubs()

    ns = {u"__name__": u"ooo_script_framework"}
    try:
        exec(code, ns)
    except Exception as err:
        missing = [e for e in exports if e not in ns]
        return (
            False,
            u"%s: %s" % (type(err).__name__, err),
            exports,
            missing,
            [],
        )

    missing = [e for e in exports if e not in ns]
    warn_names = []
    try:
        warn_names = stripped_refs_in_functions(ast.parse(src))
    except SyntaxError:
        pass

    if missing:
        return False, u"exports missing after exec", exports, missing, warn_names
    return True, None, exports, [], warn_names


def iter_root_scripts(tree):
    for name in sorted(os.listdir(tree)):
        if not name.endswith(u".py"):
            continue
        if name in SKIP_FILE_NAMES:
            continue
        yield os.path.join(tree, name)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description=u"Проверка Scripts/*.py против фильтра AlterOffice 2026"
    )
    parser.add_argument(
        u"--tree",
        default=None,
        help=u"Каталог macro-lib (по умолчанию рядом с этим скриптом)",
    )
    parser.add_argument(
        u"paths",
        nargs=u"*",
        help=u"Конкретные .py (иначе все в --tree)",
    )
    parser.add_argument(
        u"--warn-assigns",
        action=u"store_true",
        help=u"Предупреждать о ссылках на модульные константы в телах def",
    )
    parser.add_argument(
        u"--fail-on-warn",
        action=u"store_true",
        help=u"Считать предупреждения --warn-assigns ошибкой (код выхода 2)",
    )
    args = parser.parse_args(argv)

    here = os.path.dirname(os.path.abspath(__file__))
    tree = os.path.abspath(args.tree or here)
    pythonpath = os.path.join(tree, u"pythonpath")
    if os.path.isdir(pythonpath) and pythonpath not in sys.path:
        sys.path.insert(0, pythonpath)

    ensure_uno_stubs()

    if args.paths:
        files = [os.path.abspath(p) for p in args.paths]
    else:
        files = list(iter_root_scripts(tree))

    failed = 0
    warned = 0
    for fp in files:
        ok, err, exports, missing, warn_names = check_file(fp, pythonpath)
        label = os.path.basename(fp)
        if ok:
            print(u"[OK] %s exports=%s" % (label, exports or u"-"))
        else:
            failed += 1
            print(u"[FAIL] %s: %s" % (label, err))
            if missing:
                print(u"       missing: %s" % missing)
        if args.warn_assigns and warn_names:
            warned += 1
            sample = u", ".join(warn_names[:12])
            more = u"" if len(warn_names) <= 12 else u" …+%d" % (len(warn_names) - 12)
            print(
                u"[WARN] %s: %d имён вырежутся фильтром, но используются в def: %s%s"
                % (label, len(warn_names), sample, more)
            )

    if failed:
        print(u"\nОшибок: %d. См. docs и .cursor/rules про AlterOffice 2026." % failed)
        return 1
    if args.fail_on_warn and warned:
        print(u"\nПредупреждений: %d (режим --fail-on-warn)." % warned)
        return 2
    print(u"\nВсе проверенные скрипты совместимы с фильтром AO 2026 (загрузка/export).")
    if args.warn_assigns and warned:
        print(
            u"Предупреждений о константах: %d — при запуске из меню возможен NameError."
            % warned
        )
    return 0


if __name__ == u"__main__":
    sys.exit(main())
