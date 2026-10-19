#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Упаковка выбранных файлов и каталогов проекта в zip-архив по списку.

Список (по умолчанию `.archivelist` в корне проекта, параметр `-l/--list`) —
текстовый файл UTF-8:

    # комментарий, пустые строки игнорируются
    out: macro-lib-with-test.zip      # имя архива по умолчанию (необязательно)
    macro-lib                         # каталог целиком (рекурсивно)
    test -> macro-lib/test            # каталог/файл с переносом внутрь архива
    README.md                         # отдельный файл

Пути в списке — относительно корня проекта (каталог со скриптом). Выход за
пределы проекта запрещён. Необязательный маппинг `<путь> -> <путь в архиве>`
позволяет вкладывать элементы друг в друга.

Куда складывать архив:
  * по умолчанию — корень проекта;
  * `-o/--output <каталог>` — другой каталог (создаётся при необходимости);
  * `-o/--output <...>/имя.zip` — полный путь архива (имя берётся из него);
  * `-n/--name <имя.zip>` — только имя файла архива.

Служебные каталоги (`.git`, `__pycache__`, `venv`, `.venv`, `.idea`, `.vscode`,
`.gigaide`, `.cursor`, `.openpyxl_cache`) и мусор (`*.pyc`, `*.pyo`,
`.DS_Store`, lock-файлы LO) в архив не попадают; сам архив — тоже.

Примеры:
    python3 create_archive.py
    python3 create_archive.py -o ~/tmp/release
    python3 create_archive.py -l release.list -n macros.zip
    python3 create_archive.py --dry-run -v
    python3 create_archive.py --init          # создать пример .archivelist
"""

import argparse
import fnmatch
import sys
import zipfile
from collections import namedtuple
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_LIST_NAME = ".archivelist"
DEFAULT_ARCHIVE_NAME = "project-archive.zip"

# Каталоги, которые никогда не попадают в архив (VCS, кэши, окружения, IDE).
SKIP_DIR_NAMES = frozenset(
    (
        ".git",
        "__pycache__",
        "venv",
        ".venv",
        ".idea",
        ".vscode",
        ".gigaide",
        ".cursor",
        ".openpyxl_cache",
    )
)
# Файлы, которые никогда не попадают в архив.
SKIP_FILE_GLOBS = ("*.pyc", "*.pyo", ".DS_Store", ".~lock.*", "~$*")

OUT_DIRECTIVE = "out:"
MAPPING_SEP = "->"

EXAMPLE_LIST = """\
# Список содержимого архива для create_archive.py.
# Пустые строки и строки, начинающиеся с '#', игнорируются.
# Строка "out: <имя.zip>" задаёт имя архива по умолчанию
# (переопределяется параметрами --name и --output).
# Остальные строки — путь относительно корня проекта (файл или каталог);
# необязательный маппинг "<путь> -> <путь в архиве>" кладёт элемент вглубь.

out: macro-lib-with-test.zip

# Макросы и pythonpath-библиотеки
macro-lib

# Шаблоны и материалы ручной проверки — вкладываются в macro-lib/test
test -> macro-lib/test
"""

ArchiveEntry = namedtuple("ArchiveEntry", ["source", "arcname", "line_no"])


class ArchiveError(Exception):
    """Ошибка подготовки архива (сообщение показывается пользователю)."""


def _entry_label(list_path, line_no):
    return "%s:%d" % (list_path, line_no)


def parse_archive_list(list_path):
    """Прочитать список. Вернуть (entries, archive_name_или_None)."""
    try:
        text = list_path.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        raise ArchiveError(
            "Список не найден: %s\nСоздайте его командой: %s --init"
            % (list_path, Path(__file__).name)
        )
    except OSError as exc:
        raise ArchiveError("Не удалось прочитать %s: %s" % (list_path, exc))

    entries = []
    archive_name = None
    for line_no, raw_line in enumerate(text.splitlines(), 1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        if line.lower().startswith(OUT_DIRECTIVE):
            value = line[len(OUT_DIRECTIVE):].strip()
            if not value:
                raise ArchiveError("%s: пустое имя в 'out:'" % _entry_label(list_path, line_no))
            archive_name = value
            continue

        source = line
        arcname = None
        if MAPPING_SEP in line:
            source, _, arcname = (part.strip() for part in line.partition(MAPPING_SEP))
            if not source or not arcname:
                raise ArchiveError(
                    "%s: некорректный маппинг '<путь> -> <путь в архиве>': %s"
                    % (_entry_label(list_path, line_no), line)
                )
        entries.append(ArchiveEntry(source, arcname, line_no))

    return entries, archive_name


def _normalize_rel(raw_path, label, what):
    """Привести путь к виду posix и проверить, что он относительный и без '..'."""
    posix = str(raw_path).replace("\\", "/").strip().strip("/")
    parts = [part for part in posix.split("/") if part not in ("", ".")]
    if not parts or any(part == ".." for part in parts):
        raise ArchiveError("%s: %s должен быть относительным путём внутри проекта: %s"
                           % (label, what, raw_path))
    return "/".join(parts)


def resolve_entry_source(entry, list_path, project_root):
    """Проверить путь из списка и вернуть (абсолютный путь, путь в архиве)."""
    label = _entry_label(list_path, entry.line_no)
    raw = str(entry.source).replace("\\", "/").strip()
    if raw.startswith("/") or (len(raw) > 1 and raw[1] == ":"):
        raise ArchiveError("%s: абсолютный путь запрещён: %s" % (label, entry.source))

    rel_source = _normalize_rel(entry.source, label, "путь источника")
    resolved = (project_root / rel_source).resolve()
    if resolved != project_root and project_root not in resolved.parents:
        raise ArchiveError("%s: путь выходит за пределы проекта: %s" % (label, entry.source))
    if not resolved.exists():
        raise ArchiveError("%s: не найдено: %s" % (label, entry.source))

    arcname = _normalize_rel(entry.arcname or rel_source, label, "путь в архиве")
    return resolved, arcname


def _is_skipped(path, rel_parts):
    if any(part in SKIP_DIR_NAMES for part in rel_parts):
        return True
    return any(fnmatch.fnmatch(path.name, pattern) for pattern in SKIP_FILE_GLOBS)


def iter_entry_files(resolved, arcname):
    """Отдать пары (файл на диске, имя в архиве) для файла или каталога."""
    if resolved.is_file():
        yield resolved, arcname
        return

    for path in sorted(resolved.rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(resolved)
        if _is_skipped(path, rel.parts):
            continue
        yield path, "%s/%s" % (arcname, rel.as_posix())


def collect_payload(entries, list_path, project_root, skip_paths=()):
    """Собрать список (файл, имя в архиве) без дублей."""
    payload = []
    seen = {}
    for entry in entries:
        resolved, arcname = resolve_entry_source(entry, list_path, project_root)
        for path, name in iter_entry_files(resolved, arcname):
            if path in skip_paths:
                continue
            if name in seen:
                print("  ! дубль пропущен: %s (уже добавлен из %s)"
                      % (name, _entry_label(list_path, seen[name])), file=sys.stderr)
                continue
            seen[name] = entry.line_no
            payload.append((path, name))
    return payload


def resolve_output_path(output, name, list_archive_name):
    """Определить полный путь архива по --output/--name и значению 'out:'."""
    default_name = name or list_archive_name or DEFAULT_ARCHIVE_NAME
    if output is None:
        return PROJECT_ROOT / default_name

    out = Path(str(output)).expanduser()
    if not out.is_absolute():
        out = Path.cwd() / out

    if out.suffix.lower() == ".zip" and not out.is_dir():
        return out if not name else out.parent / name
    return out / default_name


def build_archive(payload, archive_path):
    """Записать payload в zip-архив."""
    archive_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for path, name in payload:
            zipf.write(path, name)


def format_size(num_bytes):
    for unit in ("Б", "КБ", "МБ", "ГБ"):
        if num_bytes < 1024 or unit == "ГБ":
            return "%.1f %s" % (num_bytes, unit)
        num_bytes /= 1024.0
    return "%.1f ГБ" % num_bytes


def print_summary(archive_path, payload):
    dirs = {}
    for _path, name in payload:
        top = name.split("/", 1)[0] if "/" in name else "(корень)"
        dirs[top] = dirs.get(top, 0) + 1

    print("\nСодержимое архива (файлов: %d):" % len(payload))
    for top in sorted(dirs):
        print("  %-24s %d" % (top, dirs[top]))
    size = archive_path.stat().st_size
    print("\n✓ Архив создан: %s (%s)" % (archive_path, format_size(size)))


def write_example_list(list_path, force=False):
    if list_path.exists() and not force:
        raise ArchiveError("Файл уже существует: %s (используйте --force)" % list_path)
    list_path.parent.mkdir(parents=True, exist_ok=True)
    list_path.write_text(EXAMPLE_LIST, encoding="utf-8")
    print("✓ Создан пример списка: %s" % list_path)
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Создать zip-архив файлов и каталогов проекта по списку (.archivelist)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=EXAMPLE_LIST,
    )
    parser.add_argument(
        "-l", "--list", dest="list_path", default=None,
        help="файл списка (по умолчанию: <%s> в корне проекта)" % DEFAULT_LIST_NAME,
    )
    parser.add_argument(
        "-o", "--output", default=None,
        help="куда положить архив: каталог либо путь к файлу .zip "
             "(по умолчанию: корень проекта)",
    )
    parser.add_argument(
        "-n", "--name", default=None,
        help="имя файла архива (по умолчанию: из 'out:' в списке, иначе %s)"
             % DEFAULT_ARCHIVE_NAME,
    )
    parser.add_argument("--init", action="store_true",
                        help="создать пример списка и выйти")
    parser.add_argument("--force", action="store_true",
                        help="перезаписать список при --init")
    parser.add_argument("--dry-run", action="store_true",
                        help="показать состав архива, ничего не записывая")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="печатать каждый добавляемый файл")
    parser.add_argument("-q", "--quiet", action="store_true",
                        help="только ошибки")
    args = parser.parse_args(argv)

    project_root = PROJECT_ROOT
    if args.list_path:
        list_path = Path(args.list_path).expanduser()
        if not list_path.is_absolute():
            list_path = Path.cwd() / list_path
    else:
        list_path = project_root / DEFAULT_LIST_NAME

    try:
        if args.init:
            return write_example_list(list_path, force=args.force)

        entries, list_archive_name = parse_archive_list(list_path)
        if not entries:
            raise ArchiveError("Список пуст (нет ни одного файла или каталога): %s" % list_path)

        archive_path = resolve_output_path(args.output, args.name, list_archive_name).resolve()

        if not args.quiet:
            print("Корень проекта:  %s" % project_root, flush=True)
            print("Список:          %s" % list_path, flush=True)
            print("Архив:           %s" % archive_path, flush=True)
            print(flush=True)

        payload = collect_payload(entries, list_path, project_root, skip_paths=(archive_path,))
        if not payload:
            raise ArchiveError("В архив нечего добавить: все элементы отфильтрованы")

        if args.verbose and not args.quiet:
            for path, name in payload:
                print("  + %s" % name)

        if args.dry_run:
            if not args.quiet:
                print("\n--dry-run: запись архива пропущена (файлов: %d)" % len(payload))
            return 0

        build_archive(payload, archive_path)
        if not args.quiet:
            print_summary(archive_path, payload)
        return 0
    except ArchiveError as exc:
        print("❌ %s" % exc, file=sys.stderr)
        return 1
    except OSError as exc:
        print("❌ Ошибка файловой системы: %s" % exc, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
