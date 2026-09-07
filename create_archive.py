#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для создания zip-архива со структурой:
    macro-lib/
        ├── (все файлы из macro-lib, включая collect_workbooks_no_docstrings.py)
        └── test/
            └── (все файлы из test)
"""

import subprocess
import sys
import zipfile
from pathlib import Path


def add_directory_to_zip(zipf, source_dir, arcname_prefix, exclude_names=None):
    """Добавить все файлы из директории в архив с префиксом пути."""
    exclude_names = set(exclude_names or ())
    source_path = Path(source_dir)
    for file_path in source_path.rglob("*"):
        if file_path.is_file():
            if file_path.name in exclude_names:
                continue
            if "__pycache__" in file_path.parts:
                continue
            rel_path = file_path.relative_to(source_path)
            arcname = str(Path(arcname_prefix) / rel_path)
            zipf.write(file_path, arcname)
            print(f"  + {arcname}")


def _run_macro_generator(macro_lib_dir, script_name, output_name, label):
    """Запустить generate_*.py в macro-lib и проверить выходной файл."""
    generator = macro_lib_dir / script_name
    if not generator.is_file():
        print("Предупреждение: нет %s, пропуск %s" % (generator, label))
        return None
    print("Сборка macro-lib/%s ..." % output_name)
    proc = subprocess.run(
        [sys.executable, str(generator)],
        cwd=str(macro_lib_dir.parent),
        capture_output=True,
        text=True,
    )
    if proc.stdout:
        print(proc.stdout.rstrip())
    if proc.returncode != 0:
        print(proc.stderr or proc.stdout, file=sys.stderr)
        raise RuntimeError("%s завершился с ошибкой" % script_name)
    out = macro_lib_dir / output_name
    if not out.is_file():
        raise RuntimeError("Не создан файл: %s" % out)
    return out


def build_derived_macro_versions(macro_lib_dir):
    """Собрать производные версии макроса (без docstring, для Python 2.7)."""
    _run_macro_generator(
        macro_lib_dir,
        "generate_no_docstrings.py",
        "collect_workbooks_no_docstrings.py",
        "без docstring",
    )
    _run_macro_generator(
        macro_lib_dir,
        "generate_py2_version.py",
        "collect_workbooks_py2.py",
        "Python 2.7",
    )


def create_archive():
    """Создать zip-архив с macro-lib и вложенным test."""
    base_dir = Path(__file__).resolve().parent
    macro_lib_dir = base_dir / "macro-lib"
    test_dir = base_dir / "test"
    output_zip = base_dir / "macro-lib-with-test.zip"

    if not macro_lib_dir.is_dir():
        print(f"Ошибка: директория не найдена: {macro_lib_dir}")
        return 1

    if not test_dir.is_dir():
        print(f"Ошибка: директория не найдена: {test_dir}")
        return 1

    build_derived_macro_versions(macro_lib_dir)

    print(f"Создание архива: {output_zip}")
    print(f"  - Источник macro-lib: {macro_lib_dir}")
    print(f"  - Источник test: {test_dir}")
    print()

    with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
        print("Добавление macro-lib/...")
        add_directory_to_zip(zipf, macro_lib_dir, "macro-lib")

        print("\nДобавление macro-lib/test/...")
        add_directory_to_zip(zipf, test_dir, "macro-lib/test")

    print(f"\n✓ Архив создан: {output_zip}")

    print("\nСтруктура архива:")
    with zipfile.ZipFile(output_zip, "r") as zipf:
        names = zipf.namelist()
        for name in names[:20]:
            print(f"  {name}")
        if len(names) > 20:
            print(f"  ... и еще {len(names) - 20} файлов")
        for derived in (
            "macro-lib/collect_workbooks_no_docstrings.py",
            "macro-lib/collect_workbooks_py2.py",
        ):
            print("  %s: %s" % (derived, "да" if derived in names else "НЕТ"))
    return 0


if __name__ == "__main__":
    sys.exit(create_archive() or 0)
