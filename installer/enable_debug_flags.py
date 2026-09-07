# -*- coding: utf-8 -*-
"""
Принудительно включить все DEBUG-флаги в macro-lib (True), независимо от текущего значения.

Затрагивает module-level *_DEBUG / DEBUG_LOG и ключи *DEBUG в dict literals —
те же паттерны, что при сборке bundled-инсталлятора.
"""
import argparse
import os

from build_ods import enable_debug_flags_tree

DEFAULT_MACRO_LIB = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "macro-lib")
)


def main():
    parser = argparse.ArgumentParser(
        description="Принудительно включить DEBUG-флаги во всех .py macro-lib",
    )
    parser.add_argument(
        "--macro-lib-dir",
        default=DEFAULT_MACRO_LIB,
        help=f"Каталог macro-lib (по умолчанию: {DEFAULT_MACRO_LIB})",
    )
    args = parser.parse_args()

    try:
        updated = enable_debug_flags_tree(args.macro_lib_dir)
    except OSError as exc:
        print(f"❌ Ошибка: {exc}")
        raise SystemExit(1)

    if not updated:
        print("Изменений нет (все DEBUG-флаги уже True или не найдены).")
        raise SystemExit(0)

    print(f"✅ DEBUG-флаги включены в {len(updated)} файле(ах):")
    base = os.path.abspath(args.macro_lib_dir)
    for path in updated:
        print(f"   {os.path.relpath(path, base)}")
    raise SystemExit(0)


if __name__ == "__main__":
    main()
