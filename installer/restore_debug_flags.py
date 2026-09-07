# -*- coding: utf-8 -*-
"""
Восстановить DEBUG-флаги в macro-lib после сборки MacroInstaller_bundled.ods.

Сборка (build_ods_bundled) перед упаковкой сбрасывает *_DEBUG / DEBUG_LOG и
ключи *DEBUG в dict literals в False и сохраняет снимок в .bundled_debug_snapshot.json.
Этот скрипт возвращает только те флаги, которые были True до сборки.
"""
import argparse
import os
import sys

from build_ods import DEFAULT_DEBUG_SNAPSHOT, restore_debug_flags_tree

DEFAULT_MACRO_LIB = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "macro-lib")
)


def main():
    parser = argparse.ArgumentParser(
        description="Восстановить DEBUG-флаги macro-lib после сборки bundled-инсталлятора",
    )
    parser.add_argument(
        "--macro-lib-dir",
        default=DEFAULT_MACRO_LIB,
        help=f"Каталог macro-lib (по умолчанию: {DEFAULT_MACRO_LIB})",
    )
    parser.add_argument(
        "--snapshot",
        default=DEFAULT_DEBUG_SNAPSHOT,
        help=f"JSON-снимок от build_ods_bundled (по умолчанию: {DEFAULT_DEBUG_SNAPSHOT})",
    )
    parser.add_argument(
        "--remove-snapshot",
        action="store_true",
        help="Удалить снимок после успешного восстановления",
    )
    args = parser.parse_args()

    try:
        restored = restore_debug_flags_tree(
            macro_lib_dir=args.macro_lib_dir,
            snapshot_path=args.snapshot,
            remove_snapshot=args.remove_snapshot,
        )
    except FileNotFoundError as exc:
        print(f"❌ {exc}")
        raise SystemExit(1)
    except (ValueError, OSError) as exc:
        print(f"❌ Ошибка: {exc}")
        raise SystemExit(1)

    if not restored:
        print("Нечего восстанавливать (снимок пуст или флаги уже True).")
        raise SystemExit(0)

    print(f"✅ DEBUG-флаги восстановлены в {len(restored)} файле(ах):")
    base = os.path.abspath(args.macro_lib_dir)
    for path in restored:
        print(f"   {os.path.relpath(path, base)}")
    raise SystemExit(0)


if __name__ == "__main__":
    main()
