#!/usr/bin/env python3
"""Скачать все файлы из каталога репозитория AlterOffice (directory listing)."""

from __future__ import annotations

import argparse
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

DEFAULT_URL = "https://repo.alteroffice.ru/release/astra/1.8/x86_64/desktop/"

# Ссылки вида <a href="file.deb">…</a> или <a href="file%5F….deb">
_HREF_RE = re.compile(
    r"""href=["']([^"'#]+)["']""",
    re.IGNORECASE,
)


def list_files(url: str) -> list[str]:
    """Вернуть имена файлов (не каталогов) из HTML-индекса."""
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "libre-macros-download/1.0"},
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        html = resp.read().decode("utf-8", errors="replace")

    names: list[str] = []
    seen: set[str] = set()
    for raw in _HREF_RE.findall(html):
        href = urllib.parse.unquote(raw.strip())
        if not href or href in (".", "./", "../", "/"):
            continue
        # абсолютные/родительские ссылки пропускаем
        if href.startswith(("http://", "https://", "/")):
            continue
        if href.startswith("../"):
            continue
        # каталоги в индексе обычно заканчиваются на /
        if href.endswith("/"):
            continue
        # служебные страницы индекса
        base = os.path.basename(href.rstrip("/"))
        if not base or base.lower().startswith("index.html"):
            continue
        if base in seen:
            continue
        seen.add(base)
        names.append(base)
    return sorted(names)


def download_file(base_url: str, name: str, dest_dir: Path, force: bool) -> str:
    """Скачать один файл. Возвращает статус: ok / skip / fail."""
    dest = dest_dir / name
    if dest.exists() and dest.stat().st_size > 0 and not force:
        return "skip"

    # в URL оставляем percent-encoding для спецсимволов
    quoted = urllib.parse.quote(name, safe=".-_")
    file_url = urllib.parse.urljoin(base_url, quoted)
    tmp = dest.with_suffix(dest.suffix + ".part")

    req = urllib.request.Request(
        file_url,
        headers={"User-Agent": "libre-macros-download/1.0"},
    )
    try:
        with urllib.request.urlopen(req, timeout=300) as resp, open(tmp, "wb") as out:
            total = resp.headers.get("Content-Length")
            total_n = int(total) if total and total.isdigit() else None
            done = 0
            while True:
                chunk = resp.read(1024 * 256)
                if not chunk:
                    break
                out.write(chunk)
                done += len(chunk)
                if total_n:
                    pct = 100.0 * done / total_n
                    print(
                        f"\r  {name}: {done:,}/{total_n:,} ({pct:.1f}%)",
                        end="",
                        flush=True,
                    )
                else:
                    print(f"\r  {name}: {done:,} bytes", end="", flush=True)
        tmp.replace(dest)
        print()
        return "ok"
    except Exception:
        if tmp.exists():
            tmp.unlink(missing_ok=True)
        raise


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Скачать все файлы из directory listing AlterOffice.",
    )
    ap.add_argument(
        "--url",
        default=DEFAULT_URL,
        help=f"URL каталога (по умолчанию: {DEFAULT_URL})",
    )
    ap.add_argument(
        "-o",
        "--output",
        default="alteroffice-desktop",
        help="Каталог назначения (по умолчанию: ./alteroffice-desktop)",
    )
    ap.add_argument(
        "-f",
        "--force",
        action="store_true",
        help="Перекачать уже существующие файлы",
    )
    args = ap.parse_args()

    url = args.url if args.url.endswith("/") else args.url + "/"
    dest_dir = Path(args.output).resolve()
    dest_dir.mkdir(parents=True, exist_ok=True)

    print(f"Список файлов: {url}")
    try:
        names = list_files(url)
    except urllib.error.URLError as e:
        print(f"Ошибка чтения индекса: {e}", file=sys.stderr)
        return 1

    if not names:
        print("Файлы не найдены.", file=sys.stderr)
        return 1

    print(f"Найдено файлов: {len(names)}")
    print(f"Каталог: {dest_dir}")

    ok = skip = fail = 0
    for i, name in enumerate(names, 1):
        print(f"[{i}/{len(names)}] {name}")
        try:
            status = download_file(url, name, dest_dir, force=args.force)
            if status == "ok":
                ok += 1
            else:
                skip += 1
                print("  уже есть, пропуск")
        except Exception as e:
            fail += 1
            print(f"  ОШИБКА: {e}", file=sys.stderr)

    print(f"Готово: скачано={ok}, пропущено={skip}, ошибок={fail}")
    return 1 if fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
