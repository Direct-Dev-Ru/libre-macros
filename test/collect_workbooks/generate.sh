#!/usr/bin/env bash
# Генерация ручных шаблонов (JSON в колонке C). См. docs/12_MANUAL_TESTING.md
# Включает сценарий прямого переноса: Способ_переноса=xml_excel_ods (сценарий 11).
# Источники: source_01..03, source_products, source_pivot, vlookup/*, csv/*.
# sources/large/source_server_logs.xlsx — 3×200k+ строк логов + справочник серверов (ВПР).
# set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$ROOT"
if [ -x "$ROOT/venv/bin/python" ]; then
    exec "$ROOT/venv/bin/python" test/collect_workbooks/generate_manual_workbooks.py "$@"
fi
exec python3 test/collect_workbooks/generate_manual_workbooks.py "$@"
