#!/usr/bin/env bash
# Устарело: автотесты удалены. Генерация ручных шаблонов — generate.sh
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
exec "$ROOT/test/collect_workbooks/generate.sh" --skip-sources "$@"
