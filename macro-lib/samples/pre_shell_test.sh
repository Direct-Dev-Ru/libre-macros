#!/bin/sh
# Тестовый pre-скрипт для макроса test_pre_shell.
set -eu
MARKER="${TMPDIR:-/tmp}/lm_pre_shell_test.marker"
echo "pre_shell_test: start pid=$$ LM_PRE_SHELL_TEST=${LM_PRE_SHELL_TEST:-}"
echo "pre_shell_test: cwd=$(pwd)"
date -Iseconds > "$MARKER" 2>/dev/null || date > "$MARKER"
echo "pre_shell_test: wrote $MARKER"
echo "pre_shell_test: ok"
exit 0
