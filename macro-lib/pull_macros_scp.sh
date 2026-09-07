#!/usr/bin/env bash
# Incremental pull of macro-lib into LibreOffice Scripts/python via scp + SHA-256.
#
# 1) Download checksums JSON from the dev machine
# 2) Compare local files; scp only mismatches / missing
# 3) Mirror Scripts/python into AlterOffice3 (same Roaming root, same tail)
#
#   ./pull_macros_scp.sh
#
# Environment:
#   REMOTE            — SSH host (default su@192.168.89.18)
#   REMOTE_SRC        — remote macro-lib directory
#   REMOTE_CHECKSUMS  — remote JSON with sha256
#   DEST              — LibreOffice Scripts/python
#   DEST_AO           — AlterOffice3 Scripts/python (optional; auto from DEST)

set -euo pipefail

REMOTE="${REMOTE:-su@192.168.89.18}"
REMOTE_SRC="${REMOTE_SRC:-/home/su/projects/python/libre-macros/macro-lib}"
REMOTE_CHECKSUMS="${REMOTE_CHECKSUMS:-/home/su/projects/python/libre-macros/scripts/macro_lib_py_checksums.json}"
DEST="${DEST:-/drives/c/Users/KuznectcovaEN_local/AppData/Roaming/LibreOffice/4/user/Scripts/python}"

alteroffice_dest() {
    local libre="$1"
    if [[ -n "${DEST_AO:-}" ]]; then
        printf '%s\n' "$DEST_AO"
        return
    fi
    # .../AppData/Roaming/LibreOffice/4/user/Scripts/python
    # → .../AppData/Roaming/AlterOffice3/user/Scripts/python
    if [[ "$libre" =~ [Aa]pp[Dd]ata[/\\][Rr]oaming ]]; then
        local prefix
        prefix="$(printf '%s\n' "$libre" | sed -E 's|(.*[Aa]pp[Dd]ata[/\\][Rr]oaming)[/\\].*|\1|')"
        printf '%s/AlterOffice3/user/Scripts/python\n' "$prefix" | sed 's|\\|/|g'
        return
    fi
    printf '%s\n' "$libre" | sed -E 's|LibreOffice[/\\][0-9]+|AlterOffice3|g'
}

is_installable() {
    local rel="$1"
    rel="${rel//\\//}"
    [[ -z "$rel" ]] && return 1
    [[ "$rel" == py2/* ]] && return 1
    [[ "$rel" == *__pycache__* ]] && return 1
    [[ "$rel" == pythonpath/* ]] && return 0
    [[ "$rel" == */* ]] && return 1
    [[ "$rel" == aoffice_* ]] && return 1
    [[ "$rel" == sync_macro_version.py ]] && return 1
    [[ "$rel" == *.py || "$rel" == *.txt ]] && return 0
    return 1
}

sha256_file() {
    local path="$1"
    if [[ ! -f "$path" ]]; then
        return 1
    fi
    if command -v sha256sum >/dev/null 2>&1; then
        sha256sum "$path" | awk '{print tolower($1)}'
    elif command -v shasum >/dev/null 2>&1; then
        shasum -a 256 "$path" | awk '{print tolower($1)}'
    else
        echo "need sha256sum or shasum" >&2
        return 1
    fi
}

copy_if_different() {
    local src="$1" dst="$2"
    [[ -f "$src" ]] || return 1
    mkdir -p "$(dirname "$dst")"
    local h1 h2
    h1="$(sha256_file "$src" || true)"
    h2="$(sha256_file "$dst" || true)"
    if [[ -n "$h1" && -n "$h2" && "$h1" == "$h2" ]]; then
        return 1
    fi
    cp -f "$src" "$dst"
    return 0
}

DEST_AO="$(alteroffice_dest "$DEST")"
mkdir -p "$DEST/pythonpath" "$DEST_AO/pythonpath"

TMP_JSON="$(mktemp "${TMPDIR:-/tmp}/macro_lib_py_checksums.XXXXXX.json")"
cleanup() { rm -f "$TMP_JSON"; }
trap cleanup EXIT

echo "Source:         $REMOTE:$REMOTE_SRC"
echo "Checksums:      $REMOTE:$REMOTE_CHECKSUMS"
echo "Destination LO: $DEST"
echo "Destination AO: $DEST_AO"
echo

echo "scp checksums JSON..."
scp "$REMOTE:$REMOTE_CHECKSUMS" "$TMP_JSON"

if ! command -v python3 >/dev/null 2>&1 && ! command -v python >/dev/null 2>&1; then
    echo "python3 required to parse checksums JSON" >&2
    exit 1
fi
PY=python3
command -v python3 >/dev/null 2>&1 || PY=python

mapfile -t ENTRIES < <("$PY" - "$TMP_JSON" <<'PY'
import json, sys
data = json.load(open(sys.argv[1], encoding="utf-8"))
for item in data.get("files") or []:
    path = str(item.get("path") or "").replace("\\", "/").strip()
    digest = str(item.get("sha256") or "").strip().lower()
    if path and digest:
        print(path + "\t" + digest)
PY
)

ok=0
downloaded=0
failed=0
ao_synced=0
total_manifest=${#ENTRIES[@]}
installable=0

echo "Manifest entries: $total_manifest"
echo

for line in "${ENTRIES[@]}"; do
    rel="${line%%$'\t'*}"
    want="${line#*$'\t'}"
    if ! is_installable "$rel"; then
        continue
    fi
    installable=$((installable + 1))
    local_path="$DEST/$rel"
    have=""
    if [[ -f "$local_path" ]]; then
        have="$(sha256_file "$local_path" || true)"
    fi

    if [[ -n "$have" && -n "$want" && "$have" == "$want" ]]; then
        echo "ok      $rel"
        ok=$((ok + 1))
    else
        why="mismatch"
        [[ -z "$have" ]] && why="missing"
        echo "scp     $rel  ($why)"
        mkdir -p "$(dirname "$local_path")"
        if scp "$REMOTE:$REMOTE_SRC/$rel" "$local_path"; then
            have2="$(sha256_file "$local_path" || true)"
            if [[ -n "$want" && -n "$have2" && "$have2" != "$want" ]]; then
                echo "WARN    $rel — sha256 still differs after scp"
            fi
            downloaded=$((downloaded + 1))
        else
            echo "FAIL    $rel"
            failed=$((failed + 1))
            continue
        fi
    fi

    ao_path="$DEST_AO/$rel"
    if copy_if_different "$local_path" "$ao_path"; then
        echo "ao      $rel"
        ao_synced=$((ao_synced + 1))
    fi
done

# Optional terms_of_use (txt may be absent from py-only manifest)
terms_rel="pythonpath/terms_of_use_ru.txt"
terms_local="$DEST/$terms_rel"
if [[ ! -f "$terms_local" ]]; then
    echo "scp     $terms_rel  (optional)"
    mkdir -p "$(dirname "$terms_local")"
    if scp "$REMOTE:$REMOTE_SRC/$terms_rel" "$terms_local" 2>/dev/null; then
        downloaded=$((downloaded + 1))
        if copy_if_different "$terms_local" "$DEST_AO/$terms_rel"; then
            ao_synced=$((ao_synced + 1))
        fi
    else
        echo "skip    $terms_rel"
    fi
fi

echo
echo "Installable: $installable"
echo "Done: ok=$ok downloaded=$downloaded ao_synced=$ao_synced failed=$failed"
[[ "$failed" -eq 0 ]]
