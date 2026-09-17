#!/usr/bin/env bash
set -euo pipefail
if [[ $# -ne 1 ]]; then echo "usage: $0 /path/to/DocumentFormat" >&2; exit 2; fi
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$1" && pwd)"
mkdir -p "$REPO/scripts" "$REPO/tests" "$REPO/references"
cp -f "$HERE"/scripts/* "$REPO/scripts/"
cp -f "$HERE"/tests/* "$REPO/tests/"
cp -f "$HERE"/references/* "$REPO/references/"
rm -rf "$REPO/PATCH_NOTES.md" "$REPO/apply_fix.ps1" "$REPO/apply_fix.sh" "$REPO/templates/generic-formal/template.docx.bak"
echo "Applied DocumentFormat template-runtime repair to $REPO"
echo "Next: python -m pip install -e '.[test]'; pytest -q"
