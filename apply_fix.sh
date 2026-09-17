#!/usr/bin/env bash
set -euo pipefail
repo="${1:-.}"
overlay="$(cd "$(dirname "$0")" && pwd)"
files=(
  SKILL.md .gitignore pyproject.toml
  references/template-contract.md references/validation.md references/word-native-structures.md
  scripts/profile_config.py scripts/fast_format_docx.py scripts/inspect_docx.py scripts/repair_cross_references.py scripts/validate_docx.py
  templates/generic-formal/profile.yaml tests/test_formatter.py
)
for f in "${files[@]}"; do
  mkdir -p "$repo/$(dirname "$f")"
  cp -f "$overlay/$f" "$repo/$f"
done
rm -rf "$repo/.idea"
rm -f "$repo/templates/generic-formal/template.docx.bak"
echo "Applied DocumentFormat fixes to $repo"
echo "Run: python -m pip install -e '.[test]' && pytest -q"
