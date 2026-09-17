param([string]$Repo = ".")
$ErrorActionPreference = "Stop"
$overlay = Split-Path -Parent $MyInvocation.MyCommand.Path
$files = @(
  "SKILL.md", ".gitignore", "pyproject.toml",
  "references/template-contract.md", "references/validation.md", "references/word-native-structures.md",
  "scripts/profile_config.py", "scripts/fast_format_docx.py", "scripts/inspect_docx.py", "scripts/repair_cross_references.py", "scripts/validate_docx.py",
  "templates/generic-formal/profile.yaml", "tests/test_formatter.py"
)
foreach ($f in $files) {
  $src = Join-Path $overlay $f
  $dst = Join-Path $Repo $f
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst) | Out-Null
  Copy-Item -Force $src $dst
}
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue (Join-Path $Repo ".idea")
Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path $Repo "templates/generic-formal/template.docx.bak")
Write-Host "Applied DocumentFormat fixes to $Repo"
Write-Host "Run: python -m pip install -e '.[test]' ; pytest -q"
