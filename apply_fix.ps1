param(
    [Parameter(Mandatory=$true)][string]$Repo
)
$ErrorActionPreference = "Stop"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Repo = (Resolve-Path $Repo).Path
$dirs = @("scripts", "tests", "references")
foreach ($dir in $dirs) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Repo $dir) | Out-Null
    Copy-Item -Force -Recurse (Join-Path $Here "$dir\*") (Join-Path $Repo $dir)
}
$obsolete = @(
    "PATCH_NOTES.md",
    "apply_fix.ps1",
    "apply_fix.sh",
    "templates\generic-formal\template.docx.bak"
)
foreach ($item in $obsolete) {
    $path = Join-Path $Repo $item
    if (Test-Path $path) { Remove-Item -Force -Recurse $path }
}
Write-Host "Applied DocumentFormat template-runtime repair to $Repo"
Write-Host "Next: python -m pip install -e '.[test]'; pytest -q"
