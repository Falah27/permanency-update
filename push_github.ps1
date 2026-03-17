$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

param(
    [string]$Message = ""
)

git add -A

git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    Write-Host "No staged changes to push."
    exit 0
}

if ([string]::IsNullOrWhiteSpace($Message)) {
    $Message = "sync $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

git commit -m $Message
git pull --rebase origin main
git push origin main

Write-Host "Push to GitHub complete."