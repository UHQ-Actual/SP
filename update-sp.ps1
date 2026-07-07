<#
update-sp.ps1 - Clone or update the SP tools repo on this machine.

Usage (from any PowerShell prompt):
    powershell -ExecutionPolicy Bypass -File .\update-sp.ps1
    powershell -ExecutionPolicy Bypass -File .\update-sp.ps1 -Dest "C:\Tools\SP"

First run clones (or downloads) the repo; later runs pull the latest.
Uses git when available; otherwise falls back to downloading the repo
ZIP from GitHub (the repo is public, so no sign-in is needed either way).
#>
param(
    [string]$Dest = (Join-Path $HOME "SP")
)

$RepoUrl = "https://github.com/UHQ-Actual/SP.git"
$ZipUrl  = "https://github.com/UHQ-Actual/SP/archive/refs/heads/main.zip"

$ErrorActionPreference = "Stop"

$haveGit = [bool](Get-Command git -ErrorAction SilentlyContinue)

if ($haveGit) {
    if (Test-Path (Join-Path $Dest ".git")) {
        Write-Host "Updating repo in $Dest"
        git -C $Dest pull --ff-only
    } else {
        Write-Host "Cloning to $Dest"
        git clone $RepoUrl $Dest
    }
} else {
    Write-Host "git not found - downloading ZIP snapshot instead"

    # Windows PowerShell 5.1 defaults to TLS 1.0, which GitHub rejects
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $tmpZip = Join-Path $env:TEMP "SP-main.zip"
    $tmpDir = Join-Path $env:TEMP "SP-main-extract"

    Invoke-WebRequest -Uri $ZipUrl -OutFile $tmpZip -UseBasicParsing
    if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force }
    Expand-Archive -Path $tmpZip -DestinationPath $tmpDir

    if (-not (Test-Path $Dest)) { New-Item -ItemType Directory -Path $Dest | Out-Null }
    Copy-Item -Path (Join-Path $tmpDir "SP-main\*") -Destination $Dest -Recurse -Force

    Remove-Item $tmpZip -Force
    Remove-Item $tmpDir -Recurse -Force
    Write-Host "Snapshot extracted to $Dest"
}

Write-Host ""
Write-Host "Done. Files in ${Dest}:"
Get-ChildItem $Dest -File | Select-Object -ExpandProperty Name
