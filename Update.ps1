# Script:   Update.ps1
# Purpose:  Orchestrator for the Beaver Add-in build and testing pipelines.
# Usage:    .\Update.ps1
#           .\Update.ps1 -SkipRuntimeTests
# ==============================================================================

[CmdletBinding()]
param(
    [switch]$SkipRuntimeTests
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  BEAVER ADD-IN: RUNNING BUILD PIPELINE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

& "$PSScriptRoot\Build.ps1"
if (-not $?) {
    Write-Host "Build pipeline failed." -ForegroundColor Red
    exit 1
}

if ($SkipRuntimeTests) {
    Write-Host ""
    Write-Host "Skipping test pipeline (-SkipRuntimeTests)." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  BEAVER ADD-IN: RUNNING TEST PIPELINE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

& "$PSScriptRoot\Test.ps1"
if (-not $?) {
    Write-Host "Test pipeline failed." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "All pipelines completed successfully!" -ForegroundColor Green
exit 0
