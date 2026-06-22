# Script:   Update.ps1
# Purpose:  Lightweight entry point stub wrapper for the Beaver Add-in pipeline.
#           Delegates execution to the .build/Update.ps1 orchestrator script.
# ==============================================================================

[CmdletBinding(PositionalBinding=$false)]
param(
    [switch]$SkipRuntimeTests,
    [string]$Filter,
    [switch]$Force,
    [switch]$ListTests,
    [switch]$SkipLint,
    [switch]$Visible,
    [switch]$Clean,
    [switch]$KeepAlive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$orchestratorPath = Join-Path $PSScriptRoot ".build\Update.ps1"
if (-not (Test-Path $orchestratorPath)) {
    Write-Error "Orchestrator script not found at: $orchestratorPath"
    exit 1
}

# Forward all arguments using PSBoundParameters
& $orchestratorPath @PSBoundParameters
exit $LASTEXITCODE
