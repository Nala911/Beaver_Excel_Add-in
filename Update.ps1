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
    [switch]$KeepAlive,
    [string]$Format = "Text",
    [switch]$FailedOnly,
    [switch]$AutoFix,
    [string]$ShowDeps,
    [switch]$LintOnly,
    [switch]$GenerateDocs,
    [switch]$SkipDocs,
    [switch]$Fast,
    [switch]$Quick,
    [string]$TestCategory,
    [switch]$Status,
    [string[]]$File,
    [switch]$ExportAddin,
    [switch]$AddFeature,
    [string]$ControlId,
    [string]$Label,
    [string]$Group,
    [string]$Tab = "BeaverTab",
    [string]$Icon = "FunctionWizard",
    [string]$Keytip,
    [string]$Shortcut,
    [string]$Screentip,
    [string]$Supertip,
    [switch]$AddHotkey,
    [string]$Key,
    [string]$Macro,
    [string]$CommandName,
    [string]$Description,
    [switch]$ValidateManifest,
    [switch]$Repair
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
