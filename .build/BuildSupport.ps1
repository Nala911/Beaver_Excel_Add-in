# Script:   BuildSupport.ps1
# Purpose:  Shared helpers, paths, and Excel COM management for Build.ps1 and Test.ps1 (Modular wrapper).
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Common Paths ---
$projectRoot = Split-Path $PSScriptRoot -Parent
$excelPath = Join-Path $projectRoot "Beaver.xlsm"
$modulesDir = Join-Path $projectRoot "Modules"
$diskThisWorkbookCls = Join-Path $projectRoot "ThisWorkbook.cls"
$ribbonXmlPath = Join-Path $projectRoot "ribbon.xml"
$featureManifestPath = Join-Path $projectRoot "features.json"
$configPath = Join-Path $projectRoot "config.json"
$testManifestPath = Join-Path $modulesDir "Tests\Test_Manifest.bas"
$commandRegistryPath = Join-Path $modulesDir "Infrastructure\Infra_CommandRegistry.bas"
$configManifestPath = Join-Path $modulesDir "Infrastructure\Infra_ConfigManifest.bas"
$uiRibbonPath = Join-Path $modulesDir "UI\UI_Ribbon.bas"
$uiHotkeysPath = Join-Path $modulesDir "UI\UI_Hotkeys.bas"
$structuredTestResultsPath = Join-Path $env:TEMP "BeaverAddin.TestResults.tsv"
$buildStatePath = Join-Path $PSScriptRoot ".build_state.json"

# --- Dot-source all modular build sub-libraries ---
. (Join-Path $PSScriptRoot "lib\Logging.ps1")
. (Join-Path $PSScriptRoot "lib\HashingState.ps1")
. (Join-Path $PSScriptRoot "lib\AstParser.ps1")
. (Join-Path $PSScriptRoot "lib\ComUtils.ps1")
. (Join-Path $PSScriptRoot "lib\Generators.ps1")
. (Join-Path $PSScriptRoot "lib\RibbonUtils.ps1")

# Run cleanup at load time to purge old logs from previous sessions
Clear-AccumulatedLogs
