# Script:   Update.ps1
# Purpose:  Orchestrator for the Beaver Add-in build and testing pipelines.
# Usage:    .\Update.ps1
#           .\Update.ps1 -SkipRuntimeTests
#           .\Update.ps1 -Filter "CleanData"
# ==============================================================================

[CmdletBinding()]
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
    [string]$TestCategory,
    [switch]$Status,
    [string[]]$File
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "BuildSupport.ps1")

# Enable orchestrator mode to share Excel session and hashes across stages
$global:BeaverOrchestratorActive = $true
$global:BeaverKeepAliveActive = $KeepAlive
$global:BeaverFormatJson = ($Format -eq "Json")
$global:BeaverSourceHashes = $null
$global:BeaverSharedExcel = $null
$global:BeaverExcelWasAlreadyOpen = $false
$global:BeaverWorkbookWasAlreadyOpen = $false
$global:BeaverBuildStateCache = $null
$global:BeaverFeatureManifestCache = $null
$global:BeaverFileContentCache = $null

try {
    if ($Status) {
        Write-Host "=========================================" -ForegroundColor Cyan
        Write-Host "     BEAVER WORKSPACE AGENT STATUS       " -ForegroundColor Cyan
        Write-Host "=========================================" -ForegroundColor Cyan

        # 1. Add-in Identity
        if (Test-Path $configPath) {
            $config = Get-Content $configPath -Raw | ConvertFrom-Json
            Write-Host "Add-in Name:    $($config.AddinIdentity.Name)" -ForegroundColor Green
            Write-Host "Version:        $($config.AddinIdentity.Version)" -ForegroundColor Green
            if ($null -ne $config.AddinIdentity.PSObject.Properties['ReleaseTier']) {
                Write-Host "Release Tier:   $($config.AddinIdentity.ReleaseTier)" -ForegroundColor Green
            }
        }

        # 2. Git Status
        Write-Host ""
        Write-Host "Git Status:" -ForegroundColor Yellow
        $gitBranch = (git branch --show-current 2>$null)
        if ($gitBranch) {
            Write-Host "  Branch: $gitBranch" -ForegroundColor Gray
        }
        $gitStatus = (git status --short 2>$null)
        if ($gitStatus) {
            $gitStatus -split "`r`n" | ForEach-Object { 
                if ($_.Trim()) { Write-Host "  $_" -ForegroundColor Yellow }
            }
        } else {
            Write-Host "  Clean (no uncommitted changes)" -ForegroundColor Green
        }

        # 3. Module Breakdown
        Write-Host ""
        Write-Host "Module Breakdown (by Category):" -ForegroundColor Yellow
        $buildState = Get-BuildState
        $categories = @{}
        $totalModules = 0
        $useCache = $null -ne $buildState -and $null -ne $buildState.Metadata
        
        if ($useCache) {
            foreach ($prop in $buildState.Metadata.PSObject.Properties) {
                $relPath = $prop.Name
                if ($relPath -match "\.(bas|cls|frm)$" -or $relPath -eq "ThisWorkbook.cls") {
                    $totalModules++
                    $meta = $prop.Value
                    $category = "Unknown"
                    if ($meta.PSObject.Properties.Name -contains "Category" -and $null -ne $meta.Category) {
                        $category = $meta.Category
                    }
                    $categories[$category] = [int]$categories[$category] + 1
                }
            }
        }
        
        if ($totalModules -eq 0 -and (Test-Path $modulesDir)) {
            $vbaFiles = Get-ChildItem -Path $modulesDir -Include *.bas, *.cls, *.frm -Recurse
            foreach ($file in $vbaFiles) {
                $headerLines = [System.IO.File]::ReadLines($file.FullName) | Select-Object -First 15
                $category = "Unknown"
                foreach ($line in $headerLines) {
                    if ($line -match "'\s*@Category:\s*([^\r\n]+)") {
                        $category = $Matches[1].Trim()
                        break
                    }
                }
                $categories[$category] = [int]$categories[$category] + 1
            }
            $totalModules = $vbaFiles.Count
        }
        
        foreach ($cat in $categories.Keys) {
            Write-Host "  $($cat): $($categories[$cat]) module(s)" -ForegroundColor Gray
        }
        Write-Host "  Total Modules: $totalModules" -ForegroundColor Gray

        # 4. Build State & Logs
        Write-Host ""
        Write-Host "Build & Test Cache State:" -ForegroundColor Yellow
        $buildState = Get-BuildState
        if ($null -ne $buildState) {
            Write-Host "  Last Build Time: $($buildState.LastBuildTime)" -ForegroundColor Gray
            $testsPassed = ($buildState.PSObject.Properties.Name.Contains("TestsPassed") -and $buildState.TestsPassed -eq $true)
            if ($testsPassed) {
                Write-Host "  Last Tests Status: PASSED" -ForegroundColor Green
            } else {
                Write-Host "  Last Tests Status: FAILED / UNRUN" -ForegroundColor Red
            }
        } else {
            Write-Host "  No build state cache found." -ForegroundColor Red
        }

        # 5. Schema Validity
        if (Test-Path $featureManifestPath) {
            $schemaPath = Join-Path $projectRoot "features.schema.json"
            if (Test-Path $schemaPath) {
                $jsonContent = Get-Content $featureManifestPath -Raw
                if (Test-Json -Json $jsonContent -SchemaFile $schemaPath) {
                    Write-Host "  features.json:   VALID against schema" -ForegroundColor Green
                } else {
                    Write-Host "  features.json:   INVALID against schema" -ForegroundColor Red
                }
            }
        }

        # 6. Syntax / Lint Check Summary
        Write-Host ""
        Write-Host "Running quick syntax / lint scan..." -ForegroundColor Yellow
        . (Join-Path $PSScriptRoot "Linter.ps1")
        $projectChanges = Get-ProjectChanges -Force:$Force
        $changedFiles = $projectChanges.ChangedFiles
        $filesToValidate = if ($Force) { $null } else { $changedFiles }
        
        $validLint = Invoke-VbaLint -SourceDir $modulesDir -FilesToProcess $filesToValidate -AutoFix:$false
        if ($validLint) {
            Write-Host "  Syntax / Lint check: PASSED" -ForegroundColor Green
        } else {
            Write-Host "  Syntax / Lint check: FAILED (see output above)" -ForegroundColor Red
        }

        exit 0
    }

    if ($LintOnly) {
        . (Join-Path $PSScriptRoot "Linter.ps1")
        $projectChanges = Get-ProjectChanges -Force:$Force
        $changedFiles = $projectChanges.ChangedFiles
        if ($null -ne $File -and $File.Count -gt 0) {
            $normalizedFiles = @()
            foreach ($f in $File) {
                $rel = $f.Replace("\", "/")
                if ($rel.StartsWith("./")) { $rel = $rel.Substring(2) }
                $normalizedFiles += $rel
            }
            $changedFiles = $normalizedFiles
        }
        $filesToValidate = if ($Force -and $null -eq $File) { $null } else { $changedFiles }
        
        $validLint = Invoke-VbaLint -SourceDir $modulesDir -FilesToProcess $filesToValidate -AutoFix:$AutoFix
        if ($validLint) {
            Write-Host "Lint check completed successfully!" -ForegroundColor Green
            exit 0
        } else {
            Write-Host "Lint check failed with errors." -ForegroundColor Red
            exit 1
        }
    }

    if ($TestCategory) {
        switch ($TestCategory.ToLower().Trim()) {
            "ui" { $Filter = "Test_UI.*" }
            "feature" { $Filter = "Test_Feat_*.*" }
            "feat" { $Filter = "Test_Feat_*.*" }
            "infrastructure" { $Filter = "Test_CommandInfrastructure.*,Test_Runner.*" }
            "infra" { $Filter = "Test_CommandInfrastructure.*,Test_Runner.*" }
            "core" { $Filter = "Test_Runner.*" }
            default {
                Write-Warning "Unknown TestCategory '$TestCategory'. Using category as direct wild-card filter."
                $Filter = "*$TestCategory*"
            }
        }
    }

    if ($ShowDeps) {
        Show-Dependencies -Target $ShowDeps
        exit 0
    }

    if ($ListTests) {
        $listParams = @{ ListTests = $true }
        if ($Filter) { $listParams["Filter"] = $Filter }
        & "$PSScriptRoot\Test.ps1" @listParams
        exit $LASTEXITCODE
    }

    if ($Clean) {
        & "$PSScriptRoot\Build.ps1" -Clean
        exit $LASTEXITCODE
    }

    # Check project changes using centralized helper
    $projectChanges = Get-ProjectChanges -Force:$Force
    $changedFiles = $projectChanges.ChangedFiles
    $deletedFiles = $projectChanges.DeletedFiles
    $manifestChanged = $projectChanges.ManifestChanged
    $manifestStructureChanged = $projectChanges.ManifestStructureChanged
    $hasAnyChanges = $projectChanges.HasAnyChanges

    if ($null -ne $File -and $File.Count -gt 0) {
        $normalizedFiles = @()
        foreach ($f in $File) {
            $rel = $f.Replace("\", "/")
            if ($rel.StartsWith("./")) { $rel = $rel.Substring(2) }
            $normalizedFiles += $rel
        }
        $changedFiles = $normalizedFiles
        $hasAnyChanges = $true
        $manifestChanged = ($changedFiles -contains "features.json")
        $manifestStructureChanged = $manifestChanged
    }

    $buildState = Get-BuildState
    $skipUnitTests = $false
    $testsPassed = $false
    $autoFilter = $null

    if ($null -ne $buildState) {
        $testsPassed = ($buildState.PSObject.Properties.Name.Contains("TestsPassed") -and $buildState.TestsPassed -eq $true)

        # Calculate if there are any VBA code changes on disk
        $hasVbaChanges = $false
        foreach ($file in $changedFiles) {
            if ($file -match "\.(bas|cls|frm)$" -or $file -eq "ThisWorkbook.cls") {
                $hasVbaChanges = $true
                break
            }
        }
        if (@($deletedFiles).Count -gt 0) {
            $hasVbaChanges = $true
        }

        $skipUnitTests = (-not $hasVbaChanges -and -not $manifestStructureChanged -and -not $Force -and $testsPassed)

        if (-not $hasAnyChanges -and -not $Force -and -not $Filter -and (Test-Path $excelPath) -and $testsPassed) {
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "  BEAVER ADD-IN: PIPELINE UP TO DATE" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "No changes detected and all tests passed on the current codebase state. Skipping build and tests." -ForegroundColor Green
            if ($null -ne $global:BeaverBuildLog) {
                $global:BeaverBuildLog.buildMode = "skipped"
            }
            Record-BuildChanges -ManifestChanged $manifestChanged -ManifestStructureChanged $manifestStructureChanged -ChangedFiles $changedFiles -DeletedFiles $deletedFiles -Force $Force
            Save-BuildLog -Status "success" -Force

            if ($GenerateDocs) {
                $genDocsPath = Join-Path $PSScriptRoot "GenerateArchitectureMap.ps1"
                if (Test-Path $genDocsPath) {
                    Write-Host ""
                    Write-Host "Running auto-generation of ARCHITECTURE.md..." -ForegroundColor Cyan
                    try {
                        & $genDocsPath
                        Write-Host "  ARCHITECTURE.md regenerated successfully." -ForegroundColor Green
                    } catch {
                        Write-Warning "Failed to auto-regenerate ARCHITECTURE.md: $($_.Exception.Message)"
                    }
                }
            }
            exit 0
        }
        
        # Smart Test Filtering: calculate affected tests dynamically using dependency tracing
        if (-not $Filter -and -not $Force -and -not $manifestChanged) {
            if (@($changedFiles).Count -gt 0) {
                $impactedTestNames = Get-ImpactedTests -ChangedFiles $changedFiles -DeletedFiles $deletedFiles
                if (@($impactedTestNames).Count -gt 0) {
                    $patterns = @()
                    foreach ($testName in $impactedTestNames) {
                        $patterns += "*$testName*"
                    }
                    $autoFilter = $patterns -join ","
                } else {
                    # No test procedures depend on the changed modules.
                    Write-Host "Smart Testing: No test procedures are affected by the changed modules. Skipping unit tests." -ForegroundColor Green
                    $skipUnitTests = $true
                }
            }
        }
    }

    Record-BuildChanges -ManifestChanged $manifestChanged -ManifestStructureChanged $manifestStructureChanged -ChangedFiles $changedFiles -DeletedFiles $deletedFiles -Force $Force

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  BEAVER ADD-IN: RUNNING BUILD PIPELINE" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $buildParams = @{}
    if ($Force) { $buildParams["Force"] = $true }
    if ($SkipLint) { $buildParams["SkipLint"] = $true }
    if ($AutoFix) { $buildParams["AutoFix"] = $true }
    if ($null -ne $File -and $File.Count -gt 0) { $buildParams["File"] = $File }

    & "$PSScriptRoot\Build.ps1" @buildParams
    if (-not $?) {
        Write-Host "Build pipeline failed." -ForegroundColor Red
        Save-BuildLog -Status "failure" -Force
        exit 1
    }

    if ($SkipRuntimeTests) {
        Write-Host ""
        Write-Host "Skipping test pipeline (-SkipRuntimeTests)." -ForegroundColor Yellow
        Save-BuildLog -Status "success" -Force
        exit 0
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  BEAVER ADD-IN: RUNNING TEST PIPELINE" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $testParams = @{}
    if ($Filter) {
        $testParams["Filter"] = $Filter
    } elseif ($autoFilter) {
        $testParams["Filter"] = $autoFilter
        Write-Host "Smart Testing: Automatically filtering tests to '$autoFilter' based on changed modules." -ForegroundColor Yellow
    }
    if ($manifestChanged) { $testParams["CheckRibbon"] = $true }
    if ($Visible) { $testParams["Visible"] = $true }
    if ($null -ne $skipUnitTests -and $skipUnitTests) {
        $testParams["SkipUnitTests"] = $true
    }
    if ($FailedOnly) {
        $testParams["FailedOnly"] = $true
    }

    & "$PSScriptRoot\Test.ps1" @testParams
    if (-not $?) {
        Set-BuildStateTestsPassed -Passed $false
        Write-Host "Test pipeline failed." -ForegroundColor Red
        Save-BuildLog -Status "failure" -Force
        exit 1
    }

    # Only mark TestsPassed = true if the full test suite was run and passed successfully
    if ($null -ne $skipUnitTests -and $skipUnitTests) {
        Set-BuildStateTestsPassed -Passed $testsPassed
    } elseif (-not $Filter -and -not $autoFilter) {
        Set-BuildStateTestsPassed -Passed $true
    } else {
        Set-BuildStateTestsPassed -Passed $false
    }

    # Save the workbook if everything succeeded and we have an open workbook session
    if ($global:BeaverSharedExcel) {
        try {
            $wbs = $global:BeaverSharedExcel.Workbooks
            foreach ($wb in $wbs) {
                if ($wb.Name -eq "Beaver Add-in.xlsm") {
                    try {
                        if ($global:BeaverSharedExcel.Calculation -ne -4105) {
                            $global:BeaverSharedExcel.Calculation = -4105 # xlCalculationAutomatic
                            Write-Host "  Ensuring calculation option is Automatic before final save." -ForegroundColor Green
                        }
                    } catch {}
                    Write-Host "Saving workbook after successful tests (Save-On-Success)..." -ForegroundColor Green
                    $wb.Save()
                    break
                }
            }
            Release-ComObjectSafely $wbs
        } catch {
            Write-Warning "Failed to save workbook on success: $($_.Exception.Message)"
        }
    }

    # Auto-generate ARCHITECTURE.md if changes occurred or if requested, and not skipped
    $shouldGenerateDocs = $GenerateDocs -or ($hasAnyChanges -and -not $SkipDocs)
    if ($shouldGenerateDocs) {
        $genDocsPath = Join-Path $PSScriptRoot "GenerateArchitectureMap.ps1"
        if (Test-Path $genDocsPath) {
            Write-Host ""
            Write-Host "Running auto-generation of ARCHITECTURE.md..." -ForegroundColor Cyan
            try {
                & $genDocsPath
                Write-Host "  ARCHITECTURE.md regenerated successfully." -ForegroundColor Green
            } catch {
                Write-Warning "Failed to auto-regenerate ARCHITECTURE.md: $($_.Exception.Message)"
            }
        }
    }

    Write-Host ""
    Write-Host "All pipelines completed successfully!" -ForegroundColor Green
    Save-BuildLog -Status "success" -Force
    exit 0
} finally {
    # Clean up persistent Excel session if orchestrator opened it
    if ($null -ne $global:BeaverSharedExcel) {
        Close-ExcelWorkbookSession -Excel $global:BeaverSharedExcel `
                                   -WasAlreadyOpen $global:BeaverExcelWasAlreadyOpen `
                                   -KeepAlive $global:BeaverKeepAliveActive `
                                   -WorkbookPath $excelPath
        $global:BeaverSharedExcel = $null
    }
    
    $global:BeaverOrchestratorActive = $false
    $global:BeaverSourceHashes = $null
    Clear-AccumulatedLogs
}
