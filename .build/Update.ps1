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
    [string]$ShowDeps
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

    # Check if manifest changed to decide if we validate the Ribbon
    $currentHashes = Get-SourceFileHashes
    $buildState = Get-BuildState
    $manifestChanged = $true
    $manifestStructureChanged = $true
    $skipUnitTests = $false
    $testsPassed = $false
    $autoFilter = $null
    $changedFiles = @()
    $deletedFiles = @()

    if ($null -ne $buildState -and $null -ne $buildState.Files) {
        $featProp = $buildState.Files.PSObject.Properties["features.json"]
        $manifestChanged = ($null -eq $featProp -or $featProp.Value -ne $currentHashes["features.json"])
        
        # Calculate changes early
        foreach ($key in $currentHashes.Keys) {
            $prop = $buildState.Files.PSObject.Properties[$key]
            if ($null -eq $prop -or $prop.Value -ne $currentHashes[$key]) {
                $changedFiles += $key
            }
        }
        foreach ($prop in $buildState.Files.PSObject.Properties) {
            $key = $prop.Name
            if (-not $currentHashes.ContainsKey($key)) {
                $deletedFiles += $key
            }
        }
        
        # Merge previously failed files to force their rebuild/re-import
        if ($buildState.PSObject.Properties.Name -contains "FailedFiles" -and $null -ne $buildState.FailedFiles) {
            foreach ($file in $buildState.FailedFiles) {
                if ($changedFiles -notcontains $file) {
                    $changedFiles += $file
                }
            }
        }

        # Check if we can skip the entire build/test pipeline
        $hasAnyChanges = (@($changedFiles).Count -gt 0 -or @($deletedFiles).Count -gt 0)
        $testsPassed = $false
        if ($buildState.PSObject.Properties.Name.Contains("TestsPassed") -and $buildState.TestsPassed -eq $true) {
            $testsPassed = $true
        }

        # Calculate structural manifest changes
        $manifestStructureChanged = $false
        if ($manifestChanged) {
            $newStructuralHash = Get-ManifestStructuralHash -Path $featureManifestPath
            $oldStructuralHash = $null
            if ($buildState.PSObject.Properties.Name.Contains("ManifestStructuralHash")) {
                $oldStructuralHash = $buildState.ManifestStructuralHash
            }
            if ($newStructuralHash -ne $oldStructuralHash) {
                $manifestStructureChanged = $true
            }
        }

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
    } else {
        $changedFiles = @($currentHashes.Keys)
    }

    Record-BuildChanges -ManifestChanged $manifestChanged -ManifestStructureChanged $manifestStructureChanged -ChangedFiles $changedFiles -DeletedFiles $deletedFiles -Force $Force

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  BEAVER ADD-IN: RUNNING BUILD PIPELINE" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $buildParams = @{}
    if ($Force) { $buildParams["Force"] = $true }
    if ($SkipLint) { $buildParams["SkipLint"] = $true }
    if ($AutoFix) { $buildParams["AutoFix"] = $true }

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
