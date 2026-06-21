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
    [switch]$Clean
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "BuildSupport.ps1")

# Enable orchestrator mode to share Excel session and hashes across stages
$global:BeaverOrchestratorActive = $true
$global:BeaverSourceHashes = $null
$global:BeaverSharedExcel = $null
$global:BeaverExcelWasAlreadyOpen = $false

try {
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
    $autoFilter = $null
    $changedFiles = @()
    $deletedFiles = @()

    if ($null -ne $buildState -and $null -ne $buildState.Files) {
        $manifestChanged = (-not $buildState.Files.PSObject.Properties.Name.Contains("features.json") -or $buildState.Files."features.json" -ne $currentHashes["features.json"])
        
        # Calculate changes early
        foreach ($key in $currentHashes.Keys) {
            if (-not $buildState.Files.PSObject.Properties.Name.Contains($key) -or $buildState.Files.$key -ne $currentHashes[$key]) {
                $changedFiles += $key
            }
        }
        foreach ($prop in $buildState.Files.PSObject.Properties) {
            $key = $prop.Name
            if (-not $currentHashes.ContainsKey($key)) {
                $deletedFiles += $key
            }
        }

        # Check if we can skip the entire build/test pipeline
        $hasAnyChanges = ($changedFiles.Count -gt 0 -or $deletedFiles.Count -gt 0)
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
        if ($deletedFiles.Count -gt 0) {
            $hasVbaChanges = $true
        }

        $skipUnitTests = (-not $hasVbaChanges -and -not $manifestStructureChanged -and -not $Force)

        if (-not $hasAnyChanges -and -not $Force -and -not $Filter -and (Test-Path $excelPath) -and $testsPassed) {
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "  BEAVER ADD-IN: PIPELINE UP TO DATE" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "No changes detected and all tests passed on the current codebase state. Skipping build and tests." -ForegroundColor Green
            exit 0
        }
        
        # Smart Test Filtering: detect changed commands to run targeted tests by default
        if (-not $Filter -and -not $Force -and -not $manifestChanged) {
            if ($changedFiles.Count -gt 0) {
                $canAutoFilter = $true
                $featureNames = [System.Collections.Generic.HashSet[string]]::new()
                
                foreach ($file in $changedFiles) {
                    $normPath = $file.Replace("\", "/")
                    if ($normPath -match "Modules/Commands/FeatCmd_(\w+)\.cls$") {
                        [void]$featureNames.Add($Matches[1])
                    } elseif ($normPath -eq "Modules/Libraries/Lib_Tests_Features.bas" -or
                              $normPath -eq "Modules/Libraries/Lib_TestManifest.bas" -or
                              $normPath -eq "Modules/Infrastructure/Infra_CommandRegistry.bas" -or
                              $normPath -eq "Modules/UI/UI_Ribbon.bas" -or
                              $normPath -eq "Modules/UI/UI_Hotkeys.bas") {
                        continue
                    } else {
                        $canAutoFilter = $false
                        break
                    }
                }
                
                if ($canAutoFilter -and $featureNames.Count -gt 0) {
                    $patterns = @()
                    foreach ($name in $featureNames) {
                        $patterns += "*$name*"
                    }
                    $autoFilter = $patterns -join ","
                }
            }
        }
    }

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  BEAVER ADD-IN: RUNNING BUILD PIPELINE" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $buildParams = @{}
    if ($Force) { $buildParams["Force"] = $true }
    if ($SkipLint) { $buildParams["SkipLint"] = $true }

    & "$PSScriptRoot\Build.ps1" @buildParams
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

    & "$PSScriptRoot\Test.ps1" @testParams
    if (-not $?) {
        Set-BuildStateTestsPassed -Passed $false
        Write-Host "Test pipeline failed." -ForegroundColor Red
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

    Write-Host ""
    Write-Host "All pipelines completed successfully!" -ForegroundColor Green
    exit 0
} finally {
    # Clean up persistent Excel session if orchestrator opened it
    if ($null -ne $global:BeaverSharedExcel) {
        if (-not $global:BeaverExcelWasAlreadyOpen) {
            Write-Host "Cleaning up persistent Excel session..." -ForegroundColor Cyan
            try {
                foreach ($wb in $global:BeaverSharedExcel.Workbooks) {
                    if ($wb.FullName -eq $excelPath) {
                        $wb.Close($false)
                    }
                }
            } catch {}
            try {
                $global:BeaverSharedExcel.Quit()
            } catch {}
        } else {
            try {
                $global:BeaverSharedExcel.Visible = $true
                $global:BeaverSharedExcel.DisplayAlerts = $true
            } catch {}
        }
        Release-ComObjectSafely $global:BeaverSharedExcel
        $global:BeaverSharedExcel = $null
    }
    
    $global:BeaverOrchestratorActive = $false
    $global:BeaverSourceHashes = $null
}
