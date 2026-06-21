[CmdletBinding()]
param(
    [switch]$CheckRibbon,
    [string]$Filter,
    [switch]$ListTests,
    [switch]$Visible,
    [switch]$SkipUnitTests
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "BuildSupport.ps1")

# --- Helper Functions ---

function Reset-StructuredTestResults {
    param([string]$Path)

    if (Test-Path $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

function Get-TestProcedureLocation {
    param([string]$ProcName)

    if ([string]::IsNullOrWhiteSpace($ProcName)) { return $null }

    # Strip any module prefix if present
    $cleanProcName = $ProcName
    if ($ProcName -match "\.([^.]+)$") {
        $cleanProcName = $Matches[1]
    }

    $projectRoot = Split-Path $PSScriptRoot -Parent
    $vbaFiles = Get-ChildItem -Path $modulesDir -Filter *.bas -Recurse
    foreach ($file in $vbaFiles) {
        $lines = Get-Content $file.FullName
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match "^\s*Public Sub\s+$cleanProcName\s*\(") {
                $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace("\", "/")
                return [pscustomobject]@{
                    File = $file.FullName
                    LineNumber = $i + 1
                    RelativePath = $relPath
                }
            }
        }
    }
    return $null
}

function Write-StructuredTestResultsSummary {
    param([pscustomobject]$StructuredResults)

    Write-Host ("  Structured test results: total={0}, passed={1}, failed={2}" -f $StructuredResults.Summary.Total, $StructuredResults.Summary.Passed, $StructuredResults.Summary.Failed) -ForegroundColor Cyan

    $failedCount = 0
    foreach ($result in @($StructuredResults.Results)) {
        if (-not $result.Passed) {
            $failedCount++
            Write-Host ("    [FAIL] {0} ({1} ms)" -f $result.Name, $result.DurationMs) -ForegroundColor Red
            
            # Retrieve definition file and line
            $loc = Get-TestProcedureLocation -ProcName $result.Name
            if ($null -ne $loc) {
                Write-Host ("           Location: {0} (Line {1})" -f $loc.RelativePath, $loc.LineNumber) -ForegroundColor DarkYellow
            }
            
            $messageText = if ([string]::IsNullOrWhiteSpace($result.Message)) { "No assertion message provided." } else { $result.Message }
            Write-Host ("           Details:  {0}" -f $messageText) -ForegroundColor Red
        }
    }

    if ($failedCount -eq 0) {
        Write-Host "    All assertions passed successfully." -ForegroundColor Green
    }
}

function Assert-StructuredTestResults {
    param(
        [pscustomobject]$StructuredResults,
        [string]$Path
    )

    if ($null -eq $StructuredResults -or $null -eq $StructuredResults.Summary) {
        throw "Structured test results file was not produced: $Path"
    }

    Write-StructuredTestResultsSummary -StructuredResults $StructuredResults

    if ($StructuredResults.Summary.Failed -gt 0) {
        $failedNames = @($StructuredResults.Results | Where-Object { -not $_.Passed } | ForEach-Object { $_.Name })
        $failedLabel = if ($failedNames.Count -gt 0) { $failedNames -join ", " } else { "unknown test(s)" }
        throw "Structured test results reported $($StructuredResults.Summary.Failed) failure(s): $failedLabel"
    }

    return "tests=$($StructuredResults.Summary.Total)"
}

function Invoke-HeadlessCallbackTests {
    param(
        $ExcelApplication,
        [object[]]$Callbacks
    )

    if ($null -eq $Callbacks -or $Callbacks.Count -eq 0) {
        Write-Host "No enabled headless-safe callbacks declared in features.json." -ForegroundColor Gray
        return "callbacks=0"
    }

    Write-Host "Running headless-safe callback tests..." -ForegroundColor Cyan
    $passed = 0

    foreach ($callbackFeature in $Callbacks) {
        $callbackName = [string]$callbackFeature.OnAction
        Write-Host "  Testing callback: $callbackName" -ForegroundColor Yellow
        $activeWindow = $null
        $originalFormulaBar = $null
        $originalHeadings = $null
        $originalWorkbookTabs = $null
        $originalHorizontalScrollBar = $null
        $originalVerticalScrollBar = $null

        try {
            $activeWindow = $ExcelApplication.ActiveWindow
            $originalFormulaBar = $ExcelApplication.DisplayFormulaBar

            if ($null -ne $activeWindow) {
                $originalHeadings = $activeWindow.DisplayHeadings
                $originalWorkbookTabs = $activeWindow.DisplayWorkbookTabs
                $originalHorizontalScrollBar = $activeWindow.DisplayHorizontalScrollBar
                $originalVerticalScrollBar = $activeWindow.DisplayVerticalScrollBar
            }

            $ExcelApplication.Run($callbackName, $null)
            Write-Host "    [PASS] $callbackName" -ForegroundColor Green
            $passed++
        } catch {
            Write-Host "    [FAIL] $callbackName - $($_.Exception.Message)" -ForegroundColor Red
            throw "Headless callback failed for '$callbackName': $($_.Exception.Message)"
        } finally {
            if ($null -ne $originalFormulaBar) {
                try { $ExcelApplication.DisplayFormulaBar = $originalFormulaBar } catch { }
            }

            if ($null -ne $activeWindow) {
                try {
                    if ($null -ne $originalHeadings) { $activeWindow.DisplayHeadings = $originalHeadings }
                    if ($null -ne $originalWorkbookTabs) { $activeWindow.DisplayWorkbookTabs = $originalWorkbookTabs }
                    if ($null -ne $originalHorizontalScrollBar) { $activeWindow.DisplayHorizontalScrollBar = $originalHorizontalScrollBar }
                    if ($null -ne $originalVerticalScrollBar) { $activeWindow.DisplayVerticalScrollBar = $originalVerticalScrollBar }
                } catch {
                    # Best-effort cleanup only.
                }
            }
        }
    }

    return "callbacks=$passed"
}

function Get-StructuredTestResultsDetails {
    param([pscustomobject]$StructuredResults)

    return "tests=$($StructuredResults.Summary.Total), passed=$($StructuredResults.Summary.Passed), failed=$($StructuredResults.Summary.Failed)"
}

function Read-StructuredTestResults {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return $null
    }

    $summary = $null
    $results = @()

    foreach ($line in Get-Content $Path) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split "`t", 6
        if ($parts[0] -eq "SUMMARY" -and $parts.Count -ge 4) {
            $summary = [pscustomobject]@{
                Total = [int]$parts[1]
                Passed = [int]$parts[2]
                Failed = [int]$parts[3]
            }
        } elseif ($parts[0] -eq "RESULT" -and $parts.Count -ge 6) {
            $results += [pscustomobject]@{
                Name = $parts[1]
                Passed = [bool]::Parse($parts[2])
                DurationMs = [int]$parts[3]
                Category = $parts[4]
                Message = $parts[5]
            }
        }
    }

    return [pscustomobject]@{
        Summary = $summary
        Results = $results
    }
}

function Get-EnabledHeadlessCallbacks {
    param(
        [string]$ManifestPath
    )

    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath
    return @($manifest.Features | Where-Object {
        $_.PSObject.Properties.Name -contains "RuntimeTestMode" -and $_.RuntimeTestMode -eq "headless"
    })
}

function Set-RibbonUiErrors {
    param ([bool]$Enabled)
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Common\General"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    $val = if ($Enabled) { 1 } else { 0 }
    Set-ItemProperty -Path $regPath -Name "ShowErrors" -Value $val -Type DWord -Force
}



# --- Execution ---

if ($ListTests) {
    Write-Host "Discovering tests in modules..." -ForegroundColor Cyan
    $tests = Get-AllTestProcedures -SourceDir $modulesDir
    if ($tests.Count -eq 0) {
        Write-Host "No tests found." -ForegroundColor Yellow
        exit 0
    }
    
    $filtered = $tests
    if ($Filter) {
        $filterPatterns = $Filter -split ","
        $filtered = @()
        foreach ($t in $tests) {
            $matchesFilter = $false
            foreach ($pat in $filterPatterns) {
                $globPat = $pat.Trim()
                if ($globPat -and $globPat -notmatch '^\*|\*$') { $globPat = "*$globPat*" }
                if ($t.FullName -like $globPat) {
                    $matchesFilter = $true
                    break
                }
            }
            if ($matchesFilter) { $filtered += $t }
        }
    }
    
    Write-Host ""
    Write-Host "Available tests:" -ForegroundColor Cyan
    $grouped = $filtered | Group-Object Module
    foreach ($grp in $grouped | Sort-Object Name) {
        Write-Host "  [$($grp.Name)]" -ForegroundColor Yellow
        foreach ($t in $grp.Group | Sort-Object Procedure) {
            Write-Host "    - $($t.Procedure)" -ForegroundColor Gray
        }
    }
    Write-Host ""
    Write-Host "Total tests found: $($filtered.Count) (out of $($tests.Count) total)" -ForegroundColor Cyan
    exit 0
}

$sharedExcel = $null
$excelWasAlreadyOpen = $false

try {
    Invoke-Stage -Stage "runtime_tests" -Action {
        Write-Host "Starting Runtime Testing..." -ForegroundColor Cyan
        Reset-StructuredTestResults -Path $structuredTestResultsPath

        $session = Initialize-ExcelWorkbookSession -Purpose "runtime testing" -Visible:$Visible
        $testExcel = $session.Excel
        $testWorkbook = $session.Workbook
        $wasAlreadyOpen = $session.WasAlreadyOpen
        $script:excelWasAlreadyOpen = $wasAlreadyOpen
        $script:sharedExcel = $testExcel

        $watcher = $null
        $ribbonUiErrorsEnabled = $false
        $testExcelPid = 0
        $signalFile = Join-Path $env:TEMP ("BeaverRibbonSignal_" + [System.Guid]::NewGuid().ToString("N") + ".tmp")

        try {
            if ($CheckRibbon) {
                Set-RibbonUiErrors -Enabled $true
                $ribbonUiErrorsEnabled = $true

                $testExcelPid = Get-ExcelProcessId -ExcelApplication $testExcel

                if ($testExcelPid -gt 0) {
                    $null = Start-ExcelWindowWatcher -ExcelPid $testExcelPid -SignalPath $signalFile -TimeoutSeconds 20

                    Write-Host "Opening workbook and checking for Ribbon UI errors..."
                    if (-not $Visible) { $testExcel.Visible = $false }
                    $testExcel.DisplayAlerts = $true

                    $testWorkbook = $testExcel.Workbooks.Open($excelPath)

                    $testExcel.Visible = $true

                    $null = New-Item -Path $signalFile -ItemType File -Force
                    $ribbonError = [WindowScraper]::StopAndGetResult()
                    if (Test-Path $signalFile) {
                        Remove-Item -Path $signalFile -Force -ErrorAction SilentlyContinue
                    }

                    if (-not $Visible) { $testExcel.Visible = $false }
                    $testExcel.DisplayAlerts = $false

                    if ($ribbonError) {
                        Write-Host "  ERROR: Ribbon UI Validation failed." -ForegroundColor Red
                        $cleanError = $ribbonError -replace "\r\n+", " | " -replace "\s+", " "
                        Write-Host "  [Diagnostics] $cleanError" -ForegroundColor Yellow
                        throw "Ribbon UI Error: $cleanError"
                    }

                    Write-Host "  Ribbon UI loaded without errors." -ForegroundColor Green
                } else {
                    $testWorkbook = $testExcel.Workbooks.Open($excelPath)
                }
            } else {
                if ($null -eq $testWorkbook) {
                    $testWorkbook = $testExcel.Workbooks.Open($excelPath)
                }
                Write-Host "  Skipped Ribbon UI validation." -ForegroundColor Yellow
            }

            if ($testWorkbook.ReadOnly) {
                throw "The workbook '$excelPath' was opened as Read-Only. Please ensure that no other Excel process is locking the file."
            }

            # Prepare filter pattern
            $filterPattern = $Filter
            if ($filterPattern) {
                $parts = $filterPattern -split ","
                $wrappedParts = foreach ($part in $parts) {
                    $trimmed = $part.Trim()
                    if ($trimmed -and $trimmed -notmatch '^\*|\*$') {
                        "*$trimmed*"
                    } else {
                        $trimmed
                    }
                }
                $filterPattern = ($wrappedParts | Where-Object { $_ }) -join ","
            }

            $structuredResults = $null
            if (-not $SkipUnitTests) {
                # Temporarily hide Excel to prevent COM hangs and speed up test runner
                if (-not $Visible) { $testExcel.Visible = $false }

                Write-Host "Running internal unit tests..." -ForegroundColor Cyan
                try {
                    $retryCount = 0
                    $maxRetries = 5
                    $runCompleted = $false
                    while (-not $runCompleted -and $retryCount -lt $maxRetries) {
                        try {
                            if ($filterPattern) {
                                Write-Host "  Running tests matching filter: '$filterPattern'..." -ForegroundColor Cyan
                                $testExcel.Run("Lib_Tests.RunTestsFilter", $filterPattern)
                            } else {
                                $testExcel.Run("Lib_Tests.RunAllTests")
                            }
                            $runCompleted = $true
                        } catch {
                            $errMsg = $_.Exception.Message + " " + $_.Exception.InnerException.Message
                            if ($errMsg -match "0x800AC472" -or $errMsg -match "800ac472") {
                                $retryCount++
                                Write-Host "  Excel is busy (0x800AC472). Retrying in 1s ($retryCount/$maxRetries)..." -ForegroundColor Yellow
                                Start-Sleep -Seconds 1
                            } else {
                                throw $_
                            }
                        }
                    }
                    if (-not $runCompleted) {
                        throw "Failed to run tests because Excel remained busy."
                    }
                    Write-Host "  SUCCESS: Unit tests completed." -ForegroundColor Green
                } catch {
                    Write-Host "  FAILURE: Unit tests failed." -ForegroundColor Red
                    
                    $vbaLogPath = Join-Path $env:TEMP "BeaverAddin_$testExcelPid.log"
                    if (Test-Path $vbaLogPath) {
                        Write-Host "`n  --- BEAVER VBA DIAGNOSTIC LOGS ---" -ForegroundColor Yellow
                        Get-Content $vbaLogPath | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
                        Write-Host "  ----------------------------------`n" -ForegroundColor Yellow
                    }
                    throw "Unit tests failed: $($_.Exception.Message)"
                }

                $structuredResults = Read-StructuredTestResults -Path $structuredTestResultsPath
                try {
                    Assert-StructuredTestResults -StructuredResults $structuredResults -Path $structuredTestResultsPath | Out-Null
                } catch {
                    $vbaLogPath = Join-Path $env:TEMP "BeaverAddin_$testExcelPid.log"
                    if (Test-Path $vbaLogPath) {
                        Write-Host "`n  --- BEAVER VBA DIAGNOSTIC LOGS ---" -ForegroundColor Yellow
                        Get-Content $vbaLogPath | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
                        Write-Host "  ----------------------------------`n" -ForegroundColor Yellow
                    }
                    throw $_
                }
            } else {
                Write-Host "  Skipping unit tests (no VBA code changes)." -ForegroundColor Yellow
            }

            if (-not $Filter -and -not $SkipUnitTests) {
                $headlessCallbacks = Get-EnabledHeadlessCallbacks -ManifestPath $featureManifestPath
                Invoke-HeadlessCallbackTests -ExcelApplication $testExcel -Callbacks $headlessCallbacks | Out-Null
            } else {
                Write-Host "  Skipped headless callback tests (Filter or SkipUnitTests active)." -ForegroundColor Yellow
            }

            $testExcel.Visible = $true

            Write-Host "Runtime testing completed with structured test collection." -ForegroundColor Green
            if ($SkipUnitTests) {
                return "tests=0, passed=0, failed=0"
            } else {
                return (Get-StructuredTestResultsDetails -StructuredResults $structuredResults)
            }
        } finally {
            if ($testExcelPid -gt 0) {
                $null = [WindowScraper]::StopAndGetResult()
            }
            if (Test-Path $signalFile) {
                Remove-Item -Path $signalFile -Force -ErrorAction SilentlyContinue
            }
            if ($ribbonUiErrorsEnabled) {
                Set-RibbonUiErrors -Enabled $false
            }
            if ($testExcel -and $excelWasAlreadyOpen) {
                try {
                    $testExcel.Visible = $true
                    $testExcel.DisplayAlerts = $true
                } catch { }
            }
        }
    } | Out-Null

    Write-StageSummary
} catch {
    Stop-Script $_.Exception.Message
} finally {
    if (-not $global:BeaverOrchestratorActive -and $null -ne $sharedExcel) {
        if (-not $excelWasAlreadyOpen) {
            try {
                foreach ($wb in $sharedExcel.Workbooks) {
                    if ($wb.FullName -eq $excelPath) {
                        $wb.Close($false)
                    }
                }
            } catch { }
            try {
                $sharedExcel.Quit()
            } catch { }
        } else {
            try {
                $sharedExcel.Visible = $true
                $sharedExcel.DisplayAlerts = $true
            } catch { }
        }
        Release-ComObjectSafely $sharedExcel
        $sharedExcel = $null
    }
}
