[CmdletBinding()]
param(
    [switch]$CheckRibbon,
    [string]$Filter,
    [switch]$ListTests,
    [switch]$Visible,
    [switch]$SkipUnitTests,
    [switch]$FailedOnly
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
    
    # 1. Try to find the file from cached build state metadata to avoid scanning all files
    $targetFile = $null
    $targetRelPath = $null
    $buildState = Get-BuildState
    if ($null -ne $buildState -and $null -ne $buildState.Metadata) {
        foreach ($prop in $buildState.Metadata.PSObject.Properties) {
            $meta = $prop.Value
            if ($null -ne $meta -and $meta.PSObject.Properties['Tests'] -and $null -ne $meta.Tests -and $meta.Tests -contains $cleanProcName) {
                $targetRelPath = $prop.Name
                $targetFile = Join-Path $projectRoot $targetRelPath
                break
            }
        }
    }

    # 2. If cached location found, scan only that single file
    if ($null -ne $targetFile -and (Test-Path $targetFile)) {
        $lines = [System.IO.File]::ReadAllLines($targetFile)
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match "^\s*Public Sub\s+$cleanProcName\s*\(") {
                return [pscustomobject]@{
                    File = $targetFile
                    LineNumber = $i + 1
                    RelativePath = $targetRelPath
                }
            }
        }
    }

    # 3. Fallback: scan all .bas files if cache miss
    $vbaFiles = Get-ChildItem -Path $modulesDir -Filter *.bas -Recurse
    foreach ($file in $vbaFiles) {
        $lines = [System.IO.File]::ReadAllLines($file.FullName)
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
            if ($null -ne $global:BeaverBuildLog) {
                $global:BeaverBuildLog.testResults.headlessCallbacks.passedCount++
            }
            $passed++
        } catch {
            Write-Host "    [FAIL] $callbackName - $($_.Exception.Message)" -ForegroundColor Red
            if ($null -ne $global:BeaverBuildLog) {
                $global:BeaverBuildLog.testResults.headlessCallbacks.status = "failure"
                [void]$global:BeaverBuildLog.testResults.headlessCallbacks.failures.Add([ordered]@{
                    name = $callbackName
                    error = $_.Exception.Message
                })
            }
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
                Release-ComObjectSafely $activeWindow
                $activeWindow = $null
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
        $script:workbookWasAlreadyOpen = $session.WorkbookWasAlreadyOpen
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
                        if ($null -ne $global:BeaverBuildLog) {
                            $global:BeaverBuildLog.testResults.ribbonValidation.status = "failure"
                            $global:BeaverBuildLog.testResults.ribbonValidation.error = $cleanError
                        }
                        throw "Ribbon UI Error: $cleanError"
                    }

                    Write-Host "  Ribbon UI loaded without errors." -ForegroundColor Green
                    if ($null -ne $global:BeaverBuildLog) {
                        $global:BeaverBuildLog.testResults.ribbonValidation.status = "success"
                    }
                } else {
                    $testWorkbook = $testExcel.Workbooks.Open($excelPath)
                }
            } else {
                if ($null -eq $testWorkbook) {
                    $testWorkbook = $testExcel.Workbooks.Open($excelPath)
                }
                Write-Host "  Skipped Ribbon UI validation." -ForegroundColor Yellow
                if ($null -ne $global:BeaverBuildLog) {
                    $global:BeaverBuildLog.testResults.ribbonValidation.status = "skipped"
                }
            }

            if ($testWorkbook.ReadOnly) {
                throw "The workbook '$excelPath' was opened as Read-Only. Please ensure that no other Excel process is locking the file."
            }

            # Resolve FailedOnly if active
            if ($FailedOnly) {
                $logPath = Join-Path $PSScriptRoot "build_log.json"
                if (Test-Path $logPath) {
                    try {
                        $logContent = Get-Content $logPath -Raw | ConvertFrom-Json
                        if ($null -ne $logContent -and $null -ne $logContent.testResults -and $null -ne $logContent.testResults.unitTests -and $null -ne $logContent.testResults.unitTests.failures) {
                            $failedTestNames = @()
                            foreach ($fail in $logContent.testResults.unitTests.failures) {
                                if ($null -ne $fail.name) {
                                    $failedTestNames += $fail.name
                                }
                            }
                            if ($failedTestNames.Count -gt 0) {
                                $wrappedFailed = foreach ($name in $failedTestNames) {
                                    "*$name*"
                                }
                                $Filter = $wrappedFailed -join ","
                                Write-Host "Smart Retry: Found $($failedTestNames.Count) previous failure(s). Filtering tests to: $Filter" -ForegroundColor Yellow
                            } else {
                                Write-Host "Smart Retry: No previous test failures recorded. Skipping tests." -ForegroundColor Green
                                $SkipUnitTests = $true
                            }
                        } else {
                            Write-Host "Smart Retry: No previous test failures recorded in log. Skipping tests." -ForegroundColor Green
                            $SkipUnitTests = $true
                        }
                    } catch {
                        Write-Warning "Failed to parse build_log.json for Smart Retry: $($_.Exception.Message)"
                    }
                } else {
                    Write-Host "Smart Retry: No build log found. Skipping tests." -ForegroundColor Green
                    $SkipUnitTests = $true
                }
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
                $unitTestStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                try {
                    $retryCount = 0
                    $maxRetries = 5
                    $runCompleted = $false
                    while (-not $runCompleted -and $retryCount -lt $maxRetries) {
                        try {
                            if ($filterPattern) {
                                Write-Host "  Running tests matching filter: '$filterPattern'..." -ForegroundColor Cyan
                                $testExcel.Run("Test_Runner.RunTestsFilter", $filterPattern)
                            } else {
                                $testExcel.Run("Test_Runner.RunAllTests")
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
                    throw "Unit tests failed: $_"
                } finally {
                    $unitTestStopwatch.Stop()
                }

                $structuredResults = Read-StructuredTestResults -Path $structuredTestResultsPath
                if ($null -ne $global:BeaverBuildLog) {
                    $global:BeaverBuildLog.testResults.runTests = $true
                    $global:BeaverBuildLog.testResults.filter = $Filter
                    if ($null -ne $structuredResults) {
                        $global:BeaverBuildLog.testResults.unitTests.total = $structuredResults.Summary.Total
                        $global:BeaverBuildLog.testResults.unitTests.passed = $structuredResults.Summary.Passed
                        $global:BeaverBuildLog.testResults.unitTests.failed = $structuredResults.Summary.Failed
                        
                        $failures = New-Object System.Collections.ArrayList
                        foreach ($res in $structuredResults.Results) {
                            if (-not $res.Passed) {
                                [void]$failures.Add([ordered]@{
                                    name = $res.Name
                                    category = $res.Category
                                    message = $res.Message
                                    durationMs = $res.DurationMs
                                })
                            }
                        }
                        $global:BeaverBuildLog.testResults.unitTests.failures = $failures
                        
                        $testDur = [int]$unitTestStopwatch.Elapsed.TotalMilliseconds
                        $global:BeaverBuildLog.testResults.unitTests.durationMs = $testDur
                    }
                }
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
            if ($null -ne $testWorkbook) {
                try {
                    if ($global:BeaverKeepAliveActive) {
                        $testWorkbook.Saved = $true
                    } else {
                        $testWorkbook.Close($false)
                    }
                } catch {}
            }
            Release-ComObjectSafely $testWorkbook
            $testWorkbook = $null
            if ($testExcel) {
                try {
                    if ($testExcel.Calculation -ne -4105) {
                        $testExcel.Calculation = -4105 # xlCalculationAutomatic
                        Write-Host "  Restored Excel calculation option to Automatic." -ForegroundColor Green
                    }
                } catch { }
                if ($excelWasAlreadyOpen) {
                    try {
                        $testExcel.Visible = $true
                        $testExcel.DisplayAlerts = $true
                    } catch { }
                }
            }
            if (-not $global:BeaverOrchestratorActive) {
                Release-ComObjectSafely $testExcel
                $testExcel = $null
            }
        }
    } | Out-Null

    Write-StageSummary
    Save-BuildLog -Status "success"
} catch {
    Stop-Script $_.Exception.Message
} finally {
    if (-not $global:BeaverOrchestratorActive -and $null -ne $sharedExcel) {
        Close-ExcelWorkbookSession -Excel $sharedExcel `
                                   -WasAlreadyOpen $excelWasAlreadyOpen `
                                   -KeepAlive $global:BeaverKeepAliveActive `
                                   -WorkbookPath $excelPath
        $sharedExcel = $null
    }
    Clear-AccumulatedLogs
}
