# Script:   Test.ps1
# Purpose:  Runs runtime verification including Ribbon UI checks and internal unit tests.
# ==============================================================================

. (Join-Path $PSScriptRoot "BuildSupport.ps1")

# --- Helper Functions ---

function Reset-StructuredTestResults {
    param([string]$Path)

    if (Test-Path $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

function Write-StructuredTestResultsSummary {
    param([pscustomobject]$StructuredResults)

    Write-Host ("  Structured test results: total={0}, passed={1}, failed={2}" -f $StructuredResults.Summary.Total, $StructuredResults.Summary.Passed, $StructuredResults.Summary.Failed) -ForegroundColor Cyan

    foreach ($result in @($StructuredResults.Results)) {
        $resultColor = if ($result.Passed) { "Green" } else { "Yellow" }
        $resultStatus = if ($result.Passed) { "PASS" } else { "FAIL" }
        $messageText = if ([string]::IsNullOrWhiteSpace($result.Message)) { "" } else { " - $($result.Message)" }
        Write-Host ("    [{0}] {1} ({2} ms, {3}){4}" -f $resultStatus, $result.Name, $result.DurationMs, $result.Category, $messageText) -ForegroundColor $resultColor
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

# --- WindowScraper C# Code ---

$scraperCode = @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class WindowScraper {
    public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumThreadDelegate lpfn, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumThreadDelegate lpfn, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

    public static string ScrapeAndClose(int processId, int timeoutSeconds, string signalFilePath) {
        var result = new StringBuilder();
        var seenTexts = new HashSet<string>();
        var startTime = DateTime.Now;

        while ((DateTime.Now - startTime).TotalSeconds < timeoutSeconds) {
            if (!string.IsNullOrEmpty(signalFilePath) && System.IO.File.Exists(signalFilePath)) {
                break;
            }
            EnumWindows((hWnd, lParam) => {
                int windowPid;
                GetWindowThreadProcessId(hWnd, out windowPid);
                if (windowPid == processId) {
                    var title = new StringBuilder(256);
                    GetWindowText(hWnd, title, 256);
                    string sTitle = title.ToString();
                    
                    if (sTitle.Contains("Microsoft Excel") || 
                        sTitle.Contains("Custom UI") || 
                        sTitle.Contains("Runtime Error") ||
                        (sTitle.Contains("Microsoft Visual Basic") && !sTitle.Contains("for Applications"))) {
                        
                        bool foundNewText = false;
                        EnumChildWindows(hWnd, (hChild, lChild) => {
                            var text = new StringBuilder(1024);
                            GetWindowText(hChild, text, 1024);
                            var sText = text.ToString().Trim();
                            if (sText.Length > 0 && 
                                !sText.Equals("OK", StringComparison.OrdinalIgnoreCase) && 
                                !sText.Equals("Cancel", StringComparison.OrdinalIgnoreCase) && 
                                !sText.Equals("Close", StringComparison.OrdinalIgnoreCase) &&
                                !sText.Equals("Help", StringComparison.OrdinalIgnoreCase) &&
                                !sText.StartsWith("MsoDock", StringComparison.OrdinalIgnoreCase) &&
                                !sText.Equals("Standard", StringComparison.OrdinalIgnoreCase) &&
                                !sText.Equals("Menu Bar", StringComparison.OrdinalIgnoreCase) &&
                                !sText.Contains("VBAProject") &&
                                !sText.Contains("Project Window") &&
                                !sText.Contains("Properties") &&
                                !seenTexts.Contains(sText)) {
                                result.AppendLine(sText);
                                seenTexts.Add(sText);
                                foundNewText = true;
                            }
                            return true;
                        }, IntPtr.Zero);

                        if (foundNewText || sTitle.Contains("Visual Basic")) {
                            PostMessage(hWnd, 0x0010, IntPtr.Zero, IntPtr.Zero);
                        }
                    }
                }
                return true;
            }, IntPtr.Zero);
            
            System.Threading.Thread.Sleep(500);
        }
        return result.ToString().Trim();
    }
}
"@
Add-Type -TypeDefinition $scraperCode -ErrorAction SilentlyContinue

# --- Execution ---

$sharedExcel = $null

try {
    Invoke-Stage -Stage "runtime_tests" -Action {
        Write-Host "Starting Runtime Testing..." -ForegroundColor Cyan
        Reset-StructuredTestResults -Path $structuredTestResultsPath

        if ($null -eq $sharedExcel) {
            $script:sharedExcel = Start-ExcelApplication -Purpose "runtime testing"
        }
        $testExcel = $sharedExcel
        $testExcel.Visible = $false
        $testExcel.DisplayAlerts = $false
        $testWorkbook = $null
        $watcher = $null
        $ribbonUiErrorsEnabled = $false
        $testExcelPid = 0
        $signalFile = Join-Path $env:TEMP ("BeaverRibbonSignal_" + [System.Guid]::NewGuid().ToString("N") + ".tmp")

        try {
            Set-RibbonUiErrors -Enabled $true
            $ribbonUiErrorsEnabled = $true

            $testExcelPid = Get-ExcelProcessId -ExcelApplication $testExcel

            if ($testExcelPid -gt 0) {
                $watcher = Start-Job -ScriptBlock {
                    param($ProcessIdToScrape, $code, $SignalPath)
                    Add-Type -TypeDefinition $code -ErrorAction SilentlyContinue
                    return [WindowScraper]::ScrapeAndClose($ProcessIdToScrape, 20, $SignalPath)
                } -ArgumentList $testExcelPid, $scraperCode, $signalFile

                Write-Host "Opening workbook and checking for Ribbon UI errors..."
                $testExcel.Visible = $false
                $testExcel.DisplayAlerts = $true

                $testWorkbook = $testExcel.Workbooks.Open($excelPath)

                $testExcel.Visible = $true

                $null = New-Item -Path $signalFile -ItemType File -Force
                $ribbonError = Receive-Job -Job $watcher -Wait
                Remove-Job $watcher -Force
                $watcher = $null
                if (Test-Path $signalFile) {
                    Remove-Item -Path $signalFile -Force -ErrorAction SilentlyContinue
                }

                $testExcel.Visible = $false
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

            Write-Host "Running internal unit tests..." -ForegroundColor Cyan
            try {
                $retryCount = 0
                $maxRetries = 5
                $runCompleted = $false
                while (-not $runCompleted -and $retryCount -lt $maxRetries) {
                    try {
                        $testExcel.Run("Lib_Tests.RunAllTests")
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

            $headlessCallbacks = Get-EnabledHeadlessCallbacks -ManifestPath $featureManifestPath
            Invoke-HeadlessCallbackTests -ExcelApplication $testExcel -Callbacks $headlessCallbacks | Out-Null

            Write-Host "Runtime testing completed with structured test collection." -ForegroundColor Green
            return (Get-StructuredTestResultsDetails -StructuredResults $structuredResults)
        } finally {
            if ($watcher) {
                Remove-Job $watcher -Force -ErrorAction SilentlyContinue
            }
            if (Test-Path $signalFile) {
                Remove-Item -Path $signalFile -Force -ErrorAction SilentlyContinue
            }
            if ($ribbonUiErrorsEnabled) {
                Set-RibbonUiErrors -Enabled $false
            }
            if ($testWorkbook) {
                try { $testWorkbook.Close($false) } catch { }
            }
            Release-ComObjectSafely $testWorkbook
            $testWorkbook = $null
        }
    } | Out-Null

    Write-StageSummary
} catch {
    Stop-Script $_.Exception.Message
} finally {
    if ($null -ne $sharedExcel) {
        Write-Host "Closing Excel application..." -ForegroundColor Gray
        try { $sharedExcel.Quit() } catch { }
        Release-ComObjectSafely $sharedExcel
        $sharedExcel = $null
    }
}
