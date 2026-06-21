# Script:   BuildSupport.ps1
# Purpose:  Shared helpers, paths, and Excel COM management for Build.ps1 and Test.ps1.
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Common Paths ---
$projectRoot = Split-Path $PSScriptRoot -Parent
$excelPath = Join-Path $projectRoot "Beaver Add-in.xlsm"
$modulesDir = Join-Path $projectRoot "Modules"
$desktopThisWorkbookCls = Join-Path $projectRoot "ThisWorkbook.cls"
$ribbonXmlPath = Join-Path $projectRoot "ribbon.xml"
$featureManifestPath = Join-Path $projectRoot "features.json"
$configPath = Join-Path $projectRoot "config.json"
$testManifestPath = Join-Path $modulesDir "Libraries\Lib_TestManifest.bas"
$commandRegistryPath = Join-Path $modulesDir "Infrastructure\Infra_CommandRegistry.bas"
$uiRibbonPath = Join-Path $modulesDir "UI\UI_Ribbon.bas"
$uiHotkeysPath = Join-Path $modulesDir "UI\UI_Hotkeys.bas"
$structuredTestResultsPath = Join-Path $env:TEMP "BeaverAddin.TestResults.tsv"

# --- Stage Execution Tracking ---
$script:StageResults = New-Object System.Collections.ArrayList

function Stop-Script {
    param(
        [string]$Message,
        [int]$ExitCode = 1
    )

    Write-StageSummary
    Write-Host $Message -ForegroundColor Red
    exit $ExitCode
}

function Add-StageResult {
    param(
        [string]$Stage,
        [string]$Status,
        [string]$Details = "",
        [double]$DurationMs = 0
    )

    [void]$script:StageResults.Add([pscustomobject]@{
        Stage = $Stage
        Status = $Status
        Details = $Details
        DurationMs = $DurationMs
        Timestamp = Get-Date
    })
}

function Write-StatusLine {
    param(
        [string]$Status,
        [string]$Stage,
        [string]$Details = "",
        [string]$Color = "Gray"
    )

    $detailText = if ([string]::IsNullOrWhiteSpace($Details)) { "" } else { " $Details" }
    Write-Host ("[{0,-6}] {1}{2}" -f $Status.ToUpperInvariant(), $Stage, $detailText) -ForegroundColor $Color
}

function Write-StageSummary {
    if ($script:StageResults.Count -eq 0) { return }

    Write-Host ""
    Write-Host "Stage Summary" -ForegroundColor Cyan
    foreach ($stage in $script:StageResults) {
        $color = if ($stage.Status -eq "success") { "Green" } elseif ($stage.Status -eq "skipped") { "Yellow" } else { "Red" }
        $durationText = if ($stage.DurationMs -gt 0) { " ({0:N1}s)" -f ($stage.DurationMs / 1000) } else { "" }
        $detailText = if ([string]::IsNullOrWhiteSpace($stage.Details)) { "" } else { " - $($stage.Details)" }
        Write-Host ("  [{0}] {1}{2}{3}" -f $stage.Status.ToUpper(), $stage.Stage, $durationText, $detailText) -ForegroundColor $color
    }

    $passed = @($script:StageResults | Where-Object { $_.Status -eq "success" }).Count
    $failed = @($script:StageResults | Where-Object { $_.Status -eq "failure" }).Count
    $skipped = @($script:StageResults | Where-Object { $_.Status -eq "skipped" }).Count
    $totalDurationMs = ($script:StageResults | Measure-Object -Property DurationMs -Sum).Sum
    if ($null -eq $totalDurationMs) { $totalDurationMs = 0 }

    Write-Host ""
    Write-Host ("Totals: passed={0}, failed={1}, skipped={2}, duration={3:N1}s" -f $passed, $failed, $skipped, ($totalDurationMs / 1000)) -ForegroundColor Cyan
}

function Invoke-Stage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Stage,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-StatusLine -Status "start" -Stage $Stage -Color "Cyan"

    try {
        $result = & $Action
        $stopwatch.Stop()

        $details = ""
        if ($null -ne $result) {
            if ($result -is [string]) {
                $details = $result
            } elseif ($result.PSObject.Properties.Name -contains "Details") {
                $details = [string]$result.Details
            }
        }

        Add-StageResult -Stage $Stage -Status "success" -Details $details -DurationMs $stopwatch.Elapsed.TotalMilliseconds
        Write-StatusLine -Status "pass" -Stage $Stage -Details ("({0:N1}s){1}" -f $stopwatch.Elapsed.TotalSeconds, $(if ([string]::IsNullOrWhiteSpace($details)) { "" } else { " $details" })) -Color "Green"
        return $result
    } catch {
        $stopwatch.Stop()
        $message = $_.Exception.Message
        Add-StageResult -Stage $Stage -Status "failure" -Details $message -DurationMs $stopwatch.Elapsed.TotalMilliseconds
        Write-StatusLine -Status "fail" -Stage $Stage -Details ("({0:N1}s) {1}" -f $stopwatch.Elapsed.TotalSeconds, $message) -Color "Red"
        throw
    }
}

function Add-SkippedStageResult {
    param(
        [string]$Stage,
        [string]$Details
    )

    Add-StageResult -Stage $Stage -Status "skipped" -Details $Details
    Write-StatusLine -Status "skip" -Stage $Stage -Details $Details -Color "Yellow"
}

# --- Excel COM Utilities ---

function Release-ComObjectSafely {
    param([object]$ComObject)

    if ($null -eq $ComObject) {
        return
    }

    try {
        if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
    } catch {
        # Best-effort cleanup only.
    }
}

function Get-ExcelProcessId {
    param(
        [Parameter(Mandatory = $true)]
        $ExcelApplication
    )

    if ($null -eq $ExcelApplication -or -not $ExcelApplication.Hwnd) {
        return 0
    }

    $excelPid = 0
    [WindowScraper]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$excelPid) | Out-Null
    return $excelPid
}

function Get-ExcelExecutablePath {
    $appPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
    )

    foreach ($path in $appPaths) {
        try {
            $key = Get-Item $path -ErrorAction Stop
            $exePath = $key.GetValue("")
            if ($exePath -and (Test-Path $exePath)) {
                return $exePath
            }
        } catch { }
    }

    return $null
}

function Remove-OrphanedExcelProcesses {
    $excelProcesses = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
    if ($excelProcesses.Count -eq 0) {
        return $false
    }

    $visibleExcel = @(
        $excelProcesses | Where-Object {
            $_.MainWindowHandle -ne 0 -or -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle)
        }
    )

    if ($visibleExcel.Count -gt 0) {
        return $false
    }

    Write-Host "  Found $($excelProcesses.Count) background Excel process(es) with no visible window. Cleaning up..." -ForegroundColor Yellow
    $stoppedAny = $false
    foreach ($process in $excelProcesses) {
        try {
            Stop-Process -Id $process.Id -Force -ErrorAction Stop
            $stoppedAny = $true
        } catch {
            Write-Warning "  Failed to stop orphaned Excel process $($process.Id): $($_.Exception.Message)"
        }
    }

    if ($stoppedAny) {
        Start-Sleep -Seconds 2
    }

    return $stoppedAny
}

function Test-FileLocked {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    try {
        $file = [System.IO.File]::Open($Path, 'Open', 'Write', 'None')
        $file.Close()
        $file.Dispose()
        return $false
    } catch {
        return $true
    }
}

function Clear-ExcelDisabledItems {
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems"
    if (Test-Path $regPath) {
        try {
            $disabledKey = Get-Item $regPath
            foreach ($valName in $disabledKey.GetValueNames()) {
                $valData = $disabledKey.GetValue($valName)
                if ($null -ne $valData -and $valData -is [byte[]]) {
                    $str = [System.Text.Encoding]::Unicode.GetString($valData)
                    if ($str -like "*beaver add-in.xlsm*") {
                        Write-Host "  Found Beaver Add-in in Excel DisabledItems. Enabling it..." -ForegroundColor Yellow
                        Remove-ItemProperty -Path $regPath -Name $valName -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        } catch {
            # Best effort
        }
    }
}

function Start-ExcelApplication {
    param(
        [string]$Purpose
    )

    Clear-ExcelDisabledItems

    try {
        return New-Object -ComObject Excel.Application
    } catch {
        $directComError = $_.Exception
    }

    $existingExcel = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
    if ($existingExcel.Count -gt 0) {
        if (Remove-OrphanedExcelProcesses) {
            try {
                return New-Object -ComObject Excel.Application
            } catch {
                $directComError = $_.Exception
                $existingExcel = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
            }
        }
    }

    if ($existingExcel.Count -gt 0) {
        throw "Failed to start Excel COM automation for ${Purpose}: $($directComError.Message)"
    }

    $excelExe = Get-ExcelExecutablePath
    if (-not $excelExe) {
        throw "Failed to start Excel COM automation for ${Purpose}: $($directComError.Message)"
    }

    try {
        $startedProcess = Start-Process -FilePath $excelExe -PassThru -ErrorAction Stop
    } catch {
        throw "Failed to start Excel COM automation for $Purpose. COM activation failed with '$($directComError.Message)', and launching EXCEL.EXE also failed: $($_.Exception.Message)"
    }

    $deadline = (Get-Date).AddSeconds(20)
    do {
        Start-Sleep -Milliseconds 500
        try {
            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            if ($excel -and $excel.Hwnd) {
                $hwndPid = 0
                [WindowScraper]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$hwndPid) | Out-Null
                if ($hwndPid -eq $startedProcess.Id) {
                    return $excel
                }
            }
        } catch { }
    } while ((Get-Date) -lt $deadline -and -not $startedProcess.HasExited)

    if (-not $startedProcess.HasExited) {
        Stop-Process -Id $startedProcess.Id -Force -ErrorAction SilentlyContinue
    }

    throw "Failed to start Excel COM automation for $Purpose. COM activation failed with '$($directComError.Message)', and Excel could not be attached after launching EXCEL.EXE."
}

function Get-FeatureManifest {
    param([string]$ManifestPath)
    if (-not (Test-Path $ManifestPath)) {
        throw "Feature manifest not found: $ManifestPath"
    }
    return Get-Content $ManifestPath -Raw | ConvertFrom-Json
}

function Get-ActiveExcelWorkbook {
    param(
        [string]$WorkbookPath
    )
    
    Clear-ExcelDisabledItems
    
    # 1. Check if Excel process is running first to avoid COM launch hangs
    $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    if (-not $excelProcesses) {
        return $null
    }
    
    # 2. Check if the workbook lock file exists to ensure it is actually open
    $fileName = Split-Path $WorkbookPath -Leaf
    $lockFile = Join-Path (Split-Path $WorkbookPath) ("~$" + $fileName)
    if (-not (Test-Path $lockFile)) {
        return $null
    }
    
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        if ($null -ne $excel) {
            foreach ($wb in $excel.Workbooks) {
                if ($wb.FullName -eq $WorkbookPath) {
                    return [pscustomobject]@{
                        Excel = $excel
                        Workbook = $wb
                        WasAlreadyOpen = $true
                    }
                }
            }
        }
    } catch {
        # Excel not running or workbook not open
    }
    return $null
}

$buildStatePath = Join-Path $PSScriptRoot ".build_state.json"

function Get-BuildState {
    if (Test-Path $buildStatePath) {
        try {
            $content = Get-Content $buildStatePath -Raw
            return $content | ConvertFrom-Json
        } catch {
            return $null
        }
    }
    return $null
}

function Save-BuildState {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$FileHashes
    )
    $projectRoot = Split-Path $PSScriptRoot -Parent
    $metadata = [ordered]@{}
    
    foreach ($relPath in $FileHashes.Keys) {
        $absPath = Join-Path $projectRoot $relPath
        if (Test-Path $absPath) {
            $file = Get-Item $absPath
            $meta = [ordered]@{
                Length = $file.Length
                LastWriteTime = $file.LastWriteTime.ToFileTime().ToString()
            }
            
            # Discovered tests from global cache or previous build state
            if ($null -ne $global:BeaverTestManifestCache -and $global:BeaverTestManifestCache.ContainsKey($relPath)) {
                $meta["Tests"] = $global:BeaverTestManifestCache[$relPath]
            } else {
                $oldState = Get-BuildState
                if ($null -ne $oldState -and $null -ne $oldState.Metadata -and $null -ne $oldState.Metadata.PSObject.Properties[$relPath]) {
                    $oldMeta = $oldState.Metadata.PSObject.Properties[$relPath].Value
                    if ($oldMeta.PSObject.Properties.Name -contains "Tests" -and $null -ne $oldMeta.Tests) {
                        $meta["Tests"] = @($oldMeta.Tests)
                    }
                }
            }
            
            # Lint status from global cache or previous build state
            if ($null -ne $global:BeaverLintStatusCache -and $global:BeaverLintStatusCache.ContainsKey($relPath)) {
                $meta["LintPassed"] = $global:BeaverLintStatusCache[$relPath]
            } else {
                $oldState = Get-BuildState
                if ($null -ne $oldState -and $null -ne $oldState.Metadata -and $null -ne $oldState.Metadata.PSObject.Properties[$relPath]) {
                    $oldMeta = $oldState.Metadata.PSObject.Properties[$relPath].Value
                    if ($oldMeta.PSObject.Properties.Name -contains "LintPassed" -and $null -ne $oldMeta.LintPassed) {
                        $meta["LintPassed"] = [bool]$oldMeta.LintPassed
                    }
                }
            }
            
            $metadata[$relPath] = $meta
        }
    }
    
    $state = [ordered]@{
        LastBuildTime = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssK")
        Files = $FileHashes
        Metadata = $metadata
        ManifestStructuralHash = (Get-ManifestStructuralHash -Path $featureManifestPath)
    }
    $stateJson = $state | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($buildStatePath, $stateJson, [System.Text.Encoding]::ASCII)
}

function Set-BuildStateTestsPassed {
    param(
        [bool]$Passed
    )
    $buildState = Get-BuildState
    if ($null -ne $buildState) {
        $buildState | Add-Member -NotePropertyName TestsPassed -NotePropertyValue $Passed -Force
        $stateJson = $buildState | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($buildStatePath, $stateJson, [System.Text.Encoding]::ASCII)
    }
}

function Get-FileHashOptimized {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) { return "" }
    
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $stream = [System.IO.File]::OpenRead($FilePath)
    $hashBytes = $md5.ComputeHash($stream)
    $stream.Close()
    $stream.Dispose()
    $md5.Dispose()

    $sb = [System.Text.StringBuilder]::new()
    foreach ($b in $hashBytes) {
        [void]$sb.Append($b.ToString("x2"))
    }
    return $sb.ToString().ToUpperInvariant()
}

function Get-SourceFileHashes {
    param(
        [switch]$Force
    )

    if (-not $Force -and $null -ne $global:BeaverSourceHashes) {
        return $global:BeaverSourceHashes
    }

    $hashes = @{}
    $buildState = Get-BuildState
    $projectRoot = Split-Path $PSScriptRoot -Parent
    
    $resolveHash = {
        param($filePath, $relPath)
        if (-not (Test-Path $filePath)) { return "" }
        $file = Get-Item $filePath
        
        if ($null -ne $buildState -and $null -ne $buildState.Metadata -and $null -ne $buildState.Metadata.PSObject.Properties[$relPath]) {
            $meta = $buildState.Metadata.PSObject.Properties[$relPath].Value
            $currentSize = $file.Length
            $currentMtime = $file.LastWriteTime.ToFileTime().ToString()
            if ($null -ne $meta.Length -and $null -ne $meta.LastWriteTime -and 
                $meta.Length -eq $currentSize -and $meta.LastWriteTime -eq $currentMtime) {
                if ($null -ne $buildState.Files -and $null -ne $buildState.Files.PSObject.Properties[$relPath]) {
                    return $buildState.Files.PSObject.Properties[$relPath].Value
                }
            }
        }
        return Get-FileHashOptimized -FilePath $filePath
    }
    
    # Manifest
    if (Test-Path $featureManifestPath) {
        $hashes["features.json"] = & $resolveHash $featureManifestPath "features.json"
    }
    
    # ThisWorkbook
    if (Test-Path $desktopThisWorkbookCls) {
        $hashes["ThisWorkbook.cls"] = & $resolveHash $desktopThisWorkbookCls "ThisWorkbook.cls"
    }
    
    # Modules
    if (Test-Path $modulesDir) {
        $vbaFiles = Get-ChildItem -Path $modulesDir -Recurse | Where-Object { $_.Extension -match "\.(bas|cls|frm)$" }
        foreach ($file in $vbaFiles) {
            $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace("\", "/")
            $hashes[$relPath] = & $resolveHash $file.FullName $relPath
            
            # If it's a form, also include the companion FRX file hash if it exists
            if ($file.Extension -eq ".frm") {
                $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
                if (Test-Path $frxPath) {
                    $frxRelPath = $frxPath.Substring($projectRoot.Length + 1).Replace("\", "/")
                    $hashes[$frxRelPath] = & $resolveHash $frxPath $frxRelPath
                }
            }
        }
    }
    
    if ($global:BeaverOrchestratorActive) {
        $global:BeaverSourceHashes = $hashes
    }
    
    return $hashes
}

function Get-AllTestProcedures {
    param([string]$SourceDir)
    
    $testProcedures = @()
    if (-not (Test-Path $SourceDir)) { return $testProcedures }
    
    $projectRoot = Split-Path $PSScriptRoot -Parent
    $buildState = Get-BuildState
    $global:BeaverTestManifestCache = @{}
    
    $moduleFiles = @(Get-ChildItem -Path $SourceDir -Filter *.bas -Recurse)
    foreach ($file in $moduleFiles) {
        if ($file.Name -eq "Lib_TestManifest.bas") { continue }
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace("\", "/")
        
        # Check cache
        $cachedTests = $null
        if ($null -ne $buildState -and $null -ne $buildState.Metadata -and $null -ne $buildState.Metadata.PSObject.Properties[$relPath]) {
            $meta = $buildState.Metadata.PSObject.Properties[$relPath].Value
            $currentSize = $file.Length
            $currentMtime = $file.LastWriteTime.ToFileTime().ToString()
            if ($null -ne $meta.Length -and $null -ne $meta.LastWriteTime -and 
                $meta.Length -eq $currentSize -and $meta.LastWriteTime -eq $currentMtime -and 
                $null -ne $meta.Tests) {
                $cachedTests = @($meta.Tests)
            }
        }
        
        if ($null -ne $cachedTests) {
            $global:BeaverTestManifestCache[$relPath] = $cachedTests
            foreach ($testName in $cachedTests) {
                $testProcedures += [pscustomobject]@{
                    Module = $moduleName
                    Procedure = $testName
                    FullName = "$moduleName.$testName"
                    File = $file.FullName
                }
            }
            continue
        }
        
        # Cache miss
        $fileTests = @()
        $matches = Select-String -Path $file.FullName -Pattern '^\s*Public Sub (Test_[A-Za-z0-9_]+)\s*\(' -ErrorAction SilentlyContinue
        foreach ($match in $matches) {
            $testName = $match.Matches[0].Groups[1].Value
            $fileTests += $testName
            $testProcedures += [pscustomobject]@{
                Module = $moduleName
                Procedure = $testName
                FullName = "$moduleName.$testName"
                File = $file.FullName
            }
        }
        $global:BeaverTestManifestCache[$relPath] = $fileTests
    }
    return $testProcedures
}


# --- WindowScraper C# Code ---
$script:scraperCode = @"
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

    private static string _scrapedText = "";
    private static System.Threading.Thread _thread = null;
    private static bool _stopRequested = false;

    public static void Start(int processId, int timeoutSeconds, string signalFilePath) {
        _scrapedText = "";
        _stopRequested = false;
        _thread = new System.Threading.Thread(() => {
            _scrapedText = ScrapeAndClose(processId, timeoutSeconds, signalFilePath);
        });
        _thread.IsBackground = true;
        _thread.Start();
    }

    public static string StopAndGetResult() {
        _stopRequested = true;
        if (_thread != null && _thread.IsAlive) {
            _thread.Join(2000);
        }
        return _scrapedText;
    }

    public static string ScrapeAndClose(int processId, int timeoutSeconds, string signalFilePath) {
        var result = new StringBuilder();
        var seenTexts = new HashSet<string>();
        var startTime = DateTime.Now;

        while ((DateTime.Now - startTime).TotalSeconds < timeoutSeconds && !_stopRequested) {
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

Add-Type -TypeDefinition $script:scraperCode -ErrorAction SilentlyContinue

function Start-ExcelWindowWatcher {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ExcelPid,
        [Parameter(Mandatory = $true)]
        [string]$SignalPath,
        [int]$TimeoutSeconds = 20
    )

    [WindowScraper]::Start($ExcelPid, $TimeoutSeconds, $SignalPath)
    return $null
}

function Initialize-ExcelWorkbookSession {
    param(
        [string]$Purpose,
        [switch]$Visible
    )

    $excelPath = Join-Path $projectRoot "Beaver Add-in.xlsm"
    $wasAlreadyOpen = $false
    $excel = $null
    $workbook = $null

    # 1. Reuse global Excel instance if orchestrator is active
    if ($global:BeaverOrchestratorActive -and $null -ne $global:BeaverSharedExcel) {
        try {
            $excel = $global:BeaverSharedExcel
            $wbs = $excel.Workbooks
            foreach ($wb in $excel.Workbooks) {
                if ($wb.FullName -eq $excelPath) {
                    $workbook = $wb
                    break
                }
            }
            $wasAlreadyOpen = $true
            Write-Host "  Reusing persistent Excel COM session." -ForegroundColor Green
        } catch {
            Write-Warning "  Persistent Excel COM session is unresponsive. Starting a new session..."
            $global:BeaverSharedExcel = $null
            $excel = $null
        }
    }

    # 2. Start a new session if needed
    if ($null -eq $excel) {
        $retryCount = 0
        $maxRetries = 10
        $ready = $false

        while (-not $ready -and $retryCount -lt $maxRetries) {
            try {
                $activeWbInfo = Get-ActiveExcelWorkbook -WorkbookPath $excelPath
                if ($null -ne $activeWbInfo) {
                    $excel = $activeWbInfo.Excel
                    $workbook = $activeWbInfo.Workbook
                    $wasAlreadyOpen = $true
                } else {
                    $excel = Start-ExcelApplication -Purpose $Purpose
                }
                $wbs = $excel.Workbooks
                $ready = $true
            } catch {
                $retryCount++
                Write-Host "  Excel is busy or initializing. Retrying in 1s ($retryCount/$maxRetries)..." -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
        }

        if (-not $ready) {
            throw "Excel COM interface remained busy or failed to initialize."
        }

        if ($wasAlreadyOpen) {
            Write-Host "  Attached to active Excel instance." -ForegroundColor Green
        }

        if ($global:BeaverOrchestratorActive) {
            $global:BeaverSharedExcel = $excel
            $global:BeaverExcelWasAlreadyOpen = $wasAlreadyOpen
        }
    }

    # 3. Open workbook if not loaded
    if ($null -eq $workbook) {
        foreach ($wb in $excel.Workbooks) {
            if ($wb.FullName -eq $excelPath) {
                $workbook = $wb
                break
            }
        }

        if ($null -eq $workbook) {
            if (-not $Visible) {
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
            } else {
                $excel.Visible = $true
            }

            $opened = $false
            $openRetry = 0
            while (-not $opened -and $openRetry -lt 5) {
                try {
                    $excel.EnableEvents = $false
                    $workbook = $excel.Workbooks.Open($excelPath)
                    $excel.EnableEvents = $true
                    $opened = $true
                } catch {
                    $openRetry++
                    $errMsg = $_.Exception.Message
                    Write-Host "  Failed to open workbook: $errMsg. Retrying in 1s ($openRetry/5)..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }
            if (-not $opened) {
                throw "Failed to open workbook: $excelPath"
            }
        } else {
            $wasAlreadyOpen = $true
        }
    }

    # 4. Check programmatic VBE access
    try {
        $null = $excel.VBE
    } catch {
        throw "Programmatic access to the Visual Basic Project is not trusted. Please enable it in Excel under File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> 'Trust access to the VBA project object model'."
    }

    # 5. Check read-only state
    if ($workbook.ReadOnly) {
        throw "The workbook '$excelPath' was opened as Read-Only. Please ensure that no other Excel process is locking the file."
    }

    # 6. Apply visibility
    if ($Visible) {
        $excel.Visible = $true
    }

    return [pscustomobject]@{
        Excel = $excel
        Workbook = $workbook
        WasAlreadyOpen = $wasAlreadyOpen
    }
}

function Get-ManifestStructuralHash {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return "" }
    
    $manifest = $null
    try {
        $manifest = Get-Content $Path -Raw | ConvertFrom-Json
    } catch {
        return ""
    }

    $getSafeProp = {
        param($obj, $name)
        if ($null -ne $obj -and $null -ne $obj.PSObject.Properties[$name]) {
            return $obj.$name
        }
        return $null
    }

    $tabs = @()
    if ($null -ne $manifest) {
        $mTabs = & $getSafeProp $manifest "Tabs"
        if ($null -eq $mTabs) {
            $mTabs = & $getSafeProp $manifest "Tab"
        }
        if ($null -ne $mTabs) {
            $tabs = @($mTabs | ForEach-Object { & $getSafeProp $_ "Id" })
        }
    }

    $groups = @()
    if ($null -ne $manifest -and $null -ne $manifest.Groups) {
        $groups = @($manifest.Groups | ForEach-Object {
            [pscustomobject]@{
                Id = & $getSafeProp $_ "Id"
                TabId = & $getSafeProp $_ "TabId"
                Features = @($_.Features)
            }
        })
    }

    $features = @()
    if ($null -ne $manifest -and $null -ne $manifest.Features) {
        $features = @($manifest.Features | ForEach-Object {
            $menuItems = & $getSafeProp $_ "MenuItems"
            [pscustomobject]@{
                ControlId = & $getSafeProp $_ "ControlId"
                Type = & $getSafeProp $_ "Type"
                OnAction = & $getSafeProp $_ "OnAction"
                Macro = & $getSafeProp $_ "Macro"
                CommandName = & $getSafeProp $_ "CommandName"
                CommandClass = & $getSafeProp $_ "CommandClass"
                RuntimeTestMode = & $getSafeProp $_ "RuntimeTestMode"
                MenuItems = if ($null -ne $menuItems) { @($menuItems) } else { $null }
            }
        })
    }

    $hotkeys = @()
    if ($null -ne $manifest -and $null -ne $manifest.Hotkeys) {
        $hotkeys = @($manifest.Hotkeys | ForEach-Object {
            [pscustomobject]@{
                Key = & $getSafeProp $_ "Key"
                Macro = & $getSafeProp $_ "Macro"
                CommandName = & $getSafeProp $_ "CommandName"
            }
        })
    }

    $canonicalStruct = [pscustomobject]@{
        Tabs = $tabs
        Groups = $groups
        Features = $features
        Hotkeys = $hotkeys
    }

    $json = $canonicalStruct | ConvertTo-Json -Depth 10

    $md5 = [System.Security.Cryptography.MD5]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $hashBytes = $md5.ComputeHash($bytes)
    $md5.Dispose()

    $sb = [System.Text.StringBuilder]::new()
    foreach ($b in $hashBytes) {
        [void]$sb.Append($b.ToString("x2"))
    }
    return $sb.ToString().ToUpperInvariant()
}


