# Script:   ComUtils.ps1
# Purpose:  Excel COM lifecycle management, process recovery, and window watcher for Beaver build pipeline.
# ==============================================================================

Set-StrictMode -Version Latest

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

    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(
        IntPtr hwnd, 
        uint dwId, 
        ref Guid riid, 
        [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    private static Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
    private const uint OBJID_NATIVEOM = 0xFFFFFFF0;

    public static void NudgeProcess(int processId) {
        EnumWindows((hWnd, lParam) => {
            int windowPid;
            GetWindowThreadProcessId(hWnd, out windowPid);
            if (windowPid == processId) {
                var className = new StringBuilder(256);
                GetClassName(hWnd, className, 256);
                if (className.ToString().Equals("XLMAIN")) {
                    ShowWindow(hWnd, 4); // SW_SHOWNOACTIVATE
                    SetForegroundWindow(hWnd);
                    return false; // stop
                }
            }
            return true;
        }, IntPtr.Zero);
    }

    public static object GetExcelObject(int processId) {
        object excelApp = null;
        EnumWindows((hWnd, lParam) => {
            int windowPid;
            GetWindowThreadProcessId(hWnd, out windowPid);
            if (windowPid == processId) {
                var className = new StringBuilder(256);
                GetClassName(hWnd, className, 256);
                if (className.ToString().Equals("XLMAIN")) {
                    EnumChildWindows(hWnd, (hChild, lChild) => {
                        var childClass = new StringBuilder(256);
                        GetClassName(hChild, childClass, 256);
                        if (childClass.ToString().Equals("EXCEL7")) {
                            object ppvObject;
                            int res = AccessibleObjectFromWindow(hChild, OBJID_NATIVEOM, ref IID_IDispatch, out ppvObject);
                            if (res == 0 && ppvObject != null) {
                                excelApp = ppvObject.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, ppvObject, null);
                                return false; // stop child enum
                            }
                        }
                        return true;
                    }, IntPtr.Zero);
                }
            }
            return excelApp == null; // continue enum if not found
        }, IntPtr.Zero);
        return excelApp;
    }

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

    # Find ONLY the background/hidden Excel processes
    $backgroundExcel = @(
        $excelProcesses | Where-Object {
            $_.MainWindowHandle -eq 0 -and [string]::IsNullOrWhiteSpace($_.MainWindowTitle)
        }
    )

    if ($backgroundExcel.Count -eq 0) {
        return $false
    }

    Write-Host "  Found $($backgroundExcel.Count) background Excel process(es) with no visible window. Cleaning up..." -ForegroundColor Yellow
    $stoppedAny = $false
    foreach ($process in $backgroundExcel) {
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

function Close-ExcelWorkbookSession {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Excel,

        [Parameter(Mandatory = $true)]
        [bool]$WasAlreadyOpen,

        [Parameter(Mandatory = $true)]
        [bool]$KeepAlive,

        [string]$WorkbookPath = $null,

        [switch]$SaveChanges,

        [switch]$ForceQuit
    )

    if ($null -eq $Excel) {
        return
    }

    # Failsafe: Ensure Excel calculation mode is set back to Automatic if leaving Excel open
    try {
        if ($Excel.Calculation -ne -4105) {
            $Excel.Calculation = -4105 # xlCalculationAutomatic
            Write-Host "  Restored Excel calculation option to Automatic." -ForegroundColor Green
        }
    } catch {}

    $excelPid = 0
    try {
        $excelPid = Get-ExcelProcessId -ExcelApplication $Excel
    } catch {}

    $isExcelVisible = $false
    try {
        $isExcelVisible = $Excel.Visible
    } catch {}

    $otherWorkbooksOpen = $false
    try {
        $wbs = $Excel.Workbooks
        foreach ($wb in $wbs) {
            if ($null -ne $WorkbookPath -and $wb.FullName -ne $WorkbookPath) {
                $otherWorkbooksOpen = $true
            }
            Release-ComObjectSafely $wb
        }
        Release-ComObjectSafely $wbs
    } catch {}

    if ($KeepAlive -and -not $ForceQuit) {
        try {
            $Excel.Visible = $true
            $Excel.DisplayAlerts = $true
        } catch {}
        Write-Host "  KeepAlive active: leaving Excel running with the workbook loaded." -ForegroundColor Yellow
    } else {
        if ($null -ne $WorkbookPath) {
            Write-Host "Closing workbook in Excel session..." -ForegroundColor Cyan
        } else {
            Write-Host "Cleaning up Excel session..." -ForegroundColor Cyan
        }
        try {
            $wbs = $Excel.Workbooks
            foreach ($wb in $wbs) {
                if ($null -ne $WorkbookPath -and $wb.FullName -eq $WorkbookPath) {
                    $wb.Close($SaveChanges)
                }
                Release-ComObjectSafely $wb
            }
            Release-ComObjectSafely $wbs
        } catch {}

        # Quit Excel if we started it, if it has no visible window, if no other workbooks are open, or if forced
        $shouldQuit = $ForceQuit -or -not $WasAlreadyOpen -or -not $isExcelVisible -or -not $otherWorkbooksOpen
        if ($shouldQuit) {
            try {
                $Excel.Quit()
            } catch {}
        }
    }

    Release-ComObjectSafely $Excel

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Hard-kill the background process if it is not supposed to remain open and it didn't shut down cleanly
    $shouldKill = $ForceQuit -or (-not $KeepAlive -and -not $WasAlreadyOpen)
    if ($shouldKill -and $excelPid -gt 0) {
        Start-Sleep -Milliseconds 500
        $proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($null -ne $proc -and $proc.Name -eq "EXCEL") {
            try {
                Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    }
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

function Clear-ExcelResiliencyItems {
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

    $recoveryPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DocumentRecovery"
    if (Test-Path $recoveryPath) {
        try {
            Remove-Item -Path $recoveryPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Cleared Excel DocumentRecovery resiliency items." -ForegroundColor Yellow
        } catch {}
    }

    $workspacePath = "HKCU:\Software\Microsoft\Office\16.0\Common\Restore Workspace"
    if (-not (Test-Path $workspacePath)) {
        try { New-Item -Path $workspacePath -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
    if (Test-Path $workspacePath) {
        try {
            Set-ItemProperty -Path $workspacePath -Name "RestoreWorkspace" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
        } catch {}
    }
}

function Start-ExcelApplication {
    param(
        [string]$Purpose
    )

    Clear-ExcelResiliencyItems

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

function Get-ActiveExcelWorkbook {
    param(
        [string]$WorkbookPath
    )
    
    Clear-ExcelResiliencyItems
    
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

    $resolvedExcelPath = if ($null -ne $excelPath) { $excelPath } else { Join-Path (Split-Path $PSScriptRoot -Parent) "Beaver Add-in.xlsm" }
    $resolvedProjectRoot = if ($null -ne $projectRoot) { $projectRoot } else { Split-Path (Split-Path $PSScriptRoot -Parent) -Parent }
    
    $wasAlreadyOpen = $false
    $workbookWasAlreadyOpen = $false
    $excel = $null
    $workbook = $null

    # 1. Reuse global Excel instance if orchestrator is active
    if ($global:BeaverOrchestratorActive -and $null -ne $global:BeaverSharedExcel) {
        try {
            $excel = $global:BeaverSharedExcel
            $wbs = $excel.Workbooks
            foreach ($wb in $wbs) {
                if ($wb.FullName -eq $resolvedExcelPath) {
                    $workbook = $wb
                } else {
                    Release-ComObjectSafely $wb
                }
            }
            Release-ComObjectSafely $wbs
            $wasAlreadyOpen = $global:BeaverExcelWasAlreadyOpen
            Write-Host "  Reusing persistent Excel COM session." -ForegroundColor Green
        } catch {
            Write-Warning "  Persistent Excel COM session is unresponsive. Starting a new session..."
            $global:BeaverSharedExcel = $null
            $excel = $null
        }
    }

    # 1.5 Try to attach via cached PID from build state (bypassing ROT registration limitations)
    if ($null -eq $excel) {
        $buildState = Get-BuildState
        if ($null -ne $buildState -and $buildState.PSObject.Properties.Name -contains "ExcelPid") {
            $cachedPid = $buildState.ExcelPid
            if ($cachedPid -gt 0) {
                $proc = Get-Process -Id $cachedPid -ErrorAction SilentlyContinue
                if ($null -ne $proc -and $proc.Name -eq "EXCEL") {
                    try {
                        # A. Try Accessibility window attachment
                        $excel = [WindowScraper]::GetExcelObject($cachedPid)
                        
                        # B. If fails, nudge the window to force window activation/registration and retry accessibility
                        if ($null -eq $excel) {
                            [WindowScraper]::NudgeProcess($cachedPid)
                            Start-Sleep -Milliseconds 200
                            $excel = [WindowScraper]::GetExcelObject($cachedPid)
                        }
                        
                        # C. Fallback: Try GetActiveObject and verify if the PID matches
                        if ($null -eq $excel) {
                            try {
                                $activeExcel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                                if ($null -ne $activeExcel) {
                                    $activePid = Get-ExcelProcessId -ExcelApplication $activeExcel
                                    if ($activePid -eq $cachedPid) {
                                        $excel = $activeExcel
                                        Write-Host "  Attached to running Excel session (PID: $cachedPid) via GetActiveObject ROT fallback." -ForegroundColor Green
                                    } else {
                                        Release-ComObjectSafely $activeExcel
                                    }
                                }
                            } catch {}
                        }
                        
                        if ($null -ne $excel) {
                            Write-Host "  Attached to running Excel session (PID: $cachedPid) successfully." -ForegroundColor Green
                            if ($null -eq $workbook) {
                                $wbs = $excel.Workbooks
                                foreach ($wb in $wbs) {
                                    if ($wb.FullName -eq $resolvedExcelPath) {
                                        $workbook = $wb
                                    } else {
                                        Release-ComObjectSafely $wb
                                    }
                                }
                                Release-ComObjectSafely $wbs
                            }
                            $wasAlreadyOpen = $true
                            if ($global:BeaverOrchestratorActive) {
                                $global:BeaverSharedExcel = $excel
                                $global:BeaverExcelWasAlreadyOpen = $wasAlreadyOpen
                            }
                        }
                    } catch {
                        $excel = $null
                    }
                }
            }
        }
    }

    # 2. Start a new session if needed
    if ($null -eq $excel) {
        # 1.8 Pre-startup lock cleanup: check if lock file exists and background processes exist
        $fileName = Split-Path $resolvedExcelPath -Leaf
        $lockFile = Join-Path $resolvedProjectRoot ("~$" + $fileName)
        $excelProcesses = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
        
        if ((Test-Path $lockFile) -and ($excelProcesses.Count -gt 0)) {
            $backgroundExcel = @(
                $excelProcesses | Where-Object {
                    $_.MainWindowHandle -eq 0 -and [string]::IsNullOrWhiteSpace($_.MainWindowTitle)
                }
            )
            
            if ($backgroundExcel.Count -gt 0) {
                Write-Host "  [PRE-STARTUP CLEANUP] Workbook is locked and background Excel process(es) found. Terminating to prevent read-only issues..." -ForegroundColor Yellow
                foreach ($proc in $backgroundExcel) {
                    try {
                        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                    } catch {}
                }
                Start-Sleep -Seconds 1
                if (Test-Path $lockFile) {
                    Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
                }
            }
        }

        $retryCount = 0
        $maxRetries = 10
        $ready = $false

        while (-not $ready -and $retryCount -lt $maxRetries) {
            try {
                $activeWbInfo = Get-ActiveExcelWorkbook -WorkbookPath $resolvedExcelPath
                if ($null -ne $activeWbInfo) {
                    $excel = $activeWbInfo.Excel
                    $workbook = $activeWbInfo.Workbook
                    $wasAlreadyOpen = $true
                } else {
                    # If workbook is not open, check if an Excel application is already running in the OS
                    $excelProcesses = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
                    if ($excelProcesses.Count -gt 0) {
                        try {
                            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                            $wasAlreadyOpen = $true
                        } catch {
                            $excel = $null
                        }
                    }
                    
                    if ($null -eq $excel) {
                        $excel = Start-ExcelApplication -Purpose $Purpose
                    }
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
            if ($wb.FullName -eq $resolvedExcelPath) {
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
                    $workbook = $excel.Workbooks.Open($resolvedExcelPath)
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
                throw "Failed to open workbook: $resolvedExcelPath"
            }
        } else {
            $wasAlreadyOpen = $true
            $workbookWasAlreadyOpen = $true
        }
    } else {
        $workbookWasAlreadyOpen = $true
    }

    # 4. Check programmatic VBE access
    try {
        $null = $excel.VBE
    } catch {
        throw "Programmatic access to the Visual Basic Project is not trusted. Please enable it in Excel under File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> 'Trust access to the VBA project object model'."
    }

    # 5. Check read-only state and perform self-healing recovery if needed
    if ($workbook.ReadOnly) {
        if ($null -ne $global:BeaverBuildLog) {
            $global:BeaverBuildLog.system.excelProcess.lockRecovered = $true
        }
        Write-Host "  [SELF-HEALING] The workbook '$resolvedExcelPath' was opened as Read-Only." -ForegroundColor Yellow
        Write-Host "  Attempting lock recovery..." -ForegroundColor Yellow
        try {
            $workbook.Close($false)
        } catch {}
        $workbook = $null

        # Terminate background Excel processes that may be holding the lock
        $excelProcesses = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
        $backgroundExcel = @(
            $excelProcesses | Where-Object {
                $_.MainWindowHandle -eq 0 -and [string]::IsNullOrWhiteSpace($_.MainWindowTitle)
            }
        )

        if ($backgroundExcel.Count -gt 0) {
            Write-Host "  Found $($backgroundExcel.Count) background Excel process(es) holding the lock. Terminating..." -ForegroundColor Yellow
            foreach ($proc in $backgroundExcel) {
                try {
                    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                } catch {}
            }
            Start-Sleep -Seconds 2
            
            # Remove the lock file if it was orphaned
            $fileName = Split-Path $resolvedExcelPath -Leaf
            $lockFile = Join-Path $resolvedProjectRoot ("~$" + $fileName)
            if (Test-Path $lockFile) {
                Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
            }
        }

        # Clear Excel disabled items registry key
        Clear-ExcelResiliencyItems

        # Ensure we have a responsive Excel application object (it might have been terminated)
        $excelResponsive = $false
        try {
            if ($null -ne $excel -and $excel.Hwnd) {
                $excelResponsive = $true
            }
        } catch {}

        if (-not $excelResponsive) {
            Write-Host "  Active Excel instance was terminated. Starting a fresh Excel instance..." -ForegroundColor Cyan
            $excel = Start-ExcelApplication -Purpose "workbook update after lock recovery"
            if ($global:BeaverOrchestratorActive) {
                $global:BeaverSharedExcel = $excel
            }
        }

        # Re-attempt opening the workbook
        Write-Host "  Re-attempting to open the workbook..." -ForegroundColor Cyan
        try {
            $excel.EnableEvents = $false
            $workbook = $excel.Workbooks.Open($resolvedExcelPath)
            $excel.EnableEvents = $true
        } catch {
            throw "Failed to open workbook during self-healing: $($_.Exception.Message)"
        }

        if ($workbook.ReadOnly) {
            throw "The workbook '$resolvedExcelPath' was opened as Read-Only even after lock recovery. Please ensure that no other Excel process is locking the file."
        } else {
            Write-Host "  Self-healing recovery successful! Workbook opened in write mode." -ForegroundColor Green
        }
    }

    # 6. Apply visibility
    if ($Visible) {
        $excel.Visible = $true
    }

    $excelPid = Get-ExcelProcessId -ExcelApplication $excel
    if ($null -ne $global:BeaverBuildLog) {
        $global:BeaverBuildLog.system.excelProcess.pid = $excelPid
        $global:BeaverBuildLog.system.excelProcess.wasAlreadyOpen = $wasAlreadyOpen
        $global:BeaverBuildLog.system.excelProcess.reusedSession = ($global:BeaverOrchestratorActive -and $null -ne $global:BeaverSharedExcel)
    }

    return [pscustomobject]@{
        Excel = $excel
        Workbook = $workbook
        WasAlreadyOpen = $wasAlreadyOpen
        WorkbookWasAlreadyOpen = $workbookWasAlreadyOpen
    }
}
