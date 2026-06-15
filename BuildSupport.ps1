# Script:   BuildSupport.ps1
# Purpose:  Shared helpers, paths, and Excel COM management for Build.ps1 and Test.ps1.
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Common Paths ---
$excelPath = Join-Path $PSScriptRoot "Beaver Add-in.xlsm"
$modulesDir = Join-Path $PSScriptRoot "Modules"
$desktopThisWorkbookCls = Join-Path $PSScriptRoot "ThisWorkbook.cls"
$ribbonXmlPath = Join-Path $PSScriptRoot "ribbon.xml"
$featureManifestPath = Join-Path $PSScriptRoot "features.json"
$configPath = Join-Path $PSScriptRoot "config.json"
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

function Start-ExcelApplication {
    param(
        [string]$Purpose
    )

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
