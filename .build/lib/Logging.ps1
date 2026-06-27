# Script:   Logging.ps1
# Purpose:  Console logging, stage results tracking, history logging, and termination for Beaver build pipeline.
# ==============================================================================

Set-StrictMode -Version Latest

# Metaprogramming override to suppress Write-Host output in JSON format mode
function Write-Host {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, ValueFromPipeline=$true)]
        [object]$Object,
        [string]$ForegroundColor,
        [string]$BackgroundColor,
        [switch]$NoNewline
    )
    if ($global:BeaverFormatJson) {
        return
    }
    $params = @{}
    if ($null -ne $Object) { $params["Object"] = $Object }
    if ($ForegroundColor) { $params["ForegroundColor"] = $ForegroundColor }
    if ($BackgroundColor) { $params["BackgroundColor"] = $BackgroundColor }
    if ($NoNewline) { $params["NoNewline"] = $true }
    Microsoft.PowerShell.Utility\Write-Host @params
}

# --- Stage Execution Tracking ---
if (-not (Get-Variable -Name "StageResults" -Scope Script -ErrorAction SilentlyContinue)) {
    $script:StageResults = New-Object System.Collections.ArrayList
}

# --- Build & Test Execution Logger ---
if (-not (Get-Variable -Name "BeaverBuildLog" -Scope Global -ErrorAction SilentlyContinue) -or $null -eq $global:BeaverBuildLog) {
    $cmdLine = if ($null -ne $MyInvocation -and $null -ne $MyInvocation.Line) { $MyInvocation.Line } else { "" }
    $global:BeaverBuildLog = [ordered]@{
        timestamp = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssK")
        status = "success"
        commandLine = $cmdLine
        buildMode = "skipped"
        system = [ordered]@{
            os = [System.Environment]::OSVersion.ToString()
            powershellVersion = $PSVersionTable.PSVersion.ToString()
            excelProcess = [ordered]@{
                pid = 0
                wasAlreadyOpen = $false
                reusedSession = $false
                lockRecovered = $false
            }
        }
        changes = [ordered]@{
            hasChanges = $false
            changedFiles = New-Object System.Collections.ArrayList
            deletedFiles = New-Object System.Collections.ArrayList
            manifestChanged = $false
            manifestStructureChanged = $false
        }
        stages = New-Object System.Collections.ArrayList
        lintResults = [ordered]@{
            checkedFiles = New-Object System.Collections.ArrayList
            errors = New-Object System.Collections.ArrayList
        }
        compileResults = [ordered]@{
            status = "success"
            errorDetails = $null
        }
        testResults = [ordered]@{
            runTests = $false
            filter = $null
            ribbonValidation = [ordered]@{
                status = "success"
                error = $null
            }
            unitTests = [ordered]@{
                total = 0
                passed = 0
                failed = 0
                durationMs = 0
                failures = New-Object System.Collections.ArrayList
            }
            headlessCallbacks = [ordered]@{
                status = "success"
                passedCount = 0
                failures = New-Object System.Collections.ArrayList
            }
        }
        totalDurationMs = 0
    }
}

function Clear-AccumulatedLogs {
    param(
        [int]$ExcludePid = 0
    )
    try {
        $tempPath = $env:TEMP
        if (Test-Path $tempPath) {
            # Find and delete BeaverAddin_*.log files
            $logFiles = Get-ChildItem -Path $tempPath -Filter "BeaverAddin_*.log" -ErrorAction SilentlyContinue
            foreach ($file in $logFiles) {
                if ($ExcludePid -gt 0 -and $file.Name -eq "BeaverAddin_$ExcludePid.log") {
                    continue
                }
                try {
                    Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
                } catch {}
            }
            # Delete test results file
            $testResultsFile = Join-Path $tempPath "BeaverAddin.TestResults.tsv"
            if (Test-Path $testResultsFile) {
                try {
                    Remove-Item -Path $testResultsFile -Force -ErrorAction SilentlyContinue
                } catch {}
            }
        }
    } catch {}
}

function Add-LintError {
    param(
        [string]$File,
        [string]$Type,
        [string]$Message,
        [int]$Line = 0
    )
    if ($null -ne $global:BeaverBuildLog) {
        [void]$global:BeaverBuildLog.lintResults.errors.Add([ordered]@{
            file = $File
            type = $Type
            message = $Message
            line = $Line
        })
    }
}

function Record-BuildChanges {
    param(
        [bool]$ManifestChanged,
        [bool]$ManifestStructureChanged,
        [string[]]$ChangedFiles,
        [string[]]$DeletedFiles,
        [bool]$Force
    )
    if ($null -eq $global:BeaverBuildLog) { return }
    $global:BeaverBuildLog.changes.hasChanges = ($ChangedFiles.Count -gt 0 -or $DeletedFiles.Count -gt 0)
    $global:BeaverBuildLog.changes.manifestChanged = $ManifestChanged
    $global:BeaverBuildLog.changes.manifestStructureChanged = $ManifestStructureChanged
    $global:BeaverBuildLog.changes.changedFiles.Clear()
    foreach ($file in $ChangedFiles) {
        [void]$global:BeaverBuildLog.changes.changedFiles.Add($file)
    }
    $global:BeaverBuildLog.changes.deletedFiles.Clear()
    foreach ($file in $DeletedFiles) {
        [void]$global:BeaverBuildLog.changes.deletedFiles.Add($file)
    }
    $global:BeaverBuildLog.buildMode = if ($Force) { "full" } else { "incremental" }
}

function Save-BuildLog {
    param(
        [string]$Status = "success",
        [switch]$Force
    )
    if ($null -eq $global:BeaverBuildLog) { return }
    if ($global:BeaverOrchestratorActive -and -not $Force -and $Status -ne "failure") {
        return
    }
    $global:BeaverBuildLog.status = $Status
    $start = [DateTime]::Parse($global:BeaverBuildLog.timestamp)
    $duration = (Get-Date) - $start
    $global:BeaverBuildLog.totalDurationMs = [Math]::Round($duration.TotalMilliseconds, 2)

    # Save build_log.json
    $logPath = Join-Path (Split-Path $PSScriptRoot -Parent) "build_log.json"
    $json = $global:BeaverBuildLog | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($logPath, $json, [System.Text.Encoding]::UTF8)

    # Save build_history.jsonl
    $historyLogPath = Join-Path (Split-Path $PSScriptRoot -Parent) "build_history.jsonl"
    $historyEntry = [ordered]@{
        timestamp = $global:BeaverBuildLog.timestamp
        status = $global:BeaverBuildLog.status
        commandLine = $global:BeaverBuildLog.commandLine
        buildMode = $global:BeaverBuildLog.buildMode
        changedFilesCount = $global:BeaverBuildLog.changes.changedFiles.Count
        deletedFilesCount = $global:BeaverBuildLog.changes.deletedFiles.Count
        totalDurationMs = $global:BeaverBuildLog.totalDurationMs
        unitTestsTotal = $global:BeaverBuildLog.testResults.unitTests.total
        unitTestsFailed = $global:BeaverBuildLog.testResults.unitTests.failed
        lintErrorsCount = $global:BeaverBuildLog.lintResults.errors.Count
    }
    $historyJson = $historyEntry | ConvertTo-Json -Compress
    [System.IO.File]::AppendAllText($historyLogPath, $historyJson + [System.Environment]::NewLine, [System.Text.Encoding]::UTF8)

    # Output systematic command line summary
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  BEAVER BUILD LOG SUMMARY ($($Status.ToUpper()))" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Timestamp:    $($global:BeaverBuildLog.timestamp)" -ForegroundColor Gray
    Write-Host "  Duration:     $([Math]::Round($global:BeaverBuildLog.totalDurationMs / 1000, 2))s" -ForegroundColor Gray
    Write-Host "  Build Mode:   $($global:BeaverBuildLog.buildMode)" -ForegroundColor Gray
    Write-Host "  Changes:      Changed=$($global:BeaverBuildLog.changes.changedFiles.Count), Deleted=$($global:BeaverBuildLog.changes.deletedFiles.Count)" -ForegroundColor Gray
    if ($global:BeaverBuildLog.system.excelProcess.pid -gt 0) {
        Write-Host "  Excel PID:    $($global:BeaverBuildLog.system.excelProcess.pid) (Reused: $($global:BeaverBuildLog.system.excelProcess.reusedSession), Lock Recovery: $($global:BeaverBuildLog.system.excelProcess.lockRecovered))" -ForegroundColor Gray
    }
    if ($global:BeaverBuildLog.lintResults.errors.Count -gt 0) {
        Write-Host "  Lint Errors:  $($global:BeaverBuildLog.lintResults.errors.Count) error(s) found" -ForegroundColor Red
    } else {
        Write-Host "  Lint Errors:  None" -ForegroundColor Green
    }
    if ($global:BeaverBuildLog.testResults.runTests) {
        $ut = $global:BeaverBuildLog.testResults.unitTests
        $tColor = if ($ut.failed -gt 0) { "Red" } else { "Green" }
        Write-Host "  Unit Tests:   Total=$($ut.total), Passed=$($ut.passed), Failed=$($ut.failed) (Duration: $($ut.durationMs)ms)" -ForegroundColor $tColor
    }
    Write-Host "  Log File:     $logPath" -ForegroundColor DarkGray
    Write-Host "========================================" -ForegroundColor Cyan
    
    if ($global:BeaverFormatJson) {
        Microsoft.PowerShell.Utility\Write-Output $json
    }
}

function Stop-Script {
    param(
        [string]$Message,
        [int]$ExitCode = 1,
        [switch]$DeleteState
    )

    Write-StageSummary
    
    $buildStatePathVar = Get-Variable -Name "buildStatePath" -ErrorAction SilentlyContinue
    $resolvedBuildStatePath = if ($null -ne $buildStatePathVar) { $buildStatePathVar.Value } else { Join-Path (Split-Path $PSScriptRoot -Parent) ".build_state.json" }

    # Save failed files in build state before exiting
    if (-not $DeleteState) {
        try {
            $failedList = @()
            $var = Get-Variable -Name "changedFiles" -ErrorAction SilentlyContinue
            if ($null -ne $var -and $null -ne $var.Value) {
                $failedList = @($var.Value)
            }
            
            $buildState = Get-BuildState
            if ($null -ne $buildState) {
                $buildState | Add-Member -NotePropertyName TestsPassed -NotePropertyValue $false -Force
                
                $existingFailed = @()
                if ($buildState.PSObject.Properties.Name -contains "FailedFiles" -and $null -ne $buildState.FailedFiles) {
                    $existingFailed = @($buildState.FailedFiles)
                }
                
                $merged = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($file in $existingFailed) {
                    if ($null -ne $file) { [void]$merged.Add($file.ToString()) }
                }
                foreach ($file in $failedList) {
                    if ($null -ne $file) { [void]$merged.Add($file.ToString()) }
                }
                
                $mergedArray = [string[]]($merged | Where-Object { $null -ne $_ })
                $buildState | Add-Member -NotePropertyName FailedFiles -NotePropertyValue $mergedArray -Force
                
                $stateJson = $buildState | ConvertTo-Json -Depth 10
                [System.IO.File]::WriteAllText($resolvedBuildStatePath, $stateJson, [System.Text.Encoding]::UTF8)
            }
        } catch {
            Write-Warning "Stop-Script: Failed to save failed state: $($_.Exception.Message)"
        }
    }
    
    # Delete build state only on catastrophic or explicit request
    if ($DeleteState -and (Test-Path $resolvedBuildStatePath)) {
        Remove-Item $resolvedBuildStatePath -Force -ErrorAction SilentlyContinue
    }

    Save-BuildLog -Status "failure" -Force
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
        if ($null -ne $global:BeaverBuildLog) {
            [void]$global:BeaverBuildLog.stages.Add([ordered]@{
                name = $Stage
                status = "success"
                durationMs = [Math]::Round($stopwatch.Elapsed.TotalMilliseconds, 2)
                details = $details
            })
        }
        Write-StatusLine -Status "pass" -Stage $Stage -Details ("({0:N1}s){1}" -f $stopwatch.Elapsed.TotalSeconds, $(if ([string]::IsNullOrWhiteSpace($details)) { "" } else { " $details" })) -Color "Green"
        return $result
    } catch {
        $stopwatch.Stop()
        $message = $_.Exception.Message
        Add-StageResult -Stage $Stage -Status "failure" -Details $message -DurationMs $stopwatch.Elapsed.TotalMilliseconds
        if ($null -ne $global:BeaverBuildLog) {
            [void]$global:BeaverBuildLog.stages.Add([ordered]@{
                name = $Stage
                status = "failure"
                durationMs = [Math]::Round($stopwatch.Elapsed.TotalMilliseconds, 2)
                details = $message
            })
        }
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
    if ($null -ne $global:BeaverBuildLog) {
        [void]$global:BeaverBuildLog.stages.Add([ordered]@{
            name = $Stage
            status = "skipped"
            durationMs = 0
            details = $Details
        })
    }
    Write-StatusLine -Status "skip" -Stage $Stage -Details $Details -Color "Yellow"
}
