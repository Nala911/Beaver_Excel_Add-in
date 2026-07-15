# Script:   Watch.ps1
# Purpose:  Continuous file system watcher that triggers incremental builds
#           and smart testing instantly while keeping the Excel session alive.
# Usage:    pwsh -File .\Watch.ps1
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$modulesDir = Join-Path $projectRoot "Modules"
$updateScript = Join-Path $projectRoot "Update.ps1"

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "      BEAVER ADD-IN: WATCH MODE          " -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Starting watch mode. Press [Ctrl+C] to stop." -ForegroundColor Gray
Write-Host "Initial run to synchronize state and start Excel session..." -ForegroundColor Gray

# Run once at startup and ensure Excel is kept alive
& pwsh -File $updateScript -KeepAlive
Write-Host "Watching for changes in Modules/ and configuration files..." -ForegroundColor Green

# Set up file system watcher
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $projectRoot
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true

# Debounce state
$global:lastTriggered = [DateTime]::MinValue
$global:isBuilding = $false

$action = {
    param($sender, $eventArgs)
    
    $filePath = $eventArgs.FullPath
    $fileName = $eventArgs.Name
    
    # Filter for relevant files only (.bas, .cls, .frm, .frx, features.json, config.json, ribbon.xml)
    if ($filePath -notmatch "\.(bas|cls|frm|frx|json|xml)$") {
        return
    }
    
    # Ignore build state cache, temp files, log files, git files, etc.
    if ($filePath -match "\\\.build" -or $filePath -match "\\\.git" -or $fileName -match "^\~\$") {
        return
    }
    
    # Debounce checks (100ms)
    $now = Get-Date
    if (($now - $global:lastTriggered).TotalMilliseconds -lt 100 -or $global:isBuilding) {
        return
    }
    
    $global:isBuilding = $true
    $global:lastTriggered = $now
    
    Write-Host ""
    Write-Host "-----------------------------------------" -ForegroundColor DarkGray
    Write-Host "Change detected: $fileName" -ForegroundColor Yellow
    Write-Host "Triggering incremental update and testing..." -ForegroundColor Yellow
    
    try {
        & pwsh -File $updateScript -KeepAlive
    } catch {
        Write-Error $_.Exception.Message
    } finally {
        $global:isBuilding = $false
    }
}

# Clean any existing event subscriptions of the same names to be robust
Get-EventSubscriber -ErrorAction SilentlyContinue | Where-Object { 
    $_.SourceIdentifier -like "BeaverWatch_*" 
} | Unregister-Event -ErrorAction SilentlyContinue

$createdEvent = Register-ObjectEvent -InputObject $watcher -EventName "Created" -SourceIdentifier "BeaverWatch_Created" -Action $action
$changedEvent = Register-ObjectEvent -InputObject $watcher -EventName "Changed" -SourceIdentifier "BeaverWatch_Changed" -Action $action
$deletedEvent = Register-ObjectEvent -InputObject $watcher -EventName "Deleted" -SourceIdentifier "BeaverWatch_Deleted" -Action $action

try {
    while ($true) {
        Start-Sleep -Seconds 1
    }
} finally {
    Write-Host "Stopping watcher and cleaning up events..." -ForegroundColor Yellow
    Unregister-Event -SourceIdentifier "BeaverWatch_Created" -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier "BeaverWatch_Changed" -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier "BeaverWatch_Deleted" -ErrorAction SilentlyContinue
    $watcher.Dispose()
    Write-Host "Watch mode stopped." -ForegroundColor Green
}
