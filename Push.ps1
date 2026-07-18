# Script:   Push.ps1
# Purpose:  Fully automated, self-sustaining verification, versioning, staging, committing, and pushing script for AI agents.
# Usage:    pwsh -File .\Push.ps1
#           pwsh -File .\Push.ps1 -DryRun
# ==============================================================================

[CmdletBinding()]
param(
    [switch]$DryRun,
    [string]$TargetBranch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "    BEAVER ADD-IN: AGENT PUSH PIPELINE   " -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

# 1. Clear locks and clean up running Excel sessions
Write-Host "Checking for active Excel processes..." -ForegroundColor Yellow
$excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host "  Found active Excel processes. Cleaning up sessions to release workbook locks..." -ForegroundColor Yellow
    foreach ($p in $excelProcesses) {
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 1
}

# 2. Run local compilation, linting and verification
Write-Host "Running compilation, lint and test pipeline..." -ForegroundColor Cyan
$updateScript = Join-Path $PSScriptRoot "Update.ps1"
if ($DryRun) {
    Write-Host "  [DryRun] Compiling and running tests headless..." -ForegroundColor Gray
}

# Invoke Update.ps1
& pwsh -File $updateScript -Force
if ($LASTEXITCODE -ne 0) {
    Write-Error "Local verification pipeline failed! Fix errors before pushing."
    exit 1
}
Write-Host "Verification pipeline completed successfully." -ForegroundColor Green

# 3. Determine changed files for semantic commit message
Write-Host "Scanning git changes..." -ForegroundColor Cyan
$gitStatus = git status --porcelain
if ([string]::IsNullOrWhiteSpace($gitStatus)) {
    Write-Host "No changes detected in source control. Nothing to commit." -ForegroundColor Green
    exit 0
}

$changedFiles = @()
$deletedFiles = @()
foreach ($line in ($gitStatus -split "`r`n")) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $status = $line.Substring(0, 2).Trim()
    $file = $line.Substring(3).Trim()
    if ($status -eq "D") {
        $deletedFiles += $file
    } else {
        $changedFiles += $file
    }
}

# Generate semantic commit message
$commitSubject = ""
$commitBody = "Automated commit by Beaver Agent.`n`nChanges:"

$commandChanges = @()
$infraChanges = @()
$testChanges = @()
$manifestChanges = @()
$otherChanges = @()

foreach ($file in ($changedFiles + $deletedFiles)) {
    $fileName = Split-Path $file -Leaf
    if ($file -match "Modules/Commands/") {
        $commandName = $fileName -replace "^FeatCmd_|\.cls$", ""
        $commandChanges += $commandName
    } elseif ($file -match "Modules/Infrastructure/") {
        $infraChanges += $fileName
    } elseif ($file -match "Modules/Tests/") {
        $testChanges += $fileName
    } elseif ($fileName -eq "features.json" -or $fileName -eq "config.json" -or $fileName -eq "ribbon.xml") {
        $manifestChanges += $fileName
    } else {
        if ($file -notmatch "\.xlsm$") { # Skip binary workbook
            $otherChanges += $file
        }
    }
}

if ($commandChanges.Count -gt 0) {
    $commitSubject = "feat(" + ($commandChanges -join ", ") + "): update features"
} elseif ($manifestChanges.Count -gt 0) {
    $commitSubject = "feat(manifest): sync feature declarations"
} elseif ($infraChanges.Count -gt 0) {
    $commitSubject = "refactor(infra): update core infrastructure"
} elseif ($testChanges.Count -gt 0) {
    $commitSubject = "test: update test cases"
} else {
    $commitSubject = "chore: general updates"
}

# Build body details
if ($commandChanges) { $commitBody += "`n- Modified command features: " + ($commandChanges -join ", ") }
if ($infraChanges) { $commitBody += "`n- Modified core infrastructure: " + ($infraChanges -join ", ") }
if ($testChanges) { $commitBody += "`n- Modified unit tests: " + ($testChanges -join ", ") }
if ($manifestChanges) { $commitBody += "`n- Modified config/manifests: " + ($manifestChanges -join ", ") }
if ($otherChanges) { $commitBody += "`n- Modified other files: " + ($otherChanges -join ", ") }

Write-Host "Generated Semantic Commit Message:" -ForegroundColor Yellow
Write-Host "-----------------------------------------" -ForegroundColor DarkGray
Write-Host "Subject: $commitSubject" -ForegroundColor Green
Write-Host "Body:`n$commitBody" -ForegroundColor Green
Write-Host "-----------------------------------------" -ForegroundColor DarkGray

# 4. Auto version bump in config.json
Write-Host "Bumping add-in version in config.json..." -ForegroundColor Cyan
$configPath = Join-Path $PSScriptRoot "config.json"
$config = Get-Content $configPath -Raw | ConvertFrom-Json
$currentVersion = $config.AddinIdentity.Version
$versionParts = $currentVersion -split "\."
if ($versionParts.Count -eq 3) {
    $patch = [int]$versionParts[2] + 1
    $newVersion = "$($versionParts[0]).$($versionParts[1]).$patch"
} else {
    $newVersion = "$currentVersion.1"
}

Write-Host "  Version Bump: $currentVersion -> $newVersion" -ForegroundColor Green

if (-not $DryRun) {
    $config.AddinIdentity.Version = $newVersion
    # Save back to config.json
    $configJson = $config | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($configPath, $configJson, [System.Text.Encoding]::UTF8)
    
    # We must rebuild/recompile once more to embed the bumped version into the compiled workbook!
    Write-Host "Re-compiling workbook to embed the new version..." -ForegroundColor Cyan
    & pwsh -File $updateScript -Force -SkipRuntimeTests
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Embed compilation failed!"
        exit 1
    }
} else {
    Write-Host "  [DryRun] Skipped config.json version bump write-back." -ForegroundColor Gray
}

# 5. Git commit & push
if (-not $DryRun) {
    Write-Host "Staging files..." -ForegroundColor Cyan
    git add config.json features.json ribbon.xml Modules/ ThisWorkbook.cls ARCHITECTURE.md
    
    $fullMessage = "$commitSubject`n`n$commitBody"
    Write-Host "Committing changes..." -ForegroundColor Cyan
    git commit -m $fullMessage
    
    $branch = $TargetBranch
    if ([string]::IsNullOrWhiteSpace($branch)) {
        # Auto-detect current branch
        $branch = (git branch --show-current).Trim()
    }
    
    Write-Host "Pushing to remote upstream branch '$branch'..." -ForegroundColor Cyan
    git push origin $branch
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Push completed successfully!" -ForegroundColor Green
    } else {
        Write-Error "Failed to push changes to remote repository."
        exit 1
    }
} else {
    Write-Host "  [DryRun] Skipped Git commit and push stages." -ForegroundColor Gray
}

Write-Host "Done!" -ForegroundColor Green
