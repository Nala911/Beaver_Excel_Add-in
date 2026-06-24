[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$SkipLint,
    [switch]$Clean
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "BuildSupport.ps1")
. (Join-Path $PSScriptRoot "Linter.ps1")

# Initialize global caches to satisfy StrictMode
$global:BeaverLintStatusCache = @{}
$global:BeaverTestManifestCache = @{}
$global:BeaverBuildStateCache = $null
$global:BeaverFeatureManifestCache = $null
$global:BeaverFileContentCache = $null

# --- Helper Functions (Moved to lib\Generators.ps1 and lib\RibbonUtils.ps1) ---
# --- Build Execution ---

if ($Clean) {
    Write-Host "Running clean stage..." -ForegroundColor Cyan
    if (Test-Path $buildStatePath) {
        Remove-Item $buildStatePath -Force -ErrorAction SilentlyContinue
        Write-Host "  Cleared build state cache (.build_state.json)." -ForegroundColor Green
    }
    $stoppedAny = Remove-OrphanedExcelProcesses
    if (-not $stoppedAny) {
        Write-Host "  No orphaned Excel processes to clean." -ForegroundColor Gray
    }
    Write-Host "Clean complete." -ForegroundColor Green
    exit 0
}

$sharedExcel = $null
$excelWasAlreadyOpen = $false

try {
    # --- Change Detection ---
    $currentHashes = Get-SourceFileHashes
    $buildState = Get-BuildState
    $manifestChanged = $true
    $changedFiles = @()
    $deletedFiles = @()

    if ($null -ne $buildState -and $null -ne $buildState.Files) {
        $featProp = $buildState.Files.PSObject.Properties["features.json"]
        $manifestChanged = ($null -eq $featProp -or $featProp.Value -ne $currentHashes["features.json"])
        
        foreach ($key in $currentHashes.Keys) {
            $prop = $buildState.Files.PSObject.Properties[$key]
            if ($null -eq $prop -or $prop.Value -ne $currentHashes[$key]) {
                $changedFiles += $key
            }
        }
        
        foreach ($prop in $buildState.Files.PSObject.Properties) {
            $key = $prop.Name
            if (-not $currentHashes.ContainsKey($key)) {
                $deletedFiles += $key
            }
        }
        
        # Merge previously failed files to force their rebuild/re-import
        if ($buildState.PSObject.Properties.Name -contains "FailedFiles" -and $null -ne $buildState.FailedFiles) {
            foreach ($file in $buildState.FailedFiles) {
                if ($changedFiles -notcontains $file) {
                    $changedFiles += $file
                }
            }
        }
    } else {
        foreach ($key in $currentHashes.Keys) {
            $changedFiles += $key
        }
    }

    $hasAnyChanges = ($changedFiles.Count -gt 0 -or $deletedFiles.Count -gt 0)

    if (-not $hasAnyChanges -and (Test-Path $excelPath) -and -not $Force) {
        Write-Host "No changes detected. Build skipped." -ForegroundColor Green
        if ($null -ne $global:BeaverBuildLog) {
            $global:BeaverBuildLog.buildMode = "skipped"
        }
        Save-BuildLog -Status "success"
        exit 0
    }

    $manifestStructureChanged = $false
    if ($manifestChanged) {
        $newStructuralHash = Get-ManifestStructuralHash -Path $featureManifestPath
        $oldStructuralHash = $null
        if ($null -ne $buildState -and $buildState.PSObject.Properties.Name.Contains("ManifestStructuralHash")) {
            $oldStructuralHash = $buildState.ManifestStructuralHash
        }
        if ($newStructuralHash -ne $oldStructuralHash) {
            $manifestStructureChanged = $true
        } else {
            Write-Host "  Manifest changed but structure is identical (metadata-only update)." -ForegroundColor Yellow
        }
    }

    Record-BuildChanges -ManifestChanged $manifestChanged -ManifestStructureChanged $manifestStructureChanged -ChangedFiles $changedFiles -DeletedFiles $deletedFiles -Force $Force

    $forceFullBuild = (-not (Test-Path $excelPath) -or $Force)

    if ($forceFullBuild) {
        Write-Host "Performing clean full build..." -ForegroundColor Cyan
    } else {
        Write-Host "Performing incremental build..." -ForegroundColor Cyan
        Write-Host "  Changed files: $($changedFiles.Count), Deleted files: $($deletedFiles.Count)" -ForegroundColor Yellow
    }

    Invoke-Stage -Stage "manifest_sync" -Action {
        if ($forceFullBuild -or $manifestChanged) {
            Sync-FeatureManifest -ManifestPath $featureManifestPath -ConfigPath $configPath -RibbonPath $ribbonXmlPath
            return "features synced from features.json"
        } else {
            return "skipped (manifest unchanged)"
        }
    } | Out-Null

    Invoke-Stage -Stage "command_registry_generation" -Action {
        $helpManifestPath = Join-Path $modulesDir "Libraries\Lib_HelpManifest.bas"
        $udfRegistryPath = Join-Path $modulesDir "Libraries\Lib_UdfRegistry.bas"
        $registryMissing = -not (Test-Path $commandRegistryPath)
        $helpMissing = -not (Test-Path $helpManifestPath)
        $udfMissing = -not (Test-Path $udfRegistryPath)
        $registryGenerated = $false
        $helpGenerated = $false
        $udfGenerated = $false
        
        if ($forceFullBuild -or $manifestStructureChanged -or $manifestChanged -or $registryMissing -or $helpMissing -or $udfMissing) {
            if ($forceFullBuild -or $manifestStructureChanged -or $registryMissing) {
                Sync-CommandRegistry -ManifestPath $featureManifestPath -OutputPath $commandRegistryPath
                $registryGenerated = $true
            }
            if ($forceFullBuild -or $manifestChanged -or $helpMissing) {
                Sync-HelpManifest -ManifestPath $featureManifestPath -OutputPath $helpManifestPath
                $helpGenerated = $true
            }
            if ($forceFullBuild -or $manifestChanged -or $udfMissing) {
                Sync-UdfRegistry -ManifestPath $featureManifestPath -OutputPath $udfRegistryPath
                $udfGenerated = $true
            }
            
            if (-not $forceFullBuild) {
                $generatedRelPaths = @()
                if ($registryGenerated) { $generatedRelPaths += "Modules/Infrastructure/Infra_CommandRegistry.bas" }
                if ($helpGenerated) { $generatedRelPaths += "Modules/Libraries/Lib_HelpManifest.bas" }
                if ($udfGenerated) { $generatedRelPaths += "Modules/Libraries/Lib_UdfRegistry.bas" }

                foreach ($relPath in $generatedRelPaths) {
                    if ($changedFiles -notcontains $relPath) {
                        if (Test-BuildStateFileChanged -RelativePath $relPath -BuildState $buildState) {
                            $script:changedFiles += $relPath
                        }
                    }
                }
            }
            
            $refreshed = @()
            if ($registryGenerated) { $refreshed += "command registry" }
            if ($helpGenerated) { $refreshed += "help manifest" }
            if ($udfGenerated) { $refreshed += "UDF registry" }
            
            if ($refreshed.Count -gt 0) {
                return ($refreshed -join ", ") + " refreshed"
            }
            return "skipped (generated files already current)"
        } else {
            return "skipped (manifest unchanged)"
        }
    } | Out-Null

    Invoke-Stage -Stage "ui_entry_generation" -Action {
        if ($forceFullBuild -or $manifestStructureChanged) {
            Sync-UiRibbonModule -ManifestPath $featureManifestPath -OutputPath $uiRibbonPath
            Sync-UiHotkeysModule -ManifestPath $featureManifestPath -OutputPath $uiHotkeysPath
            if (-not $forceFullBuild) {
                foreach ($relPath in @("Modules/UI/UI_Ribbon.bas", "Modules/UI/UI_Hotkeys.bas")) {
                    if ($changedFiles -notcontains $relPath) {
                        if (Test-BuildStateFileChanged -RelativePath $relPath -BuildState $buildState) {
                            $script:changedFiles += $relPath
                        }
                    }
                }
            }
            return "UI entry modules refreshed"
        } else {
            return "skipped (manifest unchanged)"
        }
    } | Out-Null

    Invoke-Stage -Stage "test_manifest_generation" -Action {
        $hasBasChanges = @($changedFiles | Where-Object { $_ -match "\.bas$" }).Count -gt 0
        if ($forceFullBuild -or $hasBasChanges) {
            Sync-TestManifest -SourceDir $modulesDir -OutputPath $testManifestPath
            
            # Explicitly append regenerated manifest to changed files for import
            if (-not $forceFullBuild) {
                $relPath = "Modules/Tests/Lib_TestManifest.bas"
                if ($changedFiles -notcontains $relPath) {
                    if (Test-BuildStateFileChanged -RelativePath $relPath -BuildState $buildState) {
                        $script:changedFiles += $relPath
                    }
                }
            }
            return "test manifest refreshed"
        } else {
            return "skipped (no test file changes)"
        }
    } | Out-Null

    Invoke-Stage -Stage "validation" -Action {
        $filesToValidate = if ($forceFullBuild) { $null } else { $changedFiles }
        $validRibbon = if ($forceFullBuild -or $manifestChanged) { Test-RibbonValidity -XmlPath $ribbonXmlPath -ModulesDir $modulesDir } else { $true }
        
        $validVba = $true
        $validLint = $true
        $validForms = $true
        
        if (-not $SkipLint) {
            $validVba = Invoke-VbaSyntaxCheck -SourceDir $modulesDir -FilesToProcess $filesToValidate
            $validLint = Invoke-EnhancedLinting -SourceDir $modulesDir -FilesToProcess $filesToValidate
            $validForms = Test-FormFilesValidity -SourceDir $modulesDir -FilesToProcess $filesToValidate
            $null = Invoke-TestCoverageAudit -SourceDir $modulesDir
        } else {
            Write-Host "  Skipping linting, syntax, and form validity checks (-SkipLint)." -ForegroundColor Yellow
        }

        if (-not ($validRibbon -and $validVba -and $validLint -and $validForms)) {
            throw "Pre-deployment validation failed"
        }

        if ($SkipLint) {
            return "ribbon validated (syntax and lint skipped)"
        } else {
            return "ribbon, syntax, lint, and form checks passed"
        }
    } | Out-Null

    # --- environment_checks ---
    Invoke-Stage -Stage "environment_checks" -Action {
        if (-not (Test-Path $excelPath)) {
            throw "Excel file not found: $excelPath"
        }

        if ($forceFullBuild -or $manifestChanged) {
            # 1. Check if the file is actually locked
            if (-not (Test-FileLocked -Path $excelPath)) {
                # If not locked but a lock file exists, it's orphaned. We can safely remove it.
                $lockFile = Join-Path $projectRoot ("~$" + (Split-Path $excelPath -Leaf))
                if (Test-Path $lockFile) {
                    Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
                }
                return "workbook available"
            }

            Write-Host "Excel workbook is locked. Attempting to close it..." -ForegroundColor Yellow
            $closedGracefully = $false
            
            try {
                $activeWbInfo = Get-ActiveExcelWorkbook -WorkbookPath $excelPath
                if ($null -ne $activeWbInfo) {
                    $activeExcel = $activeWbInfo.Excel
                    $wbFound = $activeWbInfo.Workbook
                    
                    $activeExcel.DisplayAlerts = $false
                    
                    $otherVisibleWorkbooks = 0
                    foreach ($otherWb in $activeExcel.Workbooks) {
                        if ($otherWb.FullName -ne $excelPath) {
                            $hasVisibleWindow = $false
                            try {
                                foreach ($win in $otherWb.Windows) {
                                    if ($win.Visible) {
                                        $hasVisibleWindow = $true
                                        break
                                    }
                                }
                            } catch {
                                $hasVisibleWindow = $true
                            }
                            if ($hasVisibleWindow) {
                                $otherVisibleWorkbooks++
                            }
                        }
                    }

                    $wbFound.Close($true)
                    Write-Host "  Closed $($wbFound.Name) successfully." -ForegroundColor Green
                    $closedGracefully = $true

                    if ($otherVisibleWorkbooks -eq 0) {
                        Write-Host "  No other visible workbooks open. Closing Excel application..." -ForegroundColor Green
                        $activeExcel.Quit()
                    }
                    
                    Release-ComObjectSafely $wbFound
                    Release-ComObjectSafely $activeExcel
                }
            } catch {
                Write-Warning "Graceful close via COM failed: $($_.Exception.Message)"
            }

            # 3. If still locked, handle termination safety
            if (Test-FileLocked -Path $excelPath) {
                # Check for visible Excel windows
                $excelProcesses = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
                $visibleExcel = @(
                    $excelProcesses | Where-Object {
                        $_.MainWindowHandle -ne 0 -or -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle)
                    }
                )

                if ($visibleExcel.Count -gt 0) {
                    throw "The workbook '$excelPath' is open and locked in a visible Excel window. Please save your work, close Excel, and rerun the script."
                }

                if ($excelProcesses.Count -gt 0) {
                    Write-Host "  Found background Excel process(es) holding the lock. Cleaning up..." -ForegroundColor Yellow
                    foreach ($process in $excelProcesses) {
                        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                    }
                    Start-Sleep -Seconds 2
                }
            }

            # Final verify
            if (Test-FileLocked -Path $excelPath) {
                throw "Excel file is still open and locked. Please close it manually and retry."
            }

            return "lock cleared"
        } else {
            return "skipped (in-place reload active)"
        }
    } | Out-Null

    # --- workbook_update ---
    Invoke-Stage -Stage "workbook_update" -Action {
        $compsToRemove = @()
        $filesToImport = @()
        $thisWorkbookChanged = ($changedFiles -contains "ThisWorkbook.cls")

        if (-not $forceFullBuild) {
            # Incremental Mode: Calculate components to remove/import
            foreach ($relPath in ($changedFiles + $deletedFiles)) {
                if ($relPath -eq "features.json" -or $relPath -eq "ThisWorkbook.cls") { continue }
                if ($relPath -notmatch "\.(bas|cls|frm)$") { continue }
                
                $filePath = Join-Path $projectRoot $relPath
                $compName = $null
                
                if ($changedFiles -contains $relPath) {
                    $compName = Get-VbaComponentNameFromFile -FilePath $filePath
                } else {
                    $fileName = Split-Path $relPath -Leaf
                    $compName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                }
                
                if ($null -ne $compName -and $compName -ne "") {
                    $compsToRemove += $compName
                }
                
                if ($changedFiles -contains $relPath) {
                    $filesToImport += $filePath
                }
            }
        }

        $hasVbaChanges = ($forceFullBuild -or $compsToRemove.Count -gt 0 -or $filesToImport.Count -gt 0 -or $thisWorkbookChanged)

        if (-not $hasVbaChanges) {
            return "skipped (no VBA code changes)"
        }

        $session = Initialize-ExcelWorkbookSession -Purpose "workbook update"
        $excel = $session.Excel
        $workbook = $session.Workbook
        $wasAlreadyOpen = $session.WasAlreadyOpen
        $script:excelWasAlreadyOpen = $wasAlreadyOpen
        $script:workbookWasAlreadyOpen = $session.WorkbookWasAlreadyOpen
        $script:sharedExcel = $excel
        
        $vbaProject = $workbook.VBProject
        $btn = $null
        $activePane = $null
        $commandBars = $null
        $cb = $null

        $origScreenUpdating = $true
        $origEvents = $true
        $origCalculation = -4105

        try {
            try {
                if ($null -ne $excel) {
                    $origScreenUpdating = $excel.ScreenUpdating
                    $origEvents = $excel.EnableEvents
                    $origCalculation = $excel.Calculation

                    $excel.ScreenUpdating = $false
                    $excel.EnableEvents = $false
                    $excel.Calculation = -4135 # xlCalculationManual
                }
            } catch {}

            Write-Host "Updating modules..."
            
            if ($forceFullBuild) {
                # Purge all components
                $compsToRemoveList = @()
                $componentsCollection = $vbaProject.VBComponents
                foreach ($comp in $componentsCollection) {
                    if (($comp.Type -ge 1 -and $comp.Type -le 3) -and ($comp.Name -ne "ThisWorkbook")) {
                        $compsToRemoveList += $comp
                    } else {
                        Release-ComObjectSafely $comp
                    }
                }
                Release-ComObjectSafely $componentsCollection
                $vbaSourceFiles = Get-ChildItem -Path $modulesDir -Include *.bas, *.cls, *.frm -Recurse
                $filesToImport = @($vbaSourceFiles | ForEach-Object { $_.FullName })
                $compsToRemove = $compsToRemoveList
            } else {
                # Incremental Mode: convert component names to VBComponent COM objects
                $compsToRemoveList = @()
                $componentsCollection = $vbaProject.VBComponents
                foreach ($compName in $compsToRemove) {
                    try {
                        $comp = $componentsCollection.Item($compName)
                        if ($null -ne $comp) {
                            $compsToRemoveList += $comp
                        }
                    } catch {
                        # Component doesn't exist
                    }
                }
                Release-ComObjectSafely $componentsCollection
                $compsToRemove = $compsToRemoveList
            }

            foreach ($comp in $compsToRemove) {
                $compName = "unknown"
                try {
                    $compName = $comp.Name
                    $vbaProject.VBComponents.Remove($comp)
                    Write-Host "  Removed component: $compName"
                } catch {
                    throw "Failed to remove existing VBA component '$compName': $($_.Exception.Message). Please ensure Excel is not in break mode or busy."
                } finally {
                    Release-ComObjectSafely $comp
                }
            }

            $tempImportDir = Join-Path ([System.IO.Path]::GetTempPath()) ("BeaverAddin-VbaImport-" + [System.Guid]::NewGuid().ToString("N"))

            try {
                foreach ($filePath in $filesToImport) {
                    $fileName = Split-Path $filePath -Leaf
                    Write-Host "  Importing $($fileName)..."
                    $importPath = New-NormalizedImportCopy -SourcePath $filePath -TempRoot $tempImportDir

                    # Only copy companion .frx if we actually normalized it (meaning $importPath is inside $tempImportDir)
                    if ($importPath -ne $filePath -and ($filePath -match "\.frm$")) {
                        $frxPath = [System.IO.Path]::ChangeExtension($filePath, ".frx")
                        if (Test-Path $frxPath) {
                            $tempFrxPath = Join-Path $tempImportDir ([System.IO.Path]::GetFileName($frxPath))
                            Copy-Item -Path $frxPath -Destination $tempFrxPath -Force
                        }
                    }

                    $vbaProject.VBComponents.Import($importPath) | Out-Null
                }
            } finally {
                if (Test-Path $tempImportDir) {
                    Remove-Item -Path $tempImportDir -Recurse -Force -ErrorAction SilentlyContinue
                }
            }

            $thisWorkbookChanged = ($changedFiles -contains "ThisWorkbook.cls")
            if ((Test-Path $desktopThisWorkbookCls) -and ($forceFullBuild -or $thisWorkbookChanged)) {
                Write-Host "  Updating ThisWorkbook..."
                $twCode = $vbaProject.VBComponents.Item("ThisWorkbook").CodeModule
                if ($twCode.CountOfLines -gt 0) { $twCode.DeleteLines(1, $twCode.CountOfLines) }
                $lines = Get-Content $desktopThisWorkbookCls | Where-Object {
                    $_ -notmatch "^VERSION\s+\d+\.\d+" -and
                    $_ -notmatch "^BEGIN\s*$" -and
                    $_ -notmatch "^\s+MultiUse\s*=" -and
                    $_ -notmatch "^END\s*$" -and
                    $_ -notmatch "^Attribute\s+"
                }
                $twCode.AddFromString([string]::Join("`r`n", $lines))
            }

            $excelPid = Get-ExcelProcessId -ExcelApplication $excel
            $signalFile = Join-Path $env:TEMP ("BeaverBuildSignal_" + [System.Guid]::NewGuid().ToString("N") + ".tmp")
            $watcher = $null

            if ($excelPid -gt 0) {
                $watcher = Start-ExcelWindowWatcher -ExcelPid $excelPid -SignalPath $signalFile -TimeoutSeconds 20
            }

            try {
                Write-Host "Compiling VBA Project..."
                if ($null -ne $excel.VBE) {
                    Write-Host "  VBE object found."

                    $missing = [System.Reflection.Missing]::Value
                    $btn = $excel.VBE.CommandBars.Item("Menu Bar").FindControl($missing, 578, $missing, $missing, $true)

                    if ($null -ne $btn) {
                        Write-Host "  Found '$($btn.Caption)' button (Enabled: $($btn.Enabled))."
                        if ($btn.Enabled) {
                            $excel.DisplayAlerts = $false

                            try {
                                Write-Host "  Executing compile..."
                                $btn.Execute()
                            } catch {
                                Write-Host "  Execute() threw an exception: $($_.Exception.Message)"
                            }

                            # Polling loop: Wait for compilation to complete (up to 1.5 seconds)
                            $waitLimit = 30
                            $waitCount = 0
                            while ($btn.Enabled -and $waitCount -lt $waitLimit) {
                                Start-Sleep -Milliseconds 50
                                $waitCount++
                            }

                            if ($btn.Enabled) {
                                Write-Host "  ERROR: VBA Compilation failed (Button still enabled)." -ForegroundColor Red
                                if ($null -ne $global:BeaverBuildLog) {
                                    $global:BeaverBuildLog.compileResults.status = "failure"
                                }

                                try {
                                    $activePane = $excel.VBE.ActiveCodePane
                                    if ($null -ne $activePane) {
                                        $modName = $activePane.CodeModule.Name
                                        $startLine = 0
                                        $startCol = 0
                                        $endLine = 0
                                        $endCol = 0
                                        $activePane.GetSelection([ref]$startLine, [ref]$startCol, [ref]$endLine, [ref]$endCol)

                                        $errorLineText = $activePane.CodeModule.Lines($startLine, 1).Trim()

                                        $diskFile = $null
                                        if ($modName -eq "ThisWorkbook") {
                                            if (Test-Path $desktopThisWorkbookCls) {
                                                $diskFile = Get-Item $desktopThisWorkbookCls
                                            }
                                        } else {
                                            $diskFile = Get-ChildItem -Path $modulesDir -Recurse | Where-Object { $_.BaseName -eq $modName -and $_.Extension -match "\.(bas|cls|frm)$" } | Select-Object -First 1
                                        }

                                        if ($null -ne $diskFile) {
                                            $diskLines = Get-Content $diskFile.FullName
                                            $lastAttr = -1
                                            for ($l = 0; $l -lt $diskLines.Count; $l++) {
                                                if ($diskLines[$l] -match "^Attribute\s+") {
                                                    $lastAttr = $l
                                                }
                                            }
                                            $offset = $lastAttr + 1
                                            $diskErrorLine = $startLine + $offset

                                            Write-Host "  [Diagnostics] Source File: $($diskFile.FullName)" -ForegroundColor Yellow
                                            Write-Host "  [Diagnostics] Error at Disk Line $diskErrorLine" -ForegroundColor Yellow

                                            $contextStart = [Math]::Max(1, $diskErrorLine - 3)
                                            $contextEnd = [Math]::Min($diskLines.Count, $diskErrorLine + 3)

                                            Write-Host "  [Diagnostics] --- Code Context ---" -ForegroundColor Yellow
                                            for ($l = $contextStart; $l -le $contextEnd; $l++) {
                                                $prefix = if ($l -eq $diskErrorLine) { ">>> " } else { "    " }
                                                $color = if ($l -eq $diskErrorLine) { "Red" } else { "DarkGray" }
                                                Write-Host ("{0}{1:D3}: {2}" -f $prefix, $l, $diskLines[$l - 1]) -ForegroundColor $color
                                            }
                                            Write-Host "  [Diagnostics] --------------------" -ForegroundColor Yellow

                                            if ($null -ne $global:BeaverBuildLog) {
                                                $global:BeaverBuildLog.compileResults.errorDetails = [ordered]@{
                                                    moduleName = $modName
                                                    errorLine = $diskErrorLine
                                                    errorText = $errorLineText
                                                    file = $diskFile.FullName
                                                }
                                            }
                                            throw "VBA Compilation failed in module '$modName' (Source: $($diskFile.FullName)) at line $($diskErrorLine): '$errorLineText'. Please fix the syntax or missing definitions."
                                        } else {
                                            Write-Host "  [Diagnostics] Module: $modName" -ForegroundColor Yellow
                                            Write-Host "  [Diagnostics] Line $($startLine): $errorLineText" -ForegroundColor Yellow
                                            if ($null -ne $global:BeaverBuildLog) {
                                                $global:BeaverBuildLog.compileResults.errorDetails = [ordered]@{
                                                    moduleName = $modName
                                                    errorLine = $startLine
                                                    errorText = $errorLineText
                                                    file = $null
                                                }
                                            }
                                            throw "VBA Compilation failed in module '$modName' at line $($startLine): '$errorLineText'. Please fix the syntax or missing definitions."
                                        }
                                    }

                                    Write-Host "  [Diagnostics] No ActiveCodePane found after failure." -ForegroundColor Yellow
                                    if ($null -ne $global:BeaverBuildLog) {
                                        $global:BeaverBuildLog.compileResults.errorDetails = "VBA Compilation failed. Check your code for syntax or definition errors."
                                    }
                                    throw "VBA Compilation failed. Check your code for syntax or definition errors."
                                } catch {
                                    if ($_.Exception.Message -match "VBA Compilation failed") {
                                        throw $_.Exception.Message
                                    }

                                    Write-Host "  [Diagnostics] Error retrieving active pane: $($_.Exception.Message)" -ForegroundColor Red
                                    if ($null -ne $global:BeaverBuildLog) {
                                        $global:BeaverBuildLog.compileResults.errorDetails = "VBA Compilation failed: $($_.Exception.Message)"
                                    }
                                    throw "VBA Compilation failed. Check your code for 'Variable not defined' or syntax errors."
                                }
                            } else {
                                Write-Host "  Compilation successful." -ForegroundColor Green
                                if ($null -ne $global:BeaverBuildLog) {
                                    $global:BeaverBuildLog.compileResults.status = "success"
                                }
                            }
                        } else {
                            Write-Host "  Project already compiled." -ForegroundColor Gray
                            if ($null -ne $global:BeaverBuildLog) {
                                $global:BeaverBuildLog.compileResults.status = "success"
                            }
                        }
                    } else {
                        Write-Host "  'Compile Project' button NOT found on VBE Menu Bar." -ForegroundColor Yellow
                    }
                } else {
                    Write-Host "  VBE object NOT found." -ForegroundColor Yellow
                }
            } finally {
                if ($excelPid -gt 0) {
                    $null = New-Item -Path $signalFile -ItemType File -Force
                    $compileError = [WindowScraper]::StopAndGetResult()
                    if ($compileError) {
                        Write-Host "  ERROR: VBE Compilation popup dialog detected." -ForegroundColor Red
                        $cleanError = $compileError -replace "\r\n+", " | " -replace "\s+", " "
                        Write-Host "  [Diagnostics] $cleanError" -ForegroundColor Yellow
                    }
                }
                if (Test-Path $signalFile) {
                    Remove-Item -Path $signalFile -Force -ErrorAction SilentlyContinue
                }
            }

            if ($null -ne $excel.VBE) {
                try {
                    $excel.VBE.MainWindow.Visible = $false
                } catch { }
            }

            if ($forceFullBuild -or $manifestChanged) {
                $workbook.Saved = $false
                $workbook.Save()
                Write-Host "  Closing workbook to release file lock for Ribbon XML injection..." -ForegroundColor Yellow
                $workbook.Close($true)
                Release-ComObjectSafely $workbook
                $workbook = $null
            } else {
                Write-Host "  Deferring workbook save until success confirmation (Transactional Build)..." -ForegroundColor Yellow
            }
            
            Release-ComObjectSafely $activePane
            Release-ComObjectSafely $btn
            Release-ComObjectSafely $cb
            Release-ComObjectSafely $commandBars
            Release-ComObjectSafely $vbaProject
            $activePane = $null
            $btn = $null
            $cb = $null
            $commandBars = $null
            $vbaProject = $null
            $savedStatus = if ($forceFullBuild -or $manifestChanged) { "modules imported and workbook saved" } else { "modules imported (save deferred)" }
            return $savedStatus
        } finally {
            if ($null -ne $excel) {
                try {
                    $excel.ScreenUpdating = $origScreenUpdating
                    $excel.EnableEvents = $origEvents
                    $excel.Calculation = $origCalculation
                } catch {}
            }
            if ($workbook) {
                if ($forceFullBuild -or $manifestChanged) {
                    try { $workbook.Close($false) } catch { }
                }
                Release-ComObjectSafely $workbook
                $workbook = $null
            }
            Release-ComObjectSafely $activePane
            Release-ComObjectSafely $btn
            Release-ComObjectSafely $cb
            Release-ComObjectSafely $commandBars
            Release-ComObjectSafely $vbaProject
            $activePane = $null
            $btn = $null
            $cb = $null
            $commandBars = $null
            $vbaProject = $null
        }
    } | Out-Null

    Invoke-Stage -Stage "ribbon_injection" -Action {
        if ($forceFullBuild -or $manifestChanged) {
            Update-RibbonInWorkbook -WorkbookPath $excelPath -RibbonXmlPath $ribbonXmlPath
            return "customUI14.xml refreshed"
        } else {
            return "skipped (ribbon unchanged)"
        }
    } | Out-Null

    # Save successful build state to persist actual final hashes of all files on disk
    Save-BuildState -FileHashes (Get-SourceFileHashes -Force)

    Write-StageSummary
    Save-BuildLog -Status "success"
} catch {
    Stop-Script $_.Exception.Message
} finally {
    if (-not $global:BeaverOrchestratorActive -and $null -ne $sharedExcel) {
        Close-ExcelWorkbookSession -Excel $sharedExcel `
                                   -WasAlreadyOpen $excelWasAlreadyOpen `
                                   -KeepAlive $global:BeaverKeepAliveActive `
                                   -WorkbookPath $excelPath `
                                   -SaveChanges
        $sharedExcel = $null
    }
}
