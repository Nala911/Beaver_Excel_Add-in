# Script:   Linter.ps1
# Purpose:  Syntax validation, custom style guides checks, and form integrity validation
#           helpers consolidated into a single-pass implementation.
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Invoke-VbaLint {
    param (
        [string]$SourceDir,
        [string[]]$FilesToProcess,
        [switch]$AutoFix
    )
    Write-Host "Linting VBA Files (Single Pass)..." -ForegroundColor Cyan

    $projectRoot = Split-Path $PSScriptRoot -Parent
    $buildState = Get-BuildState
    if (-not (Get-Variable -Name "BeaverLintStatusCache" -Scope Global -ErrorAction SilentlyContinue)) { $global:BeaverLintStatusCache = @{} }
    $global:BeaverFileContentCache = @{}

    $vbaFiles = @()
    if ($null -ne $FilesToProcess -and $FilesToProcess.Count -gt 0) {
        foreach ($file in $FilesToProcess) {
            if ($file -match "\.(bas|cls|frm)$") {
                $absPath = Join-Path $projectRoot $file
                if (Test-Path $absPath) {
                    $vbaFiles += Get-Item $absPath
                }
            }
        }
    } else {
        $vbaFiles = @(Get-ChildItem -Path $SourceDir -Include *.bas, *.cls, *.frm -Recurse)
        $thisWorkbook = Join-Path $projectRoot "ThisWorkbook.cls"
        if (Test-Path $thisWorkbook) { $vbaFiles += Get-Item $thisWorkbook }
    }

    # Define the lint worker block as a string literal (to enable marshalling across runspaces)
    $lintBlockText = @'
        param (
            [object]$file,
            [string]$projectRoot,
            [object]$buildState,
            [bool]$autoFix
        )

        $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace("\", "/")
        $fileName = $file.Name
        
        # Check cache
        $cachedPassed = $false
        if ($null -ne $buildState -and $null -ne $buildState.Metadata -and $null -ne $buildState.Metadata.PSObject.Properties[$relPath]) {
            $meta = $buildState.Metadata.PSObject.Properties[$relPath].Value
            if ($meta.PSObject.Properties['Length'] -and $meta.PSObject.Properties['LastWriteTime'] -and $meta.PSObject.Properties['LintPassed'] -and
                $meta.Length -eq $file.Length -and $meta.LastWriteTime -eq $file.LastWriteTime.ToFileTime().ToString() -and $meta.LintPassed -eq $true) {
                $cachedPassed = $true
            }
        }
        
        if ($cachedPassed) {
            return [pscustomobject]@{
                RelPath = $relPath
                Cached = $true
                Passed = $true
                Errors = @()
                Warnings = @()
                RawLines = $null
            }
        }
        
        $rawLines = [System.IO.File]::ReadAllLines($file.FullName)
        
        # Check Option Explicit and @Module header on raw content
        $contentStr = $rawLines -join "`r`n"
        $missingOptionExplicit = ($contentStr -notmatch "(?m)^Option Explicit")
        $missingModuleHeader = ($fileName -ne "Test_Manifest.bas" -and $contentStr -notmatch "' @Module:")

        if ($missingOptionExplicit -or $missingModuleHeader) {
            if ($autoFix) {
                $newLines = [System.Collections.Generic.List[string]]::new($rawLines)
                
                # Find last attribute line index
                $insertIndex = 0
                for ($idx = 0; $idx -lt $newLines.Count; $idx++) {
                    if ($newLines[$idx] -match "^Attribute\s+") {
                        $insertIndex = $idx + 1
                    }
                }
                
                # Prepend Option Explicit if missing
                if ($missingOptionExplicit) {
                    $newLines.Insert($insertIndex, "Option Explicit")
                    $newLines.Insert($insertIndex + 1, "")
                    $insertIndex += 2
                }
                
                # Prepend @Module header if missing
                if ($missingModuleHeader) {
                    $category = "Feature"
                    if ($fileName -match "^Infra_") { $category = "Infrastructure" }
                    elseif ($fileName -match "^Lib_") { $category = "Library" }
                    
                    $newLines.Insert($insertIndex, "' @Module: $($file.BaseName)")
                    $newLines.Insert($insertIndex + 1, "' @Category: $category")
                    $newLines.Insert($insertIndex + 2, "' @Description: Automatically generated header template.")
                    $newLines.Insert($insertIndex + 3, "")
                    $insertIndex += 4
                }
                
                [System.IO.File]::WriteAllLines($file.FullName, $newLines)
                $rawLines = $newLines.ToArray()
                $contentStr = $rawLines -join "`r`n"
                
                # Reset flags
                $missingOptionExplicit = $false
                $missingModuleHeader = $false
            }
        }

        $errors = New-Object System.Collections.ArrayList
        $warnings = New-Object System.Collections.ArrayList
        $allPassed = $true

        if ($missingOptionExplicit) {
            [void]$errors.Add([ordered]@{
                file = $fileName
                type = "enhanced"
                message = "Missing 'Option Explicit' at the top of the file."
                line = 1
            })
            $allPassed = $false
        }

        if ($missingModuleHeader) {
            [void]$errors.Add([ordered]@{
                file = $fileName
                type = "enhanced"
                message = "Missing '@Module' metadata header."
                line = 1
            })
            $allPassed = $false
        }

        $generatedFiles = @(
            "Test_Manifest.bas",
            "UI_Ribbon.bas",
            "UI_Hotkeys.bas",
            "Lib_HelpManifest.bas",
            "Lib_UdfRegistry.bas",
            "Infra_CommandRegistry.bas"
        )
        if ($generatedFiles -contains $fileName) {
            [void]$warnings.Add("Warning: This file is auto-generated and managed by BeaverAddin Agent. Manual changes will be overwritten on build.")
        }

        # Normalize line continuations
        $content = @()
        $originalLineNumbers = @()
        $buffer = ""
        $bufferStartLine = 1
        
        for ($i = 0; $i -lt $rawLines.Count; $i++) {
            $line = $rawLines[$i]
            if ($line -match "\s+_\s*(?:'.*)?$") {
                if ($buffer -eq "") { $bufferStartLine = $i + 1 }
                $buffer += ($line -replace "\s+_\s*(?:'.*)?$", " ")
            } else {
                $content += ($buffer + $line)
                if ($buffer -eq "") {
                    $originalLineNumbers += ($i + 1)
                } else {
                    $originalLineNumbers += $bufferStartLine
                }
                $buffer = ""
            }
        }

        $subStartRegex = '^\s*(?:Public |Private |Static )?Sub\s+'
        $subEndRegex = '^\s*End Sub'
        $funcStartRegex = '^\s*(?:Public |Private |Static )?Function\s+'
        $funcEndRegex = '^\s*End Function'
        $propStartRegex = '^\s*(?:Public |Private )?Property\s+(?:Get|Let|Set)\s+'
        $propEndRegex = '^\s*End Property'
        $ifStartRegex = '^\s*If\s+.*Then\s*(?:''.*)?$'
        $ifEndRegex = '^\s*End If'

        $subStack = New-Object System.Collections.Generic.List[int]
        $funcStack = New-Object System.Collections.Generic.List[int]
        $propStack = New-Object System.Collections.Generic.List[int]
        $ifStack = New-Object System.Collections.Generic.List[int]

        for ($i = 0; $i -lt $content.Count; $i++) {
            $line = $content[$i]
            $lineNum = $originalLineNumbers[$i]
            
            if ($line -eq "" -or $line.Trim() -eq "") { continue }

            # --- Syntax Validation Stack Checks ---
            if ($line.IndexOf("If", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($line -match $ifStartRegex) {
                    $ifStack.Add($lineNum)
                } elseif ($line -match $ifEndRegex) {
                    if ($ifStack.Count -gt 0) {
                        $ifStack.RemoveAt($ifStack.Count - 1)
                    } else {
                        [void]$errors.Add([ordered]@{
                            file = $fileName
                            type = "syntax"
                            message = "Unexpected 'End If' (No matching start found)"
                            line = $lineNum
                        })
                        $allPassed = $false
                    }
                }
            }

            if ($line.IndexOf("Sub", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($line -match $subStartRegex) {
                    $subStack.Add($lineNum)
                } elseif ($line -match $subEndRegex) {
                    if ($subStack.Count -gt 0) {
                        $subStack.RemoveAt($subStack.Count - 1)
                    } else {
                        [void]$errors.Add([ordered]@{
                            file = $fileName
                            type = "syntax"
                            message = "Unexpected 'End Sub' (No matching start found)"
                            line = $lineNum
                        })
                        $allPassed = $false
                    }
                }
            }

            if ($line.IndexOf("Function", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($line -match $funcStartRegex) {
                    $funcStack.Add($lineNum)
                } elseif ($line -match $funcEndRegex) {
                    if ($funcStack.Count -gt 0) {
                        $funcStack.RemoveAt($funcStack.Count - 1)
                    } else {
                        [void]$errors.Add([ordered]@{
                            file = $fileName
                            type = "syntax"
                            message = "Unexpected 'End Function' (No matching start found)"
                            line = $lineNum
                        })
                        $allPassed = $false
                    }
                }
            }

            if ($line.IndexOf("Property", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($line -match $propStartRegex) {
                    $propStack.Add($lineNum)
                } elseif ($line -match $propEndRegex) {
                    if ($propStack.Count -gt 0) {
                        $propStack.RemoveAt($propStack.Count - 1)
                    } else {
                        [void]$errors.Add([ordered]@{
                            file = $fileName
                            type = "syntax"
                            message = "Unexpected 'End Property' (No matching start found)"
                            line = $lineNum
                        })
                        $allPassed = $false
                    }
                }
            }

            # --- Rule A: Enforce Spill-Safe Formula Properties ---
            if ($line.IndexOf("Formula", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($fileName -ne "Lib_JsonConverter.bas" -and $line -match '\.\bFormula\b' -and $line -notmatch '".*\.Formula.*"' -and $line -notmatch '^\s*\''' -and $line -notmatch '\.Formula2' -and $line -notmatch '\.FormulaArray') {
                    [void]$errors.Add([ordered]@{
                        file = $fileName
                        type = "enhanced"
                        message = "Range.Formula usage detected. Use Range.Formula2 instead to prevent spill errors."
                        line = $lineNum
                    })
                    $allPassed = $false
                }
            }

            # --- Rule B: Multi-cell Range Property Null Check ---
            if ($line.IndexOf("NumberFormat", [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or 
                $line.IndexOf("Name", [System.StringComparison]::Ordinal) -ge 0 -or 
                $line.IndexOf("Size", [System.StringComparison]::Ordinal) -ge 0) {
                if ($line -match '\b(CStr|CInt|CLng|CDbl|CSng|CBool|CDate|CVar)\s*\(\s*(?!(?:cell\b|\w+\.Cells\b|\w+Cells\b))[a-zA-Z0-9_\.]+\.(?:NumberFormat|Font\.(?:Name|Size))\s*\)' -and $line -notmatch '^\s*''') {
                    [void]$errors.Add([ordered]@{
                        file = $fileName
                        type = "enhanced"
                        message = "Direct string/value conversion on range property without IsNull check. Mixed ranges return Null, causing Error 94."
                        line = $lineNum
                    })
                    $allPassed = $false
                }
            }

            # --- Rule C: Collection Mutation Loop Direction ---
            if ($line.IndexOf("For", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($line -match '\bFor\s+(\w+)\s*=\s*\d+\s+To\s+(?:(?:\w+\.)*Count|\w+)\b(?!\s+Step\s+-1)' -and $line -notmatch '^\s*''') {
                    $idxVar = $Matches[1]
                    $j = $i + 1
                    $hasDeletion = $false
                    while ($j -lt $content.Count -and $content[$j] -notmatch "\bNext\b") {
                        if ($content[$j] -match "\b$idxVar\b" -and ($content[$j] -match '\bDelete\b|\bRemove\b|\bRemoveAt\b') -and $content[$j] -notmatch '^\s*''') {
                            $hasDeletion = $true
                            break
                        }
                        $j++
                    }
                    if ($hasDeletion) {
                        [void]$errors.Add([ordered]@{
                            file = $fileName
                            type = "enhanced"
                            message = "Forward iteration loop with mutation detected. Use backward iteration 'For $idxVar = ... To 1 Step -1' instead to prevent skipping bugs."
                            line = $lineNum
                        })
                        $allPassed = $false
                    }
                }
            }

            # --- Rule D: Implicit Reference Check ---
            if ($fileName -ne "Lib_JsonConverter.bas" -and $fileName -notmatch "Test_") {
                if ($line.IndexOf("Range", [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                    $line.IndexOf("Cells", [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                    $line.IndexOf("Rows", [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                    $line.IndexOf("Columns", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    
                    if ($line -match '(?<!\bAs\s+)(?<!\w\.)(?<!\.)\b(Range|Cells|Rows|Columns)\b' -and 
                        $line -notmatch '^\s*''' -and
                        $line -notmatch '^\s*(?:Public |Private |Static )?(?:Sub |Function |Property |Type |Enum )\b' -and
                        $line -notmatch '".*\b(Range|Cells|Rows|Columns)\b.*"') {
                        
                        [void]$warnings.Add("Warning: Unqualified reference to '$($Matches[1])' at line $lineNum. Use explicit worksheet qualification (e.g. ws.$($Matches[1])) to prevent ActiveSheet bugs.")
                    }
                }
            }

            # --- Rule E: Localized Sheet Name Warning ---
            if ($fileName -ne "Lib_JsonConverter.bas" -and $fileName -notmatch "Test_") {
                if ($line -match '"(?:Sheet|Tabelle|Feuille|Hoja|Foglio|Planilha|Flik|Tabell)\d+"' -and $line -notmatch '^\s*''') {
                    [void]$warnings.Add("Warning: Hardcoded localized sheet name $($Matches[0]) detected at line $lineNum. This will fail in non-English Excel environments.")
                }
            }

            # --- Context Tracking Check ---
            if ($line.IndexOf("Sub", [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or 
                $line.IndexOf("Function", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $isProcStart = $false
                $procName = ""
                $isICommandExecute = $false
                
                if ($line -match "^\s*Public (?:Sub|Function)\s+([a-zA-Z0-9_]+)") {
                    $isProcStart = $true
                    $procName = $Matches[1]
                } elseif ($line -match "^\s*Private Sub (ICommand_Execute)\b") {
                    $isProcStart = $true
                    $procName = "ICommand_Execute"
                    $isICommandExecute = $true
                }

                if ($isProcStart) {
                    if ($procName -match "^(?:Workbook_|Worksheet_|App_)" -or $fileName -eq "Lib_JsonConverter.bas" -or $fileName -match "^Lib_[a-zA-Z0-9_]+Function\.bas$" -or $fileName -match "^(?:Infra_Error\.(bas|cls)|Infra_ContextTracker\.cls|Infra_Diagnostics\.bas|Infra_OperationContext\.cls|AppContainer\.cls|Infra_Config\.(cls|bas)|Infra_ConfigModel\.cls|I[A-Z][a-zA-Z0-9_\-]*\.cls|Infra_AppStateGuard\.cls|Infra_AppState\.bas|Infra_ValueConversion\.bas)$") {
                        continue
                    }

                    $className = $file.BaseName
                    $escProc = [regex]::Escape($procName)
                    $escClass = [regex]::Escape($className)
                    $procNamePattern = "(?:$escClass\.)?$escProc"

                    $j = $i + 1
                    $foundPush = $false
                    $foundPop = $false
                    $foundErrorGoto = $false
                    $foundHandleError = $false
                    
                    while ($j -lt $content.Count -and $content[$j] -notmatch "^\s*End (?:Sub|Function)") {
                        $lText = $content[$j]
                        if ($isICommandExecute) {
                            if ($lText -match 'PushContext\s+"\w+\.Execute"' -or $lText -match 'Infra_Error\.Track\("\w+\.Execute"\)') { $foundPush = $true }
                            if ($lText -match 'PopContext' -or $lText -match 'Infra_Error\.Track') { $foundPop = $true }
                            if ($lText -match 'On Error GoTo\s+\w+') { $foundErrorGoto = $true }
                            if ($lText -match 'HandleError(?:Detailed)?\s+"\w+\.Execute"') { $foundHandleError = $true }
                        } else {
                            if ($lText -match "PushContext\s+""$procNamePattern""" -or $lText -match "Infra_Error\.Track\(\s*""$procNamePattern""\s*\)") { $foundPush = $true }
                            if ($lText -match "PopContext" -or $lText -match "Infra_Error\.Track\(\s*""$procNamePattern""\s*\)") { $foundPop = $true }
                            if ($lText -match "On Error GoTo\s+\w+") { $foundErrorGoto = $true }
                            if ($lText -match "(?:Infra_Error\.)?HandleError(?:Detailed)?\s+""$procNamePattern""") { $foundHandleError = $true }
                        }
                        $j++
                    }

                    if (-not $foundPush) {
                        $msgName = if ($isICommandExecute) { "ICommand_Execute (with context ending in .Execute)" } else { "Procedure '$procName'" }
                        [void]$errors.Add([ordered]@{
                            file = $fileName
                            type = "enhanced"
                            message = "$msgName missing context tracking (PushContext or Track)."
                            line = $lineNum
                        })
                        $allPassed = $false
                    }
                    if (-not $foundPop) {
                        [void]$errors.Add([ordered]@{
                            file = $fileName
                            type = "enhanced"
                            message = "Procedure '$procName' missing 'PopContext' (or RAII Track tracker)."
                            line = $lineNum
                        })
                        $allPassed = $false
                    }
                    if (-not $foundErrorGoto) {
                        [void]$errors.Add([ordered]@{
                            file = $fileName
                            type = "enhanced"
                            message = "Procedure '$procName' missing 'On Error GoTo'."
                            line = $lineNum
                        })
                        $allPassed = $false
                    }
                    if (-not $foundHandleError) {
                        $expectedLabel = if ($isICommandExecute) { "HandleError ""[CommandName].Execute""" } else { "HandleError ""$procName""" }
                        [void]$errors.Add([ordered]@{
                            file = $fileName
                            type = "enhanced"
                            message = "Procedure '$procName' missing '$expectedLabel'."
                            line = $lineNum
                        })
                        $allPassed = $false
                    }
                }
            }
        }

        # Check final stack sizes
        foreach ($startLine in $subStack) {
            [void]$errors.Add([ordered]@{
                file = $fileName
                type = "syntax"
                message = "Mismatched 'Sub' (No matching end found)"
                line = $startLine
            })
            $allPassed = $false
        }
        foreach ($startLine in $funcStack) {
            [void]$errors.Add([ordered]@{
                file = $fileName
                type = "syntax"
                message = "Mismatched 'Function' (No matching end found)"
                line = $startLine
            })
            $allPassed = $false
        }
        foreach ($startLine in $propStack) {
            [void]$errors.Add([ordered]@{
                file = $fileName
                type = "syntax"
                message = "Mismatched 'Property' (No matching end found)"
                line = $startLine
            })
            $allPassed = $false
        }
        foreach ($startLine in $ifStack) {
            [void]$errors.Add([ordered]@{
                file = $fileName
                type = "syntax"
                message = "Mismatched 'If' (No matching end found)"
                line = $startLine
            })
            $allPassed = $false
        }

        return [pscustomobject]@{
            RelPath = $relPath
            Cached = $false
            Passed = $allPassed
            Errors = $errors
            Warnings = $warnings
            RawLines = $rawLines
        }
'@

    # Execute workers (sequentially or in parallel depending on file count)
    $results = @()
    if ($vbaFiles.Count -le 2) {
        # Sequential processing (fast path for incremental builds, avoiding thread overhead)
        $sb = [scriptblock]::Create($lintBlockText)
        foreach ($file in $vbaFiles) {
            $results += & $sb -file $file -projectRoot $projectRoot -buildState $buildState -autoFix $AutoFix
        }
    } else {
        # Parallel processing (fast path for full/clean builds, using all CPU cores)
        $results = $vbaFiles | ForEach-Object -Parallel {
            $projectRoot = $using:projectRoot
            $buildState = $using:buildState
            $autoFix = $using:AutoFix
            
            $sb = [scriptblock]::Create($using:lintBlockText)
            & $sb -file $_ -projectRoot $projectRoot -buildState $buildState -autoFix $autoFix
        }
    }

    # Process and merge results on the main thread (thread-safe)
    $allPassed = $true
    foreach ($res in $results) {
        $relPath = $res.RelPath
        if ($res.Cached) {
            $global:BeaverLintStatusCache[$relPath] = $true
            continue
        }

        $fileName = [System.IO.Path]::GetFileName($relPath)

        if ($null -ne $global:BeaverBuildLog) {
            if ($global:BeaverBuildLog.lintResults.checkedFiles -notcontains $relPath) {
                [void]$global:BeaverBuildLog.lintResults.checkedFiles.Add($relPath)
            }
        }

        # Print collected warnings
        foreach ($warn in $res.Warnings) {
            Write-Host "  [$fileName] $warn" -ForegroundColor Yellow
        }

        $global:BeaverLintStatusCache[$relPath] = $res.Passed
        if (-not $res.Passed) {
            $allPassed = $false
            foreach ($err in $res.Errors) {
                Write-Host "  [$fileName] Error: $($err.message) at line $($err.line)" -ForegroundColor Red
                Add-LintError -File $fileName -Type $err.type -Message $err.message -Line $err.line
            }
        }

        if ($null -ne $res.RawLines) {
            $global:BeaverFileContentCache[(Join-Path $projectRoot $relPath)] = $res.RawLines
        }
    }

    return $allPassed
}

function Test-FormFilesValidity {
    param (
        [string]$SourceDir,
        [string[]]$FilesToProcess
    )
    Write-Host "Checking Form Companion Files..." -ForegroundColor Cyan

    $projectRoot = Split-Path $PSScriptRoot -Parent
    $buildState = Get-BuildState
    if (-not (Get-Variable -Name "BeaverLintStatusCache" -Scope Global -ErrorAction SilentlyContinue)) { $global:BeaverLintStatusCache = @{} }

    $frmFiles = @()
    if ($null -ne $FilesToProcess -and $FilesToProcess.Count -gt 0) {
        foreach ($file in $FilesToProcess) {
            if ($file -match "\.frm$") {
                $absPath = Join-Path $projectRoot $file
                if (Test-Path $absPath) {
                    $frmFiles += Get-Item $absPath
                }
            }
        }
    } else {
        $frmFiles = @(Get-ChildItem -Path $SourceDir -Include *.frm -Recurse)
    }
    $allPassed = $true
    foreach ($file in $frmFiles) {
        $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace("\", "/")
        if ($null -ne $global:BeaverBuildLog) {
            if ($global:BeaverBuildLog.lintResults.checkedFiles -notcontains $relPath) {
                [void]$global:BeaverBuildLog.lintResults.checkedFiles.Add($relPath)
            }
        }
        
        $cachedPassed = $false
        if ($null -ne $buildState -and $null -ne $buildState.Metadata -and $null -ne $buildState.Metadata.PSObject.Properties[$relPath]) {
            $meta = $buildState.Metadata.PSObject.Properties[$relPath].Value
            if ($meta.PSObject.Properties['Length'] -and $meta.PSObject.Properties['LastWriteTime'] -and $meta.PSObject.Properties['LintPassed'] -and
                $meta.Length -eq $file.Length -and $meta.LastWriteTime -eq $file.LastWriteTime.ToFileTime().ToString() -and $meta.LintPassed -eq $true) {
                $cachedPassed = $true
            }
        }
        
        if ($cachedPassed) {
            $global:BeaverLintStatusCache[$relPath] = $true
            continue
        }
        
        if (-not $global:BeaverLintStatusCache.ContainsKey($relPath)) {
            $global:BeaverLintStatusCache[$relPath] = $true
        }

        $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
        if (-not (Test-Path $frxPath)) {
            Write-Host "  [$($file.Name)] Error: Missing companion binary file (.frx). MSForms requires a .frx file to import successfully." -ForegroundColor Red
            Add-LintError -File $file.Name -Type "form" -Message "Missing companion binary file (.frx)." -Line 1
            $allPassed = $false
            $global:BeaverLintStatusCache[$relPath] = $false
        }
    }
    return $allPassed
}

function Invoke-TestCoverageAudit {
    param (
        [string]$SourceDir
    )
    Write-Host "Running Test Coverage Audit..." -ForegroundColor Cyan

    $projectRoot = Split-Path $PSScriptRoot -Parent
    $commandsDir = Join-Path $SourceDir "Commands"
    
    # 1. Find all FeatCmd_*.cls files
    if (-not (Test-Path $commandsDir)) {
        Write-Host "  No Commands directory found at: $commandsDir" -ForegroundColor Yellow
        return $true
    }
    
    $commandFiles = Get-ChildItem -Path $commandsDir -Filter "FeatCmd_*.cls"
    if ($commandFiles.Count -eq 0) {
        Write-Host "  No FeatCmd_*.cls modules found." -ForegroundColor Yellow
        return $true
    }
    
    # 2. Read all test files content in Tests directory
    $testDir = Join-Path $SourceDir "Tests"
    if (-not (Test-Path $testDir)) {
        Write-Host "  Test directory not found: $testDir" -ForegroundColor Red
        return $false
    }
    
    $testFiles = Get-ChildItem -Path $testDir -Filter "*.bas"
    if ($testFiles.Count -eq 0) {
        Write-Host "  No test modules found in: $testDir" -ForegroundColor Red
        return $false
    }
    
    $testContent = ""
    foreach ($testFile in $testFiles) {
        if ($testFile.Name -eq "Test_Manifest.bas") { continue }
        $testContent += [System.IO.File]::ReadAllText($testFile.FullName) + "`r`n"
    }
    
    # Exclusions: commands that are mock, interactive UI, or have other reasons for no direct headless test
    $exclusions = @("Dog", "ShowHelpCenter", "Duplicate")
    $allCovered = $true

    foreach ($file in $commandFiles) {
        $cmdName = $file.BaseName -replace "^FeatCmd_", ""
        if ($exclusions -contains $cmdName) { continue }
        
        # Check if the command name is referenced in the tests
        if ($testContent -notmatch "(?i)(?:\b|_)(Test_$cmdName|FeatCmd_$cmdName|$cmdName)(?:\b|_)") {
            Write-Host "  [WARNING] Test Coverage: Command '$($file.Name)' has no corresponding tests in any test modules under '$testDir'." -ForegroundColor Yellow
            $allCovered = $false
        }
    }
    
    if ($allCovered) {
        Write-Host "  All feature commands have corresponding unit tests." -ForegroundColor Green
    }
    return $true
}
