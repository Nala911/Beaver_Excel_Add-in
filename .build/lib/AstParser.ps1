# Script:   AstParser.ps1
# Purpose:  VBA syntax analysis, regex dependency extraction, and transitive impact tracing for Beaver build pipeline.
# ==============================================================================

Set-StrictMode -Version Latest

function Get-ModuleDependencies {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) { return @() }
    
    $deps = @()
    try {
        $content = [System.IO.File]::ReadLines($FilePath)
        foreach ($line in $content) {
            if ($line -match "^\s*'\s*@Dependencies:\s*(.*)") {
                $depsStr = $Matches[1].Trim()
                if ($depsStr -eq "None" -or $depsStr -eq "") {
                    break
                }
                $deps = @($depsStr -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                break
            }
        }
    } catch {}
    return $deps
}

function Get-TestProcedureDependencies {
    param(
        [string]$FilePath,
        [string[]]$ComponentNames
    )
    if (-not (Test-Path $FilePath) -or $ComponentNames.Count -eq 0) { return @{} }
    
    $testDeps = @{}
    try {
        $lines = [System.IO.File]::ReadLines($FilePath)
        $currentTest = $null
        $currentTestDeps = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        
        # Escape component names for regex matching
        $escapedComps = @()
        foreach ($comp in $ComponentNames) {
            $escapedComps += [regex]::Escape($comp)
        }
        $componentPattern = "\b(" + ($escapedComps -join "|") + ")\b"
        
        foreach ($line in $lines) {
            if ($line -match "^\s*Public Sub (Test_[A-Za-z0-9_]+)\s*\(") {
                $currentTest = $Matches[1]
                $currentTestDeps.Clear()
            } elseif ($line -match "^\s*End Sub") {
                if ($null -ne $currentTest) {
                    $testDeps[$currentTest] = @($currentTestDeps)
                    $currentTest = $null
                }
            } elseif ($null -ne $currentTest) {
                # Scan for component references
                if ($line -match $componentPattern) {
                    $matches = [regex]::Matches($line, $componentPattern)
                    foreach ($m in $matches) {
                        [void]$currentTestDeps.Add($m.Value)
                    }
                }
                # Scan for string literals matching command names (e.g. CreateCommandContext("HelloWorld"))
                if ($line -match '"([A-Za-z0-9_]+)"') {
                    $matches = [regex]::Matches($line, '"([A-Za-z0-9_]+)"')
                    foreach ($m in $matches) {
                        $cmdName = $m.Groups[1].Value
                        if ($ComponentNames -contains "FeatCmd_$cmdName") {
                            [void]$currentTestDeps.Add("FeatCmd_$cmdName")
                        }
                    }
                }
            }
        }
    } catch {}
    return $testDeps
}

function Get-TransitiveImpact {
    param(
        [string[]]$ChangedComponents,
        $ProjectDeps # ComponentName -> [Dependencies]
    )
    
    $impacted = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $queue = [System.Collections.Generic.Queue[string]]::new()
    
    foreach ($comp in $ChangedComponents) {
        [void]$impacted.Add($comp)
        $queue.Enqueue($comp)
    }
    
    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        
        # Find which components list $current as a dependency
        foreach ($prop in $ProjectDeps.PSObject.Properties) {
            $key = $prop.Name
            $deps = @($prop.Value)
            if ($deps -contains $current) {
                if (-not $impacted.Contains($key)) {
                    [void]$impacted.Add($key)
                    $queue.Enqueue($key)
                }
            }
        }
    }
    
    return @($impacted)
}

function Get-ImpactedTests {
    param(
        [string[]]$ChangedFiles,
        [string[]]$DeletedFiles
    )
    
    $buildState = Get-BuildState
    if ($null -eq $buildState -or $null -eq $buildState.Metadata) {
        return @()
    }
    
    # 1. Collect all project component dependencies and test dependencies from metadata
    $projectDeps = [ordered]@{}
    $allTestDeps = @{}
    $changedTestFiles = @()
    
    foreach ($prop in $buildState.Metadata.PSObject.Properties) {
        $relPath = $prop.Name
        $meta = $prop.Value
        $compName = [System.IO.Path]::GetFileNameWithoutExtension($relPath)
        
        if ($meta.PSObject.Properties.Name -contains "Dependencies") {
            $projectDeps[$compName] = @($meta.Dependencies)
        }
        
        if ($meta.PSObject.Properties.Name -contains "TestDependencies" -and $null -ne $meta.TestDependencies) {
            foreach ($testProp in $meta.TestDependencies.PSObject.Properties) {
                $allTestDeps[$testProp.Name] = @($testProp.Value)
            }
        }
        
        if ($ChangedFiles -contains $relPath) {
            if ($meta.PSObject.Properties.Name -contains "Tests" -and $null -ne $meta.Tests -and $meta.Tests.Count -gt 0) {
                $changedTestFiles += $relPath
            }
        }
    }
    
    # 2. Identify changed components (excluding test files or features.json)
    $changedComps = @()
    foreach ($file in ($ChangedFiles + $DeletedFiles)) {
        if ($file -match "\.(bas|cls|frm)$") {
            $relPath = $file.Replace("\", "/")
            $meta = $buildState.Metadata.PSObject.Properties[$relPath]
            # Skip if it is primarily a test file itself (we run all tests inside it anyway)
            if ($null -ne $meta -and $meta.Value.PSObject.Properties.Name -contains "Tests" -and $meta.Value.Tests.Count -gt 0) {
                continue
            }
            $compName = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $changedComps += $compName
        }
    }
    
    # If no modules changed (only test files changed), we don't need transitive impact
    $impactedComps = @()
    if ($changedComps.Count -gt 0) {
        $impactedComps = Get-TransitiveImpact -ChangedComponents $changedComps -ProjectDeps $projectDeps
    }
    
    $testsToRun = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    
    # 3. Add tests that reference any impacted components
    foreach ($testName in $allTestDeps.Keys) {
        $deps = $allTestDeps[$testName]
        foreach ($dep in $deps) {
            if ($impactedComps -contains $dep -or $changedComps -contains $dep) {
                [void]$testsToRun.Add($testName)
                break
            }
        }
    }
    
    # 4. Add all tests from changed test files
    foreach ($file in $changedTestFiles) {
        $relPath = $file.Replace("\", "/")
        $meta = $buildState.Metadata.PSObject.Properties[$relPath].Value
        if ($null -ne $meta.Tests) {
            foreach ($t in $meta.Tests) {
                [void]$testsToRun.Add($t)
            }
        }
    }
    
    return @($testsToRun)
}

function Get-AllTestProcedures {
    param([string]$SourceDir)
    
    $testProcedures = @()
    if (-not (Test-Path $SourceDir)) { return $testProcedures }
    
    $projectRootVar = Get-Variable -Name "projectRoot" -ErrorAction SilentlyContinue
    $resolvedProjectRoot = if ($null -ne $projectRootVar) { $projectRootVar.Value } else { Split-Path (Split-Path $PSScriptRoot -Parent) -Parent }
    
    $buildState = Get-BuildState
    $global:BeaverTestManifestCache = @{}
    
    $moduleFiles = @(Get-ChildItem -Path $SourceDir -Filter *.bas -Recurse)
    foreach ($file in $moduleFiles) {
        if ($file.Name -eq "Lib_TestManifest.bas") { continue }
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $relPath = $file.FullName.Substring($resolvedProjectRoot.Length + 1).Replace("\", "/")
        
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
