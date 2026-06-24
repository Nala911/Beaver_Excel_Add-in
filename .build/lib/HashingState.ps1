# Script:   HashingState.ps1
# Purpose:  MD5 file hashing, manifest structural hashing, state tracking, and change detection for Beaver build pipeline.
# ==============================================================================

Set-StrictMode -Version Latest

function Get-BuildState {
    $buildStatePathVar = Get-Variable -Name "buildStatePath" -ErrorAction SilentlyContinue
    $resolvedBuildStatePath = if ($null -ne $buildStatePathVar) { $buildStatePathVar.Value } else { Join-Path (Split-Path $PSScriptRoot -Parent) ".build_state.json" }

    if ($null -ne $global:BeaverBuildStateCache) {
        return $global:BeaverBuildStateCache
    }
    if (Test-Path $resolvedBuildStatePath) {
        try {
            $content = Get-Content $resolvedBuildStatePath -Raw
            $global:BeaverBuildStateCache = $content | ConvertFrom-Json
            return $global:BeaverBuildStateCache
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
    
    $projectRootVar = Get-Variable -Name "projectRoot" -ErrorAction SilentlyContinue
    $resolvedProjectRoot = if ($null -ne $projectRootVar) { $projectRootVar.Value } else { Split-Path (Split-Path $PSScriptRoot -Parent) -Parent }
    
    $buildStatePathVar = Get-Variable -Name "buildStatePath" -ErrorAction SilentlyContinue
    $resolvedBuildStatePath = if ($null -ne $buildStatePathVar) { $buildStatePathVar.Value } else { Join-Path (Split-Path $PSScriptRoot -Parent) ".build_state.json" }

    $featureManifestPathVar = Get-Variable -Name "featureManifestPath" -ErrorAction SilentlyContinue
    $resolvedFeatureManifestPath = if ($null -ne $featureManifestPathVar) { $featureManifestPathVar.Value } else { Join-Path $resolvedProjectRoot "features.json" }

    $metadata = [ordered]@{}
    
    # Pre-fetch all component names to use in test dependency parsing
    $componentNames = @()
    foreach ($relPath in $FileHashes.Keys) {
        if ($relPath -match "\.(bas|cls|frm)$") {
            $componentNames += [System.IO.Path]::GetFileNameWithoutExtension($relPath)
        }
    }
    
    foreach ($relPath in $FileHashes.Keys) {
        $absPath = Join-Path $resolvedProjectRoot $relPath
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
            
            # Parse standard module dependencies
            $deps = @()
            $hasDepsCached = $false
            $oldState = Get-BuildState
            if ($null -ne $oldState -and $null -ne $oldState.Metadata -and $null -ne $oldState.Metadata.PSObject.Properties[$relPath]) {
                $oldMeta = $oldState.Metadata.PSObject.Properties[$relPath].Value
                if ($oldMeta.Length -eq $file.Length -and $oldMeta.LastWriteTime -eq $file.LastWriteTime.ToFileTime().ToString()) {
                    if ($oldMeta.PSObject.Properties.Name -contains "Dependencies") {
                        $deps = @($oldMeta.Dependencies)
                        $hasDepsCached = $true
                    }
                }
            }
            if (-not $hasDepsCached -and $relPath -match "\.(bas|cls|frm)$") {
                $deps = Get-ModuleDependencies -FilePath $absPath
            }
            $meta["Dependencies"] = $deps
            
            # Parse test dependencies
            $testDeps = @{}
            $hasTestDepsCached = $false
            if ($null -ne $oldState -and $null -ne $oldState.Metadata -and $null -ne $oldState.Metadata.PSObject.Properties[$relPath]) {
                $oldMeta = $oldState.Metadata.PSObject.Properties[$relPath].Value
                if ($oldMeta.Length -eq $file.Length -and $oldMeta.LastWriteTime -eq $file.LastWriteTime.ToFileTime().ToString()) {
                    if ($oldMeta.PSObject.Properties.Name -contains "TestDependencies") {
                        $testDeps = $oldMeta.TestDependencies
                        $hasTestDepsCached = $true
                    }
                }
            }
            if (-not $hasTestDepsCached -and ($null -ne $meta["Tests"]) -and ($meta["Tests"].Count -gt 0)) {
                $testDeps = Get-TestProcedureDependencies -FilePath $absPath -ComponentNames $componentNames
            }
            $meta["TestDependencies"] = $testDeps
            
            # Lint status from global cache or previous build state
            if ($null -ne $global:BeaverLintStatusCache -and $global:BeaverLintStatusCache.ContainsKey($relPath)) {
                $meta["LintPassed"] = $global:BeaverLintStatusCache[$relPath]
            } else {
                if ($null -ne $oldState -and $null -ne $oldState.Metadata -and $null -ne $oldState.Metadata.PSObject.Properties[$relPath]) {
                    $oldMeta = $oldState.Metadata.PSObject.Properties[$relPath].Value
                    if ($oldMeta.Length -eq $file.Length -and $oldMeta.LastWriteTime -eq $file.LastWriteTime.ToFileTime().ToString()) {
                        if ($oldMeta.PSObject.Properties.Name -contains "LintPassed" -and $null -ne $oldMeta.LintPassed) {
                            $meta["LintPassed"] = [bool]$oldMeta.LintPassed
                        }
                    }
                }
            }
            
            $metadata[$relPath] = $meta
        }
    }
    
    $excelPid = 0
    if ($null -ne $global:BeaverSharedExcel) {
        $excelPid = Get-ExcelProcessId -ExcelApplication $global:BeaverSharedExcel
    }
    if ($excelPid -eq 0) {
        $oldState = Get-BuildState
        if ($null -ne $oldState -and $oldState.PSObject.Properties.Name -contains "ExcelPid") {
            $excelPid = $oldState.ExcelPid
        }
    }
    
    $state = [ordered]@{
        LastBuildTime = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssK")
        Files = $FileHashes
        Metadata = $metadata
        ManifestStructuralHash = (Get-ManifestStructuralHash -Path $resolvedFeatureManifestPath)
        ExcelPid = $excelPid
    }
    $stateJson = $state | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($resolvedBuildStatePath, $stateJson, [System.Text.Encoding]::UTF8)
    $global:BeaverBuildStateCache = $state
}

function Set-BuildStateTestsPassed {
    param(
        [bool]$Passed
    )
    $buildStatePathVar = Get-Variable -Name "buildStatePath" -ErrorAction SilentlyContinue
    $resolvedBuildStatePath = if ($null -ne $buildStatePathVar) { $buildStatePathVar.Value } else { Join-Path (Split-Path $PSScriptRoot -Parent) ".build_state.json" }

    $buildState = Get-BuildState
    if ($null -ne $buildState) {
        $buildState | Add-Member -NotePropertyName TestsPassed -NotePropertyValue $Passed -Force
        $stateJson = $buildState | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($resolvedBuildStatePath, $stateJson, [System.Text.Encoding]::UTF8)
        $global:BeaverBuildStateCache = $buildState
    }
}

function Get-VbaStructuralHash {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) { return "" }
    
    try {
        $lines = [System.IO.File]::ReadAllLines($FilePath)
        $structuralLines = [System.Collections.Generic.List[string]]::new()
        
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        $inCodeSection = $true
        if ($extension -eq ".frm" -or $extension -eq ".cls") {
            $inCodeSection = $false
        }
        
        foreach ($line in $lines) {
            if (-not $inCodeSection) {
                if ($line -match '^Attribute\s+VB_Name\s*=') {
                    $inCodeSection = $true
                }
                if ($line -match '^Attribute\s+') {
                    [void]$structuralLines.Add($line.Trim())
                }
                continue
            }
            
            $trimmed = $line.Trim()
            if ($trimmed -eq "") { continue }
            
            # Skip pure comment lines, but keep metadata comments starting with '@'
            if ($trimmed -match "^'" -and $trimmed -notmatch "^'\s*@") {
                continue
            }
            
            # Strip trailing comments (basic logic: find "'" outside quotes)
            $cleanLine = ""
            $inQuote = $false
            for ($i = 0; $i -lt $trimmed.Length; $i++) {
                $char = $trimmed[$i]
                if ($char -eq '"') {
                    $inQuote = -not $inQuote
                }
                if ($char -eq "'" -and -not $inQuote) {
                    # Inline comment starts here, strip the rest of the line (but check if it's metadata)
                    $commentText = $trimmed.Substring($i).Trim()
                    if ($commentText -match "^'\s*@") {
                        $cleanLine += $commentText
                    }
                    break
                }
                $cleanLine += $char
            }
            
            $cleanTrimmed = $cleanLine.Trim()
            if ($cleanTrimmed -eq "") { continue }
            
            # Collapse multiple spaces
            $collapsed = [regex]::Replace($cleanTrimmed, '\s+', ' ')
            [void]$structuralLines.Add($collapsed)
        }
        
        $joined = $structuralLines -join "`n"
        
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($joined)
        $hashBytes = $md5.ComputeHash($bytes)
        $md5.Dispose()
        
        $sb = [System.Text.StringBuilder]::new()
        foreach ($b in $hashBytes) {
            [void]$sb.Append($b.ToString("x2"))
        }
        return $sb.ToString().ToUpperInvariant()
    } catch {
        # Fallback to simple file hash
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $stream = [System.IO.File]::OpenRead($FilePath)
        $hashBytes = $md5.ComputeHash($stream)
        $stream.Close()
        $md5.Dispose()
        
        $sb = [System.Text.StringBuilder]::new()
        foreach ($b in $hashBytes) {
            [void]$sb.Append($b.ToString("x2"))
        }
        return $sb.ToString().ToUpperInvariant()
    }
}

function Get-FileHashOptimized {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) { return "" }
    
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    if ($extension -in @(".bas", ".cls", ".frm")) {
        return Get-VbaStructuralHash -FilePath $FilePath
    }
    
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $stream = [System.IO.File]::OpenRead($FilePath)
    $hashBytes = $md5.ComputeHash($stream)
    $stream.Close()
    $md5.Dispose()
    
    $sb = [System.Text.StringBuilder]::new()
    foreach ($b in $hashBytes) {
        [void]$sb.Append($b.ToString("x2"))
    }
    return $sb.ToString().ToUpperInvariant()
}


function Test-BuildStateFileChanged {
    param(
        [string]$RelativePath,
        $BuildState
    )
    $projectRootVar = Get-Variable -Name "projectRoot" -ErrorAction SilentlyContinue
    $resolvedProjectRoot = if ($null -ne $projectRootVar) { $projectRootVar.Value } else { Split-Path (Split-Path $PSScriptRoot -Parent) -Parent }

    $filePath = Join-Path $resolvedProjectRoot $RelativePath
    if (-not (Test-Path $filePath)) { return $false }
    
    $currentHash = Get-FileHashOptimized -FilePath $filePath
    if ($null -eq $BuildState -or $null -eq $BuildState.Files -or $null -eq $BuildState.Files.$RelativePath) {
        return $true
    }
    return ($BuildState.Files.$RelativePath -ne $currentHash)
}

function Get-SourceFileHashes {
    param(
        [switch]$Force
    )

    if (-not $Force -and $null -ne $global:BeaverSourceHashes) {
        return $global:BeaverSourceHashes
    }

    $projectRootVar = Get-Variable -Name "projectRoot" -ErrorAction SilentlyContinue
    $resolvedProjectRoot = if ($null -ne $projectRootVar) { $projectRootVar.Value } else { Split-Path (Split-Path $PSScriptRoot -Parent) -Parent }
    
    $featureManifestPathVar = Get-Variable -Name "featureManifestPath" -ErrorAction SilentlyContinue
    $resolvedFeatureManifestPath = if ($null -ne $featureManifestPathVar) { $featureManifestPathVar.Value } else { Join-Path $resolvedProjectRoot "features.json" }

    $desktopThisWorkbookClsVar = Get-Variable -Name "desktopThisWorkbookCls" -ErrorAction SilentlyContinue
    $resolvedThisWorkbook = if ($null -ne $desktopThisWorkbookClsVar) { $desktopThisWorkbookClsVar.Value } else { Join-Path $resolvedProjectRoot "ThisWorkbook.cls" }

    $modulesDirVar = Get-Variable -Name "modulesDir" -ErrorAction SilentlyContinue
    $resolvedModulesDir = if ($null -ne $modulesDirVar) { $modulesDirVar.Value } else { Join-Path $resolvedProjectRoot "Modules" }

    $hashes = @{}
    $buildState = Get-BuildState
    
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
    if (Test-Path $resolvedFeatureManifestPath) {
        $hashes["features.json"] = & $resolveHash $resolvedFeatureManifestPath "features.json"
    }
    
    # ThisWorkbook
    if (Test-Path $resolvedThisWorkbook) {
        $hashes["ThisWorkbook.cls"] = & $resolveHash $resolvedThisWorkbook "ThisWorkbook.cls"
    }
    
    # Modules
    if (Test-Path $resolvedModulesDir) {
        $vbaFiles = Get-ChildItem -Path $resolvedModulesDir -Include *.bas, *.cls, *.frm -Recurse
        foreach ($file in $vbaFiles) {
            $relPath = $file.FullName.Substring($resolvedProjectRoot.Length + 1).Replace("\", "/")
            $hashes[$relPath] = & $resolveHash $file.FullName $relPath
            
            # If it's a form, also include the companion FRX file hash if it exists
            if ($file.Extension -eq ".frm") {
                $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
                if (Test-Path $frxPath) {
                    $frxRelPath = $frxPath.Substring($resolvedProjectRoot.Length + 1).Replace("\", "/")
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

function Write-FileIfChanged {
    param(
        [string]$Path,
        [string]$Content
    )
    if (Test-Path $Path) {
        $existing = [System.IO.File]::ReadAllText($Path)
        if ($existing -eq $Content) {
            return $false
        }
    }
    $dir = Split-Path $Path
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::ASCII)
    return $true
}

function Get-FeatureManifest {
    param([string]$ManifestPath)
    if (-not (Test-Path $ManifestPath)) {
        throw "Feature manifest not found: $ManifestPath"
    }
    if ($null -ne $global:BeaverFeatureManifestCache) {
        return $global:BeaverFeatureManifestCache
    }
    $global:BeaverFeatureManifestCache = Get-Content $ManifestPath -Raw | ConvertFrom-Json
    return $global:BeaverFeatureManifestCache
}
