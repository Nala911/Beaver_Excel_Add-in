# Script:   HashingState.ps1
# Purpose:  MD5 file hashing, manifest structural hashing, state tracking, and change detection for Beaver build pipeline.
# ==============================================================================

Set-StrictMode -Version Latest

# Compile high-performance C# helper for VBA structural hashing
$vbaHelperCode = @"
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

public static class VbaHashHelper {
    private static readonly Regex SpaceCollapseRegex = new Regex(@"\s+", RegexOptions.Compiled);
    private static readonly Regex AttributeRegex = new Regex(@"^Attribute\s+VB_Name\s*=", RegexOptions.Compiled);
    private static readonly Regex AttributeGeneralRegex = new Regex(@"^Attribute\s+", RegexOptions.Compiled);
    private static readonly Regex MetadataCommentRegex = new Regex(@"^'\s*@", RegexOptions.Compiled);

    public static string GetStructuralCode(string filePath) {
        if (!System.IO.File.Exists(filePath)) return string.Empty;
        
        string[] lines = System.IO.File.ReadAllLines(filePath);
        List<string> structuralLines = new List<string>(lines.Length);
        
        string extension = System.IO.Path.GetExtension(filePath).ToLower();
        bool inCodeSection = true;
        if (extension == ".frm" || extension == ".cls") {
            inCodeSection = false;
        }
        
        foreach (string line in lines) {
            if (!inCodeSection) {
                if (AttributeRegex.IsMatch(line)) {
                    inCodeSection = true;
                }
                if (AttributeGeneralRegex.IsMatch(line)) {
                    structuralLines.Add(line.Trim());
                }
                continue;
            }
            
            string trimmed = line.Trim();
            if (trimmed.Length == 0) continue;
            
            // Skip pure comment lines, but keep metadata comments starting with '@'
            if (trimmed.StartsWith("'")) {
                if (!MetadataCommentRegex.IsMatch(trimmed)) {
                    continue;
                }
            }
            
            // Strip trailing comments (basic logic: find "'" outside quotes)
            string cleanLine = trimmed;
            bool inQuote = false;
            int commentIdx = -1;
            for (int i = 0; i < trimmed.Length; i++) {
                char c = trimmed[i];
                if (c == '"') {
                    inQuote = !inQuote;
                }
                else if (c == '\'' && !inQuote) {
                    commentIdx = i;
                    break;
                }
            }
            
            if (commentIdx != -1) {
                string comment = trimmed.Substring(commentIdx);
                if (MetadataCommentRegex.IsMatch(comment)) {
                    cleanLine = trimmed;
                } else {
                    cleanLine = trimmed.Substring(0, commentIdx).Trim();
                }
            }
            
            if (cleanLine.Length == 0) continue;
            
            // Collapse multiple spaces
            string collapsed = SpaceCollapseRegex.Replace(cleanLine, " ");
            structuralLines.Add(collapsed);
        }
        
        return string.Join("\n", structuralLines);
    }
}
"@

if (-not ([System.Management.Automation.PSTypeName]"VbaHashHelper").Type) {
    Add-Type -TypeDefinition $vbaHelperCode
}

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
    
    $oldState = Get-BuildState
    
    foreach ($relPath in $FileHashes.Keys) {
        $currentHash = $FileHashes[$relPath]
        
        # Check if the file was in the old build state and has the same hash and contains new metadata fields
        $cachedMeta = $null
        if ($null -ne $oldState -and $null -ne $oldState.Files -and $null -ne $oldState.Metadata) {
            $oldHashProp = $oldState.Files.PSObject.Properties[$relPath]
            $oldMetaProp = $oldState.Metadata.PSObject.Properties[$relPath]
            if ($null -ne $oldHashProp -and $oldHashProp.Value -eq $currentHash -and $null -ne $oldMetaProp) {
                $oldMeta = $oldMetaProp.Value
                if ($oldMeta.PSObject.Properties.Name -contains "Category" -and $oldMeta.PSObject.Properties.Name -contains "Loc") {
                    $cachedMeta = $oldMeta
                }
            }
        }
        
        if ($null -ne $cachedMeta) {
            # Lint status from global cache if updated this session, otherwise preserve old
            if ($null -ne $global:BeaverLintStatusCache -and $global:BeaverLintStatusCache.ContainsKey($relPath)) {
                if ($cachedMeta.PSObject.Properties.Name -contains "LintPassed") {
                    $cachedMeta.LintPassed = $global:BeaverLintStatusCache[$relPath]
                } else {
                    $cachedMeta | Add-Member -NotePropertyName "LintPassed" -NotePropertyValue $global:BeaverLintStatusCache[$relPath] -Force
                }
            }
            # Test manifest from global cache if updated this session
            if ($null -ne $global:BeaverTestManifestCache -and $global:BeaverTestManifestCache.ContainsKey($relPath)) {
                if ($cachedMeta.PSObject.Properties.Name -contains "Tests") {
                    $cachedMeta.Tests = $global:BeaverTestManifestCache[$relPath]
                } else {
                    $cachedMeta | Add-Member -NotePropertyName "Tests" -NotePropertyValue $global:BeaverTestManifestCache[$relPath] -Force
                }
            }
            $metadata[$relPath] = $cachedMeta
            continue
        }
        
        # Cache miss: read file and compute metadata
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
                if ($null -ne $oldState -and $null -ne $oldState.Metadata -and $null -ne $oldState.Metadata.PSObject.Properties[$relPath]) {
                    $oldMeta = $oldState.Metadata.PSObject.Properties[$relPath].Value
                    if ($oldMeta.PSObject.Properties.Name -contains "Tests" -and $null -ne $oldMeta.Tests) {
                        $meta["Tests"] = @($oldMeta.Tests)
                    }
                }
            }
            
            # Parse standard module dependencies
            $deps = @()
            if ($relPath -match "\.(bas|cls|frm)$") {
                $deps = Get-ModuleDependencies -FilePath $absPath
            }
            $meta["Dependencies"] = $deps
            
            # Parse test dependencies
            $testDeps = @{}
            if (($null -ne $meta["Tests"]) -and ($meta["Tests"].Count -gt 0)) {
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
            
            # Extract basic module info for Architecture Map and status caching
            $category = "Unknown"
            $description = ""
            $compName = [System.IO.Path]::GetFileNameWithoutExtension($relPath)
            $loc = 0
            
            if ($relPath -match "\.(bas|cls|frm)$" -or $relPath -eq "ThisWorkbook.cls") {
                try {
                    $linesForHeader = [System.IO.File]::ReadLines($absPath) | Select-Object -First 30
                    $headerContent = $linesForHeader -join "`r`n"
                    if ($headerContent -match '(?m)^Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
                        $compName = $Matches[1]
                    }
                    if ($headerContent -match "'\s*@Category:\s*([^\r\n]+)") {
                        $category = $Matches[1].Trim()
                    }
                    if ($headerContent -match "'\s*@Description:\s*([^\r\n]+)") {
                        $description = $Matches[1].Trim()
                    }
                    $loc = [System.IO.File]::ReadAllLines($absPath).Count
                } catch {}
            }
            
            $meta["Category"] = $category
            $meta["Description"] = $description
            $meta["Name"] = $compName
            $meta["Loc"] = $loc
            
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
    Write-FileSafe -Path $resolvedBuildStatePath -Content $stateJson -Encoding ([System.Text.Encoding]::UTF8)
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
        Write-FileSafe -Path $resolvedBuildStatePath -Content $stateJson -Encoding ([System.Text.Encoding]::UTF8)
        $global:BeaverBuildStateCache = $buildState
    }
}

function Get-VbaStructuralHash {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) { return "" }
    
    try {
        $joined = [VbaHashHelper]::GetStructuralCode($FilePath)
        
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

    $diskThisWorkbookClsVar = Get-Variable -Name "diskThisWorkbookCls" -ErrorAction SilentlyContinue
    $resolvedThisWorkbook = if ($null -ne $diskThisWorkbookClsVar) { $diskThisWorkbookClsVar.Value } else { Join-Path $resolvedProjectRoot "ThisWorkbook.cls" }

    $modulesDirVar = Get-Variable -Name "modulesDir" -ErrorAction SilentlyContinue
    $resolvedModulesDir = if ($null -ne $modulesDirVar) { $modulesDirVar.Value } else { Join-Path $resolvedProjectRoot "Modules" }

    $hashes = @{}
    $buildState = Get-BuildState
    
    $resolveHash = {
        param($file, $relPath)
        if ($null -eq $file) { return "" }
        
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
        return Get-FileHashOptimized -FilePath $file.FullName
    }
    
    # Manifest
    if (Test-Path $resolvedFeatureManifestPath) {
        $hashes["features.json"] = & $resolveHash (Get-Item $resolvedFeatureManifestPath) "features.json"
    }
    
    # ThisWorkbook
    if (Test-Path $resolvedThisWorkbook) {
        $hashes["ThisWorkbook.cls"] = & $resolveHash (Get-Item $resolvedThisWorkbook) "ThisWorkbook.cls"
    }
    
    # Modules
    if (Test-Path $resolvedModulesDir) {
        $vbaFiles = Get-ChildItem -Path $resolvedModulesDir -Include *.bas, *.cls, *.frm -Recurse
        foreach ($file in $vbaFiles) {
            $relPath = $file.FullName.Substring($resolvedProjectRoot.Length + 1).Replace("\", "/")
            $hashes[$relPath] = & $resolveHash $file $relPath
            
            # If it's a form, also include the companion FRX file hash if it exists
            if ($file.Extension -eq ".frm") {
                $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
                if (Test-Path $frxPath) {
                    $frxRelPath = $frxPath.Substring($resolvedProjectRoot.Length + 1).Replace("\", "/")
                    $frxFile = Get-Item $frxPath
                    $hashes[$frxRelPath] = & $resolveHash $frxFile $frxRelPath
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

function Write-FileSafe {
    param(
        [string]$Path,
        [string]$Content,
        [System.Text.Encoding]$Encoding = [System.Text.Encoding]::UTF8
    )
    $dir = Split-Path $Path
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $tmpPath = "$Path.tmp"
    $written = $false
    for ($i = 0; $i -lt 5; $i++) {
        try {
            [System.IO.File]::WriteAllText($tmpPath, $Content, $Encoding)
            if (Test-Path $Path) {
                [System.IO.File]::Replace($tmpPath, $Path, $null, $true)
            } else {
                [System.IO.File]::Move($tmpPath, $Path)
            }
            $written = $true
            break
        } catch {
            Start-Sleep -Milliseconds (100 * ($i + 1))
        }
    }
    if (-not $written) {
        [System.IO.File]::WriteAllText($Path, $Content, $Encoding)
    }
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
    Write-FileSafe -Path $Path -Content $Content -Encoding ([System.Text.Encoding]::ASCII)
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

function Get-ProjectChanges {
    param(
        [switch]$Force
    )

    $projectRootVar = Get-Variable -Name "projectRoot" -ErrorAction SilentlyContinue
    $resolvedProjectRoot = if ($null -ne $projectRootVar) { $projectRootVar.Value } else { Split-Path (Split-Path $PSScriptRoot -Parent) -Parent }

    $featureManifestPathVar = Get-Variable -Name "featureManifestPath" -ErrorAction SilentlyContinue
    $resolvedFeatureManifestPath = if ($null -ne $featureManifestPathVar) { $featureManifestPathVar.Value } else { Join-Path $resolvedProjectRoot "features.json" }

    $currentHashes = Get-SourceFileHashes -Force:$Force
    $buildState = Get-BuildState
    
    $manifestChanged = $true
    $manifestStructureChanged = $true
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
        
        $manifestStructureChanged = $false
        if ($manifestChanged) {
            $newStructuralHash = Get-ManifestStructuralHash -Path $resolvedFeatureManifestPath
            $oldStructuralHash = $null
            if ($buildState.PSObject.Properties.Name.Contains("ManifestStructuralHash")) {
                $oldStructuralHash = $buildState.ManifestStructuralHash
            }
            if ($newStructuralHash -ne $oldStructuralHash) {
                $manifestStructureChanged = $true
            }
        }
    } else {
        $changedFiles = @($currentHashes.Keys)
    }
    
    $hasAnyChanges = ($changedFiles.Count -gt 0 -or $deletedFiles.Count -gt 0)
    
    return [pscustomobject]@{
        ChangedFiles = $changedFiles
        DeletedFiles = $deletedFiles
        ManifestChanged = $manifestChanged
        ManifestStructureChanged = $manifestStructureChanged
        HasAnyChanges = $hasAnyChanges
    }
}

