# Script:   GenerateArchitectureMap.ps1
# Purpose:  Generates ARCHITECTURE.md detailing project structure, categories,
#           and module dependencies using a Mermaid graph. Optimized for CPU and disk usage.
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "Generating architecture map..." -ForegroundColor Cyan

$projectRoot = Split-Path $PSScriptRoot -Parent
$modulesDir = Join-Path $projectRoot "Modules"
$outputPath = Join-Path $projectRoot "ARCHITECTURE.md"

# Helper function to safely read object properties under StrictMode
function Get-SafeProperty {
    param($obj, $propName)
    if ($null -ne $obj -and $null -ne $obj.PSObject.Properties[$propName]) {
        return $obj.$propName
    }
    return $null
}

$modules = @()

$buildStatePath = Join-Path $PSScriptRoot ".build_state.json"
$buildState = $null
if (Test-Path $buildStatePath) {
    try {
        $buildState = Get-Content $buildStatePath -Raw | ConvertFrom-Json
    } catch {}
}

# Find all VBA files
$vbaFiles = Get-ChildItem -Path $modulesDir -Include *.bas, *.cls, *.frm -Recurse
$thisWorkbook = Join-Path $projectRoot "ThisWorkbook.cls"
if (Test-Path $thisWorkbook) { $vbaFiles += Get-Item $thisWorkbook }

$useCache = $null -ne $buildState -and $null -ne $buildState.Metadata
if ($useCache) {
    foreach ($file in $vbaFiles) {
        $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace("\", "/")
        $metaProp = $buildState.Metadata.PSObject.Properties[$relPath]
        if ($null -eq $metaProp -or $null -eq $metaProp.Value -or $null -eq $metaProp.Value.PSObject.Properties['Category']) {
            $useCache = $false
            break
        }
    }
}

if ($useCache) {
    Write-Host "  Using cached build state metadata for module properties." -ForegroundColor Green
    foreach ($file in $vbaFiles) {
        $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace("\", "/")
        $meta = $buildState.Metadata.PSObject.Properties[$relPath].Value
        
        $compName = $meta.Name
        $category = $meta.Category
        $description = $meta.Description
        $dependencies = @()
        if ($meta.PSObject.Properties.Name -contains "Dependencies" -and $null -ne $meta.Dependencies) {
            $dependencies = @($meta.Dependencies)
        }
        $loc = $meta.Loc
        
        $modules += [pscustomobject]@{
            Name = $compName
            File = $file.Name
            RelPath = $relPath
            Category = $category
            Description = $description
            Dependencies = $dependencies
            Loc = $loc
        }
    }
} else {
    Write-Host "  Cache miss or incomplete: scanning VBA modules from disk." -ForegroundColor Yellow
    foreach ($file in $vbaFiles) {
        # CPU Optimization: Read only the first 30 lines of each file since headers are at the top
        $headerLines = [System.IO.File]::ReadLines($file.FullName) | Select-Object -First 30
        $content = $headerLines -join "`r`n"
        
        # Extract component name
        $compName = $file.BaseName
        if ($content -match '(?m)^Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
            $compName = $Matches[1]
        }
        
        # Extract category
        $category = "Unknown"
        if ($content -match "'\s*@Category:\s*([^\r\n]+)") {
            $category = $Matches[1].Trim()
        }
        
        # Extract description
        $description = ""
        if ($content -match "'\s*@Description:\s*([^\r\n]+)") {
            $description = $Matches[1].Trim()
        }
        
        # Extract dependencies
        $dependencies = @()
        if ($content -match "'\s*@Dependencies:\s*([^\r\n]+)") {
            $depString = $Matches[1].Trim()
            if ($depString -ne "None") {
                $dependencies = $depString.Split(",") | ForEach-Object { $_.Trim() }
            }
        }
        
        # Calculate LOC (Total lines of file)
        $loc = 0
        try {
            $loc = [System.IO.File]::ReadAllLines($file.FullName).Count
        } catch {
            # Fallback
        }

        # Calculate path relative to the project root
        $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace('\', '/')

        $modules += [pscustomobject]@{
            Name = $compName
            File = $file.Name
            RelPath = $relPath
            Category = $category
            Description = $description
            Dependencies = $dependencies
            Loc = $loc
        }
    }
}

# Calculate Fan-In (Used By) coupling metrics
foreach ($mod in $modules) {
    $mod | Add-Member -MemberType NoteProperty -Name "UsedBy" -Value @()
}
foreach ($mod in $modules) {
    foreach ($dep in $mod.Dependencies) {
        $target = $modules | Where-Object { $_.Name -eq $dep }
        if ($null -ne $target) {
            $target.UsedBy += $mod.Name
        }
    }
}

# Define layering precedence index
$layerPrecedence = @{
    "UI" = 4
    "Feature" = 3
    "Infrastructure" = 2
    "Library" = 1
    "Core" = 0
    "Unknown" = -1
}

# Audit layering rules violations (Ignore test modules)
$violations = @()
foreach ($mod in $modules) {
    if ($mod.Name -like "Test_*") { continue }
    
    $modLevel = $layerPrecedence[$mod.Category]
    if ($null -eq $modLevel) { $modLevel = -1 }
    
    foreach ($dep in $mod.Dependencies) {
        $depMod = $modules | Where-Object { $_.Name -eq $dep }
        if ($null -ne $depMod) {
            $depLevel = $layerPrecedence[$depMod.Category]
            if ($null -eq $depLevel) { $depLevel = -1 }
            
            # If dependency layer level is strictly higher than module layer level
            if ($depLevel -gt $modLevel) {
                $violations += [pscustomobject]@{
                    Module = $mod.Name
                    ModulePath = $mod.RelPath
                    Category = $mod.Category
                    ViolatesWith = $depMod.Name
                    ViolatesWithPath = $depMod.RelPath
                    ViolatesWithCategory = $depMod.Category
                }
            }
        }
    }
}

# Group by category for module listings
$categories = $modules | Group-Object Category | Sort-Object Name

# Load and parse features.json for user actions command mapping
$featuresJsonPath = Join-Path $projectRoot "features.json"
$uiTriggers = @()
if (Test-Path $featuresJsonPath) {
    try {
        $featuresData = Get-Content $featuresJsonPath -Raw | ConvertFrom-Json
        if ($null -ne $featuresData.Features) {
            foreach ($f in $featuresData.Features) {
                $cmdName = Get-SafeProperty $f "CommandName"
                if ([string]::IsNullOrEmpty($cmdName)) { continue }
                
                $cmdClass = "FeatCmd_" + $cmdName
                $target = $modules | Where-Object { $_.Name -eq $cmdClass }
                $classLink = if ($null -ne $target) { "**[$cmdClass]($($target.RelPath))**" } else { '`' + $cmdClass + '`' }
                $uiTriggers += [pscustomobject]@{
                    Type = "Ribbon Button"
                    ControlId = Get-SafeProperty $f "ControlId"
                    Label = Get-SafeProperty $f "Label"
                    TriggerMacro = Get-SafeProperty $f "Macro"
                    CommandClass = $classLink
                    Description = Get-SafeProperty $f "Screentip"
                }
            }
        }
        if ($null -ne $featuresData.Hotkeys) {
            foreach ($h in $featuresData.Hotkeys) {
                $cmdName = Get-SafeProperty $h "CommandName"
                if ([string]::IsNullOrEmpty($cmdName)) { continue }
                
                $cmdClass = "FeatCmd_" + $cmdName
                $target = $modules | Where-Object { $_.Name -eq $cmdClass }
                $classLink = if ($null -ne $target) { "**[$cmdClass]($($target.RelPath))**" } else { '`' + $cmdClass + '`' }
                $uiTriggers += [pscustomobject]@{
                    Type = "Hotkey"
                    ControlId = Get-SafeProperty $h "Key"
                    Label = Get-SafeProperty $h "Description"
                    TriggerMacro = Get-SafeProperty $h "Macro"
                    CommandClass = $classLink
                    Description = "Key Combination: " + (Get-SafeProperty $h "Key")
                }
            }
        }
    } catch {
        Write-Warning "Could not parse features.json for Ribbon mapping: $($_.Exception.Message)"
    }
}

# Generate Markdown content
$md = [System.Text.StringBuilder]::new()
$null = $md.AppendLine("# Beaver Add-in: System Architecture & Module Map")
$null = $md.AppendLine()
$null = $md.AppendLine("This document outlines the architectural layers and dependencies of the Beaver Excel Add-in. It is automatically generated by the build pipeline to guide human developers and AI coding agents.")
$null = $md.AppendLine()

# ----------------- ARCHITECTURAL HEALTH & VIOLATIONS SECTION -----------------
$null = $md.AppendLine("## Architectural Health & Layer Integrity")
$null = $md.AppendLine()

if ($violations.Count -gt 0) {
    $null = $md.AppendLine("> [!WARNING]")
    $null = $md.AppendLine("> **Layering Violations Detected!** Lower-level layers are importing higher-level modules. Refactor these targets to preserve architectural integrity.")
    $null = $md.AppendLine()
    $null = $md.AppendLine("| Module | Category | Depends On | Dependent Category |")
    $null = $md.AppendLine("| :--- | :--- | :--- | :--- |")
    foreach ($v in $violations) {
        $null = $md.AppendLine('| **[' + $v.Module + '](' + $v.ModulePath + ')** | `' + $v.Category + '` | **[' + $v.ViolatesWith + '](' + $v.ViolatesWithPath + ')** | `' + $v.ViolatesWithCategory + '` |')
    }
    $null = $md.AppendLine()
} else {
    $null = $md.AppendLine("> [!NOTE]")
    $null = $md.AppendLine("> **Architectural Health: 100% Compliant**. No upward layer violations detected in production code!")
    $null = $md.AppendLine()
}

# ----------------- LAYER RULE MATRIX SECTION -----------------
$null = $md.AppendLine("## Architectural Layering Rules")
$null = $md.AppendLine()
$null = $md.AppendLine("The codebase is organized into five strict horizontal layers. To prevent spaghetti code and circular dependency loops, dependencies must only flow **downward** (i.e. high-level layers can call lower-level layers, but never vice versa).")
$null = $md.AppendLine()
$null = $md.AppendLine("| Layer | Prefix / Folder | Role & Responsibility | Permitted Dependencies |")
$null = $md.AppendLine("| :--- | :--- | :--- | :--- |")
$null = $md.AppendLine("| **UI** | `UI_` / `Modules/UI` | Ribbon interface definitions, event handlers, and UserForms. | Feature, Infrastructure, Library, Core |")
$null = $md.AppendLine("| **Feature** | `FeatCmd_` / `Modules/Commands` | Individual business commands implementing `ICommand`. | Infrastructure, Library, Core |")
$null = $md.AppendLine("| **Infrastructure** | `Infra_` / `Modules/Infrastructure` | Cross-cutting systems (Undo registry, Hotkeys, Configuration, Diagnostics, Error context). | Library, Core |")
$null = $md.AppendLine("| **Library** | `Lib_` or `Udf_` / `Modules/Libraries` | Pure helper functions, utilities, and User Defined Functions (UDFs). | Core |")
$null = $md.AppendLine("| **Core** | `/Modules/Core` | Base interfaces (`ICommand`, `ICommandContext`), central contexts (`ActionContext`), and enums. | *None (Completely decoupled)* |")
$null = $md.AppendLine()
$null = $md.AppendLine("> [!IMPORTANT]")
$null = $md.AppendLine("> **Dependency Direction**: Higher-level layers (UI, Features) may depend on lower-level layers (Infrastructure, Libraries, Core), but lower-level layers must **never** depend on higher-level layers. Circular dependencies are strictly prohibited.")
$null = $md.AppendLine()

# ----------------- USER ACTION & ENTRY POINT MAPPING -----------------
if ($uiTriggers.Count -gt 0) {
    $null = $md.AppendLine("## User Action & Entry Point Mappings")
    $null = $md.AppendLine()
    $null = $md.AppendLine("This table maps user interface controls and key combinations directly to their backing command classes, helping agents trace from UI actions directly to code.")
    $null = $md.AppendLine()
    $null = $md.AppendLine("| Action Type | Control / Key ID | Label / Description | VBA Entry Point | Backing Command Class |")
    $null = $md.AppendLine("| :--- | :--- | :--- | :--- | :--- |")
    foreach ($trig in ($uiTriggers | Sort-Object Type, ControlId)) {
        $null = $md.AppendLine("| $($trig.Type) | ``$($trig.ControlId)`` | $($trig.Label) | ``$($trig.TriggerMacro)`` | $($trig.CommandClass) |")
    }
    $null = $md.AppendLine()
}

# ----------------- GRAPH SECTION -----------------
$null = $md.AppendLine("## Module Dependency Graph")
$null = $md.AppendLine()
$null = $md.AppendLine("```mermaid")
$null = $md.AppendLine("flowchart TD")
$null = $md.AppendLine("    %% Define Styles")
$null = $md.AppendLine("    classDef core fill:#d4ebf2,stroke:#0d5c75,stroke-width:1px,color:#083b4c;")
$null = $md.AppendLine("    classDef ui fill:#f5d6eb,stroke:#85145c,stroke-width:1px,color:#4a0531;")
$null = $md.AppendLine("    classDef feature fill:#d6f5d6,stroke:#148514,stroke-width:1px,color:#054a05;")
$null = $md.AppendLine("    classDef infra fill:#f5f5d6,stroke:#858514,stroke-width:1px,color:#4a4a05;")
$null = $md.AppendLine("    classDef lib fill:#f5ded6,stroke:#853714,stroke-width:1px,color:#4a1905;")
$null = $md.AppendLine()

# Output subgraphs by category
foreach ($cat in $categories) {
    $catName = $cat.Name
    $null = $md.AppendLine("    subgraph $catName [$catName Layer]")
    foreach ($mod in $cat.Group) {
        $null = $md.AppendLine('        ' + $mod.Name + '("' + $mod.Name + '")')
    }
    $null = $md.AppendLine("    end")
    $null = $md.AppendLine()
}

# Output dependency links
$null = $md.AppendLine("    %% Dependency Connections")
foreach ($mod in $modules) {
    foreach ($dep in $mod.Dependencies) {
        # Only draw link if dependency is a tracked module in our project
        $target = $modules | Where-Object { $_.Name -eq $dep }
        if ($null -ne $target) {
            $null = $md.AppendLine("    $($mod.Name) --> $dep")
        }
    }
}

$null = $md.AppendLine()
# Apply styles to classes
foreach ($cat in $categories) {
    $styleClass = switch ($cat.Name) {
        "Core" { "core" }
        "UI" { "ui" }
        "Feature" { "feature" }
        "Infrastructure" { "infra" }
        "Library" { "lib" }
        default { "core" }
    }
    foreach ($mod in $cat.Group) {
        $null = $md.AppendLine('    class ' + $mod.Name + ' ' + $styleClass + ';')
    }
}

$null = $md.AppendLine('```')
$null = $md.AppendLine()

# ----------------- MODULE DIRECTORY SECTION -----------------
$null = $md.AppendLine("## Module Directory")
$null = $md.AppendLine()

foreach ($cat in $categories) {
    $null = $md.AppendLine("### $($cat.Name) Layer")
    $null = $md.AppendLine()
    $null = $md.AppendLine("| Module | LOC | Used By (Fan-In) | Dependencies (Fan-Out) | Description |")
    $null = $md.AppendLine("| :--- | :---: | :--- | :--- | :--- |")
    foreach ($mod in ($cat.Group | Sort-Object Name)) {
        # Format outbound links
        $depLinks = @()
        foreach ($d in $mod.Dependencies) {
            $t = $modules | Where-Object { $_.Name -eq $d }
            if ($null -ne $t) {
                $depLinks += "[$d]($($t.RelPath))"
            } else {
                $depLinks += '`' + $d + '`'
            }
        }
        $depText = if ($depLinks.Count -gt 0) { $depLinks -join ", " } else { "None" }
        
        # Format inbound links
        $usedByLinks = @()
        foreach ($u in $mod.UsedBy) {
            $t = $modules | Where-Object { $_.Name -eq $u }
            if ($null -ne $t) {
                $usedByLinks += "[$u]($($t.RelPath))"
            } else {
                $usedByLinks += '`' + $u + '`'
            }
        }
        $usedByText = if ($usedByLinks.Count -gt 0) { $usedByLinks -join ", " } else { "None" }

        # Output line
        $null = $md.AppendLine("| **[$($mod.Name)]($($mod.RelPath))** | $($mod.Loc) | $usedByText | $depText | $($mod.Description) |")
    }
    $null = $md.AppendLine()
}

# ----------------- DESIGN PATTERNS SECTION -----------------
$null = $md.AppendLine("## Core Design Patterns")
$null = $md.AppendLine()
$null = $md.AppendLine('### 1. The Command Pattern (`ICommand`)')
$null = $md.AppendLine('All workbook features are implemented as class modules conforming to the `ICommand` interface:')
$null = $md.AppendLine('1. **Trigger**: An entry point in `UI_Ribbon` or `UI_Hotkeys` captures user interaction.')
$null = $md.AppendLine('2. **Context**: An `ActionContext` (implementing `ICommandContext`) is initialized to store target ranges, parameters, and application state.')
$null = $md.AppendLine('3. **Execution**: The command is dispatched to `CommandInvoker.Execute`, which validates the context, logs telemetry, tracks history for Undo, and calls the command''s custom `ICommand_Execute` method.')
$null = $md.AppendLine()
$null = $md.AppendLine('### 2. State & Safety Management (`Infra_AppState`)')
$null = $md.AppendLine('To ensure fast execution and a clean user experience, Excel screen updating, events, and automatic calculation must be disabled during heavy operations.')
$null = $md.AppendLine('* Always wrap operations in an `Infra_AppStateGuard` context or manually handle it using:')
$null = $md.AppendLine('  ```vba')
$null = $md.AppendLine('  Dim appState As Infra_AppState: Set appState = Infra_AppState.DisableSpeedSettings()')
$null = $md.AppendLine('  On Error GoTo ErrHandler')
$null = $md.AppendLine("  ' ... work ...")
$null = $md.AppendLine('  appState.Restore')
$null = $md.AppendLine('  ```')
$null = $md.AppendLine()
$null = $md.AppendLine('## AI Agent Development Conventions')
$null = $md.AppendLine('AI coding agents editing this repository must adhere to the following rules:')
$null = $md.AppendLine('* **Option Explicit**: Every source file must contain `Option Explicit` as its first line.')
$null = $md.AppendLine('* **Metadata Headers**: Every file must start with `@Module`, `@Category`, `@Description`, and `@Dependencies` comments.')
$null = $md.AppendLine('* **Context & Errors**: Every public procedure must use `Infra_Error.Track` and standard `ErrHandler` error traps.')
$null = $md.AppendLine('* **Excel References**: Always qualify ranges with sheet variables (e.g. `ws.Range` instead of `Range`). Never call conversion functions directly on range properties without checking for `Null` first.')
$null = $md.AppendLine()

# ----------------- IDEMPOTENCY CHECK & WRITE -----------------
$newContent = $md.ToString()

if (Test-Path $outputPath) {
    # Read the existing content
    $existingContent = [System.IO.File]::ReadAllText($outputPath)
    
    # Normalize line endings to avoid line-ending noise (CRLF vs LF)
    $normalizedExisting = $existingContent.Replace("`r`n", "`n")
    $normalizedNew = $newContent.Replace("`r`n", "`n")
    
    if ($normalizedExisting -eq $normalizedNew) {
        Write-Host "  Architecture map is already up to date. No write needed." -ForegroundColor Green
        return
    }
}

# Write contents
[System.IO.File]::WriteAllText($outputPath, $newContent, [System.Text.Encoding]::UTF8)
Write-Host "  Architecture map successfully written to: $outputPath" -ForegroundColor Green
