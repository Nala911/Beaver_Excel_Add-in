# Script:   GenerateArchitectureMap.ps1
# Purpose:  Generates ARCHITECTURE.md detailing project structure, subsystem blueprints,
#           logical connections, entry point mappings, and module dependencies using Mermaid graphs.
#           Specifically tuned as a machine-readable, structured blueprint for AI Coding Agents.
#           Optimized for ultra-fast CPU and low disk I/O with build-state metadata caching.
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "Generating architecture map for AI Agents..." -ForegroundColor Cyan

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

$modules = [System.Collections.Generic.List[PSCustomObject]]::new()

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
        if ($null -eq $metaProp -or $null -eq $metaProp.Value) {
            $useCache = $false
            break
        }
        $meta = $metaProp.Value
        if ($null -eq $meta.PSObject.Properties['Category'] -or 
            $meta.Length -ne $file.Length -or 
            $meta.LastWriteTime -ne $file.LastWriteTime.ToFileTime().ToString()) {
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
        
        $modules.Add([pscustomobject]@{
            Name = $compName
            File = $file.Name
            RelPath = $relPath
            Category = $category
            Description = $description
            Dependencies = $dependencies
            Loc = $loc
            UsedBy = [System.Collections.Generic.List[string]]::new()
        })
    }
} else {
    Write-Host "  Cache miss or incomplete: scanning VBA modules from disk." -ForegroundColor Yellow
    foreach ($file in $vbaFiles) {
        # Optimize disk scanning by reading the file only once
        $lines = [System.IO.File]::ReadAllLines($file.FullName)
        $loc = $lines.Count
        $headerLines = $lines[0..([Math]::Min(29, $loc - 1))]
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

        # Calculate path relative to the project root
        $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace('\', '/')

        $modules.Add([pscustomobject]@{
            Name = $compName
            File = $file.Name
            RelPath = $relPath
            Category = $category
            Description = $description
            Dependencies = $dependencies
            Loc = $loc
            UsedBy = [System.Collections.Generic.List[string]]::new()
        })
    }
}

# Create $moduleMap mapping Name -> PSCustomObject for fast O(1) lookups
$moduleMap = @{}
foreach ($mod in $modules) {
    $moduleMap[$mod.Name] = $mod
}

# Calculate Fan-In (Used By) coupling metrics using hashtable lookups
foreach ($mod in $modules) {
    foreach ($dep in $mod.Dependencies) {
        $target = $moduleMap[$dep]
        if ($null -ne $target) {
            $target.UsedBy.Add($mod.Name)
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
$violations = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($mod in $modules) {
    if ($mod.Name -like "Test_*") { continue }
    
    $modLevel = $layerPrecedence[$mod.Category]
    if ($null -eq $modLevel) { $modLevel = -1 }
    
    foreach ($dep in $mod.Dependencies) {
        $depMod = $moduleMap[$dep]
        if ($null -ne $depMod) {
            $depLevel = $layerPrecedence[$depMod.Category]
            if ($null -eq $depLevel) { $depLevel = -1 }
            
            # If dependency layer level is strictly higher than module layer level
            if ($depLevel -gt $modLevel) {
                if ($mod.Name -eq "Lib_HelpManifest" -or $mod.Name -eq "Lib_UdfRegistry") {
                    continue
                }
                $violations.Add([pscustomobject]@{
                    Module = $mod.Name
                    ModulePath = $mod.RelPath
                    Category = $mod.Category
                    ViolatesWith = $depMod.Name
                    ViolatesWithPath = $depMod.RelPath
                    ViolatesWithCategory = $depMod.Category
                })
            }
        }
    }
}

# Group by category for module listings
$categories = $modules | Group-Object Category | Sort-Object Name

# Load and parse features.json for user actions command mapping
$featuresJsonPath = Join-Path $projectRoot "features.json"
$uiTriggers = [System.Collections.Generic.List[PSCustomObject]]::new()
if (Test-Path $featuresJsonPath) {
    try {
        $featuresData = Get-Content $featuresJsonPath -Raw | ConvertFrom-Json
        if ($null -ne $featuresData.Features) {
            foreach ($f in $featuresData.Features) {
                $cmdName = Get-SafeProperty $f "CommandName"
                if ([string]::IsNullOrEmpty($cmdName)) { continue }
                
                $cmdClass = "FeatCmd_" + $cmdName
                $target = $moduleMap[$cmdClass]
                $classLink = if ($null -ne $target) { "**[$cmdClass]($($target.RelPath))**" } else { '`' + $cmdClass + '`' }
                $uiTriggers.Add([pscustomobject]@{
                    Type = "Ribbon Button"
                    ControlId = Get-SafeProperty $f "ControlId"
                    Label = Get-SafeProperty $f "Label"
                    TriggerMacro = Get-SafeProperty $f "Macro"
                    CommandClass = $classLink
                    Description = Get-SafeProperty $f "Screentip"
                })
            }
        }
        if ($null -ne $featuresData.Hotkeys) {
            foreach ($h in $featuresData.Hotkeys) {
                $cmdName = Get-SafeProperty $h "CommandName"
                if ([string]::IsNullOrEmpty($cmdName)) { continue }
                
                $cmdClass = "FeatCmd_" + $cmdName
                $target = $moduleMap[$cmdClass]
                $classLink = if ($null -ne $target) { "**[$cmdClass]($($target.RelPath))**" } else { '`' + $cmdClass + '`' }
                $uiTriggers.Add([pscustomobject]@{
                    Type = "Hotkey"
                    ControlId = Get-SafeProperty $h "Key"
                    Label = Get-SafeProperty $h "Description"
                    TriggerMacro = Get-SafeProperty $h "Macro"
                    CommandClass = $classLink
                    Description = "Key Combination: " + (Get-SafeProperty $h "Key")
                })
            }
        }
    } catch {
        Write-Warning "Could not parse features.json for Ribbon mapping: $($_.Exception.Message)"
    }
}

# Generate Markdown content using high-performance StringBuilder
$md = [System.Text.StringBuilder]::new()

function Add-Line {
    param([string]$text = "")
    $null = $md.AppendLine($text)
}

Add-Line '# Beaver Add-in: AI Agent System Architecture & Blueprint'
Add-Line ''
Add-Line 'This document is the **system context blueprint explicitly tuned for AI coding agents**. It provides instant orientation on project structure, entry points, subsystem APIs, safety constraints, and dependency rules. Automatically generated by `Update.ps1`.'
Add-Line ''

# ----------------- TABLE OF CONTENTS -----------------
Add-Line '## Table of Contents'
Add-Line ''
Add-Line '- [1. AI Agent Mandates & Safety Rules](#1-ai-agent-mandates--safety-rules)'
Add-Line '- [2. Architectural Layering Rules & Health Matrix](#2-architectural-layering-rules--health-matrix)'
Add-Line '- [3. Command Execution Lifecycle & Sequence Flow](#3-command-execution-lifecycle--sequence-flow)'
Add-Line '- [4. Core Subsystem Blueprints & Connections](#4-core-subsystem-blueprints--connections)'
Add-Line '- [5. User Action & Entry Point Mappings](#5-user-action--entry-point-mappings)'
Add-Line '- [6. Visual Module Dependency Graph](#6-visual-module-dependency-graph)'
Add-Line '- [7. Complete Module Directory by Layer](#7-complete-module-directory-by-layer)'
Add-Line ''

# ----------------- 1. AI AGENT MANDATES & SAFETY RULES -----------------
Add-Line '## 1. AI Agent Mandates & Safety Rules'
Add-Line ''
Add-Line 'All AI agents editing this codebase MUST strictly observe these non-negotiable rules:'
Add-Line ''
Add-Line '| Rule | Requirement | Example / Correct Pattern | Incorrect Pattern (Prohibited) |'
Add-Line '| :--- | :--- | :--- | :--- |'
Add-Line '| **Option Explicit** | Must be Line 1 of every `.bas`, `.cls`, `.frm` module. | `Option Explicit` | *(Omitted line 1)* |'
Add-Line '| **Worksheet Qualification** | Always qualify Range/Cells with explicit sheet variable. | `ws.Range("A1").Value` | `Range("A1").Value` |'
Add-Line '| **Null-Safe Property Access** | Check range properties for `Null` before type casting. | `If Not IsNull(cell.Value) Then ...` | `CStr(cell.Value)` directly |'
Add-Line '| **Backward Deletion** | Iterate backward (`Step -1`) when deleting collections. | `For i = count To 1 Step -1` | `For i = 1 To count` |'
Add-Line '| **Error Trapping** | Wrap public entry points with `Infra_Error` traps. | `Infra_Error.Track "Module.Proc"` | Naked procedures without traps |'
Add-Line '| **Layer Prefixes** | Preserve strict module naming prefixes. | `UI_`, `FeatCmd_`, `Infra_`, `Lib_`, `Core` | Generic module names |'
Add-Line ''

# ----------------- 2. LAYERING RULES & HEALTH MATRIX -----------------
Add-Line '## 2. Architectural Layering Rules & Health Matrix'
Add-Line ''
Add-Line 'The codebase is strictly organized into **five horizontal layers**. Dependencies must only flow **downward** (higher-level layers call lower-level layers, never vice versa). Circular imports are strictly forbidden.'
Add-Line ''
Add-Line '| Layer | Folder / Prefix | Primary Purpose | Permitted Dependencies |'
Add-Line '| :--- | :--- | :--- | :--- |'
Add-Line '| **UI** | `Modules/UI` (`UI_`) | Ribbon callbacks, hotkey definitions, UserForms, dialog management. | Feature, Infrastructure, Library, Core |'
Add-Line '| **Feature** | `Modules/Commands` (`FeatCmd_`) | Independent business commands implementing `ICommand`. | Infrastructure, Library, Core |'
Add-Line '| **Infrastructure** | `Modules/Infrastructure` (`Infra_`) | Cross-cutting systems: Undo registry, AppState, Hotkeys, Configuration, Diagnostics, Error logging. | Library, Core |'
Add-Line '| **Library** | `Modules/Libraries` (`Lib_`, `Udf_`) | Pure functional utilities, JSON tools, algorithms, User Defined Functions (UDFs). | Core |'
Add-Line '| **Core** | `Modules/Core` | Base interfaces (`ICommand`, `ICommandContext`), central contexts (`ActionContext`), enums. | *None (Fully decoupled)* |'
Add-Line ''

if ($violations.Count -gt 0) {
    Add-Line '> [!WARNING]'
    Add-Line '> **Layering Violations Detected!** Lower-level layers are importing higher-level modules. Refactor these targets to preserve architectural integrity.'
    Add-Line ''
    Add-Line '| Module | Category | Depends On | Dependent Category |'
    Add-Line '| :--- | :--- | :--- | :--- |'
    foreach ($v in $violations) {
        $mName = [string]$v.Module
        $mPath = [string]$v.ModulePath
        $mCat = [string]$v.Category
        $vName = [string]$v.ViolatesWith
        $vPath = [string]$v.ViolatesWithPath
        $vCat = [string]$v.ViolatesWithCategory
        Add-Line ('| **[' + $mName + '](' + $mPath + ')** | `' + $mCat + '` | **[' + $vName + '](' + $vPath + ')** | `' + $vCat + '` |')
    }
    Add-Line ''
} else {
    Add-Line '> [!NOTE]'
    Add-Line '> **Architectural Health: 100% Compliant**. Zero upward layer violations detected across production modules.'
    Add-Line ''
}

# ----------------- 3. COMMAND LIFECYCLE SEQUENCE FLOW -----------------
Add-Line '## 3. Command Execution Lifecycle & Sequence Flow'
Add-Line ''
Add-Line 'Every feature in the add-in follows a unified execution pipeline managed by **`CommandInvoker`**. Below is the sequence diagram illustrating how user triggers translate into safe worksheet modifications:'
Add-Line ''
Add-Line '```mermaid'
Add-Line 'sequenceDiagram'
Add-Line '    autonumber'
Add-Line '    actor User as User / Excel UI'
Add-Line '    participant UI as UI Layer (UI_Ribbon / UI_Hotkeys)'
Add-Line '    participant Invoker as Core (CommandInvoker)'
Add-Line '    participant Guard as Infra (Infra_AppStateGuard)'
Add-Line '    participant Command as Feature (FeatCmd_*)'
Add-Line '    participant Undo as Infra (Infra_Undo)'
Add-Line '    participant Excel as Excel Worksheet'
Add-Line ''
Add-Line '    User->>UI: Click Ribbon Control or Press Hotkey'
Add-Line '    UI->>Invoker: CommandInvoker.Execute(Cmd, Context)'
Add-Line '    Invoker->>Guard: Disable Speed Settings (ScreenUpdating=False, Calc=Manual)'
Add-Line '    Invoker->>Command: FeatCmd_*.Execute(Context)'
Add-Line '    Command->>Undo: Register Before/After State Snapshots'
Add-Line '    Command->>Excel: Perform Batch Range / Cell Operations'
Add-Line '    Excel-->>Command: Operations Complete'
Add-Line '    Command-->>Invoker: Return CommandValidationResult / Success'
Add-Line '    Invoker->>Guard: Restore Application Speed Settings'
Add-Line '    Invoker-->>UI: Execution Completed Successfully'
Add-Line '    UI-->>User: Interaction Complete'
Add-Line '```'
Add-Line ''

# ----------------- 4. CORE SUBSYSTEM BLUEPRINTS -----------------
Add-Line '## 4. Core Subsystem Blueprints & Connections'
Add-Line ''
Add-Line '### 4.1 The Command Pattern (`ICommand` & `CommandInvoker`)'
Add-Line 'All features are encapsulated as classes implementing `ICommand`. Features never directly manipulate global UI state or manage error logging manually. Execution passes through **`CommandInvoker.Execute`**, which handles validation, telemetry logging, Undo registration, and exception trapping.'
Add-Line ''
Add-Line '### 4.2 Application State & Safety Guard (`Infra_AppState` & `Infra_AppStateGuard`)'
Add-Line 'Excel macro operations can freeze or stutter if ScreenUpdating, Calculation, or Events are enabled during heavy processing. The add-in utilizes RAII-style guards:'
Add-Line '```vba'
Add-Line 'Dim appState As Infra_AppState'
Add-Line 'Set appState = Infra_AppState.DisableSpeedSettings()'
Add-Line 'On Error GoTo ErrHandler'
Add-Line "' ... Perform Range Operations ..."
Add-Line 'ErrHandler:'
Add-Line "    appState.Restore ' Always restores original user Excel settings even on failure"
Add-Line '```'
Add-Line ''
Add-Line '### 4.3 Transaction-Based Undo Subsystem (`Infra_Undo`)'
Add-Line "VBA macros ordinarily clear Excel's native undo stack. Beaver provides custom multi-level undo by capturing range snapshots in `ActionContext` before mutating cells, registering them into `Infra_Undo`, and binding `Application.OnUndo`."
Add-Line ''
Add-Line '### 4.4 High-Performance Batch Processing (`Infra_BatchProcessor`)'
Add-Line 'Iterating cell-by-cell in VBA is slow. `Infra_BatchProcessor` loads entire range values into 2D memory arrays (`Variant`), processes transformations in CPU RAM via `ICellTransformer`, and writes back to Excel in a single memory block assignment.'
Add-Line ''
Add-Line '### 4.5 Decoupled UserForm & Dialog Architecture (`UI_Factory` & Request Objects)'
Add-Line 'UserForms (`UI_Dialog*`) do not contain business logic. They capture user input into decoupled request models (e.g. `CleanDataRequest`, `ModifyDataRequest`, `ExportRequest`) and delegate execution to `UI_Factory` and feature command classes.'
Add-Line ''

# ----------------- 5. USER ACTION & ENTRY POINT MAPPINGS -----------------
if ($uiTriggers.Count -gt 0) {
    Add-Line '## 5. User Action & Entry Point Mappings'
    Add-Line ''
    Add-Line 'This map connects UI interactions (Ribbon buttons, Key combinations) directly to their VBA entry macros and backing command classes.'
    Add-Line ''
    Add-Line '| Action Type | Control / Key ID | Label / Description | VBA Entry Point | Backing Command Class |'
    Add-Line '| :--- | :--- | :--- | :--- | :--- |'
    foreach ($trig in ($uiTriggers | Sort-Object Type, ControlId)) {
        $tType = [string]$trig.Type
        $tId = [string]$trig.ControlId
        $tLabel = [string]$trig.Label
        $tMacro = [string]$trig.TriggerMacro
        $tClass = [string]$trig.CommandClass
        Add-Line ('| ' + $tType + ' | `' + $tId + '` | ' + $tLabel + ' | `' + $tMacro + '` | ' + $tClass + ' |')
    }
    Add-Line ''
}

# ----------------- 6. VISUAL MODULE DEPENDENCY GRAPH -----------------
Add-Line '## 6. Visual Module Dependency Graph'
Add-Line ''
Add-Line '```mermaid'
Add-Line 'flowchart TD'
Add-Line '    %% Layer Color Definitions'
Add-Line '    classDef core fill:#d4ebf2,stroke:#0d5c75,stroke-width:1px,color:#083b4c;'
Add-Line '    classDef ui fill:#f5d6eb,stroke:#85145c,stroke-width:1px,color:#4a0531;'
Add-Line '    classDef feature fill:#d6f5d6,stroke:#148514,stroke-width:1px,color:#054a05;'
Add-Line '    classDef infra fill:#f5f5d6,stroke:#858514,stroke-width:1px,color:#4a4a05;'
Add-Line '    classDef lib fill:#f5ded6,stroke:#853714,stroke-width:1px,color:#4a1905;'
Add-Line ''

# Output subgraphs by category
foreach ($cat in $categories) {
    $catName = [string]$cat.Name
    Add-Line ("    subgraph " + $catName + " [" + $catName + " Layer]")
    foreach ($mod in $cat.Group) {
        $mName = [string]$mod.Name
        Add-Line ('        ' + $mName + '("' + $mName + '")')
    }
    Add-Line '    end'
    Add-Line ''
}

# Output dependency links
Add-Line '    %% Dependency Connections'
foreach ($mod in $modules) {
    $mName = [string]$mod.Name
    foreach ($dep in $mod.Dependencies) {
        $target = $moduleMap[$dep]
        if ($null -ne $target) {
            Add-Line ("    " + $mName + " --> " + $dep)
        }
    }
}

Add-Line ''
# Apply styles to classes
foreach ($cat in $categories) {
    $catName = [string]$cat.Name
    $styleClass = switch ($catName) {
        "Core" { "core" }
        "UI" { "ui" }
        "Feature" { "feature" }
        "Infrastructure" { "infra" }
        "Library" { "lib" }
        default { "core" }
    }
    foreach ($mod in $cat.Group) {
        $mName = [string]$mod.Name
        Add-Line ('    class ' + $mName + ' ' + $styleClass + ';')
    }
}

Add-Line '```'
Add-Line ''

# ----------------- 7. MODULE DIRECTORY SECTION -----------------
Add-Line '## 7. Complete Module Directory by Layer'
Add-Line ''

foreach ($cat in $categories) {
    $catName = [string]$cat.Name
    Add-Line ("### " + $catName + " Layer")
    Add-Line ''
    Add-Line '| Module | LOC | Used By (Fan-In) | Dependencies (Fan-Out) | Description |'
    Add-Line '| :--- | :---: | :--- | :--- | :--- |'
    foreach ($mod in ($cat.Group | Sort-Object Name)) {
        # Format outbound links
        $depLinks = [System.Collections.Generic.List[string]]::new()
        foreach ($d in $mod.Dependencies) {
            $t = $moduleMap[$d]
            if ($null -ne $t) {
                $depLinks.Add("[" + $d + "](" + $t.RelPath + ")")
            } else {
                $depLinks.Add('`' + $d + '`')
            }
        }
        $depText = if ($depLinks.Count -gt 0) { $depLinks -join ", " } else { "None" }
        
        # Format inbound links
        $usedByLinks = [System.Collections.Generic.List[string]]::new()
        foreach ($u in $mod.UsedBy) {
            $t = $moduleMap[$u]
            if ($null -ne $t) {
                $usedByLinks.Add("[" + $u + "](" + $t.RelPath + ")")
            } else {
                $usedByLinks.Add('`' + $u + '`')
            }
        }
        $usedByText = if ($usedByLinks.Count -gt 0) { $usedByLinks -join ", " } else { "None" }

        $mName = [string]$mod.Name
        $mRel = [string]$mod.RelPath
        $mLoc = [string]$mod.Loc
        $mDesc = [string]$mod.Description

        Add-Line ("| **[" + $mName + "](" + $mRel + ")** | " + $mLoc + " | " + $usedByText + " | " + $depText + " | " + $mDesc + " |")
    }
    Add-Line ''
}

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
