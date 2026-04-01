# ==============================================================================
# Script:   Update.ps1
# Purpose:  Syncs VBA modules from disk into Beaver Add-in.xlsm via Excel COM.
#           Removes all existing managed components, re-imports from Modules\,
#           and replaces ThisWorkbook code from the project root's ThisWorkbook.cls.
# Usage:    .\Update.ps1  (run from project root with Excel closed)
# Prereq:   "Trust access to VBA project object model" must be enabled in Excel
#           (File > Options > Trust Center > Macro Settings)
# ==============================================================================

[CmdletBinding()]
param(
    [switch]$SkipRuntimeTests,
    [switch]$IncludeDevFeatures,
    [switch]$BumpVersion,
    [switch]$RefreshAgentsGuide
)

$excelPath = Join-Path $PSScriptRoot "Beaver Add-in.xlsm"
$modulesDir = Join-Path $PSScriptRoot "Modules"
$desktopThisWorkbookCls = Join-Path $PSScriptRoot "ThisWorkbook.cls"
$ribbonXmlPath = Join-Path $PSScriptRoot "ribbon.xml"
$agentsMdPath = Join-Path $PSScriptRoot "AGENTS.md"
$featureManifestPath = Join-Path $PSScriptRoot "features.json"
$testManifestPath = Join-Path $modulesDir "Lib_TestManifest.bas"
$structuredTestResultsPath = Join-Path $env:TEMP "BeaverAddin.TestResults.tsv"
$script:StageResults = New-Object System.Collections.ArrayList

function Stop-Script {
    param(
        [string]$Message,
        [int]$ExitCode = 1
    )

    Write-StageSummary
    Write-Host $Message -ForegroundColor Red
    exit $ExitCode
}

function Add-StageResult {
    param(
        [string]$Stage,
        [string]$Status,
        [string]$Details = ""
    )

    [void]$script:StageResults.Add([pscustomobject]@{
        Stage = $Stage
        Status = $Status
        Details = $Details
        Timestamp = Get-Date
    })
}

function Write-StageSummary {
    if ($script:StageResults.Count -eq 0) { return }

    Write-Host ""
    Write-Host "Stage Summary" -ForegroundColor Cyan
    foreach ($stage in $script:StageResults) {
        $color = if ($stage.Status -eq "success") { "Green" } elseif ($stage.Status -eq "skipped") { "Yellow" } else { "Red" }
        $detailText = if ([string]::IsNullOrWhiteSpace($stage.Details)) { "" } else { " - $($stage.Details)" }
        Write-Host ("  [{0}] {1}{2}" -f $stage.Status.ToUpper(), $stage.Stage, $detailText) -ForegroundColor $color
    }
}

function Get-FeatureManifest {
    param([string]$ManifestPath)
    if (-not (Test-Path $ManifestPath)) {
        throw "Feature manifest not found: $ManifestPath"
    }
    return Get-Content $ManifestPath -Raw | ConvertFrom-Json
}

function Test-ReleaseTierIncluded {
    param(
        [string]$ReleaseTier,
        [bool]$IncludeDev
    )

    if ([string]::IsNullOrWhiteSpace($ReleaseTier)) {
        return $true
    }

    if ($ReleaseTier -eq "dev" -and -not $IncludeDev) {
        return $false
    }

    return $true
}

function Sync-FeatureManifest {
    param(
        [string]$ManifestPath,
        [string]$ConfigPath,
        [string]$RibbonPath,
        [bool]$IncludeDev
    )

    Write-Host "Syncing feature manifest..." -ForegroundColor Cyan
    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath
    $config = if (Test-Path $ConfigPath) {
        Get-Content $ConfigPath -Raw | ConvertFrom-Json
    } else {
        [pscustomobject]@{}
    }

    $enabledFeatures = @($manifest.Features | Where-Object {
        Test-ReleaseTierIncluded -ReleaseTier $_.ReleaseTier -IncludeDev $IncludeDev
    })
    $enabledFeatureIds = @($enabledFeatures | ForEach-Object { $_.ControlId })
    $enabledHotkeys = @($manifest.Hotkeys | Where-Object {
        Test-ReleaseTierIncluded -ReleaseTier $_.ReleaseTier -IncludeDev $IncludeDev
    })

    $icons = [ordered]@{}
    foreach ($feature in $enabledFeatures) {
        $icons[$feature.ControlId] = $feature.Icon
    }

    $groupXml = foreach ($group in $manifest.Groups) {
        $groupFeatures = @($enabledFeatures | Where-Object { $group.Features -contains $_.ControlId })
        if ($groupFeatures.Count -eq 0) { continue }

        $buttonXml = foreach ($feature in $groupFeatures) {
            '          <button id="{0}" label="{1}" getImage="Ribbon_GetIcon" size="large" onAction="{2}" keytip="{3}" screentip="{4}" supertip="{5}" />' -f `
                $feature.ControlId,
                [System.Security.SecurityElement]::Escape($feature.Label),
                $feature.OnAction,
                $feature.Keytip,
                [System.Security.SecurityElement]::Escape($feature.Screentip),
                [System.Security.SecurityElement]::Escape($feature.Supertip)
        }

@"
        <group id="$($group.Id)" label="$([System.Security.SecurityElement]::Escape($group.Label))">
$($buttonXml -join "`r`n")
        </group>
"@
    }

    $ribbonContent = @"
<!--
  @Module: ribbon.xml
  @Category: UI
  @Description: Generated Ribbon UI definition for the Beaver Add-in.
  @ManagedBy: BeaverAddin Agent
  @Source: features.json via Update.ps1
-->
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="$($manifest.Tab.Id)" label="$([System.Security.SecurityElement]::Escape($manifest.Tab.Label))" keytip="$($manifest.Tab.Keytip)">
$($groupXml -join "`r`n")
      </tab>
    </tabs>
  </ribbon>
</customUI>
"@
    [System.IO.File]::WriteAllText($RibbonPath, $ribbonContent, [System.Text.Encoding]::ASCII)

    $config | Add-Member -NotePropertyName Hotkeys -NotePropertyValue $enabledHotkeys -Force
    $config | Add-Member -NotePropertyName Icons -NotePropertyValue ([pscustomobject]$icons) -Force
    if (-not $config.PSObject.Properties.Name.Contains("FeatureFlags")) {
        $config | Add-Member -NotePropertyName FeatureFlags -NotePropertyValue ([pscustomobject]@{}) -Force
    }
    $config.FeatureFlags = [pscustomobject]@{
        ReleaseTier = if ($IncludeDev) { "dev" } else { "stable" }
        IncludeDevFeatures = $IncludeDev
        ManifestFile = [System.IO.Path]::GetFileName($ManifestPath)
        GeneratedFeatureCount = $enabledFeatureIds.Count
    }

    $configJson = $config | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($ConfigPath, $configJson, [System.Text.Encoding]::ASCII)
    Write-Host "  Manifest sync complete. Features: $($enabledFeatureIds.Count), Hotkeys: $($enabledHotkeys.Count)." -ForegroundColor Green
}

function Sync-TestManifest {
    param(
        [string]$SourceDir,
        [string]$OutputPath
    )

    Write-Host "Generating test manifest..." -ForegroundColor Cyan
    $testProcedures = @()
    $moduleFiles = @(Get-ChildItem -Path $SourceDir -Filter *.bas)

    foreach ($file in $moduleFiles) {
        if ($file.Name -eq "Lib_TestManifest.bas") { continue }
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $matches = Select-String -Path $file.FullName -Pattern '^\s*Public Sub (Test_[A-Za-z0-9_]+)\s*\('
        foreach ($match in $matches) {
            $testProcedures += [pscustomobject]@{
                Module = $moduleName
                Procedure = $match.Matches[0].Groups[1].Value
            }
        }
    }

    $lines = @(
        'Attribute VB_Name = "Lib_TestManifest"',
        'Option Explicit',
        '',
        ''' @Module: Lib_TestManifest',
        ''' @Category: Infrastructure',
        ''' @Description: Generated test manifest that orchestrates all Test_* procedures.',
        ''' @ManagedBy: BeaverAddin Agent',
        ''' @Dependencies: Infra_Error'
    )
    $lines += ''
    $lines += 'Public Sub RunGeneratedTests()'
    $lines += '    Dim tracker As Object: Set tracker = Infra_Error.Track("RunGeneratedTests")'
    $lines += '    On Error GoTo ErrHandler'
    $lines += ''

    if ($testProcedures.Count -eq 0) {
        $lines += 'CleanExit:'
        $lines += '    Exit Sub'
        $lines += ''
        $lines += 'ErrHandler:'
        $lines += '    Infra_Error.HandleError "RunGeneratedTests", Err'
        $lines += '    Resume CleanExit'
    } else {
        foreach ($test in $testProcedures | Sort-Object Module, Procedure) {
            $lines += "    $($test.Module).$($test.Procedure)"
        }
        $lines += ''
        $lines += 'CleanExit:'
        $lines += '    Exit Sub'
        $lines += ''
        $lines += 'ErrHandler:'
        $lines += '    Infra_Error.HandleError "RunGeneratedTests", Err'
        $lines += '    Resume CleanExit'
    }

    $lines += 'End Sub'
    [System.IO.File]::WriteAllText($OutputPath, ($lines -join "`r`n"), [System.Text.Encoding]::ASCII)
    Write-Host "  Test manifest generated with $($testProcedures.Count) test(s)." -ForegroundColor Green
}

function Read-StructuredTestResults {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return $null
    }

    $summary = $null
    $results = @()

    foreach ($line in Get-Content $Path) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split "`t"
        if ($parts[0] -eq "SUMMARY" -and $parts.Count -ge 4) {
            $summary = [pscustomobject]@{
                Total = [int]$parts[1]
                Passed = [int]$parts[2]
                Failed = [int]$parts[3]
            }
        } elseif ($parts[0] -eq "RESULT" -and $parts.Count -ge 6) {
            $results += [pscustomobject]@{
                Name = $parts[1]
                Passed = [bool]::Parse($parts[2])
                DurationMs = [int]$parts[3]
                Category = $parts[4]
                Message = $parts[5]
            }
        }
    }

    return [pscustomobject]@{
        Summary = $summary
        Results = $results
    }
}

function Get-EnabledHeadlessCallbacks {
    param(
        [string]$ManifestPath,
        [bool]$IncludeDev
    )

    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath
    return @($manifest.Features | Where-Object {
        (Test-ReleaseTierIncluded -ReleaseTier $_.ReleaseTier -IncludeDev $IncludeDev) -and $_.RuntimeTestMode -eq "headless"
    })
}

function Get-ExcelExecutablePath {
    $appPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
    )

    foreach ($path in $appPaths) {
        try {
            $key = Get-Item $path -ErrorAction Stop
            $exePath = $key.GetValue("")
            if ($exePath -and (Test-Path $exePath)) {
                return $exePath
            }
        } catch { }
    }

    return $null
}

function Remove-OrphanedExcelProcesses {
    $excelProcesses = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
    if ($excelProcesses.Count -eq 0) {
        return $false
    }

    $visibleExcel = @(
        $excelProcesses | Where-Object {
            $_.MainWindowHandle -ne 0 -or -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle)
        }
    )

    if ($visibleExcel.Count -gt 0) {
        return $false
    }

    Write-Host "  Found $($excelProcesses.Count) background Excel process(es) with no visible window. Cleaning up..." -ForegroundColor Yellow
    $stoppedAny = $false
    foreach ($process in $excelProcesses) {
        try {
            Stop-Process -Id $process.Id -Force -ErrorAction Stop
            $stoppedAny = $true
        } catch {
            Write-Warning "  Failed to stop orphaned Excel process $($process.Id): $($_.Exception.Message)"
        }
    }

    if ($stoppedAny) {
        Start-Sleep -Seconds 2
    }

    return $stoppedAny
}

function New-NormalizedImportCopy {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$TempRoot
    )

    if (-not (Test-Path $TempRoot)) {
        New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null
    }

    $normalizedPath = Join-Path $TempRoot ([System.IO.Path]::GetFileName($SourcePath))
    $content = [System.IO.File]::ReadAllText($SourcePath)
    $content = $content -replace "(?<!`r)`n", "`r`n"
    [System.IO.File]::WriteAllText($normalizedPath, $content, [System.Text.Encoding]::ASCII)

    return $normalizedPath
}

function Start-ExcelApplication {
    param(
        [string]$Purpose
    )

    try {
        return New-Object -ComObject Excel.Application
    } catch {
        $directComError = $_.Exception
    }

    $existingExcel = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
    if ($existingExcel.Count -gt 0) {
        if (Remove-OrphanedExcelProcesses) {
            try {
                return New-Object -ComObject Excel.Application
            } catch {
                $directComError = $_.Exception
                $existingExcel = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
            }
        }
    }

    if ($existingExcel.Count -gt 0) {
        throw "Failed to start Excel COM automation for ${Purpose}: $($directComError.Message)"
    }

    $excelExe = Get-ExcelExecutablePath
    if (-not $excelExe) {
        throw "Failed to start Excel COM automation for ${Purpose}: $($directComError.Message)"
    }

    try {
        $startedProcess = Start-Process -FilePath $excelExe -PassThru -ErrorAction Stop
    } catch {
        throw "Failed to start Excel COM automation for $Purpose. COM activation failed with '$($directComError.Message)', and launching EXCEL.EXE also failed: $($_.Exception.Message)"
    }

    $deadline = (Get-Date).AddSeconds(20)
    do {
        Start-Sleep -Milliseconds 500
        try {
            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            if ($excel -and $excel.Hwnd) {
                $hwndPid = 0
                [WindowScraper]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$hwndPid) | Out-Null
                if ($hwndPid -eq $startedProcess.Id) {
                    return $excel
                }
            }
        } catch { }
    } while ((Get-Date) -lt $deadline -and -not $startedProcess.HasExited)

    if (-not $startedProcess.HasExited) {
        Stop-Process -Id $startedProcess.Id -Force -ErrorAction SilentlyContinue
    }

    throw "Failed to start Excel COM automation for $Purpose. COM activation failed with '$($directComError.Message)', and Excel could not be attached after launching EXCEL.EXE."
}

# ==============================================================================
# Helper: Set-RibbonUiErrors
# Purpose: Enables/Disables 'Show add-in user interface errors' in Registry.
# ==============================================================================
function Set-RibbonUiErrors {
    param ([bool]$Enabled)
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Common\General"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    $val = if ($Enabled) { 1 } else { 0 }
    Set-ItemProperty -Path $regPath -Name "ShowErrors" -Value $val -Type DWord -Force
}

# ==============================================================================
# Helper: Scrape-ExcelRibbonErrors (C#)
# Purpose: Finds, reads text from, and closes Ribbon UI error dialogs.
# ==============================================================================
$scraperCode = @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class WindowScraper {
    public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumThreadDelegate lpfn, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumThreadDelegate lpfn, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

    public static string ScrapeAndClose(int processId, int timeoutSeconds) {
        var result = new StringBuilder();
        var seenTexts = new HashSet<string>();
        var startTime = DateTime.Now;

        // Loop to catch windows that might appear with a delay or sequentially
        while ((DateTime.Now - startTime).TotalSeconds < timeoutSeconds) {
            EnumWindows((hWnd, lParam) => {
                int windowPid;
                GetWindowThreadProcessId(hWnd, out windowPid);
                if (windowPid == processId) {
                    var title = new StringBuilder(256);
                    GetWindowText(hWnd, title, 256);
                    string sTitle = title.ToString();
                    
                    // Check for common Office/Excel error window titles
                    if (sTitle.Contains("Microsoft Excel") || 
                        sTitle.Contains("Custom UI") || 
                        sTitle.Contains("Runtime Error") ||
                        sTitle.Contains("Microsoft Visual Basic")) {
                        
                        bool foundNewText = false;
                        EnumChildWindows(hWnd, (hChild, lChild) => {
                            var text = new StringBuilder(1024);
                            GetWindowText(hChild, text, 1024);
                            var sText = text.ToString().Trim();
                            // Collect text but avoid common UI buttons
                            if (sText.Length > 0 && 
                                !sText.Equals("OK", StringComparison.OrdinalIgnoreCase) && 
                                !sText.Equals("Cancel", StringComparison.OrdinalIgnoreCase) && 
                                !sText.Equals("Close", StringComparison.OrdinalIgnoreCase) &&
                                !sText.Equals("Help", StringComparison.OrdinalIgnoreCase) &&
                                !seenTexts.Contains(sText)) {
                                result.AppendLine(sText);
                                seenTexts.Add(sText);
                                foundNewText = true;
                            }
                            return true;
                        }, IntPtr.Zero);

                        if (foundNewText || sTitle.Contains("Visual Basic")) {
                            // Close window (WM_CLOSE = 0x10) to unblock the main Excel thread
                            PostMessage(hWnd, 0x0010, IntPtr.Zero, IntPtr.Zero);
                        }
                    }
                }
                return true;
            }, IntPtr.Zero);
            
            System.Threading.Thread.Sleep(500); // Poll every 500ms
        }
        return result.ToString().Trim();
    }
}
"@
Add-Type -TypeDefinition $scraperCode -ErrorAction SilentlyContinue

# ==============================================================================
# Function: Test-RibbonValidity
# Purpose:  Validates ribbon.xml for schema errors, duplicate IDs, and missing callbacks.
# ==============================================================================
function Test-RibbonValidity {
    param (
        [string]$XmlPath,
        [string]$ModulesDir
    )

    if (-not (Test-Path $XmlPath)) { return $true }

    Write-Host "Validating Ribbon XML..." -ForegroundColor Cyan
    $isValid = $true
    $absoluteXmlPath = Resolve-Path $XmlPath

    # 1. Schema Validation using .NET XmlReader (catches malformed XML and schema violations)
    try {
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.XmlResolver = $null # Prevent hanging on URL resolution
        $settings.ValidationType = [System.Xml.ValidationType]::Schema
        # We don't have the .xsd file locally, but we can enable 'ProcessInlineSchema' 
        # or rely on the namespace if the resolver can reach it. 
        # For Office, the schemas are standard. We'll at least catch well-formedness and basic structure.
        $settings.ValidationFlags = $settings.ValidationFlags -bor [System.Xml.Schema.XmlSchemaValidationFlags]::ProcessIdentityConstraints
        $settings.ValidationFlags = $settings.ValidationFlags -bor [System.Xml.Schema.XmlSchemaValidationFlags]::ReportValidationWarnings

        $onValidationError = [System.Xml.Schema.ValidationEventHandler] {
            param($evtSource, $e)
            # Suppress "Could not find schema information" noise - common without local XSDs
            if ($e.Message -match "Could not find schema information") { return }
            
            $script:isValid = $false
            $line = $e.Exception.LineNumber
            $col = $e.Exception.LinePosition
            Write-Host "  Ribbon XML Error [Line $line, Col $col]: $($e.Message)" -ForegroundColor Red
        }
        $settings.add_ValidationEventHandler($onValidationError)

        $reader = [System.Xml.XmlReader]::Create($absoluteXmlPath, $settings)
        while ($reader.Read()) { }
        $reader.Close()
    } catch {
        Write-Error "Ribbon XML failed to load or is malformed: $($_.Exception.Message)"
        $isValid = $false
    }

    if (-not $isValid) { return $false }

    # 2. Duplicate ID and Callback Check (Logical checks on valid XML)
    $xml = [xml](Get-Content $XmlPath -Raw)
    
    # Duplicate ID Check
    $ids = $xml.SelectNodes("//@id") | ForEach-Object { $_.Value }
    $duplicates = $ids | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($duplicates) {
        Write-Error "Duplicate IDs found in ribbon.xml: $($duplicates.Name -join ', ')"
        $isValid = $false
    }

    # Callback Verification
    $callbacks = $xml.SelectNodes("//@onAction") | ForEach-Object { $_.Value } | Select-Object -Unique
    if ($callbacks) {
        Write-Host "  Checking $($callbacks.Count) callbacks across all modules..."
        $vbaFiles = Get-ChildItem -Path $ModulesDir -Include *.bas, *.cls -Recurse
        $vbaCode = ""
        foreach ($f in $vbaFiles) { $vbaCode += Get-Content $f.FullName -Raw }
        
        foreach ($cb in $callbacks) {
            if ($vbaCode -notmatch "Sub\s+$cb\s*\(") {
                Write-Error "Ribbon callback '$cb' not found in any module in $ModulesDir"
                $isValid = $false
            }
        }
    }

    return $isValid
}

# ==============================================================================
# Function: Update-RibbonInWorkbook
# Purpose:  Injects customUI14.xml into the .xlsm archive.
# ==============================================================================
function Update-RibbonInWorkbook {
    param ([string]$WorkbookPath, [string]$RibbonXmlPath)
    if (-not (Test-Path $RibbonXmlPath)) { return }
    Write-Host "Injecting Ribbon XML..."
    try {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)
        
        $ribbonEntryPath = "customUI/customUI14.xml"
        $ribbonEntry = $zip.GetEntry($ribbonEntryPath)
        if ($null -ne $ribbonEntry) { $ribbonEntry.Delete() }
        $ribbonEntry = $zip.CreateEntry($ribbonEntryPath)
        $writer = New-Object System.IO.StreamWriter($ribbonEntry.Open())
        $writer.Write((Get-Content $RibbonXmlPath -Raw))
        $writer.Close()

        $zip.Dispose()
        Write-Host "  Ribbon XML injected successfully."
    } catch {
        Write-Error "Failed to update Ribbon XML: $($_.Exception.Message)"
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

# ==============================================================================
# Function: Invoke-VbaSyntaxCheck
# Purpose:  Performs a basic regex-based scan of VBA files for common errors
#           like missing End Sub, End If, etc.
# ==============================================================================
function Invoke-VbaSyntaxCheck {
    param ([string]$SourceDir)
    Write-Host "Linting VBA Files..." -ForegroundColor Cyan
    $vbaFiles = @(Get-ChildItem -Path $SourceDir -Include *.bas, *.cls, *.frm -Recurse)
    # Include ThisWorkbook.cls from root
    $thisWorkbook = Join-Path $PSScriptRoot "ThisWorkbook.cls"
    if (Test-Path $thisWorkbook) { $vbaFiles += Get-Item $thisWorkbook }

    $allPassed = $true
    foreach ($file in $vbaFiles) {
        $rawLines = Get-Content $file.FullName
        $fileName = $file.Name
        
        # Join line continuations (_) while tracking original line numbers
        $content = @()
        $originalLineNumbers = @() # Maps index in $content to line number in $rawLines
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

        # Check for matching blocks with line tracking
        $blocks = @(
            @{ Name = "Sub";     Start = "^\s*(?:Public |Private |Static )?Sub\s+";     End = "^\s*End Sub" }
            @{ Name = "Function";Start = "^\s*(?:Public |Private |Static )?Function\s+";End = "^\s*End Function" }
            @{ Name = "Property";Start = "^\s*(?:Public |Private )?Property\s+(?:Get|Let|Set)\s+"; End = "^\s*End Property" }
            @{ Name = "If";      Start = "^\s*If\s+.*Then\s*(?:'.*)?$"; End = "^\s*End If" } 
        )

        foreach ($b in $blocks) {
            $stack = New-Object System.Collections.Generic.List[int]
            for ($i = 0; $i -lt $content.Count; $i++) {
                $lineNum = $originalLineNumbers[$i]
                if ($content[$i] -match $b.Start) {
                    $stack.Add($lineNum)
                } elseif ($content[$i] -match $b.End) {
                    if ($stack.Count -gt 0) {
                        $stack.RemoveAt($stack.Count - 1)
                    } else {
                        Write-Host "  [$fileName] Syntax Error: Unexpected '$($b.End.Trim())' at line $lineNum (No matching start found)." -ForegroundColor Red
                        $allPassed = $false
                    }
                }
            }
            
            foreach ($startLine in $stack) {
                Write-Host "  [$fileName] Syntax Error: Mismatched '$($b.Name)' starting at line $startLine (No matching end found)." -ForegroundColor Red
                $allPassed = $false
            }
        }
    }
    return $allPassed
}

# ==============================================================================
# Function: Update-Version
# Purpose:  Increments the patch/build version in config.json.
# ==============================================================================
function Update-Version {
    param ([string]$ConfigPath)
    if (-not (Test-Path $ConfigPath)) { return }
    
    Write-Host "Updating version in config.json..." -ForegroundColor Cyan
    try {
        $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        $version = $config.AddinIdentity.Version
        $parts = $version.Split('.')
        if ($parts.Count -eq 3) {
            $build = [int]$parts[2] + 1
            $newVersion = "$($parts[0]).$($parts[1]).$build"
            $config.AddinIdentity.Version = $newVersion
            $json = $config | ConvertTo-Json -Depth 10
            [System.IO.File]::WriteAllText($ConfigPath, $json)
            Write-Host "  Version incremented: $version -> $newVersion" -ForegroundColor Green
        }
    } catch {
        Write-Warning "  Failed to auto-increment version: $($_.Exception.Message)"
    }
}

# ==============================================================================
# Function: Update-AgentsGuideDate
# Purpose:  Updates the "Last updated:" date in AGENTS.md to today.
# ==============================================================================
function Update-AgentsGuideDate {
    param ([string]$AgentsMdPath)
    if (-not (Test-Path $AgentsMdPath)) { return }
    
    Write-Host "Updating date in AGENTS.md..." -ForegroundColor Cyan
    try {
        $content = Get-Content $AgentsMdPath -Raw
        $dateStr = (Get-Date).ToString("yyyy-MM-dd")
        $newContent = $content -replace "Last updated: \d{4}-\d{2}-\d{2}", "Last updated: $dateStr"
        [System.IO.File]::WriteAllText($AgentsMdPath, $newContent)
        Write-Host "  AGENTS.md date updated to: $dateStr" -ForegroundColor Green
    } catch {
        Write-Warning "  Failed to update AGENTS.md date: $($_.Exception.Message)"
    }
}

# ==============================================================================
# Function: Invoke-EnhancedLinting
# Purpose:  Checks for Option Explicit, @Module metadata, and Error Handling 
#           boilerplate in all .bas/.cls files.
# ==============================================================================
function Invoke-EnhancedLinting {
    param ([string]$SourceDir)
    Write-Host "Running Enhanced Linting..." -ForegroundColor Cyan
    $vbaFiles = @(Get-ChildItem -Path $SourceDir -Include *.bas, *.cls, *.frm -Recurse)
    $thisWorkbook = Join-Path $PSScriptRoot "ThisWorkbook.cls"
    if (Test-Path $thisWorkbook) { $vbaFiles += Get-Item $thisWorkbook }
    $allPassed = $true

    foreach ($file in $vbaFiles) {
        $content = Get-Content $file.FullName -Raw
        $lines = Get-Content $file.FullName
        $fileName = $file.Name

        # 1. Check for Option Explicit
        if ($content -notmatch "(?m)^Option Explicit") {
            Write-Host "  [$fileName] Error: Missing 'Option Explicit' at the top of the file." -ForegroundColor Red
            $allPassed = $false
        }

        # 2. Check for @Module metadata
        if ($content -notmatch "' @Module:") {
            Write-Host "  [$fileName] Error: Missing '@Module' metadata header." -ForegroundColor Red
            $allPassed = $false
        }

        # 3. Procedure Boilerplate Check (PushContext / HandleError)
        # We look for Public Sub/Function that aren't property getters/setters or event handlers
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            if ($line -match "^\s*Public (?:Sub|Function)\s+([a-zA-Z0-9_]+)") {
                $procName = $matches[1]
                
                # Skip common Excel events or very short helper functions if they don't need context
                if ($procName -match "^(?:Workbook_|Worksheet_|App_)" -or $file.Name -eq "Lib_JsonConverter.bas" -or $file.Name -match "^(?:Infra_Error\.(bas|cls)|Infra_ContextTracker\.cls|Infra_Diagnostics\.bas|Infra_OperationContext\.cls|StateStore\.cls|AppContainer\.cls|Infra_Config\.(cls|bas)|Infra_ConfigModel\.cls|I[A-Z][a-zA-Z0-9_\-]*\.cls|Infra_AppStateGuard\.cls|Infra_AppState\.bas)$") {
                    continue
                }

                # Scan the procedure body (up to the next End Sub/Function)
                $j = $i + 1
                $foundPush = $false
                $foundPop = $false
                $foundErrorGoto = $false
                $foundHandleError = $false
                $procBody = ""
                
                while ($j -lt $lines.Count -and $lines[$j] -notmatch "^\s*End (?:Sub|Function)") {
                    $procBody += $lines[$j] + "`n"
                    if ($lines[$j] -match "PushContext\s+""$procName""" -or $lines[$j] -match "Infra_Error\.Track\s*\(""$procName""\)") { $foundPush = $true }
                    if ($lines[$j] -match "PopContext" -or $lines[$j] -match "Dim\s+\w+\s+As\s+Object:\s*Set\s+\w+\s*=\s*Infra_Error\.Track") { $foundPop = $true }
                    if ($lines[$j] -match "On Error GoTo\s+\w+") { $foundErrorGoto = $true }
                    if ($lines[$j] -match "HandleError\s+""$procName""") { $foundHandleError = $true }
                    $j++
                }

                if (-not $foundPush) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' missing context tracking (PushContext or Track)." -ForegroundColor Red
                    $allPassed = $false
                }
                if (-not $foundPop) {
                    # If we use RAII tracking, PopContext isn't needed manually
                    if ($lines[$i+1] -notmatch "Infra_Error\.Track") {
                        Write-Host "  [$fileName] Error: Procedure '$procName' missing 'PopContext'." -ForegroundColor Red
                        $allPassed = $false
                    }
                }
                if (-not $foundErrorGoto) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' missing 'On Error GoTo'." -ForegroundColor Red
                    $allPassed = $false
                }
                if (-not $foundHandleError) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' missing 'HandleError ""$procName""'." -ForegroundColor Red
                    $allPassed = $false
                }
            }
        }
    }
    return $allPassed
}

# --- 1. PRE-DEPLOYMENT VALIDATION ---
$configPath = Join-Path $PSScriptRoot "config.json"

Sync-FeatureManifest -ManifestPath $featureManifestPath -ConfigPath $configPath -RibbonPath $ribbonXmlPath -IncludeDev:$IncludeDevFeatures
Add-StageResult -Stage "manifest_sync" -Status "success" -Details "features synced from features.json"
Sync-TestManifest -SourceDir $modulesDir -OutputPath $testManifestPath
Add-StageResult -Stage "test_manifest_generation" -Status "success"

$validRibbon = Test-RibbonValidity -XmlPath $ribbonXmlPath -ModulesDir $modulesDir
$validVba = Invoke-VbaSyntaxCheck -SourceDir $modulesDir
$validLint = Invoke-EnhancedLinting -SourceDir $modulesDir

if (-not ($validRibbon -and $validVba -and $validLint)) {
    Add-StageResult -Stage "validation" -Status "failure" -Details "Pre-deployment validation failed"
    Write-Host "Pre-deployment validation failed. Fix the issues above and retry." -ForegroundColor Red
    exit 1
}
Add-StageResult -Stage "validation" -Status "success"

# --- 2. ENVIRONMENT CHECKS ---
if (-not (Test-Path $excelPath)) { Write-Error "Excel file not found."; exit }
$lockFile = Join-Path $PSScriptRoot ("~$" + (Split-Path $excelPath -Leaf))
if (Test-Path $lockFile) {
    Write-Host "Excel file is open. Attempting to close it..." -ForegroundColor Yellow
    try {
        $activeExcel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $activeExcel.DisplayAlerts = $false
        foreach ($wb in $activeExcel.Workbooks) {
            if ($wb.FullName -eq $excelPath) {
                $wb.Close($true)
                Write-Host "  Closed $($wb.Name) successfully." -ForegroundColor Green
                break
            }
        }
        $activeExcel.DisplayAlerts = $true
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($activeExcel) | Out-Null
    } catch {
        Write-Warning "Could not close gracefully via COM. Force closing Excel..."
        Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
    }
    
    # Wait for the lock file to be released
    Start-Sleep -Seconds 2
    if (Test-Path $lockFile) {
        Stop-Script "Excel file is still open. Please close it manually and retry."
    }
}

# --- 3. BEGIN UPDATE ---
Write-Host "Starting Excel... (This may take a moment)"
try {
    $excel = Start-ExcelApplication -Purpose "workbook update"
} catch {
    Stop-Script $_.Exception.Message
}
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($excelPath)
    $vbaProject = $workbook.VBProject

    Write-Host "Updating modules..."
    # Components collection can change while iterating; use a fresh list
    $compsToRemove = @()
    for ($i = 1; $i -le $vbaProject.VBComponents.Count; $i++) {
        $comp = $vbaProject.VBComponents.Item($i)
        # Type 1: Standard Module, 2: Class Module, 3: Form
        if (($comp.Type -ge 1 -and $comp.Type -le 3) -and ($comp.Name -ne "ThisWorkbook")) {
            $compsToRemove += $comp
        }
    }
    
    foreach ($comp in $compsToRemove) {
        try {
            $vbaProject.VBComponents.Remove($comp)
        } catch {
            Write-Warning "  Could not remove component: $($comp.Name)"
        }
    }

    # Import both .bas and .cls files
    # Note: -Include requires a wildcard in the path to work correctly in some PS versions
    $vbaSourceFiles = Get-ChildItem -Path $modulesDir | Where-Object { $_.Extension -match "\.(bas|cls|frm)$" }
    
    $tempImportDir = Join-Path ([System.IO.Path]::GetTempPath()) ("BeaverAddin-VbaImport-" + [System.Guid]::NewGuid().ToString("N"))
    try {
        foreach ($file in $vbaSourceFiles) {
            Write-Host "  Importing $($file.Name)..."
            $importPath = New-NormalizedImportCopy -SourcePath $file.FullName -TempRoot $tempImportDir
            
            if ($file.Extension -eq ".frm") {
                $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
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
    
    # ThisWorkbook specially handled
    if (Test-Path $desktopThisWorkbookCls) {
        Write-Host "  Updating ThisWorkbook..."
        $twCode = $vbaProject.VBComponents.Item("ThisWorkbook").CodeModule
        if ($twCode.CountOfLines -gt 0) { $twCode.DeleteLines(1, $twCode.CountOfLines) }
        # Strip header lines that Excel manages automatically (e.g., VERSION, BEGIN, Attribute)
        $lines = Get-Content $desktopThisWorkbookCls | Where-Object { 
            $_ -notmatch "^VERSION\s+\d+\.\d+" -and 
            $_ -notmatch "^BEGIN\s*$" -and 
            $_ -notmatch "^\s+MultiUse\s*=" -and 
            $_ -notmatch "^END\s*$" -and 
            $_ -notmatch "^Attribute\s+"
        }
        $twCode.AddFromString([string]::Join("`r`n", $lines))
    }

    # Compilation Check
    Write-Host "Compiling VBA Project..."
    if ($null -ne $excel.VBE) {
        Write-Host "  VBE object found."
        
        # Search all command bars recursively
        $btn = $null
        foreach ($cb in $excel.VBE.CommandBars) {
            # Helper function for recursive search
            function Find-ControlRecursive {
                param($Parent)
                foreach ($c in $Parent.Controls) {
                    try {
                        if ($c.Id -eq 578 -or $c.Caption -match "Compile") {
                            return $c
                        }
                        if ($c.Type -eq 10 -or $c.Type -eq 12) { # Popup or ButtonPopup
                            $found = Find-ControlRecursive -Parent $c
                            if ($found) { return $found }
                        }
                    } catch { }
                }
                return $null
            }
            $btn = Find-ControlRecursive -Parent $cb
            if ($btn) { break }
        }

        if ($null -ne $btn) { 
            Write-Host "  Found '$($btn.Caption)' button (Enabled: $($btn.Enabled))."
            if ($btn.Enabled) {
                # Suppress the blocking MsgBox that VBA shows on compile errors
                $excel.DisplayAlerts = $false 
                
                try {
                    Write-Host "  Executing compile..."
                    $btn.Execute()
                } catch {
                    Write-Host "  Execute() threw an exception: $($_.Exception.Message)"
                }
                
                # If still enabled after execution, compilation failed
                if ($btn.Enabled) {
                    Write-Host "  ERROR: VBA Compilation failed (Button still enabled)." -ForegroundColor Red
                    
                    # Attempt to extract the error details from the VBE
                    # When compile fails, the VBE usually highlights the error line 
                    # in the active code pane.
                    try {
                        $activePane = $excel.VBE.ActiveCodePane
                        if ($null -ne $activePane) {
                            $modName = $activePane.CodeModule.Name
                            
                            # Get the current selection (highlighted error)
                            $startLine = 0; $startCol = 0; $endLine = 0; $endCol = 0
                            $activePane.GetSelection([ref]$startLine, [ref]$startCol, [ref]$endLine, [ref]$endCol)
                            
                            $errorLineText = $activePane.CodeModule.Lines($startLine, 1).Trim()
                            
                            Write-Host "  [Diagnostics] Module: $modName" -ForegroundColor Yellow
                            Write-Host "  [Diagnostics] Line $($startLine): $errorLineText" -ForegroundColor Yellow
                            
                            throw "VBA Compilation failed in module '$modName' at line $($startLine): '$errorLineText'. Please fix the syntax or missing definitions."
                        } else {
                            Write-Host "  [Diagnostics] No ActiveCodePane found after failure." -ForegroundColor Yellow
                            throw "VBA Compilation failed. Check your code for syntax or definition errors."
                        }
                    } catch {
                        # If we fail to get the active pane details, throw a generic error
                        if ($_.Exception.Message -match "VBA Compilation failed") {
                            throw $_.Exception.Message
                        } else {
                            Write-Host "  [Diagnostics] Error retrieving active pane: $($_.Exception.Message)" -ForegroundColor Red
                            throw "VBA Compilation failed. Check your code for 'Variable not defined' or syntax errors."
                        }
                    }
                } else {
                    Write-Host "  Compilation successful." -ForegroundColor Green
                }
            } else {
                Write-Host "  Project already compiled." -ForegroundColor Gray
            }
        } else {
            Write-Host "  'Compile Project' button NOT found. Listing available CommandBars:" -ForegroundColor Yellow
            foreach ($cb in $excel.VBE.CommandBars) {
                Write-Host "    - $($cb.Name) (Visible: $($cb.Visible))"
            }
        }
    } else {
        Write-Host "  VBE object NOT found." -ForegroundColor Yellow
    }

    if ($BumpVersion) {
        Update-Version -ConfigPath $configPath
    } else {
        Write-Host "Skipping version bump (pass -BumpVersion for release builds)." -ForegroundColor Gray
    }
    if ($RefreshAgentsGuide) {
        Update-AgentsGuideDate -AgentsMdPath $agentsMdPath
    } else {
        Write-Host "Skipping AGENTS.md date refresh (pass -RefreshAgentsGuide when needed)." -ForegroundColor Gray
    }
    $workbook.Save()
    $workbook.Close($true)
    $workbook = $null
    Write-Host "SUCCESS: Modules updated."
    Add-StageResult -Stage "workbook_update" -Status "success"

    # --- RIBBON ---
    Update-RibbonInWorkbook -WorkbookPath $excelPath -RibbonXmlPath $ribbonXmlPath
    Add-StageResult -Stage "ribbon_injection" -Status "success"
} catch {
    Add-StageResult -Stage "workbook_update" -Status "failure" -Details $_.Exception.Message
    Stop-Script "An error occurred: $($_.Exception.Message)"
    # If compilation fails, we shouldn't run tests
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# --- 4. RUNTIME TESTING ---
if ($SkipRuntimeTests) {
    Add-StageResult -Stage "runtime_tests" -Status "skipped" -Details "Skipped by -SkipRuntimeTests"
    Write-StageSummary
    Write-Host "Skipping runtime testing (`-SkipRuntimeTests`)." -ForegroundColor Yellow
    exit 0
}

Write-Host "Starting Runtime Testing..." -ForegroundColor Cyan

# We use a fresh Excel instance to ensure clean state after ribbon injection
try {
    $testExcel = Start-ExcelApplication -Purpose "runtime testing"
} catch {
    Stop-Script $_.Exception.Message
}
$testExcel.Visible = $false # Keep hidden to prevent focus stealing
$testExcel.DisplayAlerts = $false
$testWorkbook = $null

try {
    # 1. Check for Ribbon UI Errors (Invalid imageMso, etc.)
    Set-RibbonUiErrors -Enabled $true
    
    # Get the PID of the Excel instance we just started
    # We use a slightly more complex way to ensure we get the right one if multiple are running
    $excelPid = 0
    try {
        $excelPid = (Get-Process -Name "EXCEL" | Sort-Object StartTime -Descending | Select-Object -First 1).Id
    } catch { }

    if ($excelPid -gt 0) {
        # Start a background job to watch for and close error dialogs
        $watcher = Start-Job -ScriptBlock {
            param($ProcessIdToScrape, $code)
            Add-Type -TypeDefinition $code -ErrorAction SilentlyContinue
            # Poll for Ribbon error dialogs for up to 20 seconds to catch all sequential errors
            return [WindowScraper]::ScrapeAndClose($ProcessIdToScrape, 20)
        } -ArgumentList $excelPid, $scraperCode
        
        Write-Host "Opening workbook and checking for Ribbon UI errors..."
        # We MUST show Excel and enable alerts for the Ribbon UI error dialog to appear
        $testExcel.Visible = $true
        $testExcel.DisplayAlerts = $true
        
        $testWorkbook = $testExcel.Workbooks.Open($excelPath)
        
        # Re-hide and disable alerts for the rest of the testing
        $testExcel.Visible = $false
        $testExcel.DisplayAlerts = $false
        
        # Wait for the watcher to finish scraping
        $ribbonError = Receive-Job -Job $watcher -Wait
        Remove-Job $watcher
        Set-RibbonUiErrors -Enabled $false
        
        if ($ribbonError) {
            Write-Host "  ERROR: Ribbon UI Validation failed." -ForegroundColor Red
            # Clean up the output to make it readable for the AI agent
            $cleanError = $ribbonError -replace "\r\n+", " | " -replace "\s+", " "
            Write-Host "  [Diagnostics] $cleanError" -ForegroundColor Yellow
            throw "Ribbon UI Error: $cleanError"
        } else {
            Write-Host "  Ribbon UI loaded without errors." -ForegroundColor Green
        }
    } else {
        $testWorkbook = $testExcel.Workbooks.Open($excelPath)
    }

    Write-Host "Running internal unit tests..." -ForegroundColor Cyan
    try {
        $testExcel.Run("Lib_Tests.RunAllTests")
        Write-Host "  SUCCESS: Unit tests passed." -ForegroundColor Green
    } catch {
        Write-Host "  FAILURE: Unit tests failed." -ForegroundColor Red
        throw "Unit tests failed: $($_.Exception.Message)"
    }

    $structuredResults = Read-StructuredTestResults -Path $structuredTestResultsPath
    if ($null -ne $structuredResults -and $null -ne $structuredResults.Summary) {
        Write-Host ("  Structured test results: total={0}, passed={1}, failed={2}" -f $structuredResults.Summary.Total, $structuredResults.Summary.Passed, $structuredResults.Summary.Failed) -ForegroundColor Cyan
        foreach ($failedResult in @($structuredResults.Results | Where-Object { -not $_.Passed })) {
            Write-Host ("  [Test Failure] {0}: {1}" -f $failedResult.Name, $failedResult.Message) -ForegroundColor Yellow
        }
    } else {
        Write-Host "  Structured test results file was not produced." -ForegroundColor Yellow
    }

    $headlessCallbacks = Get-EnabledHeadlessCallbacks -ManifestPath $featureManifestPath -IncludeDev:$IncludeDevFeatures
    if ($headlessCallbacks.Count -gt 0) {
        Write-Host "Running headless-safe callback tests..." -ForegroundColor Cyan
        foreach ($callbackFeature in $headlessCallbacks) {
            Write-Host "  Testing callback: $($callbackFeature.OnAction)" -ForegroundColor Yellow
            $testExcel.Run($callbackFeature.OnAction, $null)
        }
    } else {
        Write-Host "No enabled headless-safe callbacks declared in features.json." -ForegroundColor Gray
    }

    Add-StageResult -Stage "runtime_tests" -Status "success" -Details "Ribbon, unit tests, and headless callback checks passed"
    Write-Host "Runtime testing completed with structured test collection." -ForegroundColor Green
} catch {
    Add-StageResult -Stage "runtime_tests" -Status "failure" -Details $_.Exception.Message
    Stop-Script "Test phase failed: $($_.Exception.Message)"
} finally {
    if ($testWorkbook) { $testWorkbook.Close($false) }
    if ($testExcel) { $testExcel.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-StageSummary
