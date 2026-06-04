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
    [switch]$SkipRuntimeTests
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$excelPath = Join-Path $PSScriptRoot "Beaver Add-in.xlsm"
$modulesDir = Join-Path $PSScriptRoot "Modules"
$desktopThisWorkbookCls = Join-Path $PSScriptRoot "ThisWorkbook.cls"
$ribbonXmlPath = Join-Path $PSScriptRoot "ribbon.xml"
$featureManifestPath = Join-Path $PSScriptRoot "features.json"
$testManifestPath = Join-Path $modulesDir "Libraries\Lib_TestManifest.bas"
$commandRegistryPath = Join-Path $modulesDir "Infrastructure\Infra_CommandRegistry.bas"
$uiRibbonPath = Join-Path $modulesDir "UI\UI_Ribbon.bas"
$uiHotkeysPath = Join-Path $modulesDir "UI\UI_Hotkeys.bas"
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
        [string]$Details = "",
        [double]$DurationMs = 0
    )

    [void]$script:StageResults.Add([pscustomobject]@{
        Stage = $Stage
        Status = $Status
        Details = $Details
        DurationMs = $DurationMs
        Timestamp = Get-Date
    })
}

function Write-StatusLine {
    param(
        [string]$Status,
        [string]$Stage,
        [string]$Details = "",
        [string]$Color = "Gray"
    )

    $detailText = if ([string]::IsNullOrWhiteSpace($Details)) { "" } else { " $Details" }
    Write-Host ("[{0,-6}] {1}{2}" -f $Status.ToUpperInvariant(), $Stage, $detailText) -ForegroundColor $Color
}

function Write-StageSummary {
    if ($script:StageResults.Count -eq 0) { return }

    Write-Host ""
    Write-Host "Stage Summary" -ForegroundColor Cyan
    foreach ($stage in $script:StageResults) {
        $color = if ($stage.Status -eq "success") { "Green" } elseif ($stage.Status -eq "skipped") { "Yellow" } else { "Red" }
        $durationText = if ($stage.DurationMs -gt 0) { " ({0:N1}s)" -f ($stage.DurationMs / 1000) } else { "" }
        $detailText = if ([string]::IsNullOrWhiteSpace($stage.Details)) { "" } else { " - $($stage.Details)" }
        Write-Host ("  [{0}] {1}{2}{3}" -f $stage.Status.ToUpper(), $stage.Stage, $durationText, $detailText) -ForegroundColor $color
    }

    $passed = @($script:StageResults | Where-Object { $_.Status -eq "success" }).Count
    $failed = @($script:StageResults | Where-Object { $_.Status -eq "failure" }).Count
    $skipped = @($script:StageResults | Where-Object { $_.Status -eq "skipped" }).Count
    $totalDurationMs = ($script:StageResults | Measure-Object -Property DurationMs -Sum).Sum
    if ($null -eq $totalDurationMs) { $totalDurationMs = 0 }

    Write-Host ""
    Write-Host ("Totals: passed={0}, failed={1}, skipped={2}, duration={3:N1}s" -f $passed, $failed, $skipped, ($totalDurationMs / 1000)) -ForegroundColor Cyan
}

function Invoke-Stage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Stage,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-StatusLine -Status "start" -Stage $Stage -Color "Cyan"

    try {
        $result = & $Action
        $stopwatch.Stop()

        $details = ""
        if ($null -ne $result) {
            if ($result -is [string]) {
                $details = $result
            } elseif ($result.PSObject.Properties.Name -contains "Details") {
                $details = [string]$result.Details
            }
        }

        Add-StageResult -Stage $Stage -Status "success" -Details $details -DurationMs $stopwatch.Elapsed.TotalMilliseconds
        Write-StatusLine -Status "pass" -Stage $Stage -Details ("({0:N1}s){1}" -f $stopwatch.Elapsed.TotalSeconds, $(if ([string]::IsNullOrWhiteSpace($details)) { "" } else { " $details" })) -Color "Green"
        return $result
    } catch {
        $stopwatch.Stop()
        $message = $_.Exception.Message
        Add-StageResult -Stage $Stage -Status "failure" -Details $message -DurationMs $stopwatch.Elapsed.TotalMilliseconds
        Write-StatusLine -Status "fail" -Stage $Stage -Details ("({0:N1}s) {1}" -f $stopwatch.Elapsed.TotalSeconds, $message) -Color "Red"
        throw
    }
}

function Add-SkippedStageResult {
    param(
        [string]$Stage,
        [string]$Details
    )

    Add-StageResult -Stage $Stage -Status "skipped" -Details $Details
    Write-StatusLine -Status "skip" -Stage $Stage -Details $Details -Color "Yellow"
}

function Release-ComObjectSafely {
    param([object]$ComObject)

    if ($null -eq $ComObject) {
        return
    }

    try {
        if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
    } catch {
        # Best-effort cleanup only.
    }
}

function Get-ExcelProcessId {
    param(
        [Parameter(Mandatory = $true)]
        $ExcelApplication
    )

    if ($null -eq $ExcelApplication -or -not $ExcelApplication.Hwnd) {
        return 0
    }

    $excelPid = 0
    [WindowScraper]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$excelPid) | Out-Null
    return $excelPid
}

function Reset-StructuredTestResults {
    param([string]$Path)

    if (Test-Path $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

function Write-StructuredTestResultsSummary {
    param([pscustomobject]$StructuredResults)

    Write-Host ("  Structured test results: total={0}, passed={1}, failed={2}" -f $StructuredResults.Summary.Total, $StructuredResults.Summary.Passed, $StructuredResults.Summary.Failed) -ForegroundColor Cyan

    foreach ($result in @($StructuredResults.Results)) {
        $resultColor = if ($result.Passed) { "Green" } else { "Yellow" }
        $resultStatus = if ($result.Passed) { "PASS" } else { "FAIL" }
        $messageText = if ([string]::IsNullOrWhiteSpace($result.Message)) { "" } else { " - $($result.Message)" }
        Write-Host ("    [{0}] {1} ({2} ms, {3}){4}" -f $resultStatus, $result.Name, $result.DurationMs, $result.Category, $messageText) -ForegroundColor $resultColor
    }
}

function Assert-StructuredTestResults {
    param(
        [pscustomobject]$StructuredResults,
        [string]$Path
    )

    if ($null -eq $StructuredResults -or $null -eq $StructuredResults.Summary) {
        throw "Structured test results file was not produced: $Path"
    }

    Write-StructuredTestResultsSummary -StructuredResults $StructuredResults

    if ($StructuredResults.Summary.Failed -gt 0) {
        $failedNames = @($StructuredResults.Results | Where-Object { -not $_.Passed } | ForEach-Object { $_.Name })
        $failedLabel = if ($failedNames.Count -gt 0) { $failedNames -join ", " } else { "unknown test(s)" }
        throw "Structured test results reported $($StructuredResults.Summary.Failed) failure(s): $failedLabel"
    }

    return "tests=$($StructuredResults.Summary.Total)"
}

function Invoke-HeadlessCallbackTests {
    param(
        $ExcelApplication,
        [object[]]$Callbacks
    )

    if ($null -eq $Callbacks -or $Callbacks.Count -eq 0) {
        Write-Host "No enabled headless-safe callbacks declared in features.json." -ForegroundColor Gray
        return "callbacks=0"
    }

    Write-Host "Running headless-safe callback tests..." -ForegroundColor Cyan
    $passed = 0

    foreach ($callbackFeature in $Callbacks) {
        $callbackName = [string]$callbackFeature.OnAction
        Write-Host "  Testing callback: $callbackName" -ForegroundColor Yellow
        $activeWindow = $null
        $originalFormulaBar = $null
        $originalHeadings = $null
        $originalWorkbookTabs = $null
        $originalHorizontalScrollBar = $null
        $originalVerticalScrollBar = $null

        try {
            $activeWindow = $ExcelApplication.ActiveWindow
            $originalFormulaBar = $ExcelApplication.DisplayFormulaBar

            if ($null -ne $activeWindow) {
                $originalHeadings = $activeWindow.DisplayHeadings
                $originalWorkbookTabs = $activeWindow.DisplayWorkbookTabs
                $originalHorizontalScrollBar = $activeWindow.DisplayHorizontalScrollBar
                $originalVerticalScrollBar = $activeWindow.DisplayVerticalScrollBar
            }

            $ExcelApplication.Run($callbackName, $null)
            Write-Host "    [PASS] $callbackName" -ForegroundColor Green
            $passed++
        } catch {
            Write-Host "    [FAIL] $callbackName - $($_.Exception.Message)" -ForegroundColor Red
            throw "Headless callback failed for '$callbackName': $($_.Exception.Message)"
        } finally {
            if ($null -ne $originalFormulaBar) {
                try { $ExcelApplication.DisplayFormulaBar = $originalFormulaBar } catch { }
            }

            if ($null -ne $activeWindow) {
                try {
                    if ($null -ne $originalHeadings) { $activeWindow.DisplayHeadings = $originalHeadings }
                    if ($null -ne $originalWorkbookTabs) { $activeWindow.DisplayWorkbookTabs = $originalWorkbookTabs }
                    if ($null -ne $originalHorizontalScrollBar) { $activeWindow.DisplayHorizontalScrollBar = $originalHorizontalScrollBar }
                    if ($null -ne $originalVerticalScrollBar) { $activeWindow.DisplayVerticalScrollBar = $originalVerticalScrollBar }
                } catch {
                    # Best-effort cleanup only.
                }
            }
        }
    }

    return "callbacks=$passed"
}

function Get-StructuredTestResultsDetails {
    param([pscustomobject]$StructuredResults)

    return "tests=$($StructuredResults.Summary.Total), passed=$($StructuredResults.Summary.Passed), failed=$($StructuredResults.Summary.Failed)"
}

function Get-FeatureManifest {
    param([string]$ManifestPath)
    if (-not (Test-Path $ManifestPath)) {
        throw "Feature manifest not found: $ManifestPath"
    }
    return Get-Content $ManifestPath -Raw | ConvertFrom-Json
}



function Sync-FeatureManifest {
    param(
        [string]$ManifestPath,
        [string]$ConfigPath,
        [string]$RibbonPath
    )

    Write-Host "Syncing feature manifest..." -ForegroundColor Cyan
    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath
    $config = if (Test-Path $ConfigPath) {
        Get-Content $ConfigPath -Raw | ConvertFrom-Json
    } else {
        [pscustomobject]@{}
    }

    $enabledFeatures = @($manifest.Features)
    $enabledFeatureIds = @($enabledFeatures | ForEach-Object { $_.ControlId })
    $enabledHotkeys = @($manifest.Hotkeys)

    $icons = [ordered]@{}
    foreach ($feature in $enabledFeatures) {
        $icons[$feature.ControlId] = $feature.Icon
    }

    $tabs = if ($null -ne $manifest.PSObject.Properties['Tabs']) {
        @($manifest.Tabs)
    } else {
        @($manifest.Tab)
    }

    $tabXmls = [System.Collections.Generic.List[string]]::new()
    foreach ($tab in $tabs) {
        $tabGroups = @()
        foreach ($group in $manifest.Groups) {
            $belongs = $false
            if ($null -ne $group.PSObject.Properties['TabId']) {
                $belongs = ($group.TabId -eq $tab.Id)
            } else {
                $belongs = ($tab.Id -eq $tabs[0].Id)
            }
            if ($belongs) {
                $tabGroups += $group
            }
        }

        $groupXmls = [System.Collections.Generic.List[string]]::new()
        foreach ($group in $tabGroups) {
            $groupFeatures = @($enabledFeatures | Where-Object { $group.Features -contains $_.ControlId })
            if ($groupFeatures.Count -eq 0) { continue }

            $buttonXml = foreach ($feature in $groupFeatures) {
                '          <button id="{0}" label="{1}" imageMso="{2}" size="large" onAction="{3}" keytip="{4}" screentip="{5}" supertip="{6}" />' -f `
                    $feature.ControlId,
                    [System.Security.SecurityElement]::Escape($feature.Label),
                    [System.Security.SecurityElement]::Escape($feature.Icon),
                    $feature.OnAction,
                    $feature.Keytip,
                    [System.Security.SecurityElement]::Escape($feature.Screentip),
                    [System.Security.SecurityElement]::Escape($feature.Supertip)
            }

            $groupXmls.Add(@"
        <group id="$($group.Id)" label="$([System.Security.SecurityElement]::Escape($group.Label))">
$($buttonXml -join "`r`n")
        </group>
"@)
        }

        $tabXmls.Add(@"
      <tab id="$($tab.Id)" label="$([System.Security.SecurityElement]::Escape($tab.Label))" keytip="$($tab.Keytip)">
$($groupXmls -join "`r`n")
      </tab>
"@)
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
$($tabXmls -join "`r`n")
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
    $moduleFiles = @(Get-ChildItem -Path $SourceDir -Filter *.bas -Recurse)

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

function Sync-CommandRegistry {
    param(
        [string]$ManifestPath,
        [string]$OutputPath
    )

    Write-Host "Generating command registry..." -ForegroundColor Cyan
    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath

    $entryMap = [ordered]@{}
    $commandMap = [ordered]@{}

    foreach ($feature in @($manifest.Features)) {
        if (-not [string]::IsNullOrWhiteSpace($feature.Macro) -and -not [string]::IsNullOrWhiteSpace($feature.CommandName)) {
            $entryMap[$feature.Macro.Trim().ToUpperInvariant()] = $feature.CommandName.Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($feature.CommandName)) {
            $commandName = $feature.CommandName.Trim()
            $commandClass = if ($feature.PSObject.Properties.Name -contains "CommandClass" -and -not [string]::IsNullOrWhiteSpace($feature.CommandClass)) { $feature.CommandClass.Trim() } else { "FeatCmd_$commandName" }
            $commandMap[$commandName.ToUpperInvariant()] = [pscustomobject]@{
                CommandName = $commandName
                CommandClass = $commandClass
            }
        }
    }

    foreach ($hotkey in @($manifest.Hotkeys)) {
        if (-not [string]::IsNullOrWhiteSpace($hotkey.Macro) -and -not [string]::IsNullOrWhiteSpace($hotkey.CommandName)) {
            $entryMap[$hotkey.Macro.Trim().ToUpperInvariant()] = $hotkey.CommandName.Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($hotkey.CommandName)) {
            $commandName = $hotkey.CommandName.Trim()
            $commandClass = if ($hotkey.PSObject.Properties.Name -contains "CommandClass" -and -not [string]::IsNullOrWhiteSpace($hotkey.CommandClass)) { $hotkey.CommandClass.Trim() } else { "FeatCmd_$commandName" }
            $commandMap[$commandName.ToUpperInvariant()] = [pscustomobject]@{
                CommandName = $commandName
                CommandClass = $commandClass
            }
        }
    }

    $lines = @(
        'Attribute VB_Name = "Infra_CommandRegistry"',
        'Option Explicit',
        '',
        ''' @Module: Infra_CommandRegistry',
        ''' @Category: Infrastructure',
        ''' @Description: Generated command registry mapping entry macros and command names to command classes.',
        ''' @ManagedBy: BeaverAddin Agent',
        ''' @Dependencies: ICommand',
        '',
        'Public Function ResolveCommandName(ByVal entryMacro As String) As String',
        '    Dim tracker As Object: Set tracker = Infra_Error.Track("ResolveCommandName")',
        '    On Error GoTo ErrHandler',
        '',
        '    Select Case UCase$(Trim$(entryMacro))'
    )

    foreach ($entry in $entryMap.GetEnumerator()) {
        $lines += ('        Case "{0}"' -f $entry.Key.Replace('"', '""'))
        $lines += ('            ResolveCommandName = "{0}"' -f $entry.Value.Replace('"', '""'))
    }

    $lines += '    End Select'
    $lines += ''
    $lines += 'CleanExit:'
    $lines += '    Exit Function'
    $lines += ''
    $lines += 'ErrHandler:'
    $lines += '    Infra_Error.HandleError "ResolveCommandName", Err'
    $lines += '    Resume CleanExit'
    $lines += 'End Function'
    $lines += ''
    $lines += 'Public Function CreateCommand(ByVal commandName As String) As ICommand'
    $lines += '    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateCommand")'
    $lines += '    On Error GoTo ErrHandler'
    $lines += ''
    $lines += '    Select Case UCase$(Trim$(commandName))'

    foreach ($entry in $commandMap.GetEnumerator()) {
        $lines += ('        Case "{0}"' -f $entry.Key.Replace('"', '""'))
        $lines += ('            Set CreateCommand = New {0}' -f $entry.Value.CommandClass)
    }

    $lines += '    End Select'
    $lines += ''
    $lines += 'CleanExit:'
    $lines += '    Exit Function'
    $lines += ''
    $lines += 'ErrHandler:'
    $lines += '    Infra_Error.HandleError "CreateCommand", Err'
    $lines += '    Resume CleanExit'
    $lines += 'End Function'

    [System.IO.File]::WriteAllText($OutputPath, ($lines -join "`r`n"), [System.Text.Encoding]::ASCII)
    Write-Host "  Command registry generated with $($commandMap.Count) command(s) and $($entryMap.Count) entry point(s)." -ForegroundColor Green
}

function Get-VbaProcedureNameFromMacro {
    param([string]$MacroName)

    if ([string]::IsNullOrWhiteSpace($MacroName)) {
        throw "Macro name is required."
    }

    $parts = $MacroName.Trim().Split('.')
    return $parts[$parts.Length - 1]
}

function Sync-UiRibbonModule {
    param(
        [string]$ManifestPath,
        [string]$OutputPath
    )

    Write-Host "Generating ribbon entry module..." -ForegroundColor Cyan
    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath

    $lines = @(
        'Attribute VB_Name = "UI_Ribbon"',
        'Option Explicit',
        '',
        ''' @Module: UI_Ribbon',
        ''' @Category: UI',
        ''' @Description: Generated Ribbon callbacks for the Beaver Add-in.',
        ''' @ManagedBy: BeaverAddin Agent',
        ''' @Dependencies: AppContainer, Infra_Config, Infra_Error',
        '',
        ''' --- Dynamic UI Callbacks ---',
        '',
        ''' Returns the image object for a control based on its ID in config.json',
        'Public Sub Ribbon_GetIcon(ByVal control As Object, ByRef image As Variant)',
        '    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_GetIcon")',
        '    On Error GoTo ErrHandler',
        '    ',
        '    Dim iconName As String',
        '    iconName = Infra_Config.GetIcon(control.Id)',
        '    If iconName = "" Then iconName = "Help"',
        '    ',
        '    Set image = Application.CommandBars.GetImageMso(iconName, 32, 32)',
        '    ',
        'CleanExit:',
        '    Exit Sub',
        'ErrHandler:',
        '    Infra_Error.HandleError "Ribbon_GetIcon", Err',
        '    Resume CleanExit',
        'End Sub'
    )

    foreach ($feature in @($manifest.Features)) {
        if ([string]::IsNullOrWhiteSpace($feature.OnAction) -or [string]::IsNullOrWhiteSpace($feature.Macro)) {
            continue
        }

        $procedureName = $feature.OnAction.Trim()
        $entryMacro = $feature.Macro.Trim()

        $lines += ''
        $lines += ('Public Sub {0}(ByVal control As Object)' -f $procedureName)
        $lines += ('    Dim tracker As Object: Set tracker = Infra_Error.Track("{0}")' -f $procedureName)
        $lines += '    On Error GoTo ErrHandler'
        $lines += ''
        $lines += ('    AppContainer.ExecuteEntryPoint "{0}", "{1}", "Ribbon"' -f $entryMacro.Replace('"', '""'), $procedureName.Replace('"', '""'))
        $lines += ''
        $lines += 'CleanExit:'
        $lines += '    Exit Sub'
        $lines += 'ErrHandler:'
        $lines += ('    Infra_Error.HandleError "{0}", Err' -f $procedureName)
        $lines += '    Resume CleanExit'
        $lines += 'End Sub'
    }

    [System.IO.File]::WriteAllText($OutputPath, ($lines -join "`r`n"), [System.Text.Encoding]::ASCII)
    Write-Host "  Ribbon entry module generated." -ForegroundColor Green
}

function Sync-UiHotkeysModule {
    param(
        [string]$ManifestPath,
        [string]$OutputPath
    )

    Write-Host "Generating hotkey entry module..." -ForegroundColor Cyan
    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath

    $lines = @(
        'Attribute VB_Name = "UI_Hotkeys"',
        'Option Explicit',
        '',
        ''' @Module: UI_Hotkeys',
        ''' @Category: UI',
        ''' @Description: Generated hotkey wrappers that route through the command pipeline.',
        ''' @ManagedBy: BeaverAddin Agent',
        ''' @Dependencies: AppContainer, Infra_Error'
    )

    foreach ($hotkey in @($manifest.Hotkeys)) {
        if ([string]::IsNullOrWhiteSpace($hotkey.Macro)) {
            continue
        }

        $macroName = $hotkey.Macro.Trim()
        $procedureName = Get-VbaProcedureNameFromMacro -MacroName $macroName

        $lines += ''
        $lines += ('Public Sub {0}()' -f $procedureName)
        $lines += ('    Dim tracker As Object: Set tracker = Infra_Error.Track("{0}")' -f $procedureName)
        $lines += '    On Error GoTo ErrHandler'
        $lines += ''
        $lines += ('    AppContainer.ExecuteEntryPoint "{0}", "{1}", "Hotkey"' -f $macroName.Replace('"', '""'), $procedureName.Replace('"', '""'))
        $lines += ''
        $lines += 'CleanExit:'
        $lines += '    Exit Sub'
        $lines += 'ErrHandler:'
        $lines += ('    Infra_Error.HandleError "{0}", Err' -f $procedureName)
        $lines += '    Resume CleanExit'
        $lines += 'End Sub'
    }

    [System.IO.File]::WriteAllText($OutputPath, ($lines -join "`r`n"), [System.Text.Encoding]::ASCII)
    Write-Host "  Hotkey entry module generated." -ForegroundColor Green
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
        $parts = $line -split "`t", 6
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
        [string]$ManifestPath
    )

    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath
    return @($manifest.Features | Where-Object {
        $_.RuntimeTestMode -eq "headless"
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
    $zip = $null
    try {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)
        
        # 1. Inject the Ribbon XML
        $ribbonEntryPath = "customUI/customUI14.xml"
        $ribbonEntry = $zip.GetEntry($ribbonEntryPath)
        if ($null -ne $ribbonEntry) { $ribbonEntry.Delete() }
        $ribbonEntry = $zip.CreateEntry($ribbonEntryPath)
        $writer = New-Object System.IO.StreamWriter($ribbonEntry.Open())
        $writer.Write((Get-Content $RibbonXmlPath -Raw))
        $writer.Close()

        # 2. Ensure relationship exists in _rels/.rels
        $relsEntryPath = "_rels/.rels"
        $relsEntry = $zip.GetEntry($relsEntryPath)
        if ($null -eq $relsEntry) {
            throw "_rels/.rels not found in workbook. Invalid Excel file structure."
        }

        $relsXml = [xml]""
        $stream = $relsEntry.Open()
        try {
            $reader = New-Object System.IO.StreamReader($stream)
            $relsXml = [xml]$reader.ReadToEnd()
        } finally {
            $stream.Close()
        }

        $nsMgr = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
        $nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")
        
        $relType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
        $existingRel = $relsXml.SelectSingleNode("//r:Relationship[@Target='$ribbonEntryPath']", $nsMgr)
        
        if ($null -eq $existingRel) {
            Write-Host "  Adding Custom UI relationship to _rels/.rels..."
            # Find max rId to create a new unique one
            $ids = $relsXml.SelectNodes("//r:Relationship/@Id", $nsMgr) | ForEach-Object { 
                if ($_.Value -match "rId(\d+)") { [int]$matches[1] } else { 0 }
            }
            $nextId = ($ids | Measure-Object -Maximum).Maximum + 1
            $newId = "rId$nextId"
            
            $root = $relsXml.DocumentElement
            $newRel = $relsXml.CreateElement("Relationship", $relsXml.DocumentElement.NamespaceURI)
            $newRel.SetAttribute("Id", $newId)
            $newRel.SetAttribute("Type", $relType)
            $newRel.SetAttribute("Target", $ribbonEntryPath)
            $root.AppendChild($newRel) | Out-Null
            
            # Save updated .rels
            $relsEntry.Delete()
            $relsEntry = $zip.CreateEntry($relsEntryPath)
            $writer = New-Object System.IO.StreamWriter($relsEntry.Open())
            $relsXml.Save($writer)
            $writer.Close()
        }

        $zip.Dispose()
        Write-Host "  Ribbon XML injected and registered successfully."
    } catch {
        throw "Failed to update Ribbon XML: $($_.Exception.Message)"
    } finally {
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
                $procLineNum = $i + 1
                
                # Skip common Excel events or very short helper functions if they don't need context
                if ($procName -match "^(?:Workbook_|Worksheet_|App_)" -or $file.Name -eq "Lib_JsonConverter.bas" -or $file.Name -match "^Lib_[a-zA-Z0-9_]+Function\.bas$" -or $file.Name -match "^(?:Infra_Error\.(bas|cls)|Infra_ContextTracker\.cls|Infra_Diagnostics\.bas|Infra_OperationContext\.cls|AppContainer\.cls|Infra_Config\.(cls|bas)|Infra_ConfigModel\.cls|I[A-Z][a-zA-Z0-9_\-]*\.cls|Infra_AppStateGuard\.cls|Infra_AppState\.bas)$") {
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
                    if ($lines[$j] -match "PushContext\s+""$procName""" -or $lines[$j] -match "Infra_Error\.Track") { $foundPush = $true }
                    if ($lines[$j] -match "PopContext" -or $lines[$j] -match "Infra_Error\.Track") { $foundPop = $true }
                    if ($lines[$j] -match "On Error GoTo\s+\w+") { $foundErrorGoto = $true }
                    if ($lines[$j] -match "HandleError\s+""$procName""") { $foundHandleError = $true }
                    $j++
                }

                if (-not $foundPush) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' at line $procLineNum missing context tracking (PushContext or Track)." -ForegroundColor Red
                    $allPassed = $false
                }
                if (-not $foundPop) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' at line $procLineNum missing 'PopContext' (or RAII Track tracker)." -ForegroundColor Red
                    $allPassed = $false
                }
                if (-not $foundErrorGoto) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' at line $procLineNum missing 'On Error GoTo'." -ForegroundColor Red
                    $allPassed = $false
                }
                if (-not $foundHandleError) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' at line $procLineNum missing 'HandleError ""$procName""'." -ForegroundColor Red
                    $allPassed = $false
                }
            }
        }
    }
    return $allPassed
}

# --- 1. PRE-DEPLOYMENT VALIDATION ---
$configPath = Join-Path $PSScriptRoot "config.json"

try {
    Invoke-Stage -Stage "manifest_sync" -Action {
        Sync-FeatureManifest -ManifestPath $featureManifestPath -ConfigPath $configPath -RibbonPath $ribbonXmlPath
        return "features synced from features.json"
    } | Out-Null

    Invoke-Stage -Stage "command_registry_generation" -Action {
        Sync-CommandRegistry -ManifestPath $featureManifestPath -OutputPath $commandRegistryPath
        return "command registry refreshed"
    } | Out-Null

    Invoke-Stage -Stage "ui_entry_generation" -Action {
        Sync-UiRibbonModule -ManifestPath $featureManifestPath -OutputPath $uiRibbonPath
        Sync-UiHotkeysModule -ManifestPath $featureManifestPath -OutputPath $uiHotkeysPath
        return "UI entry modules refreshed"
    } | Out-Null

    Invoke-Stage -Stage "test_manifest_generation" -Action {
        Sync-TestManifest -SourceDir $modulesDir -OutputPath $testManifestPath
        return "test manifest refreshed"
    } | Out-Null

    Invoke-Stage -Stage "validation" -Action {
        $validRibbon = Test-RibbonValidity -XmlPath $ribbonXmlPath -ModulesDir $modulesDir
        $validVba = Invoke-VbaSyntaxCheck -SourceDir $modulesDir
        $validLint = Invoke-EnhancedLinting -SourceDir $modulesDir

        if (-not ($validRibbon -and $validVba -and $validLint)) {
            throw "Pre-deployment validation failed"
        }

        return "ribbon, syntax, and lint checks passed"
    } | Out-Null

    # --- 2. ENVIRONMENT CHECKS ---
    Invoke-Stage -Stage "environment_checks" -Action {
        if (-not (Test-Path $excelPath)) {
            throw "Excel file not found: $excelPath"
        }

        $lockFile = Join-Path $PSScriptRoot ("~$" + (Split-Path $excelPath -Leaf))
        if (-not (Test-Path $lockFile)) {
            return "workbook available"
        }

        Write-Host "Excel file is open. Attempting to close it..." -ForegroundColor Yellow
        try {
            $activeExcel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            try {
                $activeExcel.DisplayAlerts = $false
                $wbFound = $null
                foreach ($wb in $activeExcel.Workbooks) {
                    if ($wb.FullName -eq $excelPath) {
                        $wbFound = $wb
                        break
                    }
                }

                if ($null -ne $wbFound) {
                    # Count other visible workbooks to decide if we can quit the Excel instance
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

                    if ($otherVisibleWorkbooks -eq 0) {
                        Write-Host "  No other visible workbooks open. Closing Excel application..." -ForegroundColor Green
                        $activeExcel.Quit()
                    }
                }
            } finally {
                try {
                    $activeExcel.DisplayAlerts = $true
                } catch { }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($activeExcel) | Out-Null
            }
        } catch {
            Write-Warning "Could not close gracefully via COM. Force closing Excel..."
            Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
        }

        Start-Sleep -Seconds 2
        if (Test-Path $lockFile) {
            throw "Excel file is still open. Please close it manually and retry."
        }

        return "lock file cleared"
    } | Out-Null

    # --- 3. BEGIN UPDATE ---
    Invoke-Stage -Stage "workbook_update" -Action {
        Write-Host "Starting Excel... (This may take a moment)"
        $excel = Start-ExcelApplication -Purpose "workbook update"
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $null
        $vbaProject = $null
        $btn = $null
        $activePane = $null
        $commandBars = $null
        $cb = $null

        try {
            $workbook = $excel.Workbooks.Open($excelPath)
            $vbaProject = $workbook.VBProject

            Write-Host "Updating modules..."
            $compsToRemove = @()
            for ($i = 1; $i -le $vbaProject.VBComponents.Count; $i++) {
                $comp = $vbaProject.VBComponents.Item($i)
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

            $vbaSourceFiles = Get-ChildItem -Path $modulesDir -Recurse | Where-Object { $_.Extension -match "\.(bas|cls|frm)$" }
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

            if (Test-Path $desktopThisWorkbookCls) {
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

            Write-Host "Compiling VBA Project..."
            if ($null -ne $excel.VBE) {
                Write-Host "  VBE object found."

                $commandBars = $excel.VBE.CommandBars
                foreach ($cb in $commandBars) {
                    function Find-ControlRecursive {
                        param($Parent)
                        foreach ($c in $Parent.Controls) {
                            try {
                                if ($c.Id -eq 578 -or $c.Caption -match "Compile") {
                                    return $c
                                }
                                if ($c.Type -eq 10 -or $c.Type -eq 12) {
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
                        $excel.DisplayAlerts = $false

                        try {
                            Write-Host "  Executing compile..."
                            $btn.Execute()
                        } catch {
                            Write-Host "  Execute() threw an exception: $($_.Exception.Message)"
                        }

                        if ($btn.Enabled) {
                            Write-Host "  ERROR: VBA Compilation failed (Button still enabled)." -ForegroundColor Red

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

                                    # Try to map the VBE line back to the actual disk file
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

                                        # Print context
                                        $contextStart = [Math]::Max(1, $diskErrorLine - 3)
                                        $contextEnd = [Math]::Min($diskLines.Count, $diskErrorLine + 3)

                                        Write-Host "  [Diagnostics] --- Code Context ---" -ForegroundColor Yellow
                                        for ($l = $contextStart; $l -le $contextEnd; $l++) {
                                            $prefix = if ($l -eq $diskErrorLine) { ">>> " } else { "    " }
                                            $color = if ($l -eq $diskErrorLine) { "Red" } else { "DarkGray" }
                                            Write-Host ("{0}{1:D3}: {2}" -f $prefix, $l, $diskLines[$l - 1]) -ForegroundColor $color
                                        }
                                        Write-Host "  [Diagnostics] --------------------" -ForegroundColor Yellow

                                        throw "VBA Compilation failed in module '$modName' (Source: $($diskFile.FullName)) at line $($diskErrorLine): '$errorLineText'. Please fix the syntax or missing definitions."
                                    } else {
                                        Write-Host "  [Diagnostics] Module: $modName" -ForegroundColor Yellow
                                        Write-Host "  [Diagnostics] Line $($startLine): $errorLineText" -ForegroundColor Yellow
                                        throw "VBA Compilation failed in module '$modName' at line $($startLine): '$errorLineText'. Please fix the syntax or missing definitions."
                                    }
                                }

                                Write-Host "  [Diagnostics] No ActiveCodePane found after failure." -ForegroundColor Yellow
                                throw "VBA Compilation failed. Check your code for syntax or definition errors."
                            } catch {
                                if ($_.Exception.Message -match "VBA Compilation failed") {
                                    throw $_.Exception.Message
                                }

                                Write-Host "  [Diagnostics] Error retrieving active pane: $($_.Exception.Message)" -ForegroundColor Red
                                throw "VBA Compilation failed. Check your code for 'Variable not defined' or syntax errors."
                            }
                        } else {
                            Write-Host "  Compilation successful." -ForegroundColor Green
                        }
                    } else {
                        Write-Host "  Project already compiled." -ForegroundColor Gray
                    }
                } else {
                    Write-Host "  'Compile Project' button NOT found. Listing available CommandBars:" -ForegroundColor Yellow
                    foreach ($cb in $commandBars) {
                        Write-Host "    - $($cb.Name) (Visible: $($cb.Visible))"
                    }
                }
            } else {
                Write-Host "  VBE object NOT found." -ForegroundColor Yellow
            }





            $workbook.Save()
            $workbook.Close($true)
            Release-ComObjectSafely $activePane
            Release-ComObjectSafely $btn
            Release-ComObjectSafely $cb
            Release-ComObjectSafely $commandBars
            Release-ComObjectSafely $vbaProject
            Release-ComObjectSafely $workbook
            $activePane = $null
            $btn = $null
            $cb = $null
            $commandBars = $null
            $vbaProject = $null
            $workbook = $null
            Write-Host "SUCCESS: Modules updated."
            return "modules imported and workbook saved"
        } finally {
            if ($workbook) {
                try { $workbook.Close($false) } catch { }
            }
            Release-ComObjectSafely $activePane
            Release-ComObjectSafely $btn
            Release-ComObjectSafely $cb
            Release-ComObjectSafely $commandBars
            Release-ComObjectSafely $vbaProject
            Release-ComObjectSafely $workbook
            if ($excel) {
                try { $excel.Quit() } catch { }
            }
            Release-ComObjectSafely $excel
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } | Out-Null

    Invoke-Stage -Stage "ribbon_injection" -Action {
        Update-RibbonInWorkbook -WorkbookPath $excelPath -RibbonXmlPath $ribbonXmlPath
        return "customUI14.xml refreshed"
    } | Out-Null

    # --- 4. RUNTIME TESTING ---
    if ($SkipRuntimeTests) {
        Add-SkippedStageResult -Stage "runtime_tests" -Details "Skipped by -SkipRuntimeTests"
        Write-StageSummary
        Write-Host "Skipping runtime testing (`-SkipRuntimeTests`)." -ForegroundColor Yellow
        exit 0
    }

    Invoke-Stage -Stage "runtime_tests" -Action {
        Write-Host "Starting Runtime Testing..." -ForegroundColor Cyan
        Reset-StructuredTestResults -Path $structuredTestResultsPath

        $testExcel = Start-ExcelApplication -Purpose "runtime testing"
        $testExcel.Visible = $false
        $testExcel.DisplayAlerts = $false
        $testWorkbook = $null
        $watcher = $null
        $ribbonUiErrorsEnabled = $false
        $testExcelPid = 0

        try {
            Set-RibbonUiErrors -Enabled $true
            $ribbonUiErrorsEnabled = $true

            $testExcelPid = Get-ExcelProcessId -ExcelApplication $testExcel

            if ($testExcelPid -gt 0) {
                $watcher = Start-Job -ScriptBlock {
                    param($ProcessIdToScrape, $code)
                    Add-Type -TypeDefinition $code -ErrorAction SilentlyContinue
                    return [WindowScraper]::ScrapeAndClose($ProcessIdToScrape, 20)
                } -ArgumentList $testExcelPid, $scraperCode

                Write-Host "Opening workbook and checking for Ribbon UI errors..."
                $testExcel.Visible = $true
                $testExcel.DisplayAlerts = $true

                $testWorkbook = $testExcel.Workbooks.Open($excelPath)

                $testExcel.Visible = $false
                $testExcel.DisplayAlerts = $false

                $ribbonError = Receive-Job -Job $watcher -Wait
                Remove-Job $watcher -Force
                $watcher = $null

                if ($ribbonError) {
                    Write-Host "  ERROR: Ribbon UI Validation failed." -ForegroundColor Red
                    $cleanError = $ribbonError -replace "\r\n+", " | " -replace "\s+", " "
                    Write-Host "  [Diagnostics] $cleanError" -ForegroundColor Yellow
                    throw "Ribbon UI Error: $cleanError"
                }

                Write-Host "  Ribbon UI loaded without errors." -ForegroundColor Green
            } else {
                $testWorkbook = $testExcel.Workbooks.Open($excelPath)
            }

            Write-Host "Running internal unit tests..." -ForegroundColor Cyan
            try {
                $testExcel.Run("Lib_Tests.RunAllTests")
                Write-Host "  SUCCESS: Unit tests completed." -ForegroundColor Green
            } catch {
                Write-Host "  FAILURE: Unit tests failed." -ForegroundColor Red
                
                # Retrieve and print detailed VBA diagnostic logs for the current process
                $vbaLogPath = Join-Path $env:TEMP "BeaverAddin_$testExcelPid.log"
                if (Test-Path $vbaLogPath) {
                    Write-Host "`n  --- BEAVER VBA DIAGNOSTIC LOGS ---" -ForegroundColor Yellow
                    Get-Content $vbaLogPath | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
                    Write-Host "  ----------------------------------`n" -ForegroundColor Yellow
                }
                throw "Unit tests failed: $($_.Exception.Message)"
            }

            $structuredResults = Read-StructuredTestResults -Path $structuredTestResultsPath
            try {
                Assert-StructuredTestResults -StructuredResults $structuredResults -Path $structuredTestResultsPath | Out-Null
            } catch {
                # Retrieve and print detailed VBA diagnostic logs for the current process
                $vbaLogPath = Join-Path $env:TEMP "BeaverAddin_$testExcelPid.log"
                if (Test-Path $vbaLogPath) {
                    Write-Host "`n  --- BEAVER VBA DIAGNOSTIC LOGS ---" -ForegroundColor Yellow
                    Get-Content $vbaLogPath | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
                    Write-Host "  ----------------------------------`n" -ForegroundColor Yellow
                }
                throw $_
            }

            $headlessCallbacks = Get-EnabledHeadlessCallbacks -ManifestPath $featureManifestPath -IncludeDev:$false
            Invoke-HeadlessCallbackTests -ExcelApplication $testExcel -Callbacks $headlessCallbacks | Out-Null

            Write-Host "Runtime testing completed with structured test collection." -ForegroundColor Green
            return (Get-StructuredTestResultsDetails -StructuredResults $structuredResults)
        } finally {
            if ($watcher) {
                Remove-Job $watcher -Force -ErrorAction SilentlyContinue
            }
            if ($ribbonUiErrorsEnabled) {
                Set-RibbonUiErrors -Enabled $false
            }
            if ($testWorkbook) {
                try { $testWorkbook.Close($false) } catch { }
            }
            Release-ComObjectSafely $testWorkbook
            if ($testExcel) {
                try { $testExcel.Quit() } catch { }
            }
            Release-ComObjectSafely $testExcel
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } | Out-Null

    Write-StageSummary
} catch {
    Stop-Script $_.Exception.Message
}
