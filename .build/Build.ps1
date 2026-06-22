[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$SkipLint,
    [switch]$Clean
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "BuildSupport.ps1")

# Initialize global caches to satisfy StrictMode
$global:BeaverLintStatusCache = @{}
$global:BeaverTestManifestCache = @{}

# --- Helper Functions ---

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
                if ($null -ne $feature.PSObject.Properties['Type'] -and $feature.Type -eq "Menu") {
                    $menuItemsXml = [System.Collections.Generic.List[string]]::new()
                    foreach ($subId in $feature.MenuItems) {
                        $subFeature = $enabledFeatures | Where-Object { $_.ControlId -eq $subId }
                        if ($null -ne $subFeature) {
                            $menuItemsXml.Add(('            <button id="{0}" label="{1}" imageMso="{2}" onAction="{3}" keytip="{4}" screentip="{5}" supertip="{6}" />' -f `
                                $subFeature.ControlId,
                                [System.Security.SecurityElement]::Escape($subFeature.Label),
                                [System.Security.SecurityElement]::Escape($subFeature.Icon),
                                $subFeature.OnAction,
                                $subFeature.Keytip,
                                [System.Security.SecurityElement]::Escape($subFeature.Screentip),
                                [System.Security.SecurityElement]::Escape($subFeature.Supertip)))
                        }
                    }
                    "          <menu id=`"{0}`" label=`"{1}`" imageMso=`"{2}`" size=`"large`" keytip=`"{3}`" screentip=`"{4}`" supertip=`"{5}`">`r`n{6}`r`n          </menu>" -f `
                        $feature.ControlId,
                        [System.Security.SecurityElement]::Escape($feature.Label),
                        [System.Security.SecurityElement]::Escape($feature.Icon),
                        $feature.Keytip,
                        [System.Security.SecurityElement]::Escape($feature.Screentip),
                        [System.Security.SecurityElement]::Escape($feature.Supertip),
                        ($menuItemsXml -join "`r`n")
                } else {
                    '          <button id="{0}" label="{1}" imageMso="{2}" size="large" onAction="{3}" keytip="{4}" screentip="{5}" supertip="{6}" />' -f `
                        $feature.ControlId,
                        [System.Security.SecurityElement]::Escape($feature.Label),
                        [System.Security.SecurityElement]::Escape($feature.Icon),
                        $feature.OnAction,
                        $feature.Keytip,
                        [System.Security.SecurityElement]::Escape($feature.Screentip),
                        [System.Security.SecurityElement]::Escape($feature.Supertip)
                }
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
    $testProcedures = Get-AllTestProcedures -SourceDir $SourceDir

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
    $lines += 'Public Sub RunGeneratedTests(Optional ByVal filterPattern As String = "")'
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
            $testFullName = "$($test.Module).$($test.Procedure)"
            $lines += "    If MatchesFilter(""{0}"", filterPattern) Then {0}" -f $testFullName
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
    $lines += ''
    $lines += 'Private Function MatchesFilter(ByVal testName As String, ByVal filterPattern As String) As Boolean'
    $lines += '    If filterPattern = "" Then'
    $lines += '        MatchesFilter = True'
    $lines += '        Exit Function'
    $lines += '    End If'
    $lines += '    Dim patterns() As String'
    $lines += '    patterns = Split(filterPattern, ",")'
    $lines += '    Dim i As Long'
    $lines += '    For i = LBound(patterns) To UBound(patterns)'
    $lines += '        If UCase$(testName) Like UCase$(Trim$(patterns(i))) Then'
    $lines += '            MatchesFilter = True'
    $lines += '            Exit Function'
    $lines += '        End If'
    $lines += '    Next i'
    $lines += '    MatchesFilter = False'
    $lines += 'End Function'
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
        $hasMacro = $feature.PSObject.Properties.Name -contains "Macro"
        $hasCommandName = $feature.PSObject.Properties.Name -contains "CommandName"
        if ($hasMacro -and $hasCommandName -and -not [string]::IsNullOrWhiteSpace($feature.Macro) -and -not [string]::IsNullOrWhiteSpace($feature.CommandName)) {
            $entryMap[$feature.Macro.Trim().ToUpperInvariant()] = $feature.CommandName.Trim()
        }
        if ($hasCommandName -and -not [string]::IsNullOrWhiteSpace($feature.CommandName)) {
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

function Sync-HelpManifest {
    param(
        [string]$ManifestPath,
        [string]$OutputPath
    )

    Write-Host "Generating help manifest..." -ForegroundColor Cyan
    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath

    $lines = @(
        'Attribute VB_Name = "Lib_HelpManifest"',
        'Option Explicit',
        'Option Private Module',
        '',
        ''' @Module: Lib_HelpManifest',
        ''' @Category: Library',
        ''' @Description: Generated help manifest containing ribbon feature descriptions for dynamic help display.',
        ''' @ManagedBy: BeaverAddin Agent',
        ''' @Dependencies: Infra_Error',
        '',
        'Public Function GetFeatureHelp() As Collection',
        '    Dim tracker As Object: Set tracker = Infra_Error.Track("GetFeatureHelp")',
        '    On Error GoTo ErrHandler',
        '    ',
        '    Dim col As New Collection',
        '    Dim dict As Object',
        ''
    )

    foreach ($feature in @($manifest.Features)) {
        $label = if ($feature.PSObject.Properties.Name -contains "Label") { $feature.Label } else { "" }
        $screentip = if ($feature.PSObject.Properties.Name -contains "Screentip") { $feature.Screentip } else { "" }
        $supertip = if ($feature.PSObject.Properties.Name -contains "Supertip") { $feature.Supertip } else { "" }
        
        $labelEsc = if ($label) { $label.Replace('"', '""') } else { "" }
        $screentipEsc = if ($screentip) { $screentip.Replace('"', '""') } else { "" }
        $supertipEsc = if ($supertip) { $supertip.Replace('"', '""') } else { "" }

        if (-not [string]::IsNullOrWhiteSpace($labelEsc)) {
            $lines += '    Set dict = CreateObject("Scripting.Dictionary")'
            $lines += '    dict.Add "Label", "{0}"' -f $labelEsc
            $lines += '    dict.Add "Screentip", "{0}"' -f $screentipEsc
            $lines += '    dict.Add "Supertip", "{0}"' -f $supertipEsc
            $lines += '    col.Add dict'
            $lines += ''
        }
    }

    $lines += @(
        '    Set GetFeatureHelp = col',
        '',
        'CleanExit:',
        '    Exit Function',
        'ErrHandler:',
        '    Infra_Error.HandleError "GetFeatureHelp", Err',
        '    Resume CleanExit',
        'End Function'
    )

    [System.IO.File]::WriteAllText($OutputPath, ($lines -join "`r`n"), [System.Text.Encoding]::ASCII)
    Write-Host "  Help manifest generated with $($manifest.Features.Count) feature(s)." -ForegroundColor Green
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
        $hasOnAction = $feature.PSObject.Properties.Name -contains "OnAction"
        $hasMacro = $feature.PSObject.Properties.Name -contains "Macro"
        if (-not $hasOnAction -or -not $hasMacro -or [string]::IsNullOrWhiteSpace($feature.OnAction) -or [string]::IsNullOrWhiteSpace($feature.Macro)) {
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

function Test-RibbonValidity {
    param (
        [string]$XmlPath,
        [string]$ModulesDir
    )

    if (-not (Test-Path $XmlPath)) { return $true }

    Write-Host "Validating Ribbon XML..." -ForegroundColor Cyan
    $isValid = $true
    $absoluteXmlPath = Resolve-Path $XmlPath

    try {
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.XmlResolver = $null
        $settings.ValidationType = [System.Xml.ValidationType]::Schema
        $settings.ValidationFlags = $settings.ValidationFlags -bor [System.Xml.Schema.XmlSchemaValidationFlags]::ProcessIdentityConstraints
        $settings.ValidationFlags = $settings.ValidationFlags -bor [System.Xml.Schema.XmlSchemaValidationFlags]::ReportValidationWarnings

        $onValidationError = [System.Xml.Schema.ValidationEventHandler] {
            param($evtSource, $e)
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

    $xml = [xml](Get-Content $XmlPath -Raw)
    
    $ids = $xml.SelectNodes("//@id") | ForEach-Object { $_.Value }
    $duplicates = $ids | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($duplicates) {
        Write-Error "Duplicate IDs found in ribbon.xml: $($duplicates.Name -join ', ')"
        $isValid = $false
    }

    $callbacks = $xml.SelectNodes("//@onAction") | ForEach-Object { $_.Value } | Select-Object -Unique
    if ($callbacks) {
        Write-Host "  Checking $($callbacks.Count) callbacks across all modules..."
        $vbaFiles = Get-ChildItem -Path $ModulesDir -Include *.bas, *.cls -Recurse
        $sb = New-Object System.Text.StringBuilder
        foreach ($f in $vbaFiles) {
            [void]$sb.AppendLine([System.IO.File]::ReadAllText($f.FullName))
        }
        $vbaCode = $sb.ToString()
        
        foreach ($cb in $callbacks) {
            if ($vbaCode -notmatch "Sub\s+$cb\s*\(") {
                Write-Error "Ribbon callback '$cb' not found in any module in $ModulesDir"
                $isValid = $false
            }
        }
    }

    return $isValid
}

function Update-RibbonInWorkbook {
    param ([string]$WorkbookPath, [string]$RibbonXmlPath)
    if (-not (Test-Path $RibbonXmlPath)) { return }
    Write-Host "Injecting Ribbon XML..."
    $zip = $null
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

function Invoke-VbaSyntaxCheck {
    param (
        [string]$SourceDir,
        [string[]]$FilesToProcess
    )
    Write-Host "Linting VBA Files..." -ForegroundColor Cyan

    $projectRoot = Split-Path $PSScriptRoot -Parent
    $buildState = Get-BuildState
    if (-not (Get-Variable -Name "BeaverLintStatusCache" -Scope Global -ErrorAction SilentlyContinue)) { $global:BeaverLintStatusCache = @{} }

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

    $allPassed = $true
    foreach ($file in $vbaFiles) {
        $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace("\", "/")
        if ($null -ne $global:BeaverBuildLog) {
            if ($global:BeaverBuildLog.lintResults.checkedFiles -notcontains $relPath) {
                [void]$global:BeaverBuildLog.lintResults.checkedFiles.Add($relPath)
            }
        }
        
        $cachedPassed = $false
        if ($null -ne $buildState -and $null -ne $buildState.Metadata -and $null -ne $buildState.Metadata.PSObject.Properties[$relPath]) {
            $meta = $buildState.Metadata.PSObject.Properties[$relPath].Value
            if ($meta.Length -eq $file.Length -and $meta.LastWriteTime -eq $file.LastWriteTime.ToFileTime().ToString() -and $meta.LintPassed -eq $true) {
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

        $rawLines = [System.IO.File]::ReadAllLines($file.FullName)
        $fileName = $file.Name
        
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
                        Add-LintError -File $fileName -Type "syntax" -Message "Unexpected '$($b.End.Trim())' (No matching start found)" -Line $lineNum
                        $allPassed = $false
                        $global:BeaverLintStatusCache[$relPath] = $false
                    }
                }
            }
            
            foreach ($startLine in $stack) {
                Write-Host "  [$fileName] Syntax Error: Mismatched '$($b.Name)' starting at line $startLine (No matching end found)." -ForegroundColor Red
                Add-LintError -File $fileName -Type "syntax" -Message "Mismatched '$($b.Name)' (No matching end found)" -Line $startLine
                $allPassed = $false
                $global:BeaverLintStatusCache[$relPath] = $false
            }
        }
    }
    return $allPassed
}

function Invoke-EnhancedLinting {
    param (
        [string]$SourceDir,
        [string[]]$FilesToProcess
    )
    Write-Host "Running Enhanced Linting..." -ForegroundColor Cyan

    $projectRoot = Split-Path $PSScriptRoot -Parent
    $buildState = Get-BuildState
    if (-not (Get-Variable -Name "BeaverLintStatusCache" -Scope Global -ErrorAction SilentlyContinue)) { $global:BeaverLintStatusCache = @{} }

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
    $allPassed = $true

    foreach ($file in $vbaFiles) {
        $relPath = $file.FullName.Substring($projectRoot.Length + 1).Replace("\", "/")
        if ($null -ne $global:BeaverBuildLog) {
            if ($global:BeaverBuildLog.lintResults.checkedFiles -notcontains $relPath) {
                [void]$global:BeaverBuildLog.lintResults.checkedFiles.Add($relPath)
            }
        }
        
        $cachedPassed = $false
        if ($null -ne $buildState -and $null -ne $buildState.Metadata -and $null -ne $buildState.Metadata.PSObject.Properties[$relPath]) {
            $meta = $buildState.Metadata.PSObject.Properties[$relPath].Value
            if ($meta.Length -eq $file.Length -and $meta.LastWriteTime -eq $file.LastWriteTime.ToFileTime().ToString() -and $meta.LintPassed -eq $true) {
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

        $content = [System.IO.File]::ReadAllText($file.FullName)
        $lines = $content -split "`r?`n"
        $fileName = $file.Name

        if ($content -notmatch "(?m)^Option Explicit") {
            Write-Host "  [$fileName] Error: Missing 'Option Explicit' at the top of the file." -ForegroundColor Red
            Add-LintError -File $fileName -Type "enhanced" -Message "Missing 'Option Explicit' at the top of the file." -Line 1
            $allPassed = $false
            $global:BeaverLintStatusCache[$relPath] = $false
        }

        if ($file.Name -ne "Lib_TestManifest.bas" -and $content -notmatch "' @Module:") {
            Write-Host "  [$fileName] Error: Missing '@Module' metadata header." -ForegroundColor Red
            Add-LintError -File $fileName -Type "enhanced" -Message "Missing '@Module' metadata header." -Line 1
            $allPassed = $false
            $global:BeaverLintStatusCache[$relPath] = $false
        }

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]

            # --- Rule A: Enforce Spill-Safe Formula Properties ---
            if ($file.Name -ne "Lib_JsonConverter.bas" -and $line -match '\.\bFormula\b' -and $line -notmatch '".*\.Formula.*"' -and $line -notmatch '^\s*\''' -and $line -notmatch '\.Formula2' -and $line -notmatch '\.FormulaArray') {
                Write-Host "  [$fileName] Error: Range.Formula usage detected at line $($i + 1). Use Range.Formula2 instead to prevent spill errors." -ForegroundColor Red
                Add-LintError -File $fileName -Type "enhanced" -Message "Range.Formula usage detected. Use Range.Formula2 instead to prevent spill errors." -Line ($i + 1)
                $allPassed = $false
                $global:BeaverLintStatusCache[$relPath] = $false
            }

            # --- Rule B: Multi-cell Range Property Null Check ---
            if ($line -match '\b(CStr|CInt|CLng|CDbl|CSng|CBool|CDate|CVar)\s*\(\s*(?!(?:cell\b|\w+\.Cells\b|\w+Cells\b))[a-zA-Z0-9_\.]+\.(?:NumberFormat|Font\.(?:Name|Size))\s*\)' -and $line -notmatch '^\s*''') {
                Write-Host "  [$fileName] Error: Direct string/value conversion on range property without IsNull check at line $($i + 1). Mixed ranges return Null, causing Error 94." -ForegroundColor Red
                Add-LintError -File $fileName -Type "enhanced" -Message "Direct string/value conversion on range property without IsNull check. Mixed ranges return Null, causing Error 94." -Line ($i + 1)
                $allPassed = $false
                $global:BeaverLintStatusCache[$relPath] = $false
            }

            # --- Rule C: Collection Mutation Loop Direction ---
            if ($line -match '\bFor\s+(\w+)\s*=\s*\d+\s+To\s+(?:(?:\w+\.)*Count|\w+)\b(?!\s+Step\s+-1)' -and $line -notmatch '^\s*''') {
                $idxVar = $Matches[1]
                $j = $i + 1
                $hasDeletion = $false
                while ($j -lt $lines.Count -and $lines[$j] -notmatch "\bNext\b") {
                    if ($lines[$j] -match "\b$idxVar\b" -and ($lines[$j] -match '\bDelete\b|\bRemove\b|\bRemoveAt\b') -and $lines[$j] -notmatch '^\s*''') {
                        $hasDeletion = $true
                        break
                    }
                    $j++
                }
                if ($hasDeletion) {
                    Write-Host "  [$fileName] Error: Forward iteration loop with mutation detected at line $($i + 1). Use backward iteration 'For $idxVar = ... To 1 Step -1' instead to prevent skipping bugs." -ForegroundColor Red
                    Add-LintError -File $fileName -Type "enhanced" -Message "Forward iteration loop with mutation detected. Use backward iteration 'For $idxVar = ... To 1 Step -1' instead to prevent skipping bugs." -Line ($i + 1)
                    $allPassed = $false
                    $global:BeaverLintStatusCache[$relPath] = $false
                }
            }

            if ($line -match "^\s*Public (?:Sub|Function)\s+([a-zA-Z0-9_]+)") {
                $procName = $matches[1]
                $procLineNum = $i + 1
                
                if ($procName -match "^(?:Workbook_|Worksheet_|App_)" -or $file.Name -eq "Lib_JsonConverter.bas" -or $file.Name -match "^Lib_[a-zA-Z0-9_]+Function\.bas$" -or $file.Name -match "^(?:Infra_Error\.(bas|cls)|Infra_ContextTracker\.cls|Infra_Diagnostics\.bas|Infra_OperationContext\.cls|AppContainer\.cls|Infra_Config\.(cls|bas)|Infra_ConfigModel\.cls|I[A-Z][a-zA-Z0-9_\-]*\.cls|Infra_AppStateGuard\.cls|Infra_AppState\.bas|Infra_ValueConversion\.bas)$") {
                    continue
                }

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
                    Add-LintError -File $fileName -Type "enhanced" -Message "Procedure '$procName' missing context tracking (PushContext or Track)." -Line $procLineNum
                    $allPassed = $false
                    $global:BeaverLintStatusCache[$relPath] = $false
                }
                if (-not $foundPop) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' at line $procLineNum missing 'PopContext' (or RAII Track tracker)." -ForegroundColor Red
                    Add-LintError -File $fileName -Type "enhanced" -Message "Procedure '$procName' missing 'PopContext' (or RAII Track tracker)." -Line $procLineNum
                    $allPassed = $false
                    $global:BeaverLintStatusCache[$relPath] = $false
                }
                if (-not $foundErrorGoto) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' at line $procLineNum missing 'On Error GoTo'." -ForegroundColor Red
                    Add-LintError -File $fileName -Type "enhanced" -Message "Procedure '$procName' missing 'On Error GoTo'." -Line $procLineNum
                    $allPassed = $false
                    $global:BeaverLintStatusCache[$relPath] = $false
                }
                if (-not $foundHandleError) {
                    Write-Host "  [$fileName] Error: Procedure '$procName' at line $procLineNum missing 'HandleError ""$procName""'." -ForegroundColor Red
                    Add-LintError -File $fileName -Type "enhanced" -Message "Procedure '$procName' missing 'HandleError ""$procName""'." -Line $procLineNum
                    $allPassed = $false
                    $global:BeaverLintStatusCache[$relPath] = $false
                }
            }
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
            if ($meta.Length -eq $file.Length -and $meta.LastWriteTime -eq $file.LastWriteTime.ToFileTime().ToString() -and $meta.LintPassed -eq $true) {
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

function Get-VbaComponentNameFromFile {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) { return $null }

    $lines = Get-Content -Path $FilePath -TotalCount 50
    foreach ($line in $lines) {
        if ($line -match '^Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
            return $Matches[1]
        }
    }
    
    return [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
}

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
        $manifestChanged = (-not $buildState.Files.PSObject.Properties.Name.Contains("features.json") -or $buildState.Files."features.json" -ne $currentHashes["features.json"])
        
        foreach ($key in $currentHashes.Keys) {
            if (-not $buildState.Files.PSObject.Properties.Name.Contains($key) -or $buildState.Files.$key -ne $currentHashes[$key]) {
                $changedFiles += $key
            }
        }
        
        foreach ($prop in $buildState.Files.PSObject.Properties) {
            $key = $prop.Name
            if (-not $currentHashes.ContainsKey($key)) {
                $deletedFiles += $key
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
        $registryMissing = -not (Test-Path $commandRegistryPath)
        $helpMissing = -not (Test-Path $helpManifestPath)
        
        if ($forceFullBuild -or $manifestStructureChanged -or $registryMissing -or $helpMissing) {
            if ($forceFullBuild -or $manifestStructureChanged -or $registryMissing) {
                Sync-CommandRegistry -ManifestPath $featureManifestPath -OutputPath $commandRegistryPath
            }
            if ($forceFullBuild -or $manifestStructureChanged -or $helpMissing) {
                Sync-HelpManifest -ManifestPath $featureManifestPath -OutputPath $helpManifestPath
            }
            
            if (-not $forceFullBuild) {
                foreach ($relPath in @("Modules/Infrastructure/Infra_CommandRegistry.bas", "Modules/Libraries/Lib_HelpManifest.bas")) {
                    if ($changedFiles -notcontains $relPath) {
                        $script:changedFiles += $relPath
                    }
                }
            }
            return "command registry refreshed"
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
                        $script:changedFiles += $relPath
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
                $relPath = "Modules/Libraries/Lib_TestManifest.bas"
                if ($changedFiles -notcontains $relPath) {
                    $script:changedFiles += $relPath
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
        $script:sharedExcel = $excel
        
        $vbaProject = $workbook.VBProject
        $btn = $null
        $activePane = $null
        $commandBars = $null
        $cb = $null

        try {
            Write-Host "Updating modules..."
            
            if ($forceFullBuild) {
                # Purge all components
                $compsToRemoveList = @()
                foreach ($comp in $vbaProject.VBComponents) {
                    if (($comp.Type -ge 1 -and $comp.Type -le 3) -and ($comp.Name -ne "ThisWorkbook")) {
                        $compsToRemoveList += $comp
                    }
                }
                $vbaSourceFiles = Get-ChildItem -Path $modulesDir -Include *.bas, *.cls, *.frm -Recurse
                $filesToImport = @($vbaSourceFiles | ForEach-Object { $_.FullName })
                $compsToRemove = $compsToRemoveList
            } else {
                # Incremental Mode: convert component names to VBComponent COM objects
                $compsToRemoveList = @()
                foreach ($compName in $compsToRemove) {
                    try {
                        $comp = $vbaProject.VBComponents.Item($compName)
                        if ($null -ne $comp) {
                            $compsToRemoveList += $comp
                        }
                    } catch {
                        # Component doesn't exist
                    }
                }
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
                }
            }

            $tempImportDir = Join-Path ([System.IO.Path]::GetTempPath()) ("BeaverAddin-VbaImport-" + [System.Guid]::NewGuid().ToString("N"))

            try {
                foreach ($filePath in $filesToImport) {
                    $fileName = Split-Path $filePath -Leaf
                    Write-Host "  Importing $($fileName)..."
                    $importPath = New-NormalizedImportCopy -SourcePath $filePath -TempRoot $tempImportDir

                    if ($filePath -match "\.frm$") {
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

            $workbook.Saved = $false
            $workbook.Save()
            
            if ($forceFullBuild -or $manifestChanged) {
                Write-Host "  Closing workbook to release file lock for Ribbon XML injection..." -ForegroundColor Yellow
                $workbook.Close($true)
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
            Write-Host "SUCCESS: Modules updated."
            return "modules imported and workbook saved"
        } finally {
            if ($forceFullBuild -or $manifestChanged) {
                if ($workbook) {
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
        if (-not $excelWasAlreadyOpen) {
            try {
                foreach ($wb in $sharedExcel.Workbooks) {
                    if ($wb.FullName -eq $excelPath) {
                        $wb.Close($true)
                    }
                }
            } catch { }
            try {
                $sharedExcel.Quit()
            } catch { }
        } else {
            try {
                $sharedExcel.Visible = $true
                $sharedExcel.DisplayAlerts = $true
            } catch { }
        }
        Release-ComObjectSafely $sharedExcel
        $sharedExcel = $null
    }
}
