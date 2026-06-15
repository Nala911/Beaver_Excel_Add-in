# Script:   Build.ps1
# Purpose:  Syncs VBA modules from disk, compiles the VBA project, and injects Ribbon XML.
# ==============================================================================

. (Join-Path $PSScriptRoot "BuildSupport.ps1")

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
    param ([string]$SourceDir)
    Write-Host "Linting VBA Files..." -ForegroundColor Cyan
    $vbaFiles = @(Get-ChildItem -Path $SourceDir -Include *.bas, *.cls, *.frm -Recurse)
    $thisWorkbook = Join-Path $PSScriptRoot "ThisWorkbook.cls"
    if (Test-Path $thisWorkbook) { $vbaFiles += Get-Item $thisWorkbook }

    $allPassed = $true
    foreach ($file in $vbaFiles) {
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

function Invoke-EnhancedLinting {
    param ([string]$SourceDir)
    Write-Host "Running Enhanced Linting..." -ForegroundColor Cyan
    $vbaFiles = @(Get-ChildItem -Path $SourceDir -Include *.bas, *.cls, *.frm -Recurse)
    $thisWorkbook = Join-Path $PSScriptRoot "ThisWorkbook.cls"
    if (Test-Path $thisWorkbook) { $vbaFiles += Get-Item $thisWorkbook }
    $allPassed = $true

    foreach ($file in $vbaFiles) {
        $content = [System.IO.File]::ReadAllText($file.FullName)
        $lines = $content -split "`r?`n"
        $fileName = $file.Name

        if ($content -notmatch "(?m)^Option Explicit") {
            Write-Host "  [$fileName] Error: Missing 'Option Explicit' at the top of the file." -ForegroundColor Red
            $allPassed = $false
        }

        if ($content -notmatch "' @Module:") {
            Write-Host "  [$fileName] Error: Missing '@Module' metadata header." -ForegroundColor Red
            $allPassed = $false
        }

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]

            # --- Rule A: Enforce Spill-Safe Formula Properties ---
            if ($file.Name -ne "Lib_JsonConverter.bas" -and $line -match '\b\.\bFormula\b' -and $line -notmatch '".*\.Formula.*"' -and $line -notmatch '^\s*\''' -and $line -notmatch '\.Formula2' -and $line -notmatch '\.FormulaArray') {
                Write-Host "  [$fileName] Error: Range.Formula usage detected at line $($i + 1). Use Range.Formula2 instead to prevent spill errors." -ForegroundColor Red
                $allPassed = $false
            }

            # --- Rule B: Multi-cell Range Property Null Check ---
            if ($line -match '\b(CStr|CInt|CLng|CDbl|CSng|CBool|CDate|CVar)\s*\(\s*(?!(?:cell\b|\w+\.Cells\b|\w+Cells\b))[a-zA-Z0-9_\.]+\.(?:NumberFormat|Font\.(?:Name|Size))\s*\)' -and $line -notmatch '^\s*''') {
                Write-Host "  [$fileName] Error: Direct string/value conversion on range property without IsNull check at line $($i + 1). Mixed ranges return Null, causing Error 94." -ForegroundColor Red
                $allPassed = $false
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
                    $allPassed = $false
                }
            }

            if ($line -match "^\s*Public (?:Sub|Function)\s+([a-zA-Z0-9_]+)") {
                $procName = $matches[1]
                $procLineNum = $i + 1
                
                if ($procName -match "^(?:Workbook_|Worksheet_|App_)" -or $file.Name -eq "Lib_JsonConverter.bas" -or $file.Name -match "^Lib_[a-zA-Z0-9_]+Function\.bas$" -or $file.Name -match "^(?:Infra_Error\.(bas|cls)|Infra_ContextTracker\.cls|Infra_Diagnostics\.bas|Infra_OperationContext\.cls|AppContainer\.cls|Infra_Config\.(cls|bas)|Infra_ConfigModel\.cls|I[A-Z][a-zA-Z0-9_\-]*\.cls|Infra_AppStateGuard\.cls|Infra_AppState\.bas)$") {
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

function Test-FormFilesValidity {
    param ([string]$SourceDir)
    Write-Host "Checking Form Companion Files..." -ForegroundColor Cyan
    $frmFiles = @(Get-ChildItem -Path $SourceDir -Include *.frm -Recurse)
    $allPassed = $true
    foreach ($file in $frmFiles) {
        $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
        if (-not (Test-Path $frxPath)) {
            Write-Host "  [$($file.Name)] Error: Missing companion binary file (.frx). MSForms requires a .frx file to import successfully." -ForegroundColor Red
            $allPassed = $false
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

# --- Build Execution ---

$sharedExcel = $null

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
        $validForms = Test-FormFilesValidity -SourceDir $modulesDir

        if (-not ($validRibbon -and $validVba -and $validLint -and $validForms)) {
            throw "Pre-deployment validation failed"
        }

        return "ribbon, syntax, lint, and form checks passed"
    } | Out-Null

    # --- environment_checks ---
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

    # --- workbook_update ---
    Invoke-Stage -Stage "workbook_update" -Action {
        if ($null -eq $sharedExcel) {
            Write-Host "Starting Excel... (This may take a moment)"
            $script:sharedExcel = Start-ExcelApplication -Purpose "workbook update"
        }
        $excel = $sharedExcel
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

            if ($null -ne $excel.VBE) {
                try {
                    $excel.VBE.MainWindow.Visible = $false
                } catch { }
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
            $activePane = $null
            $btn = $null
            $cb = $null
            $commandBars = $null
            $vbaProject = $null
            $workbook = $null
        }
    } | Out-Null

    Invoke-Stage -Stage "ribbon_injection" -Action {
        Update-RibbonInWorkbook -WorkbookPath $excelPath -RibbonXmlPath $ribbonXmlPath
        return "customUI14.xml refreshed"
    } | Out-Null

    Write-StageSummary
} catch {
    Stop-Script $_.Exception.Message
} finally {
    if ($null -ne $sharedExcel) {
        Write-Host "Closing Excel application..." -ForegroundColor Gray
        try { $sharedExcel.Quit() } catch { }
        Release-ComObjectSafely $sharedExcel
        $sharedExcel = $null
    }
}
