# Script:   Generators.ps1
# Purpose:  VBA code and manifest generation helpers for the build pipeline.
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
    $null = Write-FileIfChanged -Path $RibbonPath -Content $ribbonContent

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
    $null = Write-FileIfChanged -Path $ConfigPath -Content $configJson
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
        'Attribute VB_Name = "Test_Manifest"',
        'Option Explicit',
        '',
        ''' @Module: Test_Manifest',
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
    $null = Write-FileIfChanged -Path $OutputPath -Content ($lines -join "`r`n")
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

    $null = Write-FileIfChanged -Path $OutputPath -Content ($lines -join "`r`n")
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

    $null = Write-FileIfChanged -Path $OutputPath -Content ($lines -join "`r`n")
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

    $null = Write-FileIfChanged -Path $OutputPath -Content ($lines -join "`r`n")
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

    $null = Write-FileIfChanged -Path $OutputPath -Content ($lines -join "`r`n")
    Write-Host "  Hotkey entry module generated." -ForegroundColor Green
}

function Sync-UdfRegistry {
    param(
        [string]$ManifestPath,
        [string]$OutputPath
    )

    Write-Host "Generating UDF registry..." -ForegroundColor Cyan
    $manifest = Get-FeatureManifest -ManifestPath $ManifestPath
    
    $udfs = @()
    if ($null -ne $manifest -and $null -ne $manifest.UDFs) {
        $udfs = @($manifest.UDFs)
    }

    $lines = @(
        'Attribute VB_Name = "Lib_UdfRegistry"',
        'Option Explicit',
        'Option Private Module',
        '',
        ''' @Module: Lib_UdfRegistry',
        ''' @Category: Library',
        ''' @Description: Generated registry of User Defined Functions (UDFs) metadata for Beaver Add-in.',
        ''' @ManagedBy: BeaverAddin Agent',
        ''' @Dependencies: Infra_Error',
        '',
        ''' Returns a collection of metadata dictionaries for all registered UDFs.',
        ''' Each dictionary contains:',
        '''   - Name: String',
        '''   - Description: String',
        '''   - Category: String',
        '''   - Syntax: String',
        '''   - ArgumentDescriptions: Variant Array of Strings',
        'Public Function GetAllUdfs() As Collection',
        '    Dim tracker As Object: Set tracker = Infra_Error.Track("GetAllUdfs")',
        '    On Error GoTo ErrHandler',
        '',
        '    Dim registry As New Collection',
        ''
    )

    foreach ($udf in $udfs) {
        $lines += '    registry.Add Get{0}Metadata()' -f $udf.Name
    }

    $lines += @(
        '',
        '    Set GetAllUdfs = registry',
        '',
        'CleanExit:',
        '    Exit Function',
        '',
        'ErrHandler:',
        '    Infra_Error.HandleError "GetAllUdfs", Err',
        '    Resume CleanExit',
        'End Function',
        ''
    )

    foreach ($udf in $udfs) {
        $name = $udf.Name
        $description = $udf.Description.Replace('"', '""')
        $category = $udf.Category.Replace('"', '""')
        $syntax = $udf.Syntax.Replace('"', '""')
        
        $argLines = @()
        if ($null -ne $udf.ArgumentDescriptions) {
            foreach ($arg in $udf.ArgumentDescriptions) {
                $argLines += '        "{0}"' -f $arg.Replace('"', '""')
            }
        }
        
        $argArrayText = if ($argLines.Count -gt 0) {
            "Array( _`r`n" + ($argLines -join ", _`r`n") + ")"
        } else {
            "Array()"
        }

        $lines += "Private Function Get{0}Metadata() As Object" -f $name
        $lines += "    Dim metadata As Object"
        $lines += "    Set metadata = CreateObject(""Scripting.Dictionary"")"
        $lines += "    metadata.Add ""Name"", ""{0}""" -f $name
        $lines += "    metadata.Add ""Description"", ""{0}""" -f $description
        $lines += "    metadata.Add ""Category"", ""{0}""" -f $category
        $lines += "    metadata.Add ""Syntax"", ""{0}""" -f $syntax
        $lines += "    metadata.Add ""ArgumentDescriptions"", {0}" -f $argArrayText
        $lines += "    Set Get{0}Metadata = metadata" -f $name
        $lines += "End Function"
        $lines += ""
    }

    $null = Write-FileIfChanged -Path $OutputPath -Content ($lines -join "`r`n")
    Write-Host "  UDF registry generated with $($udfs.Count) function(s)." -ForegroundColor Green
}

function New-NormalizedImportCopy {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$TempRoot
    )

    $content = [System.IO.File]::ReadAllText($SourcePath)
    # Check if there is any solo LF (LF not preceded by CR)
    $hasSoloLf = $content -match "(?<!`r)`n"
    
    if (-not $hasSoloLf) {
        # File is already CRLF normalized. We can import it directly from the source!
        return $SourcePath
    }

    if (-not (Test-Path $TempRoot)) {
        New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null
    }

    $normalizedPath = Join-Path $TempRoot ([System.IO.Path]::GetFileName($SourcePath))
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
