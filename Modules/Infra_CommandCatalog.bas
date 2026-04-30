Attribute VB_Name = "Infra_CommandCatalog"
Option Explicit

' @Module: Infra_CommandCatalog
' @Category: Infrastructure
' @Description: Loads manifest-driven command metadata and resolves UI entry points to command names.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_JsonConverter, Infra_Error

Private pEntryMacroToCommand As Object

Public Sub ResetCatalog()
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResetCatalog")
    On Error GoTo ErrHandler

    Set pEntryMacroToCommand = Nothing

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ResetCatalog", Err
    Resume CleanExit
End Sub

Public Function ResolveCommandName(ByVal entryMacro As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResolveCommandName")
    On Error GoTo ErrHandler

    EnsureCatalogLoaded
    If pEntryMacroToCommand Is Nothing Then GoTo CleanExit
    If pEntryMacroToCommand.Exists(NormalizeKey(entryMacro)) Then
        ResolveCommandName = CStr(pEntryMacroToCommand(NormalizeKey(entryMacro)))
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ResolveCommandName", Err
    Resume CleanExit
End Function

Private Sub EnsureCatalogLoaded()
    If Not pEntryMacroToCommand Is Nothing Then Exit Sub

    Dim tracker As Object: Set tracker = Infra_Error.Track("EnsureCatalogLoaded")
    On Error GoTo ErrHandler

    Set pEntryMacroToCommand = CreateObject("Scripting.Dictionary")
    pEntryMacroToCommand.CompareMode = vbTextCompare

    LoadManifestMappings
    If pEntryMacroToCommand.Count = 0 Then RegisterFallbackMappings

CleanExit:
    Exit Sub

ErrHandler:
    RegisterFallbackMappings
    Resume CleanExit
End Sub

Private Sub LoadManifestMappings()
    Dim tracker As Object: Set tracker = Infra_Error.Track("LoadManifestMappings")
    On Error GoTo ErrHandler

    Dim fso As Object
    Dim manifestPath As String
    Dim jsonText As String
    Dim textStream As Object
    Dim manifest As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    manifestPath = fso.BuildPath(ThisWorkbook.Path, "features.json")
    If Not fso.FileExists(manifestPath) Then GoTo CleanExit

    Set textStream = fso.OpenTextFile(manifestPath, 1)
    jsonText = textStream.ReadAll
    textStream.Close

    Set manifest = Lib_JsonConverter.ParseJson(jsonText)
    RegisterFeatureMappings manifest
    RegisterHotkeyMappings manifest

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "LoadManifestMappings", Err
    Resume CleanExit
End Sub

Private Sub RegisterFeatureMappings(ByVal manifest As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("RegisterFeatureMappings")
    On Error GoTo ErrHandler

    Dim features As Object
    Dim feature As Variant
    Dim commandName As String
    Dim entryMacro As String

    Set features = manifest("Features")
    For Each feature In features
        commandName = GetObjectText(feature, "CommandName")
        entryMacro = GetObjectText(feature, "Macro")
        RegisterMapping entryMacro, commandName
    Next feature

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "RegisterFeatureMappings", Err
    Resume CleanExit
End Sub

Private Sub RegisterHotkeyMappings(ByVal manifest As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("RegisterHotkeyMappings")
    On Error GoTo ErrHandler

    Dim hotkeys As Object
    Dim hotkey As Variant
    Dim commandName As String
    Dim entryMacro As String

    Set hotkeys = manifest("Hotkeys")
    For Each hotkey In hotkeys
        commandName = GetObjectText(hotkey, "CommandName")
        entryMacro = GetObjectText(hotkey, "Macro")
        RegisterMapping entryMacro, commandName
    Next hotkey

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "RegisterHotkeyMappings", Err
    Resume CleanExit
End Sub

Private Sub RegisterFallbackMappings()
    If pEntryMacroToCommand Is Nothing Then
        Set pEntryMacroToCommand = CreateObject("Scripting.Dictionary")
        pEntryMacroToCommand.CompareMode = vbTextCompare
    End If

    RegisterMapping "UI_Ribbon.Ribbon_OnMergeFormulas", "MergeFormulas"
    RegisterMapping "UI_Ribbon.Ribbon_OnWrapSelectionWithFormula", "WrapSelectedRange"
    RegisterMapping "UI_Ribbon.Ribbon_OnStaticSheetWorkbook", "StaticSheetWorkbook"
    RegisterMapping "UI_Ribbon.Ribbon_OnCleanData", "CleanData"
    RegisterMapping "UI_Ribbon.Ribbon_OnBreakExternalLinks", "BreakExternalLinks"
    RegisterMapping "UI_Ribbon.Ribbon_OnConvertTextToProperDate", "DateConversion"
    RegisterMapping "UI_Ribbon.Ribbon_OnDuplicate", "Duplicate"
    RegisterMapping "UI_Ribbon.Ribbon_OnExport", "ExportImageOrPdf"
    RegisterMapping "UI_Ribbon.Ribbon_OnDashboard", "Dashboard"
    RegisterMapping "UI_Ribbon.Ribbon_OnToggleFullScreen", "ToggleFullScreen"
    RegisterMapping "Infra_Hotkeys.ShowHotkeysHelp", "ShowHotkeysHelp"

    RegisterMapping "UI_Hotkeys.Hotkey_ApplyCustomNumberFormat", "ApplyCustomNumberFormat"
    RegisterMapping "UI_Hotkeys.Hotkey_MakePermanent", "MakePermanent"
    RegisterMapping "UI_Hotkeys.Hotkey_CreateNamedSheet", "CreateSheet"
    RegisterMapping "UI_Hotkeys.Hotkey_FillDown", "FillDown"
    RegisterMapping "UI_Hotkeys.Hotkey_FilterBySelectedCell", "FilterByCell"
    RegisterMapping "UI_Hotkeys.Hotkey_PasteFormat", "PasteFormat"
    RegisterMapping "UI_Hotkeys.Hotkey_FormatSelectedRange", "FormatRange"
    RegisterMapping "UI_Hotkeys.Hotkey_Backspace", "Backspace"
    RegisterMapping "UI_Hotkeys.Hotkey_Delete", "Delete"
End Sub

Private Sub RegisterMapping(ByVal entryMacro As String, ByVal commandName As String)
    If entryMacro = vbNullString Or commandName = vbNullString Then Exit Sub
    pEntryMacroToCommand(NormalizeKey(entryMacro)) = commandName
End Sub

Private Function GetObjectText(ByVal source As Object, ByVal key As String) As String
    On Error Resume Next
    GetObjectText = Trim$(CStr(source(key)))
    On Error GoTo 0
End Function

Private Function NormalizeKey(ByVal value As String) As String
    NormalizeKey = UCase$(Trim$(value))
End Function
