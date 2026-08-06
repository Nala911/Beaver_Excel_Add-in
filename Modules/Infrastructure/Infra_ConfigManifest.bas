Attribute VB_Name = "Infra_ConfigManifest"
Option Explicit
Option Private Module

' @Module: Infra_ConfigManifest
' @Category: Infrastructure
' @Description: Compiled configuration manifest generated automatically from features.json and config.json.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_HotkeyDefinition

Public Function GetEmbeddedHotkeys() As Collection
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetEmbeddedHotkeys")
    On Error GoTo ErrHandler
    
    Dim col As New Collection
    Dim hk As Infra_HotkeyDefinition

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "^+4"
    hk.MacroName = "UI_Hotkeys.Hotkey_ApplyDefaultFormat"
    hk.Description = "Apply default number format"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "^2"
    hk.MacroName = "UI_Hotkeys.Hotkey_ApplyCustomNumberFormat"
    hk.Description = "Apply custom number format with dialog"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "^+p"
    hk.MacroName = "UI_Hotkeys.Hotkey_MakePermanent"
    hk.Description = "Make selection permanent (values)"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "+{F11}"
    hk.MacroName = "UI_Hotkeys.Hotkey_CreateNamedSheet"
    hk.Description = "Create named sheet"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "^%{DOWN}"
    hk.MacroName = "UI_Hotkeys.Hotkey_FillDown"
    hk.Description = "Fill formula/format down"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "^%{RIGHT}"
    hk.MacroName = "UI_Hotkeys.Hotkey_FillRight"
    hk.Description = "Fill formula/format right"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "^+f"
    hk.MacroName = "UI_Hotkeys.Hotkey_FilterBySelectedCell"
    hk.Description = "Filter by selected cell"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "^+m"
    hk.MacroName = "UI_Hotkeys.Hotkey_PasteFormat"
    hk.Description = "Paste formats only"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "^+q"
    hk.MacroName = "UI_Hotkeys.Hotkey_FormatSelectedRange"
    hk.Description = "Format selected range"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "{BACKSPACE}"
    hk.MacroName = "UI_Hotkeys.Hotkey_Backspace"
    hk.Description = "Clear cell contents"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set hk = New Infra_HotkeyDefinition
    hk.KeyPattern = "{DELETE}"
    hk.MacroName = "UI_Hotkeys.Hotkey_Delete"
    hk.Description = "Clear or delete selection"
    hk.ReleaseTier = "stable"
    col.Add hk

    Set GetEmbeddedHotkeys = col

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetEmbeddedHotkeys", Err
    Resume CleanExit
End Function

Public Function GetEmbeddedUIConstants() As Object
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetEmbeddedUIConstants")
    On Error GoTo ErrHandler
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "DefaultFontName", "Calibri"
    dict.Add "DefaultFontSize", 10&
    dict.Add "HeaderFontSize", 11&
    dict.Add "DefaultNumberFormat", "#,##0"
    dict.Add "DisplayDateFormat", "dd/mm/yyyy"
    dict.Add "ColumnWidthThreshold", 40&
    dict.Add "MaxColumnWidth", 25&
    dict.Add "HeaderColor", "#AEAAAA"
    dict.Add "HighlightColor", "#FFC7CE"
    dict.Add "HighlightNamedRangesColor", "#DCE6F2"
    dict.Add "DefaultExportScale", 3&
    dict.Add "MaxExportScale", 10&
    
    Set GetEmbeddedUIConstants = dict

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetEmbeddedUIConstants", Err
    Resume CleanExit
End Function

Public Function GetEmbeddedSafetyConstants() As Object
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetEmbeddedSafetyConstants")
    On Error GoTo ErrHandler
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "MaxUndoCells", 1000000&
    dict.Add "MaxWrapCells", 50000&
    dict.Add "MaxFormulaCheckCells", 5000&
    dict.Add "MaxFillProximityColumns", 15&
    dict.Add "MaxFillDownAreas", 5000&
    dict.Add "ChunkRowsLimit", 20000&
    dict.Add "ChunkCellsLimit", 50000&
    
    Set GetEmbeddedSafetyConstants = dict

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetEmbeddedSafetyConstants", Err
    Resume CleanExit
End Function