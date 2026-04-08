Attribute VB_Name = "UI_Ribbon"
Option Explicit

' @Module: UI_Ribbon
' @Category: UI
' @Description: Centralized Ribbon callbacks for the Beaver Add-in.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: AppContainer, Infra_Config, Infra_Hotkeys, Infra_Error

' --- Dynamic UI Callbacks ---

' Returns the image object for a control based on its ID in config.json
Public Sub Ribbon_GetIcon(ByVal control As Object, ByRef image As Variant)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_GetIcon")
    On Error GoTo ErrHandler
    
    Dim iconName As String
    iconName = Infra_Config.GetIcon(control.Id)
    If iconName = "" Then iconName = "Help"
    
    ' Get built-in imageMso
    Set image = Application.CommandBars.GetImageMso(iconName, 32, 32)
    
CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_GetIcon", Err
    Resume CleanExit
End Sub

' --- Help Group ---

Public Sub Ribbon_OnShowHotkeysHelp(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnShowHotkeysHelp")
    On Error GoTo ErrHandler
    
    Infra_Hotkeys.ShowHotkeysHelp

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnShowHotkeysHelp", Err
    Resume CleanExit
End Sub

' --- Formatting Group ---

Public Sub Ribbon_OnMergeFormulas(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnMergeFormulas")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "MergeFormulas", "Ribbon_OnMergeFormulas", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnMergeFormulas", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnWrapSelectionWithFormula(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnWrapSelectionWithFormula")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "WrapSelectedRange", "Ribbon_OnWrapSelectionWithFormula", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnWrapSelectionWithFormula", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnStaticSheetWorkbook(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnStaticSheetWorkbook")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "StaticSheetWorkbook", "Ribbon_OnStaticSheetWorkbook", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnStaticSheetWorkbook", Err
    Resume CleanExit
End Sub

' --- Cleanup Group ---

Public Sub Ribbon_OnCleanData(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnCleanData")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "CleanData", "Ribbon_OnCleanData", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnCleanData", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnBreakExternalLinks(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnBreakExternalLinks")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "BreakExternalLinks", "Ribbon_OnBreakExternalLinks", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnBreakExternalLinks", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnConvertTextToProperDate(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnConvertTextToProperDate")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "DateConversion", "Ribbon_OnConvertTextToProperDate", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnConvertTextToProperDate", Err
    Resume CleanExit
End Sub

' --- Export Group ---

Public Sub Ribbon_OnDashboard(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnDashboard")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "Dashboard", "Ribbon_OnDashboard", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnDashboard", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnDuplicate(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnDuplicate")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "Duplicate", "Ribbon_OnDuplicate", "Ribbon"
CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnDuplicate", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnExport(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnExport")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "ExportImageOrPdf", "Ribbon_OnExport", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnExport", Err
    Resume CleanExit
End Sub

' --- Structure Group ---

Public Sub Ribbon_OnToggleFullScreen(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnToggleFullScreen")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "ToggleFullScreen", "Ribbon_OnToggleFullScreen", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnToggleFullScreen", Err
    Resume CleanExit
End Sub
