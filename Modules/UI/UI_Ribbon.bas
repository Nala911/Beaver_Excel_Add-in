Attribute VB_Name = "UI_Ribbon"
Option Explicit

' @Module: UI_Ribbon
' @Category: UI
' @Description: Generated Ribbon callbacks for the Beaver Add-in.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: AppContainer, Infra_Config, Infra_Error

' --- Dynamic UI Callbacks ---

' Returns the image object for a control based on its ID in config.json
Public Sub Ribbon_GetIcon(ByVal control As Object, ByRef image As Variant)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_GetIcon")
    On Error GoTo ErrHandler
    
    Dim iconName As String
    iconName = Infra_Config.GetIcon(control.Id)
    If iconName = "" Then iconName = "Help"
    
    Set image = Application.CommandBars.GetImageMso(iconName, 32, 32)
    
CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_GetIcon", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnWrap(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnWrap")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnWrap", "Ribbon_OnWrap", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnWrap", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnStaticSheetWorkbook(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnStaticSheetWorkbook")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnStaticSheetWorkbook", "Ribbon_OnStaticSheetWorkbook", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnStaticSheetWorkbook", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnCleanData(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnCleanData")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnCleanData", "Ribbon_OnCleanData", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnCleanData", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnModifyData(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnModifyData")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnModifyData", "Ribbon_OnModifyData", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnModifyData", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnHighlightData(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnHighlightData")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnHighlightData", "Ribbon_OnHighlightData", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnHighlightData", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnBreakExternalLinks(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnBreakExternalLinks")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnBreakExternalLinks", "Ribbon_OnBreakExternalLinks", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnBreakExternalLinks", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnDuplicate(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnDuplicate")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnDuplicate", "Ribbon_OnDuplicate", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnDuplicate", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnExport(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnExport")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnExport", "Ribbon_OnExport", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnExport", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnShowHelpCenter(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnShowHelpCenter")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnShowHelpCenter", "Ribbon_OnShowHelpCenter", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnShowHelpCenter", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnHelloWorld(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnHelloWorld")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnHelloWorld", "Ribbon_OnHelloWorld", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnHelloWorld", Err
    Resume CleanExit
End Sub