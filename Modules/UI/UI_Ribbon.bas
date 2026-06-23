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

Public Sub Ribbon_OnDateFixer(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnDateFixer")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnDateFixer", "Ribbon_OnDateFixer", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnDateFixer", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnCaseFixer(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnCaseFixer")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnCaseFixer", "Ribbon_OnCaseFixer", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnCaseFixer", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnHighlightInconsistentFormulas(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnHighlightInconsistentFormulas")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnHighlightInconsistentFormulas", "Ribbon_OnHighlightInconsistentFormulas", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnHighlightInconsistentFormulas", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnHighlightDuplicates(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnHighlightDuplicates")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnHighlightDuplicates", "Ribbon_OnHighlightDuplicates", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnHighlightDuplicates", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnHighlightErrors(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnHighlightErrors")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnHighlightErrors", "Ribbon_OnHighlightErrors", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnHighlightErrors", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnHighlightHardcodedValues(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnHighlightHardcodedValues")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnHighlightHardcodedValues", "Ribbon_OnHighlightHardcodedValues", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnHighlightHardcodedValues", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnHighlightDataValidations(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnHighlightDataValidations")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnHighlightDataValidations", "Ribbon_OnHighlightDataValidations", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnHighlightDataValidations", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnHighlightConditionalFormatting(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnHighlightConditionalFormatting")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnHighlightConditionalFormatting", "Ribbon_OnHighlightConditionalFormatting", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnHighlightConditionalFormatting", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnClearHighlights(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnClearHighlights")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnClearHighlights", "Ribbon_OnClearHighlights", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnClearHighlights", Err
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

Public Sub Ribbon_OnExportPng(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnExportPng")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnExportPng", "Ribbon_OnExportPng", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnExportPng", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnExportPdf(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnExportPdf")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnExportPdf", "Ribbon_OnExportPdf", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnExportPdf", Err
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

Public Sub Ribbon_OnTableOfContents(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnTableOfContents")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnTableOfContents", "Ribbon_OnTableOfContents", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnTableOfContents", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnUnmergeFill(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnUnmergeFill")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnUnmergeFill", "Ribbon_OnUnmergeFill", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnUnmergeFill", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnForceNumber(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnForceNumber")
    On Error GoTo ErrHandler

    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnForceNumber", "Ribbon_OnForceNumber", "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnForceNumber", Err
    Resume CleanExit
End Sub