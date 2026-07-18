VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UI_Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @Module: UI_Factory
' @Category: UI
' @Description: Centralized facade for creating and displaying standardized user prompts. Delegates to specialized dialog modules.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ActionContext, CleanDataRequest, ModifyDataRequest, HighlightDataRequest, ExportRequest, ScopedRequest, Enums, UI_DialogShared, UI_DialogCleanData, UI_DialogModifyData, UI_DialogHighlightData, UI_DialogExport, UI_DialogSheet, UI_DialogWrap

' Shows the Clean Data options via UserForm picker and returns a populated Request object.
Public Function ShowCleanDataDialog(ByVal ctx As ActionContext) As CleanDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowCleanDataDialog")
    On Error GoTo ErrHandler

    Set ShowCleanDataDialog = UI_DialogCleanData.ShowCleanDataDialog(ctx)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowCleanDataDialog", Err
    Resume CleanExit
End Function

' Shows the Modify Data options via UserForm picker and returns a populated Request object.
Public Function ShowModifyDataDialog(ByVal ctx As ActionContext, Optional ByVal commandName As String = vbNullString) As ModifyDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowModifyDataDialog")
    On Error GoTo ErrHandler

    Set ShowModifyDataDialog = UI_DialogModifyData.ShowModifyDataDialog(ctx, commandName)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowModifyDataDialog", Err
    Resume CleanExit
End Function

' Shows the Highlight Data options via UserForm picker and returns a populated Request object.
Public Function ShowHighlightDataDialog(ByVal ctx As ActionContext, Optional ByVal commandName As String = vbNullString) As HighlightDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowHighlightDataDialog")
    On Error GoTo ErrHandler

    Set ShowHighlightDataDialog = UI_DialogHighlightData.ShowHighlightDataDialog(ctx, commandName)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowHighlightDataDialog", Err
    Resume CleanExit
End Function

' Shows the Export options via UserForm picker and returns a populated Request object.
Public Function ShowExportDialog(ByVal ctx As ActionContext, Optional ByVal commandName As String = vbNullString) As ExportRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowExportDialog")
    On Error GoTo ErrHandler

    Set ShowExportDialog = UI_DialogExport.ShowExportDialog(ctx, commandName)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowExportDialog", Err
    Resume CleanExit
End Function

' Shows the conversion scope dialog for formula-to-value actions using a UserForm picker.
Public Function ShowStaticConversionDialog(ByVal ctx As ActionContext) As ScopedRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowStaticConversionDialog")
    On Error GoTo ErrHandler

    Set ShowStaticConversionDialog = UI_DialogShared.ShowStaticConversionDialog(ctx)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowStaticConversionDialog", Err
    Resume CleanExit
End Function

Public Function ShowBreakLinksDialog(ByVal ctx As ActionContext, ByVal linkInfo As String) As ScopedRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowBreakLinksDialog")
    On Error GoTo ErrHandler

    Set ShowBreakLinksDialog = UI_DialogShared.ShowBreakLinksDialog(ctx, linkInfo)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowBreakLinksDialog", Err
    Resume CleanExit
End Function

Public Function PromptForSheetInsertPosition(ByVal ctx As ActionContext, ByVal sheetName As String) As SheetInsertPosition
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForSheetInsertPosition")
    On Error GoTo ErrHandler

    PromptForSheetInsertPosition = UI_DialogSheet.PromptForSheetInsertPosition(ctx, sheetName)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForSheetInsertPosition", Err
    Resume CleanExit
End Function

Public Function PromptForSheetName(ByVal ctx As ActionContext) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForSheetName")
    On Error GoTo ErrHandler

    PromptForSheetName = UI_DialogSheet.PromptForSheetName(ctx)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForSheetName", Err
    Resume CleanExit
End Function

Public Function PromptForRelatedCell(ByVal sourceCell As Range) As Range
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForRelatedCell")
    On Error GoTo ErrHandler

    Set PromptForRelatedCell = UI_DialogWrap.PromptForRelatedCell(sourceCell)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForRelatedCell", Err
    Resume CleanExit
End Function

Public Function PromptForWrapMode(ByVal ctx As ActionContext) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForWrapMode")
    On Error GoTo ErrHandler

    PromptForWrapMode = UI_DialogWrap.PromptForWrapMode(ctx)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForWrapMode", Err
    Resume CleanExit
End Function

Public Function PromptForWrapFormulaPattern(ByVal ctx As ActionContext, ByVal placeholder As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForWrapFormulaPattern")
    On Error GoTo ErrHandler

    PromptForWrapFormulaPattern = UI_DialogWrap.PromptForWrapFormulaPattern(ctx, placeholder)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForWrapFormulaPattern", Err
    Resume CleanExit
End Function
