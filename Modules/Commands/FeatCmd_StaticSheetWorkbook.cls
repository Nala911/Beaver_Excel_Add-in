VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_StaticSheetWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand
Implements IBatchProcessor

' @Module: FeatCmd_StaticSheetWorkbook
' @Category: Feature
' @Description: Converts formulas (including spills) to static values across active sheet or whole workbook.

Private m_CurrentRequest As ScopedRequest

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("StaticSheetWorkbook.Validate")
    On Error GoTo ErrHandler

    Dim result As CommandValidationResult
    Set result = Infra_CommandSupport.ValidateHasWorksheet(context)
    If Not result.IsExecutable Then
        Set ICommand_Validate = result
        Exit Function
    End If

    Set result = Infra_CommandSupport.ValidateActiveWorkbookNotAddin(context, "Make Static")
    If Not result.IsExecutable Then
        ' Match the original warning message
        result.Message = "Make Static cannot run while the Beaver add-in workbook is active. Switch to the workbook you want to convert and try again."
        Set ICommand_Validate = result
        Exit Function
    End If

    Set result = Infra_CommandSupport.ValidateWorksheetNotProtected(context, "Cannot make formulas static on a protected worksheet. Please unprotect the sheet and try again.")
    If Not result.IsExecutable Then
        Set ICommand_Validate = result
        Exit Function
    End If

    Set ICommand_Validate = Infra_CommandSupport.ValidationSuccess()

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "StaticSheetWorkbook.Validate", Err
    Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("The workbook cannot be converted to static values right now.")
    Resume CleanExit
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("StaticSheetWorkbook.Execute")
    On Error GoTo ErrHandler

    Dim ctx As ActionContext
    Dim request As ScopedRequest

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorkbookRef Is Nothing Or ctx.WorksheetRef Is Nothing Then GoTo CleanExit

    Set request = UI_Factory.ShowStaticConversionDialog(ctx)
    If request Is Nothing Then GoTo CleanExit

    ExecuteStaticConversion request

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "StaticSheetWorkbook.Execute", Err
    Resume CleanExit
End Sub

Private Sub ExecuteStaticConversion(ByVal request As ScopedRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteStaticConversion")
    On Error GoTo ErrHandler

    If request Is Nothing Then GoTo CleanExit
    If request.Context Is Nothing Then GoTo CleanExit

    Set m_CurrentRequest = request

    Dim totalConverted As Long
    totalConverted = Infra_BatchProcessor.Execute(request.Context, request.Scope, Me, "Make Static", False)

    Dim scopeLabel As String
    scopeLabel = IIf(request.Scope = TargetScopeWorkbook, "Whole Workbook", "Active Sheet")
    Infra_Interaction.ShowInfo "Done! Converted " & totalConverted & " formula cell(s) - " & scopeLabel & ".", _
                               Infra_Config.ADDIN_NAME & " - Process Complete"

CleanExit:
    Set m_CurrentRequest = Nothing
    Exit Sub

ErrHandler:
    Set m_CurrentRequest = Nothing
    Infra_Error.HandleError "ExecuteStaticConversion", Err
    Resume CleanExit
End Sub

Private Function IBatchProcessor_ProcessRange(ByVal targetRange As Range) As Long
    IBatchProcessor_ProcessRange = 0
End Function

Private Function IBatchProcessor_ProcessWorksheet(ByVal ws As Worksheet) As Long
    IBatchProcessor_ProcessWorksheet = ConvertSheetToValues(ws)
End Function



Private Function ConvertSheetToValues(ByVal ws As Worksheet) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConvertSheetToValues")
    On Error GoTo ErrHandler

    If ws Is Nothing Then GoTo CleanExit
    ConvertSheetToValues = Infra_ValueConversion.ConvertWorksheetFormulasToValues(ws)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ConvertSheetToValues", Err
    Resume CleanExit
End Function

Public Function TestConvertSheetToValuesDirect(ByVal ws As Worksheet) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("TestConvertSheetToValuesDirect")
    On Error GoTo ErrHandler

    TestConvertSheetToValuesDirect = ConvertSheetToValues(ws)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "TestConvertSheetToValuesDirect", Err
    Resume CleanExit
End Function
