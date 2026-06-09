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

' @Module: FeatCmd_StaticSheetWorkbook
' @Category: Feature
' @Description: Converts formulas (including spills) to static values across active sheet or whole workbook.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("StaticSheetWorkbook.Validate")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)

    If ctx Is Nothing Then
        Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("There is no active workbook available for this command.")
    ElseIf ctx.WorkbookRef Is Nothing Or ctx.WorksheetRef Is Nothing Then
        Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("There is no active worksheet available for this command.")
    ElseIf ctx.WorkbookRef Is ThisWorkbook Then
        Set ICommand_Validate = Infra_CommandSupport.ValidationFailure( _
            "Make Static cannot run while the Beaver add-in workbook is active. Switch to the workbook you want to convert and try again.")
    Else
        Set ICommand_Validate = Infra_CommandSupport.ValidationSuccess()
    End If

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

    Dim ctx As Infra_ActionContext
    Dim request As Infra_ScopedRequest

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorkbookRef Is Nothing Or ctx.WorksheetRef Is Nothing Then GoTo CleanExit

    Set request = Infra_UIFactory.ShowStaticConversionDialog(ctx)
    If request Is Nothing Then GoTo CleanExit

    ExecuteStaticConversion request

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "StaticSheetWorkbook.Execute", Err
    Resume CleanExit
End Sub

Private Sub ExecuteStaticConversion(ByVal request As Infra_ScopedRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteStaticConversion")
    On Error GoTo ErrHandler

    Dim sheetsToProcess As Collection
    Dim wsItem As Variant
    Dim ws As Worksheet
    Dim convertedCount As Long
    Dim index As Long
    Dim scopeLabel As String
    Dim originalScreenUpdating As Boolean

    If request Is Nothing Then GoTo CleanExit
    If request.Context Is Nothing Then GoTo CleanExit

    Set sheetsToProcess = Infra_CommandSupport.ResolveWorksheetsToProcess(request.Context, request.Scope)
    If sheetsToProcess Is Nothing Then GoTo CleanExit

    If request.Scope = TargetScopeActiveSheet Then
        If Not Infra_Undo.SaveStateOrConfirm(request.Context.WorksheetRef.UsedRange, "Static Sheet") Then GoTo CleanExit
    End If

    originalScreenUpdating = Application.ScreenUpdating
    If request.Scope = TargetScopeWorkbook Then
        Application.ScreenUpdating = False
    End If

    Infra_Progress.StartProgress "Converting Formulas to Values", sheetsToProcess.Count
    For Each wsItem In sheetsToProcess
        index = index + 1
        Set ws = wsItem
        convertedCount = convertedCount + ConvertSheetToValues(ws)
        Infra_Progress.UpdateProgress index, True
        If Infra_Progress.UserCancelled Then Exit For
    Next wsItem

    scopeLabel = IIf(request.Scope = TargetScopeWorkbook, "Whole Workbook", "Active Sheet")
    Infra_Interaction.ShowInfo "Done! Converted " & convertedCount & " formula cell(s) - " & scopeLabel & ".", _
                               Infra_Config.ADDIN_NAME & " - Process Complete"

CleanExit:
    If Not request Is Nothing Then
        If request.Scope = TargetScopeWorkbook Then
            Application.ScreenUpdating = originalScreenUpdating
        End If
    End If
    Infra_Progress.EndProgress
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    Infra_Error.HandleError "ExecuteStaticConversion", Err
    Resume CleanExit
End Sub



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
