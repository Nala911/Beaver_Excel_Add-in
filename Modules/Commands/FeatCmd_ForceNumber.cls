VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_ForceNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_ForceNumber
' @Category: Feature
' @Description: Forces text-formatted numbers in the selection to become actual numeric values.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, Infra_ActionContext, Infra_Error, Infra_Undo, Infra_Progress, Infra_Interaction, Infra_Config

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyBulkWrite()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ForceNumber.Execute")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Dim selRange As Range
    Dim ws As Worksheet
    Dim textCells As Range
    Dim cell As Range
    Dim changeCount As Long

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit

    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then
        GoTo CleanExit
    End If

    Set selRange = ctx.SelectionRange
    Set ws = ctx.WorksheetRef
    If ws Is Nothing Or selRange Is Nothing Then GoTo CleanExit

    ' Restrict selection to UsedRange for safety
    Set selRange = Infra_CommandSupport.GetSafeProcessingRange(selRange, 50000, True)
    If selRange Is Nothing Then GoTo CleanExit

    ' Identify cells with text constants
    If selRange.CountLarge = 1 Then
        If Not selRange.HasFormula And VarType(selRange.Value) = vbString Then
            Set textCells = selRange
        End If
    Else
        On Error Resume Next
        Set textCells = selRange.SpecialCells(xlCellTypeConstants, xlTextValues)
        On Error GoTo 0
    End If

    If textCells Is Nothing Then
        Infra_Interaction.ShowInfo "No text cells found in the selection.", "Convert Text to Number"
        GoTo CleanExit
    End If

    ' Save state for Undo
    If Not Infra_Undo.SaveStateOrConfirm(selRange, "Convert Text to Number") Then GoTo CleanExit

    Infra_Progress.StartProgress "Converting Text to Numbers", 100
    Infra_Progress.UpdateProgress 10

    changeCount = 0
    Dim outVal As Variant
    For Each cell In textCells.Cells
        If Infra_ValueConversion.TryConvertToNumber(cell.Value2, outVal) Then
            If InStr(1, cell.NumberFormat, "@") > 0 Or LCase$(cell.NumberFormat) = "text" Then
                cell.NumberFormat = "General"
            End If
            cell.Value2 = outVal
            changeCount = changeCount + 1
        End If
    Next cell

    Infra_Progress.UpdateProgress 100
    Infra_Progress.EndProgress

    If changeCount > 0 Then
        Infra_Interaction.ShowInfo "Successfully converted " & changeCount & " cells to numbers.", "Convert Text to Number"
    Else
        Infra_Interaction.ShowInfo "No cells required conversion.", "Convert Text to Number"
    End If

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    Infra_Error.HandleError "ForceNumber.Execute", Err
    Resume CleanExit
End Sub

