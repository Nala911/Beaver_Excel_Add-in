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
Implements IBatchProcessor
Implements ICellTransformer

' @Module: FeatCmd_ForceNumber
' @Category: Feature
' @Description: Forces text-formatted numbers in the selection to become actual numeric values.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, IBatchProcessor, ICellTransformer, Infra_ActionContext, Infra_Error, Infra_Undo, Infra_Progress, Infra_Interaction, Infra_Config

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
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Or ctx.WorkbookRef Is Nothing Then GoTo CleanExit

    Dim totalProcessed As Long
    ' Execute batch processing via the unified engine
    totalProcessed = Infra_BatchProcessor.Execute(ctx, TargetScopeSelection, Me, "Convert Text to Number", False)

    If totalProcessed > 0 Then
        Infra_Interaction.ShowInfo "Successfully converted " & totalProcessed & " cells to numbers.", "Convert Text to Number"
    Else
        Infra_Interaction.ShowInfo "No cells required conversion.", "Convert Text to Number"
    End If

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ForceNumber.Execute", Err
    Resume CleanExit
End Sub

Private Function IBatchProcessor_ProcessRange(ByVal targetRange As Range) As Long
    IBatchProcessor_ProcessRange = Infra_CommandSupport.ProcessRangeUnified(targetRange, Me, xlTextValues)
End Function

Private Function IBatchProcessor_ProcessWorksheet(ByVal ws As Worksheet) As Long
    IBatchProcessor_ProcessWorksheet = 0
End Function

Private Function ICellTransformer_TransformCell( _
    ByVal rawValue As Variant, _
    ByVal currentFormat As String, _
    ByRef outValue As Variant, _
    ByRef outFormat As String) As Boolean

    Dim outNum As Variant
    If Infra_ValueConversion.TryConvertToNumber(rawValue, outNum) Then
        outValue = outNum
        If currentFormat = "__MIXED__" Then
            outFormat = "__FORCE_GENERAL_IF_TEXT__"
        ElseIf InStr(1, currentFormat, "@") > 0 Or LCase$(currentFormat) = "text" Then
            outFormat = "General"
        Else
            outFormat = currentFormat
        End If
        ICellTransformer_TransformCell = True
    Else
        ICellTransformer_TransformCell = False
    End If
End Function
