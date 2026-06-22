VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_UnmergeFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_UnmergeFill
' @Category: Feature
' @Description: Unmerges cells and fills the top-left value across the unmerged area.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, Infra_ActionContext, Infra_Error, Infra_Undo, Infra_Progress, Infra_Interaction, Infra_Config

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyBulkWrite()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("UnmergeFill.Execute")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Dim selRange As Range
    Dim ws As Worksheet
    Dim cell As Range
    Dim hasMerged As Boolean
    Dim unmergeCount As Long

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

    ' Check if there are merged cells
    If IsNull(selRange.MergeCells) Then
        hasMerged = True
    Else
        hasMerged = selRange.MergeCells
    End If

    If Not hasMerged Then
        Infra_Interaction.ShowInfo "No merged cells found in the selected range.", "Unmerge & Fill Down"
        GoTo CleanExit
    End If

    ' Save state for Undo
    If Not Infra_Undo.SaveStateOrConfirm(selRange, "Unmerge & Fill") Then GoTo CleanExit

    Infra_Progress.StartProgress "Unmerging & Filling", 100
    Infra_Progress.UpdateProgress 10

    Dim mArea As Range
    Dim hasFormula As Boolean
    Dim formulaStr As String
    Dim val As Variant

    On Error Resume Next
    Application.FindFormat.Clear
    Application.FindFormat.MergeCells = True
    Set cell = selRange.Find(What:="", SearchFormat:=True)
    On Error GoTo ErrHandler

    unmergeCount = 0
    Do While Not cell Is Nothing
        Set mArea = cell.MergeArea

        hasFormula = mArea.Cells(1, 1).HasFormula
        If hasFormula Then
            formulaStr = mArea.Cells(1, 1).Formula2
        End If
        val = mArea.Cells(1, 1).Value

        mArea.UnMerge

        If hasFormula Then
            mArea.Formula2 = formulaStr
        Else
            mArea.Value = val
        End If

        unmergeCount = unmergeCount + 1

        ' Find the next remaining merged cell
        On Error Resume Next
        Set cell = selRange.Find(What:="", SearchFormat:=True)
        On Error GoTo ErrHandler
        
        If unmergeCount Mod 100 = 0 Then
            DoEvents
            If Infra_Progress.UserCancelled Then Exit Do
        End If
    Loop

    On Error Resume Next
    Application.FindFormat.Clear
    On Error GoTo 0

    Infra_Progress.UpdateProgress 100
    Infra_Progress.EndProgress

    If unmergeCount > 0 Then
        Infra_Interaction.ShowInfo "Successfully unmerged and filled " & unmergeCount & " areas.", "Unmerge & Fill Down"
    Else
        Infra_Interaction.ShowInfo "No merged cells were processed.", "Unmerge & Fill Down"
    End If

CleanExit:
    On Error Resume Next
    Application.FindFormat.Clear
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    Infra_Error.HandleError "UnmergeFill.Execute", Err
    Resume CleanExit
End Sub
