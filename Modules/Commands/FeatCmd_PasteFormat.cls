VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_PasteFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_PasteFormat
' @Category: Feature
' @Description: Pastes cell formatting only. Requires user to have copied something first.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    ' Keep the current calculation mode to avoid clearing the clipboard (CutCopyMode)
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy(False, False, True, Application.Calculation)
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("PasteFormat.Execute")
    On Error GoTo ErrHandler
    
    Dim targetRange As Range
    Dim ctx As ActionContext
    Dim tempWs As Worksheet
    Dim originalActiveSheet As Object

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit
    Set targetRange = ctx.SelectionRange
    If targetRange Is Nothing Then GoTo CleanExit

    Dim safeRange As Range
    Set safeRange = Infra_CommandSupport.GetSafeProcessingRange(targetRange, 50000, True)
    If safeRange Is Nothing Then GoTo CleanExit
    Set targetRange = safeRange

    If Application.CutCopyMode = False Then
        Infra_Interaction.ShowWarning "Nothing to paste. Please copy a range first."
        GoTo CleanExit
    End If

    ' Record the original active sheet to restore it later
    Set originalActiveSheet = Application.ActiveSheet

    ' Create a temporary worksheet in the add-in workbook (ThisWorkbook) to preserve the user's active workbook Undo stack
    On Error Resume Next
    Set tempWs = ThisWorkbook.Worksheets.Add
    If tempWs Is Nothing Then
        ' Fallback: If workbook structure is protected, paste directly without undo support
        On Error Resume Next
        targetRange.PasteSpecial Paste:=xlPasteFormats
        If Err.Number <> 0 Then
            Err.Clear
            Infra_Interaction.ShowWarning "Nothing to paste. Please copy a range first."
        End If
        On Error GoTo ErrHandler
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler
    
    ' Paste clipboard formats onto the temporary sheet at the same target address coordinates
    On Error Resume Next
    tempWs.Range(targetRange.Address).PasteSpecial Paste:=xlPasteFormats
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrHandler
        Application.DisplayAlerts = False
        tempWs.Delete
        Application.DisplayAlerts = True
        Infra_Interaction.ShowWarning "Nothing to paste. Please copy a range first."
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler

    ' Hide tempWs temporarily so that Excel automatically restores active focus to originalActiveSheet
    tempWs.Visible = xlSheetHidden

    ' Save the original state of the target range for Undo (this clears/overwrites the Excel Clipboard)
    If Not Infra_Undo.SaveStateOrConfirm(targetRange, "Paste Format") Then
        Application.DisplayAlerts = False
        tempWs.Delete
        Application.DisplayAlerts = True
        GoTo CleanExit
    End If

    ' Copy the formats from the temporary sheet and paste them onto targetRange
    On Error Resume Next
    tempWs.Range(targetRange.Address).Copy
    targetRange.PasteSpecial Paste:=xlPasteFormats
    
    If Err.Number <> 0 Then
        Err.Clear
        Infra_Undo.ClearPendingUndo
        Infra_Interaction.ShowWarning "Failed to paste format."
    End If
    On Error GoTo ErrHandler

    ' Clean up the temporary sheet
    Application.DisplayAlerts = False
    tempWs.Delete
    Application.DisplayAlerts = True
    
    Application.CutCopyMode = False

CleanExit:
    If Not originalActiveSheet Is Nothing Then
        On Error Resume Next
        originalActiveSheet.Activate
        On Error GoTo 0
    End If
    Exit Sub

ErrHandler:
    If Not tempWs Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        tempWs.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
    If Not originalActiveSheet Is Nothing Then
        On Error Resume Next
        originalActiveSheet.Activate
        On Error GoTo 0
    End If
    Infra_Error.HandleError "PasteFormat.Execute", Err
    Resume CleanExit
End Sub
