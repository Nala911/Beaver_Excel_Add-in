VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_MergeFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_MergeFormulas
' @Category: Feature
' @Description: Wraps each selected cell or range item with the first reference found in a user-selected wrapper formula.

Private Const COMMAND_NAME As String = "Wrap With Formula"
Private Const MAX_CELLS_LIMIT As Long = 50000

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy(True, False, True)
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasRangeSelection(context, "Please select a cell or range before running this command.")
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("MergeFormulas.Execute")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Dim targetRange As Range
    Dim wrapperCell As Range
    Dim wrapperFormula As String
    Dim area As Range
    Dim cell As Range
    Dim totalErrors As Long

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit

    Set targetRange = ctx.SelectionRange
    If targetRange Is Nothing Then GoTo CleanExit
    If targetRange.CountLarge > MAX_CELLS_LIMIT Then
        If Not Infra_Interaction.Confirm("You have selected " & Format(targetRange.CountLarge, "#,##0") & " cells. Continue?") Then GoTo CleanExit
    End If

    Set wrapperCell = Infra_UIFactory.PromptForRelatedCell(targetRange.Cells(1, 1))
    If wrapperCell Is Nothing Then GoTo CleanExit
    Set wrapperCell = wrapperCell.Cells(1, 1)

    If Not wrapperCell.HasFormula Then
        Infra_Interaction.ShowWarning "Please select a wrapper cell that contains a formula."
        GoTo CleanExit
    End If

    wrapperFormula = CStr(wrapperCell.Formula2)
    If Left$(wrapperFormula, 1) <> "=" Then
        Infra_Interaction.ShowWarning "The selected wrapper cell does not contain a valid formula."
        GoTo CleanExit
    End If

    If Not ContainsCellReference(wrapperFormula) Then
        Infra_Interaction.ShowWarning "The wrapper formula must contain at least one cell or range reference to replace."
        GoTo CleanExit
    End If

    Infra_Undo.SaveState targetRange, COMMAND_NAME

    For Each area In targetRange.Areas
        If area.CountLarge = 1 Then
            ApplyWrapperToCell area.Cells(1, 1), wrapperFormula, totalErrors
        Else
            For Each cell In area.Cells
                ApplyWrapperToCell cell, wrapperFormula, totalErrors
            Next cell
        End If
    Next area

    If totalErrors > 0 Then
        Infra_Interaction.ShowWarning "Completed with " & totalErrors & " error(s). Some cells could not be updated."
    End If

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "MergeFormulas.Execute", Err
    Resume CleanExit
End Sub

Private Sub ApplyWrapperToCell(ByVal targetCell As Range, ByVal wrapperFormula As String, ByRef errorCount As Long)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ApplyWrapperToCell")
    On Error GoTo ErrHandler

    Dim replacementText As String
    Dim finalFormula As String

    replacementText = GetCellArgumentText(targetCell)
    finalFormula = ReplaceFirstCellReference(wrapperFormula, replacementText)

    If finalFormula = vbNullString Then
        errorCount = errorCount + 1
        Exit Sub
    End If

    On Error Resume Next
    targetCell.Formula2 = finalFormula
    If Err.Number <> 0 Then
        Err.Clear
        errorCount = errorCount + 1
        targetCell.Interior.Color = vbYellow
    End If
    On Error GoTo ErrHandler

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ApplyWrapperToCell", Err
    errorCount = errorCount + 1
    Resume CleanExit
End Sub

Private Function ReplaceFirstCellReference(ByVal formulaText As String, ByVal replacementText As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ReplaceFirstCellReference")
    On Error GoTo ErrHandler

    Dim matchText As String
    Dim startIndex As Long
    Dim tokenLength As Long

    matchText = FindFirstReferenceToken(formulaText, startIndex, tokenLength)
    If startIndex = 0 Or tokenLength = 0 Then GoTo CleanExit

    ReplaceFirstCellReference = Left$(formulaText, startIndex - 1) & replacementText & Mid$(formulaText, startIndex + tokenLength)
    Exit Function

CleanExit:
    ReplaceFirstCellReference = vbNullString
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ReplaceFirstCellReference", Err
    ReplaceFirstCellReference = vbNullString
    Resume CleanExit
End Function

Private Function FindFirstReferenceToken(ByVal formulaText As String, ByRef startIndex As Long, ByRef tokenLength As Long) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("FindFirstReferenceToken")
    On Error GoTo ErrHandler

    Dim re As Object
    Dim matches As Object
    Dim match As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = False
    re.Pattern = "(?:'[^']+'!|[A-Za-z0-9_\.]+!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?:#)?"

    Set matches = re.Execute(formulaText)
    If matches.Count = 0 Then GoTo CleanExit

    Set match = matches(0)
    startIndex = match.FirstIndex + 1
    tokenLength = Len(match.Value)
    FindFirstReferenceToken = match.Value
    Exit Function

CleanExit:
    startIndex = 0
    tokenLength = 0
    FindFirstReferenceToken = vbNullString
    Exit Function

ErrHandler:
    Infra_Error.HandleError "FindFirstReferenceToken", Err
    startIndex = 0
    tokenLength = 0
    FindFirstReferenceToken = vbNullString
    Resume CleanExit
End Function

Private Function ContainsCellReference(ByVal formulaText As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ContainsCellReference")
    On Error GoTo ErrHandler

    Dim startIndex As Long
    Dim tokenLength As Long

    ContainsCellReference = (Len(FindFirstReferenceToken(formulaText, startIndex, tokenLength)) > 0)
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ContainsCellReference", Err
    ContainsCellReference = False
End Function

Private Function GetCellArgumentText(ByVal targetCell As Range) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetCellArgumentText")
    On Error GoTo ErrHandler

    Dim v As Variant
    Dim formulaText As String

    If targetCell Is Nothing Then GoTo CleanExit

    If targetCell.HasFormula Then
        formulaText = CStr(targetCell.Formula2)
        If Left$(formulaText, 1) = "=" Then
            GetCellArgumentText = Mid$(formulaText, 2)
        Else
            GetCellArgumentText = formulaText
        End If
        Exit Function
    Else
        v = targetCell.Value2
    End If

    If IsError(v) Then
        GetCellArgumentText = CStr(targetCell.Text)
    ElseIf VarType(v) = vbString Then
        GetCellArgumentText = """" & Replace(CStr(v), """", """""") & """"
    ElseIf VarType(v) = vbBoolean Then
        GetCellArgumentText = IIf(CBool(v), "TRUE", "FALSE")
    ElseIf IsEmpty(v) Then
        GetCellArgumentText = vbNullString
    Else
        GetCellArgumentText = CStr(v)
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "GetCellArgumentText", Err
    GetCellArgumentText = vbNullString
    Resume CleanExit
End Function
