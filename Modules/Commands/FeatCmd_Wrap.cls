VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_Wrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_Wrap
' @Category: Feature
' @Description: Wraps the current selection by either applying a typed pattern or reusing a wrapper formula from another cell.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, Infra_CommandSupport, Infra_UIFactory, Infra_Interaction, Infra_Undo, Infra_Error

Private Const PLACEHOLDER As String = "[value]"
Private Const COMMAND_NAME As String = "Wrap"

Private m_RegExp As Object

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context, "The current selection cannot be wrapped.")
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Wrap.Execute")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Dim rng As Range
    Dim wrapMode As Long

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit

    Set rng = ctx.SelectionRange
    If rng Is Nothing Then GoTo CleanExit

    If rng.CountLarge > Infra_Config.MAX_WRAP_CELLS Then
        If Not Infra_Interaction.Confirm("You have selected " & Format(rng.CountLarge, "#,##0") & " cells. This may take a moment to process. Continue?") Then GoTo CleanExit
    End If

    wrapMode = Infra_UIFactory.PromptForWrapMode(ctx)
    If wrapMode = 0 Then GoTo CleanExit

    Select Case wrapMode
        Case WrapModeCell
            ExecuteWrapperCellMode rng
        Case WrapModeTyped
            ExecuteTypedPatternMode ctx, rng
    End Select

CleanExit:
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Wrap.Execute", Err
    Resume CleanExit
End Sub

Private Sub ExecuteWrapperCellMode(ByVal targetRange As Range)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteWrapperCellMode")
    On Error GoTo ErrHandler

    Dim wrapperCell As Range
    Dim wrapperFormula As String
    Dim area As Range
    Dim cell As Range
    Dim totalErrors As Long

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

    If Not Infra_Undo.SaveStateOrConfirm(targetRange, COMMAND_NAME) Then GoTo CleanExit
    Application.StatusBar = "Wrapping " & Format(targetRange.CountLarge, "#,##0") & " cells using the selected wrapper cell..."

    For Each area In targetRange.Areas
        If area.CountLarge = 1 Then
            ApplyWrapperCellToCell area.Cells(1, 1), wrapperFormula, totalErrors
        Else
            For Each cell In area.Cells
                ApplyWrapperCellToCell cell, wrapperFormula, totalErrors
            Next cell
        End If
    Next area

    If totalErrors > 0 Then
        Infra_Interaction.ShowWarning "Completed with " & totalErrors & " error(s). Some cells could not be updated."
    End If

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ExecuteWrapperCellMode", Err
    Resume CleanExit
End Sub

Private Sub ExecuteTypedPatternMode(ByVal ctx As Infra_ActionContext, ByVal rng As Range)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteTypedPatternMode")
    On Error GoTo ErrHandler

    Dim formulaCells As Range
    Dim constantCells As Range
    Dim userPattern As String
    Dim totalErrors As Long

    userPattern = Infra_UIFactory.PromptForWrapFormulaPattern(ctx, PLACEHOLDER)
    If userPattern = vbNullString Then GoTo CleanExit

    If Not Infra_Undo.SaveStateOrConfirm(rng, COMMAND_NAME) Then GoTo CleanExit
    Application.StatusBar = "Wrapping " & Format(rng.CountLarge, "#,##0") & " cells with a typed formula pattern..."

    If rng.CountLarge = 1 Then
        If rng.HasFormula Then
            ApplyWrapPatternToRange rng, userPattern, True, totalErrors
        ElseIf Not IsEmpty(rng.Value) Then
            ApplyWrapPatternToRange rng, userPattern, False, totalErrors
        End If
    Else
        On Error Resume Next
        Set formulaCells = rng.SpecialCells(xlCellTypeFormulas)
        Set constantCells = rng.SpecialCells(xlCellTypeConstants)
        On Error GoTo ErrHandler

        If Not formulaCells Is Nothing Then ApplyWrapPatternToRange formulaCells, userPattern, True, totalErrors
        If Not constantCells Is Nothing Then ApplyWrapPatternToRange constantCells, userPattern, False, totalErrors
    End If

    If totalErrors > 0 Then
        Infra_Interaction.ShowInfo "Completed with " & totalErrors & " errors. Failed cells are highlighted in yellow."
    End If

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ExecuteTypedPatternMode", Err
    Resume CleanExit
End Sub

Private Sub ApplyWrapPatternToRange(ByVal targetCells As Range, ByVal userPattern As String, ByVal isFormulaRange As Boolean, ByRef errorCount As Long)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ApplyWrapPatternToRange")
    On Error GoTo ErrHandler

    Dim area As Range
    Dim formulas As Variant
    Dim values As Variant
    Dim r As Long
    Dim c As Long
    Dim oldContent As String
    Dim newContent As String
    Dim cell As Range

    For Each area In targetCells.Areas
        If area.CountLarge = 1 Then
            oldContent = GetWrappedContent(area.Formula2, VarType(area.Value) = vbString, isFormulaRange)
            newContent = Replace(userPattern, PLACEHOLDER, oldContent, , , vbTextCompare)
            If Left$(newContent, 1) <> "=" Then newContent = "=" & newContent

            On Error Resume Next
            area.Formula2 = newContent
            If Err.Number <> 0 Then
                area.Interior.Color = vbYellow
                errorCount = errorCount + 1
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            formulas = area.Formula2
            values = area.Value

            For r = 1 To UBound(formulas, 1)
                For c = 1 To UBound(formulas, 2)
                    If Not IsEmpty(formulas(r, c)) Then
                        oldContent = GetWrappedContent(formulas(r, c), VarType(values(r, c)) = vbString, isFormulaRange)
                        newContent = Replace(userPattern, PLACEHOLDER, oldContent, , , vbTextCompare)
                        If Left$(newContent, 1) <> "=" Then newContent = "=" & newContent
                        formulas(r, c) = newContent
                    End If
                Next c
            Next r

            On Error Resume Next
            area.Formula2 = formulas
            If Err.Number <> 0 Then
                Err.Clear
                For Each cell In area.Cells
                    If Not IsEmpty(cell.Value) Then
                        oldContent = GetWrappedContent(cell.Formula2, VarType(cell.Value) = vbString, isFormulaRange)
                        newContent = Replace(userPattern, PLACEHOLDER, oldContent, , , vbTextCompare)
                        If Left$(newContent, 1) <> "=" Then newContent = "=" & newContent

                        On Error Resume Next
                        cell.Formula2 = newContent
                        If Err.Number <> 0 Then
                            cell.Interior.Color = vbYellow
                            errorCount = errorCount + 1
                            Err.Clear
                        End If
                        On Error GoTo ErrHandler
                    End If
                Next cell
            End If
            On Error GoTo ErrHandler
        End If
    Next area

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ApplyWrapPatternToRange", Err
    Resume CleanExit
End Sub

Private Sub ApplyWrapperCellToCell(ByVal targetCell As Range, ByVal wrapperFormula As String, ByRef errorCount As Long)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ApplyWrapperCellToCell")
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
    Infra_Error.HandleError "ApplyWrapperCellToCell", Err
    errorCount = errorCount + 1
    Resume CleanExit
End Sub

Private Function ReplaceFirstCellReference(ByVal formulaText As String, ByVal replacementText As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ReplaceFirstCellReference")
    On Error GoTo ErrHandler

    Dim startIndex As Long
    Dim tokenLength As Long

    Call FindFirstReferenceToken(formulaText, startIndex, tokenLength)
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

    Set re = GetRegExp()
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

Private Function GetWrappedContent(ByVal content As Variant, ByVal isStringValue As Boolean, ByVal isFormula As Boolean) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetWrappedContent")
    On Error GoTo ErrHandler

    If IsError(content) Then
        Select Case content
            Case CVErr(xlErrDiv0): GetWrappedContent = "#DIV/0!"
            Case CVErr(xlErrNull): GetWrappedContent = "#NULL!"
            Case CVErr(xlErrNA): GetWrappedContent = "#N/A"
            Case CVErr(xlErrName): GetWrappedContent = "#NAME?"
            Case CVErr(xlErrNum): GetWrappedContent = "#NUM!"
            Case CVErr(xlErrRef): GetWrappedContent = "#REF!"
            Case CVErr(xlErrValue): GetWrappedContent = "#VALUE!"
            Case Else: GetWrappedContent = "#ERROR!"
        End Select
        Exit Function
    End If

    Dim result As String
    result = CStr(content)

    If isFormula Then
        If Left$(result, 1) = "=" Then result = Mid$(result, 2)
        GetWrappedContent = "(" & result & ")"
    Else
        If isStringValue Then
            GetWrappedContent = """" & Replace(result, """", """""") & """"
        Else
            GetWrappedContent = "(" & result & ")"
        End If
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "GetWrappedContent", Err
    Resume CleanExit
End Function

Private Function GetRegExp() As Object
    If m_RegExp Is Nothing Then
        Set m_RegExp = CreateObject("VBScript.RegExp")
        m_RegExp.Global = True
        m_RegExp.IgnoreCase = True
        m_RegExp.MultiLine = False
        m_RegExp.Pattern = "(?:'[^']+'!|[A-Za-z0-9_\.]+!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?:#)?"
    End If
    Set GetRegExp = m_RegExp
End Function

Private Sub Class_Terminate()
    Set m_RegExp = Nothing
End Sub

' --- Testing Entry Points ---
Public Sub TestApplyWrapPatternDirect(ByVal targetCells As Range, ByVal userPattern As String, ByVal isFormulaRange As Boolean, ByRef errorCount As Long)
    Dim tracker As Object: Set tracker = Infra_Error.Track("TestApplyWrapPatternDirect")
    On Error GoTo ErrHandler

    ApplyWrapPatternToRange targetCells, userPattern, isFormulaRange, errorCount

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "TestApplyWrapPatternDirect", Err
    Resume CleanExit
End Sub

Public Sub TestApplyWrapperCellDirect(ByVal targetCell As Range, ByVal wrapperFormula As String, ByRef errorCount As Long)
    Dim tracker As Object: Set tracker = Infra_Error.Track("TestApplyWrapperCellDirect")
    On Error GoTo ErrHandler

    ApplyWrapperCellToCell targetCell, wrapperFormula, errorCount

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "TestApplyWrapperCellDirect", Err
    Resume CleanExit
End Sub
