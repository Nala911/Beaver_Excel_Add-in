Attribute VB_Name = "Feat_WrapSelectedRange"
Option Explicit

' @Module: Feat_WrapSelectedRange
' @Category: Feature
' @Description: Wraps selected cell formulas with a user-supplied function (e.g., IFERROR).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_AppStateGuard, Infra_Config, Infra_Error, Infra_UIFactory

Private Const PLACEHOLDER As String = "[value]"
Private Const REG_SECTION As String = "Settings"
Private Const REG_KEY_PATTERN As String = "WrapFormulaPattern"
Private Const MAX_CELLS_LIMIT As Long = 50000

Public Sub WrapSelectionWithFormula()
    PushContext "WrapSelectionWithFormula"
    On Error GoTo ErrHandler

    Dim guard As Infra_AppStateGuard
    Dim rng As Range
    Dim formulaCells As Range
    Dim constantCells As Range
    Dim userPattern As String
    Dim lastPattern As String
    Dim totalErrors As Long
    Dim ctx As Infra_ActionContext

    If Not Infra_AppState.IsRangeSelected() Then GoTo CleanExit
    Set rng = Selection
    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit
    
    ' 1. Guard against massive selections
    If rng.CountLarge > MAX_CELLS_LIMIT Then
        If MsgBox("You have selected " & Format(rng.CountLarge, "#,##0") & " cells. " & _
                  "This may take a moment to process. Continue?", _
                  vbQuestion + vbYesNo, Infra_Config.ADDIN_NAME) = vbNo Then GoTo CleanExit
    End If

    ' 2. Retrieve last used pattern or default
    lastPattern = GetSetting(Infra_Config.ADDIN_NAME, REG_SECTION, REG_KEY_PATTERN, PLACEHOLDER)

    ' 3. Prompt for the wrapper formula pattern
    userPattern = Infra_UIFactory.PromptForWrapFormulaPattern(ctx, lastPattern, PLACEHOLDER)
    If userPattern = vbNullString Then GoTo CleanExit
    
    ' 5. Save pattern for next time
    SaveSetting Infra_Config.ADDIN_NAME, REG_SECTION, REG_KEY_PATTERN, userPattern

    ' 6. Process with screen/calc guards
    Set guard = New Infra_AppStateGuard
    Application.StatusBar = "Wrapping formulas in " & Format(rng.CountLarge, "#,##0") & " cells..."

    If rng.CountLarge = 1 Then
        If rng.HasFormula Then
            ApplyWrapPatternToRange rng, userPattern, True, totalErrors
        ElseIf Not IsEmpty(rng.Value) Then
            ApplyWrapPatternToRange rng, userPattern, False, totalErrors
        End If
    Else
        ' Multiple cells: use SpecialCells to target only relevant content
        On Error Resume Next
        Set formulaCells = rng.SpecialCells(xlCellTypeFormulas)
        Set constantCells = rng.SpecialCells(xlCellTypeConstants)
        On Error GoTo ErrHandler

        If Not formulaCells Is Nothing Then ApplyWrapPatternToRange formulaCells, userPattern, True, totalErrors
        If Not constantCells Is Nothing Then ApplyWrapPatternToRange constantCells, userPattern, False, totalErrors
    End If

    ' 7. Summary report
    If totalErrors > 0 Then
        MsgBox "Completed with " & totalErrors & " errors. Failed cells are highlighted in yellow.", _
               vbInformation, Infra_Config.ADDIN_NAME
    End If

CleanExit:
    Application.StatusBar = False
    PopContext
    Exit Sub

ErrHandler:
    HandleError "WrapSelectionWithFormula", Err
    Resume CleanExit
End Sub

Private Sub ApplyWrapPatternToRange(ByVal targetCells As Range, ByVal userPattern As String, ByVal isFormulaRange As Boolean, ByRef errorCount As Long)
    PushContext "ApplyWrapPatternToRange"
    On Error GoTo ErrHandler
    
    Dim area As Range
    Dim formulas As Variant
    Dim values As Variant
    Dim r As Long, c As Long
    Dim oldContent As String
    Dim newContent As String
    
    For Each area In targetCells.Areas
        If area.CountLarge = 1 Then
            ' Single cell - direct update
            oldContent = GetWrappedContent(area.Formula2, VarType(area.Value) = vbString, isFormulaRange)
            newContent = Replace(userPattern, PLACEHOLDER, oldContent, , , vbTextCompare)
            If Left$(newContent, 1) <> "=" Then newContent = "=" & newContent
            
            On Error Resume Next
            area.Formula2 = newContent
            If Err.Number <> 0 Then
                area.Interior.Color = vbYellow
                errorCount = errorCount + 1
            End If
            On Error GoTo ErrHandler
        Else
            ' Multi-cell area - batch update via arrays for speed
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
            
            ' Write back to range in one operation
            On Error Resume Next
            area.Formula2 = formulas
            If Err.Number <> 0 Then
                ' If batch write fails (e.g. one formula is invalid), fallback to cell-by-cell for this area
                Err.Clear
                Dim cell As Range
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
    PopContext
    Exit Sub

ErrHandler:
    HandleError "ApplyWrapPatternToRange", Err
End Sub

Private Function GetWrappedContent(ByVal content As Variant, ByVal isStringValue As Boolean, ByVal isFormula As Boolean) As String
    PushContext "GetWrappedContent"
    On Error GoTo ErrHandler
    
    Dim result As String
    result = CStr(content)

    If isFormula Then
        ' Remove leading "=" if present for formulas
        If Left$(result, 1) = "=" Then result = Mid$(result, 2)
        GetWrappedContent = "(" & result & ")"
    Else
        ' Escape quotes for string constants; wrap others in parens
        If isStringValue Then
            GetWrappedContent = """" & Replace(result, """", """""") & """"
        Else
            GetWrappedContent = "(" & result & ")"
        End If
    End If

CleanExit:
    PopContext
    Exit Function

ErrHandler:
    HandleError "GetWrappedContent", Err
End Function
