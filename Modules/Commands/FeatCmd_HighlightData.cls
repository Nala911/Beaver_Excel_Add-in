VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_HighlightData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand
Implements IBatchProcessor

' @Module: FeatCmd_HighlightData
' @Category: Feature
' @Description: Advanced data highlighting tools (Inconsistent Formulas, Duplicates).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, Infra_ActionContext, Infra_HighlightDataRequest, Infra_Error, Infra_Undo, Infra_Progress, Infra_Interaction

Private m_CurrentRequest As Infra_HighlightDataRequest

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorksheet(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("HighlightData.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As Infra_ActionContext
    Dim request As Infra_HighlightDataRequest
    Dim stateInfo As Collection

    Set stateInfo = Infra_Error.NewStateInfo()

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Or ctx.WorkbookRef Is Nothing Then GoTo CleanExit
    
    Infra_Error.AddState stateInfo, "Workbook", ctx.WorkbookRef.Name
    Infra_Error.AddState stateInfo, "Worksheet", ctx.WorksheetRef.Name
    If ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing Then
        Infra_Error.AddState stateInfo, "Selection", ctx.SelectionRange.Address(False, False)
    End If

    Set request = Infra_UIFactory.ShowHighlightDataDialog(ctx, context.CommandName)
    If request Is Nothing Then GoTo CleanExit
    
    ExecuteHighlightData request

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "HighlightData.Execute", Err, stateInfo
    Resume CleanExit
End Sub

Private Sub ExecuteHighlightData(ByVal request As Infra_HighlightDataRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteHighlightData")
    On Error GoTo ErrHandler
    
    If request Is Nothing Then GoTo CleanExit
    If request.Context Is Nothing Then GoTo CleanExit
    
    Set m_CurrentRequest = request
    
    Dim totalProcessed As Long
    totalProcessed = Infra_BatchProcessor.Execute(request.Context, request.Scope, Me, "Highlight Data", False)
    
    Dim scopeLabel As String
    Select Case request.Scope
        Case TargetScopeSelection: scopeLabel = "selection"
        Case TargetScopeActiveSheet: scopeLabel = "active sheet"
        Case TargetScopeWorkbook: scopeLabel = "workbook"
    End Select
    
    Dim msg As String
    msg = "Done! Highlighted patterns and duplicates in the " & scopeLabel & "."
    Infra_Interaction.ShowInfo msg, "Highlight Data"

CleanExit:
    Set m_CurrentRequest = Nothing
    Exit Sub

ErrHandler:
    Set m_CurrentRequest = Nothing
    Infra_Error.HandleError "ExecuteHighlightData", Err
    Resume CleanExit
End Sub

Private Function IBatchProcessor_ProcessRange(ByVal targetRange As Range) As Long
    ProcessRangeForHighlight targetRange, m_CurrentRequest
    IBatchProcessor_ProcessRange = 0
End Function

Private Function IBatchProcessor_ProcessWorksheet(ByVal ws As Worksheet) As Long
    IBatchProcessor_ProcessWorksheet = 0
End Function

Private Sub ProcessRangeForHighlight(ByRef targetRange As Range, ByVal request As Infra_HighlightDataRequest)
    If targetRange Is Nothing Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = targetRange.Worksheet
    
    ' Intersect with UsedRange if targetRange is large to prevent scanning millions of empty cells
    Set targetRange = Infra_CommandSupport.GetSafeProcessingRange(targetRange, 100000, False)
    If targetRange Is Nothing Then Exit Sub
    
    Dim area As Range
    Dim totalAreas As Long
    Dim i As Long

    totalAreas = targetRange.Areas.Count
    For i = 1 To totalAreas
        Set area = targetRange.Areas(i)
        
        ' 1. Highlight Inconsistent Formulas
        If request.HighlightInconsistentFormulas Then
            HighlightInconsistentFormulasInArea area
        End If
        
        ' 2. Highlight Duplicates
        If request.HighlightDuplicates Then
            HighlightDuplicatesInArea area
        End If
        
        ' 3. Highlight Errors
        If request.HighlightErrors Then
            HighlightErrorsInArea area
        End If
        
        ' 4. Highlight Hardcoded Values
        If request.HighlightHardcodedValues Then
            HighlightHardcodedValuesInArea area
        End If
        
        DoEvents
        If Infra_Progress.UserCancelled Then Exit For
    Next i
End Sub

Private Sub HighlightInconsistentFormulasInArea(ByVal area As Range)
    Dim formulaCells As Range
    On Error Resume Next
    Set formulaCells = area.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If formulaCells Is Nothing Then Exit Sub
    
    ' Enforce MaxFormulaCheckCells safety threshold
    Dim maxCells As Long
    maxCells = Infra_Config.MAX_FORMULA_CHECK_CELLS
    If formulaCells.Cells.CountLarge > maxCells Then
        Dim msg As String
        msg = "The number of formula cells in the area (" & formulaCells.Cells.CountLarge & _
              ") exceeds the safety limit of " & maxCells & " cells. " & _
              "Inconsistent formula checking was skipped to avoid freezing Excel."
        Infra_Interaction.ShowWarning msg, "Highlight Data"
        Exit Sub
    End If
    
    Dim cell As Range
    Dim unionRange As Range
    
    For Each cell In formulaCells.Cells
        Dim isInc As Boolean
        isInc = False
        On Error Resume Next
        isInc = cell.Errors(xlInconsistentFormula).Value
        On Error GoTo 0
        
        If isInc Then
            If unionRange Is Nothing Then
                Set unionRange = cell
            Else
                Set unionRange = Application.Union(unionRange, cell)
            End If
            
            If unionRange.Areas.Count >= 500 Then
                unionRange.Interior.Color = RGB(255, 255, 153) ' Soft Yellow
                Set unionRange = Nothing
            End If
        End If
    Next cell
    
    If Not unionRange Is Nothing Then
        unionRange.Interior.Color = RGB(255, 255, 153) ' Soft Yellow
    End If
End Sub

Private Sub HighlightDuplicatesInArea(ByVal area As Range)
    If area.Cells.CountLarge <= 1 Then Exit Sub
    
    Dim dataArr As Variant
    dataArr = area.Value2
    
    If IsEmpty(dataArr) Then Exit Sub
    
    Dim valCount As Object
    Set valCount = CreateObject("Scripting.Dictionary")
    valCount.CompareMode = 1 ' Case-insensitive comparison
    
    Dim r As Long, c As Long
    Dim val As Variant
    Dim key As String
    
    ' First pass: count occurrences
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            val = dataArr(r, c)
            If Not IsEmpty(val) And Not IsError(val) Then
                key = Trim$(CStr(val))
                If key <> "" Then
                    If valCount.Exists(key) Then
                        valCount(key) = valCount(key) + 1
                    Else
                        valCount(key) = 1
                    End If
                End If
            End If
        Next c
    Next r
    
    ' Second pass: highlight cells with count > 1
    Dim unionRange As Range
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            val = dataArr(r, c)
            If Not IsEmpty(val) And Not IsError(val) Then
                key = Trim$(CStr(val))
                If key <> "" Then
                    If valCount(key) > 1 Then
                        If unionRange Is Nothing Then
                            Set unionRange = area.Cells(r, c)
                        Else
                            Set unionRange = Application.Union(unionRange, area.Cells(r, c))
                        End If
                        
                        If unionRange.Areas.Count >= 500 Then
                            unionRange.Interior.Color = RGB(255, 199, 206) ' Soft Red/Pink
                            Set unionRange = Nothing
                        End If
                    End If
                End If
            End If
        Next c
    Next r
    
    If Not unionRange Is Nothing Then
        unionRange.Interior.Color = RGB(255, 199, 206) ' Soft Red/Pink
    End If
End Sub

Private Sub HighlightErrorsInArea(ByVal area As Range)
    Dim formulaErrors As Range
    Dim constantErrors As Range
    Dim errorCells As Range
    
    On Error Resume Next
    Set formulaErrors = area.SpecialCells(xlCellTypeFormulas, xlErrors)
    Set constantErrors = area.SpecialCells(xlCellTypeConstants, xlErrors)
    On Error GoTo 0
    
    If Not formulaErrors Is Nothing And Not constantErrors Is Nothing Then
        Set errorCells = Application.Union(formulaErrors, constantErrors)
    ElseIf Not formulaErrors Is Nothing Then
        Set errorCells = formulaErrors
    ElseIf Not constantErrors Is Nothing Then
        Set errorCells = constantErrors
    End If
    
    If errorCells Is Nothing Then Exit Sub
    
    Dim subArea As Range
    Dim unionRange As Range
    
    For Each subArea In errorCells.Areas
        If unionRange Is Nothing Then
            Set unionRange = subArea
        Else
            Set unionRange = Application.Union(unionRange, subArea)
        End If
        
        If unionRange.Areas.Count >= 500 Then
            unionRange.Interior.Color = RGB(255, 204, 153) ' Soft Orange
            Set unionRange = Nothing
        End If
    Next subArea
    
    If Not unionRange Is Nothing Then
        unionRange.Interior.Color = RGB(255, 204, 153) ' Soft Orange
    End If
End Sub

Private Sub HighlightHardcodedValuesInArea(ByVal area As Range)
    Dim formulaCells As Range
    On Error Resume Next
    Set formulaCells = area.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If formulaCells Is Nothing Then Exit Sub
    
    ' Enforce MaxFormulaCheckCells safety threshold
    Dim maxCells As Long
    maxCells = Infra_Config.MAX_FORMULA_CHECK_CELLS
    If formulaCells.Cells.CountLarge > maxCells Then
        Dim msg As String
        msg = "The number of formula cells in the area (" & formulaCells.Cells.CountLarge & _
              ") exceeds the safety limit of " & maxCells & " cells. " & _
              "Hardcoded value checking was skipped to avoid freezing Excel."
        Infra_Interaction.ShowWarning msg, "Highlight Data"
        Exit Sub
    End If
    
    Dim cell As Range
    Dim unionRange As Range
    
    For Each cell In formulaCells.Cells
        Dim hasHardcoded As Boolean
        hasHardcoded = False
        On Error Resume Next
        hasHardcoded = HasHardcodedValue(CStr(cell.Formula2))
        On Error GoTo 0
        
        If hasHardcoded Then
            If unionRange Is Nothing Then
                Set unionRange = cell
            Else
                Set unionRange = Application.Union(unionRange, cell)
            End If
            
            If unionRange.Areas.Count >= 500 Then
                unionRange.Interior.Color = RGB(230, 210, 250) ' Soft Lavender/Purple
                Set unionRange = Nothing
            End If
        End If
    Next cell
    
    If Not unionRange Is Nothing Then
        unionRange.Interior.Color = RGB(230, 210, 250) ' Soft Lavender/Purple
    End If
End Sub

Public Function HasHardcodedValue(ByVal formula As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("HasHardcodedValue")
    On Error GoTo ErrHandler

    ' Remove any leading "=" sign
    If Left$(formula, 1) = "=" Then formula = Mid$(formula, 2)
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    
    ' 1. Check for non-empty text literals: "[^"]+" (excluding empty string "")
    re.Pattern = """([^""]+)"""
    If re.Test(formula) Then
        HasHardcodedValue = True
        GoTo CleanExit
    End If
    
    ' Replace all string literals with empty string to avoid parsing numbers/references in strings
    re.Pattern = """[^""]*"""
    formula = re.Replace(formula, "")
    
    ' 2. Strip table column/field references [ColumnName] or [[Header],[ColumnName]]
    re.Pattern = "\[[^\]]*\]"
    formula = re.Replace(formula, "")
    
    ' 3. Strip sheet references: 'Sheet Name'! or SheetName!
    re.Pattern = "'[^']*'!|[A-Za-z0-9_\.]+!"
    formula = re.Replace(formula, "")
    
    ' 4. Strip standard cell references: e.g. A1, $B$10, AA100, A1:B10, A:B, 1:2
    re.Pattern = "\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?"
    formula = re.Replace(formula, "")
    re.Pattern = "\$?[A-Z]{1,3}:\$?[A-Z]{1,3}"
    formula = re.Replace(formula, "")
    re.Pattern = "\$?\d+:\$?\d+"
    formula = re.Replace(formula, "")
    
    ' 5. Look for standalone numbers.
    ' Word boundary \b ensures it's a standalone token.
    re.Pattern = "\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b"
    
    Dim matches As Object
    Set matches = re.Execute(formula)
    If matches.Count > 0 Then
        ' Exclude 0 and 1 as they are common offsets/flags in MATCH, LEFT, VLOOKUP, etc.
        Dim match As Object
        For Each match In matches
            Dim val As String
            val = match.Value
            If val <> "0" And val <> "1" Then
                HasHardcodedValue = True
                GoTo CleanExit
            End If
        Next match
    End If
    
    HasHardcodedValue = False

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "HasHardcodedValue", Err
    Resume CleanExit
End Function

' --- Testing Entry Points ---
Public Sub HighlightRangeDirect(ByVal targetRange As Range)
    Dim tracker As Object: Set tracker = Infra_Error.Track("HighlightRangeDirect")
    On Error GoTo ErrHandler
    
    Dim request As New Infra_HighlightDataRequest
    ProcessRangeForHighlight targetRange, request

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "HighlightRangeDirect", Err
    Resume CleanExit
End Sub

Public Sub HighlightRangeWithOptionsDirect(ByVal targetRange As Range, ByVal request As Infra_HighlightDataRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("HighlightRangeWithOptionsDirect")
    On Error GoTo ErrHandler
    
    ProcessRangeForHighlight targetRange, request

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "HighlightRangeWithOptionsDirect", Err
    Resume CleanExit
End Sub
