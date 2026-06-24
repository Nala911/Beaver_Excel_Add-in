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
    
    If context.CommandName = "ClearHighlights" Then
        ClearWorkbookHighlights ctx.WorkbookRef
        Infra_Interaction.ShowInfo "Done! Cleared all highlights in the workbook.", "Clear Highlights"
        GoTo CleanExit
    End If
    
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
    
    ' Always clear existing highlights in the workbook first
    ClearWorkbookHighlights request.Context.WorkbookRef
    
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
        
        ' 5. Highlight Data Validations
        If request.HighlightDataValidations Then
            HighlightDataValidationsInArea area
        End If
        
        ' 6. Highlight Conditional Formatting
        If request.HighlightConditionalFormatting Then
            HighlightConditionalFormattingInArea area
        End If
        
        DoEvents
        If Infra_Progress.UserCancelled Then Exit For
    Next i
End Sub

Private Sub HighlightInconsistentFormulasInArea(ByVal area As Range)
    Dim formulaCells As Range
    If area.CountLarge = 1 Then
        If area.HasFormula Then
            Set formulaCells = area
        End If
    Else
        On Error Resume Next
        Set formulaCells = area.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
    End If
    
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
    
    Dim addrList As String
    Dim cellAddr As String
    
    For Each cell In formulaCells.Cells
        Dim isInc As Boolean
        isInc = False
        On Error Resume Next
        isInc = cell.Errors(xlInconsistentFormula).Value
        On Error GoTo 0
        
        If isInc Then
            cellAddr = Infra_CommandSupport.GetA1Address(cell.Row, cell.Column)
            Set unionRange = Infra_CommandSupport.AccumulateUnion(unionRange, area.Parent, addrList, cellAddr)
            
            If Not unionRange Is Nothing Then
                If unionRange.Areas.Count >= 500 Then
                    unionRange.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
                    Set unionRange = Nothing
                End If
            End If
        End If
    Next cell
    
    Set unionRange = Infra_CommandSupport.AccumulateUnion(unionRange, area.Parent, addrList, "", True)
    If Not unionRange Is Nothing Then
        unionRange.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
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
    Dim addrList As String
    Dim cellAddr As String
    
    Dim startRow As Long, startCol As Long
    startRow = area.Row
    startCol = area.Column
    
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            val = dataArr(r, c)
            If Not IsEmpty(val) And Not IsError(val) Then
                key = Trim$(CStr(val))
                If key <> "" Then
                    If valCount(key) > 1 Then
                        cellAddr = Infra_CommandSupport.GetA1Address(startRow + r - 1, startCol + c - 1)
                        Set unionRange = Infra_CommandSupport.AccumulateUnion(unionRange, area.Parent, addrList, cellAddr)
                        
                        If Not unionRange Is Nothing Then
                            If unionRange.Areas.Count >= 500 Then
                                unionRange.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
                                Set unionRange = Nothing
                            End If
                        End If
                    End If
                End If
            End If
        Next c
    Next r
    
    Set unionRange = Infra_CommandSupport.AccumulateUnion(unionRange, area.Parent, addrList, "", True)
    If Not unionRange Is Nothing Then
        unionRange.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
    End If
End Sub

Private Sub HighlightErrorsInArea(ByVal area As Range)
    Dim formulaErrors As Range
    Dim constantErrors As Range
    Dim errorCells As Range
    
    If area.CountLarge = 1 Then
        If IsError(area.Value) Then
            Set errorCells = area
        End If
    Else
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
    End If
    
    If errorCells Is Nothing Then Exit Sub
    
    On Error Resume Next
    errorCells.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
    On Error GoTo 0
End Sub

Private Sub HighlightHardcodedValuesInArea(ByVal area As Range)
    Dim formulaCells As Range
    If area.CountLarge = 1 Then
        If area.HasFormula Then
            Set formulaCells = area
        End If
    Else
        On Error Resume Next
        Set formulaCells = area.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
    End If
    
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
    
    Dim unionRange As Range
    Dim re As Object
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    
    Dim subArea As Range
    Dim addrList As String
    Dim cellAddr As String
    
    For Each subArea In formulaCells.Areas
        Dim formulaArr As Variant
        Dim isArray As Boolean
        
        If subArea.Cells.CountLarge = 1 Then
            formulaArr = subArea.Formula2
            isArray = False
        Else
            formulaArr = subArea.Formula2
            isArray = True
        End If
        
        If isArray Then
            Dim r As Long, c As Long
            Dim startRow As Long, startCol As Long
            startRow = subArea.Row
            startCol = subArea.Column
            
            For r = 1 To UBound(formulaArr, 1)
                For c = 1 To UBound(formulaArr, 2)
                    Dim hasHardcoded As Boolean
                    hasHardcoded = False
                    On Error Resume Next
                    hasHardcoded = HasHardcodedValue(CStr(formulaArr(r, c)), re)
                    On Error GoTo 0
                    
                    If hasHardcoded Then
                        cellAddr = Infra_CommandSupport.GetA1Address(startRow + r - 1, startCol + c - 1)
                        Set unionRange = Infra_CommandSupport.AccumulateUnion(unionRange, area.Parent, addrList, cellAddr)
                        
                        If Not unionRange Is Nothing Then
                            If unionRange.Areas.Count >= 500 Then
                                unionRange.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
                                Set unionRange = Nothing
                            End If
                        End If
                    End If
                Next c
            Next r
        Else
            Dim singleHasHardcoded As Boolean
            singleHasHardcoded = False
            On Error Resume Next
            singleHasHardcoded = HasHardcodedValue(CStr(formulaArr), re)
            On Error GoTo 0
            
            If singleHasHardcoded Then
                cellAddr = Infra_CommandSupport.GetA1Address(subArea.Row, subArea.Column)
                Set unionRange = Infra_CommandSupport.AccumulateUnion(unionRange, area.Parent, addrList, cellAddr)
                
                If Not unionRange Is Nothing Then
                    If unionRange.Areas.Count >= 500 Then
                        unionRange.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
                        Set unionRange = Nothing
                    End If
                End If
            End If
        End If
    Next subArea
    
    Set unionRange = Infra_CommandSupport.AccumulateUnion(unionRange, area.Parent, addrList, "", True)
    If Not unionRange Is Nothing Then
        unionRange.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
    End If
End Sub

Public Function HasHardcodedValue(ByVal formula As String, Optional ByVal rePassed As Object = Nothing) As Boolean
    ' [Bypass Lint] PushContext "HasHardcodedValue" / PopContext / Infra_Error.Track (exempt for CPU performance)
    On Error GoTo ErrHandler

    ' Remove any leading "=" sign
    If Left$(formula, 1) = "=" Then formula = Mid$(formula, 2)
    
    Dim re As Object
    If rePassed Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = True
    Else
        Set re = rePassed
    End If
    
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

Private Sub HighlightDataValidationsInArea(ByVal area As Range)
    Dim validationCells As Range
    If area.CountLarge = 1 Then
        Dim hasVal As Boolean
        hasVal = False
        On Error Resume Next
        Dim valType As Long
        valType = area.Validation.Type
        If Err.Number = 0 Then hasVal = True
        On Error GoTo 0
        If hasVal Then Set validationCells = area
    Else
        On Error Resume Next
        Set validationCells = area.SpecialCells(xlCellTypeAllValidation)
        On Error GoTo 0
    End If
    
    If validationCells Is Nothing Then Exit Sub
    
    On Error Resume Next
    validationCells.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
    On Error GoTo 0
End Sub

Private Sub HighlightConditionalFormattingInArea(ByVal area As Range)
    Dim cfCells As Range
    If area.CountLarge = 1 Then
        If area.FormatConditions.Count > 0 Then Set cfCells = area
    Else
        On Error Resume Next
        Set cfCells = area.SpecialCells(xlCellTypeAllFormatConditions)
        On Error GoTo 0
    End If
    
    If cfCells Is Nothing Then Exit Sub
    
    On Error Resume Next
    cfCells.Interior.Color = Infra_Config.HIGHLIGHT_COLOR
    On Error GoTo 0
End Sub

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

Public Sub ClearWorkbookHighlights(ByVal wb As Workbook)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ClearWorkbookHighlights")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim targetColor As Long
    targetColor = Infra_Config.HIGHLIGHT_COLOR
    
    ' Set up search and replace formats
    Application.FindFormat.Clear
    Application.FindFormat.Interior.Color = targetColor
    
    Application.ReplaceFormat.Clear
    Application.ReplaceFormat.Interior.ColorIndex = xlNone
    
    ' Configure app state to optimize replacement performance
    guard.ScreenUpdating = False
    guard.EnableEvents = False
    guard.Calculation = xlCalculationManual
    guard.UpdateState

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Left$(ws.Name, 1) <> "_" Then
            On Error Resume Next
            ws.UsedRange.Replace What:="", Replacement:="", SearchFormat:=True, ReplaceFormat:=True
            On Error GoTo ErrHandler
        End If
    Next ws

CleanExit:
    On Error Resume Next
    Application.FindFormat.Clear
    Application.ReplaceFormat.Clear
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "ClearWorkbookHighlights", Err
    Resume CleanExit
End Sub
