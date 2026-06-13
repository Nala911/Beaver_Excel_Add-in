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

' @Module: FeatCmd_HighlightData
' @Category: Feature
' @Description: Advanced data highlighting tools (Inconsistent Formulas, Duplicates).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, Infra_ActionContext, Infra_HighlightDataRequest, Infra_Error, Infra_Undo, Infra_Progress, Infra_Interaction

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

    Set request = Infra_UIFactory.ShowHighlightDataDialog(ctx)
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
    
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim stateInfo As Collection
    Dim scopeLabel As String
    Dim originalScreenUpdating As Boolean

    Set stateInfo = Infra_Error.NewStateInfo()
    
    If request Is Nothing Then GoTo CleanExit
    If request.Context Is Nothing Then GoTo CleanExit
    Infra_Error.AddState stateInfo, "Scope", request.Scope
    
    ' --- 1. Save State for Undo ---
    Select Case request.Scope
        Case TargetScopeSelection
            Set targetRange = request.Context.SelectionRange
            scopeLabel = "selection"
        Case TargetScopeActiveSheet
            Set targetRange = request.Context.WorksheetRef.UsedRange
            scopeLabel = "active sheet"
        Case TargetScopeWorkbook
            Set targetRange = Nothing
            scopeLabel = "workbook"
    End Select
    
    If Not targetRange Is Nothing Then
        If Not Infra_Undo.SaveStateOrConfirm(targetRange, "Highlight Data") Then GoTo CleanExit
    End If
    
    ' --- 2. Initialize Progress & Dynamic Screen Updating ---
    originalScreenUpdating = Application.ScreenUpdating
    If request.Scope = TargetScopeWorkbook Then
        Application.ScreenUpdating = False
        Infra_Progress.StartProgress "Highlighting Data", request.Context.WorkbookRef.Worksheets.Count
    Else
        Infra_Progress.StartProgress "Highlighting Data", 1
    End If
    
    Select Case request.Scope
        Case TargetScopeSelection
            ProcessRangeForHighlight request.Context.SelectionRange, request
            Infra_Progress.UpdateProgress 1, True
        Case TargetScopeActiveSheet
            ProcessRangeForHighlight request.Context.WorksheetRef.UsedRange, request
            Infra_Progress.UpdateProgress 1, True
        Case TargetScopeWorkbook
            Dim i As Long, count As Long
            count = request.Context.WorkbookRef.Worksheets.Count
            For i = 1 To count
                Set ws = request.Context.WorkbookRef.Worksheets(i)
                ProcessRangeForHighlight ws.UsedRange, request
                Infra_Progress.UpdateProgress i, True
                If Infra_Progress.UserCancelled Then Exit For
            Next i
    End Select
    
    Dim msg As String
    msg = "Done! Highlighted patterns and duplicates in the " & scopeLabel & "."
    
    Infra_Interaction.ShowInfo msg

CleanExit:
    If Not request Is Nothing Then
        If request.Scope = TargetScopeWorkbook Then
            Application.ScreenUpdating = originalScreenUpdating
        End If
    End If
    Infra_Progress.EndProgress
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ExecuteHighlightData", Err, stateInfo
    Resume CleanExit
End Sub

Private Sub ProcessRangeForHighlight(ByRef targetRange As Range, ByVal request As Infra_HighlightDataRequest)
    If targetRange Is Nothing Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = targetRange.Worksheet
    
    ' Intersect with UsedRange if targetRange is large to prevent scanning millions of empty cells
    If targetRange.Cells.CountLarge > 100000 Then
        Dim usedRange As Range
        On Error Resume Next
        Set usedRange = Application.Intersect(targetRange, ws.UsedRange)
        On Error GoTo 0
        If Not usedRange Is Nothing Then
            Set targetRange = usedRange
        Else
            Exit Sub
        End If
    End If
    
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
