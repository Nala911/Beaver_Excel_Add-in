VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_CleanData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_CleanData
' @Category: Feature
' @Description: Advanced data cleaning tools (Trim, Clean).

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorksheet(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanData.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As Infra_ActionContext
    Dim request As Infra_CleanDataRequest
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

    Set request = Infra_UIFactory.ShowCleanDataDialog(ctx)
    If request Is Nothing Then GoTo CleanExit
    
    ExecuteCleanData request

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "CleanData.Execute", Err, stateInfo
    Resume CleanExit
End Sub

Private Sub ExecuteCleanData(ByVal request As Infra_CleanDataRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteCleanData")
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim stateInfo As Collection
    Dim cleanedCount As Long
    Dim scopeLabel As String
    Dim highlightedAny As Boolean
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
        If Not Infra_Undo.SaveStateOrConfirm(targetRange, "Clean Data") Then GoTo CleanExit
    End If
    
    ' --- 2. Initialize Progress & Dynamic Screen Updating ---
    originalScreenUpdating = Application.ScreenUpdating
    If request.Scope = TargetScopeWorkbook Then
        Application.ScreenUpdating = False
        Infra_Progress.StartProgress "Cleaning Data", request.Context.WorkbookRef.Worksheets.Count
    Else
        Infra_Progress.StartProgress "Cleaning Data", 1
    End If
    
    Select Case request.Scope
        Case TargetScopeSelection
            cleanedCount = ProcessRangeForClean(request.Context.SelectionRange, request)
            If request.CleanInconsistentFormulas Then
                highlightedAny = HighlightInconsistentFormulas(request.Context.SelectionRange)
            End If
            Infra_Progress.UpdateProgress 1, True
        Case TargetScopeActiveSheet
            cleanedCount = ProcessRangeForClean(request.Context.WorksheetRef.UsedRange, request)
            If request.CleanBrokenNames Then
                CleanNamedRanges request.Context.WorkbookRef, request.Context.WorksheetRef
            End If
            If request.CleanInconsistentFormulas Then
                highlightedAny = HighlightInconsistentFormulas(request.Context.WorksheetRef.UsedRange)
            End If
            Infra_Progress.UpdateProgress 1, True
        Case TargetScopeWorkbook
            If request.CleanBrokenNames Then
                CleanNamedRanges request.Context.WorkbookRef, Nothing
            End If
            Dim i As Long, count As Long
            count = request.Context.WorkbookRef.Worksheets.Count
            For i = 1 To count
                Set ws = request.Context.WorkbookRef.Worksheets(i)
                cleanedCount = cleanedCount + ProcessRangeForClean(ws.UsedRange, request)
                If request.CleanInconsistentFormulas Then
                    If HighlightInconsistentFormulas(ws.UsedRange) Then highlightedAny = True
                End If
                Infra_Progress.UpdateProgress i, True
                If Infra_Progress.UserCancelled Then Exit For
            Next i
    End Select
    
    Dim msg As String
    msg = "Done! Cleaned text, fixed text-numbers, and removed broken names in the " & scopeLabel & "."
    If highlightedAny Then
        msg = msg & " Inconsistent formulas were highlighted in yellow."
    End If
    
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
    Infra_Error.HandleError "ExecuteCleanData", Err, stateInfo
    Resume CleanExit
End Sub

Private Function ProcessRangeForClean(ByRef targetRange As Range, ByVal request As Infra_CleanDataRequest) As Long
    ' Note: No tracker here for performance; this is a core processing loop.
    If targetRange Is Nothing Then Exit Function

    Dim textCells As Range
    Dim area As Range
    Dim dataArr As Variant
    Dim i As Long, totalAreas As Long
    Dim changeCount As Long

    On Error Resume Next
    Set textCells = targetRange.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If textCells Is Nothing Then Exit Function

    ' Convert Text format to General so text-numbers can become real numbers
    If request.CleanTextNumbers Then
        On Error Resume Next
        For Each area In textCells.Areas
            If area.NumberFormat = "@" Then area.NumberFormat = "General"
        Next area
        On Error GoTo 0
    End If

    Const CHUNK_ROW_SIZE As Long = 20000
    Dim r As Long, chunkRowsCount As Long
    Dim chunkRange As Range

    totalAreas = textCells.Areas.Count
    For i = 1 To totalAreas
        Set area = textCells.Areas(i)
        
        If area.Cells.CountLarge = 1 Then
            Dim oldVal As Variant, newVal As Variant
            oldVal = area.Value
            newVal = CleanTextValue(oldVal, request)
            If CStr(newVal) <> CStr(oldVal) Or TypeName(newVal) <> TypeName(oldVal) Then
                area.Value = newVal
                changeCount = changeCount + 1
            End If
        Else
            ' Check if area is large enough to segment into chunks
            If area.Rows.Count > CHUNK_ROW_SIZE Then
                For r = 1 To area.Rows.Count Step CHUNK_ROW_SIZE
                    chunkRowsCount = CHUNK_ROW_SIZE
                    If r + chunkRowsCount - 1 > area.Rows.Count Then
                        chunkRowsCount = area.Rows.Count - r + 1
                    End If
                    
                    Set chunkRange = area.Rows(r).Resize(chunkRowsCount)
                    
                    dataArr = chunkRange.Value
                    changeCount = changeCount + CleanTextArray(dataArr, request)
                    chunkRange.Value = dataArr
                    
                    DoEvents
                    If Infra_Progress.UserCancelled Then Exit For
                Next r
            Else
                dataArr = area.Value
                changeCount = changeCount + CleanTextArray(dataArr, request)
                area.Value = dataArr
            End If
        End If
        
        DoEvents
        If Infra_Progress.UserCancelled Then Exit For
    Next i
    
    ProcessRangeForClean = changeCount
End Function

Private Function CleanTextArray(ByRef dataArr As Variant, ByVal request As Infra_CleanDataRequest) As Long
    ' Note: No tracker here for performance.
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim oldVal As Variant
    Dim newVal As Variant
    Dim changeCount As Long

    For rowIndex = LBound(dataArr, 1) To UBound(dataArr, 1)
        For columnIndex = LBound(dataArr, 2) To UBound(dataArr, 2)
            oldVal = dataArr(rowIndex, columnIndex)
            newVal = CleanTextValue(oldVal, request)
            If CStr(newVal) <> CStr(oldVal) Or TypeName(newVal) <> TypeName(oldVal) Then
                dataArr(rowIndex, columnIndex) = newVal
                changeCount = changeCount + 1
            End If
        Next columnIndex
    Next rowIndex
    
    CleanTextArray = changeCount
End Function

Private Function CleanTextValue(ByVal rawValue As Variant, ByVal request As Infra_CleanDataRequest) As Variant
    ' Note: No tracker here. This is called in tight loops.
    If IsError(rawValue) Then
        CleanTextValue = rawValue
        Exit Function
    End If

    Dim textValue As String
    textValue = CStr(rawValue)
    
    ' Quick check for common cleaning targets to potentially skip overhead
    If textValue = vbNullString Then
        CleanTextValue = textValue
        Exit Function
    End If
    
    ' Replace non-breaking space (Chr 160)
    If request.CleanTrimSpaces Then
        If InStr(1, textValue, ChrW$(160)) > 0 Then
            textValue = Replace(textValue, ChrW$(160), " ")
        End If
    End If
    
    ' Application.Trim handles internal double spaces; Application.Clean handles non-printables.
    If request.CleanTrimSpaces And request.CleanNonPrintables Then
        textValue = Application.Trim(Application.Clean(textValue))
    ElseIf request.CleanTrimSpaces Then
        textValue = Application.Trim(textValue)
    ElseIf request.CleanNonPrintables Then
        textValue = Application.Clean(textValue)
    End If
    
    ' Fix numbers stored as text
    If request.CleanTextNumbers Then
        If IsNumeric(textValue) Then
            On Error Resume Next
            Dim dblVal As Double
            dblVal = CDbl(textValue)
            If Err.Number = 0 Then
                CleanTextValue = dblVal
                Exit Function
            End If
            On Error GoTo 0
        End If
    End If
    
    CleanTextValue = textValue
End Function

Private Sub CleanNamedRanges(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim nm As Name
    Dim i As Long
    
    On Error Resume Next
    If Not ws Is Nothing Then
        For i = ws.Names.Count To 1 Step -1
            Set nm = ws.Names(i)
            If InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Then
                nm.Delete
            End If
        Next i
    ElseIf Not wb Is Nothing Then
        For i = wb.Names.Count To 1 Step -1
            Set nm = wb.Names(i)
            If InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Then
                nm.Delete
            End If
        Next i
    End If
    On Error GoTo 0
End Sub

Private Function HighlightInconsistentFormulas(ByVal targetRange As Range) As Boolean
    If targetRange Is Nothing Then Exit Function
    
    Dim formulaCells As Range
    Dim cell As Range
    
    On Error Resume Next
    Set formulaCells = targetRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If formulaCells Is Nothing Then Exit Function
    
    ' Safety limit to prevent freezing on large worksheets
    If formulaCells.Cells.CountLarge > Infra_Config.MAX_FORMULA_CHECK_CELLS Then
        Debug.Print "BEAVER [CLEAN]: Too many formula cells (" & formulaCells.Cells.CountLarge & "). Skipping inconsistent formula highlighting."
        Exit Function
    End If
    
    On Error Resume Next
    For Each cell In formulaCells
        If cell.Errors(xlInconsistentFormula).Value = True Then
            cell.Interior.Color = vbYellow
        End If
    Next cell
    On Error GoTo 0
    
    HighlightInconsistentFormulas = True
End Function

' --- Testing Entry Points ---
' Allows headless automated testing of the range cleaning logic.
Public Function CleanRangeDirect(ByVal targetRange As Range) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanRangeDirect")
    On Error GoTo ErrHandler
    
    Dim request As New Infra_CleanDataRequest
    CleanRangeDirect = ProcessRangeForClean(targetRange, request)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CleanRangeDirect", Err
    Resume CleanExit
End Function

Public Function CleanRangeWithOptionsDirect(ByVal targetRange As Range, ByVal request As Infra_CleanDataRequest) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanRangeWithOptionsDirect")
    On Error GoTo ErrHandler
    
    CleanRangeWithOptionsDirect = ProcessRangeForClean(targetRange, request)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CleanRangeWithOptionsDirect", Err
    Resume CleanExit
End Function
