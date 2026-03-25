Attribute VB_Name = "Feat_MakeItStatic"
Option Explicit

' @Module: Feat_MakeItStatic
' @Category: Feature
' @Description: Converts formulas (including spills) to static values across selection, sheet, or workbook.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppStateGuard, Infra_AppState, Infra_Config, Infra_Error, Infra_Progress, Infra_UIFactory, Infra_StaticRequest, Infra_Undo

Private Enum StaticConversionScope
    StaticConversionScopeActiveSheet = 1
    StaticConversionScopeWorkbook = 2
End Enum

' Prompts for scope (workbook vs sheet), then converts all formulas to values.
Public Sub StaticSheetWorkbook()
    Dim tracker As Object: Set tracker = Infra_Error.Track("StaticSheetWorkbook")
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim ctx As Infra_ActionContext
    Dim request As Infra_StaticRequest

    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorkbookRef Is Nothing Or ctx.WorksheetRef Is Nothing Then GoTo CleanExit

    Set request = Infra_UIFactory.ShowStaticConversionDialog(ctx)
    If request Is Nothing Then GoTo CleanExit

    ExecuteStaticConversion request

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "StaticSheetWorkbook", Err
    Resume CleanExit
End Sub

' Converts the selected range's formulas to static values while preserving
' formatting. For small selections this is faster than sheet-level scanning.
Public Sub MakePermanent()
    Dim tracker As Object: Set tracker = Infra_Error.Track("MakePermanent")
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim rng As Range

    If Not Infra_AppState.IsRangeSelected() Then GoTo CleanExit
    Set rng = Selection
    If rng Is Nothing Then GoTo CleanExit

    Infra_Undo.SaveState rng, "Make Permanent"
    Infra_Progress.StartProgress "Converting Selection to Values", 1

    ' Direct value assignment preserves existing formatting and avoids clipboard overhead.
    rng.Value = rng.Value
    Application.CutCopyMode = False
    Infra_Progress.UpdateProgress 1, True

CleanExit:
    Infra_Progress.EndProgress
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    HandleError "MakePermanent", Err
    Resume CleanExit
End Sub

' Internal helper: converts a single cell to its static value.
' Handles spill ranges (Excel 365 / 2021) correctly.
Private Sub ConvertCellToStaticImpl(ByVal cell As Range)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConvertCellToStaticImpl")
    On Error GoTo ErrHandler
    
    Dim isSpill As Boolean

    On Error Resume Next
    isSpill = cell.HasSpill
    If Err.Number <> 0 Then
        ' HasSpill not supported â€” standard conversion
        Err.Clear
        On Error GoTo 0
        cell.Value = cell.Value
    Else
        On Error GoTo 0
        If isSpill Then
            cell.SpillingToRange.Value = cell.SpillingToRange.Value
        Else
            cell.Value = cell.Value
        End If
    End If

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "ConvertCellToStaticImpl", Err
    Resume CleanExit
End Sub

Private Sub ExecuteStaticConversion(ByVal request As Infra_StaticRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteStaticConversion")
    On Error GoTo ErrHandler

    Dim sheetsToProcess As Collection
    Dim wsItem As Variant
    Dim ws As Worksheet
    Dim convertedCount As Long
    Dim index As Long
    Dim scopeLabel As String

    If request Is Nothing Then GoTo CleanExit
    If request.Context Is Nothing Then GoTo CleanExit

    Set sheetsToProcess = BuildSheetList(request)
    If sheetsToProcess Is Nothing Then GoTo CleanExit

    If request.Scope = StaticConversionScopeActiveSheet Then
        Infra_Undo.SaveState request.Context.WorksheetRef.UsedRange, "Static Sheet"
    End If

    Infra_Progress.StartProgress "Converting Formulas to Values", sheetsToProcess.Count
    For Each wsItem In sheetsToProcess
        index = index + 1
        Set ws = wsItem
        convertedCount = convertedCount + ConvertSheetToValues(ws)
        Infra_Progress.UpdateProgress index, True
        If Infra_Progress.UserCancelled Then Exit For
    Next wsItem

    scopeLabel = IIf(request.Scope = StaticConversionScopeWorkbook, "Whole Workbook", "Active Sheet")
    MsgBox "Done! Converted " & convertedCount & " formula cell(s) - " & scopeLabel & ".", _
           vbInformation, Infra_Config.ADDIN_NAME & " - Process Complete"

CleanExit:
    Infra_Progress.EndProgress
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    HandleError "ExecuteStaticConversion", Err
    Resume CleanExit
End Sub

Private Function BuildSheetList(ByVal request As Infra_StaticRequest) As Collection
    Dim tracker As Object: Set tracker = Infra_Error.Track("BuildSheetList")
    On Error GoTo ErrHandler

    Dim sheets As New Collection
    Dim ws As Worksheet

    If request.Scope = StaticConversionScopeWorkbook Then
        For Each ws In request.Context.WorkbookRef.Worksheets
            sheets.Add ws
        Next ws
    Else
        sheets.Add request.Context.WorksheetRef
    End If

    Set BuildSheetList = sheets

CleanExit:
    Exit Function
ErrHandler:
    Set BuildSheetList = Nothing
    HandleError "BuildSheetList", Err
    Resume CleanExit
End Function

Private Function ConvertSheetToValues(ByVal ws As Worksheet) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConvertSheetToValues")
    On Error GoTo ErrHandler

    Dim formulaCells As Range
    Dim area As Range
    Dim cell As Range
    Dim hasSpillsInArea As Boolean

    If ws Is Nothing Then GoTo CleanExit
    If ws.UsedRange.Cells.Count <= 1 And IsEmpty(ws.Range("A1")) Then GoTo CleanExit

    If ws.FilterMode Then ws.ShowAllData

    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler

    If formulaCells Is Nothing Then GoTo CleanExit

    For Each area In formulaCells.Areas
        hasSpillsInArea = False

        On Error Resume Next
        If IsNull(area.HasSpill) Or area.HasSpill Then
            hasSpillsInArea = True
        End If
        On Error GoTo ErrHandler

        If Not hasSpillsInArea Then
            area.Value = area.Value
            ConvertSheetToValues = ConvertSheetToValues + area.Cells.Count
        Else
            For Each cell In area.Cells
                If cell.HasFormula Then
                    ConvertCellToStaticImpl cell
                    ConvertSheetToValues = ConvertSheetToValues + 1
                End If
            Next cell
        End If
    Next area

CleanExit:
    Exit Function
ErrHandler:
    HandleError "ConvertSheetToValues", Err
    Resume CleanExit
End Function
