VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_StaticSheetWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_StaticSheetWorkbook
' @Category: Feature
' @Description: Converts formulas (including spills) to static values across active sheet or whole workbook.

Private Enum StaticConversionScope
    StaticConversionScopeActiveSheet = 1
    StaticConversionScopeWorkbook = 2
End Enum

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("StaticSheetWorkbook.CanExecute")
    On Error GoTo ErrHandler
    ICommand_CanExecute = True ' Always can attempt, will prompt for scope.
    Exit Function
ErrHandler:
    Infra_Error.HandleError "StaticSheetWorkbook.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("StaticSheetWorkbook.Execute")
    On Error GoTo ErrHandler

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
    Infra_Error.HandleError "StaticSheetWorkbook.Execute", Err
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
    Infra_Error.HandleError "ExecuteStaticConversion", Err
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
    Infra_Error.HandleError "BuildSheetList", Err
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
    Infra_Error.HandleError "ConvertSheetToValues", Err
    Resume CleanExit
End Function

Private Sub ConvertCellToStaticImpl(ByVal cell As Range)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConvertCellToStaticImpl")
    On Error GoTo ErrHandler
    
    Dim isSpill As Boolean

    On Error Resume Next
    isSpill = cell.HasSpill
    If Err.Number <> 0 Then
        ' HasSpill not supported - standard conversion
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
    Infra_Error.HandleError "ConvertCellToStaticImpl", Err
    Resume CleanExit
End Sub
