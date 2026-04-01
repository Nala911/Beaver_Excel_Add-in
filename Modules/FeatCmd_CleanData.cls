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

Private Enum CleanDataScope
    CleanDataScopeSelection = 1
    CleanDataScopeActiveSheet = 2
    CleanDataScopeWorkbook = 3
End Enum

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanData.CanExecute")
    On Error GoTo ErrHandler
    ICommand_CanExecute = Infra_AppState.IsRangeSelected()
    Exit Function
ErrHandler:
    Infra_Error.HandleError "CleanData.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanData.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As Infra_ActionContext
    Dim request As Infra_CleanDataRequest
    Dim stateInfo As Collection

    Set stateInfo = Infra_Error.NewStateInfo()

    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit
    If Not ctx.HasRangeSelection Then GoTo CleanExit
    If ctx.SelectionRange Is Nothing Or ctx.WorksheetRef Is Nothing Or ctx.WorkbookRef Is Nothing Then GoTo CleanExit
    
    Infra_Error.AddState stateInfo, "Workbook", ctx.WorkbookRef.Name
    Infra_Error.AddState stateInfo, "Worksheet", ctx.WorksheetRef.Name
    Infra_Error.AddState stateInfo, "Selection", ctx.SelectionRange.Address(False, False)

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

    Set stateInfo = Infra_Error.NewStateInfo()
    
    If request Is Nothing Then GoTo CleanExit
    If request.Context Is Nothing Then GoTo CleanExit
    Infra_Error.AddState stateInfo, "Scope", request.Scope
    
    ' --- 1. Save State for Undo ---
    Select Case request.Scope
        Case CleanDataScopeSelection
            Set targetRange = request.Context.SelectionRange
        Case CleanDataScopeActiveSheet
            Set targetRange = request.Context.WorksheetRef.UsedRange
        Case CleanDataScopeWorkbook
            ' Capture the active sheet as the primary target for undo
            Set targetRange = request.Context.WorksheetRef.UsedRange
    End Select
    
    Infra_Undo.SaveState targetRange, "Clean Data"
    
    ' --- 2. Initialize Progress ---
    Infra_Progress.StartProgress "Cleaning Data", 100
    
    Select Case request.Scope
        Case CleanDataScopeSelection
            ProcessRangeForClean request.Context.SelectionRange, 100
        Case CleanDataScopeActiveSheet
            ProcessRangeForClean request.Context.WorksheetRef.UsedRange, 100
        Case CleanDataScopeWorkbook
            Dim i As Integer, count As Integer
            count = request.Context.WorkbookRef.Worksheets.Count
            For i = 1 To count
                Set ws = request.Context.WorkbookRef.Worksheets(i)
                Infra_Progress.UpdateProgress (i / count) * 100
                ProcessRangeForClean ws.UsedRange, 0 ' Don't sub-track within sheet to avoid overhead
                If Infra_Progress.UserCancelled Then Exit For
            Next i
    End Select
    
    Infra_Progress.EndProgress

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ExecuteCleanData", Err, stateInfo
    Resume CleanExit
End Sub

Private Sub ProcessRangeForClean(ByRef targetRange As Range, Optional ByVal ProgressWeight As Double = 0)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ProcessRangeForClean")
    On Error GoTo ErrHandler
    
    If targetRange Is Nothing Then GoTo CleanExit

    Dim textCells As Range
    Dim area As Range
    Dim dataArr As Variant
    Dim i As Long, totalAreas As Long

    On Error Resume Next
    Set textCells = targetRange.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo ErrHandler

    If textCells Is Nothing Then GoTo CleanExit

    totalAreas = textCells.Areas.Count
    For i = 1 To totalAreas
        Set area = textCells.Areas(i)
        
        If ProgressWeight > 0 Then
            Infra_Progress.UpdateProgress (i / totalAreas) * ProgressWeight
        End If
        
        If area.Cells.CountLarge = 1 Then
            area.Value = CleanTextValue(area.Value)
        Else
            dataArr = area.Value
            CleanTextArray dataArr
            area.Value = dataArr
        End If
        
        If i Mod 10 = 0 Then DoEvents ' Ensure responsiveness for many small areas
    Next i

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ProcessRangeForClean", Err
    Resume CleanExit
End Sub

Private Sub CleanTextArray(ByRef dataArr As Variant)
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanTextArray")
    On Error GoTo ErrHandler
    
    Dim rowIndex As Long
    Dim columnIndex As Long

    For rowIndex = LBound(dataArr, 1) To UBound(dataArr, 1)
        For columnIndex = LBound(dataArr, 2) To UBound(dataArr, 2)
            dataArr(rowIndex, columnIndex) = CleanTextValue(dataArr(rowIndex, columnIndex))
        Next columnIndex
    Next rowIndex

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "CleanTextArray", Err
    Resume CleanExit
End Sub

Private Function CleanTextValue(ByVal rawValue As Variant) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanTextValue")
    On Error GoTo ErrHandler
    
    Dim textValue As String
    
    textValue = CStr(rawValue)
    textValue = Replace(textValue, Chr(160), " ")
    CleanTextValue = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(textValue))

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CleanTextValue", Err
    Resume CleanExit
End Function
