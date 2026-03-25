Attribute VB_Name = "Feat_BreakExternalLinks"
Option Explicit

' @Module: Feat_BreakExternalLinks
' @Category: Feature
' @Description: Detects and breaks external Excel links, connections, and external named ranges.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_AppStateGuard, Infra_Config, Infra_Error, Infra_UIFactory

' Main entry point. Detects external elements, prompts user, then breaks them.
Public Sub BreakExternalLinks()
    PushContext "BreakExternalLinks"
    On Error GoTo ErrHandler
    
    Dim guard As New Infra_AppStateGuard
    Dim wb As Workbook
    Dim linkArr As Variant
    Dim linkInfo As String
    Dim processWholeWorkbook As Boolean
    Dim scopeChoice As Long
    Dim ctx As Infra_ActionContext
    
    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit
    
    Set wb = ctx.WorkbookRef

    ' 1. DETECTION
    linkInfo = GetExternalElementsSummary(wb, linkArr)
    If linkInfo = vbNullString Then
        MsgBox "No external links, connections, or named ranges were found.", vbInformation, Infra_Config.ADDIN_NAME
        GoTo CleanExit
    End If

    ' 2. CONFIRMATION + SCOPE
    scopeChoice = Infra_UIFactory.ShowBreakLinksScopeDialog(ctx, linkInfo)
    If scopeChoice = 0 Then GoTo CleanExit
    processWholeWorkbook = (scopeChoice = 2)

    ' 4. EXECUTION
    Dim stats As String
    stats = ExecuteBreakLinks(wb, processWholeWorkbook, linkArr, ctx.WorksheetRef)

    ' 5. COMPLETION
    MsgBox "Process Completed for: " & IIf(processWholeWorkbook, "Whole Workbook", "Active Sheet") & vbCrLf & vbCrLf & stats, vbInformation, Infra_Config.ADDIN_NAME

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "BreakExternalLinks", Err
End Sub

' Counts external links, connections, and external named ranges and returns
' a summary string. Also populates linkArr (ByRef) for later use.
Private Function GetExternalElementsSummary(ByVal wb As Workbook, ByRef linkArr As Variant) As String
    PushContext "GetExternalElementsSummary"
    On Error GoTo ErrHandler
    
    Dim info As String
    Dim extCount As Long, connCount As Long, nameCount As Long
    Dim nm As Name
    
    On Error Resume Next
    linkArr = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    On Error GoTo ErrHandler

    If Not IsEmpty(linkArr) Then extCount = UBound(linkArr) - LBound(linkArr) + 1
    connCount = wb.Connections.Count
    
    For Each nm In wb.Names
        If InStr(1, nm.RefersTo, "[", vbTextCompare) > 0 Then nameCount = nameCount + 1
    Next nm

    If extCount = 0 And connCount = 0 And nameCount = 0 Then
        GetExternalElementsSummary = vbNullString
    Else
        info = "Excel Links: " & extCount & vbCrLf & _
               "Connections: " & connCount & vbCrLf & _
               "External Named Ranges: " & nameCount
        GetExternalElementsSummary = info
    End If

CleanExit:
    PopContext
    Exit Function

ErrHandler:
    HandleError "GetExternalElementsSummary", Err
End Function

' Performs the actual link breaking. If processWholeWorkbook is True, breaks
' workbook-level links and connections first, then processes every sheet;
' otherwise processes only the active sheet.
Private Function ExecuteBreakLinks(ByVal wb As Workbook, ByVal processWholeWorkbook As Boolean, ByVal linkArr As Variant, ByVal activeSheetRef As Worksheet) As String
    PushContext "ExecuteBreakLinks"
    On Error GoTo ErrHandler
    
    Dim brokenLinks As Long, convertedFormulas As Long, pivotsConverted As Long, tablesConverted As Long
    Dim ws As Worksheet, wsToProcess As New Collection
    Dim conn As WorkbookConnection
    
    ' Global breaks
    If processWholeWorkbook Then
        If Not IsEmpty(linkArr) Then
            Dim i As Long
            For i = LBound(linkArr) To UBound(linkArr)
                On Error Resume Next
                wb.BreakLink Name:=linkArr(i), Type:=xlLinkTypeExcelLinks
                If Err.Number = 0 Then brokenLinks = brokenLinks + 1
                On Error GoTo ErrHandler
            Next i
        End If
        
        For Each conn In wb.Connections
            On Error Resume Next
            conn.Delete
            On Error GoTo ErrHandler
        Next conn
    End If

    ' Range/Sheet level
    If processWholeWorkbook Then
        For Each ws In wb.Worksheets: wsToProcess.Add ws: Next ws
    Else
        wsToProcess.Add activeSheetRef
    End If

    For Each ws In wsToProcess
        ProcessSheetInternal ws, convertedFormulas, pivotsConverted, tablesConverted
    Next ws

    ExecuteBreakLinks = "Links Broken: " & brokenLinks & vbCrLf & _
                       "Formulas Converted: " & convertedFormulas & vbCrLf & _
                       "Pivots Converted: " & pivotsConverted & vbCrLf & _
                       "Tables Converted: " & tablesConverted

CleanExit:
    PopContext
    Exit Function

ErrHandler:
    HandleError "ExecuteBreakLinks", Err
End Function

' Processes a single worksheet: converts external formulas to values,
' copies pivot table data to static values, and flattens external tables.
Private Sub ProcessSheetInternal(ByVal ws As Worksheet, ByRef fCount As Long, ByRef pCount As Long, ByRef tCount As Long)
    PushContext "ProcessSheetInternal"
    On Error GoTo ErrHandler
    
    Dim Cell As Range, fCells As Range, pvt As PivotTable, lo As ListObject
    
    On Error Resume Next
    Set fCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler
    
    If Not fCells Is Nothing Then
        For Each Cell In fCells
            If InStr(1, Cell.Formula, "[", vbTextCompare) > 0 Then
                Cell.Value = Cell.Value
                fCount = fCount + 1
            End If
        Next Cell
    End If

    For Each pvt In ws.PivotTables
        pvt.TableRange2.Copy
        pvt.TableRange2.PasteSpecial xlPasteValues
        pCount = pCount + 1
    Next pvt

    For Each lo In ws.ListObjects
        If lo.SourceType <> xlSrcRange Then
            lo.Range.Copy
            lo.Range.PasteSpecial xlPasteValues
            tCount = tCount + 1
        End If
    Next lo
    
    Application.CutCopyMode = False

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "ProcessSheetInternal", Err
End Sub
