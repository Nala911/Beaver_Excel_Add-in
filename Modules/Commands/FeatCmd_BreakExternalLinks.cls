VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_BreakExternalLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_BreakExternalLinks
' @Category: Feature
' @Description: Detects and breaks external Excel links, connections, and external named ranges.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorkbook(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("BreakExternalLinks.Execute")
    On Error GoTo ErrHandler
    
    Dim wb As Workbook
    Dim linkArr As Variant
    Dim linkInfo As String
    Dim request As Infra_ScopedRequest
    Dim ctx As Infra_ActionContext
    Dim hasWorkbookLevelItems As Boolean
    Dim hasSheetLevelItems As Boolean
    
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    
    Set wb = ctx.WorkbookRef

    ' 1. DETECTION
    linkInfo = GetExternalElementsSummary(wb, ctx.WorksheetRef, linkArr, hasWorkbookLevelItems, hasSheetLevelItems)
    If linkInfo = vbNullString Then
        Infra_Interaction.ShowInfo "No external links, connections, or named ranges were found."
        GoTo CleanExit
    End If

    ' 2. CONFIRMATION + SCOPE
    Set request = Infra_UIFactory.ShowBreakLinksDialog(ctx, linkInfo)
    If request Is Nothing Then GoTo CleanExit

    ' 4. EXECUTION
    Dim stats As String
    If request.Scope = TargetScopeActiveSheet Then
        If Not Infra_Undo.SaveStateOrConfirm(ctx.WorksheetRef.UsedRange, "Break Links") Then GoTo CleanExit
    End If
    stats = ExecuteBreakLinks(wb, request, linkArr)
    If request.Scope = TargetScopeActiveSheet And hasWorkbookLevelItems Then
        stats = stats & vbCrLf & vbCrLf & "Workbook-level links, connections, and external names were left unchanged because Sheet scope was selected."
    End If

    ' 5. COMPLETION
    Infra_Interaction.ShowInfo "Process completed for: " & IIf(request.Scope = TargetScopeWorkbook, "Whole Workbook", "Active Sheet") & vbCrLf & vbCrLf & stats

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "BreakExternalLinks.Execute", Err
    Resume CleanExit
End Sub

Private Function GetExternalElementsSummary(ByVal wb As Workbook, ByVal activeSheetRef As Worksheet, ByRef linkArr As Variant, ByRef hasWorkbookLevelItems As Boolean, ByRef hasSheetLevelItems As Boolean) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetExternalElementsSummary")
    On Error GoTo ErrHandler
    
    Dim info As String
    Dim extCount As Long, connCount As Long, nameCount As Long
    Dim nm As Name
    Dim sheetFormulaCount As Long
    Dim sheetPivotCount As Long
    Dim sheetTableCount As Long
    
    On Error Resume Next
    linkArr = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    On Error GoTo ErrHandler

    If Not IsEmpty(linkArr) Then extCount = UBound(linkArr) - LBound(linkArr) + 1
    
    If wb.Connections.Count > 0 Then
        connCount = wb.Connections.Count
    End If
    
    If wb.Names.Count > 0 Then
        For Each nm In wb.Names
            Dim refersToVal As String
            refersToVal = ""
            On Error Resume Next
            refersToVal = nm.RefersTo
            On Error GoTo ErrHandler
            If InStr(1, refersToVal, "[", vbTextCompare) > 0 Then
                nameCount = nameCount + 1
            End If
        Next nm
    End If

    GetSheetBreakableCounts activeSheetRef, sheetFormulaCount, sheetPivotCount, sheetTableCount

    hasWorkbookLevelItems = (extCount > 0 Or connCount > 0 Or nameCount > 0)
    hasSheetLevelItems = (sheetFormulaCount > 0 Or sheetPivotCount > 0 Or sheetTableCount > 0)

    If Not hasWorkbookLevelItems And Not hasSheetLevelItems Then
        GetExternalElementsSummary = vbNullString
    Else
        If hasSheetLevelItems Then
            info = "Active Sheet Items:" & vbCrLf & _
                   "Linked Formulas: " & sheetFormulaCount & vbCrLf & _
                   "Pivot Tables: " & sheetPivotCount & vbCrLf & _
                   "External Tables: " & sheetTableCount
        End If

        If hasWorkbookLevelItems Then
            If Len(info) > 0 Then info = info & vbCrLf & vbCrLf
            info = info & "Workbook-Level Items:" & vbCrLf & _
                   "Excel Links: " & extCount & vbCrLf & _
                   "Connections: " & connCount & vbCrLf & _
                   "External Named Ranges: " & nameCount
        End If
        GetExternalElementsSummary = info
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "GetExternalElementsSummary", Err
    Resume CleanExit
End Function

Private Sub GetSheetBreakableCounts(ByVal ws As Worksheet, ByRef formulaCount As Long, ByRef pivotCount As Long, ByRef tableCount As Long)
    ' No tracker here; internal optimized call
    Dim fCells As Range
    Dim area As Range
    Dim formulaArr As Variant
    Dim r As Long, c As Long

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set fCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not fCells Is Nothing Then
        For Each area In fCells.Areas
            If area.Cells.CountLarge = 1 Then
                If InStr(1, area.Formula, "[", vbTextCompare) > 0 Then formulaCount = formulaCount + 1
            Else
                formulaArr = area.Formula
                For r = 1 To UBound(formulaArr, 1)
                    For c = 1 To UBound(formulaArr, 2)
                        If InStr(1, formulaArr(r, c), "[", vbTextCompare) > 0 Then formulaCount = formulaCount + 1
                    Next c
                Next r
            End If
        Next area
    End If

    pivotCount = ws.PivotTables.Count
    tableCount = 0
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo.SourceType <> xlSrcRange Then tableCount = tableCount + 1
    Next lo
End Sub

Private Function ExecuteBreakLinks(ByVal wb As Workbook, ByVal request As Infra_ScopedRequest, ByVal linkArr As Variant) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteBreakLinks")
    On Error GoTo ErrHandler
    
    Dim brokenLinks As Long, convertedFormulas As Long, pivotsConverted As Long, tablesConverted As Long
    Dim connectionsRemoved As Long, namesRemoved As Long
    Dim ws As Worksheet, wsToProcess As New Collection
    Dim conn As WorkbookConnection
    Dim processWholeWorkbook As Boolean
    Dim index As Long
    Dim originalScreenUpdating As Boolean
 
    If request Is Nothing Then GoTo CleanExit
    processWholeWorkbook = (request.Scope = TargetScopeWorkbook)
    
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
        
        If wb.Connections.Count > 0 Then
            Dim connIndex As Long
            For connIndex = wb.Connections.Count To 1 Step -1
                Set conn = wb.Connections(connIndex)
                On Error Resume Next
                conn.Delete
                If Err.Number = 0 Then connectionsRemoved = connectionsRemoved + 1
                On Error GoTo ErrHandler
            Next connIndex
        End If

        namesRemoved = RemoveExternalWorkbookNames(wb)
    End If

    ' Range/Sheet level
    Set wsToProcess = Infra_CommandSupport.ResolveWorksheetsToProcess(request.Context, request.Scope)
    If wsToProcess Is Nothing Then GoTo CleanExit
 
    originalScreenUpdating = Application.ScreenUpdating
    If request.Scope = TargetScopeWorkbook Then
        Application.ScreenUpdating = False
    End If
 
    Infra_Progress.StartProgress "Breaking External Links", wsToProcess.Count
    For Each ws In wsToProcess
        index = index + 1
        ProcessSheetInternal ws, convertedFormulas, pivotsConverted, tablesConverted
        Infra_Progress.UpdateProgress index, True
        If Infra_Progress.UserCancelled Then Exit For
    Next ws

    ExecuteBreakLinks = "Links Broken: " & brokenLinks & vbCrLf & _
                       "Connections Removed: " & connectionsRemoved & vbCrLf & _
                       "External Names Removed: " & namesRemoved & vbCrLf & _
                       "Formulas Converted: " & convertedFormulas & vbCrLf & _
                       "Pivots Converted: " & pivotsConverted & vbCrLf & _
                       "Tables Converted: " & tablesConverted

CleanExit:
    If Not request Is Nothing Then
        If request.Scope = TargetScopeWorkbook Then
            Application.ScreenUpdating = originalScreenUpdating
        End If
    End If
    Infra_Progress.EndProgress
    Exit Function

ErrHandler:
    Infra_Progress.EndProgress
    Infra_Error.HandleError "ExecuteBreakLinks", Err
    Resume CleanExit
End Function

Private Function RemoveExternalWorkbookNames(ByVal wb As Workbook) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("RemoveExternalWorkbookNames")
    On Error GoTo ErrHandler
 
    Dim i As Long
 
    If wb Is Nothing Then GoTo CleanExit
 
    For i = wb.Names.Count To 1 Step -1
        Dim extRefersToVal As String
        extRefersToVal = ""
        On Error Resume Next
        extRefersToVal = wb.Names(i).RefersTo
        On Error GoTo ErrHandler
        
        If InStr(1, extRefersToVal, "[", vbTextCompare) > 0 Then
            On Error Resume Next
            wb.Names(i).Delete
            If Err.Number = 0 Then RemoveExternalWorkbookNames = RemoveExternalWorkbookNames + 1
            Err.Clear
            On Error GoTo ErrHandler
        End If
    Next i
 
CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "RemoveExternalWorkbookNames", Err
    Resume CleanExit
End Function

Private Sub ProcessSheetInternal(ByVal ws As Worksheet, ByRef fCount As Long, ByRef pCount As Long, ByRef tCount As Long)
    ' No tracker for performance
    Dim area As Range, fCells As Range, pvt As PivotTable, lo As ListObject
    Dim formulaArr As Variant, valueArr As Variant
    Dim r As Long, c As Long
    Dim changed As Boolean
    Dim pvtIndex As Long
    Dim loIndex As Long
    
    On Error Resume Next
    Set fCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If Not fCells Is Nothing Then
        For Each area In fCells.Areas
            If area.Cells.CountLarge = 1 Then
                If InStr(1, area.Formula, "[", vbTextCompare) > 0 Then
                    area.Value = area.Value
                    fCount = fCount + 1
                End If
            Else
                formulaArr = area.Formula
                valueArr = area.Value
                changed = False
                For r = 1 To UBound(formulaArr, 1)
                    For c = 1 To UBound(formulaArr, 2)
                        If InStr(1, formulaArr(r, c), "[", vbTextCompare) > 0 Then
                            ' Note: valueArr(r, c) already contains the result of the formula
                            formulaArr(r, c) = valueArr(r, c)
                            fCount = fCount + 1
                            changed = True
                        End If
                    Next c
                Next r
                If changed Then area.Value = formulaArr
            End If
        Next area
    End If

    If ws.PivotTables.Count > 0 Then
        For pvtIndex = ws.PivotTables.Count To 1 Step -1
            Set pvt = ws.PivotTables(pvtIndex)
            pvt.TableRange2.Copy
            pvt.TableRange2.PasteSpecial xlPasteValues
            pCount = pCount + 1
        Next pvtIndex
    End If

    If ws.ListObjects.Count > 0 Then
        For loIndex = ws.ListObjects.Count To 1 Step -1
            Set lo = ws.ListObjects(loIndex)
            If lo.SourceType <> xlSrcRange Then
                lo.Range.Copy
                lo.Range.PasteSpecial xlPasteValues
                tCount = tCount + 1
            End If
        Next loIndex
    End If
    
    Application.CutCopyMode = False
End Sub

' --- Testing Entry Points ---
' Allows headless automated testing of the link breaking and name removal logic.
Public Function ExecuteBreakLinksDirect(ByVal wb As Workbook, ByVal request As Infra_ScopedRequest, ByRef linkArr As Variant) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteBreakLinksDirect")
    On Error GoTo ErrHandler
    
    ExecuteBreakLinksDirect = ExecuteBreakLinks(wb, request, linkArr)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ExecuteBreakLinksDirect", Err
    Resume CleanExit
End Function

Public Function RemoveExternalWorkbookNamesDirect(ByVal wb As Workbook) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("RemoveExternalWorkbookNamesDirect")
    On Error GoTo ErrHandler
    
    RemoveExternalWorkbookNamesDirect = RemoveExternalWorkbookNames(wb)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "RemoveExternalWorkbookNamesDirect", Err
    Resume CleanExit
End Function
