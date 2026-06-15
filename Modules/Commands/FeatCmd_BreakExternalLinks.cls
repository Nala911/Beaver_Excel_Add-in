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
Implements IBatchProcessor

' @Module: FeatCmd_BreakExternalLinks
' @Category: Feature
' @Description: Detects and breaks external Excel links, connections, and external named ranges.

Private m_CurrentRequest As Infra_ScopedRequest
Private m_LinkArr As Variant
Private m_BrokenLinks As Long
Private m_ConnectionsRemoved As Long
Private m_NamesRemoved As Long
Private m_ConvertedFormulas As Long
Private m_PivotsConverted As Long
Private m_TablesConverted As Long

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

    Infra_CommandSupport.GetSheetBreakableCounts activeSheetRef, sheetFormulaCount, sheetPivotCount, sheetTableCount

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


Private Function ExecuteBreakLinks(ByVal wb As Workbook, ByVal request As Infra_ScopedRequest, ByVal linkArr As Variant) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteBreakLinks")
    On Error GoTo ErrHandler

    ' Reset accumulators
    m_BrokenLinks = 0
    m_ConnectionsRemoved = 0
    m_NamesRemoved = 0
    m_ConvertedFormulas = 0
    m_PivotsConverted = 0
    m_TablesConverted = 0

    Set m_CurrentRequest = request
    m_LinkArr = linkArr

    Dim processWholeWorkbook As Boolean
    processWholeWorkbook = (request.Scope = TargetScopeWorkbook)
    Dim linkIndex As Long

    ' Global workbook-level breaks
    If processWholeWorkbook Then
        If Not IsEmpty(linkArr) Then
            Dim i As Long
            For i = LBound(linkArr) To UBound(linkArr)
                On Error Resume Next
                wb.BreakLink Name:=linkArr(i), Type:=xlLinkTypeExcelLinks
                If Err.Number = 0 Then m_BrokenLinks = m_BrokenLinks + 1
                On Error GoTo ErrHandler
            Next i
        End If

        If wb.Connections.Count > 0 Then
            Dim connIndex As Long
            Dim conn As WorkbookConnection
            For connIndex = wb.Connections.Count To 1 Step -1
                Set conn = wb.Connections(connIndex)
                On Error Resume Next
                conn.Delete
                If Err.Number = 0 Then m_ConnectionsRemoved = m_ConnectionsRemoved + 1
                On Error GoTo ErrHandler
            Next connIndex
        End If

        Dim namesRemoved As Long
        Infra_CommandSupport.CleanWorkbookNames wb, Nothing, NameCleanCriteriaExternal, namesRemoved
        m_NamesRemoved = namesRemoved
    End If

    ' Execute worksheet batch conversions
    Dim totalProcessed As Long
    totalProcessed = Infra_BatchProcessor.Execute(request.Context, request.Scope, Me, "Break External Links", False)

    ' Post-process active sheet links: if a link source has no more references anywhere in the workbook, break it.
    If Not processWholeWorkbook Then
        If Not IsEmpty(linkArr) Then
            For linkIndex = LBound(linkArr) To UBound(linkArr)
                If CountLinkSourceReferences(wb, CStr(linkArr(linkIndex))) = 0 Then
                    On Error Resume Next
                    wb.BreakLink Name:=linkArr(linkIndex), Type:=xlLinkTypeExcelLinks
                    If Err.Number = 0 Then m_BrokenLinks = m_BrokenLinks + 1
                    On Error GoTo ErrHandler
                End If
            Next linkIndex
        End If
    End If

    ExecuteBreakLinks = "Links Broken: " & m_BrokenLinks & vbCrLf & _
                       "Connections Removed: " & m_ConnectionsRemoved & vbCrLf & _
                       "External Names Removed: " & m_NamesRemoved & vbCrLf & _
                       "Formulas Converted: " & m_ConvertedFormulas & vbCrLf & _
                       "Pivots Converted: " & m_PivotsConverted & vbCrLf & _
                       "Tables Converted: " & m_TablesConverted

CleanExit:
    Set m_CurrentRequest = Nothing
    m_LinkArr = Empty
    Exit Function

ErrHandler:
    Set m_CurrentRequest = Nothing
    m_LinkArr = Empty
    Infra_Error.HandleError "ExecuteBreakLinks", Err
    Resume CleanExit
End Function

Private Function IBatchProcessor_ProcessRange(ByVal targetRange As Range) As Long
    IBatchProcessor_ProcessRange = 0
End Function

Private Function IBatchProcessor_ProcessWorksheet(ByVal ws As Worksheet) As Long
    Dim convertedFormulas As Long
    Dim pivotsConverted As Long
    Dim tablesConverted As Long

    ProcessSheetInternal ws, convertedFormulas, pivotsConverted, tablesConverted

    m_ConvertedFormulas = m_ConvertedFormulas + convertedFormulas
    m_PivotsConverted = m_PivotsConverted + pivotsConverted
    m_TablesConverted = m_TablesConverted + tablesConverted

    IBatchProcessor_ProcessWorksheet = convertedFormulas + pivotsConverted + tablesConverted
End Function

Private Sub ProcessSheetInternal(ByVal ws As Worksheet, ByRef fCount As Long, ByRef pCount As Long, ByRef tCount As Long)
    ' No tracker for performance
    Dim area As Range, fCells As Range, pvt As PivotTable, lo As ListObject
    Dim formulaArr As Variant, valueArr As Variant
    Dim r As Long, c As Long
    Dim changed As Boolean
    Dim pvtIndex As Long
    Dim loIndex As Long
    
    ' Variable declarations at the top to comply with VBA compiler rules
    Dim cell As Range
    Dim isFormulaArrArray As Boolean
    Dim hasExternalLinks As Boolean
    Dim hasSpillOrArray As Boolean
    
    On Error Resume Next
    Set fCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If Not fCells Is Nothing Then
        For Each area In fCells.Areas
            If area.Cells.CountLarge = 1 Then
                If InStr(1, area.Formula2, "[", vbTextCompare) > 0 Then
                    fCount = fCount + Infra_ValueConversion.ConvertCellToStatic(area)
                End If
            Else
                formulaArr = area.Formula2
                valueArr = area.Value
                changed = False
                
                isFormulaArrArray = IsArray(formulaArr)
                hasExternalLinks = False
                
                If isFormulaArrArray Then
                    For r = 1 To UBound(formulaArr, 1)
                        For c = 1 To UBound(formulaArr, 2)
                            If InStr(1, formulaArr(r, c), "[", vbTextCompare) > 0 Then
                                hasExternalLinks = True
                                Exit For
                            End If
                        Next c
                        If hasExternalLinks Then Exit For
                    Next r
                Else
                    If InStr(1, formulaArr, "[", vbTextCompare) > 0 Then
                        hasExternalLinks = True
                    End If
                End If
                
                If hasExternalLinks Then
                    hasSpillOrArray = False
                    
                    On Error Resume Next
                    If IsNull(area.HasSpill) Or area.HasSpill = True Or IsNull(area.HasArray) Or area.HasArray = True Then
                        hasSpillOrArray = True
                    End If
                    On Error GoTo 0
                    
                    If Not hasSpillOrArray Then
                        If isFormulaArrArray Then
                            For r = 1 To UBound(formulaArr, 1)
                                For c = 1 To UBound(formulaArr, 2)
                                    If InStr(1, formulaArr(r, c), "[", vbTextCompare) > 0 Then
                                        formulaArr(r, c) = valueArr(r, c)
                                        fCount = fCount + 1
                                        changed = True
                                    End If
                                Next c
                            Next r
                            If changed Then area.Value = formulaArr
                        Else
                            area.Value = valueArr
                            fCount = fCount + 1
                        End If
                    Else
                        For Each cell In area.Cells
                            If cell.HasFormula Then
                                If InStr(1, cell.Formula2, "[", vbTextCompare) > 0 Then
                                    fCount = fCount + Infra_ValueConversion.ConvertCellToStatic(cell)
                                End If
                            End If
                        Next cell
                    End If
                End If
            End If
        Next area
    End If

    If ws.PivotTables.Count > 0 Then
        For pvtIndex = ws.PivotTables.Count To 1 Step -1
            Set pvt = ws.PivotTables(pvtIndex)
            On Error Resume Next
            pvt.TableRange2.Copy
            pvt.TableRange2.PasteSpecial xlPasteValues
            If Err.Number = 0 Then
                pCount = pCount + 1
            End If
            Err.Clear
            On Error GoTo 0
        Next pvtIndex
    End If

    If ws.ListObjects.Count > 0 Then
        For loIndex = ws.ListObjects.Count To 1 Step -1
            Set lo = ws.ListObjects(loIndex)
            If lo.SourceType <> xlSrcRange Then
                On Error Resume Next
                lo.Range.Copy
                lo.Range.PasteSpecial xlPasteValues
                If Err.Number = 0 Then
                    tCount = tCount + 1
                End If
                Err.Clear
                On Error GoTo 0
            End If
        Next loIndex
    End If
    
    Application.CutCopyMode = False
End Sub

Private Function CountLinkSourceReferences(ByVal wb As Workbook, ByVal linkName As String) As Long
    Dim ws As Worksheet
    Dim formulaCells As Range
    Dim foundCell As Range
    Dim firstAddress As String
    Dim refCount As Long
    
    ' Extract the filename with brackets if it's a full path
    Dim fileName As String
    Dim openBracketPos As Long
    openBracketPos = InStrRev(linkName, "\")
    If openBracketPos > 0 Then
        fileName = Mid(linkName, openBracketPos + 1)
    Else
        fileName = linkName
    End If
    ' Prefix with bracket to look for [filename]
    fileName = "[" & fileName & "]"

    For Each ws In wb.Worksheets
        ' Skip the undo sheet and test sheets
        If ws.Name <> "_BeaverUndo" And ws.Name <> "_BeaverTests" And Left(ws.Name, 10) <> "Test_Temp_" Then
            Set formulaCells = Nothing
            On Error Resume Next
            Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            
            If Not formulaCells Is Nothing Then
                Set foundCell = Nothing
                On Error Resume Next
                Set foundCell = formulaCells.Find(What:=fileName, LookIn:=xlFormulas, LookAt:=xlPart)
                On Error GoTo 0
                If Not foundCell Is Nothing Then
                    firstAddress = foundCell.Address
                    Do
                        refCount = refCount + 1
                        On Error Resume Next
                        Set foundCell = formulaCells.FindNext(foundCell)
                        On Error GoTo 0
                        If foundCell Is Nothing Then Exit Do
                    Loop While foundCell.Address <> firstAddress
                End If
            End If
        End If
    Next ws
    
    CountLinkSourceReferences = refCount
End Function

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
    
    Dim removedCount As Long
    Infra_CommandSupport.CleanWorkbookNames wb, Nothing, NameCleanCriteriaExternal, removedCount
    RemoveExternalWorkbookNamesDirect = removedCount

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "RemoveExternalWorkbookNamesDirect", Err
    Resume CleanExit
End Function
