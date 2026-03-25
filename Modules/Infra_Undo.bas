Attribute VB_Name = "Infra_Undo"
Option Explicit

' @Module: Infra_Undo
' @Category: Infrastructure
' @Description: Custom Undo management for macro-driven changes.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

Private Const UNDO_SHEET_NAME As String = "_BeaverUndo"
Private Const MAX_UNDO_CELLS As Long = 1000000 ' 1M cells safety limit

' Captures the state of a range and registers an Undo action.
' Call this BEFORE modifying the range.
Public Sub SaveState(ByVal Target As Range, ByVal ActionName As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("SaveState")
    On Error GoTo ErrHandler
    
    If Target Is Nothing Then GoTo CleanExit
    
    ' Safety Check: Don't capture massive ranges that would crash Excel
    If Target.Cells.CountLarge > MAX_UNDO_CELLS Then
        Debug.Print "BEAVER [UNDO]: Range too large to capture safely (" & Target.Cells.CountLarge & " cells). Skipping undo registration."
        GoTo CleanExit
    End If
    
    Dim undoSh As Worksheet
    Set undoSh = GetUndoSheet()
    
    ' Clear previous undo data
    undoSh.Cells.Clear
    
    ' Copy Target to Undo Sheet (Formulas and Formats)
    Target.Copy
    undoSh.Range("A1").PasteSpecial xlPasteAll
    
    ' Store Metadata (Target Address, Workbook, Sheet)
    undoSh.Range("ZZ1").Value = Target.Worksheet.Parent.Name ' Workbook Name
    undoSh.Range("ZZ2").Value = Target.Worksheet.Name        ' Worksheet Name
    undoSh.Range("ZZ3").Value = Target.Address               ' Address
    undoSh.Range("ZZ4").Value = ActionName                   ' Action Name
    
    ' Register the Undo macro
    Application.OnUndo "Undo " & ActionName, "Infra_Undo.PerformUndo"
    
    ' Clean up clipboard
    Application.CutCopyMode = False

CleanExit:
    Exit Sub
ErrHandler:
    HandleError "SaveState", Err
    Resume CleanExit
End Sub

' Restores the saved state. Triggered by Excel's Undo.
Public Sub PerformUndo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("PerformUndo")
    On Error GoTo ErrHandler
    
    Dim undoSh As Worksheet
    Set undoSh = GetUndoSheet()
    
    Dim wbName As String: wbName = undoSh.Range("ZZ1").Value
    Dim wsName As String: wsName = undoSh.Range("ZZ2").Value
    Dim addr As String: addr = undoSh.Range("ZZ3").Value
    
    If wbName = "" Or wsName = "" Or addr = "" Then GoTo CleanExit
    
    Dim targetWb As Workbook
    On Error Resume Next
    Set targetWb = Workbooks(wbName)
    If targetWb Is Nothing Then
        ' Fallback: maybe it's the active workbook?
        Set targetWb = ActiveWorkbook
    End If
    
    Dim targetWs As Worksheet
    Set targetWs = targetWb.Worksheets(wsName)
    If targetWs Is Nothing Then GoTo CleanExit
    On Error GoTo ErrHandler
    
    Dim targetRange As Range
    Set targetRange = targetWs.Range(addr)
    
    ' Restore data
    ' Note: We ignore the metadata column (ZZ)
    Dim dataRange As Range
    Set dataRange = undoSh.Range("A1").Resize(targetRange.Rows.Count, targetRange.Columns.Count)
    
    dataRange.Copy
    targetRange.PasteSpecial xlPasteAll
    
    ' Clear undo sheet to prevent accidental double-restore
    undoSh.Cells.Clear
    Application.CutCopyMode = False
    
    ' Select the restored range
    On Error Resume Next
    targetRange.Select
    On Error GoTo ErrHandler

CleanExit:
    Exit Sub
ErrHandler:
    HandleError "PerformUndo", Err
    Resume CleanExit
End Sub

' Returns (and creates if necessary) the hidden undo sheet.
Private Function GetUndoSheet() As Worksheet
    ' Internal helper
    On Error Resume Next
    Set GetUndoSheet = ThisWorkbook.Worksheets(UNDO_SHEET_NAME)
    If GetUndoSheet Is Nothing Then
        Set GetUndoSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetUndoSheet.Name = UNDO_SHEET_NAME
        GetUndoSheet.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0
End Function
