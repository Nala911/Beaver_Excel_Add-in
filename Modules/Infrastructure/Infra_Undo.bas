Attribute VB_Name = "Infra_Undo"
Option Explicit

' @Module: Infra_Undo
' @Category: Infrastructure
' @Description: Custom Undo management for macro-driven changes.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

Private Const UNDO_SHEET_NAME As String = "_BeaverUndo"
Private Const MAX_UNDO_CELLS As Long = 1000000 ' 1M cells safety limit
Private Const UNDO_META_WORKBOOK_NAME As String = "BeaverUndoWorkbook"
Private Const UNDO_META_WORKSHEET_NAME As String = "BeaverUndoWorksheet"
Private Const UNDO_META_ADDRESS_NAME As String = "BeaverUndoAddress"
Private Const UNDO_META_ACTION_NAME As String = "BeaverUndoAction"
Private m_PendingUndoAction As String

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
    
    ' Copy Target to Undo Sheet at the same address so relative formulas
    ' keep their original references instead of being re-based from A1.
    Target.Copy Destination:=undoSh.Range(Target.Address)

    ' Store metadata outside the undo payload so large target ranges cannot
    ' overwrite it.
    StoreUndoMetadata Target.Worksheet.Parent.Name, Target.Worksheet.Name, Target.Address, ActionName
    
    ' Stage the Undo macro (registration happens later)
    m_PendingUndoAction = ActionName
    
    ' Clean up clipboard
    Application.CutCopyMode = False

CleanExit:
    Exit Sub
ErrHandler:
    ClearUndoMetadata
    Infra_Error.HandleError "SaveState", Err
    Resume CleanExit
End Sub

' Registers the staged undo action with Excel. Called at the end of command execution.
Public Sub RegisterPendingUndo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("RegisterPendingUndo")
    On Error GoTo ErrHandler
    
    If m_PendingUndoAction <> "" Then
        Application.OnUndo "Undo " & m_PendingUndoAction, "Infra_Undo.PerformUndo"
        m_PendingUndoAction = ""
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "RegisterPendingUndo", Err
    Resume CleanExit
End Sub

' Clears any staged undo action.
Public Sub ClearPendingUndo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("ClearPendingUndo")
    On Error GoTo ErrHandler
    
    m_PendingUndoAction = ""

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "ClearPendingUndo", Err
    Resume CleanExit
End Sub

' Restores the saved state. Triggered by Excel's Undo.
Public Sub PerformUndo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("PerformUndo")
    On Error GoTo ErrHandler
    
    Dim undoSh As Worksheet
    Set undoSh = GetUndoSheet()
    
    Dim wbName As String: wbName = GetUndoMetadataValue(UNDO_META_WORKBOOK_NAME)
    Dim wsName As String: wsName = GetUndoMetadataValue(UNDO_META_WORKSHEET_NAME)
    Dim addr As String: addr = GetUndoMetadataValue(UNDO_META_ADDRESS_NAME)
    
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
    
    ' Restore data from the same address it was captured at so formula
    ' references are restored exactly as they were before the mutation.
    Dim dataRange As Range
    Set dataRange = undoSh.Range(addr)

    dataRange.Copy Destination:=targetRange
    
    ' Clear undo sheet to prevent accidental double-restore
    undoSh.Cells.Clear
    ClearUndoMetadata
    Application.CutCopyMode = False
    
    ' Select the restored range
    On Error Resume Next
    targetRange.Select
    On Error GoTo ErrHandler

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "PerformUndo", Err
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

Private Sub StoreUndoMetadata(ByVal WorkbookName As String, ByVal WorksheetName As String, ByVal AddressText As String, ByVal ActionName As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("StoreUndoMetadata")
    On Error GoTo ErrHandler

    SetUndoMetadataValue UNDO_META_WORKBOOK_NAME, WorkbookName
    SetUndoMetadataValue UNDO_META_WORKSHEET_NAME, WorksheetName
    SetUndoMetadataValue UNDO_META_ADDRESS_NAME, AddressText
    SetUndoMetadataValue UNDO_META_ACTION_NAME, ActionName

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "StoreUndoMetadata", Err
    Resume CleanExit
End Sub

Private Sub ClearUndoMetadata()
    DeleteUndoMetadataValue UNDO_META_WORKBOOK_NAME
    DeleteUndoMetadataValue UNDO_META_WORKSHEET_NAME
    DeleteUndoMetadataValue UNDO_META_ADDRESS_NAME
    DeleteUndoMetadataValue UNDO_META_ACTION_NAME
End Sub

Private Sub SetUndoMetadataValue(ByVal NameText As String, ByVal ValueText As String)
    Dim existingName As Name

    On Error Resume Next
    ThisWorkbook.Names(NameText).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=NameText, RefersTo:="=""" & Replace(ValueText, """", """""") & """", Visible:=False
End Sub

Private Sub DeleteUndoMetadataValue(ByVal NameText As String)
    On Error Resume Next
    ThisWorkbook.Names(NameText).Delete
    On Error GoTo 0
End Sub

Private Function GetUndoMetadataValue(ByVal NameText As String) As String
    Dim nameObj As Name
    Dim evaluatedValue As Variant

    On Error Resume Next
    Set nameObj = ThisWorkbook.Names(NameText)
    On Error GoTo 0

    If nameObj Is Nothing Then Exit Function

    On Error Resume Next
    evaluatedValue = Application.Evaluate(nameObj.RefersTo)
    On Error GoTo 0

    If Not IsError(evaluatedValue) Then
        GetUndoMetadataValue = CStr(evaluatedValue)
    End If
End Function
