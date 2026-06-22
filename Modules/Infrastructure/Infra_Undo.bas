Attribute VB_Name = "Infra_Undo"
Option Explicit

' @Module: Infra_Undo
' @Category: Infrastructure
' @Description: Custom Undo management for macro-driven changes.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

Private Const UNDO_SHEET_NAME As String = "_BeaverUndo"
Private Const UNDO_META_WORKBOOK_NAME As String = "BeaverUndoWorkbook"
Private Const UNDO_META_WORKSHEET_NAME As String = "BeaverUndoWorksheet"
Private Const UNDO_META_ADDRESS_NAME As String = "BeaverUndoAddress"
Private Const UNDO_META_ACTION_NAME As String = "BeaverUndoAction"
Private m_PendingUndoAction As String

' Captures the state of a range and registers an Undo action.
' Call this BEFORE modifying the range.
Public Function SaveState(ByVal Target As Range, ByVal ActionName As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("SaveState")
    Dim links As Variant
    Dim undoRange As Range
    Dim formulaCells As Range
    Dim foundCell As Range
    Dim firstAddress As String
    Dim extCells As Collection
    Dim c As Variant
    Dim i As Long
    On Error GoTo ErrHandler
    
    If Target Is Nothing Then GoTo CleanExit
    
    Dim captureRange As Range
    Set captureRange = Target
    
    ' Safety Check: Don't capture massive ranges that would crash Excel.
    ' If the target range is large, try to restrict it to the sheet's UsedRange.
    If captureRange.Cells.CountLarge > Infra_Config.MAX_UNDO_CELLS Then
        Dim usedIntersect As Range
        On Error Resume Next
        Set usedIntersect = Application.Intersect(captureRange, captureRange.Worksheet.UsedRange)
        On Error GoTo ErrHandler
        
        If Not usedIntersect Is Nothing Then
            Set captureRange = usedIntersect
        Else
            Debug.Print "BEAVER [UNDO]: Target range is too large to capture safely (" & Target.Cells.CountLarge & " cells) and does not intersect with UsedRange. Skipping undo registration."
            GoTo CleanExit
        End If
    End If
    
    ' Double check size after intersection (if the intersection itself is still too large)
    If captureRange.Cells.CountLarge > Infra_Config.MAX_UNDO_CELLS Then
        Debug.Print "BEAVER [UNDO]: Restrained capture range is still too large to capture safely (" & captureRange.Cells.CountLarge & " cells). Skipping undo registration."
        GoTo CleanExit
    End If
    
    Dim targetWb As Workbook
    Set targetWb = Target.Worksheet.Parent
    
    Dim undoSh As Worksheet
    Set undoSh = GetUndoSheet()
    If undoSh Is Nothing Then
        Debug.Print "BEAVER [UNDO]: Could not access or create undo sheet in add-in. Skipping undo registration."
        GoTo CleanExit
    End If
    
    ' Clear previous undo data
    undoSh.Cells.Clear
    
    ' Copy captureRange to Undo Sheet at the same address so relative formulas
    ' keep their original references instead of being re-based from A1.
    captureRange.Copy Destination:=undoSh.Range(captureRange.Address)

    ' Prefix external formulas in the undo sheet to prevent external links from being active in the session.
    On Error Resume Next
    links = ThisWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    On Error GoTo 0
    
    If Not IsEmpty(links) Then
        Set undoRange = undoSh.Range(captureRange.Address)
        If undoRange.CountLarge = 1 Then
            If undoRange.HasFormula Then
                Set formulaCells = undoRange
            End If
        Else
            On Error Resume Next
            Set formulaCells = undoRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
        End If
        
        If Not formulaCells Is Nothing Then
            Set extCells = New Collection
            Dim hasUndoExt As Boolean
            hasUndoExt = False
            
            If formulaCells.Cells.CountLarge = 1 Then
                If InStr(1, formulaCells.Formula2, "[", vbTextCompare) > 0 Then
                    extCells.Add formulaCells
                    hasUndoExt = True
                End If
            Else
                On Error Resume Next
                Set foundCell = formulaCells.Find(What:="[", LookIn:=xlFormulas, LookAt:=xlPart)
                If Not foundCell Is Nothing Then
                    firstAddress = foundCell.Address
                    Do
                        extCells.Add foundCell
                        Set foundCell = formulaCells.FindNext(foundCell)
                        If foundCell Is Nothing Then Exit Do
                    Loop While foundCell.Address <> firstAddress
                    hasUndoExt = True
                End If
                On Error GoTo 0
            End If
            
            If hasUndoExt Then
                For Each c In extCells
                    c.Value = "__BEAVER_UNDO_FORMULA_PREFIX__" & c.Formula2
                Next c
            End If
        End If
        
        ' Clean up any active external links created in ThisWorkbook (the add-in) due to the copy.
        ' Only run if ThisWorkbook is NOT the target workbook to prevent breaking links prematurely.
        If Not ThisWorkbook Is targetWb Then
            For i = LBound(links) To UBound(links)
                On Error Resume Next
                ThisWorkbook.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
                On Error GoTo 0
            Next i
        End If
    End If

    ' Store metadata outside the undo payload so large target ranges cannot
    ' overwrite it.
    StoreUndoMetadata Target.Worksheet.Parent.Name, Target.Worksheet.Name, captureRange.Address, ActionName
    
    ' Stage the Undo macro (registration happens later)
    m_PendingUndoAction = ActionName
    
    ' Clean up clipboard
    Application.CutCopyMode = False
    
    SaveState = True

CleanExit:
    Exit Function
ErrHandler:
    ClearUndoMetadata
    SaveState = False
    Infra_Error.HandleError "SaveState", Err
    Resume CleanExit
End Function

' Captures the state of a range and registers an Undo action.
' If capture fails (e.g. range size exceeds safety limits), explicitly prompts or warns the user before proceeding.
' Returns True if the action should proceed, False if cancelled.
Public Function SaveStateOrConfirm(ByVal Target As Range, ByVal ActionName As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("SaveStateOrConfirm")
    On Error GoTo ErrHandler

    SaveStateOrConfirm = True
    If Target Is Nothing Then GoTo CleanExit

    If Not SaveState(Target, ActionName) Then
        SaveStateOrConfirm = Infra_Interaction.Confirm( _
            "The selected range is too large to support Undo for '" & ActionName & "'." & vbCrLf & vbCrLf & _
            "Do you want to proceed with the operation anyway?", _
            ActionName & " - Undo Warning", vbDefaultButton2)
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "SaveStateOrConfirm", Err
    SaveStateOrConfirm = False
    Resume CleanExit
End Function

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
    
    Dim undoSh As Worksheet
    Set undoSh = GetUndoSheet()
    If undoSh Is Nothing Then GoTo CleanExit
    
    Dim targetRange As Range
    Set targetRange = targetWs.Range(addr)
    
    ' Restore data from the same address it was captured at so formula
    ' references are restored exactly as they were before the mutation.
    Dim dataRange As Range
    Set dataRange = undoSh.Range(addr)

    dataRange.Copy Destination:=targetRange
    
    ' Restore formulas by replacing the prefix back to empty string
    On Error Resume Next
    targetRange.Replace What:="__BEAVER_UNDO_FORMULA_PREFIX__", Replacement:="", LookAt:=xlPart
    On Error GoTo 0
    
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

' Returns (and creates if necessary) the hidden undo sheet in the specified workbook.
Private Function GetUndoSheet() As Worksheet
    ' Internal helper - always use ThisWorkbook to avoid modifying user's workbook structure
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
