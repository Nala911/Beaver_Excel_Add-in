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
Private Const UNDO_META_CAPTURE_MODE As String = "BeaverUndoCaptureMode"
Private m_PendingUndoAction As String
Private m_UseMemoryUndo As Boolean
Private m_MemoryValues As Collection
Private m_MemoryFormulas As Collection
Private m_MemoryFormats As Collection

' Captures the state of a range and registers an Undo action.
' Call this BEFORE modifying the range.
Public Function SaveState(ByVal Target As Range, ByVal ActionName As String, Optional ByVal CaptureMode As UndoCaptureMode = UndoCaptureFull) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("SaveState")
    Dim links As Variant
    Dim undoRange As Range
    Dim formulaCells As Range
    Dim foundCell As Range
    Dim firstAddress As String
    Dim extCells As Collection
    Dim c As Variant
    Dim i As Long
    Dim area As Range
    On Error GoTo ErrHandler
    
    If Target Is Nothing Then GoTo CleanExit
    
    ' Detect if the target range contains any legacy CSE array formulas.
    ' If it does, we must fallback to UndoCaptureFull because formula/value assignment
    ' via .Formula2 cannot restore CSE array boundaries.
    Dim actualCaptureMode As UndoCaptureMode
    actualCaptureMode = CaptureMode
    
    If CaptureMode = UndoCaptureFormulaOnly Then
        Dim hasLegacyArray As Boolean
        hasLegacyArray = False
        On Error Resume Next
        If IsNull(Target.HasArray) Then
            hasLegacyArray = True
        ElseIf Target.HasArray Then
            hasLegacyArray = True
        End If
        On Error GoTo ErrHandler
        
        If hasLegacyArray Then
            actualCaptureMode = UndoCaptureFull
        End If
    End If
    
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
    
    m_UseMemoryUndo = False
    If (actualCaptureMode = UndoCaptureFormulaOnly Or actualCaptureMode = UndoCaptureFormatOnly Or actualCaptureMode = UndoCaptureValueOnly) And captureRange.Cells.CountLarge <= 20000 Then
        m_UseMemoryUndo = True
        Set m_MemoryValues = New Collection
        Set m_MemoryFormulas = New Collection
        Set m_MemoryFormats = New Collection
        
        For Each area In captureRange.Areas
            If actualCaptureMode = UndoCaptureValueOnly Then
                m_MemoryValues.Add area.Value2
            Else
                m_MemoryValues.Add area.Value
            End If
            m_MemoryFormulas.Add area.Formula2
            m_MemoryFormats.Add area.NumberFormat
        Next area
        
        StoreUndoMetadata Target.Worksheet.Parent.Name, Target.Worksheet.Name, captureRange.Address, ActionName, actualCaptureMode
        m_PendingUndoAction = ActionName
        SaveState = True
        GoTo CleanExit
    End If
    
    Dim undoSh As Worksheet
    Set undoSh = GetUndoSheet()
    If undoSh Is Nothing Then
        Debug.Print "BEAVER [UNDO]: Could not access or create undo sheet in add-in. Skipping undo registration."
        GoTo CleanExit
    End If
    
    ' Clear previous undo data
    undoSh.Cells.Clear
    
    ' Copy captureRange to Undo Sheet
    If actualCaptureMode = UndoCaptureFormulaOnly Then
        For Each area In captureRange.Areas
            Dim formulas As Variant
            formulas = area.Formula2
            
            If IsArray(formulas) Then
                Dim r As Long, col As Long
                For r = LBound(formulas, 1) To UBound(formulas, 1)
                    For col = LBound(formulas, 2) To UBound(formulas, 2)
                        Dim cellForm As Variant
                        cellForm = formulas(r, col)
                        If VarType(cellForm) = vbString Then
                            If Left$(cellForm, 1) = "=" And InStr(1, cellForm, "[", vbTextCompare) > 0 Then
                                formulas(r, col) = "__BEAVER_UNDO_FORMULA_PREFIX__" & cellForm
                            End If
                        End If
                    Next col
                Next r
                undoSh.Range(area.Address).Formula2 = formulas
            Else
                If VarType(formulas) = vbString Then
                    If Left$(formulas, 1) = "=" And InStr(1, formulas, "[", vbTextCompare) > 0 Then
                        formulas = "__BEAVER_UNDO_FORMULA_PREFIX__" & formulas
                    End If
                End If
                undoSh.Range(area.Address).Formula2 = formulas
            End If
        Next area
    ElseIf actualCaptureMode = UndoCaptureFormatOnly Then
        For Each area In captureRange.Areas
            Dim formats As Variant
            formats = area.NumberFormat
            undoSh.Range(area.Address).NumberFormat = formats
        Next area
    ElseIf actualCaptureMode = UndoCaptureValueOnly Then
        For Each area In captureRange.Areas
            undoSh.Range(area.Address).Value2 = area.Value2
        Next area
    Else
        For Each area In captureRange.Areas
            area.Copy Destination:=undoSh.Range(area.Address)
        Next area

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
                        If IsActualExternalLink(formulaCells.Formula2) Then
                            extCells.Add formulaCells
                            hasUndoExt = True
                        End If
                    End If
                Else
                    On Error Resume Next
                    Set foundCell = formulaCells.Find(What:="[", LookIn:=xlFormulas, LookAt:=xlPart)
                    If Not foundCell Is Nothing Then
                        firstAddress = foundCell.Address
                        Do
                            If IsActualExternalLink(foundCell.Formula2) Then
                                extCells.Add foundCell
                                hasUndoExt = True
                            End If
                            Set foundCell = formulaCells.FindNext(foundCell)
                            If foundCell Is Nothing Then Exit Do
                        Loop While foundCell.Address <> firstAddress
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
    End If

    StoreUndoMetadata Target.Worksheet.Parent.Name, Target.Worksheet.Name, captureRange.Address, ActionName, actualCaptureMode
    
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
Public Function SaveStateOrConfirm(ByVal Target As Range, ByVal ActionName As String, Optional ByVal CaptureMode As UndoCaptureMode = UndoCaptureFull) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("SaveStateOrConfirm")
    On Error GoTo ErrHandler

    SaveStateOrConfirm = True
    If Target Is Nothing Then GoTo CleanExit

    If Not SaveState(Target, ActionName, CaptureMode) Then
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

' Saves the list of created named ranges for custom Undo.
Public Sub SaveCreatedNamesState(ByVal targetWb As Workbook, ByVal namesList As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("SaveCreatedNamesState")
    On Error GoTo ErrHandler
    
    ' Clear normal undo metadata and sheet contents
    Dim undoSh As Worksheet
    Set undoSh = GetUndoSheet()
    If Not undoSh Is Nothing Then undoSh.Cells.Clear
    ClearUndoMetadata
    
    ' Save metadata
    SetUndoMetadataValue UNDO_META_WORKBOOK_NAME, targetWb.Name
    SetUndoMetadataValue UNDO_META_ACTION_NAME, "Create Named Ranges"
    SetUndoMetadataValue "BeaverUndoCreatedNames", namesList
    
    m_PendingUndoAction = "Create Named Ranges"
    
CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "SaveCreatedNamesState", Err
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
    m_UseMemoryUndo = False
    Set m_MemoryValues = Nothing
    Set m_MemoryFormulas = Nothing
    Set m_MemoryFormats = Nothing

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
    
    Dim actionName As String: actionName = GetUndoMetadataValue(UNDO_META_ACTION_NAME)
    Dim wbName As String: wbName = GetUndoMetadataValue(UNDO_META_WORKBOOK_NAME)
    
    If actionName = "Create Named Ranges" Then
        Dim createdNamesStr As String
        createdNamesStr = GetUndoMetadataValue("BeaverUndoCreatedNames")
        
        Dim targetWb As Workbook
        On Error Resume Next
        Set targetWb = Workbooks(wbName)
        If targetWb Is Nothing Then Set targetWb = ActiveWorkbook
        On Error GoTo ErrHandler
        
        If Not targetWb Is Nothing And createdNamesStr <> "" Then
            Dim namesArr() As String
            namesArr = Split(createdNamesStr, ";")
            
            Dim i As Long
            For i = LBound(namesArr) To UBound(namesArr)
                Dim nameToDelete As String
                nameToDelete = namesArr(i)
                If nameToDelete <> "" Then
                    On Error Resume Next
                    targetWb.Names(nameToDelete).Delete
                    On Error GoTo ErrHandler
                End If
            Next i
        End If
        
        DeleteUndoMetadataValue "BeaverUndoCreatedNames"
        ClearUndoMetadata
        GoTo CleanExit
    End If
    
    Dim wsName As String: wsName = GetUndoMetadataValue(UNDO_META_WORKSHEET_NAME)
    Dim addr As String: addr = GetUndoMetadataValue(UNDO_META_ADDRESS_NAME)
    Dim capModeStr As String: capModeStr = GetUndoMetadataValue(UNDO_META_CAPTURE_MODE)
    Dim capMode As UndoCaptureMode
    If capModeStr <> "" Then
        capMode = CInt(capModeStr)
    End If
    
    If wbName = "" Or wsName = "" Or addr = "" Then GoTo CleanExit
    
    Set targetWb = Nothing
    On Error Resume Next
    Set targetWb = Workbooks(wbName)
    If targetWb Is Nothing Then
        ' Fallback: maybe it's the active workbook?
        Set targetWb = ActiveWorkbook
    End If
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
    
    If m_UseMemoryUndo And Not m_MemoryValues Is Nothing Then
        Dim areaObj As Range
        Dim areaIdx As Long
        areaIdx = 1
        For Each areaObj In targetRange.Areas
            If capMode = UndoCaptureFormulaOnly Then
                areaObj.Formula2 = m_MemoryFormulas(areaIdx)
            ElseIf capMode = UndoCaptureFormatOnly Then
                areaObj.NumberFormat = m_MemoryFormats(areaIdx)
            ElseIf capMode = UndoCaptureValueOnly Then
                areaObj.Value2 = m_MemoryValues(areaIdx)
            Else
                areaObj.Value = m_MemoryValues(areaIdx)
                areaObj.NumberFormat = m_MemoryFormats(areaIdx)
            End If
            areaIdx = areaIdx + 1
        Next areaObj
        
        Set m_MemoryValues = Nothing
        Set m_MemoryFormulas = Nothing
        Set m_MemoryFormats = Nothing
        m_UseMemoryUndo = False
        ClearUndoMetadata
        Application.CutCopyMode = False
        
        On Error Resume Next
        targetRange.Select
        On Error GoTo ErrHandler
        GoTo CleanExit
    End If
    
    ' Restore data
    Dim area As Range
    For Each area In targetRange.Areas
        If capMode = UndoCaptureFormulaOnly Then
            area.Formula2 = undoSh.Range(area.Address).Formula2
        ElseIf capMode = UndoCaptureFormatOnly Then
            area.NumberFormat = undoSh.Range(area.Address).NumberFormat
        ElseIf capMode = UndoCaptureValueOnly Then
            area.Value2 = undoSh.Range(area.Address).Value2
        Else
            undoSh.Range(area.Address).Copy Destination:=area
        End If
    Next area
    
    ' Restore formulas by replacing the prefix back to empty string
    If capMode <> UndoCaptureFormatOnly Then
        On Error Resume Next
        targetRange.Replace What:="__BEAVER_UNDO_FORMULA_PREFIX__", Replacement:="", LookAt:=xlPart
        On Error GoTo 0
    End If
    
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

Private Sub StoreUndoMetadata(ByVal WorkbookName As String, ByVal WorksheetName As String, ByVal AddressText As String, ByVal ActionName As String, ByVal CaptureMode As UndoCaptureMode)
    Dim tracker As Object: Set tracker = Infra_Error.Track("StoreUndoMetadata")
    On Error GoTo ErrHandler

    SetUndoMetadataValue UNDO_META_WORKBOOK_NAME, WorkbookName
    SetUndoMetadataValue UNDO_META_WORKSHEET_NAME, WorksheetName
    SetUndoMetadataValue UNDO_META_ADDRESS_NAME, AddressText
    SetUndoMetadataValue UNDO_META_ACTION_NAME, ActionName
    SetUndoMetadataValue UNDO_META_CAPTURE_MODE, CStr(CaptureMode)

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
    DeleteUndoMetadataValue UNDO_META_CAPTURE_MODE
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

Private Function IsActualExternalLink(ByVal formulaText As String) As Boolean
    Dim openPos As Long
    Dim closePos As Long
    Dim innerText As String
    
    openPos = InStr(1, formulaText, "[", vbTextCompare)
    Do While openPos > 0
        closePos = InStr(openPos + 1, formulaText, "]", vbTextCompare)
        If closePos > openPos Then
            innerText = Mid$(formulaText, openPos + 1, closePos - openPos - 1)
            ' Check if it looks like a workbook filename (contains file extension .xls...)
            ' or is followed by ! (which indicates an external sheet reference like [Book1]Sheet1!)
            If InStr(1, innerText, ".xl", vbTextCompare) > 0 Or InStr(closePos + 1, formulaText, "!", vbTextCompare) = closePos + 1 Then
                IsActualExternalLink = True
                Exit Function
            End If
        End If
        openPos = InStr(openPos + 1, formulaText, "[", vbTextCompare)
    Loop
    IsActualExternalLink = False
End Function
