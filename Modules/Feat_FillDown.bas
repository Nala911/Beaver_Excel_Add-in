Attribute VB_Name = "Feat_FillDown"
Option Explicit

' @Module: Feat_FillDown
' @Category: Feature
' @Description: Intelligent fill-down functionality that respects tables and filtered data.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_AppStateGuard, Infra_Config, Infra_Error

Public Sub FillDown()
    PushContext "FillDown"
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim sourceRng As Range
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim startRow As Long, lastRow As Long
    Dim targetCol As Long, maxSearchCol As Long
    Dim dist As Long
    Dim leftCol As Long, rightCol As Long
    Dim foundRef As Boolean
    Dim fillRange As Range, visibleRange As Range

    If Not Infra_AppState.IsRangeSelected() Then GoTo CleanExit
    Set sourceRng = Selection

    Set ws = sourceRng.Worksheet

    ' If multiple cells selected, grab the last row, first column of that selection
    If sourceRng.Cells.Count > 1 Then
        Set sourceRng = sourceRng.Cells(sourceRng.Rows.Count, 1)
    End If

    startRow = sourceRng.Row
    lastRow = 0
    foundRef = False

    ' --- PRIORITY 1: EXCEL TABLE (ListObject) ---
    On Error Resume Next
    Set lo = sourceRng.ListObject
    On Error GoTo ErrHandler

    If Not lo Is Nothing Then
        If Not Intersect(sourceRng, lo.DataBodyRange) Is Nothing Then
            lastRow = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
            foundRef = True
        End If
    End If

    ' --- PRIORITY 2: CURRENT REGION ---
    If Not foundRef Then
        Dim rRegion As Range
        Set rRegion = sourceRng.CurrentRegion

        If rRegion.Cells(rRegion.Cells.Count).Address <> sourceRng.Address Then
            lastRow = rRegion.Row + rRegion.Rows.Count - 1
            foundRef = True
        End If
    End If

    ' --- PRIORITY 3: PROXIMITY SEARCH (closest neighbour) ---
    If Not foundRef Then
        maxSearchCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

        For dist = 1 To maxSearchCol
            leftCol = sourceRng.Column - dist
            If leftCol >= 1 Then
                If Not IsEmpty(ws.Cells(startRow, leftCol)) Then
                    targetCol = leftCol
                    foundRef = True
                    Exit For
                End If
            End If

            rightCol = sourceRng.Column + dist
            if rightCol <= maxSearchCol Then
                If Not IsEmpty(ws.Cells(startRow, rightCol)) Then
                    targetCol = rightCol
                    foundRef = True
                    Exit For
                End If
            End If

            If leftCol < 1 And rightCol > maxSearchCol Then Exit For
        Next dist

        If foundRef Then
            If IsEmpty(ws.Cells(startRow + 1, targetCol)) Then
                lastRow = startRow
            Else
                lastRow = ws.Cells(startRow, targetCol).End(xlDown).Row
            End If
        End If
    End If

    ' --- EXECUTION: Copy/Paste to visible cells ---
    If Not foundRef Then GoTo CleanExit
    If lastRow <= startRow Then GoTo CleanExit

    Set fillRange = ws.Range( _
        ws.Cells(startRow + 1, sourceRng.Column), _
        ws.Cells(lastRow, sourceRng.Column))

    On Error Resume Next
    Set visibleRange = fillRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo ErrHandler

    If Not visibleRange Is Nothing Then
        sourceRng.Copy
        visibleRange.PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
    End If

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "FillDown", Err
End Sub
