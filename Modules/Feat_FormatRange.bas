Attribute VB_Name = "Feat_FormatRange"
Option Explicit

' @Module: Feat_FormatRange
' @Category: Feature
' @Description: Bulk formatting tools for tables, custom numbering, and specialized pasting.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_AppStateGuard, Infra_Config, Infra_Error, Infra_Undo, Infra_Progress

Public Sub FormatSelectedRange()
    PushContext "FormatSelectedRange"
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim ws As Worksheet
    Dim selRange As Range
    Dim col As Range
    Dim tbl As ListObject
    Dim firstDataCell As Range
    Dim dataRange As Range

    ' --- Step 1: Ensure user selected a range ---
    If Not Infra_AppState.IsRangeSelected() Then
        MsgBox "Please select a range before running the macro.", vbExclamation, Infra_Config.ADDIN_NAME
        GoTo CleanExit
    End If

    Set ws = ActiveSheet
    Set selRange = Selection
    If ws Is Nothing Or selRange Is Nothing Then GoTo CleanExit

    ' --- Step 2: Save State for Undo ---
    Infra_Undo.SaveState selRange, "Format Range"

    ' --- Step 3: Initialize Progress ---
    Infra_Progress.StartProgress "Formatting Range", 100
    Infra_Progress.UpdateProgress 10

    ' --- Step 4: Convert overlapping tables to plain ranges (Optimized) ---
    For Each tbl In ws.ListObjects
        If Not Intersect(tbl.Range, selRange) Is Nothing Then
            tbl.Unlist
        End If
    Next tbl
    Infra_Progress.UpdateProgress 20

    ' --- Step 5: Handle empty selection ---
    If WorksheetFunction.CountA(selRange) = 0 Then
        MsgBox "The selected range is empty. Nothing to format.", vbInformation, Infra_Config.ADDIN_NAME
        GoTo CleanExit
    End If

    ' --- Step 6: Unmerge cells ---
    selRange.UnMerge
    Infra_Progress.UpdateProgress 30

    ' --- Step 7: General formatting ---
    With selRange
        .WrapText = False
        .Font.Name = Infra_Config.DEFAULT_FONT_NAME
        .Font.Size = Infra_Config.DEFAULT_FONT_SIZE
        .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlContinuous
        .EntireRow.AutoFit
        .EntireColumn.AutoFit
    End With
    Infra_Progress.UpdateProgress 50

    ' --- Step 8: Cap wide columns ---
    Dim colCount As Long, i As Long
    colCount = selRange.Columns.Count
    For i = 1 To colCount
        Set col = selRange.Columns(i)
        If col.ColumnWidth > Infra_Config.COLUMN_WIDTH_THRESHOLD Then
            col.ColumnWidth = Infra_Config.MAX_COLUMN_WIDTH
        End If
        If i Mod 5 = 0 Then Infra_Progress.UpdateProgress 50 + (i / colCount) * 20
    Next i

    ' --- Step 9: Column-specific formatting (Date / Number) ---
    If selRange.Rows.Count > 1 Then
        Set dataRange = selRange.Offset(1, 0).Resize(selRange.Rows.Count - 1)

        For i = 1 To colCount
            Set col = selRange.Columns(i)
            Set firstDataCell = Intersect(dataRange, col).Cells(1, 1)

            If Not IsEmpty(firstDataCell.Value) Then
                If IsDate(firstDataCell.Value) Then
                    Intersect(dataRange, col).NumberFormat = Infra_Config.DISPLAY_DATE_FORMAT
                ElseIf IsNumeric(firstDataCell.Value) Then
                    Intersect(dataRange, col).NumberFormat = Infra_Config.DEFAULT_NUMBER_FORMAT
                End If
            End If
            If i Mod 5 = 0 Then Infra_Progress.UpdateProgress 70 + (i / colCount) * 20
        Next i
    End If

    ' --- Step 10: Header row formatting ---
    With selRange.Rows(1)
        .Font.Bold = True
        .Font.Size = Infra_Config.HEADER_FONT_SIZE
        .Interior.Color = Infra_Config.HEADER_COLOR
    End With
    Infra_Progress.UpdateProgress 95

    ' --- Step 11: Apply AutoFilter ---
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    selRange.Rows(1).AutoFilter

    ' --- Step 12: Land cursor on first cell ---
    selRange.Cells(1, 1).Select
    Infra_Progress.UpdateProgress 100

CleanExit:
    Infra_Progress.EndProgress
    PopContext
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    HandleError "FormatSelectedRange", Err
End Sub

' Applies a custom number format to the selected range.
Public Sub ApplyCustomNumberFormat()
    PushContext "ApplyCustomNumberFormat"
    On Error GoTo ErrHandler
    
    Dim guard As New Infra_AppStateGuard
    Dim targetRange As Range

    If Not Infra_AppState.IsRangeSelected() Then GoTo CleanExit
    Set targetRange = Selection
    If targetRange Is Nothing Then GoTo CleanExit

    targetRange.NumberFormat = Infra_Config.DEFAULT_NUMBER_FORMAT

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "ApplyCustomNumberFormat", Err
End Sub

' Pastes cell formatting only. Requires user to have copied something first.
Public Sub PasteFormat()
    PushContext "PasteFormat"
    On Error GoTo ErrHandler
    
    Dim guard As New Infra_AppStateGuard
    Dim targetRange As Range

    If Not Infra_AppState.IsRangeSelected() Then GoTo CleanExit
    Set targetRange = Selection
    If targetRange Is Nothing Then GoTo CleanExit

    On Error Resume Next
    targetRange.PasteSpecial Paste:=xlPasteFormats
    
    If Err.Number <> 0 Then
        MsgBox "Nothing to paste. Please copy a range first.", vbExclamation, Infra_Config.ADDIN_NAME
    End If
    On Error GoTo ErrHandler

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "PasteFormat", Err
End Sub
