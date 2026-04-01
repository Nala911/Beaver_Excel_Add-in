VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_FormatRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_FormatRange
' @Category: Feature
' @Description: Bulk formatting tools for tables, custom numbering, and specialized pasting.

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("FormatRange.CanExecute")
    On Error GoTo ErrHandler
    ICommand_CanExecute = Infra_AppState.IsRangeSelected()
    Exit Function
ErrHandler:
    Infra_Error.HandleError "FormatRange.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("FormatRange.Execute")
    On Error GoTo ErrHandler

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
        .Font.Name = Infra_Config.Model.DefaultFontName
        .Font.Size = Infra_Config.Model.DefaultFontSize
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
        If col.ColumnWidth > Infra_Config.Model.ColumnWidthThreshold Then
            col.ColumnWidth = Infra_Config.Model.MaxColumnWidth
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
                    Intersect(dataRange, col).NumberFormat = Infra_Config.Model.DisplayDateFormat
                ElseIf IsNumeric(firstDataCell.Value) Then
                    Intersect(dataRange, col).NumberFormat = Infra_Config.Model.DefaultNumberFormat
                End If
            End If
            If i Mod 5 = 0 Then Infra_Progress.UpdateProgress 70 + (i / colCount) * 20
        Next i
    End If

    ' --- Step 10: Header row formatting ---
    With selRange.Rows(1)
        .Font.Bold = True
        .Font.Size = Infra_Config.Model.HeaderFontSize
        .Interior.Color = Infra_Config.Model.HeaderColor
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
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    Infra_Error.HandleError "FormatRange.Execute", Err
    Resume CleanExit
End Sub
