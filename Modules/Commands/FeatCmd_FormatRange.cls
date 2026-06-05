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

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyBulkWrite()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
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
    Dim overlappingTableCount As Long
    Dim hasMergedCells As Boolean

    Dim ctx As Infra_ActionContext

    ' --- Step 1: Ensure user selected a range ---
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then
        GoTo CleanExit
    End If

    Set ws = ctx.WorksheetRef
    Set selRange = ctx.SelectionRange
    If ws Is Nothing Or selRange Is Nothing Then GoTo CleanExit

    ' --- Step 2: Initialize Progress ---
    Infra_Progress.StartProgress "Formatting Range", 100
    Infra_Progress.UpdateProgress 10

    overlappingTableCount = CountOverlappingTables(ws, selRange)
    hasMergedCells = selRange.MergeCells

    If overlappingTableCount > 0 Or hasMergedCells Then
        If Not ConfirmStructuralFormatting(overlappingTableCount, hasMergedCells) Then GoTo CleanExit
    End If

    ' --- Step 3: Save State for Undo ---
    Infra_Undo.SaveState selRange, "Format Range"

    FormatRangeDirect selRange, ws

CleanExit:
    Infra_Progress.EndProgress
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    Infra_Error.HandleError "FormatRange.Execute", Err
    Resume CleanExit
End Sub

Private Function CountOverlappingTables(ByVal ws As Worksheet, ByVal selRange As Range) As Long
    Dim tbl As ListObject

    If ws Is Nothing Then Exit Function
    If selRange Is Nothing Then Exit Function

    For Each tbl In ws.ListObjects
        If Not Intersect(tbl.Range, selRange) Is Nothing Then
            CountOverlappingTables = CountOverlappingTables + 1
        End If
    Next tbl
End Function

Private Function ConfirmStructuralFormatting(ByVal overlappingTableCount As Long, ByVal hasMergedCells As Boolean) As Boolean
    Dim message As String

    message = "Format Range will make structural changes before applying formatting." & vbCrLf & vbCrLf

    If overlappingTableCount > 0 Then
        message = message & "Excel tables to convert to normal ranges: " & overlappingTableCount & vbCrLf
    End If

    If hasMergedCells Then
        message = message & "Merged cells in the selection will be unmerged." & vbCrLf
    End If

    message = message & vbCrLf & "Continue?"
    ConfirmStructuralFormatting = Infra_Interaction.Confirm(message, Infra_Config.ADDIN_NAME & " - Format Range", vbDefaultButton2)
End Function

' --- Testing Entry Point ---
Public Sub FormatRangeDirect(ByVal selRange As Range, ByVal ws As Worksheet)
    Dim tracker As Object: Set tracker = Infra_Error.Track("FormatRangeDirect")
    On Error GoTo ErrHandler

    Dim col As Range
    Dim tbl As ListObject
    Dim firstDataCell As Range
    Dim dataRange As Range
    Dim colCount As Long
    Dim i As Long

    If ws Is Nothing Or selRange Is Nothing Then GoTo CleanExit

    ' --- Step 4: Convert overlapping tables to plain ranges (Optimized) ---
    For Each tbl In ws.ListObjects
        If Not Intersect(tbl.Range, selRange) Is Nothing Then
            tbl.Unlist
        End If
    Next tbl
    Infra_Progress.UpdateProgress 20

    ' --- Step 5: Handle empty selection ---
    If WorksheetFunction.CountA(selRange) = 0 Then
        Infra_Interaction.ShowInfo "The selected range is empty. Nothing to format."
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
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "FormatRangeDirect", Err
    Resume CleanExit
End Sub
