VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_FillDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_FillDown
' @Category: Feature
' @Description: Intelligent fill-down functionality that respects tables and filtered data.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyBulkWrite()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("FillDown.Execute")
    On Error GoTo ErrHandler

    Dim selectionRng As Range
    Dim ws As Worksheet
    Dim startRow As Long
    Dim startCol As Long, endCol As Long
    Dim colIdx As Long
    Dim sourceCell As Range
    Dim visibleRange As Range
    Dim fullGroupRange As Range
    
    Dim ctx As ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit
    Set selectionRng = ctx.SelectionRange
    Set ws = selectionRng.Worksheet

    ' Use Cells.CountLarge to prevent overflow when selecting huge ranges
    If selectionRng.Cells.CountLarge > 1 Then
        startRow = selectionRng.Row + selectionRng.Rows.Count - 1
    Else
        startRow = selectionRng.Row
    End If

    startCol = selectionRng.Column
    endCol = selectionRng.Column + selectionRng.Columns.Count - 1

    ' 1. Cache UsedRange limit
    Dim maxSearchCol As Long
    maxSearchCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    ' 2. Cache ListObject at the selection
    Dim cachedListObject As ListObject
    On Error Resume Next
    Set cachedListObject = selectionRng.ListObject
    On Error GoTo ErrHandler

    ' 3. Cache CurrentRegion of selection
    Dim cachedCurrentRegion As Range
    On Error Resume Next
    Set cachedCurrentRegion = selectionRng.CurrentRegion
    On Error GoTo ErrHandler

    ' 4. Cache proximity row values in a single COM query
    Dim maxProximityDist As Long
    maxProximityDist = Infra_Config.MAX_FILL_PROXIMITY_COLUMNS

    Dim proximityStartCol As Long
    proximityStartCol = IIf(startCol - maxProximityDist < 1, 1, startCol - maxProximityDist)
    
    Dim proximityEndCol As Long
    proximityEndCol = IIf(endCol + maxProximityDist > maxSearchCol, maxSearchCol, endCol + maxProximityDist)
    
    Dim cachedRowValues As Variant
    If proximityEndCol >= proximityStartCol Then
        cachedRowValues = ws.Range(ws.Cells(startRow, proximityStartCol), ws.Cells(startRow, proximityEndCol)).Value2
    End If

    ' Collect target fill ranges and source cells to update
    Dim colSourceCells As New Collection
    Dim colVisibleRanges As New Collection
    
    Dim undoRange As Range

    ' We will loop through the columns and group contiguous columns that share the same lastRow
    Dim groupStartCol As Long
    Dim currentGroupLastRow As Long
    
    groupStartCol = startCol
    currentGroupLastRow = -1
    
    For colIdx = startCol To endCol
        Set sourceCell = ws.Cells(startRow, colIdx)
        
        Dim lastRow As Long
        lastRow = Infra_CommandSupport.GetFillDownBoundary( _
            sourceCell, _
            maxSearchCol, _
            cachedListObject, _
            cachedCurrentRegion, _
            cachedRowValues, _
            proximityStartCol _
        )
        
        ' We want to see if we can continue the current group
        If colIdx = startCol Then
            currentGroupLastRow = lastRow
        Else
            If lastRow <> currentGroupLastRow Then
                ' Process the group from groupStartCol to colIdx - 1
                ProcessFillDownGroup ws, startRow, groupStartCol, colIdx - 1, currentGroupLastRow, colSourceCells, colVisibleRanges, undoRange
                ' Start a new group
                groupStartCol = colIdx
                currentGroupLastRow = lastRow
            End If
        End If
    Next colIdx
    
    ' Process the final group
    ProcessFillDownGroup ws, startRow, groupStartCol, endCol, currentGroupLastRow, colSourceCells, colVisibleRanges, undoRange

    ' --- EXECUTION ---
    If colSourceCells.Count = 0 Then GoTo CleanExit

    ' Safety Check: Sum up the areas of all columns' visible ranges to prevent fragmentation crashes
    Dim totalAreas As Long
    Dim i As Long
    For i = 1 To colVisibleRanges.Count
        totalAreas = totalAreas + colVisibleRanges(i).Areas.Count
    Next i
    
    If totalAreas > Infra_Config.MAX_FILL_DOWN_AREAS Then
        Err.Raise vbObjectError + 1001, "FillDown.Execute", _
            "The target range is too fragmented (exceeds " & Format(Infra_Config.MAX_FILL_DOWN_AREAS, "#,##0") & " non-contiguous areas). Action aborted to prevent Excel from freezing."
    End If

    ' Register Undo for the combined range
    If Not undoRange Is Nothing Then
        If Not Infra_Undo.SaveStateOrConfirm(undoRange, "Fill Down") Then GoTo CleanExit
    End If

    ' Copy/Paste values for each column (or use native FillDown if no cells are hidden)
    For i = 1 To colSourceCells.Count
        Set sourceCell = colSourceCells(i)
        Set visibleRange = colVisibleRanges(i)
        
        Set fullGroupRange = ws.Range(sourceCell, visibleRange)
        If visibleRange.Cells.CountLarge = fullGroupRange.Cells.CountLarge - sourceCell.Cells.CountLarge Then
            fullGroupRange.FillDown
        Else
            sourceCell.Copy
            visibleRange.PasteSpecial Paste:=xlPasteAll
        End If
    Next i
    
    Application.CutCopyMode = False

CleanExit:
    Exit Sub

ErrHandler:
    If Err.Number = vbObjectError + 1001 Then
        Infra_Error.HandleErrorDetailed "FillDown.Execute", Err, Nothing, Infra_Error.CATEGORY_VALIDATION, Infra_Error.SEVERITY_WARNING
    Else
        Infra_Error.HandleError "FillDown.Execute", Err
    End If
    Resume CleanExit
End Sub

Private Sub ProcessFillDownGroup( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal groupStartCol As Long, _
    ByVal groupEndCol As Long, _
    ByVal groupLastRow As Long, _
    ByVal colSourceCells As Collection, _
    ByVal colVisibleRanges As Collection, _
    ByRef undoRange As Range)
    
    Dim tracker As Object: Set tracker = Infra_Error.Track("FillDown.ProcessFillDownGroup")
    On Error GoTo ErrHandler
    
    If groupLastRow <= startRow Then Exit Sub
    
    Dim groupFillRange As Range
    Set groupFillRange = ws.Range(ws.Cells(startRow + 1, groupStartCol), ws.Cells(groupLastRow, groupEndCol))
    
    Dim visibleRange As Range
    If groupFillRange.CountLarge = 1 Then
        If Not groupFillRange.EntireRow.Hidden And Not groupFillRange.EntireColumn.Hidden Then
            Set visibleRange = groupFillRange
        End If
    Else
        On Error Resume Next
        Set visibleRange = groupFillRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo ErrHandler
    End If
    
    If Not visibleRange Is Nothing Then
        If visibleRange.Areas.Count = 1 And visibleRange.Cells.CountLarge = groupFillRange.Cells.CountLarge Then
            ' No hidden cells! We can do bulk copy-paste for the whole group.
            colSourceCells.Add ws.Range(ws.Cells(startRow, groupStartCol), ws.Cells(startRow, groupEndCol))
            colVisibleRanges.Add groupFillRange
        Else
            ' There are hidden cells. Fallback to column-by-column processing for this group.
            Dim cIdx As Long
            For cIdx = groupStartCol To groupEndCol
                Dim colSource As Range
                Set colSource = ws.Cells(startRow, cIdx)
                Dim colFill As Range
                Set colFill = ws.Range(ws.Cells(startRow + 1, cIdx), ws.Cells(groupLastRow, cIdx))
                
                Dim colVisible As Range
                If colFill.CountLarge = 1 Then
                    If Not colFill.EntireRow.Hidden And Not colFill.EntireColumn.Hidden Then
                        Set colVisible = colFill
                    End If
                Else
                    On Error Resume Next
                    Set colVisible = colFill.SpecialCells(xlCellTypeVisible)
                    On Error GoTo ErrHandler
                End If
                
                If Not colVisible Is Nothing Then
                    colSourceCells.Add colSource
                    colVisibleRanges.Add colVisible
                End If
            Next cIdx
        End If
        
        ' Accumulate the full modified area for a single consolidated undo action
        Dim fullGroupRange As Range
        Set fullGroupRange = ws.Range(ws.Cells(startRow, groupStartCol), ws.Cells(groupLastRow, groupEndCol))
        If undoRange Is Nothing Then
            Set undoRange = fullGroupRange
        Else
            Set undoRange = Union(undoRange, fullGroupRange)
        End If
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "FillDown.ProcessFillDownGroup", Err
    Resume CleanExit
End Sub

