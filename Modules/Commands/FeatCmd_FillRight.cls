VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_FillRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: FeatCmd_FillRight
' @Category: Feature
' @Description: Intelligent fill-right functionality that respects tables and filtered data.

Implements ICommand

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyBulkWrite()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("FillRight.Execute")
    On Error GoTo ErrHandler

    Dim selectionRng As Range
    Dim ws As Worksheet
    Dim startCol As Long
    Dim startRow As Long, endRow As Long
    Dim rowIdx As Long
    Dim sourceCell As Range
    Dim visibleRange As Range
    Dim fullGroupRange As Range
    
    Dim ctx As ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit
    Set selectionRng = ctx.SelectionRange
    If selectionRng Is Nothing Then GoTo CleanExit
    Set ws = selectionRng.Worksheet

    Dim safeRng As Range
    Set safeRng = Infra_CommandSupport.GetSafeProcessingRange(selectionRng, 50000, False)
    If Not safeRng Is Nothing Then Set selectionRng = safeRng

    ' If multiple cells are selected, start filling from the rightmost column of the selection
    If selectionRng.Cells.CountLarge > 1 Then
        startCol = selectionRng.Column + selectionRng.Columns.Count - 1
    Else
        startCol = selectionRng.Column
    End If

    startRow = selectionRng.Row
    endRow = selectionRng.Row + selectionRng.Rows.Count - 1

    ' 1. Cache UsedRange limit
    Dim maxSearchRow As Long
    maxSearchRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

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

    ' 4. Cache proximity column values in a single COM query
    Dim maxProximityDist As Long
    maxProximityDist = Infra_Config.MAX_FILL_PROXIMITY_COLUMNS

    Dim proximityStartRow As Long
    proximityStartRow = IIf(startRow - maxProximityDist < 1, 1, startRow - maxProximityDist)
    
    Dim proximityEndRow As Long
    proximityEndRow = IIf(endRow + maxProximityDist > maxSearchRow, maxSearchRow, endRow + maxProximityDist)
    
    Dim cachedColValues As Variant
    If proximityEndRow >= proximityStartRow Then
        cachedColValues = ws.Range(ws.Cells(proximityStartRow, startCol), ws.Cells(proximityEndRow, startCol)).Value2
    End If

    ' Collect target fill ranges and source cells to update
    Dim rowSourceCells As New Collection
    Dim rowVisibleRanges As New Collection
    
    Dim undoRange As Range

    ' We will loop through the rows and group contiguous rows that share the same lastCol
    Dim groupStartRow As Long
    Dim currentGroupLastCol As Long
    
    groupStartRow = startRow
    currentGroupLastCol = -1
    
    For rowIdx = startRow To endRow
        Set sourceCell = ws.Cells(rowIdx, startCol)
        
        Dim lastCol As Long
        lastCol = Infra_CommandSupport.GetFillRightBoundary( _
            sourceCell, _
            maxSearchRow, _
            cachedListObject, _
            cachedCurrentRegion, _
            cachedColValues, _
            proximityStartRow _
        )
        
        ' We want to see if we can continue the current group
        If rowIdx = startRow Then
            currentGroupLastCol = lastCol
        Else
            If lastCol <> currentGroupLastCol Then
                ' Process the group from groupStartRow to rowIdx - 1
                ProcessFillRightGroup ws, startCol, groupStartRow, rowIdx - 1, currentGroupLastCol, rowSourceCells, rowVisibleRanges, undoRange
                ' Start a new group
                groupStartRow = rowIdx
                currentGroupLastCol = lastCol
            End If
        End If
    Next rowIdx
    
    ' Process the final group
    ProcessFillRightGroup ws, startCol, groupStartRow, endRow, currentGroupLastCol, rowSourceCells, rowVisibleRanges, undoRange

    ' --- EXECUTION ---
    If rowSourceCells.Count = 0 Then GoTo CleanExit

    ' Safety Check: Sum up the areas of all rows' visible ranges to prevent fragmentation crashes
    Dim totalAreas As Long
    Dim i As Long
    For i = 1 To rowVisibleRanges.Count
        totalAreas = totalAreas + rowVisibleRanges(i).Areas.Count
    Next i
    
    If totalAreas > Infra_Config.MAX_FILL_DOWN_AREAS Then
        Err.Raise vbObjectError + 1001, "FillRight.Execute", _
            "The target range is too fragmented (exceeds " & Format(Infra_Config.MAX_FILL_DOWN_AREAS, "#,##0") & " non-contiguous areas). Action aborted to prevent Excel from freezing."
    End If

    ' Register Undo for the combined range
    If Not undoRange Is Nothing Then
        If Not Infra_Undo.SaveStateOrConfirm(undoRange, "Fill Right") Then GoTo CleanExit
    End If

    ' Copy/Paste values/formulas/formats for each row (or use native FillRight if no cells are hidden)
    For i = 1 To rowSourceCells.Count
        Set sourceCell = rowSourceCells(i)
        Set visibleRange = rowVisibleRanges(i)
        
        Set fullGroupRange = ws.Range(sourceCell, visibleRange)
        If visibleRange.Cells.CountLarge = fullGroupRange.Cells.CountLarge - sourceCell.Cells.CountLarge Then
            fullGroupRange.FillRight
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
        Infra_Error.HandleErrorDetailed "FillRight.Execute", Err, Nothing, Infra_Error.CATEGORY_VALIDATION, Infra_Error.SEVERITY_WARNING
    Else
        Infra_Error.HandleError "FillRight.Execute", Err
    End If
    Resume CleanExit
End Sub

Private Sub ProcessFillRightGroup( _
    ByVal ws As Worksheet, _
    ByVal startCol As Long, _
    ByVal groupStartRow As Long, _
    ByVal groupEndRow As Long, _
    ByVal groupLastCol As Long, _
    ByVal rowSourceCells As Collection, _
    ByVal rowVisibleRanges As Collection, _
    ByRef undoRange As Range)
    
    Dim tracker As Object: Set tracker = Infra_Error.Track("FillRight.ProcessFillRightGroup")
    On Error GoTo ErrHandler
    
    If groupLastCol <= startCol Then Exit Sub
    
    Dim groupFillRange As Range
    Set groupFillRange = ws.Range(ws.Cells(groupStartRow, startCol + 1), ws.Cells(groupEndRow, groupLastCol))
    
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
            rowSourceCells.Add ws.Range(ws.Cells(groupStartRow, startCol), ws.Cells(groupEndRow, startCol))
            rowVisibleRanges.Add groupFillRange
        Else
            ' There are hidden cells. Fallback to row-by-row processing for this group.
            Dim rIdx As Long
            For rIdx = groupStartRow To groupEndRow
                Dim rowSource As Range
                Set rowSource = ws.Cells(rIdx, startCol)
                Dim rowFill As Range
                Set rowFill = ws.Range(ws.Cells(rIdx, startCol + 1), ws.Cells(rIdx, groupLastCol))
                
                Dim rowVisible As Range
                If rowFill.CountLarge = 1 Then
                    If Not rowFill.EntireRow.Hidden And Not rowFill.EntireColumn.Hidden Then
                        Set rowVisible = rowFill
                    End If
                Else
                    On Error Resume Next
                    Set rowVisible = rowFill.SpecialCells(xlCellTypeVisible)
                    On Error GoTo ErrHandler
                End If
                
                If Not rowVisible Is Nothing Then
                    rowSourceCells.Add rowSource
                    rowVisibleRanges.Add rowVisible
                End If
            Next rIdx
        End If
        
        ' Accumulate the full modified area for a single consolidated undo action
        Dim fullGroupRange As Range
        Set fullGroupRange = ws.Range(ws.Cells(groupStartRow, startCol), ws.Cells(groupEndRow, groupLastCol))
        If undoRange Is Nothing Then
            Set undoRange = fullGroupRange
        Else
            Set undoRange = Union(undoRange, fullGroupRange)
        End If
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "FillRight.ProcessFillRightGroup", Err
    Resume CleanExit
End Sub

