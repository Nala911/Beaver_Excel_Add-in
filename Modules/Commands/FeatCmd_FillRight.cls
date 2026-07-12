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
    
    Dim ctx As ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit
    Set selectionRng = ctx.SelectionRange
    Set ws = selectionRng.Worksheet

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

    For rowIdx = startRow To endRow
        Dim sourceCell As Range
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
        
        ' If a valid reference was found and lastCol goes past startCol
        If lastCol > startCol Then
            Dim fillRange As Range
            Set fillRange = ws.Range(ws.Cells(rowIdx, startCol + 1), ws.Cells(rowIdx, lastCol))
            
            Dim visibleRange As Range
            If fillRange.CountLarge = 1 Then
                If Not fillRange.EntireRow.Hidden And Not fillRange.EntireColumn.Hidden Then
                    Set visibleRange = fillRange
                End If
            Else
                On Error Resume Next
                Set visibleRange = fillRange.SpecialCells(xlCellTypeVisible)
                On Error GoTo ErrHandler
            End If
            
            If Not visibleRange Is Nothing Then
                rowSourceCells.Add sourceCell
                rowVisibleRanges.Add visibleRange
                
                ' Accumulate the full modified area for a single consolidated undo action
                Dim fullRowRange As Range
                Set fullRowRange = ws.Range(sourceCell, ws.Cells(rowIdx, lastCol))
                If undoRange Is Nothing Then
                    Set undoRange = fullRowRange
                Else
                    Set undoRange = Union(undoRange, fullRowRange)
                End If
            End If
        End If
    Next rowIdx

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

    ' Copy/Paste values/formulas/formats for each row
    For i = 1 To rowSourceCells.Count
        Set sourceCell = rowSourceCells(i)
        Set visibleRange = rowVisibleRanges(i)
        
        sourceCell.Copy
        visibleRange.PasteSpecial Paste:=xlPasteAll
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
