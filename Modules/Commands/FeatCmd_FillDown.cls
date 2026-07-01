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
    Dim maxSearchCol As Long
    
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
    maxSearchCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    ' Collect target fill ranges and source cells to update
    Dim colSourceCells As New Collection
    Dim colVisibleRanges As New Collection
    
    Dim undoRange As Range

    For colIdx = startCol To endCol
        Dim sourceCell As Range
        Set sourceCell = ws.Cells(startRow, colIdx)
        
        Dim lastRow As Long
        lastRow = 0
        
        Dim foundRef As Boolean
        foundRef = False
        
        ' --- PRIORITY 1: EXCEL TABLE (ListObject) ---
        Dim lo As ListObject
        On Error Resume Next
        Set lo = sourceCell.ListObject
        On Error GoTo ErrHandler
        
        If Not lo Is Nothing Then
            If Not Intersect(sourceCell, lo.DataBodyRange) Is Nothing Then
                lastRow = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
                foundRef = True
            End If
        End If
        
        ' --- PRIORITY 2: CURRENT REGION ---
        If Not foundRef Then
            Dim rRegion As Range
            Set rRegion = sourceCell.CurrentRegion
            
            If rRegion.Cells(rRegion.Cells.CountLarge).Address <> sourceCell.Address Then
                lastRow = rRegion.Row + rRegion.Rows.Count - 1
                foundRef = True
            End If
        End If
        
        ' --- PRIORITY 3: PROXIMITY SEARCH (closest neighbor) ---
        ' Optimization: Limit proximity distance to 10 columns
        If Not foundRef Then
            Dim maxProximityDist As Long
            maxProximityDist = Infra_Config.MAX_FILL_PROXIMITY_COLUMNS
            Dim dist As Long
            Dim leftCol As Long, rightCol As Long
            Dim targetCol As Long
            
            For dist = 1 To maxProximityDist
                leftCol = colIdx - dist
                If leftCol >= 1 Then
                    If Not IsEmpty(ws.Cells(startRow, leftCol)) Then
                        targetCol = leftCol
                        foundRef = True
                        Exit For
                    End If
                End If
                
                rightCol = colIdx + dist
                If rightCol <= maxSearchCol Then
                    If Not IsEmpty(ws.Cells(startRow, rightCol)) Then
                        targetCol = rightCol
                        foundRef = True
                        Exit For
                    End If
                End If
            Next dist
            
            If foundRef Then
                Dim nearRegion As Range
                Set nearRegion = ws.Cells(startRow, targetCol).CurrentRegion
                lastRow = nearRegion.Row + nearRegion.Rows.Count - 1
            End If
        End If
        
        ' If a valid reference was found and lastRow goes past startRow
        If foundRef And lastRow > startRow Then
            Dim fillRange As Range
            Set fillRange = ws.Range(ws.Cells(startRow + 1, colIdx), ws.Cells(lastRow, colIdx))
            
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
                colSourceCells.Add sourceCell
                colVisibleRanges.Add visibleRange
                
                ' Accumulate the full modified area for a single consolidated undo action
                Dim fullColRange As Range
                Set fullColRange = ws.Range(sourceCell, ws.Cells(lastRow, colIdx))
                If undoRange Is Nothing Then
                    Set undoRange = fullColRange
                Else
                    Set undoRange = Union(undoRange, fullColRange)
                End If
            End If
        End If
    Next colIdx

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

    ' Copy/Paste values for each column
    For i = 1 To colSourceCells.Count
        Set sourceCell = colSourceCells(i)
        Set visibleRange = colVisibleRanges(i)
        
        sourceCell.Copy
        visibleRange.PasteSpecial Paste:=xlPasteAll
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
