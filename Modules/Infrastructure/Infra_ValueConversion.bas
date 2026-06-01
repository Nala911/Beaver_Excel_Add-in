Attribute VB_Name = "Infra_ValueConversion"
Option Explicit

' @Module: Infra_ValueConversion
' @Category: Infrastructure
' @Description: Shared helpers for converting selections, formulas, and spill ranges to static values.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

Public Function ConvertRangeToValues(ByVal targetRange As Range) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConvertRangeToValues")
    On Error GoTo ErrHandler

    If targetRange Is Nothing Then GoTo CleanExit

    ' Check if the range contains any Dynamic Array formulas with spills
    Dim formulaCells As Range
    Dim hasSpills As Boolean
    Dim cell As Range

    On Error Resume Next
    Set formulaCells = targetRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler

    If Not formulaCells Is Nothing Then
        For Each cell In formulaCells.Cells
            Dim hasSpillVal As Boolean
            On Error Resume Next
            hasSpillVal = cell.HasSpill
            On Error GoTo ErrHandler
            If hasSpillVal Then
                hasSpills = True
                Exit For
            End If
        Next cell
    End If

    If Not hasSpills Then
        For Each cell In targetRange.Cells
            Dim spillAnchor As Range
            On Error Resume Next
            Set spillAnchor = cell.SpillParent
            On Error GoTo ErrHandler
            If Not spillAnchor Is Nothing Then
                hasSpills = True
                Exit For
            End If
        Next cell
    End If

    If Not hasSpills Then
        ' Simple case: no dynamic arrays or spills involved
        targetRange.Value = targetRange.Value
        ConvertRangeToValues = targetRange.Cells.CountLarge
    Else
        ' Spill-aware case: convert formulas to static first, then flatten the range
        Dim processedCount As Long
        For Each cell In targetRange.Cells
            If cell.HasFormula Then
                processedCount = processedCount + ConvertCellToStatic(cell)
            End If
        Next cell
        
        targetRange.Value = targetRange.Value
        ConvertRangeToValues = processedCount
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ConvertRangeToValues", Err
    Resume CleanExit
End Function

Public Function ConvertWorksheetFormulasToValues(ByVal ws As Worksheet) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConvertWorksheetFormulasToValues")
    On Error GoTo ErrHandler

    Dim formulaCells As Range
    Dim area As Range
    Dim cell As Range
    Dim hasSpillsInArea As Boolean

    If ws Is Nothing Then GoTo CleanExit
    If ws.UsedRange.Cells.Count <= 1 And IsEmpty(ws.Range("A1")) Then GoTo CleanExit

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo ErrHandler

    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler

    If formulaCells Is Nothing Then GoTo CleanExit

    For Each area In formulaCells.Areas
        hasSpillsInArea = AreaHasSpill(area)

        If Not hasSpillsInArea Then
            area.Value = area.Value
            ConvertWorksheetFormulasToValues = ConvertWorksheetFormulasToValues + area.Cells.Count
        Else
            For Each cell In area.Cells
                If cell.HasFormula Then
                    ConvertWorksheetFormulasToValues = ConvertWorksheetFormulasToValues + ConvertCellToStatic(cell)
                End If
            Next cell
        End If
    Next area

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ConvertWorksheetFormulasToValues", Err
    Resume CleanExit
End Function

Public Function ConvertCellToStatic(ByVal cell As Range) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConvertCellToStatic")
    On Error GoTo ErrHandler

    Dim isSpill As Boolean

    If cell Is Nothing Then GoTo CleanExit

    On Error Resume Next
    isSpill = cell.HasSpill
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        cell.Value = cell.Value
        ConvertCellToStatic = 1
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler

    If isSpill Then
        cell.SpillingToRange.Value = cell.SpillingToRange.Value
        ConvertCellToStatic = cell.SpillingToRange.Cells.Count
    Else
        cell.Value = cell.Value
        ConvertCellToStatic = 1
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ConvertCellToStatic", Err
    Resume CleanExit
End Function

Private Function AreaHasSpill(ByVal area As Range) As Boolean
    On Error Resume Next
    If area Is Nothing Then Exit Function
    If IsNull(area.HasSpill) Then
        AreaHasSpill = True
    Else
        AreaHasSpill = area.HasSpill
    End If
    On Error GoTo 0
End Function

Public Sub WaitForCalculation()
    Dim tracker As Object: Set tracker = Infra_Error.Track("WaitForCalculation")
    On Error GoTo ErrHandler

    Dim i As Long
    For i = 1 To 1000
        If Application.CalculationState = xlDone Then Exit Sub
        DoEvents
    Next i

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "WaitForCalculation", Err
    Resume CleanExit
End Sub

Public Function ResolveSpillExpandedRange(ByVal sourceRange As Range) As Range
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResolveSpillExpandedRange")
    On Error GoTo ErrHandler

    If sourceRange Is Nothing Then GoTo CleanExit

    ' Wait for background calculations to finish
    WaitForCalculation

    Dim expanded As Range
    Set expanded = sourceRange

    Dim ws As Worksheet
    Set ws = sourceRange.Worksheet
    
    ' Safety check to prevent freezing on extremely large selections
    If sourceRange.Cells.CountLarge > 50000 Then
        Set ResolveSpillExpandedRange = sourceRange
        GoTo CleanExit
    End If

    ' Find all dynamic array spill ranges on the worksheet and check intersection
    Dim formulaCells As Range
    Dim fCell As Range

    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler

    If Not formulaCells Is Nothing Then
        For Each fCell In formulaCells.Cells
            Dim hasSpillVal As Boolean
            hasSpillVal = False
            
            On Error Resume Next
            hasSpillVal = fCell.HasSpill
            On Error GoTo ErrHandler
            
            If hasSpillVal Then
                Dim spillRange As Range
                Set spillRange = fCell.SpillingToRange
                If Not Application.Intersect(sourceRange, spillRange) Is Nothing Then
                    Set expanded = Application.Union(expanded, spillRange)
                End If
            End If
        Next fCell
    End If

    Set ResolveSpillExpandedRange = expanded

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ResolveSpillExpandedRange", Err
    Set ResolveSpillExpandedRange = sourceRange
    Resume CleanExit
End Function
