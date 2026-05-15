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

    targetRange.Value = targetRange.Value
    ConvertRangeToValues = targetRange.Cells.CountLarge

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
