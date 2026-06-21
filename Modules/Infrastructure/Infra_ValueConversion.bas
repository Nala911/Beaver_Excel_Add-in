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
    Dim processedCount As Long

    If targetRange.CountLarge = 1 Then
        If targetRange.HasFormula Then
            Set formulaCells = targetRange
        End If
    Else
        On Error Resume Next
        Set formulaCells = targetRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo ErrHandler
    End If

    hasSpills = AreaHasSpill(targetRange)

    If Not hasSpills Then
        ' Simple case: no dynamic arrays or spills involved
        On Error Resume Next
        targetRange.Value = targetRange.Value
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo ErrHandler
            ' Fallback to cell-by-cell conversion if block conversion fails
            For Each cell In targetRange.Cells
                If cell.HasFormula Then
                    processedCount = processedCount + ConvertCellToStatic(cell)
                End If
            Next cell
            targetRange.Value = targetRange.Value
            ConvertRangeToValues = processedCount
        Else
            On Error GoTo ErrHandler
            ConvertRangeToValues = targetRange.Cells.CountLarge
        End If
    Else
        ' Spill-aware case: convert formulas to static first, then flatten the range
        If Not formulaCells Is Nothing Then
            For Each cell In formulaCells.Cells
                processedCount = processedCount + ConvertCellToStatic(cell)
            Next cell
        End If
        
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
            On Error Resume Next
            area.Value = area.Value
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo ErrHandler
                ' Fallback to cell-by-cell conversion if block conversion fails
                For Each cell In area.Cells
                    If cell.HasFormula Then
                        ConvertWorksheetFormulasToValues = ConvertWorksheetFormulasToValues + ConvertCellToStatic(cell)
                    End If
                Next cell
            Else
                On Error GoTo ErrHandler
                ConvertWorksheetFormulasToValues = ConvertWorksheetFormulasToValues + area.Cells.Count
            End If
        Else
            Dim areaFormulaCells As Range
            On Error Resume Next
            Set areaFormulaCells = area.SpecialCells(xlCellTypeFormulas)
            On Error GoTo ErrHandler

            If Not areaFormulaCells Is Nothing Then
                For Each cell In areaFormulaCells.Cells
                    ConvertWorksheetFormulasToValues = ConvertWorksheetFormulasToValues + ConvertCellToStatic(cell)
                Next cell
            End If
        End If
    Next area

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ConvertWorksheetFormulasToValues", Err
    Resume CleanExit
End Function

Public Function ConvertCellToStatic(ByVal cell As Range) As Long
    On Error GoTo ErrHandler

    Dim isSpill As Boolean
    Dim isArray As Boolean

    If cell Is Nothing Then GoTo CleanExit

    On Error Resume Next
    isSpill = cell.HasSpill
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        
        ' Fallback if HasSpill is not supported (e.g. older Excel)
        ' Check if it is a legacy array formula
        Dim fallbackIsArray As Boolean
        On Error Resume Next
        fallbackIsArray = cell.HasArray
        On Error GoTo 0
        
        If fallbackIsArray Then
            Dim fallbackArrRange As Range
            On Error Resume Next
            Set fallbackArrRange = cell.CurrentArray
            On Error GoTo 0
            If Not fallbackArrRange Is Nothing Then
                fallbackArrRange.Value = fallbackArrRange.Value
                ConvertCellToStatic = fallbackArrRange.Cells.Count
                GoTo CleanExit
            End If
        End If
        
        cell.Value = cell.Value
        ConvertCellToStatic = 1
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler

    If isSpill Then
        Dim spillRange As Range
        On Error Resume Next
        Set spillRange = cell.SpillingToRange
        On Error GoTo ErrHandler
        
        If Not spillRange Is Nothing Then
            spillRange.Value = spillRange.Value
            ConvertCellToStatic = spillRange.Cells.Count
        Else
            cell.Value = cell.Value
            ConvertCellToStatic = 1
        End If
    Else
        ' Check if it is a legacy array formula
        On Error Resume Next
        isArray = cell.HasArray
        On Error GoTo ErrHandler
        
        If isArray Then
            Dim arrRange As Range
            On Error Resume Next
            Set arrRange = cell.CurrentArray
            On Error GoTo ErrHandler
            If Not arrRange Is Nothing Then
                arrRange.Value = arrRange.Value
                ConvertCellToStatic = arrRange.Cells.Count
                GoTo CleanExit
            End If
        End If
        
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

    ' 1. Expand for any intersecting legacy array formulas (CSE arrays)
    Dim formulaCellsSelection As Range
    Dim cell As Range
    On Error Resume Next
    Set formulaCellsSelection = Application.Intersect(sourceRange, ws.UsedRange.SpecialCells(xlCellTypeFormulas))
    On Error GoTo ErrHandler
    
    If Not formulaCellsSelection Is Nothing Then
        For Each cell In formulaCellsSelection.Cells
            Dim isArray As Boolean
            On Error Resume Next
            isArray = cell.HasArray
            On Error GoTo ErrHandler
            
            If isArray Then
                Dim arrRange As Range
                On Error Resume Next
                Set arrRange = cell.CurrentArray
                On Error GoTo ErrHandler
                If Not arrRange Is Nothing Then
                    Set expanded = Application.Union(expanded, arrRange)
                End If
            End If
        Next cell
    End If

    ' 2. Expand for dynamic array spill ranges intersecting the selection
    Dim scanRange As Range
    On Error Resume Next
    Set scanRange = Application.Intersect(sourceRange, ws.UsedRange)
    On Error GoTo ErrHandler

    If Not scanRange Is Nothing Then
        If AreaHasSpill(scanRange) Then
            Dim processedSpills As Range
            For Each cell In scanRange.Cells
                Dim shouldCheck As Boolean
                shouldCheck = True
                
                If Not processedSpills Is Nothing Then
                    If Not Application.Intersect(cell, processedSpills) Is Nothing Then
                        shouldCheck = False
                    End If
                End If
                
                If shouldCheck Then
                    Dim hasSpillVal As Boolean
                    hasSpillVal = False
                    
                    On Error Resume Next
                    hasSpillVal = cell.HasSpill
                    On Error GoTo ErrHandler
                    
                    If hasSpillVal Then
                        Dim spillParentCell As Range
                        On Error Resume Next
                        Set spillParentCell = cell.SpillParent
                        On Error GoTo ErrHandler
                        
                        If spillParentCell Is Nothing Then
                            Set spillParentCell = cell
                        End If
                        
                        Dim spillRange As Range
                        On Error Resume Next
                        Set spillRange = spillParentCell.SpillingToRange
                        On Error GoTo ErrHandler
                        
                        If Not spillRange Is Nothing Then
                            Set expanded = Application.Union(expanded, spillRange)
                            If processedSpills Is Nothing Then
                                Set processedSpills = spillRange
                            Else
                                Set processedSpills = Application.Union(processedSpills, spillRange)
                            End If
                        End If
                    End If
                End If
            Next cell
        End If
    End If

    Set ResolveSpillExpandedRange = expanded

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ResolveSpillExpandedRange", Err
    Set ResolveSpillExpandedRange = sourceRange
    Resume CleanExit
End Function

' Helper function to ensure any variant input (Range, Array, or Scalar) is converted to a 1-based 2D array.
Public Function Ensure2DArray(ByVal InputVal As Variant) As Variant
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ensure2DArray")
    On Error GoTo ErrHandler

    Dim result() As Variant

    If IsObject(InputVal) Then
        If InputVal Is Nothing Then
            ReDim result(1 To 1, 1 To 1)
            result(1, 1) = Empty
            Ensure2DArray = result
            GoTo CleanExit
        End If
        If TypeOf InputVal Is Range Then
            Dim r As Range: Set r = InputVal
            If r.Cells.Count = 1 Then
                ReDim result(1 To 1, 1 To 1)
                result(1, 1) = r.Value2
                Ensure2DArray = result
            Else
                Ensure2DArray = r.Value2
            End If
            GoTo CleanExit
        End If
    End If

    If IsArray(InputVal) Then
        Dim dims As Long
        dims = GetArrayDims(InputVal)
        If dims = 1 Then
            Dim i As Long, lb As Long, ub As Long
            lb = LBound(InputVal)
            ub = UBound(InputVal)
            ReDim result(1 To (ub - lb + 1), 1 To 1)
            For i = lb To ub
                result(i - lb + 1, 1) = InputVal(i)
            Next i
            Ensure2DArray = result
        ElseIf dims = 2 Then
            ' Check if it is a 1-based 2D array. If not, normalize it to 1-based.
            Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
            lb1 = LBound(InputVal, 1)
            ub1 = UBound(InputVal, 1)
            lb2 = LBound(InputVal, 2)
            ub2 = UBound(InputVal, 2)
            
            If lb1 = 1 And lb2 = 1 Then
                Ensure2DArray = InputVal
            Else
                Dim rIdx As Long, cIdx As Long
                ReDim result(1 To (ub1 - lb1 + 1), 1 To (ub2 - lb2 + 1))
                For rIdx = lb1 To ub1
                    For cIdx = lb2 To ub2
                        result(rIdx - lb1 + 1, cIdx - lb2 + 1) = InputVal(rIdx, cIdx)
                    Next cIdx
                Next rIdx
                Ensure2DArray = result
            End If
        Else
            ' Fallback for higher dimensions: use first cell
            ReDim result(1 To 1, 1 To 1)
            result(1, 1) = InputVal
            Ensure2DArray = result
        End If
    Else
        ' Scalar value
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = InputVal
        Ensure2DArray = result
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Ensure2DArray", Err
    Resume CleanExit
End Function

Private Function GetArrayDims(ByVal arr As Variant) As Long
    On Error Resume Next
    Dim i As Long, dummy As Long
    For i = 1 To 60000
        dummy = LBound(arr, i)
        If Err.Number <> 0 Then
            GetArrayDims = i - 1
            Err.Clear
            Exit Function
        End If
    Next i
    GetArrayDims = 0
End Function

' Attempts to convert a string or value to a numeric double, handling trailing minus, percent, and currency signs.
' Avoids octal or hexadecimal representations like &H.
Public Function TryConvertToNumber(ByVal rawVal As Variant, ByRef outVal As Variant) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("TryConvertToNumber")
    On Error GoTo ErrHandler

    TryConvertToNumber = False
    If IsError(rawVal) Then GoTo CleanExit
    
    Dim txt As String
    txt = CStr(rawVal)
    txt = Trim$(txt)
    
    If txt = vbNullString Then GoTo CleanExit
    
    ' Exclude hexadecimal/octal syntax which IsNumeric allows but are not standard numeric strings
    If Left$(txt, 2) = "&H" Or Left$(txt, 2) = "&h" Or Left$(txt, 2) = "&O" Or Left$(txt, 2) = "&o" Then
        GoTo CleanExit
    End If
    
    ' 1. Check if it's already directly numeric
    If IsNumeric(txt) Then
        On Error Resume Next
        outVal = CDbl(txt)
        If Err.Number = 0 Then
            TryConvertToNumber = True
            GoTo CleanExit
        End If
        Err.Clear
        On Error GoTo ErrHandler
    End If
    
    ' 2. Handle trailing minus sign (e.g. "123.45-" -> "-123.45")
    If Right$(txt, 1) = "-" Then
        txt = "-" & Left$(txt, Len(txt) - 1)
        txt = Trim$(txt)
    End If
    
    ' 3. Handle percent sign (e.g. "45%" -> 0.45)
    Dim isPercent As Boolean
    isPercent = False
    If Right$(txt, 1) = "%" Then
        isPercent = True
        txt = Left$(txt, Len(txt) - 1)
        txt = Trim$(txt)
    End If
    
    ' 4. Remove currency symbols and formatting characters
    txt = Replace(txt, "$", "")
    txt = Replace(txt, "€", "")
    txt = Replace(txt, "£", "")
    txt = Replace(txt, "¥", "")
    txt = Replace(txt, " ", "")
    
    ' Try conversion again after stripping formatting
    If IsNumeric(txt) Then
        On Error Resume Next
        Dim dVal As Double
        dVal = CDbl(txt)
        If Err.Number = 0 Then
            If isPercent Then
                dVal = dVal / 100
            End If
            outVal = dVal
            TryConvertToNumber = True
            GoTo CleanExit
        End If
        Err.Clear
        On Error GoTo ErrHandler
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "TryConvertToNumber", Err
    Resume CleanExit
End Function

