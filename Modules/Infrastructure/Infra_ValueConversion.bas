Attribute VB_Name = "Infra_ValueConversion"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' @Module: Infra_ValueConversion
' @Category: Infrastructure
' @Description: Shared helpers for converting selections, formulas, and spill ranges to static values.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Lib_ValueConversionFunction

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
        ' Spill-aware case: first try expanded bulk value flattening
        Dim expanded As Range
        Set expanded = ResolveSpillExpandedRange(targetRange)
        Dim bulkSucceeded As Boolean
        bulkSucceeded = False
        
        If Not expanded Is Nothing Then
            On Error Resume Next
            expanded.Value = expanded.Value
            If Err.Number = 0 Then
                bulkSucceeded = True
                If Not formulaCells Is Nothing Then
                    processedCount = formulaCells.Cells.CountLarge
                End If
            End If
            Err.Clear
            On Error GoTo ErrHandler
        End If
        
        If Not bulkSucceeded Then
            ' Area-by-area fallback: process only areas with spills cell-by-cell
            If Not formulaCells Is Nothing Then
                Dim area As Range
                For Each area In formulaCells.Areas
                    Dim hasSpillVal As Variant
                    hasSpillVal = Null
                    On Error Resume Next
                    hasSpillVal = area.HasSpill
                    On Error GoTo ErrHandler
                    
                    If hasSpillVal = True Then
                        ' All are spills: try bulk conversion of the expanded area
                        Dim areaExpanded As Range
                        Set areaExpanded = ResolveSpillExpandedRange(area)
                        If Not areaExpanded Is Nothing Then
                            areaExpanded.Value = areaExpanded.Value
                            processedCount = processedCount + area.Cells.CountLarge
                        End If
                    ElseIf IsNull(hasSpillVal) Then
                        ' Spills and non-spills mixed: loop cell-by-cell in this area
                        For Each cell In area.Cells
                            processedCount = processedCount + ConvertCellToStatic(cell)
                        Next cell
                    Else
                        ' No spills in this area: bulk convert
                        area.Value = area.Value
                        processedCount = processedCount + area.Cells.CountLarge
                    End If
                Next area
            End If
        End If
        
        On Error Resume Next
        targetRange.Value = targetRange.Value
        On Error GoTo ErrHandler
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
            Set areaFormulaCells = area

            If Not areaFormulaCells Is Nothing Then
                Dim expanded As Range
                Set expanded = ResolveSpillExpandedRange(area)
                
                Dim bulkSucceeded As Boolean
                bulkSucceeded = False
                
                If Not expanded Is Nothing Then
                    On Error Resume Next
                    expanded.Value = expanded.Value
                    If Err.Number = 0 Then
                        bulkSucceeded = True
                        ConvertWorksheetFormulasToValues = ConvertWorksheetFormulasToValues + areaFormulaCells.Cells.CountLarge
                    End If
                    Err.Clear
                    On Error GoTo ErrHandler
                End If
                
                If Not bulkSucceeded Then
                    For Each cell In areaFormulaCells.Cells
                        ConvertWorksheetFormulasToValues = ConvertWorksheetFormulasToValues + ConvertCellToStatic(cell)
                    Next cell
                End If
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
    For i = 1 To 100
        If Application.CalculationState = xlDone Then Exit Sub
        DoEvents
        Sleep 10
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
        ' Check if the whole selection has no array formulas in a single COM call
        Dim selectionHasArray As Variant
        selectionHasArray = Null
        On Error Resume Next
        selectionHasArray = formulaCellsSelection.HasArray
        On Error GoTo ErrHandler
        
        If IsNull(selectionHasArray) Or (selectionHasArray = True) Then
            If selectionHasArray = True Then
                ' The entire selection is one array
                Dim selectionArrRange As Range
                On Error Resume Next
                Set selectionArrRange = formulaCellsSelection.CurrentArray
                On Error GoTo ErrHandler
                If Not selectionArrRange Is Nothing Then
                    Set expanded = Application.Union(expanded, selectionArrRange)
                End If
            Else
                ' Mixed: scan cell-by-cell
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
        End If
    End If

    ' 2. Expand for dynamic array spill ranges intersecting the selection
    Dim scanRange As Range
    On Error Resume Next
    Set scanRange = Application.Intersect(sourceRange, ws.UsedRange)
    On Error GoTo ErrHandler
    
    Dim formulaCount As Long
    Dim sheetFormulas As Range
    On Error Resume Next
    Set sheetFormulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    If Not sheetFormulas Is Nothing Then
        formulaCount = sheetFormulas.Cells.CountLarge
    End If
    On Error GoTo ErrHandler

    If Not scanRange Is Nothing Then
        If AreaHasSpill(scanRange) Then
            Dim processedSpills As Range
            
            If formulaCount > 0 And formulaCount < scanRange.Cells.CountLarge Then
                ' Optimization: Scan only the formula cells on the sheet
                Dim formulaArea As Range
                For Each formulaArea In sheetFormulas.Areas
                    Dim areaSpill As Variant
                    areaSpill = False
                    On Error Resume Next
                    areaSpill = formulaArea.HasSpill
                    On Error GoTo ErrHandler
                    
                    ' Only scan this area if it contains at least one spill parent/cell
                    If IsNull(areaSpill) Or (areaSpill = True) Then
                        Dim fCell As Range
                        For Each fCell In formulaArea.Cells
                            Dim isSpill As Boolean
                            isSpill = False
                            On Error Resume Next
                            isSpill = fCell.HasSpill
                            On Error GoTo ErrHandler
                            
                            If isSpill Then
                                Dim spRange As Range
                                On Error Resume Next
                                Set spRange = fCell.SpillingToRange
                                On Error GoTo ErrHandler
                                
                                If Not spRange Is Nothing Then
                                    ' Check if this spill range intersects our scanRange
                                    If Not Application.Intersect(scanRange, spRange) Is Nothing Then
                                        Set expanded = Application.Union(expanded, spRange)
                                    End If
                                End If
                            End If
                        Next fCell
                    End If
                Next formulaArea
            Else
                ' Default path: Scan each cell in scanRange
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
    Ensure2DArray = Lib_ValueConversionFunction.Ensure2DArray(InputVal)
End Function

' Attempts to convert a string or value to a numeric double, handling trailing minus, percent, and currency signs.
Public Function TryConvertToNumber(ByVal rawVal As Variant, ByRef outVal As Variant) As Boolean
    TryConvertToNumber = Lib_ValueConversionFunction.TryConvertToNumber(rawVal, outVal)
End Function

' Centralized helper to map Excel CVErr values to their string representations.
Public Function GetExcelErrorText(ByVal errVal As Variant) As String
    GetExcelErrorText = Lib_ValueConversionFunction.GetExcelErrorText(errVal)
End Function

' Centralized helper to determine local system date pattern (dd-mm-yyyy, mm-dd-yyyy, or yyyy-mm-dd).
Public Function GetSystemDateFormatPattern() As String
    Static cachedPattern As String
    If cachedPattern <> "" Then
        GetSystemDateFormatPattern = cachedPattern
        Exit Function
    End If

    Dim dateOrder As Long
    On Error Resume Next
    dateOrder = Application.International(xlDateOrder)
    If Err.Number <> 0 Then
        dateOrder = 1 ' Default to DMY (1) if not in Excel environment
        Err.Clear
    End If
    On Error GoTo 0
    
    Select Case dateOrder
        Case 0: cachedPattern = "mm-dd-yyyy"
        Case 1: cachedPattern = "dd-mm-yyyy"
        Case 2: cachedPattern = "yyyy-mm-dd"
        Case Else: cachedPattern = "dd-mm-yyyy"
    End Select
    GetSystemDateFormatPattern = cachedPattern
End Function

' High-performance native VBA string space collapse (removes leading/trailing spaces and collapses internal multi-spaces without COM calls)
Public Function TrimSpacesVBA(ByRef text As String) As String
    If text = vbNullString Then Exit Function
    Dim s As String
    s = Trim$(text)
    Do While InStr(1, s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    TrimSpacesVBA = s
End Function

' High-performance native VBA non-printable character stripper (ASCII 0..31 without COM calls)
Public Function CleanNonPrintablesVBA(ByRef text As String) As String
    If text = vbNullString Then Exit Function
    Dim bytes() As Byte
    bytes = text
    Dim i As Long, code As Integer
    Dim hasControl As Boolean
    For i = 0 To UBound(bytes) Step 2
        code = bytes(i) + (bytes(i + 1) * 256)
        If code >= 0 And code <= 31 Then
            hasControl = True
            Exit For
        End If
    Next i
    If Not hasControl Then
        CleanNonPrintablesVBA = text
        Exit Function
    End If
    Dim sb As String
    sb = text
    For i = 0 To 31
        If InStr(1, sb, Chr$(i)) > 0 Then
            sb = Replace(sb, Chr$(i), vbNullString)
        End If
    Next i
    CleanNonPrintablesVBA = sb
End Function



