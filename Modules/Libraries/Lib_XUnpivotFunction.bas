Attribute VB_Name = "Lib_XUnpivotFunction"
Option Explicit

' @Module: Lib_XUnpivotFunction
' @Category: Library
' @Description: UDF to unpivot wide data into a long, normal form array.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

' Transforms a wide range of data into a long format (unpivot).
'
' ARGUMENTS:
'   SourceRange       : The source range to unpivot. The first row must be headers.
'   FixedColumnsCount : Number of columns on the left to repeat for each unpivoted value.
'   IgnoreBlanks      : (Optional) If True, empty cells in the unpivot area are skipped. Defaults to True.
'   AttributeHeader   : (Optional) The header name for the unpivoted column names. Defaults to "Attribute".
'   ValueHeader       : (Optional) The header name for the unpivoted values. Defaults to "Value".
'
' RETURNS: A dynamic 2D array that spills into the sheet.
' ==============================================================================
Public Function XUnpivot(SourceRange As Range, FixedColumnsCount As Long, Optional IgnoreBlanks As Variant, Optional AttributeHeader As Variant, Optional ValueHeader As Variant) As Variant
    Infra_Error.PushContext "XUnpivot"
    On Error GoTo ErrHandler
    
    Dim arrSource As Variant
    Dim resultArr() As Variant
    Dim maxRows As Long
    Dim outCols As Long
    Dim outRow As Long
    Dim r As Long, c As Long, i As Long
    Dim cellValue As Variant
    
    Dim blnIgnoreBlanks As Boolean
    Dim strAttrHeader As String
    Dim strValHeader As String
    
    ' 1. Validate inputs
    If FixedColumnsCount < 1 Or FixedColumnsCount >= SourceRange.Columns.Count Then
        XUnpivot = CVErr(xlErrValue)
        GoTo CleanExit
    End If
    
    If SourceRange.Rows.Count < 2 Then
        XUnpivot = CVErr(xlErrValue)
        GoTo CleanExit
    End If
    
    ' 2. Robust argument handling
    ' Handle IgnoreBlanks
    If IsMissing(IgnoreBlanks) Then
        blnIgnoreBlanks = True
    ElseIf IsObject(IgnoreBlanks) Then
        blnIgnoreBlanks = CBool(IgnoreBlanks.Cells(1, 1).Value)
    Else
        On Error Resume Next
        blnIgnoreBlanks = CBool(IgnoreBlanks)
        If Err.Number <> 0 Then blnIgnoreBlanks = True: Err.Clear
        On Error GoTo ErrHandler
    End If
    
    ' Handle AttributeHeader
    If IsMissing(AttributeHeader) Then
        strAttrHeader = "Attribute"
    ElseIf IsObject(AttributeHeader) Then
        strAttrHeader = CStr(AttributeHeader.Cells(1, 1).Value)
    Else
        strAttrHeader = CStr(AttributeHeader)
    End If
    
    ' Handle ValueHeader
    If IsMissing(ValueHeader) Then
        strValHeader = "Value"
    ElseIf IsObject(ValueHeader) Then
        strValHeader = CStr(ValueHeader.Cells(1, 1).Value)
    Else
        strValHeader = CStr(ValueHeader)
    End If
    
    ' 3. Read into memory array for speed
    If SourceRange.Cells.Count = 1 Then
        ' Should not happen based on rows/cols validation, but for safety
        ReDim arrSource(1 To 1, 1 To 1)
        arrSource(1, 1) = SourceRange.Value2
    Else
        arrSource = SourceRange.Value2
    End If
    
    ' 4. Calculate Dimensions
    outCols = FixedColumnsCount + 2
    maxRows = (UBound(arrSource, 1) - 1) * (UBound(arrSource, 2) - FixedColumnsCount) + 1
    
    ' 5. Initialize result array
    ReDim resultArr(1 To maxRows, 1 To outCols)
    
    ' 6. Build Header Row
    For i = 1 To FixedColumnsCount
        resultArr(1, i) = arrSource(1, i)
    Next i
    resultArr(1, FixedColumnsCount + 1) = strAttrHeader
    resultArr(1, outCols) = strValHeader
    outRow = 1
    
    ' 7. Loop through data and build unpivoted array
    For r = 2 To UBound(arrSource, 1)
        For c = FixedColumnsCount + 1 To UBound(arrSource, 2)
            cellValue = arrSource(r, c)
            
            ' Check if we should include this value
            If Not (blnIgnoreBlanks And IsEmpty(cellValue)) And Not (blnIgnoreBlanks And cellValue = "") Then
                outRow = outRow + 1
                
                ' Copy fixed columns
                For i = 1 To FixedColumnsCount
                    resultArr(outRow, i) = arrSource(r, i)
                Next i
                
                ' Add Attribute (from header)
                resultArr(outRow, FixedColumnsCount + 1) = arrSource(1, c)
                
                ' Add Value
                resultArr(outRow, outCols) = cellValue
            End If
        Next c
    Next r
    
    ' 7. Resize and Return
    If outRow = 1 Then
        ' Only headers, no data
        XUnpivot = CVErr(xlErrNull)
    Else
        Dim finalResult() As Variant
        ReDim finalResult(1 To outRow, 1 To outCols)
        For r = 1 To outRow
            For c = 1 To outCols
                finalResult(r, c) = resultArr(r, c)
            Next c
        Next r
        XUnpivot = finalResult
    End If

CleanExit:
    Infra_Error.PopContext
    Exit Function

ErrHandler:
    Infra_Error.HandleError "XUnpivot", Err
    XUnpivot = CVErr(xlErrValue)
End Function
