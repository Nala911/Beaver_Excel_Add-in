Attribute VB_Name = "Udf_XUnpivot"
Option Explicit

' @Module: Udf_XUnpivot
' @Category: Library
' @Description: UDF to unpivot wide data into long format, auto-detecting columns by trailing numeric values.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_ValueConversion

' Unpivots a 2D range of wide data into a 2D long data format.
' The first row is assumed to be headers.
' The columns to unpivot are identified by checking the second row (the first data row)
' from right to left for contiguous numerical values.
'
' ARGUMENTS:
'   data_range      : The source range or 2D array to unpivot (including header row).
'   attribute_header: Optional header name for the unpivoted attribute column (default is "Attribute").
'   value_header    : Optional header name for the unpivoted value column (default is "Value").
'   skip_blanks     : Optional boolean. If True, rows with empty or blank values in the unpivot columns are omitted (default is False).
'
' RETURNS: A dynamic array that spills into the sheet.
' ==============================================================================
Public Function XUnpivot(ByVal data_range As Variant, _
                         Optional ByVal attribute_header As String = "Attribute", _
                         Optional ByVal value_header As String = "Value", _
                         Optional ByVal skip_blanks As Boolean = False) As Variant
    Dim tracker As Object: Set tracker = Infra_Error.Track("XUnpivot")
    On Error GoTo ErrHandler
    
    ' --- Optimization: Read ranges into memory arrays ---
    Dim arr As Variant
    arr = Infra_ValueConversion.Ensure2DArray(data_range)
    
    Dim numRows As Long, numCols As Long
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    ' Check boundaries: must have at least header + 1 data row and at least 2 columns
    If numRows < 2 Or numCols < 2 Then
        XUnpivot = CVErr(xlErrValue)
        GoTo CleanExit
    End If
    
    ' Identify columns to unpivot by scanning the second row from right to left for numbers.
    Dim numUnpivot As Long: numUnpivot = 0
    Dim c As Long
    For c = numCols To 1 Step -1
        Dim val As Variant: val = arr(2, c)
        
        If IsError(val) Then
            Exit For
        End If
        If IsEmpty(val) Then
            Exit For
        End If
        If IsNumeric(val) Then
            numUnpivot = numUnpivot + 1
        Else
            Exit For
        End If
    Next c
    
    ' If no numeric columns found at the end of the second row, return a #NUM! error
    If numUnpivot = 0 Then
        XUnpivot = CVErr(xlErrNum)
        GoTo CleanExit
    End If
    
    Dim numKeys As Long
    numKeys = numCols - numUnpivot
    
    ' Estimate maximum output rows (excluding the header row, it is: (numRows - 1) * numUnpivot)
    ' Plus 1 for header row.
    Dim maxOutputRows As Long
    maxOutputRows = 1 + (numRows - 1) * numUnpivot
    
    Dim outCols As Long
    outCols = numKeys + 2
    
    Dim tempResult() As Variant
    ReDim tempResult(1 To maxOutputRows, 1 To outCols)
    
    ' 1. Write headers
    Dim kc As Long
    For kc = 1 To numKeys
        tempResult(1, kc) = arr(1, kc)
    Next kc
    tempResult(1, numKeys + 1) = attribute_header
    tempResult(1, numKeys + 2) = value_header
    
    ' 2. Process data rows
    Dim r As Long, uc As Long
    Dim outRow As Long: outRow = 1
    For r = 2 To numRows
        For uc = (numKeys + 1) To numCols
            Dim cellVal As Variant: cellVal = arr(r, uc)
            Dim shouldSkip As Boolean: shouldSkip = False
            
            If skip_blanks Then
                If IsEmpty(cellVal) Then
                    shouldSkip = True
                ElseIf VarType(cellVal) = vbString Then
                    If cellVal = "" Then shouldSkip = True
                End If
            End If
            
            If Not shouldSkip Then
                outRow = outRow + 1
                ' Copy key columns
                For kc = 1 To numKeys
                    tempResult(outRow, kc) = arr(r, kc)
                Next kc
                ' Write attribute header
                tempResult(outRow, numKeys + 1) = arr(1, uc)
                ' Write value
                tempResult(outRow, numKeys + 2) = cellVal
            End If
        Next uc
    Next r
    
    ' Populate final results array of exact size
    Dim finalResult() As Variant
    ReDim finalResult(1 To outRow, 1 To outCols)
    Dim i As Long, j As Long
    For i = 1 To outRow
        For j = 1 To outCols
            finalResult(i, j) = tempResult(i, j)
        Next j
    Next i
    
    XUnpivot = finalResult
    
CleanExit:
    Exit Function
    
ErrHandler:
    Infra_Error.HandleError "XUnpivot", Err
    XUnpivot = CVErr(xlErrValue) ' #VALUE! on general error
End Function
