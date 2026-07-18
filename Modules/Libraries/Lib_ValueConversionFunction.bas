Attribute VB_Name = "Lib_ValueConversionFunction"
Option Explicit

' @Module: Lib_ValueConversionFunction
' @Category: Library
' @Description: Pure helper functions for variant types, cell value parsing, and type conversions.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

' Helper function to ensure any variant input (Range, Array, or Scalar) is converted to a 1-based 2D array.
' Excludes Infra_Error tracking to optimize grid formula CPU execution speed.
Public Function Ensure2DArray(ByVal InputVal As Variant) As Variant
    On Error GoTo ErrHandler

    Dim result() As Variant

    If IsObject(InputVal) Then
        If InputVal Is Nothing Then
            ReDim result(1 To 1, 1 To 1)
            result(1, 1) = Empty
            Ensure2DArray = result
            Exit Function
        End If
        If TypeOf InputVal Is Excel.Range Then
            Dim r As Range: Set r = InputVal
            If r.Cells.Count = 1 Then
                ReDim result(1 To 1, 1 To 1)
                result(1, 1) = r.Value2
                Ensure2DArray = result
            Else
                Ensure2DArray = r.Value2
            End If
            Exit Function
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
    Ensure2DArray = CVErr(xlErrValue)
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
    On Error GoTo ErrHandler

    TryConvertToNumber = False
    If IsError(rawVal) Then Exit Function
    
    Dim txt As String
    txt = CStr(rawVal)
    txt = Trim$(txt)
    
    If txt = vbNullString Then Exit Function
    
    ' Exclude hexadecimal/octal syntax which IsNumeric allows but are not standard numeric strings
    If Left$(txt, 2) = "&H" Or Left$(txt, 2) = "&h" Or Left$(txt, 2) = "&O" Or Left$(txt, 2) = "&o" Then
        Exit Function
    End If
    
    ' 1. Check if it's already directly numeric
    If IsNumeric(txt) Then
        On Error Resume Next
        outVal = CDbl(txt)
        If Err.Number = 0 Then
            TryConvertToNumber = True
            Exit Function
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
            Exit Function
        End If
        Err.Clear
        On Error GoTo ErrHandler
    End If

CleanExit:
    Exit Function
ErrHandler:
    TryConvertToNumber = False
End Function

' Centralized helper to map Excel CVErr values to their string representations.
Public Function GetExcelErrorText(ByVal errVal As Variant) As String
    If Not IsError(errVal) Then
        GetExcelErrorText = CStr(errVal)
        Exit Function
    End If
    
    Select Case errVal
        Case CVErr(xlErrDiv0): GetExcelErrorText = "#DIV/0!"
        Case CVErr(xlErrNull): GetExcelErrorText = "#NULL!"
        Case CVErr(xlErrNA): GetExcelErrorText = "#N/A"
        Case CVErr(xlErrName): GetExcelErrorText = "#NAME?"
        Case CVErr(xlErrNum): GetExcelErrorText = "#NUM!"
        Case CVErr(xlErrRef): GetExcelErrorText = "#REF!"
        Case CVErr(xlErrValue): GetExcelErrorText = "#VALUE!"
        Case Else: GetExcelErrorText = "#ERROR!"
    End Select
End Function
