Attribute VB_Name = "Lib_XFilterFunction"
Option Explicit

' @Module: Lib_XFilterFunction
' @Category: Library
' @Description: UDF for advanced set filtering (Intersection, Difference) between ranges.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

' Filters Range_A based on existence (or non-existence) in Range_B.
' Acts like a set operation (INTERSECTION or DIFFERENCE).
'
' ARGUMENTS:
'   Range_A       : The source range or array to filter. Can be multi-column; 
'                   the first column of each row is used for comparison.
'   Range_B       : The reference range or array to check against.
'   code_number   : 1 = INTERSECTION (In A AND B)
'                   2 = DIFFERENCE   (In A but NOT in B)
'   if_empty      : Optional value to return if no matches are found. Defaults to "Not found".
'   case_sensitive: Optional boolean, True for case-sensitive lookups.
'
' RETURNS: A dynamic array that spills into the sheet.
' ==============================================================================
Public Function XFilter(ByVal Range_A As Variant, ByVal Range_B As Variant, Optional ByVal code_number As Integer = 1, Optional ByVal if_empty As Variant, Optional ByVal case_sensitive As Boolean = False) As Variant
    On Error GoTo ErrHandler
    
    Dim arrA As Variant, arrB As Variant
    Dim dictB As Object
    Dim resultArr() As Variant
    Dim r As Long, c As Long
    Dim resCount As Long
    Dim valA As Variant, valB As Variant
    
    ' --- Optimization: Read ranges into memory arrays ---
    arrA = Ensure2DArray(Range_A)
    arrB = Ensure2DArray(Range_B)
    
    ' 2. Use a Dictionary for O(1) lookup speed (Late Bound)
    Set dictB = CreateObject("Scripting.Dictionary")
    If case_sensitive Then
        dictB.CompareMode = 0 ' BinaryCompare (case-sensitive)
    Else
        dictB.CompareMode = 1 ' TextCompare (case-insensitive)
    End If
    
    For r = LBound(arrB, 1) To UBound(arrB, 1)
        For c = LBound(arrB, 2) To UBound(arrB, 2)
            valB = arrB(r, c)
            If Not IsError(valB) And Not IsEmpty(valB) Then
                If Not dictB.Exists(valB) Then dictB.Add valB, 1
            End If
        Next c
    Next r
    
    ' 3. Prepare Result Array
    ReDim resultArr(1 To UBound(arrA, 1), 1 To UBound(arrA, 2))
    resCount = 0
    
    ' 4. Process Range_A
    For r = LBound(arrA, 1) To UBound(arrA, 1)
        valA = arrA(r, 1) ' Key column
        
        Dim matchFound As Boolean, includeRow As Boolean
        matchFound = False
        If Not IsError(valA) And Not IsEmpty(valA) Then
            matchFound = dictB.Exists(valA)
        End If
        includeRow = False
        
        If code_number = 1 Then
            If matchFound Then includeRow = True
        ElseIf code_number = 2 Then
            If Not matchFound Then includeRow = True
        Else
            XFilter = CVErr(xlErrNum)
            GoTo CleanExit
        End If
        
        If includeRow And Not IsEmpty(valA) Then
            resCount = resCount + 1
            For c = LBound(arrA, 2) To UBound(arrA, 2)
                resultArr(resCount, c) = arrA(r, c)
            Next c
        End If
    Next r
    
    ' 5. Return Result
    If resCount = 0 Then
        If IsMissing(if_empty) Then
            XFilter = "Not found" ' Default "Not found" if no results
        Else
            XFilter = if_empty
        End If
    Else
        Dim finalResult() As Variant
        ReDim finalResult(1 To resCount, 1 To UBound(arrA, 2))
        For r = 1 To resCount
            For c = 1 To UBound(arrA, 2)
                finalResult(r, c) = resultArr(r, c)
            Next c
        Next r
        XFilter = finalResult
    End If
 
CleanExit:
    Exit Function
 
ErrHandler:
    Infra_Error.HandleError "XFilter", Err
    XFilter = CVErr(xlErrValue) ' #VALUE! on general error
End Function

' Helper function to ensure any variant input (Range, Array, or Scalar) is converted to a 1-based 2D array.
Private Function Ensure2DArray(ByVal InputVal As Variant) As Variant
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

' Returns registry info for this UDF.
' Structure:
'   - Name: String
'   - Description: String
'   - Category: String
'   - Syntax: String
'   - ArgumentDescriptions: Variant Array of Strings
Public Function GetUdfMetadata() As Object
    Dim metadata As Object
    Set metadata = CreateObject("Scripting.Dictionary")
    metadata.Add "Name", "XFilter"
    metadata.Add "Description", "Filters Range_A based on existence (or non-existence) in Range_B."
    metadata.Add "Category", "User Defined"
    metadata.Add "Syntax", "XFilter(Range_A, Range_B, [code_number], [if_empty], [case_sensitive])"
    metadata.Add "ArgumentDescriptions", Array( _
        "The source range or array to filter. Comparison is done using the first column.", _
        "The reference range or array to check against.", _
        "Optional Mode: 1 = Intersection (default, In A and B), 2 = Difference (In A but not in B).", _
        "Optional value to return if no match is found (default is ""Not found"").", _
        "Optional. Set to True for case-sensitive matching; defaults to False.")
    Set GetUdfMetadata = metadata
End Function
