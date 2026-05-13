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
'   Range_A     : The source range to filter. Can be multi-column; 
'                 the first column of each row is used for comparison.
'   Range_B     : The reference range to check against.
'   code_number : 1 = INTERSECTION (In A AND B)
'                 2 = DIFFERENCE   (In A but NOT in B)
'
' RETURNS: A dynamic array that spills into the sheet.
' ==============================================================================
Public Function XFilter(Range_A As Range, Range_B As Range, code_number As Integer) As Variant
    Infra_Error.PushContext "XFilter"
    On Error GoTo ErrHandler
    
    Dim arrA As Variant, arrB As Variant
    Dim dictB As Object
    Dim resultArr() As Variant
    Dim r As Long, c As Long
    Dim resCount As Long
    Dim valA As Variant, valB As Variant
    
    ' --- Optimization: Read ranges into memory arrays ---
    If Range_A.Cells.Count = 1 Then
        ReDim arrA(1 To 1, 1 To 1)
        arrA(1, 1) = Range_A.Value2
    Else
        arrA = Range_A.Value2
    End If
    
    If Range_B.Cells.Count = 1 Then
        ReDim arrB(1 To 1, 1 To 1)
        arrB(1, 1) = Range_B.Value2
    Else
        arrB = Range_B.Value2
    End If
    
    ' 2. Use a Dictionary for O(1) lookup speed (Late Bound)
    Set dictB = CreateObject("Scripting.Dictionary")
    dictB.CompareMode = 1 ' 1 = TextCompare
    
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
        matchFound = dictB.Exists(valA)
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
        XFilter = CVErr(xlErrNull) ' #NULL! if no results
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
    Infra_Error.PopContext
    Exit Function

ErrHandler:
    Infra_Error.HandleError "XFilter", Err
    XFilter = CVErr(xlErrValue) ' #VALUE! on general error
End Function
