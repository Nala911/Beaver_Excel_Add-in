Attribute VB_Name = "Lib_DateUtils"
Option Explicit

' @Module: Lib_DateUtils
' @Category: Library
' @Description: Robust date parsing and resolution utilities.
' @ManagedBy: BeaverAddin Agent

''
' Parses a string or numeric value into a month index (1-12).
' Supports month names (short/long) and numeric strings.
' @Param rawValue: The value to parse.
' @Return: Long (1-12), or 0 if invalid.
''
Public Function ParseMonthValue(ByVal rawValue As Variant) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("ParseMonthValue")
    On Error GoTo ErrHandler
    
    Dim textValue As String
    Dim monthIndex As Long

    textValue = Trim$(CStr(rawValue))
    If textValue = vbNullString Then GoTo CleanExit

    If IsNumeric(textValue) Then
        ParseMonthValue = CLng(textValue)
        If ParseMonthValue < 1 Or ParseMonthValue > 12 Then ParseMonthValue = 0
        GoTo CleanExit
    End If

    textValue = UCase$(Left$(textValue, 3))
    For monthIndex = 1 To 12
        If UCase$(Left$(MonthName(monthIndex, False), 3)) = textValue Then
            ParseMonthValue = monthIndex
            GoTo CleanExit
        End If
    Next monthIndex

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ParseMonthValue", Err
    Resume CleanExit
End Function

''
' Attempts to resolve a text string into a Date, using a target month to break ambiguities.
' @Param txt: The raw text string (e.g., "01-09-1990").
' @Param targetMonth: The month index (1-12) to prioritize.
' @Param outDate: [Out] The resolved date if successful.
' @Return: Boolean True if resolved.
''
Public Function TryResolveDateWithMonth(ByVal txt As String, ByVal targetMonth As Integer, ByRef outDate As Date) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("TryResolveDateWithMonth")
    On Error GoTo ErrHandler

    Dim cleanTxt As String
    Dim parts() As String
    Dim i As Long
    Dim used(0 To 2) As Boolean
    Dim d As Integer, m As Integer, y As Integer
    Dim valPart As Long
    Dim partMonth As Long
    Dim yearIdx As Integer: yearIdx = -1
    Dim monthIdx As Integer: monthIdx = -1

    ' 1. Basic cleaning and splitting
    cleanTxt = Trim$(txt)
    If Len(cleanTxt) = 0 Then GoTo CleanExit
    
    ' Standardize delimiters
    cleanTxt = Replace(cleanTxt, "-", "/")
    cleanTxt = Replace(cleanTxt, ".", "/")
    cleanTxt = Replace(cleanTxt, " ", "/")
    While InStr(cleanTxt, "//") > 0: cleanTxt = Replace(cleanTxt, "//", "/"): Wend

    parts = Split(cleanTxt, "/")
    If UBound(parts) <> 2 Then GoTo CleanExit

    ' 2. PASS 1: Find 4-digit Year (1900-2100)
    For i = 0 To 2
        If IsNumeric(parts(i)) Then
            valPart = CLng(parts(i))
            If valPart >= 1000 And valPart <= 2100 Then
                y = valPart
                yearIdx = i
                used(i) = True
                Exit For
            End If
        End If
    Next i

    ' 3. PASS 2: Find Month (matching targetMonth)
    ' This handles "09", "9", "Sep", "September"
    For i = 0 To 2
        If Not used(i) Then
            partMonth = ParseMonthValue(parts(i))
            If partMonth = targetMonth Then
                m = targetMonth
                monthIdx = i
                used(i) = True
                Exit For
            End If
        End If
    Next i

    ' 4. PASS 3: Resolve remaining (Day and possibly Year)
    ' Case A: Year and Month already found. Remaining must be Day.
    If yearIdx <> -1 And monthIdx <> -1 Then
        For i = 0 To 2
            If Not used(i) Then
                d = CInt(Val(parts(i)))
                Exit For
            End If
        Next i
    
    ' Case B: Only Month found. We need to decide which of 2 parts is Year and which is Day.
    ElseIf monthIdx <> -1 Then
        ' Heuristic: If index 2 is unused, assume it's the year (DD/MM/YY).
        ' If index 2 is used (the month), and index 0 is unused, assume index 0 is year (YY/MM/DD).
        If Not used(2) Then
            yearIdx = 2
            y = CInt(Val(parts(2)))
            If y < 100 Then y = IIf(y < 50, 2000 + y, 1900 + y) ' Basic 50-year pivot
            used(2) = True
        Else
            yearIdx = 0
            y = CInt(Val(parts(0)))
            If y < 100 Then y = IIf(y < 50, 2000 + y, 1900 + y)
            used(0) = True
        End If
        
        ' Last one is day
        For i = 0 To 2
            If Not used(i) Then
                d = CInt(Val(parts(i)))
                Exit For
            End If
        Next i
    End If

    ' 5. Final validation and Date creation
    If y = 0 Then y = Year(Date)
    If m = 0 Then m = targetMonth
    If d <= 0 Or d > 31 Then d = 1

    On Error Resume Next
    outDate = DateSerial(y, m, d)
    If Err.Number <> 0 Then
        ' Fallback for invalid days (e.g. Feb 30)
        outDate = DateSerial(y, m, 1)
        Err.Clear
    End If
    TryResolveDateWithMonth = True
    On Error GoTo ErrHandler

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "TryResolveDateWithMonth", Err
    TryResolveDateWithMonth = False
    Resume CleanExit
End Function
