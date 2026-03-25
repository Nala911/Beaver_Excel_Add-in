Attribute VB_Name = "Feat_DateConversion"
Option Explicit

' @Module: Feat_DateConversion
' @Category: Feature
' @Description: Converts text strings to proper Excel Date values with ambiguous format resolution.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_AppStateGuard, Infra_Config, Infra_Error, Infra_UIFactory

' Parses text-based dates in the selected column. Asks user for the target month
' to resolve ambiguity between day/month/year components.
Public Sub ConvertTextToProperDate()
    PushContext "ConvertTextToProperDate"
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim rng As Range
    Dim dataArr As Variant
    Dim i As Long, j As Long
    Dim txt As String
    Dim parts() As String
    Dim elem As String
    Dim d As Integer, m As Integer, y As Integer
    Dim targetMonth As Integer
    Dim valElem As Long
    Dim ctx As Infra_ActionContext

    ' 1. Validate selection
    If Not Infra_AppState.IsRangeSelected() Then GoTo CleanExit
    Set rng = Selection
    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit

    ' 2. Safety Check: single column only
    If rng.Columns.Count > 1 Then
        MsgBox "Please select only one column of dates at a time.", vbExclamation, Infra_Config.ADDIN_NAME
        GoTo CleanExit
    End If

    ' 3. Show screen so user can see context, then ask for the month
    Application.ScreenUpdating = True
    targetMonth = Infra_UIFactory.PromptForDateConversionMonth(ctx)
    If targetMonth = 0 Then GoTo CleanExit
    Application.ScreenUpdating = False

    ' 4. Read data into memory array
    If rng.Cells.Count = 1 Then
        ReDim dataArr(1 To 1, 1 To 1)
        dataArr(1, 1) = rng.Value
    Else
        dataArr = rng.Value
    End If

    ' 5. Loop through the array (process in memory)
    For i = LBound(dataArr, 1) To UBound(dataArr, 1)
        txt = Trim(CStr(dataArr(i, 1)))

        If Len(txt) > 0 Then
            ' Standardize common delimiters
            txt = Replace(txt, "-", "/")
            txt = Replace(txt, ".", "/")
            txt = Replace(txt, " ", "/")
            
            ' Reduce multiple slashes to single
            While InStr(txt, "//") > 0: txt = Replace(txt, "//", "/"): Wend
            
            parts = Split(txt, "/")

            If UBound(parts) = 2 Then
                d = 0: m = 0: y = 0

                ' Identify each element
                For j = LBound(parts) To UBound(parts)
                    elem = Trim(parts(j))

                    If IsNumeric(elem) Then
                        valElem = CLng(elem)

                        ' 1. Year check (4-digit)
                        If (valElem >= 1900 And valElem <= 2099) And y = 0 Then
                            y = valElem

                        ' 2. Month match (user-specified)
                        ' Only assign for the FIRST match found
                        ElseIf valElem = targetMonth And m = 0 Then
                            m = targetMonth

                        ' 3. Year check (2-digit fallback)
                        ' Only if we haven't found a 4-digit year yet
                        ElseIf (valElem >= 0 And valElem <= 99) And y = 0 Then
                            y = 2000 + valElem

                        ' 4. Remaining must be day
                        ElseIf d = 0 Then
                            d = CInt(valElem)
                        End If
                    End If
                Next j

                ' Validation and fallbacks
                If y = 0 Then y = Year(Date)
                If m = 0 Then m = targetMonth
                If d = 0 Or d > 31 Then d = 1

                ' Overwrite array value with the actual Date Serial
                On Error Resume Next
                dataArr(i, 1) = DateSerial(y, m, d)
                On Error GoTo ErrHandler
            End If
        End If
    Next i

    ' 6. Dump array back to Excel (one single write operation)
    rng.Value = dataArr
    rng.NumberFormat = Infra_Config.DEFAULT_DATE_FORMAT

    MsgBox "Conversion Complete!", vbInformation, Infra_Config.ADDIN_NAME

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "ConvertTextToProperDate", Err
End Sub
