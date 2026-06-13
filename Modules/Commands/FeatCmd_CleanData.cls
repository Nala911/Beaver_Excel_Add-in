VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_CleanData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_CleanData
' @Category: Feature
' @Description: Advanced data cleaning tools (Trim, Clean).

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorksheet(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanData.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As Infra_ActionContext
    Dim request As Infra_CleanDataRequest
    Dim stateInfo As Collection

    Set stateInfo = Infra_Error.NewStateInfo()

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Or ctx.WorkbookRef Is Nothing Then GoTo CleanExit
    
    Infra_Error.AddState stateInfo, "Workbook", ctx.WorkbookRef.Name
    Infra_Error.AddState stateInfo, "Worksheet", ctx.WorksheetRef.Name
    If ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing Then
        Infra_Error.AddState stateInfo, "Selection", ctx.SelectionRange.Address(False, False)
    End If

    Set request = Infra_UIFactory.ShowCleanDataDialog(ctx)
    If request Is Nothing Then GoTo CleanExit
    
    ExecuteCleanData request

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "CleanData.Execute", Err, stateInfo
    Resume CleanExit
End Sub

Private Sub ExecuteCleanData(ByVal request As Infra_CleanDataRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteCleanData")
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim stateInfo As Collection
    Dim cleanedCount As Long
    Dim scopeLabel As String
    Dim originalScreenUpdating As Boolean

    Set stateInfo = Infra_Error.NewStateInfo()
    
    If request Is Nothing Then GoTo CleanExit
    If request.Context Is Nothing Then GoTo CleanExit
    Infra_Error.AddState stateInfo, "Scope", request.Scope
    
    ' --- 1. Save State for Undo ---
    Select Case request.Scope
        Case TargetScopeSelection
            Set targetRange = request.Context.SelectionRange
            scopeLabel = "selection"
        Case TargetScopeActiveSheet
            Set targetRange = request.Context.WorksheetRef.UsedRange
            scopeLabel = "active sheet"
        Case TargetScopeWorkbook
            Set targetRange = Nothing
            scopeLabel = "workbook"
    End Select
    
    If Not targetRange Is Nothing Then
        If Not Infra_Undo.SaveStateOrConfirm(targetRange, "Clean Data") Then GoTo CleanExit
    End If
    
    ' --- 2. Initialize Progress & Dynamic Screen Updating ---
    originalScreenUpdating = Application.ScreenUpdating
    If request.Scope = TargetScopeWorkbook Then
        Application.ScreenUpdating = False
        Infra_Progress.StartProgress "Cleaning Data", request.Context.WorkbookRef.Worksheets.Count
    Else
        Infra_Progress.StartProgress "Cleaning Data", 1
    End If
    
    Select Case request.Scope
        Case TargetScopeSelection
            cleanedCount = ProcessRangeForClean(request.Context.SelectionRange, request)
            Infra_Progress.UpdateProgress 1, True
        Case TargetScopeActiveSheet
            cleanedCount = ProcessRangeForClean(request.Context.WorksheetRef.UsedRange, request)
            If request.CleanBrokenNames Then
                CleanNamedRanges request.Context.WorkbookRef, request.Context.WorksheetRef
            End If
            Infra_Progress.UpdateProgress 1, True
        Case TargetScopeWorkbook
            If request.CleanBrokenNames Then
                CleanNamedRanges request.Context.WorkbookRef, Nothing
            End If
            Dim i As Long, count As Long
            count = request.Context.WorkbookRef.Worksheets.Count
            For i = 1 To count
                Set ws = request.Context.WorkbookRef.Worksheets(i)
                cleanedCount = cleanedCount + ProcessRangeForClean(ws.UsedRange, request)
                Infra_Progress.UpdateProgress i, True
                If Infra_Progress.UserCancelled Then Exit For
            Next i
    End Select
    
    Dim msg As String
    msg = "Done! Cleaned text and removed broken names in the " & scopeLabel & "."
    
    Infra_Interaction.ShowInfo msg

CleanExit:
    If Not request Is Nothing Then
        If request.Scope = TargetScopeWorkbook Then
            Application.ScreenUpdating = originalScreenUpdating
        End If
    End If
    Infra_Progress.EndProgress
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ExecuteCleanData", Err, stateInfo
    Resume CleanExit
End Sub

Private Function ProcessRangeForClean(ByRef targetRange As Range, ByVal request As Infra_CleanDataRequest) As Long
    ' Note: No tracker here for performance; this is a core processing loop.
    If targetRange Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = targetRange.Worksheet
    
    ' Intersect with UsedRange if targetRange is large to prevent scanning millions of empty cells
    If targetRange.Cells.CountLarge > 100000 Then
        Dim usedRange As Range
        On Error Resume Next
        Set usedRange = Application.Intersect(targetRange, ws.UsedRange)
        On Error GoTo 0
        If Not usedRange Is Nothing Then
            Set targetRange = usedRange
        Else
            Exit Function
        End If
    End If

    Dim textCells As Range
    Dim changeCount As Long
    Dim dataArr As Variant

    If targetRange.CountLarge = 1 Then
        If Not targetRange.HasFormula And VarType(targetRange.Value) = vbString Then
            Set textCells = targetRange
        End If
    Else
        On Error Resume Next
        Set textCells = targetRange.SpecialCells(xlCellTypeConstants, xlTextValues)
        On Error GoTo 0
    End If

    If textCells Is Nothing Then Exit Function

    Dim chunks As Collection
    Set chunks = Infra_CommandSupport.GetChunkedRanges(textCells, 20000)
    
    Dim chunkRange As Range
    For Each chunkRange In chunks
        If chunkRange.Cells.CountLarge = 1 Then
            Dim oldVal As Variant, newVal As Variant
            oldVal = chunkRange.Value
            newVal = CleanTextValue(oldVal, request)
            If CStr(newVal) <> CStr(oldVal) Or TypeName(newVal) <> TypeName(oldVal) Then
                chunkRange.Value = newVal
                changeCount = changeCount + 1
            End If
        Else
            dataArr = chunkRange.Value
            changeCount = changeCount + CleanTextArray(dataArr, request)
            chunkRange.Value = dataArr
        End If
        
        DoEvents
        If Infra_Progress.UserCancelled Then Exit For
    Next chunkRange
    
    ProcessRangeForClean = changeCount
End Function

Private Function CleanTextArray(ByRef dataArr As Variant, ByVal request As Infra_CleanDataRequest) As Long
    ' Note: No tracker here for performance.
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim oldVal As Variant
    Dim newVal As Variant
    Dim changeCount As Long

    For rowIndex = LBound(dataArr, 1) To UBound(dataArr, 1)
        For columnIndex = LBound(dataArr, 2) To UBound(dataArr, 2)
            oldVal = dataArr(rowIndex, columnIndex)
            newVal = CleanTextValue(oldVal, request)
            If CStr(newVal) <> CStr(oldVal) Or TypeName(newVal) <> TypeName(oldVal) Then
                dataArr(rowIndex, columnIndex) = newVal
                changeCount = changeCount + 1
            End If
        Next columnIndex
    Next rowIndex
    
    CleanTextArray = changeCount
End Function

Private Function CleanTextValue(ByVal rawValue As Variant, ByVal request As Infra_CleanDataRequest) As Variant
    ' Note: No tracker here. This is called in tight loops.
    If IsError(rawValue) Then
        CleanTextValue = rawValue
        Exit Function
    End If

    Dim textValue As String
    textValue = CStr(rawValue)
    
    If textValue = vbNullString Then
        CleanTextValue = textValue
        Exit Function
    End If
    
    ' --- 1. Standardize Invisible Characters ---
    If request.CleanInvisibleChars Then
        ' Zero-width characters & BOM -> remove
        If InStr(1, textValue, ChrW$(8203)) > 0 Then textValue = Replace(textValue, ChrW$(8203), vbNullString)
        If InStr(1, textValue, ChrW$(8204)) > 0 Then textValue = Replace(textValue, ChrW$(8204), vbNullString)
        If InStr(1, textValue, ChrW$(8205)) > 0 Then textValue = Replace(textValue, ChrW$(8205), vbNullString)
        If InStr(1, textValue, ChrW$(65279)) > 0 Then textValue = Replace(textValue, ChrW$(65279), vbNullString)
        
        ' Unicode spaces -> replace with standard space
        If InStr(1, textValue, ChrW$(8201)) > 0 Then textValue = Replace(textValue, ChrW$(8201), " ") ' U+2009 thin space
        If InStr(1, textValue, ChrW$(8239)) > 0 Then textValue = Replace(textValue, ChrW$(8239), " ") ' U+202F narrow no-break space
        If InStr(1, textValue, ChrW$(8194)) > 0 Then textValue = Replace(textValue, ChrW$(8194), " ") ' U+2002 en space
        If InStr(1, textValue, ChrW$(8195)) > 0 Then textValue = Replace(textValue, ChrW$(8195), " ") ' U+2003 em space
    End If

    ' --- 2. Replace Non-Breaking Spaces (Chr 160) ---
    If request.CleanTrimSpaces Then
        If InStr(1, textValue, ChrW$(160)) > 0 Then
            textValue = Replace(textValue, ChrW$(160), " ")
        End If
    End If

    ' --- 3. Line Breaks Processing ---
    Dim hasLineBreaks As Boolean
    Dim containsLF As Boolean, containsCR As Boolean
    containsLF = (InStr(1, textValue, vbLf) > 0)
    containsCR = (InStr(1, textValue, vbCr) > 0)
    hasLineBreaks = containsLF Or containsCR

    ' Temporary placeholder for LF to protect it from Application.Clean if needed
    Const LF_PLACEHOLDER As String = "" ' Private Use Area character U+E000
    Dim lineBreaksProtected As Boolean
    lineBreaksProtected = False

    If hasLineBreaks Then
        If request.CleanReplaceLineBreaksWithSpace Then
            textValue = Replace(textValue, vbCrLf, " ")
            textValue = Replace(textValue, vbCr, " ")
            textValue = Replace(textValue, vbLf, " ")
            hasLineBreaks = False
        ElseIf request.CleanRemoveLineBreaks Then
            textValue = Replace(textValue, vbCrLf, vbNullString)
            textValue = Replace(textValue, vbCr, vbNullString)
            textValue = Replace(textValue, vbLf, vbNullString)
            hasLineBreaks = False
        ElseIf request.CleanStandardizeLineBreaks Then
            textValue = Replace(textValue, vbCrLf, vbLf)
            textValue = Replace(textValue, vbCr, vbLf)
            
            Do While InStr(1, textValue, vbLf & vbLf) > 0
                textValue = Replace(textValue, vbLf & vbLf, vbLf)
            Loop
            
            If request.CleanNonPrintables Then
                textValue = Replace(textValue, vbLf, LF_PLACEHOLDER)
                lineBreaksProtected = True
            End If
        End If
    End If

    ' --- 4. Application.Clean & Application.Trim ---
    If request.CleanTrimSpaces And request.CleanNonPrintables Then
        textValue = Application.Trim(Application.Clean(textValue))
    ElseIf request.CleanTrimSpaces Then
        textValue = Application.Trim(textValue)
    ElseIf request.CleanNonPrintables Then
        textValue = Application.Clean(textValue)
    End If

    ' --- 5. Restore Line Breaks ---
    If lineBreaksProtected Then
        textValue = Replace(textValue, LF_PLACEHOLDER, vbLf)
    End If

    ' --- 5b. Remove Special Symbols ---
    If request.CleanSpecialSymbols Then
        If InStr(1, textValue, Chr$(9)) > 0 Then textValue = Replace(textValue, Chr$(9), vbNullString)
        If InStr(1, textValue, ChrW$(8226)) > 0 Then textValue = Replace(textValue, ChrW$(8226), vbNullString)
        If InStr(1, textValue, ChrW$(8482)) > 0 Then textValue = Replace(textValue, ChrW$(8482), vbNullString)
        If InStr(1, textValue, ChrW$(174)) > 0 Then textValue = Replace(textValue, ChrW$(174), vbNullString)
        If InStr(1, textValue, ChrW$(169)) > 0 Then textValue = Replace(textValue, ChrW$(169), vbNullString)
    End If

    ' --- 5c. Standardize Dashes ---
    If request.CleanStandardizeDashes Then
        If InStr(1, textValue, ChrW$(8211)) > 0 Then textValue = Replace(textValue, ChrW$(8211), "-")
        If InStr(1, textValue, ChrW$(8212)) > 0 Then textValue = Replace(textValue, ChrW$(8212), "-")
        If InStr(1, textValue, ChrW$(8722)) > 0 Then textValue = Replace(textValue, ChrW$(8722), "-")
    End If

    ' --- 5d. Remove Accents ---
    If request.CleanRemoveAccents Then
        If HasNonAscii(textValue) Then
            textValue = RemoveAccents(textValue)
        End If
    End If

    ' --- 6. Convert Numeric Text ---
    If request.CleanConvertNumbers Then
        If IsNumeric(textValue) Then
            Dim cleanedTrimmed As String
            cleanedTrimmed = Trim$(textValue)
            If Not (Left$(cleanedTrimmed, 2) = "&H" Or Left$(cleanedTrimmed, 2) = "&h" Or _
                    Left$(cleanedTrimmed, 2) = "&O" Or Left$(cleanedTrimmed, 2) = "&o") Then
                On Error Resume Next
                Dim numericVal As Double
                numericVal = CDbl(textValue)
                If Err.Number = 0 Then
                    CleanTextValue = numericVal
                    Exit Function
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    End If

    CleanTextValue = textValue
End Function

Private Function HasNonAscii(ByRef text As String) As Boolean
    Dim i As Long, charCode As Long
    For i = 1 To Len(text)
        charCode = AscW(Mid$(text, i, 1))
        If charCode < 0 Or charCode > 127 Then
            HasNonAscii = True
            Exit Function
        End If
    Next i
End Function

Private Function RemoveAccents(ByVal text As String) As String
    Dim accentedChars As String
    accentedChars = GetAccentedChars()
    
    Const NORMAL_CHARS   As String = "AAAAAACEEEEIIIINOOOOOOUUUUYaaaaaaceeeeiiiinoooooouuuuyySsZzCcDdEeNnRrTtUu"
    
    Dim i As Long, pos As Long
    Dim char As String * 1
    Dim result As String
    
    result = text
    For i = 1 To Len(result)
        char = Mid$(result, i, 1)
        pos = InStr(1, accentedChars, char, vbBinaryCompare)
        If pos > 0 Then
            Mid$(result, i, 1) = Mid$(NORMAL_CHARS, pos, 1)
        End If
    Next i
    
    ' Handle multi-character replacements (like Æ -> AE, æ -> ae, ß -> ss, Œ -> OE, œ -> oe)
    If InStr(1, result, ChrW$(198), vbBinaryCompare) > 0 Then result = Replace(result, ChrW$(198), "AE")
    If InStr(1, result, ChrW$(230), vbBinaryCompare) > 0 Then result = Replace(result, ChrW$(230), "ae")
    If InStr(1, result, ChrW$(223), vbBinaryCompare) > 0 Then result = Replace(result, ChrW$(223), "ss")
    If InStr(1, result, ChrW$(338), vbBinaryCompare) > 0 Then result = Replace(result, ChrW$(338), "OE")
    If InStr(1, result, ChrW$(339), vbBinaryCompare) > 0 Then result = Replace(result, ChrW$(339), "oe")
    
    RemoveAccents = result
End Function

Private Function GetAccentedChars() As String
    Dim chars As String
    ' ÀÁÂÃÄÅ
    chars = chars & ChrW$(192) & ChrW$(193) & ChrW$(194) & ChrW$(195) & ChrW$(196) & ChrW$(197)
    ' Ç
    chars = chars & ChrW$(199)
    ' ÈÉÊË
    chars = chars & ChrW$(200) & ChrW$(201) & ChrW$(202) & ChrW$(203)
    ' ÌÍÎÏ
    chars = chars & ChrW$(204) & ChrW$(205) & ChrW$(206) & ChrW$(207)
    ' Ñ
    chars = chars & ChrW$(209)
    ' ÒÓÔÕÖØ
    chars = chars & ChrW$(210) & ChrW$(211) & ChrW$(212) & ChrW$(213) & ChrW$(214) & ChrW$(216)
    ' ÙÚÛÜ
    chars = chars & ChrW$(217) & ChrW$(218) & ChrW$(219) & ChrW$(220)
    ' Ý
    chars = chars & ChrW$(221)
    ' àáâãäå
    chars = chars & ChrW$(224) & ChrW$(225) & ChrW$(226) & ChrW$(227) & ChrW$(228) & ChrW$(229)
    ' ç
    chars = chars & ChrW$(231)
    ' èéêë
    chars = chars & ChrW$(232) & ChrW$(233) & ChrW$(234) & ChrW$(235)
    ' ìíîï
    chars = chars & ChrW$(236) & ChrW$(237) & ChrW$(238) & ChrW$(239)
    ' ñ
    chars = chars & ChrW$(241)
    ' òóôõöø
    chars = chars & ChrW$(242) & ChrW$(243) & ChrW$(244) & ChrW$(245) & ChrW$(246) & ChrW$(248)
    ' ùúûü
    chars = chars & ChrW$(249) & ChrW$(250) & ChrW$(251) & ChrW$(252)
    ' ýÿ
    chars = chars & ChrW$(253) & ChrW$(255)
    ' ŠšŽž
    chars = chars & ChrW$(352) & ChrW$(353) & ChrW$(381) & ChrW$(382)
    ' ČčĎďĚěŇňŘřŤťŮů
    chars = chars & ChrW$(268) & ChrW$(269) & ChrW$(270) & ChrW$(271) & ChrW$(282) & ChrW$(283)
    chars = chars & ChrW$(327) & ChrW$(328) & ChrW$(344) & ChrW$(345) & ChrW$(356) & ChrW$(357)
    chars = chars & ChrW$(366) & ChrW$(367)
    GetAccentedChars = chars
End Function

Private Sub CleanNamedRanges(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim nm As Name
    Dim i As Long
    
    On Error Resume Next
    If Not ws Is Nothing Then
        For i = ws.Names.Count To 1 Step -1
            Set nm = ws.Names(i)
            If InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Then
                nm.Delete
            End If
        Next i
    ElseIf Not wb Is Nothing Then
        For i = wb.Names.Count To 1 Step -1
            Set nm = wb.Names(i)
            If InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Then
                nm.Delete
            End If
        Next i
    End If
    On Error GoTo 0
End Sub


' --- Testing Entry Points ---
' Allows headless automated testing of the range cleaning logic.
Public Function CleanRangeDirect(ByVal targetRange As Range) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanRangeDirect")
    On Error GoTo ErrHandler
    
    Dim request As New Infra_CleanDataRequest
    CleanRangeDirect = ProcessRangeForClean(targetRange, request)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CleanRangeDirect", Err
    Resume CleanExit
End Function

Public Function CleanRangeWithOptionsDirect(ByVal targetRange As Range, ByVal request As Infra_CleanDataRequest) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanRangeWithOptionsDirect")
    On Error GoTo ErrHandler
    
    CleanRangeWithOptionsDirect = ProcessRangeForClean(targetRange, request)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CleanRangeWithOptionsDirect", Err
    Resume CleanExit
End Function
