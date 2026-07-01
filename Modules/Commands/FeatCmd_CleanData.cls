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
Implements IBatchProcessor
Implements ICellTransformer

' Trigger incremental build lock recovery test 2

' @Module: FeatCmd_CleanData
' @Category: Feature
' @Description: Advanced data cleaning tools (Trim, Clean).

Private m_CurrentRequest As CleanDataRequest

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanData.Validate")
    On Error GoTo ErrHandler

    Dim result As CommandValidationResult
    Set result = Infra_CommandSupport.ValidateHasWorksheet(context)
    If Not result.IsExecutable Then
        Set ICommand_Validate = result
        Exit Function
    End If

    Set result = Infra_CommandSupport.ValidateWorksheetNotProtected(context, "Cannot clean data on a protected worksheet. Please unprotect the sheet and try again.")
    If Not result.IsExecutable Then
        Set ICommand_Validate = result
        Exit Function
    End If

    Set ICommand_Validate = Infra_CommandSupport.ValidationSuccess()

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "CleanData.Validate", Err
    Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("The data cannot be cleaned right now.")
    Resume CleanExit
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanData.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As ActionContext
    Dim request As CleanDataRequest
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

    Set request = UI_Factory.ShowCleanDataDialog(ctx)
    If request Is Nothing Then GoTo CleanExit
    
    ExecuteCleanData request

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "CleanData.Execute", Err, stateInfo
    Resume CleanExit
End Sub

Private Sub ExecuteCleanData(ByVal request As CleanDataRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteCleanData")
    On Error GoTo ErrHandler
    
    If request Is Nothing Then GoTo CleanExit
    If request.Context Is Nothing Then GoTo CleanExit
    
    Set m_CurrentRequest = request
    
    Dim totalCleaned As Long
    totalCleaned = Infra_BatchProcessor.Execute(request.Context, request.Scope, Me, "Clean Data", False, UndoCaptureFormulaOnly)
    
    Dim scopeLabel As String
    Select Case request.Scope
        Case TargetScopeSelection: scopeLabel = "selection"
        Case TargetScopeActiveSheet: scopeLabel = "active sheet"
        Case TargetScopeWorkbook: scopeLabel = "workbook"
    End Select
    
    Dim msg As String
    msg = "Done! Cleaned text and removed broken names in the " & scopeLabel & "."
    Infra_Interaction.ShowInfo msg, "Clean Data"

CleanExit:
    Set m_CurrentRequest = Nothing
    Exit Sub

ErrHandler:
    Set m_CurrentRequest = Nothing
    Infra_Error.HandleError "ExecuteCleanData", Err
    Resume CleanExit
End Sub

Private Function IBatchProcessor_ProcessRange(ByVal targetRange As Range) As Long
    If m_CurrentRequest Is Nothing Then
        IBatchProcessor_ProcessRange = 0
        Exit Function
    End If
    IBatchProcessor_ProcessRange = ProcessRangeForClean(targetRange, m_CurrentRequest)
End Function

Private Function IBatchProcessor_ProcessWorksheet(ByVal ws As Worksheet) As Long
    If m_CurrentRequest Is Nothing Then
        IBatchProcessor_ProcessWorksheet = 0
        Exit Function
    End If

    Dim namesRemoved As Long
    Dim shapesRemoved As Long
    
    If m_CurrentRequest.CleanBrokenNames Then
        Dim tempNamesRemoved As Long
        If m_CurrentRequest.Scope = TargetScopeActiveSheet Then
            Infra_CommandSupport.CleanWorkbookNames m_CurrentRequest.Context.WorkbookRef, ws, NameCleanCriteriaBroken, tempNamesRemoved
        ElseIf m_CurrentRequest.Scope = TargetScopeWorkbook Then
            Dim isFirstSheet As Boolean
            isFirstSheet = (ws.Name = m_CurrentRequest.Context.WorkbookRef.Worksheets(1).Name)
            If isFirstSheet Then
                Infra_CommandSupport.CleanWorkbookNames m_CurrentRequest.Context.WorkbookRef, Nothing, NameCleanCriteriaBroken, tempNamesRemoved
            End If
        End If
        namesRemoved = namesRemoved + tempNamesRemoved
    End If
    
    If m_CurrentRequest.CleanSheetNames Then
        Dim sheetNamesRemoved As Long
        CleanSheetScopedNames ws, sheetNamesRemoved
        namesRemoved = namesRemoved + sheetNamesRemoved
    End If
    
    If m_CurrentRequest.CleanShapes And m_CurrentRequest.Scope <> TargetScopeSelection Then
        CleanWorksheetShapes ws, shapesRemoved
    End If
    
    IBatchProcessor_ProcessWorksheet = namesRemoved + shapesRemoved
End Function

Private Function IsDeletableShape(ByVal sh As Shape) As Boolean
    ' Filter out non-deletable shapes (Charts, Slicers, Form Controls, ActiveX, Comments)
    Select Case sh.Type
        Case msoChart, msoSlicer, msoFormControl, msoOLEControlObject, msoComment
            IsDeletableShape = False
        Case Else
            IsDeletableShape = True
    End Select
End Function

Private Sub CleanShapesInSelection(ByVal targetRange As Range, ByRef outCount As Long)
    Dim ws As Worksheet
    Set ws = targetRange.Worksheet
    Dim sh As Shape
    Dim i As Long
    outCount = 0
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If IsDeletableShape(sh) Then
            Dim intersectRange As Range
            On Error Resume Next
            Set intersectRange = Application.Intersect(sh.TopLeftCell, targetRange)
            On Error GoTo 0
            If Not intersectRange Is Nothing Then
                On Error Resume Next
                sh.Delete
                If Err.Number = 0 Then outCount = outCount + 1
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

Private Sub CleanWorksheetShapes(ByVal ws As Worksheet, ByRef outCount As Long)
    Dim sh As Shape
    Dim i As Long
    outCount = 0
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If IsDeletableShape(sh) Then
            On Error Resume Next
            sh.Delete
            If Err.Number = 0 Then outCount = outCount + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next i
End Sub

Private Sub CleanSheetScopedNames(ByVal ws As Worksheet, ByRef outCount As Long)
    Dim i As Long
    outCount = 0
    For i = ws.Names.Count To 1 Step -1
        On Error Resume Next
        ws.Names(i).Delete
        If Err.Number = 0 Then outCount = outCount + 1
        Err.Clear
        On Error GoTo 0
    Next i
End Sub




Private Function ICellTransformer_TransformCell(ByVal rawValue As Variant, ByVal currentFormat As String, ByRef outValue As Variant, ByRef outFormat As String) As Boolean
    If m_CurrentRequest Is Nothing Then
        ICellTransformer_TransformCell = False
        Exit Function
    End If
    
    Dim newVal As Variant
    newVal = CleanTextValue(rawValue, m_CurrentRequest)
    
    ' Check if value changed
    If CStr(newVal) <> CStr(rawValue) Or TypeName(newVal) <> TypeName(rawValue) Then
        outValue = newVal
        outFormat = currentFormat
        ICellTransformer_TransformCell = True
    Else
        ICellTransformer_TransformCell = False
    End If
End Function

Private Function ProcessRangeForClean(ByRef targetRange As Range, ByVal request As CleanDataRequest) As Long
    ' Note: No tracker here for performance; this is a core processing loop.
    If targetRange Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = targetRange.Worksheet
    
    ' Intersect with UsedRange if targetRange is large to prevent scanning millions of empty cells
    Set targetRange = Infra_CommandSupport.GetSafeProcessingRange(targetRange, 100000, False)
    If targetRange Is Nothing Then Exit Function

    ' Temporarily set current request to support direct/test calls
    Dim prevRequest As CleanDataRequest
    Set prevRequest = m_CurrentRequest
    Set m_CurrentRequest = request

    Dim changeCount As Long

    ' --- Apply cell-level hygiene operations ---
    If request.CleanComments Then
        Dim commentCells As Range
        If targetRange.CountLarge = 1 Then
            Dim hasComment As Boolean
            hasComment = False
            On Error Resume Next
            hasComment = Not (targetRange.Comment Is Nothing)
            On Error GoTo 0
            If hasComment Then Set commentCells = targetRange
        Else
            On Error Resume Next
            Set commentCells = targetRange.SpecialCells(xlCellTypeComments)
            On Error GoTo 0
        End If
        If Not commentCells Is Nothing Then
            changeCount = changeCount + commentCells.Cells.CountLarge
            On Error Resume Next
            targetRange.ClearComments
            On Error GoTo 0
        End If
    End If

    If request.CleanValidation Then
        Dim validationCells As Range
        If targetRange.CountLarge = 1 Then
            Dim hasVal As Boolean
            hasVal = False
            On Error Resume Next
            Dim valType As Long
            valType = targetRange.Validation.Type
            If Err.Number = 0 Then hasVal = True
            On Error GoTo 0
            If hasVal Then Set validationCells = targetRange
        Else
            On Error Resume Next
            Set validationCells = targetRange.SpecialCells(xlCellTypeAllValidation)
            On Error GoTo 0
        End If
        If Not validationCells Is Nothing Then
            changeCount = changeCount + validationCells.Cells.CountLarge
            On Error Resume Next
            targetRange.Validation.Delete
            On Error GoTo 0
        End If
    End If

    If request.CleanConditionalFormatting Then
        Dim cfCells As Range
        If targetRange.CountLarge = 1 Then
            If targetRange.FormatConditions.Count > 0 Then Set cfCells = targetRange
        Else
            On Error Resume Next
            Set cfCells = targetRange.SpecialCells(xlCellTypeAllFormatConditions)
            On Error GoTo 0
        End If
        If Not cfCells Is Nothing Then
            changeCount = changeCount + cfCells.Cells.CountLarge
            On Error Resume Next
            targetRange.FormatConditions.Delete
            On Error GoTo 0
        End If
    End If

    If request.CleanFormats Then
        changeCount = changeCount + targetRange.Cells.CountLarge
        On Error Resume Next
        targetRange.ClearFormats
        On Error GoTo 0
    End If

    If request.CleanShapes And request.Scope = TargetScopeSelection Then
        Dim shapesDeleted As Long
        CleanShapesInSelection targetRange, shapesDeleted
        changeCount = changeCount + shapesDeleted
    End If

    ' --- Apply text-level cleaning operations ---
    Dim hasTextCleaningOptions As Boolean
    hasTextCleaningOptions = request.CleanTrimSpaces Or request.CleanNonPrintables Or _
                             request.CleanInvisibleChars Or request.CleanReplaceLineBreaksWithSpace Or _
                             request.CleanRemoveLineBreaks Or request.CleanStandardizeLineBreaks Or _
                             request.CleanConvertNumbers Or request.CleanSpecialSymbols Or _
                             request.CleanStandardizeDashes Or request.CleanRemoveAccents

    If hasTextCleaningOptions Then
        changeCount = changeCount + Infra_CommandSupport.ProcessRangeUnified(targetRange, Me, xlTextValues)
    End If

    ' Restore previous request
    Set m_CurrentRequest = prevRequest

    ProcessRangeForClean = changeCount
End Function

Private Function CleanTextValue(ByVal rawValue As Variant, ByVal request As CleanDataRequest) As Variant
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
        Dim numericVal As Variant
        If Infra_ValueConversion.TryConvertToNumber(textValue, numericVal) Then
            CleanTextValue = numericVal
            Exit Function
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


' --- Testing Entry Points ---
' Allows headless automated testing of the range cleaning logic.
Public Function CleanRangeDirect(ByVal targetRange As Range) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanRangeDirect")
    On Error GoTo ErrHandler
    
    Dim request As New CleanDataRequest
    CleanRangeDirect = ProcessRangeForClean(targetRange, request)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CleanRangeDirect", Err
    Resume CleanExit
End Function

Public Function CleanRangeWithOptionsDirect(ByVal targetRange As Range, ByVal request As CleanDataRequest) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanRangeWithOptionsDirect")
    On Error GoTo ErrHandler
    
    CleanRangeWithOptionsDirect = ProcessRangeForClean(targetRange, request)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CleanRangeWithOptionsDirect", Err
    Resume CleanExit
End Function

Public Function CleanWorksheetWithOptionsDirect(ByVal ws As Worksheet, ByVal request As CleanDataRequest) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanWorksheetWithOptionsDirect")
    On Error GoTo ErrHandler
    
    Set m_CurrentRequest = request
    
    Dim wsProcessor As IBatchProcessor
    Set wsProcessor = Me
    CleanWorksheetWithOptionsDirect = wsProcessor.ProcessWorksheet(ws)

CleanExit:
    Set m_CurrentRequest = Nothing
    Exit Function

ErrHandler:
    Set m_CurrentRequest = Nothing
    Infra_Error.HandleError "CleanWorksheetWithOptionsDirect", Err
    Resume CleanExit
End Function
