Attribute VB_Name = "Lib_Tests_Feat_CleanData"
Option Explicit

' @Module: Lib_Tests_Feat_CleanData
' @Category: Library
' @Description: Integration tests for data clean features.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, AppContainer, Infra_Error, Infra_Undo, Lib_ValueConversion, FeatCmd_CleanData, Infra_CleanDataRequest

Public Sub Test_CleanData_TrimmingAndNumericalFixing()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CleanData_TrimmingAndNumericalFixing")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CleanData"

    ' Setup test values
    ws.Range("A1").Value2 = "  hello   world  " ' Text with leading/trailing and double spaces
    ws.Range("A2").NumberFormat = "@"
    ws.Range("A2").Value2 = "123.45"              ' Number stored as text
    ws.Range("A3").Value2 = "normal text"         ' Normal text (should not change)
    
    ' Inject non-breaking space (Chr(160))
    ws.Range("A4").Value2 = "hello" & ChrW$(160) & "world"

    Dim cmd As New FeatCmd_CleanData
    Dim cleanedCount As Long
    cleanedCount = cmd.CleanRangeDirect(ws.Range("A1:A4"))

    ' Asserts
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "hello world", "CleanData should trim and remove extra spaces"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, "123.45", "CleanData should leave numeric text unchanged since text number conversion was removed"
    Lib_Tests.AssertEqual ws.Range("A2").NumberFormat, "@", "Number format should remain text"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, "normal text", "CleanData should leave normal text untouched"
    Lib_Tests.AssertEqual ws.Range("A4").Value2, "hello world", "CleanData should replace non-breaking spaces with standard spaces"

    ' Cleanup
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_CleanData_TrimmingAndNumericalFixing", Err
    Resume CleanExit
End Sub

Public Sub Test_CleanData_CheckboxOptions()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CleanData_CheckboxOptions")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CleanOpt"

    ' Setup test values
    ws.Range("A1").Value2 = "  hello   world  "
    ws.Range("A2").Value2 = "hello" & ChrW$(7) & "world"

    Dim request As New Infra_CleanDataRequest
    request.CleanTrimSpaces = False
    request.CleanNonPrintables = True

    Dim cmd As New FeatCmd_CleanData
    Dim cleanedCount As Long
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("A1:A2"), request)

    Lib_Tests.AssertEqual ws.Range("A1").Value2, "  hello   world  ", "CleanData should NOT trim spaces if CleanTrimSpaces is False"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, "helloworld", "CleanData should remove non-printable characters if CleanNonPrintables is True"

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_CleanData_CheckboxOptions", Err
    Resume CleanExit
End Sub

Public Sub Test_CleanData_NewEnhancements()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CleanData_NewEnhancements")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CleanEnh"
    Dim cleanedCount As Long

    ' 1. Test standardizing invisible characters (zero-width, BOM, Unicode spaces)
    ws.Range("A1").Value2 = "hello" & ChrW$(8203) & "world"
    ws.Range("A2").Value2 = ChrW$(65279) & "BOMtest"
    ws.Range("A3").Value2 = "thin" & ChrW$(8201) & "space"
    ws.Range("A4").Value2 = "narrow" & ChrW$(8239) & "space"

    Dim req As New Infra_CleanDataRequest
    req.CleanInvisibleChars = True
    req.CleanTrimSpaces = False
    req.CleanNonPrintables = False
    req.CleanConvertNumbers = False
    req.CleanBrokenNames = False

    Dim cmd As New FeatCmd_CleanData
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("A1:A4"), req)

    Lib_Tests.AssertEqual ws.Range("A1").Value2, "helloworld", "Should remove zero-width space"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, "BOMtest", "Should remove BOM"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, "thin space", "Should convert thin space to standard space"
    Lib_Tests.AssertEqual ws.Range("A4").Value2, "narrow space", "Should convert narrow space to standard space"

    ' 2. Test Line Breaks: Replace with space
    ws.Range("B1").Value2 = "line1" & vbCrLf & "line2" & vbLf & "line3"
    req.CleanInvisibleChars = False
    req.CleanReplaceLineBreaksWithSpace = True
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("B1"), req)
    Lib_Tests.AssertEqual ws.Range("B1").Value2, "line1 line2 line3", "Should replace line breaks with spaces"

    ' 3. Test Line Breaks: Remove entirely
    ws.Range("B2").Value2 = "line1" & vbCrLf & "line2" & vbLf & "line3"
    req.CleanReplaceLineBreaksWithSpace = False
    req.CleanRemoveLineBreaks = True
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("B2"), req)
    Lib_Tests.AssertEqual ws.Range("B2").Value2, "line1line2line3", "Should remove line breaks entirely"

    ' 4. Test Line Breaks: Standardize to single LF and protect from non-printables
    ws.Range("B3").Value2 = "line1" & vbCrLf & vbCrLf & "line2" & vbCr & vbCr & "line3"
    req.CleanRemoveLineBreaks = False
    req.CleanStandardizeLineBreaks = True
    req.CleanNonPrintables = True
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("B3"), req)
    Lib_Tests.AssertEqual ws.Range("B3").Value2, "line1" & vbLf & "line2" & vbLf & "line3", "Should standardize line breaks to single LF and protect them"

    ' 5. Test numeric text conversion
    ws.Range("C1").NumberFormat = "@"
    ws.Range("C1").Value2 = "123.45"
    ws.Range("C2").NumberFormat = "@"
    ws.Range("C2").Value2 = " &HFF "
    req.CleanStandardizeLineBreaks = False
    req.CleanNonPrintables = False
    req.CleanConvertNumbers = True
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("C1:C2"), req)
    
    Lib_Tests.AssertEqual ws.Range("C1").Value2, 123.45, "Should convert numeric text to number"
    Lib_Tests.AssertEqual VarType(ws.Range("C1").Value2), vbDouble, "Converted number should be double"
    Lib_Tests.AssertEqual ws.Range("C2").Value2, " &HFF ", "Should NOT convert hex strings to number"

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_CleanData_NewEnhancements", Err
    Resume CleanExit
End Sub

Public Sub Test_CleanData_UserRequestedEnhancements()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CleanData_UserRequestedEnhancements")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CleanUser"
    Dim cleanedCount As Long
    Dim cmd As New FeatCmd_CleanData

    ' 1. Test Remove Special Symbols (Tabs, Bullets, Trademark, Registered, Copyright)
    ws.Range("A1").Value2 = "hello" & Chr$(9) & "world"
    ws.Range("A2").Value2 = "Bullet" & ChrW$(8226) & "Text"
    ws.Range("A3").Value2 = "Company" & ChrW$(8482)
    ws.Range("A4").Value2 = "Brand" & ChrW$(174)
    ws.Range("A5").Value2 = "Copyright" & ChrW$(169) & "2026"

    Dim req As New Infra_CleanDataRequest
    req.CleanTrimSpaces = False
    req.CleanNonPrintables = False
    req.CleanInvisibleChars = False
    req.CleanConvertNumbers = False
    req.CleanBrokenNames = False
    req.CleanSpecialSymbols = True

    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("A1:A5"), req)

    Lib_Tests.AssertEqual ws.Range("A1").Value2, "helloworld", "Should remove tabs"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, "BulletText", "Should remove bullets"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, "Company", "Should remove trademark"
    Lib_Tests.AssertEqual ws.Range("A4").Value2, "Brand", "Should remove registered"
    Lib_Tests.AssertEqual ws.Range("A5").Value2, "Copyright2026", "Should remove copyright"

    ' 2. Test Standardize Dashes (convert en-dash, em-dash, unicode minus sign to standard hyphen)
    ws.Range("B1").Value2 = "a" & ChrW$(8211) & "b"
    ws.Range("B2").Value2 = "c" & ChrW$(8212) & "d"
    ws.Range("B3").Value2 = ChrW$(8722) & "15.2"

    req.CleanSpecialSymbols = False
    req.CleanStandardizeDashes = True
    req.CleanConvertNumbers = True

    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("B1:B3"), req)

    Lib_Tests.AssertEqual ws.Range("B1").Value2, "a-b", "Should convert en dash to hyphen"
    Lib_Tests.AssertEqual ws.Range("B2").Value2, "c-d", "Should convert em dash to hyphen"
    Lib_Tests.AssertEqual ws.Range("B3").Value2, -15.2, "Should convert minus sign to hyphen and parse as numeric"
    Lib_Tests.AssertEqual VarType(ws.Range("B3").Value2), vbDouble, "Minus sign converted numeric should be a Double"

    ' 3. Test Remove Accents (José -> Jose, François -> Francois, Müller -> Muller)
    ws.Range("C1").Value2 = "Jos" & ChrW$(233)
    ws.Range("C2").Value2 = "Fran" & ChrW$(231) & "ois"
    ws.Range("C3").Value2 = "M" & ChrW$(252) & "ller"
    ws.Range("C4").Value2 = "Stra" & ChrW$(223) & "e & " & ChrW$(198) & "ther & " & ChrW$(339) & "uf"

    req.CleanStandardizeDashes = False
    req.CleanConvertNumbers = False
    req.CleanRemoveAccents = True

    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("C1:C4"), req)

    Lib_Tests.AssertEqual ws.Range("C1").Value2, "Jose", "Should remove accent from Jose"
    Lib_Tests.AssertEqual ws.Range("C2").Value2, "Francois", "Should remove cedilla from Francois"
    Lib_Tests.AssertEqual ws.Range("C3").Value2, "Muller", "Should remove umlaut from Muller"
    Lib_Tests.AssertEqual ws.Range("C4").Value2, "Strasse & AEther & oeuf", "Should replace multi-character accents"

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_CleanData_UserRequestedEnhancements", Err
    Resume CleanExit
End Sub

Public Sub Test_CleanData_HygieneOptions()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CleanData_HygieneOptions")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CleanHygiene"
    
    Dim cmd As New FeatCmd_CleanData
    Dim req As New Infra_CleanDataRequest
    
    ' Disable text cleaning defaults to isolate hygiene tests
    req.CleanTrimSpaces = False
    req.CleanNonPrintables = False
    req.CleanInvisibleChars = False
    req.CleanBrokenNames = False

    ' 1. Test Clear Comments/Notes
    ws.Range("A1").Value2 = "Value A1"
    ws.Range("A1").AddComment "Test Comment A1"
    
    req.CleanComments = True
    Dim cleanedCount As Long
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("A1"), req)
    
    Lib_Tests.AssertEqual True, ws.Range("A1").Comment Is Nothing, "Comment should be removed"
    req.CleanComments = False
    
    ' 2. Test Clear Validation
    ws.Range("A2").Value2 = 5
    With ws.Range("A2").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="10"
    End With
    
    req.CleanValidation = True
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("A2"), req)
    
    Dim hasValidation As Boolean
    hasValidation = False
    On Error Resume Next
    Dim valType As Long
    valType = ws.Range("A2").Validation.Type
    If Err.Number = 0 Then hasValidation = True
    On Error GoTo ErrHandler
    
    Lib_Tests.AssertEqual False, hasValidation, "Validation rule should be deleted"
    req.CleanValidation = False

    ' 3. Test Clear Conditional Formatting
    ws.Range("A3").Value2 = 10
    ws.Range("A3").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="5"
    ws.Range("A3").FormatConditions(1).Interior.Color = vbRed
    
    req.CleanConditionalFormatting = True
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("A3"), req)
    Lib_Tests.AssertEqual 0, ws.Range("A3").FormatConditions.Count, "Conditional formatting should be deleted"
    req.CleanConditionalFormatting = False

    ' 4. Test Clear Cell Formatting (keep values)
    ws.Range("A4").Value2 = "Bold Text"
    ws.Range("A4").Font.Bold = True
    ws.Range("A4").Interior.Color = vbYellow
    
    req.CleanFormats = True
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("A4"), req)
    Lib_Tests.AssertEqual "Bold Text", ws.Range("A4").Value2, "Value should be kept"
    Lib_Tests.AssertEqual False, ws.Range("A4").Font.Bold, "Font bold formatting should be cleared"
    Lib_Tests.AssertEqual xlNone, ws.Range("A4").Interior.ColorIndex, "Fill color should be cleared"
    req.CleanFormats = False

    ' 5. Test Remove Shapes/Images
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(1, 10, 10, 50, 50) ' 1 = msoShapeRectangle
    sh.Name = "TestRectangle"
    
    req.CleanShapes = True
    req.Scope = TargetScopeSelection
    
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("A1:C10"), req)
    
    Dim shapeExists As Boolean
    shapeExists = False
    Dim testSh As Shape
    On Error Resume Next
    Set testSh = ws.Shapes("TestRectangle")
    If Not testSh Is Nothing Then shapeExists = True
    On Error GoTo ErrHandler
    Lib_Tests.AssertEqual False, shapeExists, "Shape in selection should be deleted"
    req.CleanShapes = False

    ' 6. Test Remove Sheet-scoped Named Ranges
    ws.Names.Add Name:="LocalName", RefersTo:="=$A$1"
    
    req.CleanSheetNames = True
    req.Scope = TargetScopeActiveSheet
    
    Dim procCount As Long
    procCount = cmd.CleanWorksheetWithOptionsDirect(ws, req)
    
    Dim nameExists As Boolean
    nameExists = False
    Dim nm As Name
    On Error Resume Next
    Set nm = ws.Names("LocalName")
    If Not nm Is Nothing Then nameExists = True
    On Error GoTo ErrHandler
    Lib_Tests.AssertEqual False, nameExists, "Sheet-scoped name should be deleted"
    req.CleanSheetNames = False

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_CleanData_HygieneOptions", Err
    Resume CleanExit
End Sub

Public Sub Test_CleanData_LargeSelectionSafety()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CleanData_LargeSelectionSafety")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ClnWhole"

    ' Write dirty text in C3
    ws.Range("C3").Value2 = "  dirty text  "

    Dim cmd As New FeatCmd_CleanData
    Dim request As New Infra_CleanDataRequest
    Set request.Context = AppContainer.CreateCommandContext("CleanData").ActionContext
    Set request.Context.WorksheetRef = ws
    Set request.Context.WorkbookRef = ThisWorkbook
    
    ' Run clean on the entire sheet
    ' It should intersect with UsedRange and only clean C3, running instantly
    Dim changes As Long
    changes = cmd.CleanRangeWithOptionsDirect(ws.Cells, request)

    Lib_Tests.AssertEqual ws.Range("C3").Value2, "dirty text", "C3 should be trimmed"
    Lib_Tests.AssertTrue changes >= 1, "Should report at least 1 change"

    ' Cleanup
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_CleanData_LargeSelectionSafety", Err
    Resume CleanExit
End Sub

Public Sub Test_ForceNumber_Execution_And_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ForceNumber_Execution_And_Undo")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim wb As Workbook: Set wb = Workbooks.Add
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)

    ' Setup text formatted numbers
    ws.Range("A1").NumberFormat = "@"
    ws.Range("A1").Value = "1250.50"
    
    ws.Range("A2").NumberFormat = "@"
    ws.Range("A2").Value = "$1500"
    
    ws.Range("A3").NumberFormat = "@"
    ws.Range("A3").Value = "150-"
    
    ws.Range("A4").NumberFormat = "@"
    ws.Range("A4").Value = "5%"
    
    ws.Range("A1:A4").Select
    
    ' Execute
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnForceNumber", "ForceNumber", "Ribbon"
    
    ' Check values and types
    Lib_Tests.AssertEqual VarType(ws.Range("A1").Value), vbDouble, "A1 is Double"
    Lib_Tests.AssertEqual ws.Range("A1").Value, 1250.5, "A1 value is 1250.5"
    Lib_Tests.AssertEqual ws.Range("A1").NumberFormat, "General", "A1 format is General"
    
    Lib_Tests.AssertEqual VarType(ws.Range("A2").Value), vbDouble, "A2 is Double"
    Lib_Tests.AssertEqual ws.Range("A2").Value, 1500#, "A2 value is 1500"
    
    Lib_Tests.AssertEqual VarType(ws.Range("A3").Value), vbDouble, "A3 is Double"
    Lib_Tests.AssertEqual ws.Range("A3").Value, -150#, "A3 value is -150"
    
    Lib_Tests.AssertEqual VarType(ws.Range("A4").Value), vbDouble, "A4 is Double"
    Lib_Tests.AssertEqual ws.Range("A4").Value, 0.05, "A4 value is 0.05"
    
    ' Test Undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    
    ' Check if format and values are restored
    Lib_Tests.AssertEqual ws.Range("A1").NumberFormat, "@", "A1 format is text again"
    Lib_Tests.AssertEqual ws.Range("A1").Value, "1250.50", "A1 value is text again"
    Lib_Tests.AssertEqual ws.Range("A4").Value, "5%", "A4 value is text again"
    
    wb.Close SaveChanges:=False
    
CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_ForceNumber_Execution_And_Undo", Err
    Resume CleanExit
End Sub

Public Sub Test_UnmergeFill_Execution_And_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_UnmergeFill_Execution_And_Undo")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim wb As Workbook: Set wb = Workbooks.Add
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)

    ' Setup merged range
    Dim targetRange As Range: Set targetRange = ws.Range("A1:B2")
    targetRange.Merge
    targetRange.Cells(1, 1).Value = "TestMerged"

    ' Run command via AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    
    targetRange.Select
    
    Dim cmd As ICommand: Set cmd = AppContainer.ResolveCommand("UnmergeFill")
    Lib_Tests.AssertEqual Not cmd Is Nothing, True, "Resolve UnmergeFill"
    
    ' Execute
    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnUnmergeFill", "UnmergeFill", "Ribbon"
    
    ' Check if unmerged and populated
    Lib_Tests.AssertEqual targetRange.MergeCells, False, "Should be unmerged"
    Lib_Tests.AssertEqual ws.Range("A1").Value, "TestMerged", "A1 has value"
    Lib_Tests.AssertEqual ws.Range("A2").Value, "TestMerged", "A2 has value"
    Lib_Tests.AssertEqual ws.Range("B1").Value, "TestMerged", "B1 has value"
    Lib_Tests.AssertEqual ws.Range("B2").Value, "TestMerged", "B2 has value"
    
    ' Test Undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    
    ' Check if merged again and restored
    Lib_Tests.AssertEqual targetRange.MergeCells, True, "Should be merged again after Undo"
    Lib_Tests.AssertEqual ws.Range("A1").Value, "TestMerged", "A1 still has value"
    
    wb.Close SaveChanges:=False
    
CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_UnmergeFill_Execution_And_Undo", Err
    Resume CleanExit
End Sub

Public Sub Test_UnifiedHelpers_And_CleanDataDisjoint()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_UnifiedHelpers_And_CleanDataDisjoint")
    On Error GoTo ErrHandler

    ' 1. Test GetExcelErrorText
    Lib_Tests.AssertEqual Infra_ValueConversion.GetExcelErrorText(CVErr(xlErrDiv0)), "#DIV/0!", "Should return #DIV/0! for xlErrDiv0"
    Lib_Tests.AssertEqual Infra_ValueConversion.GetExcelErrorText(CVErr(xlErrNA)), "#N/A", "Should return #N/A for xlErrNA"
    Lib_Tests.AssertEqual Infra_ValueConversion.GetExcelErrorText("Hello"), "Hello", "Should return string itself for non-error string"

    ' 2. Test GetSystemDateFormatPattern
    Dim sysPattern As String
    sysPattern = Infra_ValueConversion.GetSystemDateFormatPattern()
    Lib_Tests.AssertTrue Len(sysPattern) > 0, "System date format pattern should not be empty"

    ' 3. Test Disjoint Clean Data Optimization
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_DisjointClean"

    Dim i As Long
    Dim unionRange As Range
    
    ' Populate disjoint cells (120 cells, e.g. A1, A3, A5...)
    For i = 1 To 240 Step 2
        ws.Cells(i, 1).Value2 = "  text  "
        If unionRange Is Nothing Then
            Set unionRange = ws.Cells(i, 1)
        Else
            Set unionRange = Application.Union(unionRange, ws.Cells(i, 1))
        End If
    Next i
    
    Lib_Tests.AssertEqual unionRange.Areas.Count, 120, "Should have 120 disjoint areas"

    Dim request As New Infra_CleanDataRequest
    request.CleanTrimSpaces = True
    
    ' Execute cleaning via headless entry point
    Dim cleanCmd As New FeatCmd_CleanData
    Dim cleanedCount As Long
    cleanedCount = cleanCmd.CleanRangeWithOptionsDirect(unionRange, request)
    
    Lib_Tests.AssertEqual cleanedCount, 120, "Should have cleaned 120 cells"
    
    ' Verify values are trimmed
    For i = 1 To 240 Step 2
        Lib_Tests.AssertEqual ws.Cells(i, 1).Value2, "text", "Value at row " & i & " should be trimmed"
    Next i

    ' Cleanup temporary sheet
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_UnifiedHelpers_And_CleanDataDisjoint", Err
    Resume CleanExit
End Sub

Public Sub Test_TryConvertToNumber_Unification()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_TryConvertToNumber_Unification")
    On Error GoTo ErrHandler

    Dim outVal As Variant
    Dim success As Boolean

    ' 1. Standard number
    success = Infra_ValueConversion.TryConvertToNumber("123.45", outVal)
    Lib_Tests.AssertTrue success, "Should successfully convert standard numeric string"
    Lib_Tests.AssertEqual outVal, 123.45, "Should return 123.45"

    ' 2. Trailing minus
    success = Infra_ValueConversion.TryConvertToNumber("123.45-", outVal)
    Lib_Tests.AssertTrue success, "Should successfully convert trailing minus"
    Lib_Tests.AssertEqual outVal, -123.45, "Should return -123.45"

    ' 3. Percent
    success = Infra_ValueConversion.TryConvertToNumber("45%", outVal)
    Lib_Tests.AssertTrue success, "Should successfully convert percent string"
    Lib_Tests.AssertEqual outVal, 0.45, "Should return 0.45"

    ' 4. Currency and spaces
    success = Infra_ValueConversion.TryConvertToNumber(" $ 1,234.50 ", outVal)
    Lib_Tests.AssertTrue success, "Should successfully convert formatted currency string"
    Lib_Tests.AssertEqual outVal, 1234.5, "Should return 1234.5"

    ' 5. Hex/Octal exclusion
    success = Infra_ValueConversion.TryConvertToNumber("&HFF", outVal)
    Lib_Tests.AssertTrue Not success, "Should reject hexadecimal strings"

    ' 6. Non-numeric
    success = Infra_ValueConversion.TryConvertToNumber("hello", outVal)
    Lib_Tests.AssertTrue Not success, "Should reject non-numeric string"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_TryConvertToNumber_Unification", Err
    Resume CleanExit
End Sub
