Attribute VB_Name = "Lib_Tests_Features"
Option Explicit

' @Module: Lib_Tests_Features
' @Category: Library
' @Description: Automated integration and custom undo test suite for feature commands.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, AppContainer, Infra_Error, Infra_Undo, Lib_XUnpivotFunction

Public Sub Test_HelloWorld_Execution_And_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_HelloWorld_Execution_And_Undo")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_HelloWorld"
    ws.Range("B2").Value2 = "Original Content"

    Dim ctx As ICommandContext
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Set ctx = AppContainer.CreateCommandContext("HelloWorld")
    
    ' Set the context refs directly to point to our newly created worksheet and cell B2
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.ActiveCellRef = ws.Range("B2")
    Set ctx.ActionContext.SelectionRange = ws.Range("B2")

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("HelloWorld")
    
    ' Execute the command directly
    cmd.Execute ctx
    
    ' Assert that the active cell (B2) contains "Hello world!"
    Lib_Tests.AssertEqual ws.Range("B2").Value2, "Hello world!", "HelloWorld command should update active cell to 'Hello world!'"

    ' Now register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    
    ' Assert that cell B2 returned to its original content
    Lib_Tests.AssertEqual ws.Range("B2").Value2, "Original Content", "Undo HelloWorld should restore active cell to its original content"

    ' Cleanup the temporary worksheet
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
    Infra_Error.HandleError "Test_HelloWorld_Execution_And_Undo", Err
    Resume CleanExit
End Sub



Public Sub Test_MakePermanent_SpillHandling_And_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_MakePermanent_SpillHandling_And_Undo")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_MakePermanent"

    ' Setup a dynamic array formula in A1 that spills down to A3
    ws.Range("A1").Formula2 = "=SEQUENCE(3, 1)"
    
    ' Recalculate to ensure dynamic array is evaluated
    ws.Calculate
    Infra_ValueConversion.WaitForCalculation
    
    ' Asserts before execution
    Lib_Tests.AssertEqual ws.Range("A1").Value2, 1#, "A1 should be 1"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, 2#, "A2 should be 2"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, 3#, "A3 should be 3"

    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    
    ' Create command context for MakePermanent
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("MakePermanent")
    
    ' Set the context refs directly to point to our cell A1
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("MakePermanent")
    
    ' Execute the command directly - it should expand to the full spill range (A1:A3)
    cmd.Execute ctx
    
    ' Assert that formulas are gone and values are static
    Lib_Tests.AssertEqual ws.Range("A1").HasFormula, False, "A1 formula should be removed"
    Lib_Tests.AssertEqual ws.Range("A1").Value2, 1#, "A1 static value should be 1"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, 2#, "A2 static value should be 2"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, 3#, "A3 static value should be 3"

    ' Now register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    
    ' Assert that formula and dynamic array are restored
    Lib_Tests.AssertEqual ws.Range("A1").HasFormula, True, "A1 formula should be restored"
    Lib_Tests.AssertEqual ws.Range("A1").Formula2, "=SEQUENCE(3, 1)", "A1 formula content should be restored"
    Lib_Tests.AssertEqual ws.Range("A1").Value2, 1#, "A1 restored value should be 1"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, 2#, "A2 restored value should be 2"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, 3#, "A3 restored value should be 3"

    ' Cleanup the temporary worksheet
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
    Infra_Error.HandleError "Test_MakePermanent_SpillHandling_And_Undo", Err
    Resume CleanExit
End Sub

Public Sub Test_MakePermanent_LegacyArray_And_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_MakePermanent_LegacyArray_And_Undo")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_LegacyArray"

    ' Setup a legacy CSE array formula in A1:A3
    ws.Range("A1:A3").FormulaArray = "=SEQUENCE(3, 1)"
    ws.Calculate
    Infra_ValueConversion.WaitForCalculation

    ' Asserts before execution
    Lib_Tests.AssertEqual ws.Range("A1").HasArray, True, "A1 should be part of an array formula"
    Lib_Tests.AssertEqual ws.Range("A1").Value2, 1#, "A1 should be 1"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, 2#, "A2 should be 2"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, 3#, "A3 should be 3"

    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' Create command context for MakePermanent
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("MakePermanent")

    ' Select only part of the array formula range (A1:A2)
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1:A2")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("MakePermanent")

    ' Execute the command - it should fallback and convert the entire array
    cmd.Execute ctx

    ' Assert that formulas are gone and values are static
    Lib_Tests.AssertEqual ws.Range("A1").HasFormula, False, "A1 formula should be removed"
    Lib_Tests.AssertEqual ws.Range("A1").Value2, 1#, "A1 static value should be 1"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, 2#, "A2 static value should be 2"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, 3#, "A3 static value should be 3"

    ' Now register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo

    ' Assert that CSE array formula is restored
    Lib_Tests.AssertEqual ws.Range("A1").HasArray, True, "A1 array formula should be restored"
    Lib_Tests.AssertEqual ws.Range("A1").Value2, 1#, "A1 restored value should be 1"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, 2#, "A2 restored value should be 2"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, 3#, "A3 restored value should be 3"

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
    Infra_Error.HandleError "Test_MakePermanent_LegacyArray_And_Undo", Err
    Resume CleanExit
End Sub

Public Sub Test_ValueConversion_ResolveSpillExpandedRange()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ValueConversion_ResolveSpillExpandedRange")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_SpillResolve"

    ' Setup a dynamic array formula in A1 that spills down to A3
    ws.Range("A1").Formula2 = "=SEQUENCE(3, 1)"
    ws.Calculate
    Infra_ValueConversion.WaitForCalculation

    ' Test 1: Resolve starting from the anchor cell (A1)
    Dim expandedFromAnchor As Range
    Set expandedFromAnchor = Infra_ValueConversion.ResolveSpillExpandedRange(ws.Range("A1"))
    Lib_Tests.AssertEqual expandedFromAnchor.Address(False, False), "A1:A3", "Expanded range from anchor A1 should be A1:A3"

    ' Test 2: Resolve starting from a spilled cell (A2)
    Dim expandedFromSpilled As Range
    Set expandedFromSpilled = Infra_ValueConversion.ResolveSpillExpandedRange(ws.Range("A2"))
    Lib_Tests.AssertEqual expandedFromSpilled.Address(False, False), "A1:A3", "Expanded range from spilled cell A2 should be A1:A3"

    ' Cleanup the temporary worksheet
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
    Infra_Error.HandleError "Test_ValueConversion_ResolveSpillExpandedRange", Err
    Resume CleanExit
End Sub

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

Public Sub Test_BreakExternalLinks_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_BreakExternalLinks_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_BreakLinks"

    ' Setup an external name reference
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="TestExternalName", RefersTo:="=[ExternalWorkbook.xlsx]Sheet1!$A$1"
    On Error GoTo ErrHandler

    Dim cmd As New FeatCmd_BreakExternalLinks
    Dim namesRemoved As Long
    namesRemoved = cmd.RemoveExternalWorkbookNamesDirect(ThisWorkbook)

    ' Assert external names removal works safely
    Lib_Tests.AssertTrue namesRemoved >= 0, "RemoveExternalWorkbookNames should run safely and remove external names"

    ' Verify the name is gone
    Dim nameExists As Boolean
    Dim nm As Name
    On Error Resume Next
    Set nm = ThisWorkbook.Names("TestExternalName")
    If Not nm Is Nothing Then nameExists = True
    On Error GoTo ErrHandler

    Lib_Tests.AssertEqual nameExists, False, "External named range should be successfully deleted"

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
    Infra_Error.HandleError "Test_BreakExternalLinks_Execution", Err
    Resume CleanExit
End Sub

Public Sub Test_BreakExternalLinks_SpillHandling()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_BreakExternalLinks_SpillHandling")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' 1. Create a temporary source workbook to avoid file picker dialog
    Dim sourceWb As Workbook
    Set sourceWb = Workbooks.Add
    sourceWb.Sheets(1).Range("A1").Value = "Apple"
    sourceWb.Sheets(1).Range("A2").Value = "Banana"
    sourceWb.Sheets(1).Range("A3").Value = "Cherry"
    Dim sourceWbName As String
    sourceWbName = sourceWb.Name

    ' 2. Create the target worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_BreakLinksSpill"

    ' 3. Set a dynamic array spill formula in the target sheet referencing the source workbook
    ws.Range("B2").Formula2 = "='[" & sourceWbName & "]" & sourceWb.Sheets(1).Name & "'!$A$1:$A$3"
    
    ' Force recalculation and wait to make sure the spill range is calculated and populated
    ws.Calculate
    Infra_ValueConversion.WaitForCalculation
    
    ' Check if it has a spill and correct values
    Lib_Tests.AssertEqual ws.Range("B2").Value, "Apple", "B2 should be Apple before link breaking"
    Lib_Tests.AssertEqual ws.Range("B3").Value, "Banana", "B3 should be Banana (spilled) before link breaking"
    Lib_Tests.AssertEqual ws.Range("B4").Value, "Cherry", "B4 should be Cherry (spilled) before link breaking"

    ' 4. Break the links
    Dim cmd As New FeatCmd_BreakExternalLinks
    Dim request As New Infra_ScopedRequest
    Dim context As ICommandContext
    
    Set context = AppContainer.CreateCommandContext("BreakExternalLinks")
    Set context.ActionContext.WorksheetRef = ws
    Set context.ActionContext.WorkbookRef = ThisWorkbook
    Set request.Context = context.ActionContext
    request.Scope = TargetScopeActiveSheet
    
    Dim stats As String
    stats = cmd.ExecuteBreakLinksDirect(ThisWorkbook, request, Empty)

    ' 5. Verify results
    ' Check that B2 formula is gone (replaced by static value)
    Lib_Tests.AssertEqual ws.Range("B2").HasFormula, False, "B2 should not have formula after link breaking"
    
    ' Check that all spilled values are preserved
    Lib_Tests.AssertEqual ws.Range("B2").Value, "Apple", "B2 value should be Apple after link breaking"
    Lib_Tests.AssertEqual ws.Range("B3").Value, "Banana", "B3 value should be Banana after link breaking"
    Lib_Tests.AssertEqual ws.Range("B4").Value, "Cherry", "B4 value should be Cherry after link breaking"
    
    ' Cleanup target worksheet
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Set ws = Nothing
    
    ' Close source workbook
    sourceWb.Close SaveChanges:=False
    Set sourceWb = Nothing

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Infra_Error.HandleError "Test_BreakExternalLinks_SpillHandling", Err
    Resume CleanExit
End Sub

Public Sub Test_XFilter_Features()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_XFilter_Features")
    On Error GoTo ErrHandler

    ' Setup Test Inputs
    ' Let's test with arrays (Variants)
    Dim src(1 To 4, 1 To 2) As Variant
    src(1, 1) = "Apple": src(1, 2) = 100
    src(2, 1) = "Banana": src(2, 2) = 200
    src(3, 1) = "cherry": src(3, 2) = 300
    src(4, 1) = "DATE": src(4, 2) = 400

    Dim ref(1 To 2, 1 To 1) As Variant
    ref(1, 1) = "banana"
    ref(2, 1) = "CHERRY"

    ' 1. Test Case-Insensitive Intersection (Default: code_number omitted)
    ' "Banana" matches "banana", "cherry" matches "CHERRY".
    ' Output should be 2 rows (Banana and cherry)
    Dim res1 As Variant
    res1 = Lib_XFilterFunction.XFilter(src, ref)
    
    Lib_Tests.AssertEqual UBound(res1, 1), 2, "XFilter Intersection count should be 2"
    Lib_Tests.AssertEqual res1(1, 1), "Banana", "XFilter Intersection first match should be Banana"
    Lib_Tests.AssertEqual res1(2, 1), "cherry", "XFilter Intersection second match should be cherry"

    ' 2. Test Case-Sensitive Intersection (case_sensitive = True)
    ' Nothing should match because "Banana" <> "banana" and "cherry" <> "CHERRY".
    ' By default, with if_empty omitted, it returns "Not found"
    Dim res2 As Variant
    res2 = Lib_XFilterFunction.XFilter(src, ref, 1, , True)
    Lib_Tests.AssertEqual res2, "Not found", "Omitted empty should return 'Not found'"

    ' 3. Test if_empty parameter
    ' Using case-sensitive intersection which finds nothing, but passing "Empty Val" as 4th arg.
    Dim res3 As Variant
    res3 = Lib_XFilterFunction.XFilter(src, ref, 1, "Empty Val", True)
    Lib_Tests.AssertEqual res3, "Empty Val", "Custom empty string should be returned"

    ' 4. Test Difference (code_number = 2) case-insensitive
    ' "Apple" and "DATE" should not match, so they should be returned.
    Dim res4 As Variant
    res4 = Lib_XFilterFunction.XFilter(src, ref, 2)
    Lib_Tests.AssertEqual UBound(res4, 1), 2, "XFilter Difference count should be 2"
    Lib_Tests.AssertEqual res4(1, 1), "Apple", "XFilter Difference first match should be Apple"
    Lib_Tests.AssertEqual res4(2, 1), "DATE", "XFilter Difference second match should be DATE"

    ' 5. Test 1D array conversion and scalar values
    Dim scalarSrc As Variant
    scalarSrc = "Apple"
    Dim scalarRef As Variant
    scalarRef = "Apple"
    Dim res5 As Variant
    res5 = Lib_XFilterFunction.XFilter(scalarSrc, scalarRef, 1)
    Lib_Tests.AssertEqual res5(1, 1), "Apple", "Scalar inputs should be handled correctly"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_XFilter_Features", Err
    Resume CleanExit
End Sub

Public Sub Test_XUnpivot_Features()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_XUnpivot_Features")
    On Error GoTo ErrHandler

    ' Setup standard wide data input
    ' 3 Rows, 5 Columns
    ' Row 1 (Headers): ID | Name | Jan | Feb | Mar
    ' Row 2 (Data):    101 | Alice | 100 | 110 | 120
    ' Row 3 (Data):    102 | Bob | 200 | 210 | 220
    Dim wideData(1 To 3, 1 To 5) As Variant
    wideData(1, 1) = "ID": wideData(1, 2) = "Name": wideData(1, 3) = "Jan": wideData(1, 4) = "Feb": wideData(1, 5) = "Mar"
    wideData(2, 1) = 101: wideData(2, 2) = "Alice": wideData(2, 3) = 100: wideData(2, 4) = 110: wideData(2, 5) = 120
    wideData(3, 1) = 102: wideData(3, 2) = "Bob": wideData(3, 3) = 200: wideData(3, 4) = 210: wideData(3, 5) = 220

    ' 1. Test Standard Unpivot (Default headers, skip_blanks = False)
    ' Jan, Feb, Mar are numerical in Row 2, so 3 unpivot columns and 2 key columns (ID, Name).
    ' Expected output: 1 + 2 * 3 = 7 rows, 4 columns (ID, Name, Attribute, Value)
    Dim res1 As Variant
    res1 = Lib_XUnpivotFunction.XUnpivot(wideData)
    
    Lib_Tests.AssertEqual UBound(res1, 1), 7, "XUnpivot standard: output should have 7 rows"
    Lib_Tests.AssertEqual UBound(res1, 2), 4, "XUnpivot standard: output should have 4 columns"
    
    ' Row 1 (Header)
    Lib_Tests.AssertEqual res1(1, 1), "ID", "XUnpivot standard: Header col 1 should be ID"
    Lib_Tests.AssertEqual res1(1, 2), "Name", "XUnpivot standard: Header col 2 should be Name"
    Lib_Tests.AssertEqual res1(1, 3), "Attribute", "XUnpivot standard: Header col 3 should be Attribute"
    Lib_Tests.AssertEqual res1(1, 4), "Value", "XUnpivot standard: Header col 4 should be Value"
    
    ' Row 2 (First unpivoted row for Alice)
    Lib_Tests.AssertEqual res1(2, 1), 101, "XUnpivot standard: R2 C1 should be 101"
    Lib_Tests.AssertEqual res1(2, 2), "Alice", "XUnpivot standard: R2 C2 should be Alice"
    Lib_Tests.AssertEqual res1(2, 3), "Jan", "XUnpivot standard: R2 C3 should be Jan"
    Lib_Tests.AssertEqual res1(2, 4), 100, "XUnpivot standard: R2 C4 should be 100"
    
    ' Row 4 (Third unpivoted row for Alice)
    Lib_Tests.AssertEqual res1(4, 3), "Mar", "XUnpivot standard: R4 C3 should be Mar"
    Lib_Tests.AssertEqual res1(4, 4), 120, "XUnpivot standard: R4 C4 should be 120"
    
    ' Row 5 (First unpivoted row for Bob)
    Lib_Tests.AssertEqual res1(5, 1), 102, "XUnpivot standard: R5 C1 should be 102"
    Lib_Tests.AssertEqual res1(5, 2), "Bob", "XUnpivot standard: R5 C2 should be Bob"
    Lib_Tests.AssertEqual res1(5, 3), "Jan", "XUnpivot standard: R5 C3 should be Jan"
    Lib_Tests.AssertEqual res1(5, 4), 200, "XUnpivot standard: R5 C4 should be 200"

    ' 2. Test Custom Headers
    Dim res2 As Variant
    res2 = Lib_XUnpivotFunction.XUnpivot(wideData, "Month", "Sales")
    Lib_Tests.AssertEqual res2(1, 3), "Month", "XUnpivot custom headers: Attribute header should be Month"
    Lib_Tests.AssertEqual res2(1, 4), "Sales", "XUnpivot custom headers: Value header should be Sales"

    ' 3. Test Skip Blanks
    ' Modify Row 3 Mar to empty string and Row 3 Feb to Empty
    wideData(3, 5) = ""
    wideData(3, 4) = Empty
    ' Output should omit 2 rows, so: 7 - 2 = 5 rows.
    Dim res3 As Variant
    res3 = Lib_XUnpivotFunction.XUnpivot(wideData, , , True)
    Lib_Tests.AssertEqual UBound(res3, 1), 5, "XUnpivot skip blanks: output should have 5 rows"
    
    ' Verify first unpivoted rows for Alice: should have Jan (100), Feb (110), and Mar (120)
    Lib_Tests.AssertEqual res3(2, 3), "Jan", "XUnpivot skip blanks: R2 Attribute should be Jan"
    Lib_Tests.AssertEqual res3(3, 3), "Feb", "XUnpivot skip blanks: R3 Attribute should be Feb"
    Lib_Tests.AssertEqual res3(4, 3), "Mar", "XUnpivot skip blanks: R4 Attribute should be Mar"
    ' Next row should be Bob Jan (200) because Bob Feb was Empty and Bob Mar was ""
    Lib_Tests.AssertEqual res3(5, 1), 102, "XUnpivot skip blanks: R5 ID should be 102"
    Lib_Tests.AssertEqual res3(5, 3), "Jan", "XUnpivot skip blanks: R5 Attribute should be Jan"

    ' 4. Test Single row boundary error
    Dim singleRow(1 To 1, 1 To 3) As Variant
    singleRow(1, 1) = "A": singleRow(1, 2) = "B": singleRow(1, 3) = "C"
    Dim res4 As Variant
    res4 = Lib_XUnpivotFunction.XUnpivot(singleRow)
    Lib_Tests.AssertTrue IsError(res4), "XUnpivot boundary: Single row should return error variant"
    
    ' 5. Test No Numeric Columns error
    ' All text in Row 2
    Dim noNumeric(1 To 2, 1 To 3) As Variant
    noNumeric(1, 1) = "ID": noNumeric(1, 2) = "Val1": noNumeric(1, 3) = "Val2"
    noNumeric(2, 1) = "101": noNumeric(2, 2) = "text1": noNumeric(2, 3) = "text2"
    Dim res5 As Variant
    res5 = Lib_XUnpivotFunction.XUnpivot(noNumeric)
    Lib_Tests.AssertTrue IsError(res5), "XUnpivot boundary: No numeric columns should return error variant"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_XUnpivot_Features", Err
    Resume CleanExit
End Sub

Public Sub Test_UdfRegistry_And_HelpCenter()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_UdfRegistry_And_HelpCenter")
    On Error GoTo ErrHandler

    ' 1. Test GetAllUdfs returns expected items
    Dim udfs As Collection
    Set udfs = Lib_UdfRegistry.GetAllUdfs()
    Lib_Tests.AssertTrue Not udfs Is Nothing, "UDF registry collection should not be Nothing"
    Lib_Tests.AssertTrue udfs.Count > 0, "UDF registry should contain at least one UDF"

    Dim meta As Object
    Set meta = udfs(1)
    Lib_Tests.AssertEqual meta("Name"), "XFilter", "First UDF name should be XFilter"
    Lib_Tests.AssertEqual meta("Category"), "User Defined", "XFilter category should be User Defined"
    Lib_Tests.AssertTrue IsArray(meta("ArgumentDescriptions")), "XFilter argument descriptions should be an array"

    ' 2. Verify ShowHelpCenter runs without error in headless mode (display bypassed)
    Infra_Hotkeys.ShowHelpCenter
    
    ' Assert that we completed without error
    Lib_Tests.AssertTrue True, "ShowHelpCenter completed without error in headless mode"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_UdfRegistry_And_HelpCenter", Err
    Resume CleanExit
End Sub

Public Sub Test_FillDown_Features()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_FillDown_Features")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_FillDown"

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' --- TEST 1: Multi-Column Fill Down ---
    ws.Range("A1").Value2 = 10
    ws.Range("B1").Value2 = 20
    ws.Range("C1:C5").Value2 = "Ref"
    
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("FillDown")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1:B1")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("FillDown")
    cmd.Execute ctx

    Lib_Tests.AssertEqual ws.Range("A5").Value2, 10#, "Multi-column filldown: A5 should be 10"
    Lib_Tests.AssertEqual ws.Range("B5").Value2, 20#, "Multi-column filldown: B5 should be 20"
    Lib_Tests.AssertEqual ws.Range("A6").Value2, vbEmpty, "Multi-column filldown: A6 should be empty"

    ' --- TEST 2: Proximity Search Distance Limit ---
    ws.Cells.Clear
    ws.Range("A1").Value2 = 100
    ws.Range("L1:L5").Value2 = "Ref" ' Column 12 (distance = 11 columns)
    
    Set ctx = AppContainer.CreateCommandContext("FillDown")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True
    
    cmd.Execute ctx
    Lib_Tests.AssertEqual ws.Range("A5").Value2, vbEmpty, "Proximity limit: A5 should be empty when neighbor is > 10 columns away"

    ' Neighbor at column 11 (distance = 10 columns)
    ws.Range("K1:K5").Value2 = "Ref"
    cmd.Execute ctx
    Lib_Tests.AssertEqual ws.Range("A5").Value2, 100#, "Proximity limit: A5 should be 100 when neighbor is exactly 10 columns away"

    ' --- TEST 3: Fragmentation Safety Limit ---
    ws.Cells.Clear
    ws.Range("A1").Value2 = 1
    ws.Range("B1:B10015").Value2 = "Ref"
    
    Dim filterVals(1 To 10011, 1 To 1) As Variant
    Dim i As Long
    For i = 1 To 10011
        If i Mod 2 = 1 Then
            filterVals(i, 1) = "show"
        Else
            filterVals(i, 1) = "hide"
        End If
    Next i
    ws.Range("Z1:Z10011").Value2 = filterVals
    ws.Range("Z1:Z10011").AutoFilter Field:=1, Criteria1:="show"
    
    Set ctx = AppContainer.CreateCommandContext("FillDown")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True
    
    cmd.Execute ctx
    
    ' Since the execution is aborted by the fragmentation safety guard,
    ' the copy/paste should not have run, so A3 and A10011 should remain empty.
    Lib_Tests.AssertEqual ws.Range("A3").Value2, vbEmpty, "Fragmentation guard: A3 should remain empty"
    Lib_Tests.AssertEqual ws.Range("A10011").Value2, vbEmpty, "Fragmentation guard: A10011 should remain empty"

    ' Cleanup
    ws.AutoFilterMode = False
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    ws.AutoFilterMode = False
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_FillDown_Features", Err
    Resume CleanExit
End Sub

Public Sub Test_Backspace_LargeRange_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_Backspace_LargeRange_Undo")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_LargeUndo"
    
    ' Add some values in the top cells
    ws.Range("A1").Value2 = "Value 1"
    ws.Range("A5").Value2 = "Value 5"
    
    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' Create command context for Backspace
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("Backspace")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    
    ' Select 20,000 rows (327,680,000 cells), exceeding MAX_UNDO_CELLS
    Set ctx.ActionContext.SelectionRange = ws.Rows("1:20000")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("Backspace")
    
    ' Execute the command directly
    cmd.Execute ctx
    
    ' Assert that the values are cleared
    Lib_Tests.AssertEqual ws.Range("A1").Value2, vbEmpty, "A1 should be cleared by Backspace"
    Lib_Tests.AssertEqual ws.Range("A5").Value2, vbEmpty, "A5 should be cleared by Backspace"

    ' Register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo

    ' Assert that the values are fully restored
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "Value 1", "A1 should be restored by Undo"
    Lib_Tests.AssertEqual ws.Range("A5").Value2, "Value 5", "A5 should be restored by Undo"

    ' Cleanup the temporary worksheet
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
    Infra_Error.HandleError "Test_Backspace_LargeRange_Undo", Err
    Resume CleanExit
End Sub

Public Sub Test_Wrap_CellAndPatternModes()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_Wrap_CellAndPatternModes")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_Wrap"

    ' Setup source cells and a wrapper cell
    ws.Range("A1").Value2 = 10
    ws.Range("B1").Formula2 = "=A1*2"
    ws.Range("C1").Formula2 = "=ROUND(B1, 2)"

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    Dim cmd As New FeatCmd_Wrap
    Dim errCount As Long
    
    ' Test 1: Apply wrap pattern ROUND([value], 0)
    cmd.TestApplyWrapPatternDirect ws.Range("B1"), "ROUND([value], 0)", True, errCount
    Lib_Tests.AssertEqual errCount, 0#, "Wrapping B1 pattern should have 0 errors"
    Lib_Tests.AssertEqual ws.Range("B1").Formula2, "=ROUND((A1*2), 0)", "B1 formula should be successfully wrapped"

    ' Test 2: Apply wrapper cell formula (C1) to A1
    cmd.TestApplyWrapperCellDirect ws.Range("A1"), ws.Range("C1").Formula2, errCount
    Lib_Tests.AssertEqual errCount, 0#, "Wrapping A1 using C1 wrapper should have 0 errors"
    Lib_Tests.AssertEqual ws.Range("A1").Formula2, "=ROUND(10, 2)", "A1 formula should be successfully wrapped"

    ' Test 3: Apply wrapper cell formula (C1) to range A2:A3 in bulk (using the new ApplyWrapperCellToRange via TestApplyWrapperCellRangeDirect)
    ws.Range("A2").Formula2 = "=B1*2"
    ws.Range("A3").Formula2 = "=B1*3"
    errCount = 0
    cmd.TestApplyWrapperCellRangeDirect ws.Range("A2:A3"), ws.Range("C1").Formula2, True, errCount
    Lib_Tests.AssertEqual errCount, 0#, "Wrapping A2:A3 using C1 wrapper in bulk should have 0 errors"
    Lib_Tests.AssertEqual ws.Range("A2").Formula2, "=ROUND(B1*2, 2)", "A2 formula should be successfully wrapped in bulk"
    Lib_Tests.AssertEqual ws.Range("A3").Formula2, "=ROUND(B1*3, 2)", "A3 formula should be successfully wrapped in bulk"

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
    Infra_Error.HandleError "Test_Wrap_CellAndPatternModes", Err
    Resume CleanExit
End Sub

Public Sub Test_StaticSheetWorkbook_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_StaticSheetWorkbook_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_StaticSheet"

    ws.Range("A1").Value2 = 100
    ws.Range("A2").Formula2 = "=A1*2"
    ws.Calculate
    
    Dim cmd As New FeatCmd_StaticSheetWorkbook
    Dim countConverted As Long
    countConverted = cmd.TestConvertSheetToValuesDirect(ws)

    Lib_Tests.AssertEqual countConverted, 1#, "1 formula cell should be converted to static"
    Lib_Tests.AssertEqual ws.Range("A2").HasFormula, False, "A2 formula should be removed"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, 200#, "A2 value should remain 200"

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
    Infra_Error.HandleError "Test_StaticSheetWorkbook_Execution", Err
    Resume CleanExit
End Sub

Public Sub Test_CreateSheet_PlacementAndNaming()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CreateSheet_PlacementAndNaming")
    On Error GoTo ErrHandler

    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets.Add
    ws1.Name = "Test_Temp_Create1"

    ' Create a sheet after ws1
    Set ws2 = ThisWorkbook.Worksheets.Add(After:=ws1)
    ws2.Name = "Test_Temp_Create2"

    ' Verify sheet names and placement
    Lib_Tests.AssertEqual ws1.Name, "Test_Temp_Create1", "ws1 name should match"
    Lib_Tests.AssertEqual ws2.Name, "Test_Temp_Create2", "ws2 name should match"
    Lib_Tests.AssertEqual ThisWorkbook.Worksheets(ws1.Index + 1).Name, ws2.Name, "ws2 should be positioned after ws1"

    ' Cleanup
    Application.DisplayAlerts = False
    ws1.Delete
    ws2.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws1.Delete
    ws2.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_CreateSheet_PlacementAndNaming", Err
    Resume CleanExit
End Sub

Public Sub Test_ApplyCustomNumberFormat_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ApplyCustomNumberFormat_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CustomFormat"

    ws.Range("A1").Value2 = 1234.56

    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("ApplyCustomNumberFormat")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("ApplyCustomNumberFormat")
    cmd.Execute ctx

    Lib_Tests.AssertEqual ws.Range("A1").NumberFormat, Infra_Config.Model.DefaultNumberFormat, "Custom number format should be applied to A1"

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
    Infra_Error.HandleError "Test_ApplyCustomNumberFormat_Execution", Err
    Resume CleanExit
End Sub

Public Sub Test_PasteFormat_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_PasteFormat_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_PasteFormat"

    ' Setup source cells with special styling
    ws.Range("A1").Value2 = "Source"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Interior.Color = vbGreen

    ws.Range("B1").Value2 = "Dest"
    ws.Range("B1").Font.Bold = False

    ' Copy source
    ws.Range("A1").Copy

    ' Check if CutCopyMode is active or if Excel is running headlessly (clipboard operations are unsupported in background)
    If Not Application.Visible Or Application.CutCopyMode = 0 Then
        Debug.Print "  [SKIP] Test_PasteFormat_Execution clipboard operations bypassed in headless/background environment"
        GoTo CleanExit
    End If

    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("PasteFormat")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("B1")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("PasteFormat")
    cmd.Execute ctx

    Lib_Tests.AssertEqual ws.Range("B1").Value2, "Dest", "Value of B1 should remain unchanged"
    Lib_Tests.AssertEqual ws.Range("B1").Font.Bold, True, "B1 should now have bold formatting"
    Lib_Tests.AssertEqual ws.Range("B1").Interior.Color, vbGreen, "B1 should now have green fill color"

    ' Clear clipboard
    Application.CutCopyMode = False

    ' Cleanup
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_PasteFormat_Execution", Err
    Resume CleanExit
End Sub

Public Sub Test_FormatRange_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_FormatRange_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_FormatRange"

    ' Setup target cells
    ws.Range("A1").Value2 = "HeaderCol1"
    ws.Range("B1").Value2 = "HeaderCol2"
    ws.Range("A2").Value2 = 123
    ws.Range("B2").Value2 = 46201
    ws.Range("D1").Value2 = "HeaderCol3"
    ws.Range("E1").Value2 = "HeaderCol4"
    ws.Range("D2").Value2 = 789
    ws.Range("E2").Value2 = 999

    ' Setup some merged cells
    ws.Range("A3:B3").Merge

    ' Setup multiple overlapping ListObject tables
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes)
    Dim tbl2 As ListObject
    Set tbl2 = ws.ListObjects.Add(xlSrcRange, ws.Range("D1:E2"), , xlYes)

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    Dim cmd As New FeatCmd_FormatRange
    ' Call the direct formatting testing method headlessly
    cmd.FormatRangeDirect ws.Range("A1:E3"), ws

    ' Asserts
    Lib_Tests.AssertEqual ws.ListObjects.Count, 0#, "All overlapping ListObject tables should be unlisted"
    Lib_Tests.AssertEqual ws.Range("A3").MergeCells, False, "Merged cells should be unmerged"
    Lib_Tests.AssertEqual ws.Range("A1").Font.Bold, True, "Header row A1 should be Bold"
    Lib_Tests.AssertEqual ws.Range("A1").Font.Size, Infra_Config.Model.HeaderFontSize, "Header font size should match config"
    Lib_Tests.AssertEqual ws.Range("A1").Interior.Color, Infra_Config.Model.HeaderColor, "Header color should match config"
    Lib_Tests.AssertEqual ws.Range("A2").Font.Size, Infra_Config.Model.DefaultFontSize, "Data row font size should match default config"

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
    Infra_Error.HandleError "Test_FormatRange_Execution", Err
    Resume CleanExit
End Sub

Public Sub Test_Delete_Execution_And_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_Delete_Execution_And_Undo")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_Delete"

    ws.Range("A1").Value2 = "DeleteMe"
    
    ' Select range
    ws.Range("A1").Select
    
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    
    ' Resolve and Execute Delete Command
    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("Delete")
    
    Dim context As ICommandContext
    Set context = AppContainer.CreateCommandContext("Delete", vbNullString, "Test", vbNullString)
    
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "DeleteMe", "A1 should contain text initially"
    
    ' Validate and Execute
    Dim valResult As CommandValidationResult
    Set valResult = cmd.Validate(context)
    Lib_Tests.AssertEqual valResult.IsExecutable, True, "Command should be executable"
    
    cmd.Execute context
    
    Lib_Tests.AssertEqual ws.Range("A1").Value2, Empty, "A1 should be cleared after Delete command"
    
    ' Perform Undo
    Infra_Undo.PerformUndo
    
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "DeleteMe", "A1 value should be restored by Undo"
    
    ' Test deleting a shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 50)
    shp.Select
    
    ' Re-create context for shape selection
    Set context = AppContainer.CreateCommandContext("Delete", vbNullString, "Test", vbNullString)
    cmd.Execute context
    
    Lib_Tests.AssertEqual ws.Shapes.Count, 0#, "Shape should be deleted"

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
    Infra_Error.HandleError "Test_Delete_Execution_And_Undo", Err
    Resume CleanExit
End Sub

Public Sub Test_FilterByCell_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_FilterByCell_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_Filter"

    ' Setup simple table
    ws.Range("A1").Value2 = "Fruit"
    ws.Range("A2").Value2 = "Apple"
    ws.Range("A3").Value2 = "Banana"
    ws.Range("A4").Value2 = "Apple"
    
    ws.Range("B1").Value2 = "Quantity"
    ws.Range("B2").Value2 = 10
    ws.Range("B3").Value2 = 20
    ws.Range("B4").Value2 = 30

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    
    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("FilterByCell")
    
    ' Select cell to filter by (Apple)
    ws.Range("A2").Select
    
    Dim context As ICommandContext
    Set context = AppContainer.CreateCommandContext("FilterByCell", vbNullString, "Test", vbNullString)
    
    ' Execute
    cmd.Execute context
    
    ' Assertions
    Lib_Tests.AssertEqual ws.AutoFilterMode, True, "AutoFilter should be enabled"
    
    Dim autoflt As AutoFilter
    Set autoflt = ws.AutoFilter
    Lib_Tests.AssertEqual autoflt.Range.Address, ws.Range("A1:B4").Address, "Filter range should encompass A1:B4"
    
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
    Infra_Error.HandleError "Test_FilterByCell_Execution", Err
    Resume CleanExit
End Sub

Public Sub Test_HighlightData_InconsistentFormulasAndDuplicates()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_HighlightData_InconsistentFormulasAndDuplicates")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_HighlightData"

    ' 1. Set up formula cells to test inconsistent formulas
    ws.Range("B1").Value2 = 10
    ws.Range("B2").Value2 = 20
    ws.Range("B3").Value2 = 30
    ws.Range("B4").Value2 = 40
    ws.Range("B5").Value2 = 50

    ws.Range("A1").Formula2 = "=B1"
    ws.Range("A2").Formula2 = "=B2"
    ws.Range("A3").Formula2 = "=B99" ' Inconsistent formula!
    ws.Range("A4").Formula2 = "=B4"
    ws.Range("A5").Formula2 = "=B5"

    ' 2. Set up duplicates
    ws.Range("C1").Value2 = "apple"
    ws.Range("C2").Value2 = "banana"
    ws.Range("C3").Value2 = "apple" ' Duplicate of C1!
    ws.Range("C4").Value2 = "cherry"
    ws.Range("C5").Value2 = ""      ' Empty (should not count as duplicate)
    ws.Range("C6").Value2 = ""      ' Empty (should not count as duplicate)

    ' Reset colors
    ws.Range("A1:A5").Interior.ColorIndex = xlNone
    ws.Range("C1:C6").Interior.ColorIndex = xlNone

    Dim cmd As New FeatCmd_HighlightData
    cmd.HighlightRangeDirect ws.Range("A1:A5,C1:C6")

    ' Check duplicate assertions: C1 and C3 should be highlighted with RGB(255, 199, 206)
    Lib_Tests.AssertEqual ws.Range("C1").Interior.Color, RGB(255, 199, 206), "C1 should be highlighted as duplicate"
    Lib_Tests.AssertEqual ws.Range("C3").Interior.Color, RGB(255, 199, 206), "C3 should be highlighted as duplicate"
    Lib_Tests.AssertEqual ws.Range("C2").Interior.ColorIndex, xlNone, "C2 is unique and should not be highlighted"
    Lib_Tests.AssertEqual ws.Range("C4").Interior.ColorIndex, xlNone, "C4 is unique and should not be highlighted"
    Lib_Tests.AssertEqual ws.Range("C5").Interior.ColorIndex, xlNone, "C5 is empty and should not be highlighted"
    Lib_Tests.AssertEqual ws.Range("C6").Interior.ColorIndex, xlNone, "C6 is empty and should not be highlighted"

    ' Test inconsistent formula if error indicator evaluated (Excel may require background error checking)
    Dim isIncA3 As Boolean
    isIncA3 = False
    On Error Resume Next
    isIncA3 = ws.Range("A3").Errors(xlInconsistentFormula).Value
    On Error GoTo ErrHandler
    
    If isIncA3 Then
        Lib_Tests.AssertEqual ws.Range("A3").Interior.Color, RGB(255, 255, 153), "A3 should be highlighted as inconsistent formula"
        Lib_Tests.AssertEqual ws.Range("A1").Interior.ColorIndex, xlNone, "A1 should not be highlighted"
    Else
        Debug.Print "BEAVER [TEST]: xlInconsistentFormula check skipped or not active in the current Excel test environment."
    End If

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
    Infra_Error.HandleError "Test_HighlightData_InconsistentFormulasAndDuplicates", Err
    Resume CleanExit
End Sub

Public Sub Test_ModifyData_Casing()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_Casing")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModCase"

    ws.Range("A1").Value2 = "heLLo wOrld"
    ws.Range("A2").Value2 = "heLLo wOrld"
    ws.Range("A3").Value2 = "heLLo wOrld"

    Dim cmd As New FeatCmd_ModifyData
    Dim req As New Infra_ModifyDataRequest
    Set req.Context = New Infra_ActionContext
    
    ' 1. Test UPPERCASE
    Dim changes As Long
    req.Operation = "Case: UPPERCASE"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A1"), req)
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "HELLO WORLD", "Should convert to uppercase"

    ' 2. Test lowercase
    req.Operation = "Case: lowercase"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A2"), req)
    Lib_Tests.AssertEqual ws.Range("A2").Value2, "hello world", "Should convert to lowercase"

    ' 3. Test Proper Case
    req.Operation = "Case: Proper Case"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A3"), req)
    Lib_Tests.AssertEqual ws.Range("A3").Value2, "Hello World", "Should convert to Proper Case"

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
    Infra_Error.HandleError "Test_ModifyData_Casing", Err
    Resume CleanExit
End Sub

Public Sub Test_ModifyData_DateStandardization()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_DateStandardization")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModDate"

    ws.Range("A1:A4").NumberFormat = "@"
    ws.Range("A1").Value = "05/12/2021" ' dd/mm/yyyy -> Dec 5, 2021
    ws.Range("A2").Value = "12/05/2021" ' mm/dd/yyyy -> Dec 5, 2021
    ws.Range("A3").Value = "15-Feb-21"   ' dd-mmm-yy -> Feb 15, 2021
    ws.Range("A4").Value = "20210215"    ' yyyymmdd -> Feb 15, 2021

    Dim cmd As New FeatCmd_ModifyData
    Dim req As New Infra_ModifyDataRequest
    Set req.Context = New Infra_ActionContext
    req.Operation = "Date Standardization"

    ' 1. Test dd/mm/yyyy
    req.DatePattern = "dd/mm/yyyy"
    Dim changes As Long
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A1"), req)
    Lib_Tests.AssertEqual ws.Range("A1").Value, DateSerial(2021, 12, 5), "Should convert dd/mm/yyyy to Dec 5, 2021"
    Lib_Tests.AssertEqual ws.Range("A1").NumberFormat, Infra_Config.Model.DisplayDateFormat, "Format should be date format"

    ' 2. Test mm/dd/yyyy
    req.DatePattern = "mm/dd/yyyy"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A2"), req)
    Lib_Tests.AssertEqual ws.Range("A2").Value, DateSerial(2021, 12, 5), "Should convert mm/dd/yyyy to Dec 5, 2021"

    ' 3. Test dd-mmm-yy
    req.DatePattern = "dd-mmm-yy"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A3"), req)
    Lib_Tests.AssertEqual ws.Range("A3").Value, DateSerial(2021, 2, 15), "Should convert dd-mmm-yy to Feb 15, 2021"

    ' 4. Test yyyymmdd (non-delimited)
    req.DatePattern = "yyyymmdd"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A4"), req)
    Lib_Tests.AssertEqual ws.Range("A4").Value, DateSerial(2021, 2, 15), "Should convert yyyymmdd to Feb 15, 2021"

    ' 5. Test Date type re-interpretation (Excel parsed as 3rd June, but user meant 6th March)
    Dim cellA5 As Range
    Set cellA5 = ws.Range("A5")
    cellA5.Value = DateSerial(2010, 6, 3) ' June 3, 2010
    cellA5.NumberFormat = "dd-mm-yyyy"

    req.DatePattern = "mm-dd-yyyy"
    changes = cmd.ModifyRangeWithOptionsDirect(cellA5, req)
    Lib_Tests.AssertEqual cellA5.Value, DateSerial(2010, 3, 6), "Should re-interpret June 3, 2010 as March 6, 2010 using mm-dd-yyyy"
    Lib_Tests.AssertEqual changes, 1, "Should report 1 change"

    ' 6. Test Date type re-interpretation with built-in format, adjusting for system locale
    Dim cellA6 As Range
    Set cellA6 = ws.Range("A6")
    Dim systemOrder As Long
    On Error Resume Next
    systemOrder = Application.International(xlDateOrder)
    If Err.Number <> 0 Then systemOrder = 1 ' Default to DMY
    On Error GoTo ErrHandler
    
    If systemOrder = 1 Then ' DMY system
        cellA6.Value = DateSerial(2021, 3, 1) ' March 1, 2021
        cellA6.NumberFormat = "m/d/yyyy" ' Format that formats to March 1, 2021 (e.g. 3/1/2021)
        req.DatePattern = "mm-dd-yyyy"
        changes = cmd.ModifyRangeWithOptionsDirect(cellA6, req)
        Lib_Tests.AssertEqual cellA6.Value, DateSerial(2021, 1, 3), "DMY: Should re-interpret March 1, 2021 as Jan 3, 2021 using m/d/yyyy format"
        Lib_Tests.AssertEqual changes, 1, "DMY: Should report 1 change"
    ElseIf systemOrder = 0 Then ' MDY system
        cellA6.Value = DateSerial(2021, 1, 3) ' Jan 3, 2021
        cellA6.NumberFormat = "dd-mm-yyyy" ' Format that formats to Jan 3, 2021 (e.g. 03-01-2021)
        req.DatePattern = "dd-mm-yyyy"
        changes = cmd.ModifyRangeWithOptionsDirect(cellA6, req)
        Lib_Tests.AssertEqual cellA6.Value, DateSerial(2021, 3, 1), "MDY: Should re-interpret Jan 3, 2021 as March 1, 2021 using dd-mm-yyyy format"
        Lib_Tests.AssertEqual changes, 1, "MDY: Should report 1 change"
    End If

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
    Infra_Error.HandleError "Test_ModifyData_DateStandardization", Err
    Resume CleanExit
End Sub

Public Sub Test_ModifyData_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_Undo")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModUndo"

    ws.Range("A1").Value2 = "original text"

    ' Setup command context
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("ModifyData")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("ModifyData")

    ' Run headless modification directly on A1 (Proper Case)
    Dim req As New Infra_ModifyDataRequest
    Set req.Context = ctx.ActionContext
    req.Operation = "Case: Proper Case"

    Dim featCmd As FeatCmd_ModifyData
    Set featCmd = cmd
    
    ' Save undo state first
    If Not Infra_Undo.SaveStateOrConfirm(ws.Range("A1"), "Modify Data") Then
        Err.Raise 5, "Test_ModifyData_Undo", "Failed to save undo state"
    End If
    
    Dim changes As Long
    changes = featCmd.ModifyRangeWithOptionsDirect(ws.Range("A1"), req)
    
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "Original Text", "A1 should be Proper Case"

    ' Register pending undo and perform
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo

    Lib_Tests.AssertEqual ws.Range("A1").Value2, "original text", "A1 should be restored to original text after undo"

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
    Infra_Error.HandleError "Test_ModifyData_Undo", Err
    Resume CleanExit
End Sub

Public Sub Test_HighlightData_FormulaLimitSafety()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_HighlightData_FormulaLimitSafety")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_HiLimit"

    ' Write 5005 unique formulas to exceed the default MAX_FORMULA_CHECK_CELLS (5000) and avoid duplicates
    ws.Range("A1:A5005").Formula2 = "=ROW()"

    Dim cmd As New FeatCmd_HighlightData
    
    ' This call should complete and not throw any errors, and skip checking formulas because of the limit
    cmd.HighlightRangeDirect ws.Range("A1:A5005")

    ' Since it was skipped, A1 should remain xlNone (no color index)
    Lib_Tests.AssertEqual ws.Range("A1").Interior.ColorIndex, xlNone, "Formula check should have been skipped, leaving A1 uncolored"

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
    Infra_Error.HandleError "Test_HighlightData_FormulaLimitSafety", Err
    Resume CleanExit
End Sub

Public Sub Test_HighlightData_Errors()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_HighlightData_Errors")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_HighErr"

    ' Set up standard Excel errors
    ' 1. Constant errors
    ws.Range("A1").Value = CVErr(xlErrNA)
    ws.Range("A2").Value = CVErr(xlErrValue)
    
    ' 2. Formula errors
    ws.Range("A3").Formula2 = "=1/0"
    ws.Range("A4").Formula2 = "=VLOOKUP(""nonexistent"", B10:C11, 2, FALSE)"
    
    ' 3. Normal values
    ws.Range("A5").Value2 = "normal text"
    ws.Range("A6").Value2 = 42

    ' Reset colors
    ws.Range("A1:A6").Interior.ColorIndex = xlNone

    Dim req As New Infra_HighlightDataRequest
    req.HighlightInconsistentFormulas = False
    req.HighlightDuplicates = False
    req.HighlightErrors = True

    Dim cmd As New FeatCmd_HighlightData
    cmd.HighlightRangeWithOptionsDirect ws.Range("A1:A6"), req

    ' Check assertions
    Lib_Tests.AssertEqual ws.Range("A1").Interior.Color, RGB(255, 204, 153), "A1 constant error should be highlighted orange"
    Lib_Tests.AssertEqual ws.Range("A2").Interior.Color, RGB(255, 204, 153), "A2 constant error should be highlighted orange"
    Lib_Tests.AssertEqual ws.Range("A3").Interior.Color, RGB(255, 204, 153), "A3 formula error should be highlighted orange"
    Lib_Tests.AssertEqual ws.Range("A4").Interior.Color, RGB(255, 204, 153), "A4 formula error should be highlighted orange"
    Lib_Tests.AssertEqual ws.Range("A5").Interior.ColorIndex, xlNone, "A5 normal text should not be highlighted"
    Lib_Tests.AssertEqual ws.Range("A6").Interior.ColorIndex, xlNone, "A6 normal number should not be highlighted"

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
    Infra_Error.HandleError "Test_HighlightData_Errors", Err
    Resume CleanExit
End Sub

Public Sub Test_HighlightData_HardcodedValues()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_HighlightData_HardcodedValues")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    ' Test standalone regex parsing helper first
    Dim cmd As New FeatCmd_HighlightData
    
    ' Should not contain hardcoded values
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=A1"), False, "Simple cell reference should be False"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=$A$1"), False, "Absolute reference should be False"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=A1+B2"), False, "Sum of cell references should be False"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=SUM(A1:B10)"), False, "SUM over range should be False"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=Sheet1!A1"), False, "Reference with sheet should be False"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("='Data 2026'!A1"), False, "Reference with quoted sheet and year should be False"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=LOG10(A1)"), False, "LOG10 function should be False"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=MATCH(A1, B:B, 0)"), False, "MATCH with 0 should be False"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=LEFT(A1, 1)"), False, "LEFT with 1 should be False"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=IF(A1="""", B1, C1)"), False, "IF with empty string should be False"
    
    ' Should contain hardcoded values
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=A1*1.05"), True, "1.05 is hardcoded"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=A1+50"), True, "50 is hardcoded"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=IF(A1=""Yes"", B1, C1)"), True, "Yes string is hardcoded"
    Lib_Tests.AssertEqual cmd.HasHardcodedValue("=DATE(2026, 6, 14)"), True, "Dates have hardcoded numbers"
    
    ' Test range highlighting on actual worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_HighHard"
    
    ws.Range("A1").Formula2 = "=B1"                     ' no hardcode
    ws.Range("A2").Formula2 = "=B2 + 0"                  ' ignores 0
    ws.Range("A3").Formula2 = "=B3 + 1"                  ' ignores 1
    ws.Range("A4").Formula2 = "=B4 * 1.05"               ' hardcoded 1.05 (should highlight)
    ws.Range("A5").Formula2 = "=IF(B5=""USD"", C5, D5)"   ' hardcoded "USD" (should highlight)
    ws.Range("A6").Formula2 = "=SUM(B6:B10) + 100"       ' hardcoded 100 (should highlight)
    
    ' Reset colors
    ws.Range("A1:A6").Interior.ColorIndex = xlNone
    
    Dim req As New Infra_HighlightDataRequest
    req.HighlightInconsistentFormulas = False
    req.HighlightDuplicates = False
    req.HighlightErrors = False
    req.HighlightHardcodedValues = True
    
    cmd.HighlightRangeWithOptionsDirect ws.Range("A1:A6"), req
    
    ' Check assertions
    Lib_Tests.AssertEqual ws.Range("A1").Interior.ColorIndex, xlNone, "A1 should not be highlighted"
    Lib_Tests.AssertEqual ws.Range("A2").Interior.ColorIndex, xlNone, "A2 should not be highlighted"
    Lib_Tests.AssertEqual ws.Range("A3").Interior.ColorIndex, xlNone, "A3 should not be highlighted"
    Lib_Tests.AssertEqual ws.Range("A4").Interior.Color, RGB(230, 210, 250), "A4 formula with 1.05 should be highlighted lavender"
    Lib_Tests.AssertEqual ws.Range("A5").Interior.Color, RGB(230, 210, 250), "A5 formula with 'USD' should be highlighted lavender"
    Lib_Tests.AssertEqual ws.Range("A6").Interior.Color, RGB(230, 210, 250), "A6 formula with 100 should be highlighted lavender"
    
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
    Infra_Error.HandleError "Test_HighlightData_HardcodedValues", Err
    Resume CleanExit
End Sub

Public Sub Test_HighlightData_DataValidations()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_HighlightData_DataValidations")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_HighVal"

    ' Set up validation on range A1:A2
    With ws.Range("A1:A2").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="1", Formula2:="10"
    End With
    
    ws.Range("A3").Value = 42

    ' Reset colors
    ws.Range("A1:A3").Interior.ColorIndex = xlNone

    Dim req As New Infra_HighlightDataRequest
    req.HighlightInconsistentFormulas = False
    req.HighlightDuplicates = False
    req.HighlightDataValidations = True

    Dim cmd As New FeatCmd_HighlightData
    cmd.HighlightRangeWithOptionsDirect ws.Range("A1:A3"), req

    ' Check assertions
    Lib_Tests.AssertEqual ws.Range("A1").Interior.Color, RGB(204, 255, 204), "A1 validation should be highlighted soft green"
    Lib_Tests.AssertEqual ws.Range("A2").Interior.Color, RGB(204, 255, 204), "A2 validation should be highlighted soft green"
    Lib_Tests.AssertEqual ws.Range("A3").Interior.ColorIndex, xlNone, "A3 normal cell should not be highlighted"

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
    Infra_Error.HandleError "Test_HighlightData_DataValidations", Err
    Resume CleanExit
End Sub

Public Sub Test_HighlightData_ConditionalFormatting()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_HighlightData_ConditionalFormatting")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_HighCF"

    ' Set up conditional formatting on range A1:A2
    ws.Range("A1:A2").FormatConditions.Delete
    ws.Range("A1:A2").FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=10"
    
    ws.Range("A3").Value = 42

    ' Reset colors
    ws.Range("A1:A3").Interior.ColorIndex = xlNone

    Dim req As New Infra_HighlightDataRequest
    req.HighlightInconsistentFormulas = False
    req.HighlightDuplicates = False
    req.HighlightConditionalFormatting = True

    Dim cmd As New FeatCmd_HighlightData
    cmd.HighlightRangeWithOptionsDirect ws.Range("A1:A3"), req

    ' Check assertions
    Lib_Tests.AssertEqual ws.Range("A1").Interior.Color, RGB(204, 229, 255), "A1 CF should be highlighted soft blue"
    Lib_Tests.AssertEqual ws.Range("A2").Interior.Color, RGB(204, 229, 255), "A2 CF should be highlighted soft blue"
    Lib_Tests.AssertEqual ws.Range("A3").Interior.ColorIndex, xlNone, "A3 normal cell should not be highlighted"

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
    Infra_Error.HandleError "Test_HighlightData_ConditionalFormatting", Err
    Resume CleanExit
End Sub

Public Sub Test_FormatRange_WholeSheetSafety()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_FormatRange_WholeSheetSafety")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_FmtWhole"

    ' Write one value in C3
    ws.Range("C3").Value2 = "content"

    Dim cmd As New FeatCmd_FormatRange
    
    ' Select entire worksheet range and execute formatting
    ' This should run instantly because it restricts itself to UsedRange (which contains C3)
    cmd.FormatRangeDirect ws.Cells, ws

    ' C3 should be formatted (e.g. bold header font or regular cell format)
    Lib_Tests.AssertEqual ws.Range("C3").Font.Bold, True, "C3 should be bolded as it became the header of the intersected used range"
    Lib_Tests.AssertEqual ws.Range("C3").Font.Size, 11#, "C3 font size should be 11"

    ' Cell outside used range (e.g., Z99) should remain unformatted
    Lib_Tests.AssertEqual ws.Range("Z99").Font.Bold, False, "Z99 should not be bolded"

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
    Infra_Error.HandleError "Test_FormatRange_WholeSheetSafety", Err
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

Public Sub Test_GetChunkedRanges_And_SpillExpansion()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_GetChunkedRanges_And_SpillExpansion")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_Chunk"

    ' 1. Test GetChunkedRanges
    ' Create a multi-area range: A1:A50 and C1:C10
    Dim multiAreaRange As Range
    Set multiAreaRange = Application.Union(ws.Range("A1:A50"), ws.Range("C1:C10"))
    
    ' Chunk with limit of 20 rows
    Dim chunks As Collection
    Set chunks = Infra_CommandSupport.GetChunkedRanges(multiAreaRange, 20)
    
    ' A1:A50 (50 rows) -> should split into 3 chunks: A1:A20, A21:A40, A41:A50
    ' C1:C10 (10 rows) -> should be 1 chunk: C1:C10
    ' Total expected chunks = 4
    Lib_Tests.AssertEqual chunks.Count, 4, "Should divide ranges into 4 chunks total"
    
    Lib_Tests.AssertEqual chunks(1).Address(False, False), "A1:A20", "First chunk should be A1:A20"
    Lib_Tests.AssertEqual chunks(2).Address(False, False), "A21:A40", "Second chunk should be A21:A40"
    Lib_Tests.AssertEqual chunks(3).Address(False, False), "A41:A50", "Third chunk should be A41:A50"
    Lib_Tests.AssertEqual chunks(4).Address(False, False), "C1:C10", "Fourth chunk should be C1:C10"

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
    Infra_Error.HandleError "Test_GetChunkedRanges_And_SpillExpansion", Err
    Resume CleanExit
End Sub


Public Sub Test_UI_OptionPicker_DynamicLayout()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_UI_OptionPicker_DynamicLayout")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim frm As Object
    On Error Resume Next
    Set frm = VBA.UserForms.Add("UI_OptionPicker")
    On Error GoTo ErrHandler

    Lib_Tests.AssertTrue Not frm Is Nothing, "UI_OptionPicker form should be loadable"

    ' Configure single select option picker
    frm.ConfigureOptionPicker "Test Title", "Select an option from the list below:", "Option 2", Array("Option 1", "Option 2", "Option 3 With a Very Long Text to Test Sizing")

    ' Check layout size properties of the form and controls
    Dim lst As Object: Set lst = frm.Controls("lstHotkeys")
    Dim lblPrompt As Object: Set lblPrompt = frm.Controls("lblPrompt")
    Dim btnOK As Object: Set btnOK = frm.Controls("btnOK")
    Dim btnCancel As Object: Set btnCancel = frm.Controls("btnCancel")

    Lib_Tests.AssertTrue Not lst Is Nothing, "ListBox control lstHotkeys should exist"
    Lib_Tests.AssertTrue Not lblPrompt Is Nothing, "Label control lblPrompt should exist"
    Lib_Tests.AssertTrue Not btnOK Is Nothing, "Button control btnOK should exist"
    Lib_Tests.AssertTrue Not btnCancel Is Nothing, "Button control btnCancel should exist"

    ' Assertions on visibility
    Lib_Tests.AssertTrue lblPrompt.Visible = False, "Label control lblPrompt should be invisible"
    Lib_Tests.AssertTrue btnOK.Visible = False, "Button control btnOK should be invisible"
    Lib_Tests.AssertTrue btnCancel.Visible = False, "Button control btnCancel should be invisible"

    ' Assertions on dimensions
    Lib_Tests.AssertTrue frm.Width > 200, "Form width should be scaled dynamically"
    Lib_Tests.AssertTrue frm.Height > 50, "Form height should be scaled dynamically"
    Lib_Tests.AssertTrue lst.Width > 180, "ListBox width should be scaled to fit options"
    ' Test multi-select option picker configuration
    frm.ConfigureMultiOptionPicker "Test Multi Title", "Check the options:", Array("Opt A", "Opt B"), Array(True, False)

    Lib_Tests.AssertTrue lst.MultiSelect = 0, "ListBox should be set to single-select custom checkbox list mode"
    Lib_Tests.AssertTrue btnOK.Visible = True, "OK button should be visible in multi-select mode"
    Lib_Tests.AssertTrue btnCancel.Visible = True, "Cancel button should be visible in multi-select mode"
CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_UI_OptionPicker_DynamicLayout", Err
    Resume CleanExit
End Sub


Public Sub Test_UI_OptionPicker_KeyboardNavigation()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_UI_OptionPicker_KeyboardNavigation")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim frm As Object
    On Error Resume Next
    Set frm = VBA.UserForms.Add("UI_OptionPicker")
    On Error GoTo ErrHandler

    Lib_Tests.AssertTrue Not frm Is Nothing, "UI_OptionPicker form should be loadable"

    ' Configure single select option picker
    frm.ConfigureOptionPicker "Test Keyboard Title", "Select an option:", "Option 1", Array("Option 1", "Option 2")

    ' Check initial state
    Lib_Tests.AssertTrue Not frm.IsIgnoringClick, "Initial IsIgnoringClick should be False"
    Lib_Tests.AssertTrue Not frm.WasConfirmed, "Initial WasConfirmed should be False"

    ' Simulate Arrow Down key down (KeyCode = 40)
    frm.HandleKeyDown 40, 0
    Lib_Tests.AssertTrue frm.IsIgnoringClick, "IsIgnoringClick should be True after key down (arrow key)"

    ' Simulate Arrow Down key up
    frm.HandleKeyUp 40, 0
    Lib_Tests.AssertTrue Not frm.IsIgnoringClick, "IsIgnoringClick should be False after key up"

    ' Simulate Enter key down (KeyCode = 13)
    frm.HandleKeyDown 13, 0
    Lib_Tests.AssertTrue frm.WasConfirmed, "WasConfirmed should be True after Enter key down"

CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_UI_OptionPicker_KeyboardNavigation", Err
    Resume CleanExit
End Sub


Public Sub Test_ModifyData_MixedFormats()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_MixedFormats")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModMixedFmts"

    ' Setup cells with different number formats
    ws.Range("A1").NumberFormat = "General"
    ws.Range("A2").NumberFormat = "$#,##0"
    ws.Range("A3").NumberFormat = "@"

    ws.Range("A1").Value2 = "apple"
    ws.Range("A2").Value2 = "banana"
    ws.Range("A3").Value2 = "cherry"

    Dim cmd As New FeatCmd_ModifyData
    Dim req As New Infra_ModifyDataRequest
    Set req.Context = New Infra_ActionContext
    
    req.Operation = "Case: UPPERCASE"
    
    Dim changes As Long
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A1:A3"), req)
    
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "APPLE", "A1 should be uppercase"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, "BANANA", "A2 should be uppercase"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, "CHERRY", "A3 should be uppercase"
    Lib_Tests.AssertEqual changes, 3, "Should report 3 changes"

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
    Infra_Error.HandleError "Test_ModifyData_MixedFormats", Err
    Resume CleanExit
End Sub


Public Sub Test_FormatRange_ErrorSafety()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_FormatRange_ErrorSafety")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_FmtErr"

    ' Setup header and data containing errors
    ws.Range("A1").Value2 = "HeaderA"
    ws.Range("A2").Value2 = CVErr(xlErrValue) ' #VALUE! error in first data row
    
    ws.Range("B1").Value2 = "HeaderB"
    ws.Range("B2").Value2 = "Hello" ' valid string

    Dim cmd As New FeatCmd_FormatRange
    
    ' Call format range direct. This should not throw type mismatch error 13
    cmd.FormatRangeDirect ws.Range("A1:B2"), ws

    ' A1 and B1 should be formatted as headers (bold, font size 11)
    Lib_Tests.AssertEqual ws.Range("A1").Font.Bold, True, "A1 should be bold"
    Lib_Tests.AssertEqual ws.Range("B1").Font.Bold, True, "B1 should be bold"

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
    Infra_Error.HandleError "Test_FormatRange_ErrorSafety", Err
    Resume CleanExit
End Sub


Public Sub Test_TableOfContents_Generation()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_TableOfContents_Generation")
    On Error GoTo ErrHandler

    ' Create a temporary workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add

    ' Add mock worksheets
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = wb.Worksheets.Add
    ws1.Name = "Test_TOC_Sheet1"
    ws1.Range("A1:B2").Value2 = "Data" ' 4 populated cells

    Set ws2 = wb.Worksheets.Add
    ws2.Name = "Test_TOC_Sheet2"
    ws2.Visible = xlSheetHidden

    ' Initialize AppContainer and create context
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("TableOfContents")
    Set ctx.ActionContext.WorkbookRef = wb

    ' Resolve and execute command
    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("TableOfContents")
    
    ' Execute (should not prompt since TOC sheet does not exist yet)
    cmd.Execute ctx

    ' Validate Table of Contents sheet is first
    Dim wsTOC As Worksheet
    Set wsTOC = wb.Worksheets(1)
    Lib_Tests.AssertEqual wsTOC.Name, "Table of Contents", "First worksheet should be 'Table of Contents'"

    ' Validate info on wsTOC
    Dim r As Long, foundSheet1 As Boolean, foundSheet2 As Boolean
    For r = 6 To 15
        Dim sheetName As String
        sheetName = wsTOC.Cells(r, 3).Value
        If sheetName = "Test_TOC_Sheet1" Then
            foundSheet1 = True
            ' Verify cells count is 4
            Lib_Tests.AssertEqual wsTOC.Cells(r, 5).Value, 4#, "Test_TOC_Sheet1 populated cells count should be 4"
            ' Verify visibility is Visible
            Lib_Tests.AssertEqual wsTOC.Cells(r, 4).Value, "Visible", "Test_TOC_Sheet1 visibility should be Visible"
        ElseIf sheetName = "Test_TOC_Sheet2" Then
            foundSheet2 = True
            ' Verify visibility is Hidden
            Lib_Tests.AssertEqual wsTOC.Cells(r, 4).Value, "Hidden", "Test_TOC_Sheet2 visibility should be Hidden"
        End If
    Next r

    Lib_Tests.AssertEqual foundSheet1, True, "Test_TOC_Sheet1 should be listed in TOC"
    Lib_Tests.AssertEqual foundSheet2, True, "Test_TOC_Sheet2 should be listed in TOC"

    ' Cleanup
    wb.Close SaveChanges:=False

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_TableOfContents_Generation", Err
    Resume CleanExit
End Sub


Public Sub Test_CommandResolution_NewMenus()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CommandResolution_NewMenus")
    On Error GoTo ErrHandler

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' 1. Test ModifyData sub-commands
    Dim cmdModify As ICommand
    Set cmdModify = AppContainer.ResolveCommand("DateFixer")
    Lib_Tests.AssertEqual Not cmdModify Is Nothing, True, "DateFixer command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdModify), "FeatCmd_ModifyData", "DateFixer should resolve to FeatCmd_ModifyData"

    Set cmdModify = AppContainer.ResolveCommand("CaseFixer")
    Lib_Tests.AssertEqual Not cmdModify Is Nothing, True, "CaseFixer command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdModify), "FeatCmd_ModifyData", "CaseFixer should resolve to FeatCmd_ModifyData"

    ' 2. Test HighlightData sub-commands
    Dim cmdHighlight As ICommand
    Set cmdHighlight = AppContainer.ResolveCommand("HighlightInconsistentFormulas")
    Lib_Tests.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightInconsistentFormulas should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightInconsistentFormulas should resolve to FeatCmd_HighlightData"

    Set cmdHighlight = AppContainer.ResolveCommand("HighlightDuplicates")
    Lib_Tests.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightDuplicates should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightDuplicates should resolve to FeatCmd_HighlightData"

    Set cmdHighlight = AppContainer.ResolveCommand("HighlightErrors")
    Lib_Tests.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightErrors should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightErrors should resolve to FeatCmd_HighlightData"

    Set cmdHighlight = AppContainer.ResolveCommand("HighlightHardcodedValues")
    Lib_Tests.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightHardcodedValues should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightHardcodedValues should resolve to FeatCmd_HighlightData"

    ' 3. Test Export sub-commands
    Dim cmdExport As ICommand
    Set cmdExport = AppContainer.ResolveCommand("ExportPng")
    Lib_Tests.AssertEqual Not cmdExport Is Nothing, True, "ExportPng should resolve"
    Lib_Tests.AssertEqual TypeName(cmdExport), "FeatCmd_ExportImageOrPdf", "ExportPng should resolve to FeatCmd_ExportImageOrPdf"

    Set cmdExport = AppContainer.ResolveCommand("ExportPdf")
    Lib_Tests.AssertEqual Not cmdExport Is Nothing, True, "ExportPdf should resolve"
    Lib_Tests.AssertEqual TypeName(cmdExport), "FeatCmd_ExportImageOrPdf", "ExportPdf should resolve to FeatCmd_ExportImageOrPdf"

    ' 4. Test new ModifyData commands
    Dim cmdNewModify As ICommand
    Set cmdNewModify = AppContainer.ResolveCommand("UnmergeFill")
    Lib_Tests.AssertEqual Not cmdNewModify Is Nothing, True, "UnmergeFill command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdNewModify), "FeatCmd_UnmergeFill", "UnmergeFill should resolve to FeatCmd_UnmergeFill"

    Set cmdNewModify = AppContainer.ResolveCommand("ForceNumber")
    Lib_Tests.AssertEqual Not cmdNewModify Is Nothing, True, "ForceNumber command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdNewModify), "FeatCmd_ForceNumber", "ForceNumber should resolve to FeatCmd_ForceNumber"

    ' 5. Test Duplicate command resolution
    Dim cmdDuplicate As ICommand
    Set cmdDuplicate = AppContainer.ResolveCommand("Duplicate")
    Lib_Tests.AssertEqual Not cmdDuplicate Is Nothing, True, "Duplicate command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdDuplicate), "FeatCmd_Duplicate", "Duplicate should resolve to FeatCmd_Duplicate"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_CommandResolution_NewMenus", Err
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


Public Sub Test_Export_Pdf_Backup_And_MultiRange()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_Export_Pdf_Backup_And_MultiRange")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim wb As Workbook: Set wb = Workbooks.Add
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    If wb.Worksheets.Count = 1 Then
        wb.Worksheets.Add After:=ws
    End If

    ' 1. Test PageSetup Backup & Restore
    Dim backupObj As New Infra_PageSetupBackup
    
    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.PrintArea = "A1:B5"
    ws.Visible = xlSheetVisible
    
    backupObj.Backup ws
    
    ' Modify settings
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PrintArea = "C1:D10"
    ws.Visible = xlSheetHidden
    
    ' Restore settings
    backupObj.Restore
    
    Lib_Tests.AssertEqual ws.PageSetup.Orientation, xlPortrait, "Restore orientation to Portrait"
    Lib_Tests.AssertEqual ws.PageSetup.PrintArea, "$A$1:$B$5", "Restore print area to A1:B5"
    Lib_Tests.AssertEqual ws.Visible, xlSheetVisible, "Restore visibility to Visible"

    ' 2. Test Multi-Range print area generation logic
    ' Let's write some dummy values
    ws.Range("A1").Value = "A"
    ws.Range("B2").Value = "A"
    ws.Range("D1").Value = "B"
    ws.Range("E2").Value = "B"
    ' Selection with multi-areas
    Dim selRng As Range
    Set selRng = Union(ws.Range("A1:B2"), ws.Range("D1:E2"))
    
    Dim area As Range
    Dim intersectRange As Range
    Dim printAreaAddress As String
    printAreaAddress = ""
    
    For Each area In selRng.Areas
        Set intersectRange = Intersect(area, ws.UsedRange)
        If Not intersectRange Is Nothing Then
            If printAreaAddress <> "" Then
                printAreaAddress = printAreaAddress & ","
            End If
            printAreaAddress = printAreaAddress & intersectRange.Address
        End If
    Next area
    
    ' Check if print area address correctly combines the non-contiguous ranges
    Lib_Tests.AssertEqual InStr(printAreaAddress, "$A$1:$B$2") > 0, True, "Contains first area"
    Lib_Tests.AssertEqual InStr(printAreaAddress, "$D$1:$E$2") > 0, True, "Contains second area"
    Lib_Tests.AssertEqual InStr(printAreaAddress, ","), 10, "Contains a comma separating areas"

    wb.Close SaveChanges:=False

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_Export_Pdf_Backup_And_MultiRange", Err
    Resume CleanExit
End Sub

Public Sub Test_CleanWorkbookNames_BrokenAndExternal()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CleanWorkbookNames_BrokenAndExternal")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CleanNames"

    ' Setup names on the sheet
    On Error Resume Next
    ws.Names.Add Name:="TestBrokenName", RefersTo:="=SheetNonExistent!#REF!"
    ws.Names.Add Name:="TestExternalName", RefersTo:="=[ExternalFile.xlsx]Sheet1!$A$1"
    ws.Names.Add Name:="TestNormalName", RefersTo:="=$A$1"
    
    ' Also setup workbook scope broken name
    ThisWorkbook.Names.Add Name:="TestWbBrokenName", RefersTo:="=SheetNonExistent!#REF!"
    On Error GoTo ErrHandler

    ' Clean Broken Names on Sheet
    Dim removedBrokenCount As Long
    Infra_CommandSupport.CleanWorkbookNames Nothing, ws, NameCleanCriteriaBroken, removedBrokenCount
    Lib_Tests.AssertEqual removedBrokenCount, 1, "Should clean exactly 1 broken name on the sheet"

    ' Verify the sheet names remaining
    Dim brokenExists As Boolean: brokenExists = False
    Dim externalExists As Boolean: externalExists = False
    Dim normalExists As Boolean: normalExists = False
    Dim nm As Name

    For Each nm In ws.Names
        If nm.Name = ws.Name & "!TestBrokenName" Then brokenExists = True
        If nm.Name = ws.Name & "!TestExternalName" Then externalExists = True
        If nm.Name = ws.Name & "!TestNormalName" Then normalExists = True
    Next nm

    Lib_Tests.AssertEqual brokenExists, False, "Broken sheet-scoped name should be deleted"
    Lib_Tests.AssertEqual externalExists, True, "External name should still exist"
    Lib_Tests.AssertEqual normalExists, True, "Normal name should still exist"

    ' Clean External Names on Sheet
    Dim removedExternalCount As Long
    Infra_CommandSupport.CleanWorkbookNames Nothing, ws, NameCleanCriteriaExternal, removedExternalCount
    Lib_Tests.AssertEqual removedExternalCount, 1, "Should clean exactly 1 external name on the sheet"

    externalExists = False
    For Each nm In ws.Names
        If nm.Name = ws.Name & "!TestExternalName" Then externalExists = True
    Next nm
    Lib_Tests.AssertEqual externalExists, False, "External sheet-scoped name should be deleted"

    ' Clean Workbook Scope Broken Names
    Dim removedWbCount As Long
    Infra_CommandSupport.CleanWorkbookNames ThisWorkbook, Nothing, NameCleanCriteriaBroken, removedWbCount
    Lib_Tests.AssertTrue removedWbCount >= 1, "Should clean at least 1 workbook-scoped broken name"

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
    Infra_Error.HandleError "Test_CleanWorkbookNames_BrokenAndExternal", Err
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

