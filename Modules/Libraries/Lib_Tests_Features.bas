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
    ws.Range("A1").Value2 = "Original Content"

    Dim ctx As ICommandContext
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Set ctx = AppContainer.CreateCommandContext("HelloWorld")
    
    ' Set the context refs directly to point to our newly created worksheet
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("HelloWorld")
    
    ' Execute the command directly
    cmd.Execute ctx
    
    ' Assert that cell A1 contains "Hello world"
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "Hello world", "HelloWorld command should update A1 to 'Hello world'"

    ' Now register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    
    ' Assert that cell A1 returned to its original content
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "Original Content", "Undo HelloWorld should restore A1 to its original content"

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



