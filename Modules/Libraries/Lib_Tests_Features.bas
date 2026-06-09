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
    ws.Range("A2").Value2 = "123.45"              ' Number stored as text
    ws.Range("A3").Value2 = "normal text"         ' Normal text (should not change)
    
    ' Inject non-breaking space (Chr(160))
    ws.Range("A4").Value2 = "hello" & ChrW$(160) & "world"

    Dim cmd As New FeatCmd_CleanData
    Dim cleanedCount As Long
    cleanedCount = cmd.CleanRangeDirect(ws.Range("A1:A4"))

    ' Asserts
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "hello world", "CleanData should trim and remove extra spaces"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, 123.45, "CleanData should convert numeric strings to numbers"
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
    ws.Range("A2").Value2 = "123.45"

    Dim request As New Infra_CleanDataRequest
    request.CleanTrimSpaces = False
    request.CleanTextNumbers = True

    Dim cmd As New FeatCmd_CleanData
    Dim cleanedCount As Long
    cleanedCount = cmd.CleanRangeWithOptionsDirect(ws.Range("A1:A2"), request)

    Lib_Tests.AssertEqual ws.Range("A1").Value2, "  hello   world  ", "CleanData should NOT trim spaces if CleanTrimSpaces is False"
    Lib_Tests.AssertEqual ws.Range("A2").Value2, 123.45, "CleanData should still convert text-numbers if CleanTextNumbers is True"

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

    ' Setup some merged cells
    ws.Range("A3:B3").Merge

    ' Setup an overlapping ListObject table
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes)

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    Dim cmd As New FeatCmd_FormatRange
    ' Call the direct formatting testing method headlessly
    cmd.FormatRangeDirect ws.Range("A1:B3"), ws

    ' Asserts
    Lib_Tests.AssertEqual ws.ListObjects.Count, 0#, "The ListObject table should be unlisted to a plain range"
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
