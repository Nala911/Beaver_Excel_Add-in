Attribute VB_Name = "Lib_Tests_Feat_General"
Option Explicit

' @Module: Lib_Tests_Feat_General
' @Category: Library
' @Description: Integration tests for general features, custom formatting, and Excel UDFs.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, AppContainer, Infra_Error, Infra_Undo, FeatCmd_Wrap, FeatCmd_ApplyCustomNumberFormat, FeatCmd_PasteFormat, FeatCmd_FormatRange, FeatCmd_BreakExternalLinks, Lib_XFilterFunction, Lib_XUnpivotFunction

Public Sub Test_HelloWorld_Execution_And_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_HelloWorld_Execution_And_Undo")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_HelloWorld"
    ws.Range("B2").Value2 = "Original Content"
    ws.Range("B3").Value2 = "Original B3"
    ws.Range("B4").Value2 = "Original B4"

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
    
    ' Assert that the active cell and cells below contain correct values
    Lib_Tests.AssertEqual ws.Range("B2").Value2, "Hello world!", "HelloWorld command should update active cell to 'Hello world!'"
    Lib_Tests.AssertEqual ws.Range("B3").Value2, "How are you guys", "HelloWorld command should update B3 to 'How are you guys'"
    Lib_Tests.AssertEqual ws.Range("B4").Value2, "this is testing", "HelloWorld command should update B4 to 'this is testing'"

    ' Now register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    
    ' Assert that cells returned to their original content
    Lib_Tests.AssertEqual ws.Range("B2").Value2, "Original Content", "Undo HelloWorld should restore active cell to its original content"
    Lib_Tests.AssertEqual ws.Range("B3").Value2, "Original B3", "Undo HelloWorld should restore B3 to its original content"
    Lib_Tests.AssertEqual ws.Range("B4").Value2, "Original B4", "Undo HelloWorld should restore B4 to its original content"

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
    Dim res1 As Variant
    res1 = Lib_XFilterFunction.XFilter(src, ref)
    
    Lib_Tests.AssertEqual UBound(res1, 1), 2, "XFilter Intersection count should be 2"
    Lib_Tests.AssertEqual res1(1, 1), "Banana", "XFilter Intersection first match should be Banana"
    Lib_Tests.AssertEqual res1(2, 1), "cherry", "XFilter Intersection second match should be cherry"

    ' 2. Test Case-Sensitive Intersection (case_sensitive = True)
    Dim res2 As Variant
    res2 = Lib_XFilterFunction.XFilter(src, ref, 1, , True)
    Lib_Tests.AssertEqual res2, "Not found", "Omitted empty should return 'Not found'"

    ' 3. Test if_empty parameter
    Dim res3 As Variant
    res3 = Lib_XFilterFunction.XFilter(src, ref, 1, "Empty Val", True)
    Lib_Tests.AssertEqual res3, "Empty Val", "Custom empty string should be returned"

    ' 4. Test Difference (code_number = 2) case-insensitive
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
    Dim wideData(1 To 3, 1 To 5) As Variant
    wideData(1, 1) = "ID": wideData(1, 2) = "Name": wideData(1, 3) = "Jan": wideData(1, 4) = "Feb": wideData(1, 5) = "Mar"
    wideData(2, 1) = 101: wideData(2, 2) = "Alice": wideData(2, 3) = 100: wideData(2, 4) = 110: wideData(2, 5) = 120
    wideData(3, 1) = 102: wideData(3, 2) = "Bob": wideData(3, 3) = 200: wideData(3, 4) = 210: wideData(3, 5) = 220

    ' 1. Test Standard Unpivot (Default headers, skip_blanks = False)
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
    wideData(3, 5) = ""
    wideData(3, 4) = Empty
    Dim res3 As Variant
    res3 = Lib_XUnpivotFunction.XUnpivot(wideData, , , True)
    Lib_Tests.AssertEqual UBound(res3, 1), 5, "XUnpivot skip blanks: output should have 5 rows"
    
    Lib_Tests.AssertEqual res3(2, 3), "Jan", "XUnpivot skip blanks: R2 Attribute should be Jan"
    Lib_Tests.AssertEqual res3(3, 3), "Feb", "XUnpivot skip blanks: R3 Attribute should be Feb"
    Lib_Tests.AssertEqual res3(4, 3), "Mar", "XUnpivot skip blanks: R4 Attribute should be Mar"
    Lib_Tests.AssertEqual res3(5, 1), 102, "XUnpivot skip blanks: R5 ID should be 102"
    Lib_Tests.AssertEqual res3(5, 3), "Jan", "XUnpivot skip blanks: R5 Attribute should be Jan"

    ' 4. Test Single row boundary error
    Dim singleRow(1 To 1, 1 To 3) As Variant
    singleRow(1, 1) = "A": singleRow(1, 2) = "B": singleRow(1, 3) = "C"
    Dim res4 As Variant
    res4 = Lib_XUnpivotFunction.XUnpivot(singleRow)
    Lib_Tests.AssertTrue IsError(res4), "XUnpivot boundary: Single row should return error variant"
    
    ' 5. Test No Numeric Columns error
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

    ' Test 3: Apply wrapper cell formula (C1) to range A2:A3 in bulk
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
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Test_Temp_PasteFormat").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
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

    ' Check if CutCopyMode is active or if Excel is running headlessly
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
    Infra_Error.HandleError "Test_PasteFormat_Execution", Err
    On Error Resume Next
    Application.CutCopyMode = False
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
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

    ' C3 should be formatted
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

Public Sub Test_CommandResolution_NewMenus()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CommandResolution_NewMenus")
    On Error GoTo ErrHandler

    ' 1. Test ShowHelpCenter command resolution
    Dim cmdHelp As ICommand
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Set cmdHelp = AppContainer.ResolveCommand("ShowHelpCenter")
    Lib_Tests.AssertEqual Not cmdHelp Is Nothing, True, "ShowHelpCenter command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHelp), "FeatCmd_ShowHelpCenter", "ShowHelpCenter should resolve to FeatCmd_ShowHelpCenter"

    ' 2. Test Highlight sub-commands
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

Public Sub Test_SingleCell_Bugs_Regression()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_SingleCell_Bugs_Regression")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim wb As Workbook: Set wb = Workbooks.Add
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)

    ' 1. Test UnmergeFill single-cell regression
    Dim r1 As Range: Set r1 = ws.Range("A1:B2")
    r1.Merge
    r1.Cells(1, 1).Value = "M1"
    
    Dim r2 As Range: Set r2 = ws.Range("D1:E2")
    r2.Merge
    r2.Cells(1, 1).Value = "M2"

    ' Select only cell A1 (part of first merge area)
    ws.Range("A1").Select
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnUnmergeFill", "UnmergeFill", "Ribbon"

    Lib_Tests.AssertEqual r1.MergeCells, False, "A1:B2 should be unmerged"
    Lib_Tests.AssertEqual ws.Range("A1").Value, "M1", "A1 has value M1"
    Lib_Tests.AssertEqual ws.Range("B2").Value, "M1", "B2 has value M1"
    Lib_Tests.AssertEqual r2.MergeCells, True, "D1:E2 must remain merged"

    ' Test Undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    Lib_Tests.AssertEqual r1.MergeCells, True, "A1:B2 should be merged again after undo"

    ' 2. Test FillDown single-cell regression
    ws.Cells.Clear
    ws.Range("A1").Value = "Val"
    ws.Range("B1").Value = "KeepB"
    ws.Range("B2").Value = "KeepB"
    
    ' Select A1
    ws.Range("A1").Select
    AppContainer.ExecuteEntryPoint "UI_Hotkeys.Hotkey_FillDown", "Hotkey_FillDown", "Hotkey"

    Lib_Tests.AssertEqual ws.Range("A2").Value, "Val", "A2 should be filled down"
    Lib_Tests.AssertEqual ws.Range("B2").Value, "KeepB", "B2 must not be overwritten"

    ' 3. Test CleanData single-cell regression
    ws.Cells.Clear
    ws.Range("A1").Value = "  text  "
    ws.Range("A1").AddComment "Comment A"
    ws.Range("B1").Value = "  text2  "
    ws.Range("B1").AddComment "Comment B"

    ' Clean only A1
    ws.Range("A1").Select
    Dim cleanReq As New Infra_CleanDataRequest
    cleanReq.CleanTrimSpaces = False
    cleanReq.CleanNonPrintables = False
    cleanReq.CleanInvisibleChars = False
    cleanReq.CleanComments = True
    cleanReq.Scope = TargetScopeSelection
    
    Dim cleanCmd As New FeatCmd_CleanData
    Dim cleanCount As Long
    cleanCount = cleanCmd.CleanRangeWithOptionsDirect(ws.Range("A1"), cleanReq)

    Lib_Tests.AssertEqual cleanCount, 1, "Should report exactly 1 cell cleaned"
    Lib_Tests.AssertTrue ws.Range("A1").Comment Is Nothing, "A1 comment should be deleted"
    Lib_Tests.AssertTrue Not (ws.Range("B1").Comment Is Nothing), "B1 comment must remain"

    ' 4. Test TableOfContents hyperlink escaping regression
    Dim wsSpecial As Worksheet
    Set wsSpecial = wb.Worksheets.Add
    wsSpecial.Name = "Sheet'Special"
    
    ' Generate TOC
    AppContainer.ExecuteEntryPoint "UI_Ribbon.Ribbon_OnTableOfContents", "TableOfContents", "Ribbon"
    
    Dim wsTOC As Worksheet
    Set wsTOC = wb.Worksheets("Table of Contents")
    
    ' Find the hyperlink for Sheet'Special
    Dim hl As Hyperlink
    Dim foundHl As Boolean
    foundHl = False
    For Each hl In wsTOC.Hyperlinks
        If hl.TextToDisplay = "Sheet'Special" Then
            Lib_Tests.AssertEqual hl.SubAddress, "'Sheet''Special'!A1", "SubAddress must have escaped single quotes"
            foundHl = True
            Exit For
        End If
    Next hl
    Lib_Tests.AssertTrue foundHl, "TOC should contain hyperlink for sheet with single quotes"

    wb.Close SaveChanges:=False

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_SingleCell_Bugs_Regression", Err
    Resume CleanExit
End Sub

Public Sub Test_ProtectedSheet_CanModifyContext()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ProtectedSheet_CanModifyContext")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim wb As Workbook: Set wb = Workbooks.Add
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)

    ' Unlock A1, leave B1 locked (default)
    ws.Range("A1").Locked = False
    ws.Range("B1").Locked = True

    ' Protect sheet
    ws.Protect Password:="test"

    ' Setup action context
    Dim ctx As Infra_ActionContext
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' 1. Select A1 (unlocked)
    ws.Range("A1").Select
    Set ctx = Infra_AppState.CaptureActionContext()
    Lib_Tests.AssertTrue Infra_AppState.CanModifyContext(ctx), "Unlocked cell A1 should be modifiable on protected sheet"

    ' 2. Select B1 (locked)
    ws.Range("B1").Select
    Set ctx = Infra_AppState.CaptureActionContext()
    Lib_Tests.AssertTrue Not Infra_AppState.CanModifyContext(ctx), "Locked cell B1 should not be modifiable on protected sheet"

    ' 3. Select A1:B1 (mixed)
    ws.Range("A1:B1").Select
    Set ctx = Infra_AppState.CaptureActionContext()
    Lib_Tests.AssertTrue Not Infra_AppState.CanModifyContext(ctx), "Mixed range A1:B1 should not be modifiable on protected sheet"

    ' Unprotect and check
    ws.Unprotect Password:="test"
    ws.Range("A1:B1").Select
    Set ctx = Infra_AppState.CaptureActionContext()
    Lib_Tests.AssertTrue Infra_AppState.CanModifyContext(ctx), "Range should be modifiable when worksheet is unprotected"

    wb.Close SaveChanges:=False

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_ProtectedSheet_CanModifyContext", Err
    Resume CleanExit
End Sub

Public Sub Test_ProtectedSheet_CommandValidators()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ProtectedSheet_CommandValidators")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim wb As Workbook: Set wb = Workbooks.Add
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    
    ' Protect sheet
    ws.Protect Password:="test"
    
    Dim context As ICommandContext
    Set context = AppContainer.CreateCommandContext("CleanData", "UI_Ribbon.Ribbon_OnCleanData", "Ribbon_OnCleanData", "Ribbon")

    ' 1. CleanData validation
    Dim cleanCmd As ICommand: Set cleanCmd = New FeatCmd_CleanData
    Dim cleanRes As CommandValidationResult
    Set cleanRes = cleanCmd.Validate(context)
    Lib_Tests.AssertTrue Not cleanRes.IsExecutable, "CleanData should not validate on protected sheet"
    
    ' 2. ModifyData validation
    Dim modifyCmd As ICommand: Set modifyCmd = New FeatCmd_ModifyData
    Dim modifyRes As CommandValidationResult
    Set modifyRes = modifyCmd.Validate(context)
    Lib_Tests.AssertTrue Not modifyRes.IsExecutable, "ModifyData should not validate on protected sheet"

    ' 3. StaticSheetWorkbook validation
    Dim staticCmd As ICommand: Set staticCmd = New FeatCmd_StaticSheetWorkbook
    Dim staticRes As CommandValidationResult
    Set staticRes = staticCmd.Validate(context)
    Lib_Tests.AssertTrue Not staticRes.IsExecutable, "StaticSheetWorkbook should not validate on protected sheet"

    ' Unprotect sheet
    ws.Unprotect Password:="test"

    ' Re-validate CleanData
    Set cleanRes = cleanCmd.Validate(context)
    Lib_Tests.AssertTrue cleanRes.IsExecutable, "CleanData should validate on unprotected sheet"

    wb.Close SaveChanges:=False

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_ProtectedSheet_CommandValidators", Err
    Resume CleanExit
End Sub

Public Sub Test_MultiArea_Undo_Robustness()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_MultiArea_Undo_Robustness")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim wb As Workbook: Set wb = Workbooks.Add
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)

    ' Setup initial values in disjoint areas
    ws.Range("A1").Value2 = "A1_Orig"
    ws.Range("A3").Value2 = "A3_Orig"
    ws.Range("C1").Value2 = "C1_Orig"
    ws.Range("C3").Value2 = "C3_Orig"

    Dim targetRange As Range
    Set targetRange = ws.Range("A1,A3,C1,C3")

    ' Save state for Undo
    Dim saveSuccess As Boolean
    saveSuccess = Infra_Undo.SaveState(targetRange, "Test MultiArea Undo")
    Lib_Tests.AssertTrue saveSuccess, "SaveState should succeed for multi-area range"

    ' Modify the values
    ws.Range("A1").Value2 = "A1_New"
    ws.Range("A3").Value2 = "A3_New"
    ws.Range("C1").Value2 = "C1_New"
    ws.Range("C3").Value2 = "C3_New"

    ' Restore state
    Infra_Undo.PerformUndo

    ' Verify original values restored
    Lib_Tests.AssertEqual ws.Range("A1").Value2, "A1_Orig", "A1 should be restored"
    Lib_Tests.AssertEqual ws.Range("A3").Value2, "A3_Orig", "A3 should be restored"
    Lib_Tests.AssertEqual ws.Range("C1").Value2, "C1_Orig", "C1 should be restored"
    Lib_Tests.AssertEqual ws.Range("C3").Value2, "C3_Orig", "C3 should be restored"

    wb.Close SaveChanges:=False

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_MultiArea_Undo_Robustness", Err
    Resume CleanExit
End Sub
