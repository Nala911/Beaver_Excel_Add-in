Attribute VB_Name = "Test_Feat_General"
Option Explicit

' @Module: Test_Feat_General
' @Category: Library
' @Description: Integration tests for general framework features, spill calculations, and regression checks.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, Infra_Undo, Udf_XFilter, Udf_XUnpivot, FeatCmd_CleanData, FeatCmd_ModifyData, FeatCmd_StaticSheetWorkbook, FeatCmd_UnmergeFill, FeatCmd_ForceNumber, FeatCmd_Duplicate

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
    Test_Runner.AssertEqual ws.Range("B2").Value2, "Hello world!", "HelloWorld command should update active cell to 'Hello world!'"
    Test_Runner.AssertEqual ws.Range("B3").Value2, "How are you guys", "HelloWorld command should update B3 to 'How are you guys'"
    Test_Runner.AssertEqual ws.Range("B4").Value2, "this is testing", "HelloWorld command should update B4 to 'this is testing'"

    ' Now register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    
    ' Assert that cells returned to their original content
    Test_Runner.AssertEqual ws.Range("B2").Value2, "Original Content", "Undo HelloWorld should restore active cell to its original content"
    Test_Runner.AssertEqual ws.Range("B3").Value2, "Original B3", "Undo HelloWorld should restore B3 to its original content"
    Test_Runner.AssertEqual ws.Range("B4").Value2, "Original B4", "Undo HelloWorld should restore B4 to its original content"

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
    Test_Runner.AssertEqual expandedFromAnchor.Address(False, False), "A1:A3", "Expanded range from anchor A1 should be A1:A3"

    ' Test 2: Resolve starting from a spilled cell (A2)
    Dim expandedFromSpilled As Range
    Set expandedFromSpilled = Infra_ValueConversion.ResolveSpillExpandedRange(ws.Range("A2"))
    Test_Runner.AssertEqual expandedFromSpilled.Address(False, False), "A1:A3", "Expanded range from spilled cell A2 should be A1:A3"

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
    res1 = Udf_XFilter.XFilter(src, ref)
    
    Test_Runner.AssertEqual UBound(res1, 1), 2, "XFilter Intersection count should be 2"
    Test_Runner.AssertEqual res1(1, 1), "Banana", "XFilter Intersection first match should be Banana"
    Test_Runner.AssertEqual res1(2, 1), "cherry", "XFilter Intersection second match should be cherry"

    ' 2. Test Case-Sensitive Intersection (case_sensitive = True)
    Dim res2 As Variant
    res2 = Udf_XFilter.XFilter(src, ref, 1, , True)
    Test_Runner.AssertEqual res2, "Not found", "Omitted empty should return 'Not found'"

    ' 3. Test if_empty parameter
    Dim res3 As Variant
    res3 = Udf_XFilter.XFilter(src, ref, 1, "Empty Val", True)
    Test_Runner.AssertEqual res3, "Empty Val", "Custom empty string should be returned"

    ' 4. Test Difference (code_number = 2) case-insensitive
    Dim res4 As Variant
    res4 = Udf_XFilter.XFilter(src, ref, 2)
    Test_Runner.AssertEqual UBound(res4, 1), 2, "XFilter Difference count should be 2"
    Test_Runner.AssertEqual res4(1, 1), "Apple", "XFilter Difference first match should be Apple"
    Test_Runner.AssertEqual res4(2, 1), "DATE", "XFilter Difference second match should be DATE"

    ' 5. Test 1D array conversion and scalar values
    Dim scalarSrc As Variant
    scalarSrc = "Apple"
    Dim scalarRef As Variant
    scalarRef = "Apple"
    Dim res5 As Variant
    res5 = Udf_XFilter.XFilter(scalarSrc, scalarRef, 1)
    Test_Runner.AssertEqual res5(1, 1), "Apple", "Scalar inputs should be handled correctly"

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
    res1 = Udf_XUnpivot.XUnpivot(wideData)
    
    Test_Runner.AssertEqual UBound(res1, 1), 7, "XUnpivot standard: output should have 7 rows"
    Test_Runner.AssertEqual UBound(res1, 2), 4, "XUnpivot standard: output should have 4 columns"
    
    ' Row 1 (Header)
    Test_Runner.AssertEqual res1(1, 1), "ID", "XUnpivot standard: Header col 1 should be ID"
    Test_Runner.AssertEqual res1(1, 2), "Name", "XUnpivot standard: Header col 2 should be Name"
    Test_Runner.AssertEqual res1(1, 3), "Attribute", "XUnpivot standard: Header col 3 should be Attribute"
    Test_Runner.AssertEqual res1(1, 4), "Value", "XUnpivot standard: Header col 4 should be Value"
    
    ' Row 2 (First unpivoted row for Alice)
    Test_Runner.AssertEqual res1(2, 1), 101, "XUnpivot standard: R2 C1 should be 101"
    Test_Runner.AssertEqual res1(2, 2), "Alice", "XUnpivot standard: R2 C2 should be Alice"
    Test_Runner.AssertEqual res1(2, 3), "Jan", "XUnpivot standard: R2 C3 should be Jan"
    Test_Runner.AssertEqual res1(2, 4), 100, "XUnpivot standard: R2 C4 should be 100"
    
    ' Row 4 (Third unpivoted row for Alice)
    Test_Runner.AssertEqual res1(4, 3), "Mar", "XUnpivot standard: R4 C3 should be Mar"
    Test_Runner.AssertEqual res1(4, 4), 120, "XUnpivot standard: R4 C4 should be 120"
    
    ' Row 5 (First unpivoted row for Bob)
    Test_Runner.AssertEqual res1(5, 1), 102, "XUnpivot standard: R5 C1 should be 102"
    Test_Runner.AssertEqual res1(5, 2), "Bob", "XUnpivot standard: R5 C2 should be Bob"
    Test_Runner.AssertEqual res1(5, 3), "Jan", "XUnpivot standard: R5 C3 should be Jan"
    Test_Runner.AssertEqual res1(5, 4), 200, "XUnpivot standard: R5 C4 should be 200"

    ' 2. Test Custom Headers
    Dim res2 As Variant
    res2 = Udf_XUnpivot.XUnpivot(wideData, "Month", "Sales")
    Test_Runner.AssertEqual res2(1, 3), "Month", "XUnpivot custom headers: Attribute header should be Month"
    Test_Runner.AssertEqual res2(1, 4), "Sales", "XUnpivot custom headers: Value header should be Sales"

    ' 3. Test Skip Blanks
    wideData(3, 5) = ""
    wideData(3, 4) = Empty
    Dim res3 As Variant
    res3 = Udf_XUnpivot.XUnpivot(wideData, , , True)
    Test_Runner.AssertEqual UBound(res3, 1), 5, "XUnpivot skip blanks: output should have 5 rows"
    
    Test_Runner.AssertEqual res3(2, 3), "Jan", "XUnpivot skip blanks: R2 Attribute should be Jan"
    Test_Runner.AssertEqual res3(3, 3), "Feb", "XUnpivot skip blanks: R3 Attribute should be Feb"
    Test_Runner.AssertEqual res3(4, 3), "Mar", "XUnpivot skip blanks: R4 Attribute should be Mar"
    Test_Runner.AssertEqual res3(5, 1), 102, "XUnpivot skip blanks: R5 ID should be 102"
    Test_Runner.AssertEqual res3(5, 3), "Jan", "XUnpivot skip blanks: R5 Attribute should be Jan"

    ' 4. Test Single row boundary error
    Dim singleRow(1 To 1, 1 To 3) As Variant
    singleRow(1, 1) = "A": singleRow(1, 2) = "B": singleRow(1, 3) = "C"
    Dim res4 As Variant
    res4 = Udf_XUnpivot.XUnpivot(singleRow)
    Test_Runner.AssertTrue IsError(res4), "XUnpivot boundary: Single row should return error variant"
    
    ' 5. Test No Numeric Columns error
    Dim noNumeric(1 To 2, 1 To 3) As Variant
    noNumeric(1, 1) = "ID": noNumeric(1, 2) = "Val1": noNumeric(1, 3) = "Val2"
    noNumeric(2, 1) = "101": noNumeric(2, 2) = "text1": noNumeric(2, 3) = "text2"
    Dim res5 As Variant
    res5 = Udf_XUnpivot.XUnpivot(noNumeric)
    Test_Runner.AssertTrue IsError(res5), "XUnpivot boundary: No numeric columns should return error variant"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_XUnpivot_Features", Err
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
    Test_Runner.AssertEqual chunks.Count, 4, "Should divide ranges into 4 chunks total"
    
    Test_Runner.AssertEqual chunks(1).Address(False, False), "A1:A20", "First chunk should be A1:A20"
    Test_Runner.AssertEqual chunks(2).Address(False, False), "A21:A40", "Second chunk should be A21:A40"
    Test_Runner.AssertEqual chunks(3).Address(False, False), "A41:A50", "Third chunk should be A41:A50"
    Test_Runner.AssertEqual chunks(4).Address(False, False), "C1:C10", "Fourth chunk should be C1:C10"

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

    Test_Runner.AssertEqual r1.MergeCells, False, "A1:B2 should be unmerged"
    Test_Runner.AssertEqual ws.Range("A1").Value, "M1", "A1 has value M1"
    Test_Runner.AssertEqual ws.Range("B2").Value, "M1", "B2 has value M1"
    Test_Runner.AssertEqual r2.MergeCells, True, "D1:E2 must remain merged"

    ' Test Undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    Test_Runner.AssertEqual r1.MergeCells, True, "A1:B2 should be merged again after undo"

    ' 2. Test FillDown single-cell regression
    ws.Cells.Clear
    ws.Range("A1").Value = "Val"
    ws.Range("B1").Value = "KeepB"
    ws.Range("B2").Value = "KeepB"
    
    ' Select A1
    ws.Range("A1").Select
    AppContainer.ExecuteEntryPoint "UI_Hotkeys.Hotkey_FillDown", "Hotkey_FillDown", "Hotkey"

    Test_Runner.AssertEqual ws.Range("A2").Value, "Val", "A2 should be filled down"
    Test_Runner.AssertEqual ws.Range("B2").Value, "KeepB", "B2 must not be overwritten"

    ' 3. Test CleanData single-cell regression
    ws.Cells.Clear
    ws.Range("A1").Value = "  text  "
    ws.Range("A1").AddComment "Comment A"
    ws.Range("B1").Value = "  text2  "
    ws.Range("B1").AddComment "Comment B"

    ' Clean only A1
    ws.Range("A1").Select
    Dim cleanReq As New CleanDataRequest
    cleanReq.CleanTrimSpaces = False
    cleanReq.CleanNonPrintables = False
    cleanReq.CleanInvisibleChars = False
    cleanReq.CleanComments = True
    cleanReq.Scope = TargetScopeSelection
    
    Dim cleanCmd As New FeatCmd_CleanData
    Dim cleanCount As Long
    cleanCount = cleanCmd.CleanRangeWithOptionsDirect(ws.Range("A1"), cleanReq)

    Test_Runner.AssertEqual cleanCount, 1, "Should report exactly 1 cell cleaned"
    Test_Runner.AssertTrue ws.Range("A1").Comment Is Nothing, "A1 comment should be deleted"
    Test_Runner.AssertTrue Not (ws.Range("B1").Comment Is Nothing), "B1 comment must remain"

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
            Test_Runner.AssertEqual hl.SubAddress, "'Sheet''Special'!A1", "SubAddress must have escaped single quotes"
            foundHl = True
            Exit For
        End If
    Next hl
    Test_Runner.AssertTrue foundHl, "TOC should contain hyperlink for sheet with single quotes"

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
    Dim ctx As ActionContext
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' 1. Select A1 (unlocked)
    ws.Range("A1").Select
    Set ctx = Infra_AppState.CaptureActionContext()
    Test_Runner.AssertTrue Infra_AppState.CanModifyContext(ctx), "Unlocked cell A1 should be modifiable on protected sheet"

    ' 2. Select B1 (locked)
    ws.Range("B1").Select
    Set ctx = Infra_AppState.CaptureActionContext()
    Test_Runner.AssertTrue Not Infra_AppState.CanModifyContext(ctx), "Locked cell B1 should not be modifiable on protected sheet"

    ' 3. Select A1:B1 (mixed)
    ws.Range("A1:B1").Select
    Set ctx = Infra_AppState.CaptureActionContext()
    Test_Runner.AssertTrue Not Infra_AppState.CanModifyContext(ctx), "Mixed range A1:B1 should not be modifiable on protected sheet"

    ' Unprotect and check
    ws.Unprotect Password:="test"
    ws.Range("A1:B1").Select
    Set ctx = Infra_AppState.CaptureActionContext()
    Test_Runner.AssertTrue Infra_AppState.CanModifyContext(ctx), "Range should be modifiable when worksheet is unprotected"

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
    Test_Runner.AssertTrue Not cleanRes.IsExecutable, "CleanData should not validate on protected sheet"
    
    ' 2. ModifyData validation
    Dim modifyCmd As ICommand: Set modifyCmd = New FeatCmd_ModifyData
    Dim modifyRes As CommandValidationResult
    Set modifyRes = modifyCmd.Validate(context)
    Test_Runner.AssertTrue Not modifyRes.IsExecutable, "ModifyData should not validate on protected sheet"

    ' 3. StaticSheetWorkbook validation
    Dim staticCmd As ICommand: Set staticCmd = New FeatCmd_StaticSheetWorkbook
    Dim staticRes As CommandValidationResult
    Set staticRes = staticCmd.Validate(context)
    Test_Runner.AssertTrue Not staticRes.IsExecutable, "StaticSheetWorkbook should not validate on protected sheet"

    ' Unprotect sheet
    ws.Unprotect Password:="test"

    ' Re-validate CleanData
    Set cleanRes = cleanCmd.Validate(context)
    Test_Runner.AssertTrue cleanRes.IsExecutable, "CleanData should validate on unprotected sheet"

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
    Test_Runner.AssertTrue saveSuccess, "SaveState should succeed for multi-area range"

    ' Modify the values
    ws.Range("A1").Value2 = "A1_New"
    ws.Range("A3").Value2 = "A3_New"
    ws.Range("C1").Value2 = "C1_New"
    ws.Range("C3").Value2 = "C3_New"

    ' Restore state
    Infra_Undo.PerformUndo

    ' Verify original values restored
    Test_Runner.AssertEqual ws.Range("A1").Value2, "A1_Orig", "A1 should be restored"
    Test_Runner.AssertEqual ws.Range("A3").Value2, "A3_Orig", "A3 should be restored"
    Test_Runner.AssertEqual ws.Range("C1").Value2, "C1_Orig", "C1 should be restored"
    Test_Runner.AssertEqual ws.Range("C3").Value2, "C3_Orig", "C3 should be restored"

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
