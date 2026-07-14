Attribute VB_Name = "Test_Feat_Editing"
Option Explicit

' @Module: Test_Feat_Editing
' @Category: Library
' @Description: Integration tests for basic cell editing, deleting, fill, and filter features.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, Infra_Undo, FeatCmd_MakePermanent, FeatCmd_FillDown, FeatCmd_FillRight, FeatCmd_Delete, FeatCmd_FilterByCell

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
    Test_Runner.AssertEqual ws.Range("A1").Value2, 1#, "A1 should be 1"
    Test_Runner.AssertEqual ws.Range("A2").Value2, 2#, "A2 should be 2"
    Test_Runner.AssertEqual ws.Range("A3").Value2, 3#, "A3 should be 3"

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
    Test_Runner.AssertEqual ws.Range("A1").HasFormula, False, "A1 formula should be removed"
    Test_Runner.AssertEqual ws.Range("A1").Value2, 1#, "A1 static value should be 1"
    Test_Runner.AssertEqual ws.Range("A2").Value2, 2#, "A2 static value should be 2"
    Test_Runner.AssertEqual ws.Range("A3").Value2, 3#, "A3 static value should be 3"

    ' Now register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo
    
    ' Assert that formula and dynamic array are restored
    Test_Runner.AssertEqual ws.Range("A1").HasFormula, True, "A1 formula should be restored"
    Test_Runner.AssertEqual ws.Range("A1").Formula2, "=SEQUENCE(3, 1)", "A1 formula content should be restored"
    Test_Runner.AssertEqual ws.Range("A1").Value2, 1#, "A1 restored value should be 1"
    Test_Runner.AssertEqual ws.Range("A2").Value2, 2#, "A2 restored value should be 2"
    Test_Runner.AssertEqual ws.Range("A3").Value2, 3#, "A3 restored value should be 3"

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
    Test_Runner.AssertEqual ws.Range("A1").HasArray, True, "A1 should be part of an array formula"
    Test_Runner.AssertEqual ws.Range("A1").Value2, 1#, "A1 should be 1"
    Test_Runner.AssertEqual ws.Range("A2").Value2, 2#, "A2 should be 2"
    Test_Runner.AssertEqual ws.Range("A3").Value2, 3#, "A3 should be 3"

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
    Test_Runner.AssertEqual ws.Range("A1").HasFormula, False, "A1 formula should be removed"
    Test_Runner.AssertEqual ws.Range("A1").Value2, 1#, "A1 static value should be 1"
    Test_Runner.AssertEqual ws.Range("A2").Value2, 2#, "A2 static value should be 2"
    Test_Runner.AssertEqual ws.Range("A3").Value2, 3#, "A3 static value should be 3"

    ' Now register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo

    ' Assert that CSE array formula is restored
    Test_Runner.AssertEqual ws.Range("A1").HasArray, True, "A1 array formula should be restored"
    Test_Runner.AssertEqual ws.Range("A1").Value2, 1#, "A1 restored value should be 1"
    Test_Runner.AssertEqual ws.Range("A2").Value2, 2#, "A2 restored value should be 2"
    Test_Runner.AssertEqual ws.Range("A3").Value2, 3#, "A3 restored value should be 3"

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

    Test_Runner.AssertEqual ws.Range("A5").Value2, 10#, "Multi-column filldown: A5 should be 10"
    Test_Runner.AssertEqual ws.Range("B5").Value2, 20#, "Multi-column filldown: B5 should be 20"
    Test_Runner.AssertEqual ws.Range("A6").Value2, vbEmpty, "Multi-column filldown: A6 should be empty"

    ' --- TEST 2: Proximity Search Distance Limit ---
    ws.Cells.Clear
    ws.Range("A1").Value2 = 100
    ws.Range("Q1:Q5").Value2 = "Ref" ' Column 17 (distance = 16 columns)
    
    Set ctx = AppContainer.CreateCommandContext("FillDown")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True
    
    cmd.Execute ctx
    Test_Runner.AssertEqual ws.Range("A5").Value2, vbEmpty, "Proximity limit: A5 should be empty when neighbor is > 15 columns away"

    ' Neighbor at column 16 (distance = 15 columns)
    ws.Range("P1:P5").Value2 = "Ref"
    cmd.Execute ctx
    Test_Runner.AssertEqual ws.Range("A5").Value2, 100#, "Proximity limit: A5 should be 100 when neighbor is exactly 15 columns away"

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
    Test_Runner.AssertEqual ws.Range("A3").Value2, vbEmpty, "Fragmentation guard: A3 should remain empty"
    Test_Runner.AssertEqual ws.Range("A10011").Value2, vbEmpty, "Fragmentation guard: A10011 should remain empty"

    ' --- TEST 4: No Neighbors and Source in Data Range ---
    ws.Cells.Clear
    ws.Range("A1").Value2 = 100
    ws.Range("A2:A5").Value2 = 50
    
    Set ctx = AppContainer.CreateCommandContext("FillDown")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True
    
    cmd.Execute ctx
    Test_Runner.AssertEqual ws.Range("A2").Value2, 50#, "No neighbors: A2 should remain 50"
    Test_Runner.AssertEqual ws.Range("A5").Value2, 50#, "No neighbors: A5 should remain 50"

    ' --- TEST 5: Skip Closer Neighbor with No Extra Rows to Fill ---
    ws.Cells.Clear
    ws.Range("A1:A30").Value2 = 10
    ws.Range("D1").Value2 = 100
    ws.Range("E1").Value2 = 200
    
    Set ctx = AppContainer.CreateCommandContext("FillDown")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("D1")
    ctx.ActionContext.HasRangeSelection = True
    
    cmd.Execute ctx
    Test_Runner.AssertEqual ws.Range("D2").Value2, 100#, "Skip closer invalid neighbor: D2 should be 100"
    Test_Runner.AssertEqual ws.Range("D30").Value2, 100#, "Skip closer invalid neighbor: D30 should be 100"
    Test_Runner.AssertEqual ws.Range("D31").Value2, vbEmpty, "Skip closer invalid neighbor: D31 should be empty"

    ' --- TEST 6: User Scenario with A1:A20 and E1:G1 ---
    ws.Cells.Clear
    ws.Range("A1:A20").Value2 = 10
    ws.Range("E1:G1").Value2 = 200
    
    Set ctx = AppContainer.CreateCommandContext("FillDown")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("F1")
    ctx.ActionContext.HasRangeSelection = True
    
    cmd.Execute ctx
    Test_Runner.AssertEqual ws.Range("F2").Value2, 200#, "User scenario: F2 should be 200"
    Test_Runner.AssertEqual ws.Range("F20").Value2, 200#, "User scenario: F20 should be 200"
    Test_Runner.AssertEqual ws.Range("F21").Value2, vbEmpty, "User scenario: F21 should be empty"

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

Public Sub Test_FillRight_Features()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_FillRight_Features")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_FillRight"

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' --- TEST 1: Multi-Row Fill Right ---
    ws.Range("A1").Value2 = 10
    ws.Range("A2").Value2 = 20
    ws.Range("A3:G3").Value2 = "Ref"
    
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("FillRight")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1:A2")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("FillRight")
    cmd.Execute ctx

    Test_Runner.AssertEqual ws.Range("G1").Value2, 10#, "Multi-row fillright: G1 should be 10"
    Test_Runner.AssertEqual ws.Range("G2").Value2, 20#, "Multi-row fillright: G2 should be 20"
    Test_Runner.AssertEqual ws.Range("H1").Value2, vbEmpty, "Multi-row fillright: H1 should be empty"

    ' --- TEST 2: Proximity Search Distance Limit ---
    ws.Cells.Clear
    ws.Range("A1").Value2 = 100
    ws.Range("A17:Q17").Value2 = "Ref" ' Row 17 (distance = 16 rows)
    
    Set ctx = AppContainer.CreateCommandContext("FillRight")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True
    
    cmd.Execute ctx
    Test_Runner.AssertEqual ws.Range("Q1").Value2, vbEmpty, "Proximity limit: Q1 should be empty when neighbor is > 15 rows away"

    ' Neighbor at row 16 (distance = 15 rows)
    ws.Range("A16:P16").Value2 = "Ref"
    cmd.Execute ctx
    Test_Runner.AssertEqual ws.Range("P1").Value2, 100#, "Proximity limit: P1 should be 100 when neighbor is exactly 15 rows away"

    ' --- TEST 3: No Neighbors and Source in Data Range ---
    ws.Cells.Clear
    ws.Range("A1").Value2 = 100
    ws.Range("B1:E1").Value2 = 50
    
    Set ctx = AppContainer.CreateCommandContext("FillRight")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True
    
    cmd.Execute ctx
    Test_Runner.AssertEqual ws.Range("B1").Value2, 50#, "No neighbors: B1 should remain 50"
    Test_Runner.AssertEqual ws.Range("E1").Value2, 50#, "No neighbors: E1 should remain 50"

    ' --- TEST 4: Skip Closer Neighbor with No Extra Columns to Fill ---
    ws.Cells.Clear
    ws.Range("A1:AD1").Value2 = 10
    ws.Range("A4").Value2 = 100
    ws.Range("A5").Value2 = 200
    
    Set ctx = AppContainer.CreateCommandContext("FillRight")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A4")
    ctx.ActionContext.HasRangeSelection = True
    
    cmd.Execute ctx
    Test_Runner.AssertEqual ws.Range("B4").Value2, 100#, "Skip closer invalid neighbor: B4 should be 100"
    Test_Runner.AssertEqual ws.Range("AD4").Value2, 100#, "Skip closer invalid neighbor: AD4 should be 100"
    Test_Runner.AssertEqual ws.Range("AE4").Value2, vbEmpty, "Skip closer invalid neighbor: AE4 should be empty"

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
    Infra_Error.HandleError "Test_FillRight_Features", Err
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
    Test_Runner.AssertEqual ws.Range("A1").Value2, vbEmpty, "A1 should be cleared by Backspace"
    Test_Runner.AssertEqual ws.Range("A5").Value2, vbEmpty, "A5 should be cleared by Backspace"

    ' Register and perform undo
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo

    ' Assert that the values are fully restored
    Test_Runner.AssertEqual ws.Range("A1").Value2, "Value 1", "A1 should be restored by Undo"
    Test_Runner.AssertEqual ws.Range("A5").Value2, "Value 5", "A5 should be restored by Undo"

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
    
    Test_Runner.AssertEqual ws.Range("A1").Value2, "DeleteMe", "A1 should contain text initially"
    
    ' Validate and Execute
    Dim valResult As CommandValidationResult
    Set valResult = cmd.Validate(context)
    Test_Runner.AssertEqual valResult.IsExecutable, True, "Command should be executable"
    
    cmd.Execute context
    
    Test_Runner.AssertEqual ws.Range("A1").Value2, Empty, "A1 should be cleared after Delete command"
    
    ' Perform Undo
    Infra_Undo.PerformUndo
    
    Test_Runner.AssertEqual ws.Range("A1").Value2, "DeleteMe", "A1 value should be restored by Undo"
    
    ' Test deleting a shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 50)
    shp.Select
    
    ' Re-create context for shape selection
    Set context = AppContainer.CreateCommandContext("Delete", vbNullString, "Test", vbNullString)
    cmd.Execute context
    
    Test_Runner.AssertEqual ws.Shapes.Count, 0#, "Shape should be deleted"

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
    Test_Runner.AssertEqual ws.AutoFilterMode, True, "AutoFilter should be enabled"
    
    Dim autoflt As AutoFilter
    Set autoflt = ws.AutoFilter
    Test_Runner.AssertEqual autoflt.Range.Address, ws.Range("A1:B4").Address, "Filter range should encompass A1:B4"
    
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
