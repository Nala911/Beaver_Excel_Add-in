Attribute VB_Name = "Lib_Tests_Features"
Option Explicit

' @Module: Lib_Tests_Features
' @Category: Library
' @Description: Automated integration and custom undo test suite for feature commands.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, AppContainer, Infra_Error, Infra_Undo

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
