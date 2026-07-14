Attribute VB_Name = "Test_Feat_Wrap"
Option Explicit

' @Module: Test_Feat_Wrap
' @Category: Library
' @Description: Unit and integration tests for formula wrapping feature.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, FeatCmd_Wrap

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
    Test_Runner.AssertEqual errCount, 0#, "Wrapping B1 pattern should have 0 errors"
    Test_Runner.AssertEqual ws.Range("B1").Formula2, "=ROUND((A1*2), 0)", "B1 formula should be successfully wrapped"

    ' Test 2: Apply wrapper cell formula (C1) to A1
    cmd.TestApplyWrapperCellDirect ws.Range("A1"), ws.Range("C1").Formula2, errCount
    Test_Runner.AssertEqual errCount, 0#, "Wrapping A1 using C1 wrapper should have 0 errors"
    Test_Runner.AssertEqual ws.Range("A1").Formula2, "=ROUND(10, 2)", "A1 formula should be successfully wrapped"

    ' Test 3: Apply wrapper cell formula (C1) to range A2:A3 in bulk
    ws.Range("A2").Formula2 = "=B1*2"
    ws.Range("A3").Formula2 = "=B1*3"
    errCount = 0
    cmd.TestApplyWrapperCellRangeDirect ws.Range("A2:A3"), ws.Range("C1").Formula2, True, errCount
    Test_Runner.AssertEqual errCount, 0#, "Wrapping A2:A3 using C1 wrapper in bulk should have 0 errors"
    Test_Runner.AssertEqual ws.Range("A2").Formula2, "=ROUND(B1*2, 2)", "A2 formula should be successfully wrapped in bulk"
    Test_Runner.AssertEqual ws.Range("A3").Formula2, "=ROUND(B1*3, 2)", "A3 formula should be successfully wrapped in bulk"

    ' Test 4: Targeting user's bug scenario (Single cell)
    ws.Range("X8").Value2 = 100
    ws.Range("V8").Value2 = 5
    ws.Range("V10").Formula2 = "=V8"
    errCount = 0
    cmd.TestApplyWrapperCellDirect ws.Range("V10"), "=X8+V10", errCount
    Test_Runner.AssertEqual errCount, 0#, "Wrapping V10 with =X8+V10 should have 0 errors"
    Test_Runner.AssertEqual ws.Range("V10").Formula2, "=X8+V8", "V10 should be wrapped by replacing the correct target reference"

    ' Test 5: Targeting user's bug scenario (Range)
    ws.Range("V9").Value2 = 6
    ws.Range("V10").Formula2 = "=V8"
    ws.Range("V11").Formula2 = "=V9"
    errCount = 0
    cmd.TestApplyWrapperCellRangeDirect ws.Range("V10:V11"), "=X8+V10", True, errCount
    Test_Runner.AssertEqual errCount, 0#, "Wrapping V10:V11 with =X8+V10 in range mode should have 0 errors"
    Test_Runner.AssertEqual ws.Range("V10").Formula2, "=X8+V8", "V10 should be wrapped correctly in range mode"
    Test_Runner.AssertEqual ws.Range("V11").Formula2, "=X8+V9", "V11 should be wrapped correctly in range mode"

    ' Test 6: Targeting user's second bug scenario (A1=10, B1=12, A5="=A1", wrapper="=B1*A5")
    ws.Range("A1").Value2 = 10
    ws.Range("B1").Value2 = 12
    ws.Range("A5").Formula2 = "=A1"
    errCount = 0
    cmd.TestApplyWrapperCellDirect ws.Range("A5"), "=B1*A5", errCount
    Test_Runner.AssertEqual errCount, 0#, "Wrapping A5 with =B1*A5 should have 0 errors"
    Test_Runner.AssertEqual ws.Range("A5").Formula2, "=B1*A1", "A5 should be wrapped to =B1*A1"

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
' Trigger rebuild
