Attribute VB_Name = "Lib_Tests_Feat_Wrap"
Option Explicit

' @Module: Lib_Tests_Feat_Wrap
' @Category: Library
' @Description: Unit and integration tests for formula wrapping feature.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, AppContainer, Infra_Error, FeatCmd_Wrap

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
