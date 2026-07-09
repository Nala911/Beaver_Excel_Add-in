Attribute VB_Name = "Test_Feat_FinancialModelling"
Option Explicit

' @Module: Test_Feat_FinancialModelling
' @Category: Library
' @Description: Unit tests for financial model coloring auto-formatter.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, FeatCmd_FinancialModelling

Public Sub Test_FinancialModelling_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_FinancialModelling_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_FinMod"

    ' Setup constant numeric cells (should be Blue)
    ws.Range("A1").Value2 = 123.45
    ws.Range("A2").Value2 = DateSerial(2026, 7, 10)

    ' Setup constant text cells (should be Black)
    ws.Range("B1").Value2 = "Revenue"
    ws.Range("B2").Value2 = "Q1 2026"

    ' Setup formula cells (should be Black)
    ws.Range("A3").Formula2 = "=A1*2"

    ' Setup cross-sheet formula cells (should be Green)
    ' First add another temp sheet to reference
    Dim wsRef As Worksheet
    Set wsRef = ThisWorkbook.Worksheets.Add
    wsRef.Name = "Test_Temp_FinRef"
    wsRef.Range("A1").Value2 = 100
    
    ws.Range("A4").Formula2 = "=Test_Temp_FinRef!A1+50"

    ' Setup external formula cells (should be Red)
    ' Using formula returning a bracket literal to avoid external file dialog prompts during testing
    ws.Range("A5").Formula2 = "=""["""

    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' Setup context
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("HighlightFinancialModelling")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1:B5")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("HighlightFinancialModelling")
    cmd.Execute ctx

    ' Assertions
    ' A1: Hardcoded numeric -> Blue
    Test_Runner.AssertEqual ws.Range("A1").Font.Color, RGB(0, 0, 255), "A1 (number) font should be Blue"
    ' A2: Hardcoded date -> Blue
    Test_Runner.AssertEqual ws.Range("A2").Font.Color, RGB(0, 0, 255), "A2 (date) font should be Blue"
    
    ' B1: Hardcoded text -> Black
    Test_Runner.AssertEqual ws.Range("B1").Font.Color, RGB(0, 0, 0), "B1 (text) font should be Black"
    ' B2: Hardcoded text -> Black
    Test_Runner.AssertEqual ws.Range("B2").Font.Color, RGB(0, 0, 0), "B2 (text) font should be Black"

    ' A3: Standard formula -> Black
    Test_Runner.AssertEqual ws.Range("A3").Font.Color, RGB(0, 0, 0), "A3 (local formula) font should be Black"

    ' A4: Cross-sheet formula -> Green
    Test_Runner.AssertEqual ws.Range("A4").Font.Color, RGB(0, 128, 0), "A4 (cross-sheet reference) font should be Green"

    ' A5: External formula -> Red
    Test_Runner.AssertEqual ws.Range("A5").Font.Color, RGB(180, 0, 0), "A5 (external reference) font should be Red"

    ' Cleanup
    Application.DisplayAlerts = False
    ws.Delete
    wsRef.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    wsRef.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_FinancialModelling_Execution", Err
    Resume CleanExit
End Sub
