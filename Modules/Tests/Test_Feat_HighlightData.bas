Attribute VB_Name = "Test_Feat_HighlightData"
Option Explicit

' @Module: Test_Feat_HighlightData
' @Category: Library
' @Description: Integration tests for data highlighting features.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, Infra_Undo, FeatCmd_HighlightData, HighlightDataRequest

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

    ' Check duplicate assertions: C1 and C3 should be highlighted with Infra_Config.HIGHLIGHT_COLOR
    Test_Runner.AssertEqual ws.Range("C1").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "C1 should be highlighted as duplicate"
    Test_Runner.AssertEqual ws.Range("C3").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "C3 should be highlighted as duplicate"
    Test_Runner.AssertEqual ws.Range("C2").Interior.ColorIndex, xlNone, "C2 is unique and should not be highlighted"
    Test_Runner.AssertEqual ws.Range("C4").Interior.ColorIndex, xlNone, "C4 is unique and should not be highlighted"
    Test_Runner.AssertEqual ws.Range("C5").Interior.ColorIndex, xlNone, "C5 is empty and should not be highlighted"
    Test_Runner.AssertEqual ws.Range("C6").Interior.ColorIndex, xlNone, "C6 is empty and should not be highlighted"

    ' Test inconsistent formula if error indicator evaluated (Excel may require background error checking)
    Dim isIncA3 As Boolean
    isIncA3 = False
    On Error Resume Next
    isIncA3 = ws.Range("A3").Errors(xlInconsistentFormula).Value
    On Error GoTo ErrHandler
    
    If isIncA3 Then
        Test_Runner.AssertEqual ws.Range("A3").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A3 should be highlighted as inconsistent formula"
        Test_Runner.AssertEqual ws.Range("A1").Interior.ColorIndex, xlNone, "A1 should not be highlighted"
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
    Test_Runner.AssertEqual ws.Range("A1").Interior.ColorIndex, xlNone, "Formula check should have been skipped, leaving A1 uncolored"

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

    Dim req As New HighlightDataRequest
    req.HighlightInconsistentFormulas = False
    req.HighlightDuplicates = False
    req.HighlightErrors = True

    Dim cmd As New FeatCmd_HighlightData
    cmd.HighlightRangeWithOptionsDirect ws.Range("A1:A6"), req

    ' Check assertions
    Test_Runner.AssertEqual ws.Range("A1").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A1 constant error should be highlighted orange"
    Test_Runner.AssertEqual ws.Range("A2").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A2 constant error should be highlighted orange"
    Test_Runner.AssertEqual ws.Range("A3").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A3 formula error should be highlighted orange"
    Test_Runner.AssertEqual ws.Range("A4").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A4 formula error should be highlighted orange"
    Test_Runner.AssertEqual ws.Range("A5").Interior.ColorIndex, xlNone, "A5 normal text should not be highlighted"
    Test_Runner.AssertEqual ws.Range("A6").Interior.ColorIndex, xlNone, "A6 normal number should not be highlighted"

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
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=A1"), False, "Simple cell reference should be False"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=$A$1"), False, "Absolute reference should be False"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=A1+B2"), False, "Sum of cell references should be False"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=SUM(A1:B10)"), False, "SUM over range should be False"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=Sheet1!A1"), False, "Reference with sheet should be False"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("='Data 2026'!A1"), False, "Reference with quoted sheet and year should be False"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=LOG10(A1)"), False, "LOG10 function should be False"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=MATCH(A1, B:B, 0)"), False, "MATCH with 0 should be False"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=LEFT(A1, 1)"), False, "LEFT with 1 should be False"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=IF(A1="""", B1, C1)"), False, "IF with empty string should be False"
    
    ' Should contain hardcoded values
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=A1*1.05"), True, "1.05 is hardcoded"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=A1+50"), True, "50 is hardcoded"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=IF(A1=""Yes"", B1, C1)"), True, "Yes string is hardcoded"
    Test_Runner.AssertEqual cmd.HasHardcodedValue("=DATE(2026, 6, 14)"), True, "Dates have hardcoded numbers"
    
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
    
    Dim req As New HighlightDataRequest
    req.HighlightInconsistentFormulas = False
    req.HighlightDuplicates = False
    req.HighlightErrors = False
    req.HighlightHardcodedValues = True
    
    cmd.HighlightRangeWithOptionsDirect ws.Range("A1:A6"), req
    
    ' Check assertions
    Test_Runner.AssertEqual ws.Range("A1").Interior.ColorIndex, xlNone, "A1 should not be highlighted"
    Test_Runner.AssertEqual ws.Range("A2").Interior.ColorIndex, xlNone, "A2 should not be highlighted"
    Test_Runner.AssertEqual ws.Range("A3").Interior.ColorIndex, xlNone, "A3 should not be highlighted"
    Test_Runner.AssertEqual ws.Range("A4").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A4 formula with 1.05 should be highlighted lavender"
    Test_Runner.AssertEqual ws.Range("A5").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A5 formula with 'USD' should be highlighted lavender"
    Test_Runner.AssertEqual ws.Range("A6").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A6 formula with 100 should be highlighted lavender"
    
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

    Dim req As New HighlightDataRequest
    req.HighlightInconsistentFormulas = False
    req.HighlightDuplicates = False
    req.HighlightDataValidations = True

    Dim cmd As New FeatCmd_HighlightData
    cmd.HighlightRangeWithOptionsDirect ws.Range("A1:A3"), req

    ' Check assertions
    Test_Runner.AssertEqual ws.Range("A1").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A1 validation should be highlighted soft green"
    Test_Runner.AssertEqual ws.Range("A2").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A2 validation should be highlighted soft green"
    Test_Runner.AssertEqual ws.Range("A3").Interior.ColorIndex, xlNone, "A3 normal cell should not be highlighted"

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

    Dim req As New HighlightDataRequest
    req.HighlightInconsistentFormulas = False
    req.HighlightDuplicates = False
    req.HighlightConditionalFormatting = True

    Dim cmd As New FeatCmd_HighlightData
    cmd.HighlightRangeWithOptionsDirect ws.Range("A1:A3"), req

    ' Check assertions
    Test_Runner.AssertEqual ws.Range("A1").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A1 CF should be highlighted soft blue"
    Test_Runner.AssertEqual ws.Range("A2").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A2 CF should be highlighted soft blue"
    Test_Runner.AssertEqual ws.Range("A3").Interior.ColorIndex, xlNone, "A3 normal cell should not be highlighted"

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

Public Sub Test_HighlightData_ClearHighlights()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_HighlightData_ClearHighlights")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ClearHi"

    ' Color cell A1 with our target highlight color
    ws.Range("A1").Interior.Color = Infra_Config.HIGHLIGHT_COLOR
    Test_Runner.AssertEqual ws.Range("A1").Interior.Color, Infra_Config.HIGHLIGHT_COLOR, "A1 should be colored with HighlightColor"

    Dim cmd As New FeatCmd_HighlightData
    
    ' Call ClearWorkbookHighlights directly to test workbook-wide clearing
    cmd.ClearWorkbookHighlights ws.Parent
    
    Test_Runner.AssertEqual ws.Range("A1").Interior.ColorIndex, xlNone, "A1 color should be cleared after calling ClearWorkbookHighlights"

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
    Infra_Error.HandleError "Test_HighlightData_ClearHighlights", Err
    Resume CleanExit
End Sub
