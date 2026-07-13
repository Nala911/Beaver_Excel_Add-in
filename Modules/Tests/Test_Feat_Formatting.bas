Attribute VB_Name = "Test_Feat_Formatting"
Option Explicit

' @Module: Test_Feat_Formatting
' @Category: Library
' @Description: Unit and integration tests for formatting feature commands.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, Infra_Undo, FeatCmd_ApplyCustomNumberFormat, FeatCmd_ApplyDefaultFormat, FeatCmd_PasteFormat, FeatCmd_FormatRange

Public Sub Test_ApplyDefaultFormat_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ApplyDefaultFormat_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_DefaultFormat"

    ws.Range("A1").Value2 = 1234.56

    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("ApplyDefaultFormat")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("ApplyDefaultFormat")
    cmd.Execute ctx

    Test_Runner.AssertEqual ws.Range("A1").NumberFormat, Infra_Config.Model.DefaultNumberFormat, "Default number format should be applied to A1"

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
    Infra_Error.HandleError "Test_ApplyDefaultFormat_Execution", Err
    Resume CleanExit
End Sub

Public Sub Test_ApplyDefaultFormat_WholeSheetSafety()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ApplyDefaultFormat_WholeSheetSafety")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_DefFmtWhole"

    ' Set value in C3
    ws.Range("C3").Value2 = 1234.56

    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("ApplyDefaultFormat")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Cells ' Entire worksheet selected
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("ApplyDefaultFormat")
    cmd.Execute ctx

    ' C3 should be formatted
    Test_Runner.AssertEqual ws.Range("C3").NumberFormat, Infra_Config.Model.DefaultNumberFormat, "C3 should be formatted since it is in the UsedRange"

    ' Z99 (outside UsedRange) should remain unformatted (General format)
    Test_Runner.AssertEqual ws.Range("Z99").NumberFormat, "General", "Z99 should not be formatted"

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
    Infra_Error.HandleError "Test_ApplyDefaultFormat_WholeSheetSafety", Err
    Resume CleanExit
End Sub

Public Sub Test_ApplyCustomNumberFormat_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ApplyCustomNumberFormat_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CustomFormat"

    ws.Range("A1").Value2 = 1234.56
    ws.Range("A1").NumberFormat = "General"

    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' Pass the mock custom format via sourceName (TEST_FORMAT:<format>)
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("ApplyCustomNumberFormat", , "TEST_FORMAT:#,##0.0 "" kwh""")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("ApplyCustomNumberFormat")
    cmd.Execute ctx

    Test_Runner.AssertEqual ws.Range("A1").NumberFormat, "#,##0.0 "" kwh""", "Custom number format should be applied via mock SourceName"

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

Public Sub Test_ApplyCustomNumberFormat_WholeSheetSafety()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ApplyCustomNumberFormat_WholeSheetSafety")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CustFmtWhole"

    ' Set value in C3
    ws.Range("C3").Value2 = 1234.56
    ws.Range("C3").NumberFormat = "General"

    ' Initialize AppContainer
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' Pass mock custom format via sourceName
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("ApplyCustomNumberFormat", , "TEST_FORMAT:#,##0.0 "" kwh""")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Cells ' Entire worksheet selected
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("ApplyCustomNumberFormat")
    cmd.Execute ctx

    ' C3 should be formatted
    Test_Runner.AssertEqual ws.Range("C3").NumberFormat, "#,##0.0 "" kwh""", "C3 should be formatted via mock format"

    ' Z99 (outside UsedRange) should remain unformatted (General format)
    Test_Runner.AssertEqual ws.Range("Z99").NumberFormat, "General", "Z99 should not be formatted"

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
    Infra_Error.HandleError "Test_ApplyCustomNumberFormat_WholeSheetSafety", Err
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
        On Error Resume Next
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        On Error GoTo ErrHandler
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

    Test_Runner.AssertEqual ws.Range("B1").Value2, "Dest", "Value of B1 should remain unchanged"
    Test_Runner.AssertEqual ws.Range("B1").Font.Bold, True, "B1 should now have bold formatting"
    Test_Runner.AssertEqual ws.Range("B1").Interior.Color, vbGreen, "B1 should now have green fill color"

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
    Test_Runner.AssertEqual ws.ListObjects.Count, 0#, "All overlapping ListObject tables should be unlisted"
    Test_Runner.AssertEqual ws.Range("A3").MergeCells, False, "Merged cells should be unmerged"
    Test_Runner.AssertEqual ws.Range("A1").Font.Bold, True, "Header row A1 should be Bold"
    Test_Runner.AssertEqual ws.Range("A1").Font.Size, Infra_Config.Model.HeaderFontSize, "Header font size should match config"
    Test_Runner.AssertEqual ws.Range("A1").Interior.Color, Infra_Config.Model.HeaderColor, "Header color should match config"
    Test_Runner.AssertEqual ws.Range("A2").Font.Size, Infra_Config.Model.DefaultFontSize, "Data row font size should match default config"

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
    Test_Runner.AssertEqual ws.Range("C3").Font.Bold, True, "C3 should be bolded as it became the header of the intersected used range"
    Test_Runner.AssertEqual ws.Range("C3").Font.Size, 11#, "C3 font size should be 11"

    ' Cell outside used range (e.g., Z99) should remain unformatted
    Test_Runner.AssertEqual ws.Range("Z99").Font.Bold, False, "Z99 should not be bolded"

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
    Test_Runner.AssertEqual ws.Range("A1").Font.Bold, True, "A1 should be bold"
    Test_Runner.AssertEqual ws.Range("B1").Font.Bold, True, "B1 should be bold"

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
