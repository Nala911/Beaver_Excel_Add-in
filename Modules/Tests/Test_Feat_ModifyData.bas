Attribute VB_Name = "Test_Feat_ModifyData"
Option Explicit

' @Module: Test_Feat_ModifyData
' @Category: Library
' @Description: Integration tests for data modification (casing, dates) features.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, Infra_Undo, FeatCmd_ModifyData, ModifyDataRequest

Public Sub Test_ModifyData_Casing()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_Casing")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModCase"

    ws.Range("A1").Value2 = "heLLo wOrld"
    ws.Range("A2").Value2 = "heLLo wOrld"
    ws.Range("A3").Value2 = "heLLo wOrld"

    Dim cmd As New FeatCmd_ModifyData
    Dim req As New ModifyDataRequest
    Set req.Context = New ActionContext
    
    ' 1. Test UPPERCASE
    Dim changes As Long
    req.Operation = "Case: UPPERCASE"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A1"), req)
    Test_Runner.AssertEqual ws.Range("A1").Value2, "HELLO WORLD", "Should convert to uppercase"

    ' 2. Test lowercase
    req.Operation = "Case: lowercase"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A2"), req)
    Test_Runner.AssertEqual ws.Range("A2").Value2, "hello world", "Should convert to lowercase"

    ' 3. Test Proper Case
    req.Operation = "Case: Proper Case"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A3"), req)
    Test_Runner.AssertEqual ws.Range("A3").Value2, "Hello World", "Should convert to Proper Case"

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
    Infra_Error.HandleError "Test_ModifyData_Casing", Err
    Resume CleanExit
End Sub

Public Sub Test_ModifyData_DateStandardization()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_DateStandardization")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModDate"

    ws.Range("A1:A4").NumberFormat = "@"
    ws.Range("A1").Value = "05/12/2021" ' dd/mm/yyyy -> Dec 5, 2021
    ws.Range("A2").Value = "12/05/2021" ' mm/dd/yyyy -> Dec 5, 2021
    ws.Range("A3").Value = "15-Feb-21"   ' dd-mmm-yy -> Feb 15, 2021
    ws.Range("A4").Value = "20210215"    ' yyyymmdd -> Feb 15, 2021

    Dim cmd As New FeatCmd_ModifyData
    Dim req As New ModifyDataRequest
    Set req.Context = New ActionContext
    req.Operation = "Date Standardization"

    ' 1. Test dd/mm/yyyy
    req.DatePattern = "dd/mm/yyyy"
    Dim changes As Long
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A1"), req)
    Test_Runner.AssertEqual ws.Range("A1").Value, DateSerial(2021, 12, 5), "Should convert dd/mm/yyyy to Dec 5, 2021"
    Test_Runner.AssertEqual ws.Range("A1").NumberFormat, Infra_Config.Model.DisplayDateFormat, "Format should be date format"

    ' 2. Test mm/dd/yyyy
    req.DatePattern = "mm/dd/yyyy"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A2"), req)
    Test_Runner.AssertEqual ws.Range("A2").Value, DateSerial(2021, 12, 5), "Should convert mm/dd/yyyy to Dec 5, 2021"

    ' 3. Test dd-mmm-yy
    req.DatePattern = "dd-mmm-yy"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A3"), req)
    Test_Runner.AssertEqual ws.Range("A3").Value, DateSerial(2021, 2, 15), "Should convert dd-mmm-yy to Feb 15, 2021"

    ' 4. Test yyyymmdd (non-delimited)
    req.DatePattern = "yyyymmdd"
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A4"), req)
    Test_Runner.AssertEqual ws.Range("A4").Value, DateSerial(2021, 2, 15), "Should convert yyyymmdd to Feb 15, 2021"

    ' 5. Test Date type re-interpretation (Excel parsed as 3rd June, but user meant 6th March)
    Dim cellA5 As Range
    Set cellA5 = ws.Range("A5")
    cellA5.Value = DateSerial(2010, 6, 3) ' June 3, 2010
    cellA5.NumberFormat = "dd-mm-yyyy"

    req.DatePattern = "mm-dd-yyyy"
    changes = cmd.ModifyRangeWithOptionsDirect(cellA5, req)
    Test_Runner.AssertEqual cellA5.Value, DateSerial(2010, 3, 6), "Should re-interpret June 3, 2010 as March 6, 2010 using mm-dd-yyyy"
    Test_Runner.AssertEqual changes, 1, "Should report 1 change"

    ' 6. Test Date type re-interpretation with built-in format, adjusting for system locale
    Dim cellA6 As Range
    Set cellA6 = ws.Range("A6")
    Dim systemOrder As Long
    On Error Resume Next
    systemOrder = Application.International(xlDateOrder)
    If Err.Number <> 0 Then systemOrder = 1 ' Default to DMY
    On Error GoTo ErrHandler
    
    If systemOrder = 1 Then ' DMY system
        cellA6.Value = DateSerial(2021, 3, 1) ' March 1, 2021
        cellA6.NumberFormat = "m/d/yyyy" ' Format that formats to March 1, 2021 (e.g. 3/1/2021)
        req.DatePattern = "mm-dd-yyyy"
        changes = cmd.ModifyRangeWithOptionsDirect(cellA6, req)
        Test_Runner.AssertEqual cellA6.Value, DateSerial(2021, 1, 3), "DMY: Should re-interpret March 1, 2021 as Jan 3, 2021 using m/d/yyyy format"
        Test_Runner.AssertEqual changes, 1, "DMY: Should report 1 change"
    ElseIf systemOrder = 0 Then ' MDY system
        cellA6.Value = DateSerial(2021, 1, 3) ' Jan 3, 2021
        cellA6.NumberFormat = "dd-mm-yyyy" ' Format that formats to Jan 3, 2021 (e.g. 03-01-2021)
        req.DatePattern = "dd-mm-yyyy"
        changes = cmd.ModifyRangeWithOptionsDirect(cellA6, req)
        Test_Runner.AssertEqual cellA6.Value, DateSerial(2021, 3, 1), "MDY: Should re-interpret Jan 3, 2021 as March 1, 2021 using dd-mm-yyyy format"
        Test_Runner.AssertEqual changes, 1, "MDY: Should report 1 change"
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
    Infra_Error.HandleError "Test_ModifyData_DateStandardization", Err
    Resume CleanExit
End Sub

Public Sub Test_ModifyData_Undo()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_Undo")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModUndo"

    ws.Range("A1").Value2 = "original text"

    ' Setup command context
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("ModifyData")
    Set ctx.ActionContext.WorksheetRef = ws
    Set ctx.ActionContext.WorkbookRef = ThisWorkbook
    Set ctx.ActionContext.SelectionRange = ws.Range("A1")
    ctx.ActionContext.HasRangeSelection = True

    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("ModifyData")

    ' Run headless modification directly on A1 (Proper Case)
    Dim req As New ModifyDataRequest
    Set req.Context = ctx.ActionContext
    req.Operation = "Case: Proper Case"

    Dim featCmd As FeatCmd_ModifyData
    Set featCmd = cmd
    
    ' Save undo state first
    If Not Infra_Undo.SaveStateOrConfirm(ws.Range("A1"), "Modify Data") Then
        Err.Raise 5, "Test_ModifyData_Undo", "Failed to save undo state"
    End If
    
    Dim changes As Long
    changes = featCmd.ModifyRangeWithOptionsDirect(ws.Range("A1"), req)
    
    Test_Runner.AssertEqual ws.Range("A1").Value2, "Original Text", "A1 should be Proper Case"

    ' Register pending undo and perform
    Infra_Undo.RegisterPendingUndo
    Infra_Undo.PerformUndo

    Test_Runner.AssertEqual ws.Range("A1").Value2, "original text", "A1 should be restored to original text after undo"

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
    Infra_Error.HandleError "Test_ModifyData_Undo", Err
    Resume CleanExit
End Sub

Public Sub Test_ModifyData_MixedFormats()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_MixedFormats")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModMixedFmts"

    ' Setup cells with different number formats
    ws.Range("A1").NumberFormat = "General"
    ws.Range("A2").NumberFormat = "$#,##0"
    ws.Range("A3").NumberFormat = "@"

    ws.Range("A1").Value2 = "apple"
    ws.Range("A2").Value2 = "banana"
    ws.Range("A3").Value2 = "cherry"

    Dim cmd As New FeatCmd_ModifyData
    Dim req As New ModifyDataRequest
    Set req.Context = New ActionContext
    
    req.Operation = "Case: UPPERCASE"
    
    Dim changes As Long
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A1:A3"), req)
    
    Test_Runner.AssertEqual ws.Range("A1").Value2, "APPLE", "A1 should be uppercase"
    Test_Runner.AssertEqual ws.Range("A2").Value2, "BANANA", "A2 should be uppercase"
    Test_Runner.AssertEqual ws.Range("A3").Value2, "CHERRY", "A3 should be uppercase"
    Test_Runner.AssertEqual changes, 3, "Should report 3 changes"

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
    Infra_Error.HandleError "Test_ModifyData_MixedFormats", Err
    Resume CleanExit
End Sub

Public Sub Test_ModifyData_MixedFormats_DateStandardization()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_MixedFormats_DateStandardization")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModMixedDates"

    ' Setup cells with mixed formats: 2 dates and 1 currency cell (66% date changes, which is > 50%)
    ws.Range("A1").NumberFormat = "@"
    ws.Range("A1").Value = "15/02/2021"
    
    ws.Range("A2").NumberFormat = "@"
    ws.Range("A2").Value = "20/02/2021"
    
    ws.Range("A3").NumberFormat = "$#,##0"
    ws.Range("A3").Value = 1000

    Dim cmd As New FeatCmd_ModifyData
    Dim req As New ModifyDataRequest
    Set req.Context = New ActionContext
    
    req.Operation = "Date Standardization"
    req.DatePattern = "dd/mm/yyyy"
    
    Dim changes As Long
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A1:A3"), req)
    
    ' A1 and A2 must be converted to date values and formatted as standard dates
    Test_Runner.AssertEqual ws.Range("A1").Value, DateSerial(2021, 2, 15), "A1 should be standardized to Feb 15, 2021"
    Test_Runner.AssertEqual ws.Range("A1").NumberFormat, Infra_Config.Model.DisplayDateFormat, "A1 format should be standard date format"
    
    Test_Runner.AssertEqual ws.Range("A2").Value, DateSerial(2021, 2, 20), "A2 should be standardized to Feb 20, 2021"
    Test_Runner.AssertEqual ws.Range("A2").NumberFormat, Infra_Config.Model.DisplayDateFormat, "A2 format should be standard date format"
    
    ' A3 (Currency) must NOT have its format overwritten, even though formatting changes exceed 50% of the range
    Test_Runner.AssertEqual ws.Range("A3").NumberFormat, "$#,##0", "A3 currency format must be preserved"
    Test_Runner.AssertEqual ws.Range("A3").Value, 1000, "A3 value must be preserved"
    
    Test_Runner.AssertEqual changes, 2, "Should report exactly 2 changes"

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
    Infra_Error.HandleError "Test_ModifyData_MixedFormats_DateStandardization", Err
    Resume CleanExit
End Sub

Public Sub Test_ModifyData_BulkFormattingOverride()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ModifyData_BulkFormattingOverride")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ModFmt"

    ' Setup cells
    ws.Range("A1:A2").NumberFormat = "@"
    ws.Range("A1").Value2 = "05/12/2021" ' dd/mm/yyyy
    ws.Range("A2").Value2 = "15/02/2021" ' dd/mm/yyyy

    Dim cmd As New FeatCmd_ModifyData
    Dim req As New ModifyDataRequest
    Set req.Context = New ActionContext
    req.Operation = "Date Standardization"
    req.DatePattern = "dd/mm/yyyy"

    ' Execute
    Dim changes As Long
    changes = cmd.ModifyRangeWithOptionsDirect(ws.Range("A1:A2"), req)
    
    Test_Runner.AssertEqual changes, 2#, "Should perform 2 changes"
    Test_Runner.AssertEqual ws.Range("A1").NumberFormat, Infra_Config.Model.DisplayDateFormat, "A1 format should be display date format"
    Test_Runner.AssertEqual ws.Range("A2").NumberFormat, Infra_Config.Model.DisplayDateFormat, "A2 format should be display date format"

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
    Infra_Error.HandleError "Test_ModifyData_BulkFormattingOverride", Err
    Resume CleanExit
End Sub
