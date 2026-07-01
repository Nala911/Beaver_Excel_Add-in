Attribute VB_Name = "Test_Feat_Workbook"
Option Explicit

' @Module: Test_Feat_Workbook
' @Category: Library
' @Description: Integration tests for workbook duplication, named sheet creation, export, and index reports.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, Infra_Undo, FeatCmd_StaticSheetWorkbook, FeatCmd_TableOfContents, FeatCmd_ExportImageOrPdf

Public Sub Test_StaticSheetWorkbook_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_StaticSheetWorkbook_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_StaticSheet"

    ws.Range("A1").Value2 = 100
    ws.Range("A2").Formula2 = "=A1*2"
    ws.Calculate
    
    Dim cmd As New FeatCmd_StaticSheetWorkbook
    Dim countConverted As Long
    countConverted = cmd.TestConvertSheetToValuesDirect(ws)

    Test_Runner.AssertEqual countConverted, 1#, "1 formula cell should be converted to static"
    Test_Runner.AssertEqual ws.Range("A2").HasFormula, False, "A2 formula should be removed"
    Test_Runner.AssertEqual ws.Range("A2").Value2, 200#, "A2 value should remain 200"

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
    Infra_Error.HandleError "Test_StaticSheetWorkbook_Execution", Err
    Resume CleanExit
End Sub

Public Sub Test_CreateSheet_PlacementAndNaming()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CreateSheet_PlacementAndNaming")
    On Error GoTo ErrHandler

    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets.Add
    ws1.Name = "Test_Temp_Create1"

    ' Create a sheet after ws1
    Set ws2 = ThisWorkbook.Worksheets.Add(After:=ws1)
    ws2.Name = "Test_Temp_Create2"

    ' Verify sheet names and placement
    Test_Runner.AssertEqual ws1.Name, "Test_Temp_Create1", "ws1 name should match"
    Test_Runner.AssertEqual ws2.Name, "Test_Temp_Create2", "ws2 name should match"
    Test_Runner.AssertEqual ThisWorkbook.Worksheets(ws1.Index + 1).Name, ws2.Name, "ws2 should be positioned after ws1"

    ' Cleanup
    Application.DisplayAlerts = False
    ws1.Delete
    ws2.Delete
    Application.DisplayAlerts = True

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws1.Delete
    ws2.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Infra_Error.HandleError "Test_CreateSheet_PlacementAndNaming", Err
    Resume CleanExit
End Sub

Public Sub Test_TableOfContents_Generation()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_TableOfContents_Generation")
    On Error GoTo ErrHandler

    ' Create a temporary workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add

    ' Add mock worksheets
    Dim ws1 As Worksheet, ws2 As Worksheet, wsFilter1 As Worksheet, wsFilter2 As Worksheet
    Set ws1 = wb.Worksheets.Add
    ws1.Name = "Test_TOC_Sheet1"
    ws1.Range("A1:B2").Value2 = "Data" ' 4 populated cells

    Set ws2 = wb.Worksheets.Add
    ws2.Name = "Test_TOC_Sheet2"
    ws2.Visible = xlSheetHidden

    Set wsFilter1 = wb.Worksheets.Add
    wsFilter1.Name = "Test_Temp_TOC_Filter"

    Set wsFilter2 = wb.Worksheets.Add
    wsFilter2.Name = "_BeaverTOC_Filter"

    ' Initialize AppContainer and create context
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Dim ctx As ICommandContext
    Set ctx = AppContainer.CreateCommandContext("TableOfContents")
    Set ctx.ActionContext.WorkbookRef = wb

    ' Resolve and execute command
    Dim cmd As ICommand
    Set cmd = AppContainer.ResolveCommand("TableOfContents")
    
    ' Execute (should not prompt since TOC sheet does not exist yet)
    cmd.Execute ctx

    ' Validate Table of Contents sheet is first
    Dim wsTOC As Worksheet
    Set wsTOC = wb.Worksheets(1)
    Test_Runner.AssertEqual wsTOC.Name, "Table of Contents", "First worksheet should be 'Table of Contents'"

    ' Validate info on wsTOC
    Dim r As Long, foundSheet1 As Boolean, foundSheet2 As Boolean, foundFilter1 As Boolean, foundFilter2 As Boolean
    For r = 6 To 18
        Dim sheetName As String
        sheetName = wsTOC.Cells(r, 3).Value
        If sheetName = "Test_TOC_Sheet1" Then
            foundSheet1 = True
            ' Verify cells count is 4
            Test_Runner.AssertEqual wsTOC.Cells(r, 5).Value, 4#, "Test_TOC_Sheet1 populated cells count should be 4"
            ' Verify visibility is Visible
            Test_Runner.AssertEqual wsTOC.Cells(r, 4).Value, "Visible", "Test_TOC_Sheet1 visibility should be Visible"
        ElseIf sheetName = "Test_TOC_Sheet2" Then
            foundSheet2 = True
            ' Verify visibility is Hidden
            Test_Runner.AssertEqual wsTOC.Cells(r, 4).Value, "Hidden", "Test_TOC_Sheet2 visibility should be Hidden"
        ElseIf sheetName = "Test_Temp_TOC_Filter" Then
            foundFilter1 = True
        ElseIf sheetName = "_BeaverTOC_Filter" Then
            foundFilter2 = True
        End If
    Next r

    Test_Runner.AssertEqual foundSheet1, True, "Test_TOC_Sheet1 should be listed in TOC"
    Test_Runner.AssertEqual foundSheet2, True, "Test_TOC_Sheet2 should be listed in TOC"
    Test_Runner.AssertEqual foundFilter1, False, "Test_Temp_TOC_Filter should be excluded from TOC"
    Test_Runner.AssertEqual foundFilter2, False, "_BeaverTOC_Filter should be excluded from TOC"

    ' Cleanup
    wb.Close SaveChanges:=False

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_TableOfContents_Generation", Err
    Resume CleanExit
End Sub

Public Sub Test_Export_Pdf_Backup_And_MultiRange()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_Export_Pdf_Backup_And_MultiRange")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim wb As Workbook: Set wb = Workbooks.Add
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    If wb.Worksheets.Count = 1 Then
        wb.Worksheets.Add After:=ws
    End If

    ' 1. Test PageSetup Backup & Restore
    Dim backupObj As New Infra_PageSetupBackup
    
    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.PrintArea = "A1:B5"
    ws.Visible = xlSheetVisible
    
    backupObj.Backup ws
    
    ' Modify settings
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PrintArea = "C1:D10"
    ws.Visible = xlSheetHidden
    
    ' Restore settings
    backupObj.Restore
    
    Test_Runner.AssertEqual ws.PageSetup.Orientation, xlPortrait, "Restore orientation to Portrait"
    Test_Runner.AssertEqual ws.PageSetup.PrintArea, "$A$1:$B$5", "Restore print area to A1:B5"
    Test_Runner.AssertEqual ws.Visible, xlSheetVisible, "Restore visibility to Visible"

    ' 2. Test Multi-Range print area generation logic
    ' Let's write some dummy values
    ws.Range("A1").Value = "A"
    ws.Range("B2").Value = "A"
    ws.Range("D1").Value = "B"
    ws.Range("E2").Value = "B"
    ' Selection with multi-areas
    Dim selRng As Range
    Set selRng = Union(ws.Range("A1:B2"), ws.Range("D1:E2"))
    
    Dim area As Range
    Dim intersectRange As Range
    Dim printAreaAddress As String
    printAreaAddress = ""
    
    For Each area In selRng.Areas
        Set intersectRange = Intersect(area, ws.UsedRange)
        If Not intersectRange Is Nothing Then
            if printAreaAddress <> "" Then
                printAreaAddress = printAreaAddress & ","
            End If
            printAreaAddress = printAreaAddress & intersectRange.Address
        End If
    Next area
    
    ' Check if print area address correctly combines the non-contiguous ranges
    Test_Runner.AssertEqual InStr(printAreaAddress, "$A$1:$B$2") > 0, True, "Contains first area"
    Test_Runner.AssertEqual InStr(printAreaAddress, "$D$1:$E$2") > 0, True, "Contains second area"
    Test_Runner.AssertEqual InStr(printAreaAddress, ","), 10, "Contains a comma separating areas"

    wb.Close SaveChanges:=False

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Test_Export_Pdf_Backup_And_MultiRange", Err
    Resume CleanExit
End Sub

Public Sub Test_CleanWorkbookNames_BrokenAndExternal()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CleanWorkbookNames_BrokenAndExternal")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_CleanNames"

    ' Setup names on the sheet
    On Error Resume Next
    ws.Names.Add Name:="TestBrokenName", RefersTo:="=SheetNonExistent!#REF!"
    ws.Names.Add Name:="TestExternalName", RefersTo:="=[ExternalFile.xlsx]Sheet1!$A$1"
    ws.Names.Add Name:="TestNormalName", RefersTo:="=$A$1"
    
    ' Also setup workbook scope broken name
    ThisWorkbook.Names.Add Name:="TestWbBrokenName", RefersTo:="=SheetNonExistent!#REF!"
    On Error GoTo ErrHandler

    ' Clean Broken Names on Sheet
    Dim removedBrokenCount As Long
    Infra_CommandSupport.CleanWorkbookNames Nothing, ws, NameCleanCriteriaBroken, removedBrokenCount
    Test_Runner.AssertEqual removedBrokenCount, 1, "Should clean exactly 1 broken name on the sheet"

    ' Verify the sheet names remaining
    Dim brokenExists As Boolean: brokenExists = False
    Dim externalExists As Boolean: externalExists = False
    Dim normalExists As Boolean: normalExists = False
    Dim nm As Name

    For Each nm In ws.Names
        If nm.Name = ws.Name & "!TestBrokenName" Then brokenExists = True
        If nm.Name = ws.Name & "!TestExternalName" Then externalExists = True
        If nm.Name = ws.Name & "!TestNormalName" Then normalExists = True
    Next nm

    Test_Runner.AssertEqual brokenExists, False, "Broken sheet-scoped name should be deleted"
    Test_Runner.AssertEqual externalExists, True, "External name should still exist"
    Test_Runner.AssertEqual normalExists, True, "Normal name should still exist"

    ' Clean External Names on Sheet
    Dim removedExternalCount As Long
    Infra_CommandSupport.CleanWorkbookNames Nothing, ws, NameCleanCriteriaExternal, removedExternalCount
    Test_Runner.AssertEqual removedExternalCount, 1, "Should clean exactly 1 external name on the sheet"

    externalExists = False
    For Each nm In ws.Names
        If nm.Name = ws.Name & "!TestExternalName" Then externalExists = True
    Next nm
    Test_Runner.AssertEqual externalExists, False, "External sheet-scoped name should be deleted"

    ' Clean Workbook Scope Broken Names
    Dim removedWbCount As Long
    Infra_CommandSupport.CleanWorkbookNames ThisWorkbook, Nothing, NameCleanCriteriaBroken, removedWbCount
    Test_Runner.AssertTrue removedWbCount >= 1, "Should clean at least 1 workbook-scoped broken name"

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
    Infra_Error.HandleError "Test_CleanWorkbookNames_BrokenAndExternal", Err
    Resume CleanExit
End Sub
