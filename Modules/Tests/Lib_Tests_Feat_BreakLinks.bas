Attribute VB_Name = "Lib_Tests_Feat_BreakLinks"
Option Explicit

' @Module: Lib_Tests_Feat_BreakLinks
' @Category: Library
' @Description: Unit and integration tests for link breaking commands.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, AppContainer, Infra_Error, FeatCmd_BreakExternalLinks

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

Public Sub Test_BreakExternalLinks_SpillHandling()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_BreakExternalLinks_SpillHandling")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider

    ' 1. Create a temporary source workbook to avoid file picker dialog
    Dim sourceWb As Workbook
    Set sourceWb = Workbooks.Add
    sourceWb.Sheets(1).Range("A1").Value = "Apple"
    sourceWb.Sheets(1).Range("A2").Value = "Banana"
    sourceWb.Sheets(1).Range("A3").Value = "Cherry"
    Dim sourceWbName As String
    sourceWbName = sourceWb.Name

    ' 2. Create the target worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_BreakLinksSpill"

    ' 3. Set a dynamic array spill formula in the target sheet referencing the source workbook
    ws.Range("B2").Formula2 = "='[" & sourceWbName & "]" & sourceWb.Sheets(1).Name & "'!$A$1:$A$3"
    
    ' Force recalculation and wait to make sure the spill range is calculated and populated
    ws.Calculate
    Infra_ValueConversion.WaitForCalculation
    
    ' Check if it has a spill and correct values
    Lib_Tests.AssertEqual ws.Range("B2").Value, "Apple", "B2 should be Apple before link breaking"
    Lib_Tests.AssertEqual ws.Range("B3").Value, "Banana", "B3 should be Banana (spilled) before link breaking"
    Lib_Tests.AssertEqual ws.Range("B4").Value, "Cherry", "B4 should be Cherry (spilled) before link breaking"

    ' 4. Break the links
    Dim cmd As New FeatCmd_BreakExternalLinks
    Dim request As New ScopedRequest
    Dim context As ICommandContext
    
    Set context = AppContainer.CreateCommandContext("BreakExternalLinks")
    Set context.ActionContext.WorksheetRef = ws
    Set context.ActionContext.WorkbookRef = ThisWorkbook
    Set request.Context = context.ActionContext
    request.Scope = TargetScopeActiveSheet
    
    Dim stats As String
    stats = cmd.ExecuteBreakLinksDirect(ThisWorkbook, request, Empty)

    ' 5. Verify results
    ' Check that B2 formula is gone (replaced by static value)
    Lib_Tests.AssertEqual ws.Range("B2").HasFormula, False, "B2 should not have formula after link breaking"
    
    ' Check that all spilled values are preserved
    Lib_Tests.AssertEqual ws.Range("B2").Value, "Apple", "B2 value should be Apple after link breaking"
    Lib_Tests.AssertEqual ws.Range("B3").Value, "Banana", "B3 value should be Banana after link breaking"
    Lib_Tests.AssertEqual ws.Range("B4").Value, "Cherry", "B4 value should be Cherry after link breaking"
    
    ' Cleanup target worksheet
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Set ws = Nothing
    
    ' Close source workbook
    sourceWb.Close SaveChanges:=False
    Set sourceWb = Nothing

CleanExit:
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Infra_Error.HandleError "Test_BreakExternalLinks_SpillHandling", Err
    Resume CleanExit
End Sub
