Attribute VB_Name = "Test_Feat_CreateNamedRanges"
Option Explicit

' @Module: Test_Feat_CreateNamedRanges
' @Category: Library
' @Description: Unit tests for Bulk Named Ranges feature.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, FeatCmd_CreateNamedRanges

Public Sub Test_CreateNamedRanges_Sanitization()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CreateNamedRanges_Sanitization")
    On Error GoTo ErrHandler

    Dim cmd As New FeatCmd_CreateNamedRanges
    
    ' Run assertions for sanitization logic
    Test_Runner.AssertEqual cmd.SanitizeExcelName("WACC %"), "WACC_", "WACC % should replace spaces and symbols with underscores and collapse them"
    Test_Runner.AssertEqual cmd.SanitizeExcelName("2026 Forecast"), "_2026_Forecast", "2026 Forecast should prepend underscore and replace space"
    Test_Runner.AssertEqual cmd.SanitizeExcelName("A1"), "A1_", "A1 coordinate conflict should append underscore"
    Test_Runner.AssertEqual cmd.SanitizeExcelName("Terminal Value - Year 5"), "Terminal_Value_Year_5", "Terminal Value hyphen and spaces should become a single collapsed underscore"
    Test_Runner.AssertEqual cmd.SanitizeExcelName("C"), "C_", "C coordinate conflict should append underscore"
    Test_Runner.AssertEqual cmd.SanitizeExcelName("R"), "R_", "R coordinate conflict should append underscore"
    Test_Runner.AssertEqual cmd.SanitizeExcelName("USD100"), "USD100_", "Valid-looking cell coordinate USD100 should append underscore"
    Test_Runner.AssertEqual cmd.SanitizeExcelName("  Discount Rate (WACC)  "), "Discount_Rate_WACC_", "Leading/trailing spaces trimmed, brackets/spaces collapsed to single underscore"
    
    ' Reserved word checks
    Test_Runner.AssertEqual cmd.SanitizeExcelName("Print_Area"), "Print_Area_", "Print_Area reserved word should append underscore"
    Test_Runner.AssertEqual cmd.SanitizeExcelName("Database"), "Database_", "Database reserved word should append underscore"
    Test_Runner.AssertEqual cmd.SanitizeExcelName("CRITERIA"), "CRITERIA_", "CRITERIA reserved word should append underscore"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_CreateNamedRanges_Sanitization", Err
    Resume CleanExit
End Sub

Public Sub Test_CreateNamedRanges_Execution()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CreateNamedRanges_Execution")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_NamedR"

    ' 1. Test Parallel Matching (Workbook Scope)
    ws.Range("A1").Value2 = "Assumption A"
    ws.Range("A2").Value2 = "Assumption B"
    ws.Range("A3").Value2 = "Assumption C"
    
    ws.Range("B1").Value2 = 100
    ws.Range("B2").Value2 = 200
    ws.Range("B3").Value2 = 300

    Dim cmd As New FeatCmd_CreateNamedRanges
    Dim createdCount As Long
    Dim createdList As String
    Dim success As Boolean

    success = cmd.ExecuteBulkNamedRangesDirect( _
        ws.Range("B1:B3"), ws.Range("A1:A3"), "Workbook", ThisWorkbook, ws, True, False, createdCount, createdList)

    Test_Runner.AssertEqual success, True, "Bulk naming parallel should return True"
    Test_Runner.AssertEqual createdCount, 3#, "Should have created 3 names"
    Test_Runner.AssertEqual createdList, "Assumption_A;Assumption_B;Assumption_C", "Lists of names should match"

    ' Assert workbook names were created
    Dim nameObj As Name
    Set nameObj = ThisWorkbook.Names("Assumption_A")
    Test_Runner.AssertEqual Not nameObj Is Nothing, True, "Assumption_A should exist in Workbook names"
    Test_Runner.AssertEqual nameObj.RefersToRange.Address, ws.Range("B1").Address, "Assumption_A should refer to B1"

    ' 2. Test Transposed Matching (Worksheet Scope)
    ' vertical labels, horizontal values
    ws.Range("A5").Value2 = "Label X"
    ws.Range("A6").Value2 = "Label Y"
    
    ws.Range("E5").Value2 = 500
    ws.Range("F5").Value2 = 600

    createdCount = 0
    createdList = ""
    success = cmd.ExecuteBulkNamedRangesDirect( _
        ws.Range("E5:F5"), ws.Range("A5:A6"), "Worksheet", ThisWorkbook, ws, True, False, createdCount, createdList)

    Test_Runner.AssertEqual success, True, "Bulk naming transposed should return True"
    Test_Runner.AssertEqual createdCount, 2#, "Should have created 2 names"
    
    ' Worksheet names are prefixed with worksheet name (e.g. Test_Temp_NamedR!Label_X)
    Dim expectedWorksheetNameRef As String
    expectedWorksheetNameRef = ws.Name & "!Label_X"
    
    Dim wsNameObj As Name
    On Error Resume Next
    Set wsNameObj = ws.Names("Label_X")
    On Error GoTo ErrHandler
    
    Test_Runner.AssertEqual Not wsNameObj Is Nothing, True, "Label_X should exist in Worksheet names"
    Test_Runner.AssertEqual wsNameObj.RefersToRange.Address, ws.Range("E5").Address, "Label_X should refer to E5"

    ' 3. Test Custom Undo Revert
    ' Register the pending undo
    Infra_Undo.SaveCreatedNamesState ThisWorkbook, "Assumption_A;Assumption_B;Assumption_C;" & ws.Name & "!Label_X;" & ws.Name & "!Label_Y"
    
    ' Execute undo
    Infra_Undo.PerformUndo

    ' Verify names are removed
    Dim nameDeletedCheck As Boolean
    
    On Error Resume Next
    Set nameObj = ThisWorkbook.Names("Assumption_A")
    nameDeletedCheck = (Err.Number <> 0 Or nameObj Is Nothing)
    On Error GoTo ErrHandler
    Test_Runner.AssertEqual nameDeletedCheck, True, "Assumption_A workbook name should be deleted on Undo"

    On Error Resume Next
    Set wsNameObj = ws.Names("Label_X")
    nameDeletedCheck = (Err.Number <> 0 Or wsNameObj Is Nothing)
    On Error GoTo ErrHandler
    Test_Runner.AssertEqual nameDeletedCheck, True, "Label_X worksheet name should be deleted on Undo"

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
    Infra_Error.HandleError "Test_CreateNamedRanges_Execution", Err
    Resume CleanExit
End Sub

Public Sub Test_CreateNamedRanges_SmartValidation()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CreateNamedRanges_SmartValidation")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_SmartR"

    Dim cmd As New FeatCmd_CreateNamedRanges
    Dim createdCount As Long
    Dim createdList As String
    Dim success As Boolean

    ' Set up initial state
    ' A1 = 10 (will be named "Dog")
    ' B1 = 20 (will try to name "Dog" -> should block/warn)
    ' A2 = 30 (has name "Cat", will change to "Cat_New" -> should rename)
    
    ws.Range("A1").Value2 = 10
    ws.Range("B1").Value2 = 20
    ws.Range("A2").Value2 = 30

    ' 1. Initially create name "Dog" for A1
    ws.Names.Add Name:="Dog", RefersTo:=ws.Range("A1")
    
    ' 2. Verify baseline naming
    Dim nameObj As Name
    On Error Resume Next
    Set nameObj = ws.Names("Dog")
    On Error GoTo ErrHandler
    Test_Runner.AssertEqual Not nameObj Is Nothing, True, "Dog should point to A1 initially"
    Test_Runner.AssertEqual nameObj.RefersToRange.Address, ws.Range("A1").Address, "Dog points to A1"

    ' 3. Test Skip Redundant: Try to name A1 as "Dog" again. It should skip it silently.
    createdCount = 0
    createdList = ""
    ws.Range("C1").Value2 = "Dog"
    success = cmd.ExecuteBulkNamedRangesDirect( _
        ws.Range("A1"), ws.Range("C1"), "Worksheet", ThisWorkbook, ws, True, False, createdCount, createdList)
        
    Test_Runner.AssertEqual success, True, "Skip redundant should return True"
    Test_Runner.AssertEqual createdCount, 0#, "Should have created 0 names since it was skipped"
    Test_Runner.AssertEqual createdList, "", "Created list should be empty"

    ' 4. Test Block Conflict: Try to name B1 as "Dog". It should not allow it.
    createdCount = 0
    createdList = ""
    ws.Range("C2").Value2 = "Dog"
    success = cmd.ExecuteBulkNamedRangesDirect( _
        ws.Range("B1"), ws.Range("C2"), "Worksheet", ThisWorkbook, ws, True, False, createdCount, createdList)
        
    Test_Runner.AssertEqual success, True, "Block conflict should return True (keeps executing other cells)"
    Test_Runner.AssertEqual createdCount, 0#, "Should have created 0 names due to block conflict"
    Test_Runner.AssertEqual createdList, "", "Created list should be empty"

    ' 5. Test Rename Clean: A2 has name "Cat", we rename A2 to "Cat_New"
    ws.Names.Add Name:="Cat", RefersTo:=ws.Range("A2")
    
    ' Check baseline "Cat" exists on A2
    Dim catName As Name
    On Error Resume Next
    Set catName = ws.Names("Cat")
    On Error GoTo ErrHandler
    Test_Runner.AssertEqual Not catName Is Nothing, True, "Cat should point to A2 initially"
    
    ' Now name A2 as "Cat_New"
    createdCount = 0
    createdList = ""
    ws.Range("C3").Value2 = "Cat_New"
    success = cmd.ExecuteBulkNamedRangesDirect( _
        ws.Range("A2"), ws.Range("C3"), "Worksheet", ThisWorkbook, ws, True, False, createdCount, createdList)
        
    Test_Runner.AssertEqual success, True, "Rename execution should succeed"
    Test_Runner.AssertEqual createdCount, 1#, "Should have created 1 name"
    Test_Runner.AssertEqual createdList, ws.Name & "!Cat_New", "Created name list should contain worksheet qualified Cat_New"
    
    ' Verify A2 is named "Cat_New"
    Dim catNewName As Name
    On Error Resume Next
    Set catNewName = ws.Names("Cat_New")
    On Error GoTo ErrHandler
    Test_Runner.AssertEqual Not catNewName Is Nothing, True, "Cat_New should exist"
    Test_Runner.AssertEqual catNewName.RefersToRange.Address, ws.Range("A2").Address, "Cat_New points to A2"
    
    ' Verify old name "Cat" was deleted to prevent duplicates
    On Error Resume Next
    Set catName = ws.Names("Cat")
    Dim catDeleted As Boolean: catDeleted = (Err.Number <> 0 Or catName Is Nothing)
    On Error GoTo ErrHandler
    Test_Runner.AssertEqual catDeleted, True, "Old name Cat should have been deleted"

    ' Verify other cell name "Dog" on A1 was not affected by the rename of A2
    Dim dogNameCheck As Name
    On Error Resume Next
    Set dogNameCheck = ws.Names("Dog")
    On Error GoTo ErrHandler
    Test_Runner.AssertEqual Not dogNameCheck Is Nothing, True, "Dog should still point to A1 after A2 is renamed"

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
    Infra_Error.HandleError "Test_CreateNamedRanges_SmartValidation", Err
    Resume CleanExit
End Sub

Public Sub Test_CreateNamedRanges_EmptyValueSkipping()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CreateNamedRanges_EmptyValueSkipping")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_EmptyVal"

    ' Setup labels in col A
    ws.Range("A1").Value2 = "Label_Valid"
    ws.Range("A2").Value2 = "Label_EmptyVal"
    ws.Range("A3").Value2 = "Label_SpaceVal"
    ws.Range("A4").Value2 = "Label_ErrorVal"

    ' Setup values in col B
    ws.Range("B1").Value2 = 42                ' Valid
    ' ws.Range("B2") remains completely empty
    ws.Range("B3").Value2 = "   "             ' Empty space string
    ws.Range("B4").Formula2 = "=1/0"          ' Formula Error (#DIV/0!)

    Dim cmd As New FeatCmd_CreateNamedRanges
    Dim createdCount As Long
    Dim createdList As String
    Dim success As Boolean

    success = cmd.ExecuteBulkNamedRangesDirect( _
        ws.Range("B1:B4"), ws.Range("A1:A4"), "Workbook", ThisWorkbook, ws, True, False, createdCount, createdList)

    Test_Runner.AssertEqual success, True, "Execution should succeed when skipping empty values"
    Test_Runner.AssertEqual createdCount, 2#, "Should have created exactly 2 named ranges (Valid and Error)"
    Test_Runner.AssertEqual createdList, "Label_Valid;Label_ErrorVal", "Should only name Label_Valid and Label_ErrorVal"

    ' Cleanup names
    On Error Resume Next
    ThisWorkbook.Names("Label_Valid").Delete
    ThisWorkbook.Names("Label_ErrorVal").Delete
    On Error GoTo ErrHandler

    ' Cleanup sheet
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
    Infra_Error.HandleError "Test_CreateNamedRanges_EmptyValueSkipping", Err
    Resume CleanExit
End Sub

Public Sub Test_CreateNamedRanges_FullColumnOptimized()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CreateNamedRanges_FullColumnOptimized")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_FullCol"

    ' Setup a few non-empty rows at the top
    ws.Range("A1").Value2 = "Assumption_X"
    ws.Range("A2").Value2 = ""                 ' Empty label
    ws.Range("A3").Value2 = "Assumption_Y"
    
    ws.Range("B1").Value2 = 100
    ws.Range("B2").Value2 = 200
    ws.Range("B3").Value2 = 300

    Dim cmd As New FeatCmd_CreateNamedRanges
    Dim createdCount As Long
    Dim createdList As String
    Dim success As Boolean

    ' Run bulk naming on full column ranges A:A and B:B
    success = cmd.ExecuteBulkNamedRangesDirect( _
        ws.Range("B:B"), ws.Range("A:A"), "Workbook", ThisWorkbook, ws, True, False, createdCount, createdList)

    Test_Runner.AssertEqual success, True, "Full column execution should succeed quickly"
    Test_Runner.AssertEqual createdCount, 2#, "Should have created exactly 2 named ranges (Assumption_X and Assumption_Y)"
    Test_Runner.AssertEqual createdList, "Assumption_X;Assumption_Y", "Only B1 and B3 named, row 2 skipped"

    ' Test Overwriting Collision (A1 is now renamed to Assumption_Z)
    ws.Range("A1").Value2 = "Assumption_Z"
    
    success = cmd.ExecuteBulkNamedRangesDirect( _
        ws.Range("B1"), ws.Range("A1"), "Workbook", ThisWorkbook, ws, True, False, createdCount, createdList)
        
    Test_Runner.AssertEqual success, True, "Overwriting should succeed"
    
    ' Old name Assumption_X should have been deleted, new name Assumption_Z created
    Dim nameDeletedCheck As Boolean
    Dim nameObj As Name
    On Error Resume Next
    Set nameObj = ThisWorkbook.Names("Assumption_X")
    nameDeletedCheck = (Err.Number <> 0 Or nameObj Is Nothing)
    On Error GoTo ErrHandler
    Test_Runner.AssertEqual nameDeletedCheck, True, "Old name Assumption_X should be deleted when overwritten"
    
    Dim newNameObj As Name
    On Error Resume Next
    Set newNameObj = ThisWorkbook.Names("Assumption_Z")
    On Error GoTo ErrHandler
    Test_Runner.AssertEqual Not newNameObj Is Nothing, True, "New name Assumption_Z should exist"

    ' Cleanup names
    On Error Resume Next
    ThisWorkbook.Names("Assumption_Z").Delete
    ThisWorkbook.Names("Assumption_Y").Delete
    On Error GoTo ErrHandler

    ' Cleanup sheet
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
    Infra_Error.HandleError "Test_CreateNamedRanges_FullColumnOptimized", Err
    Resume CleanExit
End Sub

Public Sub Test_CreateNamedRanges_ApplyNames()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CreateNamedRanges_ApplyNames")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Test_Temp_ApplyN"

    ws.Range("A1").Value2 = "Val_Rate"
    ws.Range("B1").Value2 = 0.05
    
    ws.Range("A2").Value2 = "Val_Principal"
    ws.Range("B2").Value2 = 1000
    
    ' Formula referencing B1 and B2 before naming
    ws.Range("B3").Formula2 = "=B2 * (1 + B1)"
    Test_Runner.AssertEqual ws.Range("B3").Formula2, "=B2 * (1 + B1)", "Initial formula has cell references"

    Dim cmd As New FeatCmd_CreateNamedRanges
    Dim createdCount As Long
    Dim createdList As String
    Dim success As Boolean

    ' Execute bulk named ranges
    success = cmd.ExecuteBulkNamedRangesDirect( _
        ws.Range("B1:B2"), ws.Range("A1:A2"), "Workbook", ThisWorkbook, ws, True, False, createdCount, createdList)

    Test_Runner.AssertEqual success, True, "Execution should succeed"
    Test_Runner.AssertEqual createdCount, 2#, "Should create 2 names"
    
    ' Check if formula in B3 was automatically updated to use named ranges
    Dim formulaResult As String: formulaResult = UCase$(ws.Range("B3").Formula2)
    Dim ratePresent As Boolean: ratePresent = (InStr(formulaResult, "VAL_RATE") > 0)
    Dim principalPresent As Boolean: principalPresent = (InStr(formulaResult, "VAL_PRINCIPAL") > 0)
    
    Test_Runner.AssertEqual ratePresent, True, "Formula should contain Val_Rate"
    Test_Runner.AssertEqual principalPresent, True, "Formula should contain Val_Principal"

    ' Cleanup names
    On Error Resume Next
    ThisWorkbook.Names("Val_Rate").Delete
    ThisWorkbook.Names("Val_Principal").Delete
    On Error GoTo ErrHandler

    ' Cleanup sheet
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
    Infra_Error.HandleError "Test_CreateNamedRanges_ApplyNames", Err
    Resume CleanExit
End Sub

