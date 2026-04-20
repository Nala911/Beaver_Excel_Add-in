Attribute VB_Name = "Lib_Tests_DateConversion"
Option Explicit

' @Module: Lib_Tests_DateConversion
' @Category: Infrastructure
' @Description: Unit tests for robust date parsing logic.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, Lib_DateUtils

Public Sub Test_ResolveDate_UserCase()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ResolveDate_UserCase")
    On Error GoTo ErrHandler
    
    Dim d As Date
    Dim success As Boolean
    
    ' Case: 01-09-1990 with target month 9
    success = Lib_DateUtils.TryResolveDateWithMonth("01-09-1990", 9, d)
    
    Lib_Tests.AssertTrue success, "Should resolve 01-09-1990"
    Lib_Tests.AssertEqual Day(d), 1, "Day should be 1"
    Lib_Tests.AssertEqual Month(d), 9, "Month should be 9"
    Lib_Tests.AssertEqual Year(d), 1990, "Year should be 1990"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_ResolveDate_UserCase", Err
    Resume CleanExit
End Sub

Public Sub Test_ResolveDate_TextMonth()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ResolveDate_TextMonth")
    On Error GoTo ErrHandler
    
    Dim d As Date
    Dim success As Boolean
    
    ' Case: 01-Sep-1990 with target month 9
    success = Lib_DateUtils.TryResolveDateWithMonth("01-Sep-1990", 9, d)
    
    Lib_Tests.AssertTrue success, "Should resolve 01-Sep-1990"
    Lib_Tests.AssertEqual Day(d), 1, "Day should be 1"
    Lib_Tests.AssertEqual Month(d), 9, "Month should be 9"
    Lib_Tests.AssertEqual Year(d), 1990, "Year should be 1990"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_ResolveDate_TextMonth", Err
    Resume CleanExit
End Sub

Public Sub Test_ResolveDate_Ambiguous_DMY()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ResolveDate_Ambiguous_DMY")
    On Error GoTo ErrHandler
    
    Dim d As Date
    Dim success As Boolean
    
    ' Case: 01/02/03 with target month 2
    success = Lib_DateUtils.TryResolveDateWithMonth("01/02/03", 2, d)
    
    Lib_Tests.AssertTrue success, "Should resolve 01/02/03"
    Lib_Tests.AssertEqual Month(d), 2, "Month should be 2"
    Lib_Tests.AssertEqual Day(d), 1, "Day should be 1 (DD/MM/YY heuristic)"
    Lib_Tests.AssertEqual Year(d), 2003, "Year should be 2003"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_ResolveDate_Ambiguous_DMY", Err
    Resume CleanExit
End Sub

Public Sub Test_ResolveDate_Ambiguous_MDY()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ResolveDate_Ambiguous_MDY")
    On Error GoTo ErrHandler
    
    Dim d As Date
    Dim success As Boolean
    
    ' Case: 01/02/03 with target month 1
    success = Lib_DateUtils.TryResolveDateWithMonth("01/02/03", 1, d)
    
    Lib_Tests.AssertTrue success, "Should resolve 01/02/03"
    Lib_Tests.AssertEqual Month(d), 1, "Month should be 1"
    Lib_Tests.AssertEqual Day(d), 2, "Day should be 2 (MM/DD/YY heuristic)"
    Lib_Tests.AssertEqual Year(d), 2003, "Year should be 2003"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_ResolveDate_Ambiguous_MDY", Err
    Resume CleanExit
End Sub
