Attribute VB_Name = "Lib_Tests"
Option Explicit

' @Module: Lib_Tests
' @Category: Infrastructure
' @Description: Lightweight unit testing framework with structured result export for automated verification.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Config, Infra_Hotkeys, Lib_TestManifest

Private pResults As Collection
Private pSuiteStartTime As Double
Private Const TEST_RESULTS_FILE_NAME As String = "BeaverAddin.TestResults.tsv"

' --- PUBLIC INTERFACE ---

' Main entry point for the test runner.
' Executes the generated test manifest built during Update.ps1.
Public Sub RunAllTests()
    Dim tracker As Object: Set tracker = Infra_Error.Track("RunAllTests")
    On Error GoTo ErrHandler
    
    Set pResults = New Collection
    pSuiteStartTime = Timer
    Debug.Print "--- BEAVER ADD-IN: STARTING UNIT TESTS ---"
    ClearPersistedResults

    Lib_TestManifest.RunGeneratedTests
    
    ReportResults

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "RunAllTests", Err
    Resume CleanExit
End Sub

' --- TEST SUITES ---

Public Sub Test_Infrastructure_Basics()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_Infrastructure_Basics")
    On Error GoTo ErrHandler
    
    ' Example Test: Verify Config is loaded
    AssertNotEqual Infra_Config.ADDIN_NAME, "", "Addin Name should not be empty"
    AssertNotEqual Infra_Config.ADDIN_VERSION, "", "Addin Version should not be empty"
    AssertNotEqual Infra_Config.RELEASE_TIER, "", "Release tier should not be empty"

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "Test_Infrastructure_Basics", Err
    Resume CleanExit
End Sub

Public Sub Test_ConfigProvidesTypedHotkeys()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_ConfigProvidesTypedHotkeys")
    On Error GoTo ErrHandler

    AssertTrue Not Infra_Config.Hotkeys Is Nothing, "Hotkey collection should be available"
    AssertTrue Infra_Config.Hotkeys.Count > 0, "Hotkey collection should not be empty"

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "Test_ConfigProvidesTypedHotkeys", Err
    Resume CleanExit
End Sub

Public Sub Test_TranslateHotkeyHandlesModifiers()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_TranslateHotkeyHandlesModifiers")
    On Error GoTo ErrHandler

    AssertEqual Infra_Hotkeys.TranslateHotkey("^+p"), "Ctrl + Shift + P", "TranslateHotkey should expand modifiers"
    AssertEqual Infra_Hotkeys.TranslateHotkey("{DELETE}"), "DELETE", "TranslateHotkey should normalize braced keys"

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "Test_TranslateHotkeyHandlesModifiers", Err
    Resume CleanExit
End Sub

' --- ASSERTIONS ---

Public Sub AssertTrue(ByVal condition As Boolean, ByVal msg As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("AssertTrue")
    On Error GoTo ErrHandler
    
    LogResult "AssertTrue", condition, msg

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "AssertTrue", Err
    Resume CleanExit
End Sub

Public Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal msg As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("AssertEqual")
    On Error GoTo ErrHandler
    
    LogResult "AssertEqual", (actual = expected), msg & " (Expected: " & expected & ", Actual: " & actual & ")"

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "AssertEqual", Err
    Resume CleanExit
End Sub

Public Sub AssertNotEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal msg As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("AssertNotEqual")
    On Error GoTo ErrHandler
    
    LogResult "AssertNotEqual", (actual <> expected), msg

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "AssertNotEqual", Err
    Resume CleanExit
End Sub

' --- PRIVATE HELPERS ---

Private Sub LogResult(ByVal testName As String, ByVal Passed As Boolean, ByVal msg As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("LogResult")
    On Error GoTo ErrHandler
    
    Dim res As New Infra_TestResult
    res.Name = testName
    res.Passed = Passed
    res.Message = msg
    res.DurationMs = CLng(TimerElapsedMilliseconds(pSuiteStartTime))
    res.Category = IIf(Passed, "pass", "fail")
    
    If pResults Is Nothing Then Set pResults = New Collection
    pResults.Add res
    
    If Passed Then
        Debug.Print "  [PASS] " & msg
    Else
        Debug.Print "  [FAIL] " & msg
    End If

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "LogResult", Err
    Resume CleanExit
End Sub

Private Sub ReportResults()
    Dim tracker As Object: Set tracker = Infra_Error.Track("ReportResults")
    On Error GoTo ErrHandler
    
    Dim passCount As Long, failCount As Long
    Dim res As Variant
    
    For Each res In pResults
        If res.Passed Then passCount = passCount + 1 Else failCount = failCount + 1
    Next res
    
    Debug.Print "--- TEST SUMMARY ---"
    Debug.Print "Total: " & pResults.Count
    Debug.Print "Passed: " & passCount
    Debug.Print "Failed: " & failCount
    Debug.Print "--------------------"
    PersistResults pResults, passCount, failCount
    
    If failCount > 0 Then
        Err.Raise 9999, "Lib_Tests", "Unit tests failed. See Immediate Window for details."
    End If

CleanExit:
    Exit Sub

ErrHandler:
    HandleError "ReportResults", Err
    Resume CleanExit
End Sub

Private Sub ClearPersistedResults()
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(GetResultsFilePath()) Then
        fso.DeleteFile GetResultsFilePath(), True
    End If
    On Error GoTo 0
End Sub

Private Sub PersistResults(ByVal results As Collection, ByVal passCount As Long, ByVal failCount As Long)
    Dim tracker As Object: Set tracker = Infra_Error.Track("PersistResults")
    On Error GoTo ErrHandler

    Dim fso As Object
    Dim stream As Object
    Dim res As Infra_TestResult
    Dim totalCount As Long

    If results Is Nothing Then GoTo CleanExit

    totalCount = results.Count
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.OpenTextFile(GetResultsFilePath(), 2, True)
    stream.WriteLine "SUMMARY" & vbTab & totalCount & vbTab & passCount & vbTab & failCount

    For Each res In results
        stream.WriteLine "RESULT" & vbTab & EscapeTab(res.Name) & vbTab & CStr(res.Passed) & vbTab & res.DurationMs & vbTab & EscapeTab(res.Category) & vbTab & EscapeTab(res.Message)
    Next res

    stream.Close

CleanExit:
    Exit Sub
ErrHandler:
    HandleErrorDetailed "PersistResults", Err, Nothing, Infra_Error.CATEGORY_TEST, Infra_Error.SEVERITY_WARNING
    Resume CleanExit
End Sub

Private Function GetResultsFilePath() As String
    GetResultsFilePath = Environ$("TEMP") & "\" & TEST_RESULTS_FILE_NAME
End Function

Private Function EscapeTab(ByVal value As String) As String
    EscapeTab = Replace(value, vbTab, " ")
    EscapeTab = Replace(EscapeTab, vbCrLf, " ")
    EscapeTab = Replace(EscapeTab, vbCr, " ")
    EscapeTab = Replace(EscapeTab, vbLf, " ")
End Function

Private Function TimerElapsedMilliseconds(ByVal startedAt As Double) As Double
    TimerElapsedMilliseconds = (Timer - startedAt) * 1000#
    If TimerElapsedMilliseconds < 0 Then
        TimerElapsedMilliseconds = TimerElapsedMilliseconds + (86400# * 1000#)
    End If
End Function
