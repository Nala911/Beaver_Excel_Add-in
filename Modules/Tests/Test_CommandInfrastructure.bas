Attribute VB_Name = "Test_CommandInfrastructure"
Option Explicit

' @Module: Test_CommandInfrastructure
' @Category: Infrastructure
' @Description: Tests for manifest-driven command resolution and typed command context capture.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, Infra_CommandRegistry, AppContainer, Infra_Error

Public Sub Test_CommandRegistryResolvesRibbonEntries()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CommandRegistryResolvesRibbonEntries")
    On Error GoTo ErrHandler

    AssertEqual Infra_CommandRegistry.ResolveCommandName("UI_Ribbon.Ribbon_OnExportPng"), "ExportPng", "Ribbon export PNG callback should resolve to ExportPng"
    AssertEqual Infra_CommandRegistry.ResolveCommandName("UI_Ribbon.Ribbon_OnExportPdf"), "ExportPdf", "Ribbon export PDF callback should resolve to ExportPdf"
    AssertEqual Infra_CommandRegistry.ResolveCommandName("UI_Ribbon.Ribbon_OnShowHelpCenter"), "ShowHelpCenter", "Hotkeys help callback should resolve to ShowHelpCenter"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_CommandRegistryResolvesRibbonEntries", Err
    Resume CleanExit
End Sub

Public Sub Test_CommandRegistryResolvesHotkeyEntries()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CommandRegistryResolvesHotkeyEntries")
    On Error GoTo ErrHandler

    AssertEqual Infra_CommandRegistry.ResolveCommandName("UI_Hotkeys.Hotkey_FormatSelectedRange"), "FormatRange", "Format hotkey should resolve to FormatRange"
    AssertEqual Infra_CommandRegistry.ResolveCommandName("UI_Hotkeys.Hotkey_Delete"), "Delete", "Delete hotkey should resolve to Delete"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_CommandRegistryResolvesHotkeyEntries", Err
    Resume CleanExit
End Sub

Public Sub Test_CommandContextIncludesMetadataAndActionContext()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CommandContextIncludesMetadataAndActionContext")
    On Error GoTo ErrHandler

    Dim ctx As ICommandContext

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Set ctx = AppContainer.CreateCommandContext("FormatRange", "UI_Hotkeys.Hotkey_FormatSelectedRange", "Hotkey_FormatSelectedRange", "Hotkey")

    AssertEqual ctx.CommandName, "FormatRange", "Command context should store the resolved command name"
    AssertEqual ctx.EntryMacro, "UI_Hotkeys.Hotkey_FormatSelectedRange", "Command context should store the entry macro"
    AssertEqual ctx.TriggerKind, "Hotkey", "Command context should store the trigger kind"
    AssertEqual ctx.SourceName, "Hotkey_FormatSelectedRange", "Command context should store the source procedure"
    AssertTrue Not ctx.ActionContext Is Nothing, "Command context should capture an action context snapshot"
    AssertTrue Not ctx.ActionContext.WorkbookRef Is Nothing, "Action context should capture the active workbook"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_CommandContextIncludesMetadataAndActionContext", Err
    Resume CleanExit
End Sub

Public Sub Test_CommandRegistryCreatesKnownCommands()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CommandRegistryCreatesKnownCommands")
    On Error GoTo ErrHandler

    AssertTrue Not Infra_CommandRegistry.CreateCommand("FormatRange") Is Nothing, "Command registry should create FormatRange command instances"
    AssertTrue Not Infra_CommandRegistry.CreateCommand("ShowHelpCenter") Is Nothing, "Command registry should create ShowHelpCenter command instances"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_CommandRegistryCreatesKnownCommands", Err
    Resume CleanExit
End Sub

Public Sub Test_DiagnosticsEscapeJson()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_DiagnosticsEscapeJson")
    On Error GoTo ErrHandler

    Dim testStr As String
    testStr = "Hello " & Chr$(34) & "World" & Chr$(34) & " \ " & vbCrLf & vbTab
    
    Dim escaped As String
    escaped = Infra_Diagnostics.EscapeJson(testStr)
    
    Dim expected As String
    expected = "Hello \" & Chr$(34) & "World\" & Chr$(34) & " \\ \n\t"
    
    AssertEqual escaped, expected, "JSON escaping should correctly escape quotes, backslashes, newlines, and tabs"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_DiagnosticsEscapeJson", Err
    Resume CleanExit
End Sub

Public Sub Test_DiagnosticsLogsJSON()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_DiagnosticsLogsJSON")
    On Error GoTo ErrHandler

    Dim testMsg As String
    testMsg = "Test warning message " & Format$(Now, "hhnnss")
    
    ' Log warning
    Infra_Diagnostics.LogWarning "Test_DiagnosticsLogsJSON", testMsg
    
    ' Read log file
    Dim fso As Object
    Dim logPath As String
    Dim stream As Object
    Dim lastLine As String
    Dim lineText As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = Environ$("TEMP") & "\BeaverAddin_" & Infra_Diagnostics.GetPid() & ".log"
    
    AssertTrue fso.FileExists(logPath), "Log file should exist"
    
    Set stream = fso.OpenTextFile(logPath, 1)
    Do Until stream.AtEndOfStream
        lineText = stream.ReadLine
        If Trim$(lineText) <> "" Then
            lastLine = lineText
        End If
    Loop
    stream.Close
    
    ' Verify lastLine is valid JSON containing our testMsg
    AssertTrue Left$(lastLine, 1) = "{", "Log line should start with {"
    AssertTrue Right$(lastLine, 1) = "}", "Log line should end with }"
    AssertTrue InStr(lastLine, """event"":""warning""") > 0, "Log should contain warning event name"
    AssertTrue InStr(lastLine, """procedure"":""Test_DiagnosticsLogsJSON""") > 0, "Log should contain procedure name"
    AssertTrue InStr(lastLine, """message"":""" & testMsg & """") > 0, "Log should contain the message payload"

CleanExit:
    Set stream = Nothing
    Set fso = Nothing
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_DiagnosticsLogsJSON", Err
    Resume CleanExit
End Sub

