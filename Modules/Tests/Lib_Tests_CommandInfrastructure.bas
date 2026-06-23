Attribute VB_Name = "Lib_Tests_CommandInfrastructure"
Option Explicit

' @Module: Lib_Tests_CommandInfrastructure
' @Category: Infrastructure
' @Description: Tests for manifest-driven command resolution and typed command context capture.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, Infra_CommandRegistry, AppContainer, Infra_Error

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
