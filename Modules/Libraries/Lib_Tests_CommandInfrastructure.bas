Attribute VB_Name = "Lib_Tests_CommandInfrastructure"
Option Explicit

' @Module: Lib_Tests_CommandInfrastructure
' @Category: Infrastructure
' @Description: Tests for manifest-driven command resolution and typed command context capture.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, Infra_CommandCatalog, Infra_CommandRegistry, AppContainer, Infra_Error

Public Sub Test_CommandCatalogResolvesRibbonEntries()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CommandCatalogResolvesRibbonEntries")
    On Error GoTo ErrHandler

    AssertEqual Infra_CommandCatalog.ResolveCommandName("UI_Ribbon.Ribbon_OnExport"), "ExportImageOrPdf", "Ribbon export callback should resolve to ExportImageOrPdf"
    AssertEqual Infra_CommandCatalog.ResolveCommandName("UI_Ribbon.Ribbon_OnToggleFullScreen"), "ToggleFullScreen", "Ribbon focus-mode callback should resolve to ToggleFullScreen"
    AssertEqual Infra_CommandCatalog.ResolveCommandName("Infra_Hotkeys.ShowHotkeysHelp"), "ShowHotkeysHelp", "Hotkeys help callback should resolve to ShowHotkeysHelp"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_CommandCatalogResolvesRibbonEntries", Err
    Resume CleanExit
End Sub

Public Sub Test_CommandCatalogResolvesHotkeyEntries()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CommandCatalogResolvesHotkeyEntries")
    On Error GoTo ErrHandler

    AssertEqual Infra_CommandCatalog.ResolveCommandName("UI_Hotkeys.Hotkey_FormatSelectedRange"), "FormatRange", "Format hotkey should resolve to FormatRange"
    AssertEqual Infra_CommandCatalog.ResolveCommandName("UI_Hotkeys.Hotkey_Delete"), "Delete", "Delete hotkey should resolve to Delete"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_CommandCatalogResolvesHotkeyEntries", Err
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
    AssertTrue Not Infra_CommandRegistry.CreateCommand("ShowHotkeysHelp") Is Nothing, "Command registry should create ShowHotkeysHelp command instances"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Test_CommandRegistryCreatesKnownCommands", Err
    Resume CleanExit
End Sub
