Attribute VB_Name = "Lib_Tests_UI"
Option Explicit

' @Module: Lib_Tests_UI
' @Category: Library
' @Description: Unit and integration tests for userforms and UI components.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_Tests, AppContainer, Infra_Error, Infra_Hotkeys, Lib_UdfRegistry

Public Sub Test_UI_OptionPicker_DynamicLayout()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_UI_OptionPicker_DynamicLayout")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim frm As Object
    On Error Resume Next
    Set frm = VBA.UserForms.Add("UI_OptionPicker")
    On Error GoTo ErrHandler

    Dim status As Boolean
    status = Not frm Is Nothing
    Lib_Tests.AssertTrue status, "UI_OptionPicker form should be loadable"

    ' Configure single select option picker
    frm.ConfigureOptionPicker "Test Title", "Select an option from the list below:", "Option 2", Array("Option 1", "Option 2", "Option 3 With a Very Long Text to Test Sizing")

    ' Check layout size properties of the form and controls
    Dim lst As Object: Set lst = frm.Controls("lstHotkeys")
    Dim lblPrompt As Object: Set lblPrompt = frm.Controls("lblPrompt")
    Dim btnOK As Object: Set btnOK = frm.Controls("btnOK")
    Dim btnCancel As Object: Set btnCancel = frm.Controls("btnCancel")

    Lib_Tests.AssertTrue Not lst Is Nothing, "ListBox control lstHotkeys should exist"
    Lib_Tests.AssertTrue Not lblPrompt Is Nothing, "Label control lblPrompt should exist"
    Lib_Tests.AssertTrue Not btnOK Is Nothing, "Button control btnOK should exist"
    Lib_Tests.AssertTrue Not btnCancel Is Nothing, "Button control btnCancel should exist"

    ' Assertions on visibility
    Lib_Tests.AssertTrue lblPrompt.Visible = False, "Label control lblPrompt should be invisible"
    Lib_Tests.AssertTrue btnOK.Visible = False, "Button control btnOK should be invisible"
    Lib_Tests.AssertTrue btnCancel.Visible = False, "Button control btnCancel should be invisible"

    ' Assertions on dimensions
    Lib_Tests.AssertTrue frm.Width > 200, "Form width should be scaled dynamically"
    Lib_Tests.AssertTrue frm.Height > 50, "Form height should be scaled dynamically"
    Lib_Tests.AssertTrue lst.Width > 180, "ListBox width should be scaled to fit options"

    ' Test multi-select option picker configuration
    frm.ConfigureMultiOptionPicker "Test Multi Title", "Check the options:", Array("Opt A", "Opt B"), Array(True, False)

    Lib_Tests.AssertTrue lst.MultiSelect = 0, "ListBox should be set to single-select custom checkbox list mode"
    Lib_Tests.AssertTrue btnOK.Visible = True, "OK button should be visible in multi-select mode"
    Lib_Tests.AssertTrue btnCancel.Visible = True, "Cancel button should be visible in multi-select mode"

    ' 1. Test ToggleItem (from Checked to Unchecked)
    frm.TestToggleItem 0
    Lib_Tests.AssertEqual Left$(lst.List(0), 1), ChrW$(9744), "Opt A should be unchecked after toggle"
    
    ' 2. Test ToggleItem (from Unchecked to Checked)
    frm.TestToggleItem 0
    Lib_Tests.AssertEqual Left$(lst.List(0), 1), ChrW$(9745), "Opt A should be checked again after toggle"

    ' 3. Test mutual exclusivity with custom exclusivePrefixes
    Dim optArray As Variant
    optArray = Array("Group 1: Val A", "Group 1: Val B", "Group 2: Val A")
    frm.ConfigureMultiOptionPicker "Test Exclusivity", "Check options:", optArray, Array(True, False, False), , Array("Group 1:")
    
    ' Toggle second item in Group 1 (should uncheck the first item in Group 1)
    frm.TestToggleItem 1
    Lib_Tests.AssertEqual Left$(lst.List(0), 1), ChrW$(9744), "Group 1: Val A should be unchecked"
    Lib_Tests.AssertEqual Left$(lst.List(1), 1), ChrW$(9745), "Group 1: Val B should be checked"
    
    ' Toggle third item (not in Group 1, but starts with Group 2) - should not affect Group 1: Val B
    frm.TestToggleItem 2
    Lib_Tests.AssertEqual Left$(lst.List(1), 1), ChrW$(9745), "Group 1: Val B should remain checked"
    Lib_Tests.AssertEqual Left$(lst.List(2), 1), ChrW$(9745), "Group 2: Val A should be checked"

CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_UI_OptionPicker_DynamicLayout", Err
    Resume CleanExit
End Sub

Public Sub Test_UI_OptionPicker_KeyboardNavigation()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_UI_OptionPicker_KeyboardNavigation")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim frm As Object
    On Error Resume Next
    Set frm = VBA.UserForms.Add("UI_OptionPicker")
    On Error GoTo ErrHandler

    Dim status As Boolean
    status = Not frm Is Nothing
    Lib_Tests.AssertTrue status, "UI_OptionPicker form should be loadable"

    ' Configure single select option picker
    frm.ConfigureOptionPicker "Test Keyboard Title", "Select an option:", "Option 1", Array("Option 1", "Option 2")

    ' Check initial state
    Lib_Tests.AssertTrue Not frm.IsIgnoringClick, "Initial IsIgnoringClick should be False"
    Lib_Tests.AssertTrue Not frm.WasConfirmed, "Initial WasConfirmed should be False"

    ' Simulate Arrow Down key down (KeyCode = 40)
    frm.HandleKeyDown 40, 0
    Lib_Tests.AssertTrue frm.IsIgnoringClick, "IsIgnoringClick should be True after key down (arrow key)"

    ' Simulate Arrow Down key up
    frm.HandleKeyUp 40, 0
    Lib_Tests.AssertTrue Not frm.IsIgnoringClick, "IsIgnoringClick should be False after key up"

    ' Simulate Enter key down (KeyCode = 13)
    frm.HandleKeyDown 13, 0
    Lib_Tests.AssertTrue frm.WasConfirmed, "WasConfirmed should be True after Enter key down"

CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_UI_OptionPicker_KeyboardNavigation", Err
    Resume CleanExit
End Sub

Public Sub Test_UdfRegistry_And_HelpCenter()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_UdfRegistry_And_HelpCenter")
    On Error GoTo ErrHandler

    ' 1. Test GetAllUdfs returns expected items
    Dim udfs As Collection
    Set udfs = Lib_UdfRegistry.GetAllUdfs()
    Lib_Tests.AssertTrue Not udfs Is Nothing, "UDF registry collection should not be Nothing"
    Lib_Tests.AssertTrue udfs.Count > 0, "UDF registry should contain at least one UDF"

    Dim meta As Object
    Set meta = udfs(1)
    Lib_Tests.AssertEqual meta("Name"), "XFilter", "First UDF name should be XFilter"
    Lib_Tests.AssertEqual meta("Category"), "User Defined", "XFilter category should be User Defined"
    Lib_Tests.AssertTrue IsArray(meta("ArgumentDescriptions")), "XFilter argument descriptions should be an array"

    ' 2. Verify ShowHelpCenter runs without error in headless mode (display bypassed)
    Infra_Hotkeys.ShowHelpCenter
    
    ' Assert that we completed without error
    Lib_Tests.AssertTrue True, "ShowHelpCenter completed without error in headless mode"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_UdfRegistry_And_HelpCenter", Err
    Resume CleanExit
End Sub

Public Sub Test_CommandResolution_NewMenus()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Test_CommandResolution_NewMenus")
    On Error GoTo ErrHandler

    ' 1. Test ShowHelpCenter command resolution
    Dim cmdHelp As ICommand
    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Set cmdHelp = AppContainer.ResolveCommand("ShowHelpCenter")
    Lib_Tests.AssertEqual Not cmdHelp Is Nothing, True, "ShowHelpCenter command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHelp), "FeatCmd_ShowHelpCenter", "ShowHelpCenter should resolve to FeatCmd_ShowHelpCenter"

    ' 2. Test Highlight sub-commands
    Dim cmdHighlight As ICommand
    Set cmdHighlight = AppContainer.ResolveCommand("HighlightInconsistentFormulas")
    Lib_Tests.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightInconsistentFormulas should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightInconsistentFormulas should resolve to FeatCmd_HighlightData"

    Set cmdHighlight = AppContainer.ResolveCommand("HighlightDuplicates")
    Lib_Tests.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightDuplicates should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightDuplicates should resolve to FeatCmd_HighlightData"

    Set cmdHighlight = AppContainer.ResolveCommand("HighlightErrors")
    Lib_Tests.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightErrors should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightErrors should resolve to FeatCmd_HighlightData"

    Set cmdHighlight = AppContainer.ResolveCommand("HighlightHardcodedValues")
    Lib_Tests.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightHardcodedValues should resolve"
    Lib_Tests.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightHardcodedValues should resolve to FeatCmd_HighlightData"

    ' 3. Test Export sub-commands
    Dim cmdExport As ICommand
    Set cmdExport = AppContainer.ResolveCommand("ExportPng")
    Lib_Tests.AssertEqual Not cmdExport Is Nothing, True, "ExportPng should resolve"
    Lib_Tests.AssertEqual TypeName(cmdExport), "FeatCmd_ExportImageOrPdf", "ExportPng should resolve to FeatCmd_ExportImageOrPdf"

    Set cmdExport = AppContainer.ResolveCommand("ExportPdf")
    Lib_Tests.AssertEqual Not cmdExport Is Nothing, True, "ExportPdf should resolve"
    Lib_Tests.AssertEqual TypeName(cmdExport), "FeatCmd_ExportImageOrPdf", "ExportPdf should resolve to FeatCmd_ExportImageOrPdf"

    ' 4. Test new ModifyData commands
    Dim cmdNewModify As ICommand
    Set cmdNewModify = AppContainer.ResolveCommand("UnmergeFill")
    Lib_Tests.AssertEqual Not cmdNewModify Is Nothing, True, "UnmergeFill command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdNewModify), "FeatCmd_UnmergeFill", "UnmergeFill should resolve to FeatCmd_UnmergeFill"

    Set cmdNewModify = AppContainer.ResolveCommand("ForceNumber")
    Lib_Tests.AssertEqual Not cmdNewModify Is Nothing, True, "ForceNumber command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdNewModify), "FeatCmd_ForceNumber", "ForceNumber should resolve to FeatCmd_ForceNumber"

    ' 5. Test Duplicate command resolution
    Dim cmdDuplicate As ICommand
    Set cmdDuplicate = AppContainer.ResolveCommand("Duplicate")
    Lib_Tests.AssertEqual Not cmdDuplicate Is Nothing, True, "Duplicate command should resolve"
    Lib_Tests.AssertEqual TypeName(cmdDuplicate), "FeatCmd_Duplicate", "Duplicate should resolve to FeatCmd_Duplicate"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_CommandResolution_NewMenus", Err
    Resume CleanExit
End Sub
