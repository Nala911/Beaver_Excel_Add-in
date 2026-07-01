Attribute VB_Name = "Test_UI"
Option Explicit

' @Module: Test_UI
' @Category: Library
' @Description: Unit and integration tests for userforms and UI components.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Test_Runner, AppContainer, Infra_Error, Infra_Hotkeys, Lib_UdfRegistry

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
    Test_Runner.AssertTrue status, "UI_OptionPicker form should be loadable"

    ' Configure single select option picker
    frm.ConfigureOptionPicker "Test Title", "Select an option from the list below:", "Option 2", Array("Option 1", "Option 2", "Option 3 With a Very Long Text to Test Sizing")

    ' Check layout size properties of the form and controls
    Dim lst As Object: Set lst = frm.Controls("lstHotkeys")
    Dim lblPrompt As Object: Set lblPrompt = frm.Controls("lblPrompt")
    Dim btnOK As Object: Set btnOK = frm.Controls("btnOK")
    Dim btnCancel As Object: Set btnCancel = frm.Controls("btnCancel")

    Test_Runner.AssertTrue Not lst Is Nothing, "ListBox control lstHotkeys should exist"
    Test_Runner.AssertTrue Not lblPrompt Is Nothing, "Label control lblPrompt should exist"
    Test_Runner.AssertTrue Not btnOK Is Nothing, "Button control btnOK should exist"
    Test_Runner.AssertTrue Not btnCancel Is Nothing, "Button control btnCancel should exist"

    ' Assertions on visibility
    Test_Runner.AssertTrue lblPrompt.Visible = False, "Label control lblPrompt should be invisible"
    Test_Runner.AssertTrue btnOK.Visible = False, "Button control btnOK should be invisible"
    Test_Runner.AssertTrue btnCancel.Visible = False, "Button control btnCancel should be invisible"

    ' Assertions on dimensions
    Test_Runner.AssertTrue frm.Width > 200, "Form width should be scaled dynamically"
    Test_Runner.AssertTrue frm.Height > 50, "Form height should be scaled dynamically"
    Test_Runner.AssertTrue lst.Width > 180, "ListBox width should be scaled to fit options"

    ' Test multi-select option picker configuration
    frm.ConfigureMultiOptionPicker "Test Multi Title", "Check the options:", Array("Opt A", "Opt B"), Array(True, False)

    Test_Runner.AssertTrue lst.MultiSelect = 0, "ListBox should be set to single-select custom checkbox list mode"
    Test_Runner.AssertTrue btnOK.Visible = True, "OK button should be visible in multi-select mode"
    Test_Runner.AssertTrue btnCancel.Visible = True, "Cancel button should be visible in multi-select mode"

    ' 1. Test ToggleItem (from Checked to Unchecked)
    frm.TestToggleItem 0
    Test_Runner.AssertEqual Left$(lst.List(0), 1), ChrW$(9744), "Opt A should be unchecked after toggle"
    
    ' 2. Test ToggleItem (from Unchecked to Checked)
    frm.TestToggleItem 0
    Test_Runner.AssertEqual Left$(lst.List(0), 1), ChrW$(9745), "Opt A should be checked again after toggle"

    ' 3. Test mutual exclusivity with custom exclusivePrefixes
    Dim optArray As Variant
    optArray = Array("Group 1: Val A", "Group 1: Val B", "Group 2: Val A")
    frm.ConfigureMultiOptionPicker "Test Exclusivity", "Check options:", optArray, Array(True, False, False), , Array("Group 1:")
    
    ' Toggle second item in Group 1 (should uncheck the first item in Group 1)
    frm.TestToggleItem 1
    Test_Runner.AssertEqual Left$(lst.List(0), 1), ChrW$(9744), "Group 1: Val A should be unchecked"
    Test_Runner.AssertEqual Left$(lst.List(1), 1), ChrW$(9745), "Group 1: Val B should be checked"
    
    ' Toggle third item (not in Group 1, but starts with Group 2) - should not affect Group 1: Val B
    frm.TestToggleItem 2
    Test_Runner.AssertEqual Left$(lst.List(1), 1), ChrW$(9745), "Group 1: Val B should remain checked"
    Test_Runner.AssertEqual Left$(lst.List(2), 1), ChrW$(9745), "Group 2: Val A should be checked"

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
    Test_Runner.AssertTrue status, "UI_OptionPicker form should be loadable"

    ' Configure single select option picker
    frm.ConfigureOptionPicker "Test Keyboard Title", "Select an option:", "Option 1", Array("Option 1", "Option 2")

    ' Check initial state
    Test_Runner.AssertTrue Not frm.IsIgnoringClick, "Initial IsIgnoringClick should be False"
    Test_Runner.AssertTrue Not frm.WasConfirmed, "Initial WasConfirmed should be False"

    ' Simulate Arrow Down key down (KeyCode = 40)
    frm.HandleKeyDown 40, 0
    Test_Runner.AssertTrue frm.IsIgnoringClick, "IsIgnoringClick should be True after key down (arrow key)"

    ' Simulate Arrow Down key up
    frm.HandleKeyUp 40, 0
    Test_Runner.AssertTrue Not frm.IsIgnoringClick, "IsIgnoringClick should be False after key up"

    ' Simulate Enter key down (KeyCode = 13)
    frm.HandleKeyDown 13, 0
    Test_Runner.AssertTrue frm.WasConfirmed, "WasConfirmed should be True after Enter key down"

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
    Test_Runner.AssertTrue Not udfs Is Nothing, "UDF registry collection should not be Nothing"
    Test_Runner.AssertTrue udfs.Count > 0, "UDF registry should contain at least one UDF"

    Dim meta As Object
    Set meta = udfs(1)
    Test_Runner.AssertEqual meta("Name"), "XFilter", "First UDF name should be XFilter"
    Test_Runner.AssertEqual meta("Category"), "User Defined", "XFilter category should be User Defined"
    Test_Runner.AssertTrue IsArray(meta("ArgumentDescriptions")), "XFilter argument descriptions should be an array"

    ' 2. Verify ShowHelpCenter runs without error in headless mode (display bypassed)
    Infra_Hotkeys.ShowHelpCenter
    
    ' Assert that we completed without error
    Test_Runner.AssertTrue True, "ShowHelpCenter completed without error in headless mode"

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
    Test_Runner.AssertEqual Not cmdHelp Is Nothing, True, "ShowHelpCenter command should resolve"
    Test_Runner.AssertEqual TypeName(cmdHelp), "FeatCmd_ShowHelpCenter", "ShowHelpCenter should resolve to FeatCmd_ShowHelpCenter"

    ' 2. Test Highlight sub-commands
    Dim cmdHighlight As ICommand
    Set cmdHighlight = AppContainer.ResolveCommand("HighlightInconsistentFormulas")
    Test_Runner.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightInconsistentFormulas should resolve"
    Test_Runner.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightInconsistentFormulas should resolve to FeatCmd_HighlightData"

    Set cmdHighlight = AppContainer.ResolveCommand("HighlightDuplicates")
    Test_Runner.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightDuplicates should resolve"
    Test_Runner.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightDuplicates should resolve to FeatCmd_HighlightData"

    Set cmdHighlight = AppContainer.ResolveCommand("HighlightErrors")
    Test_Runner.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightErrors should resolve"
    Test_Runner.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightErrors should resolve to FeatCmd_HighlightData"

    Set cmdHighlight = AppContainer.ResolveCommand("HighlightHardcodedValues")
    Test_Runner.AssertEqual Not cmdHighlight Is Nothing, True, "HighlightHardcodedValues should resolve"
    Test_Runner.AssertEqual TypeName(cmdHighlight), "FeatCmd_HighlightData", "HighlightHardcodedValues should resolve to FeatCmd_HighlightData"

    ' 3. Test Export sub-commands
    Dim cmdExport As ICommand
    Set cmdExport = AppContainer.ResolveCommand("ExportPng")
    Test_Runner.AssertEqual Not cmdExport Is Nothing, True, "ExportPng should resolve"
    Test_Runner.AssertEqual TypeName(cmdExport), "FeatCmd_ExportImageOrPdf", "ExportPng should resolve to FeatCmd_ExportImageOrPdf"

    Set cmdExport = AppContainer.ResolveCommand("ExportPdf")
    Test_Runner.AssertEqual Not cmdExport Is Nothing, True, "ExportPdf should resolve"
    Test_Runner.AssertEqual TypeName(cmdExport), "FeatCmd_ExportImageOrPdf", "ExportPdf should resolve to FeatCmd_ExportImageOrPdf"

    ' 4. Test new ModifyData commands
    Dim cmdNewModify As ICommand
    Set cmdNewModify = AppContainer.ResolveCommand("UnmergeFill")
    Test_Runner.AssertEqual Not cmdNewModify Is Nothing, True, "UnmergeFill command should resolve"
    Test_Runner.AssertEqual TypeName(cmdNewModify), "FeatCmd_UnmergeFill", "UnmergeFill should resolve to FeatCmd_UnmergeFill"

    Set cmdNewModify = AppContainer.ResolveCommand("ForceNumber")
    Test_Runner.AssertEqual Not cmdNewModify Is Nothing, True, "ForceNumber command should resolve"
    Test_Runner.AssertEqual TypeName(cmdNewModify), "FeatCmd_ForceNumber", "ForceNumber should resolve to FeatCmd_ForceNumber"

    ' 5. Test Duplicate command resolution
    Dim cmdDuplicate As ICommand
    Set cmdDuplicate = AppContainer.ResolveCommand("Duplicate")
    Test_Runner.AssertEqual Not cmdDuplicate Is Nothing, True, "Duplicate command should resolve"
    Test_Runner.AssertEqual TypeName(cmdDuplicate), "FeatCmd_Duplicate", "Duplicate should resolve to FeatCmd_Duplicate"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Test_CommandResolution_NewMenus", Err
    Resume CleanExit
End Sub
