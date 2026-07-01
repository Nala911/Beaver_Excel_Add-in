Attribute VB_Name = "Lib_TestManifest"
Option Explicit

' @Module: Lib_TestManifest
' @Category: Infrastructure
' @Description: Generated test manifest that orchestrates all Test_* procedures.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

Public Sub RunGeneratedTests(Optional ByVal filterPattern As String = "")
    Dim tracker As Object: Set tracker = Infra_Error.Track("RunGeneratedTests")
    On Error GoTo ErrHandler

    If MatchesFilter("Test_CommandInfrastructure.Test_CommandContextIncludesMetadataAndActionContext", filterPattern) Then Test_CommandInfrastructure.Test_CommandContextIncludesMetadataAndActionContext
    If MatchesFilter("Test_CommandInfrastructure.Test_CommandRegistryCreatesKnownCommands", filterPattern) Then Test_CommandInfrastructure.Test_CommandRegistryCreatesKnownCommands
    If MatchesFilter("Test_CommandInfrastructure.Test_CommandRegistryResolvesHotkeyEntries", filterPattern) Then Test_CommandInfrastructure.Test_CommandRegistryResolvesHotkeyEntries
    If MatchesFilter("Test_CommandInfrastructure.Test_CommandRegistryResolvesRibbonEntries", filterPattern) Then Test_CommandInfrastructure.Test_CommandRegistryResolvesRibbonEntries
    If MatchesFilter("Test_CommandInfrastructure.Test_DiagnosticsEscapeJson", filterPattern) Then Test_CommandInfrastructure.Test_DiagnosticsEscapeJson
    If MatchesFilter("Test_CommandInfrastructure.Test_DiagnosticsLogsJSON", filterPattern) Then Test_CommandInfrastructure.Test_DiagnosticsLogsJSON
    If MatchesFilter("Test_Feat_BreakLinks.Test_BreakExternalLinks_Execution", filterPattern) Then Test_Feat_BreakLinks.Test_BreakExternalLinks_Execution
    If MatchesFilter("Test_Feat_BreakLinks.Test_BreakExternalLinks_SpillHandling", filterPattern) Then Test_Feat_BreakLinks.Test_BreakExternalLinks_SpillHandling
    If MatchesFilter("Test_Feat_CleanData.Test_CleanData_CheckboxOptions", filterPattern) Then Test_Feat_CleanData.Test_CleanData_CheckboxOptions
    If MatchesFilter("Test_Feat_CleanData.Test_CleanData_HygieneOptions", filterPattern) Then Test_Feat_CleanData.Test_CleanData_HygieneOptions
    If MatchesFilter("Test_Feat_CleanData.Test_CleanData_LargeSelectionSafety", filterPattern) Then Test_Feat_CleanData.Test_CleanData_LargeSelectionSafety
    If MatchesFilter("Test_Feat_CleanData.Test_CleanData_NewEnhancements", filterPattern) Then Test_Feat_CleanData.Test_CleanData_NewEnhancements
    If MatchesFilter("Test_Feat_CleanData.Test_CleanData_TrimmingAndNumericalFixing", filterPattern) Then Test_Feat_CleanData.Test_CleanData_TrimmingAndNumericalFixing
    If MatchesFilter("Test_Feat_CleanData.Test_CleanData_UserRequestedEnhancements", filterPattern) Then Test_Feat_CleanData.Test_CleanData_UserRequestedEnhancements
    If MatchesFilter("Test_Feat_CleanData.Test_ForceNumber_Execution_And_Undo", filterPattern) Then Test_Feat_CleanData.Test_ForceNumber_Execution_And_Undo
    If MatchesFilter("Test_Feat_CleanData.Test_TryConvertToNumber_Unification", filterPattern) Then Test_Feat_CleanData.Test_TryConvertToNumber_Unification
    If MatchesFilter("Test_Feat_CleanData.Test_UnifiedHelpers_And_CleanDataDisjoint", filterPattern) Then Test_Feat_CleanData.Test_UnifiedHelpers_And_CleanDataDisjoint
    If MatchesFilter("Test_Feat_CleanData.Test_UnmergeFill_Execution_And_Undo", filterPattern) Then Test_Feat_CleanData.Test_UnmergeFill_Execution_And_Undo
    If MatchesFilter("Test_Feat_Editing.Test_Backspace_LargeRange_Undo", filterPattern) Then Test_Feat_Editing.Test_Backspace_LargeRange_Undo
    If MatchesFilter("Test_Feat_Editing.Test_Delete_Execution_And_Undo", filterPattern) Then Test_Feat_Editing.Test_Delete_Execution_And_Undo
    If MatchesFilter("Test_Feat_Editing.Test_FillDown_Features", filterPattern) Then Test_Feat_Editing.Test_FillDown_Features
    If MatchesFilter("Test_Feat_Editing.Test_FilterByCell_Execution", filterPattern) Then Test_Feat_Editing.Test_FilterByCell_Execution
    If MatchesFilter("Test_Feat_Editing.Test_MakePermanent_LegacyArray_And_Undo", filterPattern) Then Test_Feat_Editing.Test_MakePermanent_LegacyArray_And_Undo
    If MatchesFilter("Test_Feat_Editing.Test_MakePermanent_SpillHandling_And_Undo", filterPattern) Then Test_Feat_Editing.Test_MakePermanent_SpillHandling_And_Undo
    If MatchesFilter("Test_Feat_Formatting.Test_ApplyCustomNumberFormat_Execution", filterPattern) Then Test_Feat_Formatting.Test_ApplyCustomNumberFormat_Execution
    If MatchesFilter("Test_Feat_Formatting.Test_FormatRange_ErrorSafety", filterPattern) Then Test_Feat_Formatting.Test_FormatRange_ErrorSafety
    If MatchesFilter("Test_Feat_Formatting.Test_FormatRange_Execution", filterPattern) Then Test_Feat_Formatting.Test_FormatRange_Execution
    If MatchesFilter("Test_Feat_Formatting.Test_FormatRange_WholeSheetSafety", filterPattern) Then Test_Feat_Formatting.Test_FormatRange_WholeSheetSafety
    If MatchesFilter("Test_Feat_Formatting.Test_PasteFormat_Execution", filterPattern) Then Test_Feat_Formatting.Test_PasteFormat_Execution
    If MatchesFilter("Test_Feat_General.Test_GetChunkedRanges_And_SpillExpansion", filterPattern) Then Test_Feat_General.Test_GetChunkedRanges_And_SpillExpansion
    If MatchesFilter("Test_Feat_General.Test_HelloWorld_Execution_And_Undo", filterPattern) Then Test_Feat_General.Test_HelloWorld_Execution_And_Undo
    If MatchesFilter("Test_Feat_General.Test_MultiArea_Undo_Robustness", filterPattern) Then Test_Feat_General.Test_MultiArea_Undo_Robustness
    If MatchesFilter("Test_Feat_General.Test_ProtectedSheet_CanModifyContext", filterPattern) Then Test_Feat_General.Test_ProtectedSheet_CanModifyContext
    If MatchesFilter("Test_Feat_General.Test_ProtectedSheet_CommandValidators", filterPattern) Then Test_Feat_General.Test_ProtectedSheet_CommandValidators
    If MatchesFilter("Test_Feat_General.Test_SingleCell_Bugs_Regression", filterPattern) Then Test_Feat_General.Test_SingleCell_Bugs_Regression
    If MatchesFilter("Test_Feat_General.Test_ValueConversion_ResolveSpillExpandedRange", filterPattern) Then Test_Feat_General.Test_ValueConversion_ResolveSpillExpandedRange
    If MatchesFilter("Test_Feat_General.Test_XFilter_Features", filterPattern) Then Test_Feat_General.Test_XFilter_Features
    If MatchesFilter("Test_Feat_General.Test_XUnpivot_Features", filterPattern) Then Test_Feat_General.Test_XUnpivot_Features
    If MatchesFilter("Test_Feat_HighlightData.Test_HighlightData_ClearHighlights", filterPattern) Then Test_Feat_HighlightData.Test_HighlightData_ClearHighlights
    If MatchesFilter("Test_Feat_HighlightData.Test_HighlightData_ConditionalFormatting", filterPattern) Then Test_Feat_HighlightData.Test_HighlightData_ConditionalFormatting
    If MatchesFilter("Test_Feat_HighlightData.Test_HighlightData_DataValidations", filterPattern) Then Test_Feat_HighlightData.Test_HighlightData_DataValidations
    If MatchesFilter("Test_Feat_HighlightData.Test_HighlightData_Errors", filterPattern) Then Test_Feat_HighlightData.Test_HighlightData_Errors
    If MatchesFilter("Test_Feat_HighlightData.Test_HighlightData_FormulaLimitSafety", filterPattern) Then Test_Feat_HighlightData.Test_HighlightData_FormulaLimitSafety
    If MatchesFilter("Test_Feat_HighlightData.Test_HighlightData_HardcodedValues", filterPattern) Then Test_Feat_HighlightData.Test_HighlightData_HardcodedValues
    If MatchesFilter("Test_Feat_HighlightData.Test_HighlightData_InconsistentFormulasAndDuplicates", filterPattern) Then Test_Feat_HighlightData.Test_HighlightData_InconsistentFormulasAndDuplicates
    If MatchesFilter("Test_Feat_ModifyData.Test_ModifyData_BulkFormattingOverride", filterPattern) Then Test_Feat_ModifyData.Test_ModifyData_BulkFormattingOverride
    If MatchesFilter("Test_Feat_ModifyData.Test_ModifyData_Casing", filterPattern) Then Test_Feat_ModifyData.Test_ModifyData_Casing
    If MatchesFilter("Test_Feat_ModifyData.Test_ModifyData_DateStandardization", filterPattern) Then Test_Feat_ModifyData.Test_ModifyData_DateStandardization
    If MatchesFilter("Test_Feat_ModifyData.Test_ModifyData_MixedFormats", filterPattern) Then Test_Feat_ModifyData.Test_ModifyData_MixedFormats
    If MatchesFilter("Test_Feat_ModifyData.Test_ModifyData_MixedFormats_DateStandardization", filterPattern) Then Test_Feat_ModifyData.Test_ModifyData_MixedFormats_DateStandardization
    If MatchesFilter("Test_Feat_ModifyData.Test_ModifyData_Undo", filterPattern) Then Test_Feat_ModifyData.Test_ModifyData_Undo
    If MatchesFilter("Test_Feat_Workbook.Test_CleanWorkbookNames_BrokenAndExternal", filterPattern) Then Test_Feat_Workbook.Test_CleanWorkbookNames_BrokenAndExternal
    If MatchesFilter("Test_Feat_Workbook.Test_CreateSheet_PlacementAndNaming", filterPattern) Then Test_Feat_Workbook.Test_CreateSheet_PlacementAndNaming
    If MatchesFilter("Test_Feat_Workbook.Test_Export_Pdf_Backup_And_MultiRange", filterPattern) Then Test_Feat_Workbook.Test_Export_Pdf_Backup_And_MultiRange
    If MatchesFilter("Test_Feat_Workbook.Test_StaticSheetWorkbook_Execution", filterPattern) Then Test_Feat_Workbook.Test_StaticSheetWorkbook_Execution
    If MatchesFilter("Test_Feat_Workbook.Test_TableOfContents_Generation", filterPattern) Then Test_Feat_Workbook.Test_TableOfContents_Generation
    If MatchesFilter("Test_Feat_Wrap.Test_Wrap_CellAndPatternModes", filterPattern) Then Test_Feat_Wrap.Test_Wrap_CellAndPatternModes
    If MatchesFilter("Test_Runner.Test_ConfigProvidesTypedHotkeys", filterPattern) Then Test_Runner.Test_ConfigProvidesTypedHotkeys
    If MatchesFilter("Test_Runner.Test_Infrastructure_Basics", filterPattern) Then Test_Runner.Test_Infrastructure_Basics
    If MatchesFilter("Test_Runner.Test_TranslateHotkeyHandlesModifiers", filterPattern) Then Test_Runner.Test_TranslateHotkeyHandlesModifiers
    If MatchesFilter("Test_UI.Test_CommandResolution_NewMenus", filterPattern) Then Test_UI.Test_CommandResolution_NewMenus
    If MatchesFilter("Test_UI.Test_UdfRegistry_And_HelpCenter", filterPattern) Then Test_UI.Test_UdfRegistry_And_HelpCenter
    If MatchesFilter("Test_UI.Test_UI_OptionPicker_DynamicLayout", filterPattern) Then Test_UI.Test_UI_OptionPicker_DynamicLayout
    If MatchesFilter("Test_UI.Test_UI_OptionPicker_KeyboardNavigation", filterPattern) Then Test_UI.Test_UI_OptionPicker_KeyboardNavigation

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "RunGeneratedTests", Err
    Resume CleanExit
End Sub

Private Function MatchesFilter(ByVal testName As String, ByVal filterPattern As String) As Boolean
    If filterPattern = "" Then
        MatchesFilter = True
        Exit Function
    End If
    Dim patterns() As String
    patterns = Split(filterPattern, ",")
    Dim i As Long
    For i = LBound(patterns) To UBound(patterns)
        If UCase$(testName) Like UCase$(Trim$(patterns(i))) Then
            MatchesFilter = True
            Exit Function
        End If
    Next i
    MatchesFilter = False
End Function