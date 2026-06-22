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

    If MatchesFilter("Lib_Tests.Test_ConfigProvidesTypedHotkeys", filterPattern) Then Lib_Tests.Test_ConfigProvidesTypedHotkeys
    If MatchesFilter("Lib_Tests.Test_Infrastructure_Basics", filterPattern) Then Lib_Tests.Test_Infrastructure_Basics
    If MatchesFilter("Lib_Tests.Test_TranslateHotkeyHandlesModifiers", filterPattern) Then Lib_Tests.Test_TranslateHotkeyHandlesModifiers
    If MatchesFilter("Lib_Tests_CommandInfrastructure.Test_CommandContextIncludesMetadataAndActionContext", filterPattern) Then Lib_Tests_CommandInfrastructure.Test_CommandContextIncludesMetadataAndActionContext
    If MatchesFilter("Lib_Tests_CommandInfrastructure.Test_CommandRegistryCreatesKnownCommands", filterPattern) Then Lib_Tests_CommandInfrastructure.Test_CommandRegistryCreatesKnownCommands
    If MatchesFilter("Lib_Tests_CommandInfrastructure.Test_CommandRegistryResolvesHotkeyEntries", filterPattern) Then Lib_Tests_CommandInfrastructure.Test_CommandRegistryResolvesHotkeyEntries
    If MatchesFilter("Lib_Tests_CommandInfrastructure.Test_CommandRegistryResolvesRibbonEntries", filterPattern) Then Lib_Tests_CommandInfrastructure.Test_CommandRegistryResolvesRibbonEntries
    If MatchesFilter("Lib_Tests_Features.Test_ApplyCustomNumberFormat_Execution", filterPattern) Then Lib_Tests_Features.Test_ApplyCustomNumberFormat_Execution
    If MatchesFilter("Lib_Tests_Features.Test_Backspace_LargeRange_Undo", filterPattern) Then Lib_Tests_Features.Test_Backspace_LargeRange_Undo
    If MatchesFilter("Lib_Tests_Features.Test_BreakExternalLinks_Execution", filterPattern) Then Lib_Tests_Features.Test_BreakExternalLinks_Execution
    If MatchesFilter("Lib_Tests_Features.Test_BreakExternalLinks_SpillHandling", filterPattern) Then Lib_Tests_Features.Test_BreakExternalLinks_SpillHandling
    If MatchesFilter("Lib_Tests_Features.Test_CleanData_CheckboxOptions", filterPattern) Then Lib_Tests_Features.Test_CleanData_CheckboxOptions
    If MatchesFilter("Lib_Tests_Features.Test_CleanData_HygieneOptions", filterPattern) Then Lib_Tests_Features.Test_CleanData_HygieneOptions
    If MatchesFilter("Lib_Tests_Features.Test_CleanData_LargeSelectionSafety", filterPattern) Then Lib_Tests_Features.Test_CleanData_LargeSelectionSafety
    If MatchesFilter("Lib_Tests_Features.Test_CleanData_NewEnhancements", filterPattern) Then Lib_Tests_Features.Test_CleanData_NewEnhancements
    If MatchesFilter("Lib_Tests_Features.Test_CleanData_TrimmingAndNumericalFixing", filterPattern) Then Lib_Tests_Features.Test_CleanData_TrimmingAndNumericalFixing
    If MatchesFilter("Lib_Tests_Features.Test_CleanData_UserRequestedEnhancements", filterPattern) Then Lib_Tests_Features.Test_CleanData_UserRequestedEnhancements
    If MatchesFilter("Lib_Tests_Features.Test_CleanWorkbookNames_BrokenAndExternal", filterPattern) Then Lib_Tests_Features.Test_CleanWorkbookNames_BrokenAndExternal
    If MatchesFilter("Lib_Tests_Features.Test_CommandResolution_NewMenus", filterPattern) Then Lib_Tests_Features.Test_CommandResolution_NewMenus
    If MatchesFilter("Lib_Tests_Features.Test_CreateSheet_PlacementAndNaming", filterPattern) Then Lib_Tests_Features.Test_CreateSheet_PlacementAndNaming
    If MatchesFilter("Lib_Tests_Features.Test_Delete_Execution_And_Undo", filterPattern) Then Lib_Tests_Features.Test_Delete_Execution_And_Undo
    If MatchesFilter("Lib_Tests_Features.Test_Export_Pdf_Backup_And_MultiRange", filterPattern) Then Lib_Tests_Features.Test_Export_Pdf_Backup_And_MultiRange
    If MatchesFilter("Lib_Tests_Features.Test_FillDown_Features", filterPattern) Then Lib_Tests_Features.Test_FillDown_Features
    If MatchesFilter("Lib_Tests_Features.Test_FilterByCell_Execution", filterPattern) Then Lib_Tests_Features.Test_FilterByCell_Execution
    If MatchesFilter("Lib_Tests_Features.Test_ForceNumber_Execution_And_Undo", filterPattern) Then Lib_Tests_Features.Test_ForceNumber_Execution_And_Undo
    If MatchesFilter("Lib_Tests_Features.Test_FormatRange_ErrorSafety", filterPattern) Then Lib_Tests_Features.Test_FormatRange_ErrorSafety
    If MatchesFilter("Lib_Tests_Features.Test_FormatRange_Execution", filterPattern) Then Lib_Tests_Features.Test_FormatRange_Execution
    If MatchesFilter("Lib_Tests_Features.Test_FormatRange_WholeSheetSafety", filterPattern) Then Lib_Tests_Features.Test_FormatRange_WholeSheetSafety
    If MatchesFilter("Lib_Tests_Features.Test_GetChunkedRanges_And_SpillExpansion", filterPattern) Then Lib_Tests_Features.Test_GetChunkedRanges_And_SpillExpansion
    If MatchesFilter("Lib_Tests_Features.Test_HelloWorld_Execution_And_Undo", filterPattern) Then Lib_Tests_Features.Test_HelloWorld_Execution_And_Undo
    If MatchesFilter("Lib_Tests_Features.Test_HighlightData_ConditionalFormatting", filterPattern) Then Lib_Tests_Features.Test_HighlightData_ConditionalFormatting
    If MatchesFilter("Lib_Tests_Features.Test_HighlightData_DataValidations", filterPattern) Then Lib_Tests_Features.Test_HighlightData_DataValidations
    If MatchesFilter("Lib_Tests_Features.Test_HighlightData_Errors", filterPattern) Then Lib_Tests_Features.Test_HighlightData_Errors
    If MatchesFilter("Lib_Tests_Features.Test_HighlightData_FormulaLimitSafety", filterPattern) Then Lib_Tests_Features.Test_HighlightData_FormulaLimitSafety
    If MatchesFilter("Lib_Tests_Features.Test_HighlightData_HardcodedValues", filterPattern) Then Lib_Tests_Features.Test_HighlightData_HardcodedValues
    If MatchesFilter("Lib_Tests_Features.Test_HighlightData_InconsistentFormulasAndDuplicates", filterPattern) Then Lib_Tests_Features.Test_HighlightData_InconsistentFormulasAndDuplicates
    If MatchesFilter("Lib_Tests_Features.Test_MakePermanent_LegacyArray_And_Undo", filterPattern) Then Lib_Tests_Features.Test_MakePermanent_LegacyArray_And_Undo
    If MatchesFilter("Lib_Tests_Features.Test_MakePermanent_SpillHandling_And_Undo", filterPattern) Then Lib_Tests_Features.Test_MakePermanent_SpillHandling_And_Undo
    If MatchesFilter("Lib_Tests_Features.Test_ModifyData_Casing", filterPattern) Then Lib_Tests_Features.Test_ModifyData_Casing
    If MatchesFilter("Lib_Tests_Features.Test_ModifyData_DateStandardization", filterPattern) Then Lib_Tests_Features.Test_ModifyData_DateStandardization
    If MatchesFilter("Lib_Tests_Features.Test_ModifyData_MixedFormats", filterPattern) Then Lib_Tests_Features.Test_ModifyData_MixedFormats
    If MatchesFilter("Lib_Tests_Features.Test_ModifyData_Undo", filterPattern) Then Lib_Tests_Features.Test_ModifyData_Undo
    If MatchesFilter("Lib_Tests_Features.Test_PasteFormat_Execution", filterPattern) Then Lib_Tests_Features.Test_PasteFormat_Execution
    If MatchesFilter("Lib_Tests_Features.Test_SingleCell_Bugs_Regression", filterPattern) Then Lib_Tests_Features.Test_SingleCell_Bugs_Regression
    If MatchesFilter("Lib_Tests_Features.Test_StaticSheetWorkbook_Execution", filterPattern) Then Lib_Tests_Features.Test_StaticSheetWorkbook_Execution
    If MatchesFilter("Lib_Tests_Features.Test_TableOfContents_Generation", filterPattern) Then Lib_Tests_Features.Test_TableOfContents_Generation
    If MatchesFilter("Lib_Tests_Features.Test_TryConvertToNumber_Unification", filterPattern) Then Lib_Tests_Features.Test_TryConvertToNumber_Unification
    If MatchesFilter("Lib_Tests_Features.Test_UdfRegistry_And_HelpCenter", filterPattern) Then Lib_Tests_Features.Test_UdfRegistry_And_HelpCenter
    If MatchesFilter("Lib_Tests_Features.Test_UI_OptionPicker_DynamicLayout", filterPattern) Then Lib_Tests_Features.Test_UI_OptionPicker_DynamicLayout
    If MatchesFilter("Lib_Tests_Features.Test_UI_OptionPicker_KeyboardNavigation", filterPattern) Then Lib_Tests_Features.Test_UI_OptionPicker_KeyboardNavigation
    If MatchesFilter("Lib_Tests_Features.Test_UnifiedHelpers_And_CleanDataDisjoint", filterPattern) Then Lib_Tests_Features.Test_UnifiedHelpers_And_CleanDataDisjoint
    If MatchesFilter("Lib_Tests_Features.Test_UnmergeFill_Execution_And_Undo", filterPattern) Then Lib_Tests_Features.Test_UnmergeFill_Execution_And_Undo
    If MatchesFilter("Lib_Tests_Features.Test_ValueConversion_ResolveSpillExpandedRange", filterPattern) Then Lib_Tests_Features.Test_ValueConversion_ResolveSpillExpandedRange
    If MatchesFilter("Lib_Tests_Features.Test_Wrap_CellAndPatternModes", filterPattern) Then Lib_Tests_Features.Test_Wrap_CellAndPatternModes
    If MatchesFilter("Lib_Tests_Features.Test_XFilter_Features", filterPattern) Then Lib_Tests_Features.Test_XFilter_Features
    If MatchesFilter("Lib_Tests_Features.Test_XUnpivot_Features", filterPattern) Then Lib_Tests_Features.Test_XUnpivot_Features

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