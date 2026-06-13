Attribute VB_Name = "Lib_TestManifest"
Option Explicit

' @Module: Lib_TestManifest
' @Category: Infrastructure
' @Description: Generated test manifest that orchestrates all Test_* procedures.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

Public Sub RunGeneratedTests()
    Dim tracker As Object: Set tracker = Infra_Error.Track("RunGeneratedTests")
    On Error GoTo ErrHandler

    Lib_Tests.Test_ConfigProvidesTypedHotkeys
    Lib_Tests.Test_Infrastructure_Basics
    Lib_Tests.Test_TranslateHotkeyHandlesModifiers
    Lib_Tests_CommandInfrastructure.Test_CommandContextIncludesMetadataAndActionContext
    Lib_Tests_CommandInfrastructure.Test_CommandRegistryCreatesKnownCommands
    Lib_Tests_CommandInfrastructure.Test_CommandRegistryResolvesHotkeyEntries
    Lib_Tests_CommandInfrastructure.Test_CommandRegistryResolvesRibbonEntries
    Lib_Tests_Features.Test_ApplyCustomNumberFormat_Execution
    Lib_Tests_Features.Test_Backspace_LargeRange_Undo
    Lib_Tests_Features.Test_BreakExternalLinks_Execution
    Lib_Tests_Features.Test_CleanData_CheckboxOptions
    Lib_Tests_Features.Test_CleanData_LargeSelectionSafety
    Lib_Tests_Features.Test_CleanData_NewEnhancements
    Lib_Tests_Features.Test_CleanData_TrimmingAndNumericalFixing
    Lib_Tests_Features.Test_CleanData_UserRequestedEnhancements
    Lib_Tests_Features.Test_CreateSheet_PlacementAndNaming
    Lib_Tests_Features.Test_Delete_Execution_And_Undo
    Lib_Tests_Features.Test_FillDown_Features
    Lib_Tests_Features.Test_FilterByCell_Execution
    Lib_Tests_Features.Test_FormatRange_Execution
    Lib_Tests_Features.Test_FormatRange_WholeSheetSafety
    Lib_Tests_Features.Test_GetChunkedRanges_And_SpillExpansion
    Lib_Tests_Features.Test_HelloWorld_Execution_And_Undo
    Lib_Tests_Features.Test_HighlightData_Errors
    Lib_Tests_Features.Test_HighlightData_FormulaLimitSafety
    Lib_Tests_Features.Test_HighlightData_InconsistentFormulasAndDuplicates
    Lib_Tests_Features.Test_MakePermanent_LegacyArray_And_Undo
    Lib_Tests_Features.Test_MakePermanent_SpillHandling_And_Undo
    Lib_Tests_Features.Test_ModifyData_Casing
    Lib_Tests_Features.Test_ModifyData_DateStandardization
    Lib_Tests_Features.Test_ModifyData_Undo
    Lib_Tests_Features.Test_PasteFormat_Execution
    Lib_Tests_Features.Test_StaticSheetWorkbook_Execution
    Lib_Tests_Features.Test_UdfRegistry_And_HelpCenter
    Lib_Tests_Features.Test_ValueConversion_ResolveSpillExpandedRange
    Lib_Tests_Features.Test_Wrap_CellAndPatternModes
    Lib_Tests_Features.Test_XFilter_Features
    Lib_Tests_Features.Test_XUnpivot_Features

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "RunGeneratedTests", Err
    Resume CleanExit
End Sub