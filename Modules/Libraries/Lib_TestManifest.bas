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
    Lib_Tests_Features.Test_HelloWorld_Execution_And_Undo

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "RunGeneratedTests", Err
    Resume CleanExit
End Sub