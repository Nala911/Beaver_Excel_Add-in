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

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "RunGeneratedTests", Err
    Resume CleanExit
End Sub