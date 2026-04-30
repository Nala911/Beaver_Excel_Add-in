Attribute VB_Name = "Infra_Bootstrap"
Option Explicit

' @Module: Infra_Bootstrap
' @Category: Infrastructure
' @Description: Centralized startup and shutdown workflow for the add-in host.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: AppContainer, Infra_Config, Infra_Error, Infra_Hotkeys, ExcelContextProvider

Public Sub Startup()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Startup")
    On Error GoTo ErrHandler

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    Infra_Hotkeys.RegisterHotkeys

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Startup", Err
    Resume CleanExit
End Sub

Public Sub Shutdown()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Shutdown")
    On Error GoTo ErrHandler

    Infra_Hotkeys.UnregisterHotkeys

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Shutdown", Err
    Resume CleanExit
End Sub

Public Sub EnsureStarted()
    Dim tracker As Object: Set tracker = Infra_Error.Track("EnsureStarted")
    On Error GoTo ErrHandler

    AppContainer.EnsureInitialized

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "EnsureStarted", Err
    Resume CleanExit
End Sub
