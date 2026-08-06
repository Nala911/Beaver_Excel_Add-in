Attribute VB_Name = "Infra_Bootstrap"
Option Explicit

' @Module: Infra_Bootstrap
' @Category: Infrastructure
' @Description: Centralized startup and shutdown workflow for the add-in host.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: AppContainer, Infra_Config, Infra_Error, Infra_Hotkeys, ExcelContextProvider, Lib_UdfRegistry

Public Sub Auto_Open()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Auto_Open")
    On Error GoTo ErrHandler

    Startup

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Auto_Open", Err
    Resume CleanExit
End Sub

Public Sub Startup()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Startup")
    On Error GoTo ErrHandler

    AppContainer.Initialize Infra_Config, Infra_Error, ExcelContextProvider
    
    ' Register hotkeys directly during startup sequence
    Infra_Hotkeys.RegisterHotkeys
    
    RegisterUDFs

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
    ThisWorkbook.Saved = True

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

Private Sub RegisterUDFs()
    Dim tracker As Object: Set tracker = Infra_Error.Track("RegisterUDFs")
    On Error GoTo ErrHandler

    Dim udfs As Collection
    Dim meta As Object

    Set udfs = Lib_UdfRegistry.GetAllUdfs()
    For Each meta In udfs
        ' Register UDF with descriptions in the Function Wizard
        Application.MacroOptions _
            Macro:="'" & ThisWorkbook.Name & "'!" & meta("Name"), _
            Description:=meta("Description"), _
            Category:=meta("Category"), _
            ArgumentDescriptions:=meta("ArgumentDescriptions")
    Next meta

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "RegisterUDFs", Err
    Resume CleanExit
End Sub
