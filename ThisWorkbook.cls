VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' @Module: ThisWorkbook
' @Category: Infrastructure
' @Description: Workbook event bridge for Excel add-in startup and shutdown hooks.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Hotkeys, Infra_Error

' ==============================================================================
' ThisWorkbook - Event module for the Beaver Add-in workbook.
' Kept intentionally thin: delegates all logic to standard modules.
'   Workbook_Open        -> Hotkeys.RegisterHotkeys
'   Workbook_BeforeClose -> Hotkeys.UnregisterHotkeys
' ==============================================================================

Private Sub Workbook_Open()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Workbook_Open")
    On Error GoTo ErrHandler

    Infra_Bootstrap.Startup

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Workbook_Open", Err
    Resume CleanExit
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Workbook_BeforeClose")
    On Error GoTo ErrHandler

    Infra_Bootstrap.Shutdown

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Workbook_BeforeClose", Err
    Resume CleanExit
    ' Trigger incremental build compile test
End Sub
