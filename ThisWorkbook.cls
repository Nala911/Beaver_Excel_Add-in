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
' @Description: Workbook event bridge for add-in startup and shutdown hooks.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Hotkeys, Infra_Error

' ==============================================================================
' ThisWorkbook â€” Event module for the Beaver Add-in workbook.
' Kept intentionally thin: delegates all logic to standard modules.
'   Workbook_Open        -> Hotkeys.RegisterHotkeys
'   Workbook_BeforeClose -> Hotkeys.UnregisterHotkeys
' ==============================================================================

Private Sub Workbook_Open()
    Infra_Error.PushContext "Workbook_Open"
    On Error GoTo ErrHandler
    Infra_Hotkeys.RegisterHotkeys
    Infra_Error.PopContext
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Workbook_Open", Err
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Infra_Error.PushContext "Workbook_BeforeClose"
    On Error GoTo ErrHandler
    Infra_Hotkeys.UnregisterHotkeys
    Infra_Error.PopContext
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Workbook_BeforeClose", Err
End Sub
