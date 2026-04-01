Attribute VB_Name = "UI_Hotkeys"
Option Explicit

' @Module: UI_Hotkeys
' @Category: UI
' @Description: Hotkey wrappers that route through the CommandInvoker.

Public Sub Hotkey_MakePermanent()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Hotkey_MakePermanent")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "MakePermanent", "Hotkey_MakePermanent", "Hotkey"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Hotkey_MakePermanent", Err
    Resume CleanExit
End Sub

Public Sub Hotkey_ApplyCustomNumberFormat()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Hotkey_ApplyCustomNumberFormat")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "ApplyCustomNumberFormat", "Hotkey_ApplyCustomNumberFormat", "Hotkey"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Hotkey_ApplyCustomNumberFormat", Err
    Resume CleanExit
End Sub

Public Sub Hotkey_CreateNamedSheet()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Hotkey_CreateNamedSheet")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "CreateSheet", "Hotkey_CreateNamedSheet", "Hotkey"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Hotkey_CreateNamedSheet", Err
    Resume CleanExit
End Sub

Public Sub Hotkey_FillDown()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Hotkey_FillDown")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "FillDown", "Hotkey_FillDown", "Hotkey"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Hotkey_FillDown", Err
    Resume CleanExit
End Sub

Public Sub Hotkey_FilterBySelectedCell()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Hotkey_FilterBySelectedCell")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "FilterByCell", "Hotkey_FilterBySelectedCell", "Hotkey"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Hotkey_FilterBySelectedCell", Err
    Resume CleanExit
End Sub

Public Sub Hotkey_PasteFormat()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Hotkey_PasteFormat")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "PasteFormat", "Hotkey_PasteFormat", "Hotkey"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Hotkey_PasteFormat", Err
    Resume CleanExit
End Sub

Public Sub Hotkey_FormatSelectedRange()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Hotkey_FormatSelectedRange")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "FormatRange", "Hotkey_FormatSelectedRange", "Hotkey"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Hotkey_FormatSelectedRange", Err
    Resume CleanExit
End Sub

Public Sub Hotkey_Backspace()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Hotkey_Backspace")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "Backspace", "Hotkey_Backspace", "Hotkey"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Hotkey_Backspace", Err
    Resume CleanExit
End Sub

Public Sub Hotkey_Delete()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Hotkey_Delete")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "Delete", "Hotkey_Delete", "Hotkey"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Hotkey_Delete", Err
    Resume CleanExit
End Sub
