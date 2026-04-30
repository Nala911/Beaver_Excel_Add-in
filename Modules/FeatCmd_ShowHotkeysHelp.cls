VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_ShowHotkeysHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: FeatCmd_ShowHotkeysHelp
' @Category: Feature
' @Description: Displays the Beaver hotkey reference dialog.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ICommandContext, Infra_Hotkeys, Infra_Error

Implements ICommand

Private Function ICommand_CanExecute(ByVal context As ICommandContext) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowHotkeysHelp.CanExecute")
    On Error GoTo ErrHandler

    ICommand_CanExecute = True

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ShowHotkeysHelp.CanExecute", Err
    ICommand_CanExecute = False
    Resume CleanExit
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowHotkeysHelp.Execute")
    On Error GoTo ErrHandler

    Infra_Hotkeys.ShowHotkeysHelp

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ShowHotkeysHelp.Execute", Err
    Resume CleanExit
End Sub
