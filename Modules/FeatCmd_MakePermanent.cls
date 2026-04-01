VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_MakePermanent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_MakePermanent
' @Category: Feature
' @Description: Converts selection to values.

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("MakePermanent.CanExecute")
    On Error GoTo ErrHandler
    ICommand_CanExecute = Infra_AppState.IsRangeSelected()
    Exit Function
ErrHandler:
    Infra_Error.HandleError "MakePermanent.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("MakePermanent.Execute")
    On Error GoTo ErrHandler

    Dim rng As Range

    If Not Infra_AppState.IsRangeSelected() Then GoTo CleanExit
    Set rng = Selection
    If rng Is Nothing Then GoTo CleanExit

    Infra_Undo.SaveState rng, "Make Permanent"
    Infra_Progress.StartProgress "Converting Selection to Values", 1

    ' Direct value assignment preserves existing formatting and avoids clipboard overhead.
    rng.Value = rng.Value
    Application.CutCopyMode = False
    Infra_Progress.UpdateProgress 1, True

CleanExit:
    Infra_Progress.EndProgress
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    Infra_Error.HandleError "MakePermanent.Execute", Err
    Resume CleanExit
End Sub
