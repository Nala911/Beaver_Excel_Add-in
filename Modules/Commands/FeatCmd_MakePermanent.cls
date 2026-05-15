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

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasRangeSelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("MakePermanent.Execute")
    On Error GoTo ErrHandler

    Dim rng As Range
    Dim ctx As Infra_ActionContext

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit
    Set rng = ctx.SelectionRange
    If rng Is Nothing Then GoTo CleanExit

    Infra_Undo.SaveState rng, "Make Permanent"
    Infra_Progress.StartProgress "Converting Selection to Values", 1

    ' Reuse the shared value-conversion helper so hotkey and ribbon flows stay aligned.
    Infra_ValueConversion.ConvertRangeToValues rng
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
