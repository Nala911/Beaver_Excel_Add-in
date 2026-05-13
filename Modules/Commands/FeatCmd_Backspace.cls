VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_Backspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_Backspace
' @Category: Feature
' @Description: Custom handlers for Backspace with protection awareness.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Backspace.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As Infra_ActionContext

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_AppState.CanModifyContext(ctx) Then GoTo CleanExit
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit

    Infra_Undo.SaveState ctx.SelectionRange, "Clear"
    ctx.SelectionRange.ClearContents
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Backspace.Execute", Err
    Resume CleanExit
End Sub
