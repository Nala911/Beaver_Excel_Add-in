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

Private Function ICommand_CanExecute(ByVal context As ICommandContext) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("Backspace.CanExecute")
    On Error GoTo ErrHandler
    Dim ctx As Infra_ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    ICommand_CanExecute = Infra_AppState.CanModifyContext(ctx)
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Backspace.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Backspace.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As Infra_ActionContext

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_AppState.CanModifyContext(ctx) Then GoTo CleanExit
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit

    ctx.SelectionRange.ClearContents
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Backspace.Execute", Err
    Resume CleanExit
End Sub
