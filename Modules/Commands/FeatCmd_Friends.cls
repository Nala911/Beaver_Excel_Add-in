VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_Friends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: FeatCmd_Friends
' @Category: Feature
' @Description: Puts 'Hello friends , this is just testing' in the active cell.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ICommandContext, Infra_Error, Infra_CommandSupport, Infra_Undo

Implements ICommand

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Friends.Execute")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit

    Dim targetCell As Range
    Set targetCell = ctx.ActiveCellRef
    If targetCell Is Nothing Then GoTo CleanExit

    If Not Infra_Undo.SaveStateOrConfirm(targetCell, "Say Hello Friends") Then GoTo CleanExit

    targetCell.Value2 = "Hello friends , this is just testing"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Friends.Execute", Err
    Resume CleanExit
End Sub
