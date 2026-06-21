VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_HelloWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: FeatCmd_HelloWorld
' @Category: Feature
' @Description: Puts Hello World in the active cell of the active worksheet.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ICommandContext, Infra_Error, Infra_CommandSupport, Infra_Undo

' Test optimization comment
Implements ICommand

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorksheet(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("HelloWorld.Execute")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit

    If Not Infra_Undo.SaveStateOrConfirm(ctx.ActiveCellRef, "Hello World") Then GoTo CleanExit

    ctx.ActiveCellRef.Value2 = "Hello world!"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "HelloWorld.Execute", Err
    Resume CleanExit
End Sub
