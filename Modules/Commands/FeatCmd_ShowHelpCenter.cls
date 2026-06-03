VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_ShowHelpCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: FeatCmd_ShowHelpCenter
' @Category: Feature
' @Description: Displays the Beaver Help Center dialog containing hotkey shortcuts and UDF details (e.g. XFilter).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ICommandContext, Infra_Hotkeys, Infra_Error

Implements ICommand

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidationSuccess()
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowHelpCenter.Execute")
    On Error GoTo ErrHandler

    Infra_Hotkeys.ShowHelpCenter

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ShowHelpCenter.Execute", Err
    Resume CleanExit
End Sub
