VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_FilterAppState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_FilterAppState
' @Category: Infrastructure
' @Description: Command filter that manages Excel application state toggles based on command execution policy.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommandFilter, ICommand, ICommandContext, CommandExecutionPolicy, Infra_AppStateGuard, Infra_CommandSupport, Infra_Error

Implements ICommandFilter

Private m_Guards As Object

Private Sub Class_Initialize()
    Set m_Guards = CreateObject("Scripting.Dictionary")
End Sub

Private Sub ICommandFilter_BeforeExecute(ByVal cmd As ICommand, ByVal context As ICommandContext, ByRef cancel As Boolean)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_FilterAppState.BeforeExecute")
    On Error GoTo ErrHandler

    If cmd Is Nothing Or context Is Nothing Then Exit Sub

    Dim policy As CommandExecutionPolicy
    Set policy = cmd.ExecutionPolicy
    If policy Is Nothing Then Set policy = Infra_CommandSupport.CreateExecutionPolicy()

    If policy.UseAppStateGuard Then
        Dim guard As New Infra_AppStateGuard
        guard.ScreenUpdating = policy.ScreenUpdating
        guard.EnableEvents = policy.EnableEvents
        guard.DisplayAlerts = policy.DisplayAlerts
        guard.Calculation = policy.Calculation
        guard.UpdateState
        
        Set m_Guards(context) = guard
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Infra_FilterAppState.BeforeExecute", Err
    Resume CleanExit
End Sub

Private Sub ICommandFilter_AfterExecute(ByVal cmd As ICommand, ByVal context As ICommandContext, ByVal success As Boolean)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_FilterAppState.AfterExecute")
    On Error GoTo ErrHandler

    If context Is Nothing Then Exit Sub

    If m_Guards.Exists(context) Then
        Dim guard As Infra_AppStateGuard
        Set guard = m_Guards(context)
        m_Guards.Remove context
        Set guard = Nothing
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Infra_FilterAppState.AfterExecute", Err
    Resume CleanExit
End Sub
