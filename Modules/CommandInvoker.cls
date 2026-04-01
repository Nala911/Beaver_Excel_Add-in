VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandInvoker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: CommandInvoker
' @Category: Infrastructure
' @Description: Centralized pipeline for executing commands. Handles guards and errors.

Public Sub ExecuteCommand(ByVal cmd As ICommand, ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteCommand")
    On Error GoTo ErrHandler
    
    If cmd Is Nothing Then GoTo CleanExit
    
    If Not cmd.CanExecute() Then
        MsgBox "This command cannot be executed in the current context.", vbExclamation, "Command Invalid"
        GoTo CleanExit
    End If
    
    ' Unified Guard
    Dim guard As New Infra_AppStateGuard
    guard.ScreenUpdating = False
    guard.EnableEvents = False
    guard.DisplayAlerts = True
    guard.Calculation = xlCalculationManual
    guard.UpdateState
    
    ' Execute
    cmd.Execute context

CleanExit:
    Exit Sub
    
ErrHandler:
    Infra_Error.HandleError "ExecuteCommand", Err
    Resume CleanExit
End Sub
