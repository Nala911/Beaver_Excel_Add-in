VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandExecutionPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: CommandExecutionPolicy
' @Category: Infrastructure
' @Description: Describes how the command invoker should configure Excel application state for a command.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Public UseAppStateGuard As Boolean
Public ScreenUpdating As Boolean
Public EnableEvents As Boolean
Public DisplayAlerts As Boolean
Public Calculation As XlCalculation

Private Sub Class_Initialize()
    UseAppStateGuard = True
    ScreenUpdating = False
    EnableEvents = False
    DisplayAlerts = True
    Calculation = xlCalculationManual
End Sub
