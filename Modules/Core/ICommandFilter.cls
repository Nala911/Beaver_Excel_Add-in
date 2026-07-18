VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: ICommandFilter
' @Category: Core
' @Description: Interface for intercepting command execution (middleware pattern).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ICommandContext

Public Sub BeforeExecute(ByVal cmd As ICommand, ByVal context As ICommandContext, ByRef cancel As Boolean)
    ' Pre-execution hook
End Sub

Public Sub AfterExecute(ByVal cmd As ICommand, ByVal context As ICommandContext, ByVal success As Boolean)
    ' Post-execution hook
End Sub
