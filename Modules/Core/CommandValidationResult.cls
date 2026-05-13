VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: CommandValidationResult
' @Category: Infrastructure
' @Description: Captures whether a command can execute, and the user-facing feedback to show when it cannot.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Public IsExecutable As Boolean
Public Message As String
Public Title As String
Public MessageStyle As VbMsgBoxStyle
Public ShowMessage As Boolean

Private Sub Class_Initialize()
    IsExecutable = True
    MessageStyle = vbExclamation
End Sub
