VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: ICommand
' @Category: Infrastructure
' @Description: Standard command pattern interface for all features.
 
Public Property Get ExecutionPolicy() As CommandExecutionPolicy
End Property

Public Function Validate(ByVal context As ICommandContext) As CommandValidationResult
End Function

Public Sub Execute(ByVal context As ICommandContext)
End Sub
