VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: ICommandContext
' @Category: Infrastructure
' @Description: Provides runtime context to commands (e.g., Ribbon state, arguments).

Public Property Get SourceName() As String
End Property

Public Property Get TriggerKind() As String
End Property
