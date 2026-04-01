VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicCommandContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommandContext

' @Module: BasicCommandContext
' @Category: Infrastructure
' @Description: Simple implementation of ICommandContext for standard commands.

Private mSourceName As String
Private mTriggerKind As String

Public Property Get SourceName() As String
    SourceName = mSourceName
End Property

Public Property Let SourceName(ByVal value As String)
    mSourceName = value
End Property

Public Property Get TriggerKind() As String
    TriggerKind = mTriggerKind
End Property

Public Property Let TriggerKind(ByVal value As String)
    mTriggerKind = value
End Property

Private Property Get ICommandContext_SourceName() As String
    ICommandContext_SourceName = Me.SourceName
End Property

Private Property Get ICommandContext_TriggerKind() As String
    ICommandContext_TriggerKind = Me.TriggerKind
End Property
