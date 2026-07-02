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
' @Category: Core
' @Description: Simple implementation of ICommandContext for standard commands.

Private mSourceName As String
Private mTriggerKind As String
Private mCommandName As String
Private mEntryMacro As String
Private mActionContext As ActionContext

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

Public Property Get CommandName() As String
    CommandName = mCommandName
End Property

Public Property Let CommandName(ByVal value As String)
    mCommandName = value
End Property

Public Property Get EntryMacro() As String
    EntryMacro = mEntryMacro
End Property

Public Property Let EntryMacro(ByVal value As String)
    mEntryMacro = value
End Property

Public Property Get ActionContext() As ActionContext
    Set ActionContext = mActionContext
End Property

Public Property Set ActionContext(ByVal value As ActionContext)
    Set mActionContext = value
End Property

Private Property Get ICommandContext_SourceName() As String
    ICommandContext_SourceName = Me.SourceName
End Property

Private Property Get ICommandContext_TriggerKind() As String
    ICommandContext_TriggerKind = Me.TriggerKind
End Property

Private Property Get ICommandContext_CommandName() As String
    ICommandContext_CommandName = Me.CommandName
End Property

Private Property Get ICommandContext_EntryMacro() As String
    ICommandContext_EntryMacro = Me.EntryMacro
End Property

Private Property Get ICommandContext_ActionContext() As ActionContext
    Set ICommandContext_ActionContext = Me.ActionContext
End Property
