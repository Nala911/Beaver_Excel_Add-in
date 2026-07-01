VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScopedRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRequest

' @Module: ScopedRequest
' @Category: Core
' @Description: Typed request describing the target scope and action context for a scoped command execution.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ActionContext, TargetScope

Public Context As ActionContext
Public Scope As TargetScope

Private Property Get IRequest_Context() As ActionContext
    Set IRequest_Context = Me.Context
End Property

Private Property Get IRequest_Scope() As TargetScope
    IRequest_Scope = Me.Scope
End Property
