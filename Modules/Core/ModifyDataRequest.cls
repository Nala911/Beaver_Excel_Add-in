VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModifyDataRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRequest

' @Module: ModifyDataRequest
' @Category: Core
' @Description: Typed request describing the target selection, action context, and data modification options.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ActionContext

Public Context As ActionContext
Public Operation As String
Public DatePattern As String

Private Property Get IRequest_Context() As ActionContext
    Set IRequest_Context = Me.Context
End Property

Private Property Get IRequest_Scope() As TargetScope
    IRequest_Scope = TargetScopeSelection
End Property
