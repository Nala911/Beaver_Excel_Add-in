VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: IRequest
' @Category: Core
' @Description: Interface for all typed request objects in the command pipeline.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ActionContext, TargetScope

Public Property Get Context() As ActionContext
End Property

Public Property Get Scope() As TargetScope
End Property
