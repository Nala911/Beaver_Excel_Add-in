VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_ModifyDataRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_ModifyDataRequest
' @Category: Core
' @Description: Typed request describing the target selection, action context, and data modification options.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_ActionContext

Public Context As Infra_ActionContext
Public Operation As String
Public DatePattern As String
