VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_CleanDataRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_CleanDataRequest
' @Category: Infrastructure
' @Description: Typed request describing the target scope for a clean-data operation.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_ActionContext

Public Context As Infra_ActionContext
Public Scope As Long
