VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_BreakLinksRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_BreakLinksRequest
' @Category: Infrastructure
' @Description: Typed request for break-links workflows.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_ActionContext

Public Context As Infra_ActionContext
Public Scope As BreakLinksScope
