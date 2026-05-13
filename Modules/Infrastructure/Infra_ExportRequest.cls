VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_ExportRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_ExportRequest
' @Category: Infrastructure
' @Description: Typed request describing what range to export and which export mode/options to use.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_ActionContext

Public Context As Infra_ActionContext
Public SourceRange As Range
Public ExportAsPng As Boolean
Public ScaleFactor As Long
