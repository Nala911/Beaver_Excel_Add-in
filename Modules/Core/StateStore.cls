VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StateStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @Module: StateStore
' @Category: Infrastructure
' @Description: Compatibility wrapper around ExcelContextProvider for older callers that still reference StateStore.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ExcelContextProvider, Infra_ActionContext

Public Property Get IsRangeSelected() As Boolean
    IsRangeSelected = ExcelContextProvider.IsRangeSelected
End Property

Public Property Get CanModifyActiveCell() As Boolean
    CanModifyActiveCell = ExcelContextProvider.CanModifyActiveCell
End Property

Public Function CaptureActionContext() As Infra_ActionContext
    Set CaptureActionContext = ExcelContextProvider.CaptureActionContext()
End Function

Public Function CanModifyContext(ByVal ctx As Infra_ActionContext) As Boolean
    CanModifyContext = ExcelContextProvider.CanModifyContext(ctx)
End Function
