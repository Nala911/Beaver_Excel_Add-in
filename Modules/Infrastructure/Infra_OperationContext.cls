VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_OperationContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_OperationContext
' @Category: Infrastructure
' @Description: RAII operation scope with diagnostics and breadcrumb cleanup.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Diagnostics, Infra_Error

Private mProcedureName As String
Private mOperationId As String
Private mStartedAt As Double

Public Sub Init(ByVal procName As String)
    mProcedureName = procName
    mOperationId = Infra_Diagnostics.NextOperationId(procName)
    mStartedAt = Timer
    Infra_Diagnostics.LogOperationStart mOperationId, mProcedureName
End Sub

Public Property Get OperationId() As String
    OperationId = mOperationId
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    If mProcedureName <> "" Then
        Infra_Diagnostics.LogOperationFinish mOperationId, mProcedureName, TimerElapsedSeconds(mStartedAt)
        Infra_Error.PopContext
    End If
    On Error GoTo 0
End Sub

Private Function TimerElapsedSeconds(ByVal startedAt As Double) As Double
    TimerElapsedSeconds = Timer - startedAt
    If TimerElapsedSeconds < 0 Then
        TimerElapsedSeconds = TimerElapsedSeconds + 86400#
    End If
End Function
