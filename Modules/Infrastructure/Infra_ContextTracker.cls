VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_ContextTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_ContextTracker
' @Category: Infrastructure
' @Description: RAII-style tracker for the execution stack (Infra_Error).
' @ManagedBy: BeaverAddin Agent

Private m_ProcedureName As String

Public Sub Init(ByVal procName As String)
    m_ProcedureName = procName
    ' Note: We don't call PushContext here if it's already done by the factory
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If m_ProcedureName <> "" Then
        Infra_Error.PopContext
    End If
    On Error GoTo 0
End Sub
