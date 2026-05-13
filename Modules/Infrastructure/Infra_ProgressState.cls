VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_ProgressState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_ProgressState
' @Category: Infrastructure
' @Description: Mutable progress session state for Infra_Progress.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Public Title As String
Public CurrentStep As Double
Public TotalSteps As Double
Public LastUpdateTime As Double
Public UserCancelled As Boolean
