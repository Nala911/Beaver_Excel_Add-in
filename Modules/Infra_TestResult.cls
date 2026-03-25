VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_TestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_TestResult
' @Category: Infrastructure
' @Description: Simple object to hold a single unit test result for Lib_Tests.
' @ManagedBy: BeaverAddin Agent

Public Name As String
Public Passed As Boolean
Public Message As String
Public DurationMs As Long
Public Category As String
