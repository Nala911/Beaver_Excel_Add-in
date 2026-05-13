VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_ErrorContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' @Module: Infra_ErrorContext
' @Category: Infrastructure
' @Description: Structured payload describing a single runtime error instance.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

'===================================================================================================='
'|     Class: Infra_ErrorContext
'|   Purpose: Holds all relevant information about a single error event in a structured way.
'|
'===================================================================================================='

Option Explicit

Public Timestamp As Date
Public ProcedureName As String
Public Severity As String
Public Category As String
Public ErrorNumber As Long
Public ErrorDescription As String
Public ErrorSource As String
Public Username As String
Public AppVersion As String
Public Breadcrumbs As String
Public WorkbookName As String
Public WorksheetName As String
Public SelectionAddress As String
Public StateDump As String
