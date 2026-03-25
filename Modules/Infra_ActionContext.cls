VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_ActionContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_ActionContext
' @Category: Infrastructure
' @Description: Typed snapshot of the active Excel workbook, sheet, and selection state for a macro invocation.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Public WorkbookRef As Workbook
Public WorksheetRef As Worksheet
Public SelectionRange As Range
Public ActiveCellRef As Range
Public HasRangeSelection As Boolean
