VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBatchProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: IBatchProcessor
' @Category: Core
' @Description: Interface for custom processing logic executed by Infra_BatchProcessor.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Public Function ProcessRange(ByVal targetRange As Range) As Long
End Function

Public Function ProcessWorksheet(ByVal ws As Worksheet) As Long
End Function
