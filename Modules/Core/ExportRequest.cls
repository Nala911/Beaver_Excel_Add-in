VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: ExportRequest
' @Category: Core
' @Description: Typed request describing what range to export and which export mode/options to use.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ActionContext

Public Context As ActionContext
Public SourceRange As Range
Public ExportScope As String
Public ExportAsPng As Boolean
Public ScaleFactor As Long
Public OutputPath As String
Public CopyToClipboard As Boolean
