VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICellTransformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: ICellTransformer
' @Category: Core
' @Description: Interface for a cell value transformer used by the unified batch processor engine.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Public Function TransformCell( _
    ByVal rawValue As Variant, _
    ByVal currentFormat As String, _
    ByRef outValue As Variant, _
    ByRef outFormat As String) As Boolean
End Function
