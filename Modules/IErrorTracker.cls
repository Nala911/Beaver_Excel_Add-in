VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: IErrorTracker
' @Category: Infrastructure
' @Description: Interface for tracking and handling errors in a centralized way.

Public Function Track(ByVal SourceName As String) As Object
End Function

Public Sub HandleError(ByVal SourceName As String, ByVal errObj As ErrObject)
End Sub
