VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: IConfig
' @Category: Infrastructure
' @Description: Interface for accessing application configuration.

Public Property Get Model() As Object
End Property

Public Property Get GetIcon(ByVal controlId As String) As String
End Property
