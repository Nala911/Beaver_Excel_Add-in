VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: IAppGuard
' @Category: Infrastructure
' @Description: Interface for Application State Guard.

Public Property Let ScreenUpdating(ByVal val As Boolean)
End Property
Public Property Get ScreenUpdating() As Boolean
End Property

Public Property Let EnableEvents(ByVal val As Boolean)
End Property
Public Property Get EnableEvents() As Boolean
End Property

Public Property Let DisplayAlerts(ByVal val As Boolean)
End Property
Public Property Get DisplayAlerts() As Boolean
End Property

Public Property Let Calculation(ByVal val As XlCalculation)
End Property
Public Property Get Calculation() As XlCalculation
End Property

Public Property Let Silent(ByVal val As Boolean)
End Property

Public Sub UpdateState()
End Sub
