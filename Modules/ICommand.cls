VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: ICommand
' @Category: Infrastructure
' @Description: Standard command pattern interface for all features.

Public Sub Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Execute")
    On Error GoTo ErrHandler
CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Execute", Err
    Resume CleanExit
End Sub

Public Function CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("CanExecute")
    On Error GoTo ErrHandler
CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "CanExecute", Err
    Resume CleanExit
End Function
