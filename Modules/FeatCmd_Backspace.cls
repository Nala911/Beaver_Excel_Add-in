VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_Backspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_Backspace
' @Category: Feature
' @Description: Custom handlers for Backspace with protection awareness.

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("Backspace.CanExecute")
    On Error GoTo ErrHandler
    ICommand_CanExecute = Infra_AppState.CanModifyActiveCell()
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Backspace.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Backspace.Execute")
    On Error GoTo ErrHandler
    
    If Not Infra_AppState.CanModifyActiveCell() Then GoTo CleanExit
    
    ' Checks if the current selection is a range of cells
    If TypeName(Selection) = "Range" Then
        Selection.ClearContents
    End If
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Backspace.Execute", Err
    Resume CleanExit
End Sub
