VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_ToggleFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_ToggleFullScreen
' @Category: Feature
' @Description: Toggles Excel's full screen (Ribbon and status bar) visibility.

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ToggleFullScreen.CanExecute")
    On Error GoTo ErrHandler
    ICommand_CanExecute = Not (ActiveWindow Is Nothing)
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ToggleFullScreen.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ToggleFullScreen.Execute")
    On Error GoTo ErrHandler
    
    If ActiveWindow Is Nothing Then GoTo CleanExit

    With ActiveWindow
        .DisplayHeadings = Not .DisplayHeadings
        .DisplayWorkbookTabs = Not .DisplayWorkbookTabs
        .DisplayHorizontalScrollBar = Not .DisplayHorizontalScrollBar
        .DisplayVerticalScrollBar = Not .DisplayVerticalScrollBar
    End With
    
    Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ToggleFullScreen.Execute", Err
    Resume CleanExit
End Sub
