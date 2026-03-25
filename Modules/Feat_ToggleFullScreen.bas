Attribute VB_Name = "Feat_ToggleFullScreen"
Option Explicit

' @Module: Feat_ToggleFullScreen
' @Category: Feature
' @Description: Toggles Excel's full screen (Ribbon and status bar) visibility.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

Public Sub ToggleFullScreen()
    PushContext "ToggleFullScreen"
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
    PopContext
    Exit Sub

ErrHandler:
    HandleError "ToggleFullScreen", Err
End Sub
