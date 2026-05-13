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

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWindow(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ToggleFullScreen.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As Infra_ActionContext

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WindowRef Is Nothing Then GoTo CleanExit

    Dim targetState As Boolean
    
    ' Determine target state based on current heading visibility.
    ' If headings are visible, we are entering Focus Mode (targetState = False).
    ' If headings are hidden, we are leaving Focus Mode (targetState = True).
    targetState = Not ctx.WindowRef.DisplayHeadings

    With ctx.WindowRef
        .DisplayHeadings = targetState
        .DisplayWorkbookTabs = targetState
        .DisplayHorizontalScrollBar = targetState
        .DisplayVerticalScrollBar = targetState
    End With
    
    Application.DisplayFormulaBar = targetState
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ToggleFullScreen.Execute", Err
    Resume CleanExit
End Sub
