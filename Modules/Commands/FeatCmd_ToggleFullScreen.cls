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
' @Description: Toggles a focused full-screen workspace by hiding or restoring available Excel chrome.

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

    Dim enterFocusMode As Boolean

    enterFocusMode = (ctx.WindowRef.DisplayHeadings Or _
                      ctx.WindowRef.DisplayWorkbookTabs Or _
                      ctx.WindowRef.DisplayHorizontalScrollBar Or _
                      ctx.WindowRef.DisplayVerticalScrollBar Or _
                      Application.DisplayFormulaBar Or _
                      Application.DisplayStatusBar Or _
                      Not Application.DisplayFullScreen)

    ApplyFocusMode ctx.WindowRef, enterFocusMode
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ToggleFullScreen.Execute", Err
    Resume CleanExit
End Sub

Private Sub ApplyFocusMode(ByVal targetWindow As Window, ByVal enableFocusMode As Boolean)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ApplyFocusMode")
    On Error GoTo ErrHandler

    If targetWindow Is Nothing Then GoTo CleanExit

    With targetWindow
        .DisplayHeadings = Not enableFocusMode
        .DisplayWorkbookTabs = Not enableFocusMode
        .DisplayHorizontalScrollBar = Not enableFocusMode
        .DisplayVerticalScrollBar = Not enableFocusMode
    End With

    Application.DisplayFormulaBar = Not enableFocusMode
    Application.DisplayStatusBar = Not enableFocusMode
    Application.DisplayFullScreen = enableFocusMode
    SetRibbonVisible Not enableFocusMode

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ApplyFocusMode", Err
    Resume CleanExit
End Sub

Private Sub SetRibbonVisible(ByVal isVisible As Boolean)
    Dim tracker As Object: Set tracker = Infra_Error.Track("SetRibbonVisible")
    On Error GoTo ErrHandler

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon""," & IIf(isVisible, "True", "False") & ")"

CleanExit:
    Exit Sub

ErrHandler:
    ' Some Excel environments do not allow Ribbon toggling through Excel4 macros.
    Err.Clear
    Resume CleanExit
End Sub
