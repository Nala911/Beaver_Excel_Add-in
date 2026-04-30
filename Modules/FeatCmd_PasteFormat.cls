VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_PasteFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_PasteFormat
' @Category: Feature
' @Description: Pastes cell formatting only. Requires user to have copied something first.

Private Function ICommand_CanExecute(ByVal context As ICommandContext) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("PasteFormat.CanExecute")
    On Error GoTo ErrHandler
    Dim ctx As Infra_ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    ICommand_CanExecute = Infra_CommandSupport.HasRangeSelectionInContext(ctx)
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PasteFormat.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("PasteFormat.Execute")
    On Error GoTo ErrHandler
    
    Dim targetRange As Range
    Dim ctx As Infra_ActionContext

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit
    Set targetRange = ctx.SelectionRange
    If targetRange Is Nothing Then GoTo CleanExit

    On Error Resume Next
    targetRange.PasteSpecial Paste:=xlPasteFormats
    
    If Err.Number <> 0 Then
        MsgBox "Nothing to paste. Please copy a range first.", vbExclamation, "Beaver Add-in"
    End If
    On Error GoTo ErrHandler

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "PasteFormat.Execute", Err
    Resume CleanExit
End Sub
