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

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasRangeSelection(context)
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
        Err.Clear
        Infra_Interaction.ShowWarning "Nothing to paste. Please copy a range first."
    End If
    On Error GoTo ErrHandler

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "PasteFormat.Execute", Err
    Resume CleanExit
End Sub
