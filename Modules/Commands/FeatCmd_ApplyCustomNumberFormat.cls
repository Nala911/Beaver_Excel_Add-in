VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_ApplyCustomNumberFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_ApplyCustomNumberFormat
' @Category: Feature
' @Description: Applies a custom number format to the selected range.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyBulkWrite()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ApplyCustomNumberFormat.Execute")
    On Error GoTo ErrHandler
    
    Dim targetRange As Range
    Dim ctx As Infra_ActionContext

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit
    Set targetRange = ctx.SelectionRange
    If targetRange Is Nothing Then GoTo CleanExit

    Infra_Undo.SaveState targetRange, "Apply Format"
    targetRange.NumberFormat = Infra_Config.Model.DefaultNumberFormat

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ApplyCustomNumberFormat.Execute", Err
    Resume CleanExit
End Sub
