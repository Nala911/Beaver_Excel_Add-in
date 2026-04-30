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

Private Function ICommand_CanExecute(ByVal context As ICommandContext) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ApplyCustomNumberFormat.CanExecute")
    On Error GoTo ErrHandler
    Dim ctx As Infra_ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    ICommand_CanExecute = Infra_CommandSupport.HasRangeSelectionInContext(ctx)
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ApplyCustomNumberFormat.CanExecute", Err
    ICommand_CanExecute = False
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

    targetRange.NumberFormat = Infra_Config.Model.DefaultNumberFormat

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ApplyCustomNumberFormat.Execute", Err
    Resume CleanExit
End Sub
