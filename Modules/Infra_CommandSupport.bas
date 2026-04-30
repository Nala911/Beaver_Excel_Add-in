Attribute VB_Name = "Infra_CommandSupport"
Option Explicit

' @Module: Infra_CommandSupport
' @Category: Infrastructure
' @Description: Shared helpers for consuming typed command context inside feature commands.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_Error, ICommandContext, Infra_ActionContext

Public Function ActionContextFromCommandContext(ByVal context As ICommandContext) As Infra_ActionContext
    Dim tracker As Object: Set tracker = Infra_Error.Track("ActionContextFromCommandContext")
    On Error GoTo ErrHandler

    If Not context Is Nothing Then
        Set ActionContextFromCommandContext = context.ActionContext
    End If

    If ActionContextFromCommandContext Is Nothing Then
        Set ActionContextFromCommandContext = Infra_AppState.CaptureActionContext()
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ActionContextFromCommandContext", Err
    Resume CleanExit
End Function

Public Function HasRangeSelectionInContext(ByVal ctx As Infra_ActionContext) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("HasRangeSelectionInContext")
    On Error GoTo ErrHandler

    If ctx Is Nothing Then GoTo CleanExit
    HasRangeSelectionInContext = ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "HasRangeSelectionInContext", Err
    Resume CleanExit
End Function
