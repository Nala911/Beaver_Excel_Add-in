Attribute VB_Name = "Infra_CommandCatalog"
Option Explicit

' @Module: Infra_CommandCatalog
' @Category: Infrastructure
' @Description: Compatibility facade for command metadata lookups backed by the generated command registry.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_CommandRegistry, Infra_Error

Public Sub ResetCatalog()
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResetCatalog")
    On Error GoTo ErrHandler

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ResetCatalog", Err
    Resume CleanExit
End Sub

Public Function ResolveCommandName(ByVal entryMacro As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResolveCommandName")
    On Error GoTo ErrHandler

    ResolveCommandName = Infra_CommandRegistry.ResolveCommandName(entryMacro)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ResolveCommandName", Err
    Resume CleanExit
End Function
