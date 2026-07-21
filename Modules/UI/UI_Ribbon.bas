Attribute VB_Name = "UI_Ribbon"
Option Explicit

' @Module: UI_Ribbon
' @Category: UI
' @Description: Generated Ribbon callbacks for the Beaver Add-in.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: AppContainer, Infra_Config, Infra_Error

' --- Dynamic UI Callbacks ---

' Returns the image object for a control based on its ID in config.json
Public Sub Ribbon_GetIcon(ByVal control As Object, ByRef image As Variant)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_GetIcon")
    On Error GoTo ErrHandler
    
    Dim iconName As String
    iconName = Infra_Config.GetIcon(control.Id)
    If iconName = "" Then iconName = "Help"
    
    Set image = Application.CommandBars.GetImageMso(iconName, 32, 32)
    
CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_GetIcon", Err
    Resume CleanExit
End Sub

Public Sub Ribbon_OnAction(ByVal control As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Ribbon_OnAction")
    On Error GoTo ErrHandler

    EnsureUIServicesRegistered
    AppContainer.ExecuteEntryPoint control.Id, control.Id, "Ribbon"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Ribbon_OnAction", Err
    Resume CleanExit
End Sub

Public Sub EnsureUIServicesRegistered()
    Dim tracker As Object: Set tracker = Infra_Error.Track("EnsureUIServicesRegistered")
    On Error GoTo ErrHandler

    AppContainer.Register "IUIFactory", UI_Factory
    AppContainer.Register "IUIHelpCenter", UI_HelpCenter

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "EnsureUIServicesRegistered", Err
    Resume CleanExit
End Sub

Public Function GetUIFactory() As Object
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetUIFactory")
    On Error GoTo ErrHandler

    Set GetUIFactory = UI_Factory

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetUIFactory", Err
    Resume CleanExit
End Function

Public Function GetUIHelpCenter() As Object
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetUIHelpCenter")
    On Error GoTo ErrHandler

    Set GetUIHelpCenter = UI_HelpCenter

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetUIHelpCenter", Err
    Resume CleanExit
End Function