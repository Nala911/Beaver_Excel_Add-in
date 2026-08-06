Attribute VB_Name = "Infra_RibbonManager"
Option Explicit

' @Module: Infra_RibbonManager
' @Category: Infrastructure
' @Description: Centralized ribbon state manager, UI invalidation, and dynamic callback handler.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Config, Infra_Error

Private m_RibbonUI As Object

''' Stores reference to IRibbonUI instance from onLoad callback.
Public Sub SetRibbon(ByVal ribbon As Object)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_RibbonManager.SetRibbon")
    On Error GoTo ErrHandler

    Set m_RibbonUI = ribbon

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Infra_RibbonManager.SetRibbon", Err
    Resume CleanExit
End Sub

''' Gets current stored IRibbonUI reference.
Public Function GetRibbon() As Object
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_RibbonManager.GetRibbon")
    On Error GoTo ErrHandler

    Set GetRibbon = m_RibbonUI

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Infra_RibbonManager.GetRibbon", Err
    Resume CleanExit
End Function

''' Forces full ribbon UI refresh.
Public Sub InvalidateUI()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_RibbonManager.InvalidateUI")
    On Error GoTo ErrHandler

    If Not m_RibbonUI Is Nothing Then
        m_RibbonUI.Invalidate
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Infra_RibbonManager.InvalidateUI", Err
    Resume CleanExit
End Sub

''' Invalidates a single ribbon control by ID.
Public Sub InvalidateControl(ByVal controlId As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_RibbonManager.InvalidateControl")
    On Error GoTo ErrHandler

    If Not m_RibbonUI Is Nothing And controlId <> "" Then
        m_RibbonUI.InvalidateControl controlId
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Infra_RibbonManager.InvalidateControl", Err
    Resume CleanExit
End Sub

''' Resolves dynamic enabled state for a ribbon control based on context.
Public Function GetControlEnabled(ByVal controlId As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_RibbonManager.GetControlEnabled")
    On Error GoTo ErrHandler

    GetControlEnabled = True
    If Application.ActiveWorkbook Is Nothing Then
        GetControlEnabled = False
        GoTo CleanExit
    End If
    If Application.ActiveSheet Is Nothing Then
        GetControlEnabled = False
        GoTo CleanExit
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Infra_RibbonManager.GetControlEnabled", Err
    Resume CleanExit
End Function

''' Resolves dynamic visible state for a ribbon control.
Public Function GetControlVisible(ByVal controlId As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_RibbonManager.GetControlVisible")
    On Error GoTo ErrHandler

    GetControlVisible = True

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Infra_RibbonManager.GetControlVisible", Err
    Resume CleanExit
End Function

''' Resolves dynamic icon name for a control.
Public Function GetControlIcon(ByVal controlId As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_RibbonManager.GetControlIcon")
    On Error GoTo ErrHandler

    Dim iconName As String
    iconName = Infra_Config.GetIcon(controlId)
    If iconName = "" Then iconName = "Help"
    GetControlIcon = iconName

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Infra_RibbonManager.GetControlIcon", Err
    Resume CleanExit
End Function
