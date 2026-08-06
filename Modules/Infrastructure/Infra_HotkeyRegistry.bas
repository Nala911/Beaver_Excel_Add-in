Attribute VB_Name = "Infra_HotkeyRegistry"
Option Explicit

' @Module: Infra_HotkeyRegistry
' @Category: Infrastructure
' @Description: Centralized hotkey registration, binding, unbinding, and shortcut lifecycle management.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_ConfigManifest, Infra_HotkeyDefinition, Infra_Error

Private m_RegisteredHotkeys As Collection

''' Binds all hotkeys defined in the compiled configuration manifest.
Public Sub RegisterAllHotkeys()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_HotkeyRegistry.RegisterAllHotkeys")
    On Error GoTo ErrHandler

    UnregisterAllHotkeys

    Set m_RegisteredHotkeys = New Collection
    Dim embeddedHotkeys As Collection
    Set embeddedHotkeys = Infra_ConfigManifest.GetEmbeddedHotkeys()
    If embeddedHotkeys Is Nothing Then GoTo CleanExit

    Dim hk As Infra_HotkeyDefinition
    Dim i As Long
    For i = 1 To embeddedHotkeys.Count
        Set hk = embeddedHotkeys.Item(i)
        If hk.KeyPattern <> "" And hk.MacroName <> "" Then
            On Error Resume Next
            Application.OnKey hk.KeyPattern, hk.MacroName
            If Err.Number = 0 Then
                m_RegisteredHotkeys.Add hk.KeyPattern, hk.KeyPattern
            End If
            On Error GoTo ErrHandler
        End If
    Next i

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Infra_HotkeyRegistry.RegisterAllHotkeys", Err
    Resume CleanExit
End Sub

''' Unbinds all currently active add-in hotkeys from Excel.
Public Sub UnregisterAllHotkeys()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_HotkeyRegistry.UnregisterAllHotkeys")
    On Error GoTo ErrHandler

    If m_RegisteredHotkeys Is Nothing Then Exit Sub

    Dim i As Long
    Dim keyPattern As String
    For i = m_RegisteredHotkeys.Count To 1 Step -1
        keyPattern = CStr(m_RegisteredHotkeys.Item(i))
        On Error Resume Next
        Application.OnKey keyPattern
        On Error GoTo ErrHandler
    Next i

    Set m_RegisteredHotkeys = New Collection

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Infra_HotkeyRegistry.UnregisterAllHotkeys", Err
    Resume CleanExit
End Sub

''' Binds an individual hotkey pattern to a target VBA macro.
Public Sub RegisterHotkey(ByVal keyPattern As String, ByVal macroName As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_HotkeyRegistry.RegisterHotkey")
    On Error GoTo ErrHandler

    If keyPattern = "" Or macroName = "" Then Exit Sub
    If m_RegisteredHotkeys Is Nothing Then Set m_RegisteredHotkeys = New Collection

    On Error Resume Next
    Application.OnKey keyPattern, macroName
    If Err.Number = 0 Then
        On Error Resume Next
        m_RegisteredHotkeys.Remove keyPattern
        On Error GoTo ErrHandler
        m_RegisteredHotkeys.Add keyPattern, keyPattern
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Infra_HotkeyRegistry.RegisterHotkey", Err
    Resume CleanExit
End Sub

''' Unbinds a single hotkey pattern.
Public Sub UnregisterHotkey(ByVal keyPattern As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_HotkeyRegistry.UnregisterHotkey")
    On Error GoTo ErrHandler

    If keyPattern = "" Then Exit Sub

    On Error Resume Next
    Application.OnKey keyPattern
    If Not m_RegisteredHotkeys Is Nothing Then
        m_RegisteredHotkeys.Remove keyPattern
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "Infra_HotkeyRegistry.UnregisterHotkey", Err
    Resume CleanExit
End Sub

''' Checks if a key pattern is currently registered.
Public Function IsHotkeyRegistered(ByVal keyPattern As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_HotkeyRegistry.IsHotkeyRegistered")
    On Error GoTo ErrHandler

    If m_RegisteredHotkeys Is Nothing Then GoTo CleanExit
    On Error Resume Next
    Dim val As String
    val = m_RegisteredHotkeys.Item(keyPattern)
    IsHotkeyRegistered = (Err.Number = 0 And val <> "")
    On Error GoTo ErrHandler

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Infra_HotkeyRegistry.IsHotkeyRegistered", Err
    Resume CleanExit
End Function
