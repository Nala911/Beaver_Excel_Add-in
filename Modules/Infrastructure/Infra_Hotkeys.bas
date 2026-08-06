Attribute VB_Name = "Infra_Hotkeys"
Option Explicit

' @Module: Infra_Hotkeys
' @Category: Infrastructure
' @Description: Central registry for all keyboard shortcuts, now loaded from config.json.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Config, Infra_Error, Lib_UdfRegistry

Private pAreHotkeysRegistered As Boolean

' Returns a 2D array of all hotkey definitions from JSON.
'   Column 1 = OnKey pattern  (e.g. "^+p")
'   Column 2 = Macro name     (e.g. "Feat_MakeItStatic.MakePermanent")
'   Column 3 = Human-readable description for ShowHelpCenter
Public Function HotkeyDefinitions() As Variant
    Dim tracker As Object: Set tracker = Infra_Error.Track("HotkeyDefinitions")
    On Error GoTo ErrHandler
    
    Dim hkColl As Collection
    Set hkColl = Infra_Config.Hotkeys
    
    If hkColl Is Nothing Then
        HotkeyDefinitions = Array()
        GoTo CleanExit
    End If
    
    Dim count As Long
    count = hkColl.count
    
    If count = 0 Then
        HotkeyDefinitions = Array()
        GoTo CleanExit
    End If
    
    Dim defs() As Variant
    ReDim defs(1 To count, 1 To 3)
    
    Dim i As Long
    Dim item As Infra_HotkeyDefinition
    For i = 1 To count
        Set item = hkColl(i)
        defs(i, 1) = item.KeyPattern
        defs(i, 2) = item.MacroName
        defs(i, 3) = item.Description
    Next i
    
    HotkeyDefinitions = defs

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "HotkeyDefinitions", Err
    Resume CleanExit
End Function

Public Sub EnsureHotkeysRegistered()
    Dim tracker As Object: Set tracker = Infra_Error.Track("EnsureHotkeysRegistered")
    On Error GoTo ErrHandler

    If Not pAreHotkeysRegistered Then
        RegisterHotkeys
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "EnsureHotkeysRegistered", Err
    Resume CleanExit
End Sub

' Binds all shortcuts defined in HotkeyDefinitions via Application.OnKey.
' Called by ThisWorkbook.Workbook_Open.
Public Sub RegisterHotkeys()
    Dim tracker As Object: Set tracker = Infra_Error.Track("RegisterHotkeys")
    On Error GoTo ErrHandler
    
    Dim defs As Variant
    Dim i As Long

    defs = HotkeyDefinitions()
    If IsEmpty(defs) Then GoTo CleanExit

    On Error Resume Next
    For i = LBound(defs, 1) To UBound(defs, 1)
        If defs(i, 1) <> "" And defs(i, 2) <> "" Then
            Dim macroName As String
            macroName = CStr(defs(i, 2))
            
            ' Strip module prefix (e.g. "UI_Hotkeys.Hotkey_Name" -> "Hotkey_Name")
            ' to ensure Excel can resolve the macro call in add-in/hidden workbook mode
            Dim dotPos As Long
            dotPos = InStr(macroName, ".")
            If dotPos > 0 Then
                macroName = Mid(macroName, dotPos + 1)
            End If
            
            ' Register with workbook qualification so Excel routes to this add-in
            Application.OnKey CStr(defs(i, 1)), "'" & ThisWorkbook.Name & "'!" & macroName
            ' Also register without workbook qualification so global macro namespace resolves in add-in mode
            Application.OnKey CStr(defs(i, 1)), macroName
        End If
    Next i
    On Error GoTo ErrHandler
    pAreHotkeysRegistered = True

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "RegisterHotkeys", Err
    Resume CleanExit
End Sub

' Clears all shortcut bindings by passing "" as the procedure to Application.OnKey.
' Called by ThisWorkbook.Workbook_BeforeClose.
Public Sub UnregisterHotkeys()
    Dim tracker As Object: Set tracker = Infra_Error.Track("UnregisterHotkeys")
    On Error GoTo ErrHandler
    
    Dim defs As Variant
    Dim i As Long

    defs = HotkeyDefinitions()
    If IsEmpty(defs) Then GoTo CleanExit

    On Error Resume Next
    For i = LBound(defs, 1) To UBound(defs, 1)
        If defs(i, 1) <> "" Then
            Application.OnKey CStr(defs(i, 1)), ""
        End If
    Next i
    On Error GoTo ErrHandler
    pAreHotkeysRegistered = False

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "UnregisterHotkeys", Err
    Resume CleanExit
End Sub

' Shows a human-readable list of all shortcuts, UDFs, and diagnostics in a scrollable UserForm.
Public Sub ShowHelpCenter()
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowHelpCenter")
    On Error GoTo ErrHandler
    
    Dim uiHelpCenter As Object
    Set uiHelpCenter = AppContainer.Resolve("IUIHelpCenter")
    If Not uiHelpCenter Is Nothing Then
        uiHelpCenter.DisplayHelp
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "ShowHelpCenter", Err
    Resume CleanExit
End Sub

' Converts Application.OnKey patterns (like "^+p") into human-readable text (like "Ctrl + Shift + P").
Public Function TranslateHotkey(ByVal pattern As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("TranslateHotkey")
    On Error GoTo ErrHandler
    
    Dim modifiers As String
    Dim key As String
    Dim char As String
    
    key = pattern
    modifiers = ""
    
    ' Extract modifiers from the beginning (Order: Ctrl, Alt, Shift)
    Do While Len(key) > 0
        char = Left(key, 1)
        If char = "^" Then
            modifiers = modifiers & "Ctrl + "
            key = Mid(key, 2)
        ElseIf char = "%" Then
            modifiers = modifiers & "Alt + "
            key = Mid(key, 2)
        ElseIf char = "+" Then
            modifiers = modifiers & "Shift + "
            key = Mid(key, 2)
        Else
            Exit Do
        End If
    Loop
    
    ' Handle special keys in braces (e.g., {+})
    key = Replace(key, "{", "")
    key = Replace(key, "}", "")
    
    ' Capitalize lone letters
    If Len(key) = 1 Then
        If key >= "a" And key <= "z" Then key = UCase(key)
    End If
    
    TranslateHotkey = modifiers & key

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "TranslateHotkey", Err
    Resume CleanExit
End Function
