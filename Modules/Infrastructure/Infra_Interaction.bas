Attribute VB_Name = "Infra_Interaction"
Option Explicit

' @Module: Infra_Interaction
' @Category: Infrastructure
' @Description: Centralized user interaction helpers for prompts, confirmations, and notifications.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Config, Infra_Error

Private Const OPTION_PICKER_FORM_NAME As String = "UI_OptionPicker"

Public Function ShowMessage(ByVal message As String, Optional ByVal style As VbMsgBoxStyle = vbInformation, Optional ByVal title As String = vbNullString) As VbMsgBoxResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowMessage")
    On Error GoTo ErrHandler

    If Application.Visible Then
        ShowMessage = MsgBox(message, style, ResolveTitle(title))
    Else
        Debug.Print "BEAVER [INTERACTION]: (" & style & ") " & message
        If (style And vbYesNo) = vbYesNo Then
            ShowMessage = vbYes
        ElseIf (style And vbOKCancel) = vbOKCancel Then
            ShowMessage = vbOK
        ElseIf (style And vbRetryCancel) = vbRetryCancel Then
            ShowMessage = vbRetry
        ElseIf (style And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
            ShowMessage = vbIgnore
        Else
            ShowMessage = vbOK
        End If
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ShowMessage", Err
    Resume CleanExit
End Function

Public Sub ShowInfo(ByVal message As String, Optional ByVal title As String = vbNullString)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowInfo")
    On Error GoTo ErrHandler

    ShowMessage message, vbInformation, title

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ShowInfo", Err
    Resume CleanExit
End Sub

Public Sub ShowWarning(ByVal message As String, Optional ByVal title As String = vbNullString)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowWarning")
    On Error GoTo ErrHandler

    ShowMessage message, vbExclamation, title

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ShowWarning", Err
    Resume CleanExit
End Sub

Public Sub ShowCritical(ByVal message As String, Optional ByVal title As String = vbNullString)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowCritical")
    On Error GoTo ErrHandler

    ShowMessage message, vbCritical, title

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ShowCritical", Err
    Resume CleanExit
End Sub

Public Function Confirm(ByVal message As String, Optional ByVal title As String = vbNullString, Optional ByVal defaultButton As VbMsgBoxStyle = vbDefaultButton1) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("Confirm")
    On Error GoTo ErrHandler

    Confirm = (ShowMessage(message, vbQuestion + vbYesNo + defaultButton, title) = vbYes)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "Confirm", Err
    Resume CleanExit
End Function

Public Function PromptText(ByVal promptMsg As String, ByVal title As String, ByVal defaultText As String, ByRef outResult As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptText")
    On Error GoTo ErrHandler

    Dim result As String
    result = InputBox(promptMsg, ResolveTitle(title), defaultText)
    If StrPtr(result) = 0 Then
        PromptText = False
    Else
        outResult = result
        PromptText = True
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "PromptText", Err
    Resume CleanExit
End Function

Public Function PromptRange(ByVal promptMsg As String, ByVal title As String, ByRef outRange As Range) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptRange")
    On Error GoTo ErrHandler

    Dim selectedRange As Range

    On Error Resume Next
    Set selectedRange = Application.InputBox(Prompt:=promptMsg, Title:=ResolveTitle(title), Type:=8)
    On Error GoTo ErrHandler

    If selectedRange Is Nothing Then
        PromptRange = False
    Else
        Set outRange = selectedRange
        PromptRange = True
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "PromptRange", Err
    Resume CleanExit
End Function

Public Function PromptOption(ByVal promptMsg As String, ByVal title As String, ByVal defaultChoice As String, ByVal options As Variant, ByRef outResult As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptOption")
    On Error GoTo ErrHandler

    Dim frm As Object
    Dim wasConfirmed As Boolean
    Dim selectedValue As String

    On Error Resume Next
    Set frm = VBA.UserForms.Add(OPTION_PICKER_FORM_NAME)
    On Error GoTo ErrHandler

    If frm Is Nothing Then
        Infra_Interaction.ShowCritical "Could not load the option picker form.", ResolveTitle(title)
        GoTo CleanExit
    End If

    frm.ConfigureOptionPicker ResolveTitle(title), promptMsg, defaultChoice, options
    frm.Show

    On Error Resume Next
    wasConfirmed = frm.WasConfirmed
    If Err.Number <> 0 Then
        Err.Clear
        wasConfirmed = False
    ElseIf wasConfirmed Then
        selectedValue = frm.SelectedValue
        If Err.Number <> 0 Then
            Err.Clear
            wasConfirmed = False
        End If
    End If
    On Error GoTo ErrHandler

    If wasConfirmed Then
        outResult = selectedValue
        PromptOption = True
    End If

CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    On Error GoTo 0
    Exit Function

ErrHandler:
    Infra_Error.HandleError "PromptOption", Err
    Resume CleanExit
End Function

Public Function PromptMultiOption( _
    ByVal promptMsg As String, _
    ByVal title As String, _
    ByVal options As Variant, _
    ByVal defaultChecked As Variant, _
    ByRef outSelectedIndices As Variant) As Boolean
    
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptMultiOption")
    On Error GoTo ErrHandler

    Dim frm As Object
    Dim wasConfirmed As Boolean

    On Error Resume Next
    Set frm = VBA.UserForms.Add(OPTION_PICKER_FORM_NAME)
    On Error GoTo ErrHandler

    If frm Is Nothing Then
        Infra_Interaction.ShowCritical "Could not load the option picker form.", ResolveTitle(title)
        GoTo CleanExit
    End If

    frm.ConfigureMultiOptionPicker ResolveTitle(title), promptMsg, options, defaultChecked
    frm.Show

    On Error Resume Next
    wasConfirmed = frm.WasConfirmed
    If Err.Number <> 0 Then
        Err.Clear
        wasConfirmed = False
    ElseIf wasConfirmed Then
        outSelectedIndices = frm.SelectedIndices
        If Err.Number <> 0 Then
            Err.Clear
            wasConfirmed = False
        End If
    End If
    On Error GoTo ErrHandler

    If wasConfirmed Then
        PromptMultiOption = True
    End If

CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    On Error GoTo 0
    Exit Function

ErrHandler:
    Infra_Error.HandleError "PromptMultiOption", Err
    Resume CleanExit
End Function

Public Function PromptSaveAsPath(ByVal dialogTitle As String, ByVal initialPath As String, ByVal fileFilter As String, ByRef outPath As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptSaveAsPath")
    On Error GoTo ErrHandler

    Dim selectedPath As Variant

    selectedPath = Application.GetSaveAsFilename(InitialFileName:=initialPath, FileFilter:=fileFilter, Title:=ResolveTitle(dialogTitle))
    If VarType(selectedPath) = vbBoolean Then
        If selectedPath = False Then GoTo CleanExit
    End If

    outPath = CStr(selectedPath)
    PromptSaveAsPath = (Len(Trim$(outPath)) > 0)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "PromptSaveAsPath", Err
    Resume CleanExit
End Function

' Formats a dialog or message box title consistently with the Add-in brand name.
' If dialogName is provided, returns "Add-in Name - dialogName"
' Otherwise returns "Add-in Name"
Public Function FormatTitle(Optional ByVal dialogName As String = vbNullString) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("FormatTitle")
    On Error GoTo ErrHandler

    Dim nameTrimmed As String
    nameTrimmed = Trim$(dialogName)
    If Len(nameTrimmed) > 0 Then
        FormatTitle = Infra_Config.ADDIN_NAME & " - " & nameTrimmed
    Else
        FormatTitle = Infra_Config.ADDIN_NAME
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "FormatTitle", Err
    Resume CleanExit
End Function

Private Function ResolveTitle(ByVal explicitTitle As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResolveTitle")
    On Error GoTo ErrHandler

    Dim titleTrimmed As String
    titleTrimmed = Trim$(explicitTitle)
    If Len(titleTrimmed) > 0 Then
        If InStr(1, titleTrimmed, Infra_Config.ADDIN_NAME, vbTextCompare) = 1 Then
            ResolveTitle = titleTrimmed
        Else
            ResolveTitle = FormatTitle(titleTrimmed)
        End If
    Else
        ResolveTitle = FormatTitle()
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ResolveTitle", Err
    Resume CleanExit
End Function
