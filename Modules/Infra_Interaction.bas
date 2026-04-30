Attribute VB_Name = "Infra_Interaction"
Option Explicit

' @Module: Infra_Interaction
' @Category: Infrastructure
' @Description: Centralized user interaction helpers for prompts, confirmations, and notifications.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Config, Infra_Error

Public Function ShowMessage(ByVal message As String, Optional ByVal style As VbMsgBoxStyle = vbInformation, Optional ByVal title As String = vbNullString) As VbMsgBoxResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowMessage")
    On Error GoTo ErrHandler

    ShowMessage = MsgBox(message, style, ResolveTitle(title))

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

Private Function ResolveTitle(ByVal explicitTitle As String) As String
    If Len(Trim$(explicitTitle)) > 0 Then
        ResolveTitle = explicitTitle
    Else
        ResolveTitle = Infra_Config.ADDIN_NAME
    End If
End Function
