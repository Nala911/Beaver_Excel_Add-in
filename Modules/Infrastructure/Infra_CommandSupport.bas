Attribute VB_Name = "Infra_CommandSupport"
Option Explicit

' @Module: Infra_CommandSupport
' @Category: Infrastructure
' @Description: Shared helpers for command validation, execution policy, and typed command context access.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_Error, ICommandContext, Infra_ActionContext, CommandExecutionPolicy, CommandValidationResult, Infra_Interaction

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

Public Function CreateExecutionPolicy(Optional ByVal screenUpdating As Boolean = False, Optional ByVal enableEvents As Boolean = False, Optional ByVal displayAlerts As Boolean = True, Optional ByVal calculation As XlCalculation = xlCalculationManual, Optional ByVal useGuard As Boolean = True) As CommandExecutionPolicy
    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateExecutionPolicy")
    On Error GoTo ErrHandler

    Dim policy As New CommandExecutionPolicy
    policy.UseAppStateGuard = useGuard
    policy.ScreenUpdating = screenUpdating
    policy.EnableEvents = enableEvents
    policy.DisplayAlerts = displayAlerts
    policy.Calculation = calculation
    Set CreateExecutionPolicy = policy

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CreateExecutionPolicy", Err
    Resume CleanExit
End Function

' Standard execution policy preset for commands that show user forms, progress dialogs, or require screen updates.
Public Function PolicyInteractiveUI() As CommandExecutionPolicy
    Dim tracker As Object: Set tracker = Infra_Error.Track("PolicyInteractiveUI")
    On Error GoTo ErrHandler

    Set PolicyInteractiveUI = CreateExecutionPolicy(True, False, True)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PolicyInteractiveUI", Err
    Resume CleanExit
End Function

' Standard execution policy preset for background range modifications and silent operations that require max performance.
Public Function PolicyBulkWrite() As CommandExecutionPolicy
    Dim tracker As Object: Set tracker = Infra_Error.Track("PolicyBulkWrite")
    On Error GoTo ErrHandler

    Set PolicyBulkWrite = CreateExecutionPolicy(False, False, True)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PolicyBulkWrite", Err
    Resume CleanExit
End Function

Public Function ValidationSuccess() As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidationSuccess")
    On Error GoTo ErrHandler

    Dim result As New CommandValidationResult
    result.IsExecutable = True
    Set ValidationSuccess = result

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidationSuccess", Err
    Resume CleanExit
End Function

Public Function ValidationFailure(ByVal message As String, Optional ByVal style As VbMsgBoxStyle = vbExclamation, Optional ByVal title As String = vbNullString, Optional ByVal showMessage As Boolean = True) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidationFailure")
    On Error GoTo ErrHandler

    Dim result As New CommandValidationResult
    result.IsExecutable = False
    result.Message = message
    result.MessageStyle = style
    result.Title = title
    result.ShowMessage = showMessage
    Set ValidationFailure = result

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidationFailure", Err
    Resume CleanExit
End Function

Public Sub ShowValidationFailure(ByVal validation As CommandValidationResult)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowValidationFailure")
    On Error GoTo ErrHandler

    If validation Is Nothing Then GoTo CleanExit
    If validation.IsExecutable Then GoTo CleanExit
    If Not validation.ShowMessage Then GoTo CleanExit

    If Len(validation.Message) > 0 Then
        Infra_Interaction.ShowMessage validation.Message, validation.MessageStyle, validation.Title
    End If

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ShowValidationFailure", Err
    Resume CleanExit
End Sub

Public Function ValidateHasRangeSelection(ByVal context As ICommandContext, Optional ByVal message As String = "Please select a range before running this command.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateHasRangeSelection")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If HasRangeSelectionInContext(ctx) Then
        Set ValidateHasRangeSelection = ValidationSuccess()
    Else
        Set ValidateHasRangeSelection = ValidationFailure(message)
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateHasRangeSelection", Err
    Set ValidateHasRangeSelection = ValidationFailure(message)
    Resume CleanExit
End Function

Public Function ValidateHasWorkbook(ByVal context As ICommandContext, Optional ByVal message As String = "There is no active workbook available for this command.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateHasWorkbook")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not ctx Is Nothing Then
        If Not ctx.WorkbookRef Is Nothing Then
            Set ValidateHasWorkbook = ValidationSuccess()
            Exit Function
        End If
    End If

    Set ValidateHasWorkbook = ValidationFailure(message)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateHasWorkbook", Err
    Set ValidateHasWorkbook = ValidationFailure(message)
    Resume CleanExit
End Function

Public Function ValidateHasWorksheet(ByVal context As ICommandContext, Optional ByVal message As String = "There is no active worksheet available for this command.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateHasWorksheet")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not ctx Is Nothing Then
        If Not ctx.WorksheetRef Is Nothing Then
            Set ValidateHasWorksheet = ValidationSuccess()
            Exit Function
        End If
    End If

    Set ValidateHasWorksheet = ValidationFailure(message)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateHasWorksheet", Err
    Set ValidateHasWorksheet = ValidationFailure(message)
    Resume CleanExit
End Function

Public Function ValidateHasWindow(ByVal context As ICommandContext, Optional ByVal message As String = "There is no active window available for this command.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateHasWindow")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not ctx Is Nothing Then
        If Not ctx.WindowRef Is Nothing Then
            Set ValidateHasWindow = ValidationSuccess()
            Exit Function
        End If
    End If

    Set ValidateHasWindow = ValidationFailure(message)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateHasWindow", Err
    Set ValidateHasWindow = ValidationFailure(message)
    Resume CleanExit
End Function

Public Function ValidateCanModifySelection(ByVal context As ICommandContext, Optional ByVal message As String = "The current selection cannot be modified.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateCanModifySelection")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not HasRangeSelectionInContext(ctx) Then
        Set ValidateCanModifySelection = ValidationFailure("Please select a range before running this command.")
    ElseIf Infra_AppState.CanModifyContext(ctx) Then
        Set ValidateCanModifySelection = ValidationSuccess()
    Else
        Set ValidateCanModifySelection = ValidationFailure(message)
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateCanModifySelection", Err
    Set ValidateCanModifySelection = ValidationFailure(message)
    Resume CleanExit
End Function
