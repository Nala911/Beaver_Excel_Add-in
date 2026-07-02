VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandInvoker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: CommandInvoker
' @Category: Core
' @Description: Centralized pipeline for executing commands. Handles guards and errors.

Public Sub ExecuteCommand(ByVal cmd As ICommand, ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteCommand")
    On Error GoTo ErrHandler

    Dim validation As CommandValidationResult
    Dim policy As CommandExecutionPolicy
    Dim guard As Infra_AppStateGuard
    
    Dim opId As String
    Dim startTime As Double
    Dim commandName As String
    Dim successFlag As Boolean
    
    If cmd Is Nothing Then GoTo CleanExit
    If context Is Nothing Then GoTo CleanExit
    
    commandName = context.CommandName
    If commandName = vbNullString Then commandName = TypeName(cmd)
    opId = Infra_Diagnostics.NextOperationId(commandName)
    startTime = Timer
    successFlag = False
    
    Infra_Diagnostics.LogOperationStart opId, commandName
    
    ' Clear transient cache at the start of command execution
    Infra_Cache.Clear
    
    Infra_Undo.ClearPendingUndo

    Set validation = cmd.Validate(context)
    If validation Is Nothing Then
        Set validation = Infra_CommandSupport.ValidationSuccess()
    End If

    If Not validation.IsExecutable Then
        Infra_CommandSupport.ShowValidationFailure validation
        GoTo CleanExit
    End If

    Set policy = cmd.ExecutionPolicy
    If policy Is Nothing Then Set policy = Infra_CommandSupport.CreateExecutionPolicy()

    If policy.UseAppStateGuard Then
        Set guard = New Infra_AppStateGuard
        guard.ScreenUpdating = policy.ScreenUpdating
        guard.EnableEvents = policy.EnableEvents
        guard.DisplayAlerts = policy.DisplayAlerts
        guard.Calculation = policy.Calculation
        guard.UpdateState
    End If

    cmd.Execute context
    
    ' Ensure AppStateGuard is terminated so Application state is restored
    ' before we register the Undo macro.
    If Not guard Is Nothing Then Set guard = Nothing
    
    Infra_Undo.RegisterPendingUndo
    
    successFlag = True

CleanExit:
    If Not guard Is Nothing Then Set guard = Nothing
    
    If opId <> vbNullString Then
        Dim elapsed As Double
        elapsed = Infra_Diagnostics.TimerElapsedSeconds(startTime)
        
        Dim cellsCount As Double
        cellsCount = 0
        On Error Resume Next
        If TypeName(Selection) = "Range" Then
            cellsCount = Selection.Cells.CountLarge
        End If
        Err.Clear
        On Error GoTo 0
        
        Dim wbName As String
        wbName = ""
        On Error Resume Next
        wbName = ActiveWorkbook.Name
        Err.Clear
        On Error GoTo 0
        
        Dim payload As String
        payload = "{" & _
                  """elapsed_seconds"":" & Format$(elapsed, "0.000") & "," & _
                  """cells_count"":" & cellsCount & "," & _
                  """workbook"":""" & Infra_Diagnostics.EscapeJson(wbName) & """," & _
                  Trim$("""status"":""" & IIf(successFlag, "success", "failure") & """") & _
                  "}"
        Infra_Diagnostics.LogEvent "operation_finish", commandName, opId, payload
    End If
    Exit Sub

ErrHandler:
    Infra_Undo.ClearPendingUndo
    Infra_Error.HandleError "ExecuteCommand", Err
    Resume CleanExit
End Sub
