Attribute VB_Name = "Infra_Progress"
Option Explicit

' @Module: Infra_Progress
' @Category: Infrastructure
' @Description: Centralized progress reporting and UI responsiveness management.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_ProgressState

Private pStateStack As Collection

' Frequency of DoEvents in seconds (e.g., 0.1s for smooth UI without too much overhead)
Private Const UPDATE_INTERVAL As Double = 0.1

' Initializes a new progress tracking session.
Public Sub StartProgress(ByVal Title As String, ByVal TotalSteps As Double)
    Dim tracker As Object: Set tracker = Infra_Error.Track("StartProgress")
    On Error GoTo ErrHandler

    Dim state As Infra_ProgressState
    Set state = New Infra_ProgressState
    state.Title = Title
    state.TotalSteps = TotalSteps
    state.CurrentStep = 0
    state.LastUpdateTime = Timer
    state.UserCancelled = False

    GetStateStack().Add state
    UpdateUI state, True

CleanExit:
    Exit Sub
ErrHandler:
    HandleError "StartProgress", Err
    Resume CleanExit
End Sub

' Updates the current progress and refreshes the UI if the interval has passed.
Public Sub UpdateProgress(ByVal CurrentStep As Double, Optional ByVal ForceRefresh As Boolean = False)
    Dim tracker As Object: Set tracker = Infra_Error.Track("UpdateProgress")
    On Error GoTo ErrHandler

    Dim state As Infra_ProgressState
    Set state = CurrentState()
    If state Is Nothing Then GoTo CleanExit

    state.CurrentStep = CurrentStep

    If ForceRefresh Or TimerElapsedSeconds(state.LastUpdateTime) > UPDATE_INTERVAL Then
        UpdateUI state, False
        state.LastUpdateTime = Timer
        DoEvents ' Keep Excel responsive
    End If

CleanExit:
    Exit Sub
ErrHandler:
    HandleError "UpdateProgress", Err
    Resume CleanExit
End Sub

' Increments progress by one and refreshes UI if needed.
Public Sub Increment()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Increment")
    On Error GoTo ErrHandler

    Dim state As Infra_ProgressState
    Set state = CurrentState()
    If state Is Nothing Then GoTo CleanExit

    UpdateProgress state.CurrentStep + 1

CleanExit:
    Exit Sub
ErrHandler:
    HandleError "Increment", Err
    Resume CleanExit
End Sub

' Cleans up progress reporting and restores the status bar.
Public Sub EndProgress()
    Dim tracker As Object: Set tracker = Infra_Error.Track("EndProgress")
    On Error GoTo ErrHandler

    If pStateStack Is Nothing Then GoTo CleanExit
    If pStateStack.Count = 0 Then GoTo CleanExit

    pStateStack.Remove pStateStack.Count
    If pStateStack.Count = 0 Then
        Application.StatusBar = False
    Else
        UpdateUI CurrentState(), True
    End If

CleanExit:
    Exit Sub
ErrHandler:
    HandleError "EndProgress", Err
    Resume CleanExit
End Sub

' Returns True if the user has requested cancellation (to be implemented with a form later).
Public Property Get UserCancelled() As Boolean
    Dim state As Infra_ProgressState
    Set state = CurrentState()
    If state Is Nothing Then Exit Property
    UserCancelled = state.UserCancelled
End Property

' Internal: Updates the Status Bar. Can be expanded to update a UserForm.
Private Sub UpdateUI(ByVal state As Infra_ProgressState, Optional ByVal IsInitial As Boolean = False)
    ' Internal helper, typically doesn't need its own tracker to avoid overhead in loops
    On Error Resume Next
    
    Dim percent As Integer
    Dim bar As String
    Dim barLength As Integer: barLength = 20
    
    If state Is Nothing Then
        Application.StatusBar = False
        GoTo CleanExit
    End If

    If state.TotalSteps > 0 Then
        percent = Int((state.CurrentStep / state.TotalSteps) * 100)
        If percent > 100 Then percent = 100
        If percent < 0 Then percent = 0
        
        ' Build a simple text-based progress bar: [##########----------]
        Dim filled As Integer: filled = Int((state.CurrentStep / state.TotalSteps) * barLength)
        If filled > barLength Then filled = barLength
        If filled < 0 Then filled = 0
        
        bar = String(filled, ChrW(&H2588)) & String(barLength - filled, ChrW(&H2591))
        
        Application.StatusBar = "BEAVER | " & state.Title & " [" & bar & "] " & percent & "%"
    Else
        Application.StatusBar = "BEAVER | " & state.Title & " (Processing...)"
    End If
    
CleanExit:
    On Error GoTo 0
End Sub

Private Function CurrentState() As Infra_ProgressState
    If pStateStack Is Nothing Then Exit Function
    If pStateStack.Count = 0 Then Exit Function
    Set CurrentState = pStateStack(pStateStack.Count)
End Function

Private Function GetStateStack() As Collection
    If pStateStack Is Nothing Then Set pStateStack = New Collection
    Set GetStateStack = pStateStack
End Function

Private Function TimerElapsedSeconds(ByVal startedAt As Double) As Double
    TimerElapsedSeconds = Timer - startedAt
    If TimerElapsedSeconds < 0 Then
        TimerElapsedSeconds = TimerElapsedSeconds + 86400#
    End If
End Function
