VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_AppStateGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_AppStateGuard
' @Category: Infrastructure
' @Description: RAII-style wrapper for Excel's global application state.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

Private m_OriginalScreenUpdating As Boolean
Private m_OriginalEnableEvents As Boolean
Private m_OriginalDisplayAlerts As Boolean
Private m_OriginalCalculation As XlCalculation

' Granular Controls (Public to allow overrides)
Public ScreenUpdating As Boolean
Public EnableEvents As Boolean
Public DisplayAlerts As Boolean
Public Calculation As XlCalculation

Private Sub Class_Initialize()
    PushContext "Class_Initialize"
    On Error GoTo ErrHandler

    ' Capture Current State
    With Application
        m_OriginalScreenUpdating = .ScreenUpdating
        m_OriginalEnableEvents = .EnableEvents
        m_OriginalDisplayAlerts = .DisplayAlerts
        m_OriginalCalculation = .Calculation
    End With
    
    ' Default to "Silent Mode" (Everything disabled/manual)
    Me.ScreenUpdating = False
    Me.EnableEvents = False
    Me.DisplayAlerts = False
    Me.Calculation = xlCalculationManual
    
    ' Apply Default Silent State
    UpdateState

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "Class_Initialize", Err
End Sub

' Property to toggle all settings at once (Legacy support + convenience)
Public Property Let Silent(ByVal val As Boolean)
    Me.ScreenUpdating = Not val
    Me.EnableEvents = Not val
    Me.DisplayAlerts = Not val
    Me.Calculation = IIf(val, xlCalculationManual, xlCalculationAutomatic)
    UpdateState
End Property

' Applies current granular settings to the Application.
Public Sub UpdateState()
    PushContext "UpdateState"
    On Error GoTo ErrHandler
    
    With Application
        If .ScreenUpdating <> Me.ScreenUpdating Then .ScreenUpdating = Me.ScreenUpdating
        If .EnableEvents <> Me.EnableEvents Then .EnableEvents = Me.EnableEvents
        If .DisplayAlerts <> Me.DisplayAlerts Then .DisplayAlerts = Me.DisplayAlerts
        If .Calculation <> Me.Calculation Then .Calculation = Me.Calculation
    End With

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "UpdateState", Err
End Sub

Private Sub Class_Terminate()
    PushContext "Class_Terminate"
    On Error GoTo ErrHandler

    With Application
        .ScreenUpdating = m_OriginalScreenUpdating
        .EnableEvents = m_OriginalEnableEvents
        .DisplayAlerts = m_OriginalDisplayAlerts
        .Calculation = m_OriginalCalculation
    End With

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "Class_Terminate", Err
End Sub
