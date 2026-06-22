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

Private m_OriginalScreenUpdating As Boolean
Private m_OriginalEnableEvents As Boolean
Private m_OriginalDisplayAlerts As Boolean
Private m_OriginalCalculation As XlCalculation

Private m_ScreenUpdating As Boolean
Private m_EnableEvents As Boolean
Private m_DisplayAlerts As Boolean
Private m_Calculation As XlCalculation

Private Sub Class_Initialize()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_AppStateGuard.Class_Initialize")
    On Error GoTo ErrHandler

    ' Capture Current State
    With Application
        m_OriginalScreenUpdating = .ScreenUpdating
        m_OriginalEnableEvents = .EnableEvents
        m_OriginalDisplayAlerts = .DisplayAlerts
        m_OriginalCalculation = .Calculation
    End With
    
    ' Default to "Silent Mode" (Everything disabled/manual)
    m_ScreenUpdating = False
    m_EnableEvents = False
    m_DisplayAlerts = False
    m_Calculation = xlCalculationManual
    
    ' Register nested level and capture outer-most states
    If Infra_AppState.GuardNestingLevel = 0 Then
        Infra_AppState.OriginalScreenUpdating = m_OriginalScreenUpdating
        Infra_AppState.OriginalEnableEvents = m_OriginalEnableEvents
        Infra_AppState.OriginalDisplayAlerts = m_OriginalDisplayAlerts
        Infra_AppState.OriginalCalculation = m_OriginalCalculation
    End If
    Infra_AppState.GuardNestingLevel = Infra_AppState.GuardNestingLevel + 1
    
    ' Apply Default Silent State
    UpdateState

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Infra_AppStateGuard.Class_Initialize", Err
    Resume CleanExit
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    ' Clear active guard registration
    Infra_AppState.GuardNestingLevel = Infra_AppState.GuardNestingLevel - 1
    If Infra_AppState.GuardNestingLevel < 0 Then Infra_AppState.GuardNestingLevel = 0

    ' Restore original application state in a failsafe manner
    With Application
        If .ScreenUpdating <> m_OriginalScreenUpdating Then .ScreenUpdating = m_OriginalScreenUpdating
        If .EnableEvents <> m_OriginalEnableEvents Then .EnableEvents = m_OriginalEnableEvents
        If .DisplayAlerts <> m_OriginalDisplayAlerts Then .DisplayAlerts = m_OriginalDisplayAlerts
        If .Calculation <> m_OriginalCalculation Then .Calculation = m_OriginalCalculation
    End With
    On Error GoTo 0
End Sub

' --- Public properties (for direct access without interface if needed) ---
Public Property Get OriginalCalculation() As XlCalculation
    OriginalCalculation = m_OriginalCalculation
End Property

Public Property Get ScreenUpdating() As Boolean
    ScreenUpdating = m_ScreenUpdating
End Property
Public Property Let ScreenUpdating(ByVal val As Boolean)
    m_ScreenUpdating = val
End Property

Public Property Get EnableEvents() As Boolean
    EnableEvents = m_EnableEvents
End Property
Public Property Let EnableEvents(ByVal val As Boolean)
    m_EnableEvents = val
End Property

Public Property Get DisplayAlerts() As Boolean
    DisplayAlerts = m_DisplayAlerts
End Property
Public Property Let DisplayAlerts(ByVal val As Boolean)
    m_DisplayAlerts = val
End Property

Public Property Get Calculation() As XlCalculation
    Calculation = m_Calculation
End Property
Public Property Let Calculation(ByVal val As XlCalculation)
    m_Calculation = val
End Property

Public Property Let Silent(ByVal val As Boolean)
    m_ScreenUpdating = Not val
    m_EnableEvents = Not val
    m_DisplayAlerts = Not val
    m_Calculation = IIf(val, xlCalculationManual, xlCalculationAutomatic)
    UpdateState
End Property

Public Sub UpdateState()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_AppStateGuard.UpdateState")
    On Error GoTo ErrHandler
    
    With Application
        If .ScreenUpdating <> m_ScreenUpdating Then .ScreenUpdating = m_ScreenUpdating
        If .EnableEvents <> m_EnableEvents Then .EnableEvents = m_EnableEvents
        If .DisplayAlerts <> m_DisplayAlerts Then .DisplayAlerts = m_DisplayAlerts
        If .Calculation <> m_Calculation Then .Calculation = m_Calculation
    End With

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Infra_AppStateGuard.UpdateState", Err
    Resume CleanExit
End Sub
