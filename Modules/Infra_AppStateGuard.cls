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

Implements IAppGuard

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
    Dim tracker As Object: Set tracker = Infra_Error.Track("AppStateGuard.Class_Initialize")
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
    
    ' Apply Default Silent State
    UpdateState

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "AppStateGuard.Class_Initialize", Err
    Resume CleanExit
End Sub

Private Sub Class_Terminate()
    Dim tracker As Object: Set tracker = Infra_Error.Track("AppStateGuard.Class_Terminate")
    On Error GoTo ErrHandler

    With Application
        .ScreenUpdating = m_OriginalScreenUpdating
        .EnableEvents = m_OriginalEnableEvents
        .DisplayAlerts = m_OriginalDisplayAlerts
        .Calculation = m_OriginalCalculation
    End With

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "AppStateGuard.Class_Terminate", Err
    Resume CleanExit
End Sub

' --- Public properties (for direct access without interface if needed) ---
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
    Dim tracker As Object: Set tracker = Infra_Error.Track("AppStateGuard.UpdateState")
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
    Infra_Error.HandleError "AppStateGuard.UpdateState", Err
    Resume CleanExit
End Sub

' --- IAppGuard Implementation ---

Private Property Let IAppGuard_ScreenUpdating(ByVal val As Boolean)
    Me.ScreenUpdating = val
End Property
Private Property Get IAppGuard_ScreenUpdating() As Boolean
    IAppGuard_ScreenUpdating = Me.ScreenUpdating
End Property

Private Property Let IAppGuard_EnableEvents(ByVal val As Boolean)
    Me.EnableEvents = val
End Property
Private Property Get IAppGuard_EnableEvents() As Boolean
    IAppGuard_EnableEvents = Me.EnableEvents
End Property

Private Property Let IAppGuard_DisplayAlerts(ByVal val As Boolean)
    Me.DisplayAlerts = val
End Property
Private Property Get IAppGuard_DisplayAlerts() As Boolean
    IAppGuard_DisplayAlerts = Me.DisplayAlerts
End Property

Private Property Let IAppGuard_Calculation(ByVal val As XlCalculation)
    Me.Calculation = val
End Property
Private Property Get IAppGuard_Calculation() As XlCalculation
    IAppGuard_Calculation = Me.Calculation
End Property

Private Property Let IAppGuard_Silent(ByVal val As Boolean)
    Me.Silent = val
End Property

Private Sub IAppGuard_UpdateState()
    Me.UpdateState
End Sub
