VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StateStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @Module: StateStore
' @Category: Infrastructure
' @Description: Centralized State Management holding application-level state and UI toggles.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_ActionContext

Private m_IsRangeSelected As Boolean
Private m_CanModifyActiveCell As Boolean
Private m_ActionContext As Infra_ActionContext

Private Sub Class_Initialize()
    ' Wait until explicit operations populate state
End Sub

Public Property Get IsRangeSelected() As Boolean
    IsRangeSelected = (TypeName(Selection) = "Range")
End Property

Public Property Get CanModifyActiveCell() As Boolean
    CanModifyActiveCell = CanModifyContext(CaptureActionContext())
End Property

Public Function CaptureActionContext() As Infra_ActionContext
    Dim ctx As New Infra_ActionContext

    If Not ActiveWorkbook Is Nothing Then Set ctx.WorkbookRef = ActiveWorkbook
    If TypeName(ActiveSheet) = "Worksheet" Then Set ctx.WorksheetRef = ActiveSheet

    On Error Resume Next
    ctx.SelectionTypeName = TypeName(Selection)
    Set ctx.SelectionRef = Selection
    If Not ActiveWindow Is Nothing Then Set ctx.WindowRef = ActiveWindow
    On Error GoTo 0

    If ctx.SelectionTypeName = "Range" Then
        ctx.HasRangeSelection = True
        Set ctx.SelectionRange = Selection
    End If

    If TypeName(ActiveCell) = "Range" Then
        Set ctx.ActiveCellRef = ActiveCell
    End If
    
    Set CaptureActionContext = ctx
End Function

Public Function CanModifyContext(ByVal ctx As Infra_ActionContext) As Boolean
    If ctx Is Nothing Then Exit Function
    If Not ctx.HasRangeSelection Then Exit Function
    If ctx.WorksheetRef Is Nothing Then Exit Function
    If ctx.ActiveCellRef Is Nothing Then Exit Function

    If Not ctx.WorksheetRef.ProtectContents Then
        CanModifyContext = True
    Else
        CanModifyContext = Not ctx.ActiveCellRef.Locked
    End If
End Function
