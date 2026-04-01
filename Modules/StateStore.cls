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
    If TypeName(Selection) <> "Range" Then Exit Property
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Property
    If ActiveCell Is Nothing Then Exit Property

    If Not ActiveSheet.ProtectContents Then
        CanModifyActiveCell = True
    Else
        CanModifyActiveCell = Not ActiveCell.Locked
    End If
End Property

Public Function CaptureActionContext() As Infra_ActionContext
    Dim ctx As New Infra_ActionContext
    
    If Not ActiveWorkbook Is Nothing Then Set ctx.WorkbookRef = ActiveWorkbook
    If TypeName(ActiveSheet) = "Worksheet" Then Set ctx.WorksheetRef = ActiveSheet
    
    If TypeName(Selection) = "Range" Then
        ctx.HasRangeSelection = True
        Set ctx.SelectionRange = Selection
    End If

    If TypeName(ActiveCell) = "Range" Then
        Set ctx.ActiveCellRef = ActiveCell
    End If
    
    Set CaptureActionContext = ctx
End Function
