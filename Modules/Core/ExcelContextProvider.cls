VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelContextProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @Module: ExcelContextProvider
' @Category: Core
' @Description: Captures and evaluates the current Excel workbook, sheet, window, and selection context.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ActionContext

Public Property Get IsRangeSelected() As Boolean
    IsRangeSelected = (TypeName(Selection) = "Range")
End Property

Public Property Get CanModifyActiveCell() As Boolean
    CanModifyActiveCell = CanModifyContext(CaptureActionContext())
End Property

Public Function CaptureActionContext() As ActionContext
    Dim tracker As Object: Set tracker = Infra_Error.Track("CaptureActionContext")
    On Error GoTo ErrHandler

    Dim ctx As New ActionContext

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

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CaptureActionContext", Err
    Resume CleanExit
End Function

Public Function CanModifyContext(ByVal ctx As ActionContext) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("CanModifyContext")
    On Error GoTo ErrHandler

    If ctx Is Nothing Then Exit Function
    If Not ctx.HasRangeSelection Then Exit Function
    If ctx.WorksheetRef Is Nothing Then Exit Function
    If ctx.ActiveCellRef Is Nothing Then Exit Function

    If Not ctx.WorksheetRef.ProtectContents Then
        CanModifyContext = True
    Else
        Dim hasLocked As Boolean
        If ctx.SelectionRange.Cells.CountLarge = 1 Then
            hasLocked = ctx.SelectionRange.Locked
        Else
            hasLocked = IsNull(ctx.SelectionRange.Locked)
            If Not hasLocked Then
                hasLocked = ctx.SelectionRange.Locked
            End If
        End If
        CanModifyContext = Not hasLocked
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CanModifyContext", Err
    Resume CleanExit
End Function
