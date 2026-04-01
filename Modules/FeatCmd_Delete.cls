VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_Delete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_Delete
' @Category: Feature
' @Description: Custom handlers for Delete key with protection awareness.

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("Delete.CanExecute")
    On Error GoTo ErrHandler
    
    If TypeName(Selection) = "Range" Then
        ICommand_CanExecute = Infra_AppState.CanModifyActiveCell()
    Else
        ICommand_CanExecute = True ' Allow deleting objects unconditionally for now
    End If
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Delete.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Delete.Execute")
    On Error GoTo ErrHandler
    
    ' Guard: for range selections, check modifiability
    If TypeName(Selection) = "Range" Then
        If Not Infra_AppState.CanModifyActiveCell() Then GoTo CleanExit
    End If

    Dim selType As String
    selType = TypeName(Selection)
    
    Select Case selType
        Case "Range"
            ' Clear values + formats
            Selection.Clear

        Case "Picture", "Shape", "DrawingObjects", "OLEObject", "ChartObject"
            ' Directly deletable objects
            Selection.Delete

        Case "ChartArea"
            Selection.Parent.Parent.Delete

        Case "PlotArea", "Legend", "Axis", "Series", "ChartTitle", "DataLabel", "Floor", "Walls"
            ' Deeper chart sub-elements
            Selection.Parent.Parent.Delete

        Case Else
            ' Last resort: try to delete
            On Error Resume Next
            Selection.Delete
            On Error GoTo ErrHandler
    End Select
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Delete.Execute", Err
    Resume CleanExit
End Sub
