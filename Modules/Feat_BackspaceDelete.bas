Attribute VB_Name = "Feat_BackspaceDelete"
Option Explicit

' @Module: Feat_BackspaceDelete
' @Category: Feature
' @Description: Custom handlers for Backspace and Delete keys with protection awareness.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_Error

' Assigned to the Backspace key.
' Clears the contents of the selected range.
Public Sub Backspace()
    PushContext "Backspace"
    On Error GoTo ErrHandler
    
    ' Guard: must be a worksheet and cell must be modifiable
    If Not Infra_AppState.CanModifyActiveCell() Then GoTo CleanExit
    
    ' Checks if the current selection is a range of cells
    If TypeName(Selection) = "Range" Then
        Selection.ClearContents
    End If
    
CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "Backspace", Err
End Sub

' Assigned to the Delete key.
' Clears a range or deletes an object depending on what is selected.
Public Sub Delete()
    PushContext "Delete"
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
    PopContext
    Exit Sub

ErrHandler:
    HandleError "Delete", Err
End Sub
