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

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("Delete.Validate")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then
        Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("There is no active Excel selection available for this command.")
    ElseIf ctx.SelectionTypeName = "Range" Then
        Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
    Else
        Set ICommand_Validate = Infra_CommandSupport.ValidationSuccess()
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "Delete.Validate", Err
    Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("The current selection cannot be deleted.")
    Resume CleanExit
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Delete.Execute")
    On Error GoTo ErrHandler
    
    ' Guard: for range selections, check modifiability
    Dim ctx As Infra_ActionContext
    Dim selectionRef As Object

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit

    If ctx.SelectionTypeName = "Range" Then
        If Not Infra_AppState.CanModifyContext(ctx) Then GoTo CleanExit
    End If

    Dim selType As String
    selType = ctx.SelectionTypeName
    Set selectionRef = ctx.SelectionRef
    
    Select Case selType
        Case "Range"
            ctx.SelectionRange.Clear

        Case "Picture", "Shape", "DrawingObjects", "OLEObject", "ChartObject"
            selectionRef.Delete

        Case "ChartArea"
            selectionRef.Parent.Parent.Delete

        Case "PlotArea", "Legend", "Axis", "Series", "ChartTitle", "DataLabel", "Floor", "Walls"
            selectionRef.Parent.Parent.Delete

        Case Else
            On Error Resume Next
            selectionRef.Delete
            On Error GoTo ErrHandler
    End Select
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "Delete.Execute", Err
    Resume CleanExit
End Sub
