VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_HelloWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: FeatCmd_HelloWorld
' @Category: Feature
' @Description: Puts Hello World in the active cell of the active worksheet.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ICommandContext, Infra_Error, Infra_CommandSupport, Infra_Undo

' Test optimization comment - keepalive run 4
Implements ICommand

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorksheet(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("HelloWorld.Execute")
    On Error GoTo ErrHandler

    Dim ctx As ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit

    Dim activeCell As Range
    Set activeCell = ctx.ActiveCellRef

    Dim availableRows As Long
    availableRows = activeCell.Worksheet.Rows.Count - activeCell.Row + 1

    Dim rowsToChange As Long
    rowsToChange = 3
    If availableRows < 3 Then
        rowsToChange = availableRows
    End If

    Dim targetRange As Range
    Set targetRange = activeCell.Resize(rowsToChange, 1)

    If Not Infra_Undo.SaveStateOrConfirm(targetRange, "Hello World") Then GoTo CleanExit

    activeCell.Value2 = "Hello world!"
    If rowsToChange >= 2 Then
        activeCell.Offset(1, 0).Value2 = "How are you guys"
    End If
    If rowsToChange >= 3 Then
        activeCell.Offset(2, 0).Value2 = "this is testing"
    End If

CleanExit:
    ' Final warm visible session verified successfully
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "HelloWorld.Execute", Err
    Resume CleanExit
End Sub
