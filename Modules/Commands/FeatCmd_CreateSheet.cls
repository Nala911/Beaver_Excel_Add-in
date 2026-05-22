VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_CreateSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_CreateSheet
' @Category: Feature
' @Description: Creates a new worksheet with a user-supplied name and explicit placement.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorkbook(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateSheet.Execute")
    On Error GoTo ErrHandler

    Dim sheetName As String
    Dim newSheet As Worksheet
    Dim ctx As Infra_ActionContext
    Dim insertPosition As SheetInsertPosition

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit

    sheetName = Infra_UIFactory.PromptForSheetName(ctx)
    If sheetName = vbNullString Then GoTo CleanExit

    insertPosition = Infra_UIFactory.PromptForSheetInsertPosition(ctx, sheetName)
    If insertPosition = 0 Then GoTo CleanExit

    If insertPosition = SheetInsertPositionBeforeCurrent Then
        Set newSheet = ctx.WorkbookRef.Worksheets.Add(Before:=ctx.WorksheetRef)
    Else
        Set newSheet = ctx.WorkbookRef.Worksheets.Add(After:=ctx.WorksheetRef)
    End If

    newSheet.Name = sheetName

CleanExit:
    Exit Sub

ErrHandler:
    If Not newSheet Is Nothing Then
        Application.DisplayAlerts = False
        newSheet.Delete
        Application.DisplayAlerts = True
    End If
    Infra_Error.HandleError "CreateSheet.Execute", Err
    Resume CleanExit
End Sub
