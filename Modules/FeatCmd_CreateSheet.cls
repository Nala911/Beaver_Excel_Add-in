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
' @Description: Creates a new worksheet with a user-supplied name and smart positioning.

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateSheet.CanExecute")
    On Error GoTo ErrHandler
    ICommand_CanExecute = True
    Exit Function
ErrHandler:
    Infra_Error.HandleError "CreateSheet.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateSheet.Execute")
    On Error GoTo ErrHandler

    Dim sheetName As Variant
    Dim newSheet As Worksheet
    Dim nameLower As String
    Dim ctx As Infra_ActionContext

    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit

    sheetName = InputBox("Please enter the name for the new sheet:", "Beaver Add-in â€” Create Sheet")

    If StrPtr(sheetName) = 0 Then GoTo CleanExit

    If Trim(CStr(sheetName)) = "" Then
        MsgBox "Sheet name cannot be blank.", vbExclamation, "Beaver Add-in"
        GoTo CleanExit
    End If

    nameLower = LCase(CStr(sheetName))

    If (Left(nameLower, 7) = "summary") Or (Left(nameLower, 5) = "recon") Then
        Set newSheet = ctx.WorkbookRef.Worksheets.Add(Before:=ctx.WorksheetRef)
    Else
        Set newSheet = ctx.WorkbookRef.Worksheets.Add(After:=ctx.WorksheetRef)
    End If

    newSheet.Name = CStr(sheetName)

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
