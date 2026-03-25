Attribute VB_Name = "Feat_CreateSheet"
Option Explicit

' @Module: Feat_CreateSheet
' @Category: Feature
' @Description: Creates a new worksheet with a user-supplied name and smart positioning.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppStateGuard, Infra_Config, Infra_Error, Infra_AppState

Public Sub CreateNamedSheet()
    PushContext "CreateNamedSheet"
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim sheetName As Variant
    Dim newSheet As Worksheet
    Dim nameLower As String
    Dim ctx As Infra_ActionContext

    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit

    sheetName = InputBox("Please enter the name for the new sheet:", Infra_Config.ADDIN_NAME & " â€” Create Sheet")

    If StrPtr(sheetName) = 0 Then GoTo CleanExit

    If Trim(CStr(sheetName)) = "" Then
        MsgBox "Sheet name cannot be blank.", vbExclamation, Infra_Config.ADDIN_NAME
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
    PopContext
    Exit Sub

ErrHandler:
    If Not newSheet Is Nothing Then
        Application.DisplayAlerts = False
        newSheet.Delete
        Application.DisplayAlerts = True
    End If
    HandleError "CreateNamedSheet", Err
End Sub
