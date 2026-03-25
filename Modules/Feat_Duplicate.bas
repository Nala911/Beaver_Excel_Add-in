Attribute VB_Name = "Feat_Duplicate"
Option Explicit

' @Module: Feat_Duplicate
' @Category: Feature
' @Description: Duplicates the current workbook to the Desktop as a macro-free .xlsx file.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_AppStateGuard, Infra_Config, Infra_Error, Infra_UIFactory

Public Sub Duplicate()
    PushContext "Duplicate"
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim desktopPath As String
    Dim fullPath As String
    Dim sourceWB As Workbook
    Dim copyWB As Workbook
    Dim baseName As String
    Dim ctx As Infra_ActionContext
    
    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit

    Set sourceWB = ctx.WorkbookRef

    ' --- Extract current workbook name without extension ---
    baseName = sourceWB.Name
    If InStrRev(baseName, ".") > 0 Then
        baseName = Left(baseName, InStrRev(baseName, ".") - 1)
    End If

    ' --- Show screen so prompt is visible ---
    Application.ScreenUpdating = True

    ' --- Prompt user for new name ---
    Dim inputResult As Variant
    inputResult = Infra_UIFactory.PromptForDuplicateName(ctx, baseName)
    If VarType(inputResult) = vbBoolean And inputResult = False Then GoTo CleanExit

    baseName = CStr(inputResult)

    Application.ScreenUpdating = False

    ' --- Build full path ---
    desktopPath = Infra_AppState.GetDesktopPath()
    If desktopPath = "" Then
        MsgBox "Could not locate the Desktop path. Cannot save the file.", vbCritical, Infra_Config.ADDIN_NAME
        GoTo CleanExit
    End If

    fullPath = desktopPath & "\" & baseName & ".xlsx"

    ' --- Copy all sheets to a brand-new workbook ---
    sourceWB.Sheets.Copy
    Set copyWB = ActiveWorkbook   ' Sheets.Copy makes the new workbook active

    ' --- Save the copy as genuine .xlsx (no macros), then close it ---
    Application.DisplayAlerts = False
    copyWB.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    copyWB.Close SaveChanges:=False
    Application.DisplayAlerts = True

    ' --- Re-activate the original workbook ---
    sourceWB.Activate

    MsgBox "Duplicate saved to Desktop:" & vbCrLf & baseName & ".xlsx", vbInformation, Infra_Config.ADDIN_NAME

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not copyWB Is Nothing Then copyWB.Close SaveChanges:=False
    On Error GoTo 0
    HandleError "Duplicate", Err
End Sub
