VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_Duplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_Duplicate
' @Category: Feature
' @Description: Duplicates the current workbook to the Desktop as a macro-free .xlsx file.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy(True, False, True)
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorkbook(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Duplicate.Execute")
    On Error GoTo ErrHandler

    Dim desktopPath As String
    Dim fullPath As String
    Dim sourceWB As Workbook
    Dim copyWB As Workbook
    Dim baseName As String
    Dim ctx As Infra_ActionContext
    
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit

    Set sourceWB = ctx.WorkbookRef

    ' --- Extract current workbook name without extension ---
    baseName = sourceWB.Name
    If InStrRev(baseName, ".") > 0 Then
        baseName = Left(baseName, InStrRev(baseName, ".") - 1)
    End If

    ' --- Prompt user for new name ---
    Dim inputResult As Variant
    inputResult = Infra_UIFactory.PromptForDuplicateName(ctx, baseName)
    If VarType(inputResult) = vbBoolean And inputResult = False Then GoTo CleanExit

    baseName = CStr(inputResult)

    ' --- Build full path ---
    desktopPath = Infra_AppState.GetDesktopPath()
    If desktopPath = "" Then
        Infra_Interaction.ShowCritical "Could not locate the Desktop path. Cannot save the file."
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

    Infra_Interaction.ShowInfo "Duplicate saved to Desktop:" & vbCrLf & baseName & ".xlsx"

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not copyWB Is Nothing Then copyWB.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Duplicate.Execute", Err
    Resume CleanExit
End Sub
