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
' @Description: Duplicates the current workbook as a macro-free .xlsx file and lets the user choose where to save it.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorkbook(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Duplicate.Execute")
    On Error GoTo ErrHandler

    Dim fullPath As String
    Dim sourceWB As Workbook
    Dim copyWB As Workbook
    Dim baseName As String
    Dim ext As String
    Dim dotPos As Long
    Dim desktopPath As String
    Dim affix As String
    Dim counter As Long
    Dim ctx As Infra_ActionContext
    
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit

    Set sourceWB = ctx.WorkbookRef
    If sourceWB Is Nothing Then GoTo CleanExit

    ' --- Determine Desktop Path ---
    desktopPath = Infra_AppState.GetDesktopPath()
    If desktopPath = vbNullString Then
        desktopPath = Environ("USERPROFILE") & "\Desktop"
    End If

    ' --- Extract current workbook name and extension ---
    baseName = sourceWB.Name
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then
        ext = Mid(baseName, dotPos + 1)
        baseName = Left(baseName, dotPos - 1)
    Else
        ext = "xlsx" ' Default for unsaved workbooks
    End If

    ' --- Build Unique File Path ---
    affix = "_Copy"
    fullPath = Infra_AppState.CombinePath(desktopPath, baseName & affix & "." & ext)
    
    counter = 1
    Do While Infra_AppState.FileExists(fullPath)
        fullPath = Infra_AppState.CombinePath(desktopPath, baseName & affix & " (" & counter & ")." & ext)
        counter = counter + 1
    Loop

    ' --- Save Copy ---
    ' SaveCopyAs creates a file on disk matching the current in-memory state of sourceWB.
    Application.DisplayAlerts = False
    sourceWB.SaveCopyAs fullPath
    Application.DisplayAlerts = True

    ' --- Open the Copy ---
    Set copyWB = Workbooks.Open(fullPath)

    ' --- Close the Original Workbook (Deferred) ---
    ' We defer this using Application.OnTime because closing the workbook that initiated a
    ' Ribbon/Hotkey callback during execution causes Excel to throw an Application-defined
    ' or object-defined error (1004) when the callback returns.
    Dim originalName As String
    originalName = sourceWB.Name
    Application.OnTime EarliestTime:=Now, Procedure:="'Infra_CommandSupport.CloseWorkbookDeferred """ & Replace(originalName, "'", "''") & """'"

    Infra_Interaction.ShowInfo "Duplicate workbook opened:" & vbCrLf & fullPath

CleanExit:
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not copyWB Is Nothing Then copyWB.Close SaveChanges:=False
    On Error GoTo 0
    Infra_Error.HandleError "Duplicate.Execute", Err
    Resume CleanExit
End Sub
